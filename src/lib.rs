use pyo3::prelude::*;
use rayon::prelude::*;
use rayon::slice::ParallelSliceMut;
use std::collections::HashMap;
use rand::prelude::*;

// --- 採点ロジック ---
fn calculate_single_score(
    schedule: &Vec<Vec<i32>>,
    roles: &Vec<String>,
    constraints: &HashMap<(usize, usize), String>,
    days: usize,
    staff_count: usize
) -> i32 {
    let mut score = 100;

    // 1. 制約チェック (希望休など)
    for staff_idx in 0..staff_count {
        for day in 0..days {
            if let Some(constraint) = constraints.get(&(staff_idx, day)) {
                let current_shift = schedule[staff_idx][day];
                if constraint == "NG" && current_shift != 0 { score -= 10000; }
                else if constraint == "NO_MORNING" && current_shift == 1 { score -= 10000; }
                else if constraint == "NO_NIGHT" && current_shift == 2 { score -= 10000; }
            }
        }
    }

    // 2. 個人チェック (勤務間隔・日数・連勤)
    for (staff_idx, staff_row) in schedule.iter().enumerate() {
        
        // --- ★追加: 連勤チェック ---
        let mut consecutive_days = 0;
        for day in 0..days {
            if staff_row[day] != 0 {
                // 出勤ならカウントアップ
                consecutive_days += 1;
                
                // ★ルール: 7連勤以上になったら大幅減点
                // (6連勤までは許容する設定です。厳しくするなら > 5 にしてください)
                if consecutive_days > 5 {
                    score -= 100;
                }
            } else {
                // 休みならリセット
                consecutive_days = 0;
            }
        }
        // ---------------------------

        // 夜勤のあとのインターバルチェック
        for day in 0..days {
            if staff_row[day] == 2 {
                // 翌日が朝番ならダメ
                if day + 1 < days && staff_row[day+1] == 1 {
                    score -= 100;
                }
                // 翌々日が朝番ならダメ (2日空ける)
                if day + 2 < days && staff_row[day+2] == 1 {
                    score -= 100;
                }
            }
        }

        // 勤務日数の目標
        let work_days = staff_row.iter().filter(|&&s| s != 0).count() as i32;
        let role = &roles[staff_idx];
        let target_days = if role == "Assist" { 10 } else { 21 };
        
        let diff = (work_days - target_days).abs();
        score -= diff * 10;
    }

    // 3. 運営チェック (人数)
    for day in 0..days {
        let mut morning_count = 0;
        let mut night_count = 0;
        let mut total_workers = 0;
        let mut chief_leader_count = 0;

        for staff_idx in 0..staff_count {
            let shift = schedule[staff_idx][day];
            if shift != 0 {
                total_workers += 1;
                if shift == 1 { morning_count += 1; }
                else if shift == 2 { night_count += 1; }

                let role = &roles[staff_idx];
                if role == "Chief" || role == "Leader" { chief_leader_count += 1; }
            }
        }

        if morning_count < 5 { score -= 50; }
        if night_count < 5 { score -= 50; }
        if total_workers < 10 { score -= 50; }
        if chief_leader_count < 2 { score -= 30; }
    }
    score
}

// --- 遺伝的アルゴリズム本体 ---
#[pyfunction]
fn run_genetic_algorithm(
    roles: Vec<String>,
    constraints: HashMap<(usize, usize), String>,
    days: usize,
    staff_count: usize,
    population_size: usize,
    generations: usize
) -> PyResult<(Vec<Vec<i32>>, i32)> {

    // 初期個体
    let mut population: Vec<Vec<Vec<i32>>> = (0..population_size).into_par_iter().map(|_| {
        let mut rng = rand::thread_rng();
        let mut schedule = Vec::with_capacity(staff_count);
        for _ in 0..staff_count {
            let mut row = Vec::with_capacity(days);
            for _ in 0..days {
                row.push(rng.gen_range(0..=2));
            }
            schedule.push(row);
        }
        schedule
    }).collect();

    for generation_idx in 0..generations {
        // 採点
        let mut scored_population: Vec<(i32, Vec<Vec<i32>>)> = population.par_iter()
            .map(|sch| {
                let s = calculate_single_score(sch, &roles, &constraints, days, staff_count);
                (s, sch.clone())
            })
            .collect();

        scored_population.par_sort_unstable_by(|a, b| b.0.cmp(&a.0));

        if scored_population[0].0 == 100 {
            let (best_score, best_schedule) = &scored_population[0];
            println!("Rust: Generation {} found 100 score!", generation_idx);
            return Ok((best_schedule.clone(), *best_score));
        }
        
        if generation_idx % 5 == 0 {
            println!("Rust: Gen {} Best Score = {}", generation_idx, scored_population[0].0);
        }

        // 次世代生成
        let elite_count = (population_size as f64 * 0.2) as usize;
        let mut next_gen = Vec::with_capacity(population_size);

        for i in 0..elite_count {
            next_gen.push(scored_population[i].1.clone());
        }

        let num_children = population_size - elite_count;
        let children: Vec<Vec<Vec<i32>>> = (0..num_children).into_par_iter().map(|_| {
            let mut rng = rand::thread_rng();
            let parent_idx1 = rng.gen_range(0..(population_size / 2));
            let parent_idx2 = rng.gen_range(0..(population_size / 2));
            let parent1 = &scored_population[parent_idx1].1;
            let parent2 = &scored_population[parent_idx2].1;

            let mut child = parent1.clone();
            let split = rng.gen_range(1..staff_count);
            for i in split..staff_count {
                child[i] = parent2[i].clone();
            }

            if rng.gen_bool(0.2) {
                let m_staff = rng.gen_range(0..staff_count);
                let m_day = rng.gen_range(0..days);
                child[m_staff][m_day] = rng.gen_range(0..=2);
            }
            child
        }).collect();

        next_gen.extend(children);
        population = next_gen;
    }

    let best_schedule = &population[0];
    let score = calculate_single_score(best_schedule, &roles, &constraints, days, staff_count);
    Ok((best_schedule.clone(), score))
}

#[pymodule]
#[pyo3(name = "ShiftScheduler")]
fn ShiftScheduler(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(run_genetic_algorithm, m)?)?;
    Ok(())
}