//! `office-service serve`: one JSON request per stdin line, one answer per
//! stdout line (see `wire`). `office-service bench <request.json> [runs]`:
//! process CPU time, wall time and memory high-water of each run of one
//! request, as JSON lines on stdout.

use std::io::{BufRead, Write};
use std::time::Instant;

/// Engine calls recurse over document trees; give them the stack a deep
/// document needs rather than the platform's main-thread default.
const STACK_BYTES: usize = 256 << 20;

fn main() {
    let args: Vec<String> = std::env::args().collect();
    let worker = std::thread::Builder::new()
        .stack_size(STACK_BYTES)
        .spawn(move || match args.get(1).map(String::as_str) {
            Some("serve") => serve(),
            Some("bench") => bench(&args[2..]),
            _ => {
                eprintln!("usage: office-service serve | bench <request.json> [runs]");
                std::process::exit(2);
            }
        })
        .expect("engine thread starts");
    if worker.join().is_err() {
        std::process::exit(101);
    }
}

fn serve() {
    let stdin = std::io::stdin();
    let mut stdout = std::io::stdout().lock();
    for line in stdin.lock().lines() {
        let Ok(line) = line else { break };
        if line.trim().is_empty() {
            continue;
        }
        let answer = office_service::wire::call(line.as_bytes());
        if stdout.write_all(&answer).is_err() || stdout.write_all(b"\n").is_err() {
            break;
        }
        let _ = stdout.flush();
    }
}

fn bench(args: &[String]) {
    let Some(path) = args.first() else {
        eprintln!("usage: office-service bench <request.json> [runs]");
        std::process::exit(2);
    };
    let runs: usize = args.get(1).and_then(|runs| runs.parse().ok()).unwrap_or(1);
    let request = std::fs::read(path).unwrap_or_else(|error| {
        eprintln!("{path}: {error}");
        std::process::exit(2);
    });
    let (_, rss_before) = usage::memory();
    for run in 0..runs {
        let cpu = usage::cpu_ms();
        let wall = Instant::now();
        let answer = office_service::wire::call(&request);
        let wall_ms = wall.elapsed().as_secs_f64() * 1000.0;
        let cpu_ms = usage::cpu_ms() - cpu;
        let (peak, rss) = usage::memory();
        let error = serde_json::from_slice::<serde_json::Value>(&answer)
            .ok()
            .and_then(|answer| answer.get("error").cloned());
        println!(
            "{}",
            serde_json::json!({
                "run": run,
                "cpuMs": cpu_ms.round(),
                "wallMs": wall_ms.round(),
                "peakMiB": (peak as f64 / 1048576.0 * 10.0).round() / 10.0,
                "rssBeforeMiB": (rss_before as f64 / 1048576.0 * 10.0).round() / 10.0,
                "rssAfterMiB": (rss as f64 / 1048576.0 * 10.0).round() / 10.0,
                "answerBytes": answer.len(),
                "error": error,
            })
        );
    }
}

#[cfg(windows)]
mod usage {
    #[repr(C)]
    #[derive(Default)]
    struct FileTime {
        low: u32,
        high: u32,
    }

    #[repr(C)]
    #[derive(Default)]
    struct MemoryCounters {
        cb: u32,
        page_fault_count: u32,
        peak_working_set_size: usize,
        working_set_size: usize,
        quota_peak_paged_pool_usage: usize,
        quota_paged_pool_usage: usize,
        quota_peak_non_paged_pool_usage: usize,
        quota_non_paged_pool_usage: usize,
        pagefile_usage: usize,
        peak_pagefile_usage: usize,
    }

    #[link(name = "kernel32")]
    unsafe extern "system" {
        fn GetCurrentProcess() -> isize;
        fn GetProcessTimes(
            process: isize,
            creation: *mut FileTime,
            exit: *mut FileTime,
            kernel: *mut FileTime,
            user: *mut FileTime,
        ) -> i32;
        fn K32GetProcessMemoryInfo(process: isize, counters: *mut MemoryCounters, cb: u32) -> i32;
    }

    fn ticks(time: &FileTime) -> f64 {
        ((time.high as u64) << 32 | time.low as u64) as f64 / 10_000.0
    }

    /// User plus kernel CPU of this process, in milliseconds.
    pub fn cpu_ms() -> f64 {
        let (mut creation, mut exit, mut kernel, mut user) = Default::default();
        // SAFETY: the pointers are to live locals of the declared layout.
        unsafe {
            GetProcessTimes(
                GetCurrentProcess(),
                &mut creation,
                &mut exit,
                &mut kernel,
                &mut user,
            );
        }
        ticks(&kernel) + ticks(&user)
    }

    /// Peak and current working set, in bytes.
    pub fn memory() -> (usize, usize) {
        let mut counters = MemoryCounters {
            cb: std::mem::size_of::<MemoryCounters>() as u32,
            ..Default::default()
        };
        // SAFETY: `counters` is a live MemoryCounters with its size in `cb`.
        unsafe {
            K32GetProcessMemoryInfo(GetCurrentProcess(), &mut counters, counters.cb);
        }
        (counters.peak_working_set_size, counters.working_set_size)
    }
}

#[cfg(not(windows))]
mod usage {
    #[repr(C)]
    struct TimeVal {
        sec: i64,
        usec: i64,
    }

    #[repr(C)]
    struct RUsage {
        user: TimeVal,
        system: TimeVal,
        max_rss: i64,
        rest: [i64; 13],
    }

    unsafe extern "C" {
        fn getrusage(who: i32, usage: *mut RUsage) -> i32;
    }

    fn rusage() -> RUsage {
        let mut usage = RUsage {
            user: TimeVal { sec: 0, usec: 0 },
            system: TimeVal { sec: 0, usec: 0 },
            max_rss: 0,
            rest: [0; 13],
        };
        // SAFETY: `usage` is a live struct of getrusage's layout on 64-bit Linux.
        unsafe {
            getrusage(0, &mut usage);
        }
        usage
    }

    pub fn cpu_ms() -> f64 {
        let usage = rusage();
        let ms = |time: &TimeVal| time.sec as f64 * 1000.0 + time.usec as f64 / 1000.0;
        ms(&usage.user) + ms(&usage.system)
    }

    /// Peak resident set (current is not tracked), in bytes.
    pub fn memory() -> (usize, usize) {
        let peak = rusage().max_rss as usize * 1024;
        (peak, peak)
    }
}
