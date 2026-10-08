//! `office-service serve`: one JSON request per stdin line, one answer per
//! stdout line (see `wire`). `office-service bench <request.json>...`: runs
//! the requests in order in this process (a replica hit after its miss) and
//! prints each one's process CPU time, wall time and memory high-water as a
//! JSON line on stdout. With `OFFICE_SERVICE_STACK_STATS` set, `serve` runs
//! each request on a fresh thread and writes `stack <method> <bytes>` to
//! stderr: the stack the call committed (Windows), to size the engine thread.

use std::io::{BufRead, Write};
use std::time::Instant;

/// Engine calls recurse over document trees; give them the stack a deep
/// document needs rather than the platform's main-thread default. The
/// deepest recorded call commits 0.45 MiB in release (API.md, "Stack"); a
/// debug build needs far more (it overflowed the 1 MiB main thread), and the
/// reservation costs nothing until used.
const STACK_BYTES: usize = 256 << 20;

fn main() {
    let args: Vec<String> = std::env::args().collect();
    let worker = std::thread::Builder::new()
        .stack_size(STACK_BYTES)
        .spawn(move || match args.get(1).map(String::as_str) {
            Some("serve") => serve(),
            Some("bench") => bench(&args[2..]),
            _ => {
                eprintln!("usage: office-service serve | bench <request.json>...");
                std::process::exit(2);
            }
        })
        .expect("engine thread starts");
    if worker.join().is_err() {
        std::process::exit(101);
    }
}

fn serve() {
    let stats = std::env::var_os("OFFICE_SERVICE_STACK_STATS").is_some();
    let stdin = std::io::stdin();
    let mut stdout = std::io::stdout().lock();
    for line in stdin.lock().lines() {
        let Ok(line) = line else { break };
        if line.trim().is_empty() {
            continue;
        }
        let answer = if stats {
            stack_stats(line)
        } else {
            office_service::wire::call(line.as_bytes())
        };
        if stdout.write_all(&answer).is_err() || stdout.write_all(b"\n").is_err() {
            break;
        }
        let _ = stdout.flush();
    }
}

/// One request on a fresh thread, reporting the stack it committed.
fn stack_stats(line: String) -> Vec<u8> {
    let method = serde_json::from_str::<serde_json::Value>(&line)
        .ok()
        .and_then(|request| request.get("method")?.as_str().map(str::to_owned))
        .unwrap_or_default();
    let (answer, used) = std::thread::Builder::new()
        .stack_size(STACK_BYTES)
        .spawn(move || {
            let answer = office_service::wire::call(line.as_bytes());
            (answer, usage::stack_committed())
        })
        .expect("call thread starts")
        .join()
        .unwrap_or_else(|_| std::process::exit(101));
    eprintln!("stack {method} {}", used.map_or(-1, |used| used as i64));
    answer
}

fn bench(paths: &[String]) {
    if paths.is_empty() {
        eprintln!("usage: office-service bench <request.json>...");
        std::process::exit(2);
    }
    let (_, rss_before) = usage::memory();
    for (run, path) in paths.iter().enumerate() {
        let request = std::fs::read(path).unwrap_or_else(|error| {
            eprintln!("{path}: {error}");
            std::process::exit(2);
        });
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
                "request": path,
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
        fn GetCurrentThreadStackLimits(low: *mut usize, high: *mut usize);
        fn VirtualQuery(address: usize, info: *mut MemoryInfo, length: usize) -> usize;
    }

    #[repr(C)]
    #[derive(Default)]
    struct MemoryInfo {
        base: usize,
        allocation_base: usize,
        allocation_protect: u32,
        partition: u16,
        region: usize,
        state: u32,
        protect: u32,
        kind: u32,
    }

    /// Bytes of this thread's stack committed so far, its high-water mark:
    /// Windows commits stack pages as the stack grows and keeps them.
    pub fn stack_committed() -> Option<usize> {
        const MEM_COMMIT: u32 = 0x1000;
        let (mut low, mut high) = (0usize, 0usize);
        // SAFETY: the pointers are to live locals; VirtualQuery writes at most
        // `size_of::<MemoryInfo>()` bytes into `info`.
        unsafe {
            GetCurrentThreadStackLimits(&mut low, &mut high);
            let mut address = low;
            while address < high {
                let mut info = MemoryInfo::default();
                if VirtualQuery(address, &mut info, std::mem::size_of::<MemoryInfo>()) == 0 {
                    return None;
                }
                if info.state == MEM_COMMIT {
                    return Some(high - address);
                }
                address = info.base + info.region;
            }
        }
        None
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

    /// Not measured off Windows.
    pub fn stack_committed() -> Option<usize> {
        None
    }

    /// Peak resident set (current is not tracked), in bytes.
    pub fn memory() -> (usize, usize) {
        let peak = rusage().max_rss as usize * 1024;
        (peak, peak)
    }
}
