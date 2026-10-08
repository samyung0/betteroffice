//! Hostile parts from the 2026-10-08 review, decoded under a counting
//! allocator: a CTF directory of 256 entries each spanning a 1, 8 or 15 MiB
//! block (which used to copy every entry, up to 3.8 GiB), and a 150-byte part
//! that unpacks 16 MiB before the CTF stage refuses it.

use std::alloc::{GlobalAlloc, Layout, System};
use std::sync::atomic::{AtomicUsize, Ordering};
use std::time::{Duration, Instant};

use ooxml_text::decode_embedded_font;

struct Counting;

static CURRENT: AtomicUsize = AtomicUsize::new(0);
static PEAK: AtomicUsize = AtomicUsize::new(0);

fn grew(by: usize) {
    let now = CURRENT.fetch_add(by, Ordering::Relaxed) + by;
    PEAK.fetch_max(now, Ordering::Relaxed);
}

unsafe impl GlobalAlloc for Counting {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        let pointer = unsafe { System.alloc(layout) };
        if !pointer.is_null() {
            grew(layout.size());
        }
        pointer
    }

    unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
        unsafe { System.dealloc(pointer, layout) };
        CURRENT.fetch_sub(layout.size(), Ordering::Relaxed);
    }

    unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, size: usize) -> *mut u8 {
        let moved = unsafe { System.realloc(pointer, layout, size) };
        if !moved.is_null() {
            if size > layout.size() {
                grew(size - layout.size());
            } else {
                CURRENT.fetch_sub(layout.size() - size, Ordering::Relaxed);
            }
        }
        moved
    }
}

#[global_allocator]
static ALLOCATOR: Counting = Counting;

const MIB: usize = 1024 * 1024;

#[test]
fn hostile_parts_are_refused_within_bounded_memory_and_time() {
    for (name, part) in [
        (
            "256 tables over 1 MiB",
            &include_bytes!("fonts/hostile-tables-256x1.fntdata")[..],
        ),
        (
            "256 tables over 8 MiB",
            include_bytes!("fonts/hostile-tables-256x8.fntdata"),
        ),
        (
            "256 tables over 15 MiB",
            include_bytes!("fonts/hostile-tables-256x15.fntdata"),
        ),
        (
            "16 MiB expansion",
            include_bytes!("fonts/hostile-expand.fntdata"),
        ),
    ] {
        let base = CURRENT.load(Ordering::Relaxed);
        PEAK.store(base, Ordering::Relaxed);
        let started = Instant::now();
        let result = decode_embedded_font(part, "ppt/fonts/font1.fntdata").map(|font| font.len());
        let elapsed = started.elapsed();
        let peak = PEAK.load(Ordering::Relaxed) - base;
        println!("{name}: {result:?}, peak {peak} bytes, {elapsed:?}");
        assert!(result.is_err(), "{name}: {result:?}");
        // The block itself (at most 16 MiB) is the only large allocation.
        assert!(peak <= 17 * MIB, "{name}: peak {peak} bytes");
        assert!(elapsed < Duration::from_secs(30), "{name}: {elapsed:?}");
    }
}
