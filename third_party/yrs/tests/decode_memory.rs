//! Patched for BetterOffice: decoding allocates for what the input holds,
//! not for the lengths it declares. Its own test binary, so the counting
//! allocator sees only this test.
use std::alloc::{GlobalAlloc, Layout, System};
use std::sync::atomic::{AtomicUsize, Ordering};

use yrs::encoding::read::Cursor;
use yrs::encoding::write::Write;
use yrs::updates::decoder::Decode;
use yrs::{Any, StateVector, Update};

struct Counting;

static LIVE: AtomicUsize = AtomicUsize::new(0);
static PEAK: AtomicUsize = AtomicUsize::new(0);

unsafe impl GlobalAlloc for Counting {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        let live = LIVE.fetch_add(layout.size(), Ordering::SeqCst) + layout.size();
        PEAK.fetch_max(live, Ordering::SeqCst);
        unsafe { System.alloc(layout) }
    }

    unsafe fn dealloc(&self, ptr: *mut u8, layout: Layout) {
        LIVE.fetch_sub(layout.size(), Ordering::SeqCst);
        unsafe { System.dealloc(ptr, layout) }
    }
}

#[global_allocator]
static ALLOCATOR: Counting = Counting;

/// Bytes allocated at the peak while `run` ran, above what was live before.
fn peak_of(run: impl FnOnce()) -> usize {
    let before = LIVE.load(Ordering::SeqCst);
    PEAK.store(before, Ordering::SeqCst);
    run();
    PEAK.load(Ordering::SeqCst) - before
}

#[test]
fn declared_lengths_reserve_nothing() {
    // 64 nested maps each declaring 2^22 entries and holding one (449 bytes;
    // it reserved about 24 GiB before).
    let mut maps = Vec::new();
    for _ in 0..64 {
        maps.write_u8(118);
        maps.write_var(1u64 << 22);
        maps.write_string("k");
    }
    maps.write_u8(126);
    let peak = peak_of(|| assert!(Any::decode(&mut Cursor::new(&maps)).is_err()));
    assert!(peak < 1 << 20, "{} bytes", peak);

    // An update declaring 2^30 clients, a client declaring 2^30 structs, a
    // delete set declaring 2^30 ranges, and a state vector declaring 2^30
    // entries: each is an end of input, not a reservation.
    let mut clients = Vec::new();
    clients.write_var(1u32 << 30);
    let mut structs = Vec::new();
    structs.write_var(1u32);
    structs.write_var(1u32 << 30);
    structs.write_var(7u64);
    structs.write_var(0u32);
    let mut ranges = Vec::new();
    ranges.write_var(0u32);
    ranges.write_var(1u32);
    ranges.write_var(7u64);
    ranges.write_var(1u32 << 30);
    let mut vector = Vec::new();
    vector.write_var(1u32 << 30);
    for bytes in [&clients, &structs, &ranges] {
        let peak = peak_of(|| assert!(Update::decode_v1(bytes).is_err()));
        assert!(peak < 1 << 20, "{} bytes", peak);
    }
    let peak = peak_of(|| assert!(StateVector::decode_v1(&vector).is_err()));
    assert!(peak < 1 << 20, "{} bytes", peak);
}
