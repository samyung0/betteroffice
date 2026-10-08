//! Security regression (notes REVIEW2 M6): the patched yrs decoders reserved
//! what an update declared (a map's or array's length, the counts of clients,
//! structs, delete ranges and state vector entries) before reading it, so a
//! few bytes made a peer's update write hundreds of MiB. Decoding now grows
//! with what the input holds. Its own test binary, so the counting allocator
//! sees only this test.

use std::alloc::{GlobalAlloc, Layout, System};
use std::sync::atomic::{AtomicUsize, Ordering};

use yrs::encoding::write::Write;
use yrs::updates::decoder::Decode;
use yrs::{StateVector, Update};

struct Counting;

static ALLOCATED: AtomicUsize = AtomicUsize::new(0);

unsafe impl GlobalAlloc for Counting {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        ALLOCATED.fetch_add(layout.size(), Ordering::SeqCst);
        unsafe { System.alloc(layout) }
    }

    unsafe fn dealloc(&self, ptr: *mut u8, layout: Layout) {
        unsafe { System.dealloc(ptr, layout) }
    }

    unsafe fn realloc(&self, ptr: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        ALLOCATED.fetch_add(new_size.saturating_sub(layout.size()), Ordering::SeqCst);
        unsafe { System.realloc(ptr, layout, new_size) }
    }
}

#[global_allocator]
static ALLOCATOR: Counting = Counting;

/// Bytes allocated while `run` ran.
fn allocated(run: impl FnOnce()) -> usize {
    let before = ALLOCATED.load(Ordering::SeqCst);
    run();
    ALLOCATED.load(Ordering::SeqCst) - before
}

/// An update of one client with one item holding `value` as its `Any`
/// content, under the root text "t", cut off right after `value`.
fn any_item(value: &[u8]) -> Vec<u8> {
    let mut bytes = Vec::new();
    bytes.write_var(1u32);
    bytes.write_var(1u32);
    bytes.write_var(77u64);
    bytes.write_var(0u32);
    bytes.write_u8(8);
    bytes.write_var(1u32);
    bytes.write_string("t");
    bytes.write_var(1u32);
    bytes.extend_from_slice(value);
    bytes
}

#[test]
fn declared_lengths_reserve_nothing_the_input_cannot_hold() {
    // 2^20 entries: tens of MiB each when reserved up front, a few bytes here.
    let declared = 1u32 << 20;
    let mut map = vec![118];
    map.write_var(declared);
    let mut array = vec![117];
    array.write_var(declared);
    let mut clients = Vec::new();
    clients.write_var(declared);
    let mut structs = Vec::new();
    structs.write_var(1u32);
    structs.write_var(declared);
    structs.write_var(77u64);
    structs.write_var(0u32);
    let mut deletes = Vec::new();
    deletes.write_var(0u32); // no structs
    deletes.write_var(1u32); // one client in the delete set
    deletes.write_var(77u64);
    deletes.write_var(declared); // its ranges
    let updates = [
        ("a map", any_item(&map)),
        ("an array", any_item(&array)),
        ("the clients", clients),
        ("a client's structs", structs),
        ("a client's delete ranges", deletes),
    ];
    for (name, bytes) in updates {
        let mut decoded = None;
        let bytes_allocated = allocated(|| decoded = Some(Update::decode_v1(&bytes).is_err()));
        assert_eq!(decoded, Some(true), "{name}: decoded");
        assert!(
            bytes_allocated < 64 << 10,
            "{name}: {bytes_allocated} bytes"
        );
    }
    let mut vector = Vec::new();
    vector.write_var(declared);
    let bytes_allocated = allocated(|| assert!(StateVector::decode_v1(&vector).is_err()));
    assert!(
        bytes_allocated < 64 << 10,
        "a state vector: {bytes_allocated} bytes"
    );
}
