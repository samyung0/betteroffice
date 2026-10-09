//! Global allocator for the WASM modules: dlmalloc, growing linear memory in
//! steps of up to 64 MiB instead of the 64 KiB each request needs. Every grow
//! costs the browser time that rises with memory size; thousands of them made
//! large opens slow (a 248-page DOCX view spent 102 s of 121 s growing).
//! Linked with `use wasm_alloc as _;`.

/// Bytes one step reserves once memory is `pages` 64 KiB pages: as much as
/// memory, at most 64 MiB. Counted in pages, since `pages * 64 KiB` overflows
/// 32 bits at the 4 GiB limit (65,536 pages).
#[cfg(any(target_arch = "wasm32", test))]
fn step_bytes(pages: u32) -> u32 {
    const PAGE: u32 = 64 << 10;
    const MAX_STEP: u32 = 64 << 20;
    pages.min(MAX_STEP / PAGE) * PAGE
}

#[cfg(target_arch = "wasm32")]
mod wasm {
    use super::step_bytes;
    use core::arch::wasm32::memory_size;
    use core::sync::atomic::{AtomicU32, Ordering::Relaxed};
    use std::alloc::{GlobalAlloc, Layout, System};

    /// Linear memory grows, dlmalloc's and the steps.
    static GROWS: AtomicU32 = AtomicU32::new(0);

    struct StepAlloc;

    #[global_allocator]
    static GLOBAL: StepAlloc = StepAlloc;

    /// Runs one `System` call; when it grew memory, reserves a step as large
    /// as memory, capped.
    fn stepped(allocate: impl FnOnce() -> *mut u8) -> *mut u8 {
        let before = memory_size::<0>();
        let ptr = allocate();
        let grown = memory_size::<0>();
        if grown != before {
            GROWS.fetch_add(1, Relaxed);
            reserve(step_bytes(grown as u32) as usize);
        }
        ptr
    }

    /// Allocates and frees `bytes`, which dlmalloc keeps as free memory.
    fn reserve(bytes: usize) {
        if bytes == 0 {
            return;
        }
        let Ok(layout) = Layout::from_size_align(bytes, 16) else {
            return;
        };
        let before = memory_size::<0>();
        // SAFETY: `bytes` is not zero, so the layout is not zero-sized.
        let block = unsafe { System.alloc(layout) };
        if memory_size::<0>() != before {
            GROWS.fetch_add(1, Relaxed);
        }
        if !block.is_null() {
            // SAFETY: allocated just above with the same layout.
            unsafe { System.dealloc(block, layout) };
        }
    }

    // SAFETY: every method forwards to `System` with its own arguments.
    unsafe impl GlobalAlloc for StepAlloc {
        unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
            stepped(|| unsafe { System.alloc(layout) })
        }

        unsafe fn alloc_zeroed(&self, layout: Layout) -> *mut u8 {
            stepped(|| unsafe { System.alloc_zeroed(layout) })
        }

        unsafe fn realloc(&self, ptr: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
            stepped(|| unsafe { System.realloc(ptr, layout, new_size) })
        }

        unsafe fn dealloc(&self, ptr: *mut u8, layout: Layout) {
            unsafe { System.dealloc(ptr, layout) }
        }
    }

    /// Linear memory grows since the module started, for tests and benches.
    #[unsafe(no_mangle)]
    pub extern "C" fn betteroffice_memory_grows() -> u32 {
        GROWS.load(Relaxed)
    }
}

#[cfg(test)]
mod tests {
    use super::step_bytes;

    // Security regression WM1: at exactly 4 GiB the step overflowed 32 bits,
    // trapping inside the allocator (debug) or reserving a zero-size layout.
    #[test]
    fn a_step_is_memory_size_capped_at_64_mib_up_to_4_gib() {
        assert_eq!(step_bytes(1), 64 << 10);
        assert_eq!(step_bytes(1024), 64 << 20);
        assert_eq!(step_bytes(1025), 64 << 20);
        assert_eq!(step_bytes(65_536), 64 << 20);
    }
}
