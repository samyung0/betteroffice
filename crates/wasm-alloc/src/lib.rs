//! Global allocator for the WASM modules: dlmalloc, growing linear memory in
//! steps of up to 64 MiB instead of the 64 KiB each request needs. Every grow
//! costs the browser time that rises with memory size; thousands of them made
//! large opens slow (a 248-page DOCX view spent 102 s of 121 s growing).
//! Linked with `use wasm_alloc as _;`.

#[cfg(target_arch = "wasm32")]
mod wasm {
    use core::arch::wasm32::memory_size;
    use core::sync::atomic::{AtomicU32, Ordering::Relaxed};
    use std::alloc::{GlobalAlloc, Layout, System};

    const PAGE: usize = 64 << 10;
    /// Caps what one step reserves beyond the need.
    const MAX_STEP: usize = 64 << 20;

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
            reserve((grown * PAGE).min(MAX_STEP));
        }
        ptr
    }

    /// Allocates and frees `bytes`, which dlmalloc keeps as free memory.
    fn reserve(bytes: usize) {
        let Ok(layout) = Layout::from_size_align(bytes, 16) else {
            return;
        };
        let before = memory_size::<0>();
        // SAFETY: `bytes` is a whole number of pages, so the layout is not
        // zero-sized.
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
