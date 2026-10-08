# betteroffice-wasm-alloc

Global allocator for BetterOffice's WASM modules: Rust's dlmalloc, growing
linear memory in steps of up to 64 MiB instead of the 64 KiB a request needs.
A module links it with `use wasm_alloc as _;`; other targets keep theirs.
