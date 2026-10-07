//! Native twin of the editor's resident engine worker: opens a DOCX, lays it
//! out through the region pass, builds the first frame, then types at the end
//! of the first and the last body paragraph. Prints one JSON line with the
//! allocator high-water per stage and the per-key stage medians.
//!
//! usage: type_probe <docx> <request.json> [keys] [latin.ttf] [cjk.otf]
//! `request.json` is `{"request": <resident region request>}` as the host
//! builds it (`buildResidentRegionLayoutRequest` plus its `measurement`).

use std::alloc::{GlobalAlloc, Layout, System};
use std::sync::atomic::{AtomicUsize, Ordering};
use std::time::Instant;

use docx_edit::{
    EditCtx, EngineSession, FormatPolicy, Position, package_media, parse_docx_for_edit,
    seed_parsed_docx,
};
use serde_json::{Map, Value, json};

struct Counting;

static LIVE: AtomicUsize = AtomicUsize::new(0);
static PEAK: AtomicUsize = AtomicUsize::new(0);

unsafe impl GlobalAlloc for Counting {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        let ptr = unsafe { System.alloc(layout) };
        if !ptr.is_null() {
            grow(layout.size());
        }
        ptr
    }
    unsafe fn dealloc(&self, ptr: *mut u8, layout: Layout) {
        unsafe { System.dealloc(ptr, layout) };
        LIVE.fetch_sub(layout.size(), Ordering::Relaxed);
    }
    unsafe fn realloc(&self, ptr: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        let next = unsafe { System.realloc(ptr, layout, new_size) };
        if !next.is_null() {
            LIVE.fetch_sub(layout.size(), Ordering::Relaxed);
            grow(new_size);
        }
        next
    }
}

fn grow(size: usize) {
    let live = LIVE.fetch_add(size, Ordering::Relaxed) + size;
    PEAK.fetch_max(live, Ordering::Relaxed);
}

#[global_allocator]
static ALLOCATOR: Counting = Counting;

fn mb(bytes: usize) -> f64 {
    (bytes as f64 / 1_048_576.0 * 10.0).round() / 10.0
}

/// `{live, peak}` in MB, then resets the peak to the live size.
fn mark() -> Value {
    let live = LIVE.load(Ordering::Relaxed);
    let peak = PEAK.swap(live, Ordering::Relaxed);
    json!({"live": mb(live), "peak": mb(peak)})
}

/// Process CPU time (user + system) in ms.
#[cfg(unix)]
fn cpu_ms() -> f64 {
    #[repr(C)]
    struct Rusage {
        utime: [i64; 2],
        stime: [i64; 2],
        rest: [i64; 14],
    }
    unsafe extern "C" {
        fn getrusage(who: i32, usage: *mut Rusage) -> i32;
    }
    let mut usage = std::mem::MaybeUninit::<Rusage>::zeroed();
    unsafe { getrusage(0, usage.as_mut_ptr()) };
    let usage = unsafe { usage.assume_init() };
    let ms = |t: [i64; 2]| t[0] as f64 * 1000.0 + (t[1] & 0xffff_ffff) as f64 / 1000.0;
    ms(usage.utime) + ms(usage.stime)
}

/// Process CPU time (user + kernel) in ms.
#[cfg(windows)]
fn cpu_ms() -> f64 {
    unsafe extern "system" {
        fn GetCurrentProcess() -> isize;
        fn GetProcessTimes(
            process: isize,
            creation: *mut u64,
            exit: *mut u64,
            kernel: *mut u64,
            user: *mut u64,
        ) -> i32;
    }
    let [mut creation, mut exit, mut kernel, mut user] = [0_u64; 4];
    unsafe {
        GetProcessTimes(
            GetCurrentProcess(),
            &mut creation,
            &mut exit,
            &mut kernel,
            &mut user,
        )
    };
    (kernel + user) as f64 / 10_000.0
}

fn median(values: &mut [f64]) -> f64 {
    values.sort_by(f64::total_cmp);
    let value = values.get(values.len() / 2).copied().unwrap_or(0.0);
    (value * 10.0).round() / 10.0
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let args: Vec<String> = std::env::args().skip(1).collect();
    let [docx, request_path, rest @ ..] = args.as_slice() else {
        return Err("usage: type_probe <docx> <request.json> [keys] [latin] [cjk]".into());
    };
    let keys: usize = rest.first().map_or(Ok(20), |value| value.parse())?;
    let fonts = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
    let latin = rest.get(1).map_or_else(
        || fonts.join("crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf"),
        Into::into,
    );
    let cjk = rest.get(2).map_or_else(
        || fonts.join("packages/fonts-cjk/assets/NotoSansTC-Regular.otf"),
        Into::into,
    );
    let wire: Value = serde_json::from_slice(&std::fs::read(request_path)?)?;
    let request = wire["request"].to_string();
    let mut out = Map::new();
    out.insert("file".into(), json!(docx.rsplit('/').next()));

    let engine = EngineSession::new(7);
    let bytes = std::fs::read(docx)?;
    let started = Instant::now();
    let envelope = parse_docx_for_edit(&bytes)?;
    engine.set_media(package_media(&envelope));
    seed_parsed_docx(engine.doc(), envelope)?;
    out.insert("openMs".into(), json!(started.elapsed().as_millis()));
    out.insert("memOpen".into(), mark());
    docx_layout::clear_measure_fonts();
    let ids = [
        docx_layout::register_measure_font(&std::fs::read(latin)?).map_err(|_| "latin font")?,
        docx_layout::register_measure_font(&std::fs::read(cjk)?).map_err(|_| "cjk font")?,
    ];
    if ids != [0, 1] {
        return Err(format!("font ids {ids:?} differ from the request's [0, 1]").into());
    }
    out.insert("memFonts".into(), mark());

    let started = Instant::now();
    let layout = engine.layout_document_with_regions_retained_json(&request)?;
    out.insert("layoutMs".into(), json!(started.elapsed().as_millis()));
    out.insert("pages".into(), json!(engine.stats().retained_pages));
    // What the editor's worker host sends (encodeDisplayListFrameExtras).
    let mut extras = Map::new();
    if let Some(headers_footers) = serde_json::from_str::<Value>(&layout)?.get("headersFooters") {
        extras.insert("headersFooters".into(), headers_footers.clone());
    }
    extras.insert(
        "fontChains".into(),
        wire["request"]["measurement"]["fontChains"].clone(),
    );
    let extras = Value::Object(extras).to_string();
    drop(layout);
    out.insert("memLayout".into(), mark());
    let started = Instant::now();
    let first = engine.build_display_list_frame(&extras, 0)?;
    out.insert("firstFrameMs".into(), json!(started.elapsed().as_millis()));
    out.insert("firstFrameMB".into(), json!(mb(first.len())));
    drop(first);
    out.insert("memFrame".into(), mark());
    // The retained display list's share: a clone allocates what it holds.
    let live = LIVE.load(Ordering::Relaxed);
    let list = engine
        .with_display_list(Clone::clone)
        .ok_or("no display list")?;
    let list_bytes = LIVE.load(Ordering::Relaxed) - live;
    let (mut primitives, mut glyphs) = (0, 0);
    for page in &list.pages {
        let regions = [&page.header, &page.footer];
        let all = page
            .primitives
            .iter()
            .chain(regions.into_iter().flatten().flat_map(|r| &r.primitives))
            .chain(page.note_areas.iter().flat_map(|a| &a.primitives));
        for primitive in all {
            primitives += 1;
            if let docx_layout::display_list::Primitive::GlyphRun(run) = primitive {
                glyphs += run.glyphs.len();
            }
        }
    }
    drop(list);
    PEAK.store(LIVE.load(Ordering::Relaxed), Ordering::Relaxed);
    out.insert(
        "displayList".into(),
        json!({"mb": mb(list_bytes), "primitives": primitives, "glyphs": glyphs,
            "primitiveBytes": std::mem::size_of::<docx_layout::display_list::Primitive>()}),
    );

    let paragraphs: Vec<_> = engine
        .doc()
        .paragraphs("body")?
        .into_iter()
        .filter(|paragraph| {
            paragraph.text.chars().any(|c| !c.is_whitespace())
                && !paragraph
                    .text
                    .split_once(". ")
                    .is_some_and(|(head, _)| head.chars().all(|c| c.is_ascii_digit()))
        })
        .collect();
    let targets = [
        (
            "first",
            paragraphs.first().ok_or("no body text")?.para_id.clone(),
        ),
        (
            "last",
            paragraphs.last().ok_or("no body text")?.para_id.clone(),
        ),
    ];
    let mut epoch = engine.stats().frame_epoch;
    for (label, para_id) in targets {
        let start = engine.doc().paragraph_mark_position(&para_id)?.index;
        let mut rows: Vec<Map<String, Value>> = Vec::new();
        let before = engine.stats();
        for (key, head) in (0..keys).zip(start..) {
            let text = if key % 6 == 5 { " " } else { "a" };
            let cpu = cpu_ms();
            let started = Instant::now();
            engine.doc().insert_text(
                &EditCtx::local("", ""),
                Position::new("body", head),
                text,
                FormatPolicy::Inherit,
            )?;
            let clock = Instant::now();
            let (frame, profile) = engine.apply_and_layout_profiled("body", epoch, &mut || {
                clock.elapsed().as_secs_f64() * 1000.0
            })?;
            epoch = engine.stats().frame_epoch;
            let mut row = match serde_json::to_value(profile)? {
                Value::Object(map) => map,
                _ => unreachable!(),
            };
            row.insert(
                "engineMs".into(),
                json!(started.elapsed().as_secs_f64() * 1000.0),
            );
            row.insert("engineCpuMs".into(), json!(cpu_ms() - cpu));
            row.insert("frameKB".into(), json!(frame.len() as f64 / 1024.0));
            rows.push(row);
        }
        let after = engine.stats();
        let mut medians = Map::new();
        for key in rows
            .first()
            .map(|row| row.keys().cloned().collect::<Vec<_>>())
            .unwrap_or_default()
        {
            let mut values: Vec<f64> = rows.iter().filter_map(|row| row[&key].as_f64()).collect();
            medians.insert(key, json!(median(&mut values)));
        }
        medians.insert(
            "fastPathKeys".into(),
            json!(after.region_fast_path_hits - before.region_fast_path_hits),
        );
        medians.insert("fallback".into(), json!(after.region_fast_path_fallback));
        medians.insert(
            "rebuiltDisplayPages".into(),
            json!(after.rebuilt_display_pages - before.rebuilt_display_pages),
        );
        out.insert(format!("keys_{label}"), Value::Object(medians));
        out.insert(format!("memKeys_{label}"), mark());
    }
    println!("{}", Value::Object(out));
    Ok(())
}
