//! Security regression (parity-px N1): pruning a slide's timing after a
//! delete repeated a full pass per link of a chain of nodes each waiting on
//! the one before (`p:tn`), so a crafted deck made a save quadratic. The
//! prune is one pass plus a worklist now. Its own test binary, so the counting
//! allocator sees only this test: doubling the chain must about double the
//! save's allocations, where the repeated passes quadrupled them.

use std::alloc::{GlobalAlloc, Layout, System};
use std::sync::atomic::{AtomicUsize, Ordering};

use pptx_edit::{DeckSession, EditCtx};

struct Counting;

static ALLOCATIONS: AtomicUsize = AtomicUsize::new(0);

unsafe impl GlobalAlloc for Counting {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        ALLOCATIONS.fetch_add(1, Ordering::SeqCst);
        unsafe { System.alloc(layout) }
    }

    unsafe fn dealloc(&self, ptr: *mut u8, layout: Layout) {
        unsafe { System.dealloc(ptr, layout) }
    }

    unsafe fn realloc(&self, ptr: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        ALLOCATIONS.fetch_add(1, Ordering::SeqCst);
        unsafe { System.realloc(ptr, layout, new_size) }
    }
}

#[global_allocator]
static ALLOCATOR: Counting = Counting;

const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const REL: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const P: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";

fn shape(id: u32, name: &str) -> String {
    format!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="{id}" name="{name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp>"#
    )
}

/// A slide whose timing is a chain of `links` steps, each starting when the
/// one before ends; the first shows shape 2, the rest shape 3.
fn chained_deck(links: usize) -> Vec<u8> {
    let mut steps = String::new();
    for link in 0..links {
        let id = 100 + link;
        let start = match link {
            0 => r#"<p:cond delay="0"/>"#.to_owned(),
            _ => format!(
                r#"<p:cond evt="onEnd" delay="0"><p:tn val="{}"/></p:cond>"#,
                id - 1
            ),
        };
        let target = if link == 0 { 2 } else { 3 };
        steps.push_str(&format!(
            r#"<p:par><p:cTn id="{id}" fill="hold"><p:stCondLst>{start}</p:stCondLst><p:childTnLst><p:set><p:cBhvr><p:cTn id="{}" dur="1"/><p:tgtEl><p:spTgt spid="{target}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set></p:childTnLst></p:cTn></p:par>"#,
            1_000_000 + link
        ));
    }
    let tree = r#"<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>"#;
    let namespaces = format!(r#"xmlns:a="{A}" xmlns:r="{R}" xmlns:p="{P}""#);
    let xml = |body: String| {
        format!(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>{body}"#).into_bytes()
    };
    let rels = |body: String| {
        xml(format!(
            r#"<Relationships xmlns="{REL}">{body}</Relationships>"#
        ))
    };
    let parts: Vec<(String, Vec<u8>)> = vec![
        (
            "[Content_Types].xml".into(),
            xml(r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/></Types>"#.to_owned()),
        ),
        (
            "_rels/.rels".into(),
            rels(format!(r#"<Relationship Id="rId1" Type="{R}/officeDocument" Target="ppt/presentation.xml"/>"#)),
        ),
        (
            "ppt/presentation.xml".into(),
            xml(format!(r#"<p:presentation {namespaces}><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>"#)),
        ),
        (
            "ppt/_rels/presentation.xml.rels".into(),
            rels(format!(r#"<Relationship Id="rId1" Type="{R}/slideMaster" Target="slideMasters/slideMaster1.xml"/><Relationship Id="rId2" Type="{R}/slide" Target="slides/slide1.xml"/>"#)),
        ),
        (
            "ppt/slides/slide1.xml".into(),
            xml(format!(
                r#"<p:sld {namespaces}><p:cSld><p:spTree>{tree}{}{}</p:spTree></p:cSld><p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>{steps}</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing></p:sld>"#,
                shape(2, "First"),
                shape(3, "Rest")
            )),
        ),
        (
            "ppt/slides/_rels/slide1.xml.rels".into(),
            rels(format!(r#"<Relationship Id="rId1" Type="{R}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>"#)),
        ),
        (
            "ppt/slideLayouts/slideLayout1.xml".into(),
            xml(format!(r#"<p:sldLayout {namespaces} type="blank"><p:cSld><p:spTree>{tree}</p:spTree></p:cSld></p:sldLayout>"#)),
        ),
        (
            "ppt/slideLayouts/_rels/slideLayout1.xml.rels".into(),
            rels(format!(r#"<Relationship Id="rId1" Type="{R}/slideMaster" Target="../slideMasters/slideMaster1.xml"/>"#)),
        ),
        (
            "ppt/slideMasters/slideMaster1.xml".into(),
            xml(format!(r#"<p:sldMaster {namespaces}><p:cSld><p:spTree>{tree}</p:spTree></p:cSld><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>"#)),
        ),
        (
            "ppt/slideMasters/_rels/slideMaster1.xml.rels".into(),
            rels(format!(r#"<Relationship Id="rId1" Type="{R}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>"#)),
        ),
    ];
    ooxml_opc::rezip_parts(&parts).unwrap()
}

/// Deletes the shape the chain's first step shows and saves, returning the
/// allocations the save made and the saved slide.
fn delete_first_and_save(links: usize) -> (usize, String) {
    let session = DeckSession::open(&chained_deck(links), 4401).unwrap();
    let snapshot = session.snapshot().unwrap();
    let slide = &snapshot.slides[0];
    let first = slide
        .shapes
        .iter()
        .find(|shape| shape.name == "First")
        .unwrap();
    session
        .remove_shape(&EditCtx::local("work"), &slide.id, &first.id)
        .unwrap();
    let before = ALLOCATIONS.load(Ordering::SeqCst);
    let saved = session.save().unwrap();
    let allocations = ALLOCATIONS.load(Ordering::SeqCst) - before;
    let (_, slide) = ooxml_opc::unzip_parts(&saved)
        .unwrap()
        .into_iter()
        .find(|(path, _)| path == "ppt/slides/slide1.xml")
        .unwrap();
    (allocations, String::from_utf8(slide).unwrap())
}

#[test]
fn a_chain_of_waiting_steps_prunes_in_work_linear_in_its_length() {
    let (short, short_slide) = delete_first_and_save(1_000);
    let (long, long_slide) = delete_first_and_save(2_000);
    // Every step waited on the first, so the whole timing goes.
    for slide in [&short_slide, &long_slide] {
        assert!(!slide.contains("<p:timing"), "the chain was kept");
        assert!(slide.contains(r#"name="Rest""#));
    }
    assert!(
        long < short * 3,
        "allocations grew {short} -> {long} for twice the chain"
    );
}
