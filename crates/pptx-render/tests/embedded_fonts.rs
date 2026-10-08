use std::alloc::{GlobalAlloc, Layout, System};
use std::cell::Cell;

use ooxml_text::{FontStore, decode_embedded_font, shape};
use pptx_edit::DeckSession;
use pptx_parse::EmbeddedFont;
use pptx_render::{EmbeddedFace, PositionedTextLine, Primitive, SlideRenderer};

const DECK: &[u8] = include_bytes!("fixtures/line-spacing.pptx");
/// Lato as Google Slides embeds it (EOT, MicroType Express).
const LATO: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/Lato-regular.fntdata");
const LATO_BOLD: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/Lato-bold.fntdata");
const SANS: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
const TEXT: &str = "Embedded as text it wraps later";
const GUID_PART: &str = "{001B70DC-AA60-4AD5-90EC-18A0948E1EAE}.odttf";

const FONT_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font";

/// `line-spacing.pptx` with shape 4's first run in Lato saying [`TEXT`],
/// and `faces` (slot, part file name, bytes) embedded for Lato.
fn deck(faces: &[(&str, &str, &[u8])], bold: bool) -> Vec<u8> {
    let mut slots = String::new();
    let mut relationships = String::new();
    let mut parts = Vec::new();
    for (index, (slot, name, bytes)) in faces.iter().enumerate() {
        slots += &format!(r#"<p:{slot} r:id="rIdFont{index}"/>"#);
        relationships += &format!(
            r#"<Relationship Id="rIdFont{index}" Type="{FONT_RELATIONSHIP}" Target="fonts/{name}"/>"#
        );
        if !bytes.is_empty() {
            parts.push((format!("ppt/fonts/{name}"), bytes.to_vec()));
        }
    }
    let list = format!(r#"<p:embeddedFont><p:font typeface="Lato"/>{slots}</p:embeddedFont>"#);
    build(&list, &relationships, parts, bold)
}

/// `count` typefaces, each with all four slots on `bytes`: one part for all
/// of them, or a part of its own each.
fn many(count: usize, distinct: bool, bytes: &[u8]) -> Vec<u8> {
    let parts_count = if distinct { count } else { 1 };
    let relationships: String = (0..parts_count)
        .map(|index| {
            format!(
                r#"<Relationship Id="rIdFont{index}" Type="{FONT_RELATIONSHIP}" Target="fonts/font{index}.fntdata"/>"#
            )
        })
        .collect();
    let list: String = (0..count)
        .map(|index| {
            let id = if distinct { index } else { 0 };
            let slots: String = ["regular", "bold", "italic", "boldItalic"]
                .iter()
                .map(|slot| format!(r#"<p:{slot} r:id="rIdFont{id}"/>"#))
                .collect();
            let typeface = if index == 0 {
                "Lato".to_owned()
            } else {
                format!("Face{index}")
            };
            format!(r#"<p:embeddedFont><p:font typeface="{typeface}"/>{slots}</p:embeddedFont>"#)
        })
        .collect();
    let parts = (0..parts_count)
        .map(|index| (format!("ppt/fonts/font{index}.fntdata"), bytes.to_vec()))
        .collect();
    build(&list, &relationships, parts, false)
}

fn build(list: &str, relationships: &str, fonts: Vec<(String, Vec<u8>)>, bold: bool) -> Vec<u8> {
    let mut parts = ooxml_opc::unzip_parts(DECK).unwrap();
    parts.extend(fonts);
    let list = format!(r#"<p:embeddedFontLst>{list}</p:embeddedFontLst></p:presentation>"#);
    let b = if bold { r#" b="1""# } else { "" };
    let run = format!(
        r#"<a:rPr sz="3200"{b}><a:solidFill><a:srgbClr val="2040B0"/></a:solidFill><a:latin typeface="Lato"/></a:rPr><a:t>{TEXT}</a:t>"#
    );
    for (part, from, to) in [
        ("ppt/presentation.xml", "</p:presentation>", list.as_str()),
        (
            "ppt/_rels/presentation.xml.rels",
            "</Relationships>",
            &format!("{relationships}</Relationships>"),
        ),
        (
            "[Content_Types].xml",
            "</Types>",
            r#"<Default Extension="fntdata" ContentType="application/x-fontdata"/><Default Extension="odttf" ContentType="application/vnd.openxmlformats-officedocument.obfuscatedFont"/></Types>"#,
        ),
        (
            "ppt/slides/slide1.xml",
            r#"<a:rPr sz="3200"><a:solidFill><a:srgbClr val="2040B0"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>Exact 72 pt</a:t>"#,
            run.as_str(),
        ),
    ] {
        let (_, bytes) = parts.iter_mut().find(|(path, _)| path == part).unwrap();
        let xml = String::from_utf8(bytes.clone()).unwrap();
        assert!(xml.contains(from), "{part}");
        *bytes = xml.replacen(from, to, 1).into_bytes();
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
}

struct Laid {
    faces: Vec<EmbeddedFace>,
    lines: Vec<PositionedTextLine>,
}

fn lay_out(deck: &[u8]) -> Laid {
    lay_out_refusing(deck, false)
}

/// Lays out shape 4, after the page refused every embedded face if `refuse`.
fn lay_out_refusing(deck: &[u8], refuse: bool) -> Laid {
    let session = DeckSession::open(deck, 8_021).unwrap();
    let mut renderer = SlideRenderer::new();
    renderer.register_font("Arial", false, false, SANS).unwrap();
    let faces = renderer.register_embedded_fonts(session.package());
    if refuse {
        for face in &faces {
            renderer.refuse_embedded_font(face.font_id);
        }
    }
    let lines = renderer
        .layout_slide(session.package(), &session.snapshot().unwrap(), 0)
        .unwrap()
        .display_list
        .primitives
        .into_iter()
        .find_map(|primitive| match primitive {
            Primitive::TextBox {
                object_id: 4,
                lines,
                ..
            } => Some(lines),
            _ => None,
        })
        .unwrap();
    Laid { faces, lines }
}

fn texts(lines: &[PositionedTextLine]) -> Vec<String> {
    lines
        .iter()
        .map(|line| line.runs.iter().map(|run| run.text.as_str()).collect())
        .collect()
}

fn lato_ttf() -> Vec<u8> {
    decode_embedded_font(LATO, "ppt/fonts/Lato-regular.fntdata").unwrap()
}

#[test]
fn the_presentation_lists_each_embedded_face_with_its_part() {
    let deck = include_bytes!("fixtures/embedded-lato.pptx");
    let session = DeckSession::open(deck, 8_021).unwrap();
    assert_eq!(
        session.package().presentation.embedded_fonts,
        [
            EmbeddedFont {
                typeface: "Lato".into(),
                bold: false,
                italic: false,
                part_path: "ppt/fonts/Lato-regular.fntdata".into(),
            },
            EmbeddedFont {
                typeface: "Lato".into(),
                bold: true,
                italic: false,
                part_path: "ppt/fonts/Lato-bold.fntdata".into(),
            },
        ]
    );
}

#[test]
fn an_embedded_face_measures_and_wraps_the_text_it_names() {
    let fallback = lay_out(&deck(&[], false));
    let embedded = lay_out(&deck(&[("regular", "Lato-regular.fntdata", LATO)], false));
    assert!(fallback.faces.is_empty());
    let [face] = embedded.faces.as_slice() else {
        panic!("one face: {:?}", embedded.faces);
    };
    assert_eq!(
        (face.typeface.as_str(), face.bold, face.italic),
        ("Lato", false, false)
    );
    // Display lists name the face by its alias, so the page's own "Lato"
    // (or "Fustat") is never restyled by the deck.
    assert!(face.family.starts_with("bo-embedded-"), "{}", face.family);

    assert_eq!(
        texts(&fallback.lines),
        ["Embedded as ", "text it wraps later", "Second line"]
    );
    assert_eq!(
        texts(&embedded.lines),
        ["Embedded as text ", "it wraps later", "Second line"]
    );
    // Lato's own advances and kerning, at the run's 32 pt.
    let mut store = FontStore::new();
    let lato = store.register(lato_ttf()).unwrap();
    for line in &embedded.lines[..2] {
        let run = &line.runs[0];
        let text = run.text.as_str();
        assert_eq!(run.font_id, face.font_id);
        assert_eq!(run.font_family, face.family);
        let width: f32 = shape(&store, lato, text, run.font_size_px, &[])
            .unwrap()
            .iter()
            .map(|glyph| glyph.x_advance)
            .sum();
        assert!(
            (run.width - width).abs() < 0.01,
            "{text}: {} vs {width}",
            run.width
        );
    }
    // Lato's own ascent places the first baseline.
    assert!((embedded.lines[0].baseline - fallback.lines[0].baseline - 1.0225).abs() < 0.001);
    // The control run in Arial keeps the host face.
    assert_eq!(embedded.lines[2].runs, fallback.lines[2].runs);
}

#[test]
fn plain_and_obfuscated_parts_lay_out_as_the_compressed_one() {
    let compressed = lay_out(&deck(&[("regular", "Lato-regular.fntdata", LATO)], false));
    let plain = lay_out(&deck(
        &[("regular", "Lato-regular.ttf", &lato_ttf())],
        false,
    ));
    let key = [
        0xAE, 0x1E, 0x8E, 0x94, 0xA0, 0x18, 0xEC, 0x90, 0xD5, 0x4A, 0x60, 0xAA, 0xDC, 0x70, 0x1B,
        0x00,
    ];
    let mut obfuscated = lato_ttf();
    for (index, byte) in obfuscated.iter_mut().take(32).enumerate() {
        *byte ^= key[index % 16];
    }
    let obfuscated = lay_out(&deck(&[("regular", GUID_PART, &obfuscated)], false));
    assert_eq!(plain.faces, compressed.faces);
    assert_eq!(plain.lines, compressed.lines);
    assert_eq!(obfuscated.faces, compressed.faces);
    assert_eq!(obfuscated.lines, compressed.lines);
}

#[test]
fn a_broken_or_missing_part_keeps_the_host_face() {
    let fallback = lay_out(&deck(&[], false));
    let mut truncated = LATO[..LATO.len() / 2].to_vec();
    let size = truncated.len() as u32;
    truncated[..4].copy_from_slice(&size.to_le_bytes());
    for faces in [
        [("regular", "Lato-regular.fntdata", &truncated[..])],
        [("regular", "Lato-regular.fntdata", &b"not a font at all"[..])],
        [("regular", "missing.fntdata", &[][..])],
    ] {
        let laid = lay_out(&deck(&faces, false));
        assert!(laid.faces.is_empty());
        assert_eq!(laid.lines, fallback.lines);
    }
}

#[test]
fn each_style_takes_its_own_face_and_a_lone_style_serves_the_family() {
    let both = [
        ("regular", "Lato-regular.fntdata", LATO),
        ("bold", "Lato-bold.fntdata", LATO_BOLD),
    ];
    let bold = lay_out(&deck(&both, true));
    let bold_face = bold.faces.iter().find(|face| face.bold).unwrap();
    assert_eq!(bold.lines[0].runs[0].font_id, bold_face.font_id);
    let regular = lay_out(&deck(&both, false));
    let regular_face = regular.faces.iter().find(|face| !face.bold).unwrap();
    assert_eq!(regular.lines[0].runs[0].font_id, regular_face.font_id);

    // Only the bold face embedded: regular text still draws in Lato Bold,
    // the face the browser picks for the name.
    let lone = lay_out(&deck(&[("bold", "Lato-bold.fntdata", LATO_BOLD)], false));
    assert_eq!(lone.lines[0].runs[0].font_id, lone.faces[0].font_id);
}

#[test]
fn a_host_face_registered_later_leaves_the_embedded_one_in_place() {
    let deck = deck(&[("regular", "Lato-regular.fntdata", LATO)], false);
    let session = DeckSession::open(&deck, 8_021).unwrap();
    let mut renderer = SlideRenderer::new();
    let faces = renderer.register_embedded_fonts(session.package());
    renderer.register_font("Arial", false, false, SANS).unwrap();
    renderer.register_font("Lato", false, false, SANS).unwrap();
    let slide = renderer
        .layout_slide(session.package(), &session.snapshot().unwrap(), 0)
        .unwrap();
    let lines = slide
        .display_list
        .primitives
        .iter()
        .find_map(|primitive| match primitive {
            Primitive::TextBox {
                object_id: 4,
                lines,
                ..
            } => Some(lines),
            _ => None,
        })
        .unwrap();
    assert_eq!(lines[0].runs[0].font_id, faces[0].font_id);
    // The host face, not the embedded one, stays the fallback.
    assert_ne!(
        renderer.fallback_font().map(|id| id.to_u32()),
        Some(faces[0].font_id)
    );
}

/// Security regression (REVIEW1 S2).
#[test]
fn a_face_the_page_refuses_leaves_layout_too() {
    let fallback = lay_out(&deck(&[], false));
    let refused = lay_out_refusing(
        &deck(&[("regular", "Lato-regular.fntdata", LATO)], false),
        true,
    );
    assert_eq!(refused.faces.len(), 1);
    assert_eq!(refused.lines, fallback.lines);
}

const EXPAND: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/hostile-expand.fntdata");
const RLE: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/hostile-rle.fntdata");

/// Bytes this thread allocates: the security tests bound the work a deck's
/// fonts take by allocation, not by time.
struct Counting;

thread_local! {
    static ALLOCATED: Cell<usize> = const { Cell::new(0) };
}

fn count(bytes: usize) {
    let _ = ALLOCATED.try_with(|total| total.set(total.get() + bytes));
}

unsafe impl GlobalAlloc for Counting {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        count(layout.size());
        unsafe { System.alloc(layout) }
    }

    unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
        unsafe { System.dealloc(pointer, layout) }
    }

    unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, size: usize) -> *mut u8 {
        count(size.saturating_sub(layout.size()));
        unsafe { System.realloc(pointer, layout, size) }
    }
}

#[global_allocator]
static ALLOCATOR: Counting = Counting;

const MIB: usize = 1024 * 1024;

/// Registers `deck`'s embedded fonts: the faces, and the bytes allocated
/// doing it.
fn register(deck: &[u8]) -> (Vec<EmbeddedFace>, usize) {
    let session = DeckSession::open(deck, 8_021).unwrap();
    let mut renderer = SlideRenderer::new();
    renderer.register_font("Arial", false, false, SANS).unwrap();
    let before = ALLOCATED.with(Cell::get);
    let faces = renderer.register_embedded_fonts(session.package());
    (faces, ALLOCATED.with(Cell::get) - before)
}

#[test]
fn a_part_named_by_many_slots_is_decoded_once() {
    // Security regression (REVIEW1 M2): 25 typefaces x 4 slots on one
    // 150-byte part that unpacks 16 MiB before failing. Decoded per slot,
    // that was 100 unpacks (1.6 GiB); decoded once, one.
    let (faces, allocated) = register(&many(25, false, EXPAND));
    assert!(faces.is_empty());
    assert!(allocated < 24 * MIB, "{allocated} bytes");
    // One good part named by 1,000 typefaces: one face, every slot served.
    let laid = lay_out(&many(1_000, false, LATO));
    assert_eq!(laid.faces.len(), 1);
    assert_eq!(laid.lines[0].runs[0].font_id, laid.faces[0].font_id);
}

#[test]
fn failing_parts_spend_the_deck_budget_before_their_work() {
    // Security regression (REVIEW1 M2): 20 distinct copies, each charged its
    // declared 16 MiB before it unpacks, so the 64 MiB budget stops the work
    // after three (charged only on success, all twenty unpacked: 320 MiB).
    let (faces, allocated) = register(&many(20, true, EXPAND));
    assert!(faces.is_empty());
    assert!(allocated < 56 * MIB, "{allocated} bytes");
    // Security regression (REVIEW1 M2): distinct good parts register up to
    // the face cap.
    let (faces, _) = register(&many(70, true, LATO));
    assert_eq!(faces.len(), pptx_render::MAX_EMBEDDED_FACES);
}

#[test]
fn run_length_expansion_spends_the_deck_budget_as_it_grows() {
    // Security regression (REVIEW2 N1): 200 distinct 145-byte parts whose
    // run-length stage expands each to 32 MiB. Charged as it grows, two
    // expand before the 64 MiB budget runs out; charged by declared length
    // only (393 KB each), about 170 did (5 GiB).
    let (faces, allocated) = register(&many(200, true, RLE));
    assert!(faces.is_empty());
    assert!(allocated < 80 * MIB, "{allocated} bytes");
}
