use ooxml_text::embedded_font::MAX_EMBEDDED_FONT_BYTES;
use ooxml_text::{EmbeddedFontError, FontStore, decode_embedded_font};
use skrifa::instance::{LocationRef, Size};
use skrifa::outline::pen::ControlBoundsPen;
use skrifa::outline::{DrawSettings, Engine, HintingInstance, HintingOptions, Target};
use skrifa::raw::TableProvider;
use skrifa::{FontRef, GlyphId, MetadataProvider};

/// Lato Regular as Google Slides embeds it: EOT 2.2, MicroType Express.
const LATO: &[u8] = include_bytes!("fonts/Lato-regular.fntdata");
const PART: &str = "ppt/fonts/Lato-regular.fntdata";

fn lato() -> Vec<u8> {
    decode_embedded_font(LATO, PART).expect("Lato decodes")
}

/// An EOT 2.2 container around `data` with `flags`.
fn eot(data: &[u8], flags: u32) -> Vec<u8> {
    let mut header = LATO[..LATO.len() - u32_le(LATO, 4) as usize].to_vec();
    header[12..16].copy_from_slice(&flags.to_le_bytes());
    header[4..8].copy_from_slice(&(data.len() as u32).to_le_bytes());
    let mut container = header;
    container.extend(data);
    let size = container.len() as u32;
    container[..4].copy_from_slice(&size.to_le_bytes());
    container
}

fn u32_le(bytes: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes(bytes[offset..offset + 4].try_into().unwrap())
}

#[test]
fn a_micro_type_express_part_decodes_to_the_whole_font() {
    let font = lato();
    let face = FontRef::new(&font).unwrap();
    let count = face.maxp().unwrap().num_glyphs();
    assert_eq!(count, 272);
    // Each rebuilt glyph's box starts where its side bearing says.
    let (hmtx, loca, glyf) = (
        face.hmtx().unwrap(),
        face.loca(None).unwrap(),
        face.glyf().unwrap(),
    );
    let mut outlined = 0;
    for gid in 0..count {
        let gid = GlyphId::new(gid.into());
        if let Some(glyph) = loca.get_glyf(gid, &glyf).unwrap() {
            assert_eq!(hmtx.side_bearing(gid), Some(glyph.x_min()), "glyph {gid}");
            outlined += 1;
        }
    }
    assert_eq!(outlined, 268);
    // Every glyph program, rejoined from the push and code blocks, runs.
    let outlines = face.outline_glyphs();
    let options = HintingOptions {
        engine: Engine::Interpreter,
        target: Target::Mono,
    };
    let hinter =
        HintingInstance::new(&outlines, Size::new(13.0), LocationRef::default(), options).unwrap();
    for gid in 0..count {
        let glyph = outlines.get(GlyphId::new(gid.into())).unwrap();
        glyph
            .draw(
                DrawSettings::hinted(&hinter, true),
                &mut ControlBoundsPen::default(),
            )
            .unwrap_or_else(|error| panic!("glyph {gid}: {error}"));
    }

    let mut store = FontStore::new();
    let id = store.register(font).unwrap();
    assert_eq!(store.metrics(id).unwrap().units_per_em, 2000);
    assert_eq!(store.advance_width(id, 'A').unwrap(), Some(1360.0));
}

#[test]
fn plain_xor_encrypted_and_uncompressed_eot_forms_give_the_same_font() {
    let font = lato();
    assert_eq!(decode_embedded_font(&font, PART).unwrap(), font);
    assert_eq!(decode_embedded_font(&eot(&font, 0), PART).unwrap(), font);
    let encrypted: Vec<u8> = font.iter().map(|byte| byte ^ 0x50).collect();
    assert_eq!(
        decode_embedded_font(&eot(&encrypted, 0x1000_0000), PART).unwrap(),
        font
    );
    let mtx = &LATO[LATO.len() - u32_le(LATO, 4) as usize..];
    let encrypted: Vec<u8> = mtx.iter().map(|byte| byte ^ 0x50).collect();
    assert_eq!(
        decode_embedded_font(&eot(&encrypted, 0x1000_0004), PART).unwrap(),
        font
    );
}

#[test]
fn an_obfuscated_part_is_read_with_the_guid_its_name_carries() {
    let font = lato();
    let key = [
        0xAE, 0x1E, 0x8E, 0x94, 0xA0, 0x18, 0xEC, 0x90, 0xD5, 0x4A, 0x60, 0xAA, 0xDC, 0x70, 0x1B,
        0x00,
    ];
    let mut obfuscated = font.clone();
    for (index, byte) in obfuscated.iter_mut().take(32).enumerate() {
        *byte ^= key[index % 16];
    }
    let name = "ppt/fonts/{001B70DC-AA60-4AD5-90EC-18A0948E1EAE}.odttf";
    assert_eq!(decode_embedded_font(&obfuscated, name).unwrap(), font);
    assert_eq!(
        decode_embedded_font(&obfuscated, "ppt/fonts/font1.odttf"),
        Err(EmbeddedFontError::Unrecognized)
    );
}

#[test]
fn damaged_parts_fail_without_panicking() {
    let header = LATO.len() - u32_le(LATO, 4) as usize;
    for length in [
        header,
        header + 9,
        header + 200,
        LATO.len() / 2,
        LATO.len() - 1,
    ] {
        let mut cut = LATO[..length].to_vec();
        let size = cut.len() as u32;
        cut[..4].copy_from_slice(&size.to_le_bytes());
        let data = (length - header) as u32;
        cut[4..8].copy_from_slice(&data.to_le_bytes());
        assert!(decode_embedded_font(&cut, PART).is_err(), "cut at {length}");
    }
    for position in (header..LATO.len()).step_by(997) {
        let mut flipped = LATO.to_vec();
        flipped[position] ^= 0x5A;
        if let Ok(font) = decode_embedded_font(&flipped, PART) {
            assert!(font.len() <= MAX_EMBEDDED_FONT_BYTES);
        }
    }
}
