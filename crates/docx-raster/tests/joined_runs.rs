//! A joined glyph run (`clusterRuns`) rasterizes exactly like the one-cluster
//! runs it was joined from, run transforms (character scale) included. (The
//! rasterizer refuses w14 effects, so gradient fills are checked in
//! docx-layout's join test only.)

use docx_layout::display_list::DisplayList;
use docx_layout::measure_blocks::{MeasurementConfig, measure_blocks};
use docx_layout::types::{Input, LayoutBlock, MeasuredBlock};
use docx_raster::{FontChains, ImageMap, RenderResources, render_png};
use ooxml_text::FontStore;
use serde_json::{Value, json};

const LIBERATION: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

/// The display list with every joined run split back into one run per
/// glyph cluster (clusters are UTF-8 byte offsets into the run's text).
fn split_cluster_runs(list: &DisplayList) -> DisplayList {
    let mut value = serde_json::to_value(list).unwrap();
    for page in value["pages"].as_array_mut().unwrap() {
        let primitives = page["primitives"].as_array_mut().unwrap();
        *primitives = std::mem::take(primitives)
            .into_iter()
            .flat_map(|primitive| {
                if primitive["clusterRuns"] != json!(true) {
                    return vec![primitive];
                }
                let text = primitive["text"].as_str().unwrap();
                let glyphs = primitive["glyphs"].as_array().unwrap();
                let mut starts: Vec<u64> = Vec::new();
                for glyph in glyphs {
                    let cluster = glyph["cluster"].as_u64().unwrap();
                    if !starts.contains(&cluster) {
                        starts.push(cluster);
                    }
                }
                starts
                    .iter()
                    .enumerate()
                    .map(|(index, &start)| {
                        let end = starts
                            .get(index + 1)
                            .map_or(text.len(), |&end| end as usize);
                        let mut piece = primitive.clone();
                        piece.as_object_mut().unwrap().remove("clusterRuns");
                        piece["text"] = json!(text[start as usize..end]);
                        piece["glyphs"] = glyphs
                            .iter()
                            .filter(|glyph| glyph["cluster"].as_u64() == Some(start))
                            .map(|glyph| {
                                let mut glyph = glyph.clone();
                                glyph["cluster"] = json!(0);
                                glyph
                            })
                            .collect();
                        piece
                    })
                    .collect()
            })
            .collect();
    }
    serde_json::from_value(value).unwrap()
}

#[test]
fn joined_glyph_runs_rasterize_like_their_cluster_runs() {
    docx_layout::clear_measure_fonts();
    let measure_id = docx_layout::register_measure_font(LIBERATION).unwrap();
    let mut fonts = FontStore::new();
    let raster_id = fonts.register(LIBERATION.to_vec()).unwrap();
    assert_eq!(measure_id, raster_id.to_u32());
    let chains = json!({ "liberation sans|0|0": [measure_id] });
    let lorem = "The committee reviews every application against the published \
                 criteria before the spring deadline. Late submissions wait. ";
    let paragraphs = [
        ("justify", lorem.repeat(2), json!({})),
        (
            "left",
            "Scaled office text with fluffy affirmations and first flights.".to_owned(),
            json!({ "horizontalScale": 150.0 }),
        ),
        (
            "left",
            "Condensed text office file with fifty fine flags".to_owned(),
            json!({ "horizontalScale": 80.0 }),
        ),
        (
            "left",
            "Cafe\u{301} nai\u{308}ve re\u{301}sume\u{301} office first fluff waffle".to_owned(),
            json!({}),
        ),
        (
            "left",
            "Letter spaced words in the official office file".to_owned(),
            json!({ "letterSpacing": 2.0 }),
        ),
    ];
    let mut pm = 0_usize;
    let mut blocks: Vec<LayoutBlock> = Vec::new();
    for (index, (alignment, text, formatting)) in paragraphs.iter().enumerate() {
        let start = pm;
        let len = text.encode_utf16().count();
        let mut run = json!({
            "kind": "text",
            "text": text,
            "fontFamily": "Liberation Sans",
            "fontSize": 11.0,
            "pmStart": start + 1,
            "pmEnd": start + 1 + len
        });
        run.as_object_mut()
            .unwrap()
            .extend(formatting.as_object().unwrap().clone());
        pm = start + len + 2;
        blocks.push(
            serde_json::from_value(json!({
                "kind": "paragraph",
                "id": format!("p{index}"),
                "runs": [run],
                "attrs": {
                    "alignment": alignment,
                    "defaultFontFamily": "Liberation Sans",
                    "defaultFontSize": 11.0
                },
                "pmStart": start,
                "pmEnd": pm
            }))
            .unwrap(),
        );
    }
    let config = MeasurementConfig {
        font_chains: serde_json::from_value(chains.clone()).unwrap(),
        defaults: json!({ "fontSize": 11.0, "fontFamily": "Liberation Sans" }),
        compat: Value::Null,
        authoritative_shaping: true,
    };
    let extents = measure_blocks(&mut blocks, 400.0, &config).unwrap();
    let mut input = Input {
        measured: blocks
            .into_iter()
            .zip(extents)
            .map(|(block, measure)| MeasuredBlock { block, measure })
            .collect(),
        options: serde_json::from_value(json!({
            "pageSize": { "w": 500.0, "h": 400.0 },
            "margins": { "top": 50.0, "right": 50.0, "bottom": 50.0, "left": 50.0 }
        }))
        .unwrap(),
    };
    let layout = docx_layout::compute_layout_input(&mut input).unwrap();
    let extras = json!({ "fontChains": chains }).to_string();
    let joined =
        docx_layout::build_display_list_value_from_resident(&input, &layout, &extras).unwrap();
    docx_layout::clear_measure_fonts();
    let split = split_cluster_runs(&joined);

    let glyph_runs = |list: &DisplayList| -> Vec<Value> {
        serde_json::to_value(list).unwrap()["pages"]
            .as_array()
            .unwrap()
            .iter()
            .flat_map(|page| page["primitives"].as_array().unwrap().clone())
            .filter(|primitive| primitive["kind"] == "glyphRun")
            .collect()
    };
    let joined_runs = glyph_runs(&joined);
    assert!(
        joined_runs
            .iter()
            .any(|run| run["clusterRuns"] == json!(true))
    );
    assert!(
        joined_runs
            .iter()
            .any(|run| run.get("horizontalScale").is_some())
    );

    let chains = FontChains::from([("liberation sans|0|0".to_owned(), vec![raster_id])]);
    let images = ImageMap::new();
    let resources = RenderResources::new(&fonts, &chains, &images);
    for page in 0..joined.pages.len() {
        assert!(
            render_png(&joined, page, &resources).unwrap()
                == render_png(&split, page, &resources).unwrap(),
            "page {page} rasterizes differently once its runs are joined"
        );
    }
}
