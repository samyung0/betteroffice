//! Every PPTX fixture in the repository, saved with a text box and a preset
//! shape the editor adds to its first slide, writes its slides byte for
//! byte as the save did before the chat agent's layout placeholders (the
//! Editor's condition, 2026-10-22): the pins were computed on fork commit
//! 7c43320ab, before `ShapeAdd.placeholder`. "-" marks a fixture the
//! editor does not open or that has no slide.

use std::path::PathBuf;

use pptx_edit::{DeckSession, EditCtx, PresetShapeDraft, ShapeDraft, ShapeRect, TextStyle};
use sha2::{Digest, Sha256};

const FIXTURES: &[(&str, &str)] = &[
    (
        "apps/demo/public/betteroffice-demo.pptx",
        "70039f4165965cddaf3165365ea0d2ab1d0b4e1ed2a2cbaaecb926ea16735fbf",
    ),
    (
        "crates/ooxml-drawingml/tests/fixtures/preset-adjustments.pptx",
        "16d3395663760befcf129cb74c72da400209998259a6ad51ddb722d3f63d1fef",
    ),
    (
        "crates/ooxml-opc/tests/fixtures/betteroffice-demo.pptx",
        "70039f4165965cddaf3165365ea0d2ab1d0b4e1ed2a2cbaaecb926ea16735fbf",
    ),
    (
        "crates/pptx-edit/tests/fixtures/blip-shadow.pptx",
        "0d394a58a376e5d91c88befc1162ab3e8936293b6c6f3eb3110a919a808bd753",
    ),
    (
        "crates/pptx-edit/tests/fixtures/chart-text-overflow.pptx",
        "e5f171e3b4f79041a76c603bb8434eb31f57e5fbc01961c15aa144d84646f4bc",
    ),
    (
        "crates/pptx-edit/tests/fixtures/deck-schema-v2-connectors.pptx",
        "d7683748bdedf5a1169c5fcf5237cbf2e8e40b6d8908c5b470b13c47fd5c888c",
    ),
    (
        "crates/pptx-edit/tests/fixtures/deck-schema-v2-nested-connectors.pptx",
        "b69b230428466d43ac904a7a37246bb762e61cc79f744be7c37821bdd9c4cb06",
    ),
    (
        "crates/pptx-edit/tests/fixtures/hidden-shapes.pptx",
        "817a7d4cb36174e7f416f651bc79e5de8c23b9847c0d03aac1874d33eee27c9d",
    ),
    (
        "crates/pptx-edit/tests/fixtures/metafile-tracking.pptx",
        "844f999d8e06cc9aba3faf741e61da229a2901af3c6c8aef9a704beca2eb34b3",
    ),
    (
        "crates/pptx-edit/tests/fixtures/modern-comments.pptx",
        "70039f4165965cddaf3165365ea0d2ab1d0b4e1ed2a2cbaaecb926ea16735fbf",
    ),
    (
        "crates/pptx-edit/tests/fixtures/run-spacing-shadow.pptx",
        "0d0b27bdc5f780f8e05f9e8a30595bd5559e133447942d9a73f692857f893af0",
    ),
    (
        "crates/pptx-parse/tests/fixtures/chart-deck.pptx",
        "6af2bc9f01bf08a8bcbe4dcafdfb04b8155f0b3fab5ebb14125b1712df4add72",
    ),
    (
        "crates/pptx-parse/tests/fixtures/custom-geometry.pptx",
        "7048388b8a2ae05227cc1a37fcc140b0cd18b5d53076f185ad0347edd763a6ef",
    ),
    (
        "crates/pptx-parse/tests/fixtures/gradient-outline.pptx",
        "1f710f5f8a373ce4b4e72ae6d92ee8e693bc8d74586701753cc559ae7270d14e",
    ),
    (
        "crates/pptx-parse/tests/fixtures/master-style-deck.pptx",
        "9665de345b8a9759d21eebad6fcba6575c504d91138e0863c126ed1f36bcb3b7",
    ),
    (
        "crates/pptx-parse/tests/fixtures/picture-fill.pptx",
        "df4c8fd780625c7a0c7ad8242b5a5c79bcb480b89306702770b2d8454bf71650",
    ),
    (
        "crates/pptx-parse/tests/fixtures/run-gradfill.pptx",
        "4744bf46b6d641d2090f4b11625abffa68c72f1da79b70a865e516c579b3ce30",
    ),
    (
        "crates/pptx-parse/tests/fixtures/shape-style.pptx",
        "756b96ac379b2da6093ab6a5232fe36136bfcf57f4fd40ea86864af445b2bfd8",
    ),
    (
        "crates/pptx-parse/tests/fixtures/slide-number-fields.pptx",
        "87d5724a4cdc209dbb98a7779ea5aa25e54892ba5bb58dcb3dc34a701ca6bda6",
    ),
    (
        "crates/pptx-parse/tests/fixtures/style-matrix-deck.pptx",
        "3f8385bdfd01f8c60f9f40c6a25e430adfda7b61999463981a869296b7c0592f",
    ),
    (
        "crates/pptx-render/tests/fixtures/alternate-content.pptx",
        "6fd8faa0eacde7ae362d2a1f052176f82644494f513d822787670f57b48a3520",
    ),
    (
        "crates/pptx-render/tests/fixtures/arrow-adjustments.pptx",
        "443d73c84617f457fd4a3626b5690d792401674200364b4173c2c3a49a1ec998",
    ),
    (
        "crates/pptx-render/tests/fixtures/autonumber-bullets.pptx",
        "e29bc88e443946e7ac61d0493f645129c13dc47671f14e40835f4ed98dd12dcc",
    ),
    (
        "crates/pptx-render/tests/fixtures/blip-effects.pptx",
        "7a5fb949c2f952a88a611dd47f7f7db9a0f61cfb7f80bfaf730311cadcc51771",
    ),
    (
        "crates/pptx-render/tests/fixtures/blip-lum.pptx",
        "162f64969cac1ce258d62b345b19f94cb437594b34029e476797b9fd0c040daf",
    ),
    (
        "crates/pptx-render/tests/fixtures/chart-legend.pptx",
        "84ce6caf9539d9d038a9eb650c969ed506215d4e829aac933d122810fa061328",
    ),
    (
        "crates/pptx-render/tests/fixtures/chart-series-line.pptx",
        "24d0a4d6de3172f18fce5f1eb9bc5f34303fbcbe6b88f4589e81444ee8215b2d",
    ),
    (
        "crates/pptx-render/tests/fixtures/chart-space-fill.pptx",
        "db7b3cc7d7380b5dba3a52fbf1e8d87326e1da3e5c487e62903b55ac8e67a969",
    ),
    (
        "crates/pptx-render/tests/fixtures/chart-text-properties.pptx",
        "db7b3cc7d7380b5dba3a52fbf1e8d87326e1da3e5c487e62903b55ac8e67a969",
    ),
    (
        "crates/pptx-render/tests/fixtures/data-labels.pptx",
        "900ced96636e99896c0754c5c2235cc8c63c12e9ec846bc8d03362a75621f374",
    ),
    (
        "crates/pptx-render/tests/fixtures/embedded-lato.pptx",
        "2c1fdbb661806031ac50dbc47f4d4a475fd04cd4f965a9fa3d332db88db96c02",
    ),
    (
        "crates/pptx-render/tests/fixtures/gradient-stop-order.pptx",
        "215b54c4308e15876a15109fc8f1f2dd9fe3190e2cc6e140af1d1d7d668649c7",
    ),
    (
        "crates/pptx-render/tests/fixtures/horizontal-bar-axes.pptx",
        "900ced96636e99896c0754c5c2235cc8c63c12e9ec846bc8d03362a75621f374",
    ),
    (
        "crates/pptx-render/tests/fixtures/line-ends.pptx",
        "4e7e06c1f2daa61d37f1a74a899723341415986d63ea5a9652abf9978be899ce",
    ),
    (
        "crates/pptx-render/tests/fixtures/line-spacing.pptx",
        "9505bbc87e33ebbcd08e4dbe0d5e2d6346fb937c760522861f1a6c594e900fd7",
    ),
    (
        "crates/pptx-render/tests/fixtures/list-style-bullets.pptx",
        "4a6fd3caa0d3beafedd35b850f2bb3bf8b5cc4aad14ea1a72504f6320b5dd7ed",
    ),
    (
        "crates/pptx-render/tests/fixtures/list-style-plain-paragraph.pptx",
        "bf4f9e38a4e2b425f3922a1aea30d7a3c3f43ec68fb71e87d1c4ce05542c3bdf",
    ),
    (
        "crates/pptx-render/tests/fixtures/metafile-gradient.pptx",
        "c6b0491a7df9bd854c9f9de39486c8eeba2bad68a2da4ca40680de5089ca88e8",
    ),
    (
        "crates/pptx-render/tests/fixtures/metafile-pictures.pptx",
        "3b0fb89b1ad7f287249e6546d1786f0924d43574fb73ac92072aabfb5d477467",
    ),
    (
        "crates/pptx-render/tests/fixtures/outer-shadow-picture.pptx",
        "2516c37b3b80f79933e0e6f08cc9a62a0630d48e470ff69fa6827bb249d78a34",
    ),
    (
        "crates/pptx-render/tests/fixtures/outer-shadow-scale.pptx",
        "e42069927b96fd50f0afff5aa5c4190905aa3eb9ed38e6d7a11c29122504b3ab",
    ),
    (
        "crates/pptx-render/tests/fixtures/outer-shadow.pptx",
        "b9dd2d43df946e004e464813ebd5cf89b4a84f5064db1621f786b5e648b65442",
    ),
    (
        "crates/pptx-render/tests/fixtures/paragraph-right-margin.pptx",
        "efdc2b53416b0448926cd01e25bd9b9ae9ef0084cfcdb94c33ff0869d43797c9",
    ),
    (
        "crates/pptx-render/tests/fixtures/paragraph-spacing.pptx",
        "d2e414f18c3b9ed74cec7ffca872a25a42b12110125f47fe0eae3cee007c6a6a",
    ),
    (
        "crates/pptx-render/tests/fixtures/picture-crop-mask.pptx",
        "864b5e4c58fdff51c68e0ae7eddbf14ade892dface7575e134782414f584522c",
    ),
    (
        "crates/pptx-render/tests/fixtures/prstgeom.pptx",
        "4872641e3ec6531676fa4605d251b2016c22d0e60c886ccaa004e9d935efbc59",
    ),
    (
        "crates/pptx-render/tests/fixtures/run-caps.pptx",
        "8af079a847b4ec28d0127be6fcf27b7c34c614e492d3002c58ab6a3581edebb1",
    ),
    (
        "crates/pptx-render/tests/fixtures/run-paint-scope.pptx",
        "25ce41068eae8a51dd24e73dd8b3095961bfd71df03222bdf7b3cb0a28a627f5",
    ),
    (
        "crates/pptx-render/tests/fixtures/run-spacing.pptx",
        "19bd623a80670a541d635436852594a99351221668eef37c711a4381f87a73f1",
    ),
    (
        "crates/pptx-render/tests/fixtures/star-inner-radius.pptx",
        "23c182911b13a4337fed79b704a3934e11d1dd634519fb4497a248fa82c73dab",
    ),
    (
        "crates/pptx-render/tests/fixtures/symbol-bullets.pptx",
        "cfd0a1143e981972ea0edb921a22ec550e41aeb3d682239f91dee088d8516cce",
    ),
    (
        "crates/pptx-render/tests/fixtures/table-basic.pptx",
        "92cb00f7eacb375f79b953f06a4263f00e19fce147917265e4389ec4c3bb366f",
    ),
    (
        "crates/pptx-render/tests/fixtures/table-rotated.pptx",
        "5735b3e85a9dc17ecd686e6b9d9fce71c78d8c0d3034de7c830bfb5109172109",
    ),
    (
        "crates/pptx-render/tests/fixtures/text-baseline-script.pptx",
        "cb6573480d5d8b638d217a2c408d0786ce03f370893dd62fb35c18d44bcb4085",
    ),
    (
        "crates/pptx-render/tests/fixtures/text-orientation.pptx",
        "dc25bf4cf0dc70fb6b6b8f0389df56b2dac80e7a17456b64cb64293912f92f12",
    ),
    (
        "crates/pptx-render/tests/fixtures/text-overflow.pptx",
        "a9660d74a886f49b4c7ba6268b62735814f64ec9cee4500f03374ad639703c89",
    ),
    (
        "crates/pptx-render/tests/fixtures/theme-fonts.pptx",
        "34e4828358a721a3264399ea852b29c5c2846b3d0b956655eaac929b370c205c",
    ),
    (
        "crates/pptx-render/tests/fixtures/value-axis-autoscale.pptx",
        "900ced96636e99896c0754c5c2235cc8c63c12e9ec846bc8d03362a75621f374",
    ),
    (
        "crates/pptx-render/tests/fixtures/vertical-writing-modes.pptx",
        "2235fce2dc91d85d59306248d34dda74f3a9c75d0da786f205a54c8dd3f5de34",
    ),
    (
        "fuzz/corpus/pptx-package-parse/custom-geometry.pptx",
        "7048388b8a2ae05227cc1a37fcc140b0cd18b5d53076f185ad0347edd763a6ef",
    ),
    (
        "fuzz/corpus/pptx-package-parse/line-ends.pptx",
        "4e7e06c1f2daa61d37f1a74a899723341415986d63ea5a9652abf9978be899ce",
    ),
    (
        "fuzz/corpus/pptx-package-parse/modern-comments.pptx",
        "70039f4165965cddaf3165365ea0d2ab1d0b4e1ed2a2cbaaecb926ea16735fbf",
    ),
    (
        "fuzz/corpus/pptx-package-parse/slide-number-fields.pptx",
        "87d5724a4cdc209dbb98a7779ea5aa25e54892ba5bb58dcb3dc34a701ca6bda6",
    ),
    (
        "packages/pptx/src/render/fixtures/tiff-image.pptx",
        "df4c8fd780625c7a0c7ad8242b5a5c79bcb480b89306702770b2d8454bf71650",
    ),
    (
        "poc/fixtures/deck-50.pptx",
        "92dec8d0a24cae1d4ca3339d46c6ceefaa49abd533a460ba09cbc2c1204963f7",
    ),
    (
        "poc/fixtures/feature-rich.pptx",
        "46cb5337a8a65e6231229583db55ec6ba5d82a387a528ad050e5d77858afe496",
    ),
    (
        "poc/fixtures/lecture.pptx",
        "13a6155e5c4ce447af6f3954381702f819ebbcb99af6ba97560f7f1c662694fa",
    ),
];

/// The fixture's saved slide parts, names and bytes, hashed in name order
/// after the editor's two shapes went on its first slide.
fn digest(path: &str) -> String {
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../..");
    let bytes = std::fs::read(root.join(path)).unwrap();
    let Ok(session) = DeckSession::open(&bytes, 7) else {
        return "-".to_owned();
    };
    let Some(slide) = session
        .snapshot()
        .ok()
        .and_then(|deck| deck.slides.first().map(|slide| slide.id.clone()))
    else {
        return "-".to_owned();
    };
    let context = EditCtx::local("test");
    let rect = |x: i64| ShapeRect {
        x,
        y: 200_000,
        width: 1_000_000,
        height: 500_000,
    };
    session
        .add_text_box(
            &context,
            &slide,
            &ShapeDraft {
                name: "Note".to_owned(),
                rect: rect(100_000),
                text: "Saved note".to_owned(),
                style: TextStyle::default(),
            },
        )
        .unwrap();
    session
        .add_shape(
            &context,
            &slide,
            &PresetShapeDraft {
                name: "Badge".to_owned(),
                geometry: "roundRect".to_owned(),
                rect: rect(1_200_000),
                fill: Some("#00AAFF".to_owned()),
            },
        )
        .unwrap();
    let mut parts = ooxml_opc::unzip_parts(&session.save().unwrap()).unwrap();
    parts.sort();
    let mut hasher = Sha256::new();
    for (name, bytes) in parts {
        if name.starts_with("ppt/slides/slide") && name.ends_with(".xml") {
            hasher.update(name.as_bytes());
            hasher.update(&bytes);
        }
    }
    hasher
        .finalize()
        .iter()
        .map(|byte| format!("{byte:02x}"))
        .collect()
}

#[test]
fn every_fixture_saves_an_added_shape_as_before() {
    let mut found = Vec::new();
    for (path, _) in FIXTURES {
        let line = format!("{path} {}", digest(path));
        eprintln!("fixture {line}");
        found.push(line);
    }
    let pinned: Vec<String> = FIXTURES
        .iter()
        .map(|(path, digest)| format!("{path} {digest}"))
        .collect();
    assert_eq!(found, pinned);
}
