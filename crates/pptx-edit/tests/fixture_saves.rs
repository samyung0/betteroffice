//! Every PPTX fixture in the repository, saved with a text box and a preset
//! shape the editor adds to its first slide, writes its slides byte for
//! byte as the save did before the chat agent's layout placeholders (the
//! Editor's condition, 2026-10-22): the pins were computed on fork commit
//! 7c43320ab, before `ShapeAdd.placeholder`, and re-pinned in the 0.5.1 sync
//! for the added text box's `txBox="1"` (upstream #1507), the only change:
//! without it every fixture hashed to its old pin. Re-pinned in the second
//! sync for upstream #1860's splice save, which keeps the source's markup
//! around the added shapes: 45 fixtures changed, each saved slide equal to the
//! previous save under XML canonicalization, and without the splice every
//! fixture hashed to its old pin. "-" marks a fixture the editor does not open
//! or that has no slide.

use std::path::PathBuf;

use pptx_edit::{DeckSession, EditCtx, PresetShapeDraft, ShapeDraft, ShapeRect, TextStyle};
use sha2::{Digest, Sha256};

const FIXTURES: &[(&str, &str)] = &[
    (
        "apps/demo/public/betteroffice-demo.pptx",
        "327c41068bcd499cbc864f7c754e43d73029c10ee1263f3a2ea4a776028db543",
    ),
    (
        "crates/ooxml-drawingml/tests/fixtures/preset-adjustments.pptx",
        "000652df7c76a3bbd599d6de98dab7173dfc6374be950754294143aeb1a35ce7",
    ),
    (
        "crates/ooxml-opc/tests/fixtures/betteroffice-demo.pptx",
        "327c41068bcd499cbc864f7c754e43d73029c10ee1263f3a2ea4a776028db543",
    ),
    (
        "crates/pptx-edit/tests/fixtures/blip-shadow.pptx",
        "2db25a2fa978ff8aa4445e2ba9f56e9e0e37d449499a002935a72e1368cb381a",
    ),
    (
        "crates/pptx-edit/tests/fixtures/chart-text-overflow.pptx",
        "37bb9cd0640ff8c15e6206d3f0c6fe60b9e7b6bc3311c8ec8c2a5334d092c60e",
    ),
    (
        "crates/pptx-edit/tests/fixtures/deck-schema-v2-connectors.pptx",
        "9c86a89c4622def86e7e8e831bc0c052002f57b1ed9e132a5957740a54134fd2",
    ),
    (
        "crates/pptx-edit/tests/fixtures/deck-schema-v2-nested-connectors.pptx",
        "5c1921cf466dffd1d943a4fd79cb818d3f723cd4e5cac89998a304614e5d3ee1",
    ),
    (
        "crates/pptx-edit/tests/fixtures/hidden-shapes.pptx",
        "66fe0fd37c8f5b2d1b272402b29773a3c9827b5177b15efe9d6f4bed22b6cfe4",
    ),
    (
        "crates/pptx-edit/tests/fixtures/metafile-tracking.pptx",
        "aec7187272b94663665a5f3debc902a1c2869732fe0476e768cd1d91afb8f83e",
    ),
    (
        "crates/pptx-edit/tests/fixtures/modern-comments.pptx",
        "327c41068bcd499cbc864f7c754e43d73029c10ee1263f3a2ea4a776028db543",
    ),
    (
        "crates/pptx-edit/tests/fixtures/run-spacing-shadow.pptx",
        "5ca7775e0c10ceb429c7062c58da7edbce644df57eb8f0147984da9fe4f8ce78",
    ),
    (
        "crates/pptx-parse/tests/fixtures/chart-deck.pptx",
        "279e82e3620ad3b79d8788279d085a2bae64db067791a6f9826f3ea938c038dd",
    ),
    (
        "crates/pptx-parse/tests/fixtures/custom-geometry.pptx",
        "f08640c5e18f2534f20b51b3fb1ee4372834c1b338a49185c8972fc8026e1699",
    ),
    (
        "crates/pptx-parse/tests/fixtures/gradient-outline.pptx",
        "1d585e970ca8c7c0b3e72c5c9aa7f9001dab60703e727a2b2c9383560f2ab668",
    ),
    (
        "crates/pptx-parse/tests/fixtures/master-style-deck.pptx",
        "7ec48ee9c2df92f8dd2d6c40f049238580b30221fd84834d0c61ac087ca63158",
    ),
    (
        "crates/pptx-parse/tests/fixtures/picture-fill.pptx",
        "1c7bd4a7ef5e562d19c5ef7ea00a82ae49cf8191f6894fe8c2d4bdc432934f1d",
    ),
    (
        "crates/pptx-parse/tests/fixtures/run-gradfill.pptx",
        "21aad12bcec381595fe56c3aa1c9c860f8a8b50c44c1ac66eb1f8a26e6ec2882",
    ),
    (
        "crates/pptx-parse/tests/fixtures/shape-style.pptx",
        "4e19a4308f83e3110600bc26c4c97eba9c07819bc9dacafaf64d48b8136183ff",
    ),
    (
        "crates/pptx-parse/tests/fixtures/slide-number-fields.pptx",
        "bb150acd246646e93c3a8f8f9ef25ce6a3623d4194f43e818f8544639bf0c31d",
    ),
    (
        "crates/pptx-parse/tests/fixtures/style-matrix-deck.pptx",
        "9fa2bad157743cc8595551e1a19dfa8e90bc8e1f8fd98cc6592f49530cf9c068",
    ),
    (
        "crates/pptx-render/tests/fixtures/alternate-content.pptx",
        "cdaf6870f93133df2fda6d30725eb726a678587e9ee12edbe05434ca300c779a",
    ),
    (
        "crates/pptx-render/tests/fixtures/arrow-adjustments.pptx",
        "4137e02a1b32fb544a5805647c3f7f919741cca7b520abde92deeb2d69f98f4c",
    ),
    (
        "crates/pptx-render/tests/fixtures/autonumber-bullets.pptx",
        "e8e2d88e821167e9d1e7346ab6a8dd3abb2645c40e56831cfcd6a5c308fe05d2",
    ),
    (
        "crates/pptx-render/tests/fixtures/blip-effects.pptx",
        "9f3e056c1f9473def08356a434bc97894a9cb7bdb5fcd7a0d3685c39c84d4f05",
    ),
    (
        "crates/pptx-render/tests/fixtures/blip-lum.pptx",
        "2de3f0e4e8a4a75b6d5f966d9ef5aef2aee86a5af0581d81c33617e4a56a77da",
    ),
    (
        "crates/pptx-render/tests/fixtures/chart-legend.pptx",
        "4b4cf570fbb5b812dfed343540d43eaabf9f0ba1b8db9befb8960a6b1bea1d1d",
    ),
    (
        "crates/pptx-render/tests/fixtures/chart-series-line.pptx",
        "df6d74ae0698b062d8a65cdefaecce83d0e6e3ff04e92f0ca2b922ef9deb874c",
    ),
    (
        "crates/pptx-render/tests/fixtures/chart-space-fill.pptx",
        "ca813454dd000b273c5ddec64e6785e98ce5e077ca2f53fb24bc771edec07f39",
    ),
    (
        "crates/pptx-render/tests/fixtures/chart-text-properties.pptx",
        "ca813454dd000b273c5ddec64e6785e98ce5e077ca2f53fb24bc771edec07f39",
    ),
    (
        "crates/pptx-render/tests/fixtures/data-labels.pptx",
        "af83e374ec97707aa069fa82c68017f6f56037321585417ea4f3593aeb0ac6c7",
    ),
    (
        "crates/pptx-render/tests/fixtures/embedded-lato.pptx",
        "3769c908884e331e56fe01c269c1b81bca6c8494bff9fa7260d871bb82f7943d",
    ),
    (
        "crates/pptx-render/tests/fixtures/gradient-stop-order.pptx",
        "2a730ae7311c9ec5304f32a2c7ec5ab3318cde6b6a1b7bba6d8154e294a6e1e9",
    ),
    (
        "crates/pptx-render/tests/fixtures/horizontal-bar-axes.pptx",
        "af83e374ec97707aa069fa82c68017f6f56037321585417ea4f3593aeb0ac6c7",
    ),
    (
        "crates/pptx-render/tests/fixtures/line-ends.pptx",
        "972d72bf892913a60011a9ec3f5a44607e6616ce0079ee30b38ef927f0d4522c",
    ),
    (
        "crates/pptx-render/tests/fixtures/line-spacing.pptx",
        "4c43a621ccdab1779fd1a6d29b620f41feb817be161f6929c38d15e54098b7b2",
    ),
    (
        "crates/pptx-render/tests/fixtures/list-style-bullets.pptx",
        "104a206344b40d1ef7373a545a893717fa6478eef9f9a2fa19b07d4a772dee3b",
    ),
    (
        "crates/pptx-render/tests/fixtures/list-style-plain-paragraph.pptx",
        "bd4a798dab2d259e06dd00d168fbc61aa36fc5bc23936809303c0bde967af7ab",
    ),
    (
        "crates/pptx-render/tests/fixtures/metafile-gradient.pptx",
        "923fae127095baf2fe5a24c694e7369e2e1ac628957d9c7e653fc615cb30ba63",
    ),
    (
        "crates/pptx-render/tests/fixtures/metafile-pictures.pptx",
        "b0b83e17b9b5826789a853a6a02b395848a3132c765fe992947307a6c7791d3b",
    ),
    (
        "crates/pptx-render/tests/fixtures/outer-shadow-picture.pptx",
        "cb452fc18730517b20d908f5722f8f04b38ea1bd698957d93aa54e49a14c2568",
    ),
    (
        "crates/pptx-render/tests/fixtures/outer-shadow-scale.pptx",
        "9b4875db287be9d90b61d42cbb719cab4ab2ac2c451541fd42fb2177360a5f41",
    ),
    (
        "crates/pptx-render/tests/fixtures/outer-shadow.pptx",
        "ede2ae34d2a31deef91e71c6b20d51ee6db453f3f44d71e23f0aeff9cfb0db97",
    ),
    (
        "crates/pptx-render/tests/fixtures/paragraph-right-margin.pptx",
        "982f90afc2ebf15a57e1fbe7024276689a7e4c576e48341da18f31b013754d18",
    ),
    (
        "crates/pptx-render/tests/fixtures/paragraph-spacing.pptx",
        "9322f2f0d7876aeef35f0033f0da0ea62990fe24d6db602eb9502c637f02e339",
    ),
    (
        "crates/pptx-render/tests/fixtures/picture-crop-mask.pptx",
        "83598e357559fb79eccbf36910b5cd211f4f233c5e7cae9aa69ec3ca36743de5",
    ),
    (
        "crates/pptx-render/tests/fixtures/prstgeom.pptx",
        "08aeea30cc1d2916c9a0b24526314270f86c233d59da8807882499c08da91cf9",
    ),
    (
        "crates/pptx-render/tests/fixtures/run-caps.pptx",
        "3aef31273d7bafe90d931e5126168fcb84dc4f5f84083c825f3c347d8d02db67",
    ),
    (
        "crates/pptx-render/tests/fixtures/run-paint-scope.pptx",
        "d44374fca77eaba4f45330ff528616f5753242a006a3529d52796a6575e77428",
    ),
    (
        "crates/pptx-render/tests/fixtures/run-spacing.pptx",
        "2201053fd434bf5b5d3a5bfe233391d354c994275eb149f640a077bd746dad38",
    ),
    (
        "crates/pptx-render/tests/fixtures/star-inner-radius.pptx",
        "d5cf141664a2221a9eeb40c10739d3f6cc9cf24a19a7047abb767da17fa99dea",
    ),
    (
        "crates/pptx-render/tests/fixtures/symbol-bullets.pptx",
        "451d37c70f69b7c3b627be78949b373417a300a988f01bcb5da638bdf29e377a",
    ),
    (
        "crates/pptx-render/tests/fixtures/table-basic.pptx",
        "8e4610a688b046a7125a7994af35b5ace562fa38a42f236c862a02b5615ef7e0",
    ),
    (
        "crates/pptx-render/tests/fixtures/table-rotated.pptx",
        "df07768b826e5ac1acd257967f72527887cf9f845ea3edf6ccc41192ae8785c8",
    ),
    (
        "crates/pptx-render/tests/fixtures/text-baseline-script.pptx",
        "21884be9a82edffa3d90f3109eb17a89f136bf6bad5462cad6568869f664f0a4",
    ),
    (
        "crates/pptx-render/tests/fixtures/text-orientation.pptx",
        "1c26e30798f4b38a1dcbf2b6e98d53029ff42a488114b45ae407ef18674af86e",
    ),
    (
        "crates/pptx-render/tests/fixtures/text-overflow.pptx",
        "9f9942b752c5e4ef8b77697f266b631e44d23774e02d29906051b16548de03c6",
    ),
    (
        "crates/pptx-render/tests/fixtures/theme-fonts.pptx",
        "9470847c1f88e2f6bad9f2b7e6e9d71a877f4c893f67beb81354227611887e37",
    ),
    (
        "crates/pptx-render/tests/fixtures/value-axis-autoscale.pptx",
        "af83e374ec97707aa069fa82c68017f6f56037321585417ea4f3593aeb0ac6c7",
    ),
    (
        "crates/pptx-render/tests/fixtures/vertical-writing-modes.pptx",
        "7d5b57a77ea5ec99180ddc5372ebb616e255fb108006fe840379391aaede6786",
    ),
    (
        "fuzz/corpus/pptx-package-parse/custom-geometry.pptx",
        "f08640c5e18f2534f20b51b3fb1ee4372834c1b338a49185c8972fc8026e1699",
    ),
    (
        "fuzz/corpus/pptx-package-parse/line-ends.pptx",
        "972d72bf892913a60011a9ec3f5a44607e6616ce0079ee30b38ef927f0d4522c",
    ),
    (
        "fuzz/corpus/pptx-package-parse/modern-comments.pptx",
        "327c41068bcd499cbc864f7c754e43d73029c10ee1263f3a2ea4a776028db543",
    ),
    (
        "fuzz/corpus/pptx-package-parse/slide-number-fields.pptx",
        "bb150acd246646e93c3a8f8f9ef25ce6a3623d4194f43e818f8544639bf0c31d",
    ),
    (
        "packages/pptx/src/render/fixtures/tiff-image.pptx",
        "1c7bd4a7ef5e562d19c5ef7ea00a82ae49cf8191f6894fe8c2d4bdc432934f1d",
    ),
    (
        "poc/fixtures/deck-50.pptx",
        "c681bdf50a377810527831c555cd5c084e8655d0e5348b4e9f6b34dca92099ed",
    ),
    (
        "poc/fixtures/feature-rich.pptx",
        "49b44d0392e4c2ca010a17a5b8f637c75c148deade3a6200bc6e3b08a3cda8ef",
    ),
    (
        "poc/fixtures/lecture.pptx",
        "aaa4b0c894065e77b8b768c178008b07f730d271bd9eb66e1be2be657c913f3b",
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
