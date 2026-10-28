//! Every PPTX fixture in the repository, saved with a text box and a preset
//! shape the editor adds to its first slide, writes its slides byte for
//! byte as the save did before the chat agent's layout placeholders (the
//! Editor's condition, 2026-10-22): the pins were computed on fork commit
//! 7c43320ab, before `ShapeAdd.placeholder`, and re-pinned in the 0.5.1 sync
//! for the added text box's `txBox="1"` (upstream #1507), the only change:
//! without it every fixture hashed to its old pin. "-" marks a fixture the
//! editor does not open or that has no slide.

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
        "6026427fe29ffe82339f5fe2d4a47922650d816184034d218bce7ab29adfc86d",
    ),
    (
        "crates/ooxml-opc/tests/fixtures/betteroffice-demo.pptx",
        "327c41068bcd499cbc864f7c754e43d73029c10ee1263f3a2ea4a776028db543",
    ),
    (
        "crates/pptx-edit/tests/fixtures/blip-shadow.pptx",
        "463086b31e9a303af5dc6a5b906d565ffea09b9b0ae635b04540b6903fb1cc27",
    ),
    (
        "crates/pptx-edit/tests/fixtures/chart-text-overflow.pptx",
        "5e9abf5400b20348f7707b6e721fd69933d180f1a593af039590b9e6ef730d47",
    ),
    (
        "crates/pptx-edit/tests/fixtures/deck-schema-v2-connectors.pptx",
        "9c86a89c4622def86e7e8e831bc0c052002f57b1ed9e132a5957740a54134fd2",
    ),
    (
        "crates/pptx-edit/tests/fixtures/deck-schema-v2-nested-connectors.pptx",
        "112ca9a89435e8a5156e51fc28e3b83cc1ef60b4b3e6be0815fb023588c92746",
    ),
    (
        "crates/pptx-edit/tests/fixtures/hidden-shapes.pptx",
        "66fe0fd37c8f5b2d1b272402b29773a3c9827b5177b15efe9d6f4bed22b6cfe4",
    ),
    (
        "crates/pptx-edit/tests/fixtures/metafile-tracking.pptx",
        "6b7f571e23494b7881a6d87933debf737b5b1adf7b090a6881b9b4bf2cab3a2e",
    ),
    (
        "crates/pptx-edit/tests/fixtures/modern-comments.pptx",
        "327c41068bcd499cbc864f7c754e43d73029c10ee1263f3a2ea4a776028db543",
    ),
    (
        "crates/pptx-edit/tests/fixtures/run-spacing-shadow.pptx",
        "8462b9178fd4d97d7d237850086ab3ad52cc30ae00d3fe33b74e9c6e4f05512b",
    ),
    (
        "crates/pptx-parse/tests/fixtures/chart-deck.pptx",
        "279e82e3620ad3b79d8788279d085a2bae64db067791a6f9826f3ea938c038dd",
    ),
    (
        "crates/pptx-parse/tests/fixtures/custom-geometry.pptx",
        "01f368686fe1d137a82474bde8f2ed77a1838fbd8aa8cd32cc9c99d1a13f24f1",
    ),
    (
        "crates/pptx-parse/tests/fixtures/gradient-outline.pptx",
        "1d585e970ca8c7c0b3e72c5c9aa7f9001dab60703e727a2b2c9383560f2ab668",
    ),
    (
        "crates/pptx-parse/tests/fixtures/master-style-deck.pptx",
        "6517eeb50963c370dcc0c2c8f4af4bf94d404a25af26aa5d3eb897b3232ff4f8",
    ),
    (
        "crates/pptx-parse/tests/fixtures/picture-fill.pptx",
        "e4eba934b578f47b91c849a95ddacb2b12de7b784e84e6fb3c70a98e906bc750",
    ),
    (
        "crates/pptx-parse/tests/fixtures/run-gradfill.pptx",
        "424e2c2e36a811e5ddde2ed95214bf36061d21e9c4888246940e17796b440357",
    ),
    (
        "crates/pptx-parse/tests/fixtures/shape-style.pptx",
        "7d4f3e98d420641968c7fda41fc9aac450562c48c0e64dfdcde7ba982c267791",
    ),
    (
        "crates/pptx-parse/tests/fixtures/slide-number-fields.pptx",
        "bb150acd246646e93c3a8f8f9ef25ce6a3623d4194f43e818f8544639bf0c31d",
    ),
    (
        "crates/pptx-parse/tests/fixtures/style-matrix-deck.pptx",
        "d74182e134a60e780f9acc6d25fd8f157da497bd3e4aeba8ca33d6ba4e8d7a84",
    ),
    (
        "crates/pptx-render/tests/fixtures/alternate-content.pptx",
        "69584e03a8f9a99bf89c894f5d56f913c39ebe77db9c67a05ac453f24f586e09",
    ),
    (
        "crates/pptx-render/tests/fixtures/arrow-adjustments.pptx",
        "3e6dc3451f205d4f61b9f8f03b9abaac31f6d781cfd71d9dd40c4cfcfacf7303",
    ),
    (
        "crates/pptx-render/tests/fixtures/autonumber-bullets.pptx",
        "2a7ebf75bdc164facaed24d14bb2483afa9b10933e439ef88c9975eada4b29c0",
    ),
    (
        "crates/pptx-render/tests/fixtures/blip-effects.pptx",
        "db6ac873b31c0457cf93c52e9df50598274b510d563ce56517bf10162e3a57d1",
    ),
    (
        "crates/pptx-render/tests/fixtures/blip-lum.pptx",
        "f2cba930c2a69cefb1b5bc3f231e5b816840b3d111ceb7dc4d2f1b9d90df07c4",
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
        "985bedcfa04a4784763229196c46939ba33153f9a1d68175fe73f9df6b3cc674",
    ),
    (
        "crates/pptx-render/tests/fixtures/gradient-stop-order.pptx",
        "3011bb000657dca9affb4567df8376cf49749ea13905325edc8688abf342f7b6",
    ),
    (
        "crates/pptx-render/tests/fixtures/horizontal-bar-axes.pptx",
        "af83e374ec97707aa069fa82c68017f6f56037321585417ea4f3593aeb0ac6c7",
    ),
    (
        "crates/pptx-render/tests/fixtures/line-ends.pptx",
        "60c65bb3b3e729dc4ea0b0a1ff38916d2afd0baa462aa1a730b4271ec63d926a",
    ),
    (
        "crates/pptx-render/tests/fixtures/line-spacing.pptx",
        "86f9294b91f9f85d96c7d6c4f9aca663c807b358a3885e3a516e3ff6c75f58d4",
    ),
    (
        "crates/pptx-render/tests/fixtures/list-style-bullets.pptx",
        "aeb068e31cdc102476b85c48dadafc43e8a26f378c79ee18df07f47f562f0f5c",
    ),
    (
        "crates/pptx-render/tests/fixtures/list-style-plain-paragraph.pptx",
        "a113de17312839421d0ac9e831d57877bfd8b0352612f943b6159d6b847c20b7",
    ),
    (
        "crates/pptx-render/tests/fixtures/metafile-gradient.pptx",
        "4d373a44a2588d741c7125c1235f838bb09b40d4a5c2c97a2156c14bd83ff61d",
    ),
    (
        "crates/pptx-render/tests/fixtures/metafile-pictures.pptx",
        "36cab83f6df2b9224ff3cdfd1cf69384eb2a5c736d11049d6a4272064825aa29",
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
        "dd547ab94bae8a22dcd0ac230e099cb7f7cf94c8e5b66be62c81dccba0413c9a",
    ),
    (
        "crates/pptx-render/tests/fixtures/paragraph-spacing.pptx",
        "7a891a5cb2000a0441c186bd0919dff4b79886f857987393586547cf707f74a8",
    ),
    (
        "crates/pptx-render/tests/fixtures/picture-crop-mask.pptx",
        "83598e357559fb79eccbf36910b5cd211f4f233c5e7cae9aa69ec3ca36743de5",
    ),
    (
        "crates/pptx-render/tests/fixtures/prstgeom.pptx",
        "4d8a15d44164353842b8035d86b978546e7ccc02ae747827304181ed0761ca16",
    ),
    (
        "crates/pptx-render/tests/fixtures/run-caps.pptx",
        "f50f690cb5380092e2cb420893211e4d1163e5df8eb705bc22da91f466ec92e2",
    ),
    (
        "crates/pptx-render/tests/fixtures/run-paint-scope.pptx",
        "3de9e635717e0144d7183f8f74b1cdcf9e2f3cd5dd94516159b377546f778096",
    ),
    (
        "crates/pptx-render/tests/fixtures/run-spacing.pptx",
        "0aa854d8ccd586e9dc91a17701f57637b94401e5c7d19361eaf362f7ceaf3ace",
    ),
    (
        "crates/pptx-render/tests/fixtures/star-inner-radius.pptx",
        "795f701eb092def720b4d71c46f5a5dd02f8fed05b1a23aefb8f3db85ef5d006",
    ),
    (
        "crates/pptx-render/tests/fixtures/symbol-bullets.pptx",
        "fc33f8c81123148cbee485ce448f3b0a760e600c2cc8e13b773accfddb42744f",
    ),
    (
        "crates/pptx-render/tests/fixtures/table-basic.pptx",
        "0754189a707fa406327266c393f0bab568fcb4e81d57bb7cd12307b9b35c49bd",
    ),
    (
        "crates/pptx-render/tests/fixtures/table-rotated.pptx",
        "0018342620e3f6ca7154653bdbdd99a44e523ad1229b73b2f2ad2a9f83d02e89",
    ),
    (
        "crates/pptx-render/tests/fixtures/text-baseline-script.pptx",
        "21884be9a82edffa3d90f3109eb17a89f136bf6bad5462cad6568869f664f0a4",
    ),
    (
        "crates/pptx-render/tests/fixtures/text-orientation.pptx",
        "2cc87281fac533c12a73a70f15477b8ac1b6fe36a0b5dd16cc7c884b03d89e48",
    ),
    (
        "crates/pptx-render/tests/fixtures/text-overflow.pptx",
        "04d86c00d7aa1acec0660e1a0d3e2d346cee33c38577de8803554c5fe33f38e8",
    ),
    (
        "crates/pptx-render/tests/fixtures/theme-fonts.pptx",
        "c261977c86205fa8a4c17b8685dbcc490c95dbed765df00aa1f0d74589be0ee2",
    ),
    (
        "crates/pptx-render/tests/fixtures/value-axis-autoscale.pptx",
        "af83e374ec97707aa069fa82c68017f6f56037321585417ea4f3593aeb0ac6c7",
    ),
    (
        "crates/pptx-render/tests/fixtures/vertical-writing-modes.pptx",
        "977181aad0462ed34f621350ac53e78ac2a2e68dfd563dd1d693aa7636b5a5de",
    ),
    (
        "fuzz/corpus/pptx-package-parse/custom-geometry.pptx",
        "01f368686fe1d137a82474bde8f2ed77a1838fbd8aa8cd32cc9c99d1a13f24f1",
    ),
    (
        "fuzz/corpus/pptx-package-parse/line-ends.pptx",
        "60c65bb3b3e729dc4ea0b0a1ff38916d2afd0baa462aa1a730b4271ec63d926a",
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
        "e4eba934b578f47b91c849a95ddacb2b12de7b784e84e6fb3c70a98e906bc750",
    ),
    (
        "poc/fixtures/deck-50.pptx",
        "2ca9ff7a67448258f46355054385906f399a42085528458c4daa07115c072f29",
    ),
    (
        "poc/fixtures/feature-rich.pptx",
        "ac6a38708e29fc891c8647aea128a2c42eba4de8b1f8764bd9d3361876eebf90",
    ),
    (
        "poc/fixtures/lecture.pptx",
        "1b6f45a1074a11faaf7c0d067635d5405aeba2d5eb14c118663692785a81ce79",
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
