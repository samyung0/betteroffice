use pptx_edit::{DeckSession, EditCtx, TextStyle, TextStylePatch};
use yrs::types::Attrs;
use yrs::updates::decoder::Decode;
use yrs::{Any, Doc, Map, ReadTxn, Text, TextRef, Transact, Update};

const DECK: &[u8] = include_bytes!("../../pptx-render/tests/fixtures/run-spacing.pptx");
/// `DECK` with `kern="0"` on its first run and in its second shape's list style.
const KERNED: &[u8] = include_bytes!("fixtures/run-kern.pptx");
const STORY: &str = "story:slide:0:256:shape:0:0";

fn run_kern(session: &DeckSession) -> Option<f64> {
    session.story(STORY).unwrap().paragraphs[0].runs[0]
        .style
        .kern_pt
}

fn story(txn: &yrs::TransactionMut<'_>) -> TextRef {
    txn.get_map("pptx:stories")
        .unwrap()
        .get(txn, STORY)
        .unwrap()
        .cast::<TextRef>()
        .unwrap()
}

fn slide_xml(bytes: &[u8]) -> String {
    let parts = ooxml_opc::unzip_parts(bytes).unwrap();
    let (_, xml) = parts
        .iter()
        .find(|(path, _)| path == "ppt/slides/slide1.xml")
        .unwrap();
    String::from_utf8(xml.clone()).unwrap()
}

#[test]
fn a_runs_own_kern_threshold_reaches_the_snapshot_and_survives_an_edit() {
    let session = DeckSession::open(KERNED, 32508).unwrap();
    assert_eq!(run_kern(&session), Some(0.0));

    let context = EditCtx::local("test");
    let end = session.story(STORY).unwrap().length - 1;
    session
        .format_text(
            &context,
            STORY,
            0,
            end,
            &TextStylePatch {
                bold: Some(true),
                ..Default::default()
            },
        )
        .unwrap();
    let saved = session.save().unwrap();
    assert!(slide_xml(&saved).contains(r#"kern="0""#));
    let reopened = DeckSession::open(&saved, 32509).unwrap();
    assert_eq!(run_kern(&reopened), Some(0.0));
}

#[test]
fn an_inserted_runs_kern_threshold_must_be_a_size() {
    let session = DeckSession::open(DECK, 32510).unwrap();
    let context = EditCtx::local("test");
    let before = session.encode_state_as_update_v1();
    for kern in [f64::NAN, f64::INFINITY, -1.0, 4000.01] {
        assert!(
            session
                .insert_text(
                    &context,
                    STORY,
                    0,
                    "X",
                    &TextStyle {
                        kern_pt: Some(kern),
                        ..Default::default()
                    }
                )
                .is_err(),
            "{kern}"
        );
    }
    assert_eq!(session.encode_state_as_update_v1(), before);
}

#[test]
fn a_remote_kern_outside_the_schema_range_is_rejected() {
    let session = DeckSession::open(DECK, 32511).unwrap();
    let before = session.encode_state_as_update_v1();
    let remote = |client: u64, kern: Any| {
        let doc = Doc::with_client_id(client);
        doc.transact_mut()
            .apply_update(Update::decode_v1(&before).unwrap())
            .unwrap();
        let mut txn = doc.transact_mut();
        story(&txn).format(&mut txn, 0, 3, Attrs::from([("kern".into(), kern)]));
        txn.encode_update_v1()
    };
    for (client, kern) in [
        Any::Number(-1.0),
        Any::Number(4_000.01),
        Any::String("12".into()),
    ]
    .into_iter()
    .enumerate()
    {
        let update = remote(32520 + client as u64, kern.clone());
        assert!(session.apply_update_v1(&update).is_err(), "{kern:?}");
        assert_eq!(session.encode_state_as_update_v1(), before);
    }
    session
        .apply_update_v1(&remote(32530, Any::Number(f64::NAN)))
        .unwrap();
    assert_eq!(run_kern(&session), None, "yrs drops a NaN format");
    session
        .apply_update_v1(&remote(32531, Any::Number(12.0)))
        .unwrap();
    assert_eq!(run_kern(&session), Some(12.0));
}

#[test]
fn a_run_retyped_without_kern_is_saved_without_it() {
    let session = DeckSession::open(KERNED, 32519).unwrap();
    let context = EditCtx::local("test");
    let text = "OUR GREEN INITIATIVES";
    session
        .delete_text(&context, STORY, 0, text.encode_utf16().count() as u32)
        .unwrap();
    session
        .insert_text(
            &context,
            STORY,
            0,
            text,
            &TextStyle {
                font_size_pt: Some(32.0),
                ..Default::default()
            },
        )
        .unwrap();
    assert_eq!(run_kern(&session), None);
    let stored = session.encode_state_as_update_v1();
    let reattached = DeckSession::open_from_update_with_source(&stored, KERNED, 32520).unwrap();
    assert_eq!(run_kern(&reattached), None);
    // Only the list style's `kern` is left.
    let saved = slide_xml(&reattached.save().unwrap());
    assert_eq!(saved.matches(r#"kern="0""#).count(), 1, "{saved}");
}

#[test]
fn clearing_formatting_drops_a_runs_own_kern() {
    let session = DeckSession::open(KERNED, 32524).unwrap();
    let end = session.story(STORY).unwrap().length - 1;
    session
        .clear_text_formatting(&EditCtx::local("test"), STORY, 0, end, None)
        .unwrap();
    assert_eq!(run_kern(&session), None);
}

#[test]
fn reattaching_a_source_without_kern_leaves_the_session_untouched() {
    let stored = DeckSession::open(DECK, 32521)
        .unwrap()
        .encode_state_as_update_v1();
    let reattached = DeckSession::open_from_update_with_source(&stored, DECK, 32522).unwrap();
    assert_eq!(reattached.encode_state_as_update_v1(), stored);
    assert_eq!(run_kern(&reattached), None);
}
