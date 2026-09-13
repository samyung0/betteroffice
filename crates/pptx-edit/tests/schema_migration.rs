//! `deck-schema-v1.update.bin` was produced by release 4bdccdd: it opens
//! `betteroffice-demo.pptx`, adds a text box and edits its story, then persists
//! `encode_state_as_update_v1()`.

use pptx_edit::{DeckSession, DeckSnapshot, EditCtx, EditError, TextStyle};
use std::sync::Arc;
use yrs::updates::decoder::Decode;
use yrs::{Any, Doc, Map, MapRef, Out, ReadTxn, StateVector, Transact, Update};

const V1_UPDATE: &[u8] = include_bytes!("fixtures/deck-schema-v1.update.bin");
const SOURCE: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");
const META: &str = "pptx:meta";
const SHAPE_ID: &str = "shape:4242:0";
const STORY_ID: &str = "story:shape:4242:0:0";

#[test]
fn released_v1_snapshot_migrates_and_round_trips_as_v3() {
    assert_eq!(stamped_version(V1_UPDATE), Some(1.0));

    let session = DeckSession::open_from_update(V1_UPDATE, 901).unwrap();
    assert_v1_content(&session);

    let migrated = session.encode_state_as_update_v1();
    assert_eq!(stamped_version(&migrated), Some(3.0));
    assert!(
        package_json(&migrated).contains("\"charts\""),
        "the migrated package must carry the v2 chart field"
    );

    let reopened = DeckSession::open_from_update(&migrated, 902).unwrap();
    assert_v1_content(&reopened);
    assert_eq!(
        snapshot_shape_ids(&session.snapshot().unwrap()),
        snapshot_shape_ids(&reopened.snapshot().unwrap())
    );
    assert_eq!(
        reopened.encode_state_as_update_v1().len(),
        migrated.len(),
        "reopening a v3 snapshot must not migrate again"
    );
    let seeded = DeckSession::open(SOURCE, 909).unwrap();
    seeded.apply_update_v1(V1_UPDATE).unwrap();
    assert_v1_content(&seeded);
}

#[test]
fn a_migrated_session_still_edits() {
    let session = DeckSession::open_from_update(V1_UPDATE, 903).unwrap();
    session
        .insert_text(
            &EditCtx::local("test"),
            STORY_ID,
            0,
            "re-",
            &TextStyle::default(),
        )
        .unwrap();
    assert_eq!(
        session.story(STORY_ID).unwrap().plain_text(),
        "re-edited persisted on v1"
    );
    let reopened =
        DeckSession::open_from_update(&session.encode_state_as_update_v1(), 904).unwrap();
    assert_eq!(
        reopened.story(STORY_ID).unwrap().plain_text(),
        "re-edited persisted on v1"
    );
}

#[test]
fn two_clients_migrating_the_same_v1_snapshot_converge() {
    let left = DeckSession::open_from_update(V1_UPDATE, 907).unwrap();
    let right = DeckSession::open_from_update(V1_UPDATE, 908).unwrap();

    right
        .apply_update_v1(&left.encode_state_as_update_v1())
        .unwrap();
    left.apply_update_v1(&right.encode_state_as_update_v1())
        .unwrap();

    assert_eq!(left.snapshot().unwrap(), right.snapshot().unwrap());
    assert_eq!(
        stamped_version(&left.encode_state_as_update_v1()),
        Some(3.0)
    );
    assert_eq!(
        package_json(&left.encode_state_as_update_v1()),
        package_json(&right.encode_state_as_update_v1())
    );
    assert_eq!(left.package().media, right.package().media);
}

#[test]
fn binary_media_seeding_and_legacy_migration_preserve_bytes_without_json_history() {
    let mut parts = ooxml_opc::unzip_parts(SOURCE).unwrap();
    let (_, image) = parts
        .iter_mut()
        .find(|(path, _)| path.starts_with("ppt/media/"))
        .unwrap();
    image.extend((0..1024 * 1024).map(|index| index as u8));
    let source = ooxml_opc::rezip_parts(&parts).unwrap();
    let fresh = DeckSession::open(&source, 910).unwrap();
    let fresh_state = fresh.encode_state_as_update_v1();
    assert_binary_media(&fresh_state, &fresh.package().media);

    let legacy = hydrated(&fresh_state);
    {
        let mut txn = legacy.transact_mut();
        let meta = txn.get_map(META).unwrap();
        meta.insert(&mut txn, "schemaVersion", 2.0);
        meta.insert(
            &mut txn,
            "packageJson",
            Any::Buffer(Arc::from(serde_json::to_vec(fresh.package()).unwrap())),
        );
        meta.remove(&mut txn, "media");
    }
    let legacy_state = legacy
        .transact()
        .encode_state_as_update_v1(&StateVector::default());
    let left = DeckSession::open_from_update(&legacy_state, 911).unwrap();
    let right = DeckSession::open_from_update(&legacy_state, 912).unwrap();
    let expected_package: pptx_parse::PptxPackage =
        serde_json::from_slice(&serde_json::to_vec(fresh.package()).unwrap()).unwrap();
    assert_eq!(left.package(), &expected_package);
    assert_eq!(left.snapshot().unwrap(), fresh.snapshot().unwrap());
    assert_eq!(left.package().media, fresh.package().media);
    assert!(!left.can_undo());
    let migrated = left.encode_state_as_update_v1();
    assert_binary_media(&migrated, &fresh.package().media);
    assert!(
        migrated.len() < legacy_state.len() / 2,
        "legacy={} migrated={}",
        legacy_state.len(),
        migrated.len()
    );
    assert!(migrated.len() < fresh_state.len() + 1024);

    left.apply_update_v1(&right.encode_state_as_update_v1())
        .unwrap();
    left.apply_update_v1(&legacy_state).unwrap();
    right
        .apply_update_v1(&left.encode_state_as_update_v1())
        .unwrap();
    let merged = left.encode_state_as_update_v1();
    assert!(
        merged.len() < fresh_state.len() + 2048,
        "legacy or losing migration payload survived: {}",
        merged.len()
    );
    assert_binary_media(&merged, &fresh.package().media);
    assert_eq!(left.snapshot().unwrap(), right.snapshot().unwrap());

    let attached = DeckSession::open_from_update_with_source(&merged, &source, 913).unwrap();
    let exported = pptx_parse::parse_pptx(&attached.save().unwrap()).unwrap();
    assert_eq!(exported.media, fresh.package().media);
    eprintln!(
        "binary media bytes: legacy={}, fresh={}, migrated={}, merged={}",
        legacy_state.len(),
        fresh_state.len(),
        migrated.len(),
        merged.len()
    );
}

#[test]
fn legacy_text_updates_merge_after_migration_without_resurrecting_json_media() {
    let legacy = hydrated(V1_UPDATE);
    let vector = legacy.transact().state_vector();
    let migrated = DeckSession::open_from_update(V1_UPDATE, 914).unwrap();
    {
        use yrs::{Text, TextRef};
        let mut txn = legacy.transact_mut();
        let stories = txn.get_map("pptx:stories").unwrap();
        let story = stories
            .get(&txn, STORY_ID)
            .unwrap()
            .cast::<TextRef>()
            .unwrap();
        story.insert(&mut txn, 0, "late ");
    }
    let late = legacy.transact().encode_diff_v1(&vector);
    migrated.apply_update_v1(&late).unwrap();
    migrated.apply_update_v1(V1_UPDATE).unwrap();
    assert_eq!(
        migrated.story(STORY_ID).unwrap().plain_text(),
        "late edited persisted on v1"
    );
    assert_binary_media(
        &migrated.encode_state_as_update_v1(),
        &migrated.package().media,
    );
}

#[test]
fn fresh_editor_and_legacy_parent_converge_through_state_vector_sync() {
    for parent_first in [false, true] {
        for legacy_state in [V1_UPDATE.to_vec(), restamped(V1_UPDATE, Some(2.0))] {
            let parent = hydrated(&legacy_state);
            let editor = DeckSession::open(SOURCE, 918).unwrap();
            let initial_vector = editor.encode_state_vector_v1();
            if parent_first {
                sync_parent_to_editor(&parent, &editor);
                sync_editor_to_parent(&editor, &parent);
            } else {
                sync_editor_to_parent(&editor, &parent);
                sync_parent_to_editor(&parent, &editor);
            }
            sync_editor_to_parent(&editor, &parent);
            sync_parent_to_editor(&parent, &editor);

            let parent_state = parent
                .transact()
                .encode_state_as_update_v1(&StateVector::default());
            let editor_state = editor.encode_state_as_update_v1();
            assert_binary_media(&parent_state, &editor.package().media);
            assert_eq!(package_json(&parent_state), package_json(&editor_state));
            assert_eq!(parent_state, editor_state);
            assert_v1_content(&editor);
            let restored_parent = DeckSession::open_from_update(&parent_state, 919).unwrap();
            assert_v1_content(&restored_parent);
            assert_eq!(
                editor.snapshot().unwrap(),
                restored_parent.snapshot().unwrap()
            );
            assert_eq!(
                parent.transact().state_vector(),
                editor.yrs_doc().transact().state_vector()
            );
            assert_ne!(editor.encode_state_vector_v1(), initial_vector);
            assert!(!editor.can_undo());
        }
    }
}

fn sync_parent_to_editor(parent: &Doc, editor: &DeckSession) {
    let vector = StateVector::decode_v1(&editor.encode_state_vector_v1()).unwrap();
    editor
        .apply_update_v1(&parent.transact().encode_diff_v1(&vector))
        .unwrap();
}

fn sync_editor_to_parent(editor: &DeckSession, parent: &Doc) {
    use yrs::updates::encoder::Encode;
    let vector = parent.transact().state_vector().encode_v1();
    parent
        .transact_mut()
        .apply_update(Update::decode_v1(&editor.encode_diff_v1(&vector).unwrap()).unwrap())
        .unwrap();
}

#[test]
fn concurrent_legacy_migration_is_upgraded_and_broadcast_by_a_live_session() {
    use pptx_edit::UpdateOrigin;
    use std::sync::Mutex;

    let migrated = DeckSession::open_from_update(V1_UPDATE, 1).unwrap();
    let initial_migration = migrated.encode_state_as_update_v1();
    let legacy = Doc::with_client_id(999_999);
    legacy
        .transact_mut()
        .apply_update(Update::decode_v1(V1_UPDATE).unwrap())
        .unwrap();
    {
        let mut txn = legacy.transact_mut();
        let meta = txn.get_map(META).unwrap();
        meta.insert(&mut txn, "schemaVersion", 2.0);
        meta.insert(
            &mut txn,
            "packageJson",
            Any::Buffer(Arc::from(serde_json::to_vec(migrated.package()).unwrap())),
        );
    }
    let events = Arc::new(Mutex::new(Vec::new()));
    let observed = Arc::clone(&events);
    let _subscription = migrated
        .observe_update_v1(move |event| observed.lock().unwrap().push(event))
        .unwrap();
    migrated
        .apply_update_v1(
            &legacy
                .transact()
                .encode_state_as_update_v1(&StateVector::default()),
        )
        .unwrap();
    assert_v1_content(&migrated);
    assert_binary_media(
        &migrated.encode_state_as_update_v1(),
        &migrated.package().media,
    );
    assert!(!migrated.can_undo());
    assert!(
        events
            .lock()
            .unwrap()
            .iter()
            .any(|event| event.origin == UpdateOrigin::Local)
    );
    legacy
        .transact_mut()
        .apply_update(Update::decode_v1(&initial_migration).unwrap())
        .unwrap();
    for event in events.lock().unwrap().iter() {
        legacy
            .transact_mut()
            .apply_update(Update::decode_v1(&event.update).unwrap())
            .unwrap();
    }
    let converged = legacy
        .transact()
        .encode_state_as_update_v1(&StateVector::default());
    assert_binary_media(&converged, &migrated.package().media);
}

#[test]
fn current_schema_requires_binary_media_and_rejects_malformed_entries() {
    let seeded = DeckSession::open(SOURCE, 915)
        .unwrap()
        .encode_state_as_update_v1();
    for value in [
        None,
        Some(Any::Null),
        Some(Any::Array(Arc::from([Any::Array(Arc::from([
            Any::from("ppt/media/image.png"),
            Any::from("image/png"),
            Any::from("not bytes"),
        ]))]))),
    ] {
        let doc = hydrated(&seeded);
        {
            let mut txn = doc.transact_mut();
            let meta = txn.get_map(META).unwrap();
            match value {
                Some(value) => {
                    meta.insert(&mut txn, "media", value);
                }
                None => {
                    meta.remove(&mut txn, "media");
                }
            }
        }
        let invalid = doc
            .transact()
            .encode_state_as_update_v1(&StateVector::default());
        assert!(matches!(
            DeckSession::open_from_update(&invalid, 916),
            Err(EditError::InvalidState(_))
        ));
        let receiver = DeckSession::open_from_update(&seeded, 917).unwrap();
        assert!(matches!(
            receiver.apply_update_v1(&invalid),
            Err(EditError::InvalidState(_))
        ));
        assert_eq!(receiver.encode_state_as_update_v1(), seeded);
    }
}

fn assert_binary_media(update: &[u8], expected: &[pptx_parse::MediaPart]) {
    let json: serde_json::Value = serde_json::from_str(&package_json(update)).unwrap();
    assert_eq!(json["media"], serde_json::json!([]));
    assert_eq!(stamped_version(update), Some(3.0));
    let doc = hydrated(update);
    let txn = doc.transact();
    let Some(Out::Any(Any::Array(media))) = txn.get_map(META).unwrap().get(&txn, "media") else {
        panic!("missing binary media");
    };
    assert_eq!(media.len(), expected.len());
    for (entry, part) in media.iter().zip(expected) {
        assert_eq!(
            entry,
            &Any::Array(Arc::from([
                Any::from(part.part_path.as_str()),
                Any::from(part.content_type.as_str()),
                Any::Buffer(Arc::from(part.bytes.as_slice())),
            ]))
        );
    }
}

#[test]
fn unmigratable_schema_versions_stay_rejected() {
    for version in [0.0, 1.5, 5.0] {
        assert!(
            matches!(
                DeckSession::open_from_update(&restamped(V1_UPDATE, Some(version)), 905),
                Err(EditError::InvalidState(message))
                    if message == "unsupported deck schema version"
            ),
            "schema version {version} must be rejected"
        );
    }
    assert!(matches!(
        DeckSession::open_from_update(&restamped(V1_UPDATE, None), 906),
        Err(EditError::InvalidState(message))
            if message == "unsupported deck schema version"
    ));
}

fn assert_v1_content(session: &DeckSession) {
    let snapshot = session.snapshot().unwrap();
    assert_eq!(snapshot.width_emu, 12_192_000);
    assert_eq!(snapshot.height_emu, 6_858_000);
    assert_eq!(snapshot.slides.len(), 3);
    assert_eq!(snapshot.slides[0].id, "slide:0:256");
    assert!(
        snapshot_shape_ids(&snapshot)
            .iter()
            .any(|id| id == SHAPE_ID)
    );
    assert_eq!(
        session.story(STORY_ID).unwrap().plain_text(),
        "edited persisted on v1"
    );
    assert!(session.package().charts.is_empty());
}

fn snapshot_shape_ids(snapshot: &DeckSnapshot) -> Vec<String> {
    snapshot
        .slides
        .iter()
        .flat_map(|slide| slide.shapes.iter())
        .map(|shape| shape.id.clone())
        .collect()
}

fn hydrated(update: &[u8]) -> Doc {
    let doc = Doc::new();
    doc.transact_mut()
        .apply_update(Update::decode_v1(update).unwrap())
        .unwrap();
    doc
}

fn meta(doc: &Doc) -> MapRef {
    doc.transact().get_map(META).unwrap()
}

fn stamped_version(update: &[u8]) -> Option<f64> {
    let doc = hydrated(update);
    let meta = meta(&doc);
    match meta.get(&doc.transact(), "schemaVersion") {
        Some(Out::Any(Any::Number(value))) => Some(value),
        _ => None,
    }
}

fn package_json(update: &[u8]) -> String {
    let doc = hydrated(update);
    let meta = meta(&doc);
    match meta.get(&doc.transact(), "packageJson") {
        Some(Out::Any(Any::Buffer(bytes))) => String::from_utf8(bytes.to_vec()).unwrap(),
        _ => panic!("missing packageJson"),
    }
}

fn restamped(update: &[u8], version: Option<f64>) -> Vec<u8> {
    let doc = hydrated(update);
    let meta = meta(&doc);
    {
        let mut txn = doc.transact_mut();
        match version {
            Some(version) => {
                meta.insert(&mut txn, "schemaVersion", version);
            }
            None => {
                meta.remove(&mut txn, "schemaVersion");
            }
        }
    }
    doc.transact()
        .encode_state_as_update_v1(&StateVector::default())
}
