use betteroffice_pptx::{EditCtx, Error, Presentation, ShapeNode};

const FIXTURE: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");
const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

#[test]
fn opens_edits_renders_saves_and_reopens() {
    let mut presentation = Presentation::open(FIXTURE).unwrap();
    assert_eq!(presentation.slides().len(), 3);
    assert_eq!(presentation.layouts().len(), 1);
    assert_eq!(presentation.masters().len(), 1);
    assert!(presentation.slides().iter().any(|slide| {
        slide.shapes.iter().any(|shape| {
            matches!(shape, ShapeNode::Shape(shape) if shape.text.as_ref().is_some_and(|body| {
                body.paragraphs.iter().flat_map(|paragraph| &paragraph.runs).any(|run| run.text.contains("Rust"))
            }))
        })
    }));

    let before = presentation.snapshot().unwrap();
    let slide_id = before.slides[0].id.clone();
    let shape_id = before.slides[0].shapes[0].id.clone();
    let receipt = presentation
        .move_shape(
            &EditCtx::local("facade-test"),
            &slide_id,
            &shape_id,
            1_111_111,
            2_222_222,
        )
        .unwrap();
    assert_eq!((receipt.after.x, receipt.after.y), (1_111_111, 2_222_222));
    let after = presentation.snapshot().unwrap();
    assert_eq!(
        (after.slides[0].shapes[0].x, after.slides[0].shapes[0].y),
        (1_111_111, 2_222_222)
    );

    for bold in [false, true] {
        presentation
            .register_font("Inter", bold, false, FONT)
            .unwrap();
    }
    let rendered = presentation.render_slide(0).unwrap();
    assert_eq!(
        (rendered.display_list.width, rendered.display_list.height),
        (1280.0, 720.0)
    );
    assert!(!rendered.display_list.primitives.is_empty());

    let saved = presentation.save().unwrap();
    let reopened = Presentation::open(&saved).unwrap();
    assert_eq!(reopened.slides().len(), 3);
    assert_eq!(reopened.package().presentation.slides.len(), 3);
    let snapshot = reopened.snapshot().unwrap();
    assert_eq!(
        (
            snapshot.slides[0].shapes[0].x,
            snapshot.slides[0].shapes[0].y
        ),
        (1_111_111, 2_222_222)
    );
}

#[test]
fn collaborative_facade_exchanges_typed_updates() {
    let left = Presentation::open_collaborative(FIXTURE, 101).unwrap();
    let right = Presentation::open_collaborative(FIXTURE, 202).unwrap();
    let before = left.snapshot().unwrap();
    let slide_id = before.slides[0].id.clone();
    let shape_id = before.slides[0].shapes[0].id.clone();

    left.move_shape(
        &EditCtx::local("left"),
        &slide_id,
        &shape_id,
        3_333_333,
        4_444_444,
    )
    .unwrap();
    let update = left
        .encode_diff_v1(&right.encode_state_vector_v1())
        .unwrap();
    let snapshot = right.apply_update_v1(&update).unwrap();

    assert_eq!(
        (
            snapshot.slides[0].shapes[0].x,
            snapshot.slides[0].shapes[0].y
        ),
        (3_333_333, 4_444_444)
    );
}

#[test]
fn facade_deletes_a_paragraph_break() {
    let presentation = Presentation::open(FIXTURE).unwrap();
    let story_id = presentation
        .snapshot()
        .unwrap()
        .slides
        .into_iter()
        .flat_map(|slide| slide.shapes)
        .flat_map(|shape| shape.text_stories)
        .map(|story| story.id)
        .next()
        .unwrap();
    let before = presentation.story(&story_id).unwrap();
    let context = EditCtx::local("facade-test");
    presentation
        .insert_paragraph_break(&context, &story_id, 1)
        .unwrap();
    presentation
        .delete_paragraph_break(&context, &story_id, 1)
        .unwrap();
    assert_eq!(presentation.story(&story_id).unwrap(), before);
}

#[test]
fn a_facade_seeded_update_accepts_the_source_file() {
    let presentation = Presentation::open(FIXTURE).unwrap();
    let update = presentation.encode_state_as_update_v1();
    let replica = pptx_edit::DeckSession::open_from_update_with_source(&update, FIXTURE, 303)
        .expect("the facade fingerprints the file bytes");
    assert!(replica.save().is_ok());
}

#[test]
fn reports_native_parse_errors() {
    assert!(matches!(
        Presentation::open(b"not a presentation"),
        Err(Error::Parse(_))
    ));
}

#[test]
fn facade_paragraph_proposals_preview_save_and_undo() {
    use betteroffice_pptx::{
        ProposalEdit, ProposalRequest, ProposalShapeChangeKind, ProposalTextChangeKind,
        TextParagraphDraft, TextRunDraft, TextStyle,
    };
    let fixture = include_bytes!("../../pptx-edit/tests/fixtures/proposal-shapes.pptx");
    let presentation = Presentation::open(fixture).unwrap();
    let before = presentation.snapshot().unwrap();
    let shape = &before.slides[0].shapes[0];
    let proposal = presentation
        .propose(ProposalRequest {
            agent_id: "agent".into(),
            note: None,
            edits: vec![ProposalEdit::SetStoryParagraphs {
                story_id: shape.text_stories[0].id.clone(),
                paragraphs: vec![
                    TextParagraphDraft {
                        alignment: Some("ctr".into()),
                        bullet: Some(true),
                        runs: vec![TextRunDraft {
                            text: "First bullet".into(),
                            style: TextStyle {
                                bold: Some(true),
                                ..Default::default()
                            },
                        }],
                    },
                    TextParagraphDraft {
                        alignment: Some("r".into()),
                        bullet: Some(false),
                        runs: vec![TextRunDraft {
                            text: "Second point".into(),
                            style: TextStyle::default(),
                        }],
                    },
                ],
            }],
        })
        .unwrap();
    let diff = presentation
        .layout_proposal_diff_slide(&proposal.id, 0)
        .unwrap();
    assert!(
        diff.text_changes
            .iter()
            .any(|change| change.kind == ProposalTextChangeKind::Insertion)
    );
    assert!(
        diff.text_changes
            .iter()
            .any(|change| change.kind == ProposalTextChangeKind::Deletion)
    );
    assert_eq!(diff.shape_changes[0].kind, ProposalShapeChangeKind::Changed);
    assert_eq!(presentation.snapshot().unwrap(), before);
    let preview = presentation.preview_proposal(&proposal.id).unwrap();
    assert_eq!(
        presentation
            .accept_proposal(&proposal.id, false)
            .unwrap()
            .snapshot,
        preview.snapshot
    );
    let reopened = Presentation::open(&presentation.save().unwrap()).unwrap();
    let story = &reopened.snapshot().unwrap().slides[0].shapes[0].text_stories[0];
    assert_eq!(story.plain_text(), "First bullet\nSecond point");
    assert_eq!(story.paragraphs[0].alignment.as_deref(), Some("ctr"));
    assert_eq!(story.paragraphs[0].runs[0].style.bold, Some(true));
    assert!(presentation.undo());
    assert_eq!(presentation.snapshot().unwrap(), before);
}

#[test]
fn facade_slide_property_proposals_render_inherit_accept_reject_and_undo() {
    use betteroffice_pptx::{ProposalEdit, ProposalRequest, ProposalStripStatus};
    let fixture = include_bytes!("../../pptx-edit/tests/fixtures/proposal-slide-props.pptx");
    for kind in ["setSlideLayout", "setSlideBackground"] {
        let presentation = Presentation::open(fixture).unwrap();
        let before = presentation.snapshot().unwrap();
        let slide = &before.slides[0];
        let edit: ProposalEdit = serde_json::from_value(if kind == "setSlideLayout" {
            serde_json::json!({"type": kind, "slideId": slide.id, "layoutPartPath": "ppt/slideLayouts/slideLayout2.xml", "mapping": [[slide.shapes[0].id, 8]]})
        } else {
            serde_json::json!({"type": kind, "slideId": slide.id, "background": {"type": "solid", "color": {"rgb": "123456"}}})
        }).unwrap();
        let request = || ProposalRequest {
            agent_id: "agent".into(),
            note: None,
            edits: vec![edit.clone()],
        };
        let proposal = presentation.propose(request()).unwrap();
        assert_eq!(
            presentation.proposal_slide_strip(&proposal.id).unwrap()[0].status,
            ProposalStripStatus::Changed
        );
        let by_id = presentation
            .layout_proposal_diff_slide_by_id(&proposal.id, &slide.id)
            .unwrap();
        let by_index = presentation
            .layout_proposal_diff_slide(&proposal.id, 0)
            .unwrap();
        assert_eq!(by_id.frame, by_index.frame);
        let frame = serde_json::to_string(&by_id.frame).unwrap();
        if kind == "setSlideLayout" {
            assert!(frame.contains("ppt/media/proposal-pixel.png"));
            assert!(frame.contains("#123456"));
        } else {
            assert!(frame.contains("#123456"));
        }
        assert_eq!(presentation.snapshot().unwrap(), before);
        assert!(presentation.reject_proposal(&proposal.id));
        let proposal = presentation.propose(request()).unwrap();
        let preview = presentation.preview_proposal(&proposal.id).unwrap();
        presentation.accept_proposal(&proposal.id, false).unwrap();
        assert_eq!(presentation.snapshot().unwrap(), preview.snapshot);
        assert!(presentation.undo());
        assert_eq!(presentation.snapshot().unwrap(), before);
    }
}

#[test]
fn facade_table_proposals_preview_accept_reject_and_undo() {
    use betteroffice_pptx::{ProposalEdit, ProposalRequest, ProposalShapeChangeKind};
    let mut parts = ooxml_opc::unzip_parts(include_bytes!(
        "../../pptx-render/tests/fixtures/table-basic.pptx"
    ))
    .unwrap();
    let (_, bytes) = parts
        .iter_mut()
        .find(|(name, _)| name == "ppt/slides/slide1.xml")
        .unwrap();
    let mut xml = String::from_utf8(bytes.clone()).unwrap();
    for attribute in [
        "gridSpan=\"2\"",
        "rowSpan=\"2\"",
        "hMerge=\"1\"",
        "vMerge=\"1\"",
    ] {
        xml = xml.replace(attribute, "");
    }
    *bytes = xml.into_bytes();
    let fixture = ooxml_opc::rezip_parts(&parts).unwrap();
    for kind in [
        "insertTableRow",
        "deleteTableRow",
        "insertTableColumn",
        "deleteTableColumn",
    ] {
        let presentation = Presentation::open(&fixture).unwrap();
        let before = presentation.snapshot().unwrap();
        let slide = &before.slides[0];
        let mut wire = serde_json::json!({
            "type": kind, "slideId": slide.id, "shapeId": slide.shapes[0].id, "index": 0,
        });
        if kind.starts_with("insert") {
            wire["position"] = "before".into();
            wire["cells"] = serde_json::json!(["Native cell"]);
        }
        let edit: ProposalEdit = serde_json::from_value(wire).unwrap();
        let request = || ProposalRequest {
            agent_id: "agent".into(),
            note: None,
            edits: vec![edit.clone()],
        };
        let proposal = presentation.propose(request()).unwrap();
        let preview = presentation.preview_proposal(&proposal.id).unwrap();
        let by_id = presentation
            .layout_proposal_diff_slide_by_id(&proposal.id, &slide.id)
            .unwrap();
        let indexed = presentation
            .layout_proposal_diff_slide(&proposal.id, 0)
            .unwrap();
        assert_eq!(by_id.frame, indexed.frame);
        assert_eq!(
            by_id.shape_changes[0].kind,
            ProposalShapeChangeKind::Changed
        );
        assert_eq!(
            by_id.snapshot.slides[0].shapes[0].graphic,
            preview.snapshot.slides[0].shapes[0].graphic
        );
        assert_eq!(presentation.snapshot().unwrap(), before);
        assert!(presentation.reject_proposal(&proposal.id));
        let proposal = presentation.propose(request()).unwrap();
        let accepted = presentation.preview_proposal(&proposal.id).unwrap();
        presentation.accept_proposal(&proposal.id, false).unwrap();
        assert_eq!(presentation.snapshot().unwrap(), accepted.snapshot);
        assert!(presentation.undo());
        assert_eq!(presentation.snapshot().unwrap(), before);
    }
}

#[test]
fn facade_reorder_proposals_preview_accept_reject_and_undo() {
    use betteroffice_pptx::{
        ProposalEdit, ProposalRequest, ProposalShapeChangeKind, ProposalShapeOrder,
    };
    let fixture = include_bytes!("../../pptx-edit/tests/fixtures/proposal-shapes.pptx");
    for to in [
        ProposalShapeOrder::Front,
        ProposalShapeOrder::Back,
        ProposalShapeOrder::Forward,
        ProposalShapeOrder::Backward,
    ] {
        let presentation = Presentation::open(fixture).unwrap();
        let before = presentation.snapshot().unwrap();
        let slide = &before.slides[0];
        let from = usize::from(matches!(
            to,
            ProposalShapeOrder::Back | ProposalShapeOrder::Backward
        ));
        let shape = &slide.shapes[from];
        let request = || ProposalRequest {
            agent_id: "agent".into(),
            note: None,
            edits: vec![ProposalEdit::ReorderShape {
                slide_id: slide.id.clone(),
                shape_id: shape.id.clone(),
                to,
            }],
        };
        let proposal = presentation.propose(request()).unwrap();
        let preview = presentation.preview_proposal(&proposal.id).unwrap();
        assert_eq!(preview.snapshot.slides[0].shapes[1 - from].id, shape.id);
        let by_id = presentation
            .layout_proposal_diff_slide_by_id(&proposal.id, &slide.id)
            .unwrap();
        let indexed = presentation
            .layout_proposal_diff_slide(&proposal.id, 0)
            .unwrap();
        assert_eq!(by_id.frame, indexed.frame);
        assert_eq!(
            by_id.shape_changes[0].kind,
            ProposalShapeChangeKind::Reordered
        );
        assert_eq!(presentation.snapshot().unwrap(), before);
        assert!(presentation.reject_proposal(&proposal.id));
        let proposal = presentation.propose(request()).unwrap();
        presentation.accept_proposal(&proposal.id, false).unwrap();
        assert_eq!(presentation.snapshot().unwrap(), preview.snapshot);
        assert!(presentation.undo());
        assert_eq!(presentation.snapshot().unwrap(), before);
        assert!(!presentation.can_undo());
    }
}
