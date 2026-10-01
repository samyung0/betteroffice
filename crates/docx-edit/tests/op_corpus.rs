//! Editing-operation behavior corpus.

use std::collections::BTreeMap;
use std::sync::Arc;

use docx_edit::*;
use yrs::Any;

const DATE: &str = "2026-07-13T12:00:00Z";

fn ctx() -> EditCtx {
    EditCtx::local("Ada", DATE)
}

fn sug(author: &str) -> EditCtx {
    EditCtx::local(author, DATE).suggesting()
}

fn doc_with(text: &str) -> (EditingDoc, ParagraphId) {
    let doc = EditingDoc::new(7);
    let para = doc.create_story("body", text, "Normal", "left").unwrap();
    (doc, para)
}

fn active(attrs: &BTreeMap<String, Any>, key: &str) -> bool {
    matches!(attrs.get(key), Some(value) if *value != Any::Null)
}

fn revision_id_of(attrs: &BTreeMap<String, Any>, key: &str) -> Option<String> {
    let Any::Map(revision) = attrs.get(key)? else {
        return None;
    };
    match revision.get("id") {
        Some(Any::String(id)) => Some(id.to_string()),
        _ => None,
    }
}

fn revision_author_of(attrs: &BTreeMap<String, Any>, key: &str) -> Option<String> {
    let Any::Map(revision) = attrs.get(key)? else {
        return None;
    };
    match revision.get("author") {
        Some(Any::String(author)) => Some(author.to_string()),
        _ => None,
    }
}

/// Attributes of the first text segment containing `marker`.
fn seg_attrs(doc: &EditingDoc, marker: &str) -> BTreeMap<String, Any> {
    doc.story_segments("body")
        .unwrap()
        .into_iter()
        .find_map(|segment| match segment.content {
            SegmentContent::Text(value) if value.contains(marker) => Some(segment.attributes),
            _ => None,
        })
        .unwrap_or_else(|| panic!("text segment {marker:?} not found"))
}

fn raw_text(doc: &EditingDoc) -> String {
    doc.story_segments("body")
        .unwrap()
        .into_iter()
        .filter_map(|segment| match segment.content {
            SegmentContent::Text(value) => Some(value),
            _ => None,
        })
        .collect()
}

#[test]
fn mid_paragraph_breaks_keep_their_units_and_render_between_text_fragments() {
    for (open, close) in [
        ("", ""),
        ("<w:hyperlink w:anchor=\"target\">", "</w:hyperlink>"),
        ("<w:ins w:id=\"9\" w:author=\"Ada\">", "</w:ins>"),
    ] {
        let xml = format!(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="11111111">{open}<w:r><w:t>before</w:t><w:br w:type="page"/><w:t>after</w:t></w:r>{close}</w:p><w:p w14:paraId="22222222"><w:r><w:t>next</w:t></w:r></w:p></w:body></w:document>"#
        );
        let bytes =
            ooxml_opc::rezip_parts(&[("word/document.xml".into(), xml.into_bytes())]).unwrap();
        let doc = EditingDoc::new(7);
        seed_from_docx(&doc, &bytes).unwrap();
        assert_eq!(slot_units(&doc), "before[pageBreak]after¶next¶");
        doc.insert_text(&ctx(), Position::new("body", 0), "Q", FormatPolicy::Inherit)
            .unwrap();
        for para in ["11111111", "22222222"] {
            doc.set_paragraph_attr(
                para,
                "numPr",
                Any::Map(Arc::new(
                    [(String::from("numId"), Any::Number(1.0))]
                        .into_iter()
                        .collect(),
                )),
            )
            .unwrap();
            doc.set_paragraph_attr(para, "listMarker", Any::String("%1.".into()))
                .unwrap();
            doc.set_paragraph_attr(para, "listNumFmt", Any::String("decimal".into()))
                .unwrap();
        }
        let blocks =
            bridge::yrs_doc_to_layout_blocks(&doc, "body", &bridge::RenderEnv::default()).unwrap();
        let value = serde_json::to_value(&blocks).unwrap();
        assert_eq!(
            value
                .as_array()
                .unwrap()
                .iter()
                .map(|block| block["kind"].as_str().unwrap())
                .collect::<Vec<_>>(),
            ["paragraph", "pageBreak", "paragraph", "paragraph"]
        );
        assert_eq!(value[0]["attrs"]["listMarker"], "1.");
        assert!(value[2]["attrs"].get("listMarker").is_none());
        assert_eq!(value[3]["attrs"]["listMarker"], "2.");
        doc.delete_range(&ctx(), StoryRange::new("body", 7, 8))
            .unwrap();
        assert_eq!(slot_units(&doc), "Qbeforeafter¶next¶");
    }
}

#[test]
fn enter_after_a_source_trailing_column_break_keeps_the_new_line_after_it() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="11111111"><w:r><w:t>prev</w:t><w:br w:type="column"/></w:r></w:p><w:p w14:paraId="22222222"><w:r><w:t>Heading</w:t></w:r></w:p></w:body></w:document>"#;
    let bytes = ooxml_opc::rezip_parts(&[("word/document.xml".into(), xml.to_vec())]).unwrap();
    let doc = EditingDoc::new(7);
    seed_from_docx(&doc, &bytes).unwrap();
    doc.split_paragraph(&ctx(), Position::new("body", 5), None)
        .unwrap();
    assert_eq!(slot_units(&doc), "prev¶[columnBreak]¶Heading¶");
}

#[test]
fn continued_field_ends_follow_edits_and_disappear_with_their_owner() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="11111111"><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>first</w:t></w:r></w:p><w:p w14:paraId="22222222"><w:r><w:t>second</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t>after</w:t></w:r></w:p></w:body></w:document>"#;
    let bytes = ooxml_opc::rezip_parts(&[("word/document.xml".into(), xml.to_vec())]).unwrap();
    let doc = EditingDoc::new(7);
    seed_from_docx(&doc, &bytes).unwrap();
    let ends = |doc: &EditingDoc| {
        marks(doc)
            .into_iter()
            .flat_map(|properties| {
                let Some(Any::Array(markers)) = properties.get("bookmarks") else {
                    return Vec::new();
                };
                markers
                    .iter()
                    .filter_map(|marker| {
                        let Any::Map(marker) = marker else {
                            return None;
                        };
                        (marker.get("kind") == Some(&Any::String("fieldend".into())))
                            .then(|| marker["offset"].clone())
                    })
                    .collect::<Vec<_>>()
            })
            .collect::<Vec<_>>()
    };
    assert_eq!(ends(&doc), vec![Any::Number(6.0)]);
    doc.insert_text(&ctx(), Position::new("body", 2), "x", FormatPolicy::Inherit)
        .unwrap();
    assert_eq!(ends(&doc), vec![Any::Number(7.0)]);
    doc.split_paragraph(&ctx(), Position::new("body", 5), None)
        .unwrap();
    assert_eq!(ends(&doc), vec![Any::Number(4.0)]);
    doc.delete_range(&ctx(), StoryRange::new("body", 0, 1))
        .unwrap();
    assert!(ends(&doc).is_empty());
}

#[test]
fn nested_continued_fields_keep_both_owners_boundaries() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="11111111"><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>IF</w:instrText></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>PAGE</w:instrText></w:r></w:p><w:p w14:paraId="22222222"><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>yes</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:body></w:document>"#;
    let bytes = ooxml_opc::rezip_parts(&[("word/document.xml".into(), xml.to_vec())]).unwrap();
    let doc = EditingDoc::new(7);
    seed_from_docx(&doc, &bytes).unwrap();
    let paragraphs = doc.paragraphs("body").unwrap();
    let Some(Any::Array(markers)) = paragraphs[1].properties.get("bookmarks") else {
        panic!("nested field boundaries are missing");
    };
    assert_eq!(markers.len(), 4);
    doc.delete_range(&ctx(), StoryRange::new("body", 0, 1))
        .unwrap();
    assert!(
        !doc.paragraphs("body").unwrap()[1]
            .properties
            .contains_key("bookmarks")
    );
}

#[test]
fn coincident_markers_keep_source_order_and_starts_before_ends() {
    let kinds = |doc: &EditingDoc, paragraph: usize| {
        let Any::Array(markers) =
            doc.paragraphs("body").unwrap()[paragraph].properties["bookmarks"].clone()
        else {
            panic!("markers are missing");
        };
        markers
            .iter()
            .map(
                |marker| match (map_get(marker, "kind"), map_get(marker, "id")) {
                    (Some(Any::String(kind)), Some(Any::Number(id))) => format!("{kind}{id}"),
                    (Some(Any::String(kind)), _) => kind.to_string(),
                    _ => panic!("marker without a kind"),
                },
            )
            .collect::<Vec<_>>()
            .join(" ")
    };
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="11111111"><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>IF</w:instrText></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>PAGE</w:instrText></w:r></w:p><w:p w14:paraId="22222222"><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:bookmarkStart w:id="5" w:name="m"/><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:bookmarkEnd w:id="5"/><w:r><w:t>yes</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:body></w:document>"#;
    let bytes = ooxml_opc::rezip_parts(&[("word/document.xml".into(), xml.to_vec())]).unwrap();
    for client in 7..12 {
        let doc = EditingDoc::new(client);
        seed_from_docx(&doc, &bytes).unwrap();
        assert_eq!(
            kinds(&doc, 1),
            "fieldseparate fieldend start5 fieldseparate end5 fieldend"
        );
    }
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="11111111"><w:r><w:t>ab</w:t></w:r><w:bookmarkStart w:id="3" w:name="a"/><w:bookmarkStart w:id="4" w:name="b"/><w:r><w:t>c</w:t></w:r></w:p><w:p w14:paraId="22222222"><w:r><w:t>d</w:t></w:r><w:bookmarkStart w:id="5" w:name="c"/><w:bookmarkEnd w:id="3"/><w:r><w:t>e</w:t></w:r><w:bookmarkEnd w:id="4"/><w:bookmarkEnd w:id="5"/><w:r><w:t>f</w:t></w:r></w:p></w:body></w:document>"#;
    let bytes = ooxml_opc::rezip_parts(&[("word/document.xml".into(), xml.to_vec())]).unwrap();
    let doc = EditingDoc::new(7);
    seed_from_docx(&doc, &bytes).unwrap();
    doc.delete_range(&ctx(), StoryRange::new("body", 2, 6))
        .unwrap();
    assert_eq!(kinds(&doc, 0), "start3 end3 start4 end4 start5 end5");
}

#[test]
fn a_continued_fields_cached_blocks_do_not_keep_deleted_owners_alive() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="11111111"><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>IF</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p><w:p w14:paraId="22222222"><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>PAGE</w:instrText></w:r></w:p><w:p w14:paraId="33333333"><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:body></w:document>"#;
    let bytes = ooxml_opc::rezip_parts(&[("word/document.xml".into(), xml.to_vec())]).unwrap();
    let doc = EditingDoc::new(7);
    seed_from_docx(&doc, &bytes).unwrap();
    doc.delete_range(&ctx(), StoryRange::new("body", 2, 3))
        .unwrap();
    let paragraphs = doc.paragraphs("body").unwrap();
    let Some(Any::Array(markers)) = paragraphs[2].properties.get("bookmarks") else {
        panic!("outer field's end is missing");
    };
    assert_eq!(markers.len(), 1);
    assert_eq!(
        map_get(&markers[0], "kind"),
        Some(&Any::String("fieldend".into()))
    );
}

#[test]
fn continued_field_boundaries_keep_and_resolve_their_revision() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="11111111"><w:sdt><w:sdtContent><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC</w:instrText></w:r></w:sdtContent></w:sdt></w:p><w:p w14:paraId="22222222"><w:ins w:id="9" w:author="Ada"><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>result</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:ins></w:p></w:body></w:document>"#;
    let bytes = ooxml_opc::rezip_parts(&[("word/document.xml".into(), xml.to_vec())]).unwrap();
    let doc = EditingDoc::new(7);
    seed_from_docx(&doc, &bytes).unwrap();
    let markers = |doc: &EditingDoc| {
        let paragraphs = doc.paragraphs("body").unwrap();
        let Any::Array(markers) = paragraphs[1].properties["bookmarks"].clone() else {
            panic!("continued field boundaries are missing");
        };
        markers
    };
    assert_eq!(markers(&doc).len(), 2);
    assert!(markers(&doc).iter().all(|marker| {
        map_get(marker, "attributes")
            .and_then(|attrs| map_get(attrs, "ins"))
            .is_some()
    }));
    doc.accept_change(&ctx(), &ChangeTarget::All).unwrap();
    assert_eq!(markers(&doc).len(), 2);
    assert!(markers(&doc).iter().all(|marker| {
        map_get(marker, "attributes")
            .and_then(|attrs| map_get(attrs, "ins"))
            .is_none()
    }));
}

#[test]
fn seeded_bookmarks_follow_typing_splits_joins_and_remote_typing() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="11111111"><w:r><w:t>A</w:t></w:r><w:bookmarkStart w:id="5" w:name="mark"/><w:r><w:t>BC</w:t></w:r><w:bookmarkEnd w:id="5"/><w:r><w:t>D</w:t></w:r></w:p></w:body></w:document>"#;
    let bytes = ooxml_opc::rezip_parts(&[("word/document.xml".into(), xml.to_vec())]).unwrap();
    let doc = EditingDoc::new(7);
    seed_from_docx(&doc, &bytes).unwrap();
    let peer = EditingDoc::new(8);
    peer.apply_update_v1(&doc.encode_state_as_update_v1())
        .unwrap();
    let boundaries = |doc: &EditingDoc| {
        doc.paragraphs("body")
            .unwrap()
            .into_iter()
            .map(|paragraph| {
                let Some(Any::Array(bookmarks)) = paragraph.properties.get("bookmarks") else {
                    return Vec::new();
                };
                bookmarks
                    .iter()
                    .map(|marker| {
                        (
                            map_get(marker, "kind").cloned().unwrap(),
                            map_get(marker, "offset").cloned().unwrap(),
                        )
                    })
                    .collect::<Vec<_>>()
            })
            .collect::<Vec<_>>()
    };
    doc.insert_text(&ctx(), Position::new("body", 0), "Z", FormatPolicy::Inherit)
        .unwrap();
    doc.split_paragraph(&ctx(), Position::new("body", 3), None)
        .unwrap();
    assert_eq!(
        boundaries(&doc),
        vec![
            vec![(Any::from("start"), Any::Number(2.0))],
            vec![(Any::from("end"), Any::Number(1.0))]
        ]
    );
    doc.merge_paragraphs(&ctx(), "11111111", MergeDirection::Forward)
        .unwrap();
    assert_eq!(
        boundaries(&doc),
        vec![vec![
            (Any::from("start"), Any::Number(2.0)),
            (Any::from("end"), Any::Number(4.0))
        ]]
    );
    peer.insert_text(&ctx(), Position::new("body", 0), "Y", FormatPolicy::Inherit)
        .unwrap();
    doc.apply_update_v1(&peer.encode_state_as_update_v1())
        .unwrap();
    peer.apply_update_v1(&doc.encode_state_as_update_v1())
        .unwrap();
    assert_eq!(boundaries(&doc), boundaries(&peer));
    assert_eq!(
        boundaries(&doc),
        vec![vec![
            (Any::from("start"), Any::Number(3.0)),
            (Any::from("end"), Any::Number(5.0))
        ]]
    );
}

#[test]
fn empty_and_deleted_bookmarks_stay_before_new_text() {
    for text in ["", "AB"] {
        let xml = format!(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="11111111"><w:bookmarkStart w:id="5" w:name="mark"/><w:r><w:t>{text}</w:t></w:r><w:bookmarkEnd w:id="5"/></w:p></w:body></w:document>"#
        );
        let bytes =
            ooxml_opc::rezip_parts(&[("word/document.xml".into(), xml.into_bytes())]).unwrap();
        let doc = EditingDoc::new(7);
        seed_from_docx(&doc, &bytes).unwrap();
        if !text.is_empty() {
            doc.delete_range(&ctx(), StoryRange::new("body", 0, 2))
                .unwrap();
        }
        doc.insert_text(&ctx(), Position::new("body", 0), "Z", FormatPolicy::Inherit)
            .unwrap();
        let paras = doc.paragraphs("body").unwrap();
        let Any::Array(markers) = &paras[0].properties["bookmarks"] else {
            panic!("missing bookmarks");
        };
        assert_eq!(markers.len(), 2);
        for marker in markers.iter() {
            assert_eq!(map_get(marker, "offset"), Some(&Any::Number(0.0)));
        }
    }
}

#[test]
fn imported_run_property_changes_are_listed_and_resolve_after_typing() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:rPr><w:b/><w:rPrChange w:id="4" w:author="A"><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr><w:t>changed</w:t></w:r></w:p></w:body></w:document>"#;
    let bytes = ooxml_opc::rezip_parts(&[("word/document.xml".into(), xml.to_vec())]).unwrap();
    for accept in [false, true] {
        let doc = EditingDoc::new(7);
        seed_from_docx(&doc, &bytes).unwrap();
        let changes = doc.list_changes("body").unwrap();
        assert_eq!(changes.len(), 1);
        assert_eq!(changes[0].kind, ChangeKind::RunPropertiesChanged);
        doc.insert_text(&ctx(), Position::new("body", 2), "Z", FormatPolicy::Inherit)
            .unwrap();
        let target = ChangeTarget::Revision("4".into());
        if accept {
            doc.accept_change(&ctx(), &target).unwrap();
        } else {
            doc.reject_change(&ctx(), &target).unwrap();
        }
        assert!(doc.list_changes("body").unwrap().is_empty());
        let attrs = seg_attrs(&doc, "ch");
        assert_eq!(active(&attrs, "bold"), accept);
        assert_eq!(active(&attrs, "italic"), !accept);
        assert!(!active(&attrs, "rPrChange"));
        assert_eq!(raw_text(&doc), "chZanged");
    }
}

#[test]
fn accepting_a_field_link_uncovered_with_bookmarks_anchors_them_after_inserting_its_text() {
    let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> REF mark </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:ins w:id="4" w:author="A"><w:hyperlink w:anchor="mark"><w:r><w:t>A</w:t></w:r><w:bookmarkStart w:id="5" w:name="mark"/><w:r><w:t>BC</w:t></w:r><w:bookmarkEnd w:id="5"/><w:r><w:t>D</w:t></w:r></w:hyperlink></w:ins><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:body></w:document>"#;
    let bytes = ooxml_opc::rezip_parts(&[("word/document.xml".into(), xml.to_vec())]).unwrap();
    let doc = EditingDoc::new(7);
    seed_from_docx(&doc, &bytes).unwrap();
    doc.accept_change(&ctx(), &ChangeTarget::All).unwrap();
    assert_eq!(slot_units(&doc), "ABCD[field]¶");
    let paras = doc.paragraphs("body").unwrap();
    let Any::Array(markers) = &paras[0].properties["bookmarks"] else {
        panic!("missing bookmarks")
    };
    assert_eq!(markers.len(), 2);
    assert_eq!(map_get(&markers[0], "offset"), Some(&Any::Number(1.0)));
    assert_eq!(map_get(&markers[1], "offset"), Some(&Any::Number(3.0)));
}

fn map_get<'a>(value: &'a Any, key: &str) -> Option<&'a Any> {
    match value {
        Any::Map(map) => map.get(key),
        _ => None,
    }
}

// ---------------------------------------------------------------------------
// Text ops
// ---------------------------------------------------------------------------

#[test]
fn delete_range_suggesting_is_three_way() {
    let (doc, _) = doc_with("abcdef");
    // Bob's own pending insertion.
    doc.insert_text(
        &sug("Bob"),
        Position::new("body", 3),
        "XY",
        FormatPolicy::Plain,
    )
    .unwrap();
    assert_eq!(raw_text(&doc), "abcXYdef");
    // Bob suggest-deletes a range covering plain text AND his own insertion.
    let receipt = doc
        .delete_range(&sug("Bob"), StoryRange::new("body", 1, 7))
        .unwrap();
    assert_eq!(receipt.revision_ids.len(), 1);
    // Own insertion physically removed; retained text del-stamped by Bob. ("bc" and "de" stay
    // separate yrs chunks after the removal between them.)
    assert_eq!(raw_text(&doc), "abcdef");
    for marker in ["bc", "de"] {
        let stamped = seg_attrs(&doc, marker);
        assert_eq!(revision_author_of(&stamped, "del").as_deref(), Some("Bob"));
        assert_eq!(
            revision_id_of(&stamped, "del").as_ref(),
            Some(&receipt.revision_ids[0])
        );
    }
    assert!(!active(&seg_attrs(&doc, "a"), "del"));
    // Vanilla still shows the deleted text (pending until accepted).
    assert_eq!(
        doc.para_text(
            &doc.paragraphs("body").unwrap()[0].para_id,
            TextView::Vanilla
        )
        .unwrap(),
        "abcdef"
    );
}

#[test]
fn delete_range_suggesting_marks_pilcrow_ppr_del_and_keeps_other_authors_ins() {
    let (doc, first) = doc_with("onetwo");
    let split = doc
        .split_paragraph(&ctx(), Position::new("body", 3), None)
        .unwrap();
    assert_eq!(split.first_para_id, first);
    // Alice suggests an insertion; Bob suggest-deletes across it and the pilcrow.
    doc.insert_text(
        &sug("Alice"),
        Position::new("body", 2),
        "ZZ",
        FormatPolicy::Plain,
    )
    .unwrap();
    // Story: "onZZe" ¶ "two" ¶ — delete from 1 to 7 (covers Alice's ins + first pilcrow).
    doc.delete_range(&sug("Bob"), StoryRange::new("body", 1, 7))
        .unwrap();
    // Alice's pending insertion is NOT Bob's own → retained, nested ins+del.
    let nested = seg_attrs(&doc, "ZZ");
    assert_eq!(revision_author_of(&nested, "ins").as_deref(), Some("Alice"));
    assert_eq!(revision_author_of(&nested, "del").as_deref(), Some("Bob"));
    // The boundary pilcrow is retained and carries pPrDel; both paragraphs still exist.
    let paragraphs = doc.paragraphs("body").unwrap();
    assert_eq!(paragraphs.len(), 2);
    assert!(active(&paragraphs[0].properties, "pPrDel"));
    // The story's FINAL pilcrow is never stamped.
    assert!(!active(&paragraphs[1].properties, "pPrDel"));
}

#[test]
fn plain_delete_across_pilcrow_adopts_first_paragraph_identity() {
    let (doc, first) = doc_with("aaabbb");
    let split = doc
        .split_paragraph(&ctx(), Position::new("body", 3), None)
        .unwrap();
    let second = split.second_para_id.clone();
    // Give the FIRST paragraph distinctive pPr.
    doc.set_paragraph_attrs(
        &ctx(),
        &ParaSelector::One(first.clone()),
        &ParaAttrDelta {
            alignment: Patch::Set("right".into()),
            ..ParaAttrDelta::default()
        },
    )
    .unwrap();
    // Delete across the boundary pilcrow (units 2..5 cover "a", ¶, "b").
    doc.delete_range(&ctx(), StoryRange::new("body", 2, 5))
        .unwrap();
    let paragraphs = doc.paragraphs("body").unwrap();
    assert_eq!(paragraphs.len(), 1);
    assert_eq!(paragraphs[0].text, "aabb");
    // R6: the survivor adopted the FIRST affected paragraph's paraId + pPr.
    assert_eq!(paragraphs[0].para_id, first);
    assert_ne!(paragraphs[0].para_id, second);
    assert_eq!(
        paragraphs[0].properties.get("alignment"),
        Some(&Any::from("right"))
    );
}

#[test]
fn plain_delete_never_removes_final_pilcrow() {
    let (doc, first) = doc_with("abc");
    // Range covers the entire story including the final pilcrow.
    doc.delete_range(&ctx(), StoryRange::new("body", 0, 4))
        .unwrap();
    let paragraphs = doc.paragraphs("body").unwrap();
    assert_eq!(paragraphs.len(), 1);
    assert_eq!(paragraphs[0].text, "");
    assert_eq!(paragraphs[0].para_id, first);
}

#[test]
fn replace_range_suggesting_shares_one_revision_id() {
    let (doc, _) = doc_with("alpha beta gamma");
    let range = doc
        .resolve_search("body", None, "beta", TextView::Vanilla)
        .unwrap();
    let story_range = doc.locate_range(&range).unwrap();
    let receipt = doc.replace_range(&sug("Bob"), story_range, "BETA").unwrap();
    assert_eq!(receipt.revision_ids.len(), 1);
    let shared = &receipt.revision_ids[0];
    let del_attrs = seg_attrs(&doc, "beta");
    let ins_attrs = seg_attrs(&doc, "BETA");
    assert_eq!(revision_id_of(&del_attrs, "del").as_ref(), Some(shared));
    assert_eq!(revision_id_of(&ins_attrs, "ins").as_ref(), Some(shared));
    // Vanilla: pending insertion invisible, pending deletion still present.
    let para = doc.paragraphs("body").unwrap()[0].para_id.clone();
    assert_eq!(
        doc.para_text(&para, TextView::Vanilla).unwrap(),
        "alpha beta gamma"
    );
    assert_eq!(
        doc.para_text(&para, TextView::Raw).unwrap(),
        "alpha BETAbeta gamma"
    );
}

#[test]
fn same_author_typing_and_splits_coalesce_to_one_revision_id() {
    let (doc, _) = doc_with("");
    let first = doc
        .insert_text(
            &sug("Ada"),
            Position::new("body", 0),
            "abc",
            FormatPolicy::Inherit,
        )
        .unwrap()
        .revision_ids[0]
        .clone();
    let split = doc
        .split_paragraph(&sug("Ada"), Position::new("body", 3), None)
        .unwrap();
    assert_eq!(split.revision_ids, vec![first.clone()]);
    let second = doc
        .insert_text(
            &sug("Ada"),
            Position::new("body", 4),
            "def",
            FormatPolicy::Inherit,
        )
        .unwrap();
    assert_eq!(second.revision_ids, vec![first.clone()]);
    let second_split = doc
        .split_paragraph(&sug("Ada"), Position::new("body", 7), None)
        .unwrap();
    assert_eq!(second_split.revision_ids, vec![first.clone()]);
    let third = doc
        .insert_text(
            &sug("Ada"),
            Position::new("body", 8),
            "ghi",
            FormatPolicy::Inherit,
        )
        .unwrap();
    assert_eq!(third.revision_ids, vec![first.clone()]);

    let ids: std::collections::HashSet<_> = doc
        .list_changes("body")
        .unwrap()
        .into_iter()
        .map(|change| change.revision_id)
        .collect();
    assert_eq!(ids, std::collections::HashSet::from([first]));
}

#[test]
fn backspace_over_own_pending_split_retracts_the_paragraph_mark() {
    let (doc, _) = doc_with("abc");
    let split = doc
        .split_paragraph(&sug("Ada"), Position::new("body", 3), None)
        .unwrap();
    doc.merge_paragraphs(&sug("Ada"), &split.second_para_id, MergeDirection::Backward)
        .unwrap();

    let paragraphs = doc.paragraphs("body").unwrap();
    assert_eq!(paragraphs.len(), 1);
    assert_eq!(paragraphs[0].text, "abc");
    assert!(doc.list_changes("body").unwrap().is_empty());
}

#[test]
fn paragraph_properties_coalesce_with_same_author_inserted_text() {
    let (doc, para_id) = doc_with("");
    let revision_id = doc
        .insert_text(
            &sug("Ada"),
            Position::new("body", 0),
            "item",
            FormatPolicy::Inherit,
        )
        .unwrap()
        .revision_ids[0]
        .clone();
    let mut delta = ParaAttrDelta::default();
    delta.other.insert(
        "numPr".to_owned(),
        Some(Any::Map(Arc::new(std::collections::HashMap::from([
            ("numId".to_owned(), Any::Number(2.0)),
            ("ilvl".to_owned(), Any::Number(0.0)),
        ])))),
    );
    let formatting = doc
        .set_paragraph_attrs(&sug("Ada"), &ParaSelector::One(para_id), &delta)
        .unwrap();
    assert_eq!(formatting.revision_ids, vec![revision_id.clone()]);
    let ids: std::collections::HashSet<_> = doc
        .list_changes("body")
        .unwrap()
        .into_iter()
        .map(|change| change.revision_id)
        .collect();
    assert_eq!(ids, std::collections::HashSet::from([revision_id]));
}

#[test]
fn same_author_typing_reuses_a_pending_paragraph_property_revision() {
    let (doc, para_id) = doc_with("");
    let mut delta = ParaAttrDelta::default();
    delta.other.insert(
        "numPr".to_owned(),
        Some(Any::Map(Arc::new(std::collections::HashMap::from([
            ("numId".to_owned(), Any::Number(2.0)),
            ("ilvl".to_owned(), Any::Number(0.0)),
        ])))),
    );
    let formatting_id = doc
        .set_paragraph_attrs(&sug("Ada"), &ParaSelector::One(para_id), &delta)
        .unwrap()
        .revision_ids[0]
        .clone();

    let typing = doc
        .insert_text(
            &sug("Ada"),
            Position::new("body", 0),
            "item",
            FormatPolicy::Inherit,
        )
        .unwrap();
    assert_eq!(typing.revision_ids, vec![formatting_id.clone()]);
    let ids: std::collections::HashSet<_> = doc
        .list_changes("body")
        .unwrap()
        .into_iter()
        .map(|change| change.revision_id)
        .collect();
    assert_eq!(ids, std::collections::HashSet::from([formatting_id]));
}

#[test]
fn replace_range_plain_keeps_first_replaced_formatting() {
    let (doc, _) = doc_with("alpha beta gamma");
    doc.format_range(
        &ctx(),
        StoryRange::new("body", 6, 10),
        &InlineFormatDelta {
            bold: Patch::Set(true),
            ..InlineFormatDelta::default()
        },
    )
    .unwrap();
    doc.replace_range(&ctx(), StoryRange::new("body", 6, 10), "BETA")
        .unwrap();
    assert_eq!(raw_text(&doc), "alpha BETA gamma");
    // Type-over keeps the replaced text's formatting.
    assert_eq!(seg_attrs(&doc, "BETA").get("bold"), Some(&Any::Bool(true)));
}

#[test]
fn insert_text_rejects_breaks_and_empty_ranges_error() {
    let (doc, _) = doc_with("abc");
    assert_eq!(
        doc.insert_text(
            &ctx(),
            Position::new("body", 1),
            "a\nb",
            FormatPolicy::Plain
        ),
        Err(OpError::TextContainsBreak)
    );
    assert_eq!(
        doc.delete_range(&ctx(), StoryRange::new("body", 1, 1)),
        Err(OpError::EmptyRange)
    );
    assert_eq!(
        doc.toggle_format(&ctx(), StoryRange::new("body", 1, 1), SimpleFormat::Bold),
        Err(OpError::EmptyRange)
    );
}

#[test]
fn insert_text_inherit_copies_formatting_but_never_tracked_changes() {
    let (doc, _) = doc_with("ab");
    doc.format_range(
        &ctx(),
        StoryRange::new("body", 0, 2),
        &InlineFormatDelta {
            bold: Patch::Set(true),
            ..InlineFormatDelta::default()
        },
    )
    .unwrap();
    // Suggest-insert (carries ins) then type after it with Inherit.
    doc.insert_text(
        &sug("Alice"),
        Position::new("body", 2),
        "X",
        FormatPolicy::Plain,
    )
    .unwrap();
    doc.insert_text(&ctx(), Position::new("body", 2), "Q", FormatPolicy::Inherit)
        .unwrap();
    let attrs = seg_attrs(&doc, "Q");
    assert_eq!(attrs.get("bold"), Some(&Any::Bool(true)));
    assert!(!active(&attrs, "ins"), "ins must not be inherited");
    // Plain policy inherits nothing.
    doc.insert_text(&ctx(), Position::new("body", 0), "P", FormatPolicy::Plain)
        .unwrap();
    assert!(!active(&seg_attrs(&doc, "P"), "bold"));
}

#[test]
fn tab_is_a_char_and_break_is_an_embed_both_invisible_in_vanilla() {
    let (doc, para) = doc_with("ab");
    doc.insert_tab(&ctx(), Position::new("body", 1)).unwrap();
    doc.insert_hard_break(&ctx(), Position::new("body", 3))
        .unwrap();
    assert_eq!(doc.para_text(&para, TextView::Raw).unwrap(), "a\tb\n");
    assert_eq!(doc.para_text(&para, TextView::Vanilla).unwrap(), "ab");
    assert_eq!(doc.story_len("body").unwrap(), 5); // a, tab, b, break, pilcrow
}

// ---------------------------------------------------------------------------
// Split / merge
// ---------------------------------------------------------------------------

#[test]
fn split_first_half_keeps_original_para_id_and_full_ppr() {
    let (doc, original) = doc_with("hello world");
    doc.set_paragraph_attr(&original, "borders", Any::from("boxed"))
        .unwrap();
    doc.set_paragraph_attrs(
        &ctx(),
        &ParaSelector::One(original.clone()),
        &ParaAttrDelta {
            alignment: Patch::Set("center".into()),
            ..ParaAttrDelta::default()
        },
    )
    .unwrap();
    let split = doc
        .split_paragraph(&ctx(), Position::new("body", 5), None)
        .unwrap();
    assert_eq!(split.first_para_id, original);
    assert_ne!(split.second_para_id, original);
    let paragraphs = doc.paragraphs("body").unwrap();
    assert_eq!(paragraphs[0].text, "hello");
    assert_eq!(paragraphs[1].text, " world");
    assert_eq!(paragraphs[0].para_id, original);
    assert_eq!(paragraphs[1].para_id, split.second_para_id);
    // First half: full pPr copy including borders.
    assert_eq!(
        paragraphs[0].properties.get("alignment"),
        Some(&Any::from("center"))
    );
    assert_eq!(
        paragraphs[0].properties.get("borders"),
        Some(&Any::from("boxed"))
    );
    // Mid-split second half keeps its pPr — except borders, ALWAYS cleared.
    assert_eq!(
        paragraphs[1].properties.get("alignment"),
        Some(&Any::from("center"))
    );
    assert_eq!(paragraphs[1].properties.get("borders"), None);
}

#[test]
fn split_at_end_inherits_subset_and_reduces_dtf_to_font_size_color() {
    let (doc, original) = doc_with("heading");
    let dtf: BTreeMap<String, Any> = [
        ("fontFamily".to_string(), Any::from("Georgia")),
        ("fontSize".to_string(), Any::Number(48.0)),
        ("color".to_string(), Any::from("336699")),
        ("bold".to_string(), Any::Bool(true)),
    ]
    .into_iter()
    .collect();
    doc.set_paragraph_default_format(&ctx(), &ParaSelector::One(original.clone()), Some(&dtf))
        .unwrap();
    doc.set_paragraph_attrs(
        &ctx(),
        &ParaSelector::One(original.clone()),
        &ParaAttrDelta {
            alignment: Patch::Set("center".into()),
            space_after: Patch::Set(240.0),
            ..ParaAttrDelta::default()
        },
    )
    .unwrap();
    // Split at the end of the paragraph text → empty second half.
    doc.split_paragraph(&ctx(), Position::new("body", 7), None)
        .unwrap();
    let paragraphs = doc.paragraphs("body").unwrap();
    assert_eq!(paragraphs[1].text, "");
    let second = &paragraphs[1].properties;
    // Inherited subset kept: pStyle, spaceAfter; alignment (not in the subset) dropped.
    assert_eq!(second.get("pStyle"), Some(&Any::from("Normal")));
    assert_eq!(second.get("spaceAfter"), Some(&Any::Number(240.0)));
    assert_eq!(second.get("alignment"), None);
    // dtf reduced to the font/size/color carry — bold dropped.
    let dtf = second.get("defaultTextFormatting").expect("dtf kept");
    assert_eq!(map_get(dtf, "fontFamily"), Some(&Any::from("Georgia")));
    assert_eq!(map_get(dtf, "fontSize"), Some(&Any::Number(48.0)));
    assert_eq!(map_get(dtf, "color"), Some(&Any::from("336699")));
    assert_eq!(map_get(dtf, "bold"), None);
    // First half untouched.
    assert_eq!(
        paragraphs[0].properties.get("alignment"),
        Some(&Any::from("center"))
    );
}

#[test]
fn split_at_end_with_next_style_switches_the_second_half() {
    let (doc, _) = doc_with("Heading text");
    let next = ResolvedStyleProjection {
        style_id: "BodyText".into(),
        known: true,
        paragraph_attrs: [("spaceBefore".to_string(), Any::Number(120.0))]
            .into_iter()
            .collect(),
        run_marks: BTreeMap::new(),
    };
    doc.split_paragraph(&ctx(), Position::new("body", 12), Some(&next))
        .unwrap();
    let paragraphs = doc.paragraphs("body").unwrap();
    let second = &paragraphs[1].properties;
    assert_eq!(second.get("pStyle"), Some(&Any::from("BodyText")));
    assert_eq!(second.get("spaceBefore"), Some(&Any::Number(120.0)));
    assert_eq!(second.get("alignment"), None); // reset by the projection
    assert_eq!(second.get("borders"), None);
    // First half keeps the source style.
    assert_eq!(
        paragraphs[0].properties.get("pStyle"),
        Some(&Any::from("Normal"))
    );

    // An unknown next style must fail BEFORE mutating.
    let unknown = ResolvedStyleProjection {
        style_id: "Ghost".into(),
        known: false,
        ..ResolvedStyleProjection::default()
    };
    assert_eq!(
        doc.split_paragraph(&ctx(), Position::new("body", 3), Some(&unknown)),
        Err(OpError::UnknownStyle("Ghost".into()))
    );
    assert_eq!(doc.paragraphs("body").unwrap().len(), 2);
}

#[test]
fn suggesting_split_stamps_ppr_ins_on_the_first_half() {
    let (doc, original) = doc_with("onetwo");
    let split = doc
        .split_paragraph(&sug("Alice"), Position::new("body", 3), None)
        .unwrap();
    assert_eq!(split.revision_ids.len(), 1);
    let paragraphs = doc.paragraphs("body").unwrap();
    assert_eq!(paragraphs[0].para_id, original);
    assert!(active(&paragraphs[0].properties, "pPrIns"));
    assert!(!active(&paragraphs[1].properties, "pPrIns"));
    // The pilcrow embed itself carries the ins stamp.
    let segments = doc.story_segments("body").unwrap();
    let pilcrow_attrs = segments
        .iter()
        .find_map(|segment| match &segment.content {
            SegmentContent::Pilcrow(props) if props.para_id == original => {
                Some(segment.attributes.clone())
            }
            _ => None,
        })
        .unwrap();
    assert_eq!(
        revision_author_of(&pilcrow_attrs, "ins").as_deref(),
        Some("Alice")
    );
}

#[test]
fn merge_directions_adopt_the_earlier_paragraph_and_guard_edges() {
    let (doc, first) = doc_with("onetwo");
    let split = doc
        .split_paragraph(&ctx(), Position::new("body", 3), None)
        .unwrap();
    // Edges error.
    assert_eq!(
        doc.merge_paragraphs(&ctx(), &split.second_para_id, MergeDirection::Forward),
        Err(OpError::CannotMergeFinalParagraph(
            split.second_para_id.clone()
        ))
    );
    assert_eq!(
        doc.merge_paragraphs(&ctx(), &first, MergeDirection::Backward),
        Err(OpError::NoParagraphBefore(first.clone()))
    );
    // Backward on the second paragraph merges at the same boundary as Forward on the first.
    doc.set_paragraph_attrs(
        &ctx(),
        &ParaSelector::One(first.clone()),
        &ParaAttrDelta {
            alignment: Patch::Set("right".into()),
            ..ParaAttrDelta::default()
        },
    )
    .unwrap();
    doc.merge_paragraphs(&ctx(), &split.second_para_id, MergeDirection::Backward)
        .unwrap();
    let paragraphs = doc.paragraphs("body").unwrap();
    assert_eq!(paragraphs.len(), 1);
    assert_eq!(paragraphs[0].text, "onetwo");
    assert_eq!(paragraphs[0].para_id, first);
    assert_eq!(
        paragraphs[0].properties.get("alignment"),
        Some(&Any::from("right"))
    );
}

#[test]
fn suggesting_merge_retains_the_mark_with_ppr_del() {
    let (doc, first) = doc_with("onetwo");
    doc.split_paragraph(&ctx(), Position::new("body", 3), None)
        .unwrap();
    let receipt = doc
        .merge_paragraphs(&sug("Bob"), &first, MergeDirection::Forward)
        .unwrap();
    assert_eq!(receipt.revision_ids.len(), 1);
    let paragraphs = doc.paragraphs("body").unwrap();
    assert_eq!(paragraphs.len(), 2, "mark retained until accepted");
    assert!(active(&paragraphs[0].properties, "pPrDel"));
}

// ---------------------------------------------------------------------------
// Inline formatting
// ---------------------------------------------------------------------------

#[test]
fn toggle_format_sets_partial_ranges_and_clears_full_ranges() {
    let (doc, _) = doc_with("abcd");
    // Partially bold → toggle bolds the whole range.
    doc.format_range(
        &ctx(),
        StoryRange::new("body", 0, 2),
        &InlineFormatDelta {
            bold: Patch::Set(true),
            ..InlineFormatDelta::default()
        },
    )
    .unwrap();
    doc.toggle_format(&ctx(), StoryRange::new("body", 0, 4), SimpleFormat::Bold)
        .unwrap();
    assert_eq!(seg_attrs(&doc, "cd").get("bold"), Some(&Any::Bool(true)));
    // Fully bold → toggle removes it everywhere.
    doc.toggle_format(&ctx(), StoryRange::new("body", 0, 4), SimpleFormat::Bold)
        .unwrap();
    assert!(!active(&seg_attrs(&doc, "ab"), "bold"));
    // Underline toggles on with the default single style.
    doc.toggle_format(
        &ctx(),
        StoryRange::new("body", 0, 2),
        SimpleFormat::Underline,
    )
    .unwrap();
    let underline = seg_attrs(&doc, "ab");
    assert_eq!(
        map_get(underline.get("underline").unwrap(), "style"),
        Some(&Any::from("single"))
    );
    // Superscript and subscript are mutually exclusive.
    doc.toggle_format(
        &ctx(),
        StoryRange::new("body", 0, 2),
        SimpleFormat::Subscript,
    )
    .unwrap();
    doc.toggle_format(
        &ctx(),
        StoryRange::new("body", 0, 2),
        SimpleFormat::Superscript,
    )
    .unwrap();
    let attrs = seg_attrs(&doc, "ab");
    assert_eq!(attrs.get("superscript"), Some(&Any::Bool(true)));
    assert!(!active(&attrs, "subscript"));
}

#[test]
fn format_range_lowers_the_word_value_shapes() {
    let (doc, _) = doc_with("styled text");
    doc.format_range(
        &ctx(),
        StoryRange::new("body", 0, 6),
        &InlineFormatDelta {
            font_size: Patch::Set(12.0),
            font_family: Patch::Set(FontFamilyPatch {
                ascii: "Arial".into(),
                h_ansi: None,
            }),
            color: Patch::Set(ColorPatch::Rgb("FF0000".into())),
            highlight: Patch::Set("#FFFF00".into()),
            ..InlineFormatDelta::default()
        },
    )
    .unwrap();
    let attrs = seg_attrs(&doc, "styled");
    // fontSize: points → half-points, written to BOTH size and sizeCs (w:sz + w:szCs).
    let font_size = attrs.get("fontSize").unwrap();
    assert_eq!(map_get(font_size, "size"), Some(&Any::Number(24.0)));
    assert_eq!(map_get(font_size, "sizeCs"), Some(&Any::Number(24.0)));
    // fontFamily: hAnsi defaults to ascii (w:ascii + w:hAnsi).
    let family = attrs.get("fontFamily").unwrap();
    assert_eq!(map_get(family, "ascii"), Some(&Any::from("Arial")));
    assert_eq!(map_get(family, "hAnsi"), Some(&Any::from("Arial")));
    // color: rgb XOR theme.
    let color = attrs.get("textColor").unwrap();
    assert_eq!(map_get(color, "rgb"), Some(&Any::from("FF0000")));
    assert_eq!(map_get(color, "themeColor"), Some(&Any::Null));
    // highlight: hex mapped onto Word's named palette.
    assert_eq!(
        map_get(attrs.get("highlight").unwrap(), "color"),
        Some(&Any::from("yellow"))
    );

    // Theme color replaces rgb (XOR).
    doc.format_range(
        &ctx(),
        StoryRange::new("body", 0, 6),
        &InlineFormatDelta {
            color: Patch::Set(ColorPatch::Theme("accent1".into())),
            ..InlineFormatDelta::default()
        },
    )
    .unwrap();
    let color = seg_attrs(&doc, "styled");
    let color = color.get("textColor").unwrap();
    assert_eq!(map_get(color, "rgb"), Some(&Any::Null));
    assert_eq!(map_get(color, "themeColor"), Some(&Any::from("accent1")));

    // Tri-state: Keep leaves fields alone, Clear removes, Set(false) clears booleans.
    doc.format_range(
        &ctx(),
        StoryRange::new("body", 0, 6),
        &InlineFormatDelta {
            highlight: Patch::Clear,
            ..InlineFormatDelta::default()
        },
    )
    .unwrap();
    let attrs = seg_attrs(&doc, "styled");
    assert!(!active(&attrs, "highlight"));
    assert!(active(&attrs, "fontSize"), "Keep left fontSize untouched");

    // Unmapped highlight hex values pass through unchanged.
    assert_eq!(highlight_color_name("#123456"), "#123456");
    assert_eq!(highlight_color_name("00ffff"), "cyan");
}

#[test]
fn clear_formatting_keeps_hyperlinks_and_tracked_changes() {
    let (doc, _) = doc_with("read the docs now");
    let link_attrs: BTreeMap<String, Any> = [
        ("hyperlink".to_string(), Any::from("https://example.com")),
        ("bold".to_string(), Any::Bool(true)),
    ]
    .into_iter()
    .collect();
    doc.insert_text(
        &ctx(),
        Position::new("body", 9),
        "LINK",
        FormatPolicy::Explicit(link_attrs),
    )
    .unwrap();
    doc.insert_text(
        &sug("Alice"),
        Position::new("body", 0),
        "NEW ",
        FormatPolicy::Plain,
    )
    .unwrap();
    let len = doc.story_len("body").unwrap();
    doc.clear_formatting(&ctx(), StoryRange::new("body", 0, len - 1))
        .unwrap();
    let link = seg_attrs(&doc, "LINK");
    assert!(!active(&link, "bold"), "formatting stripped");
    assert_eq!(
        link.get("hyperlink"),
        Some(&Any::from("https://example.com")),
        "hyperlink retained (Word Ctrl+Space)"
    );
    let pending = seg_attrs(&doc, "NEW");
    assert_eq!(
        revision_author_of(&pending, "ins").as_deref(),
        Some("Alice"),
        "tracked-change stamps retained"
    );
}

// ---------------------------------------------------------------------------
// Paragraph attrs + style
// ---------------------------------------------------------------------------

#[test]
fn indent_wrappers_default_to_720_and_clamp_to_zero() {
    let (doc, para) = doc_with("indented");
    let selector = ParaSelector::One(para.clone());
    doc.increase_indent(&ctx(), &selector, None).unwrap();
    doc.increase_indent(&ctx(), &selector, None).unwrap();
    assert_eq!(
        doc.paragraphs("body").unwrap()[0]
            .properties
            .get("indentLeft"),
        Some(&Any::Number(1440.0))
    );
    doc.decrease_indent(&ctx(), &selector, None).unwrap();
    assert_eq!(
        doc.paragraphs("body").unwrap()[0]
            .properties
            .get("indentLeft"),
        Some(&Any::Number(720.0))
    );
    // Reaching zero clears the attribute.
    doc.decrease_indent(&ctx(), &selector, None).unwrap();
    assert_eq!(
        doc.paragraphs("body").unwrap()[0]
            .properties
            .get("indentLeft"),
        None
    );
    doc.decrease_indent(&ctx(), &selector, None).unwrap();
    assert_eq!(
        doc.paragraphs("body").unwrap()[0]
            .properties
            .get("indentLeft"),
        None
    );
}

#[test]
fn tab_stops_add_replace_and_remove() {
    let (doc, para) = doc_with("tabbed");
    let selector = ParaSelector::One(para.clone());
    doc.add_tab_stop(
        &ctx(),
        &selector,
        &TabStop {
            pos: 1440.0,
            alignment: "center".into(),
            leader: None,
        },
    )
    .unwrap();
    doc.add_tab_stop(
        &ctx(),
        &selector,
        &TabStop {
            pos: 720.0,
            alignment: "left".into(),
            leader: Some("dot".into()),
        },
    )
    .unwrap();
    let tabs = doc.paragraphs("body").unwrap()[0]
        .properties
        .get("tabs")
        .cloned()
        .unwrap();
    let Any::Array(stops) = &tabs else {
        panic!("tabs must be an array")
    };
    assert_eq!(stops.len(), 2);
    assert_eq!(map_get(&stops[0], "pos"), Some(&Any::Number(720.0))); // sorted by pos
    assert_eq!(map_get(&stops[1], "pos"), Some(&Any::Number(1440.0)));
    doc.remove_tab_stop(&ctx(), &selector, 720.0).unwrap();
    doc.remove_tab_stop(&ctx(), &selector, 1440.0).unwrap();
    assert_eq!(
        doc.paragraphs("body").unwrap()[0].properties.get("tabs"),
        None,
        "empty tab list clears the attr"
    );
}

#[test]
fn apply_paragraph_style_resets_attrs_sweeps_marks_and_errs_before_mutating() {
    let (doc, para) = doc_with("styled paragraph");
    let selector = ParaSelector::One(para.clone());
    // Direct formatting that the style must reset/sweep.
    doc.set_paragraph_attrs(
        &ctx(),
        &selector,
        &ParaAttrDelta {
            alignment: Patch::Set("center".into()),
            ..ParaAttrDelta::default()
        },
    )
    .unwrap();
    doc.format_range(
        &ctx(),
        StoryRange::new("body", 0, 6),
        &InlineFormatDelta {
            italic: Patch::Set(true),
            ..InlineFormatDelta::default()
        },
    )
    .unwrap();

    // Unknown style: error BEFORE any mutation.
    let before = doc.story_segments("body").unwrap();
    let unknown = ResolvedStyleProjection {
        style_id: "Nope".into(),
        known: false,
        ..ResolvedStyleProjection::default()
    };
    assert_eq!(
        doc.apply_paragraph_style(&ctx(), &selector, &unknown),
        Err(OpError::UnknownStyle("Nope".into()))
    );
    assert_eq!(doc.story_segments("body").unwrap(), before);

    // Known style: styleId set, style-controlled attrs reset, 7 marks swept, run formats added.
    let heading = ResolvedStyleProjection {
        style_id: "Heading1".into(),
        known: true,
        paragraph_attrs: [("spaceBefore".to_string(), Any::Number(240.0))]
            .into_iter()
            .collect(),
        run_marks: [
            ("bold".to_string(), Any::Bool(true)),
            (
                "fontSize".to_string(),
                Any::Map(Arc::new(
                    [
                        ("size".to_string(), Any::Number(32.0)),
                        ("sizeCs".to_string(), Any::Number(32.0)),
                    ]
                    .into_iter()
                    .collect(),
                )),
            ),
        ]
        .into_iter()
        .collect(),
    };
    doc.apply_paragraph_style(&ctx(), &selector, &heading)
        .unwrap();
    let paragraph = &doc.paragraphs("body").unwrap()[0];
    assert_eq!(
        paragraph.properties.get("pStyle"),
        Some(&Any::from("Heading1"))
    );
    assert_eq!(
        paragraph.properties.get("spaceBefore"),
        Some(&Any::Number(240.0))
    );
    assert_eq!(
        paragraph.properties.get("alignment"),
        None,
        "style-controlled attr reset"
    );
    let attrs = seg_attrs(&doc, "styled");
    assert!(!active(&attrs, "italic"), "old style-controlled mark swept");
    assert_eq!(attrs.get("bold"), Some(&Any::Bool(true)));
    assert_eq!(
        map_get(attrs.get("fontSize").unwrap(), "size"),
        Some(&Any::Number(32.0))
    );
}

#[test]
fn apply_paragraph_style_clears_a_widow_control_off_the_new_style_does_not_author() {
    let (doc, para) = doc_with("styled paragraph");
    let selector = ParaSelector::One(para.clone());
    doc.set_paragraph_attr(&para, "widowControl", Any::Bool(false))
        .unwrap();

    // Absence encodes default-on, so a style that authors nothing must clear
    // false.
    let plain = ResolvedStyleProjection {
        style_id: "Body".into(),
        known: true,
        ..ResolvedStyleProjection::default()
    };
    doc.apply_paragraph_style(&ctx(), &selector, &plain)
        .unwrap();
    assert_eq!(
        doc.paragraphs("body").unwrap()[0]
            .properties
            .get("widowControl"),
        None
    );

    let off = ResolvedStyleProjection {
        style_id: "Tight".into(),
        known: true,
        paragraph_attrs: [("widowControl".to_string(), Any::Bool(false))]
            .into_iter()
            .collect(),
        ..ResolvedStyleProjection::default()
    };
    doc.apply_paragraph_style(&ctx(), &selector, &off).unwrap();
    assert_eq!(
        doc.paragraphs("body").unwrap()[0]
            .properties
            .get("widowControl"),
        Some(&Any::Bool(false))
    );
}

#[test]
fn dedupe_para_ids_first_occurrence_keeps_its_id() {
    // Concurrent splits of the same paragraph give both new pilcrows the ORIGINAL paraId.
    let base = EditingDoc::new(1);
    let original = base
        .create_story("body", "abcdef", "Normal", "left")
        .unwrap();
    let update = base.encode_state_as_update_v1();
    let a = EditingDoc::new(2);
    let b = EditingDoc::new(3);
    a.apply_update_v1(&update).unwrap();
    b.apply_update_v1(&update).unwrap();
    a.split_paragraph(&ctx(), Position::new("body", 2), None)
        .unwrap();
    b.split_paragraph(&ctx(), Position::new("body", 4), None)
        .unwrap();
    let from_a = a.encode_state_as_update_v1();
    let from_b = b.encode_state_as_update_v1();
    a.apply_update_v1(&from_b).unwrap();
    b.apply_update_v1(&from_a).unwrap();
    let ids: Vec<ParagraphId> = a
        .paragraphs("body")
        .unwrap()
        .into_iter()
        .map(|p| p.para_id)
        .collect();
    assert_eq!(
        ids.iter().filter(|id| **id == original).count(),
        2,
        "concurrent splits duplicate the original id"
    );
    let renames = a.dedupe_para_ids(DATE).unwrap();
    assert_eq!(renames.len(), 1);
    assert_eq!(renames[0].0, original);
    let ids: Vec<ParagraphId> = a
        .paragraphs("body")
        .unwrap()
        .into_iter()
        .map(|p| p.para_id)
        .collect();
    assert_eq!(ids.iter().filter(|id| **id == original).count(), 1);
    assert_eq!(ids[0], original, "FIRST occurrence keeps the id");
}

// ---------------------------------------------------------------------------
// Undo
// ---------------------------------------------------------------------------

#[test]
fn undo_tracks_local_only_with_barriers() {
    let (doc, _) = doc_with("base");
    let mut undo = doc.undo_manager();
    assert!(!undo.can_undo());

    // Agent and system edits are untracked.
    let agent = EditCtx {
        author: "Agent".into(),
        origin: EditOrigin::Agent,
        suggesting: None,
        now_iso: DATE.into(),
    };
    doc.insert_text(&agent, Position::new("body", 0), "A", FormatPolicy::Plain)
        .unwrap();
    doc.insert_text(
        &EditCtx::system(DATE),
        Position::new("body", 0),
        "S",
        FormatPolicy::Plain,
    )
    .unwrap();
    assert!(
        !undo.can_undo(),
        "agent/system origins never enter local undo"
    );

    // Local edits group within the capture window; a barrier splits them.
    doc.insert_text(&ctx(), Position::new("body", 2), "x", FormatPolicy::Plain)
        .unwrap();
    undo.add_undo_barrier();
    doc.insert_text(&ctx(), Position::new("body", 3), "y", FormatPolicy::Plain)
        .unwrap();
    assert!(undo.can_undo());
    assert_eq!(undo.undo_depth(), 2);
    assert!(undo.undo());
    assert_eq!(
        raw_text(&doc),
        "SAxbase",
        "only the post-barrier edit reverted"
    );
    assert!(undo.undo());
    assert_eq!(raw_text(&doc), "SAbase");
    assert!(undo.can_redo());
    assert!(undo.redo());
    assert_eq!(raw_text(&doc), "SAxbase");
}

// ---------------------------------------------------------------------------
// Read queries
// ---------------------------------------------------------------------------

#[test]
fn vanilla_view_and_resolve_search_map_back_to_raw_offsets() {
    let (doc, para) = doc_with("abcdef");
    doc.insert_text(
        &sug("Alice"),
        Position::new("body", 2),
        "XX",
        FormatPolicy::Plain,
    )
    .unwrap();
    assert_eq!(doc.para_text(&para, TextView::Raw).unwrap(), "abXXcdef");
    assert_eq!(doc.para_text(&para, TextView::Vanilla).unwrap(), "abcdef");
    // Search in the vanilla view returns RAW paragraph offsets.
    let range = doc
        .resolve_search("body", None, "cd", TextView::Vanilla)
        .unwrap();
    assert_eq!(range.start.para, para);
    assert_eq!((range.start.offset, range.end.offset), (4, 6));
    // The mapped range drives a real op.
    let story_range = doc.locate_range(&range).unwrap();
    doc.delete_range(&ctx(), story_range).unwrap();
    assert_eq!(doc.para_text(&para, TextView::Raw).unwrap(), "abXXef");

    // Ambiguity and not-found are typed.
    assert_eq!(
        doc.resolve_search("body", None, "zz", TextView::Vanilla),
        Err(OpError::SearchNotFound("zz".into()))
    );
    let (doc2, _) = doc_with("ab ab");
    assert_eq!(
        doc2.resolve_search("body", None, "ab", TextView::Vanilla),
        Err(OpError::AmbiguousSearch {
            needle: "ab".into(),
            occurrences: 2
        })
    );
}

#[test]
fn text_between_uses_the_requested_view() {
    let (doc, para) = doc_with("hello world");
    doc.insert_text(
        &sug("Alice"),
        Position::new("body", 5),
        " INS",
        FormatPolicy::Plain,
    )
    .unwrap();
    // Raw: "hello INS world"
    let range = LocRange {
        start: Loc::new("body", para.clone(), 0),
        end: Loc::new("body", para.clone(), 15),
    };
    assert_eq!(
        doc.text_between(&range, TextView::Raw).unwrap(),
        "hello INS world"
    );
    assert_eq!(
        doc.text_between(&range, TextView::Vanilla).unwrap(),
        "hello world"
    );
}

#[test]
fn find_in_document_windows_and_skips_ambiguous_paragraphs() {
    let (doc, _) = doc_with("prefix prefix target here");
    // Second paragraph contains the needle twice → skipped entirely.
    doc.split_paragraph(&ctx(), Position::new("body", 25), None)
        .unwrap();
    let end = doc.story_len("body").unwrap() - 1;
    doc.insert_text(
        &ctx(),
        Position::new("body", end),
        "target and target again",
        FormatPolicy::Plain,
    )
    .unwrap();
    let matches = doc
        .find_in_document("body", "TARGET", FindOptions::default())
        .unwrap();
    assert_eq!(matches.len(), 1, "ambiguous paragraph skipped");
    assert_eq!(matches[0].match_text, "target", "original casing returned");
    assert_eq!(matches[0].before, "prefix prefix ");
    assert_eq!(matches[0].after, " here");
    // Case-sensitive narrows to nothing.
    let matches = doc
        .find_in_document(
            "body",
            "TARGET",
            FindOptions {
                case_sensitive: true,
                ..FindOptions::default()
            },
        )
        .unwrap();
    assert!(matches.is_empty());
}

#[test]
fn selection_info_reports_vanilla_slices_of_the_start_paragraph() {
    let (doc, para) = doc_with("hello brave world");
    let info = doc
        .selection_info(
            &Loc::new("body", para.clone(), 6),
            &Loc::new("body", para.clone(), 11),
        )
        .unwrap();
    assert_eq!(info.para_id, para);
    assert_eq!(info.before, "hello ");
    assert_eq!(info.selected_text, "brave");
    assert_eq!(info.after, " world");
    assert_eq!(info.paragraph_text, "hello brave world");
}

#[test]
fn list_changes_and_find_change_range_cover_text_and_marks() {
    let (doc, para) = doc_with("alpha beta");
    let ins_receipt = doc
        .insert_text(
            &sug("Alice"),
            Position::new("body", 5),
            " NEW",
            FormatPolicy::Plain,
        )
        .unwrap();
    let del_receipt = doc
        .delete_range(&sug("Bob"), StoryRange::new("body", 0, 2))
        .unwrap();
    let split = doc
        .split_paragraph(&sug("Carol"), Position::new("body", 9), None)
        .unwrap();
    let changes = doc.list_changes("body").unwrap();
    let kinds: Vec<ChangeKind> = changes.iter().map(|change| change.kind).collect();
    assert!(kinds.contains(&ChangeKind::Insertion));
    assert!(kinds.contains(&ChangeKind::Deletion));
    assert!(kinds.contains(&ChangeKind::ParagraphMarkInsertion));
    let ins_id = &ins_receipt.revision_ids[0];
    let ins_change = changes
        .iter()
        .find(|change| change.revision_id == *ins_id)
        .unwrap();
    assert_eq!(ins_change.author, "Alice");
    assert_eq!(ins_change.range.start.para, para);
    // find_change_range agrees with list_changes.
    let range = doc.find_change_range(ins_id).unwrap();
    assert_eq!(range, ins_change.range);
    assert!(doc.find_change_range(&del_receipt.revision_ids[0]).is_ok());
    assert!(doc.find_change_range(&split.revision_ids[0]).is_ok());
    assert_eq!(
        doc.find_change_range("999:999"),
        Err(OpError::UnknownChange("999:999".into()))
    );
}

#[test]
fn comments_list_and_range_resolution() {
    let (doc, para) = doc_with("alpha beta gamma");
    let comment_id = doc
        .add_comment(
            &[StoryRange::new("body", 6, 10)],
            "Reviewer",
            DATE,
            Any::from("nice word"),
        )
        .unwrap();
    let range = doc.find_comment_range(&comment_id).unwrap();
    assert_eq!(range.start.para, para);
    assert_eq!((range.start.offset, range.end.offset), (6, 10));
    let comments = doc.list_comments().unwrap();
    assert_eq!(comments.len(), 1);
    assert_eq!(comments[0].id, comment_id);
    assert_eq!(comments[0].author, "Reviewer");
    assert!(!comments[0].done);
    assert_eq!(comments[0].ranges, vec![range]);
    assert_eq!(
        doc.find_comment_range("missing"),
        Err(OpError::UnknownComment("missing".into()))
    );
}

#[test]
fn nav_boundary_walks_graphemes_words_and_paragraph_edges() {
    let (doc, para) = doc_with("a\u{1F44D}\u{1F3FD}b");
    // Grapheme: the emoji + modifier is ONE cluster of 4 UTF-16 units.
    let next = doc
        .nav_boundary(
            &Loc::new("body", para.clone(), 1),
            NavUnit::Grapheme,
            NavDirection::Next,
        )
        .unwrap();
    assert_eq!(next.offset, 5);
    let prev = doc
        .nav_boundary(
            &Loc::new("body", para.clone(), 5),
            NavUnit::Grapheme,
            NavDirection::Prev,
        )
        .unwrap();
    assert_eq!(prev.offset, 1);

    let (doc2, para2) = doc_with("foo bar");
    let split = doc2
        .split_paragraph(&ctx(), Position::new("body", 7), None)
        .unwrap();
    assert_eq!(split.first_para_id, para2);
    let next = doc2
        .nav_boundary(
            &Loc::new("body", para2.clone(), 0),
            NavUnit::Word,
            NavDirection::Next,
        )
        .unwrap();
    assert_eq!(next.offset, 3);
    // At the paragraph end, Next crosses the pilcrow into the following paragraph.
    let crossed = doc2
        .nav_boundary(
            &Loc::new("body", para2.clone(), 7),
            NavUnit::Word,
            NavDirection::Next,
        )
        .unwrap();
    assert_eq!(crossed.para, split.second_para_id);
    assert_eq!(crossed.offset, 0);
    // Prev from the second paragraph's start lands on the first paragraph's end.
    let back = doc2
        .nav_boundary(
            &Loc::new("body", split.second_para_id.clone(), 0),
            NavUnit::Word,
            NavDirection::Prev,
        )
        .unwrap();
    assert_eq!((back.para, back.offset), (para2, 7));
}

struct FakeBridge {
    pages: Vec<Vec<ParagraphId>>,
}

impl LayoutBridge for FakeBridge {
    fn page_count(&self) -> u32 {
        self.pages.len() as u32
    }

    fn paragraphs_on_page(&self, page_number: u32) -> Vec<ParagraphId> {
        self.pages[(page_number - 1) as usize].clone()
    }
}

#[test]
fn page_content_dedupes_split_paragraph_fragments() {
    let (doc, first) = doc_with("onetwo");
    let split = doc
        .split_paragraph(&ctx(), Position::new("body", 3), None)
        .unwrap();
    // The first paragraph is split across the page boundary → its id repeats.
    let bridge = FakeBridge {
        pages: vec![vec![
            first.clone(),
            first.clone(),
            split.second_para_id.clone(),
        ]],
    };
    let page = doc
        .page_content(1, &bridge, TextView::Vanilla)
        .unwrap()
        .unwrap();
    assert_eq!(page.paragraphs.len(), 2, "deduped by paraId");
    assert_eq!(page.paragraphs[0].para_id, first);
    assert_eq!(page.paragraphs[0].text, "one");
    assert_eq!(page.paragraphs[0].style_id.as_deref(), Some("Normal"));
    assert_eq!(
        page.text,
        format!("[{first}] one\n[{}] two", split.second_para_id)
    );
    assert_eq!(
        doc.page_content(2, &bridge, TextView::Vanilla).unwrap(),
        None
    );
    assert_eq!(
        doc.page_content(0, &bridge, TextView::Vanilla).unwrap(),
        None
    );
}

// Slots that open with a table or page break

/// Units as text: `¶` a pilcrow, `[kind]` any other embed.
fn slot_units(doc: &EditingDoc) -> String {
    doc.story_segments("body")
        .unwrap()
        .into_iter()
        .map(|segment| match segment.content {
            SegmentContent::Text(text) => text,
            SegmentContent::Pilcrow(_) => "¶".to_owned(),
            SegmentContent::OtherEmbed { kind, .. } => format!("[{kind}]"),
        })
        .collect()
}

/// `Before¶[pageBreak]Chapter¶`: a page break opens the second slot.
fn page_break_slot() -> (EditingDoc, ParagraphId, ParagraphId) {
    let (doc, _) = doc_with("BeforeChapter");
    let split = doc
        .split_paragraph(&ctx(), Position::new("body", 6), None)
        .unwrap();
    doc.insert_embed(&ctx(), Position::new("body", 7), "pageBreak", vec![])
        .unwrap();
    assert_eq!(slot_units(&doc), "Before¶[pageBreak]Chapter¶");
    (doc, split.first_para_id, split.second_para_id)
}

/// Merges `para` Delete-style (`forward`) or Backspace-style.
fn merge(doc: &EditingDoc, context: &EditCtx, para: &ParagraphId, forward: bool) -> Receipt {
    let direction = if forward {
        MergeDirection::Forward
    } else {
        MergeDirection::Backward
    };
    doc.merge_paragraphs(context, para, direction).unwrap()
}

fn para_ids(doc: &EditingDoc) -> Vec<ParagraphId> {
    doc.paragraphs("body")
        .unwrap()
        .into_iter()
        .map(|paragraph| paragraph.para_id)
        .collect()
}

#[test]
fn a_merge_into_a_slot_that_opens_with_a_break_removes_the_break() {
    for kind in ["pageBreak", "columnBreak"] {
        for forward in [true, false] {
            let (doc, first) = block_slot(kind);
            let second = para_ids(&doc)[1].clone();
            let receipt = merge(
                &doc,
                &ctx(),
                if forward { &first } else { &second },
                forward,
            );
            assert_eq!(slot_units(&doc), "Before¶After¶", "{kind}");
            assert_eq!(para_ids(&doc), [first.clone(), second.clone()]);
            // The caret stays where it was.
            let caret = if forward {
                Loc::new("body", first, 6)
            } else {
                Loc::new("body", second, 0)
            };
            assert_eq!(receipt.range.unwrap().start, caret);
        }
    }
    // Suggesting marks the break deleted.
    let (doc, _, second) = page_break_slot();
    let receipt = merge(&doc, &sug("Bob"), &second, false);
    assert_eq!(slot_units(&doc), "Before¶[pageBreak]Chapter¶");
    let marked = doc
        .story_segments("body")
        .unwrap()
        .into_iter()
        .find(|segment| matches!(segment.content, SegmentContent::OtherEmbed { .. }))
        .unwrap();
    assert_eq!(
        revision_id_of(&marked.attributes, "del"),
        receipt.revision_ids.first().cloned()
    );
}

#[test]
fn a_merge_into_a_slot_that_opens_with_a_table_removes_only_an_empty_paragraph() {
    for kind in ["table", "blockSdt"] {
        // A paragraph with content stays, and so does the caret.
        let (doc, first) = block_slot(kind);
        let second = para_ids(&doc)[1].clone();
        let before = doc.encode_state_as_update_v1();
        let forward = merge(&doc, &ctx(), &first, true);
        let backward = merge(&doc, &ctx(), &second, false);
        assert_eq!(doc.encode_state_as_update_v1(), before, "{kind}");
        assert_eq!(forward.range.unwrap().start, Loc::new("body", first, 6));
        assert_eq!(backward.range.unwrap().start, Loc::new("body", second, 0));
        // An empty one goes; the slot keeps its own paragraph.
        for forward in [true, false] {
            let (doc, _) = block_slot(kind);
            doc.split_paragraph(&ctx(), Position::new("body", 6), None)
                .unwrap();
            assert_eq!(slot_units(&doc), format!("Before¶¶[{kind}]After¶"));
            let ids = para_ids(&doc);
            let receipt = merge(&doc, &ctx(), &ids[if forward { 1 } else { 2 }], forward);
            assert_eq!(slot_units(&doc), format!("Before¶[{kind}]After¶"));
            assert_eq!(para_ids(&doc), [ids[0].clone(), ids[2].clone()]);
            assert_eq!(
                receipt.range.unwrap().start,
                Loc::new("body", ids[2].clone(), 0)
            );
        }
    }
    // Suggesting marks the empty paragraph's mark deleted.
    let (doc, _) = block_slot("table");
    doc.split_paragraph(&ctx(), Position::new("body", 6), None)
        .unwrap();
    let receipt = merge(&doc, &sug("Bob"), &para_ids(&doc)[1], true);
    assert_eq!(slot_units(&doc), "Before¶¶[table]After¶");
    let marks: Vec<_> = doc
        .story_segments("body")
        .unwrap()
        .into_iter()
        .filter(|segment| matches!(segment.content, SegmentContent::Pilcrow(_)))
        .map(|segment| revision_id_of(&segment.attributes, "del"))
        .collect();
    assert_eq!(marks, [None, receipt.revision_ids.first().cloned(), None]);
}

#[test]
fn a_merge_never_joins_two_tables() {
    // Before¶[table]¶[table]After¶: the paragraph between the tables belongs
    // to the first table's slot.
    let (doc, _) = block_slot("table");
    doc.split_paragraph(&ctx(), Position::new("body", 8), None)
        .unwrap();
    doc.insert_table(&ctx(), Position::new("body", 9), 1, 1)
        .unwrap();
    assert_eq!(slot_units(&doc), "Before¶[table]¶[table]After¶");
    let ids = para_ids(&doc);
    let before = doc.encode_state_as_update_v1();
    merge(&doc, &ctx(), &ids[1], true);
    merge(&doc, &ctx(), &ids[2], false);
    assert_eq!(doc.encode_state_as_update_v1(), before);
}

#[test]
fn inline_content_meant_ahead_of_a_slot_opening_block_lands_after_it() {
    let (doc, _, second) = page_break_slot();
    let receipt = doc
        .insert_text(
            &ctx(),
            Position::new("body", 7),
            "Typed ",
            FormatPolicy::Inherit,
        )
        .unwrap();
    assert_eq!(slot_units(&doc), "Before¶[pageBreak]Typed Chapter¶");
    assert_eq!(
        receipt.range.unwrap(),
        LocRange::new(
            Loc::new("body", second.clone(), 1),
            Loc::new("body", second, 7)
        )
    );
    doc.insert_tab(&ctx(), Position::new("body", 7)).unwrap();
    doc.insert_hard_break(&ctx(), Position::new("body", 7))
        .unwrap();
    doc.replace_range(&ctx(), StoryRange::new("body", 7, 7), "x")
        .unwrap();
    assert_eq!(
        slot_units(&doc),
        "Before¶[pageBreak]x[break]\tTyped Chapter¶"
    );
    // A block goes where it is asked to: a second break also opens the slot.
    doc.insert_embed(&ctx(), Position::new("body", 7), "pageBreak", vec![])
        .unwrap();
    assert_eq!(
        slot_units(&doc),
        "Before¶[pageBreak][pageBreak]x[break]\tTyped Chapter¶"
    );
}

#[test]
fn a_replacement_of_a_leading_block_lands_after_the_blocks_left() {
    let (doc, _, _) = page_break_slot();
    doc.insert_embed(&ctx(), Position::new("body", 7), "pageBreak", vec![])
        .unwrap();
    // Replacing the first of two leading breaks keeps the second ahead of the text.
    doc.replace_range(&ctx(), StoryRange::new("body", 7, 8), "Typed ")
        .unwrap();
    assert_eq!(slot_units(&doc), "Before¶[pageBreak]Typed Chapter¶");
}

/// `Before¶[kind]After¶`: a block of `kind` opens the second slot.
fn block_slot(kind: &str) -> (EditingDoc, ParagraphId) {
    let (doc, _) = doc_with("BeforeAfter");
    let split = doc
        .split_paragraph(&ctx(), Position::new("body", 6), None)
        .unwrap();
    if kind == "table" {
        doc.insert_table(&ctx(), Position::new("body", 7), 1, 1)
            .unwrap();
    } else {
        doc.insert_embed(&ctx(), Position::new("body", 7), kind, vec![])
            .unwrap();
    }
    assert_eq!(slot_units(&doc), format!("Before¶[{kind}]After¶"));
    (doc, split.first_para_id)
}

#[test]
fn a_range_delete_ending_at_a_block_led_slot_keeps_the_mark_before_it() {
    for kind in ["table", "pageBreak", "columnBreak", "blockSdt"] {
        let (doc, first) = block_slot(kind);
        doc.delete_range(&ctx(), StoryRange::new("body", 3, 7))
            .unwrap();
        assert_eq!(slot_units(&doc), format!("Bef¶[{kind}]After¶"), "{kind}");
        assert_eq!(doc.paragraphs("body").unwrap()[0].para_id, first);
    }
}

#[test]
fn a_range_delete_across_paragraphs_keeps_the_last_mark_with_the_first_identity() {
    let (doc, first) = block_slot("table");
    doc.split_paragraph(&ctx(), Position::new("body", 3), None)
        .unwrap();
    assert_eq!(slot_units(&doc), "Bef¶ore¶[table]After¶");
    doc.delete_range(&ctx(), StoryRange::new("body", 1, 8))
        .unwrap();
    assert_eq!(slot_units(&doc), "B¶[table]After¶");
    assert_eq!(doc.paragraphs("body").unwrap()[0].para_id, first);
    // A range that also takes one of the slot's leading blocks keeps the mark
    // and deletes that block.
    let (doc, _) = block_slot("pageBreak");
    doc.insert_embed(&ctx(), Position::new("body", 7), "table", vec![])
        .unwrap();
    doc.delete_range(&ctx(), StoryRange::new("body", 3, 8))
        .unwrap();
    assert_eq!(slot_units(&doc), "Bef¶[pageBreak]After¶");
}

#[test]
fn a_range_delete_from_an_empty_paragraph_or_one_block_still_deletes_it() {
    // Nothing is left ahead of the block, so the mark goes.
    let (doc, _) = block_slot("table");
    doc.delete_range(&ctx(), StoryRange::new("body", 0, 7))
        .unwrap();
    assert_eq!(slot_units(&doc), "[table]After¶");
    // Backspace after a block deletes that block.
    let (doc, _) = block_slot("pageBreak");
    doc.delete_range(&ctx(), StoryRange::new("body", 7, 8))
        .unwrap();
    assert_eq!(slot_units(&doc), "Before¶After¶");
}

#[test]
fn a_suggested_range_delete_or_replacement_keeps_the_mark_before_a_block() {
    let (doc, _) = block_slot("table");
    doc.delete_range(&sug("Bob"), StoryRange::new("body", 3, 7))
        .unwrap();
    assert_eq!(slot_units(&doc), "Before¶[table]After¶");
    assert!(active(&seg_attrs(&doc, "ore"), "del"));
    let pilcrow = doc
        .story_segments("body")
        .unwrap()
        .into_iter()
        .find(|segment| matches!(segment.content, SegmentContent::Pilcrow(_)))
        .unwrap();
    assert!(!active(&pilcrow.attributes, "del"));
    let (doc, _) = block_slot("table");
    doc.replace_range(&ctx(), StoryRange::new("body", 3, 7), "X")
        .unwrap();
    assert_eq!(slot_units(&doc), "BefX¶[table]After¶");
}

#[test]
fn a_replacement_from_a_paragraph_start_to_a_block_led_slot_keeps_the_mark() {
    for kind in ["table", "blockSdt"] {
        // Type-over and paste write their text in the paragraph before the block.
        let (doc, first) = block_slot(kind);
        doc.replace_range(&ctx(), StoryRange::new("body", 0, 7), "X")
            .unwrap();
        assert_eq!(slot_units(&doc), format!("X¶[{kind}]After¶"), "{kind}");
        assert_eq!(para_ids(&doc)[0], first);
        let (doc, _) = block_slot(kind);
        let run = RichRun {
            text: "X".into(),
            attrs: BTreeMap::new(),
        };
        doc.replace_range_rich(&ctx(), StoryRange::new("body", 0, 7), &[run])
            .unwrap();
        assert_eq!(slot_units(&doc), format!("X¶[{kind}]After¶"));
        let (doc, _) = block_slot(kind);
        let receipt = doc
            .replace_range(&sug("Bob"), StoryRange::new("body", 0, 7), "X")
            .unwrap();
        doc.accept_change(
            &ctx(),
            &ChangeTarget::Revision(receipt.revision_ids[0].clone()),
        )
        .unwrap();
        assert_eq!(slot_units(&doc), format!("X¶[{kind}]After¶"));
        // A delete from the paragraph's start takes the whole paragraph.
        let (doc, _) = block_slot(kind);
        doc.delete_range(&ctx(), StoryRange::new("body", 0, 7))
            .unwrap();
        assert_eq!(slot_units(&doc), format!("[{kind}]After¶"));
    }
}

#[test]
fn delete_or_backspace_beside_an_empty_paragraph_before_a_break_removes_the_paragraph() {
    for kind in ["pageBreak", "columnBreak"] {
        for forward in [true, false] {
            for (suggesting, context) in [(false, ctx()), (true, sug("Bob"))] {
                // Enter at the break's slot, then Delete or Backspace.
                let (doc, _) = block_slot(kind);
                let split = doc
                    .split_paragraph(&context, Position::new("body", 7), None)
                    .unwrap();
                assert_eq!(slot_units(&doc), format!("Before¶¶[{kind}]After¶"));
                let target = if forward {
                    &split.first_para_id
                } else {
                    &split.second_para_id
                };
                let receipt = merge(&doc, &context, target, forward);
                let case = format!("{kind} forward={forward} suggesting={suggesting}");
                assert_eq!(slot_units(&doc), format!("Before¶[{kind}]After¶"), "{case}");
                // Suggesting mode withdraws the author's own pending paragraph.
                assert!(receipt.revision_ids.is_empty(), "{case}");
                assert!(
                    doc.story_segments("body")
                        .unwrap()
                        .iter()
                        .all(|segment| revision_id_of(&segment.attributes, "del").is_none()),
                    "{case}"
                );
            }
        }
        // With text before the break, Delete still removes the break.
        let (doc, first) = block_slot(kind);
        merge(&doc, &ctx(), &first, true);
        assert_eq!(slot_units(&doc), "Before¶After¶", "{kind}");
    }
}

#[test]
fn an_empty_paragraph_ending_a_section_keeps_its_section_before_a_block() {
    for kind in ["pageBreak", "columnBreak", "table"] {
        for forward in [true, false] {
            let (doc, _) = block_slot(kind);
            let split = doc
                .split_paragraph(&ctx(), Position::new("body", 7), None)
                .unwrap();
            doc.set_paragraph_attr(
                &split.first_para_id,
                "sectionBreakType",
                Any::String("nextPage".into()),
            )
            .unwrap();
            let target = if forward {
                &split.first_para_id
            } else {
                &split.second_para_id
            };
            merge(&doc, &ctx(), target, forward);
            let case = format!("{kind} forward={forward}");
            // Before a break the break goes; before a table nothing changes.
            let expected = if kind == "table" {
                "Before¶¶[table]After¶".to_owned()
            } else {
                "Before¶¶After¶".to_owned()
            };
            assert_eq!(slot_units(&doc), expected, "{case}");
            assert!(
                doc.paragraphs("body").unwrap()[1]
                    .properties
                    .contains_key("sectionBreakType"),
                "{case}"
            );
        }
    }
}

#[test]
fn a_removed_empty_paragraph_before_a_block_hands_its_bookmarks_to_the_next() {
    let bookmark = |kind: &str| {
        Any::Map(Arc::new(
            [
                ("id".to_owned(), Any::from(5_i64)),
                ("kind".to_owned(), Any::String(kind.into())),
                ("offset".to_owned(), Any::from(0_i64)),
            ]
            .into(),
        ))
    };
    for kind in ["pageBreak", "columnBreak", "table"] {
        let (doc, _) = block_slot(kind);
        let split = doc
            .split_paragraph(&ctx(), Position::new("body", 7), None)
            .unwrap();
        doc.set_paragraph_attr(
            &split.first_para_id,
            "bookmarks",
            Any::Array(vec![bookmark("start"), bookmark("end")].into()),
        )
        .unwrap();
        merge(&doc, &ctx(), &split.first_para_id, true);
        let paragraphs = doc.paragraphs("body").unwrap();
        assert_eq!(paragraphs.len(), 2, "{kind}");
        let Some(Any::Array(moved)) = paragraphs[1].properties.get("bookmarks") else {
            panic!("{kind}: the bookmarks did not move");
        };
        assert_eq!(moved.len(), 2, "{kind}");
        // Ahead of the break the slot opens with; a table is no break.
        let Any::Map(first) = &moved[0] else { panic!() };
        assert_eq!(
            first.get("breaksAfter").is_some(),
            kind != "table",
            "{kind}"
        );
    }
}

#[test]
fn a_suggested_delete_in_an_original_empty_paragraph_before_a_break_marks_it_deleted() {
    let (doc, _, _) = page_break_slot();
    let split = doc
        .split_paragraph(&ctx(), Position::new("body", 7), None)
        .unwrap();
    let receipt = merge(&doc, &sug("Bob"), &split.first_para_id, true);
    assert_eq!(slot_units(&doc), "Before¶¶[pageBreak]Chapter¶");
    assert_eq!(receipt.revision_ids.len(), 1);
    let paragraph = &doc.paragraphs("body").unwrap()[1];
    assert!(paragraph.properties.contains_key("pPrDel"));
}

#[test]
fn accepting_a_deleted_mark_keeps_it_while_its_paragraph_holds_content_before_a_block() {
    for typed in [true, false] {
        let (doc, _) = block_slot("table");
        doc.split_paragraph(&ctx(), Position::new("body", 6), None)
            .unwrap();
        let receipt = merge(&doc, &sug("Bob"), &para_ids(&doc)[1], true);
        if typed {
            doc.insert_text(
                &ctx(),
                Position::new("body", 7),
                "typed",
                FormatPolicy::Inherit,
            )
            .unwrap();
        }
        doc.accept_change(
            &ctx(),
            &ChangeTarget::Revision(receipt.revision_ids[0].clone()),
        )
        .unwrap();
        let expected = if typed {
            "Before¶typed¶[table]After¶"
        } else {
            "Before¶[table]After¶"
        };
        assert_eq!(slot_units(&doc), expected);
        assert!(doc.list_changes("body").unwrap().is_empty());
    }
}

#[test]
fn enter_then_delete_before_a_block_leaves_its_paragraph_as_it_was() {
    for kind in ["table", "blockSdt"] {
        for context in [ctx(), sug("Bob")] {
            let (doc, _) = block_slot(kind);
            // The paragraph after the block has borders, which a split drops.
            doc.apply_raw_ops(
                "body",
                vec![RawOp::SetEmbedAttr {
                    index: 13,
                    key: "borders".into(),
                    value: Any::String("top".into()),
                }],
                &ctx(),
            )
            .unwrap();
            let before = doc.story_segments("body").unwrap();
            let split = doc
                .split_paragraph(&context, Position::new("body", 7), None)
                .unwrap();
            assert_eq!(slot_units(&doc), format!("Before¶¶[{kind}]After¶"));
            // Enter inserts an empty paragraph before the block.
            assert_eq!(para_ids(&doc)[2], split.second_para_id);
            assert_eq!(doc.story_segments("body").unwrap()[3..], before[2..]);
            merge(&doc, &context, &split.first_para_id, true);
            assert_eq!(doc.story_segments("body").unwrap(), before, "{kind}");
        }
    }
}

#[test]
fn a_paragraph_holding_only_comment_reference_fields_is_empty_before_a_table() {
    let (doc, _) = block_slot("table");
    doc.split_paragraph(&ctx(), Position::new("body", 6), None)
        .unwrap();
    let reference = vec![(
        "modelKind".to_owned(),
        Any::String("commentReference".into()),
    )];
    doc.insert_embed(&ctx(), Position::new("body", 7), "field", reference)
        .unwrap();
    assert_eq!(slot_units(&doc), "Before¶[field]¶[table]After¶");
    let slot = para_ids(&doc)[2].clone();
    let receipt = merge(&doc, &ctx(), &para_ids(&doc)[1], true);
    assert_eq!(slot_units(&doc), "Before¶[table]After¶");
    assert_eq!(receipt.range.unwrap().start, Loc::new("body", slot, 0));
}

#[test]
fn removing_a_comment_removes_its_reference_field() {
    let (doc, _) = doc_with("alpha beta");
    let comment = |id: &str, start: u32, end: u32| RawOp::SetComment {
        id: id.into(),
        ranges: vec![(start, end)],
        author: "Ada".into(),
        date: DATE.into(),
        body: Any::Null,
    };
    let reference = |id: Any| {
        vec![
            (
                "modelKind".to_owned(),
                Any::String("commentReference".into()),
            ),
            ("commentId".to_owned(), id),
        ]
    };
    doc.apply_raw_ops(
        "body",
        vec![comment("7", 0, 5), comment("8", 6, 10)],
        &ctx(),
    )
    .unwrap();
    // A seeded document names comments by string or numeric id.
    doc.insert_embed(
        &ctx(),
        Position::new("body", 10),
        "field",
        reference(Any::Number(8.0)),
    )
    .unwrap();
    doc.insert_embed(
        &ctx(),
        Position::new("body", 5),
        "field",
        reference(Any::String("7".into())),
    )
    .unwrap();
    assert_eq!(slot_units(&doc), "alpha[field] beta[field]¶");
    doc.apply_raw_ops(
        "body",
        vec![RawOp::RemoveComment { id: "7".into() }],
        &ctx(),
    )
    .unwrap();
    assert_eq!(slot_units(&doc), "alpha beta[field]¶");
    doc.apply_raw_ops(
        "body",
        vec![RawOp::RemoveComment { id: "8".into() }],
        &ctx(),
    )
    .unwrap();
    assert_eq!(slot_units(&doc), "alpha beta¶");
}

/// Each pilcrow's property values, in story order.
fn marks(doc: &EditingDoc) -> Vec<BTreeMap<String, Any>> {
    doc.story_segments("body")
        .unwrap()
        .into_iter()
        .filter_map(|segment| match segment.content {
            SegmentContent::Pilcrow(properties) => Some(properties.values),
            _ => None,
        })
        .collect()
}

#[test]
fn rejecting_an_inserted_mark_keeps_it_while_its_paragraph_holds_content_before_a_block() {
    for typist in [None, Some("Bob"), Some("Ann")] {
        let (doc, _) = block_slot("table");
        // Ann suggests Enter before the table; someone types into the new paragraph.
        let split = doc
            .split_paragraph(&sug("Ann"), Position::new("body", 7), None)
            .unwrap();
        let context = typist.map_or_else(ctx, sug);
        doc.insert_text(
            &context,
            Position::new("body", 7),
            "hello",
            FormatPolicy::Inherit,
        )
        .unwrap();
        doc.reject_change(
            &ctx(),
            &ChangeTarget::Revision(split.revision_ids[0].clone()),
        )
        .unwrap();
        // Ann's own typing goes with her paragraph; anyone else's keeps it.
        let expected = if typist == Some("Ann") {
            "Before¶[table]After¶"
        } else {
            "Before¶hello¶[table]After¶"
        };
        assert_eq!(slot_units(&doc), expected, "{typist:?}");
        assert!(marks(&doc).iter().all(|mark| !active(mark, "pPrIns")));
    }
}

#[test]
fn undoing_a_comment_removal_restores_the_comment_with_its_field() {
    let (doc, _) = doc_with("alpha beta");
    let system = EditCtx::system(DATE);
    doc.apply_raw_ops(
        "body",
        vec![RawOp::SetComment {
            id: "7".into(),
            ranges: vec![(0, 5)],
            author: "Ada".into(),
            date: DATE.into(),
            body: Any::Null,
        }],
        &system,
    )
    .unwrap();
    let reference = vec![
        (
            "modelKind".to_owned(),
            Any::String("commentReference".into()),
        ),
        ("commentId".to_owned(), Any::String("7".into())),
    ];
    doc.insert_embed(&system, Position::new("body", 5), "field", reference)
        .unwrap();
    let mut undo = doc.undo_manager();
    doc.apply_raw_ops(
        "body",
        vec![RawOp::RemoveComment { id: "7".into() }],
        &ctx(),
    )
    .unwrap();
    assert_eq!(slot_units(&doc), "alpha beta¶");
    assert!(undo.undo());
    assert_eq!(slot_units(&doc), "alpha[field] beta¶");
    assert_eq!(doc.resolve_comment("7").unwrap().len(), 1);
    assert!(undo.redo());
    assert!(doc.resolve_comment("7").is_err());
}

#[test]
fn a_split_leaves_the_section_with_the_mark_that_ends_it_and_borders_off_the_new_half() {
    let section = || {
        let attr = |key: &str, value: Any| RawOp::SetEmbedAttr {
            index: 9,
            key: key.into(),
            value,
        };
        let original = Any::Map(Arc::new(
            [
                ("borders".to_owned(), Any::String("top".into())),
                ("spaceAfter".to_owned(), Any::Number(480.0)),
            ]
            .into(),
        ));
        vec![
            attr(
                "sectPr",
                Any::Map(Arc::new(
                    [("sectionStart".to_owned(), Any::String("nextPage".into()))].into(),
                )),
            ),
            attr("sectionBreakType", Any::String("nextPage".into())),
            attr("borders", Any::String("top".into())),
            attr("_originalFormatting", original),
        ]
    };
    let original_borders = |mark: &BTreeMap<String, Any>| match mark.get("_originalFormatting") {
        Some(Any::Map(original)) => original.contains_key("borders"),
        _ => false,
    };
    for at in [5, 9] {
        // "AfterText" ends a section; Enter mid-paragraph or at its end.
        let (doc, _) = doc_with("AfterText");
        doc.apply_raw_ops("body", section(), &ctx()).unwrap();
        doc.split_paragraph(&ctx(), Position::new("body", at), None)
            .unwrap();
        let [first, second] = marks(&doc).try_into().unwrap();
        assert!(
            !first.contains_key("sectPr") && !first.contains_key("sectionBreakType"),
            "{at}"
        );
        assert!(
            second.contains_key("sectPr") && second.contains_key("sectionBreakType"),
            "{at}"
        );
        assert!(
            !second.contains_key("borders") && !original_borders(&second),
            "{at}"
        );
    }
    // Enter before a table: the new paragraph takes neither the section nor borders.
    let (doc, _) = block_slot("table");
    doc.apply_raw_ops(
        "body",
        section()
            .into_iter()
            .map(|op| match op {
                RawOp::SetEmbedAttr { key, value, .. } => RawOp::SetEmbedAttr {
                    index: 13,
                    key,
                    value,
                },
                op => op,
            })
            .collect(),
        &ctx(),
    )
    .unwrap();
    doc.split_paragraph(&ctx(), Position::new("body", 7), None)
        .unwrap();
    let [_, inserted, slot] = marks(&doc).try_into().unwrap();
    assert!(!inserted.contains_key("sectPr") && !inserted.contains_key("borders"));
    assert!(!original_borders(&inserted));
    assert!(slot.contains_key("sectPr") && slot.contains_key("borders") && original_borders(&slot));
}
