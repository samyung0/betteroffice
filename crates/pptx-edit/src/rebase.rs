use std::collections::{BTreeMap, HashMap, HashSet};
use std::sync::Arc;

use sha2::{Digest, Sha256};
use yrs::{Any, Doc, Map, Out, ReadTxn, Transact, WriteTxn};

use crate::{DeckSession, DeckSnapshot, EditError, EditResult, META, ShapeSnapshot};

const SOURCE_OVERLAY: &str = "sourceOverlay";

/// A fresh current state and a transient indexed projection in its identity space.
pub struct CheckpointRebase {
    pub state: Vec<u8>,
    pub indexed_state: Vec<u8>,
}

impl DeckSession {
    pub fn rebase_checkpoint(
        old_source: &[u8],
        captured_state: &[u8],
        latest_state: &[u8],
        new_source: &[u8],
        client_id: u64,
    ) -> EditResult<CheckpointRebase> {
        let captured = Self::open_from_update_with_source(captured_state, old_source, client_id)?;
        let latest = Self::open_from_update_with_source(latest_state, old_source, client_id)?;
        let exported = captured.save()?;
        if parts(&exported)? != parts(new_source)? {
            return Err(invalid(
                "published package does not match captured checkpoint",
            ));
        }
        let current_source = latest.save()?;
        let current = Self::open(&current_source, client_id)?;
        let indexed = Self::open(new_source, client_id)?;
        let captured_snapshot = captured.snapshot()?;
        let latest_snapshot = latest.snapshot()?;
        let current_snapshot = current.snapshot()?;
        let mut indexed_snapshot = indexed.snapshot()?;
        let current_ids = export_ids(&latest_snapshot, &current_snapshot)?;
        let indexed_ids = export_ids(&captured_snapshot, &indexed_snapshot)?;
        remap_baseline(&mut indexed_snapshot, &indexed_ids, &current_ids)?;
        crate::deck::seed_snapshot(&indexed.doc, &indexed_snapshot)?;
        set_overlay(&current.doc, new_source, &current_source)?;
        let state = current.encode_state_as_update_v1();
        let reopened = Self::open_from_update_with_source(&state, new_source, client_id)?;
        if parts(&reopened.save()?)? != parts(&current_source)? {
            return Err(invalid(
                "rebased source does not reproduce the saved checkpoint",
            ));
        }
        Ok(CheckpointRebase {
            state,
            indexed_state: indexed.encode_state_as_update_v1(),
        })
    }
}

fn invalid(message: &str) -> EditError {
    EditError::InvalidState(message.to_owned())
}

fn parts(bytes: &[u8]) -> EditResult<BTreeMap<String, Vec<u8>>> {
    ooxml_opc::unzip_parts(bytes)
        .map(|parts| parts.into_iter().collect())
        .map_err(|error| EditError::Parse(error.to_string()))
}

/// Parts that differ from the published base are stored as bytes, media included.
fn set_overlay(doc: &Doc, base: &[u8], current: &[u8]) -> EditResult<()> {
    let fingerprint = format!("{:x}", Sha256::digest(base));
    let base = parts(base)?;
    let current = parts(current)?;
    let mut overlay = Vec::new();
    for (path, bytes) in &current {
        if base.get(path) == Some(bytes) {
            continue;
        }
        overlay.push(Any::Array(Arc::from([
            Any::from(path.as_str()),
            Any::Buffer(Arc::from(bytes.as_slice())),
        ])));
    }
    for path in base.keys().filter(|path| !current.contains_key(*path)) {
        overlay.push(Any::Array(Arc::from([Any::from(path.as_str())])));
    }
    let mut txn = doc.transact_mut_with("pptx:rebase");
    let meta = txn.get_or_insert_map(META);
    meta.insert(&mut txn, "fingerprint", fingerprint);
    meta.insert(&mut txn, SOURCE_OVERLAY, Any::Array(Arc::from(overlay)));
    Ok(())
}

pub(crate) fn source_package(doc: &Doc, source: &[u8]) -> EditResult<pptx_parse::PptxPackage> {
    let txn = doc.transact();
    let meta = txn
        .get_map(META)
        .ok_or_else(|| invalid("missing metadata"))?;
    let Some(value) = meta.get(&txn, SOURCE_OVERLAY) else {
        return pptx_parse::parse_pptx(source).map_err(|error| EditError::Parse(error.to_string()));
    };
    let Out::Any(Any::Array(overlay)) = value else {
        return Err(invalid("invalid source overlay"));
    };
    let mut parts = parts(source)?;
    let mut seen = HashSet::new();
    for entry in overlay.iter() {
        let Any::Array(fields) = entry else {
            return Err(invalid("invalid source overlay entry"));
        };
        let Some(Any::String(path)) = fields.first() else {
            return Err(invalid("invalid source overlay path"));
        };
        if !seen.insert(path.to_string()) {
            return Err(invalid("duplicate source overlay path"));
        }
        match fields.as_ref() {
            [_] => {
                parts.remove(path.as_ref());
            }
            [_, Any::Buffer(bytes)] => {
                parts.insert(path.to_string(), bytes.to_vec());
            }
            _ => return Err(invalid("invalid source overlay content")),
        }
    }
    let bytes = ooxml_opc::rezip_parts(&parts.into_iter().collect::<Vec<_>>())
        .map_err(|error| EditError::Parse(error.to_string()))?;
    pptx_parse::parse_pptx(&bytes).map_err(|error| EditError::Parse(error.to_string()))
}

#[derive(Default)]
struct ExportIds {
    ids: HashMap<String, String>,
    shapes: HashMap<String, u32>,
}

fn export_ids(authored: &DeckSnapshot, exported: &DeckSnapshot) -> EditResult<ExportIds> {
    if authored.slides.len() != exported.slides.len() {
        return Err(invalid("export changed slide count"));
    }
    let mut ids = ExportIds::default();
    for (old, new) in authored.slides.iter().zip(&exported.slides) {
        if old.source_part_path.is_some() && old.source_part_path != new.source_part_path {
            return Err(invalid("export changed a source slide binding"));
        }
        insert_id(&mut ids, &old.id, &new.id)?;
        shape_ids(&mut ids, &old.shapes, &new.shapes)?;
    }
    Ok(ids)
}

fn insert_id(ids: &mut ExportIds, old: &str, new: &str) -> EditResult<()> {
    if ids.ids.insert(old.to_owned(), new.to_owned()).is_some() {
        return Err(invalid("duplicate authored identity"));
    }
    Ok(())
}

fn shape_ids(
    ids: &mut ExportIds,
    authored: &[ShapeSnapshot],
    exported: &[ShapeSnapshot],
) -> EditResult<()> {
    if authored.len() != exported.len() {
        return Err(invalid("export changed shape count"));
    }
    for (old, new) in authored.iter().zip(exported) {
        if old.kind != new.kind
            || old.text_stories.len() != new.text_stories.len()
            || (old.source_id != 0 && old.source_id != new.source_id)
        {
            return Err(invalid("export changed shape structure"));
        }
        insert_id(ids, &old.id, &new.id)?;
        ids.shapes.insert(old.id.clone(), new.source_id);
        for (old, new) in old.text_stories.iter().zip(&new.text_stories) {
            if old.plain_text() != new.plain_text() {
                return Err(invalid("export changed authored text"));
            }
            insert_id(ids, &old.id, &new.id)?;
            if old.paragraphs.len() != new.paragraphs.len() {
                return Err(invalid("export changed paragraph count"));
            }
            for (old, new) in old.paragraphs.iter().zip(&new.paragraphs) {
                insert_id(ids, &old.id, &new.id)?;
            }
        }
        shape_ids(ids, &old.children, &new.children)?;
    }
    Ok(())
}

fn remap_baseline(
    snapshot: &mut DeckSnapshot,
    indexed: &ExportIds,
    current: &ExportIds,
) -> EditResult<()> {
    let reverse = indexed
        .ids
        .iter()
        .map(|(old, fresh)| (fresh.as_str(), old.as_str()))
        .collect::<HashMap<_, _>>();
    let mapped = |id: &str| -> EditResult<String> {
        let old = reverse
            .get(id)
            .ok_or_else(|| invalid("missing indexed export identity"))?;
        Ok(current
            .ids
            .get(*old)
            .cloned()
            .unwrap_or_else(|| format!("rebase:removed:{id}")))
    };
    fn shapes(
        values: &mut [ShapeSnapshot],
        reverse: &HashMap<&str, &str>,
        current: &ExportIds,
        mapped: &impl Fn(&str) -> EditResult<String>,
    ) -> EditResult<()> {
        for shape in values {
            if let Some(source_id) = reverse
                .get(shape.id.as_str())
                .and_then(|old| current.shapes.get(*old))
            {
                shape.source_id = *source_id;
            }
            shape.id = mapped(&shape.id)?;
            for story in &mut shape.text_stories {
                story.id = mapped(&story.id)?;
                for paragraph in &mut story.paragraphs {
                    paragraph.id = mapped(&paragraph.id)?;
                }
            }
            shapes(&mut shape.children, reverse, current, mapped)?;
        }
        Ok(())
    }
    for slide in &mut snapshot.slides {
        slide.id = mapped(&slide.id)?;
        shapes(&mut slide.shapes, &reverse, current, &mapped)?;
    }
    Ok(())
}
