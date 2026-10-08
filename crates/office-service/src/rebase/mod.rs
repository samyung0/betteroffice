//! office-checkpoint.ts `rebaseOffice` and office-rebase.ts: the edits saved
//! after a capture landed on seed(export), with their effects against the
//! export's derived baseline.

mod align;
mod docx;
mod transplant;

use std::collections::{HashMap, HashSet};

use pptx_edit::DeckSession;
use yrs::{Map, Out, Transact};

use crate::common::{CHECKPOINT_MISMATCH, sha256_hex};
use crate::env;
use crate::error::{Error, Result};
use crate::jsv::{V, stringify};
use crate::session::{compare_baselines, unshifted_effects};
use crate::types::{Checkpoint, EffectKind, Format, NetEffect, Rebased};
use crate::xlsx;
use transplant::{DOCX_LINEAGE, PPTX_LINEAGE, open, ordered_keys, root_map, transplant as land};

/// A refusal: `Office rebase: <message>`.
pub(crate) fn fail(message: impl Into<String>) -> Error {
    Error::Rebase(message.into())
}

pub(crate) fn rebase(
    base: &[u8],
    captured: Checkpoint,
    latest: Checkpoint,
    exported: &[u8],
) -> Result<Rebased> {
    let format = captured.format;
    if base.is_empty() || exported.is_empty() {
        return Err(Error::Invalid(
            "Expected supported Office format and nonempty source bytes".into(),
        ));
    }
    let base_sha256 = sha256_hex(base);
    for checkpoint in [captured, latest] {
        if checkpoint.format != format
            || checkpoint.schema_version != 1
            || checkpoint.base_sha256 != base_sha256
            || checkpoint.state.is_empty()
        {
            return Err(Error::Engine(CHECKPOINT_MISMATCH.into()));
        }
    }
    let exported_sha256 = sha256_hex(exported);
    if format == Format::Xlsx {
        let state = xlsx::rebase(base, captured.state, latest.state, exported)?;
        let checkpoint = Checkpoint {
            format,
            schema_version: 1,
            base_sha256: &exported_sha256,
            state: &state,
        };
        xlsx::check(exported, checkpoint)?;
        let effects = xlsx::pending_effects(exported, checkpoint, None)?;
        return Ok(Rebased { state, effects });
    }
    // Override layout (spike): the transplant reads whole stories, so it runs
    // on materialized states and lands on the export's materialization; the
    // room keeps the copies the landed edits name.
    let chunked = format == Format::Docx && crate::docx::chunked_state(captured.state);
    let materialized_captured;
    let materialized_latest;
    let (captured, latest) = if chunked {
        materialized_captured = crate::docx::materialized(base, captured.state)?;
        materialized_latest = crate::docx::materialized(base, latest.state)?;
        (
            Checkpoint {
                state: &materialized_captured,
                ..captured
            },
            Checkpoint {
                state: &materialized_latest,
                ..latest
            },
        )
    } else {
        (captured, latest)
    };
    let seed = match format {
        Format::Docx if chunked => crate::docx::materialized(exported, &[])?,
        Format::Docx => crate::docx::seed(exported)?,
        _ => crate::pptx::seed(exported)?,
    };
    let ids = if format == Format::Docx {
        validate_export_backing(base, exported)?;
        docx::docx_ids(captured.state, &seed)?
    } else {
        pptx_ids(&deck(base, Some(captured.state))?, &deck(exported, None)?)?
    };
    let client = env::next_client();
    let lineage = if format == Format::Docx {
        &DOCX_LINEAGE
    } else {
        &PPTX_LINEAGE
    };
    let landed = land(lineage, captured.state, latest.state, &seed, &ids, client)?;
    let state = landed.state;
    if format == Format::Docx {
        assert_docx_restorations(base, latest.state, exported, &state, client, &landed.ids)?;
        assert_docx_renders(exported, &state)?;
    } else {
        assert_pptx_restorations(&seed, &state)?;
    }
    let (seed, state) = if chunked {
        (
            crate::docx::seed_of_layout(exported, true),
            crate::docx::shared_state(exported, &seed, &state)?,
        )
    } else {
        (seed, state)
    };
    let exported_checkpoint = |state| Checkpoint {
        format,
        schema_version: 1,
        base_sha256: &exported_sha256,
        state,
    };
    let (from, to) = (
        crate::baseline_unguarded(exported, exported_checkpoint(&seed))?,
        crate::baseline_unguarded(exported, exported_checkpoint(&state))?,
    );
    let saved = compare_baselines(
        &crate::baseline_unguarded(base, captured)?,
        &crate::baseline_unguarded(base, latest)?,
    );
    // The check compares every effect, moves included; the effects returned,
    // as save effects, leave out entries an edit only shifted (Epo, 2026-10-07).
    if change(format, &saved) != change(format, &compare_baselines(&from, &to)) {
        return Err(fail("the rebased state does not carry the saved edits"));
    }
    Ok(Rebased {
        state,
        effects: unshifted_effects(&from, &to),
    })
}

/// What a list of effects changes, as the rebase compares it: DOCX
/// formatting passes through an export its own way, and a comment body
/// compares by author and text.
fn change(format: Format, list: &[NetEffect]) -> String {
    let comment = |effect: &NetEffect, value: &Option<String>| -> V {
        match value {
            Some(value)
                if !value.is_empty()
                    && format == Format::Docx
                    && effect.id.starts_with("comment:") =>
            {
                V::from(comment_text(value))
            }
            Some(value) => V::from(value.as_str()),
            None => V::Null,
        }
    };
    let mut entries: Vec<String> = list
        .iter()
        .filter(|effect| format == Format::Pptx || effect.kind != EffectKind::Visual)
        .map(|effect| {
            let operation = serde_json::to_value(effect.operation).unwrap_or_default();
            let kind = serde_json::to_value(effect.kind).unwrap_or_default();
            stringify(&V::array(vec![
                V::from(&operation),
                V::from(&kind),
                comment(effect, &effect.before),
                comment(effect, &effect.after),
                effect.image_sha256.clone().map(V::from).unwrap_or(V::Null),
            ]))
            .unwrap_or_default()
        })
        .collect();
    entries.sort_by(|a, b| crate::js::utf16_compare(a, b));
    stringify(&V::array(entries.into_iter().map(V::from).collect())).unwrap_or_default()
}

/// A DOCX comment entry's author and visible text (`commentText`).
fn comment_text(value: &str) -> String {
    let parsed = V::parse(value).unwrap_or_default();
    let mut texts = String::new();
    fn walk(node: &V, texts: &mut String) {
        if let Some(items) = node.arr() {
            for item in items.items() {
                walk(&item, texts);
            }
        } else if let Some(object) = node.obj() {
            for (key, child) in object.entries() {
                match (&*key, child.as_str()) {
                    ("text", Some(text)) => texts.push_str(&text),
                    _ => walk(&child, texts),
                }
            }
        }
    }
    walk(&parsed.get("body"), &mut texts);
    stringify(&V::array(vec![parsed.get("author"), V::from(texts)])).unwrap_or_default()
}

/// The deck snapshot of a package, or of a state over it.
fn deck(bytes: &[u8], state: Option<&[u8]>) -> Result<V> {
    let session = match state {
        Some(state) => DeckSession::open_from_update_with_source(state, bytes, env::next_client()),
        None => DeckSession::open(bytes, env::next_client()),
    }
    .map_err(Error::engine)?;
    crate::pptx::deck_json(&session)
        .and_then(|json| V::parse(&crate::js::stringify(&json)).map_err(Error::Engine))
}

/// Captured PPTX ids to seed(export)'s, by their place in the deck (`pptxIds`).
fn pptx_ids(captured: &V, exported: &V) -> Result<Vec<(String, String)>> {
    let mut ids: Vec<(String, String)> = Vec::new();
    fn set(ids: &mut Vec<(String, String)>, from: &V, to: &V) {
        let (from, to) = (from.to_js_string(), to.to_js_string());
        match ids.iter_mut().find(|(key, _)| *key == from) {
            Some(entry) => entry.1 = to,
            None => ids.push((from, to)),
        }
    }
    let (mine, theirs) = (
        captured.get("slides").items(),
        exported.get("slides").items(),
    );
    if mine.len() != theirs.len() {
        return Err(fail("the export changed the slide count"));
    }
    fn shapes(ids: &mut Vec<(String, String)>, a: &[V], b: &[V]) -> Result<()> {
        if a.len() != b.len() {
            return Err(fail("the export changed a slide's shapes"));
        }
        for (shape, other) in a.iter().zip(b) {
            let (stories, others) = (
                shape.get("textStories").items(),
                other.get("textStories").items(),
            );
            if !shape.get("kind").same(&other.get("kind")) || stories.len() != others.len() {
                return Err(fail("the export changed a shape"));
            }
            set(ids, &shape.get("id"), &other.get("id"));
            for (story, theirs) in stories.iter().zip(&others) {
                let (paragraphs, their_paragraphs) = (
                    story.get("paragraphs").items(),
                    theirs.get("paragraphs").items(),
                );
                if paragraphs.len() != their_paragraphs.len() {
                    return Err(fail("the export changed a text's paragraphs"));
                }
                set(ids, &story.get("id"), &theirs.get("id"));
                for (paragraph, their) in paragraphs.iter().zip(&their_paragraphs) {
                    set(ids, &paragraph.get("id"), &their.get("id"));
                }
            }
            shapes(
                ids,
                &shape.get("children").items(),
                &other.get("children").items(),
            )?;
        }
        Ok(())
    }
    for (slide, other) in mine.iter().zip(&theirs) {
        set(&mut ids, &slide.get("id"), &other.get("id"));
        shapes(
            &mut ids,
            &slide.get("shapes").items(),
            &other.get("shapes").items(),
        )?;
    }
    let key = |comment: &V, id: &dyn Fn(String) -> String| {
        stringify(&V::array(vec![
            V::from(id(comment.get("slideId").to_js_string())),
            comment.get("author"),
            comment.get("text"),
            comment.get("created").or_else(|| V::Null),
        ]))
        .unwrap_or_default()
    };
    let mut theirs_by_key: HashMap<String, Vec<String>> = HashMap::new();
    for comment in exported.get("comments").items() {
        theirs_by_key
            .entry(key(&comment, &|value| value))
            .or_default()
            .push(comment.get("id").to_js_string());
    }
    for comment in captured.get("comments").items() {
        let lookup = |value: String| {
            ids.iter()
                .find(|(from, _)| *from == value)
                .map(|(_, to)| to.clone())
                .unwrap_or(value)
        };
        let found = theirs_by_key.get(&key(&comment, &lookup)).cloned();
        if let Some([single]) = found.as_deref() {
            set(&mut ids, &comment.get("id"), &V::from(single.as_str()));
        }
    }
    Ok(ids)
}

/// A PPTX slide or shape the later edits brought back is bound to source XML
/// the export dropped (`assertPptxRestorations`).
fn assert_pptx_restorations(seed: &[u8], state: &[u8]) -> Result<()> {
    let published = open(&PPTX_LINEAGE, seed, None)?;
    let rebased = open(&PPTX_LINEAGE, state, None)?;
    let (published, rebased) = (published.transact(), rebased.transact());
    for (root, bound) in [
        ("pptx:slides", "sourcePartPath"),
        ("pptx:shapes", "sourceId"),
    ] {
        let (kept, entities) = (root_map(&published, root), root_map(&rebased, root));
        for id in ordered_keys(&rebased, &entities) {
            if kept.contains_key(&published, &id) {
                continue;
            }
            let Some(Out::YMap(entity)) = entities.get(&rebased, &id) else {
                continue;
            };
            let bound_value = match entity.get(&rebased, bound) {
                Some(Out::Any(any)) => V::from(&serde_json::to_value(any).unwrap_or_default()),
                Some(_) => V::Bool(true),
                None => V::Undef,
            };
            if bound_value.truthy() {
                return Err(fail(format!(
                    "{root} {id} is bound to source XML the export dropped"
                )));
            }
        }
    }
    Ok(())
}

/// `decodeURIComponent`.
fn decode_uri_component(text: &str) -> Result<String> {
    let bytes = text.as_bytes();
    let mut out = Vec::with_capacity(bytes.len());
    let mut at = 0;
    while at < bytes.len() {
        if bytes[at] == b'%' {
            let hex = text
                .get(at + 1..at + 3)
                .ok_or_else(|| Error::engine("URI malformed"))?;
            out.push(u8::from_str_radix(hex, 16).map_err(|_| Error::engine("URI malformed"))?);
            at += 3;
        } else {
            out.push(bytes[at]);
            at += 1;
        }
    }
    String::from_utf8(out).map_err(|_| Error::engine("URI malformed"))
}

/// `decodeURIComponent(new URL(target, "https://docx.invalid/" + owner).pathname.slice(1))`.
fn target_path(owner: &str, target: &str) -> Result<String> {
    let internal = || fail("DOCX needs an internal package target");
    let target = target.replace('\\', "/");
    let scheme = target.find(':').is_some_and(|colon| {
        let head = &target[..colon];
        !head.is_empty()
            && head
                .chars()
                .next()
                .is_some_and(|ch| ch.is_ascii_alphabetic())
            && head
                .chars()
                .all(|ch| ch.is_ascii_alphanumeric() || matches!(ch, '+' | '-' | '.'))
    });
    if scheme || target.starts_with("//") {
        return Err(internal());
    }
    let (rest, fragment) = target.split_once('#').unwrap_or((&target, ""));
    let (path, query) = rest.split_once('?').unwrap_or((rest, ""));
    if !query.is_empty() || !fragment.is_empty() {
        return Err(internal());
    }
    let joined = if path.starts_with('/') {
        path.to_owned()
    } else {
        let directory = owner.rfind('/').map(|slash| &owner[..=slash]).unwrap_or("");
        format!("/{directory}{path}")
    };
    let mut segments: Vec<&str> = Vec::new();
    let parts: Vec<&str> = joined.split('/').skip(1).collect();
    for (index, segment) in parts.iter().enumerate() {
        let last = index + 1 == parts.len();
        match *segment {
            "." | "%2e" | "%2E" => {
                if last {
                    segments.push("");
                }
            }
            ".." | ".%2e" | "%2e." | "%2e%2e" | ".%2E" | "%2E." | "%2E%2E" => {
                segments.pop();
                if last {
                    segments.push("");
                }
            }
            other => segments.push(other),
        }
    }
    let encoded = segments
        .join("/")
        .replace(' ', "%20")
        .replace('"', "%22")
        .replace('<', "%3C")
        .replace('>', "%3E")
        .replace('`', "%60");
    decode_uri_component(&encoded)
}

/// office-checkpoint.ts `parseRelationshipsXmlWithRust` after a `TextDecoder`.
fn relationships(bytes: &[u8], path: &str) -> Result<docx_parse::RelationshipMap> {
    let text = String::from_utf8_lossy(bytes);
    let text = text.strip_prefix('\u{feff}').unwrap_or(&text);
    let limits = docx_parse::xml::ParseLimits::default();
    let mut budget = docx_parse::xml::ParseBudget::new(&limits);
    docx_parse::parse_relationships(text.as_bytes(), path, &mut budget).map_err(Error::engine)
}

/// The export keeps every package part a later edit may still name (`validateExportBacking`).
fn validate_export_backing(base: &[u8], exported: &[u8]) -> Result<()> {
    let old_parts = ooxml_opc::unzip_parts(base).map_err(Error::Engine)?;
    let new_parts: HashMap<String, Vec<u8>> = ooxml_opc::unzip_parts(exported)
        .map_err(Error::Engine)?
        .into_iter()
        .collect();
    let mut owned: HashSet<String> = [
        "[Content_Types].xml",
        "word/document.xml",
        "word/footnotes.xml",
        "word/endnotes.xml",
        "word/comments.xml",
        "word/commentsExtended.xml",
        "word/commentsIds.xml",
        "word/commentsExtensible.xml",
        "word/numbering.xml",
        "docProps/core.xml",
    ]
    .into_iter()
    .map(str::to_owned)
    .collect();
    for (path, bytes) in &old_parts {
        if !path.ends_with(".rels") {
            continue;
        }
        let next = new_parts
            .get(path)
            .ok_or_else(|| fail(format!("the export lost relationship part {path}")))?;
        let previous = relationships(bytes, path)?;
        let current = relationships(next, path)?;
        for (id, relationship) in &previous {
            let changed = current.get(id).is_none_or(|replacement| {
                replacement.relationship_type != relationship.relationship_type
                    || replacement.target != relationship.target
                    || replacement.target_mode != relationship.target_mode
            });
            if changed {
                return Err(fail(format!(
                    "the export changed source relationship {path}:{id}"
                )));
            }
            if relationship.relationship_type.ends_with("/header")
                || relationship.relationship_type.ends_with("/footer")
            {
                let owner = path.replacen("/_rels/", "/", 1);
                let owner = owner.strip_suffix(".rels").unwrap_or(&owner);
                owned.insert(target_path(owner, &relationship.target)?);
            }
        }
    }
    for (path, bytes) in &old_parts {
        if owned.contains(path) || path.ends_with(".rels") {
            continue;
        }
        if new_parts.get(path) != Some(bytes) {
            return Err(fail(format!(
                "the export changed or lost unmodeled source part {path}"
            )));
        }
    }
    Ok(())
}

/// The editor can open the rebased state: every story lowers (`assertDocxRenders`).
fn assert_docx_renders(source: &[u8], state: &[u8]) -> Result<()> {
    let session = crate::docx::DocxSession::open(source, Some(state))?;
    for story in crate::docx::read::story_ids(session.engine().doc()) {
        if let Err(error) = session
            .engine()
            .lower_story_json(&story, &docx_edit::bridge::RenderEnv::default())
        {
            return Err(fail(format!(
                "the rebased story {story} does not render: {error}"
            )));
        }
    }
    Ok(())
}

/// Every paragraph mark the rebase wrote takes the same raw inline markup
/// from the export as it took from the old source (`assertDocxRestorations`).
fn assert_docx_restorations(
    base: &[u8],
    latest: &[u8],
    exported: &[u8],
    state: &[u8],
    client: u64,
    ids: &[(String, String)],
) -> Result<()> {
    let written = docx::docx_written_paragraphs(state, client)?;
    if written.is_empty() {
        return Ok(());
    }
    let back: HashMap<&str, &str> = ids
        .iter()
        .map(|(from, to)| (to.as_str(), from.as_str()))
        .collect();
    let before = crate::docx::DocxSession::open(base, Some(latest))?.raw_inlines()?;
    let after = crate::docx::DocxSession::open(exported, Some(state))?.raw_inlines()?;
    for (story, id) in written {
        let source_story = back.get(story.as_str()).copied().unwrap_or(&story);
        let source_id = back.get(id.as_str()).copied().unwrap_or(&id);
        if before.get(&format!("{source_story}\u{0}{source_id}"))
            != after.get(&format!("{story}\u{0}{id}"))
        {
            return Err(fail(format!(
                "DOCX paragraph {id} needs content the export dropped"
            )));
        }
    }
    Ok(())
}
