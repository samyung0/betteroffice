//! Lockstep differential test of the DOCX override layout (spike): today's
//! seeded sessions and override-layout sessions run the same random
//! schedules (2–3 peers, partial syncs, Undo and Redo). Every peer reads the
//! same after every operation and every sync in both layouts, both worlds
//! converge, a fresh session on the room's override state reads the same,
//! and the two rooms export the same bytes.
//!
//! `CHUNKED_SEEDS` (schedules per fixture and peer count, default 6),
//! `CHUNKED_SEED` (one schedule), `CHUNKED_DOCX_DIR` (more fixtures),
//! `CHUNKED_TRACE` (a seed to trace).

use std::cell::RefCell;
use std::collections::{BTreeSet, HashMap};
use std::panic::AssertUnwindSafe;
use std::path::Path;
use std::rc::Rc;

use docx_edit::overlay;
use docx_edit::*;
use office_service::{Checkpoint, Determinism, Format};
use yrs::updates::decoder::Decode;
use yrs::{Any, ReadTxn, Transact};

#[path = "../../docx-edit/tests/support/chunked_fixtures.rs"]
mod chunked_fixtures;

const DATE: &str = "2026-10-08T00:00:00Z";

fn ctx() -> EditCtx {
    EditCtx::local("Ada", DATE)
}

struct Rng(u64);

impl Rng {
    fn next(&mut self) -> u64 {
        self.0 ^= self.0 << 13;
        self.0 ^= self.0 >> 7;
        self.0 ^= self.0 << 17;
        self.0
    }

    fn below(&mut self, n: usize) -> usize {
        if n == 0 {
            0
        } else {
            (self.next() % n as u64) as usize
        }
    }

    fn chance(&mut self, percent: u64) -> bool {
        self.next() % 100 < percent
    }
}

struct Peer {
    doc: EditingDoc,
    undo: UndoSession,
    outbox: Rc<RefCell<Vec<Vec<u8>>>>,
    _subscription: yrs::Subscription,
}

struct World {
    chunked: bool,
    peers: Vec<Peer>,
    /// Every message sent, with its author.
    log: Vec<(usize, Vec<u8>)>,
    /// Whether a message came from applying a peer's update (a forwarded
    /// cleanup, a rename or a field re-read) rather than from an op.
    system: Vec<bool>,
    seen: Vec<BTreeSet<usize>>,
    /// Cleanup messages a peer sent while applying another cleanup message.
    echoes: usize,
}

fn open_peer(base: &[u8], client: u64, chunked: bool) -> Peer {
    let doc = EditingDoc::new(client);
    if chunked {
        let envelope = parse_docx_for_edit(base).unwrap();
        overlay::open_chunked(&doc, envelope, &overlay::fingerprint(base)).unwrap();
    } else {
        seed_from_docx(&doc, base).unwrap();
    }
    let outbox = Rc::new(RefCell::new(Vec::new()));
    let sink = Rc::clone(&outbox);
    // D3: replicas send their formatting cleanup unless the mode is `on`.
    let subscription = doc
        .observe_shared_updates(cleanup_mode() != CleanupMode::On, move |update| {
            sink.borrow_mut().push(update)
        })
        .unwrap();
    let undo = UndoSession::new();
    undo.track(&doc);
    Peer {
        doc,
        undo,
        outbox,
        _subscription: subscription,
    }
}

impl World {
    fn new(base: &[u8], peers: usize, chunked: bool) -> Self {
        Self {
            chunked,
            peers: (0..peers)
                .map(|peer| open_peer(base, 701 + peer as u64, chunked))
                .collect(),
            log: Vec::new(),
            system: Vec::new(),
            seen: vec![BTreeSet::new(); peers],
            echoes: 0,
        }
    }

    fn drain(&mut self, peer: usize) {
        self.drain_as(peer, false);
    }

    /// Logs what `peer` sent; returns how many of them were cleanups.
    fn drain_as(&mut self, peer: usize, system: bool) -> usize {
        let sent: Vec<_> = self.peers[peer].outbox.borrow_mut().drain(..).collect();
        let mut cleanups = 0;
        for message in sent {
            cleanups += usize::from(system && blockless(&message));
            self.seen[peer].insert(self.log.len());
            self.log.push((peer, message));
            self.system.push(system);
        }
        cleanups
    }

    /// Delivers to `to` every message `from` has that `to` lacks.
    fn sync(&mut self, from: usize, to: usize) -> Result<(), String> {
        let missing: Vec<usize> = self.seen[from]
            .difference(&self.seen[to])
            .copied()
            .collect();
        for index in missing {
            let message = self.log[index].1.clone();
            self.seen[to].insert(index);
            self.peers[to]
                .doc
                .apply_peer_update_v1(&message)
                .map_err(|error| error.to_string())?;
            let cleanups = self.drain_as(to, true);
            if self.system[index] && blockless(&message) {
                self.echoes += cleanups;
            }
        }
        Ok(())
    }

    fn sync_all(&mut self) {
        let mut passes = 0;
        loop {
            passes += 1;
            if passes > 40 {
                let tail: Vec<String> = self.log[self.log.len().saturating_sub(4)..]
                    .iter()
                    .map(|(peer, message)| {
                        let update = yrs::Update::decode_v1(message).unwrap();
                        let text = format!("{update:?}");
                        format!(
                            "peer {peer}, {} B: {}",
                            message.len(),
                            &text[..text.len().min(1200)]
                        )
                    })
                    .collect();
                panic!(
                    "sync_all does not settle (chunked {}): {}",
                    self.chunked,
                    tail.join(
                        "
"
                    )
                );
            }
            let before = self.log.len();
            for from in 0..self.peers.len() {
                for to in 0..self.peers.len() {
                    if from != to {
                        self.sync(from, to).unwrap();
                    }
                }
            }
            if self.log.len() == before && self.seen.iter().all(|seen| seen.len() == before) {
                break;
            }
        }
    }

    /// The room after every message: the seed of its layout plus each one.
    fn room(&self, base: &[u8]) -> Vec<u8> {
        let seed = if self.chunked {
            overlay::meta_seed(&overlay::fingerprint(base))
        } else {
            let doc = EditingDoc::new(1);
            seed_from_docx(&doc, base).unwrap();
            doc.encode_state_as_update_v1()
        };
        let mut options = yrs::Options::with_client_id(yrs::ClientID::new(9));
        options.offset_kind = yrs::OffsetKind::Utf16;
        let room = yrs::Doc::with_options(options);
        {
            let mut txn = room.transact_mut();
            txn.apply_update(yrs::Update::decode_v1(&seed).unwrap())
                .unwrap();
            for (_, message) in &self.log {
                txn.apply_update(yrs::Update::decode_v1(message).unwrap())
                    .unwrap();
            }
        }
        let txn = room.transact();
        assert!(
            txn.store().pending_update().is_none() && txn.store().pending_ds().is_none(),
            "the room holds pending content"
        );
        txn.encode_state_as_update_v1(&yrs::StateVector::default())
    }
}

#[derive(Clone, Debug)]
enum Op {
    Insert(String, u32, String),
    Delete(String, u32, u32),
    Replace(String, u32, u32, String),
    Split(String, u32),
    /// A paragraph by story and index (its id differs by layout once renamed).
    Merge(String, usize, bool),
    Bold(String, u32, u32),
    Italic(String, u32, u32),
    Align(String, usize),
    List(String, usize),
    Paste(String, u32, u32, Vec<String>),
    PageBreak(String, u32),
    SectionBreak(String, u32),
    Table(String, u32),
    Row(String),
    DeleteTable(String),
    Comment(String, u32, u32),
    Uncomment(String),
    Accept(String, u32, u32),
    Reject(String, u32, u32),
    Suggest(String, u32, String),
    Toc,
    SelectAll(String),
    Undo,
    Redo,
}

/// A story's length and its paragraphs as (paraId, first unit, mark index).
fn layout(doc: &EditingDoc, story: &str) -> (u32, Vec<(String, u32, u32)>) {
    let mut offset = 0;
    let mut start = 0;
    let mut paragraphs = Vec::new();
    for segment in doc.story_segments(story).unwrap_or_default() {
        match segment.content {
            SegmentContent::Text(text) => offset += text.encode_utf16().count() as u32,
            SegmentContent::Pilcrow(properties) => {
                paragraphs.push((properties.para_id, start, offset));
                offset += 1;
                start = offset;
            }
            SegmentContent::OtherEmbed { .. } => offset += 1,
        }
    }
    (offset, paragraphs)
}

fn stories(doc: &EditingDoc) -> Vec<String> {
    let mut ids: Vec<String> = doc.all_story_segments().into_keys().collect();
    ids.sort();
    ids
}

fn comment_ids(doc: &EditingDoc) -> Vec<String> {
    let txn = doc.yrs_doc().transact();
    let mut ids: Vec<String> = txn
        .get_map("comments")
        .map(|map| {
            use yrs::Map;
            map.keys(&txn).map(str::to_owned).collect()
        })
        .unwrap_or_default();
    ids.sort();
    ids
}

fn choose(rng: &mut Rng, doc: &EditingDoc, focus: Option<&str>, marker: String) -> Op {
    let all = stories(doc);
    let story = if rng.chance(75) || all.len() == 1 {
        "body".to_owned()
    } else {
        all[rng.below(all.len())].clone()
    };
    let (len, paragraphs) = layout(doc, &story);
    if paragraphs.is_empty() || len == 0 {
        return Op::Undo;
    }
    let focused = focus.and_then(|id| paragraphs.iter().find(|(para, _, _)| para == id));
    let focused_index = focus.and_then(|id| paragraphs.iter().position(|(para, _, _)| para == id));
    let in_body = story == "body";
    let count = paragraphs.len();
    let pick_index = move |rng: &mut Rng| -> usize {
        match focused_index {
            Some(index) if in_body => index,
            _ => rng.below(count),
        }
    };
    let pick = |rng: &mut Rng| -> (String, u32, u32) {
        match focused {
            Some(paragraph) if story == "body" => paragraph.clone(),
            _ => paragraphs[rng.below(paragraphs.len())].clone(),
        }
    };
    // A position in a paragraph, often at its start or end (chunk boundaries).
    let at = |rng: &mut Rng, (_, start, mark): &(String, u32, u32)| -> u32 {
        match rng.below(4) {
            0 => *start,
            1 => *mark,
            _ => start + rng.below((mark - start + 1) as usize) as u32,
        }
    };
    let kind = if focus.is_some() {
        rng.below(3)
    } else {
        rng.below(32)
    };
    match kind {
        0..=4 => {
            let paragraph = pick(rng);
            Op::Insert(story, at(rng, &paragraph), marker)
        }
        5 | 6 => {
            let paragraph = pick(rng);
            let from = at(rng, &paragraph);
            let to = (from + 1 + rng.below(6) as u32).min(len - 1);
            if to <= from {
                return Op::Undo;
            }
            Op::Delete(story, from, to)
        }
        7 => {
            let first = rng.below(paragraphs.len());
            let last = (first + 1 + rng.below(3)).min(paragraphs.len() - 1);
            let from = at(rng, &paragraphs[first]);
            let to = at(rng, &paragraphs[last]).max(from + 1).min(len - 1);
            if to <= from {
                return Op::Undo;
            }
            Op::Delete(story, from, to)
        }
        8 => {
            let paragraph = pick(rng);
            let from = at(rng, &paragraph);
            let to = (from + rng.below(8) as u32).min(len - 1);
            Op::Replace(story, from, to.max(from), marker)
        }
        9..=11 => {
            let paragraph = pick(rng);
            Op::Split(story, at(rng, &paragraph))
        }
        12 | 13 => {
            let index = pick_index(rng);
            Op::Merge(story, index, rng.chance(50))
        }
        14 => {
            let first = rng.below(paragraphs.len());
            let last = (first + rng.below(3)).min(paragraphs.len() - 1);
            let from = at(rng, &paragraphs[first]);
            let to = at(rng, &paragraphs[last]).max(from + 1).min(len);
            Op::Bold(story, from, to)
        }
        15 => {
            let index = pick_index(rng);
            Op::Align(story, index)
        }
        16 => {
            let index = pick_index(rng);
            Op::List(story, index)
        }
        17 => {
            let first = rng.below(paragraphs.len());
            let last = (first + rng.below(3)).min(paragraphs.len() - 1);
            let from = at(rng, &paragraphs[first]);
            let to = at(rng, &paragraphs[last]).max(from).min(len - 1);
            Op::Paste(
                story,
                from,
                to,
                vec![
                    format!("{marker}a"),
                    format!("{marker}b"),
                    format!("{marker}c"),
                ],
            )
        }
        18 => {
            let paragraph = pick(rng);
            Op::PageBreak(story, at(rng, &paragraph))
        }
        19 => {
            let paragraph = pick(rng);
            Op::SectionBreak(story, paragraph.2)
        }
        20 => {
            let paragraph = pick(rng);
            Op::Table(story, paragraph.1)
        }
        21 if rng.chance(30) => Op::DeleteTable(story),
        21 => Op::Row(story),
        22 => {
            let first = rng.below(paragraphs.len());
            let last = (first + 1 + rng.below(3)).min(paragraphs.len() - 1);
            let from = at(rng, &paragraphs[first]);
            let to = at(rng, &paragraphs[last]).max(from + 1).min(len);
            Op::Comment(story, from, to)
        }
        23 => {
            let ids = comment_ids(doc);
            if ids.is_empty() {
                return Op::Redo;
            }
            Op::Uncomment(ids[rng.below(ids.len())].clone())
        }
        24 | 25 => {
            let first = rng.below(paragraphs.len());
            let last = (first + rng.below(4)).min(paragraphs.len() - 1);
            let from = paragraphs[first].1;
            let to = paragraphs[last].2.max(from + 1).min(len);
            if kind == 24 {
                Op::Accept(story, from, to)
            } else {
                Op::Reject(story, from, to)
            }
        }
        26 => {
            let paragraph = pick(rng);
            Op::Suggest(story, at(rng, &paragraph), marker)
        }
        27 => Op::Toc,
        28 if rng.chance(15) => Op::SelectAll(story),
        28 | 29 => Op::Undo,
        _ => Op::Redo,
    }
}

fn run(peer: &Peer, op: &Op) -> Result<(), String> {
    let doc = &peer.doc;
    let para_at = |story: &str, index: usize| -> Result<String, String> {
        layout(doc, story)
            .1
            .get(index)
            .map(|paragraph| paragraph.0.clone())
            .ok_or_else(|| format!("no paragraph {index} in {story}"))
    };
    let text = |result: Result<(), String>| result;
    let result: Result<(), String> = match op {
        Op::Insert(story, at, text) => doc
            .insert_text(&ctx(), Position::new(story, *at), text, FormatPolicy::Inherit)
            .map(drop)
            .map_err(|error| error.to_string()),
        Op::Delete(story, from, to) => doc
            .delete_range(&ctx(), StoryRange::new(story, *from, *to))
            .map(drop)
            .map_err(|error| error.to_string()),
        Op::Replace(story, from, to, text) => doc
            .replace_range(&ctx(), StoryRange::new(story, *from, *to), text)
            .map(drop)
            .map_err(|error| error.to_string()),
        Op::Split(story, at) => doc
            .split_paragraph(&ctx(), Position::new(story, *at))
            .map(drop)
            .map_err(|error| error.to_string()),
        Op::Merge(story, index, forward) => para_at(story, *index).and_then(|para| doc
            .merge_paragraphs(
                &ctx(),
                &para,
                if *forward {
                    MergeDirection::Forward
                } else {
                    MergeDirection::Backward
                },
            )
            .map(drop)
            .map_err(|error| error.to_string())),
        Op::Bold(story, from, to) => doc
            .toggle_format(&ctx(), StoryRange::new(story, *from, *to), SimpleFormat::Bold)
            .map(drop)
            .map_err(|error| error.to_string()),
        Op::Italic(story, from, to) => doc
            .toggle_format(&ctx(), StoryRange::new(story, *from, *to), SimpleFormat::Italic)
            .map(drop)
            .map_err(|error| error.to_string()),
        Op::Align(story, index) => para_at(story, *index).and_then(|para| doc
            .set_paragraph_attrs(
                &ctx(),
                &ParaSelector::One(para),
                &ParaAttrDelta {
                    alignment: Patch::Set("center".into()),
                    ..Default::default()
                },
            )
            .map(drop)
            .map_err(|error| error.to_string())),
        Op::List(story, index) => para_at(story, *index).and_then(|para| doc
            .set_paragraph_attrs(
                &ctx(),
                &ParaSelector::One(para),
                &ParaAttrDelta {
                    other: [(
                        "numPr".to_owned(),
                        Some(Any::Map(std::sync::Arc::new(HashMap::from([
                            ("numId".to_owned(), Any::Number(1.0)),
                            ("ilvl".to_owned(), Any::Number(0.0)),
                        ])))),
                    )]
                    .into_iter()
                    .collect(),
                    ..Default::default()
                },
            )
            .map(drop)
            .map_err(|error| error.to_string())),
        Op::Paste(story, from, to, lines) => (|| {
            let index_of = |landed: Receipt| -> Result<u32, String> {
                let end = landed.range.ok_or("no landed range")?.end;
                layout(doc, story)
                    .1
                    .iter()
                    .find(|(para, _, _)| *para == end.para)
                    .map(|(_, start, _)| start + end.offset)
                    .ok_or_else(|| "landed paragraph not found".to_owned())
            };
            let mut caret = index_of(
                doc.replace_range(&ctx(), StoryRange::new(story, *from, *to), &lines[0])
                    .map_err(|error| error.to_string())?,
            )?;
            for line in &lines[1..] {
                doc.split_paragraph(&ctx(), Position::new(story, caret))
                    .map_err(|error| error.to_string())?;
                caret += 1;
                caret = index_of(
                    doc.insert_text(
                        &ctx(),
                        Position::new(story, caret),
                        line,
                        FormatPolicy::Inherit,
                    )
                    .map_err(|error| error.to_string())?,
                )?;
            }
            Ok(())
        })(),
        Op::PageBreak(story, at) => doc
            .insert_embed(&ctx(), Position::new(story, *at), "pageBreak", vec![])
            .map(drop)
            .map_err(|error| error.to_string()),
        Op::SectionBreak(story, at) => doc
            .insert_embed(
                &ctx(),
                Position::new(story, *at),
                "sectionBreak",
                vec![("type".to_owned(), Any::from("nextPage"))],
            )
            .map(drop)
            .map_err(|error| error.to_string()),
        Op::Table(story, at) => doc
            .insert_table(&ctx(), Position::new(story, *at), 2, 2, 15)
            .map(drop)
            .map_err(|error| error.to_string()),
        Op::Row(story) => doc
            .insert_row(&ctx(), &CellLoc::new(story, 0, 0, 0), true)
            .map(drop)
            .map_err(|error| error.to_string()),
        Op::DeleteTable(story) => doc
            .delete_table(&ctx(), &TableLocator::new(story, 0))
            .map(drop)
            .map_err(|error| error.to_string()),
        Op::Comment(story, from, to) => doc
            .add_comment(
                &[StoryRange::new(story, *from, *to)],
                "Ada",
                DATE,
                Any::from_json(
                    r#"[{"type":"paragraph","content":[{"type":"run","content":[{"type":"text","text":"a comment"}]}]}]"#,
                )
                .unwrap(),
            )
            .map(drop)
            .map_err(|error| error.to_string()),
        Op::Uncomment(id) => doc
            .apply_raw_ops("body", vec![RawOp::RemoveComment { id: id.clone() }], &ctx())
            .map_err(|error| error.to_string()),
        Op::Accept(story, from, to) => doc
            .accept_change(
                &ctx(),
                &ChangeTarget::Range(StoryRange::new(story, *from, *to)),
            )
            .map(drop)
            .map_err(|error| error.to_string()),
        Op::Reject(story, from, to) => doc
            .reject_change(
                &ctx(),
                &ChangeTarget::Range(StoryRange::new(story, *from, *to)),
            )
            .map(drop)
            .map_err(|error| error.to_string()),
        Op::Suggest(story, at, text) => doc
            .insert_text(
                &EditCtx::local("Bo", DATE).suggesting(),
                Position::new(story, *at),
                text,
                FormatPolicy::Inherit,
            )
            .map(drop)
            .map_err(|error| error.to_string()),
        Op::Toc => doc
            .update_toc(
                &ctx(),
                None,
                &TocLayout {
                    pages: HashMap::new(),
                    tab_twips: 9000,
                    empty_text: "No table of contents entries found.".into(),
                },
            )
            .map(drop)
            .map_err(|error| error.to_string()),
        Op::SelectAll(story) => {
            let (len, _) = layout(doc, story);
            if len < 2 {
                Ok(())
            } else {
                doc.delete_range(&ctx(), StoryRange::new(story, 0, len - 1))
                    .map(drop)
                    .map_err(|error| error.to_string())
            }
        }
        Op::Undo => {
            peer.undo.undo();
            Ok(())
        }
        Op::Redo => {
            peer.undo.redo();
            Ok(())
        }
    };
    peer.undo.add_undo_barrier();
    text(result)
}

/// Runs `op` and catches an engine panic, which then must happen alike.
fn attempt(peer: &Peer, op: &Op) -> Result<(), String> {
    match std::panic::catch_unwind(AssertUnwindSafe(|| run(peer, op))) {
        Ok(result) => result,
        Err(panic) => Err(format!(
            "panic: {}",
            panic
                .downcast_ref::<String>()
                .cloned()
                .or_else(|| panic.downcast_ref::<&str>().map(|s| (*s).to_owned()))
                .unwrap_or_default()
        )),
    }
}

/// Baseline object and image ids are their embeds' item ids
/// (`<client#clock>`), which differ by layout; positions and values do not.
fn object_ids_hidden(text: &str) -> String {
    let mut out = String::with_capacity(text.len());
    let mut chars = text.char_indices().peekable();
    while let Some((at, ch)) = chars.next() {
        if ch == '<' {
            let tail = &text[at + 1..];
            let digits = |s: &str| s.bytes().take_while(u8::is_ascii_digit).count();
            let client = digits(tail);
            if client > 0 && tail[client..].starts_with('#') {
                let clock = digits(&tail[client + 1..]);
                if clock > 0 && tail[client + 1 + clock..].starts_with('>') {
                    out.push_str("<id>");
                    for _ in 0..client + 1 + clock + 1 {
                        chars.next();
                    }
                    continue;
                }
            }
        }
        out.push(ch);
    }
    out
}

/// The document as the lockstep compares it: [`overlay::read_dump`] with
/// renamed paragraph ids (`{client}.{clock}`, from the rename after concurrent
/// splits) kept to their client. The clock is the restored mark's item clock,
/// which an Undo's restore order sets differently by layout.
fn dump(doc: &EditingDoc) -> String {
    let text = overlay::read_dump(doc);
    let mut out = String::with_capacity(text.len());
    let mut rest = text.as_str();
    while let Some(at) = rest.find("\"pilcrow\":\"") {
        let start = at + "\"pilcrow\":\"".len();
        out.push_str(&rest[..start]);
        rest = &rest[start..];
        let end = rest.find('"').unwrap_or(0);
        let id = &rest[..end];
        match id.split_once('.') {
            Some((client, clock))
                if !client.is_empty()
                    && client.bytes().all(|byte| byte.is_ascii_digit())
                    && !clock.is_empty()
                    && clock.bytes().all(|byte| byte.is_ascii_digit()) =>
            {
                out.push_str(client);
                out.push_str(".*");
            }
            _ => out.push_str(id),
        }
        rest = &rest[end..];
    }
    out.push_str(rest);
    out
}

/// Whether two exported packages differ only in `w14:paraId` values (a
/// renamed paragraph id carries an item clock, see [`dump`]).
fn same_but_para_ids(left: &[u8], right: &[u8]) -> bool {
    let mask = |xml: &[u8]| -> String {
        let text = String::from_utf8_lossy(xml);
        let mut out = String::with_capacity(text.len());
        let mut rest = text.as_ref();
        while let Some(at) = rest.find("w14:paraId=\"") {
            let start = at + "w14:paraId=\"".len();
            out.push_str(&rest[..start]);
            rest = &rest[start..];
            rest = &rest[rest.find('"').unwrap_or(0)..];
        }
        out.push_str(rest);
        out
    };
    let (Ok(left), Ok(right)) = (ooxml_opc::unzip_parts(left), ooxml_opc::unzip_parts(right))
    else {
        return false;
    };
    left.len() == right.len()
        && left
            .iter()
            .zip(&right)
            .all(|((a, x), (b, y))| a == b && (x == y || mask(x) == mask(y)))
}

/// Renamed paragraph ids (`{client}.{clock}`) kept to their client, as in
/// [`dump`].
fn renamed_ids_hidden(text: &str) -> String {
    let mut out = String::with_capacity(text.len());
    let mut rest = text;
    while let Some(at) = rest.find(":paragraph:") {
        let start = at + ":paragraph:".len();
        out.push_str(&rest[..start]);
        rest = &rest[start..];
        let client = rest.bytes().take_while(u8::is_ascii_digit).count();
        let clock = rest[client..]
            .strip_prefix('.')
            .map(|tail| tail.bytes().take_while(u8::is_ascii_digit).count());
        if let Some(clock) = clock.filter(|clock| client > 0 && *clock > 0) {
            out.push_str(&rest[..client]);
            out.push_str(".*");
            rest = &rest[client + 1 + clock..];
        }
    }
    out.push_str(rest);
    out
}

/// The items around paragraph mark `para_id` in `story`, for diagnosis.
fn item_window(doc: &EditingDoc, story: &str, para_id: &str) -> String {
    use yrs::Map;
    let txn = doc.yrs_doc().transact();
    let Some(yrs::Out::YText(text)) = txn.get_map("stories").and_then(|map| map.get(&txn, story))
    else {
        return String::new();
    };
    let mut items = Vec::new();
    let mut target = None;
    let mut item = <yrs::TextRef as AsRef<yrs::branch::Branch>>::as_ref(&text).start();
    while let Some(current) = item {
        let summary = match current.content() {
            yrs::block::ItemContent::String(text) => {
                if !para_id.is_empty() && text.as_str().contains(para_id) {
                    target = Some(items.len());
                }
                format!("{:?}", text.as_str())
            }
            yrs::block::ItemContent::Format(key, value) => format!("fmt {key}={value}"),
            yrs::block::ItemContent::Type(_) => {
                let map =
                    yrs::MapRef::from(yrs::branch::BranchPtr::from(match current.content() {
                        yrs::block::ItemContent::Type(branch) => branch.as_ref(),
                        _ => unreachable!(),
                    }));
                let kind = map
                    .get(&txn, "_kind")
                    .map(|value| value.to_string(&txn))
                    .unwrap_or_default();
                let id = map
                    .get(&txn, "paraId")
                    .map(|value| value.to_string(&txn))
                    .unwrap_or_default();
                if id == para_id {
                    target = Some(items.len());
                }
                format!("embed {kind} {id}")
            }
            yrs::block::ItemContent::Deleted(len) => format!("deleted {len}"),
            other => format!("{other:?}").chars().take(40).collect(),
        };
        let id = current.id();
        let client = if id.client.get() >= overlay::RESERVED_CLIENTS {
            "W".to_owned()
        } else {
            id.client.get().to_string()
        };
        items.push(format!(
            "{}{client}:{} {summary}",
            if current.is_deleted() { "x " } else { "  " },
            id.clock
        ));
        item = current.right();
    }
    let at = target.unwrap_or(0);
    let after: usize = std::env::var("CHUNKED_WINDOW_AFTER")
        .ok()
        .and_then(|v| v.parse().ok())
        .unwrap_or(6);
    items[at.saturating_sub(
        std::env::var("CHUNKED_WINDOW_SIZE")
            .ok()
            .and_then(|v| v.parse().ok())
            .unwrap_or(14),
    )..(at + after).min(items.len())]
        .join(
            "
",
        )
}

/// A story's live items (content and format markers, no ids, no deleted
/// items): equal in both layouts when an op wrote the same structure.
fn live_structure(doc: &EditingDoc, story: &str) -> Vec<String> {
    use yrs::Map;
    let txn = doc.yrs_doc().transact();
    let Some(yrs::Out::YText(text)) = txn.get_map("stories").and_then(|map| map.get(&txn, story))
    else {
        return Vec::new();
    };
    let mut out = Vec::new();
    let mut item = <yrs::TextRef as AsRef<yrs::branch::Branch>>::as_ref(&text).start();
    while let Some(current) = item {
        if !current.is_deleted() {
            out.push(match current.content() {
                yrs::block::ItemContent::String(text) => format!("{:?}", text.as_str()),
                yrs::block::ItemContent::Format(key, value) => {
                    format!("fmt {key}={}", canon(value))
                }
                yrs::block::ItemContent::Type(_) => "embed".to_owned(),
                other => format!("{other:?}").chars().take(30).collect(),
            });
        }
        item = current.right();
    }
    fn canon(value: &Any) -> String {
        match value {
            Any::Map(map) => {
                let mut entries: Vec<_> = map
                    .iter()
                    .map(|(key, value)| format!("{key}:{}", canon(value)))
                    .collect();
                entries.sort();
                format!("{{{}}}", entries.join(","))
            }
            Any::Array(values) => format!(
                "[{}]",
                values.iter().map(canon).collect::<Vec<_>>().join(",")
            ),
            other => other.to_string(),
        }
    }
    // Text items split differently; join adjacent strings.
    let mut joined: Vec<String> = Vec::new();
    for entry in out {
        if entry.starts_with('"') && joined.last().is_some_and(|last| last.starts_with('"')) {
            let last = joined.pop().unwrap();
            joined.push(format!("{}{}", &last[..last.len() - 1], &entry[1..]));
        } else {
            joined.push(entry);
        }
    }
    joined
}

/// Quoted renamed paragraph ids (`"{client}.{clock}"`) kept to their
/// client, as in [`dump`].
fn quoted_renamed_ids_hidden(text: &str) -> String {
    let mut out = String::with_capacity(text.len());
    let mut rest = text;
    while let Some(at) = rest.find('"') {
        out.push_str(&rest[..=at]);
        rest = &rest[at + 1..];
        let end = rest.find('"').unwrap_or(rest.len());
        let value = &rest[..end];
        match value.split_once('.') {
            Some((client, clock))
                if !client.is_empty()
                    && !clock.is_empty()
                    && client.bytes().all(|byte| byte.is_ascii_digit())
                    && clock.bytes().all(|byte| byte.is_ascii_digit()) =>
            {
                out.push_str(client);
                out.push_str(".*");
            }
            _ => out.push_str(value),
        }
        rest = &rest[end..];
        if let Some(stripped) = rest.strip_prefix('"') {
            out.push('"');
            rest = stripped;
        }
    }
    out.push_str(rest);
    out
}

fn first_difference(left: &str, right: &str) -> String {
    for (index, (a, b)) in left.lines().zip(right.lines()).enumerate() {
        if a != b {
            return format!("line {index}:\n  seeded:  {a}\n  chunked: {b}");
        }
    }
    format!(
        "lengths {} / {} lines",
        left.lines().count(),
        right.lines().count()
    )
}

#[derive(Default, Debug)]
struct Stats {
    applied: usize,
    diverged: usize,
    exported: usize,
    refused: usize,
    panicked: usize,
    room_seeded: usize,
    room_chunked: usize,
    copies: usize,
}

fn lockstep(name: &str, base: &[u8], seed: u64, peers: usize, rounds: usize, ops: usize) -> Stats {
    let trace = std::env::var("CHUNKED_TRACE").ok() == Some(seed.to_string());
    let mut rng = Rng(seed.wrapping_mul(0x9E37_79B9_7F4A_7C15) | 1);
    // CHUNKED_CONTROL=1 runs today's layout in both worlds (a control).
    let control = std::env::var("CHUNKED_CONTROL").is_ok_and(|value| value == "1");
    let mut worlds = [
        World::new(base, peers, false),
        World::new(base, peers, !control),
    ];
    let mut stats = Stats::default();
    let focus = layout(&worlds[0].peers[0].doc, "body")
        .1
        .get(1)
        .map(|paragraph| paragraph.0.clone());
    let context = |round: usize, what: &str| format!("{name} seed {seed} round {round}: {what}");
    for round in 0..rounds {
        for step in 0..ops {
            for peer in 0..peers {
                let marker = format!("<{peer}.{round}.{step}>");
                let focused = (round == 0 && step == 0)
                    .then_some(focus.as_deref())
                    .flatten();
                let op = choose(&mut rng, &worlds[0].peers[peer].doc, focused, marker);
                let seeded = attempt(&worlds[0].peers[peer], &op);
                let chunked = attempt(&worlds[1].peers[peer], &op);
                if trace {
                    eprintln!("r{round} s{step} p{peer} {op:?} -> {seeded:?}");
                }
                assert_eq!(
                    seeded,
                    chunked,
                    "{}",
                    context(round, &format!("peer {peer} {op:?}"))
                );
                match &seeded {
                    Ok(()) => stats.applied += 1,
                    Err(error) if error.starts_with("panic") => stats.panicked += 1,
                    Err(_) => stats.refused += 1,
                }
                for world in worlds.iter_mut() {
                    world.drain(peer);
                }
                if let (Ok(step_id), Ok(marker)) = (
                    std::env::var("CHUNKED_WINDOW_STEP"),
                    std::env::var("CHUNKED_WINDOW"),
                ) && step_id == format!("r{round}s{step}p{peer}")
                {
                    eprintln!(
                        "after r{round}s{step}p{peer} seeded:
{}",
                        item_window(&worlds[0].peers[peer].doc, "body", &marker)
                    );
                    eprintln!(
                        "after r{round}s{step}p{peer} chunked:
{}",
                        item_window(&worlds[1].peers[peer].doc, "body", &marker)
                    );
                }
                if std::env::var("CHUNKED_STRUCTURE").is_ok() {
                    let (a, b) = (
                        live_structure(&worlds[0].peers[peer].doc, "body"),
                        live_structure(&worlds[1].peers[peer].doc, "body"),
                    );
                    if a != b {
                        let at = a
                            .iter()
                            .zip(&b)
                            .position(|(x, y)| x != y)
                            .unwrap_or(a.len().min(b.len()));
                        eprintln!(
                            "structure differs after r{round}s{step}p{peer} {op:?} at {at}:
  seeded:  {:?}
  chunked: {:?}",
                            &a[at.saturating_sub(4)..(at + 6).min(a.len())],
                            &b[at.saturating_sub(4)..(at + 6).min(b.len())]
                        );
                    }
                }
                let (left, right) = (
                    dump(&worlds[0].peers[peer].doc),
                    dump(&worlds[1].peers[peer].doc),
                );
                assert!(
                    left == right,
                    "{}",
                    context(
                        round,
                        &format!(
                            "peer {peer} after {op:?}: {}",
                            first_difference(&left, &right)
                        )
                    )
                );
            }
        }
        let pairs: Vec<(usize, usize)> = (0..peers * 2)
            .map(|_| (rng.below(peers), rng.below(peers)))
            .filter(|(from, to)| from != to)
            .collect();
        for (from, to) in pairs {
            let watch =
                std::env::var("CHUNKED_WINDOW_SYNC").ok() == Some(format!("r{round} {from}->{to}"));
            if watch && let Ok(marker) = std::env::var("CHUNKED_WINDOW") {
                eprintln!(
                    "before seeded:
{}",
                    item_window(&worlds[0].peers[to].doc, "body", &marker)
                );
                eprintln!(
                    "before chunked:
{}",
                    item_window(&worlds[1].peers[to].doc, "body", &marker)
                );
            }
            for world in worlds.iter_mut() {
                let chunked = world.chunked;
                world.sync(from, to).unwrap_or_else(|error| {
                    panic!(
                        "{}",
                        context(
                            round,
                            &format!("chunked {chunked} sync {from}->{to}: {error}")
                        )
                    )
                });
            }
            if trace {
                eprintln!("r{round} sync {from}->{to}");
            }
            if watch && let Ok(marker) = std::env::var("CHUNKED_WINDOW") {
                eprintln!(
                    "after seeded:
{}",
                    item_window(&worlds[0].peers[to].doc, "body", &marker)
                );
                eprintln!(
                    "after chunked:
{}",
                    item_window(&worlds[1].peers[to].doc, "body", &marker)
                );
            }
            if std::env::var("CHUNKED_STRUCTURE").is_ok() {
                let (a, b) = (
                    live_structure(&worlds[0].peers[to].doc, "body"),
                    live_structure(&worlds[1].peers[to].doc, "body"),
                );
                if a != b {
                    let at = a
                        .iter()
                        .zip(&b)
                        .position(|(x, y)| x != y)
                        .unwrap_or(a.len().min(b.len()));
                    eprintln!(
                        "structure differs after r{round} sync {from}->{to} at {at}:
  seeded:  {:?}
  chunked: {:?}",
                        &a[at.saturating_sub(6)..(at + 8).min(a.len())],
                        &b[at.saturating_sub(6)..(at + 8).min(b.len())]
                    );
                }
            }
            let (left, right) = (
                dump(&worlds[0].peers[to].doc),
                dump(&worlds[1].peers[to].doc),
            );
            if left != right {
                if let Ok(para) = std::env::var("CHUNKED_WINDOW") {
                    eprintln!(
                        "seeded:
{}",
                        item_window(&worlds[0].peers[to].doc, "body", &para)
                    );
                    eprintln!(
                        "chunked:
{}",
                        item_window(&worlds[1].peers[to].doc, "body", &para)
                    );
                }
                panic!(
                    "{}",
                    context(
                        round,
                        &format!("sync {from}->{to}: {}", first_difference(&left, &right))
                    )
                );
            }
        }
    }
    finish(
        &format!("{name} seed {seed}"),
        base,
        &mut worlds,
        &mut stats,
    );
    stats
}

fn checkpoint<'a>(sha: &'a str, state: &'a [u8]) -> Checkpoint<'a> {
    Checkpoint {
        format: Format::Docx,
        schema_version: 1,
        base_sha256: sha,
        state,
    }
}

/// Converges both worlds and checks them: every peer reads the same, the
/// worlds read the same, a fresh session on the override room reads the
/// same, and both rooms export the same file and the same baseline.
fn finish(label: &str, base: &[u8], worlds: &mut [World; 2], stats: &mut Stats) {
    let phase = |name: &str| {
        if std::env::var("CHUNKED_PHASES").is_ok() {
            eprintln!("{label}: {name}");
        }
    };
    for world in worlds.iter_mut() {
        world.sync_all();
    }
    phase("synced");
    // Peers of today's layout that do not converge (a fault of today's
    // engine, e.g. anchors that follow yrs's local-only redone links) must
    // read the same in the override layout, peer by peer.
    let dumps: Vec<Vec<String>> = worlds
        .iter()
        .map(|world| world.peers.iter().map(|peer| dump(&peer.doc)).collect())
        .collect();
    for (index, (left, right)) in dumps[0].iter().zip(&dumps[1]).enumerate() {
        assert!(
            left == right,
            "{label}: converged peer {index} differs by layout: {}",
            first_difference(left, right)
        );
    }
    if let Some(index) = dumps[0].iter().position(|other| *other != dumps[0][0]) {
        eprintln!(
            "{label}: today's layout does not converge (peer {index}), alike in both: {}",
            first_difference(&dumps[0][0], &dumps[0][index])
        );
        stats.diverged += 1;
        return;
    }
    phase("dumped");
    let (left, right) = (dump(&worlds[0].peers[0].doc), dump(&worlds[1].peers[0].doc));
    assert!(
        left == right,
        "{label}: converged worlds differ: {}",
        first_difference(&left, &right)
    );
    // The layout engine's input (the render bridge's blocks) agrees too.
    let blocks = |doc: &EditingDoc| {
        bridge::yrs_doc_to_layout_blocks(doc, "body", &bridge::RenderEnv::default())
            .map(|blocks| serde_json::to_string(&blocks).unwrap())
            .map_err(|error| format!("{error:?}"))
    };
    let (left_blocks, right_blocks) = (
        quoted_renamed_ids_hidden(&blocks(&worlds[0].peers[0].doc).unwrap_or_else(|error| error)),
        quoted_renamed_ids_hidden(&blocks(&worlds[1].peers[0].doc).unwrap_or_else(|error| error)),
    );
    if left_blocks != right_blocks {
        let split = |text: &str| {
            text.replace(
                "},{", "},
{",
            )
        };
        panic!(
            "{label}: layout blocks differ: {}",
            first_difference(&split(&left_blocks), &split(&right_blocks))
        );
    }
    phase("blocks compared");
    let rooms = [worlds[0].room(base), worlds[1].room(base)];
    if !worlds[1].chunked {
        return;
    }
    let fresh = EditingDoc::new(799);
    let envelope = parse_docx_for_edit(base).unwrap();
    overlay::open_chunked(&fresh, envelope, &overlay::fingerprint(base)).unwrap();
    fresh.apply_shared_update(&rooms[1]).unwrap();
    let reopened = dump(&fresh);
    assert!(
        reopened == left,
        "{label}: a fresh session on the chunked room differs: {}",
        first_difference(&left, &reopened)
    );
    // Read straight from the room and the source, nothing materialized.
    let logical = overlay::LogicalBase::new(base).unwrap();
    let room_doc = {
        let mut options = yrs::Options::with_client_id(yrs::ClientID::new(9));
        options.offset_kind = yrs::OffsetKind::Utf16;
        let doc = yrs::Doc::with_options(options);
        doc.transact_mut()
            .apply_update(yrs::Update::decode_v1(&rooms[1]).unwrap())
            .unwrap();
        doc
    };
    let (engine_view, room_view) = (
        overlay::session_segments_dump(&worlds[1].peers[0].doc),
        overlay::room_segments_dump(&logical, &room_doc),
    );
    assert!(
        engine_view == room_view,
        "{label}: the room read through the source differs: {}",
        first_difference(&engine_view, &room_view)
    );
    stats.copies = fresh.overlay().map_or(0, |overlay| overlay.shared().len());
    stats.room_seeded = rooms[0].len();
    stats.room_chunked = rooms[1].len();
    let sha = office_service::sha256_hex(base);
    let determinism = Determinism {
        seed: "0000000000000000000000000000000000000000000000000000000000000000",
        now: "2026-10-08T00:00:00.000Z",
    };
    phase("room read through the source");
    let exports: Vec<_> = rooms
        .iter()
        .map(|room| {
            let export = office_service::export(base, checkpoint(&sha, room), determinism);
            phase("exported one room");
            export
        })
        .collect();
    if exports[0].is_ok() {
        stats.exported += 1;
    }
    match (&exports[0], &exports[1]) {
        (Ok(left), Ok(right)) => assert!(
            left == right || same_but_para_ids(left, right),
            "{label}: exports differ"
        ),
        (left, right) => assert_eq!(
            left.as_ref().err().map(ToString::to_string),
            right.as_ref().err().map(ToString::to_string),
            "{label}: export outcome differs"
        ),
    }
    phase("exports compared");
    let baselines: Vec<_> = rooms
        .iter()
        .map(|room| {
            office_service::baseline(base, checkpoint(&sha, room))
                .map(|entries| renamed_ids_hidden(&object_ids_hidden(&format!("{entries:#?}"))))
                .map_err(|error| error.to_string())
        })
        .collect();
    if baselines[0] != baselines[1] {
        let text =
            |baseline: &Result<String, String>| baseline.clone().unwrap_or_else(|error| error);
        panic!(
            "{label}: baselines differ: {}",
            first_difference(&text(&baselines[0]), &text(&baselines[1]))
        );
    }
}

fn fixtures() -> Vec<(String, Vec<u8>)> {
    let mut fixtures = chunked_fixtures::xml_fixtures();
    let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
    for path in [
        "crates/office-service/tests/fixtures/stories.docx",
        "crates/betteroffice-docx/tests/corpus/fixtures/wordprocessingml-comprehensive.docx",
        "poc/fixtures/feature-rich.docx",
    ] {
        fixtures.push((
            Path::new(path)
                .file_stem()
                .unwrap()
                .to_string_lossy()
                .into_owned(),
            std::fs::read(root.join(path)).unwrap(),
        ));
    }
    if let Ok(dir) = std::env::var("CHUNKED_DOCX_DIR") {
        let mut paths: Vec<_> = std::fs::read_dir(dir)
            .unwrap()
            .flatten()
            .map(|entry| entry.path())
            .filter(|path| path.extension().is_some_and(|ext| ext == "docx"))
            .collect();
        paths.sort();
        for path in paths {
            fixtures.push((
                path.file_stem().unwrap().to_string_lossy().into_owned(),
                std::fs::read(&path).unwrap(),
            ));
        }
    }
    fixtures
}

#[test]
fn random_schedules_match_the_seeded_layout() {
    let seeds: u64 = std::env::var("CHUNKED_SEEDS")
        .ok()
        .and_then(|value| value.parse().ok())
        .unwrap_or(6);
    let only: Option<u64> = std::env::var("CHUNKED_SEED")
        .ok()
        .and_then(|value| value.parse().ok());
    let rounds: usize = std::env::var("CHUNKED_ROUNDS")
        .ok()
        .and_then(|value| value.parse().ok())
        .unwrap_or(4);
    let mut totals = Stats::default();
    let mut schedules = 0;
    let mut failures: Vec<String> = Vec::new();
    let fixture = std::env::var("CHUNKED_FIXTURE").ok();
    for (name, base) in fixtures() {
        if fixture.as_ref().is_some_and(|only| *only != name) {
            continue;
        }
        for seed in 0..seeds {
            for peers in [2, 3] {
                let seed = seed * 31 + peers as u64;
                if only.is_some_and(|only| only != seed) {
                    continue;
                }
                let outcome = std::panic::catch_unwind(AssertUnwindSafe(|| {
                    lockstep(&name, &base, seed, peers, rounds, 3)
                }));
                let Ok(stats) = outcome else {
                    failures.push(format!("{name} seed {seed} peers {peers}"));
                    continue;
                };
                eprintln!("{name} seed {seed} peers {peers}: {stats:?}");
                totals.applied += stats.applied;
                totals.refused += stats.refused;
                totals.panicked += stats.panicked;
                totals.copies += stats.copies;
                totals.diverged += stats.diverged;
                totals.exported += stats.exported;
                schedules += 1;
            }
        }
    }
    eprintln!(
        "lockstep: {schedules} schedules passed, {} failed, {totals:?}",
        failures.len()
    );
    assert!(failures.is_empty(), "failed schedules: {failures:#?}");
    assert!(totals.applied > totals.refused);
}

enum Step {
    Do(usize, Op),
    Sync(usize, usize),
}

/// The clocks of the session clients (below the reserved range).
fn user_clocks(doc: &EditingDoc) -> Vec<(u64, u32)> {
    let txn = doc.yrs_doc().transact();
    let mut clocks: Vec<(u64, u32)> = txn
        .state_vector()
        .iter()
        .map(|(client, clock)| (client.get(), *clock))
        .filter(|(client, _)| *client != 0 && *client < overlay::RESERVED_CLIENTS)
        .collect();
    clocks.sort();
    clocks
}

fn compare(worlds: &[World; 2], peer: usize, what: &str) {
    assert_eq!(
        user_clocks(&worlds[0].peers[peer].doc),
        user_clocks(&worlds[1].peers[peer].doc),
        "{what}: peer {peer} wrote different items"
    );
    let (left, right) = (
        dump(&worlds[0].peers[peer].doc),
        dump(&worlds[1].peers[peer].doc),
    );
    assert!(
        left == right,
        "{what}: peer {peer} reads differ: {}",
        first_difference(&left, &right)
    );
}

/// Runs scripted steps in both layouts, comparing reads after each, then
/// [`finish`]. Every scripted op must apply.
fn script(name: &str, base: &[u8], peers: usize, steps: Vec<Step>) -> [World; 2] {
    let mut worlds = [
        World::new(base, peers, false),
        World::new(base, peers, true),
    ];
    for (index, step) in steps.into_iter().enumerate() {
        match step {
            Step::Do(peer, op) => {
                let seeded = attempt(&worlds[0].peers[peer], &op);
                let chunked = attempt(&worlds[1].peers[peer], &op);
                assert_eq!(seeded, chunked, "{name} step {index} {op:?}");
                assert!(seeded.is_ok(), "{name} step {index} {op:?}: {seeded:?}");
                for world in worlds.iter_mut() {
                    world.drain(peer);
                }
                compare(&worlds, peer, &format!("{name} step {index} {op:?}"));
            }
            Step::Sync(from, to) => {
                for world in worlds.iter_mut() {
                    world.sync(from, to).unwrap();
                }
                compare(
                    &worlds,
                    to,
                    &format!("{name} step {index} sync {from}->{to}"),
                );
            }
        }
    }
    let mut stats = Stats::default();
    finish(name, base, &mut worlds, &mut stats);
    worlds
}

fn mixed() -> Vec<u8> {
    chunked_fixtures::xml_fixtures()
        .into_iter()
        .find(|(name, _)| name == "xml-mixed")
        .unwrap()
        .1
}

/// Paragraph `index` of the body as (paraId, first unit, mark index).
fn para(base: &[u8], index: usize) -> (String, u32, u32) {
    let doc = EditingDoc::new(1);
    seed_from_docx(&doc, base).unwrap();
    layout(&doc, "body").1[index].clone()
}

/// The reserved (copy) writers a room state holds.
fn copies_in(state: &[u8]) -> BTreeSet<u64> {
    let update = yrs::Update::decode_v1(state).unwrap();
    update
        .blocks()
        .map(|block| match block {
            yrs::UpdateBlock::Item(item) => item.id().client.get(),
            yrs::UpdateBlock::Gc(range) | yrs::UpdateBlock::Skip(range) => range.client.get(),
        })
        .filter(|client| *client >= overlay::RESERVED_CLIENTS)
        .collect()
}

fn body_text(doc: &EditingDoc) -> String {
    doc.story_segments("body")
        .unwrap()
        .into_iter()
        .filter_map(|segment| match segment.content {
            SegmentContent::Text(text) => Some(text),
            _ => None,
        })
        .collect()
}

fn insert(story: &str, at: u32, text: &str) -> Op {
    Op::Insert(story.to_owned(), at, text.to_owned())
}

#[test]
fn concurrent_first_edits_of_one_chunk_converge_on_one_copy() {
    let base = mixed();
    let (_, start, mark) = para(&base, 1);
    let worlds = script(
        "concurrent first edits",
        &base,
        3,
        vec![
            Step::Do(0, insert("body", start, "<A>")),
            Step::Do(1, insert("body", start + 4, "<B>")),
            Step::Do(2, insert("body", mark, "<C>")),
            Step::Sync(1, 0),
            Step::Sync(2, 0),
        ],
    );
    let room = worlds[1].room(&base);
    let overlay = worlds[1].peers[0].doc.overlay().unwrap();
    let body = overlay
        .stories
        .iter()
        .find(|story| story.id == "body")
        .unwrap();
    assert_eq!(
        copies_in(&room),
        BTreeSet::from([body.writer, body.chunk_writers[1]]),
        "one story copy and one chunk copy"
    );
    let text = body_text(&worlds[1].peers[0].doc);
    for marker in ["<A>", "<B>", "<C>", "Beta", "tracked insert"] {
        assert_eq!(text.matches(marker).count(), 1, "{marker} once in {text}");
    }
}

#[test]
fn enter_at_both_sides_of_a_chunk_boundary() {
    let base = mixed();
    let (_, _, end_of_first) = para(&base, 0);
    let (_, start_of_second, _) = para(&base, 1);
    script(
        "enter across a chunk boundary",
        &base,
        2,
        vec![
            Step::Do(0, Op::Split("body".into(), end_of_first)),
            Step::Do(1, Op::Split("body".into(), start_of_second)),
            Step::Do(1, insert("body", start_of_second, "<new>")),
            Step::Sync(0, 1),
            Step::Sync(1, 0),
            Step::Do(0, Op::Undo),
            Step::Do(1, Op::Undo),
            Step::Do(1, Op::Redo),
        ],
    );
}

#[test]
fn backspace_and_delete_join_paragraphs_across_chunks() {
    let base = mixed();
    let (_, start_of_third, _) = para(&base, 2);
    script(
        "joins across chunks",
        &base,
        2,
        vec![
            Step::Do(0, Op::Merge("body".into(), 2, false)),
            Step::Do(1, Op::Merge("body".into(), 1, true)),
            Step::Sync(0, 1),
            Step::Do(1, insert("body", start_of_third - 2, "<joined>")),
            Step::Sync(1, 0),
            Step::Do(0, Op::Undo),
        ],
    );
}

#[test]
fn select_all_delete_against_a_concurrent_insert() {
    let base = mixed();
    let (_, start, _) = para(&base, 4);
    let worlds = script(
        "select-all delete",
        &base,
        2,
        vec![
            Step::Do(0, Op::SelectAll("body".into())),
            Step::Do(1, insert("body", start + 2, "<kept>")),
            Step::Sync(0, 1),
            Step::Sync(1, 0),
        ],
    );
    assert!(body_text(&worlds[1].peers[0].doc).contains("<kept>"));
}

#[test]
fn paste_over_paragraphs_in_several_chunks() {
    let base = mixed();
    let (_, from, _) = para(&base, 1);
    let (_, to, _) = para(&base, 3);
    script(
        "paste spanning chunks",
        &base,
        2,
        vec![
            Step::Do(
                0,
                Op::Paste(
                    "body".into(),
                    from + 3,
                    to + 2,
                    vec!["<one>".into(), "<two>".into(), "<three>".into()],
                ),
            ),
            Step::Do(1, insert("body", to + 1, "<mid>")),
            Step::Sync(0, 1),
            Step::Sync(1, 0),
            Step::Do(1, Op::Undo),
        ],
    );
}

#[test]
fn tables_rows_cells_and_deletion() {
    let base = mixed();
    script(
        "tables",
        &base,
        2,
        vec![
            Step::Do(0, Op::Row("body".into())),
            Step::Do(1, insert("body:t0:r0c0", 0, "<cell>")),
            Step::Sync(0, 1),
            Step::Sync(1, 0),
            Step::Do(0, Op::DeleteTable("body".into())),
            Step::Do(1, insert("body:t0:r0c1", 1, "<late>")),
            Step::Sync(0, 1),
            Step::Sync(1, 0),
            Step::Do(0, Op::Undo),
        ],
    );
}

#[test]
fn tracked_changes_accept_reject_and_suggest() {
    let base = mixed();
    let (_, start, mark) = para(&base, 1);
    script(
        "tracked changes",
        &base,
        2,
        vec![
            Step::Do(0, Op::Accept("body".into(), start, mark)),
            Step::Do(1, Op::Reject("body".into(), start, mark)),
            Step::Do(1, Op::Suggest("body".into(), start + 1, "<sug>".into())),
            Step::Sync(0, 1),
            Step::Sync(1, 0),
            Step::Do(0, Op::Accept("body".into(), start, mark + 6)),
            Step::Sync(0, 1),
        ],
    );
}

#[test]
fn a_comment_over_source_text_copies_no_chunk() {
    let base = mixed();
    let (_, from, _) = para(&base, 1);
    let (_, _, to) = para(&base, 4);
    let worlds = script(
        "comment across chunks",
        &base,
        2,
        vec![
            Step::Do(0, Op::Comment("body".into(), from + 2, to)),
            Step::Sync(0, 1),
        ],
    );
    assert!(
        copies_in(&worlds[1].room(&base)).is_empty(),
        "a comment's anchors name source items without copying them"
    );
    let (_, start, _) = para(&base, 2);
    script(
        "comment then edits",
        &base,
        2,
        vec![
            Step::Do(0, Op::Comment("body".into(), from + 2, to)),
            Step::Do(1, insert("body", start + 1, "<in>")),
            Step::Sync(0, 1),
            Step::Sync(1, 0),
            Step::Do(1, Op::Uncomment("701:0".into())),
            Step::Sync(1, 0),
        ],
    );
}

#[test]
fn fields_toc_update_and_typing_in_its_result() {
    let base = mixed();
    let (_, toc, _) = para(&base, 4);
    script(
        "fields",
        &base,
        2,
        vec![
            Step::Do(1, insert("body", toc + 3, "<toc>")),
            Step::Do(0, Op::Toc),
            Step::Sync(0, 1),
            Step::Sync(1, 0),
            Step::Do(1, Op::Undo),
        ],
    );
}

#[test]
fn list_numbering_section_and_page_breaks() {
    let base = mixed();
    let (_, start, _) = para(&base, 2);
    let (_, _, section_end) = para(&base, 7);
    let (_, leading, _) = para(&base, 9);
    script(
        "lists, sections, breaks",
        &base,
        2,
        vec![
            Step::Do(0, Op::List("body".into(), 5)),
            Step::Do(1, Op::Split("body".into(), start + 3)),
            Step::Do(0, Op::SectionBreak("body".into(), section_end)),
            Step::Do(1, Op::PageBreak("body".into(), leading + 2)),
            Step::Do(0, Op::Merge("body".into(), 10, false)),
            Step::Sync(0, 1),
            Step::Sync(1, 0),
            Step::Do(1, Op::Align("body".into(), 2)),
        ],
    );
}

#[test]
fn concurrent_mid_splits_rename_paragraph_ids_alike() {
    let base = mixed();
    let (_, start, _) = para(&base, 1);
    script(
        "paraIds",
        &base,
        3,
        vec![
            Step::Do(0, Op::Split("body".into(), start + 2)),
            Step::Do(1, Op::Split("body".into(), start + 5)),
            Step::Do(2, Op::Split("body".into(), start + 9)),
            Step::Sync(0, 1),
            Step::Sync(2, 1),
            Step::Sync(1, 0),
            Step::Sync(1, 2),
        ],
    );
}

#[test]
fn duplicated_source_paragraph_ids_rename_alike() {
    let base = chunked_fixtures::xml_fixtures()
        .into_iter()
        .find(|(name, _)| name == "xml-collapsed-bookmark")
        .unwrap()
        .1;
    let (_, start, _) = para(&base, 0);
    script(
        "duplicate source paraIds",
        &base,
        2,
        vec![
            Step::Do(0, Op::Split("body".into(), start + 1)),
            Step::Sync(0, 1),
            Step::Sync(1, 0),
        ],
    );
}

/// The publication rebase reads whole stories off its states: in the
/// override layout it runs on materialized states (source plus room), and
/// lands the same edits as today's.
#[test]
fn rebase_on_materialized_chunked_states_lands_as_today() {
    let base = mixed();
    let (_, start, mark) = para(&base, 1);
    let (_, later, _) = para(&base, 3);
    let captured_ops = [
        insert("body", start + 2, "<cap>"),
        Op::Split("body".into(), mark),
        Op::Comment("body".into(), start, later),
    ];
    let latest_ops = [
        insert("body", later + 1, "<late>"),
        Op::Bold("body".into(), start, start + 4),
    ];
    let peers = [open_peer(&base, 701, false), open_peer(&base, 701, true)];
    for op in &captured_ops {
        for peer in &peers {
            run(peer, op).unwrap();
        }
    }
    // The rooms: today's whole state; the override layout's meta seed,
    // copies and edits.
    let room_of = |world: usize, peer: &Peer| -> Vec<u8> {
        if world == 0 {
            return peer.doc.encode_state_as_update_v1();
        }
        let meta = overlay::meta_seed(&overlay::fingerprint(&base));
        let mut updates = vec![yrs::Update::decode_v1(&meta).unwrap()];
        updates.extend(
            peer.outbox
                .borrow()
                .iter()
                .map(|message| yrs::Update::decode_v1(message).unwrap()),
        );
        yrs::updates::encoder::Encode::encode_v1(&yrs::Update::merge_updates(updates))
    };
    let captured: Vec<Vec<u8>> = peers
        .iter()
        .enumerate()
        .map(|(world, peer)| room_of(world, peer))
        .collect();
    for op in &latest_ops {
        for peer in &peers {
            run(peer, op).unwrap();
        }
    }
    let latest: Vec<Vec<u8>> = peers
        .iter()
        .enumerate()
        .map(|(world, peer)| room_of(world, peer))
        .collect();
    let sha = office_service::sha256_hex(&base);
    let determinism = Determinism {
        seed: "0000000000000000000000000000000000000000000000000000000000000000",
        now: "2026-10-08T00:00:00.000Z",
    };
    let exported =
        office_service::export(&base, checkpoint(&sha, &captured[0]), determinism).unwrap();
    let export_sha = office_service::sha256_hex(&exported);
    let rebased: Vec<_> = (0..2)
        .map(|world| {
            office_service::rebase(
                &base,
                checkpoint(&sha, &captured[world]),
                checkpoint(&sha, &latest[world]),
                &exported,
            )
            .map_err(|error| error.to_string())
        })
        .collect();
    let landed: Vec<_> = rebased
        .iter()
        .map(|result| {
            result.as_ref().map(|rebased| {
                office_service::export(
                    &exported,
                    checkpoint(&export_sha, &rebased.state),
                    determinism,
                )
                .unwrap()
            })
        })
        .collect();
    assert!(
        rebased[0].is_ok(),
        "today's rebase: {:?}",
        rebased[0].as_ref().err()
    );
    assert!(
        landed[0] == landed[1],
        "rebased exports differ: chunked {:?}",
        rebased[1].as_ref().err()
    );
    let (today, chunked) = (rebased[0].as_ref().unwrap(), rebased[1].as_ref().unwrap());
    assert_eq!(
        format!("{:?}", today.effects),
        format!("{:?}", chunked.effects).replace("", ""),
        "rebase effects differ"
    );
    assert!(
        overlay::is_chunked_state(&chunked.state),
        "the rebased room is in the override layout"
    );
    eprintln!(
        "rebased room: today {} B, override layout {} B",
        today.state.len(),
        chunked.state.len()
    );
}

// D3 (spike measurement): yrs's formatting cleanup `on` (every replica and
// the room clean for themselves, as production does today), `off` (nobody
// cleans) or `sent` (each replica sends the deletions its cleanup made; the
// room does not clean). `CHUNKED_CLEANUP_MODE`; `off` also needs
// `DOCX_EDIT_FORMAT_CLEANUP=0` for the sessions. Results go to `D3_OUT` as
// JSON lines; run the `d3_` tests with `--ignored --test-threads=1`.

#[derive(Clone, Copy, PartialEq, Debug)]
enum CleanupMode {
    On,
    Off,
    Sent,
}

fn cleanup_mode() -> CleanupMode {
    let mode = match std::env::var("CHUNKED_CLEANUP_MODE").as_deref() {
        Ok("on") => CleanupMode::On,
        Ok("off") => CleanupMode::Off,
        _ => CleanupMode::Sent,
    };
    let sessions_clean = !std::env::var("DOCX_EDIT_FORMAT_CLEANUP").is_ok_and(|value| value == "0");
    if std::env::var("CHUNKED_CLEANUP_MODE").is_ok() {
        assert_eq!(
            mode == CleanupMode::Off,
            !sessions_clean,
            "CHUNKED_CLEANUP_MODE=off goes with DOCX_EDIT_FORMAT_CLEANUP=0"
        );
    }
    mode
}

/// An update without structs: a deletion only.
fn blockless(update: &[u8]) -> bool {
    yrs::Update::decode_v1(update).is_ok_and(|update| update.blocks().next().is_none())
}

struct Counting;
static LIVE: std::sync::atomic::AtomicIsize = std::sync::atomic::AtomicIsize::new(0);
unsafe impl std::alloc::GlobalAlloc for Counting {
    unsafe fn alloc(&self, layout: std::alloc::Layout) -> *mut u8 {
        LIVE.fetch_add(layout.size() as isize, std::sync::atomic::Ordering::Relaxed);
        unsafe { std::alloc::System.alloc(layout) }
    }
    unsafe fn dealloc(&self, ptr: *mut u8, layout: std::alloc::Layout) {
        LIVE.fetch_sub(layout.size() as isize, std::sync::atomic::Ordering::Relaxed);
        unsafe { std::alloc::System.dealloc(ptr, layout) }
    }
    unsafe fn realloc(&self, ptr: *mut u8, layout: std::alloc::Layout, new_size: usize) -> *mut u8 {
        LIVE.fetch_add(
            new_size as isize - layout.size() as isize,
            std::sync::atomic::Ordering::Relaxed,
        );
        unsafe { std::alloc::System.realloc(ptr, layout, new_size) }
    }
}
#[global_allocator]
static ALLOCATOR: Counting = Counting;

/// The server's room: a yrs document applying every logged message in its
/// own transaction, cleaning only in the `on` mode.
struct RoomTrack {
    doc: yrs::Doc,
    applied: usize,
    regular_us: Vec<f64>,
    cleanup_us: Vec<f64>,
    saves: Vec<usize>,
    vector: yrs::StateVector,
}

impl RoomTrack {
    fn new(seed: &[u8], clean: bool) -> Self {
        let mut options = yrs::Options::with_client_id(yrs::ClientID::new(9));
        options.offset_kind = yrs::OffsetKind::Utf16;
        options.cleanup_formatting = clean;
        let doc = yrs::Doc::with_options(options);
        doc.transact_mut()
            .apply_update(yrs::Update::decode_v1(seed).unwrap())
            .unwrap();
        let vector = doc.transact().state_vector();
        Self {
            doc,
            applied: 0,
            regular_us: Vec::new(),
            cleanup_us: Vec::new(),
            saves: Vec::new(),
            vector,
        }
    }

    fn catch_up(&mut self, world: &World) {
        while self.applied < world.log.len() {
            let message = &world.log[self.applied].1;
            let cleanup = world.system[self.applied] && blockless(message);
            let started = std::time::Instant::now();
            self.doc
                .transact_mut()
                .apply_update(yrs::Update::decode_v1(message).unwrap())
                .unwrap();
            let micros = started.elapsed().as_secs_f64() * 1e6;
            if cleanup {
                self.cleanup_us.push(micros);
            } else {
                self.regular_us.push(micros);
            }
            self.applied += 1;
        }
    }

    /// The change a save stores: the room's diff against the last save.
    fn save(&mut self) {
        let txn = self.doc.transact();
        self.saves
            .push(txn.encode_state_as_update_v1(&self.vector).len());
        self.vector = txn.state_vector();
    }

    fn state(&self) -> Vec<u8> {
        self.doc
            .transact()
            .encode_state_as_update_v1(&yrs::StateVector::default())
    }
}

/// Live format items over every story, and how many of them change nothing:
/// a value the attribute already has, or one overridden before any content.
fn format_items(doc: &yrs::Doc) -> (usize, usize) {
    use yrs::block::ItemContent;
    use yrs::types::Attrs;
    let txn = doc.transact();
    let Some(stories) = txn.get_map("stories") else {
        return (0, 0);
    };
    let (mut live, mut redundant) = (0, 0);
    for (_, story) in yrs::Map::iter(&stories, &txn) {
        let yrs::Out::YText(story) = story else {
            continue;
        };
        let mut current: Attrs = Attrs::new();
        let mut since_content: HashMap<std::sync::Arc<str>, usize> = HashMap::new();
        let mut item = <yrs::TextRef as AsRef<yrs::branch::Branch>>::as_ref(&story).start();
        while let Some(at) = item {
            item = at.right();
            if at.is_deleted() {
                continue;
            }
            match at.content() {
                ItemContent::Format(key, value) => {
                    live += 1;
                    let now = current.get(key).cloned().unwrap_or(Any::Null);
                    if now == **value || since_content.contains_key(key) {
                        // A no-op, or the earlier marker for the key is
                        // overridden before any content.
                        redundant += 1;
                    }
                    *since_content.entry(key.clone()).or_default() += 1;
                    if **value == Any::Null {
                        current.remove(key);
                    } else {
                        current.insert(key.clone(), (**value).clone());
                    }
                }
                _ => since_content.clear(),
            }
        }
    }
    (live, redundant)
}

/// `word/document.xml`: its size, its runs, and adjacent text runs with the
/// same properties (a split a clean document would not have).
fn xml_runs(docx: &[u8]) -> (usize, usize, usize) {
    let parts = ooxml_opc::unzip_parts(docx).unwrap();
    let xml = parts
        .iter()
        .find(|(name, _)| name == "word/document.xml")
        .map(|(_, bytes)| String::from_utf8_lossy(bytes).into_owned())
        .unwrap_or_default();
    let (mut runs, mut mergeable) = (0, 0);
    let mut previous: Option<(usize, String, bool)> = None;
    let mut at = 0;
    while let Some(offset) = xml[at..].find("<w:r") {
        let start = at + offset;
        let next = xml.as_bytes().get(start + 4).copied();
        if !matches!(next, Some(b'>') | Some(b' ')) {
            at = start + 4;
            continue;
        }
        let Some(length) = xml[start..].find("</w:r>") else {
            break;
        };
        let end = start + length + 6;
        let body = &xml[start..end];
        let properties = body
            .find("<w:rPr>")
            .and_then(|open| {
                body[open..]
                    .find("</w:rPr>")
                    .map(|close| body[open..open + close].to_owned())
            })
            .unwrap_or_default();
        let others = [
            "<w:drawing",
            "<w:fldChar",
            "<w:instrText",
            "<w:br",
            "<w:tab",
            "<w:sym",
            "<w:footnoteReference",
            "<w:endnoteReference",
            "<w:commentReference",
            "<w:object",
            "<w:pict",
        ];
        let text_only = body.contains("<w:t") && others.iter().all(|tag| !body.contains(tag));
        if let Some((previous_end, previous_properties, previous_text)) = &previous
            && *previous_end == start
            && *previous_properties == properties
            && *previous_text
            && text_only
        {
            mergeable += 1;
        }
        runs += 1;
        previous = Some((end, properties, text_only));
        at = end;
    }
    (xml.len(), runs, mergeable)
}

fn median_of(mut values: Vec<f64>) -> f64 {
    if values.is_empty() {
        return 0.0;
    }
    values.sort_by(|a, b| a.partial_cmp(b).unwrap());
    values[values.len() / 2]
}

/// Everything D3 compares, for a world and its room after a full sync.
fn d3_metrics(base: &[u8], world: &World, room: &RoomTrack) -> serde_json::Value {
    let state = room.state();
    let sha = office_service::sha256_hex(base);
    let determinism = Determinism {
        seed: "0000000000000000000000000000000000000000000000000000000000000000",
        now: "2026-10-08T00:00:00.000Z",
    };
    let heap = {
        let before = LIVE.load(std::sync::atomic::Ordering::Relaxed);
        let doc = yrs::Doc::new();
        doc.transact_mut()
            .apply_update(yrs::Update::decode_v1(&state).unwrap())
            .unwrap();
        let heap = LIVE.load(std::sync::atomic::Ordering::Relaxed) - before;
        drop(doc);
        heap
    };
    let export = office_service::export(base, checkpoint(&sha, &state), determinism)
        .map_err(|error| error.to_string());
    let (xml_bytes, runs, mergeable) = export.as_ref().map_or((0, 0, 0), |bytes| xml_runs(bytes));
    let baseline = office_service::baseline(base, checkpoint(&sha, &state))
        .map(|entries| {
            (
                entries.len(),
                office_service::sha256_hex(
                    renamed_ids_hidden(&object_ids_hidden(&format!("{entries:?}"))).as_bytes(),
                ),
            )
        })
        .unwrap_or((0, String::new()));
    let room_view = {
        let doc = EditingDoc::new(990);
        doc.load_state_v1(&state).unwrap();
        let text = renamed_ids_hidden(&dump(&doc));
        if let Ok(path) = std::env::var("D3_DUMP") {
            std::fs::write(format!("{path}.{}.txt", mode_name(cleanup_mode())), &text).unwrap();
        }
        office_service::sha256_hex(text.as_bytes())
    };
    let dumps: Vec<String> = world.peers.iter().map(|peer| dump(&peer.doc)).collect();
    let diverged = dumps.iter().filter(|other| **other != dumps[0]).count();
    let (room_live, room_redundant) = format_items(&room.doc);
    let (peer_live, peer_redundant) = format_items(world.peers[0].doc.yrs_doc());
    let cleanups: Vec<&Vec<u8>> = world
        .log
        .iter()
        .zip(&world.system)
        .filter(|((_, message), system)| **system && blockless(message))
        .map(|((_, message), _)| message)
        .collect();
    let distinct: BTreeSet<String> = cleanups
        .iter()
        .map(|message| {
            format!(
                "{:?}",
                yrs::Update::decode_v1(message).unwrap().delete_set()
            )
        })
        .collect();
    serde_json::json!({
        "roomBytes": state.len(),
        "roomHeap": heap,
        "saves": room.saves,
        "storedTotal": room.saves.iter().sum::<usize>(),
        "storedLast": room.saves.last().copied().unwrap_or(0),
        "roomFormatLive": room_live,
        "roomFormatRedundant": room_redundant,
        "peerFormatLive": peer_live,
        "peerFormatRedundant": peer_redundant,
        "exportBytes": export.as_ref().map_or(0, Vec::len),
        "exportError": export.as_ref().err(),
        "documentXmlBytes": xml_bytes,
        "runs": runs,
        "mergeableRuns": mergeable,
        "baselineEntries": baseline.0,
        "baselineHash": baseline.1,
        "roomViewHash": room_view,
        "peersDiverged": diverged,
        "messages": world.log.len(),
        "cleanupMessages": cleanups.len(),
        "cleanupBytes": cleanups.iter().map(|message| message.len()).sum::<usize>(),
        "distinctCleanups": distinct.len(),
        "echoes": world.echoes,
        "roomRegularApplyUsMedian": median_of(room.regular_us.clone()),
        "roomCleanupApplyUsMedian": median_of(room.cleanup_us.clone()),
        "roomCleanupApplyUsTotal": room.cleanup_us.iter().sum::<f64>(),
        "roomRegularApplyUsTotal": room.regular_us.iter().sum::<f64>(),
    })
}

fn d3_write(line: serde_json::Value) {
    use std::io::Write;
    let path = std::env::var("D3_OUT").expect("D3_OUT names the output file");
    let mut out = std::fs::OpenOptions::new()
        .create(true)
        .append(true)
        .open(path)
        .unwrap();
    writeln!(out, "{line}").unwrap();
}

fn mode_name(mode: CleanupMode) -> &'static str {
    match mode {
        CleanupMode::On => "on",
        CleanupMode::Off => "off",
        CleanupMode::Sent => "sent",
    }
}

fn seed_state_of(base: &[u8]) -> Vec<u8> {
    let doc = EditingDoc::new(1);
    seed_from_docx(&doc, base).unwrap();
    doc.encode_state_as_update_v1()
}

/// The lockstep's random schedule in today's layout only, with the room
/// saved at every round's end.
fn d3_schedule(name: &str, base: &[u8], seed: u64, peers: usize, rounds: usize, ops: usize) {
    let mode = cleanup_mode();
    let mut rng = Rng(seed.wrapping_mul(0x9E37_79B9_7F4A_7C15) | 1);
    let mut world = World::new(base, peers, false);
    let mut room = RoomTrack::new(&seed_state_of(base), mode == CleanupMode::On);
    let focus = layout(&world.peers[0].doc, "body")
        .1
        .get(1)
        .map(|paragraph| paragraph.0.clone());
    let mut applied = 0;
    let trace = std::env::var("CHUNKED_TRACE").ok() == Some(seed.to_string());
    for round in 0..rounds {
        for step in 0..ops {
            for peer in 0..peers {
                let marker = format!("<{peer}.{round}.{step}>");
                let focused = (round == 0 && step == 0)
                    .then_some(focus.as_deref())
                    .flatten();
                let op = choose(&mut rng, &world.peers[peer].doc, focused, marker);
                let outcome = attempt(&world.peers[peer], &op);
                if trace {
                    eprintln!("r{round} s{step} p{peer} {op:?} -> {outcome:?}");
                }
                applied += usize::from(outcome.is_ok());
                world.drain(peer);
            }
        }
        let pairs: Vec<(usize, usize)> = (0..peers * 2)
            .map(|_| (rng.below(peers), rng.below(peers)))
            .filter(|(from, to)| from != to)
            .collect();
        for (from, to) in pairs {
            world.sync(from, to).unwrap();
            if trace {
                eprintln!("r{round} sync {from}->{to}");
            }
        }
        room.catch_up(&world);
        room.save();
    }
    world.sync_all();
    room.catch_up(&world);
    room.save();
    let mut line = d3_metrics(base, &world, &room);
    let object = line.as_object_mut().unwrap();
    object.insert("kind".into(), "schedule".into());
    object.insert("fixture".into(), name.into());
    object.insert("seed".into(), seed.into());
    object.insert("peers".into(), peers.into());
    object.insert("mode".into(), mode_name(mode).into());
    object.insert("applied".into(), applied.into());
    d3_write(line);
}

#[test]
#[ignore = "D3 measurement: CHUNKED_CLEANUP_MODE, D3_OUT"]
fn d3_random_schedules() {
    let seeds: u64 = std::env::var("CHUNKED_SEEDS")
        .ok()
        .and_then(|value| value.parse().ok())
        .unwrap_or(6);
    let only: Option<u64> = std::env::var("CHUNKED_SEED")
        .ok()
        .and_then(|value| value.parse().ok());
    let fixture = std::env::var("CHUNKED_FIXTURE").ok();
    for (name, base) in fixtures() {
        if fixture.as_ref().is_some_and(|only| *only != name) {
            continue;
        }
        for seed in 0..seeds {
            for peers in [2, 3] {
                let seed = seed * 31 + peers as u64;
                if only.is_some_and(|only| only != seed) {
                    continue;
                }
                let outcome = std::panic::catch_unwind(AssertUnwindSafe(|| {
                    d3_schedule(&name, &base, seed, peers, 4, 3)
                }));
                if let Err(panic) = outcome {
                    let message = panic
                        .downcast_ref::<String>()
                        .cloned()
                        .or_else(|| panic.downcast_ref::<&str>().map(|text| (*text).to_owned()))
                        .unwrap_or_default();
                    d3_write(serde_json::json!({
                        "kind": "schedule", "fixture": name, "seed": seed, "peers": peers,
                        "mode": mode_name(cleanup_mode()), "panicked": message,
                    }));
                }
            }
        }
    }
}

/// The paragraph `index` of `peer`'s body: its first unit and its mark.
fn d3_para(peer: &Peer, index: usize) -> (u32, u32) {
    let (_, paragraphs) = layout(&peer.doc, "body");
    let (_, start, mark) = paragraphs[index.min(paragraphs.len() - 1)].clone();
    (start, mark)
}

fn d3_op(world: &mut World, peer: usize, op: Op) {
    let _ = attempt(&world.peers[peer], &op);
    world.drain(peer);
}

/// The bold-then-italic span of a round: the last 20 units of a paragraph
/// and the first 10 of the next.
fn d3_span(peer: &Peer, paragraph: usize) -> (u32, u32) {
    let (start, mark) = d3_para(peer, paragraph);
    let (_, next_mark) = d3_para(peer, paragraph + 1);
    (
        mark.saturating_sub(20).max(start),
        (mark + 10).min(next_mark),
    )
}

/// Three peers typing, formatting, pasting and undoing on a long file, with
/// one select-all delete and its Undo; the room saved every round.
fn d3_session(name: &str, base: &[u8], rounds: usize) {
    let mode = cleanup_mode();
    let mut world = World::new(base, 3, false);
    let mut room = RoomTrack::new(&seed_state_of(base), mode == CleanupMode::On);
    let mut timeline = Vec::new();
    for round in 0..rounds {
        let paragraph = 4 + 3 * round;
        // Peer 0 types a word at the end of a paragraph, key by key.
        for character in format!(" word{round}").chars() {
            let (_, mark) = d3_para(&world.peers[0], paragraph);
            d3_op(
                &mut world,
                0,
                Op::Insert("body".into(), mark, character.to_string()),
            );
        }
        // Peer 1 bolds across the paragraph's end, then italicizes inside.
        let (from, to) = d3_span(&world.peers[1], paragraph);
        d3_op(&mut world, 1, Op::Bold("body".into(), from, to));
        d3_op(
            &mut world,
            1,
            Op::Italic("body".into(), from + 5, to.saturating_sub(5).max(from + 6)),
        );
        // Peer 2 pastes three lines over a few characters further down.
        let (start, mark) = d3_para(&world.peers[2], paragraph + 10);
        d3_op(
            &mut world,
            2,
            Op::Paste(
                "body".into(),
                start,
                (start + 5).min(mark),
                vec![
                    format!("pasted {round} a"),
                    format!("pasted {round} b"),
                    format!("pasted {round} c"),
                ],
            ),
        );
        for (from, to) in [(0, 1), (1, 2), (2, 0)] {
            world.sync(from, to).unwrap();
        }
        // Peer 0 undoes three keys; peer 1 takes the bold off again.
        for _ in 0..3 {
            d3_op(&mut world, 0, Op::Undo);
        }
        let (from, to) = d3_span(&world.peers[1], paragraph);
        d3_op(&mut world, 1, Op::Bold("body".into(), from, to));
        if round == rounds / 2 && !std::env::var("D3_SELECT_ALL").is_ok_and(|value| value == "0") {
            // Peer 2 selects all and deletes while peer 0 types; then Undo.
            d3_op(&mut world, 2, Op::SelectAll("body".into()));
            let (_, mark) = d3_para(&world.peers[0], paragraph + 2);
            d3_op(
                &mut world,
                0,
                Op::Insert("body".into(), mark, "late".into()),
            );
            for (from, to) in [(2, 0), (0, 2), (2, 1), (0, 1)] {
                world.sync(from, to).unwrap();
            }
            d3_op(&mut world, 2, Op::Undo);
        }
        for (from, to) in [(0, 1), (1, 0), (2, 0), (0, 2), (1, 2)] {
            world.sync(from, to).unwrap();
        }
        room.catch_up(&world);
        room.save();
        let (live, redundant) = format_items(&room.doc);
        timeline.push(serde_json::json!([live, redundant]));
    }
    world.sync_all();
    room.catch_up(&world);
    room.save();
    let mut line = d3_metrics(base, &world, &room);
    let object = line.as_object_mut().unwrap();
    object.insert(
        "kind".into(),
        if std::env::var("D3_SELECT_ALL").is_ok_and(|value| value == "0") {
            "session-without-select-all"
        } else {
            "session"
        }
        .into(),
    );
    object.insert("fixture".into(), name.into());
    object.insert("mode".into(), mode_name(mode).into());
    object.insert("roomFormatTimeline".into(), timeline.into());
    d3_write(line);
}

#[test]
#[ignore = "D3 measurement: CHUNKED_CLEANUP_MODE, D3_OUT, D3_FILES"]
fn d3_sessions_on_long_files() {
    let files = std::env::var("D3_FILES").expect("D3_FILES lists .docx paths, separated by ;");
    let rounds: usize = std::env::var("D3_ROUNDS")
        .ok()
        .and_then(|value| value.parse().ok())
        .unwrap_or(8);
    for path in files.split(';') {
        let base = std::fs::read(path).unwrap();
        let name = Path::new(path)
            .file_stem()
            .unwrap()
            .to_string_lossy()
            .into_owned();
        d3_session(&name, &base, rounds);
    }
}

/// The worst case: two peers turn bold on and off over overlapping words at
/// the same time, again and again (one peer alone leaves nothing redundant:
/// a local format deletes the markers inside its range), one of them also
/// typing inside; both synced after every step.
#[test]
#[ignore = "D3 measurement: CHUNKED_CLEANUP_MODE, D3_OUT"]
fn d3_bold_toggled_over_one_range() {
    let mode = cleanup_mode();
    let base =
        std::fs::read(Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/stories.docx"))
            .unwrap();
    let mut world = World::new(&base, 2, false);
    let mut room = RoomTrack::new(&seed_state_of(&base), mode == CleanupMode::On);
    let toggles: usize = std::env::var("D3_TOGGLES")
        .ok()
        .and_then(|value| value.parse().ok())
        .unwrap_or(60);
    let mut timeline = Vec::new();
    for toggle in 1..=toggles {
        let (start, mark) = d3_para(&world.peers[0], 1);
        d3_op(
            &mut world,
            0,
            Op::Bold("body".into(), start, mark.min(start + 12).max(start + 1)),
        );
        let (start, mark) = d3_para(&world.peers[1], 1);
        d3_op(
            &mut world,
            1,
            Op::Bold(
                "body".into(),
                (start + 4).min(mark),
                mark.min(start + 16).max(start + 5),
            ),
        );
        if toggle % 5 == 0 {
            d3_op(
                &mut world,
                1,
                Op::Insert("body".into(), start + 6, "x".into()),
            );
        }
        world.sync(0, 1).unwrap();
        world.sync(1, 0).unwrap();
        room.catch_up(&world);
        if [10, 20, 40, 60, 100, 200].contains(&toggle) || toggle == toggles {
            room.save();
            let (room_live, room_redundant) = format_items(&room.doc);
            let (peer_live, peer_redundant) = format_items(world.peers[0].doc.yrs_doc());
            timeline.push(serde_json::json!({
                "toggles": toggle,
                "roomLive": room_live,
                "roomRedundant": room_redundant,
                "peerLive": peer_live,
                "peerRedundant": peer_redundant,
                "roomBytes": room.state().len(),
            }));
        }
    }
    world.sync_all();
    room.catch_up(&world);
    room.save();
    let mut line = d3_metrics(&base, &world, &room);
    let object = line.as_object_mut().unwrap();
    object.insert("kind".into(), "toggle".into());
    object.insert("fixture".into(), "stories".into());
    object.insert("mode".into(), mode_name(mode).into());
    object.insert("timeline".into(), timeline.into());
    d3_write(line);
}

/// A worst case that leaves markers behind: a peer bolds a fresh word while
/// another peer deletes it, again and again; each round ends synced.
#[test]
#[ignore = "D3 measurement: CHUNKED_CLEANUP_MODE, D3_OUT"]
fn d3_bold_against_a_concurrent_delete() {
    let mode = cleanup_mode();
    let base =
        std::fs::read(Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/stories.docx"))
            .unwrap();
    let mut world = World::new(&base, 2, false);
    let mut room = RoomTrack::new(&seed_state_of(&base), mode == CleanupMode::On);
    let rounds: usize = std::env::var("D3_TOGGLES")
        .ok()
        .and_then(|value| value.parse().ok())
        .unwrap_or(60);
    let mut timeline = Vec::new();
    for round in 1..=rounds {
        let word = format!("w{round:03}");
        let (start, _) = d3_para(&world.peers[0], 1);
        d3_op(
            &mut world,
            0,
            Op::Insert("body".into(), start, word.clone()),
        );
        world.sync(0, 1).unwrap();
        let length = word.encode_utf16().count() as u32;
        d3_op(
            &mut world,
            0,
            Op::Bold("body".into(), start, start + length),
        );
        d3_op(
            &mut world,
            1,
            Op::Delete("body".into(), start, start + length),
        );
        world.sync(0, 1).unwrap();
        world.sync(1, 0).unwrap();
        world.sync(0, 1).unwrap();
        room.catch_up(&world);
        if [10, 20, 40, 60, 100, 200].contains(&round) || round == rounds {
            room.save();
            let (room_live, room_redundant) = format_items(&room.doc);
            let (peer_live, peer_redundant) = format_items(world.peers[0].doc.yrs_doc());
            let (other_live, other_redundant) = format_items(world.peers[1].doc.yrs_doc());
            timeline.push(serde_json::json!({
                "rounds": round,
                "roomLive": room_live,
                "roomRedundant": room_redundant,
                "peerLive": peer_live,
                "peerRedundant": peer_redundant,
                "otherLive": other_live,
                "otherRedundant": other_redundant,
                "roomBytes": room.state().len(),
            }));
        }
    }
    world.sync_all();
    room.catch_up(&world);
    room.save();
    let mut line = d3_metrics(&base, &world, &room);
    let object = line.as_object_mut().unwrap();
    object.insert("kind".into(), "format-vs-delete".into());
    object.insert("fixture".into(), "stories".into());
    object.insert("mode".into(), mode_name(mode).into());
    object.insert("timeline".into(), timeline.into());
    d3_write(line);
}
