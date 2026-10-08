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
    seen: Vec<BTreeSet<usize>>,
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
    let subscription = doc
        .observe_shared_updates(true, move |update| sink.borrow_mut().push(update))
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
            seen: vec![BTreeSet::new(); peers],
        }
    }

    fn drain(&mut self, peer: usize) {
        let sent: Vec<_> = self.peers[peer].outbox.borrow_mut().drain(..).collect();
        for message in sent {
            self.seen[peer].insert(self.log.len());
            self.log.push((peer, message));
        }
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
            self.drain(to);
        }
        Ok(())
    }

    fn sync_all(&mut self) {
        loop {
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
    Merge(String, bool),
    Bold(String, u32, u32),
    Align(String),
    List(String),
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
        12 | 13 => Op::Merge(pick(rng).0, rng.chance(50)),
        14 => {
            let first = rng.below(paragraphs.len());
            let last = (first + rng.below(3)).min(paragraphs.len() - 1);
            let from = at(rng, &paragraphs[first]);
            let to = at(rng, &paragraphs[last]).max(from + 1).min(len);
            Op::Bold(story, from, to)
        }
        15 => Op::Align(pick(rng).0),
        16 => Op::List(pick(rng).0),
        17 => {
            let first = rng.below(paragraphs.len());
            let last = (first + rng.below(3)).min(paragraphs.len() - 1);
            let from = at(rng, &paragraphs[first]);
            let to = at(rng, &paragraphs[last]).max(from).min(len - 1);
            Op::Paste(
                story,
                from,
                to,
                vec![format!("{marker}a"), format!("{marker}b"), format!("{marker}c")],
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
            .split_paragraph(&ctx(), Position::new(story, *at), None)
            .map(drop)
            .map_err(|error| error.to_string()),
        Op::Merge(para, forward) => doc
            .merge_paragraphs(
                &ctx(),
                para,
                if *forward {
                    MergeDirection::Forward
                } else {
                    MergeDirection::Backward
                },
            )
            .map(drop)
            .map_err(|error| error.to_string()),
        Op::Bold(story, from, to) => doc
            .toggle_format(&ctx(), StoryRange::new(story, *from, *to), SimpleFormat::Bold)
            .map(drop)
            .map_err(|error| error.to_string()),
        Op::Align(para) => doc
            .set_paragraph_attrs(
                &ctx(),
                &ParaSelector::One(para.clone()),
                &ParaAttrDelta {
                    alignment: Patch::Set("center".into()),
                    ..Default::default()
                },
            )
            .map(drop)
            .map_err(|error| error.to_string()),
        Op::List(para) => doc
            .set_paragraph_attrs(
                &ctx(),
                &ParaSelector::One(para.clone()),
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
            .map_err(|error| error.to_string()),
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
                doc.split_paragraph(&ctx(), Position::new(story, caret), None)
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
                Any::from("a comment"),
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
    let (Ok(left), Ok(right)) = (
        ooxml_opc::unzip_parts(left),
        ooxml_opc::unzip_parts(right),
    ) else {
        return false;
    };
    left.len() == right.len()
        && left.iter().zip(&right).all(|((a, x), (b, y))| a == b && (x == y || mask(x) == mask(y)))
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
    refused: usize,
    panicked: usize,
    room_seeded: usize,
    room_chunked: usize,
    copies: usize,
}

fn lockstep(name: &str, base: &[u8], seed: u64, peers: usize, rounds: usize, ops: usize) -> Stats {
    let trace = std::env::var("CHUNKED_TRACE").ok() == Some(seed.to_string());
    let mut rng = Rng(seed.wrapping_mul(0x9E37_79B9_7F4A_7C15) | 1);
    let mut worlds = [World::new(base, peers, false), World::new(base, peers, true)];
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
                let focused = (round == 0 && step == 0).then_some(focus.as_deref()).flatten();
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
                let (left, right) = (
                    dump(&worlds[0].peers[peer].doc),
                    dump(&worlds[1].peers[peer].doc),
                );
                assert!(
                    left == right,
                    "{}",
                    context(
                        round,
                        &format!("peer {peer} after {op:?}: {}", first_difference(&left, &right))
                    )
                );
            }
        }
        let pairs: Vec<(usize, usize)> = (0..peers * 2)
            .map(|_| (rng.below(peers), rng.below(peers)))
            .filter(|(from, to)| from != to)
            .collect();
        for (from, to) in pairs {
            for world in worlds.iter_mut() {
                let chunked = world.chunked;
                world.sync(from, to).unwrap_or_else(|error| {
                    panic!(
                        "{}",
                        context(round, &format!("chunked {chunked} sync {from}->{to}: {error}"))
                    )
                });
            }
            if trace {
                eprintln!("r{round} sync {from}->{to}");
            }
            let (left, right) = (
                dump(&worlds[0].peers[to].doc),
                dump(&worlds[1].peers[to].doc),
            );
            assert!(
                left == right,
                "{}",
                context(
                    round,
                    &format!("sync {from}->{to}: {}", first_difference(&left, &right))
                )
            );
        }
    }
    finish(&format!("{name} seed {seed}"), base, &mut worlds, &mut stats);
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
    for world in worlds.iter_mut() {
        world.sync_all();
        let first = dump(&world.peers[0].doc);
        for (index, peer) in world.peers.iter().enumerate().skip(1) {
            let other = dump(&peer.doc);
            if first != other {
                if let Ok(dir) = std::env::var("CHUNKED_DUMP_DIR") {
                    let dir = Path::new(&dir);
                    let state = |doc: &EditingDoc| doc.encode_state_as_update_v1();
                    std::fs::write(dir.join("peer0.txt"), &first).unwrap();
                    std::fs::write(dir.join(format!("peer{index}.txt")), &other).unwrap();
                    std::fs::write(dir.join("peer0.bin"), state(&world.peers[0].doc)).unwrap();
                    std::fs::write(dir.join(format!("peer{index}.bin")), state(&peer.doc)).unwrap();
                    for (at, (_, message)) in world.log.iter().enumerate() {
                        std::fs::write(dir.join(format!("msg{at:03}.bin")), message).unwrap();
                    }
                }
                panic!(
                    "{label}: chunked {} peer {index} diverges: {}",
                    world.chunked,
                    first_difference(&first, &other),
                );
            }
        }
    }
    let (left, right) = (
        dump(&worlds[0].peers[0].doc),
        dump(&worlds[1].peers[0].doc),
    );
    assert!(
        left == right,
        "{label}: converged worlds differ: {}",
        first_difference(&left, &right)
    );
    let rooms = [worlds[0].room(base), worlds[1].room(base)];
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
    stats.copies = fresh.overlay().map_or(0, |overlay| overlay.shared().len());
    stats.room_seeded = rooms[0].len();
    stats.room_chunked = rooms[1].len();
    let sha = office_service::sha256_hex(base);
    let determinism = Determinism {
        seed: "0000000000000000000000000000000000000000000000000000000000000000",
        now: "2026-10-08T00:00:00.000Z",
    };
    let exports: Vec<_> = rooms
        .iter()
        .map(|room| office_service::export(base, checkpoint(&sha, room), determinism))
        .collect();
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
    let baselines: Vec<_> = rooms
        .iter()
        .map(|room| {
            office_service::baseline(base, checkpoint(&sha, room))
                .map(|entries| renamed_ids_hidden(&object_ids_hidden(&format!("{entries:#?}"))))
                .map_err(|error| error.to_string())
        })
        .collect();
    if baselines[0] != baselines[1] {
        let text = |baseline: &Result<String, String>| baseline.clone().unwrap_or_else(|error| error);
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
    for (name, base) in fixtures() {
        for seed in 0..seeds {
            for peers in [2, 3] {
                let seed = seed * 31 + peers as u64;
                if only.is_some_and(|only| only != seed) {
                    continue;
                }
                let stats = lockstep(&name, &base, seed, peers, rounds, 3);
                eprintln!("{name} seed {seed} peers {peers}: {stats:?}");
                totals.applied += stats.applied;
                totals.refused += stats.refused;
                totals.panicked += stats.panicked;
                totals.copies += stats.copies;
                schedules += 1;
            }
        }
    }
    eprintln!("lockstep: {schedules} schedules, {totals:?}");
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
    let mut worlds = [World::new(base, peers, false), World::new(base, peers, true)];
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
                compare(&worlds, to, &format!("{name} step {index} sync {from}->{to}"));
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
    let body = overlay.stories.iter().find(|story| story.id == "body").unwrap();
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
    let (second, _, _) = para(&base, 1);
    let (third, start_of_third, _) = para(&base, 2);
    script(
        "joins across chunks",
        &base,
        2,
        vec![
            Step::Do(0, Op::Merge(third, false)),
            Step::Do(1, Op::Merge(second, true)),
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
        vec![Step::Do(0, Op::Comment("body".into(), from + 2, to)), Step::Sync(0, 1)],
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
    let (item, start, _) = para(&base, 2);
    let (_, _, section_end) = para(&base, 7);
    let (_, leading, _) = para(&base, 9);
    script(
        "lists, sections, breaks",
        &base,
        2,
        vec![
            Step::Do(0, Op::List(para(&base, 5).0)),
            Step::Do(1, Op::Split("body".into(), start + 3)),
            Step::Do(0, Op::SectionBreak("body".into(), section_end)),
            Step::Do(1, Op::PageBreak("body".into(), leading + 2)),
            Step::Do(0, Op::Merge(para(&base, 10).0, false)),
            Step::Sync(0, 1),
            Step::Sync(1, 0),
            Step::Do(1, Op::Align(item)),
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
