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
                    overlay::read_dump(&worlds[0].peers[peer].doc),
                    overlay::read_dump(&worlds[1].peers[peer].doc),
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
                overlay::read_dump(&worlds[0].peers[to].doc),
                overlay::read_dump(&worlds[1].peers[to].doc),
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
    for world in worlds.iter_mut() {
        world.sync_all();
        let first = overlay::read_dump(&world.peers[0].doc);
        for (index, peer) in world.peers.iter().enumerate().skip(1) {
            let other = overlay::read_dump(&peer.doc);
            if first != other {
                let state = |doc: &EditingDoc| {
                    let txn = doc.yrs_doc().transact();
                    format!(
                        "pending {} ds {} sv {:?}",
                        txn.store().pending_update().is_some(),
                        txn.store().pending_ds().is_some(),
                        txn.state_vector()
                    )
                };
                if let Ok(dir) = std::env::var("CHUNKED_DUMP_DIR") {
                    let dir = Path::new(&dir);
                    std::fs::write(dir.join("peer0.txt"), &first).unwrap();
                    std::fs::write(dir.join(format!("peer{index}.txt")), &other).unwrap();
                    std::fs::write(dir.join("peer0.bin"), world.peers[0].doc.encode_state_as_update_v1()).unwrap();
                    std::fs::write(dir.join(format!("peer{index}.bin")), peer.doc.encode_state_as_update_v1()).unwrap();
                    let log: Vec<String> = world.log.iter().map(|(author, message)| format!("{author} {}", message.len())).collect();
                    std::fs::write(dir.join("log.txt"), log.join("
")).unwrap();
                    for (index, (_, message)) in world.log.iter().enumerate() {
                        std::fs::write(dir.join(format!("msg{index:03}.bin")), message).unwrap();
                    }
                }
                panic!(
                    "{name} seed {seed}: chunked {} peer {index} diverges: {}
  peer 0 {}
  peer {index} {}",
                    world.chunked,
                    first_difference(&first, &other),
                    state(&world.peers[0].doc),
                    state(&peer.doc)
                );
            }
        }
    }
    let (left, right) = (
        overlay::read_dump(&worlds[0].peers[0].doc),
        overlay::read_dump(&worlds[1].peers[0].doc),
    );
    assert!(
        left == right,
        "{name} seed {seed}: converged worlds differ: {}",
        first_difference(&left, &right)
    );
    // The rooms hold everything a fresh session needs.
    let rooms = [worlds[0].room(base), worlds[1].room(base)];
    let fresh = EditingDoc::new(799);
    let envelope = parse_docx_for_edit(base).unwrap();
    overlay::open_chunked(&fresh, envelope, &overlay::fingerprint(base)).unwrap();
    fresh.apply_shared_update(&rooms[1]).unwrap();
    let reopened = overlay::read_dump(&fresh);
    assert!(
        reopened == left,
        "{name} seed {seed}: a fresh session on the chunked room differs: {}",
        first_difference(&left, &reopened)
    );
    stats.copies = fresh.overlay().map_or(0, |overlay| overlay.shared().len());
    stats.room_seeded = rooms[0].len();
    stats.room_chunked = rooms[1].len();
    // Both rooms export the same file and the same baseline.
    let sha = office_service::sha256_hex(base);
    let checkpoint = |state: &[u8]| Checkpoint {
        format: Format::Docx,
        schema_version: 1,
        base_sha256: &sha,
        state: unsafe { std::mem::transmute::<&[u8], &'static [u8]>(state) },
    };
    let determinism = Determinism {
        seed: "0000000000000000000000000000000000000000000000000000000000000000",
        now: "2026-10-08T00:00:00.000Z",
    };
    let exports: Vec<_> = rooms
        .iter()
        .map(|room| office_service::export(base, checkpoint(room), determinism))
        .collect();
    match (&exports[0], &exports[1]) {
        (Ok(left), Ok(right)) => assert!(left == right, "{name} seed {seed}: exports differ"),
        (left, right) => assert_eq!(
            left.as_ref().err().map(ToString::to_string),
            right.as_ref().err().map(ToString::to_string),
            "{name} seed {seed}: export outcome differs"
        ),
    }
    let baselines: Vec<_> = rooms
        .iter()
        .map(|room| {
            office_service::baseline(base, checkpoint(room))
                .map(|entries| object_ids_hidden(&format!("{entries:?}")))
                .map_err(|error| error.to_string())
        })
        .collect();
    if baselines[0] != baselines[1] {
        let (left, right) = (
            baselines[0].clone().unwrap_or_else(|error| error).replace("BaselineEntry", "
BaselineEntry"),
            baselines[1].clone().unwrap_or_else(|error| error).replace("BaselineEntry", "
BaselineEntry"),
        );
        panic!(
            "{name} seed {seed}: baselines differ: {}",
            first_difference(&left, &right)
        );
    }
    stats
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
