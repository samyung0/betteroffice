//! office-rebase.ts `transplant` (`landLater`) on yrs: the later edits
//! (latest − captured), observed as Yjs changes at captured positions, land
//! at the corresponding positions of seed(export).
//!
//! yrs hands a transaction's events and a map event's keys over in hash map
//! order, and integrates a map's entries in hash map order, so steps, keys
//! and copied entries are put in a fixed order (roots in lineage order, then
//! path; keys by name; entries by their item ids). Yjs used its integration
//! order. Order only decides the clocks of the items the rebase writes, not
//! what the rebased state holds.

use std::collections::{HashMap, HashSet};
use std::rc::Rc;
use std::sync::Arc;
use std::sync::{Mutex, MutexGuard};

use yrs::block::{ItemContent, Prelim};
use yrs::branch::{Branch, BranchPtr};
use yrs::types::text::YChange;
use yrs::types::{Delta, EntryChange, Event, PathSegment, TypeRef};
use yrs::updates::decoder::Decode;
use yrs::{
    Any, Array, ArrayRef, DeepObservable, Doc, Map, MapRef, OffsetKind, Options, Out, ReadTxn,
    Text, TextRef, Transact, TransactionMut, Update,
};

use super::align::{
    Alignment, Units, align, child_at_offset, past, point_at, unit_attributes, units,
};
use super::{docx, fail};
use crate::error::{Error, Result};

pub(crate) struct Lineage {
    pub maps: &'static [&'static str],
    pub arrays: &'static [&'static str],
    /// Root maps keyed by entity id: their keys, and id strings in values, are translated.
    pub keyed: &'static [&'static str],
    /// DOCX: anchors are rewritten once every text landed, `settle` and
    /// `check` run, and story deltas renumber field children.
    pub docx: bool,
}

pub(crate) const DOCX_LINEAGE: Lineage = Lineage {
    maps: &["stories", "comments", "bookmarks"],
    arrays: &[],
    keyed: &["stories", "comments", "bookmarks"],
    docx: true,
};

pub(crate) const PPTX_LINEAGE: Lineage = Lineage {
    maps: &[
        "pptx:meta",
        "pptx:slides",
        "pptx:shapes",
        "pptx:stories",
        "pptx:comments",
    ],
    arrays: &["pptx:slide-order"],
    keyed: &[
        "pptx:slides",
        "pptx:shapes",
        "pptx:stories",
        "pptx:comments",
    ],
    docx: false,
};

/// The key whose value names positions in texts, and the roots holding it.
pub(crate) const POSITIONS_KEY: &str = "anchors";
pub(crate) const POSITION_ROOTS: [&str; 2] = ["comments", "bookmarks"];

/// A document read as Yjs reads one: UTF-16 positions.
pub(crate) fn new_doc(client: Option<u64>) -> Doc {
    let mut options = Options {
        offset_kind: OffsetKind::Utf16,
        ..Options::default()
    };
    if let Some(client) = client {
        options.client_id = yrs::block::ClientID::new(client);
    }
    Doc::with_options(options)
}

/// office-rebase.ts `open`: the lineage's roots, then the update.
pub(crate) fn open(lineage: &Lineage, update: &[u8], client: Option<u64>) -> Result<Doc> {
    let doc = new_doc(client);
    for name in lineage.maps {
        doc.get_or_insert_map(*name);
    }
    for name in lineage.arrays {
        doc.get_or_insert_array(*name);
    }
    let update = Update::decode_v1(update).map_err(|error| fail(error.to_string()))?;
    let mut txn = doc.transact_mut();
    txn.apply_update(update)
        .map_err(|error| fail(error.to_string()))?;
    if txn.store().pending_update().is_some() || txn.store().pending_ds().is_some() {
        return Err(fail("a state refers to content it does not hold"));
    }
    drop(txn);
    Ok(doc)
}

pub(crate) fn root_map<T: ReadTxn>(txn: &T, name: &str) -> MapRef {
    txn.get_map(name).expect("the lineage's roots exist")
}

/// A value copied into the rebased document; maps integrate their entries
/// in the order given.
pub(crate) enum Copied {
    Any(Any),
    Text(Vec<Delta<Copied>>),
    Array(Vec<Copied>),
    Map(Vec<(Arc<str>, Copied)>),
}

impl Prelim for Copied {
    type Return = Out;

    fn into_content(self, _txn: &mut TransactionMut) -> (ItemContent, Option<Self>) {
        match self {
            Copied::Any(any) => (ItemContent::Any(vec![any]), None),
            other => {
                let type_ref = match &other {
                    Copied::Text(_) => TypeRef::Text,
                    Copied::Array(_) => TypeRef::Array,
                    _ => TypeRef::Map,
                };
                (ItemContent::Type(Branch::new(type_ref)), Some(other))
            }
        }
    }

    fn integrate(self, txn: &mut TransactionMut, inner_ref: BranchPtr) {
        match self {
            Copied::Text(delta) => TextRef::from(inner_ref).apply_delta(txn, delta),
            Copied::Array(items) => {
                let array = ArrayRef::from(inner_ref);
                for (index, item) in items.into_iter().enumerate() {
                    array.insert(txn, index as u32, item);
                }
            }
            Copied::Map(entries) => {
                let map = MapRef::from(inner_ref);
                for (key, value) in entries {
                    map.insert(txn, key, value);
                }
            }
            Copied::Any(_) => {}
        }
    }
}

impl From<Any> for Copied {
    fn from(any: Any) -> Self {
        Copied::Any(any)
    }
}

fn type_name(value: &Out) -> &'static str {
    match value {
        Out::YMap(_) => "YMap",
        Out::YArray(_) => "YArray",
        Out::YText(_) => "YText",
        Out::YXmlElement(_) => "YXmlElement",
        Out::YXmlFragment(_) => "YXmlFragment",
        Out::YXmlText(_) => "YXmlText",
        Out::YDoc(_) => "Doc",
        _ => "value",
    }
}

/// Both values are the same kind of shared type (`a?.constructor === b.constructor`).
fn same_type(a: Option<&Out>, b: &Out) -> bool {
    matches!(
        (a, b),
        (Some(Out::YMap(_)), Out::YMap(_))
            | (Some(Out::YText(_)), Out::YText(_))
            | (Some(Out::YArray(_)), Out::YArray(_))
    )
}

/// Strings in a value through `id` (`translate`).
pub(crate) fn translate(value: &Any, id: &dyn Fn(&str) -> String) -> Any {
    match value {
        Any::String(text) => Any::String(id(text).into()),
        Any::Array(items) => Any::Array(items.iter().map(|item| translate(item, id)).collect()),
        Any::Map(map) => Any::Map(Arc::new(
            map.iter()
                .map(|(key, item)| (key.clone(), translate(item, id)))
                .collect(),
        )),
        other => other.clone(),
    }
}

/// A map's keys in their entries' item order.
pub(crate) fn ordered_keys<T: ReadTxn>(txn: &T, map: &MapRef) -> Vec<String> {
    let branch: &Branch = map.as_ref();
    let items: HashMap<&str, (u64, u32)> = branch
        .map_items()
        .map(|(key, item)| (key.as_ref(), (item.id().client.get(), item.id().clock)))
        .collect();
    let mut keys: Vec<(u64, u32, String)> = map
        .keys(txn)
        .map(|key| {
            let id = items.get(key).copied().unwrap_or_default();
            (id.0, id.1, key.to_owned())
        })
        .collect();
    keys.sort();
    keys.into_iter().map(|(_, _, key)| key).collect()
}

/// office-rebase.ts `copy`; a positioned top-level entity leaves its anchors out.
pub(crate) fn copy<T: ReadTxn>(
    txn: &T,
    value: &Out,
    positioned: bool,
    id: &dyn Fn(&str) -> String,
) -> Result<(Copied, Option<Any>)> {
    Ok(match value {
        Out::YMap(map) => {
            let mut entries = Vec::new();
            let mut deferred = None;
            for key in ordered_keys(txn, map) {
                let item = map.get(txn, &key).unwrap_or(Out::Any(Any::Undefined));
                if positioned && key == POSITIONS_KEY {
                    deferred = Some(match item {
                        Out::Any(any) => any,
                        _ => Any::Undefined,
                    });
                } else {
                    entries.push((Arc::<str>::from(key), copy(txn, &item, false, id)?.0));
                }
            }
            (Copied::Map(entries), deferred)
        }
        Out::YArray(array) => {
            let items = array
                .iter(txn)
                .map(|item| copy(txn, &item, false, id).map(|copied| copied.0))
                .collect::<Result<Vec<_>>>()?;
            (Copied::Array(items), None)
        }
        Out::YText(text) => {
            let mut delta = Vec::new();
            for diff in text.diff(txn, YChange::identity) {
                let insert = match &diff.insert {
                    Out::Any(Any::String(text)) => Copied::Any(Any::String(text.clone())),
                    other => copy(txn, other, false, id)?.0,
                };
                delta.push(Delta::Inserted(insert, diff.attributes));
            }
            (Copied::Text(delta), None)
        }
        Out::Any(any) => (Copied::Any(translate(any, id)), None),
        other => return Err(fail(format!("unexpected {} value", type_name(other)))),
    })
}

pub(crate) type Attrs = HashMap<Arc<str>, Any>;

#[derive(Clone)]
pub(crate) enum Op {
    Insert(Out, Option<Attrs>),
    Retain(u32, Option<Attrs>),
    Delete(u32),
}

/// `JSON.stringify(a ?? null) === JSON.stringify(b ?? null)` on attribute values.
pub(crate) fn same_attribute(a: Option<&Any>, b: Option<&Any>) -> bool {
    let null = |value: Option<&Any>| match value {
        None | Some(Any::Null) | Some(Any::Undefined) => None,
        Some(other) => Some(other.clone()),
    };
    null(a) == null(b)
}

/// A captured-position text delta as the same edit of the seed text (`landDelta`).
#[allow(clippy::too_many_arguments)]
fn land_delta<T: ReadTxn>(
    edited: &T,
    f: &Alignment,
    delta: &[Op],
    had: &[Rc<Attrs>],
    id: &dyn Fn(&str) -> String,
    place: &str,
    attributes: &dyn Fn(&Attrs, usize) -> Attrs,
) -> Result<Vec<Delta<Copied>>> {
    let mut out: Vec<Delta<Copied>> = Vec::new();
    let mut c: usize = 0;
    let mut s: i64 = 0;
    fn advance(out: &mut Vec<Delta<Copied>>, s: &mut i64, to: i64) {
        if to > *s {
            out.push(Delta::Retain((to - *s) as u32, None));
        }
        *s = to;
    }
    enum Kind {
        Delete,
        Format(Attrs),
    }
    // A range lands in runs of units that stay together in the seed.
    let range = |out: &mut Vec<Delta<Copied>>,
                 s: &mut i64,
                 c: &mut usize,
                 length: usize,
                 kind: &Kind|
     -> Result<()> {
        let mut at = *c;
        while at < *c + length {
            let target = f.map[at];
            if target < 0 {
                return Err(fail(format!(
                    "a change at {place}:{at} touches content the export wrote differently"
                )));
            }
            // Reference fields only the seed holds, inside a delete, go with it.
            if matches!(kind, Kind::Delete) && at > *c && target > *s && past(f, *s) == target {
                out.push(Delta::Deleted((target - *s) as u32));
                *s = target;
            }
            let mut run = 1;
            while at + run < *c + length && f.map[at + run] == target + run as i64 {
                run += 1;
            }
            advance(out, s, target);
            out.push(match kind {
                Kind::Delete => Delta::Deleted(run as u32),
                Kind::Format(changed) => Delta::Retain(run as u32, Some(Box::new(changed.clone()))),
            });
            *s += run as i64;
            at += run;
        }
        *c += length;
        Ok(())
    };
    for op in delta {
        match op {
            Op::Insert(value, attrs) => {
                let at = point_at(f, c).ok_or_else(|| {
                    fail(format!(
                        "an insertion at {place}:{c} lands in content the export wrote differently"
                    ))
                })?;
                advance(&mut out, &mut s, at);
                let insert = match value {
                    Out::Any(Any::String(text)) => Copied::Any(Any::String(text.clone())),
                    other => copy(edited, other, false, id)?.0,
                };
                let attrs = attrs.as_ref().map(|attrs| Box::new(attributes(attrs, c)));
                out.push(Delta::Inserted(insert, attrs));
            }
            Op::Delete(length) => range(&mut out, &mut s, &mut c, *length as usize, &Kind::Delete)?,
            Op::Retain(length, Some(attrs)) => {
                // Next to a formatting change Yjs repeats attributes the
                // captured units already had; only those that change land.
                let end = c + *length as usize;
                while c < end {
                    let mut run = 1;
                    while c + run < end && Rc::ptr_eq(&had[c + run], &had[c]) {
                        run += 1;
                    }
                    let mut changed: Attrs = HashMap::new();
                    for (key, value) in attrs {
                        let before = had.get(c).and_then(|had| had.get(key));
                        if !same_attribute(Some(value), before) {
                            changed.insert(key.clone(), value.clone());
                        }
                    }
                    if changed.is_empty() {
                        c += run;
                    } else {
                        range(&mut out, &mut s, &mut c, run, &Kind::Format(changed))?;
                    }
                }
            }
            Op::Retain(length, None) => c += *length as usize,
        }
    }
    Ok(out)
}

/// A text's unit keys apart from the units it shares with nothing on the other side.
fn outside(text: &Units, apart: &HashSet<(u64, u32)>) -> Vec<super::align::Key> {
    text.keys
        .iter()
        .zip(&text.ids)
        .filter(|(_, id)| !apart.contains(id))
        .map(|(key, _)| key.clone())
        .collect()
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum Action {
    Add,
    Update,
    Delete,
}

#[derive(Clone, Debug, PartialEq, Eq, PartialOrd, Ord)]
enum Segment {
    Key(String),
    Index(u32),
}

impl Segment {
    fn text(&self) -> String {
        match self {
            Segment::Key(key) => key.clone(),
            Segment::Index(index) => index.to_string(),
        }
    }
}

enum StepKind {
    Text {
        delta: Vec<Op>,
        target: TextRef,
    },
    Array {
        target: ArrayRef,
    },
    Map {
        keys: Vec<(String, Action)>,
        target: MapRef,
    },
}

struct Step {
    root: &'static str,
    path: Vec<Segment>,
    kind: StepKind,
}

impl Step {
    fn target(&self) -> Out {
        match &self.kind {
            StepKind::Text { target, .. } => Out::YText(target.clone()),
            StepKind::Array { target } => Out::YArray(target.clone()),
            StepKind::Map { target, .. } => Out::YMap(target.clone()),
        }
    }

    fn place(&self) -> String {
        std::iter::once(self.root.to_owned())
            .chain(self.path.iter().map(Segment::text))
            .collect::<Vec<_>>()
            .join("/")
    }
}

pub(crate) struct Landed {
    pub state: Vec<u8>,
    pub ids: Vec<(String, String)>,
}

/// latest − captured landed on `seed`, written by `client`. `ids` maps
/// captured entity ids to the seed's (in insertion order). Any failure
/// refuses the rebase.
pub(crate) fn transplant(
    lineage: &Lineage,
    captured: &[u8],
    latest: &[u8],
    seed: &[u8],
    ids: &[(String, String)],
    client: u64,
) -> Result<Landed> {
    land_later(lineage, captured, latest, seed, ids, client).map_err(|error| match error {
        Error::Rebase(_) => error,
        other => Error::Rebase(other.message()),
    })
}

fn lock<T>(mutex: &Mutex<T>) -> MutexGuard<'_, T> {
    mutex
        .lock()
        .unwrap_or_else(|poisoned| poisoned.into_inner())
}

fn observe_steps(lineage: &Lineage, edited: &Doc, later: &Doc) -> Result<Vec<Step>> {
    let steps: Arc<Mutex<Vec<Step>>> = Arc::default();
    let failure: Arc<Mutex<Option<Error>>> = Arc::default();
    let mut subscriptions = Vec::new();
    let roots: Vec<(&'static str, Out)> = {
        let txn = edited.transact();
        lineage
            .maps
            .iter()
            .map(|name| (*name, Out::YMap(root_map(&txn, name))))
            .chain(lineage.arrays.iter().map(|name| {
                (
                    *name,
                    Out::YArray(txn.get_array(*name).expect("the lineage's roots exist")),
                )
            }))
            .collect()
    };
    for (name, root) in roots {
        let steps = steps.clone();
        let failure = failure.clone();
        let observe = move |txn: &TransactionMut, events: &yrs::types::Events| {
            for event in events.iter() {
                let path: Vec<Segment> = event
                    .path()
                    .into_iter()
                    .map(|segment| match segment {
                        PathSegment::Key(key) => Segment::Key(key.to_string()),
                        PathSegment::Index(index) => Segment::Index(index),
                    })
                    .collect();
                let kind = match event {
                    Event::Text(event) => StepKind::Text {
                        delta: event
                            .delta(txn)
                            .iter()
                            .map(|op| match op {
                                Delta::Inserted(value, attrs) => {
                                    Op::Insert(value.clone(), attrs.as_deref().cloned())
                                }
                                Delta::Retain(length, attrs) => {
                                    Op::Retain(*length, attrs.as_deref().cloned())
                                }
                                Delta::Deleted(length) => Op::Delete(*length),
                            })
                            .collect(),
                        target: event.target().clone(),
                    },
                    Event::Array(event) => StepKind::Array {
                        target: event.target().clone(),
                    },
                    Event::Map(event) => {
                        let mut keys: Vec<(String, Action)> = event
                            .keys(txn)
                            .iter()
                            .map(|(key, change)| {
                                let action = match change {
                                    EntryChange::Inserted(_) => Action::Add,
                                    EntryChange::Updated(..) => Action::Update,
                                    EntryChange::Removed(_) => Action::Delete,
                                };
                                (key.to_string(), action)
                            })
                            .collect();
                        keys.sort_by(|a, b| a.0.cmp(&b.0));
                        StepKind::Map {
                            keys,
                            target: event.target().clone(),
                        }
                    }
                    _ => {
                        lock(&failure).get_or_insert_with(|| fail("unexpected YXmlElement change"));
                        continue;
                    }
                };
                lock(&steps).push(Step {
                    root: name,
                    path,
                    kind,
                });
            }
        };
        subscriptions.push(match root {
            Out::YArray(array) => array.observe_deep(observe),
            Out::YMap(map) => map.observe_deep(observe),
            _ => unreachable!("roots are maps or arrays"),
        });
    }
    let diff = later
        .transact()
        .encode_state_as_update_v1(&edited.transact().state_vector());
    let update = Update::decode_v1(&diff).map_err(|error| fail(error.to_string()))?;
    edited
        .transact_mut()
        .apply_update(update)
        .map_err(|error| fail(error.to_string()))?;
    drop(subscriptions);
    if let Some(error) = lock(&failure).take() {
        return Err(error);
    }
    let order = |root: &str| {
        lineage
            .maps
            .iter()
            .chain(lineage.arrays)
            .position(|name| *name == root)
            .unwrap_or(usize::MAX)
    };
    let mut steps = std::mem::take(&mut *lock(&steps));
    steps.sort_by(|a, b| {
        order(a.root)
            .cmp(&order(b.root))
            .then_with(|| a.path.cmp(&b.path))
    });
    Ok(steps)
}

fn land_later(
    lineage: &Lineage,
    captured: &[u8],
    latest: &[u8],
    seed: &[u8],
    ids: &[(String, String)],
    client: u64,
) -> Result<Landed> {
    let before = open(lineage, captured, None)?;
    let edited = open(lineage, captured, None)?;
    let later = open(lineage, latest, None)?;
    let result = open(lineage, seed, Some(client))?;
    let steps = observe_steps(lineage, &edited, &later)?;
    let keyed = |root: &str| lineage.keyed.contains(&root);
    let id_map: HashMap<&str, &str> = ids
        .iter()
        .map(|(from, to)| (from.as_str(), to.as_str()))
        .collect();

    let mut renamed: Vec<(String, String)> = Vec::new();
    {
        let txn = result.transact();
        for step in &steps {
            let StepKind::Map { keys, .. } = &step.kind else {
                continue;
            };
            if !step.path.is_empty() || !keyed(step.root) {
                continue;
            }
            let taken = root_map(&txn, step.root);
            for (key, action) in keys {
                if *action != Action::Add || !taken.contains_key(&txn, key) {
                    continue;
                }
                let mut n = 1;
                while taken.contains_key(&txn, &format!("{key}~{n}"))
                    || renamed
                        .iter()
                        .any(|(_, name)| *name == format!("{key}~{n}"))
                {
                    n += 1;
                }
                match renamed.iter_mut().find(|(from, _)| from == key) {
                    Some((_, name)) => *name = format!("{key}~{n}"),
                    None => renamed.push((key.clone(), format!("{key}~{n}"))),
                }
            }
        }
    }
    let renamed_map: HashMap<String, String> = renamed.iter().cloned().collect();
    let id = |value: &str| -> String {
        id_map
            .get(value)
            .map(|to| (*to).to_owned())
            .or_else(|| renamed_map.get(value).cloned())
            .unwrap_or_else(|| value.to_owned())
    };

    // Every captured entity a change touches has a counterpart in the seed;
    // a text's positions correspond through its alignment.
    let mut touched: Vec<(&'static str, Vec<String>)> = Vec::new();
    for step in &steps {
        if !keyed(step.root) {
            continue;
        }
        let slot = match touched.iter().position(|(root, _)| *root == step.root) {
            Some(slot) => slot,
            None => {
                touched.push((step.root, Vec::new()));
                touched.len() - 1
            }
        };
        let keys = &mut touched[slot].1;
        let mut add = |key: &String| {
            if !keys.contains(key) {
                keys.push(key.clone());
            }
        };
        match (&step.kind, step.path.first()) {
            (_, Some(Segment::Key(first))) => add(first),
            (StepKind::Map { keys: changed, .. }, None) => {
                changed.iter().for_each(|(key, _)| add(key))
            }
            _ => {}
        }
    }
    let mut alignments: HashMap<(String, String), Alignment> = HashMap::new();
    {
        let before_txn = before.transact();
        let result_txn = result.transact();
        for (name, keys) in &touched {
            let (from, to) = (root_map(&before_txn, name), root_map(&result_txn, name));
            for key in keys {
                let Some(value) = from
                    .get(&before_txn, key)
                    .filter(|value| !matches!(value, Out::Any(Any::Null | Any::Undefined)))
                else {
                    continue;
                };
                let counterpart = if id_map.contains_key(key.as_str()) {
                    to.get(&result_txn, &id(key))
                } else {
                    None
                };
                let matches = match (&counterpart, &value) {
                    (Some(Out::Any(_)), Out::Any(_)) => true,
                    (counterpart, value) => same_type(counterpart.as_ref(), value),
                };
                if !matches {
                    return Err(fail(format!(
                        "{name} {key} has no counterpart in the export"
                    )));
                }
                if let (Out::YText(source), Some(Out::YText(target))) = (&value, &counterpart) {
                    alignments.insert(
                        ((*name).to_owned(), key.clone()),
                        align(units(&before_txn, source), units(&result_txn, target))?,
                    );
                }
            }
        }
    }

    // An embed's changes land on its seed counterpart, found through the
    // captured position of the same item before any text changes.
    let mut targets: Vec<Out> = Vec::with_capacity(steps.len());
    {
        let txn = result.transact();
        for step in &steps {
            let embedded = step
                .path
                .iter()
                .any(|segment| matches!(segment, Segment::Index(_)));
            let target = if !embedded {
                let mut target = Some(if lineage.arrays.contains(&step.root) {
                    Out::YArray(txn.get_array(step.root).expect("the lineage's roots exist"))
                } else {
                    Out::YMap(root_map(&txn, step.root))
                });
                for (depth, segment) in step.path.iter().enumerate() {
                    target = match (&target, segment) {
                        (Some(Out::YMap(map)), Segment::Key(key)) => {
                            let key = if depth == 0 && keyed(step.root) {
                                id(key)
                            } else {
                                key.clone()
                            };
                            map.get(&txn, &key)
                        }
                        _ => None,
                    };
                }
                target
            } else {
                let unexpected = || fail(format!("unexpected change at {}", step.place()));
                let (Some(Segment::Key(key)), Some(Segment::Index(_)), 2) =
                    (step.path.first(), step.path.get(1), step.path.len())
                else {
                    return Err(unexpected());
                };
                let counterpart = root_map(&txn, step.root).get(&txn, &id(key));
                let f = alignments
                    .get(&(step.root.to_owned(), key.clone()))
                    .ok_or_else(unexpected)?;
                let branch_id = match step.target() {
                    Out::YMap(map) => map.as_ref().id(),
                    Out::YText(text) => AsRef::<Branch>::as_ref(&text).id(),
                    Out::YArray(array) => array.as_ref().id(),
                    _ => return Err(unexpected()),
                };
                let yrs::BranchID::Nested(item) = branch_id else {
                    return Err(unexpected());
                };
                let c = f
                    .from
                    .ids
                    .iter()
                    .position(|unit| *unit == (item.client.get(), item.clock));
                let s = c.map(|c| f.map[c]).unwrap_or(-1);
                match (s, counterpart) {
                    (s, Some(Out::YText(text))) if s >= 0 => child_at_offset(&txn, &text, s as u32),
                    _ => None,
                }
            };
            if !same_type(target.as_ref(), &step.target()) {
                return Err(fail(format!("{} is missing from the export", step.place())));
            }
            targets.push(target.unwrap_or_else(|| step.target()));
        }
    }

    // Parents first: an entity lands before changes inside it.
    let mut order: Vec<usize> = (0..steps.len()).collect();
    order.sort_by_key(|index| steps[*index].path.len());
    {
        let before_txn = before.transact();
        let edited_txn = edited.transact();
        let later_txn = later.transact();
        let mut txn = result.transact_mut();
        let mut positions: Vec<(MapRef, Any)> = Vec::new();
        for index in order {
            let step = &steps[index];
            match (&step.kind, &targets[index]) {
                (StepKind::Text { delta, .. }, Out::YText(text)) => {
                    let unexpected = || fail(format!("unexpected text at {}", step.place()));
                    let (Some(Segment::Key(key)), 1) = (step.path.first(), step.path.len()) else {
                        return Err(unexpected());
                    };
                    let f = alignments
                        .get(&(step.root.to_owned(), key.clone()))
                        .ok_or_else(unexpected)?;
                    let Some(Out::YText(captured_text)) =
                        root_map(&before_txn, step.root).get(&before_txn, key)
                    else {
                        return Err(unexpected());
                    };
                    let had = unit_attributes(&before_txn, &captured_text);
                    let fields = (lineage.docx && step.root == "stories").then(|| {
                        docx::field_attributes(&before_txn, &captured_text, &txn, text, f)
                    });
                    let rewrite = |attrs: &Attrs, at: usize| match &fields {
                        Some(fields) => fields.rewrite(attrs, at),
                        None => attrs.clone(),
                    };
                    let landed =
                        land_delta(&edited_txn, f, delta, &had, &id, &step.place(), &rewrite)?;
                    text.apply_delta(&mut txn, landed);
                }
                (StepKind::Array { target: source }, Out::YArray(array)) => {
                    let length = array.len(&txn);
                    if length > 0 {
                        array.remove_range(&mut txn, 0, length);
                    }
                    let items = source
                        .iter(&edited_txn)
                        .map(|item| copy(&edited_txn, &item, false, &id).map(|copied| copied.0))
                        .collect::<Result<Vec<_>>>()?;
                    for (at, item) in items.into_iter().enumerate() {
                        array.insert(&mut txn, at as u32, item);
                    }
                }
                (
                    StepKind::Map {
                        keys,
                        target: source,
                    },
                    Out::YMap(map),
                ) => {
                    let entities = step.path.is_empty() && keyed(step.root);
                    let positioned = lineage.docx && POSITION_ROOTS.contains(&step.root);
                    for (key, action) in keys {
                        let name = if entities { id(key) } else { key.clone() };
                        if *action == Action::Delete {
                            map.remove(&mut txn, &name);
                            continue;
                        }
                        let value = source
                            .get(&edited_txn, key)
                            .unwrap_or(Out::Any(Any::Undefined));
                        if positioned && !entities && key == POSITIONS_KEY {
                            if let Out::Any(any) = value {
                                positions.push((map.clone(), any));
                            }
                            continue;
                        }
                        if lineage.docx && !step.path.is_empty() {
                            let old = map.get(&txn, key);
                            let (next, _) = copy(&edited_txn, &value, false, &id)?;
                            map.insert(
                                &mut txn,
                                name,
                                docx::keep_continuation(key, old.as_ref(), next),
                            );
                            continue;
                        }
                        let (copied, anchors) =
                            copy(&edited_txn, &value, positioned && entities, &id)?;
                        let inserted = map.insert(&mut txn, name, copied);
                        if let (Some(anchors), Out::YMap(inserted)) = (anchors, inserted) {
                            positions.push((inserted, anchors));
                        }
                    }
                }
                _ => return Err(fail(format!("{} is missing from the export", step.place()))),
            }
        }
        if lineage.docx {
            let mut cache = docx::AnchorAlignments::default();
            for (map, value) in positions {
                let rewritten = docx::rewrite_anchors(&value, &edited_txn, &txn, &id, &mut cache)?;
                map.insert(&mut txn, POSITIONS_KEY, rewritten);
            }
            docx::settle(&mut txn);
            docx::check(&txn, &later_txn, &id, &before_txn)?;
        }
    }

    // Every touched entity reads as in the latest state, but for what the
    // export wrote its own way in a text it kept.
    let replaced: HashSet<(String, String)> = steps
        .iter()
        .filter(|step| step.path.is_empty() && keyed(step.root))
        .filter_map(|step| match &step.kind {
            StepKind::Map { keys, .. } => Some(
                keys.iter()
                    .filter(|(_, action)| *action != Action::Delete)
                    .map(|(key, _)| (step.root.to_owned(), key.clone()))
                    .collect::<Vec<_>>(),
            ),
            _ => None,
        })
        .flatten()
        .collect();
    {
        let later_txn = later.transact();
        let result_txn = result.transact();
        for (name, keys) in &touched {
            for key in keys {
                let value = root_map(&later_txn, name).get(&later_txn, key);
                let next = root_map(&result_txn, name).get(&result_txn, &id(key));
                if value.is_none() != next.is_none() {
                    return Err(fail(format!("{name} {key} did not land")));
                }
                let Some(Out::YText(value)) = value else {
                    continue;
                };
                let f = if replaced.contains(&((*name).to_owned(), key.clone())) {
                    None
                } else {
                    alignments.get(&((*name).to_owned(), key.clone()))
                };
                let mut from_apart = HashSet::new();
                let mut to_apart: HashSet<(u64, u32)> = f
                    .map(|f| f.to.ids.iter().copied().collect())
                    .unwrap_or_default();
                if let Some(f) = f {
                    for (index, target) in f.map.iter().enumerate() {
                        if *target < 0 {
                            from_apart.insert(f.from.ids[index]);
                        } else {
                            to_apart.remove(&f.to.ids[*target as usize]);
                        }
                    }
                }
                let next_units = match &next {
                    Some(Out::YText(next)) => units(&result_txn, next),
                    _ => Units::default(),
                };
                if outside(&units(&later_txn, &value), &from_apart)
                    != outside(&next_units, &to_apart)
                {
                    return Err(fail(format!("{name} {key} did not land")));
                }
            }
        }
    }
    let state = result
        .transact()
        .encode_state_as_update_v1(&yrs::StateVector::default());
    let mut all_ids = ids.to_vec();
    for (from, to) in renamed {
        match all_ids.iter_mut().find(|(key, _)| *key == from) {
            Some(entry) => entry.1 = to,
            None => all_ids.push((from, to)),
        }
    }
    Ok(Landed {
        state,
        ids: all_ids,
    })
}
