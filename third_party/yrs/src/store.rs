use crate::block::{Block, BlockRange, ClientID, ItemContent, ItemPtr};
use crate::block_store::BlockStore;
use crate::branch::{Branch, BranchPtr};
use crate::doc::{DocAddr, Options};
use crate::error::Error;
use crate::event::SubdocsEvent;
use crate::id_set::DeleteSet;
use crate::slice::ItemSlice;
use crate::types::{Path, PathSegment, TypeRef};
use crate::update::PendingUpdate;
use crate::updates::encoder::{Encode, Encoder};
use crate::{
    Doc, Observer, OffsetKind, Snapshot, TransactionCleanupEvent, TransactionMut, UpdateEvent,
    Uuid, ID,
};
use crate::{IdSet, StateVector};
use arc_swap::{ArcSwap, DefaultStrategy, Guard};
use async_lock::futures::{Read, Write};
use async_lock::{RwLock, RwLockReadGuard, RwLockWriteGuard};
use std::borrow::Borrow;
use std::collections::hash_map::Entry;
use std::collections::{HashMap, HashSet};
use std::ops::Deref;
use std::sync::Arc;

/// Store is a core element of a document. It contains all of the information, like block store
/// map of root types, pending updates waiting to be applied once a missing update information
/// arrives and all subscribed callbacks.
pub struct Store {
    pub(crate) client_id: ClientID,
    pub(crate) offset_kind: OffsetKind,
    pub(crate) skip_gc: bool,
    pub(crate) cleanup_formatting: bool,

    /// Root types (a.k.a. top-level types). These types are defined by users at the document level,
    /// they have their own unique names and represent core shared types that expose operations
    /// which can be called concurrently by remote peers in a conflict-free manner.
    pub(crate) types: HashMap<Arc<str>, Box<Branch>>,

    /// A block store of a current document. It represent all blocks (inserted or tombstoned
    /// operations) integrated - and therefore visible - into a current document.
    pub(crate) blocks: BlockStore,

    /// A pending update. It contains blocks, which are not yet integrated into `blocks`, usually
    /// because due to issues in update exchange, there were some missing blocks that need to be
    /// integrated first before the data from `pending` can be applied safely.
    pub(crate) pending: Option<PendingUpdate>,

    /// A pending delete set. Just like `pending`, it contains deleted ranges of blocks that have
    /// not been yet applied due to missing blocks that prevent `pending` update to be integrated
    /// into `blocks`.
    pub(crate) pending_ds: Option<IdSet>,

    pub(crate) subdocs: HashMap<DocAddr, Doc>,

    pub(crate) events: Option<Box<StoreEvents>>,

    /// Pointer to a parent block - present only if a current document is a sub-document of another
    /// document.
    pub(crate) parent: Option<ItemPtr>,

    /// Dependencies between items and weak links pointing to these items.
    pub(crate) linked_by: HashMap<ItemPtr, HashSet<BranchPtr>>,
}

impl Store {
    /// Create a new empty store in context of a given `client_id`.
    pub(crate) fn new(options: &Options) -> Self {
        Store {
            client_id: options.client_id,
            offset_kind: options.offset_kind,
            skip_gc: options.skip_gc,
            cleanup_formatting: options.cleanup_formatting,
            types: HashMap::default(),
            blocks: BlockStore::default(),
            subdocs: HashMap::default(),
            linked_by: HashMap::default(),
            events: None,
            pending: None,
            pending_ds: None,
            parent: None,
        }
    }

    /// If there are any missing updates, this method will return a pending update which contains
    /// updates waiting for their predecessors to arrive in order to be integrated.
    pub fn pending_update(&self) -> Option<&PendingUpdate> {
        self.pending.as_ref()
    }

    /// Returns a mutable reference to the pending update if it exists.
    pub fn pending_update_mut(&mut self) -> Option<&mut PendingUpdate> {
        self.pending.as_mut()
    }

    /// If there are some delete updates waiting for missing updates to arrive in order to be
    /// applied, this method will return them.
    pub fn pending_ds(&self) -> Option<&IdSet> {
        self.pending_ds.as_ref()
    }

    /// Returns a mutable reference to the pending delete set if it exists.
    pub fn pending_ds_mut(&mut self) -> Option<&mut IdSet> {
        self.pending_ds.as_mut()
    }

    pub fn is_subdoc(&self) -> bool {
        self.parent.is_some()
    }

    /// Get the latest clock sequence number observed and integrated into a current store client.
    /// This is exclusive value meaning it describes a clock value of the beginning of the next
    /// block that's about to be inserted. You cannot use that clock value to find any existing
    /// block content.
    pub fn get_local_state(&self) -> u32 {
        self.blocks.get_clock(&self.client_id)
    }

    /// Returns a branch reference to a complex type identified by its pointer. Returns `None` if
    /// no such type could be found or was ever defined.
    pub(crate) fn get_type<K: Borrow<str>>(&self, key: K) -> Option<BranchPtr> {
        let ptr = BranchPtr::from(self.types.get(key.borrow())?);
        Some(ptr)
    }

    /// Returns a branch reference to a complex type identified by its pointer. Returns `None` if
    /// no such type could be found or was ever defined.
    pub(crate) fn get_or_create_type<K: Into<Arc<str>>>(
        &mut self,
        key: K,
        type_ref: TypeRef,
    ) -> BranchPtr {
        let key = key.into();
        match self.types.entry(key.clone()) {
            Entry::Occupied(e) => {
                let mut branch = BranchPtr::from(e.get());
                branch.repair_type_ref(type_ref);
                branch
            }
            Entry::Vacant(e) => {
                let mut branch = Branch::new(type_ref);
                let mut branch_ref = BranchPtr::from(&mut branch);
                branch_ref.name = Some(key);
                e.insert(branch);
                branch_ref
            }
        }
    }

    /// Encodes all changes from current transaction block store up to a given `snapshot`.
    /// This enables to encode state of a document at some specific point in the past.
    pub fn encode_state_from_snapshot<E: Encoder>(
        &self,
        snapshot: &Snapshot,
        encoder: &mut E,
    ) -> Result<(), Error> {
        if !self.skip_gc {
            return Err(Error::Gc);
        }
        self.write_blocks_to(&snapshot.state_map, encoder);
        snapshot.delete_set.encode(encoder);

        Ok(())
    }

    pub(crate) fn write_blocks_to<E: Encoder>(&self, sv: &StateVector, encoder: &mut E) {
        let local_sv = self.blocks.get_state_vector();
        let mut diff = Vec::with_capacity(sv.len());
        for (&client_id, &clock) in sv.iter() {
            if local_sv.contains_client(&client_id) {
                diff.push((client_id, clock.min(local_sv.get(&client_id))));
            }
        }
        // Write items with higher client ids first
        // This heavily improves the conflict algorithm.
        diff.sort_by(|a, b| b.0.cmp(&a.0));

        encoder.write_var(diff.len());
        for (client, clock) in diff {
            let blocks = self.blocks.get_client(&client).unwrap();
            let clock = clock.min(blocks.clock() + 1);
            let last_idx = blocks.find_index(clock - 1).unwrap();
            // write # encoded structs
            encoder.write_var(last_idx + 1);
            encoder.write_client(client);
            encoder.write_var(0);
            for i in 0..last_idx {
                let block = blocks[i].as_slice();
                block.encode(encoder);
            }
            let last_block = &blocks[last_idx];
            // write first struct with an offset
            let mut slice = last_block.as_slice();
            slice.trim_end(slice.clock_end() - (clock - 1));
            slice.encode(encoder);
        }
    }

    /// Compute a diff to sync with another client.
    ///
    /// This is the most efficient method to sync with another client by only
    /// syncing the differences.
    ///
    /// The sync protocol in Yrs/js is:
    /// * Send StateVector to the other client.
    /// * The other client comutes a minimal diff to sync by using the StateVector.
    pub fn encode_diff<E: Encoder>(&self, sv: &StateVector, encoder: &mut E) {
        //TODO: this could be actually 2 steps:
        // 1. create Diff of block store and remote state vector (it can have lifetime of bock store)
        // 2. make Diff implement Encode trait and encode it
        // this way we can add some extra utility method on top of Diff (like introspection) without need of decoding it.
        self.write_blocks_from(sv, encoder);
        let delete_set = IdSet::from_store(&self.blocks);
        delete_set.encode(encoder);
    }

    pub(crate) fn write_blocks_from<E: Encoder>(&self, sv: &StateVector, encoder: &mut E) {
        let local_sv = self.blocks.get_state_vector();
        let diff = Self::diff_state_vectors(&local_sv, sv);
        self.write_blocks_since(diff, encoder);
    }

    /// The blocks of each `(client, clock)` from that clock on. A transaction's
    /// update passes the clients it inserted into, so encoding it costs
    /// O(changed clients) rather than O(clients).
    pub(crate) fn write_blocks_since<E: Encoder>(
        &self,
        mut diff: Vec<(ClientID, u32)>,
        encoder: &mut E,
    ) {
        // Write items with higher client ids first
        // This heavily improves the conflict algorithm.
        diff.sort_by(|a, b| b.0.cmp(&a.0));

        encoder.write_var(diff.len());
        for (client, clock) in diff {
            let blocks = self.blocks.get_client(&client).unwrap();
            let clock = clock.max(
                blocks
                    .get(0)
                    .map(|i| i.as_ref().clock_start())
                    .unwrap_or_default(),
            ); // make sure the first id exists
            let start = blocks.find_index(clock).unwrap();
            // write # encoded structs
            encoder.write_var(blocks.len() - start);
            encoder.write_client(client);
            encoder.write_var(clock);
            let first_block = blocks.get(start).unwrap().as_ref();
            // write first struct with an offset
            let offset = clock - first_block.clock_start();
            let mut slice = first_block.as_slice();
            slice.trim_start(offset);
            slice.encode(encoder);
            for i in (start + 1)..blocks.len() {
                let block = &blocks[i];
                block.as_slice().encode(encoder);
            }
        }
    }

    fn diff_state_vectors(local_sv: &StateVector, remote_sv: &StateVector) -> Vec<(ClientID, u32)> {
        let mut diff = Vec::new();
        for (client, &remote_clock) in remote_sv.iter() {
            let local_clock = local_sv.get(client);
            if local_clock > remote_clock {
                diff.push((*client, remote_clock));
            }
        }
        for (client, _) in local_sv.iter() {
            if !remote_sv.contains_client(client) {
                diff.push((*client, 0));
            }
        }
        diff
    }

    pub fn get_type_from_path(&self, path: &Path) -> Option<BranchPtr> {
        let mut i = path.iter();
        if let Some(PathSegment::Key(root_name)) = i.next() {
            let mut current = self.get_type(root_name.clone())?;
            while let Some(segment) = i.next() {
                match segment {
                    PathSegment::Key(key) => {
                        let child = current.map.get(key)?;
                        if let ItemContent::Type(child_branch) = &child.content {
                            current = BranchPtr::from(child_branch.as_ref());
                        } else {
                            return None;
                        }
                    }
                    PathSegment::Index(index) => {
                        if let Some((ItemContent::Type(child_branch), _)) = current.get_at(*index) {
                            current = child_branch.into();
                        } else {
                            return None;
                        }
                    }
                }
            }
            Some(current)
        } else {
            None
        }
    }

    /// Consumes current block slice view, materializing it into actual block representation equivalent,
    /// splitting underlying block along [ItemSlice::start]/[ItemSlice::end] offsets.
    ///
    /// Returns a block created this way, that represents the boundaries that current [ItemSlice]
    /// was representing.
    pub(crate) fn materialize(&mut self, mut slice: ItemSlice) -> ItemPtr {
        let id = slice.id().clone();
        let (blocks, work) = self.blocks.client_and_work(&id.client).unwrap();
        let mut links = None;
        let item = slice.ptr.deref();
        if item.info.is_linked() {
            links = self.linked_by.get(&slice.ptr).cloned();
        }

        let mut index = None;
        let mut ptr = if slice.adjacent_left() {
            slice.ptr
        } else {
            let mut i = blocks.find_index(id.clock).unwrap();
            if let Some(new) = slice.ptr.splice(slice.start, OffsetKind::Utf16) {
                if let Some(source) = links.clone() {
                    let dest = self
                        .linked_by
                        .entry(ItemPtr::from(new.as_ref()))
                        .or_default();
                    dest.extend(source);
                }
                blocks.insert(i + 1, Block::Item(new), work);
                i += 1;
                //todo: txn merge blocks insert?
                index = Some(i);
            }
            let ptr = blocks[i].as_item().unwrap();
            slice = ItemSlice::new(ptr, 0, slice.end - slice.start);
            ptr
        };

        if !slice.adjacent_right() {
            // split block on the right side
            let i = if let Some(i) = index {
                i
            } else {
                let last_id = slice.last_id();
                blocks.find_index(last_id.clock).unwrap()
            };
            let new = ptr.splice(slice.len(), OffsetKind::Utf16).unwrap();
            if let Some(source) = links {
                let dest = self
                    .linked_by
                    .entry(ItemPtr::from(new.as_ref()))
                    .or_default();
                dest.extend(source);
            }
            blocks.insert(i + 1, Block::Item(new), work);
            //todo: txn merge blocks insert?
        }

        ptr
    }

    /// Returns a collection of sub documents linked within the structures of this document store.
    pub fn subdocs(&self) -> SubdocsIter {
        SubdocsIter(self.subdocs.values())
    }

    /// Returns a collection of globally unique identifiers of sub documents linked within
    /// the structures of this document store.
    pub fn subdoc_guids(&self) -> SubdocGuids {
        SubdocGuids(self.subdocs.values())
    }

    /// Patched for BetterOffice: the item holding `id`, unless that unit was
    /// garbage-collected or is not held.
    pub fn get_item(&self, id: &ID) -> Option<ItemPtr> {
        self.blocks.get_item(id)
    }

    /// Patched for BetterOffice: the range of the block holding `id` (an item
    /// or a garbage-collected run) and its item, if it is one.
    pub fn get_block_range(&self, id: &ID) -> Option<(BlockRange, Option<ItemPtr>)> {
        let block = self.blocks.get_block(id)?;
        let block = block.as_ref();
        Some((block.range(), block.as_item()))
    }

    /// Patched for BetterOffice: the blocks of `client` that hold clocks in
    /// `clocks` (items or garbage-collected runs) with their items, in clock
    /// order. One search finds the first, so a walk over what an update
    /// inserted reads each block once.
    pub fn blocks_in(
        &self,
        client: &ClientID,
        clocks: std::ops::Range<u32>,
    ) -> impl Iterator<Item = (BlockRange, Option<ItemPtr>)> + '_ {
        let list = self.blocks.get_client(client);
        let first = list
            .filter(|list| list.len() > 0)
            .and_then(|list| list.find_index(clocks.start));
        list.into_iter()
            .flat_map(move |list| {
                (first.unwrap_or(list.len())..list.len()).filter_map(move |at| list.get(at))
            })
            .map(|block| {
                let block = block.as_ref();
                (block.range(), block.as_item())
            })
            .take_while(move |(range, _)| range.clock < clocks.end)
    }

    /// Patched for BetterOffice: the ID and parent of the first live, countable
    /// item right of the item holding `id` (also a deleted one), found by
    /// walking right from it instead of computing an index.
    pub fn next_live_item(&self, id: &ID) -> Option<(ID, BranchPtr)> {
        let mut next = self.blocks.get_item(id)?.right;
        while let Some(item) = next {
            if !item.is_deleted() && item.is_countable() {
                return Some((*item.id(), *item.parent.as_branch()?));
            }
            next = item.right;
        }
        None
    }

    /// The item slice `id` resolves to once redone, at the same offset into
    /// each redone item (as Yjs's `followRedone` keeps its `diff`).
    pub(crate) fn follow_redone(&self, id: &ID) -> Option<ItemSlice> {
        let mut next_id = Some(*id);
        let mut slice = None;
        while let Some(next) = next_id.as_mut() {
            slice = self.blocks.get_item_clean_start(next);
            if let Some(slice) = &slice {
                next_id = slice
                    .ptr
                    .redone
                    .map(|redone| ID::new(redone.client, redone.clock + slice.start));
            } else {
                break;
            }
        }
        slice
    }
}

impl Encode for Store {
    /// Encodes the document state to a binary format.
    ///
    /// Document updates are idempotent and commutative. Caveats:
    /// * It doesn't matter in which order document updates are applied.
    /// * As long as all clients receive the same document updates, all clients
    ///   end up with the same content.
    /// * Even if an update contains known information, the unknown information
    ///   is extracted and integrated into the document structure.
    fn encode<E: Encoder>(&self, encoder: &mut E) {
        self.encode_diff(&StateVector::default(), encoder)
    }
}

impl std::fmt::Debug for Store {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        std::fmt::Display::fmt(self, f)
    }
}

impl std::fmt::Display for Store {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let mut s = f.debug_struct(&self.client_id.to_string());
        if !self.types.is_empty() {
            s.field("root types", &self.types);
        }
        if !self.blocks.is_empty() {
            s.field("blocks", &self.blocks);
        }
        if let Some(pending) = self.pending.as_ref() {
            s.field("pending", pending);
        }
        if let Some(pending_ds) = self.pending_ds.as_ref() {
            s.field("pending delete set", pending_ds);
        }

        if let Some(parent) = self.parent.as_ref() {
            s.field("parent block", parent.id());
        }
        if !self.linked_by.is_empty() {
            s.field("links", &self.linked_by);
        }
        s.finish()
    }
}

#[repr(transparent)]
#[derive(Debug, Clone)]
pub(crate) struct DocStore(pub(crate) Arc<StoreInner>);

impl DocStore {
    pub fn new(options: Options, parent: Option<ItemPtr>) -> Self {
        let mut store = Store::new(&options);
        let options = ArcSwap::new(options.into());
        store.parent = parent;
        DocStore(Arc::new(StoreInner {
            options,
            store: RwLock::new(store),
        }))
    }

    pub(crate) fn try_read(&self) -> Option<RwLockReadGuard<Store>> {
        self.0.store.try_read()
    }

    #[cfg(not(target_family = "wasm"))]
    pub(crate) fn read_blocking(&self) -> RwLockReadGuard<Store> {
        self.0.store.read_blocking()
    }

    #[cfg(target_family = "wasm")]
    pub(crate) fn read_blocking(&self) -> RwLockReadGuard<Store> {
        self.0.store.try_read().unwrap()
    }

    pub(crate) fn read_async(&self) -> Read<Store> {
        self.0.store.read()
    }

    pub(crate) fn try_write(&self) -> Option<RwLockWriteGuard<Store>> {
        self.0.store.try_write()
    }

    #[cfg(not(target_family = "wasm"))]
    pub(crate) fn write_blocking(&self) -> RwLockWriteGuard<Store> {
        self.0.store.write_blocking()
    }

    #[cfg(target_family = "wasm")]
    pub(crate) fn write_blocking(&self) -> RwLockWriteGuard<Store> {
        self.0.store.try_write().unwrap()
    }

    pub(crate) fn write_async(&self) -> Write<Store> {
        self.0.store.write()
    }

    pub(crate) fn options(&self) -> Guard<Arc<Options>, DefaultStrategy> {
        self.0.options.load()
    }

    /// Sets [Doc::should_load] flag, returning previous value.
    pub(crate) fn set_should_load(&self, should_load: bool) -> bool {
        self.0
            .options
            .rcu(|options| {
                let mut options = options.deref().clone();
                options.should_load = should_load;
                options
            })
            .should_load
    }

    pub(crate) fn set_subdoc_data(&self, client_id: ClientID, collection_id: Option<Arc<str>>) {
        self.0.options.rcu(|options| {
            let mut options = options.deref().clone();
            options.client_id = client_id;
            if options.collection_id.is_none() {
                options.collection_id = collection_id.clone();
            }
            options
        });
    }
}

#[derive(Debug)]
pub(crate) struct StoreInner {
    options: ArcSwap<Options>,
    store: RwLock<Store>,
}

#[repr(transparent)]
pub struct SubdocsIter<'doc>(std::collections::hash_map::Values<'doc, DocAddr, Doc>);

impl<'doc> Iterator for SubdocsIter<'doc> {
    type Item = &'doc Doc;

    fn next(&mut self) -> Option<Self::Item> {
        self.0.next()
    }
}

#[repr(transparent)]
pub struct SubdocGuids<'doc>(std::collections::hash_map::Values<'doc, DocAddr, Doc>);

impl<'doc> Iterator for SubdocGuids<'doc> {
    type Item = Uuid;

    fn next(&mut self) -> Option<Self::Item> {
        let d = self.0.next()?;
        Some(d.guid())
    }
}

macro_rules! define_event_type {
    ($name:ident ($($args:tt)*)) => {
        #[cfg(feature = "sync")]
        pub type $name = Box<dyn FnMut($($args)*) + Send + Sync + 'static>;
        #[cfg(not(feature = "sync"))]
        pub type $name = Box<dyn FnMut($($args)*) + 'static>;
    };
}

define_event_type!(TransactionCleanupFn(
    &TransactionMut,
    &TransactionCleanupEvent
));
define_event_type!(AfterTransactionFn(&mut TransactionMut));
define_event_type!(UpdateFn(&TransactionMut, &UpdateEvent));
define_event_type!(SubdocsFn(&TransactionMut, &SubdocsEvent));
define_event_type!(DestroyFn(&TransactionMut, &Doc));
define_event_type!(BeforeObserverCallsFn(&TransactionMut));

#[derive(Default)]
pub struct StoreEvents {
    /// Handles subscriptions for the transaction cleanup event. Events are called with the
    /// newest updates once they are committed and compacted.
    pub transaction_cleanup_events: Observer<TransactionCleanupFn>,

    /// Handles subscriptions for the `afterTransactionCleanup` event. Events are called with the
    /// newest updates once they are committed and compacted.
    pub after_transaction_events: Observer<AfterTransactionFn>,

    /// A subscription handler. It contains all callbacks with registered by user functions that
    /// are supposed to be called, once a new update arrives.
    pub update_v1_events: Observer<UpdateFn>,

    /// A subscription handler. It contains all callbacks with registered by user functions that
    /// are supposed to be called, once a new update arrives.
    pub update_v2_events: Observer<UpdateFn>,

    /// Handles subscriptions for subdocs events.
    pub subdocs_events: Observer<SubdocsFn>,

    pub destroy_events: Observer<DestroyFn>,

    /// Handles subscriptions for the `beforeObserverCalls` event. Callbacks are called after
    /// the transaction body completes but before type-level observers are triggered.
    pub before_observer_calls_events: Observer<BeforeObserverCallsFn>,
}

impl StoreEvents {
    pub fn emit_update_v1(&mut self, txn: &TransactionMut) {
        if self.update_v1_events.has_subscribers() {
            // O(changed clients): comparing state vectors costs O(clients).
            if !txn.delete_set.is_empty() || !txn.insert_set.is_empty() {
                let update = UpdateEvent::new_v1(txn);
                self.update_v1_events
                    .trigger(|callback| callback(txn, &update));
            }
        }
    }

    pub fn emit_update_v2(&mut self, txn: &TransactionMut) {
        if self.update_v2_events.has_subscribers() {
            if !txn.delete_set.is_empty() || !txn.insert_set.is_empty() {
                let update = UpdateEvent::new_v2(txn);
                self.update_v2_events.trigger(|fun| fun(txn, &update));
            }
        }
    }

    pub fn emit_after_transaction(&mut self, txn: &mut TransactionMut) {
        self.after_transaction_events.trigger(|fun| fun(txn));
    }

    pub fn emit_transaction_cleanup(&mut self, txn: &TransactionMut) {
        if self.transaction_cleanup_events.has_subscribers() {
            let event = TransactionCleanupEvent::new(txn);
            self.transaction_cleanup_events
                .trigger(|fun| fun(txn, &event));
        }
    }

    pub fn emit_before_observer_calls(&mut self, txn: &TransactionMut) {
        self.before_observer_calls_events.trigger(|fun| fun(txn));
    }
}

#[cfg(test)]
mod inspection_test {
    use crate::block::ItemContent;
    use crate::types::text::Text;
    use crate::{ClientID, Doc, GetString, ReadTxn, Transact, ID};

    fn id(client: u64, clock: u32) -> ID {
        ID::new(ClientID::new(client), clock)
    }

    // Patched for BetterOffice: the read-only accessors walk what the store holds.
    #[test]
    fn reads_items_links_and_blocks() {
        let doc = Doc::with_client_id(7);
        let text = doc.get_or_insert_text("t");
        let mut txn = doc.transact_mut();
        text.insert(&mut txn, 0, "abc");
        text.insert(&mut txn, 3, "def");
        text.remove_range(&mut txn, 1, 1);
        drop(txn);
        let txn = doc.transact();
        assert_eq!(text.get_string(&txn), "acdef");
        let store = txn.store();
        let first = store.get_item(&id(7, 0)).unwrap();
        assert_eq!(first.origin(), None);
        assert!(matches!(first.content(), ItemContent::String(_)));
        let branch = first.parent_branch().unwrap();
        assert_eq!(branch.start(), Some(first));
        assert!(branch.item().is_none());
        let deleted = first.right().unwrap();
        assert!(deleted.is_deleted());
        assert_eq!(deleted.origin(), Some(&id(7, 0)));
        assert_eq!(deleted.left(), Some(first));
        // "c" and "def" squashed into one item after the transaction.
        let last = store.get_item(&id(7, 4)).unwrap();
        assert_eq!(last.id(), &id(7, 2));
        assert_eq!(last.origin(), Some(&id(7, 1)));
        assert_eq!(last.right_origin(), None);
        let (range, item) = store.get_block_range(&id(7, 4)).unwrap();
        assert_eq!((range.clock, range.len), (2, 4));
        assert_eq!(item, Some(last));
        assert!(store.get_item(&id(8, 0)).is_none());
    }
}
