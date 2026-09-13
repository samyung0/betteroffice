use super::*;
use serde::{Deserialize, Serialize};
use yrs::{Any, Doc, Map, Out, ReadTxn, Transact, WriteTxn};

pub(crate) const ROOT: &str = "xlsx:rebase";
const MANIFEST: &str = "$manifest";

#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub(crate) struct AliasSpan {
    pub start: u64,
    pub len: u64,
    pub target: Option<u64>,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
pub(crate) struct SheetAlias {
    pub target: String,
    pub rows: Vec<AliasSpan>,
    pub cols: Vec<AliasSpan>,
}

#[derive(Clone, Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
struct Manifest {
    version: u32,
    base_sha256: String,
    part_order: Option<Vec<String>>,
    aliases: Option<Vec<SheetAlias>>,
}

#[derive(Clone)]
pub(crate) struct RebaseData {
    manifest: Manifest,
    pub(crate) values: BTreeMap<String, Any>,
}

impl RebaseData {
    fn new(manifest: Manifest, parts: BTreeMap<String, Vec<u8>>) -> Result<Self> {
        let mut values = parts
            .into_iter()
            .map(|(path, bytes)| (format!("part:{path}"), Any::Buffer(bytes.into())))
            .collect::<BTreeMap<_, _>>();
        values.insert(
            MANIFEST.into(),
            Any::String(
                serde_json::to_string(&manifest)
                    .map_err(|error| Error::CollaborativeState(error.to_string()))?
                    .into(),
            ),
        );
        Ok(Self { manifest, values })
    }

    pub(crate) fn read(update: &[u8]) -> Result<Option<Self>> {
        let doc = Doc::new();
        crate::authority::hydrate_doc(&doc, update).map_err(Error::CollaborativeState)?;
        let txn = doc.transact();
        let Some(root) = txn.get_map(ROOT) else {
            return Ok(None);
        };
        let values = root
            .iter(&txn)
            .map(|(key, value)| match value {
                Out::Any(value) => Ok((key.to_owned(), value)),
                _ => Err(Error::CollaborativeState(
                    "invalid rebase package value".into(),
                )),
            })
            .collect::<Result<BTreeMap<_, _>>>()?;
        let Some(Any::String(json)) = values.get(MANIFEST) else {
            return Err(Error::CollaborativeState("missing rebase manifest".into()));
        };
        let manifest: Manifest = serde_json::from_str(json)
            .map_err(|error| Error::CollaborativeState(error.to_string()))?;
        if manifest.version != 1 {
            return Err(Error::CollaborativeState(
                "unsupported rebase manifest".into(),
            ));
        }
        if values.iter().any(|(key, value)| {
            key != MANIFEST && (!key.starts_with("part:") || !matches!(value, Any::Buffer(_)))
        }) {
            return Err(Error::CollaborativeState(
                "rebase parts must be binary".into(),
            ));
        }
        Ok(Some(Self { manifest, values }))
    }

    pub(crate) fn write(&self, doc: &Doc) {
        let mut txn = doc.transact_mut();
        let map = txn.get_or_insert_map(ROOT);
        for (key, value) in &self.values {
            map.insert(&mut txn, key.as_str(), value.clone());
        }
    }

    pub(crate) fn validate<T: ReadTxn>(&self, txn: &T) -> std::result::Result<(), String> {
        let root = txn.get_map(ROOT).ok_or("missing rebase package")?;
        if root.len(txn) as usize != self.values.len()
            || self
                .values
                .iter()
                .any(|(key, value)| root.get(txn, key) != Some(Out::Any(value.clone())))
        {
            return Err("rebase package cannot change within an editing epoch".into());
        }
        Ok(())
    }

    pub(crate) fn bind_values(&mut self, doc: &Doc) {
        let txn = doc.transact();
        let root = txn.get_map(ROOT).expect("validated rebase root");
        for (key, value) in &mut self.values {
            if let Some(Out::Any(current)) = root.get(&txn, key) {
                *value = current;
            }
        }
    }

    fn source(&self, original: &[(String, Vec<u8>)], source_sha: &str) -> Result<Option<Vec<u8>>> {
        if self.manifest.base_sha256 != source_sha {
            return Err(Error::CollaborativeState(
                "rebase package does not match source bytes".into(),
            ));
        }
        if let Some(aliases) = &self.manifest.aliases {
            if self.manifest.part_order.is_some() {
                return Err(Error::CollaborativeState(
                    "live rebase cannot contain indexed aliases".into(),
                ));
            }
            let mut targets = BTreeSet::new();
            for alias in aliases {
                if !targets.insert(&alias.target) {
                    return Err(Error::CollaborativeState(
                        "duplicate indexed sheet alias".into(),
                    ));
                }
                for (spans, limit) in [
                    (&alias.rows, u64::from(MAX_ROWS)),
                    (&alias.cols, u64::from(MAX_COLS)),
                ] {
                    let mut position = 0_u64;
                    for span in spans {
                        if span.start != position
                            || span.len == 0
                            || span.len > limit.saturating_sub(position)
                            || span
                                .target
                                .is_some_and(|target| target > limit || span.len > limit - target)
                        {
                            return Err(Error::CollaborativeState(
                                "invalid indexed axis aliases".into(),
                            ));
                        }
                        position += span.len;
                    }
                    if position != limit {
                        return Err(Error::CollaborativeState(
                            "incomplete indexed axis aliases".into(),
                        ));
                    }
                }
            }
        }
        let Some(order) = &self.manifest.part_order else {
            if self.values.len() != 1 {
                return Err(Error::CollaborativeState(
                    "unexpected indexed package parts".into(),
                ));
            }
            return Ok(None);
        };
        let base = original
            .iter()
            .map(|(path, bytes)| (path.as_str(), bytes))
            .collect::<BTreeMap<_, _>>();
        let names = order.iter().map(String::as_str).collect::<BTreeSet<_>>();
        if names.len() != order.len()
            || self.values.keys().any(|key| {
                key != MANIFEST
                    && !key
                        .strip_prefix("part:")
                        .is_some_and(|path| names.contains(path))
            })
        {
            return Err(Error::CollaborativeState(
                "invalid rebase package manifest".into(),
            ));
        }
        let parts = order
            .iter()
            .map(|path| {
                let bytes = match self.values.get(&format!("part:{path}")) {
                    Some(Any::Buffer(bytes)) => bytes.to_vec(),
                    None => base
                        .get(path.as_str())
                        .ok_or_else(|| {
                            Error::CollaborativeState("rebase package part absent".into())
                        })?
                        .to_vec(),
                    _ => {
                        return Err(Error::CollaborativeState(
                            "invalid binary package part".into(),
                        ));
                    }
                };
                Ok((path.clone(), bytes))
            })
            .collect::<Result<Vec<_>>>()?;
        ooxml_opc::rezip_parts(&parts)
            .map(Some)
            .map_err(Error::Package)
    }
}

pub struct RebasedCheckpoint {
    pub state: Vec<u8>,
    pub indexed_state: Vec<u8>,
}

impl Workbook {
    /// Rebuild a saved checkpoint and indexed projection against one published source.
    pub fn rebase_checkpoint(
        old_source: &[u8],
        captured: &[u8],
        latest: &[u8],
        new_source: &[u8],
        client_id: u64,
    ) -> Result<RebasedCheckpoint> {
        let options = CalculationOptions::default();
        let mut before = Self::open_collaborative(old_source, client_id)?;
        before.apply_update_v1(captured, options)?;
        let mut current = Self::open_collaborative(old_source, client_id)?;
        current.apply_update_v1(latest, options)?;
        let mut indexed = Self::open_collaborative(new_source, client_id)?;
        if before.model.sheets.len() != indexed.model.sheets.len()
            || before
                .model
                .sheets
                .iter()
                .zip(&indexed.model.sheets)
                .any(|(a, b)| a.name != b.name)
        {
            return Err(Error::CollaborativeState(
                "published workbook does not match captured sheet order".into(),
            ));
        }
        let aliases = before
            .authority
            .rebase_aliases(&current.authority)
            .map_err(authority_error)?;
        let exported = current.save()?;
        let mut rebased = Self::open_collaborative(&exported, client_id)?;
        let base_parts = ooxml_opc::unzip_parts(new_source).map_err(Error::Package)?;
        let current_parts = ooxml_opc::unzip_parts(&exported).map_err(Error::Package)?;
        let base = base_parts
            .iter()
            .map(|(path, bytes)| (path.as_str(), bytes))
            .collect::<BTreeMap<_, _>>();
        let changed = current_parts
            .iter()
            .filter(|(path, bytes)| {
                base.get(path.as_str())
                    .is_none_or(|original| *original != bytes)
            })
            .cloned()
            .collect();
        let source_sha = format!("{:x}", Sha256::digest(new_source));
        let data = RebaseData::new(
            Manifest {
                version: 1,
                base_sha256: source_sha.clone(),
                part_order: Some(current_parts.iter().map(|(path, _)| path.clone()).collect()),
                aliases: None,
            },
            changed,
        )?;
        rebased
            .authority
            .attach_rebase(data)
            .map_err(authority_error)?;
        let data = RebaseData::new(
            Manifest {
                version: 1,
                base_sha256: source_sha,
                part_order: None,
                aliases: Some(aliases),
            },
            BTreeMap::new(),
        )?;
        indexed
            .authority
            .attach_rebase(data)
            .map_err(authority_error)?;
        let state = rebased.encode_state_as_update_v1();
        let indexed_state = indexed.encode_state_as_update_v1();
        validate_collaboration_size(&state)?;
        validate_collaboration_size(&indexed_state)?;
        Ok(RebasedCheckpoint {
            state,
            indexed_state,
        })
    }

    pub(super) fn restore_rebase(
        &mut self,
        update: &[u8],
        options: CalculationOptions,
    ) -> Result<bool> {
        if !self.authority.is_pristine() || self.authority.has_rebase() {
            return Ok(false);
        }
        let Some(data) = RebaseData::read(update)? else {
            return Ok(false);
        };
        let source_sha = self
            .source_sha
            .as_deref()
            .ok_or_else(|| Error::CollaborativeState("rebase requires source bytes".into()))?;
        let source = self
            .source_package
            .as_ref()
            .ok_or_else(|| Error::CollaborativeState("rebase source package absent".into()))?;
        let parts = source.parts();
        let mut candidate = match data.source(parts, source_sha)? {
            Some(bytes) => Self::open_collaborative(&bytes, self.client_id())?,
            None => Self::from_source(
                source.model().clone(),
                Some(source.clone()),
                self.source_active_sheet,
                false,
                Some(self.client_id()),
                &[],
                Some(source_sha),
            )?,
        };
        candidate.authority.expect_rebase(data);
        if !candidate.restore_snapshot(update, options)? {
            return Err(Error::CollaborativeState(
                "incomplete rebased checkpoint".into(),
            ));
        }
        let observers = self.update_observers.clone();
        candidate.source_sha = self.source_sha.clone();
        candidate.update_observers = observers;
        *self = candidate;
        self.emit_update(UpdateEvent {
            update: self.encode_state_as_update_v1(),
            origin: UpdateOrigin::Local,
        });
        Ok(true)
    }

    /// Identities used only when constructing a compact indexed baseline.
    pub fn checkpoint_sheet_ids(&self) -> Result<Vec<String>> {
        if let Some(aliases) = self.authority.rebase_alias_projection() {
            if aliases.len() != self.model.sheets.len() {
                return Err(Error::CollaborativeState(
                    "indexed projection sheet count changed".into(),
                ));
            }
            return Ok(aliases.iter().map(|alias| alias.target.clone()).collect());
        }
        Ok(self.sheet_info()?.sheet_ids)
    }

    pub fn checkpoint_cell_identities(
        &self,
        cells: impl IntoIterator<Item = (SheetId, CellRef)>,
    ) -> Result<Vec<String>> {
        let Some(aliases) = self.authority.rebase_alias_projection() else {
            return self.cell_identities(cells);
        };
        cells
            .into_iter()
            .map(|(sheet, at)| {
                let alias = aliases.get(sheet.0 as usize).ok_or_else(|| {
                    Error::CollaborativeState("indexed sheet identity absent".into())
                })?;
                #[derive(Serialize)]
                struct IdentityPoint {
                    run: &'static str,
                    offset: u64,
                }
                let point = |value: u32, spans: &[AliasSpan]| -> Result<IdentityPoint> {
                    let value = u64::from(value);
                    let index = spans.partition_point(|span| span.start + span.len <= value);
                    let span = spans
                        .get(index)
                        .filter(|span| value >= span.start)
                        .ok_or_else(|| {
                            Error::CollaborativeState("indexed axis identity absent".into())
                        })?;
                    Ok(match span.target {
                        Some(target) => IdentityPoint {
                            run: "base",
                            offset: target + value - span.start,
                        },
                        None => IdentityPoint {
                            run: "indexed-deleted",
                            offset: value,
                        },
                    })
                };
                Ok(format!(
                    "{}:{}",
                    alias.target,
                    serde_json::to_string(&(
                        point(at.row, &alias.rows)?,
                        point(at.col, &alias.cols)?
                    ))
                    .map_err(|error| Error::CollaborativeState(error.to_string()))?
                ))
            })
            .collect()
    }
}

impl RebaseData {
    pub(crate) fn aliases(&self) -> Option<&[SheetAlias]> {
        self.manifest.aliases.as_deref()
    }
}
