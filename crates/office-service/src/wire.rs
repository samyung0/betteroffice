//! The JSON boundary: the Office worker's `{method, args}` messages with
//! `Uint8Array`s as `{"$bytes": base64}`, answered with `{"value"}` or
//! `{"error", "kind"}`. The parity harness and the benchmark CLI use it;
//! a server uses the typed API.

use std::sync::OnceLock;

use base64::Engine as _;
use base64::engine::general_purpose::STANDARD;
use serde::Deserialize;
use serde_json::{Value, json};

use crate::env::{Env, with_env};
use crate::types::{BaselineEntry, Checkpoint, Command, Determinism, Format, ObjectRef};
use crate::{Error, XlsxReplicas};

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct Request {
    method: String,
    #[serde(default)]
    args: Vec<Value>,
    /// Client ids for `randomInt` in call order, and `Date.now()`.
    #[serde(default)]
    clients: Vec<u64>,
    #[serde(default)]
    now_ms: Option<f64>,
    #[serde(default)]
    perf_ms: Option<f64>,
}

fn replicas() -> &'static XlsxReplicas {
    static REPLICAS: OnceLock<XlsxReplicas> = OnceLock::new();
    REPLICAS.get_or_init(|| XlsxReplicas::new(0))
}

pub fn bytes(value: &Value) -> Result<Vec<u8>, Error> {
    value
        .get("$bytes")
        .and_then(Value::as_str)
        .and_then(|text| STANDARD.decode(text).ok())
        .ok_or_else(|| Error::Invalid("expected {\"$bytes\": base64}".into()))
}

pub fn to_bytes(bytes: &[u8]) -> Value {
    json!({ "$bytes": STANDARD.encode(bytes) })
}

struct OwnedCheckpoint {
    format: Format,
    schema_version: u32,
    base_sha256: String,
    state: Vec<u8>,
}

impl OwnedCheckpoint {
    fn parse(value: &Value) -> Result<Self, Error> {
        let field = |key: &str| value.get(key).cloned().unwrap_or(Value::Null);
        Ok(Self {
            format: serde_json::from_value(field("format")).map_err(invalid)?,
            schema_version: field("schemaVersion").as_u64().unwrap_or_default() as u32,
            base_sha256: field("baseSha256").as_str().unwrap_or_default().to_owned(),
            state: bytes(&field("state"))?,
        })
    }

    fn view(&self) -> Checkpoint<'_> {
        Checkpoint {
            format: self.format,
            schema_version: self.schema_version,
            base_sha256: &self.base_sha256,
            state: &self.state,
        }
    }
}

fn invalid(error: impl std::fmt::Display) -> Error {
    Error::Invalid(error.to_string())
}

fn arg(args: &[Value], index: usize) -> &Value {
    args.get(index).unwrap_or(&Value::Null)
}

fn typed<T: serde::de::DeserializeOwned>(value: &Value) -> Result<T, Error> {
    serde_json::from_value(value.clone()).map_err(invalid)
}

fn value(result: impl serde::Serialize) -> Result<Value, Error> {
    serde_json::to_value(result).map_err(invalid)
}

fn dispatch(method: &str, args: &[Value]) -> Result<Value, Error> {
    let base = || bytes(arg(args, 0));
    let checkpoint = |index| OwnedCheckpoint::parse(arg(args, index));
    match method {
        "seedOffice" => {
            let format: Format = typed(arg(args, 0))?;
            let base = bytes(arg(args, 1))?;
            let state = crate::seed(format, &base)?;
            Ok(json!({
                "format": format,
                "schemaVersion": 1,
                "baseSha256": crate::sha256_hex(&base),
                "state": to_bytes(&state),
            }))
        }
        "officeBaseline" => value(crate::baseline(&base()?, checkpoint(1)?.view())?),
        "compareBaselines" => {
            let from: Vec<BaselineEntry> = typed(arg(args, 0))?;
            let to: Vec<BaselineEntry> = typed(arg(args, 1))?;
            value(crate::compare_baselines(&from, &to))
        }
        "saveEffects" => {
            let indexed: Vec<BaselineEntry> = typed(arg(args, 1))?;
            value(crate::save_effects(
                &base()?,
                &indexed,
                checkpoint(2)?.view(),
            )?)
        }
        "compare" => value(crate::compare(
            &base()?,
            checkpoint(1)?.view(),
            checkpoint(2)?.view(),
        )?),
        "xlsxPendingEffects" => {
            let (base, checkpoint) = (base()?, checkpoint(1)?);
            match arg(args, 2) {
                Value::String(room) if !room.is_empty() => {
                    value(replicas().effects(&base, checkpoint.view(), room)?)
                }
                Value::Null | Value::String(_) => {
                    value(crate::xlsx_pending_effects(&base, checkpoint.view())?)
                }
                _ => Err(Error::Invalid("Expected a room name".into())),
            }
        }
        "configureOfficeReplicas" => {
            // JS `Number.isSafeInteger(budget) && budget >= 0`.
            const SAFE: f64 = 9_007_199_254_740_991.0;
            let budget = arg(args, 0)
                .as_f64()
                .filter(|budget| budget.fract() == 0.0 && (0.0..=SAFE).contains(budget))
                .map(|budget| budget as u64)
                .ok_or_else(|| {
                    Error::Invalid("Expected a nonnegative replica budget in bytes".into())
                })?;
            replicas().configure(budget);
            Ok(Value::Null)
        }
        "dropOfficeReplica" => {
            let room = arg(args, 0)
                .as_str()
                .ok_or_else(|| Error::Invalid("Expected a room name".into()))?;
            replicas().drop_room(room);
            Ok(Value::Null)
        }
        "officeReplicaStats" => {
            let stats = replicas().stats();
            Ok(json!({
                "replicas": stats.replicas,
                "replicaBytes": stats.replica_bytes,
                "hits": stats.hits,
                "misses": stats.misses,
                "evictions": stats.evictions,
            }))
        }
        "rebaseOffice" => {
            let rebased = crate::rebase(
                &base()?,
                checkpoint(1)?.view(),
                checkpoint(2)?.view(),
                &bytes(arg(args, 3))?,
            )?;
            Ok(json!({ "state": to_bytes(&rebased.state), "effects": value(rebased.effects)? }))
        }
        "resolveAsset" => {
            let object_ref: ObjectRef = typed(arg(args, 2))?;
            let asset = crate::resolve_asset(&base()?, checkpoint(1)?.view(), &object_ref)?;
            Ok(json!({
                "bytes": to_bytes(&asset.bytes),
                "mimeType": asset.mime_type,
                "sha256": asset.sha256,
            }))
        }
        "exportOffice" => {
            let determinism = arg(args, 2);
            let text = |key: &str| {
                determinism
                    .get(key)
                    .and_then(Value::as_str)
                    .unwrap_or_default()
            };
            let exported = crate::export(
                &base()?,
                checkpoint(1)?.view(),
                Determinism {
                    seed: text("seed"),
                    now: text("now"),
                },
            )?;
            Ok(to_bytes(&exported))
        }
        "inspectOffice" => value(crate::inspect(&base()?, checkpoint(1)?.view())?),
        "applyOfficeCommands" => {
            let commands: Vec<Command> = typed(arg(args, 2))?;
            let result = crate::apply_commands(&base()?, checkpoint(1)?.view(), &commands)?;
            Ok(json!({
                "state": to_bytes(&result.state),
                "inverse": value(result.inverse)?,
                "targets": value(result.targets)?,
            }))
        }
        "locateOfficeTargets" => {
            let ids: Vec<String> = typed(arg(args, 2))?;
            value(crate::locate_targets(
                &base()?,
                checkpoint(1)?.view(),
                &ids,
            )?)
        }
        "runtimeManifest" => value(crate::runtime_manifest()),
        _ => Err(Error::Invalid(format!("unknown method {method}"))),
    }
}

fn kind(error: &Error) -> &'static str {
    match error {
        Error::Edit { .. } => "edit",
        Error::Rebase(_) => "rebase",
        Error::Invalid(_) => "invalid",
        Error::Engine(_) => "engine",
        Error::Panic(_) => "panic",
    }
}

/// Answers one request.
pub fn call(request: &[u8]) -> Vec<u8> {
    let answer = match serde_json::from_slice::<Request>(request) {
        Err(error) => json!({ "error": error.to_string(), "kind": "invalid" }),
        Ok(request) => {
            let env = Env {
                clients: request.clients,
                now_ms: request.now_ms,
                perf_ms: request.perf_ms,
            };
            match with_env(env, || dispatch(&request.method, &request.args)) {
                Ok(value) => json!({ "value": value }),
                Err(error) => json!({ "error": error.to_string(), "kind": kind(&error) }),
            }
        }
    };
    serde_json::to_vec(&answer).unwrap_or_default()
}
