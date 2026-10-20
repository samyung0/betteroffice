use base64::Engine as _;
use sha2::{Digest, Sha256};

use crate::error::{EditCode, Error, Result};
use crate::js::utf16_len;
use crate::types::{Asset, Checkpoint, Command, Format};

pub fn sha256_hex(bytes: impl AsRef<[u8]>) -> String {
    let digest = Sha256::digest(bytes.as_ref());
    let mut out = String::with_capacity(64);
    for byte in digest {
        out.push(char::from_digit((byte >> 4) as u32, 16).unwrap_or('0'));
        out.push(char::from_digit((byte & 15) as u32, 16).unwrap_or('0'));
    }
    out
}

pub(crate) const CHECKPOINT_MISMATCH: &str =
    "Office checkpoint does not match the exact base package and schema";

/// office-checkpoint.ts `assertCheckpoint`.
pub(crate) fn assert_checkpoint(
    format: Format,
    base: &[u8],
    checkpoint: Option<Checkpoint>,
) -> Result<()> {
    if base.is_empty() {
        return Err(Error::Invalid("Expected nonempty Office base bytes".into()));
    }
    if let Some(checkpoint) = checkpoint
        && (checkpoint.schema_version != 1
            || checkpoint.format != format
            || checkpoint.base_sha256 != sha256_hex(base)
            || checkpoint.state.is_empty())
    {
        return Err(Error::Engine(CHECKPOINT_MISMATCH.into()));
    }
    Ok(())
}

/// office-checkpoint.ts `assetFromDataUrl` for a string `src`.
pub(crate) fn asset_from_data_url(src: &str) -> Result<Asset> {
    let unsupported =
        || Error::Engine("Image object does not contain supported embedded image bytes".into());
    let rest = src.strip_prefix("data:").ok_or_else(unsupported)?;
    let (mime, data) = rest.split_once(";base64,").ok_or_else(unsupported)?;
    let subtype = mime.strip_prefix("image/").ok_or_else(unsupported)?;
    let mime_ok = !subtype.is_empty()
        && subtype
            .bytes()
            .all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'.' | b'+' | b'-'));
    let body = data.trim_end_matches('=');
    let padding = data.len() - body.len();
    let data_ok = padding <= 2
        && body
            .bytes()
            .all(|byte| byte.is_ascii_alphanumeric() || byte == b'+' || byte == b'/');
    if !mime_ok || !data_ok {
        return Err(unsupported());
    }
    let bytes = node_base64(body);
    Ok(Asset {
        sha256: sha256_hex(&bytes),
        bytes,
        mime_type: mime.to_owned(),
    })
}

/// `Buffer.from(text, "base64")` on validated base64 without its padding:
/// trailing bits that do not make a byte are dropped, as Node does.
fn node_base64(body: &str) -> Vec<u8> {
    let usable = body.len() - body.len() % 4;
    let mut bytes = base64::engine::general_purpose::STANDARD_NO_PAD
        .decode(&body[..usable])
        .unwrap_or_default();
    let tail = &body[usable..];
    if tail.len() >= 2 {
        let config = base64::engine::GeneralPurposeConfig::new()
            .with_decode_allow_trailing_bits(true)
            .with_decode_padding_mode(base64::engine::DecodePaddingMode::Indifferent);
        let engine = base64::engine::GeneralPurpose::new(&base64::alphabet::STANDARD, config);
        bytes.extend(engine.decode(tail).unwrap_or_default());
    }
    bytes
}

/// office-checkpoint.ts `imageMimeType`.
pub(crate) fn image_mime_type(part: &str) -> Result<&'static str> {
    let extension = part.rsplit('.').next().unwrap_or_default().to_lowercase();
    let mime = match extension.as_str() {
        "png" => "image/png",
        "jpg" | "jpeg" => "image/jpeg",
        "gif" => "image/gif",
        "svg" => "image/svg+xml",
        "webp" => "image/webp",
        "emf" => "image/emf",
        "wmf" => "image/wmf",
        "tiff" | "tif" => "image/tiff",
        "bmp" => "image/bmp",
        _ => {
            return Err(Error::Engine(format!(
                "Unsupported image part type: {part}"
            )));
        }
    };
    Ok(mime)
}

/// office-checkpoint.ts `checkReplacement`: `expected_text` occurs exactly
/// once in the paragraph, counted as note edits count it (an empty one only
/// in an empty paragraph). Returns the span of that occurrence the
/// replacement changes, in paragraph offsets, and the paragraph's new text.
/// A refusal of the match carries the paragraph as it reads now.
pub(crate) fn check_replacement(
    current: &str,
    command: &Command,
    format: &str,
) -> Result<(Span, String)> {
    let Command::ReplaceText {
        target_id,
        expected_text,
        text,
    } = command
    else {
        return Err(Error::edit(
            EditCode::UnsupportedOperation,
            format!("{format} sources support replace_text only"),
        ));
    };
    if text.contains('\n') {
        return Err(Error::edit(
            EditCode::InvalidInput,
            "replacement text must stay within one paragraph",
        ));
    }
    let refuse = |code, message: String| Error::Edit {
        code,
        message,
        details: Some(serde_json::json!({
            "block": { "id": target_id, "type": "paragraph", "text": current }
        })),
    };
    if expected_text.is_empty() && !current.is_empty() {
        return Err(refuse(
            EditCode::InvalidInput,
            "expected_text must not be empty unless the paragraph is; include surrounding text"
                .into(),
        ));
    }
    let found: Vec<usize> = current
        .match_indices(expected_text.as_str())
        .map(|(at, _)| at)
        .collect();
    let at = match found[..] {
        [at] => at,
        [] => {
            return Err(refuse(
                EditCode::StaleTarget,
                "expected text was not found in the paragraph".into(),
            ));
        }
        _ => {
            return Err(refuse(
                EditCode::InvalidInput,
                format!(
                    "expected text occurs {} times in the paragraph; include more context",
                    found.len()
                ),
            ));
        }
    };
    let offset = utf16_len(&current[..at]);
    let mut span = changed_span(expected_text, text);
    span.start += offset;
    span.end += offset;
    let after = format!(
        "{}{text}{}",
        &current[..at],
        &current[at + expected_text.len()..]
    );
    Ok((span, after))
}

/// Whether an Office file can store `ch`: XML 1.0's `Char` (no control
/// character but tab, line feed and carriage return; no U+FFFE or U+FFFF).
fn storable(ch: char) -> bool {
    matches!(ch, '\t' | '\n' | '\r')
        || matches!(ch, ' '..='\u{d7ff}' | '\u{e000}'..='\u{fffd}' | '\u{10000}'..)
}

/// office-checkpoint.ts `storableText`: `text` without the characters an
/// Office file cannot store. An agent's text loses them before its edit
/// applies, and an Undo's inverse restores its target without them (Epo
/// 2026-10-15).
pub(crate) fn storable_text(text: &str) -> String {
    text.chars().filter(|&ch| storable(ch)).collect()
}

/// office-checkpoint.ts `storableCommand`: the command with its text or value
/// as an Office file can store it, each dropped character added to `removed`.
pub(crate) fn storable_command(command: &Command, removed: &mut Vec<char>) -> Command {
    let mut command = command.clone();
    let mut keep = |text: &mut String| {
        let before = removed.len();
        removed.extend(text.chars().filter(|&ch| !storable(ch)));
        if removed.len() > before {
            *text = storable_text(text);
        }
    };
    match &mut command {
        Command::ReplaceText { text, .. }
        | Command::SetCell { value: text, .. }
        | Command::AddSheet { name: text, .. }
        | Command::RenameSheet { name: text, .. } => keep(text),
        Command::SetRange { values, .. } => values.iter_mut().flatten().for_each(keep),
        Command::FormatCells { .. }
        | Command::DeleteSheet { .. }
        | Command::InsertRows { .. }
        | Command::DeleteRows { .. }
        | Command::InsertColumns { .. }
        | Command::DeleteColumns { .. } => {}
    }
    command
}

/// office-checkpoint.ts `removedNotice`: what an edit's tool result tells the
/// agent about the characters its text lost, each distinct one named once.
pub(crate) fn removed_notice(removed: &[char]) -> Option<String> {
    if removed.is_empty() {
        return None;
    }
    let mut names: Vec<String> = Vec::new();
    for ch in removed {
        let name = format!("U+{:04X}", *ch as u32);
        if !names.contains(&name) {
            names.push(name);
        }
    }
    let count = removed.len();
    let noun = if count == 1 {
        "character"
    } else {
        "characters"
    };
    Some(format!(
        "Removed {count} {noun} Office files cannot store ({}).",
        names.join(", ")
    ))
}

/// The span of `before` that `after` changes, in UTF-16 offsets, never
/// splitting a surrogate pair (office-checkpoint.ts `changedSpan`).
pub(crate) struct Span {
    pub start: usize,
    pub end: usize,
    pub text: String,
}

fn changed_span(before: &str, after: &str) -> Span {
    let a: Vec<u16> = before.encode_utf16().collect();
    let b: Vec<u16> = after.encode_utf16().collect();
    let limit = a.len().min(b.len());
    let mut prefix = 0;
    while prefix < limit && a[prefix] == b[prefix] {
        prefix += 1;
    }
    if prefix > 0 && (0xd800..=0xdbff).contains(&a[prefix - 1]) {
        prefix -= 1;
    }
    let mut suffix = 0;
    while suffix < limit - prefix && a[a.len() - 1 - suffix] == b[b.len() - 1 - suffix] {
        suffix += 1;
    }
    if suffix > 0 && (0xdc00..=0xdfff).contains(&a[a.len() - suffix]) {
        suffix -= 1;
    }
    Span {
        start: prefix,
        end: a.len() - suffix,
        text: String::from_utf16_lossy(&b[prefix..b.len() - suffix]),
    }
}

/// office-checkpoint.ts `unzippedBytes`: the ZIP entries' uncompressed sizes,
/// none for ZIP64 or a damaged directory.
pub(crate) fn unzipped_bytes(bytes: &[u8]) -> Option<u64> {
    let u16_at = |at: usize| u16::from_le_bytes([bytes[at], bytes[at + 1]]) as usize;
    let u32_at = |at: usize| {
        u32::from_le_bytes([bytes[at], bytes[at + 1], bytes[at + 2], bytes[at + 3]]) as usize
    };
    if bytes.len() < 22 {
        return None;
    }
    let last = bytes.len() - 22;
    let first = last.saturating_sub(0xffff);
    for end in (first..=last).rev() {
        if u32_at(end) != 0x0605_4b50 {
            continue;
        }
        let count = u16_at(end + 10);
        let mut at = u32_at(end + 16);
        if count == 0xffff || at == 0xffff_ffff {
            return None;
        }
        let mut total = 0_u64;
        for _ in 0..count {
            if at + 46 > bytes.len() || u32_at(at) != 0x0201_4b50 {
                return None;
            }
            let size = u32_at(at + 24);
            if size == 0xffff_ffff {
                return None;
            }
            total += size as u64;
            at += 46 + u16_at(at + 28) + u16_at(at + 30) + u16_at(at + 32);
        }
        return Some(total);
    }
    None
}

/// UTF-16 length as the u32 the engines take.
pub(crate) fn len16(text: &str) -> u32 {
    utf16_len(text) as u32
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn changed_span_keeps_surrogate_pairs_whole() {
        let span = changed_span("a😀b", "a😁b");
        assert_eq!((span.start, span.end, span.text.as_str()), (1, 3, "😁"));
        let span = changed_span("abc", "abc");
        assert_eq!((span.start, span.end, span.text.as_str()), (3, 3, ""));
    }

    #[test]
    fn data_urls_decode_as_node_does() {
        let asset = asset_from_data_url("data:image/png;base64,iVBORw0KGgo=").unwrap();
        assert_eq!(asset.bytes, [0x89, b'P', b'N', b'G', 13, 10, 26, 10]);
        assert!(asset_from_data_url("data:text/plain;base64,AA==").is_err());
    }
}
