//! OPC (DOCX) container read/write, compiled to WASM.
//!
//! Replaces jszip at the unzip/rezip trust boundary. The decompression budget
//! (total inflated bytes + entry count) and path-traversal rejection are
//! enforced here by construction, so a malicious `.docx` cannot exhaust memory
//! or carry out-of-tree part names into the parser. The pure `*_parts` functions
//! hold the logic and are unit-tested natively; the `#[wasm_bindgen]` wrappers
//! only marshal to/from JS `{ path: Uint8Array }` objects.

use std::collections::HashSet;
use std::io::{Cursor, Read, Write};

#[cfg(feature = "wasm")]
use js_sys::{Object, Reflect, Uint8Array};
#[cfg(feature = "wasm")]
use wasm_bindgen::prelude::*;

pub mod progress;
mod retained;
mod sanitize;
mod source_bytes;

pub use retained::{PackageBytes, RetainedPackage};
pub use source_bytes::SourceContainerBuilder;
#[doc(hidden)]
pub use source_bytes::SourceContainerInitializer;

pub use sanitize::{
    ContentTypes, DocumentKind, DocumentKindError, detect_package_kind, parse_content_types,
    sanitize_package, sanitize_package_for_format,
};

/// A well-formed document stays far under this; a decompression bomb blows past it.
/// The most bytes one package may inflate to.
pub const MAX_TOTAL_UNCOMPRESSED_BYTES: u64 = 512 * 1024 * 1024;

/// No legitimate package carries this many parts.
const MAX_ENTRY_COUNT: usize = 5000;

/// Reject absolute, drive-letter, and any `..` entry name (checked on both
/// separators, since producers may emit backslashes).
#[cfg(test)]
fn is_safe_entry_path(name: &str) -> bool {
    normalized_security_path(name).is_some()
}

/// Normalize only for security comparisons. Returned archive paths retain
/// their authored spelling; this key catches slash/case/dot aliases that could
/// otherwise name the same security-sensitive OPC part twice.
fn normalized_security_path(name: &str) -> Option<String> {
    let normalized = name.replace('\\', "/");
    if normalized.starts_with('/') {
        return None;
    }
    if normalized.len() >= 2 && normalized.as_bytes()[1] == b':' {
        return None;
    }
    let mut segments = Vec::new();
    for segment in normalized.split('/') {
        match segment {
            "" | "." => {}
            ".." => return None,
            segment => segments.push(segment),
        }
    }
    if segments.is_empty() {
        return None;
    }
    Some(segments.join("/").to_lowercase())
}

/// Extract every non-directory entry as `(path, bytes)`, enforcing the
/// decompression budget and path-traversal guard. Reading is bounded by the
/// remaining byte budget so a lying size header cannot force unbounded output.
pub fn unzip_parts(data: &[u8]) -> Result<Vec<(String, Vec<u8>)>, String> {
    unzip_parts_with_limits(data, MAX_TOTAL_UNCOMPRESSED_BYTES)
}

/// As [`unzip_parts`], with a caller-supplied expanded-data budget. The budget only
/// tightens: it is capped at the container ceiling, never raised above it.
pub fn unzip_parts_with_limits(
    data: &[u8],
    max_expanded_bytes: u64,
) -> Result<Vec<(String, Vec<u8>)>, String> {
    unzip_parts_where(data, max_expanded_bytes, |_| true)
}

/// As [`unzip_parts_with_limits`], inflating only the entries `keep` accepts.
/// Every entry's path is still checked.
pub fn unzip_parts_where(
    data: &[u8],
    max_expanded_bytes: u64,
    keep: impl Fn(&str) -> bool,
) -> Result<Vec<(String, Vec<u8>)>, String> {
    let budget = max_expanded_bytes.min(MAX_TOTAL_UNCOMPRESSED_BYTES);
    let mut archive =
        zip::ZipArchive::new(Cursor::new(data)).map_err(|e| format!("bad zip: {e}"))?;

    if archive.len() > MAX_ENTRY_COUNT {
        return Err(format!("zip entry count exceeds {MAX_ENTRY_COUNT}"));
    }

    let mut parts = Vec::with_capacity(archive.len());
    let mut total: u64 = 0;
    let mut seen_paths = HashSet::new();

    for i in 0..archive.len() {
        progress::pulse();
        let mut entry = archive
            .by_index(i)
            .map_err(|e| format!("bad zip entry: {e}"))?;
        if entry.is_dir() {
            continue;
        }
        let name = entry
            .name()
            .map_err(|e| format!("bad zip entry name: {e}"))?
            .into_owned();
        let Some(security_path) = normalized_security_path(&name) else {
            return Err(format!("unsafe zip entry path: {name}"));
        };
        if !seen_paths.insert(security_path) {
            return Err(format!("duplicate normalized zip entry path: {name}"));
        }
        if !keep(&name) {
            continue;
        }

        // read at most (budget - total) + 1 bytes: one over the limit proves a bomb
        let remaining = budget - total;
        let mut buf = Vec::new();
        entry
            .by_ref()
            .take(remaining + 1)
            .read_to_end(&mut buf)
            .map_err(|e| format!("read failed for {name}: {e}"))?;
        if buf.len() as u64 > remaining {
            return Err(format!("inflated size exceeds {budget} bytes"));
        }
        total += buf.len() as u64;
        parts.push((name, buf));
    }

    Ok(parts)
}

/// Why [`Package::read`] answered no part.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum PartError {
    /// It inflates past the bytes the read allows.
    TooLarge(String),
    /// Its entry does not read (a damaged entry, a bad checksum).
    Unreadable(String),
}

impl std::fmt::Display for PartError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::TooLarge(why) | Self::Unreadable(why) => f.write_str(why),
        }
    }
}

/// A package opened once to read several parts by path: its central
/// directory parsed and its part names normalized once, however many parts
/// are read (review r2 L3).
pub struct Package<'a> {
    archive: zip::ZipArchive<Cursor<&'a [u8]>>,
    /// Each normalized part name's first entry.
    names: std::collections::HashMap<String, usize>,
}

impl<'a> Package<'a> {
    pub fn open(data: &'a [u8]) -> Result<Self, String> {
        let archive =
            zip::ZipArchive::new(Cursor::new(data)).map_err(|e| format!("bad zip: {e}"))?;
        if archive.len() > MAX_ENTRY_COUNT {
            return Err(format!("zip entry count exceeds {MAX_ENTRY_COUNT}"));
        }
        let mut names = std::collections::HashMap::new();
        for index in 0..archive.len() {
            if let Some(name) = archive
                .name_for_index(index)
                .and_then(|name| normalized_security_path(&name.ok()?))
            {
                names.entry(name).or_insert(index);
            }
        }
        Ok(Self { archive, names })
    }

    /// One part's bytes, found by its path as OPC names parts (no leading
    /// slash, case-insensitive), without inflating the rest; `Ok(None)` when
    /// the package holds no such part. A part inflating past `max_bytes` is
    /// refused as a bomb is by [`unzip_parts`].
    pub fn read(&mut self, path: &str, max_bytes: u64) -> Result<Option<Vec<u8>>, PartError> {
        let Some(&index) = normalized_security_path(path.trim_start_matches('/'))
            .and_then(|wanted| self.names.get(&wanted))
        else {
            return Ok(None);
        };
        let budget = max_bytes.min(MAX_TOTAL_UNCOMPRESSED_BYTES);
        let mut entry = self
            .archive
            .by_index(index)
            .map_err(|e| PartError::Unreadable(format!("bad zip entry: {e}")))?;
        let mut buf = Vec::new();
        // Read no more than one byte past the budget, whatever the entry says
        // its size is.
        entry
            .by_ref()
            .take(budget + 1)
            .read_to_end(&mut buf)
            .map_err(|e| PartError::Unreadable(format!("read failed for {path}: {e}")))?;
        if buf.len() as u64 > budget {
            return Err(PartError::TooLarge(format!(
                "{path} inflates past {budget} bytes"
            )));
        }
        Ok(Some(buf))
    }
}

/// Write `(path, bytes)` entries into a deflated zip, in the given order.
pub fn rezip_parts(entries: &[(String, Vec<u8>)]) -> Result<Vec<u8>, String> {
    rezip_parts_borrowed(entries)
}

/// `rezip_parts` over borrowed entry bytes.
pub fn rezip_parts_borrowed<S: AsRef<[u8]>>(entries: &[(String, S)]) -> Result<Vec<u8>, String> {
    validate_parts(entries)?;

    let mut cursor = Cursor::new(Vec::new());
    {
        let mut writer = zip::ZipWriter::new(&mut cursor);
        for (name, bytes) in entries {
            progress::pulse();
            write_deflated(&mut writer, name, bytes.as_ref())?;
        }
        writer.finish().map_err(|e| format!("finish: {e}"))?;
    }
    Ok(cursor.into_inner())
}

fn validate_parts<S: AsRef<[u8]>>(entries: &[(String, S)]) -> Result<(), String> {
    if entries.len() > MAX_ENTRY_COUNT {
        return Err(format!("zip entry count exceeds {MAX_ENTRY_COUNT}"));
    }
    let mut seen_paths = HashSet::new();
    let mut total = 0_u64;
    for (name, bytes) in entries {
        let bytes = bytes.as_ref();
        let Some(security_path) = normalized_security_path(name) else {
            return Err(format!("unsafe zip entry path: {name}"));
        };
        if !seen_paths.insert(security_path) {
            return Err(format!("duplicate normalized zip entry path: {name}"));
        }
        total = total
            .checked_add(bytes.len() as u64)
            .ok_or_else(|| format!("inflated size exceeds {MAX_TOTAL_UNCOMPRESSED_BYTES} bytes"))?;
        if total > MAX_TOTAL_UNCOMPRESSED_BYTES {
            return Err(format!(
                "inflated size exceeds {MAX_TOTAL_UNCOMPRESSED_BYTES} bytes"
            ));
        }
    }
    Ok(())
}

fn write_deflated(
    writer: &mut zip::ZipWriter<&mut Cursor<Vec<u8>>>,
    name: &str,
    bytes: &[u8],
) -> Result<(), String> {
    writer
        .start_file(
            name,
            zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Deflated),
        )
        .map_err(|e| format!("start_file {name}: {e}"))?;
    writer
        .write_all(bytes)
        .map_err(|e| format!("write {name}: {e}"))
}

/// Source bytes retained so unchanged members re-emit verbatim on save; equality is always true.
#[derive(Clone, Default)]
pub struct SourceContainer(std::sync::Arc<[u8]>);

impl SourceContainer {
    pub fn new(bytes: Vec<u8>) -> Self {
        Self(bytes.into())
    }

    pub fn as_bytes(&self) -> &[u8] {
        &self.0
    }

    /// Whether the archive's members are exactly `names` plus directory entries, each
    /// listed once, as both a strict central-directory walk and the ZIP reader see them.
    /// False for anything either reads differently or cannot read, ZIP64 included.
    pub fn holds_exactly<'a>(&self, names: impl IntoIterator<Item = &'a str>) -> bool {
        let Some(members) = central_directory_entries(self.as_bytes()) else {
            return false;
        };
        let Ok(mut archive) = zip::ZipArchive::new(Cursor::new(self.as_bytes())) else {
            return false;
        };
        if archive.len() != members.len() {
            return false;
        }
        let mut files = HashSet::with_capacity(members.len());
        for (index, (name, header)) in members.iter().enumerate() {
            match archive.by_index_raw(index) {
                Ok(file)
                    if file.name().is_ok_and(|entry| entry == name.as_str())
                        && file.header_start() == *header => {}
                _ => return false,
            }
            if !files.insert(name.as_str()) {
                return false;
            }
        }
        files.retain(|name| !name.ends_with('/') && !name.ends_with('\\'));
        let mut count = 0;
        for name in names {
            if !files.contains(name) {
                return false;
            }
            count += 1;
        }
        count == files.len()
    }
}

/// The member names and local header offsets of a single-disk, non-ZIP64 archive whose
/// central directory ends where its only end-of-central-directory record starts.
fn central_directory_entries(bytes: &[u8]) -> Option<Vec<(String, u64)>> {
    const EOCD: u32 = 0x0605_4b50;
    const ENTRY: u32 = 0x0201_4b50;
    let u16_at = |at: usize| Some(u16::from_le_bytes(bytes.get(at..at + 2)?.try_into().ok()?));
    let u32_at = |at: usize| Some(u32::from_le_bytes(bytes.get(at..at + 4)?.try_into().ok()?));
    let last = bytes.len().checked_sub(22)?;
    let mut found = None;
    for pos in (last.saturating_sub(u16::MAX as usize)..=bytes.len() - 4).rev() {
        if u32_at(pos)? == EOCD {
            if found.is_some()
                || pos > last
                || pos + 22 + usize::from(u16_at(pos + 20)?) != bytes.len()
            {
                return None;
            }
            found = Some(pos);
        }
    }
    let eocd = found?;
    let total = u16_at(eocd + 10)?;
    let size = u32_at(eocd + 12)?;
    let offset = u32_at(eocd + 16)?;
    if u16_at(eocd + 4)? != 0
        || u16_at(eocd + 6)? != 0
        || u16_at(eocd + 8)? != total
        || total == u16::MAX
        || size == u32::MAX
        || offset == u32::MAX
        || (offset as usize).checked_add(size as usize)? != eocd
    {
        return None;
    }
    let mut at = offset as usize;
    let mut entries = Vec::with_capacity(usize::from(total));
    for _ in 0..total {
        if u32_at(at)? != ENTRY {
            return None;
        }
        let name_len = usize::from(u16_at(at + 28)?);
        let rest = usize::from(u16_at(at + 30)?) + usize::from(u16_at(at + 32)?);
        let header = u32_at(at + 42)?;
        if header == u32::MAX {
            return None;
        }
        let name = bytes.get(at + 46..at + 46 + name_len)?;
        entries.push((String::from_utf8(name.to_vec()).ok()?, u64::from(header)));
        at += 46 + name_len + rest;
    }
    (at == eocd).then_some(entries)
}

impl std::fmt::Debug for SourceContainer {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("SourceContainer")
            .field(&self.0.len())
            .finish()
    }
}

impl PartialEq for SourceContainer {
    fn eq(&self, _: &Self) -> bool {
        true
    }
}

impl Eq for SourceContainer {}

/// As [`rezip_parts`], but unchanged members copy the source's compressed
/// payload verbatim; unreadable `source` falls back to full rezip.
pub fn rezip_parts_preserving<S: AsRef<[u8]>>(
    entries: &[(String, S)],
    source: &[u8],
) -> Result<Vec<u8>, String> {
    validate_parts(entries)?;

    let Ok(mut archive) = zip::ZipArchive::new(Cursor::new(source)) else {
        return rezip_parts_borrowed(entries);
    };
    if archive.len() > MAX_ENTRY_COUNT {
        return rezip_parts_borrowed(entries);
    }

    let mut cursor = Cursor::new(Vec::new());
    {
        let mut writer = zip::ZipWriter::new(&mut cursor);
        for (name, bytes) in entries {
            progress::pulse();
            let bytes = bytes.as_ref();
            if !copy_unchanged_member(&mut archive, &mut writer, name, bytes)? {
                write_deflated(&mut writer, name, bytes)?;
            }
        }
        writer.finish().map_err(|e| format!("finish: {e}"))?;
    }
    Ok(cursor.into_inner())
}

/// Copy `name`'s compressed source member into `writer`; false means re-deflate.
fn copy_unchanged_member(
    archive: &mut zip::ZipArchive<Cursor<&[u8]>>,
    writer: &mut zip::ZipWriter<&mut Cursor<Vec<u8>>>,
    name: &str,
    bytes: &[u8],
) -> Result<bool, String> {
    let Some(index) = archive.index_for_name(name) else {
        return Ok(false);
    };
    let Ok(mut file) = archive.by_index(index) else {
        return Ok(false);
    };
    // A raw copy keeps a Unicode Path field (0x7075) and writes the effective name as the
    // new header name, so a field that did not apply when the source was read could rename
    // the part on the next read. Such a member gets a fresh header (upstream #2138).
    if file.is_dir()
        || file.size() != bytes.len() as u64
        || file
            .extra_data_fields()
            .any(|field| matches!(field, zip::ExtraField::UnicodePath(_)))
        || !inflated_matches(&mut file, bytes)
    {
        return Ok(false);
    }
    drop(file);
    let file = archive
        .by_index(index)
        .map_err(|e| format!("reopen {name}: {e}"))?;
    writer
        .raw_copy_file_rename(file, name)
        .map_err(|e| format!("copy {name}: {e}"))?;
    Ok(true)
}

fn inflated_matches(file: &mut zip::read::ZipFile<'_, Cursor<&[u8]>>, bytes: &[u8]) -> bool {
    let mut rest = bytes;
    let mut chunk = [0u8; 64 * 1024];
    loop {
        progress::tick();
        match file.read(&mut chunk) {
            Ok(0) => return rest.is_empty(),
            Ok(n) => {
                if rest.len() < n || rest[..n] != chunk[..n] {
                    return false;
                }
                rest = &rest[n..];
            }
            Err(_) => return false,
        }
    }
}

/// Unzip a DOCX; returns a JS object `{ [path]: Uint8Array }`.
#[cfg(feature = "wasm")]
#[wasm_bindgen]
pub fn unzip_docx(data: &[u8]) -> Result<JsValue, JsValue> {
    let parts = unzip_parts(data).map_err(|e| JsValue::from_str(&e))?;
    let out = Object::new();
    for (name, bytes) in parts {
        let arr = Uint8Array::from(bytes.as_slice());
        Reflect::set(&out, &JsValue::from_str(&name), &arr)?;
    }
    Ok(out.into())
}

/// Rezip from a JS object `{ [path]: Uint8Array }` into a DOCX byte array.
#[cfg(feature = "wasm")]
#[wasm_bindgen]
pub fn rezip_docx(entries: JsValue) -> Result<Vec<u8>, JsValue> {
    let obj: Object = entries
        .dyn_into()
        .map_err(|_| JsValue::from_str("rezip_docx: expected an object"))?;
    let mut collected: Vec<(String, Vec<u8>)> = Vec::new();
    let keys = Object::keys(&obj);
    for key in keys.iter() {
        let name = key
            .as_string()
            .ok_or_else(|| JsValue::from_str("rezip_docx: non-string key"))?;
        let value = Reflect::get(&obj, &key)?;
        let arr = Uint8Array::new(&value);
        collected.push((name, arr.to_vec()));
    }
    rezip_parts(&collected).map_err(|e| JsValue::from_str(&e))
}

#[cfg(feature = "wasm")]
#[wasm_bindgen(js_name = sanitizeOoxml)]
pub fn sanitize_ooxml(data: &[u8], expected_format: &str) -> Result<Vec<u8>, JsValue> {
    sanitize_package_for_format(data, expected_format).map_err(|error| JsValue::from_str(&error))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn sample() -> Vec<(String, Vec<u8>)> {
        vec![
            ("[Content_Types].xml".into(), b"<Types/>".to_vec()),
            ("word/document.xml".into(), b"<w:document/>".to_vec()),
            ("word/media/image1.png".into(), vec![0x89, 0x50, 0x4e, 0x47]),
        ]
    }

    #[test]
    fn round_trips_parts() {
        let zipped = rezip_parts(&sample()).expect("rezip");
        let back = unzip_parts(&zipped).expect("unzip");
        assert_eq!(back, sample());
    }

    fn names(parts: &[(String, Vec<u8>)]) -> Vec<&str> {
        parts.iter().map(|(name, _)| name.as_str()).collect()
    }

    fn with_comment(mut bytes: Vec<u8>, comment: &[u8]) -> Vec<u8> {
        let end = bytes.len();
        bytes[end - 2..].copy_from_slice(&(comment.len() as u16).to_le_bytes());
        bytes.extend_from_slice(comment);
        bytes
    }

    #[test]
    fn source_container_holds_exactly_its_parts() {
        let parts = sample();
        let source = SourceContainer::new(rezip_parts(&parts).unwrap());
        assert!(source.holds_exactly(names(&parts)));
        assert!(!source.holds_exactly(names(&parts[1..])));
        let mut extra = names(&parts);
        extra.push("word/extra.xml");
        assert!(!source.holds_exactly(extra));
    }

    #[test]
    fn source_container_ignores_directory_entries_and_archive_comments() {
        let parts = sample();
        let mut cursor = Cursor::new(Vec::new());
        {
            let mut writer = zip::ZipWriter::new(&mut cursor);
            writer
                .add_directory("word/", zip::write::SimpleFileOptions::default())
                .unwrap();
            for (name, bytes) in &parts {
                writer
                    .start_file(name, zip::write::SimpleFileOptions::default())
                    .unwrap();
                writer.write_all(bytes).unwrap();
            }
            writer.finish().unwrap();
        }
        let bytes = cursor.into_inner();
        assert!(SourceContainer::new(bytes.clone()).holds_exactly(names(&parts)));
        for comment in [b"archive comment".to_vec(), vec![0; u16::MAX as usize]] {
            let source = SourceContainer::new(with_comment(bytes.clone(), &comment));
            assert!(source.holds_exactly(names(&parts)));
        }
        let source = SourceContainer::new(with_comment(bytes, b"comment PK\x05\x06 signature"));
        assert!(!source.holds_exactly(names(&parts)));
    }

    #[test]
    fn source_container_refuses_ambiguous_or_unreadable_archives() {
        let parts = sample();
        let bytes = rezip_parts(&parts).unwrap();
        let mut decoy = bytes[bytes.len() - 22..].to_vec();
        decoy[12..20].fill(0xfe);
        let mut truncated = bytes.clone();
        truncated.pop();
        let mut trailing = bytes.clone();
        trailing.push(0);
        let mut zip64 = bytes.clone();
        let end = zip64.len() - 22;
        zip64[end + 16..end + 20].fill(0xff);
        for bytes in [
            with_comment(bytes.clone(), &decoy),
            truncated,
            trailing,
            zip64,
            Vec::new(),
            b"not a zip".to_vec(),
        ] {
            assert!(!SourceContainer::new(bytes).holds_exactly(names(&parts)));
        }
    }

    #[test]
    fn rejects_traversal_paths() {
        assert!(!is_safe_entry_path("../evil.xml"));
        assert!(!is_safe_entry_path("word/../../etc/passwd"));
        assert!(!is_safe_entry_path("/etc/passwd"));
        assert!(!is_safe_entry_path("C:/windows"));
        assert!(!is_safe_entry_path("word\\..\\..\\x"));
        assert!(is_safe_entry_path("word/document.xml"));
        assert!(is_safe_entry_path("word/my..file.xml"));
        assert!(!is_safe_entry_path("."));
    }

    #[test]
    fn rejects_traversal_on_read_and_write() {
        assert!(rezip_parts(&[("../escape.xml".into(), b"x".to_vec())]).is_err());

        // Bypass rezip_parts to exercise the independent read-side guard.
        let mut cursor = Cursor::new(Vec::new());
        {
            let mut writer = zip::ZipWriter::new(&mut cursor);
            writer
                .start_file("../escape.xml", zip::write::SimpleFileOptions::default())
                .unwrap();
            writer.write_all(b"x").unwrap();
            writer.finish().unwrap();
        }
        let zipped = cursor.into_inner();
        let err = unzip_parts(&zipped).unwrap_err();
        assert!(err.contains("unsafe"));
    }

    #[test]
    fn rejects_duplicate_normalized_paths_on_read_and_write() {
        let entries = vec![
            ("word/document.xml".into(), b"a".to_vec()),
            ("WORD//./document.xml".into(), b"b".to_vec()),
        ];
        assert!(rezip_parts(&entries).unwrap_err().contains("duplicate"));

        let mut cursor = Cursor::new(Vec::new());
        {
            let mut writer = zip::ZipWriter::new(&mut cursor);
            for (name, bytes) in &entries {
                writer
                    .start_file(name, zip::write::SimpleFileOptions::default())
                    .unwrap();
                writer.write_all(bytes).unwrap();
            }
            writer.finish().unwrap();
        }
        assert!(
            unzip_parts(&cursor.into_inner())
                .unwrap_err()
                .contains("duplicate")
        );
    }

    /// Security regression (Q240 review r1 M4, r2 M1, r4 L1): one part's
    /// read stops at its byte budget however small the entry compresses: an
    /// entry whose checksum is wrong answers "inflates past" because the read
    /// never reaches its end (read whole, it fails its checksum), finds a
    /// part by its normalized name only (stored names normalized too, so an
    /// entry stored in other case or with backslashes is found by its OPC
    /// name), and never answers a traversal or drive name, requested or
    /// stored.
    #[test]
    fn a_part_read_holds_its_budget_and_refuses_traversal_names() {
        let mut cursor = Cursor::new(Vec::new());
        {
            let mut writer = zip::ZipWriter::new(&mut cursor);
            for (name, bytes) in [
                ("word/media/zeros.bin", vec![0_u8; 1 << 20]),
                ("word/media/ones.bin", vec![1_u8; 1 << 20]),
                ("Word/Media/Upper.BIN", b"upper".to_vec()),
                ("word\\media\\back.bin", b"back".to_vec()),
                ("../escape.xml", b"x".to_vec()),
            ] {
                writer
                    .start_file(name, zip::write::SimpleFileOptions::default())
                    .unwrap();
                writer.write_all(&bytes).unwrap();
            }
            writer.finish().unwrap();
        }
        let mut zipped = cursor.into_inner();
        assert!(zipped.len() < 1 << 14, "{}", zipped.len());
        // ones.bin's stored checksum, wherever the zip records it, made wrong.
        let crc = zip::ZipArchive::new(Cursor::new(&zipped))
            .unwrap()
            .by_name("word/media/ones.bin")
            .unwrap()
            .crc32()
            .to_le_bytes();
        let places: Vec<usize> = (0..zipped.len() - 3)
            .filter(|&at| zipped[at..at + 4] == crc)
            .collect();
        assert!(places.len() >= 2, "{places:?}");
        for at in places {
            zipped[at] ^= 0xff;
        }
        let read = |name: &str, max_bytes: u64| {
            Package::open(&zipped)
                .unwrap()
                .read(name, max_bytes)
                .map_err(|error| error.to_string())
        };
        let whole = read("word/media/ones.bin", 2 << 20).unwrap_err();
        assert!(whole.contains("checksum"), "{whole}");
        let error = read("word/media/ones.bin", 1_000).unwrap_err();
        assert!(error.contains("inflates past 1000 bytes"), "{error}");
        assert_eq!(
            read("/WORD/media/./zeros.bin", 1 << 20)
                .unwrap()
                .map(|bytes| bytes.len()),
            Some(1 << 20)
        );
        assert_eq!(
            read("word/media/upper.bin", 1 << 20).unwrap(),
            Some(b"upper".to_vec())
        );
        assert_eq!(
            read("word/media/back.bin", 1 << 20).unwrap(),
            Some(b"back".to_vec())
        );
        for name in [
            "escape.xml",
            "../escape.xml",
            "C:/escape.xml",
            "word/missing.xml",
        ] {
            assert_eq!(read(name, 1 << 20).unwrap(), None, "{name}");
        }
    }

    /// One opened package reads several parts by normalized name, and says
    /// whether a part it refuses is too large or unreadable (review r2 L3,
    /// N4).
    #[test]
    fn a_package_reads_several_parts_from_one_open() {
        let zipped = rezip_parts(&sample()).unwrap();
        let mut package = Package::open(&zipped).unwrap();
        assert_eq!(
            package.read("/Word/Document.xml", 100).unwrap(),
            Some(b"<w:document/>".to_vec())
        );
        assert_eq!(
            package.read("word/media/image1.png", 100).unwrap(),
            Some(vec![0x89, 0x50, 0x4e, 0x47])
        );
        assert!(matches!(
            package.read("word/document.xml", 3),
            Err(PartError::TooLarge(_))
        ));
        assert_eq!(package.read("word/none.xml", 100).unwrap(), None);
    }

    fn stored_zip(entries: &[(String, Vec<u8>)], stored: &str) -> Vec<u8> {
        let mut cursor = Cursor::new(Vec::new());
        {
            let mut writer = zip::ZipWriter::new(&mut cursor);
            for (name, bytes) in entries {
                let options = zip::write::SimpleFileOptions::default().compression_method(
                    if name == stored {
                        zip::CompressionMethod::Stored
                    } else {
                        zip::CompressionMethod::Deflated
                    },
                );
                writer.start_file(name, options).unwrap();
                writer.write_all(bytes).unwrap();
            }
            writer.finish().unwrap();
        }
        cursor.into_inner()
    }

    fn entry_method(zip_bytes: &[u8], name: &str) -> zip::CompressionMethod {
        let mut archive = zip::ZipArchive::new(Cursor::new(zip_bytes)).unwrap();
        archive.by_name(name).unwrap().compression()
    }

    #[test]
    fn preserving_copies_verbatim_and_keeps_methods() {
        let entries = sample();
        let source = stored_zip(&entries, "word/media/image1.png");
        let out = rezip_parts_preserving(&entries, &source).expect("rezip");
        assert_eq!(unzip_parts(&out).unwrap(), entries);
        assert_eq!(
            entry_method(&out, "word/media/image1.png"),
            zip::CompressionMethod::Stored
        );
        assert_eq!(
            entry_method(&out, "word/document.xml"),
            zip::CompressionMethod::Deflated
        );
    }

    #[test]
    fn preserving_deflates_changed_and_new_parts() {
        let source = stored_zip(&sample(), "word/media/image1.png");
        let mut entries = sample();
        entries[1].1 = b"<w:document>edited</w:document>".to_vec();
        entries.push(("word/extra.xml".into(), b"<x/>".to_vec()));
        let out = rezip_parts_preserving(&entries, &source).expect("rezip");
        assert_eq!(unzip_parts(&out).unwrap(), entries);
        assert_eq!(
            entry_method(&out, "word/document.xml"),
            zip::CompressionMethod::Deflated
        );
    }

    #[test]
    fn preserving_rejects_same_size_forgery() {
        let source = stored_zip(&sample(), "word/media/image1.png");
        let mut entries = sample();
        // Same length, different content: must re-deflate, not copy.
        entries[2].1 = vec![0x89, 0x50, 0x4e, 0x00];
        let out = rezip_parts_preserving(&entries, &source).expect("rezip");
        assert_eq!(unzip_parts(&out).unwrap(), entries);
    }

    fn crc32(bytes: &[u8]) -> u32 {
        let mut crc = !0u32;
        for byte in bytes {
            crc ^= u32::from(*byte);
            for _ in 0..8 {
                crc = (crc >> 1) ^ (0xedb8_8320 & (crc & 1).wrapping_neg());
            }
        }
        !crc
    }

    fn unicode_path(name_raw: &[u8], path: &str) -> Vec<u8> {
        let mut field = vec![0x75, 0x70];
        field.extend_from_slice(&(5 + path.len() as u16).to_le_bytes());
        field.push(1);
        field.extend_from_slice(&crc32(name_raw).to_le_bytes());
        field.extend_from_slice(path.as_bytes());
        field
    }

    fn raw_stored_zip(name: &[u8], extra: &[u8], data: &[u8]) -> Vec<u8> {
        let crc = crc32(data).to_le_bytes();
        let size = (data.len() as u32).to_le_bytes();
        let lengths = [
            (name.len() as u16).to_le_bytes(),
            (extra.len() as u16).to_le_bytes(),
        ]
        .concat();
        let mut zip = [
            &[0x50, 0x4b, 3, 4, 20, 0, 0, 0, 0, 0, 0, 0, 0x21, 0][..],
            &crc,
            &size,
            &size,
            &lengths,
        ]
        .concat();
        zip.extend_from_slice(name);
        zip.extend_from_slice(extra);
        zip.extend_from_slice(data);
        let central_start = zip.len() as u32;
        zip.extend_from_slice(&[0x50, 0x4b, 1, 2, 20, 0, 20, 0, 0, 0, 0, 0, 0, 0, 0x21, 0]);
        zip.extend_from_slice(&[&crc[..], &size, &size, &lengths, &[0; 14]].concat());
        zip.extend_from_slice(name);
        zip.extend_from_slice(extra);
        let central_size = zip.len() as u32 - central_start;
        zip.extend_from_slice(&[0x50, 0x4b, 5, 6, 0, 0, 0, 0, 1, 0, 1, 0]);
        zip.extend_from_slice(&central_size.to_le_bytes());
        zip.extend_from_slice(&central_start.to_le_bytes());
        zip.extend_from_slice(&[0, 0]);
        zip
    }

    fn effective_names(zip_bytes: &[u8]) -> Vec<String> {
        let mut archive = zip::ZipArchive::new(Cursor::new(zip_bytes)).unwrap();
        (0..archive.len())
            .map(|index| {
                archive
                    .by_index_raw(index)
                    .unwrap()
                    .name()
                    .unwrap()
                    .into_owned()
            })
            .collect()
    }

    // Security regression: zip 9 refuses to write these Info-ZIP field ids without its
    // `unreserved` feature, so a document holding them could not be saved at all.
    #[test]
    fn preserving_saves_members_holding_info_zip_extra_fields() {
        // 0x7875 unix uid/gid and 0x554e, the shapes real xlsx and docx files carry.
        let unix = [0x75, 0x78, 11, 0, 1, 4, 0xe8, 3, 0, 0, 4, 0xe8, 3, 0, 0];
        let nu = [0x4e, 0x55, 6, 0, 0x2e, 0, b'x', b'x', b'x', b'x'];
        for extra in [&unix[..], &nu[..]] {
            let source = raw_stored_zip(b"word/document.xml", extra, b"<x/>");
            let parts = unzip_parts(&source).expect("unzip");
            let out = rezip_parts_preserving(&parts, &source).expect("rezip");
            assert_eq!(unzip_parts(&out).expect("reopen"), parts);
            // Copied, so the member keeps its stored bytes and its field.
            assert_eq!(
                entry_method(&out, "word/document.xml"),
                zip::CompressionMethod::Stored
            );
            assert!(out.windows(extra.len()).any(|bytes| bytes == extra));
        }
    }

    // Security regression: a raw copy keeps the field and writes the effective name, so a
    // chain that reads as safe.xml would save as ../escape.xml, whichever order it is in.
    #[test]
    fn preserving_never_lets_a_copied_unicode_path_rename_a_part() {
        let first = unicode_path(b"old.xml", "safe.xml");
        let second = unicode_path(b"safe.xml", "../escape.xml");
        for extra in [
            [first.clone(), second.clone()].concat(),
            [second, first].concat(),
        ] {
            let source = raw_stored_zip(b"old.xml", &extra, b"<x/>");
            let parts = unzip_parts(&source).expect("unzip");
            assert_eq!(names(&parts), ["safe.xml"]);
            let out = rezip_parts_preserving(&parts, &source).expect("rezip");
            assert_eq!(effective_names(&out), ["safe.xml"]);
            assert_eq!(unzip_parts(&out).expect("reopen"), parts);
        }
    }

    #[test]
    fn saved_packages_reopen_with_the_names_they_were_given() {
        let mut cases = vec![(sample(), stored_zip(&sample(), "word/media/image1.png"))];
        for target in [
            "../escape.xml",
            "/escape.xml",
            "word\\..\\..\\escape.xml",
            "C:/escape.xml",
            "word/escape.xml",
        ] {
            let mut extra = unicode_path(b"old.xml", "safe.xml");
            extra.extend(unicode_path(b"safe.xml", target));
            let source = raw_stored_zip(b"old.xml", &extra, b"<x/>");
            cases.push((unzip_parts(&source).expect("unzip"), source));
        }
        let renamed = raw_stored_zip(
            b"old.xml",
            &unicode_path(b"old.xml", "word/renamed.xml"),
            b"<x/>",
        );
        cases.push((unzip_parts(&renamed).expect("unzip"), renamed));
        for (parts, source) in cases {
            let intended: Vec<&str> = names(&parts);
            for saved in [
                rezip_parts(&parts).expect("rezip"),
                rezip_parts_preserving(&parts, &source).expect("rezip preserving"),
            ] {
                let reopened = effective_names(&saved);
                assert_eq!(reopened, intended);
                assert!(
                    reopened
                        .iter()
                        .all(|name| is_safe_entry_path(name) && !name.contains('\\'))
                );
                let retained = RetainedPackage::new(std::sync::Arc::from(saved)).expect("retained");
                assert_eq!(
                    retained.parts().map(|(name, _)| name).collect::<Vec<_>>(),
                    intended
                );
            }
        }
    }

    // A Unicode Path that applies is the name every reader checks: it cannot carry a part
    // past the path guard on the way in.
    #[test]
    fn every_reader_checks_the_name_a_unicode_path_gives() {
        let source = raw_stored_zip(
            b"old.xml",
            &unicode_path(b"old.xml", "../escape.xml"),
            b"<x/>",
        );
        assert!(unzip_parts(&source).unwrap_err().contains("unsafe"));
        let retained = RetainedPackage::new(std::sync::Arc::from(source.clone()));
        assert!(retained.err().is_some_and(|e| e.contains("unsafe")));
        let mut package = Package::open(&source).expect("open");
        assert_eq!(package.read("old.xml", 1024), Ok(None));
    }

    #[test]
    fn preserving_falls_back_when_source_is_not_a_zip() {
        let out = rezip_parts_preserving(&sample(), b"not a zip").expect("rezip");
        assert_eq!(unzip_parts(&out).unwrap(), sample());
    }

    #[test]
    fn preserving_still_enforces_entry_guards() {
        assert!(rezip_parts_preserving(&[("../x".into(), b"x".to_vec())], b"").is_err());
        let entries = vec![
            ("word/document.xml".into(), b"a".to_vec()),
            ("WORD//./document.xml".into(), b"b".to_vec()),
        ];
        assert!(
            rezip_parts_preserving(&entries, b"")
                .unwrap_err()
                .contains("duplicate")
        );
    }

    #[test]
    fn caller_budget_only_tightens_the_expanded_data_ceiling() {
        let mut entries = sample();
        entries.push(("word/media/filler.bin".into(), vec![0; 4 * 1024 * 1024]));
        let zipped = rezip_parts(&entries).expect("rezip");
        assert!(zipped.len() < 64 * 1024, "fixture must stay compressible");

        assert!(
            unzip_parts_with_limits(&zipped, 1024 * 1024)
                .unwrap_err()
                .contains("inflated size exceeds")
        );
        assert_eq!(
            unzip_parts_with_limits(&zipped, MAX_TOTAL_UNCOMPRESSED_BYTES).unwrap(),
            entries
        );
        assert_eq!(unzip_parts_with_limits(&zipped, u64::MAX).unwrap(), entries);
    }
}
