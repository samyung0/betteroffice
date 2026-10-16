//! UTF-16 code-unit lengths.

/// Returns a string's UTF-16 code-unit length.
pub fn utf16_len(s: &str) -> usize {
    s.encode_utf16().count()
}
