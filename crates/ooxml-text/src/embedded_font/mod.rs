//! Embedded font parts (`ppt/fonts/*.fntdata`, `*.odttf`) decoded to plain
//! OpenType bytes: an sfnt as is, an Embedded OpenType container (with its
//! XOR encryption and MicroType Express compression, which Google Slides and
//! PowerPoint write), or a part whose first 32 bytes are obfuscated with the
//! GUID its file name carries (ECMA-376 Part 2, as `.odttf`).
//!
//! MicroType Express drops the TrueType device tables `hdmx` and `VDMX`
//! (layout and paint never read them); every other table, outlines and glyph
//! programs included, comes back as the font had it.

mod ctf;
mod lzcomp;

/// Largest font one embedded part may decode to.
pub const MAX_EMBEDDED_FONT_BYTES: usize = 32 * 1024 * 1024;

const EOT_MIN_HEADER: usize = 82;
const EOT_MAGIC: u16 = 0x504C;
const EOT_COMPRESSED: u32 = 0x4;
const EOT_XOR_ENCRYPTED: u32 = 0x1000_0000;
const OBFUSCATED_PREFIX: usize = 32;

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum EmbeddedFontError {
    /// Neither an sfnt, an EOT container nor an obfuscated sfnt.
    Unrecognized,
    /// The data ends inside a structure it declares.
    Truncated,
    /// A structure holds a value the format does not allow.
    Malformed(&'static str),
    /// The decoded font would exceed [`MAX_EMBEDDED_FONT_BYTES`].
    TooLarge,
}

impl std::fmt::Display for EmbeddedFontError {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Unrecognized => formatter.write_str("not a font, EOT or obfuscated font"),
            Self::Truncated => formatter.write_str("embedded font data is truncated"),
            Self::Malformed(what) => write!(formatter, "malformed embedded font: {what}"),
            Self::TooLarge => write!(
                formatter,
                "embedded font exceeds {MAX_EMBEDDED_FONT_BYTES} bytes"
            ),
        }
    }
}

impl std::error::Error for EmbeddedFontError {}

/// Decodes the bytes of the embedded font part `part_name` into an sfnt.
pub fn decode_embedded_font(bytes: &[u8], part_name: &str) -> Result<Vec<u8>, EmbeddedFontError> {
    if bytes.len() > MAX_EMBEDDED_FONT_BYTES {
        return Err(EmbeddedFontError::TooLarge);
    }
    if is_sfnt(bytes) {
        return Ok(bytes.to_vec());
    }
    if is_eot(bytes) {
        return decode_eot(bytes);
    }
    let key = obfuscation_key(part_name).ok_or(EmbeddedFontError::Unrecognized)?;
    let mut clear = bytes.to_vec();
    for (index, byte) in clear.iter_mut().take(OBFUSCATED_PREFIX).enumerate() {
        *byte ^= key[index % key.len()];
    }
    if is_sfnt(&clear) {
        Ok(clear)
    } else if is_eot(&clear) {
        decode_eot(&clear)
    } else {
        Err(EmbeddedFontError::Unrecognized)
    }
}

fn is_sfnt(bytes: &[u8]) -> bool {
    bytes.len() >= 12 && matches!(&bytes[..4], [0x00, 0x01, 0x00, 0x00] | b"OTTO" | b"true")
}

fn is_eot(bytes: &[u8]) -> bool {
    bytes.len() >= EOT_MIN_HEADER && u16::from_le_bytes([bytes[34], bytes[35]]) == EOT_MAGIC
}

fn decode_eot(bytes: &[u8]) -> Result<Vec<u8>, EmbeddedFontError> {
    let read = |offset: usize| {
        u32::from_le_bytes([
            bytes[offset],
            bytes[offset + 1],
            bytes[offset + 2],
            bytes[offset + 3],
        ])
    };
    if read(0) as usize != bytes.len() {
        return Err(EmbeddedFontError::Malformed(
            "EOT size differs from the part",
        ));
    }
    if !matches!(read(8), 0x0001_0000 | 0x0002_0001 | 0x0002_0002) {
        return Err(EmbeddedFontError::Malformed("unknown EOT version"));
    }
    let data_size = read(4) as usize;
    if data_size > bytes.len() - EOT_MIN_HEADER {
        return Err(EmbeddedFontError::Malformed(
            "EOT font data overruns the part",
        ));
    }
    let flags = read(12);
    // The font data is the container's last field.
    let mut data = bytes[bytes.len() - data_size..].to_vec();
    if flags & EOT_XOR_ENCRYPTED != 0 {
        for byte in &mut data {
            *byte ^= 0x50;
        }
    }
    let font = if flags & EOT_COMPRESSED != 0 {
        decode_mtx(&data)?
    } else {
        data
    };
    if !is_sfnt(&font) {
        return Err(EmbeddedFontError::Malformed("EOT font data is not an sfnt"));
    }
    Ok(font)
}

/// MicroType Express: three LZCOMP blocks holding the font in Compact Table
/// Format, its glyphs' pushed values and their remaining instructions.
fn decode_mtx(data: &[u8]) -> Result<Vec<u8>, EmbeddedFontError> {
    if data.len() < 10 {
        return Err(EmbeddedFontError::Truncated);
    }
    let version = data[0];
    if !(1..=3).contains(&version) {
        return Err(EmbeddedFontError::Malformed("unknown MTX version"));
    }
    let u24 = |offset: usize| {
        (usize::from(data[offset]) << 16)
            | (usize::from(data[offset + 1]) << 8)
            | usize::from(data[offset + 2])
    };
    let (second, third) = (u24(4), u24(7));
    if !(10 <= second && second <= third && third <= data.len()) {
        return Err(EmbeddedFontError::Malformed("MTX block offsets"));
    }
    let mut budget = MAX_EMBEDDED_FONT_BYTES;
    let mut unpack = |block: &[u8]| {
        let bytes = lzcomp::unpack(block, version, budget)?;
        budget -= bytes.len();
        Ok::<_, EmbeddedFontError>(bytes)
    };
    let rest = unpack(&data[10..second])?;
    let push = unpack(&data[second..third])?;
    let code = unpack(&data[third..])?;
    ctf::to_sfnt(&rest, &push, &code)
}

/// The XOR key a GUID file name (`{0123…}.odttf`) stands for: its 16 bytes
/// in reverse order.
fn obfuscation_key(part_name: &str) -> Option<[u8; 16]> {
    let file = part_name.rsplit('/').next()?;
    let stem = file.split('.').next()?;
    let hex: Vec<u8> = stem
        .bytes()
        .filter(|byte| !matches!(byte, b'{' | b'}' | b'-'))
        .collect();
    if hex.len() != 32 {
        return None;
    }
    let mut key = [0_u8; 16];
    for (index, pair) in hex.chunks(2).enumerate() {
        let text = std::str::from_utf8(pair).ok()?;
        key[15 - index] = u8::from_str_radix(text, 16).ok()?;
    }
    Some(key)
}

/// Big-endian reader over one block, failing with `Truncated` past its end.
struct Reader<'a> {
    data: &'a [u8],
    position: usize,
}

impl<'a> Reader<'a> {
    fn new(data: &'a [u8]) -> Self {
        Self { data, position: 0 }
    }

    fn bytes(&mut self, count: usize) -> Result<&'a [u8], EmbeddedFontError> {
        let end = self
            .position
            .checked_add(count)
            .filter(|end| *end <= self.data.len())
            .ok_or(EmbeddedFontError::Truncated)?;
        let bytes = &self.data[self.position..end];
        self.position = end;
        Ok(bytes)
    }

    fn u8(&mut self) -> Result<u8, EmbeddedFontError> {
        Ok(self.bytes(1)?[0])
    }

    fn u16(&mut self) -> Result<u16, EmbeddedFontError> {
        let bytes = self.bytes(2)?;
        Ok(u16::from_be_bytes([bytes[0], bytes[1]]))
    }

    fn i16(&mut self) -> Result<i16, EmbeddedFontError> {
        Ok(self.u16()? as i16)
    }

    /// `255USHORT`.
    fn u16_255(&mut self) -> Result<u16, EmbeddedFontError> {
        Ok(match self.u8()? {
            253 => self.u16()?,
            255 => 253 + u16::from(self.u8()?),
            254 => 506 + u16::from(self.u8()?),
            code => u16::from(code),
        })
    }

    /// `255SHORT`.
    fn i16_255(&mut self) -> Result<i16, EmbeddedFontError> {
        let mut code = self.u8()?;
        if code == 253 {
            return self.i16();
        }
        let negative = code == 250;
        if negative {
            code = self.u8()?;
        }
        let value = match code {
            255 => 250 + i16::from(self.u8()?),
            254 => 500 + i16::from(self.u8()?),
            code => i16::from(code),
        };
        Ok(if negative { -value } else { value })
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn a_guid_file_name_gives_the_reversed_key() {
        assert_eq!(
            obfuscation_key("ppt/fonts/{001B70DC-AA60-4AD5-90EC-18A0948E1EAE}.odttf"),
            Some([
                0xAE, 0x1E, 0x8E, 0x94, 0xA0, 0x18, 0xEC, 0x90, 0xD5, 0x4A, 0x60, 0xAA, 0xDC, 0x70,
                0x1B, 0x00
            ])
        );
        assert_eq!(obfuscation_key("ppt/fonts/font1.fntdata"), None);
    }

    #[test]
    fn unknown_bytes_are_refused() {
        for bytes in [&b""[..], b"not a font at all", &[0_u8; 200]] {
            assert_eq!(
                decode_embedded_font(bytes, "ppt/fonts/font1.fntdata"),
                Err(EmbeddedFontError::Unrecognized)
            );
        }
    }

    #[test]
    fn an_eot_whose_size_field_disagrees_is_refused() {
        let mut eot = vec![0_u8; 120];
        eot[34..36].copy_from_slice(&EOT_MAGIC.to_le_bytes());
        eot[0..4].copy_from_slice(&121_u32.to_le_bytes());
        assert!(matches!(
            decode_embedded_font(&eot, "ppt/fonts/font1.fntdata"),
            Err(EmbeddedFontError::Malformed(_))
        ));
    }

    #[test]
    fn short_numbers_read_their_255_forms() {
        let data = [7, 255, 3, 254, 0, 253, 0x12, 0x34];
        let mut reader = Reader::new(&data);
        let values: Vec<u16> = (0..4).map(|_| reader.u16_255().unwrap()).collect();
        assert_eq!(values, [7, 256, 506, 0x1234]);
        let data = [250, 255, 0, 249, 253, 0xFF, 0xFE, 250, 3];
        let mut reader = Reader::new(&data);
        let values: Vec<i16> = (0..4).map(|_| reader.i16_255().unwrap()).collect();
        assert_eq!(values, [-250, 249, -2, -3]);
    }
}
