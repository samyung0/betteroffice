//! Compact Table Format back to a TrueType sfnt: `glyf` from its compact
//! glyphs with their programs rejoined from the push and code blocks, `loca`
//! rebuilt, `cvt ` from its deltas; `hdmx` and `VDMX` dropped.

use super::{EmbeddedFontError, MAX_EMBEDDED_FONT_BYTES, Reader};

const MAX_TABLES: usize = 256;
const HOP3: u8 = 251;
const HOP4: u8 = 252;

const ARG_1_AND_2_ARE_WORDS: u16 = 0x0001;
const WE_HAVE_A_SCALE: u16 = 0x0008;
const MORE_COMPONENTS: u16 = 0x0020;
const WE_HAVE_AN_X_AND_Y_SCALE: u16 = 0x0040;
const WE_HAVE_A_TWO_BY_TWO: u16 = 0x0080;
const WE_HAVE_INSTRUCTIONS: u16 = 0x0100;

pub(super) fn to_sfnt(rest: &[u8], push: &[u8], code: &[u8]) -> Result<Vec<u8>, EmbeddedFontError> {
    let mut directory = Reader::new(rest);
    let version = directory.bytes(4)?;
    let count = usize::from(directory.u16()?);
    if count > MAX_TABLES {
        return Err(EmbeddedFontError::Malformed("CTF table count"));
    }
    directory.bytes(6)?;
    let mut tables = Vec::with_capacity(count);
    for _ in 0..count {
        let tag: [u8; 4] = directory.bytes(4)?.try_into().expect("four bytes");
        directory.bytes(4)?;
        let offset = directory.bytes(4)?;
        let length = directory.bytes(4)?;
        let offset = u32::from_be_bytes(offset.try_into().expect("four bytes")) as usize;
        let length = u32::from_be_bytes(length.try_into().expect("four bytes")) as usize;
        let data = offset
            .checked_add(length)
            .and_then(|end| rest.get(offset..end))
            .ok_or(EmbeddedFontError::Truncated)?;
        tables.push((tag, data));
    }
    let table = |tag: &[u8; 4]| {
        tables
            .iter()
            .find(|(candidate, _)| candidate == tag)
            .map(|(_, data)| *data)
    };

    let mut out: Vec<([u8; 4], Vec<u8>)> = Vec::with_capacity(count);
    let glyphs = match table(b"glyf") {
        Some(glyf) => {
            let maxp = table(b"maxp").ok_or(EmbeddedFontError::Malformed("CTF without maxp"))?;
            let glyph_count = Reader::new(maxp.get(4..).unwrap_or_default()).u16()?;
            Some(decode_glyphs(glyf, push, code, glyph_count)?)
        }
        None => None,
    };
    for (tag, data) in &tables {
        let data = match (tag, &glyphs) {
            (b"hdmx" | b"VDMX", _) => continue,
            (b"glyf", Some(glyphs)) => glyphs.glyf.clone(),
            (b"loca", Some(glyphs)) => glyphs.loca.clone(),
            (b"loca", None) => continue,
            (b"cvt ", _) => decode_cvt(data)?,
            (b"head", Some(_)) => {
                let mut head = data.to_vec();
                // Long `loca` offsets.
                head.get_mut(50..52)
                    .ok_or(EmbeddedFontError::Truncated)?
                    .copy_from_slice(&1_u16.to_be_bytes());
                head
            }
            (b"maxp", Some(glyphs)) if data.len() >= 32 => {
                let mut maxp = data.to_vec();
                let declared = u16::from_be_bytes([maxp[26], maxp[27]]);
                let largest = declared.max(glyphs.max_instructions);
                maxp[26..28].copy_from_slice(&largest.to_be_bytes());
                maxp
            }
            _ => data.to_vec(),
        };
        out.push((*tag, data));
    }
    if glyphs.is_some() && !out.iter().any(|(tag, _)| tag == b"loca") {
        let loca = glyphs
            .as_ref()
            .map(|glyphs| glyphs.loca.clone())
            .unwrap_or_default();
        out.push((*b"loca", loca));
    }
    assemble(version, out)
}

struct Glyphs {
    glyf: Vec<u8>,
    loca: Vec<u8>,
    max_instructions: u16,
}

fn decode_glyphs(
    glyf: &[u8],
    push: &[u8],
    code: &[u8],
    count: u16,
) -> Result<Glyphs, EmbeddedFontError> {
    let mut input = Reader::new(glyf);
    let mut push = Reader::new(push);
    let mut code = Reader::new(code);
    let mut output = Vec::with_capacity(glyf.len() * 2);
    let mut loca = Vec::with_capacity((usize::from(count) + 1) * 4);
    let mut max_instructions = 0_u16;
    for _ in 0..count {
        loca.extend((output.len() as u32).to_be_bytes());
        let instructions = decode_glyph(&mut input, &mut push, &mut code, &mut output)?;
        max_instructions = max_instructions.max(instructions);
        output.resize(output.len().next_multiple_of(4), 0);
        if output.len() > MAX_EMBEDDED_FONT_BYTES {
            return Err(EmbeddedFontError::TooLarge);
        }
    }
    loca.extend((output.len() as u32).to_be_bytes());
    if [input, push, code]
        .iter()
        .any(|block| block.position != block.data.len())
    {
        return Err(EmbeddedFontError::Malformed("CTF glyph data left over"));
    }
    Ok(Glyphs {
        glyf: output,
        loca,
        max_instructions,
    })
}

/// Appends one glyph in TrueType form; returns its program's length.
fn decode_glyph(
    input: &mut Reader<'_>,
    push: &mut Reader<'_>,
    code: &mut Reader<'_>,
    output: &mut Vec<u8>,
) -> Result<u16, EmbeddedFontError> {
    let contours = input.i16()?;
    if contours == 0 {
        return Ok(0);
    }
    if contours == -1 {
        output.extend(contours.to_be_bytes());
        output.extend(input.bytes(8)?);
        let mut flags;
        loop {
            flags = input.u16()?;
            output.extend(flags.to_be_bytes());
            let mut size = 2 + if flags & ARG_1_AND_2_ARE_WORDS != 0 {
                4
            } else {
                2
            };
            if flags & WE_HAVE_A_SCALE != 0 {
                size += 2;
            } else if flags & WE_HAVE_AN_X_AND_Y_SCALE != 0 {
                size += 4;
            } else if flags & WE_HAVE_A_TWO_BY_TWO != 0 {
                size += 8;
            }
            output.extend(input.bytes(size)?);
            if flags & MORE_COMPONENTS == 0 {
                break;
            }
        }
        if flags & WE_HAVE_INSTRUCTIONS == 0 {
            return Ok(0);
        }
        let program = program(input, push, code)?;
        output.extend((program.len() as u16).to_be_bytes());
        output.extend(&program);
        return Ok(program.len() as u16);
    }
    let (contours, stored_bounds) = if contours == 0x7FFF {
        let contours = input.i16()?;
        let bounds = input.bytes(8)?;
        (contours, Some(bounds))
    } else {
        (contours, None)
    };
    if contours < 0 {
        return Err(EmbeddedFontError::Malformed("CTF contour count"));
    }
    let mut ends = Vec::with_capacity(contours as usize);
    let mut end = 0_u32;
    for index in 0..contours {
        let value = u32::from(input.u16_255()?);
        end = if index == 0 { value } else { end + value };
        if end > u32::from(u16::MAX) {
            return Err(EmbeddedFontError::Malformed("CTF contour end"));
        }
        ends.push(end as u16);
    }
    let points = end as usize + 1;
    let flags = input.bytes(points)?;
    let mut coordinates = Vec::with_capacity(points);
    let (mut x, mut y) = (0_i32, 0_i32);
    for flag in flags {
        let (dx, dy) = triplet(flag & 0x7F, input)?;
        x += dx;
        y += dy;
        coordinates.push((x, y, flag & 0x80 == 0));
    }
    let program = program(input, push, code)?;
    output.extend(contours.to_be_bytes());
    match stored_bounds {
        Some(bounds) => output.extend(bounds),
        None => {
            let (x, y, _) = coordinates[0];
            let (mut x_min, mut y_min, mut x_max, mut y_max) = (x, y, x, y);
            for &(x, y, _) in &coordinates {
                (x_min, x_max) = (x_min.min(x), x_max.max(x));
                (y_min, y_max) = (y_min.min(y), y_max.max(y));
            }
            for value in [x_min, y_min, x_max, y_max] {
                output.extend((value as i16).to_be_bytes());
            }
        }
    }
    for end in ends {
        output.extend(end.to_be_bytes());
    }
    output.extend((program.len() as u16).to_be_bytes());
    output.extend(&program);
    write_points(&coordinates, output);
    Ok(program.len() as u16)
}

/// One `(dx, dy)` of the triplet encoding MTX shares with WOFF2.
fn triplet(index: u8, input: &mut Reader<'_>) -> Result<(i32, i32), EmbeddedFontError> {
    let signed = |positive: u8, value: i32| if positive & 1 == 1 { value } else { -value };
    let index_value = i32::from(index);
    Ok(match index {
        0..10 => {
            let value = i32::from(input.u8()?);
            (0, signed(index, ((index_value & 14) << 7) + value))
        }
        10..20 => {
            let value = i32::from(input.u8()?);
            (signed(index, (((index_value - 10) & 14) << 7) + value), 0)
        }
        20..84 => {
            let base = index_value - 20;
            let value = i32::from(input.u8()?);
            (
                signed(index, 1 + (base & 0x30) + (value >> 4)),
                signed(index >> 1, 1 + ((base & 0x0C) << 2) + (value & 0x0F)),
            )
        }
        84..120 => {
            let base = index_value - 84;
            let bytes = input.bytes(2)?;
            (
                signed(index, 1 + ((base / 12) << 8) + i32::from(bytes[0])),
                signed(
                    index >> 1,
                    1 + (((base % 12) >> 2) << 8) + i32::from(bytes[1]),
                ),
            )
        }
        120..124 => {
            let bytes = input.bytes(3)?;
            let (a, b, c) = (
                i32::from(bytes[0]),
                i32::from(bytes[1]),
                i32::from(bytes[2]),
            );
            (
                signed(index, (a << 4) + (b >> 4)),
                signed(index >> 1, ((b & 0x0F) << 8) + c),
            )
        }
        _ => {
            let bytes = input.bytes(4)?;
            (
                signed(index, i32::from(u16::from_be_bytes([bytes[0], bytes[1]]))),
                signed(
                    index >> 1,
                    i32::from(u16::from_be_bytes([bytes[2], bytes[3]])),
                ),
            )
        }
    })
}

/// The glyph program: its pushed values as push instructions, then the rest
/// of its instructions from the code block.
fn program(
    input: &mut Reader<'_>,
    push: &mut Reader<'_>,
    code: &mut Reader<'_>,
) -> Result<Vec<u8>, EmbeddedFontError> {
    let push_count = usize::from(input.u16_255()?);
    let code_size = usize::from(input.u16_255()?);
    let mut values: Vec<i16> = Vec::with_capacity(push_count);
    while values.len() < push_count {
        let hop = match push.data.get(push.position) {
            Some(&HOP3) => 3,
            Some(&HOP4) => 5,
            _ => 0,
        };
        if hop == 0 {
            values.push(push.i16_255()?);
            continue;
        }
        if values.len() < 2 || values.len() + hop > push_count {
            return Err(EmbeddedFontError::Malformed("CTF push hop"));
        }
        push.u8()?;
        let repeated = values[values.len() - 2];
        values.push(repeated);
        for _ in 0..hop / 2 {
            values.push(push.i16_255()?);
            values.push(repeated);
        }
    }
    let mut program = Vec::with_capacity(push_count * 2 + code_size + 4);
    let mut index = 0;
    while index < values.len() {
        let bytes = (0..=255).contains(&values[index]);
        let run = values[index..]
            .iter()
            .take(255)
            .take_while(|value| (0..=255).contains(*value) == bytes)
            .count();
        let (short, long) = if bytes { (0xB0, 0x40) } else { (0xB8, 0x41) };
        if run <= 8 {
            program.push(short + run as u8 - 1);
        } else {
            program.extend([long, run as u8]);
        }
        for value in &values[index..index + run] {
            if bytes {
                program.push(*value as u8);
            } else {
                program.extend(value.to_be_bytes());
            }
        }
        index += run;
    }
    program.extend(code.bytes(code_size)?);
    if program.len() > usize::from(u16::MAX) {
        return Err(EmbeddedFontError::Malformed("CTF glyph program length"));
    }
    Ok(program)
}

/// Simple-glyph flags and coordinates, short forms where they fit.
fn write_points(points: &[(i32, i32, bool)], output: &mut Vec<u8>) {
    const ON_CURVE: u8 = 0x01;
    const X_SHORT: u8 = 0x02;
    const Y_SHORT: u8 = 0x04;
    const X_SAME_OR_POSITIVE: u8 = 0x10;
    const Y_SAME_OR_POSITIVE: u8 = 0x20;
    let mut flags = Vec::with_capacity(points.len());
    let mut xs = Vec::with_capacity(points.len() * 2);
    let mut ys = Vec::with_capacity(points.len() * 2);
    let mut previous = (0, 0);
    for &(x, y, on_curve) in points {
        let mut flag = if on_curve { ON_CURVE } else { 0 };
        for (delta, short, same, bytes) in [
            (x - previous.0, X_SHORT, X_SAME_OR_POSITIVE, &mut xs),
            (y - previous.1, Y_SHORT, Y_SAME_OR_POSITIVE, &mut ys),
        ] {
            if delta == 0 {
                flag |= same;
            } else if delta.abs() < 256 {
                flag |= short | if delta > 0 { same } else { 0 };
                bytes.push(delta.unsigned_abs() as u8);
            } else {
                bytes.extend((delta as i16).to_be_bytes());
            }
        }
        flags.push(flag);
        previous = (x, y);
    }
    output.extend(flags);
    output.extend(xs);
    output.extend(ys);
}

/// `cvt `: an entry count, then each value as a delta from the previous one.
fn decode_cvt(data: &[u8]) -> Result<Vec<u8>, EmbeddedFontError> {
    const LOWEST: i32 = 238;
    let mut input = Reader::new(data);
    let count = input.u16()?;
    let mut output = Vec::with_capacity(usize::from(count) * 2);
    let mut value = 0_i16;
    for _ in 0..count {
        let delta = match input.u8()? {
            code @ 0..238 => i32::from(code),
            238 => i32::from(input.i16()?),
            code @ 239..248 => -(i32::from(input.u8()?) + i32::from(code - 239) * LOWEST),
            code => i32::from(input.u8()?) + i32::from(code - 247) * LOWEST,
        };
        value = value.wrapping_add(delta as i16);
        output.extend(value.to_be_bytes());
    }
    Ok(output)
}

fn assemble(
    version: &[u8],
    mut tables: Vec<([u8; 4], Vec<u8>)>,
) -> Result<Vec<u8>, EmbeddedFontError> {
    tables.sort_by_key(|(tag, _)| *tag);
    let count = tables.len() as u16;
    let mut power = 1_u16;
    let mut selector = 0_u16;
    while power * 2 <= count {
        power *= 2;
        selector += 1;
    }
    let range = power * 16;
    let header = 12 + 16 * tables.len();
    let total = header
        + tables
            .iter()
            .map(|(_, data)| data.len().next_multiple_of(4))
            .sum::<usize>();
    if total > MAX_EMBEDDED_FONT_BYTES {
        return Err(EmbeddedFontError::TooLarge);
    }
    let mut font = Vec::with_capacity(total);
    font.extend(version);
    for value in [count, range, selector, count * 16 - range] {
        font.extend(value.to_be_bytes());
    }
    let mut offset = header;
    let mut head_offset = None;
    for (tag, data) in &tables {
        if tag == b"head" && data.len() >= 12 {
            head_offset = Some(offset);
        }
        let mut sum_data = data.clone();
        if tag == b"head" && sum_data.len() >= 12 {
            sum_data[8..12].fill(0);
        }
        font.extend(tag);
        font.extend(checksum(&sum_data).to_be_bytes());
        font.extend((offset as u32).to_be_bytes());
        font.extend((data.len() as u32).to_be_bytes());
        offset += data.len().next_multiple_of(4);
    }
    for (tag, data) in &tables {
        let start = font.len();
        font.extend(data);
        if tag == b"head" && data.len() >= 12 {
            font[start + 8..start + 12].fill(0);
        }
        font.resize(font.len().next_multiple_of(4), 0);
    }
    if let Some(offset) = head_offset {
        let adjustment = 0xB1B0_AFBA_u32.wrapping_sub(checksum(&font));
        font[offset + 8..offset + 12].copy_from_slice(&adjustment.to_be_bytes());
    }
    Ok(font)
}

fn checksum(data: &[u8]) -> u32 {
    data.chunks(4).fold(0_u32, |sum, chunk| {
        let mut word = [0_u8; 4];
        word[..chunk.len()].copy_from_slice(chunk);
        sum.wrapping_add(u32::from_be_bytes(word))
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn cvt_values_accumulate_their_deltas() {
        // 5, then +238*2+1 (two-byte positive), then -(238+3), then a word.
        let data = [0, 4, 5, 249, 1, 240, 3, 238, 0x80, 0x00];
        let values: Vec<i16> = decode_cvt(&data)
            .unwrap()
            .chunks(2)
            .map(|pair| i16::from_be_bytes([pair[0], pair[1]]))
            .collect();
        assert_eq!(values, [5, 482, 241, 241_i16.wrapping_add(i16::MIN)]);
    }

    #[test]
    fn pushed_values_expand_hops_and_come_back_as_push_instructions() {
        // pushCount 7, codeSize 1; values A=1, X1=300, Hop4 X2=2 X3=3.
        let glyph = [7, 1];
        let push = [1, 255, 50, HOP4, 2, 3];
        let code = [0x2F];
        let program = program(
            &mut Reader::new(&glyph),
            &mut Reader::new(&push),
            &mut Reader::new(&code),
        )
        .unwrap();
        assert_eq!(
            program,
            [0xB0, 1, 0xB8, 0x01, 0x2C, 0xB4, 1, 2, 1, 3, 1, 0x2F]
        );
    }

    #[test]
    fn a_hop_without_two_values_before_it_is_refused() {
        let program = program(
            &mut Reader::new(&[3, 0]),
            &mut Reader::new(&[HOP3, 5]),
            &mut Reader::new(&[]),
        );
        assert_eq!(program, Err(EmbeddedFontError::Malformed("CTF push hop")));
    }
}
