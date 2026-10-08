//! Built and mutated parts: the run-length stage on a whole font, refused
//! directories and coordinates, and a structured fuzz over the real Lato
//! blocks (seeded; `EMBEDDED_FONT_FUZZ_CASES` raises the case count). Run
//! with overflow checks, as `cargo test` does.

use super::lzcomp::encoder::{Encoder, literal_block};
use super::*;

const LATO: &[u8] = include_bytes!("../../tests/fonts/Lato-regular.fntdata");
const PART: &str = "ppt/fonts/font1.fntdata";

fn eot(data: &[u8], flags: u32) -> Vec<u8> {
    let mut part = vec![0_u8; EOT_MIN_HEADER];
    part[8..12].copy_from_slice(&0x0002_0001_u32.to_le_bytes());
    part[12..16].copy_from_slice(&flags.to_le_bytes());
    part[34..36].copy_from_slice(&EOT_MAGIC.to_le_bytes());
    part[4..8].copy_from_slice(&(data.len() as u32).to_le_bytes());
    part.extend(data);
    let size = part.len() as u32;
    part[..4].copy_from_slice(&size.to_le_bytes());
    part
}

fn mtx(rest: &[u8], push: &[u8], code: &[u8]) -> Vec<u8> {
    let second = 10 + rest.len();
    let third = second + push.len();
    let mut data = vec![3_u8];
    data.extend([0, 0, 0]);
    data.extend(&(second as u32).to_be_bytes()[1..]);
    data.extend(&(third as u32).to_be_bytes()[1..]);
    data.extend(rest);
    data.extend(push);
    data.extend(code);
    data
}

fn literal_part(rest: &[u8], push: &[u8], code: &[u8]) -> Vec<u8> {
    eot(
        &mtx(
            &literal_block(rest),
            &literal_block(push),
            &literal_block(code),
        ),
        EOT_COMPRESSED,
    )
}

/// The three decoded blocks of the real Lato part.
fn lato_blocks() -> [Vec<u8>; 3] {
    let data = &LATO[LATO.len() - u32::from_le_bytes(LATO[4..8].try_into().unwrap()) as usize..];
    let u24 = |at: usize| {
        usize::from(data[at]) << 16 | usize::from(data[at + 1]) << 8 | usize::from(data[at + 2])
    };
    let (second, third) = (u24(4), u24(7));
    [&data[10..second], &data[second..third], &data[third..]].map(|block| {
        lzcomp::unpack(block, data[0], MAX_EMBEDDED_FONT_BYTES, &mut { usize::MAX }).unwrap()
    })
}

/// A CTF directory over `entries` (tag, offset, length).
fn directory(entries: &[([u8; 4], u32, u32)]) -> Vec<u8> {
    let mut rest = vec![0, 1, 0, 0];
    rest.extend((entries.len() as u16).to_be_bytes());
    rest.extend([0; 6]);
    for (tag, offset, length) in entries {
        rest.extend(tag);
        rest.extend([0; 4]);
        rest.extend(offset.to_be_bytes());
        rest.extend(length.to_be_bytes());
    }
    rest
}

/// Run-length codes `bytes` with the byte it uses least as the escape.
fn run_length(bytes: &[u8]) -> Vec<u8> {
    let mut counts = [0_usize; 256];
    for byte in bytes {
        counts[usize::from(*byte)] += 1;
    }
    let escape = (0..=255_u8)
        .min_by_key(|byte| counts[usize::from(*byte)])
        .unwrap();
    let mut coded = vec![escape];
    let mut index = 0;
    while index < bytes.len() {
        let byte = bytes[index];
        let run = bytes[index..]
            .iter()
            .take(255)
            .take_while(|next| **next == byte)
            .count();
        if run >= 4 {
            coded.extend([escape, run as u8, byte]);
            index += run;
        } else {
            if byte == escape {
                coded.extend([escape, 0]);
            } else {
                coded.push(byte);
            }
            index += 1;
        }
    }
    coded
}

#[test]
fn a_font_whose_blocks_use_the_run_length_stage_decodes_the_same() {
    let expected = decode_embedded_font(LATO, PART).unwrap();
    let coded = lato_blocks().map(|block| run_length(&block));
    // The CTF block's zero runs come back through the repeat code.
    assert!(
        coded[0]
            .windows(2)
            .any(|pair| pair[0] == coded[0][0] && pair[1] > 0)
    );
    let blocks = coded.map(|coded| Encoder::new(coded.len(), true).literals(&coded).finish());
    let part = eot(&mtx(&blocks[0], &blocks[1], &blocks[2]), EOT_COMPRESSED);
    assert_eq!(decode_embedded_font(&part, PART).unwrap(), expected);
}

#[test]
fn directories_with_no_table_a_table_twice_or_too_much_are_refused() {
    let [rest, push, code] = lato_blocks();
    let none = directory(&[]);
    assert_eq!(
        decode_embedded_font(&literal_part(&none, &push, &code), PART),
        Err(EmbeddedFontError::Malformed("CTF table count"))
    );
    let mut twice = rest.clone();
    let first_tag = twice[12..16].to_vec();
    twice[28..32].copy_from_slice(&first_tag);
    assert_eq!(
        decode_embedded_font(&literal_part(&twice, &push, &code), PART),
        Err(EmbeddedFontError::Malformed("CTF table listed twice"))
    );
    // 256 tables each covering the whole block (48 MiB in all): refused
    // before any copy.
    let size = 192 * 1024;
    let entries: Vec<_> = (0..256)
        .map(|index| (format!("t{index:03}").as_bytes().try_into().unwrap(), 0, 0))
        .collect();
    let mut overlapping = directory(&entries);
    overlapping.resize(size, 0);
    for index in 0..256 {
        let length = 12 + 16 * index + 12;
        overlapping[length..length + 4].copy_from_slice(&(size as u32).to_be_bytes());
    }
    assert_eq!(
        decode_embedded_font(&literal_part(&overlapping, &[], &[]), PART).map(|font| font.len()),
        Err(EmbeddedFontError::TooLarge)
    );
}

#[test]
fn coordinates_beyond_sixteen_bits_are_refused() {
    // One glyph of 65,536 points, each 65,535 units right of the last.
    let mut glyf = 1_i16.to_be_bytes().to_vec();
    glyf.push(253);
    glyf.extend(65_535_u16.to_be_bytes());
    glyf.extend(std::iter::repeat_n(125_u8, 65_536));
    for _ in 0..65_536 {
        glyf.extend([0xFF, 0xFF, 0, 0]);
    }
    glyf.extend([0, 0]);
    let mut head = vec![0_u8; 54];
    head[18..20].copy_from_slice(&1000_u16.to_be_bytes());
    let mut maxp = vec![0_u8; 32];
    maxp[..4].copy_from_slice(&0x0001_0000_u32.to_be_bytes());
    maxp[4..6].copy_from_slice(&1_u16.to_be_bytes());
    let base = 12 + 16 * 3;
    let glyf_length = glyf.len() as u32;
    let mut rest = directory(&[
        (*b"glyf", base, glyf_length),
        (*b"head", base + glyf_length, 54),
        (*b"maxp", base + glyf_length + 54, 32),
    ]);
    rest.extend(&glyf);
    rest.extend(&head);
    rest.extend(&maxp);
    assert_eq!(
        decode_embedded_font(&literal_part(&rest, &[], &[]), PART),
        Err(EmbeddedFontError::Malformed("CTF coordinate out of range"))
    );
}

struct Rng(u64);

impl Rng {
    fn next(&mut self) -> u64 {
        self.0 ^= self.0 << 13;
        self.0 ^= self.0 >> 7;
        self.0 ^= self.0 << 17;
        self.0
    }

    fn below(&mut self, bound: usize) -> usize {
        (self.next() % bound as u64) as usize
    }
}

const WORDS: [u32; 12] = [
    0,
    1,
    2,
    3,
    0x7F,
    0xFF,
    0xFFFF,
    0x1_0000,
    0x7FFF_FFFF,
    0x8000_0000,
    0xFFFF_FFFE,
    0xFFFF_FFFF,
];
const BYTES: [u8; 14] = [
    0, 1, 0x7F, 0x80, 0xFF, 249, 250, 251, 252, 253, 254, 255, 0x7E, 2,
];

/// One structured mutation of the Lato part: directory fields and count,
/// `numGlyphs`, glyph bytes set to 255USHORT/255SHORT/hop codes, push and code
/// blocks cut or extended, EOT and MTX headers, bit flips, and LZCOMP declared
/// lengths that differ from the content, with and without the run-length
/// flag.
fn mutate(rng: &mut Rng, [rest, push, code]: &[Vec<u8>; 3]) -> Vec<u8> {
    let count = usize::from(u16::from_be_bytes([rest[4], rest[5]]));
    let record = |index: usize| 12 + 16 * index;
    let find = |tag: &[u8]| {
        (0..count)
            .find(|index| &rest[record(*index)..record(*index) + 4] == tag)
            .unwrap()
    };
    let word = |at: usize| u32::from_be_bytes(rest[at..at + 4].try_into().unwrap()) as usize;
    let (glyf, maxp) = (find(b"glyf"), find(b"maxp"));
    let (glyf_offset, glyf_length) = (word(record(glyf) + 8), word(record(glyf) + 12));
    let maxp_offset = word(record(maxp) + 8);
    let data_start = LATO.len() - u32::from_le_bytes(LATO[4..8].try_into().unwrap()) as usize;
    let (mut rest, mut push, mut code) = (rest.clone(), push.clone(), code.clone());
    match rng.below(9) {
        0 => {
            let field = record(rng.below(count)) + if rng.below(2) == 0 { 8 } else { 12 };
            let value = if rng.below(2) == 0 {
                WORDS[rng.below(WORDS.len())]
            } else {
                (rest.len() as u32)
                    .wrapping_add(rng.below(5) as u32)
                    .wrapping_sub(2)
            };
            rest[field..field + 4].copy_from_slice(&value.to_be_bytes());
        }
        1 => {
            let value = [0_u16, 1, 2, 15, 16, 17, 255, 256, 257, 0xFFFF][rng.below(10)];
            rest[4..6].copy_from_slice(&value.to_be_bytes());
        }
        2 => {
            let value = [0_u16, 1, 271, 273, 1000, 0x7FFF, 0xFFFF][rng.below(7)];
            rest[maxp_offset + 4..maxp_offset + 6].copy_from_slice(&value.to_be_bytes());
        }
        3 | 4 => {
            for _ in 0..1 + rng.below(4) {
                rest[glyf_offset + rng.below(glyf_length)] = BYTES[rng.below(BYTES.len())];
            }
        }
        5 => match rng.below(4) {
            0 => push.truncate(rng.below(push.len() + 1)),
            1 => {
                for _ in 0..1 + rng.below(4) {
                    let at = rng.below(push.len());
                    push[at] = BYTES[rng.below(BYTES.len())];
                }
            }
            2 => code.truncate(rng.below(code.len() + 1)),
            _ => {
                let value = BYTES[rng.below(BYTES.len())];
                push.extend(std::iter::repeat_n(value, 1 + rng.below(64)));
            }
        },
        6 => {
            let mut part = LATO.to_vec();
            let field = [0_usize, 4, 8, 12][rng.below(4)];
            let value = if field == 12 {
                [0_u32, 4, 0x1000_0000, 0x1000_0004, 1, 5, 0xFFFF_FFFF][rng.below(7)]
            } else {
                WORDS[rng.below(WORDS.len())]
            };
            part[field..field + 4].copy_from_slice(&value.to_le_bytes());
            return part;
        }
        7 => {
            let mut part = LATO.to_vec();
            match rng.below(4) {
                0 => part[data_start] = BYTES[rng.below(BYTES.len())],
                which @ (1 | 2) => {
                    let at = data_start + if which == 1 { 4 } else { 7 };
                    let value = WORDS[rng.below(WORDS.len())].min(0xFF_FFFF);
                    part[at..at + 3].copy_from_slice(&value.to_be_bytes()[1..]);
                }
                _ => {
                    for _ in 0..2 + rng.below(6) {
                        let at = data_start + 10 + rng.below(part.len() - data_start - 10);
                        part[at] ^= 1_u8.rotate_left(rng.below(8) as u32);
                    }
                }
            }
            return part;
        }
        _ => {
            let delta = [-5_i64, -1, 1, 7, 1 << 20][rng.below(5)];
            let declared = (rest.len() as i64 + delta).clamp(0, (1 << 24) - 1) as usize;
            let block = Encoder::new(declared, rng.below(2) == 1)
                .literals(&rest[..declared.min(rest.len())])
                .finish();
            return eot(
                &mtx(&block, &literal_block(&push), &literal_block(&code)),
                EOT_COMPRESSED,
            );
        }
    }
    literal_part(&rest, &push, &code)
}

#[test]
fn structured_mutations_of_a_real_part_fail_or_decode_within_limits() {
    let cases = std::env::var("EMBEDDED_FONT_FUZZ_CASES")
        .ok()
        .and_then(|cases| cases.parse().ok())
        .unwrap_or(150);
    let blocks = lato_blocks();
    let mut rng = Rng(0x9E37_79B9_7F4A_7C15);
    let mut decoded = 0;
    for _ in 0..cases {
        let part = mutate(&mut rng, &blocks);
        if let Ok(font) = decode_embedded_font(&part, PART) {
            assert!(is_sfnt(&font) && font.len() <= MAX_EMBEDDED_FONT_BYTES);
            decoded += 1;
        }
    }
    assert!(decoded < cases);
}
