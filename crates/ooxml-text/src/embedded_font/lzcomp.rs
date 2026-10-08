//! LZCOMP, the LZ77 variant with adaptive Huffman codes that MicroType
//! Express compresses each block with (W3C MTX submission, appendix C).

use super::{EmbeddedFontError, charge};

/// Bytes the history starts with so early copies have something to reach.
const PRELOAD: usize = 2 * 32 * 96 + 4 * 256;
/// Copies from at least this far back are one byte longer than encoded.
const MAX_2BYTE_DISTANCE: usize = 512;
const ROOT: usize = 1;

/// The output length a block declares, before its run-length stage.
pub(super) fn declared_length(data: &[u8], version: u8) -> Result<usize, EmbeddedFontError> {
    let mut bits = Bits::new(data);
    if version != 1 {
        bits.bit()?;
    }
    Ok(bits.value(24)? as usize)
}

/// Unpacks one block of at most `limit` bytes, charging `budget` for each run
/// the run-length stage expands as it is produced.
pub(super) fn unpack(
    data: &[u8],
    version: u8,
    limit: usize,
    budget: &mut usize,
) -> Result<Vec<u8>, EmbeddedFontError> {
    let mut bits = Bits::new(data);
    let run_length = version != 1 && bits.bit()?;
    let mut distances = Tree::new(8);
    let mut lengths = Tree::new(8);
    let out_len = bits.value(24)? as usize;
    if out_len > limit {
        return Err(EmbeddedFontError::TooLarge);
    }
    let mut ranges = 1;
    while (1_usize << (3 * ranges)) < out_len {
        ranges += 1;
    }
    let dup2 = 256 + 8 * ranges;
    let mut symbols = Tree::new(dup2 + 3);

    let mut history = preload();
    history.reserve(out_len);
    let mut output = Output::new(run_length, limit);
    let end = PRELOAD + out_len;
    while history.len() < end {
        let symbol = symbols.read(&mut bits)?;
        if symbol < 256 {
            output.push(&mut history, symbol as u8, budget)?;
        } else if (dup2..dup2 + 3).contains(&symbol) {
            let value = history[history.len() - 2 * (symbol - dup2 + 1)];
            output.push(&mut history, value, budget)?;
        } else {
            let code = symbol - 256;
            let distance_symbols = code / 8 + 1;
            let mut length = read_length(code % 8, &mut lengths, &mut bits)?;
            let mut distance = 0_usize;
            for _ in 0..distance_symbols {
                distance = (distance << 3) | distances.read(&mut bits)?;
            }
            distance += 1;
            if distance >= MAX_2BYTE_DISTANCE {
                length += 1;
            }
            let position = history.len();
            if distance + length - 1 > position || position + length > end {
                return Err(EmbeddedFontError::Malformed("LZCOMP copy out of range"));
            }
            let start = position + 1 - distance - length;
            for offset in 0..length {
                let value = history[start + offset];
                output.push(&mut history, value, budget)?;
            }
        }
    }
    Ok(output.finish(history))
}

/// A copy's length: 2-bit groups, each with a stop bit above it, the first
/// from the copy symbol and the rest from the length tree.
fn read_length(
    first: usize,
    lengths: &mut Tree,
    bits: &mut Bits<'_>,
) -> Result<usize, EmbeddedFontError> {
    let mut group = first;
    let mut value = 0_usize;
    for _ in 0..12 {
        value = (value << 2) | (group & 3);
        if group & 4 == 0 {
            return Ok(value + 2);
        }
        group = lengths.read(bits)?;
    }
    Err(EmbeddedFontError::Malformed("LZCOMP length"))
}

fn preload() -> Vec<u8> {
    let mut history = Vec::with_capacity(PRELOAD);
    for high in 0..32_u8 {
        for low in 0..96_u8 {
            history.extend([high, low]);
        }
    }
    for value in 0..=255_u8 {
        history.extend([value; 4]);
    }
    history
}

/// Where decoded bytes go: straight out, or through the run-length stage.
struct Output {
    run_length: Option<RunLength>,
    limit: usize,
}

struct RunLength {
    bytes: Vec<u8>,
    escape: Option<u8>,
    state: RunState,
}

enum RunState {
    Normal,
    Escaped,
    Repeat(u8),
}

impl Output {
    fn new(run_length: bool, limit: usize) -> Self {
        Self {
            run_length: run_length.then(|| RunLength {
                bytes: Vec::new(),
                escape: None,
                state: RunState::Normal,
            }),
            limit,
        }
    }

    fn push(
        &mut self,
        history: &mut Vec<u8>,
        value: u8,
        budget: &mut usize,
    ) -> Result<(), EmbeddedFontError> {
        history.push(value);
        let Some(run) = &mut self.run_length else {
            return Ok(());
        };
        let Some(escape) = run.escape else {
            run.escape = Some(value);
            return Ok(());
        };
        match run.state {
            RunState::Normal if value == escape => run.state = RunState::Escaped,
            RunState::Normal => run.bytes.push(value),
            RunState::Escaped if value == 0 => {
                run.bytes.push(escape);
                run.state = RunState::Normal;
            }
            RunState::Escaped => run.state = RunState::Repeat(value),
            RunState::Repeat(count) => {
                // The block's declared length paid for the three coded bytes,
                // not for what they expand to.
                charge(budget, usize::from(count))?;
                run.bytes
                    .extend(std::iter::repeat_n(value, usize::from(count)));
                run.state = RunState::Normal;
            }
        }
        if run.bytes.len() > self.limit {
            return Err(EmbeddedFontError::TooLarge);
        }
        Ok(())
    }

    fn finish(self, mut history: Vec<u8>) -> Vec<u8> {
        match self.run_length {
            Some(run) => run.bytes,
            None => {
                history.drain(..PRELOAD);
                history
            }
        }
    }
}

/// Most significant bit first.
struct Bits<'a> {
    data: &'a [u8],
    position: usize,
    buffer: u8,
    left: u8,
}

impl<'a> Bits<'a> {
    fn new(data: &'a [u8]) -> Self {
        Self {
            data,
            position: 0,
            buffer: 0,
            left: 0,
        }
    }

    fn bit(&mut self) -> Result<bool, EmbeddedFontError> {
        if self.left == 0 {
            self.buffer = *self
                .data
                .get(self.position)
                .ok_or(EmbeddedFontError::Truncated)?;
            self.position += 1;
            self.left = 8;
        }
        self.left -= 1;
        Ok(self.buffer >> self.left & 1 == 1)
    }

    fn value(&mut self, count: u32) -> Result<u32, EmbeddedFontError> {
        let mut value = 0;
        for _ in 0..count {
            value = value << 1 | u32::from(self.bit()?);
        }
        Ok(value)
    }
}

#[derive(Clone, Copy, Default)]
struct Node {
    up: usize,
    left: usize,
    right: usize,
    /// The leaf's symbol; `None` for an internal node.
    symbol: Option<usize>,
    weight: u32,
}

/// Adaptive Huffman tree over symbols `0..range`, kept in sibling order.
struct Tree {
    nodes: Vec<Node>,
    leaf: Vec<usize>,
}

impl Tree {
    fn new(range: usize) -> Self {
        let mut nodes = vec![Node::default(); 2 * range];
        for (index, node) in nodes.iter_mut().enumerate().skip(1) {
            if index >= range {
                node.symbol = Some(index - range);
            } else {
                node.left = 2 * index;
                node.right = 2 * index + 1;
            }
            if index > ROOT {
                node.up = index / 2;
                node.weight = 1;
            }
        }
        for index in (ROOT..range).rev() {
            nodes[index].weight = nodes[2 * index].weight + nodes[2 * index + 1].weight;
        }
        let mut tree = Self {
            nodes,
            leaf: (range..2 * range).collect(),
        };
        if range > 256 && range < 512 {
            tree.update(tree.leaf[256]);
            tree.update(tree.leaf[257]);
            for _ in 0..12 {
                tree.update(tree.leaf[range - 3]);
            }
            for _ in 0..6 {
                tree.update(tree.leaf[range - 2]);
            }
        } else {
            for _ in 0..2 {
                for symbol in 0..range {
                    tree.update(tree.leaf[symbol]);
                }
            }
        }
        tree
    }

    fn read(&mut self, bits: &mut Bits<'_>) -> Result<usize, EmbeddedFontError> {
        let mut index = ROOT;
        loop {
            let node = self.nodes[index];
            if let Some(symbol) = node.symbol {
                self.update(index);
                return Ok(symbol);
            }
            index = if bits.bit()? { node.right } else { node.left };
        }
    }

    /// Counts one more use of the node at `index`, swapping it ahead of the
    /// first node of equal weight to keep the sibling property.
    fn update(&mut self, mut index: usize) {
        while index != ROOT {
            let weight = self.nodes[index].weight;
            let mut first = index - 1;
            if self.nodes[first].weight == weight {
                while self.nodes[first].weight == weight {
                    first -= 1;
                }
                first += 1;
                if first > ROOT {
                    self.swap(index, first);
                    index = first;
                }
            }
            self.nodes[index].weight = weight + 1;
            index = self.nodes[index].up;
        }
        self.nodes[ROOT].weight += 1;
    }

    fn swap(&mut self, a: usize, b: usize) {
        let (up_a, up_b) = (self.nodes[a].up, self.nodes[b].up);
        self.nodes.swap(a, b);
        self.nodes[a].up = up_a;
        self.nodes[b].up = up_b;
        for index in [a, b] {
            let node = self.nodes[index];
            match node.symbol {
                Some(symbol) => self.leaf[symbol] = index,
                None => {
                    self.nodes[node.left].up = index;
                    self.nodes[node.right].up = index;
                }
            }
        }
    }
}

/// The encoder half of LZCOMP, for building test streams.
#[cfg(test)]
pub(super) mod encoder {
    use super::{ROOT, Tree};

    pub(in super::super) struct Encoder {
        bits: Vec<bool>,
        symbols: Tree,
        distances: Tree,
        lengths: Tree,
        ranges: usize,
    }

    impl Encoder {
        /// A version-3 block declaring `declared` bytes, flagged for the
        /// run-length stage when `run_length` is set.
        pub(in super::super) fn new(declared: usize, run_length: bool) -> Self {
            let mut bits = vec![run_length];
            bits.extend((0..24).rev().map(|shift| declared >> shift & 1 == 1));
            let mut ranges = 1;
            while (1_usize << (3 * ranges)) < declared {
                ranges += 1;
            }
            Self {
                bits,
                symbols: Tree::new(256 + 8 * ranges + 3),
                distances: Tree::new(8),
                lengths: Tree::new(8),
                ranges,
            }
        }

        pub(in super::super) fn literals(mut self, bytes: &[u8]) -> Self {
            for byte in bytes {
                write(&mut self.symbols, usize::from(*byte), &mut self.bits);
            }
            self
        }

        /// `length` bytes from `distance` back.
        pub(in super::super) fn copy(mut self, distance: usize, length: usize) -> Self {
            let mut value = if distance >= 512 { length - 1 } else { length } - 2;
            let mut groups = Vec::new();
            loop {
                groups.push(value & 3);
                value >>= 2;
                if value == 0 {
                    break;
                }
            }
            groups.reverse();
            let last = groups.len() - 1;
            let groups: Vec<usize> = groups
                .iter()
                .enumerate()
                .map(|(index, group)| if index < last { group | 4 } else { *group })
                .collect();
            let mut value = distance - 1;
            let mut digits = Vec::new();
            loop {
                digits.push(value & 7);
                value >>= 3;
                if value == 0 {
                    break;
                }
            }
            digits.reverse();
            assert!(digits.len() <= self.ranges);
            let symbol = 256 + 8 * (digits.len() - 1) + groups[0];
            write(&mut self.symbols, symbol, &mut self.bits);
            for group in &groups[1..] {
                write(&mut self.lengths, *group, &mut self.bits);
            }
            for digit in digits {
                write(&mut self.distances, digit, &mut self.bits);
            }
            self
        }

        pub(in super::super) fn finish(self) -> Vec<u8> {
            self.bits
                .chunks(8)
                .map(|chunk| {
                    chunk.iter().enumerate().fold(0, |byte, (index, bit)| {
                        byte | (u8::from(*bit) << (7 - index))
                    })
                })
                .collect()
        }
    }

    /// A literal block of `bytes`, without the run-length stage.
    pub(in super::super) fn literal_block(bytes: &[u8]) -> Vec<u8> {
        Encoder::new(bytes.len(), false).literals(bytes).finish()
    }

    fn write(tree: &mut Tree, symbol: usize, out: &mut Vec<bool>) {
        let mut index = tree.leaf[symbol];
        let leaf = index;
        let mut path = Vec::new();
        while index != ROOT {
            let up = tree.nodes[index].up;
            path.push(tree.nodes[up].right == index);
            index = up;
        }
        out.extend(path.iter().rev());
        tree.update(leaf);
    }
}

#[cfg(test)]
mod tests {
    use super::encoder::{Encoder, literal_block};
    use super::*;

    #[test]
    fn literals_round_trip_through_the_adaptive_tree() {
        let text = b"MicroType Express keeps adapting its codes as symbols repeat";
        assert_eq!(
            unpack(&literal_block(text), 3, 1024, &mut { usize::MAX }).unwrap(),
            text
        );
        assert_eq!(
            unpack(&literal_block(text), 3, text.len() - 1, &mut { usize::MAX }),
            Err(EmbeddedFontError::TooLarge)
        );
        assert_eq!(declared_length(&literal_block(text), 3), Ok(text.len()));
    }

    #[test]
    fn copies_reach_back_into_the_output_and_the_preload() {
        // A copy names its length and how far back its last byte is; from
        // 512 back it is one byte longer than its symbols say.
        let block = Encoder::new(4 + 2 + 2 + 600 + 3, false)
            .copy(4, 4)
            .literals(b"ab")
            .copy(1, 2)
            .copy(1, 600)
            .copy(600, 3)
            .finish();
        let mut expected = preload();
        for (back, length) in [(4, 4), (0, 0), (1, 2), (1, 600), (600, 3)] {
            if length == 0 {
                expected.extend(b"ab");
                continue;
            }
            let start = expected.len() - back - length + 1;
            for offset in 0..length {
                expected.push(expected[start + offset]);
            }
        }
        assert_eq!(
            unpack(&block, 3, 4096, &mut { usize::MAX }).unwrap(),
            expected[PRELOAD..]
        );
        assert_eq!(&expected[PRELOAD + 4..PRELOAD + 8], b"abab");
    }

    #[test]
    fn the_run_length_stage_expands_runs_and_escaped_escapes() {
        // Escape 0xEE; "a", a literal 0xEE (escape, 0), five 'z' (escape, 5,
        // 'z'), "b".
        let coded = [0xEE, b'a', 0xEE, 0, 0xEE, 5, b'z', b'b'];
        let block = Encoder::new(coded.len(), true).literals(&coded).finish();
        let mut budget = 100;
        assert_eq!(
            unpack(&block, 3, 64, &mut budget).unwrap(),
            [b'a', 0xEE, b'z', b'z', b'z', b'z', b'z', b'b']
        );
        // The repeat's five bytes are charged as they are produced.
        assert_eq!(budget, 95);
        assert_eq!(
            unpack(&block, 3, 7, &mut { usize::MAX }),
            Err(EmbeddedFontError::TooLarge)
        );
        let mut short = 4;
        assert_eq!(
            unpack(&block, 3, 64, &mut short),
            Err(EmbeddedFontError::TooLarge)
        );
    }

    #[test]
    fn a_stream_ending_early_is_truncated() {
        let stream = literal_block(b"truncated stream");
        assert_eq!(
            unpack(&stream[..stream.len() - 3], 3, 1024, &mut { usize::MAX }),
            Err(EmbeddedFontError::Truncated)
        );
    }

    #[test]
    fn the_preload_holds_pairs_then_quads() {
        let history = preload();
        assert_eq!(history.len(), PRELOAD);
        assert_eq!(&history[..4], &[0, 0, 0, 1]);
        assert_eq!(&history[PRELOAD - 4..], &[255; 4]);
    }
}
