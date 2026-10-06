//! A bounded RFC 1951 decoder used for ZIP method 8 entries.

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum DeflateErrorKind {
    UnexpectedEnd,
    InvalidBlockType,
    InvalidStoredLength,
    InvalidHuffmanTree,
    InvalidCode,
    InvalidDistance,
    OutputLimit,
    OperationLimit,
    SizeMismatch,
    TrailingData,
    AllocationFailed,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct DeflateError {
    pub kind: DeflateErrorKind,
    pub byte_offset: usize,
}

impl core::fmt::Display for DeflateError {
    fn fmt(&self, formatter: &mut core::fmt::Formatter<'_>) -> core::fmt::Result {
        write!(
            formatter,
            "invalid deflate stream ({:?}) at byte {}",
            self.kind, self.byte_offset
        )
    }
}

impl std::error::Error for DeflateError {}

struct BitReader<'a> {
    bytes: &'a [u8],
    bit_offset: usize,
    operations_left: usize,
}

impl<'a> BitReader<'a> {
    fn new(bytes: &'a [u8], max_operations: usize) -> Self {
        Self {
            bytes,
            bit_offset: 0,
            operations_left: max_operations,
        }
    }

    fn error(&self, kind: DeflateErrorKind) -> DeflateError {
        DeflateError {
            kind,
            byte_offset: self.bit_offset / 8,
        }
    }

    fn read_bits(&mut self, count: u8) -> Result<u32, DeflateError> {
        debug_assert!(count <= 16);
        self.consume_operations(1)?;
        let end = self
            .bit_offset
            .checked_add(count as usize)
            .ok_or_else(|| self.error(DeflateErrorKind::UnexpectedEnd))?;
        if end > self.bytes.len().saturating_mul(8) {
            return Err(self.error(DeflateErrorKind::UnexpectedEnd));
        }

        let mut value = 0_u32;
        for shift in 0..count {
            let byte = self.bytes[self.bit_offset / 8];
            let bit = (byte >> (self.bit_offset % 8)) & 1;
            value |= u32::from(bit) << shift;
            self.bit_offset += 1;
        }
        Ok(value)
    }

    fn consume_operations(&mut self, count: usize) -> Result<(), DeflateError> {
        self.operations_left = self
            .operations_left
            .checked_sub(count)
            .ok_or_else(|| self.error(DeflateErrorKind::OperationLimit))?;
        Ok(())
    }

    fn align_to_byte(&mut self) {
        self.bit_offset = (self.bit_offset + 7) & !7;
    }

    fn remaining_bits(&self) -> usize {
        self.bytes
            .len()
            .saturating_mul(8)
            .saturating_sub(self.bit_offset)
    }
}

struct Huffman {
    codes_by_length: [Vec<(u16, u16)>; 16],
    max_length: usize,
}

impl Huffman {
    fn from_lengths(
        lengths: &[u8],
        allow_empty: bool,
        require_complete: bool,
    ) -> Result<Self, DeflateErrorKind> {
        let mut counts = [0_u16; 16];
        for &length in lengths {
            if length > 15 {
                return Err(DeflateErrorKind::InvalidHuffmanTree);
            }
            if length != 0 {
                counts[length as usize] = counts[length as usize]
                    .checked_add(1)
                    .ok_or(DeflateErrorKind::InvalidHuffmanTree)?;
            }
        }

        let symbol_count: usize = counts.iter().map(|&count| usize::from(count)).sum();
        if symbol_count == 0 {
            if allow_empty {
                return Ok(Self {
                    codes_by_length: std::array::from_fn(|_| Vec::new()),
                    max_length: 0,
                });
            }
            return Err(DeflateErrorKind::InvalidHuffmanTree);
        }

        // RFC 1951 canonical-code oversubscription check. Incomplete trees are legal,
        // including the common single-distance-code form.
        let mut slots = 1_i32;
        for &count in counts.iter().skip(1) {
            slots = (slots << 1) - i32::from(count);
            if slots < 0 {
                return Err(DeflateErrorKind::InvalidHuffmanTree);
            }
        }
        if slots > 0 && (require_complete || symbol_count != 1 || counts[1] != 1) {
            return Err(DeflateErrorKind::InvalidHuffmanTree);
        }

        let mut next_code = [0_u16; 16];
        let mut code = 0_u16;
        for bits in 1..=15 {
            code = code
                .checked_add(counts[bits - 1])
                .and_then(|value| value.checked_shl(1))
                .ok_or(DeflateErrorKind::InvalidHuffmanTree)?;
            next_code[bits] = code;
        }

        let mut codes_by_length: [Vec<(u16, u16)>; 16] = std::array::from_fn(|_| Vec::new());
        let mut max_length = 0;
        for (symbol, &length) in lengths.iter().enumerate() {
            if length == 0 {
                continue;
            }
            let length = usize::from(length);
            let canonical = next_code[length];
            next_code[length] = canonical
                .checked_add(1)
                .ok_or(DeflateErrorKind::InvalidHuffmanTree)?;
            codes_by_length[length].push((canonical, symbol as u16));
            max_length = max_length.max(length);
        }

        Ok(Self {
            codes_by_length,
            max_length,
        })
    }

    fn decode(&self, reader: &mut BitReader<'_>) -> Result<u16, DeflateError> {
        if self.max_length == 0 {
            return Err(reader.error(DeflateErrorKind::InvalidCode));
        }

        let mut code = 0_u16;
        for length in 1..=self.max_length {
            code = (code << 1) | reader.read_bits(1)? as u16;
            let codes = &self.codes_by_length[length];
            // Canonical codes of a given length are consecutive and stored in order.
            if let Some(&(first, _)) = codes.first()
                && let Some(index) = code.checked_sub(first)
                && let Some(&(_, symbol)) = codes.get(usize::from(index))
            {
                return Ok(symbol);
            }
        }
        Err(reader.error(DeflateErrorKind::InvalidCode))
    }
}

const LENGTH_BASE: [usize; 29] = [
    3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131,
    163, 195, 227, 258,
];
const LENGTH_EXTRA: [u8; 29] = [
    0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0,
];
const DISTANCE_BASE: [usize; 30] = [
    1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537,
    2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577,
];
const DISTANCE_EXTRA: [u8; 30] = [
    0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13,
    13,
];
const CODE_LENGTH_ORDER: [usize; 19] = [
    16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15,
];

fn fixed_trees() -> Result<(Huffman, Huffman), DeflateErrorKind> {
    let mut literal_lengths = [0_u8; 288];
    literal_lengths[..=143].fill(8);
    literal_lengths[144..=255].fill(9);
    literal_lengths[256..=279].fill(7);
    literal_lengths[280..].fill(8);
    let distance_lengths = [5_u8; 32];
    Ok((
        Huffman::from_lengths(&literal_lengths, false, false)?,
        Huffman::from_lengths(&distance_lengths, false, false)?,
    ))
}

fn dynamic_trees(reader: &mut BitReader<'_>) -> Result<(Huffman, Huffman), DeflateError> {
    let literal_count = reader.read_bits(5)? as usize + 257;
    let distance_count = reader.read_bits(5)? as usize + 1;
    let code_length_count = reader.read_bits(4)? as usize + 4;
    if literal_count > 286 || distance_count > 32 {
        return Err(reader.error(DeflateErrorKind::InvalidHuffmanTree));
    }

    let mut code_lengths = [0_u8; 19];
    for &index in &CODE_LENGTH_ORDER[..code_length_count] {
        code_lengths[index] = reader.read_bits(3)? as u8;
    }
    let code_tree =
        Huffman::from_lengths(&code_lengths, false, true).map_err(|kind| reader.error(kind))?;

    let total = literal_count + distance_count;
    let mut lengths = Vec::new();
    lengths
        .try_reserve_exact(total)
        .map_err(|_| reader.error(DeflateErrorKind::AllocationFailed))?;
    while lengths.len() < total {
        match code_tree.decode(reader)? {
            value @ 0..=15 => lengths.push(value as u8),
            16 => {
                let previous = *lengths
                    .last()
                    .ok_or_else(|| reader.error(DeflateErrorKind::InvalidHuffmanTree))?;
                let repeat = reader.read_bits(2)? as usize + 3;
                if lengths.len().saturating_add(repeat) > total {
                    return Err(reader.error(DeflateErrorKind::InvalidHuffmanTree));
                }
                lengths.extend(std::iter::repeat_n(previous, repeat));
            }
            17 => {
                let repeat = reader.read_bits(3)? as usize + 3;
                if lengths.len().saturating_add(repeat) > total {
                    return Err(reader.error(DeflateErrorKind::InvalidHuffmanTree));
                }
                lengths.extend(std::iter::repeat_n(0, repeat));
            }
            18 => {
                let repeat = reader.read_bits(7)? as usize + 11;
                if lengths.len().saturating_add(repeat) > total {
                    return Err(reader.error(DeflateErrorKind::InvalidHuffmanTree));
                }
                lengths.extend(std::iter::repeat_n(0, repeat));
            }
            _ => return Err(reader.error(DeflateErrorKind::InvalidHuffmanTree)),
        }
    }

    let (literal_lengths, distance_lengths) = lengths.split_at(literal_count);
    if literal_lengths.get(256).copied().unwrap_or_default() == 0 {
        return Err(reader.error(DeflateErrorKind::InvalidHuffmanTree));
    }
    let literal_tree =
        Huffman::from_lengths(literal_lengths, false, false).map_err(|kind| reader.error(kind))?;
    let distance_tree =
        Huffman::from_lengths(distance_lengths, true, false).map_err(|kind| reader.error(kind))?;
    Ok((literal_tree, distance_tree))
}

fn reserve_output(
    output: &mut Vec<u8>,
    additional: usize,
    max_output: usize,
    reader: &BitReader<'_>,
) -> Result<(), DeflateError> {
    let required = output
        .len()
        .checked_add(additional)
        .ok_or_else(|| reader.error(DeflateErrorKind::OutputLimit))?;
    if required > max_output {
        return Err(reader.error(DeflateErrorKind::OutputLimit));
    }
    output
        .try_reserve(additional)
        .map_err(|_| reader.error(DeflateErrorKind::AllocationFailed))
}

fn decode_compressed_block(
    reader: &mut BitReader<'_>,
    output: &mut Vec<u8>,
    literal_tree: &Huffman,
    distance_tree: &Huffman,
    max_output: usize,
    allow_invalid_distance: bool,
) -> Result<(), DeflateError> {
    loop {
        let symbol = literal_tree.decode(reader)?;
        match symbol {
            0..=255 => {
                reserve_output(output, 1, max_output, reader)?;
                output.push(symbol as u8);
            }
            256 => return Ok(()),
            257..=285 => {
                let length_index = usize::from(symbol - 257);
                let length = LENGTH_BASE[length_index]
                    + reader.read_bits(LENGTH_EXTRA[length_index])? as usize;
                let distance_symbol = distance_tree.decode(reader)?;
                if distance_symbol > 29 && !allow_invalid_distance {
                    return Err(reader.error(DeflateErrorKind::InvalidDistance));
                }
                let distance = if distance_symbol > 29 {
                    0
                } else {
                    let distance_index = usize::from(distance_symbol);
                    DISTANCE_BASE[distance_index]
                        + reader.read_bits(DISTANCE_EXTRA[distance_index])? as usize
                };
                if (distance == 0 || distance > output.len()) && !allow_invalid_distance {
                    return Err(reader.error(DeflateErrorKind::InvalidDistance));
                }

                reserve_output(output, length, max_output, reader)?;
                reader.consume_operations(length)?;
                for _ in 0..length {
                    let byte = output
                        .len()
                        .checked_sub(distance)
                        .and_then(|index| output.get(index))
                        .copied()
                        .unwrap_or(0);
                    output.push(byte);
                }
            }
            _ => return Err(reader.error(DeflateErrorKind::InvalidCode)),
        }
    }
}

/// Decodes an RFC 1951 raw Deflate stream while enforcing an output ceiling.
pub fn decompress(
    input: &[u8],
    max_output: usize,
    expected_size: Option<usize>,
) -> Result<Vec<u8>, DeflateError> {
    let operation_limit = max_output
        .saturating_mul(2)
        .saturating_add(input.len().saturating_mul(256))
        .saturating_add(1_024);
    decompress_with_limit(input, max_output, expected_size, operation_limit)
        .map(|(output, _)| output)
}

/// Decodes the complete prefix of a truncated Deflate stream.
///
/// PDF readers commonly recover content from producer-generated streams that
/// omit the final block marker. ZIP callers must keep using [`decompress`],
/// which remains strict about stream termination and declared sizes.
#[cfg(any(
    feature = "native-formats",
    feature = "iwork-formats",
    feature = "pdf-formats",
    feature = "legacy-office-formats"
))]
pub(crate) fn decompress_prefix(input: &[u8], max_output: usize) -> Result<Vec<u8>, DeflateError> {
    let operation_limit = max_output
        .saturating_mul(2)
        .saturating_add(input.len().saturating_mul(256))
        .saturating_add(1_024);
    decompress_with_policy(input, max_output, None, operation_limit, true, false, false)
        .map(|(output, _)| output)
}

#[cfg(feature = "xps-formats")]
pub(crate) fn decompress_output_prefix(
    input: &[u8],
    max_output: usize,
) -> Result<Vec<u8>, DeflateError> {
    let operation_limit = max_output
        .saturating_mul(2)
        .saturating_add(input.len().saturating_mul(256))
        .saturating_add(1_024);
    decompress_with_policy(input, max_output, None, operation_limit, false, false, true)
        .map(|(output, _)| output)
}

/// Returns the valid prefix preceding structural corruption in a PDF content stream.
#[cfg(any(feature = "iwork-formats", feature = "pdf-formats"))]
pub(crate) fn decompress_damaged_prefix(
    input: &[u8],
    max_output: usize,
) -> Result<Vec<u8>, DeflateError> {
    let operation_limit = max_output
        .saturating_mul(2)
        .saturating_add(input.len().saturating_mul(256))
        .saturating_add(1_024);
    decompress_with_policy(input, max_output, None, operation_limit, true, true, false)
        .map(|(output, _)| output)
}

/// Variant used by the ZIP layer to apply a document-wide operation budget.
pub(crate) fn decompress_with_limit(
    input: &[u8],
    max_output: usize,
    expected_size: Option<usize>,
    max_operations: usize,
) -> Result<(Vec<u8>, usize), DeflateError> {
    decompress_with_policy(
        input,
        max_output,
        expected_size,
        max_operations,
        false,
        false,
        false,
    )
}

fn decompress_with_policy(
    input: &[u8],
    max_output: usize,
    expected_size: Option<usize>,
    max_operations: usize,
    allow_truncated_prefix: bool,
    allow_damaged_prefix: bool,
    allow_output_prefix: bool,
) -> Result<(Vec<u8>, usize), DeflateError> {
    let effective_limit = expected_size.map_or(max_output, |size| size.min(max_output));
    let mut reader = BitReader::new(input, max_operations);
    let mut output = Vec::new();
    output
        .try_reserve(input.len().saturating_mul(2).min(effective_limit))
        .map_err(|_| reader.error(DeflateErrorKind::AllocationFailed))?;

    let decode_result = (|| -> Result<(), DeflateError> {
        let mut final_block = false;
        while !final_block {
            // Charge a fixed setup cost so a stream made of millions of empty blocks
            // cannot turn tiny output into unbounded tree-building work.
            reader.consume_operations(256)?;
            final_block = reader.read_bits(1)? != 0;
            match reader.read_bits(2)? {
                0 => {
                    reader.align_to_byte();
                    let length = reader.read_bits(16)? as usize;
                    let complement = reader.read_bits(16)? as u16;
                    if (length as u16) ^ complement != u16::MAX {
                        return Err(reader.error(DeflateErrorKind::InvalidStoredLength));
                    }
                    reserve_output(&mut output, length, effective_limit, &reader)?;
                    for _ in 0..length {
                        output.push(reader.read_bits(8)? as u8);
                    }
                }
                1 => {
                    let (literal_tree, distance_tree) =
                        fixed_trees().map_err(|kind| reader.error(kind))?;
                    decode_compressed_block(
                        &mut reader,
                        &mut output,
                        &literal_tree,
                        &distance_tree,
                        effective_limit,
                        allow_damaged_prefix,
                    )?;
                }
                2 => {
                    let (literal_tree, distance_tree) = dynamic_trees(&mut reader)?;
                    decode_compressed_block(
                        &mut reader,
                        &mut output,
                        &literal_tree,
                        &distance_tree,
                        effective_limit,
                        allow_damaged_prefix,
                    )?;
                }
                _ => return Err(reader.error(DeflateErrorKind::InvalidBlockType)),
            }
        }

        if reader.remaining_bits() >= 8 {
            return Err(reader.error(DeflateErrorKind::TrailingData));
        }
        if expected_size.is_some_and(|size| output.len() != size) {
            return Err(reader.error(DeflateErrorKind::SizeMismatch));
        }
        Ok(())
    })();
    if let Err(error) = decode_result
        && !(allow_output_prefix
            && error.kind == DeflateErrorKind::OutputLimit
            && !output.is_empty()
            || allow_damaged_prefix
                && !output.is_empty()
                && !matches!(
                    error.kind,
                    DeflateErrorKind::OutputLimit
                        | DeflateErrorKind::OperationLimit
                        | DeflateErrorKind::AllocationFailed
                )
            || allow_truncated_prefix
                && error.kind == DeflateErrorKind::UnexpectedEnd
                && error.byte_offset >= input.len()
                && !output.is_empty())
    {
        return Err(error);
    }
    let operations_used = max_operations - reader.operations_left;
    Ok((output, operations_used))
}

#[cfg(test)]
mod tests {
    #[cfg(feature = "xps-formats")]
    use super::decompress_output_prefix;
    use super::{DeflateErrorKind, decompress, decompress_with_limit};

    #[test]
    fn canonical_lookup_preserves_symbols_errors_and_operation_budgets() {
        let (literal, distance) = super::fixed_trees().unwrap();
        let mut lengths: Vec<u8> = (1..15).collect();
        lengths.extend([15, 15]);
        let trees = [
            literal,
            distance,
            super::Huffman::from_lengths(&lengths, false, true).unwrap(),
            super::Huffman::from_lengths(&[0, 1], false, false).unwrap(),
            super::Huffman::from_lengths(&[], true, false).unwrap(),
        ];
        for tree in trees {
            for bits in 0..=u16::MAX {
                let bytes = bits.to_le_bytes();
                for budget in [0, 1, 7, 15] {
                    let mut actual = super::BitReader::new(&bytes, budget);
                    let mut reference = super::BitReader::new(&bytes, budget);
                    let expected = (|| {
                        let mut code = 0;
                        for length in 1..=tree.max_length {
                            code |= (reference.read_bits(1)? as u16) << (length - 1);
                            if let Some((_, symbol)) =
                                tree.codes_by_length[length].iter().find(|(canonical, _)| {
                                    canonical.reverse_bits() >> (16 - length) == code
                                })
                            {
                                return Ok(*symbol);
                            }
                        }
                        Err(reference.error(DeflateErrorKind::InvalidCode))
                    })();
                    assert_eq!(tree.decode(&mut actual), expected);
                    assert_eq!(actual.bit_offset, reference.bit_offset);
                    assert_eq!(actual.operations_left, reference.operations_left);
                }
            }
        }
    }

    #[test]
    fn decodes_stored_block() {
        let stream = b"\x01\x05\x00\xfa\xffhello";
        assert_eq!(decompress(stream, 5, Some(5)).unwrap(), b"hello");
    }

    #[test]
    fn decodes_fixed_huffman_block() {
        let stream = [0xcb, 0x48, 0xcd, 0xc9, 0xc9, 0x07, 0x00];
        assert_eq!(decompress(&stream, 5, Some(5)).unwrap(), b"hello");
    }

    #[cfg(feature = "xps-formats")]
    #[test]
    fn decodes_a_bounded_output_prefix() {
        let stream = [0xcb, 0x48, 0xcd, 0xc9, 0xc9, 0x07, 0x00];
        assert_eq!(decompress_output_prefix(&stream, 4).unwrap(), b"hell");
    }

    #[test]
    fn decodes_dynamic_huffman_block() {
        let stream = [
            0xed, 0xcb, 0xd1, 0x09, 0xc0, 0x20, 0x0c, 0x40, 0xc1, 0x55, 0x32, 0x40, 0xe9, 0x24,
            0x2e, 0x21, 0x36, 0x94, 0x80, 0x1a, 0x31, 0x71, 0xff, 0xee, 0x51, 0xde, 0xfd, 0x5f,
            0xf1, 0xad, 0x43, 0x6c, 0xc5, 0x19, 0xf2, 0x78, 0xf7, 0x2d, 0x61, 0x29, 0x75, 0x68,
            0x5e, 0xd2, 0x7c, 0x86, 0xb6, 0xd4, 0x3c, 0x5b, 0xea, 0x63, 0xcb, 0xa2, 0xd9, 0x7c,
            0x45, 0xbb, 0xe5, 0x2d, 0x85, 0x48, 0x24, 0x12, 0x89, 0x44, 0x22, 0x91, 0x48, 0x24,
            0x12, 0x89, 0x44, 0x22, 0x91, 0x48, 0x24, 0x12, 0x89, 0x7f, 0x8f, 0x1f,
        ];
        let expected = b"Lorem ipsum dolor sit amet, consectetur adipiscing elit. ".repeat(100);
        assert_eq!(
            decompress(&stream, expected.len(), Some(expected.len())).unwrap(),
            expected
        );
    }

    #[test]
    fn output_limit_stops_a_bomb_early() {
        let stream = [
            0x73, 0x74, 0x1c, 0x05, 0xa3, 0x60, 0x14, 0x0c, 0x77, 0x00, 0x00,
        ];
        let error = decompress(&stream, 64, None).unwrap_err();
        assert_eq!(error.kind, DeflateErrorKind::OutputLimit);
    }

    #[test]
    fn rejects_bad_stored_length_and_trailing_data() {
        let bad_length = b"\x01\x05\x00\x00\x00hello";
        assert_eq!(
            decompress(bad_length, 32, None).unwrap_err().kind,
            DeflateErrorKind::InvalidStoredLength
        );

        let trailing = [0x03, 0x00, 0x00];
        assert_eq!(
            decompress(&trailing, 32, Some(0)).unwrap_err().kind,
            DeflateErrorKind::TrailingData
        );
    }

    #[test]
    fn operation_budget_bounds_cpu_work() {
        let stream = [0xcb, 0x48, 0xcd, 0xc9, 0xc9, 0x07, 0x00];
        assert_eq!(
            decompress_with_limit(&stream, 5, Some(5), 1)
                .unwrap_err()
                .kind,
            DeflateErrorKind::OperationLimit
        );
    }

    #[test]
    fn accepts_rfc_edge_vectors() {
        let cases: &[(&[u8], &[u8])] = &[
            (&[0x03, 0xfc], b""),
            (&[0x01, 0x00, 0x00, 0xff, 0xff], b""),
            (&[0x01, 0x01, 0x00, 0xfe, 0xff, 0x41], b"A"),
            (&[0x00, 0x01, 0x00, 0xfe, 0xff, 0x41, 0x03, 0x00], b"A"),
            (&[0x73, 0x04, 0x00], b"A"),
            (&[0x73, 0x74, 0xc4, 0x04, 0x00], b"AAAAAAAAAAAAAAAAAAAA"),
            (
                &[
                    0x05, 0xc0, 0x81, 0x08, 0x00, 0x00, 0x00, 0x00, 0x20, 0x7f, 0xeb, 0x03,
                ],
                b"",
            ),
            (
                &[
                    0x05, 0x83, 0x85, 0x00, 0x00, 0x00, 0x00, 0x80, 0x2a, 0x7f, 0xe8, 0x6e,
                ],
                b"",
            ),
        ];
        for &(stream, expected) in cases {
            assert_eq!(
                decompress(stream, expected.len(), Some(expected.len())).unwrap(),
                expected
            );
        }
    }

    #[test]
    fn rejects_invalid_rfc_edge_vectors() {
        let cases: &[&[u8]] = &[
            &[0x07],
            &[0x01, 0x01, 0x00, 0xff, 0xff, 0x41],
            &[0x01, 0x01, 0x00, 0xfe, 0xff],
            &[0x03, 0x02, 0x00],
            &[0x1b, 0x03],
            &[0x03, 0x3e],
            &[0x05, 0x00, 0x92, 0x04],
            &[0x05, 0x00, 0x02, 0x00],
            &[0x05, 0x00, 0x02, 0x24],
            &[
                0x05, 0xc0, 0x81, 0x08, 0x00, 0x00, 0x00, 0x00, 0x20, 0x7f, 0x6e,
            ],
            &[0x05, 0x00, 0x80, 0xe4, 0x7f, 0x1b],
            &[0x03, 0x00, 0x00],
        ];
        for &stream in cases {
            assert!(decompress(stream, 1024, None).is_err(), "{stream:02x?}");
        }
    }
}
