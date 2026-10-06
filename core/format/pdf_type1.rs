use std::collections::BTreeMap;

const MAX_GLYPHS: usize = 4_096;
const MAX_SUBRS: usize = 4_096;
const MAX_CHARSTRING_BYTES: usize = 1 << 20;
const MAX_SUBR_DEPTH: usize = 16;

pub(super) struct ConvertedType1 {
    pub cff: Vec<u8>,
    pub code_to_gid: BTreeMap<u8, u16>,
    pub glyph_count: usize,
    pub units_per_em: Option<u16>,
}

pub(super) fn convert_type1_to_cff(
    bytes: &[u8],
    encoding: &BTreeMap<u8, String>,
) -> Option<ConvertedType1> {
    let pfb;
    let bytes = if bytes.starts_with(b"\x80\x01") {
        pfb = decode_pfb_segments(bytes)?;
        pfb.as_slice()
    } else {
        bytes
    };
    let units_per_em = type1_units_per_em(bytes);
    let shear = type1_font_matrix(bytes)
        .and_then(|matrix| {
            let scale = *matrix.first()?;
            (scale != 0.0).then_some((matrix[2] / scale) as f32)
        })
        .filter(|shear| shear.is_finite())
        .unwrap_or(0.0);
    let encrypted = encrypted_program(bytes)?;
    let mut parser = Tokenizer::new(&encrypted);
    // Type 1 private dictionaries may declare lenIV after Subrs. It applies to
    // every encrypted charstring regardless of declaration order, so discover
    // it before decoding either Subrs or CharStrings.
    let mut len_iv = type1_len_iv(&encrypted).unwrap_or(4);
    let mut subrs = Vec::<Vec<u8>>::new();
    let mut glyphs = Vec::<(String, Vec<u8>)>::new();
    while let Some(token) = parser.token() {
        if token != "/" {
            continue;
        }
        let Some(dictionary_key) = parser.token() else {
            break;
        };
        match dictionary_key.as_str() {
            "lenIV" => len_iv = parser.token()?.parse().ok()?,
            "Subrs" if subrs.is_empty() => {
                let count = parser.token()?.parse::<usize>().ok()?;
                if count > MAX_SUBRS {
                    return None;
                }
                subrs.resize(count, Vec::new());
                parser.seek_token("array")?;
                while parser.peek_token().as_deref() == Some("dup") {
                    parser.token();
                    let index = parser.token()?.parse::<usize>().ok()?;
                    let length = parser.token()?.parse::<usize>().ok()?;
                    parser.token()?;
                    let encoded = parser.binary(length)?;
                    if let Some(slot) = subrs.get_mut(index) {
                        *slot = decrypt_charstring(encoded, len_iv)?;
                    }
                    while matches!(
                        parser.peek_token().as_deref(),
                        Some("noaccess" | "readonly" | "executeonly")
                    ) {
                        parser.token();
                    }
                    if matches!(
                        parser.peek_token().as_deref(),
                        Some("NP" | "ND" | "put" | "|" | "|-")
                    ) {
                        parser.token();
                    }
                }
            }
            "CharStrings" if glyphs.is_empty() => {
                let capacity = parser.token()?.parse::<usize>().ok()?;
                if capacity == 0 {
                    return None;
                }
                parser.seek_token("begin")?;
                while let Some(token) = parser.token() {
                    if token == "end" {
                        break;
                    }
                    if token != "/" {
                        continue;
                    }
                    // A subset may retain the original dictionary capacity.
                    // Bound actual entries instead of rejecting a sparse CJK font.
                    if glyphs.len() >= MAX_GLYPHS {
                        return None;
                    }
                    let Some(name) = parser.token() else {
                        break;
                    };
                    let Some(length) = parser.token().and_then(|value| value.parse::<usize>().ok())
                    else {
                        break;
                    };
                    if parser.token().is_none() {
                        break;
                    }
                    let Some(encoded) = parser
                        .binary(length)
                        .and_then(|bytes| decrypt_charstring(bytes, len_iv))
                    else {
                        break;
                    };
                    if matches!(
                        parser.peek_token().as_deref(),
                        Some("ND" | "NP" | "def" | "|" | "|-")
                    ) {
                        parser.token();
                    }
                    let mut converter = CharStringConverter {
                        shear,
                        ..CharStringConverter::default()
                    };
                    let conversion_failed = converter.convert(&encoded, &subrs, 0).is_none();
                    if converter.output.last() != Some(&14) {
                        converter.output.push(14);
                    }
                    let browser_safe = browser_safe_type2_charstring(&converter.output);
                    let invalid = conversion_failed || !browser_safe;
                    if invalid {
                        // A damaged or unsupported glyph must not discard the
                        // rest of an otherwise usable embedded font.
                        converter.output.clear();
                        converter.output.push(14);
                    }
                    glyphs.push((name, converter.output));
                }
                // Binary eexec sections are commonly followed by encrypted
                // padding and a clear-text trailer. Once CharStrings has been
                // consumed, scanning that arbitrary tail as PostScript can
                // turn an otherwise valid font into a whole-font failure.
                if !glyphs.is_empty() {
                    break;
                }
            }
            _ => {}
        }
    }
    if glyphs.is_empty() {
        return None;
    }
    if let Some(index) = glyphs.iter().position(|(name, _)| name == ".notdef") {
        glyphs.swap(0, index);
    } else {
        glyphs.insert(0, (".notdef".to_owned(), vec![14]));
    }
    let glyph_ids = glyphs
        .iter()
        .enumerate()
        .map(|(gid, (name, _))| Some((name.clone(), u16::try_from(gid).ok()?)))
        .collect::<Option<BTreeMap<_, _>>>()?;
    let code_to_gid: BTreeMap<u8, u16> = encoding
        .iter()
        .filter_map(|(&code, name)| glyph_ids.get(name).map(|&gid| (code, gid)))
        .collect();
    let glyph_count = glyphs.len();
    let cff = build_cff(&glyphs)?;
    Some(ConvertedType1 {
        cff,
        code_to_gid,
        glyph_count,
        units_per_em,
    })
}

pub(super) fn browser_safe_type2_charstring(bytes: &[u8]) -> bool {
    let mut position = 0_usize;
    let mut stack = 0_usize;
    let mut found_width = false;
    let mut found_endchar = false;
    let mut stems = 0_usize;
    let mut vertical_stem_seen = false;
    let mut drawing_started = false;
    while position < bytes.len() {
        let byte = bytes[position];
        position += 1;
        if byte >= 32 || byte == 28 {
            let extra = match byte {
                28 => 3,
                247..=254 => 2,
                255 => 5,
                _ => 1,
            };
            if position
                .checked_add(extra - 1)
                .is_none_or(|end| end > bytes.len())
            {
                return false;
            }
            position += extra - 1;
            stack += 1;
            if stack > 48 {
                return false;
            }
            continue;
        }
        let command = if byte == 12 {
            let Some(&escaped) = bytes.get(position) else {
                return false;
            };
            position += 1;
            0x0c00 | u16::from(escaped)
        } else {
            u16::from(byte)
        };
        let valid = match command {
            1 | 3 => {
                let valid = stack >= 2
                    && (stack.is_multiple_of(2) || !found_width && (stack - 1).is_multiple_of(2))
                    && !drawing_started
                    && (command != 1 || !vertical_stem_seen);
                stems += stack / 2;
                found_width = true;
                vertical_stem_seen |= command == 3;
                valid && stems <= 96
            }
            4 | 22 => {
                let valid = stack == 1 || !found_width && stack == 2;
                found_width = true;
                valid
            }
            21 => {
                let valid = stack == 2 || !found_width && stack == 3;
                found_width = true;
                valid
            }
            5 => found_width && stack >= 2 && stack.is_multiple_of(2),
            6 | 7 => found_width && stack >= 1,
            8 => found_width && stack >= 6 && stack.is_multiple_of(6),
            30 | 31 => {
                found_width
                    && stack >= 4
                    && ((stack - 4).is_multiple_of(8)
                        || stack >= 5 && (stack - 5).is_multiple_of(8)
                        || stack >= 8 && (stack - 8).is_multiple_of(8)
                        || stack >= 9 && (stack - 9).is_multiple_of(8))
            }
            0x0c23 => found_width && stack == 13,
            14 => {
                found_endchar = true;
                found_width = true;
                true
            }
            _ => false,
        };
        if !valid {
            return false;
        }
        drawing_started |= matches!(command, 4..=8 | 21 | 22 | 30 | 31 | 0x0c23);
        stack = 0;
    }
    found_endchar
}

fn type1_len_iv(bytes: &[u8]) -> Option<i32> {
    let marker = b"/lenIV";
    let start = bytes
        .windows(marker.len())
        .position(|window| window == marker)?
        .checked_add(marker.len())?;
    let tail = bytes.get(start..bytes.len().min(start.checked_add(32)?))?;
    let start = tail.iter().position(|byte| !byte.is_ascii_whitespace())?;
    let end = tail[start..]
        .iter()
        .position(|byte| !matches!(byte, b'-' | b'0'..=b'9'))
        .map_or(tail.len(), |length| start + length);
    std::str::from_utf8(tail.get(start..end)?)
        .ok()?
        .parse()
        .ok()
}

fn decode_pfb_segments(bytes: &[u8]) -> Option<Vec<u8>> {
    let mut position = 0_usize;
    let mut output = Vec::new();
    while position < bytes.len() {
        if bytes.get(position) != Some(&0x80) {
            return None;
        }
        let kind = *bytes.get(position.checked_add(1)?)?;
        position = position.checked_add(2)?;
        if kind == 3 {
            return (!output.is_empty()).then_some(output);
        }
        if !matches!(kind, 1 | 2) {
            return None;
        }
        let length = usize::try_from(u32::from_le_bytes(
            bytes
                .get(position..position.checked_add(4)?)?
                .try_into()
                .ok()?,
        ))
        .ok()?;
        position = position.checked_add(4)?;
        let Some(end) = position.checked_add(length) else {
            return None;
        };
        let Some(segment) = bytes.get(position..end) else {
            // Some PDFs retain the final PFB ASCII-segment header after
            // stripping its optional cleartomark payload (Length3 = 0).
            // The preceding font program is still complete and usable.
            return (position == bytes.len() && !output.is_empty()).then_some(output);
        };
        output.extend_from_slice(segment);
        position = end;
    }
    (!output.is_empty()).then_some(output)
}

fn type1_units_per_em(bytes: &[u8]) -> Option<u16> {
    super::font_matrix_units_per_em(&type1_font_matrix(bytes)?)
}

fn type1_font_matrix(bytes: &[u8]) -> Option<Vec<f64>> {
    let marker = b"/FontMatrix";
    let start = bytes
        .windows(marker.len())
        .position(|window| window == marker)?
        .checked_add(marker.len())?;
    let matrix = bytes.get(start..bytes.len().min(start.checked_add(256)?))?;
    let open = matrix.iter().position(|byte| *byte == b'[')?;
    let values = matrix.get(open + 1..)?;
    let close = values.iter().position(|byte| *byte == b']')?;
    let values = std::str::from_utf8(values.get(..close)?)
        .ok()?
        .split_ascii_whitespace()
        .map(str::parse::<f64>)
        .collect::<Result<Vec<_>, _>>()
        .ok()?;
    (values.len() == 6).then_some(values)
}

fn encrypted_program(bytes: &[u8]) -> Option<Vec<u8>> {
    let marker = b"eexec";
    let start = bytes
        .windows(marker.len())
        .position(|window| window == marker)?
        .checked_add(marker.len())?;
    let data = bytes.get(start..)?;
    let first = data.iter().position(|byte| !byte.is_ascii_whitespace())?;
    let data = &data[first..];
    let is_hex = data
        .iter()
        .filter(|byte| !byte.is_ascii_whitespace())
        .take(8)
        .all(u8::is_ascii_hexdigit);
    let ciphertext = if is_hex {
        let mut decoded = Vec::with_capacity(data.len() / 2);
        let mut high = None;
        for &byte in data {
            if !byte.is_ascii_hexdigit() {
                continue;
            }
            let value = (byte as char).to_digit(16)? as u8;
            if let Some(high) = high.take() {
                decoded.push(high << 4 | value);
            } else {
                high = Some(value);
            }
        }
        decoded
    } else {
        data.to_vec()
    };
    decrypt(&ciphertext, 55_665, 4)
}

fn decrypt_charstring(bytes: &[u8], len_iv: i32) -> Option<Vec<u8>> {
    if len_iv == -1 {
        return Some(bytes.to_vec());
    }
    decrypt(bytes, 4_330, usize::try_from(len_iv).ok()?)
}

fn decrypt(bytes: &[u8], key: u16, discard: usize) -> Option<Vec<u8>> {
    if discard > bytes.len() {
        return None;
    }
    let mut state = u32::from(key);
    let mut output = Vec::with_capacity(bytes.len() - discard);
    for (index, &ciphertext) in bytes.iter().enumerate() {
        let plaintext = ciphertext ^ (state >> 8) as u8;
        state = ((u32::from(ciphertext) + state) * 52_845 + 22_719) & 0xffff;
        if index >= discard {
            output.push(plaintext);
        }
    }
    Some(output)
}

struct Tokenizer<'a> {
    bytes: &'a [u8],
    position: usize,
    peeked: Option<String>,
}

impl<'a> Tokenizer<'a> {
    fn new(bytes: &'a [u8]) -> Self {
        Self {
            bytes,
            position: 0,
            peeked: None,
        }
    }

    fn peek_token(&mut self) -> Option<String> {
        if self.peeked.is_none() {
            self.peeked = self.next_token();
        }
        self.peeked.clone()
    }

    fn token(&mut self) -> Option<String> {
        self.peeked.take().or_else(|| self.next_token())
    }

    fn next_token(&mut self) -> Option<String> {
        loop {
            while self
                .bytes
                .get(self.position)
                .is_some_and(u8::is_ascii_whitespace)
            {
                self.position += 1;
            }
            if self.bytes.get(self.position) != Some(&b'%') {
                break;
            }
            while self
                .bytes
                .get(self.position)
                .is_some_and(|byte| !matches!(byte, b'\r' | b'\n'))
            {
                self.position += 1;
            }
        }
        let first = *self.bytes.get(self.position)?;
        if matches!(first, b'/' | b'[' | b']' | b'{' | b'}' | b'(' | b')') {
            self.position += 1;
            return Some((first as char).to_string());
        }
        let start = self.position;
        while self.bytes.get(self.position).is_some_and(|byte| {
            !byte.is_ascii_whitespace()
                && !matches!(byte, b'/' | b'[' | b']' | b'{' | b'}' | b'(' | b')')
        }) {
            self.position += 1;
        }
        (self.position > start)
            .then(|| String::from_utf8_lossy(&self.bytes[start..self.position]).into_owned())
    }

    fn seek_token(&mut self, expected: &str) -> Option<()> {
        while let Some(token) = self.token() {
            if token == expected {
                return Some(());
            }
        }
        None
    }

    fn binary(&mut self, length: usize) -> Option<&'a [u8]> {
        self.peeked = None;
        if self
            .bytes
            .get(self.position)
            .is_some_and(u8::is_ascii_whitespace)
        {
            self.position += 1;
        }
        let end = self.position.checked_add(length)?;
        let result = self.bytes.get(self.position..end)?;
        self.position = end;
        Some(result)
    }
}

#[derive(Default)]
struct CharStringConverter {
    stack: Vec<f32>,
    output: Vec<u8>,
    flexing: bool,
    shear: f32,
}

impl CharStringConverter {
    fn convert(&mut self, bytes: &[u8], subrs: &[Vec<u8>], depth: usize) -> Option<()> {
        if depth > MAX_SUBR_DEPTH || bytes.len() > MAX_CHARSTRING_BYTES {
            return None;
        }
        let mut position = 0;
        while position < bytes.len() {
            let byte = bytes[position];
            position += 1;
            if byte >= 32 {
                let value = match byte {
                    32..=246 => f32::from(byte) - 139.0,
                    247..=250 => {
                        let next = f32::from(*bytes.get(position)?);
                        position += 1;
                        f32::from(byte - 247) * 256.0 + next + 108.0
                    }
                    251..=254 => {
                        let next = f32::from(*bytes.get(position)?);
                        position += 1;
                        -(f32::from(byte - 251) * 256.0) - next - 108.0
                    }
                    255 => {
                        let value =
                            i32::from_be_bytes(bytes.get(position..position + 4)?.try_into().ok()?);
                        position += 4;
                        value as f32
                    }
                    _ => return None,
                };
                self.stack.push(value);
                continue;
            }
            let command = if byte == 12 {
                let escaped = *bytes.get(position)?;
                position += 1;
                0x0c00 | u16::from(escaped)
            } else {
                u16::from(byte)
            };
            match command {
                1 | 3 | 0x0c01 | 0x0c02 => self.stack.clear(),
                4 if self.flexing => {
                    let dy = self.stack.pop()?;
                    self.stack.extend([0.0, dy]);
                }
                4 => {
                    let dy = self.stack.pop()?;
                    self.stack.extend([self.shear * dy, dy]);
                    self.emit(2, &[21])?;
                }
                5 => self.emit_transformed_pairs(2, &[5])?,
                6 => self.emit(1, &[6])?,
                7 => {
                    let dy = self.stack.pop()?;
                    self.stack.extend([self.shear * dy, dy]);
                    self.emit(2, &[5])?;
                }
                8 => self.emit_transformed_pairs(6, &[8])?,
                9 | 0x0c00 => self.stack.clear(),
                10 => {
                    let index = self.stack.pop()?.round();
                    if index < 0.0 {
                        return None;
                    }
                    self.convert(subrs.get(index as usize)?, subrs, depth + 1)?;
                }
                11 => return Some(()),
                13 => {
                    let width = self.stack.pop()?;
                    let side_bearing = self.stack.pop()?;
                    self.stack.extend([width, side_bearing]);
                    self.emit(2, &[22])?;
                }
                14 => {
                    self.output.push(14);
                    self.stack.clear();
                }
                21 => {
                    if self.flexing {
                        if self.stack.len() < 2 {
                            return None;
                        }
                    } else {
                        self.emit_transformed_pairs(2, &[21])?;
                    }
                }
                22 if self.flexing => {
                    let dx = self.stack.pop()?;
                    self.stack.extend([dx, 0.0]);
                }
                22 => self.emit(1, &[22])?,
                30 => {
                    let start = self.stack.len().checked_sub(4)?;
                    let values = self.stack.split_off(start);
                    let [dy1, dx2, dy2, dx3] = values.as_slice() else {
                        return None;
                    };
                    self.stack.extend([
                        self.shear * dy1,
                        *dy1,
                        dx2 + self.shear * dy2,
                        *dy2,
                        *dx3,
                        0.0,
                    ]);
                    self.emit(6, &[8])?;
                }
                31 => {
                    let start = self.stack.len().checked_sub(4)?;
                    let values = self.stack.split_off(start);
                    let [dx1, dx2, dy2, dy3] = values.as_slice() else {
                        return None;
                    };
                    self.stack.extend([
                        *dx1,
                        0.0,
                        dx2 + self.shear * dy2,
                        *dy2,
                        self.shear * dy3,
                        *dy3,
                    ]);
                    self.emit(6, &[8])?;
                }
                0x0c06 => self.emit(4, &[14])?,
                0x0c07 => {
                    self.stack.pop();
                    let width = self.stack.pop()?;
                    let side_y = self.stack.pop()?;
                    let side_x = self.stack.pop()?;
                    self.stack.extend([width, side_x, side_y]);
                    self.transform_pairs_at_stack_end(2, 2)?;
                    self.emit(3, &[21])?;
                }
                0x0c0c => {
                    let denominator = self.stack.pop()?;
                    let numerator = self.stack.pop()?;
                    if denominator == 0.0 {
                        return None;
                    }
                    self.stack.push(numerator / denominator);
                }
                0x0c10 => {
                    let subr = self.stack.pop()?.round() as i32;
                    let arguments = self.stack.pop()?.round() as usize;
                    if subr == 1 && arguments == 0 {
                        self.flexing = true;
                    } else if subr == 0 && arguments == 3 && self.stack.len() >= 17 {
                        let start = self.stack.len() - 17;
                        let values = self.stack.split_off(start);
                        self.stack.extend([
                            values[0] + values[2],
                            values[1] + values[3],
                            values[4],
                            values[5],
                            values[6],
                            values[7],
                            values[8],
                            values[9],
                            values[10],
                            values[11],
                            values[12],
                            values[13],
                            values[14],
                        ]);
                        self.transform_pairs_at_stack_end(13, 12)?;
                        self.emit(13, &[12, 35])?;
                        self.flexing = false;
                        self.stack.extend([values[15], values[16]]);
                    }
                }
                0x0c11 => {}
                0x0c21 => self.stack.clear(),
                _ => return None,
            }
        }
        Some(())
    }

    fn emit(&mut self, count: usize, command: &[u8]) -> Option<()> {
        if self.stack.len() < count {
            return None;
        }
        let start = self.stack.len() - count;
        for value in self.stack.drain(start..) {
            encode_type2_number(value, &mut self.output)?;
        }
        self.output.extend_from_slice(command);
        self.stack.clear();
        Some(())
    }

    fn emit_transformed_pairs(&mut self, count: usize, command: &[u8]) -> Option<()> {
        self.transform_pairs_at_stack_end(count, count)?;
        self.emit(count, command)
    }

    fn transform_pairs_at_stack_end(&mut self, count: usize, pair_values: usize) -> Option<()> {
        if self.stack.len() < count || !pair_values.is_multiple_of(2) || pair_values > count {
            return None;
        }
        let start = self.stack.len() - count;
        for index in (start..start + pair_values).step_by(2) {
            self.stack[index] += self.shear * self.stack[index + 1];
        }
        Some(())
    }
}

fn encode_type2_number(value: f32, output: &mut Vec<u8>) -> Option<()> {
    if value.fract() != 0.0 {
        let fixed = (value * 65_536.0).round();
        if fixed < i32::MIN as f32 || fixed > i32::MAX as f32 {
            return None;
        }
        output.push(255);
        output.extend_from_slice(&(fixed as i32).to_be_bytes());
        return Some(());
    }
    let value = value as i32;
    match value {
        -107..=107 => output.push((value + 139) as u8),
        108..=1_131 => {
            let adjusted = value - 108;
            output.extend([(adjusted / 256 + 247) as u8, (adjusted % 256) as u8]);
        }
        -1_131..=-108 => {
            let adjusted = -value - 108;
            output.extend([(adjusted / 256 + 251) as u8, (adjusted % 256) as u8]);
        }
        _ if i16::try_from(value).is_ok() => {
            output.push(28);
            output.extend_from_slice(&(value as i16).to_be_bytes());
        }
        _ => {
            output.push(255);
            output.extend_from_slice(&value.checked_mul(65_536)?.to_be_bytes());
        }
    }
    Some(())
}

fn build_cff(glyphs: &[(String, Vec<u8>)]) -> Option<Vec<u8>> {
    let name_index = cff_index(&[b"OfficeViewerType1".to_vec()])?;
    let strings = glyphs
        .iter()
        .skip(1)
        .map(|(name, _)| name.as_bytes().to_vec())
        .collect::<Vec<_>>();
    let string_index = cff_index(&strings)?;
    let global_subrs = cff_index(&[])?;
    let mut charset = vec![0];
    for index in 0..glyphs.len().saturating_sub(1) {
        charset.extend_from_slice(&u16::try_from(391 + index).ok()?.to_be_bytes());
    }
    let charstrings = cff_index(
        &glyphs
            .iter()
            .map(|(_, charstring)| charstring.clone())
            .collect::<Vec<_>>(),
    )?;
    let dummy_top = cff_index(&[cff_top_dict(0, 0)])?;
    let charset_offset = 4_usize
        .checked_add(name_index.len())?
        .checked_add(dummy_top.len())?
        .checked_add(string_index.len())?
        .checked_add(global_subrs.len())?;
    let charstrings_offset = charset_offset.checked_add(charset.len())?;
    let top_index = cff_index(&[cff_top_dict(charset_offset, charstrings_offset)])?;
    let mut output = vec![1, 0, 4, 4];
    output.extend(name_index);
    output.extend(top_index);
    output.extend(string_index);
    output.extend(global_subrs);
    output.extend(charset);
    output.extend(charstrings);
    Some(output)
}

fn cff_top_dict(charset: usize, charstrings: usize) -> Vec<u8> {
    let mut dict = Vec::with_capacity(12);
    dict.push(29);
    dict.extend_from_slice(&(charset as i32).to_be_bytes());
    dict.push(15);
    dict.push(29);
    dict.extend_from_slice(&(charstrings as i32).to_be_bytes());
    dict.push(17);
    dict
}

fn cff_index(values: &[Vec<u8>]) -> Option<Vec<u8>> {
    let count = u16::try_from(values.len()).ok()?;
    let mut output = count.to_be_bytes().to_vec();
    if values.is_empty() {
        return Some(output);
    }
    output.push(4);
    let mut offset = 1_u32;
    output.extend_from_slice(&offset.to_be_bytes());
    for value in values {
        offset = offset.checked_add(u32::try_from(value.len()).ok()?)?;
        output.extend_from_slice(&offset.to_be_bytes());
    }
    for value in values {
        output.extend_from_slice(value);
    }
    Some(output)
}

#[cfg(test)]
mod tests {
    use super::{
        CharStringConverter, browser_safe_type2_charstring, convert_type1_to_cff, decrypt,
        type1_units_per_em,
    };
    use std::collections::BTreeMap;

    fn encrypt(bytes: &[u8], key: u16) -> Vec<u8> {
        let mut state = u32::from(key);
        bytes
            .iter()
            .map(|&plaintext| {
                let ciphertext = plaintext ^ (state >> 8) as u8;
                state = ((u32::from(ciphertext) + state) * 52_845 + 22_719) & 0xffff;
                ciphertext
            })
            .collect()
    }

    #[test]
    fn converts_sparse_cjk_dictionary_and_bounds_actual_entries() {
        let converted = convert_type1_to_cff(
            include_bytes!("fixtures/type1-sparse-dictionary.pfa"),
            &BTreeMap::from([(33, "gid307".to_owned()), (34, "gid552".to_owned())]),
        )
        .expect("A 47102-slot dictionary contains only three subset glyphs");
        assert_eq!(converted.glyph_count, 3);
        assert_eq!(converted.code_to_gid.len(), 2);

        for count in [super::MAX_GLYPHS, super::MAX_GLYPHS + 1] {
            let mut private = vec![0, 0, 0, 0];
            private
                .extend_from_slice(b"/lenIV -1 def /Subrs 0 array /CharStrings 1 dict dup begin ");
            for index in 0..count {
                let name = if index == 0 {
                    ".notdef".to_owned()
                } else {
                    format!("g{index}")
                };
                private.extend_from_slice(format!("/{name} 1 RD ").as_bytes());
                private.push(14);
                private.extend_from_slice(b" ND ");
            }
            private.extend_from_slice(b"end");
            let mut font = b"%!PS-AdobeFont-1.0\ncurrentfile eexec\n".to_vec();
            font.extend(encrypt(&private, 55_665));
            assert_eq!(
                convert_type1_to_cff(&font, &BTreeMap::new()).is_some(),
                count <= super::MAX_GLYPHS
            );
        }
    }

    #[test]
    fn converts_a_bounded_type1_program_to_name_keyed_cff() {
        let mut private = vec![0, 0, 0, 0];
        private.extend_from_slice(
            b"/lenIV -1 def /Subrs 0 array /CharStrings 2 dict dup begin \
              /.notdef 1 RD ",
        );
        private.push(14);
        private.extend_from_slice(b" ND /A 4 RD ");
        private.extend_from_slice(&[139, 248, 236, 13]);
        private.extend_from_slice(b" ND end");
        let mut font = b"%!PS-AdobeFont-1.0\ncurrentfile eexec\n".to_vec();
        font.extend(encrypt(&private, 55_665));
        let converted =
            convert_type1_to_cff(&font, &BTreeMap::from([(65, "A".to_owned())])).unwrap();
        assert_eq!(converted.cff[..4], [1, 0, 4, 4]);
        assert_eq!(converted.glyph_count, 2);
        assert_eq!(converted.code_to_gid.get(&65), Some(&1));
    }

    #[test]
    fn parses_compact_type1_subroutine_aliases() {
        let mut private = vec![0, 0, 0, 0];
        private.extend_from_slice(
            b"/lenIV -1 def /Subrs 2 array \
              dup 0 1 -| ",
        );
        private.push(11);
        private.extend_from_slice(b" | dup 1 7 -| ");
        private.extend_from_slice(&[239, 239, 21, 239, 139, 5, 11]);
        private.extend_from_slice(b" | /CharStrings 2 dict dup begin /.notdef 1 -| ");
        private.push(14);
        private.extend_from_slice(b" |- /A 8 -| ");
        private.extend_from_slice(&[139, 248, 136, 13, 140, 10, 9, 14]);
        private.extend_from_slice(b" |- end");
        let mut font = b"%!PS-AdobeFont-1.0\ncurrentfile eexec\n".to_vec();
        font.extend(encrypt(&private, 55_665));

        let converted = convert_type1_to_cff(&font, &BTreeMap::from([(65, "A".to_owned())]))
            .expect("compact Type 1 font");
        let header_size = usize::from(converted.cff[2]);
        let (_, after_names) = super::super::cff_index(&converted.cff, header_size).unwrap();
        let (top_dicts, _) = super::super::cff_index(&converted.cff, after_names).unwrap();
        let offset =
            usize::try_from(super::super::cff_dict_integer_operands(top_dicts[0], 17).unwrap()[0])
                .unwrap();
        let (charstrings, _) = super::super::cff_index(&converted.cff, offset).unwrap();
        let glyph = charstrings[usize::from(converted.code_to_gid[&65])];

        assert_eq!(glyph, [248, 136, 139, 22, 239, 239, 21, 239, 139, 5, 14]);
    }

    #[test]
    fn applies_type1_font_matrix_shear_to_glyph_deltas() {
        let mut converter = CharStringConverter {
            shear: 0.167,
            ..CharStringConverter::default()
        };
        converter
            .convert(&[139, 248, 136, 13, 239, 7, 14], &[], 0)
            .unwrap();
        assert_eq!(converter.output.last(), Some(&14));
        assert!(browser_safe_type2_charstring(&converter.output));
        assert_ne!(converter.output, vec![248, 136, 139, 22, 239, 7, 14]);
    }

    #[test]
    fn normalizes_axis_movements_inside_type1_flex_sequences() {
        let mut converter = CharStringConverter::default();
        let charstring = [
            139, 248, 136, 13, // 0 500 hsbw
            139, 140, 12, 16, // 0 1 callothersubr: start flex
            149, 4, // 10 vmoveto
            159, 22, // 20 hmoveto
            149, 149, 21, 149, 149, 21, 149, 149, 21, 149, 149, 21, 149, 149, 21, 189, 139, 139,
            142, 139, 12, 16, // flex depth/results, 3 0 callothersubr
            12, 17, 12, 17, 12, 33, 14, // pop pop setcurrentpoint endchar
        ];
        converter.convert(&charstring, &[], 0).unwrap();
        assert!(
            converter.output.windows(2).any(|window| window == [12, 35]),
            "converted flex operator missing: {:?}",
            converter.output
        );
        assert!(
            browser_safe_type2_charstring(&converter.output),
            "invalid converted flex: {:?}",
            converter.output
        );
    }

    #[test]
    fn reads_non_default_type1_font_matrix_units() {
        assert_eq!(
            type1_units_per_em(
                b"%!PS-AdobeFont-1.0\n/FontMatrix [0.0005 0 0 0.0005 0 0] readonly def\n"
            ),
            Some(2_000)
        );
    }

    #[test]
    fn type1_cipher_round_trips_after_the_random_prefix() {
        let plaintext = b"randpayload";
        assert_eq!(
            decrypt(&encrypt(plaintext, 4_330), 4_330, 4).unwrap(),
            b"payload",
        );
    }

    #[test]
    fn drops_type1_hints_that_follow_the_type2_side_bearing_move() {
        let mut converter = CharStringConverter::default();
        converter
            .convert(&[139, 248, 136, 13, 139, 149, 1, 139, 159, 3, 14], &[], 0)
            .unwrap();
        assert_eq!(converter.output, [248, 136, 139, 22, 14]);
        assert!(browser_safe_type2_charstring(&converter.output));
        assert!(!browser_safe_type2_charstring(&[
            248, 136, 139, 22, 139, 149, 1, 14,
        ]));
    }

    #[test]
    fn drops_type1_triple_stem_hints_from_type2_charstrings() {
        let mut converter = CharStringConverter::default();
        converter
            .convert(
                &[
                    139, 149, 159, 169, 179, 189, 12, 2, // hstem3
                    139, 149, 159, 169, 179, 189, 12, 1, // vstem3
                    14,
                ],
                &[],
                0,
            )
            .unwrap();
        assert_eq!(converter.output, [14]);
        assert!(browser_safe_type2_charstring(&converter.output));
        assert!(!browser_safe_type2_charstring(&[
            139, 149, 3, 139, 159, 1, 14,
        ]));
    }

    #[test]
    fn validates_three_byte_type2_short_integers() {
        assert!(browser_safe_type2_charstring(&[
            28, 0x04, 0xcd, 28, 0x05, 0x00, 22, 14,
        ]));
        assert!(!browser_safe_type2_charstring(&[28, 0x04]));
    }
}
