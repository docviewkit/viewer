//! Small bounded primitives required by the PDF Standard Security Handler.

use sha2::{Digest, Sha384, Sha512};

const MD5_SHIFTS: [u32; 64] = [
    7, 12, 17, 22, 7, 12, 17, 22, 7, 12, 17, 22, 7, 12, 17, 22, 5, 9, 14, 20, 5, 9, 14, 20, 5, 9,
    14, 20, 5, 9, 14, 20, 4, 11, 16, 23, 4, 11, 16, 23, 4, 11, 16, 23, 4, 11, 16, 23, 6, 10, 15,
    21, 6, 10, 15, 21, 6, 10, 15, 21, 6, 10, 15, 21,
];

const MD5_CONSTANTS: [u32; 64] = [
    0xd76a_a478,
    0xe8c7_b756,
    0x2420_70db,
    0xc1bd_ceee,
    0xf57c_0faf,
    0x4787_c62a,
    0xa830_4613,
    0xfd46_9501,
    0x6980_98d8,
    0x8b44_f7af,
    0xffff_5bb1,
    0x895c_d7be,
    0x6b90_1122,
    0xfd98_7193,
    0xa679_438e,
    0x49b4_0821,
    0xf61e_2562,
    0xc040_b340,
    0x265e_5a51,
    0xe9b6_c7aa,
    0xd62f_105d,
    0x0244_1453,
    0xd8a1_e681,
    0xe7d3_fbc8,
    0x21e1_cde6,
    0xc337_07d6,
    0xf4d5_0d87,
    0x455a_14ed,
    0xa9e3_e905,
    0xfcef_a3f8,
    0x676f_02d9,
    0x8d2a_4c8a,
    0xfffa_3942,
    0x8771_f681,
    0x6d9d_6122,
    0xfde5_380c,
    0xa4be_ea44,
    0x4bde_cfa9,
    0xf6bb_4b60,
    0xbebf_bc70,
    0x289b_7ec6,
    0xeaa1_27fa,
    0xd4ef_3085,
    0x0488_1d05,
    0xd9d4_d039,
    0xe6db_99e5,
    0x1fa2_7cf8,
    0xc4ac_5665,
    0xf429_2244,
    0x432a_ff97,
    0xab94_23a7,
    0xfc93_a039,
    0x655b_59c3,
    0x8f0c_cc92,
    0xffef_f47d,
    0x8584_5dd1,
    0x6fa8_7e4f,
    0xfe2c_e6e0,
    0xa301_4314,
    0x4e08_11a1,
    0xf753_7e82,
    0xbd3a_f235,
    0x2ad7_d2bb,
    0xeb86_d391,
];

pub(super) fn md5(input: &[u8]) -> [u8; 16] {
    let bit_length = (input.len() as u64).wrapping_mul(8);
    let mut padded = input.to_vec();
    padded.push(0x80);
    while padded.len() % 64 != 56 {
        padded.push(0);
    }
    padded.extend_from_slice(&bit_length.to_le_bytes());
    let mut state = [0x6745_2301_u32, 0xefcd_ab89, 0x98ba_dcfe, 0x1032_5476];
    for chunk in padded.chunks_exact(64) {
        let mut words = [0_u32; 16];
        for (word, bytes) in words.iter_mut().zip(chunk.chunks_exact(4)) {
            *word = u32::from_le_bytes(bytes.try_into().expect("four-byte MD5 word"));
        }
        let [mut a, mut b, mut c, mut d] = state;
        for index in 0..64 {
            let (value, word) = match index {
                0..=15 => ((b & c) | (!b & d), index),
                16..=31 => ((d & b) | (!d & c), (5 * index + 1) % 16),
                32..=47 => (b ^ c ^ d, (3 * index + 5) % 16),
                _ => (c ^ (b | !d), (7 * index) % 16),
            };
            let next = a
                .wrapping_add(value)
                .wrapping_add(MD5_CONSTANTS[index])
                .wrapping_add(words[word])
                .rotate_left(MD5_SHIFTS[index])
                .wrapping_add(b);
            a = d;
            d = c;
            c = b;
            b = next;
        }
        state[0] = state[0].wrapping_add(a);
        state[1] = state[1].wrapping_add(b);
        state[2] = state[2].wrapping_add(c);
        state[3] = state[3].wrapping_add(d);
    }
    let mut digest = [0_u8; 16];
    for (target, word) in digest.chunks_exact_mut(4).zip(state) {
        target.copy_from_slice(&word.to_le_bytes());
    }
    digest
}

const SHA256_INITIAL: [u32; 8] = [
    0x6a09_e667,
    0xbb67_ae85,
    0x3c6e_f372,
    0xa54f_f53a,
    0x510e_527f,
    0x9b05_688c,
    0x1f83_d9ab,
    0x5be0_cd19,
];

const SHA256_CONSTANTS: [u32; 64] = [
    0x428a_2f98,
    0x7137_4491,
    0xb5c0_fbcf,
    0xe9b5_dba5,
    0x3956_c25b,
    0x59f1_11f1,
    0x923f_82a4,
    0xab1c_5ed5,
    0xd807_aa98,
    0x1283_5b01,
    0x2431_85be,
    0x550c_7dc3,
    0x72be_5d74,
    0x80de_b1fe,
    0x9bdc_06a7,
    0xc19b_f174,
    0xe49b_69c1,
    0xefbe_4786,
    0x0fc1_9dc6,
    0x240c_a1cc,
    0x2de9_2c6f,
    0x4a74_84aa,
    0x5cb0_a9dc,
    0x76f9_88da,
    0x983e_5152,
    0xa831_c66d,
    0xb003_27c8,
    0xbf59_7fc7,
    0xc6e0_0bf3,
    0xd5a7_9147,
    0x06ca_6351,
    0x1429_2967,
    0x27b7_0a85,
    0x2e1b_2138,
    0x4d2c_6dfc,
    0x5338_0d13,
    0x650a_7354,
    0x766a_0abb,
    0x81c2_c92e,
    0x9272_2c85,
    0xa2bf_e8a1,
    0xa81a_664b,
    0xc24b_8b70,
    0xc76c_51a3,
    0xd192_e819,
    0xd699_0624,
    0xf40e_3585,
    0x106a_a070,
    0x19a4_c116,
    0x1e37_6c08,
    0x2748_774c,
    0x34b0_bcb5,
    0x391c_0cb3,
    0x4ed8_aa4a,
    0x5b9c_ca4f,
    0x682e_6ff3,
    0x748f_82ee,
    0x78a5_636f,
    0x84c8_7814,
    0x8cc7_0208,
    0x90be_fffa,
    0xa450_6ceb,
    0xbef9_a3f7,
    0xc671_78f2,
];

pub(super) fn sha256(input: &[u8]) -> [u8; 32] {
    let bit_length = (input.len() as u64).wrapping_mul(8);
    let mut padded = input.to_vec();
    padded.push(0x80);
    while padded.len() % 64 != 56 {
        padded.push(0);
    }
    padded.extend_from_slice(&bit_length.to_be_bytes());
    let mut state = SHA256_INITIAL;
    for chunk in padded.chunks_exact(64) {
        let mut words = [0_u32; 64];
        for (word, bytes) in words[..16].iter_mut().zip(chunk.chunks_exact(4)) {
            *word = u32::from_be_bytes(bytes.try_into().expect("four-byte SHA-256 word"));
        }
        for index in 16..64 {
            let s0 = words[index - 15].rotate_right(7)
                ^ words[index - 15].rotate_right(18)
                ^ (words[index - 15] >> 3);
            let s1 = words[index - 2].rotate_right(17)
                ^ words[index - 2].rotate_right(19)
                ^ (words[index - 2] >> 10);
            words[index] = words[index - 16]
                .wrapping_add(s0)
                .wrapping_add(words[index - 7])
                .wrapping_add(s1);
        }
        let [mut a, mut b, mut c, mut d, mut e, mut f, mut g, mut h] = state;
        for index in 0..64 {
            let sum1 = e.rotate_right(6) ^ e.rotate_right(11) ^ e.rotate_right(25);
            let choice = (e & f) ^ (!e & g);
            let temp1 = h
                .wrapping_add(sum1)
                .wrapping_add(choice)
                .wrapping_add(SHA256_CONSTANTS[index])
                .wrapping_add(words[index]);
            let sum0 = a.rotate_right(2) ^ a.rotate_right(13) ^ a.rotate_right(22);
            let majority = (a & b) ^ (a & c) ^ (b & c);
            let temp2 = sum0.wrapping_add(majority);
            h = g;
            g = f;
            f = e;
            e = d.wrapping_add(temp1);
            d = c;
            c = b;
            b = a;
            a = temp1.wrapping_add(temp2);
        }
        for (target, value) in state.iter_mut().zip([a, b, c, d, e, f, g, h]) {
            *target = target.wrapping_add(value);
        }
    }
    let mut digest = [0_u8; 32];
    for (target, word) in digest.chunks_exact_mut(4).zip(state) {
        target.copy_from_slice(&word.to_be_bytes());
    }
    digest
}

pub(super) fn sha384(input: &[u8]) -> [u8; 48] {
    Sha384::digest(input).into()
}

pub(super) fn sha512(input: &[u8]) -> [u8; 64] {
    Sha512::digest(input).into()
}

pub(super) fn rc4(key: &[u8], input: &[u8]) -> Vec<u8> {
    let mut state = [0_u8; 256];
    for (index, value) in state.iter_mut().enumerate() {
        *value = index as u8;
    }
    let mut j = 0_usize;
    for index in 0..256 {
        j = (j + state[index] as usize + key[index % key.len()] as usize) & 255;
        state.swap(index, j);
    }
    let (mut i, mut j) = (0_usize, 0_usize);
    input
        .iter()
        .map(|byte| {
            i = (i + 1) & 255;
            j = (j + state[i] as usize) & 255;
            state.swap(i, j);
            byte ^ state[(state[i] as usize + state[j] as usize) & 255]
        })
        .collect()
}

const AES_SBOX: [u8; 256] = [
    0x63, 0x7c, 0x77, 0x7b, 0xf2, 0x6b, 0x6f, 0xc5, 0x30, 0x01, 0x67, 0x2b, 0xfe, 0xd7, 0xab, 0x76,
    0xca, 0x82, 0xc9, 0x7d, 0xfa, 0x59, 0x47, 0xf0, 0xad, 0xd4, 0xa2, 0xaf, 0x9c, 0xa4, 0x72, 0xc0,
    0xb7, 0xfd, 0x93, 0x26, 0x36, 0x3f, 0xf7, 0xcc, 0x34, 0xa5, 0xe5, 0xf1, 0x71, 0xd8, 0x31, 0x15,
    0x04, 0xc7, 0x23, 0xc3, 0x18, 0x96, 0x05, 0x9a, 0x07, 0x12, 0x80, 0xe2, 0xeb, 0x27, 0xb2, 0x75,
    0x09, 0x83, 0x2c, 0x1a, 0x1b, 0x6e, 0x5a, 0xa0, 0x52, 0x3b, 0xd6, 0xb3, 0x29, 0xe3, 0x2f, 0x84,
    0x53, 0xd1, 0x00, 0xed, 0x20, 0xfc, 0xb1, 0x5b, 0x6a, 0xcb, 0xbe, 0x39, 0x4a, 0x4c, 0x58, 0xcf,
    0xd0, 0xef, 0xaa, 0xfb, 0x43, 0x4d, 0x33, 0x85, 0x45, 0xf9, 0x02, 0x7f, 0x50, 0x3c, 0x9f, 0xa8,
    0x51, 0xa3, 0x40, 0x8f, 0x92, 0x9d, 0x38, 0xf5, 0xbc, 0xb6, 0xda, 0x21, 0x10, 0xff, 0xf3, 0xd2,
    0xcd, 0x0c, 0x13, 0xec, 0x5f, 0x97, 0x44, 0x17, 0xc4, 0xa7, 0x7e, 0x3d, 0x64, 0x5d, 0x19, 0x73,
    0x60, 0x81, 0x4f, 0xdc, 0x22, 0x2a, 0x90, 0x88, 0x46, 0xee, 0xb8, 0x14, 0xde, 0x5e, 0x0b, 0xdb,
    0xe0, 0x32, 0x3a, 0x0a, 0x49, 0x06, 0x24, 0x5c, 0xc2, 0xd3, 0xac, 0x62, 0x91, 0x95, 0xe4, 0x79,
    0xe7, 0xc8, 0x37, 0x6d, 0x8d, 0xd5, 0x4e, 0xa9, 0x6c, 0x56, 0xf4, 0xea, 0x65, 0x7a, 0xae, 0x08,
    0xba, 0x78, 0x25, 0x2e, 0x1c, 0xa6, 0xb4, 0xc6, 0xe8, 0xdd, 0x74, 0x1f, 0x4b, 0xbd, 0x8b, 0x8a,
    0x70, 0x3e, 0xb5, 0x66, 0x48, 0x03, 0xf6, 0x0e, 0x61, 0x35, 0x57, 0xb9, 0x86, 0xc1, 0x1d, 0x9e,
    0xe1, 0xf8, 0x98, 0x11, 0x69, 0xd9, 0x8e, 0x94, 0x9b, 0x1e, 0x87, 0xe9, 0xce, 0x55, 0x28, 0xdf,
    0x8c, 0xa1, 0x89, 0x0d, 0xbf, 0xe6, 0x42, 0x68, 0x41, 0x99, 0x2d, 0x0f, 0xb0, 0x54, 0xbb, 0x16,
];

fn aes_inverse_sbox() -> [u8; 256] {
    let mut inverse = [0_u8; 256];
    for (index, value) in AES_SBOX.iter().enumerate() {
        inverse[*value as usize] = index as u8;
    }
    inverse
}

fn aes_expand_key(key: &[u8]) -> Option<(Vec<u8>, usize)> {
    let words = match key.len() {
        16 => 4,
        32 => 8,
        _ => return None,
    };
    let rounds = words + 6;
    let mut expanded = vec![0_u8; 16 * (rounds + 1)];
    expanded[..key.len()].copy_from_slice(key);
    let mut generated = key.len();
    let mut rcon = 1_u8;
    let mut temp = [0_u8; 4];
    while generated < expanded.len() {
        temp.copy_from_slice(&expanded[generated - 4..generated]);
        if generated.is_multiple_of(key.len()) {
            temp.rotate_left(1);
            for byte in &mut temp {
                *byte = AES_SBOX[*byte as usize];
            }
            temp[0] ^= rcon;
            rcon = xtime(rcon);
        } else if key.len() == 32 && generated % key.len() == 16 {
            for byte in &mut temp {
                *byte = AES_SBOX[*byte as usize];
            }
        }
        for byte in temp {
            expanded[generated] = expanded[generated - key.len()] ^ byte;
            generated += 1;
        }
    }
    Some((expanded, rounds))
}

fn xtime(value: u8) -> u8 {
    (value << 1) ^ if value & 0x80 != 0 { 0x1b } else { 0 }
}

fn gf_mul(mut left: u8, mut right: u8) -> u8 {
    let mut result = 0_u8;
    while right != 0 {
        if right & 1 != 0 {
            result ^= left;
        }
        left = xtime(left);
        right >>= 1;
    }
    result
}

fn aes_decrypt_block(block: &mut [u8; 16], expanded: &[u8], rounds: usize) {
    let inverse = aes_inverse_sbox();
    for (byte, key) in block
        .iter_mut()
        .zip(&expanded[16 * rounds..16 * (rounds + 1)])
    {
        *byte ^= key;
    }
    for round in (1..rounds).rev() {
        let prior = *block;
        block[0] = inverse[prior[0] as usize];
        block[1] = inverse[prior[13] as usize];
        block[2] = inverse[prior[10] as usize];
        block[3] = inverse[prior[7] as usize];
        block[4] = inverse[prior[4] as usize];
        block[5] = inverse[prior[1] as usize];
        block[6] = inverse[prior[14] as usize];
        block[7] = inverse[prior[11] as usize];
        block[8] = inverse[prior[8] as usize];
        block[9] = inverse[prior[5] as usize];
        block[10] = inverse[prior[2] as usize];
        block[11] = inverse[prior[15] as usize];
        block[12] = inverse[prior[12] as usize];
        block[13] = inverse[prior[9] as usize];
        block[14] = inverse[prior[6] as usize];
        block[15] = inverse[prior[3] as usize];
        for index in 0..16 {
            block[index] ^= expanded[16 * round + index];
        }
        for column in 0..4 {
            let offset = column * 4;
            let values = [
                block[offset],
                block[offset + 1],
                block[offset + 2],
                block[offset + 3],
            ];
            block[offset] = gf_mul(values[0], 14)
                ^ gf_mul(values[1], 11)
                ^ gf_mul(values[2], 13)
                ^ gf_mul(values[3], 9);
            block[offset + 1] = gf_mul(values[0], 9)
                ^ gf_mul(values[1], 14)
                ^ gf_mul(values[2], 11)
                ^ gf_mul(values[3], 13);
            block[offset + 2] = gf_mul(values[0], 13)
                ^ gf_mul(values[1], 9)
                ^ gf_mul(values[2], 14)
                ^ gf_mul(values[3], 11);
            block[offset + 3] = gf_mul(values[0], 11)
                ^ gf_mul(values[1], 13)
                ^ gf_mul(values[2], 9)
                ^ gf_mul(values[3], 14);
        }
    }
    let prior = *block;
    block[0] = inverse[prior[0] as usize];
    block[1] = inverse[prior[13] as usize];
    block[2] = inverse[prior[10] as usize];
    block[3] = inverse[prior[7] as usize];
    block[4] = inverse[prior[4] as usize];
    block[5] = inverse[prior[1] as usize];
    block[6] = inverse[prior[14] as usize];
    block[7] = inverse[prior[11] as usize];
    block[8] = inverse[prior[8] as usize];
    block[9] = inverse[prior[5] as usize];
    block[10] = inverse[prior[2] as usize];
    block[11] = inverse[prior[15] as usize];
    block[12] = inverse[prior[12] as usize];
    block[13] = inverse[prior[9] as usize];
    block[14] = inverse[prior[6] as usize];
    block[15] = inverse[prior[3] as usize];
    for index in 0..16 {
        block[index] ^= expanded[index];
    }
}

fn aes_encrypt_block(block: &mut [u8; 16], expanded: &[u8], rounds: usize) {
    for (byte, key) in block.iter_mut().zip(&expanded[..16]) {
        *byte ^= key;
    }
    for round in 1..=rounds {
        for byte in block.iter_mut() {
            *byte = AES_SBOX[*byte as usize];
        }
        let prior = *block;
        block[0] = prior[0];
        block[1] = prior[5];
        block[2] = prior[10];
        block[3] = prior[15];
        block[4] = prior[4];
        block[5] = prior[9];
        block[6] = prior[14];
        block[7] = prior[3];
        block[8] = prior[8];
        block[9] = prior[13];
        block[10] = prior[2];
        block[11] = prior[7];
        block[12] = prior[12];
        block[13] = prior[1];
        block[14] = prior[6];
        block[15] = prior[11];
        if round != rounds {
            for column in 0..4 {
                let offset = column * 4;
                let values = [
                    block[offset],
                    block[offset + 1],
                    block[offset + 2],
                    block[offset + 3],
                ];
                block[offset] = gf_mul(values[0], 2) ^ gf_mul(values[1], 3) ^ values[2] ^ values[3];
                block[offset + 1] =
                    values[0] ^ gf_mul(values[1], 2) ^ gf_mul(values[2], 3) ^ values[3];
                block[offset + 2] =
                    values[0] ^ values[1] ^ gf_mul(values[2], 2) ^ gf_mul(values[3], 3);
                block[offset + 3] =
                    gf_mul(values[0], 3) ^ values[1] ^ values[2] ^ gf_mul(values[3], 2);
            }
        }
        for (byte, key) in block
            .iter_mut()
            .zip(&expanded[16 * round..16 * (round + 1)])
        {
            *byte ^= key;
        }
    }
}

pub(super) fn aes_cbc_encrypt(key: &[u8], iv: &[u8; 16], input: &[u8]) -> Option<Vec<u8>> {
    if !input.len().is_multiple_of(16) {
        return None;
    }
    let (expanded, rounds) = aes_expand_key(key)?;
    let mut output = Vec::with_capacity(input.len());
    let mut previous = *iv;
    for chunk in input.chunks_exact(16) {
        let mut block: [u8; 16] = chunk.try_into().ok()?;
        for index in 0..16 {
            block[index] ^= previous[index];
        }
        aes_encrypt_block(&mut block, &expanded, rounds);
        output.extend_from_slice(&block);
        previous = block;
    }
    Some(output)
}

pub(super) fn aes_cbc_decrypt(
    key: &[u8],
    iv: &[u8; 16],
    input: &[u8],
    remove_padding: bool,
) -> Option<Vec<u8>> {
    if !input.len().is_multiple_of(16) {
        return None;
    }
    let (expanded, rounds) = aes_expand_key(key)?;
    let mut output = Vec::with_capacity(input.len());
    let mut previous = *iv;
    for chunk in input.chunks_exact(16) {
        let encrypted: [u8; 16] = chunk.try_into().ok()?;
        let mut block = encrypted;
        aes_decrypt_block(&mut block, &expanded, rounds);
        for index in 0..16 {
            block[index] ^= previous[index];
        }
        output.extend_from_slice(&block);
        previous = encrypted;
    }
    if remove_padding && !output.is_empty() {
        let padding = *output.last()? as usize;
        if padding == 0
            || padding > 16
            || padding > output.len()
            || !output[output.len() - padding..]
                .iter()
                .all(|byte| *byte as usize == padding)
        {
            return None;
        }
        output.truncate(output.len() - padding);
    }
    Some(output)
}

#[cfg(test)]
mod tests {
    use super::{aes_cbc_decrypt, aes_cbc_encrypt, md5, rc4, sha256};
    use super::{sha384, sha512};

    fn hex(value: &str) -> Vec<u8> {
        value
            .as_bytes()
            .chunks_exact(2)
            .map(|pair| u8::from_str_radix(std::str::from_utf8(pair).unwrap(), 16).unwrap())
            .collect()
    }

    #[test]
    fn cryptographic_primitives_match_published_vectors() {
        assert_eq!(
            md5(b"abc").to_vec(),
            hex("900150983cd24fb0d6963f7d28e17f72")
        );
        assert_eq!(
            sha256(b"abc").to_vec(),
            hex("ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad")
        );
        assert_eq!(
            sha384(b"abc").to_vec(),
            hex(concat!(
                "cb00753f45a35e8bb5a03d699ac65007",
                "272c32ab0eded1631a8b605a43ff5bed",
                "8086072ba1e7cc2358baeca134c825a7",
            )),
        );
        assert_eq!(
            sha512(b"abc").to_vec(),
            hex(concat!(
                "ddaf35a193617abacc417349ae204131",
                "12e6fa4e89a97ea20a9eeee64b55d39a",
                "2192992a274fc1a836ba3c23a3feebbd",
                "454d4423643ce80e2a9ac94fa54ca49f",
            )),
        );
        assert_eq!(rc4(b"Key", b"Plaintext"), hex("bbf316e8d940af0ad3"));
        let key = hex("2b7e151628aed2a6abf7158809cf4f3c");
        let iv: [u8; 16] = hex("000102030405060708090a0b0c0d0e0f").try_into().unwrap();
        let encrypted = hex("7649abac8119b246cee98e9b12e9197d");
        assert_eq!(
            aes_cbc_decrypt(&key, &iv, &encrypted, false).unwrap(),
            hex("6bc1bee22e409f96e93d7e117393172a"),
        );
        assert_eq!(
            aes_cbc_encrypt(
                &hex("2b7e151628aed2a6abf7158809cf4f3c"),
                &iv,
                &hex("6bc1bee22e409f96e93d7e117393172a"),
            )
            .unwrap(),
            hex("7649abac8119b246cee98e9b12e9197d"),
        );
        let key = hex(concat!(
            "603deb1015ca71be2b73aef0857d7781",
            "1f352c073b6108d72d9810a30914dff4",
        ));
        let encrypted = hex("f58c4c04d6e5f1ba779eabfb5f7bfbd6");
        assert_eq!(
            aes_cbc_decrypt(&key, &iv, &encrypted, false).unwrap(),
            hex("6bc1bee22e409f96e93d7e117393172a"),
        );
    }
}
