//! SHA-256 (FIPS 180-4), the core's twin of `VLA_Digest.bas`'s
//! `VlaSha256Hex` and of `check_rule_coverage.ps1`'s `Get-SourceHash`
//! (PORT.7, slice 7c).
//!
//! Three implementations, one baseline: `tools/check_hash_twin.ps1` holds
//! the PowerShell to the FIPS vectors and the whitespace pair, holds the
//! VBA's test pins to the same digest strings, and (since this slice) holds
//! this file's tests to them too, so no side can drift alone. Depends on
//! nothing, as the VBA does: the core has no dependency, and SD-13 is read
//! off the wasm's import section.
//!
//! The skipping variant is `EnglishSourceHash`'s identity for a phrasebook:
//! the digest of the bytes that are not a tab, a line feed, a carriage
//! return or a space, with how many there were, so that the same text in
//! CRLF and LF, indented or not, has one identity. The build stamp
//! (`Frazaro.Build`) uses it for the sentences and every file used, so a
//! workbook's stamp reads the same whatever line endings its sources had.

const K: [u32; 64] = [
    0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1, 0x923f82a4, 0xab1c5ed5,
    0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3, 0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174,
    0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc, 0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da,
    0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147, 0x06ca6351, 0x14292967,
    0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13, 0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85,
    0xa2bfe8a1, 0xa81a664b, 0xc24b8b70, 0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070,
    0x19a4c116, 0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f, 0x682e6ff3,
    0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208, 0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2,
];

const H0: [u32; 8] = [
    0x6a09e667, 0xbb67ae85, 0x3c6ef372, 0xa54ff53a, 0x510e527f, 0x9b05688c, 0x1f83d9ab, 0x5be0cd19,
];

fn compress(state: &mut [u32; 8], block: &[u8]) {
    let mut w = [0u32; 64];
    for (i, chunk) in block.chunks_exact(4).enumerate().take(16) {
        w[i] = u32::from_be_bytes([chunk[0], chunk[1], chunk[2], chunk[3]]);
    }
    for i in 16..64 {
        let s0 = w[i - 15].rotate_right(7) ^ w[i - 15].rotate_right(18) ^ (w[i - 15] >> 3);
        let s1 = w[i - 2].rotate_right(17) ^ w[i - 2].rotate_right(19) ^ (w[i - 2] >> 10);
        w[i] = w[i - 16]
            .wrapping_add(s0)
            .wrapping_add(w[i - 7])
            .wrapping_add(s1);
    }
    let [mut a, mut b, mut c, mut d, mut e, mut f, mut g, mut h] = *state;
    for i in 0..64 {
        let s1 = e.rotate_right(6) ^ e.rotate_right(11) ^ e.rotate_right(25);
        let ch = (e & f) ^ (!e & g);
        let t1 = h
            .wrapping_add(s1)
            .wrapping_add(ch)
            .wrapping_add(K[i])
            .wrapping_add(w[i]);
        let s0 = a.rotate_right(2) ^ a.rotate_right(13) ^ a.rotate_right(22);
        let maj = (a & b) ^ (a & c) ^ (b & c);
        let t2 = s0.wrapping_add(maj);
        h = g;
        g = f;
        f = e;
        e = d.wrapping_add(t1);
        d = c;
        c = b;
        b = a;
        a = t1.wrapping_add(t2);
    }
    for (s, v) in state.iter_mut().zip([a, b, c, d, e, f, g, h]) {
        *s = s.wrapping_add(v);
    }
}

/// The SHA-256 digest of the bytes.
pub fn sha256(data: &[u8]) -> [u8; 32] {
    let mut state = H0;
    let mut blocks = data.chunks_exact(64);
    for block in &mut blocks {
        compress(&mut state, block);
    }
    // The padding: 0x80, zeros to 56 mod 64, then the bit length big-endian.
    let rest = blocks.remainder();
    let mut tail = Vec::with_capacity(128);
    tail.extend_from_slice(rest);
    tail.push(0x80);
    while tail.len() % 64 != 56 {
        tail.push(0);
    }
    tail.extend_from_slice(&((data.len() as u64) * 8).to_be_bytes());
    for block in tail.chunks_exact(64) {
        compress(&mut state, block);
    }
    let mut out = [0u8; 32];
    for (i, word) in state.iter().enumerate() {
        out[i * 4..i * 4 + 4].copy_from_slice(&word.to_be_bytes());
    }
    out
}

/// A digest as sixty-four upper-case hex digits, as the VBA and the
/// PowerShell spell it.
pub fn hex_upper(digest: &[u8; 32]) -> String {
    digest.iter().map(|b| format!("{b:02X}")).collect()
}

/// `VlaSha256Hex`: the digest of the bytes, in hex.
pub fn sha256_hex(data: &[u8]) -> String {
    hex_upper(&sha256(data))
}

/// `VlaSha256HexSkippingWhitespace`: the digest of the bytes that are not
/// tab, line feed, carriage return or space, and how many there were.
pub fn sha256_hex_skipping_whitespace(data: &[u8]) -> (String, usize) {
    let packed: Vec<u8> = data
        .iter()
        .copied()
        .filter(|b| !matches!(b, 9 | 10 | 13 | 32))
        .collect();
    (sha256_hex(&packed), packed.len())
}

#[cfg(test)]
mod tests {
    use super::*;

    // The baseline of tools/check_hash_twin.ps1: the FIPS 180-4 vectors
    // and the block-boundary cases, which the check reads from this file as
    // it reads them from the VBA's pins. Changing a string here without
    // changing the check's baseline fails the check, on purpose.
    #[test]
    fn the_fips_vectors_and_the_block_boundaries() {
        assert_eq!(
            sha256_hex(b"abc"),
            "BA7816BF8F01CFEA414140DE5DAE2223B00361A396177A9CB410FF61F20015AD"
        );
        assert_eq!(
            sha256_hex(b"abcdbcdecdefdefgefghfghighijhijkijkljklmklmnlmnomnopnopq"),
            "248D6A61D20638B8E5C026930C3E6039A33CE45964FF2167F6ECEDD419DB06C1"
        );
        assert_eq!(
            sha256_hex(b""),
            "E3B0C44298FC1C149AFBF4C8996FB92427AE41E4649B934CA495991B7852B855"
        );
        assert_eq!(
            sha256_hex(&[b'x'; 55]),
            "D5E285683CD4EFC02D021A5C62014694958901005D6F71E89E0989FAC77E4072"
        );
        assert_eq!(
            sha256_hex(&[b'x'; 56]),
            "04C26261370EE7541549D16DEE320C723E3FD14671E66A099AFE0A377C16888E"
        );
        assert_eq!(
            sha256_hex(&[b'x'; 64]),
            "7CE100971F64E7001E8FE5A51973ECDFE1CED42BEFE7EE8D5FD6219506B5393C"
        );
        assert_eq!(
            sha256_hex(&[b'x'; 65]),
            "9537C5FDF120482F7D58D25E9ED583F52C02B4E304EA814DB1633AD565AED7E9"
        );
    }

    #[test]
    fn the_whitespace_pair_has_one_identity() {
        let crlf = sha256_hex_skipping_whitespace(b"(say\t\"hello\")\r\n");
        let lf = sha256_hex_skipping_whitespace(b" ( say \"hello\" ) \n");
        assert_eq!(crlf, lf);
        assert_eq!(
            crlf,
            (
                "F5885BDE1D136B3F76E2F8392A31D8EBF4F84AEA0445CF23F251B9F27A98A728".to_string(),
                12
            )
        );
    }

    #[test]
    fn a_long_message_crosses_many_blocks() {
        // A million 'a's: the FIPS long-message vector.
        let data = vec![b'a'; 1_000_000];
        assert_eq!(
            sha256_hex(&data),
            "CDC76E5C9914FB9281A1C7E284D73E67F1809A48A497200E046D39CCC7112CD0"
        );
    }
}
