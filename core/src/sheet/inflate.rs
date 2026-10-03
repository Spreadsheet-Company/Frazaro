//! Inflate, RFC 1951's decoder, for the parts of an existing workbook the
//! writer must read before it edits them (PORT.7, slice 7d).
//!
//! New ground, said plainly: the reference opens a workbook through Excel.
//! `frazaro build --into model.xlsx` copies every part of the model as the
//! compressed bytes it already is and edits four or five of them (the
//! workbook, its relationships, the content types, the styles, the cell
//! metadata when there is one), which it must first decode. Decoding alone
//! is enough: what the writer itself writes is stored, so no encoder, no
//! dependency, and the bytes stay deterministic. Decoding is canonical, so
//! this is the one implementation of the format everyone agrees on, in the
//! shape of zlib's puff.c: canonical Huffman codes read bit by bit, which is
//! slow by zlib's standards and fast enough for the parts of a workbook.
//!
//! Bounds, since the model is a file someone else made: the output never
//! grows past `limit`, a back-reference never reaches before the output's
//! start, and the input ending early is an error, not a panic.

const MAX_BITS: usize = 15;
const FIXED_LIT_SYMBOLS: usize = 288;
const DIST_SYMBOLS: usize = 30;

const LENGTH_BASE: [u16; 29] = [
    3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131,
    163, 195, 227, 258,
];
const LENGTH_EXTRA: [u8; 29] = [
    0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0,
];
const DIST_BASE: [u16; 30] = [
    1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537,
    2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577,
];
const DIST_EXTRA: [u8; 30] = [
    0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13,
    13,
];
/// The order the code-length code's lengths are stored in a dynamic block.
const CL_ORDER: [usize; 19] = [
    16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15,
];

struct Bits<'a> {
    data: &'a [u8],
    pos: usize,
    bit_buf: u32,
    bit_count: u32,
}

impl<'a> Bits<'a> {
    fn new(data: &'a [u8]) -> Bits<'a> {
        Bits {
            data,
            pos: 0,
            bit_buf: 0,
            bit_count: 0,
        }
    }

    fn need(&mut self, n: u32) -> Result<(), &'static str> {
        while self.bit_count < n {
            let byte = *self.data.get(self.pos).ok_or("the data ended early")?;
            self.pos += 1;
            self.bit_buf |= u32::from(byte) << self.bit_count;
            self.bit_count += 8;
        }
        Ok(())
    }

    fn bits(&mut self, n: u32) -> Result<u32, &'static str> {
        if n == 0 {
            return Ok(0);
        }
        self.need(n)?;
        let v = self.bit_buf & ((1u32 << n) - 1);
        self.bit_buf >>= n;
        self.bit_count -= n;
        Ok(v)
    }

    /// Drop the bits left in the current byte, for a stored block.
    fn align(&mut self) {
        self.bit_buf = 0;
        self.bit_count = 0;
    }
}

/// A canonical Huffman code: how many codes of each length, and the
/// symbols in code order.
struct Huffman {
    counts: [u16; MAX_BITS + 1],
    symbols: Vec<u16>,
}

impl Huffman {
    /// From the code length of each symbol; a length of zero is an unused
    /// symbol. Over-subscribed codes are refused; incomplete codes are
    /// allowed, as the format allows them for a single distance code.
    fn new(lengths: &[u8]) -> Result<Huffman, &'static str> {
        let mut counts = [0u16; MAX_BITS + 1];
        for &l in lengths {
            counts[usize::from(l)] += 1;
        }
        counts[0] = 0;
        let mut left: i32 = 1;
        for &c in &counts[1..] {
            left <<= 1;
            left -= i32::from(c);
            if left < 0 {
                return Err("an over-subscribed Huffman code");
            }
        }
        let mut offsets = [0u16; MAX_BITS + 2];
        for len in 1..=MAX_BITS {
            offsets[len + 1] = offsets[len] + counts[len];
        }
        let mut symbols = vec![0u16; lengths.len()];
        for (sym, &l) in lengths.iter().enumerate() {
            if l != 0 {
                symbols[usize::from(offsets[usize::from(l)])] = sym as u16;
                offsets[usize::from(l)] += 1;
            }
        }
        Ok(Huffman { counts, symbols })
    }

    fn decode(&self, bits: &mut Bits) -> Result<u16, &'static str> {
        let mut code: i32 = 0;
        let mut first: i32 = 0;
        let mut index: i32 = 0;
        for len in 1..=MAX_BITS {
            code |= bits.bits(1)? as i32;
            let count = i32::from(self.counts[len]);
            if code - count < first {
                return Ok(self.symbols[(index + (code - first)) as usize]);
            }
            index += count;
            first += count;
            first <<= 1;
            code <<= 1;
        }
        Err("an invalid Huffman code")
    }
}

fn fixed_codes() -> (Huffman, Huffman) {
    let mut lengths = [0u8; FIXED_LIT_SYMBOLS];
    for (i, l) in lengths.iter_mut().enumerate() {
        *l = match i {
            0..=143 => 8,
            144..=255 => 9,
            256..=279 => 7,
            _ => 8,
        };
    }
    let lit = Huffman::new(&lengths).expect("the fixed literal code is well formed");
    let dist = Huffman::new(&[5u8; DIST_SYMBOLS]).expect("the fixed distance code is well formed");
    (lit, dist)
}

fn dynamic_codes(bits: &mut Bits) -> Result<(Huffman, Huffman), &'static str> {
    let nlen = bits.bits(5)? as usize + 257;
    let ndist = bits.bits(5)? as usize + 1;
    let ncode = bits.bits(4)? as usize + 4;
    if nlen > 286 || ndist > 30 {
        return Err("a dynamic block with too many codes");
    }
    let mut cl_lengths = [0u8; 19];
    for &i in CL_ORDER.iter().take(ncode) {
        cl_lengths[i] = bits.bits(3)? as u8;
    }
    let cl = Huffman::new(&cl_lengths)?;
    let mut lengths = vec![0u8; nlen + ndist];
    let mut i = 0;
    while i < nlen + ndist {
        let sym = cl.decode(bits)?;
        match sym {
            0..=15 => {
                lengths[i] = sym as u8;
                i += 1;
            }
            16 => {
                if i == 0 {
                    return Err("a repeat with no length before it");
                }
                let prev = lengths[i - 1];
                let n = 3 + bits.bits(2)? as usize;
                if i + n > lengths.len() {
                    return Err("a repeat past the end of the code lengths");
                }
                lengths[i..i + n].fill(prev);
                i += n;
            }
            17 | 18 => {
                let n = if sym == 17 {
                    3 + bits.bits(3)? as usize
                } else {
                    11 + bits.bits(7)? as usize
                };
                if i + n > lengths.len() {
                    return Err("a repeat past the end of the code lengths");
                }
                i += n; // already zero
            }
            _ => return Err("an invalid code length code"),
        }
    }
    if lengths[256] == 0 {
        return Err("a dynamic block with no end-of-block code");
    }
    let lit = Huffman::new(&lengths[..nlen])?;
    let dist = Huffman::new(&lengths[nlen..])?;
    Ok((lit, dist))
}

fn inflate_codes(
    bits: &mut Bits,
    out: &mut Vec<u8>,
    lit: &Huffman,
    dist: &Huffman,
    limit: usize,
) -> Result<(), &'static str> {
    loop {
        let sym = lit.decode(bits)? as usize;
        if sym < 256 {
            if out.len() >= limit {
                return Err("longer than the limit allows");
            }
            out.push(sym as u8);
        } else if sym == 256 {
            return Ok(());
        } else {
            let li = sym - 257;
            if li >= LENGTH_BASE.len() {
                return Err("an invalid length code");
            }
            let len =
                usize::from(LENGTH_BASE[li]) + bits.bits(u32::from(LENGTH_EXTRA[li]))? as usize;
            let di = dist.decode(bits)? as usize;
            if di >= DIST_BASE.len() {
                return Err("an invalid distance code");
            }
            let d = usize::from(DIST_BASE[di]) + bits.bits(u32::from(DIST_EXTRA[di]))? as usize;
            if d > out.len() {
                return Err("a reference before the start of the output");
            }
            if out.len() + len > limit {
                return Err("longer than the limit allows");
            }
            let start = out.len() - d;
            for k in 0..len {
                let b = out[start + k];
                out.push(b);
            }
        }
    }
}

/// The bytes a raw deflate stream encodes, up to `limit` bytes; `hint` is
/// the size the container declares, used to size the buffer and nothing
/// else. An error names what was wrong with the stream.
pub fn inflate(data: &[u8], hint: usize, limit: usize) -> Result<Vec<u8>, &'static str> {
    let mut bits = Bits::new(data);
    let mut out: Vec<u8> = Vec::with_capacity(hint.min(limit));
    loop {
        let last = bits.bits(1)? == 1;
        match bits.bits(2)? {
            0 => {
                bits.align();
                let len = usize::from(*bits.data.get(bits.pos).ok_or("the data ended early")?)
                    | (usize::from(*bits.data.get(bits.pos + 1).ok_or("the data ended early")?)
                        << 8);
                let nlen = usize::from(*bits.data.get(bits.pos + 2).ok_or("the data ended early")?)
                    | (usize::from(*bits.data.get(bits.pos + 3).ok_or("the data ended early")?)
                        << 8);
                if len != (!nlen & 0xFFFF) {
                    return Err("a stored block whose length check failed");
                }
                bits.pos += 4;
                let block = bits
                    .data
                    .get(bits.pos..bits.pos + len)
                    .ok_or("the data ended early")?;
                if out.len() + len > limit {
                    return Err("longer than the limit allows");
                }
                out.extend_from_slice(block);
                bits.pos += len;
            }
            1 => {
                let (lit, dist) = fixed_codes();
                inflate_codes(&mut bits, &mut out, &lit, &dist, limit)?;
            }
            2 => {
                let (lit, dist) = dynamic_codes(&mut bits)?;
                inflate_codes(&mut bits, &mut out, &lit, &dist, limit)?;
            }
            _ => return Err("an invalid block type"),
        }
        if last {
            return Ok(out);
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    // Vectors made with .NET's DeflateStream (2026-10-03): "abc" at
    // Optimal is one fixed-Huffman block; at NoCompression one stored
    // block; a repetitive text is fixed codes with back-references; a
    // worksheet's XML is a dynamic block.
    const ABC_FIXED: &[u8] = &[0x4b, 0x4c, 0x4a, 0x06, 0x00];
    const ABC_STORED: &[u8] = &[0x01, 0x03, 0x00, 0xfc, 0xff, 0x61, 0x62, 0x63];
    const REP_FIXED: &[u8] = &[
        0x73, 0x2b, 0x4a, 0xac, 0x4a, 0x2c, 0xca, 0x57, 0x28, 0x2f, 0xca, 0x2c, 0x49, 0x2d, 0x56,
        0x28, 0xcf, 0x2f, 0xca, 0x4e, 0xca, 0xcf, 0xcf, 0x2e, 0xd6, 0x53, 0x70, 0x1b, 0x95, 0xc1,
        0x23, 0x03, 0x00,
    ];
    const XML_DYNAMIC: &[u8] = &[
        0x75, 0xd0, 0xd1, 0x0a, 0xc2, 0x20, 0x14, 0x06, 0xe0, 0x57, 0x19, 0xde, 0xd7, 0x99, 0x2e,
        0x22, 0xc2, 0x09, 0x1b, 0xbd, 0x88, 0x98, 0xcb, 0x68, 0x4e, 0x51, 0xd9, 0x7a, 0xfc, 0xce,
        0x24, 0x64, 0x2c, 0xba, 0xf3, 0xfc, 0xff, 0xf1, 0x03, 0xe5, 0x8b, 0x0b, 0xaf, 0x68, 0xb4,
        0x4e, 0xd5, 0xdb, 0x8e, 0x53, 0x6c, 0x89, 0x49, 0xc9, 0x5f, 0x01, 0xa2, 0x32, 0xda, 0xca,
        0x78, 0x74, 0x5e, 0x4f, 0xd8, 0x0c, 0x2e, 0x58, 0x99, 0x70, 0x0c, 0x0f, 0x88, 0x3e, 0x68,
        0x79, 0xcf, 0x97, 0xec, 0x08, 0xac, 0xae, 0xcf, 0x60, 0xe5, 0x73, 0x22, 0x82, 0xe7, 0xec,
        0x26, 0x93, 0x14, 0x3c, 0xb8, 0xa5, 0x0a, 0x2d, 0xa1, 0x98, 0xaa, 0xf5, 0xd0, 0x51, 0x52,
        0xa5, 0x96, 0x44, 0x9c, 0x67, 0x51, 0x73, 0x98, 0x05, 0x07, 0xf5, 0xed, 0x7a, 0x9a, 0x53,
        0x8a, 0x54, 0x29, 0x00, 0x81, 0xa2, 0xb0, 0xa2, 0xb0, 0x8d, 0x42, 0x77, 0x0a, 0xcb, 0xe9,
        0xe5, 0x1f, 0xd2, 0x14, 0xa4, 0xd9, 0x20, 0x6c, 0x87, 0xac, 0x5b, 0x83, 0xe8, 0xe9, 0xa1,
        0xc7, 0x66, 0x58, 0x37, 0x4e, 0x3f, 0x20, 0x6c, 0xde, 0x09, 0xe5, 0x03, 0xc5, 0x07,
    ];
    const XML: &str = "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData><row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c><c r=\"B1\"><v>1200</v></c></row><row r=\"2\"><c r=\"A2\" t=\"s\"><v>1</v></c><c r=\"B2\"><v>800</v></c></row><row r=\"3\"><c r=\"A3\" t=\"s\"><v>2</v></c><c r=\"B3\"><f>B1-B2</f><v>400</v></c></row></sheetData></worksheet>";

    #[test]
    fn each_block_kind_decodes() {
        assert_eq!(inflate(ABC_FIXED, 3, 1 << 20).unwrap(), b"abc");
        assert_eq!(inflate(ABC_STORED, 3, 1 << 20).unwrap(), b"abc");
        let rep = "Frazaro writes workbooks. ".repeat(12);
        assert_eq!(inflate(REP_FIXED, 312, 1 << 20).unwrap(), rep.as_bytes());
        assert_eq!(inflate(XML_DYNAMIC, 332, 1 << 20).unwrap(), XML.as_bytes());
        assert_eq!(XML.len(), 332);
    }

    #[test]
    fn what_is_wrong_is_named_and_nothing_panics() {
        assert_eq!(
            inflate(&ABC_FIXED[..3], 3, 1 << 20),
            Err("the data ended early")
        );
        assert_eq!(inflate(&[], 0, 1 << 20), Err("the data ended early"));
        assert_eq!(inflate(&[0x07], 0, 1 << 20), Err("an invalid block type"));
        assert_eq!(
            inflate(
                &[0x01, 0x03, 0x00, 0x00, 0x00, 0x61, 0x62, 0x63],
                3,
                1 << 20
            ),
            Err("a stored block whose length check failed")
        );
        assert_eq!(
            inflate(REP_FIXED, 312, 100),
            Err("longer than the limit allows")
        );
        assert_eq!(
            inflate(ABC_STORED, 3, 2),
            Err("longer than the limit allows")
        );
        // A back-reference to before the start: a fixed block (bits 1, 1, 0)
        // whose first symbol is length code 257 (0000001) followed by
        // distance code 0 (00000), a distance of 1 into an empty output.
        assert_eq!(
            inflate(&[0x03, 0x02], 0, 1 << 20),
            Err("a reference before the start of the output")
        );
        // Random bytes never panic.
        for seed in 0u32..64 {
            let bytes: Vec<u8> = (0..32)
                .map(|i| (seed.wrapping_mul(2654435761).wrapping_add(i * 40503) >> 13) as u8)
                .collect();
            let _ = inflate(&bytes, 0, 1 << 16);
        }
    }
}
