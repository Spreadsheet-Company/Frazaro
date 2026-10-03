//! A zip container with stored entries and no clock (PORT.7, slice 7a).
//!
//! New ground, said plainly: the reference saves through Excel and never
//! writes a zip. Every entry here is STORED (method 0), so the bytes of a
//! part are the bytes in the file and the only arithmetic the format asks for
//! is its CRC-32. Every entry carries the zip epoch, 1980-01-01 00:00:00, as
//! its modification stamp, the one value that reads as "no clock"; entries go
//! in the order given; no extra field, no comment, no zip64. Two calls over
//! the same parts give the same bytes, which is what the build golden
//! (`scripts/build/*_golden.xlsx`) and `frazaro rebuild` stand on, and what
//! `tools/check_build_golden.ps1` reads off the artifact.
//!
//! Sizes are the format's 32-bit ones. The writer answers `None` when a part,
//! an offset or the count of parts would not fit, and the build turns that
//! into the catalogue's refusal; nothing here panics. The reader at the foot
//! walks an archive's central directory back, for the tests and for
//! `rebuild`, which reads only archives this writer made.

const LOCAL_HEADER: u32 = 0x0403_4B50;
const CENTRAL_HEADER: u32 = 0x0201_4B50;
const END_OF_CENTRAL_DIRECTORY: u32 = 0x0605_4B50;
/// "Version needed" and "version made by": 2.0, plain MS-DOS attributes.
const VERSION: u16 = 20;
/// The compression method of every entry: none.
pub const STORED: u16 = 0;
/// The modification stamp of every entry, in MS-DOS form: 00:00:00 on
/// 1980-01-01 (day 1 of month 1 of year 1980, the format's year zero).
pub const EPOCH_TIME: u16 = 0;
pub const EPOCH_DATE: u16 = 0x0021;

const fn crc_table() -> [u32; 256] {
    let mut table = [0u32; 256];
    let mut n = 0;
    while n < 256 {
        let mut c = n as u32;
        let mut k = 0;
        while k < 8 {
            c = if c & 1 != 0 {
                0xEDB8_8320 ^ (c >> 1)
            } else {
                c >> 1
            };
            k += 1;
        }
        table[n] = c;
        n += 1;
    }
    table
}
const CRC_TABLE: [u32; 256] = crc_table();

/// The IEEE CRC-32 (reflected polynomial EDB88320) as the zip format uses
/// it; the check value for "123456789" is CBF43926.
pub fn crc32(data: &[u8]) -> u32 {
    let mut c = 0xFFFF_FFFFu32;
    for &b in data {
        c = CRC_TABLE[((c ^ u32::from(b)) & 0xFF) as usize] ^ (c >> 8);
    }
    c ^ 0xFFFF_FFFF
}

fn put_u16(out: &mut Vec<u8>, v: u16) {
    out.extend_from_slice(&v.to_le_bytes());
}

fn put_u32(out: &mut Vec<u8>, v: u32) {
    out.extend_from_slice(&v.to_le_bytes());
}

/// An entry written as given (slice 7d): `bytes` go into the file under
/// `method`, and `crc` and `size` describe the uncompressed content. A
/// stored entry's bytes are its content; an entry copied from another
/// archive keeps that archive's compressed bytes, method, CRC and size, so
/// the copy is byte-exact without decoding.
pub struct RawEntry<'a> {
    pub name: &'a str,
    pub method: u16,
    pub crc: u32,
    pub size: u32,
    pub bytes: &'a [u8],
}

/// The archive holding `parts`, each stored, in the order given; `None`
/// when a size would not fit the format's fields.
pub fn write_stored(parts: &[(String, Vec<u8>)]) -> Option<Vec<u8>> {
    let mut raws = Vec::with_capacity(parts.len());
    for (name, data) in parts {
        raws.push(RawEntry {
            name,
            method: STORED,
            crc: crc32(data),
            size: u32::try_from(data.len()).ok()?,
            bytes: data,
        });
    }
    write_entries(&raws)
}

/// The archive holding `entries` as given, in the order given, every one
/// stamped with the epoch; `None` when a size would not fit the format's
/// fields.
pub fn write_entries(entries: &[RawEntry]) -> Option<Vec<u8>> {
    let count = u16::try_from(entries.len()).ok()?;
    let mut out: Vec<u8> = Vec::new();
    let mut central: Vec<u8> = Vec::new();
    for e in entries {
        let name_len = u16::try_from(e.name.len()).ok()?;
        let compressed = u32::try_from(e.bytes.len()).ok()?;
        let offset = u32::try_from(out.len()).ok()?;
        // The local file header, then the entry's bytes as they are.
        put_u32(&mut out, LOCAL_HEADER);
        put_u16(&mut out, VERSION);
        put_u16(&mut out, 0); // general purpose flags
        put_u16(&mut out, e.method);
        put_u16(&mut out, EPOCH_TIME);
        put_u16(&mut out, EPOCH_DATE);
        put_u32(&mut out, e.crc);
        put_u32(&mut out, compressed);
        put_u32(&mut out, e.size);
        put_u16(&mut out, name_len);
        put_u16(&mut out, 0); // extra field length
        out.extend_from_slice(e.name.as_bytes());
        out.extend_from_slice(e.bytes);
        // The central directory's record of it.
        put_u32(&mut central, CENTRAL_HEADER);
        put_u16(&mut central, VERSION); // version made by
        put_u16(&mut central, VERSION); // version needed
        put_u16(&mut central, 0);
        put_u16(&mut central, e.method);
        put_u16(&mut central, EPOCH_TIME);
        put_u16(&mut central, EPOCH_DATE);
        put_u32(&mut central, e.crc);
        put_u32(&mut central, compressed);
        put_u32(&mut central, e.size);
        put_u16(&mut central, name_len);
        put_u16(&mut central, 0); // extra field length
        put_u16(&mut central, 0); // comment length
        put_u16(&mut central, 0); // disk number start
        put_u16(&mut central, 0); // internal attributes
        put_u32(&mut central, 0); // external attributes
        put_u32(&mut central, offset);
        central.extend_from_slice(e.name.as_bytes());
    }
    let cd_offset = u32::try_from(out.len()).ok()?;
    let cd_size = u32::try_from(central.len()).ok()?;
    out.extend_from_slice(&central);
    put_u32(&mut out, END_OF_CENTRAL_DIRECTORY);
    put_u16(&mut out, 0); // this disk
    put_u16(&mut out, 0); // the disk the directory starts on
    put_u16(&mut out, count);
    put_u16(&mut out, count);
    put_u32(&mut out, cd_size);
    put_u32(&mut out, cd_offset);
    put_u16(&mut out, 0); // comment length
    Some(out)
}

/// One entry of an archive's central directory, with where its data starts.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Entry {
    pub name: String,
    pub method: u16,
    pub time: u16,
    pub date: u16,
    pub crc: u32,
    pub compressed_size: u32,
    pub size: u32,
    /// The offset of the entry's bytes, behind its local header.
    pub data_offset: usize,
}

fn read_u16(b: &[u8], at: usize) -> Option<u16> {
    Some(u16::from_le_bytes([*b.get(at)?, *b.get(at + 1)?]))
}

fn read_u32(b: &[u8], at: usize) -> Option<u32> {
    Some(u32::from_le_bytes([
        *b.get(at)?,
        *b.get(at + 1)?,
        *b.get(at + 2)?,
        *b.get(at + 3)?,
    ]))
}

const ZIP64_LOCATOR: u32 = 0x0706_4B50;

/// Where the end-of-central-directory record starts: the last 22 bytes
/// when the archive has no comment, and up to 65,535 bytes earlier when it
/// has one (the record's comment length must then reach the end).
fn find_eocd(bytes: &[u8]) -> Option<usize> {
    if bytes.len() < 22 {
        return None;
    }
    let last = bytes.len() - 22;
    let first = last.saturating_sub(65_535);
    (first..=last).rev().find(|&i| {
        read_u32(bytes, i) == Some(END_OF_CENTRAL_DIRECTORY)
            && read_u16(bytes, i + 20).map(usize::from) == Some(bytes.len() - 22 - i)
    })
}

/// Why `entries` would answer `None` for these bytes, in a few words, or
/// `None` when it would not.
pub fn why_not_an_archive(bytes: &[u8]) -> Option<&'static str> {
    if bytes.starts_with(&[0xD0, 0xCF, 0x11, 0xE0]) {
        return Some("it is an older .xls or an encrypted workbook, not an .xlsx package");
    }
    let Some(eocd) = find_eocd(bytes) else {
        return Some("it has no end-of-central-directory record, so it is not a zip archive");
    };
    if read_u16(bytes, eocd + 10) == Some(0xFFFF)
        || read_u32(bytes, eocd + 16) == Some(0xFFFF_FFFF)
        || (eocd >= 20 && read_u32(bytes, eocd - 20) == Some(ZIP64_LOCATOR))
    {
        return Some("it is a zip64 archive, which this version does not read");
    }
    if entries(bytes).is_none() {
        return Some("its central directory does not read whole");
    }
    None
}

/// The central directory of an archive, in directory order; `None` when
/// the bytes are not an archive this reader knows (zip64 included). A
/// comment at the end is allowed; a data descriptor after an entry is
/// harmless, since the sizes come from the directory. Each entry's data
/// offset is read from its own local header, as a reader must.
pub fn entries(bytes: &[u8]) -> Option<Vec<Entry>> {
    let eocd = find_eocd(bytes)?;
    let count = usize::from(read_u16(bytes, eocd + 10)?);
    let cd_size = read_u32(bytes, eocd + 12)? as usize;
    let cd_offset = read_u32(bytes, eocd + 16)? as usize;
    if count == 0xFFFF || cd_offset == 0xFFFF_FFFF {
        return None; // zip64
    }
    if cd_offset.checked_add(cd_size)? > eocd {
        return None;
    }
    let mut pos = cd_offset;
    let mut list = Vec::with_capacity(count);
    for _ in 0..count {
        if read_u32(bytes, pos)? != CENTRAL_HEADER {
            return None;
        }
        let method = read_u16(bytes, pos + 10)?;
        let time = read_u16(bytes, pos + 12)?;
        let date = read_u16(bytes, pos + 14)?;
        let crc = read_u32(bytes, pos + 16)?;
        let compressed_size = read_u32(bytes, pos + 20)?;
        let size = read_u32(bytes, pos + 24)?;
        let name_len = usize::from(read_u16(bytes, pos + 28)?);
        let extra_len = usize::from(read_u16(bytes, pos + 30)?);
        let comment_len = usize::from(read_u16(bytes, pos + 32)?);
        let local = read_u32(bytes, pos + 42)? as usize;
        let name = std::str::from_utf8(bytes.get(pos + 46..pos + 46 + name_len)?)
            .ok()?
            .to_string();
        if read_u32(bytes, local)? != LOCAL_HEADER {
            return None;
        }
        let local_name_len = usize::from(read_u16(bytes, local + 26)?);
        let local_extra_len = usize::from(read_u16(bytes, local + 28)?);
        let data_offset = local + 30 + local_name_len + local_extra_len;
        if data_offset.checked_add(compressed_size as usize)? > bytes.len() {
            return None;
        }
        list.push(Entry {
            name,
            method,
            time,
            date,
            crc,
            compressed_size,
            size,
            data_offset,
        });
        pos += 46 + name_len + extra_len + comment_len;
    }
    Some(list)
}

/// A stored entry's bytes; `None` for an entry that is not stored.
pub fn stored_data<'a>(bytes: &'a [u8], e: &Entry) -> Option<&'a [u8]> {
    if e.method != STORED {
        return None;
    }
    bytes.get(e.data_offset..e.data_offset + e.size as usize)
}

/// An entry's bytes as the archive holds them, compressed or not, for a
/// byte-exact copy into another archive.
pub fn raw_data<'a>(bytes: &'a [u8], e: &Entry) -> Option<&'a [u8]> {
    bytes.get(e.data_offset..e.data_offset + e.compressed_size as usize)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn the_crc_check_value() {
        assert_eq!(crc32(b"123456789"), 0xCBF4_3926);
        assert_eq!(crc32(b""), 0);
    }

    fn two_parts() -> Vec<(String, Vec<u8>)> {
        vec![
            ("[Content_Types].xml".to_string(), b"<Types/>".to_vec()),
            ("xl/workbook.xml".to_string(), b"<workbook/>".to_vec()),
        ]
    }

    #[test]
    fn the_archive_walks_back_through_its_own_directory() {
        let bytes = write_stored(&two_parts()).expect("two small parts fit");
        assert_eq!(&bytes[..4], &[0x50, 0x4B, 0x03, 0x04]);
        let list = entries(&bytes).expect("a central directory");
        assert_eq!(list.len(), 2);
        assert_eq!(list[0].name, "[Content_Types].xml");
        assert_eq!(list[1].name, "xl/workbook.xml");
        for (e, (_, data)) in list.iter().zip(two_parts().iter()) {
            assert_eq!(e.method, STORED);
            assert_eq!((e.time, e.date), (EPOCH_TIME, EPOCH_DATE));
            assert_eq!(e.size, data.len() as u32);
            assert_eq!(e.compressed_size, data.len() as u32);
            assert_eq!(e.crc, crc32(data));
            assert_eq!(stored_data(&bytes, e).unwrap(), data.as_slice());
        }
    }

    #[test]
    fn the_same_parts_give_the_same_bytes() {
        assert_eq!(write_stored(&two_parts()), write_stored(&two_parts()));
    }

    #[test]
    fn an_entry_copies_as_it_is_and_a_comment_or_zip64_is_told_apart() {
        // A "deflated" entry carried as given: the bytes are not decoded,
        // the method, CRC and size are what the source said.
        let raw = RawEntry {
            name: "xl/styles.xml",
            method: 8,
            crc: 0x1234_5678,
            size: 999,
            bytes: &[0x4b, 0x4c, 0x4a, 0x06, 0x00],
        };
        let own = RawEntry {
            name: "own.xml",
            method: STORED,
            crc: crc32(b"<a/>"),
            size: 4,
            bytes: b"<a/>",
        };
        let bytes = write_entries(&[raw, own]).unwrap();
        let list = entries(&bytes).unwrap();
        assert_eq!(list[0].method, 8);
        assert_eq!(list[0].crc, 0x1234_5678);
        assert_eq!(list[0].size, 999);
        assert_eq!(list[0].compressed_size, 5);
        assert_eq!(
            raw_data(&bytes, &list[0]).unwrap(),
            &[0x4b, 0x4c, 0x4a, 0x06, 0x00]
        );
        assert_eq!(stored_data(&bytes, &list[0]), None);
        assert_eq!(stored_data(&bytes, &list[1]).unwrap(), b"<a/>");
        assert_eq!(why_not_an_archive(&bytes), None);
        // An archive comment after the record is allowed.
        let mut commented = bytes.clone();
        let n = commented.len();
        commented[n - 2] = 5;
        commented.extend_from_slice(b"hello");
        assert_eq!(entries(&commented).unwrap().len(), 2);
        // What is not an archive is named.
        assert_eq!(
            why_not_an_archive(b"not a zip"),
            Some("it has no end-of-central-directory record, so it is not a zip archive")
        );
        assert_eq!(
            why_not_an_archive(&[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]),
            Some("it is an older .xls or an encrypted workbook, not an .xlsx package")
        );
        let mut zip64 = bytes.clone();
        let n = zip64.len();
        zip64[n - 12] = 0xFF;
        zip64[n - 11] = 0xFF;
        zip64[n - 10] = 0xFF;
        zip64[n - 9] = 0xFF;
        assert_eq!(
            why_not_an_archive(&zip64),
            Some("it is a zip64 archive, which this version does not read")
        );
        assert!(entries(&zip64).is_none());
    }

    #[test]
    fn what_does_not_fit_is_refused_not_truncated() {
        let many: Vec<(String, Vec<u8>)> =
            (0..65_536).map(|_| ("a".to_string(), Vec::new())).collect();
        assert!(write_stored(&many).is_none());
        assert!(entries(b"not a zip at all, not even twenty-two").is_none());
        assert!(entries(b"").is_none());
    }
}
