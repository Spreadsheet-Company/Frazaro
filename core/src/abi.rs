//! The C-ABI door (PORT.6, slice 6h): the translate API over plain C linkage,
//! for a wasm host or a native embedding, with no binding layer. Every
//! function is `#[no_mangle] extern "C"`; on wasm32 the module exports them
//! beside `memory`, and its import section stays empty
//! (tools/check_core_imports.ps1), so a page that loads it cannot phone home
//! even by mistake (SD-13).
//!
//! **Memory.** The host allocates every input with [`frazaro_alloc`], writes
//! UTF-8 into it and passes pointer and length; the core answers with a
//! record it allocated the same way, writes the record's length to `out_len`,
//! and the host frees inputs and record alike with [`frazaro_free`], each
//! with the length it was allocated with. On wasm32 a pointer is an offset
//! into the module's one linear memory.
//!
//! **The record**, little-endian: four `u32` (status, line, id length, text
//! length), then the id's bytes, then the text's bytes. Status 0: the text is
//! the output. Status 1: a refusal, the text its words, the id the
//! catalogue's, the line the program line it stands on (0 when none, as for
//! a phrasebook refused at load or a compile refusal). Status 2: an input was
//! not UTF-8, the text names which.
//!
//! **Phrasebooks** travel as one buffer, the texts separated by a NUL byte
//! (a phrasebook never holds one; the reader refuses the character), in the
//! order they load, named `vocab-1`, `vocab-2`, ... as `VLA_Browser.bas`
//! names them. The gate a door asks before loading a text it did not ship
//! ([`frazaro_vocab_gate`]) is exported too, so the page's consent checkbox
//! and the CLI's `--allow-raw` refuse with the one catalogue text.
//!
//! **The build** (PORT.7, slice 7e): [`frazaro_build_xlsx`] takes the same
//! three inputs and answers with the workbook `frazaro build` writes from
//! them. Its status-0 record is the one place the text field holds bytes
//! that are not UTF-8: the workbook itself; the id field holds the bytes'
//! SHA-256 in upper-case hex, the digest the door prints beside its byte
//! count, so a page shows the digest of exactly what it hands over. The ABI
//! number stays 1: an export was added and no signature changed its meaning.

use std::alloc::{alloc, dealloc, Layout};

use crate::api;

const STATUS_OK: u32 = 0;
const STATUS_REFUSED: u32 = 1;
const STATUS_NOT_UTF8: u32 = 2;

/// A byte buffer's layout: `len` bytes, at least one so that the allocator
/// is never asked for nothing.
fn layout(len: u32) -> Layout {
    Layout::array::<u8>((len as usize).max(1)).expect("a byte array fits in memory")
}

/// Allocate `len` bytes the host may write into (an input) or that the core
/// hands back (a record). Never null; a zero length gets one byte.
///
/// # Safety
/// The caller owns the buffer and frees it with [`frazaro_free`] and the
/// same `len`.
#[no_mangle]
pub unsafe extern "C" fn frazaro_alloc(len: u32) -> *mut u8 {
    unsafe { alloc(layout(len)) }
}

/// Free a buffer that [`frazaro_alloc`] returned, or that a translate call
/// returned, with the length it was allocated with. A null pointer is
/// ignored.
///
/// # Safety
/// `ptr` must have come from [`frazaro_alloc`] (directly, or as a returned
/// record) with this `len`, and must not be used again.
#[no_mangle]
pub unsafe extern "C" fn frazaro_free(ptr: *mut u8, len: u32) {
    if !ptr.is_null() {
        unsafe { dealloc(ptr, layout(len)) }
    }
}

/// The bytes at `ptr`; a null pointer or a zero length is the empty slice.
unsafe fn bytes<'a>(ptr: *const u8, len: u32) -> &'a [u8] {
    if ptr.is_null() || len == 0 {
        &[]
    } else {
        unsafe { std::slice::from_raw_parts(ptr, len as usize) }
    }
}

/// A record handed to the host: allocated with [`frazaro_alloc`]'s layout so
/// that [`frazaro_free`] takes it back. `text` is UTF-8 for every answer but
/// a build's, whose text field is the workbook's bytes.
fn record(status: u32, line: u32, id: &str, text: &[u8], out_len: *mut u32) -> *mut u8 {
    let total = 16 + id.len() + text.len();
    let mut buf: Vec<u8> = Vec::with_capacity(total);
    buf.extend_from_slice(&status.to_le_bytes());
    buf.extend_from_slice(&line.to_le_bytes());
    buf.extend_from_slice(&(id.len() as u32).to_le_bytes());
    buf.extend_from_slice(&(text.len() as u32).to_le_bytes());
    buf.extend_from_slice(id.as_bytes());
    buf.extend_from_slice(text);
    unsafe {
        let p = frazaro_alloc(total as u32);
        std::ptr::copy_nonoverlapping(buf.as_ptr(), p, total);
        if !out_len.is_null() {
            *out_len = total as u32;
        }
        p
    }
}

/// An input as text, or the status-2 record naming it.
fn text<'a>(b: &'a [u8], what: &str, out_len: *mut u32) -> Result<&'a str, *mut u8> {
    std::str::from_utf8(b).map_err(|_| {
        record(
            STATUS_NOT_UTF8,
            0,
            "",
            format!("{what} is not UTF-8").as_bytes(),
            out_len,
        )
    })
}

/// The three inputs a translate or a build call takes, as text and the
/// phrasebooks split at their NUL bytes, or the status-2 record naming the
/// input that is not UTF-8.
type Inputs<'a> = (&'a str, &'a str, Vec<&'a str>);

#[allow(clippy::too_many_arguments)]
unsafe fn inputs<'a>(
    program: *const u8,
    program_len: u32,
    prelude: *const u8,
    prelude_len: u32,
    books: *const u8,
    books_len: u32,
    out_len: *mut u32,
) -> Result<Inputs<'a>, *mut u8> {
    let (program, prelude, books) = unsafe {
        (
            bytes(program, program_len),
            bytes(prelude, prelude_len),
            bytes(books, books_len),
        )
    };
    let program = text(program, "the program", out_len)?;
    let prelude = text(prelude, "the prelude", out_len)?;
    let books = text(books, "a phrasebook", out_len)?;
    let vocabs: Vec<&str> = if books.is_empty() {
        Vec::new()
    } else {
        books.split('\0').collect()
    };
    Ok((program, prelude, vocabs))
}

/// The two translate functions' shared body.
#[allow(clippy::too_many_arguments)]
unsafe fn translate(
    to_vba: bool,
    program: *const u8,
    program_len: u32,
    prelude: *const u8,
    prelude_len: u32,
    books: *const u8,
    books_len: u32,
    out_len: *mut u32,
) -> *mut u8 {
    let (program, prelude, vocabs) = match unsafe {
        inputs(
            program,
            program_len,
            prelude,
            prelude_len,
            books,
            books_len,
            out_len,
        )
    } {
        Ok(t) => t,
        Err(p) => return p,
    };
    let out = if to_vba {
        api::english_translate_text_to_vba(program, prelude, &vocabs)
    } else {
        api::english_translate_text_to_vla(program, prelude, &vocabs)
    };
    match out {
        Ok(t) => record(STATUS_OK, 0, "", t.as_bytes(), out_len),
        Err(e) => record(
            STATUS_REFUSED,
            e.line,
            &e.refusal.id,
            e.refusal.text.as_bytes(),
            out_len,
        ),
    }
}

/// `EnglishTranslateTextToVla` over C linkage: the program's VLA, or its
/// refusal, as a record (see the module's doc).
///
/// # Safety
/// Every pointer is a buffer of the given length in the module's memory
/// (from [`frazaro_alloc`]), or null with length 0; `out_len` is writable or
/// null.
#[no_mangle]
pub unsafe extern "C" fn frazaro_translate_vla(
    program: *const u8,
    program_len: u32,
    prelude: *const u8,
    prelude_len: u32,
    books: *const u8,
    books_len: u32,
    out_len: *mut u32,
) -> *mut u8 {
    unsafe {
        translate(
            false,
            program,
            program_len,
            prelude,
            prelude_len,
            books,
            books_len,
            out_len,
        )
    }
}

/// `EnglishTranslateTextToVba` over C linkage: the program's VBA, or its
/// refusal, as a record.
///
/// # Safety
/// As [`frazaro_translate_vla`].
#[no_mangle]
pub unsafe extern "C" fn frazaro_translate_vba(
    program: *const u8,
    program_len: u32,
    prelude: *const u8,
    prelude_len: u32,
    books: *const u8,
    books_len: u32,
    out_len: *mut u32,
) -> *mut u8 {
    unsafe {
        translate(
            true,
            program,
            program_len,
            prelude,
            prelude_len,
            books,
            books_len,
            out_len,
        )
    }
}

/// `frazaro build` over C linkage (PORT.7, slice 7e): the program translated
/// and written as a workbook, from the same three inputs as a translate
/// call. Status 0: the text field is the workbook's bytes, not UTF-8, and
/// the id field their SHA-256 in upper-case hex, the digest the door prints.
/// Status 1: a sentence's refusal with its line, or a build's (`build-*`),
/// whose text names the line while the line field is 0, as the API gives it.
///
/// # Safety
/// As [`frazaro_translate_vla`].
#[no_mangle]
pub unsafe extern "C" fn frazaro_build_xlsx(
    program: *const u8,
    program_len: u32,
    prelude: *const u8,
    prelude_len: u32,
    books: *const u8,
    books_len: u32,
    out_len: *mut u32,
) -> *mut u8 {
    let (program, prelude, vocabs) = match unsafe {
        inputs(
            program,
            program_len,
            prelude,
            prelude_len,
            books,
            books_len,
            out_len,
        )
    } {
        Ok(t) => t,
        Err(p) => return p,
    };
    match api::english_build_xlsx(program, prelude, &vocabs) {
        Ok(bytes) => {
            let digest = crate::sha256::sha256_hex(&bytes);
            record(STATUS_OK, 0, &digest, &bytes, out_len)
        }
        Err(e) => record(
            STATUS_REFUSED,
            e.line,
            &e.refusal.id,
            e.refusal.text.as_bytes(),
            out_len,
        ),
    }
}

/// The door's gate for a phrasebook text it did not ship: F.10's capability
/// check, then SEC.2's consent for a `(raw ...)` form, given by `allow_raw`
/// (a checkbox, a flag). A status-0 record says the text may load; a
/// status-1 record carries the refusal. `name` is the name the refusal shows.
///
/// # Safety
/// As [`frazaro_translate_vla`].
#[no_mangle]
pub unsafe extern "C" fn frazaro_vocab_gate(
    text_ptr: *const u8,
    text_len: u32,
    name: *const u8,
    name_len: u32,
    allow_raw: u32,
    out_len: *mut u32,
) -> *mut u8 {
    let (t, n) = unsafe { (bytes(text_ptr, text_len), bytes(name, name_len)) };
    let t = match text(t, "the phrasebook", out_len) {
        Ok(s) => s,
        Err(p) => return p,
    };
    let n = match text(n, "the phrasebook's name", out_len) {
        Ok(s) => s,
        Err(p) => return p,
    };
    match api::vocab_gate(t, n, allow_raw != 0) {
        Ok(()) => record(STATUS_OK, 0, "", b"", out_len),
        Err(e) => record(STATUS_REFUSED, 0, &e.id, e.text.as_bytes(), out_len),
    }
}

/// The crate's version as a status-0 record, for a page's footer.
///
/// # Safety
/// `out_len` is writable or null.
#[no_mangle]
pub unsafe extern "C" fn frazaro_version_text(out_len: *mut u32) -> *mut u8 {
    record(STATUS_OK, 0, "", crate::VERSION.as_bytes(), out_len)
}

#[cfg(test)]
mod tests {
    use super::*;

    const PRELUDE: &str = include_str!("../../scripts/prelude.vla");
    const ENGLISH: &str = include_str!("../../scripts/polyglotta/english.vla");
    const FIXTURE: &str = include_str!("../../scripts/build/fixture.txt");
    const GOLDEN: &[u8] = include_bytes!("../../scripts/build/fixture_golden.xlsx");

    /// A record as the host reads it: the text field still bytes.
    struct Raw {
        status: u32,
        line: u32,
        id: String,
        bytes: Vec<u8>,
    }

    struct Rec {
        status: u32,
        line: u32,
        id: String,
        text: String,
    }

    fn u32_at(b: &[u8], i: usize) -> u32 {
        u32::from_le_bytes([b[i], b[i + 1], b[i + 2], b[i + 3]])
    }

    unsafe fn put(s: &[u8]) -> (*mut u8, u32) {
        let p = unsafe { frazaro_alloc(s.len() as u32) };
        unsafe { std::ptr::copy_nonoverlapping(s.as_ptr(), p, s.len()) };
        (p, s.len() as u32)
    }

    unsafe fn take_raw(ptr: *mut u8, len: u32) -> Raw {
        let b = unsafe { std::slice::from_raw_parts(ptr, len as usize) }.to_vec();
        unsafe { frazaro_free(ptr, len) };
        let (id_len, text_len) = (u32_at(&b, 8) as usize, u32_at(&b, 12) as usize);
        assert_eq!(
            b.len(),
            16 + id_len + text_len,
            "the record's lengths add up"
        );
        Raw {
            status: u32_at(&b, 0),
            line: u32_at(&b, 4),
            id: String::from_utf8(b[16..16 + id_len].to_vec()).unwrap(),
            bytes: b[16 + id_len..].to_vec(),
        }
    }

    unsafe fn take(ptr: *mut u8, len: u32) -> Rec {
        let r = unsafe { take_raw(ptr, len) };
        Rec {
            status: r.status,
            line: r.line,
            id: r.id,
            text: String::from_utf8(r.bytes).unwrap(),
        }
    }

    fn build(program: &[u8], books: &[&str]) -> Raw {
        unsafe {
            let (pp, pl) = put(program);
            let (qp, ql) = put(PRELUDE.as_bytes());
            let joined = books.join("\0");
            let (bp, bl) = put(joined.as_bytes());
            let mut out_len = 0u32;
            let ptr = frazaro_build_xlsx(pp, pl, qp, ql, bp, bl, &mut out_len);
            frazaro_free(pp, pl);
            frazaro_free(qp, ql);
            frazaro_free(bp, bl);
            take_raw(ptr, out_len)
        }
    }

    fn translate(to_vba: bool, program: &[u8], books: &[&str]) -> Rec {
        unsafe {
            let (pp, pl) = put(program);
            let (qp, ql) = put(PRELUDE.as_bytes());
            let joined = books.join("\0");
            let (bp, bl) = put(joined.as_bytes());
            let mut out_len = 0u32;
            let ptr = if to_vba {
                frazaro_translate_vba(pp, pl, qp, ql, bp, bl, &mut out_len)
            } else {
                frazaro_translate_vla(pp, pl, qp, ql, bp, bl, &mut out_len)
            };
            frazaro_free(pp, pl);
            frazaro_free(qp, ql);
            frazaro_free(bp, bl);
            take(ptr, out_len)
        }
    }

    fn gate(text: &str, allow_raw: u32) -> Rec {
        unsafe {
            let (tp, tl) = put(text.as_bytes());
            let (np, nl) = put(b"pasted.vla");
            let mut out_len = 0u32;
            let ptr = frazaro_vocab_gate(tp, tl, np, nl, allow_raw, &mut out_len);
            frazaro_free(tp, tl);
            frazaro_free(np, nl);
            take(ptr, out_len)
        }
    }

    #[test]
    fn a_translation_comes_back_as_a_status_0_record() {
        let r = translate(false, b"Set total to 5.", &[ENGLISH]);
        assert_eq!((r.status, r.line, r.id.as_str()), (0, 0, ""));
        assert!(r.text.contains("(set! total 5)"), "{}", r.text);
        let r = translate(true, b"Log 1.", &[ENGLISH]);
        assert_eq!(r.status, 0);
        assert!(r.text.contains("Debug.Print 1"), "{}", r.text);
    }

    #[test]
    fn a_refusal_comes_back_with_its_id_and_line() {
        let r = translate(false, b"Log 1.\nSet total to $5.", &[ENGLISH]);
        assert_eq!(
            (r.status, r.line, r.id.as_str()),
            (1, 2, "english-unknown-character")
        );
        assert!(r.text.contains("'$'"), "{}", r.text);
        // The same refusal through the VBA function.
        let v = translate(true, b"Log 1.\nSet total to $5.", &[ENGLISH]);
        assert_eq!((v.status, v.line, v.id), (r.status, r.line, r.id));
    }

    #[test]
    fn phrasebooks_are_nul_separated_and_named_in_order() {
        let r = translate(false, b"Log 1.", &[ENGLISH, "hello"]);
        assert_eq!(
            (r.status, r.line, r.id.as_str()),
            (1, 0, "english-vocab-expected-directive")
        );
        assert!(r.text.starts_with("vocab-2 line 0: "), "{}", r.text);
        // No phrasebook at all: the built-in grammar alone.
        let r = translate(false, b"Set total to 5.", &[]);
        assert_eq!(r.status, 0);
        assert!(r.text.contains("(set! total 5)"), "{}", r.text);
    }

    #[test]
    fn an_input_that_is_not_utf8_is_status_2() {
        let r = translate(false, &[0xFF, 0xFE, 0x41], &[ENGLISH]);
        assert_eq!((r.status, r.line), (2, 0));
        assert_eq!(r.text, "the program is not UTF-8");
    }

    #[test]
    fn the_gate_refuses_a_raw_form_without_consent_and_a_capability_always() {
        let raw = "(english-vla \"char cell {r:cell}\" (raw \"Debug.Print 2\"))";
        let r = gate(raw, 0);
        assert_eq!(
            (r.status, r.id.as_str()),
            (1, "english-vocab-raw-consent-declined")
        );
        assert!(r.text.contains("'pasted.vla' was not loaded"), "{}", r.text);
        let r = gate(raw, 1);
        assert_eq!((r.status, r.text.as_str()), (0, ""));
        let r = gate("(requires-capability \"network\")", 1);
        assert_eq!(
            (r.status, r.id.as_str()),
            (1, "english-vocab-requires-capability-ungranted")
        );
        let r = gate("(english-vla \"wobble {x:expr}\" (debug-print {x}))", 0);
        assert_eq!(r.status, 0);
    }

    #[test]
    fn the_version_text_and_the_allocator_s_edges() {
        unsafe {
            let mut out_len = 0u32;
            let ptr = frazaro_version_text(&mut out_len);
            let r = take(ptr, out_len);
            assert_eq!((r.status, r.text.as_str()), (0, crate::VERSION));
            let p = frazaro_alloc(0);
            assert!(!p.is_null());
            frazaro_free(p, 0);
            frazaro_free(std::ptr::null_mut(), 5);
        }
    }

    #[test]
    fn a_build_comes_back_as_the_golden_s_bytes_with_their_digest_as_its_id() {
        // The page's Download hands over exactly what frazaro build writes:
        // the fixture through the record is the golden, byte for byte, and
        // the id is the digest the door prints beside its byte count.
        let r = build(FIXTURE.as_bytes(), &[ENGLISH]);
        assert_eq!((r.status, r.line), (0, 0));
        assert_eq!(r.id, crate::sha256::sha256_hex(GOLDEN));
        assert_eq!(r.bytes.len(), GOLDEN.len());
        assert!(r.bytes == GOLDEN, "the record's bytes are the golden");
        // The rows of a page carry no final newline; the file does. Same bytes.
        let r = build(FIXTURE.trim_end().as_bytes(), &[ENGLISH]);
        assert!(
            r.bytes == GOLDEN,
            "a program without a final newline builds the same"
        );
    }

    #[test]
    fn a_build_s_refusals_come_back_through_the_record() {
        // A sentence the writer cannot hold: the build's refusal, its text
        // naming the line, the line field 0 as the API gives it.
        let r = build(b"Put 5 into cell B2.\nSay \"hello\".\n", &[ENGLISH]);
        assert_eq!(
            (r.status, r.line, r.id.as_str()),
            (1, 0, "build-not-representable")
        );
        let text = String::from_utf8(r.bytes).unwrap();
        assert!(
            text.contains("Line 2 asks for more: Say \"hello\"."),
            "{text}"
        );
        // A sentence the language refuses: the translation's refusal, with its line.
        let r = build(b"Put 5 into cell B2.\nSet total to $5.\n", &[ENGLISH]);
        assert_eq!(
            (r.status, r.line, r.id.as_str()),
            (1, 2, "english-unknown-character")
        );
        // An input that is not UTF-8.
        let r = build(&[0xFF, 0xFE, 0x41], &[ENGLISH]);
        assert_eq!((r.status, r.line), (2, 0));
        assert_eq!(
            String::from_utf8(r.bytes).unwrap(),
            "the program is not UTF-8"
        );
    }
}
