//! The language's C surface (KERNEL.22, 2026-10-09): the machine's four
//! calls over plain C linkage, for a wasm host or a native embedding that
//! wants the bare grid without an engine's devices (`Alonzo/SPEC.md`
//! section 4.6: `vla_load`, `vla_write`, `vla_step` and `vla_view`, with
//! the records an engine's calls answer).
//!
//! **Exported only when asked for.** Every function here is compiled and
//! tested always, and carries `#[no_mangle]` only under the `c-abi`
//! feature: a C export of a library crate is exported again by every
//! module built on it (measured 2026-10-09 on rustc 1.99 for
//! wasm32-unknown-unknown), and neither `frazaro-core`'s module nor an
//! engine's should hold a second door to the grid past its own. So
//! `cargo build --release -p vla-lang --features c-abi --target
//! wasm32-unknown-unknown` writes the language's module with these nine
//! exports and `memory`, which `tools/check_wasm_exports.ps1` holds exactly,
//! and the crates above call [`crate::machine`] in Rust instead.
//!
//! **Memory and the record**, as `frazaro-core`'s C surface has them: the
//! host allocates every input with [`vla_alloc`] and frees inputs and
//! records alike with [`vla_free`] and the length each had; every answer is
//! one record, four little-endian `u32` (status, line, id length, text
//! length), then the id's bytes and the text's bytes, its total length
//! written to `out_len`. Status 0: the text is the answer. Status 1: a
//! refusal, the id the catalogue's, the text its words, the line the row it
//! stands on or 0. Status 2: an input that is not UTF-8, named. The plane's
//! answer is the one record whose text is bytes and not UTF-8.
//!
//! **The handles** live in one table for the module's life, from 1, never
//! reused, at most 16 at once (`machine::Handles`). A load uses the
//! language's own library, the day-one functions (`AD-1`).

use std::alloc::{alloc, dealloc, Layout};
use std::cell::RefCell;

use crate::calc::library;
use crate::machine::{unloaded_row, written_row, Handles, Machine, Viewed};
use crate::messages::Refusal;

const STATUS_OK: u32 = 0;
const STATUS_REFUSED: u32 = 1;
const STATUS_NOT_UTF8: u32 = 2;

/// The surface's version: it stays while an export is added and no
/// signature changes its meaning, as `frazaro_abi_version` does.
pub const ABI_VERSION: u32 = 1;

thread_local! {
    // On wasm32-unknown-unknown a thread-local is a plain static, and the
    // module's import section stays empty.
    static GRIDS: RefCell<Handles<Machine>> = const { RefCell::new(Handles::new()) };
}

fn layout(len: u32) -> Layout {
    Layout::array::<u8>((len as usize).max(1)).expect("a byte array fits in memory")
}

/// The surface's version, a bare number.
#[cfg_attr(feature = "c-abi", no_mangle)]
pub extern "C" fn vla_abi_version() -> u32 {
    ABI_VERSION
}

/// Allocate `len` bytes the host may write into (an input) or that the
/// module hands back (a record). A zero length gets one byte.
///
/// # Safety
/// The caller owns the buffer and frees it with [`vla_free`] and the same
/// `len`.
#[cfg_attr(feature = "c-abi", no_mangle)]
pub unsafe extern "C" fn vla_alloc(len: u32) -> *mut u8 {
    unsafe { alloc(layout(len)) }
}

/// Free a buffer [`vla_alloc`] returned, or a record a call returned, with
/// the length it was allocated with. A null pointer is ignored.
///
/// # Safety
/// `ptr` came from [`vla_alloc`], directly or as a record, with this `len`,
/// and is not used again.
#[cfg_attr(feature = "c-abi", no_mangle)]
pub unsafe extern "C" fn vla_free(ptr: *mut u8, len: u32) {
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

/// A record handed to the host, allocated with [`vla_alloc`]'s layout.
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
        let p = vla_alloc(total as u32);
        if p.is_null() {
            return p;
        }
        std::ptr::copy_nonoverlapping(buf.as_ptr(), p, total);
        if !out_len.is_null() {
            *out_len = total as u32;
        }
        p
    }
}

fn answer(text: &[u8], out_len: *mut u32) -> *mut u8 {
    record(STATUS_OK, 0, "", text, out_len)
}

fn refused(line: u32, r: &Refusal, out_len: *mut u32) -> *mut u8 {
    record(STATUS_REFUSED, line, &r.id, r.text.as_bytes(), out_len)
}

/// An input as text, or the status-2 record naming it.
fn text<'a>(b: &'a [u8], what: &str, out_len: *mut u32) -> Result<&'a str, *mut u8> {
    std::str::from_utf8(b).map_err(|_| {
        record(
            STATUS_NOT_UTF8,
            0,
            "",
            format!("{what} not UTF-8").as_bytes(),
            out_len,
        )
    })
}

/// The crate's version as a status-0 record.
///
/// # Safety
/// `out_len` is writable or null.
#[cfg_attr(feature = "c-abi", no_mangle)]
pub unsafe extern "C" fn vla_version_text(out_len: *mut u32) -> *mut u8 {
    answer(crate::VERSION.as_bytes(), out_len)
}

/// `load`: a text of rows (`crate::machine::load`) made into a grid under a
/// new handle, answered `(loaded <handle> <sheets> <cells> <formulas>)`;
/// refused with the row's line, or `grid-handle-limit` when 16 grids are
/// live.
///
/// # Safety
/// `rows` came from [`vla_alloc`] with `rows_len`, or is null with 0;
/// `out_len` is writable or null.
#[cfg_attr(feature = "c-abi", no_mangle)]
pub unsafe extern "C" fn vla_load(rows: *const u8, rows_len: u32, out_len: *mut u32) -> *mut u8 {
    let text = match text(unsafe { bytes(rows, rows_len) }, "the rows are", out_len) {
        Ok(t) => t,
        Err(p) => return p,
    };
    if let Err(r) = GRIDS.with(|g| g.borrow().room()) {
        return refused(0, &r, out_len);
    }
    let machine = match Machine::load(text, library::language()) {
        Ok(m) => m,
        Err(e) => return refused(e.line, &e.refusal, out_len),
    };
    let loaded = machine.loaded();
    GRIDS.with(|g| match g.borrow_mut().insert(machine) {
        Ok(handle) => answer(loaded.row(handle).as_bytes(), out_len),
        Err(r) => refused(0, &r, out_len),
    })
}

/// `write`: rows, one a line, applied all or none, answered `(written <n>)`.
///
/// # Safety
/// As [`vla_load`].
#[cfg_attr(feature = "c-abi", no_mangle)]
pub unsafe extern "C" fn vla_write(
    handle: u32,
    rows: *const u8,
    rows_len: u32,
    out_len: *mut u32,
) -> *mut u8 {
    GRIDS.with(|g| {
        let mut g = g.borrow_mut();
        let machine = match g.get_mut(handle) {
            Ok(m) => m,
            Err(r) => return refused(0, &r, out_len),
        };
        let text = match text(unsafe { bytes(rows, rows_len) }, "the rows are", out_len) {
            Ok(t) => t,
            Err(p) => return p,
        };
        match machine.write(text) {
            Ok(n) => answer(written_row(n).as_bytes(), out_len),
            Err(e) => refused(e.line, &e.refusal, out_len),
        }
    })
}

/// `step`: at most `budget` cells, 0 for no limit, answered `(step <frame>
/// <evaluated> <of> done|yielded)`.
///
/// # Safety
/// `out_len` is writable or null.
#[cfg_attr(feature = "c-abi", no_mangle)]
pub unsafe extern "C" fn vla_step(handle: u32, budget: u32, out_len: *mut u32) -> *mut u8 {
    GRIDS.with(|g| {
        let mut g = g.borrow_mut();
        match g.get_mut(handle) {
            Ok(m) => answer(m.step(budget as usize).row().as_bytes(), out_len),
            Err(r) => refused(0, &r, out_len),
        }
    })
}

/// `view`: a projection's name (`grid` or `plane`), a sheet's name and a
/// window (`A1:F20`; empty for the sheet's extent), answered with the
/// record's lines or the plane's bytes.
///
/// # Safety
/// Each pointer came from [`vla_alloc`] with its length, or is null with 0;
/// `out_len` is writable or null.
#[allow(clippy::too_many_arguments)]
#[cfg_attr(feature = "c-abi", no_mangle)]
pub unsafe extern "C" fn vla_view(
    handle: u32,
    projection: *const u8,
    projection_len: u32,
    sheet: *const u8,
    sheet_len: u32,
    window: *const u8,
    window_len: u32,
    out_len: *mut u32,
) -> *mut u8 {
    GRIDS.with(|g| {
        let g = g.borrow();
        let machine = match g.get(handle) {
            Ok(m) => m,
            Err(r) => return refused(0, &r, out_len),
        };
        let projection = match text(
            unsafe { bytes(projection, projection_len) },
            "the projection is",
            out_len,
        ) {
            Ok(t) => t,
            Err(p) => return p,
        };
        let sheet = match text(
            unsafe { bytes(sheet, sheet_len) },
            "the sheet's name is",
            out_len,
        ) {
            Ok(t) => t,
            Err(p) => return p,
        };
        let window = match text(
            unsafe { bytes(window, window_len) },
            "the window is",
            out_len,
        ) {
            Ok(t) => t,
            Err(p) => return p,
        };
        let window = (!window.is_empty()).then_some(window);
        match machine.view(projection, sheet, window) {
            Ok(Viewed::Text(t)) => answer(t.as_bytes(), out_len),
            Ok(Viewed::Bytes(b)) => answer(&b, out_len),
            Err(r) => refused(0, &r, out_len),
        }
    })
}

/// `unload`: the grid freed and its handle never given again, answered
/// `(unloaded <handle>)`.
///
/// # Safety
/// `out_len` is writable or null.
#[cfg_attr(feature = "c-abi", no_mangle)]
pub unsafe extern "C" fn vla_unload(handle: u32, out_len: *mut u32) -> *mut u8 {
    GRIDS.with(|g| match g.borrow_mut().remove(handle) {
        Ok(_) => answer(unloaded_row(handle).as_bytes(), out_len),
        Err(r) => refused(0, &r, out_len),
    })
}

#[cfg(test)]
mod tests {
    //! The surface as a host meets it: inputs allocated with the module's
    //! allocator, every answer read as the record's bytes and freed with the
    //! length it holds. One test, since the handle table is the thread's for
    //! the module's life and a second test on the same thread would see it.
    use super::*;

    /// A record as the host reads it, the text still bytes.
    struct Rec {
        status: u32,
        line: u32,
        id: String,
        text: Vec<u8>,
    }

    impl Rec {
        fn text(&self) -> &str {
            std::str::from_utf8(&self.text).unwrap()
        }
    }

    /// An input in the module's memory: its pointer and length.
    fn put(bytes: &[u8]) -> (*mut u8, u32) {
        let len = bytes.len() as u32;
        let p = unsafe { vla_alloc(len) };
        unsafe { std::ptr::copy_nonoverlapping(bytes.as_ptr(), p, bytes.len()) };
        (p, len)
    }

    fn take(p: *mut u8, out_len: u32) -> Rec {
        let b = unsafe { std::slice::from_raw_parts(p, out_len as usize) }.to_vec();
        unsafe { vla_free(p, out_len) };
        let word = |i: usize| u32::from_le_bytes([b[i], b[i + 1], b[i + 2], b[i + 3]]);
        let (status, line, id_len, text_len) = (word(0), word(4), word(8) as usize, word(12));
        assert_eq!(
            16 + id_len + text_len as usize,
            b.len(),
            "the record's length"
        );
        Rec {
            status,
            line,
            id: String::from_utf8(b[16..16 + id_len].to_vec()).unwrap(),
            text: b[16 + id_len..].to_vec(),
        }
    }

    fn load(rows: &[u8]) -> Rec {
        let (p, n) = put(rows);
        let mut out = 0u32;
        let r = unsafe { vla_load(p, n, &mut out) };
        unsafe { vla_free(p, n) };
        take(r, out)
    }

    fn write(h: u32, rows: &str) -> Rec {
        let (p, n) = put(rows.as_bytes());
        let mut out = 0u32;
        let r = unsafe { vla_write(h, p, n, &mut out) };
        unsafe { vla_free(p, n) };
        take(r, out)
    }

    fn step(h: u32, budget: u32) -> Rec {
        let mut out = 0u32;
        take(unsafe { vla_step(h, budget, &mut out) }, out)
    }

    fn view(h: u32, projection: &[u8], sheet: &str, window: &str) -> Rec {
        let (pp, pn) = put(projection);
        let (sp, sn) = put(sheet.as_bytes());
        let (wp, wn) = put(window.as_bytes());
        let mut out = 0u32;
        let r = unsafe { vla_view(h, pp, pn, sp, sn, wp, wn, &mut out) };
        unsafe {
            vla_free(pp, pn);
            vla_free(sp, sn);
            vla_free(wp, wn);
        }
        take(r, out)
    }

    fn unload(h: u32) -> Rec {
        let mut out = 0u32;
        take(unsafe { vla_unload(h, &mut out) }, out)
    }

    #[test]
    fn the_surface_answers_the_four_calls_in_records() {
        assert_eq!(vla_abi_version(), 1);
        let mut out = 0u32;
        let v = take(unsafe { vla_version_text(&mut out) }, out);
        assert_eq!((v.status, v.text()), (0, crate::VERSION));

        let rows = b"(cell \"S\" \"A1\" 2)\n(formula \"S\" \"B1\" \"=S.last!B1+A1\")\n";
        let loaded = load(rows);
        assert_eq!((loaded.status, loaded.line, loaded.id.as_str()), (0, 0, ""));
        let h: u32 = loaded
            .text()
            .strip_prefix("(loaded ")
            .and_then(|t| t.split(' ').next())
            .and_then(|t| t.parse().ok())
            .unwrap();
        assert_eq!(loaded.text(), format!("(loaded {h} 1 2 1)"));

        let s = step(h, 0);
        assert_eq!((s.status, s.text()), (0, "(step 1 1 1 done)"));
        assert_eq!(step(h, 0).text(), "(step 2 1 1 done)");
        let g = view(h, b"grid", "S", "");
        assert_eq!(g.status, 0);
        assert!(
            g.text().contains("(value \"S\" \"B1\" 4)\n"),
            "{}",
            g.text()
        );
        let p = view(h, b"plane", "S", "A1:C1");
        assert_eq!((p.status, p.text.as_slice()), (0, &[2u8, 4, 0][..]));

        assert_eq!(write(h, "(cell \"S\" \"A1\" 3)").text(), "(written 1)");
        let w = write(h, "(cell \"S\" \"A1\" 4)\n(cell \"S.last\" \"A1\" 1)");
        assert_eq!((w.status, w.line, w.id.as_str()), (1, 2, "grid-write-last"));
        let w = write(h, "(derived \"S\" \"B1\" 1)");
        assert_eq!(
            (w.status, w.line, w.id.as_str()),
            (1, 1, "grid-write-derived")
        );
        assert!(w.text().contains("S!B1"), "{}", w.text());

        // A frame in progress refuses a write, and a write that is not UTF-8
        // is named.
        assert_eq!(
            write(h, "(formula \"S\" \"C1:C9\" \"=A1\")").text(),
            "(written 9)"
        );
        assert_eq!(step(h, 4).text(), "(step 3 4 10 yielded)");
        let w = write(h, "(cell \"S\" \"A1\" 5)");
        assert_eq!((w.status, w.id.as_str()), (1, "grid-write-during-step"));
        assert_eq!(step(h, 0).text(), "(step 3 10 10 done)");
        let (p8, n8) = put(&[0xff, 0xfe]);
        let mut out = 0u32;
        let bad = take(unsafe { vla_write(h, p8, n8, &mut out) }, out);
        unsafe { vla_free(p8, n8) };
        assert_eq!((bad.status, bad.text()), (2, "the rows are not UTF-8"));
        let bad = view(h, &[0xff], "S", "");
        assert_eq!((bad.status, bad.text()), (2, "the projection is not UTF-8"));
        let bad = view(h, b"raster", "S", "");
        assert_eq!(
            (bad.status, bad.id.as_str()),
            (1, "view-projection-unknown")
        );

        // A refused load holds the row's line; an unknown handle is named.
        let r = load(b"(cell \"S\" \"A1\" 1)\n(cell \"S\" \"A1\" 2)\n");
        assert_eq!(
            (r.status, r.line, r.id.as_str()),
            (1, 2, "grid-cell-written-twice")
        );
        assert_eq!(unload(h).text(), format!("(unloaded {h})"));
        for gone in [h, 0] {
            let r = step(gone, 0);
            assert_eq!((r.status, r.id.as_str()), (1, "grid-handle-unknown"));
            assert!(
                r.text().contains(&format!("handle {gone};")),
                "{}",
                r.text()
            );
            assert_eq!(unload(gone).id, "grid-handle-unknown");
        }

        // Sixteen live at once; the next is refused; a handle is never given
        // again.
        let mut live = Vec::new();
        for _ in 0..16 {
            let r = load(b"(cell \"S\" \"A1\" 1)\n");
            assert_eq!(r.status, 0, "{}", r.text());
            live.push(r);
        }
        let r = load(b"(cell \"S\" \"A1\" 1)\n");
        assert_eq!((r.status, r.id.as_str()), (1, "grid-handle-limit"));
        let handles: Vec<u32> = live
            .iter()
            .map(|r| r.text().split(' ').nth(1).unwrap().parse().unwrap())
            .collect();
        assert_eq!(handles, (h + 1..=h + 16).collect::<Vec<u32>>());
        for handle in handles {
            assert_eq!(unload(handle).status, 0);
        }
        let again = load(b"(cell \"S\" \"A1\" 1)\n");
        assert_eq!(again.text(), format!("(loaded {} 1 1 0)", h + 17));
        assert_eq!(unload(h + 17).status, 0);
    }
}
