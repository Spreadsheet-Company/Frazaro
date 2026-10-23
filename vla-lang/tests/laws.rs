//! The language's laws, held over generated programs (VLA.1). Each law was
//! pinned before on a few fixed inputs (the corpus, five view goldens, one
//! grid); here a seeded generator writes the inputs, the same ones on every
//! run, so that each law is held over programs no one wrote by hand, through
//! the public API an outside user has. A failing case names its seed and
//! prints its input.
//!
//! 1. Reading what the printer wrote gives the same forms, and printing
//!    them again gives the same text.
//! 2. Expansion reaches its fixpoint in one pass: expanding what it wrote
//!    fires no macro and changes nothing.
//! 3. Loading a grid's view and viewing it again gives the same view,
//!    before a frame and after each one, over grids whose formats hold
//!    codes among them, whose `text` rows the load skips (KERNEL.47), and
//!    what they say of a font, an alignment and a border (KERNEL.50).
//! 4. A frame stepped in slices of any budget is the frame stepped whole.
//! 5. A save resumed and stepped once is the original stepped once more.
//! 6. Two machines given the same rows show the same grid and the same
//!    plane, byte for byte, frame after frame, which a hash order leaking
//!    into a frame would break.
//! 7. A grid written between frames is planned as a load of the grid it
//!    leaves plans it: the census after a write and a frame is the census of
//!    a machine made from the same grid (KERNEL.34).
//! 8. A grid's frames are the same however its formulas are written, one
//!    fill each, one fill a row, one fill a column or a formula a cell: the
//!    record cell for cell, frame after frame, the values where a formula a
//!    cell is written with a text of its own; a grid written by rows past
//!    sixteen of them is one whose readers the machine indexes (KERNEL.35).
//! 9. A window's values are its grid record less its formula rows, every
//!    sheet's, twins included, before a frame, after each and after each
//!    write between frames; and a record's window and extent, as the machine
//!    keeps them, are the sheet's walked (KERNEL.37). Its grids' formats
//!    hold codes and looks map values to them, so each record's `text` rows
//!    are held too (KERNEL.47).
//! 10. Each twin after a frame is its sheet's values as the frame left
//!     them, however many of its cells the frame changed: the twin's record
//!     holds the sheet's value cells and its formulas' values as cells, and
//!     the plane the machine keeps is a fresh load's, over generated grids
//!     most of whose cells change at every frame, with writes between
//!     frames (KERNEL.36).
//! 11. A window's values at the frame before the last complete one, patched
//!     by its changes, are its values now, byte for byte, the changes
//!     naming no cell exactly when the two are equal, or the changes answer
//!     whole: every sheet's and twin's, whole and through windows, right
//!     after a frame, after the writes before the next and while it yields,
//!     over the grids of laws 9, 10 and 8, and grids whose values turn
//!     between ones their records spell alike (KERNEL.42); a cell's `text`
//!     row among its rows (KERNEL.47).
//! 12. A frame's count is the formula cells whose values it changed to the
//!     bit, read before its first step and after its last, each sheet's
//!     count is its share, and the counts during a yield are the last
//!     frame's, stepped whole and in slices; and a frame that changed
//!     nothing, with no write before it, is a fixed point: the next frame
//!     with none before it either changes nothing and leaves every value as
//!     it was (KERNEL.48).
//! 13. Far-apart cells of one text compute what the same cells compute each
//!     written with a text of its own, frame for frame through value writes:
//!     a few texts each copied to scattered cells, one shape however far
//!     apart its cells stand, which the machine cuts into pieces, held to
//!     shapes never cut (KERNEL.54).

use std::collections::{BTreeMap, BTreeSet};

use vla_lang::calc::graph::is_formula;
use vla_lang::calc::library;
use vla_lang::calc::{CellId, Values};
use vla_lang::expand::expand_text;
use vla_lang::form::Form;
use vla_lang::machine::{Draft, Machine, Viewed};
use vla_lang::printer::write_datum;
use vla_lang::reader::read_forms;
use vla_lang::refers::shift_a1_references;
use vla_lang::rows::quoted;
use vla_lang::sheet::{parse_a1_range, A1Range};
use vla_lang::view::{view_text, window_of};

const PRELUDE: &str = include_str!("../testdata/prelude.vla");
const GOLDEN: &str = include_str!("../testdata/instructions_golden.vla");

/// splitmix64, the generator's whole randomness: seeded, so every run
/// writes the same programs.
struct Rng(u64);

impl Rng {
    fn new(seed: u64) -> Rng {
        Rng(seed)
    }

    fn next(&mut self) -> u64 {
        self.0 = self.0.wrapping_add(0x9E37_79B9_7F4A_7C15);
        let mut z = self.0;
        z = (z ^ (z >> 30)).wrapping_mul(0xBF58_476D_1CE4_E5B9);
        z = (z ^ (z >> 27)).wrapping_mul(0x94D0_49BB_1331_11EB);
        z ^ (z >> 31)
    }

    fn below(&mut self, n: usize) -> usize {
        (self.next() % n as u64) as usize
    }

    fn pick<'a, T>(&mut self, xs: &'a [T]) -> &'a T {
        &xs[self.below(xs.len())]
    }
}

// ---- 1. reading and printing ------------------------------------------

/// What a symbol may be made of: anything but a delimiter.
const SYMBOL_CHARS: &[char] = &[
    'a', 'b', 'z', 'A', 'Z', '0', '9', '-', '+', '*', '/', '<', '=', '>', '!', '?', '.', '_', '&',
    '#', '$', '%', ':', '\'', '`', ',', '@', '[', ']', '{', '}', '|', '~', '^', 'é', 'λ',
];

/// What a string may hold, the printer's escapes and the delimiters among it.
const STRING_CHARS: &[char] = &[
    'a', ' ', '"', '\\', '(', ')', ';', '\t', '\n', '\r', 'é', 'λ', '=', '1',
];

fn chars(r: &mut Rng, pool: &[char], most: usize) -> String {
    let n = 1 + r.below(most);
    (0..n).map(|_| *r.pick(pool)).collect()
}

/// One form's text: a symbol, a string written with the reader's escapes,
/// or a list of up to three of them, six levels deep at most.
fn form_text(r: &mut Rng, depth: u32) -> String {
    match r.below(if depth >= 5 { 2 } else { 4 }) {
        0 => chars(r, SYMBOL_CHARS, 6),
        1 => {
            let body = chars(r, STRING_CHARS, 8)
                .replace('\\', "\\\\")
                .replace('"', "\\\"");
            format!("\"{body}\"")
        }
        _ => {
            let k = r.below(4);
            let items: Vec<String> = (0..k).map(|_| form_text(r, depth + 1)).collect();
            format!("({})", items.join(" "))
        }
    }
}

#[test]
fn reading_what_the_printer_wrote_gives_the_same_forms() {
    for seed in 1..=3u64 {
        let mut r = Rng::new(seed);
        for case in 0..1000 {
            let src = form_text(&mut r, 0);
            let forms = read_forms(&src)
                .unwrap_or_else(|e| panic!("seed {seed} case {case}: {src:?} refused: {}", e.text));
            let printed: Vec<String> = forms.iter().map(write_datum).collect();
            let again = read_forms(&printed.join("\n")).unwrap_or_else(|e| {
                panic!("seed {seed} case {case}: {printed:?} refused: {}", e.text)
            });
            assert_eq!(again, forms, "seed {seed} case {case}: {src:?}");
            let reprinted: Vec<String> = again.iter().map(write_datum).collect();
            assert_eq!(reprinted, printed, "seed {seed} case {case}: {src:?}");
        }
    }
}

// ---- 2. expansion ------------------------------------------------------

/// The prelude's macros: each name, its fixed parameters' count, and
/// whether a rest parameter follows them.
fn prelude_macros() -> Vec<(String, usize, bool)> {
    let forms = read_forms(PRELUDE).expect("the prelude reads");
    forms
        .iter()
        .filter_map(|f| {
            let Form::List(l) = f else { return None };
            let (Some(Form::Sym(head)), Some(Form::List(sig))) = (l.items.first(), l.items.get(1))
            else {
                return None;
            };
            if head != "defmacro" {
                return None;
            }
            let Some(Form::Sym(name)) = sig.items.first() else {
                return None;
            };
            let params = &sig.items[1..];
            let rest = params
                .iter()
                .position(|p| matches!(p, Form::Sym(s) if s == "&"));
            Some((name.clone(), rest.unwrap_or(params.len()), rest.is_some()))
        })
        .collect()
}

/// A call of one of the prelude's macros, its arguments variables, numbers,
/// quoted lists, core forms, or calls again, three levels deep at most.
fn call(r: &mut Rng, macros: &[(String, usize, bool)], depth: u32) -> String {
    let (name, fixed, rest) = r.pick(macros);
    let extra = if *rest { r.below(3) } else { 0 };
    let args: Vec<String> = (0..fixed + extra)
        .map(|_| argument(r, macros, depth + 1))
        .collect();
    if args.is_empty() {
        format!("({name})")
    } else {
        format!("({name} {})", args.join(" "))
    }
}

fn argument(r: &mut Rng, macros: &[(String, usize, bool)], depth: u32) -> String {
    match r.below(if depth >= 3 { 4 } else { 5 }) {
        0 => format!("x{}", r.below(9)),
        1 => r.below(100).to_string(),
        2 => {
            let items: Vec<String> = (0..r.below(4)).map(|i| format!("q{i}")).collect();
            format!("(quote ({}))", items.join(" "))
        }
        3 => format!("(debug-print x{})", r.below(9)),
        _ => call(r, macros, depth),
    }
}

/// Expanding what an expansion wrote fires nothing and changes nothing.
fn assert_fixpoint(program: &str, what: &str) {
    let (once, _) = expand_text(program, PRELUDE, true)
        .unwrap_or_else(|e| panic!("{what}: refused: {}", e.text));
    let (twice, fired) = expand_text(&once, PRELUDE, true)
        .unwrap_or_else(|e| panic!("{what}: its expansion refused: {}\n{once}", e.text));
    assert_eq!(
        fired, 0,
        "{what}: a second pass fired {fired} macro(s)\n{program}"
    );
    assert_eq!(
        twice, once,
        "{what}: a second pass changed the text\n{program}"
    );
}

#[test]
fn expansion_reaches_its_fixpoint_in_one_pass() {
    assert_fixpoint(GOLDEN, "the corpus golden");
    let macros = prelude_macros();
    assert_eq!(macros.len(), 46, "the prelude's macros, read off it");
    let (mut expanded, mut refused) = (0, 0);
    for seed in 1..=4u64 {
        let mut r = Rng::new(seed);
        for case in 0..30 {
            let program = call(&mut r, &macros, 0);
            if expand_text(&program, PRELUDE, true).is_err() {
                // A call whose arguments a macro cannot take (a number where
                // a quoted list belongs) is refused by name; the law speaks
                // of what expands.
                refused += 1;
                continue;
            }
            expanded += 1;
            assert_fixpoint(&program, &format!("seed {seed} case {case}"));
        }
    }
    assert!(
        expanded >= 60,
        "{expanded} programs expanded and {refused} were refused: too few to hold the law over"
    );
}

// ---- 3 to 6. the machine -----------------------------------------------

const SHEET: &str = "G";
const COLUMNS: &[&str] = &["A", "B", "C", "D", "E", "F"];
const ROWS: usize = 5;

/// A grid's rows over A1:F5: empty cells, numbers, texts and formulas. A
/// formula reads only the cells before it in reading order this frame, and
/// any cell of the last frame through the twin, so no frame holds a cycle.
fn grid_rows(r: &mut Rng) -> String {
    let all: Vec<String> = (0..ROWS)
        .flat_map(|row| COLUMNS.iter().map(move |c| format!("{c}{}", row + 1)))
        .collect();
    let mut out = format!("(sheet \"{SHEET}\" visible)\n");
    for (i, a) in all.iter().enumerate() {
        match r.below(10) {
            0..=2 => {}
            3 | 4 => {
                let n = r.below(201) as f64 / 2.0;
                out.push_str(&format!("(cell \"{SHEET}\" \"{a}\" {n})\n"));
            }
            5 => {
                let t = r.below(10);
                out.push_str(&format!("(cell \"{SHEET}\" \"{a}\" \"t{t}\")\n"));
            }
            _ => {
                let f = expr(r, &all[..i], &all, i / COLUMNS.len(), 0);
                out.push_str(&format!("(formula \"{SHEET}\" \"{a}\" \"={f}\")\n"));
            }
        }
    }
    out
}

/// A formula over the earlier cells, the last frame's, numbers and the
/// language library's functions, three levels deep at most, the lookups
/// and the criteria among them over the last frame's columns and rows
/// every `$` holds and over the rows above (KERNEL.8). `row` is the
/// formula's row less one, so `A1:F{row}` is every row above it.
fn expr(r: &mut Rng, earlier: &[String], all: &[String], row: usize, depth: u32) -> String {
    if depth >= 3 || r.below(3) == 0 {
        return match r.below(4) {
            0 if !earlier.is_empty() => r.pick(earlier).clone(),
            1 => format!("{SHEET}.last!{}", r.pick(all)),
            _ => r.below(10).to_string(),
        };
    }
    let e = |r: &mut Rng| expr(r, earlier, all, row, depth + 1);
    match r.below(11) {
        0 => format!("{}+{}", e(r), e(r)),
        1 => format!("{}-{}", e(r), e(r)),
        2 => format!("({})*{}", e(r), e(r)),
        3 => format!("{}/{}", e(r), e(r)),
        4 => format!("{}&{}", e(r), e(r)),
        5 => format!("IF({}>{},{},{})", e(r), e(r), e(r), e(r)),
        6 if row > 0 => format!("SUM(A1:F{row})"),
        6 => format!("SUM({},{})", e(r), e(r)),
        7 => format!("MOD({},{})", e(r), e(r)),
        8 => format!("CHOOSE({},{},{})", 1 + r.below(2), e(r), e(r)),
        9 => {
            let c = r.pick(COLUMNS);
            match r.below(9) {
                0 => format!(
                    "MATCH({},{SHEET}.last!${c}$1:${c}${ROWS},{})",
                    e(r),
                    r.pick(&["0", "1", "-1"])
                ),
                1 => {
                    let n = 1 + r.below(ROWS);
                    format!("XMATCH({},{SHEET}.last!$A${n}:$F${n},1)", e(r))
                }
                2 => format!("COUNTIF({SHEET}.last!$A$1:$F${ROWS},{})", e(r)),
                3 if row > 0 => format!("SUMIF(A1:F{row},{},A1:F{row})", e(r)),
                4 => format!(
                    "IFERROR(VLOOKUP({},{SHEET}.last!$A$1:$C${ROWS},2,FALSE),{})",
                    e(r),
                    e(r)
                ),
                // Arrays (KERNEL.8, slice f): an array parameter given an
                // expression, an array function, and `:` between two
                // expressions over the rows above.
                6 => format!(
                    "SUMPRODUCT(({SHEET}.last!${c}$1:${c}${ROWS}>{})*{SHEET}.last!$A$1:$A${ROWS})",
                    e(r)
                ),
                7 => format!(
                    "INDEX(SORT({SHEET}.last!${c}$1:${c}${ROWS},1,{}),{})",
                    r.pick(&["1", "-1"]),
                    1 + r.below(ROWS)
                ),
                8 if row > 0 => format!("SUM(A1:INDEX(A1:F{row},{}))", 1 + r.below(row)),
                _ => format!(
                    "INDEX({SHEET}.last!$A$1:$F${ROWS},{},{})",
                    1 + r.below(ROWS),
                    1 + r.below(6)
                ),
            }
        }
        _ => format!("MAX(ABS({}),INT({}))", e(r), e(r)),
    }
}

fn load(rows: &str, what: &str) -> Machine {
    Machine::load(rows, library::language())
        .unwrap_or_else(|e| panic!("{what}: line {}: {}\n{rows}", e.line, e.refusal.text))
}

/// A machine's save: the grid view of every sheet a program sees.
fn save(m: &Machine) -> String {
    (0..m.sheets())
        .filter_map(|i| m.sheet_name(i))
        .map(|name| match m.view("grid", name, None) {
            Ok(Viewed::Text(t)) => t,
            other => panic!("{other:?}"),
        })
        .collect()
}

fn plane(m: &Machine) -> Vec<u8> {
    match m.view("plane", SHEET, Some("A1:F5")) {
        Ok(Viewed::Bytes(b)) => b,
        other => panic!("{other:?}"),
    }
}

#[test]
fn loading_a_view_and_viewing_it_again_gives_the_same_view() {
    let mut texts = 0;
    let mut groups = 0;
    for seed in 1..=50u64 {
        let mut r = Rng::new(seed);
        let rows = grid_rows(&mut r);
        // Every second grid dressed with formats whose codes print `text`
        // rows, which the load skips and the view prints again (KERNEL.47).
        let rows = match seed % 2 {
            0 => dressed(&mut r, &rows),
            _ => rows,
        };
        let mut m = load(&rows, &format!("seed {seed}"));
        for frame in 0..=2 {
            let saved = save(&m);
            let again = load(&saved, &format!("seed {seed}, frame {frame}'s save"));
            assert_eq!(save(&again), saved, "seed {seed}, frame {frame}\n{rows}");
            texts += saved.lines().filter(|l| l.starts_with("(text ")).count();
            groups += saved
                .lines()
                .filter(|l| l.starts_with("(format ") && l.contains(" (font "))
                .count();
            assert!(m.step(0).done);
        }
    }
    // Measured 2026-10-17 (KERNEL.47): 318 text rows in the saves.
    assert!(
        texts >= 150,
        "too few text rows for the law to hold over: {texts}"
    );
    // Measured 2026-10-21 (KERNEL.50): 126 format rows saying a font,
    // an alignment and a border in the saves.
    assert!(
        groups >= 60,
        "too few format rows saying a font for the law to hold over: {groups}"
    );
}

#[test]
fn a_frame_stepped_in_slices_is_the_frame_stepped_whole() {
    for seed in 1..=50u64 {
        let mut r = Rng::new(seed);
        let rows = grid_rows(&mut r);
        let mut whole = load(&rows, &format!("seed {seed}"));
        let mut sliced = load(&rows, &format!("seed {seed}"));
        for frame in 1..=3 {
            assert!(whole.step(0).done);
            let mut done = false;
            for _ in 0..1000 {
                if sliced.step(1 + r.below(4)).done {
                    done = true;
                    break;
                }
            }
            assert!(
                done,
                "seed {seed}, frame {frame}: the slices never finished"
            );
            assert_eq!(
                save(&sliced),
                save(&whole),
                "seed {seed}, frame {frame}\n{rows}"
            );
        }
    }
}

#[test]
fn a_save_resumed_and_stepped_once_is_the_original_stepped_once_more() {
    for seed in 1..=50u64 {
        let mut r = Rng::new(seed);
        let rows = grid_rows(&mut r);
        let mut m = load(&rows, &format!("seed {seed}"));
        for _ in 0..1 + r.below(3) {
            assert!(m.step(0).done);
        }
        let mut again = load(&save(&m), &format!("seed {seed}'s save"));
        again
            .resume_at(m.frame())
            .unwrap_or_else(|e| panic!("seed {seed}: {}", e.text));
        assert!(m.step(0).done);
        assert!(again.step(0).done);
        assert_eq!(again.frame(), m.frame(), "seed {seed}");
        assert_eq!(save(&again), save(&m), "seed {seed}\n{rows}");
    }
}

#[test]
fn two_machines_given_the_same_rows_show_the_same_frames() {
    for seed in 1..=50u64 {
        let mut r = Rng::new(seed);
        let rows = grid_rows(&mut r);
        let mut a = load(&rows, &format!("seed {seed}"));
        let mut b = load(&rows, &format!("seed {seed}"));
        for frame in 1..=3 {
            assert!(a.step(0).done && b.step(0).done);
            assert_eq!(save(&a), save(&b), "seed {seed}, frame {frame}\n{rows}");
            assert_eq!(plane(&a), plane(&b), "seed {seed}, frame {frame}\n{rows}");
        }
    }
}

/// A write between two frames: a formula over a rectangle of the grid,
/// reading the last frame and a number from its first cell, which becomes
/// one fill; the last such write again, which writes nothing where its
/// cells are as it left them; or a number over a cell, perhaps a fill's
/// first, which gives it up.
fn write_between(r: &mut Rng, last: &mut Option<String>) -> String {
    let (top, left) = (1 + r.below(ROWS), r.below(COLUMNS.len()));
    let bottom = top + r.below(ROWS - top + 1);
    let right = left + r.below(COLUMNS.len() - left);
    let first = format!("{}{top}", COLUMNS[left]);
    match r.below(4) {
        0 => format!("(cell \"{SHEET}\" \"{first}\" {})", r.below(10)),
        1 if last.is_some() => last.clone().unwrap_or_default(),
        _ => {
            let write = format!(
                "(formula \"{SHEET}\" \"{first}:{}{bottom}\" \"={SHEET}.last!{first}*2+{}\")",
                COLUMNS[right],
                r.below(10)
            );
            *last = Some(write.clone());
            write
        }
    }
}

#[test]
fn a_grid_written_between_frames_is_planned_as_a_load_of_it() {
    let mut written = 0;
    for seed in 1..=50u64 {
        let mut r = Rng::new(seed);
        let rows = grid_rows(&mut r);
        let mut m = load(&rows, &format!("seed {seed}"));
        let mut last = None;
        for frame in 1..=4 {
            assert!(m.step(0).done);
            let write = write_between(&mut r, &mut last);
            let when = format!("seed {seed}, frame {frame}: {write}\n{rows}");
            m.write(&write)
                .unwrap_or_else(|e| panic!("{when}: {}", e.refusal.text));
            written += 1;
            assert!(m.step(0).done);
            let sheets = m.sheets();
            let mut workbook = m.workbook();
            workbook.sheets.truncate(sheets);
            let draft = Draft::of_workbook(&workbook);
            let mut again = Machine::new(draft, library::language())
                .unwrap_or_else(|e| panic!("{when}: line {}: {}", e.line, e.refusal.text));
            assert!(again.step(0).done);
            assert_eq!(m.census(), again.census(), "{when}");
        }
    }
    assert_eq!(written, 200);
}

// ---- 8. the ways a grid's formulas are written ---------------------------

/// The columns a value sheet holds, and the columns a fill covers.
const VALUES: &[&str] = &["A", "B", "C", "D", "E", "F", "G", "H"];

/// A fill's formula at its first cell, in row `top`: the value sheet's cells
/// around its own, a column held, a cell held, a range and a part every `$`
/// holds, the last frame's cells, and the cell above its own once it is
/// not the first row, under the language library's functions, the lookups
/// and the criteria among them over the value sheet's columns every `$`
/// holds, whose indexes a fill's cells share (KERNEL.8).
fn fill_expr(r: &mut Rng, top: usize, depth: u32) -> String {
    if depth >= 2 || r.below(3) == 0 {
        return match r.below(6) {
            0 => format!("V!{}{top}", r.pick(&["A", "B", "C"])),
            1 => format!("V!${}{}", r.pick(&["A", "B"]), top + r.below(2)),
            2 => format!("V!$A${}", 1 + r.below(5)),
            3 => format!("F.last!{}{top}", r.pick(&["C", "D", "E"])),
            4 if top > 1 => format!("C{}", top - 1),
            _ => r.below(10).to_string(),
        };
    }
    let e = |r: &mut Rng| fill_expr(r, top, depth + 1);
    match r.below(7) {
        6 => match r.below(4) {
            0 => format!("XLOOKUP({},V!$A$1:$A$5,V!$B$1:$B$5,-1)", e(r)),
            1 => format!("COUNTIF(V!$A$1:$B$5,{})", e(r)),
            // An array parameter given an expression (KERNEL.8, slice f).
            2 => format!("SUMPRODUCT(V!$A$1:$A$5*(V!$B$1:$B$5>{}))", e(r)),
            _ => format!("INDEX(V!$A$1:$B$5,MOD({},5)+1,2)", e(r)),
        },
        0 => format!("{}+{}", e(r), e(r)),
        1 => format!("{}*{}", e(r), e(r)),
        2 => format!("SUM(V!A{top}:B{})", top + 1),
        3 => format!("IF(V!$A$1>{},{},{})", r.below(5), e(r), e(r)),
        4 => format!("MOD({},7)", e(r)),
        _ => format!("MAX({},SUM(V!$A$1:$B$2))", e(r)),
    }
}

/// A grid of a value sheet `V` and a sheet `F` of fills over `C1:H{rows}`,
/// each a band of up to six rows written as one fill.
fn fill_rows(r: &mut Rng, rows: usize) -> String {
    let mut out = String::from("(sheet \"F\" visible)\n(sheet \"V\" visible)\n");
    for row in 1..=rows + 2 {
        for col in VALUES {
            if r.below(4) != 0 {
                out.push_str(&format!("(cell \"V\" \"{col}{row}\" {})\n", r.below(10)));
            }
        }
    }
    let mut top = 1;
    while top <= rows {
        let bottom = (top + r.below(6)).min(rows);
        let f = fill_expr(r, top, 0);
        out.push_str(&format!("(formula \"F\" \"C{top}:H{bottom}\" \"={f}\")\n"));
        top = bottom + 1;
    }
    out
}

/// A grid's rows with each fill written `"rows"` one fill a row, `"cols"`
/// one fill a column, or `"cells"` one formula a cell, each with a text of
/// its own that computes the same value; each part's text the fill's moved
/// to the part's first cell, as the fill reads there.
fn written(rows: &str, how: &str) -> String {
    let mut out = String::new();
    let mut n = 0;
    for line in rows.lines() {
        let fill = read_forms(line).ok().and_then(|forms| match forms.first() {
            Some(Form::List(l)) => match l.items.as_slice() {
                [Form::Sym(head), Form::Str(sheet), Form::Str(addr), Form::Str(text)]
                    if head == "formula" =>
                {
                    Some((sheet.clone(), parse_a1_range(addr)?, text.clone()))
                }
                _ => None,
            },
            _ => None,
        });
        let Some((sheet, r, text)) = fill.filter(|f| f.1.cells() > 1) else {
            out.push_str(line);
            out.push('\n');
            continue;
        };
        let moved = |row: u32, col: u32| {
            let (d_row, d_col) = (i64::from(row - r.top), i64::from(col - r.left));
            shift_a1_references(&text[1..], d_row, d_col)
        };
        let mut put = |part: A1Range, text: String| {
            out.push_str(&format!(
                "(formula {} {} {})\n",
                quoted(&sheet),
                quoted(&part.text()),
                quoted(&text)
            ));
        };
        for row in r.top..=r.bottom {
            for col in r.left..=r.right {
                let (first_row, first_col) = (row == r.top, col == r.left);
                match how {
                    "rows" if first_col => put(
                        A1Range {
                            top: row,
                            bottom: row,
                            ..r
                        },
                        format!("={}", moved(row, col)),
                    ),
                    "cols" if first_row => put(
                        A1Range {
                            left: col,
                            right: col,
                            ..r
                        },
                        format!("={}", moved(row, col)),
                    ),
                    "cells" => {
                        n += 1;
                        let cell = A1Range {
                            top: row,
                            left: col,
                            bottom: row,
                            right: col,
                        };
                        put(cell, format!("=({})+0*{n}", moved(row, col)));
                    }
                    _ => {}
                }
            }
        }
    }
    out
}

/// A machine's value rows, every sheet a program sees: what a formula a
/// cell with a text of its own computes alike.
fn values(m: &Machine) -> Vec<String> {
    save(m)
        .lines()
        .filter(|l| l.starts_with("(value "))
        .map(str::to_string)
        .collect()
}

#[test]
fn a_grid_s_frames_are_the_same_however_its_formulas_are_written() {
    let mut indexed = 0;
    for seed in 1..=30u64 {
        let mut r = Rng::new(seed);
        let rows = 8 + r.below(17);
        let grid = fill_rows(&mut r, rows);
        let ways = ["one", "rows", "cols", "cells"];
        let mut machines: Vec<Machine> = ways
            .iter()
            .map(|&how| {
                let text = if how == "one" {
                    grid.clone()
                } else {
                    written(&grid, how)
                };
                load(&text, &format!("seed {seed}, {how}"))
            })
            .collect();
        for frame in 1..=4 {
            for m in &mut machines {
                assert!(m.step(0).done);
            }
            let when = format!("seed {seed}, frame {frame}");
            let one = save(&machines[0]);
            assert_eq!(save(&machines[1]), one, "{when}, by rows\n{grid}");
            assert_eq!(save(&machines[2]), one, "{when}, by columns\n{grid}");
            assert_eq!(
                values(&machines[3]),
                values(&machines[0]),
                "{when}, a cell each\n{grid}"
            );
            let cell = format!("{}{}", r.pick(VALUES), 1 + r.below(rows + 2));
            let write = format!("(cell \"V\" \"{cell}\" {})", r.below(10));
            for m in &mut machines {
                m.write(&write)
                    .unwrap_or_else(|e| panic!("{when}: {write}: {}", e.refusal.text));
            }
        }
        if rows >= 16 && machines[1].census().contains(" (most ") {
            indexed += 1;
        }
    }
    assert!(
        indexed >= 8,
        "{indexed} grids written by rows were indexed: too few for the law to hold over"
    );
}

// ---- 9. a window's values ------------------------------------------------

/// A grid's rows dressed: formats, and a `style` or a `sentence` row after
/// some of its written cells, so that the rows the values keep beside a
/// formula's stand in the records (KERNEL.37). Three of the formats hold
/// codes, a number's, a sign's and a text's, and two looks map a value to
/// them, so that a cell's `text` row stands in the records too (KERNEL.47);
/// two say a font, an alignment and borders, so that their groups stand in
/// the records (KERNEL.50).
fn dressed(r: &mut Rng, rows: &str) -> String {
    let cell = format!("(cell \"{SHEET}\" \"");
    let formula = format!("(formula \"{SHEET}\" \"");
    let mut out = String::new();
    for line in rows.lines() {
        out.push_str(line);
        out.push('\n');
        if line.starts_with("(sheet ") {
            out.push_str("(format 1 \"F7F4FC\" general nowrap (font bold 14 \"6B4FC8\") (align center middle) (border (left thin) (bottom double \"FF0000\")))\n");
            out.push_str("(format 2 none \"0.00\" nowrap (font not-italic 9.5) (align right top) (border (top medium-dash-dot)))\n");
            out.push_str("(format 3 none \"#,##0.0;(#,##0.0);\\\"zero\\\"\" nowrap)\n");
            out.push_str("(format 4 \"E2D8F6\" \"0;-0;0;\\\"<\\\"@\\\">\\\"\" nowrap)\n");
            out.push_str(&format!(
                "(look \"{SHEET}\" 0 3)\n(look \"{SHEET}\" \"t1\" 4)\n"
            ));
        }
        let Some(addr) = line
            .strip_prefix(&cell)
            .or_else(|| line.strip_prefix(&formula))
            .and_then(|rest| rest.split('"').next())
        else {
            continue;
        };
        if r.below(3) == 0 {
            let style = 1 + r.below(4);
            out.push_str(&format!("(style \"{SHEET}\" \"{addr}\" {style})\n"));
        }
        if r.below(3) == 0 {
            let row = 1 + r.below(9);
            out.push_str(&format!("(sentence \"{SHEET}\" \"{addr}\" {row})\n"));
        }
    }
    out
}

/// A window of a sheet, as a projection prints it.
fn record(m: &Machine, projection: &str, sheet: &str, window: Option<&str>) -> String {
    match m.view(projection, sheet, window) {
        Ok(Viewed::Text(t)) => t,
        other => panic!("{projection} {sheet} {window:?}: {other:?}"),
    }
}

/// A record's lines less its `formula` lines.
fn less_formulas(record: &str) -> String {
    record
        .lines()
        .filter(|l| !l.starts_with("(formula "))
        .map(|l| format!("{l}\n"))
        .collect()
}

/// A record's window and extent rows.
fn header(record: &str) -> Vec<&str> {
    record
        .lines()
        .filter(|l| l.starts_with("(window ") || l.starts_with("(extent "))
        .collect()
}

/// Every sheet's values, twins included, held to its grid record less its
/// formula lines, whole and through a window; and each whole record's window
/// and extent, read from the extents the machine keeps, held to the sheet's
/// walked in a copy of the workbook. The records held, counted.
fn values_held(m: &mut Machine, when: &str) -> usize {
    let names: Vec<String> = (0..)
        .map_while(|i| m.sheet_name(i).map(str::to_string))
        .collect();
    let mut held = 0;
    let mut wholes = Vec::new();
    for name in &names {
        for window in [None, Some("B2:D4")] {
            let grid = record(m, "grid", name, window);
            let shown = record(m, "values", name, window);
            assert_eq!(shown, less_formulas(&grid), "{when}: {name} {window:?}");
            held += 1;
            if window.is_none() {
                wholes.push(grid);
            }
        }
    }
    let book = m.workbook();
    for (name, grid) in names.iter().zip(&wholes) {
        let w = window_of(&book, name, None).expect("a sheet the machine holds");
        let walked = view_text(&book, &w);
        assert_eq!(
            header(grid),
            header(&walked),
            "{when}: {name}'s window and extent"
        );
    }
    held
}

#[test]
fn a_window_s_values_are_its_grid_less_its_formula_rows() {
    let mut held = 0;
    let mut texts = 0;
    let count = |m: &Machine| {
        record(m, "values", SHEET, None)
            .lines()
            .filter(|l| l.starts_with("(text "))
            .count()
    };
    for seed in 1..=50u64 {
        let mut r = Rng::new(seed);
        let rows = grid_rows(&mut r);
        let rows = dressed(&mut r, &rows);
        let mut m = load(&rows, &format!("seed {seed}"));
        let mut last = None;
        for frame in 0..=4 {
            if frame > 0 {
                assert!(m.step(0).done);
            }
            held += values_held(&mut m, &format!("seed {seed}, frame {frame}\n{rows}"));
            texts += count(&m);
            let write = write_between(&mut r, &mut last);
            let when = format!("seed {seed}, frame {frame}: {write}\n{rows}");
            m.write(&write)
                .unwrap_or_else(|e| panic!("{when}: {}", e.refusal.text));
            held += values_held(&mut m, &when);
            texts += count(&m);
        }
    }
    // Fifty grids, five frames, before and after a write, two sheets (the
    // grid's and its twin), each whole and through a window.
    assert_eq!(held, 50 * 5 * 2 * 2 * 2);
    // Measured 2026-10-17 (KERNEL.47): 2,181 text rows in the values.
    assert!(
        texts >= 1000,
        "too few text rows for the law to hold over: {texts}"
    );
}

// ---- 10. a frame's end at any share -------------------------------------

/// A grid most of whose cells change at every frame, and its rows: a sheet
/// `G` of value cells in column A and formulas over `B1:F{rows}` in bands,
/// each band one formula, most reading the cell `T!A1` a test writes before
/// every frame, some reading `G`'s twin, some a text and a number by turns,
/// some keeping their values; in some grids a value cell far off, so that
/// `G` is too sparse for an array (KERNEL.36).
fn churning_rows(r: &mut Rng) -> (String, usize) {
    let rows = 4 + r.below(40);
    let mut out =
        String::from("(sheet \"T\" visible)\n(cell \"T\" \"A1\" 0)\n(sheet \"G\" visible)\n");
    for row in 1..=rows {
        if r.below(3) != 0 {
            out.push_str(&format!("(cell \"G\" \"A{row}\" {})\n", r.below(10)));
        }
    }
    let mut top = 1;
    while top <= rows {
        let bottom = (top + r.below(5)).min(rows);
        let (a, b) = (1 + r.below(9), 1 + r.below(9));
        let f = match r.below(6) {
            0 => format!("ROW()*{a}+COLUMN()"),
            1 => format!("IF(MOD(T!$A$1+ROW(),2)=0,\"t\"&ROW(),ROW()*{a})"),
            2 => format!("G.last!{}{top}+T!$A$1", r.pick(&["B", "C", "D", "E", "F"])),
            3 => format!("MOD(T!$A$1*{a}+ROW()*{b}+COLUMN(),97)+$A{top}"),
            _ => format!("MOD(T!$A$1*{a}+ROW()*{b}+COLUMN(),97)"),
        };
        out.push_str(&format!(
            "(formula \"G\" \"B{top}:F{bottom}\" {})\n",
            quoted(&format!("={f}"))
        ));
        top = bottom + 1;
    }
    if r.below(3) == 0 {
        out.push_str("(cell \"G\" \"AZ999\" 1)\n");
    }
    (out, rows)
}

/// A record's data by cell, each datum printed again: the rows headed by
/// one of `heads`, a value cell's `cell` row or a formula's `value` row.
fn data(record: &str, heads: &[&str]) -> BTreeMap<String, String> {
    read_forms(record)
        .expect("a record reads")
        .iter()
        .filter_map(|f| match f {
            Form::List(l) => match l.items.as_slice() {
                [Form::Sym(h), Form::Str(_), Form::Str(addr), datum]
                    if heads.contains(&h.as_str()) =>
                {
                    Some((addr.clone(), write_datum(datum)))
                }
                _ => None,
            },
            _ => None,
        })
        .collect()
}

#[test]
fn each_twin_is_its_sheet_s_values_however_many_cells_a_frame_changes() {
    let mut checked = 0;
    for seed in 1..=40u64 {
        let mut r = Rng::new(seed ^ 0x36);
        let (rows, n) = churning_rows(&mut r);
        let mut m = load(&rows, &format!("seed {seed}"));
        for frame in 1..=8 {
            let when = format!("seed {seed}, frame {frame}\n{rows}");
            let mut writes = vec![format!("(cell \"T\" \"A1\" {frame})")];
            match r.below(6) {
                0 => writes.push(format!(
                    "(cell \"G\" \"A{}\" {})",
                    1 + r.below(n),
                    r.below(10)
                )),
                1 => writes.push(format!(
                    "(cell \"G\" \"{}{}\" 5)",
                    r.pick(&["B", "C", "D", "E", "F"]),
                    1 + r.below(n)
                )),
                _ => {}
            }
            for w in &writes {
                m.write(w)
                    .unwrap_or_else(|e| panic!("{when}: {w}: {}", e.refusal.text));
            }
            assert!(m.step(0).done);
            let grid = |name: &str| match m.view("grid", name, None) {
                Ok(Viewed::Text(t)) => t,
                other => panic!("{when}: {other:?}"),
            };
            for sheet in ["T", "G"] {
                let values: BTreeMap<String, String> = data(&grid(sheet), &["cell", "value"])
                    .into_iter()
                    .filter(|(_, d)| !d.starts_with("(not-computed"))
                    .collect();
                let twin = data(&grid(&format!("{sheet}.last")), &["cell"]);
                assert_eq!(twin, values, "{when}: {sheet}'s twin");
            }
            let kept = match m.view("plane", "G", Some("A1:F44")) {
                Ok(Viewed::Bytes(b)) => b,
                other => panic!("{when}: {other:?}"),
            };
            let again = load(&save(&m), &format!("{when}: its save"));
            match again.view("plane", "G", Some("A1:F44")) {
                Ok(Viewed::Bytes(b)) => assert_eq!(kept, b, "{when}: G's plane"),
                other => panic!("{when}: {other:?}"),
            }
            checked += 1;
        }
    }
    assert_eq!(checked, 320);
}

// ---- 11. a window's changes ---------------------------------------------

/// The cell a record's row belongs to: a `cell`, `value`, `text`, `style`,
/// `sentence` or `empty` row's.
fn cell_of(line: &str) -> Option<(u32, u32)> {
    let forms = read_forms(line).ok()?;
    let Some(Form::List(l)) = forms.first() else {
        return None;
    };
    match l.items.as_slice() {
        [Form::Sym(head), Form::Str(_), Form::Str(addr), ..]
            if matches!(
                head.as_str(),
                "cell" | "value" | "text" | "style" | "sentence" | "empty"
            ) =>
        {
            let r = parse_a1_range(addr)?;
            Some((r.top, r.left))
        }
        _ => None,
    }
}

/// A record's rows before its cells', and each cell's rows.
fn rows_by_cell(record: &str) -> (String, BTreeMap<(u32, u32), String>) {
    let mut head = String::new();
    let mut cells: BTreeMap<(u32, u32), String> = BTreeMap::new();
    for line in record.lines() {
        match cell_of(line) {
            Some(cell) => cells
                .entry(cell)
                .or_default()
                .push_str(&format!("{line}\n")),
            None => head.push_str(&format!("{line}\n")),
        }
    }
    (head, cells)
}

/// The rows a changes answer gives each cell it names, none for an
/// `empty` row's; `None` for changes answered whole.
fn given(changes: &str) -> Option<BTreeMap<(u32, u32), String>> {
    let mut lines = changes.lines();
    if lines.next()?.ends_with(" whole)") {
        return None;
    }
    let mut given: BTreeMap<(u32, u32), String> = BTreeMap::new();
    for line in lines {
        let cell = cell_of(line).unwrap_or_else(|| panic!("a changes row names a cell: {line}"));
        let rows = given.entry(cell).or_default();
        if !line.starts_with("(empty ") {
            rows.push_str(&format!("{line}\n"));
        }
    }
    Some(given)
}

/// A values record patched by a window's changes, as a host patches the
/// record it holds: each cell the changes name has its rows replaced with
/// theirs, an `empty` row taking them away, every cell in the record's
/// order; `None` for changes answered whole.
fn patched(record: &str, changes: &str) -> Option<String> {
    let (head, mut cells) = rows_by_cell(record);
    for (cell, rows) in given(changes)? {
        match rows.is_empty() {
            true => cells.remove(&cell),
            false => cells.insert(cell, rows),
        };
    }
    Some(head + &cells.into_values().collect::<String>())
}

/// The windows the law is held over: every sheet's and twin's, whole and
/// through three windows, the last one of law 11's own grids' values that
/// a record spells alike.
fn change_windows(m: &Machine) -> Vec<(String, Option<&'static str>)> {
    (0..)
        .map_while(|i| m.sheet_name(i).map(str::to_string))
        .flat_map(|s| [None, Some("B2:D4"), Some("A3:F3"), Some("C4:C9")].map(|w| (s.clone(), w)))
        .collect()
}

fn values_of(m: &Machine, windows: &[(String, Option<&str>)]) -> Vec<String> {
    windows
        .iter()
        .map(|(s, w)| record(m, "values", s, *w))
        .collect()
}

/// The windows held: patched, still among them, and answered whole.
#[derive(Debug, Default)]
struct Held {
    patched: usize,
    still: usize,
    whole: usize,
    /// The `text` rows the changes of the windows patched gave (KERNEL.47).
    texts: usize,
}

/// The law at one moment: each window's values at `base`, the frame before
/// the last complete one, patched by its changes, against its values now.
/// The changes' first row names the window as the values' own `window` row
/// does, and the two frames.
fn changes_held(
    m: &Machine,
    windows: &[(String, Option<&str>)],
    base: &[String],
    held: &mut Held,
    when: &str,
) {
    let at = m.frame();
    for ((sheet, window), was) in windows.iter().zip(base) {
        let changes = record(m, "changes", sheet, *window);
        let now = record(m, "values", sheet, *window);
        let named = now
            .lines()
            .find_map(|l| l.strip_prefix("(window ").and_then(|w| w.strip_suffix(')')))
            .expect("a window row");
        let head = changes.lines().next().expect("a changes row");
        let frames = format!("(changes {named} {} {at} ", at.saturating_sub(1));
        assert!(
            head.starts_with(&frames),
            "{when}: {sheet} {window:?}: {head}"
        );
        match patched(was, &changes) {
            None => held.whole += 1,
            Some(got) => {
                assert_eq!(got, now, "{when}: {sheet} {window:?} patched by\n{changes}");
                let n: usize = head[frames.len()..head.len() - 1].parse().expect("a count");
                assert_eq!(
                    n == 0,
                    *was == now,
                    "{when}: {sheet} {window:?}: no cell named exactly when still\n{changes}"
                );
                // And a cell named exactly when its rows changed.
                let (_, before) = rows_by_cell(was);
                let (_, after) = rows_by_cell(&now);
                let changed = given(&changes).expect("not whole");
                assert_eq!(
                    changed.len(),
                    n,
                    "{when}: {sheet} {window:?}: the count\n{changes}"
                );
                for cell in changed.keys() {
                    assert_ne!(
                        before.get(cell),
                        after.get(cell),
                        "{when}: {sheet} {window:?}: {cell:?} named, its rows as they were\n{changes}"
                    );
                }
                held.patched += 1;
                held.texts += changes.lines().filter(|l| l.starts_with("(text ")).count();
                held.still += usize::from(n == 0);
            }
        }
    }
}

/// One grid's frames under the law: before each, the rows `writes` gives
/// written; then the frame stepped in slices of `budget` cells, 0 for one
/// step; the law held after the writes, while the frame yields, and after
/// it.
fn changes_over(
    m: &mut Machine,
    frames: u64,
    budget: usize,
    writes: &mut dyn FnMut(u64) -> Vec<String>,
    held: &mut Held,
    what: &str,
) {
    let windows = change_windows(m);
    let mut base: Option<Vec<String>> = None;
    let mut at = values_of(m, &windows);
    for frame in 1..=frames {
        for w in writes(frame) {
            m.write(&w)
                .unwrap_or_else(|e| panic!("{what}, frame {frame}: {w}: {}", e.refusal.text));
        }
        if let Some(b) = &base {
            changes_held(
                m,
                &windows,
                b,
                held,
                &format!("{what}, before frame {frame}"),
            );
        }
        if !m.step(budget).done {
            if let Some(b) = &base {
                changes_held(
                    m,
                    &windows,
                    b,
                    held,
                    &format!("{what}, during frame {frame}"),
                );
            }
            while !m.step(budget).done {}
        }
        let was = std::mem::replace(&mut at, values_of(m, &windows));
        changes_held(m, &windows, &was, held, &format!("{what}, frame {frame}"));
        base = Some(was);
    }
}

/// A grid whose formulas turn at every frame between values their records
/// spell alike: 0 and -0, read from cells holding them, in `C1:C{n}`, more
/// than the 64 a frame's end keeps in some grids; and reasons, one range
/// and another where one value is wanted, in `D1:D3`.
fn spelled_rows(r: &mut Rng) -> String {
    let n = 1 + r.below(80);
    let mut out = String::from(concat!(
        "(sheet \"Z\" visible)\n(cell \"Z\" \"A1\" 0)\n(cell \"Z\" \"A2\" -0)\n",
        "(formula \"Z\" \"B1\" \"=IF(Z.last!B1=1,0,1)\")\n",
        "(formula \"Z\" \"D1:D3\" \"=IF($B$1=1,X1:X2,Y1:Y2)\")\n",
        "(cell \"Z\" \"E1:E5\" 4)\n",
    ));
    out.push_str(&format!(
        "(formula \"Z\" \"C1:C{n}\" \"=IF($B$1=1,$A$1,$A$2)\")\n"
    ));
    out
}

#[test]
fn a_window_s_values_patched_by_its_changes_are_its_values_now() {
    // Each family's windows counted apart, so that each is held patched,
    // changing and still, and answered whole.
    let mut held: [Held; 4] = Default::default();
    // Law 9's grids, dressed, written before every second frame as law 7
    // writes them, some stepped in slices.
    for seed in 1..=50u64 {
        let mut r = Rng::new(seed ^ 0x42);
        let rows = grid_rows(&mut r);
        let rows = dressed(&mut r, &rows);
        let mut m = load(&rows, &format!("seed {seed}"));
        let budget = [0, 1, 3, 7][r.below(4)];
        let mut last = None;
        let mut writes = |frame: u64| match frame % 2 {
            0 => vec![write_between(&mut r, &mut last)],
            _ => Vec::new(),
        };
        let what = format!("seed {seed}\n{rows}");
        changes_over(&mut m, 6, budget, &mut writes, &mut held[0], &what);
    }
    // Law 10's, most of whose cells change at every frame, past the room a
    // frame's end keeps in some.
    for seed in 1..=40u64 {
        let mut r = Rng::new(seed ^ 0x4242);
        let (rows, n) = churning_rows(&mut r);
        let mut m = load(&rows, &format!("seed {seed}"));
        let budget = [0, 5, 40][r.below(3)];
        let mut writes = |frame: u64| {
            let mut w = vec![format!("(cell \"T\" \"A1\" {frame})")];
            if r.below(3) == 0 {
                w.push(format!(
                    "(cell \"G\" \"A{}\" {})",
                    1 + r.below(n),
                    r.below(10)
                ));
            }
            w
        };
        let what = format!("churning seed {seed}\n{rows}");
        changes_over(&mut m, 8, budget, &mut writes, &mut held[1], &what);
    }
    // Law 8's fills, one written whole and one a cell each.
    for seed in 1..=20u64 {
        let mut r = Rng::new(seed ^ 0x8);
        let rows = 8 + r.below(17);
        let grid = fill_rows(&mut r, rows);
        for text in [grid.clone(), written(&grid, "cells")] {
            let mut m = load(&text, &format!("fills seed {seed}"));
            let mut writes = |_: u64| {
                let cell = format!("{}{}", r.pick(VALUES), 1 + r.below(rows + 2));
                vec![format!("(cell \"V\" \"{cell}\" {})", r.below(10))]
            };
            let what = format!("fills seed {seed}");
            changes_over(&mut m, 4, 0, &mut writes, &mut held[2], &what);
        }
    }
    // Values that turn between ones their records spell alike.
    for seed in 1..=20u64 {
        let mut r = Rng::new(seed ^ 0x5A);
        let rows = spelled_rows(&mut r);
        let mut m = load(&rows, &format!("spelled seed {seed}"));
        let mut writes = |frame: u64| match frame % 3 {
            0 => {
                let v = r.pick(&["0", "-0", "4", "\"t\""]);
                vec![format!("(cell \"Z\" \"E{}\" {v})", 1 + r.below(5))]
            }
            _ => Vec::new(),
        };
        let what = format!("spelled seed {seed}\n{rows}");
        changes_over(&mut m, 6, 0, &mut writes, &mut held[3], &what);
    }
    // Measured 2026-10-15: patched 1,442, 7,833, 2,282 and 1,060, of which
    // still 670, 3,240, 1,924 and 135; whole 4,262, 4,343, 2,198 and 700.
    // The first family measured again 2026-10-17, its grids dressed with
    // codes (KERNEL.47): patched 1,851, still 834, whole 3,989. All four
    // measured again 2026-10-22, a cell added and a twin cell gained no
    // longer answering whole (KERNEL.56): patched 1,911, 9,535, 3,114 and
    // 1,060, of which still 813, 3,950, 2,585 and 398; whole 3,929, 2,641,
    // 1,366 and 700; 338 text rows.
    for (family, h) in ["grids", "churning", "fills", "spelled"].iter().zip(&held) {
        assert!(
            h.patched - h.still >= 150 && h.still >= 60 && h.whole >= 300,
            "{family}: too few windows of a kind for the law to hold over: {h:?}"
        );
    }
    // Measured 2026-10-17 (KERNEL.47), the grids' formats holding codes: 254
    // text rows in the changes of the windows patched.
    assert!(
        held[0].texts >= 120,
        "grids: too few text rows for the law to hold over: {:?}",
        held[0]
    );
}

// ---- 12. a frame's own count ---------------------------------------------

/// Every formula cell's value of the last complete frame, by its place, read
/// through the machine's values as any caller reads them and spelled to the
/// bit, a -0 apart from a 0.
fn formula_values(m: &mut Machine) -> BTreeMap<CellId, String> {
    let sheets = m.sheets();
    let ids: Vec<CellId> = {
        let book = m.workbook();
        (0..sheets)
            .flat_map(|s| {
                book.sheets[s]
                    .cells
                    .iter()
                    .filter(|(_, cell)| is_formula(&cell.content))
                    .map(move |(&(row, col), _)| (s, row, col))
            })
            .collect()
    };
    let m: &Machine = m;
    ids.into_iter()
        .filter_map(|id| Some((id, format!("{:?}", Values::computed(m, id)?))))
        .collect()
}

/// The frames counted: changing and still; stepped in slices and yielding;
/// and settled, a frame stepped after a still one with no write before
/// either.
#[derive(Debug, Default)]
struct Counted {
    changing: usize,
    still: usize,
    yielded: usize,
    settled: usize,
}

/// A frame stepped in slices of `budget` cells, 0 for one step, its counts
/// during a yield the last complete frame's.
fn step_counted(m: &mut Machine, budget: usize, counted: &mut Counted, when: &str) {
    let last = (m.changed(), m.reached());
    if !m.step(budget).done {
        counted.yielded += 1;
        let now = (m.changed(), m.reached());
        assert_eq!(now, last, "{when}: the counts during a yield");
        while !m.step(budget).done {}
    }
}

/// A frame's counts held to its formulas' values before its first step and
/// after its last: the count the cells whose values differ, a value gained
/// or spelled otherwise; each sheet's its share, a twin's none, none past
/// the grid's sheets; and the cells computed no fewer than those changed
/// and no more than the formulas.
fn counts_held(
    m: &Machine,
    before: &BTreeMap<CellId, String>,
    after: &BTreeMap<CellId, String>,
    when: &str,
) {
    let mut by_sheet: BTreeMap<usize, usize> = BTreeMap::new();
    let ids: BTreeSet<&CellId> = before.keys().chain(after.keys()).collect();
    for id in ids {
        if before.get(id) != after.get(id) {
            *by_sheet.entry(id.0).or_default() += 1;
        }
    }
    let total: usize = by_sheet.values().sum();
    assert_eq!(m.changed(), total, "{when}: the count");
    let sheets = (0..).map_while(|i| m.sheet_name(i)).count();
    for s in 0..sheets {
        let share = by_sheet.get(&s).copied().unwrap_or(0);
        let name = m.sheet_name(s).unwrap_or_default();
        assert_eq!(m.changed_on(s), Some(share), "{when}: {name}'s count");
    }
    assert_eq!(m.changed_on(sheets), None, "{when}: past the sheets");
    let (changed, reached) = (m.changed(), m.reached());
    assert!(
        changed <= reached && reached <= m.formulas().0,
        "{when}: {changed} changed of {reached} computed"
    );
}

/// The law over one grid's frames: before each, the rows `writes` gives
/// written; the frame stepped in slices of `budget` cells and its counts
/// held; and a frame that changed nothing with no write before it followed
/// by one more with none before it either, which changes nothing and leaves
/// every value as it was.
fn counts_over(
    m: &mut Machine,
    frames: u64,
    budget: usize,
    writes: &mut dyn FnMut(u64) -> Vec<String>,
    counted: &mut Counted,
    what: &str,
) {
    // Every count the machine answers: the frame's, each sheet's, the cells
    // it computed.
    let all = |m: &Machine| -> Vec<Option<usize>> {
        let sheets = (0..).map_while(|i| m.sheet_name(i)).count();
        let mut v: Vec<Option<usize>> = (0..sheets).map(|s| m.changed_on(s)).collect();
        v.extend([Some(m.changed()), Some(m.reached())]);
        v
    };
    for frame in 1..=frames {
        let rows = writes(frame);
        let kept = all(m);
        for w in &rows {
            m.write(w)
                .unwrap_or_else(|e| panic!("{what}, frame {frame}: {w}: {}", e.refusal.text));
        }
        let when = format!("{what}, frame {frame}");
        assert_eq!(
            all(m),
            kept,
            "{when}: the writes before it leave the counts"
        );
        let before = formula_values(m);
        step_counted(m, budget, counted, &when);
        let after = formula_values(m);
        counts_held(m, &before, &after, &when);
        match m.changed() {
            0 => counted.still += 1,
            _ => counted.changing += 1,
        }
        if rows.is_empty() && m.changed() == 0 {
            step_counted(m, budget, counted, &format!("{when}, settled"));
            assert_eq!(m.changed(), 0, "{when}: the frame after a fixed point");
            let again = formula_values(m);
            assert_eq!(again, after, "{when}: the values after a fixed point");
            counted.settled += 1;
        }
    }
}

#[test]
fn a_frame_s_count_is_the_formula_cells_whose_values_it_changed() {
    // Each family's frames counted apart, so that each is held changing,
    // and still, yielding and settled where its grids do.
    let mut counted: [Counted; 4] = Default::default();
    // Law 9's grids, dressed, written before every second frame as law 7
    // writes them, each stepped whole or in slices.
    for seed in 1..=50u64 {
        let mut r = Rng::new(seed ^ 0x48);
        let rows = grid_rows(&mut r);
        let rows = dressed(&mut r, &rows);
        let mut m = load(&rows, &format!("seed {seed}"));
        let budget = [0, 1, 7, 64][r.below(4)];
        let mut last = None;
        let mut writes = |frame: u64| match frame % 2 {
            0 => vec![write_between(&mut r, &mut last)],
            _ => Vec::new(),
        };
        let what = format!("seed {seed}, budget {budget}\n{rows}");
        counts_over(&mut m, 8, budget, &mut writes, &mut counted[0], &what);
    }
    // Law 10's, most of whose cells change while the cell they read is
    // written before every frame, and which settle or not once it is not.
    for seed in 1..=40u64 {
        let mut r = Rng::new(seed ^ 0x4848);
        let (rows, n) = churning_rows(&mut r);
        let mut m = load(&rows, &format!("churning seed {seed}"));
        let budget = [0, 7, 64][r.below(3)];
        let mut writes = |frame: u64| {
            if frame > 4 {
                return Vec::new();
            }
            let mut w = vec![format!("(cell \"T\" \"A1\" {frame})")];
            if r.below(3) == 0 {
                w.push(format!(
                    "(cell \"G\" \"A{}\" {})",
                    1 + r.below(n),
                    r.below(10)
                ));
            }
            w
        };
        let what = format!("churning seed {seed}, budget {budget}\n{rows}");
        counts_over(&mut m, 8, budget, &mut writes, &mut counted[1], &what);
    }
    // Law 8's fills, one written whole and one a cell each, their value
    // sheet written before the first frames and then not.
    for seed in 1..=20u64 {
        let mut r = Rng::new(seed ^ 0x88);
        let rows = 8 + r.below(17);
        let grid = fill_rows(&mut r, rows);
        for text in [grid.clone(), written(&grid, "cells")] {
            let mut m = load(&text, &format!("fills seed {seed}"));
            let mut writes = |frame: u64| match frame {
                1..=2 => {
                    let cell = format!("{}{}", r.pick(VALUES), 1 + r.below(rows + 2));
                    vec![format!("(cell \"V\" \"{cell}\" {})", r.below(10))]
                }
                _ => Vec::new(),
            };
            let what = format!("fills seed {seed}\n{text}");
            counts_over(&mut m, 6, 0, &mut writes, &mut counted[2], &what);
        }
    }
    // Law 11's values that turn between ones their records spell alike: 0
    // and -0, and a reason and another.
    for seed in 1..=20u64 {
        let mut r = Rng::new(seed ^ 0x5A48);
        let rows = spelled_rows(&mut r);
        let mut m = load(&rows, &format!("spelled seed {seed}"));
        let mut writes = |frame: u64| match frame % 3 {
            0 => {
                let v = r.pick(&["0", "-0", "4", "\"t\""]);
                vec![format!("(cell \"Z\" \"E{}\" {v})", 1 + r.below(5))]
            }
            _ => Vec::new(),
        };
        let what = format!("spelled seed {seed}\n{rows}");
        counts_over(&mut m, 6, 0, &mut writes, &mut counted[3], &what);
    }
    // Measured 2026-10-18, changing, still, yielding and settled: the grids
    // 374, 26, 155 and 16; the churning 234, 86, 256 and 86; the fills 91,
    // 149, none stepped in slices, and 128; the spelled 120 changing, their
    // first cell turning at every frame, so never still.
    let floors = [
        (180, 12, 75, 8),
        (115, 40, 125, 40),
        (45, 70, 0, 60),
        (60, 0, 0, 0),
    ];
    let families = ["grids", "churning", "fills", "spelled"];
    for ((family, c), floor) in families.iter().zip(&counted).zip(floors) {
        assert!(
            c.changing >= floor.0
                && c.still >= floor.1
                && c.yielded >= floor.2
                && c.settled >= floor.3,
            "{family}: too few frames of a kind for the law to hold over: {c:?}"
        );
    }
}

// ---- 13. far-apart cells of one text -------------------------------------

/// An expression written at `C1`: references moving over the value sheet
/// `V` and over the cell's own sheet left of it, a range moving with them,
/// held rectangles and parts, so that every copy of it reads as it does
/// where it is moved, and none reads a cell right of its own.
fn scattered_expr(r: &mut Rng, depth: u32) -> String {
    if depth >= 2 || r.below(3) == 0 {
        return match r.below(6) {
            0 => format!("V!{}{}", r.pick(&["A", "B", "C"]), 1 + r.below(2)),
            1 => format!("V!${}{}", r.pick(&["A", "B"]), 1 + r.below(2)),
            2 => format!("V!$A${}", 1 + r.below(5)),
            3 => format!("{}{}", r.pick(&["A", "B"]), 1 + r.below(2)),
            _ => r.below(10).to_string(),
        };
    }
    let e = |r: &mut Rng| scattered_expr(r, depth + 1);
    match r.below(6) {
        0 => format!("{}+{}", e(r), e(r)),
        1 => format!("{}*{}", e(r), e(r)),
        2 => format!("SUM(V!A1:B2)+{}", e(r)),
        3 => format!("IF(V!$A$1>{},{},{})", r.below(5), e(r), e(r)),
        4 => format!("MOD({},7)", e(r)),
        _ => format!("MAX({},SUM(V!$A$1:$B$2))", e(r)),
    }
}

/// A grid of far-apart cells of one text: a value sheet `V`, and sheets `F`
/// and `G` with values in columns A and B, on which a few texts are each
/// copied to scattered cells of columns C to J over 400 rows, a copy now and
/// then beside the one before it; each copy the text moved to its cell. The
/// grid's rows, and the same grid with each copy given a text of its own
/// (`+0*k`) that computes the same value; and whether a text's copies stand
/// far enough apart for its shape to be cut (KERNEL.54's line: a box more
/// than 64 rows or columns, its cells less than half of it).
fn scattered_rows(r: &mut Rng) -> (String, String, bool) {
    let mut head =
        String::from("(sheet \"F\" visible)\n(sheet \"G\" visible)\n(sheet \"V\" visible)\n");
    for row in 1..=402 {
        for col in ["A", "B", "C"] {
            if r.below(4) != 0 {
                head.push_str(&format!("(cell \"V\" \"{col}{row}\" {})\n", r.below(10)));
            }
        }
    }
    head.push_str("(cell \"F\" \"A1:B402\" 2)\n(cell \"G\" \"A1:B402\" 3)\n");
    let (mut one, mut own) = (head.clone(), head);
    let mut taken: BTreeSet<(&str, u32, u32)> = BTreeSet::new();
    let (mut n, mut cut) = (0, false);
    for _ in 0..2 + r.below(4) {
        let text = scattered_expr(r, 0);
        let mut at: Vec<(&str, u32, u32)> = Vec::new();
        for _ in 0..2 + r.below(9) {
            let cell = match at.last() {
                Some(&(sheet, row, col)) if r.below(4) == 0 => (sheet, row + 1, col),
                _ => (
                    *r.pick(&["F", "G"]),
                    1 + r.below(400) as u32,
                    3 + r.below(8) as u32,
                ),
            };
            if taken.insert(cell) {
                at.push(cell);
            }
        }
        for &(sheet, row, col) in &at {
            let a1 = A1Range {
                top: row,
                left: col,
                bottom: row,
                right: col,
            }
            .text();
            let moved = shift_a1_references(&text, i64::from(row) - 1, i64::from(col) - 3);
            n += 1;
            one.push_str(&format!(
                "(formula \"{sheet}\" \"{a1}\" {})\n",
                quoted(&format!("={moved}"))
            ));
            own.push_str(&format!(
                "(formula \"{sheet}\" \"{a1}\" {})\n",
                quoted(&format!("=({moved})+0*{n}"))
            ));
        }
        let span = |f: fn(&(&str, u32, u32)) -> u32| {
            let (lo, hi) = at
                .iter()
                .fold((u32::MAX, 0), |(lo, hi), c| (lo.min(f(c)), hi.max(f(c))));
            u64::from(hi - lo + 1)
        };
        let (rows, cols) = (span(|c| c.1), span(|c| c.2));
        let sheets: BTreeSet<&str> = at.iter().map(|c| c.0).collect();
        let area = rows * cols * sheets.len() as u64;
        cut |= (rows > 64 || cols > 64) && (at.len() as u64) * 2 < area;
    }
    (one, own, cut)
}

#[test]
fn far_apart_cells_of_one_text_compute_what_each_written_alone_computes() {
    let mut cut = 0;
    for seed in 1..=30u64 {
        let mut r = Rng::new(seed);
        let (one, own, any_cut) = scattered_rows(&mut r);
        cut += usize::from(any_cut);
        let mut pieced = load(&one, &format!("seed {seed}, one text a shape"));
        let mut alone = load(&own, &format!("seed {seed}, a text a copy"));
        for frame in 1..=5 {
            assert!(pieced.step(0).done && alone.step(0).done);
            assert_eq!(
                values(&pieced),
                values(&alone),
                "seed {seed}, frame {frame}\n{one}"
            );
            let write = match r.below(3) {
                0 => format!(
                    "(cell \"V\" \"{}{}\" {})",
                    r.pick(&["A", "B", "C"]),
                    1 + r.below(402),
                    r.below(10)
                ),
                _ => format!(
                    "(cell \"{}\" \"{}{}\" {})",
                    r.pick(&["F", "G"]),
                    r.pick(&["A", "B"]),
                    1 + r.below(402),
                    r.below(10)
                ),
            };
            for m in [&mut pieced, &mut alone] {
                m.write(&write).unwrap_or_else(|e| {
                    panic!("seed {seed}, frame {frame}: {write}: {}", e.refusal.text)
                });
            }
        }
    }
    assert!(
        cut >= 20,
        "{cut} grids of 30 held a text whose shape is cut: too few for the law to hold over"
    );
}
