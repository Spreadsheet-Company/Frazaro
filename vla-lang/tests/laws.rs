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
//!    before a frame and after each one.
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

use std::collections::HashMap;

use vla_lang::calc::library;
use vla_lang::expand::expand_text;
use vla_lang::form::Form;
use vla_lang::machine::{Draft, Machine, Viewed};
use vla_lang::printer::write_datum;
use vla_lang::reader::read_forms;
use vla_lang::refers::shift_a1_references;
use vla_lang::rows::quoted;
use vla_lang::sheet::{parse_a1_range, A1Range};

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
/// language library's functions, three levels deep at most. `row` is the
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
    match r.below(10) {
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
    for seed in 1..=50u64 {
        let mut r = Rng::new(seed);
        let rows = grid_rows(&mut r);
        let mut m = load(&rows, &format!("seed {seed}"));
        for frame in 0..=2 {
            let saved = save(&m);
            let again = load(&saved, &format!("seed {seed}, frame {frame}'s save"));
            assert_eq!(save(&again), saved, "seed {seed}, frame {frame}\n{rows}");
            assert!(m.step(0).done);
        }
    }
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
            let mut workbook = m.workbook().clone();
            workbook.sheets.truncate(sheets);
            let cells = workbook.sheets.iter().map(|s| s.cells.len()).sum();
            let draft = Draft {
                workbook,
                values: HashMap::new(),
                lines: HashMap::new(),
                cells,
            };
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
/// not the first row, under the language library's functions.
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
    match r.below(6) {
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
