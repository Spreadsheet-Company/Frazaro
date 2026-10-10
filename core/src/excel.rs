//! Excel's library, registered through the language's functions seam
//! (KERNEL.7, 2026-10-08): the declared subset `frazaro calc` computes, each
//! function with the evidence that admitted it.
//!
//! The rule is `KERNEL.6`'s: a function enters the subset with its fixture
//! and its measured frequency, never by taste. The measure is the one the
//! roadmap cites, the two spreadsheet corpora of Jansen and Hermans, "Enron
//! versus EUSES: A Comparison of Two Spreadsheet Corpora" (2015, arXiv
//! 1503.04055), Table IV: the fifteen most used functions in each corpus,
//! which cover 69% of the Enron corpus's spreadsheets and 70% of EUSES's.
//! Of the Enron fifteen (SUM, IF, AVERAGE, VLOOKUP, ROUND, SUBTOTAL, OFFSET,
//! CONCATENATE, NOW, DAVERAGE, SUMIF, INDEX, MATCH, LOOKUP, MONTH) this
//! slice computes the five that are arithmetic and logic over values; the
//! lookups, the database functions, the date and the text functions are
//! `KERNEL.8`'s, `OFFSET` is unreadable by the reader's own rule, and `NOW`
//! is volatile, which an engine refuses and a door shows as not computed.
//! Of the EUSES fifteen (SUM, IF, ROUND, HYPERLINK, CONCATENATE, AND,
//! COUNTIF, AVERAGE, OR, INDIRECT, MIN, ISNUMBER, MAX, VLOOKUP, ISBLANK) it
//! computes nine. `IFS` enters on its fixture alone,
//! `scripts/reflect/saved.xlsx`, the owner's first Excel-saved workbook,
//! whose cached values are this oracle's first row. The language's own
//! fourteen (`vla_lang::calc::library`) are here because
//! `Alonzo/SPEC.md` section 13 named them for the engine's first
//! cartridges; `IF`, `SUM`, `AND`, `OR`, `MIN` and `MAX` are in the tables
//! too.
//!
//! Every function here is held to the language's exact tests or to its
//! fixture: `scripts/recalc/subset.txt`, built by the door into
//! `subset.xlsx` with one formula per function and per coercion this slice
//! claims. Its cached values are Excel's the day the owner opens it in
//! Excel 365 and saves it beside as `subset_saved.xlsx`, and
//! `tools/check_recalc_golden.ps1` then holds every row of that file to
//! `agree`; until then the door prints `unchecked` beside each, which is
//! the honest word. A function outside this list shows `not computed here`
//! with its name (`HORIZON.md` section 12.6), and the one line that admits
//! it is a row in [`SUBSET`] with its evidence.

use vla_lang::calc::graph::cells_in;
use vla_lang::calc::library::{language, number_arg, numbers};
use vla_lang::calc::value::{round15, text_to_number};
use vla_lang::calc::{Area, Arg, Ctx, ErrorKind, Expr, Library, Reason, Value};
use vla_lang::intrinsics::fold;

/// One function of the subset and why it is in.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Entry {
    pub name: &'static str,
    /// The measurement or the fixture that admitted it.
    pub evidence: &'static str,
}

/// The declared subset, the language's fourteen first, then Excel's.
pub const SUBSET: &[Entry] = &[
    Entry {
        name: "IF",
        evidence: "Enron 2, EUSES 2; SPEC.md s13",
    },
    Entry {
        name: "AND",
        evidence: "EUSES 6; SPEC.md s13",
    },
    Entry {
        name: "OR",
        evidence: "EUSES 9; SPEC.md s13",
    },
    Entry {
        name: "NOT",
        evidence: "SPEC.md s13",
    },
    Entry {
        name: "SUM",
        evidence: "Enron 1, EUSES 1; SPEC.md s13",
    },
    Entry {
        name: "MIN",
        evidence: "EUSES 11; SPEC.md s13",
    },
    Entry {
        name: "MAX",
        evidence: "EUSES 13; SPEC.md s13",
    },
    Entry {
        name: "ABS",
        evidence: "SPEC.md s13",
    },
    Entry {
        name: "INT",
        evidence: "SPEC.md s13",
    },
    Entry {
        name: "MOD",
        evidence: "SPEC.md s13: every counter wraps",
    },
    Entry {
        name: "ROW",
        evidence: "SPEC.md s13: a cell painting at the mouse",
    },
    Entry {
        name: "COLUMN",
        evidence: "SPEC.md s13: a cell painting at the mouse",
    },
    Entry {
        name: "CHOOSE",
        evidence: "SPEC.md s13: a score until INDEX comes",
    },
    Entry {
        name: "SIN",
        evidence: "SPEC.md s13: a slow oscillator",
    },
    Entry {
        name: "AVERAGE",
        evidence: "Enron 3, EUSES 8",
    },
    Entry {
        name: "ROUND",
        evidence: "Enron 5, EUSES 3",
    },
    Entry {
        name: "SUMIF",
        evidence: "Enron 11",
    },
    Entry {
        name: "COUNTIF",
        evidence: "EUSES 7",
    },
    Entry {
        name: "ISNUMBER",
        evidence: "EUSES 12",
    },
    Entry {
        name: "ISBLANK",
        evidence: "EUSES 15",
    },
    Entry {
        name: "IFS",
        evidence: "scripts/reflect/saved.xlsx, the first Excel-saved fixture",
    },
];

/// The functions named by the two corpora's top fifteen that this slice
/// does not compute, each with the slice that takes it: what the honest
/// label names most often on a real workbook.
pub const NAMED_NOT_COMPUTED: &[(&str, &str)] = &[
    ("VLOOKUP", "KERNEL.8, the lookups"),
    ("INDEX", "KERNEL.8, the lookups"),
    ("MATCH", "KERNEL.8, the lookups"),
    ("LOOKUP", "KERNEL.8, the lookups"),
    ("SUBTOTAL", "KERNEL.8, the lookups"),
    ("DAVERAGE", "KERNEL.8, the database functions"),
    ("CONCATENATE", "KERNEL.8, the text functions"),
    ("MONTH", "KERNEL.8, the dates"),
    ("NOW", "volatile: an engine refuses it, a door shows it"),
    ("OFFSET", "unreadable by the reader's rule"),
    ("INDIRECT", "unreadable by the reader's rule"),
    (
        "HYPERLINK",
        "a link, not a value; the formula sink refuses it",
    ),
];

/// Excel's library: the language's functions and the seven above.
pub fn library() -> Library {
    let mut lib = language();
    lib.register_form("IFS", ifs);
    lib.register("AVERAGE", average);
    lib.register("ROUND", round);
    lib.register("SUMIF", sumif);
    lib.register("COUNTIF", countif);
    lib.register("ISNUMBER", isnumber);
    lib.register("ISBLANK", isblank);
    lib
}

/// `IFS(c1, v1, c2, v2, ...)`: the value beside the first condition that
/// holds; a condition's error is the answer; none holding is `#N/A`. Only
/// the value chosen is computed.
fn ifs<'e>(ctx: &Ctx<'_>, args: &'e [Expr]) -> Result<Arg<'e>, Reason> {
    if args.len() < 2 || args.len() % 2 == 1 {
        return Ok(Arg::Scalar(Value::Error(ErrorKind::NA)));
    }
    for pair in args.chunks(2) {
        match ctx.truth(&pair[0])? {
            Ok(true) => return ctx.eval(&pair[1]),
            Ok(false) => {}
            Err(k) => return Ok(Arg::Scalar(Value::Error(k))),
        }
    }
    Ok(Arg::Scalar(Value::Error(ErrorKind::NA)))
}

/// `AVERAGE` over the numbers `SUM` would read; none is `#DIV/0!`.
fn average(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    Ok(match numbers(ctx, args)? {
        Err(k) => Value::Error(k),
        Ok(ns) if ns.is_empty() => Value::Error(ErrorKind::Div0),
        Ok(ns) => Value::Number(ns.iter().sum::<f64>() / ns.len() as f64),
    })
}

/// `ROUND(x, digits)`: half away from zero, the digits read down to a
/// whole number, a negative count rounding to tens and hundreds. The scaled
/// value is first rounded to Excel's fifteen significant digits, so that
/// `ROUND(2.675, 2)` is 2.68 as Excel has it and not the 2.67 the double
/// 267.49999999999997 would give.
fn round(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    if args.len() != 2 {
        return Ok(Value::Error(ErrorKind::Value));
    }
    let x = match number_arg(ctx, args, 0)? {
        Ok(n) => n,
        Err(k) => return Ok(Value::Error(k)),
    };
    let digits = match number_arg(ctx, args, 1)? {
        Ok(n) => n.trunc(),
        Err(k) => return Ok(Value::Error(k)),
    };
    let rounded = if digits >= 0.0 {
        let m = 10f64.powi(digits.min(300.0) as i32);
        round15(x * m).round() / m
    } else {
        let p = 10f64.powi((-digits).min(300.0) as i32);
        round15(x / p).round() * p
    };
    Ok(vla_lang::calc::eval::finite(rounded))
}

/// What a criterion of `SUMIF` or `COUNTIF` asks of a cell.
#[derive(Clone, Debug, PartialEq)]
enum Criterion {
    Eq(Target),
    Ne(Target),
    Lt(Target),
    Gt(Target),
    Le(Target),
    Ge(Target),
}

#[derive(Clone, Debug, PartialEq)]
enum Target {
    Number(f64),
    /// A text with `*`, `?` and `~` as Excel's wildcards, folded.
    Text(String),
    Bool(bool),
    Error(ErrorKind),
    Blank,
}

/// A criterion as Excel reads one: a number, a truth value or an error
/// matches cells equal to it; a text may begin with `=`, `<>`, `<`, `>`,
/// `<=` or `>=`, and what follows is a number when it spells one, `TRUE` or
/// `FALSE` when it spells either, blank when empty, a text pattern
/// otherwise.
fn criterion(v: &Value) -> Criterion {
    let text = match v {
        Value::Number(n) => return Criterion::Eq(Target::Number(*n)),
        Value::Bool(b) => return Criterion::Eq(Target::Bool(*b)),
        Value::Error(k) => return Criterion::Eq(Target::Error(*k)),
        Value::Empty => return Criterion::Eq(Target::Blank),
        Value::Text(t) => t.as_str(),
    };
    let (op, rest): (fn(Target) -> Criterion, &str) = if let Some(r) = text.strip_prefix("<>") {
        (Criterion::Ne, r)
    } else if let Some(r) = text.strip_prefix("<=") {
        (Criterion::Le, r)
    } else if let Some(r) = text.strip_prefix(">=") {
        (Criterion::Ge, r)
    } else if let Some(r) = text.strip_prefix('<') {
        (Criterion::Lt, r)
    } else if let Some(r) = text.strip_prefix('>') {
        (Criterion::Gt, r)
    } else if let Some(r) = text.strip_prefix('=') {
        (Criterion::Eq, r)
    } else {
        (Criterion::Eq, text)
    };
    let target = if rest.is_empty() {
        Target::Blank
    } else if let Some(n) = text_to_number(rest) {
        Target::Number(n)
    } else if let Some(k) = ErrorKind::parse(rest) {
        Target::Error(k)
    } else {
        match fold(rest).as_str() {
            "true" => Target::Bool(true),
            "false" => Target::Bool(false),
            _ => Target::Text(fold(rest)),
        }
    };
    op(target)
}

/// Whether a cell's value meets the criterion. A text that spells a number
/// counts as the number for equality, as `COUNTIF` counts it; a comparison
/// by size holds only between values of one kind; `<>` holds for every
/// cell that `=` would not count, an empty cell included.
fn meets(cell: &Value, c: &Criterion) -> bool {
    use std::cmp::Ordering::*;
    let ordering = |target: &Target| -> Option<std::cmp::Ordering> {
        match (cell, target) {
            (Value::Number(x), Target::Number(n)) => x.partial_cmp(n),
            (Value::Text(t), Target::Number(n)) => {
                text_to_number(t).and_then(|x| (x == *n).then_some(Equal))
            }
            (Value::Text(t), Target::Text(pattern)) => {
                if wildcard(pattern, &fold(t)) {
                    Some(Equal)
                } else {
                    Some(fold(t).cmp(pattern))
                }
            }
            (Value::Bool(b), Target::Bool(want)) => Some(b.cmp(want)),
            (Value::Error(k), Target::Error(want)) => (k == want).then_some(Equal),
            (Value::Empty, Target::Blank) => Some(Equal),
            (Value::Text(t), Target::Blank) if t.is_empty() => Some(Equal),
            _ => None,
        }
    };
    match c {
        Criterion::Eq(t) => ordering(t) == Some(Equal),
        Criterion::Ne(t) => ordering(t) != Some(Equal),
        Criterion::Lt(t) => matches!(
            (ordering(t), t),
            (Some(Less), Target::Number(_) | Target::Text(_))
        ),
        Criterion::Gt(t) => matches!(
            (ordering(t), t),
            (Some(Greater), Target::Number(_) | Target::Text(_))
        ),
        Criterion::Le(t) => matches!(
            (ordering(t), t),
            (
                Some(Less) | Some(Equal),
                Target::Number(_) | Target::Text(_)
            )
        ),
        Criterion::Ge(t) => matches!(
            (ordering(t), t),
            (
                Some(Greater) | Some(Equal),
                Target::Number(_) | Target::Text(_)
            )
        ),
    }
}

/// Excel's wildcards over folded text: `*` any run, `?` one character, `~`
/// the next character itself.
fn wildcard(pattern: &str, text: &str) -> bool {
    fn go(p: &[char], t: &[char]) -> bool {
        match p.split_first() {
            None => t.is_empty(),
            Some(('*', rest)) => (0..=t.len()).any(|i| go(rest, &t[i..])),
            Some(('?', rest)) => !t.is_empty() && go(rest, &t[1..]),
            Some(('~', rest)) => match (rest.split_first(), t.split_first()) {
                (Some((want, rest2)), Some((have, t2))) => want == have && go(rest2, t2),
                (None, _) => t.is_empty(),
                _ => false,
            },
            Some((want, rest)) => match t.split_first() {
                Some((have, t2)) => want == have && go(rest, t2),
                None => false,
            },
        }
    }
    let p: Vec<char> = pattern.chars().collect();
    let t: Vec<char> = text.chars().collect();
    if !p.iter().any(|c| matches!(c, '*' | '?' | '~')) {
        return p == t;
    }
    go(&p, &t)
}

/// One area from an argument, for the range of `SUMIF` and `COUNTIF`.
fn one_area(arg: Option<&Arg>) -> Option<Area> {
    match arg {
        Some(Arg::Area { areas, .. }) if areas.len() == 1 => Some(areas[0]),
        _ => None,
    }
}

/// `COUNTIF(range, criterion)`: the cells of the range meeting the
/// criterion; a criterion a blank meets counts the empty cells of the whole
/// range, as Excel counts them, however far it reaches.
fn countif(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    if args.len() != 2 {
        return Ok(Value::Error(ErrorKind::Value));
    }
    let Some(area) = one_area(args.first()) else {
        return Ok(Value::Error(ErrorKind::Value));
    };
    let c = criterion(&ctx.one(&args[1])?);
    let model = ctx.env.model();
    let Some(sheet) = model.sheets.get(area.sheet) else {
        return Ok(Value::Error(ErrorKind::Ref));
    };
    let mut count = 0u64;
    let mut present = 0u64;
    for ((row, col), _) in cells_in(sheet, ctx.env.extent(area.sheet), &area) {
        let v = ctx.env.read((area.sheet, row, col))?;
        if v != Value::Empty {
            present += 1;
            if meets(&v, &c) {
                count += 1;
            }
        }
    }
    let cells = u64::from(area.bottom - area.top + 1) * u64::from(area.right - area.left + 1);
    if meets(&Value::Empty, &c) {
        count += cells - present;
    }
    Ok(Value::Number(count as f64))
}

/// `SUMIF(range, criterion, [sum_range])`: the numbers of the sum range
/// (the range itself when none is given) at the places where the range's
/// cell meets the criterion; the sum range is read from its top-left cell
/// in the range's shape, as Excel reads it.
fn sumif(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    if args.len() < 2 || args.len() > 3 {
        return Ok(Value::Error(ErrorKind::Value));
    }
    let Some(area) = one_area(args.first()) else {
        return Ok(Value::Error(ErrorKind::Value));
    };
    let c = criterion(&ctx.one(&args[1])?);
    let sum_area = match args.get(2) {
        None => area,
        Some(a) => match one_area(Some(a)) {
            Some(s) => Area {
                sheet: s.sheet,
                top: s.top,
                left: s.left,
                bottom: s.top.saturating_add(area.bottom - area.top),
                right: s.left.saturating_add(area.right - area.left),
            },
            None => return Ok(Value::Error(ErrorKind::Value)),
        },
    };
    let model = ctx.env.model();
    let Some(sum_sheet) = model.sheets.get(sum_area.sheet) else {
        return Ok(Value::Error(ErrorKind::Ref));
    };
    // Every contribution is a present cell of the sum range; the criterion
    // is read at the same offset in the range, an absent cell being empty.
    let mut total = 0.0;
    for ((row, col), _) in cells_in(sum_sheet, ctx.env.extent(sum_area.sheet), &sum_area) {
        let value = ctx.env.read((sum_area.sheet, row, col))?;
        let n = match value {
            Value::Number(n) => n,
            Value::Error(k) => return Ok(Value::Error(k)),
            _ => continue,
        };
        let at = (
            area.sheet,
            area.top + (row - sum_area.top),
            area.left + (col - sum_area.left),
        );
        let tested = ctx.env.read(at)?;
        if meets(&tested, &c) {
            total += n;
        }
    }
    Ok(vla_lang::calc::eval::finite(total))
}

/// `ISNUMBER(x)`: whether the value is a number; an error is not, and does
/// not propagate.
fn isnumber(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    if args.len() != 1 {
        return Ok(Value::Error(ErrorKind::Value));
    }
    Ok(Value::Bool(matches!(ctx.one(&args[0])?, Value::Number(_))))
}

/// `ISBLANK(x)`: whether the cell is empty; a formula's empty result is 0
/// and so is not.
fn isblank(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    if args.len() != 1 {
        return Ok(Value::Error(ErrorKind::Value));
    }
    Ok(Value::Bool(ctx.one(&args[0])? == Value::Empty))
}

/// The names the subset holds, for a page to list: the language's first.
pub fn subset_names() -> Vec<&'static str> {
    SUBSET.iter().map(|e| e.name).collect()
}

#[cfg(test)]
mod tests {
    use super::*;
    use vla_lang::calc::library::NAMES as LANGUAGE_NAMES;
    use vla_lang::calc::{Calc, Computed};
    use vla_lang::sheet::{parse_a1_range, Cell, Content, Workbook};

    fn model(rows: &[(&str, &str, &str)]) -> Workbook {
        let mut wb = Workbook::new();
        for (sheet, addr, text) in rows {
            let i = wb.ensure_sheet(sheet);
            let r = parse_a1_range(addr).unwrap();
            let content = if let Some(f) = text.strip_prefix('=') {
                Content::Formula(f.to_string())
            } else if let Ok(n) = text.parse::<f64>() {
                Content::Number(n)
            } else if *text == "TRUE" || *text == "FALSE" {
                Content::Bool(*text == "TRUE")
            } else {
                Content::Text(text.to_string())
            };
            wb.sheets[i].set(r.top, r.left, Cell { content, style: 0 });
        }
        wb
    }

    fn at(calc: &Calc<'_>, addr: &str) -> Computed {
        let r = parse_a1_range(addr).unwrap();
        calc.computed((0, r.top, r.left)).cloned().unwrap()
    }

    fn number(calc: &Calc<'_>, addr: &str) -> f64 {
        match at(calc, addr) {
            Computed::Value(Value::Number(n)) => n,
            other => panic!("{addr}: {other:?}"),
        }
    }

    #[test]
    fn the_subset_table_is_the_library_and_the_language_comes_first() {
        let lib = library();
        let mut names: Vec<&str> = SUBSET.iter().map(|e| e.name).collect();
        names.sort_unstable();
        assert_eq!(
            lib.names(),
            names,
            "every function admitted has its evidence row, and no other"
        );
        assert_eq!(
            &SUBSET[..LANGUAGE_NAMES.len()]
                .iter()
                .map(|e| e.name)
                .collect::<Vec<_>>(),
            &LANGUAGE_NAMES.to_vec()
        );
        for e in SUBSET {
            assert!(!e.evidence.is_empty(), "{} has no evidence", e.name);
        }
        for (name, _) in NAMED_NOT_COMPUTED {
            assert!(!lib.has(name), "{name} is listed as not computed");
        }
    }

    #[test]
    fn excel_s_functions_compute_as_excel_does() {
        let wb = model(&[
            ("S", "A1", "10"),
            ("S", "A2", "20"),
            ("S", "A3", "x"),
            ("S", "A4", "30"),
            ("S", "B1", "1"),
            ("S", "B2", "2"),
            ("S", "B4", "4"),
            ("S", "C1", "=AVERAGE(A1:A4)"),
            ("S", "C2", "=AVERAGE(A3)"),
            ("S", "C3", "=ROUND(2.675,2)"),
            ("S", "C4", "=ROUND(-2.5,0)"),
            ("S", "C5", "=ROUND(1234.5678,-2)"),
            ("S", "C6", "=ROUND(1.234,1.9)"),
            ("S", "C7", "=SUMIF(A1:A4,\">15\")"),
            ("S", "C8", "=SUMIF(A1:A4,\">15\",B1:B4)"),
            ("S", "C9", "=COUNTIF(A1:A4,\">15\")"),
            ("S", "C10", "=COUNTIF(A1:A4,\"x\")"),
            ("S", "C11", "=COUNTIF(A1:A6,\"\")"),
            ("S", "C12", "=COUNTIF(A1:A4,\"<>10\")"),
            ("S", "C13", "=COUNTIF(A1:A4,\"*\")"),
            ("S", "C14", "=COUNTIF(A1:A4,10)"),
            ("S", "C15", "=ISNUMBER(A1)"),
            ("S", "C16", "=ISNUMBER(A3)"),
            ("S", "C17", "=ISNUMBER(1/0)"),
            ("S", "C18", "=ISBLANK(A5)"),
            ("S", "C19", "=ISBLANK(A1)"),
            ("S", "C20", "=IFS(A1>15,\"big\",TRUE,\"small\")"),
            ("S", "C21", "=IFS(A1>15,\"big\")"),
            ("S", "C22", "=IFS(A3,1,TRUE,2)"),
            ("S", "C23", "=COUNTIF(B1:B4,\"<>\")"),
            ("S", "C24", "=COUNTIF(A1:A4,\"?\")"),
            ("S", "C25", "=SUMIF(A1:A4,\"x\",B1:B4)"),
        ]);
        let lib = library();
        let calc = Calc::run(&wb, &lib);
        assert_eq!(number(&calc, "C1"), 20.0, "the text is skipped");
        assert_eq!(
            at(&calc, "C2"),
            Computed::Value(Value::Error(ErrorKind::Div0))
        );
        assert_eq!(number(&calc, "C3"), 2.68);
        assert_eq!(number(&calc, "C4"), -3.0);
        assert_eq!(number(&calc, "C5"), 1200.0);
        assert_eq!(number(&calc, "C6"), 1.2);
        assert_eq!(number(&calc, "C7"), 50.0);
        assert_eq!(number(&calc, "C8"), 6.0);
        assert_eq!(number(&calc, "C9"), 2.0);
        assert_eq!(number(&calc, "C10"), 1.0);
        assert_eq!(number(&calc, "C11"), 2.0, "A5 and A6 are blank");
        assert_eq!(number(&calc, "C12"), 3.0, "the text counts as not 10");
        assert_eq!(number(&calc, "C13"), 1.0, "* matches texts alone");
        assert_eq!(number(&calc, "C14"), 1.0);
        assert_eq!(at(&calc, "C15"), Computed::Value(Value::Bool(true)));
        assert_eq!(at(&calc, "C16"), Computed::Value(Value::Bool(false)));
        assert_eq!(
            at(&calc, "C17"),
            Computed::Value(Value::Bool(false)),
            "an error is not a number and does not propagate"
        );
        assert_eq!(at(&calc, "C18"), Computed::Value(Value::Bool(true)));
        assert_eq!(at(&calc, "C19"), Computed::Value(Value::Bool(false)));
        assert_eq!(
            at(&calc, "C20"),
            Computed::Value(Value::Text("small".to_string()))
        );
        assert_eq!(
            at(&calc, "C21"),
            Computed::Value(Value::Error(ErrorKind::NA))
        );
        assert_eq!(
            at(&calc, "C22"),
            Computed::Value(Value::Error(ErrorKind::Value)),
            "a text condition"
        );
        assert_eq!(number(&calc, "C23"), 3.0, "the non-blank cells");
        assert_eq!(number(&calc, "C24"), 1.0);
        assert_eq!(number(&calc, "C25"), 0.0, "B3 is absent");
    }

    #[test]
    fn criteria_read_as_excel_reads_them() {
        assert_eq!(
            criterion(&Value::Text(">=5".into())),
            Criterion::Ge(Target::Number(5.0))
        );
        assert_eq!(
            criterion(&Value::Text("<>".into())),
            Criterion::Ne(Target::Blank)
        );
        assert_eq!(
            criterion(&Value::Text("=TRUE".into())),
            Criterion::Eq(Target::Bool(true))
        );
        assert_eq!(
            criterion(&Value::Text("a*".into())),
            Criterion::Eq(Target::Text("a*".into()))
        );
        assert_eq!(
            criterion(&Value::Number(3.0)),
            Criterion::Eq(Target::Number(3.0))
        );
        assert!(wildcard("a*", "abc"));
        assert!(wildcard("a?c", "abc"));
        assert!(!wildcard("a?c", "abbc"));
        assert!(wildcard("10~*", "10*"));
        assert!(wildcard("*", ""));
        assert!(meets(
            &Value::Text("5".into()),
            &Criterion::Eq(Target::Number(5.0))
        ));
        assert!(
            !meets(
                &Value::Text("5".into()),
                &Criterion::Gt(Target::Number(1.0))
            ),
            "a text is not compared by size against a number"
        );
        assert!(meets(&Value::Empty, &Criterion::Ne(Target::Number(5.0))));
        assert!(meets(
            &Value::Text("ABC".into()),
            &Criterion::Eq(Target::Text("abc".into()))
        ));
    }
}
