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
//! CONCATENATE, NOW, DAVERAGE, SUMIF, INDEX, MATCH, LOOKUP, MONTH) `KERNEL.7`
//! computed the five that are arithmetic and logic over values, and
//! `KERNEL.8`'s first slice the four lookups (2026-10-18); the database
//! functions, the dates and the text functions are its later slices',
//! `OFFSET` is unreadable by the reader's own rule, and `NOW` is volatile,
//! which an engine refuses and a door shows as not computed. Of the EUSES
//! fifteen (SUM, IF, ROUND, HYPERLINK, CONCATENATE, AND, COUNTIF, AVERAGE,
//! OR, INDIRECT, MIN, ISNUMBER, MAX, VLOOKUP, ISBLANK) it computes ten. `IFS` enters on its fixture alone,
//! `scripts/reflect/saved.xlsx`, the owner's first Excel-saved workbook,
//! whose cached values are this oracle's first row. The language's own
//! functions (`vla_lang::calc::library`) are here as they are there: its
//! day-one fourteen, which `Alonzo/SPEC.md` section 13 named for the
//! engine's first cartridges (`IF`, `SUM`, `AND`, `OR`, `MIN` and `MAX` in
//! the tables too), and the thirteen `KERNEL.8` brought, the lookups, the
//! error tests a lookup's miss is read by, and `SUMIF` and `COUNTIF`, moved
//! down from this file with their criteria since a vector's index keys by
//! them and a cartridge counts duplicates (Alonzo's `CART.13`).
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

use vla_lang::calc::library::{language, number_arg, numbers};
use vla_lang::calc::value::round15;
use vla_lang::calc::{Arg, Ctx, ErrorKind, Expr, Library, Reason, Value};

/// One function of the subset and why it is in.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Entry {
    pub name: &'static str,
    /// The measurement or the fixture that admitted it.
    pub evidence: &'static str,
}

/// The declared subset: the language's functions first, its day-one
/// fourteen and the thirteen `KERNEL.8` brought, then Excel's.
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
        name: "INDEX",
        evidence: "Enron 12; Alonzo's KIT.3 and CART.5, Buffet's MISE.6; scripts/recalc/lookups.txt",
    },
    Entry {
        name: "MATCH",
        evidence: "Enron 13; scripts/recalc/lookups.txt",
    },
    Entry {
        name: "XMATCH",
        evidence: "Excel 365's spelling of MATCH's search; scripts/recalc/lookups.txt",
    },
    Entry {
        name: "VLOOKUP",
        evidence: "Enron 4, EUSES 14; scripts/recalc/lookups.txt",
    },
    Entry {
        name: "HLOOKUP",
        evidence: "VLOOKUP's search across a row; scripts/recalc/lookups.txt",
    },
    Entry {
        name: "LOOKUP",
        evidence: "Enron 14; scripts/recalc/lookups.txt",
    },
    Entry {
        name: "XLOOKUP",
        evidence: "Excel 365's spelling of the lookups; Alonzo's ALGEBRA.md s12, the word game; scripts/recalc/lookups.txt",
    },
    Entry {
        name: "IFERROR",
        evidence: "a lookup's miss, read; scripts/recalc/lookups.txt",
    },
    Entry {
        name: "IFNA",
        evidence: "a lookup's miss, read; scripts/recalc/lookups.txt",
    },
    Entry {
        name: "ISERROR",
        evidence: "a lookup's miss, read; scripts/recalc/lookups.txt",
    },
    Entry {
        name: "ISNA",
        evidence: "a lookup's miss, read; scripts/recalc/lookups.txt",
    },
    Entry {
        name: "COUNTIF",
        evidence: "EUSES 7; moved into the language, KERNEL.8",
    },
    Entry {
        name: "SUMIF",
        evidence: "Enron 11; moved into the language, KERNEL.8",
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

/// The functions named by the two corpora's top fifteen that the subset
/// does not compute, each with the slice that takes it: what the honest
/// label names most often on a real workbook.
pub const NAMED_NOT_COMPUTED: &[(&str, &str)] = &[
    ("SUBTOTAL", "KERNEL.8, the aggregates by criteria"),
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

/// Excel's library: the language's functions and the five of its own below.
pub fn library() -> Library {
    let mut lib = language();
    lib.register_form("IFS", ifs);
    lib.register("AVERAGE", average);
    lib.register("ROUND", round);
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
}
