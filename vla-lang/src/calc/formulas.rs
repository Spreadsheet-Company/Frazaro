//! A formula read by a formula (KERNEL.53, 2026-10-21): `ISFORMULA`,
//! whether the cell a reference names holds a formula, and `FORMULATEXT`,
//! that formula's text as Excel's formula bar shows it, each a function of
//! a reference and never of a value, Microsoft's page on each its source.
//!
//! The text is the one the view record prints for the cell
//! ([`crate::sheet::formula_text`], an `=` before it): the formula as its
//! file or its row writes it, the file's prefixes dropped and a fill's cell
//! its master's text moved to it, which for a workbook Excel saved is the
//! spelling Excel saved. Where Excel's bar spells a formula otherwise the
//! text is labelled and never guessed: a construct the language does not
//! read, by its own name, which is how the parser names `@`, `A1#`, a
//! table's reference and another workbook's; `SINGLE` and `ANCHORARRAY`,
//! the file's names for `@` and `#`; and a legacy formula Excel 365 may show
//! with an `@` ([`may_show_at`], the owner's decision of 2026-10-21). A
//! legacy array formula stands in the braces the bar shows it in, and a
//! text longer than the 8,192 characters Excel holds, its `=` aside, is
//! `#N/A`, as Excel answers.
//!
//! A formula changes only by a write, and a write that changes one has the
//! next frame compute every cell (`crate::machine`), so that a formula
//! rewritten to the value it had computes its readers again. The plan reads
//! the reference as it reads every reference, an edge, so that a cell
//! reading its own formula is a cycle, as a `ROW` of its own cell is, where
//! Excel computes it (the owner's decision of 2026-10-21).
//!
//! A reference of many cells: its first cell's answer where one value is
//! wanted, as Microsoft's page on `FORMULATEXT` has it, and each of its
//! cells' where the expression is read as arrays, as `ROW` and `COLUMN`
//! answer since KERNEL.8's sixth slice and as `SUMPRODUCT(--ISFORMULA(r))`
//! counts a range's formulas.

use super::array::{self, Array};
use super::eval::{Arg, Ctx, Param, Reason, Written};
use super::formula::{self, Expr};
use super::graph::CellId;
use super::value::{ErrorKind, Value};

/// How deep a defined name may read a name before [`may_show_at`] gives up
/// and says an `@` may show.
const NAME_DEPTH: u32 = 8;

/// The label of a formula whose text Excel 365 may show with an `@`.
const SHOWN_WITH_AT: &str = "a formula Excel shows with @";

/// `ISFORMULA(reference)`: whether the cell holds a formula.
pub(crate) fn isformula(ctx: &Ctx<'_>, args: &[Arg<'_>]) -> Result<Arg<'static>, Reason> {
    of_cells(ctx, args, Value::Bool(false), |_, _, _| {
        Ok(Value::Bool(true))
    })
}

/// `FORMULATEXT(reference)`: the formula's text as Excel's bar shows it,
/// `#N/A` for a cell holding none.
pub(crate) fn formulatext(ctx: &Ctx<'_>, args: &[Arg<'_>]) -> Result<Arg<'static>, Reason> {
    of_cells(ctx, args, Value::Error(ErrorKind::NA), text_at)
}

/// A function of the cells a reference names, `at` each cell holding a
/// formula and `none` at every other place, which for both functions is a
/// value cell's answer too (KERNEL.49: the grid answers a cell's formula, and
/// a cell holding none is as no cell): the error a reference to no cell
/// reaches it as (`#REF!`); `#VALUE!` for a value or an array, which name
/// no cell, as for a call of more arguments or fewer, and for a reference
/// over a span of sheets, as Excel 365's save of the fixture answers both
/// functions there; a reference's first cell's answer where one value is
/// wanted; and where the expression is read as arrays, an array of each
/// place's answer.
fn of_cells(
    ctx: &Ctx<'_>,
    args: &[Arg<'_>],
    none: Value,
    at: fn(&Ctx<'_>, CellId, &Written<'_>) -> Result<Value, Reason>,
) -> Result<Arg<'static>, Reason> {
    let areas = match args {
        [Arg::Area { areas, .. }] => areas,
        [Arg::Scalar(Value::Error(k))] => return Ok(Arg::Scalar(Value::Error(*k))),
        _ => return Ok(Arg::Scalar(Value::Error(ErrorKind::Value))),
    };
    let a = match &areas[..] {
        [a] => *a,
        [] => return Ok(Arg::Scalar(Value::Error(ErrorKind::Ref))),
        _ => return Ok(Arg::Scalar(Value::Error(ErrorKind::Value))),
    };
    if a.sheet >= ctx.env.sheets() {
        return Ok(Arg::Scalar(Value::Error(ErrorKind::Ref)));
    }
    if a.is_cell() || !ctx.arrays() {
        let id = (a.sheet, a.top, a.left);
        return Ok(Arg::Scalar(match ctx.env.written(id) {
            Some(w) => at(ctx, id, &w)?,
            None => none,
        }));
    }
    let rows = u64::from(a.bottom - a.top) + 1;
    let columns = u64::from(a.right - a.left) + 1;
    let (rows, columns) = match array::room(rows, columns)? {
        Ok(rc) => rc,
        Err(k) => return Ok(Arg::Scalar(Value::Error(k))),
    };
    // Every place `none`, then each cell the sheet holds inside the area,
    // so that a whole column is walked by its cells and not its places.
    let width = columns as usize;
    let mut values = vec![none; rows as usize * width];
    for (row, col) in ctx.env.walk(&a) {
        let id = (a.sheet, row, col);
        if let Some(w) = ctx.env.written(id) {
            values[(row - a.top) as usize * width + (col - a.left) as usize] = at(ctx, id, &w)?;
        }
    }
    Ok(Arg::Array(Array::new(rows, columns, values)))
}

/// `FORMULATEXT`'s answer at one cell: `#N/A` for a cell holding no formula,
/// for a fill's cell whose master the sheet has lost, for which the record
/// shows none either, and for a text longer than Excel's 8,192 characters,
/// its `=` aside, as the parser counts them (`formula::LENGTH_LIMIT`); else
/// the record's text with its `=`, in braces for a legacy array formula,
/// labelled where Excel's bar spells it otherwise ([`spelled_alike`]).
fn text_at(ctx: &Ctx<'_>, (sheet, _, _): CellId, written: &Written<'_>) -> Result<Value, Reason> {
    let Some(text) = written.text() else {
        return Ok(Value::Error(ErrorKind::NA));
    };
    if text.chars().count() > formula::LENGTH_LIMIT {
        return Ok(Value::Error(ErrorKind::NA));
    }
    spelled_alike(ctx, written, sheet, &text)?;
    Ok(Value::Text(match written.array() {
        true => format!("{{={text}}}"),
        false => format!("={text}"),
    }))
}

/// Whether Excel's bar spells a formula as its file does, the label where it
/// may not: a construct the language does not read, by the name the parser
/// gives it; `SINGLE` or `ANCHORARRAY` anywhere, which the bar shows as `@`
/// and `#`; and in a legacy formula, an `@` Excel 365 may show there.
fn spelled_alike(
    ctx: &Ctx<'_>,
    written: &Written<'_>,
    home: usize,
    text: &str,
) -> Result<(), Reason> {
    let e = formula::parse(text).map_err(Reason::Construct)?;
    let calls = formula::calls(text);
    if calls.iter().any(|name| name == "SINGLE") {
        return Err(Reason::Construct(SHOWN_WITH_AT.to_string()));
    }
    if calls.iter().any(|name| name == "ANCHORARRAY") {
        return Err(Reason::Construct(
            "a formula Excel shows with #".to_string(),
        ));
    }
    if written.legacy() && may_show_at(ctx, &e, home, 0) {
        return Err(Reason::Construct(SHOWN_WITH_AT.to_string()));
    }
    Ok(())
}

/// How a legacy formula reads an expression, for [`may_show_at`]: one value
/// wanted, where a range is intersected with the formula's cell; handed
/// whole to a parameter that takes a reference (`SUM`'s), where a reference
/// is not, though an operator inside still reads one value of each operand;
/// or inside a parameter that takes an array (`SUMPRODUCT`'s), where
/// nothing is but what a function forces to one value (KERNEL.8's sixth
/// slice: `IFERROR`'s and `IFNA`'s arguments, a lookup's but its table).
#[derive(Clone, Copy, PartialEq, Eq)]
enum Read {
    One,
    Whole,
    Arrays,
}

impl Read {
    /// How an operator's operand, or an argument of one value, is read here.
    fn value(self) -> Read {
        match self {
            Read::Arrays => Read::Arrays,
            _ => Read::One,
        }
    }

    /// How an argument of this class is read here.
    fn of(self, class: Param) -> Read {
        match class {
            Param::Value => self.value(),
            Param::Scalar => Read::One,
            Param::Range if self == Read::Arrays => Read::Arrays,
            Param::Range => Read::Whole,
            _ => Read::Arrays,
        }
    }
}

/// Whether Excel 365 may show an `@` in a legacy formula, which Microsoft's
/// page "Implicit intersection operator: @" says it adds where implicit
/// intersection could occur: before a range, and before a function that
/// could return a range or an array, unless what reads it takes a range or
/// an array. So where one value is wanted ([`Read::One`]: at the formula's
/// top, at an operator's operands, at a parameter of one value): a
/// reference of more than one cell, `:` between expressions, a defined name
/// that may name more, or a call that may answer a reference of many cells,
/// `INDEX`, `XLOOKUP` or a function the library does not hold, as Excel
/// shows one of an add-in. And anywhere but inside a parameter read as
/// arrays, a call that answers an array where a dynamic formula reads it
/// and one value where a legacy formula does, the page's exception for a
/// range's parameter being a reference's: a function answering an operand,
/// or `ROW`, `COLUMN`, `ISFORMULA` or `FORMULATEXT` of more than one cell.
/// The branches of `IF`, `CHOOSE` and `IFS` are read as their call is. It
/// may say an `@` may show where Excel shows none, and the text is then
/// labelled where it would have computed; Excel's save of the fixture tells.
fn may_show_at(ctx: &Ctx<'_>, e: &Expr, home: usize, depth: u32) -> bool {
    // An explicit stack, so that a chain of 4,096 operators walks in one
    // frame, as the evaluator walks it (VLA.1).
    let mut stack: Vec<(&Expr, Read)> = vec![(e, Read::One)];
    while let Some((e, read)) = stack.pop() {
        let one = read == Read::One;
        let shows = match e {
            Expr::Ref(r) => one && !r.is_cell(),
            Expr::Range(_) => one,
            Expr::Name { name, sheet } => {
                one && !names_one(ctx, name, sheet.as_deref(), home, depth)
            }
            Expr::Neg(x) | Expr::Pos(x) | Expr::Percent(x) => {
                stack.push((x, read.value()));
                false
            }
            Expr::Binary(_, l, r) => {
                stack.push((l, read.value()));
                stack.push((r, read.value()));
                false
            }
            Expr::Call(name, args) => {
                call_may_show_at(ctx, name, args, read, home, depth, &mut stack)
            }
            _ => false,
        };
        if shows {
            return true;
        }
    }
    false
}

/// A call's part of [`may_show_at`]: whether the call itself may answer
/// many where one value is wanted, and each argument pushed with how it is
/// read.
fn call_may_show_at<'e>(
    ctx: &Ctx<'_>,
    name: &str,
    args: &'e [Expr],
    read: Read,
    home: usize,
    depth: u32,
    stack: &mut Vec<(&'e Expr, Read)>,
) -> bool {
    let lib = ctx.lib;
    let one = read == Read::One;
    let each = |stack: &mut Vec<(&'e Expr, Read)>, how: &dyn Fn(usize) -> Read| {
        for (j, a) in args.iter().enumerate() {
            stack.push((a, how(j)));
        }
    };
    match name {
        // A branch answers what the call answers.
        "IF" | "CHOOSE" => each(stack, &|j| if j == 0 { read.value() } else { read }),
        "IFS" => each(stack, &|j| if j % 2 == 0 { read.value() } else { read }),
        "IFERROR" | "IFNA" => each(stack, &|_| Read::One),
        "INDEX" | "XLOOKUP" => {
            if one {
                return true;
            }
            let arrays: &[usize] = if name == "INDEX" { &[0] } else { &[1, 2] };
            each(stack, &|j| {
                if arrays.contains(&j) {
                    Read::Arrays
                } else {
                    read.value()
                }
            });
        }
        // A reference read for its place or its formula, never its values,
        // which answers an array of many cells where read as arrays.
        "ROW" | "COLUMN" | "ISFORMULA" | "FORMULATEXT" => {
            if let [a] = args {
                if read != Read::Arrays && !one_cell(ctx, a, home, depth) {
                    return true;
                }
            }
            each(stack, &|_| Read::Whole);
        }
        _ if !lib.has(name) => {
            if one {
                return true;
            }
            each(stack, &|_| read.value());
        }
        _ if lib.arrayed(name).is_some() => {
            if read != Read::Arrays {
                return true;
            }
            let classes = lib.classes(name);
            each(stack, &|j| classes.map_or(Read::Arrays, |c| read.of(c(j))));
        }
        _ => match lib.classes(name) {
            Some(classes) => each(stack, &|j| read.of(classes(j))),
            // A form of a crate above, whose parameters say nothing.
            None => each(stack, &|_| read.value()),
        },
    }
    false
}

/// Whether an argument names one cell for certain: a reference to one cell,
/// or a defined name of one.
fn one_cell(ctx: &Ctx<'_>, a: &Expr, home: usize, depth: u32) -> bool {
    match a {
        Expr::Ref(r) => r.is_cell(),
        Expr::Name { name, sheet } => names_one(ctx, name, sheet.as_deref(), home, depth),
        _ => false,
    }
}

/// Whether a defined name stands for one value where one is wanted: one
/// no `@` could show before, its text read as [`may_show_at`] reads a
/// formula; a name the workbook does not define is `#NAME?`, one value.
fn names_one(ctx: &Ctx<'_>, name: &str, sheet: Option<&str>, home: usize, depth: u32) -> bool {
    if depth >= NAME_DEPTH {
        return false;
    }
    let scope = match sheet {
        None => None,
        Some(s) => match ctx.env.find_sheet(s) {
            Some(i) => Some(i),
            None => return true,
        },
    };
    let Some(text) = ctx.env.name(name, scope, home) else {
        return true;
    };
    match formula::parse(text) {
        Ok(parsed) => !may_show_at(ctx, &parsed, scope.unwrap_or(home), depth + 1),
        Err(_) => false,
    }
}

#[cfg(test)]
mod tests {
    use crate::calc::{library, Calc};
    use crate::sheet::{parse_a1_range, Cell, Content, Workbook};

    const NA: &str = "(error \"#N/A\")";
    const VALUE: &str = "(error \"#VALUE!\")";
    const SHOWN_WITH_AT: &str = "(not-computed \"a formula Excel shows with @\")";

    /// The cells the formulas read, on sheet `T`: a number, a text, a
    /// formula, nothing, a formula erring, an error typed, a formula of a
    /// truth value and one of a text holding a quote mark; a fill over
    /// `B1:B4`; a dynamic formula and a legacy array formula in `C1` and
    /// `C2`; numbers in `D1:D3`; and `E1` onwards each a formula below.
    fn model(targets: &[&str], formulas: &[&str]) -> Workbook {
        let mut wb = Workbook::new();
        let t = wb.ensure_sheet("T");
        let put = |wb: &mut Workbook, addr: &str, content: Content| {
            let r = parse_a1_range(addr).unwrap();
            wb.sheets[t].set(r.top, r.left, Cell { content, style: 0 });
        };
        put(&mut wb, "A1", Content::Number(10.0));
        put(&mut wb, "A2", Content::Text("x".to_string()));
        put(&mut wb, "A3", Content::Formula("A1*2".to_string()));
        put(&mut wb, "A5", Content::Formula("1/0".to_string()));
        put(&mut wb, "A6", Content::Error("#N/A".to_string()));
        put(&mut wb, "A7", Content::Formula("TRUE".to_string()));
        put(&mut wb, "A8", Content::Formula("\"a\"\"b\"".to_string()));
        wb.sheets[t].set_formula(parse_a1_range("B1:B4").unwrap(), "A1+100", 0);
        put(
            &mut wb,
            "C1",
            Content::DynamicFormula("_xlfn.SEQUENCE(1)".to_string()),
        );
        put(
            &mut wb,
            "C2",
            Content::ArrayFormula("SUM(D1:D3*2)".to_string()),
        );
        put(&mut wb, "D1", Content::Number(1.0));
        put(&mut wb, "D2", Content::Number(2.0));
        put(&mut wb, "D3", Content::Number(3.0));
        for (i, text) in targets.iter().enumerate() {
            let content = Content::Formula(text.to_string());
            wb.sheets[t].set(i as u32 + 1, 5, Cell { content, style: 0 });
        }
        let u = wb.ensure_sheet("U");
        wb.sheets[u].set(
            1,
            1,
            Cell {
                content: Content::Formula("T!A1*3".to_string()),
                style: 0,
            },
        );
        let f = wb.ensure_sheet("F");
        for (i, text) in formulas.iter().enumerate() {
            let content = Content::Formula(text.to_string());
            wb.sheets[f].set(i as u32 + 1, 1, Cell { content, style: 0 });
        }
        wb
    }

    /// Each formula of the pairs, on sheet `F`, against the value it is to
    /// compute, `targets` written into `T!E1` onwards first.
    fn holds(targets: &[&str], cases: &[(&str, &str)]) {
        let formulas: Vec<&str> = cases.iter().map(|(f, _)| *f).collect();
        let wb = model(targets, &formulas);
        let f = wb.find_sheet("F").unwrap();
        let lib = library::language();
        let calc = Calc::run(&wb, &lib);
        for (i, (formula, want)) in cases.iter().enumerate() {
            let got = match calc.computed((f, i as u32 + 1, 1)) {
                Some(c) => c.spell(),
                None => "nothing".to_string(),
            };
            assert_eq!(got, *want, "={formula}");
        }
    }

    #[test]
    fn isformula_tells_a_formula_from_a_value_and_reads_no_value() {
        holds(
            &[],
            &[
                ("ISFORMULA(T!A1)", "false"),
                ("ISFORMULA(T!A2)", "false"),
                ("ISFORMULA(T!A3)", "true"),
                ("ISFORMULA(T!A4)", "false"),
                ("ISFORMULA(T!A5)", "true"),
                ("ISFORMULA(T!A6)", "false"),
                ("ISFORMULA(T!B1)", "true"),
                ("ISFORMULA(T!B3)", "true"),
                ("ISFORMULA(T!C1)", "true"),
                ("ISFORMULA(T!C2)", "true"),
                ("ISFORMULA(U!A1)", "true"),
                ("ISFORMULA(INDEX(T!A1:A8,3))", "true"),
                ("ISFORMULA(IF(TRUE,T!A3,T!A1))", "true"),
                ("NOT(ISFORMULA(T!A1))", "true"),
            ],
        );
    }

    #[test]
    fn formulatext_is_the_text_the_record_shows() {
        holds(
            &["SUM(D1:D3)", "_xlfn.CONCAT(\"a\",\"b\")", "A1 + 1"],
            &[
                ("FORMULATEXT(T!A3)", "\"=A1*2\""),
                ("FORMULATEXT(T!A5)", "\"=1/0\""),
                ("FORMULATEXT(T!A7)", "\"=TRUE\""),
                ("FORMULATEXT(T!A8)", "\"=\\\"a\\\"\\\"b\\\"\""),
                ("FORMULATEXT(T!B1)", "\"=A1+100\""),
                ("FORMULATEXT(T!B3)", "\"=A3+100\""),
                ("FORMULATEXT(T!C1)", "\"=SEQUENCE(1)\""),
                ("FORMULATEXT(T!C2)", "\"{=SUM(D1:D3*2)}\""),
                ("FORMULATEXT(T!E1)", "\"=SUM(D1:D3)\""),
                ("FORMULATEXT(T!E2)", "\"=CONCAT(\\\"a\\\",\\\"b\\\")\""),
                ("FORMULATEXT(T!E3)", "\"=A1 + 1\""),
                ("FORMULATEXT(U!A1)", "\"=T!A1*3\""),
                ("FORMULATEXT(T!A1)", NA),
                ("FORMULATEXT(T!A2)", NA),
                ("FORMULATEXT(T!A4)", NA),
                ("FORMULATEXT(T!A6)", NA),
                ("LEN(FORMULATEXT(T!A3))", "5"),
                ("FORMULATEXT(INDEX(T!A1:A8,3))", "\"=A1*2\""),
            ],
        );
    }

    #[test]
    fn a_reference_of_many_cells_answers_its_first_where_one_value_is_wanted_and_each_as_arrays() {
        holds(
            &[],
            &[
                ("ISFORMULA(T!A1:A8)", "false"),
                ("ISFORMULA(T!A3:A8)", "true"),
                ("FORMULATEXT(T!A1:A8)", NA),
                ("FORMULATEXT(T!A3:A8)", "\"=A1*2\""),
                ("SUMPRODUCT(--ISFORMULA(T!A1:A8))", "4"),
                ("SUMPRODUCT(--ISFORMULA(T!A1:B4))", "5"),
                ("SUMPRODUCT(--ISFORMULA(T!B:B))", "4"),
                ("INDEX(FORMULATEXT(T!A3:B4),2,2)", "\"=A4+100\""),
                ("INDEX(FORMULATEXT(T!A3:A5),2)", NA),
                ("ROWS(FORMULATEXT(T!A1:A8))", "8"),
                // Handed whole to a range's parameter in a legacy formula,
                // the reference is still read where one value is wanted.
                ("SUM(--ISFORMULA(T!A1:A8))", "0"),
                ("COUNTA(FORMULATEXT(T!A1:A8))", "1"),
            ],
        );
    }

    #[test]
    fn an_argument_naming_no_cell_is_its_error_or_a_wrong_value() {
        holds(
            &[],
            &[
                ("ISFORMULA(1)", VALUE),
                ("ISFORMULA(\"T!A3\")", VALUE),
                ("ISFORMULA({1,2})", VALUE),
                ("ISFORMULA(#REF!)", "(error \"#REF!\")"),
                ("FORMULATEXT(1/0)", "(error \"#DIV/0!\")"),
                ("FORMULATEXT(Nowhere!A1)", "(error \"#REF!\")"),
                // Excel 365's save of the fixture: #VALUE! for both.
                ("ISFORMULA(T:U!A1)", VALUE),
                ("FORMULATEXT(T:U!A1)", VALUE),
            ],
        );
    }

    #[test]
    fn a_text_excel_s_bar_spells_otherwise_is_labelled() {
        holds(
            &[
                "D1:D3",
                "D1:D3*2",
                "SUM(D1:D3*2)",
                "IF(A1>0,D1,D2)",
                "INDEX(D1:D3,2)",
                "_xlfn.SINGLE(D1:D3)",
                "ROW(D1:D3)",
                "ROW(D1)",
                "LEN(D1:D3)",
                "VLOOKUP(1,D1:D3,1,FALSE)",
                "COUNTIF(D1:D3,\">1\")",
                "FACT(3)",
                "SUM(FACT(3))",
                "Sales[Amount]",
                "_xlfn.ANCHORARRAY(D1)",
                "IFERROR(D1:D3,0)",
                "SUM(IF(A1>0,D1:D3,0))",
                "SUMPRODUCT(D1:D3*2)",
                "-D1",
                "SUM(IFERROR(D1:D3,0))",
                "SUMPRODUCT(IFERROR(D1:D3,0))",
                "TRANSPOSE(D1:D3)",
                "SUM(TRANSPOSE(D1:D3))",
                "SUM(--ISFORMULA(D1:D3))",
                "SUM(ISFORMULA(D1:D3))",
                "SUMPRODUCT(TRANSPOSE(D1:D3))",
                "SUMPRODUCT(--ISFORMULA(D1:D3))",
                "D1:INDEX(D1:D3,2)",
                "-D1:D3",
            ],
            &[
                ("FORMULATEXT(T!E1)", SHOWN_WITH_AT),
                ("FORMULATEXT(T!E2)", SHOWN_WITH_AT),
                ("FORMULATEXT(T!E3)", SHOWN_WITH_AT),
                ("FORMULATEXT(T!E4)", "\"=IF(A1>0,D1,D2)\""),
                ("FORMULATEXT(T!E5)", SHOWN_WITH_AT),
                ("FORMULATEXT(T!E6)", SHOWN_WITH_AT),
                ("FORMULATEXT(T!E7)", SHOWN_WITH_AT),
                ("FORMULATEXT(T!E8)", "\"=ROW(D1)\""),
                ("FORMULATEXT(T!E9)", SHOWN_WITH_AT),
                ("FORMULATEXT(T!E10)", "\"=VLOOKUP(1,D1:D3,1,FALSE)\""),
                ("FORMULATEXT(T!E11)", "\"=COUNTIF(D1:D3,\\\">1\\\")\""),
                ("FORMULATEXT(T!E12)", SHOWN_WITH_AT),
                ("FORMULATEXT(T!E13)", "\"=SUM(FACT(3))\""),
                ("FORMULATEXT(T!E14)", "(not-computed \"Sales[Amount]\")"),
                (
                    "FORMULATEXT(T!E15)",
                    "(not-computed \"a formula Excel shows with #\")",
                ),
                ("FORMULATEXT(T!E16)", SHOWN_WITH_AT),
                ("FORMULATEXT(T!E17)", "\"=SUM(IF(A1>0,D1:D3,0))\""),
                ("FORMULATEXT(T!E18)", "\"=SUMPRODUCT(D1:D3*2)\""),
                ("FORMULATEXT(T!E19)", "\"=-D1\""),
                // Forced to one value in a legacy formula, inside an array
                // parameter too.
                ("FORMULATEXT(T!E20)", SHOWN_WITH_AT),
                ("FORMULATEXT(T!E21)", SHOWN_WITH_AT),
                // A function answering an array, where one value is wanted
                // and where its answer is handed whole, a legacy formula
                // reading its first value there.
                ("FORMULATEXT(T!E22)", SHOWN_WITH_AT),
                ("FORMULATEXT(T!E23)", SHOWN_WITH_AT),
                // A reference read for its formula, of many cells, as `ROW`'s.
                ("FORMULATEXT(T!E24)", SHOWN_WITH_AT),
                ("FORMULATEXT(T!E25)", SHOWN_WITH_AT),
                // Inside an array parameter, read as arrays: none.
                ("FORMULATEXT(T!E26)", "\"=SUMPRODUCT(TRANSPOSE(D1:D3))\""),
                ("FORMULATEXT(T!E27)", "\"=SUMPRODUCT(--ISFORMULA(D1:D3))\""),
                // `:` between expressions, and a sign over a range.
                ("FORMULATEXT(T!E28)", SHOWN_WITH_AT),
                ("FORMULATEXT(T!E29)", SHOWN_WITH_AT),
                ("ISFORMULA(T!E14)", "true"),
            ],
        );
    }

    #[test]
    fn a_dynamic_formula_and_an_array_formula_show_no_at() {
        let mut wb = model(
            &[],
            &[
                "FORMULATEXT(T!G1)",
                "FORMULATEXT(T!G2)",
                "FORMULATEXT(T!G3)",
            ],
        );
        let t = wb.find_sheet("T").unwrap();
        let dynamic = Content::DynamicFormula("INDEX(D1:D3,2)".to_string());
        wb.sheets[t].set(
            1,
            7,
            Cell {
                content: dynamic,
                style: 0,
            },
        );
        let array = Content::ArrayFormula("D1:D3*2".to_string());
        wb.sheets[t].set(
            2,
            7,
            Cell {
                content: array,
                style: 0,
            },
        );
        let single = Content::DynamicFormula("_xlfn.SINGLE(D1:D3)".to_string());
        wb.sheets[t].set(
            3,
            7,
            Cell {
                content: single,
                style: 0,
            },
        );
        let lib = library::language();
        let calc = Calc::run(&wb, &lib);
        let f = wb.find_sheet("F").unwrap();
        let spelled = |row: u32| calc.computed((f, row, 1)).unwrap().spell();
        assert_eq!(spelled(1), "\"=INDEX(D1:D3,2)\"");
        assert_eq!(spelled(2), "\"{=D1:D3*2}\"");
        assert_eq!(spelled(3), SHOWN_WITH_AT, "SINGLE is the @ in either kind");
    }

    #[test]
    fn a_name_of_one_cell_is_read_as_its_cell_and_a_name_of_more_may_show_an_at() {
        let mut wb = model(
            &["Here*2", "Many*2", "SUM(Many)", "Three+1", "Nobody+1"],
            &[
                "ISFORMULA(Here)",
                "FORMULATEXT(Here)",
                "ISFORMULA(Three)",
                "FORMULATEXT(T!E1)",
                "FORMULATEXT(T!E2)",
                "FORMULATEXT(T!E3)",
                "FORMULATEXT(T!E4)",
                "FORMULATEXT(T!E5)",
            ],
        );
        wb.defined_names
            .push(("Here".to_string(), "T!$A$3".to_string()));
        wb.defined_names
            .push(("Many".to_string(), "T!$D$1:$D$3".to_string()));
        wb.defined_names
            .push(("Three".to_string(), "3".to_string()));
        let lib = library::language();
        let calc = Calc::run(&wb, &lib);
        let f = wb.find_sheet("F").unwrap();
        let spelled = |row: u32| calc.computed((f, row, 1)).unwrap().spell();
        assert_eq!(spelled(1), "true");
        assert_eq!(spelled(2), "\"=A1*2\"");
        assert_eq!(spelled(3), VALUE, "a name that is no reference");
        assert_eq!(spelled(4), "\"=Here*2\"");
        assert_eq!(spelled(5), SHOWN_WITH_AT);
        assert_eq!(spelled(6), "\"=SUM(Many)\"");
        assert_eq!(spelled(7), "\"=Three+1\"");
        assert_eq!(spelled(8), "\"=Nobody+1\"", "#NAME?, one value");
    }

    #[test]
    fn a_text_longer_than_excel_shows_is_na() {
        // 8,192 characters after the `=`, the parser's limit, then one more.
        let at_most = format!("{}11", "1+".repeat(4_095));
        let past = format!("{at_most}1");
        assert_eq!(at_most.len(), crate::calc::formula::LENGTH_LIMIT);
        let wb = model(
            &[&at_most, &past],
            &["LEN(FORMULATEXT(T!E1))", "FORMULATEXT(T!E2)"],
        );
        let lib = library::language();
        let calc = Calc::run(&wb, &lib);
        let f = wb.find_sheet("F").unwrap();
        assert_eq!(calc.computed((f, 1, 1)).unwrap().spell(), "8193");
        assert_eq!(calc.computed((f, 2, 1)).unwrap().spell(), NA);
    }

    #[test]
    fn a_cell_reading_its_own_formula_is_a_cycle_as_a_row_of_its_own_cell_is() {
        let wb = model(&[], &["ISFORMULA(A1)", "ROW(A2)", "FORMULATEXT(A3)"]);
        let lib = library::language();
        let calc = Calc::run(&wb, &lib);
        let f = wb.find_sheet("F").unwrap();
        for row in 1..=3 {
            assert_eq!(
                calc.computed((f, row, 1)).unwrap().spell(),
                "(not-computed cycle)",
                "the owner's decision of 2026-10-21: a reference read for its formula is an edge"
            );
        }
    }
}
