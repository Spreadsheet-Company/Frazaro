//! The information functions (KERNEL.68, 2026-10-30): what a value or a
//! reference is, told without computing with it, so that a formula says "an
//! empty cell", "a text" or "`#DIV/0!`" where it could only say "an error".
//!
//! Each edge is Excel 365's own save of `scripts/recalc/info.txt`, 787 formulas
//! over every kind of cell, every value given directly and a range read every way
//! (`info_saved.xlsx`, held at `no-differ`), and `scripts/recalc/typed.txt`, the
//! same read as a person's typing reads them.
//!
//! `ISNUMBER`, `ISTEXT`, `ISNONTEXT`, `ISLOGICAL`, `ISBLANK`, `ISERR` and
//! `ERROR.TYPE` read one value, as `ISERROR` and `ISNA` do: lifted over an array,
//! a legacy formula meeting a range at its cell and a span of sheets reading as
//! `#REF!`. `NA` reads nothing. The other four tell a reference from what a sign
//! or an operator makes of it, which a value parameter cannot, so that each is a
//! form:
//! - `ISREF` is true for a reference on one sheet, `INDEX`'s, `IF`'s, `CHOOSE`'s
//!   and `XLOOKUP`'s among them, and false for anything else: a span of sheets,
//!   `+A1`, `#REF!` and a name nothing defines. It is never lifted.
//! - `TYPE` is 1, 2, 4 or 16 of one value, an empty cell a number, and 64 of an
//!   array of any size. A range of many cells is an array where the expression
//!   is read as arrays, and in a legacy formula met at its cell, never lifted.
//! - `ISEVEN` and `ISODD` read a number down toward zero, a text as its number
//!   and nothing as 0; a truth value is `#VALUE!` and an error its own. A
//!   reference of many cells is `#VALUE!` even at its own row, as Excel's older
//!   add-in functions read one, and an array lifts them, `+C1:C3` among them.
//!
//! A reference to its own cell is a cycle (`ISREF(A25)` in `A25`), as `ROW` and
//! `ISFORMULA` of their own cell are (the owner's decision of 2026-10-21), where
//! Excel answers `TRUE`. A run calls each at each cell and a program leaves each
//! shape to `Ctx`: what a test reads, an error, an empty cell or a text, no
//! register holds, and its answer, a truth value, would fail a called function's
//! guard at every evaluation.

use super::array::{self, Lifted};
use super::eval::{numeric, Arg, Ctx, Reason};
use super::formula::Expr;
use super::value::{ErrorKind, Value};

/// A test's answer over the one value its call's one argument reads as;
/// `#VALUE!` for a call of more arguments or fewer.
fn tested(ctx: &Ctx<'_>, args: &[Arg<'_>], test: fn(&Value) -> Value) -> Result<Value, Reason> {
    match args {
        [a] => Ok(test(&ctx.one(a)?)),
        _ => Ok(Value::Error(ErrorKind::Value)),
    }
}

/// `ISNUMBER(value)`: whether it is a number; an error is not, and does not
/// propagate. A formula's empty result is 0 and so is one.
pub(crate) fn isnumber(ctx: &Ctx<'_>, args: &[Arg<'_>]) -> Result<Value, Reason> {
    tested(ctx, args, |v| Value::Bool(matches!(v, Value::Number(_))))
}

/// `ISTEXT(value)`: whether it is a text, the empty text among them.
pub(crate) fn istext(ctx: &Ctx<'_>, args: &[Arg<'_>]) -> Result<Value, Reason> {
    tested(ctx, args, |v| Value::Bool(matches!(v, Value::Text(_))))
}

/// `ISNONTEXT(value)`: whether it is anything but a text, nothing and an error
/// among them.
pub(crate) fn isnontext(ctx: &Ctx<'_>, args: &[Arg<'_>]) -> Result<Value, Reason> {
    tested(ctx, args, |v| Value::Bool(!matches!(v, Value::Text(_))))
}

/// `ISLOGICAL(value)`: whether it is a truth value.
pub(crate) fn islogical(ctx: &Ctx<'_>, args: &[Arg<'_>]) -> Result<Value, Reason> {
    tested(ctx, args, |v| Value::Bool(matches!(v, Value::Bool(_))))
}

/// `ISBLANK(value)`: whether a cell holds nothing; a formula's empty result is
/// 0 and an empty text a text, and neither is.
pub(crate) fn isblank(ctx: &Ctx<'_>, args: &[Arg<'_>]) -> Result<Value, Reason> {
    tested(ctx, args, |v| Value::Bool(*v == Value::Empty))
}

/// `ISERR(value)`: whether it is an error other than `#N/A`.
pub(crate) fn iserr(ctx: &Ctx<'_>, args: &[Arg<'_>]) -> Result<Value, Reason> {
    tested(ctx, args, |v| {
        Value::Bool(matches!(v, Value::Error(k) if *k != ErrorKind::NA))
    })
}

/// `ERROR.TYPE(value)`: an error's number, Microsoft's 1 to 7 and Excel's save's
/// 14 for `#CALC!`, and `#N/A` for anything else. Excel's other errors, `#SPILL!`
/// and the rest, are no value of the language (`KERNEL.55`).
pub(crate) fn error_type(ctx: &Ctx<'_>, args: &[Arg<'_>]) -> Result<Value, Reason> {
    tested(ctx, args, |v| match v {
        Value::Error(k) => Value::Number(match k {
            ErrorKind::Null => 1.0,
            ErrorKind::Div0 => 2.0,
            ErrorKind::Value => 3.0,
            ErrorKind::Ref => 4.0,
            ErrorKind::Name => 5.0,
            ErrorKind::Num => 6.0,
            ErrorKind::NA => 7.0,
            ErrorKind::Calc => 14.0,
        }),
        _ => Value::Error(ErrorKind::NA),
    })
}

/// `NA()`: the error itself.
pub(crate) fn na(_: &Ctx<'_>, args: &[Arg<'_>]) -> Result<Value, Reason> {
    Ok(Value::Error(match args {
        [] => ErrorKind::NA,
        _ => ErrorKind::Value,
    }))
}

/// `TRUE()`: the value, as a literal is, and OpenFormula's one spelling of it, as
/// LibreOffice writes a boolean (KERNEL.58's *Found* 14).
pub(crate) fn true_(_: &Ctx<'_>, args: &[Arg<'_>]) -> Result<Value, Reason> {
    Ok(match args {
        [] => Value::Bool(true),
        _ => Value::Error(ErrorKind::Value),
    })
}

/// `FALSE()`: as [`true_`].
pub(crate) fn false_(_: &Ctx<'_>, args: &[Arg<'_>]) -> Result<Value, Reason> {
    Ok(match args {
        [] => Value::Bool(false),
        _ => Value::Error(ErrorKind::Value),
    })
}

/// `ISREF(value)`: whether its argument answers a reference on one sheet, and
/// not what a sign makes of one, which Excel reads as a value.
pub(crate) fn isref<'e>(ctx: &Ctx<'_>, args: &'e [Expr]) -> Result<Arg<'e>, Reason> {
    let [e] = args else {
        return Ok(Arg::Scalar(Value::Error(ErrorKind::Value)));
    };
    let answer = ctx.eval(e)?;
    let one_sheet = matches!(&answer, Arg::Area { areas, .. } if areas.len() == 1);
    Ok(Arg::Scalar(Value::Bool(
        one_sheet && !matches!(e, Expr::Pos(_)),
    )))
}

/// `TYPE`'s number for one value.
fn type_of(v: &Value) -> f64 {
    match v {
        Value::Number(_) | Value::Empty => 1.0,
        Value::Text(_) => 2.0,
        Value::Bool(_) => 4.0,
        Value::Error(_) => 16.0,
    }
}

/// `TYPE(value)`, one number whatever it is given. See the module's comment.
pub(crate) fn type_<'e>(ctx: &Ctx<'_>, args: &'e [Expr]) -> Result<Arg<'e>, Reason> {
    let [e] = args else {
        return Ok(Arg::Scalar(Value::Error(ErrorKind::Value)));
    };
    let answer = ctx.eval(e)?;
    let number = match &answer {
        Arg::Scalar(v) => type_of(v),
        Arg::Array(_) => 64.0,
        Arg::Area { areas, .. } => match &areas[..] {
            [a] if a.is_cell() => type_of(&ctx.env.read((a.sheet, a.top, a.left))?),
            [_] if ctx.arrays() => 64.0,
            _ => type_of(&ctx.one(&answer)?),
        },
    };
    Ok(Arg::Scalar(Value::Number(number)))
}

/// One value's answer to `ISEVEN`, or to `ISODD` when `odd`. A double past 2^53
/// is even, which `ISEVEN(1E+300)` is in Excel's save too, so that the rest of
/// its division by 2 is the rule at every size.
fn parity_of(v: &Value, odd: bool) -> Result<Value, Reason> {
    let n = match v {
        Value::Error(k) => return Ok(Value::Error(*k)),
        Value::Bool(_) => return Ok(Value::Error(ErrorKind::Value)),
        v => match numeric(v)? {
            Ok(n) => n,
            Err(k) => return Ok(Value::Error(k)),
        },
    };
    Ok(Value::Bool((n.trunc() % 2.0 != 0.0) == odd))
}

/// `ISEVEN(number)` and `ISODD(number)`. See the module's comment: a sign reads
/// its operand as an operator does, an array where the expression is read as
/// arrays and a range's cell where it is not, and any other reference of many
/// cells is `#VALUE!`.
fn parity<'e>(ctx: &Ctx<'_>, args: &'e [Expr], odd: bool) -> Result<Arg<'e>, Reason> {
    let [e] = args else {
        return Ok(Arg::Scalar(Value::Error(ErrorKind::Value)));
    };
    let given = match e {
        Expr::Pos(_) => ctx.operand(e)?,
        _ => match ctx.eval(e)? {
            Arg::Area { areas, .. } => match &areas[..] {
                [a] if a.is_cell() => Lifted::One(ctx.env.read((a.sheet, a.top, a.left))?),
                _ => return Ok(Arg::Scalar(Value::Error(ErrorKind::Value))),
            },
            other => ctx.lifted(other)?,
        },
    };
    Ok(match given {
        Lifted::One(v) => Arg::Scalar(parity_of(&v, odd)?),
        Lifted::Many(a) => Arg::Array(array::lift1(&a, |v| parity_of(v, odd))?),
    })
}

/// `ISEVEN(number)` ([`parity`]).
pub(crate) fn iseven<'e>(ctx: &Ctx<'_>, args: &'e [Expr]) -> Result<Arg<'e>, Reason> {
    parity(ctx, args, false)
}

/// `ISODD(number)` ([`parity`]).
pub(crate) fn isodd<'e>(ctx: &Ctx<'_>, args: &'e [Expr]) -> Result<Arg<'e>, Reason> {
    parity(ctx, args, true)
}

#[cfg(test)]
mod tests {
    use crate::calc::{library, Calc};
    use crate::machine::Machine;
    use crate::sheet::{parse_a1_range, Cell, Content, Workbook};

    const NA: &str = "(error \"#N/A\")";
    const VALUE: &str = "(error \"#VALUE!\")";

    /// A sheet `T` of each kind of cell: `A1` a number, `A2` a text, `A3` nothing,
    /// `A4` a truth value, `A5` an error, `A6` `#N/A`, `A7` an empty text, `A8` an
    /// empty read, `A9` a numeric text; numbers in `C1:C3`; and a sheet `U` after
    /// it, so that `T:U` is a span of sheets.
    const CELLS: &str = "(sheet \"T\" visible)\n\
(cell \"T\" \"A1\" 10)\n\
(cell \"T\" \"A2\" \"x\")\n\
(formula \"T\" \"A4\" \"=TRUE\")\n\
(formula \"T\" \"A5\" \"=1/0\")\n\
(formula \"T\" \"A6\" \"=NA()\")\n\
(formula \"T\" \"A7\" \"=\\\"\\\"\")\n\
(formula \"T\" \"A8\" \"=A3\")\n\
(cell \"T\" \"A9\" \"5\")\n\
(cell \"T\" \"C1\" 1)\n\
(cell \"T\" \"C2\" 2)\n\
(cell \"T\" \"C3\" 3)\n\
(sheet \"U\" visible)\n\
(cell \"U\" \"A1\" 7)\n";

    /// Each formula's value at `E<n>` of a machine over [`CELLS`], its frame's
    /// answer, held to a door's `Calc` over the same grid.
    fn answers(formulas: &[&str]) -> Vec<String> {
        let mut rows = String::from(CELLS);
        for (k, f) in formulas.iter().enumerate() {
            let text = crate::rows::quoted(f);
            rows.push_str(&format!("(formula \"T\" \"E{}\" {text})\n", k + 20));
        }
        let mut m = Machine::load(&rows, library::language())
            .unwrap_or_else(|e| panic!("line {}: {}", e.line, e.refusal.text));
        assert!(m.step(0).done);
        let book = m.workbook();
        let lib = library::language();
        let calc = Calc::run(&book, &lib);
        let at = |k: usize| {
            let r = parse_a1_range(&format!("E{}", k + 20)).unwrap();
            (0usize, r.top, r.left)
        };
        (0..formulas.len())
            .map(|k| {
                let id = at(k);
                let door = calc.computed(id).expect("computed").spell();
                let kept = crate::calc::Values::computed(&m, id).expect("computed");
                assert_eq!(
                    kept.spell(),
                    door,
                    "{}: the door and the machine",
                    formulas[k]
                );
                door
            })
            .collect()
    }

    fn holds(cases: &[(&str, &str)]) {
        let formulas: Vec<&str> = cases.iter().map(|c| c.0).collect();
        for (got, (f, want)) in answers(&formulas).iter().zip(cases) {
            assert_eq!(got, want, "{f}");
        }
    }

    #[test]
    fn each_test_tells_each_kind_of_cell_as_excel_s_save_does() {
        holds(&[
            ("=ISNUMBER(A1)", "true"),
            ("=ISNUMBER(A2)", "false"),
            ("=ISNUMBER(A3)", "false"),
            ("=ISNUMBER(A4)", "false"),
            ("=ISNUMBER(A5)", "false"),
            ("=ISNUMBER(A8)", "true"),
            ("=ISNUMBER(A9)", "false"),
            ("=ISTEXT(A2)", "true"),
            ("=ISTEXT(A7)", "true"),
            ("=ISTEXT(A9)", "true"),
            ("=ISTEXT(A3)", "false"),
            ("=ISNONTEXT(A3)", "true"),
            ("=ISNONTEXT(A5)", "true"),
            ("=ISNONTEXT(A7)", "false"),
            ("=ISLOGICAL(A4)", "true"),
            ("=ISLOGICAL(A1)", "false"),
            ("=ISBLANK(A3)", "true"),
            ("=ISBLANK(A7)", "false"),
            ("=ISBLANK(A8)", "false"),
            ("=ISBLANK(0)", "false"),
            ("=ISBLANK(\"\")", "false"),
            ("=ISBLANK(Z99)", "true"),
            ("=ISERR(A5)", "true"),
            ("=ISERR(A6)", "false"),
            ("=ISERR(A1)", "false"),
            ("=ERROR.TYPE(A5)", "2"),
            ("=ERROR.TYPE(A6)", "7"),
            ("=ERROR.TYPE(A1)", NA),
            ("=ERROR.TYPE(A3)", NA),
            ("=ERROR.TYPE(#NULL!)", "1"),
            ("=ERROR.TYPE(#VALUE!)", "3"),
            ("=ERROR.TYPE(#REF!)", "4"),
            ("=ERROR.TYPE(#NAME?)", "5"),
            ("=ERROR.TYPE(#NUM!)", "6"),
            ("=ERROR.TYPE(FILTER(C1:C3,C1:C3>100))", "14"),
            ("=NA()", NA),
            ("=ISNA(NA())", "true"),
            ("=TYPE(A1)", "1"),
            ("=TYPE(A3)", "1"),
            ("=TYPE(A2)", "2"),
            ("=TYPE(A7)", "2"),
            ("=TYPE(A4)", "4"),
            ("=TYPE(A5)", "16"),
        ]);
    }

    #[test]
    fn a_range_is_met_at_its_cell_and_lifted_where_read_as_arrays() {
        holds(&[
            // Off the range's rows: `#VALUE!`, which each reads as an error.
            ("=ISNUMBER(C1:C3)", "false"),
            ("=ISERR(C1:C3)", "true"),
            ("=ERROR.TYPE(C1:C3)", "3"),
            ("=TYPE(C1:C3)", "16"),
            ("=TYPE(C1:C3*2)", "16"),
            ("=SUMPRODUCT(--ISNUMBER(C1:C3))", "3"),
            ("=SUMPRODUCT(--ISBLANK(A1:A4))", "1"),
            ("=SUMPRODUCT(ISNONTEXT(A1:A3)*1)", "2"),
            ("=ROWS(ISTEXT(C1:C3))", "3"),
            ("=ROWS(ERROR.TYPE(C1:C3))", "3"),
            // A span of sheets reads as `#REF!`.
            ("=ISNUMBER(T:U!A1)", "false"),
            ("=ERROR.TYPE(T:U!A1)", "4"),
            ("=TYPE(T:U!A1)", "16"),
        ]);
    }

    #[test]
    fn isref_tells_a_reference_from_what_a_sign_makes_of_it() {
        holds(&[
            ("=ISREF(A1)", "true"),
            ("=ISREF(A3)", "true"),
            ("=ISREF(C1:C3)", "true"),
            ("=ISREF(C:C)", "true"),
            ("=ISREF((A1))", "true"),
            ("=ISREF(A1:A1)", "true"),
            ("=ISREF(INDEX(C1:C3,2))", "true"),
            ("=ISREF(INDEX(C1:C3,0))", "true"),
            ("=ISREF(IF(TRUE,A1,A2))", "true"),
            ("=ISREF(CHOOSE(1,A1,A2))", "true"),
            ("=ISREF(A1:INDEX(C1:C3,2))", "true"),
            ("=ISREF(1)", "false"),
            ("=ISREF(\"A1\")", "false"),
            ("=ISREF(TRUE)", "false"),
            ("=ISREF(#REF!)", "false"),
            ("=ISREF(1/0)", "false"),
            ("=ISREF({1,2})", "false"),
            ("=ISREF(IF(FALSE,A1,2))", "false"),
            ("=ISREF(A1+0)", "false"),
            ("=ISREF(+A1)", "false"),
            ("=ISREF(+C1:C3)", "false"),
            ("=ISREF(-A1)", "false"),
            ("=ISREF(NOSUCHNAME)", "false"),
            ("=ISREF(T:U!A1)", "false"),
            ("=SUMPRODUCT(--ISREF(C1:C3))", "1"),
            ("=ROWS(ISREF(C1:C3))", "1"),
        ]);
    }

    #[test]
    fn type_is_one_number_whatever_it_is_given() {
        holds(&[
            ("=TYPE({1,2})", "64"),
            ("=TYPE({1})", "64"),
            ("=TYPE(SEQUENCE(1))", "64"),
            ("=TYPE(FILTER(C1:C3,C1:C3>1))", "64"),
            ("=TYPE(FILTER(C1:C3,C1:C3>100))", "16"),
            ("=TYPE(INDEX(C1:C3,2))", "1"),
            ("=TYPE(C2:C2)", "1"),
            ("=TYPE(+C2)", "1"),
            ("=TYPE(TRUE+0)", "1"),
            ("=TYPE(#REF!)", "16"),
            ("=TYPE(1/0)", "16"),
            ("=SUMPRODUCT(TYPE(C1:C3))", "64"),
            ("=SUMPRODUCT(TYPE(+C1:C3))", "64"),
            ("=SUMPRODUCT(TYPE(C1:C1))", "1"),
            ("=SUMPRODUCT(--TYPE(C1:C3*1))", "64"),
            ("=ROWS(TYPE(C1:C3))", "1"),
            ("=INDEX(TYPE(C1:C3),2)", "(error \"#REF!\")"),
        ]);
    }

    #[test]
    fn iseven_and_isodd_read_a_number_down_and_refuse_a_range() {
        holds(&[
            ("=ISEVEN(A1)", "true"),
            ("=ISEVEN(A3)", "true"),
            ("=ISODD(A3)", "false"),
            ("=ISEVEN(A8)", "true"),
            ("=ISODD(A9)", "true"),
            ("=ISEVEN(2.9)", "true"),
            ("=ISEVEN(-2.5)", "true"),
            ("=ISEVEN(-0.5)", "true"),
            ("=ISODD(-0.5)", "false"),
            ("=ISODD(1.9)", "true"),
            ("=ISODD(-1.9)", "true"),
            ("=ISEVEN(0)", "true"),
            ("=ISEVEN(1E+300)", "true"),
            ("=ISODD(-1E+300)", "false"),
            ("=ISEVEN(\"2\")", "true"),
            ("=ISEVEN(\"x\")", VALUE),
            ("=ISEVEN(A2)", VALUE),
            ("=ISEVEN(A7)", VALUE),
            ("=ISEVEN(TRUE)", VALUE),
            ("=ISODD(A4)", VALUE),
            ("=ISEVEN(A5)", "(error \"#DIV/0!\")"),
            ("=ISODD(A6)", NA),
            ("=ISEVEN(#REF!)", "(error \"#REF!\")"),
            // A reference of many cells, even at its own row, or through a call.
            ("=ISEVEN(C1:C3)", VALUE),
            ("=ISEVEN(C:C)", VALUE),
            ("=ISEVEN(INDEX(C1:C3,0))", VALUE),
            ("=ISEVEN(IF(TRUE,C1:C3))", VALUE),
            ("=ISEVEN(T:U!A1)", VALUE),
            ("=SUMPRODUCT(--ISEVEN(C1:C3))", VALUE),
            ("=INDEX(ISEVEN(C1:C3),2)", VALUE),
            // One cell, however it is spelled.
            ("=ISEVEN(C2:C2)", "true"),
            ("=ISEVEN(INDEX(C1:C3,2))", "true"),
            ("=ISEVEN(IF(TRUE,C2))", "true"),
            // An array lifts, a sign making an array of a range where read as
            // arrays and meeting it at its cell where not.
            ("=ISODD({1,2})", "true"),
            ("=SUMPRODUCT(--ISEVEN(C1:C3*1))", "1"),
            ("=SUMPRODUCT(--ISODD(C1:C3*1))", "2"),
            ("=SUMPRODUCT(--ISEVEN(+C1:C3))", "1"),
            ("=SUMPRODUCT(--ISODD(+C1:C3))", "2"),
            ("=ROWS(ISEVEN(+C1:C3))", "3"),
            ("=ISEVEN(+C1:C3)", VALUE),
        ]);
    }

    #[test]
    fn a_count_no_excel_cell_holds_is_value() {
        holds(&[
            ("=ISNUMBER()", VALUE),
            ("=ISNUMBER(1,2)", VALUE),
            ("=ISREF()", VALUE),
            ("=ISREF(A1,A2)", VALUE),
            ("=TYPE()", VALUE),
            ("=ISEVEN()", VALUE),
            ("=NA(1)", VALUE),
        ]);
    }

    #[test]
    fn a_test_of_its_own_cell_is_a_cycle_as_row_s_is() {
        // The owner's decision of 2026-10-21, as for `ROW` and `ISFORMULA`.
        let mut wb = Workbook::new();
        let s = wb.ensure_sheet("T");
        let r = parse_a1_range("A1").unwrap();
        wb.sheets[s].set(
            r.top,
            r.left,
            Cell {
                content: Content::Formula("ISREF(A1)".to_string()),
                style: 0,
            },
        );
        let lib = library::language();
        let calc = Calc::run(&wb, &lib);
        assert_eq!(
            calc.computed((s, r.top, r.left)).expect("computed").spell(),
            "(not-computed cycle)"
        );
    }
}
