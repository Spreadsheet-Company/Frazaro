//! The lookups (KERNEL.8, slice a, 2026-10-18): `INDEX`, `MATCH`, `XMATCH`,
//! `VLOOKUP`, `HLOOKUP`, `LOOKUP` and `XLOOKUP`, and the four error tests a
//! lookup's miss is read by, `IFERROR`, `IFNA`, `ISERROR` and `ISNA`. The
//! language's own, since a game reaches the language alone (Alonzo's
//! `AD-1`) and Alonzo's `KIT.3`, `CART.3`, `CART.5` and `CART.13` and
//! Buffet's `MISE.6` ask for them; Excel's library has them through it.
//!
//! Each rule is Excel's as its documentation states it; where the
//! documentation is silent, the slice's fixture holds the rule to Excel's
//! own saved answer (`scripts/recalc/lookups.txt`, saved by Excel as
//! `lookups_saved.xlsx`). An exact search goes by the lookups' equality, no
//! two kinds equal (5 is not `"5"` to a lookup), a number by its value
//! exactly, where an operator holds two numbers equal to fifteen digits
//! equal (`MATCH(0.3, …, 0)` finds no `0.1+0.2` in Excel's save), a text
//! without case, beyond ASCII too (`value::fold_char`), a truth value; an
//! error, the blank and an empty probe found by nothing. Every search is
//! `super::index`'s, by a vector's index where its environment keeps one
//! and by the scan that is its meaning where it does not; an approximate
//! match by the bisection stated there, over the vector as it stands, but
//! `MATCH`'s -1, which walks from the first cell (`index::descending`). A
//! reason read answers, as a cell not computed answers wherever it is read.

use std::cmp::Ordering;

use super::criteria::{as_written, has_wildcard, wildcard};
use super::eval::{exact_order, Arg, Ctx, Reason};
use super::formula::Expr;
use super::graph::{Area, Areas};
use super::index::{self, Line};
use super::library::picked;
use super::value::{fold_cmp, ErrorKind, Value};

fn error<'e>(k: ErrorKind) -> Result<Arg<'e>, Reason> {
    Ok(Arg::Scalar(Value::Error(k)))
}

/// An argument read as a number, as an operand is: `Err` the error it
/// answers instead.
fn number_of(ctx: &Ctx<'_>, arg: &Arg<'_>) -> Result<Result<f64, ErrorKind>, Reason> {
    Ok(ctx.one(arg)?.number())
}

/// An argument that may be left out, read as a number; left out or given
/// as nothing, `default`.
fn number_or(
    ctx: &Ctx<'_>,
    arg: Option<&Arg<'_>>,
    default: f64,
) -> Result<Result<f64, ErrorKind>, Reason> {
    let Some(arg) = arg else {
        return Ok(Ok(default));
    };
    Ok(match ctx.one(arg)? {
        Value::Empty => Ok(default),
        v => v.number(),
    })
}

/// [`number_or`] over an argument as written, for the forms.
fn expr_number_or(
    ctx: &Ctx<'_>,
    e: Option<&Expr>,
    default: f64,
) -> Result<Result<f64, ErrorKind>, Reason> {
    let Some(e) = e else {
        return Ok(Ok(default));
    };
    Ok(match ctx.scalar(e)? {
        Value::Empty => Ok(default),
        v => v.number(),
    })
}

/// The one area of a reference, `None` for a value or a span of sheets.
fn area_of(arg: &Arg<'_>) -> Option<Area> {
    match arg {
        Arg::Area { areas, .. } => match &areas[..] {
            [a] => Some(*a),
            _ => None,
        },
        Arg::Scalar(_) => None,
    }
}

/// Whether one value given where a vector is wanted, as Excel reads a value
/// there, an array of one, is kept by a search's rule: of the probe's kind
/// and `keep` keeping how it compares with the probe.
fn one_kept(v: &Value, probe: &Value, keep: impl Fn(Ordering) -> bool) -> bool {
    let kind = |x: &Value| match x {
        Value::Number(_) => Some(0),
        Value::Text(_) => Some(1),
        Value::Bool(_) => Some(2),
        _ => None,
    };
    if kind(v).is_none() || kind(v) != kind(probe) {
        return false;
    }
    let o = match (v, probe) {
        (Value::Number(a), Value::Number(b)) => exact_order(*a, *b),
        (Value::Text(a), Value::Text(b)) => fold_cmp(a, b),
        (Value::Bool(a), Value::Bool(b)) => a.cmp(b),
        _ => return false,
    };
    keep(o)
}

/// An exact match as `MATCH` 0, `VLOOKUP` and `HLOOKUP` take it: a text
/// holding `*`, `?` or `~` a pattern, any other probe by its key.
fn exactly(ctx: &Ctx<'_>, line: Line, probe: &Value) -> Result<Option<u32>, Reason> {
    match probe {
        Value::Text(t) if has_wildcard(t) => index::pattern(ctx, line, t, false),
        _ => index::exact(ctx, line, probe, false),
    }
}

/// `INDEX(reference, row, [column], [area])`: the cell of the reference's
/// area at the row and the column counted from its corner, as a reference
/// (`library::picked`): a reference one row high takes a single number as
/// its column; a row of 0 is a whole column, a column of 0 a whole row; a
/// number read down to a whole one; outside the area `#REF!`, below 0
/// `#VALUE!`. A value given as the reference is an array of one. Only area
/// 1 of the one area a reference names, a union being no reference this
/// version reads.
pub(crate) fn index<'e>(ctx: &Ctx<'_>, args: &'e [Expr]) -> Result<Arg<'e>, Reason> {
    if !(2..=4).contains(&args.len()) {
        return error(ErrorKind::Value);
    }
    let reference = ctx.eval(&args[0])?;
    let row = match ctx.scalar(&args[1])?.number() {
        Ok(n) => n,
        Err(k) => return error(k),
    };
    let column = match args.get(2) {
        None => None,
        Some(e) => match ctx.scalar(e)?.number() {
            Ok(n) => Some(n),
            Err(k) => return error(k),
        },
    };
    if let Some(e) = args.get(3) {
        match expr_number_or(ctx, Some(e), 1.0)? {
            Ok(n) if n.trunc() == 1.0 => {}
            Ok(_) => return error(ErrorKind::Ref),
            Err(k) => return error(k),
        }
    }
    match reference {
        Arg::Area {
            areas,
            written,
            moved,
        } => {
            let [a] = &areas[..] else {
                return error(ErrorKind::Ref);
            };
            let (rows, cols) = (a.bottom - a.top + 1, a.right - a.left + 1);
            match picked(rows, cols, row, column) {
                Ok(p) => Ok(Arg::Area {
                    areas: Areas::one(Area {
                        sheet: a.sheet,
                        top: a.top + p.0,
                        left: a.left + p.1,
                        bottom: a.top + p.2,
                        right: a.left + p.3,
                    }),
                    written,
                    moved,
                }),
                Err(k) => error(k),
            }
        }
        Arg::Scalar(v) => match picked(1, 1, row, column) {
            Ok(_) => Ok(Arg::Scalar(v)),
            Err(k) => error(k),
        },
    }
}

/// [`index`] as a value, for a program's call through its pointer
/// (`super::program`, KERNEL.8): the cell it picks read as a read is, a
/// pick of many cells labelled as `Ctx` labels a range where one value is
/// wanted, a value given as the reference an array of one. Its operands are
/// those a program hands a call, a reference as the cells it names or
/// `#REF!` where it leaves the sheet, and values; the same meaning as the
/// form's, which the machine's tests hold to each other.
pub(crate) fn index_value(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    if !(2..=3).contains(&args.len()) {
        return Ok(Value::Error(ErrorKind::Value));
    }
    let row = match ctx.one(&args[1])?.number() {
        Ok(n) => n,
        Err(k) => return Ok(Value::Error(k)),
    };
    let column = match args.get(2) {
        None => None,
        Some(a) => match ctx.one(a)?.number() {
            Ok(n) => Some(n),
            Err(k) => return Ok(Value::Error(k)),
        },
    };
    match &args[0] {
        Arg::Area {
            areas,
            written,
            moved,
        } => {
            let [a] = &areas[..] else {
                return Ok(Value::Error(ErrorKind::Ref));
            };
            let (rows, cols) = (a.bottom - a.top + 1, a.right - a.left + 1);
            match picked(rows, cols, row, column) {
                Ok((t, l, b, r)) if t == b && l == r => {
                    ctx.env.read((a.sheet, a.top + t, a.left + l))
                }
                // A pick of many cells, where one value is wanted: met at
                // the cell in a legacy formula, an array in a dynamic one,
                // labelled until arrays come (KERNEL.8, slice b).
                Ok(_) if ctx.env.dynamic() => Err(as_written(written, *moved)),
                Ok((t, l, b, r)) => ctx.intersect(&Area {
                    sheet: a.sheet,
                    top: a.top + t,
                    left: a.left + l,
                    bottom: a.top + b,
                    right: a.left + r,
                }),
                Err(k) => Ok(Value::Error(k)),
            }
        }
        Arg::Scalar(v) => match picked(1, 1, row, column) {
            Ok(_) => Ok(v.clone()),
            Err(k) => Ok(Value::Error(k)),
        },
    }
}

/// `MATCH(value, vector, [type])`: the position, from 1, of the value in a
/// vector of one row or one column (any other `#N/A`): type 0 exactly, a
/// text holding a wildcard a pattern; type 1, or none, the bisection's
/// largest not past the value; type -1 the smallest not under it on a
/// descending vector, by the walk from its first cell
/// (`index::descending`); none `#N/A`; a value that is an error answers it.
/// A type given as nothing, `MATCH(x, v, )` or an empty cell, is 0, as a
/// number read from nothing is.
pub(crate) fn match_(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    if !(2..=3).contains(&args.len()) {
        return Ok(Value::Error(ErrorKind::Value));
    }
    let probe = ctx.one(&args[0])?;
    if let Value::Error(k) = probe {
        return Ok(Value::Error(k));
    }
    let kind = match args.get(2) {
        None => 1.0,
        Some(a) => match number_of(ctx, a)? {
            Ok(n) => n,
            Err(k) => return Ok(Value::Error(k)),
        },
    };
    let found = match &args[1] {
        Arg::Scalar(v) => {
            let kept = if kind == 0.0 {
                match &probe {
                    Value::Text(t) if has_wildcard(t) => {
                        matches!(v, Value::Text(x) if wildcard(t, x))
                    }
                    _ => one_kept(v, &probe, |o| o == Ordering::Equal),
                }
            } else if kind > 0.0 {
                one_kept(v, &probe, |o| o != Ordering::Greater)
            } else {
                one_kept(v, &probe, |o| o != Ordering::Less)
            };
            kept.then_some(0)
        }
        arg => {
            let Some(line) = area_of(arg).and_then(Line::of) else {
                return Ok(Value::Error(ErrorKind::NA));
            };
            if kind == 0.0 {
                exactly(ctx, line, &probe)?
            } else if kind > 0.0 {
                index::bisect(ctx, line, &probe, |o| o != Ordering::Greater)?
            } else {
                index::descending(ctx, line, &probe)?
            }
        }
    };
    Ok(match found {
        Some(k) => Value::Number(f64::from(k) + 1.0),
        None => Value::Error(ErrorKind::NA),
    })
}

/// The mode and the search of `XMATCH` and `XLOOKUP`, each read down to a
/// whole number: mode 0, -1, 1 or 2 and search 1, -1, 2 or -2, mode 2 taking
/// no bisection; any other `#VALUE!`. Mode 3, the regular expressions Excel
/// has added, is labelled (`match mode 3`), not computed: Excel's save of
/// the lookups' fixture takes it (`Lookups!A46`).
fn modes(mode: f64, search: f64) -> Result<Result<(i8, i8), ErrorKind>, Reason> {
    let (mode, search) = (mode.trunc(), search.trunc());
    let modes = [0.0, -1.0, 1.0, 2.0, 3.0];
    let searches = [1.0, -1.0, 2.0, -2.0];
    if !modes.contains(&mode) || !searches.contains(&search) {
        return Ok(Err(ErrorKind::Value));
    }
    if mode == 3.0 {
        return Err(Reason::Construct("match mode 3".to_string()));
    }
    if mode == 2.0 && search.abs() == 2.0 {
        return Ok(Err(ErrorKind::Value));
    }
    Ok(Ok((mode as i8, search as i8)))
}

/// The search of `XMATCH` and `XLOOKUP` ([`modes`]): mode 0 exactly with no
/// wildcard, -1 exactly or the next smaller, 1 exactly or the next larger,
/// 2 a pattern for a text and exactly for any other value; search 1 first to
/// last, -1 last to first, 2 and -2 the bisection over an ascending and a
/// descending vector.
fn xsearch(
    ctx: &Ctx<'_>,
    line: Line,
    probe: &Value,
    (mode, search): (i8, i8),
) -> Result<Option<u32>, Reason> {
    if search.abs() == 1 {
        let from_end = search < 0;
        return match (mode, probe) {
            (2, Value::Text(t)) if has_wildcard(t) => index::pattern(ctx, line, t, from_end),
            (-1, _) => index::nearest(ctx, line, probe, false, from_end),
            (1, _) => index::nearest(ctx, line, probe, true, from_end),
            _ => index::exact(ctx, line, probe, from_end),
        };
    }
    let equal = |k: Option<u32>| -> Result<Option<u32>, Reason> {
        let Some(k) = k else {
            return Ok(None);
        };
        let v = ctx.env.read(line.cell(k))?;
        Ok(one_kept(&v, probe, |o| o == Ordering::Equal).then_some(k))
    };
    if search > 0 {
        match mode {
            -1 => index::bisect(ctx, line, probe, |o| o != Ordering::Greater),
            1 => {
                let under = index::bisect(ctx, line, probe, |o| o == Ordering::Less)?;
                index::next_of_kind(ctx, line, probe, under)
            }
            _ => equal(index::bisect(ctx, line, probe, |o| o != Ordering::Greater)?),
        }
    } else {
        match mode {
            1 => index::bisect(ctx, line, probe, |o| o != Ordering::Less),
            -1 => {
                let over = index::bisect(ctx, line, probe, |o| o == Ordering::Greater)?;
                index::next_of_kind(ctx, line, probe, over)
            }
            _ => equal(index::bisect(ctx, line, probe, |o| o != Ordering::Less)?),
        }
    }
}

/// One value given where `XMATCH`'s or `XLOOKUP`'s vector is wanted, an
/// array of one: kept by the mode's rule.
fn xsearch_one(v: &Value, probe: &Value, mode: i8) -> Option<u32> {
    let kept = match mode {
        -1 => one_kept(v, probe, |o| o != Ordering::Greater),
        1 => one_kept(v, probe, |o| o != Ordering::Less),
        2 => match probe {
            Value::Text(t) => matches!(v, Value::Text(x) if wildcard(t, x)),
            _ => one_kept(v, probe, |o| o == Ordering::Equal),
        },
        _ => one_kept(v, probe, |o| o == Ordering::Equal),
    };
    kept.then_some(0)
}

/// `XMATCH(value, vector, [mode], [search])`: the position, from 1, of the
/// value in a vector of one row or one column (any other `#VALUE!`), by
/// [`xsearch`]; none `#N/A`.
pub(crate) fn xmatch(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    if !(2..=4).contains(&args.len()) {
        return Ok(Value::Error(ErrorKind::Value));
    }
    let probe = ctx.one(&args[0])?;
    if let Value::Error(k) = probe {
        return Ok(Value::Error(k));
    }
    let mode = match number_or(ctx, args.get(2), 0.0)? {
        Ok(n) => n,
        Err(k) => return Ok(Value::Error(k)),
    };
    let search = match number_or(ctx, args.get(3), 1.0)? {
        Ok(n) => n,
        Err(k) => return Ok(Value::Error(k)),
    };
    let modes = match modes(mode, search)? {
        Ok(m) => m,
        Err(e) => return Ok(Value::Error(e)),
    };
    let found = match &args[1] {
        Arg::Scalar(v) => xsearch_one(v, &probe, modes.0),
        arg => {
            let Some(line) = area_of(arg).and_then(Line::of) else {
                return Ok(Value::Error(ErrorKind::Value));
            };
            xsearch(ctx, line, &probe, modes)?
        }
    };
    Ok(match found {
        Some(k) => Value::Number(f64::from(k) + 1.0),
        None => Value::Error(ErrorKind::NA),
    })
}

/// `VLOOKUP(value, table, column, [approximate])`, or `HLOOKUP` `across`:
/// the table's first column (row) searched as `MATCH` searches it, exactly
/// when the fourth argument is FALSE or 0, by the bisection when it is TRUE,
/// another number or left out, and the value in the column (row) asked for;
/// a column below 1 `#VALUE!`, past the table `#REF!`; none found `#N/A`;
/// an empty cell found answers as an empty cell, which a formula's cell
/// holds as 0.
fn table_lookup(ctx: &Ctx<'_>, args: &[Arg], across: bool) -> Result<Value, Reason> {
    if !(3..=4).contains(&args.len()) {
        return Ok(Value::Error(ErrorKind::Value));
    }
    let probe = ctx.one(&args[0])?;
    if let Value::Error(k) = probe {
        return Ok(Value::Error(k));
    }
    let at = match number_of(ctx, &args[2])? {
        Ok(n) => n.trunc(),
        Err(k) => return Ok(Value::Error(k)),
    };
    let approximate = match args.get(3) {
        None => true,
        Some(a) => match ctx.one(a)?.truth() {
            Ok(b) => b,
            Err(k) => return Ok(Value::Error(k)),
        },
    };
    if at < 1.0 {
        return Ok(Value::Error(ErrorKind::Value));
    }
    let table = match &args[1] {
        Arg::Scalar(v) => {
            // A value given as the table, an array of one.
            if at > 1.0 {
                return Ok(Value::Error(ErrorKind::Ref));
            }
            let kept = if approximate {
                one_kept(v, &probe, |o| o != Ordering::Greater)
            } else {
                match &probe {
                    Value::Text(t) if has_wildcard(t) => {
                        matches!(v, Value::Text(x) if wildcard(t, x))
                    }
                    _ => one_kept(v, &probe, |o| o == Ordering::Equal),
                }
            };
            return Ok(if kept {
                v.clone()
            } else {
                Value::Error(ErrorKind::NA)
            });
        }
        arg => match area_of(arg) {
            Some(a) => a,
            None => return Ok(Value::Error(ErrorKind::Ref)),
        },
    };
    let width = if across {
        table.bottom - table.top + 1
    } else {
        table.right - table.left + 1
    };
    if at > f64::from(width) {
        return Ok(Value::Error(ErrorKind::Ref));
    }
    let at = at as u32 - 1;
    let line = Line {
        area: if across {
            Area {
                bottom: table.top,
                ..table
            }
        } else {
            Area {
                right: table.left,
                ..table
            }
        },
    };
    let found = if approximate {
        index::bisect(ctx, line, &probe, |o| o != Ordering::Greater)?
    } else {
        exactly(ctx, line, &probe)?
    };
    match found {
        None => Ok(Value::Error(ErrorKind::NA)),
        Some(k) => {
            let id = if across {
                (table.sheet, table.top + at, table.left + k)
            } else {
                (table.sheet, table.top + k, table.left + at)
            };
            ctx.env.read(id)
        }
    }
}

/// `VLOOKUP` ([`table_lookup`] down the first column).
pub(crate) fn vlookup(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    table_lookup(ctx, args, false)
}

/// `HLOOKUP` ([`table_lookup`] across the first row).
pub(crate) fn hlookup(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    table_lookup(ctx, args, true)
}

/// The cell `k` along a reference read from its first cell, a row where it
/// is one row high and wider than one cell, a column otherwise, as Excel
/// reads `LOOKUP`'s results; `None` past the reference's end, which Excel
/// reads and a plan orders nothing after (`criteria::sumif`'s sum range).
fn along(r: &Area, k: u32) -> Option<(usize, u32, u32)> {
    if r.top == r.bottom && r.left != r.right {
        let col = r.left.checked_add(k).filter(|&c| c <= r.right)?;
        Some((r.sheet, r.top, col))
    } else {
        let row = r.top.checked_add(k).filter(|&row| row <= r.bottom)?;
        Some((r.sheet, row, r.left))
    }
}

/// `LOOKUP(value, vector, [results])` and `LOOKUP(value, array)`: always
/// the bisection; the answer at the same position of the results, read from
/// their first cell whichever way they run, a position past their end
/// labelled as they are written; the array form searching its first row
/// when it is wider than it is tall and its first column otherwise, and
/// answering from its last row or column; none `#N/A`.
pub(crate) fn lookup(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    if !(2..=3).contains(&args.len()) {
        return Ok(Value::Error(ErrorKind::Value));
    }
    let probe = ctx.one(&args[0])?;
    if let Value::Error(k) = probe {
        return Ok(Value::Error(k));
    }
    let not_past = |o: Ordering| o != Ordering::Greater;
    let area = match &args[1] {
        Arg::Scalar(v) => {
            if !one_kept(v, &probe, not_past) {
                return Ok(Value::Error(ErrorKind::NA));
            }
            return Ok(match args.get(2) {
                None => v.clone(),
                Some(Arg::Scalar(r)) => r.clone(),
                Some(arg) => match area_of(arg) {
                    Some(r) => ctx.env.read((r.sheet, r.top, r.left))?,
                    None => Value::Error(ErrorKind::NA),
                },
            });
        }
        arg => match area_of(arg) {
            Some(a) => a,
            None => return Ok(Value::Error(ErrorKind::NA)),
        },
    };
    if let Some(results) = args.get(2) {
        let Some(line) = Line::of(area) else {
            return Ok(Value::Error(ErrorKind::NA));
        };
        let Some(k) = index::bisect(ctx, line, &probe, not_past)? else {
            return Ok(Value::Error(ErrorKind::NA));
        };
        return match results {
            Arg::Scalar(r) => Ok(if k == 0 {
                r.clone()
            } else {
                Value::Error(ErrorKind::NA)
            }),
            Arg::Area {
                areas,
                written,
                moved,
            } if areas.len() == 1 => match along(&areas[0], k) {
                Some(id) => ctx.env.read(id),
                None => Err(as_written(written, *moved)),
            },
            Arg::Area { .. } => Ok(Value::Error(ErrorKind::NA)),
        };
    }
    let (rows, cols) = (area.bottom - area.top + 1, area.right - area.left + 1);
    let across = cols > rows;
    let line = Line {
        area: if across {
            Area {
                bottom: area.top,
                ..area
            }
        } else {
            Area {
                right: area.left,
                ..area
            }
        },
    };
    let Some(k) = index::bisect(ctx, line, &probe, not_past)? else {
        return Ok(Value::Error(ErrorKind::NA));
    };
    let id = if across {
        (area.sheet, area.bottom, area.left + k)
    } else {
        (area.sheet, area.top + k, area.right)
    };
    ctx.env.read(id)
}

/// `XLOOKUP(value, vector, results, [missing], [mode], [search])`: the
/// value searched in a vector of one row or one column by [`xsearch`], and
/// the results' cells at its position, as a reference: their row where the
/// vector is a column of their height, their column where it is a row of
/// their width (any other `#VALUE!`); results of many cells, whose answer
/// is many, labelled where one value is wanted until arrays come. None
/// found, `missing`, computed only then, or `#N/A` without it.
pub(crate) fn xlookup<'e>(ctx: &Ctx<'_>, args: &'e [Expr]) -> Result<Arg<'e>, Reason> {
    if !(3..=6).contains(&args.len()) {
        return error(ErrorKind::Value);
    }
    let probe = ctx.scalar(&args[0])?;
    if let Value::Error(k) = probe {
        return error(k);
    }
    let vector = ctx.eval(&args[1])?;
    let results = ctx.eval(&args[2])?;
    let mode = match expr_number_or(ctx, args.get(4), 0.0)? {
        Ok(n) => n,
        Err(k) => return error(k),
    };
    let search = match expr_number_or(ctx, args.get(5), 1.0)? {
        Ok(n) => n,
        Err(k) => return error(k),
    };
    let modes = match modes(mode, search)? {
        Ok(m) => m,
        Err(e) => return error(e),
    };
    let (found, length, across) = match &vector {
        Arg::Scalar(v) => (xsearch_one(v, &probe, modes.0), 1, false),
        arg => {
            let Some(line) = area_of(arg).and_then(Line::of) else {
                return error(ErrorKind::Value);
            };
            let across = line.area.left != line.area.right;
            (xsearch(ctx, line, &probe, modes)?, line.len(), across)
        }
    };
    let answer_at = |k: u32| -> Result<Arg<'e>, Reason> {
        match results {
            Arg::Scalar(v) => {
                if length == 1 && k == 0 {
                    Ok(Arg::Scalar(v))
                } else {
                    error(ErrorKind::Value)
                }
            }
            Arg::Area {
                areas,
                written,
                moved,
            } => {
                let [r] = &areas[..] else {
                    return error(ErrorKind::Value);
                };
                let (rows, cols) = (r.bottom - r.top + 1, r.right - r.left + 1);
                let area = if across {
                    if cols != length {
                        return error(ErrorKind::Value);
                    }
                    Area {
                        left: r.left + k,
                        right: r.left + k,
                        ..*r
                    }
                } else {
                    if rows != length {
                        return error(ErrorKind::Value);
                    }
                    Area {
                        top: r.top + k,
                        bottom: r.top + k,
                        ..*r
                    }
                };
                Ok(Arg::Area {
                    areas: Areas::one(area),
                    written,
                    moved,
                })
            }
        }
    };
    match found {
        Some(k) => answer_at(k),
        None => match args.get(3) {
            Some(e) if !matches!(e, Expr::Empty) => ctx.eval(e),
            _ => error(ErrorKind::NA),
        },
    }
}

/// `IFERROR(value, other)`, or `IFNA` where `na` alone is caught: the
/// value, or `other` computed when the value is an error, `#N/A` alone for
/// `IFNA`; a reason is no error, and passes through. A run computes it by
/// itself (`super::vector`, its `if_error`): a change here is a change
/// there, and the machine's tests hold the two to each other.
fn or_else<'e>(ctx: &Ctx<'_>, args: &'e [Expr], na: bool) -> Result<Arg<'e>, Reason> {
    if args.len() != 2 {
        return error(ErrorKind::Value);
    }
    match ctx.scalar(&args[0])? {
        Value::Error(k) if caught(k, na) => ctx.eval(&args[1]),
        v => Ok(Arg::Scalar(v)),
    }
}

/// Whether `IFERROR` (or, `na`, `IFNA`) catches an error.
pub(crate) fn caught(k: ErrorKind, na: bool) -> bool {
    !na || k == ErrorKind::NA
}

/// `IFERROR` ([`or_else`]).
pub(crate) fn iferror<'e>(ctx: &Ctx<'_>, args: &'e [Expr]) -> Result<Arg<'e>, Reason> {
    or_else(ctx, args, false)
}

/// `IFNA` ([`or_else`]).
pub(crate) fn ifna<'e>(ctx: &Ctx<'_>, args: &'e [Expr]) -> Result<Arg<'e>, Reason> {
    or_else(ctx, args, true)
}

/// `ISERROR(value)`: whether it is one of the seven errors, never
/// propagating it.
pub(crate) fn iserror(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    if args.len() != 1 {
        return Ok(Value::Error(ErrorKind::Value));
    }
    Ok(Value::Bool(matches!(ctx.one(&args[0])?, Value::Error(_))))
}

/// `ISNA(value)`: whether it is `#N/A`, never propagating it.
pub(crate) fn isna(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    if args.len() != 1 {
        return Ok(Value::Error(ErrorKind::Value));
    }
    Ok(Value::Bool(matches!(
        ctx.one(&args[0])?,
        Value::Error(ErrorKind::NA)
    )))
}

#[cfg(test)]
mod tests {
    use crate::calc::{library, Calc, ErrorKind};
    use crate::sheet::{parse_a1_range, Cell, Content, Workbook};

    /// The table the lookups search, on a sheet `T`: ascending numbers in
    /// A1:A6 with letters beside them, texts in two cases and beyond ASCII
    /// in C, descending numbers in D, every kind with an empty cell and an
    /// error in E, numbers in no order in G, and a row of numbers over a row
    /// of their names in A8:F9. A text spelling a number is written after a
    /// `'`.
    const TABLE: &[(&str, &str)] = &[
        ("A1", "10"),
        ("A2", "20"),
        ("A3", "30"),
        ("A4", "40"),
        ("A5", "50"),
        ("A6", "60"),
        ("B1", "a"),
        ("B2", "b"),
        ("B3", "c"),
        ("B4", "d"),
        ("B5", "e"),
        ("B6", "f"),
        ("C1", "Apple"),
        ("C2", "banana"),
        ("C3", "Cherry"),
        ("C4", "apple"),
        ("C5", "date"),
        ("C6", "Éclair"),
        ("D1", "60"),
        ("D2", "50"),
        ("D3", "40"),
        ("D4", "30"),
        ("D5", "20"),
        ("D6", "10"),
        ("E1", "TRUE"),
        ("E2", "5"),
        ("E3", "'5"),
        ("E5", "#N/A"),
        ("E6", "5"),
        ("G1", "50"),
        ("G2", "10"),
        ("G3", "70"),
        ("G4", "20"),
        ("G5", "80"),
        ("G6", "30"),
        ("A8", "1"),
        ("B8", "2"),
        ("C8", "3"),
        ("D8", "4"),
        ("E8", "5"),
        ("F8", "6"),
        ("A9", "one"),
        ("B9", "two"),
        ("C9", "three"),
        ("D9", "four"),
        ("E9", "five"),
        ("F9", "six"),
    ];

    /// Each formula's value over [`TABLE`] as `Calc` computes it, spelled,
    /// the formulas down column A of a sheet `F`.
    fn computed(formulas: &[&str]) -> Vec<String> {
        let mut wb = Workbook::new();
        let t = wb.ensure_sheet("T");
        for (addr, text) in TABLE {
            let r = parse_a1_range(addr).unwrap();
            let content = if let Some(s) = text.strip_prefix('\'') {
                Content::Text(s.to_string())
            } else if let Ok(n) = text.parse::<f64>() {
                Content::Number(n)
            } else if *text == "TRUE" || *text == "FALSE" {
                Content::Bool(*text == "TRUE")
            } else if let Some(e) = ErrorKind::parse(text) {
                Content::Error(e.text().to_string())
            } else {
                Content::Text(text.to_string())
            };
            wb.sheets[t].set(r.top, r.left, Cell { content, style: 0 });
        }
        let f = wb.ensure_sheet("F");
        for (i, formula) in formulas.iter().enumerate() {
            let content = Content::Formula(formula.to_string());
            wb.sheets[f].set(i as u32 + 1, 1, Cell { content, style: 0 });
        }
        let lib = library::language();
        let calc = Calc::run(&wb, &lib);
        (0..formulas.len())
            .map(|i| match calc.computed((f, i as u32 + 1, 1)) {
                Some(c) => c.spell(),
                None => "nothing".to_string(),
            })
            .collect()
    }

    /// Each formula of the pairs against the value it is to compute.
    fn holds(cases: &[(&str, &str)]) {
        let formulas: Vec<&str> = cases.iter().map(|(f, _)| *f).collect();
        for ((formula, want), got) in cases.iter().zip(computed(&formulas)) {
            assert_eq!(got, *want, "={formula}");
        }
    }

    const NA: &str = "(error \"#N/A\")";
    const REF: &str = "(error \"#REF!\")";
    const VALUE: &str = "(error \"#VALUE!\")";

    #[test]
    fn match_finds_as_excel_documents() {
        holds(&[
            ("MATCH(30,T!A1:A6,0)", "3"),
            ("MATCH(35,T!A1:A6,1)", "3"),
            ("MATCH(35,T!A1:A6)", "3"),
            ("MATCH(5,T!A1:A6)", NA),
            ("MATCH(99,T!A1:A6)", "6"),
            ("MATCH(35,T!D1:D6,-1)", "3"),
            // Type -1 walks from the first cell, as Excel's save answers,
            // where a bisection would keep a later one.
            ("MATCH(15,T!G1:G6,-1)", "1"),
            ("MATCH(55,T!G1:G6,-1)", NA),
            ("MATCH(5,T!G1:G6,-1)", "6"),
            ("MATCH(65,T!E1:E6,-1)", NA),
            ("MATCH(5,T!E1:E6,-1)", "6"),
            ("MATCH(\"APPLE\",T!C1:C6,0)", "1"),
            ("MATCH(\"b*\",T!C1:C6,0)", "2"),
            ("MATCH(\"?pple\",T!C1:C6,0)", "1"),
            // No two kinds equal to a lookup.
            ("MATCH(\"5\",T!A1:A6,0)", NA),
            ("MATCH(5,T!E1:E6,0)", "2"),
            ("MATCH(\"5\",T!E1:E6,0)", "3"),
            ("MATCH(TRUE,T!E1:E6,0)", "1"),
            ("MATCH(30,T!A1:B6,0)", NA),
            // A type given as nothing is 0: exactly.
            ("MATCH(35,T!A1:A6,)", NA),
            ("MATCH(3,T!A8:F8,0)", "3"),
            ("MATCH(T!E5,T!A1:A6,0)", NA),
            ("MATCH(1/0,T!A1:A6,0)", "(error \"#DIV/0!\")"),
        ]);
    }

    #[test]
    fn xmatch_finds_by_each_mode_and_search() {
        holds(&[
            ("XMATCH(\"apple\",T!C1:C6)", "1"),
            ("XMATCH(\"apple\",T!C1:C6,0,-1)", "4"),
            ("XMATCH(35,T!A1:A6)", NA),
            ("XMATCH(35,T!A1:A6,1)", "4"),
            ("XMATCH(35,T!A1:A6,-1)", "3"),
            ("XMATCH(35,T!A1:A6,1,2)", "4"),
            ("XMATCH(35,T!A1:A6,-1,2)", "3"),
            ("XMATCH(35,T!D1:D6,1,-2)", "3"),
            ("XMATCH(35,T!D1:D6,-1,-2)", "4"),
            ("XMATCH(\"c*\",T!C1:C6,2)", "3"),
            ("XMATCH(\"c*\",T!C1:C6)", NA),
            // Mode 3, Excel's regular expressions, is labelled; a mode or a
            // search Excel does not take is `#VALUE!` before it.
            ("XMATCH(30,T!A1:A6,3)", "(not-computed \"match mode 3\")"),
            ("XMATCH(30,T!A1:A6,3,-2)", "(not-computed \"match mode 3\")"),
            ("XMATCH(30,T!A1:A6,3,3)", VALUE),
            ("XMATCH(30,T!A1:A6,4)", VALUE),
            ("XMATCH(30,T!A1:A6,0,3)", VALUE),
            ("XMATCH(\"c*\",T!C1:C6,2,2)", VALUE),
            ("XMATCH(30,T!A1:B6)", VALUE),
            ("XMATCH(4,T!A8:F8)", "4"),
        ]);
    }

    #[test]
    fn the_table_lookups_answer_from_the_column_or_row_asked() {
        holds(&[
            ("VLOOKUP(30,T!A1:B6,2,FALSE)", "\"c\""),
            ("VLOOKUP(35,T!A1:B6,2)", "\"c\""),
            ("VLOOKUP(35,T!A1:B6,2,TRUE)", "\"c\""),
            ("VLOOKUP(35,T!A1:B6,2,FALSE)", NA),
            // A fourth argument given as nothing is FALSE: exactly.
            ("VLOOKUP(35,T!A1:B6,2,)", NA),
            ("VLOOKUP(30,T!A1:B6,3,FALSE)", REF),
            ("VLOOKUP(30,T!A1:B6,0,FALSE)", VALUE),
            ("VLOOKUP(\"APPLE\",T!C1:D6,2,FALSE)", "60"),
            ("VLOOKUP(\"ban*\",T!C1:D6,2,FALSE)", "50"),
            ("VLOOKUP(10,T!A1:E6,5,FALSE)", "true"),
            // An empty cell found is 0 in the formula's cell.
            ("VLOOKUP(40,T!A1:E6,5,FALSE)", "0"),
            ("VLOOKUP(50,T!A1:E6,5,FALSE)", NA),
            ("HLOOKUP(3,T!A8:F9,2,FALSE)", "\"three\""),
            ("HLOOKUP(3.5,T!A8:F9,2)", "\"three\""),
            ("HLOOKUP(3,T!A8:F9,3,FALSE)", REF),
            ("HLOOKUP(0,T!A8:F9,2)", NA),
        ]);
    }

    #[test]
    fn lookup_answers_from_its_results_or_its_last_column() {
        holds(&[
            ("LOOKUP(35,T!A1:A6,T!B1:B6)", "\"c\""),
            ("LOOKUP(35,T!A1:B6)", "\"c\""),
            ("LOOKUP(5,T!A1:A6,T!B1:B6)", NA),
            ("LOOKUP(3.5,T!A8:F9)", "\"three\""),
            ("LOOKUP(9.99E+307,T!A1:A6)", "60"),
            ("LOOKUP(35,T!A1:A6,T!A9:F9)", "\"three\""),
            // Results shorter than the vector, read past their end: labelled.
            ("LOOKUP(99,T!A1:A6,T!B1:B3)", "(not-computed \"T!B1:B3\")"),
            ("LOOKUP(25,T!A1:A6,T!B1:B3)", "\"b\""),
        ]);
    }

    #[test]
    fn xlookup_answers_the_results_at_the_match_or_its_missing() {
        holds(&[
            ("XLOOKUP(30,T!A1:A6,T!B1:B6)", "\"c\""),
            ("XLOOKUP(31,T!A1:A6,T!B1:B6,\"none\")", "\"none\""),
            ("XLOOKUP(31,T!A1:A6,T!B1:B6)", NA),
            ("XLOOKUP(31,T!A1:A6,T!B1:B6,)", NA),
            ("XLOOKUP(35,T!A1:A6,T!B1:B6,,1)", "\"d\""),
            ("XLOOKUP(35,T!A1:A6,T!B1:B6,,-1,2)", "\"c\""),
            ("XLOOKUP(30,T!A1:A6,T!B1:B5)", VALUE),
            ("XLOOKUP(3,T!A8:F8,T!A9:F9)", "\"three\""),
            ("XLOOKUP(\"apple\",T!C1:C6,T!A1:A6,,0,-1)", "40"),
            ("XLOOKUP(30,T!A1:A6,T!A1:B6)", "30"),
            (
                "XLOOKUP(30,T!A1:A6,T!B1:B6,,3)",
                "(not-computed \"match mode 3\")",
            ),
        ]);
    }

    #[test]
    fn index_picks_a_cell_a_row_or_a_column() {
        holds(&[
            ("INDEX(T!A1:B6,3,2)", "\"c\""),
            ("INDEX(T!A1:A6,4)", "40"),
            ("INDEX(T!A8:F8,3)", "3"),
            ("INDEX(T!A1:B6,7,1)", REF),
            ("INDEX(T!A1:B6,1,3)", REF),
            ("INDEX(T!A1:B6,-1,1)", VALUE),
            ("INDEX(T!A1:B6,2.9,1.2)", "20"),
            ("SUM(INDEX(T!A1:B6,0,1))", "210"),
            ("SUM(INDEX(T!A1:D6,2,0))", "70"),
            ("INDEX(T!A1:A6,MATCH(40,T!A1:A6,0))", "40"),
            // A row picked where one value is wanted meets the formula's
            // column, A, in a legacy formula (KERNEL.8, slice b); a row with
            // no cell there is `#VALUE!`.
            ("INDEX(T!A1:B6,2,0)", "20"),
            ("INDEX(T!C1:D6,2,0)", VALUE),
            ("INDEX(5,1)", "5"),
            ("INDEX(5,2)", REF),
            ("INDEX(T!A1:B6,1,1,1)", "10"),
            ("INDEX(T!A1:B6,1,1,2)", REF),
            ("INDEX(T!A1:B6,\"x\")", VALUE),
        ]);
    }

    #[test]
    fn the_error_tests_read_a_miss() {
        holds(&[
            ("IFERROR(VLOOKUP(31,T!A1:B6,2,FALSE),\"none\")", "\"none\""),
            ("IFERROR(VLOOKUP(30,T!A1:B6,2,FALSE),\"none\")", "\"c\""),
            ("IFNA(1/0,2)", "(error \"#DIV/0!\")"),
            ("IFERROR(1/0,2)", "2"),
            ("IFNA(MATCH(31,T!A1:A6,0),-1)", "-1"),
            ("ISNA(MATCH(31,T!A1:A6,0))", "true"),
            ("ISERROR(T!A1)", "false"),
            ("ISERROR(T!E5)", "true"),
            ("ISNA(1/0)", "false"),
            ("ISERROR(T!E4)", "false"),
            // A reason is no error: it passes through.
            ("IFERROR(FACT(3),1)", "(not-computed \"FACT\")"),
            ("ISERROR(FACT(3))", "(not-computed \"FACT\")"),
            ("IF(ISNA(MATCH(31,T!A1:A6,0)),0,1)", "0"),
        ]);
    }

    #[test]
    fn the_criteria_count_and_sum_as_kernel_7_wrote_them() {
        holds(&[
            ("COUNTIF(T!C1:C6,\"apple\")", "2"),
            ("COUNTIF(T!A1:A6,\">25\")", "4"),
            ("COUNTIF(T!A1:A6,\"<>30\")", "5"),
            ("COUNTIF(T!E1:E6,5)", "3"),
            ("COUNTIF(T!E1:E6,\"\")", "1"),
            ("COUNTIF(T!E1:E6,\"<>\")", "5"),
            ("COUNTIF(T!E1:E6,TRUE)", "1"),
            ("COUNTIF(T!E1:E6,\"#N/A\")", "1"),
            ("COUNTIF(T!C1:C6,\"É*\")", "1"),
            ("COUNTIF(T!C1:C6,\"*a*\")", "5"),
            ("SUMIF(T!C1:C6,\"apple\",T!A1:A6)", "50"),
            ("SUMIF(T!A1:A6,\">25\")", "180"),
            ("SUMIF(T!A1:A6,\">25\",T!D1:D6)", "100"),
            // A sum range larger than the range is read in the range's shape.
            ("SUMIF(T!A1:A3,\">15\",T!D1:D6)", "90"),
            // One smaller, which Excel reads past its end: labelled.
            ("SUMIF(T!A1:A6,\">25\",T!D1)", "(not-computed \"T!D1\")"),
        ]);
    }
}
