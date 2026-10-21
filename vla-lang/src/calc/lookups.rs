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

use super::array::{self, Array, Lifted};
use super::criteria::{as_written, has_wildcard, wildcard};
use super::eval::{exact_order, Arg, Ctx, Reason};
use super::formula::Expr;
use super::graph::{Area, Areas};
use super::index::{self, Line, Seek};
use super::library::picked;
use super::order;
use super::value::{ErrorKind, Value};

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

/// The one area of a reference, `None` for a value, an array or a span of
/// sheets.
fn area_of(arg: &Arg<'_>) -> Option<Area> {
    match arg {
        Arg::Area { areas, .. } => match &areas[..] {
            [a] => Some(*a),
            _ => None,
        },
        _ => None,
    }
}

/// The vector an argument is where a lookup searches one: a reference's one
/// row or column, or an array's (KERNEL.8, slice f); `None` for any other.
fn seek_of<'a>(arg: &'a Arg<'_>) -> Option<Seek<'a>> {
    match arg {
        Arg::Array(a) if a.is_vector() => Some(Seek::Values(a.values())),
        arg => area_of(arg).and_then(Line::of).map(Seek::Line),
    }
}

/// Whether a vector runs across a row, a reference's or an array's.
fn runs_across(arg: &Arg<'_>) -> bool {
    match arg {
        Arg::Array(a) => a.rows() == 1 && a.columns() > 1,
        arg => area_of(arg).is_some_and(|a| a.left != a.right),
    }
}

/// Whether one value given where a vector is wanted, as Excel reads a value
/// there, an array of one, is kept by a search's rule: of the probe's kind
/// and `keep` keeping how it compares with the probe; two texts in Excel's
/// order of texts, labelled where a character beyond its known order
/// decides an answer that turns on it, which an exact match's never does.
fn one_kept(v: &Value, probe: &Value, keep: impl Fn(Ordering) -> bool) -> Result<bool, Reason> {
    let kind = |x: &Value| match x {
        Value::Number(_) => Some(0),
        Value::Text(_) => Some(1),
        Value::Bool(_) => Some(2),
        _ => None,
    };
    if kind(v).is_none() || kind(v) != kind(probe) {
        return Ok(false);
    }
    let o = match (v, probe) {
        (Value::Number(a), Value::Number(b)) => exact_order(*a, *b),
        (Value::Text(a), Value::Text(b)) => {
            let c = order::compare(a, b);
            if !c.known
                && c.order != Ordering::Equal
                && keep(Ordering::Less) != keep(Ordering::Greater)
            {
                return Err(order::unknown());
            }
            c.order
        }
        (Value::Bool(a), Value::Bool(b)) => a.cmp(b),
        _ => return Ok(false),
    };
    Ok(keep(o))
}

/// An exact match as `MATCH` 0, `VLOOKUP` and `HLOOKUP` take it: a text
/// holding `*`, `?` or `~` a pattern, any other probe by its key.
fn exactly(ctx: &Ctx<'_>, seek: Seek<'_>, probe: &Value) -> Result<Option<u32>, Reason> {
    match probe {
        Value::Text(t) if has_wildcard(t) => index::pattern(ctx, seek, t, false),
        _ => index::exact(ctx, seek, probe, false),
    }
}

/// The part of an array `INDEX` picks (`library::picked`'s offsets): its one
/// value, or the array of its rows and columns there (KERNEL.8, slice f).
fn sub_array(a: &Array, (t, l, b, r): (u32, u32, u32, u32)) -> Arg<'static> {
    if t == b && l == r {
        return Arg::Scalar(a.get(t, l).clone());
    }
    let mut values = Vec::with_capacity(((b - t + 1) * (r - l + 1)) as usize);
    for row in t..=b {
        for col in l..=r {
            values.push(a.get(row, col).clone());
        }
    }
    Arg::Array(Array::new(b - t + 1, r - l + 1, values))
}

/// `INDEX(reference, row, [column], [area])`: the cell of the reference's
/// area at the row and the column counted from its corner, as a reference
/// (`library::picked`): a reference one row high takes a single number as
/// its column; a row of 0 is a whole column, a column of 0 a whole row; a
/// number read down to a whole one; outside the area `#REF!`, below 0
/// `#VALUE!`. A value given as the reference is an array of one, and an
/// array is picked from as a reference is, its values answered (KERNEL.8,
/// slice f): the reference's expression is read as arrays, and a row or a
/// column given as an array lifts the pick, each place's one value. Only
/// area 1 of the one area a reference names, a union being no reference
/// this version reads.
pub(crate) fn index<'e>(ctx: &Ctx<'_>, args: &'e [Expr]) -> Result<Arg<'e>, Reason> {
    if !(2..=4).contains(&args.len()) {
        return error(ErrorKind::Value);
    }
    let reference = ctx.as_arrays(|| ctx.eval(&args[0]))?;
    let row = ctx.operand(&args[1])?;
    let column = match args.get(2) {
        None => None,
        Some(e) => Some(ctx.operand(e)?),
    };
    if let Some(e) = args.get(3) {
        match expr_number_or(ctx, Some(e), 1.0)? {
            Ok(n) if n.trunc() == 1.0 => {}
            Ok(_) => return error(ErrorKind::Ref),
            Err(k) => return error(k),
        }
    }
    let many = matches!(row, Lifted::Many(_)) || matches!(column, Some(Lifted::Many(_)));
    if many {
        // Each place's pick, one value each.
        let column = column.unwrap_or(Lifted::One(Value::Empty));
        let left_out = args.get(2).is_none();
        let at = array::lift2(&row, &column, |r, c| {
            let (r, c) = match (r.number(), c.number()) {
                (Ok(r), Ok(c)) => (r, c),
                (Err(k), _) | (_, Err(k)) => return Ok(Value::Error(k)),
            };
            let picked = pick(&reference, r, (!left_out).then_some(c))?;
            ctx.one(&picked)
        })?;
        return Ok(Arg::Array(at));
    }
    let row = match row {
        Lifted::One(v) => match v.number() {
            Ok(n) => n,
            Err(k) => return error(k),
        },
        Lifted::Many(_) => unreachable!("a lift answered above"),
    };
    let column = match column {
        None => None,
        Some(Lifted::One(v)) => match v.number() {
            Ok(n) => Some(n),
            Err(k) => return error(k),
        },
        Some(Lifted::Many(_)) => unreachable!("a lift answered above"),
    };
    pick(&reference, row, column)
}

/// [`index`]'s pick of a reference, a value or an array at one row and
/// column.
fn pick<'e>(reference: &Arg<'e>, row: f64, column: Option<f64>) -> Result<Arg<'e>, Reason> {
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
                    written: written.clone(),
                    moved: *moved,
                }),
                Err(k) => error(k),
            }
        }
        Arg::Array(a) => match picked(a.rows(), a.columns(), row, column) {
            Ok(p) => Ok(sub_array(a, p)),
            Err(k) => error(k),
        },
        Arg::Scalar(v) => match picked(1, 1, row, column) {
            Ok(_) => Ok(Arg::Scalar(v.clone())),
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
        arg => match pick(arg, row, column)? {
            Arg::Scalar(v) => Ok(v),
            other => ctx.one(&other),
        },
    }
}

/// `MATCH(value, vector, [type])`: the position, from 1, of the value in a
/// vector of one row or one column, a reference's or an array's (any other
/// `#N/A`): type 0 exactly, a text holding a wildcard a pattern; type 1, or
/// none, the bisection's largest not past the value; type -1 the smallest
/// not under it on a descending vector, by the walk from its first cell
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
                    _ => one_kept(v, &probe, |o| o == Ordering::Equal)?,
                }
            } else if kind > 0.0 {
                one_kept(v, &probe, |o| o != Ordering::Greater)?
            } else {
                one_kept(v, &probe, |o| o != Ordering::Less)?
            };
            kept.then_some(0)
        }
        arg => {
            let Some(seek) = seek_of(arg) else {
                return Ok(Value::Error(ErrorKind::NA));
            };
            if kind == 0.0 {
                exactly(ctx, seek, &probe)?
            } else if kind > 0.0 {
                index::bisect(ctx, seek, &probe, |o| o != Ordering::Greater)?
            } else {
                index::descending(ctx, seek, &probe)?
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
    seek: Seek<'_>,
    probe: &Value,
    (mode, search): (i8, i8),
) -> Result<Option<u32>, Reason> {
    if search.abs() == 1 {
        let from_end = search < 0;
        return match (mode, probe) {
            (2, Value::Text(t)) if has_wildcard(t) => index::pattern(ctx, seek, t, from_end),
            (-1, _) => index::nearest(ctx, seek, probe, false, from_end),
            (1, _) => index::nearest(ctx, seek, probe, true, from_end),
            _ => index::exact(ctx, seek, probe, from_end),
        };
    }
    let equal = |k: Option<u32>| -> Result<Option<u32>, Reason> {
        let Some(k) = k else {
            return Ok(None);
        };
        let v = seek.read(ctx, k)?;
        Ok(one_kept(&v, probe, |o| o == Ordering::Equal)?.then_some(k))
    };
    if search > 0 {
        match mode {
            -1 => index::bisect(ctx, seek, probe, |o| o != Ordering::Greater),
            1 => {
                let under = index::bisect(ctx, seek, probe, |o| o == Ordering::Less)?;
                index::next_of_kind(ctx, seek, probe, under)
            }
            _ => equal(index::bisect(ctx, seek, probe, |o| o != Ordering::Greater)?),
        }
    } else {
        match mode {
            1 => index::bisect(ctx, seek, probe, |o| o != Ordering::Less),
            -1 => {
                let over = index::bisect(ctx, seek, probe, |o| o == Ordering::Greater)?;
                index::next_of_kind(ctx, seek, probe, over)
            }
            _ => equal(index::bisect(ctx, seek, probe, |o| o != Ordering::Less)?),
        }
    }
}

/// One value given where `XMATCH`'s or `XLOOKUP`'s vector is wanted, an
/// array of one: kept by the mode's rule.
fn xsearch_one(v: &Value, probe: &Value, mode: i8) -> Result<Option<u32>, Reason> {
    let kept = match mode {
        -1 => one_kept(v, probe, |o| o != Ordering::Greater)?,
        1 => one_kept(v, probe, |o| o != Ordering::Less)?,
        2 => match probe {
            Value::Text(t) => matches!(v, Value::Text(x) if wildcard(t, x)),
            _ => one_kept(v, probe, |o| o == Ordering::Equal)?,
        },
        _ => one_kept(v, probe, |o| o == Ordering::Equal)?,
    };
    Ok(kept.then_some(0))
}

/// `XMATCH(value, vector, [mode], [search])`: the position, from 1, of the
/// value in a vector of one row or one column, a reference's or an array's
/// (any other `#VALUE!`), by [`xsearch`]; none `#N/A`.
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
        Arg::Scalar(v) => xsearch_one(v, &probe, modes.0)?,
        arg => {
            let Some(seek) = seek_of(arg) else {
                return Ok(Value::Error(ErrorKind::Value));
            };
            xsearch(ctx, seek, &probe, modes)?
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
/// holds as 0. A table may be an array, searched and answered as a
/// reference is (KERNEL.8, slice f).
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
    let search = |seek: Seek<'_>| -> Result<Option<u32>, Reason> {
        if approximate {
            index::bisect(ctx, seek, &probe, |o| o != Ordering::Greater)
        } else {
            exactly(ctx, seek, &probe)
        }
    };
    let table = match &args[1] {
        Arg::Scalar(v) => {
            // A value given as the table, an array of one.
            if at > 1.0 {
                return Ok(Value::Error(ErrorKind::Ref));
            }
            let kept = if approximate {
                one_kept(v, &probe, |o| o != Ordering::Greater)?
            } else {
                match &probe {
                    Value::Text(t) if has_wildcard(t) => {
                        matches!(v, Value::Text(x) if wildcard(t, x))
                    }
                    _ => one_kept(v, &probe, |o| o == Ordering::Equal)?,
                }
            };
            return Ok(if kept {
                v.clone()
            } else {
                Value::Error(ErrorKind::NA)
            });
        }
        Arg::Array(a) => {
            let width = if across { a.rows() } else { a.columns() };
            if at > f64::from(width) {
                return Ok(Value::Error(ErrorKind::Ref));
            }
            let at = at as u32 - 1;
            let first: Vec<Value> = if across {
                (0..a.columns()).map(|c| a.get(0, c).clone()).collect()
            } else {
                (0..a.rows()).map(|r| a.get(r, 0).clone()).collect()
            };
            return Ok(match search(Seek::Values(&first))? {
                None => Value::Error(ErrorKind::NA),
                Some(k) if across => a.get(at, k).clone(),
                Some(k) => a.get(k, at).clone(),
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
    match search(Seek::Line(line))? {
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

/// The value `k` along an array as [`along`] reads a reference: its first
/// row where it is one row high and wider than one value, its first column
/// otherwise; `None` past its end.
fn along_array(a: &Array, k: u32) -> Option<&Value> {
    if a.rows() == 1 && a.columns() > 1 {
        (k < a.columns()).then(|| a.get(0, k))
    } else {
        (k < a.rows()).then(|| a.get(k, 0))
    }
}

/// `LOOKUP(value, vector, [results])` and `LOOKUP(value, array)`: always
/// the bisection; the answer at the same position of the results, read from
/// their first cell whichever way they run, a position past their end
/// labelled as they are written; the array form searching its first row
/// when it is wider than it is tall and its first column otherwise, and
/// answering from its last row or column; none `#N/A`. A vector, results
/// or an array form may be arrays, read as references are (KERNEL.8, slice
/// f), a position past an array's end `#N/A`.
pub(crate) fn lookup(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    if !(2..=3).contains(&args.len()) {
        return Ok(Value::Error(ErrorKind::Value));
    }
    let probe = ctx.one(&args[0])?;
    if let Value::Error(k) = probe {
        return Ok(Value::Error(k));
    }
    let not_past = |o: Ordering| o != Ordering::Greater;
    // The results at a position found.
    let answer = |k: u32| -> Result<Value, Reason> {
        match args.get(2) {
            Some(Arg::Scalar(r)) => Ok(if k == 0 {
                r.clone()
            } else {
                Value::Error(ErrorKind::NA)
            }),
            Some(Arg::Area {
                areas,
                written,
                moved,
            }) if areas.len() == 1 => match along(&areas[0], k) {
                Some(id) => ctx.env.read(id),
                None => Err(as_written(written, *moved)),
            },
            Some(Arg::Array(r)) => Ok(along_array(r, k)
                .cloned()
                .unwrap_or(Value::Error(ErrorKind::NA))),
            _ => Ok(Value::Error(ErrorKind::NA)),
        }
    };
    match &args[1] {
        Arg::Scalar(v) => {
            if !one_kept(v, &probe, not_past)? {
                return Ok(Value::Error(ErrorKind::NA));
            }
            return match args.get(2) {
                None => Ok(v.clone()),
                Some(Arg::Area { .. }) => match area_of(&args[2]) {
                    Some(r) => ctx.env.read((r.sheet, r.top, r.left)),
                    None => Ok(Value::Error(ErrorKind::NA)),
                },
                Some(_) => answer(0),
            };
        }
        Arg::Array(a) => {
            if args.get(2).is_some() {
                let Some(seek) = seek_of(&args[1]) else {
                    return Ok(Value::Error(ErrorKind::NA));
                };
                return match index::bisect(ctx, seek, &probe, not_past)? {
                    Some(k) => answer(k),
                    None => Ok(Value::Error(ErrorKind::NA)),
                };
            }
            // The array form: its first row or column searched, its last
            // answering.
            let across = a.columns() > a.rows();
            let first: Vec<Value> = if across {
                (0..a.columns()).map(|c| a.get(0, c).clone()).collect()
            } else {
                (0..a.rows()).map(|r| a.get(r, 0).clone()).collect()
            };
            return Ok(
                match index::bisect(ctx, Seek::Values(&first), &probe, not_past)? {
                    None => Value::Error(ErrorKind::NA),
                    Some(k) if across => a.get(a.rows() - 1, k).clone(),
                    Some(k) => a.get(k, a.columns() - 1).clone(),
                },
            );
        }
        _ => {}
    }
    let Some(area) = area_of(&args[1]) else {
        return Ok(Value::Error(ErrorKind::NA));
    };
    if args.get(2).is_some() {
        let Some(line) = Line::of(area) else {
            return Ok(Value::Error(ErrorKind::NA));
        };
        return match index::bisect(ctx, Seek::Line(line), &probe, not_past)? {
            Some(k) => answer(k),
            None => Ok(Value::Error(ErrorKind::NA)),
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
    let Some(k) = index::bisect(ctx, Seek::Line(line), &probe, not_past)? else {
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
/// is many, an array where the formula is read as arrays and labelled where
/// one value is wanted. None found, `missing`, computed only then, or `#N/A`
/// without it. The vector and the results are read as arrays and may be
/// arrays (KERNEL.8, slice f), and a value sought given as an array lifts
/// the search, each place's one value.
pub(crate) fn xlookup<'e>(ctx: &Ctx<'_>, args: &'e [Expr]) -> Result<Arg<'e>, Reason> {
    if !(3..=6).contains(&args.len()) {
        return error(ErrorKind::Value);
    }
    let probe = ctx.operand(&args[0])?;
    if let Lifted::One(Value::Error(k)) = probe {
        return error(k);
    }
    let vector = ctx.as_arrays(|| ctx.eval(&args[1]))?;
    let results = ctx.as_arrays(|| ctx.eval(&args[2]))?;
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
    match probe {
        Lifted::One(probe) => xlooked(ctx, args, &probe, &vector, &results, modes),
        Lifted::Many(probes) => {
            let answers = array::lift1(&probes, |p| {
                if let Value::Error(k) = p {
                    return Ok(Value::Error(*k));
                }
                let answer = xlooked(ctx, args, p, &vector, &results, modes)?;
                ctx.one(&answer)
            })?;
            Ok(Arg::Array(answers))
        }
    }
}

/// [`xlookup`]'s search of one value and its answer.
fn xlooked<'e>(
    ctx: &Ctx<'_>,
    args: &'e [Expr],
    probe: &Value,
    vector: &Arg<'_>,
    results: &Arg<'e>,
    modes: (i8, i8),
) -> Result<Arg<'e>, Reason> {
    let (found, length, across) = match vector {
        Arg::Scalar(v) => (xsearch_one(v, probe, modes.0)?, 1, false),
        arg => {
            let Some(seek) = seek_of(arg) else {
                return error(ErrorKind::Value);
            };
            (
                xsearch(ctx, seek, probe, modes)?,
                seek.len(),
                runs_across(arg),
            )
        }
    };
    let answer_at = |k: u32| -> Result<Arg<'e>, Reason> {
        match results {
            Arg::Scalar(v) => {
                if length == 1 && k == 0 {
                    Ok(Arg::Scalar(v.clone()))
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
                    written: written.clone(),
                    moved: *moved,
                })
            }
            Arg::Array(r) => {
                if across {
                    if r.columns() != length {
                        return error(ErrorKind::Value);
                    }
                    Ok(sub_array(r, (0, k, r.rows() - 1, k)))
                } else {
                    if r.rows() != length {
                        return error(ErrorKind::Value);
                    }
                    Ok(sub_array(r, (k, 0, k, r.columns() - 1)))
                }
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
/// there, and the machine's tests hold the two to each other. In a legacy
/// formula both are values forced, Excel's class for them (`Param::Scalar`,
/// KERNEL.8, slice f): an array's first value, a range met at the cell even
/// inside an array parameter, as Excel's own save of the arrays' fixture
/// computes `IFERROR({1,#N/A,3},0)` as 1. Where the formula reads arrays by
/// its kind, a value that is an array is caught at each of its places,
/// `other` read there, and a place not caught leaves `other` free to fall
/// short, as `IF` leaves the branch it does not take.
fn or_else<'e>(ctx: &Ctx<'_>, args: &'e [Expr], na: bool) -> Result<Arg<'e>, Reason> {
    if args.len() != 2 {
        return error(ErrorKind::Value);
    }
    if !ctx.reads_arrays_by_kind() {
        return Ok(Arg::Scalar(match ctx.forced(&args[0])? {
            Value::Error(k) if caught(k, na) => ctx.forced(&args[1])?,
            v => v,
        }));
    }
    match ctx.operand(&args[0])? {
        Lifted::One(Value::Error(k)) if caught(k, na) => ctx.eval(&args[1]),
        Lifted::One(v) => Ok(Arg::Scalar(v)),
        Lifted::Many(a) => {
            let value = Lifted::Many(a);
            let other = ctx.operand(&args[1])?;
            let short = || Value::Error(ErrorKind::NA);
            let lifted = array::lift_by(&[&value, &other], |r, c| {
                Ok(match value.at(r, c) {
                    None => short(),
                    Some(Value::Error(k)) if caught(*k, na) => {
                        other.at(r, c).cloned().unwrap_or_else(short)
                    }
                    Some(v) => v.clone(),
                })
            })?;
            Ok(Arg::Array(lifted))
        }
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

/// `ISERROR(value)`: whether it is one of the errors, never propagating it.
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
    /// error in E, numbers in no order in G, a row of numbers over a row of
    /// their names in A8:F9, and in I texts Excel's order decides, one
    /// beyond its table (`ø`) among them. A text spelling a number is written
    /// after a `'`.
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
        ("I1", "a"),
        ("I2", "ø"),
        ("I3", "c"),
        ("I4", "co-op"),
        ("I5", "coop"),
        ("I6", "Straße"),
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

    #[test]
    fn a_search_orders_texts_as_excel_does_and_labels_one_beyond_it() {
        // KERNEL.8, slice c: Excel's order of texts in a search and a
        // criterion, `ß` and `ss` one, and a search that a character beyond
        // its table decides labelled, by the scan and by the index alike:
        // each probe asked twice, the second answered from the vector's
        // index, built at its second probe.
        const ORDER: &str = "(not-computed \"a character's order\")";
        holds(&[
            ("XMATCH(\"b\",T!I1:I3,1)", ORDER),
            ("XMATCH(\"b\",T!I1:I3,1)", ORDER),
            ("XMATCH(\"b\",T!I1:I3,-1)", ORDER),
            ("XMATCH(\"b\",T!I1:I3,-1)", ORDER),
            // An exact match answers, whatever the order beside it.
            ("XMATCH(\"Ø\",T!I1:I3,1)", "2"),
            ("XMATCH(\"Ø\",T!I1:I3,1)", "2"),
            ("XMATCH(\"ø\",T!I1:I3,0)", "2"),
            ("MATCH(\"b\",T!I1:I3,1)", ORDER),
            // "~" stands before "a" in Excel's order, after it in code points.
            ("MATCH(\"~\",T!I1,1)", NA),
            ("MATCH(\"co-op\",T!I4:I5,0)", "1"),
            ("COUNTIF(T!I1:I6,\"<b\")", ORDER),
            ("COUNTIF(T!I1:I6,\"strasse\")", "1"),
            ("COUNTIF(T!I1:I6,\"strasse\")", "1"),
            ("VLOOKUP(\"STRASSE\",T!I1:I6,1,FALSE)", "\"Straße\""),
            // One value where a vector is wanted, by size and exactly.
            ("XMATCH(\"b\",\"ø\",1)", ORDER),
            ("XMATCH(\"ø\",\"Ø\",1)", "1"),
            ("MATCH(\"b\",\"ø\",0)", NA),
        ]);
    }
}
