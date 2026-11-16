//! The functions that take a function (KERNEL.29, 2026-11-15): `MAP`,
//! `REDUCE`, `SCAN`, `BYROW`, `BYCOL` and `MAKEARRAY`, and `ISOMITTED`.
//!
//! Each takes its function as a value its last argument answers, a `LAMBDA`
//! written there, a name bound to one, or a named function by its name, and
//! calls it once a place (`Ctx::call_function`, an application its array
//! bounds and so not counted against a cell's calls). What it hands the
//! function: of a reference, each cell as a reference to that cell, a row or
//! a column as a reference to it, so that `ROW` and `ISREF` read it as they
//! read any reference; of an array, each value, or a row or a column as an
//! array; one value as itself. What the function answers at a place must be
//! one value, a reference to one cell read, an array of one its value, and
//! anything else is `#CALC!` there, as Excel's page of `BYROW` says of its
//! own. A function's value that is no function, or one whose parameters are
//! not as many as the function hands it, is `#VALUE!`, Excel's "Incorrect
//! Parameters". Each answers an array, which a cell shows by spilling and an
//! aggregate reads whole (`KERNEL.55`), and is made after it is sized
//! (`super::array::room`).
//!
//! Where Excel's pages are silent the slice's fixture holds them to its save
//! (`scripts/recalc/lambda`): a range walked row by row, as every walk of a
//! reference goes; `MAP` over arrays of unlike shapes, each place past an
//! array's shape `#N/A`, an array of one value standing at every place, as a
//! lift is; a `MAKEARRAY` count not a whole number `#VALUE!`, where
//! `SEQUENCE` reads its own down.

use std::borrow::Cow;

use super::array::{self, Array};
use super::arrays::omitted;
use super::eval::{arrays, Arg, Ctx, Function, Library, Reason};
use super::formula::Expr;
use super::graph::{Area, Areas};
use super::value::{ErrorKind, Value};

/// Register the functions that take a function.
pub(crate) fn register(lib: &mut Library) {
    lib.register_arrayed("MAP", map, arrays);
    lib.register_arrayed("REDUCE", reduce, arrays);
    lib.register_arrayed("SCAN", scan, arrays);
    lib.register_arrayed("BYROW", byrow, arrays);
    lib.register_arrayed("BYCOL", bycol, arrays);
    lib.register_arrayed("MAKEARRAY", makearray, super::eval::values);
    lib.register_form("ISOMITTED", isomitted);
}

fn error(k: ErrorKind) -> Result<Arg<'static>, Reason> {
    Ok(Arg::Scalar(Value::Error(k)))
}

/// The function a helper's last argument answers, taking `params` arguments:
/// `Err` the answer where it is none, or takes another count.
fn function(arg: Option<&Arg<'_>>, params: usize) -> Result<Function, ErrorKind> {
    match arg {
        Some(Arg::Function(f)) if f.parameters() == params => Ok(f.clone()),
        Some(Arg::Scalar(Value::Error(k))) => Err(*k),
        _ => Err(ErrorKind::Value),
    }
}

/// An argument a helper walks: a reference's one area, an array, or one
/// value; `#VALUE!` for a span of sheets or a function's value.
enum Walked<'a> {
    Area(Area),
    Array(&'a Array),
    One(&'a Value),
}

impl Walked<'_> {
    fn of<'a>(arg: &'a Arg<'_>) -> Result<Walked<'a>, ErrorKind> {
        match arg {
            Arg::Area { areas, .. } => match &areas[..] {
                [a] => Ok(Walked::Area(*a)),
                _ => Err(ErrorKind::Value),
            },
            Arg::Array(a) => Ok(Walked::Array(a)),
            Arg::Scalar(Value::Error(k)) => Err(*k),
            Arg::Scalar(v) => Ok(Walked::One(v)),
            Arg::Function(_) => Err(ErrorKind::Value),
        }
    }

    fn shape(&self) -> (u32, u32) {
        match self {
            Walked::Area(a) => (a.bottom - a.top + 1, a.right - a.left + 1),
            Walked::Array(a) => (a.rows(), a.columns()),
            Walked::One(_) => (1, 1),
        }
    }

    /// The place at a row and a column from 0, as the function is handed
    /// it: a cell as a reference to it, an array's value, one value at every
    /// place; `None` past its shape.
    fn at(&self, ctx: &Ctx<'_>, row: u32, col: u32) -> Option<Arg<'static>> {
        match self {
            Walked::Area(a) => {
                let (rows, cols) = self.shape();
                (row < rows && col < cols).then(|| {
                    reference(
                        ctx,
                        Area {
                            sheet: a.sheet,
                            top: a.top + row,
                            left: a.left + col,
                            bottom: a.top + row,
                            right: a.left + col,
                        },
                    )
                })
            }
            Walked::Array(a) if a.is_single() => Some(Arg::Scalar(a.get(0, 0).clone())),
            Walked::Array(a) => a.at(row, col).map(|v| Arg::Scalar(v.clone())),
            Walked::One(v) => Some(Arg::Scalar((*v).clone())),
        }
    }

    /// The row from 0, or the column, as the function is handed it: a
    /// reference to it, or an array of it.
    fn line(&self, ctx: &Ctx<'_>, k: u32, across: bool) -> Arg<'static> {
        match self {
            Walked::Area(a) => reference(
                ctx,
                if across {
                    Area {
                        left: a.left + k,
                        right: a.left + k,
                        ..*a
                    }
                } else {
                    Area {
                        top: a.top + k,
                        bottom: a.top + k,
                        ..*a
                    }
                },
            ),
            Walked::Array(a) => {
                let (rows, cols) = (a.rows(), a.columns());
                Arg::Array(if across {
                    Array::new(rows, 1, (0..rows).map(|r| a.get(r, k).clone()).collect())
                } else {
                    Array::new(1, cols, (0..cols).map(|c| a.get(k, c).clone()).collect())
                })
            }
            Walked::One(v) => Arg::Scalar((*v).clone()),
        }
    }
}

/// A reference to an area, its label the area as written, its sheet named
/// where it is not the cell's own.
fn reference(ctx: &Ctx<'_>, a: Area) -> Arg<'static> {
    let mut written = crate::sheet::A1Range {
        top: a.top,
        left: a.left,
        bottom: a.bottom,
        right: a.right,
    }
    .text();
    if a.sheet != ctx.env.here().0 {
        if let Some(name) = ctx.env.sheet_name(a.sheet) {
            written = format!("{}!{written}", crate::rows::sheet_prefix(name));
        }
    }
    Arg::Area {
        areas: Areas::one(a),
        written: Cow::Owned(written),
        moved: (0, 0),
    }
}

/// What the function answered at a place, as one value: a reference to one
/// cell read, an array of one its value, any other answer `#CALC!`.
fn one_value(ctx: &Ctx<'_>, arg: &Arg<'_>) -> Result<Value, Reason> {
    match arg {
        Arg::Scalar(v) => Ok(v.clone()),
        Arg::Array(a) if a.is_single() => Ok(a.get(0, 0).clone()),
        Arg::Area { areas, .. } => match &areas[..] {
            [a] if a.is_cell() => ctx.env.read((a.sheet, a.top, a.left)),
            _ => Ok(Value::Error(ErrorKind::Calc)),
        },
        _ => Ok(Value::Error(ErrorKind::Calc)),
    }
}

/// An array of a shape made place by place, row by row; one value for an
/// array of one.
fn made(
    rows: u32,
    cols: u32,
    mut f: impl FnMut(u32, u32) -> Result<Value, Reason>,
) -> Result<Arg<'static>, Reason> {
    let (rows, cols) = match array::room(u64::from(rows), u64::from(cols))? {
        Ok(rc) => rc,
        Err(k) => return error(k),
    };
    let mut values = Vec::with_capacity((rows as usize) * (cols as usize));
    for r in 0..rows {
        for c in 0..cols {
            values.push(f(r, c)?);
        }
    }
    Ok(if values.len() == 1 {
        Arg::Scalar(values.pop().expect("one value"))
    } else {
        Arg::Array(Array::new(rows, cols, values))
    })
}

/// `MAP(array, ..., lambda)`: the function at each place of its arrays, each
/// array's value there an argument, an array of their shape; a place past an
/// array's shape `#N/A`.
fn map(ctx: &Ctx<'_>, args: &[Arg<'_>]) -> Result<Arg<'static>, Reason> {
    if args.len() < 2 {
        return error(ErrorKind::Value);
    }
    let (arrays, last) = args.split_at(args.len() - 1);
    let f = match function(last.first(), arrays.len()) {
        Ok(f) => f,
        Err(k) => return error(k),
    };
    let mut walked = Vec::with_capacity(arrays.len());
    for a in arrays {
        match Walked::of(a) {
            Ok(w) => walked.push(w),
            Err(k) => return error(k),
        }
    }
    let (rows, cols) = array::shape_of(walked.iter().map(Walked::shape));
    made(rows, cols, |r, c| {
        let mut call = Vec::with_capacity(walked.len());
        for w in &walked {
            match w.at(ctx, r, c) {
                Some(v) => call.push(Some(v)),
                None => return Ok(Value::Error(ErrorKind::NA)),
            }
        }
        let answer = ctx.call_function(&f, call, false)?;
        one_value(ctx, &answer)
    })
}

/// `REDUCE([initial], array, lambda(accumulator, value))`: the function
/// folded over the array's places row by row, its answer the accumulator
/// handed to the next; with no initial value, the array's first value the
/// accumulator at the start, as Excel's page says.
fn reduce(ctx: &Ctx<'_>, args: &[Arg<'_>]) -> Result<Arg<'static>, Reason> {
    fold(ctx, args, false)
}

/// `SCAN([initial], array, lambda(accumulator, value))`: [`reduce`]'s
/// accumulator at each place, an array of the array's shape.
fn scan(ctx: &Ctx<'_>, args: &[Arg<'_>]) -> Result<Arg<'static>, Reason> {
    fold(ctx, args, true)
}

fn fold(ctx: &Ctx<'_>, args: &[Arg<'_>], scanned: bool) -> Result<Arg<'static>, Reason> {
    if args.len() != 3 {
        return error(ErrorKind::Value);
    }
    let f = match function(args.get(2), 2) {
        Ok(f) => f,
        Err(k) => return error(k),
    };
    let walked = match Walked::of(&args[1]) {
        Ok(w) => w,
        Err(k) => return error(k),
    };
    let (rows, cols) = walked.shape();
    let mut places = (0..rows).flat_map(|r| (0..cols).map(move |c| (r, c)));
    let mut acc: Arg<'static> = if omitted(args.first()) {
        let Some((r, c)) = places.next() else {
            return error(ErrorKind::Calc);
        };
        let first = walked.at(ctx, r, c).expect("a place of its shape");
        // The first value the start, a place of the scan's answer.
        Arg::Scalar(one_value(ctx, &first)?)
    } else {
        args[0].clone().into_owned()
    };
    let mut scan: Vec<Value> = Vec::new();
    if scanned {
        scan.reserve((rows as usize) * (cols as usize));
        if omitted(args.first()) {
            scan.push(one_value(ctx, &acc)?);
        }
    }
    for (r, c) in places {
        let value = walked.at(ctx, r, c).expect("a place of its shape");
        acc = ctx.call_function(&f, vec![Some(acc), Some(value)], false)?;
        if scanned {
            scan.push(one_value(ctx, &acc)?);
        }
    }
    if !scanned {
        return Ok(acc);
    }
    let mut at = scan.into_iter();
    made(rows, cols, |_, _| {
        Ok(at.next().unwrap_or(Value::Error(ErrorKind::NA)))
    })
}

/// `BYROW(array, lambda(row))`: the function at each row, a column of its
/// answers.
fn byrow(ctx: &Ctx<'_>, args: &[Arg<'_>]) -> Result<Arg<'static>, Reason> {
    by_line(ctx, args, false)
}

/// `BYCOL(array, lambda(column))`: the function at each column, a row of its
/// answers.
fn bycol(ctx: &Ctx<'_>, args: &[Arg<'_>]) -> Result<Arg<'static>, Reason> {
    by_line(ctx, args, true)
}

fn by_line(ctx: &Ctx<'_>, args: &[Arg<'_>], across: bool) -> Result<Arg<'static>, Reason> {
    if args.len() != 2 {
        return error(ErrorKind::Value);
    }
    let f = match function(args.get(1), 1) {
        Ok(f) => f,
        Err(k) => return error(k),
    };
    let walked = match Walked::of(&args[0]) {
        Ok(w) => w,
        Err(k) => return error(k),
    };
    let (rows, cols) = walked.shape();
    let (lines, shape) = if across {
        (cols, (1, cols))
    } else {
        (rows, (rows, 1))
    };
    made(shape.0, shape.1, |r, c| {
        let k = if across { c } else { r };
        debug_assert!(k < lines);
        let line = walked.line(ctx, k, across);
        let answer = ctx.call_function(&f, vec![Some(line)], false)?;
        one_value(ctx, &answer)
    })
}

/// `MAKEARRAY(rows, cols, lambda(row, col))`: the function at each place,
/// handed its row and its column from 1; a count below 1 or no number
/// `#VALUE!`, as Excel's page says, and one not a whole number too, as its
/// save says (`MAKEARRAY(2.7,2,...)`), where `SEQUENCE` reads its own down.
fn makearray(ctx: &Ctx<'_>, args: &[Arg<'_>]) -> Result<Arg<'static>, Reason> {
    if args.len() != 3 {
        return error(ErrorKind::Value);
    }
    let f = match function(args.get(2), 2) {
        Ok(f) => f,
        Err(k) => return error(k),
    };
    let count = |a: &Arg<'_>| -> Result<Result<u32, ErrorKind>, Reason> {
        Ok(match ctx.one(a)? {
            Value::Number(n) if n.is_finite() && n >= 1.0 && n.fract() == 0.0 => {
                Ok(n.min(f64::from(u32::MAX)) as u32)
            }
            Value::Error(k) => Err(k),
            _ => Err(ErrorKind::Value),
        })
    };
    let rows = match count(&args[0])? {
        Ok(n) => n,
        Err(k) => return error(k),
    };
    let cols = match count(&args[1])? {
        Ok(n) => n,
        Err(k) => return error(k),
    };
    made(rows, cols, |r, c| {
        let at = vec![
            Some(Arg::Scalar(Value::Number(f64::from(r + 1)))),
            Some(Arg::Scalar(Value::Number(f64::from(c + 1)))),
        ];
        let answer = ctx.call_function(&f, at, false)?;
        one_value(ctx, &answer)
    })
}

/// `ISOMITTED(argument)`: whether a name a `LAMBDA` takes was given no
/// argument, `LAMBDA(x,y,ISOMITTED(y))(1,)` (Excel's page's example); an
/// argument left out, as the inlining writes one in, `TRUE`; anything else
/// `FALSE`.
fn isomitted<'e>(ctx: &Ctx<'_>, args: &'e [Expr]) -> Result<Arg<'e>, Reason> {
    let [a] = args else {
        return error(ErrorKind::Value);
    };
    Ok(Arg::Scalar(Value::Bool(match a {
        Expr::Local(l) => ctx.omitted(l),
        Expr::Empty => true,
        _ => false,
    })))
}
