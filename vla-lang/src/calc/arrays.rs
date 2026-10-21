//! The array functions (KERNEL.8, slice f, 2026-10-20): `SEQUENCE`,
//! `FILTER`, `SORT`, `SORTBY`, `UNIQUE`, `TRANSPOSE`, `SUMPRODUCT`,
//! `MMULT`, `ROWS` and `COLUMNS`, and the shapers `CHOOSECOLS`,
//! `CHOOSEROWS`, `TAKE`, `DROP`, `HSTACK`, `VSTACK`, `TOCOL`, `TOROW`,
//! `WRAPROWS`, `WRAPCOLS` and `EXPAND` (the owner's decision 1,
//! 2026-10-20: the outline's list, `SORTBY`, `ROWS` and `COLUMNS`, and the
//! shape functions with them).
//!
//! Each answers as Excel's documentation states it, and where it is silent
//! as the slice's fixture holds it to Excel's saved answer
//! (`scripts/recalc/arrays.txt`). An argument read as an array is a
//! reference's one area read whole, every place of it, an empty cell
//! `Value::Empty`; an array itself; or one value, an array of one. An array
//! with nothing in it is `#CALC!`, Excel's answer for one (`FILTER` keeping
//! no row without its third argument, `SEQUENCE` of no rows, `TAKE` of
//! none). Every array made is sized before it is made, an array past the
//! grid's limit labelled (`super::array::room`). `SUMPRODUCT`'s and
//! `MMULT`'s sums are added in a stated order, place by place in rows and
//! each product's factors in their arguments' order, `MMULT`'s from the
//! first column along (`Alonzo/docs/ALGEBRA.md` section 7), so that every
//! path that computes them adds alike.
//!
//! The functions answering arrays are registered as such
//! (`eval::Arrayed`), their array parameters read as arrays and never
//! lifted, their other parameters read as one value each; `SUMPRODUCT`,
//! `ROWS` and `COLUMNS` answer one value and lift nothing, every parameter
//! an array's.

use std::cmp::Ordering;
use std::collections::HashMap;

use super::array::{self, Array};
use super::eval::{arrays, finite, values, Arg, Ctx, Library, Param, Reason};
use super::order;
use super::value::{ErrorKind, Value};

/// Register the array functions with their parameters' classes.
pub(crate) fn register(lib: &mut Library) {
    lib.register_arrayed("SEQUENCE", sequence, values);
    lib.register_arrayed("FILTER", filter, filtered);
    lib.register_arrayed("SORT", sort, sorted);
    lib.register_arrayed("SORTBY", sortby, sorted_by);
    lib.register_arrayed("UNIQUE", unique, first_array);
    lib.register_arrayed("TRANSPOSE", transpose, arrays);
    lib.register_classed("SUMPRODUCT", sumproduct, arrays);
    lib.register_arrayed("MMULT", mmult, arrays);
    lib.register_classed("ROWS", rows, arrays);
    lib.register_classed("COLUMNS", columns, arrays);
    lib.register_arrayed("CHOOSECOLS", choosecols, picked);
    lib.register_arrayed("CHOOSEROWS", chooserows, picked);
    lib.register_arrayed("TAKE", take, first_array);
    lib.register_arrayed("DROP", drop, first_array);
    lib.register_arrayed("HSTACK", hstack, arrays);
    lib.register_arrayed("VSTACK", vstack, arrays);
    lib.register_arrayed("TOCOL", tocol, first_array);
    lib.register_arrayed("TOROW", torow, first_array);
    lib.register_arrayed("WRAPROWS", wraprows, wrapped);
    lib.register_arrayed("WRAPCOLS", wrapcols, wrapped);
    lib.register_arrayed("EXPAND", expand, expanded);
}

/// The first parameter an array, the rest one value each: `UNIQUE`'s,
/// `TAKE`'s, `DROP`'s, `TOCOL`'s and `TOROW`'s.
fn first_array(j: usize) -> Param {
    if j == 0 {
        Param::Array
    } else {
        Param::Value
    }
}

/// `FILTER`'s: the array, the array of what to keep, and what to answer
/// when nothing is kept, which may be an array, handed whole.
fn filtered(j: usize) -> Param {
    match j {
        0 | 1 => Param::Array,
        _ => Param::Range,
    }
}

/// `SORT`'s: the array, the index and the order, each of which may be an
/// array of them, handed whole, and whether it sorts columns.
fn sorted(j: usize) -> Param {
    match j {
        0 => Param::Array,
        1 | 2 => Param::Range,
        _ => Param::Value,
    }
}

/// `SORTBY`'s: the array, then an array to sort by and its order in turn.
fn sorted_by(j: usize) -> Param {
    if j % 2 == 1 || j == 0 {
        Param::Array
    } else {
        Param::Range
    }
}

/// `CHOOSECOLS`'s and `CHOOSEROWS`'s: the array, then the places, each of
/// which may be an array of them, handed whole.
fn picked(j: usize) -> Param {
    if j == 0 {
        Param::Array
    } else {
        Param::Range
    }
}

/// `WRAPROWS`'s and `WRAPCOLS`'s: the vector, the count, and the value
/// that pads, read as one value.
fn wrapped(j: usize) -> Param {
    match j {
        0 => Param::Array,
        1 => Param::Value,
        _ => Param::Range,
    }
}

/// `EXPAND`'s: the array, the rows and the columns, and the value that pads.
fn expanded(j: usize) -> Param {
    match j {
        0 => Param::Array,
        1 | 2 => Param::Value,
        _ => Param::Range,
    }
}

/// An error answered as the function's value.
fn error(k: ErrorKind) -> Result<Arg<'static>, Reason> {
    Ok(Arg::Scalar(Value::Error(k)))
}

/// An answer that is an error, or the value it holds, for the macros below.
type Got<T> = Result<Result<T, ErrorKind>, Reason>;

/// The value inside an answer, or the function's return of its error.
macro_rules! take {
    ($e:expr) => {
        match $e? {
            Ok(v) => v,
            Err(k) => return error(k),
        }
    };
}

/// An argument read as an array: a reference's one area read whole, an
/// array itself, one value an array of one, an error among them as a value
/// it holds; `#VALUE!` for a span of sheets.
fn array_of(ctx: &Ctx<'_>, arg: &Arg<'_>) -> Got<Array> {
    Ok(match arg {
        Arg::Array(a) => Ok(a.clone()),
        Arg::Area { areas, .. } => match &areas[..] {
            [a] => Ok(ctx.area_array(a)?),
            _ => Err(ErrorKind::Value),
        },
        Arg::Scalar(v) => Ok(Array::single(v.clone())),
    })
}

/// Whether an argument was left out or given as nothing.
fn omitted(arg: Option<&Arg<'_>>) -> bool {
    matches!(arg, None | Some(Arg::Scalar(Value::Empty)))
}

/// A number argument, `None` left out or given as nothing; an empty cell a
/// reference names reads as 0, as a number read from it does.
fn number(ctx: &Ctx<'_>, arg: Option<&Arg<'_>>) -> Got<Option<f64>> {
    match arg {
        _ if omitted(arg) => Ok(Ok(None)),
        Some(a) => Ok(ctx.one(a)?.number().map(Some)),
        None => Ok(Ok(None)),
    }
}

/// A count or a place, its number read down to a whole one.
fn whole(ctx: &Ctx<'_>, arg: Option<&Arg<'_>>) -> Got<Option<f64>> {
    Ok(number(ctx, arg)?.map(|n| n.map(f64::trunc)))
}

/// A truth value argument, `default` left out or given as nothing.
fn truth_or(ctx: &Ctx<'_>, arg: Option<&Arg<'_>>, default: bool) -> Got<bool> {
    match arg {
        _ if omitted(arg) => Ok(Ok(default)),
        Some(a) => Ok(ctx.one(a)?.truth()),
        None => Ok(Ok(default)),
    }
}

/// Every number of an argument that may be an array of them, each read down
/// to a whole one, `default` alone where it is left out.
fn wholes(ctx: &Ctx<'_>, arg: Option<&Arg<'_>>, default: f64) -> Got<Vec<f64>> {
    let Some(a) = arg.filter(|_| !omitted(arg)) else {
        return Ok(Ok(vec![default]));
    };
    let a = match array_of(ctx, a)? {
        Ok(a) => a,
        Err(k) => return Ok(Err(k)),
    };
    let mut out = Vec::with_capacity(a.len());
    for v in a.values() {
        match v.number() {
            Ok(n) => out.push(n.trunc()),
            Err(k) => return Ok(Err(k)),
        }
    }
    Ok(Ok(out))
}

/// An array made of these rows and columns by `at`, sized first; `#CALC!`
/// for none.
fn made(
    rows: u64,
    columns: u64,
    mut at: impl FnMut(u32, u32) -> Value,
) -> Result<Arg<'static>, Reason> {
    let (rows, columns) = match array::room(rows, columns)? {
        Ok(rc) => rc,
        Err(k) => return error(k),
    };
    let mut values = Vec::with_capacity((rows as usize) * (columns as usize));
    for r in 0..rows {
        for c in 0..columns {
            values.push(at(r, c));
        }
    }
    Ok(Arg::Array(Array::new(rows, columns, values)))
}

/// An array's lines, its rows or (`columns`) its columns: how many, and the
/// value at a line's place.
struct Lines<'a> {
    array: &'a Array,
    columns: bool,
}

impl Lines<'_> {
    fn count(&self) -> u32 {
        if self.columns {
            self.array.columns()
        } else {
            self.array.rows()
        }
    }

    fn span(&self) -> u32 {
        if self.columns {
            self.array.rows()
        } else {
            self.array.columns()
        }
    }

    fn at(&self, line: u32, k: u32) -> &Value {
        if self.columns {
            self.array.get(k, line)
        } else {
            self.array.get(line, k)
        }
    }

    /// The lines picked, in order, as an array.
    fn picked(&self, lines: &[u32]) -> Result<Arg<'static>, Reason> {
        let span = self.span();
        if self.columns {
            made(u64::from(span), lines.len() as u64, |r, c| {
                self.at(lines[c as usize], r).clone()
            })
        } else {
            made(lines.len() as u64, u64::from(span), |r, c| {
                self.at(lines[r as usize], c).clone()
            })
        }
    }
}

/// `SEQUENCE(rows, [columns], [start], [step])`: the numbers from `start`
/// by `step`, row by row, each `start` and its place times `step`; the
/// counts read down to whole numbers, one where left out, below 0
/// `#VALUE!`, 0 `#CALC!`.
fn sequence(ctx: &Ctx<'_>, args: &[Arg<'_>]) -> Result<Arg<'static>, Reason> {
    if !(1..=4).contains(&args.len()) {
        return error(ErrorKind::Value);
    }
    let rows = take!(whole(ctx, args.first())).unwrap_or(1.0);
    let columns = take!(whole(ctx, args.get(1))).unwrap_or(1.0);
    let start = take!(number(ctx, args.get(2))).unwrap_or(1.0);
    let step = take!(number(ctx, args.get(3))).unwrap_or(1.0);
    if rows < 0.0 || columns < 0.0 {
        return error(ErrorKind::Value);
    }
    let width = columns as u64;
    made(rows as u64, width, |r, c| {
        let k = u64::from(r) * width + u64::from(c);
        finite(start + step * k as f64)
    })
}

/// `FILTER(array, include, [if_empty])`: the rows of the array whose place
/// in `include`, a column as tall as it, is true, or the columns where
/// `include` is a row as wide as it; a number is true but 0, a truth value
/// itself, an empty value false, a text `#VALUE!`, an error the answer.
/// Nothing kept: `if_empty` as it stands, or `#CALC!` without it.
fn filter(ctx: &Ctx<'_>, args: &[Arg<'_>]) -> Result<Arg<'static>, Reason> {
    if !(2..=3).contains(&args.len()) {
        return error(ErrorKind::Value);
    }
    let a = take!(array_of(ctx, &args[0]));
    let include = take!(array_of(ctx, &args[1]));
    let columns = if include.columns() == 1 && include.rows() == a.rows() {
        false
    } else if include.rows() == 1 && include.columns() == a.columns() {
        true
    } else {
        return error(ErrorKind::Value);
    };
    let mut kept = Vec::new();
    for (k, v) in include.values().iter().enumerate() {
        match v {
            Value::Error(e) => return error(*e),
            Value::Bool(true) => kept.push(k as u32),
            Value::Number(n) if *n != 0.0 => kept.push(k as u32),
            Value::Text(_) => return error(ErrorKind::Value),
            _ => {}
        }
    }
    if kept.is_empty() {
        return match args.get(2) {
            Some(otherwise) => Ok(otherwise.clone().into_owned()),
            None => error(ErrorKind::Calc),
        };
    }
    Lines { array: &a, columns }.picked(&kept)
}

/// The order of a value among those `SORT` and `SORTBY` sort: numbers by
/// value, then texts in Excel's order of texts, then FALSE and TRUE, then
/// the errors, each one; a descending sort the reverse; an empty value last
/// either way, as Excel's sorts place a blank. Two texts a character beyond
/// Excel's known order decides are noted in `unknown`, the sort labelled.
fn sort_order(a: &Value, b: &Value, descending: bool, unknown: &mut Option<Reason>) -> Ordering {
    let rank = |v: &Value| match v {
        Value::Number(_) => 0,
        Value::Text(_) => 1,
        Value::Bool(_) => 2,
        Value::Error(_) => 3,
        Value::Empty => 4,
    };
    let o = match (a, b) {
        (Value::Empty, Value::Empty) => return Ordering::Equal,
        (Value::Empty, _) => return Ordering::Greater,
        (_, Value::Empty) => return Ordering::Less,
        (Value::Number(x), Value::Number(y)) => x.partial_cmp(y).unwrap_or(Ordering::Equal),
        (Value::Text(x), Value::Text(y)) => {
            let c = order::compare(x, y);
            if !c.known && c.order != Ordering::Equal {
                unknown.get_or_insert_with(order::unknown);
            }
            c.order
        }
        (Value::Bool(x), Value::Bool(y)) => x.cmp(y),
        _ => rank(a).cmp(&rank(b)),
    };
    if descending {
        o.reverse()
    } else {
        o
    }
}

/// The lines of an array sorted, stably, by keys in turn: `value` a line's
/// value at a key, `descending` whether each key sorts descending.
fn sorted_lines<'a>(
    count: u32,
    descending: &[bool],
    value: impl Fn(u32, usize) -> &'a Value,
) -> Result<Vec<u32>, Reason> {
    let mut lines: Vec<u32> = (0..count).collect();
    let mut unknown = None;
    lines.sort_by(|&x, &y| {
        for (k, &d) in descending.iter().enumerate() {
            let o = sort_order(value(x, k), value(y, k), d, &mut unknown);
            if o != Ordering::Equal {
                return o;
            }
        }
        Ordering::Equal
    });
    match unknown {
        Some(r) => Err(r),
        None => Ok(lines),
    }
}

/// A sort order's number as a direction: 1 ascending, -1 descending, any
/// other `#VALUE!`.
fn direction(n: f64) -> Result<bool, ErrorKind> {
    if n == 1.0 {
        Ok(false)
    } else if n == -1.0 {
        Ok(true)
    } else {
        Err(ErrorKind::Value)
    }
}

/// `SORT(array, [sort_index], [sort_order], [by_col])`: the array's rows,
/// or columns `by_col`, in order by the values at `sort_index`, the first
/// where left out, and at each further index of an array of them;
/// `sort_order` 1 ascending, the default, or -1 descending, one for every
/// index or one each; an index outside the array or an order other than 1
/// or -1 `#VALUE!`. Rows of one order keep theirs.
fn sort(ctx: &Ctx<'_>, args: &[Arg<'_>]) -> Result<Arg<'static>, Reason> {
    if !(1..=4).contains(&args.len()) {
        return error(ErrorKind::Value);
    }
    let a = take!(array_of(ctx, &args[0]));
    let columns = take!(truth_or(ctx, args.get(3), false));
    let lines = Lines { array: &a, columns };
    let indices = take!(wholes(ctx, args.get(1), 1.0));
    let orders = take!(wholes(ctx, args.get(2), 1.0));
    if orders.len() != 1 && orders.len() != indices.len() {
        return error(ErrorKind::Value);
    }
    let span = f64::from(lines.span());
    let mut keys: Vec<(u32, bool)> = Vec::with_capacity(indices.len());
    for (k, &i) in indices.iter().enumerate() {
        if i < 1.0 || i > span {
            return error(ErrorKind::Value);
        }
        let descending = match direction(orders[if orders.len() == 1 { 0 } else { k }]) {
            Ok(d) => d,
            Err(e) => return error(e),
        };
        keys.push((i as u32 - 1, descending));
    }
    let descending: Vec<bool> = keys.iter().map(|&(_, d)| d).collect();
    let order = sorted_lines(lines.count(), &descending, |line, k| {
        lines.at(line, keys[k].0)
    })?;
    lines.picked(&order)
}

/// `SORTBY(array, by_array, [sort_order], ...)`: the array's rows in order
/// by each `by_array` in turn, a column as tall as the array, or its
/// columns by rows as wide as it, every one of one kind; each order 1, the
/// default, or -1; any other shape or order `#VALUE!`.
fn sortby(ctx: &Ctx<'_>, args: &[Arg<'_>]) -> Result<Arg<'static>, Reason> {
    if args.len() < 2 {
        return error(ErrorKind::Value);
    }
    let a = take!(array_of(ctx, &args[0]));
    let mut by: Vec<(Array, bool)> = Vec::new();
    let mut columns: Option<bool> = None;
    let mut j = 1;
    while j < args.len() {
        let b = take!(array_of(ctx, &args[j]));
        let descending = match take!(whole(ctx, args.get(j + 1))) {
            None => false,
            Some(n) => match direction(n) {
                Ok(d) => d,
                Err(e) => return error(e),
            },
        };
        let down = b.columns() == 1 && b.rows() == a.rows();
        let across = b.rows() == 1 && b.columns() == a.columns();
        let kind = match (down, across, columns) {
            (true, _, None | Some(false)) => false,
            (_, true, None | Some(true)) => true,
            _ => return error(ErrorKind::Value),
        };
        columns = Some(kind);
        by.push((b, descending));
        j += 2;
    }
    let columns = columns.unwrap_or(false);
    let lines = Lines { array: &a, columns };
    let descending: Vec<bool> = by.iter().map(|&(_, d)| d).collect();
    let order = sorted_lines(lines.count(), &descending, |line, k| {
        &by[k].0.values()[line as usize]
    })?;
    lines.picked(&order)
}

/// A value's part of a line's key, by which `UNIQUE` tells lines one: a
/// number by its value exactly, 0 and -0 one; a text as Excel's order of
/// texts holds two one, case aside; a truth value; an error by its kind; an
/// empty value apart from every other.
fn key_part(v: &Value, out: &mut String) {
    match v {
        Value::Number(n) => {
            let n = if *n == 0.0 { 0.0 } else { *n };
            out.push('n');
            out.push_str(&n.to_bits().to_string());
        }
        Value::Text(t) => {
            out.push('t');
            out.push_str(&order::key(t));
        }
        Value::Bool(b) => out.push_str(if *b { "b1" } else { "b0" }),
        Value::Error(k) => {
            out.push('e');
            out.push_str(k.text());
        }
        Value::Empty => out.push('_'),
    }
    out.push('\u{0}');
}

/// `UNIQUE(array, [by_col], [exactly_once])`: the array's distinct rows, or
/// columns `by_col`, each at its first place; `exactly_once`, those that
/// stand once alone; none `#CALC!`.
fn unique(ctx: &Ctx<'_>, args: &[Arg<'_>]) -> Result<Arg<'static>, Reason> {
    if !(1..=3).contains(&args.len()) {
        return error(ErrorKind::Value);
    }
    let a = take!(array_of(ctx, &args[0]));
    let columns = take!(truth_or(ctx, args.get(1), false));
    let once = take!(truth_or(ctx, args.get(2), false));
    let lines = Lines { array: &a, columns };
    let mut seen: HashMap<String, usize> = HashMap::new();
    let mut firsts: Vec<(u32, u32)> = Vec::new();
    let mut key = String::new();
    for line in 0..lines.count() {
        key.clear();
        for k in 0..lines.span() {
            key_part(lines.at(line, k), &mut key);
        }
        match seen.get(&key) {
            Some(&i) => firsts[i].1 += 1,
            None => {
                seen.insert(key.clone(), firsts.len());
                firsts.push((line, 1));
            }
        }
    }
    let kept: Vec<u32> = firsts
        .iter()
        .filter(|&&(_, n)| !once || n == 1)
        .map(|&(line, _)| line)
        .collect();
    if kept.is_empty() {
        return error(ErrorKind::Calc);
    }
    lines.picked(&kept)
}

/// `TRANSPOSE(array)`: its rows as columns.
fn transpose(ctx: &Ctx<'_>, args: &[Arg<'_>]) -> Result<Arg<'static>, Reason> {
    let [a] = args else {
        return error(ErrorKind::Value);
    };
    let a = take!(array_of(ctx, a));
    made(u64::from(a.columns()), u64::from(a.rows()), |r, c| {
        a.get(c, r).clone()
    })
}

/// `SUMPRODUCT(array, ...)`: the sum of the products of the arrays' values
/// place by place, every array of one shape (any other `#VALUE!`); a value
/// that is no number counts as 0, an error is the answer; added place by
/// place in rows from 0, each product's factors in their arguments' order.
fn sumproduct(ctx: &Ctx<'_>, args: &[Arg<'_>]) -> Result<Value, Reason> {
    if args.is_empty() || args.len() > 255 {
        return Ok(Value::Error(ErrorKind::Value));
    }
    let mut held = Vec::with_capacity(args.len());
    for a in args {
        match array_of(ctx, a)? {
            Ok(a) => held.push(a),
            Err(k) => return Ok(Value::Error(k)),
        }
    }
    let (rows, columns) = (held[0].rows(), held[0].columns());
    if held
        .iter()
        .any(|a| a.rows() != rows || a.columns() != columns)
    {
        return Ok(Value::Error(ErrorKind::Value));
    }
    let mut sum = 0.0;
    for k in 0..held[0].len() {
        let mut product = 1.0;
        for a in &held {
            match &a.values()[k] {
                Value::Number(n) => product *= n,
                Value::Error(e) => return Ok(Value::Error(*e)),
                _ => product = 0.0,
            }
        }
        sum += product;
    }
    Ok(finite(sum))
}

/// `MMULT(array1, array2)`: the matrix product, as many rows as the first
/// and columns as the second, each value the sum along the first's row and
/// the second's column, added from the first column along; the first's
/// columns not the second's rows, or a value no number, `#VALUE!`, an error
/// the answer.
fn mmult(ctx: &Ctx<'_>, args: &[Arg<'_>]) -> Result<Arg<'static>, Reason> {
    let [x, y] = args else {
        return error(ErrorKind::Value);
    };
    let x = take!(array_of(ctx, x));
    let y = take!(array_of(ctx, y));
    if x.columns() != y.rows() {
        return error(ErrorKind::Value);
    }
    for v in x.values().iter().chain(y.values()) {
        match v {
            Value::Number(_) => {}
            Value::Error(k) => return error(*k),
            _ => return error(ErrorKind::Value),
        }
    }
    let n = |v: &Value| match v {
        Value::Number(n) => *n,
        _ => 0.0,
    };
    made(u64::from(x.rows()), u64::from(y.columns()), |r, c| {
        let mut sum = 0.0;
        for k in 0..x.columns() {
            sum += n(x.get(r, k)) * n(y.get(k, c));
        }
        finite(sum)
    })
}

/// `ROWS(array)` and `COLUMNS` (`across`): how many rows (columns) a
/// reference's one area or an array spans, one for one value, an error the
/// answer; a span of sheets `#VALUE!`.
fn measure(args: &[Arg<'_>], across: bool) -> Result<Value, Reason> {
    let [a] = args else {
        return Ok(Value::Error(ErrorKind::Value));
    };
    let n = match a {
        Arg::Area { areas, .. } => match &areas[..] {
            [r] if across => r.right - r.left + 1,
            [r] => r.bottom - r.top + 1,
            _ => return Ok(Value::Error(ErrorKind::Value)),
        },
        Arg::Array(x) if across => x.columns(),
        Arg::Array(x) => x.rows(),
        Arg::Scalar(Value::Error(k)) => return Ok(Value::Error(*k)),
        Arg::Scalar(_) => 1,
    };
    Ok(Value::Number(f64::from(n)))
}

/// `ROWS(array)`.
fn rows(_: &Ctx<'_>, args: &[Arg<'_>]) -> Result<Value, Reason> {
    measure(args, false)
}

/// `COLUMNS(array)`.
fn columns(_: &Ctx<'_>, args: &[Arg<'_>]) -> Result<Value, Reason> {
    measure(args, true)
}

/// `CHOOSECOLS(array, col, ...)` and `CHOOSEROWS` (`columns` false): the
/// lines at the places given, in their order, a place an array of them as
/// well; a place below 0 counted from the last line; 0 or past the lines
/// `#VALUE!`.
fn choose_lines(ctx: &Ctx<'_>, args: &[Arg<'_>], columns: bool) -> Result<Arg<'static>, Reason> {
    if args.len() < 2 {
        return error(ErrorKind::Value);
    }
    let a = take!(array_of(ctx, &args[0]));
    let lines = Lines { array: &a, columns };
    let count = f64::from(lines.count());
    let mut picks = Vec::new();
    for arg in &args[1..] {
        for n in take!(wholes(ctx, Some(arg), 0.0)) {
            if n >= 1.0 && n <= count {
                picks.push(n as u32 - 1);
            } else if n <= -1.0 && -n <= count {
                picks.push((count + n) as u32);
            } else {
                return error(ErrorKind::Value);
            }
        }
    }
    lines.picked(&picks)
}

/// `CHOOSECOLS(array, col_num, ...)`.
fn choosecols(ctx: &Ctx<'_>, args: &[Arg<'_>]) -> Result<Arg<'static>, Reason> {
    choose_lines(ctx, args, true)
}

/// `CHOOSEROWS(array, row_num, ...)`.
fn chooserows(ctx: &Ctx<'_>, args: &[Arg<'_>]) -> Result<Arg<'static>, Reason> {
    choose_lines(ctx, args, false)
}

/// The places along one side that `TAKE` keeps (`drop` false) or `DROP`
/// leaves: a count from the first, or from the last below 0, as many as
/// there are past them, every place where left out; `None` where nothing is
/// left, which is `#CALC!`.
fn kept(n: Option<f64>, size: u32, drop: bool) -> Option<(u32, u32)> {
    let size64 = f64::from(size);
    let Some(n) = n else {
        return Some((0, size));
    };
    let (from, to) = match (drop, n >= 0.0) {
        (false, true) => (0.0, n.min(size64)),
        (false, false) => ((size64 + n).max(0.0), size64),
        (true, true) => (n.min(size64), size64),
        (true, false) => (0.0, (size64 + n).max(0.0)),
    };
    (to > from).then_some((from as u32, to as u32))
}

/// `TAKE(array, rows, [columns])` and `DROP` (`drop`): the rows and the
/// columns kept, each count from the first or below 0 from the last, every
/// one where left out; nothing left `#CALC!`.
fn take_or_drop(ctx: &Ctx<'_>, args: &[Arg<'_>], drop: bool) -> Result<Arg<'static>, Reason> {
    if !(2..=3).contains(&args.len()) {
        return error(ErrorKind::Value);
    }
    let a = take!(array_of(ctx, &args[0]));
    let rows = take!(whole(ctx, args.get(1)));
    let columns = take!(whole(ctx, args.get(2)));
    let (Some((top, bottom)), Some((left, right))) =
        (kept(rows, a.rows(), drop), kept(columns, a.columns(), drop))
    else {
        return error(ErrorKind::Calc);
    };
    made(u64::from(bottom - top), u64::from(right - left), |r, c| {
        a.get(top + r, left + c).clone()
    })
}

/// `TAKE(array, rows, [columns])`.
fn take(ctx: &Ctx<'_>, args: &[Arg<'_>]) -> Result<Arg<'static>, Reason> {
    take_or_drop(ctx, args, false)
}

/// `DROP(array, rows, [columns])`.
fn drop(ctx: &Ctx<'_>, args: &[Arg<'_>]) -> Result<Arg<'static>, Reason> {
    take_or_drop(ctx, args, true)
}

/// `HSTACK(array, ...)` and `VSTACK` (`down`): the arrays side by side, or
/// one under another, as many rows (columns) as the tallest (widest), a
/// place an array does not reach `#N/A`.
fn stack(ctx: &Ctx<'_>, args: &[Arg<'_>], down: bool) -> Result<Arg<'static>, Reason> {
    if args.is_empty() {
        return error(ErrorKind::Value);
    }
    let mut held = Vec::with_capacity(args.len());
    for a in args {
        held.push(take!(array_of(ctx, a)));
    }
    // Each array's first place along the stacking side.
    let mut starts = Vec::with_capacity(held.len());
    let mut along = 0u64;
    let mut across = 0u32;
    for a in &held {
        starts.push(along);
        if down {
            along += u64::from(a.rows());
            across = across.max(a.columns());
        } else {
            along += u64::from(a.columns());
            across = across.max(a.rows());
        }
    }
    let at = |k: u64, other: u32| -> Value {
        let i = starts.partition_point(|&s| s <= k) - 1;
        let a = &held[i];
        let k = (k - starts[i]) as u32;
        let (r, c) = if down { (k, other) } else { (other, k) };
        if r < a.rows() && c < a.columns() {
            a.get(r, c).clone()
        } else {
            Value::Error(ErrorKind::NA)
        }
    };
    if down {
        made(along, u64::from(across), |r, c| at(u64::from(r), c))
    } else {
        made(u64::from(across), along, |r, c| at(u64::from(c), r))
    }
}

/// `HSTACK(array, ...)`.
fn hstack(ctx: &Ctx<'_>, args: &[Arg<'_>]) -> Result<Arg<'static>, Reason> {
    stack(ctx, args, false)
}

/// `VSTACK(array, ...)`.
fn vstack(ctx: &Ctx<'_>, args: &[Arg<'_>]) -> Result<Arg<'static>, Reason> {
    stack(ctx, args, true)
}

/// `TOCOL(array, [ignore], [scan_by_column])` and `TOROW` (`row`): the
/// array's values in one column (row), read row by row or, by its third
/// argument, column by column; `ignore` 1 passing empty values by, 2 the
/// errors, 3 both, 0 the default none, any other `#VALUE!`; none left
/// `#CALC!`.
fn to_line(ctx: &Ctx<'_>, args: &[Arg<'_>], row: bool) -> Result<Arg<'static>, Reason> {
    if !(1..=3).contains(&args.len()) {
        return error(ErrorKind::Value);
    }
    let a = take!(array_of(ctx, &args[0]));
    let ignore = take!(whole(ctx, args.get(1))).unwrap_or(0.0);
    if !(0.0..=3.0).contains(&ignore) {
        return error(ErrorKind::Value);
    }
    let ignore = ignore as u8;
    let by_column = take!(truth_or(ctx, args.get(2), false));
    let mut out = Vec::with_capacity(a.len());
    let (outer, inner) = if by_column {
        (a.columns(), a.rows())
    } else {
        (a.rows(), a.columns())
    };
    for i in 0..outer {
        for j in 0..inner {
            let v = if by_column { a.get(j, i) } else { a.get(i, j) };
            let passed = match v {
                Value::Empty => ignore & 1 == 1,
                Value::Error(_) => ignore & 2 == 2,
                _ => false,
            };
            if !passed {
                out.push(v.clone());
            }
        }
    }
    let n = out.len() as u64;
    if row {
        made(1, n, |_, c| out[c as usize].clone())
    } else {
        made(n, 1, |r, _| out[r as usize].clone())
    }
}

/// `TOCOL(array, [ignore], [scan_by_column])`.
fn tocol(ctx: &Ctx<'_>, args: &[Arg<'_>]) -> Result<Arg<'static>, Reason> {
    to_line(ctx, args, false)
}

/// `TOROW(array, [ignore], [scan_by_column])`.
fn torow(ctx: &Ctx<'_>, args: &[Arg<'_>]) -> Result<Arg<'static>, Reason> {
    to_line(ctx, args, true)
}

/// The value that pads, `#N/A` where left out.
fn pad(ctx: &Ctx<'_>, arg: Option<&Arg<'_>>) -> Result<Value, Reason> {
    match arg {
        _ if omitted(arg) => Ok(Value::Error(ErrorKind::NA)),
        Some(a) => ctx.one(a),
        None => Ok(Value::Error(ErrorKind::NA)),
    }
}

/// `WRAPROWS(vector, wrap_count, [pad_with])` and `WRAPCOLS` (`columns`):
/// a row or a column of values laid in rows (columns) of `wrap_count`,
/// the last padded; an array of more than one row and column `#VALUE!`, a
/// count below 1 `#NUM!`.
fn wrap(ctx: &Ctx<'_>, args: &[Arg<'_>], columns: bool) -> Result<Arg<'static>, Reason> {
    if !(2..=3).contains(&args.len()) {
        return error(ErrorKind::Value);
    }
    let v = take!(array_of(ctx, &args[0]));
    if !v.is_vector() {
        return error(ErrorKind::Value);
    }
    let Some(count) = take!(whole(ctx, args.get(1))) else {
        return error(ErrorKind::Value);
    };
    if count < 1.0 {
        return error(ErrorKind::Num);
    }
    let filler = pad(ctx, args.get(2))?;
    let n = v.len() as u64;
    let count = (count as u64).min(n.max(1));
    let lines = n.div_ceil(count);
    let at = |k: u64| -> Value {
        match v.values().get(k as usize) {
            Some(x) => x.clone(),
            None => filler.clone(),
        }
    };
    if columns {
        made(count, lines, |r, c| at(u64::from(c) * count + u64::from(r)))
    } else {
        made(lines, count, |r, c| at(u64::from(r) * count + u64::from(c)))
    }
}

/// `WRAPROWS(vector, wrap_count, [pad_with])`.
fn wraprows(ctx: &Ctx<'_>, args: &[Arg<'_>]) -> Result<Arg<'static>, Reason> {
    wrap(ctx, args, false)
}

/// `WRAPCOLS(vector, wrap_count, [pad_with])`.
fn wrapcols(ctx: &Ctx<'_>, args: &[Arg<'_>]) -> Result<Arg<'static>, Reason> {
    wrap(ctx, args, true)
}

/// `EXPAND(array, rows, [columns], [pad_with])`: the array grown to the
/// rows and the columns given, its own where left out, padded; fewer than
/// it holds `#VALUE!`.
fn expand(ctx: &Ctx<'_>, args: &[Arg<'_>]) -> Result<Arg<'static>, Reason> {
    if !(2..=4).contains(&args.len()) {
        return error(ErrorKind::Value);
    }
    let a = take!(array_of(ctx, &args[0]));
    let rows = take!(whole(ctx, args.get(1))).unwrap_or(f64::from(a.rows()));
    let columns = take!(whole(ctx, args.get(2))).unwrap_or(f64::from(a.columns()));
    if rows < f64::from(a.rows()) || columns < f64::from(a.columns()) {
        return error(ErrorKind::Value);
    }
    let filler = pad(ctx, args.get(3))?;
    made(rows as u64, columns as u64, |r, c| {
        if r < a.rows() && c < a.columns() {
            a.get(r, c).clone()
        } else {
            filler.clone()
        }
    })
}

#[cfg(test)]
mod tests {
    use crate::calc::{library, Calc};
    use crate::sheet::{parse_a1_range, Cell, Content, Workbook};

    /// A sheet `S` of three numbers in A1:A3, three in B1:B3 and texts in
    /// C1:C2 above an empty C3, and the cells given: a formula's text after
    /// `=`, `~=` for a dynamic formula, `{=` for a legacy array formula.
    fn book(cells: &[(&str, &str)]) -> Workbook {
        let mut wb = Workbook::new();
        let s = wb.ensure_sheet("S");
        let data = [
            ("A1", "1"),
            ("A2", "2"),
            ("A3", "3"),
            ("B1", "4"),
            ("B2", "5"),
            ("B3", "6"),
            ("C1", "a"),
            ("C2", "bc"),
        ];
        for (addr, text) in data.iter().chain(cells) {
            let r = parse_a1_range(addr).unwrap();
            let content = if let Some(f) = text.strip_prefix("~=") {
                Content::DynamicFormula(f.to_string())
            } else if let Some(f) = text.strip_prefix("{=") {
                Content::ArrayFormula(f.to_string())
            } else if let Some(f) = text.strip_prefix('=') {
                Content::Formula(f.to_string())
            } else if let Ok(n) = text.parse::<f64>() {
                Content::Number(n)
            } else {
                Content::Text(text.to_string())
            };
            wb.sheets[s].set(r.top, r.left, Cell { content, style: 0 });
        }
        wb
    }

    /// Each formula's answer as a row spells it, held to `want`.
    fn holds(cases: &[(&str, &str, &str)]) {
        holds_by(&library::language(), cases);
    }

    /// The same, computed by the library given.
    fn holds_by(lib: &crate::calc::Library, cases: &[(&str, &str, &str)]) {
        let cells: Vec<(&str, &str)> = cases.iter().map(|&(a, f, _)| (a, f)).collect();
        let wb = book(&cells);
        let calc = Calc::run(&wb, lib);
        for &(addr, formula, want) in cases {
            let r = parse_a1_range(addr).unwrap();
            let got = calc.computed((0, r.top, r.left)).map(|c| c.spell());
            assert_eq!(got.as_deref(), Some(want), "{addr}: {formula}");
        }
    }

    #[test]
    fn an_array_constant_reads_as_its_values_and_its_first_in_a_legacy_cell() {
        holds(&[
            ("E10", "={1,2;3,4}", "1"),
            ("E11", "=SUM({1,2;3,4})", "10"),
            ("E12", "=SUM({1,\"2\",TRUE})", "1"),
            ("E13", "=ROWS({1,2;3,4})", "2"),
            ("E14", "=COLUMNS({1,2,3})", "3"),
            ("E15", "=COUNTA({1,\"a\",TRUE,#N/A})", "4"),
            ("E16", "={-1,2}+0", "-1"),
            ("E17", "=AND({TRUE,FALSE})", "false"),
            ("E18", "~={1,2}", "(not-computed \"a spill\")"),
            ("E19", "~={7}", "7"),
        ]);
    }

    #[test]
    fn an_operator_lifts_over_arrays_by_excel_s_rule() {
        // The most rows and columns, a row or a column standing at every
        // place, `#N/A` where one falls short.
        holds(&[
            ("E10", "=SUM({1,2}*{3,4})", "11"),
            ("E11", "=SUM({1,2,3}+{10;20})", "102"),
            ("E12", "=SUM({1,2,3}+{10,20})", "(error \"#N/A\")"),
            ("E13", "=SUM(-{1,2})", "-3"),
            ("E14", "=SUM({50,150}%)", "2"),
            ("E15", "=SUM(({1,2}>1)*1)", "1"),
            ("E16", "=SUM(LEN({\"a\",\"bc\"}&\"x\"))", "5"),
        ]);
    }

    #[test]
    fn a_value_parameter_lifts_its_function_and_an_array_parameter_reads_arrays() {
        holds(&[
            ("E10", "=SUM(LEN({\"a\",\"bc\"}))", "3"),
            ("E11", "=SUM(ABS({-1,-2}))", "3"),
            ("E12", "=SUM(IF({TRUE,FALSE,TRUE},{1,2,3},0))", "4"),
            ("E13", "=SUM(CHOOSE({1,2},10,20))", "30"),
            // A value forced in a legacy formula, Excel's class for `IFERROR`'s
            // and `VLOOKUP`'s index, is an array's first value; a dynamic
            // formula lifts (Excel's own save, `arrays.txt`'s `A28` and `A33`).
            ("E14", "=SUM(IFERROR({1,#N/A},5))", "1"),
            ("E40", "~=SUM(IFERROR({1,#N/A},5))", "6"),
            ("E15", "=SUM(COUNTIF(A1:A3,{1,3}))", "2"),
            ("E16", "=SUM(INDEX({10,20,30},{1,3}))", "40"),
            ("E17", "=SUM(VLOOKUP(1,{1,2,3},{2,3},FALSE))", "2"),
            ("E41", "~=SUM(VLOOKUP(1,{1,2,3},{2,3},FALSE))", "5"),
            ("E18", "=SUM(MATCH({3,1},A1:A3,0))", "4"),
            // A range where one value is wanted: met at the cell in a legacy
            // formula (row 20 holds none of A1:A3), read whole in a dynamic
            // one and inside an array parameter of either.
            ("E20", "=SUM(A1:A3*B1:B3)", "(error \"#VALUE!\")"),
            ("E21", "~=SUM(A1:A3*B1:B3)", "32"),
            ("E22", "=SUMPRODUCT(A1:A3*B1:B3)", "32"),
            ("E23", "=SUMPRODUCT((A1:A3>1)*B1:B3)", "11"),
            ("E24", "=SUMPRODUCT(ROW(A1:A3)*A1:A3)", "14"),
            ("E25", "=SUM(ROW(A1:A3))", "1"),
            ("E26", "~=SUM(ROW(A1:A3))", "6"),
            ("E27", "~=SUM(LEN(C1:C3))", "3"),
            ("E28", "=MATCH(4,A1:A3*2,0)", "2"),
            ("E29", "=INDEX(A1:A3*10,2)", "20"),
            // `IF` lifted over its condition and both branches, `#N/A` where
            // the condition or the branch it takes falls short, a branch not
            // taken free to, as Excel's own save computes `arrays.txt`'s `A26`.
            (
                "E30",
                "=SUM(IF({TRUE,FALSE,TRUE},{1,2},0))",
                "(error \"#N/A\")",
            ),
            ("E31", "=SUM(IF({FALSE,FALSE,FALSE},{1,2},5))", "15"),
            ("E42", "=SUM(CHOOSE({2,2,2},{10,20},5))", "15"),
            ("E32", "=SUM(IF({TRUE;FALSE},{1,2},{10;20}))", "43"),
            // A function of one value lifted over two arrays, `#N/A` where
            // the shorter does not reach, as an operator's lift.
            ("E33", "=SUM(MOD({5,6,7},{2,3}))", "(error \"#N/A\")"),
            ("E34", "=SUM(IFNA(MOD({5,6,7},{2,3}),100))", "1"),
            ("E43", "~=SUM(IFNA(MOD({5,6,7},{2,3}),100))", "101"),
            // Inside an array parameter of a legacy formula, `IF`'s condition
            // and `IFERROR`'s value keep their own class: a range there is met
            // at the cell (row 44 and 46 hold none of A1:A3), where a dynamic
            // formula reads it whole.
            ("E44", "=SUMPRODUCT(IF(A1:A3>1,1,0))", "(error \"#VALUE!\")"),
            ("E45", "~=SUMPRODUCT(IF(A1:A3>1,1,0))", "2"),
            ("E46", "=SUMPRODUCT(IFERROR(1/(A1:A3-3),7))", "7"),
            ("E47", "~=SUMPRODUCT(IFERROR(1/(A1:A3-3),7))", "5.5"),
            // `COUNTA` counts an array's empty value, where a range's empty
            // cell it passes by (`arrays.txt`'s `A119`).
            ("E48", "=COUNTA(TRANSPOSE(C1:C3))", "3"),
            ("E49", "=COUNTA(C1:C3)", "2"),
            // `IFERROR` lifted in a dynamic formula reads `other` at the place it
            // catches, and leaves it free to fall short where it does not.
            ("E50", "~=SUM(IFERROR({1,#N/A,3}*1,{10,20,30}))", "24"),
            ("E51", "~=SUM(IFERROR({1,2,3}*1,{10,20}))", "6"),
            // A value forced is one value inside an array parameter too, where a
            // value read as arrays would be an array where one is wanted.
            ("E52", "=SUMPRODUCT(VLOOKUP(1,{1,2,3},{2,3},FALSE))", "2"),
            // Every value `VLOOKUP` and `HLOOKUP` take but the table is forced
            // in a legacy formula, the value sought among them, as Excel's own
            // save computes `arrays.txt`'s `A203`, and the match's kind; a
            // dynamic formula and a legacy array formula lift, and a range
            // sought inside an array parameter is met at the cell (row 59
            // holds none of A1:A3).
            ("E53", "=SUM(VLOOKUP({2,3},A1:B3,2,FALSE))", "5"),
            ("E54", "~=SUM(VLOOKUP({2,3},A1:B3,2,FALSE))", "11"),
            ("E61", "{=SUM(VLOOKUP({2,3},A1:B3,2,FALSE))", "11"),
            ("E55", "=SUM(HLOOKUP({1,2},{1,2;10,20},2,FALSE))", "10"),
            ("E56", "~=SUM(HLOOKUP({1,2},{1,2;10,20},2,FALSE))", "30"),
            ("E57", "=SUM(VLOOKUP(2.5,A1:B3,2,{TRUE,FALSE}))", "5"),
            (
                "E58",
                "~=SUM(VLOOKUP(2.5,A1:B3,2,{TRUE,FALSE}))",
                "(error \"#N/A\")",
            ),
            (
                "E59",
                "=SUMPRODUCT(VLOOKUP(A1:A3,A1:B3,2,FALSE))",
                "(error \"#VALUE!\")",
            ),
            ("E60", "~=SUMPRODUCT(VLOOKUP(A1:A3,A1:B3,2,FALSE))", "15"),
        ]);
    }

    #[test]
    fn a_range_s_reach_orders_the_cells_its_span_holds() {
        // The plan's edges for `:` between expressions are its reach, the
        // rectangle spanning every reference either end may answer, `IF`'s
        // other branch among them: `C5` lies in the range `IF` answers,
        // `A4:G6`, and in no reference the formula's text holds, so only
        // the reach orders it before the cell that sums it.
        let wb = book(&[
            ("H1", "=SUM(IF(FALSE,G2,A6):G4)"),
            ("C5", "=A1*100"),
            ("G4", "7"),
        ]);
        let lib = library::language();
        let calc = Calc::run(&wb, &lib);
        let at = |addr: &str| {
            let r = parse_a1_range(addr).unwrap();
            calc.computed((0, r.top, r.left)).map(|c| c.spell())
        };
        assert_eq!(at("C5").as_deref(), Some("100"));
        assert_eq!(at("H1").as_deref(), Some("107"), "A4:G6, C5 and G4 in it");
    }

    #[test]
    fn a_sort_keeps_the_order_of_lines_it_ranks_alike() {
        // Forty lines, which a sort of fewer than twenty would keep in
        // their order however it sorts: the even numbers first, as their
        // key 0 sorts before 1, each half in the order the lines stood.
        let evens: Vec<String> = (2..=40).step_by(2).map(|n| n.to_string()).collect();
        let odds: Vec<String> = (1..=39).step_by(2).map(|n| n.to_string()).collect();
        let want = format!("\"{},{}\"", evens.join(","), odds.join(","));
        holds(&[(
            "E10",
            "=TEXTJOIN(\",\",TRUE,SORTBY(SEQUENCE(40),MOD(SEQUENCE(40),2)))",
            &want,
        )]);
    }

    #[test]
    fn an_array_where_one_value_is_wanted_is_its_first_or_labelled_by_the_formula_s_kind() {
        // A form of the test's own reading its argument as one value, as
        // a form of a crate above may (Excel's `IFS` in `frazaro-core`): an
        // array of many there is its first value in a legacy formula, as
        // Excel's `@` takes it, and labelled where the formula reads
        // arrays, which would lift the form instead.
        fn one<'e>(
            ctx: &crate::calc::eval::Ctx<'_>,
            args: &'e [crate::calc::formula::Expr],
        ) -> Result<crate::calc::eval::Arg<'e>, crate::calc::eval::Reason> {
            Ok(crate::calc::eval::Arg::Scalar(ctx.scalar(&args[0])?))
        }
        let mut lib = library::language();
        lib.register_form("ONE", one);
        let wanted = "(not-computed \"an array where one value is wanted\")";
        holds_by(
            &lib,
            &[
                ("E10", "=ONE({5,6})", "5"),
                ("E11", "~=ONE({5,6})", wanted),
                ("E12", "{=ONE({5,6})", wanted),
                ("E13", "~=ONE({7})", "7"),
                ("E14", "=SUMPRODUCT(ONE({1,2}))", wanted),
            ],
        );
    }

    #[test]
    fn a_cell_holds_one_value_of_an_answer_of_many_by_its_kind() {
        holds(&[
            ("E10", "~=A1:A3*2", "(not-computed \"a spill\")"),
            ("E11", "{=A1:A3*2", "2"),
            ("E12", "{=SUM(A1:A3*B1:B3)", "32"),
            ("E13", "=TRANSPOSE(A1:A3)", "1"),
            ("E14", "{=A1:A3", "1"),
            ("E15", "~=SINGLE({5,6})", "5"),
        ]);
    }

    #[test]
    fn a_range_between_expressions_spans_its_ends() {
        holds(&[
            ("E10", "=SUM(A1:INDEX(A1:A3,2))", "3"),
            ("E11", "=SUM(INDEX(A1:A3,2):A3)", "5"),
            ("E12", "=ROWS(A1:INDEX(A:A,3))", "3"),
            ("E13", "=SUM(A1:INDEX(B1:B3,3))", "21"),
            ("E14", "=SUM(IF(TRUE,A1,B1):A2)", "3"),
            ("E15", "=SUM(A1:CHOOSE(2,A1,B2))", "12"),
            ("E16", "=SUM(A1:INDEX(A1:A3,4))", "(error \"#REF!\")"),
        ]);
    }

    #[test]
    fn a_range_whose_end_no_plan_bounds_is_labelled() {
        // `:` with a defined name at an end: the plan's edges for the range
        // are its reach, which a name's text does not give, so the range is
        // labelled rather than computed in an order that never read it.
        let mut wb = book(&[
            ("E10", "=SUM(A1:Ends)"),
            ("E11", "=SUM(A1:INDEX(Ends,1))"),
            ("E12", "=SUM(Ends)"),
        ]);
        wb.defined_names
            .push(("Ends".to_string(), "S!$A$3".to_string()));
        let lib = library::language();
        let calc = Calc::run(&wb, &lib);
        let at = |addr: &str| {
            let r = parse_a1_range(addr).unwrap();
            calc.computed((0, r.top, r.left)).map(|c| c.spell())
        };
        let unbounded = "(not-computed \"a range between computed ends\")";
        assert_eq!(at("E10").as_deref(), Some(unbounded));
        assert_eq!(at("E11").as_deref(), Some(unbounded));
        assert_eq!(at("E12").as_deref(), Some("3"));
    }

    #[test]
    fn the_array_functions_answer_as_excel_documents_them() {
        holds(&[
            ("E10", "=SUM(SEQUENCE(3))", "6"),
            ("E11", "=INDEX(SEQUENCE(2,3,10,10),2,3)", "60"),
            ("E12", "=SEQUENCE(0)", "(error \"#CALC!\")"),
            ("E13", "=SUM(FILTER(A1:A3,A1:A3>1))", "5"),
            ("E14", "=FILTER(A1:A3,A1:A3>5)", "(error \"#CALC!\")"),
            ("E15", "=FILTER(A1:A3,A1:A3>5,\"none\")", "\"none\""),
            ("E16", "=INDEX(SORT({3;1;2}),1)", "1"),
            ("E17", "=INDEX(SORT({3;1;2},1,-1),1)", "3"),
            (
                "E18",
                "=INDEX(SORTBY({\"a\";\"b\";\"c\"},{3;1;2}),1)",
                "\"b\"",
            ),
            ("E19", "=ROWS(UNIQUE({1;2;1;\"a\";\"A\"}))", "3"),
            ("E20", "=ROWS(UNIQUE({1;2;1},,TRUE))", "1"),
            ("E21", "=INDEX(TRANSPOSE({1,2;3,4}),1,2)", "3"),
            ("E22", "=INDEX(MMULT({1,2;3,4},{5;6}),2,1)", "39"),
            ("E23", "=MMULT({1,2},{1,2})", "(error \"#VALUE!\")"),
            ("E24", "=SUMPRODUCT({1,2},{3,4})", "11"),
            ("E25", "=SUMPRODUCT({1,2},{3})", "(error \"#VALUE!\")"),
            ("E26", "=SUMPRODUCT({1,\"a\"},{3,4})", "3"),
            ("E27", "=INDEX(CHOOSECOLS({1,2,3},-1),1)", "3"),
            ("E28", "=SUM(CHOOSEROWS({1;2;3},{1,3}))", "4"),
            ("E29", "=SUM(TAKE({1;2;3},2))", "3"),
            ("E30", "=SUM(TAKE({1;2;3},-1))", "3"),
            ("E31", "=TAKE({1;2},0)", "(error \"#CALC!\")"),
            ("E32", "=SUM(DROP({1;2;3},1))", "5"),
            ("E33", "=COLUMNS(HSTACK({1,2},{3}))", "3"),
            ("E34", "=INDEX(HSTACK({1;2},{3}),2,2)", "(error \"#N/A\")"),
            ("E35", "=ROWS(VSTACK({1,2},{3}))", "2"),
            ("E36", "=ROWS(TOCOL({1,2;3,4}))", "4"),
            ("E37", "=INDEX(TOCOL({1,2;3,4},0,TRUE),2)", "3"),
            ("E38", "=ROWS(TOCOL(C1:C3,1))", "2"),
            (
                "E39",
                "=INDEX(WRAPROWS({1,2,3,4,5},2),3,2)",
                "(error \"#N/A\")",
            ),
            ("E40", "=INDEX(WRAPCOLS({1,2,3,4,5},2,0),2,3)", "0"),
            ("E41", "=INDEX(EXPAND({1},2,2,0),2,2)", "0"),
            ("E42", "=EXPAND({1,2},1,1)", "(error \"#VALUE!\")"),
            ("E43", "=ROWS(A1:B3)", "3"),
            ("E44", "=COLUMNS(A1:B3)", "2"),
            ("E45", "=XLOOKUP(2,{1,2,3},{\"x\",\"y\",\"z\"})", "\"y\""),
            ("E46", "=SUM(XLOOKUP({1,3},{1,2,3},{10,20,30}))", "40"),
            ("E47", "=VLOOKUP(2,{1,\"a\";2,\"b\"},2,FALSE)", "\"b\""),
            ("E48", "=MATCH(2,{1,2,3},0)", "2"),
            ("E49", "=LOOKUP(2.5,{1,2,3},{\"a\",\"b\",\"c\"})", "\"b\""),
            // A number keeps its row but 0; an empty value sorts last; the
            // kinds in Excel's order, numbers, texts, truth values, errors.
            ("E50", "=ROWS(FILTER({1;2;3},{1;0;2}))", "2"),
            ("E51", "=INDEX(SORT(C1:C3),1)", "\"a\""),
            ("E52", "=INDEX(SORT({#N/A;TRUE;1;\"a\"}),3)", "true"),
        ]);
    }

    #[test]
    fn an_array_past_the_grid_s_limit_is_labelled() {
        let past = "(not-computed \"an array past the grid's limit\")";
        holds(&[
            ("E10", "=SUM(SEQUENCE(2000000))", past),
            ("E11", "=SUM(SEQUENCE(1000000))", "500000500000"),
            ("E12", "=ROWS(EXPAND(1,2000,2000))", past),
        ]);
    }
}
