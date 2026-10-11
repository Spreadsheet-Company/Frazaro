//! A shape evaluated at many of its cells at once (KERNEL.20, slice 3,
//! 2026-10-10): the vectorized evaluation of a shared formula, the first of
//! the structural advantages `KERNEL.20` names.
//!
//! An engine's machine evaluates a shape at each of its cells
//! ([`super::shape`]), and a cell's evaluation walks the parse: each node
//! matched, each call's function looked up by name, each operand made and
//! dropped, once a cell. Here the walk is once a run. A run is cells of one
//! shape on one sheet, consecutive in the order of evaluation, none reading
//! another, so that any of them may be computed before any other
//! (`Shapes::plan` finds them, in an order that puts a shape's cells of one
//! level together); each node is evaluated at every cell of the run before
//! its parent is, into a lane, one answer a cell.
//!
//! The meaning is the cell by cell evaluation's, and every rule of it is
//! written once. A value is made by [`binary`], over two numbers by its own
//! [`arithmetic`] and [`compares`], or by the library: a strict function
//! called at each cell with the operands it would be called with there, or,
//! for the language's aggregates and `NOT`, the library's own rule and fold
//! (`library::Fold`) driven here over the run's operands, a range of a sheet
//! kept as an array walked by index (KERNEL.26). A reason propagates as it
//! does there, each cell's answer the first its own evaluation meets, and
//! `IF` computes each branch at the cells that take it and at no other. An
//! operand computed at a cell whose answer an earlier operand has already
//! decided is computed for nothing and changes nothing, an evaluation having
//! no effect but its answer. A constant, and a reference every `$` holds, is
//! computed once for the run. What is not computed here is computed at each
//! cell by [`Ctx`] itself: a defined name, a form other than the language's
//! `IF`, a range where one value is wanted, and a call with such an operand.
//!
//! A lane holds each cell's answer as a value, or, while every answer is
//! one, as a number or a truth value (KERNEL.26): what a read of numbers,
//! the arithmetic and the comparisons over them, the aggregates and `IF`
//! make and read with no value's kind to match, a lane of values from the
//! first answer of another kind.
//!
//! Its buffers are its caller's, kept from one frame to the next ([`Lanes`])
//! and made in the first frame a run is evaluated in, as many as its parse
//! can hold at once whatever branch its cells take ([`holds`]), so that a
//! later frame makes none; a value the same at every cell is copied as often
//! as alone. It recurses once a level of nesting as [`Ctx`] does, a level's
//! frame holding no array, since `VLA.1`'s limits are pinned on a stack of
//! 1 MiB. Bit for bit with the cell by cell evaluation, which the machine's
//! tests hold it to over generated grids of every kind of node, and with
//! `Calc`, which plans and parses on its own.

use std::cell::RefCell;
use std::cmp::Ordering;
use std::convert::Infallible;

use crate::refers::Shape as Part;

use super::dense::Dense;
use super::eval::{
    arithmetic, binary, compare_numbers, compares, finite, Arg, Computed, Ctx, Env, Known, Library,
    Reason, Strict,
};
use super::formula::{BinOp, Expr, RefExpr};
use super::graph::{Area, CellId};
use super::library::{self, Fold, Ranged, Rule};
use super::value::{ErrorKind, Value};

/// The most cells evaluated at once: a lane of them is a few kilobytes, so
/// that the lanes a formula holds at once stay in the processor's cache.
pub(crate) const CHUNK: usize = 256;

/// What an expression answers at one cell, as [`Ctx::scalar`] answers it.
type Got = Result<Value, Reason>;

/// What an expression answers at each cell of a selection.
enum Lane {
    /// The same at every cell, computed once: a constant, or a reference
    /// every `$` holds.
    Same(Got),
    /// Each cell's own, in the selection's order.
    Each(Vec<Got>),
    /// Each cell's own, every one a number, held as one (KERNEL.26): what a
    /// read of numbers, the arithmetic over them and `SUM`, `MIN` and `MAX`
    /// answer, until a value of another kind turns the lane into
    /// [`Lane::Each`].
    Numbers(Vec<f64>),
    /// Each cell's own, every one a truth value (KERNEL.26): what a
    /// comparison of numbers, `AND`, `OR` and `NOT` answer.
    Truths(Vec<bool>),
}

impl Lane {
    /// What `f` makes of the answer at the selection's `k`th cell, the
    /// answer lent where it stands, a number or a truth value made a value
    /// for the while.
    #[inline]
    fn with<R>(&self, k: usize, f: impl FnOnce(Result<&Value, &Reason>) -> R) -> R {
        match self {
            Lane::Same(got) => f(got.as_ref()),
            Lane::Each(each) => f(each[k].as_ref()),
            Lane::Numbers(ns) => f(Ok(&Value::Number(ns[k]))),
            Lane::Truths(ts) => f(Ok(&Value::Bool(ts[k]))),
        }
    }

    /// The answer at the selection's `k`th cell, moved out of the lane,
    /// which nothing reads there again, or copied when it is the same at
    /// every cell; the copy at the `last` cell moved, so that the copies are
    /// as many as alone.
    fn take_at(&mut self, k: usize, last: bool) -> Got {
        match self {
            Lane::Same(got) if last => std::mem::replace(got, Ok(Value::Empty)),
            Lane::Same(got) => got.clone(),
            Lane::Each(each) => std::mem::replace(&mut each[k], Ok(Value::Empty)),
            Lane::Numbers(ns) => Ok(Value::Number(ns[k])),
            Lane::Truths(ts) => Ok(Value::Bool(ts[k])),
        }
    }
}

/// A lane as a lane of values holds it (KERNEL.26): what every node but
/// those over numbers and truth values reads, a lane of numbers or of truth
/// values made one first.
enum Values {
    Same(Got),
    Each(Vec<Got>),
}

/// Whether an operator is arithmetic, whose answer over two numbers is a
/// number or an error; the others but `&` compare.
fn is_arithmetic(op: BinOp) -> bool {
    matches!(
        op,
        BinOp::Add | BinOp::Sub | BinOp::Mul | BinOp::Div | BinOp::Pow
    )
}

/// The buffers an evaluation of many cells holds, kept by its caller from
/// one run to the next so that they are made once (KERNEL.20): of values,
/// of numbers and of truth values (KERNEL.26), each as many as the parse
/// holds lanes at once, and of positions.
#[derive(Debug, Default)]
pub(crate) struct Lanes {
    values: Vec<Vec<Got>>,
    numbers: Vec<Vec<f64>>,
    truths: Vec<Vec<bool>>,
    positions: Vec<Vec<u32>>,
}

impl Lanes {
    /// As many buffers made as a run's expression holds at once
    /// ([`holds`]), before the run begins: a lane of values, of numbers or
    /// of truth values each, whichever its cells' answers come to be.
    fn reserve(&mut self, (values, positions): (usize, usize)) {
        while self.values.len() < values {
            self.values.push(Vec::with_capacity(CHUNK));
        }
        while self.numbers.len() < values {
            self.numbers.push(Vec::with_capacity(CHUNK));
        }
        while self.truths.len() < values {
            self.truths.push(Vec::with_capacity(CHUNK));
        }
        while self.positions.len() < positions {
            self.positions.push(Vec::with_capacity(CHUNK));
        }
    }

    /// How many buffers of each kind are made, values, numbers, truth values
    /// and positions, and the most any holds.
    #[cfg(test)]
    pub(crate) fn made(&self) -> (usize, usize, usize, usize, usize) {
        let most = self.values.iter().map(Vec::capacity);
        let most = most
            .chain(self.numbers.iter().map(Vec::capacity))
            .chain(self.truths.iter().map(Vec::capacity))
            .chain(self.positions.iter().map(Vec::capacity))
            .max();
        (
            self.values.len(),
            self.numbers.len(),
            self.truths.len(),
            self.positions.len(),
            most.unwrap_or(0),
        )
    }
}

/// The most buffers an evaluation of the expression at many cells holds at
/// once, of values and of positions, whichever branches its cells take. A
/// run makes them before it begins, so that the first frame it is evaluated
/// in makes every buffer a later frame holds, and a later frame, whose cells
/// may take a branch no cell took before, makes none. An upper bound,
/// walked as the evaluation recurses and round a loop along a chain.
pub(crate) fn holds(lib: &Library, e: &Expr) -> (usize, usize) {
    match e {
        Expr::Number(_) | Expr::Text(_) | Expr::Bool(_) | Expr::Error(_) | Expr::Empty => (0, 0),
        Expr::Ref(_) | Expr::Name { .. } => (1, 0),
        Expr::Neg(x) | Expr::Pos(x) | Expr::Percent(x) => holds(lib, x),
        Expr::Binary(..) => {
            // The accumulated lane held while each right operand is evaluated.
            let (mut values, mut positions) = (0, 0);
            let mut left = e;
            while let Expr::Binary(_, l, r) = left {
                let (v, p) = holds(lib, r);
                values = values.max(1 + v);
                positions = positions.max(p);
                left = l;
            }
            let (v, p) = holds(lib, left);
            (values.max(v), positions.max(p))
        }
        Expr::Call(name, args) => {
            if lib.form(name).is_some() {
                if lib.known(name) != Some(Known::If) || !(2..=3).contains(&args.len()) {
                    return (1, 0);
                }
                // The condition's lane, and the answers' beside it while the
                // condition is a lane of numbers or truth values (KERNEL.26),
                // then each branch beside the answers and the three lists of
                // positions that place its answers.
                let (mut values, mut positions) = holds(lib, &args[0]);
                values = values.max(2);
                for a in &args[1..] {
                    let (v, p) = holds(lib, a);
                    values = values.max(1 + v);
                    positions = positions.max(3 + p);
                }
                (values, positions)
            } else if lib.strict(name).is_none() {
                (0, 0)
            } else if !args.iter().all(|a| is_operand(lib, a)) {
                (1, 0)
            } else {
                // Each operand beside those before it, then the answers.
                let (mut values, mut positions) = (args.len() + 1, 0);
                for (j, a) in args.iter().enumerate() {
                    let (v, p) = holds(lib, a);
                    values = values.max(j + v);
                    positions = positions.max(p);
                }
                (values, positions)
            }
        }
    }
}

/// Whether an argument can be a strict call's operand in an evaluation of
/// many cells: a reference written as one, or anything whose evaluation is a
/// value at every cell, never an area.
fn is_operand(lib: &Library, a: &Expr) -> bool {
    match a {
        Expr::Ref(_) => true,
        Expr::Pos(x) => is_operand(lib, x),
        _ => yields_value(lib, a),
    }
}

/// The reference an operand is made from at every cell, when it is one: a
/// reference written as one, under any unary plus, which returns its operand
/// as it is.
fn ref_operand<'e>(lib: &Library, a: &'e Expr) -> Option<&'e RefExpr> {
    match a {
        Expr::Ref(r) => Some(r),
        Expr::Pos(x) if !yields_value(lib, a) => ref_operand(lib, x),
        _ => None,
    }
}

/// Whether [`Ctx::eval`] of an expression is a value at every cell: anything
/// but a reference, a name, or a form whose answer may be one.
fn yields_value(lib: &Library, e: &Expr) -> bool {
    match e {
        Expr::Ref(_) | Expr::Name { .. } => false,
        Expr::Pos(x) => yields_value(lib, x),
        Expr::Call(name, args) => match lib.known(name) {
            Some(Known::If) => args.iter().skip(1).all(|a| yields_value(lib, a)),
            _ => lib.form(name).is_none(),
        },
        _ => true,
    }
}

/// The cells of the run a lane is for: every one, or some, by position.
#[derive(Clone, Copy)]
enum Sel<'s> {
    All(usize),
    Some(&'s [u32]),
}

impl Sel<'_> {
    fn len(self) -> usize {
        match self {
            Sel::All(n) => n,
            Sel::Some(at) => at.len(),
        }
    }

    /// The position in the run of the selection's `k`th cell.
    fn at(self, k: usize) -> usize {
        match self {
            Sel::All(_) => k,
            Sel::Some(at) => at[k] as usize,
        }
    }
}

/// An operand of a strict call at every cell: a reference written as one,
/// made at each cell as [`Ctx::eval`] makes it, or a value.
enum Operand<'e> {
    Ref(&'e RefExpr),
    Value(Lane),
}

/// The most operands of an aggregate the run computes itself (KERNEL.26);
/// a call of more is left to the library's function at each cell.
const COUNTED: usize = 8;

/// How `IF`'s conditions part the cells of a selection.
enum Parted {
    /// Every cell takes the one branch.
    All(bool),
    /// The answer at every cell, the condition a reason or an error.
    Answer(Lane),
    /// Each cell's own: the answers' lane, a place held for each, and the
    /// cells that take each branch, by place in the selection.
    Cells(Lane, Vec<u32>, Vec<u32>),
}

/// One run's evaluation.
struct Many<'a> {
    env: &'a dyn Env,
    lib: &'a Library,
    /// The run's cells, in the order of evaluation, all on one sheet.
    cells: &'a [CellId],
    /// The cell the shape was read at, from which each cell's offset is.
    anchor: CellId,
    lanes: RefCell<&'a mut Lanes>,
}

/// A shape's parse evaluated at each of a run's cells, the cells read at
/// `anchor`'s offset from each, each cell's answer pushed to `out` in the
/// run's order as the shape's evaluation at the cell alone answers it
/// (`Shapes::evaluate`); `held` is what [`holds`] says of the parse.
#[allow(clippy::too_many_arguments)]
pub(crate) fn evaluate(
    env: &dyn Env,
    lib: &Library,
    expr: &Expr,
    held: (usize, usize),
    anchor: CellId,
    cells: &[CellId],
    lanes: &mut Lanes,
    out: &mut Vec<Computed>,
) {
    lanes.reserve(held);
    let many = Many {
        env,
        lib,
        cells,
        anchor,
        lanes: RefCell::new(lanes),
    };
    match many.scalars(expr, Sel::All(cells.len())) {
        Lane::Same(got) => {
            // The last cell takes the answer itself, the others a copy, as
            // many as alone.
            if let Some((_, rest)) = cells.split_last() {
                for _ in rest {
                    out.push(computed(got.clone()));
                }
                out.push(computed(got));
            }
        }
        Lane::Each(mut each) => {
            out.extend(each.drain(..).map(computed));
            many.give(each);
        }
        Lane::Numbers(numbers) => {
            out.extend(numbers.iter().map(|&n| Computed::Value(Value::Number(n))));
            many.give_numbers(numbers);
        }
        Lane::Truths(truths) => {
            out.extend(truths.iter().map(|&b| Computed::Value(Value::Bool(b))));
            many.give_truths(truths);
        }
    }
}

/// A cell's answer as its formula holds it.
fn computed(got: Got) -> Computed {
    match got {
        Ok(v) => Computed::Value(v.settled()),
        Err(r) => Computed::NotComputed(r),
    }
}

impl<'a> Many<'a> {
    /// A cell's offset from the cell the shape was read at.
    fn offset(&self, pos: usize) -> (i64, i64) {
        let (_, row, col) = self.cells[pos];
        (
            i64::from(row) - i64::from(self.anchor.1),
            i64::from(col) - i64::from(self.anchor.2),
        )
    }

    /// The evaluation at one cell of the run, the environment told it is
    /// the cell being computed.
    fn ctx(&self, pos: usize) -> Ctx<'a> {
        self.env.enter(self.cells[pos]);
        let (d_row, d_col) = self.offset(pos);
        Ctx::at(self.env, self.lib, d_row, d_col)
    }

    fn take(&self) -> Vec<Got> {
        let kept = self.lanes.borrow_mut().values.pop();
        kept.unwrap_or_else(|| Vec::with_capacity(CHUNK))
    }

    fn give(&self, mut lane: Vec<Got>) {
        lane.clear();
        self.lanes.borrow_mut().values.push(lane);
    }

    fn take_numbers(&self) -> Vec<f64> {
        let kept = self.lanes.borrow_mut().numbers.pop();
        kept.unwrap_or_else(|| Vec::with_capacity(CHUNK))
    }

    fn give_numbers(&self, mut lane: Vec<f64>) {
        lane.clear();
        self.lanes.borrow_mut().numbers.push(lane);
    }

    fn take_truths(&self) -> Vec<bool> {
        let kept = self.lanes.borrow_mut().truths.pop();
        kept.unwrap_or_else(|| Vec::with_capacity(CHUNK))
    }

    fn give_truths(&self, mut lane: Vec<bool>) {
        lane.clear();
        self.lanes.borrow_mut().truths.push(lane);
    }

    /// A lane's buffer given back, whatever it holds.
    fn give_lane(&self, lane: Lane) {
        match lane {
            Lane::Same(_) => {}
            Lane::Each(each) => self.give(each),
            Lane::Numbers(numbers) => self.give_numbers(numbers),
            Lane::Truths(truths) => self.give_truths(truths),
        }
    }

    /// A lane as values, each cell's answer as the lane of values holds it
    /// (KERNEL.26).
    fn values(&self, lane: Lane) -> Values {
        match lane {
            Lane::Same(got) => Values::Same(got),
            Lane::Each(each) => Values::Each(each),
            Lane::Numbers(numbers) => {
                let mut each = self.take();
                each.extend(numbers.iter().map(|&n| Ok(Value::Number(n))));
                self.give_numbers(numbers);
                Values::Each(each)
            }
            Lane::Truths(truths) => {
                let mut each = self.take();
                each.extend(truths.iter().map(|&b| Ok(Value::Bool(b))));
                self.give_truths(truths);
                Values::Each(each)
            }
        }
    }

    /// [`Many::values`], as a lane.
    fn valued(&self, lane: Lane) -> Lane {
        match self.values(lane) {
            Values::Same(got) => Lane::Same(got),
            Values::Each(each) => Lane::Each(each),
        }
    }

    /// An arithmetic operator at each cell of a lane of numbers, its left
    /// operand or its right, `f` the operator's answer at the `k`th cell over
    /// the number there: each answer written where its operand stood, a lane
    /// of numbers while each is a number (KERNEL.26).
    fn arithmetic_in_place(&self, mut numbers: Vec<f64>, f: impl Fn(usize, f64) -> Value) -> Lane {
        for k in 0..numbers.len() {
            match f(k, numbers[k]) {
                Value::Number(n) => numbers[k] = n,
                other => {
                    let mut each = self.take();
                    each.extend(numbers[..k].iter().map(|&n| Ok(Value::Number(n))));
                    each.push(Ok(other));
                    each.extend((k + 1..numbers.len()).map(|j| Ok(f(j, numbers[j]))));
                    self.give_numbers(numbers);
                    return Lane::Each(each);
                }
            }
        }
        Lane::Numbers(numbers)
    }

    /// A comparison at each of `n` cells, `ordering` how its operands compare
    /// at the `k`th: a lane of truth values (KERNEL.26).
    fn compared(&self, op: BinOp, n: usize, ordering: impl Fn(usize) -> Ordering) -> Lane {
        let mut truths = self.take_truths();
        truths.extend((0..n).map(|k| compares(op, ordering(k))));
        Lane::Truths(truths)
    }

    /// One more cell's answer pushed to a lane being made, which stays a
    /// lane of numbers or of truth values while every answer is one, and
    /// becomes a lane of values at the first that is not (KERNEL.26).
    fn push(&self, lane: &mut Lane, got: Got) {
        match (&mut *lane, got) {
            (Lane::Each(each), got) => each.push(got),
            (Lane::Numbers(numbers), Ok(Value::Number(n))) => numbers.push(n),
            (Lane::Truths(truths), Ok(Value::Bool(b))) => truths.push(b),
            (_, got) => {
                let made = std::mem::replace(lane, Lane::Same(Ok(Value::Empty)));
                *lane = self.valued(made);
                if let Lane::Each(each) = lane {
                    each.push(got);
                }
            }
        }
    }

    fn positions(&self) -> Vec<u32> {
        let kept = self.lanes.borrow_mut().positions.pop();
        kept.unwrap_or_else(|| Vec::with_capacity(CHUNK))
    }

    fn give_positions(&self, mut at: Vec<u32>) {
        at.clear();
        self.lanes.borrow_mut().positions.push(at);
    }

    /// [`Ctx::scalar`] at every cell of the selection.
    fn scalars(&self, e: &Expr, sel: Sel) -> Lane {
        match e {
            Expr::Number(v) => Lane::Same(Ok(Value::Number(*v))),
            Expr::Text(t) => Lane::Same(Ok(Value::Text(t.clone()))),
            Expr::Bool(b) => Lane::Same(Ok(Value::Bool(*b))),
            Expr::Error(k) => Lane::Same(Ok(Value::Error(*k))),
            Expr::Empty => Lane::Same(Ok(Value::Empty)),
            Expr::Ref(r) if matches!(r.shape, Part::Cell) && r.last_sheet.is_none() => {
                self.reads(r, sel)
            }
            Expr::Neg(x) => self.unary(x, sel, |n| -n),
            // Unary plus is its operand, which a scalar reads as one value.
            Expr::Pos(x) => self.scalars(x, sel),
            Expr::Percent(x) => self.unary(x, sel, |n| n / 100.0),
            Expr::Binary(..) => self.chain(e, sel),
            Expr::Call(name, args) => self.call(e, name, args, sel),
            // A range where one value is wanted, a defined name: each cell's.
            Expr::Ref(_) | Expr::Name { .. } => self.each_cell(e, sel),
        }
    }

    /// A reference to one cell as written, read at every cell: its sheet
    /// found as `graph::areas_at` finds it, the same at every cell of the
    /// run; read once when every `$` holds it; a sheet the environment keeps
    /// as an array read by index, an empty slot through the environment,
    /// where a reason may stand beside it. The values read are a lane of
    /// numbers while every one is a number (KERNEL.26).
    fn reads(&self, r: &RefExpr, sel: Sel) -> Lane {
        if sel.len() == 0 {
            return Lane::Each(self.take());
        }
        let sheet = match (r.found, &r.sheet) {
            (Some((first, _)), _) => usize::from(first),
            (None, None) => self.cells[sel.at(0)].0,
            (None, Some(name)) => match self.env.model().find_sheet(name) {
                Some(s) => s,
                None => return Lane::Same(Ok(Value::Error(ErrorKind::Ref))),
            },
        };
        let dense = self.env.dense(sheet);
        // A number read where it stands, or the answer of any other kind.
        let at = |k: usize| -> Result<f64, Got> {
            let (d_row, d_col) = self.offset(sel.at(k));
            let Some((row, col, _, _)) = r.moved(d_row, d_col) else {
                return Err(Ok(Value::Error(ErrorKind::Ref)));
            };
            match dense.and_then(|d| d.get(row, col)) {
                Some(Value::Number(n)) => Ok(*n),
                Some(v) if !matches!(v, Value::Empty) => Err(Ok(v.clone())),
                _ => match self.env.read((sheet, row, col)) {
                    Ok(Value::Number(n)) => Ok(n),
                    got => Err(got),
                },
            }
        };
        let got = |k: usize| -> Got {
            match at(k) {
                Ok(n) => Ok(Value::Number(n)),
                Err(other) => other,
            }
        };
        if r.corners.row_abs1 && r.corners.col_abs1 {
            return Lane::Same(got(0));
        }
        let mut numbers = self.take_numbers();
        for k in 0..sel.len() {
            match at(k) {
                Ok(n) => numbers.push(n),
                Err(other) => {
                    let mut out = self.take();
                    out.extend(numbers.iter().map(|&n| Ok(Value::Number(n))));
                    self.give_numbers(numbers);
                    out.push(other);
                    out.extend((k + 1..sel.len()).map(got));
                    return Lane::Each(out);
                }
            }
        }
        Lane::Numbers(numbers)
    }

    /// A sign or a percent at every cell, over the number each value reads as.
    fn unary(&self, x: &Expr, sel: Sel, f: fn(f64) -> f64) -> Lane {
        let lane = self.scalars(x, sel);
        self.signed(lane, f)
    }

    /// [`Many::unary`] over its operand's lane, in a frame of its own, so
    /// that nothing here stands beneath the evaluation of a deeper level.
    #[inline(never)]
    fn signed(&self, lane: Lane, f: fn(f64) -> f64) -> Lane {
        let number = |v: &mut Value| {
            *v = match v.number() {
                Ok(n) => finite(f(n)),
                Err(k) => Value::Error(k),
            }
        };
        match lane {
            // Over numbers, as numbers while each answer is one (KERNEL.26).
            Lane::Numbers(numbers) => self.numbers_mapped(numbers, |n| finite(f(n))),
            lane => match self.values(lane) {
                Values::Same(mut got) => {
                    if let Ok(v) = got.as_mut() {
                        number(v);
                    }
                    Lane::Same(got)
                }
                Values::Each(mut each) => {
                    for v in each.iter_mut().flatten() {
                        number(v);
                    }
                    Lane::Each(each)
                }
            },
        }
    }

    /// A lane of numbers, each made a value by `f`: a lane of numbers while
    /// each answer is a number, a lane of values from the first that is not
    /// (KERNEL.26).
    fn numbers_mapped(&self, mut numbers: Vec<f64>, f: impl Fn(f64) -> Value) -> Lane {
        for k in 0..numbers.len() {
            match f(numbers[k]) {
                Value::Number(n) => numbers[k] = n,
                other => {
                    let mut each = self.take();
                    each.extend(numbers[..k].iter().map(|&n| Ok(Value::Number(n))));
                    each.push(Ok(other));
                    each.extend(numbers[k + 1..].iter().map(|&n| Ok(f(n))));
                    self.give_numbers(numbers);
                    return Lane::Each(each);
                }
            }
        }
        Lane::Numbers(numbers)
    }

    /// A binary operator, or a chain of them, at every cell: the leftmost
    /// operand, then each link in turn, as `Ctx::chain` goes round its left
    /// spine, the right operands on a stack.
    fn chain(&self, e: &Expr, sel: Sel) -> Lane {
        const NEAR: usize = 16;
        let mut near: [(BinOp, &Expr); NEAR] = [(BinOp::Add, e); NEAR];
        let mut far: Vec<(BinOp, &Expr)> = Vec::new();
        let mut links = 0;
        let mut left = e;
        while let Expr::Binary(op, l, r) = left {
            if links < NEAR {
                near[links] = (*op, r);
            } else {
                far.push((*op, r));
            }
            links += 1;
            left = l;
        }
        let mut acc = self.scalars(left, sel);
        for i in (0..links).rev() {
            let (op, r) = if i < NEAR { near[i] } else { far[i - NEAR] };
            let b = self.scalars(r, sel);
            acc = self.combine(op, acc, b);
        }
        acc
    }

    /// [`binary`] at every cell: the left operand's reason first, then the
    /// right's, then the operator over the two values; over two lanes of
    /// numbers, or one and a number the same at every cell, the operator's
    /// rule over numbers, `binary`'s own (KERNEL.26).
    fn combine(&self, op: BinOp, a: Lane, b: Lane) -> Lane {
        let (a, b) = match (op, a, b) {
            (BinOp::Concat, a, b) => (self.values(a), self.values(b)),
            (op, Lane::Numbers(xs), Lane::Numbers(ys)) => {
                let lane = if is_arithmetic(op) {
                    self.arithmetic_in_place(xs, |k, x| arithmetic(op, x, ys[k]))
                } else {
                    let lane = self.compared(op, xs.len(), |k| compare_numbers(xs[k], ys[k]));
                    self.give_numbers(xs);
                    lane
                };
                self.give_numbers(ys);
                return lane;
            }
            (op, Lane::Numbers(xs), Lane::Same(Ok(Value::Number(y)))) => {
                return if is_arithmetic(op) {
                    self.arithmetic_in_place(xs, |_, x| arithmetic(op, x, y))
                } else {
                    let lane = self.compared(op, xs.len(), |k| compare_numbers(xs[k], y));
                    self.give_numbers(xs);
                    lane
                };
            }
            (op, Lane::Same(Ok(Value::Number(x))), Lane::Numbers(ys)) => {
                return if is_arithmetic(op) {
                    self.arithmetic_in_place(ys, |_, y| arithmetic(op, x, y))
                } else {
                    let lane = self.compared(op, ys.len(), |k| compare_numbers(x, ys[k]));
                    self.give_numbers(ys);
                    lane
                };
            }
            (_, a, b) => (self.values(a), self.values(b)),
        };
        match (a, b) {
            (Values::Same(x), Values::Same(y)) => Lane::Same(match (x, &y) {
                (Err(r), _) => Err(r),
                (Ok(_), Err(r)) => Err(r.clone()),
                (Ok(x), Ok(y)) => Ok(binary(op, &x, y)),
            }),
            (Values::Each(mut xs), Values::Same(y)) => {
                for x in xs.iter_mut() {
                    if let Ok(a) = x {
                        match &y {
                            Ok(b) => *a = binary(op, a, b),
                            Err(r) => *x = Err(r.clone()),
                        }
                    }
                }
                Lane::Each(xs)
            }
            (Values::Same(x), Values::Each(mut ys)) => {
                for y in ys.iter_mut() {
                    match &x {
                        Err(r) => *y = Err(r.clone()),
                        Ok(a) => {
                            if let Ok(b) = y {
                                *b = binary(op, a, b);
                            }
                        }
                    }
                }
                Lane::Each(ys)
            }
            (Values::Each(mut xs), Values::Each(mut ys)) => {
                for (x, y) in xs.iter_mut().zip(ys.iter_mut()) {
                    if let Ok(a) = x {
                        match y {
                            Ok(b) => *a = binary(op, a, b),
                            Err(r) => *x = Err(std::mem::replace(r, Reason::Cycle)),
                        }
                    }
                }
                self.give(ys);
                Lane::Each(xs)
            }
        }
    }

    /// A call at every cell: the language's `IF` here, another form at
    /// each cell, one of the language's aggregates or `NOT` here by the
    /// library's own rule and fold, any other strict function over its
    /// operands, a function the library does not hold named, in the order
    /// `Ctx` looks them up.
    fn call(&self, e: &Expr, name: &str, args: &[Expr], sel: Sel) -> Lane {
        if self.lib.form(name).is_some() {
            return match self.lib.known(name) {
                Some(Known::If) => self.if_(args, sel),
                _ => self.each_cell(e, sel),
            };
        }
        let Some(f) = self.lib.strict(name) else {
            return Lane::Same(Err(Reason::Function(name.to_string())));
        };
        if !args.iter().all(|a| is_operand(self.lib, a)) {
            return self.each_cell(e, sel);
        }
        match args {
            [a] => {
                let mut x = self.operand(a, sel);
                let lane = self.apply_one(f, name, &mut x, sel);
                self.drop_operand(x);
                lane
            }
            [a, b] => {
                let mut x = self.operand(a, sel);
                let mut y = self.operand(b, sel);
                let lane = self.apply_two(f, name, &mut x, &mut y, sel);
                self.drop_operand(x);
                self.drop_operand(y);
                lane
            }
            _ => {
                let mut ops: Vec<Operand> = args.iter().map(|a| self.operand(a, sel)).collect();
                let lane = self.apply_many(f, name, &mut ops, sel);
                for x in ops {
                    self.drop_operand(x);
                }
                lane
            }
        }
    }

    /// One of the language's own functions the run knows, over its operands
    /// computed (KERNEL.26), by the run itself when its operands are ones it
    /// walks.
    fn known_call(&self, name: &str, ops: &mut [Operand<'_>], sel: Sel) -> Option<Lane> {
        self.known(self.lib.known(name)?, ops, sel)
    }

    /// An operand at every cell, [`is_operand`] having said it can be one.
    fn operand<'e>(&self, a: &'e Expr, sel: Sel) -> Operand<'e> {
        match ref_operand(self.lib, a) {
            Some(r) => Operand::Ref(r),
            None => Operand::Value(self.scalars(a, sel)),
        }
    }

    /// One of the language's own functions the run computes itself
    /// (KERNEL.26): `SUM`, `MIN` and `MAX`, `AND` and `OR` by the library's
    /// rule and fold (`library::Fold`), over values and over ranges of sheets
    /// the environment keeps as arrays, and `NOT` over a value; `None` for a
    /// call left to the library's function at each cell: an operand naming a
    /// span of sheets or a range of a sheet walked through its cells, `NOT`
    /// of a reference, more operands than [`COUNTED`].
    fn known(&self, known: Known, ops: &mut [Operand<'_>], sel: Sel) -> Option<Lane> {
        if known == Known::Not {
            // The operand's lane taken, its values negated.
            let [Operand::Value(lane)] = ops else {
                return None;
            };
            let lane = std::mem::replace(lane, Lane::Same(Ok(Value::Empty)));
            return Some(self.negated(lane, sel));
        }
        if ops.len() > COUNTED {
            return None;
        }
        // Each reference's sheet, found as `graph::areas_at` finds it, the
        // same at every cell of the run, and its array; `None` for a sheet
        // the grid does not hold, where the reference is `#REF!`.
        let mut kept: [Option<(usize, &Dense)>; COUNTED] = [None; COUNTED];
        for (j, x) in ops.iter().enumerate() {
            let Operand::Ref(r) = x else {
                continue;
            };
            if r.last_sheet.is_some() {
                return None;
            }
            let sheet = match (r.found, &r.sheet) {
                (Some((first, _)), _) => Some(usize::from(first)),
                (None, None) => Some(self.cells[0].0),
                (None, Some(name)) => self.env.model().find_sheet(name),
            };
            if let Some(s) = sheet {
                kept[j] = Some((s, self.env.dense(s)?));
            }
        }
        let kept = &kept[..ops.len()];
        // The answers a lane of numbers, or of truth values, while each is.
        let numbers = || Lane::Numbers(self.take_numbers());
        let truths = || Lane::Truths(self.take_truths());
        Some(match known {
            Known::Sum => self.aggregate::<library::Sum>(ops, kept, sel, numbers()),
            Known::Min => self.aggregate::<library::Least>(ops, kept, sel, numbers()),
            Known::Max => self.aggregate::<library::Most>(ops, kept, sel, numbers()),
            Known::And => self.aggregate::<library::All>(ops, kept, sel, truths()),
            Known::Or => self.aggregate::<library::Some_>(ops, kept, sel, truths()),
            // A form, and `NOT` above.
            Known::If | Known::Not => return None,
        })
    }

    /// `NOT` at every cell, over the value its operand is there, as
    /// `library::negated` answers it: in place over a lane of values, a lane
    /// of truth values made over one of numbers or of truth values.
    fn negated(&self, lane: Lane, sel: Sel) -> Lane {
        match lane {
            Lane::Same(got) => Lane::Same(got.map(|v| library::negated(&v))),
            Lane::Each(mut each) => {
                for v in each.iter_mut().flatten() {
                    *v = library::negated(v);
                }
                Lane::Each(each)
            }
            typed => {
                let mut out = Lane::Truths(self.take_truths());
                for k in 0..sel.len() {
                    let got = typed.with(k, |g| g.map(library::negated).map_err(Reason::clone));
                    self.push(&mut out, got);
                }
                self.give_lane(typed);
                out
            }
        }
    }

    /// An aggregate at every cell, by its rule and its fold, as
    /// `library::folded` computes it at one; `kept` holds each reference's
    /// sheet and array, `out` the lane its answers are pushed to.
    #[inline(never)]
    fn aggregate<F: Fold>(
        &self,
        ops: &[Operand<'_>],
        kept: &[Option<(usize, &Dense)>],
        sel: Sel,
        mut out: Lane,
    ) -> Lane {
        for k in 0..sel.len() {
            let got = self.fold_at::<F>(ops, kept, sel.at(k), k);
            self.push(&mut out, got);
        }
        out
    }

    /// An aggregate at one cell: each operand taken left to right, as
    /// `take_arg` takes it, a reason the answer; then each value counted in
    /// the arguments' order as `library::each_item` counts it, a value given
    /// directly coerced and its error the answer at once, a range moved to the
    /// cell and walked by index to its end, its first error the answer.
    fn fold_at<F: Fold>(
        &self,
        ops: &[Operand<'_>],
        kept: &[Option<(usize, &Dense)>],
        pos: usize,
        k: usize,
    ) -> Got {
        for x in ops {
            if let Operand::Value(lane) = x {
                if let Some(r) = lane.with(k, |g| g.err().cloned()) {
                    return Err(r);
                }
            }
        }
        let (d_row, d_col) = self.offset(pos);
        let mut fold = F::default();
        for (x, kept) in ops.iter().zip(kept) {
            let r = match x {
                Operand::Value(lane) => {
                    match lane.with(k, |g| g.ok().map(F::Rule::given)) {
                        Some(Ok(item)) => fold.add(item),
                        Some(Err(e)) => return Ok(Value::Error(e)),
                        None => {}
                    }
                    continue;
                }
                Operand::Ref(r) => r,
            };
            let area = kept.and_then(|(sheet, array)| {
                let (top, left, bottom, right) = r.moved(d_row, d_col)?;
                let area = Area {
                    sheet,
                    top,
                    left,
                    bottom,
                    right,
                };
                Some((area, array))
            });
            // A reference that names no cell is `#REF!`, a value given.
            let Some((area, array)) = area else {
                match F::Rule::given(&Value::Error(ErrorKind::Ref)) {
                    Ok(item) => fold.add(item),
                    Err(e) => return Ok(Value::Error(e)),
                }
                continue;
            };
            let mut err: Option<ErrorKind> = None;
            let walked = array.each_ref::<Infallible, _>(&area, |v| {
                match F::Rule::ranged(v) {
                    Ranged::Counts(item) => fold.add(item),
                    Ranged::Error(e) => {
                        err.get_or_insert(e);
                    }
                    Ranged::Skipped => {}
                }
                Ok(())
            });
            let Ok(()) = walked;
            if let Some(e) = err {
                return Ok(Value::Error(e));
            }
        }
        Ok(fold.answer())
    }

    fn drop_operand(&self, x: Operand<'_>) {
        if let Operand::Value(lane) = x {
            self.give_lane(lane);
        }
    }

    /// A strict function of one operand at every cell, the operand on the
    /// stack in a frame of its own (as `Ctx::strict_one`'s): computed by the
    /// run itself when it is one of the language's that the run knows and
    /// its operand is one it walks (KERNEL.26), else by the function at each
    /// cell, its answers a lane of numbers while each is one.
    #[inline(never)]
    fn apply_one(&self, f: Strict, name: &str, x: &mut Operand<'_>, sel: Sel) -> Lane {
        if let Some(lane) = self.known_call(name, std::slice::from_mut(x), sel) {
            return lane;
        }
        let (mut out, n) = (Lane::Numbers(self.take_numbers()), sel.len());
        for k in 0..n {
            let ctx = self.ctx(sel.at(k));
            let got = match take_arg(x, k, k + 1 == n, &ctx) {
                Ok(a) => f(&ctx, std::slice::from_ref(&a)),
                Err(r) => Err(r),
            };
            self.push(&mut out, got);
        }
        out
    }

    /// Two operands, left to right, the first not computed the answer; the
    /// pair held side by side while the run asks whether it computes the call
    /// itself (KERNEL.26), each given back to its place after.
    #[inline(never)]
    fn apply_two<'e>(
        &self,
        f: Strict,
        name: &str,
        x: &mut Operand<'e>,
        y: &mut Operand<'e>,
        sel: Sel,
    ) -> Lane {
        let held = || Operand::Value(Lane::Same(Ok(Value::Empty)));
        let mut pair = [std::mem::replace(x, held()), std::mem::replace(y, held())];
        let known = self.known_call(name, &mut pair, sel);
        let [a, b] = pair;
        (*x, *y) = (a, b);
        if let Some(lane) = known {
            return lane;
        }
        let (mut out, n) = (Lane::Numbers(self.take_numbers()), sel.len());
        for k in 0..n {
            let ctx = self.ctx(sel.at(k));
            let got = match take_arg(x, k, k + 1 == n, &ctx) {
                Err(r) => Err(r),
                Ok(a) => match take_arg(y, k, k + 1 == n, &ctx) {
                    Err(r) => Err(r),
                    Ok(b) => f(&ctx, &[a, b]),
                },
            };
            self.push(&mut out, got);
        }
        out
    }

    /// Any other count, in a vector made once for the run.
    #[inline(never)]
    fn apply_many(&self, f: Strict, name: &str, ops: &mut [Operand<'_>], sel: Sel) -> Lane {
        if let Some(lane) = self.known_call(name, ops, sel) {
            return lane;
        }
        let (mut out, n) = (Lane::Numbers(self.take_numbers()), sel.len());
        let mut args = Vec::with_capacity(ops.len());
        for k in 0..n {
            let ctx = self.ctx(sel.at(k));
            args.clear();
            let mut stopped = None;
            for x in ops.iter_mut() {
                match take_arg(x, k, k + 1 == n, &ctx) {
                    Ok(a) => args.push(a),
                    Err(r) => {
                        stopped = Some(r);
                        break;
                    }
                }
            }
            let got = match stopped {
                Some(r) => Err(r),
                None => f(&ctx, &args),
            };
            self.push(&mut out, got);
        }
        out
    }

    /// The language's `IF` at every cell, as `library::if_` computes it at
    /// one: the condition at each, a reason or an error the cell's answer,
    /// then each branch at the cells that take it and at no other.
    fn if_(&self, args: &[Expr], sel: Sel) -> Lane {
        if args.len() < 2 || args.len() > 3 {
            return Lane::Same(Ok(Value::Error(ErrorKind::Value)));
        }
        let conditions = self.scalars(&args[0], sel);
        let (mut out, yes, no) = match self.parted(conditions, sel) {
            Parted::All(taken) => return self.branch(args, taken, sel),
            Parted::Answer(lane) => return lane,
            Parted::Cells(out, yes, no) => (out, yes, no),
        };
        self.fill(&mut out, args, true, sel, &yes);
        self.fill(&mut out, args, false, sel, &no);
        self.give_positions(yes);
        self.give_positions(no);
        out
    }

    /// How `IF`'s conditions at every cell of the selection part its cells,
    /// in a frame of its own, so that nothing here stands beneath the
    /// evaluation of a deeper level.
    #[inline(never)]
    fn parted(&self, conditions: Lane, sel: Sel) -> Parted {
        let (mut yes, mut no) = match conditions {
            Lane::Same(Err(r)) => return Parted::Answer(Lane::Same(Err(r))),
            Lane::Same(Ok(v)) => {
                return match v.truth() {
                    Ok(taken) => Parted::All(taken),
                    Err(k) => Parted::Answer(Lane::Same(Ok(Value::Error(k)))),
                }
            }
            _ => (self.positions(), self.positions()),
        };
        let out = match conditions {
            // The conditions' own lane holds the answers, a cell whose
            // condition is a reason or an error keeping it.
            Lane::Each(mut out) => {
                for (k, got) in out.iter_mut().enumerate() {
                    if let Ok(v) = got {
                        match v.truth() {
                            Ok(true) => yes.push(k as u32),
                            Ok(false) => no.push(k as u32),
                            Err(kind) => *v = Value::Error(kind),
                        }
                    }
                }
                Lane::Each(out)
            }
            // Over numbers or truth values, the answers a lane of numbers
            // while each branch's are (KERNEL.26), a place held for each.
            typed => {
                let mut out = Lane::Numbers(self.take_numbers());
                for k in 0..sel.len() {
                    match typed.with(k, |g| g.map(Value::truth).map_err(Reason::clone)) {
                        Ok(Ok(taken)) => {
                            if taken { &mut yes } else { &mut no }.push(k as u32);
                            self.push(&mut out, Ok(Value::Number(0.0)));
                        }
                        Ok(Err(kind)) => self.push(&mut out, Ok(Value::Error(kind))),
                        Err(r) => self.push(&mut out, Err(r)),
                    }
                }
                self.give_lane(typed);
                out
            }
        };
        Parted::Cells(out, yes, no)
    }

    /// One branch at the cells that take it, `ks` their places in the
    /// selection, each answer placed in `out`.
    fn fill(&self, out: &mut Lane, args: &[Expr], taken: bool, sel: Sel, ks: &[u32]) {
        if ks.is_empty() {
            return;
        }
        let lane = if ks.len() == sel.len() {
            self.branch(args, taken, sel)
        } else {
            let mut at = self.positions();
            at.extend(ks.iter().map(|&k| sel.at(k as usize) as u32));
            let lane = self.branch(args, taken, Sel::Some(&at));
            self.give_positions(at);
            lane
        };
        self.place(out, lane, sel, ks);
    }

    /// A branch's answers placed in `out` at the cells that take it, in a
    /// frame of its own as [`Many::parted`] is: a lane of numbers while the
    /// branch's answers are numbers (KERNEL.26), else a lane of values.
    #[inline(never)]
    fn place(&self, out: &mut Lane, lane: Lane, sel: Sel, ks: &[u32]) {
        if let Lane::Numbers(numbers) = out {
            match &lane {
                Lane::Same(Ok(Value::Number(n))) => {
                    for &k in ks {
                        numbers[k as usize] = *n;
                    }
                    return;
                }
                Lane::Numbers(answers) => {
                    for (j, &k) in ks.iter().enumerate() {
                        numbers[k as usize] = answers[j];
                    }
                    self.give_lane(lane);
                    return;
                }
                _ => {}
            }
        }
        let made = std::mem::replace(out, Lane::Same(Ok(Value::Empty)));
        let mut each = match self.values(made) {
            Values::Each(each) => each,
            Values::Same(got) => {
                let mut each = self.take();
                each.resize(sel.len(), got);
                each
            }
        };
        match lane {
            Lane::Same(got) => {
                if let Some((&last, rest)) = ks.split_last() {
                    for &k in rest {
                        each[k as usize] = got.clone();
                    }
                    each[last as usize] = got;
                }
            }
            mut lane => {
                for (j, &k) in ks.iter().enumerate() {
                    each[k as usize] = lane.take_at(j, false);
                }
                self.give_lane(lane);
            }
        }
        *out = Lane::Each(each);
    }

    /// The branch `IF` takes on a truth value, at the cells selected: a
    /// branch left out reads as 0 when taken, a missing else as FALSE.
    fn branch(&self, args: &[Expr], taken: bool, sel: Sel) -> Lane {
        match (taken, args.get(1), args.get(2)) {
            (true, Some(Expr::Empty), _) | (false, _, Some(Expr::Empty)) => {
                Lane::Same(Ok(Value::Number(0.0)))
            }
            (true, Some(then), _) => self.scalars(then, sel),
            (false, _, Some(otherwise)) => self.scalars(otherwise, sel),
            (false, _, None) => Lane::Same(Ok(Value::Bool(false))),
            (true, None, _) => Lane::Same(Ok(Value::Error(ErrorKind::Value))),
        }
    }

    /// An expression evaluated at each cell by [`Ctx`] itself.
    #[inline(never)]
    fn each_cell(&self, e: &Expr, sel: Sel) -> Lane {
        let mut out = self.take();
        for k in 0..sel.len() {
            out.push(self.ctx(sel.at(k)).scalar(e));
        }
        Lane::Each(out)
    }
}

/// A strict call's operand at the selection's `k`th cell: a reference made
/// there, a value moved out of its lane, which nothing reads again; a value
/// the same at every cell copied, and moved at the `last` cell, so that the
/// copies are as many as alone.
fn take_arg<'e>(
    x: &mut Operand<'e>,
    k: usize,
    last: bool,
    ctx: &Ctx<'_>,
) -> Result<Arg<'e>, Reason> {
    match x {
        Operand::Ref(r) => ctx.reference(r),
        Operand::Value(lane) => lane.take_at(k, last).map(Arg::Scalar),
    }
}
