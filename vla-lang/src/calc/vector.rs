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
//! The meaning is the cell by cell evaluation's. Every value is made by the
//! same functions, [`binary`] and the library's own; a reason propagates as
//! it does there, each cell's answer the first its own evaluation meets; a
//! strict call's function is called at each cell with the operands it would
//! be called with there; and `IF` computes each branch at the cells that
//! take it and at no other. An operand computed at a cell whose answer an
//! earlier operand has already decided is computed for nothing and changes
//! nothing, an evaluation having no effect but its answer. A constant, and a
//! reference every `$` holds, is computed once for the run. What is not
//! computed here is computed at each cell by [`Ctx`] itself: a defined name,
//! a form other than the language's `IF`, a range where one value is
//! wanted, and a call with such an operand.
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

use crate::refers::Shape as Part;

use super::eval::{binary, finite, Arg, Computed, Ctx, Env, Known, Library, Reason, Strict};
use super::formula::{BinOp, Expr, RefExpr};
use super::graph::CellId;
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
}

/// The buffers an evaluation of many cells holds, kept by its caller from
/// one run to the next so that they are made once (KERNEL.20).
#[derive(Debug, Default)]
pub(crate) struct Lanes {
    values: Vec<Vec<Got>>,
    positions: Vec<Vec<u32>>,
}

impl Lanes {
    /// As many buffers made as a run's expression holds at once
    /// ([`holds`]), before the run begins.
    fn reserve(&mut self, (values, positions): (usize, usize)) {
        while self.values.len() < values {
            self.values.push(Vec::with_capacity(CHUNK));
        }
        while self.positions.len() < positions {
            self.positions.push(Vec::with_capacity(CHUNK));
        }
    }

    /// How many buffers of each kind are made, and the most any holds.
    #[cfg(test)]
    pub(crate) fn made(&self) -> (usize, usize, usize) {
        let most = self.values.iter().map(Vec::capacity);
        let most = most.chain(self.positions.iter().map(Vec::capacity)).max();
        (self.values.len(), self.positions.len(), most.unwrap_or(0))
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
                // The condition's lane, then each branch beside it and the
                // three lists of positions that place its answers.
                let (mut values, mut positions) = holds(lib, &args[0]);
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

/// Whether [`Ctx::eval`] of an expression is a value at every cell: anything
/// but a reference, a name, or a form whose answer may be one.
fn yields_value(lib: &Library, e: &Expr) -> bool {
    match e {
        Expr::Ref(_) | Expr::Name { .. } => false,
        Expr::Pos(x) => yields_value(lib, x),
        Expr::Call(name, args) => match lib.known(name) {
            Some(Known::If) => args.iter().skip(1).all(|a| yields_value(lib, a)),
            None => lib.form(name).is_none(),
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
    /// where a reason may stand beside it.
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
        let at = |k: usize| -> Got {
            let (d_row, d_col) = self.offset(sel.at(k));
            let Some((row, col, _, _)) = r.moved(d_row, d_col) else {
                return Ok(Value::Error(ErrorKind::Ref));
            };
            match dense.and_then(|d| d.get(row, col)) {
                Some(v) if !matches!(v, Value::Empty) => Ok(v.clone()),
                _ => self.env.read((sheet, row, col)),
            }
        };
        if r.corners.row_abs1 && r.corners.col_abs1 {
            return Lane::Same(at(0));
        }
        let mut out = self.take();
        out.extend((0..sel.len()).map(at));
        Lane::Each(out)
    }

    /// A sign or a percent at every cell, over the number each value reads as.
    fn unary(&self, x: &Expr, sel: Sel, f: fn(f64) -> f64) -> Lane {
        let number = |v: &mut Value| {
            *v = match v.number() {
                Ok(n) => finite(f(n)),
                Err(k) => Value::Error(k),
            }
        };
        match self.scalars(x, sel) {
            Lane::Same(mut got) => {
                if let Ok(v) = got.as_mut() {
                    number(v);
                }
                Lane::Same(got)
            }
            Lane::Each(mut each) => {
                for v in each.iter_mut().flatten() {
                    number(v);
                }
                Lane::Each(each)
            }
        }
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
    /// right's, then the operator over the two values.
    fn combine(&self, op: BinOp, a: Lane, b: Lane) -> Lane {
        match (a, b) {
            (Lane::Same(x), Lane::Same(y)) => Lane::Same(match (x, &y) {
                (Err(r), _) => Err(r),
                (Ok(_), Err(r)) => Err(r.clone()),
                (Ok(x), Ok(y)) => Ok(binary(op, &x, y)),
            }),
            (Lane::Each(mut xs), Lane::Same(y)) => {
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
            (Lane::Same(x), Lane::Each(mut ys)) => {
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
            (Lane::Each(mut xs), Lane::Each(mut ys)) => {
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
    /// each cell, a strict function over its operands, a function the
    /// library does not hold named, in the order `Ctx` looks them up.
    fn call(&self, e: &Expr, name: &str, args: &[Expr], sel: Sel) -> Lane {
        if self.lib.form(name).is_some() {
            return match self.lib.known(name) {
                Some(Known::If) => self.if_(args, sel),
                None => self.each_cell(e, sel),
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
                let lane = self.apply_one(f, &mut x, sel);
                self.drop_operand(x);
                lane
            }
            [a, b] => {
                let mut x = self.operand(a, sel);
                let mut y = self.operand(b, sel);
                let lane = self.apply_two(f, &mut x, &mut y, sel);
                self.drop_operand(x);
                self.drop_operand(y);
                lane
            }
            _ => {
                let mut ops: Vec<Operand> = args.iter().map(|a| self.operand(a, sel)).collect();
                let lane = self.apply_many(f, &mut ops, sel);
                for x in ops {
                    self.drop_operand(x);
                }
                lane
            }
        }
    }

    /// An operand at every cell, [`is_operand`] having said it can be one.
    fn operand<'e>(&self, a: &'e Expr, sel: Sel) -> Operand<'e> {
        match a {
            Expr::Ref(r) => Operand::Ref(r),
            Expr::Pos(x) if !yields_value(self.lib, a) => self.operand(x, sel),
            _ => Operand::Value(self.scalars(a, sel)),
        }
    }

    fn drop_operand(&self, x: Operand<'_>) {
        if let Operand::Value(Lane::Each(each)) = x {
            self.give(each);
        }
    }

    /// A strict function of one operand at every cell, the operand on the
    /// stack in a frame of its own (as `Ctx::strict_one`'s).
    #[inline(never)]
    fn apply_one(&self, f: Strict, x: &mut Operand<'_>, sel: Sel) -> Lane {
        let (mut out, n) = (self.take(), sel.len());
        for k in 0..n {
            let ctx = self.ctx(sel.at(k));
            out.push(match take_arg(x, k, k + 1 == n, &ctx) {
                Ok(a) => f(&ctx, std::slice::from_ref(&a)),
                Err(r) => Err(r),
            });
        }
        Lane::Each(out)
    }

    /// Two operands, left to right, the first not computed the answer.
    #[inline(never)]
    fn apply_two(&self, f: Strict, x: &mut Operand<'_>, y: &mut Operand<'_>, sel: Sel) -> Lane {
        let (mut out, n) = (self.take(), sel.len());
        for k in 0..n {
            let ctx = self.ctx(sel.at(k));
            out.push(match take_arg(x, k, k + 1 == n, &ctx) {
                Err(r) => Err(r),
                Ok(a) => match take_arg(y, k, k + 1 == n, &ctx) {
                    Err(r) => Err(r),
                    Ok(b) => f(&ctx, &[a, b]),
                },
            });
        }
        Lane::Each(out)
    }

    /// Any other count, in a vector made once for the run.
    #[inline(never)]
    fn apply_many(&self, f: Strict, ops: &mut [Operand<'_>], sel: Sel) -> Lane {
        let (mut out, n) = (self.take(), sel.len());
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
            out.push(match stopped {
                Some(r) => Err(r),
                None => f(&ctx, &args),
            });
        }
        Lane::Each(out)
    }

    /// The language's `IF` at every cell, as `library::if_` computes it at
    /// one: the condition at each, a reason or an error the cell's answer,
    /// then each branch at the cells that take it and at no other.
    fn if_(&self, args: &[Expr], sel: Sel) -> Lane {
        if args.len() < 2 || args.len() > 3 {
            return Lane::Same(Ok(Value::Error(ErrorKind::Value)));
        }
        let mut out = match self.scalars(&args[0], sel) {
            Lane::Same(Err(r)) => return Lane::Same(Err(r)),
            Lane::Same(Ok(v)) => {
                return match v.truth() {
                    Ok(taken) => self.branch(args, taken, sel),
                    Err(k) => Lane::Same(Ok(Value::Error(k))),
                }
            }
            Lane::Each(conditions) => conditions,
        };
        let (mut yes, mut no) = (self.positions(), self.positions());
        for (k, got) in out.iter_mut().enumerate() {
            if let Ok(v) = got {
                match v.truth() {
                    Ok(true) => yes.push(k as u32),
                    Ok(false) => no.push(k as u32),
                    Err(kind) => *v = Value::Error(kind),
                }
            }
        }
        self.fill(&mut out, args, true, sel, &yes);
        self.fill(&mut out, args, false, sel, &no);
        self.give_positions(yes);
        self.give_positions(no);
        Lane::Each(out)
    }

    /// One branch at the cells that take it, `ks` their places in the
    /// selection, each answer written into its place in `out`.
    fn fill(&self, out: &mut [Got], args: &[Expr], taken: bool, sel: Sel, ks: &[u32]) {
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
        match lane {
            Lane::Same(got) => {
                if let Some((&last, rest)) = ks.split_last() {
                    for &k in rest {
                        out[k as usize] = got.clone();
                    }
                    out[last as usize] = got;
                }
            }
            Lane::Each(mut each) => {
                for (j, &k) in ks.iter().enumerate() {
                    out[k as usize] = std::mem::replace(&mut each[j], Ok(Value::Empty));
                }
                self.give(each);
            }
        }
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
        Operand::Value(Lane::Same(got)) if last => {
            std::mem::replace(got, Ok(Value::Empty)).map(Arg::Scalar)
        }
        Operand::Value(Lane::Same(got)) => got.clone().map(Arg::Scalar),
        Operand::Value(Lane::Each(each)) => {
            std::mem::replace(&mut each[k], Ok(Value::Empty)).map(Arg::Scalar)
        }
    }
}
