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
//! written once. A value is made by [`operate`], over two numbers by
//! `binary`'s own [`arithmetic`] and [`compares`], or by the library: a
//! strict function called at each cell with the operands it would be called
//! with there, or, for the language's aggregates and `NOT`, the library's
//! own rule and fold (`library::Fold`) driven here over the run's operands,
//! a range of a sheet kept as an array walked by index (KERNEL.26). A
//! reason propagates as it does there, each cell's answer the first its own
//! evaluation meets, and `IF` computes each branch at the cells that take
//! it and at no other. An operand computed at a cell whose answer an
//! earlier operand has already decided is computed for nothing and changes
//! nothing, an evaluation having no effect but its answer. A constant, and
//! a reference every `$` holds, is computed once for the run. What is not
//! computed here is computed at each cell by [`Ctx`] itself: a defined
//! name, a form other than the language's `IF` and `CHOOSE`, a range where
//! one value is wanted, and a call with such an operand.
//!
//! The language's other functions are computed here by the library's own
//! meaning of each (KERNEL.28): `ABS`, `INT`, `SIN` and `MOD` at the cells
//! where every operand is a number, computed or read where it stands, by the
//! library's function of numbers, and at any other cell by the library's
//! strict function there, so that no coercion, nor the order in which a call
//! meets its operands' reasons and errors, is written here; `ROW` and
//! `COLUMN` from the run's own cells, or from a reference's first cell moved
//! to each; and `CHOOSE` as `IF` is, its index at every cell and each value
//! at the cells that choose it and at no other.
//!
//! A lane holds each cell's answer as a value, or, while every answer is
//! one, as a number or a truth value (KERNEL.26): what a read of numbers,
//! the arithmetic and the comparisons over them, the aggregates and `IF`
//! make and read with no value's kind to match, a lane of values from the
//! first answer of another kind.
//!
//! A part a formula writes more than once is computed once where its places
//! stand in one branch of every `IF` around them (KERNEL.27), as Life's rule
//! writes `SUM(Screen.last!A1:C3)-Screen.last!B2` twice: computed where it is
//! first written, its lane kept, every later place taking a copy and the
//! last the lane itself. The parts are found once a shape ([`Repeats`]), each
//! place by its path from the parse's root, and every value is the one each
//! place would compute, an evaluation having no effect but its answer.
//!
//! Its buffers are its caller's, kept from one frame to the next ([`Lanes`])
//! and made in the first frame a run is evaluated in, as many as its parse
//! can hold at once whatever branch its cells take ([`holds`]), with a lane
//! for each repeated part, so that a later frame makes none; a value the
//! same at every cell is copied as often as alone, and so is a repeated
//! part's. It recurses once a level of nesting as [`Ctx`] does, a level's
//! frame holding no array, since `VLA.1`'s limits are pinned on a stack of
//! 1 MiB. Bit for bit with the cell by cell evaluation, which the machine's
//! tests hold it to over generated grids of every kind of node, and with
//! `Calc`, which plans and parses on its own.
//!
//! The paths it takes are read off a parse without computing it, for a
//! plan's census (KERNEL.30, [`paths`]): each read's cost class, by a walk
//! that asks what the dispatch asks, held to it by a test that notes the
//! paths runs take ([`taken`]).

use std::cell::{Cell, OnceCell, RefCell};
use std::convert::Infallible;
use std::hash::Hasher;

use crate::refers::Shape as Part;

use super::dense::Dense;
use super::eval::{
    arithmetic, compare_numbers, compares, finite, intersection, may_merge, near_bounds, operate,
    Arg, Computed, Ctx, Env, Known, Library, Reason, Strict,
};
use super::formula::{BinOp, Expr, RefExpr};
use super::graph::{Area, CellId};
use super::library::{self, Axis, Fold, Ranged, Rule};
use super::lookups;
use super::value::{ErrorKind, Value};

/// The most cells evaluated at once: a lane of them is a few kilobytes, so
/// that the lanes a formula holds at once stay in the processor's cache.
pub(crate) const CHUNK: usize = 256;

/// What an expression answers at one cell, as [`Ctx::scalar`] answers it.
type Got = Result<Value, Reason>;

/// What an expression answers at each cell of a selection.
#[derive(Debug)]
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

/// A comparison's truth value pushed for each pair of numbers, exactly, as
/// `compares` over `eval::exact_order` answers for the finite numbers a
/// lane holds, the operator's test chosen once and each test a loop of its
/// own; and whether `marked` held of any pair (KERNEL.8).
#[inline(always)]
fn tested(
    truths: &mut Vec<bool>,
    pairs: impl Iterator<Item = (f64, f64)>,
    op: BinOp,
    marked: impl Fn(f64, f64) -> bool + Copy,
) -> bool {
    match op {
        BinOp::Eq => tested_by(truths, pairs, |x, y| x == y, marked),
        BinOp::Ne => tested_by(truths, pairs, |x, y| x != y, marked),
        BinOp::Lt => tested_by(truths, pairs, |x, y| x < y, marked),
        BinOp::Gt => tested_by(truths, pairs, |x, y| x > y, marked),
        BinOp::Le => tested_by(truths, pairs, |x, y| x <= y, marked),
        _ => tested_by(truths, pairs, |x, y| x >= y, marked),
    }
}

/// [`tested`] by one test.
#[inline(always)]
fn tested_by(
    truths: &mut Vec<bool>,
    pairs: impl Iterator<Item = (f64, f64)>,
    test: impl Fn(f64, f64) -> bool,
    marked: impl Fn(f64, f64) -> bool,
) -> bool {
    let mut any = false;
    truths.extend(pairs.map(|(x, y)| {
        any |= marked(x, y);
        test(x, y)
    }));
    any
}

/// The buffers an evaluation of many cells holds, kept by its caller from
/// one run to the next so that they are made once (KERNEL.20): of values,
/// of numbers and of truth values (KERNEL.26), each as many as the parse
/// holds lanes at once, and of positions; and the places of the parse's
/// repeated parts and the lanes kept for them (KERNEL.27).
#[derive(Debug, Default)]
pub(crate) struct Lanes {
    values: Vec<Vec<Got>>,
    numbers: Vec<Vec<f64>>,
    truths: Vec<Vec<bool>>,
    positions: Vec<Vec<u32>>,
    found: Vec<Key>,
    kept: Vec<Option<Lane>>,
    /// By a run's cell, where the cells side by side in its row from it end
    /// (KERNEL.43).
    ends: Vec<u32>,
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
        // Room for a chunk's stretches, the last run's emptied first, since
        // `reserve` counts from the length.
        self.ends.clear();
        self.ends.reserve(CHUNK);
    }

    /// How many buffers of each kind are made, values, numbers, truth values
    /// and positions, the most any holds, and the room of the stretches'
    /// (KERNEL.43).
    #[cfg(test)]
    pub(crate) fn made(&self) -> (usize, usize, usize, usize, usize, usize) {
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
            self.ends.capacity(),
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
                let known = lib.known(name);
                if known == Some(Known::Choose) && args.len() >= 2 {
                    // The index's lane, and the answers' beside it while the
                    // cells are labelled by the value each chooses and
                    // gathered by value, three lists of positions (KERNEL.28);
                    // then each value beside the answers, two of the lists
                    // and the positions its cells stand at.
                    let (mut values, mut positions) = holds(lib, &args[0]);
                    values = values.max(2);
                    positions = positions.max(3);
                    for a in &args[1..] {
                        let (v, p) = holds(lib, a);
                        values = values.max(1 + v);
                        positions = positions.max(3 + p);
                    }
                    return (values, positions);
                }
                if known == Some(Known::Index) && index_by_run(lib, args) {
                    // The row's lane and the column's, each beside those
                    // before it, then the answers (KERNEL.8).
                    let (mut values, mut positions) = (args.len() + 1, 0);
                    for (j, a) in args.iter().enumerate().skip(1) {
                        let (v, p) = holds(lib, a);
                        values = values.max(j + v);
                        positions = positions.max(p);
                    }
                    return (values, positions);
                }
                if matches!(known, Some(Known::IfError | Known::IfNa)) && args.len() == 2 {
                    // The value's lane, which holds the answers; then the
                    // other beside it, with the cells it is computed at and
                    // their positions in the run (KERNEL.8).
                    let (mut values, mut positions) = holds(lib, &args[0]);
                    values = values.max(1);
                    let (v, p) = holds(lib, &args[1]);
                    values = values.max(1 + v);
                    positions = positions.max(2 + p);
                    return (values, positions);
                }
                if known != Some(Known::If) || !(2..=3).contains(&args.len()) {
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

/// Whether a run computes `INDEX` by itself (KERNEL.8): of a reference written
/// as one, on one sheet, and a row and maybe a column; any other at each cell
/// by `Ctx`.
fn index_by_run(lib: &Library, args: &[Expr]) -> bool {
    (2..=3).contains(&args.len())
        && ref_operand(lib, &args[0]).is_some_and(|r| r.last_sheet.is_none())
}

/// Whether a reference names the same cells at every cell, a `$` holding
/// each corner it would move by (KERNEL.8): a vector a run's cells search
/// alike, whose index they share, the census's `keyed`.
pub(crate) fn held_whole(r: &RefExpr) -> bool {
    let c = &r.corners;
    r.last_sheet.is_none()
        && match r.shape {
            Part::Cell => c.row_abs1 && c.col_abs1,
            Part::Range => c.row_abs1 && c.col_abs1 && c.row_abs2 && c.col_abs2,
            Part::Column => c.col_abs1 && c.col_abs2,
            Part::Row => c.row_abs1 && c.row_abs2,
            Part::None => false,
        }
}

/// Whether a reference to one cell or a range moves with the cell's column
/// at every corner it has, no `$` before a column (KERNEL.43): so that a
/// run's cells side by side in a row read it as slices one column apart. A
/// whole row or column is read as an array walked, never by slices.
fn moves_with_column(r: &RefExpr) -> bool {
    match r.shape {
        Part::Cell => !r.corners.col_abs1,
        Part::Range => !r.corners.col_abs1 && !r.corners.col_abs2,
        _ => false,
    }
}

/// The reference an operand is made from at every cell, when it is one: a
/// reference written as one, under any unary plus, which returns its operand
/// as it is. A program's operands are told by the same rule (KERNEL.44,
/// `super::program`).
pub(crate) fn ref_operand<'e>(lib: &Library, a: &'e Expr) -> Option<&'e RefExpr> {
    match a {
        Expr::Ref(r) => Some(r),
        Expr::Pos(x) if !yields_value(lib, a) => ref_operand(lib, x),
        _ => None,
    }
}

/// Whether [`Ctx::eval`] of an expression is a value at every cell: anything
/// but a reference, a name, or a form whose answer may be one; the
/// language's `IF` and `CHOOSE` answer one of their own arguments, and are a
/// value when every one they may answer is (KERNEL.28). A program's operands
/// are told by the same rule (KERNEL.44).
pub(crate) fn yields_value(lib: &Library, e: &Expr) -> bool {
    match e {
        Expr::Ref(_) | Expr::Name { .. } => false,
        Expr::Pos(x) => yields_value(lib, x),
        Expr::Call(name, args) => match lib.known(name) {
            Some(Known::If | Known::Choose) => args.iter().skip(1).all(|a| yields_value(lib, a)),
            // The value, read as one, or the other as `Ctx` evaluates it
            // (KERNEL.8).
            Some(Known::IfError | Known::IfNa) => {
                args.get(1).map_or(true, |a| yields_value(lib, a))
            }
            _ => lib.form(name).is_none(),
        },
        _ => true,
    }
}

/// The parts written more than once in a shape's parse that a run computes
/// once (KERNEL.27). A part is a node of the parse the run evaluates over its
/// cells, not a literal, nor a reference a strict call takes as one, nor
/// anything inside a call [`Ctx`] computes at each cell; two places are the
/// same part when their parses are the same to the bit and they stand in the
/// same branch of every `IF` around them, so that the run reaches both over
/// the same cells. The run computes a part where it is first written and
/// keeps its lane, every later place taking a copy and the last the lane
/// itself, so that the copies are as many as alone; a place inside a later
/// place of a larger part is never reached and is left out, the larger part
/// taken first. Each place is held by its path from the parse's root, a
/// child's index a step, so that a copy of the parse is found the same way.
/// Found once a shape, as [`holds`] is; most parses have none, and hold
/// nothing.
#[derive(Clone, Debug, Default, PartialEq)]
pub(crate) struct Repeats {
    /// Each place in the parse's preorder: its part, whether it is the
    /// part's last place, and where its path ends in `steps`.
    places: Vec<(u16, bool, u32)>,
    /// The places' paths, one after another.
    steps: Vec<u16>,
    parts: usize,
}

impl Repeats {
    /// The repeated parts of a parse.
    pub(crate) fn of(lib: &Library, e: &Expr) -> Repeats {
        // The parse in preorder, each node with its parent, the branch it
        // stands in (0 for none, else an `IF`'s place and which branch), and
        // whether a run reaches it, round a stack, since a chain of operators
        // nests as deep as it is long.
        struct Node<'e> {
            e: &'e Expr,
            parent: u32,
            branch: u32,
            reached: bool,
            size: u32,
            hash: u64,
        }
        let mut nodes: Vec<Node> = Vec::new();
        let mut stack: Vec<(&Expr, u32, u32, bool)> = vec![(e, u32::MAX, 0, true)];
        let mut candidates = 0;
        while let Some((x, parent, branch, reached)) = stack.pop() {
            let at = nodes.len() as u32;
            candidates += usize::from(reached && !literal(x));
            nodes.push(Node {
                e: x,
                parent,
                branch,
                reached,
                size: 1,
                hash: 0,
            });
            match x {
                Expr::Neg(y) | Expr::Pos(y) | Expr::Percent(y) => {
                    stack.push((y, at, branch, reached));
                }
                // The left operand a link of a chain, which the run reaches
                // as one when it is a part.
                Expr::Binary(_, l, r) => {
                    stack.push((r, at, branch, reached));
                    stack.push((l, at, branch, reached));
                }
                Expr::Call(name, args) => {
                    let how = reaches(lib, name, args);
                    for (j, a) in args.iter().enumerate().rev() {
                        let (b, r) = match how {
                            Reaches::If if j > 0 => (at * 4 + j as u32, reached),
                            Reaches::If => (branch, reached),
                            Reaches::Operands => (branch, reached && ref_operand(lib, a).is_none()),
                            Reaches::None => (branch, false),
                        };
                        stack.push((a, at, b, r));
                    }
                }
                _ => {}
            }
        }
        if candidates < 2 {
            return Repeats::default();
        }
        // Each node's size and its parse's hash, children before parents.
        for i in (1..nodes.len()).rev() {
            let p = nodes[i].parent as usize;
            nodes[p].size += nodes[i].size;
        }
        for i in (0..nodes.len()).rev() {
            let mut h = crate::fx::FxHasher::default();
            node_hash(nodes[i].e, &mut h);
            let mut c = i + 1;
            for _ in 0..arity(nodes[i].e) {
                h.write_u64(nodes[c].hash);
                c += nodes[c].size as usize;
            }
            nodes[i].hash = h.finish();
        }
        // The places a run reaches, by hash, branch and place, each place
        // once, so that the order is the parse's whatever a sort keeps.
        let mut reached: Vec<Key> = (1..nodes.len() as u32)
            .filter(|&i| nodes[i as usize].reached && !literal(nodes[i as usize].e))
            .map(|i| (nodes[i as usize].hash as usize, nodes[i as usize].branch, i))
            .collect();
        reached.sort_unstable();
        // Among the places of one hash and branch, those the same to the
        // bit, each part's places one after another in `members`, each part
        // by its size, the largest first, and where its places begin.
        let mut members: Vec<u32> = Vec::new();
        let mut groups: Vec<Key> = Vec::new();
        let mut taken = vec![false; reached.len()];
        for a in 0..reached.len() {
            if taken[a] {
                continue;
            }
            let (hash, branch, first) = reached[a];
            let start = members.len();
            for (b, &(h, br, i)) in reached.iter().enumerate().skip(a) {
                if (h, br) != (hash, branch) {
                    break;
                }
                if !taken[b] && same_part(nodes[first as usize].e, nodes[i as usize].e) {
                    taken[b] = true;
                    members.push(i);
                }
            }
            match members.len() - start {
                0 | 1 => members.truncate(start),
                n => {
                    let size = nodes[first as usize].size as usize;
                    groups.push((usize::MAX - size, start as u32, n as u32));
                }
            }
        }
        groups.sort_unstable();
        // Each part kept while two of its places are still reached, the
        // places inside a later place of a part kept before it never being.
        let mut hidden = vec![false; nodes.len()];
        let mut kept: Vec<Key> = Vec::new();
        let mut parts = 0;
        for &(_, start, n) in &groups {
            let at = start as usize;
            let places: Vec<u32> = members[at..at + n as usize]
                .iter()
                .copied()
                .filter(|&i| !hidden[i as usize])
                .collect();
            if places.len() < 2 || parts == usize::from(u16::MAX) {
                continue;
            }
            for (k, &i) in places.iter().enumerate() {
                kept.push((i as usize, parts as u32, u32::from(k + 1 == places.len())));
                if k > 0 {
                    let end = i + nodes[i as usize].size;
                    for h in &mut hidden[i as usize + 1..end as usize] {
                        *h = true;
                    }
                }
            }
            parts += 1;
        }
        if parts == 0 {
            return Repeats::default();
        }
        kept.sort_unstable();
        // Each place's path: the child's index at each step from the root.
        let mut out = Repeats {
            parts,
            ..Repeats::default()
        };
        let mut path: Vec<u16> = Vec::new();
        for (i, part, last) in kept {
            let (i, part, last) = (i as u32, part as u16, last == 1);
            path.clear();
            let mut c = i as usize;
            while nodes[c].parent != u32::MAX {
                let p = nodes[c].parent as usize;
                // Which child of its parent: the children before it in
                // preorder, each a whole subtree.
                let (mut child, mut s) = (0u16, p + 1);
                while s != c {
                    s += nodes[s].size as usize;
                    child += 1;
                }
                path.push(child);
                c = p;
            }
            out.steps.extend(path.iter().rev());
            out.places.push((part, last, out.steps.len() as u32));
        }
        out
    }

    /// How many lanes a run keeps for them at once: one a part.
    pub(crate) fn parts(&self) -> usize {
        self.parts
    }

    /// How many places the parts are written at that a run reaches.
    #[cfg(test)]
    pub(crate) fn places(&self) -> usize {
        self.places.len()
    }

    /// Each place's node in a parse of the same shape, by its address, with
    /// its part and whether it is the part's last place (`part << 1 | last`),
    /// sorted by address.
    fn found(&self, root: &Expr, out: &mut Vec<Key>) {
        out.clear();
        let mut from = 0;
        for &(part, last, to) in &self.places {
            let path = &self.steps[from..to as usize];
            from = to as usize;
            let mut e = Some(root);
            for &k in path {
                e = match (e, usize::from(k)) {
                    (Some(Expr::Neg(x) | Expr::Pos(x) | Expr::Percent(x)), 0) => Some(&**x),
                    (Some(Expr::Binary(_, l, _)), 0) => Some(&**l),
                    (Some(Expr::Binary(_, _, r)), 1) => Some(&**r),
                    (Some(Expr::Call(_, args)), j) => args.get(j),
                    _ => None,
                };
            }
            // A path is its own parse's, found on a parse of the same shape;
            // were one not, its place would be computed where it stands.
            debug_assert!(e.is_some(), "a place's path leaves its parse");
            if let Some(e) = e {
                let at = e as *const Expr as usize;
                out.push((at, u32::from(part) << 1 | u32::from(last), 0));
            }
        }
        out.sort_unstable();
    }
}

/// What [`Repeats`] sorts, the places by hash, branch and place, the parts by
/// size, and the places' addresses, each a key the shape of a [`CellId`], so
/// that the build's one sort of cells sorts them all (KERNEL.27): each sort's
/// own kind of key cost the language's module some kilobytes. On a 32-bit
/// target a hash keeps its low half, which only puts more places side by
/// side for [`same_part`] to tell apart.
type Key = (usize, u32, u32);

/// How a run reaches a call's arguments: each branch of the language's `IF`
/// in a branch of its own, a strict call's operands where the call stands,
/// or none, the call computed at each cell or answering without them.
enum Reaches {
    If,
    Operands,
    None,
}

/// [`Reaches`] for a call, as [`Many::call`] dispatches it.
fn reaches(lib: &Library, name: &str, args: &[Expr]) -> Reaches {
    if lib.form(name).is_some() {
        match lib.known(name) {
            Some(Known::If) if (2..=3).contains(&args.len()) => Reaches::If,
            // The value at every cell, the other in a branch of its own
            // (KERNEL.8).
            Some(Known::IfError | Known::IfNa) if args.len() == 2 => Reaches::If,
            // The row and the column at every cell, the reference read
            // where it picks (KERNEL.8).
            Some(Known::Index) if index_by_run(lib, args) => Reaches::Operands,
            _ => Reaches::None,
        }
    } else if lib.strict(name).is_none() || !args.iter().all(|a| is_operand(lib, a)) {
        Reaches::None
    } else {
        Reaches::Operands
    }
}

/// Whether a node is a literal, which a run makes the same at every cell as
/// cheaply as it would copy it, and is never a repeated part (KERNEL.27).
fn literal(e: &Expr) -> bool {
    matches!(
        e,
        Expr::Number(_) | Expr::Text(_) | Expr::Bool(_) | Expr::Error(_) | Expr::Empty
    )
}

/// How many children a node has, in the order [`Repeats::of`] walks them.
fn arity(e: &Expr) -> usize {
    match e {
        Expr::Neg(_) | Expr::Pos(_) | Expr::Percent(_) => 1,
        Expr::Binary(..) => 2,
        Expr::Call(_, args) => args.len(),
        _ => 0,
    }
}

/// A node's own part of its parse's hash: its kind and what it holds but its
/// children.
fn node_hash(e: &Expr, h: &mut crate::fx::FxHasher) {
    use std::hash::Hash;
    std::mem::discriminant(e).hash(h);
    match e {
        Expr::Number(v) => h.write_u64(v.to_bits()),
        Expr::Text(t) => t.hash(h),
        Expr::Bool(b) => b.hash(h),
        Expr::Error(k) => k.hash(h),
        Expr::Ref(r) => {
            r.written.hash(h);
            r.sheet.hash(h);
            (r.top, r.left, r.bottom, r.right).hash(h);
        }
        Expr::Name { name, sheet } => {
            name.hash(h);
            sheet.hash(h);
        }
        Expr::Call(name, args) => {
            name.hash(h);
            args.len().hash(h);
        }
        Expr::Binary(op, ..) => std::mem::discriminant(op).hash(h),
        Expr::Empty | Expr::Neg(_) | Expr::Pos(_) | Expr::Percent(_) => {}
    }
}

/// Whether two parses are the same to the bit, a number's sign of zero
/// included: round a stack, as [`Repeats::of`] walks.
fn same_part(a: &Expr, b: &Expr) -> bool {
    let mut stack = vec![(a, b)];
    while let Some(pair) = stack.pop() {
        match pair {
            (Expr::Number(x), Expr::Number(y)) if x.to_bits() == y.to_bits() => {}
            (Expr::Text(x), Expr::Text(y)) if x == y => {}
            (Expr::Bool(x), Expr::Bool(y)) if x == y => {}
            (Expr::Error(x), Expr::Error(y)) if x == y => {}
            (Expr::Empty, Expr::Empty) => {}
            (Expr::Ref(x), Expr::Ref(y)) if x == y => {}
            (Expr::Name { name, sheet }, Expr::Name { name: n, sheet: s })
                if name == n && sheet == s => {}
            (Expr::Call(f, xs), Expr::Call(g, ys)) if f == g && xs.len() == ys.len() => {
                stack.extend(xs.iter().zip(ys));
            }
            (Expr::Neg(x), Expr::Neg(y))
            | (Expr::Pos(x), Expr::Pos(y))
            | (Expr::Percent(x), Expr::Percent(y)) => stack.push((x, y)),
            (Expr::Binary(o, l, r), Expr::Binary(p, m, s)) if o == p => {
                stack.push((r, s));
                stack.push((l, m));
            }
            _ => return false,
        }
    }
    true
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

/// A range of an aggregate's as [`Many::fold_stretch`] reads it (KERNEL.43):
/// its array, its rows, its first column at the stretch's first cell, its
/// width, and whether it moves with the cells or stands still.
#[derive(Clone, Copy)]
struct Slices<'d> {
    array: &'d Dense,
    top: u32,
    bottom: u32,
    left: u32,
    width: usize,
    moves: bool,
}

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

/// How `CHOOSE`'s index at every cell parts the cells of a selection
/// (KERNEL.28), as [`Parted`] has `IF`'s conditions.
enum Picked {
    /// Every cell chooses the one value.
    All(usize),
    /// The answer at every cell, the index a reason or an error.
    Answer(Lane),
    /// Each cell's own: the answers' lane, a place held for each cell that
    /// chooses a value; the cells' places in the selection by the value they
    /// choose; and where each value's places begin among them.
    Cells(Lane, Vec<u32>, Vec<u32>),
}

/// One run's evaluation.
struct Many<'a> {
    env: &'a dyn Env,
    lib: &'a Library,
    /// The run's cells, in the order of evaluation, all on one sheet.
    cells: &'a [CellId],
    /// By cell, where the cells side by side in its row from it end, the
    /// position after the last (KERNEL.43), found when a read or a fold
    /// first asks: a run none of whose references asks, as a spreadsheet's
    /// column of formulas reading its own sheet, never finds them.
    ends: OnceCell<Vec<u32>>,
    /// The buffer they are found in, until they are.
    spare: Cell<Vec<u32>>,
    /// The cell the shape was read at, from which each cell's offset is.
    anchor: CellId,
    lanes: RefCell<&'a mut Lanes>,
    /// The places of the parse's repeated parts (KERNEL.27), each by its
    /// node's address in this parse, its part and whether it is the last;
    /// sorted by address.
    found: Vec<Key>,
    /// By part, the lane its first place computed, while a later place has
    /// yet to take it.
    kept: RefCell<Vec<Option<Lane>>>,
    /// The address of the place [`Many::again`] has just asked
    /// [`Many::scalars`] to compute, 0 for none, so that the node's own
    /// frame computes it, the question of whether it is a place asked of
    /// this marker alone.
    computing: Cell<usize>,
}

/// A shape's parse evaluated at each of a run's cells, the cells read at
/// `anchor`'s offset from each, each cell's answer pushed to `out` in the
/// run's order as the shape's evaluation at the cell alone answers it
/// (`Shapes::evaluate`); `held` is what [`holds`] says of the parse, with a
/// lane for each of its repeated parts, `repeats` (KERNEL.27). Where every
/// answer is a number, the answers are `numbers` instead, `out` left empty,
/// so that no value is made for a cell whose number is the one it holds
/// (KERNEL.43).
#[allow(clippy::too_many_arguments)]
pub(crate) fn evaluate(
    env: &dyn Env,
    lib: &Library,
    expr: &Expr,
    held: (usize, usize),
    repeats: &Repeats,
    anchor: CellId,
    cells: &[CellId],
    lanes: &mut Lanes,
    out: &mut Vec<Computed>,
    numbers: &mut Vec<f64>,
) {
    lanes.reserve(held);
    let mut found = std::mem::take(&mut lanes.found);
    repeats.found(expr, &mut found);
    let mut kept = std::mem::take(&mut lanes.kept);
    kept.clear();
    kept.resize_with(repeats.parts(), || None);
    let many = Many {
        env,
        lib,
        cells,
        ends: OnceCell::new(),
        spare: Cell::new(std::mem::take(&mut lanes.ends)),
        anchor,
        lanes: RefCell::new(lanes),
        found,
        kept: RefCell::new(kept),
        computing: Cell::new(0),
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
        Lane::Numbers(lane) => {
            numbers.extend_from_slice(&lane);
            many.give_numbers(lane);
        }
        Lane::Truths(truths) => {
            out.extend(truths.iter().map(|&b| Computed::Value(Value::Bool(b))));
            many.give_truths(truths);
        }
    }
    // A lane kept for a part whose last place was not reached given back,
    // and the buffers of the places to the caller (KERNEL.27).
    for k in 0..repeats.parts() {
        let lane = many.kept.borrow_mut()[k].take();
        if let Some(lane) = lane {
            many.give_lane(lane);
        }
    }
    let Many {
        lanes,
        found,
        kept,
        ends,
        spare,
        ..
    } = many;
    let lanes = lanes.into_inner();
    lanes.found = found;
    lanes.kept = kept.into_inner();
    lanes.ends = ends.into_inner().unwrap_or_else(|| spare.into_inner());
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

    /// A comparison of two lanes of numbers at each of their cells: a lane
    /// of truth values (KERNEL.26), each as [`compare_numbers`] answers.
    /// Each pair compared exactly by the operator's own test, chosen once
    /// for the lane and not at every cell, and marked where the two stand
    /// near enough for fifteen digits to merge ([`may_merge`]), those alone
    /// compared again (KERNEL.8).
    fn compared(&self, op: BinOp, xs: &[f64], ys: &[f64]) -> Lane {
        let mut truths = self.take_truths();
        let from = truths.len();
        let pairs = || xs.iter().zip(ys).map(|(&x, &y)| (x, y));
        let near = tested(&mut truths, pairs(), op, may_merge);
        if near {
            self.compared_again(op, &mut truths[from..], |k| (xs[k], ys[k]));
        }
        Lane::Truths(truths)
    }

    /// A comparison of a lane of numbers with one number `y`, the lane's on
    /// the left, or on the right where `flipped`: a lane of truth values,
    /// each as [`compare_numbers`] answers (KERNEL.8). Each number compared
    /// exactly by the operator's own test, chosen once for the lane, as
    /// `compares` over `exact_order` answers, and marked where fifteen
    /// digits could answer otherwise: between `y` and its [`near_bounds`]
    /// on the side where the operator's exact answer would change, below
    /// `y` for `<` and `>=`, above it for `>` and `<=`, either side for `=`
    /// and `<>`, the marks sharing the test's own comparisons; the lane
    /// compared again only where one is marked. The loop the lanes ran
    /// before chose the operator's test at every cell.
    fn compared_with(&self, op: BinOp, xs: &[f64], y: f64, flipped: bool) -> Lane {
        // The operator read with the lane's number on its left.
        let lane_op = match (flipped, op) {
            (true, BinOp::Lt) => BinOp::Gt,
            (true, BinOp::Gt) => BinOp::Lt,
            (true, BinOp::Le) => BinOp::Ge,
            (true, BinOp::Ge) => BinOp::Le,
            (_, op) => op,
        };
        let (low, high) = near_bounds(y);
        let mut truths = self.take_truths();
        let from = truths.len();
        let pairs = xs.iter().map(|&x| (x, y));
        let t = &mut truths;
        let near = match lane_op {
            BinOp::Eq => tested_by(
                t,
                pairs,
                |x, y| x == y,
                move |x, y| ((x < y) & (x > low)) | ((x > y) & (x < high)),
            ),
            BinOp::Ne => tested_by(
                t,
                pairs,
                |x, y| x != y,
                move |x, y| ((x < y) & (x > low)) | ((x > y) & (x < high)),
            ),
            BinOp::Lt => tested_by(t, pairs, |x, y| x < y, move |x, y| (x < y) & (x > low)),
            BinOp::Gt => tested_by(t, pairs, |x, y| x > y, move |x, y| (x > y) & (x < high)),
            BinOp::Le => tested_by(t, pairs, |x, y| x <= y, move |x, y| (x > y) & (x < high)),
            _ => tested_by(t, pairs, |x, y| x >= y, move |x, y| (x < y) & (x > low)),
        };
        if near {
            let operands = |k: usize| if flipped { (y, xs[k]) } else { (xs[k], y) };
            self.compared_again(op, &mut truths[from..], operands);
        }
        Lane::Truths(truths)
    }

    /// The truth values of a lane of comparisons made again where the two
    /// numbers [`may_merge`], by [`compare_numbers`] itself.
    #[cold]
    #[inline(never)]
    fn compared_again(
        &self,
        op: BinOp,
        truths: &mut [bool],
        operands: impl Fn(usize) -> (f64, f64),
    ) {
        for (k, truth) in truths.iter_mut().enumerate() {
            let (x, y) = operands(k);
            if may_merge(x, y) {
                *truth = compares(op, compare_numbers(x, y));
            }
        }
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

    /// [`Ctx::scalar`] at every cell of the selection; a place of a repeated
    /// part from its part's lane (KERNEL.27).
    fn scalars(&self, e: &Expr, sel: Sel) -> Lane {
        if self.repeated(e) {
            return self.again(e, &sel);
        }
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

    /// Whether an expression is a place of a repeated part (KERNEL.27), by
    /// its node's address; never the place [`Many::again`] is computing,
    /// which it asks [`Many::scalars`] for once. A question of a `bool`, so
    /// that the frame a level of nesting stacks holds nothing more for it.
    fn repeated(&self, e: &Expr) -> bool {
        let at = e as *const Expr as usize;
        !self.found.is_empty() && self.computing.replace(0) != at && self.part_at(at).is_some()
    }

    /// A place's part, and whether it is the part's last place, by its
    /// node's address.
    fn part_at(&self, at: usize) -> Option<(u16, bool)> {
        let i = self.found.binary_search_by_key(&at, |&(a, _, _)| a).ok()?;
        let (_, place, _) = self.found[i];
        Some(((place >> 1) as u16, place & 1 == 1))
    }

    /// A place of a repeated part, in a frame of its own (KERNEL.27): the
    /// lane its part's first place computed and kept, copied, or taken at
    /// the last place; at the first, the part computed by [`Many::scalars`]
    /// and a copy kept.
    #[inline(never)]
    fn again(&self, e: &Expr, sel: &Sel) -> Lane {
        let at = e as *const Expr as usize;
        let Some((part, last)) = self.part_at(at) else {
            // Never: the node was found a place. Computed where it stands.
            self.computing.set(at);
            return self.scalars(e, *sel);
        };
        let k = usize::from(part);
        let held = self.kept.borrow_mut()[k].take();
        match held {
            Some(lane) if last => lane,
            Some(lane) => {
                let copy = self.copied(&lane);
                self.kept.borrow_mut()[k] = Some(lane);
                copy
            }
            None => {
                self.computing.set(at);
                let lane = self.scalars(e, *sel);
                if !last {
                    let copy = self.copied(&lane);
                    self.kept.borrow_mut()[k] = Some(copy);
                }
                lane
            }
        }
    }

    /// A lane copied into a buffer of its own kind (KERNEL.27).
    fn copied(&self, lane: &Lane) -> Lane {
        match lane {
            Lane::Same(got) => Lane::Same(got.clone()),
            Lane::Each(each) => {
                let mut copy = self.take();
                copy.extend(each.iter().cloned());
                Lane::Each(copy)
            }
            Lane::Numbers(numbers) => {
                let mut copy = self.take_numbers();
                copy.extend_from_slice(numbers);
                Lane::Numbers(copy)
            }
            Lane::Truths(truths) => {
                let mut copy = self.take_truths();
                copy.extend_from_slice(truths);
                Lane::Truths(copy)
            }
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
        #[cfg(test)]
        taken::note(if r.corners.row_abs1 && r.corners.col_abs1 {
            taken::Path::Once
        } else if dense.is_some() && !r.corners.col_abs1 {
            taken::Path::Stride
        } else if dense.is_some() {
            taken::Path::Slot
        } else {
            taken::Path::Lookup
        });
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
        // A reference that holds its column, read by cells of one row, reads
        // one cell at every one of them (KERNEL.33): Sand's gate,
        // `Gate!$B2`, before a branch most cells skip. A run's cells stand in
        // the order by row and then by column, and a selection keeps their
        // order, so its first and last cells share a row only when every one
        // does.
        if r.corners.col_abs1 && self.cells[sel.at(0)].1 == self.cells[sel.at(sel.len() - 1)].1 {
            return Lane::Same(got(0));
        }
        // The cells side by side in one row read together (KERNEL.43): a
        // reference that moves with the column, over a sheet kept as an
        // array, as one slice of the array's numbers from the first cell's
        // to the last's or the array's edge, the cells past the edge alone;
        // one whose column a `$` holds, as one read, the same cell at each.
        // The cells one below the other in one column, a reference no `$`
        // holds, over a sheet kept as an array, as one column of its numbers
        // (KERNEL.31), or one by one where a slot holds no number. Any other
        // cell, and the cells of a slice not wholly numbers, one by one. A
        // cell outside the array is empty, which a lane of numbers never
        // holds, so a slice starts at the first cell's or not at all.
        let mut numbers = self.take_numbers();
        let n = sel.len();
        let mut k = 0;
        let mut odd: Option<(usize, Got)> = None;
        let down = !r.corners.row_abs1 && !r.corners.col_abs1;
        'cells: while k < n {
            let m = if dense.is_some() || r.corners.col_abs1 {
                self.stretch(sel, k)
            } else {
                1
            };
            if m == 1 && down {
                if let Some(d) = dense {
                    let below = self.stretch_down(sel, k);
                    if below > 1 {
                        let (d_row, d_col) = self.offset(sel.at(k));
                        let column = r.moved(d_row, d_col).is_some_and(|(row, col, _, _)| {
                            d.numbers_in_column(row, col, below, &mut numbers)
                        });
                        if column {
                            #[cfg(test)]
                            sliced::note(sliced::DOWN);
                            k += below;
                            continue;
                        }
                        // Read one by one, the column's slots not all numbers.
                        for j in k..k + below {
                            match at(j) {
                                Ok(x) => numbers.push(x),
                                Err(other) => {
                                    odd = Some((j, other));
                                    break 'cells;
                                }
                            }
                        }
                        k += below;
                        continue;
                    }
                }
            }
            // How many of the stretch's cells its slice read.
            let mut read = 0;
            if m > 1 {
                if r.corners.col_abs1 {
                    if let Ok(x) = at(k) {
                        #[cfg(test)]
                        sliced::note(sliced::HELD);
                        numbers.extend(std::iter::repeat(x).take(m));
                        k += m;
                        continue;
                    }
                } else if let Some(d) = dense {
                    let (d_row, d_col) = self.offset(sel.at(k));
                    let slice = r.moved(d_row, d_col).and_then(|(row, col, _, _)| {
                        match d.inside((row, row), col, 1, m, true) {
                            (0, b @ 2..) => d.numbers_in_row(row, col, b),
                            _ => None,
                        }
                    });
                    if let Some(slice) = slice {
                        #[cfg(test)]
                        sliced::note(sliced::READ);
                        numbers.extend_from_slice(slice);
                        read = slice.len();
                    }
                }
            }
            for j in k + read..k + m {
                match at(j) {
                    Ok(x) => numbers.push(x),
                    Err(other) => {
                        odd = Some((j, other));
                        break 'cells;
                    }
                }
            }
            k += m;
        }
        // At the first cell whose value is not a number, a lane of values,
        // the rest read one by one.
        if let Some((j, other)) = odd {
            let mut out = self.take();
            out.extend(numbers.iter().map(|&x| Ok(Value::Number(x))));
            self.give_numbers(numbers);
            out.push(other);
            out.extend((j + 1..n).map(got));
            return Lane::Each(out);
        }
        Lane::Numbers(numbers)
    }

    /// How many of the selection's cells from its `k`th stand side by side
    /// in one row, each the cell right of the one before (KERNEL.43): a run's
    /// cells of one row, or of a row's stretch a branch or a frame's reach
    /// selects.
    #[inline]
    fn stretch(&self, sel: Sel, k: usize) -> usize {
        let ends = self.ends();
        match sel {
            Sel::All(n) => (ends[k] as usize).min(n) - k,
            Sel::Some(at) => {
                let p = at[k] as usize;
                let end = ends[p] as usize;
                let mut m = 1;
                while k + m < at.len() && at[k + m] as usize == p + m && p + m < end {
                    m += 1;
                }
                m
            }
        }
    }

    /// How many of the selection's cells from its `k`th stand one below the
    /// other in one column, each the cell under the one before (KERNEL.31):
    /// a run's cells of a column, as a spreadsheet's column of formulas
    /// stands, or the stretch of one a branch or a frame's reach selects.
    /// Counted as it is asked, a column's stretch being read at its first
    /// cell alone.
    #[inline]
    fn stretch_down(&self, sel: Sel, k: usize) -> usize {
        let (s, row, col) = self.cells[sel.at(k)];
        let mut m = 1;
        while k + m < sel.len() && self.cells[sel.at(k + m)] == (s, row + m as u32, col) {
            m += 1;
        }
        m
    }

    /// Each cell's stretch, found the first time one is asked for, from the
    /// last cell back (KERNEL.43).
    fn ends(&self) -> &[u32] {
        self.ends.get_or_init(|| {
            let cells = self.cells;
            let mut ends = self.spare.take();
            ends.clear();
            ends.resize(cells.len(), 0);
            for p in (0..cells.len()).rev() {
                let (s, row, col) = cells[p];
                let beside = cells.get(p + 1) == Some(&(s, row, col + 1));
                ends[p] = if beside { ends[p + 1] } else { (p + 1) as u32 };
            }
            ends
        })
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
    /// spine, the right operands on a stack; a link below the first that is
    /// a place of a repeated part the leftmost operand (KERNEL.27), which
    /// computes the same values in the same order.
    fn chain(&self, e: &Expr, sel: Sel) -> Lane {
        const NEAR: usize = 16;
        let mut near: [(BinOp, &Expr); NEAR] = [(BinOp::Add, e); NEAR];
        let mut far: Vec<(BinOp, &Expr)> = Vec::new();
        let mut links = 0;
        let mut left = e;
        while let Expr::Binary(op, l, r) = left {
            if links > 0 && self.repeated(left) {
                break;
            }
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

    /// [`operate`] at every cell: the left operand's reason first, then the
    /// right's, then the operator over the two values, or its label; over
    /// two lanes of numbers, or one and a number the same at every cell, the
    /// operator's rule over numbers, `binary`'s own (KERNEL.26).
    fn combine(&self, op: BinOp, a: Lane, b: Lane) -> Lane {
        let (a, b) = match (op, a, b) {
            (BinOp::Concat, a, b) => (self.values(a), self.values(b)),
            (op, Lane::Numbers(xs), Lane::Numbers(ys)) => {
                let lane = if is_arithmetic(op) {
                    self.arithmetic_in_place(xs, |k, x| arithmetic(op, x, ys[k]))
                } else {
                    let lane = self.compared(op, &xs, &ys);
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
                    let lane = self.compared_with(op, &xs, y, false);
                    self.give_numbers(xs);
                    lane
                };
            }
            (op, Lane::Same(Ok(Value::Number(x))), Lane::Numbers(ys)) => {
                return if is_arithmetic(op) {
                    self.arithmetic_in_place(ys, |_, y| arithmetic(op, x, y))
                } else {
                    let lane = self.compared_with(op, &ys, x, true);
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
                (Ok(x), Ok(y)) => operate(op, &x, y),
            }),
            (Values::Each(mut xs), Values::Same(y)) => {
                for x in xs.iter_mut() {
                    if let Ok(a) = x {
                        *x = match &y {
                            Ok(b) => operate(op, a, b),
                            Err(r) => Err(r.clone()),
                        };
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
                                *y = operate(op, a, b);
                            }
                        }
                    }
                }
                Lane::Each(ys)
            }
            (Values::Each(mut xs), Values::Each(mut ys)) => {
                for (x, y) in xs.iter_mut().zip(ys.iter_mut()) {
                    if let Ok(a) = x {
                        *x = match y {
                            Ok(b) => operate(op, a, b),
                            Err(r) => Err(std::mem::replace(r, Reason::Cycle)),
                        };
                    }
                }
                self.give(ys);
                Lane::Each(xs)
            }
        }
    }

    /// A call at every cell: the language's `IF` and `CHOOSE` here, another
    /// form at each cell, one of the language's functions here by the
    /// library's own meaning, any other strict function over its operands, a
    /// function the library does not hold named, in the order `Ctx` looks
    /// them up.
    fn call(&self, e: &Expr, name: &str, args: &[Expr], sel: Sel) -> Lane {
        if self.lib.form(name).is_some() {
            return match self.lib.known(name) {
                Some(Known::If) => self.if_(args, sel),
                Some(Known::Choose) => self.choose(args, sel),
                Some(Known::Index) if index_by_run(self.lib, args) => self.index(e, args, sel),
                Some(Known::IfError) => self.if_error(args, sel, false),
                Some(Known::IfNa) => self.if_error(args, sel, true),
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
    /// computed, by the run itself: an aggregate or `NOT` when its operands
    /// are ones it walks (KERNEL.26), a function of numbers or of a cell's
    /// place as the paths below take it (KERNEL.28); `None` for a call left to
    /// the library's function `f` at each cell.
    fn known_call(&self, f: Strict, name: &str, ops: &mut [Operand<'_>], sel: Sel) -> Option<Lane> {
        match self.lib.known(name)? {
            // Each function of numbers its own instance of the path, so that
            // the library's function is compiled into it (as `aggregate` is
            // compiled for each fold).
            Known::Abs => self.of_numbers(f, 1, ops, sel, |a, _| library::abs_of(a)),
            Known::Int => self.of_numbers(f, 1, ops, sel, |a, _| library::int_of(a)),
            Known::Sin => self.of_numbers(f, 1, ops, sel, |a, _| library::sin_of(a)),
            Known::Mod => self.of_numbers(f, 2, ops, sel, library::mod_of),
            Known::Row => self.placed(Axis::Row, ops, sel),
            Known::Column => self.placed(Axis::Column, ops, sel),
            known => self.known(known, ops, sel),
        }
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
            return Some(self.negated(lane));
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
        #[cfg(test)]
        if matches!(
            known,
            Known::Sum | Known::Min | Known::Max | Known::And | Known::Or
        ) {
            for (x, k) in ops.iter().zip(kept) {
                if let (Operand::Ref(r), Some(_)) = (x, k) {
                    taken::note(if moves_with_column(r) {
                        taken::Path::Stride
                    } else {
                        taken::Path::Array
                    });
                }
            }
        }
        // The answers a lane of numbers, or of truth values, while each is.
        let numbers = || Lane::Numbers(self.take_numbers());
        let truths = || Lane::Truths(self.take_truths());
        Some(match known {
            Known::Sum => self.aggregate::<library::Sum>(ops, kept, sel, numbers()),
            Known::Min => self.aggregate::<library::Least>(ops, kept, sel, numbers()),
            Known::Max => self.aggregate::<library::Most>(ops, kept, sel, numbers()),
            Known::And => self.aggregate::<library::All>(ops, kept, sel, truths()),
            Known::Or => self.aggregate::<library::Some_>(ops, kept, sel, truths()),
            // The forms, `NOT` above, and the functions `known_call` takes
            // to their own paths (KERNEL.28); the lookups, the criteria and
            // the error tests, called at each cell (KERNEL.8).
            Known::If
            | Known::Choose
            | Known::Not
            | Known::Abs
            | Known::Int
            | Known::Sin
            | Known::Mod
            | Known::Row
            | Known::Column
            | Known::Index
            | Known::IfError
            | Known::IfNa
            | Known::IsError
            | Known::IsNa
            | Known::Match
            | Known::XMatch
            | Known::VLookup
            | Known::HLookup
            | Known::Lookup
            | Known::CountIf
            | Known::SumIf => return None,
        })
    }

    /// `ROW` and `COLUMN` at every cell (KERNEL.28), as `library::placed`
    /// answers at one: of no operand, the cell itself, which the run knows
    /// as the cell `Env::enter` would tell the environment, a lane of numbers
    /// made with no call; of a reference written as one whose sheets the
    /// shape found as it was placed, or that names none, its first cell moved
    /// to the cell, `#VALUE!` where the move leaves the sheet; `None`, the
    /// library's function at each cell, for any other operand or count, and
    /// for a reference whose sheet is found by name at each evaluation.
    #[inline(never)]
    fn placed(&self, axis: Axis, ops: &mut [Operand<'_>], sel: Sel) -> Option<Lane> {
        match ops {
            [] => {
                let mut numbers = self.take_numbers();
                numbers.extend((0..sel.len()).map(|k| {
                    let (_, row, column) = self.cells[sel.at(k)];
                    axis.of(row, column)
                }));
                Some(Lane::Numbers(numbers))
            }
            [Operand::Ref(r)]
                if r.found.is_some() || (r.sheet.is_none() && r.last_sheet.is_none()) =>
            {
                let mut out = Lane::Numbers(self.take_numbers());
                for k in 0..sel.len() {
                    let (d_row, d_col) = self.offset(sel.at(k));
                    let cell = r.moved(d_row, d_col).map(|(top, left, _, _)| (top, left));
                    self.push(&mut out, Ok(library::placed(axis, cell)));
                }
                Some(out)
            }
            _ => None,
        }
    }

    /// `ABS`, `INT`, `SIN` and `MOD` at every cell (KERNEL.28): at a cell
    /// where each of its `count` operands is a number, computed, the same at
    /// every cell, or read where it stands from a reference written as one
    /// cell, the library's function of numbers `g`, a function of one number
    /// reading no second; at any other cell the library's function `f` there,
    /// its operands taken as alone, so that the coercions and the order in
    /// which a call meets its operands' reasons and errors stay the library's
    /// and `Ctx`'s. `None`, the library's function at each cell, for another
    /// count of operands, and for a reference to more than one cell or
    /// across sheets. The references are read, and the cells computed, each
    /// in a frame of its own, so that neither stands beneath the other.
    #[inline(never)]
    fn of_numbers<G: Fn(f64, f64) -> Value>(
        &self,
        f: Strict,
        count: usize,
        ops: &mut [Operand<'_>],
        sel: Sel,
        g: G,
    ) -> Option<Lane> {
        if ops.len() != count {
            return None;
        }
        let one_cell = |r: &RefExpr| matches!(r.shape, Part::Cell) && r.last_sheet.is_none();
        if ops
            .iter()
            .any(|x| matches!(x, Operand::Ref(r) if !one_cell(r)))
        {
            return None;
        }
        let read = self.read_refs(ops, sel);
        Some(self.numbers_by(f, count, ops, read, sel, g))
    }

    /// Each reference written as one among a call's operands read where it
    /// stands, as the run reads one (KERNEL.28).
    #[inline(never)]
    fn read_refs(&self, ops: &[Operand<'_>], sel: Sel) -> [Option<Lane>; 2] {
        let mut read: [Option<Lane>; 2] = [None, None];
        for (j, x) in ops.iter().enumerate().take(2) {
            if let Operand::Ref(r) = x {
                read[j] = Some(self.reads(r, sel));
            }
        }
        read
    }

    /// [`Many::of_numbers`] at each cell, the references read.
    #[inline(never)]
    fn numbers_by<G: Fn(f64, f64) -> Value>(
        &self,
        f: Strict,
        count: usize,
        ops: &mut [Operand<'_>],
        read: [Option<Lane>; 2],
        sel: Sel,
        g: G,
    ) -> Lane {
        let n = sel.len();
        let mut out = Lane::Numbers(self.take_numbers());
        for k in 0..n {
            let a = number_at(operand_lane(&read[0], &ops[0]), k);
            let b = match count {
                2 => number_at(operand_lane(&read[1], &ops[1]), k),
                _ => Some(0.0),
            };
            let got = match (a, b) {
                (Some(a), Some(b)) => Ok(g(a, b)),
                _ => self.alone(f, ops, k, k + 1 == n, sel),
            };
            self.push(&mut out, got);
        }
        for lane in read.into_iter().flatten() {
            self.give_lane(lane);
        }
        out
    }

    /// A strict function at the selection's `k`th cell alone, its operands
    /// taken there as the run's other calls take them (KERNEL.28).
    fn alone(&self, f: Strict, ops: &mut [Operand<'_>], k: usize, last: bool, sel: Sel) -> Got {
        let ctx = self.ctx(sel.at(k));
        match ops {
            [x] => match take_arg(x, k, last, &ctx) {
                Ok(a) => f(&ctx, std::slice::from_ref(&a)),
                Err(r) => Err(r),
            },
            [x, y] => match take_arg(x, k, last, &ctx) {
                Err(r) => Err(r),
                Ok(a) => match take_arg(y, k, last, &ctx) {
                    Err(r) => Err(r),
                    Ok(b) => f(&ctx, &[a, b]),
                },
            },
            _ => {
                let mut args = Vec::with_capacity(ops.len());
                for x in ops.iter_mut() {
                    args.push(take_arg(x, k, last, &ctx)?);
                }
                f(&ctx, &args)
            }
        }
    }

    /// `CHOOSE` at every cell (KERNEL.28), as `library::choose` computes it
    /// at one: the index at each, then each value at the cells that choose
    /// it and at no other, its answers placed among the others as `IF`'s
    /// branches' are.
    fn choose(&self, args: &[Expr], sel: Sel) -> Lane {
        if args.len() < 2 {
            return Lane::Same(Ok(Value::Error(ErrorKind::Value)));
        }
        let index = self.scalars(&args[0], sel);
        let (mut out, order, starts) = match self.picked(index, args.len() - 1, sel) {
            Picked::All(j) => return self.scalars(&args[j], sel),
            Picked::Answer(lane) => return lane,
            Picked::Cells(out, order, starts) => (out, order, starts),
        };
        for j in 1..args.len() {
            let (from, to) = (starts[j] as usize, starts[j + 1] as usize);
            if from < to {
                self.fill_value(&mut out, &args[j], sel, &order[from..to]);
            }
        }
        self.give_positions(order);
        self.give_positions(starts);
        out
    }

    /// How `CHOOSE`'s index at every cell of the selection parts its cells
    /// (KERNEL.28), in a frame of its own, as [`Many::parted`] is: each cell
    /// labelled by the value it chooses, `library::chosen`, or answered by its
    /// index's reason or error; then the cells by value, a count of the
    /// labels, `order` the places of each value's cells, value `j`'s from
    /// `starts[j]` to `starts[j + 1]`, so that every value's cells are found
    /// in one pass over them whatever the count of values.
    #[inline(never)]
    fn picked(&self, index: Lane, count: usize, sel: Sel) -> Picked {
        if let Lane::Same(got) = &index {
            return match got {
                Err(r) => Picked::Answer(Lane::Same(Err(r.clone()))),
                Ok(v) => match library::chosen(v, count) {
                    Ok(j) => Picked::All(j),
                    Err(e) => Picked::Answer(Lane::Same(Ok(Value::Error(e)))),
                },
            };
        }
        let n = sel.len();
        let mut labels = self.positions();
        let mut out = Lane::Numbers(self.take_numbers());
        for k in 0..n {
            let chose = index.with(k, |g| match g {
                Err(r) => Err(Err(r.clone())),
                Ok(v) => library::chosen(v, count).map_err(|e| Ok(Value::Error(e))),
            });
            match chose {
                // A place held for the value's answer.
                Ok(j) => {
                    labels.push(j as u32);
                    self.push(&mut out, Ok(Value::Number(0.0)));
                }
                Err(answer) => {
                    labels.push(0);
                    self.push(&mut out, answer);
                }
            }
        }
        self.give_lane(index);
        // A count of the labels, each a value's last place and one more, then
        // the places from the last back, each value's taking its places in
        // the selection's order and its count down to its first.
        let mut starts = self.positions();
        starts.resize(count + 2, 0);
        for &j in labels.iter() {
            starts[j as usize] += 1;
        }
        for j in 1..starts.len() {
            starts[j] += starts[j - 1];
        }
        let mut order = self.positions();
        order.resize(n, 0);
        for k in (0..n).rev() {
            let j = labels[k] as usize;
            starts[j] -= 1;
            order[starts[j] as usize] = k as u32;
        }
        self.give_positions(labels);
        Picked::Cells(out, order, starts)
    }

    /// One of `CHOOSE`'s values at the cells that choose it (KERNEL.28), `ks`
    /// their places in the selection, its answers placed in `out`.
    fn fill_value(&self, out: &mut Lane, value: &Expr, sel: Sel, ks: &[u32]) {
        let lane = if ks.len() == sel.len() {
            self.scalars(value, sel)
        } else {
            let mut at = self.positions();
            at.extend(ks.iter().map(|&k| sel.at(k as usize) as u32));
            let lane = self.scalars(value, Sel::Some(&at));
            self.give_positions(at);
            lane
        };
        self.place(out, lane, sel, ks);
    }

    /// `INDEX` at every cell (KERNEL.8), as `lookups::index` computes it at
    /// one, of a reference written as one: its row and its column computed
    /// at every cell, the reference moved to each, and the cell
    /// `library::picked` names read there as a read of it answers; a reason
    /// or an error of the row or the column the cell's answer; a pick of
    /// many cells where one value is wanted the label `Ctx` gives it, the
    /// reference as written, moved; a reference moved off its sheet, or to a
    /// sheet the grid does not hold, a value given there, `#REF!` past the
    /// pick's own errors, as an array of one.
    fn index(&self, e: &Expr, args: &[Expr], sel: Sel) -> Lane {
        let Some(r) = ref_operand(self.lib, &args[0]) else {
            // Never: `index_by_run` holds a reference written as one.
            return self.each_cell(e, sel);
        };
        if sel.len() == 0 {
            return Lane::Each(self.take());
        }
        let sheet = match (r.found, &r.sheet) {
            (Some((first, _)), _) => Some(usize::from(first)),
            (None, None) => Some(self.cells[sel.at(0)].0),
            (None, Some(name)) => self.env.model().find_sheet(name),
        };
        #[cfg(test)]
        if let Some(s) = sheet {
            taken::note(if self.env.dense(s).is_some() {
                taken::Path::Slot
            } else {
                taken::Path::Lookup
            });
        }
        let rows = self.scalars(&args[1], sel);
        let cols = args.get(2).map(|a| self.scalars(a, sel));
        let mut out = Lane::Numbers(self.take_numbers());
        for k in 0..sel.len() {
            let got = self.picked_at(r, sheet, &rows, cols.as_ref(), k, sel);
            self.push(&mut out, got);
        }
        self.give_lane(rows);
        if let Some(c) = cols {
            self.give_lane(c);
        }
        out
    }

    /// `INDEX`'s answer at the selection's `k`th cell (KERNEL.8).
    fn picked_at(
        &self,
        r: &RefExpr,
        sheet: Option<usize>,
        rows: &Lane,
        cols: Option<&Lane>,
        k: usize,
        sel: Sel,
    ) -> Got {
        let number = |lane: &Lane| -> Result<Result<f64, ErrorKind>, Reason> {
            lane.with(k, |g| match g {
                Ok(v) => Ok(v.number()),
                Err(r) => Err(r.clone()),
            })
        };
        let row = match number(rows)? {
            Ok(n) => n,
            Err(e) => return Ok(Value::Error(e)),
        };
        let col = match cols {
            None => None,
            Some(lane) => match number(lane)? {
                Ok(n) => Some(n),
                Err(e) => return Ok(Value::Error(e)),
            },
        };
        let (d_row, d_col) = self.offset(sel.at(k));
        let Some((s, (top, left, bottom, right))) =
            sheet.and_then(|s| r.moved(d_row, d_col).map(|m| (s, m)))
        else {
            return Ok(Value::Error(match library::picked(1, 1, row, col) {
                Ok(_) => ErrorKind::Ref,
                Err(e) => e,
            }));
        };
        match library::picked(bottom - top + 1, right - left + 1, row, col) {
            Err(e) => Ok(Value::Error(e)),
            Ok((t, l, b, rr)) if t == b && l == rr => self.env.read((s, top + t, left + l)),
            // A pick of many cells, met at the cell in a legacy formula and
            // labelled in a dynamic one, as `lookups::index_value` answers
            // (KERNEL.8, slice b), the cell entered to be asked its kind.
            Ok((t, l, b, rr)) => {
                let cell = self.cells[sel.at(k)];
                self.env.enter(cell);
                if self.env.dynamic() {
                    return Err(Reason::Construct(if (d_row, d_col) == (0, 0) {
                        r.written.clone()
                    } else {
                        crate::refers::shift_a1_references(&r.written, d_row, d_col)
                    }));
                }
                let picked = Area {
                    sheet: s,
                    top: top + t,
                    left: left + l,
                    bottom: top + b,
                    right: left + rr,
                };
                match intersection(&picked, cell) {
                    Some(id) => self.env.read(id),
                    None => Ok(Value::Error(ErrorKind::Value)),
                }
            }
        }
    }

    /// `IFERROR`, or `IFNA` where `na`, at every cell (KERNEL.8), as
    /// `lookups::iferror` computes it at one: the value at each, the answer
    /// but at the cells whose value is an error caught, where the other is
    /// computed and at no other, its answers placed as `CHOOSE`'s values'.
    fn if_error(&self, args: &[Expr], sel: Sel, na: bool) -> Lane {
        if args.len() != 2 {
            return Lane::Same(Ok(Value::Error(ErrorKind::Value)));
        }
        let mut out = self.scalars(&args[0], sel);
        let caught = |g: Result<&Value, &Reason>| matches!(g, Ok(Value::Error(k)) if lookups::caught(*k, na));
        match &out {
            Lane::Same(got) => {
                return if caught(got.as_ref()) {
                    self.scalars(&args[1], sel)
                } else {
                    out
                };
            }
            // No error stands in a lane of numbers or of truth values.
            Lane::Numbers(_) | Lane::Truths(_) => return out,
            Lane::Each(_) => {}
        }
        let mut ks = self.positions();
        if let Lane::Each(each) = &out {
            ks.extend(
                each.iter()
                    .enumerate()
                    .filter(|(_, g)| caught(g.as_ref()))
                    .map(|(k, _)| k as u32),
            );
        }
        if !ks.is_empty() {
            self.fill_value(&mut out, &args[1], sel, &ks);
        }
        self.give_positions(ks);
        out
    }

    /// `NOT` at every cell, over the value its operand is there, as
    /// `library::negated` answers it: in place over a lane of values or of
    /// truth values, a lane of truth values made over one of numbers, each
    /// number's truth turned (KERNEL.43).
    fn negated(&self, lane: Lane) -> Lane {
        match lane {
            Lane::Same(got) => Lane::Same(got.map(|v| library::negated(&v))),
            Lane::Each(mut each) => {
                for v in each.iter_mut().flatten() {
                    *v = library::negated(v);
                }
                Lane::Each(each)
            }
            Lane::Truths(mut truths) => {
                for t in truths.iter_mut() {
                    *t = !*t;
                }
                Lane::Truths(truths)
            }
            Lane::Numbers(numbers) => {
                let mut truths = self.take_truths();
                truths.extend(numbers.iter().map(|&n| n == 0.0));
                self.give_numbers(numbers);
                Lane::Truths(truths)
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
        // The cells side by side in one row folded together where every
        // operand allows it (KERNEL.43), the cells whose ranges leave their
        // arrays at the row's ends alone, and any other cell alone; the
        // folds made once, for every stretch.
        let mut folds: [F; CHUNK] = std::array::from_fn(|_| F::default());
        let n = sel.len();
        let mut k = 0;
        while k < n {
            let m = self.stretch(sel, k).min(CHUNK);
            let (a, b) = match m {
                1 => (1, 1),
                _ => self.fold_stretch::<F>(ops, kept, sel, k, m, &mut folds),
            };
            for j in k..k + a {
                let got = self.fold_at::<F>(ops, kept, sel.at(j), j);
                self.push(&mut out, got);
            }
            for fold in folds[..b - a].iter_mut() {
                let answer = std::mem::take(fold).answer();
                match (&mut out, answer) {
                    (Lane::Numbers(ns), Value::Number(x)) => ns.push(x),
                    (Lane::Truths(ts), Value::Bool(b)) => ts.push(b),
                    (_, v) => self.push(&mut out, Ok(v)),
                }
            }
            for j in k + b..k + m {
                let got = self.fold_at::<F>(ops, kept, sel.at(j), j);
                self.push(&mut out, got);
            }
            k += m;
        }
        out
    }

    /// An aggregate at the `m` cells from the selection's `k`th, which stand
    /// side by side in one row (KERNEL.43): the places `a..b` of the cells
    /// whose ranges each stay inside its array and hold numbers there, their
    /// folds the first `b - a` of `folds`, each operand's items added a lane
    /// at a time, in the order [`Many::fold_at`] adds them at each cell, a
    /// value given by its lane and a range by its rows' slices of numbers,
    /// one column apart from cell to cell where its columns move with the
    /// cell and the same slices where a `$` holds them at both corners; so
    /// that each cell's answer is its own fold's, the same additions in the
    /// same order. `(m, m)`, nothing folded, where an operand is not one it
    /// adds so: a lane of values, a range not wholly numbers inside its
    /// array, a whole row or column, a range held at one corner and not the
    /// other.
    #[inline(never)]
    fn fold_stretch<F: Fold>(
        &self,
        ops: &[Operand<'_>],
        kept: &[Option<(usize, &Dense)>],
        sel: Sel,
        k: usize,
        m: usize,
        folds: &mut [F; CHUNK],
    ) -> (usize, usize) {
        let none = (m, m);
        let mut areas: [Option<Slices>; COUNTED] = [None; COUNTED];
        let (d_row, d_col) = self.offset(sel.at(k));
        // The places every range keeps inside its array.
        let (mut a, mut b) = (0, m);
        for (j, (x, kept)) in ops.iter().zip(kept).enumerate() {
            match x {
                Operand::Value(Lane::Numbers(_) | Lane::Truths(_)) => {}
                Operand::Value(Lane::Same(Ok(Value::Number(_) | Value::Bool(_)))) => {}
                Operand::Value(_) => return none,
                Operand::Ref(r) => {
                    let Some((_, array)) = kept else {
                        return none;
                    };
                    let moves = moves_with_column(r);
                    let held = match r.shape {
                        Part::Cell => r.corners.col_abs1,
                        Part::Range => r.corners.col_abs1 && r.corners.col_abs2,
                        _ => false,
                    };
                    if !moves && !held {
                        return none;
                    }
                    let Some((top, left, bottom, right)) = r.moved(d_row, d_col) else {
                        return none;
                    };
                    let width = (right - left + 1) as usize;
                    let (x, y) = array.inside((top, bottom), left, width, m, moves);
                    (a, b) = (a.max(x), b.min(y));
                    if a + 1 >= b {
                        return none;
                    }
                    areas[j] = Some(Slices {
                        array,
                        top,
                        bottom,
                        left,
                        width,
                        moves,
                    });
                }
            }
        }
        // Every range's rows hold numbers over the places.
        let span = |at: &Slices| {
            if at.moves {
                at.width + (b - a) - 1
            } else {
                at.width
            }
        };
        let from = |at: &Slices| {
            if at.moves {
                at.left + a as u32
            } else {
                at.left
            }
        };
        for at in areas.iter().flatten() {
            for row in at.top..=at.bottom {
                if at.array.numbers_in_row(row, from(at), span(at)).is_none() {
                    return none;
                }
            }
        }
        // A number or a truth value given, and a number read out of a range,
        // is an item under either rule, never an error nor skipped: the
        // rule's own items of them (`Rule::given_number` and its kin).
        let folds = &mut folds[..b - a];
        for fold in folds.iter_mut() {
            *fold = F::default();
        }
        let (first, last) = (k + a, k + b);
        for (x, area) in ops.iter().zip(&areas) {
            match (x, area) {
                (Operand::Value(Lane::Numbers(ns)), _) => {
                    for (fold, &x) in folds.iter_mut().zip(&ns[first..last]) {
                        fold.add(F::Rule::given_number(x));
                    }
                }
                (Operand::Value(Lane::Truths(ts)), _) => {
                    for (fold, &t) in folds.iter_mut().zip(&ts[first..last]) {
                        fold.add(F::Rule::given_truth(t));
                    }
                }
                (Operand::Value(Lane::Same(Ok(v))), _) => {
                    let item = match v {
                        Value::Number(x) => F::Rule::given_number(*x),
                        Value::Bool(t) => F::Rule::given_truth(*t),
                        // Never: each operand was read above.
                        _ => return none,
                    };
                    for fold in folds.iter_mut() {
                        fold.add(item);
                    }
                }
                (Operand::Ref(_), Some(at)) => {
                    for row in at.top..=at.bottom {
                        let Some(slice) = at.array.numbers_in_row(row, from(at), span(at)) else {
                            debug_assert!(false, "a range's rows were read before");
                            return none;
                        };
                        for c in 0..at.width {
                            if at.moves {
                                for (fold, &x) in folds.iter_mut().zip(&slice[c..]) {
                                    fold.add(F::Rule::ranged_number(x));
                                }
                            } else {
                                let item = F::Rule::ranged_number(slice[c]);
                                for fold in folds.iter_mut() {
                                    fold.add(item);
                                }
                            }
                        }
                    }
                }
                // Never: each operand was read above.
                _ => return none,
            }
        }
        #[cfg(test)]
        sliced::note(sliced::FOLDED);
        (a, b)
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
        if let Some(lane) = self.known_call(f, name, std::slice::from_mut(x), sel) {
            return lane;
        }
        #[cfg(test)]
        self.noted_call(name, &[&*x]);
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
        let known = self.known_call(f, name, &mut pair, sel);
        let [a, b] = pair;
        (*x, *y) = (a, b);
        if let Some(lane) = known {
            return lane;
        }
        #[cfg(test)]
        self.noted_call(name, &[&*x, &*y]);
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
        if let Some(lane) = self.known_call(f, name, ops, sel) {
            return lane;
        }
        #[cfg(test)]
        self.noted_call(name, &ops.iter().collect::<Vec<_>>());
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
        if let [_, a, b] = args {
            if let Some(lane) = self.chosen(&conditions, a, b) {
                #[cfg(test)]
                sliced::note(sliced::CHOSEN);
                self.give_lane(conditions);
                return lane;
            }
        }
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

    /// `IF` between two literals of one kind, a number or a truth value each,
    /// over conditions of numbers or truth values (KERNEL.43): each cell's
    /// answer the literal its condition's truth takes, none parted, a lane of
    /// numbers or of truth values; `None` for any other.
    #[inline(never)]
    fn chosen(&self, conditions: &Lane, a: &Expr, b: &Expr) -> Option<Lane> {
        match (conditions, a, b) {
            (Lane::Truths(ts), Expr::Number(x), Expr::Number(y)) => {
                let mut numbers = self.take_numbers();
                numbers.extend(ts.iter().map(|&t| if t { *x } else { *y }));
                Some(Lane::Numbers(numbers))
            }
            (Lane::Numbers(cs), Expr::Number(x), Expr::Number(y)) => {
                let mut numbers = self.take_numbers();
                numbers.extend(cs.iter().map(|&c| if c != 0.0 { *x } else { *y }));
                Some(Lane::Numbers(numbers))
            }
            (Lane::Truths(ts), Expr::Bool(x), Expr::Bool(y)) => {
                let mut truths = self.take_truths();
                truths.extend(ts.iter().map(|&t| if t { *x } else { *y }));
                Some(Lane::Truths(truths))
            }
            (Lane::Numbers(cs), Expr::Bool(x), Expr::Bool(y)) => {
                let mut truths = self.take_truths();
                truths.extend(cs.iter().map(|&c| if c != 0.0 { *x } else { *y }));
                Some(Lane::Truths(truths))
            }
            _ => None,
        }
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

    /// The library's function called at each cell over operands the run
    /// computed, and the references it reads there, noted against the run's
    /// shape in a test (KERNEL.30's oracle).
    #[cfg(test)]
    fn noted_call(&self, name: &str, ops: &[&Operand<'_>]) {
        taken::note(taken::Path::Call(name.to_ascii_uppercase()));
        let known = self.lib.known(name);
        for (j, x) in ops.iter().enumerate() {
            let Operand::Ref(r) = x else {
                continue;
            };
            // A vector a lookup or a criterion searches, the same at every
            // cell, its index shared by the run's cells (KERNEL.8).
            let found = match (r.found, &r.sheet) {
                (Some(_), _) | (None, None) => true,
                (None, Some(name)) => self.env.model().find_sheet(name).is_some(),
            };
            if found && known.is_some_and(|k| k.searches(j)) && held_whole(r) {
                taken::note(taken::Path::Keyed);
                continue;
            }
            let one = matches!(r.shape, Part::Cell);
            let array = match (r.found, &r.sheet) {
                _ if r.last_sheet.is_some() => Some(false),
                (Some((first, _)), _) => Some(self.env.dense(usize::from(first)).is_some()),
                (None, None) => Some(self.env.dense(self.cells[0].0).is_some()),
                (None, Some(name)) => self
                    .env
                    .model()
                    .find_sheet(name)
                    .map(|s| self.env.dense(s).is_some()),
            };
            taken::note(match (one, array) {
                (_, None) => continue,
                (true, Some(true)) => taken::Path::Slot,
                (false, Some(true)) => taken::Path::Array,
                (true, Some(false)) => taken::Path::Lookup,
                (false, Some(false)) => taken::Path::Cells,
            });
        }
    }

    /// An expression evaluated at each cell by [`Ctx`] itself.
    #[inline(never)]
    fn each_cell(&self, e: &Expr, sel: Sel) -> Lane {
        #[cfg(test)]
        taken::note(taken::Path::Each);
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

/// An operand's lane for a function of numbers (KERNEL.28): a reference's,
/// read where it stands, or a value's own.
fn operand_lane<'l>(read: &'l Option<Lane>, op: &'l Operand<'_>) -> Option<&'l Lane> {
    match (read, op) {
        (Some(lane), _) => Some(lane),
        (None, Operand::Value(lane)) => Some(lane),
        (None, Operand::Ref(_)) => None,
    }
}

/// The number an operand's lane holds at the selection's `k`th cell, when it
/// holds one there (KERNEL.28).
fn number_at(lane: Option<&Lane>, k: usize) -> Option<f64> {
    match lane? {
        Lane::Numbers(ns) => Some(ns[k]),
        Lane::Same(Ok(Value::Number(n))) => Some(*n),
        Lane::Each(each) => match &each[k] {
            Ok(Value::Number(n)) => Some(*n),
            _ => None,
        },
        Lane::Same(_) | Lane::Truths(_) => None,
    }
}

// ---- the paths a run takes, read off a parse (KERNEL.30) -------------------

/// The paths a run of a shape takes through its parse, read off the parse
/// without computing it, as [`Many::scalars`] and the calls below it choose
/// each node's (KERNEL.30's census): each the cost class of a kind of read.
/// The walk stands beside the dispatch and asks what it asks, by the same
/// predicates ([`is_operand`], [`ref_operand`], [`COUNTED`], the library's
/// known functions), and a test notes the paths runs take ([`taken`]) and
/// holds each to a kind this names.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub(crate) struct Paths {
    /// A reference into a sheet kept as an array whose column moves with
    /// the cell (KERNEL.43): one cell read for a run's cells side by side in
    /// a row as one slice of the array's numbers, or for its cells one below
    /// the other as one column of them where its row moves too (KERNEL.31),
    /// or a range folded over its rows' slices.
    pub(crate) stride: bool,
    /// One cell of a sheet kept as an array, read by its index; its column
    /// held, once for a run's cells of one row.
    pub(crate) slot: bool,
    /// A range of a sheet kept as an array, walked by index; its columns
    /// held at both corners, the same range at each of a run's cells.
    pub(crate) array: bool,
    /// One cell through the environment's lookups.
    pub(crate) lookup: bool,
    /// A range walked through its cells, at each cell.
    pub(crate) cells: bool,
    /// A cell every `$` holds, read once a run.
    pub(crate) once: bool,
    /// A node [`Ctx`] computes at each cell, everything under it with it.
    pub(crate) each: bool,
    /// A vector a lookup or a criterion searches, the same at every cell of
    /// the run, by its index or a bisection and never walked whole at each
    /// (KERNEL.8).
    pub(crate) keyed: bool,
    /// The functions called at each cell, by the library over operands the
    /// run computed or by [`Ctx`], in upper case, sorted, each once.
    pub(crate) calls: Vec<String>,
}

/// The paths of a shape's parse, `arrays` saying which sheets the
/// environment keeps as arrays (a machine's twins, as its last frame's end
/// left them, and its own sheets a read takes as arrays, KERNEL.31), `homes`
/// the sheets the shape's cells stand on, which a reference that names no
/// sheet reads, `runs` whether the shape has runs, whose cells side by side
/// read a reference that moves with them by slices (KERNEL.43), where a cell
/// alone reads it by index. The walk keeps its own stack, as every walk of a
/// parse does, so that a chain as long as Excel holds costs it no depth.
pub(crate) fn paths(
    lib: &Library,
    model: &crate::sheet::Workbook,
    e: &Expr,
    arrays: &dyn Fn(usize) -> bool,
    homes: &[usize],
    runs: bool,
) -> Paths {
    let mut walk = Walk {
        lib,
        model,
        arrays,
        homes,
        runs,
        paths: Paths::default(),
    };
    let mut stack: Vec<&Expr> = vec![e];
    while let Some(e) = stack.pop() {
        walk.node(e, &mut stack);
    }
    let mut paths = walk.paths;
    paths.calls.sort_unstable();
    paths.calls.dedup();
    paths
}

/// One parse's walk.
struct Walk<'a> {
    lib: &'a Library,
    model: &'a crate::sheet::Workbook,
    arrays: &'a dyn Fn(usize) -> bool,
    homes: &'a [usize],
    runs: bool,
    paths: Paths,
}

impl Walk<'_> {
    /// A node as [`Many::scalars`] takes it, its operands onto the stack.
    fn node<'e>(&mut self, e: &'e Expr, stack: &mut Vec<&'e Expr>) {
        match e {
            Expr::Number(_) | Expr::Text(_) | Expr::Bool(_) | Expr::Error(_) | Expr::Empty => {}
            Expr::Ref(r) if matches!(r.shape, Part::Cell) && r.last_sheet.is_none() => self.one(r),
            Expr::Neg(x) | Expr::Pos(x) | Expr::Percent(x) => stack.push(x),
            Expr::Binary(_, a, b) => {
                stack.push(b);
                stack.push(a);
            }
            Expr::Call(name, args) => self.call(e, name, args, stack),
            Expr::Ref(_) | Expr::Name { .. } => self.each(e),
        }
    }

    /// The sheet a reference to one sheet reads, as the run finds it:
    /// `Some(Some(sheet))`, `Some(None)` for the formula's own, each sheet
    /// the shape's cells stand on; `None` for a sheet the grid does not hold.
    fn sheet_of(&self, r: &RefExpr) -> Option<Option<usize>> {
        match (r.found, &r.sheet) {
            (Some((first, _)), _) => Some(Some(usize::from(first))),
            (None, None) => Some(None),
            (None, Some(name)) => self.model.find_sheet(name).map(Some),
        }
    }

    /// Whether a reference's sheet is kept as an array: the formula's own
    /// where every sheet its cells stand on is (KERNEL.31).
    fn kept(&self, sheet: Option<usize>) -> bool {
        match sheet {
            Some(s) => (self.arrays)(s),
            None => !self.homes.is_empty() && self.homes.iter().all(|&s| (self.arrays)(s)),
        }
    }

    /// One cell as [`Many::reads`] reads it.
    fn one(&mut self, r: &RefExpr) {
        match self.sheet_of(r) {
            None => {}
            Some(_) if r.corners.row_abs1 && r.corners.col_abs1 => self.paths.once = true,
            Some(sheet) if self.kept(sheet) && self.runs && !r.corners.col_abs1 => {
                self.paths.stride = true
            }
            Some(sheet) if self.kept(sheet) => self.paths.slot = true,
            Some(_) => self.paths.lookup = true,
        }
    }

    /// A reference read at each cell, by the library or by [`Ctx`]: an
    /// array walked by index where its sheet is kept as one, its cells
    /// through lookups where it is not; a span of sheets through its cells.
    fn at_each(&mut self, r: &RefExpr) {
        let array = match self.sheet_of(r) {
            _ if r.last_sheet.is_some() => false,
            None => return,
            Some(sheet) => self.kept(sheet),
        };
        match (matches!(r.shape, Part::Cell), array) {
            (true, true) => self.paths.slot = true,
            (false, true) => self.paths.array = true,
            (true, false) => self.paths.lookup = true,
            (false, false) => self.paths.cells = true,
        }
    }

    /// A node [`Ctx`] computes at each cell: everything under it at each
    /// cell with it, every function it calls a call at each cell.
    fn each(&mut self, e: &Expr) {
        self.paths.each = true;
        let mut stack = vec![e];
        while let Some(x) = stack.pop() {
            match x {
                Expr::Ref(r) => self.at_each(r),
                Expr::Call(name, args) => {
                    self.paths.calls.push(name.to_ascii_uppercase());
                    stack.extend(args.iter());
                }
                Expr::Neg(y) | Expr::Pos(y) | Expr::Percent(y) => stack.push(y),
                Expr::Binary(_, a, b) => {
                    stack.push(a);
                    stack.push(b);
                }
                Expr::Name { .. }
                | Expr::Number(_)
                | Expr::Text(_)
                | Expr::Bool(_)
                | Expr::Error(_)
                | Expr::Empty => {}
            }
        }
    }

    /// A call as [`Many::call`] and the paths below it take it.
    fn call<'e>(&mut self, e: &'e Expr, name: &str, args: &'e [Expr], stack: &mut Vec<&'e Expr>) {
        if self.lib.form(name).is_some() {
            match self.lib.known(name) {
                Some(Known::If) if (2..=3).contains(&args.len()) => stack.extend(args.iter()),
                Some(Known::Choose) if args.len() >= 2 => stack.extend(args.iter()),
                Some(Known::If | Known::Choose) => {}
                Some(Known::IfError | Known::IfNa) if args.len() == 2 => stack.extend(args.iter()),
                Some(Known::IfError | Known::IfNa) => {}
                // The cell picked read as one cell is (KERNEL.8): by its slot
                // where its sheet is kept as an array, else by its position.
                Some(Known::Index) if index_by_run(self.lib, args) => {
                    if let Some(r) = ref_operand(self.lib, &args[0]) {
                        if let Some(sheet) = self.sheet_of(r) {
                            if self.kept(sheet) {
                                self.paths.slot = true;
                            } else {
                                self.paths.lookup = true;
                            }
                        }
                    }
                    stack.extend(args[1..].iter());
                }
                _ => self.each(e),
            }
            return;
        }
        // A function the library does not hold: its operands are never
        // computed, and the shape's own list names it.
        if self.lib.strict(name).is_none() {
            return;
        }
        if !args.iter().all(|a| is_operand(self.lib, a)) {
            self.each(e);
            return;
        }
        let mut refs: Vec<&RefExpr> = Vec::new();
        let mut searched: Vec<bool> = Vec::new();
        let known = self.lib.known(name);
        for (j, a) in args.iter().enumerate() {
            match ref_operand(self.lib, a) {
                Some(r) => {
                    refs.push(r);
                    searched.push(known.is_some_and(|k| k.searches(j)));
                }
                None => stack.push(a),
            }
        }
        let one_cell = |r: &RefExpr| matches!(r.shape, Part::Cell) && r.last_sheet.is_none();
        // Whether the run computes the call itself, as `Many::known_call`
        // answers; and the reads it makes when it does.
        let by_run = match known {
            Some(Known::Abs | Known::Int | Known::Sin) => {
                args.len() == 1 && refs.iter().all(|r| one_cell(r))
            }
            Some(Known::Mod) => args.len() == 2 && refs.iter().all(|r| one_cell(r)),
            Some(Known::Row | Known::Column) => match (args.len(), refs.as_slice()) {
                (0, _) => true,
                (1, [r]) => r.found.is_some() || (r.sheet.is_none() && r.last_sheet.is_none()),
                _ => false,
            },
            Some(Known::Not) => args.len() == 1 && refs.is_empty(),
            Some(Known::Sum | Known::Min | Known::Max | Known::And | Known::Or) => {
                args.len() <= COUNTED
                    && refs.iter().all(|r| {
                        r.last_sheet.is_none()
                            && match self.sheet_of(r) {
                                None => true,
                                Some(sheet) => self.kept(sheet),
                            }
                    })
            }
            _ => false,
        };
        if !by_run {
            self.paths.calls.push(name.to_ascii_uppercase());
            for (r, searched) in refs.into_iter().zip(searched) {
                // A vector the run's cells search alike, by the index they
                // share or a bisection (KERNEL.8).
                if self.runs && searched && held_whole(r) && self.sheet_of(r).is_some() {
                    self.paths.keyed = true;
                } else {
                    self.at_each(r);
                }
            }
            return;
        }
        match known {
            Some(Known::Sum | Known::Min | Known::Max | Known::And | Known::Or) => {
                for r in refs {
                    if self.sheet_of(r).is_none() {
                        continue;
                    }
                    if self.runs && moves_with_column(r) {
                        self.paths.stride = true;
                    } else {
                        self.paths.array = true;
                    }
                }
            }
            Some(Known::Abs | Known::Int | Known::Sin | Known::Mod) => {
                for r in refs {
                    self.one(r);
                }
            }
            _ => {}
        }
    }
}

/// How often a run read a reference by slices or a held column once, folded
/// an aggregate by slices, and chose `IF`'s literals by its condition's lane
/// (KERNEL.43), and read a reference's cells one below the other as one
/// column of an array (KERNEL.31), counted in a test on its own thread, so
/// that a law over runs can say that each was taken.
#[cfg(test)]
pub(crate) mod sliced {
    use std::cell::Cell;

    pub(crate) const READ: usize = 0;
    pub(crate) const HELD: usize = 1;
    pub(crate) const FOLDED: usize = 2;
    pub(crate) const CHOSEN: usize = 3;
    pub(crate) const DOWN: usize = 4;

    thread_local! {
        static COUNTS: Cell<[usize; 5]> = const { Cell::new([0; 5]) };
    }

    pub(crate) fn note(i: usize) {
        COUNTS.with(|c| {
            let mut a = c.get();
            a[i] += 1;
            c.set(a);
        });
    }

    /// The counts since the last time they were taken.
    pub(crate) fn take() -> [usize; 5] {
        COUNTS.with(|c| c.replace([0; 5]))
    }
}

/// The paths runs took, noted in a test against the shape each run is of
/// (KERNEL.30's oracle: the census names every one of them). Noting is on
/// only between [`taken::begin`] and [`taken::end`], on the thread that
/// asked, and only in a test build.
#[cfg(test)]
pub(crate) mod taken {
    use std::cell::{Cell, RefCell};
    use std::collections::BTreeSet;

    /// A path a run took.
    #[derive(Clone, Debug, PartialEq, Eq, PartialOrd, Ord)]
    pub(crate) enum Path {
        Stride,
        Slot,
        Array,
        Lookup,
        Cells,
        Once,
        Each,
        /// A vector a lookup or a criterion searches by its index or a
        /// bisection, the same at every cell (KERNEL.8).
        Keyed,
        /// The library's function, called at each cell.
        Call(String),
    }

    thread_local! {
        static SHAPE: Cell<usize> = const { Cell::new(usize::MAX) };
        static NOTED: RefCell<Option<BTreeSet<(usize, Path)>>> = const { RefCell::new(None) };
    }

    /// Noting begins on this thread.
    pub(crate) fn begin() {
        NOTED.with(|n| *n.borrow_mut() = Some(BTreeSet::new()));
    }

    /// Every path noted since [`begin`], by shape; noting ends.
    pub(crate) fn end() -> BTreeSet<(usize, Path)> {
        NOTED.with(|n| n.borrow_mut().take().unwrap_or_default())
    }

    /// The shape of the run whose paths are noted next.
    pub(crate) fn shape(s: usize) {
        SHAPE.with(|c| c.set(s));
    }

    /// A path taken by the run being evaluated.
    pub(crate) fn note(p: Path) {
        NOTED.with(|n| {
            if let Some(set) = n.borrow_mut().as_mut() {
                set.insert((SHAPE.with(Cell::get), p));
            }
        });
    }
}
