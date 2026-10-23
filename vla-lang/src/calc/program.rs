//! A cell alone lowered once to a register program (KERNEL.44, 2026-10-16).
//!
//! A cell no run holds, a shape's only cell or one of a run's few, was
//! evaluated by [`Ctx`] walking its parse, and the walk cost 25 to 90 ns a
//! node of its own: a frame a node, a name's case checked and looked up at
//! each call, an operand made and moved, where the arithmetic itself cost a
//! few (measured in the item's scoping: Kairos's cells 250 ns each, `chain`'s
//! balances 83). Here a shape's parse is lowered once, when the plan keeps
//! the shape (`super::shape`, beside what a run holds), and each cell of it
//! evaluated alone runs the program at its offset from the shape's first
//! cell, with no walk: about 18 ns a balance of `chain` and 40 a cell of
//! Kairos.
//!
//! A program's registers hold numbers, a truth value as 1 or 0, each
//! register's kind read off the parse. Its operations are the evaluator's
//! own functions over the values the evaluator would hand them, in the
//! evaluator's order: [`arithmetic`] and [`compares`] over
//! [`compare_numbers`], a sign and a percent through [`finite`], `IF` and
//! `CHOOSE` as jumps that compute the branch taken alone, the aggregates by
//! the library's own rule and fold (`library::Rule`, `library::Fold`) over a
//! value given and over a range walked as [`Ctx::each`] walks it, `NOT`,
//! `ABS`, `INT`, `SIN` and `MOD` by the library's meanings, `ROW` and
//! `COLUMN` of the cell or of a reference moved to it, and any other strict
//! function through its pointer with the operands [`Ctx`] would hand it. A
//! reference is held as its parts, the sheet its placement found or the
//! cell's own, its shape, corners and rectangle, which
//! `formula::moved_at` moves as [`RefExpr::moved`] does.
//!
//! Its answer is `Ctx`'s by construction. Every read and every operation
//! whose answer may be a value no register holds is guarded: a read of a
//! text, a truth value, an error or a reason, an operation answering an
//! error, a fold meeting in a range a value neither a number nor nothing,
//! `CHOOSE`'s index outside its values, a call answering anything but a
//! number. A guard that fails leaves the cell to `Ctx`, which evaluates it
//! from its start as it always has, an evaluation having no effect but its
//! answer; so a program answers only where every value it met was a number,
//! nothing, or a truth value it computed. Nothing reads as 0, which every
//! node a program holds reads alike: the arithmetic, a comparison of
//! numbers, a condition, the rule for a value given, a cell's settled
//! answer. A comparison of a number with a truth value is left to `Ctx`
//! (nothing compares with a truth value as FALSE), and so is everything
//! [`Why`] names.
//!
//! The machine runs a program for every cell it evaluates alone and for a
//! run's selection of fewer cells than `shape::SHORT`, its reads through
//! [`Source`]; nothing here allocates but a call through the library's
//! pointer of three operands or more, which `Ctx` allocates for too.

use std::borrow::Cow;

use crate::refers::Shape as Part;

use super::eval::{
    arithmetic, compare_numbers, compares, finite, Arg, Ctx, Env, Known, Library, Reason, Strict,
};
use super::formula::{moved_at, BinOp, Corners, Expr, RefExpr};
use super::graph::{clamped, Area, Areas, CellId};
use super::library::{self, All, Axis, Fold, Least, Most, Rule, Some_, Sum};
use super::lookups;
use super::value::{ErrorKind, Value};
use super::vector::{ref_operand, yields_value};

/// The most registers a program holds; a parse that needs more is left to
/// [`Ctx`].
const REGISTERS: usize = 32;

/// How deeply the lowering nests before it leaves a parse to [`Ctx`]: the
/// lowering recurses as the parse nests, beneath the plan's frame, and a
/// formula nests seldom past a few levels.
const DEPTH: u32 = 48;

/// Why a parse is left to [`Ctx`]: the first node the lowering met that a
/// program does not hold, named in a census row (`(unlowered <why>)`).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum Why {
    /// A text where one value is wanted, or `&`.
    Text,
    /// An error written as a value, or a call `Ctx` answers with an error
    /// whatever its operands (a count of operands the function never
    /// takes, `ROW` of a value).
    Error,
    /// A defined name.
    Name,
    /// A range, a whole column or a whole row where one value is wanted.
    Range,
    /// A form other than the language's `IF` and `CHOOSE`, or a function
    /// the library does not hold.
    Form,
    /// A comparison of a number with a truth value, or the values of an
    /// `IF` or a `CHOOSE` of two kinds.
    Kinds,
    /// A span of sheets.
    Span,
    /// A sheet found by its name at each evaluation, its placement having
    /// found none.
    Sheet,
    /// An omitted argument where one value is wanted, or `CHOOSE`'s omitted
    /// value handed to a function through its pointer, which `Ctx` hands
    /// nothing.
    Omitted,
    /// More registers, levels or operations than a program holds.
    Deep,
    /// An array anywhere in the parse (KERNEL.8, slice f, `vector::arrays_in`),
    /// which no register holds.
    Array,
}

impl Why {
    /// The word a census row spells.
    pub(crate) fn word(self) -> &'static str {
        match self {
            Why::Text => "text",
            Why::Error => "error",
            Why::Name => "name",
            Why::Range => "range",
            Why::Form => "form",
            Why::Kinds => "kinds",
            Why::Span => "span",
            Why::Sheet => "sheet",
            Why::Omitted => "omitted",
            Why::Deep => "deep",
            Why::Array => "array",
        }
    }
}

/// What a register holds, read off the parse: a number, or a truth value as
/// 1 or 0.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum Kind {
    Number,
    Truth,
}

/// A program's answer, as its cell holds it.
#[derive(Clone, Copy, Debug, PartialEq)]
pub(crate) enum Answer {
    Number(f64),
    Truth(bool),
}

/// An operand of an aggregate or a call: a register and its kind, or a
/// reference, which the call takes as the cells it names.
#[derive(Clone, Copy, Debug, PartialEq)]
enum Operand {
    Reg(u8, Kind),
    Ref(u16),
}

/// One operation. `to` is the register an operation's answer goes to, and
/// the one its first operand stands in; `b` its second operand's.
#[derive(Clone, Copy, Debug, PartialEq)]
enum Op {
    /// A number, or a truth value as 1 or 0.
    Const {
        to: u8,
        x: f64,
    },
    /// A cell's number, by its reference.
    Read {
        to: u8,
        at: u16,
    },
    Arith {
        to: u8,
        op: BinOp,
        b: u8,
    },
    Compare {
        to: u8,
        op: BinOp,
        b: u8,
    },
    Neg {
        to: u8,
    },
    Percent {
        to: u8,
    },
    Not {
        to: u8,
    },
    Abs {
        to: u8,
    },
    Int {
        to: u8,
    },
    Sin {
        to: u8,
    },
    Mod {
        to: u8,
        b: u8,
    },
    /// `ROW()` or `COLUMN()`, the cell's own.
    Here {
        to: u8,
        axis: Axis,
    },
    /// `ROW` or `COLUMN` of a reference, its first cell moved to the cell.
    Placed {
        to: u8,
        axis: Axis,
        at: u16,
    },
    /// An aggregate over `operands[from..from + count]`.
    Fold {
        to: u8,
        known: Known,
        from: u16,
        count: u16,
    },
    /// A strict function, `calls[f]`, over `operands[from..from + count]`.
    Call {
        to: u8,
        f: u16,
        from: u16,
        count: u16,
    },
    /// On to `to` where the condition is false.
    Unless {
        cond: u8,
        to: u16,
    },
    Jump {
        to: u16,
    },
    /// On to the value `CHOOSE`'s index chooses, `table[from + k - 1]`.
    Choose {
        index: u8,
        from: u16,
        count: u16,
    },
    Done {
        from: u8,
    },
}

/// A reference as a program holds it: the sheet its placement found, or
/// none for the cell's own; its shape, its corners as written and its
/// rectangle, which [`moved_at`] moves; and the place of its text among the
/// program's, where a call takes it as cells and labels them with it.
#[derive(Clone, Copy, Debug, PartialEq)]
struct Ref {
    sheet: Option<u8>,
    shape: Part,
    corners: Corners,
    rect: (u32, u32, u32, u32),
    text: u16,
}

impl Ref {
    #[inline]
    fn moved(&self, off: (i64, i64)) -> Option<(u32, u32, u32, u32)> {
        moved_at(self.shape, &self.corners, self.rect, off.0, off.1)
    }

    #[inline]
    fn sheet(&self, home: usize) -> usize {
        self.sheet.map_or(home, usize::from)
    }

    /// The area it names from a cell of sheet `home` at an offset, as
    /// `graph::areas_at` names it for a reference to one sheet; `None` where
    /// a corner leaves the sheet, which is `#REF!`.
    #[inline]
    fn area(&self, home: usize, off: (i64, i64)) -> Option<Area> {
        let (top, left, bottom, right) = self.moved(off)?;
        Some(Area {
            sheet: self.sheet(home),
            top,
            left,
            bottom,
            right,
        })
    }
}

/// A shape's parse lowered. See the module's comment.
#[derive(Clone)]
pub(crate) struct Program {
    code: Vec<Op>,
    refs: Vec<Ref>,
    /// The texts of the references a call takes as cells.
    texts: Vec<String>,
    operands: Vec<Operand>,
    /// `CHOOSE`'s values, where each begins.
    table: Vec<u16>,
    calls: Vec<Strict>,
    kind: Kind,
    /// The call at the program's root, the cell's whole formula, by its
    /// place among the calls: a value of it that is no number is the cell's
    /// answer all the same, handed to the environment (`Source::answered`),
    /// which is `Ctx`'s answer for the cell, so that a lookup answering a
    /// text or a miss in a cell alone is not computed again from the start
    /// (KERNEL.8). Every program's loop is the one it was: an operation of
    /// the root's own in it cost the chain fixture's cells alone a tenth of
    /// their time, and one of `INDEX`'s half again (measured, 0.235 ms a
    /// frame to 0.260 and 0.363).
    root: Option<u16>,
}

impl std::fmt::Debug for Program {
    /// The program whole, but for its functions' addresses, counted.
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("Program")
            .field("code", &self.code)
            .field("refs", &self.refs)
            .field("texts", &self.texts)
            .field("operands", &self.operands)
            .field("table", &self.table)
            .field("calls", &self.calls.len())
            .field("kind", &self.kind)
            .field("root", &self.root)
            .finish()
    }
}

/// What a program reads its cells through: the environment, and a cell's
/// value as a register holds it.
pub(crate) trait Source: Env {
    /// A cell's number, as [`number_of`] reads [`Env::read`]'s answer; an
    /// environment that can say it without making the value says it so.
    fn number(&self, id: CellId) -> Option<f64> {
        number_of(self.read(id))
    }

    /// The answer of the call at a program's root that is no number, the
    /// cell's all the same (KERNEL.8), told just before the program answers
    /// `None`: an environment that keeps it takes it as the cell's value; one
    /// that does not passes it by, and the cell is `Ctx`'s.
    fn answered(&self, _value: Value) {}
}

/// A read's answer as a register holds it: a number itself, 0 for a cell
/// holding nothing; `None` for a text, a truth value, an error or a reason,
/// which no register holds.
pub(crate) fn number_of(got: Result<Value, Reason>) -> Option<f64> {
    match got {
        Ok(Value::Number(n)) => Some(n),
        Ok(Value::Empty) => Some(0.0),
        _ => None,
    }
}

#[inline]
fn truth(b: bool) -> f64 {
    if b {
        1.0
    } else {
        0.0
    }
}

/// An operation's answer as a register holds it: a number, or the guard
/// that fails for an error.
#[inline]
fn num(v: Value) -> Option<f64> {
    match v {
        Value::Number(n) => Some(n),
        _ => guard(guarded::ANSWER),
    }
}

/// A guard failed: the cell is `Ctx`'s. Counted in a test.
#[inline]
fn guard<T>(_which: usize) -> Option<T> {
    #[cfg(test)]
    guarded::note(_which);
    None
}

impl Program {
    /// A parse lowered, or why it is left to [`Ctx`].
    pub(crate) fn lower(lib: &Library, e: &Expr) -> Result<Program, Why> {
        // A parse holding an array anywhere is `Ctx`'s (KERNEL.8, slice f).
        if super::vector::arrays_in(lib, e) {
            return Err(Why::Array);
        }
        let mut l = Lowering {
            lib,
            code: Vec::new(),
            refs: Vec::new(),
            texts: Vec::new(),
            operands: Vec::new(),
            table: Vec::new(),
            calls: Vec::new(),
            depth: 0,
            root: None,
        };
        let kind = l.expr(e, 0)?;
        l.code.push(Op::Done { from: 0 });
        if l.code.len() > usize::from(u16::MAX) {
            return Err(Why::Deep);
        }
        l.code.shrink_to_fit();
        l.refs.shrink_to_fit();
        l.operands.shrink_to_fit();
        Ok(Program {
            code: l.code,
            refs: l.refs,
            texts: l.texts,
            operands: l.operands,
            table: l.table,
            calls: l.calls,
            kind,
            root: l.root,
        })
    }

    /// The program run at a cell, `off` the cell's offset from its shape's
    /// first cell: its answer, or `None` where a guard failed and the cell
    /// is `Ctx`'s. The environment is told the cell first, as `Ctx` is.
    pub(crate) fn run<S: Source>(
        &self,
        src: &S,
        lib: &Library,
        id: CellId,
        off: (i64, i64),
    ) -> Option<Answer> {
        let mut reg = [0.0f64; REGISTERS];
        let mut pc = 0usize;
        loop {
            match self.code[pc] {
                Op::Const { to, x } => reg[usize::from(to)] = x,
                Op::Read { to, at } => {
                    let r = &self.refs[usize::from(at)];
                    let Some((row, col, _, _)) = r.moved(off) else {
                        return guard(guarded::OFF);
                    };
                    let Some(x) = src.number((r.sheet(id.0), row, col)) else {
                        return guard(guarded::READ);
                    };
                    reg[usize::from(to)] = x;
                }
                Op::Arith { to, op, b } => {
                    let t = usize::from(to);
                    reg[t] = num(arithmetic(op, reg[t], reg[usize::from(b)]))?;
                }
                Op::Compare { to, op, b } => {
                    let t = usize::from(to);
                    reg[t] = truth(compares(op, compare_numbers(reg[t], reg[usize::from(b)])));
                }
                Op::Neg { to } => {
                    let t = usize::from(to);
                    reg[t] = num(finite(-reg[t]))?;
                }
                Op::Percent { to } => {
                    let t = usize::from(to);
                    reg[t] = num(finite(reg[t] / 100.0))?;
                }
                // `library::negated`: a number is true unless 0.
                Op::Not { to } => {
                    let t = usize::from(to);
                    reg[t] = truth(reg[t] == 0.0);
                }
                Op::Abs { to } => {
                    let t = usize::from(to);
                    reg[t] = num(library::abs_of(reg[t]))?;
                }
                Op::Int { to } => {
                    let t = usize::from(to);
                    reg[t] = num(library::int_of(reg[t]))?;
                }
                Op::Sin { to } => {
                    let t = usize::from(to);
                    reg[t] = num(library::sin_of(reg[t]))?;
                }
                Op::Mod { to, b } => {
                    let t = usize::from(to);
                    reg[t] = num(library::mod_of(reg[t], reg[usize::from(b)]))?;
                }
                Op::Here { to, axis } => reg[usize::from(to)] = axis.of(id.1, id.2),
                Op::Placed { to, axis, at } => {
                    let Some((top, left, _, _)) = self.refs[usize::from(at)].moved(off) else {
                        return guard(guarded::OFF);
                    };
                    reg[usize::from(to)] = axis.of(top, left);
                }
                Op::Fold {
                    to,
                    known,
                    from,
                    count,
                } => {
                    let from = usize::from(from);
                    let ops = &self.operands[from..from + usize::from(count)];
                    reg[usize::from(to)] = match known {
                        Known::Sum => self.fold::<Sum, S>(src, ops, &reg, id, off),
                        Known::Min => self.fold::<Least, S>(src, ops, &reg, id, off),
                        Known::Max => self.fold::<Most, S>(src, ops, &reg, id, off),
                        Known::And => self.fold::<All, S>(src, ops, &reg, id, off),
                        _ => self.fold::<Some_, S>(src, ops, &reg, id, off),
                    }?;
                }
                Op::Call { to, f, from, count } => {
                    let from = usize::from(from);
                    let ops = &self.operands[from..from + usize::from(count)];
                    reg[usize::from(to)] = self.number(src, lib, f, ops, &reg, id, off)?;
                }
                Op::Unless { cond, to } => {
                    if reg[usize::from(cond)] == 0.0 {
                        pc = usize::from(to);
                        continue;
                    }
                }
                Op::Jump { to } => {
                    pc = usize::from(to);
                    continue;
                }
                Op::Choose { index, from, count } => {
                    let index = Value::Number(reg[usize::from(index)]);
                    let Ok(k) = library::chosen(&index, usize::from(count)) else {
                        return guard(guarded::INDEX);
                    };
                    pc = usize::from(self.table[usize::from(from) + k - 1]);
                    continue;
                }
                Op::Done { from } => {
                    let x = reg[usize::from(from)];
                    return Some(match self.kind {
                        Kind::Number => Answer::Number(x),
                        Kind::Truth => Answer::Truth(x != 0.0),
                    });
                }
            }
            pc += 1;
        }
    }

    /// An aggregate over its operands by the library's rule and fold: a
    /// register by the rule for a value given (1 or 0 being a truth value's
    /// number), a reference by the rule for a range, its cells walked as
    /// `Ctx::each` walks them.
    #[inline(never)]
    fn fold<F: Fold, S: Source>(
        &self,
        src: &S,
        ops: &[Operand],
        reg: &[f64; REGISTERS],
        id: CellId,
        off: (i64, i64),
    ) -> Option<f64> {
        let mut fold = F::default();
        for op in ops {
            match *op {
                Operand::Reg(r, _) => {
                    fold.add(<F::Rule as Rule>::given_number(reg[usize::from(r)]));
                }
                Operand::Ref(at) => {
                    let Some(a) = self.refs[usize::from(at)].area(id.0, off) else {
                        return guard(guarded::OFF);
                    };
                    walk::<F::Rule, S, _>(src, &a, &mut |x| fold.add(x))?;
                }
            }
        }
        match fold.answer() {
            Value::Number(n) => Some(n),
            Value::Bool(b) => Some(truth(b)),
            _ => guard(guarded::ANSWER),
        }
    }

    /// The call at `calls[f]`, its answer guarded to be a number; at the
    /// program's root another answer handed to the environment as the
    /// cell's (`root`, KERNEL.8), no guard failing. Out of the loop, as every
    /// value a call makes, so that the loop holds none.
    #[allow(clippy::too_many_arguments)]
    #[inline(never)]
    fn number<S: Source>(
        &self,
        src: &S,
        lib: &Library,
        f: u16,
        ops: &[Operand],
        reg: &[f64; REGISTERS],
        id: CellId,
        off: (i64, i64),
    ) -> Option<f64> {
        match self.call(src, lib, self.calls[usize::from(f)], ops, reg, id, off) {
            Ok(Value::Number(n)) => Some(n),
            Ok(v) if self.root == Some(f) => {
                src.answered(v.settled());
                None
            }
            _ => guard(guarded::CALL),
        }
    }

    /// A strict function through its pointer, over the operands `Ctx` would
    /// hand it: a reference as the cells it names, labelled with its text, or
    /// `#REF!` where it leaves the sheet; a register as its value. Its answer
    /// as it stands, which a call inside the program guards to be a number
    /// and a call at its root keeps whatever its kind (KERNEL.8).
    #[allow(clippy::too_many_arguments)]
    #[inline(never)]
    fn call<S: Source>(
        &self,
        src: &S,
        lib: &Library,
        f: Strict,
        ops: &[Operand],
        reg: &[f64; REGISTERS],
        id: CellId,
        off: (i64, i64),
    ) -> Result<Value, Reason> {
        let ctx = Ctx::at(src, lib, off.0, off.1);
        let arg = |op: &Operand| -> Arg<'_> {
            match *op {
                Operand::Reg(r, Kind::Number) => Arg::Scalar(Value::Number(reg[usize::from(r)])),
                Operand::Reg(r, Kind::Truth) => {
                    Arg::Scalar(Value::Bool(reg[usize::from(r)] != 0.0))
                }
                Operand::Ref(at) => {
                    let r = &self.refs[usize::from(at)];
                    match r.area(id.0, off) {
                        Some(a) => Arg::Area {
                            areas: Areas::one(a),
                            written: Cow::Borrowed(&self.texts[usize::from(r.text)]),
                            moved: off,
                        },
                        None => Arg::Scalar(Value::Error(ErrorKind::Ref)),
                    }
                }
            }
        };
        match ops {
            [] => f(&ctx, &[]),
            [a] => f(&ctx, &[arg(a)]),
            [a, b] => f(&ctx, &[arg(a), arg(b)]),
            _ => {
                let args: Vec<Arg> = ops.iter().map(arg).collect();
                f(&ctx, &args)
            }
        }
    }
}

/// A range's cells walked as `Ctx::each` walks them, a sheet kept as an
/// array by index over the sheet's extent and any other through the cells
/// the environment walks for it, each number to `f` as the rule reads a
/// number in a range, an empty slot and a cell not held skipped by the walks
/// themselves; the guard fails at any other value and at a reason, a held
/// cell read as nothing among them (a formula with no place, which the
/// machine never leaves, and which `Ctx` skips).
fn walk<R: Rule, S: Source, F: FnMut(R::Item)>(src: &S, a: &Area, f: &mut F) -> Option<()> {
    if let Some(kept) = src.dense(a.sheet) {
        let Some(a) = clamped(a, src.extent(a.sheet)) else {
            return Some(());
        };
        let walked = kept.each_read(&a, |_, _, read| match read {
            Ok(Value::Number(n)) => {
                f(R::ranged_number(*n));
                Ok(())
            }
            _ => Err(()),
        });
        return match walked {
            Ok(()) => Some(()),
            Err(()) => guard(guarded::RANGE),
        };
    }
    for (row, col) in src.walk(a) {
        match src.read((a.sheet, row, col)) {
            Ok(Value::Number(n)) => f(R::ranged_number(n)),
            _ => return guard(guarded::RANGE),
        }
    }
    Some(())
}

/// Whether an operand may be nothing where a function takes it: `CHOOSE`'s
/// omitted value, which `Ctx` hands on as nothing (an `IF`'s omitted branch
/// is 0 and its missing else FALSE), at any depth of the `IF`s and `CHOOSE`s
/// that answer one of their values.
fn may_be_nothing(lib: &Library, e: &Expr) -> bool {
    match e {
        Expr::Empty => true,
        Expr::Pos(x) => may_be_nothing(lib, x),
        Expr::Call(name, args) if lib.form(name).is_some() => match lib.known(name) {
            Some(Known::If) => args
                .iter()
                .skip(1)
                .any(|a| !matches!(a, Expr::Empty) && may_be_nothing(lib, a)),
            Some(Known::Choose) => args.iter().skip(1).any(|a| may_be_nothing(lib, a)),
            _ => false,
        },
        _ => false,
    }
}

struct Lowering<'l> {
    lib: &'l Library,
    code: Vec<Op>,
    refs: Vec<Ref>,
    texts: Vec<String>,
    operands: Vec<Operand>,
    table: Vec<u16>,
    calls: Vec<Strict>,
    depth: u32,
    /// The call at the root, where the formula is one (KERNEL.8).
    root: Option<u16>,
}

impl Lowering<'_> {
    fn reg(at: usize) -> Result<u8, Why> {
        if at < REGISTERS {
            Ok(at as u8)
        } else {
            Err(Why::Deep)
        }
    }

    /// The next operation's place.
    fn here(&self) -> Result<u16, Why> {
        u16::try_from(self.code.len()).map_err(|_| Why::Deep)
    }

    /// The expression's value into register `at`, as `Ctx::scalar` computes
    /// it, the registers above it its scratch; its kind.
    fn expr(&mut self, e: &Expr, at: usize) -> Result<Kind, Why> {
        self.depth += 1;
        let kind = if self.depth > DEPTH {
            Err(Why::Deep)
        } else {
            self.node(e, at)
        };
        self.depth -= 1;
        kind
    }

    fn node(&mut self, e: &Expr, at: usize) -> Result<Kind, Why> {
        let to = Self::reg(at)?;
        match e {
            Expr::Number(v) => {
                self.code.push(Op::Const { to, x: *v });
                Ok(Kind::Number)
            }
            Expr::Bool(b) => {
                self.code.push(Op::Const { to, x: truth(*b) });
                Ok(Kind::Truth)
            }
            Expr::Ref(r) => {
                if r.last_sheet.is_some() {
                    return Err(Why::Span);
                }
                if !matches!(r.shape, Part::Cell) {
                    return Err(Why::Range);
                }
                let at = self.refer(r, false)?;
                self.code.push(Op::Read { to, at });
                Ok(Kind::Number)
            }
            Expr::Neg(x) => {
                self.expr(x, at)?;
                self.code.push(Op::Neg { to });
                Ok(Kind::Number)
            }
            Expr::Percent(x) => {
                self.expr(x, at)?;
                self.code.push(Op::Percent { to });
                Ok(Kind::Number)
            }
            // Unary plus is its operand, which a value reads as one.
            Expr::Pos(x) => self.expr(x, at),
            Expr::Binary(..) => self.chain(e, at),
            Expr::Call(name, args) => self.call(name, args, at),
            Expr::Text(_) => Err(Why::Text),
            Expr::Error(_) => Err(Why::Error),
            Expr::Empty => Err(Why::Omitted),
            Expr::Name { .. } => Err(Why::Name),
            Expr::Array(_) | Expr::Range(_) => Err(Why::Array),
        }
    }

    /// A chain of operators, its left spine round a loop as `Ctx::chain`
    /// evaluates it: the leftmost operand into `at`, each right operand into
    /// the register above, each link's answer into `at`.
    fn chain(&mut self, e: &Expr, at: usize) -> Result<Kind, Why> {
        let mut links: Vec<(BinOp, &Expr)> = Vec::new();
        let mut left = e;
        while let Expr::Binary(op, l, r) = left {
            links.push((*op, r));
            left = l;
        }
        let (to, b) = (Self::reg(at)?, Self::reg(at + 1)?);
        let mut kind = self.expr(left, at)?;
        for &(op, r) in links.iter().rev() {
            if op == BinOp::Concat {
                return Err(Why::Text);
            }
            let k = self.expr(r, at + 1)?;
            kind = match op {
                BinOp::Add | BinOp::Sub | BinOp::Mul | BinOp::Div | BinOp::Pow => {
                    self.code.push(Op::Arith { to, op, b });
                    Kind::Number
                }
                _ => {
                    if k != kind {
                        return Err(Why::Kinds);
                    }
                    self.code.push(Op::Compare { to, op, b });
                    Kind::Truth
                }
            };
        }
        Ok(kind)
    }

    /// A call, found once: the language's `IF` and `CHOOSE` as jumps, an
    /// aggregate or another of the language's functions by its meaning, any
    /// other strict function through its pointer.
    fn call(&mut self, name: &str, args: &[Expr], at: usize) -> Result<Kind, Why> {
        let to = Self::reg(at)?;
        if self.lib.form(name).is_some() {
            return match self.lib.known(name) {
                Some(Known::If) => self.if_(args, at),
                Some(Known::Choose) => self.choose(args, at),
                Some(Known::Index) => self.index(args, at),
                _ => Err(Why::Form),
            };
        }
        let Some(f) = self.lib.strict(name) else {
            return Err(Why::Form);
        };
        // A count of operands the function never takes is answered with an
        // error by `Ctx`, whatever they are.
        let one = |args: &[Expr]| {
            if args.len() == 1 {
                Ok(())
            } else {
                Err(Why::Error)
            }
        };
        match self.lib.known(name) {
            Some(known @ (Known::Sum | Known::Min | Known::Max | Known::And | Known::Or)) => {
                let (from, count) = self.operands(args, at, false)?;
                self.code.push(Op::Fold {
                    to,
                    known,
                    from,
                    count,
                });
                Ok(match known {
                    Known::And | Known::Or => Kind::Truth,
                    _ => Kind::Number,
                })
            }
            Some(Known::Not) => {
                one(args)?;
                self.expr(&args[0], at)?;
                self.code.push(Op::Not { to });
                Ok(Kind::Truth)
            }
            Some(known @ (Known::Abs | Known::Int | Known::Sin)) => {
                one(args)?;
                self.expr(&args[0], at)?;
                self.code.push(match known {
                    Known::Abs => Op::Abs { to },
                    Known::Int => Op::Int { to },
                    _ => Op::Sin { to },
                });
                Ok(Kind::Number)
            }
            Some(Known::Mod) => {
                let [n, d] = args else {
                    return Err(Why::Error);
                };
                self.expr(n, at)?;
                self.expr(d, at + 1)?;
                let b = Self::reg(at + 1)?;
                self.code.push(Op::Mod { to, b });
                Ok(Kind::Number)
            }
            Some(known @ (Known::Row | Known::Column)) => {
                let axis = if known == Known::Row {
                    Axis::Row
                } else {
                    Axis::Column
                };
                match args {
                    [] => self.code.push(Op::Here { to, axis }),
                    [a] => {
                        let mut a = a;
                        while let Expr::Pos(x) = a {
                            a = x;
                        }
                        let Expr::Ref(r) = a else {
                            return Err(Why::Error);
                        };
                        let at = self.refer(r, false)?;
                        self.code.push(Op::Placed { to, axis, at });
                    }
                    _ => return Err(Why::Error),
                }
                Ok(Kind::Number)
            }
            Some(Known::If | Known::Choose | Known::Index | Known::IfError | Known::IfNa) => {
                Err(Why::Form)
            }
            // An error is what the error tests read, and no register holds
            // one (KERNEL.8).
            Some(Known::IsError | Known::IsNa) => Err(Why::Error),
            Some(
                Known::Match
                | Known::XMatch
                | Known::VLookup
                | Known::HLookup
                | Known::Lookup
                | Known::CountIf
                | Known::SumIf,
            )
            | None => self.called(f, args, at),
        }
    }

    /// A strict function's call through its pointer: its operands, then the
    /// call, its answer guarded to be a number; or, the cell's whole
    /// formula, the program's root, answering its value whatever its kind
    /// once the program's operations are done (KERNEL.8).
    fn called(&mut self, f: Strict, args: &[Expr], at: usize) -> Result<Kind, Why> {
        let to = Self::reg(at)?;
        let (from, count) = self.operands(args, at, true)?;
        let f_at = u16::try_from(self.calls.len()).map_err(|_| Why::Deep)?;
        self.calls.push(f);
        if self.depth == 1 {
            self.root = Some(f_at);
        }
        self.code.push(Op::Call {
            to,
            f: f_at,
            from,
            count,
        });
        Ok(Kind::Number)
    }

    /// `INDEX` of a reference written as one, at a row and maybe a column
    /// (KERNEL.8): a call of `lookups::index_value` through its pointer, its
    /// answer guarded as any call's, or the cell's at the root. An operation
    /// of its own in the loop every program runs cost the chain fixture's
    /// cells alone half their time again (measured, 0.24 to 0.36 ms a frame),
    /// and the root's a tenth. Any other form of it is left to `Ctx`.
    fn index(&mut self, args: &[Expr], at: usize) -> Result<Kind, Why> {
        if !(2..=3).contains(&args.len()) || ref_operand(self.lib, &args[0]).is_none() {
            return Err(Why::Form);
        }
        self.called(lookups::index_value, args, at)
    }

    /// A call's operands: a reference written as one, which the call takes
    /// as the cells it names (`vector::ref_operand`'s rule), or anything
    /// whose evaluation is a value at every cell (`vector::yields_value`'s),
    /// into a register of its own above `at`. A function called through its
    /// pointer is handed a reference's text with its cells, and never an
    /// operand that may be nothing.
    fn operands(&mut self, args: &[Expr], at: usize, call: bool) -> Result<(u16, u16), Why> {
        let mut ops = Vec::with_capacity(args.len());
        for (j, a) in args.iter().enumerate() {
            let op = match ref_operand(self.lib, a) {
                Some(r) => Operand::Ref(self.refer(r, call)?),
                None if yields_value(self.lib, a) => {
                    if call && may_be_nothing(self.lib, a) {
                        return Err(Why::Omitted);
                    }
                    let r = at + 1 + j;
                    let kind = self.expr(a, r)?;
                    Operand::Reg(Self::reg(r)?, kind)
                }
                // An `IF` or a `CHOOSE` that may answer a reference.
                None => return Err(Why::Range),
            };
            ops.push(op);
        }
        let from = u16::try_from(self.operands.len()).map_err(|_| Why::Deep)?;
        let count = u16::try_from(ops.len()).map_err(|_| Why::Deep)?;
        self.operands.extend(ops);
        Ok((from, count))
    }

    /// A reference kept as its parts, its sheet the one its placement found
    /// or the cell's own, and its text where a call takes it.
    fn refer(&mut self, r: &RefExpr, text: bool) -> Result<u16, Why> {
        if r.last_sheet.is_some() {
            return Err(Why::Span);
        }
        if r.found.is_none() && r.sheet.is_some() {
            return Err(Why::Sheet);
        }
        let text = if text {
            let t = u16::try_from(self.texts.len()).map_err(|_| Why::Deep)?;
            self.texts.push(r.written.clone());
            t
        } else {
            u16::MAX
        };
        let at = u16::try_from(self.refs.len()).map_err(|_| Why::Deep)?;
        self.refs.push(Ref {
            sheet: r.found.map(|(first, _)| first),
            shape: r.shape,
            corners: r.corners,
            rect: (r.top, r.left, r.bottom, r.right),
            text,
        });
        Ok(at)
    }

    /// `IF`, as `library::if_` computes it: the condition, then the branch
    /// taken alone, an omitted branch 0 and a missing else FALSE; its two
    /// values of one kind.
    fn if_(&mut self, args: &[Expr], at: usize) -> Result<Kind, Why> {
        if !(2..=3).contains(&args.len()) {
            return Err(Why::Error);
        }
        let to = Self::reg(at)?;
        self.expr(&args[0], at)?;
        let unless = self.code.len();
        self.code.push(Op::Unless { cond: to, to: 0 });
        let then = self.branch(&args[1], at)?;
        let skip = self.code.len();
        self.code.push(Op::Jump { to: 0 });
        self.code[unless] = Op::Unless {
            cond: to,
            to: self.here()?,
        };
        let other = match args.get(2) {
            None => {
                self.code.push(Op::Const { to, x: 0.0 });
                Kind::Truth
            }
            Some(e) => self.branch(e, at)?,
        };
        self.code[skip] = Op::Jump { to: self.here()? };
        if then == other {
            Ok(then)
        } else {
            Err(Why::Kinds)
        }
    }

    /// One of `IF`'s branches or `CHOOSE`'s values into `at`: one left out
    /// is 0 (`IF`'s answer for it, and `CHOOSE`'s nothing as a number).
    fn branch(&mut self, e: &Expr, at: usize) -> Result<Kind, Why> {
        match e {
            Expr::Empty => {
                let to = Self::reg(at)?;
                self.code.push(Op::Const { to, x: 0.0 });
                Ok(Kind::Number)
            }
            _ => self.expr(e, at),
        }
    }

    /// `CHOOSE`, as `library::choose` computes it: the index, then the value
    /// chosen alone; its values of one kind.
    fn choose(&mut self, args: &[Expr], at: usize) -> Result<Kind, Why> {
        if args.len() < 2 {
            return Err(Why::Error);
        }
        let to = Self::reg(at)?;
        self.expr(&args[0], at)?;
        let from = u16::try_from(self.table.len()).map_err(|_| Why::Deep)?;
        let count = u16::try_from(args.len() - 1).map_err(|_| Why::Deep)?;
        self.table.extend(std::iter::repeat(0).take(args.len() - 1));
        self.code.push(Op::Choose {
            index: to,
            from,
            count,
        });
        let mut kind: Option<Kind> = None;
        let mut ends = Vec::with_capacity(args.len() - 1);
        for (j, a) in args[1..].iter().enumerate() {
            self.table[usize::from(from) + j] = self.here()?;
            let k = self.branch(a, at)?;
            if kind.is_some_and(|was| was != k) {
                return Err(Why::Kinds);
            }
            kind = Some(k);
            ends.push(self.code.len());
            self.code.push(Op::Jump { to: 0 });
        }
        let end = self.here()?;
        for e in ends {
            self.code[e] = Op::Jump { to: end };
        }
        Ok(kind.unwrap_or(Kind::Number))
    }
}

/// Each guard that failed, counted in a test on its own thread, so that a law
/// over programs can say that each was taken (as `vector::sliced` counts a
/// run's paths).
pub(crate) mod guarded {
    /// A read of a value no register holds: a text, a truth value, an error
    /// or a reason.
    pub(crate) const READ: usize = 0;
    /// A reference moved off the sheet, `#REF!`.
    pub(crate) const OFF: usize = 1;
    /// An operation or a fold answering an error.
    pub(crate) const ANSWER: usize = 2;
    /// A fold meeting in a range a value neither a number nor nothing.
    pub(crate) const RANGE: usize = 3;
    /// `CHOOSE`'s index outside its values, or an error.
    pub(crate) const INDEX: usize = 4;
    /// A call through the library's pointer answering anything but a
    /// number.
    pub(crate) const CALL: usize = 5;
    #[cfg(test)]
    pub(crate) const GUARDS: usize = 6;

    #[cfg(test)]
    thread_local! {
        static COUNTS: std::cell::Cell<[usize; GUARDS]> = const { std::cell::Cell::new([0; GUARDS]) };
    }

    #[cfg(test)]
    pub(crate) fn note(i: usize) {
        COUNTS.with(|c| {
            let mut a = c.get();
            a[i] += 1;
            c.set(a);
        });
    }

    /// The counts since the last time they were taken.
    #[cfg(test)]
    pub(crate) fn take() -> [usize; GUARDS] {
        COUNTS.with(|c| c.replace([0; GUARDS]))
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::calc::formula::parse;
    use crate::calc::library;

    fn twice(_: &Ctx<'_>, _: &[Arg<'_>]) -> Result<Value, Reason> {
        Ok(Value::Number(2.0))
    }

    fn first<'e>(ctx: &Ctx<'_>, args: &'e [Expr]) -> Result<Arg<'e>, Reason> {
        ctx.eval(&args[0])
    }

    /// A formula's parse lowered, unplaced, under the language's library and
    /// a function and a form of the test's own: its kind, or why it is left.
    fn lowered(text: &str) -> Result<Kind, Why> {
        let mut lib = library::language();
        lib.register("TWICE", twice);
        lib.register_form("FIRST", first);
        let e = parse(text).unwrap_or_else(|c| panic!("{text}: {c}"));
        Program::lower(&lib, &e).map(|p| p.kind)
    }

    #[test]
    fn a_parse_is_lowered_or_left_by_the_first_node_a_program_does_not_hold() {
        // KERNEL.44: each node a program holds, and its kind; each node it
        // leaves, by its reason, the first met in the order the evaluator
        // meets them; a span of sheets among them, which no grid's lowering
        // meets, a shape reading one being risky and never lowered.
        let deep = format!("{}1", "-".repeat(52));
        let shallow = format!("{}1", "-".repeat(46));
        let wide = format!("SUM({}1)", "A1+0,".repeat(40));
        let cases: Vec<(&str, Result<Kind, Why>)> = vec![
            ("1+2*3", Ok(Kind::Number)),
            ("-A1%", Ok(Kind::Number)),
            ("A1>2", Ok(Kind::Truth)),
            ("(A1>2)=(B1<1)", Ok(Kind::Truth)),
            ("AND(A1,B1:B3,1>0)", Ok(Kind::Truth)),
            ("SUM(A1:B2,3)-MIN(A1)", Ok(Kind::Number)),
            ("NOT(A1)", Ok(Kind::Truth)),
            ("MOD(ABS(A1),INT(SIN(B1)))", Ok(Kind::Number)),
            ("ROW()+COLUMN(B2:C3)", Ok(Kind::Number)),
            ("IF(A1,,2)", Ok(Kind::Number)),
            ("IF(A1,A2>1)", Ok(Kind::Truth)),
            ("CHOOSE(A1,1,,3)", Ok(Kind::Number)),
            ("TWICE(A1:A2,TRUE)", Ok(Kind::Number)),
            ("TWICE(IF(A1,,3))", Ok(Kind::Number)),
            ("SUM(CHOOSE(1,,2))", Ok(Kind::Number)),
            (&shallow, Ok(Kind::Number)),
            ("\"x\"", Err(Why::Text)),
            ("A1&B1", Err(Why::Text)),
            ("1+\"x\"", Err(Why::Text)),
            ("#N/A", Err(Why::Error)),
            ("NOT(A1,B1)", Err(Why::Error)),
            ("ABS()", Err(Why::Error)),
            ("MOD(1)", Err(Why::Error)),
            ("ROW(1)", Err(Why::Error)),
            ("ROW(A1,B1)", Err(Why::Error)),
            ("IF(A1)", Err(Why::Error)),
            ("CHOOSE(1)", Err(Why::Error)),
            ("Rate", Err(Why::Name)),
            ("A1:B2", Err(Why::Range)),
            ("A:A+1", Err(Why::Range)),
            ("2:2*3", Err(Why::Range)),
            ("SUM(IF(A1,A1:A2,1))", Err(Why::Range)),
            ("FIRST(1)", Err(Why::Form)),
            ("A1=TRUE", Err(Why::Kinds)),
            ("IF(A1,1,TRUE)", Err(Why::Kinds)),
            ("IF(A1,1)", Err(Why::Kinds)),
            ("CHOOSE(1,2,TRUE)", Err(Why::Kinds)),
            ("Jan:Dec!A1", Err(Why::Span)),
            ("SUM(Jan:Dec!A1)", Err(Why::Span)),
            ("ROW(Jan:Dec!A1)", Err(Why::Span)),
            ("Data!A1", Err(Why::Sheet)),
            ("SUM(Data!A1:A2)", Err(Why::Sheet)),
            ("TWICE(CHOOSE(1,,2))", Err(Why::Omitted)),
            ("TWICE(IF(A1,CHOOSE(2,1,),3))", Err(Why::Omitted)),
            ("TWICE(+CHOOSE(2,1,))", Err(Why::Omitted)),
            ("TWICE(,1)", Err(Why::Omitted)),
            ("SUM(1,,2)", Err(Why::Omitted)),
            (&deep, Err(Why::Deep)),
            (&wide, Err(Why::Deep)),
            // The first met: a text before a name, a name before a text.
            ("\"x\"+Rate", Err(Why::Text)),
            ("Rate+\"x\"", Err(Why::Name)),
        ];
        for (text, want) in cases {
            assert_eq!(lowered(text), want, "{text}");
        }
    }

    #[test]
    fn each_reason_is_spelled_as_a_census_row_spells_it() {
        let all = [
            (Why::Text, "text"),
            (Why::Error, "error"),
            (Why::Name, "name"),
            (Why::Range, "range"),
            (Why::Form, "form"),
            (Why::Kinds, "kinds"),
            (Why::Span, "span"),
            (Why::Sheet, "sheet"),
            (Why::Omitted, "omitted"),
            (Why::Deep, "deep"),
        ];
        for (why, word) in all {
            assert_eq!(why.word(), word);
        }
    }
}
