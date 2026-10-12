//! The language's own functions (KERNEL.7, 2026-10-08): what a cartridge
//! rule needs on day one, `Alonzo/SPEC.md` section 13's list to this item.
//!
//! `IF`, `AND`, `OR`, `NOT`, `SUM`, `MIN`, `MAX`, `ABS`, `INT`, `MOD`,
//! `ROW`, `COLUMN`, `CHOOSE` and `SIN`, with the comparisons and the
//! arithmetic the evaluator holds as operators: `ROW` and `COLUMN` because
//! a cell painting at the mouse must know where it is, `MOD` because every
//! counter wraps, `CHOOSE` for a score until `INDEX` comes, `SIN` for a
//! slow oscillator. Each is the formula language's own, with the semantics
//! ISO/IEC 29500 and OpenFormula publish and Excel computes, held here by
//! tests with exact values and no tolerance (`Alonzo/docs/ALGEBRA.md`
//! section 7, crack 1); an engine's replay golden is the oracle above
//! them. Excel's wider library registers through the same seam from
//! `frazaro-core`, by measurement, and a function the engine comes to
//! need moves down here with its tests.
//!
//! The rule every aggregate follows, Excel's: a value given directly is
//! coerced (a truth value is 1 or 0, a numeric text its number, another
//! text `#VALUE!`), a value read out of a range counts only when it is of
//! the kind the function wants (`SUM` skips a text and a truth value in a
//! range), an empty cell is skipped, and an error anywhere is the answer.
//!
//! The rule is written once, for numbers and for truth values ([`Rule`]),
//! and so is each aggregate's fold ([`Fold`]: where `SUM`, `MIN`, `MAX`,
//! `AND` and `OR` start, what each value they count does, what they answer),
//! and `NOT`'s answer ([`negated`]); a cell's evaluation drives them over its
//! arguments here, and an evaluation of many cells at once drives the same
//! rule and the same fold over its own (`super::vector`, KERNEL.26), so that
//! no meaning of the library is written twice.
//!
//! So is the meaning of each of the others (KERNEL.28): `ABS`, `INT`, `SIN`
//! and `MOD` over the numbers their arguments read as ([`abs_of`],
//! [`int_of`], [`sin_of`], [`mod_of`]), `ROW` and `COLUMN` over the cell or
//! the reference they place ([`placed`]), and `CHOOSE`'s index ([`chosen`]).
//! The functions below call them after their own coercions, and an
//! evaluation of many cells at once calls them over its lanes, a function of
//! numbers where every operand at a cell is a number and the function below
//! at any other cell, so that a coercion is written here alone.

use super::eval::{finite, Arg, Ctx, Known, Library, Reason};
use super::formula::Expr;
use super::value::{ErrorKind, Value};

/// The language's library, every function above registered, each as one an
/// evaluation of many cells at once computes by itself: `IF` as the form it
/// computes as `if_` computes it at one (KERNEL.20), the aggregates and `NOT`
/// by the same rule and fold (KERNEL.26), and the others, `CHOOSE` among
/// them as a form, by the same meanings (KERNEL.28).
pub fn language() -> Library {
    let mut lib = Library::empty();
    lib.register_known("IF", if_, Known::If);
    lib.register_known_strict("AND", and, Known::And);
    lib.register_known_strict("OR", or, Known::Or);
    lib.register_known_strict("NOT", not, Known::Not);
    lib.register_known_strict("SUM", sum, Known::Sum);
    lib.register_known_strict("MIN", min, Known::Min);
    lib.register_known_strict("MAX", max, Known::Max);
    lib.register_known_strict("ABS", abs, Known::Abs);
    lib.register_known_strict("INT", int, Known::Int);
    lib.register_known_strict("MOD", modulo, Known::Mod);
    lib.register_known_strict("ROW", row, Known::Row);
    lib.register_known_strict("COLUMN", column, Known::Column);
    lib.register_known("CHOOSE", choose, Known::Choose);
    lib.register_known_strict("SIN", sin, Known::Sin);
    lib
}

/// The names the language's library holds, for a page to list.
pub const NAMES: [&str; 14] = [
    "IF", "AND", "OR", "NOT", "SUM", "MIN", "MAX", "ABS", "INT", "MOD", "ROW", "COLUMN", "CHOOSE",
    "SIN",
];

/// What a numeric aggregate collects: the numbers, or the first error.
pub type Numbers = Result<Vec<f64>, ErrorKind>;

/// What an aggregate makes of a value read out of a range (KERNEL.26): an
/// item it counts, an error it keeps to answer with once the range is
/// walked, or nothing.
#[derive(Clone, Copy, Debug, PartialEq)]
pub(crate) enum Ranged<T> {
    Counts(T),
    Error(ErrorKind),
    Skipped,
}

/// The aggregate rule above for one kind of item (KERNEL.26): how a value
/// given directly is coerced, and what a value read out of a range is.
pub(crate) trait Rule {
    type Item: Copy;
    /// A value given directly: its item, or the error that is the answer.
    fn given(v: &Value) -> Result<Self::Item, ErrorKind>;
    /// A value read out of a range.
    fn ranged(v: &Value) -> Ranged<Self::Item>;
}

/// The rule for numbers: a value given directly reads as a number or is
/// `#VALUE!`; in a range only a number counts.
pub(crate) struct OverNumbers;

impl Rule for OverNumbers {
    type Item = f64;

    fn given(v: &Value) -> Result<f64, ErrorKind> {
        v.number()
    }

    fn ranged(v: &Value) -> Ranged<f64> {
        match v {
            Value::Number(n) => Ranged::Counts(*n),
            Value::Error(k) => Ranged::Error(*k),
            _ => Ranged::Skipped,
        }
    }
}

/// The rule for truth values: a value given directly must read as one
/// (`#VALUE!` otherwise); in a range a truth value counts, a number counts
/// as `0` or not, a text is skipped.
pub(crate) struct OverTruths;

impl Rule for OverTruths {
    type Item = bool;

    fn given(v: &Value) -> Result<bool, ErrorKind> {
        v.truth()
    }

    fn ranged(v: &Value) -> Ranged<bool> {
        match v {
            Value::Bool(b) => Ranged::Counts(*b),
            Value::Number(n) => Ranged::Counts(*n != 0.0),
            Value::Error(k) => Ranged::Error(*k),
            _ => Ranged::Skipped,
        }
    }
}

/// Each item of the arguments by a rule, to `f` in the arguments' order and
/// a range's walk: a value given directly that is an error, or does not read
/// as an item, answers at once; a range is walked to its end and answers
/// with its first error, so that a cell in it not computed still answers
/// with its reason. An aggregate folds over this as it walks, and nothing is
/// collected (KERNEL.25).
fn each_item<R: Rule, F: FnMut(R::Item)>(
    ctx: &Ctx<'_>,
    args: &[Arg],
    mut f: F,
) -> Result<Result<(), ErrorKind>, Reason> {
    for a in args {
        match a {
            Arg::Scalar(v) => match R::given(v) {
                Ok(x) => f(x),
                Err(k) => return Ok(Err(k)),
            },
            Arg::Area { areas, .. } => {
                let mut err: Option<ErrorKind> = None;
                ctx.each(areas, |v| {
                    match R::ranged(&v) {
                        Ranged::Counts(x) => f(x),
                        Ranged::Error(k) => {
                            err.get_or_insert(k);
                        }
                        Ranged::Skipped => {}
                    }
                    Ok(())
                })?;
                if let Some(k) = err {
                    return Ok(Err(k));
                }
            }
        }
    }
    Ok(Ok(()))
}

/// Each number of the arguments by the aggregate rule above, to `f`, as
/// [`each_item`] walks them ([`OverNumbers`]).
pub fn each_number<F: FnMut(f64)>(
    ctx: &Ctx<'_>,
    args: &[Arg],
    f: F,
) -> Result<Result<(), ErrorKind>, Reason> {
    each_item::<OverNumbers, F>(ctx, args, f)
}

/// Each truth value of the arguments by the same rule, to `f`, as
/// [`each_number`] walks ([`OverTruths`]).
pub fn each_truth<F: FnMut(bool)>(
    ctx: &Ctx<'_>,
    args: &[Arg],
    f: F,
) -> Result<Result<(), ErrorKind>, Reason> {
    each_item::<OverTruths, F>(ctx, args, f)
}

/// An aggregate's meaning (KERNEL.26): its rule, where it starts, what each
/// item it counts does, and its answer once its arguments are walked. A
/// cell's evaluation drives it through [`folded`]; an evaluation of many
/// cells at once drives it through its own walk (`super::vector`).
pub(crate) trait Fold: Default {
    type Rule: Rule;
    fn add(&mut self, x: <Self::Rule as Rule>::Item);
    fn answer(self) -> Value;
}

/// `SUM` adds in the walk's order from -0.0, `Iterator::sum`'s own start,
/// so that its bits are those of the numbers collected and summed.
pub(crate) struct Sum(f64);

impl Default for Sum {
    fn default() -> Sum {
        Sum(-0.0)
    }
}

impl Fold for Sum {
    type Rule = OverNumbers;

    fn add(&mut self, n: f64) {
        self.0 += n;
    }

    fn answer(self) -> Value {
        finite(self.0)
    }
}

/// `MIN` of no number is 0.
pub(crate) struct Least {
    least: f64,
    any: bool,
}

impl Default for Least {
    fn default() -> Least {
        Least {
            least: f64::INFINITY,
            any: false,
        }
    }
}

impl Fold for Least {
    type Rule = OverNumbers;

    fn add(&mut self, n: f64) {
        self.least = self.least.min(n);
        self.any = true;
    }

    fn answer(self) -> Value {
        Value::Number(self.least.min(if self.any { f64::INFINITY } else { 0.0 }))
    }
}

/// `MAX` of no number is 0.
pub(crate) struct Most {
    most: f64,
    any: bool,
}

impl Default for Most {
    fn default() -> Most {
        Most {
            most: f64::NEG_INFINITY,
            any: false,
        }
    }
}

impl Fold for Most {
    type Rule = OverNumbers;

    fn add(&mut self, n: f64) {
        self.most = self.most.max(n);
        self.any = true;
    }

    fn answer(self) -> Value {
        Value::Number(
            self.most
                .max(if self.any { f64::NEG_INFINITY } else { 0.0 }),
        )
    }
}

/// `AND` of no truth value is `#VALUE!`.
pub(crate) struct All {
    all: bool,
    any: bool,
}

impl Default for All {
    fn default() -> All {
        All {
            all: true,
            any: false,
        }
    }
}

impl Fold for All {
    type Rule = OverTruths;

    fn add(&mut self, b: bool) {
        self.all &= b;
        self.any = true;
    }

    fn answer(self) -> Value {
        if self.any {
            Value::Bool(self.all)
        } else {
            Value::Error(ErrorKind::Value)
        }
    }
}

/// `OR` of no truth value is `#VALUE!`.
#[derive(Default)]
pub(crate) struct Some_ {
    some: bool,
    any: bool,
}

impl Fold for Some_ {
    type Rule = OverTruths;

    fn add(&mut self, b: bool) {
        self.some |= b;
        self.any = true;
    }

    fn answer(self) -> Value {
        if self.any {
            Value::Bool(self.some)
        } else {
            Value::Error(ErrorKind::Value)
        }
    }
}

/// An aggregate over a cell's arguments: its fold driven by its rule's walk.
fn folded<F: Fold>(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    let mut fold = F::default();
    Ok(match each_item::<F::Rule, _>(ctx, args, |x| fold.add(x))? {
        Err(k) => Value::Error(k),
        Ok(()) => fold.answer(),
    })
}

/// `NOT`'s answer over the one value it reads (KERNEL.26).
pub(crate) fn negated(v: &Value) -> Value {
    match v.truth() {
        Ok(b) => Value::Bool(!b),
        Err(k) => Value::Error(k),
    }
}

/// `ABS`'s answer over the number its argument reads as (KERNEL.28).
pub(crate) fn abs_of(n: f64) -> Value {
    Value::Number(n.abs())
}

/// `INT`'s, rounding down to the nearest integer, so that `INT(-1.5)` is -2.
pub(crate) fn int_of(n: f64) -> Value {
    Value::Number(n.floor())
}

/// `SIN`'s, of an angle in radians.
pub(crate) fn sin_of(n: f64) -> Value {
    finite(n.sin())
}

/// `MOD(n, d)`'s: `n - d * INT(n / d)`, so that the result takes the
/// divisor's sign; a divisor of 0 is `#DIV/0!`.
pub(crate) fn mod_of(n: f64, d: f64) -> Value {
    if d == 0.0 {
        return Value::Error(ErrorKind::Div0);
    }
    finite(n - d * (n / d).floor())
}

/// Which of a cell's coordinates `ROW` and `COLUMN` answer (KERNEL.28).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum Axis {
    Row,
    Column,
}

impl Axis {
    /// A cell's row or column, as a number.
    pub(crate) fn of(self, row: u32, column: u32) -> f64 {
        f64::from(match self {
            Axis::Row => row,
            Axis::Column => column,
        })
    }
}

/// `ROW`'s and `COLUMN`'s answer (KERNEL.28): the row or the column of the
/// cell its argument places, the cell being computed or a reference's first
/// cell; `#VALUE!` for an argument that places none, a value or a reference
/// to no cell (a sheet the grid does not hold, a cell moved off the sheet),
/// which reaches it as the value `#REF!`.
pub(crate) fn placed(axis: Axis, cell: Option<(u32, u32)>) -> Value {
    match cell {
        Some((row, column)) => Value::Number(axis.of(row, column)),
        None => Value::Error(ErrorKind::Value),
    }
}

/// `CHOOSE`'s index (KERNEL.28): which of its `count` values it picks, the
/// index read down to a whole number; an error the index's own, and
/// `#VALUE!` outside 1 to the count.
pub(crate) fn chosen(index: &Value, count: usize) -> Result<usize, ErrorKind> {
    let k = index.number()?.floor();
    if k < 1.0 || k > count as f64 {
        return Err(ErrorKind::Value);
    }
    Ok(k as usize)
}

/// The numbers of the arguments by the aggregate rule above, collected.
pub fn numbers(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Numbers, Reason> {
    let mut out = Vec::new();
    let walked = each_number(ctx, args, |n| out.push(n))?;
    Ok(walked.map(|()| out))
}

/// The truth values of the arguments by the same rule, collected.
pub fn truths(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Result<Vec<bool>, ErrorKind>, Reason> {
    let mut out = Vec::new();
    let walked = each_truth(ctx, args, |b| out.push(b))?;
    Ok(walked.map(|()| out))
}

/// One numeric argument, coerced as an operand is.
pub fn number_arg(
    ctx: &Ctx<'_>,
    args: &[Arg],
    at: usize,
) -> Result<Result<f64, ErrorKind>, Reason> {
    let Some(arg) = args.get(at) else {
        return Ok(Err(ErrorKind::Value));
    };
    Ok(ctx.one(arg)?.number())
}

/// `IF`, which an evaluation of many cells at once computes by itself
/// (`super::vector`, its `if_`): a change here is a change there, and the
/// machine's tests hold the two to each other (KERNEL.20).
fn if_<'e>(ctx: &Ctx<'_>, args: &'e [Expr]) -> Result<Arg<'e>, Reason> {
    if args.len() < 2 || args.len() > 3 {
        return Ok(Arg::Scalar(Value::Error(ErrorKind::Value)));
    }
    let condition = match ctx.truth(&args[0])? {
        Ok(b) => b,
        Err(k) => return Ok(Arg::Scalar(Value::Error(k))),
    };
    // A branch left out reads as 0 when taken; a missing else is FALSE.
    match (condition, args.get(1), args.get(2)) {
        (true, Some(Expr::Empty), _) | (false, _, Some(Expr::Empty)) => {
            Ok(Arg::Scalar(Value::Number(0.0)))
        }
        (true, Some(then), _) => ctx.eval(then),
        (false, _, Some(otherwise)) => ctx.eval(otherwise),
        (false, _, None) => Ok(Arg::Scalar(Value::Bool(false))),
        (true, None, _) => Ok(Arg::Scalar(Value::Error(ErrorKind::Value))),
    }
}

fn and(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    folded::<All>(ctx, args)
}

fn or(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    folded::<Some_>(ctx, args)
}

fn not(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    if args.len() != 1 {
        return Ok(Value::Error(ErrorKind::Value));
    }
    Ok(negated(&ctx.one(&args[0])?))
}

fn sum(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    folded::<Sum>(ctx, args)
}

fn min(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    folded::<Least>(ctx, args)
}

fn max(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    folded::<Most>(ctx, args)
}

/// `ABS` ([`abs_of`]).
fn abs(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    if args.len() != 1 {
        return Ok(Value::Error(ErrorKind::Value));
    }
    Ok(match number_arg(ctx, args, 0)? {
        Ok(n) => abs_of(n),
        Err(k) => Value::Error(k),
    })
}

/// `INT` ([`int_of`]).
fn int(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    if args.len() != 1 {
        return Ok(Value::Error(ErrorKind::Value));
    }
    Ok(match number_arg(ctx, args, 0)? {
        Ok(n) => int_of(n),
        Err(k) => Value::Error(k),
    })
}

/// `MOD` ([`mod_of`]), the first argument's error before the second's.
fn modulo(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    if args.len() != 2 {
        return Ok(Value::Error(ErrorKind::Value));
    }
    let n = match number_arg(ctx, args, 0)? {
        Ok(n) => n,
        Err(k) => return Ok(Value::Error(k)),
    };
    let d = match number_arg(ctx, args, 1)? {
        Ok(d) => d,
        Err(k) => return Ok(Value::Error(k)),
    };
    Ok(mod_of(n, d))
}

/// The row of the first cell of a reference, or of the cell being computed.
fn row(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    position(ctx, args, Axis::Row)
}

/// The column of the first cell of a reference, or of the cell being computed.
fn column(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    position(ctx, args, Axis::Column)
}

/// [`placed`] over a call's arguments: none, the cell being computed; a
/// reference, its first area's first cell; anything else, no cell.
fn position(ctx: &Ctx<'_>, args: &[Arg], axis: Axis) -> Result<Value, Reason> {
    let (_, row, column) = ctx.env.here();
    Ok(match args {
        [] => placed(axis, Some((row, column))),
        [Arg::Area { areas, .. }] => match areas.first() {
            Some(a) => placed(axis, Some((a.top, a.left))),
            None => Value::Error(ErrorKind::Ref),
        },
        _ => placed(axis, None),
    })
}

/// `CHOOSE(k, v1, v2, ...)`: the k-th value ([`chosen`]). Only the value
/// chosen is computed, and a reference chosen stays a reference. An
/// evaluation of many cells at once computes each value at the cells that
/// choose it (`super::vector`, its `choose`, KERNEL.28): a change here is a
/// change there, and the machine's tests hold the two to each other.
fn choose<'e>(ctx: &Ctx<'_>, args: &'e [Expr]) -> Result<Arg<'e>, Reason> {
    if args.len() < 2 {
        return Ok(Arg::Scalar(Value::Error(ErrorKind::Value)));
    }
    match chosen(&ctx.scalar(&args[0])?, args.len() - 1) {
        Ok(k) => ctx.eval(&args[k]),
        Err(e) => Ok(Arg::Scalar(Value::Error(e))),
    }
}

/// `SIN` ([`sin_of`]).
fn sin(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    if args.len() != 1 {
        return Ok(Value::Error(ErrorKind::Value));
    }
    Ok(match number_arg(ctx, args, 0)? {
        Ok(n) => sin_of(n),
        Err(k) => Value::Error(k),
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn the_language_registers_the_day_one_list_exactly() {
        let lib = language();
        let mut want: Vec<&str> = NAMES.to_vec();
        want.sort_unstable();
        assert_eq!(lib.names(), want);
        assert!(
            lib.form("IF").is_some(),
            "IF computes only the branch it takes"
        );
        assert!(
            lib.form("CHOOSE").is_some(),
            "CHOOSE computes only the value chosen"
        );
        assert!(lib.strict("SUM").is_some());
    }
}
