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

use super::eval::{finite, Arg, Ctx, Library, Reason};
use super::formula::Expr;
use super::value::{ErrorKind, Value};

/// The language's library, every function above registered.
pub fn language() -> Library {
    let mut lib = Library::empty();
    lib.register_form("IF", if_);
    lib.register("AND", and);
    lib.register("OR", or);
    lib.register("NOT", not);
    lib.register("SUM", sum);
    lib.register("MIN", min);
    lib.register("MAX", max);
    lib.register("ABS", abs);
    lib.register("INT", int);
    lib.register("MOD", modulo);
    lib.register("ROW", row);
    lib.register("COLUMN", column);
    lib.register_form("CHOOSE", choose);
    lib.register("SIN", sin);
    lib
}

/// The names the language's library holds, for a page to list.
pub const NAMES: [&str; 14] = [
    "IF", "AND", "OR", "NOT", "SUM", "MIN", "MAX", "ABS", "INT", "MOD", "ROW", "COLUMN", "CHOOSE",
    "SIN",
];

/// What a numeric aggregate collects: the numbers, or the first error.
pub type Numbers = Result<Vec<f64>, ErrorKind>;

/// The numbers of the arguments by the aggregate rule above.
pub fn numbers(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Numbers, Reason> {
    let mut out = Vec::new();
    for a in args {
        match a {
            Arg::Scalar(v) => match v.number() {
                Ok(n) => out.push(n),
                Err(k) => return Ok(Err(k)),
            },
            Arg::Area { areas, .. } => {
                let mut err: Option<ErrorKind> = None;
                ctx.each(areas, &mut |v| {
                    match v {
                        Value::Number(n) => out.push(n),
                        Value::Error(k) => {
                            err.get_or_insert(k);
                        }
                        _ => {}
                    }
                    Ok(())
                })?;
                if let Some(k) = err {
                    return Ok(Err(k));
                }
            }
        }
    }
    Ok(Ok(out))
}

/// The truth values of the arguments by the same rule: a value given
/// directly must read as one (`#VALUE!` otherwise); in a range a truth
/// value counts, a number counts as `0` or not, a text is skipped.
pub fn truths(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Result<Vec<bool>, ErrorKind>, Reason> {
    let mut out = Vec::new();
    for a in args {
        match a {
            Arg::Scalar(v) => match v.truth() {
                Ok(b) => out.push(b),
                Err(k) => return Ok(Err(k)),
            },
            Arg::Area { areas, .. } => {
                let mut err: Option<ErrorKind> = None;
                ctx.each(areas, &mut |v| {
                    match v {
                        Value::Bool(b) => out.push(b),
                        Value::Number(n) => out.push(n != 0.0),
                        Value::Error(k) => {
                            err.get_or_insert(k);
                        }
                        _ => {}
                    }
                    Ok(())
                })?;
                if let Some(k) = err {
                    return Ok(Err(k));
                }
            }
        }
    }
    Ok(Ok(out))
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

fn if_(ctx: &Ctx<'_>, args: &[Expr]) -> Result<Arg, Reason> {
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
    Ok(match truths(ctx, args)? {
        Err(k) => Value::Error(k),
        Ok(bs) if bs.is_empty() => Value::Error(ErrorKind::Value),
        Ok(bs) => Value::Bool(bs.iter().all(|b| *b)),
    })
}

fn or(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    Ok(match truths(ctx, args)? {
        Err(k) => Value::Error(k),
        Ok(bs) if bs.is_empty() => Value::Error(ErrorKind::Value),
        Ok(bs) => Value::Bool(bs.iter().any(|b| *b)),
    })
}

fn not(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    if args.len() != 1 {
        return Ok(Value::Error(ErrorKind::Value));
    }
    Ok(match ctx.truth_arg(&args[0])? {
        Ok(b) => Value::Bool(!b),
        Err(k) => Value::Error(k),
    })
}

fn sum(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    Ok(match numbers(ctx, args)? {
        Err(k) => Value::Error(k),
        Ok(ns) => finite(ns.iter().sum()),
    })
}

fn min(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    Ok(match numbers(ctx, args)? {
        Err(k) => Value::Error(k),
        Ok(ns) => Value::Number(
            ns.iter()
                .copied()
                .fold(f64::INFINITY, f64::min)
                .min(if ns.is_empty() { 0.0 } else { f64::INFINITY }),
        ),
    })
}

fn max(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    Ok(match numbers(ctx, args)? {
        Err(k) => Value::Error(k),
        Ok(ns) => Value::Number(ns.iter().copied().fold(f64::NEG_INFINITY, f64::max).max(
            if ns.is_empty() {
                0.0
            } else {
                f64::NEG_INFINITY
            },
        )),
    })
}

fn abs(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    if args.len() != 1 {
        return Ok(Value::Error(ErrorKind::Value));
    }
    Ok(match number_arg(ctx, args, 0)? {
        Ok(n) => Value::Number(n.abs()),
        Err(k) => Value::Error(k),
    })
}

/// `INT` rounds down to the nearest integer, so `INT(-1.5)` is -2.
fn int(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    if args.len() != 1 {
        return Ok(Value::Error(ErrorKind::Value));
    }
    Ok(match number_arg(ctx, args, 0)? {
        Ok(n) => Value::Number(n.floor()),
        Err(k) => Value::Error(k),
    })
}

/// `MOD(n, d)` is `n - d * INT(n / d)`, so the result takes the divisor's
/// sign; a divisor of 0 is `#DIV/0!`.
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
    if d == 0.0 {
        return Ok(Value::Error(ErrorKind::Div0));
    }
    Ok(finite(n - d * (n / d).floor()))
}

/// The row of the first cell of a reference, or of the cell being computed.
fn row(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    position(ctx, args, |here, area| match area {
        Some(a) => a.top,
        None => here.1,
    })
}

/// The column of the first cell of a reference, or of the cell being computed.
fn column(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    position(ctx, args, |here, area| match area {
        Some(a) => a.left,
        None => here.2,
    })
}

fn position(
    ctx: &Ctx<'_>,
    args: &[Arg],
    pick: fn((usize, u32, u32), Option<&super::graph::Area>) -> u32,
) -> Result<Value, Reason> {
    let here = ctx.env.here();
    match args {
        [] => Ok(Value::Number(f64::from(pick(here, None)))),
        [Arg::Area { areas, .. }] => match areas.first() {
            Some(a) => Ok(Value::Number(f64::from(pick(here, Some(a))))),
            None => Ok(Value::Error(ErrorKind::Ref)),
        },
        _ => Ok(Value::Error(ErrorKind::Value)),
    }
}

/// `CHOOSE(k, v1, v2, ...)`: the k-th value, k read down to a whole
/// number; outside 1 to the count, `#VALUE!`. Only the value chosen is
/// computed, and a reference chosen stays a reference.
fn choose(ctx: &Ctx<'_>, args: &[Expr]) -> Result<Arg, Reason> {
    if args.len() < 2 {
        return Ok(Arg::Scalar(Value::Error(ErrorKind::Value)));
    }
    let k = match ctx.scalar(&args[0])?.number() {
        Ok(n) => n.floor(),
        Err(e) => return Ok(Arg::Scalar(Value::Error(e))),
    };
    if k < 1.0 || k > (args.len() - 1) as f64 {
        return Ok(Arg::Scalar(Value::Error(ErrorKind::Value)));
    }
    ctx.eval(&args[k as usize])
}

fn sin(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    if args.len() != 1 {
        return Ok(Value::Error(ErrorKind::Value));
    }
    Ok(match number_arg(ctx, args, 0)? {
        Ok(n) => finite(n.sin()),
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
