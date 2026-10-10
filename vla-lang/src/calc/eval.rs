//! The evaluator and the functions seam (KERNEL.7, 2026-10-08).
//!
//! One formula at a time: an [`Expr`] against an [`Env`], the grid as the
//! cell being computed sees it, through a [`Ctx`] that carries the
//! [`Library`] of functions. The library is the seam: a function is a
//! Rust function registered under its upper-case name, strict (its
//! arguments evaluated first, as `SUM`'s are) or a form (its arguments
//! handed over unevaluated, as `IF`'s are, so that only the branch taken
//! is computed). The language registers its own functions
//! (`super::library`), the ones a cartridge rule needs on day one; a crate
//! above registers more through the same calls, `frazaro-core` Excel's
//! library by measurement, and an engine consumes the language's alone.
//!
//! What an operand is: a [`Value`], or an area, a rectangle of cells a
//! reference names, which a function such as `SUM` walks and which a
//! scalar operator reads as one cell when it is one cell. A range where
//! one value is wanted is not computed in this slice ([`Reason`]), as the
//! parser's page says. An error value is a value and propagates as Excel
//! propagates it, the left operand's first; it never refuses anything.

use std::borrow::Cow;
use std::cell::Cell as Counter;
use std::cmp::Ordering;
use std::collections::BTreeMap;

use crate::intrinsics::fold;
use crate::rows::quoted;
use crate::sheet::Workbook;

use super::formula::{self, BinOp, Expr, RefExpr};
use super::graph::{areas_at, cells_in, Area, Areas, CellId, Extent};
use super::value::{ErrorKind, Value};

/// Why a cell was not computed: the honest static label, named.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum Reason {
    /// A function the library does not hold, by name.
    Function(String),
    /// A construct this slice does not read, as written.
    Construct(String),
    /// The cell is in a cycle, or reads one.
    Cycle,
}

impl Reason {
    /// The label as a datum: `(not-computed "VLOOKUP")`, `(not-computed cycle)`.
    pub fn spell(&self) -> String {
        match self {
            Reason::Function(name) | Reason::Construct(name) => {
                format!("(not-computed {})", quoted(name))
            }
            Reason::Cycle => "(not-computed cycle)".to_string(),
        }
    }
}

/// What a formula cell holds after evaluation.
#[derive(Clone, Debug, PartialEq)]
pub enum Computed {
    Value(Value),
    NotComputed(Reason),
}

impl Computed {
    /// The datum a `value` or `calc` row prints.
    pub fn spell(&self) -> String {
        match self {
            Computed::Value(v) => v.spell(),
            Computed::NotComputed(r) => r.spell(),
        }
    }
}

/// An evaluated operand: one value, or the cells a reference names. `'e` is
/// the life of the expression it was evaluated from, whose text labels an
/// area (KERNEL.25).
#[derive(Clone, Debug, PartialEq)]
pub enum Arg<'e> {
    Scalar(Value),
    Area {
        areas: Areas,
        /// The reference as written, for a label: borrowed from the parse,
        /// so that a reference evaluated costs no copy of its text, or owned
        /// when the parse was the evaluation's own, a defined name's text.
        written: Cow<'e, str>,
        /// The offset the reference was evaluated at (KERNEL.22, a shape at
        /// a cell other than its first), by which the label's text is moved
        /// before it is shown; (0, 0) everywhere else.
        moved: (i64, i64),
    },
}

impl Arg<'_> {
    /// The operand with its label owned, to outlive the parse it was
    /// evaluated from.
    pub fn into_owned(self) -> Arg<'static> {
        match self {
            Arg::Scalar(v) => Arg::Scalar(v),
            Arg::Area {
                areas,
                written,
                moved,
            } => Arg::Area {
                areas,
                written: Cow::Owned(written.into_owned()),
                moved,
            },
        }
    }
}

/// The grid as the cell being computed sees it.
pub trait Env {
    /// The model.
    fn model(&self) -> &Workbook;
    /// The cell being computed.
    fn here(&self) -> CellId;
    /// A cell's value now: a constant as it stands, a formula's computed
    /// value, `Empty` for a cell with nothing; a formula not computed is
    /// the reason, which propagates.
    fn read(&self, id: CellId) -> Result<Value, Reason>;
    /// A sheet's used rectangle, cached by the caller.
    fn extent(&self, sheet: usize) -> Extent;
    /// The text of a defined name, resolved as `graph::Names::resolve`
    /// resolves it.
    fn name(&self, name: &str, scope: Option<usize>, home: usize) -> Option<&str>;
}

/// A strict function: its arguments evaluated, a value or a reason out.
pub type Strict = fn(&Ctx<'_>, &[Arg<'_>]) -> Result<Value, Reason>;

/// A form: its arguments as written, an operand or a reason out, the operand
/// living as long as the arguments it was evaluated from.
pub type Form = for<'e> fn(&Ctx<'_>, &'e [Expr]) -> Result<Arg<'e>, Reason>;

/// The functions a formula may call, by upper-case name: the seam.
#[derive(Default)]
pub struct Library {
    strict: BTreeMap<String, Strict>,
    forms: BTreeMap<String, Form>,
}

impl Library {
    /// A library with nothing in it.
    pub fn empty() -> Library {
        Library::default()
    }

    /// Register a strict function; a later registration of the same name
    /// replaces the earlier.
    pub fn register(&mut self, name: &str, f: Strict) {
        let key = name.to_ascii_uppercase();
        self.forms.remove(&key);
        self.strict.insert(key, f);
    }

    /// Register a form.
    pub fn register_form(&mut self, name: &str, f: Form) {
        let key = name.to_ascii_uppercase();
        self.strict.remove(&key);
        self.forms.insert(key, f);
    }

    /// The strict function of that name, without case.
    pub fn strict(&self, name: &str) -> Option<Strict> {
        self.strict.get(&*key(name)).copied()
    }

    /// The form of that name, without case.
    pub fn form(&self, name: &str) -> Option<Form> {
        self.forms.get(&*key(name)).copied()
    }

    /// Whether the library holds a function of that name.
    pub fn has(&self, name: &str) -> bool {
        let wanted = key(name);
        self.strict.contains_key(&*wanted) || self.forms.contains_key(&*wanted)
    }

    /// Every name the library holds, sorted.
    pub fn names(&self) -> Vec<&str> {
        let mut names: Vec<&str> = self
            .strict
            .keys()
            .chain(self.forms.keys())
            .map(String::as_str)
            .collect();
        names.sort_unstable();
        names
    }

    /// How many functions the library holds.
    pub fn len(&self) -> usize {
        self.strict.len() + self.forms.len()
    }

    /// Whether the library holds none.
    pub fn is_empty(&self) -> bool {
        self.len() == 0
    }
}

/// A name as the library keys it, in ASCII upper case: the name as it
/// stands when it holds no ASCII small letter, as a parsed call's name
/// never does (`formula.rs` upper-cases it), so that a lookup at every
/// evaluation makes no string (KERNEL.25).
fn key(name: &str) -> Cow<'_, str> {
    if name.bytes().any(|b| b.is_ascii_lowercase()) {
        Cow::Owned(name.to_ascii_uppercase())
    } else {
        Cow::Borrowed(name)
    }
}

/// How deep a name may read a name before `#NAME?`.
const NAME_DEPTH: u32 = 16;

/// One cell's evaluation: the environment and the library.
pub struct Ctx<'a> {
    pub env: &'a dyn Env,
    pub lib: &'a Library,
    depth: Counter<u32>,
    /// The offset of the cell being computed from the cell its formula was
    /// read at: (0, 0) but in an engine's machine, which parses a shape once
    /// and evaluates it at each of its cells (KERNEL.22).
    offset: Counter<(i64, i64)>,
}

impl<'a> Ctx<'a> {
    pub fn new(env: &'a dyn Env, lib: &'a Library) -> Ctx<'a> {
        Ctx::at(env, lib, 0, 0)
    }

    /// An evaluation of a formula read at another cell, `d_row` rows up and
    /// `d_col` columns left of the cell being computed: every reference the
    /// formula writes is moved by that much, as a filled formula's are, and
    /// the references inside a defined name's text are not, since a name
    /// does not move with the cell that reads it (KERNEL.22).
    pub fn at(env: &'a dyn Env, lib: &'a Library, d_row: i64, d_col: i64) -> Ctx<'a> {
        Ctx {
            env,
            lib,
            depth: Counter::new(0),
            offset: Counter::new((d_row, d_col)),
        }
    }

    /// An expression as an operand.
    pub fn eval<'e>(&self, e: &'e Expr) -> Result<Arg<'e>, Reason> {
        match e {
            Expr::Number(v) => Ok(Arg::Scalar(Value::Number(*v))),
            Expr::Text(t) => Ok(Arg::Scalar(Value::Text(t.clone()))),
            Expr::Bool(b) => Ok(Arg::Scalar(Value::Bool(*b))),
            Expr::Error(k) => Ok(Arg::Scalar(Value::Error(*k))),
            Expr::Empty => Ok(Arg::Scalar(Value::Empty)),
            Expr::Ref(r) => self.reference(r),
            Expr::Name { name, sheet } => self.named(name, sheet.as_deref()),
            Expr::Call(name, args) => self.call(name, args),
            Expr::Neg(x) => {
                let v = self.scalar(x)?;
                Ok(Arg::Scalar(match v.number() {
                    Ok(n) => finite(-n),
                    Err(k) => Value::Error(k),
                }))
            }
            // Unary plus returns its operand as it is, an area included.
            Expr::Pos(x) => self.eval(x),
            Expr::Percent(x) => {
                let v = self.scalar(x)?;
                Ok(Arg::Scalar(match v.number() {
                    Ok(n) => finite(n / 100.0),
                    Err(k) => Value::Error(k),
                }))
            }
            Expr::Binary(op, l, r) => {
                let a = self.scalar(l)?;
                let b = self.scalar(r)?;
                Ok(Arg::Scalar(binary(*op, &a, &b)))
            }
        }
    }

    /// An expression as one value: a scalar, or a reference to one cell;
    /// a reference to more is a construct this slice does not read.
    pub fn scalar(&self, e: &Expr) -> Result<Value, Reason> {
        let arg = self.eval(e)?;
        self.one(&arg)
    }

    /// An operand as one value, by the same rule.
    pub fn one(&self, arg: &Arg<'_>) -> Result<Value, Reason> {
        match arg {
            Arg::Scalar(v) => Ok(v.clone()),
            Arg::Area {
                areas,
                written,
                moved,
            } => match &areas[..] {
                [a] if a.is_cell() => self.env.read((a.sheet, a.top, a.left)),
                _ if *moved == (0, 0) => Err(Reason::Construct(written.to_string())),
                // The label a moved text would show: the cell's own reference.
                _ => Err(Reason::Construct(crate::refers::shift_a1_references(
                    written, moved.0, moved.1,
                ))),
            },
        }
    }

    /// A condition as `IF`, `IFS` and `NOT` read one: the operand as one
    /// value, then its truth value (`Value::truth`), a text read through a
    /// reference coerced as a typed one is. Measured: Excel's save of the
    /// subset fixture computes `IF(A6,1,2)` over the text `TRUE` as 1, and
    /// over the text `x` as `#VALUE!` (`scripts/recalc/subset_saved.xlsx`,
    /// `logic!C6` and `C3`). `AND` and `OR` read a reference as a range
    /// instead, skipping its texts (`logic!C13`).
    pub fn truth(&self, e: &Expr) -> Result<Result<bool, ErrorKind>, Reason> {
        Ok(self.scalar(e)?.truth())
    }

    /// [`Ctx::truth`] over an evaluated operand.
    pub fn truth_arg(&self, arg: &Arg<'_>) -> Result<Result<bool, ErrorKind>, Reason> {
        Ok(self.one(arg)?.truth())
    }

    /// Every cell the areas hold, in area order then row-major, to `f`;
    /// a cell nothing holds is skipped, as Excel skips it in a range. `f` is
    /// a type of its own and not a pointer, so that an aggregate's fold is
    /// compiled into the walk (KERNEL.25).
    pub fn each<F>(&self, areas: &[Area], mut f: F) -> Result<(), Reason>
    where
        F: FnMut(Value) -> Result<(), Reason>,
    {
        let model = self.env.model();
        for a in areas {
            let Some(sheet) = model.sheets.get(a.sheet) else {
                continue;
            };
            for ((row, col), _) in cells_in(sheet, self.env.extent(a.sheet), a) {
                let v = self.env.read((a.sheet, row, col))?;
                if v != Value::Empty {
                    f(v)?;
                }
            }
        }
        Ok(())
    }

    fn reference<'e>(&self, r: &'e RefExpr) -> Result<Arg<'e>, Reason> {
        let home = self.env.here().0;
        let (d_row, d_col) = self.offset.get();
        match areas_at(self.env.model(), home, r, d_row, d_col) {
            Ok(areas) => Ok(Arg::Area {
                areas,
                written: Cow::Borrowed(&r.written),
                moved: (d_row, d_col),
            }),
            Err(k) => Ok(Arg::Scalar(Value::Error(k))),
        }
    }

    /// A defined name's value: its text parsed here and evaluated, so that
    /// an area it names carries its label owned, the parse being this
    /// call's own.
    fn named(&self, name: &str, sheet: Option<&str>) -> Result<Arg<'static>, Reason> {
        let model = self.env.model();
        let home = self.env.here().0;
        let scope = match sheet {
            None => None,
            Some(s) => match model.find_sheet(s) {
                Some(i) => Some(i),
                None => return Ok(Arg::Scalar(Value::Error(ErrorKind::Ref))),
            },
        };
        let Some(text) = self.env.name(name, scope, home) else {
            return Ok(Arg::Scalar(Value::Error(ErrorKind::Name)));
        };
        if self.depth.get() >= NAME_DEPTH {
            return Ok(Arg::Scalar(Value::Error(ErrorKind::Name)));
        }
        let parsed = formula::parse(text).map_err(Reason::Construct)?;
        // A name's references are where the name says, whatever cell reads it.
        let offset = self.offset.replace((0, 0));
        self.depth.set(self.depth.get() + 1);
        let result = self.eval(&parsed).map(Arg::into_owned);
        self.depth.set(self.depth.get() - 1);
        self.offset.set(offset);
        result
    }

    fn call<'e>(&self, name: &str, args: &'e [Expr]) -> Result<Arg<'e>, Reason> {
        if let Some(form) = self.lib.form(name) {
            return form(self, args);
        }
        if let Some(f) = self.lib.strict(name) {
            let mut values = Vec::with_capacity(args.len());
            for a in args {
                values.push(self.eval(a)?);
            }
            return f(self, &values).map(Arg::Scalar);
        }
        Err(Reason::Function(name.to_string()))
    }
}

/// A result no cell can hold is `#NUM!`.
pub fn finite(v: f64) -> Value {
    if v == 0.0 {
        // No negative zero in a cell: Excel has none, and an empty sum is one.
        Value::Number(0.0)
    } else if v.is_finite() {
        Value::Number(v)
    } else {
        Value::Error(ErrorKind::Num)
    }
}

/// A binary operator over two values: the left operand's error first,
/// then the right's; arithmetic over the numbers each coerces to, `&` over
/// their texts, a comparison in Excel's order of kinds.
pub fn binary(op: BinOp, a: &Value, b: &Value) -> Value {
    if let Value::Error(k) = a {
        return Value::Error(*k);
    }
    if let Value::Error(k) = b {
        return Value::Error(*k);
    }
    match op {
        BinOp::Add | BinOp::Sub | BinOp::Mul | BinOp::Div | BinOp::Pow => {
            let x = match a.number() {
                Ok(n) => n,
                Err(k) => return Value::Error(k),
            };
            let y = match b.number() {
                Ok(n) => n,
                Err(k) => return Value::Error(k),
            };
            match op {
                BinOp::Add => finite(x + y),
                BinOp::Sub => finite(x - y),
                BinOp::Mul => finite(x * y),
                BinOp::Div => {
                    if y == 0.0 {
                        Value::Error(ErrorKind::Div0)
                    } else {
                        finite(x / y)
                    }
                }
                _ => power(x, y),
            }
        }
        BinOp::Concat => match (a.text(), b.text()) {
            (Ok(s), Ok(t)) => Value::Text(s + &t),
            (Err(k), _) | (_, Err(k)) => Value::Error(k),
        },
        BinOp::Eq => Value::Bool(compare(a, b) == Ordering::Equal),
        BinOp::Ne => Value::Bool(compare(a, b) != Ordering::Equal),
        BinOp::Lt => Value::Bool(compare(a, b) == Ordering::Less),
        BinOp::Gt => Value::Bool(compare(a, b) == Ordering::Greater),
        BinOp::Le => Value::Bool(compare(a, b) != Ordering::Greater),
        BinOp::Ge => Value::Bool(compare(a, b) != Ordering::Less),
    }
}

/// `^` as Excel has it: `0^0` is `#NUM!`, `0` to a negative power is
/// `#DIV/0!`, a negative base to a fractional power is `#NUM!`.
fn power(x: f64, y: f64) -> Value {
    if x == 0.0 && y == 0.0 {
        return Value::Error(ErrorKind::Num);
    }
    if x == 0.0 && y < 0.0 {
        return Value::Error(ErrorKind::Div0);
    }
    finite(x.powf(y))
}

/// Excel's order of kinds in a comparison: every number is less than every
/// text, every text less than every truth value; two texts compare without
/// case; an empty cell is 0, `""` or FALSE, whichever the other side is.
pub fn compare(a: &Value, b: &Value) -> Ordering {
    use Value::*;
    match (a, b) {
        (Empty, Empty) => Ordering::Equal,
        (Empty, Number(y)) => 0.0f64.partial_cmp(y).unwrap_or(Ordering::Equal),
        (Number(x), Empty) => x.partial_cmp(&0.0).unwrap_or(Ordering::Equal),
        (Empty, Text(t)) => {
            if t.is_empty() {
                Ordering::Equal
            } else {
                Ordering::Less
            }
        }
        (Text(t), Empty) => {
            if t.is_empty() {
                Ordering::Equal
            } else {
                Ordering::Greater
            }
        }
        (Empty, Bool(y)) => false.cmp(y),
        (Bool(x), Empty) => x.cmp(&false),
        (Number(x), Number(y)) => x.partial_cmp(y).unwrap_or(Ordering::Equal),
        (Text(x), Text(y)) => fold(x).cmp(&fold(y)),
        (Bool(x), Bool(y)) => x.cmp(y),
        _ => rank(a).cmp(&rank(b)),
    }
}

fn rank(v: &Value) -> u8 {
    match v {
        Value::Number(_) | Value::Empty => 0,
        Value::Text(_) => 1,
        Value::Bool(_) => 2,
        Value::Error(_) => 3,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn operators_coerce_as_excel_does() {
        let n = |v: f64| Value::Number(v);
        let t = |s: &str| Value::Text(s.to_string());
        assert_eq!(binary(BinOp::Add, &t("5"), &n(1.0)), n(6.0));
        assert_eq!(
            binary(BinOp::Add, &t("x"), &n(1.0)),
            Value::Error(ErrorKind::Value)
        );
        assert_eq!(binary(BinOp::Mul, &Value::Bool(true), &n(3.0)), n(3.0));
        assert_eq!(binary(BinOp::Add, &Value::Empty, &n(1.0)), n(1.0));
        assert_eq!(
            binary(BinOp::Div, &n(1.0), &n(0.0)),
            Value::Error(ErrorKind::Div0)
        );
        assert_eq!(
            binary(BinOp::Div, &n(1.0), &Value::Empty),
            Value::Error(ErrorKind::Div0)
        );
        assert_eq!(
            binary(BinOp::Pow, &n(0.0), &n(0.0)),
            Value::Error(ErrorKind::Num)
        );
        assert_eq!(
            binary(BinOp::Pow, &n(0.0), &n(-1.0)),
            Value::Error(ErrorKind::Div0)
        );
        assert_eq!(
            binary(BinOp::Pow, &n(-8.0), &n(1.0 / 3.0)),
            Value::Error(ErrorKind::Num)
        );
        assert_eq!(binary(BinOp::Pow, &n(2.0), &n(10.0)), n(1024.0));
        assert_eq!(
            binary(BinOp::Mul, &n(1e308), &n(10.0)),
            Value::Error(ErrorKind::Num)
        );
        assert_eq!(binary(BinOp::Concat, &n(1.0), &n(2.0)), t("12"));
        assert_eq!(
            binary(BinOp::Concat, &Value::Bool(true), &Value::Empty),
            t("TRUE")
        );
        assert_eq!(binary(BinOp::Concat, &n(0.1 + 0.2), &t("")), t("0.3"));
        // The left error first.
        assert_eq!(
            binary(
                BinOp::Add,
                &Value::Error(ErrorKind::NA),
                &Value::Error(ErrorKind::Div0)
            ),
            Value::Error(ErrorKind::NA)
        );
    }

    #[test]
    fn comparisons_follow_excel_s_order_of_kinds() {
        let n = |v: f64| Value::Number(v);
        let t = |s: &str| Value::Text(s.to_string());
        let yes = Value::Bool(true);
        assert_eq!(binary(BinOp::Eq, &t("ABC"), &t("abc")), yes);
        assert_eq!(binary(BinOp::Lt, &n(1e9), &t("")), yes);
        assert_eq!(binary(BinOp::Gt, &Value::Bool(false), &t("zzz")), yes);
        assert_eq!(binary(BinOp::Eq, &Value::Empty, &n(0.0)), yes);
        assert_eq!(binary(BinOp::Eq, &Value::Empty, &t("")), yes);
        assert_eq!(binary(BinOp::Eq, &Value::Empty, &Value::Bool(false)), yes);
        assert_eq!(binary(BinOp::Ne, &n(1.0), &t("1")), yes);
        assert_eq!(binary(BinOp::Le, &n(2.0), &n(2.0)), yes);
        assert_eq!(binary(BinOp::Ge, &n(1.0), &n(2.0)), Value::Bool(false));
    }

    #[test]
    fn the_library_is_keyed_without_case_and_a_name_holds_one_kind() {
        let mut lib = Library::empty();
        fn one(_: &Ctx<'_>, _: &[Arg]) -> Result<Value, Reason> {
            Ok(Value::Number(1.0))
        }
        fn two<'e>(_: &Ctx<'_>, _: &'e [Expr]) -> Result<Arg<'e>, Reason> {
            Ok(Arg::Scalar(Value::Number(2.0)))
        }
        lib.register("one", one);
        lib.register_form("Two", two);
        assert!(lib.has("ONE") && lib.has("two"));
        assert_eq!(lib.names(), vec!["ONE", "TWO"]);
        assert!(lib.strict("one").is_some() && lib.form("one").is_none());
        lib.register_form("one", two);
        assert!(lib.strict("one").is_none() && lib.form("one").is_some());
        assert_eq!(lib.len(), 2);
    }

    #[test]
    fn a_name_is_upper_cased_only_when_that_would_change_it() {
        // KERNEL.25: a parsed call's name, already in upper case, is looked
        // up as it stands; any other has its ASCII letters upper-cased first,
        // so a letter outside ASCII keeps its case, as it always has.
        fn one(_: &Ctx<'_>, _: &[Arg]) -> Result<Value, Reason> {
            Ok(Value::Number(1.0))
        }
        assert!(matches!(key("SUM"), Cow::Borrowed("SUM")));
        assert!(matches!(key("ÄRGER"), Cow::Borrowed("ÄRGER")));
        assert_eq!(key("Sum"), "SUM");
        assert_eq!(key("ärger"), "äRGER");
        let mut lib = Library::empty();
        lib.register("Ärger", one);
        lib.register("sum", one);
        assert_eq!(lib.names(), vec!["SUM", "ÄRGER"]);
        for (name, found) in [
            ("SUM", true),
            ("Sum", true),
            ("sum", true),
            ("ÄRGER", true),
            ("Ärger", true),
            ("ärger", false),
            ("äRGER", false),
        ] {
            assert_eq!(lib.strict(name).is_some(), found, "{name}");
            assert_eq!(lib.has(name), found, "{name}");
            assert!(lib.form(name).is_none(), "{name}");
        }
    }
}
