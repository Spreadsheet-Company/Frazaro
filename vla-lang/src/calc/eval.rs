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

use crate::rows::quoted;
use crate::sheet::Workbook;

use super::dense::Dense;
use super::formula::{self, BinOp, Expr, RefExpr};
use super::graph::{areas_at, cells_in, places, Area, Areas, CellId, Extent, Names};
use super::index::Indexes;
use super::value::{round15, ErrorKind, Value};

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
    /// The label as a datum: `(not-computed "FACT")`, `(not-computed cycle)`.
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
    /// A sheet's values kept as an array over its extent, when the
    /// environment keeps one (KERNEL.20: an engine's machine, for each
    /// `.last` twin; and since KERNEL.31 for each of its frame's own sheets
    /// its plan reads by a range or a run's moving cell, while none of their
    /// formulas holds a reason): each slot the value a read of its cell
    /// answers, or `Value::Empty` for a cell holding none, which a walk skips
    /// as a walk of the cells skips it; so a range of the sheet is walked by
    /// index and never cell by cell. `None`, a sheet walked through its
    /// cells, unless an environment says otherwise.
    fn dense(&self, _sheet: usize) -> Option<&Dense> {
        None
    }

    /// The cell being computed from now on, told by an evaluation of many
    /// cells at once before it computes at each (KERNEL.20, `super::vector`),
    /// so that [`Env::here`] answers it: an engine's machine keeps it as it
    /// keeps the cell it evaluates alone. Nothing, unless an environment
    /// says otherwise.
    fn enter(&self, _id: CellId) {}

    /// The indexes of the vectors the lookups and the criteria search, when
    /// the environment keeps them (KERNEL.8, `super::index`): an engine's
    /// machine, which tells them each cell that changes, and `Calc`, whose
    /// one pass changes none. `None`, every search a scan, unless an
    /// environment says otherwise.
    fn indexes(&self) -> Option<&Indexes> {
        None
    }

    /// Whether the formula of the cell being computed is a dynamic-array
    /// formula (KERNEL.8, slice b): one a host saved as an array formula,
    /// Excel 365's kind, in which a range where one value is wanted is an
    /// array, labelled until arrays come; a formula of the legacy kind
    /// intersects such a range with its cell instead ([`Ctx::one`]).
    /// `false`, the legacy kind, unless an environment says otherwise: an
    /// engine's grid, loaded from rows, holds that kind alone.
    fn dynamic(&self) -> bool {
        false
    }

    /// Whether reasons stand on a sheet beside its cells, at places no cell
    /// of it holds (KERNEL.8, slice b: an engine's machine, whose `.last`
    /// twin holds no cell for a formula not computed, its reason beside):
    /// a walk of an area of such a sheet reads every place of the area
    /// inside the sheet's extent, so that it meets a reason where a walk of
    /// the sheet the twin copies meets it. `false`, unless an environment
    /// says otherwise.
    fn reasons_beside(&self, _sheet: usize) -> bool {
        false
    }
}

/// The cell of an area a legacy formula reads where one value is wanted,
/// implicit intersection (KERNEL.8, slice b; Microsoft's page "Implicit
/// intersection operator: @"): the area's cell in the row of the cell being
/// computed, for an area one column wide; in its column, for one a row
/// high; in both, for any other; `None` where that cell is not in the area,
/// which reads as `#VALUE!`. An area on another sheet is met by the same
/// row and column.
pub(crate) fn intersection(a: &Area, here: CellId) -> Option<CellId> {
    let (_, row, col) = here;
    let row = if a.top == a.bottom {
        a.top
    } else {
        (a.top..=a.bottom).contains(&row).then_some(row)?
    };
    let col = if a.left == a.right {
        a.left
    } else {
        (a.left..=a.right).contains(&col).then_some(col)?
    };
    Some((a.sheet, row, col))
}

/// A strict function: its arguments evaluated, a value or a reason out.
pub type Strict = fn(&Ctx<'_>, &[Arg<'_>]) -> Result<Value, Reason>;

/// A form: its arguments as written, an operand or a reason out, the operand
/// living as long as the arguments it was evaluated from.
pub type Form = for<'e> fn(&Ctx<'_>, &'e [Expr]) -> Result<Arg<'e>, Reason>;

/// One of the language's own functions that an evaluation of many cells at
/// once computes by itself (`super::vector`), known by the name the
/// language's library registers it under: `IF` (KERNEL.20); the aggregates
/// and `NOT`, which it computes by the library's own rule and fold
/// (KERNEL.26, `super::library::Fold`); and the language's other functions,
/// which it computes by the library's own meaning of each (KERNEL.28). Any
/// other registration of the name replaces it, so that a crate above that
/// registers a meaning of its own has that meaning computed, cell by cell.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum Known {
    /// `library::if_`.
    If,
    /// `SUM`, `MIN` and `MAX`: `library::Sum`, `Least` and `Most`.
    Sum,
    Min,
    Max,
    /// `AND` and `OR`: `library::All` and `Some_`.
    And,
    Or,
    /// `NOT`: `library::negated`.
    Not,
    /// `ABS`, `INT` and `SIN`: `library::abs_of`, `int_of` and `sin_of`.
    Abs,
    Int,
    Sin,
    /// `MOD`: `library::mod_of`.
    Mod,
    /// `ROW` and `COLUMN`: `library::placed`.
    Row,
    Column,
    /// `CHOOSE`: `library::chosen`.
    Choose,
    /// `INDEX` (KERNEL.8), a form answering a reference, which a run
    /// computes by itself as it computes `CHOOSE`, its row and column in
    /// lanes and the cell picked read at each cell, and a program lowers:
    /// `library::picked`.
    Index,
    /// `IFERROR` and `IFNA` (KERNEL.8), forms a run computes by itself as it
    /// computes `IF`: the value at every cell, the other at the cells whose
    /// value is an error.
    IfError,
    IfNa,
    /// `ISERROR` and `ISNA` (KERNEL.8), which a run calls at each cell and a
    /// program leaves to `Ctx`, an error being what they read and no
    /// register holds.
    IsError,
    IsNa,
    /// The lookups and the criteria (KERNEL.8), called at each cell, each
    /// searching one of its operands by an index or a bisection
    /// ([`Known::searches`]).
    Match,
    XMatch,
    VLookup,
    HLookup,
    Lookup,
    CountIf,
    SumIf,
}

impl Known {
    /// Whether the function searches its `j`th operand, by an index or a
    /// bisection, rather than walking it whole at each cell (KERNEL.8): the
    /// census names such a read `keyed` where every `$` holds it.
    pub(crate) fn searches(self, j: usize) -> bool {
        match self {
            Known::Match | Known::XMatch | Known::VLookup | Known::HLookup | Known::Lookup => {
                j == 1
            }
            Known::CountIf => j == 0,
            Known::SumIf => j == 0 || j == 2,
            _ => false,
        }
    }
}

/// The functions a formula may call, by upper-case name: the seam. Every
/// function registered is a function of its arguments and of the cell it is
/// computed at, and of nothing else: no clock, no chance, no state of its
/// own. An engine's machine computes a cell again only when a cell it reads
/// has changed (KERNEL.20, `super::reach`), so a function that answered
/// differently over the same arguments would be computed once and kept.
#[derive(Default)]
pub struct Library {
    /// Each strict function, and whether it is one an evaluation of many
    /// cells at once knows (KERNEL.26).
    strict: BTreeMap<String, (Strict, Option<Known>)>,
    /// Each form, and whether it is one an evaluation of many cells at once
    /// knows (KERNEL.20).
    forms: BTreeMap<String, (Form, Option<Known>)>,
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
        self.strict.insert(key, (f, None));
    }

    /// Register a form.
    pub fn register_form(&mut self, name: &str, f: Form) {
        let key = name.to_ascii_uppercase();
        self.strict.remove(&key);
        self.forms.insert(key, (f, None));
    }

    /// Register one of the language's own forms as one an evaluation of
    /// many cells at once computes by itself (KERNEL.20).
    pub(crate) fn register_known(&mut self, name: &str, f: Form, known: Known) {
        let key = name.to_ascii_uppercase();
        self.strict.remove(&key);
        self.forms.insert(key, (f, Some(known)));
    }

    /// Register one of the language's own strict functions as one an
    /// evaluation of many cells at once computes by itself, by the same rule
    /// and fold (KERNEL.26).
    pub(crate) fn register_known_strict(&mut self, name: &str, f: Strict, known: Known) {
        let key = name.to_ascii_uppercase();
        self.forms.remove(&key);
        self.strict.insert(key, (f, Some(known)));
    }

    /// The known function registered under that name, without case: a form
    /// or a strict function.
    pub(crate) fn known(&self, name: &str) -> Option<Known> {
        let key = key(name);
        match self.forms.get(&*key) {
            Some(&(_, known)) => known,
            None => self.strict.get(&*key).and_then(|&(_, known)| known),
        }
    }

    /// The strict function of that name, without case.
    pub fn strict(&self, name: &str) -> Option<Strict> {
        self.strict.get(&*key(name)).map(|&(f, _)| f)
    }

    /// The form of that name, without case.
    pub fn form(&self, name: &str) -> Option<Form> {
        self.forms.get(&*key(name)).map(|&(f, _)| f)
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

/// The construct a defined name leaves unread for a formula read at a sheet
/// (KERNEL.8, slice b): the first name the parse reads, or a name such a
/// name's text reads, as deep as an evaluation follows names, whose text the
/// parse does not read, and the construct that parse names; `None` where
/// every name it reaches is read, or is none the workbook defines, which an
/// evaluation answers `#NAME?`. What a load refuses with the constructs
/// (`calc-construct-not-read`), so that no cell of an engine's grid holds
/// the reason [`Ctx`] would label it with at a step.
pub(crate) fn unread_name(
    model: &Workbook,
    names: &Names,
    e: &Expr,
    home: usize,
) -> Option<String> {
    fn go(model: &Workbook, names: &Names, e: &Expr, home: usize, depth: u32) -> Option<String> {
        let mut found = None;
        e.each_read(&mut |read| {
            let formula::Read::Name { name, sheet } = read else {
                return;
            };
            if found.is_some() || depth >= NAME_DEPTH {
                return;
            }
            let scope = match sheet {
                None => None,
                Some(s) => match model.find_sheet(s) {
                    Some(i) => Some(i),
                    None => return,
                },
            };
            let Some(text) = names.resolve(name, scope, home) else {
                return;
            };
            found = match formula::parse(text) {
                Err(construct) => Some(construct),
                Ok(parsed) => go(model, names, &parsed, home, depth + 1),
            };
        });
        found
    }
    go(model, names, e, home, 0)
}

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
            // One operator, as it always was; a chain of them goes round a
            // loop instead (VLA.1).
            Expr::Binary(op, l, r) if !matches!(**l, Expr::Binary(..)) => {
                let a = self.scalar(l)?;
                let b = self.scalar(r)?;
                Ok(Arg::Scalar(operate(*op, &a, &b)?))
            }
            Expr::Binary(..) => self.chain(e),
        }
    }

    /// A chain of binary operators (`1+2+3`, which Excel reads left to
    /// right) is a left spine; it is evaluated as the recursion would, the
    /// leftmost operand first and each link in turn, but with its right
    /// operands on a stack, so that a formula as long as Excel holds, 4,096
    /// links, needs no deeper stack than a short one (VLA.1). The first
    /// sixteen are held in place, so that a chain such as a neighbour count
    /// allocates nothing (KERNEL.25).
    fn chain<'e>(&self, e: &'e Expr) -> Result<Arg<'e>, Reason> {
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
        let mut acc = self.scalar(left)?;
        for i in (0..links).rev() {
            let (op, r) = if i < NEAR { near[i] } else { far[i - NEAR] };
            let b = self.scalar(r)?;
            acc = operate(op, &acc, &b)?;
        }
        Ok(Arg::Scalar(acc))
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
                // A legacy formula intersects the area with its cell, and a
                // span of sheets is `#REF!`, as Excel's save of the
                // intersection's fixture holds it (KERNEL.8, slice b); a
                // dynamic-array formula's range is an array, labelled.
                [a] if !self.env.dynamic() => self.intersect(a),
                _ if !self.env.dynamic() => Ok(Value::Error(ErrorKind::Ref)),
                _ if *moved == (0, 0) => Err(Reason::Construct(written.to_string())),
                // The label a moved text would show: the cell's own reference.
                _ => Err(Reason::Construct(crate::refers::shift_a1_references(
                    written, moved.0, moved.1,
                ))),
            },
        }
    }

    /// An area's cell at the cell being computed, by [`intersection`], read;
    /// `#VALUE!` where the area holds none there.
    pub fn intersect(&self, a: &Area) -> Result<Value, Reason> {
        match intersection(a, self.env.here()) {
            Some(id) => self.env.read(id),
            None => Ok(Value::Error(ErrorKind::Value)),
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
    /// compiled into the walk (KERNEL.25). A sheet the environment keeps as
    /// an array is walked by index, the same values in the same order
    /// (KERNEL.20); a sheet holding reasons beside its cells, at every place,
    /// so that a reason answers where it stands (KERNEL.8, slice b).
    pub fn each<F>(&self, areas: &[Area], mut f: F) -> Result<(), Reason>
    where
        F: FnMut(Value) -> Result<(), Reason>,
    {
        let model = self.env.model();
        for a in areas {
            if let Some(kept) = self.env.dense(a.sheet) {
                kept.each(a, &mut f)?;
                continue;
            }
            let Some(sheet) = model.sheets.get(a.sheet) else {
                continue;
            };
            if self.env.reasons_beside(a.sheet) {
                for (row, col) in places(self.env.extent(a.sheet), a) {
                    let v = self.env.read((a.sheet, row, col))?;
                    if v != Value::Empty {
                        f(v)?;
                    }
                }
                continue;
            }
            for ((row, col), _) in cells_in(sheet, self.env.extent(a.sheet), a) {
                let v = self.env.read((a.sheet, row, col))?;
                if v != Value::Empty {
                    f(v)?;
                }
            }
        }
        Ok(())
    }

    pub(super) fn reference<'e>(&self, r: &'e RefExpr) -> Result<Arg<'e>, Reason> {
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
            // One argument or two, the counts a call mostly has, held on the
            // stack in a frame of their own (KERNEL.20), so that this frame,
            // which evaluation stacks once a level of nesting, holds no array
            // (`VLA.1`'s limits are pinned on a small stack); any other count
            // in a vector. Left to right, stopping at the first not computed.
            return match args {
                [a] => self.strict_one(f, a),
                [a, b] => self.strict_two(f, a, b),
                _ => {
                    let mut values = Vec::with_capacity(args.len());
                    for a in args {
                        values.push(self.eval(a)?);
                    }
                    f(self, &values)
                }
            }
            .map(Arg::Scalar);
        }
        Err(Reason::Function(name.to_string()))
    }

    /// A strict call of one argument, the argument on the stack (KERNEL.20).
    fn strict_one(&self, f: Strict, a: &Expr) -> Result<Value, Reason> {
        let x = self.eval(a)?;
        f(self, std::slice::from_ref(&x))
    }

    /// A strict call of two arguments, evaluated left to right, on the stack.
    fn strict_two(&self, f: Strict, a: &Expr, b: &Expr) -> Result<Value, Reason> {
        let pair = [self.eval(a)?, self.eval(b)?];
        f(self, &pair)
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
/// their texts, a comparison in Excel's order of kinds. Its arithmetic and
/// its comparison of two numbers are [`arithmetic`] and [`compares`] over
/// [`compare_numbers`], which an evaluation of many cells at once calls over
/// its lanes of numbers (KERNEL.26), so that each rule is written once. A
/// cell computes it through [`operate`], which labels what this answers
/// with the language's own reading: a comparison by size of two texts that
/// a character beyond Excel's known order decides, weighed here after the
/// known ones by its code point, and a text `&` cuts through a surrogate
/// pair, cut here before the pair.
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
            arithmetic(op, x, y)
        }
        // Cut to Excel's 32,767 units, as its save holds `&` (KERNEL.8,
        // slice c).
        BinOp::Concat => match (a.text(), b.text()) {
            (Ok(s), Ok(t)) => Value::Text(super::texts::cut_to_most(s + &t).0),
            (Err(k), _) | (_, Err(k)) => Value::Error(k),
        },
        _ => Value::Bool(compares(op, compare(a, b))),
    }
}

/// [`binary`] as a cell computes it (KERNEL.8, slice c): a comparison by
/// size of two texts, an empty value among them as `""`, that a character
/// beyond Excel's known order decides is labelled (`order::sized`), and a
/// text `&` would cut through a surrogate pair at its 32,767th unit,
/// where Excel keeps half a character, is labelled (`texts::halved`). The
/// evaluator's operators and a lane's call it, so that a run answers as a
/// cell alone does.
pub(crate) fn operate(op: BinOp, a: &Value, b: &Value) -> Result<Value, Reason> {
    if matches!(a, Value::Error(_)) || matches!(b, Value::Error(_)) {
        return Ok(binary(op, a, b));
    }
    match op {
        BinOp::Lt | BinOp::Gt | BinOp::Le | BinOp::Ge => {
            if let Some((x, y)) = texts_of(a, b) {
                return Ok(Value::Bool(compares(op, super::order::sized(x, y)?)));
            }
        }
        BinOp::Concat => {
            if let (Ok(s), Ok(t)) = (a.text(), b.text()) {
                return match super::texts::cut_to_most(s + &t) {
                    (t, true) => Ok(Value::Text(t)),
                    (_, false) => Err(super::texts::halved()),
                };
            }
        }
        _ => {}
    }
    Ok(binary(op, a, b))
}

/// Two operands as the texts a comparison orders: two texts, or a text and
/// an empty value, which is `""` beside a text.
fn texts_of<'v>(a: &'v Value, b: &'v Value) -> Option<(&'v str, &'v str)> {
    match (a, b) {
        (Value::Text(x), Value::Text(y)) => Some((x, y)),
        (Value::Empty, Value::Text(y)) => Some(("", y)),
        (Value::Text(x), Value::Empty) => Some((x, "")),
        _ => None,
    }
}

/// An arithmetic operator over two numbers (`+`, `-`, `*`, `/`, `^`).
#[inline]
pub(crate) fn arithmetic(op: BinOp, x: f64, y: f64) -> Value {
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

/// Whether a comparison operator holds of how its operands compare.
#[inline]
pub(crate) fn compares(op: BinOp, ordering: Ordering) -> bool {
    match op {
        BinOp::Eq => ordering == Ordering::Equal,
        BinOp::Ne => ordering != Ordering::Equal,
        BinOp::Lt => ordering == Ordering::Less,
        BinOp::Gt => ordering == Ordering::Greater,
        BinOp::Le => ordering != Ordering::Greater,
        _ => ordering != Ordering::Less,
    }
}

/// How two numbers compare under Excel's comparison operators and its
/// criteria (KERNEL.8): equal where they are equal to fifteen significant
/// digits, as Excel's save of the lookups' fixture holds `0.1+0.2=0.3` TRUE
/// and `COUNTIF` of `0.1+0.2` against 0.3 one (`Criteria!H30`, `H32`), and
/// otherwise by value, as rounding to fifteen digits never reverses an
/// order. Two numbers apart by more than [`NEAR`] of the larger, or of two
/// signs, are further apart than fifteen digits can merge (two numbers
/// merged differ by at most a unit of their fifteenth digit, 10^-14 of the
/// larger), so that only the rare pair near enough is rounded, out of line.
/// The lookups compare exactly ([`exact_order`]).
#[inline]
pub(crate) fn compare_numbers(x: f64, y: f64) -> Ordering {
    let o = exact_order(x, y);
    if o != Ordering::Equal && near(x, y) {
        to_fifteen_digits(x, y, o)
    } else {
        o
    }
}

/// How near two numbers stand, as a share of the larger, for fifteen digits
/// to merge them: ten times the most they can differ by.
const NEAR: f64 = 1e-13;

/// Whether two numbers stand near enough for fifteen digits to merge them:
/// a superset, which [`to_fifteen_digits`] decides, the sum of their sizes
/// standing for the larger as plain arithmetic, and larger.
#[inline]
fn near(x: f64, y: f64) -> bool {
    (x - y).abs() <= NEAR * (x.abs() + y.abs())
}

/// Whether [`compare_numbers`] might answer otherwise than
/// [`exact_order`]: two numbers unequal and [`near`] each other.
#[inline]
pub(crate) fn may_merge(x: f64, y: f64) -> bool {
    (x != y) & near(x, y)
}

/// The bounds a number `y` is compared by over a lane (KERNEL.8): a number
/// under the first is less than `y` and one over the second greater, to
/// fifteen digits as exactly; one between them is `y`, or is compared again
/// by [`compare_numbers`]. [`NEAR`] of `y` either side, which holds every
/// number fifteen digits could merge with it, since such a number differs
/// from `y` by at most 10^-14 of the larger of the two.
#[inline]
pub(crate) fn near_bounds(y: f64) -> (f64, f64) {
    let reach = NEAR * y.abs();
    (y - reach, y + reach)
}

/// [`compare_numbers`] of two numbers [`near`] each other.
#[cold]
#[inline(never)]
fn to_fifteen_digits(x: f64, y: f64, o: Ordering) -> Ordering {
    if round15(x) == round15(y) {
        Ordering::Equal
    } else {
        o
    }
}

/// How two numbers compare exactly, by value: the lookups' order and their
/// equality, as Excel's save of the lookups' fixture finds no `0.1+0.2`
/// for 0.3 by `MATCH` or `VLOOKUP` exactly (`Criteria!H31`, `H33`).
#[inline]
pub(crate) fn exact_order(x: f64, y: f64) -> Ordering {
    x.partial_cmp(&y).unwrap_or(Ordering::Equal)
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
/// text, every text less than every truth value; two numbers equal to
/// fifteen significant digits are equal ([`compare_numbers`]); two texts in
/// Excel's order of texts, case aside, a character beyond its known order
/// weighed after the known ones by its code point (`super::order`, KERNEL.8);
/// an empty cell is 0, `""` or FALSE, whichever the other side is.
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
        (Number(x), Number(y)) => compare_numbers(*x, *y),
        (Text(x), Text(y)) => super::order::compare(x, y).order,
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
        // KERNEL.8, as Excel's save holds them: two numbers equal to
        // fifteen digits are equal, and a letter's case beyond ASCII is no
        // difference.
        let no = Value::Bool(false);
        assert_eq!(binary(BinOp::Eq, &n(0.1 + 0.2), &n(0.3)), yes);
        assert_eq!(binary(BinOp::Ne, &n(0.1 + 0.2), &n(0.3)), no);
        assert_eq!(binary(BinOp::Lt, &n(0.3), &n(0.1 + 0.2)), no);
        assert_eq!(binary(BinOp::Le, &n(0.1 + 0.2), &n(0.3)), yes);
        assert_eq!(binary(BinOp::Eq, &n(0.3), &n(0.3000000000001)), no);
        assert_eq!(binary(BinOp::Eq, &t("É"), &t("é")), yes);
        assert_eq!(binary(BinOp::Eq, &t("é"), &t("e")), no);
        // KERNEL.8, slice c, as Excel's save of the text fixture holds them:
        // its order of texts, `ß` and `ss` one.
        assert_eq!(binary(BinOp::Eq, &t("ß"), &t("ss")), yes);
        assert_eq!(binary(BinOp::Lt, &t("~"), &t("a")), yes);
        assert_eq!(binary(BinOp::Lt, &t("coop"), &t("co-op")), yes);
        assert_eq!(binary(BinOp::Lt, &Value::Empty, &t("-")), yes);
    }

    #[test]
    fn a_cell_labels_what_binary_answers_by_the_language_s_reading() {
        let t = |s: &str| Value::Text(s.to_string());
        let unknown = Err(super::super::order::unknown());
        // A comparison by size a character beyond the table decides.
        assert_eq!(operate(BinOp::Lt, &t("ø"), &t("o")), unknown);
        assert_eq!(operate(BinOp::Ge, &t("a\u{2013}"), &t("a")), unknown);
        assert_eq!(binary(BinOp::Lt, &t("o"), &t("ø")), Value::Bool(true));
        // Equality, a difference before the character, and an error first
        // are answered.
        assert_eq!(operate(BinOp::Eq, &t("ø"), &t("o")), Ok(Value::Bool(false)));
        assert_eq!(operate(BinOp::Le, &t("Ø"), &t("ø")), Ok(Value::Bool(true)));
        assert_eq!(
            operate(BinOp::Lt, &t("aø"), &t("bø")),
            Ok(Value::Bool(true))
        );
        assert_eq!(
            operate(BinOp::Lt, &Value::Error(ErrorKind::NA), &t("ø")),
            Ok(Value::Error(ErrorKind::NA))
        );
        // `&` cut to 32,767 units, and through a pair labelled.
        let x = t(&"x".repeat(32_767));
        assert_eq!(operate(BinOp::Concat, &x, &t("y")), Ok(x.clone()));
        let pairs = t(&"\u{1F600}".repeat(16_383));
        assert_eq!(
            operate(BinOp::Concat, &pairs, &t("\u{1F600}")),
            Err(super::super::texts::halved())
        );
        assert_eq!(binary(BinOp::Concat, &pairs, &t("\u{1F600}")), pairs);
    }

    #[test]
    fn two_numbers_compare_as_their_fifteen_digits_do() {
        // KERNEL.8: the comparison's shortcut never decides a pair that
        // rounding to fifteen digits would merge. Numbers of every
        // magnitude a double holds, each against the 300 doubles above it
        // and below it, and against 0 and its own negation: equal exactly
        // where their fifteen digits are, and otherwise ordered by value.
        let mut state = 0x9E37_79B9_7F4A_7C15u64;
        let mut next = || {
            state = state.wrapping_add(0x9E37_79B9_7F4A_7C15);
            let mut z = state;
            z = (z ^ (z >> 30)).wrapping_mul(0xBF58_476D_1CE4_E5B9);
            z = (z ^ (z >> 27)).wrapping_mul(0x94D0_49BB_1331_11EB);
            z ^ (z >> 31)
        };
        let mut merged = 0usize;
        for i in 0..600 {
            let x = match i % 4 {
                0 => f64::from_bits(next() & 0x7FEF_FFFF_FFFF_FFFF),
                1 => (next() % 1_000_000) as f64 / 1_000.0,
                2 => 0.1 * (next() % 100) as f64 + 0.2,
                _ => -f64::from_bits(next() & 0x7FEF_FFFF_FFFF_FFFF) * 1e-200,
            };
            let mut others = vec![0.0, -x, round15(x)];
            let mut up = x;
            let mut down = x;
            for _ in 0..300 {
                up = f64::from_bits(up.to_bits() + 1);
                down = f64::from_bits(down.to_bits().wrapping_sub(1));
                others.push(up);
                others.push(down);
            }
            let rx = round15(x);
            for y in others.into_iter().filter(|y| y.is_finite()) {
                let want = if rx == round15(y) {
                    Ordering::Equal
                } else {
                    x.partial_cmp(&y).unwrap()
                };
                assert_eq!(compare_numbers(x, y), want, "{x:e} against {y:e}");
                assert_eq!(compare_numbers(y, x), want.reverse(), "{y:e} against {x:e}");
                merged += usize::from(want == Ordering::Equal && x != y);
            }
        }
        assert!(merged > 5_000, "the pairs merged: {merged}");
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
