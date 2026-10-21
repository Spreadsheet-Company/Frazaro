//! The formula parser (KERNEL.7, 2026-10-08): operators and calls over the
//! reference scanner the language already has.
//!
//! `refers::scan` reads every reference out of a formula's text and nothing
//! else, held to the refers golden the reference writes; this module reads
//! the rest, the literals, the operators and the function calls, splicing
//! the scanner's records in by position, so that the one reading of what
//! `A1`, `Data!B2:C9`, `Rate` or `'Q1 Data'!A:A` means stays in `refers.rs`.
//! The grammar is the file's: a formula is stored in the invariant spelling,
//! English function names, `,` between arguments and `.` as the decimal
//! mark, whatever locale typed it (ISO/IEC 29500-1, 18.17). Excel's
//! precedence, lowest first: the comparisons; `&`; `+` and `-`; `*` and
//! `/`; `^`, left to right; a sign, so that `-2^2` is 4; `%`; and `:`
//! between two expressions that answer references (`A1:INDEX(A:A,5)`),
//! above them all.
//!
//! An array constant (`{1,2;3,4}`: `,` between the values of a row, `;`
//! between rows, each a number, a text, a truth value or an error) reads
//! as the array it holds since `KERNEL.8`'s sixth slice, and `:` between
//! expressions as a range whose plan's edges are its reach
//! ([`RangeExpr::reach`]).
//!
//! A construct this slice does not read is not an error: the parse answers
//! with the construct as written, and the cell shows `not computed here`
//! naming it, the honest static label of `HORIZON.md` section 12.6. Those
//! are: implicit intersection by `@` or by a space between two references,
//! a union in parentheses, a spill range (`A1#`), a structured reference
//! (`Sales[Amount]`), and a link into another workbook
//! (`[Rates.xlsx]Sheet1!A1`). A range where one value is wanted (`=B1:B2*2`)
//! a legacy formula intersects and a dynamic one reads whole (`KERNEL.8`'s
//! second and sixth slices), by the kind the file gives the formula.

use crate::messages::{raise, Refusal};
use crate::refers::{self, is_ident_char, FormulaRef, Kind, Shape};
use crate::sheet::{MAX_COLUMN, MAX_ROW};

use super::array::Array;
use super::value::{ErrorKind, Value};

/// One reference as the parser hands it to the evaluator: a rectangle on
/// one sheet, or the same rectangle on every sheet of a span.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct RefExpr {
    /// The sheet named in front of `!`, or `None` for the formula's own.
    pub sheet: Option<String>,
    /// The last sheet of a 3D span, `Jan:Dec!A1`.
    pub last_sheet: Option<String>,
    pub top: u32,
    pub left: u32,
    pub bottom: u32,
    pub right: u32,
    /// The reference as written, for a label.
    pub written: String,
    /// The reference's part as the scanner read it (a cell, a range, a
    /// whole column or row) and its corners as written with their `$`
    /// marks, which a move at an offset reads (KERNEL.22): the rectangle
    /// above is these corners sorted.
    pub shape: Shape,
    pub corners: Corners,
    /// The sheets named in front of `!`, the first and the last, found once
    /// against the grid an engine's machine placed the formula in
    /// (KERNEL.20), whose sheets are fixed for its life. `None` for a
    /// reference to the formula's own sheet, for a sheet the grid does not
    /// hold, for one past the 256th, and for any parse not placed, each of
    /// which finds its sheets by name at every evaluation, as before. A byte
    /// each, so that the field fits the padding the struct already had: a
    /// parse's size is what `VLA.1`'s limits are pinned on, a stack of 1 MiB,
    /// and the reference is the largest kind of expression (measured: two
    /// words would have made every expression 24 bytes larger and the
    /// deepest formula's load 192 KiB deeper).
    pub found: Option<(u8, u8)>,
}

/// A reference's corners as written, each number with its `$` mark, as
/// `refers::FormulaRef` holds them (a cell's second corner 0): what
/// [`RefExpr::moved`] moves, corner by corner, as
/// `refers::shift_a1_references` moves the text.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Corners {
    pub row1: u32,
    pub col1: u32,
    pub row2: u32,
    pub col2: u32,
    pub row_abs1: bool,
    pub col_abs1: bool,
    pub row_abs2: bool,
    pub col_abs2: bool,
}

impl RefExpr {
    /// Whether the reference names one cell.
    pub fn is_cell(&self) -> bool {
        self.top == self.bottom && self.left == self.right && self.last_sheet.is_none()
    }

    /// The rectangle this reference names in a cell `d_row` rows down and
    /// `d_col` columns right of the cell whose formula it was read from: each
    /// relative number moved and each `$` one kept, corner by corner as
    /// written, then sorted, which is what parsing the text that
    /// `refers::shift_a1_references` moves gives (KERNEL.22, a shape
    /// evaluated at an offset); `None` when a corner leaves the sheet,
    /// where the mover writes `#REF!`. At no offset it is the rectangle as
    /// read.
    pub fn moved(&self, d_row: i64, d_col: i64) -> Option<(u32, u32, u32, u32)> {
        moved_at(
            self.shape,
            &self.corners,
            (self.top, self.left, self.bottom, self.right),
            d_row,
            d_col,
        )
    }
}

/// [`RefExpr::moved`]'s rule over a reference's parts: its shape, its
/// corners as written and its rectangle as read, which is what a program
/// keeps of a reference (KERNEL.44, `super::program`), so that the rule is
/// written once. Inlined where it is called, so that `RefExpr::moved` is
/// compiled as it was before the rule was shared: called out of line, it
/// cost Sand's frame 3 to 5% in wasm, every read of a run moving a reference
/// (measured, alternated, KERNEL.44's record).
#[inline(always)]
pub(crate) fn moved_at(
    shape: Shape,
    c: &Corners,
    rect: (u32, u32, u32, u32),
    d_row: i64,
    d_col: i64,
) -> Option<(u32, u32, u32, u32)> {
    if d_row == 0 && d_col == 0 {
        return Some(rect);
    }
    let row = |n: u32, abs: bool| shifted(n, abs, d_row, MAX_ROW);
    let col = |n: u32, abs: bool| shifted(n, abs, d_col, MAX_COLUMN);
    match shape {
        Shape::Cell => {
            let k = col(c.col1, c.col_abs1)?;
            let r = row(c.row1, c.row_abs1)?;
            Some((r, k, r, k))
        }
        Shape::Range => {
            let k1 = col(c.col1, c.col_abs1)?;
            let r1 = row(c.row1, c.row_abs1)?;
            let k2 = col(c.col2, c.col_abs2)?;
            let r2 = row(c.row2, c.row_abs2)?;
            Some((r1.min(r2), k1.min(k2), r1.max(r2), k1.max(k2)))
        }
        Shape::Column => {
            let k1 = col(c.col1, c.col_abs1)?;
            let k2 = col(c.col2, c.col_abs2)?;
            Some((1, k1.min(k2), MAX_ROW, k1.max(k2)))
        }
        Shape::Row => {
            let r1 = row(c.row1, c.row_abs1)?;
            let r2 = row(c.row2, c.row_abs2)?;
            Some((r1.min(r2), 1, r1.max(r2), MAX_COLUMN))
        }
        Shape::None => None,
    }
}

/// A number moved by a distance unless its `$` holds it, `None` off the
/// sheet: `refers.rs`'s rule for the text, kept the same here.
fn shifted(n: u32, is_abs: bool, d: i64, max: u32) -> Option<u32> {
    if is_abs {
        return Some(n);
    }
    let m = i64::from(n) + d;
    (m >= 1 && m <= i64::from(max)).then_some(m as u32)
}

/// The binary operators.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum BinOp {
    Add,
    Sub,
    Mul,
    Div,
    Pow,
    Concat,
    Eq,
    Ne,
    Lt,
    Gt,
    Le,
    Ge,
}

/// A parsed formula. A list that may grow (KERNEL.8's decision 9).
#[derive(Clone, Debug, PartialEq)]
#[non_exhaustive]
pub enum Expr {
    Number(f64),
    Text(String),
    Bool(bool),
    Error(ErrorKind),
    /// An argument left out, `IF(A1,,B1)`.
    Empty,
    /// An array constant, `{1,2;3,4}` (KERNEL.8, slice f).
    Array(Array),
    Ref(RefExpr),
    /// `:` between two expressions that answer references,
    /// `A1:INDEX(A:A,5)` (KERNEL.8, slice f).
    Range(Box<RangeExpr>),
    /// A defined name, with the sheet that scopes it when one is written.
    Name {
        name: String,
        sheet: Option<String>,
    },
    /// A function call, the name in upper case.
    Call(String, Vec<Expr>),
    Neg(Box<Expr>),
    Pos(Box<Expr>),
    Percent(Box<Expr>),
    Binary(BinOp, Box<Expr>, Box<Expr>),
}

/// `:` between two expressions that answer references (KERNEL.8, slice f):
/// the rectangle spanning the two it evaluates to, and the rectangle every
/// such span lies in, which is what a plan reads it by.
#[derive(Clone, Debug, PartialEq)]
pub struct RangeExpr {
    pub left: Expr,
    pub right: Expr,
    /// The rectangle spanning every reference either end may answer, as a
    /// reference written at the formula's cell: a reference's own, `INDEX`'s
    /// its reference's, `IF`'s and `CHOOSE`'s their branches', on the one
    /// sheet they share. The plan's edges for the range, as a reference's
    /// are for it. `None` where an end is anything else (a defined name, a
    /// function the plan cannot bound) or the ends stand on two sheets,
    /// which the evaluation labels.
    pub reach: Option<RefExpr>,
    /// Whether the reach moves with the formula's cell as every span it
    /// bounds does, each of its edges held by a `$` in every end or in none,
    /// so that a shape read at one cell and evaluated at another reads it
    /// moved (KERNEL.22); a range whose ends move apart is labelled there.
    pub rigid: bool,
}

impl RangeExpr {
    /// The range between two ends, its reach found.
    pub fn new(left: Expr, right: Expr) -> RangeExpr {
        let mut ends = Vec::new();
        let (reach, rigid) = if bounded(&left, &mut ends) && bounded(&right, &mut ends) {
            spanned(&ends)
        } else {
            (None, false)
        };
        RangeExpr {
            left,
            right,
            reach,
            rigid,
        }
    }
}

/// The references an end of a range may answer, into `out`: whether every
/// one is bounded ([`RangeExpr::reach`]).
fn bounded<'e>(e: &'e Expr, out: &mut Vec<&'e RefExpr>) -> bool {
    match e {
        Expr::Ref(r) => {
            out.push(r);
            true
        }
        Expr::Pos(x) => bounded(x, out),
        Expr::Range(r) => match &r.reach {
            Some(reach) if r.rigid => {
                out.push(reach);
                true
            }
            _ => false,
        },
        Expr::Call(name, args) => match (name.as_str(), args.as_slice()) {
            ("INDEX", [first, ..]) => bounded(first, out),
            ("IF", [_, yes]) => bounded(yes, out),
            ("IF", [_, yes, no]) => bounded(yes, out) && bounded(no, out),
            ("CHOOSE", [_, values @ ..]) if !values.is_empty() => {
                values.iter().all(|v| bounded(v, out))
            }
            _ => false,
        },
        _ => false,
    }
}

/// The rectangle spanning references on one sheet, as a reference written
/// at the formula's cell, and whether it moves as each of them does: an
/// edge whose every end is held by a `$`, or none is, is held or moves with
/// them; a whole column's rows and a whole row's columns are whole.
fn spanned(ends: &[&RefExpr]) -> (Option<RefExpr>, bool) {
    let Some(first) = ends.first() else {
        return (None, false);
    };
    let same_sheet = |r: &RefExpr| {
        r.last_sheet.is_none()
            && match (&r.sheet, &first.sheet) {
                (None, None) => true,
                (Some(a), Some(b)) => a.eq_ignore_ascii_case(b),
                _ => false,
            }
    };
    if !ends.iter().all(|r| same_sheet(r)) {
        return (None, false);
    }
    // Each edge's number and whether a `$` holds it, from the corner of
    // each end that gives it.
    struct Edge {
        n: u32,
        held: Option<bool>,
        mixed: bool,
    }
    let fold = |edge: &mut Edge, n: u32, held: bool, pick: fn(u32, u32) -> u32| {
        edge.n = pick(edge.n, n);
        match edge.held {
            None => edge.held = Some(held),
            Some(h) if h != held => edge.mixed = true,
            _ => {}
        }
    };
    let mut top = Edge {
        n: u32::MAX,
        held: None,
        mixed: false,
    };
    let mut left = Edge {
        n: u32::MAX,
        held: None,
        mixed: false,
    };
    let mut bottom = Edge {
        n: 0,
        held: None,
        mixed: false,
    };
    let mut right = Edge {
        n: 0,
        held: None,
        mixed: false,
    };
    let (mut whole_rows, mut whole_columns) = (false, false);
    for r in ends {
        let c = &r.corners;
        match r.shape {
            Shape::Cell => {
                fold(&mut top, c.row1, c.row_abs1, u32::min);
                fold(&mut bottom, c.row1, c.row_abs1, u32::max);
                fold(&mut left, c.col1, c.col_abs1, u32::min);
                fold(&mut right, c.col1, c.col_abs1, u32::max);
            }
            Shape::Range => {
                for (row, held) in [(c.row1, c.row_abs1), (c.row2, c.row_abs2)] {
                    fold(&mut top, row, held, u32::min);
                    fold(&mut bottom, row, held, u32::max);
                }
                for (col, held) in [(c.col1, c.col_abs1), (c.col2, c.col_abs2)] {
                    fold(&mut left, col, held, u32::min);
                    fold(&mut right, col, held, u32::max);
                }
            }
            Shape::Column => {
                whole_rows = true;
                for (col, held) in [(c.col1, c.col_abs1), (c.col2, c.col_abs2)] {
                    fold(&mut left, col, held, u32::min);
                    fold(&mut right, col, held, u32::max);
                }
            }
            Shape::Row => {
                whole_columns = true;
                for (row, held) in [(c.row1, c.row_abs1), (c.row2, c.row_abs2)] {
                    fold(&mut top, row, held, u32::min);
                    fold(&mut bottom, row, held, u32::max);
                }
            }
            Shape::None => return (None, false),
        }
    }
    let rows_rigid = whole_rows || !(top.mixed || bottom.mixed);
    let columns_rigid = whole_columns || !(left.mixed || right.mixed);
    let held = |e: &Edge| e.held.unwrap_or(false);
    let (shape, corners, rect) = if whole_rows {
        // Every row: a whole column's span, whatever rows the others name;
        // and every column too where a whole row stands beside it.
        let (l, r, lh, rh) = if whole_columns {
            (1, MAX_COLUMN, true, true)
        } else {
            (left.n, right.n, held(&left), held(&right))
        };
        (
            Shape::Column,
            Corners {
                row1: 0,
                col1: l,
                row2: 0,
                col2: r,
                row_abs1: false,
                col_abs1: lh,
                row_abs2: false,
                col_abs2: rh,
            },
            (1, l, MAX_ROW, r),
        )
    } else if whole_columns {
        (
            Shape::Row,
            Corners {
                row1: top.n,
                col1: 0,
                row2: bottom.n,
                col2: 0,
                row_abs1: held(&top),
                col_abs1: false,
                row_abs2: held(&bottom),
                col_abs2: false,
            },
            (top.n, 1, bottom.n, MAX_COLUMN),
        )
    } else {
        (
            Shape::Range,
            Corners {
                row1: top.n,
                col1: left.n,
                row2: bottom.n,
                col2: right.n,
                row_abs1: held(&top),
                col_abs1: held(&left),
                row_abs2: held(&bottom),
                col_abs2: held(&right),
            },
            (top.n, left.n, bottom.n, right.n),
        )
    };
    let text = crate::sheet::A1Range {
        top: rect.0,
        left: rect.1,
        bottom: rect.2,
        right: rect.3,
    }
    .text();
    let written = match &first.sheet {
        Some(s) => format!("{}!{text}", crate::rows::sheet_prefix(s)),
        None => text,
    };
    let reach = RefExpr {
        sheet: first.sheet.clone(),
        last_sheet: None,
        top: rect.0,
        left: rect.1,
        bottom: rect.2,
        right: rect.3,
        written,
        shape,
        corners,
        found: None,
    };
    (Some(reach), rows_rigid && columns_rigid)
}

/// What an expression reads besides its literals: a reference, or a defined
/// name with the sheet written in front of it when there is one.
#[derive(Clone, Copy, Debug, PartialEq)]
pub enum Read<'e> {
    Ref(&'e RefExpr),
    Name {
        name: &'e str,
        sheet: Option<&'e str>,
    },
}

impl Expr {
    /// Every reference and every defined name the expression reads, in the
    /// order they are written (KERNEL.22: a shape's edges are read off its
    /// parse, as `graph::dependencies` reads them off the text).
    pub fn each_read<'e>(&'e self, f: &mut dyn FnMut(Read<'e>)) {
        let mut stack = vec![self];
        while let Some(e) = stack.pop() {
            match e {
                Expr::Ref(r) => f(Read::Ref(r)),
                Expr::Name { name, sheet } => f(Read::Name {
                    name,
                    sheet: sheet.as_deref(),
                }),
                // A range reads its reach, the plan's edges for it, and
                // whatever its ends read (KERNEL.8, slice f).
                Expr::Range(r) => {
                    if let Some(reach) = &r.reach {
                        f(Read::Ref(reach));
                    }
                    e.push_children(&mut stack);
                }
                other => other.push_children(&mut stack),
            }
        }
    }

    /// Whether the expression holds `:` between two expressions anywhere
    /// (KERNEL.8, slice f), the range operator Excel's shared formulas cannot
    /// hold (`PtgRange`).
    pub fn has_range(&self) -> bool {
        let mut stack = vec![self];
        while let Some(e) = stack.pop() {
            if matches!(e, Expr::Range(_)) {
                return true;
            }
            e.push_children(&mut stack);
        }
        false
    }

    /// The functions the expression calls, upper-cased, in the order they
    /// are written, repeats kept: [`calls`] over a parse.
    pub fn each_call<'e>(&'e self, f: &mut dyn FnMut(&'e str)) {
        let mut stack = vec![self];
        while let Some(e) = stack.pop() {
            if let Expr::Call(name, _) = e {
                f(name);
            }
            e.push_children(&mut stack);
        }
    }

    /// The walks' step (VLA.1): an expression's operands onto a stack, last
    /// first, so that popping visits them in the order they are written.
    /// The walks keep their own stack rather than the call stack, since a
    /// chain as long as Excel holds is a tree 4,095 deep down its left side,
    /// and a machine's plan reads every shape's references at its load.
    fn push_children<'e>(&'e self, stack: &mut Vec<&'e Expr>) {
        match self {
            Expr::Call(_, args) => stack.extend(args.iter().rev()),
            Expr::Neg(x) | Expr::Pos(x) | Expr::Percent(x) => stack.push(x),
            Expr::Binary(_, a, b) => {
                stack.push(b);
                stack.push(a);
            }
            Expr::Range(r) => {
                stack.push(&r.right);
                stack.push(&r.left);
            }
            Expr::Ref(_)
            | Expr::Name { .. }
            | Expr::Number(_)
            | Expr::Text(_)
            | Expr::Bool(_)
            | Expr::Error(_)
            | Expr::Empty
            | Expr::Array(_) => {}
        }
    }
}

#[derive(Clone, Debug, PartialEq)]
enum Tok {
    Num(f64),
    Str(String),
    Bool(bool),
    Err(ErrorKind),
    Ref(RefExpr),
    Name {
        name: String,
        sheet: Option<String>,
    },
    /// An identifier followed by `(`, upper-cased; the `(` follows as its
    /// own token.
    Func(String),
    Op(BinOp),
    Percent,
    LParen,
    RParen,
    Comma,
    /// An array constant's braces and its row separator, and `:` between
    /// two expressions (KERNEL.8, slice f).
    LBrace,
    RBrace,
    Semi,
    Colon,
    /// A construct this slice does not read, as written.
    Other(String),
}

/// The scanner's record as a reference expression, for the kinds that are
/// a rectangle: a cell, a range, a whole column, a whole row, on one sheet
/// or across a span. `None` for every other kind.
pub fn ref_expr(r: &FormulaRef) -> Option<RefExpr> {
    if !matches!(
        r.kind,
        Kind::Cell | Kind::Range | Kind::Column | Kind::Row | Kind::ThreeD
    ) {
        return None;
    }
    let (top, left, bottom, right) = match r.shape {
        Shape::Cell => (r.row1, r.col1, r.row1, r.col1),
        Shape::Range => (
            r.row1.min(r.row2),
            r.col1.min(r.col2),
            r.row1.max(r.row2),
            r.col1.max(r.col2),
        ),
        Shape::Column => (1, r.col1.min(r.col2), MAX_ROW, r.col1.max(r.col2)),
        Shape::Row => (r.row1.min(r.row2), 1, r.row1.max(r.row2), MAX_COLUMN),
        Shape::None => return None,
    };
    let (sheet, last_sheet) = if r.sheet_name.is_empty() {
        (None, None)
    } else {
        match r.sheet_name.split_once(':') {
            Some((first, last)) => (Some(first.to_string()), Some(last.to_string())),
            None => (Some(r.sheet_name.clone()), None),
        }
    };
    Some(RefExpr {
        sheet,
        last_sheet,
        top,
        left,
        bottom,
        right,
        written: r.written.clone(),
        shape: r.shape,
        corners: Corners {
            row1: r.row1,
            col1: r.col1,
            row2: r.row2,
            col2: r.col2,
            row_abs1: r.row_abs1,
            col_abs1: r.col_abs1,
            row_abs2: r.row_abs2,
            col_abs2: r.col_abs2,
        },
        found: None,
    })
}

fn tokenize(text: &str) -> Vec<Tok> {
    let s: Vec<char> = text.chars().collect();
    let refs: Vec<FormulaRef> = refers::scan(text)
        .into_iter()
        .filter(|r| r.kind != Kind::Unreadable)
        .collect();
    let mut toks = Vec::new();
    let mut i = usize::from(s.first() == Some(&'='));
    let mut next_ref = 0;
    while i < s.len() {
        while next_ref < refs.len() && refs[next_ref].pos < i {
            next_ref += 1;
        }
        if next_ref < refs.len() && refs[next_ref].pos == i {
            let r = &refs[next_ref];
            next_ref += 1;
            toks.push(match r.kind {
                Kind::Cell | Kind::Range | Kind::Column | Kind::Row | Kind::ThreeD => {
                    match ref_expr(r) {
                        Some(re) => Tok::Ref(re),
                        None => Tok::Other(r.written.clone()),
                    }
                }
                Kind::Name => Tok::Name {
                    name: r.part.clone(),
                    sheet: (!r.sheet_name.is_empty()).then(|| r.sheet_name.clone()),
                },
                Kind::Broken => Tok::Err(ErrorKind::Ref),
                Kind::Structured | Kind::External | Kind::Spill | Kind::Unreadable => {
                    Tok::Other(r.written.clone())
                }
            });
            i += r.span.max(1);
            continue;
        }
        let c = s[i];
        if c.is_whitespace() {
            i += 1;
        } else if c == '"' {
            let (text, end) = read_string(&s, i);
            toks.push(Tok::Str(text));
            i = end;
        } else if c.is_ascii_digit()
            || (c == '.' && s.get(i + 1).is_some_and(|d| d.is_ascii_digit()))
        {
            let (tok, end) = read_number(&s, i);
            toks.push(tok);
            i = end;
        } else if c == '#' {
            let mut k = i + 1;
            while k < s.len() && (s[k].is_ascii_alphanumeric() || s[k] == '/' || s[k] == '_') {
                k += 1;
            }
            if matches!(s.get(k), Some('!') | Some('?')) {
                k += 1;
            }
            let tok: String = s[i..k].iter().collect();
            toks.push(match ErrorKind::parse(&tok) {
                Some(kind) if kind.is_literal() => Tok::Err(kind),
                _ => Tok::Other(tok),
            });
            i = k;
        } else if c.is_ascii_alphabetic() || c == '_' {
            let mut k = i;
            while k < s.len() && is_ident_char(s[k]) {
                k += 1;
            }
            let tok: String = s[i..k].iter().collect();
            if s.get(k) == Some(&'(') {
                toks.push(Tok::Func(tok.to_ascii_uppercase()));
            } else {
                match tok.to_ascii_lowercase().as_str() {
                    "true" => toks.push(Tok::Bool(true)),
                    "false" => toks.push(Tok::Bool(false)),
                    _ => toks.push(Tok::Other(tok)),
                }
            }
            i = k;
        } else {
            let two: String = s[i..(i + 2).min(s.len())].iter().collect();
            let (tok, len) = match two.as_str() {
                "<>" => (Tok::Op(BinOp::Ne), 2),
                "<=" => (Tok::Op(BinOp::Le), 2),
                ">=" => (Tok::Op(BinOp::Ge), 2),
                _ => match c {
                    '<' => (Tok::Op(BinOp::Lt), 1),
                    '>' => (Tok::Op(BinOp::Gt), 1),
                    '=' => (Tok::Op(BinOp::Eq), 1),
                    '+' => (Tok::Op(BinOp::Add), 1),
                    '-' => (Tok::Op(BinOp::Sub), 1),
                    '*' => (Tok::Op(BinOp::Mul), 1),
                    '/' => (Tok::Op(BinOp::Div), 1),
                    '^' => (Tok::Op(BinOp::Pow), 1),
                    '&' => (Tok::Op(BinOp::Concat), 1),
                    '%' => (Tok::Percent, 1),
                    '(' => (Tok::LParen, 1),
                    ')' => (Tok::RParen, 1),
                    ',' => (Tok::Comma, 1),
                    '{' => (Tok::LBrace, 1),
                    '}' => (Tok::RBrace, 1),
                    ';' => (Tok::Semi, 1),
                    ':' => (Tok::Colon, 1),
                    other => (Tok::Other(other.to_string()), 1),
                },
            };
            toks.push(tok);
            i += len;
        }
    }
    toks
}

/// A string literal at `i` (its opening quote): the text with `""` read as
/// `"`, and the position after the closing quote, or the end of the text
/// when it never closes.
fn read_string(s: &[char], i: usize) -> (String, usize) {
    let mut out = String::new();
    let mut k = i + 1;
    while k < s.len() {
        if s[k] == '"' {
            if s.get(k + 1) == Some(&'"') {
                out.push('"');
                k += 2;
            } else {
                return (out, k + 1);
            }
        } else {
            out.push(s[k]);
            k += 1;
        }
    }
    (out, k)
}

/// A number literal at `i`: digits, an optional fraction, an optional
/// exponent when digits follow it. A literal no double holds is `#NUM!`.
fn read_number(s: &[char], i: usize) -> (Tok, usize) {
    let mut k = i;
    while k < s.len() && s[k].is_ascii_digit() {
        k += 1;
    }
    if k < s.len() && s[k] == '.' {
        k += 1;
        while k < s.len() && s[k].is_ascii_digit() {
            k += 1;
        }
    }
    if k < s.len() && (s[k] == 'e' || s[k] == 'E') {
        let mut j = k + 1;
        if j < s.len() && (s[j] == '+' || s[j] == '-') {
            j += 1;
        }
        let first = j;
        while j < s.len() && s[j].is_ascii_digit() {
            j += 1;
        }
        if j > first {
            k = j;
        }
    }
    let text: String = s[i..k].iter().collect();
    match text.parse::<f64>() {
        Ok(v) if v.is_finite() => (Tok::Num(v), k),
        _ => (Tok::Err(ErrorKind::Num), k),
    }
}

/// The deepest a formula may nest, Excel's own limit (VLA.1): each pair of
/// parentheses, each call's arguments, each sign before an operand and each
/// `%` after one counts a level, so that evaluation, which recurses once a
/// level, never goes deeper. The deepest real formula nests 9 (Alonzo's
/// Sand, its Screen's rule; Life's nests 6). A chain of binary operators is
/// not nesting: it is read in a loop and evaluated in one ([`LENGTH_LIMIT`]
/// bounds it).
pub const NESTING_LIMIT: u32 = 64;

/// The longest a formula's text may be, its leading `=` aside: Excel's own
/// limit of 8,192 characters (VLA.1). It bounds a chain like `1+1+...+1` at
/// 4,096 links, whose tree is freed recursively.
pub const LENGTH_LIMIT: usize = 8192;

/// The label [`parse`] answers for a formula nested past [`NESTING_LIMIT`].
pub const TOO_DEEP: &str = "nesting deeper than 64";

/// The label [`parse`] answers for a formula longer than [`LENGTH_LIMIT`].
pub const TOO_LONG: &str = "longer than 8,192 characters";

/// The refusal for a formula cell whose text does not parse (VLA.1): one of
/// Excel's two limits by its own id, any other label as the construct this
/// version does not read. An engine's load raises it; `frazaro calc` shows
/// the label in the cell instead.
pub fn unread_refusal(cell: &str, construct: &str) -> Refusal {
    match construct {
        TOO_DEEP => raise(
            "calc-formula-too-deep",
            &[("cell", cell), ("limit", &NESTING_LIMIT.to_string())],
        ),
        TOO_LONG => raise(
            "calc-formula-too-long",
            &[("cell", cell), ("limit", "8,192")],
        ),
        _ => raise(
            "calc-construct-not-read",
            &[("cell", cell), ("construct", construct)],
        ),
    }
}

struct Parser {
    toks: Vec<Tok>,
    at: usize,
    /// The nesting levels open where the parser stands (VLA.1).
    depth: u32,
}

impl Parser {
    /// `f` one nesting level deeper, refused past [`NESTING_LIMIT`].
    fn nested<T>(&mut self, f: impl FnOnce(&mut Self) -> Result<T, String>) -> Result<T, String> {
        if self.depth >= NESTING_LIMIT {
            return Err(TOO_DEEP.to_string());
        }
        self.depth += 1;
        let r = f(self);
        self.depth -= 1;
        r
    }

    fn peek(&self) -> Option<&Tok> {
        self.toks.get(self.at)
    }

    fn take(&mut self) -> Option<Tok> {
        let t = self.toks.get(self.at).cloned();
        if t.is_some() {
            self.at += 1;
        }
        t
    }

    fn peek_op(&self) -> Option<BinOp> {
        match self.peek() {
            Some(Tok::Op(op)) => Some(*op),
            _ => None,
        }
    }

    /// What the token spells, for the label of a construct not read.
    fn label(t: &Tok) -> String {
        match t {
            Tok::Other(s) => s.clone(),
            Tok::Op(op) => match op {
                BinOp::Add => "+",
                BinOp::Sub => "-",
                BinOp::Mul => "*",
                BinOp::Div => "/",
                BinOp::Pow => "^",
                BinOp::Concat => "&",
                BinOp::Eq => "=",
                BinOp::Ne => "<>",
                BinOp::Lt => "<",
                BinOp::Gt => ">",
                BinOp::Le => "<=",
                BinOp::Ge => ">=",
            }
            .to_string(),
            Tok::Percent => "%".to_string(),
            Tok::LParen => "(".to_string(),
            Tok::RParen => ")".to_string(),
            Tok::Comma => ",".to_string(),
            Tok::LBrace => "{".to_string(),
            Tok::RBrace => "}".to_string(),
            Tok::Semi => ";".to_string(),
            Tok::Colon => ":".to_string(),
            Tok::Num(v) => format!("{v}"),
            Tok::Str(s) => format!("\"{s}\""),
            Tok::Bool(b) => if *b { "TRUE" } else { "FALSE" }.to_string(),
            Tok::Err(k) => k.text().to_string(),
            Tok::Ref(r) => r.written.clone(),
            Tok::Name { name, .. } => name.clone(),
            Tok::Func(name) => name.clone(),
        }
    }

    fn starts_operand(t: &Tok) -> bool {
        matches!(
            t,
            Tok::Num(_)
                | Tok::Str(_)
                | Tok::Bool(_)
                | Tok::Err(_)
                | Tok::Ref(_)
                | Tok::Name { .. }
                | Tok::Func(_)
                | Tok::LParen
                | Tok::LBrace
        )
    }

    fn comparison(&mut self) -> Result<Expr, String> {
        let mut left = self.concat()?;
        while let Some(op) = self.peek_op() {
            if !matches!(
                op,
                BinOp::Eq | BinOp::Ne | BinOp::Lt | BinOp::Gt | BinOp::Le | BinOp::Ge
            ) {
                break;
            }
            self.at += 1;
            let right = self.concat()?;
            left = Expr::Binary(op, Box::new(left), Box::new(right));
        }
        Ok(left)
    }

    fn concat(&mut self) -> Result<Expr, String> {
        let mut left = self.additive()?;
        while self.peek_op() == Some(BinOp::Concat) {
            self.at += 1;
            let right = self.additive()?;
            left = Expr::Binary(BinOp::Concat, Box::new(left), Box::new(right));
        }
        Ok(left)
    }

    fn additive(&mut self) -> Result<Expr, String> {
        let mut left = self.term()?;
        while let Some(op @ (BinOp::Add | BinOp::Sub)) = self.peek_op() {
            self.at += 1;
            let right = self.term()?;
            left = Expr::Binary(op, Box::new(left), Box::new(right));
        }
        Ok(left)
    }

    fn term(&mut self) -> Result<Expr, String> {
        let mut left = self.power()?;
        while let Some(op @ (BinOp::Mul | BinOp::Div)) = self.peek_op() {
            self.at += 1;
            let right = self.power()?;
            left = Expr::Binary(op, Box::new(left), Box::new(right));
        }
        Ok(left)
    }

    fn power(&mut self) -> Result<Expr, String> {
        let mut left = self.unary()?;
        while self.peek_op() == Some(BinOp::Pow) {
            self.at += 1;
            let right = self.unary()?;
            left = Expr::Binary(BinOp::Pow, Box::new(left), Box::new(right));
        }
        Ok(left)
    }

    fn unary(&mut self) -> Result<Expr, String> {
        match self.peek_op() {
            Some(BinOp::Sub) => {
                self.at += 1;
                Ok(Expr::Neg(Box::new(self.nested(Self::unary)?)))
            }
            Some(BinOp::Add) => {
                self.at += 1;
                Ok(Expr::Pos(Box::new(self.nested(Self::unary)?)))
            }
            _ => self.postfix(),
        }
    }

    // Every level of nesting passes through this frame, and VLA.1's limits
    // are pinned on 1 MiB in a debug build, which gives each temporary a
    // slot of its own. A `?` holds temporaries of its own: written with
    // them, the `:` check below made a level 146 bytes deeper; matched by
    // hand, the two calls leave a level 44 bytes shallower than before the
    // check (KERNEL.8, slice f).
    #[allow(clippy::question_mark)]
    fn postfix(&mut self) -> Result<Expr, String> {
        let mut e = match self.primary() {
            Ok(e) => e,
            Err(why) => return Err(why),
        };
        // `:` binds tightest of all (KERNEL.8, slice f): its loop a frame of
        // its own, entered only where one follows, writing into the
        // primary's place.
        if self.peek() == Some(&Tok::Colon) {
            if let Err(why) = self.ranged(&mut e) {
                return Err(why);
            }
        }
        let mut percents = 0;
        while self.peek() == Some(&Tok::Percent) {
            // Each `%` wraps what came before, a level its evaluation
            // recurses through (VLA.1).
            if self.depth + percents >= NESTING_LIMIT {
                return Err(TOO_DEEP.to_string());
            }
            percents += 1;
            self.at += 1;
            e = Expr::Percent(Box::new(e));
        }
        if let Some(t) = self.peek() {
            if Self::starts_operand(t) {
                // Two operands with nothing between them: Excel's
                // intersection operator, a space.
                return Err("intersection".to_string());
            }
        }
        Ok(e)
    }

    /// `:` after a primary read to another (KERNEL.8, slice f), the
    /// tightest binding of all, the range written into the primary's place:
    /// each `:` a level its evaluation recurses through (VLA.1). Two
    /// references side by side with `:` between them the scanner reads as
    /// one range; this is the `:` it leaves, where an end is a call or a
    /// name.
    fn ranged(&mut self, e: &mut Expr) -> Result<(), String> {
        let mut ranges = 0;
        while self.peek() == Some(&Tok::Colon) {
            if self.depth + ranges >= NESTING_LIMIT {
                return Err(TOO_DEEP.to_string());
            }
            ranges += 1;
            self.at += 1;
            let right = self.primary()?;
            let left = std::mem::replace(e, Expr::Empty);
            *e = Expr::Range(Box::new(RangeExpr::new(left, right)));
        }
        Ok(())
    }

    /// An array constant's values after its `{`, through its `}` (KERNEL.8,
    /// slice f): `,` between the values of a row and `;` between rows, every
    /// row as long as the first, each value a number with its sign, a text,
    /// a truth value or an error, as Excel holds one; anything else, a
    /// reference among them, is a construct the formula cannot hold.
    fn array_constant(&mut self) -> Result<Expr, String> {
        let mut rows: Vec<Vec<Value>> = vec![Vec::new()];
        loop {
            let v = self.array_value()?;
            if let Some(row) = rows.last_mut() {
                row.push(v);
            }
            match self.take() {
                Some(Tok::Comma) => {}
                Some(Tok::Semi) => rows.push(Vec::new()),
                Some(Tok::RBrace) => break,
                Some(other) => return Err(Self::label(&other)),
                None => return Err("end of formula".to_string()),
            }
        }
        let width = rows[0].len();
        if rows.iter().any(|r| r.len() != width) {
            return Err("an array constant of rows unlike".to_string());
        }
        let height = rows.len();
        let values: Vec<Value> = rows.into_iter().flatten().collect();
        Ok(Expr::Array(Array::new(height as u32, width as u32, values)))
    }

    /// One value of an array constant.
    fn array_value(&mut self) -> Result<Value, String> {
        let signed = |p: &mut Self, sign: f64| match p.take() {
            // No negative zero in an array, as in a cell.
            Some(Tok::Num(v)) => Ok(Value::Number(if v == 0.0 { 0.0 } else { sign * v })),
            Some(other) => Err(Self::label(&other)),
            None => Err("end of formula".to_string()),
        };
        match self.take() {
            Some(Tok::Num(v)) => Ok(Value::Number(v)),
            Some(Tok::Op(BinOp::Sub)) => signed(self, -1.0),
            Some(Tok::Op(BinOp::Add)) => signed(self, 1.0),
            Some(Tok::Str(s)) => Ok(Value::Text(s)),
            Some(Tok::Bool(b)) => Ok(Value::Bool(b)),
            Some(Tok::Err(k)) => Ok(Value::Error(k)),
            Some(other) => Err(Self::label(&other)),
            None => Err("end of formula".to_string()),
        }
    }

    fn primary(&mut self) -> Result<Expr, String> {
        let Some(t) = self.take() else {
            return Err("end of formula".to_string());
        };
        match t {
            Tok::Num(v) => Ok(Expr::Number(v)),
            Tok::Str(s) => Ok(Expr::Text(s)),
            Tok::Bool(b) => Ok(Expr::Bool(b)),
            Tok::Err(k) => Ok(Expr::Error(k)),
            Tok::LBrace => self.array_constant(),
            Tok::Ref(r) => Ok(Expr::Ref(r)),
            Tok::Name { name, sheet } => Ok(Expr::Name { name, sheet }),
            Tok::Func(name) => {
                if self.take() != Some(Tok::LParen) {
                    return Err(name);
                }
                self.nested(|p| p.call(name))
            }
            Tok::LParen => {
                let e = self.nested(Self::comparison)?;
                match self.take() {
                    Some(Tok::RParen) => Ok(e),
                    Some(Tok::Comma) => Err("union".to_string()),
                    Some(other) => Err(Self::label(&other)),
                    None => Err("end of formula".to_string()),
                }
            }
            other => Err(Self::label(&other)),
        }
    }

    /// A call's arguments after its `(`, through its `)`, a level deeper.
    fn call(&mut self, name: String) -> Result<Expr, String> {
        let mut args = Vec::new();
        if self.peek() == Some(&Tok::RParen) {
            self.at += 1;
            return Ok(Expr::Call(name, args));
        }
        loop {
            match self.peek() {
                Some(Tok::Comma) | Some(Tok::RParen) => args.push(Expr::Empty),
                _ => args.push(self.comparison()?),
            }
            match self.take() {
                Some(Tok::Comma) => continue,
                Some(Tok::RParen) => break,
                Some(other) => return Err(Self::label(&other)),
                None => return Err("end of formula".to_string()),
            }
        }
        Ok(Expr::Call(name, args))
    }
}

/// A formula's text, with or without its leading `=`, as an expression; or
/// the construct this slice does not read, as written, for the label: one
/// of [`TOO_LONG`] and [`TOO_DEEP`] past Excel's two limits (VLA.1).
pub fn parse(text: &str) -> Result<Expr, String> {
    if text.strip_prefix('=').unwrap_or(text).chars().count() > LENGTH_LIMIT {
        return Err(TOO_LONG.to_string());
    }
    let mut p = Parser {
        toks: tokenize(text),
        at: 0,
        depth: 0,
    };
    if p.toks.is_empty() {
        return Err("end of formula".to_string());
    }
    let e = p.comparison()?;
    match p.peek() {
        None => Ok(e),
        Some(t) => Err(Parser::label(t)),
    }
}

/// The functions a formula's text calls, upper-cased, in order, repeats
/// kept: what a histogram over a workbook counts (`frazaro reflect
/// --functions`), read before anything is evaluated.
pub fn calls(text: &str) -> Vec<String> {
    tokenize(text)
        .into_iter()
        .filter_map(|t| match t {
            Tok::Func(name) => Some(name),
            _ => None,
        })
        .collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    fn cell(written: &str, row: u32, col: u32) -> Expr {
        Expr::Ref(RefExpr {
            sheet: None,
            last_sheet: None,
            top: row,
            left: col,
            bottom: row,
            right: col,
            written: written.to_string(),
            shape: Shape::Cell,
            corners: Corners {
                row1: row,
                col1: col,
                row2: 0,
                col2: 0,
                row_abs1: false,
                col_abs1: false,
                row_abs2: false,
                col_abs2: false,
            },
            found: None,
        })
    }

    #[test]
    fn literals_and_references_are_read() {
        assert_eq!(parse("=5"), Ok(Expr::Number(5.0)));
        assert_eq!(parse("1.5e3"), Ok(Expr::Number(1500.0)));
        assert_eq!(parse(".5"), Ok(Expr::Number(0.5)));
        assert_eq!(parse("=\"a\"\"b\""), Ok(Expr::Text("a\"b".to_string())));
        assert_eq!(parse("=TRUE"), Ok(Expr::Bool(true)));
        assert_eq!(parse("=false"), Ok(Expr::Bool(false)));
        assert_eq!(parse("=#N/A"), Ok(Expr::Error(ErrorKind::NA)));
        assert_eq!(parse("=#REF!"), Ok(Expr::Error(ErrorKind::Ref)));
        assert_eq!(parse("=B2"), Ok(cell("B2", 2, 2)));
        assert_eq!(
            parse("=Rate"),
            Ok(Expr::Name {
                name: "Rate".to_string(),
                sheet: None
            })
        );
        assert_eq!(
            parse("=Model!Local"),
            Ok(Expr::Name {
                name: "Local".to_string(),
                sheet: Some("Model".to_string())
            })
        );
        let r = match parse("='Q1 Data'!A:A").unwrap() {
            Expr::Ref(r) => r,
            other => panic!("{other:?}"),
        };
        assert_eq!(r.sheet.as_deref(), Some("Q1 Data"));
        assert_eq!((r.top, r.left, r.bottom, r.right), (1, 1, MAX_ROW, 1));
        let span = match parse("=Model:Scratch!B1").unwrap() {
            Expr::Ref(r) => r,
            other => panic!("{other:?}"),
        };
        assert_eq!(span.sheet.as_deref(), Some("Model"));
        assert_eq!(span.last_sheet.as_deref(), Some("Scratch"));
        assert_eq!((span.top, span.left, span.bottom, span.right), (1, 2, 1, 2));
        assert!(
            !span.is_cell(),
            "a span is one cell on each of its sheets, not one cell"
        );
    }

    #[test]
    fn precedence_is_excel_s() {
        // -2^2 is 4: the sign binds tighter than the power.
        assert_eq!(
            parse("=-2^2"),
            Ok(Expr::Binary(
                BinOp::Pow,
                Box::new(Expr::Neg(Box::new(Expr::Number(2.0)))),
                Box::new(Expr::Number(2.0))
            ))
        );
        // 1+2*3 groups the product; 2^3^2 groups to the left.
        assert_eq!(
            parse("=1+2*3"),
            Ok(Expr::Binary(
                BinOp::Add,
                Box::new(Expr::Number(1.0)),
                Box::new(Expr::Binary(
                    BinOp::Mul,
                    Box::new(Expr::Number(2.0)),
                    Box::new(Expr::Number(3.0))
                ))
            ))
        );
        assert_eq!(
            parse("=2^3^2"),
            Ok(Expr::Binary(
                BinOp::Pow,
                Box::new(Expr::Binary(
                    BinOp::Pow,
                    Box::new(Expr::Number(2.0)),
                    Box::new(Expr::Number(3.0))
                )),
                Box::new(Expr::Number(2.0))
            ))
        );
        // & below arithmetic, the comparisons below &.
        assert_eq!(
            parse("=1&2=\"12\""),
            Ok(Expr::Binary(
                BinOp::Eq,
                Box::new(Expr::Binary(
                    BinOp::Concat,
                    Box::new(Expr::Number(1.0)),
                    Box::new(Expr::Number(2.0))
                )),
                Box::new(Expr::Text("12".to_string()))
            ))
        );
        assert_eq!(
            parse("=50%"),
            Ok(Expr::Percent(Box::new(Expr::Number(50.0))))
        );
        assert_eq!(
            parse("=(1+2)*3").map(|e| matches!(e, Expr::Binary(BinOp::Mul, ..))),
            Ok(true)
        );
    }

    #[test]
    fn calls_take_their_arguments_empty_ones_included() {
        assert_eq!(
            parse("=IF(B2>3,\"big\",\"small\")"),
            Ok(Expr::Call(
                "IF".to_string(),
                vec![
                    Expr::Binary(
                        BinOp::Gt,
                        Box::new(cell("B2", 2, 2)),
                        Box::new(Expr::Number(3.0))
                    ),
                    Expr::Text("big".to_string()),
                    Expr::Text("small".to_string()),
                ]
            ))
        );
        assert_eq!(
            parse("=sum(A1,,2)"),
            Ok(Expr::Call(
                "SUM".to_string(),
                vec![cell("A1", 1, 1), Expr::Empty, Expr::Number(2.0)]
            ))
        );
        assert_eq!(parse("=PI()"), Ok(Expr::Call("PI".to_string(), vec![])));
        assert_eq!(parse("=row()"), Ok(Expr::Call("ROW".to_string(), vec![])));
        assert_eq!(
            calls("=IF(SUM(A1:A3)>0,INDIRECT(\"B1\"),offset(A1,1,0))"),
            vec!["IF", "SUM", "INDIRECT", "OFFSET"]
        );
        assert_eq!(calls("=\"SUM(\"&A1"), Vec::<String>::new());
    }

    #[test]
    fn an_array_constant_and_a_range_between_expressions_parse() {
        // KERNEL.8, slice f: an array constant as the array it holds, each
        // value with its sign, every row as long as the first; `:` between
        // two expressions a range, its reach the rectangle spanning what
        // its ends may answer, moving with the cell where every end's `$`
        // marks agree, unbounded at a name; `:` above a sign; `#CALC!` no
        // literal.
        let Ok(Expr::Array(a)) = parse("={1,-2;\"a\",TRUE}") else {
            panic!("an array constant");
        };
        assert_eq!((a.rows(), a.columns()), (2, 2));
        assert_eq!(
            a.values(),
            &[
                Value::Number(1.0),
                Value::Number(-2.0),
                Value::Text("a".to_string()),
                Value::Bool(true)
            ]
        );
        assert!(matches!(parse("={#N/A,-0}"), Ok(Expr::Array(_))));
        assert_eq!(
            parse("={1,2;3}"),
            Err("an array constant of rows unlike".to_string())
        );
        assert_eq!(parse("={A1}"), Err("A1".to_string()));
        let Ok(Expr::Range(r)) = parse("=A1:INDEX(C1:C10,5)") else {
            panic!("a range");
        };
        let reach = r.reach.as_ref().unwrap();
        assert_eq!(
            (reach.top, reach.left, reach.bottom, reach.right),
            (1, 1, 10, 3)
        );
        assert!(r.rigid);
        let Ok(Expr::Range(r)) = parse("=$A$1:INDEX(C1:C10,5)") else {
            panic!("a range");
        };
        assert!(!r.rigid, "an edge held by a $ at one end and not the other");
        let Ok(Expr::Range(r)) = parse("=A1:Rate") else {
            panic!("a range");
        };
        assert!(r.reach.is_none(), "a name no plan can bound");
        let Ok(Expr::Range(r)) = parse("=A5:INDEX(B:B,3)") else {
            panic!("a range");
        };
        let reach = r.reach.as_ref().unwrap();
        assert_eq!(reach.shape, Shape::Column);
        assert_eq!((reach.left, reach.right), (1, 2));
        assert!(
            matches!(parse("=-A1:INDEX(B:B,3)"), Ok(Expr::Neg(x)) if matches!(*x, Expr::Range(_)))
        );
        assert_eq!(parse("=#CALC!"), Err("#CALC!".to_string()));
    }

    #[test]
    fn a_construct_this_slice_does_not_read_is_named_as_written() {
        assert_eq!(parse("=@A1:A3"), Err("@".to_string()));
        assert_eq!(parse("=A1 B1"), Err("intersection".to_string()));
        assert_eq!(parse("=SUM((A1,B2))"), Err("union".to_string()));
        assert_eq!(parse("=A1#"), Err("A1#".to_string()));
        assert_eq!(
            parse("=SUM(Sales[Amount])"),
            Err("Sales[Amount]".to_string())
        );
        assert_eq!(
            parse("=[Rates.xlsx]Sheet1!A1"),
            Err("[Rates.xlsx]Sheet1!A1".to_string())
        );
        assert_eq!(parse("=#SPILL!"), Err("#SPILL!".to_string()));
        assert_eq!(parse("="), Err("end of formula".to_string()));
        assert_eq!(parse("=1+"), Err("end of formula".to_string()));
        assert_eq!(parse("=)"), Err(")".to_string()));
    }

    #[test]
    fn excel_s_two_limits_are_labels_of_their_own() {
        // VLA.1: 64 levels of nesting parse and a 65th is TOO_DEEP, whether
        // the levels are parentheses, a call's arguments, signs or percents.
        let wrap =
            |n: usize, open: &str, close: &str| format!("={}1{}", open.repeat(n), close.repeat(n));
        for (open, close) in [("(", ")"), ("ABS(", ")"), ("-", ""), ("+", ""), ("", "%")] {
            assert!(
                parse(&wrap(64, open, close)).is_ok(),
                "{open}...{close} at 64"
            );
            assert_eq!(
                parse(&wrap(65, open, close)),
                Err(TOO_DEEP.to_string()),
                "{open}...{close} at 65"
            );
        }
        // 8,192 characters after the `=` parse, a chain of 4,095 links; one
        // more character is TOO_LONG.
        let chain = format!("={}11", "1+".repeat(4095));
        assert_eq!(chain.len() - 1, LENGTH_LIMIT);
        assert!(parse(&chain).is_ok());
        assert_eq!(parse(&format!("{chain}1")), Err(TOO_LONG.to_string()));
    }
}
