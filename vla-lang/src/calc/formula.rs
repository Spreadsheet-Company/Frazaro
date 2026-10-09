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
//! `/`; `^`, left to right; a sign, so that `-2^2` is 4; `%`.
//!
//! A construct this slice does not read is not an error: the parse answers
//! with the construct as written, and the cell shows `not computed here`
//! naming it, the honest static label of `HORIZON.md` section 12.6. Those
//! are: an array constant (`{1,2;3,4}`), implicit intersection by `@` or by
//! a space between two references, a union in parentheses, a spill range
//! (`A1#`), a structured reference (`Sales[Amount]`), a link into another
//! workbook (`[Rates.xlsx]Sheet1!A1`), and a range where one value is
//! wanted (`=B1:B2*2`), which a legacy formula intersects and a dynamic one
//! spills; the reader cannot tell the two apart from the text alone, so
//! neither is claimed (`KERNEL.8` takes them with the `@` operator).

use crate::refers::{self, is_ident_char, FormulaRef, Kind, Shape};
use crate::sheet::{MAX_COLUMN, MAX_ROW};

use super::value::ErrorKind;

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
        if d_row == 0 && d_col == 0 {
            return Some((self.top, self.left, self.bottom, self.right));
        }
        let c = &self.corners;
        let row = |n: u32, abs: bool| shifted(n, abs, d_row, MAX_ROW);
        let col = |n: u32, abs: bool| shifted(n, abs, d_col, MAX_COLUMN);
        match self.shape {
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

/// A parsed formula.
#[derive(Clone, Debug, PartialEq)]
pub enum Expr {
    Number(f64),
    Text(String),
    Bool(bool),
    Error(ErrorKind),
    /// An argument left out, `IF(A1,,B1)`.
    Empty,
    Ref(RefExpr),
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
        match self {
            Expr::Ref(r) => f(Read::Ref(r)),
            Expr::Name { name, sheet } => f(Read::Name {
                name,
                sheet: sheet.as_deref(),
            }),
            Expr::Call(_, args) => args.iter().for_each(|a| a.each_read(f)),
            Expr::Neg(x) | Expr::Pos(x) | Expr::Percent(x) => x.each_read(f),
            Expr::Binary(_, a, b) => {
                a.each_read(f);
                b.each_read(f);
            }
            Expr::Number(_) | Expr::Text(_) | Expr::Bool(_) | Expr::Error(_) | Expr::Empty => {}
        }
    }

    /// The functions the expression calls, upper-cased, in the order they
    /// are written, repeats kept: [`calls`] over a parse.
    pub fn each_call<'e>(&'e self, f: &mut dyn FnMut(&'e str)) {
        match self {
            Expr::Call(name, args) => {
                f(name);
                args.iter().for_each(|a| a.each_call(f));
            }
            Expr::Neg(x) | Expr::Pos(x) | Expr::Percent(x) => x.each_call(f),
            Expr::Binary(_, a, b) => {
                a.each_call(f);
                b.each_call(f);
            }
            Expr::Ref(_)
            | Expr::Name { .. }
            | Expr::Number(_)
            | Expr::Text(_)
            | Expr::Bool(_)
            | Expr::Error(_)
            | Expr::Empty => {}
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
                Some(kind) => Tok::Err(kind),
                None => Tok::Other(tok),
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

struct Parser {
    toks: Vec<Tok>,
    at: usize,
}

impl Parser {
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
                Ok(Expr::Neg(Box::new(self.unary()?)))
            }
            Some(BinOp::Add) => {
                self.at += 1;
                Ok(Expr::Pos(Box::new(self.unary()?)))
            }
            _ => self.postfix(),
        }
    }

    fn postfix(&mut self) -> Result<Expr, String> {
        let mut e = self.primary()?;
        while self.peek() == Some(&Tok::Percent) {
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

    fn primary(&mut self) -> Result<Expr, String> {
        let Some(t) = self.take() else {
            return Err("end of formula".to_string());
        };
        match t {
            Tok::Num(v) => Ok(Expr::Number(v)),
            Tok::Str(s) => Ok(Expr::Text(s)),
            Tok::Bool(b) => Ok(Expr::Bool(b)),
            Tok::Err(k) => Ok(Expr::Error(k)),
            Tok::Ref(r) => Ok(Expr::Ref(r)),
            Tok::Name { name, sheet } => Ok(Expr::Name { name, sheet }),
            Tok::Func(name) => {
                if self.take() != Some(Tok::LParen) {
                    return Err(name);
                }
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
            Tok::LParen => {
                let e = self.comparison()?;
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
}

/// A formula's text, with or without its leading `=`, as an expression; or
/// the construct this slice does not read, as written, for the label.
pub fn parse(text: &str) -> Result<Expr, String> {
    let mut p = Parser {
        toks: tokenize(text),
        at: 0,
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
    fn a_construct_this_slice_does_not_read_is_named_as_written() {
        assert_eq!(parse("={1,2;3,4}"), Err("{".to_string()));
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
}
