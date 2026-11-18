//! The file format's names for Excel's newer functions (PORT.7, slice 7b).
//!
//! New ground, said plainly: the reference writes a formula through
//! `Range.Formula2`, and Excel's object model spells the file for it. In the
//! file, a function added after Excel 2007 is stored under a prefix,
//! `_xlfn.` (and `_xlfn._xlws.` for two of the dynamic-array functions), and
//! a formula written without it shows `#NAME?` until a person re-enters the
//! cell. Excel strips the prefix again when it shows the formula, so a user
//! never sees it; Sheets, Calc and Numbers strip it too. The table is the
//! "future functions" list of the format's documentation (MS-XLSX), as far
//! as this writer knows it; a function missing from it shows up as `#NAME?`
//! in a built workbook, and the fix is one more row here.
//!
//! The names `LET` binds and `LAMBDA` takes are stored under `_xlpm.` where
//! they are bound and wherever they are read (KERNEL.29,
//! [`prefix_parameters`]): a name the file holds bare reads as a defined name,
//! which Excel answers `#NAME?`. Excel strips that prefix too, and so does the
//! reader (`strip_future_prefixes`).

pub use super::strip_future_prefixes;
use super::{is_name_char, is_name_start};
use crate::refers;

/// Functions stored as `_xlfn.<NAME>`.
const XLFN: &[&str] = &[
    // Excel 2010
    "AGGREGATE",
    "BETA.DIST",
    "BETA.INV",
    "BINOM.DIST",
    "BINOM.INV",
    "CEILING.PRECISE",
    "CHISQ.DIST",
    "CHISQ.DIST.RT",
    "CHISQ.INV",
    "CHISQ.INV.RT",
    "CHISQ.TEST",
    "CONFIDENCE.NORM",
    "CONFIDENCE.T",
    "COVARIANCE.P",
    "COVARIANCE.S",
    "ERF.PRECISE",
    "ERFC.PRECISE",
    "EXPON.DIST",
    "F.DIST",
    "F.DIST.RT",
    "F.INV",
    "F.INV.RT",
    "F.TEST",
    "FLOOR.PRECISE",
    "GAMMA.DIST",
    "GAMMA.INV",
    "GAMMALN.PRECISE",
    "HYPGEOM.DIST",
    "ISO.CEILING",
    "LOGNORM.DIST",
    "LOGNORM.INV",
    "MODE.MULT",
    "MODE.SNGL",
    "NEGBINOM.DIST",
    "NETWORKDAYS.INTL",
    "NORM.DIST",
    "NORM.INV",
    "NORM.S.DIST",
    "NORM.S.INV",
    "PERCENTILE.EXC",
    "PERCENTILE.INC",
    "PERCENTRANK.EXC",
    "PERCENTRANK.INC",
    "POISSON.DIST",
    "QUARTILE.EXC",
    "QUARTILE.INC",
    "RANK.AVG",
    "RANK.EQ",
    "STDEV.P",
    "STDEV.S",
    "T.DIST",
    "T.DIST.2T",
    "T.DIST.RT",
    "T.INV",
    "T.INV.2T",
    "T.TEST",
    "VAR.P",
    "VAR.S",
    "WEIBULL.DIST",
    "WORKDAY.INTL",
    "Z.TEST",
    // Excel 2013
    "ACOT",
    "ACOTH",
    "ARABIC",
    "BASE",
    "BINOM.DIST.RANGE",
    "BITAND",
    "BITLSHIFT",
    "BITOR",
    "BITRSHIFT",
    "BITXOR",
    "CEILING.MATH",
    "COMBINA",
    "COT",
    "COTH",
    "CSC",
    "CSCH",
    "DAYS",
    "DECIMAL",
    "ENCODEURL",
    "FILTERXML",
    "FLOOR.MATH",
    "FORMULATEXT",
    "GAMMA",
    "GAUSS",
    "IFNA",
    "IMCOSH",
    "IMCOT",
    "IMCSC",
    "IMCSCH",
    "IMSEC",
    "IMSECH",
    "IMSINH",
    "IMTAN",
    "ISFORMULA",
    "ISOWEEKNUM",
    "MUNIT",
    "NUMBERVALUE",
    "PDURATION",
    "PERMUTATIONA",
    "PHI",
    "RRI",
    "SEC",
    "SECH",
    "SHEET",
    "SHEETS",
    "SKEW.P",
    "UNICHAR",
    "UNICODE",
    "WEBSERVICE",
    "XOR",
    // Excel 2016 and 2019
    "CONCAT",
    "FORECAST.ETS",
    "FORECAST.ETS.CONFINT",
    "FORECAST.ETS.SEASONALITY",
    "FORECAST.ETS.STAT",
    "FORECAST.LINEAR",
    "IFS",
    "MAXIFS",
    "MINIFS",
    "SWITCH",
    "TEXTJOIN",
    // Microsoft 365
    // Excel's `#`, as its file spells the operator (KERNEL.55).
    "ANCHORARRAY",
    "ARRAYTOTEXT",
    "BYCOL",
    "BYROW",
    "CHOOSECOLS",
    "CHOOSEROWS",
    "DETECTLANGUAGE",
    "DROP",
    "EXPAND",
    "GROUPBY",
    "HSTACK",
    "IMAGE",
    "ISOMITTED",
    "LAMBDA",
    "LET",
    "MAKEARRAY",
    "MAP",
    "PERCENTOF",
    "PIVOTBY",
    "RANDARRAY",
    "REDUCE",
    "REGEXEXTRACT",
    "REGEXREPLACE",
    "REGEXTEST",
    "SCAN",
    "SEQUENCE",
    // Excel's `@`, as its file spells the operator (KERNEL.8, slice b).
    "SINGLE",
    "SORTBY",
    "STOCKHISTORY",
    "TAKE",
    "TEXTAFTER",
    "TEXTBEFORE",
    "TEXTSPLIT",
    "TOCOL",
    "TOROW",
    "TRANSLATE",
    "TRIMRANGE",
    "UNIQUE",
    "VALUETOTEXT",
    "VSTACK",
    "WRAPCOLS",
    "WRAPROWS",
    "XLOOKUP",
    "XMATCH",
];

/// Functions stored as `_xlfn._xlws.<NAME>`.
const XLWS: &[&str] = &["FILTER", "SORT"];

/// Functions that can return a reference or an array, so that Excel 365
/// marks a legacy formula calling one with the implicit-intersection `@`
/// (`=@IFS(...)`, seen on the owner's first live pass of slice 7b) and
/// stores a `Formula2` entry of one as a dynamic-array formula. The writer
/// stores such a formula the same way (`sheet::ooxml`, the `cm` metadata),
/// so the formula bar reads as the add-in's would. The table is Excel's own
/// classification as far as this writer knows it: a function missing here
/// shows `@` in the bar and, filled over a range, computes one value per
/// cell; one wrongly here is a single-cell dynamic array, harmless in Excel
/// and a braced array in an older host. `IF`, `IFERROR` and the lookup
/// functions that return one value are not here, as old workbooks full of
/// them open without `@`.
const ARRAY_CAPABLE: &[&str] = &[
    "ANCHORARRAY",
    "BYCOL",
    "BYROW",
    "CHOOSE",
    "CHOOSECOLS",
    "CHOOSEROWS",
    "DROP",
    "EXPAND",
    "FILTER",
    "FREQUENCY",
    "GROUPBY",
    "GROWTH",
    "HSTACK",
    "IFS",
    "INDEX",
    "INDIRECT",
    "LAMBDA",
    "LET",
    "LINEST",
    "LOGEST",
    "LOOKUP",
    "MAKEARRAY",
    "MAP",
    "MINVERSE",
    "MMULT",
    "MODE.MULT",
    "MUNIT",
    "OFFSET",
    "PIVOTBY",
    "RANDARRAY",
    "REDUCE",
    "REGEXEXTRACT",
    "SCAN",
    "SEQUENCE",
    "SORT",
    "SORTBY",
    "SWITCH",
    "TAKE",
    "TEXTSPLIT",
    "TOCOL",
    "TOROW",
    "TRANSPOSE",
    "TREND",
    "TRIMRANGE",
    "UNIQUE",
    "VSTACK",
    "WRAPCOLS",
    "WRAPROWS",
    "XLOOKUP",
];

/// Whether a formula's text (without its `=`, prefixes applied or not)
/// calls a function that can return a reference or an array, and so must be
/// stored as a dynamic-array formula to read without `@`; or calls in place
/// what a call answers, `IF(c,SQ,PLUS1)(5)` (KERNEL.29), which Excel reads in
/// a dynamic-array formula alone, repairing the file where a legacy one holds
/// it (the third save of `scripts/recalc/lambda.xlsx`).
pub fn can_return_array(formula: &str) -> bool {
    let chars: Vec<char> = formula.chars().collect();
    let mut i = 0;
    while i < chars.len() {
        let c = chars[i];
        if c == '"' {
            i += 1;
            while i < chars.len() {
                if chars[i] == '"' {
                    if i + 1 < chars.len() && chars[i + 1] == '"' {
                        i += 2;
                        continue;
                    }
                    i += 1;
                    break;
                }
                i += 1;
            }
            continue;
        }
        if c == ')' && chars.get(i + 1) == Some(&'(') {
            return true;
        }
        if is_name_start(c) && (i == 0 || !is_name_char(chars[i - 1])) {
            let start = i;
            while i < chars.len() && is_name_char(chars[i]) {
                i += 1;
            }
            let mut j = i;
            while j < chars.len() && chars[j] == ' ' {
                j += 1;
            }
            if j < chars.len() && chars[j] == '(' {
                let name: String = chars[start..i]
                    .iter()
                    .collect::<String>()
                    .to_ascii_uppercase();
                let bare = name
                    .strip_prefix("_XLFN._XLWS.")
                    .or_else(|| name.strip_prefix("_XLFN."))
                    .unwrap_or(&name);
                if ARRAY_CAPABLE.contains(&bare) {
                    return true;
                }
            }
            continue;
        }
        i += 1;
    }
    false
}

/// The prefix a function name takes in the file, or `None` for one Excel
/// 2007 already had (or one this table does not know).
pub fn prefix_for(upper_name: &str) -> Option<&'static str> {
    if XLWS.contains(&upper_name) {
        Some("_xlfn._xlws.")
    } else if XLFN.contains(&upper_name) {
        Some("_xlfn.")
    } else {
        None
    }
}

/// A formula's text (without its leading `=`) with every call of a newer
/// function written under its prefix, in the upper case Excel stores. Text
/// inside a string literal is left alone; a name already prefixed is one
/// the table does not hold (`_xlfn.IFS` is not `IFS`), so it is left too.
/// A name is a call only when `(` follows it, so a defined name or a cell
/// reference spelled like a function is untouched.
pub fn prefix_future_functions(formula: &str) -> String {
    let chars: Vec<char> = formula.chars().collect();
    let mut out = String::with_capacity(formula.len() + 16);
    let mut i = 0;
    while i < chars.len() {
        let c = chars[i];
        if c == '"' {
            // A string literal, to its closing quote; a doubled quote stays
            // inside it.
            out.push(c);
            i += 1;
            while i < chars.len() {
                out.push(chars[i]);
                if chars[i] == '"' {
                    if i + 1 < chars.len() && chars[i + 1] == '"' {
                        out.push('"');
                        i += 2;
                        continue;
                    }
                    i += 1;
                    break;
                }
                i += 1;
            }
            continue;
        }
        if is_name_start(c) && (i == 0 || !is_name_char(chars[i - 1])) {
            let start = i;
            while i < chars.len() && is_name_char(chars[i]) {
                i += 1;
            }
            let name: String = chars[start..i].iter().collect();
            let mut j = i;
            while j < chars.len() && chars[j] == ' ' {
                j += 1;
            }
            let is_call = j < chars.len() && chars[j] == '(';
            let upper = name.to_ascii_uppercase();
            match (is_call, prefix_for(&upper)) {
                (true, Some(prefix)) => {
                    out.push_str(prefix);
                    out.push_str(&upper);
                }
                _ => out.push_str(&name),
            }
            continue;
        }
        out.push(c);
        i += 1;
    }
    out
}

/// A formula's text as a file holds it (without its `=`): each spill read
/// through `#` as the file spells it (KERNEL.55, [`spills_filed`]), its newer
/// functions under their prefixes and the names its `LET`s and `LAMBDA`s bind
/// under `_xlpm.` (KERNEL.29), a cell's, a rule's and a defined name's alike.
pub fn as_filed(formula: &str) -> String {
    prefix_parameters(&prefix_future_functions(&spills_filed(formula)))
}

/// A formula's text with each spill a cell's `#` reads, `A1#` or
/// `Spills!A1#`, written as the file spells it, `ANCHORARRAY(A1)`, for
/// [`prefix_future_functions`] to prefix (KERNEL.55): by the reference
/// scanner's reading, so that a `#` inside a text, a quoted sheet's name or an
/// error's spelling stands as it is, and so does a defined name's spill,
/// `Name#`, which the language does not read.
fn spills_filed(formula: &str) -> std::borrow::Cow<'_, str> {
    let is_spill =
        |r: &refers::FormulaRef| r.kind == refers::Kind::Spill && r.shape == refers::Shape::Cell;
    let scanned = refers::scan(formula);
    if !scanned.iter().any(is_spill) {
        return std::borrow::Cow::Borrowed(formula);
    }
    let chars: Vec<char> = formula.chars().collect();
    let mut out = String::with_capacity(formula.len() + 16);
    let mut at = 0;
    for r in scanned.iter().filter(|r| is_spill(r)) {
        let end = (r.pos + r.span).min(chars.len());
        out.extend(&chars[at..r.pos]);
        let written: String = chars[r.pos..end].iter().collect();
        out.push_str("ANCHORARRAY(");
        out.push_str(written.strip_suffix('#').unwrap_or(&written));
        out.push(')');
        at = end;
    }
    out.extend(&chars[at..]);
    std::borrow::Cow::Owned(out)
}

/// A defined name's text as a file holds it: [`as_filed`], with no `=`,
/// which a machine's `name` row writes before its text and a file never does.
pub fn name_as_filed(text: &str) -> String {
    as_filed(text.strip_prefix('=').unwrap_or(text))
}

/// One piece of a formula's text as [`prefix_parameters`] reads it.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Piece {
    /// A name: letters, digits, `_` and `.` from a letter or `_`.
    Name,
    LParen,
    RParen,
    Comma,
    Bang,
    /// A string literal, a quoted sheet's name, a bracket, a space or any
    /// other character, written as it stands.
    Other,
}

/// A formula's text in pieces, each with where it starts and ends.
fn pieces(chars: &[char]) -> Vec<(Piece, usize, usize)> {
    let mut out = Vec::new();
    let mut i = 0;
    // To the closing `close` of a literal opened at `i`, a doubled one inside.
    let literal = |i: usize, close: char| {
        let mut k = i + 1;
        while k < chars.len() {
            if chars[k] == close {
                if k + 1 < chars.len() && chars[k + 1] == close {
                    k += 2;
                    continue;
                }
                return k + 1;
            }
            k += 1;
        }
        k
    };
    while i < chars.len() {
        let c = chars[i];
        let (piece, end) = match c {
            '"' => (Piece::Other, literal(i, '"')),
            '\'' => (Piece::Other, literal(i, '\'')),
            '[' => {
                let mut depth = 0;
                let mut k = i;
                while k < chars.len() {
                    match chars[k] {
                        '[' => depth += 1,
                        ']' => {
                            depth -= 1;
                            if depth == 0 {
                                k += 1;
                                break;
                            }
                        }
                        _ => {}
                    }
                    k += 1;
                }
                (Piece::Other, k)
            }
            '(' => (Piece::LParen, i + 1),
            ')' => (Piece::RParen, i + 1),
            ',' => (Piece::Comma, i + 1),
            '!' => (Piece::Bang, i + 1),
            c if is_name_start(c) && (i == 0 || !is_name_char(chars[i - 1])) => {
                let mut k = i;
                while k < chars.len() && is_name_char(chars[k]) {
                    k += 1;
                }
                (Piece::Name, k)
            }
            _ => (Piece::Other, i + 1),
        };
        out.push((piece, i, end));
        i = end;
    }
    out
}

/// A name's text as a binding keys it, its `_xlpm.` and its case aside.
fn binding_key(name: &str) -> String {
    let bare = match name.get(..6) {
        Some(p) if p.eq_ignore_ascii_case("_xlpm.") => &name[6..],
        _ => name,
    };
    bare.to_ascii_lowercase()
}

/// A formula's text (without its leading `=`) with every name a `LET` binds
/// or a `LAMBDA` takes written under `_xlpm.`, where it is bound and wherever
/// it is read or called in its scope, as Excel's files hold them (KERNEL.29):
/// `LET(x,1,x+1)` is `LET(_xlpm.x,1,_xlpm.x+1)`. A string literal, a quoted
/// sheet's name and a bracket are left as they stand, and so are a name a
/// sheet is named by or scopes and a name already prefixed. A `LET` or a
/// `LAMBDA` whose binders are not each one name is left to Excel's reading.
pub fn prefix_parameters(formula: &str) -> String {
    let chars: Vec<char> = formula.chars().collect();
    let ps = pieces(&chars);
    // Each `(`'s `)`.
    let mut close = vec![usize::MAX; ps.len()];
    let mut open: Vec<usize> = Vec::new();
    for (k, p) in ps.iter().enumerate() {
        match p.0 {
            Piece::LParen => open.push(k),
            Piece::RParen => {
                if let Some(o) = open.pop() {
                    close[o] = k;
                }
            }
            _ => {}
        }
    }
    let mut out = String::with_capacity(formula.len() + 16);
    let mut scope: Vec<String> = Vec::new();
    Prefixer {
        chars: &chars,
        ps: &ps,
        close: &close,
    }
    .write(0, ps.len(), &mut scope, &mut out);
    out
}

struct Prefixer<'a> {
    chars: &'a [char],
    ps: &'a [(Piece, usize, usize)],
    close: &'a [usize],
}

impl Prefixer<'_> {
    fn text(&self, k: usize) -> String {
        let (_, a, b) = self.ps[k];
        self.chars[a..b].iter().collect()
    }

    /// Whether piece `k` is a space, which a binder's argument may hold
    /// around its name.
    fn blank(&self, k: usize) -> bool {
        let (p, a, b) = self.ps[k];
        p == Piece::Other && self.chars[a..b].iter().all(|c| c.is_whitespace())
    }

    /// The arguments of the call whose `(` is piece `o`: each the pieces
    /// between its commas at the call's own depth.
    fn arguments(&self, o: usize) -> Vec<(usize, usize)> {
        let end = self.close[o];
        let mut args = Vec::new();
        let mut from = o + 1;
        let mut k = o + 1;
        while k < end {
            match self.ps[k].0 {
                Piece::LParen if self.close[k] != usize::MAX => {
                    k = self.close[k] + 1;
                    continue;
                }
                Piece::Comma => {
                    args.push((from, k));
                    from = k + 1;
                }
                _ => {}
            }
            k += 1;
        }
        args.push((from, end));
        args
    }

    /// The one name an argument holds, spaces aside, by its piece.
    fn binder(&self, (a, b): (usize, usize)) -> Option<usize> {
        let mut found = None;
        for k in a..b {
            if self.blank(k) {
                continue;
            }
            if self.ps[k].0 != Piece::Name || found.is_some() {
                return None;
            }
            found = Some(k);
        }
        found
    }

    fn binding_word(&self, k: usize) -> Option<&'static str> {
        if self.ps[k].0 != Piece::Name {
            return None;
        }
        let name = self.text(k).to_ascii_uppercase();
        match name.strip_prefix("_XLFN.").unwrap_or(&name) {
            "LET" => Some("LET"),
            "LAMBDA" => Some("LAMBDA"),
            _ => None,
        }
    }

    fn name(&self, k: usize, out: &mut String) {
        let t = self.text(k);
        if t.len() >= 6 && t[..6].eq_ignore_ascii_case("_xlpm.") {
            out.push_str(&t);
        } else {
            out.push_str("_xlpm.");
            out.push_str(&t);
        }
    }

    fn write(&self, from: usize, to: usize, scope: &mut Vec<String>, out: &mut String) {
        let mut k = from;
        while k < to {
            let (piece, _, _) = self.ps[k];
            if let Some(word) = self.binding_word(k) {
                if k + 1 < to
                    && self.ps[k + 1].0 == Piece::LParen
                    && self.close[k + 1] != usize::MAX
                {
                    let o = k + 1;
                    let end = self.close[o];
                    out.push_str(&self.text(k));
                    out.push('(');
                    self.binding(word, o, scope, out);
                    out.push(')');
                    k = end + 1;
                    continue;
                }
            }
            let bound = piece == Piece::Name
                && !(k + 1 < self.ps.len() && self.ps[k + 1].0 == Piece::Bang)
                && !(k > 0 && self.ps[k - 1].0 == Piece::Bang)
                && scope.contains(&binding_key(&self.text(k)));
            if bound {
                self.name(k, out);
            } else {
                out.push_str(&self.text(k));
            }
            k += 1;
        }
    }

    /// A `LET`'s or a `LAMBDA`'s arguments, its binders under `_xlpm.` and
    /// each in scope where it binds.
    fn binding(&self, word: &str, o: usize, scope: &mut Vec<String>, out: &mut String) {
        let args = self.arguments(o);
        let n = args.len();
        let binders: Vec<usize> = match word {
            "LAMBDA" => (0..n.saturating_sub(1)).collect(),
            _ if n >= 3 && n % 2 == 1 => (0..n - 1).step_by(2).collect(),
            _ => Vec::new(),
        };
        let found: Option<Vec<usize>> = binders.iter().map(|&j| self.binder(args[j])).collect();
        let Some(found) = found else {
            // Not each binder one name: the arguments as they stand.
            let end = self.close[o];
            self.write(o + 1, end, scope, out);
            return;
        };
        let mark = scope.len();
        // A `LET`'s name waits for its value, in scope for every value after
        // it and its body; a `LAMBDA`'s parameters are in scope for its body.
        let mut waiting: Option<String> = None;
        for (j, &(a, b)) in args.iter().enumerate() {
            if j > 0 {
                out.push(',');
            }
            match binders.iter().position(|&x| x == j) {
                Some(at) => {
                    for k in a..b {
                        if k == found[at] {
                            self.name(k, out);
                        } else {
                            out.push_str(&self.text(k));
                        }
                    }
                    let key = binding_key(&self.text(found[at]));
                    if word == "LET" {
                        waiting = Some(key);
                    } else {
                        scope.push(key);
                    }
                }
                None => {
                    self.write(a, b, scope, out);
                    if let Some(key) = waiting.take() {
                        scope.push(key);
                    }
                }
            }
        }
        scope.truncate(mark);
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn let_and_lambda_s_names_are_written_under_their_prefix() {
        assert_eq!(
            prefix_parameters("_xlfn.LET(x,1,x*2)"),
            "_xlfn.LET(_xlpm.x,1,_xlpm.x*2)"
        );
        assert_eq!(
            prefix_parameters("LAMBDA(x,y,x+y)(1,2)"),
            "LAMBDA(_xlpm.x,_xlpm.y,_xlpm.x+_xlpm.y)(1,2)"
        );
        // A name in scope for the values after it and the body, a call of
        // one prefixed, a shadow inside its own scope.
        assert_eq!(
            prefix_parameters("LET(f,LAMBDA(n,n+1),g,f(2),LET(f,3,f+g))"),
            "LET(_xlpm.f,LAMBDA(_xlpm.n,_xlpm.n+1),_xlpm.g,_xlpm.f(2),LET(_xlpm.f,3,_xlpm.f+_xlpm.g))"
        );
        // Outside its scope a name is a defined name; a name's own value
        // does not see it; text, quoted sheets and a sheet's name untouched.
        assert_eq!(
            prefix_parameters("LET(x,x,x)+x"),
            "LET(_xlpm.x,x,_xlpm.x)+x"
        );
        assert_eq!(
            prefix_parameters("LET(x,\"x\",'x'!A1&x&x!B2)"),
            "LET(_xlpm.x,\"x\",'x'!A1&_xlpm.x&x!B2)"
        );
        // Already prefixed, and spaces around a binder.
        assert_eq!(
            prefix_parameters("LET(_xlpm.x,1,_xlpm.x)"),
            "LET(_xlpm.x,1,_xlpm.x)"
        );
        assert_eq!(
            prefix_parameters("LET( x ,1, x )"),
            "LET( _xlpm.x ,1, _xlpm.x )"
        );
        // A binding whose binders are not names is left as it stands.
        assert_eq!(prefix_parameters("LET(A1+1,2,3)"), "LET(A1+1,2,3)");
        assert_eq!(prefix_parameters("SUM(A1:A3)"), "SUM(A1:A3)");
    }

    #[test]
    fn newer_functions_take_their_prefix_and_older_ones_do_not() {
        assert_eq!(
            prefix_future_functions("IFS(B2>3,\"big\",TRUE,\"small\")"),
            "_xlfn.IFS(B2>3,\"big\",TRUE,\"small\")"
        );
        assert_eq!(prefix_future_functions("SUM(B2:B9)"), "SUM(B2:B9)");
        assert_eq!(prefix_future_functions("ifs(1,2)"), "_xlfn.IFS(1,2)");
        assert_eq!(
            prefix_future_functions("FILTER(A:A,B:B>1)"),
            "_xlfn._xlws.FILTER(A:A,B:B>1)"
        );
        assert_eq!(
            prefix_future_functions("STDEV.S(B2:B50)+STDEV(B2:B50)"),
            "_xlfn.STDEV.S(B2:B50)+STDEV(B2:B50)"
        );
        assert_eq!(
            prefix_future_functions("UNICHAR(160)&CHAR(120)"),
            "_xlfn.UNICHAR(160)&CHAR(120)"
        );
        // Excel's `@`: without its prefix Excel reads `SINGLE` as a function
        // of no add-in and answers `#NAME?` (KERNEL.8, slice b, Excel's save
        // of the intersection's fixture).
        assert_eq!(
            prefix_future_functions("ROW(SINGLE(A1:A8))"),
            "ROW(_xlfn.SINGLE(A1:A8))"
        );
    }

    #[test]
    fn what_is_not_a_call_is_left_alone() {
        // A string literal, with a doubled quote inside it.
        assert_eq!(
            prefix_future_functions("\"IFS(\"\"x\"\")\"&SWITCH(1,1,\"a\")"),
            "\"IFS(\"\"x\"\")\"&_xlfn.SWITCH(1,1,\"a\")"
        );
        // Already prefixed: the name read is `_xlfn.IFS`, not in the table.
        assert_eq!(prefix_future_functions("_xlfn.IFS(1,2)"), "_xlfn.IFS(1,2)");
        // A name or a reference spelled like a function, with no `(` after it.
        assert_eq!(prefix_future_functions("LET+1"), "LET+1");
        assert_eq!(prefix_future_functions("Data!IFS1"), "Data!IFS1");
        // A space before the parenthesis still makes a call.
        assert_eq!(prefix_future_functions("IFS (1,2)"), "_xlfn.IFS (1,2)");
    }

    #[test]
    fn a_function_that_can_return_an_array_is_known() {
        assert!(can_return_array("IFS(B2>3,\"big\",TRUE,\"small\")"));
        assert!(can_return_array("_xlfn.IFS(1,2)"));
        assert!(can_return_array("_xlfn._xlws.FILTER(A:A,B:B>1)"));
        assert!(can_return_array("1+INDEX(A1:A3,2)"));
        assert!(can_return_array("sequence(3)"));
        assert!(!can_return_array("B2*2"));
        assert!(!can_return_array("SUM(B2:B9)"));
        assert!(!can_return_array("IF(B2>3,\"big\",\"small\")"));
        assert!(!can_return_array("VLOOKUP(A1,B:C,2,FALSE)"));
        assert!(!can_return_array("\"INDEX(\"&A1"));
        assert!(!can_return_array("INDEX+1"));
        // KERNEL.29: a call in place, which a legacy formula cannot hold.
        assert!(can_return_array("IF(FALSE,SQ,PLUS1)(5)"));
        assert!(can_return_array("IFERROR(SQ,PLUS1)(3)"));
        assert!(can_return_array("_xlfn.LAMBDA(_xlpm.x,_xlpm.x+1)(5)"));
        assert!(!can_return_array("\")(\"&A1"));
        assert!(!can_return_array("(A1)+(B1)"));
        assert!(!can_return_array("SQ(3)"));
    }

    #[test]
    fn a_cell_s_spill_is_filed_as_excel_files_it() {
        // KERNEL.55's second slice: `A1#` under `_xlfn.ANCHORARRAY`, as the
        // save of the item's probe spells it, a sheet's cell's too, the
        // function's own spelling prefixed once; a `#` inside a text, a
        // defined name's spill and an error's spelling left as they stand.
        assert_eq!(as_filed("SUM(A1#)"), "SUM(_xlfn.ANCHORARRAY(A1))");
        assert_eq!(
            as_filed("Spills!$B$2#*10+A1#"),
            "_xlfn.ANCHORARRAY(Spills!$B$2)*10+_xlfn.ANCHORARRAY(A1)"
        );
        assert_eq!(as_filed("ANCHORARRAY(D1)"), "_xlfn.ANCHORARRAY(D1)");
        assert_eq!(as_filed("_xlfn.ANCHORARRAY(D1)"), "_xlfn.ANCHORARRAY(D1)");
        assert_eq!(as_filed("\"A1#\"&A1"), "\"A1#\"&A1");
        assert_eq!(as_filed("Rate#"), "Rate#");
        assert_eq!(as_filed("IFERROR(A1,#N/A)"), "IFERROR(A1,#N/A)");
    }
}
