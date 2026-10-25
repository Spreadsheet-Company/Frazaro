//! The formula syntax (VLA.4, 2026-10-17): a VLA expression written as a
//! formula's text, and a formula's text read back as a VLA expression.
//!
//! [`write`] takes an expression the expander has finished with, `(if (> A1
//! 3) "big" "small")`, and answers the formula's text in Excel's invariant
//! spelling, `=IF(A1>3,"big","small")`: the spelling a file and the view
//! record hold (ISO/IEC 29500-1, 18.17), English function names, `,`
//! between arguments, `.` as the decimal mark, A1 references, and none of
//! the `_xlfn.` prefixes a file writer adds (`sheet::strip_future_prefixes`
//! is the reading side of that rule). [`read`] is its inverse, for tools and
//! tests: formula text, with or without its `=`, to the canonical
//! expression. The expander's `formula-expand` is the writer at expand time
//! (`expand.rs`), so that a cartridge's source composes a formula from
//! expressions and its macros, where a template macro could only fuse a
//! flat one from texts (Alonzo's SPEC 7.8; KIT.1's "a macro's call is no
//! argument of `&expand`").
//!
//! **Canonical, both ways.** The text is written with no space outside a
//! text, the fewest parentheses Excel's precedence allows, every function
//! name, truth value, error and column letter upper-case, and a sheet quoted
//! only when `refers::quote_sheet` says it must be; a number and a text are
//! written as they stand. The expression is read with lower-case heads, a
//! left chain of one operator as one form of many operands, a sign on a
//! number as the signed number, and a reference as the bare symbol it can
//! be (`Input!B5`) or, where a symbol cannot hold it, `(ref "Q1 Data"
//! "A1")`. So `read(write(v))` is `v` in that canonical form, and
//! `write(read(f))` is `f` in the canonical text, byte for byte: the laws
//! the tests hold, over every formula of the corpus as well as the cases.
//! What is canonicalized away is what Excel computes the same: spaces,
//! redundant parentheses, `sum(` for `SUM(`, `'Input'!B5` for `Input!B5`.
//! What equivalence means here is the language's own reading: the
//! evaluator's parser (`calc::formula::parse`) reads the text written and
//! the text read as one tree. Nothing is claimed of Excel's values, of a
//! function's existence or of its arity, which are evaluation's.
//!
//! **Precedence**, Excel's as `calc::formula` reads it, lowest first: the
//! comparisons; `&`; `+` and `-`; `*` and `/`; `^`; a sign; `%`; and `:`
//! between two expressions that answer references (KERNEL.8, slice f),
//! `(: A1 (index A:A 5))` for `A1:INDEX(A:A,5)`, above them all; every
//! binary operator left-associative, so `2^3^2` is 64. A left operand is
//! parenthesized when it binds looser than its operator, a right one when
//! it binds looser or equally; a sign's operand when it is any binary
//! operation, since a sign binds tighter than `^` (`(- (^ x 2))` is
//! `-(x^2)`, and `(^ (- x) 2)` is `-x^2`, Excel's own trap); `%`'s when it
//! is a sign or a binary operation. A form of many operands is the left
//! fold VLA's own operators are (`EmitChain`, `EvalOpChain`): `(- a b c)`
//! is `a-b-c`, and `(< a b c)` is `a<b<c`, which compares `a<b` with `c`.
//!
//! **The atoms.** A number is the formula's numeral, `5`, `1.50`, `.5`,
//! `1e3`, as written, a `-` before it the canonical sign; VBA's own
//! numerals (`&HFF`, `1d5`, `+5`) are refused, since a formula does not
//! read them. A text is a text literal, `""` for a quote mark inside it,
//! at most 255 characters, Excel's limit. `true` and `false` are `TRUE` and
//! `FALSE`; an error is `(error "#N/A")`, as the rows spell a value, the
//! seven a formula holds. A symbol the reference scanner reads whole as one
//! reference or one name is that reference or that name, its column letters
//! upper-cased and its sheet quoted by the one quoting rule; any other
//! symbol is refused, never renamed (`tax-rate` would read as `tax` minus
//! `rate`). Inside the syntax a string is always a text: a formula's text
//! kept raw is the string a row holds outside it.
//!
//! **The heads.** The twelve operators and `%`; `error`, `ref`, `last` and
//! `lambda`, the syntax's own forms, `(last X)` the twin of the one
//! reference on a sheet X is, `Input.last!B13` for `Input!B13` (VLA.6); a
//! head the head table declares no
//! formula's (`set!`, `begin`, `.`, `new`; the table's formula column,
//! F.7's declaration) refused; any other head a function call, its name
//! letters, digits, `.` and `_` from a letter on. `if`, `not` and `mod`,
//! VLA's own heads, have their arity checked; an Excel function's arity is
//! Excel's. `quote` is an array constant (KERNEL.8, slice f): `(quote (1 2
//! 3))` one row, `{1,2,3}`, and `(quote ((1 2) (3 4)))` rows, `{1,2;3,4}`,
//! each value a numeral, a text, a truth value or an error, every row as
//! long as the first.
//!
//! **What [`read`] refuses** is told apart: a text that is no formula at
//! all (`vla-formula-read-malformed`), and a construct Excel reads that
//! this version does not (`vla-formula-read-unsupported`): `@`, an
//! intersection, a union, a spill, a structured reference, a link into
//! another workbook, an omitted argument, a `_xlfn.` spelling, an error
//! beyond the seven, a LAMBDA called in place, and a function named like a word the
//! syntax keeps for its own forms. The reader has no VBA twin, as the
//! evaluator has none: the add-in's formulas are read by Excel, and its two
//! ids are raised by no VBA site. The writer has one (`VLA.bas`'s
//! `FormulaSyntaxWrite`), held to this one by the expand golden.

use crate::calc::formula::{LENGTH_LIMIT, NESTING_LIMIT};
use crate::calc::value::ErrorKind;
use crate::form::{Form, List};
use crate::headtable::{self, FormulaSupport};
use crate::intrinsics::{fold, is_numeric};
use crate::messages::{raise, Refusal};
use crate::printer::write_datum;
use crate::reader::FORM_DEPTH_LIMIT;
use crate::refers::{self, is_ident_char, quote_sheet, FormulaRef, Kind, Shape};
use crate::sheet::column_letters;

/// The longest text a formula's literal holds, Excel's own limit, in UTF-16
/// code units as Excel and VBA's `Len` count them.
pub const TEXT_LIMIT: usize = 255;

// The precedence levels, lowest first.
const CMP: u8 = 1;
const CAT: u8 = 2;
const ADD: u8 = 3;
const MUL: u8 = 4;
const POW: u8 = 5;
const SIGN: u8 = 6;
const PCT: u8 = 7;
const RANGE: u8 = 8;
const ATOM: u8 = 9;

/// The binary operators, as the head spells each and the text writes it.
const BINARY: &[&str] = &[
    "=", "<>", "<", ">", "<=", ">=", "&", "+", "-", "*", "/", "^",
];

fn binary_level(op: &str) -> u8 {
    match op {
        "=" | "<>" | "<" | ">" | "<=" | ">=" => CMP,
        "&" => CAT,
        "+" | "-" => ADD,
        "*" | "/" => MUL,
        _ => POW,
    }
}

// ---------------------------------------------------------------------
//  The writer
// ---------------------------------------------------------------------

/// An expression the expander has finished with, as a formula's canonical
/// text with its leading `=`; refused by name when it is not one.
pub fn write(expr: &Form) -> Result<String, Refusal> {
    Ok(format!("={}", text(expr, 1)?))
}

/// The expression's level of precedence, as the writer will spell it.
fn level(f: &Form) -> u8 {
    match f {
        Form::Sym(s) if is_negative_numeral(s) => SIGN,
        Form::List(l) => {
            let Some(Form::Sym(head)) = l.items.first() else {
                return ATOM;
            };
            let h = fold(head);
            let n = l.items.len() - 1;
            match h.as_str() {
                "+" | "-" if n == 1 => SIGN,
                "%" if n == 1 => PCT,
                ":" if n >= 2 => RANGE,
                op if n >= 2 && BINARY.contains(&op) => binary_level(op),
                _ => ATOM,
            }
        }
        _ => ATOM,
    }
}

/// An operand, parenthesized when the level it sits at needs it: `right`
/// for every operand of a binary operator after its first.
fn operand(f: &Form, at: u8, right: bool, depth: u32) -> Result<String, Refusal> {
    let t = text(f, depth + 1)?;
    let l = level(f);
    Ok(if l < at || (right && l == at) {
        format!("({t})")
    } else {
        t
    })
}

fn text(f: &Form, depth: u32) -> Result<String, Refusal> {
    if depth > FORM_DEPTH_LIMIT {
        return Err(raise(
            "vla-form-too-deep",
            &[("limit", &FORM_DEPTH_LIMIT.to_string()), ("loc", "")],
        ));
    }
    match f {
        Form::Str(s) => text_literal(s),
        Form::Sym(s) => atom(s),
        Form::List(l) => list(l, depth),
    }
}

/// A text literal: its quote marks doubled, at most [`TEXT_LIMIT`] long.
fn text_literal(s: &str) -> Result<String, Refusal> {
    let n = s.encode_utf16().count();
    if n > TEXT_LIMIT {
        return Err(raise("vla-formula-text-too-long", &[("n", &n.to_string())]));
    }
    Ok(format!("\"{}\"", s.replace('"', "\"\"")))
}

/// A symbol: a numeral, a truth value, a reference or a name, each in its
/// canonical spelling; anything else refused.
fn atom(s: &str) -> Result<String, Refusal> {
    if is_numeral(s) || is_negative_numeral(s) {
        let unsigned = s.strip_prefix('-').unwrap_or(s);
        if !unsigned.parse::<f64>().is_ok_and(f64::is_finite) {
            return Err(raise("vla-formula-numeral-unspelled", &[("tok", s)]));
        }
        return Ok(s.to_string());
    }
    match fold(s).as_str() {
        "true" => return Ok("TRUE".to_string()),
        "false" => return Ok("FALSE".to_string()),
        _ => {}
    }
    match whole_reference(s) {
        Some(r) => match r.kind {
            Kind::Cell | Kind::Range | Kind::Column | Kind::Row | Kind::ThreeD | Kind::Name => {
                Ok(reference_text(&r))
            }
            _ => Err(raise(
                "vla-formula-unsupported",
                &[("construct", &construct_of(&r))],
            )),
        },
        None if is_numeric(s) => Err(raise("vla-formula-numeral-unspelled", &[("tok", s)])),
        None => Err(raise("vla-formula-atom-unspelled", &[("tok", s)])),
    }
}

/// The one record the scanner reads over the whole of `s`, or none.
pub(crate) fn whole_reference(s: &str) -> Option<FormulaRef> {
    if s.is_empty() || s.starts_with('=') {
        return None;
    }
    let mut found = refers::scan(s);
    if found.len() != 1 {
        return None;
    }
    let r = found.pop()?;
    (r.pos == 0 && r.span == s.chars().count()).then_some(r)
}

/// What a construct outside this version is called in a refusal.
fn construct_of(r: &FormulaRef) -> String {
    let what = match r.kind {
        Kind::Spill => "the spill reference",
        Kind::Structured => "the structured reference",
        Kind::External => "the reference into another workbook",
        Kind::Broken => "the broken reference",
        _ => "the reference",
    };
    format!("{what} {}", r.written)
}

/// A reference's canonical spelling: its sheet or span quoted only where it
/// must be, then its part ([`part_text`]).
fn reference_text(r: &FormulaRef) -> String {
    let part = part_text(r);
    if r.sheet_name.is_empty() {
        part
    } else {
        format!("{}!{part}", sheet_text(&r.sheet_name))
    }
}

/// A reference's part, after its sheet, in its canonical spelling: its
/// column letters upper-case, its `$` marks as written; a name as written.
/// VLA.6: apart from [`reference_text`], for the expander's `range-expand`.
pub(crate) fn part_text(r: &FormulaRef) -> String {
    match (r.kind, r.shape) {
        (Kind::Name, _) | (_, Shape::None) => r.part.clone(),
        (_, Shape::Cell) => cell_text(r.col1, r.col_abs1, r.row1, r.row_abs1),
        (_, Shape::Range) => format!(
            "{}:{}",
            cell_text(r.col1, r.col_abs1, r.row1, r.row_abs1),
            cell_text(r.col2, r.col_abs2, r.row2, r.row_abs2)
        ),
        (_, Shape::Column) => format!(
            "{}{}:{}{}",
            mark(r.col_abs1),
            column_letters(r.col1),
            mark(r.col_abs2),
            column_letters(r.col2)
        ),
        (_, Shape::Row) => format!(
            "{}{}:{}{}",
            mark(r.row_abs1),
            r.row1,
            mark(r.row_abs2),
            r.row2
        ),
    }
}

fn cell_text(col: u32, col_abs: bool, row: u32, row_abs: bool) -> String {
    format!(
        "{}{}{}{}",
        mark(col_abs),
        column_letters(col),
        mark(row_abs),
        row
    )
}

fn mark(abs: bool) -> &'static str {
    if abs {
        "$"
    } else {
        ""
    }
}

/// A sheet, or a span of sheets, as a qualifier spells it: a span is quoted
/// whole when either end needs quotes, as Excel writes `'Jan 1:Dec'!A1`.
fn sheet_text(sheet: &str) -> String {
    match sheet.split_once(':') {
        Some((first, last)) => {
            if quote_sheet(first) == first && quote_sheet(last) == last {
                format!("{first}:{last}")
            } else {
                format!(
                    "'{}:{}'",
                    first.replace('\'', "''"),
                    last.replace('\'', "''")
                )
            }
        }
        None => quote_sheet(sheet),
    }
}

/// A list: an operator, one of the syntax's own forms, or a call.
fn list(l: &List, depth: u32) -> Result<String, Refusal> {
    let Some(head_form) = l.items.first() else {
        return Err(raise("vla-formula-head-not-a-function", &[("head", "()")]));
    };
    let Form::Sym(head) = head_form else {
        return Err(raise(
            "vla-formula-head-not-a-function",
            &[("head", &write_datum(head_form))],
        ));
    };
    let h = headtable::resolve_head_alias(&fold(head));
    let args = &l.items[1..];
    let n = args.len();
    match h.as_str() {
        "+" | "-" if n == 1 => Ok(format!("{h}{}", operand(&args[0], SIGN, false, depth)?)),
        "%" if n == 1 => Ok(format!("{}%", operand(&args[0], PCT, false, depth)?)),
        "%" => Err(arity(head, "one operand", n)),
        op if BINARY.contains(&op) => {
            if n < 2 {
                return Err(arity(
                    head,
                    if op == "+" || op == "-" {
                        "one operand or more"
                    } else {
                        "two operands or more"
                    },
                    n,
                ));
            }
            let at = binary_level(op);
            let mut out = operand(&args[0], at, false, depth)?;
            for a in &args[1..] {
                // `#N/A`, the one error that ends in neither `!` nor `?`,
                // is read with a `/` after it as one token by the
                // language's parser (`calc::formula`), so before a `/` it
                // is parenthesized: `x*(#N/A)/y`.
                if op == "/" && out.ends_with("#N/A") {
                    out.truncate(out.len() - 4);
                    out.push_str("(#N/A)");
                }
                out.push_str(op);
                out.push_str(&operand(a, at, true, depth)?);
            }
            Ok(out)
        }
        // `:` between two expressions that answer references, above every
        // other operator (KERNEL.8, slice f).
        ":" => {
            if n < 2 {
                return Err(arity(head, "two operands or more", n));
            }
            let mut out = operand(&args[0], RANGE, false, depth)?;
            for a in &args[1..] {
                out.push(':');
                out.push_str(&operand(a, RANGE, true, depth)?);
            }
            Ok(out)
        }
        "error" => error_literal(head, args),
        "ref" => ref_form(head, args),
        "last" => last_form(head, args),
        "lambda" => lambda(head, args, depth),
        "quote" => quote_form(head, args),
        _ => {
            if let Some(row) = headtable::row(&h) {
                if row.formula != FormulaSupport::Yes {
                    return Err(raise("vla-formula-head-refused", &[("head", head)]));
                }
            }
            let want = match h.as_str() {
                "if" => (!(2..=3).contains(&n)).then_some("two or three arguments"),
                "not" => (n != 1).then_some("one argument"),
                "mod" => (n != 2).then_some("two arguments"),
                _ => None,
            };
            if let Some(want) = want {
                return Err(arity(head, want, n));
            }
            let name = if h == fold(head) {
                head.as_str()
            } else {
                h.as_str()
            };
            if !is_function_name(name) {
                return Err(raise("vla-formula-head-not-a-function", &[("head", head)]));
            }
            let mut parts = Vec::with_capacity(n);
            for a in args {
                parts.push(text(a, depth + 1)?);
            }
            Ok(format!(
                "{}({})",
                name.to_ascii_uppercase(),
                parts.join(",")
            ))
        }
    }
}

fn arity(head: &str, want: &str, n: usize) -> Refusal {
    raise(
        "vla-formula-arity",
        &[("head", head), ("want", want), ("n", &n.to_string())],
    )
}

/// `(error "#N/A")`: one of the seven errors a formula holds.
fn error_literal(head: &str, args: &[Form]) -> Result<String, Refusal> {
    let [Form::Str(t)] = args else {
        return Err(arity(head, "one text, such as \"#N/A\"", args.len()));
    };
    match ErrorKind::parse(t) {
        Some(k) if k.is_literal() => Ok(k.text().to_string()),
        _ => Err(raise(
            "vla-formula-unsupported",
            &[("construct", &format!("the error {t}"))],
        )),
    }
}

/// Whether a form is an error literal, `(error "#N/A")`, a value an array
/// constant holds.
fn is_error_form(f: &Form) -> bool {
    match f {
        Form::List(l) => matches!(l.items.first(), Some(Form::Sym(h)) if fold(h) == "error"),
        _ => false,
    }
}

/// `(quote (1 2 3))`, one row, or `(quote ((1 2) (3 4)))`, rows: an array
/// constant, `{1,2,3}` and `{1,2;3,4}` (KERNEL.8, slice f), each value a
/// numeral, a text, a truth value or `(error "#N/A")`, every row as long as
/// the first; anything else refused by name.
fn quote_form(head: &str, args: &[Form]) -> Result<String, Refusal> {
    let [Form::List(l)] = args else {
        return Err(arity(head, "one list of values or of rows", args.len()));
    };
    let unsupported =
        |construct: &str| raise("vla-formula-unsupported", &[("construct", construct)]);
    if l.items.is_empty() {
        return Err(unsupported("an empty array constant, (quote ())"));
    }
    let is_row = |f: &Form| matches!(f, Form::List(_)) && !is_error_form(f);
    let rows: Vec<&[Form]> = if l.items.iter().all(is_row) {
        l.items
            .iter()
            .map(|f| match f {
                Form::List(r) => r.items.as_slice(),
                _ => &[],
            })
            .collect()
    } else if l.items.iter().any(is_row) {
        return Err(unsupported("an array constant mixing values and rows"));
    } else {
        vec![l.items.as_slice()]
    };
    let width = rows[0].len();
    if width == 0 || rows.iter().any(|r| r.len() != width) {
        return Err(unsupported("an array constant whose rows differ in length"));
    }
    let mut out = String::from("{");
    for (i, row) in rows.iter().enumerate() {
        if i > 0 {
            out.push(';');
        }
        for (j, v) in row.iter().enumerate() {
            if j > 0 {
                out.push(',');
            }
            let value = match v {
                Form::Str(s) => text_literal(s)?,
                Form::Sym(s)
                    if is_numeral(s)
                        || is_negative_numeral(s)
                        || matches!(fold(s).as_str(), "true" | "false") =>
                {
                    atom(s)?
                }
                Form::List(e) if is_error_form(v) => error_literal("error", &e.items[1..])?,
                other => {
                    return Err(unsupported(&format!(
                        "{} in an array constant",
                        write_datum(other)
                    )))
                }
            };
            out.push_str(&value);
        }
    }
    out.push('}');
    Ok(out)
}

/// `(ref "Sheet" "A1")` or `(ref "A1")`: one reference or one name, built
/// from a sheet and an address, the sheet quoted as it must be.
fn ref_form(head: &str, args: &[Form]) -> Result<String, Refusal> {
    let built = match args {
        [Form::Str(part)] => part.clone(),
        [Form::Str(sheet), Form::Str(part)] => format!("{}!{part}", sheet_text(sheet)),
        _ => {
            return Err(arity(
                head,
                "a sheet and an address as texts, or an address alone",
                args.len(),
            ))
        }
    };
    match whole_reference(&built) {
        Some(r)
            if matches!(
                r.kind,
                Kind::Cell | Kind::Range | Kind::Column | Kind::Row | Kind::ThreeD | Kind::Name
            ) =>
        {
            Ok(reference_text(&r))
        }
        _ => Err(raise("vla-formula-ref-invalid", &[("text", &built)])),
    }
}

/// VLA.6, VBA `FsLastForm`: `(last X)`, X's twin: the one reference on a
/// sheet X is ([`sheet_reference`]), with `.last` after its sheet's name,
/// quoted by the sheet rule, its marks as written: `Input.last!B13` for
/// `Input!B13`. A twin's twin is refused, since no sheet a grid may hold ends
/// in `.last` ([`crate::machine::twin_base`], the one rule).
fn last_form(head: &str, args: &[Form]) -> Result<String, Refusal> {
    let [reference] = args else {
        return Err(arity(head, "one reference on a sheet", args.len()));
    };
    let mut r = sheet_reference(reference, "last")?;
    if crate::machine::twin_base(&r.sheet_name).is_some() {
        return Err(raise(
            "vla-formula-last-twin",
            &[("value", &write_datum(reference))],
        ));
    }
    r.sheet_name.push_str(".last");
    Ok(reference_text(&r))
}

/// VLA.6, VBA `FsSheetReference`: one reference on one sheet, as `last` and
/// the expander's `sheet-expand` and `range-expand` read it: a symbol the
/// scanner reads whole as a cell, a range, a column or a row behind its
/// sheet, or `(ref sheet address)`, whose own refusals stand. Refused by
/// name: a reference with no sheet, since a formula does not know the sheet
/// it will stand on when it is written; and a name, a span of sheets, a text
/// or any other form.
pub(crate) fn sheet_reference(f: &Form, op: &str) -> Result<FormulaRef, Refusal> {
    let expected = |what: &str| {
        raise(
            "vla-reference-expected",
            &[("op", op), ("value", &write_datum(f)), ("what", what)],
        )
    };
    let spelled = match f {
        Form::List(l) if f.head_is("ref") => {
            let Some(Form::Sym(head)) = l.items.first() else {
                unreachable!("a list whose head is ref holds it as a symbol")
            };
            ref_form(head, &l.items[1..])?
        }
        Form::List(_) => return Err(expected("no reference")),
        Form::Str(_) => return Err(expected("a text")),
        Form::Sym(s) => s.clone(),
    };
    let Some(r) = whole_reference(&spelled) else {
        return Err(expected("no reference"));
    };
    match r.kind {
        Kind::Cell | Kind::Range | Kind::Column | Kind::Row if r.sheet_name.is_empty() => {
            Err(raise(
                "vla-reference-no-sheet",
                &[("op", op), ("value", &write_datum(f))],
            ))
        }
        Kind::Cell | Kind::Range | Kind::Column | Kind::Row => Ok(r),
        Kind::ThreeD => Err(expected("a span of sheets")),
        Kind::Name => Err(expected("a name")),
        _ => Err(expected("no reference")),
    }
}

/// `(lambda (x y) body)`: Excel's `LAMBDA(x,y,body)`.
fn lambda(head: &str, args: &[Form], depth: u32) -> Result<String, Refusal> {
    let [params, body] = args else {
        return Err(arity(head, "a list of parameters and one body", args.len()));
    };
    let Form::List(params) = params else {
        return Err(raise(
            "vla-formula-lambda-params",
            &[("tok", &write_datum(params))],
        ));
    };
    let mut out = String::from("LAMBDA(");
    for p in &params.items {
        let name = match p {
            Form::Sym(s) => match whole_reference(s) {
                Some(r) if r.kind == Kind::Name && r.sheet_name.is_empty() => s.clone(),
                _ => return Err(raise("vla-formula-lambda-params", &[("tok", s)])),
            },
            other => {
                return Err(raise(
                    "vla-formula-lambda-params",
                    &[("tok", &write_datum(other))],
                ))
            }
        };
        out.push_str(&name);
        out.push(',');
    }
    out.push_str(&text(body, depth + 1)?);
    out.push(')');
    Ok(out)
}

/// The formula's numeral, unsigned: digits with an optional fraction, or a
/// fraction alone, then an optional exponent with digits.
pub fn is_numeral(s: &str) -> bool {
    let b = s.as_bytes();
    let mut i = 0;
    let mut digits = 0;
    while i < b.len() && b[i].is_ascii_digit() {
        i += 1;
        digits += 1;
    }
    if i < b.len() && b[i] == b'.' {
        i += 1;
        while i < b.len() && b[i].is_ascii_digit() {
            i += 1;
            digits += 1;
        }
    }
    if digits == 0 {
        return false;
    }
    if i < b.len() && (b[i] == b'e' || b[i] == b'E') {
        i += 1;
        if i < b.len() && (b[i] == b'+' || b[i] == b'-') {
            i += 1;
        }
        let first = i;
        while i < b.len() && b[i].is_ascii_digit() {
            i += 1;
        }
        if i == first {
            return false;
        }
    }
    i == b.len()
}

pub(crate) fn is_negative_numeral(s: &str) -> bool {
    s.strip_prefix('-').is_some_and(is_numeral)
}

/// A function's name as a formula spells one: a letter, then letters,
/// digits, `.` and `_`.
fn is_function_name(s: &str) -> bool {
    let mut cs = s.chars();
    cs.next().is_some_and(|c| c.is_ascii_alphabetic())
        && cs.all(|c| c.is_ascii_alphanumeric() || c == '.' || c == '_')
}

// ---------------------------------------------------------------------
//  The reader
// ---------------------------------------------------------------------

/// A formula's text, with or without its `=`, as the canonical expression;
/// refused as malformed or as a construct this version does not read.
pub fn read(formula_text: &str) -> Result<Form, Refusal> {
    let body = formula_text.strip_prefix('=').unwrap_or(formula_text);
    if body.chars().count() > LENGTH_LIMIT {
        return Err(malformed(
            "it is longer than 8,192 characters, Excel's own limit",
        ));
    }
    let toks = lex(formula_text)?;
    if toks.is_empty() {
        return Err(malformed("it is empty"));
    }
    let mut p = Reader {
        toks,
        at: 0,
        depth: 0,
    };
    let e = p.comparison()?;
    match p.peek() {
        None => Ok(e),
        Some(t) => Err(p.stray(t)),
    }
}

fn malformed(what: &str) -> Refusal {
    raise("vla-formula-read-malformed", &[("what", what)])
}

fn unsupported(construct: &str) -> Refusal {
    raise("vla-formula-read-unsupported", &[("construct", construct)])
}

#[derive(Clone, Debug, PartialEq)]
enum Tok {
    /// A numeral, as written.
    Num(String),
    Text(String),
    Bool(bool),
    Err(ErrorKind),
    /// A reference or a name, the scanner's record.
    Ref(Box<FormulaRef>),
    /// A name followed by `(`, as written; the `(` follows as its own token.
    Func(String),
    Op(&'static str),
    Percent,
    Open,
    Close,
    Comma,
    /// An array constant's braces and its row separator, and `:` between
    /// two expressions (KERNEL.8, slice f).
    OpenBrace,
    CloseBrace,
    Semi,
    Colon,
}

/// The tokens of a formula's text, the references and names read by
/// `refers::scan`, the one reading of them.
fn lex(text: &str) -> Result<Vec<Tok>, Refusal> {
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
            let r = refs[next_ref].clone();
            next_ref += 1;
            i += r.span.max(1);
            match r.kind {
                Kind::Cell | Kind::Range | Kind::Column | Kind::Row | Kind::ThreeD | Kind::Name => {
                    toks.push(Tok::Ref(Box::new(r)))
                }
                Kind::Broken if r.written == r.part && fold(&r.part) == "#ref!" => {
                    toks.push(Tok::Err(ErrorKind::Ref))
                }
                _ => return Err(unsupported(&construct_of(&r))),
            }
            continue;
        }
        let c = s[i];
        if c.is_whitespace() {
            i += 1;
        } else if c == '"' {
            let mut out = String::new();
            let mut k = i + 1;
            let mut closed = false;
            while k < s.len() {
                if s[k] == '"' {
                    if s.get(k + 1) == Some(&'"') {
                        out.push('"');
                        k += 2;
                        continue;
                    }
                    closed = true;
                    k += 1;
                    break;
                }
                out.push(s[k]);
                k += 1;
            }
            if !closed {
                return Err(malformed("a text that never closes"));
            }
            if out.encode_utf16().count() > TEXT_LIMIT {
                return Err(malformed(
                    "a text longer than 255 characters, Excel's own limit",
                ));
            }
            toks.push(Tok::Text(out));
            i = k;
        } else if c.is_ascii_digit()
            || (c == '.' && s.get(i + 1).is_some_and(|d| d.is_ascii_digit()))
        {
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
            let numeral: String = s[i..k].iter().collect();
            if !numeral.parse::<f64>().is_ok_and(f64::is_finite) {
                return Err(malformed(&format!(
                    "the number {numeral}, which no double holds"
                )));
            }
            toks.push(Tok::Num(numeral));
            i = k;
        } else if c == '#' {
            let mut k = i + 1;
            while k < s.len() && (s[k].is_ascii_alphanumeric() || s[k] == '/' || s[k] == '_') {
                k += 1;
            }
            if matches!(s.get(k), Some('!') | Some('?')) {
                k += 1;
            }
            let tok: String = s[i..k].iter().collect();
            match ErrorKind::parse(&tok) {
                Some(kind) if kind.is_literal() => toks.push(Tok::Err(kind)),
                None if fold(&tok).starts_with("#n/a/") => return Err(unsupported(
                    "#N/A right before a /, which this version reads only parenthesized, (#N/A)",
                )),
                // An error no formula writes as a literal, `#CALC!` among them.
                _ => return Err(unsupported(&format!("the error {tok}"))),
            }
            i = k;
        } else if c.is_ascii_alphabetic() || c == '_' || (c as u32) > 127 {
            let mut k = i;
            while k < s.len() && is_ident_char(s[k]) {
                k += 1;
            }
            let tok: String = s[i..k].iter().collect();
            if s.get(k) == Some(&'(') {
                toks.push(Tok::Func(tok));
            } else {
                match fold(&tok).as_str() {
                    "true" => toks.push(Tok::Bool(true)),
                    "false" => toks.push(Tok::Bool(false)),
                    _ => return Err(malformed(&format!("'{tok}', which reads as no name"))),
                }
            }
            i = k;
        } else {
            let two: String = s[i..(i + 2).min(s.len())].iter().collect();
            let (tok, len) = match two.as_str() {
                "<>" => (Tok::Op("<>"), 2),
                "<=" => (Tok::Op("<="), 2),
                ">=" => (Tok::Op(">="), 2),
                _ => match c {
                    '<' => (Tok::Op("<"), 1),
                    '>' => (Tok::Op(">"), 1),
                    '=' => (Tok::Op("="), 1),
                    '+' => (Tok::Op("+"), 1),
                    '-' => (Tok::Op("-"), 1),
                    '*' => (Tok::Op("*"), 1),
                    '/' => (Tok::Op("/"), 1),
                    '^' => (Tok::Op("^"), 1),
                    '&' => (Tok::Op("&"), 1),
                    '%' => (Tok::Percent, 1),
                    '(' => (Tok::Open, 1),
                    ')' => (Tok::Close, 1),
                    ',' => (Tok::Comma, 1),
                    '{' => (Tok::OpenBrace, 1),
                    '}' => (Tok::CloseBrace, 1),
                    ';' => (Tok::Semi, 1),
                    ':' => (Tok::Colon, 1),
                    '@' => return Err(unsupported("the implicit intersection @")),
                    '[' => return Err(unsupported("a structured reference")),
                    other => {
                        return Err(malformed(&format!(
                            "'{other}', which no formula holds there"
                        )))
                    }
                },
            };
            toks.push(tok);
            i += len;
        }
    }
    Ok(toks)
}

struct Reader {
    toks: Vec<Tok>,
    at: usize,
    depth: u32,
}

impl Reader {
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

    fn peek_op(&self) -> Option<&'static str> {
        match self.peek() {
            Some(Tok::Op(op)) => Some(op),
            _ => None,
        }
    }

    /// `f` one level of nesting deeper, Excel's limit of 64 held.
    fn nested<T>(&mut self, f: impl FnOnce(&mut Self) -> Result<T, Refusal>) -> Result<T, Refusal> {
        if self.depth >= NESTING_LIMIT {
            return Err(malformed("it nests deeper than 64, Excel's own limit"));
        }
        self.depth += 1;
        let r = f(self);
        self.depth -= 1;
        r
    }

    /// The refusal of a token found where nothing more was wanted.
    fn stray(&self, t: &Tok) -> Refusal {
        match t {
            Tok::Close => malformed("a ')' that opens nothing"),
            Tok::Comma => malformed("a ',' outside a call"),
            Tok::CloseBrace => malformed("a '}' that opens nothing"),
            Tok::Semi => malformed("a ';' outside an array constant"),
            _ => malformed("two operands with no operator between them"),
        }
    }

    /// A left-associative run of the operators `ops` over `next`, a run of
    /// one operator read as one form of many operands.
    fn binary(
        &mut self,
        ops: &[&'static str],
        next: fn(&mut Self) -> Result<Form, Refusal>,
    ) -> Result<Form, Refusal> {
        let mut left = next(self)?;
        while let Some(op) = self.peek_op() {
            if !ops.contains(&op) {
                break;
            }
            self.at += 1;
            let right = next(self)?;
            left = match left {
                Form::List(mut l)
                    if l.items.len() >= 3
                        && matches!(l.items.first(), Some(Form::Sym(h)) if h == op) =>
                {
                    l.items.push(right);
                    Form::List(l)
                }
                other => Form::list(vec![Form::sym(op), other, right]),
            };
        }
        Ok(left)
    }

    fn comparison(&mut self) -> Result<Form, Refusal> {
        self.binary(&["=", "<>", "<", ">", "<=", ">="], Self::concat)
    }

    fn concat(&mut self) -> Result<Form, Refusal> {
        self.binary(&["&"], Self::additive)
    }

    fn additive(&mut self) -> Result<Form, Refusal> {
        self.binary(&["+", "-"], Self::term)
    }

    fn term(&mut self) -> Result<Form, Refusal> {
        self.binary(&["*", "/"], Self::power)
    }

    fn power(&mut self) -> Result<Form, Refusal> {
        self.binary(&["^"], Self::unary)
    }

    fn unary(&mut self) -> Result<Form, Refusal> {
        match self.peek_op() {
            Some(op @ ("-" | "+")) => {
                self.at += 1;
                let x = self.nested(Self::unary)?;
                Ok(match (op, x) {
                    ("-", Form::Sym(n)) if is_numeral(&n) => Form::Sym(format!("-{n}")),
                    (_, x) => Form::list(vec![Form::sym(op), x]),
                })
            }
            _ => self.postfix(),
        }
    }

    fn postfix(&mut self) -> Result<Form, Refusal> {
        let mut e = self.ranged()?;
        let mut levels = 0;
        while self.peek() == Some(&Tok::Percent) {
            if self.depth + levels >= NESTING_LIMIT {
                return Err(malformed("it nests deeper than 64, Excel's own limit"));
            }
            levels += 1;
            self.at += 1;
            e = Form::list(vec![Form::sym("%"), e]);
        }
        if let Some(t) = self.peek() {
            if starts_operand(t) {
                let both_refs = matches!(self.toks.get(self.at - 1), Some(Tok::Ref(_)))
                    && matches!(t, Tok::Ref(_));
                return Err(if both_refs {
                    unsupported("the intersection of two references, a space between them")
                } else {
                    malformed("two operands with no operator between them")
                });
            }
        }
        Ok(e)
    }

    /// A primary, and `:` after it to another (KERNEL.8, slice f): a run of
    /// them one form of many operands, `(: a b c)`, the tightest binding of
    /// all; a `:` between two references side by side the scanner reads as
    /// one range, a symbol.
    fn ranged(&mut self) -> Result<Form, Refusal> {
        let mut left = self.primary()?;
        while self.peek() == Some(&Tok::Colon) {
            self.at += 1;
            let right = self.primary()?;
            left = match left {
                Form::List(mut l)
                    if l.items.len() >= 3
                        && matches!(l.items.first(), Some(Form::Sym(h)) if h == ":") =>
                {
                    l.items.push(right);
                    Form::List(l)
                }
                other => Form::list(vec![Form::sym(":"), other, right]),
            };
        }
        Ok(left)
    }

    /// An array constant's values after its `{`, through its `}` (KERNEL.8,
    /// slice f): `(quote (1 2 3))` for one row, `(quote ((1 2) (3 4)))` for
    /// rows, each value a numeral with its sign, a text, a truth value or an
    /// error, every row as long as the first.
    fn array_constant(&mut self) -> Result<Form, Refusal> {
        let mut rows: Vec<Vec<Form>> = vec![Vec::new()];
        loop {
            let value = match self.take() {
                Some(Tok::Num(n)) => Form::Sym(n),
                Some(Tok::Op("-")) => match self.take() {
                    Some(Tok::Num(n)) => Form::Sym(format!("-{n}")),
                    _ => return Err(malformed("a sign before no number in an array constant")),
                },
                Some(Tok::Op("+")) => match self.take() {
                    Some(Tok::Num(n)) => Form::Sym(n),
                    _ => return Err(malformed("a sign before no number in an array constant")),
                },
                Some(Tok::Text(s)) => Form::Str(s),
                Some(Tok::Bool(b)) => Form::sym(if b { "true" } else { "false" }),
                Some(Tok::Err(k)) => Form::list(vec![Form::sym("error"), Form::string(k.text())]),
                Some(_) => {
                    return Err(malformed(
                        "an array constant holding what is no number, text, truth value or error",
                    ))
                }
                None => return Err(malformed("a '{' that never closes")),
            };
            if let Some(row) = rows.last_mut() {
                row.push(value);
            }
            match self.take() {
                Some(Tok::Comma) => {}
                Some(Tok::Semi) => rows.push(Vec::new()),
                Some(Tok::CloseBrace) => break,
                _ => return Err(malformed("a '{' that never closes")),
            }
        }
        let width = rows[0].len();
        if rows.iter().any(|r| r.len() != width) {
            return Err(malformed("an array constant whose rows differ in length"));
        }
        let held = if rows.len() == 1 {
            Form::list(rows.pop().unwrap_or_default())
        } else {
            Form::list(rows.into_iter().map(Form::list).collect())
        };
        Ok(Form::list(vec![Form::sym("quote"), held]))
    }

    fn primary(&mut self) -> Result<Form, Refusal> {
        let Some(t) = self.take() else {
            return Err(malformed("it ends where an operand was wanted"));
        };
        match t {
            Tok::Num(n) => Ok(Form::Sym(n)),
            Tok::Text(s) => Ok(Form::Str(s)),
            Tok::Bool(b) => Ok(Form::sym(if b { "true" } else { "false" })),
            Tok::Err(k) => Ok(Form::list(vec![Form::sym("error"), Form::string(k.text())])),
            Tok::Ref(r) => Ok(reference_form(&r)),
            Tok::Func(name) => {
                self.at += 1; // the `(` the lexer put after the name
                self.nested(|p| p.call(name))
            }
            Tok::Open => {
                let e = self.nested(Self::comparison)?;
                match self.take() {
                    Some(Tok::Close) => Ok(e),
                    Some(Tok::Comma) => Err(unsupported("the union of references, (A1,B2)")),
                    Some(other) => Err(self.stray(&other)),
                    None => Err(malformed("a '(' that never closes")),
                }
            }
            Tok::Close => Err(malformed("a ')' where an operand was wanted")),
            Tok::Comma => Err(malformed("a ',' where an operand was wanted")),
            Tok::Op(op) => Err(malformed(&format!(
                "the operator {op} where an operand was wanted"
            ))),
            Tok::Percent => Err(malformed("a '%' where an operand was wanted")),
            Tok::OpenBrace => self.array_constant(),
            Tok::CloseBrace => Err(malformed("a '}' that opens nothing")),
            Tok::Semi => Err(malformed("a ';' outside an array constant")),
            Tok::Colon => Err(malformed("a ':' where an operand was wanted")),
        }
    }

    /// A call's arguments after its `(`, through its `)`.
    fn call(&mut self, name: String) -> Result<Form, Refusal> {
        let folded = fold(&name);
        if folded.starts_with("_xl") {
            return Err(unsupported(&format!("the file's spelling {name}")));
        }
        if !is_function_name(&name) {
            return Err(unsupported(&format!("the function name {name}")));
        }
        // VLA.6: `last` among them, since the writer writes `(last X)` as a
        // twin and never as `LAST(X)`.
        let reserved = matches!(folded.as_str(), "error" | "ref" | "quote" | "last")
            || headtable::row(&folded).is_some_and(|r| r.formula != FormulaSupport::Yes);
        if reserved {
            return Err(unsupported(&format!(
                "a function named {name}, a word the syntax keeps for a form of its own"
            )));
        }
        let mut args = Vec::new();
        if self.peek() == Some(&Tok::Close) {
            self.at += 1;
        } else {
            loop {
                if matches!(self.peek(), Some(Tok::Comma) | Some(Tok::Close)) {
                    return Err(unsupported("an omitted argument"));
                }
                args.push(self.comparison()?);
                match self.take() {
                    Some(Tok::Comma) => continue,
                    Some(Tok::Close) => break,
                    Some(other) => return Err(self.stray(&other)),
                    None => return Err(malformed(&format!("{name}( never closes"))),
                }
            }
        }
        if self.peek() == Some(&Tok::Open) {
            return Err(unsupported(
                "a function called in place, such as LAMBDA(x,x)(1)",
            ));
        }
        let n = args.len();
        let bad = match folded.as_str() {
            "if" => (!(2..=3).contains(&n)).then_some("IF takes two or three arguments"),
            "not" => (n != 1).then_some("NOT takes one argument"),
            "mod" => (n != 2).then_some("MOD takes two arguments"),
            "lambda" => (n == 0).then_some("LAMBDA takes at least its body"),
            _ => None,
        };
        if let Some(bad) = bad {
            return Err(malformed(bad));
        }
        if folded == "lambda" {
            let body = args.pop().expect("LAMBDA's body, counted above");
            let mut params = Vec::with_capacity(args.len());
            for a in args {
                match &a {
                    Form::Sym(s)
                        if whole_reference(s)
                            .is_some_and(|r| r.kind == Kind::Name && r.sheet_name.is_empty()) =>
                    {
                        params.push(a)
                    }
                    _ => return Err(malformed("a LAMBDA parameter that is no name")),
                }
            }
            return Ok(Form::list(vec![
                Form::sym("lambda"),
                Form::list(params),
                body,
            ]));
        }
        let mut items = Vec::with_capacity(n + 1);
        items.push(Form::Sym(folded));
        items.extend(args);
        Ok(Form::list(items))
    }
}

fn starts_operand(t: &Tok) -> bool {
    matches!(
        t,
        Tok::Num(_)
            | Tok::Text(_)
            | Tok::Bool(_)
            | Tok::Err(_)
            | Tok::Ref(_)
            | Tok::Func(_)
            | Tok::Open
            | Tok::OpenBrace
    )
}

/// A reference or a name as the canonical expression holds it: the bare
/// symbol of its canonical spelling, or `(ref sheet part)` where a symbol
/// cannot hold that spelling.
fn reference_form(r: &FormulaRef) -> Form {
    let canonical = reference_text(r);
    if canonical
        .chars()
        .any(|c| matches!(c, ' ' | '\t' | '\r' | '\n' | '(' | ')' | ';' | '"'))
    {
        let part = match canonical.rfind('!') {
            Some(at) => canonical[at + 1..].to_string(),
            None => canonical.clone(),
        };
        Form::list(vec![
            Form::sym("ref"),
            Form::string(&r.sheet_name),
            Form::string(&part),
        ])
    } else {
        Form::Sym(canonical)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::read_forms;

    fn w(src: &str) -> String {
        write(&read_forms(src).unwrap()[0]).unwrap_or_else(|e| panic!("{src}: {}", e.text))
    }

    fn w_err(src: &str) -> String {
        write(&read_forms(src).unwrap()[0]).unwrap_err().id
    }

    fn r(text: &str) -> String {
        write_datum(&read(text).unwrap_or_else(|e| panic!("{text}: {}", e.text)))
    }

    fn r_err(text: &str) -> String {
        read(text).unwrap_err().id
    }

    #[test]
    fn the_atoms_are_written_in_their_canonical_spelling() {
        assert_eq!(w("5"), "=5");
        assert_eq!(w("1.50"), "=1.50");
        assert_eq!(w(".5"), "=.5");
        assert_eq!(w("1e3"), "=1e3");
        assert_eq!(w("-1"), "=-1");
        assert_eq!(w("\"say \\\"hi\\\"\""), "=\"say \"\"hi\"\"\"");
        assert_eq!(w("\"\""), "=\"\"");
        assert_eq!(w("True"), "=TRUE");
        assert_eq!(w("(error \"#div/0!\")"), "=#DIV/0!");
        assert_eq!(w("a1"), "=A1");
        assert_eq!(w("$a$1:B$2"), "=$A$1:B$2");
        assert_eq!(w("e:E"), "=E:E");
        assert_eq!(w("5:$7"), "=5:$7");
        assert_eq!(w("'Input'!b5"), "=Input!B5");
        assert_eq!(w("Tally.last!B16"), "=Tally.last!B16");
        assert_eq!(w("'2026'!A1"), "='2026'!A1");
        assert_eq!(w("Jan:Dec!A1"), "=Jan:Dec!A1");
        assert_eq!(w("Rate"), "=Rate");
        assert_eq!(w("Model!Local"), "=Model!Local");
        assert_eq!(w("(ref \"Q1 Data\" \"a1\")"), "='Q1 Data'!A1");
        assert_eq!(w("(ref \"Input\" \"B5\")"), "=Input!B5");
        assert_eq!(w("(ref \"C3\")"), "=C3");
    }

    #[test]
    fn precedence_is_written_with_the_fewest_parentheses() {
        assert_eq!(w("(+ 1 (* 2 3))"), "=1+2*3");
        assert_eq!(w("(* (+ 1 2) 3)"), "=(1+2)*3");
        assert_eq!(w("(- a (- b c))"), "=a-(b-c)");
        assert_eq!(w("(- a b c)"), "=a-b-c");
        assert_eq!(w("(- (- a b) c)"), "=a-b-c");
        assert_eq!(w("(^ 2 (^ 3 2))"), "=2^(3^2)");
        assert_eq!(w("(^ 2 3 2)"), "=2^3^2");
        assert_eq!(w("(- (^ x 2))"), "=-(x^2)");
        assert_eq!(w("(^ (- x) 2)"), "=-x^2");
        assert_eq!(w("(^ -2 2)"), "=-2^2");
        assert_eq!(w("(% -5)"), "=(-5)%");
        assert_eq!(w("(- (% 5))"), "=-5%");
        assert_eq!(w("(% (+ a b))"), "=(a+b)%");
        assert_eq!(w("(+ x)"), "=+x");
        assert_eq!(w("(= (& 1 2) \"12\")"), "=1&2=\"12\"");
        assert_eq!(w("(& (= a b) c)"), "=(a=b)&c");
        assert_eq!(w("(< a b c)"), "=a<b<c");
        assert_eq!(w("(* a -1)"), "=a*-1");
        assert_eq!(w("(- 2 -1)"), "=2--1");
        assert_eq!(w("(- (- x))"), "=--x");
        // `#N/A` before a `/` is parenthesized, the language's parser
        // reading `#N/A/x` as one token; the generator's seed 222 found it.
        assert_eq!(w("(/ (error \"#N/A\") x)"), "=(#N/A)/x");
        assert_eq!(w("(/ (^ 2 (error \"#N/A\")) x)"), "=2^(#N/A)/x");
        assert_eq!(w("(/ x (error \"#N/A\") y)"), "=x/(#N/A)/y");
        assert_eq!(w("(/ x (error \"#N/A\"))"), "=x/#N/A");
        assert_eq!(w("(/ (error \"#NUM!\") x)"), "=#NUM!/x");
    }

    #[test]
    fn calls_and_the_syntax_s_own_forms() {
        assert_eq!(
            w("(if (> A1 3) (stdev.s A1:A9) (pi))"),
            "=IF(A1>3,STDEV.S(A1:A9),PI())"
        );
        assert_eq!(w("(SUM a)"), "=SUM(a)");
        assert_eq!(w("(si (> a 1) 1 0)"), "=IF(a>1,1,0)");
        assert_eq!(w("(log10 LOG10)"), "=LOG10(LOG10)");
        assert_eq!(
            w("(let x 1 (lambda (y) (+ x y)))"),
            "=LET(x,1,LAMBDA(y,x+y))"
        );
        assert_eq!(w("(lambda () 1)"), "=LAMBDA(1)");
        assert_eq!(
            w("(and true (not b) (or c d))"),
            "=AND(TRUE,NOT(b),OR(c,d))"
        );
        assert_eq!(w("(mod n 2)"), "=MOD(n,2)");
        assert_eq!(w("(xor (type A1) 1)"), "=XOR(TYPE(A1),1)");
        assert_eq!(w("(row)"), "=ROW()");
        assert_eq!(w("\"=A1+1\""), "=\"=A1+1\"");
    }

    #[test]
    fn the_writer_refuses_what_it_cannot_spell_by_name() {
        assert_eq!(w_err("tax-rate"), "vla-formula-atom-unspelled");
        assert_eq!(w_err("A1+B1"), "vla-formula-atom-unspelled");
        assert_eq!(w_err("#N/A"), "vla-formula-atom-unspelled");
        assert_eq!(w_err("(f :key 1)"), "vla-formula-atom-unspelled");
        assert_eq!(w_err("&HFF"), "vla-formula-numeral-unspelled");
        assert_eq!(w_err("1d5"), "vla-formula-numeral-unspelled");
        assert_eq!(w_err("+5"), "vla-formula-numeral-unspelled");
        assert_eq!(w_err("1e999"), "vla-formula-numeral-unspelled");
        assert_eq!(
            w_err("((lambda (x) x) 1)"),
            "vla-formula-head-not-a-function"
        );
        assert_eq!(w_err("(\"sum\" 1)"), "vla-formula-head-not-a-function");
        assert_eq!(w_err("(a-b 1)"), "vla-formula-head-not-a-function");
        assert_eq!(
            w_err("(_xlfn.xlookup 1)"),
            "vla-formula-head-not-a-function"
        );
        assert_eq!(w_err("(set! x 1)"), "vla-formula-head-refused");
        assert_eq!(w_err("(fijar! x 1)"), "vla-formula-head-refused");
        assert_eq!(w_err("(new x)"), "vla-formula-head-refused");
        assert_eq!(w_err("(if (> x 1) (then 2))"), "vla-formula-head-refused");
        assert_eq!(w_err("(+)"), "vla-formula-arity");
        assert_eq!(w_err("(* a)"), "vla-formula-arity");
        assert_eq!(w_err("(% a b)"), "vla-formula-arity");
        assert_eq!(w_err("(not a b)"), "vla-formula-arity");
        assert_eq!(w_err("(if c)"), "vla-formula-arity");
        assert_eq!(w_err("(mod a)"), "vla-formula-arity");
        assert_eq!(w_err("(lambda x)"), "vla-formula-arity");
        // An array constant (KERNEL.8, slice f), and what it cannot hold.
        assert_eq!(w("(quote (1 2))"), "={1,2}");
        assert_eq!(
            w("(quote ((1 \"a\") (true (error \"#n/a\"))))"),
            "={1,\"a\";TRUE,#N/A}"
        );
        assert_eq!(w("(: a1 (index a:a 5))"), "=A1:INDEX(A:A,5)");
        assert_eq!(w("(+ 1 (: a1 (index a:a 5)))"), "=1+A1:INDEX(A:A,5)");
        for bad in [
            "(quote ((1 2) (3)))",
            "(quote (A1))",
            "(quote ((1) 2))",
            "(quote ())",
            "(quote ((error \"#calc!\")))",
        ] {
            assert_eq!(w_err(bad), "vla-formula-unsupported", "{bad}");
        }
        assert_eq!(w_err("(quote)"), "vla-formula-arity");
        assert_eq!(w_err("(: a1)"), "vla-formula-arity");
        assert_eq!(w_err("(error \"#SPILL!\")"), "vla-formula-unsupported");
        assert_eq!(w_err("A1#"), "vla-formula-unsupported");
        assert_eq!(w_err("(lambda ((x)) x)"), "vla-formula-lambda-params");
        assert_eq!(w_err("(lambda (A1) A1)"), "vla-formula-lambda-params");
        assert_eq!(w_err("(ref \"Data\" \"A1+1\")"), "vla-formula-ref-invalid");
        let long = format!("\"{}\"", "a".repeat(256));
        assert_eq!(w_err(&long), "vla-formula-text-too-long");
        assert_eq!(w(&format!("\"{}\"", "a".repeat(255))).len(), 258);
    }

    #[test]
    fn the_reader_gives_the_canonical_expression() {
        assert_eq!(r("=5"), "5");
        assert_eq!(r("=1.50"), "1.50");
        assert_eq!(r("=-1"), "-1");
        assert_eq!(r("=--1"), "(- -1)");
        assert_eq!(r("=-(1)"), "-1");
        assert_eq!(r("=-5%"), "(- (% 5))");
        assert_eq!(r("=(-5)%"), "(% -5)");
        assert_eq!(r("=\"say \"\"hi\"\"\""), "\"say \\\"hi\\\"\"");
        assert_eq!(r("=true"), "true");
        assert_eq!(r("=#n/a"), "(error \"#N/A\")");
        assert_eq!(r("=a1"), "A1");
        assert_eq!(r("='Input'!B5"), "Input!B5");
        assert_eq!(r("='Q1 Data'!A1"), "(ref \"Q1 Data\" \"A1\")");
        assert_eq!(r("= A1 + 1"), "(+ A1 1)");
        assert_eq!(r("=1+2+3"), "(+ 1 2 3)");
        assert_eq!(r("=1+(2+3)"), "(+ 1 (+ 2 3))");
        assert_eq!(r("=(1+2)+3"), "(+ 1 2 3)");
        assert_eq!(r("=-a-b"), "(- (- a) b)");
        assert_eq!(r("=a<b<c"), "(< a b c)");
        assert_eq!(r("=sum(a1:b2)"), "(sum A1:B2)");
        assert_eq!(r("=LAMBDA(x,y,x+y)"), "(lambda (x y) (+ x y))");
        assert_eq!(r("=LET(x,1,x+1)"), "(let x 1 (+ x 1))");
        assert_eq!(r("=PI()"), "(pi)");
        assert_eq!(r("=#REF!"), "(error \"#REF!\")");
    }

    #[test]
    fn the_reader_tells_malformed_from_unsupported() {
        for bad in [
            "=",
            "=1+",
            "=(1",
            "=)",
            "=\"abc",
            "=1 2",
            "=SUM(1,2",
            "=IF(1)",
            "=NOT(1,2)",
            // Array constants Excel cannot hold (KERNEL.8, slice f).
            "={1,2;3}",
            "={A1}",
            "={1,",
            "={}",
            "=1}",
            "=1;2",
            "=A1:",
        ] {
            assert_eq!(r_err(bad), "vla-formula-read-malformed", "{bad}");
        }
        for valid in [
            "=@A1:A3",
            "=A1 B1",
            "=SUM((A1,B2))",
            "=A1#",
            "=SUM(Sales[Amount])",
            "=[Rates.xlsx]Sheet1!A1",
            "=IF(A1,,B1)",
            "=_xlfn.XLOOKUP(1,A:A,B:B)",
            "=#SPILL!",
            "=LAMBDA(x,x)(1)",
            "=Sheet1!#REF!",
            "=CALL(1)",
            "=REF(1)",
        ] {
            assert_eq!(r_err(valid), "vla-formula-read-unsupported", "{valid}");
        }
        let long = format!("=\"{}\"", "a".repeat(256));
        assert_eq!(r_err(&long), "vla-formula-read-malformed");
    }

    /// F.7's test (VLA.4): no row of the head table is undeclared in a
    /// formula; the writer refuses every head the table refuses, by name;
    /// and every head it declares a formula's is written, or refused only
    /// for its shape.
    #[test]
    fn every_head_of_the_table_is_declared_and_the_writer_holds_to_it() {
        for row in headtable::rows() {
            assert_ne!(
                row.formula,
                FormulaSupport::Undeclared,
                "{} is undeclared",
                row.symbol
            );
            let form = Form::list(vec![Form::sym(&row.symbol), Form::sym("1"), Form::sym("2")]);
            match (row.formula, write(&form)) {
                (FormulaSupport::Refused, Err(e)) => {
                    assert_eq!(e.id, "vla-formula-head-refused", "{}", row.symbol)
                }
                (FormulaSupport::Refused, Ok(t)) => panic!("{} written as {t}", row.symbol),
                (_, Ok(_)) => {}
                (_, Err(e)) => assert!(
                    matches!(
                        e.id.as_str(),
                        "vla-formula-arity"
                            | "vla-formula-unsupported"
                            | "vla-formula-lambda-params"
                    ),
                    "{}: {}",
                    row.symbol,
                    e.id
                ),
            }
        }
    }

    #[test]
    fn writing_and_reading_are_inverse_on_the_canonical_forms() {
        for text in [
            "=IF(B11=1,Day.last!$B$3+$B$6,Hustle.last!B13)",
            "=Tally.last!B1+IF(AND(Input!B1=1,Input.last!B1=0),1,0)",
            "=INT(B15/60)&\":\"&IF(MOD(B15,60)=0,\"00\",\"30\")",
            "=(COLUMN()-Gate!$A2)^2<=Gate!$B2",
            "='Q1 Data'!A1+'2026'!B2",
            "=-(x^2)+-x^2",
            "=(-5)%+-5%",
            "=a-(b-c)",
            "=2^(3^2)",
            // KERNEL.8, slice f: array constants and `:` between expressions.
            "={1,2,3}",
            "={1,-2.5;\"a\",TRUE}",
            "={#N/A}",
            "=SUM({1;2;3})",
            "=A1:INDEX(A:A,5)",
            "=SUM(INDEX(A1:A9,2):INDEX(A1:A9,5))",
            "=-A1:INDEX(A:A,5)",
            "=A1:INDEX(A:A,5)%",
            "=Data!A1:Data!B2",
        ] {
            let v = read(text).unwrap();
            assert_eq!(write(&v).unwrap(), text, "{text}");
            assert_eq!(read(&write(&v).unwrap()).unwrap(), v, "{text}");
        }
    }
}
