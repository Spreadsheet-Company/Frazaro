//! Values and their coercions (KERNEL.7, 2026-10-08): what a formula
//! computes to, and how one kind becomes another when an operator or a
//! function wants it.
//!
//! The kinds are a cell's: a number (an IEEE 754 double, Excel's own and
//! VBA's `Double`; `Alonzo/SPEC.md` section 6, rule 5), a text, a truth
//! value, one of the seven error values a cell can hold, and the empty
//! cell, which a reference to nothing yields and which a formula never
//! returns: a formula's empty result settles to 0, as Excel's does.
//! Coercion is where a spreadsheet's semantics hide (the roadmap's catch
//! for `KERNEL.7`), so each rule here is one method with Excel's rule in
//! its comment, and each enters the oracle with a fixture that shows it
//! (`scripts/recalc/`). The language's own functions carry no tolerance;
//! the agreement of a computed value with a host's cached one carries the
//! one stated tolerance, [`agrees`], fifteen significant digits, which is
//! the precision Excel documents for its arithmetic (`Alonzo/docs/ALGEBRA.md`
//! section 7, crack 1: written here before the oracle was).

use crate::intrinsics::fold;
use crate::rows;
use crate::sheet::number_text;

/// The seven error values a cell can hold, Excel's spelling.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum ErrorKind {
    /// `#NULL!`: an intersection of two ranges that do not meet.
    Null,
    /// `#DIV/0!`
    Div0,
    /// `#VALUE!`: an operand of the wrong kind.
    Value,
    /// `#REF!`: a reference to a cell or a sheet that is not there.
    Ref,
    /// `#NAME?`: a name nothing defines.
    Name,
    /// `#NUM!`: a number no cell can hold.
    Num,
    /// `#N/A`
    NA,
}

impl ErrorKind {
    /// The text a cell shows, which is the text the file stores.
    pub fn text(self) -> &'static str {
        match self {
            ErrorKind::Null => "#NULL!",
            ErrorKind::Div0 => "#DIV/0!",
            ErrorKind::Value => "#VALUE!",
            ErrorKind::Ref => "#REF!",
            ErrorKind::Name => "#NAME?",
            ErrorKind::Num => "#NUM!",
            ErrorKind::NA => "#N/A",
        }
    }

    /// The kind an error text names, without case; `None` for an error this
    /// version does not hold (`#SPILL!`, `#CALC!`, `#GETTING_DATA`).
    pub fn parse(text: &str) -> Option<ErrorKind> {
        match fold(text.trim()).as_str() {
            "#null!" => Some(ErrorKind::Null),
            "#div/0!" => Some(ErrorKind::Div0),
            "#value!" => Some(ErrorKind::Value),
            "#ref!" => Some(ErrorKind::Ref),
            "#name?" => Some(ErrorKind::Name),
            "#num!" => Some(ErrorKind::Num),
            "#n/a" => Some(ErrorKind::NA),
            _ => None,
        }
    }
}

/// What a cell holds or a formula computes.
#[derive(Clone, Debug, PartialEq)]
pub enum Value {
    Number(f64),
    Text(String),
    Bool(bool),
    Error(ErrorKind),
    /// A cell with nothing in it. An operand that is empty reads as 0, as
    /// `""` or as FALSE, whichever the operator wants; a formula whose
    /// result is empty settles to 0 ([`Value::settled`]).
    Empty,
}

impl Value {
    /// The number an arithmetic operator reads: a number itself; a truth
    /// value 1 or 0; an empty cell 0; a text its number when it spells one
    /// (`"5"`, `" 1e3 "`, `"50%"`) and `#VALUE!` otherwise; an error the
    /// error. Excel's rule for `+`, `-`, `*`, `/`, `^` and a function's
    /// numeric argument given directly (a value read out of a range follows
    /// the range rule instead, which each function states).
    pub fn number(&self) -> Result<f64, ErrorKind> {
        match self {
            Value::Number(v) => Ok(*v),
            Value::Bool(b) => Ok(if *b { 1.0 } else { 0.0 }),
            Value::Empty => Ok(0.0),
            Value::Text(t) => text_to_number(t).ok_or(ErrorKind::Value),
            Value::Error(k) => Err(*k),
        }
    }

    /// The text `&` reads: a text itself; a number in Excel's General
    /// spelling ([`number_to_text`]); a truth value `TRUE` or `FALSE`; an
    /// empty cell `""`; an error the error.
    pub fn text(&self) -> Result<String, ErrorKind> {
        match self {
            Value::Text(t) => Ok(t.clone()),
            Value::Number(v) => Ok(number_to_text(*v)),
            Value::Bool(true) => Ok("TRUE".to_string()),
            Value::Bool(false) => Ok("FALSE".to_string()),
            Value::Empty => Ok(String::new()),
            Value::Error(k) => Err(*k),
        }
    }

    /// The truth value `IF`, `AND`, `OR` and `NOT` read: a truth value
    /// itself; a number true unless 0; an empty cell false; a text only
    /// when it spells `TRUE` or `FALSE`, `#VALUE!` otherwise; an error the
    /// error.
    pub fn truth(&self) -> Result<bool, ErrorKind> {
        match self {
            Value::Bool(b) => Ok(*b),
            Value::Number(v) => Ok(*v != 0.0),
            Value::Empty => Ok(false),
            Value::Text(t) => match fold(t).as_str() {
                "true" => Ok(true),
                "false" => Ok(false),
                _ => Err(ErrorKind::Value),
            },
            Value::Error(k) => Err(*k),
        }
    }

    /// A formula's result as its cell holds it: an empty result is 0, as
    /// `=A1` over an empty A1 shows 0 and is a number to `ISNUMBER`.
    pub fn settled(self) -> Value {
        match self {
            Value::Empty => Value::Number(0.0),
            other => other,
        }
    }

    /// The value as a datum, spelled as a `cell` row spells one
    /// (`crate::rows::datum`): a number in its shortest spelling, a text
    /// quoted, `true` or `false`, `(error "#DIV/0!")`; a number no cell can
    /// hold is `#NUM!`, as the writer spells it; an empty value is 0.
    pub fn spell(&self) -> String {
        match self {
            Value::Number(v) if *v == 0.0 => "0".to_string(),
            Value::Number(v) => match number_text(*v) {
                Some(text) => text,
                None => rows::datum(&rows::Value::Error(ErrorKind::Num.text().to_string())),
            },
            Value::Text(t) => rows::quoted(t),
            Value::Bool(true) => "true".to_string(),
            Value::Bool(false) => "false".to_string(),
            Value::Error(k) => rows::datum(&rows::Value::Error(k.text().to_string())),
            Value::Empty => "0".to_string(),
        }
    }

    /// A cached value as the file holds it, read as a computed kind: the
    /// file's number text parsed, a text, a truth value, one of the seven
    /// errors. `None` for a value this version does not compare: a date
    /// cell, a number text that is not a number, an error outside the
    /// seven.
    pub fn from_file(v: &rows::Value) -> Option<Value> {
        match v {
            rows::Value::Number(text) => {
                let n: f64 = text.trim().parse().ok()?;
                n.is_finite().then_some(Value::Number(n))
            }
            rows::Value::Text(t) => Some(Value::Text(t.clone())),
            rows::Value::Bool(b) => Some(Value::Bool(*b)),
            rows::Value::Error(text) => ErrorKind::parse(text).map(Value::Error),
            rows::Value::Date(_) => None,
        }
    }
}

/// A number rounded to fifteen significant digits, the precision Excel
/// documents for its arithmetic; 0 and a value that is not finite stay as
/// they are.
pub fn round15(v: f64) -> f64 {
    if v == 0.0 || !v.is_finite() {
        return v;
    }
    format!("{v:.14e}").parse().unwrap_or(v)
}

/// A number's text as Excel's General format spells it for `&` and for a
/// text argument: fifteen significant digits at most, fixed while the
/// decimal exponent is between -14 and 14, and Excel's scientific spelling
/// outside that, `1E+20`, `1.5E-15`, the exponent signed and at least two
/// digits. Measured: Excel's save of the subset fixture spells 1E-7 fixed,
/// `0.0000001`, and 1E+20 and 1E+21 scientific
/// (`scripts/recalc/subset_saved.xlsx`, `Output!C21` to `C23`). The two
/// boundaries are LibreOffice's automatic number format's (`nExp <= -15 ||
/// nExp >= 15` in `sal/rtl/strtmpl.hxx`, `doubleToString`), which both
/// measured points agree with; `scripts/recalc/edges.txt` holds them for an
/// Excel save to witness.
pub fn number_to_text(v: f64) -> String {
    if v == 0.0 {
        return "0".to_string();
    }
    let r = round15(v);
    let sci = format!("{r:.14e}");
    let (mantissa, exponent) = sci.split_once('e').unwrap_or((sci.as_str(), "0"));
    let exponent: i32 = exponent.parse().unwrap_or(0);
    if (-14..=14).contains(&exponent) {
        // Fixed: the shortest spelling of a value already rounded to fifteen
        // significant digits is those digits, and Rust never writes an
        // exponent in this form.
        return format!("{r}");
    }
    let mantissa = mantissa.trim_end_matches('0').trim_end_matches('.');
    format!(
        "{mantissa}E{}{:02}",
        if exponent < 0 { '-' } else { '+' },
        exponent.abs()
    )
}

/// The number a text spells, as an arithmetic operator reads it: blanks
/// around it ignored, an optional sign, digits with at most one `.`, an
/// optional exponent, an optional `%` at the end dividing by 100. Nothing
/// else is a number here: a thousands separator, a currency sign and a
/// date are a locale's, and the locale is the host's.
pub fn text_to_number(text: &str) -> Option<f64> {
    let t = text.trim();
    let (t, percent) = match t.strip_suffix('%') {
        Some(rest) => (rest.trim_end(), true),
        None => (t, false),
    };
    let bytes = t.as_bytes();
    if bytes.is_empty() {
        return None;
    }
    let mut i = 0;
    if bytes[i] == b'+' || bytes[i] == b'-' {
        i += 1;
    }
    let mut digits = 0;
    let mut dot = false;
    while i < bytes.len() {
        match bytes[i] {
            b'0'..=b'9' => digits += 1,
            b'.' if !dot => dot = true,
            _ => break,
        }
        i += 1;
    }
    if digits == 0 {
        return None;
    }
    if i < bytes.len() && (bytes[i] == b'e' || bytes[i] == b'E') {
        let mut j = i + 1;
        if j < bytes.len() && (bytes[j] == b'+' || bytes[j] == b'-') {
            j += 1;
        }
        let first = j;
        while j < bytes.len() && bytes[j].is_ascii_digit() {
            j += 1;
        }
        if j == first {
            return None;
        }
        i = j;
    }
    if i != bytes.len() {
        return None;
    }
    let v: f64 = t.parse().ok()?;
    let v = if percent { v / 100.0 } else { v };
    v.is_finite().then_some(v)
}

/// Whether a computed value agrees with the value a host saved in the
/// file: two numbers when they round to the same fifteen significant
/// digits (the stated tolerance), two texts when equal, two truth values
/// or two errors when the same; `None` when the file's value is one this
/// version does not compare (a date; see [`Value::from_file`]).
pub fn agrees(computed: &Value, cached: &rows::Value) -> Option<bool> {
    let cached = Value::from_file(cached)?;
    let computed = computed.clone().settled();
    Some(match (&computed, &cached) {
        (Value::Number(a), Value::Number(b)) => {
            let key = |x: f64| format!("{:.14e}", if x == 0.0 { 0.0 } else { round15(x) });
            key(*a) == key(*b)
        }
        (Value::Text(a), Value::Text(b)) => a == b,
        (Value::Bool(a), Value::Bool(b)) => a == b,
        (Value::Error(a), Value::Error(b)) => a == b,
        _ => false,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn the_seven_errors_spell_and_parse() {
        for k in [
            ErrorKind::Null,
            ErrorKind::Div0,
            ErrorKind::Value,
            ErrorKind::Ref,
            ErrorKind::Name,
            ErrorKind::Num,
            ErrorKind::NA,
        ] {
            assert_eq!(ErrorKind::parse(k.text()), Some(k));
            assert_eq!(ErrorKind::parse(&k.text().to_lowercase()), Some(k));
        }
        assert_eq!(ErrorKind::parse("#SPILL!"), None);
    }

    #[test]
    fn numbers_are_read_as_excel_reads_an_operand() {
        assert_eq!(Value::Text("5".into()).number(), Ok(5.0));
        assert_eq!(Value::Text(" 1e3 ".into()).number(), Ok(1000.0));
        assert_eq!(Value::Text("50%".into()).number(), Ok(0.5));
        assert_eq!(Value::Text("-.5".into()).number(), Ok(-0.5));
        assert_eq!(Value::Text("abc".into()).number(), Err(ErrorKind::Value));
        assert_eq!(Value::Text("".into()).number(), Err(ErrorKind::Value));
        assert_eq!(Value::Text("1,000".into()).number(), Err(ErrorKind::Value));
        assert_eq!(Value::Bool(true).number(), Ok(1.0));
        assert_eq!(Value::Empty.number(), Ok(0.0));
        assert_eq!(Value::Error(ErrorKind::NA).number(), Err(ErrorKind::NA));
    }

    #[test]
    fn texts_are_excel_s_general_spelling() {
        assert_eq!(number_to_text(15.0), "15");
        assert_eq!(number_to_text(0.1 + 0.2), "0.3");
        assert_eq!(number_to_text(1.0 / 3.0), "0.333333333333333");
        assert_eq!(number_to_text(-2.5), "-2.5");
        assert_eq!(number_to_text(0.000001), "0.000001");
        // Measured: Excel's save of the subset fixture (subset_saved.xlsx,
        // Output!C21 to C23).
        assert_eq!(number_to_text(0.0000001), "0.0000001");
        assert_eq!(number_to_text(1e20), "1E+20");
        assert_eq!(number_to_text(1e21), "1E+21");
        // The boundaries the rule takes from LibreOffice, held for Excel by
        // scripts/recalc/edges.txt.
        assert_eq!(number_to_text(1e14), "100000000000000");
        assert_eq!(number_to_text(1e15), "1E+15");
        assert_eq!(number_to_text(999_999_999_999_999.0), "999999999999999");
        assert_eq!(number_to_text(999_999_999_999_999.9), "1E+15");
        assert_eq!(number_to_text(2f64.powi(57)), "1.44115188075856E+17");
        assert_eq!(number_to_text(1e-14), "0.00000000000001");
        assert_eq!(number_to_text(1e-15), "1E-15");
        assert_eq!(number_to_text(1.5e-7), "0.00000015");
        assert_eq!(number_to_text(-0.0000001), "-0.0000001");
        assert_eq!(
            number_to_text(1.0 / 3.0 * 1e-10),
            "0.0000000000333333333333333"
        );
        assert_eq!(number_to_text(2f64.sqrt()), "1.4142135623731");
        assert_eq!(number_to_text(0.0), "0");
        assert_eq!(Value::Bool(true).text(), Ok("TRUE".to_string()));
        assert_eq!(Value::Empty.text(), Ok(String::new()));
    }

    #[test]
    fn truth_is_read_from_numbers_and_the_two_words() {
        assert_eq!(Value::Number(2.0).truth(), Ok(true));
        assert_eq!(Value::Number(0.0).truth(), Ok(false));
        assert_eq!(Value::Text("true".into()).truth(), Ok(true));
        assert_eq!(Value::Text("yes".into()).truth(), Err(ErrorKind::Value));
        assert_eq!(Value::Empty.truth(), Ok(false));
    }

    #[test]
    fn a_value_spells_as_a_cell_row_does() {
        assert_eq!(Value::Number(15.0).spell(), "15");
        assert_eq!(Value::Number(0.6).spell(), "0.6");
        assert_eq!(Value::Text("big".into()).spell(), "\"big\"");
        assert_eq!(Value::Bool(false).spell(), "false");
        assert_eq!(Value::Error(ErrorKind::Div0).spell(), "(error \"#DIV/0!\")");
        assert_eq!(Value::Number(f64::INFINITY).spell(), "(error \"#NUM!\")");
        assert_eq!(Value::Empty.spell(), "0");
        assert_eq!(Value::Empty.settled(), Value::Number(0.0));
    }

    #[test]
    fn agreement_carries_the_stated_tolerance_and_nothing_more() {
        let cached = |s: &str| rows::Value::Number(s.to_string());
        assert_eq!(
            agrees(&Value::Number(0.1 + 0.2), &cached("0.3")),
            Some(true)
        );
        assert_eq!(
            agrees(&Value::Number(0.1 + 0.2), &cached("0.30000000000000004")),
            Some(true)
        );
        assert_eq!(agrees(&Value::Number(15.0), &cached("15")), Some(true));
        assert_eq!(
            agrees(&Value::Number(15.0), &cached("15.000000000001")),
            Some(false)
        );
        assert_eq!(agrees(&Value::Empty, &cached("0")), Some(true));
        assert_eq!(
            agrees(&Value::Text("big".into()), &rows::Value::Text("big".into())),
            Some(true)
        );
        assert_eq!(
            agrees(&Value::Text("big".into()), &rows::Value::Text("Big".into())),
            Some(false)
        );
        assert_eq!(
            agrees(
                &Value::Error(ErrorKind::Div0),
                &rows::Value::Error("#DIV/0!".into())
            ),
            Some(true)
        );
        assert_eq!(
            agrees(&Value::Number(1.0), &rows::Value::Date("2026-10-08".into())),
            None
        );
        assert_eq!(
            agrees(&Value::Number(1.0), &rows::Value::Bool(true)),
            Some(false)
        );
    }
}
