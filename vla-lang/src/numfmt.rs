//! A number format's code (KERNEL.47, 2026-10-17): Excel's own spelling of
//! how a number is drawn, beside the two formats the grid had, `general`
//! and `text`, so that a cell holding 1 under `0.00` draws `1.00`. The view
//! record prints a code in its `format` row and a number's text under it in
//! a `text` row (`crate::view`), so that a viewport draws the text and
//! computes nothing, and the writer gives Excel the code (`frazaro-core`).
//!
//! What this version takes, in Excel's syntax as `Range.NumberFormat` reads
//! it, spelled the same on every machine (the point `.`, the thousands
//! `,`): one to four sections by `;`, a positive number's, a negative's,
//! zero's and a text's. In a number's section, the digit placeholders `0`,
//! a digit or 0, and `#`, a digit or nothing; one point; `,` between
//! integer placeholders for the thousands, and after the last placeholder
//! for each thousandth; `%` for each hundredfold; and literals: a quoted
//! text, `\` and the character after it, the characters Excel draws
//! unquoted (`$ - + ( ) : ! ^ & ' ~ { } < > =` and the space), and any
//! character outside ASCII, none of which means anything in a code, so that
//! `▲0.00;▼0.00;=0.00` draws a glyph beside its number. In a text's section,
//! literals and `@`, the text itself. A bare `$` is a dollar here, the
//! record having no reading machine whose currency it could be. Anything
//! else is refused by name ([`Unsupported`]): a bracket (a color, a
//! condition, a locale), a date's or a time's letters, an exponent, a
//! fraction, a width or a fill, `General` inside a code.
//!
//! How a number becomes its text, as Excel draws it: the section is chosen
//! by the number's sign before it is rounded (one section for every number,
//! a negative drawn with its `-`, but for one that draws as zero; two, the
//! second for below zero, drawn without its `-`; three, the third for
//! zero); the number's fifteen significant digits, the precision Excel
//! documents and `round15`, `ROUND` and `&` already take, the nearest and a
//! tie toward zero, are moved in decimal for each `%` and each thousandth,
//! never by a multiply in binary; then rounded half away from zero to the
//! section's decimals, the digits past the fifteenth drawn as zeros. So
//! 2.675 and 1.005 under `0.00` are `2.68` and `1.01`, where Rust's own
//! `{:.2}` rounds the binary value with ties to even and spells them `2.67`
//! and `1.00`; -0.4 under `0` is `0`; and 1234567890123455 under `0` is
//! `1234567890123450`. The digits come from the shortest spelling when it
//! has fifteen significant digits or fewer, which is then the fifteen-digit
//! rounding exactly (a double's half step is at most a fifth of half the
//! fifteenth digit's unit), and from `{:.14e}` when it has more. The two
//! rules after "but" and "toward zero" are Excel's word against the draft's
//! prediction, from the reference's first run of the golden (2026-10-17).
//! The reference is Excel itself: `scripts/formats.txt` and the text Excel
//! draws for each of its cases, `scripts/formats_golden.txt`, written by
//! `VlaWriteFormatsGolden` (VLA_Tests.bas), which this module's test reads.

use std::fmt;
use std::sync::Arc;

use crate::calc::Value;

/// The longest code Excel takes, in characters.
pub const MOST_CHARS: usize = 255;

/// The most decimals a section draws, Excel's ceiling.
pub const MOST_DECIMALS: usize = 30;

/// The codes of Excel's built-in formats that Excel draws alike in every
/// locale, by their ids in a styles part. 0 is General and 49 is Text, the
/// grid's two symbols; 5 to 8, 14 and the rest draw by the reading
/// machine's locale and are never written for a code.
const BUILTINS: [(&str, u32); 8] = [
    ("0", 1),
    ("0.00", 2),
    ("#,##0", 3),
    ("#,##0.00", 4),
    ("0%", 9),
    ("0.00%", 10),
    ("#,##0 ;(#,##0)", 37),
    ("#,##0.00;(#,##0.00)", 39),
];

/// A code this version draws, parsed once and kept with its text as
/// written. Two codes are the same exactly when their texts are.
#[derive(Clone)]
pub struct Code(Arc<Parsed>);

struct Parsed {
    text: String,
    sections: Vec<Section>,
}

/// What a code holds that this version does not draw: the construct, as
/// `grid-format-code` names it, and what draws instead.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Unsupported {
    pub what: String,
    pub fix: String,
}

impl Unsupported {
    fn new(what: impl Into<String>, fix: impl Into<String>) -> Unsupported {
        Unsupported {
            what: what.into(),
            fix: fix.into(),
        }
    }
}

/// One piece of a section as it draws.
#[derive(Clone, Debug, PartialEq)]
enum Part {
    /// A digit placeholder: `0` (true) draws a digit or 0, `#` a digit or
    /// nothing.
    Digit(bool),
    Point,
    Percent,
    Literal(String),
    /// The text itself, in a text's section.
    At,
}

/// One section, read for drawing.
#[derive(Clone, Debug, Default)]
struct Section {
    parts: Vec<Part>,
    /// Whether the section holds a digit placeholder at all.
    placeholders: bool,
    /// The placeholders before the point.
    integer: usize,
    /// The integer digits always drawn: the placeholders from the first
    /// `0` before the point to the point.
    min_integer: usize,
    /// The placeholders after the point, `0` true.
    decimals: Vec<bool>,
    /// Whether the integer digits are grouped by thousands.
    grouping: bool,
    /// The places the point moves right before rounding: 2 for each `%`,
    /// -3 for each thousandth.
    shift: i32,
}

/// A token of a code, before its section is read.
#[derive(Clone, Debug, PartialEq)]
enum Token {
    Digit(bool),
    Point,
    Comma,
    Percent,
    Literal(String),
    At,
}

impl Code {
    /// A code as written, or what it holds that this version does not draw.
    pub fn parse(text: &str) -> Result<Code, Unsupported> {
        if text.is_empty() {
            return Err(Unsupported::new(
                "no code at all",
                "write general for Excel's General format",
            ));
        }
        if text.eq_ignore_ascii_case("general") {
            return Err(Unsupported::new(
                "Excel's General format as a code",
                "write it general, a symbol, unquoted",
            ));
        }
        if text == "@" {
            return Err(Unsupported::new(
                "Excel's Text format as a code",
                "write it text, a symbol, unquoted",
            ));
        }
        if text.chars().count() > MOST_CHARS {
            return Err(Unsupported::new(
                format!("more than {MOST_CHARS} characters"),
                format!("shorten it to Excel's {MOST_CHARS}"),
            ));
        }
        let raw = tokens(text)?;
        if raw.len() > 4 {
            return Err(Unsupported::new(
                "a fifth section",
                "keep four, a positive number's, a negative's, zero's and a text's",
            ));
        }
        let mut sections = Vec::with_capacity(raw.len());
        for (i, toks) in raw.into_iter().enumerate() {
            sections.push(if i == 3 {
                text_section(toks)?
            } else {
                number_section(toks)?
            });
        }
        Ok(Code(Arc::new(Parsed {
            text: text.to_string(),
            sections,
        })))
    }

    /// The code as written.
    pub fn as_str(&self) -> &str {
        &self.0.text
    }

    /// The id of the built-in format the code is, where Excel has one that
    /// it draws alike in every locale.
    pub fn builtin(&self) -> Option<u32> {
        BUILTINS
            .iter()
            .find(|(code, _)| *code == self.0.text)
            .map(|&(_, id)| id)
    }

    /// The text a value draws as under the code, as Excel's `Range.Text`
    /// gives it in a column wide enough: a number's, and a text's where the
    /// code has a text's section; `None` for a text under a code without
    /// one, a truth value, an error and an empty value, each drawn as the
    /// record spells it, and for a number no cell can hold.
    pub fn text(&self, value: &Value) -> Option<String> {
        match value {
            Value::Number(n) => self.number(*n),
            Value::Text(t) => self.0.sections.get(3).map(|s| s.text_of(t)),
            _ => None,
        }
    }

    fn number(&self, x: f64) -> Option<String> {
        if !x.is_finite() {
            return None;
        }
        let sections = &self.0.sections;
        let numeric = sections.len().min(3);
        let (section, minus) = if x == 0.0 {
            // 0 and -0 alike: zero's section where there is one.
            (&sections[if numeric == 3 { 2 } else { 0 }], false)
        } else if x < 0.0 {
            if numeric >= 2 {
                (&sections[1], false)
            } else {
                (&sections[0], true)
            }
        } else {
            (&sections[0], false)
        };
        let mut out = String::new();
        let zero = section.draw(x.abs(), &mut out);
        // One section's minus, before every literal; none where the number
        // draws as zero, as Excel draws -0.4 under 0 as 0 (the reference's
        // golden, 2026-10-17).
        if minus && !zero {
            out.insert(0, '-');
        }
        Some(out)
    }
}

impl PartialEq for Code {
    fn eq(&self, other: &Code) -> bool {
        self.0.text == other.0.text
    }
}

impl Eq for Code {}

impl fmt::Debug for Code {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_tuple("Code").field(&self.0.text).finish()
    }
}

/// A code's tokens, a list a section, split at each `;` outside quotes; or
/// the first thing it holds that this version does not draw.
fn tokens(text: &str) -> Result<Vec<Vec<Token>>, Unsupported> {
    let chars: Vec<char> = text.chars().collect();
    let mut sections: Vec<Vec<Token>> = vec![Vec::new()];
    let mut i = 0;
    while i < chars.len() {
        let c = chars[i];
        let toks = sections.last_mut().expect("one section at least");
        match c {
            ';' => sections.push(Vec::new()),
            '"' => {
                let start = i + 1;
                let Some(len) = chars[start..].iter().position(|&x| x == '"') else {
                    return Err(Unsupported::new("a quote left open", "close it"));
                };
                let quoted: String = chars[start..start + len].iter().collect();
                if quoted.chars().any(is_control) {
                    return Err(control());
                }
                toks.push(Token::Literal(quoted));
                i = start + len;
            }
            '\\' => {
                let Some(&next) = chars.get(i + 1) else {
                    return Err(Unsupported::new(
                        "a backslash with nothing after it",
                        "drop it, or write after it the character it draws",
                    ));
                };
                if is_control(next) {
                    return Err(control());
                }
                toks.push(Token::Literal(next.to_string()));
                i += 1;
            }
            '0' => toks.push(Token::Digit(true)),
            '#' => toks.push(Token::Digit(false)),
            '.' => toks.push(Token::Point),
            ',' => toks.push(Token::Comma),
            '%' => toks.push(Token::Percent),
            '@' => toks.push(Token::At),
            '$' | '-' | '+' | '(' | ')' | ':' | '!' | '^' | '&' | '\'' | '~' | '{' | '}' | '<'
            | '>' | '=' | ' ' => toks.push(Token::Literal(c.to_string())),
            '[' => {
                let bracket: String = match chars[i..].iter().position(|&x| x == ']') {
                    Some(len) => chars[i..=i + len].iter().collect(),
                    None => chars[i..].iter().collect(),
                };
                return Err(Unsupported::new(
                    format!("{bracket}, a color, a condition, a locale or an elapsed time"),
                    "drop it",
                ));
            }
            '?' => {
                return Err(Unsupported::new(
                    "?, a digit's width",
                    "write 0 or # in its place",
                ))
            }
            '/' => {
                return Err(Unsupported::new(
                    "/, a fraction's bar",
                    "quote it, \"/\", to draw it as written",
                ))
            }
            '_' => {
                return Err(Unsupported::new(
                    "_, a character's width",
                    "drop it, or quote a space in its place",
                ))
            }
            '*' => return Err(Unsupported::new("*, a fill", "drop it")),
            c if c.is_ascii_alphabetic() => return Err(letter(&chars, i)),
            c if c.is_ascii_digit() => {
                return Err(Unsupported::new(
                    format!("{c}, a digit outside quotes"),
                    format!("quote it, \"{c}\", to draw it as written"),
                ))
            }
            c if is_control(c) => return Err(control()),
            c if c.is_ascii() => {
                return Err(Unsupported::new(
                    format!("{c}, a character Excel draws only quoted"),
                    format!("quote it, \"{c}\", to draw it as written"),
                ))
            }
            c => toks.push(Token::Literal(c.to_string())),
        }
        i += 1;
    }
    Ok(sections)
}

fn is_control(c: char) -> bool {
    (c as u32) < 0x20 || c == '\u{7F}'
}

fn control() -> Unsupported {
    Unsupported::new("a control character", "drop it")
}

/// A letter outside quotes: `General` inside a code, an exponent, a date's
/// or a time's letter, or a letter Excel draws only quoted.
fn letter(chars: &[char], i: usize) -> Unsupported {
    let word: String = chars[i..].iter().take(7).collect();
    if word.eq_ignore_ascii_case("general") {
        return Unsupported::new(
            "General inside a code",
            "write the digits with 0 and # in its place",
        );
    }
    let c = chars[i];
    let sign = chars.get(i + 1).copied();
    if (c == 'E' || c == 'e') && (sign == Some('+') || sign == Some('-')) {
        let sign = sign.unwrap_or('+');
        return Unsupported::new(
            format!("{c}{sign}, a scientific exponent"),
            "write the digits with 0 and # alone",
        );
    }
    match c.to_ascii_lowercase() {
        'y' | 'm' | 'd' | 'h' | 's' | 'a' | 'p' | 'b' | 'e' | 'g' => Unsupported::new(
            format!("{c}, a date's or a time's letter"),
            format!("quote it, \"{c}\", to draw it as written; this version draws no date or time"),
        ),
        _ => Unsupported::new(
            format!("{c}, a letter outside quotes"),
            format!("quote it, \"{c}\", to draw it as written"),
        ),
    }
}

/// A number's section from its tokens.
fn number_section(toks: Vec<Token>) -> Result<Section, Unsupported> {
    if toks.contains(&Token::At) {
        return Err(Unsupported::new(
            "@ before the fourth section",
            "put a text's section fourth, after zero's",
        ));
    }
    if toks.iter().filter(|t| **t == Token::Point).count() > 1 {
        return Err(Unsupported::new(
            "a second point",
            "keep one, or quote the other",
        ));
    }
    let point = toks.iter().position(|t| *t == Token::Point);
    let is_digit = |t: &Token| matches!(t, Token::Digit(_));
    let before_point = |k: usize| point.map_or(true, |p| k < p);
    let mut s = Section {
        placeholders: toks.iter().any(is_digit),
        ..Section::default()
    };
    let mut zeros: Vec<bool> = Vec::new();
    for (k, t) in toks.iter().enumerate() {
        match t {
            Token::Digit(zero) => {
                if before_point(k) {
                    s.integer += 1;
                    zeros.push(*zero);
                } else {
                    s.decimals.push(*zero);
                }
                s.parts.push(Part::Digit(*zero));
            }
            Token::Point | Token::Percent if !s.placeholders => {
                let what = if *t == Token::Point { "." } else { "%" };
                return Err(Unsupported::new(
                    format!("{what} beside no digit placeholder"),
                    format!("write a digit with 0 or #, or quote it, \"{what}\""),
                ));
            }
            Token::Point => s.parts.push(Part::Point),
            Token::Percent => {
                s.shift += 2;
                s.parts.push(Part::Percent);
            }
            Token::Literal(text) => match s.parts.last_mut() {
                Some(Part::Literal(last)) => last.push_str(text),
                _ => s.parts.push(Part::Literal(text.clone())),
            },
            Token::Comma => {
                if !toks[..k].iter().any(is_digit) {
                    return Err(Unsupported::new(
                        "a comma before any digit placeholder",
                        "drop it, or quote it, \",\"",
                    ));
                }
                match toks[k + 1..].iter().position(is_digit).map(|j| k + 1 + j) {
                    // After the last placeholder: a thousandth.
                    None => s.shift -= 3,
                    // Between integer placeholders: the thousands.
                    Some(j) if before_point(j) => s.grouping = true,
                    Some(_) if !before_point(k) => {
                        return Err(Unsupported::new(
                            "a comma among the decimals",
                            "drop it, or quote it, \",\"",
                        ))
                    }
                    Some(_) => {
                        return Err(Unsupported::new(
                            "a comma just before the point",
                            "move it after the last digit placeholder for thousandths, or drop it",
                        ))
                    }
                }
            }
            Token::At => unreachable!("refused above"),
        }
    }
    if s.decimals.len() > MOST_DECIMALS {
        return Err(Unsupported::new(
            format!("more than {MOST_DECIMALS} decimals"),
            format!("keep {MOST_DECIMALS}, Excel's most"),
        ));
    }
    s.min_integer = zeros
        .iter()
        .position(|zero| *zero)
        .map_or(0, |first| zeros.len() - first);
    Ok(s)
}

/// A text's section from its tokens: literals and `@` alone.
fn text_section(toks: Vec<Token>) -> Result<Section, Unsupported> {
    let mut s = Section::default();
    for t in toks {
        let what = match t {
            Token::Literal(text) => {
                match s.parts.last_mut() {
                    Some(Part::Literal(last)) => last.push_str(&text),
                    _ => s.parts.push(Part::Literal(text)),
                }
                continue;
            }
            Token::At => {
                s.parts.push(Part::At);
                continue;
            }
            Token::Digit(true) => "0",
            Token::Digit(false) => "#",
            Token::Point => ".",
            Token::Comma => ",",
            Token::Percent => "%",
        };
        return Err(Unsupported::new(
            format!("{what} in a text's section"),
            format!("quote it, \"{what}\", to draw it as written"),
        ));
    }
    Ok(s)
}

impl Section {
    /// A text drawn by a text's section.
    fn text_of(&self, text: &str) -> String {
        let mut out = String::new();
        for p in &self.parts {
            match p {
                Part::Literal(l) => out.push_str(l),
                Part::At => out.push_str(text),
                _ => {}
            }
        }
        out
    }

    /// A number's magnitude drawn by a number's section onto `out`; whether
    /// it drew as zero, its digits all 0 once rounded. A section with no
    /// placeholder draws no number, and so never a zero.
    fn draw(&self, magnitude: f64, out: &mut String) -> bool {
        if !self.placeholders {
            for p in &self.parts {
                if let Part::Literal(l) = p {
                    out.push_str(l);
                }
            }
            return false;
        }
        let (digits, point) = decimal(magnitude);
        let (int, frac) = rounded(&digits, point + self.shift, self.decimals.len());
        let zero = int.is_empty() && frac.iter().all(|&d| d == 0);
        // A `#` drops a trailing 0; a `0` keeps its digit and every one
        // before it.
        let mut kept = frac.len();
        while kept > 0 && !self.decimals[kept - 1] && frac[kept - 1] == 0 {
            kept -= 1;
        }
        let mut int = int;
        if int.len() < self.min_integer {
            let mut padded = vec![0; self.min_integer - int.len()];
            padded.extend_from_slice(&int);
            int = padded;
        }
        let len = int.len();
        let push_digit = |j: usize, out: &mut String| {
            out.push(char::from(b'0' + int[j]));
            let from_right = len - 1 - j;
            if self.grouping && from_right > 0 && from_right % 3 == 0 {
                out.push(',');
            }
        };
        let mut after_point = false;
        let (mut k, mut d) = (0, 0);
        for p in &self.parts {
            match p {
                Part::Digit(_) if !after_point => {
                    // The integer digits stand right to left across the
                    // placeholders, the leftmost taking every digit past
                    // the others.
                    let r = self.integer - 1 - k;
                    if len > r {
                        let first = if k == 0 { 0 } else { len - 1 - r };
                        for j in first..=len - 1 - r {
                            push_digit(j, out);
                        }
                    }
                    k += 1;
                }
                Part::Digit(_) => {
                    if d < kept {
                        out.push(char::from(b'0' + frac[d]));
                    }
                    d += 1;
                }
                Part::Point => {
                    if self.integer == 0 {
                        // No placeholder before the point: its digits, if
                        // any, stand before it.
                        for j in 0..len {
                            push_digit(j, out);
                        }
                    }
                    out.push('.');
                    after_point = true;
                }
                Part::Percent => out.push('%'),
                Part::Literal(l) => out.push_str(l),
                Part::At => {}
            }
        }
        zero
    }
}

/// A positive finite number's significant digits, no 0 at either end, and
/// the count of its integer digits: its shortest spelling's where that has
/// fifteen significant digits or fewer, which is then its fifteen-digit
/// rounding exactly, and the fifteen of `{:.14e}` where it has more, the
/// nearest. A number exactly halfway between two of fifteen digits rounds
/// toward zero, as Excel draws 1234567890123455 as 1234567890123450 (the
/// reference's golden, 2026-10-17), where `{:.14e}` rounds a tie to even;
/// such a number's shortest spelling is the tie itself, sixteen digits
/// ending in 5. Zero has none.
fn decimal(magnitude: f64) -> (Vec<u8>, i32) {
    if magnitude == 0.0 {
        return (Vec::new(), 0);
    }
    let spelled = |text: String| -> (Vec<u8>, i32) {
        let (mantissa, exponent) = text.split_once('e').expect("an exponent");
        let digits: Vec<u8> = mantissa
            .bytes()
            .filter(u8::is_ascii_digit)
            .map(|b| b - b'0')
            .collect();
        (
            digits,
            exponent.parse::<i32>().expect("a whole exponent") + 1,
        )
    };
    let (mut digits, mut point) = spelled(format!("{magnitude:e}"));
    if digits.len() == 16 && digits[15] == 5 && exactly(magnitude, &digits, point) {
        digits.pop();
    } else if digits.len() > 15 {
        (digits, point) = spelled(format!("{magnitude:.14e}"));
    }
    while digits.last() == Some(&0) {
        digits.pop();
    }
    (digits, point)
}

/// Whether a positive number is exactly the decimal its digits and point
/// spell, read in whole numbers: the decimal `d` times 10^p is a double only
/// as a whole number, or, for p below 0, where 5^-p divides `d`, the whole
/// number d / 5^-p over 2^-p, a power of two scaling the double exactly.
fn exactly(magnitude: f64, digits: &[u8], point: i32) -> bool {
    let d = digits
        .iter()
        .fold(0u128, |acc, &x| acc * 10 + u128::from(x));
    let p = point - digits.len() as i32;
    if p >= 0 {
        if p > 20 || magnitude.fract() != 0.0 || magnitude >= 1e38 {
            return false;
        }
        magnitude as u128 == d * 10u128.pow(p as u32)
    } else {
        let q = (-p) as u32;
        if q > 40 {
            return false;
        }
        let five = 5u128.pow(q);
        if d % five != 0 {
            return false;
        }
        let scaled = magnitude * 2f64.powi(q as i32);
        scaled.fract() == 0.0 && scaled < 1e38 && scaled as u128 == d / five
    }
}

/// Digits standing for a number with `point` integer digits (the digit at
/// index j is worth 10^(point-1-j)), rounded half away from zero to
/// `decimals` places: the integer digits, none for a number under 1, and
/// the decimals, exactly `decimals` of them.
fn rounded(digits: &[u8], point: i32, decimals: usize) -> (Vec<u8>, Vec<u8>) {
    let at = |j: i32| -> u8 {
        if j < 0 {
            0
        } else {
            digits.get(j as usize).copied().unwrap_or(0)
        }
    };
    let keep = point + decimals as i32;
    let mut kept: Vec<u8> = (0..keep.max(0)).map(at).collect();
    if keep >= 0 && at(keep) >= 5 {
        let mut i = kept.len();
        loop {
            if i == 0 {
                kept.insert(0, 1);
                break;
            }
            i -= 1;
            if kept[i] == 9 {
                kept[i] = 0;
            } else {
                kept[i] += 1;
                break;
            }
        }
    }
    // The last digit kept stands at the last decimal's place, a carry past
    // the first or not; a number under the first decimal's place has its
    // zeros put before it.
    if kept.len() < decimals {
        let mut padded = vec![0; decimals - kept.len()];
        padded.extend_from_slice(&kept);
        kept = padded;
    }
    let split = kept.len() - decimals;
    let frac = kept[split..].to_vec();
    let int = match kept[..split].iter().position(|&x| x != 0) {
        Some(lead) => kept[lead..split].to_vec(),
        None => Vec::new(),
    };
    (int, frac)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn code(text: &str) -> Code {
        Code::parse(text).unwrap_or_else(|u| panic!("{text}: {u:?}"))
    }

    fn drawn(text: &str, x: f64) -> String {
        code(text).text(&Value::Number(x)).expect("a number draws")
    }

    fn refused(text: &str) -> Unsupported {
        match Code::parse(text) {
            Ok(c) => panic!("{text} was taken: {c:?}"),
            Err(u) => u,
        }
    }

    #[test]
    fn a_number_draws_with_its_decimals() {
        assert_eq!(drawn("0.00", 1.0), "1.00");
        assert_eq!(drawn("0.00", 12.3), "12.30");
        assert_eq!(drawn("0.00", 0.0), "0.00");
        assert_eq!(drawn("0.00", -0.0), "0.00");
        assert_eq!(drawn("0", 0.4), "0");
        assert_eq!(drawn("#,##0.00", 1234.5), "1,234.50");
        assert_eq!(drawn("#,##0", 1234567.0), "1,234,567");
        assert_eq!(drawn("#,##0", 999.0), "999");
        assert_eq!(drawn("0.00", 1e16), "10000000000000000.00");
    }

    #[test]
    fn rounding_is_excel_s_half_away_from_zero_over_fifteen_digits() {
        // Rust's own {:.2} spells these 2.67, 1.00 and 0.12, and {:.0}
        // spells 2.5 as 2: ties to even on the binary value.
        assert_eq!(drawn("0.00", 2.675), "2.68");
        assert_eq!(drawn("0.00", 1.005), "1.01");
        assert_eq!(drawn("0.00", 0.125), "0.13");
        assert_eq!(drawn("0", 2.5), "3");
        assert_eq!(drawn("0", 0.5), "1");
        assert_eq!(drawn("0", -2.5), "-3");
        assert_eq!(drawn("0.00", 9.995), "10.00");
        assert_eq!(drawn("0.00", 0.005), "0.01");
        assert_eq!(drawn("0.00", 0.0049), "0.00");
        assert_eq!(drawn("0.00", 0.1 + 0.2), "0.30");
        // The fifteenth digit rounded by the sixteenth, the rest zeros.
        assert_eq!(drawn("0", 123456789012345678.0), "123456789012346000");
        // An exact tie at the fifteenth digit toward zero, as Excel draws it
        // (the reference's golden), where {:.14e} rounds to even: whole, and
        // with a fraction, the tie read in whole numbers either way.
        assert_eq!(drawn("0", 1234567890123455.0), "1234567890123450");
        assert_eq!(drawn("0", 1234567890123465.0), "1234567890123460");
        assert_eq!(drawn("0", 617283945061727.5), "617283945061727");
        // Not a tie: the nearest, the sixteenth digit 5 with more after it,
        // written out, or past the double's shortest spelling, which reads
        // 0.2234567890123455 for 0.22345678901234550250...
        assert_eq!(drawn("0", 1234567890123455.5), "1234567890123460");
        assert_eq!(
            drawn("0.000000000000000", 0.2234567890123455),
            "0.223456789012346"
        );
        assert_eq!(drawn("0", 1e20), "100000000000000000000");
        assert_eq!(drawn("0.00", 5e-324), "0.00");
    }

    #[test]
    fn sections_choose_by_sign_before_rounding() {
        assert_eq!(drawn("0.00", -1.5), "-1.50");
        // One section's minus none where the number draws as zero, as Excel
        // draws it (the reference's golden); a second section's draws it.
        assert_eq!(drawn("0.00", -0.001), "0.00");
        assert_eq!(drawn("0", -0.4), "0");
        assert_eq!(drawn("#,##0.00", -0.001), "0.00");
        assert_eq!(drawn("0.00;(0.00)", -1.5), "(1.50)");
        assert_eq!(drawn("0.00;(0.00)", -0.001), "(0.00)");
        assert_eq!(drawn("0.00;(0.00)", 0.0), "0.00");
        assert_eq!(drawn("0.00;(0.00);\"zero\"", 0.0), "zero");
        assert_eq!(drawn("0.00;(0.00);\"zero\"", -0.0), "zero");
        assert_eq!(drawn("0.00;(0.00);\"zero\"", 0.001), "0.00");
        assert_eq!(drawn("▲0.00;▼0.00;=0.00", 1.5), "▲1.50");
        assert_eq!(drawn("▲0.00;▼0.00;=0.00", -1.5), "▼1.50");
        assert_eq!(drawn("▲0.00;▼0.00;=0.00", 0.0), "=0.00");
        assert_eq!(drawn("0;;0", -3.0), "");
        assert_eq!(drawn("0.00;-0.00", -2.0), "-2.00");
        assert_eq!(drawn("\"x\"", -5.0), "-x");
    }

    #[test]
    fn placeholders_draw_digits_zeros_and_nothing() {
        assert_eq!(drawn("#", 0.0), "");
        assert_eq!(drawn("#,###", 0.0), "");
        assert_eq!(drawn("#.##", 0.5), ".5");
        assert_eq!(drawn("0.#", 1.0), "1.");
        assert_eq!(drawn("0.0#", 1.5), "1.5");
        assert_eq!(drawn("0.0#", 1.25), "1.25");
        assert_eq!(drawn("0.#0", 1.5), "1.50");
        assert_eq!(drawn("000", 5.0), "005");
        assert_eq!(drawn("#0#", 5.0), "05");
        assert_eq!(drawn("000-00-0000", 123456789.0), "123-45-6789");
        assert_eq!(drawn("000-00-0000", 5.0), "000-00-0005");
        assert_eq!(drawn("0-0", 123.0), "12-3");
        assert_eq!(drawn(".00", 12.5), "12.50");
        assert_eq!(drawn(".00", 0.5), ".50");
    }

    #[test]
    fn percents_and_thousandths_move_the_point_in_decimal() {
        assert_eq!(drawn("0%", 0.5), "50%");
        assert_eq!(drawn("0.00%", 0.12345), "12.35%");
        assert_eq!(drawn("0.0%", 0.145), "14.5%");
        assert_eq!(drawn("#,##0,", 1234567.891), "1,235");
        assert_eq!(drawn("0.0,,", 12200000.0), "12.2");
        assert_eq!(drawn("#,##0,\"K\"", 2500.0), "3K");
        assert_eq!(drawn("0%%", 0.5), "5000%%");
    }

    #[test]
    fn literals_draw_as_written() {
        assert_eq!(drawn("\\$#,##0.00", 1234.5), "$1,234.50");
        assert_eq!(drawn("$#,##0.00", -5.0), "-$5.00");
        assert_eq!(drawn("0.00\" kg\"", 2.0), "2.00 kg");
        assert_eq!(drawn("(000) 000-0000", 5551234567.0), "(555) 123-4567");
        assert_eq!(drawn("#,##0 ;(#,##0)", 1234.0), "1,234 ");
        assert_eq!(drawn("#,##0 ;(#,##0)", -1234.0), "(1,234)");
        assert_eq!(drawn("\\€#,##0.00", 3.0), "€3.00");
    }

    #[test]
    fn a_text_draws_by_the_fourth_section_alone() {
        let c = code("0;-0;0;\"<\"@\">\"");
        assert_eq!(c.text(&Value::Text("abc".into())), Some("<abc>".into()));
        assert_eq!(code("0.00").text(&Value::Text("abc".into())), None);
        assert_eq!(
            code("0;-0;0;").text(&Value::Text("abc".into())),
            Some(String::new())
        );
        assert_eq!(code("0.00").text(&Value::Bool(true)), None);
        assert_eq!(code("0.00").text(&Value::Empty), None);
        assert_eq!(code("0.00").text(&Value::Number(f64::INFINITY)), None);
    }

    #[test]
    fn excel_s_built_ins_are_known_by_their_codes() {
        assert_eq!(code("0.00").builtin(), Some(2));
        assert_eq!(code("#,##0.00").builtin(), Some(4));
        assert_eq!(code("#,##0 ;(#,##0)").builtin(), Some(37));
        assert_eq!(code("0.000").builtin(), None);
        assert_eq!(code("▲0.00;▼0.00;=0.00").builtin(), None);
        assert_eq!(code("0.00"), code("0.00"));
        assert_ne!(code("0.00"), code("0.00;-0.00"));
    }

    #[test]
    fn what_this_version_does_not_draw_is_named() {
        let named = |text: &str, word: &str| {
            let u = refused(text);
            assert!(u.what.contains(word), "{text}: {u:?}");
            assert!(!u.fix.is_empty(), "{text}: {u:?}");
        };
        named("", "no code");
        named("General", "General format");
        named("general", "General format");
        named("@", "Text format");
        named("[Red]0.00", "[Red]");
        named("0.00;[Red]-0.00", "color");
        named("yyyy-mm-dd", "date");
        named("0.00E+00", "E+");
        named("# ?/?", "?");
        named("0/0", "/");
        named("_(0_)", "_");
        named("* 0", "*");
        named("General;-General", "General inside");
        named("0 kg", "k");
        named("0 1", "a digit");
        named("0|", "|");
        named("0.00.0", "second point");
        named("0;0;0;@;0", "fifth section");
        named("@;0", "@ before");
        named("0;-0;0;0", "in a text's section");
        named(",0", "before any digit");
        named("0.0,0", "among the decimals");
        named("0,.00", "just before the point");
        named("\"x\".", "beside no digit");
        named("\"open", "quote left open");
        named("0\\", "backslash");
        named(&format!("0.{}", "0".repeat(31)), "decimals");
        named(&"0".repeat(256), "characters");
        named("0\t", "control");
        assert!(Code::parse(&format!("0.{}", "0".repeat(30))).is_ok());
    }
}
