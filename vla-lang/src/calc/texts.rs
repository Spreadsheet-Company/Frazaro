//! The text functions (KERNEL.8, slice c, 2026-10-19): `CONCATENATE`,
//! `CONCAT`, `TEXTJOIN`, `LEFT`, `RIGHT`, `MID`, `LEN`, `FIND`, `SEARCH`,
//! `SUBSTITUTE`, `REPLACE`, `UPPER`, `LOWER`, `PROPER`, `TRIM`, `REPT`,
//! `EXACT`, `VALUE`, `T`, `N`, `UNICODE`, `UNICHAR`, `CHAR`, `CODE` and
//! `TEXT`. The language's own, since a game reaches the language alone
//! (Alonzo's `AD-1`) and Alonzo's `CART.13`, the word game, reads its words
//! a character at a time; Excel's library has them through it.
//!
//! Each rule is Excel's as its documentation states it, held where the
//! documentation is silent to Excel's own saved answer
//! (`scripts/recalc/texts.txt`, saved by Excel as `texts_saved.xlsx`).
//! A value is read as text as `&` reads it (`Value::text`): a number in
//! Excel's General spelling, a truth value `TRUE` or `FALSE`, an empty cell
//! `""`, an error the error; a count, a start or an instance as a number,
//! read down to the whole number at or below it ([`count`]); an argument
//! where one value is wanted met at the cell, as `Ctx::one` meets it.
//!
//! Excel counts a text by its characters, an emoji one, as its save of the
//! fixture holds (`LEN` of an emoji and two letters 3, `LEFT` of it by 2
//! the emoji and a letter): every count and position here is a character's
//! of the language, so that no cut halves one. Excel's limit on a cell's
//! text, 32,767, is in UTF-16 units, an emoji two (`REPT` of 16,384 emoji
//! `#VALUE!`): a text a function makes past it is `#VALUE!` ([`made`]), and
//! one `&` or `CONCATENATE` makes is cut to it, as Excel's save holds them
//! ([`cut_to_most`]), and labelled where the cut would halve a surrogate
//! pair, where Excel keeps half a character. `CHAR` and `CODE` read the
//! host's code page in Excel, Windows' 1252 under the owner's, which the
//! language has none of: they compute over ASCII, the same in every code
//! page, and label the rest (the owner's decision, 2026-10-19). `FIXED` and
//! `DOLLAR`, a locale's separators and its currency, are not here.

use super::criteria::wildcard_prefix;
use super::eval::{Arg, Ctx, Reason};
use super::value::{number_to_text, text_to_number, ErrorKind, Value};
use crate::numfmt::Code;

/// The most UTF-16 units a text a formula makes may hold, Excel's limit on
/// a cell's text, an emoji two.
pub const MOST: usize = 32_767;

/// A text's size in UTF-16 units, as Excel's limit on a cell's text
/// measures it.
pub fn units(t: &str) -> usize {
    if t.is_ascii() {
        t.len()
    } else {
        t.encode_utf16().count()
    }
}

/// A text's length in characters, as Excel counts it, an emoji one.
pub fn length(t: &str) -> usize {
    if t.is_ascii() {
        t.len()
    } else {
        t.chars().count()
    }
}

/// A text `SUBSTITUTE`, `REPLACE` or `TEXT` makes: itself, or `#VALUE!` past
/// [`MOST`] units, as Excel's save answers `SUBSTITUTE` (`REPT`, `CONCAT`
/// and `TEXTJOIN` hold the limit themselves). A text's UTF-8 bytes are never
/// fewer than its UTF-16 units, so a text of [`MOST`] bytes or fewer is
/// never counted.
pub(crate) fn made(t: String) -> Value {
    if t.len() > MOST && units(&t) > MOST {
        Value::Error(ErrorKind::Value)
    } else {
        Value::Text(t)
    }
}

/// A text `&` or `CONCATENATE` makes, cut to [`MOST`] units, as Excel's
/// save holds `LEN(REPT("x",32767)&"x")` and its `CONCATENATE` 32,767
/// (`Functions!A82`, `A85`), and whether the cut fell between characters:
/// where it would halve a surrogate pair, the pair's character is cut too,
/// one unit short of the half Excel keeps.
pub(crate) fn cut_to_most(mut t: String) -> (String, bool) {
    if t.len() <= MOST || units(&t) <= MOST {
        return (t, true);
    }
    let mut u = 0;
    for (b, c) in t.char_indices() {
        let next = u + c.len_utf16();
        if next > MOST {
            t.truncate(b);
            return (t, u == MOST);
        }
        u = next;
    }
    (t, true)
}

/// The label of a text `&` or `CONCATENATE` would cut through a surrogate
/// pair at its 32,767th unit: half a character, which no text of the
/// language holds.
pub(crate) fn halved() -> Reason {
    Reason::Construct("half a surrogate pair".to_string())
}

/// What a text function answers other than its value: an error, which is
/// the cell's value, or a reason, which labels the cell.
enum Fault {
    Error(ErrorKind),
    Reason(Reason),
}

impl From<ErrorKind> for Fault {
    fn from(k: ErrorKind) -> Fault {
        Fault::Error(k)
    }
}

impl From<Reason> for Fault {
    fn from(r: Reason) -> Fault {
        Fault::Reason(r)
    }
}

/// A function's answer from its rule: a fault's error the value, its
/// reason the label.
fn answer(rule: impl FnOnce() -> Result<Value, Fault>) -> Result<Value, Reason> {
    match rule() {
        Ok(v) => Ok(v),
        Err(Fault::Error(k)) => Ok(Value::Error(k)),
        Err(Fault::Reason(r)) => Err(r),
    }
}

/// The answer to a call with the wrong count of arguments, or one out of
/// its range.
fn wrong<T>() -> Result<T, Fault> {
    Err(Fault::Error(ErrorKind::Value))
}

/// A label: a construct this version does not answer, as Excel would.
fn label<T>(what: &str) -> Result<T, Fault> {
    Err(Fault::Reason(Reason::Construct(what.to_string())))
}

/// An argument's text, as `&` reads it.
fn text(ctx: &Ctx<'_>, arg: &Arg<'_>) -> Result<String, Fault> {
    Ok(ctx.one(arg)?.text()?)
}

/// A count, a start or an instance, read as a number down to the whole
/// number at or below it, as Excel's save answers `#VALUE!` for
/// `LEFT(text,-0.5)` (`Functions!A13`), where truncation toward zero would
/// make it 0; `default` where the argument is left out.
fn count(ctx: &Ctx<'_>, arg: Option<&Arg<'_>>, default: f64) -> Result<f64, Fault> {
    match arg {
        None => Ok(default),
        Some(a) => Ok(ctx.one(a)?.number()?.floor()),
    }
}

/// A count not below 0 held to a text's length.
fn at_most(n: f64, len: usize) -> usize {
    if n >= len as f64 {
        len
    } else {
        n as usize
    }
}

/// The byte where character `k` of `t` begins, `t`'s length where `k` is
/// its length or past it.
fn byte_at(t: &str, k: usize) -> usize {
    if t.is_ascii() {
        return k.min(t.len());
    }
    t.char_indices().nth(k).map_or(t.len(), |(b, _)| b)
}

/// `t`'s characters from `from` to `to`.
fn cut(t: &str, from: usize, to: usize) -> &str {
    let a = byte_at(t, from);
    &t[a..byte_at(t, to).max(a)]
}

/// The character `t[..b]` ends at, plus one: a position as Excel counts.
fn position(t: &str, b: usize) -> Value {
    Value::Number((length(&t[..b]) + 1) as f64)
}

/// `LEN(text)`: its length in characters, an emoji one.
pub(crate) fn len(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    answer(|| {
        let [t] = args else { return wrong() };
        Ok(Value::Number(length(&text(ctx, t)?) as f64))
    })
}

/// `LEFT(text, [count])`: its first `count` characters, 1 when left out;
/// the whole text for a count past its length, `#VALUE!` for one below 0.
pub(crate) fn left(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    answer(|| {
        if !(1..=2).contains(&args.len()) {
            return wrong();
        }
        let t = text(ctx, &args[0])?;
        let n = count(ctx, args.get(1), 1.0)?;
        if n < 0.0 {
            return wrong();
        }
        let n = at_most(n, length(&t));
        Ok(Value::Text(cut(&t, 0, n).to_string()))
    })
}

/// `RIGHT(text, [count])`: its last `count` characters, as `LEFT` counts.
pub(crate) fn right(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    answer(|| {
        if !(1..=2).contains(&args.len()) {
            return wrong();
        }
        let t = text(ctx, &args[0])?;
        let n = count(ctx, args.get(1), 1.0)?;
        if n < 0.0 {
            return wrong();
        }
        let len = length(&t);
        let n = at_most(n, len);
        Ok(Value::Text(cut(&t, len - n, len).to_string()))
    })
}

/// `MID(text, start, count)`: `count` characters from the `start`th, `""`
/// from a start past the end; `#VALUE!` for a start below 1 or a count
/// below 0.
pub(crate) fn mid(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    answer(|| {
        let [t, start, n] = args else { return wrong() };
        let t = text(ctx, t)?;
        let start = count(ctx, Some(start), 1.0)?;
        let n = count(ctx, Some(n), 0.0)?;
        if start < 1.0 || n < 0.0 {
            return wrong();
        }
        let len = length(&t);
        if start > len as f64 {
            return Ok(Value::Text(String::new()));
        }
        let from = start as usize - 1;
        let to = from + at_most(n, len - from);
        Ok(Value::Text(cut(&t, from, to).to_string()))
    })
}

/// Where `FIND` and `SEARCH` begin: the start read, 1 when left out, or
/// `#VALUE!` for one below 1 or past the text, where an empty text sought
/// is found at the start itself, up to one past the end.
fn start_of(
    ctx: &Ctx<'_>,
    args: &[Arg],
    sought: &str,
    within: &str,
) -> Result<Result<usize, Value>, Fault> {
    let start = count(ctx, args.get(2), 1.0)?;
    let len = length(within) as f64;
    if start < 1.0 || start > len + 1.0 || (start > len && !sought.is_empty()) {
        return wrong();
    }
    if sought.is_empty() {
        return Ok(Err(Value::Number(start)));
    }
    Ok(Ok(byte_at(within, start as usize - 1)))
}

/// `FIND(sought, within, [start])`: where the text sought first stands in
/// the other from the start on, exactly, case and all; `#VALUE!` where it
/// stands nowhere.
pub(crate) fn find(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    answer(|| {
        if !(2..=3).contains(&args.len()) {
            return wrong();
        }
        let sought = text(ctx, &args[0])?;
        let within = text(ctx, &args[1])?;
        let from = match start_of(ctx, args, &sought, &within)? {
            Ok(b) => b,
            Err(found) => return Ok(found),
        };
        match within[from..].find(sought.as_str()) {
            Some(b) => Ok(position(&within, from + b)),
            None => wrong(),
        }
    })
}

/// `SEARCH(sought, within, [start])`: as `FIND`, without case and with
/// Excel's wildcards, `*`, `?` and `~` (`criteria::wildcard_prefix`), a `?`
/// one character, an emoji among them.
pub(crate) fn search(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    answer(|| {
        if !(2..=3).contains(&args.len()) {
            return wrong();
        }
        let sought = text(ctx, &args[0])?;
        let within = text(ctx, &args[1])?;
        let from = match start_of(ctx, args, &sought, &within)? {
            Ok(b) => b,
            Err(found) => return Ok(found),
        };
        let starts = within[from..]
            .char_indices()
            .map(|(b, _)| from + b)
            .chain(std::iter::once(within.len()));
        for b in starts {
            if wildcard_prefix(&sought, &within[b..]) {
                return Ok(position(&within, b));
            }
        }
        wrong()
    })
}

/// `SUBSTITUTE(text, old, new, [instance])`: every `old` in the text, from
/// the left and not overlapping, made `new`, or the `instance`th alone;
/// the text itself for an empty `old`; `#VALUE!` for an instance below 1.
pub(crate) fn substitute(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    answer(|| {
        if !(3..=4).contains(&args.len()) {
            return wrong();
        }
        let t = text(ctx, &args[0])?;
        let old = text(ctx, &args[1])?;
        let new = text(ctx, &args[2])?;
        let instance = match args.get(3) {
            None => None,
            Some(a) => {
                let i = count(ctx, Some(a), 1.0)?;
                if i < 1.0 {
                    return wrong();
                }
                Some(i)
            }
        };
        if old.is_empty() {
            return Ok(Value::Text(t));
        }
        Ok(match instance {
            None => made(t.replace(old.as_str(), &new)),
            Some(i) => {
                let nth = (i as usize).saturating_sub(1);
                match t.match_indices(old.as_str()).nth(nth) {
                    Some((b, _)) => made(format!("{}{new}{}", &t[..b], &t[b + old.len()..])),
                    None => Value::Text(t),
                }
            }
        })
    })
}

/// `REPLACE(text, start, count, new)`: the `count` characters from the
/// `start`th made `new`, a start past the end adding `new` at the end;
/// `#VALUE!` for a start below 1 or a count below 0.
pub(crate) fn replace(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    answer(|| {
        let [t, start, n, new] = args else {
            return wrong();
        };
        let t = text(ctx, t)?;
        let start = count(ctx, Some(start), 1.0)?;
        let n = count(ctx, Some(n), 0.0)?;
        let new = text(ctx, new)?;
        if start < 1.0 || n < 0.0 {
            return wrong();
        }
        let len = length(&t);
        let from = at_most(start - 1.0, len);
        let to = from + at_most(n, len - from);
        let (a, b) = (byte_at(&t, from), byte_at(&t, to));
        Ok(made(format!("{}{new}{}", &t[..a], &t[b..])))
    })
}

/// A character's upper case by Unicode's simple mapping, one character for
/// one, as Excel's `UPPER` and `LOWER` draw them: the one character Unicode
/// gives where it gives one; where it gives more, the character itself
/// where the simple mapping has none (`ß`, as Excel's save keeps it), and a
/// Greek letter with a subscript iota its title-case letter, the simple
/// mapping Unicode's data names.
fn upper_char(c: char) -> char {
    if c.is_ascii() {
        return c.to_ascii_uppercase();
    }
    let mut u = c.to_uppercase();
    match (u.next(), u.next()) {
        (Some(x), None) => x,
        _ => match u32::from(c) {
            0x1F80..=0x1F87 | 0x1F90..=0x1F97 | 0x1FA0..=0x1FA7 => {
                char::from_u32(u32::from(c) + 8).unwrap_or(c)
            }
            0x1FB3 | 0x1FC3 | 0x1FF3 => char::from_u32(u32::from(c) + 9).unwrap_or(c),
            _ => c,
        },
    }
}

/// [`upper_char`]'s lower case: the character `char::to_lowercase` begins
/// with, the simple mapping, so that `İ` is `i`, as Excel's save holds
/// `LOWER("İ")` (`Functions!A68`).
fn lower_char(c: char) -> char {
    if c.is_ascii() {
        return c.to_ascii_lowercase();
    }
    c.to_lowercase().next().unwrap_or(c)
}

/// `UPPER(text)`: every letter in upper case.
pub(crate) fn upper(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    answer(|| {
        let [t] = args else { return wrong() };
        Ok(Value::Text(text(ctx, t)?.chars().map(upper_char).collect()))
    })
}

/// `LOWER(text)`: every letter in lower case.
pub(crate) fn lower(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    answer(|| {
        let [t] = args else { return wrong() };
        Ok(Value::Text(text(ctx, t)?.chars().map(lower_char).collect()))
    })
}

/// `PROPER(text)`: a letter after anything but a letter in upper case,
/// every other in lower, so that `o'neil` is `O'Neil` and `2nd` is `2Nd`.
pub(crate) fn proper(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    answer(|| {
        let [t] = args else { return wrong() };
        let mut after_letter = false;
        let out = text(ctx, t)?
            .chars()
            .map(|c| {
                let letter = c.is_alphabetic();
                let drawn = match (letter, after_letter) {
                    (false, _) => c,
                    (true, false) => upper_char(c),
                    (true, true) => lower_char(c),
                };
                after_letter = letter;
                drawn
            })
            .collect();
        Ok(Value::Text(out))
    })
}

/// `TRIM(text)`: its spaces, U+0020 alone, gone from both ends and each run
/// between words made one.
pub(crate) fn trim(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    answer(|| {
        let [t] = args else { return wrong() };
        let t = text(ctx, t)?;
        let words: Vec<&str> = t.split(' ').filter(|w| !w.is_empty()).collect();
        Ok(Value::Text(words.join(" ")))
    })
}

/// `REPT(text, times)`: the text so many times over; `#VALUE!` for a count
/// below 0 or a text past [`MOST`].
pub(crate) fn rept(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    answer(|| {
        let [t, times] = args else { return wrong() };
        let t = text(ctx, t)?;
        let times = count(ctx, Some(times), 0.0)?;
        if times < 0.0 {
            return wrong();
        }
        let each = units(&t);
        if each == 0 || times < 1.0 {
            return Ok(Value::Text(String::new()));
        }
        if each as f64 * times > MOST as f64 {
            return wrong();
        }
        Ok(Value::Text(t.repeat(times as usize)))
    })
}

/// `EXACT(a, b)`: whether the two texts are one, case and all.
pub(crate) fn exact(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    answer(|| {
        let [a, b] = args else { return wrong() };
        let a = text(ctx, a)?;
        let b = text(ctx, b)?;
        Ok(Value::Bool(a == b))
    })
}

/// `VALUE(text)`: the number a text spells, as an operator reads one; a
/// number itself, an empty value 0, a truth value `#VALUE!`. A text no
/// operator reads is `#VALUE!` when it holds no digit, and labelled when it
/// does: a thousands separator, a currency or a date is read by the host's
/// locale in Excel (`VALUE("1,000")` is 1000 in an English one), which the
/// language has none of (the line's Found of 2026-10-18: dates decide it).
pub(crate) fn value(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    answer(|| {
        let [v] = args else { return wrong() };
        match ctx.one(v)? {
            Value::Number(n) => Ok(Value::Number(n)),
            Value::Empty => Ok(Value::Number(0.0)),
            Value::Bool(_) => wrong(),
            Value::Error(k) => Ok(Value::Error(k)),
            Value::Text(t) => match text_to_number(&t) {
                Some(n) => Ok(Value::Number(n)),
                None if t.bytes().any(|b| b.is_ascii_digit()) => label("a locale's number"),
                None => wrong(),
            },
        }
    })
}

/// `T(value)`: a text itself, an error the error, anything else `""`.
pub(crate) fn t(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    answer(|| {
        let [v] = args else { return wrong() };
        Ok(match ctx.one(v)? {
            Value::Text(t) => Value::Text(t),
            Value::Error(k) => Value::Error(k),
            _ => Value::Text(String::new()),
        })
    })
}

/// `N(value)`: a number itself, a truth value 1 or 0, an error the error,
/// anything else 0.
pub(crate) fn n(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    answer(|| {
        let [v] = args else { return wrong() };
        Ok(match ctx.one(v)? {
            Value::Number(n) => Value::Number(n),
            Value::Bool(b) => Value::Number(if b { 1.0 } else { 0.0 }),
            Value::Error(k) => Value::Error(k),
            _ => Value::Number(0.0),
        })
    })
}

/// `UNICODE(text)`: its first character's code point, a pair's whole;
/// `#VALUE!` for an empty text.
pub(crate) fn unicode(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    answer(|| {
        let [t] = args else { return wrong() };
        match text(ctx, t)?.chars().next() {
            Some(c) => Ok(Value::Number(f64::from(u32::from(c)))),
            None => wrong(),
        }
    })
}

/// `UNICHAR(number)`: the character of a code point, 1 to 1,114,111;
/// `#N/A` for a surrogate's, which is half a character; `#VALUE!` outside.
pub(crate) fn unichar(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    answer(|| {
        let [n] = args else { return wrong() };
        let n = count(ctx, Some(n), 0.0)?;
        if !(1.0..=1_114_111.0).contains(&n) {
            return wrong();
        }
        match char::from_u32(n as u32) {
            Some(c) => Ok(Value::Text(c.to_string())),
            None => Ok(Value::Error(ErrorKind::NA)),
        }
    })
}

/// `CHAR(number)`: the character of a code 1 to 255 in the host's code
/// page; computed for ASCII's 1 to 127, the same in every page, and
/// labelled for 128 to 255; `#VALUE!` outside.
pub(crate) fn char_(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    answer(|| {
        let [n] = args else { return wrong() };
        let n = count(ctx, Some(n), 0.0)?;
        if !(1.0..=255.0).contains(&n) {
            return wrong();
        }
        if n >= 128.0 {
            return label("a code page's character");
        }
        Ok(Value::Text(char::from(n as u8).to_string()))
    })
}

/// `CODE(text)`: its first character's code in the host's code page;
/// computed for ASCII, labelled for any other; `#VALUE!` for an empty text.
pub(crate) fn code(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    answer(|| {
        let [t] = args else { return wrong() };
        match text(ctx, t)?.chars().next() {
            Some(c) if c.is_ascii() => Ok(Value::Number(f64::from(c as u8))),
            Some(_) => label("a code page's code"),
            None => wrong(),
        }
    })
}

/// A text growing toward [`MOST`]: past it `#VALUE!`, as Excel documents
/// `CONCAT` and `TEXTJOIN`, or cut to it, as its save holds `CONCATENATE`.
struct Joined {
    out: String,
    over: bool,
}

impl Joined {
    fn new() -> Joined {
        Joined {
            out: String::new(),
            over: false,
        }
    }

    fn push(&mut self, t: &str) {
        if self.over {
            return;
        }
        self.out.push_str(t);
        self.over = self.out.len() > MOST && units(&self.out) > MOST;
    }

    fn value(self) -> Value {
        if self.over {
            Value::Error(ErrorKind::Value)
        } else {
            Value::Text(self.out)
        }
    }

    /// The text cut to [`MOST`] units ([`cut_to_most`]), its first units
    /// being all that was kept once it passed them; the label of a cut
    /// through a surrogate pair.
    fn cut(self) -> Result<Value, Fault> {
        match cut_to_most(self.out) {
            (t, true) => Ok(Value::Text(t)),
            (_, false) => Err(Fault::Reason(halved())),
        }
    }
}

/// `CONCATENATE(text, ...)`: the texts one after the other, each argument
/// one value, cut to [`MOST`] units as `&` cuts them.
pub(crate) fn concatenate(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    answer(|| {
        if args.is_empty() || args.len() > 255 {
            return wrong();
        }
        let mut joined = Joined::new();
        for a in args {
            joined.push(&text(ctx, a)?);
        }
        joined.cut()
    })
}

/// Every value an argument holds, in order, to `f`: a value given
/// directly, or each cell of a reference holding one, row by row, an empty
/// cell passed by; the first error the answer.
fn each_value(
    ctx: &Ctx<'_>,
    arg: &Arg<'_>,
    mut f: impl FnMut(Value),
) -> Result<Option<ErrorKind>, Fault> {
    match arg {
        Arg::Scalar(Value::Error(k)) => Ok(Some(*k)),
        Arg::Scalar(v) => {
            f(v.clone());
            Ok(None)
        }
        Arg::Area { areas, .. } => {
            let mut error = None;
            ctx.each(areas, |v| {
                match v {
                    Value::Error(k) => {
                        error.get_or_insert(k);
                    }
                    v if error.is_none() => f(v),
                    _ => {}
                }
                Ok(())
            })?;
            Ok(error)
        }
    }
}

/// `CONCAT(text, ...)`: every value its arguments hold, a reference's cells
/// row by row, one after the other.
pub(crate) fn concat(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    answer(|| {
        if args.is_empty() || args.len() > 253 {
            return wrong();
        }
        let mut joined = Joined::new();
        for a in args {
            let error = each_value(ctx, a, |v| {
                if let Ok(t) = v.text() {
                    joined.push(&t);
                }
            })?;
            if let Some(k) = error {
                return Ok(Value::Error(k));
            }
        }
        Ok(joined.value())
    })
}

/// The most cells a range of delimiters is read from.
const MOST_DELIMITERS: usize = 1024;

/// `TEXTJOIN(delimiter, ignore_empty, text, ...)`: every value its texts
/// hold, a reference's cells row by row, with the delimiter between them;
/// a range of delimiters taken in turn; an empty value and an empty text
/// passed by where `ignore_empty` is TRUE, and joined as `""` where it is
/// FALSE, an empty cell of a reference among them, so that its delimiters
/// stand.
pub(crate) fn textjoin(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    answer(|| {
        if args.len() < 3 || args.len() > 252 {
            return wrong();
        }
        let delimiters = delimiters(ctx, &args[0])?;
        let ignore = ctx.one(&args[1])?.truth()?;
        // An empty cell joins its delimiters where it is not passed by, so a
        // reference is read at every place; where every delimiter is empty
        // it adds nothing, and its cells alone are read.
        let every_place = !ignore && delimiters.iter().any(|d| !d.is_empty());
        let mut joined = Joined::new();
        let mut first = true;
        let mut k = 0usize;
        let mut push = |t: &str, joined: &mut Joined| {
            if ignore && t.is_empty() {
                return;
            }
            if !first && !delimiters.is_empty() {
                joined.push(&delimiters[k % delimiters.len()]);
                k += 1;
            }
            first = false;
            joined.push(t);
        };
        for a in &args[2..] {
            match a {
                Arg::Area { areas, .. } if every_place => {
                    for area in areas.iter() {
                        for row in area.top..=area.bottom {
                            for col in area.left..=area.right {
                                match ctx.env.read((area.sheet, row, col))? {
                                    Value::Error(e) => return Ok(Value::Error(e)),
                                    v => push(&v.text()?, &mut joined),
                                }
                                if joined.over {
                                    return wrong();
                                }
                            }
                        }
                    }
                }
                a => {
                    let mut texts = Vec::new();
                    if let Some(e) = each_value(ctx, a, |v| texts.push(v.text()))? {
                        return Ok(Value::Error(e));
                    }
                    for t in texts {
                        push(&t?, &mut joined);
                    }
                }
            }
        }
        Ok(joined.value())
    })
}

/// `TEXTJOIN`'s delimiters: a value's text, or each cell's of a range,
/// an empty cell `""`, up to [`MOST_DELIMITERS`] of them, a larger range
/// labelled.
fn delimiters(ctx: &Ctx<'_>, arg: &Arg<'_>) -> Result<Vec<String>, Fault> {
    let Arg::Area { areas, .. } = arg else {
        return Ok(vec![ctx.one(arg)?.text()?]);
    };
    let mut out = Vec::new();
    for area in areas.iter() {
        let cells =
            (u64::from(area.bottom - area.top) + 1) * (u64::from(area.right - area.left) + 1);
        if out.len() as u64 + cells > MOST_DELIMITERS as u64 {
            return label("a range of more than 1024 delimiters");
        }
        for row in area.top..=area.bottom {
            for col in area.left..=area.right {
                out.push(ctx.env.read((area.sheet, row, col))?.text()?);
            }
        }
    }
    Ok(out)
}

/// `TEXT(value, code)`: a number drawn by a number format's code
/// (`crate::numfmt`), a date's and a time's among them, as `Range.Text`
/// draws it; a text spelling a number, and an empty value, read as the
/// number; any other text drawn by the code's text's section, or itself
/// where it has none; a truth value its text. `General` draws Excel's
/// General spelling, `@` a value's text and an empty code `""`. A code this
/// version does not draw is labelled; a number its code draws no date for,
/// below 0 or past 9999, is `#VALUE!`.
pub(crate) fn text_(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    answer(|| {
        let [v, code] = args else { return wrong() };
        let v = ctx.one(v)?;
        let code = text(ctx, code)?;
        let number = match &v {
            Value::Error(k) => return Ok(Value::Error(*k)),
            Value::Number(n) => Some(*n),
            Value::Empty => Some(0.0),
            Value::Text(t) => text_to_number(t),
            Value::Bool(_) => None,
        };
        if code.is_empty() {
            return Ok(Value::Text(String::new()));
        }
        if code == "@" {
            return Ok(made(v.text()?));
        }
        if code.eq_ignore_ascii_case("general") {
            return Ok(match number {
                Some(n) => Value::Text(number_to_text(n)),
                None => Value::Text(v.text()?),
            });
        }
        let Ok(parsed) = Code::parse(&code) else {
            return label("TEXT's code");
        };
        Ok(match number {
            Some(n) => match parsed.text(&Value::Number(n)) {
                Some(drawn) => made(drawn),
                None => Value::Error(ErrorKind::Value),
            },
            None => {
                let t = v.text()?;
                match parsed.text(&Value::Text(t.clone())) {
                    Some(drawn) => made(drawn),
                    None => Value::Text(t),
                }
            }
        })
    })
}

#[cfg(test)]
mod tests {
    use crate::calc::{library, Calc, ErrorKind};
    use crate::sheet::{parse_a1_range, Cell, Content, Workbook};

    /// The texts the functions read, on a sheet `T`: words, spaces, an
    /// apostrophe and digits among letters, letters beyond ASCII, an emoji
    /// (a surrogate pair, two UTF-16 units) first and between letters, a
    /// number, a truth value, an error; a row of texts with an empty cell
    /// among them in B, delimiters in D, and texts a locale might read in C.
    /// A text spelling a number is written after a `'`.
    const TABLE: &[(&str, &str)] = &[
        ("A1", "Hello World"),
        ("A2", "  lots   of   space  "),
        ("A3", "o'neil 2nd-place"),
        ("A4", "Éclair ß"),
        ("A5", "😀ab"),
        ("A6", "a😀b"),
        ("A7", "123.5"),
        ("A8", "TRUE"),
        ("A9", "#N/A"),
        ("B1", "x"),
        ("B3", "z"),
        ("C1", "'1,000"),
        ("C2", "12abc"),
        ("C3", "'50%"),
        ("C4", "' 7 "),
        ("C5", "abc"),
        ("D1", "-"),
        ("D2", "+"),
    ];

    /// Each formula's value over [`TABLE`] as `Calc` computes it, spelled,
    /// the formulas down column A of a sheet `F`.
    fn computed(formulas: &[&str]) -> Vec<String> {
        let mut wb = Workbook::new();
        let t = wb.ensure_sheet("T");
        for (addr, text) in TABLE {
            let r = parse_a1_range(addr).unwrap();
            let content = if let Some(s) = text.strip_prefix('\'') {
                Content::Text(s.to_string())
            } else if let Ok(n) = text.parse::<f64>() {
                Content::Number(n)
            } else if *text == "TRUE" || *text == "FALSE" {
                Content::Bool(*text == "TRUE")
            } else if let Some(e) = ErrorKind::parse(text) {
                Content::Error(e.text().to_string())
            } else {
                Content::Text(text.to_string())
            };
            wb.sheets[t].set(r.top, r.left, Cell { content, style: 0 });
        }
        let f = wb.ensure_sheet("F");
        for (i, formula) in formulas.iter().enumerate() {
            let content = Content::Formula(formula.to_string());
            wb.sheets[f].set(i as u32 + 1, 1, Cell { content, style: 0 });
        }
        let lib = library::language();
        let calc = Calc::run(&wb, &lib);
        (0..formulas.len())
            .map(|i| match calc.computed((f, i as u32 + 1, 1)) {
                Some(c) => c.spell(),
                None => "nothing".to_string(),
            })
            .collect()
    }

    /// Each formula of the pairs against the value it is to compute.
    fn holds(cases: &[(&str, &str)]) {
        let formulas: Vec<&str> = cases.iter().map(|(f, _)| *f).collect();
        for ((formula, want), got) in cases.iter().zip(computed(&formulas)) {
            assert_eq!(got, *want, "={formula}");
        }
    }

    const NA: &str = "(error \"#N/A\")";
    const VALUE: &str = "(error \"#VALUE!\")";
    const HALF: &str = "(not-computed \"half a surrogate pair\")";

    #[test]
    fn a_text_is_counted_and_cut_by_its_characters() {
        holds(&[
            ("LEN(T!A1)", "11"),
            ("LEN(T!A5)", "3"),
            ("LEN(123.5)", "5"),
            ("LEN(TRUE)", "4"),
            ("LEN(\"\")", "0"),
            ("LEN(T!A9)", NA),
            ("LEFT(T!A1)", "\"H\""),
            ("LEFT(T!A1,5)", "\"Hello\""),
            ("LEFT(T!A1,99)", "\"Hello World\""),
            ("LEFT(T!A1,0)", "\"\""),
            ("LEFT(T!A1,2.9)", "\"He\""),
            ("LEFT(T!A1,-1)", VALUE),
            ("LEFT(T!A1,-0.5)", VALUE),
            ("LEFT(1234,2)", "\"12\""),
            ("LEFT(T!A5,2)", "\"😀a\""),
            ("LEFT(T!A5,1)", "\"😀\""),
            ("RIGHT(T!A1,5)", "\"World\""),
            ("RIGHT(T!A6,3)", "\"a😀b\""),
            ("RIGHT(T!A6,2)", "\"😀b\""),
            ("MID(T!A1,7,5)", "\"World\""),
            ("MID(T!A1,7,99)", "\"World\""),
            ("MID(T!A1,20,5)", "\"\""),
            ("MID(T!A1,0,5)", VALUE),
            ("MID(T!A1,1,-1)", VALUE),
            ("MID(T!A1,1,-0.5)", VALUE),
            ("MID(T!A6,2,2)", "\"😀b\""),
            ("MID(T!A6,3,1)", "\"b\""),
        ]);
    }

    #[test]
    fn find_and_search_count_where_the_text_stands() {
        holds(&[
            ("FIND(\"o\",T!A1)", "5"),
            ("FIND(\"o\",T!A1,6)", "8"),
            ("FIND(\"O\",T!A1)", VALUE),
            ("FIND(\"x\",T!A1)", VALUE),
            ("FIND(\"o\",T!A1,0)", VALUE),
            ("FIND(\"o\",T!A1,12)", VALUE),
            ("FIND(\"\",T!A1)", "1"),
            ("FIND(\"\",T!A1,3)", "3"),
            ("FIND(\"b\",T!A6)", "3"),
            ("FIND(\"b\",T!A6,3)", "3"),
            ("SEARCH(\"O\",T!A1)", "5"),
            ("SEARCH(\"w?r\",T!A1)", "7"),
            ("SEARCH(\"l*d\",T!A1)", "3"),
            ("SEARCH(\"~*\",\"a*b\")", "2"),
            ("SEARCH(\"é\",T!A4)", "1"),
            ("SEARCH(\"?\",T!A6)", "1"),
            ("SEARCH(\"?b\",T!A6)", "2"),
            ("SEARCH(\"b\",T!A6)", "3"),
        ]);
    }

    #[test]
    fn substitute_and_replace_make_a_text_of_another() {
        holds(&[
            ("SUBSTITUTE(\"aaa\",\"aa\",\"b\")", "\"ba\""),
            ("SUBSTITUTE(\"a-b-c\",\"-\",\"+\")", "\"a+b+c\""),
            ("SUBSTITUTE(\"a-b-c\",\"-\",\"+\",2)", "\"a-b+c\""),
            ("SUBSTITUTE(\"a-b-c\",\"-\",\"+\",3)", "\"a-b-c\""),
            ("SUBSTITUTE(\"abc\",\"\",\"x\")", "\"abc\""),
            ("SUBSTITUTE(\"abc\",\"B\",\"x\")", "\"abc\""),
            ("SUBSTITUTE(\"abc\",\"b\",\"x\",0)", VALUE),
            ("REPLACE(\"abcdef\",2,3,\"X\")", "\"aXef\""),
            ("REPLACE(\"abc\",10,1,\"X\")", "\"abcX\""),
            ("REPLACE(\"abc\",0,1,\"X\")", VALUE),
            ("REPLACE(\"abc\",1,-1,\"X\")", VALUE),
            ("REPLACE(T!A5,2,1,\"X\")", "\"😀Xb\""),
            ("REPLACE(T!A5,3,1,\"X\")", "\"😀aX\""),
        ]);
    }

    #[test]
    fn a_letter_s_case_and_a_text_s_spaces() {
        holds(&[
            ("UPPER(T!A4)", "\"ÉCLAIR ß\""),
            ("LOWER(\"ÉCLAIR\")", "\"éclair\""),
            ("LOWER(\"İ\")", "\"i\""),
            ("UPPER(\"\u{1FB3}\")", "\"\u{1FBC}\""),
            ("UPPER(\"\u{FB01}\")", "\"\u{FB01}\""),
            ("PROPER(T!A3)", "\"O'Neil 2Nd-Place\""),
            ("PROPER(\"hELLO wORLD\")", "\"Hello World\""),
            ("TRIM(T!A2)", "\"lots of space\""),
            // A tab is no space to TRIM.
            ("FIND(CHAR(9),TRIM(\" a\"&CHAR(9)&\" b \"))", "2"),
            ("EXACT(\"a\",\"A\")", "false"),
            ("EXACT(\"a\",\"a\")", "true"),
            ("EXACT(1,\"1\")", "true"),
        ]);
    }

    #[test]
    fn no_text_a_formula_makes_passes_32_767_units() {
        holds(&[
            ("REPT(\"ab\",3)", "\"ababab\""),
            ("REPT(\"ab\",0)", "\"\""),
            ("REPT(\"ab\",-1)", VALUE),
            ("REPT(\"ab\",20000)", VALUE),
            ("LEN(REPT(\"x\",32767))", "32767"),
            ("LEN(REPT(\"x\",32767)&\"x\")", "32767"),
            ("RIGHT(REPT(\"x\",32766)&\"yz\")", "\"y\""),
            ("LEN(REPT(\"😀\",16383)&\"x\")", "16384"),
            ("LEN(REPT(\"😀\",16383)&\"😀\")", HALF),
            ("LEN(REPT(\"😀\",16384))", VALUE),
            ("LEN(CONCATENATE(REPT(\"x\",32767),\"x\"))", "32767"),
            ("LEN(CONCATENATE(REPT(\"😀\",16383),\"😀\"))", HALF),
            ("LEN(CONCAT(REPT(\"x\",32767),\"x\"))", VALUE),
            ("LEN(SUBSTITUTE(REPT(\"x\",20000),\"x\",\"yy\"))", VALUE),
        ]);
    }

    #[test]
    fn a_value_is_read_as_a_number_or_a_text() {
        holds(&[
            ("VALUE(T!C3)", "0.5"),
            ("VALUE(T!C4)", "7"),
            ("VALUE(5)", "5"),
            ("VALUE(T!C5)", VALUE),
            ("VALUE(TRUE)", VALUE),
            ("VALUE(T!C1)", "(not-computed \"a locale's number\")"),
            ("VALUE(T!C2)", "(not-computed \"a locale's number\")"),
            ("T(T!A1)", "\"Hello World\""),
            ("T(5)", "\"\""),
            ("T(T!A9)", NA),
            ("N(5)", "5"),
            ("N(TRUE)", "1"),
            ("N(\"x\")", "0"),
            ("N(T!A9)", NA),
        ]);
    }

    #[test]
    fn a_character_and_its_code() {
        holds(&[
            ("UNICODE(\"A\")", "65"),
            ("UNICODE(T!A5)", "128512"),
            ("UNICODE(\"\")", VALUE),
            ("UNICHAR(65)", "\"A\""),
            ("UNICHAR(128512)", "\"😀\""),
            ("UNICHAR(55296)", NA),
            ("UNICHAR(0)", VALUE),
            ("UNICHAR(1114112)", VALUE),
            ("CHAR(65)", "\"A\""),
            ("CHAR(200)", "(not-computed \"a code page's character\")"),
            ("CHAR(0)", VALUE),
            ("CHAR(256)", VALUE),
            ("CODE(\"A\")", "65"),
            ("CODE(\"é\")", "(not-computed \"a code page's code\")"),
            ("CODE(\"\")", VALUE),
        ]);
    }

    #[test]
    fn texts_are_joined_and_a_range_s_cells_row_by_row() {
        holds(&[
            ("CONCATENATE(\"a\",1,TRUE)", "\"a1TRUE\""),
            ("CONCATENATE(T!A9,\"x\")", NA),
            ("CONCAT(T!B1:B3)", "\"xz\""),
            ("CONCAT(\"a\",T!B1:B3,\"b\")", "\"axzb\""),
            ("CONCAT(T!A8:A9)", NA),
            ("TEXTJOIN(\",\",TRUE,T!B1:B3)", "\"x,z\""),
            ("TEXTJOIN(\",\",FALSE,T!B1:B3)", "\"x,,z\""),
            ("TEXTJOIN(\",\",TRUE,\"a\",\"\",\"b\")", "\"a,b\""),
            ("TEXTJOIN(\",\",FALSE,\"a\",\"\",\"b\")", "\"a,,b\""),
            ("TEXTJOIN(\"\",FALSE,T!B1:B3)", "\"xz\""),
            ("TEXTJOIN(T!D1:D2,TRUE,\"a\",\"b\",\"c\")", "\"a-b+c\""),
            ("TEXTJOIN(\",\",TRUE,T!A9)", NA),
            ("LEN(TEXTJOIN(\",\",FALSE,T!B:B))", VALUE),
        ]);
    }

    #[test]
    fn text_draws_a_value_by_its_code() {
        holds(&[
            ("TEXT(1234.5,\"#,##0.00\")", "\"1,234.50\""),
            ("TEXT(45292,\"yyyy-mm-dd\")", "\"2024-01-01\""),
            ("TEXT(45292,\"dddd, mmmm d\")", "\"Monday, January 1\""),
            ("TEXT(0.75,\"h:mm AM/PM\")", "\"6:00 PM\""),
            ("TEXT(\"12\",\"0.00\")", "\"12.00\""),
            ("TEXT(\"abc\",\"0.00\")", "\"abc\""),
            (
                "TEXT(\"abc\",\"0;0;0;[@]\")",
                "(not-computed \"TEXT's code\")",
            ),
            ("TEXT(\"abc\",\"0;0;0;\"\"<\"\"@\"\">\"\"\")", "\"<abc>\""),
            ("TEXT(TRUE,\"0\")", "\"TRUE\""),
            ("TEXT(-1,\"yyyy\")", VALUE),
            ("TEXT(1.5,\"General\")", "\"1.5\""),
            ("TEXT(5,\"@\")", "\"5\""),
            ("TEXT(5,\"\")", "\"\""),
            ("TEXT(5,\"[Red]0\")", "(not-computed \"TEXT's code\")"),
            ("TEXT(T!A9,\"0\")", NA),
        ]);
    }
}
