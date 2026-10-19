//! Excel's criteria (KERNEL.8, slice a, 2026-10-18): what a criterion of
//! `COUNTIF` and `SUMIF` asks of a cell, and the two functions, moved down
//! from `frazaro-core`'s Excel library with their meanings and their tests,
//! since a vector's index (`super::index`) keys by a criterion's equality and
//! a cartridge counts duplicates (Alonzo's `CART.13`).
//!
//! A criterion reads as Excel reads one: a number, a truth value or an error
//! matches the cells equal to it, and an empty cell is the number 0; a text
//! may begin with `=`, `<>`, `<`, `>`, `<=` or `>=`, and what follows is a
//! number when it spells one, `TRUE` or `FALSE` when it spells either, blank
//! when empty, and a text pattern otherwise, `*` any run of characters, `?`
//! one, `~` the next as itself. Two numbers equal to fifteen significant
//! digits are equal, as Excel's operators hold them
//! (`eval::compare_numbers`). A text that spells a number is equal to the
//! number and to nothing else of it: `<>` holds of it, `<` and `>` do not; a
//! comparison by size holds only between values of one kind; `<>` holds for
//! every other cell that `=` would not count, an empty cell included. The
//! rules Excel's save of the lookups' fixture settled: an empty cell is 0
//! (`Criteria!H16`), `"<>20"` counts the text `"20"` (`H4`, `H20`),
//! `0.1+0.2` meets 0.3 (`H32`), and `"é"` meets `"É"` (`H34`, `H36`).
//!
//! Every text here is compared as `value::fold_char` leaves its characters,
//! lowered beyond ASCII too, character by character as it is read
//! (`value::fold_cmp`), so that no comparison makes a string, where the
//! core's folded each cell's text into a new string at every test; and a
//! pattern is matched by the walk that remembers its last `*`, linear in the
//! text for each, allocating nothing, where the core's recursion tried every
//! split at each `*` and collected two vectors of characters at every test
//! (the scoping's Found). Both answer as the core's did, folding alike,
//! which a test holds over generated patterns.

use std::cmp::Ordering;

use super::eval::{compare_numbers, finite, Arg, Ctx, Reason};
use super::graph::Area;
use super::index::walk;
use super::value::{fold_char, fold_cmp, fold_eq, fold_text, text_to_number, ErrorKind, Value};

/// What a criterion of `COUNTIF` or `SUMIF` asks of a cell.
#[derive(Clone, Debug, PartialEq)]
pub enum Criterion {
    Eq(Target),
    Ne(Target),
    Lt(Target),
    Gt(Target),
    Le(Target),
    Ge(Target),
}

/// What a criterion compares a cell with.
#[derive(Clone, Debug, PartialEq)]
pub enum Target {
    Number(f64),
    /// A text with `*`, `?` and `~` as Excel's wildcards, folded
    /// (`value::fold_text`).
    Text(String),
    Bool(bool),
    Error(ErrorKind),
    Blank,
}

impl Target {
    /// Whether a text target holds a wildcard, which only a scan matches.
    pub fn is_pattern(&self) -> bool {
        matches!(self, Target::Text(t) if has_wildcard(t))
    }
}

/// Whether a text holds one of Excel's wildcards, `*`, `?` or `~`.
pub fn has_wildcard(t: &str) -> bool {
    t.bytes().any(|b| matches!(b, b'*' | b'?' | b'~'))
}

/// A criterion as Excel reads one (see the module's comment).
pub fn criterion(v: &Value) -> Criterion {
    let text = match v {
        Value::Number(n) => return Criterion::Eq(Target::Number(*n)),
        Value::Bool(b) => return Criterion::Eq(Target::Bool(*b)),
        Value::Error(k) => return Criterion::Eq(Target::Error(*k)),
        // An empty cell is 0, as Excel's save counts it: the zeros of a
        // range, not its blanks (`Criteria!H16`).
        Value::Empty => return Criterion::Eq(Target::Number(0.0)),
        Value::Text(t) => t.as_str(),
    };
    let (op, rest): (fn(Target) -> Criterion, &str) = if let Some(r) = text.strip_prefix("<>") {
        (Criterion::Ne, r)
    } else if let Some(r) = text.strip_prefix("<=") {
        (Criterion::Le, r)
    } else if let Some(r) = text.strip_prefix(">=") {
        (Criterion::Ge, r)
    } else if let Some(r) = text.strip_prefix('<') {
        (Criterion::Lt, r)
    } else if let Some(r) = text.strip_prefix('>') {
        (Criterion::Gt, r)
    } else if let Some(r) = text.strip_prefix('=') {
        (Criterion::Eq, r)
    } else {
        (Criterion::Eq, text)
    };
    let target = if rest.is_empty() {
        Target::Blank
    } else if let Some(n) = text_to_number(rest) {
        Target::Number(n)
    } else if let Some(k) = ErrorKind::parse(rest) {
        Target::Error(k)
    } else {
        let folded = fold_text(rest);
        match folded.as_str() {
            "true" => Target::Bool(true),
            "false" => Target::Bool(false),
            _ => Target::Text(folded),
        }
    };
    op(target)
}

/// Whether a cell's value meets the criterion (see the module's comment).
pub fn meets(cell: &Value, c: &Criterion) -> bool {
    use Ordering::*;
    match c {
        Criterion::Eq(t) => equal(cell, t),
        // A text spelling the number is not the number to `<>`, as Excel's
        // save counts it (`Criteria!H4`, `H20`): only a number equal to it
        // fails.
        Criterion::Ne(Target::Number(n)) => {
            !matches!(cell, Value::Number(x) if compare_numbers(*x, *n) == Equal)
        }
        Criterion::Ne(t) => !equal(cell, t),
        Criterion::Lt(t) => sized(cell, t) == Some(Less),
        Criterion::Gt(t) => sized(cell, t) == Some(Greater),
        Criterion::Le(t) => matches!(sized(cell, t), Some(Less | Equal)),
        Criterion::Ge(t) => matches!(sized(cell, t), Some(Greater | Equal)),
    }
}

/// Whether a cell is equal to a target, as `=` and `<>` ask: a number to
/// fifteen digits, a text spelling a number to the number, a text by
/// Excel's wildcards, a truth value, an error, and the blank an empty cell
/// or an empty text.
fn equal(cell: &Value, target: &Target) -> bool {
    match (cell, target) {
        (Value::Number(x), Target::Number(n)) => compare_numbers(*x, *n) == Ordering::Equal,
        (Value::Text(t), Target::Number(n)) => {
            text_to_number(t).is_some_and(|x| compare_numbers(x, *n) == Ordering::Equal)
        }
        (Value::Text(t), Target::Text(pattern)) => wildcard(pattern, t),
        (Value::Bool(b), Target::Bool(want)) => b == want,
        (Value::Error(k), Target::Error(want)) => k == want,
        (Value::Empty, Target::Blank) => true,
        (Value::Text(t), Target::Blank) => t.is_empty(),
        _ => false,
    }
}

/// How a cell compares by size with a target, as `<`, `>`, `<=` and `>=`
/// ask: a number with a number, to fifteen digits; a text with a text, a
/// match of its wildcards equal; `None` between two kinds, a text spelling a
/// number with the number among them.
fn sized(cell: &Value, target: &Target) -> Option<Ordering> {
    match (cell, target) {
        (Value::Number(x), Target::Number(n)) => Some(compare_numbers(*x, *n)),
        (Value::Text(t), Target::Text(pattern)) => Some(if wildcard(pattern, t) {
            Ordering::Equal
        } else {
            fold_cmp(t, pattern)
        }),
        _ => None,
    }
}

/// One token of a pattern: `*` any run of characters, `?` one, or a
/// character to match without case, `~`'s next character among them.
#[derive(Clone, Copy, PartialEq, Eq)]
enum Token {
    Star,
    Any,
    Char(char),
}

/// The token at byte `at` of a pattern and the byte after it; `None` at the
/// end, and at a `~` with nothing after it, which matches nothing.
#[inline]
fn token(pattern: &str, at: usize) -> Option<(Token, usize)> {
    let mut chars = pattern[at..].chars();
    let c = chars.next()?;
    let after = at + c.len_utf8();
    Some(match c {
        '*' => (Token::Star, after),
        '?' => (Token::Any, after),
        '~' => {
            let next = chars.next()?;
            (Token::Char(next), after + next.len_utf8())
        }
        other => (Token::Char(other), after),
    })
}

/// Excel's wildcards, a pattern against a text, both compared as
/// `value::fold_char` leaves their characters: `*` any run of characters,
/// `?` one, `~` the next character
/// itself, a `~` with nothing after it matching nothing more; a pattern with
/// none of the three is the text itself. The walk that remembers its last
/// `*`: on a mismatch it lets that `*` take one more character and goes on
/// from just after it, which is every split a recursion would try that can
/// still match, so that a pattern costs at most its length times the
/// text's, and nothing is collected.
pub fn wildcard(pattern: &str, text: &str) -> bool {
    if !has_wildcard(pattern) {
        return fold_eq(pattern, text);
    }
    let (mut p, mut t) = (0usize, 0usize);
    // The pattern just after the last `*`, and where in the text it took up.
    let mut star: Option<(usize, usize)> = None;
    loop {
        let step = match token(pattern, p) {
            Some((Token::Star, after)) => {
                star = Some((after, t));
                p = after;
                continue;
            }
            Some((Token::Any, after)) => {
                text[t..].chars().next().map(|c| (after, t + c.len_utf8()))
            }
            Some((Token::Char(want), after)) => match text[t..].chars().next() {
                Some(have) if fold_char(have) == fold_char(want) => {
                    Some((after, t + have.len_utf8()))
                }
                _ => None,
            },
            None if t == text.len() => return true,
            None => None,
        };
        match step {
            Some((np, nt)) => {
                p = np;
                t = nt;
            }
            None => {
                // The last `*` takes one more character, or the match fails.
                let Some((after, from)) = star else {
                    return false;
                };
                let Some(c) = text[from..].chars().next() else {
                    return false;
                };
                let from = from + c.len_utf8();
                star = Some((after, from));
                p = after;
                t = from;
            }
        }
    }
}

/// One area from an argument, for the range of `SUMIF` and `COUNTIF`.
pub(crate) fn one_area(arg: Option<&Arg>) -> Option<Area> {
    match arg {
        Some(Arg::Area { areas, .. }) if areas.len() == 1 => Some(areas[0]),
        _ => None,
    }
}

/// `COUNTIF(range, criterion)`: the cells of the range meeting the
/// criterion; a criterion a blank meets counts the empty cells of the whole
/// range, as Excel counts them, however far it reaches. Answered from the
/// range's index where the environment keeps one and the criterion is an
/// equality or its `<>` (`super::index`), by the scan below otherwise; the
/// two are one meaning, which a test holds.
pub(crate) fn countif(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    if args.len() != 2 {
        return Ok(Value::Error(ErrorKind::Value));
    }
    let Some(area) = one_area(args.first()) else {
        return Ok(Value::Error(ErrorKind::Value));
    };
    let c = criterion(&ctx.one(&args[1])?);
    if ctx.env.model().sheets.get(area.sheet).is_none() {
        return Ok(Value::Error(ErrorKind::Ref));
    }
    if let Some(indexes) = ctx.env.indexes() {
        if let Some(answer) = indexes.count(ctx, area, &c) {
            return answer;
        }
    }
    scan_count(ctx, area, &c)
}

/// `COUNTIF`'s scan, its meaning: each cell the range holds read in order,
/// a reason the answer at once, then the cells meeting the criterion
/// counted, with the empty cells of the whole range where a blank meets it.
/// The cells walked as an index's build walks them, by the array a sheet is
/// kept as where it is one.
pub(crate) fn scan_count(ctx: &Ctx<'_>, area: Area, c: &Criterion) -> Result<Value, Reason> {
    if ctx.env.model().sheets.get(area.sheet).is_none() {
        return Ok(Value::Error(ErrorKind::Ref));
    }
    let (mut count, mut present) = (0u64, 0u64);
    let mut stop: Option<Reason> = None;
    walk(ctx, &area, |_, _, read| match read {
        Ok(v) => {
            present += 1;
            count += u64::from(meets(v, c));
            true
        }
        Err(r) => {
            stop = Some(r.clone());
            false
        }
    });
    if let Some(r) = stop {
        return Err(r);
    }
    if meets(&Value::Empty, c) {
        count += cells_of(&area) - present;
    }
    Ok(Value::Number(count as f64))
}

/// How many cells an area spans, every one counted, held or not.
pub(crate) fn cells_of(area: &Area) -> u64 {
    u64::from(area.bottom - area.top + 1) * u64::from(area.right - area.left + 1)
}

/// `SUMIF(range, criterion, [sum_range])`: the numbers of the sum range
/// (the range itself when none is given) at the places where the range's
/// cell meets the criterion, an error at such a place the answer; the sum
/// range is read from its top-left cell in the range's shape, as Excel
/// reads it. Answered from the pair's index
/// where the environment keeps one and the criterion is an equality, by the
/// scan below otherwise. A sum range smaller than the range, which Excel
/// reads past the cells it names, is labelled as written: a plan orders a
/// cell after the cells its references name, and no other, so that a cell
/// read past them could be one not computed yet (KERNEL.8, its Found).
pub(crate) fn sumif(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    if args.len() < 2 || args.len() > 3 {
        return Ok(Value::Error(ErrorKind::Value));
    }
    let Some(area) = one_area(args.first()) else {
        return Ok(Value::Error(ErrorKind::Value));
    };
    let c = criterion(&ctx.one(&args[1])?);
    let sums = match args.get(2) {
        None => area,
        Some(Arg::Area {
            areas,
            written,
            moved,
        }) if areas.len() == 1 => {
            let s = areas[0];
            let sums = sum_area(&area, &s);
            if sums.bottom > s.bottom || sums.right > s.right {
                return Err(as_written(written, *moved));
            }
            sums
        }
        Some(_) => return Ok(Value::Error(ErrorKind::Value)),
    };
    if ctx.env.model().sheets.get(sums.sheet).is_none() {
        return Ok(Value::Error(ErrorKind::Ref));
    }
    if let Some(indexes) = ctx.env.indexes() {
        if let Some(answer) = indexes.sum(ctx, area, sums, &c) {
            return answer;
        }
    }
    scan_sum(ctx, area, sums, &c)
}

/// A reference's label as written at the cell computing it, as `Ctx::one`
/// labels a range where one value is wanted.
pub(crate) fn as_written(written: &str, moved: (i64, i64)) -> Reason {
    Reason::Construct(crate::refers::shift_a1_references(
        written, moved.0, moved.1,
    ))
}

/// The sum range of `SUMIF` as Excel reads it: from its top-left cell, in
/// the criteria range's shape.
pub(crate) fn sum_area(area: &Area, s: &Area) -> Area {
    Area {
        sheet: s.sheet,
        top: s.top,
        left: s.left,
        bottom: s.top.saturating_add(area.bottom - area.top),
        right: s.left.saturating_add(area.right - area.left),
    }
}

/// `SUMIF`'s scan, its meaning: each cell the sum range holds read in order,
/// a reason the answer at once; at a number or an error, the cell at the
/// same place in the criteria range read, an absent cell being empty, and
/// where it meets the criterion the number added from 0 in that order, or
/// the error the answer at once. An error at a place not met is passed by,
/// as Excel's save sums it (`Criteria!H23`). The sum range walked as an
/// index's build walks it.
pub(crate) fn scan_sum(
    ctx: &Ctx<'_>,
    area: Area,
    sums: Area,
    c: &Criterion,
) -> Result<Value, Reason> {
    if ctx.env.model().sheets.get(sums.sheet).is_none() {
        return Ok(Value::Error(ErrorKind::Ref));
    }
    let mut total = 0.0;
    let mut stop: Option<Result<ErrorKind, Reason>> = None;
    walk(ctx, &sums, |row, col, read| {
        let held = match read {
            Ok(Value::Number(n)) => Ok(*n),
            Ok(Value::Error(k)) => Err(*k),
            Ok(_) => return true,
            Err(r) => {
                stop = Some(Err(r.clone()));
                return false;
            }
        };
        let at = (
            area.sheet,
            area.top + (row - sums.top),
            area.left + (col - sums.left),
        );
        match ctx.env.read(at) {
            Ok(tested) if meets(&tested, c) => match held {
                Ok(n) => {
                    total += n;
                    true
                }
                Err(k) => {
                    stop = Some(Ok(k));
                    false
                }
            },
            Ok(_) => true,
            Err(r) => {
                stop = Some(Err(r));
                false
            }
        }
    });
    match stop {
        Some(Ok(k)) => Ok(Value::Error(k)),
        Some(Err(r)) => Err(r),
        None => Ok(finite(total)),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// The core's matcher as it stood before this slice (`excel.rs`,
    /// `KERNEL.7`), kept as the reference the linear walk is held to.
    fn recursive(pattern: &str, text: &str) -> bool {
        fn go(p: &[char], t: &[char]) -> bool {
            match p.split_first() {
                None => t.is_empty(),
                Some(('*', rest)) => (0..=t.len()).any(|i| go(rest, &t[i..])),
                Some(('?', rest)) => !t.is_empty() && go(rest, &t[1..]),
                Some(('~', rest)) => match (rest.split_first(), t.split_first()) {
                    (Some((want, rest2)), Some((have, t2))) => want == have && go(rest2, t2),
                    (None, _) => t.is_empty(),
                    _ => false,
                },
                Some((want, rest)) => match t.split_first() {
                    Some((have, t2)) => want == have && go(rest, t2),
                    None => false,
                },
            }
        }
        let pattern = fold_text(pattern);
        let text = fold_text(text);
        let p: Vec<char> = pattern.chars().collect();
        let t: Vec<char> = text.chars().collect();
        if !p.iter().any(|c| matches!(c, '*' | '?' | '~')) {
            return p == t;
        }
        go(&p, &t)
    }

    #[test]
    fn criteria_read_as_excel_reads_them() {
        assert_eq!(
            criterion(&Value::Text(">=5".into())),
            Criterion::Ge(Target::Number(5.0))
        );
        assert_eq!(
            criterion(&Value::Text("<>".into())),
            Criterion::Ne(Target::Blank)
        );
        assert_eq!(
            criterion(&Value::Text("=TRUE".into())),
            Criterion::Eq(Target::Bool(true))
        );
        assert_eq!(
            criterion(&Value::Text("a*".into())),
            Criterion::Eq(Target::Text("a*".into()))
        );
        assert_eq!(
            criterion(&Value::Number(3.0)),
            Criterion::Eq(Target::Number(3.0))
        );
        assert!(wildcard("a*", "abc"));
        assert!(wildcard("a?c", "abc"));
        assert!(!wildcard("a?c", "abbc"));
        assert!(wildcard("10~*", "10*"));
        assert!(wildcard("*", ""));
        assert!(meets(
            &Value::Text("5".into()),
            &Criterion::Eq(Target::Number(5.0))
        ));
        assert!(
            !meets(
                &Value::Text("5".into()),
                &Criterion::Gt(Target::Number(1.0))
            ),
            "a text is not compared by size against a number"
        );
        assert!(meets(&Value::Empty, &Criterion::Ne(Target::Number(5.0))));
        assert!(meets(
            &Value::Text("ABC".into()),
            &Criterion::Eq(Target::Text("abc".into()))
        ));
        // KERNEL.8, the rules Excel's save of the lookups' fixture settled.
        assert_eq!(
            criterion(&Value::Empty),
            Criterion::Eq(Target::Number(0.0)),
            "an empty cell is 0"
        );
        assert!(!meets(&Value::Empty, &criterion(&Value::Empty)));
        assert!(meets(&Value::Number(0.0), &criterion(&Value::Empty)));
        let text = |s: &str| Value::Text(s.to_string());
        let not_20 = criterion(&text("<>20"));
        assert!(meets(&text("20"), &not_20), "a spelled number is not it");
        assert!(!meets(&Value::Number(20.0), &not_20));
        assert!(meets(&text("20"), &criterion(&text("=20"))));
        assert!(!meets(&text("20"), &criterion(&text(">=20"))));
        assert!(!meets(&text("20"), &criterion(&text("<=20"))));
        let tenths = Value::Number(0.1 + 0.2);
        assert!(meets(&tenths, &criterion(&Value::Number(0.3))));
        assert!(meets(&tenths, &criterion(&text("<=0.3"))));
        assert!(!meets(&tenths, &criterion(&text("<>0.3"))));
        assert!(meets(
            &text("0.30000000000000004"),
            &criterion(&Value::Number(0.3))
        ));
        assert_eq!(
            criterion(&text("<>ÉCLAIR")),
            Criterion::Ne(Target::Text("éclair".into()))
        );
        assert!(meets(&text("É"), &criterion(&text("é"))));
        assert!(!meets(&text("e"), &criterion(&text("é"))));
        assert!(wildcard("é*", "Éclair"));
        assert!(wildcard("?clair", "Éclair"));
    }

    #[test]
    fn the_walk_matches_as_the_recursion_did() {
        // Every pattern of up to four tokens over an alphabet holding the
        // three wildcards, a letter in two cases and a letter beyond ASCII,
        // against every text of up to four characters, a letter in two
        // cases, one beyond ASCII and a wildcard among them: the linear
        // walk answers as the core's recursion answered, both folding.
        let alphabet = ['a', 'A', 'b', '*', '?', '~', 'É'];
        let mut patterns: Vec<String> = vec![String::new()];
        let mut layer: Vec<String> = vec![String::new()];
        for _ in 0..4 {
            let mut next = Vec::new();
            for p in &layer {
                for c in alphabet {
                    let mut q = p.clone();
                    q.push(c);
                    next.push(q);
                }
            }
            patterns.extend(next.iter().cloned());
            layer = next;
        }
        let letters = ['a', 'A', 'é', '*'];
        let mut texts: Vec<String> = vec![String::new()];
        let mut layer: Vec<String> = vec![String::new()];
        for _ in 0..4 {
            let mut next = Vec::new();
            for t in &layer {
                for c in letters {
                    let mut u = t.clone();
                    u.push(c);
                    next.push(u);
                }
            }
            texts.extend(next.iter().cloned());
            layer = next;
        }
        let mut checked = 0usize;
        for p in &patterns {
            for t in &texts {
                assert_eq!(wildcard(p, t), recursive(p, t), "pattern {p:?}, text {t:?}");
                checked += 1;
            }
        }
        assert_eq!(checked, 2_801 * 341, "every pair");
    }

    #[test]
    fn a_pattern_of_many_stars_costs_its_length_times_the_text_s() {
        // A pattern the recursion took about n^k steps over: twenty stars,
        // each before an `a`, against a text of 4,000 `a`s and no `b`.
        let pattern = format!("{}b", "*a".repeat(20));
        let text = "a".repeat(4_000);
        assert!(!wildcard(&pattern, &text));
        assert!(wildcard(&format!("{}*", "*a".repeat(20)), &text));
    }
}
