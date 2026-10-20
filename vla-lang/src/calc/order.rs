//! Excel's order of texts (KERNEL.8, slice c, 2026-10-20): how `<`, `>`,
//! `<=` and `>=`, a lookup's sorted search, an index's sorted texts and a
//! criterion by size order two texts, and when two texts are one, as Excel
//! 365's save of the text fixture ranks 163 characters and texts among
//! themselves (`scripts/recalc/texts.txt`, its Order sheet, saved by Excel
//! as `texts_saved.xlsx`).
//!
//! Excel orders texts by levels, as Windows sorts words. First, each
//! character's place, which [`SYMBOLS`] and the digits and letters give:
//! the space, the no-break space, the punctuation and currency, `+`, `<`,
//! `=`, `>` and the other signs, then the digits, then the letters A to Z,
//! each with its accented forms, then Greek, Cyrillic, an emoji and a Chinese
//! character, the five beyond Latin the save ranks. So `"10"` stands before
//! `"2"`, `"a b"` before `"a1"`, `"~"` before `"a"` and `"é"` before `"f"`.
//! Second, where the places are one, the accents, character by character
//! from the first: none, then acute, grave, circumflex, diaeresis and ring
//! (`"resume"` before `"résumé"`). Case counts for nothing (`"A"` and `"a"`
//! one, `"résumé"` and `"Résumé"`). Last, the apostrophe, the hyphen and the
//! em dash, passed by until then and then compared in that order, a text
//! without one first, and of two with one at different places, the one
//! with it later first: `"coop"` before `"co-op"` before `"cop"`, `"a"`
//! before `"a-"` before `"-a"`. `ß` is `ss` and `æ` is `ae` at every level,
//! so that `"ß"="ss"` is TRUE.
//!
//! The table holds the characters the save ranks and what follows from its
//! levels: ASCII's 95 printable characters in both cases, the no-break
//! space, `¡`, `¿`, `’`, `€`, `×`, `÷`, `°` and `µ`, the em dash, the
//! letters with an acute, a grave, a circumflex or a diaeresis (the save
//! ranks the four on a, e and u, three on o and the diaeresis on y, and the
//! rest take their accent's place, the same on every letter), `å`, `ç` and
//! `ñ`, `æ` and `ß`, and `α`, `ω`, `ж`, the emoji `😀` and `日`, each in both
//! cases where it has two. A character beyond it has a place Excel knows
//! and the save does not show (`ø` and `œ`, `ı` and `İ`, a tab, any other
//! script), so a comparison by size that such a character decides is
//! labelled where its answer turns on it (`a character's order`), never
//! guessed: [`compare`] weighs it after the table, by its code point, and
//! says the answer is not known. A comparison the table decides before such
//! a character is reached, in two texts' common beginning or after their
//! first difference, is known. Whether two texts are one is answered
//! always: one where their characters are, case aside (`value::fold_char`),
//! `ß` and `ss` one, `æ` and `ae` one, which beyond the table is the fold's
//! claim, as the lookups' save held it beyond ASCII.

use std::cmp::Ordering;

use super::eval::Reason;
use super::value::fold_char;

/// The symbols in Excel's order, the space first; a symbol's place is its
/// index here plus one.
const SYMBOLS: [char; 40] = [
    ' ', '\u{A0}', '!', '"', '#', '$', '%', '&', '(', ')', '*', ',', '.', '/', ':', ';', '?', '@',
    '[', '\\', ']', '^', '_', '`', '{', '|', '}', '~', '\u{A1}', '\u{BF}', '\u{2019}', '\u{20AC}',
    '+', '<', '=', '>', '\u{D7}', '\u{F7}', '\u{B0}', '\u{B5}',
];

/// The place of `0`; the digits follow it, after every symbol.
const DIGITS: u32 = 100;

/// The place of `a`; the letters follow it, after every digit.
const LETTERS: u32 = 200;

/// The place of `α`; `ω`, `ж`, `😀` and `日` follow it, after every letter.
const OTHERS: u32 = 300;

/// Where a character beyond the table is weighed from, by its code point:
/// after every character in it.
const BEYOND: u32 = 0x1_0000;

/// The accents, in Excel's order; none is 0.
const ACUTE: u8 = 1;
const GRAVE: u8 = 2;
const CIRCUMFLEX: u8 = 3;
const DIAERESIS: u8 = 4;
const RING: u8 = 5;
const CEDILLA: u8 = 6;
const TILDE: u8 = 7;

/// A character's part in the order, as its fold leaves it.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Element {
    /// A place and an accent: a symbol, a digit or a letter.
    Place(u32, u8),
    /// Two unaccented places: `æ`, `ß`.
    Two(u32, u32),
    /// Passed by until last: the apostrophe 1, the hyphen 2, the em dash 3.
    Last(u8),
    /// Beyond the table, weighed after it by its code point.
    Beyond(u32),
}

/// An ASCII character's element, as [`ASCII`] holds it.
const fn ascii_element(b: u8) -> Element {
    match b {
        b'\'' => Element::Last(1),
        b'-' => Element::Last(2),
        b'0'..=b'9' => Element::Place(DIGITS + (b - b'0') as u32, 0),
        b'a'..=b'z' => Element::Place(LETTERS + (b - b'a') as u32, 0),
        b'A'..=b'Z' => Element::Place(LETTERS + (b - b'A') as u32, 0),
        _ => {
            let mut i = 0;
            while i < SYMBOLS.len() {
                if SYMBOLS[i] as u32 == b as u32 {
                    return Element::Place(i as u32 + 1, 0);
                }
                i += 1;
            }
            Element::Beyond(BEYOND + b as u32)
        }
    }
}

/// Each ASCII character's element, both cases alike; a control character
/// is beyond the table.
const ASCII: [Element; 128] = {
    let mut table = [Element::Beyond(BEYOND); 128];
    let mut b = 0;
    while b < 128 {
        table[b] = ascii_element(b as u8);
        b += 1;
    }
    table
};

/// A letter's place with an accent.
fn accented(letter: u8, accent: u8) -> Element {
    Element::Place(LETTERS + u32::from(letter - b'a'), accent)
}

/// A character's element, the character as its fold leaves it.
fn element(c: char) -> Element {
    if c.is_ascii() {
        return ASCII[c as usize];
    }
    beyond_ascii(c)
}

/// [`element`] beyond ASCII, out of line.
#[inline(never)]
fn beyond_ascii(c: char) -> Element {
    match c {
        'á' => accented(b'a', ACUTE),
        'à' => accented(b'a', GRAVE),
        'â' => accented(b'a', CIRCUMFLEX),
        'ä' => accented(b'a', DIAERESIS),
        'å' => accented(b'a', RING),
        'ç' => accented(b'c', CEDILLA),
        'é' => accented(b'e', ACUTE),
        'è' => accented(b'e', GRAVE),
        'ê' => accented(b'e', CIRCUMFLEX),
        'ë' => accented(b'e', DIAERESIS),
        'í' => accented(b'i', ACUTE),
        'ì' => accented(b'i', GRAVE),
        'î' => accented(b'i', CIRCUMFLEX),
        'ï' => accented(b'i', DIAERESIS),
        'ñ' => accented(b'n', TILDE),
        'ó' => accented(b'o', ACUTE),
        'ò' => accented(b'o', GRAVE),
        'ô' => accented(b'o', CIRCUMFLEX),
        'ö' => accented(b'o', DIAERESIS),
        'ú' => accented(b'u', ACUTE),
        'ù' => accented(b'u', GRAVE),
        'û' => accented(b'u', CIRCUMFLEX),
        'ü' => accented(b'u', DIAERESIS),
        'ý' => accented(b'y', ACUTE),
        'ÿ' => accented(b'y', DIAERESIS),
        'æ' => Element::Two(LETTERS, LETTERS + u32::from(b'e' - b'a')),
        'ß' => Element::Two(
            LETTERS + u32::from(b's' - b'a'),
            LETTERS + u32::from(b's' - b'a'),
        ),
        '\u{2014}' => Element::Last(3),
        'α' => Element::Place(OTHERS, 0),
        'ω' => Element::Place(OTHERS + 1, 0),
        'ж' => Element::Place(OTHERS + 2, 0),
        '\u{1F600}' => Element::Place(OTHERS + 3, 0),
        '日' => Element::Place(OTHERS + 4, 0),
        _ => match SYMBOLS.iter().position(|&s| s == c) {
            Some(i) => Element::Place(i as u32 + 1, 0),
            None => Element::Beyond(BEYOND + u32::from(c)),
        },
    }
}

/// A text's elements, character by character as the fold leaves each.
fn elements(t: &str) -> impl Iterator<Item = Element> + '_ {
    t.chars().map(|c| element(fold_char(c)))
}

/// A text's places in order, the apostrophe, the hyphen and the em dash
/// passed by, each with whether the table holds it.
fn places(t: &str) -> impl Iterator<Item = (u32, bool)> + '_ {
    elements(t)
        .flat_map(|e| match e {
            Element::Place(p, _) => [Some((p, true)), None],
            Element::Two(p, q) => [Some((p, true)), Some((q, true))],
            Element::Last(_) => [None, None],
            Element::Beyond(p) => [Some((p, false)), None],
        })
        .flatten()
}

/// A text's accents, one at each of its places.
fn accents(t: &str) -> impl Iterator<Item = u8> + '_ {
    elements(t)
        .flat_map(|e| match e {
            Element::Place(_, a) => [Some(a), None],
            Element::Two(..) => [Some(0), Some(0)],
            Element::Last(_) => [None, None],
            Element::Beyond(_) => [Some(0), None],
        })
        .flatten()
}

/// A text's apostrophes, hyphens and em dashes, each with how many places
/// stand before it.
fn lasts(t: &str) -> impl Iterator<Item = (u32, u8)> + '_ {
    let mut before = 0u32;
    elements(t).filter_map(move |e| match e {
        Element::Place(..) | Element::Beyond(_) => {
            before += 1;
            None
        }
        Element::Two(..) => {
            before += 2;
            None
        }
        Element::Last(w) => Some((before, w)),
    })
}

/// How two texts compare in Excel's order, and whether the table decided
/// it.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) struct Compared {
    /// The order, a character beyond the table weighed after it by its code
    /// point: a total order, which a sort may take.
    pub(crate) order: Ordering,
    /// Whether the table decided it, as Excel's would; two texts one are
    /// always known.
    pub(crate) known: bool,
}

/// The bytes of two texts' common beginning, character by character as the
/// fold leaves each, which orders nothing: in the first text and in the
/// second, which may differ where a character folds to one of another
/// length (the Kelvin sign to `k`). A shortcut, byte by byte while both are
/// ASCII: [`compare`]'s levels pass a common beginning by as they go, equal
/// places beyond the table among them, and answer alike without it.
fn beginning(a: &str, b: &str) -> (usize, usize) {
    let (x, y) = (a.as_bytes(), b.as_bytes());
    let mut i = 0;
    while i < x.len()
        && i < y.len()
        && x[i] < 0x80
        && y[i] < 0x80
        && x[i].eq_ignore_ascii_case(&y[i])
    {
        i += 1;
    }
    if i == x.len() || i == y.len() || (x[i] < 0x80 && y[i] < 0x80) {
        return (i, i);
    }
    let (mut j, mut k) = (i, i);
    for (p, q) in a[i..].chars().zip(b[i..].chars()) {
        if fold_char(p) != fold_char(q) {
            break;
        }
        j += p.len_utf8();
        k += q.len_utf8();
    }
    (j, k)
}

/// How two texts compare in Excel's order (see the module's comment): their
/// common beginning passed by ([`beginning`]); then their places, the first
/// that differ deciding, known where the table holds both and unknown where
/// it holds either not, two equal places passed by whatever they are, and
/// where one text runs out of places, the other first, known where the
/// other holds a place the table holds; then their accents; then their
/// apostrophes, hyphens and em dashes, a later one first.
pub(crate) fn compare(a: &str, b: &str) -> Compared {
    let (i, j) = beginning(a, b);
    let (a, b) = (&a[i..], &b[j..]);
    let (mut xs, mut ys) = (places(a), places(b));
    loop {
        match (xs.next(), ys.next()) {
            (Some((p, pk)), Some((q, qk))) => {
                if p != q {
                    return Compared {
                        order: p.cmp(&q),
                        known: pk && qk,
                    };
                }
            }
            (None, None) => break,
            (None, Some((_, k))) => {
                return Compared {
                    order: Ordering::Less,
                    known: k || ys.any(|(_, k)| k),
                }
            }
            (Some((_, k)), None) => {
                return Compared {
                    order: Ordering::Greater,
                    known: k || xs.any(|(_, k)| k),
                }
            }
        }
    }
    // Every place one, each beyond the table matched by itself: the rest is
    // the table's.
    let order = accents(a).cmp(accents(b)).then_with(|| lasts_order(a, b));
    Compared { order, known: true }
}

/// Two texts' apostrophes, hyphens and em dashes compared: a text without
/// one before a text with one, and of two at different places the later
/// first, else the apostrophe before the hyphen before the em dash.
fn lasts_order(a: &str, b: &str) -> Ordering {
    let (mut xs, mut ys) = (lasts(a), lasts(b));
    loop {
        match (xs.next(), ys.next()) {
            (None, None) => return Ordering::Equal,
            (None, Some(_)) => return Ordering::Less,
            (Some(_), None) => return Ordering::Greater,
            (Some((p, v)), Some((q, w))) => {
                let o = q.cmp(&p).then(v.cmp(&w));
                if o != Ordering::Equal {
                    return o;
                }
            }
        }
    }
}

/// A text's characters as the equality of texts keys them: each as the fold
/// leaves it, `ß` as `ss` and `æ` as `ae`, so that two texts are one where
/// their keys are.
pub(crate) fn keyed(t: &str) -> impl Iterator<Item = char> + '_ {
    t.chars()
        .flat_map(|c| match fold_char(c) {
            'ß' => [Some('s'), Some('s')],
            'æ' => [Some('a'), Some('e')],
            f => [Some(f), None],
        })
        .flatten()
}

/// The key a text is one with another by ([`keyed`]): the form an index
/// keeps a text in.
pub(crate) fn key(t: &str) -> String {
    if t.is_ascii() {
        t.to_ascii_lowercase()
    } else {
        keyed(t).collect()
    }
}

/// Whether two texts are one in Excel's order: [`compare`]'s `Equal`, with
/// no string made, byte by byte while both are ASCII.
#[inline]
pub(crate) fn equal(a: &str, b: &str) -> bool {
    if a.is_ascii() && b.is_ascii() {
        a.eq_ignore_ascii_case(b)
    } else {
        keyed(a).eq(keyed(b))
    }
}

/// Whether the table holds every character of a text, so that its order
/// against any other text the table holds is known.
pub(crate) fn within(t: &str) -> bool {
    if t.is_ascii() {
        t.bytes().all(|b| (0x20..0x7F).contains(&b))
    } else {
        elements(t).all(|e| !matches!(e, Element::Beyond(_)))
    }
}

/// The label of a comparison a character beyond the table decides.
pub(crate) fn unknown() -> Reason {
    Reason::Construct("a character's order".to_string())
}

/// How two texts compare where the answer turns on their order, as `<` and
/// a sorted search ask: [`compare`]'s order where it is known or the two
/// are one, the label otherwise.
pub(crate) fn sized(a: &str, b: &str) -> Result<Ordering, Reason> {
    let c = compare(a, b);
    if c.known || c.order == Ordering::Equal {
        Ok(c.order)
    } else {
        Err(unknown())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// The texts of the save's Order sheet, each with its rank, the count of
    /// texts Excel's `<` holds below it (`texts_saved.xlsx`, `Order!B1:B163`),
    /// in rank order.
    const RANKS: &[(u32, &str)] = &[
        (0, "'"),
        (1, "-"),
        (2, "\u{2014}"),
        (3, " "),
        (4, "  "),
        (5, "\u{A0}"),
        (6, "!"),
        (7, "\""),
        (8, "#"),
        (9, "$"),
        (10, "%"),
        (11, "&"),
        (12, "("),
        (13, ")"),
        (14, "*"),
        (15, ","),
        (16, "."),
        (17, "/"),
        (18, ":"),
        (19, ";"),
        (20, "?"),
        (21, "@"),
        (22, "["),
        (23, "\\"),
        (24, "]"),
        (25, "^"),
        (26, "_"),
        (27, "`"),
        (28, "{"),
        (29, "|"),
        (30, "}"),
        (31, "~"),
        (32, "\u{A1}"),
        (33, "\u{BF}"),
        (34, "\u{2019}"),
        (35, "\u{20AC}"),
        (36, "+"),
        (37, "<"),
        (38, "="),
        (39, ">"),
        (40, "\u{D7}"),
        (41, "\u{F7}"),
        (42, "\u{B0}"),
        (43, "\u{B5}"),
        (44, "0"),
        (45, "1"),
        (46, "10"),
        (47, "2"),
        (47, "2"),
        (49, "3"),
        (50, "4"),
        (51, "5"),
        (52, "6"),
        (53, "7"),
        (54, "8"),
        (55, "9"),
        (56, "A"),
        (56, "a"),
        (58, "a-"),
        (59, "-a"),
        (60, "á"),
        (61, "à"),
        (62, "â"),
        (63, "ä"),
        (64, "å"),
        (65, "a b"),
        (66, "A1"),
        (66, "a1"),
        (68, "A10"),
        (69, "A2"),
        (70, "aa"),
        (70, "aA"),
        (70, "Aa"),
        (73, "ab"),
        (74, "a'b"),
        (75, "a-b"),
        (76, "æ"),
        (76, "ae"),
        (78, "af"),
        (79, "B"),
        (79, "b"),
        (81, "C"),
        (81, "c"),
        (83, "ç"),
        (84, "co op"),
        (85, "coo"),
        (86, "coop"),
        (87, "co-op"),
        (88, "cop"),
        (89, "D"),
        (89, "d"),
        (91, "E"),
        (91, "e"),
        (93, "é"),
        (94, "è"),
        (95, "ê"),
        (96, "ë"),
        (97, "F"),
        (97, "f"),
        (99, "G"),
        (99, "g"),
        (101, "H"),
        (101, "h"),
        (103, "I"),
        (103, "i"),
        (105, "ı"),
        (106, "İ"),
        (107, "J"),
        (107, "j"),
        (109, "K"),
        (109, "k"),
        (111, "L"),
        (111, "l"),
        (113, "M"),
        (113, "m"),
        (115, "N"),
        (115, "n"),
        (117, "ñ"),
        (118, "O"),
        (118, "o"),
        (120, "ò"),
        (121, "ô"),
        (122, "ö"),
        (123, "ø"),
        (124, "œ"),
        (125, "P"),
        (125, "p"),
        (127, "Q"),
        (127, "q"),
        (129, "R"),
        (129, "r"),
        (131, "resume"),
        (132, "résumé"),
        (132, "Résumé"),
        (134, "S"),
        (134, "s"),
        (136, "ß"),
        (136, "ss"),
        (138, "st"),
        (139, "T"),
        (139, "t"),
        (141, "U"),
        (141, "u"),
        (143, "ú"),
        (144, "ù"),
        (145, "û"),
        (146, "ü"),
        (147, "V"),
        (147, "v"),
        (149, "W"),
        (149, "w"),
        (151, "X"),
        (151, "x"),
        (153, "Y"),
        (153, "y"),
        (155, "ÿ"),
        (156, "Z"),
        (156, "z"),
        (158, "α"),
        (159, "Ω"),
        (160, "Ж"),
        (161, "\u{1F600}"),
        (162, "日"),
    ];

    #[test]
    fn the_order_is_excel_s_rank_for_rank() {
        assert_eq!(RANKS.len(), 163);
        for &(r, a) in RANKS {
            for &(s, b) in RANKS {
                let c = compare(a, b);
                if within(a) && within(b) {
                    assert!(c.known, "{a:?} {b:?}");
                    assert_eq!(c.order, r.cmp(&s), "{a:?} {b:?}");
                    assert_eq!(equal(a, b), r == s, "{a:?} {b:?}");
                } else {
                    // One of ø, œ, ı and İ: its place is the save's alone, so
                    // the order is not known, but against itself.
                    assert_eq!(c.known, a == b, "{a:?} {b:?}");
                    assert_eq!(sized(a, b).is_ok(), a == b, "{a:?} {b:?}");
                }
            }
        }
        // A control character is beyond the table, the printable ones in it.
        assert!(!within("a\tb"));
        assert!(!within("\u{7F}"));
        assert!(within(" ~"));
        // Each text but those four is the table's.
        let beyond: Vec<&str> = RANKS
            .iter()
            .map(|&(_, t)| t)
            .filter(|t| !within(t))
            .collect();
        assert_eq!(beyond, ["ı", "İ", "ø", "œ"]);
    }

    #[test]
    fn the_functions_sheet_s_comparisons_are_excel_s() {
        // texts_saved.xlsx, Functions!A173 to A192, each as Excel answers.
        let lt = |a: &str, b: &str| compare(a, b).order == Ordering::Less;
        assert!(equal("a", "A"));
        assert!(equal("é", "É"));
        assert!(!equal("é", "e"));
        assert!(equal("ß", "ss"));
        assert!(lt("a", "B"));
        assert!(!lt("B", "a"));
        assert!(!lt("Z", "a"));
        assert!(lt("-", "a"));
        assert!(lt(" ", "0"));
        assert!(lt("0", "a"));
        assert!(lt("_", "a"));
        assert!(lt("[", "a"));
        assert!(lt("~", "a"));
        assert!(lt("é", "f"));
        assert!(lt("e", "é"));
        assert!(!lt("co-op", "coop"));
        assert!(lt("coop", "co-op"));
        assert!(lt("a b", "ab"));
        assert!(lt("a", "a "));
        assert!(lt("10", "9"));
    }

    #[test]
    fn a_character_beyond_the_table_is_known_only_where_it_decides_nothing() {
        let known = |a: &str, b: &str| compare(a, b).known;
        // A common beginning orders nothing, whatever it holds.
        assert!(known("øa", "øb"));
        assert!(known("a\tb", "a\tc"));
        assert!(known("ŁÓDŹ", "łódź"));
        // A difference before the character decides.
        assert!(known("aø", "bø"));
        assert!(!known("ab", "a\u{2013}c"));
        // Where it decides, or may be passed by, it is not known.
        assert!(!known("ø", "o"));
        assert!(!known("a", "a\t"));
        assert!(!known("", "\u{1}"));
        assert!(!known("ıa", "ib"));
        // A text running out before another's place the table holds is
        // first, whatever stands between.
        assert!(known("a", "a\u{2013}b"));
        assert_eq!(compare("a", "a\u{2013}b").order, Ordering::Less);
        // The one beyond the table matched by itself: the rest decides.
        assert!(known("éøb", "eøc"));
        assert_eq!(sized("éøb", "eøc"), Ok(Ordering::Less));
        assert_eq!(sized("ø", "o"), Err(unknown()));
        // Two texts one are known, and one is never unknown to `=`.
        assert!(equal("Øre", "øRE"));
        assert!(!equal("ø", "o"));
        assert!(!equal("İ", "i"));
        assert!(!equal("ı", "i"));
    }

    #[test]
    fn the_order_is_total_and_its_equality_its_key_s() {
        let texts = [
            "",
            "a",
            "A",
            "b",
            "ab",
            "aB",
            "a-b",
            "a'b",
            "a--b",
            "-ab",
            "ab-",
            "a-",
            "-a",
            "'",
            "-",
            "\u{2014}",
            "é",
            "É",
            "e",
            "è",
            "éé",
            "eé",
            "ée",
            "æ",
            "ae",
            "Æ",
            "ß",
            "ss",
            "SS",
            "ẞ",
            "st",
            "s",
            "\u{212A}",
            "k",
            "K",
            "İ",
            "i",
            "ı",
            "ø",
            "o",
            "œ",
            "oe",
            "\t",
            "a\tb",
            "Ж",
            "ж",
            "α",
            "Ω",
            "ω",
            "日",
            "\u{1F600}",
            "10",
            "2",
            " ",
            "a b",
            "\u{A0}",
            "~",
            "co-op",
            "coop",
            "cop",
            "résumé",
            "resume",
            "ŁÓDŹ",
            "łódź",
        ];
        for a in texts {
            for b in texts {
                let ab = compare(a, b);
                assert_eq!(ab.order, compare(b, a).order.reverse(), "{a:?} {b:?}");
                assert_eq!(ab.known, compare(b, a).known, "{a:?} {b:?}");
                assert_eq!(equal(a, b), ab.order == Ordering::Equal, "{a:?} {b:?}");
                assert_eq!(equal(a, b), key(a) == key(b), "{a:?} {b:?}");
                assert_eq!(key(a), keyed(a).collect::<String>(), "{a:?}");
                for c in texts {
                    if ab.order != Ordering::Greater && compare(b, c).order != Ordering::Greater {
                        assert_ne!(compare(a, c).order, Ordering::Greater, "{a:?} {b:?} {c:?}");
                    }
                }
            }
        }
        assert_eq!(key("Straße"), "strasse");
        assert_eq!(key("Ærø"), "aerø");
    }
}
