//! The formula-reference reader (PORT.8, slice 8b): `VLA_Refers.bas`
//! (`AXM.7`) ported one arm per arm, held to the refers golden.
//!
//! Pure text to records: the references inside a formula's text, read from
//! the text alone, which is what `refers(from, to)` is built from. A parser
//! for references, not for the formula language: operators and functions
//! are stepped over, and the binders of `LET` and `LAMBDA` read as names (a
//! recorded limit). The kinds, the record and the three readings are the
//! reference's (`src/VLA_Refers.bas`, whose header says what each reads);
//! `scripts/refers.txt` to `scripts/refers_golden.txt`, written by the
//! reference's `VlaWriteRefersGolden`, is the golden
//! `the_refers_golden_is_reproduced` holds this module to, and
//! `tools/check_refers_golden.ps1` holds the pair's shape.
//!
//! Over the one scan, four renderers: [`spell`], the second field of a
//! `refers` row in the treaty's spelling; [`r1c1`], the formula relative to
//! a cell as Excel's `FormulaR1C1` writes it; [`shift_a1_references`], the
//! writer's mover (every relative part moved by a fill's distance, `#REF!`
//! off the sheet), which was `sheet/refs.rs`'s hand scanner until this
//! slice and is held by the build golden; and [`resolve_books`], a file's
//! `[1]Sheet1!A1` to the formula bar's `[Rates.xlsx]Sheet1!A1` through the
//! external books the package names. [`quote_sheet`] is the one sheet
//! quoting rule, used at both ends of a `refers` row, so that the relation
//! joins to itself.

use crate::intrinsics::fold;
use crate::sheet::{column_letters, MAX_COLUMN, MAX_ROW};

/// The eleven kinds of reference.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum Kind {
    Cell,
    Range,
    Column,
    Row,
    Name,
    Structured,
    External,
    /// A 3D span, `Jan:Dec!A1`.
    ThreeD,
    Spill,
    /// `INDIRECT` or `OFFSET`, unreadable by name.
    Unreadable,
    /// `#REF!`, a broken reference.
    Broken,
}

impl Kind {
    /// The word the golden's first field spells.
    pub fn word(self) -> &'static str {
        match self {
            Kind::Cell => "cell",
            Kind::Range => "range",
            Kind::Column => "column",
            Kind::Row => "row",
            Kind::Name => "name",
            Kind::Structured => "structured",
            Kind::External => "external",
            Kind::ThreeD => "3d",
            Kind::Spill => "spill",
            Kind::Unreadable => "unreadable",
            Kind::Broken => "broken",
        }
    }
}

/// The reference part's own shape where it has numbers.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Shape {
    None,
    Cell,
    Range,
    Column,
    Row,
}

/// One reference, as [`scan`] found it: `FormulaRef` of the reference.
/// Positions are character indices into the text the scan was given.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct FormulaRef {
    pub kind: Kind,
    pub shape: Shape,
    /// External: the text inside the `[ ]` (`1`, `Book.xlsx`); empty otherwise.
    pub book: String,
    /// The sheet as named: unquoted, `''` undoubled; a span as `Jan:Dec`;
    /// empty where none is written.
    pub sheet_name: String,
    /// The whole token as it stands in the formula.
    pub written: String,
    /// The part after any qualifier or `#REF!`: `$A$1`, `A1:B2`, `A1#`,
    /// `Rate`, `Sales[Amount]`, `INDIRECT`, `#REF!`.
    pub part: String,
    pub row1: u32,
    pub col1: u32,
    pub row2: u32,
    pub col2: u32,
    pub row_abs1: bool,
    pub col_abs1: bool,
    pub row_abs2: bool,
    pub col_abs2: bool,
    /// Where the token begins.
    pub pos: usize,
    /// Its length in characters.
    pub span: usize,
    /// Where `part` begins, for a renderer.
    pub part_pos: usize,
}

impl FormulaRef {
    fn at(pos: usize) -> FormulaRef {
        FormulaRef {
            kind: Kind::Name,
            shape: Shape::None,
            book: String::new(),
            sheet_name: String::new(),
            written: String::new(),
            part: String::new(),
            row1: 0,
            col1: 0,
            row2: 0,
            col2: 0,
            row_abs1: false,
            col_abs1: false,
            row_abs2: false,
            col_abs2: false,
            pos,
            span: 0,
            part_pos: pos,
        }
    }

    /// The canonical name of an unreadable call, as its row spells it.
    pub fn unreadable_name(&self) -> &'static str {
        if fold(&self.part) == "indirect" {
            "INDIRECT"
        } else {
            "OFFSET"
        }
    }
}

// ---------------------------------------------------------------------
//  The scan (RefersScan)
// ---------------------------------------------------------------------

/// Every reference in a formula's text, in formula order, duplicates kept.
/// A leading `=` is stepped over.
pub fn scan(formula_text: &str) -> Vec<FormulaRef> {
    let s: Vec<char> = formula_text.chars().collect();
    let n = s.len();
    let mut refs = Vec::new();
    let mut i = 0;
    if n >= 1 && s[0] == '=' {
        i = 1;
    }
    while i < n {
        if s[i] == '"' {
            i = skip_delimited(&s, i, '"');
        } else if !read_reference(&s, &mut i, &mut refs) {
            i += 1;
        }
    }
    refs
}

/// What can begin at `i`: a qualifier and its part, a bare part, a bare
/// structured reference, an error literal, a function name, a number, or a
/// quoted text that qualifies nothing. Moves `i` past what it read and
/// answers true; false where nothing of the kind begins.
fn read_reference(s: &[char], i: &mut usize, refs: &mut Vec<FormulaRef>) -> bool {
    let ch = s[*i];
    let mut r = FormulaRef::at(*i);
    let mut j = *i;
    if read_qualifier(s, &mut j, &mut r) {
        r.part_pos = j;
        if read_part(s, &mut j, &mut r) {
            finish_ref(s, j, r, refs);
        }
        *i = j;
        return true;
    }
    if ch == '\'' {
        *i = skip_delimited(s, *i, '\'');
        true
    } else if ch == '[' {
        let end = bracket_group_end(s, *i);
        r.kind = Kind::Structured;
        r.part_pos = *i;
        r.part = s[*i..end].iter().collect();
        finish_ref(s, end, r, refs);
        *i = end;
        true
    } else if ch == '#' || ch == '$' || is_ident_char(ch) {
        r.part_pos = *i;
        if read_part(s, &mut j, &mut r) {
            finish_ref(s, j, r, refs);
        }
        if j == *i {
            j = *i + 1;
        }
        *i = j;
        true
    } else {
        false
    }
}

/// A sheet qualifier at `j`: `'quoted'!`, `[book]sheet!`, `[book]!`,
/// `sheet!`, or a span `first:last!`. Fills the book and the sheet, moves
/// `j` past the `!` and answers true; otherwise leaves `j` and answers false.
fn read_qualifier(s: &[char], j: &mut usize, r: &mut FormulaRef) -> bool {
    let n = s.len();
    let ch = s[*j];
    if ch == '\'' {
        let Some(k) = delimited_end(s, *j, '\'') else {
            return false;
        };
        if s.get(k) != Some(&'!') {
            return false;
        }
        let inner: String = s[*j + 1..k - 1].iter().collect();
        let q = inner.replace("''", "'");
        match (q.find('['), q.find(']')) {
            (Some(b), Some(e)) if e > b => {
                r.book = q[b + 1..e].to_string();
                r.sheet_name = q[e + 1..].to_string();
            }
            _ => r.sheet_name = q,
        }
        *j = k + 1;
        return true;
    }
    let mut k = *j;
    let mut e = 0usize;
    if ch == '[' {
        let Some(e_rel) = s[*j + 1..].iter().position(|&c| c == ']') else {
            return false;
        };
        e = *j + 1 + e_rel;
        if let Some(b_rel) = s[*j + 1..].iter().position(|&c| c == '[') {
            if *j + 1 + b_rel < e {
                return false; // a nested [: a structured reference, not a book
            }
        }
        k = e + 1;
        if s.get(k) == Some(&'!') {
            r.book = s[*j + 1..e].iter().collect();
            r.sheet_name.clear();
            *j = k + 1;
            return true;
        }
    }
    let mut p = k;
    while p < n && is_ident_char(s[p]) {
        p += 1;
    }
    if p == k {
        return false;
    }
    let mut q: String = s[k..p].iter().collect();
    if s.get(p) == Some(&':') {
        let mut p2 = p + 1;
        while p2 < n && is_ident_char(s[p2]) {
            p2 += 1;
        }
        if p2 > p + 1 && s.get(p2) == Some(&'!') {
            q.push(':');
            q.extend(s[p + 1..p2].iter());
            p = p2;
        }
    }
    if s.get(p) != Some(&'!') {
        return false;
    }
    if ch == '[' {
        r.book = s[*j + 1..e].iter().collect();
    }
    r.sheet_name = q;
    *j = p + 1;
    true
}

/// The reference part at `j`, after a qualifier or bare: `#REF!` (with the
/// cell or range a deleted sheet leaves after it), a cell, a range, a
/// column, a row, a structured reference with its table, a name, either's
/// spill `#`, or an unreadable call's name. A function name, `TRUE`,
/// `FALSE`, a number or another error literal is stepped over with false;
/// `j` stays only where nothing begins.
fn read_part(s: &[char], j: &mut usize, r: &mut FormulaRef) -> bool {
    if s.get(*j) == Some(&'#') {
        let k = error_literal_end(s, *j);
        let tok: String = s[*j..k].iter().collect();
        *j = k;
        if fold(&tok) != "#ref!" {
            return false;
        }
        r.kind = Kind::Broken;
        r.part = tok;
        if read_cell_like(s, j, r) {
            r.part_pos = *j - r.part.chars().count();
        }
        return true;
    }
    let k = token_end(s, *j);
    if k == *j {
        return false;
    }
    let tok: String = s[*j..k].iter().collect();
    let nxt = s.get(k).copied();
    if nxt == Some('(') {
        *j = k;
        let f = fold(&tok);
        if f == "indirect" || f == "offset" {
            r.kind = Kind::Unreadable;
            r.part = tok;
            return true;
        }
        return false;
    }
    if nxt == Some('[') {
        let end = bracket_group_end(s, k);
        r.kind = Kind::Structured;
        r.part = s[*j..end].iter().collect();
        *j = end;
        return true;
    }
    if read_cell_like(s, j, r) {
        r.kind = match r.shape {
            Shape::Cell => {
                if s.get(*j) == Some(&'#') {
                    r.part.push('#');
                    *j += 1;
                    Kind::Spill
                } else {
                    Kind::Cell
                }
            }
            Shape::Range => Kind::Range,
            Shape::Column => Kind::Column,
            Shape::Row | Shape::None => Kind::Row,
        };
        return true;
    }
    let mut first = tok.chars().next();
    if first == Some('$') {
        first = tok.chars().nth(1);
    }
    if first == Some('.') || first.is_some_and(|c| c.is_ascii_digit()) {
        *j = k; // a number: 1.5, 2E (its +3 follows), .5
        return false;
    }
    let f = fold(&tok);
    if f == "true" || f == "false" {
        *j = k;
        return false;
    }
    r.kind = Kind::Name;
    r.part = tok;
    *j = k;
    if s.get(*j) == Some(&'#') {
        r.kind = Kind::Spill;
        r.part.push('#');
        *j += 1;
    }
    true
}

/// A cell, range, whole column or whole row at `j`, its second half read
/// after a colon when the two halves are of one shape. Sets the shape, the
/// part and the numbers, moves `j` and answers true; otherwise leaves
/// everything as it was.
fn read_cell_like(s: &[char], j: &mut usize, r: &mut FormulaRef) -> bool {
    let k = token_end(s, *j);
    if k == *j {
        return false;
    }
    let tok: String = s[*j..k].iter().collect();
    let mut tok2 = String::new();
    let mut k2 = k;
    if s.get(k) == Some(&':') {
        k2 = token_end(s, k + 1);
        if k2 > k + 1 {
            tok2 = s[k + 1..k2].iter().collect();
        }
    }
    let mut shape = Shape::None;
    let (mut r1, mut c1, mut r2, mut c2) = (0u32, 0u32, 0u32, 0u32);
    let (mut ra1, mut ca1, mut ra2, mut ca2) = (false, false, false, false);
    if let Some((row, col, ra, ca)) = parse_cell_part(&tok) {
        r1 = row;
        c1 = col;
        ra1 = ra;
        ca1 = ca;
        shape = Shape::Cell;
        if !tok2.is_empty() {
            if let Some((row, col, ra, ca)) = parse_cell_part(&tok2) {
                r2 = row;
                c2 = col;
                ra2 = ra;
                ca2 = ca;
                shape = Shape::Range;
            }
        }
    } else if !tok2.is_empty() {
        if let (Some((col, ca)), Some((col2, ca_two))) =
            (parse_col_part(&tok), parse_col_part(&tok2))
        {
            c1 = col;
            ca1 = ca;
            c2 = col2;
            ca2 = ca_two;
            shape = Shape::Column;
        } else if let (Some((row, ra)), Some((row2, ra_two))) =
            (parse_row_part(&tok), parse_row_part(&tok2))
        {
            r1 = row;
            ra1 = ra;
            r2 = row2;
            ra2 = ra_two;
            shape = Shape::Row;
        }
    }
    if shape == Shape::None {
        return false;
    }
    r.shape = shape;
    r.row1 = r1;
    r.col1 = c1;
    r.row_abs1 = ra1;
    r.col_abs1 = ca1;
    if shape == Shape::Cell {
        r.part = tok;
        *j = k;
    } else {
        r.row2 = r2;
        r.col2 = c2;
        r.row_abs2 = ra2;
        r.col_abs2 = ca2;
        r.part = s[*j..k2].iter().collect();
        *j = k2;
    }
    true
}

/// The record is complete: its text, and the kind a qualifier decides.
fn finish_ref(s: &[char], j: usize, mut r: FormulaRef, refs: &mut Vec<FormulaRef>) {
    r.span = j - r.pos;
    r.written = s[r.pos..j].iter().collect();
    if !r.book.is_empty() {
        r.kind = Kind::External;
    } else if r.sheet_name.contains(':') {
        r.kind = Kind::ThreeD;
    }
    refs.push(r);
}

// ---------------------------------------------------------------------
//  The pieces of a part
// ---------------------------------------------------------------------

/// `$?LETTERS$?DIGITS`, the whole token, within the sheet: (row, column,
/// row absolute, column absolute).
fn parse_cell_part(tok: &str) -> Option<(u32, u32, bool, bool)> {
    let c: Vec<char> = tok.chars().collect();
    let n = c.len();
    let mut p = 0;
    let mut col_abs = false;
    if c.first() == Some(&'$') {
        col_abs = true;
        p = 1;
    }
    let mut k = p;
    while k < n && c[k].is_ascii_alphabetic() {
        k += 1;
    }
    let letters: String = c[p..k].iter().collect();
    if letters.is_empty() || letters.len() > 3 {
        return None;
    }
    let mut row_abs = false;
    if c.get(k) == Some(&'$') {
        row_abs = true;
        k += 1;
    }
    let p2 = k;
    while k < n && c[k].is_ascii_digit() {
        k += 1;
    }
    let digits: String = c[p2..k].iter().collect();
    if digits.is_empty() || digits.len() > 7 {
        return None;
    }
    if k < n {
        return None; // something after the digits: a name (ABC123X)
    }
    let col = letters_to_col(&letters)?;
    let row: u32 = digits.parse().ok()?;
    if row == 0 || row > MAX_ROW {
        return None;
    }
    Some((row, col, row_abs, col_abs))
}

/// `$?LETTERS`, the whole token, for one end of a whole-column reference.
fn parse_col_part(tok: &str) -> Option<(u32, bool)> {
    let (letters, abs) = match tok.strip_prefix('$') {
        Some(rest) => (rest, true),
        None => (tok, false),
    };
    let count = letters.chars().count();
    if count == 0 || count > 3 || !letters.chars().all(|c| c.is_ascii_alphabetic()) {
        return None;
    }
    Some((letters_to_col(letters)?, abs))
}

/// `$?DIGITS`, the whole token, for one end of a whole-row reference.
fn parse_row_part(tok: &str) -> Option<(u32, bool)> {
    let (digits, abs) = match tok.strip_prefix('$') {
        Some(rest) => (rest, true),
        None => (tok, false),
    };
    let count = digits.chars().count();
    if count == 0 || count > 7 || !digits.chars().all(|c| c.is_ascii_digit()) {
        return None;
    }
    let row: u32 = digits.parse().ok()?;
    if row == 0 || row > MAX_ROW {
        return None;
    }
    Some((row, abs))
}

/// A = 1 ... XFD = 16384; `None` past the last column.
fn letters_to_col(letters: &str) -> Option<u32> {
    let mut v: u32 = 0;
    for ch in fold(letters).chars() {
        let c = ch as u32;
        if !(97..=122).contains(&c) {
            return None;
        }
        v = v * 26 + (c - 96);
    }
    if v == 0 || v > MAX_COLUMN {
        return None;
    }
    Some(v)
}

// ---------------------------------------------------------------------
//  Walking the text
// ---------------------------------------------------------------------

/// The position after the run of identifier characters and `$` signs at `j`.
fn token_end(s: &[char], j: usize) -> usize {
    let mut k = j;
    while k < s.len() && (s[k] == '$' || is_ident_char(s[k])) {
        k += 1;
    }
    k
}

/// The position after the error literal at `j` (its `#`): letters, digits,
/// `/` and `_`, then a closing `!` or `?` when there is one.
fn error_literal_end(s: &[char], j: usize) -> usize {
    let mut k = j + 1;
    while k < s.len() && (s[k].is_ascii_alphanumeric() || s[k] == '/' || s[k] == '_') {
        k += 1;
    }
    if matches!(s.get(k), Some('!') | Some('?')) {
        k += 1;
    }
    k
}

/// The position after the closing delimiter of the delimited text at `j`;
/// a doubled delimiter stays inside. `None` when it never closes.
fn delimited_end(s: &[char], j: usize, d: char) -> Option<usize> {
    let n = s.len();
    let mut k = j + 1;
    while k < n {
        if s[k] != d {
            k += 1;
        } else if s.get(k + 1) == Some(&d) {
            k += 2;
        } else {
            return Some(k + 1);
        }
    }
    None
}

/// [`delimited_end`], or the end of the text when the delimiter never closes.
fn skip_delimited(s: &[char], j: usize, d: char) -> usize {
    delimited_end(s, j, d).unwrap_or(s.len())
}

/// The position after the `]` that closes the `[` at `i`, nested brackets
/// counted and a `'` inside taken as the escape it is there; the end of the
/// text when it never closes.
fn bracket_group_end(s: &[char], i: usize) -> usize {
    let n = s.len();
    let mut k = i;
    let mut depth: i32 = 0;
    while k < n {
        let ch = s[k];
        if ch == '\'' {
            k += 2;
        } else {
            if ch == '[' {
                depth += 1;
            }
            if ch == ']' {
                depth -= 1;
            }
            k += 1;
            if depth == 0 {
                return k;
            }
        }
    }
    n
}

/// A letter, a digit, `_` or `.`, or any character past ASCII (a letter in
/// another alphabet, as a name or a sheet may hold).
fn is_ident_char(ch: char) -> bool {
    ch.is_ascii_alphanumeric() || ch == '_' || ch == '.' || (ch as u32) > 127
}

// ---------------------------------------------------------------------
//  The spellings (RefersSpell, RefersQuoteSheet)
// ---------------------------------------------------------------------

/// The second field of a `refers` row for one record, the home sheet
/// qualifying a reference written bare.
pub fn spell(r: &FormulaRef, home_sheet: &str) -> String {
    let q = if r.sheet_name.is_empty() {
        format!("{}!", quote_sheet(home_sheet))
    } else {
        format!("{}!", quote_sheet(&r.sheet_name))
    };
    match r.kind {
        Kind::Cell | Kind::Range | Kind::Column | Kind::Row => format!("{q}{}", spelled_part(r)),
        Kind::Spill => {
            if r.shape == Shape::Cell {
                format!("{q}{}#", spelled_part(r))
            } else {
                r.written.clone()
            }
        }
        Kind::Broken => {
            if r.shape != Shape::None {
                r.written.clone()
            } else {
                format!("{q}#REF!")
            }
        }
        Kind::Unreadable => format!("(unreadable \"{}\")", r.unreadable_name()),
        Kind::Name | Kind::Structured | Kind::External | Kind::ThreeD => r.written.clone(),
    }
}

/// The part in A1 with its marks dropped and its letters upper-cased.
fn spelled_part(r: &FormulaRef) -> String {
    match r.shape {
        Shape::Cell => format!("{}{}", column_letters(r.col1), r.row1),
        Shape::Range => format!(
            "{}{}:{}{}",
            column_letters(r.col1),
            r.row1,
            column_letters(r.col2),
            r.row2
        ),
        Shape::Column => format!("{}:{}", column_letters(r.col1), column_letters(r.col2)),
        Shape::Row => format!("{}:{}", r.row1, r.row2),
        Shape::None => String::new(),
    }
}

/// A sheet name as a reference spells it: quoted when it must be, an
/// apostrophe inside doubled.
pub fn quote_sheet(sheet_name: &str) -> String {
    if sheet_needs_quotes(sheet_name) {
        format!("'{}'", sheet_name.replace('\'', "''"))
    } else {
        sheet_name.to_string()
    }
}

/// Excel's rule as the owner's pass of 2026-10-04 confirmed it: a character
/// outside letters, digits, `_` and `.`, a leading digit, or a name that is
/// itself a cell or an R1C1 reference.
fn sheet_needs_quotes(sheet_name: &str) -> bool {
    let Some(first) = sheet_name.chars().next() else {
        return true;
    };
    if first.is_ascii_digit() {
        return true;
    }
    if !sheet_name.chars().all(is_ident_char) {
        return true;
    }
    if parse_cell_part(sheet_name).is_some() {
        return true;
    }
    is_r1c1_shaped(sheet_name)
}

/// `R`, `C`, `R1`, `C1`, `RC`, `R1C1`, `R12C34`: a sheet so named is quoted.
fn is_r1c1_shaped(sheet_name: &str) -> bool {
    let f = fold(sheet_name);
    let mut chars = f.chars();
    let mut seen_c = match chars.next() {
        Some('c') => true,
        Some('r') => false,
        _ => return false,
    };
    for ch in chars {
        if ch == 'c' && !seen_c {
            seen_c = true;
        } else if !ch.is_ascii_digit() {
            return false;
        }
    }
    true
}

// ---------------------------------------------------------------------
//  R1C1, the mover, the books, and the home cell
// ---------------------------------------------------------------------

/// The text with every part that has numbers rewritten by `render`,
/// everything else as written: the shape every renderer shares.
fn rewrite(formula_text: &str, render: impl Fn(&FormulaRef) -> String) -> String {
    let chars: Vec<char> = formula_text.chars().collect();
    let mut out = String::with_capacity(formula_text.len() + 8);
    let mut p = 0;
    for r in &scan(formula_text) {
        if r.shape != Shape::None {
            out.extend(chars[p..r.part_pos].iter());
            out.push_str(&render(r));
            if r.kind == Kind::Spill {
                out.push('#');
            }
            p = r.pos + r.span;
        }
    }
    out.extend(chars[p..].iter());
    out
}

/// The formula rendered in R1C1 relative to the cell at (`home_row`,
/// `home_col`), as Excel's `FormulaR1C1` writes it.
pub fn r1c1(formula_text: &str, home_row: u32, home_col: u32) -> String {
    rewrite(formula_text, |r| r1c1_part(r, home_row, home_col))
}

/// One part in R1C1: a whole column or row whose two ends render alike is
/// written once, as Excel writes it.
fn r1c1_part(r: &FormulaRef, home_row: u32, home_col: u32) -> String {
    let cell = |row: u32, ra: bool, col: u32, ca: bool| {
        format!(
            "R{}C{}",
            rel_spell(row, ra, home_row),
            rel_spell(col, ca, home_col)
        )
    };
    match r.shape {
        Shape::Cell => cell(r.row1, r.row_abs1, r.col1, r.col_abs1),
        Shape::Range => format!(
            "{}:{}",
            cell(r.row1, r.row_abs1, r.col1, r.col_abs1),
            cell(r.row2, r.row_abs2, r.col2, r.col_abs2)
        ),
        Shape::Column => {
            let a = format!("C{}", rel_spell(r.col1, r.col_abs1, home_col));
            let b = format!("C{}", rel_spell(r.col2, r.col_abs2, home_col));
            if a == b {
                a
            } else {
                format!("{a}:{b}")
            }
        }
        Shape::Row => {
            let a = format!("R{}", rel_spell(r.row1, r.row_abs1, home_row));
            let b = format!("R{}", rel_spell(r.row2, r.row_abs2, home_row));
            if a == b {
                a
            } else {
                format!("{a}:{b}")
            }
        }
        Shape::None => String::new(),
    }
}

/// `5` for an absolute 5; `[2]`, `[-3]` or nothing for a relative one.
fn rel_spell(n: u32, is_abs: bool, home: u32) -> String {
    if is_abs {
        n.to_string()
    } else if n == home {
        String::new()
    } else {
        format!("[{}]", i64::from(n) - i64::from(home))
    }
}

/// The formula's text with every relative reference moved `d_row` rows down
/// and `d_col` columns right, as the formula of a cell that far from the
/// first cell of a fill; a part moved off the sheet becomes `#REF!`, as
/// Excel writes it. The writer's mover (PORT.7, slice 7c), held by the build
/// golden; a shared formula's children in the reader are rendered with it.
pub fn shift_a1_references(formula: &str, d_row: i64, d_col: i64) -> String {
    if d_row == 0 && d_col == 0 {
        return formula.to_string();
    }
    rewrite(formula, |r| shifted_part(r, d_row, d_col))
}

fn shifted(n: u32, is_abs: bool, d: i64, max: u32) -> Option<u32> {
    if is_abs {
        return Some(n);
    }
    let m = i64::from(n) + d;
    (m >= 1 && m <= i64::from(max)).then_some(m as u32)
}

fn mark(is_abs: bool) -> &'static str {
    if is_abs {
        "$"
    } else {
        ""
    }
}

fn shifted_part(r: &FormulaRef, d_row: i64, d_col: i64) -> String {
    let cell = |row: u32, ra: bool, col: u32, ca: bool| -> Option<String> {
        let col = shifted(col, ca, d_col, MAX_COLUMN)?;
        let row = shifted(row, ra, d_row, MAX_ROW)?;
        Some(format!(
            "{}{}{}{}",
            mark(ca),
            column_letters(col),
            mark(ra),
            row
        ))
    };
    let rendered = match r.shape {
        Shape::Cell => cell(r.row1, r.row_abs1, r.col1, r.col_abs1),
        Shape::Range => match (
            cell(r.row1, r.row_abs1, r.col1, r.col_abs1),
            cell(r.row2, r.row_abs2, r.col2, r.col_abs2),
        ) {
            (Some(a), Some(b)) => Some(format!("{a}:{b}")),
            _ => None,
        },
        Shape::Column => match (
            shifted(r.col1, r.col_abs1, d_col, MAX_COLUMN),
            shifted(r.col2, r.col_abs2, d_col, MAX_COLUMN),
        ) {
            (Some(a), Some(b)) => Some(format!(
                "{}{}:{}{}",
                mark(r.col_abs1),
                column_letters(a),
                mark(r.col_abs2),
                column_letters(b)
            )),
            _ => None,
        },
        Shape::Row => match (
            shifted(r.row1, r.row_abs1, d_row, MAX_ROW),
            shifted(r.row2, r.row_abs2, d_row, MAX_ROW),
        ) {
            (Some(a), Some(b)) => Some(format!(
                "{}{}:{}{}",
                mark(r.row_abs1),
                a,
                mark(r.row_abs2),
                b
            )),
            _ => None,
        },
        Shape::None => None,
    };
    rendered.unwrap_or_else(|| "#REF!".to_string())
}

/// A file's `[1]Sheet1!A1` and `[1]!Rate` as the formula bar shows them,
/// `[Rates.xlsx]Sheet1!A1`, through the external books in the order the
/// file numbers them (`books[0]` is `[1]`); a number past the list, or a
/// book with no name, stays as the file wrote it.
pub fn resolve_books(formula_text: &str, books: &[String]) -> String {
    let chars: Vec<char> = formula_text.chars().collect();
    let mut out = String::with_capacity(formula_text.len() + 16);
    let mut p = 0;
    for r in &scan(formula_text) {
        if r.kind != Kind::External || r.book.is_empty() {
            continue;
        }
        if !r.book.chars().all(|c| c.is_ascii_digit()) {
            continue;
        }
        let Some(name) = r
            .book
            .parse::<usize>()
            .ok()
            .and_then(|n| n.checked_sub(1))
            .and_then(|i| books.get(i))
            .filter(|name| !name.is_empty())
        else {
            continue;
        };
        // The book's text sits right after the token's first bracket.
        let Some(open) = r.written.find('[') else {
            continue;
        };
        let book_len = r.book.len();
        let mut token = String::with_capacity(r.written.len() + name.len());
        token.push_str(&r.written[..open + 1]);
        token.push_str(name);
        token.push_str(&r.written[open + 1 + book_len..]);
        out.extend(chars[p..r.pos].iter());
        out.push_str(&token);
        p = r.pos + r.span;
    }
    out.extend(chars[p..].iter());
    out
}

/// A formula's own cell as a `refers` row spells its first field,
/// `Model!B3` or `'Q1 Data'!C5`, read with the same scan: (sheet, row,
/// column). `None` for anything else: a bare `B3`, a range, a book, more
/// text.
pub fn parse_home(home_text: &str) -> Option<(String, u32, u32)> {
    let refs = scan(home_text);
    if refs.len() != 1 {
        return None;
    }
    let r = &refs[0];
    if r.kind != Kind::Cell || r.sheet_name.is_empty() {
        return None;
    }
    if r.pos != 0 || r.span != home_text.chars().count() {
        return None;
    }
    Some((r.sheet_name.clone(), r.row1, r.col1))
}

// ---------------------------------------------------------------------
//  The golden's report (RefersReportFor, VlaWriteRefersGolden)
// ---------------------------------------------------------------------

/// One case's record, as `RefersReportFor` writes it: the body's first
/// line is the cell, the rest the formula, blank lines at its end dropped;
/// one `<kind>\t<written>\t<spelled>` line per reference or `NONE`, then
/// the `R1C1` line. `BAD HOME` names a cell line that does not read.
pub fn report_for(body: &[&str]) -> String {
    let home_text = body.first().copied().unwrap_or("");
    let mut formula: String = if body.len() > 1 {
        body[1..].join("\n")
    } else {
        String::new()
    };
    while formula.ends_with('\n') {
        formula.pop();
    }
    let Some((home_sheet, home_row, home_col)) = parse_home(home_text) else {
        return format!("BAD HOME\t{home_text}\r\n");
    };
    let refs = scan(&formula);
    let mut out = String::new();
    if refs.is_empty() {
        out.push_str("NONE\r\n");
    } else {
        for r in &refs {
            out.push_str(&format!(
                "{}\t{}\t{}\r\n",
                r.kind.word(),
                r.written,
                spell(r, &home_sheet)
            ));
        }
    }
    out.push_str(&format!("R1C1\t{}\r\n", r1c1(&formula, home_row, home_col)));
    out
}

/// `VlaWriteRefersGolden`'s text for a fixture: each `=== ` header, then
/// its case's report; the lines before the first header are ignored.
pub fn fixture_report(fixture: &str) -> String {
    let text = fixture.replace("\r\n", "\n");
    let mut out = String::new();
    let mut header: Option<&str> = None;
    let mut body: Vec<&str> = Vec::new();
    for line in text.split('\n') {
        if line.starts_with("=== ") {
            if let Some(h) = header {
                out.push_str(h);
                out.push_str("\r\n");
                out.push_str(&report_for(&body));
            }
            header = Some(line);
            body.clear();
        } else if header.is_some() {
            body.push(line);
        }
    }
    if let Some(h) = header {
        out.push_str(h);
        out.push_str("\r\n");
        out.push_str(&report_for(&body));
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    const FIXTURE: &str = include_str!("../../scripts/refers.txt");
    const GOLDEN: &str = include_str!("../../scripts/refers_golden.txt");

    #[test]
    fn the_refers_golden_is_reproduced() {
        // scripts/refers_golden.txt is written by VlaWriteRefersGolden
        // (VLA_Tests.bas) from the reference's VLA_Refers; this is the
        // reader's oracle, compared as the treaty compares a golden (LF,
        // trailing blank lines dropped).
        let norm = |s: &str| s.replace("\r\n", "\n").trim_end_matches('\n').to_string();
        let got = norm(&fixture_report(FIXTURE));
        let want = norm(GOLDEN);
        if got != want {
            let at = got
                .chars()
                .zip(want.chars())
                .position(|(a, b)| a != b)
                .unwrap_or(got.len().min(want.len()));
            let line = want[..at].matches('\n').count() + 1;
            let context = |s: &str| {
                let start = s[..at].rfind('\n').map(|i| i + 1).unwrap_or(0);
                let end = s[at..].find('\n').map(|i| at + i).unwrap_or(s.len());
                s[start..end].to_string()
            };
            panic!(
                "the refers golden differs at char {at} (line {line}):\n  want: {}\n  got:  {}",
                context(&want),
                context(&got)
            );
        }
    }

    /// Prints the fixture's report, for checking a prediction of the golden:
    /// `cargo test -p frazaro-core print_refers_report -- --ignored --nocapture`.
    #[test]
    #[ignore]
    fn print_refers_report() {
        print!("{}", fixture_report(FIXTURE));
    }

    #[test]
    fn the_record_carries_its_fields() {
        let refs = scan("='Q1 Data'!$A$1:B2");
        assert_eq!(refs.len(), 1);
        let r = &refs[0];
        assert_eq!((r.kind, r.shape), (Kind::Range, Shape::Range));
        assert_eq!((r.sheet_name.as_str(), r.book.as_str()), ("Q1 Data", ""));
        assert_eq!(r.part, "$A$1:B2");
        assert_eq!((r.row1, r.col1, r.row2, r.col2), (1, 1, 2, 2));
        assert_eq!(
            (r.row_abs1, r.col_abs1, r.row_abs2, r.col_abs2),
            (true, true, false, false)
        );
        assert_eq!((r.pos, r.span, r.part_pos), (1, 17, 11));
        let refs = scan("=#REF!A1");
        assert_eq!((refs[0].kind, refs[0].shape), (Kind::Broken, Shape::Cell));
        assert_eq!((refs[0].part.as_str(), refs[0].part_pos), ("A1", 6));
        let refs = scan("=INDIRECT(A1)");
        assert_eq!(refs.len(), 2);
        assert_eq!(refs[0].kind, Kind::Unreadable);
        assert_eq!(refs[0].unreadable_name(), "INDIRECT");
        assert_eq!(refs[1].kind, Kind::Cell);
        assert_eq!(scan("A1+B1").len(), 2);
        assert!(scan("=1+1").is_empty());
    }

    #[test]
    fn the_home_cell_parses() {
        assert_eq!(parse_home("Model!B3"), Some(("Model".to_string(), 3, 2)));
        assert_eq!(
            parse_home("'Q1 Data'!C5"),
            Some(("Q1 Data".to_string(), 5, 3))
        );
        assert_eq!(parse_home("B3"), None);
        assert_eq!(parse_home("Model!A:A"), None);
        assert_eq!(parse_home("Model!B3+1"), None);
    }

    #[test]
    fn a_sheet_is_quoted_as_a_reference_quotes_it() {
        assert_eq!(quote_sheet("Data"), "Data");
        assert_eq!(quote_sheet("Q1 Data"), "'Q1 Data'");
        assert_eq!(quote_sheet("It's"), "'It''s'");
        assert_eq!(quote_sheet("A1"), "'A1'");
        assert_eq!(quote_sheet("2024"), "'2024'");
        assert_eq!(quote_sheet("Q1.Data"), "Q1.Data");
        assert_eq!(quote_sheet("R1C1"), "'R1C1'");
        assert_eq!(quote_sheet("RC"), "'RC'");
        assert_eq!(quote_sheet("Tax_Rate"), "Tax_Rate");
        assert_eq!(quote_sheet("Umsatz\u{d6}"), "Umsatz\u{d6}");
        assert_eq!(quote_sheet(""), "''");
    }

    #[test]
    fn a_name_past_ascii_is_a_name_and_a_sheet_past_ascii_is_a_sheet() {
        let refs = scan("=Umsatz_\u{d6}*2");
        assert_eq!((refs.len(), refs[0].kind), (1, Kind::Name));
        assert_eq!(refs[0].written, "Umsatz_\u{d6}");
        let refs = scan("=\u{dc}bersicht!A1");
        assert_eq!(
            (refs[0].kind, refs[0].sheet_name.as_str()),
            (Kind::Cell, "\u{dc}bersicht")
        );
        assert_eq!(spell(&refs[0], "Model"), "\u{dc}bersicht!A1");
        assert_eq!(
            r1c1("=\u{dc}bersicht!A1", 3, 2),
            "=\u{dc}bersicht!R[-2]C[-1]"
        );
    }

    #[test]
    fn books_resolve_to_their_file_names() {
        let books = vec!["Rates.xlsx".to_string(), String::new()];
        assert_eq!(
            resolve_books("=[1]Sheet1!A1+[1]!Rate", &books),
            "=[Rates.xlsx]Sheet1!A1+[Rates.xlsx]!Rate"
        );
        assert_eq!(
            resolve_books("='[1]Q1 Data'!A1", &books),
            "='[Rates.xlsx]Q1 Data'!A1"
        );
        assert_eq!(resolve_books("=[2]Sheet1!A1", &books), "=[2]Sheet1!A1");
        assert_eq!(resolve_books("=[3]Sheet1!A1", &books), "=[3]Sheet1!A1");
        assert_eq!(
            resolve_books("=[Book.xlsx]Sheet1!A1", &books),
            "=[Book.xlsx]Sheet1!A1"
        );
        assert_eq!(resolve_books("=A1+[Amount]", &books), "=A1+[Amount]");
    }

    // The mover's tests, as sheet/refs.rs held them (PORT.7, slice 7c).
    fn s(f: &str, r: i64, c: i64) -> String {
        shift_a1_references(f, r, c)
    }

    #[test]
    fn relative_parts_move_and_absolute_parts_stay() {
        assert_eq!(s("B2+B3", 1, 0), "B3+B4");
        assert_eq!(s("$B$2+B$2+$B2", 1, 1), "$B$2+C$2+$B3");
        assert_eq!(s("SUM(A1:A3)", 0, 1), "SUM(B1:B3)");
        assert_eq!(
            s("IFS(B2>3,\"big\",TRUE,\"small\")", 2, 0),
            "IFS(B4>3,\"big\",TRUE,\"small\")"
        );
        assert_eq!(s("b2*2", 1, 0), "B3*2");
        assert_eq!(s("Z1+AA1", 0, 1), "AA1+AB1");
    }

    #[test]
    fn whole_columns_rows_and_sheets() {
        assert_eq!(s("SUM(A:A)", 0, 1), "SUM(B:B)");
        assert_eq!(s("SUM($A:A)", 0, 1), "SUM($A:B)");
        assert_eq!(s("SUM(1:3)", 1, 0), "SUM(2:4)");
        assert_eq!(s("Data!A1+'Q1 Data'!B2", 1, 0), "Data!A2+'Q1 Data'!B3");
        assert_eq!(s("'It''s'!A1", 1, 0), "'It''s'!A2");
        assert_eq!(
            s("SUM(Jan:Dec!A1)+[1]Sheet1!A1+A1#", 1, 0),
            "SUM(Jan:Dec!A2)+[1]Sheet1!A2+A2#"
        );
    }

    #[test]
    fn what_is_not_a_reference_is_left_alone() {
        assert_eq!(s("\"A1\"&A1", 1, 0), "\"A1\"&A2");
        assert_eq!(s("\"say \"\"A1\"\"\"&A1", 1, 0), "\"say \"\"A1\"\"\"&A2");
        assert_eq!(s("LOG10(A1)", 1, 0), "LOG10(A2)");
        assert_eq!(s("Table1[Col]+A1", 1, 0), "Table1[Col]+A2");
        assert_eq!(
            s("Table1[[#This Row],[A1]]", 1, 0),
            "Table1[[#This Row],[A1]]"
        );
        assert_eq!(s("TRUE+A1", 1, 0), "TRUE+A2");
        assert_eq!(s("ABC123X+A1", 1, 0), "ABC123X+A2");
        assert_eq!(s("Rate*A1", 1, 0), "Rate*A2");
        assert_eq!(s("_xlfn.IFS(A1,1)", 1, 0), "_xlfn.IFS(A2,1)");
        assert_eq!(s("1.5+A1", 1, 0), "1.5+A2");
        assert_eq!(s("A1", 0, 0), "A1");
        assert_eq!(s("INDIRECT(\"A1\")+A1", 1, 0), "INDIRECT(\"A1\")+A2");
    }

    #[test]
    fn off_the_sheet_is_ref_error() {
        assert_eq!(s("A1", -1, 0), "#REF!");
        assert_eq!(s("A1+B1", 0, -1), "#REF!+A1");
        assert_eq!(s("XFD1", 0, 1), "#REF!");
        assert_eq!(s("A1:B2", -1, 0), "#REF!");
        assert_eq!(s("$A$1", -5, -5), "$A$1");
        assert_eq!(s("SUM(A:A)", 0, -1), "SUM(#REF!)");
    }
}
