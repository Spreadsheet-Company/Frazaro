//! A1 references inside a formula's text, shifted as Excel shifts them when
//! a formula is filled (PORT.7, slice 7c).
//!
//! New ground, said plainly: when the reference runs `Range.Formula2 =` over
//! more than one cell, Excel enters the formula in each cell with its
//! relative references moved by that cell's distance from the first, and
//! the file holds the result. Slice 7b wrote a filled formula as a shared
//! formula and left the moving to the host; a filled formula whose function
//! can return an array must instead be a dynamic-array formula in every cell
//! (sheet::ooxml, the `cm` metadata), and a dynamic-array formula cannot be
//! shared, so the writer moves the references itself.
//!
//! What moves: a cell (`B2`), a range (`B2:C4`), a whole column (`A:A`) or a
//! whole row (`1:3`), sheet-qualified or not (`Data!B2`, `'Q1 Data'!B2`),
//! each part without a `$` moved by the row or column distance, each part
//! with one left where it is. A reference moved off the sheet becomes
//! `#REF!`, as Excel writes it. What does not move: text inside a string
//! literal, a structured reference inside `[...]`, a name or a function
//! (`LOG10(`) that is spelled like a cell, `TRUE` and `FALSE`.
//!
//! This is the file-format half of `AXM.7`'s formula-reference reader, built
//! here because the writer needs it first; when `AXM.7` lands in the VBA, its
//! port is what this scanner becomes.

use super::{column_letters, MAX_COLUMN, MAX_ROW};

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct CellRef {
    col_abs: bool,
    col: u32,
    row_abs: bool,
    row: u32,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Token {
    Cell(CellRef),
    Range(CellRef, CellRef),
    /// `A:A`: (absolute?, column) twice.
    Columns((bool, u32), (bool, u32)),
    /// `1:3`: (absolute?, row) twice.
    Rows((bool, u32), (bool, u32)),
}

fn is_ident_char(c: char) -> bool {
    c.is_ascii_alphanumeric() || c == '_' || c == '.'
}

fn letters_to_col(letters: &str) -> Option<u32> {
    if letters.is_empty() || letters.len() > 3 {
        return None;
    }
    let mut col: u32 = 0;
    for c in letters.chars() {
        col = col * 26 + (c.to_ascii_uppercase() as u32 - 'A' as u32 + 1);
    }
    (col <= MAX_COLUMN).then_some(col)
}

/// `$?LETTERS$?DIGITS` at `i`, with where it ends.
fn parse_cell(chars: &[char], i: usize) -> Option<(CellRef, usize)> {
    let mut j = i;
    let col_abs = chars.get(j) == Some(&'$');
    if col_abs {
        j += 1;
    }
    let ls = j;
    while j < chars.len() && chars[j].is_ascii_alphabetic() {
        j += 1;
    }
    let letters: String = chars[ls..j].iter().collect();
    let col = letters_to_col(&letters)?;
    let row_abs = chars.get(j) == Some(&'$');
    if row_abs {
        j += 1;
    }
    let ds = j;
    while j < chars.len() && chars[j].is_ascii_digit() {
        j += 1;
    }
    if j == ds || j - ds > 7 {
        return None;
    }
    let row: u32 = chars[ds..j].iter().collect::<String>().parse().ok()?;
    if row == 0 || row > MAX_ROW {
        return None;
    }
    Some((
        CellRef {
            col_abs,
            col,
            row_abs,
            row,
        },
        j,
    ))
}

/// `$?LETTERS` at `i`, for a whole-column reference.
fn parse_col_part(chars: &[char], i: usize) -> Option<((bool, u32), usize)> {
    let mut j = i;
    let abs = chars.get(j) == Some(&'$');
    if abs {
        j += 1;
    }
    let ls = j;
    while j < chars.len() && chars[j].is_ascii_alphabetic() {
        j += 1;
    }
    let col = letters_to_col(&chars[ls..j].iter().collect::<String>())?;
    Some(((abs, col), j))
}

/// `$?DIGITS` at `i`, for a whole-row reference.
fn parse_row_part(chars: &[char], i: usize) -> Option<((bool, u32), usize)> {
    let mut j = i;
    let abs = chars.get(j) == Some(&'$');
    if abs {
        j += 1;
    }
    let ds = j;
    while j < chars.len() && chars[j].is_ascii_digit() {
        j += 1;
    }
    if j == ds || j - ds > 7 {
        return None;
    }
    let row: u32 = chars[ds..j].iter().collect::<String>().parse().ok()?;
    if row == 0 || row > MAX_ROW {
        return None;
    }
    Some(((abs, row), j))
}

/// The token must end where an identifier would end, and not at a `(`.
fn ends_token(chars: &[char], j: usize) -> bool {
    match chars.get(j) {
        None => true,
        Some(c) => !is_ident_char(*c) && *c != '(' && *c != '$',
    }
}

/// A reference token at `i`, with where it ends.
fn parse_token(chars: &[char], i: usize) -> Option<(Token, usize)> {
    if let Some((a, j)) = parse_cell(chars, i) {
        if chars.get(j) == Some(&':') {
            if let Some((b, k)) = parse_cell(chars, j + 1) {
                if ends_token(chars, k) {
                    return Some((Token::Range(a, b), k));
                }
            }
        }
        if ends_token(chars, j) {
            return Some((Token::Cell(a), j));
        }
        return None;
    }
    if let Some((a, j)) = parse_col_part(chars, i) {
        if chars.get(j) == Some(&':') {
            if let Some((b, k)) = parse_col_part(chars, j + 1) {
                if ends_token(chars, k) {
                    return Some((Token::Columns(a, b), k));
                }
            }
        }
        return None;
    }
    if let Some((a, j)) = parse_row_part(chars, i) {
        if chars.get(j) == Some(&':') {
            if let Some((b, k)) = parse_row_part(chars, j + 1) {
                if ends_token(chars, k) {
                    return Some((Token::Rows(a, b), k));
                }
            }
        }
    }
    None
}

fn shifted(n: u32, abs: bool, d: i64, max: u32) -> Option<u32> {
    if abs {
        return Some(n);
    }
    let m = i64::from(n) + d;
    (m >= 1 && m <= i64::from(max)).then_some(m as u32)
}

fn render_cell(c: &CellRef, d_row: i64, d_col: i64) -> Option<String> {
    let col = shifted(c.col, c.col_abs, d_col, MAX_COLUMN)?;
    let row = shifted(c.row, c.row_abs, d_row, MAX_ROW)?;
    Some(format!(
        "{}{}{}{}",
        if c.col_abs { "$" } else { "" },
        column_letters(col),
        if c.row_abs { "$" } else { "" },
        row
    ))
}

fn render_token(t: &Token, d_row: i64, d_col: i64) -> String {
    let rendered = match t {
        Token::Cell(c) => render_cell(c, d_row, d_col),
        Token::Range(a, b) => match (render_cell(a, d_row, d_col), render_cell(b, d_row, d_col)) {
            (Some(x), Some(y)) => Some(format!("{x}:{y}")),
            _ => None,
        },
        Token::Columns(a, b) => {
            match (
                shifted(a.1, a.0, d_col, MAX_COLUMN),
                shifted(b.1, b.0, d_col, MAX_COLUMN),
            ) {
                (Some(x), Some(y)) => Some(format!(
                    "{}{}:{}{}",
                    if a.0 { "$" } else { "" },
                    column_letters(x),
                    if b.0 { "$" } else { "" },
                    column_letters(y)
                )),
                _ => None,
            }
        }
        Token::Rows(a, b) => match (
            shifted(a.1, a.0, d_row, MAX_ROW),
            shifted(b.1, b.0, d_row, MAX_ROW),
        ) {
            (Some(x), Some(y)) => Some(format!(
                "{}{}:{}{}",
                if a.0 { "$" } else { "" },
                x,
                if b.0 { "$" } else { "" },
                y
            )),
            _ => None,
        },
    };
    rendered.unwrap_or_else(|| "#REF!".to_string())
}

/// The formula's text (without its `=`) with every relative reference
/// moved `d_row` rows down and `d_col` columns right, as the formula of a
/// cell that far from the first cell of a fill.
pub fn shift_a1_references(formula: &str, d_row: i64, d_col: i64) -> String {
    if d_row == 0 && d_col == 0 {
        return formula.to_string();
    }
    let chars: Vec<char> = formula.chars().collect();
    let mut out = String::with_capacity(formula.len() + 8);
    let mut i = 0;
    while i < chars.len() {
        let c = chars[i];
        if c == '"' || c == '\'' {
            // A string literal, or a quoted sheet name; a doubled quote stays inside.
            out.push(c);
            i += 1;
            while i < chars.len() {
                out.push(chars[i]);
                if chars[i] == c {
                    if i + 1 < chars.len() && chars[i + 1] == c {
                        out.push(c);
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
        if c == '[' {
            // A structured reference, to its closing bracket.
            while i < chars.len() {
                out.push(chars[i]);
                i += 1;
                if chars[i - 1] == ']' {
                    break;
                }
            }
            continue;
        }
        let at_boundary = i == 0 || !is_ident_char(chars[i - 1]) && chars[i - 1] != '$';
        if at_boundary && (c == '$' || c.is_ascii_alphanumeric()) {
            if let Some((token, end)) = parse_token(&chars, i) {
                out.push_str(&render_token(&token, d_row, d_col));
                i = end;
                continue;
            }
            // Not a reference: copy the whole identifier so its digits are
            // never taken for a row reference midway.
            let start = i;
            while i < chars.len() && (is_ident_char(chars[i]) || chars[i] == '$') {
                i += 1;
            }
            if i == start {
                out.push(c);
                i += 1;
            } else {
                out.extend(chars[start..i].iter());
            }
            continue;
        }
        out.push(c);
        i += 1;
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;

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
    }

    #[test]
    fn off_the_sheet_is_ref_error() {
        assert_eq!(s("A1", -1, 0), "#REF!");
        assert_eq!(s("A1+B1", 0, -1), "#REF!+A1");
        assert_eq!(s("XFD1", 0, 1), "#REF!");
        assert_eq!(s("A1:B2", -1, 0), "#REF!");
        assert_eq!(s("$A$1", -5, -5), "$A$1");
    }
}
