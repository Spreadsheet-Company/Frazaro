//! The sheet model (PORT.7): what a workbook holds before it is written.
//! Sheets of cells, each cell a text, a number, a truth value or a formula's
//! text; column widths, a hidden column, a style per cell or per column from
//! a small style table; and nothing that computes. docs/HORIZON.md section
//! 12: the core has no calc engine, so a formula is text the host evaluates
//! when it opens the file, which `fullCalcOnLoad` (ooxml.rs) makes it do.
//!
//! New ground, said plainly: the reference has no model of a sheet; it
//! writes through Excel's object model. What is modelled is the little the
//! room needs (`BuildWorkspace`, VLA_IDE.bas) and what slice 7b's static
//! subset writes, and the model grows only as a slice needs it.

pub mod ooxml;
pub mod xlfn;
pub mod xml;
pub mod zip;

use std::collections::BTreeMap;

use crate::intrinsics::fold;

/// A colour, as `Interior.Color = RGB(r, g, b)` names it.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Rgb(pub u8, pub u8, pub u8);

impl Rgb {
    /// The file's spelling: opaque alpha, then the three bytes in hex.
    pub fn argb_hex(self) -> String {
        format!("FF{:02X}{:02X}{:02X}", self.0, self.1, self.2)
    }
}

/// The number formats a style names: Excel's built-in General and Text
/// (`NumberFormat = "@"`, so a sentence is never read as a formula).
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum NumFmt {
    #[default]
    General,
    Text,
}

impl NumFmt {
    /// The built-in format's id in the styles part.
    pub fn id(self) -> u32 {
        match self {
            NumFmt::General => 0,
            NumFmt::Text => 49,
        }
    }
}

/// One cell format: a solid fill or none, a number format, wrapped text or
/// not. Index 0 of a [`Styles`] table is this type's default.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Style {
    pub fill: Option<Rgb>,
    pub num_fmt: NumFmt,
    pub wrap: bool,
}

/// The workbook's cell formats, each once; a cell or a column names one by
/// its index.
#[derive(Clone, Debug)]
pub struct Styles {
    xfs: Vec<Style>,
}

impl Default for Styles {
    fn default() -> Self {
        Self::new()
    }
}

impl Styles {
    /// A table holding only the default format, at index 0.
    pub fn new() -> Styles {
        Styles {
            xfs: vec![Style::default()],
        }
    }

    /// The index of a format, added if it is new.
    pub fn id(&mut self, style: Style) -> u32 {
        if let Some(i) = self.xfs.iter().position(|s| *s == style) {
            return i as u32;
        }
        self.xfs.push(style);
        (self.xfs.len() - 1) as u32
    }

    /// Every format, in index order.
    pub fn xfs(&self) -> &[Style] {
        &self.xfs
    }

    /// The distinct fills, in the order they were first used.
    pub fn fills(&self) -> Vec<Rgb> {
        let mut fills: Vec<Rgb> = Vec::new();
        for s in &self.xfs {
            if let Some(f) = s.fill {
                if !fills.contains(&f) {
                    fills.push(f);
                }
            }
        }
        fills
    }
}

/// What a cell holds. A formula is its text without the leading `=`, as the
/// file stores it; the host computes it. A formula written into a range of
/// cells is a shared formula, which is what Excel itself writes when a
/// formula is filled: the top-left cell is the master and holds the text
/// and the range, the others point at it by its index, and the host
/// adjusts the references for each cell as `Range.Formula2` would have.
#[derive(Clone, Debug, PartialEq)]
pub enum Content {
    Text(String),
    Number(f64),
    Bool(bool),
    Formula(String),
    SharedMaster {
        text: String,
        range: String,
        si: u32,
    },
    SharedChild {
        si: u32,
    },
}

#[derive(Clone, Debug, PartialEq)]
pub struct Cell {
    pub content: Content,
    pub style: u32,
}

/// A column's settings: `index` is 1-based; `width` is in characters, as
/// `ColumnWidth` counts them; `style` is the format of its cells that have
/// none of their own.
#[derive(Clone, Debug, PartialEq)]
pub struct Column {
    pub index: u32,
    pub width: Option<f64>,
    pub hidden: bool,
    pub style: Option<u32>,
}

/// A rectangle of cells, rows and columns 1-based and inclusive.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct A1Range {
    pub top: u32,
    pub left: u32,
    pub bottom: u32,
    pub right: u32,
}

impl A1Range {
    pub fn cells(&self) -> u64 {
        u64::from(self.bottom - self.top + 1) * u64::from(self.right - self.left + 1)
    }

    /// `B2` for one cell, `B2:C4` for more.
    pub fn text(&self) -> String {
        if self.top == self.bottom && self.left == self.right {
            cell_ref(self.top, self.left)
        } else {
            format!(
                "{}:{}",
                cell_ref(self.top, self.left),
                cell_ref(self.bottom, self.right)
            )
        }
    }
}

/// Excel's last column (XFD) and last row.
pub const MAX_COLUMN: u32 = 16_384;
pub const MAX_ROW: u32 = 1_048_576;

fn parse_a1_cell(s: &str) -> Option<(u32, u32)> {
    let s = s.replace('$', "");
    let letters: String = s.chars().take_while(|c| c.is_ascii_alphabetic()).collect();
    let digits = &s[letters.len()..];
    if letters.is_empty() || letters.len() > 3 || digits.is_empty() {
        return None;
    }
    if !digits.chars().all(|c| c.is_ascii_digit()) {
        return None;
    }
    let mut col: u32 = 0;
    for c in letters.chars() {
        col = col * 26 + (c.to_ascii_uppercase() as u32 - 'A' as u32 + 1);
    }
    let row: u32 = digits.parse().ok()?;
    if col == 0 || col > MAX_COLUMN || row == 0 || row > MAX_ROW {
        return None;
    }
    Some((row, col))
}

/// An A1 reference as a program writes it (`b2`, `A1:B5`, `$C$2`), in any
/// case, with the corners in any order; `None` for anything else, a whole
/// column or row included.
pub fn parse_a1_range(s: &str) -> Option<A1Range> {
    let s = s.trim();
    let (a, b) = match s.split_once(':') {
        Some((a, b)) => (a, b),
        None => (s, s),
    };
    let (r1, c1) = parse_a1_cell(a)?;
    let (r2, c2) = parse_a1_cell(b)?;
    Some(A1Range {
        top: r1.min(r2),
        left: c1.min(c2),
        bottom: r1.max(r2),
        right: c1.max(c2),
    })
}

/// One sheet: its cells by (row, column), both 1-based.
#[derive(Clone, Debug, PartialEq)]
pub struct Sheet {
    pub name: String,
    pub cells: BTreeMap<(u32, u32), Cell>,
    pub columns: Vec<Column>,
    pub gridlines: bool,
    /// The selected cell when the sheet opens, (row, column).
    pub active_cell: (u32, u32),
    /// How many shared formulas the sheet holds; the next one's index.
    pub shared_formulas: u32,
}

impl Sheet {
    pub fn new(name: &str) -> Sheet {
        Sheet {
            name: name.to_string(),
            cells: BTreeMap::new(),
            columns: Vec::new(),
            gridlines: true,
            active_cell: (1, 1),
            shared_formulas: 0,
        }
    }

    pub fn set(&mut self, row: u32, col: u32, cell: Cell) {
        self.cells.insert((row, col), cell);
    }

    /// A formula's text (without its `=`) into every cell of a range: one
    /// plain formula for one cell, a shared formula for more.
    pub fn set_formula(&mut self, range: A1Range, text: &str, style: u32) {
        if range.cells() == 1 {
            self.set(
                range.top,
                range.left,
                Cell {
                    content: Content::Formula(text.to_string()),
                    style,
                },
            );
            return;
        }
        let si = self.shared_formulas;
        self.shared_formulas += 1;
        for row in range.top..=range.bottom {
            for col in range.left..=range.right {
                let content = if row == range.top && col == range.left {
                    Content::SharedMaster {
                        text: text.to_string(),
                        range: range.text(),
                        si,
                    }
                } else {
                    Content::SharedChild { si }
                };
                self.set(row, col, Cell { content, style });
            }
        }
    }

    /// The used range as ((first row, first column), (last row, last
    /// column)), or `None` for a sheet with no cell.
    pub fn extent(&self) -> Option<((u32, u32), (u32, u32))> {
        let mut it = self.cells.keys();
        let &(r0, c0) = it.next()?;
        let (mut top, mut left, mut bottom, mut right) = (r0, c0, r0, c0);
        for &(r, c) in it {
            top = top.min(r);
            left = left.min(c);
            bottom = bottom.max(r);
            right = right.max(c);
        }
        Some(((top, left), (bottom, right)))
    }
}

/// The workbook: its sheets in tab order, their formats, and which tab is
/// open when the file opens.
#[derive(Clone, Debug)]
pub struct Workbook {
    pub sheets: Vec<Sheet>,
    pub styles: Styles,
    pub active_sheet: usize,
}

impl Default for Workbook {
    fn default() -> Self {
        Self::new()
    }
}

impl Workbook {
    pub fn new() -> Workbook {
        Workbook {
            sheets: Vec::new(),
            styles: Styles::new(),
            active_sheet: 0,
        }
    }

    /// The index of the sheet of that name, compared as Excel compares sheet
    /// names, without case.
    pub fn find_sheet(&self, name: &str) -> Option<usize> {
        let want = fold(name);
        self.sheets.iter().position(|s| fold(&s.name) == want)
    }

    /// The sheet of that name, added last with the name as given when there
    /// is none (`VlaEnsureSheet`'s shape).
    pub fn ensure_sheet(&mut self, name: &str) -> usize {
        if let Some(i) = self.find_sheet(name) {
            return i;
        }
        self.sheets.push(Sheet::new(name));
        self.sheets.len() - 1
    }
}

/// A column's letters: 1 is A, 26 is Z, 27 is AA.
pub fn column_letters(col: u32) -> String {
    let mut n = col;
    let mut letters = Vec::new();
    while n > 0 {
        let rem = (n - 1) % 26;
        letters.push((b'A' + rem as u8) as char);
        n = (n - 1) / 26;
    }
    letters.iter().rev().collect()
}

/// A cell's reference: row 1, column 2 is `B1`.
pub fn cell_ref(row: u32, col: u32) -> String {
    format!("{}{}", column_letters(col), row)
}

/// `ColumnWidth` in characters to the file's `width` attribute, as Excel
/// computes it for the default font, Calibri 11, whose widest digit is 7
/// pixels: the characters plus 5 pixels of padding, in 1/256ths of a
/// character. 72 gives 72.7109375, which is what Excel itself writes for a
/// column set to 72.
pub fn column_width_attr(chars: f64) -> f64 {
    const MAX_DIGIT_WIDTH: f64 = 7.0;
    ((chars * MAX_DIGIT_WIDTH + 5.0) / MAX_DIGIT_WIDTH * 256.0).trunc() / 256.0
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn references_and_letters() {
        assert_eq!(column_letters(1), "A");
        assert_eq!(column_letters(26), "Z");
        assert_eq!(column_letters(27), "AA");
        assert_eq!(column_letters(702), "ZZ");
        assert_eq!(column_letters(703), "AAA");
        assert_eq!(cell_ref(1, 2), "B1");
        assert_eq!(cell_ref(12, 3), "C12");
    }

    #[test]
    fn the_room_widths_as_excel_writes_them() {
        assert_eq!(column_width_attr(72.0), 72.7109375);
        assert_eq!(column_width_attr(60.0), 60.7109375);
        assert_eq!(column_width_attr(8.43), 9.140625);
    }

    #[test]
    fn styles_dedup_and_fills_keep_first_use_order() {
        let mut s = Styles::new();
        let a = s.id(Style {
            fill: Some(Rgb(1, 2, 3)),
            ..Style::default()
        });
        let b = s.id(Style {
            fill: Some(Rgb(4, 5, 6)),
            wrap: true,
            ..Style::default()
        });
        let a2 = s.id(Style {
            fill: Some(Rgb(1, 2, 3)),
            ..Style::default()
        });
        assert_eq!((a, b, a2), (1, 2, 1));
        assert_eq!(s.fills(), vec![Rgb(1, 2, 3), Rgb(4, 5, 6)]);
        assert_eq!(Rgb(247, 244, 252).argb_hex(), "FFF7F4FC");
    }

    #[test]
    fn a1_references_as_a_program_writes_them() {
        let r = parse_a1_range("b2").unwrap();
        assert_eq!((r.top, r.left, r.bottom, r.right), (2, 2, 2, 2));
        assert_eq!(r.text(), "B2");
        let r = parse_a1_range("c4:A1").unwrap();
        assert_eq!((r.top, r.left, r.bottom, r.right), (1, 1, 4, 3));
        assert_eq!(r.text(), "A1:C4");
        assert_eq!(r.cells(), 12);
        assert_eq!(parse_a1_range("$B$2:$B$4").unwrap().text(), "B2:B4");
        assert_eq!(parse_a1_range("XFD1048576").unwrap().text(), "XFD1048576");
        for bad in [
            "c:c", "1:3", "XFE1", "A0", "A1048577", "AAAA1", "b", "2", "", "b2:",
        ] {
            assert!(parse_a1_range(bad).is_none(), "{bad} should not parse");
        }
    }

    #[test]
    fn a_formula_over_a_range_is_shared_from_its_top_left_cell() {
        let mut sh = Sheet::new("S");
        sh.set_formula(parse_a1_range("c2:c3").unwrap(), "B2+B3", 0);
        sh.set_formula(parse_a1_range("d2").unwrap(), "B2*2", 0);
        sh.set_formula(parse_a1_range("e2:e3").unwrap(), "1", 0);
        assert_eq!(
            sh.cells[&(2, 3)].content,
            Content::SharedMaster {
                text: "B2+B3".to_string(),
                range: "C2:C3".to_string(),
                si: 0
            }
        );
        assert_eq!(sh.cells[&(3, 3)].content, Content::SharedChild { si: 0 });
        assert_eq!(
            sh.cells[&(2, 4)].content,
            Content::Formula("B2*2".to_string())
        );
        assert_eq!(
            sh.cells[&(2, 5)].content,
            Content::SharedMaster {
                text: "1".to_string(),
                range: "E2:E3".to_string(),
                si: 1
            }
        );
        assert_eq!(sh.shared_formulas, 2);
    }

    #[test]
    fn sheets_are_found_without_case_and_made_once() {
        let mut wb = Workbook::new();
        assert_eq!(wb.find_sheet("Output"), None);
        let i = wb.ensure_sheet("Output");
        assert_eq!(wb.ensure_sheet("output"), i);
        assert_eq!(wb.find_sheet("OUTPUT"), Some(i));
        assert_eq!(wb.sheets.len(), 1);
        assert_eq!(wb.sheets[0].name, "Output");
    }

    #[test]
    fn a_sheet_knows_its_extent() {
        let mut sh = Sheet::new("S");
        assert_eq!(sh.extent(), None);
        sh.set(
            3,
            2,
            Cell {
                content: Content::Number(1.0),
                style: 0,
            },
        );
        sh.set(
            1,
            3,
            Cell {
                content: Content::Bool(true),
                style: 0,
            },
        );
        assert_eq!(sh.extent(), Some(((1, 2), (3, 3))));
    }
}
