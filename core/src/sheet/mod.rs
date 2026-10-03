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
pub mod xml;
pub mod zip;

use std::collections::BTreeMap;

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
/// file stores it; the host computes it.
#[derive(Clone, Debug, PartialEq)]
pub enum Content {
    Text(String),
    Number(f64),
    Bool(bool),
    Formula(String),
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

/// One sheet: its cells by (row, column), both 1-based.
#[derive(Clone, Debug, PartialEq)]
pub struct Sheet {
    pub name: String,
    pub cells: BTreeMap<(u32, u32), Cell>,
    pub columns: Vec<Column>,
    pub gridlines: bool,
    /// The selected cell when the sheet opens, (row, column).
    pub active_cell: (u32, u32),
}

impl Sheet {
    pub fn new(name: &str) -> Sheet {
        Sheet {
            name: name.to_string(),
            cells: BTreeMap::new(),
            columns: Vec::new(),
            gridlines: true,
            active_cell: (1, 1),
        }
    }

    pub fn set(&mut self, row: u32, col: u32, cell: Cell) {
        self.cells.insert((row, col), cell);
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
