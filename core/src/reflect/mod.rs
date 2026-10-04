//! The reader (PORT.8): a workbook's file read as the relations of
//! `REFLECT` (docs/SINGULARITY.md, Stage 0.1; the roadmap's `AXM.8`), with
//! no host on the machine.
//!
//! New ground, said plainly: the reference reads a workbook through Excel's
//! object model and never opens the file, so there is no VBA arm to port
//! for the file formats; what the relations are, and how a formula's
//! references are read (`AXM.7`, slice 8b), is the language's and follows
//! the VBA as everything else does. A reader prints seven of the eight:
//! `sheet(name, state)`, `name(name, refersto)`, `table(name, sheet,
//! range)`, `cell(sheet, addr, value)`, `formula(sheet, addr, text)` and,
//! from 8b, `refers(from, to)`; `changed` is `diff`'s (8c) and `ran` comes
//! from no file.
//!
//! The reader streams. It never builds a model of the workbook: each sheet
//! part is inflated whole (the bound every part already has, 256 MB), walked
//! once by the cursor in `cursor.rs`, and dropped before the next is read;
//! the walk hands each relation row to a [`Sink`] as it finds it. What is
//! held is the shared-string table, since any cell may name any string, and
//! the sink's own state: nothing for printing, counters for `--counts`.
//!
//! The fixed order, which is what makes the printed text a golden: every
//! `sheet` row in tab order; every `name` row sorted by its name without
//! case; every `table` row by its sheet's tab order, then its name; then
//! sheet by sheet in tab order, cell by cell in document order, each cell's
//! `cell` row, then its `formula` row. `print.rs` has the spelling.

pub mod cursor;
pub mod ooxml;
pub mod print;

use crate::messages::Refusal;

/// A cell's value as the file holds it.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum Value {
    /// The file's own text of a number, never a float round trip.
    Number(String),
    Text(String),
    Bool(bool),
    /// An error value's text, `#DIV/0!`.
    Error(String),
    /// An ISO 8601 date cell (`t="d"`), as written.
    Date(String),
}

/// A sheet's state, as the workbook part has it.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Visibility {
    Visible,
    Hidden,
    /// `veryHidden`: not listed by Excel's Unhide dialog.
    VeryHidden,
}

impl Visibility {
    /// The word a `sheet` row prints.
    pub fn word(self) -> &'static str {
        match self {
            Visibility::Visible => "visible",
            Visibility::Hidden => "hidden",
            Visibility::VeryHidden => "very-hidden",
        }
    }
}

/// One row of a relation, borrowed from the walk that found it.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Row<'r> {
    Sheet {
        name: &'r str,
        visibility: Visibility,
    },
    Name {
        name: &'r str,
        refers_to: &'r str,
    },
    Table {
        name: &'r str,
        sheet: &'r str,
        range: &'r str,
    },
    Cell {
        sheet: &'r str,
        addr: &'r str,
        value: &'r Value,
    },
    Formula {
        sheet: &'r str,
        addr: &'r str,
        text: &'r str,
    },
}

/// Where rows go as the walk emits them.
pub trait Sink {
    fn row(&mut self, row: &Row<'_>);
}

/// A sink that keeps nothing, for `--counts`.
pub struct Discard;

impl Sink for Discard {
    fn row(&mut self, _row: &Row<'_>) {}
}

/// What one sheet's walk counted: cells with a value or a formula, formula
/// cells, array-formula anchors, the last row and column that hold one,
/// and the part's inflated size.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct SheetStats {
    pub cells: u64,
    pub formulas: u64,
    pub array_anchors: u64,
    pub rows: u32,
    pub columns: u32,
    pub part_bytes: usize,
}

/// A sheet's name as a reference spells it in front of `!`: quoted, with an
/// apostrophe doubled, when it holds anything but letters, digits,
/// underscores and periods, or begins with a digit.
pub fn sheet_prefix(name: &str) -> String {
    let plain = !name.is_empty()
        && !name.starts_with(|c: char| c.is_ascii_digit())
        && name
            .chars()
            .all(|c| c.is_alphanumeric() || c == '_' || c == '.');
    if plain {
        name.to_string()
    } else {
        format!("'{}'", name.replace('\'', "''"))
    }
}

/// The whole of a workbook as text, one relation row a line in the fixed
/// order: the API's surface, the tests' and a door's that holds the text.
pub fn reflect_text(bytes: &[u8], label: &str) -> Result<String, Refusal> {
    let package = ooxml::Package::open(bytes, label)?;
    let mut printer = print::Printer::default();
    package.header_rows(&mut printer);
    for i in 0..package.sheets().len() {
        package.walk_sheet(i, &mut printer)?;
    }
    Ok(printer.out)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn a_sheet_name_is_quoted_as_a_reference_quotes_it() {
        assert_eq!(sheet_prefix("Model"), "Model");
        assert_eq!(sheet_prefix("Sheet.1"), "Sheet.1");
        assert_eq!(sheet_prefix("Q1 Data"), "'Q1 Data'");
        assert_eq!(sheet_prefix("It's"), "'It''s'");
        assert_eq!(sheet_prefix("2026"), "'2026'");
        assert_eq!(sheet_prefix("Donn\u{e9}es"), "Donn\u{e9}es");
        assert_eq!(sheet_prefix("a-b"), "'a-b'");
    }
}
