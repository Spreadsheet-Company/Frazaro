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
//! `cell` row, then its `formula` row, then its `refers` rows, one per
//! distinct target, sorted by the target as printed (slice 8b, through
//! `crate::refers`, the port of `VLA_Refers.bas`). `print.rs` has the
//! spelling; `cone.rs` sizes a cell's cone through an index of the walk.

pub mod cone;
pub mod cursor;
pub mod ooxml;
pub mod print;

use crate::messages::Refusal;
use crate::refers::{FormulaRef, Kind};

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

/// The second field of a `refers` row: a reference in the treaty's
/// spelling, or an unreadable call named.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum RefersTo {
    /// `Model!B1`, `Data!A:A`, `Rate`, `Sales[Amount]`, `[Rates.xlsx]Sheet1!A1`.
    Reference(String),
    /// `(unreadable "INDIRECT")`, `(unreadable "OFFSET")`.
    Unreadable(&'static str),
}

impl RefersTo {
    /// What one record of a formula's scan refers to, from the sheet that
    /// holds the formula.
    pub fn of(r: &FormulaRef, home_sheet: &str) -> RefersTo {
        match r.kind {
            Kind::Unreadable => RefersTo::Unreadable(r.unreadable_name()),
            _ => RefersTo::Reference(crate::refers::spell(r, home_sheet)),
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
    /// `(refers "Model!B3" "Model!B1")`: the cell holding the formula, and
    /// one thing it refers to.
    Refers {
        sheet: &'r str,
        addr: &'r str,
        to: &'r RefersTo,
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
/// the part's inflated size, and from 8b the distinct formulas in R1C1 and
/// the `INDIRECT` and `OFFSET` calls met.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct SheetStats {
    pub cells: u64,
    pub formulas: u64,
    pub array_anchors: u64,
    pub rows: u32,
    pub columns: u32,
    pub part_bytes: usize,
    /// Distinct formulas on the sheet, compared in R1C1 relative to their
    /// cells, so that a formula filled down counts once.
    pub distinct_r1c1: u64,
    /// `INDIRECT` and `OFFSET` calls on the sheet: where a cone is blind.
    pub unreadable: u64,
}

/// A sheet's name as a reference spells it in front of `!`: the one quoting
/// rule, `refers::quote_sheet` (the port of `RefersQuoteSheet`), so that a
/// `name` row's sheet, a `refers` row's two ends and the formula text agree.
pub fn sheet_prefix(name: &str) -> String {
    crate::refers::quote_sheet(name)
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
        // AXM.7's two clauses, confirmed in Excel on 2026-10-04.
        assert_eq!(sheet_prefix("A1"), "'A1'");
        assert_eq!(sheet_prefix("R1C1"), "'R1C1'");
    }

    #[test]
    fn a_target_is_a_reference_or_an_unreadable_call() {
        let refs = crate::refers::scan("=OFFSET(B1,1,0)+'Q1 Data'!A1");
        assert_eq!(
            RefersTo::of(&refs[0], "Model"),
            RefersTo::Unreadable("OFFSET")
        );
        assert_eq!(
            RefersTo::of(&refs[1], "Model"),
            RefersTo::Reference("Model!B1".to_string())
        );
        assert_eq!(
            RefersTo::of(&refs[2], "Model"),
            RefersTo::Reference("'Q1 Data'!A1".to_string())
        );
    }
}
