//! The relation rows and their spelling (PORT.8, slice 8a; cut out of
//! `frazaro-core`'s `reflect` with the language, PORT.12): one row a line,
//! in the proof corpus's notation, so that an expected-relations file is a
//! golden the core writes, the owner checks in Excel by eye, and
//! `VlaReadForms` reads later for the VBA's own `REFLECT` test without a new
//! parser. The reader of a workbook's file prints these rows from the file;
//! the view record ([`crate::view`]) prints the `sheet`, `cell` and
//! `formula` rows from the grid; an engine's loader reads them back as the
//! inverse.
//!
//! The spelling, fixed by the treaty's amendment for oracle 8: a text is a
//! VLA string as `WriteDatum` writes one; a number is the file's own text,
//! never a float round trip; a truth value is `true` or `false`; an error is
//! `(error "#DIV/0!")`; an ISO date cell is `(date "2026-10-03")`; a sheet's
//! state is `visible`, `hidden` or `very-hidden`.

use crate::form::Form;
use crate::printer::write_datum;
use crate::refers::{self, FormulaRef, Kind};

/// A cell's value as the file holds it.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum Value {
    /// The file's own text of a number, never a float round trip.
    Number(String),
    Text(String),
    Bool(bool),
    /// An error value's text, `#DIV/0!`.
    Error(String),
    /// An ISO 8601 date cell (`t="d"`; an OpenDocument date or time value),
    /// as written.
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
            _ => RefersTo::Reference(refers::spell(r, home_sheet)),
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

/// Where rows go as a walk emits them.
pub trait Sink {
    fn row(&mut self, row: &Row<'_>);
}

/// A sheet's name as a reference spells it in front of `!`: the one quoting
/// rule, `refers::quote_sheet` (the port of `RefersQuoteSheet`), so that a
/// `name` row's sheet, a `refers` row's two ends and the formula text agree.
pub fn sheet_prefix(name: &str) -> String {
    refers::quote_sheet(name)
}

/// A text as a VLA string literal, as `WriteDatum` writes one.
pub fn quoted(s: &str) -> String {
    write_datum(&Form::string(s))
}

/// A `refers` row's second field as printed: a quoted reference, or the
/// `(unreadable "...")` form. The rows of one cell sort by this text.
pub fn target(to: &RefersTo) -> String {
    match to {
        RefersTo::Reference(text) => quoted(text),
        RefersTo::Unreadable(name) => format!("(unreadable {})", quoted(name)),
    }
}

/// A value as a datum.
pub fn datum(v: &Value) -> String {
    match v {
        Value::Number(text) => text.clone(),
        Value::Text(text) => quoted(text),
        Value::Bool(true) => "true".to_string(),
        Value::Bool(false) => "false".to_string(),
        Value::Error(text) => format!("(error {})", quoted(text)),
        Value::Date(text) => format!("(date {})", quoted(text)),
    }
}

/// One row as its line, without the line break.
pub fn line(row: &Row<'_>) -> String {
    match row {
        Row::Sheet { name, visibility } => {
            format!("(sheet {} {})", quoted(name), visibility.word())
        }
        Row::Name { name, refers_to } => {
            format!("(name {} {})", quoted(name), quoted(refers_to))
        }
        Row::Table { name, sheet, range } => {
            format!(
                "(table {} {} {})",
                quoted(name),
                quoted(sheet),
                quoted(range)
            )
        }
        Row::Cell { sheet, addr, value } => {
            format!("(cell {} {} {})", quoted(sheet), quoted(addr), datum(value))
        }
        Row::Formula { sheet, addr, text } => {
            format!(
                "(formula {} {} {})",
                quoted(sheet),
                quoted(addr),
                quoted(text)
            )
        }
        Row::Refers { sheet, addr, to } => {
            format!(
                "(refers {} {})",
                quoted(&format!("{}!{}", sheet_prefix(sheet), addr)),
                target(to)
            )
        }
    }
}

/// A sink that holds the lines as one text, each ending in a line feed.
#[derive(Default)]
pub struct Printer {
    pub out: String,
}

impl Sink for Printer {
    fn row(&mut self, row: &Row<'_>) {
        self.out.push_str(&line(row));
        self.out.push('\n');
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn each_row_prints_in_the_corpus_notation() {
        assert_eq!(
            line(&Row::Sheet {
                name: "Q1 Data",
                visibility: Visibility::VeryHidden
            }),
            "(sheet \"Q1 Data\" very-hidden)"
        );
        assert_eq!(
            line(&Row::Name {
                name: "Rate",
                refers_to: "Model!$B$1"
            }),
            "(name \"Rate\" \"Model!$B$1\")"
        );
        assert_eq!(
            line(&Row::Table {
                name: "Sales",
                sheet: "Data",
                range: "A1:B4"
            }),
            "(table \"Sales\" \"Data\" \"A1:B4\")"
        );
        assert_eq!(
            line(&Row::Cell {
                sheet: "Model",
                addr: "B1",
                value: &Value::Number("1200".to_string())
            }),
            "(cell \"Model\" \"B1\" 1200)"
        );
        assert_eq!(
            line(&Row::Cell {
                sheet: "Model",
                addr: "A1",
                value: &Value::Text("say \"hi\"".to_string())
            }),
            "(cell \"Model\" \"A1\" \"say \\\"hi\\\"\")"
        );
        assert_eq!(datum(&Value::Bool(true)), "true");
        assert_eq!(datum(&Value::Bool(false)), "false");
        assert_eq!(
            datum(&Value::Error("#DIV/0!".to_string())),
            "(error \"#DIV/0!\")"
        );
        assert_eq!(
            datum(&Value::Date("2026-10-03".to_string())),
            "(date \"2026-10-03\")"
        );
        assert_eq!(
            line(&Row::Formula {
                sheet: "Model",
                addr: "B3",
                text: "=B1-B2"
            }),
            "(formula \"Model\" \"B3\" \"=B1-B2\")"
        );
        assert_eq!(
            line(&Row::Refers {
                sheet: "Q1 Data",
                addr: "B3",
                to: &RefersTo::Reference("Model!B1".to_string())
            }),
            "(refers \"'Q1 Data'!B3\" \"Model!B1\")"
        );
        assert_eq!(
            line(&Row::Refers {
                sheet: "Model",
                addr: "D2",
                to: &RefersTo::Unreadable("INDIRECT")
            }),
            "(refers \"Model!D2\" (unreadable \"INDIRECT\"))"
        );
        // A quoted reference sorts before the unreadable form: '"' < '('.
        assert!(
            target(&RefersTo::Reference("Model!B1".to_string()))
                < target(&RefersTo::Unreadable("OFFSET"))
        );
        let mut p = Printer::default();
        p.row(&Row::Sheet {
            name: "A",
            visibility: Visibility::Visible,
        });
        p.row(&Row::Sheet {
            name: "B",
            visibility: Visibility::Hidden,
        });
        assert_eq!(p.out, "(sheet \"A\" visible)\n(sheet \"B\" hidden)\n");
    }

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
        let refs = refers::scan("=OFFSET(B1,1,0)+'Q1 Data'!A1");
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
