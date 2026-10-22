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
//!
//! The record's own rows (KERNEL.57): a sheet's gridlines, a column's and a
//! row's settings, a format, a formula's value and a cell's format, which the
//! relations do not print, are spelled here too ([`Record`], [`record_line`]),
//! so that the view record and a reader printing a workbook as the record
//! prints it (`frazaro reflect --looks`) print them alike.

use crate::form::Form;
use crate::printer::write_datum;
use crate::refers::{self, FormulaRef, Kind};
use crate::sheet::{column_letters, number_text, Column, NumFmt, RowSettings, Style};

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

/// One of the record's own rows (KERNEL.57), which the relations do not
/// print: the view record prints them (`crate::view`), and a reader printing
/// a workbook as the record prints a sheet whole hands them to its sink's
/// [`Sink::record`], so that both spell them by [`record_line`]. A list that
/// may grow.
#[derive(Clone, Copy, Debug, PartialEq)]
#[non_exhaustive]
pub enum Record<'r> {
    /// `(gridlines "<sheet>" on|off)`.
    Gridlines { sheet: &'r str, on: bool },
    /// `(column "<sheet>" "<letters>" <width>|none shown|hidden <format>|none)`.
    Column { sheet: &'r str, column: &'r Column },
    /// `(row "<sheet>" <number> <height>|none shown|hidden)`.
    Row {
        sheet: &'r str,
        number: u32,
        settings: &'r RowSettings,
    },
    /// `(format <index> <fill> <number format> wrap|nowrap <groups>)`: the
    /// format at `index`, the default format where `style` is `None`; and
    /// `code`, where given, the number format as a file holds it, printed in
    /// quotes in place of the style's own, a code the language may not draw
    /// (KERNEL.57: a reader prints it as written, and a load refuses it by
    /// name).
    Format {
        index: u32,
        style: Option<&'r Style>,
        code: Option<&'r str>,
    },
    /// `(value "<sheet>" "<addr>" <value>)`: a formula's value, spelled as a
    /// `cell` row spells one, or the label of a formula not computed.
    Value {
        sheet: &'r str,
        addr: &'r str,
        value: &'r str,
    },
    /// `(style "<sheet>" "<addr>" <format>)`: the cell's format, where it is
    /// not 0.
    Style {
        sheet: &'r str,
        addr: &'r str,
        index: u32,
    },
}

/// A record's row as its line, without the line break.
pub fn record_line(record: &Record<'_>) -> String {
    match record {
        Record::Gridlines { sheet, on } => {
            format!(
                "(gridlines {} {})",
                quoted(sheet),
                if *on { "on" } else { "off" }
            )
        }
        Record::Column { sheet, column } => {
            let width = column
                .width
                .and_then(number_text)
                .unwrap_or_else(|| "none".to_string());
            let format = column
                .style
                .map_or_else(|| "none".to_string(), |s| s.to_string());
            format!(
                "(column {} {} {} {} {})",
                quoted(sheet),
                quoted(&column_letters(column.index)),
                width,
                if column.hidden { "hidden" } else { "shown" },
                format
            )
        }
        Record::Row {
            sheet,
            number,
            settings,
        } => {
            let height = settings
                .height
                .and_then(number_text)
                .unwrap_or_else(|| "none".to_string());
            format!(
                "(row {} {} {} {})",
                quoted(sheet),
                number,
                height,
                if settings.hidden { "hidden" } else { "shown" }
            )
        }
        Record::Format { index, style, code } => format_text(*index, *style, *code),
        Record::Value { sheet, addr, value } => {
            format!("(value {} {} {})", quoted(sheet), quoted(addr), value)
        }
        Record::Style { sheet, addr, index } => {
            format!("(style {} {} {})", quoted(sheet), quoted(addr), index)
        }
    }
}

/// A `format` row: the index, the fill as six hex digits or `none`, the
/// number format, `general`, `text` or a code quoted (KERNEL.47), whether
/// text wraps, and the groups of what it says of a font, an alignment and a
/// border (KERNEL.50). An index the style table does not hold reads as the
/// default format; a code given is printed in place of the style's own.
fn format_text(index: u32, style: Option<&Style>, code: Option<&str>) -> String {
    let default = Style::default();
    let s = style.unwrap_or(&default);
    let fill = match s.fill {
        Some(rgb) => quoted(&rgb.hex()),
        None => "none".to_string(),
    };
    let num = match (code, &s.num_fmt) {
        (Some(code), _) => quoted(code),
        (None, NumFmt::General) => "general".to_string(),
        (None, NumFmt::Text) => "text".to_string(),
        (None, NumFmt::Code(code)) => quoted(code.as_str()),
    };
    // What it says of a font, an alignment and a border (KERNEL.50), after
    // its wrap; nothing for a format that says none of them.
    let groups = crate::appearance::groups_text(&s.font, &s.align, &s.border);
    format!(
        "(format {index} {fill} {num} {}{groups})",
        if s.wrap { "wrap" } else { "nowrap" }
    )
}

/// An array formula's kind, as the reader finds it beside its anchor's
/// `formula` row (KERNEL.8, slice b).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ArrayKind {
    /// Excel 365's dynamic-array formula: `t="array"`, its `cm` naming the
    /// dynamic-array metadata (`XLDAPR`), over its own cell or the range it
    /// spills into.
    Dynamic,
    /// A legacy array formula, entered with Ctrl+Shift+Enter: `t="array"`
    /// and no such `cm`; and an OpenDocument matrix formula, whose file
    /// tells no dynamic one from it.
    Legacy,
}

/// Where rows go as a walk emits them.
pub trait Sink {
    fn row(&mut self, row: &Row<'_>);

    /// The `formula` row just emitted, at this sheet and address, is an
    /// array formula's anchor of this kind (KERNEL.8, slice b): what a
    /// door computing the workbook needs to tell a formula that intersects
    /// a range where one value is wanted from one that reads it as an
    /// array. Nothing, unless a sink says otherwise: the printer ignores it,
    /// so that a formula row stays the file's own text.
    fn array(&mut self, _sheet: &str, _addr: &str, _kind: ArrayKind) {}

    /// A row of this sheet is hidden (KERNEL.8, slice e): what a door
    /// computing the workbook needs for `SUBTOTAL`, whose 101 to 111 pass a
    /// hidden row by. Nothing, unless a sink says otherwise: the printer
    /// ignores it.
    fn hidden_row(&mut self, _sheet: &str, _row: u32) {}

    /// One of the record's own rows (KERNEL.57), from a walk that prints a
    /// workbook as the record prints a sheet whole. Nothing, unless a sink
    /// says otherwise: the printer prints it.
    fn record(&mut self, _record: &Record<'_>) {}
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

    fn record(&mut self, record: &Record<'_>) {
        self.out.push_str(&record_line(record));
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
