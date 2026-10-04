//! The relations as printed (PORT.8, slice 8a): one row a line, in the
//! proof corpus's notation, so that an expected-relations file is a golden
//! the core writes, the owner checks in Excel by eye, and `VlaReadForms`
//! reads later for the VBA's own `REFLECT` test without a new parser.
//!
//! The spelling, fixed by the treaty's amendment for oracle 8: a text is a
//! VLA string as `WriteDatum` writes one; a number is the file's own text,
//! never a float round trip; a truth value is `true` or `false`; an error is
//! `(error "#DIV/0!")`; an ISO date cell is `(date "2026-10-03")`; a sheet's
//! state is `visible`, `hidden` or `very-hidden`.

use super::{Row, Sink, Value};
use crate::form::Form;
use crate::printer::write_datum;

fn quoted(s: &str) -> String {
    write_datum(&Form::string(s))
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
    use crate::reflect::Visibility;

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
}
