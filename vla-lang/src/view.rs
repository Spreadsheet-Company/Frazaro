//! The view record (KERNEL.4, 2026-10-07): one window of the sheet model as
//! lines, the first implementation of the kernel's projections seam
//! ([`crate::projection::Projection`]; the design is `web/CALLOSUM.md` §7,
//! decisions 1 and 2, and §8, slice 1).
//!
//! A view is a pure function of the model and a window: no handle kept, no
//! state, nothing read back from the `.xlsx` bytes. The model is the one the
//! writer builds (`build::build_workbook`), so a door that translates and
//! builds a program draws it without writing a file, and a viewport
//! (KERNEL.5) draws from this record alone. One form a line, in the proof
//! corpus's notation, as `reflect` prints a file:
//!
//! ```text
//! (sheet "Frazaro" visible)          every sheet of the model, in tab order
//! (window "Output" "A1:F20")         the window, its sheet as the model spells it
//! (extent "Output" "B1:D4")          the sheet's used rectangle, or none
//! (gridlines "Output" on)
//! (column "Frazaro" "B" 72 shown 1)  a column of the window with settings:
//!                                    width or none, shown or hidden, format or none
//! (format 1 "F7F4FC" text nowrap)    format 0 and every format the window uses:
//!                                    fill or none, general or text, wrap or nowrap
//! (cell "Output" "B2" 5)             a value, spelled as reflect spells it
//! (formula "Output" "B3" "=B2*2")    a formula's text as the bar shows it
//! (value "Output" "B3" 10)           its computed value, once values exist
//! (style "Output" "B2" 1)            the cell's format, when it is not 0
//! (sentence "Output" "B2" 4)         the row of the sentence that wrote the cell
//! ```
//!
//! The cells come in row-major order, each `cell` or `formula` row followed
//! by its `value` row (a formula's, once values exist), its `style` row and
//! its `sentence` row. The `cell` and `formula` rows are the reader's own,
//! spelled by `reflect::print`, which is the free oracle: the view of a
//! built model's sheet, whole, is `reflect`'s reading of the file the
//! writer writes from that model, row for row. A `cell` row is a value the
//! sentences put there, a `formula` row is text the host would compute, and
//! the record never shows one as the other. Since recalculation landed
//! (KERNEL.7, 2026-10-08; `Alonzo/SPEC.md` section 4.5 and decision 14) a
//! `(value "<sheet>" "<addr>" <v>)` row follows each formula row when the
//! caller computed the model ([`view_text_valued`], a [`Grid`] holding a
//! [`Calc`]): the value spelled as a `cell` row's is, or
//! `(not-computed "<name>")` for a formula this version does not compute,
//! naming the function or the construct, `(not-computed cycle)` for one in
//! a cycle. The bare [`view_text`] prints no value row, which is what a
//! grid shows before its first step.

use std::collections::BTreeSet;

use crate::calc::Calc;
use crate::messages::{raise, Refusal};
use crate::projection::{Projection, Window};
use crate::rows::{line as relation, quoted};
use crate::rows::{Row, Value, Visibility};
use crate::sheet::{
    cell_ref, column_letters, formula_text, number_text, parse_a1_range, shared_masters, A1Range,
    Cell, Column, Content, NumFmt, Sheet, Style, Workbook,
};

/// The grid: the view record of one window, the first projection. With a
/// [`Calc`] over the same model, each formula row is followed by its
/// `value` row; without one, the record shows formulas as text alone.
#[derive(Clone, Copy, Debug, Default)]
pub struct Grid<'v> {
    pub values: Option<&'v Calc<'v>>,
}

impl<'v> Grid<'v> {
    /// The grid over a computed model.
    pub fn valued(calc: &'v Calc<'v>) -> Grid<'v> {
        Grid { values: Some(calc) }
    }
}

impl Projection for Grid<'_> {
    fn name(&self) -> &str {
        "grid"
    }

    fn project(&self, model: &Workbook, window: &Window, out: &mut dyn FnMut(&str)) {
        for s in &model.sheets {
            out(&relation(&Row::Sheet {
                name: &s.name,
                visibility: Visibility::Visible,
            }));
        }
        let Some(i) = model.find_sheet(&window.sheet) else {
            // A window onto no sheet of the model, which `window_of` never
            // makes: the sheet rows and the window asked for, nothing more.
            out(&format!(
                "(window {} {})",
                quoted(&window.sheet),
                quoted(&window.range.text())
            ));
            return;
        };
        let sheet = &model.sheets[i];
        let name = sheet.name.as_str();
        let range = window.range;
        out(&format!(
            "(window {} {})",
            quoted(name),
            quoted(&range.text())
        ));
        match extent_of(sheet) {
            Some(e) => out(&format!("(extent {} {})", quoted(name), quoted(&e.text()))),
            None => out(&format!("(extent {} none)", quoted(name))),
        }
        out(&format!(
            "(gridlines {} {})",
            quoted(name),
            if sheet.gridlines { "on" } else { "off" }
        ));
        let mut columns: Vec<&Column> = sheet
            .columns
            .iter()
            .filter(|c| c.index >= range.left && c.index <= range.right)
            .collect();
        columns.sort_by_key(|c| c.index);
        let mut used: BTreeSet<u32> = BTreeSet::new();
        used.insert(0);
        for c in &columns {
            let width = c
                .width
                .and_then(number_text)
                .unwrap_or_else(|| "none".to_string());
            let style = match c.style {
                Some(s) => {
                    used.insert(s);
                    s.to_string()
                }
                None => "none".to_string(),
            };
            out(&format!(
                "(column {} {} {} {} {})",
                quoted(name),
                quoted(&column_letters(c.index)),
                width,
                if c.hidden { "hidden" } else { "shown" },
                style
            ));
        }
        let cells: Vec<(&(u32, u32), &Cell)> = sheet
            .cells
            .range((range.top, 0)..=(range.bottom, u32::MAX))
            .filter(|((_, col), _)| *col >= range.left && *col <= range.right)
            .collect();
        for (_, cell) in &cells {
            used.insert(cell.style);
        }
        for index in used {
            out(&format_row(index, model.styles.xfs().get(index as usize)));
        }
        let masters = shared_masters(sheet);
        for ((row, col), cell) in cells {
            let addr = cell_ref(*row, *col);
            let printed = match &cell.content {
                Content::Text(t) => {
                    let value = Value::Text(t.clone());
                    out(&relation(&Row::Cell {
                        sheet: name,
                        addr: &addr,
                        value: &value,
                    }));
                    true
                }
                Content::Number(n) => {
                    // What the writer puts in the file: the number's text,
                    // or Excel's #NUM! for one no cell can hold.
                    let value = match number_text(*n) {
                        Some(text) => Value::Number(text),
                        None => Value::Error("#NUM!".to_string()),
                    };
                    out(&relation(&Row::Cell {
                        sheet: name,
                        addr: &addr,
                        value: &value,
                    }));
                    true
                }
                Content::Bool(b) => {
                    let value = Value::Bool(*b);
                    out(&relation(&Row::Cell {
                        sheet: name,
                        addr: &addr,
                        value: &value,
                    }));
                    true
                }
                Content::Error(e) => {
                    let value = Value::Error(e.clone());
                    out(&relation(&Row::Cell {
                        sheet: name,
                        addr: &addr,
                        value: &value,
                    }));
                    true
                }
                other => match formula_text(other, *row, *col, &masters) {
                    Some(text) => {
                        let text = format!("={text}");
                        out(&relation(&Row::Formula {
                            sheet: name,
                            addr: &addr,
                            text: &text,
                        }));
                        if let Some(computed) =
                            self.values.and_then(|calc| calc.computed((i, *row, *col)))
                        {
                            out(&format!(
                                "(value {} {} {})",
                                quoted(name),
                                quoted(&addr),
                                computed.spell()
                            ));
                        }
                        true
                    }
                    None => false,
                },
            };
            if !printed {
                continue;
            }
            if cell.style != 0 {
                out(&format!(
                    "(style {} {} {})",
                    quoted(name),
                    quoted(&addr),
                    cell.style
                ));
            }
            if let Some(r) = sheet.sentences.get(&(*row, *col)) {
                out(&format!(
                    "(sentence {} {} {})",
                    quoted(name),
                    quoted(&addr),
                    r
                ));
            }
        }
    }
}

/// A `format` row: the index, the fill as six hex digits or `none`, the
/// number format, and whether text wraps. An index the style table does not
/// hold reads as the default format.
fn format_row(index: u32, style: Option<&Style>) -> String {
    let default = Style::default();
    let s = style.unwrap_or(&default);
    let fill = match s.fill {
        Some(rgb) => quoted(&format!("{:02X}{:02X}{:02X}", rgb.0, rgb.1, rgb.2)),
        None => "none".to_string(),
    };
    let num = match s.num_fmt {
        NumFmt::General => "general",
        NumFmt::Text => "text",
    };
    format!(
        "(format {index} {fill} {num} {})",
        if s.wrap { "wrap" } else { "nowrap" }
    )
}

/// The sheet's used rectangle, or `None` for a sheet with no cell.
fn extent_of(sheet: &Sheet) -> Option<A1Range> {
    sheet
        .extent()
        .map(|((top, left), (bottom, right))| A1Range {
            top,
            left,
            bottom,
            right,
        })
}

/// The window over a whole sheet: its extent, or `A1` alone for a sheet
/// with no cell.
pub fn whole(sheet: &Sheet) -> A1Range {
    extent_of(sheet).unwrap_or(A1Range {
        top: 1,
        left: 1,
        bottom: 1,
        right: 1,
    })
}

/// The window a door asked for, against the model: the sheet by name as
/// Excel compares names, without case, spelled in the window as the model
/// spells it; the range as written (`A1:F20`, `b2`, corners in any order),
/// or the sheet's whole extent when none is given. A sheet the model does
/// not hold is refused naming the ones it does (`view-sheet-unknown`); a
/// text that is not a rectangle of cells, a whole column or row included,
/// is refused as written (`view-window-not-a-range`).
pub fn window_of(model: &Workbook, sheet: &str, range: Option<&str>) -> Result<Window, Refusal> {
    let Some(i) = model.find_sheet(sheet) else {
        let sheets = model
            .sheets
            .iter()
            .map(|s| s.name.as_str())
            .collect::<Vec<_>>()
            .join(", ");
        return Err(raise(
            "view-sheet-unknown",
            &[("name", sheet), ("sheets", &sheets)],
        ));
    };
    let s = &model.sheets[i];
    let range = match range {
        Some(text) => parse_a1_range(text)
            .ok_or_else(|| raise("view-window-not-a-range", &[("text", text)]))?,
        None => whole(s),
    };
    Ok(Window {
        sheet: s.name.clone(),
        range,
    })
}

/// The view record of one window as text, each line ending in a line feed:
/// the API's surface, the tests' and a door's that holds the text.
pub fn view_text(model: &Workbook, window: &Window) -> String {
    let mut out = String::new();
    Grid::default().project(model, window, &mut |line| {
        out.push_str(line);
        out.push('\n');
    });
    out
}

/// The view record with each formula's computed value after it: the record
/// a door prints once it has recalculated the model (`frazaro view` through
/// the core's Excel library; an engine after a step). `calc` must be over
/// `model`.
pub fn view_text_valued(model: &Workbook, window: &Window, calc: &Calc<'_>) -> String {
    let mut out = String::new();
    Grid::valued(calc).project(model, window, &mut |line| {
        out.push_str(line);
        out.push('\n');
    });
    out
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::sheet::{Rgb, Styles};

    fn range(text: &str) -> A1Range {
        parse_a1_range(text).unwrap()
    }

    fn lines(text: &str) -> Vec<&str> {
        text.lines().collect()
    }

    /// A small model made by hand: two sheets, the first with columns and
    /// formats, a shared formula, a dynamic one, a truth value, a number no
    /// cell can hold, and the rows of the sentences that wrote them.
    fn a_model() -> Workbook {
        let mut wb = Workbook::new();
        let mut styles = Styles::new();
        let lavender = styles.id(Style {
            fill: Some(Rgb(247, 244, 252)),
            num_fmt: NumFmt::Text,
            wrap: false,
        });
        let gray = styles.id(Style {
            fill: Some(Rgb(242, 242, 242)),
            num_fmt: NumFmt::General,
            wrap: true,
        });
        wb.styles = styles;
        let mut model = Sheet::new("Model");
        model.gridlines = false;
        model.columns = vec![
            Column {
                index: 3,
                width: Some(60.0),
                hidden: false,
                style: Some(gray),
            },
            Column {
                index: 1,
                width: None,
                hidden: true,
                style: None,
            },
            Column {
                index: 2,
                width: Some(72.0),
                hidden: false,
                style: Some(lavender),
            },
        ];
        model.set(
            1,
            2,
            Cell {
                content: Content::Text("Revenue".to_string()),
                style: lavender,
            },
        );
        model.set(
            2,
            2,
            Cell {
                content: Content::Number(1200.5),
                style: 0,
            },
        );
        model.set(
            3,
            2,
            Cell {
                content: Content::Bool(true),
                style: 0,
            },
        );
        model.set(
            4,
            2,
            Cell {
                content: Content::Number(f64::INFINITY),
                style: 0,
            },
        );
        model.set_formula(range("C2:C4"), "B2*2", 0);
        model.set_formula_dynamic(range("D2"), "_xlfn.IFS(B2>3,\"big\",TRUE,\"small\")", 0);
        model.wrote(range("B1"), 1);
        model.wrote(range("B2:B4"), 2);
        model.wrote(range("C2:C4"), 3);
        model.wrote(range("D2"), 4);
        wb.sheets.push(model);
        wb.sheets.push(Sheet::new("Empty"));
        wb
    }

    #[test]
    fn the_record_s_rows_have_the_fixed_spelling_in_the_fixed_order() {
        let wb = a_model();
        let text = view_text(
            &wb,
            &Window {
                sheet: "Model".to_string(),
                range: range("A1:F20"),
            },
        );
        assert_eq!(
            lines(&text),
            vec![
                "(sheet \"Model\" visible)",
                "(sheet \"Empty\" visible)",
                "(window \"Model\" \"A1:F20\")",
                "(extent \"Model\" \"B1:D4\")",
                "(gridlines \"Model\" off)",
                "(column \"Model\" \"A\" none hidden none)",
                "(column \"Model\" \"B\" 72 shown 1)",
                "(column \"Model\" \"C\" 60 shown 2)",
                "(format 0 none general nowrap)",
                "(format 1 \"F7F4FC\" text nowrap)",
                "(format 2 \"F2F2F2\" general wrap)",
                "(cell \"Model\" \"B1\" \"Revenue\")",
                "(style \"Model\" \"B1\" 1)",
                "(sentence \"Model\" \"B1\" 1)",
                "(cell \"Model\" \"B2\" 1200.5)",
                "(sentence \"Model\" \"B2\" 2)",
                "(formula \"Model\" \"C2\" \"=B2*2\")",
                "(sentence \"Model\" \"C2\" 3)",
                "(formula \"Model\" \"D2\" \"=IFS(B2>3,\\\"big\\\",TRUE,\\\"small\\\")\")",
                "(sentence \"Model\" \"D2\" 4)",
                "(cell \"Model\" \"B3\" true)",
                "(sentence \"Model\" \"B3\" 2)",
                "(formula \"Model\" \"C3\" \"=B3*2\")",
                "(sentence \"Model\" \"C3\" 3)",
                "(cell \"Model\" \"B4\" (error \"#NUM!\"))",
                "(sentence \"Model\" \"B4\" 2)",
                "(formula \"Model\" \"C4\" \"=B4*2\")",
                "(sentence \"Model\" \"C4\" 3)",
            ]
        );
    }

    #[test]
    fn a_window_clips_to_its_rectangle_and_an_empty_one_prints_its_header_alone() {
        let wb = a_model();
        // The clip: the shared formula's children keep their moved text, and
        // only column C's settings are in the window.
        let text = view_text(
            &wb,
            &Window {
                sheet: "Model".to_string(),
                range: range("C3:D4"),
            },
        );
        assert_eq!(
            lines(&text),
            vec![
                "(sheet \"Model\" visible)",
                "(sheet \"Empty\" visible)",
                "(window \"Model\" \"C3:D4\")",
                "(extent \"Model\" \"B1:D4\")",
                "(gridlines \"Model\" off)",
                "(column \"Model\" \"C\" 60 shown 2)",
                "(format 0 none general nowrap)",
                "(format 2 \"F2F2F2\" general wrap)",
                "(formula \"Model\" \"C3\" \"=B3*2\")",
                "(sentence \"Model\" \"C3\" 3)",
                "(formula \"Model\" \"C4\" \"=B4*2\")",
                "(sentence \"Model\" \"C4\" 3)",
            ]
        );
        // Past the extent: the header says where the cells are, and no cell row follows.
        let text = view_text(
            &wb,
            &Window {
                sheet: "Model".to_string(),
                range: range("H10:J12"),
            },
        );
        assert_eq!(
            lines(&text)[2..],
            [
                "(window \"Model\" \"H10:J12\")",
                "(extent \"Model\" \"B1:D4\")",
                "(gridlines \"Model\" off)",
                "(format 0 none general nowrap)",
            ]
        );
        // An empty sheet: no extent, gridlines on, the default format.
        let text = view_text(
            &wb,
            &Window {
                sheet: "Empty".to_string(),
                range: whole(&wb.sheets[1]),
            },
        );
        assert_eq!(
            lines(&text)[2..],
            [
                "(window \"Empty\" \"A1\")",
                "(extent \"Empty\" none)",
                "(gridlines \"Empty\" on)",
                "(format 0 none general nowrap)",
            ]
        );
        // A window onto no sheet, which window_of never makes: the sheets,
        // the window, and nothing else.
        let text = view_text(
            &wb,
            &Window {
                sheet: "Nowhere".to_string(),
                range: range("A1"),
            },
        );
        assert_eq!(lines(&text).len(), 3);
        assert_eq!(lines(&text)[2], "(window \"Nowhere\" \"A1\")");
    }

    #[test]
    fn window_of_finds_the_sheet_without_case_and_refuses_the_rest_by_name() {
        let wb = a_model();
        let w = window_of(&wb, "MODEL", None).unwrap();
        assert_eq!(w.sheet, "Model");
        assert_eq!(w.range, range("B1:D4"));
        let w = window_of(&wb, "empty", None).unwrap();
        assert_eq!((w.sheet.as_str(), w.range.text().as_str()), ("Empty", "A1"));
        let w = window_of(&wb, "Model", Some("f20:a1")).unwrap();
        assert_eq!(w.range.text(), "A1:F20");
        let r = window_of(&wb, "Summary", None).unwrap_err();
        assert_eq!(
            (r.id.as_str(), r.source.as_str()),
            ("view-sheet-unknown", "VLA-View")
        );
        assert!(r.text.contains("Summary"), "{}", r.text);
        assert!(r.text.contains("Model, Empty"), "{}", r.text);
        for text in ["A:A", "3:3", "nonsense", "", "A1:B"] {
            let r = window_of(&wb, "Model", Some(text)).unwrap_err();
            assert_eq!(r.id, "view-window-not-a-range", "{text}");
            assert!(r.text.contains(text), "{}", r.text);
        }
        assert_eq!(Grid::default().name(), "grid");
    }
}
