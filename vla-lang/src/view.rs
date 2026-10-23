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
//! (sheet "Frazaro" visible)          every sheet of the model, in tab order, with
//!                                    its own state: visible, hidden, very-hidden
//! (window "Output" "A1:F20")         the window, its sheet as the model spells it
//! (extent "Output" "B1:D4")          the sheet's used rectangle, or none
//! (gridlines "Output" on)
//! (column "Frazaro" "B" 72 shown 1)  a column of the window with settings:
//!                                    width or none, shown or hidden, format or none
//! (row "Board" 3 30 shown)           a row of the window with settings (KERNEL.22):
//!                                    height or none, shown or hidden
//! (format 1 "F7F4FC" text nowrap)    format 0 and every format the window uses:
//!                                    fill or none, general, text or a code of
//!                                    Excel's in quotes (KERNEL.47), wrap or nowrap,
//!                                    then what it says of a font, an alignment
//!                                    and a border, (font ...) (align ...)
//!                                    (border ...), where it says any (KERNEL.50;
//!                                    `crate::appearance`)
//! (look "Board" 1 2)                 the sheet's look map (KERNEL.22): a cell
//!                                    holding the value draws with the format
//! (cell "Output" "B2" 5)             a value, spelled as reflect spells it
//! (formula "Output" "B3" "=B2*2")    a formula's text as the bar shows it
//! (value "Output" "B3" 10)           its computed value, once values exist
//! (text "Output" "B3" "10.00")       a number's text under its format's code
//!                                    (KERNEL.47), after its cell or value row
//! (style "Output" "B2" 1)            the cell's format, when it is not 0
//! (sentence "Output" "B2" 4)         the row of the sentence that wrote the cell
//! ```
//!
//! The cells come in row-major order, each `cell` or `formula` row followed
//! by its `value` row (a formula's, once values exist), its `text` row, its
//! `style` row and its `sentence` row. A `text` row is the text Excel draws
//! for the cell's number, or for its text under a code's fourth section,
//! under the code of the format the cell draws with, its look's where its
//! value has one (KERNEL.47; `crate::numfmt`): so a viewport draws the text
//! and never respells a value. A record whose formats hold no code prints
//! none. The `cell` and `formula` rows are the reader's own,
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
//!
//! The second projection, the values ([`ValueGrid`], `values`; KERNEL.37,
//! 2026-10-12), is the same record with its `formula` rows left out: every
//! other row the grid's, in the grid's order, so that for any window it is
//! the grid record with its `(formula ` lines removed, byte for byte. It is
//! what an engine's host draws every frame ([`view_text_values`]), a
//! formula's text being the bar's, one cell at a time, where the grid moves
//! every shared child's text from its master at every view. A formula with
//! no value yet prints no row of its own, its `style` and `sentence` rows
//! kept; a shared formula's cell whose master the sheet has lost prints
//! nothing in either record. Both read a sheet's `extent` row from the
//! extent their values keep ([`Values::kept_extent`]), where they keep one.
//!
//! The changes of a window (KERNEL.42, 2026-10-15) are an engine's machine's
//! to answer, since they need two frames, and are printed here: a `changes`
//! row naming the sheet, the window, the frame they are since and the frame
//! they are at, and how many cells changed, then each changed cell's rows as
//! the values print them, by the same code, in the same order; a cell that
//! now prints no row is `(empty "<sheet>" "<addr>")`. Where the machine
//! cannot say which cells changed, the row ends `whole` and none follows,
//! and the host views the window. A host holding the window's values at the
//! first frame, and replacing each named cell's rows with the rows given it,
//! holds the window's values at the second, byte for byte.

use std::collections::BTreeSet;

use crate::calc::{self, Extent, Values};
use crate::grid::{text_of, Cells, Constant, Formula, Held, Holds, Masters, Page};
use crate::machine::load::Datum;
use crate::messages::{raise, Refusal};
use crate::projection::{Projection, Window};
use crate::rows::{line as relation, quoted, record_line, Record};
use crate::rows::{Row, Value};
use crate::sheet::{
    cell_ref, number_text, parse_a1_range, A1Range, Column, NumFmt, Sheet, Styles, Workbook,
};

/// The grid: the view record of one window, the first projection. With the
/// values of a recalculation over the same model (a [`crate::calc::Calc`],
/// or an engine's machine since KERNEL.22), each formula row is followed by
/// its `value` row; without them, the record shows formulas as text alone.
#[derive(Clone, Copy, Debug, Default)]
pub struct Grid<'v> {
    pub values: Option<&'v dyn Values>,
}

impl<'v> Grid<'v> {
    /// The grid over a computed model.
    pub fn valued(values: &'v dyn Values) -> Grid<'v> {
        Grid {
            values: Some(values),
        }
    }
}

impl Projection for Grid<'_> {
    fn name(&self) -> &str {
        "grid"
    }

    fn project(&self, model: &Workbook, window: &Window, out: &mut dyn FnMut(&str)) {
        self.rows(model, window, true, out);
    }
}

/// The values: the view record of one window with its `formula` rows left
/// out, the second projection (KERNEL.37). With the values of a
/// recalculation or an engine's machine, each formula's `value` row stands
/// where the grid prints it, after the row the grid prints before it; the
/// record of a window holding no formula is the grid's.
#[derive(Clone, Copy, Debug, Default)]
pub struct ValueGrid<'v> {
    pub values: Option<&'v dyn Values>,
}

impl<'v> ValueGrid<'v> {
    /// The values of a computed model.
    pub fn valued(values: &'v dyn Values) -> ValueGrid<'v> {
        ValueGrid {
            values: Some(values),
        }
    }
}

impl Projection for ValueGrid<'_> {
    fn name(&self) -> &str {
        "values"
    }

    fn project(&self, model: &Workbook, window: &Window, out: &mut dyn FnMut(&str)) {
        Grid {
            values: self.values,
        }
        .rows(model, window, false, out);
    }
}

impl Grid<'_> {
    /// The record's rows of one window: every row the grid prints, or with
    /// `formulas` false every row but the `formula` rows, which is the values
    /// (KERNEL.37); over a grid, a workbook's or an engine's machine's, whose
    /// twins' cells it makes from their sheets (KERNEL.49).
    pub(crate) fn rows<G: Cells + ?Sized>(
        &self,
        model: &G,
        window: &Window,
        formulas: bool,
        out: &mut dyn FnMut(&str),
    ) {
        for s in 0..model.sheet_count() {
            let page = model.page(s);
            out(&relation(&Row::Sheet {
                name: page.name,
                visibility: page.visibility,
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
        let page = model.page(i);
        let name = page.name;
        let range = window.range;
        out(&format!(
            "(window {} {})",
            quoted(name),
            quoted(&range.text())
        ));
        match extent_in(model, i, self.values) {
            Some(e) => out(&format!("(extent {} {})", quoted(name), quoted(&e.text()))),
            None => out(&format!("(extent {} none)", quoted(name))),
        }
        out(&record_line(&Record::Gridlines {
            sheet: name,
            on: page.gridlines,
        }));
        let mut columns: Vec<&Column> = page
            .columns
            .iter()
            .filter(|c| c.index >= range.left && c.index <= range.right)
            .collect();
        columns.sort_by_key(|c| c.index);
        let mut used: BTreeSet<u32> = BTreeSet::new();
        used.insert(0);
        for c in &columns {
            if let Some(s) = c.style {
                used.insert(s);
            }
            out(&record_line(&Record::Column {
                sheet: name,
                column: c,
            }));
        }
        // The rows of the window with settings, the column rows' twins
        // (KERNEL.22; Alonzo/SPEC.md decision 19).
        for (n, r) in page.rows.range(range.top..=range.bottom) {
            out(&record_line(&Record::Row {
                sheet: name,
                number: *n,
                settings: r,
            }));
        }
        // The window's cells, as the grid holds them, an engine's twin's
        // made from its sheet (KERNEL.49).
        let mut cells: Vec<((u32, u32), Held<'_>)> = Vec::new();
        model.cells(i, &range, &mut |at, cell| cells.push((at, cell)));
        for (_, cell) in &cells {
            used.insert(cell.style);
        }
        for (_, index) in page.looks {
            used.insert(*index);
        }
        for index in used {
            out(&record_line(&Record::Format {
                index,
                style: model.styles().xfs().get(index as usize),
                code: None,
            }));
        }
        // The sheet's look map, a value to the format a cell holding it draws
        // with (KERNEL.22; Alonzo/SPEC.md decision 19).
        for (value, index) in page.looks {
            out(&format!("(look {} {} {})", quoted(name), value, index));
        }
        // The sheet's fills' masters, found at the first cell that needs
        // them (KERNEL.37): the grid moves a fill's cell's text from its
        // master; the values ask only whether a cell with no value has one.
        let mut masters = None;
        let coded = Coded::of(model.styles(), page.looks);
        #[cfg(test)]
        crate::sheet::work::viewed(cells.len());
        for (at, cell) in &cells {
            self.cell_rows(
                model,
                &page,
                i,
                *at,
                cell,
                formulas,
                &mut masters,
                coded.as_ref(),
                out,
            );
        }
    }

    /// One cell's rows, as the record prints them (KERNEL.42: the changes
    /// print a changed cell's by this code too): its `cell` row, or its
    /// `formula` row where the record has them and its `value` row once it
    /// has one, then its `text` row where its value draws under a code
    /// (KERNEL.47), then its `style` and `sentence` rows; whether it printed
    /// any. `masters` are the sheet's fills', found once, and `coded` what a
    /// `text` row needs, where the style table holds a code.
    #[allow(clippy::too_many_arguments)]
    fn cell_rows<'g, G: Cells + ?Sized>(
        &self,
        model: &'g G,
        page: &Page<'_>,
        i: usize,
        (row, col): (u32, u32),
        cell: &Held<'_>,
        formulas: bool,
        masters: &mut Option<Masters<'g>>,
        coded: Option<&Coded<'_>>,
        out: &mut dyn FnMut(&str),
    ) -> bool {
        let name = page.name;
        let addr = cell_ref(row, col);
        // Whether the cell printed a row, and the value a code draws.
        let (printed, shown) = match &cell.holds {
            Holds::Constant(Constant::Text(t)) => {
                let value = Value::Text(t.to_string());
                out(&relation(&Row::Cell {
                    sheet: name,
                    addr: &addr,
                    value: &value,
                }));
                (true, Some(calc::Value::Text(t.to_string())))
            }
            Holds::Constant(Constant::Number(n)) => {
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
                (true, Some(calc::Value::Number(*n)))
            }
            Holds::Constant(Constant::Bool(b)) => {
                let value = Value::Bool(*b);
                out(&relation(&Row::Cell {
                    sheet: name,
                    addr: &addr,
                    value: &value,
                }));
                (true, None)
            }
            Holds::Constant(Constant::Error(e)) => {
                let value = Value::Error(e.to_string());
                out(&relation(&Row::Cell {
                    sheet: name,
                    addr: &addr,
                    value: &value,
                }));
                (true, None)
            }
            Holds::Formula(formula) => {
                let computed = self.values.and_then(|calc| calc.computed((i, row, col)));
                // A formula with a text: every one but a fill's cell whose
                // master the sheet has lost, which shows nothing in either
                // record.
                let held = if formulas {
                    let walked = masters.get_or_insert_with(|| model.masters(i));
                    match text_of(*formula, (row, col), walked) {
                        Some(text) => {
                            let text = format!("={text}");
                            out(&relation(&Row::Formula {
                                sheet: name,
                                addr: &addr,
                                text: &text,
                            }));
                            true
                        }
                        None => false,
                    }
                } else {
                    match formula {
                        // A value is only ever computed for a formula with
                        // a text, so a fill's cell with one has its master;
                        // one without asks the sheet.
                        Formula::Fill(si) if computed.is_none() => masters
                            .get_or_insert_with(|| model.masters(i))
                            .contains_key(si),
                        _ => true,
                    }
                };
                let mut shown = None;
                if let Some(computed) = computed.filter(|_| held) {
                    out(&record_line(&Record::Value {
                        sheet: name,
                        addr: &addr,
                        value: &computed.spell(),
                    }));
                    // The value as the row spells it: an empty result is 0.
                    if let calc::Computed::Value(v) = computed {
                        shown = Some(v.clone().settled());
                    }
                }
                (held, shown)
            }
        };
        if !printed {
            return false;
        }
        if let (Some(coded), Some(v)) = (coded, shown) {
            if let Some(text) = coded.text(cell.style, &v) {
                out(&format!(
                    "(text {} {} {})",
                    quoted(name),
                    quoted(&addr),
                    quoted(&text)
                ));
            }
        }
        if cell.style != 0 {
            out(&record_line(&Record::Style {
                sheet: name,
                addr: &addr,
                index: cell.style,
            }));
        }
        if let Some(r) = cell.sentence {
            out(&format!(
                "(sentence {} {} {})",
                quoted(name),
                quoted(&addr),
                r
            ));
        }
        true
    }
}

/// What a record needs for a cell's `text` row (KERNEL.47): the style table
/// and the sheet's looks read back from their spellings, made once a view
/// and only where the table holds a code, so that a record whose formats
/// hold none prints as it did and finds no look.
struct Coded<'m> {
    styles: &'m Styles,
    looks: Vec<(Look, u32)>,
}

/// A look's value, as a value a code draws meets it: a number by its value,
/// so that 7 and 7.0 are one and so are 0 and -0, and a text by its
/// characters, the viewport's rule (Alonzo's `SPEC.md` §3.1). A look of a
/// truth value or an error matches no value a code draws.
enum Look {
    Number(f64),
    Text(String),
}

impl<'m> Coded<'m> {
    fn of(styles: &'m Styles, looks: &[(String, u32)]) -> Option<Coded<'m>> {
        if !styles.has_code() {
            return None;
        }
        let looks = looks
            .iter()
            .filter_map(|(spelled, index)| {
                let forms = crate::reader::read_forms(spelled).ok()?;
                let look = match crate::machine::load::datum(forms.first()?)? {
                    Datum::Number(n) => Look::Number(n),
                    Datum::Text(t) => Look::Text(t),
                    _ => return None,
                };
                Some((look, *index))
            })
            .collect();
        Some(Coded { styles, looks })
    }

    /// The text a value draws as in a cell of format `style`: under its
    /// look's format in place of the cell's own where its value has one, as
    /// the viewport draws a look (Alonzo's `SPEC.md` decision 19), the first
    /// look of the record's that it meets; `None` where that format holds no
    /// code or the code draws the value as the record spells it.
    fn text(&self, style: u32, value: &calc::Value) -> Option<String> {
        let met = self.looks.iter().find(|(look, _)| match (look, value) {
            (Look::Number(a), calc::Value::Number(b)) => a == b,
            (Look::Text(a), calc::Value::Text(b)) => a == b,
            _ => false,
        });
        let index = met.map_or(style, |(_, index)| *index);
        match self.styles.get(index).map(|s| &s.num_fmt) {
            Some(NumFmt::Code(code)) => code.text(value),
            _ => None,
        }
    }
}

/// The used rectangle of sheet `i`, read from the extent `values` keep
/// where they keep one, else the grid's (KERNEL.37): a workbook's walked, a
/// machine's as it keeps it.
fn extent_in<G: Cells + ?Sized>(
    model: &G,
    i: usize,
    values: Option<&dyn Values>,
) -> Option<A1Range> {
    match values.and_then(|v| v.kept_extent(i)) {
        Some(kept) => rectangle(kept),
        None => rectangle(model.extent(i)),
    }
}

/// An extent as a rectangle of cells.
fn rectangle(extent: Extent) -> Option<A1Range> {
    extent.map(|((top, left), (bottom, right))| A1Range {
        top,
        left,
        bottom,
        right,
    })
}

/// The window over a whole sheet: its extent, or `A1` alone for a sheet
/// with no cell.
pub fn whole(sheet: &Sheet) -> A1Range {
    or_a1(rectangle(sheet.extent()))
}

/// A rectangle, or `A1` alone for none.
fn or_a1(range: Option<A1Range>) -> A1Range {
    range.unwrap_or(A1Range {
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
    window_over(model, sheet, range, None)
}

/// [`window_of`] over a grid, a workbook's or an engine's machine's
/// (KERNEL.49), a whole sheet's extent read from the extent `values` keep
/// where they keep one (KERNEL.37): the window an engine's machine views.
pub(crate) fn window_over<G: Cells + ?Sized>(
    model: &G,
    sheet: &str,
    range: Option<&str>,
    values: Option<&dyn Values>,
) -> Result<Window, Refusal> {
    let Some(i) = model.find_sheet(sheet) else {
        let sheets = (0..model.sheet_count())
            .filter_map(|s| model.sheet_name(s))
            .collect::<Vec<_>>()
            .join(", ");
        return Err(raise(
            "view-sheet-unknown",
            &[("name", sheet), ("sheets", &sheets)],
        ));
    };
    let range = match range {
        Some(text) => parse_a1_range(text)
            .ok_or_else(|| raise("view-window-not-a-range", &[("text", text)]))?,
        None => or_a1(extent_in(model, i, values)),
    };
    Ok(Window {
        sheet: model.sheet_name(i).unwrap_or_default().to_string(),
        range,
    })
}

/// The view record of one window as text, each line ending in a line feed:
/// the API's surface, the tests' and a door's that holds the text.
pub fn view_text(model: &Workbook, window: &Window) -> String {
    record_text(model, window, None, true)
}

/// The view record with each formula's computed value after it: the record
/// a door prints once it has recalculated the model (`frazaro view` through
/// the core's Excel library; an engine after a step). `values` must be over
/// `model`.
pub fn view_text_valued(model: &Workbook, window: &Window, values: &dyn Values) -> String {
    record_text(model, window, Some(values), true)
}

/// The values record of one window (KERNEL.37): [`view_text_valued`]'s
/// record with its `formula` lines left out, what an engine's host draws
/// every frame. `values` must be over `model`.
pub fn view_text_values(model: &Workbook, window: &Window, values: &dyn Values) -> String {
    record_text(model, window, Some(values), false)
}

/// A window's record as text over a grid, a workbook's or an engine's
/// machine's (KERNEL.49): the grid's, or with `formulas` false the values',
/// each line ending in a line feed.
pub(crate) fn record_text<G: Cells + ?Sized>(
    model: &G,
    window: &Window,
    values: Option<&dyn Values>,
    formulas: bool,
) -> String {
    let mut out = String::new();
    Grid { values }.rows(model, window, formulas, &mut |line| {
        out.push_str(line);
        out.push('\n');
    });
    out
}

/// The changes of one window (KERNEL.42) as text, each line ending in a
/// line feed: `(changes "<sheet>" "<range>" <since> <at> <n>)`, the window
/// as its values record names it, then the rows of the `n` cells `cells`
/// names, row by row, each as the values print it, a cell that prints none
/// `(empty "<sheet>" "<addr>")`; or, `cells` being `None`, the row ending
/// `whole` and no other. An engine's machine says which cells changed, and
/// only it can (`crate::machine::Machine::view`); `values` must be over
/// `model`, and `cells` inside the window, in row-major order.
pub(crate) fn changes_text<G: Cells + ?Sized>(
    model: &G,
    window: &Window,
    values: &dyn Values,
    (since, at): (u64, u64),
    cells: Option<&[(u32, u32)]>,
) -> String {
    let Some(i) = model.find_sheet(&window.sheet) else {
        return String::new();
    };
    let page = model.page(i);
    let head = |n: &str| {
        format!(
            "(changes {} {} {since} {at} {n})\n",
            quoted(page.name),
            quoted(&window.range.text())
        )
    };
    let Some(cells) = cells else {
        return head("whole");
    };
    let mut out = head(&cells.len().to_string());
    #[cfg(test)]
    crate::sheet::work::viewed(cells.len());
    let grid = Grid::valued(values);
    let coded = Coded::of(model.styles(), page.looks);
    let mut masters = None;
    let mut push = |line: &str| {
        out.push_str(line);
        out.push('\n');
    };
    for &(row, col) in cells {
        // The cell as the grid holds it, an engine's twin's made from its
        // sheet (KERNEL.49).
        let printed = model.cell(i, row, col).is_some_and(|cell| {
            grid.cell_rows(
                model,
                &page,
                i,
                (row, col),
                &cell,
                false,
                &mut masters,
                coded.as_ref(),
                &mut push,
            )
        });
        if !printed {
            push(&format!(
                "(empty {} {})",
                quoted(page.name),
                quoted(&cell_ref(row, col))
            ));
        }
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::sheet::{Cell, Content};
    use crate::sheet::{Rgb, Style, Styles};

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
        let lavender = styles.id(Style::new(Some(Rgb(247, 244, 252)), NumFmt::Text, false));
        let gray = styles.id(Style::new(Some(Rgb(242, 242, 242)), NumFmt::General, true));
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
    fn the_changes_print_each_cell_s_rows_as_the_values_print_them() {
        // KERNEL.42: the changes' text of the cells a machine named, each
        // cell's rows the values' own, printed by the same code: a value
        // cell's with its style and sentence rows, a formula's value row and
        // its sentence row; a cell that prints no row, absent or a shared
        // formula's cell whose master the sheet has lost, spelled empty; and
        // the row alone, ending whole, where the machine cannot say.
        let mut wb = a_model();
        let library = crate::calc::library::language();
        let calc = crate::calc::Calc::run(&wb, &library);
        let w = Window {
            sheet: "Model".to_string(),
            range: range("A1:F20"),
        };
        let values = view_text_values(&wb, &w, &calc);
        let rows = |addr: &str| -> String {
            let at = format!(" \"{addr}\" ");
            values
                .lines()
                .filter(|l| l.contains(&at))
                .map(|l| format!("{l}\n"))
                .collect()
        };
        let text = changes_text(&wb, &w, &calc, (6, 7), Some(&[(1, 2), (2, 3), (9, 9)]));
        assert_eq!(
            text,
            format!(
                "(changes \"Model\" \"A1:F20\" 6 7 3)\n{}{}(empty \"Model\" \"I9\")\n",
                rows("B1"),
                rows("C2")
            )
        );
        assert_eq!(rows("B1").lines().count(), 3, "{values}");
        assert_eq!(
            changes_text(&wb, &w, &calc, (6, 7), None),
            "(changes \"Model\" \"A1:F20\" 6 7 whole)\n"
        );
        wb.sheets[0].cells.remove(&(2, 3));
        let lost = crate::calc::Calc::run(&wb, &library);
        assert_eq!(
            changes_text(&wb, &w, &lost, (0, 1), Some(&[(3, 3)])),
            "(changes \"Model\" \"A1:F20\" 0 1 1)\n(empty \"Model\" \"C3\")\n",
            "a fill's cell whose master is lost"
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
    fn row_settings_looks_and_hidden_sheets_print_in_their_places() {
        // KERNEL.22: the `row` rows after the `column` rows, the `look` rows
        // after the `format` rows with their formats among them, and each
        // sheet's own visibility.
        let mut wb = a_model();
        wb.sheets[1].visibility = crate::rows::Visibility::Hidden;
        let model = &mut wb.sheets[0];
        model.rows.insert(
            3,
            crate::sheet::RowSettings {
                height: Some(30.0),
                hidden: false,
            },
        );
        model.rows.insert(
            9,
            crate::sheet::RowSettings {
                height: None,
                hidden: true,
            },
        );
        model.looks.push(("1".to_string(), 2));
        let text = view_text(
            &wb,
            &Window {
                sheet: "Model".to_string(),
                range: range("B2:B4"),
            },
        );
        assert_eq!(
            lines(&text)[..10],
            [
                "(sheet \"Model\" visible)",
                "(sheet \"Empty\" hidden)",
                "(window \"Model\" \"B2:B4\")",
                "(extent \"Model\" \"B1:D4\")",
                "(gridlines \"Model\" off)",
                "(column \"Model\" \"B\" 72 shown 1)",
                "(row \"Model\" 3 30 shown)",
                "(format 0 none general nowrap)",
                "(format 1 \"F7F4FC\" text nowrap)",
                "(format 2 \"F2F2F2\" general wrap)",
            ]
        );
        assert_eq!(lines(&text)[10], "(look \"Model\" 1 2)");
        assert!(
            !text.contains("(row \"Model\" 9"),
            "row 9 is outside the window"
        );
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

    /// A record less its `formula` lines: what the values must be (KERNEL.37).
    fn less_formulas(record: &str) -> String {
        record
            .lines()
            .filter(|l| !l.starts_with("(formula "))
            .map(|l| format!("{l}\n"))
            .collect()
    }

    /// The values record of a window, over the values given or none.
    fn values_text(wb: &Workbook, window: &Window, values: Option<&dyn Values>) -> String {
        let mut out = String::new();
        ValueGrid { values }.project(wb, window, &mut |line| {
            out.push_str(line);
            out.push('\n');
        });
        out
    }

    /// What `f` walked, counted (`sheet::work`).
    fn walks(f: impl FnOnce()) -> crate::sheet::work::Work {
        crate::sheet::work::begin();
        f();
        crate::sheet::work::end()
    }

    #[test]
    fn the_values_are_the_grid_less_its_formula_rows() {
        // KERNEL.37: every window of the model, computed and not, its values
        // its grid record with the formula lines left out, a formula's style
        // and sentence rows kept whether it has a value or not.
        let mut wb = a_model();
        wb.sheets[0].cells.get_mut(&(3, 3)).unwrap().style = 2;
        let library = crate::calc::library::language();
        let calc = crate::calc::Calc::run(&wb, &library);
        for (sheet, window) in [
            ("Model", "A1:F20"),
            ("Model", "C3:D4"),
            ("Model", "B1:B4"),
            ("Model", "H10:J12"),
            ("Empty", "A1"),
        ] {
            let w = Window {
                sheet: sheet.to_string(),
                range: range(window),
            };
            let bare = values_text(&wb, &w, None);
            assert_eq!(bare, less_formulas(&view_text(&wb, &w)), "{window}");
            let valued = view_text_values(&wb, &w, &calc);
            assert_eq!(
                valued,
                less_formulas(&view_text_valued(&wb, &w, &calc)),
                "{window}"
            );
            assert!(!valued.contains("(formula "), "{valued}");
        }
        let w = Window {
            sheet: "Model".to_string(),
            range: range("A1:F20"),
        };
        let bare = values_text(&wb, &w, None);
        assert!(bare.contains("(style \"Model\" \"C3\" 2)\n(sentence \"Model\" \"C3\" 3)\n"));
        assert!(!bare.contains("(value "), "{bare}");
        let valued = view_text_values(&wb, &w, &calc);
        assert!(
            valued.contains("(value \"Model\" \"C2\" 2401)\n(sentence \"Model\" \"C2\" 3)\n"),
            "{valued}"
        );
        // A window of values alone is the grid's own record.
        let w = Window {
            sheet: "Model".to_string(),
            range: range("B1:B4"),
        };
        assert_eq!(values_text(&wb, &w, None), view_text(&wb, &w));
        assert_eq!(ValueGrid::default().name(), "values");
    }

    #[test]
    fn a_fill_s_cell_whose_master_is_lost_shows_in_neither_record() {
        // KERNEL.37: a value over a fill's first cell, as the core's build
        // writes one, leaves the fill's other cells children of a master the
        // sheet no longer holds. Neither record shows them, not even their
        // style and sentence rows, and with no value to vouch for them the
        // values find that out by one walk of the sheet.
        let mut wb = Workbook::new();
        let mut sheet = Sheet::new("S");
        sheet.set(
            2,
            2,
            Cell {
                content: Content::Number(5.0),
                style: 0,
            },
        );
        sheet.set_formula(range("C2:C4"), "B2+1", 1);
        sheet.wrote(range("C2:C4"), 2);
        sheet.set(
            2,
            3,
            Cell {
                content: Content::Number(9.0),
                style: 0,
            },
        );
        sheet.wrote(range("C2"), 3);
        wb.sheets.push(sheet);
        let w = Window {
            sheet: "S".to_string(),
            range: range("A1:D5"),
        };
        let library = crate::calc::library::language();
        let calc = crate::calc::Calc::run(&wb, &library);
        let grid = view_text(&wb, &w);
        let valued = view_text_valued(&wb, &w, &calc);
        let mut bare = String::new();
        let work = walks(|| bare = values_text(&wb, &w, None));
        assert_eq!(bare, less_formulas(&grid));
        assert_eq!(work.walked, 4, "the sheet's four cells walked once");
        let mut values = String::new();
        let work = walks(|| values = view_text_values(&wb, &w, &calc));
        assert_eq!(values, less_formulas(&valued));
        assert_eq!(work.walked, 4, "no value vouches for C3, so one walk");
        assert_eq!(work.extents, 0, "the extent the calculation keeps");
        for record in [&grid, &valued, &bare, &values] {
            assert!(
                !record.contains("\"C3\"") && !record.contains("\"C4\""),
                "{record}"
            );
            assert!(record.contains("(cell \"S\" \"C2\" 9)\n(sentence \"S\" \"C2\" 3)\n"));
        }
    }

    #[test]
    fn the_values_of_a_computed_fill_walk_nothing() {
        // KERNEL.37: where every cell of a fill has a value, the values walk
        // the sheet neither for its masters nor for its extent; the grid
        // walks it once for its masters, to move their texts; with no
        // values, both walk it for the extent.
        let mut wb = Workbook::new();
        let mut sheet = Sheet::new("S");
        for row in 1..=4 {
            sheet.set(
                row,
                1,
                Cell {
                    content: Content::Number(f64::from(row)),
                    style: 0,
                },
            );
        }
        sheet.set_formula(range("B1:C4"), "A1*2", 0);
        wb.sheets.push(sheet);
        let w = Window {
            sheet: "S".to_string(),
            range: range("B2:C3"),
        };
        let library = crate::calc::library::language();
        let calc = crate::calc::Calc::run(&wb, &library);
        let work = walks(|| {
            view_text_values(&wb, &w, &calc);
        });
        assert_eq!((work.walked, work.extents), (0, 0));
        let work = walks(|| {
            view_text_valued(&wb, &w, &calc);
        });
        assert_eq!((work.walked, work.extents), (12, 0));
        let work = walks(|| {
            values_text(&wb, &w, None);
        });
        assert_eq!((work.walked, work.extents), (12, 12));
        let work = walks(|| {
            view_text(&wb, &w);
        });
        assert_eq!((work.walked, work.extents), (12, 12));
    }
}
