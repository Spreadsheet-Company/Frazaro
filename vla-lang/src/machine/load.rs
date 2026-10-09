//! The loader (KERNEL.22, 2026-10-09): rows to a grid, the inverse of the
//! printers in `crate::rows` and `crate::view`, so that what a door prints
//! the machine loads, and a sheet viewed whole and loaded again views the
//! same (`Alonzo/SPEC.md` section 7.3, and its free oracle in section 13).
//!
//! The rows taken: `sheet`, `cell` and `formula` over a cell or a range (a
//! value fills the range, a formula fills it as a shared formula, moved as
//! Excel moves a filled formula), `value` after its formula (the formula's
//! last value, never a second write), `name`, `gridlines`, `column`, `row`,
//! `format`, `look`, `style` and `sentence`. Skipped by name, being a
//! view's or a reader's own derivation or a thing no model holds before
//! `KERNEL.8`: `window`, `extent`, `refers` and `table`. A `sheet` row
//! naming `X.last` is taken as the twin's own declaration, which every
//! record of a save lists, when the rows hold `X`; the machine makes the
//! twins itself.
//!
//! One write per cell (decision 23 of the page): a cell the rows write
//! twice, by two rows or by a range and a row inside it, is refused naming
//! both lines; a setting given twice, as a save repeats a sheet row and
//! format 0 in every record it concatenates, is taken when it says the same
//! thing and refused when it does not. A grid holds at most [`CELLS`] cells
//! beside its twins, so that one range row cannot ask for the seventeen
//! billion of Excel's sheet.

use std::collections::HashMap;

use crate::calc::{CellId, Computed, ErrorKind, Reason, Value};
use crate::form::Form;
use crate::intrinsics::fold;
use crate::messages::raise;
use crate::printer::write_datum;
use crate::rows::{quoted, Visibility};
use crate::sheet::{
    number_text, parse_a1_range, A1Range, Cell, Column, Content, NumFmt, Rgb, RowSettings, Style,
    Workbook, MAX_COLUMN, MAX_ROW,
};

use super::{LineRefusal, CELLS};
use crate::calc::shape::spell_cell;

/// The rows a load takes, the four it skips last, as `grid-row-unknown`
/// lists them.
pub const LOAD_ROWS: &str = "sheet, cell, formula, value, name, gridlines, column, row, format, look, style, sentence, window, extent, refers and table";

/// The most formats a grid's table holds, Excel's own count of cell
/// formats a workbook may hold, rounded to a power of two.
pub const FORMATS: u32 = 65_536;

/// The grid as the rows wrote it, before a machine takes it: an engine
/// reads it between the two (its device layouts, its Clock), which is why
/// the load is two calls, [`read_rows`] and `Machine::new`.
#[derive(Clone, Debug)]
pub struct Draft {
    pub workbook: Workbook,
    /// Each formula cell's last value, from its `value` row.
    pub values: HashMap<CellId, Computed>,
    /// The line of the row that wrote each cell, for a refusal's line.
    pub lines: HashMap<CellId, u32>,
    /// How many cells the rows wrote.
    pub cells: usize,
}

/// A value as a row spells it.
#[derive(Clone, Debug, PartialEq)]
pub(crate) enum Datum {
    Number(f64),
    Text(String),
    Bool(bool),
    Error(String),
}

impl Datum {
    /// The cell content the value is.
    pub(crate) fn content(&self) -> Content {
        match self {
            Datum::Number(n) => Content::Number(*n),
            Datum::Text(t) => Content::Text(t.clone()),
            Datum::Bool(b) => Content::Bool(*b),
            Datum::Error(e) => Content::Error(e.clone()),
        }
    }

    /// The value spelled as the record spells a `cell` row's.
    fn spell(&self) -> String {
        match self {
            Datum::Number(n) => number_text(*n).unwrap_or_else(|| "0".to_string()),
            Datum::Text(t) => quoted(t),
            Datum::Bool(true) => "true".to_string(),
            Datum::Bool(false) => "false".to_string(),
            Datum::Error(e) => format!("(error {})", quoted(e)),
        }
    }
}

/// A value read off a row: a number in the record's spelling, a text, `true`
/// or `false`, or `(error "#N/A")`; `None` for anything else.
pub(crate) fn datum(f: &Form) -> Option<Datum> {
    match f {
        Form::Str(s) => Some(Datum::Text(s.clone())),
        Form::Sym(s) if s == "true" => Some(Datum::Bool(true)),
        Form::Sym(s) if s == "false" => Some(Datum::Bool(false)),
        Form::Sym(s) => number(s).map(Datum::Number),
        Form::List(l) => match l.items.as_slice() {
            [Form::Sym(head), Form::Str(e)] if head == "error" && !e.is_empty() => {
                Some(Datum::Error(e.clone()))
            }
            _ => None,
        },
    }
}

/// A number as a row writes it: an optional `-`, digits with at most one
/// `.`, an optional exponent; finite. Nothing else is read as one, so a
/// word such as `inf` is no number.
pub(crate) fn number(s: &str) -> Option<f64> {
    let b = s.as_bytes();
    let mut i = usize::from(b.first() == Some(&b'-'));
    let mut digits = 0;
    let mut dot = false;
    while i < b.len() {
        match b[i] {
            b'0'..=b'9' => digits += 1,
            b'.' if !dot => dot = true,
            _ => break,
        }
        i += 1;
    }
    if digits == 0 {
        return None;
    }
    if i < b.len() && (b[i] == b'e' || b[i] == b'E') {
        let mut j = i + 1;
        if j < b.len() && (b[j] == b'+' || b[j] == b'-') {
            j += 1;
        }
        let first = j;
        while j < b.len() && b[j].is_ascii_digit() {
            j += 1;
        }
        if j == first {
            return None;
        }
        i = j;
    }
    if i != b.len() {
        return None;
    }
    s.parse::<f64>().ok().filter(|v| v.is_finite())
}

/// Whether a name can name a sheet of a grid: 1 to 26 characters, so that
/// its twin's name fits Excel's 31; none of `\ / ? * [ ] :`; no apostrophe
/// at either end; and not ending in `.last`, the twins' own suffix.
pub fn sheet_name_ok(name: &str) -> bool {
    let n = name.chars().count();
    (1..=26).contains(&n)
        && !name
            .chars()
            .any(|c| matches!(c, '\\' | '/' | '?' | '*' | '[' | ']' | ':'))
        && !name.starts_with('\'')
        && !name.ends_with('\'')
        && !fold(name).ends_with(".last")
}

/// The sheet a twin's name is the twin of: `Screen` for `Screen.last`,
/// compared without case; `None` for any other name.
pub fn twin_base(name: &str) -> Option<&str> {
    let n = name.len();
    (n > 5 && name.is_char_boundary(n - 5) && fold(&name[n - 5..]) == ".last")
        .then(|| &name[..n - 5])
}

/// The shapes a row is held to, in the notation, for `grid-row-malformed`.
pub(crate) fn shape_of(kind: &str) -> &'static str {
    match kind {
        "sheet" => "(sheet \"<name>\" visible|hidden|very-hidden)",
        "cell" => "(cell \"<sheet>\" \"<cell or range>\" <value>)",
        "formula" => "(formula \"<sheet>\" \"<cell or range>\" \"=<text>\")",
        "derived" => "(derived \"<sheet>\" \"<cell or range>\" <value>)",
        "value" => {
            "(formula \"<sheet>\" \"<cell>\" \"=<text>\") (value \"<sheet>\" \"<cell>\" <value>)"
        }
        "name" => "(name \"<name>\" \"<refers-to>\")",
        "gridlines" => "(gridlines \"<sheet>\" on|off)",
        "column" => "(column \"<sheet>\" \"<letters>\" <width>|none shown|hidden <format>|none)",
        "row" => "(row \"<sheet>\" <number> <height>|none shown|hidden)",
        "format" => "(format <index> \"<RRGGBB>\"|none general|text wrap|nowrap)",
        "look" => "(look \"<sheet>\" <value> <format>)",
        "style" => "(cell \"<sheet>\" \"<cell>\" <value>) (style \"<sheet>\" \"<cell>\" <format>)",
        "sentence" => {
            "(cell \"<sheet>\" \"<cell>\" <value>) (sentence \"<sheet>\" \"<cell>\" <row>)"
        }
        _ => "(<kind> ...)",
    }
}

/// `grid-row-malformed` for a form.
pub(crate) fn malformed(line: u32, form: &Form, kind: &str) -> LineRefusal {
    LineRefusal {
        line,
        refusal: raise(
            "grid-row-malformed",
            &[
                ("line", &line.to_string()),
                ("row", &write_datum(form)),
                ("shape", shape_of(kind)),
            ],
        ),
    }
}

/// `grid-row-unknown` for a row whose head is not one taken.
pub(crate) fn unknown(line: u32, head: &str, kinds: &str) -> LineRefusal {
    LineRefusal {
        line,
        refusal: raise(
            "grid-row-unknown",
            &[
                ("line", &line.to_string()),
                ("head", head),
                ("kinds", kinds),
            ],
        ),
    }
}

/// `grid-cell-written-twice`: `what` written at `line` and first at `first`.
pub(crate) fn twice(line: u32, what: &str, first: u32) -> LineRefusal {
    LineRefusal {
        line,
        refusal: raise(
            "grid-cell-written-twice",
            &[
                ("line", &line.to_string()),
                ("what", what),
                ("first", &first.to_string()),
            ],
        ),
    }
}

/// `grid-sheet-name-invalid` for a name a grid cannot hold.
pub(crate) fn bad_name(line: u32, name: &str) -> LineRefusal {
    LineRefusal {
        line,
        refusal: raise(
            "grid-sheet-name-invalid",
            &[("line", &line.to_string()), ("name", name)],
        ),
    }
}

/// `grid-write-last`: a row writing into a twin.
pub(crate) fn into_twin(line: u32, sheet: &str, addr: &str) -> LineRefusal {
    let source = twin_base(sheet).unwrap_or(sheet);
    LineRefusal {
        line,
        refusal: raise(
            "grid-write-last",
            &[
                ("line", &line.to_string()),
                ("target", &format!("{sheet}!{addr}")),
                ("sheet", sheet),
                ("source", source),
            ],
        ),
    }
}

/// `grid-too-many-cells`: a row whose range would bring the grid past the cap.
pub(crate) fn too_many(line: u32, sheet: &str, addr: &str, count: u64) -> LineRefusal {
    LineRefusal {
        line,
        refusal: raise(
            "grid-too-many-cells",
            &[
                ("line", &line.to_string()),
                ("target", &format!("{sheet}!{addr}")),
                ("cells", &thousands(count)),
                ("limit", &thousands(CELLS as u64)),
            ],
        ),
    }
}

/// A count with its thousands separated by commas, as the host's progress
/// sentence spells one.
pub fn thousands(n: u64) -> String {
    let digits = n.to_string();
    let mut out = String::with_capacity(digits.len() + digits.len() / 3);
    for (i, c) in digits.chars().enumerate() {
        if i > 0 && (digits.len() - i) % 3 == 0 {
            out.push(',');
        }
        out.push(c);
    }
    out
}

/// A row's head and items, or `None` for a form that is no row.
pub(crate) fn row_parts(form: &Form) -> Option<(&str, &[Form], u32)> {
    let Form::List(l) = form else { return None };
    match l.items.first() {
        Some(Form::Sym(head)) => Some((head.as_str(), &l.items[1..], l.line)),
        _ => None,
    }
}

/// The loader's state over one text of rows.
struct Reader {
    draft: Draft,
    sheet_rows: HashMap<usize, (Visibility, u32)>,
    twin_rows: Vec<(String, u32)>,
    names: HashMap<String, (String, u32)>,
    gridlines: HashMap<usize, (bool, u32)>,
    columns: HashMap<(usize, u32), u32>,
    row_settings: HashMap<(usize, u32), u32>,
    formats: HashMap<u32, (Style, u32)>,
    looks: HashMap<(usize, String), (u32, u32)>,
    styles: HashMap<CellId, (u32, u32)>,
    sentences: HashMap<CellId, (u32, u32)>,
    value_lines: HashMap<CellId, u32>,
}

/// The forms of a text of rows read into a grid, the inverse of the record:
/// every row taken or skipped by name, or the first refusal with its line.
/// Nothing is computed and no twin is made: that is `Machine::new`'s.
pub fn read_rows(forms: &[Form]) -> Result<Draft, LineRefusal> {
    let mut r = Reader {
        draft: Draft {
            workbook: Workbook::new(),
            values: HashMap::new(),
            lines: HashMap::new(),
            cells: 0,
        },
        sheet_rows: HashMap::new(),
        twin_rows: Vec::new(),
        names: HashMap::new(),
        gridlines: HashMap::new(),
        columns: HashMap::new(),
        row_settings: HashMap::new(),
        formats: HashMap::new(),
        looks: HashMap::new(),
        styles: HashMap::new(),
        sentences: HashMap::new(),
        value_lines: HashMap::new(),
    };
    for form in forms {
        r.row(form)?;
    }
    // A twin's declaration stands only beside the sheet it is the twin of.
    for (name, line) in &r.twin_rows {
        let base = twin_base(name).unwrap_or("");
        if r.draft.workbook.find_sheet(base).is_none() {
            return Err(bad_name(*line, name));
        }
    }
    Ok(r.draft)
}

impl Reader {
    fn row(&mut self, form: &Form) -> Result<(), LineRefusal> {
        let Some((head, items, line)) = row_parts(form) else {
            let line = form.as_list().map(|l| l.line).unwrap_or(0);
            return Err(malformed(line, form, ""));
        };
        match head {
            "sheet" => self.sheet(form, items, line),
            "cell" => self.cell(form, items, line),
            "formula" => self.formula(form, items, line),
            "value" => self.value(form, items, line),
            "name" => self.name(form, items, line),
            "gridlines" => self.gridlines(form, items, line),
            "column" => self.column(form, items, line),
            "row" => self.row_settings(form, items, line),
            "format" => self.format(form, items, line),
            "look" => self.look(form, items, line),
            "style" => self.style(form, items, line),
            "sentence" => self.sentence(form, items, line),
            "window" | "extent" | "refers" | "table" => Ok(()),
            other => Err(unknown(line, other, LOAD_ROWS)),
        }
    }

    /// A sheet of the grid by name, made when it is new; a twin's name and a
    /// name no sheet may hold refused.
    fn sheet_index(&mut self, name: &str, line: u32, addr: &str) -> Result<usize, LineRefusal> {
        if let Some(i) = self.draft.workbook.find_sheet(name) {
            return Ok(i);
        }
        if twin_base(name).is_some() {
            return Err(into_twin(line, name, addr));
        }
        if !sheet_name_ok(name) {
            return Err(bad_name(line, name));
        }
        Ok(self.draft.workbook.ensure_sheet(name))
    }

    fn sheet(&mut self, form: &Form, items: &[Form], line: u32) -> Result<(), LineRefusal> {
        let [Form::Str(name), Form::Sym(word)] = items else {
            return Err(malformed(line, form, "sheet"));
        };
        let visibility = match word.as_str() {
            "visible" => Visibility::Visible,
            "hidden" => Visibility::Hidden,
            "very-hidden" => Visibility::VeryHidden,
            _ => return Err(malformed(line, form, "sheet")),
        };
        if twin_base(name).is_some() {
            // A twin's own declaration, as every record lists the twins.
            if visibility != Visibility::Hidden {
                return Err(bad_name(line, name));
            }
            self.twin_rows.push((name.clone(), line));
            return Ok(());
        }
        let i = self.sheet_index(name, line, "A1")?;
        if let Some((v, first)) = self.sheet_rows.get(&i) {
            if *v != visibility {
                return Err(twice(line, &format!("(sheet {})", quoted(name)), *first));
            }
            return Ok(());
        }
        self.sheet_rows.insert(i, (visibility, line));
        self.draft.workbook.sheets[i].visibility = visibility;
        Ok(())
    }

    /// A row's sheet and its rectangle.
    fn target(
        &mut self,
        form: &Form,
        kind: &str,
        sheet: &str,
        addr: &str,
        line: u32,
    ) -> Result<(usize, A1Range), LineRefusal> {
        let range = parse_a1_range(addr).ok_or_else(|| malformed(line, form, kind))?;
        let i = self.sheet_index(sheet, line, addr)?;
        Ok((i, range))
    }

    /// The cells a range row writes, refused when one was written before or
    /// when they would bring the grid past its cap.
    fn claim(
        &mut self,
        si: usize,
        range: A1Range,
        sheet: &str,
        addr: &str,
        line: u32,
    ) -> Result<(), LineRefusal> {
        let count = range.cells();
        if self.draft.cells as u64 + count > CELLS as u64 {
            return Err(too_many(line, sheet, addr, self.draft.cells as u64 + count));
        }
        for row in range.top..=range.bottom {
            for col in range.left..=range.right {
                if let Some(first) = self.draft.lines.get(&(si, row, col)) {
                    let what = spell_cell(&self.draft.workbook, (si, row, col));
                    return Err(twice(line, &what, *first));
                }
            }
        }
        for row in range.top..=range.bottom {
            for col in range.left..=range.right {
                self.draft.lines.insert((si, row, col), line);
            }
        }
        self.draft.cells += count as usize;
        Ok(())
    }

    fn cell(&mut self, form: &Form, items: &[Form], line: u32) -> Result<(), LineRefusal> {
        let [Form::Str(sheet), Form::Str(addr), v] = items else {
            return Err(malformed(line, form, "cell"));
        };
        let value = datum(v).ok_or_else(|| malformed(line, form, "cell"))?;
        let (si, range) = self.target(form, "cell", sheet, addr, line)?;
        self.claim(si, range, sheet, addr, line)?;
        let s = &mut self.draft.workbook.sheets[si];
        for row in range.top..=range.bottom {
            for col in range.left..=range.right {
                s.set(
                    row,
                    col,
                    Cell {
                        content: value.content(),
                        style: 0,
                    },
                );
            }
        }
        Ok(())
    }

    fn formula(&mut self, form: &Form, items: &[Form], line: u32) -> Result<(), LineRefusal> {
        let [Form::Str(sheet), Form::Str(addr), Form::Str(text)] = items else {
            return Err(malformed(line, form, "formula"));
        };
        let Some(body) = text.strip_prefix('=').filter(|b| !b.is_empty()) else {
            return Err(malformed(line, form, "formula"));
        };
        let (si, range) = self.target(form, "formula", sheet, addr, line)?;
        self.claim(si, range, sheet, addr, line)?;
        self.draft.workbook.sheets[si].set_formula(range, body, 0);
        Ok(())
    }

    fn value(&mut self, form: &Form, items: &[Form], line: u32) -> Result<(), LineRefusal> {
        let [Form::Str(sheet), Form::Str(addr), v] = items else {
            return Err(malformed(line, form, "value"));
        };
        let computed = computed(v).ok_or_else(|| malformed(line, form, "value"))?;
        let Some(range) = parse_a1_range(addr).filter(|r| r.cells() == 1) else {
            return Err(malformed(line, form, "value"));
        };
        let Some(si) = self.draft.workbook.find_sheet(sheet) else {
            return Err(malformed(line, form, "value"));
        };
        let id = (si, range.top, range.left);
        let is_formula = self.draft.workbook.sheets[si]
            .cells
            .get(&(range.top, range.left))
            .is_some_and(|c| crate::calc::graph::is_formula(&c.content));
        if !is_formula {
            return Err(malformed(line, form, "value"));
        }
        if let Some(first) = self.value_lines.get(&id) {
            let what = format!("(value {} {})", quoted(sheet), quoted(addr));
            return Err(twice(line, &what, *first));
        }
        self.value_lines.insert(id, line);
        self.draft.values.insert(id, computed);
        Ok(())
    }

    fn name(&mut self, form: &Form, items: &[Form], line: u32) -> Result<(), LineRefusal> {
        let [Form::Str(name), Form::Str(text)] = items else {
            return Err(malformed(line, form, "name"));
        };
        if name.is_empty() || text.is_empty() {
            return Err(malformed(line, form, "name"));
        }
        let key = fold(name);
        if let Some((said, first)) = self.names.get(&key) {
            if said != text {
                return Err(twice(line, &format!("(name {})", quoted(name)), *first));
            }
            return Ok(());
        }
        self.names.insert(key, (text.clone(), line));
        self.draft
            .workbook
            .defined_names
            .push((name.clone(), text.clone()));
        Ok(())
    }

    fn gridlines(&mut self, form: &Form, items: &[Form], line: u32) -> Result<(), LineRefusal> {
        let [Form::Str(sheet), Form::Sym(word)] = items else {
            return Err(malformed(line, form, "gridlines"));
        };
        let on = match word.as_str() {
            "on" => true,
            "off" => false,
            _ => return Err(malformed(line, form, "gridlines")),
        };
        let si = self.sheet_index(sheet, line, "A1")?;
        if let Some((said, first)) = self.gridlines.get(&si) {
            if *said != on {
                return Err(twice(
                    line,
                    &format!("(gridlines {})", quoted(sheet)),
                    *first,
                ));
            }
            return Ok(());
        }
        self.gridlines.insert(si, (on, line));
        self.draft.workbook.sheets[si].gridlines = on;
        Ok(())
    }

    fn column(&mut self, form: &Form, items: &[Form], line: u32) -> Result<(), LineRefusal> {
        let [Form::Str(sheet), Form::Str(letters), width, Form::Sym(shown), style] = items else {
            return Err(malformed(line, form, "column"));
        };
        let Some(index) = column_index(letters) else {
            return Err(malformed(line, form, "column"));
        };
        let width = match width {
            Form::Sym(w) if w == "none" => None,
            Form::Sym(w) => match number(w) {
                Some(n) if n >= 0.0 => Some(n),
                _ => return Err(malformed(line, form, "column")),
            },
            _ => return Err(malformed(line, form, "column")),
        };
        let hidden = match shown.as_str() {
            "shown" => false,
            "hidden" => true,
            _ => return Err(malformed(line, form, "column")),
        };
        let style = match style {
            Form::Sym(s) if s == "none" => None,
            Form::Sym(s) => match s.parse::<u32>() {
                Ok(n) => Some(n),
                Err(_) => return Err(malformed(line, form, "column")),
            },
            _ => return Err(malformed(line, form, "column")),
        };
        let si = self.sheet_index(sheet, line, "A1")?;
        let column = Column {
            index,
            width,
            hidden,
            style,
        };
        if let Some(first) = self.columns.get(&(si, index)) {
            let said = self.draft.workbook.sheets[si]
                .columns
                .iter()
                .find(|c| c.index == index);
            if said != Some(&column) {
                let what = format!("(column {} {})", quoted(sheet), quoted(letters));
                return Err(twice(line, &what, *first));
            }
            return Ok(());
        }
        self.columns.insert((si, index), line);
        self.draft.workbook.sheets[si].columns.push(column);
        Ok(())
    }

    fn row_settings(&mut self, form: &Form, items: &[Form], line: u32) -> Result<(), LineRefusal> {
        let [Form::Str(sheet), Form::Sym(n), height, Form::Sym(shown)] = items else {
            return Err(malformed(line, form, "row"));
        };
        let n = match n.parse::<u32>() {
            Ok(n) if (1..=MAX_ROW).contains(&n) => n,
            _ => return Err(malformed(line, form, "row")),
        };
        let height = match height {
            Form::Sym(h) if h == "none" => None,
            Form::Sym(h) => match number(h) {
                Some(v) if v >= 0.0 => Some(v),
                _ => return Err(malformed(line, form, "row")),
            },
            _ => return Err(malformed(line, form, "row")),
        };
        let hidden = match shown.as_str() {
            "shown" => false,
            "hidden" => true,
            _ => return Err(malformed(line, form, "row")),
        };
        let si = self.sheet_index(sheet, line, "A1")?;
        let settings = RowSettings { height, hidden };
        if let Some(first) = self.row_settings.get(&(si, n)) {
            if self.draft.workbook.sheets[si].rows.get(&n) != Some(&settings) {
                return Err(twice(line, &format!("(row {} {n})", quoted(sheet)), *first));
            }
            return Ok(());
        }
        self.row_settings.insert((si, n), line);
        self.draft.workbook.sheets[si].rows.insert(n, settings);
        Ok(())
    }

    fn format(&mut self, form: &Form, items: &[Form], line: u32) -> Result<(), LineRefusal> {
        let [Form::Sym(index), fill, Form::Sym(num), Form::Sym(wrap)] = items else {
            return Err(malformed(line, form, "format"));
        };
        let index = match index.parse::<u32>() {
            Ok(i) if i < FORMATS => i,
            _ => return Err(malformed(line, form, "format")),
        };
        let fill = match fill {
            Form::Sym(f) if f == "none" => None,
            Form::Str(hex) => match rgb(hex) {
                Some(c) => Some(c),
                None => return Err(malformed(line, form, "format")),
            },
            _ => return Err(malformed(line, form, "format")),
        };
        let num_fmt = match num.as_str() {
            "general" => NumFmt::General,
            "text" => NumFmt::Text,
            _ => return Err(malformed(line, form, "format")),
        };
        let wrap = match wrap.as_str() {
            "wrap" => true,
            "nowrap" => false,
            _ => return Err(malformed(line, form, "format")),
        };
        let style = Style {
            fill,
            num_fmt,
            wrap,
        };
        if let Some((said, first)) = self.formats.get(&index) {
            if *said != style {
                return Err(twice(line, &format!("(format {index})"), *first));
            }
            return Ok(());
        }
        self.formats.insert(index, (style.clone(), line));
        self.draft.workbook.styles.put(index, style);
        Ok(())
    }

    fn look(&mut self, form: &Form, items: &[Form], line: u32) -> Result<(), LineRefusal> {
        let [Form::Str(sheet), v, Form::Sym(index)] = items else {
            return Err(malformed(line, form, "look"));
        };
        let value = datum(v).ok_or_else(|| malformed(line, form, "look"))?;
        let index = match index.parse::<u32>() {
            Ok(i) if i < FORMATS => i,
            _ => return Err(malformed(line, form, "look")),
        };
        let si = self.sheet_index(sheet, line, "A1")?;
        let spelled = value.spell();
        let key = (si, spelled.clone());
        if let Some((said, first)) = self.looks.get(&key) {
            if *said != index {
                return Err(twice(
                    line,
                    &format!("(look {} {spelled})", quoted(sheet)),
                    *first,
                ));
            }
            return Ok(());
        }
        self.looks.insert(key, (index, line));
        self.draft.workbook.sheets[si].looks.push((spelled, index));
        Ok(())
    }

    /// The cell a `style` or `sentence` row names, which an earlier row wrote.
    fn written_cell(
        &self,
        form: &Form,
        kind: &str,
        sheet: &str,
        addr: &str,
        line: u32,
    ) -> Result<CellId, LineRefusal> {
        let range = parse_a1_range(addr)
            .filter(|r| r.cells() == 1)
            .ok_or_else(|| malformed(line, form, kind))?;
        let si = self
            .draft
            .workbook
            .find_sheet(sheet)
            .ok_or_else(|| malformed(line, form, kind))?;
        if !self.draft.workbook.sheets[si]
            .cells
            .contains_key(&(range.top, range.left))
        {
            return Err(malformed(line, form, kind));
        }
        Ok((si, range.top, range.left))
    }

    fn style(&mut self, form: &Form, items: &[Form], line: u32) -> Result<(), LineRefusal> {
        let [Form::Str(sheet), Form::Str(addr), Form::Sym(index)] = items else {
            return Err(malformed(line, form, "style"));
        };
        let index = match index.parse::<u32>() {
            Ok(i) if i < FORMATS => i,
            _ => return Err(malformed(line, form, "style")),
        };
        let id = self.written_cell(form, "style", sheet, addr, line)?;
        if let Some((said, first)) = self.styles.get(&id) {
            if *said != index {
                let what = format!("(style {} {})", quoted(sheet), quoted(addr));
                return Err(twice(line, &what, *first));
            }
            return Ok(());
        }
        self.styles.insert(id, (index, line));
        if let Some(cell) = self.draft.workbook.sheets[id.0]
            .cells
            .get_mut(&(id.1, id.2))
        {
            cell.style = index;
        }
        Ok(())
    }

    fn sentence(&mut self, form: &Form, items: &[Form], line: u32) -> Result<(), LineRefusal> {
        let [Form::Str(sheet), Form::Str(addr), Form::Sym(n)] = items else {
            return Err(malformed(line, form, "sentence"));
        };
        let n = match n.parse::<u32>() {
            Ok(n) if n > 0 => n,
            _ => return Err(malformed(line, form, "sentence")),
        };
        let id = self.written_cell(form, "sentence", sheet, addr, line)?;
        if let Some((said, first)) = self.sentences.get(&id) {
            if *said != n {
                let what = format!("(sentence {} {})", quoted(sheet), quoted(addr));
                return Err(twice(line, &what, *first));
            }
            return Ok(());
        }
        self.sentences.insert(id, (n, line));
        self.draft.workbook.sheets[id.0]
            .sentences
            .insert((id.1, id.2), n);
        Ok(())
    }
}

/// A `value` row's value: a value, or the label of a formula not computed.
fn computed(f: &Form) -> Option<Computed> {
    if let Form::List(l) = f {
        match l.items.as_slice() {
            [Form::Sym(head), Form::Str(name)] if head == "not-computed" => {
                return Some(Computed::NotComputed(Reason::Construct(name.clone())));
            }
            [Form::Sym(head), Form::Sym(word)] if head == "not-computed" && word == "cycle" => {
                return Some(Computed::NotComputed(Reason::Cycle));
            }
            _ => {}
        }
    }
    Some(Computed::Value(match datum(f)? {
        Datum::Number(n) => Value::Number(n),
        Datum::Text(t) => Value::Text(t),
        Datum::Bool(b) => Value::Bool(b),
        Datum::Error(e) => Value::Error(ErrorKind::parse(&e)?),
    }))
}

/// A column's letters as its index, `A` to `XFD`.
fn column_index(letters: &str) -> Option<u32> {
    if letters.is_empty() || letters.len() > 3 || !letters.chars().all(|c| c.is_ascii_alphabetic())
    {
        return None;
    }
    let mut n: u32 = 0;
    for c in letters.chars() {
        n = n * 26 + (c.to_ascii_uppercase() as u32 - 'A' as u32 + 1);
    }
    (1..=MAX_COLUMN).contains(&n).then_some(n)
}

/// Six hex digits as a colour.
fn rgb(hex: &str) -> Option<Rgb> {
    if hex.len() != 6 || !hex.chars().all(|c| c.is_ascii_hexdigit()) {
        return None;
    }
    let byte = |i: usize| u8::from_str_radix(&hex[i..i + 2], 16).ok();
    Some(Rgb(byte(0)?, byte(2)?, byte(4)?))
}
