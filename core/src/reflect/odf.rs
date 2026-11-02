//! The OpenDocument reader (PORT.8, slice 8e): an `.ods` file opened into
//! the same [`Source`] as an OOXML package, so that `frazaro reflect`,
//! `diff`, `audit` and `--cone` read either format and print the same
//! rows in the same spelling.
//!
//! What is read, and from where. The `mimetype` entry says the file is a
//! spreadsheet (another OpenDocument kind is refused by name). Everything
//! else comes from `content.xml`, held whole, since it holds every sheet:
//! the automatic styles, for the table style whose `display` is `false`
//! (a hidden sheet; OpenDocument has no very-hidden); each `table:table`,
//! a sheet, remembered as its slice of the part so that a sheet is walked
//! from its own slice; the named expressions, workbook-wide and
//! sheet-local, as `name` rows; the database ranges, Calc's named
//! rectangles with a header row, as `table` rows (the nearest kin of a
//! Table; the anonymous ones autofilters leave are not); and, at open, the
//! other files the formulas reach, for `--counts`. The relations decode no
//! style, setting, manifest or metadata; under `--looks` (KERNEL.57) the
//! sheet is printed as the record prints it whole, its styles, its data
//! styles and its gridlines read from `content.xml`, `styles.xml` and
//! `settings.xml` (`odf_looks.rs`), a sheet's part walked twice, its first
//! pass for its columns, its rows of settings, its extent and the formats
//! its cells use.
//!
//! A cell's value is what the file holds: `office:value` as written for a
//! float, a percentage or a currency; a boolean; a date or a time under
//! `date`, a time as its ISO duration; a string from its paragraphs, each
//! `text:p` a line, a `text:s` its spaces, with annotations and drawings
//! passed over; an error where Calc's `calcext:value-type` says so, or,
//! lacking that marker, where a formula cell's string is one of Excel's
//! seven error literals, the gap the format leaves. A repeated row or cell
//! is read once per repetition, which is what the file says; a matrix
//! formula is its anchor's, an array anchor, and the cells it spills into
//! are values.
//!
//! A formula's text is OpenFormula (`of:=SUM([.B1:.B2];[Data.A1])`) and is
//! printed as the formula bar shows it, in Excel and in Calc alike
//! (`=SUM(B1:B2,Data!A1)`): the namespace prefix dropped; each bracketed
//! reference spelled in A1, its sheet quoted by the one rule, a sheet
//! marked absolute with `$` unmarked, two sheets as a 3D span, a source
//! (`'file:///…/Rates.xlsx'#`) as the file's name in brackets; `;`
//! between arguments as `,`; the reference operators `~` and `!` as `,`
//! and a space; `COM.MICROSOFT.` dropped from a function's name. So the
//! one reference reader, `crate::refers`, reads both formats, and the
//! goldens keep one spelling.

use std::cell::{OnceCell, RefCell};
use std::collections::{HashMap, HashSet};

use super::cursor::{
    attr, attr_full, attributes_full, skip_element, unescape_entities, Cursor, Event,
};
use super::looks::{
    iso_duration_days, iso_serial, ColumnSpec, Look, LooksStats, RuleGate, SheetScan,
};
use super::odf_looks::{self, OdsLooks};
use super::ooxml::{not_a_workbook, step, xml_refused};
use super::{
    emit_formula, emit_formula_with, NameInfo, Record, Row, SheetInfo, SheetStats, Sink, Source,
    Summary, TableInfo, Value, Visibility,
};
use crate::intrinsics::fold;
use crate::messages::Refusal;
use crate::refers;
use crate::sheet::merge::part_text_raw;
use crate::sheet::zip;
use crate::sheet::{
    cell_ref, number_text, parse_a1_range, A1Range, RowSettings, Style, MAX_COLUMN, MAX_ROW,
};
use vla_lang::machine::load::FORMATS;
use vla_lang::rows::{datum, ArrayKind};
use vla_lang::rule::{Operator, Rule, Test};

const MIMETYPE: &str = "application/vnd.oasis.opendocument.spreadsheet";
const CONTENT: &str = "content.xml";

/// Whether these bytes are an OpenDocument file: a zip whose `mimetype`
/// entry says so, or one with a `content.xml` and no workbook part.
pub fn is_opendocument(bytes: &[u8]) -> bool {
    let Some(entries) = zip::entries(bytes) else {
        return false;
    };
    if let Some(e) = entries.iter().find(|e| e.name == "mimetype") {
        if let Some(data) = zip::stored_data(bytes, e) {
            return data.starts_with(b"application/vnd.oasis.opendocument");
        }
        return matches!(
            part_text_raw(bytes, &entries, "mimetype"),
            Ok(Some(ref t)) if t.starts_with("application/vnd.oasis.opendocument")
        );
    }
    let has = |n: &str| entries.iter().any(|e| e.name == n);
    has(CONTENT) && !has("xl/workbook.xml")
}

/// An OpenDocument spreadsheet opened far enough to walk.
pub struct Document {
    label: String,
    /// `content.xml`, inflated.
    text: String,
    sheets: Vec<SheetInfo>,
    /// Each sheet's slice of the text, from its start tag to its end tag.
    spans: Vec<(usize, usize)>,
    names: Vec<NameInfo>,
    tables: Vec<TableInfo>,
    books: Vec<String>,
    /// `styles.xml` and `settings.xml`, inflated, where the file has them,
    /// read for the looks alone (KERNEL.57).
    styles: Option<String>,
    settings: Option<String>,
    /// Whether the document counts its dates from 1904, its null date's.
    date1904: bool,
    /// The styles as the looks read them, the first time a walk asks.
    looks: OnceCell<Result<OdsLooks, Refusal>>,
    /// The formats the rules name, each different one once, past the cell formats,
    /// in the order a walk first meets it (KERNEL.58).
    rule_looks: RefCell<Vec<Look>>,
    /// What a walk asks of a rule, made for the first rule it meets.
    rule_gate: OnceCell<RuleGate>,
}

impl std::fmt::Debug for Document {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("Document")
            .field("label", &self.label)
            .field("sheets", &self.sheets)
            .field("names", &self.names)
            .field("tables", &self.tables)
            .field("books", &self.books)
            .finish()
    }
}

/// What the structure pass found before any sheet is walked.
struct Structure {
    /// Each sheet's name, table style and slice.
    sheets: Vec<(String, Option<String>, usize, usize)>,
    hidden_styles: HashSet<String>,
    names: Vec<NameInfo>,
    /// Each database range's name and target, `Data.A1:Data.B4`.
    ranges: Vec<(String, String)>,
    books: Vec<String>,
    /// The null date, `table:null-date`, where the document says one.
    null_date: Option<String>,
}

impl Document {
    /// The file read far enough to walk; `label` is what a refusal calls
    /// the file.
    pub fn open(bytes: &[u8], label: &str) -> Result<Document, Refusal> {
        let entries = zip::entries(bytes).ok_or_else(|| {
            not_a_workbook(
                label,
                zip::why_not_an_archive(bytes)
                    .unwrap_or("it is not a zip archive this reader knows"),
            )
        })?;
        let text_of = |name: &str| -> Result<Option<String>, Refusal> {
            part_text_raw(bytes, &entries, name).map_err(|why| not_a_workbook(label, &why))
        };
        if let Some(mime) = text_of("mimetype")? {
            let mime = mime.trim();
            if mime != MIMETYPE {
                let kind = mime.rsplit('.').next().unwrap_or("file");
                return Err(not_a_workbook(
                    label,
                    &format!("it is an OpenDocument {kind} file, not a spreadsheet"),
                ));
            }
        }
        let text = text_of(CONTENT)?.ok_or_else(|| {
            not_a_workbook(
                label,
                "it has no content.xml part, so it is not an OpenDocument spreadsheet",
            )
        })?;
        let found = structure(&text, label)?;
        let sheets: Vec<SheetInfo> = found
            .sheets
            .iter()
            .map(|(name, style, _, _)| SheetInfo {
                name: name.clone(),
                visibility: match style {
                    Some(s) if found.hidden_styles.contains(s) => Visibility::Hidden,
                    _ => Visibility::Visible,
                },
                part: CONTENT.to_string(),
            })
            .collect();
        let spans: Vec<(usize, usize)> = found.sheets.iter().map(|s| (s.2, s.3)).collect();
        let mut tables: Vec<TableInfo> = Vec::new();
        for (name, target) in &found.ranges {
            let Some((first, second)) = split_range(target) else {
                continue;
            };
            let (sheet_name, from) = parse_part(first);
            let (_, to) = parse_part(second);
            let wanted = fold(&sheet_name);
            let Some(sheet) = sheets.iter().position(|s| fold(&s.name) == wanted) else {
                continue;
            };
            let range = if to.is_empty() || to == from {
                from.to_string()
            } else {
                format!("{from}:{to}")
            };
            tables.push(TableInfo {
                name: name.clone(),
                sheet,
                range,
                columns: Vec::new(),
            });
        }
        Ok(Document {
            label: label.to_string(),
            text,
            sheets,
            spans,
            names: found.names,
            tables,
            books: found.books,
            styles: text_of("styles.xml")?,
            settings: text_of("settings.xml")?,
            date1904: found
                .null_date
                .as_deref()
                .is_some_and(|d| d.starts_with("1904")),
            looks: OnceCell::new(),
            rule_looks: RefCell::new(Vec::new()),
            rule_gate: OnceCell::new(),
        })
    }
}

/// One pass over `content.xml` for its structure: the hidden table styles,
/// each sheet's name, style and slice, the named expressions (sheet-local
/// inside a table, workbook-wide outside), the database ranges, and the
/// files the formulas reach.
fn structure(text: &str, label: &str) -> Result<Structure, Refusal> {
    let part = CONTENT;
    let mut found = Structure {
        sheets: Vec::new(),
        hidden_styles: HashSet::new(),
        names: Vec::new(),
        ranges: Vec::new(),
        books: Vec::new(),
        null_date: None,
    };
    let mut cursor = Cursor::new(text);
    let mut in_styles = false;
    let mut current_style: Option<String> = None;
    let mut table_depth = 0usize;
    let mut current_sheet: Option<usize> = None;
    loop {
        let before = cursor.position();
        let Some(event) = step(&mut cursor, label, part)? else {
            break;
        };
        match event {
            Event::Start {
                name: "automatic-styles",
                empty: false,
                ..
            } => in_styles = true,
            Event::End {
                name: "automatic-styles",
            } => in_styles = false,
            Event::Start {
                name: "style",
                attrs,
                empty,
            } if in_styles => {
                if !empty && attr(attrs, "family").as_deref() == Some("table") {
                    current_style = attr(attrs, "name").map(|c| c.into_owned());
                }
            }
            Event::End { name: "style" } => current_style = None,
            Event::Start {
                name: "table-properties",
                attrs,
                ..
            } => {
                if let Some(style) = &current_style {
                    if attr(attrs, "display").as_deref() == Some("false") {
                        found.hidden_styles.insert(style.clone());
                    }
                }
            }
            Event::Start {
                name: "table",
                attrs,
                empty,
            } => {
                if table_depth == 0 {
                    let name = attr(attrs, "name")
                        .map(|c| c.into_owned())
                        .unwrap_or_else(|| format!("Sheet{}", found.sheets.len() + 1));
                    let style = attr(attrs, "style-name").map(|c| c.into_owned());
                    found.sheets.push((name, style, before, cursor.position()));
                    current_sheet = Some(found.sheets.len() - 1);
                    if empty {
                        current_sheet = None;
                        continue;
                    }
                }
                if !empty {
                    table_depth += 1;
                }
            }
            Event::End { name: "table" } => {
                table_depth = table_depth.saturating_sub(1);
                if table_depth == 0 {
                    if let Some(i) = current_sheet {
                        found.sheets[i].3 = cursor.position();
                    }
                    current_sheet = None;
                }
            }
            Event::Start {
                name: "named-range",
                attrs,
                ..
            } => {
                if let Some(name) = attr(attrs, "name") {
                    let target = attr(attrs, "cell-range-address").unwrap_or_default();
                    let refers_to = convert_reference(&target, &mut found.books);
                    found.names.push(NameInfo {
                        name: name.into_owned(),
                        refers_to,
                        sheet: if table_depth > 0 { current_sheet } else { None },
                        hidden: false,
                        placeholder: false,
                    });
                }
            }
            Event::Start {
                name: "named-expression",
                attrs,
                ..
            } => {
                if let Some(name) = attr(attrs, "name") {
                    let expression = attr(attrs, "expression").unwrap_or_default();
                    let refers_to = formula_to_a1(&expression, &mut found.books);
                    found.names.push(NameInfo {
                        name: name.into_owned(),
                        refers_to: refers_to.trim_start_matches('=').to_string(),
                        sheet: if table_depth > 0 { current_sheet } else { None },
                        hidden: false,
                        placeholder: false,
                    });
                }
            }
            Event::Start {
                name: "null-date",
                attrs,
                ..
            } => found.null_date = attr(attrs, "date-value").map(|d| d.into_owned()),
            Event::Start {
                name: "database-range",
                attrs,
                ..
            } => {
                if let (Some(name), Some(target)) =
                    (attr(attrs, "name"), attr(attrs, "target-range-address"))
                {
                    if !name.starts_with("__Anonymous") {
                        found.ranges.push((name.into_owned(), target.into_owned()));
                    }
                }
            }
            Event::Start {
                name: "table-cell" | "covered-table-cell",
                attrs,
                ..
            } => {
                if let Some(formula) = attr(attrs, "formula") {
                    let _ = formula_to_a1(&formula, &mut found.books);
                }
            }
            _ => {}
        }
    }
    Ok(found)
}

/// Excel's seven error literals, which an OpenDocument string may hold in
/// a formula cell where a host could not say "error".
fn is_error_literal(text: &str) -> bool {
    matches!(
        text,
        "#NULL!" | "#DIV/0!" | "#VALUE!" | "#REF!" | "#NAME?" | "#NUM!" | "#N/A"
    )
}

/// One cell of a row as the file gives it, before its position is known.
struct CellData {
    repeat: u32,
    value: Option<Value>,
    formula: Option<String>,
    matrix: bool,
    /// Its own style, `table:style-name`, which the looks read (KERNEL.57).
    style: Option<String>,
    /// Whether it merges the cells after it or below it.
    spans: bool,
    /// Whether its text holds a run of a style of its own, `text:span`.
    rich: bool,
}

/// The text of a `text:p` just opened: its runs, `text:s` as spaces,
/// `text:tab` and `text:line-break` as what they say, spans and links
/// transparent, anything else inside passed over; `rich` set where a span
/// names a style of its own.
fn paragraph_text(cursor: &mut Cursor, label: &str, rich: &mut bool) -> Result<String, Refusal> {
    let part = CONTENT;
    let mut out = String::new();
    let mut depth = 0u32;
    loop {
        match step(cursor, label, part)? {
            None => return Err(xml_refused(label, part, "a paragraph never closes")),
            Some(Event::Start { name, attrs, empty }) => match name {
                "s" => {
                    let n = attr(attrs, "c")
                        .and_then(|v| v.parse::<usize>().ok())
                        .unwrap_or(1)
                        .min(1 << 16);
                    for _ in 0..n {
                        out.push(' ');
                    }
                    if !empty {
                        depth += 1;
                    }
                }
                "tab" => {
                    out.push('\t');
                    if !empty {
                        depth += 1;
                    }
                }
                "line-break" => {
                    out.push('\n');
                    if !empty {
                        depth += 1;
                    }
                }
                "span" | "a" | "ruby" | "ruby-base" | "meta" | "meta-field" => {
                    if name == "span" && attr(attrs, "style-name").is_some() {
                        *rich = true;
                    }
                    if !empty {
                        depth += 1;
                    }
                }
                _ => {
                    if !empty {
                        skip_element(cursor).map_err(|why| xml_refused(label, part, why))?;
                    }
                }
            },
            Some(Event::End { .. }) => {
                if depth == 0 {
                    return Ok(out);
                }
                depth -= 1;
            }
            Some(Event::Text { text, raw }) => {
                if raw {
                    out.push_str(text);
                } else {
                    out.push_str(&unescape_entities(text));
                }
            }
        }
    }
}

/// The paragraphs of a cell just opened, to its end tag, one line each.
fn cell_text(cursor: &mut Cursor, label: &str, rich: &mut bool) -> Result<String, Refusal> {
    let part = CONTENT;
    let mut paragraphs: Vec<String> = Vec::new();
    loop {
        match step(cursor, label, part)? {
            None => return Err(xml_refused(label, part, "a cell never closes")),
            Some(Event::Start {
                name: "p", empty, ..
            }) => {
                if empty {
                    paragraphs.push(String::new());
                } else {
                    paragraphs.push(paragraph_text(cursor, label, rich)?);
                }
            }
            Some(Event::Start { empty: false, .. }) => {
                skip_element(cursor).map_err(|why| xml_refused(label, part, why))?;
            }
            Some(Event::End { .. }) => return Ok(paragraphs.join("\n")),
            _ => {}
        }
    }
}

/// A cell from its start tag: what it holds, read as the module says.
fn read_cell(
    cursor: &mut Cursor,
    attrs: &str,
    empty: bool,
    label: &str,
) -> Result<CellData, Refusal> {
    let repeat = attr(attrs, "number-columns-repeated")
        .and_then(|v| v.parse::<u32>().ok())
        .unwrap_or(1)
        .max(1);
    let matrix = attr(attrs, "number-matrix-rows-spanned").is_some()
        || attr(attrs, "number-matrix-columns-spanned").is_some();
    let formula = attr(attrs, "formula").map(|f| formula_to_a1(&f, &mut Vec::new()));
    // `office:value-type` beside Calc's `calcext:value-type`: the same
    // local name, two namespaces.
    let mut value_type: Option<String> = None;
    let mut extension_type: Option<String> = None;
    for (name, value) in attributes_full(attrs) {
        let local = name.rsplit(':').next().unwrap_or(name);
        if local == "value-type" {
            if name.starts_with("calcext:") {
                extension_type = Some(unescape_entities(value).into_owned());
            } else {
                value_type = Some(unescape_entities(value).into_owned());
            }
        }
    }
    let mut rich = false;
    let text = if empty {
        String::new()
    } else {
        cell_text(cursor, label, &mut rich)?
    };
    let value = match value_type.as_deref() {
        Some("float") | Some("percentage") | Some("currency") => {
            Some(Value::Number(match attr(attrs, "value") {
                Some(v) => v.trim().to_string(),
                None => text.trim().to_string(),
            }))
        }
        Some("boolean") => attr(attrs, "boolean-value")
            .map(|v| Value::Bool(v.trim().eq_ignore_ascii_case("true")))
            .or_else(|| Some(Value::Bool(text.trim().eq_ignore_ascii_case("true")))),
        Some("date") => attr(attrs, "date-value").map(|v| Value::Date(v.trim().to_string())),
        Some("time") => attr(attrs, "time-value").map(|v| Value::Date(v.trim().to_string())),
        Some("string") => {
            // LibreOffice writes an error's `string-value` empty, its text in the
            // paragraph.
            let error = extension_type.as_deref() == Some("error");
            let text = attr(attrs, "string-value")
                .filter(|v| !(error && v.is_empty()))
                .map(|v| v.into_owned())
                .unwrap_or(text);
            if error || (formula.is_some() && is_error_literal(text.trim())) {
                Some(Value::Error(text.trim().to_string()))
            } else {
                Some(Value::Text(text))
            }
        }
        // Excel's own spelling of an error cell in an .ods (2026-10-05):
        // a value type the format does not define, with the text beside it.
        Some("error") => {
            let text = attr(attrs, "string-value")
                .map(|v| v.into_owned())
                .unwrap_or(text);
            Some(Value::Error(text.trim().to_string()))
        }
        Some("void") => None,
        _ => {
            if formula.is_none() && !text.is_empty() {
                Some(Value::Text(text))
            } else {
                None
            }
        }
    };
    let _ = attr_full(attrs, "office:value-type");
    let spanned = |name: &str| {
        attr(attrs, name)
            .and_then(|v| v.parse::<u32>().ok())
            .unwrap_or(1)
    };
    Ok(CellData {
        repeat,
        value,
        formula,
        matrix,
        style: attr(attrs, "style-name").map(|s| s.into_owned()),
        spans: spanned("number-columns-spanned") > 1 || spanned("number-rows-spanned") > 1,
        rich,
    })
}

/// A row from its start tag to its end tag, cell by cell.
fn read_row(cursor: &mut Cursor, label: &str) -> Result<Vec<CellData>, Refusal> {
    let part = CONTENT;
    let mut cells: Vec<CellData> = Vec::new();
    loop {
        match step(cursor, label, part)? {
            None => return Err(xml_refused(label, part, "a row never closes")),
            Some(Event::Start {
                name: "table-cell" | "covered-table-cell",
                attrs,
                empty,
            }) => cells.push(read_cell(cursor, attrs, empty, label)?),
            Some(Event::Start { empty: false, .. }) => {
                skip_element(cursor).map_err(|why| xml_refused(label, part, why))?;
            }
            Some(Event::End { name: "table-row" }) => return Ok(cells),
            _ => {}
        }
    }
}

impl Source for Document {
    fn label(&self) -> &str {
        &self.label
    }

    fn sheets(&self) -> &[SheetInfo] {
        &self.sheets
    }

    fn names(&self) -> &[NameInfo] {
        &self.names
    }

    fn tables(&self) -> &[TableInfo] {
        &self.tables
    }

    fn external_books(&self) -> &[String] {
        &self.books
    }

    fn summary(&self) -> Summary {
        Summary {
            names: self.names.len(),
            placeholders: 0,
            tables: self.tables.len(),
            external_books: self.books.len(),
            strings: 0,
            string_bytes: 0,
        }
    }

    fn walk_sheet_with(
        &self,
        index: usize,
        sink: &mut dyn Sink,
        book_distinct: &mut HashSet<String>,
    ) -> Result<SheetStats, Refusal> {
        self.walk(index, sink, book_distinct, None)
            .map(|(stats, _)| stats)
    }

    /// One sheet as the record prints it whole, with its looks (KERNEL.57):
    /// its slice read twice, the first pass ([`Document::first_pass`]) for
    /// its own rows of the record, then the walk, each cell's rows in the
    /// record's spelling.
    fn walk_sheet_looks(
        &self,
        index: usize,
        sink: &mut dyn Sink,
        book_distinct: &mut HashSet<String>,
    ) -> Result<(SheetStats, LooksStats), Refusal> {
        let looks = self.looks()?;
        self.walk(index, sink, book_distinct, Some(looks))
    }
}

/// A run of columns and the cell style its cells take where they name
/// none, `table:default-cell-style-name`.
type ColumnDefaults = Vec<(u32, u32, Option<String>)>;

/// The style a column's cells take where they name none.
fn column_default(columns: &ColumnDefaults, col: u32) -> Option<&str> {
    columns
        .iter()
        .find(|(min, max, _)| *min <= col && col <= *max)
        .and_then(|(_, _, s)| s.as_deref())
}

impl Document {
    /// The document's styles as the looks read them, read the first time a
    /// walk asks (KERNEL.57).
    fn looks(&self) -> Result<&OdsLooks, Refusal> {
        self.looks
            .get_or_init(|| {
                odf_looks::read(
                    &self.text,
                    self.styles.as_deref(),
                    self.settings.as_deref(),
                    &self.label,
                )
            })
            .as_ref()
            .map_err(Clone::clone)
    }

    /// A value as the record spells it (KERNEL.57): a date as Excel's serial
    /// in the document's calendar, a time, an ISO duration, as its days; any
    /// other as the relations spell it.
    fn record_value(&self, value: Value) -> Value {
        if let Value::Date(iso) = &value {
            let days = if iso.trim_start_matches('-').starts_with('P') {
                iso_duration_days(iso)
            } else {
                iso_serial(iso, self.date1904)
            };
            if let Some(n) = days.and_then(number_text) {
                return Value::Number(n);
            }
        }
        value
    }

    /// The looks' first pass over a sheet's slice (KERNEL.57): its
    /// gridlines, its columns, their widths, hidden and formats, the rows a
    /// person set a height for or hid, the extent of the cells the walk
    /// prints and the formats they use, and what the record cannot hold
    /// among them, counted: a cell that keeps a format of its own and no
    /// value, a row formatted whole (its default cell style, or, as Excel's
    /// save writes one, its empty cells to the last column in a style of
    /// their own), a merge, a look of no word; and the sheet's conditional formats
    /// as rows (KERNEL.58): LibreOffice's `calcext` elements where the sheet holds
    /// any, else its cell styles' conditions, each style's over the cells naming it.
    fn first_pass(
        &self,
        index: usize,
        looks: &OdsLooks,
    ) -> Result<(SheetScan, ColumnDefaults), Refusal> {
        let part = CONTENT;
        let label = self.label.as_str();
        let info = &self.sheets[index];
        let (start, end) = self.spans[index];
        let mut cursor = Cursor::new(&self.text[start..end]);
        let mut scan = SheetScan::new();
        scan.gridlines = looks.gridlines(&info.name);
        let mut columns: ColumnDefaults = Vec::new();
        let mut col_no: u32 = 0;
        let mut row_no: u32 = 0;
        // The rows of settings as runs, kept to the extent at the end.
        let mut runs: Vec<(u32, u32, RowSettings)> = Vec::new();
        // The styles naming conditions the sheet's cells take, each with its
        // cells, in the order first met; the conditions LibreOffice writes beside
        // its table; and the range of the format holding the condition being read.
        let mut conditioned: Vec<(String, Gathered)> = Vec::new();
        let mut conditioned_at: HashMap<String, usize> = HashMap::new();
        let mut calc_rules: Vec<FileCondition> = Vec::new();
        let mut target: Option<Vec<A1Range>> = None;
        let mut opened = false;
        while let Some(event) = step(&mut cursor, label, part)? {
            match event {
                Event::Start {
                    name: "table",
                    empty,
                    ..
                } if !opened => {
                    if empty {
                        break;
                    }
                    opened = true;
                }
                Event::End { name: "table" } => break,
                Event::Start {
                    name: "table-column",
                    attrs,
                    ..
                } => {
                    let repeat = attr(attrs, "number-columns-repeated")
                        .and_then(|v| v.parse::<u32>().ok())
                        .unwrap_or(1)
                        .max(1);
                    let min = col_no.saturating_add(1);
                    let max = col_no.saturating_add(repeat).min(MAX_COLUMN);
                    col_no = col_no.saturating_add(repeat);
                    let style = attr(attrs, "style-name").map(|s| s.into_owned());
                    let default = attr(attrs, "default-cell-style-name").map(|s| s.into_owned());
                    let width = looks.column_width(style.as_deref());
                    let hidden = attr(attrs, "visibility").as_deref() == Some("collapse")
                        || width == Some(0.0);
                    let format = looks.format(default.as_deref());
                    if min <= max {
                        scan.columns.push(ColumnSpec {
                            min,
                            max,
                            width: if hidden { None } else { width },
                            hidden,
                            style: (format != 0).then_some(format),
                        });
                        columns.push((min, max, default));
                    }
                }
                Event::Start {
                    name: "table-row",
                    attrs,
                    empty,
                } => {
                    let repeat = attr(attrs, "number-rows-repeated")
                        .and_then(|v| v.parse::<u32>().ok())
                        .unwrap_or(1)
                        .max(1);
                    let style = attr(attrs, "style-name").map(|s| s.into_owned());
                    let height = looks.row_height(style.as_deref());
                    let hidden = matches!(
                        attr(attrs, "visibility").as_deref(),
                        Some("collapse") | Some("filter")
                    );
                    if height.is_some() || hidden {
                        runs.push((
                            row_no.saturating_add(1),
                            repeat,
                            RowSettings { height, hidden },
                        ));
                    }
                    let row_default =
                        attr(attrs, "default-cell-style-name").map(|s| s.into_owned());
                    if looks.format(row_default.as_deref()) != 0 {
                        scan.stats.row_formats += u64::from(repeat);
                    }
                    if empty {
                        row_no = row_no.saturating_add(repeat);
                        continue;
                    }
                    let cells = read_row(&mut cursor, label)?;
                    let mut col: u32 = 0;
                    for cell in &cells {
                        let first = col.saturating_add(1);
                        col = col.saturating_add(cell.repeat);
                        if first > MAX_COLUMN {
                            break;
                        }
                        let inherited = row_default
                            .as_deref()
                            .or_else(|| column_default(&columns, first));
                        let named = cell.style.as_deref().or(inherited);
                        // The conditions these cells take, a repeated cell of no
                        // style of its own or its row's by each column's.
                        let last = col.min(MAX_COLUMN);
                        let mut gather = |style: &str, left: u32, right: u32| {
                            if looks.maps_of(style).is_empty() {
                                return;
                            }
                            let at =
                                *conditioned_at.entry(style.to_string()).or_insert_with(|| {
                                    conditioned.push((style.to_string(), Gathered::default()));
                                    conditioned.len() - 1
                                });
                            conditioned[at].1.add(A1Range {
                                top: row_no.saturating_add(1).min(MAX_ROW),
                                left,
                                bottom: row_no.saturating_add(repeat).min(MAX_ROW),
                                right,
                            });
                        };
                        match cell.style.as_deref().or(row_default.as_deref()) {
                            Some(style) => gather(style, first, last),
                            None => {
                                for (min, max, default) in &columns {
                                    if let Some(style) = default {
                                        if *min <= last && first <= *max {
                                            gather(style, first.max(*min), last.min(*max));
                                        }
                                    }
                                }
                            }
                        }
                        if cell.spans {
                            scan.stats.merged += 1;
                        }
                        let format = looks.format(named);
                        if cell.value.is_none() && cell.formula.is_none() {
                            // A cell that keeps a format of its own and no
                            // value; Excel's row of them to the last column, a
                            // row's own format.
                            if cell.style.is_some()
                                && format != 0
                                && format != looks.format(inherited)
                            {
                                if col >= MAX_COLUMN {
                                    scan.stats.row_formats += u64::from(repeat);
                                } else {
                                    scan.stats.formatted_empty +=
                                        u64::from(cell.repeat) * u64::from(repeat);
                                }
                            }
                            continue;
                        }
                        let look = &looks.looks[format as usize];
                        let last_col = col.min(MAX_COLUMN);
                        let last_row = row_no.saturating_add(repeat).min(MAX_ROW);
                        let count = u64::from(last_col - first + 1)
                            * u64::from(last_row.saturating_sub(row_no));
                        scan.hold(row_no.saturating_add(1), first);
                        scan.hold(last_row, last_col);
                        scan.used.insert(format);
                        if format != 0 {
                            scan.stats.styled += count;
                        }
                        if look.unheld || cell.rich {
                            scan.stats.unheld += count;
                        }
                        if look.undrawn.is_some() {
                            scan.stats.undrawn += count;
                        }
                    }
                    row_no = row_no.saturating_add(repeat);
                }
                Event::Start {
                    name: "conditional-format",
                    attrs,
                    ..
                } => target = attr(attrs, "target-range-address").and_then(|a| range_list(&a)),
                Event::Start {
                    name: "condition",
                    attrs,
                    ..
                } => calc_rules.push(FileCondition {
                    ranges: target.clone(),
                    condition: calc_condition(&attr(attrs, "value").unwrap_or_default()),
                    base: attr(attrs, "base-cell-address").map(|b| b.into_owned()),
                    applied: attr(attrs, "apply-style-name").map(|s| s.into_owned()),
                }),
                // A date's rule calls `TODAY`, and the other three compute a look no
                // row says.
                Event::Start {
                    name: "date-is" | "color-scale" | "data-bar" | "icon-set",
                    ..
                } => calc_rules.push(FileCondition {
                    ranges: target.clone(),
                    condition: Condition::NotComputed,
                    base: None,
                    applied: None,
                }),
                Event::Start {
                    name:
                        "table-header-rows"
                        | "table-rows"
                        | "table-row-group"
                        | "table-columns"
                        | "table-header-columns"
                        | "table-column-group"
                        | "conditional-formats",
                    ..
                } => {}
                Event::Start { empty: false, .. } => {
                    skip_element(&mut cursor).map_err(|why| xml_refused(label, part, why))?;
                }
                _ => {}
            }
        }
        let conditions = if calc_rules.is_empty() {
            conditioned
                .iter_mut()
                .for_each(|(_, cells)| cells.close_row());
            conditioned
                .iter()
                .flat_map(|(style, cells)| {
                    looks.maps_of(style).iter().map(|map| FileCondition {
                        ranges: Some(cells.ranges.clone()),
                        condition: style_condition(&map.condition),
                        base: map.base.clone(),
                        applied: Some(map.applied.clone()),
                    })
                })
                .collect()
        } else {
            calc_rules
        };
        for (at, condition) in conditions.into_iter().enumerate() {
            self.take_rule(condition, at as u32 + 1, index, looks, &mut scan);
        }
        if let Some(extent) = scan.extent {
            for (first, count, settings) in runs {
                let last = first.saturating_add(count - 1).min(extent.bottom);
                for row in first.max(extent.top)..=last {
                    scan.rows.insert(row, settings.clone());
                }
            }
        }
        Ok((scan, columns))
    }

    /// A condition read, as a row or as a count: one the language does not compute,
    /// or whose formula a load would refuse, counted and left out; and one the record
    /// cannot say, for want of a range, a test or a style it applies, counted as a look
    /// of no word, a style the document does not declare applying nothing. A cell
    /// takes the style of the first condition true over it and no other, across
    /// formats too: LibreOffice's help and its `ScDocument::GetCondResult`. So each
    /// row stops, what a file's rule is here where a package's may go on.
    fn take_rule(
        &self,
        read: FileCondition,
        priority: u32,
        home: usize,
        looks: &OdsLooks,
        scan: &mut SheetScan,
    ) {
        scan.stats.conditional += 1;
        if matches!(read.condition, Condition::NotComputed) {
            scan.stats.not_computed += 1;
            return;
        }
        let (Some(ranges), Some(applied)) = (read.ranges, read.applied) else {
            scan.stats.unheld += 1;
            return;
        };
        let mut row = Rule {
            priority,
            ranges,
            test: Test::Formula(String::new()),
            format: None,
            stop: true,
        };
        let Some(test) = read.condition.test(row.corner(), read.base.as_deref()) else {
            scan.stats.unheld += 1;
            return;
        };
        row.test = test;
        let Some(look) = looks.rule_look(&applied) else {
            scan.stats.unheld += 1;
            return;
        };
        let gate = self.rule_gate.get_or_init(|| RuleGate::of(self));
        if !gate.takes(&row, home) {
            scan.stats.not_computed += 1;
            return;
        }
        row.format = self.rule_format(look, looks.looks.len(), scan);
        scan.stats.rule_cells += row.cells();
        scan.rules.push(row);
    }

    /// The format a rule's style says, as the record numbers it past the `cells`
    /// formats, each different one once, and kept for the sheet's rows: `None` for one
    /// that says nothing the record has a word for or that a grid cannot hold, counted.
    fn rule_format(&self, look: Look, cells: usize, scan: &mut SheetScan) -> Option<u32> {
        if look.unheld {
            scan.stats.unheld += 1;
        }
        if look.style == Style::partial() {
            return None;
        }
        let mut kept = self.rule_looks.borrow_mut();
        let at = kept.iter().position(|l| *l == look).unwrap_or_else(|| {
            kept.push(look.clone());
            kept.len() - 1
        });
        let Some(index) = u32::try_from(cells + at).ok().filter(|i| *i < FORMATS) else {
            scan.stats.unheld += 1;
            return None;
        };
        scan.rule_formats.insert(index, look);
        Some(index)
    }

    /// One sheet's walk: the relations' rows, or, given the document's
    /// styles (KERNEL.57), the sheet as the record prints it whole, its own
    /// rows of the record before any cell's, a date as its serial, a
    /// formula's saved value its `value` row after the formula, and each
    /// cell's `style` row after its rows and before its `refers` rows.
    fn walk(
        &self,
        index: usize,
        sink: &mut dyn Sink,
        book_distinct: &mut HashSet<String>,
        looks: Option<&OdsLooks>,
    ) -> Result<(SheetStats, LooksStats), Refusal> {
        let part = CONTENT;
        let label = self.label.as_str();
        let info = &self.sheets[index];
        let (start, end) = self.spans[index];
        let slice = &self.text[start..end];
        let mut stats = SheetStats {
            part_bytes: slice.len(),
            ..SheetStats::default()
        };
        // The looks' first pass, and the sheet's own rows of the record
        // before any cell's (KERNEL.57).
        let pass = match looks {
            Some(l) => {
                let (mut scan, columns) = self.first_pass(index, l)?;
                scan.print(&info.name, &l.looks, sink);
                Some((scan, columns))
            }
            None => None,
        };
        let mut sheet_distinct: HashSet<String> = HashSet::new();
        let mut cursor = Cursor::new(slice);
        let mut opened = false;
        let mut row_no: u32 = 0;
        while let Some(event) = step(&mut cursor, label, part)? {
            match event {
                Event::Start {
                    name: "table",
                    empty,
                    ..
                } if !opened => {
                    if empty {
                        break;
                    }
                    opened = true;
                }
                Event::End { name: "table" } => break,
                Event::Start {
                    name: "table-row",
                    attrs,
                    empty,
                } => {
                    let repeat = attr(attrs, "number-rows-repeated")
                        .and_then(|v| v.parse::<u32>().ok())
                        .unwrap_or(1)
                        .max(1);
                    // KERNEL.8, slice e: a row hidden by hand (`collapse`)
                    // or by a filter (`filter`), which `SUBTOTAL`'s 101 to
                    // 111 pass by, told to the sink as the xlsx reader tells
                    // its one flag; a row holding no value or formula is
                    // told nothing, since no aggregate counts it either way.
                    let hidden = matches!(
                        attr(attrs, "visibility").as_deref(),
                        Some("collapse") | Some("filter")
                    );
                    let row_default =
                        attr(attrs, "default-cell-style-name").map(|s| s.into_owned());
                    if empty {
                        row_no = row_no.saturating_add(repeat);
                        continue;
                    }
                    let cells = read_row(&mut cursor, label)?;
                    // Excel's filler below and beside the data is one empty
                    // row repeated a million times holding one empty cell
                    // repeated sixteen thousand: nothing in it is held, so
                    // nothing in it is visited.
                    if cells
                        .iter()
                        .all(|c| c.value.is_none() && c.formula.is_none())
                    {
                        row_no = row_no.saturating_add(repeat);
                        continue;
                    }
                    for r in 0..repeat {
                        let row = row_no.saturating_add(1).saturating_add(r);
                        if row > MAX_ROW {
                            break;
                        }
                        if hidden {
                            sink.hidden_row(&info.name, row);
                        }
                        let mut col: u32 = 0;
                        for cell in &cells {
                            if cell.value.is_none() && cell.formula.is_none() {
                                col = col.saturating_add(cell.repeat);
                                continue;
                            }
                            for c in 0..cell.repeat {
                                let at = col.saturating_add(1).saturating_add(c);
                                if at > MAX_COLUMN {
                                    break;
                                }
                                let addr = cell_ref(row, at);
                                stats.cells += 1;
                                stats.rows = stats.rows.max(row);
                                stats.columns = stats.columns.max(at);
                                let sheet = info.name.as_str();
                                match (looks, &pass) {
                                    (Some(l), Some((_, columns))) => {
                                        let named = cell.style.as_deref().or(row_default
                                            .as_deref()
                                            .or_else(|| column_default(columns, at)));
                                        let format = l.format(named);
                                        let value =
                                            cell.value.clone().map(|v| self.record_value(v));
                                        match &cell.formula {
                                            Some(text) => {
                                                if cell.matrix {
                                                    stats.array_anchors += 1;
                                                }
                                                let saved = value.as_ref().map(datum);
                                                emit_formula_with(
                                                    sink,
                                                    sheet,
                                                    &addr,
                                                    text,
                                                    row,
                                                    at,
                                                    &mut stats,
                                                    &mut sheet_distinct,
                                                    book_distinct,
                                                    &mut |sink| {
                                                        if let Some(saved) = &saved {
                                                            sink.record(&Record::Value {
                                                                sheet,
                                                                addr: &addr,
                                                                value: saved,
                                                            });
                                                        }
                                                        if format != 0 {
                                                            sink.record(&Record::Style {
                                                                sheet,
                                                                addr: &addr,
                                                                index: format,
                                                            });
                                                        }
                                                    },
                                                );
                                                if cell.matrix {
                                                    sink.array(sheet, &addr, ArrayKind::Legacy);
                                                }
                                            }
                                            None => {
                                                if let Some(value) = &value {
                                                    sink.row(&Row::Cell {
                                                        sheet,
                                                        addr: &addr,
                                                        value,
                                                    });
                                                }
                                                if format != 0 {
                                                    sink.record(&Record::Style {
                                                        sheet,
                                                        addr: &addr,
                                                        index: format,
                                                    });
                                                }
                                            }
                                        }
                                    }
                                    _ => {
                                        if let Some(value) = &cell.value {
                                            sink.row(&Row::Cell {
                                                sheet,
                                                addr: &addr,
                                                value,
                                            });
                                        }
                                        if let Some(text) = &cell.formula {
                                            if cell.matrix {
                                                stats.array_anchors += 1;
                                            }
                                            emit_formula(
                                                sink,
                                                sheet,
                                                &addr,
                                                text,
                                                row,
                                                at,
                                                &mut stats,
                                                &mut sheet_distinct,
                                                book_distinct,
                                            );
                                            // A matrix formula's anchor, of
                                            // the kind the format can tell
                                            // (KERNEL.8, slice b).
                                            if cell.matrix {
                                                sink.array(sheet, &addr, ArrayKind::Legacy);
                                            }
                                        }
                                    }
                                }
                            }
                            col = col.saturating_add(cell.repeat);
                        }
                    }
                    row_no = row_no.saturating_add(repeat);
                }
                Event::Start {
                    name: "table-header-rows" | "table-rows" | "table-row-group",
                    ..
                } => {}
                Event::Start { empty: false, .. } => {
                    skip_element(&mut cursor).map_err(|why| xml_refused(label, part, why))?;
                }
                _ => {}
            }
        }
        stats.distinct_r1c1 = sheet_distinct.len() as u64;
        Ok((stats, pass.map(|(scan, _)| scan.stats).unwrap_or_default()))
    }
}

// ---- A conditional format's condition, as a rule ---------------------------

/// The rectangles a style's conditions answer for, gathered from the cells naming
/// it, in the order rows come: a row's runs whole, each then growing the rectangle
/// above it where that is the same run of columns.
#[derive(Debug, Default)]
struct Gathered {
    ranges: Vec<A1Range>,
    /// The last rectangle of each run of columns.
    open: HashMap<(u32, u32), usize>,
    /// The runs of the row being read.
    runs: Vec<A1Range>,
}

impl Gathered {
    /// A cell's rectangle, or a repeated cell's, the next this style's cells hold.
    fn add(&mut self, cells: A1Range) {
        if self.runs.first().is_some_and(|run| run.top != cells.top) {
            self.close_row();
        }
        match self.runs.last_mut() {
            Some(run) if run.right.saturating_add(1) == cells.left => run.right = cells.right,
            _ => self.runs.push(cells),
        }
    }

    /// The row being read, each of its runs growing a rectangle or making one.
    fn close_row(&mut self) {
        for run in std::mem::take(&mut self.runs) {
            match self.open.get(&(run.left, run.right)) {
                Some(&at) if self.ranges[at].bottom.saturating_add(1) == run.top => {
                    self.ranges[at].bottom = run.bottom;
                }
                _ => {
                    self.open.insert((run.left, run.right), self.ranges.len());
                    self.ranges.push(run);
                }
            }
        }
    }
}

/// A condition as a file holds it: the range it answers for, what it asks, the
/// cell its references are read at, and the style it applies.
struct FileCondition {
    ranges: Option<Vec<A1Range>>,
    condition: Condition,
    base: Option<String>,
    applied: Option<String>,
}

/// What a condition asks of a cell.
enum Condition {
    /// A test, its formulas the bar's text, read at the condition's base.
    Test(Test),
    /// A formula around the cell the rule is read at, what a kind of LibreOffice's
    /// own says, Excel's formula for its kin: the text before the cell and after it.
    AtCell(String, String),
    /// A kind the language does not compute.
    NotComputed,
    /// A text this version does not read.
    Unread,
}

impl Condition {
    /// The test a rule's row holds, read at `corner`: a formula read at a base
    /// other than the corner moved from it, as a row's are read at the corner.
    fn test(self, corner: (u32, u32), base: Option<&str>) -> Option<Test> {
        let (row, col) = corner;
        match self {
            Condition::Test(test) => {
                let (from_row, from_col) = base.and_then(base_cell).unwrap_or(corner);
                let (d_row, d_col) = (
                    i64::from(row) - i64::from(from_row),
                    i64::from(col) - i64::from(from_col),
                );
                let moved = |text: &str| refers::shift_a1_references(text, d_row, d_col);
                Some(match test {
                    Test::Formula(text) => Test::Formula(moved(&text)),
                    Test::Cell {
                        operator,
                        first,
                        second,
                    } => Test::Cell {
                        operator,
                        first: moved(&first),
                        second: second.as_deref().map(moved),
                    },
                })
            }
            Condition::AtCell(before, after) => {
                let cell = cell_ref(row, col);
                Some(Test::Formula(format!("{before}{cell}{after}")))
            }
            Condition::NotComputed | Condition::Unread => None,
        }
    }
}

/// A cell an OpenDocument address names, `Sheet.C2`, `$'Q1 Data'.$C$2`: its row and
/// its column.
fn base_cell(address: &str) -> Option<(u32, u32)> {
    let (_, cell) = parse_part(address.trim());
    let range = parse_a1_range(cell)?;
    Some((range.top, range.left))
}

/// A range list an OpenDocument file writes, `S.C2:S.Y2 'Q1 Data'.C4:'Q1 Data'.C4`,
/// as its rectangles; `None` for a list of none or one this version does not read.
fn range_list(address: &str) -> Option<Vec<A1Range>> {
    let mut ranges = Vec::new();
    let mut quoted = false;
    let mut start = 0;
    let mut pieces: Vec<&str> = Vec::new();
    for (i, c) in address.char_indices() {
        match c {
            '\'' => quoted = !quoted,
            ' ' if !quoted => {
                pieces.push(&address[start..i]);
                start = i + 1;
            }
            _ => {}
        }
    }
    pieces.push(&address[start..]);
    for piece in pieces.into_iter().filter(|p| !p.is_empty()) {
        let (first, second) = split_range(piece)?;
        let (_, from) = parse_part(first);
        let to = if second.is_empty() {
            from
        } else {
            parse_part(second).1
        };
        ranges.push(parse_a1_range(&format!("{from}:{to}"))?);
    }
    (!ranges.is_empty()).then_some(ranges)
}

/// An OpenFormula text as a row holds it, without its `=`; `None` for one of nothing.
fn bar_text(text: &str) -> Option<String> {
    let text = formula_to_a1(text.trim(), &mut Vec::new());
    let text = text.strip_prefix('=').unwrap_or(&text).trim();
    (!text.is_empty()).then(|| text.to_string())
}

/// The inside of `<word>(...)`.
fn inside<'t>(text: &'t str, word: &str) -> Option<&'t str> {
    text.strip_prefix(word)?
        .strip_prefix('(')?
        .strip_suffix(')')
}

/// A cell's value against one bound, `<5`, `!=[.A1]`, as a condition.
fn compared(text: &str) -> Condition {
    const SYMBOLS: [(&str, Operator); 6] = [
        ("<=", Operator::LessOrEqual),
        (">=", Operator::GreaterOrEqual),
        ("!=", Operator::NotEqual),
        ("<", Operator::Less),
        (">", Operator::Greater),
        ("=", Operator::Equal),
    ];
    let said = SYMBOLS.iter().find_map(|(symbol, operator)| {
        let first = bar_text(text.strip_prefix(symbol)?)?;
        Some(Test::Cell {
            operator: *operator,
            first,
            second: None,
        })
    });
    said.map_or(Condition::Unread, Condition::Test)
}

/// A cell's value against two bounds, `5,9`, split at the comma outside any text,
/// reference or call, as a condition.
fn between(operator: Operator, text: &str) -> Condition {
    let chars: Vec<char> = text.chars().collect();
    let mut depth = 0u32;
    let mut i = 0;
    while i < chars.len() {
        match chars[i] {
            '"' => {
                i += 1;
                while i < chars.len() && chars[i] != '"' {
                    i += 1;
                }
            }
            '[' | '(' => depth += 1,
            ']' | ')' => depth = depth.saturating_sub(1),
            ',' if depth == 0 => {
                let first: String = chars[..i].iter().collect();
                let second: String = chars[i + 1..].iter().collect();
                return match (bar_text(&first), bar_text(&second)) {
                    (Some(first), Some(second)) => Condition::Test(Test::Cell {
                        operator,
                        first,
                        second: Some(second),
                    }),
                    _ => Condition::Unread,
                };
            }
            _ => {}
        }
        i += 1;
    }
    Condition::Unread
}

/// A formula, as a condition.
fn formula_is(text: &str) -> Condition {
    bar_text(text).map_or(Condition::Unread, |f| Condition::Test(Test::Formula(f)))
}

/// A cell style's condition (`style:map`): Excel's save writes `of:is-true-formula(...)`,
/// `of:cell-content()<5` and `of:cell-content-is-between(5,9)`, and LibreOffice's the same
/// with no prefix.
fn style_condition(text: &str) -> Condition {
    let text = text.trim();
    let text = match text.split_once(':') {
        Some((prefix, rest)) if prefix.chars().all(|c| c.is_ascii_alphanumeric()) => rest,
        _ => text,
    };
    if let Some(inner) = inside(text, "is-true-formula") {
        return formula_is(inner);
    }
    if let Some(inner) = inside(text, "cell-content-is-between") {
        return between(Operator::Between, inner);
    }
    if let Some(inner) = inside(text, "cell-content-is-not-between") {
        return between(Operator::NotBetween, inner);
    }
    match text.strip_prefix("cell-content()") {
        Some(rest) => compared(rest),
        None => Condition::Unread,
    }
}

/// A condition of LibreOffice's (`calcext:value`): a formula, `formula-is(...)`; a
/// cell's value against its bounds; a kind of its own over a cell's text or an error,
/// as Excel's formula for that kind; or a kind the language does not compute.
fn calc_condition(value: &str) -> Condition {
    let value = value.trim();
    if let Some(inner) = inside(value, "formula-is") {
        return formula_is(inner);
    }
    if let Some(inner) = inside(value, "between") {
        return between(Operator::Between, inner);
    }
    if let Some(inner) = inside(value, "not-between") {
        return between(Operator::NotBetween, inner);
    }
    type Around = fn(&str) -> (String, String);
    let over_text: [(&str, Around); 4] = [
        ("contains-text", |t| {
            (format!("NOT(ISERROR(SEARCH({t},"), ")))".to_string())
        }),
        ("not-contains-text", |t| {
            (format!("ISERROR(SEARCH({t},"), "))".to_string())
        }),
        ("begins-with", |t| {
            ("LEFT(".to_string(), format!(",LEN({t}))={t}"))
        }),
        ("ends-with", |t| {
            ("RIGHT(".to_string(), format!(",LEN({t}))={t}"))
        }),
    ];
    for (word, around) in over_text {
        if let Some(inner) = inside(value, word) {
            return bar_text(inner).map_or(Condition::Unread, |text| {
                let (before, after) = around(&text);
                Condition::AtCell(before, after)
            });
        }
    }
    match value {
        "is-error" => return Condition::AtCell("ISERROR(".to_string(), ")".to_string()),
        "is-no-error" => return Condition::AtCell("NOT(ISERROR(".to_string(), "))".to_string()),
        "duplicate"
        | "unique"
        | "above-average"
        | "below-average"
        | "above-equal-average"
        | "below-equal-average" => return Condition::NotComputed,
        _ => {}
    }
    const RANKS: [&str; 4] = [
        "top-elements",
        "bottom-elements",
        "top-percent",
        "bottom-percent",
    ];
    if RANKS.iter().any(|word| inside(value, word).is_some()) {
        return Condition::NotComputed;
    }
    compared(value)
}

// ---- OpenFormula to the formula bar's spelling -----------------------------

/// An OpenFormula text as the formula bar spells it; the files its
/// references reach are added to `books`, each once.
pub fn formula_to_a1(text: &str, books: &mut Vec<String>) -> String {
    // The namespace prefix, `of:=` or another writer's, up to its `:=`.
    let body = match text.find(":=") {
        Some(i)
            if text[..i]
                .chars()
                .all(|c| c.is_ascii_alphanumeric() || c == '.' || c == '_')
                && !text[..i].is_empty() =>
        {
            &text[i + 1..]
        }
        _ => text,
    };
    let chars: Vec<char> = body.chars().collect();
    let mut out = String::with_capacity(body.len() + 1);
    if chars.first() != Some(&'=') {
        out.push('=');
    }
    out.push_str(&convert_body(&chars, books));
    out
}

/// The index of the `)` that closes the `(` at `open`, strings and
/// bracketed references passed over; the end when none closes it.
fn matching_paren(chars: &[char], open: usize) -> usize {
    let mut depth = 0usize;
    let mut i = open;
    while i < chars.len() {
        match chars[i] {
            '"' => {
                i += 1;
                while i < chars.len() {
                    if chars[i] == '"' {
                        if chars.get(i + 1) == Some(&'"') {
                            i += 2;
                            continue;
                        }
                        break;
                    }
                    i += 1;
                }
            }
            '[' => {
                let mut quoted = false;
                i += 1;
                while i < chars.len() {
                    if chars[i] == '\'' {
                        quoted = !quoted;
                    } else if chars[i] == ']' && !quoted {
                        break;
                    }
                    i += 1;
                }
            }
            '(' => depth += 1,
            ')' => {
                depth -= 1;
                if depth == 0 {
                    return i;
                }
            }
            _ => {}
        }
        i += 1;
    }
    chars.len()
}

/// Whether the text at `i` begins with `word`, at a token boundary.
fn ahead_is(chars: &[char], i: usize, word: &str) -> bool {
    let at_boundary = i == 0 || !(chars[i - 1].is_ascii_alphanumeric() || chars[i - 1] == '.');
    at_boundary
        && chars.len() >= i + word.chars().count()
        && chars[i..i + word.chars().count()]
            .iter()
            .copied()
            .eq(word.chars())
}

/// A formula's body, after its `=`, converted; recursive for the argument
/// of `COM.MICROSOFT.SINGLE`, Excel's spelling of the `@` operator.
fn convert_body(chars: &[char], books: &mut Vec<String>) -> String {
    const SINGLE: &str = "COM.MICROSOFT.SINGLE(";
    const MICROSOFT: &str = "COM.MICROSOFT.";
    let mut out = String::with_capacity(chars.len());
    let mut i = 0;
    while i < chars.len() {
        let c = chars[i];
        match c {
            '"' => {
                // A string, with its doubled quotes, copied as it stands.
                out.push('"');
                i += 1;
                while i < chars.len() {
                    out.push(chars[i]);
                    if chars[i] == '"' {
                        if chars.get(i + 1) == Some(&'"') {
                            out.push('"');
                            i += 2;
                            continue;
                        }
                        i += 1;
                        break;
                    }
                    i += 1;
                }
            }
            '[' => {
                let mut j = i + 1;
                let mut quoted = false;
                while j < chars.len() {
                    if chars[j] == '\'' {
                        quoted = !quoted;
                    } else if chars[j] == ']' && !quoted {
                        break;
                    }
                    j += 1;
                }
                let inner: String = chars[i + 1..j.min(chars.len())].iter().collect();
                out.push_str(&convert_reference(&inner, books));
                i = j + 1;
            }
            '#' => {
                // An error literal, its `!` or `?` included.
                let mut j = i + 1;
                while j < chars.len()
                    && (chars[j].is_ascii_alphanumeric() || matches!(chars[j], '/' | '?' | '!'))
                {
                    j += 1;
                }
                out.extend(&chars[i..j]);
                i = j;
            }
            ';' | '~' => {
                out.push(',');
                i += 1;
            }
            '!' => {
                out.push(' ');
                i += 1;
            }
            // Excel writes a named expression as `$$Name`; the bar shows `Name`.
            '$' if chars.get(i + 1) == Some(&'$') => {
                i += 2;
            }
            _ => {
                // Excel writes the `@` operator, implicit intersection, as
                // `COM.MICROSOFT.SINGLE(...)`: the bar shows `@` before the
                // argument. Any other `COM.MICROSOFT.` function keeps its
                // name without the prefix.
                if ahead_is(chars, i, SINGLE) {
                    let open = i + SINGLE.chars().count() - 1;
                    let close = matching_paren(chars, open);
                    out.push('@');
                    out.push_str(&convert_body(
                        &chars[open + 1..close.min(chars.len())],
                        books,
                    ));
                    i = close + 1;
                    continue;
                }
                if ahead_is(chars, i, MICROSOFT) {
                    i += MICROSOFT.chars().count();
                    continue;
                }
                out.push(c);
                i += 1;
            }
        }
    }
    out
}

/// The two halves of a range address at the `:` outside quotes, or the
/// one address and an empty second half.
fn split_range(inner: &str) -> Option<(&str, &str)> {
    let mut quoted = false;
    for (i, c) in inner.char_indices() {
        match c {
            '\'' => quoted = !quoted,
            ':' if !quoted => return Some((&inner[..i], &inner[i + 1..])),
            _ => {}
        }
    }
    Some((inner, ""))
}

/// One address of a reference, `$'Q1 Data'.$B$2` or `.A1` or `.B`: the
/// sheet named, its quotes undone and its `$` dropped (empty when none),
/// and the cell part as written.
fn parse_part(part: &str) -> (String, &str) {
    let part = part.strip_prefix('$').unwrap_or(part);
    if let Some(rest) = part.strip_prefix('\'') {
        // A quoted sheet name to its closing quote, `''` an apostrophe.
        let chars: Vec<(usize, char)> = rest.char_indices().collect();
        let mut name = String::new();
        let mut k = 0;
        while k < chars.len() {
            let (at, c) = chars[k];
            if c == '\'' {
                if chars.get(k + 1).map(|&(_, n)| n) == Some('\'') {
                    name.push('\'');
                    k += 2;
                    continue;
                }
                let after = &rest[at + 1..];
                return (name, after.strip_prefix('.').unwrap_or(after));
            }
            name.push(c);
            k += 1;
        }
        return (name, "");
    }
    match part.find('.') {
        Some(i) => (part[..i].to_string(), &part[i + 1..]),
        None => (String::new(), part),
    }
}

/// The last segment of an IRI, percent-decoded: `Rates.xlsx` from
/// `file:///C:/Models/Rates.xlsx`.
fn file_name_of(iri: &str) -> String {
    let tail = iri.rsplit(['/', '\\']).next().unwrap_or(iri);
    let bytes = tail.as_bytes();
    let mut out: Vec<u8> = Vec::with_capacity(bytes.len());
    let mut i = 0;
    while i < bytes.len() {
        if bytes[i] == b'%' {
            if let Some(v) = tail
                .get(i + 1..i + 3)
                .and_then(|hex| u8::from_str_radix(hex, 16).ok())
            {
                out.push(v);
                i += 3;
                continue;
            }
        }
        out.push(bytes[i]);
        i += 1;
    }
    String::from_utf8(out).unwrap_or_else(|_| tail.to_string())
}

/// A bracketed reference's inside, spelled as the formula bar spells it:
/// `.A1` as `A1`, `Data.A1:.B2` as `Data!A1:B2`, `Model.B1:Scratch.B1` as
/// `Model:Scratch!B1`, `'Rates.xlsx'#$Sheet1.A1` as `[Rates.xlsx]Sheet1!A1`.
fn convert_reference(inner: &str, books: &mut Vec<String>) -> String {
    let mut rest = inner.trim();
    // A source: a quoted IRI and `#`.
    let mut book: Option<String> = None;
    if let Some(after_quote) = rest.strip_prefix('\'') {
        let mut end: Option<usize> = None;
        let chars: Vec<(usize, char)> = after_quote.char_indices().collect();
        let mut k = 0;
        while k < chars.len() {
            if chars[k].1 == '\'' {
                if chars.get(k + 1).map(|&(_, c)| c) == Some('\'') {
                    k += 2;
                    continue;
                }
                end = Some(chars[k].0);
                break;
            }
            k += 1;
        }
        if let Some(e) = end {
            if after_quote[e + 1..].starts_with('#') {
                let iri = after_quote[..e].replace("''", "'");
                let name = file_name_of(&iri);
                if !books.iter().any(|b| fold(b) == fold(&name)) {
                    books.push(name.clone());
                }
                book = Some(name);
                rest = &after_quote[e + 2..];
            }
        }
    }
    let Some((first, second)) = split_range(rest) else {
        return rest.to_string();
    };
    let (sheet1, addr1) = parse_part(first);
    let (sheet2, addr2) = if second.is_empty() {
        (String::new(), "")
    } else {
        parse_part(second)
    };
    let addr = if addr2.is_empty()
        || (addr2 == addr1 && !sheet2.is_empty() && fold(&sheet2) != fold(&sheet1))
    {
        addr1.to_string()
    } else {
        format!("{addr1}:{addr2}")
    };
    let prefix = if sheet1.is_empty() && sheet2.is_empty() {
        String::new()
    } else if sheet2.is_empty() || fold(&sheet2) == fold(&sheet1) {
        let name = if sheet1.is_empty() { &sheet2 } else { &sheet1 };
        match &book {
            Some(b) => quote_book_sheet(b, name),
            None => format!("{}!", refers::quote_sheet(name)),
        }
    } else {
        // A 3D span: both sheets, quoted together when either needs it.
        let span = format!("{sheet1}:{sheet2}");
        let needs =
            refers::quote_sheet(&sheet1) != sheet1 || refers::quote_sheet(&sheet2) != sheet2;
        let spelled = if needs {
            format!("'{}'", span.replace('\'', "''"))
        } else {
            span
        };
        match &book {
            Some(b) => format!("[{b}]{spelled}!"),
            None => format!("{spelled}!"),
        }
    };
    format!("{prefix}{addr}")
}

/// `[Book.xlsx]Sheet!` or, when the sheet needs quoting, `'[Book.xlsx]Sheet 1'!`.
fn quote_book_sheet(book: &str, sheet: &str) -> String {
    if refers::quote_sheet(sheet) == sheet {
        format!("[{book}]{sheet}!")
    } else {
        format!("'[{book}]{}'!", sheet.replace('\'', "''"))
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reflect::{open, reflect_text};

    const OFFICE: &str = "xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\" xmlns:of=\"urn:oasis:names:tc:opendocument:xmlns:of:1.2\" xmlns:calcext=\"urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0\" office:version=\"1.2\"";

    fn ods(content: &str, mimetype: &str) -> Vec<u8> {
        zip::write_stored(&[
            ("mimetype".to_string(), mimetype.as_bytes().to_vec()),
            (CONTENT.to_string(), content.as_bytes().to_vec()),
        ])
        .unwrap()
    }

    #[test]
    fn an_openformula_text_reads_as_the_formula_bar_spells_it() {
        let mut books = Vec::new();
        let cases = [
            ("of:=SUM([.B1:.B2])", "=SUM(B1:B2)"),
            ("of:=[.B1]*2", "=B1*2"),
            ("of:=SUM([Data.B2:Data.B4])", "=SUM(Data!B2:B4)"),
            ("of:=SUM([Data.B2:.B4])", "=SUM(Data!B2:B4)"),
            (
                "of:=['Q1 Data'.B2]+['It''s'.A1]",
                "='Q1 Data'!B2+'It''s'!A1",
            ),
            ("of:=SUM([Model.B1:Scratch.B1])", "=SUM(Model:Scratch!B1)"),
            (
                "of:=SUM(['Q1 Data'.A1:'Q4 Data'.A1])",
                "=SUM('Q1 Data:Q4 Data'!A1)",
            ),
            ("of:=SUM([Jan.A1:Mar.B2])", "=SUM(Jan:Mar!A1:B2)"),
            ("of:=['Rates.xlsx'#$Sheet1.A1]", "=[Rates.xlsx]Sheet1!A1"),
            (
                "of:=['file:///C:/Models/Rates%20Q1.xlsx'#$'Sheet 1'.A1:.B2]",
                "='[Rates Q1.xlsx]Sheet 1'!A1:B2",
            ),
            ("of:=SUM([.B:.B])", "=SUM(B:B)"),
            ("of:=SUM([.1:.3])", "=SUM(1:3)"),
            ("of:=INDIRECT(\"B1\")", "=INDIRECT(\"B1\")"),
            ("of:=OFFSET([.B1];1;0)", "=OFFSET(B1,1,0)"),
            ("of:=COM.MICROSOFT.SEQUENCE(2)", "=SEQUENCE(2)"),
            ("of:=[.A1]&\"a;b~c!\"", "=A1&\"a;b~c!\""),
            ("of:=\"x\"\"y\"", "=\"x\"\"y\""),
            ("of:=[.#REF!]", "=#REF!"),
            ("of:=#REF!+#N/A", "=#REF!+#N/A"),
            ("of:=SUM([.A1:.A3]~[.C1])", "=SUM(A1:A3,C1)"),
            ("of:=SUM([.A1:.B2]![.B1:.C3])", "=SUM(A1:B2 B1:C3)"),
            ("oooc:=[.A1]", "=A1"),
            ("=A1", "=A1"),
            ("A1+1", "=A1+1"),
            ("of:=[$Model.$B$2]", "=Model!$B$2"),
            ("of:=[Model.B1]+[.B2]*Rate", "=Model!B1+B2*Rate"),
            (
                "of:=ORG.OPENOFFICE.STYLE(\"Default\")",
                "=ORG.OPENOFFICE.STYLE(\"Default\")",
            ),
            // Excel's own .ods writing (its save of the fixture, 2026-10-05):
            // a name as `$$Name` inside `COM.MICROSOFT.SINGLE`, the `@` operator.
            ("of:=COM.MICROSOFT.SINGLE($$HiddenName)", "=@HiddenName"),
            ("of:=[.B3]*COM.MICROSOFT.SINGLE($$Rate)", "=B3*@Rate"),
            (
                "of:=COM.MICROSOFT.SINGLE(INDIRECT(\"B1\"))",
                "=@INDIRECT(\"B1\")",
            ),
            (
                "of:=COM.MICROSOFT.SINGLE([.A1:.A3])+COM.MICROSOFT.SINGLE($$Deep)",
                "=@A1:A3+@Deep",
            ),
            ("of:=COM.MICROSOFT.SINGLE($$Rate", "=@Rate"),
        ];
        for (given, want) in cases {
            assert_eq!(formula_to_a1(given, &mut books), want, "{given}");
        }
        assert_eq!(
            books,
            vec!["Rates.xlsx".to_string(), "Rates Q1.xlsx".to_string()]
        );
        // A name's refers-to from its range address.
        assert_eq!(
            convert_reference("$Data.$A$2:.$B$4", &mut books),
            "Data!$A$2:$B$4"
        );
        assert_eq!(convert_reference("$Model.$B$2", &mut books), "Model!$B$2");
        assert_eq!(file_name_of("file:///C:/x/y/Rates.xlsx"), "Rates.xlsx");
        assert_eq!(file_name_of("Rates.xlsx"), "Rates.xlsx");
    }

    #[test]
    fn a_document_reads_as_the_relations() {
        let content = format!(concat!(
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>",
            "<office:document-content {OFFICE}>",
            "<office:automatic-styles>",
            "<style:style style:name=\"ta1\" style:family=\"table\"><style:table-properties table:display=\"true\"/></style:style>",
            "<style:style style:name=\"ta2\" style:family=\"table\"><style:table-properties table:display=\"false\"/></style:style>",
            "<style:style style:name=\"ce1\" style:family=\"table-cell\"/>",
            "</office:automatic-styles>",
            "<office:body><office:spreadsheet>",
            "<table:table table:name=\"Model\" table:style-name=\"ta1\">",
            "<table:table-column table:number-columns-repeated=\"4\"/>",
            "<table:table-row>",
            "<table:table-cell office:value-type=\"string\"><text:p>Rate <text:span text:style-name=\"T1\">applies</text:span></text:p></table:table-cell>",
            "<table:table-cell office:value-type=\"float\" office:value=\"1200\"><text:p>1200</text:p></table:table-cell>",
            "<table:table-cell table:formula=\"of:=[.B1]*2\" office:value-type=\"float\" office:value=\"2400\"><text:p>2400</text:p></table:table-cell>",
            "<table:table-cell table:formula=\"of:=1/0\" office:value-type=\"string\" calcext:value-type=\"error\"><text:p>#DIV/0!</text:p></table:table-cell>",
            "<table:table-cell table:formula=\"of:=NA()\" office:value-type=\"error\" office:string-value=\"#N/A\"><text:p>#N/A</text:p></table:table-cell>",
            "</table:table-row>",
            "<table:table-row table:number-rows-repeated=\"2\">",
            "<table:table-cell office:value-type=\"boolean\" office:boolean-value=\"true\"><text:p>TRUE</text:p></table:table-cell>",
            "<table:table-cell table:style-name=\"ce1\"/>",
            "<table:table-cell office:value-type=\"date\" office:date-value=\"2025-09-30\"><text:p>30/09/25</text:p></table:table-cell>",
            "<table:table-cell office:value-type=\"time\" office:time-value=\"PT12H30M00S\"><text:p>12:30</text:p></table:table-cell>",
            "</table:table-row>",
            "<table:table-row>",
            "<table:table-cell table:formula=\"of:=[.B1:.B2]*2\" table:number-matrix-columns-spanned=\"1\" table:number-matrix-rows-spanned=\"2\" office:value-type=\"float\" office:value=\"2400\"><text:p>2400</text:p></table:table-cell>",
            "<table:table-cell office:value-type=\"string\" table:number-columns-repeated=\"2\"><text:p>a<text:s text:c=\"2\"/>b</text:p><text:p>c</text:p></table:table-cell>",
            "<table:table-cell table:formula=\"of:=['Rates.xlsx'#$Sheet1.A1]\" office:value-type=\"string\"><text:p>#REF!</text:p><office:annotation><text:p>a note</text:p></office:annotation></table:table-cell>",
            "</table:table-row>",
            "<table:named-expressions><table:named-range table:name=\"Local\" table:base-cell-address=\"$Model.$A$1\" table:cell-range-address=\"$Model.$B$1\"/></table:named-expressions>",
            "</table:table>",
            "<table:table table:name=\"Q1 Data\" table:style-name=\"ta2\">",
            "<table:table-row><table:table-cell table:number-columns-repeated=\"2\"/></table:table-row>",
            "<table:table-row><table:table-cell/><table:table-cell office:value-type=\"percentage\" office:value=\"0.2\"><text:p>20%</text:p></table:table-cell></table:table-row>",
            "</table:table>",
            "<table:named-expressions>",
            "<table:named-expression table:name=\"Rate\" table:base-cell-address=\"$Model.$A$1\" table:expression=\"of:=0.2\"/>",
            "<table:named-range table:name=\"Range1\" table:base-cell-address=\"$Model.$A$1\" table:cell-range-address=\"$'Q1 Data'.$A$2:.$B$2\"/>",
            "</table:named-expressions>",
            "<table:database-ranges>",
            "<table:database-range table:name=\"__Anonymous_Sheet_DB__0\" table:target-range-address=\"Model.A1:Model.D1\"/>",
            "<table:database-range table:name=\"Sales\" table:target-range-address=\"'Q1 Data'.A1:'Q1 Data'.B2\"/>",
            "</table:database-ranges>",
            "</office:spreadsheet></office:body></office:document-content>"), OFFICE = OFFICE);
        let bytes = ods(&content, MIMETYPE);
        assert!(is_opendocument(&bytes));
        let got = reflect_text(&bytes, "t.ods").unwrap();
        let want = concat!(
            "(sheet \"Model\" visible)\n",
            "(sheet \"Q1 Data\" hidden)\n",
            "(name \"Model!Local\" \"Model!$B$1\")\n",
            "(name \"Range1\" \"'Q1 Data'!$A$2:$B$2\")\n",
            "(name \"Rate\" \"0.2\")\n",
            "(table \"Sales\" \"Q1 Data\" \"A1:B2\")\n",
            "(cell \"Model\" \"A1\" \"Rate applies\")\n",
            "(cell \"Model\" \"B1\" 1200)\n",
            "(cell \"Model\" \"C1\" 2400)\n",
            "(formula \"Model\" \"C1\" \"=B1*2\")\n",
            "(refers \"Model!C1\" \"Model!B1\")\n",
            "(cell \"Model\" \"D1\" (error \"#DIV/0!\"))\n",
            "(formula \"Model\" \"D1\" \"=1/0\")\n",
            "(cell \"Model\" \"E1\" (error \"#N/A\"))\n",
            "(formula \"Model\" \"E1\" \"=NA()\")\n",
            "(cell \"Model\" \"A2\" true)\n",
            "(cell \"Model\" \"C2\" (date \"2025-09-30\"))\n",
            "(cell \"Model\" \"D2\" (date \"PT12H30M00S\"))\n",
            "(cell \"Model\" \"A3\" true)\n",
            "(cell \"Model\" \"C3\" (date \"2025-09-30\"))\n",
            "(cell \"Model\" \"D3\" (date \"PT12H30M00S\"))\n",
            "(cell \"Model\" \"A4\" 2400)\n",
            "(formula \"Model\" \"A4\" \"=B1:B2*2\")\n",
            "(refers \"Model!A4\" \"Model!B1:B2\")\n",
            // Two paragraphs are two lines; the printer writes the break as
            // the reference's WriteDatum does, as itself.
            "(cell \"Model\" \"B4\" \"a  b\nc\")\n",
            "(cell \"Model\" \"C4\" \"a  b\nc\")\n",
            "(cell \"Model\" \"D4\" (error \"#REF!\"))\n",
            "(formula \"Model\" \"D4\" \"=[Rates.xlsx]Sheet1!A1\")\n",
            "(refers \"Model!D4\" \"[Rates.xlsx]Sheet1!A1\")\n",
            "(cell \"Q1 Data\" \"B2\" 0.2)\n",
        );
        assert_eq!(got, want);
        let doc = Document::open(&bytes, "t.ods").unwrap();
        let summary = doc.summary();
        assert_eq!(
            (summary.names, summary.tables, summary.external_books),
            (3, 1, 1)
        );
        let mut distinct = HashSet::new();
        let stats = doc
            .walk_sheet_with(0, &mut super::super::Discard, &mut distinct)
            .unwrap();
        assert_eq!(
            (
                stats.cells,
                stats.formulas,
                stats.array_anchors,
                stats.rows,
                stats.columns
            ),
            (15, 5, 1, 4, 5)
        );
        assert_eq!((stats.distinct_r1c1, stats.unreadable), (5, 0));
    }

    #[test]
    fn a_matrix_formula_s_kind_goes_beside_its_row() {
        // KERNEL.8, slice b: a matrix formula's anchor told to the sink after
        // its formula row, legacy, the format telling no dynamic-array
        // formula from it; a plain formula told nothing.
        #[derive(Default)]
        struct Kinds(Vec<(String, ArrayKind)>);
        impl Sink for Kinds {
            fn row(&mut self, _: &Row<'_>) {}
            fn array(&mut self, _sheet: &str, addr: &str, kind: ArrayKind) {
                self.0.push((addr.to_string(), kind));
            }
        }
        let content = format!(concat!(
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>",
            "<office:document-content {OFFICE}>",
            "<office:body><office:spreadsheet>",
            "<table:table table:name=\"Model\">",
            "<table:table-row>",
            "<table:table-cell office:value-type=\"float\" office:value=\"1\"><text:p>1</text:p></table:table-cell>",
            "<table:table-cell table:formula=\"of:=[.A1]*2\" office:value-type=\"float\" office:value=\"2\"><text:p>2</text:p></table:table-cell>",
            "<table:table-cell table:formula=\"of:=[.A1:.A2]*2\" table:number-matrix-columns-spanned=\"1\" table:number-matrix-rows-spanned=\"2\" office:value-type=\"float\" office:value=\"2\"><text:p>2</text:p></table:table-cell>",
            "</table:table-row>",
            "</table:table>",
            "</office:spreadsheet></office:body></office:document-content>"), OFFICE = OFFICE);
        let document = Document::open(&ods(&content, MIMETYPE), "t.ods").unwrap();
        let mut kinds = Kinds::default();
        document.walk_sheet(0, &mut kinds).unwrap();
        assert_eq!(kinds.0, vec![("C1".to_string(), ArrayKind::Legacy)]);
    }

    #[test]
    fn a_row_hidden_by_hand_or_by_a_filter_is_told_hidden() {
        // KERNEL.8, slice e: `table:visibility`'s `collapse`, a row hidden by
        // hand, and `filter`, a row a filter hides, each told to the sink as
        // the one flag the xlsx reader tells, at every row a repeat spans; a
        // row shown, or holding no value, told nothing. The door's
        // `SUBTOTAL` 109 passes both by, and its 9, which would count the
        // one and not the other, is labelled over them, as over an xlsx
        // file's.
        #[derive(Default)]
        struct Hidden(Vec<u32>);
        impl Sink for Hidden {
            fn row(&mut self, _: &Row<'_>) {}
            fn hidden_row(&mut self, _sheet: &str, row: u32) {
                self.0.push(row);
            }
        }
        let number = |n: u32| {
            format!("<table:table-cell office:value-type=\"float\" office:value=\"{n}\"><text:p>{n}</text:p></table:table-cell>")
        };
        let content = format!(
            concat!(
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>",
                "<office:document-content {OFFICE}>",
                "<office:body><office:spreadsheet>",
                "<table:table table:name=\"Model\">",
                "<table:table-row>{one}</table:table-row>",
                "<table:table-row table:visibility=\"collapse\">{two}</table:table-row>",
                "<table:table-row table:visibility=\"filter\" table:number-rows-repeated=\"2\">{four}</table:table-row>",
                "<table:table-row table:visibility=\"visible\">{eight}</table:table-row>",
                "<table:table-row table:visibility=\"collapse\"><table:table-cell/></table:table-row>",
                "<table:table-row>",
                "<table:table-cell table:formula=\"of:=SUBTOTAL(109;[.A1:.A6])\" office:value-type=\"float\" office:value=\"9\"><text:p>9</text:p></table:table-cell>",
                "<table:table-cell table:formula=\"of:=SUBTOTAL(9;[.A1:.A6])\" office:value-type=\"float\" office:value=\"11\"><text:p>11</text:p></table:table-cell>",
                "</table:table-row>",
                "</table:table>",
                "</office:spreadsheet></office:body></office:document-content>"
            ),
            OFFICE = OFFICE,
            one = number(1),
            two = number(2),
            four = number(4),
            eight = number(8),
        );
        let bytes = ods(&content, MIMETYPE);
        let mut hidden = Hidden::default();
        Document::open(&bytes, "h.ods")
            .unwrap()
            .walk_sheet(0, &mut hidden)
            .unwrap();
        assert_eq!(hidden.0, vec![2, 3, 4]);
        let library = crate::excel::library();
        let rows = crate::reflect::calc::calc_text(&bytes, "h.ods", &library).unwrap();
        assert_eq!(
            rows,
            concat!(
                "(calc \"Model\" \"A7\" \"=SUBTOTAL(109,A1:A6)\" 9 9 agree)\n",
                "(calc \"Model\" \"B7\" \"=SUBTOTAL(9,A1:A6)\" (not-computed \"a row a filter may hide\") 11 unchecked)\n",
            )
        );
    }

    /// A document in LibreOffice Calc's own dialect, written by hand from its
    /// saves' shapes (KERNEL.57, the third slice): a family's default style
    /// and the named chain `Default`, `Heading`, `Heading 1` in styles.xml,
    /// data styles in content.xml, a column's and a row's default cell style,
    /// a row's height in inches, a date, a time past a day, a merge with its
    /// covered cell formatted, a run of its own style, a formula's saved
    /// value, conditional formats beside the table, and settings.xml's
    /// gridlines a sheet each.
    const CALC_STYLES: &str = concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>",
        "<office:document-styles {OFFICE} xmlns:fo=\"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0\" xmlns:number=\"urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0\" xmlns:svg=\"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0\">",
        "<office:font-face-decls>",
        "<style:font-face style:name=\"Liberation Sans\" svg:font-family=\"&apos;Liberation Sans&apos;\" style:font-family-generic=\"swiss\" style:font-pitch=\"variable\"/>",
        "<style:font-face style:name=\"DejaVu Serif\" svg:font-family=\"&apos;DejaVu Serif&apos;\" style:font-family-generic=\"roman\"/>",
        "</office:font-face-decls>",
        "<office:styles>",
        "<style:default-style style:family=\"table-cell\"><style:paragraph-properties style:tab-stop-distance=\"0.5in\"/>",
        "<style:text-properties style:font-name=\"Liberation Sans\" fo:font-size=\"10pt\" fo:language=\"en\" fo:country=\"US\"/></style:default-style>",
        "<number:number-style style:name=\"N0\"><number:number number:min-integer-digits=\"1\"/></number:number-style>",
        "<style:style style:name=\"Default\" style:family=\"table-cell\"/>",
        "<style:style style:name=\"Heading\" style:family=\"table-cell\" style:parent-style-name=\"Default\"><style:text-properties fo:color=\"#000000\" fo:font-size=\"24pt\" fo:font-style=\"normal\" fo:font-weight=\"bold\"/></style:style>",
        "<style:style style:name=\"Heading_20_1\" style:display-name=\"Heading 1\" style:family=\"table-cell\" style:parent-style-name=\"Heading\"><style:text-properties fo:font-size=\"18pt\"/></style:style>",
        "<style:style style:name=\"Good\" style:family=\"table-cell\" style:parent-style-name=\"Default\"><style:table-cell-properties fo:background-color=\"#ccffcc\"/><style:text-properties fo:color=\"#006600\"/></style:style>",
        "<style:style style:name=\"Note\" style:family=\"table-cell\" style:parent-style-name=\"Default\"><style:text-properties style:font-name=\"DejaVu Serif\"/></style:style>",
        "</office:styles>",
        "<office:automatic-styles><style:page-layout style:name=\"Mpm1\"/></office:automatic-styles>",
        "<office:master-styles><style:master-page style:name=\"Default\" style:page-layout-name=\"Mpm1\"><style:header><text:p><text:sheet-name>???</text:sheet-name></text:p></style:header></style:master-page></office:master-styles>",
        "</office:document-styles>"
    );

    const CALC_CONTENT: &str = concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>",
        "<office:document-content {OFFICE} xmlns:fo=\"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0\" xmlns:number=\"urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0\">",
        "<office:automatic-styles>",
        "<style:style style:name=\"co1\" style:family=\"table-column\"><style:table-column-properties fo:break-before=\"auto\" style:column-width=\"0.889in\"/></style:style>",
        "<style:style style:name=\"co2\" style:family=\"table-column\"><style:table-column-properties fo:break-before=\"auto\" style:column-width=\"1.5in\"/></style:style>",
        "<style:style style:name=\"ro1\" style:family=\"table-row\"><style:table-row-properties style:row-height=\"0.178in\" fo:break-before=\"auto\" style:use-optimal-row-height=\"true\"/></style:style>",
        "<style:style style:name=\"ro2\" style:family=\"table-row\"><style:table-row-properties style:row-height=\"0.4167in\" fo:break-before=\"auto\" style:use-optimal-row-height=\"false\"/></style:style>",
        "<style:style style:name=\"ta1\" style:family=\"table\" style:master-page-name=\"Default\"><style:table-properties table:display=\"true\" style:writing-mode=\"lr-tb\"/></style:style>",
        "<number:number-style style:name=\"N2\"><number:number number:decimal-places=\"2\" number:min-decimal-places=\"2\" number:min-integer-digits=\"1\"/></number:number-style>",
        "<number:date-style style:name=\"N37\" number:automatic-order=\"true\"><number:month number:style=\"long\"/><number:text>/</number:text><number:day number:style=\"long\"/><number:text>/</number:text><number:year/></number:date-style>",
        "<number:time-style style:name=\"N46\" number:truncate-on-overflow=\"false\"><number:hours/><number:text>:</number:text><number:minutes number:style=\"long\"/></number:time-style>",
        "<style:style style:name=\"ce1\" style:family=\"table-cell\" style:parent-style-name=\"Default\" style:data-style-name=\"N2\"/>",
        "<style:style style:name=\"ce2\" style:family=\"table-cell\" style:parent-style-name=\"Heading_20_1\"/>",
        "<style:style style:name=\"ce3\" style:family=\"table-cell\" style:parent-style-name=\"Default\" style:data-style-name=\"N37\"/>",
        "<style:style style:name=\"ce4\" style:family=\"table-cell\" style:parent-style-name=\"Default\" style:data-style-name=\"N46\"/>",
        "<style:style style:name=\"ce5\" style:family=\"table-cell\" style:parent-style-name=\"Default\"><style:table-cell-properties style:text-align-source=\"fix\" style:repeat-content=\"false\" fo:wrap-option=\"wrap\" fo:border=\"0.74pt solid #000000\" style:vertical-align=\"middle\"/><style:paragraph-properties fo:text-align=\"center\" fo:margin-left=\"0in\"/></style:style>",
        "<style:style style:name=\"ce6\" style:family=\"table-cell\" style:parent-style-name=\"Good\"/>",
        "<style:style style:name=\"ce7\" style:family=\"table-cell\" style:parent-style-name=\"Default\"><style:table-cell-properties fo:background-color=\"#ffff00\"/></style:style>",
        "<style:style style:name=\"ce8\" style:family=\"table-cell\" style:parent-style-name=\"Default\"><style:map style:condition=\"cell-content()&gt;5\" style:apply-style-name=\"Good\" style:base-cell-address=\"Sheet1.B6\"/></style:style>",
        "<style:style style:name=\"ce9\" style:family=\"table-cell\" style:parent-style-name=\"Note\"/>",
        "<style:style style:name=\"T1\" style:family=\"text\"><style:text-properties fo:font-weight=\"bold\"/></style:style>",
        "</office:automatic-styles>",
        "<office:body><office:spreadsheet>",
        "<table:table table:name=\"Sheet1\" table:style-name=\"ta1\">",
        "<table:table-column table:style-name=\"co1\" table:default-cell-style-name=\"Default\"/>",
        "<table:table-column table:style-name=\"co2\" table:default-cell-style-name=\"ce1\"/>",
        "<table:table-column table:style-name=\"co1\" table:number-columns-repeated=\"1022\" table:default-cell-style-name=\"Default\"/>",
        "<table:table-row table:style-name=\"ro1\">",
        "<table:table-cell table:style-name=\"ce2\" office:value-type=\"string\" calcext:value-type=\"string\"><text:p>Title</text:p></table:table-cell>",
        "<table:table-cell office:value-type=\"float\" office:value=\"1.5\" calcext:value-type=\"float\"><text:p>1.50</text:p></table:table-cell>",
        "<table:table-cell table:number-columns-repeated=\"1022\"/>",
        "</table:table-row>",
        "<table:table-row table:style-name=\"ro2\">",
        "<table:table-cell table:style-name=\"ce3\" office:value-type=\"date\" office:date-value=\"2025-01-21\" calcext:value-type=\"date\"><text:p>01/21/25</text:p></table:table-cell>",
        "<table:table-cell table:style-name=\"ce4\" office:value-type=\"time\" office:time-value=\"PT36H30M00S\" calcext:value-type=\"time\"><text:p>36:30</text:p></table:table-cell>",
        "<table:table-cell table:number-columns-repeated=\"1022\"/>",
        "</table:table-row>",
        "<table:table-row table:style-name=\"ro1\">",
        "<table:table-cell table:style-name=\"ce5\" table:number-columns-spanned=\"2\" table:number-rows-spanned=\"1\" office:value-type=\"string\" calcext:value-type=\"string\"><text:p>Merged</text:p></table:table-cell>",
        "<table:covered-table-cell table:style-name=\"ce5\"/>",
        "<table:table-cell table:number-columns-repeated=\"1022\"/>",
        "</table:table-row>",
        "<table:table-row table:style-name=\"ro1\" table:default-cell-style-name=\"ce7\">",
        "<table:table-cell office:value-type=\"float\" office:value=\"4\" calcext:value-type=\"float\"><text:p>4</text:p></table:table-cell>",
        "<table:table-cell table:number-columns-repeated=\"1023\"/>",
        "</table:table-row>",
        "<table:table-row table:style-name=\"ro1\">",
        "<table:table-cell table:style-name=\"ce6\" office:value-type=\"string\" calcext:value-type=\"string\"><text:p><text:span text:style-name=\"T1\">Go</text:span>od</text:p></table:table-cell>",
        "<table:table-cell table:style-name=\"ce7\"/>",
        "<table:table-cell table:style-name=\"ce9\" office:value-type=\"string\" calcext:value-type=\"string\"><text:p>serif</text:p></table:table-cell>",
        "<table:table-cell table:number-columns-repeated=\"1021\"/>",
        "</table:table-row>",
        "<table:table-row table:style-name=\"ro1\">",
        "<table:table-cell table:formula=\"of:=[.B1]*2\" office:value-type=\"float\" office:value=\"3\" calcext:value-type=\"float\"><text:p>3</text:p></table:table-cell>",
        "<table:table-cell table:style-name=\"ce8\" office:value-type=\"float\" office:value=\"6\" calcext:value-type=\"float\"><text:p>6</text:p></table:table-cell>",
        "<table:table-cell table:number-columns-repeated=\"1022\"/>",
        "</table:table-row>",
        "<table:table-row table:style-name=\"ro1\" table:number-rows-repeated=\"1048570\"><table:table-cell table:number-columns-repeated=\"1024\"/></table:table-row>",
        "<calcext:conditional-formats>",
        "<calcext:conditional-format calcext:target-range-address=\"Sheet1.B6:Sheet1.B6\"><calcext:condition calcext:apply-style-name=\"Good\" calcext:value=\"&gt;5\" calcext:base-cell-address=\"Sheet1.B6\"/></calcext:conditional-format>",
        "<calcext:conditional-format calcext:target-range-address=\"Sheet1.A1:Sheet1.A6\"><calcext:color-scale><calcext:color-scale-entry calcext:value=\"0\" calcext:type=\"minimum\" calcext:color=\"#ff0000\"/><calcext:color-scale-entry calcext:value=\"0\" calcext:type=\"maximum\" calcext:color=\"#00a933\"/></calcext:color-scale></calcext:conditional-format>",
        "</calcext:conditional-formats>",
        "</table:table>",
        "<table:table table:name=\"Sheet2\" table:style-name=\"ta1\">",
        "<table:table-column table:style-name=\"co1\" table:number-columns-repeated=\"1024\" table:default-cell-style-name=\"Default\"/>",
        "<table:table-row table:style-name=\"ro1\"><table:table-cell office:value-type=\"float\" office:value=\"1\" calcext:value-type=\"float\"><text:p>1</text:p></table:table-cell><table:table-cell table:number-columns-repeated=\"1023\"/></table:table-row>",
        "</table:table>",
        "</office:spreadsheet></office:body></office:document-content>"
    );

    const CALC_SETTINGS: &str = concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>",
        "<office:document-settings xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:config=\"urn:oasis:names:tc:opendocument:xmlns:config:1.0\" office:version=\"1.3\"><office:settings>",
        "<config:config-item-set config:name=\"ooo:view-settings\">",
        "<config:config-item config:name=\"VisibleAreaTop\" config:type=\"int\">0</config:config-item>",
        "<config:config-item-map-indexed config:name=\"Views\"><config:config-item-map-entry>",
        "<config:config-item config:name=\"ViewId\" config:type=\"string\">view1</config:config-item>",
        "<config:config-item-map-named config:name=\"Tables\">",
        "<config:config-item-map-entry config:name=\"Sheet1\"><config:config-item config:name=\"CursorPositionX\" config:type=\"int\">0</config:config-item><config:config-item config:name=\"ShowGrid\" config:type=\"boolean\">true</config:config-item></config:config-item-map-entry>",
        "<config:config-item-map-entry config:name=\"Sheet2\"><config:config-item config:name=\"ShowGrid\" config:type=\"boolean\">false</config:config-item></config:config-item-map-entry>",
        "</config:config-item-map-named>",
        "<config:config-item config:name=\"ActiveTable\" config:type=\"string\">Sheet1</config:config-item>",
        "<config:config-item config:name=\"ShowGrid\" config:type=\"boolean\">true</config:config-item>",
        "</config:config-item-map-entry></config:config-item-map-indexed>",
        "</config:config-item-set>",
        "<config:config-item-set config:name=\"ooo:configuration-settings\"><config:config-item config:name=\"ShowGrid\" config:type=\"boolean\">true</config:config-item></config:config-item-set>",
        "</office:settings></office:document-settings>"
    );

    /// The document in Calc's dialect, its three parts beside its
    /// mimetype, its content's spreadsheet opening with `settings`.
    /// A document of a sheet `S` holding `table`, its automatic styles `automatic` and
    /// its named ones `named`: its `format` rows but format 0, its `conditional` rows,
    /// and its counts.
    fn rules_read(automatic: &str, named: &str, table: &str) -> (Vec<String>, LooksStats) {
        let namespaces = format!("{OFFICE} xmlns:fo=\"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0\" xmlns:number=\"urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0\"");
        let content = format!("<office:document-content {namespaces}><office:automatic-styles>{automatic}</office:automatic-styles><office:body><office:spreadsheet><table:table table:name=\"S\">{table}</table:table><table:table table:name=\"T\"><table:table-row><table:table-cell/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>");
        let styles = format!("<office:document-styles {namespaces}><office:styles><style:style style:name=\"Default\" style:family=\"table-cell\"/>{named}</office:styles></office:document-styles>");
        let bytes = zip::write_stored(&[
            ("mimetype".to_string(), MIMETYPE.as_bytes().to_vec()),
            (CONTENT.to_string(), content.into_bytes()),
            ("styles.xml".to_string(), styles.into_bytes()),
        ])
        .unwrap();
        let (text, stats) = crate::reflect::reflect_looks(&bytes, "rules.ods").unwrap();
        vla_lang::machine::Machine::load(&text, crate::excel::library())
            .unwrap_or_else(|e| panic!("line {}: {}\n{text}", e.line, e.refusal.text));
        let rows = text
            .lines()
            .filter(|l| {
                (l.starts_with("(format ") && !l.starts_with("(format 0 "))
                    || l.starts_with("(conditional ")
            })
            .map(str::to_string)
            .collect();
        (rows, stats)
    }

    const LIT: &str = "<style:style style:name=\"Lit\" style:display-name=\"Lit Up\" style:family=\"table-cell\"><style:table-cell-properties fo:background-color=\"#FFE699\"/></style:style>";

    /// A value in a cell, in a row's cell text.
    const ONE: &str = "<table:table-cell office:value-type=\"float\" office:value=\"1\"><text:p>1</text:p></table:table-cell>";

    /// Each condition LibreOffice's save writes is its row, the first true alone
    /// drawing; a range of several rectangles, a sheet's name quoted, an address's `$`.
    #[test]
    fn libreoffice_s_conditions_read_as_rows() {
        let condition = |range: &str, value: &str, base: &str| {
            format!("<calcext:conditional-format calcext:target-range-address=\"{range}\"><calcext:condition calcext:apply-style-name=\"Lit Up\" calcext:value=\"{value}\" calcext:base-cell-address=\"{base}\"/></calcext:conditional-format>")
        };
        let mut formats = String::new();
        for (range, value, base) in [
            ("S.A1:S.B2", "formula-is([.$A$1]=1)", "S.A1"),
            (
                "S.A1:S.A1 S.C3:S.D4",
                "formula-is(COUNTIF([.A1:.A9];1)&gt;0)",
                "S.A1",
            ),
            ("$S.$A$1:$S.$A$2", "&lt;5", "S.A1"),
            ("S.A1", "&lt;=5", "S.A1"),
            ("S.A1", "=&quot;a,b&quot;", "S.A1"),
            ("S.A1", "!=[.B1]", "S.A1"),
            ("S.A1", "&gt;=5", "S.A1"),
            ("S.A1", "&gt;5", "S.A1"),
            ("S.A1", "between(1,9)", "S.A1"),
            ("S.A1", "not-between([.B1],MAX(1;2))", "S.A1"),
            ("S.C3", "contains-text(&quot;a&quot;)", "S.C3"),
            ("S.C3", "not-contains-text(&quot;a&quot;)", "S.C3"),
            ("S.C3", "begins-with(&quot;a&quot;)", "S.C3"),
            ("S.C3", "ends-with(&quot;a&quot;)", "S.C3"),
            ("S.C3", "is-error", "S.C3"),
            ("S.C3", "is-no-error", "S.C3"),
            ("'S'.C3:'S'.C4", "formula-is(1)", "S.C3"),
            (
                "'My S'.C3:'My S'.C4 'My S'.E5",
                "between(&quot;a,b&quot;,['a,b'.A1])",
                "'My S'.C3",
            ),
            ("S.C3", "not-between(['a,b'.A1],(1))", "S.C3"),
        ] {
            formats.push_str(&condition(range, value, base));
        }
        let table = format!("<table:table-row>{ONE}</table:table-row><calcext:conditional-formats>{formats}</calcext:conditional-formats>");
        let (rows, stats) = rules_read("", LIT, &table);
        let spelled = |priority: u32, range: &str, test: &str| {
            format!("(conditional \"S\" {priority} \"{range}\" {test} 2 stop)")
        };
        let want = [
            "(format 2 (fill \"FFE699\"))".to_string(),
            spelled(1, "A1:B2", "\"=$A$1=1\""),
            spelled(2, "A1 C3:D4", "\"=COUNTIF(A1:A9,1)>0\""),
            spelled(3, "A1:A2", "(cell less \"=5\")"),
            spelled(4, "A1", "(cell less-or-equal \"=5\")"),
            spelled(5, "A1", "(cell equal \"=\\\"a,b\\\"\")"),
            spelled(6, "A1", "(cell not-equal \"=B1\")"),
            spelled(7, "A1", "(cell greater-or-equal \"=5\")"),
            spelled(8, "A1", "(cell greater \"=5\")"),
            spelled(9, "A1", "(cell between \"=1\" \"=9\")"),
            spelled(10, "A1", "(cell not-between \"=B1\" \"=MAX(1,2)\")"),
            spelled(11, "C3", "\"=NOT(ISERROR(SEARCH(\\\"a\\\",C3)))\""),
            spelled(12, "C3", "\"=ISERROR(SEARCH(\\\"a\\\",C3))\""),
            spelled(13, "C3", "\"=LEFT(C3,LEN(\\\"a\\\"))=\\\"a\\\"\""),
            spelled(14, "C3", "\"=RIGHT(C3,LEN(\\\"a\\\"))=\\\"a\\\"\""),
            spelled(15, "C3", "\"=ISERROR(C3)\""),
            spelled(16, "C3", "\"=NOT(ISERROR(C3))\""),
            spelled(17, "C3:C4", "\"=1\""),
            spelled(
                18,
                "C3:C4 E5",
                "(cell between \"=\\\"a,b\\\"\" \"='a,b'!A1\")",
            ),
            spelled(19, "C3", "(cell not-between \"='a,b'!A1\" \"=(1)\")"),
        ];
        assert_eq!(rows, want, "{stats:?}");
        assert_eq!(stats.conditional, 19);
        assert_eq!(stats.not_computed, 0);
        assert_eq!(stats.unheld, 0);
        assert_eq!(stats.rule_cells, 4 + 5 + 2 + 7 + 6 + 2 + 3 + 1);
    }

    /// A formula is read at its base and a row's at its corner, so that one moves from
    /// the first to the second: LibreOffice's base, and a style's that a cell's own
    /// other rule cut away from it. An absolute reference stands.
    #[test]
    fn a_formula_moves_from_its_base_to_the_rule_s_corner() {
        let table = format!(
            "<table:table-row>{ONE}</table:table-row><calcext:conditional-formats><calcext:conditional-format calcext:target-range-address=\"S.D12:S.D13 S.F10:S.F11\"><calcext:condition calcext:apply-style-name=\"Lit\" calcext:value=\"formula-is([.H12]=[.$H$12])\" calcext:base-cell-address=\"S.D12\"/><calcext:condition calcext:apply-style-name=\"Lit\" calcext:value=\"between([.H12],[.H13])\" calcext:base-cell-address=\"$S.$D$12\"/></calcext:conditional-format></calcext:conditional-formats>"
        );
        let (rows, _) = rules_read("", LIT, &table);
        assert_eq!(
            rows[1..],
            [
                "(conditional \"S\" 1 \"D12:D13 F10:F11\" \"=H10=$H$12\" 2 stop)",
                "(conditional \"S\" 2 \"D12:D13 F10:F11\" (cell between \"=H10\" \"=H11\") 2 stop)",
            ]
        );
    }

    /// What the language does not compute is left out and counted, each kind of
    /// LibreOffice's and a condition a load would refuse; what says too little to be a
    /// row, or applies a style the document does not declare, counts as a look of no word.
    #[test]
    fn what_a_condition_cannot_say_is_counted() {
        let condition = |value: &str| {
            format!(
                "<calcext:condition calcext:apply-style-name=\"Lit\" calcext:value=\"{value}\"/>"
            )
        };
        let mut conditions = String::new();
        for value in [
            "duplicate",
            "unique",
            "above-average",
            "below-average",
            "above-equal-average",
            "below-equal-average",
            "top-elements(1)",
            "bottom-elements(1)",
            "top-percent(1)",
            "bottom-percent(1)",
            "formula-is(TODAY()=[.A1])",
            "formula-is([.A1#])",
        ] {
            conditions.push_str(&condition(value));
        }
        let not_computed = 12 + 4;
        for value in [
            "no-such-kind",
            "formula-is()",
            "between(1)",
            "between(,1)",
            "between(1,)",
            "&lt;",
            "contains-text()",
        ] {
            conditions.push_str(&condition(value));
        }
        let table = format!(
            "<table:table-row>{ONE}</table:table-row><calcext:conditional-formats><calcext:conditional-format calcext:target-range-address=\"S.A1\">{conditions}<calcext:condition calcext:apply-style-name=\"Nowhere\" calcext:value=\"formula-is(1)\"/><calcext:condition calcext:value=\"formula-is(1)\"/><calcext:date-is calcext:style=\"Lit\" calcext:date=\"today\"/><calcext:color-scale><calcext:color-scale-entry calcext:value=\"0\" calcext:type=\"minimum\" calcext:color=\"#f8696b\"/></calcext:color-scale><calcext:data-bar calcext:min-length=\"10\"/><calcext:icon-set calcext:icon-set-type=\"3Arrows\"/></calcext:conditional-format><calcext:conditional-format calcext:target-range-address=\"not an address\"><calcext:condition calcext:apply-style-name=\"Lit\" calcext:value=\"formula-is(1)\"/></calcext:conditional-format><calcext:conditional-format calcext:target-range-address=\" \"><calcext:condition calcext:apply-style-name=\"Lit\" calcext:value=\"formula-is(1)\"/></calcext:conditional-format><calcext:conditional-format><calcext:condition calcext:apply-style-name=\"Lit\" calcext:value=\"formula-is(1)\"/></calcext:conditional-format></calcext:conditional-formats>"
        );
        let (rows, stats) = rules_read("", LIT, &table);
        assert_eq!(rows, Vec::<String>::new(), "{stats:?}");
        assert_eq!(stats.conditional, 12 + 7 + 2 + 4 + 3);
        assert_eq!(stats.not_computed, not_computed);
        assert_eq!(stats.unheld, 7 + 2 + 3);
        assert_eq!(stats.rule_cells, 0);
        assert_eq!(
            stats.left_out().as_deref(),
            Some("the looks leave out what the record cannot hold: 12 looks the record has no word for, 16 conditional formats the language does not compute")
        );
    }

    /// Excel's save writes a sheet's conditions on the styles its cells name: each a row
    /// over those cells, in the order first met, each style's in the order written, with
    /// or without a namespace; a row's run whole and the rows' grown into a rectangle, a
    /// repeated row and cell among them, a style a column or a row gives its cells; a
    /// sheet of `calcext` elements reads them alone.
    #[test]
    fn a_style_s_conditions_read_as_rows_over_the_cells_naming_it() {
        let map = |condition: &str, base: &str| {
            format!(
                "<style:map style:condition=\"{condition}\" style:apply-style-name=\"Lit\"{base}/>"
            )
        };
        let cell_style = |name: &str, maps: &str| {
            format!("<style:style style:name=\"{name}\" style:family=\"table-cell\">{maps}</style:style>")
        };
        let automatic = [
            cell_style(
                "two",
                &[
                    map("of:is-true-formula([.$A$1]=1)", ""),
                    map("of:cell-content()&lt;5", ""),
                ]
                .concat(),
            ),
            cell_style(
                "near",
                &map(
                    "is-true-formula([.E2]=1)",
                    " style:base-cell-address=\"S.B2\"",
                ),
            ),
            cell_style("span", &map("of:cell-content-is-between(1,[.$A$1])", "")),
            cell_style("out", &map("of:cell-content-is-not-between(1,2)", "")),
            cell_style("column", &map("cell-content()&gt;=1", "")),
            cell_style("row", &map("ooow:cell-content()!=1", "")),
            cell_style("odd", &map("value()&gt;1", "")),
            cell_style(
                "no-style",
                &map("of:cell-content()=1", "").replace("Lit", "Nowhere"),
            ),
            cell_style("step", &map("of:is-true-formula(2)", "")),
        ]
        .concat();
        let cell = |style: &str| format!("<table:table-cell table:style-name=\"{style}\"/>");
        let table = [
            "<table:table-column table:default-cell-style-name=\"Default\" table:number-columns-repeated=\"7\"/>".to_string(),
            "<table:table-column table:default-cell-style-name=\"column\"/>".to_string(),
            // Two rows of a run of two cells and of one repeated.
            format!("<table:table-row>{ONE}{}{}<table:table-cell table:style-name=\"span\" table:number-columns-repeated=\"2\"/></table:table-row>", cell("two"), cell("two")),
            format!("<table:table-row><table:table-cell/>{}{}<table:table-cell table:number-columns-repeated=\"2\" table:style-name=\"span\"/>{}</table:table-row>", cell("two"), cell("two"), cell("near")),
            // A cut in the run below it, a run apart in its row, and a repeated row.
            format!("<table:table-row table:number-rows-repeated=\"2\">{}<table:table-cell/>{}{}</table:table-row>", cell("two"), cell("two"), cell("out")),
            // A style a column gives the cell, and one the row gives it, nearer.
            "<table:table-row><table:table-cell table:number-columns-repeated=\"8\"/></table:table-row>".to_string(),
            "<table:table-row table:default-cell-style-name=\"row\"><table:table-cell/></table:table-row>".to_string(),
            format!("<table:table-row>{}{}</table:table-row>", cell("odd"), cell("no-style")),
            // A run the column after the last row's, and one a row below a gap.
            format!("<table:table-row>{}</table:table-row>", cell("step")),
            format!("<table:table-row><table:table-cell/>{}</table:table-row>", cell("step")),
            "<table:table-row><table:table-cell/></table:table-row>".to_string(),
            format!("<table:table-row><table:table-cell/>{}</table:table-row>", cell("step")),
        ]
        .concat();
        let (rows, stats) = rules_read(&automatic, LIT, &table);
        assert_eq!(
            rows,
            [
                "(format 11 (fill \"FFE699\"))",
                "(conditional \"S\" 1 \"B1:C2 A3:A4 C3:C4\" \"=$A$1=1\" 11 stop)",
                "(conditional \"S\" 2 \"B1:C2 A3:A4 C3:C4\" (cell less \"=5\") 11 stop)",
                "(conditional \"S\" 3 \"D1:E2\" (cell between \"=1\" \"=$A$1\") 11 stop)",
                "(conditional \"S\" 4 \"F2\" \"=I2=1\" 11 stop)",
                "(conditional \"S\" 5 \"D3:D4\" (cell not-between \"=1\" \"=2\") 11 stop)",
                "(conditional \"S\" 6 \"H5\" (cell greater-or-equal \"=1\") 11 stop)",
                "(conditional \"S\" 7 \"A6\" (cell not-equal \"=1\") 11 stop)",
                "(conditional \"S\" 10 \"A8 B9 B11\" \"=2\" 11 stop)",
            ],
            "{stats:?}"
        );
        assert_eq!(stats.conditional, 10);
        assert_eq!(stats.unheld, 2);
    }

    /// A rule is asked of on the sheet that holds it: a name scoped to another sheet,
    /// which a load would not read from there, keeps no rule out.
    #[test]
    fn a_condition_is_asked_of_on_its_own_sheet() {
        let rule = "<calcext:conditional-formats><calcext:conditional-format calcext:target-range-address=\"X.A1\"><calcext:condition calcext:apply-style-name=\"Lit\" calcext:value=\"formula-is(Odd=1)\"/></calcext:conditional-format></calcext:conditional-formats>";
        let names = "<table:named-expressions><table:named-expression table:name=\"Odd\" table:base-cell-address=\"$S.$A$1\" table:expression=\"of:A1#\"/></table:named-expressions>";
        let namespaces = format!(
            "{OFFICE} xmlns:fo=\"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0\""
        );
        let row = format!("<table:table-row>{ONE}</table:table-row>");
        let content = format!("<office:document-content {namespaces}><office:body><office:spreadsheet><table:table table:name=\"First\">{row}{rule}</table:table><table:table table:name=\"Second\">{row}{rule}{names}</table:table></office:spreadsheet></office:body></office:document-content>");
        let styles = format!("<office:document-styles {namespaces}><office:styles>{LIT}</office:styles></office:document-styles>");
        let bytes = zip::write_stored(&[
            ("mimetype".to_string(), MIMETYPE.as_bytes().to_vec()),
            (CONTENT.to_string(), content.into_bytes()),
            ("styles.xml".to_string(), styles.into_bytes()),
        ])
        .unwrap();
        let (text, stats) = crate::reflect::reflect_looks(&bytes, "two.ods").unwrap();
        assert!(
            text.contains("(conditional \"First\" 1 \"A1\" \"=Odd=1\" "),
            "{text}"
        );
        assert!(!text.contains("(conditional \"Second\""), "{text}");
        assert_eq!(stats.not_computed, 1);
        vla_lang::machine::Machine::load(&text, crate::excel::library())
            .unwrap_or_else(|e| panic!("line {}: {}\n{text}", e.line, e.refusal.text));
    }

    /// A rule's format past what a grid holds is none, and counted.
    #[test]
    fn a_rule_s_format_past_a_grid_s_formats_is_counted() {
        let document = Document::open(&calc_document(""), "calc.ods").unwrap();
        let mut scan = SheetScan::new();
        let look = |fill| {
            let mut style = Style::partial();
            style.say_fill(Some(fill));
            Look {
                style,
                ..Look::default()
            }
        };
        let at = document.rule_format(look(crate::sheet::Rgb(1, 2, 3)), 65_535, &mut scan);
        assert_eq!(at, Some(65_535));
        assert_eq!(scan.stats.unheld, 0);
        let past = document.rule_format(look(crate::sheet::Rgb(4, 5, 6)), 65_535, &mut scan);
        assert_eq!(past, None);
        assert_eq!(scan.stats.unheld, 1);
        assert_eq!(scan.rule_formats.len(), 1);
    }

    /// A sheet of LibreOffice's elements reads them alone, its styles' conditions being
    /// the same rules written again, and each sheet decides for itself.
    #[test]
    fn a_format_of_libreoffice_s_goes_before_a_style_s_condition() {
        let automatic = "<style:style style:name=\"two\" style:family=\"table-cell\"><style:map style:condition=\"is-true-formula(2)\" style:apply-style-name=\"Lit\"/></style:style>";
        let table = "<table:table-row><table:table-cell table:style-name=\"two\"/></table:table-row><calcext:conditional-formats><calcext:conditional-format calcext:target-range-address=\"S.A1\"><calcext:condition calcext:apply-style-name=\"Lit\" calcext:value=\"formula-is(1)\"/></calcext:conditional-format></calcext:conditional-formats>";
        let (rows, stats) = rules_read(automatic, LIT, table);
        assert_eq!(rows[1..], ["(conditional \"S\" 1 \"A1\" \"=1\" 3 stop)"]);
        assert_eq!(stats.conditional, 1);
    }

    /// A rule's format says what its style's chain says less `Default`, each word and a
    /// said default among them, and no more; the same format twice is one; a family and a
    /// vertical alignment are left out and count nothing, and a word of no row counts.
    #[test]
    fn a_rule_s_style_reads_as_a_format_of_groups_alone() {
        let style = |name: &str, attrs: &str, body: &str| {
            format!("<style:style style:name=\"{name}\" style:family=\"table-cell\"{attrs}>{body}</style:style>")
        };
        let cases: Vec<(String, &str, u64)> = vec![
            (style("A", "", "<style:table-cell-properties fo:background-color=\"transparent\"/>"), "(fill none)", 0),
            (style("A", "", "<style:text-properties fo:font-weight=\"normal\" fo:font-style=\"normal\"/>"), "(font not-bold not-italic)", 0),
            (style("A", "", "<style:text-properties fo:font-weight=\"bold\" fo:font-style=\"italic\" fo:font-size=\"20pt\" fo:color=\"#0000FF\"/>"), "(font bold italic 20 \"0000FF\")", 0),
            (style("A", "", "<style:text-properties style:text-underline-style=\"none\" style:text-line-through-style=\"none\"/>"), "(font not-underline not-strike)", 0),
            (style("A", "", "<style:text-properties style:text-underline-style=\"solid\" style:text-underline-type=\"double\" style:text-line-through-style=\"solid\"/>"), "(font double-underline strike)", 0),
            (style("A", "", "<style:paragraph-properties fo:text-align=\"center\"/>"), "(align center)", 0),
            (style("A", "", "<style:table-cell-properties fo:wrap-option=\"wrap\"/>"), "(wrap wrap)", 0),
            (style("A", "", "<style:table-cell-properties fo:wrap-option=\"no-wrap\"/>"), "(wrap nowrap)", 0),
            (style("A", "", "<style:table-cell-properties fo:border-left=\"thin solid #800020\" fo:border-right=\"none\"/>"), "(border (left thin \"800020\") (right none))", 0),
            (style("A", "", "<style:table-cell-properties fo:border=\"none\"/>"), "(border (left none) (right none) (top none) (bottom none))", 0),
            (style("A", " style:data-style-name=\"N2\"", ""), "(number \"0.00\")", 0),
            (style("A", " style:data-style-name=\"N0\"", ""), "(number general)", 0),
            (style("A", " style:data-style-name=\"Nowhere\"", ""), "", 0),
            // The chain: nearest first, `Default` said by no rule.
            (
                format!("{}{}{}", style("A", " style:parent-style-name=\"B\"", "<style:text-properties fo:font-weight=\"bold\"/>"), style("B", " style:parent-style-name=\"Default\"", "<style:text-properties fo:font-weight=\"normal\" fo:font-style=\"italic\"/>"), style("Default", "", "<style:text-properties fo:font-size=\"20pt\"/>")),
                "(font bold italic)",
                0,
            ),
            // What Excel does not draw of a rule, left out and uncounted.
            (style("A", "", "<style:text-properties style:font-name=\"Consolas\"/><style:table-cell-properties style:vertical-align=\"top\"/>"), "", 0),
            // A word of no row beside a word.
            (style("A", "", "<style:text-properties fo:font-weight=\"bold\" style:text-position=\"super 58%\"/>"), "(font bold)", 1),
            (style("A", "", "<style:table-cell-properties fo:background-color=\"red\"/>"), "", 1),
            (style("A", "", "<style:text-properties style:use-window-font-color=\"true\"/>"), "", 1),
            (style("A", "", "<style:text-properties fo:color=\"red\"/>"), "", 1),
            (style("A", "", "<style:text-properties fo:font-size=\"0.01pt\"/>"), "", 1),
            (style("A", " style:data-style-name=\"Odd\"", ""), "", 1),
        ];
        let data = "<number:number-style style:name=\"N2\"><number:number number:decimal-places=\"2\" number:min-decimal-places=\"2\" number:min-integer-digits=\"1\"/></number:number-style><number:number-style style:name=\"N0\"><number:number number:min-integer-digits=\"1\"/></number:number-style><number:number-style style:name=\"Odd\"><style:text-properties fo:color=\"#ff0000\"/><number:number number:decimal-places=\"2\" number:min-decimal-places=\"2\" number:min-integer-digits=\"1\"/></number:number-style>";
        for (styles, words, unheld) in &cases {
            let table = format!("<table:table-row>{ONE}</table:table-row><calcext:conditional-formats><calcext:conditional-format calcext:target-range-address=\"S.A1\"><calcext:condition calcext:apply-style-name=\"A\" calcext:value=\"formula-is(1)\"/></calcext:conditional-format></calcext:conditional-formats>");
            let named = format!("{data}{styles}");
            let (rows, stats) = rules_read("", &named, &table);
            let row = rows.last().expect("its rule");
            if words.is_empty() {
                assert_eq!(rows.len(), 1, "{styles}: {rows:?}");
                assert_eq!(
                    row, "(conditional \"S\" 1 \"A1\" \"=1\" none stop)",
                    "{styles}"
                );
            } else {
                assert_eq!(rows.len(), 2, "{styles}: {rows:?}");
                assert!(
                    rows[0].ends_with(&format!(" {words})")),
                    "{styles}: {}",
                    rows[0]
                );
            }
            assert_eq!(stats.unheld, *unheld, "{styles}");
        }
        // One format of two rules alike.
        let table = format!("<table:table-row>{ONE}</table:table-row><calcext:conditional-formats><calcext:conditional-format calcext:target-range-address=\"S.A1\"><calcext:condition calcext:apply-style-name=\"Lit\" calcext:value=\"formula-is(1)\"/><calcext:condition calcext:apply-style-name=\"Lit Up\" calcext:value=\"formula-is(2)\"/><calcext:condition calcext:apply-style-name=\"Other\" calcext:value=\"formula-is(3)\"/></calcext:conditional-format></calcext:conditional-formats>");
        let other = "<style:style style:name=\"Other\" style:family=\"table-cell\"><style:text-properties fo:font-weight=\"bold\"/></style:style>";
        let (rows, _) = rules_read("", &format!("{LIT}{other}"), &table);
        assert_eq!(
            rows,
            [
                "(format 3 (fill \"FFE699\"))",
                "(format 4 (font bold))",
                "(conditional \"S\" 1 \"A1\" \"=1\" 3 stop)",
                "(conditional \"S\" 2 \"A1\" \"=2\" 3 stop)",
                "(conditional \"S\" 3 \"A1\" \"=3\" 4 stop)",
            ]
        );
    }

    /// LibreOffice writes an error cell's `string-value` empty and its text in the
    /// paragraph, which a load takes where it refuses an error of no name.
    #[test]
    fn libreoffice_s_error_value_reads_from_its_paragraph() {
        let table = "<table:table-row><table:table-cell table:formula=\"of:=#DIV/0!\" office:value-type=\"string\" office:string-value=\"\" calcext:value-type=\"error\"><text:p>#DIV/0!</text:p></table:table-cell><table:table-cell office:value-type=\"string\" office:string-value=\"\"><text:p>shown</text:p></table:table-cell></table:table-row>";
        let namespaces = format!(
            "{OFFICE} xmlns:fo=\"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0\""
        );
        let content = format!("<office:document-content {namespaces}><office:body><office:spreadsheet><table:table table:name=\"S\">{table}</table:table></office:spreadsheet></office:body></office:document-content>");
        let bytes = ods(&content, MIMETYPE);
        let (text, _) = crate::reflect::reflect_looks(&bytes, "error.ods").unwrap();
        assert!(
            text.contains("(value \"S\" \"A1\" (error \"#DIV/0!\"))"),
            "{text}"
        );
        assert!(
            text.contains("(cell \"S\" \"B1\" \"\")"),
            "the text a file says: {text}"
        );
        vla_lang::machine::Machine::load(&text, crate::excel::library())
            .unwrap_or_else(|e| panic!("line {}: {}\n{text}", e.line, e.refusal.text));
    }

    fn calc_document(settings: &str) -> Vec<u8> {
        let part = |text: &str| text.replace("{OFFICE}", OFFICE).into_bytes();
        let content = CALC_CONTENT.replace(
            "<office:spreadsheet>",
            &format!("<office:spreadsheet>{settings}"),
        );
        zip::write_stored(&[
            ("mimetype".to_string(), MIMETYPE.as_bytes().to_vec()),
            (CONTENT.to_string(), part(&content)),
            ("styles.xml".to_string(), part(CALC_STYLES)),
            (
                "settings.xml".to_string(),
                CALC_SETTINGS.as_bytes().to_vec(),
            ),
        ])
        .unwrap()
    }

    #[test]
    fn a_calc_document_s_looks_read_as_the_record_prints_them() {
        let (text, stats) = crate::reflect::reflect_looks(&calc_document(""), "calc.ods").unwrap();
        // Format 0 the `Default` style over the family's default; the
        // others in the order the parts declare them, content's automatic
        // styles first, `ce8` drawing as format 0 (its rule alone) and
        // `ce9` a format of its own for its family, a face's name and its
        // generic; Heading 1's bold from its parent `Heading`, its size
        // its own; column B's format its default cell style, which B1
        // takes and B2's own style passes over; row 4's its default cell
        // style, which A4 takes; a date as its serial, a time past a day as
        // its days, a formula's saved value after it; each sheet's
        // gridlines its own.
        assert_eq!(
            text,
            concat!(
                "(sheet \"Sheet1\" visible)\n",
                "(sheet \"Sheet2\" visible)\n",
                "(gridlines \"Sheet1\" on)\n",
                "(column \"Sheet1\" \"A\" 11.43 shown none)\n",
                "(column \"Sheet1\" \"B\" 19.86 shown 1)\n",
                "(column \"Sheet1\" \"C\" 11.43 shown none)\n",
                "(row \"Sheet1\" 2 30 shown)\n",
                "(format 0 none general nowrap)\n",
                "(format 1 none \"0.00\" nowrap)\n",
                "(format 2 none general nowrap (font bold 18 \"000000\"))\n",
                "(format 3 none \"mm/dd/yy\" nowrap)\n",
                "(format 4 none \"[h]:mm\" nowrap)\n",
                "(format 5 none general wrap (align center middle) (border (left thin \"000000\") (right thin \"000000\") (top thin \"000000\") (bottom thin \"000000\")))\n",
                "(format 6 \"CCFFCC\" general nowrap (font \"006600\"))\n",
                "(format 7 \"FFFF00\" general nowrap)\n",
                "(format 9 none general nowrap (font (family \"DejaVu Serif\" serif)))\n",
                "(format 14 (fill \"CCFFCC\") (font \"006600\"))\n",
                "(conditional \"Sheet1\" 1 \"B6\" (cell greater \"=5\") 14 stop)\n",
                "(cell \"Sheet1\" \"A1\" \"Title\")\n",
                "(style \"Sheet1\" \"A1\" 2)\n",
                "(cell \"Sheet1\" \"B1\" 1.5)\n",
                "(style \"Sheet1\" \"B1\" 1)\n",
                "(cell \"Sheet1\" \"A2\" 45678)\n",
                "(style \"Sheet1\" \"A2\" 3)\n",
                "(cell \"Sheet1\" \"B2\" 1.5208333333333333)\n",
                "(style \"Sheet1\" \"B2\" 4)\n",
                "(cell \"Sheet1\" \"A3\" \"Merged\")\n",
                "(style \"Sheet1\" \"A3\" 5)\n",
                "(cell \"Sheet1\" \"A4\" 4)\n",
                "(style \"Sheet1\" \"A4\" 7)\n",
                "(cell \"Sheet1\" \"A5\" \"Good\")\n",
                "(style \"Sheet1\" \"A5\" 6)\n",
                "(cell \"Sheet1\" \"C5\" \"serif\")\n",
                "(style \"Sheet1\" \"C5\" 9)\n",
                "(formula \"Sheet1\" \"A6\" \"=B1*2\")\n",
                "(value \"Sheet1\" \"A6\" 3)\n",
                "(refers \"Sheet1!A6\" \"Sheet1!B1\")\n",
                "(cell \"Sheet1\" \"B6\" 6)\n",
                "(gridlines \"Sheet2\" off)\n",
                "(column \"Sheet2\" \"A\" 11.43 shown none)\n",
                "(format 0 none general nowrap)\n",
                "(cell \"Sheet2\" \"A1\" 1)\n",
            )
        );
        // The covered cell and B5 keep a format and no value; the run of
        // A5 and C5's family the record cannot hold; LibreOffice's two
        // rules beside the table, its style's map passed over for them, its color
        // scale counted and the other a row.
        assert_eq!(
            (
                stats.styled,
                stats.formatted_empty,
                stats.unheld,
                stats.undrawn,
                stats.merged,
                stats.conditional,
                stats.row_formats
            ),
            (8, 2, 1, 0, 1, 2, 1)
        );
        // A document of 1904's dates counts its serials from its null date.
        let nineteen_o_four = calc_document(
            "<table:calculation-settings><table:null-date table:date-value=\"1904-01-01\"/></table:calculation-settings>",
        );
        let (text, _) = crate::reflect::reflect_looks(&nineteen_o_four, "1904.ods").unwrap();
        assert!(text.contains("(cell \"Sheet1\" \"A2\" 44216)\n"), "{text}");
    }

    /// The looks' first pass and the walk print the same cells: every
    /// OpenDocument file of the tree, and the Calc document, read under the
    /// looks prints the cells and formulas its relations print, in order.
    #[test]
    fn the_first_pass_reads_the_cells_the_walk_prints() {
        #[derive(Default)]
        struct Places(Vec<String>);
        impl Sink for Places {
            fn row(&mut self, row: &Row<'_>) {
                match row {
                    Row::Cell { sheet, addr, .. } | Row::Formula { sheet, addr, .. } => {
                        let place = format!("{sheet}!{addr}");
                        if self.0.last() != Some(&place) {
                            self.0.push(place);
                        }
                    }
                    _ => {}
                }
            }
        }
        let dir = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../scripts/reflect");
        let mut files: Vec<(String, Vec<u8>)> = std::fs::read_dir(&dir)
            .unwrap()
            .flatten()
            .map(|e| e.path())
            .filter(|p| p.extension().and_then(|e| e.to_str()) == Some("ods"))
            .map(|p| (p.display().to_string(), std::fs::read(&p).unwrap()))
            .collect();
        assert!(files.len() >= 3, "{} OpenDocument files", files.len());
        files.push(("calc.ods".to_string(), calc_document("")));
        for (label, bytes) in &files {
            let document = Document::open(bytes, label).unwrap();
            for i in 0..document.sheets().len() {
                let mut related = Places::default();
                document.walk_sheet(i, &mut related).unwrap();
                let mut looked = Places::default();
                document
                    .walk_sheet_looks(i, &mut looked, &mut HashSet::new())
                    .unwrap();
                assert_eq!(related.0, looked.0, "{label}, sheet {i}");
                // The first pass's extent is the printed cells' rectangle.
                let (scan, _) = document.first_pass(i, document.looks().unwrap()).unwrap();
                let places: Vec<(u32, u32)> = related
                    .0
                    .iter()
                    .map(|p| {
                        let a =
                            crate::sheet::parse_a1_range(p.rsplit('!').next().unwrap()).unwrap();
                        (a.top, a.left)
                    })
                    .collect();
                let rectangle = (!places.is_empty()).then(|| crate::sheet::A1Range {
                    top: places.iter().map(|p| p.0).min().unwrap(),
                    left: places.iter().map(|p| p.1).min().unwrap(),
                    bottom: places.iter().map(|p| p.0).max().unwrap(),
                    right: places.iter().map(|p| p.1).max().unwrap(),
                });
                assert_eq!(scan.extent, rectangle, "{label}, sheet {i}");
            }
        }
    }

    #[test]
    fn what_is_not_a_spreadsheet_is_refused_by_name() {
        let text_doc = ods(
            "<office:document-content/>",
            "application/vnd.oasis.opendocument.text",
        );
        assert!(is_opendocument(&text_doc));
        let r = open(&text_doc, "letter.odt").unwrap_err();
        assert_eq!(r.id, "reflect-not-a-workbook");
        assert!(
            r.text.contains("OpenDocument text file, not a spreadsheet"),
            "{}",
            r.text
        );
        let no_content =
            zip::write_stored(&[("mimetype".to_string(), MIMETYPE.as_bytes().to_vec())]).unwrap();
        let r = open(&no_content, "x.ods").unwrap_err();
        assert!(r.text.contains("no content.xml"), "{}", r.text);
        let doctype = ods("<!DOCTYPE x><office:document-content/>", MIMETYPE);
        let r = open(&doctype, "d.ods").unwrap_err();
        assert_eq!(r.id, "reflect-xml-refused");
        // A plain package is not an OpenDocument file.
        let pkg =
            zip::write_stored(&[("xl/workbook.xml".to_string(), b"<workbook/>".to_vec())]).unwrap();
        assert!(!is_opendocument(&pkg));
        assert!(!is_opendocument(b"nope"));
    }
}
