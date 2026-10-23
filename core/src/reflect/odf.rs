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

use std::cell::OnceCell;
use std::collections::{BTreeSet, HashSet};

use super::cursor::{
    attr, attr_full, attributes_full, skip_element, unescape_entities, Cursor, Event,
};
use super::looks::{iso_duration_days, iso_serial, ColumnSpec, LooksStats, SheetScan};
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
use crate::sheet::{cell_ref, number_text, RowSettings, MAX_COLUMN, MAX_ROW};
use vla_lang::rows::{datum, ArrayKind};

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
            let text = attr(attrs, "string-value")
                .map(|v| v.into_owned())
                .unwrap_or(text);
            if extension_type.as_deref() == Some("error")
                || (formula.is_some() && is_error_literal(text.trim()))
            {
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
    /// their own), a merge, a conditional format's rules, a look of no word.
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
        // The styles naming conditions the sheet's cells take, and the
        // conditional formats LibreOffice writes beside its table.
        let mut conditioned: BTreeSet<String> = BTreeSet::new();
        let mut calc_rules: u64 = 0;
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
                        if let Some(n) = named {
                            if looks.rules_of(n) > 0 {
                                conditioned.insert(n.to_string());
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
                    name: "condition" | "date-is" | "color-scale" | "data-bar" | "icon-set",
                    ..
                } => calc_rules += 1,
                Event::Start {
                    name:
                        "table-header-rows"
                        | "table-rows"
                        | "table-row-group"
                        | "table-columns"
                        | "table-header-columns"
                        | "table-column-group"
                        | "conditional-formats"
                        | "conditional-format",
                    ..
                } => {}
                Event::Start { empty: false, .. } => {
                    skip_element(&mut cursor).map_err(|why| xml_refused(label, part, why))?;
                }
                _ => {}
            }
        }
        scan.stats.conditional = if calc_rules > 0 {
            calc_rules
        } else {
            conditioned
                .iter()
                .map(|n| u64::from(looks.rules_of(n)))
                .sum()
        };
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
        // `ce9` a format of its own for its family, which the record has
        // no word for; Heading 1's bold from its parent `Heading`, its size
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
                "(format 9 none general nowrap)\n",
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
        // rules beside the table, its style's map passed over for them.
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
            (8, 2, 2, 0, 1, 2, 1)
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
