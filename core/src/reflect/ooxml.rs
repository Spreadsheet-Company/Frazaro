//! OOXML in (PORT.8, slice 8a): a workbook's package read as relations.
//!
//! New ground, said plainly: the reference opens a workbook in Excel and
//! reads cells through the object model; the package is never seen. What is
//! read here, through `sheet::zip`, `sheet::inflate` and the cursor: the
//! workbook part (its sheet list with each sheet's state, its defined names,
//! its external references), the workbook's relationships (which part each
//! sheet is), the shared-string table (held whole, since any cell may name
//! any string), each sheet's relationships and the Table parts they name,
//! each external link's relationships (the linked file's name), and, one at
//! a time as [`Package::walk_sheet`] is called, the sheet parts: `sheetData`,
//! its rows, each cell's address, type, formula and value. Styles, the
//! theme, the calculation chain, comments, drawings and the properties are
//! never decoded.
//!
//! What a cell gives: a `cell` row when the file holds a value (a constant,
//! or a formula's cached value), a `formula` row when it holds a formula
//! (its text with `=` in front and the file's `_xlfn.` prefixes dropped,
//! which is what `Range.Formula` returns, and a link's `[1]` resolved to
//! the linked file's name, `[Rates.xlsx]`, as the formula bar shows it; a
//! shared formula's children get the first cell's text with its references
//! moved by the cell's distance, as Excel shows them; an array formula is
//! its anchor's text, and the cells it spills into are values alone), then
//! from 8b its `refers` rows, one per distinct thing the formula refers to
//! as `crate::refers` reads it, sorted by the target as printed; nothing
//! for a formatted cell with no value. Excel's own `_xlfn.` placeholder
//! names, which no one defined, print in no `name` row and are counted
//! apart.

use std::borrow::Cow;
use std::collections::HashSet;

use super::cursor::{attr, element_text, skip_element, Cursor, Event};
use super::{print, sheet_prefix, RefersTo, Row, SheetStats, Sink, Value, Visibility};
use crate::intrinsics::fold;
use crate::messages::{raise, Refusal};
use crate::refers;
use crate::sheet::merge::part_text_raw;
use crate::sheet::zip::{self, Entry};
use crate::sheet::{cell_ref, parse_a1_range, xlfn};

/// The most distinct R1C1 formulas a count keeps, per sheet and per
/// workbook; past it the count stands still (AXM.1's own cap).
pub const DISTINCT_CAP: usize = 1 << 20;

const WORKBOOK: &str = "xl/workbook.xml";
const WORKBOOK_RELS: &str = "xl/_rels/workbook.xml.rels";
const SHARED_STRINGS: &str = "xl/sharedStrings.xml";

/// One sheet of the workbook, in tab order.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SheetInfo {
    pub name: String,
    pub visibility: Visibility,
    /// The sheet's part, `xl/worksheets/sheet1.xml`.
    pub part: String,
}

/// One defined name.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct NameInfo {
    pub name: String,
    /// What it refers to, as the file holds it: an A1 formula without `=`.
    pub refers_to: String,
    /// The sheet a sheet-scoped name belongs to, by tab index.
    pub sheet: Option<usize>,
    pub hidden: bool,
    /// Excel's own `_xlfn.` placeholder for a newer function, which no one
    /// defined: counted apart and printed in no row.
    pub placeholder: bool,
}

/// One Table (a ListObject), by the name Excel shows.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct TableInfo {
    pub name: String,
    /// The sheet it sits on, by tab index.
    pub sheet: usize,
    pub range: String,
    pub columns: Vec<String>,
}

/// What `--counts` reports beyond the sheets.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Summary {
    pub names: usize,
    pub placeholders: usize,
    pub tables: usize,
    pub external_books: usize,
    pub strings: usize,
    pub string_bytes: usize,
}

struct Rel {
    id: String,
    kind: String,
    target: String,
}

/// A package opened far enough to walk: its sheet list, names, Tables,
/// external books and shared strings. The sheet parts are read one at a
/// time by [`Package::walk_sheet`] and dropped.
pub struct Package<'a> {
    label: String,
    bytes: &'a [u8],
    entries: Vec<Entry>,
    sheets: Vec<SheetInfo>,
    names: Vec<NameInfo>,
    tables: Vec<TableInfo>,
    external_books: Vec<String>,
    strings: Vec<String>,
    string_bytes: usize,
}

impl std::fmt::Debug for Package<'_> {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("Package")
            .field("label", &self.label)
            .field("bytes", &self.bytes.len())
            .field("sheets", &self.sheets)
            .field("names", &self.names)
            .field("tables", &self.tables)
            .field("external_books", &self.external_books)
            .field("strings", &self.strings.len())
            .finish()
    }
}

fn not_a_workbook(label: &str, why: &str) -> Refusal {
    raise("reflect-not-a-workbook", &[("path", label), ("why", why)])
}

fn xml_refused(label: &str, part: &str, why: &str) -> Refusal {
    raise(
        "reflect-xml-refused",
        &[("path", label), ("part", part), ("why", why)],
    )
}

fn unsupported(label: &str, part: &str, why: &str) -> Refusal {
    raise(
        "reflect-unsupported",
        &[("path", label), ("part", part), ("why", why)],
    )
}

/// The next event of a part, the cursor's refusal given the part's name.
fn step<'t>(
    cursor: &mut Cursor<'t>,
    label: &str,
    part: &str,
) -> Result<Option<Event<'t>>, Refusal> {
    cursor
        .next_event()
        .map_err(|why| xml_refused(label, part, why))
}

/// The text of the element just opened, under the part's refusal.
fn inner_text(
    cursor: &mut Cursor,
    skip: &[&str],
    excel_escapes: bool,
    label: &str,
    part: &str,
) -> Result<String, Refusal> {
    element_text(cursor, skip, excel_escapes).map_err(|why| xml_refused(label, part, why))
}

/// A relationship's target resolved against the folder of the part that
/// names it: `worksheets/sheet1.xml` from `xl/` is `xl/worksheets/sheet1.xml`,
/// `../tables/table1.xml` from `xl/worksheets/` is `xl/tables/table1.xml`,
/// and a target that begins with `/` is from the package root.
fn resolve(base_dir: &str, target: &str) -> String {
    let joined = match target.strip_prefix('/') {
        Some(absolute) => absolute.to_string(),
        None => format!("{base_dir}{target}"),
    };
    let mut parts: Vec<&str> = Vec::new();
    for segment in joined.split('/') {
        match segment {
            "" | "." => {}
            ".." => {
                parts.pop();
            }
            s => parts.push(s),
        }
    }
    parts.join("/")
}

/// The folder of a part with its slash: `xl/worksheets/` for
/// `xl/worksheets/sheet1.xml`.
fn dir_of(part: &str) -> &str {
    match part.rfind('/') {
        Some(i) => &part[..=i],
        None => "",
    }
}

/// The relationships part beside a part: `xl/_rels/workbook.xml.rels`.
fn rels_of(part: &str) -> String {
    let dir = dir_of(part);
    format!("{dir}_rels/{}.rels", &part[dir.len()..])
}

/// The last segment of a path or URL, percent-decoded: `Rates Q1.xlsx` for
/// `file:///C:/Models/Rates%20Q1.xlsx`.
fn file_name_of(target: &str) -> String {
    let last = target.rsplit(['/', '\\']).next().unwrap_or(target);
    let bytes = last.as_bytes();
    let mut out: Vec<u8> = Vec::with_capacity(bytes.len());
    let mut i = 0;
    while i < bytes.len() {
        if bytes[i] == b'%' && i + 2 < bytes.len() {
            if let Ok(v) = u8::from_str_radix(&last[i + 1..i + 3], 16) {
                out.push(v);
                i += 3;
                continue;
            }
        }
        out.push(bytes[i]);
        i += 1;
    }
    String::from_utf8(out).unwrap_or_else(|_| last.to_string())
}

fn read_rels(text: &str, label: &str, part: &str) -> Result<Vec<Rel>, Refusal> {
    let mut cursor = Cursor::new(text);
    let mut out = Vec::new();
    while let Some(event) = step(&mut cursor, label, part)? {
        if let Event::Start {
            name: "Relationship",
            attrs,
            ..
        } = event
        {
            out.push(Rel {
                id: attr(attrs, "Id").unwrap_or_default().into_owned(),
                kind: attr(attrs, "Type").unwrap_or_default().into_owned(),
                target: attr(attrs, "Target").unwrap_or_default().into_owned(),
            });
        }
    }
    Ok(out)
}

fn read_strings(text: &str, label: &str, part: &str) -> Result<Vec<String>, Refusal> {
    let mut cursor = Cursor::new(text);
    let mut out = Vec::new();
    while let Some(event) = step(&mut cursor, label, part)? {
        if let Event::Start {
            name: "si", empty, ..
        } = event
        {
            if empty {
                out.push(String::new());
            } else {
                out.push(inner_text(&mut cursor, &["rPh"], true, label, part)?);
            }
        }
    }
    Ok(out)
}

fn read_table(
    text: &str,
    label: &str,
    part: &str,
    sheet: usize,
) -> Result<Option<TableInfo>, Refusal> {
    let mut cursor = Cursor::new(text);
    let mut info: Option<TableInfo> = None;
    while let Some(event) = step(&mut cursor, label, part)? {
        match event {
            Event::Start {
                name: "table",
                attrs,
                ..
            } => {
                let name = attr(attrs, "displayName")
                    .or_else(|| attr(attrs, "name"))
                    .unwrap_or_default()
                    .into_owned();
                info = Some(TableInfo {
                    name,
                    sheet,
                    range: attr(attrs, "ref").unwrap_or_default().into_owned(),
                    columns: Vec::new(),
                });
            }
            Event::Start {
                name: "tableColumn",
                attrs,
                ..
            } => {
                if let Some(t) = info.as_mut() {
                    t.columns
                        .push(attr(attrs, "name").unwrap_or_default().into_owned());
                }
            }
            _ => {}
        }
    }
    Ok(info)
}

impl<'a> Package<'a> {
    /// The package read far enough to walk; `label` is what a refusal calls
    /// the file.
    pub fn open(bytes: &'a [u8], label: &str) -> Result<Package<'a>, Refusal> {
        let entries = zip::entries(bytes).ok_or_else(|| {
            not_a_workbook(
                label,
                zip::why_not_an_archive(bytes)
                    .unwrap_or("it is not a zip archive this reader knows"),
            )
        })?;
        let has = |name: &str| entries.iter().any(|e| e.name == name);
        if !has(WORKBOOK) {
            if has("content.xml") && has("mimetype") {
                return Err(not_a_workbook(
                    label,
                    "it is an OpenDocument spreadsheet, which this version does not read yet (PORT.8, slice 8e)",
                ));
            }
            return Err(not_a_workbook(
                label,
                "it has no xl/workbook.xml part, so it is not a spreadsheet package",
            ));
        }
        let text = |name: &str| -> Result<Option<String>, Refusal> {
            part_text_raw(bytes, &entries, name).map_err(|why| not_a_workbook(label, &why))
        };
        let workbook = text(WORKBOOK)?.unwrap_or_default();
        let rels_text = text(WORKBOOK_RELS)?
            .ok_or_else(|| not_a_workbook(label, "it has no xl/_rels/workbook.xml.rels part"))?;
        let rels = read_rels(&rels_text, label, WORKBOOK_RELS)?;
        let target_of = |rid: &str| {
            rels.iter()
                .find(|r| r.id == rid)
                .map(|r| resolve("xl/", &r.target))
        };

        // The workbook part: the sheets, the names, the external references.
        let mut sheets: Vec<SheetInfo> = Vec::new();
        let mut names: Vec<NameInfo> = Vec::new();
        let mut external_rids: Vec<String> = Vec::new();
        let mut cursor = Cursor::new(&workbook);
        while let Some(event) = step(&mut cursor, label, WORKBOOK)? {
            match event {
                Event::Start {
                    name: "sheet",
                    attrs,
                    ..
                } => {
                    let sheet_name = attr(attrs, "name").unwrap_or_default().into_owned();
                    let visibility = match attr(attrs, "state").as_deref() {
                        Some("hidden") => Visibility::Hidden,
                        Some("veryHidden") => Visibility::VeryHidden,
                        _ => Visibility::Visible,
                    };
                    let rid = attr(attrs, "id").unwrap_or_default();
                    let part = target_of(&rid).ok_or_else(|| {
                        unsupported(
                            label,
                            WORKBOOK,
                            &format!(
                                "sheet {sheet_name} names a relationship the workbook does not have"
                            ),
                        )
                    })?;
                    sheets.push(SheetInfo {
                        name: sheet_name,
                        visibility,
                        part,
                    });
                }
                Event::Start {
                    name: "definedName",
                    attrs,
                    empty,
                } => {
                    let name = attr(attrs, "name").unwrap_or_default().into_owned();
                    let sheet = attr(attrs, "localSheetId").and_then(|v| v.parse::<usize>().ok());
                    let hidden =
                        matches!(attr(attrs, "hidden").as_deref(), Some("1") | Some("true"));
                    let refers_to = if empty {
                        String::new()
                    } else {
                        inner_text(&mut cursor, &[], false, label, WORKBOOK)?
                    };
                    let placeholder = fold(&name).starts_with("_xlfn.");
                    names.push(NameInfo {
                        name,
                        refers_to,
                        sheet,
                        hidden,
                        placeholder,
                    });
                }
                Event::Start {
                    name: "externalReference",
                    attrs,
                    ..
                } => {
                    if let Some(rid) = attr(attrs, "id") {
                        external_rids.push(rid.into_owned());
                    }
                }
                _ => {}
            }
        }
        if sheets.is_empty() {
            return Err(unsupported(label, WORKBOOK, "it lists no sheet"));
        }

        // The shared strings, held whole.
        let strings_part = rels
            .iter()
            .find(|r| r.kind.ends_with("/sharedStrings"))
            .map(|r| resolve("xl/", &r.target))
            .unwrap_or_else(|| SHARED_STRINGS.to_string());
        let (strings, string_bytes) = match text(&strings_part)? {
            None => (Vec::new(), 0),
            Some(t) => (read_strings(&t, label, &strings_part)?, t.len()),
        };

        // The Tables, through each sheet's relationships.
        let mut tables: Vec<TableInfo> = Vec::new();
        for (i, sheet) in sheets.iter().enumerate() {
            let rels_part = rels_of(&sheet.part);
            let Some(rels_text) = text(&rels_part)? else {
                continue;
            };
            for rel in read_rels(&rels_text, label, &rels_part)? {
                if !rel.kind.ends_with("/table") {
                    continue;
                }
                let table_part = resolve(dir_of(&sheet.part), &rel.target);
                let Some(table_text) = text(&table_part)? else {
                    continue;
                };
                if let Some(table) = read_table(&table_text, label, &table_part, i)? {
                    tables.push(table);
                }
            }
        }

        // The external books, each by the name of its file.
        let mut external_books: Vec<String> = Vec::new();
        for rid in external_rids {
            let Some(link_part) = target_of(&rid) else {
                continue;
            };
            let link_rels = rels_of(&link_part);
            let name = match text(&link_rels)? {
                Some(rels_text) => read_rels(&rels_text, label, &link_rels)?
                    .into_iter()
                    .find(|r| r.kind.ends_with("/externalLinkPath"))
                    .map(|r| file_name_of(&r.target))
                    .unwrap_or_default(),
                None => String::new(),
            };
            external_books.push(name);
        }

        Ok(Package {
            label: label.to_string(),
            bytes,
            entries,
            sheets,
            names,
            tables,
            external_books,
            strings,
            string_bytes,
        })
    }

    /// The sheets in tab order.
    pub fn sheets(&self) -> &[SheetInfo] {
        &self.sheets
    }

    /// Every defined name, placeholders included, in the file's order.
    pub fn names(&self) -> &[NameInfo] {
        &self.names
    }

    /// Every Table, in the order the sheets' relationships name them.
    pub fn tables(&self) -> &[TableInfo] {
        &self.tables
    }

    /// The linked workbooks' file names, in the order `[1]`, `[2]`, ... of a
    /// formula's text; an empty name for a link of another kind.
    pub fn external_books(&self) -> &[String] {
        &self.external_books
    }

    /// The counts `--counts` prints for the workbook.
    pub fn summary(&self) -> Summary {
        let placeholders = self.names.iter().filter(|n| n.placeholder).count();
        Summary {
            names: self.names.len() - placeholders,
            placeholders,
            tables: self.tables.len(),
            external_books: self.external_books.len(),
            strings: self.strings.len(),
            string_bytes: self.string_bytes,
        }
    }

    /// A name as a row prints it: a sheet-scoped name behind its sheet.
    /// A `name` row's first field: the name, a sheet-scoped one behind its
    /// sheet as a reference quotes it (`Model!Local`). The key `diff` matches
    /// names by.
    pub(crate) fn printed_name(&self, name: &NameInfo) -> String {
        match name.sheet.and_then(|i| self.sheets.get(i)) {
            Some(sheet) => format!("{}!{}", sheet_prefix(&sheet.name), name.name),
            None => name.name.clone(),
        }
    }

    /// The rows that come before any cell, in the fixed order: every
    /// `sheet` row in tab order, every `name` row sorted by its name without
    /// case, every `table` row by its sheet's tab order then its name.
    pub fn header_rows(&self, sink: &mut dyn Sink) {
        for sheet in &self.sheets {
            sink.row(&Row::Sheet {
                name: &sheet.name,
                visibility: sheet.visibility,
            });
        }
        let mut names: Vec<(String, &str)> = self
            .names
            .iter()
            .filter(|n| !n.placeholder)
            .map(|n| (self.printed_name(n), n.refers_to.as_str()))
            .collect();
        names.sort_by(|a, b| fold(&a.0).cmp(&fold(&b.0)).then_with(|| a.0.cmp(&b.0)));
        for (name, refers_to) in &names {
            sink.row(&Row::Name { name, refers_to });
        }
        let mut tables: Vec<&TableInfo> = self.tables.iter().collect();
        tables.sort_by(|a, b| {
            a.sheet
                .cmp(&b.sheet)
                .then_with(|| fold(&a.name).cmp(&fold(&b.name)))
        });
        for table in tables {
            sink.row(&Row::Table {
                name: &table.name,
                sheet: &self.sheets[table.sheet].name,
                range: &table.range,
            });
        }
    }

    /// One sheet's cells in document order, to the sink: a `cell` row for a
    /// value the file holds, a `formula` row for a formula, then its
    /// `refers` rows. The sheet's part is read here and dropped on return;
    /// what the walk counted comes back.
    pub fn walk_sheet(&self, index: usize, sink: &mut dyn Sink) -> Result<SheetStats, Refusal> {
        self.walk_sheet_with(index, sink, &mut HashSet::new())
    }

    /// [`Package::walk_sheet`], with the workbook-wide set of distinct R1C1
    /// formulas carried across sheets for `--counts`.
    pub fn walk_sheet_with(
        &self,
        index: usize,
        sink: &mut dyn Sink,
        book_distinct: &mut HashSet<String>,
    ) -> Result<SheetStats, Refusal> {
        let info = &self.sheets[index];
        let mut sheet_distinct: HashSet<String> = HashSet::new();
        let part = info.part.as_str();
        let label = self.label.as_str();
        let text = part_text_raw(self.bytes, &self.entries, part)
            .map_err(|why| not_a_workbook(label, &why))?
            .ok_or_else(|| {
                unsupported(label, part, "the sheet's part is missing from the package")
            })?;
        let mut stats = SheetStats {
            part_bytes: text.len(),
            ..SheetStats::default()
        };
        let mut cursor = Cursor::new(&text);
        // Shared formulas by index: the first cell's text and position.
        let mut shared: Vec<Option<(String, u32, u32)>> = Vec::new();
        let mut in_data = false;
        let mut row_no: u32 = 0;
        let mut col_no: u32 = 0;
        while let Some(event) = step(&mut cursor, label, part)? {
            match event {
                Event::Start {
                    name: "sheetData",
                    empty,
                    ..
                } => {
                    if empty {
                        break;
                    }
                    in_data = true;
                }
                Event::End { name: "sheetData" } => break,
                Event::Start {
                    name: "row", attrs, ..
                } if in_data => {
                    row_no = attr(attrs, "r")
                        .and_then(|v| v.parse::<u32>().ok())
                        .unwrap_or(row_no + 1);
                    col_no = 0;
                }
                Event::Start {
                    name: "c",
                    attrs,
                    empty,
                } if in_data => {
                    let (row, col) = match attr(attrs, "r").and_then(|r| parse_a1_range(&r)) {
                        Some(a) => (a.top, a.left),
                        None => (row_no, col_no + 1),
                    };
                    row_no = row;
                    col_no = col;
                    let kind: Option<Cow<str>> = attr(attrs, "t");
                    let mut formula_part: Option<(Option<String>, Option<String>, String)> = None;
                    let mut value_text: Option<String> = None;
                    let mut inline: Option<String> = None;
                    if !empty {
                        loop {
                            match step(&mut cursor, label, part)? {
                                None => {
                                    return Err(xml_refused(label, part, "a cell never closes"))
                                }
                                Some(Event::Start {
                                    name: "f",
                                    attrs: f_attrs,
                                    empty: f_empty,
                                }) => {
                                    let t = attr(f_attrs, "t").map(Cow::into_owned);
                                    let si = attr(f_attrs, "si").map(Cow::into_owned);
                                    let f_text = if f_empty {
                                        String::new()
                                    } else {
                                        inner_text(&mut cursor, &[], false, label, part)?
                                    };
                                    formula_part = Some((t, si, f_text));
                                }
                                Some(Event::Start {
                                    name: "v",
                                    empty: v_empty,
                                    ..
                                }) => {
                                    value_text = Some(if v_empty {
                                        String::new()
                                    } else {
                                        inner_text(&mut cursor, &[], false, label, part)?
                                    });
                                }
                                Some(Event::Start {
                                    name: "is",
                                    empty: is_empty,
                                    ..
                                }) => {
                                    inline = Some(if is_empty {
                                        String::new()
                                    } else {
                                        inner_text(&mut cursor, &["rPh"], true, label, part)?
                                    });
                                }
                                Some(Event::Start { empty: false, .. }) => {
                                    skip_element(&mut cursor)
                                        .map_err(|why| xml_refused(label, part, why))?;
                                }
                                Some(Event::End { name: "c" }) => break,
                                _ => {}
                            }
                        }
                    }
                    let addr = cell_ref(row, col);
                    let formula = match formula_part {
                        None => None,
                        Some((t, si, f_text)) => {
                            match t.as_deref() {
                                Some("shared") => {
                                    let i: usize = si
                                    .as_deref()
                                    .and_then(|s| s.trim().parse().ok())
                                    .ok_or_else(|| {
                                        unsupported(
                                            label,
                                            part,
                                            &format!("cell {addr} holds a shared formula with no index"),
                                        )
                                    })?;
                                    if f_text.is_empty() {
                                        let (master, m_row, m_col) =
                                        shared.get(i).and_then(|m| m.as_ref()).ok_or_else(|| {
                                            unsupported(
                                                label,
                                                part,
                                                &format!("cell {addr} continues a shared formula whose first cell has not been read"),
                                            )
                                        })?;
                                        Some(refers::shift_a1_references(
                                            master,
                                            i64::from(row) - i64::from(*m_row),
                                            i64::from(col) - i64::from(*m_col),
                                        ))
                                    } else {
                                        if shared.len() <= i {
                                            shared.resize(i + 1, None);
                                        }
                                        shared[i] = Some((f_text.clone(), row, col));
                                        Some(f_text)
                                    }
                                }
                                Some("array") => {
                                    if f_text.is_empty() {
                                        None
                                    } else {
                                        stats.array_anchors += 1;
                                        Some(f_text)
                                    }
                                }
                                Some("dataTable") => {
                                    return Err(unsupported(
                                        label,
                                        part,
                                        &format!("cell {addr} holds a data-table formula"),
                                    ))
                                }
                                _ => {
                                    if f_text.is_empty() {
                                        None
                                    } else {
                                        Some(f_text)
                                    }
                                }
                            }
                        }
                    };
                    let value_text = value_text.filter(|s| !s.is_empty());
                    let value = match kind.as_deref() {
                        Some("s") => match value_text {
                            Some(s) => {
                                let i: usize = s.trim().parse().map_err(|_| {
                                    unsupported(
                                        label,
                                        part,
                                        &format!("cell {addr} names a shared string by {s}, which is not an index"),
                                    )
                                })?;
                                let text = self.strings.get(i).cloned().ok_or_else(|| {
                                    unsupported(
                                        label,
                                        part,
                                        &format!("cell {addr} names shared string {i}, past the end of the table"),
                                    )
                                })?;
                                Some(Value::Text(text))
                            }
                            None => None,
                        },
                        Some("str") => value_text.map(Value::Text),
                        Some("b") => value_text.map(|s| {
                            let s = s.trim();
                            Value::Bool(s == "1" || s.eq_ignore_ascii_case("true"))
                        }),
                        Some("e") => value_text.map(|s| Value::Error(s.trim().to_string())),
                        Some("d") => value_text.map(|s| Value::Date(s.trim().to_string())),
                        Some("inlineStr") => inline.map(Value::Text),
                        _ => value_text.map(|s| Value::Number(s.trim().to_string())),
                    };
                    if value.is_none() && formula.is_none() {
                        continue;
                    }
                    stats.cells += 1;
                    stats.rows = stats.rows.max(row);
                    stats.columns = stats.columns.max(col);
                    if let Some(value) = &value {
                        sink.row(&Row::Cell {
                            sheet: &info.name,
                            addr: &addr,
                            value,
                        });
                    }
                    if let Some(f_text) = formula {
                        stats.formulas += 1;
                        let text = format!("={}", xlfn::strip_future_prefixes(&f_text));
                        let text = refers::resolve_books(&text, &self.external_books);
                        sink.row(&Row::Formula {
                            sheet: &info.name,
                            addr: &addr,
                            text: &text,
                        });
                        let found = refers::scan(&text);
                        stats.unreadable += found
                            .iter()
                            .filter(|r| r.kind == refers::Kind::Unreadable)
                            .count() as u64;
                        let mut targets: Vec<RefersTo> =
                            found.iter().map(|r| RefersTo::of(r, &info.name)).collect();
                        targets.sort_by_key(print::target);
                        targets.dedup();
                        for to in &targets {
                            sink.row(&Row::Refers {
                                sheet: &info.name,
                                addr: &addr,
                                to,
                            });
                        }
                        let rendered = refers::r1c1(&text, row, col);
                        if sheet_distinct.len() < DISTINCT_CAP {
                            sheet_distinct.insert(rendered.clone());
                        }
                        if book_distinct.len() < DISTINCT_CAP {
                            book_distinct.insert(rendered);
                        }
                    }
                }
                _ => {}
            }
        }
        stats.distinct_r1c1 = sheet_distinct.len() as u64;
        Ok(stats)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reflect::reflect_text;

    const NS: &str = "xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"";
    const REL_SHEET: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
    const REL_SST: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
    const REL_TABLE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";

    /// A package of one sheet, with a string table when given.
    fn one_sheet(sheet_data: &str, sst: Option<&str>) -> Vec<u8> {
        let mut parts: Vec<(String, Vec<u8>)> = vec![
            ("xl/workbook.xml".to_string(), format!("<workbook {NS}><sheets><sheet name=\"Model\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>").into_bytes()),
            ("xl/_rels/workbook.xml.rels".to_string(), format!("<Relationships><Relationship Id=\"rId1\" Type=\"{REL_SHEET}\" Target=\"worksheets/sheet1.xml\"/><Relationship Id=\"rId2\" Type=\"{REL_SST}\" Target=\"sharedStrings.xml\"/></Relationships>").into_bytes()),
            ("xl/worksheets/sheet1.xml".to_string(), format!("<worksheet {NS}><dimension ref=\"A1\"/><sheetData>{sheet_data}</sheetData><pageMargins left=\"0.7\"/></worksheet>").into_bytes()),
        ];
        if let Some(s) = sst {
            parts.push((
                "xl/sharedStrings.xml".to_string(),
                format!("<sst {NS}>{s}</sst>").into_bytes(),
            ));
        }
        zip::write_stored(&parts).unwrap()
    }

    #[test]
    fn each_kind_of_cell_reads() {
        let sst = "<si><t>Revenue</t></si><si><r><t>Rate </t></r><r><rPr><b/></rPr><t>applies</t></r><rPh sb=\"0\" eb=\"1\"><t>x</t></rPh></si>";
        let data = concat!(
            "<row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c><c r=\"B1\"><v>1200</v></c><c r=\"C1\" s=\"3\"/><c r=\"D1\" t=\"inlineStr\"><is><t>in &amp; line</t></is></c></row>",
            "<row r=\"2\"><c r=\"A2\" t=\"s\"><v>1</v></c><c r=\"B2\" t=\"b\"><v>1</v></c><c r=\"C2\" t=\"e\"><f>1/0</f><v>#DIV/0!</v></c><c r=\"D2\" t=\"d\"><v>2026-10-03</v></c></row>",
            "<row r=\"3\"><c r=\"A3\" t=\"str\"><f>A1&amp;\"!\"</f><v>Revenue!</v></c><c r=\"B3\"><f>B1*2</f><v>2400</v></c><c r=\"C3\"><f t=\"shared\" ref=\"C3:C4\" si=\"0\">B3+1</f><v>2401</v></c></row>",
            "<row r=\"4\"><c r=\"C4\"><f t=\"shared\" si=\"0\"/><v>1</v></c><c r=\"D4\" cm=\"1\"><f t=\"array\" ref=\"D4:D5\">_xlfn.SEQUENCE(2)</f><v>1</v></c></row>",
            "<row r=\"5\"><c r=\"D5\"><v>2</v></c><c><v>7</v></c><c t=\"s\"><v>0</v></c></row>",
            "<row><c r=\"A6\"><f></f><v>0.30000000000000004</v></c><c r=\"B6\"><v></v></c></row>",
        );
        let bytes = one_sheet(data, Some(sst));
        let text = reflect_text(&bytes, "t.xlsx").unwrap();
        assert_eq!(
            text,
            concat!(
                "(sheet \"Model\" visible)\n",
                "(cell \"Model\" \"A1\" \"Revenue\")\n",
                "(cell \"Model\" \"B1\" 1200)\n",
                "(cell \"Model\" \"D1\" \"in & line\")\n",
                "(cell \"Model\" \"A2\" \"Rate applies\")\n",
                "(cell \"Model\" \"B2\" true)\n",
                "(cell \"Model\" \"C2\" (error \"#DIV/0!\"))\n",
                "(formula \"Model\" \"C2\" \"=1/0\")\n",
                "(cell \"Model\" \"D2\" (date \"2026-10-03\"))\n",
                "(cell \"Model\" \"A3\" \"Revenue!\")\n",
                "(formula \"Model\" \"A3\" \"=A1&\\\"!\\\"\")\n",
                "(refers \"Model!A3\" \"Model!A1\")\n",
                "(cell \"Model\" \"B3\" 2400)\n",
                "(formula \"Model\" \"B3\" \"=B1*2\")\n",
                "(refers \"Model!B3\" \"Model!B1\")\n",
                "(cell \"Model\" \"C3\" 2401)\n",
                "(formula \"Model\" \"C3\" \"=B3+1\")\n",
                "(refers \"Model!C3\" \"Model!B3\")\n",
                "(cell \"Model\" \"C4\" 1)\n",
                "(formula \"Model\" \"C4\" \"=B4+1\")\n",
                "(refers \"Model!C4\" \"Model!B4\")\n",
                "(cell \"Model\" \"D4\" 1)\n",
                "(formula \"Model\" \"D4\" \"=SEQUENCE(2)\")\n",
                "(cell \"Model\" \"D5\" 2)\n",
                "(cell \"Model\" \"E5\" 7)\n",
                "(cell \"Model\" \"F5\" \"Revenue\")\n",
                "(cell \"Model\" \"A6\" 0.30000000000000004)\n",
            )
        );
        let package = Package::open(&bytes, "t.xlsx").unwrap();
        let stats = package.walk_sheet(0, &mut super::super::Discard).unwrap();
        assert_eq!(stats.cells, 16);
        assert_eq!(stats.formulas, 6);
        assert_eq!(stats.array_anchors, 1);
        // =1/0, =RC[-2]&"!", =RC[-1]*2, =RC[-1]+1 (C3 and C4 as one), =SEQUENCE(2).
        assert_eq!((stats.distinct_r1c1, stats.unreadable), (5, 0));
        assert_eq!((stats.rows, stats.columns), (6, 6));
        assert!(stats.part_bytes > 500);
        let summary = package.summary();
        assert_eq!((summary.strings, summary.names, summary.tables), (2, 0, 0));
    }

    #[test]
    fn the_header_rows_come_in_the_fixed_order() {
        let parts: Vec<(String, Vec<u8>)> = vec![
            ("xl/workbook.xml".to_string(), format!("<workbook {NS}><sheets><sheet name=\"Model\" sheetId=\"1\" r:id=\"rId1\"/><sheet name=\"Q1 Data\" sheetId=\"3\" state=\"hidden\" r:id=\"rId2\"/><sheet name=\"Secret\" sheetId=\"4\" state=\"veryHidden\" r:id=\"rId3\"/></sheets><externalReferences><externalReference r:id=\"rId9\"/></externalReferences><definedNames><definedName name=\"rate\">0.2</definedName><definedName name=\"_xlfn.SINGLE\" hidden=\"1\">#NAME?</definedName><definedName name=\"Local\" localSheetId=\"1\">'Q1 Data'!$B$2</definedName><definedName name=\"Broken\" hidden=\"1\">Model!#REF!</definedName><definedName name=\"Alpha\">Model!$A$1</definedName></definedNames></workbook>").into_bytes()),
            ("xl/_rels/workbook.xml.rels".to_string(), format!("<Relationships><Relationship Id=\"rId1\" Type=\"{REL_SHEET}\" Target=\"worksheets/sheet1.xml\"/><Relationship Id=\"rId2\" Type=\"{REL_SHEET}\" Target=\"/xl/worksheets/sheet2.xml\"/><Relationship Id=\"rId3\" Type=\"{REL_SHEET}\" Target=\"worksheets/sheet3.xml\"/><Relationship Id=\"rId9\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink\" Target=\"externalLinks/externalLink1.xml\"/></Relationships>").into_bytes()),
            ("xl/worksheets/sheet1.xml".to_string(), format!("<worksheet {NS}><sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData><tableParts count=\"2\"><tablePart r:id=\"rId1\"/><tablePart r:id=\"rId2\"/></tableParts></worksheet>").into_bytes()),
            ("xl/worksheets/_rels/sheet1.xml.rels".to_string(), format!("<Relationships><Relationship Id=\"rId1\" Type=\"{REL_TABLE}\" Target=\"../tables/table1.xml\"/><Relationship Id=\"rId2\" Type=\"{REL_TABLE}\" Target=\"../tables/table2.xml\"/></Relationships>").into_bytes()),
            ("xl/tables/table1.xml".to_string(), format!("<table {NS} id=\"1\" name=\"Table1\" displayName=\"Sales\" ref=\"A1:B4\"><tableColumns count=\"2\"><tableColumn id=\"1\" name=\"Item\"/><tableColumn id=\"2\" name=\"Amount\"/></tableColumns></table>").into_bytes()),
            ("xl/tables/table2.xml".to_string(), format!("<table {NS} id=\"2\" name=\"Costs\" displayName=\"Costs\" ref=\"D1:D9\"><tableColumns count=\"1\"><tableColumn id=\"1\" name=\"Cost\"/></tableColumns></table>").into_bytes()),
            ("xl/worksheets/sheet2.xml".to_string(), format!("<worksheet {NS}><sheetData/></worksheet>").into_bytes()),
            ("xl/worksheets/sheet3.xml".to_string(), format!("<worksheet {NS}><sheetData><row r=\"2\"><c r=\"B2\"><v>10</v></c></row></sheetData></worksheet>").into_bytes()),
            ("xl/externalLinks/externalLink1.xml".to_string(), format!("<externalLink {NS}><externalBook r:id=\"rId1\"><sheetNames><sheetName val=\"Sheet1\"/></sheetNames></externalBook></externalLink>").into_bytes()),
            ("xl/externalLinks/_rels/externalLink1.xml.rels".to_string(), "<Relationships><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath\" Target=\"file:///C:/Models/Rates%20Q1.xlsx\" TargetMode=\"External\"/></Relationships>".to_string().into_bytes()),
        ];
        let bytes = zip::write_stored(&parts).unwrap();
        let text = reflect_text(&bytes, "t.xlsx").unwrap();
        assert_eq!(
            text,
            concat!(
                "(sheet \"Model\" visible)\n",
                "(sheet \"Q1 Data\" hidden)\n",
                "(sheet \"Secret\" very-hidden)\n",
                "(name \"'Q1 Data'!Local\" \"'Q1 Data'!$B$2\")\n",
                "(name \"Alpha\" \"Model!$A$1\")\n",
                "(name \"Broken\" \"Model!#REF!\")\n",
                "(name \"rate\" \"0.2\")\n",
                "(table \"Costs\" \"Model\" \"D1:D9\")\n",
                "(table \"Sales\" \"Model\" \"A1:B4\")\n",
                "(cell \"Model\" \"A1\" 1)\n",
                "(cell \"Secret\" \"B2\" 10)\n",
            )
        );
        let package = Package::open(&bytes, "t.xlsx").unwrap();
        assert_eq!(package.external_books(), ["Rates Q1.xlsx"]);
        assert_eq!(package.tables()[0].columns, ["Item", "Amount"]);
        let summary = package.summary();
        assert_eq!(
            (
                summary.names,
                summary.placeholders,
                summary.tables,
                summary.external_books
            ),
            (4, 1, 2, 1)
        );
        assert!(package
            .names()
            .iter()
            .any(|n| n.name == "Broken" && n.hidden));
    }

    #[test]
    fn what_is_refused_is_named_through_the_catalogue() {
        let r = Package::open(b"nope", "x.xlsx").unwrap_err();
        assert_eq!(r.id, "reflect-not-a-workbook");
        assert!(r.text.starts_with("x.xlsx is not a workbook frazaro reflect can read: it has no end-of-central-directory record"), "{}", r.text);
        let ole = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, 0, 0, 0, 0];
        assert!(Package::open(&ole, "old.xls")
            .unwrap_err()
            .text
            .contains("older .xls"));
        let no_wb = zip::write_stored(&[("a.txt".to_string(), b"hi".to_vec())]).unwrap();
        assert!(Package::open(&no_wb, "x.zip")
            .unwrap_err()
            .text
            .contains("no xl/workbook.xml"));
        let ods = zip::write_stored(&[
            (
                "mimetype".to_string(),
                b"application/vnd.oasis.opendocument.spreadsheet".to_vec(),
            ),
            (
                "content.xml".to_string(),
                b"<office:document-content/>".to_vec(),
            ),
        ])
        .unwrap();
        assert!(Package::open(&ods, "x.ods")
            .unwrap_err()
            .text
            .contains("OpenDocument spreadsheet"));
        // A declaration in a part is refused before anything else is read.
        let doctype = one_sheet(
            "<!DOCTYPE x><row r=\"1\"><c r=\"A1\"><v>1</v></c></row>",
            None,
        );
        let r = reflect_text(&doctype, "d.xlsx").unwrap_err();
        assert_eq!(r.id, "reflect-xml-refused");
        assert!(
            r.text
                .contains("xl/worksheets/sheet1.xml, it declares a DOCTYPE"),
            "{}",
            r.text
        );
        // A shape this version does not read is named with its cell.
        let table = one_sheet("<row r=\"1\"><c r=\"B2\"><f t=\"dataTable\" ref=\"B2:B3\" dt2D=\"0\" r1=\"A1\"/><v>1</v></c></row>", None);
        let r = reflect_text(&table, "d.xlsx").unwrap_err();
        assert_eq!(r.id, "reflect-unsupported");
        assert!(
            r.text.contains("cell B2 holds a data-table formula"),
            "{}",
            r.text
        );
        let orphan = one_sheet(
            "<row r=\"1\"><c r=\"A1\"><f t=\"shared\" si=\"4\"/><v>1</v></c></row>",
            None,
        );
        assert!(reflect_text(&orphan, "d.xlsx")
            .unwrap_err()
            .text
            .contains("cell A1 continues a shared formula"));
        let past = one_sheet(
            "<row r=\"1\"><c r=\"A1\" t=\"s\"><v>3</v></c></row>",
            Some("<si><t>a</t></si>"),
        );
        assert!(reflect_text(&past, "d.xlsx")
            .unwrap_err()
            .text
            .contains("shared string 3, past the end"));
        // A damaged part is the archive's fault, not a shape's.
        let mut damaged = one_sheet("<row r=\"1\"><c r=\"A1\"><v>1</v></c></row>", None);
        let e = zip::entries(&damaged)
            .unwrap()
            .into_iter()
            .find(|e| e.name == "xl/worksheets/sheet1.xml")
            .unwrap();
        damaged[e.data_offset + 10] ^= 1;
        let r = reflect_text(&damaged, "d.xlsx").unwrap_err();
        assert_eq!(r.id, "reflect-not-a-workbook");
        assert!(r.text.contains("checksum does not match"), "{}", r.text);
    }

    /// The treaty's oracle 8 (conformance/README.md, the amendment of
    /// 2026-10-03 for PORT.8): each fixture's relations are its golden
    /// whole, line endings aside. The goldens are the core's own output,
    /// blessed by the owner reading each fixture in Excel; a deliberate
    /// change regenerates them with the door and raises the check's floors.
    #[test]
    fn the_reflect_goldens_are_reproduced() {
        let rows: [(&str, &[u8], &str); 5] = [
            // The first build golden as Excel 365 saved it (the owner's
            // live pass, 2026-10-04): a cell row for every formula's cached
            // value, and the stamp's long text as _xlfn._LONGTEXT.
            (
                "scripts/reflect/saved.xlsx",
                include_bytes!("../../../scripts/reflect/saved.xlsx"),
                include_str!("../../../scripts/reflect/saved_relations.vla"),
            ),
            (
                "scripts/reflect/fixture.xlsx",
                include_bytes!("../../../scripts/reflect/fixture.xlsx"),
                include_str!("../../../scripts/reflect/fixture_relations.vla"),
            ),
            (
                "scripts/build/fixture_golden.xlsx",
                include_bytes!("../../../scripts/build/fixture_golden.xlsx"),
                include_str!("../../../scripts/reflect/build_fixture_relations.vla"),
            ),
            (
                "scripts/build/into_golden.xlsx",
                include_bytes!("../../../scripts/build/into_golden.xlsx"),
                include_str!("../../../scripts/reflect/build_into_relations.vla"),
            ),
            (
                "scripts/build/model.xlsx",
                include_bytes!("../../../scripts/build/model.xlsx"),
                include_str!("../../../scripts/reflect/model_relations.vla"),
            ),
        ];
        for (label, bytes, golden) in rows {
            let got = reflect_text(bytes, label)
                .unwrap_or_else(|r| panic!("{label} did not read: {}", r.text));
            let want = golden.replace("\r\n", "\n");
            if got != want {
                let line = got
                    .lines()
                    .zip(want.lines())
                    .position(|(a, b)| a != b)
                    .map(|i| i + 1)
                    .unwrap_or(got.lines().count().min(want.lines().count()) + 1);
                panic!("{label}: the relations differ from the golden at line {line}");
            }
        }
    }

    #[test]
    fn paths_resolve_as_the_package_names_them() {
        assert_eq!(
            resolve("xl/", "worksheets/sheet1.xml"),
            "xl/worksheets/sheet1.xml"
        );
        assert_eq!(
            resolve("xl/", "/xl/worksheets/sheet1.xml"),
            "xl/worksheets/sheet1.xml"
        );
        assert_eq!(
            resolve("xl/worksheets/", "../tables/table1.xml"),
            "xl/tables/table1.xml"
        );
        assert_eq!(
            resolve("xl/worksheets/", "./a/../b.xml"),
            "xl/worksheets/b.xml"
        );
        assert_eq!(rels_of("xl/workbook.xml"), "xl/_rels/workbook.xml.rels");
        assert_eq!(
            rels_of("xl/worksheets/sheet1.xml"),
            "xl/worksheets/_rels/sheet1.xml.rels"
        );
        assert_eq!(
            file_name_of("file:///C:/Models/Rates%20Q1.xlsx"),
            "Rates Q1.xlsx"
        );
        assert_eq!(file_name_of("Rates.xlsx"), "Rates.xlsx");
        assert_eq!(file_name_of("..\\Shared\\Rates.xlsx"), "Rates.xlsx");
        assert_eq!(file_name_of("odd%2"), "odd%2");
    }
}
