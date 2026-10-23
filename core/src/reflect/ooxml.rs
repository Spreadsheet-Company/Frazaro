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
//!
//! The looks (KERNEL.57, `--looks`): the styles part and the theme, read
//! once a workbook and only for the looks, each `cellXfs` entry a format of
//! the record's at its own index ([`super::looks::Look`]), its own font,
//! fill, border, alignment and number format whatever its `apply` flags and
//! its cell style say, what the Normal style's font says left unsaid; and a
//! sheet's part read twice, a first pass for its first view's gridlines,
//! its columns, the rows a person set a height or hid, the extent of the
//! cells the walk prints and the formats they use, and what the record
//! cannot hold, counted: a cell that keeps a format and no value, a look
//! the record has no word for, a merged range, a conditional format, a row
//! formatted whole.

use std::borrow::Cow;
use std::cell::OnceCell;
use std::collections::{HashMap, HashSet};

use super::cursor::{attr, element_text, skip_element, Cursor, Event};
use super::looks::{
    font_size, iso_serial, rgb_from_hex, width_chars, Color, ColorRef, Colors, ColumnSpec, Look,
    LooksStats, Resolved, SheetScan, THEME_SLOTS,
};
use super::{
    emit_formula, emit_formula_with, NameInfo, Record, Row, SheetInfo, SheetStats, Sink, Source,
    Summary, TableInfo, Value, Visibility,
};
use crate::intrinsics::fold;
use crate::messages::{raise, Refusal};
use crate::refers;
use crate::sheet::merge::part_text_raw;
use crate::sheet::zip::{self, Entry};
use crate::sheet::{
    cell_ref, number_text, parse_a1_range, xlfn, Edge, Horizontal, Line, Rgb, RowSettings, Vertical,
};
use vla_lang::rows::{datum, ArrayKind};

/// The most distinct R1C1 formulas a count keeps, per sheet and per
/// workbook; past it the count stands still (AXM.1's own cap).
const WORKBOOK: &str = "xl/workbook.xml";
const WORKBOOK_RELS: &str = "xl/_rels/workbook.xml.rels";
const SHARED_STRINGS: &str = "xl/sharedStrings.xml";

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
    /// The cell-metadata index a dynamic-array formula's `cm` names, from
    /// the metadata part, and whether the package holds that part
    /// (KERNEL.8, slice b).
    dynamic_cm: Option<u32>,
    has_metadata: bool,
    /// The error each value-metadata block names, by a cell's `vm` less one:
    /// `#CALC!` and `#SPILL!`, which Excel writes as `#VALUE!` beside a rich
    /// value (KERNEL.8, slice f).
    rich_errors: Vec<Option<String>>,
    /// Whether the workbook counts its dates from 1904 (`workbookPr`'s
    /// `date1904`, KERNEL.8, slice d).
    date1904: bool,
    /// Whether each shared string is rich text, a run with properties of
    /// its own, which no row says (KERNEL.57).
    rich: Vec<bool>,
    /// The styles part and the theme part, by the workbook's relationships
    /// (KERNEL.57).
    styles_part: Option<String>,
    theme_part: Option<String>,
    /// The cell formats as the looks read them, read the first time a walk
    /// asks for them (KERNEL.57).
    looks: OnceCell<Result<StyleTable, Refusal>>,
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

pub(crate) fn not_a_workbook(label: &str, why: &str) -> Refusal {
    raise("reflect-not-a-workbook", &[("path", label), ("why", why)])
}

pub(crate) fn xml_refused(label: &str, part: &str, why: &str) -> Refusal {
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
pub(crate) fn step<'t>(
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

/// The shared strings, each its text, and whether each is rich text, a run
/// with properties of its own, which no row says (KERNEL.57).
fn read_strings(text: &str, label: &str, part: &str) -> Result<(Vec<String>, Vec<bool>), Refusal> {
    let mut cursor = Cursor::new(text);
    let mut out = Vec::new();
    let mut rich = Vec::new();
    while let Some(event) = step(&mut cursor, label, part)? {
        if let Event::Start {
            name: "si", empty, ..
        } = event
        {
            if empty {
                out.push(String::new());
                rich.push(false);
            } else {
                let start = cursor.position();
                out.push(inner_text(&mut cursor, &["rPh"], true, label, part)?);
                rich.push(has_run_properties(&text[start..cursor.position()]));
            }
        }
    }
    Ok((out, rich))
}

/// Whether a string's markup holds a run's own properties, `<rPr>` under
/// any prefix: a run of bold or of color inside a cell's text.
fn has_run_properties(xml: &str) -> bool {
    xml.contains("<rPr") || xml.contains(":rPr")
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
        let mut date1904 = false;
        let mut cursor = Cursor::new(&workbook);
        while let Some(event) = step(&mut cursor, label, WORKBOOK)? {
            match event {
                Event::Start {
                    name: "workbookPr",
                    attrs,
                    ..
                } => {
                    date1904 =
                        matches!(attr(attrs, "date1904").as_deref(), Some("1") | Some("true"));
                }
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
        let (strings, rich, string_bytes) = match text(&strings_part)? {
            None => (Vec::new(), Vec::new(), 0),
            Some(t) => {
                let (strings, rich) = read_strings(&t, label, &strings_part)?;
                (strings, rich, t.len())
            }
        };
        // The styles part and the theme, read only for the looks, the first
        // time a walk asks for them (KERNEL.57).
        let part_of = |suffix: &str, default: &str| {
            rels.iter()
                .find(|r| r.kind.ends_with(suffix))
                .map(|r| resolve("xl/", &r.target))
                .or_else(|| has(default).then(|| default.to_string()))
        };
        let styles_part = part_of("/styles", "xl/styles.xml");
        let theme_part = part_of("/theme", "xl/theme/theme1.xml");

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

        // The metadata part, for the index a dynamic-array formula's `cm`
        // names (KERNEL.8, slice b).
        let metadata_part = rels
            .iter()
            .find(|r| r.kind.ends_with("/sheetMetadata"))
            .map(|r| resolve("xl/", &r.target))
            .unwrap_or_else(|| "xl/metadata.xml".to_string());
        let metadata = text(&metadata_part)?;
        let has_metadata = metadata.is_some();
        let dynamic_cm = metadata
            .as_deref()
            .and_then(crate::sheet::merge::dynamic_cm_of);
        // The rich values an error with no cell error value of its own names
        // through a cell's `vm` (KERNEL.8, slice f).
        let rich_part = |suffix: &str| {
            rels.iter()
                .find(|r| r.kind.ends_with(suffix))
                .map(|r| resolve("xl/", &r.target))
        };
        let rich_errors = match (
            metadata.as_deref(),
            rich_part("/rdRichValue"),
            rich_part("/rdRichValueStructure"),
        ) {
            (Some(m), Some(v), Some(s)) => match (text(&v)?, text(&s)?) {
                (Some(values), Some(structures)) => {
                    crate::sheet::merge::rich_errors_of(m, &values, &structures)
                }
                _ => Vec::new(),
            },
            _ => Vec::new(),
        };

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
            dynamic_cm,
            has_metadata,
            rich_errors,
            date1904,
            rich,
            styles_part,
            theme_part,
            looks: OnceCell::new(),
        })
    }

    /// An array formula's kind by its cell's `cm` (KERNEL.8, slice b):
    /// dynamic where it names the dynamic-array metadata, or names anything
    /// in a package with no metadata part to say otherwise; legacy, entered
    /// with Ctrl+Shift+Enter, where it is absent or names other metadata.
    fn array_kind(&self, cm: Option<u32>) -> ArrayKind {
        match (cm, self.dynamic_cm) {
            (Some(c), Some(d)) if c == d => ArrayKind::Dynamic,
            (Some(_), None) if !self.has_metadata => ArrayKind::Dynamic,
            _ => ArrayKind::Legacy,
        }
    }
}

impl<'a> Source for Package<'a> {
    fn label(&self) -> &str {
        &self.label
    }

    /// The sheets in tab order.
    fn sheets(&self) -> &[SheetInfo] {
        &self.sheets
    }

    /// Every defined name, placeholders included, in the file's order.
    fn names(&self) -> &[NameInfo] {
        &self.names
    }

    /// Every Table, in the order the sheets' relationships name them.
    fn tables(&self) -> &[TableInfo] {
        &self.tables
    }

    /// The linked workbooks' file names, in the order `[1]`, `[2]`, ... of a
    /// formula's text; an empty name for a link of another kind.
    fn external_books(&self) -> &[String] {
        &self.external_books
    }

    fn date1904(&self) -> bool {
        self.date1904
    }

    fn summary(&self) -> Summary {
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

    /// One sheet's cells in document order, to the sink: a `cell` row for a
    /// value the file holds, a `formula` row for a formula, then its
    /// `refers` rows. The sheet's part is read here and dropped on return;
    /// what the walk counted comes back, with the workbook-wide set of
    /// distinct R1C1 formulas carried across sheets for `--counts`.
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
    /// its part read twice, the first pass ([`Package::first_pass`]) for its
    /// own rows of the record, then the walk, each cell's rows in the
    /// record's spelling.
    fn walk_sheet_looks(
        &self,
        index: usize,
        sink: &mut dyn Sink,
        book_distinct: &mut HashSet<String>,
    ) -> Result<(SheetStats, LooksStats), Refusal> {
        let table = self.style_table()?;
        self.walk(index, sink, book_distinct, Some(table))
    }
}

impl<'a> Package<'a> {
    /// One sheet's walk: the relations' rows, or, given the workbook's cell
    /// formats (KERNEL.57), the sheet as the record prints it whole, its own
    /// rows of the record before any cell's, a date as its serial, a
    /// formula's saved value its `value` row after the formula, and each
    /// cell's `style` row after its rows and before its `refers` rows.
    fn walk(
        &self,
        index: usize,
        sink: &mut dyn Sink,
        book_distinct: &mut HashSet<String>,
        looks: Option<&StyleTable>,
    ) -> Result<(SheetStats, LooksStats), Refusal> {
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
        // The looks' first pass, and the sheet's own rows of the record
        // before any cell's (KERNEL.57).
        let scan = match looks {
            Some(table) => {
                let mut scan = self.first_pass(&text, table, label, part)?;
                scan.print(&info.name, &table.looks, sink);
                Some(scan)
            }
            None => None,
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
                    // KERNEL.8, slice e: a hidden row, which `SUBTOTAL`'s
                    // 101 to 111 pass by.
                    if matches!(attr(attrs, "hidden").as_deref(), Some("1") | Some("true")) {
                        sink.hidden_row(&info.name, row_no);
                    }
                }
                Event::Start {
                    name: "c",
                    attrs,
                    empty,
                } if in_data => {
                    let mut cell = read_c(
                        &mut cursor,
                        &text,
                        attrs,
                        empty,
                        (row_no, col_no),
                        label,
                        part,
                    )?;
                    let (row, col) = (cell.row, cell.col);
                    row_no = row;
                    col_no = col;
                    let addr = cell_ref(row, col);
                    let (formula, array_anchor) =
                        cell.formula_text(&addr, &mut shared, &mut stats, label, part)?;
                    let value = cell.value(self, &addr, label, part)?;
                    if value.is_none() && formula.is_none() {
                        continue;
                    }
                    stats.cells += 1;
                    stats.rows = stats.rows.max(row);
                    stats.columns = stats.columns.max(col);
                    let formula = formula.map(|f_text| {
                        let text = format!("={}", xlfn::strip_future_prefixes(&f_text));
                        refers::resolve_books(&text, &self.external_books)
                    });
                    let sheet = info.name.as_str();
                    match looks {
                        None => {
                            if let Some(value) = &value {
                                sink.row(&Row::Cell {
                                    sheet,
                                    addr: &addr,
                                    value,
                                });
                            }
                            if let Some(f) = &formula {
                                emit_formula(
                                    sink,
                                    sheet,
                                    &addr,
                                    f,
                                    row,
                                    col,
                                    &mut stats,
                                    &mut sheet_distinct,
                                    book_distinct,
                                );
                            }
                        }
                        Some(table) => {
                            let style = if table.part { cell.style } else { 0 };
                            let value = value.map(|v| self.record_value(v));
                            match &formula {
                                Some(f) => {
                                    let saved = value.as_ref().map(datum);
                                    emit_formula_with(
                                        sink,
                                        sheet,
                                        &addr,
                                        f,
                                        row,
                                        col,
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
                                            if style != 0 {
                                                sink.record(&Record::Style {
                                                    sheet,
                                                    addr: &addr,
                                                    index: style,
                                                });
                                            }
                                        },
                                    );
                                }
                                None => {
                                    if let Some(value) = &value {
                                        sink.row(&Row::Cell {
                                            sheet,
                                            addr: &addr,
                                            value,
                                        });
                                    }
                                    if style != 0 {
                                        sink.record(&Record::Style {
                                            sheet,
                                            addr: &addr,
                                            index: style,
                                        });
                                    }
                                }
                            }
                        }
                    }
                    if formula.is_some() && array_anchor {
                        sink.array(sheet, &addr, self.array_kind(cell.cm));
                    }
                }
                _ => {}
            }
        }
        stats.distinct_r1c1 = sheet_distinct.len() as u64;
        Ok((stats, scan.map(|s| s.stats).unwrap_or_default()))
    }

    /// A value as the record spells it (KERNEL.57): a date, the file's ISO
    /// text, as Excel's serial in the workbook's calendar; any other as the
    /// relations spell it.
    fn record_value(&self, value: Value) -> Value {
        if let Value::Date(iso) = &value {
            if let Some(n) = iso_serial(iso, self.date1904).and_then(number_text) {
                return Value::Number(n);
            }
        }
        value
    }

    /// The looks' first pass over a sheet's part (KERNEL.57): its first
    /// view's gridlines, its own default width and height, its columns, the
    /// rows a person set a height for or hid, the extent of the cells the
    /// walk prints and the formats they use, and what the record cannot hold
    /// among them, counted; its merged ranges and conditional formats after
    /// its cells. A cell or a column naming a format past the styles part's
    /// list is refused by name.
    fn first_pass(
        &self,
        text: &str,
        table: &StyleTable,
        label: &str,
        part: &str,
    ) -> Result<SheetScan, Refusal> {
        let mut scan = SheetScan::new();
        let mut cursor = Cursor::new(text);
        let mut views = 0u32;
        let mut in_data = false;
        let mut row_no: u32 = 0;
        let mut col_no: u32 = 0;
        while let Some(event) = step(&mut cursor, label, part)? {
            match event {
                // An extension's elements, conditional formats Excel writes
                // twice among them, are no look of the record's.
                Event::Start {
                    name: "extLst",
                    empty: false,
                    ..
                } => skip_element(&mut cursor).map_err(|why| xml_refused(label, part, why))?,
                Event::Start {
                    name: "sheetView",
                    attrs,
                    ..
                } if !in_data => {
                    views += 1;
                    if views == 1 && truthy_off(attr(attrs, "showGridLines")) {
                        scan.gridlines = false;
                    }
                }
                Event::Start {
                    name: "sheetFormatPr",
                    attrs,
                    ..
                } => {
                    // A sheet's own default width or height, which no row
                    // of the record says.
                    if attr(attrs, "defaultColWidth").is_some() {
                        scan.stats.unheld += 1;
                    }
                    if attr(attrs, "baseColWidth")
                        .and_then(|v| v.trim().parse::<f64>().ok())
                        .is_some_and(|w| w != 8.0)
                    {
                        scan.stats.unheld += 1;
                    }
                    if truthy(attr(attrs, "customHeight")) {
                        scan.stats.unheld += 1;
                    }
                }
                Event::Start {
                    name: "col", attrs, ..
                } if !in_data => {
                    let number = |name: &str| attr(attrs, name).and_then(|v| v.trim().parse().ok());
                    let (Some(min), Some(max)) = (number("min"), number("max")) else {
                        continue;
                    };
                    let width = attr(attrs, "width").and_then(|v| v.trim().parse::<f64>().ok());
                    let hidden = truthy(attr(attrs, "hidden")) || width == Some(0.0);
                    let style: Option<u32> = if table.part { number("style") } else { None };
                    if let Some(s) = style {
                        table.check(s, label, part, &format!("the columns {min} to {max}"))?;
                    }
                    scan.columns.push(ColumnSpec {
                        min,
                        max,
                        width: if hidden {
                            None
                        } else {
                            width.and_then(width_chars)
                        },
                        hidden,
                        style,
                    });
                }
                Event::Start {
                    name: "sheetData",
                    empty,
                    ..
                } => in_data = !empty,
                Event::End { name: "sheetData" } => in_data = false,
                Event::Start {
                    name: "row", attrs, ..
                } if in_data => {
                    row_no = attr(attrs, "r")
                        .and_then(|v| v.parse::<u32>().ok())
                        .unwrap_or(row_no + 1);
                    col_no = 0;
                    // A height a person set; one Excel fitted to the row's
                    // text is the host's to fit again (decision 10).
                    let height = if truthy(attr(attrs, "customHeight")) {
                        attr(attrs, "ht").and_then(|v| v.trim().parse::<f64>().ok())
                    } else {
                        None
                    };
                    let hidden = truthy(attr(attrs, "hidden"));
                    if height.is_some() || hidden {
                        scan.rows.insert(row_no, RowSettings { height, hidden });
                    }
                    if truthy(attr(attrs, "customFormat"))
                        && attr(attrs, "s")
                            .and_then(|v| v.trim().parse::<u32>().ok())
                            .is_some_and(|s| s != 0)
                    {
                        scan.stats.row_formats += 1;
                    }
                }
                Event::Start {
                    name: "c",
                    attrs,
                    empty,
                } if in_data => {
                    let cell = read_c(
                        &mut cursor,
                        text,
                        attrs,
                        empty,
                        (row_no, col_no),
                        label,
                        part,
                    )?;
                    row_no = cell.row;
                    col_no = cell.col;
                    let style = if table.part { cell.style } else { 0 };
                    if !cell.prints(self) {
                        if style != 0 {
                            scan.stats.formatted_empty += 1;
                        }
                        continue;
                    }
                    if style != 0 {
                        let what = format!("the cell {}", cell_ref(cell.row, cell.col));
                        table.check(style, label, part, &what)?;
                        scan.stats.styled += 1;
                    }
                    scan.used.insert(style);
                    let look = &table.looks[style as usize];
                    let rich = cell.inline_rich
                        || (cell.kind.as_deref() == Some("s")
                            && cell
                                .value_text
                                .as_deref()
                                .and_then(|v| v.trim().parse::<usize>().ok())
                                .and_then(|i| self.rich.get(i).copied())
                                .unwrap_or(false));
                    if look.unheld || rich {
                        scan.stats.unheld += 1;
                    }
                    if look.undrawn.is_some() {
                        scan.stats.undrawn += 1;
                    }
                    scan.hold(cell.row, cell.col);
                }
                Event::Start {
                    name: "mergeCell", ..
                } => scan.stats.merged += 1,
                Event::Start { name: "cfRule", .. } => scan.stats.conditional += 1,
                _ => {}
            }
        }
        Ok(scan)
    }

    /// The workbook's cell formats as the looks read them, read the first
    /// time a walk asks (KERNEL.57).
    fn style_table(&self) -> Result<&StyleTable, Refusal> {
        self.looks
            .get_or_init(|| self.read_style_table())
            .as_ref()
            .map_err(Clone::clone)
    }

    /// The styles part and the theme read into the workbook's cell formats;
    /// a package with no styles part reads every cell at format 0.
    fn read_style_table(&self) -> Result<StyleTable, Refusal> {
        let label = self.label.as_str();
        let read = |part: &str| {
            part_text_raw(self.bytes, &self.entries, part)
                .map_err(|why| not_a_workbook(label, &why))
        };
        let none = StyleTable {
            looks: vec![Look::default()],
            part: false,
        };
        let Some(styles_part) = self.styles_part.as_deref() else {
            return Ok(none);
        };
        let Some(text) = read(styles_part)? else {
            return Ok(none);
        };
        let theme = match self.theme_part.as_deref() {
            Some(theme_part) => match read(theme_part)? {
                Some(t) => Some(read_theme(&t, label, theme_part)?),
                None => None,
            },
            None => None,
        };
        Ok(read_styles(&text, label, styles_part)?.table(theme))
    }
}

/// One `<c>` element read whole: what the walk and the looks' first pass
/// both ask of a cell, before its formula and its value are made of it.
struct CellRead {
    row: u32,
    col: u32,
    /// Its `s`, the cell format it names, 0 where it names none.
    style: u32,
    /// Its `t`, the kind of its value.
    kind: Option<String>,
    cm: Option<u32>,
    vm: Option<usize>,
    /// Its `<f>`: the formula's kind, its shared index and its text.
    formula: Option<(Option<String>, Option<String>, String)>,
    /// Its `<v>`'s text, empty where the element is empty.
    value_text: Option<String>,
    /// Its `<is>`'s text.
    inline: Option<String>,
    /// Whether its `<is>` is rich text (KERNEL.57).
    inline_rich: bool,
}

/// A `<c>` element from its start tag to its end: its place, by its `r` or
/// by where it stands after the cell before it, and what it holds; `text`
/// is the part the cursor walks.
fn read_c(
    cursor: &mut Cursor,
    text: &str,
    attrs: &str,
    empty: bool,
    (row_no, col_no): (u32, u32),
    label: &str,
    part: &str,
) -> Result<CellRead, Refusal> {
    let (row, col) = match attr(attrs, "r").and_then(|r| parse_a1_range(&r)) {
        Some(a) => (a.top, a.left),
        None => (row_no, col_no + 1),
    };
    let mut cell = CellRead {
        row,
        col,
        style: attr(attrs, "s")
            .and_then(|v| v.trim().parse().ok())
            .unwrap_or(0),
        kind: attr(attrs, "t").map(Cow::into_owned),
        cm: attr(attrs, "cm").and_then(|v| v.trim().parse().ok()),
        vm: attr(attrs, "vm").and_then(|v| v.trim().parse().ok()),
        formula: None,
        value_text: None,
        inline: None,
        inline_rich: false,
    };
    if !empty {
        loop {
            match step(cursor, label, part)? {
                None => return Err(xml_refused(label, part, "a cell never closes")),
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
                        inner_text(cursor, &[], false, label, part)?
                    };
                    cell.formula = Some((t, si, f_text));
                }
                Some(Event::Start {
                    name: "v",
                    empty: v_empty,
                    ..
                }) => {
                    cell.value_text = Some(if v_empty {
                        String::new()
                    } else {
                        inner_text(cursor, &[], false, label, part)?
                    });
                }
                Some(Event::Start {
                    name: "is",
                    empty: is_empty,
                    ..
                }) => {
                    if is_empty {
                        cell.inline = Some(String::new());
                    } else {
                        let start = cursor.position();
                        cell.inline = Some(inner_text(cursor, &["rPh"], true, label, part)?);
                        cell.inline_rich = has_run_properties(&text[start..cursor.position()]);
                    }
                }
                Some(Event::Start { empty: false, .. }) => {
                    skip_element(cursor).map_err(|why| xml_refused(label, part, why))?;
                }
                Some(Event::End { name: "c" }) => break,
                _ => {}
            }
        }
    }
    Ok(cell)
}

impl CellRead {
    /// The cell's formula as the walk prints it, its text with no `=`: a
    /// shared formula's child the first cell's text with its references
    /// moved by the cell's distance, an array formula's anchor its text and
    /// a cell it spills into none; and whether the cell is an array's
    /// anchor.
    fn formula_text(
        &mut self,
        addr: &str,
        shared: &mut Vec<Option<(String, u32, u32)>>,
        stats: &mut SheetStats,
        label: &str,
        part: &str,
    ) -> Result<(Option<String>, bool), Refusal> {
        let Some((t, si, f_text)) = self.formula.take() else {
            return Ok((None, false));
        };
        let (row, col) = (self.row, self.col);
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
                    Ok((
                        Some(refers::shift_a1_references(
                            master,
                            i64::from(row) - i64::from(*m_row),
                            i64::from(col) - i64::from(*m_col),
                        )),
                        false,
                    ))
                } else {
                    if shared.len() <= i {
                        shared.resize(i + 1, None);
                    }
                    shared[i] = Some((f_text.clone(), row, col));
                    Ok((Some(f_text), false))
                }
            }
            Some("array") => {
                if f_text.is_empty() {
                    Ok((None, false))
                } else {
                    stats.array_anchors += 1;
                    Ok((Some(f_text), true))
                }
            }
            Some("dataTable") => Err(unsupported(
                label,
                part,
                &format!("cell {addr} holds a data-table formula"),
            )),
            _ => {
                if f_text.is_empty() {
                    Ok((None, false))
                } else {
                    Ok((Some(f_text), false))
                }
            }
        }
    }

    /// The value the file holds for the cell, by its kind, as the walk
    /// prints it.
    fn value(
        &mut self,
        package: &Package,
        addr: &str,
        label: &str,
        part: &str,
    ) -> Result<Option<Value>, Refusal> {
        let kind = self.kind.as_deref();
        // A formula's text value may be empty, `t="str"` and an empty `<v>`,
        // as a host saves `=""` (KERNEL.8, slice c); every other kind's
        // empty `<v>` holds nothing.
        let raw = self.value_text.take();
        let empty_text = kind == Some("str") && raw.is_some();
        let value_text = raw.filter(|s| !s.is_empty());
        Ok(match kind {
            Some("s") => {
                match value_text {
                    Some(s) => {
                        let i: usize = s.trim().parse().map_err(|_| {
                        unsupported(
                            label,
                            part,
                            &format!("cell {addr} names a shared string by {s}, which is not an index"),
                        )
                    })?;
                        let text = package.strings.get(i).cloned().ok_or_else(|| {
                        unsupported(
                            label,
                            part,
                            &format!("cell {addr} names shared string {i}, past the end of the table"),
                        )
                    })?;
                        Some(Value::Text(text))
                    }
                    None => None,
                }
            }
            Some("str") => match value_text {
                Some(t) => Some(Value::Text(t)),
                None if empty_text => Some(Value::Text(String::new())),
                None => None,
            },
            Some("b") => value_text.map(|s| {
                let s = s.trim();
                Value::Bool(s == "1" || s.eq_ignore_ascii_case("true"))
            }),
            // An error with no cell error value of its own, read off the
            // rich value its `vm` names, where Excel wrote `#VALUE!`
            // (KERNEL.8, slice f).
            Some("e") => match self
                .vm
                .and_then(|i| package.rich_errors.get(i.checked_sub(1)?))
                .and_then(Option::as_ref)
            {
                Some(rich) => Some(Value::Error(rich.clone())),
                None => value_text.map(|s| Value::Error(s.trim().to_string())),
            },
            Some("d") => value_text.map(|s| Value::Date(s.trim().to_string())),
            Some("inlineStr") => self.inline.take().map(Value::Text),
            _ => value_text.map(|s| Value::Number(s.trim().to_string())),
        })
    }

    /// Whether the walk prints a row for the cell, a value the file holds
    /// or a formula, asked before its value and its formula are made, with
    /// none of the sheet's shared formulas: the looks' first pass's
    /// question (KERNEL.57), answered as the walk answers it.
    fn prints(&self, package: &Package) -> bool {
        if let Some((t, _, f_text)) = &self.formula {
            let formula = match t.as_deref() {
                Some("shared") | Some("dataTable") => true,
                _ => !f_text.is_empty(),
            };
            if formula {
                return true;
            }
        }
        let filled = self.value_text.as_deref().is_some_and(|s| !s.is_empty());
        match self.kind.as_deref() {
            Some("str") => self.value_text.is_some(),
            Some("e") => {
                filled
                    || self
                        .vm
                        .and_then(|i| package.rich_errors.get(i.checked_sub(1)?))
                        .and_then(Option::as_ref)
                        .is_some()
            }
            Some("inlineStr") => self.inline.is_some(),
            _ => filled,
        }
    }
}

/// Whether an attribute says yes, `1` or `true`.
fn truthy(value: Option<Cow<'_, str>>) -> bool {
    matches!(value.as_deref(), Some("1") | Some("true"))
}

/// Whether an attribute says no, `0` or `false`.
fn truthy_off(value: Option<Cow<'_, str>>) -> bool {
    matches!(value.as_deref(), Some("0") | Some("false"))
}

/// A package's cell formats as the looks read them (KERNEL.57): each
/// `cellXfs` entry as the record says it, at its own index, format 0 always
/// among them.
#[derive(Debug)]
pub(crate) struct StyleTable {
    looks: Vec<Look>,
    /// Whether the package has a styles part: one with none reads every
    /// cell at format 0.
    part: bool,
}

impl StyleTable {
    /// A format a cell or a column names, refused by name past the styles
    /// part's list or past the formats a grid holds.
    fn check(&self, index: u32, label: &str, part: &str, what: &str) -> Result<(), Refusal> {
        if index >= vla_lang::machine::load::FORMATS {
            return Err(unsupported(
                label,
                part,
                &format!("{what} names the cell format {index}, past the 65,536 a grid holds"),
            ));
        }
        if index as usize >= self.looks.len() {
            return Err(unsupported(
                label,
                part,
                &format!(
                    "{what} names the cell format {index}, past the {} its styles part holds",
                    self.looks.len()
                ),
            ));
        }
        Ok(())
    }
}

/// A font of a styles part, as the looks read it.
#[derive(Clone, Debug, Default, PartialEq)]
struct FileFont {
    bold: bool,
    italic: bool,
    size: Option<f64>,
    color: Option<ColorRef>,
    name: Option<String>,
    /// An underline, a strike, a raised or lowered text, an outline, a
    /// shadow, a condensed or extended face: what no word of the record's
    /// says.
    unheld: bool,
}

/// A fill of a styles part: a solid pattern's foreground, or another fill
/// the record cannot hold.
#[derive(Clone, Debug, Default)]
struct FileFill {
    solid: bool,
    fg: Option<ColorRef>,
    unheld: bool,
}

/// A border of a styles part: each of a cell's four edges, left, right, top
/// and bottom, its line's name and its color; a diagonal, which the record
/// cannot hold.
#[derive(Clone, Debug, Default)]
struct FileBorder {
    edges: [Option<(String, Option<ColorRef>)>; 4],
    unheld: bool,
}

/// A cell format of a styles part's `cellXfs`: its ids and its alignment.
#[derive(Clone, Debug, Default)]
struct FileXf {
    num_fmt: u32,
    font: usize,
    fill: usize,
    border: usize,
    horizontal: Option<String>,
    vertical: Option<String>,
    wrap: bool,
    /// An indent, a rotation, shrinking to fit, a right-to-left order or a
    /// justified last line: what no word of the record's says.
    unheld: bool,
}

/// A styles part read for the looks.
#[derive(Debug, Default)]
struct Parsed {
    codes: HashMap<u32, String>,
    fonts: Vec<FileFont>,
    fills: Vec<FileFill>,
    borders: Vec<FileBorder>,
    /// Each cell style's format's font, `cellStyleXfs` by index.
    style_fonts: Vec<usize>,
    xfs: Vec<FileXf>,
    /// The Normal cell style's format, `cellStyles`' `builtinId="0"`.
    normal: Option<usize>,
    /// The workbook's own palette, where it has one.
    palette: Vec<Rgb>,
}

/// A color element's attributes as a color, `None` for one naming none.
fn color_of(attrs: &str) -> Option<ColorRef> {
    let tint = attr(attrs, "tint")
        .and_then(|v| v.trim().parse::<f64>().ok())
        .unwrap_or(0.0);
    let color = if truthy(attr(attrs, "auto")) {
        Color::Auto
    } else if let Some(rgb) = attr(attrs, "rgb") {
        Color::Rgb(rgb_from_hex(&rgb)?)
    } else if let Some(t) = attr(attrs, "theme").and_then(|v| v.trim().parse().ok()) {
        Color::Theme(t)
    } else {
        Color::Indexed(attr(attrs, "indexed").and_then(|v| v.trim().parse().ok())?)
    };
    Some(ColorRef { color, tint })
}

/// Whether a font's flag element says yes: present, and not `val="0"` or
/// `val="false"`.
fn flag_on(attrs: &str) -> bool {
    !truthy_off(attr(attrs, "val"))
}

/// The list of a styles part being read.
#[derive(Clone, Copy, PartialEq, Eq)]
enum List {
    None,
    NumFmts,
    Fonts,
    Fills,
    Borders,
    StyleXfs,
    CellXfs,
    CellStyles,
    IndexedColors,
}

/// A styles part read: its number formats, fonts, fills, borders, cell
/// styles' fonts, cell formats, the Normal style and the palette; its
/// differential formats, its table styles and its extensions passed over.
fn read_styles(text: &str, label: &str, part: &str) -> Result<Parsed, Refusal> {
    let mut parsed = Parsed::default();
    let mut list = List::None;
    let mut cursor = Cursor::new(text);
    let mut font: Option<FileFont> = None;
    let mut fill: Option<FileFill> = None;
    let mut in_gradient = false;
    let mut border: Option<FileBorder> = None;
    let mut side: Option<usize> = None;
    let mut xf: Option<FileXf> = None;
    let skip =
        |cursor: &mut Cursor| skip_element(cursor).map_err(|why| xml_refused(label, part, why));
    while let Some(event) = step(&mut cursor, label, part)? {
        match event {
            Event::Start { name, attrs, empty } => match (name, list) {
                ("extLst" | "dxfs" | "tableStyles" | "mruColors", _) if !empty => {
                    skip(&mut cursor)?
                }
                ("numFmts", _) => list = if empty { List::None } else { List::NumFmts },
                ("numFmt", List::NumFmts) => {
                    if let (Some(id), Some(code)) = (
                        attr(attrs, "numFmtId").and_then(|v| v.trim().parse::<u32>().ok()),
                        attr(attrs, "formatCode"),
                    ) {
                        parsed.codes.insert(id, code.into_owned());
                    }
                }
                ("fonts", _) => list = if empty { List::None } else { List::Fonts },
                ("font", List::Fonts) => {
                    if empty {
                        parsed.fonts.push(FileFont::default());
                    } else {
                        font = Some(FileFont::default());
                    }
                }
                (child, List::Fonts) => {
                    if let Some(f) = font.as_mut() {
                        match child {
                            "b" => f.bold = flag_on(attrs),
                            "i" => f.italic = flag_on(attrs),
                            "strike" | "outline" | "shadow" | "condense" | "extend" => {
                                if flag_on(attrs) {
                                    f.unheld = true;
                                }
                            }
                            "u" => {
                                if attr(attrs, "val").as_deref() != Some("none") {
                                    f.unheld = true;
                                }
                            }
                            "vertAlign" => {
                                if attr(attrs, "val").as_deref() != Some("baseline") {
                                    f.unheld = true;
                                }
                            }
                            "sz" => f.size = attr(attrs, "val").and_then(|v| v.trim().parse().ok()),
                            "color" => f.color = color_of(attrs),
                            "name" => f.name = attr(attrs, "val").map(Cow::into_owned),
                            _ => {}
                        }
                    }
                }
                ("fills", _) => list = if empty { List::None } else { List::Fills },
                ("fill", List::Fills) => {
                    if empty {
                        parsed.fills.push(FileFill::default());
                    } else {
                        fill = Some(FileFill::default());
                    }
                }
                ("patternFill", List::Fills) => {
                    if let Some(f) = fill.as_mut() {
                        match attr(attrs, "patternType").as_deref() {
                            Some("solid") => f.solid = true,
                            None | Some("none") => {}
                            Some(_) => f.unheld = true,
                        }
                    }
                }
                ("fgColor", List::Fills) if !in_gradient => {
                    if let Some(f) = fill.as_mut() {
                        f.fg = color_of(attrs);
                    }
                }
                ("gradientFill", List::Fills) => {
                    if let Some(f) = fill.as_mut() {
                        f.unheld = true;
                    }
                    in_gradient = !empty;
                }
                ("borders", _) => list = if empty { List::None } else { List::Borders },
                ("border", List::Borders) => {
                    if empty {
                        parsed.borders.push(FileBorder::default());
                    } else {
                        border = Some(FileBorder::default());
                    }
                }
                (
                    "left" | "start" | "right" | "end" | "top" | "bottom" | "diagonal",
                    List::Borders,
                ) => {
                    if let Some(b) = border.as_mut() {
                        let style = attr(attrs, "style").map(Cow::into_owned);
                        let at = match name {
                            "left" | "start" => Some(0),
                            "right" | "end" => Some(1),
                            "top" => Some(2),
                            "bottom" => Some(3),
                            _ => None,
                        };
                        match (at, style) {
                            (Some(i), Some(style)) if style != "none" => {
                                b.edges[i] = Some((style, None));
                                side = (!empty).then_some(i);
                            }
                            (None, Some(style)) if style != "none" => b.unheld = true,
                            _ => {}
                        }
                    }
                }
                ("color", List::Borders) => {
                    if let (Some(b), Some(i)) = (border.as_mut(), side) {
                        if let Some(edge) = b.edges[i].as_mut() {
                            edge.1 = color_of(attrs);
                        }
                    }
                }
                ("cellStyleXfs", _) => list = if empty { List::None } else { List::StyleXfs },
                ("xf", List::StyleXfs) => {
                    let font = attr(attrs, "fontId").and_then(|v| v.trim().parse().ok());
                    parsed.style_fonts.push(font.unwrap_or(0));
                    if !empty {
                        skip(&mut cursor)?;
                    }
                }
                ("cellXfs", _) => list = if empty { List::None } else { List::CellXfs },
                ("xf", List::CellXfs) => {
                    let id = |name: &str| attr(attrs, name).and_then(|v| v.trim().parse().ok());
                    let read = FileXf {
                        num_fmt: id("numFmtId").map_or(0, |v: usize| v as u32),
                        font: id("fontId").unwrap_or(0),
                        fill: id("fillId").unwrap_or(0),
                        border: id("borderId").unwrap_or(0),
                        ..FileXf::default()
                    };
                    if empty {
                        parsed.xfs.push(read);
                    } else {
                        xf = Some(read);
                    }
                }
                ("alignment", List::CellXfs) => {
                    if let Some(x) = xf.as_mut() {
                        x.horizontal = attr(attrs, "horizontal").map(Cow::into_owned);
                        x.vertical = attr(attrs, "vertical").map(Cow::into_owned);
                        x.wrap = truthy(attr(attrs, "wrapText"));
                        let nonzero = |name: &str| {
                            attr(attrs, name)
                                .and_then(|v| v.trim().parse::<f64>().ok())
                                .is_some_and(|v| v != 0.0)
                        };
                        if nonzero("indent")
                            || nonzero("textRotation")
                            || nonzero("relativeIndent")
                            || truthy(attr(attrs, "shrinkToFit"))
                            || truthy(attr(attrs, "justifyLastLine"))
                            || attr(attrs, "readingOrder").as_deref() == Some("2")
                        {
                            x.unheld = true;
                        }
                    }
                }
                ("cellStyles", _) => list = if empty { List::None } else { List::CellStyles },
                ("cellStyle", List::CellStyles) => {
                    if attr(attrs, "builtinId").as_deref() == Some("0") {
                        parsed.normal = attr(attrs, "xfId").and_then(|v| v.trim().parse().ok());
                    }
                }
                ("indexedColors", _) => {
                    list = if empty {
                        List::None
                    } else {
                        List::IndexedColors
                    }
                }
                ("rgbColor", List::IndexedColors) => {
                    let rgb = attr(attrs, "rgb").and_then(|v| rgb_from_hex(&v));
                    parsed.palette.push(rgb.unwrap_or(Rgb(0, 0, 0)));
                }
                _ => {}
            },
            Event::End { name } => match name {
                "font" if list == List::Fonts => {
                    if let Some(f) = font.take() {
                        parsed.fonts.push(f);
                    }
                }
                "fill" if list == List::Fills => {
                    if let Some(f) = fill.take() {
                        parsed.fills.push(f);
                    }
                    in_gradient = false;
                }
                "gradientFill" => in_gradient = false,
                "left" | "start" | "right" | "end" | "top" | "bottom" | "diagonal" => side = None,
                "border" if list == List::Borders => {
                    if let Some(b) = border.take() {
                        parsed.borders.push(b);
                    }
                }
                "xf" if list == List::CellXfs => {
                    if let Some(x) = xf.take() {
                        parsed.xfs.push(x);
                    }
                }
                "numFmts" | "fonts" | "fills" | "borders" | "cellStyleXfs" | "cellXfs"
                | "cellStyles" | "indexedColors" => list = List::None,
                _ => {}
            },
            Event::Text { .. } => {}
        }
    }
    Ok(parsed)
}

impl Parsed {
    /// The cell formats as the record says them, against the workbook's
    /// default font, the Normal style's (font 0 where the part names none),
    /// each color resolved through the theme and the palette.
    fn table(&self, theme: Option<[Option<Rgb>; 12]>) -> StyleTable {
        let colors = Colors::new(
            theme,
            (!self.palette.is_empty()).then(|| self.palette.clone()),
        );
        let default = self
            .normal
            .and_then(|x| self.style_fonts.get(x))
            .and_then(|f| self.fonts.get(*f))
            .or_else(|| self.fonts.first())
            .cloned()
            .unwrap_or_default();
        let mut looks: Vec<Look> = self
            .xfs
            .iter()
            .map(|xf| self.look_of(xf, &default, &colors))
            .collect();
        if looks.is_empty() {
            looks.push(Look::default());
        }
        StyleTable { looks, part: true }
    }

    /// One cell format as the record says it (decision 5): its own number
    /// format, fill, font, alignment and border, whatever its `apply` flags
    /// and its cell style say, as Excel draws it (Excel's reading: each flag
    /// off, all of them and none written each drawn as the entry says); a
    /// property the default font has unsaid; what the record has no word for
    /// counted, the rest said.
    fn look_of(&self, xf: &FileXf, default: &FileFont, colors: &Colors) -> Look {
        let mut look = Look::default();
        // The number format: a code the part defines, or a built-in's as
        // Excel reads it, a locale's id as its twin and an id with no code
        // General.
        match self.codes.get(&xf.num_fmt) {
            Some(code) => look.set_code(code),
            None => {
                let twin = super::looks::builtin_twin(xf.num_fmt);
                if let Some(code) = super::looks::builtin_code(twin) {
                    look.set_code(code);
                }
            }
        }
        // The fill: a solid pattern's foreground; Excel's automatic, or no
        // color at all, white, as Excel draws a solid fill of either.
        if let Some(fill) = self.fills.get(xf.fill) {
            look.unheld |= fill.unheld;
            if fill.solid {
                match fill.fg.map(|c| colors.resolve(&c)) {
                    Some(Resolved::Rgb(c)) => look.style.fill = Some(c),
                    Some(Resolved::Unresolved) => look.unheld = true,
                    Some(Resolved::Automatic) | None => look.style.fill = Some(Rgb(255, 255, 255)),
                }
            }
        }
        // The font, against the workbook's default.
        let font = self.fonts.get(xf.font).cloned().unwrap_or_default();
        if font.bold != default.bold {
            look.style.font.bold = Some(font.bold);
        }
        if font.italic != default.italic {
            look.style.font.italic = Some(font.italic);
        }
        if font.size != default.size {
            if let Some(points) = font.size {
                match font_size(points) {
                    Some(size) => look.style.font.size = Some(size),
                    None => look.unheld = true,
                }
            }
        }
        if font.color != default.color {
            match font.color.map(|c| colors.resolve(&c)) {
                Some(Resolved::Rgb(c)) => look.style.font.color = Some(c),
                Some(Resolved::Unresolved) => look.unheld = true,
                Some(Resolved::Automatic) | None => {}
            }
        }
        if font.name != default.name || font.unheld {
            look.unheld = true;
        }
        // The alignment and the wrap.
        match xf.horizontal.as_deref() {
            None | Some("general") => {}
            Some(word) => match Horizontal::from_word(word) {
                Some(h) => look.style.align.horizontal = Some(h),
                None => look.unheld = true,
            },
        }
        match xf.vertical.as_deref() {
            None | Some("bottom") => {}
            Some("top") => look.style.align.vertical = Some(Vertical::Top),
            Some("center") => look.style.align.vertical = Some(Vertical::Middle),
            Some(_) => look.unheld = true,
        }
        look.style.wrap = xf.wrap;
        look.unheld |= xf.unheld;
        // The border: each edge a line names, its color resolved.
        if let Some(border) = self.borders.get(xf.border) {
            look.unheld |= border.unheld;
            for (i, edge) in border.edges.iter().enumerate() {
                let Some((style, color)) = edge else {
                    continue;
                };
                let Some(line) = Line::from_excel(style) else {
                    look.unheld = true;
                    continue;
                };
                if line == Line::None {
                    continue;
                }
                let color = match color.map(|c| colors.resolve(&c)) {
                    Some(Resolved::Rgb(c)) => Some(c),
                    Some(Resolved::Unresolved) => {
                        look.unheld = true;
                        None
                    }
                    Some(Resolved::Automatic) | None => None,
                };
                let edge = Some(Edge::new(line, color));
                match i {
                    0 => look.style.border.left = edge,
                    1 => look.style.border.right = edge,
                    2 => look.style.border.top = edge,
                    _ => look.style.border.bottom = edge,
                }
            }
        }
        look
    }
}

/// A theme part's color scheme, its twelve slots in Excel's count
/// ([`THEME_SLOTS`]): each a color of its own, `srgbClr`, or the system's
/// by the color it was when the file was saved, `sysClr`'s `lastClr`.
fn read_theme(text: &str, label: &str, part: &str) -> Result<[Option<Rgb>; 12], Refusal> {
    let mut slots = [None; 12];
    let mut cursor = Cursor::new(text);
    let mut in_scheme = false;
    let mut current: Option<usize> = None;
    while let Some(event) = step(&mut cursor, label, part)? {
        match event {
            Event::Start {
                name: "clrScheme",
                empty: false,
                ..
            } => in_scheme = true,
            Event::End { name: "clrScheme" } => break,
            Event::Start {
                name: "srgbClr",
                attrs,
                ..
            } if current.is_some() => {
                if let Some(i) = current {
                    slots[i] = attr(attrs, "val").and_then(|v| rgb_from_hex(&v));
                }
            }
            Event::Start {
                name: "sysClr",
                attrs,
                ..
            } if current.is_some() => {
                if let Some(i) = current {
                    slots[i] = attr(attrs, "lastClr").and_then(|v| rgb_from_hex(&v)).or(
                        match attr(attrs, "val").as_deref() {
                            Some("windowText") => Some(Rgb(0, 0, 0)),
                            Some("window") => Some(Rgb(255, 255, 255)),
                            _ => None,
                        },
                    );
                }
            }
            Event::Start { name, .. } if in_scheme && current.is_none() => {
                current = THEME_SLOTS.iter().position(|s| *s == name);
            }
            Event::End { name } if current.is_some_and(|i| THEME_SLOTS[i] == name) => {
                current = None
            }
            _ => {}
        }
    }
    Ok(slots)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reflect::reflect_text;
    use crate::sheet::A1Range;

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
    fn an_array_formula_s_kind_goes_beside_its_row() {
        // KERNEL.8, slice b: an array formula's anchor told to the sink after
        // its formula row, dynamic where its `cm` names the dynamic-array
        // metadata, legacy where it has none or names other metadata, or
        // where the package holds no metadata part and its `cm` names
        // nothing to say otherwise; a plain formula told nothing; the
        // printer's record the same either way.
        #[derive(Default)]
        struct Kinds(Vec<(String, ArrayKind)>);
        impl Sink for Kinds {
            fn row(&mut self, _: &Row<'_>) {}
            fn array(&mut self, _sheet: &str, addr: &str, kind: ArrayKind) {
                self.0.push((addr.to_string(), kind));
            }
        }
        let data = concat!(
            "<row r=\"1\"><c r=\"A1\" cm=\"1\"><f t=\"array\" ref=\"A1:A2\">_xlfn.SEQUENCE(2)</f><v>1</v></c>",
            "<c r=\"B1\"><f t=\"array\" ref=\"B1\">SUM(C1:C2*D1:D2)</f><v>0</v></c>",
            "<c r=\"C1\"><f>D1+1</f><v>1</v></c>",
            "<c r=\"E1\" cm=\"2\"><f t=\"array\" ref=\"E1\">D1</f><v>0</v></c></row>",
        );
        let kinds = |bytes: &[u8]| {
            let p = Package::open(bytes, "t.xlsx").unwrap();
            let mut k = Kinds::default();
            p.walk_sheet(0, &mut k).unwrap();
            k.0
        };
        let without = one_sheet(data, None);
        let mut parts = zip::entries(&without)
            .unwrap()
            .iter()
            .map(|e| {
                (
                    e.name.clone(),
                    part_text_raw(&without, &zip::entries(&without).unwrap(), &e.name)
                        .unwrap()
                        .unwrap()
                        .into_bytes(),
                )
            })
            .collect::<Vec<_>>();
        parts.push((
            "xl/metadata.xml".to_string(),
            crate::sheet::ooxml::metadata_xml().into_bytes(),
        ));
        let with = zip::write_stored(&parts).unwrap();
        let dynamic = ArrayKind::Dynamic;
        let legacy = ArrayKind::Legacy;
        assert_eq!(
            kinds(&with),
            vec![
                ("A1".to_string(), dynamic),
                ("B1".to_string(), legacy),
                ("E1".to_string(), legacy)
            ],
            "the metadata names cm 1 alone"
        );
        assert_eq!(
            kinds(&without),
            vec![
                ("A1".to_string(), dynamic),
                ("B1".to_string(), legacy),
                ("E1".to_string(), dynamic)
            ],
            "no metadata part: a cm marks a dynamic formula"
        );
        assert_eq!(
            reflect_text(&with, "t.xlsx").unwrap(),
            reflect_text(&without, "t.xlsx").unwrap(),
            "the printer ignores the kind"
        );
    }

    #[test]
    fn an_error_a_rich_value_holds_reads_as_its_own() {
        // KERNEL.8, slice f: Excel writes `#SPILL!` and `#CALC!`, which have
        // no cell error value, as `#VALUE!` beside a `vm` naming a rich value
        // of the `_error` structure, its `errorType` 8 or 13; the reader reads
        // the error the rich value names, and a `vm` naming nothing it knows
        // leaves the cell's own value. The parts are Excel's own save of the
        // arrays' fixture, its keys in its order.
        const RICH: &str = "http://schemas.microsoft.com/office/2017/06/relationships";
        const META: &str =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata";
        let metadata = concat!(
            "<metadata xmlns:xlrd=\"http://schemas.microsoft.com/office/spreadsheetml/2017/richdata\">",
            "<metadataTypes count=\"2\"><metadataType name=\"XLDAPR\" minSupportedVersion=\"120000\"/>",
            "<metadataType name=\"XLRICHVALUE\" minSupportedVersion=\"120000\"/></metadataTypes>",
            "<futureMetadata name=\"XLDAPR\" count=\"1\"><bk><extLst><ext uri=\"{bdbb8cdc}\">",
            "<xda:dynamicArrayProperties fDynamic=\"1\" fCollapsed=\"0\"/></ext></extLst></bk></futureMetadata>",
            "<futureMetadata name=\"XLRICHVALUE\" count=\"2\"><bk><extLst><ext uri=\"{3e2802c4}\">",
            "<xlrd:rvb i=\"0\"/></ext></extLst></bk><bk><extLst><ext uri=\"{3e2802c4}\">",
            "<xlrd:rvb i=\"1\"/></ext></extLst></bk></futureMetadata>",
            "<cellMetadata count=\"1\"><bk><rc t=\"1\" v=\"0\"/></bk></cellMetadata>",
            "<valueMetadata count=\"2\"><bk><rc t=\"2\" v=\"0\"/></bk><bk><rc t=\"2\" v=\"1\"/></bk>",
            "</valueMetadata></metadata>",
        );
        let values = concat!(
            "<rvData count=\"2\"><rv s=\"0\"><v>0</v><v>8</v><v>5</v><v>1</v></rv>",
            "<rv s=\"1\"><v>13</v><v>3</v></rv></rvData>",
        );
        let structures = concat!(
            "<rvStructures count=\"2\"><s t=\"_error\"><k n=\"colOffset\" t=\"i\"/>",
            "<k n=\"errorType\" t=\"i\"/><k n=\"rwOffset\" t=\"i\"/><k n=\"subType\" t=\"i\"/></s>",
            "<s t=\"_error\"><k n=\"errorType\" t=\"i\"/><k n=\"subType\" t=\"i\"/></s></rvStructures>",
        );
        let data = concat!(
            "<row r=\"1\"><c r=\"A1\" t=\"e\" cm=\"1\" vm=\"1\"><f t=\"array\" ref=\"A1\">B1:B6*2</f><v>#VALUE!</v></c>",
            "<c r=\"B1\" t=\"e\" cm=\"1\" vm=\"2\"><f t=\"array\" ref=\"B1\">_xlfn.SEQUENCE(0)</f><v>#VALUE!</v></c>",
            "<c r=\"C1\" t=\"e\" vm=\"3\"><f>1/0</f><v>#DIV/0!</v></c>",
            "<c r=\"D1\" t=\"e\"><f>NA()</f><v>#N/A</v></c></row>",
        );
        let rels = format!(
            "<Relationships><Relationship Id=\"rId1\" Type=\"{REL_SHEET}\" Target=\"worksheets/sheet1.xml\"/><Relationship Id=\"rId2\" Type=\"{META}\" Target=\"metadata.xml\"/><Relationship Id=\"rId3\" Type=\"{RICH}/rdRichValue\" Target=\"richData/rdrichvalue.xml\"/><Relationship Id=\"rId4\" Type=\"{RICH}/rdRichValueStructure\" Target=\"richData/rdrichvaluestructure.xml\"/></Relationships>"
        );
        let parts: Vec<(String, Vec<u8>)> = vec![
            ("xl/workbook.xml".to_string(), format!("<workbook {NS}><sheets><sheet name=\"Model\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>").into_bytes()),
            ("xl/_rels/workbook.xml.rels".to_string(), rels.into_bytes()),
            ("xl/worksheets/sheet1.xml".to_string(), format!("<worksheet {NS}><sheetData>{data}</sheetData></worksheet>").into_bytes()),
            ("xl/metadata.xml".to_string(), metadata.as_bytes().to_vec()),
            ("xl/richData/rdrichvalue.xml".to_string(), values.as_bytes().to_vec()),
            ("xl/richData/rdrichvaluestructure.xml".to_string(), structures.as_bytes().to_vec()),
        ];
        let bytes = zip::write_stored(&parts).unwrap();
        #[derive(Default)]
        struct Values(Vec<(String, Value)>);
        impl Sink for Values {
            fn row(&mut self, row: &Row<'_>) {
                if let Row::Cell { addr, value, .. } = row {
                    self.0.push((addr.to_string(), (*value).clone()));
                }
            }
        }
        let p = Package::open(&bytes, "t.xlsx").unwrap();
        let mut got = Values::default();
        p.walk_sheet(0, &mut got).unwrap();
        let error = |a: &str, t: &str| (a.to_string(), Value::Error(t.to_string()));
        assert_eq!(
            got.0,
            vec![
                error("A1", "#SPILL!"),
                error("B1", "#CALC!"),
                error("C1", "#DIV/0!"),
                error("D1", "#N/A"),
            ]
        );
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
    fn a_workbook_of_1904_says_so_and_its_calendar_s_functions_are_labelled() {
        // KERNEL.8, slice d: `workbookPr`'s `date1904`, a Mac's workbook,
        // read; the door labels the functions whose answer turns on the
        // calendar there, and computes the rest, and a workbook without it
        // counts from 1900.
        let sheet = concat!(
            "<row r=\"1\"><c r=\"A1\"><f>YEAR(1)</f><v>1904</v></c>",
            "<c r=\"B1\"><f>DAYS(10,1)</f><v>9</v></c></row>",
        );
        let plain = one_sheet(sheet, None);
        let parts: Vec<(String, Vec<u8>)> = vec![
            ("xl/workbook.xml".to_string(), format!("<workbook {NS}><workbookPr date1904=\"1\"/><sheets><sheet name=\"Model\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>").into_bytes()),
            ("xl/_rels/workbook.xml.rels".to_string(), format!("<Relationships><Relationship Id=\"rId1\" Type=\"{REL_SHEET}\" Target=\"worksheets/sheet1.xml\"/></Relationships>").into_bytes()),
            ("xl/worksheets/sheet1.xml".to_string(), format!("<worksheet {NS}><sheetData>{sheet}</sheetData></worksheet>").into_bytes()),
        ];
        let mac = zip::write_stored(&parts).unwrap();
        assert!(Package::open(&mac, "m.xlsx").unwrap().date1904());
        assert!(!Package::open(&plain, "p.xlsx").unwrap().date1904());
        let library = crate::excel::library();
        let rows = crate::reflect::calc::calc_text(&mac, "m.xlsx", &library).unwrap();
        assert_eq!(
            rows,
            concat!(
                "(calc \"Model\" \"A1\" \"=YEAR(1)\" (not-computed \"the 1904 date system\") 1904 unchecked)\n",
                "(calc \"Model\" \"B1\" \"=DAYS(10,1)\" 9 9 agree)\n",
            )
        );
        let rows = crate::reflect::calc::calc_text(&plain, "p.xlsx", &library).unwrap();
        assert!(rows.starts_with("(calc \"Model\" \"A1\" \"=YEAR(1)\" 1900 1904 differ)\n"));
    }

    #[test]
    fn a_hidden_row_is_read_and_subtotal_s_101_to_111_pass_it_by() {
        // KERNEL.8, slice e: `<row hidden="1">` told to the sink, `true`
        // spelled as well, and a row hidden `0` not; the door's `SUBTOTAL`
        // 109 sums the rows shown, and its 9 over a range a hidden row
        // crosses is labelled, the file's one flag for a row a filter hides
        // and one a person hides, which 1 to 11 tell apart; over a range no
        // hidden row crosses, it computes. The printer's record is the same
        // with the flags as without them.
        #[derive(Default)]
        struct Hidden(Vec<(String, u32)>);
        impl Sink for Hidden {
            fn row(&mut self, _row: &Row<'_>) {}
            fn hidden_row(&mut self, sheet: &str, row: u32) {
                self.0.push((sheet.to_string(), row));
            }
        }
        let sheet = concat!(
            "<row r=\"1\"><c r=\"A1\"><v>1</v></c></row>",
            "<row r=\"2\" hidden=\"1\"><c r=\"A2\"><v>2</v></c></row>",
            "<row r=\"3\" hidden=\"0\"><c r=\"A3\"><v>4</v></c></row>",
            "<row r=\"4\" hidden=\"true\"><c r=\"A4\"><v>8</v></c></row>",
            "<row r=\"5\"><c r=\"B5\"><f>SUBTOTAL(109,A1:A4)</f><v>5</v></c>",
            "<c r=\"C5\"><f>SUBTOTAL(9,A1:A4)</f><v>15</v></c>",
            "<c r=\"D5\"><f>SUBTOTAL(9,A1,A3)</f><v>5</v></c></row>",
        );
        let book = one_sheet(sheet, None);
        let mut hidden = Hidden::default();
        Package::open(&book, "h.xlsx")
            .unwrap()
            .walk_sheet(0, &mut hidden)
            .unwrap();
        let model = "Model".to_string();
        assert_eq!(hidden.0, vec![(model.clone(), 2), (model, 4)]);
        let library = crate::excel::library();
        let computed = crate::reflect::calc::calc_text(&book, "h.xlsx", &library).unwrap();
        assert_eq!(
            computed,
            concat!(
                "(calc \"Model\" \"B5\" \"=SUBTOTAL(109,A1:A4)\" 5 5 agree)\n",
                "(calc \"Model\" \"C5\" \"=SUBTOTAL(9,A1:A4)\" (not-computed \"a row a filter may hide\") 15 unchecked)\n",
                "(calc \"Model\" \"D5\" \"=SUBTOTAL(9,A1,A3)\" 5 5 agree)\n",
            )
        );
        let plain = [" hidden=\"1\"", " hidden=\"0\"", " hidden=\"true\""]
            .iter()
            .fold(sheet.to_string(), |s, flag| s.replace(flag, ""));
        let shown = one_sheet(&plain, None);
        assert_eq!(
            reflect_text(&book, "h.xlsx").unwrap(),
            reflect_text(&shown, "h.xlsx").unwrap()
        );
    }

    #[test]
    fn a_formula_s_empty_text_reads_as_one() {
        // KERNEL.8, slice c: a host saves `=""` as `t="str"` and an empty
        // `<v>`, which is the empty text; with no `<v>`, or a number's empty
        // `<v>`, the cell holds nothing.
        let data = concat!(
            "<row r=\"1\"><c r=\"A1\" t=\"str\"><f>\"\"&amp;\"\"</f><v></v></c>",
            "<c r=\"B1\" t=\"str\"><f>\"\"&amp;\"\"</f></c>",
            "<c r=\"C1\"><f>1+1</f><v></v></c></row>",
        );
        let text = reflect_text(&one_sheet(data, None), "t.xlsx").unwrap();
        assert_eq!(
            text,
            concat!(
                "(sheet \"Model\" visible)\n",
                "(cell \"Model\" \"A1\" \"\")\n",
                "(formula \"Model\" \"A1\" \"=\\\"\\\"&\\\"\\\"\")\n",
                "(formula \"Model\" \"B1\" \"=\\\"\\\"&\\\"\\\"\")\n",
                "(formula \"Model\" \"C1\" \"=1+1\")\n",
            )
        );
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
        // An OpenDocument file opens through its own reader (8e); given to
        // this one directly, it is a package with no workbook part.
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
            .contains("no xl/workbook.xml"));
        assert!(crate::reflect::odf::is_opendocument(&ods));
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
        let rows: [(&str, &[u8], &str); 8] = [
            // The twin as Excel 365 saved it (the owner's pass, 2026-10-05):
            // Excel's own ODF dialect, the filler rows repeated a million
            // times, a sheet renamed, the link dropped, `@` as SINGLE.
            (
                "scripts/reflect/opendocument_saved.ods",
                include_bytes!("../../../scripts/reflect/opendocument_saved.ods"),
                include_str!("../../../scripts/reflect/opendocument_saved_relations.vla"),
            ),
            // The fixture's changed copy (8c) and its OpenDocument twin (8e).
            (
                "scripts/reflect/changed.xlsx",
                include_bytes!("../../../scripts/reflect/changed.xlsx"),
                include_str!("../../../scripts/reflect/changed_relations.vla"),
            ),
            (
                "scripts/reflect/opendocument.ods",
                include_bytes!("../../../scripts/reflect/opendocument.ods"),
                include_str!("../../../scripts/reflect/opendocument_relations.vla"),
            ),
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
    // ---- the looks (KERNEL.57) ----------------------------------------

    const REL_STYLES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
    const REL_THEME: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";

    /// Office's theme of 2023, the one Excel 365's saves carry.
    const THEME: &str = concat!(
        "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Office Theme\"><a:themeElements><a:clrScheme name=\"Office\">",
        "<a:dk1><a:sysClr val=\"windowText\" lastClr=\"000000\"/></a:dk1><a:lt1><a:sysClr val=\"window\" lastClr=\"FFFFFF\"/></a:lt1>",
        "<a:dk2><a:srgbClr val=\"0E2841\"/></a:dk2><a:lt2><a:srgbClr val=\"E8E8E8\"/></a:lt2>",
        "<a:accent1><a:srgbClr val=\"156082\"/></a:accent1><a:accent2><a:srgbClr val=\"E97132\"/></a:accent2>",
        "<a:accent3><a:srgbClr val=\"196B24\"/></a:accent3><a:accent4><a:srgbClr val=\"0F9ED5\"/></a:accent4>",
        "<a:accent5><a:srgbClr val=\"A02B93\"/></a:accent5><a:accent6><a:srgbClr val=\"4EA72E\"/></a:accent6>",
        "<a:hlink><a:srgbClr val=\"467886\"/></a:hlink><a:folHlink><a:srgbClr val=\"96607D\"/></a:folHlink>",
        "</a:clrScheme></a:themeElements></a:theme>",
    );

    /// A package of one sheet with a styles part and, when given, a theme.
    fn styled(sheet: &str, styles: &str, theme: Option<&str>, sst: &str) -> Vec<u8> {
        let mut rels = format!("<Relationships><Relationship Id=\"rId1\" Type=\"{REL_SHEET}\" Target=\"worksheets/sheet1.xml\"/><Relationship Id=\"rId2\" Type=\"{REL_SST}\" Target=\"sharedStrings.xml\"/><Relationship Id=\"rId3\" Type=\"{REL_STYLES}\" Target=\"styles.xml\"/>");
        let mut parts: Vec<(String, Vec<u8>)> = vec![
            ("xl/workbook.xml".to_string(), format!("<workbook {NS}><sheets><sheet name=\"Model\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>").into_bytes()),
            ("xl/worksheets/sheet1.xml".to_string(), format!("<worksheet {NS}>{sheet}</worksheet>").into_bytes()),
            ("xl/sharedStrings.xml".to_string(), format!("<sst {NS}>{sst}</sst>").into_bytes()),
            ("xl/styles.xml".to_string(), format!("<styleSheet {NS}>{styles}</styleSheet>").into_bytes()),
        ];
        if let Some(t) = theme {
            rels.push_str(&format!(
                "<Relationship Id=\"rId4\" Type=\"{REL_THEME}\" Target=\"theme/theme1.xml\"/>"
            ));
            parts.push(("xl/theme/theme1.xml".to_string(), t.as_bytes().to_vec()));
        }
        rels.push_str("</Relationships>");
        parts.push(("xl/_rels/workbook.xml.rels".to_string(), rels.into_bytes()));
        zip::write_stored(&parts).unwrap()
    }

    const STYLES: &str = concat!(
        "<numFmts count=\"2\"><numFmt numFmtId=\"164\" formatCode=\"0.0%\"/><numFmt numFmtId=\"165\" formatCode=\"[Red]0.00\"/></numFmts>",
        "<fonts count=\"6\">",
        "<font><sz val=\"11\"/><color theme=\"1\"/><name val=\"Aptos Narrow\"/><family val=\"2\"/><scheme val=\"minor\"/></font>",
        "<font><b/><sz val=\"11\"/><color theme=\"1\"/><name val=\"Aptos Narrow\"/></font>",
        "<font><i/><sz val=\"22\"/><color rgb=\"FFC00000\"/><name val=\"Aptos Narrow\"/></font>",
        "<font><u/><sz val=\"11\"/><color theme=\"5\" tint=\"0.39997558519241921\"/><name val=\"Courier New\"/></font>",
        "<font><b val=\"0\"/><sz val=\"11\"/><color theme=\"1\"/><name val=\"Aptos Narrow\"/></font>",
        "<font><sz val=\"11\"/><color theme=\"0\"/><name val=\"Aptos Narrow\"/></font>",
        "</fonts>",
        "<fills count=\"6\">",
        "<fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill>",
        "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFF7F4FC\"/><bgColor indexed=\"64\"/></patternFill></fill>",
        "<fill><patternFill patternType=\"solid\"><fgColor theme=\"4\" tint=\"0.79998168889431442\"/><bgColor indexed=\"64\"/></patternFill></fill>",
        "<fill><patternFill patternType=\"solid\"><fgColor indexed=\"22\"/><bgColor indexed=\"64\"/></patternFill></fill>",
        "<fill><patternFill patternType=\"darkGray\"><fgColor rgb=\"FF000000\"/></patternFill></fill>",
        "</fills>",
        "<borders count=\"3\">",
        "<border><left/><right/><top/><bottom/><diagonal/></border>",
        "<border><left style=\"thin\"><color indexed=\"64\"/></left><right/><top style=\"medium\"><color rgb=\"FF6B4FC8\"/></top><bottom style=\"double\"><color auto=\"1\"/></bottom><diagonal/></border>",
        "<border diagonalUp=\"1\"><left/><right/><top/><bottom/><diagonal style=\"thin\"><color indexed=\"64\"/></diagonal></border>",
        "</borders>",
        "<cellStyleXfs count=\"2\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/><xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>",
        "<cellXfs count=\"12\">",
        "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>",
        "<xf numFmtId=\"2\" fontId=\"1\" fillId=\"2\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"1\"/>",
        "<xf numFmtId=\"0\" fontId=\"2\" fillId=\"0\" borderId=\"1\" xfId=\"0\" applyAlignment=\"1\"><alignment horizontal=\"center\" vertical=\"center\" wrapText=\"1\"/></xf>",
        "<xf numFmtId=\"164\" fontId=\"0\" fillId=\"3\" borderId=\"0\" xfId=\"0\"/>",
        "<xf numFmtId=\"165\" fontId=\"0\" fillId=\"4\" borderId=\"0\" xfId=\"0\"/>",
        "<xf numFmtId=\"44\" fontId=\"3\" fillId=\"5\" borderId=\"2\" xfId=\"0\"><alignment horizontal=\"centerContinuous\" indent=\"1\"/></xf>",
        "<xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"0\"/>",
        "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"1\" applyFont=\"0\"/>",
        "<xf numFmtId=\"14\" fontId=\"4\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>",
        "<xf numFmtId=\"49\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"><alignment vertical=\"top\" horizontal=\"right\"/></xf>",
        "<xf numFmtId=\"0\" fontId=\"1\" fillId=\"2\" borderId=\"0\" xfId=\"0\"/>",
        "<xf numFmtId=\"0\" fontId=\"5\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>",
        "</cellXfs>",
        "<cellStyles count=\"2\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/><cellStyle name=\"Heading\" xfId=\"1\"/></cellStyles>",
        "<dxfs count=\"1\"><dxf><font><b/></font><fill><patternFill><bgColor rgb=\"FFFF0000\"/></patternFill></fill></dxf></dxfs>",
    );

    const SHEET: &str = concat!(
        "<dimension ref=\"A1:F7\"/><sheetViews><sheetView showGridLines=\"0\" workbookViewId=\"0\"/></sheetViews>",
        "<sheetFormatPr defaultRowHeight=\"14.5\"/>",
        "<cols><col min=\"1\" max=\"1\" width=\"20.7109375\" customWidth=\"1\" style=\"1\"/><col min=\"2\" max=\"3\" width=\"9.140625\" style=\"9\"/>",
        "<col min=\"4\" max=\"4\" width=\"0\" hidden=\"1\" customWidth=\"1\"/><col min=\"6\" max=\"16384\" width=\"30.7109375\" customWidth=\"1\"/></cols>",
        "<sheetData>",
        "<row r=\"1\" ht=\"30\" customHeight=\"1\"><c r=\"A1\" s=\"1\" t=\"s\"><v>0</v></c><c r=\"B1\" s=\"2\"><v>1.5</v></c><c r=\"C1\" s=\"3\"><v>0.25</v></c></row>",
        "<row r=\"2\" ht=\"28.5\"><c r=\"A2\" s=\"4\"><v>-2</v></c><c r=\"B2\" s=\"5\"><v>3</v></c><c r=\"C2\" s=\"6\" t=\"s\"><v>1</v></c></row>",
        "<row r=\"3\" hidden=\"1\"><c r=\"A3\" s=\"7\"><f>B1*2</f><v>3</v></c><c r=\"B3\" s=\"8\" t=\"d\"><v>2025-09-30</v></c></row>",
        "<row r=\"4\" s=\"2\" customFormat=\"1\"><c r=\"E4\" s=\"9\" t=\"inlineStr\"><is><r><rPr><b/></rPr><t>bold</t></r><r><t xml:space=\"preserve\"> plain</t></r></is></c></row>",
        "<row r=\"5\" ht=\"12\" customHeight=\"1\"/>",
        "<row r=\"6\"><c r=\"D6\" s=\"11\"><v>8</v></c><c r=\"E6\"><v>9</v></c></row>",
        "<row r=\"7\"><c r=\"F7\" s=\"10\"/></row>",
        "</sheetData>",
        "<mergeCells count=\"1\"><mergeCell ref=\"A5:B5\"/></mergeCells>",
        "<conditionalFormatting sqref=\"A1:A10\"><cfRule type=\"cellIs\" dxfId=\"0\" priority=\"1\" operator=\"equal\"><formula>1</formula></cfRule></conditionalFormatting>",
        "<extLst><ext uri=\"{78C0D931-6437-407d-A8EE-F0AAD7539E65}\"><x14:conditionalFormattings xmlns:x14=\"x\"><x14:conditionalFormatting><x14:cfRule type=\"dataBar\"/></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>",
    );

    const SST: &str = "<si><t>Title</t></si><si><r><rPr><b/></rPr><t>rich</t></r></si>";

    #[test]
    fn a_workbook_s_looks_print_as_the_record_prints_a_sheet_whole() {
        // Decisions 2 to 10 over one sheet: its gridlines; the columns of
        // its extent, a width read as Excel reads it, Excel's default none,
        // a span one row a column, a hidden one and one past the extent
        // left out; the rows a person set or hid, a fitted height none;
        // format 0 and every format the printed cells and columns use, a
        // format only a formatted empty cell uses left out; each format at
        // its own index as the record says it, against the Normal style's
        // font, its colors resolved through the theme, a tint and the
        // palette, its codes Excel's built-ins and the part's own, a code
        // the language does not draw as written, an `apply` flag off and a
        // cell style passed over, what the record has no word for left out;
        // and each cell's rows, a saved value a `value` row, a date its
        // serial, its `style` row after them and its `refers` rows last.
        let bytes = styled(SHEET, STYLES, Some(THEME), SST);
        let (text, stats) = crate::reflect::reflect_looks(&bytes, "looks.xlsx").unwrap();
        // A tint as the file writes it, read as the reader reads it.
        let tint = |text: &str| text.parse::<f64>().unwrap();
        let fill3 = super::super::looks::tinted(Rgb(0x15, 0x60, 0x82), tint("0.79998168889431442"));
        let ink5 = super::super::looks::tinted(Rgb(0xE9, 0x71, 0x32), tint("0.39997558519241921"));
        let accounting = vla_lang::rows::quoted(super::super::looks::builtin_code(44).unwrap());
        let want = [
            "(sheet \"Model\" visible)".to_string(),
            "(gridlines \"Model\" off)".to_string(),
            "(column \"Model\" \"A\" 20 shown 1)".to_string(),
            "(column \"Model\" \"B\" none shown 9)".to_string(),
            "(column \"Model\" \"C\" none shown 9)".to_string(),
            "(column \"Model\" \"D\" none hidden none)".to_string(),
            "(row \"Model\" 1 30 shown)".to_string(),
            "(row \"Model\" 3 none hidden)".to_string(),
            "(row \"Model\" 5 12 shown)".to_string(),
            "(format 0 none general nowrap)".to_string(),
            "(format 1 \"F7F4FC\" \"0.00\" nowrap (font bold))".to_string(),
            "(format 2 none general wrap (font italic 22 \"C00000\") (align center middle) (border (left thin) (top medium \"6B4FC8\") (bottom double)))".to_string(),
            format!("(format 3 \"{}\" \"0.0%\" nowrap)", fill3.hex()),
            "(format 4 \"C0C0C0\" \"[Red]0.00\" nowrap)".to_string(),
            format!("(format 5 none {accounting} nowrap (font \"{}\"))", ink5.hex()),
            "(format 6 none general nowrap (font bold))".to_string(),
            "(format 7 none general nowrap)".to_string(),
            "(format 8 none \"m/d/yyyy\" nowrap)".to_string(),
            "(format 9 none text nowrap (align right top))".to_string(),
            "(format 11 none general nowrap (font \"FFFFFF\"))".to_string(),
            "(cell \"Model\" \"A1\" \"Title\")".to_string(),
            "(style \"Model\" \"A1\" 1)".to_string(),
            "(cell \"Model\" \"B1\" 1.5)".to_string(),
            "(style \"Model\" \"B1\" 2)".to_string(),
            "(cell \"Model\" \"C1\" 0.25)".to_string(),
            "(style \"Model\" \"C1\" 3)".to_string(),
            "(cell \"Model\" \"A2\" -2)".to_string(),
            "(style \"Model\" \"A2\" 4)".to_string(),
            "(cell \"Model\" \"B2\" 3)".to_string(),
            "(style \"Model\" \"B2\" 5)".to_string(),
            "(cell \"Model\" \"C2\" \"rich\")".to_string(),
            "(style \"Model\" \"C2\" 6)".to_string(),
            "(formula \"Model\" \"A3\" \"=B1*2\")".to_string(),
            "(value \"Model\" \"A3\" 3)".to_string(),
            "(style \"Model\" \"A3\" 7)".to_string(),
            "(refers \"Model!A3\" \"Model!B1\")".to_string(),
            "(cell \"Model\" \"B3\" 45930)".to_string(),
            "(style \"Model\" \"B3\" 8)".to_string(),
            "(cell \"Model\" \"E4\" \"bold plain\")".to_string(),
            "(style \"Model\" \"E4\" 9)".to_string(),
            "(cell \"Model\" \"D6\" 8)".to_string(),
            "(style \"Model\" \"D6\" 11)".to_string(),
            "(cell \"Model\" \"E6\" 9)".to_string(),
        ];
        let got: Vec<&str> = text.lines().collect();
        assert_eq!(got, want.iter().map(String::as_str).collect::<Vec<_>>());
        assert_eq!(
            stats,
            LooksStats {
                formats: 11,
                styled: 10,
                formatted_empty: 1,
                unheld: 3,
                undrawn: 2,
                merged: 1,
                conditional: 1,
                row_formats: 1,
            }
        );
        // A tint is a hue kept, lighter: the scoping's prediction until
        // Excel's reading holds it (the item's second slice).
        assert!(fill3.0 > 0x15 && fill3.1 > 0x60 && fill3.2 > 0x82);
        // The relations of the same file stand as they were: a saved value
        // a `cell` row before its formula, a date a date, no look.
        let relations = reflect_text(&bytes, "looks.xlsx").unwrap();
        assert!(
            relations.contains("(cell \"Model\" \"A3\" 3)\n(formula \"Model\" \"A3\" \"=B1*2\")\n")
        );
        assert!(relations.contains("(cell \"Model\" \"B3\" (date \"2025-09-30\"))\n"));
        for head in [
            "(gridlines ",
            "(column ",
            "(row ",
            "(format ",
            "(value ",
            "(style ",
        ] {
            assert!(!relations.contains(head), "{head} in the relations");
        }
    }

    #[test]
    fn a_workbook_with_no_theme_or_no_styles_part_reads_what_it_can() {
        // A theme's color with no theme part: unsaid and counted.
        let bytes = styled(SHEET, STYLES, None, SST);
        let (text, stats) = crate::reflect::reflect_looks(&bytes, "looks.xlsx").unwrap();
        assert!(text.contains("(format 3 none \"0.0%\" nowrap)\n"), "{text}");
        assert!(text.contains("(format 5 none "), "{text}");
        assert_eq!(
            stats.unheld, 5,
            "the theme's fill and the theme's white ink counted too"
        );
        // A package with no styles part: every cell at format 0.
        let plain = one_sheet("<row r=\"1\"><c r=\"A1\" s=\"5\"><v>1</v></c></row>", None);
        let (text, stats) = crate::reflect::reflect_looks(&plain, "plain.xlsx").unwrap();
        assert_eq!(
            text,
            "(sheet \"Model\" visible)\n(gridlines \"Model\" on)\n(format 0 none general nowrap)\n(cell \"Model\" \"A1\" 1)\n"
        );
        assert_eq!(stats.styled, 0);
    }

    #[test]
    fn a_workbook_s_own_palette_and_its_normal_style_s_font_are_its_own() {
        // A palette of the workbook's own (`indexedColors`) resolves its
        // indices; the default font is the Normal style's, here font 1,
        // Arial 10, so that format 0 says nothing and a cell in Calibri 11
        // says its size and counts its family.
        let styles = concat!(
            "<fonts count=\"2\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font><font><sz val=\"10\"/><name val=\"Arial\"/></font></fonts>",
            "<fills count=\"3\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill>",
            "<fill><patternFill patternType=\"solid\"><fgColor indexed=\"1\"/></patternFill></fill></fills>",
            "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>",
            "<cellXfs count=\"3\"><xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>",
            "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>",
            "<xf numFmtId=\"0\" fontId=\"1\" fillId=\"2\" borderId=\"0\" xfId=\"0\"/></cellXfs>",
            "<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>",
            "<colors><indexedColors><rgbColor rgb=\"FF000000\"/><rgbColor rgb=\"FF6B4FC8\"/></indexedColors></colors>",
        );
        let bytes = styled(
            "<sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c><c r=\"B1\" s=\"1\"><v>2</v></c><c r=\"C1\" s=\"2\"><v>3</v></c></row></sheetData>",
            styles,
            None,
            "",
        );
        let (text, stats) = crate::reflect::reflect_looks(&bytes, "own.xlsx").unwrap();
        assert!(
            text.contains(
                "(format 0 none general nowrap)
(format 1 none general nowrap (font 11))
(format 2 \"6B4FC8\" general nowrap)
"
            ),
            "{text}"
        );
        assert_eq!(stats.unheld, 1, "Calibri beside the default's Arial");
    }

    #[test]
    fn a_cell_naming_a_format_past_the_styles_part_is_refused_by_name() {
        let styles = "<cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellXfs>";
        let bytes = styled(
            "<sheetData><row r=\"1\"><c r=\"B1\" s=\"3\"><v>1</v></c></row></sheetData>",
            styles,
            None,
            "",
        );
        let refusal = crate::reflect::reflect_looks(&bytes, "past.xlsx").unwrap_err();
        assert_eq!(refusal.id, "reflect-unsupported");
        assert!(
            refusal
                .text
                .contains("the cell B1 names the cell format 3, past the 1 its styles part holds"),
            "{}",
            refusal.text
        );
        // The relations read the same file, the format unread.
        assert!(reflect_text(&bytes, "past.xlsx").is_ok());
    }

    /// The looks' first pass and the walk print the same cells: every
    /// package of the tree read under the looks prints the cells and
    /// formulas its relations print, each extent and window as the cells'.
    #[test]
    fn the_first_pass_reads_the_cells_the_walk_prints() {
        use std::path::Path;
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
        let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("..");
        let mut packages = 0;
        for dir in [
            "scripts/reflect",
            "scripts/build",
            "scripts",
            "scripts/recalc",
        ] {
            let Ok(entries) = std::fs::read_dir(root.join(dir)) else {
                continue;
            };
            for entry in entries.flatten() {
                let path = entry.path();
                if path.extension().and_then(|e| e.to_str()) != Some("xlsx") {
                    continue;
                }
                let bytes = std::fs::read(&path).unwrap();
                let label = path.display().to_string();
                let package = Package::open(&bytes, &label).unwrap();
                let table = package.style_table().unwrap();
                for i in 0..package.sheets().len() {
                    let mut related = Places::default();
                    package.walk_sheet(i, &mut related).unwrap();
                    let mut looked = Places::default();
                    package
                        .walk_sheet_looks(i, &mut looked, &mut HashSet::new())
                        .unwrap();
                    assert_eq!(related.0, looked.0, "{label}, sheet {i}");
                    // The first pass's extent is the printed cells' rectangle.
                    let part = package.sheets[i].part.clone();
                    let text = part_text_raw(&bytes, &package.entries, &part)
                        .unwrap()
                        .unwrap();
                    let scan = package.first_pass(&text, table, &label, &part).unwrap();
                    let places: Vec<(u32, u32)> = related
                        .0
                        .iter()
                        .map(|p| {
                            let a = parse_a1_range(p.rsplit('!').next().unwrap()).unwrap();
                            (a.top, a.left)
                        })
                        .collect();
                    let rectangle = (!places.is_empty()).then(|| A1Range {
                        top: places.iter().map(|p| p.0).min().unwrap(),
                        left: places.iter().map(|p| p.1).min().unwrap(),
                        bottom: places.iter().map(|p| p.0).max().unwrap(),
                        right: places.iter().map(|p| p.1).max().unwrap(),
                    });
                    assert_eq!(scan.extent, rectangle, "{label}, sheet {i}");
                }
                packages += 1;
            }
        }
        assert!(packages >= 15, "{packages} packages read");
    }
}
