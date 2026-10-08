//! The reader (PORT.8): a workbook's file read as the relations of
//! `REFLECT` (docs/SINGULARITY.md, Stage 0.1; the roadmap's `AXM.8`), with
//! no host on the machine.
//!
//! New ground, said plainly: the reference reads a workbook through Excel's
//! object model and never opens the file, so there is no VBA arm to port
//! for the file formats; what the relations are, and how a formula's
//! references are read (`AXM.7`, slice 8b), is the language's and follows
//! the VBA as everything else does. A reader prints seven of the eight:
//! `sheet(name, state)`, `name(name, refersto)`, `table(name, sheet,
//! range)`, `cell(sheet, addr, value)`, `formula(sheet, addr, text)` and,
//! from 8b, `refers(from, to)`; `changed` is `diff`'s (8c, `diff.rs`: two
//! files read through this reader and compared) and `ran` comes from no
//! file. `audit.rs` (8d) asks Stage 2.3's audit list of the relations:
//! six named walks, one finding a row.
//!
//! Two formats, one reader (8e): an OOXML package (`ooxml.rs`) and an
//! OpenDocument spreadsheet (`odf.rs`) each open into a [`Source`], the one
//! shape the door, `diff`, `audit` and `--cone` are written against, and
//! print the same rows in the same spelling, a formula from either file as
//! the formula bar shows it. [`open`] tells them apart by the file.
//!
//! The reader streams. It never builds a model of the workbook: each sheet
//! part is inflated whole (the bound every part already has, 256 MB), walked
//! once by the cursor in `cursor.rs`, and dropped before the next is read;
//! the walk hands each relation row to a [`Sink`] as it finds it. What is
//! held is the shared-string table, since any cell may name any string, and
//! the sink's own state: nothing for printing, counters for `--counts`. An
//! OpenDocument file holds every sheet in one part, so that part is held
//! whole and each sheet walked from its own slice of it.
//!
//! The fixed order, which is what makes the printed text a golden: every
//! `sheet` row in tab order; every `name` row sorted by its name without
//! case; every `table` row by its sheet's tab order, then its name; then
//! sheet by sheet in tab order, cell by cell in document order, each cell's
//! `cell` row, then its `formula` row, then its `refers` rows, one per
//! distinct target, sorted by the target as printed (slice 8b, through
//! `crate::refers`, the port of `VLA_Refers.bas`). `print.rs` has the
//! spelling; `cone.rs` sizes a cell's cone through an index of the walk.

pub mod audit;
pub mod cone;
pub mod cursor;
pub mod diff;
pub mod odf;
pub mod ooxml;
pub mod print;

use std::collections::HashSet;

use crate::intrinsics::fold;
use crate::messages::Refusal;
use crate::refers::{self, Kind};

pub use vla_lang::rows::{sheet_prefix, RefersTo, Row, Sink, Value, Visibility};

/// The cap on the sets of distinct R1C1 formulas `--counts` keeps, a
/// sheet's and the workbook's.
pub const DISTINCT_CAP: usize = 1 << 20;

/// A sink that keeps nothing, for `--counts`.
pub struct Discard;

impl Sink for Discard {
    fn row(&mut self, _row: &Row<'_>) {}
}

/// What one sheet's walk counted: cells with a value or a formula, formula
/// cells, array-formula anchors, the last row and column that hold one,
/// the part's inflated size, and from 8b the distinct formulas in R1C1 and
/// the `INDIRECT` and `OFFSET` calls met.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct SheetStats {
    pub cells: u64,
    pub formulas: u64,
    pub array_anchors: u64,
    pub rows: u32,
    pub columns: u32,
    pub part_bytes: usize,
    /// Distinct formulas on the sheet, compared in R1C1 relative to their
    /// cells, so that a formula filled down counts once.
    pub distinct_r1c1: u64,
    /// `INDIRECT` and `OFFSET` calls on the sheet: where a cone is blind.
    pub unreadable: u64,
}

/// One sheet of the workbook, in tab order.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SheetInfo {
    pub name: String,
    pub visibility: Visibility,
    /// The part that holds it: `xl/worksheets/sheet1.xml` in a package,
    /// `content.xml` in an OpenDocument file.
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

/// One Table (a ListObject), by the name Excel shows; in an OpenDocument
/// file a database range, Calc's named rectangle with a header row.
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

/// A workbook opened far enough to walk, in either format: what the door,
/// `diff`, `audit` and `--cone` ask of a file. The sheet parts are read one
/// at a time by [`Source::walk_sheet`] and dropped.
pub trait Source: std::fmt::Debug {
    /// What a refusal calls the file.
    fn label(&self) -> &str;
    fn sheets(&self) -> &[SheetInfo];
    fn names(&self) -> &[NameInfo];
    fn tables(&self) -> &[TableInfo];
    /// The other workbooks the file links to, by file name.
    fn external_books(&self) -> &[String];
    /// The counts `--counts` prints for the workbook.
    fn summary(&self) -> Summary;

    /// One sheet's cells in document order, to the sink: a `cell` row for a
    /// value the file holds, a `formula` row for a formula, then its
    /// `refers` rows; with the workbook-wide set of distinct R1C1 formulas
    /// carried across sheets for `--counts`. What the walk counted comes
    /// back.
    fn walk_sheet_with(
        &self,
        index: usize,
        sink: &mut dyn Sink,
        book_distinct: &mut HashSet<String>,
    ) -> Result<SheetStats, Refusal>;

    /// [`Source::walk_sheet_with`] with no workbook-wide set.
    fn walk_sheet(&self, index: usize, sink: &mut dyn Sink) -> Result<SheetStats, Refusal> {
        self.walk_sheet_with(index, sink, &mut HashSet::new())
    }

    /// A `name` row's first field: the name, a sheet-scoped one behind its
    /// sheet as a reference quotes it (`Model!Local`). The key `diff` and
    /// `audit` match names by.
    fn printed_name(&self, name: &NameInfo) -> String {
        match name.sheet.and_then(|i| self.sheets().get(i)) {
            Some(sheet) => format!("{}!{}", sheet_prefix(&sheet.name), name.name),
            None => name.name.clone(),
        }
    }

    /// The rows that come before any cell, in the fixed order: every
    /// `sheet` row in tab order, every `name` row sorted by its name without
    /// case, every `table` row by its sheet's tab order then its name.
    fn header_rows(&self, sink: &mut dyn Sink) {
        for sheet in self.sheets() {
            sink.row(&Row::Sheet {
                name: &sheet.name,
                visibility: sheet.visibility,
            });
        }
        let mut names: Vec<(String, &str)> = self
            .names()
            .iter()
            .filter(|n| !n.placeholder)
            .map(|n| (self.printed_name(n), n.refers_to.as_str()))
            .collect();
        names.sort_by(|a, b| fold(&a.0).cmp(&fold(&b.0)).then_with(|| a.0.cmp(&b.0)));
        for (name, refers_to) in &names {
            sink.row(&Row::Name { name, refers_to });
        }
        let mut tables: Vec<&TableInfo> = self.tables().iter().collect();
        tables.sort_by(|a, b| {
            a.sheet
                .cmp(&b.sheet)
                .then_with(|| fold(&a.name).cmp(&fold(&b.name)))
        });
        let sheets = self.sheets();
        for table in tables {
            let Some(sheet) = sheets.get(table.sheet) else {
                continue;
            };
            sink.row(&Row::Table {
                name: &table.name,
                sheet: &sheet.name,
                range: &table.range,
            });
        }
    }
}

/// The file opened by its format: an OpenDocument spreadsheet by its
/// `mimetype` entry (or a `content.xml` with no workbook part), an OOXML
/// package otherwise. `label` is what a refusal calls the file.
pub fn open<'a>(bytes: &'a [u8], label: &str) -> Result<Box<dyn Source + 'a>, Refusal> {
    if odf::is_opendocument(bytes) {
        Ok(Box::new(odf::Document::open(bytes, label)?))
    } else {
        Ok(Box::new(ooxml::Package::open(bytes, label)?))
    }
}

/// A formula's rows and counts, the same from either reader: the `formula`
/// row, then its `refers` rows, one per distinct target sorted as printed;
/// the unreadable calls counted; the R1C1 rendering added to the sheet's
/// and the workbook's distinct sets, each capped. `text` begins with `=`.
#[allow(clippy::too_many_arguments)]
pub(crate) fn emit_formula(
    sink: &mut dyn Sink,
    sheet: &str,
    addr: &str,
    text: &str,
    row: u32,
    col: u32,
    stats: &mut SheetStats,
    sheet_distinct: &mut HashSet<String>,
    book_distinct: &mut HashSet<String>,
) {
    stats.formulas += 1;
    sink.row(&Row::Formula { sheet, addr, text });
    let found = refers::scan(text);
    stats.unreadable += found.iter().filter(|r| r.kind == Kind::Unreadable).count() as u64;
    let mut targets: Vec<RefersTo> = found.iter().map(|r| RefersTo::of(r, sheet)).collect();
    targets.sort_by_key(print::target);
    targets.dedup();
    for to in &targets {
        sink.row(&Row::Refers { sheet, addr, to });
    }
    let rendered = refers::r1c1(text, row, col);
    if sheet_distinct.len() < DISTINCT_CAP {
        sheet_distinct.insert(rendered.clone());
    }
    if book_distinct.len() < DISTINCT_CAP {
        book_distinct.insert(rendered);
    }
}

/// The whole of a workbook as text, one relation row a line in the fixed
/// order: the API's surface, the tests' and a door's that holds the text.
pub fn reflect_text(bytes: &[u8], label: &str) -> Result<String, Refusal> {
    let source = open(bytes, label)?;
    let mut printer = print::Printer::default();
    source.header_rows(&mut printer);
    for i in 0..source.sheets().len() {
        source.walk_sheet(i, &mut printer)?;
    }
    Ok(printer.out)
}
