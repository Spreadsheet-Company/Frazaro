//! The loader (KERNEL.22, 2026-10-09): rows to a grid, the inverse of the
//! printers in `crate::rows` and `crate::view`, so that what a door prints
//! the machine loads, and a sheet viewed whole and loaded again views the
//! same (`Alonzo/SPEC.md` section 7.3, and its free oracle in section 13).
//! Since KERNEL.49's third slice the rows are read into a draft of records,
//! each row that writes cells its sheet, its rectangle, its line and its
//! value or its formula's text, a range never made into its cells
//! ([`Draft`]); a door's workbook is made from them whole when one is asked
//! for ([`Draft::workbook`]), and a machine's grid without one.
//!
//! The rows taken: `sheet`, `cell` and `formula` over a cell or a range (a
//! value fills the range, a formula fills it as a shared formula, moved as
//! Excel moves a filled formula), `value` after its formula (the formula's
//! last value, never a second write) or since KERNEL.72 over a range, every
//! cell of it a formula, one value every cell's or as many as its cells, row
//! by row, a save's spelling (`crate::view`), `name`, `gridlines`, `column`,
//! `row`,
//! `format`, `look`, `style` and `sentence`; a `format` row's number format
//! is `general`, `text` or, since KERNEL.47, a code of Excel's in quotes,
//! refused by `grid-format-code` naming what it holds that this version
//! does not draw (`crate::numfmt`), and since KERNEL.50 its wrap is followed
//! by what it says of a font, an alignment and a border, each a group
//! (`crate::appearance`), refused by `grid-format-property` naming what a
//! group holds that a format does not take; since KERNEL.58 a format of
//! groups alone says no more than its groups, three more saying a fill, a
//! number format and a wrap; and a `conditional` row holds a rule on its
//! sheet, whose range needs no earlier row, and a `lit` row a rule's last answer
//! at a cell, as a `value` row a formula's, a rule with no such row waiting for
//! the first step. A `style` row names a cell or,
//! since KERNEL.50, a range, every cell of it written by an earlier row, as
//! a `sentence` row's cell is; a cell no earlier row wrote is refused by
//! `grid-cell-unwritten`, naming it. Since KERNEL.67 a range's format is held
//! as its rectangle and never a format a cell (`grid::CellStyles`), and its
//! cells are found written by counting what its records hold of it. Skipped
//! by name, being a view's or a reader's own derivation or a thing no model
//! holds before `KERNEL.8`:
//! `window`, `extent`, `refers`, `table` and `text`, a number's text under
//! its code, which a save prints and the code makes again. A `sheet` row
//! naming `X.last` is taken as the twin's own declaration, which every
//! record of a save lists, when the rows hold `X`; the machine makes the
//! twins itself.
//!
//! One write per cell (decision 23 of the page): a cell the rows write
//! twice, by two rows or by a range and a row inside it, is refused naming
//! both lines, at the first such cell in row-major order, found by the
//! records' rectangles and never by a cell; a setting given twice, as a save
//! repeats format 0 in every record it concatenates, and until KERNEL.72 a
//! sheet row, is taken when it says the same thing and refused when it does
//! not. A grid holds at most [`CELLS`] cells beside its twins, so that one
//! range row cannot ask for the seventeen billion of Excel's sheet.

use std::collections::{BTreeMap, HashMap};
use std::sync::OnceLock;

use crate::appearance::{self, Wrong};
use crate::calc::{CellId, Computed, ErrorKind, Extent, Reason, Value};
use crate::form::Form;
use crate::grid::{outside, outside_held, overlap, CellStyles, Kind, Rects, RowMajor};
use crate::intrinsics::fold;
use crate::messages::raise;
use crate::numfmt::{Code, Unsupported};
use crate::printer::write_datum;
use crate::refers::{quote_sheet, shift_a1_references};
use crate::rows::{quoted, Visibility};
use crate::rule::{is_number, Operator, Rule, Stop, StopKind, Test, Uncomputed, BAR_LENGTHS};
use crate::sheet::{
    cell_ref, number_text, parse_a1_range, A1Range, Cell, Column, Content, NumFmt, Rgb,
    RowSettings, Sheet, Style, Styles, Unsaid, Workbook, MAX_COLUMN, MAX_ROW,
};

use super::{LineRefusal, CELLS};

/// The rows a load takes, the five it skips last, as `grid-row-unknown`
/// lists them.
pub const LOAD_ROWS: &str = "sheet, cell, formula, value, name, gridlines, column, row, format, look, conditional, merge, style, lit, sentence, window, extent, refers, table and text";

/// The most formats a grid's table holds, Excel's own count of cell
/// formats a workbook may hold, rounded to a power of two.
pub const FORMATS: u32 = 65_536;

/// The grid as the rows wrote it, before a machine takes it (KERNEL.49's
/// decision 8): each sheet's name and settings, and each row that wrote
/// cells as a record of its sheet, its rectangle, its line and its value or
/// its formula's text, a range's never made into its cells and no line kept
/// for a cell, so that what a load holds before the machine is the size of
/// its rows. An engine reads it between the two calls, [`read_rows`] and
/// `Machine::new` (its device layouts, its Clock): a sheet by name, made
/// when new ([`Draft::find_sheet`], [`Draft::ensure_sheet`]); a cell's
/// content ([`Draft::content`]); a sheet's cells walked ([`Draft::cells`]);
/// every formula's text, once a record ([`Draft::formulas`]); a cell put or
/// taken out, counted ([`Draft::put`], [`Draft::take`]); and a cell's line
/// ([`Draft::line`]). A door's workbook is made from it whole
/// ([`Draft::workbook`]).
#[derive(Clone, Debug)]
pub struct Draft {
    pub(crate) sheets: Vec<DraftSheet>,
    /// How many of the sheets the rows name: all of them, until a machine
    /// makes each rule's answer sheet after them (KERNEL.58).
    pub(crate) rows: usize,
    pub(crate) styles: Styles,
    pub(crate) defined_names: Vec<(String, String)>,
    /// Each formula cell's last value by its cell, made from the `value` rows
    /// the first time a caller asks ([`Draft::values`]), which a load never
    /// does (KERNEL.72).
    pub(crate) values: OnceLock<HashMap<CellId, Computed>>,
    /// How many cells the rows wrote.
    pub(crate) cells: usize,
}

/// One sheet of a draft: its name and settings, and its records.
#[derive(Clone, Debug)]
pub(crate) struct DraftSheet {
    /// Its name and settings as the file model's sheet holds them, its
    /// sentences and its count of shared formulas among them; no cell.
    pub(crate) sheet: Sheet,
    /// Each record, in the order the rows wrote them; `None` where one was
    /// taken out whole.
    pub(crate) records: Vec<Option<Record>>,
    /// The records of more than one cell, each by its rectangle.
    ranges: Rects,
    /// The records of one cell, by place.
    ones: BTreeMap<(u32, u32), usize>,
    /// Each cell's format, as rectangles (KERNEL.67): the cells a row's
    /// format names, a row's format 0 among them, read as none.
    pub(crate) styles: CellStyles,
    /// The line of each rule on its sheet, by priority (KERNEL.58).
    rule_lines: BTreeMap<u32, u32>,
    /// Each `lit` row: the rules it says applied at a cell, and its line.
    lit: BTreeMap<(u32, u32), (Vec<u32>, u32)>,
    /// Its `value` rows (KERNEL.72).
    pub(crate) said: SaidRows,
}

/// A rule as a machine answers it (KERNEL.58): its priority, its answer sheet,
/// and whether it stops.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) struct Answering {
    pub(crate) priority: u32,
    pub(crate) sheet: usize,
    pub(crate) stop: bool,
}

/// A rule a machine plans only inside its sheet's extent, its grid's rules past the
/// room (KERNEL.58's *Found* 15): its rows' sheet, its place among that sheet's rules,
/// its answer sheet, and how far it is planned.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) struct Planned {
    pub(crate) sheet: usize,
    pub(crate) at: usize,
    pub(crate) answer: usize,
    pub(crate) area: Option<A1Range>,
}

/// An extent as the rectangle it is.
pub(crate) fn extent_area(extent: Extent) -> Option<A1Range> {
    extent.map(|((top, left), (bottom, right))| A1Range {
        top,
        left,
        bottom,
        right,
    })
}

/// What of a rule's cells stands inside `area` and outside `was`, as pieces that never
/// overlap: what a write growing its sheet's extent plans of it.
pub(crate) fn pieces_past(
    rule: &Rule,
    was: Option<A1Range>,
    area: Option<A1Range>,
) -> Vec<A1Range> {
    let mut held = Rects::default();
    let mut past = Vec::new();
    for (inside, new) in [(was, false), (area, true)] {
        let Some(inside) = inside else { continue };
        for range in rule.ranges.iter().filter_map(|r| overlap(r, &inside)) {
            for piece in outside_held(range, &held) {
                held.insert(piece, 0);
                if new {
                    past.push(piece);
                }
            }
        }
    }
    past
}

/// What a `value` row says of the formula cells of its range (KERNEL.72): one
/// value, every cell's; or each cell's, row by row, kept as numbers where
/// they all are, 8 bytes a cell where a value is 32.
#[derive(Clone, Debug)]
pub(crate) enum Said {
    One(Computed),
    Numbers(Vec<f64>),
    Each(Vec<Computed>),
}

/// A sheet's `value` rows (KERNEL.72): each its range, its line and what it
/// says, kept as a draft keeps its records, by rectangle or by place, so that
/// a cell said twice is found by what its row's range meets, no cell of a
/// range made, and no hashed entry a cell of a save.
#[derive(Clone, Debug, Default)]
pub(crate) struct SaidRows {
    pub(crate) rows: Vec<(A1Range, u32, Said)>,
    ranges: Rects,
    ones: BTreeMap<(u32, u32), usize>,
}

impl SaidRows {
    /// The first cell of a range, in row-major order, a row already says, and
    /// that row's line.
    fn first_said(&self, range: &A1Range) -> Option<((u32, u32), u32)> {
        let (at, k) = first_in(&self.ranges, &self.ones, range)?;
        Some((at, self.rows[k].1))
    }

    fn add(&mut self, range: A1Range, line: u32, said: Said) {
        let k = self.rows.len();
        if range.cells() == 1 {
            self.ones.insert((range.top, range.left), k);
        } else {
            self.ranges.insert(range, k as u32);
        }
        self.rows.push((range, line, said));
    }
}

/// The first cell of a range, in row-major order, among rectangles and cells
/// alone that never overlap, and the record or row holding it: each rectangle
/// meeting the range found by the rows they share, its first cell there the
/// corner they share.
fn first_in(
    ranges: &Rects,
    ones: &BTreeMap<(u32, u32), usize>,
    range: &A1Range,
) -> Option<((u32, u32), usize)> {
    let mut first: Option<((u32, u32), usize)> = None;
    let mut meet = |at: (u32, u32), k: usize| {
        if !first.is_some_and(|(f, _)| at >= f) {
            first = Some((at, k));
        }
    };
    for (r, k) in ranges.meeting(*range) {
        if let Some(shared) = overlap(&r, range) {
            meet((shared.top, shared.left), k as usize);
        }
    }
    // The cells written alone, by the rows the range spans where they are
    // fewer than its rows, else row by row.
    let rows = u64::from(range.bottom - range.top) + 1;
    if (ones.len() as u64) < rows {
        for (&(row, col), &k) in ones.range((range.top, 0)..=(range.bottom, u32::MAX)) {
            if col >= range.left && col <= range.right {
                meet((row, col), k);
                break;
            }
        }
    } else {
        for row in range.top..=range.bottom {
            let mut in_row = ones.range((row, range.left)..=(row, range.right));
            if let Some((&at, &k)) = in_row.next() {
                meet(at, k);
                break;
            }
        }
    }
    first
}

/// One row's cells: their rectangle, the row's line, and what they hold.
#[derive(Clone, Debug, PartialEq)]
pub(crate) struct Record {
    pub(crate) range: A1Range,
    pub(crate) line: u32,
    pub(crate) what: What,
}

/// What a record's cells hold: a value, every cell the same; or a formula,
/// its text without its `=` at its first cell, `anchor`, each other cell's
/// that text moved there; with the shared formula's index the row's range
/// of more than one cell gave it and that range, which a workbook made
/// from the draft names (a fill cut by a cell taken keeps all of it, its
/// first cell left its anchor, as a machine's write leaves one, KERNEL.34);
/// and its kind, which a row writes legacy and a workbook may say otherwise
/// (KERNEL.8, slice f).
#[derive(Clone, Debug, PartialEq)]
pub(crate) enum What {
    Value(Content),
    Formula {
        text: String,
        anchor: (u32, u32),
        si: Option<u32>,
        written: A1Range,
        kind: Kind,
    },
}

/// A cell of a draft as a reader of it meets it (KERNEL.49's decision 8).
#[derive(Clone, Copy, Debug, PartialEq)]
pub enum Drafted<'a> {
    /// A value cell's content.
    Value(&'a Content),
    /// A formula cell: its formula's text without its `=` where its row
    /// wrote it, at the first cell of its range, and the cell's distance
    /// from there.
    Formula { text: &'a str, moved: (i64, i64) },
}

impl Drafted<'_> {
    /// Whether the cell holds a formula.
    pub fn is_formula(&self) -> bool {
        matches!(self, Drafted::Formula { .. })
    }

    /// The cell's content: a value as written, a formula as its own text,
    /// moved to the cell.
    pub fn content(&self) -> Content {
        match *self {
            Drafted::Value(content) => content.clone(),
            Drafted::Formula {
                text,
                moved: (0, 0),
            } => Content::Formula(text.to_string()),
            Drafted::Formula {
                text,
                moved: (d_row, d_col),
            } => Content::Formula(shift_a1_references(text, d_row, d_col)),
        }
    }
}

impl DraftSheet {
    fn new(name: &str) -> DraftSheet {
        DraftSheet {
            sheet: Sheet::new(name),
            records: Vec::new(),
            ranges: Rects::default(),
            ones: BTreeMap::new(),
            styles: CellStyles::default(),
            rule_lines: BTreeMap::new(),
            lit: BTreeMap::new(),
            said: SaidRows::default(),
        }
    }

    /// The rectangle of its records' cells, as `Sheet::extent` walks a
    /// sheet's cells for it, no cell walked (KERNEL.49).
    pub(crate) fn extent(&self) -> Extent {
        let mut out: Extent = None;
        for r in self.records.iter().flatten() {
            let x = r.range;
            out = Some(match out {
                None => ((x.top, x.left), (x.bottom, x.right)),
                Some(((top, left), (bottom, right))) => (
                    (top.min(x.top), left.min(x.left)),
                    (bottom.max(x.bottom), right.max(x.right)),
                ),
            });
        }
        out
    }

    /// A record kept, by its first cell.
    fn insert(&mut self, record: Record) -> usize {
        let at = (record.range.top, record.range.left);
        let k = self.records.len();
        if record.range.cells() == 1 {
            self.ones.insert(at, k);
        } else {
            // A record's index a `u32`, the cell cap holding the records far
            // below it.
            self.ranges.insert(record.range, k as u32);
        }
        self.records.push(Some(record));
        k
    }

    /// A record given up: its keys gone and its place left empty.
    fn remove(&mut self, k: usize) -> Option<Record> {
        let record = self.records.get_mut(k)?.take()?;
        let at = (record.range.top, record.range.left);
        match record.range.cells() {
            1 => self.ones.remove(&at),
            _ => self.ranges.remove(&record.range).map(|k| k as usize),
        };
        Some(record)
    }

    /// The record holding a cell.
    pub(crate) fn find(&self, row: u32, col: u32) -> Option<usize> {
        if let Some(&k) = self.ones.get(&(row, col)) {
            return Some(k);
        }
        self.ranges.at(row, col).map(|(_, k)| k as usize)
    }

    /// Whether the records hold every cell of a range (KERNEL.67): the cells
    /// they hold in it counted, the ranges' by the rectangles they share with
    /// it and the cells written alone among its rows, the records never
    /// overlapping, so that no cell of the range is looked up.
    fn covers(&self, range: &A1Range) -> bool {
        self.held_in(range, &|_| true) == range.cells()
    }

    /// Whether every cell of a range holds a formula (KERNEL.72), counted as
    /// [`DraftSheet::covers`] counts.
    fn all_formulas(&self, range: &A1Range) -> bool {
        self.held_in(range, &|r| matches!(r.what, What::Formula { .. })) == range.cells()
    }

    /// How many cells of a range the records `taken` lets by hold, the records
    /// never overlapping.
    fn held_in(&self, range: &A1Range, taken: &dyn Fn(&Record) -> bool) -> u64 {
        let taken = |k: usize| self.records[k].as_ref().is_some_and(taken);
        let mut held: u64 = self
            .ranges
            .meeting(*range)
            .filter(|&(_, k)| taken(k as usize))
            .filter_map(|(r, _)| overlap(&r, range))
            .map(|shared| shared.cells())
            .sum();
        // The cells written alone, by the rows the range spans where they are
        // fewer than its rows, else row by row.
        let rows = u64::from(range.bottom - range.top) + 1;
        held += if (self.ones.len() as u64) < rows {
            self.ones
                .range((range.top, 0)..=(range.bottom, u32::MAX))
                .filter(|(&(_, col), &k)| col >= range.left && col <= range.right && taken(k))
                .count() as u64
        } else {
            (range.top..=range.bottom)
                .map(|row| {
                    self.ones
                        .range((row, range.left)..=(row, range.right))
                        .filter(|(_, &k)| taken(k))
                        .count() as u64
                })
                .sum()
        };
        held
    }

    /// The first cell of a range, in row-major order, a record already
    /// holds, and that record.
    fn first_written(&self, range: &A1Range) -> Option<((u32, u32), usize)> {
        first_in(&self.ranges, &self.ones, range)
    }

    /// The cell's content as the draft holds it.
    fn drafted(&self, k: usize, (row, col): (u32, u32)) -> Option<Drafted<'_>> {
        Some(match &self.records[k].as_ref()?.what {
            What::Value(content) => Drafted::Value(content),
            What::Formula { text, anchor, .. } => Drafted::Formula {
                text,
                moved: (
                    i64::from(row) - i64::from(anchor.0),
                    i64::from(col) - i64::from(anchor.1),
                ),
            },
        })
    }

    /// One cell taken out of the record holding it, the rest of the record
    /// kept as at most four (KERNEL.49's decision 2); a fill's cut at its
    /// first cell given its first cell left, its text moved there, or each
    /// of its cells its own text where the move would not keep the text
    /// whole, as a machine's write leaves a fill (KERNEL.34). Whether a cell
    /// was there.
    fn cut_cell(&mut self, row: u32, col: u32) -> bool {
        let Some(k) = self.find(row, col) else {
            return false;
        };
        let record = self.remove(k).expect("found");
        let out = A1Range {
            top: row,
            left: col,
            bottom: row,
            right: col,
        };
        let (pieces, _) = crate::grid::cut(record.range, &out);
        for range in pieces {
            self.insert(Record {
                range,
                line: record.line,
                what: record.what.clone(),
            });
        }
        if let What::Formula {
            text,
            anchor,
            si: Some(si),
            ..
        } = &record.what
        {
            if *anchor == (row, col) {
                self.anchor_again(*si, (row, col), text);
            }
        }
        true
    }

    /// A fill whose first cell was taken out, given its first cell left as
    /// its anchor, its text moved there, where the move keeps the text whole;
    /// else each of its cells a formula of its own, its text moved there.
    fn anchor_again(&mut self, si: u32, was: (u32, u32), text: &str) {
        let pieces: Vec<usize> = self
            .records
            .iter()
            .enumerate()
            .filter(|(_, r)| matches!(r, Some(Record { what: What::Formula { si: Some(s), .. }, .. }) if *s == si))
            .map(|(k, _)| k)
            .collect();
        let Some(first) = pieces
            .iter()
            .filter_map(|&k| self.records[k].as_ref())
            .map(|r| (r.range.top, r.range.left))
            .min()
        else {
            return;
        };
        let at = (
            i64::from(first.0) - i64::from(was.0),
            i64::from(first.1) - i64::from(was.1),
        );
        if crate::refers::moves_whole(text, at, at) {
            let moved = shift_a1_references(text, at.0, at.1);
            for k in pieces {
                if let Some(Record {
                    what: What::Formula { text, anchor, .. },
                    ..
                }) = &mut self.records[k]
                {
                    text.clone_from(&moved);
                    *anchor = first;
                }
            }
            return;
        }
        for k in pieces {
            let record = self.remove(k).expect("a piece");
            for row in record.range.top..=record.range.bottom {
                for col in record.range.left..=record.range.right {
                    let own = shift_a1_references(
                        text,
                        i64::from(row) - i64::from(was.0),
                        i64::from(col) - i64::from(was.1),
                    );
                    let cell = A1Range {
                        top: row,
                        left: col,
                        bottom: row,
                        right: col,
                    };
                    self.insert(Record {
                        range: cell,
                        line: record.line,
                        what: What::Formula {
                            text: own,
                            anchor: (row, col),
                            si: None,
                            written: cell,
                            kind: Kind::Plain,
                        },
                    });
                }
            }
        }
    }
}

impl Draft {
    fn new() -> Draft {
        Draft {
            sheets: Vec::new(),
            rows: 0,
            styles: Styles::new(),
            defined_names: Vec::new(),
            values: OnceLock::new(),
            cells: 0,
        }
    }

    /// How many sheets the rows hold, the engine's own made since among them.
    pub fn sheets(&self) -> usize {
        self.sheets.len()
    }

    /// A sheet's name, by index.
    pub fn sheet_name(&self, sheet: usize) -> Option<&str> {
        self.sheets.get(sheet).map(|s| s.sheet.name.as_str())
    }

    /// A sheet's index by name, compared as Excel compares names
    /// ([`Workbook::find_sheet`]).
    pub fn find_sheet(&self, name: &str) -> Option<usize> {
        self.sheets
            .iter()
            .position(|s| s.sheet.name.eq_ignore_ascii_case(name))
    }

    /// The sheet of that name, added last with the name as given when there
    /// is none.
    pub fn ensure_sheet(&mut self, name: &str) -> usize {
        if let Some(i) = self.find_sheet(name) {
            return i;
        }
        self.sheets.push(DraftSheet::new(name));
        self.sheets.len() - 1
    }

    /// A cell as the rows wrote it; `None` for a cell no row wrote.
    pub fn content(&self, sheet: usize, row: u32, col: u32) -> Option<Drafted<'_>> {
        let s = self.sheets.get(sheet)?;
        s.drafted(s.find(row, col)?, (row, col))
    }

    /// A sheet's cells, row by row, each as the rows wrote it: every cell
    /// of a value's range and of a formula's, made one at a time as the walk
    /// meets it and never kept.
    pub fn cells(&self, sheet: usize) -> impl Iterator<Item = ((u32, u32), Drafted<'_>)> + '_ {
        let rects = self.sheets.get(sheet).map_or_else(Vec::new, |s| {
            s.records
                .iter()
                .enumerate()
                .filter_map(|(k, r)| Some((r.as_ref()?.range, k as u32)))
                .collect()
        });
        let s = self.sheets.get(sheet);
        RowMajor::new(rects).filter_map(move |(row, col, k)| {
            Some(((row, col), s?.drafted(k as usize, (row, col))?))
        })
    }

    /// Every formula the rows wrote, once a record: its sheet and the first
    /// cell of its range, its text without its `=` at its anchor, and its
    /// row's line; by sheet and first cell, the order a walk of the cells
    /// meets a fill's first cell in. A rule's among them (KERNEL.58), at its
    /// range's corner: it computes as a cell's does, so an engine's check of
    /// what a formula may call holds it too.
    pub fn formulas(&self) -> impl Iterator<Item = (CellId, &str, u32)> + '_ {
        let mut out: Vec<(CellId, &str, u32)> = Vec::new();
        for (si, s) in self.sheets.iter().enumerate() {
            for r in s.records.iter().flatten() {
                if let What::Formula { text, .. } = &r.what {
                    out.push(((si, r.range.top, r.range.left), text.as_str(), r.line));
                }
            }
            for rule in s.sheet.rules() {
                let (row, col) = rule.corner();
                let line = s.rule_lines.get(&rule.priority).copied().unwrap_or(0);
                for text in rule.formulas() {
                    out.push(((si, row, col), text, line));
                }
            }
        }
        out.sort_by_key(|&(id, _, _)| id);
        out.into_iter()
    }

    /// A value written into one cell by the engine, as a `cell` row of one
    /// cell writes it, at no line: the cell taken out of the record that
    /// held it first, its format kept, and counted when it is new.
    pub fn put(&mut self, sheet: usize, row: u32, col: u32, content: Content) {
        if !self.sheets[sheet].cut_cell(row, col) {
            self.cells += 1;
        }
        let one = A1Range {
            top: row,
            left: col,
            bottom: row,
            right: col,
        };
        self.sheets[sheet].insert(Record {
            range: one,
            line: 0,
            what: What::Value(content),
        });
    }

    /// One cell taken out of the draft, its format with it, and no longer
    /// counted; whether the rows wrote it.
    pub fn take(&mut self, sheet: usize, row: u32, col: u32) -> bool {
        let Some(s) = self.sheets.get_mut(sheet) else {
            return false;
        };
        if !s.cut_cell(row, col) {
            return false;
        }
        s.styles.take(row, col);
        self.cells -= 1;
        true
    }

    /// The line of the row that wrote a cell, 0 for one the engine put.
    pub fn line(&self, sheet: usize, row: u32, col: u32) -> Option<u32> {
        let s = self.sheets.get(sheet)?;
        Some(s.records[s.find(row, col)?].as_ref()?.line)
    }

    /// How many cells the rows wrote, and the engine since.
    pub fn count(&self) -> usize {
        self.cells
    }

    /// Each formula cell's last value, from its `value` row: made by its
    /// cell the first time it is asked for (KERNEL.72).
    pub fn values(&self) -> &HashMap<CellId, Computed> {
        self.values.get_or_init(|| {
            let mut out = HashMap::new();
            for (si, s) in self.sheets.iter().enumerate() {
                for (range, _, said) in &s.said.rows {
                    let mut k = 0;
                    for row in range.top..=range.bottom {
                        for col in range.left..=range.right {
                            let value = match said {
                                Said::One(c) => c.clone(),
                                Said::Numbers(n) => Computed::Value(Value::Number(n[k])),
                                Said::Each(c) => c[k].clone(),
                            };
                            out.insert((si, row, col), value);
                            k += 1;
                        }
                    }
                }
            }
            out
        })
    }

    /// The grid as a workbook, every cell of every record made, a formula
    /// over a range a shared formula from its first cell, as the writer and
    /// a door's view take it.
    pub fn workbook(&self) -> Workbook {
        let mut wb = Workbook::new();
        wb.styles = self.styles.clone();
        wb.defined_names = self.defined_names.clone();
        for s in &self.sheets {
            let mut sheet = s.sheet.clone();
            for record in s.records.iter().flatten() {
                let range = record.range;
                for row in range.top..=range.bottom {
                    for col in range.left..=range.right {
                        let content = match &record.what {
                            What::Value(content) => content.clone(),
                            What::Formula {
                                text,
                                si: None,
                                kind,
                                ..
                            } => match kind {
                                Kind::Plain => Content::Formula(text.clone()),
                                Kind::Dynamic => Content::DynamicFormula(text.clone()),
                                Kind::Array => Content::ArrayFormula(text.clone()),
                            },
                            What::Formula {
                                text,
                                anchor,
                                si: Some(si),
                                written,
                                ..
                            } if *anchor == (row, col) => Content::SharedMaster {
                                text: text.clone(),
                                range: written.text(),
                                si: *si,
                            },
                            What::Formula { si: Some(si), .. } => Content::SharedChild { si: *si },
                        };
                        let style = s.styles.at(row, col);
                        sheet.set(row, col, Cell { content, style });
                    }
                }
            }
            wb.sheets.push(sheet);
        }
        wb
    }

    /// A draft of a workbook's cells, each a record of its own at no line, a
    /// fill's cells pieces of one fill as its master names it, a fill's cell
    /// whose master the sheet has lost left out: what a door's workbook, or
    /// the grid a machine lends ([`super::Machine::workbook`]), loads as, its
    /// fills and its formulas' kinds kept (KERNEL.49).
    pub fn of_workbook(workbook: &Workbook) -> Draft {
        let mut draft = Draft::new();
        draft.styles = workbook.styles.clone();
        draft.defined_names = workbook.defined_names.clone();
        for sheet in &workbook.sheets {
            let mut s = DraftSheet::new(&sheet.name);
            s.sheet = Sheet {
                cells: BTreeMap::new(),
                ..sheet.clone()
            };
            let masters = crate::sheet::shared_masters(sheet);
            for (&(row, col), cell) in &sheet.cells {
                let one = A1Range {
                    top: row,
                    left: col,
                    bottom: row,
                    right: col,
                };
                let own = |text: &String, kind: Kind| What::Formula {
                    text: text.clone(),
                    anchor: (row, col),
                    si: None,
                    written: one,
                    kind,
                };
                let what = match &cell.content {
                    Content::SharedMaster { text, range, si } => What::Formula {
                        text: text.clone(),
                        anchor: (row, col),
                        si: Some(*si),
                        written: parse_a1_range(range).unwrap_or(one),
                        kind: Kind::Plain,
                    },
                    Content::SharedChild { si } => {
                        let Some(&(text, m_row, m_col)) = masters.get(si) else {
                            continue;
                        };
                        let written = match &sheet.cells[&(m_row, m_col)].content {
                            Content::SharedMaster { range, .. } => {
                                parse_a1_range(range).unwrap_or(one)
                            }
                            _ => one,
                        };
                        What::Formula {
                            text: text.to_string(),
                            anchor: (m_row, m_col),
                            si: Some(*si),
                            written,
                            kind: Kind::Plain,
                        }
                    }
                    Content::Formula(text) => own(text, Kind::Plain),
                    Content::DynamicFormula(text) => own(text, Kind::Dynamic),
                    Content::ArrayFormula(text) => own(text, Kind::Array),
                    value => What::Value(value.clone()),
                };
                if cell.style != 0 {
                    s.styles.put(one, cell.style);
                }
                s.insert(Record {
                    range: one,
                    line: 0,
                    what,
                });
                draft.cells += 1;
            }
            draft.sheets.push(s);
        }
        draft
    }

    /// Each rule's answer sheet made after the rows' sheets (KERNEL.58's
    /// decision 6), hidden and named as no reference can spell: its rectangles
    /// one fill of the rule's answer, read at the range's corner, a rectangle
    /// overlapping one before it cut, and each `lit` row its last answer, as a
    /// `value` row a formula's. A grid of no rule makes none. For each rows'
    /// sheet its rules, none at all where no sheet holds one, and how many
    /// cells they answer for, at most [`CELLS`] beside the grid's. Rules asking
    /// more than that room are each planned only inside their sheet's extent, and
    /// said so ([`Planned`]): what no record of a sheet whole shows, so that rules over
    /// whole columns cost what their sheets hold (KERNEL.58's *Found* 15).
    pub(crate) fn answer_sheets(&mut self) -> Result<Answers, LineRefusal> {
        let rows = self.sheets.len();
        self.rows = rows;
        if self.sheets.iter().all(|s| s.sheet.rules().is_empty()) {
            return Ok((Vec::new(), 0, Vec::new()));
        }
        let asked: u64 = self
            .sheets
            .iter()
            .flat_map(|s| s.sheet.rules())
            .filter(|rule| rule.computed())
            .map(Rule::cells)
            .sum();
        let past_room = asked > CELLS as u64;
        let mut planned = Vec::new();
        let mut cells: u64 = 0;
        let mut answering = Vec::with_capacity(rows);
        for si in 0..rows {
            let name = self.sheets[si].sheet.name.clone();
            let rules: Vec<Rule> = self.sheets[si].sheet.rules().to_vec();
            let area = extent_area(self.sheets[si].extent()).filter(|_| past_room);
            let mut mine = Vec::with_capacity(rules.len());
            for (at, rule) in rules.iter().enumerate() {
                if !rule.computed() {
                    continue;
                }
                let line = self.sheets[si]
                    .rule_lines
                    .get(&rule.priority)
                    .copied()
                    .unwrap_or(0);
                let mut answer = DraftSheet::new(&format!("{name}[{}]", rule.priority));
                answer.sheet.visibility = Visibility::Hidden;
                let text = rule.answer_text(&name);
                let (anchor, written) = (rule.corner(), rule.around());
                let mut before = Rects::default();
                for &range in &rule.ranges {
                    let range = if past_room {
                        match area.and_then(|area| overlap(&range, &area)) {
                            Some(inside) => inside,
                            None => continue,
                        }
                    } else {
                        range
                    };
                    for piece in outside_held(range, &before) {
                        before.insert(piece, 0);
                        cells += piece.cells();
                        answer.insert(Record {
                            range: piece,
                            line,
                            what: What::Formula {
                                text: text.clone(),
                                anchor,
                                si: Some(0),
                                written,
                                kind: Kind::Plain,
                            },
                        });
                    }
                    if cells > CELLS as u64 {
                        return Err(rule_cells(line, &name, &rule.range_text(), cells));
                    }
                }
                mine.push(Answering {
                    priority: rule.priority,
                    sheet: self.sheets.len(),
                    stop: rule.stop,
                });
                if past_room {
                    planned.push(Planned {
                        sheet: si,
                        at,
                        answer: self.sheets.len(),
                        area,
                    });
                }
                self.sheets.push(answer);
            }
            // Each rule a `lit` row names, by its priority, the rules being in
            // that order: a save of many rules and many lit cells is read in
            // the time of its rows.
            let lit = std::mem::take(&mut self.sheets[si].lit);
            for ((row, col), (lit, line)) in lit {
                for priority in &lit {
                    if let Ok(k) = mine.binary_search_by_key(priority, |a: &Answering| a.priority) {
                        let cell = A1Range {
                            top: row,
                            left: col,
                            bottom: row,
                            right: col,
                        };
                        let answered = Computed::Value(Value::Number(1.0));
                        self.sheets[mine[k].sheet]
                            .said
                            .add(cell, line, Said::One(answered));
                    }
                }
            }
            answering.push(mine);
        }
        Ok((answering, cells as usize, planned))
    }

    /// A formula cell's kind said: a legacy array formula, as a workbook may
    /// hold one and no row writes, or a dynamic-array formula, which a row's
    /// word says of one cell since KERNEL.70 (a machine read over either kind,
    /// the tests' own, KERNEL.8, slice f).
    #[cfg(test)]
    pub(crate) fn set_kind(&mut self, sheet: usize, row: u32, col: u32, to: Kind) {
        let s = &mut self.sheets[sheet];
        let k = s.find(row, col).expect("a cell");
        if let Some(Record {
            what: What::Formula { kind, .. },
            ..
        }) = &mut s.records[k]
        {
            *kind = to;
        }
    }
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

/// A `formula` row's fields: its sheet, its address and what it writes, and its
/// kind, legacy, or dynamic by the word after them (KERNEL.70), as Excel 365
/// reads the same text typed; `None` for any other shape. The third is any form,
/// for the caller to read as it reads one.
pub(crate) fn formula_row(items: &[Form]) -> Option<(&String, &String, &Form, Kind)> {
    match items {
        [Form::Str(sheet), Form::Str(addr), third] => Some((sheet, addr, third, Kind::Plain)),
        [Form::Str(sheet), Form::Str(addr), third, Form::Sym(word)] if word == "dynamic" => {
            Some((sheet, addr, third, Kind::Dynamic))
        }
        _ => None,
    }
}

/// A dynamic formula's row over more than one cell (KERNEL.70): a person types
/// one cell, a fill holds no kind, and a spill is no fill.
pub(crate) fn dynamic_range(line: u32, sheet: &str, addr: &str) -> LineRefusal {
    LineRefusal {
        line,
        refusal: raise(
            "grid-dynamic-range",
            &[("target", &format!("{sheet}!{addr}"))],
        ),
    }
}

/// The shapes a row is held to, in the notation, for `grid-row-malformed`.
pub(crate) fn shape_of(kind: &str) -> &'static str {
    match kind {
        "sheet" => "(sheet \"<name>\" visible|hidden|very-hidden)",
        "cell" => "(cell \"<sheet>\" \"<cell or range>\" <value>)",
        "formula" => "(formula \"<sheet>\" \"<cell or range>\" \"=<text>\") or (formula \"<sheet>\" \"<cell>\" \"=<text>\" dynamic)",
        "derived" => "(derived \"<sheet>\" \"<cell or range>\" <value>)",
        "value" => {
            "(formula \"<sheet>\" \"<cell or range>\" \"=<text>\") (value \"<sheet>\" \"<cell or range>\" <value> ...)"
        }
        "name" => "(name \"<name>\" \"<refers-to>\")",
        "gridlines" => "(gridlines \"<sheet>\" on|off)",
        "column" => "(column \"<sheet>\" \"<letters>\" <width>|none shown|hidden <format>|none)",
        "row" => "(row \"<sheet>\" <number> <height>|none shown|hidden)",
        "format" => "(format <index> \"<RRGGBB>\"|none general|text|\"<code>\" wrap|nowrap [(font ...)] [(align ...)] [(border ...)]) or a format of groups alone, (format <index> [(fill ...)] [(number ...)] [(wrap ...)] [(font ...)] [(align ...)] [(border ...)]), saying at least one",
        "look" => "(look \"<sheet>\" <value> <format>)",
        "conditional" => "(conditional \"<sheet>\" <priority> \"<range>\" \"=<formula>\"|(color-scale ...)|(data-bar ...)|(icon-set ...)|(top <rank> [percent])|(bottom <rank> [percent])|(above-average ...)|(below-average ...)|(duplicate)|(unique)|(cell <operator> \"=<bound>\" [\"=<bound>\"]) <format>|none [stop]), its operator one of less, less-or-equal, equal, not-equal, greater-or-equal, greater, between and not-between, the last two taking two bounds",
        "merge" => "(merge \"<sheet>\" \"<range of more than one cell>\")",
        "lit" => "(conditional \"<sheet>\" <priority> \"<range>\" <test> <format>|none) (lit \"<sheet>\" \"<cell>\" <priority> ...)",
        "style" => "(cell \"<sheet>\" \"<cell or range>\" <value>) (style \"<sheet>\" \"<cell or range>\" <format>)",
        "sentence" => {
            "(cell \"<sheet>\" \"<cell>\" <value>) (sentence \"<sheet>\" \"<cell>\" <row>)"
        }
        _ => "(<kind> ...)",
    }
}

// The refusals below carry the row's line in the answer's own field and
// never in their words (KERNEL.23), so a write of one row, whose line is
// always 1, does not say it.

/// `grid-row-malformed` for a form.
pub(crate) fn malformed(line: u32, form: &Form, kind: &str) -> LineRefusal {
    LineRefusal {
        line,
        refusal: raise(
            "grid-row-malformed",
            &[("row", &write_datum(form)), ("shape", shape_of(kind))],
        ),
    }
}

/// `grid-format-code` for a `format` row's code that holds what this version
/// does not draw (KERNEL.47).
pub(crate) fn format_code(line: u32, code: &str, unsupported: &Unsupported) -> LineRefusal {
    LineRefusal {
        line,
        refusal: raise(
            "grid-format-code",
            &[
                ("code", &quoted(code)),
                ("what", &unsupported.what),
                ("fix", &unsupported.fix),
            ],
        ),
    }
}

/// `grid-format-property` for a `format` row's group that holds what a
/// format does not take (KERNEL.50), with what the group takes.
pub(crate) fn format_property(
    line: u32,
    index: &str,
    group: &str,
    what: &str,
    fix: &str,
) -> LineRefusal {
    LineRefusal {
        line,
        refusal: raise(
            "grid-format-property",
            &[
                ("index", index),
                ("group", group),
                ("what", what),
                ("fix", fix),
            ],
        ),
    }
}

/// `grid-cell-unwritten` for a `style` or `sentence` row naming a cell that
/// no row before it wrote (KERNEL.50), the cell named.
pub(crate) fn unwritten(
    line: u32,
    form: &Form,
    kind: &str,
    sheet: &str,
    (row, col): (u32, u32),
) -> LineRefusal {
    LineRefusal {
        line,
        refusal: raise(
            "grid-cell-unwritten",
            &[
                ("row", &write_datum(form)),
                ("cell", &format!("{sheet}!{}", cell_ref(row, col))),
                ("kind", kind),
            ],
        ),
    }
}

/// `grid-row-unknown` for a row whose head is not one taken.
pub(crate) fn unknown(line: u32, head: &str, kinds: &str) -> LineRefusal {
    LineRefusal {
        line,
        refusal: raise("grid-row-unknown", &[("head", head), ("kinds", kinds)]),
    }
}

/// `grid-cell-written-twice`: `what` written at `line` and first at `first`.
pub(crate) fn twice(line: u32, what: &str, first: u32) -> LineRefusal {
    LineRefusal {
        line,
        refusal: raise(
            "grid-cell-written-twice",
            &[("what", what), ("first", &first.to_string())],
        ),
    }
}

/// `grid-sheet-name-invalid` for a name a grid cannot hold.
pub(crate) fn bad_name(line: u32, name: &str) -> LineRefusal {
    LineRefusal {
        line,
        refusal: raise("grid-sheet-name-invalid", &[("name", name)]),
    }
}

/// `grid-write-last`: a row writing cells of a twin, told to write the same
/// cells of the sheet the twin copies.
pub(crate) fn into_twin(line: u32, sheet: &str, addr: &str) -> LineRefusal {
    let source = twin_base(sheet).unwrap_or(sheet);
    LineRefusal {
        line,
        refusal: raise(
            "grid-write-last",
            &[
                ("target", &format!("{sheet}!{addr}")),
                ("sheet", sheet),
                ("source", &format!("{source}!{addr}")),
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
                ("target", &format!("{sheet}!{addr}")),
                ("cells", &thousands(count)),
                ("limit", &thousands(CELLS as u64)),
            ],
        ),
    }
}

/// How many rectangles a `lit` row reads of a rule one by one: past them it reads
/// the rule's pieces, found by their place.
const FEW_RECTANGLES: usize = 16;

/// A rule's cells as pieces that never overlap, what its answer sheet holds.
fn rule_pieces(rule: &Rule) -> Rects {
    let mut pieces = Rects::default();
    for &range in &rule.ranges {
        for piece in outside_held(range, &pieces) {
            pieces.insert(piece, 0);
        }
    }
    pieces
}

/// What a draft's rules make of it: each sheet's rules as a machine answers them,
/// how many cells they answer for, and each planned inside its sheet's extent.
pub(crate) type Answers = (Vec<Vec<Answering>>, usize, Vec<Planned>);

/// `grid-merge-overlap`: a merged range over one already merged (KERNEL.58).
fn merge_overlap(line: u32, sheet: &str, range: &A1Range, held: &A1Range) -> LineRefusal {
    LineRefusal {
        line,
        refusal: raise(
            "grid-merge-overlap",
            &[
                ("target", &format!("{sheet}!{}", range.text())),
                ("held", &held.text()),
            ],
        ),
    }
}

/// `grid-rule-cells`: a rule whose range would bring the cells rules answer for
/// past their cap (KERNEL.58).
pub(crate) fn rule_cells(line: u32, sheet: &str, range: &str, count: u64) -> LineRefusal {
    LineRefusal {
        line,
        refusal: raise(
            "grid-rule-cells",
            &[
                ("target", &format!("{sheet}!{range}")),
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
    /// The pieces of a rule of many rectangles a `lit` row has named, by its sheet
    /// and its place, `None` for one past the room, whose pieces might be more than
    /// its rows.
    lit_rules: BTreeMap<(usize, usize), Option<Rects>>,
    sheet_rows: HashMap<usize, (Visibility, u32)>,
    twin_rows: Vec<(String, u32)>,
    names: HashMap<String, (String, u32)>,
    gridlines: HashMap<usize, (bool, u32)>,
    columns: HashMap<(usize, u32), u32>,
    row_settings: HashMap<(usize, u32), u32>,
    formats: HashMap<u32, (Style, u32)>,
    looks: HashMap<(usize, String), (u32, u32)>,
    /// Each `style` row taken, its sheet, its range and its line, in their
    /// order: the first holding a cell gave it its format, and is the line a
    /// `twice` refusal names (KERNEL.67).
    styled: Vec<(usize, A1Range, u32)>,
    sentences: HashMap<CellId, (u32, u32)>,
}

/// The forms of a text of rows read into a draft, the inverse of the record:
/// every row taken or skipped by name, or the first refusal with its line.
/// Nothing is computed and no twin is made: that is `Machine::new`'s.
pub fn read_rows(forms: &[Form]) -> Result<Draft, LineRefusal> {
    let mut r = Reader::new();
    for form in forms {
        r.row(form)?;
    }
    r.finish()
}

/// [`read_rows`] over forms made one at a time (KERNEL.72;
/// [`crate::reader::each_form`]), each given up before the next is made, so
/// that the read holds the draft and one row: [`super::Machine::load`]'s, and
/// an engine's reading its rows between the two calls.
pub fn read_rows_each(forms: impl IntoIterator<Item = Form>) -> Result<Draft, LineRefusal> {
    let mut r = Reader::new();
    for form in forms {
        r.row(&form)?;
    }
    r.finish()
}

impl Reader {
    fn new() -> Reader {
        Reader {
            draft: Draft::new(),
            lit_rules: BTreeMap::new(),
            sheet_rows: HashMap::new(),
            twin_rows: Vec::new(),
            names: HashMap::new(),
            gridlines: HashMap::new(),
            columns: HashMap::new(),
            row_settings: HashMap::new(),
            formats: HashMap::new(),
            looks: HashMap::new(),
            styled: Vec::new(),
            sentences: HashMap::new(),
        }
    }

    /// The rows all read: a twin's declaration stands only beside the sheet
    /// it is the twin of.
    fn finish(self) -> Result<Draft, LineRefusal> {
        for (name, line) in &self.twin_rows {
            let base = twin_base(name).unwrap_or("");
            if self.draft.find_sheet(base).is_none() {
                return Err(bad_name(*line, name));
            }
        }
        Ok(self.draft)
    }

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
            "conditional" => self.conditional(form, items, line),
            "merge" => self.merge(form, items, line),
            "lit" => self.lit(form, items, line),
            "style" => self.style(form, items, line),
            "sentence" => self.sentence(form, items, line),
            "window" | "extent" | "refers" | "table" | "text" => Ok(()),
            other => Err(unknown(line, other, LOAD_ROWS)),
        }
    }

    /// A sheet of the grid by name, made when it is new; a twin's name and a
    /// name no sheet may hold refused. A row writing cells (`addr` given)
    /// into a twin is `grid-write-last`; a setting naming a twin is
    /// `grid-sheet-name-invalid`, as a `sheet` row naming one alone is,
    /// since a twin copies its sheet's values and never its settings.
    fn sheet_index(
        &mut self,
        name: &str,
        line: u32,
        addr: Option<&str>,
    ) -> Result<usize, LineRefusal> {
        if let Some(i) = self.draft.find_sheet(name) {
            return Ok(i);
        }
        if twin_base(name).is_some() {
            return Err(match addr {
                Some(addr) => into_twin(line, name, addr),
                None => bad_name(line, name),
            });
        }
        if !sheet_name_ok(name) {
            return Err(bad_name(line, name));
        }
        Ok(self.draft.ensure_sheet(name))
    }

    /// A sheet's name and settings.
    fn settings(&mut self, si: usize) -> &mut Sheet {
        &mut self.draft.sheets[si].sheet
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
        let i = self.sheet_index(name, line, None)?;
        if let Some((v, first)) = self.sheet_rows.get(&i) {
            if *v != visibility {
                return Err(twice(line, &format!("(sheet {})", quoted(name)), *first));
            }
            return Ok(());
        }
        self.sheet_rows.insert(i, (visibility, line));
        self.settings(i).visibility = visibility;
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
        let i = self.sheet_index(sheet, line, Some(addr))?;
        Ok((i, range))
    }

    /// The cells a range row writes, refused when one was written before,
    /// the first such in row-major order named with its row's line, or when
    /// they would bring the grid past its cap: found by the records the
    /// range meets, and no cell of it made (KERNEL.49's decision 8).
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
        let s = &self.draft.sheets[si];
        if let Some(((row, col), k)) = s.first_written(&range) {
            let what = format!("{}!{}", quote_sheet(&s.sheet.name), cell_ref(row, col));
            let first = s.records[k].as_ref().map_or(0, |r| r.line);
            return Err(twice(line, &what, first));
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
        self.draft.sheets[si].insert(Record {
            range,
            line,
            what: What::Value(value.content()),
        });
        Ok(())
    }

    fn formula(&mut self, form: &Form, items: &[Form], line: u32) -> Result<(), LineRefusal> {
        let Some((sheet, addr, third, kind)) = formula_row(items) else {
            return Err(malformed(line, form, "formula"));
        };
        let Form::Str(text) = third else {
            return Err(malformed(line, form, "formula"));
        };
        let Some(body) = text.strip_prefix('=').filter(|b| !b.is_empty()) else {
            return Err(malformed(line, form, "formula"));
        };
        let (si, range) = self.target(form, "formula", sheet, addr, line)?;
        if kind == Kind::Dynamic && range.cells() != 1 {
            return Err(dynamic_range(line, sheet, addr));
        }
        self.claim(si, range, sheet, addr, line)?;
        // A range's formula a shared formula, its index the sheet's next, as
        // `Sheet::set_formula` numbers them.
        let shared = self.settings(si).next_shared(range);
        self.draft.sheets[si].insert(Record {
            range,
            line,
            what: What::Formula {
                text: body.to_string(),
                anchor: (range.top, range.left),
                si: shared,
                written: range,
                kind,
            },
        });
        Ok(())
    }

    /// A `value` row: a formula's last value, or since KERNEL.72 a range's,
    /// every cell of it a formula an earlier row wrote: one value, every
    /// cell's, or as many as its cells, row by row. Each cell's value said
    /// once, the first cell a row says again refused with the line of the row
    /// that said it, spelled as the row spells it where the row names one cell.
    fn value(&mut self, form: &Form, items: &[Form], line: u32) -> Result<(), LineRefusal> {
        let [Form::Str(sheet), Form::Str(addr), values @ ..] = items else {
            return Err(malformed(line, form, "value"));
        };
        let said = said_of(values).ok_or_else(|| malformed(line, form, "value"))?;
        let Some(range) =
            parse_a1_range(addr).filter(|r| values.len() == 1 || values.len() as u64 == r.cells())
        else {
            return Err(malformed(line, form, "value"));
        };
        let Some(si) = self.draft.find_sheet(sheet) else {
            return Err(malformed(line, form, "value"));
        };
        let s = &mut self.draft.sheets[si];
        if !s.all_formulas(&range) {
            return Err(malformed(line, form, "value"));
        }
        if let Some(((row, col), first)) = s.said.first_said(&range) {
            let at = match range.cells() {
                1 => addr.clone(),
                _ => cell_ref(row, col),
            };
            let what = format!("(value {} {})", quoted(sheet), quoted(&at));
            return Err(twice(line, &what, first));
        }
        s.said.add(range, line, said);
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
        self.draft.defined_names.push((name.clone(), text.clone()));
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
        let si = self.sheet_index(sheet, line, None)?;
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
        self.settings(si).gridlines = on;
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
        let si = self.sheet_index(sheet, line, None)?;
        let column = Column {
            index,
            width,
            hidden,
            style,
        };
        if let Some(first) = self.columns.get(&(si, index)) {
            let said = self.draft.sheets[si]
                .sheet
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
        self.settings(si).columns.push(column);
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
        let si = self.sheet_index(sheet, line, None)?;
        let settings = RowSettings { height, hidden };
        if let Some(first) = self.row_settings.get(&(si, n)) {
            if self.draft.sheets[si].sheet.rows.get(&n) != Some(&settings) {
                return Err(twice(line, &format!("(row {} {n})", quoted(sheet)), *first));
            }
            return Ok(());
        }
        self.row_settings.insert((si, n), line);
        self.settings(si).rows.insert(n, settings);
        Ok(())
    }

    fn format(&mut self, form: &Form, items: &[Form], line: u32) -> Result<(), LineRefusal> {
        let Some(Form::Sym(index)) = items.first() else {
            return Err(malformed(line, form, "format"));
        };
        let index = match index.parse::<u32>() {
            Ok(i) if i < FORMATS => i,
            _ => return Err(malformed(line, form, "format")),
        };
        let (mut style, groups) = match &items[1..] {
            // A format of groups alone (KERNEL.58), a list where a fill's
            // place stands: all it says is in them.
            groups @ [Form::List(_), ..] => (Style::partial(), groups),
            [fill, num, Form::Sym(wrap), groups @ ..] => {
                let fill = match fill {
                    Form::Sym(f) if f == "none" => None,
                    Form::Str(hex) => match Rgb::from_hex(hex) {
                        Some(c) => Some(c),
                        None => return Err(malformed(line, form, "format")),
                    },
                    _ => return Err(malformed(line, form, "format")),
                };
                let num_fmt = match num {
                    Form::Sym(s) if s == "general" => NumFmt::General,
                    Form::Sym(s) if s == "text" => NumFmt::Text,
                    // A code of Excel's (KERNEL.47), refused naming what it
                    // holds that this version does not draw.
                    Form::Str(code) => match Code::parse(code) {
                        Ok(code) => NumFmt::Code(code),
                        Err(unsupported) => return Err(format_code(line, code, &unsupported)),
                    },
                    _ => return Err(malformed(line, form, "format")),
                };
                let wrap = match wrap.as_str() {
                    "wrap" => true,
                    "nowrap" => false,
                    _ => return Err(malformed(line, form, "format")),
                };
                (Style::new(fill, num_fmt, wrap), groups)
            }
            _ => return Err(malformed(line, form, "format")),
        };
        // What its groups say of a font, an alignment and a border
        // (KERNEL.50), and in a format of groups alone of a fill, a number
        // format and a wrap (KERNEL.58), which a format spelling the three
        // says in their places alone; a group holding what a format does not
        // take refused naming it.
        match appearance::read_groups(groups) {
            Ok(said) => {
                if said.says_the_places() && style.unsaid == Unsaid::NOTHING {
                    return Err(malformed(line, form, "format"));
                }
                if let Some(fill) = said.fill {
                    style.say_fill(fill);
                }
                if let Some(number) = said.number {
                    style.say_number(number);
                }
                if let Some(wrap) = said.wrap {
                    style.say_wrap(wrap);
                }
                style.font = said.font;
                style.family = said.family.map(Box::new);
                style.align = said.align;
                style.border = said.border;
            }
            Err(Wrong::Shape) => return Err(malformed(line, form, "format")),
            Err(Wrong::Property { group, what, fix }) => {
                return Err(format_property(line, &index.to_string(), group, &what, fix))
            }
            Err(Wrong::Code { code, why }) => return Err(format_code(line, &code, &why)),
        }
        if let Some((said, first)) = self.formats.get(&index) {
            if *said != style {
                return Err(twice(line, &format!("(format {index})"), *first));
            }
            return Ok(());
        }
        self.formats.insert(index, (style.clone(), line));
        self.draft.styles.put(index, style);
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
        let si = self.sheet_index(sheet, line, None)?;
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
        self.settings(si).looks.push((spelled, index));
        Ok(())
    }

    /// A `conditional` row (KERNEL.58): a rule on its sheet, the sheet made
    /// when it is new, a twin refused as a setting naming one is. A second row
    /// of its priority is taken when it says the same and refused when it does
    /// not, as a look is.
    fn conditional(&mut self, form: &Form, items: &[Form], line: u32) -> Result<(), LineRefusal> {
        let bad = || malformed(line, form, "conditional");
        let (items, stop) = match items {
            [before @ .., Form::Sym(word)] if word == "stop" => (before, true),
            _ => (items, false),
        };
        let [Form::Str(sheet), Form::Sym(priority), Form::Str(range), test, Form::Sym(format)] =
            items
        else {
            return Err(bad());
        };
        let priority = priority.parse::<u32>().map_err(|_| bad())?;
        let ranges: Option<Vec<A1Range>> = range.split(' ').map(parse_a1_range).collect();
        let ranges = ranges.ok_or_else(bad)?;
        let test = rule_test(test).ok_or_else(bad)?;
        let format = match format.as_str() {
            "none" => None,
            n => match n.parse::<u32>() {
                Ok(i) if i < FORMATS => Some(i),
                _ => return Err(bad()),
            },
        };
        if let Test::Uncomputed(held) = &test {
            if !held.takes_a_format() && (format.is_some() || stop) {
                return Err(bad());
            }
        }
        let rule = Rule {
            priority,
            ranges,
            test,
            format,
            stop,
        };
        let si = self.sheet_index(sheet, line, None)?;
        let s = &mut self.draft.sheets[si];
        if let Some(&first) = s.rule_lines.get(&priority) {
            if s.sheet.rules().contains(&rule) {
                return Ok(());
            }
            return Err(twice(
                line,
                &format!("(conditional {} {priority})", quoted(sheet)),
                first,
            ));
        }
        s.rule_lines.insert(priority, line);
        s.sheet.put_rule(rule);
        Ok(())
    }

    /// A `merge` row (KERNEL.58): a range of more than one cell, a second alike
    /// taken and one overlapping another refused by name.
    fn merge(&mut self, form: &Form, items: &[Form], line: u32) -> Result<(), LineRefusal> {
        let bad = || malformed(line, form, "merge");
        let [Form::Str(sheet), Form::Str(range)] = items else {
            return Err(bad());
        };
        let area = parse_a1_range(range)
            .filter(|r| r.cells() > 1)
            .ok_or_else(bad)?;
        let si = self.sheet_index(sheet, line, None)?;
        self.draft.sheets[si]
            .sheet
            .put_merge(area)
            .map_err(|held| merge_overlap(line, sheet, &area, &held))
    }

    /// A `lit` row (KERNEL.58): the rules a cell's record says applied, each
    /// a rule of the sheet that answers for the cell, or for a covered cell its
    /// merge's first; a second alike taken.
    fn lit(&mut self, form: &Form, items: &[Form], line: u32) -> Result<(), LineRefusal> {
        let bad = || malformed(line, form, "lit");
        let [Form::Str(sheet), Form::Str(addr), priorities @ ..] = items else {
            return Err(bad());
        };
        if priorities.is_empty() {
            return Err(bad());
        }
        let range = parse_a1_range(addr)
            .filter(|r| r.cells() == 1)
            .ok_or_else(bad)?;
        let si = self.draft.find_sheet(sheet).ok_or_else(bad)?;
        let s = &mut self.draft.sheets[si];
        let (row, col) = s
            .sheet
            .merges()
            .at(range.top, range.left)
            .map_or((range.top, range.left), |first| (first.top, first.left));
        let mut lit = Vec::with_capacity(priorities.len());
        for p in priorities {
            let Form::Sym(p) = p else {
                return Err(bad());
            };
            let p = p.parse::<u32>().map_err(|_| bad())?;
            let rules = s.sheet.rules();
            let Ok(k) = rules.binary_search_by_key(&p, |r| r.priority) else {
                return Err(bad());
            };
            let rule = &rules[k];
            if !rule.computed() {
                return Err(bad());
            }
            let covers = if rule.ranges.len() <= FEW_RECTANGLES {
                rule.covers(row, col)
            } else {
                let pieces = self
                    .lit_rules
                    .entry((si, k))
                    .or_insert_with(|| (rule.cells() <= CELLS as u64).then(|| rule_pieces(rule)));
                match pieces {
                    Some(pieces) => pieces.at(row, col).is_some(),
                    None => rule.covers(row, col),
                }
            };
            if !covers {
                return Err(bad());
            }
            lit.push(p);
        }
        let at = (range.top, range.left);
        if let Some((said, first)) = s.lit.get(&at) {
            if *said != lit {
                let what = format!("(lit {} {})", quoted(sheet), quoted(addr));
                return Err(twice(line, &what, *first));
            }
            return Ok(());
        }
        s.lit.insert(at, (lit, line));
        Ok(())
    }

    /// The sheet of the cells a `style` or `sentence` row names, each of
    /// which an earlier row wrote; the first, in row-major order, that none
    /// did refused by name (KERNEL.50). The cells are found written by a count
    /// of what the records hold of them (KERNEL.67), and only a range they
    /// fall short of is walked, cell by cell, to the cell it names.
    fn written(
        &self,
        form: &Form,
        kind: &str,
        sheet: &str,
        range: A1Range,
        line: u32,
    ) -> Result<usize, LineRefusal> {
        let first = (range.top, range.left);
        let Some(si) = self.draft.find_sheet(sheet) else {
            return Err(unwritten(line, form, kind, sheet, first));
        };
        let s = &self.draft.sheets[si];
        if s.covers(&range) {
            return Ok(si);
        }
        for row in range.top..=range.bottom {
            for col in range.left..=range.right {
                if s.find(row, col).is_none() {
                    return Err(unwritten(line, form, kind, sheet, (row, col)));
                }
            }
        }
        Ok(si)
    }

    /// A `style` row: a cell's format, or since KERNEL.50 every cell of a
    /// range's, as `cell` and `formula` rows take a range; each cell's format
    /// said once, a second alike taken. Since KERNEL.67 the range is held as
    /// its rectangle: the rectangles an earlier row left meeting it, the first
    /// cell, row by row, it shares with one of another format refused, naming
    /// the first row to name that cell, as the cell by cell walk named it;
    /// and the cells its own format's rectangles leave given the format, as
    /// rectangles.
    fn style(&mut self, form: &Form, items: &[Form], line: u32) -> Result<(), LineRefusal> {
        let [Form::Str(sheet), Form::Str(addr), Form::Sym(index)] = items else {
            return Err(malformed(line, form, "style"));
        };
        let index = match index.parse::<u32>() {
            Ok(i) if i < FORMATS => i,
            _ => return Err(malformed(line, form, "style")),
        };
        let range = parse_a1_range(addr).ok_or_else(|| malformed(line, form, "style"))?;
        let si = self.written(form, "style", sheet, range, line)?;
        let mut other: Option<(u32, u32)> = None;
        let mut alike: Vec<A1Range> = Vec::new();
        for (piece, said) in self.draft.sheets[si].styles.meeting(range) {
            if said == index {
                alike.push(piece);
            } else if let Some(shared) = overlap(&piece, &range) {
                let at = (shared.top, shared.left);
                if other.map_or(true, |first| at < first) {
                    other = Some(at);
                }
            }
        }
        if let Some((row, col)) = other {
            let first = self
                .styled
                .iter()
                .find(|&&(s, r, _)| {
                    s == si && row >= r.top && row <= r.bottom && col >= r.left && col <= r.right
                })
                .map(|&(_, _, first)| first)
                .expect("a row gave the cell its format");
            let what = format!("(style {} {})", quoted(sheet), quoted(&cell_ref(row, col)));
            return Err(twice(line, &what, first));
        }
        for piece in outside(range, &alike) {
            self.draft.sheets[si].styles.put(piece, index);
        }
        self.styled.push((si, range, line));
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
        let range = parse_a1_range(addr)
            .filter(|r| r.cells() == 1)
            .ok_or_else(|| malformed(line, form, "sentence"))?;
        let si = self.written(form, "sentence", sheet, range, line)?;
        let id = (si, range.top, range.left);
        if let Some((said, first)) = self.sentences.get(&id) {
            if *said != n {
                let what = format!("(sentence {} {})", quoted(sheet), quoted(addr));
                return Err(twice(line, &what, *first));
            }
            return Ok(());
        }
        self.sentences.insert(id, (n, line));
        self.settings(id.0).sentences.insert((id.1, id.2), n);
        Ok(())
    }
}

/// What a `value` row's values say (KERNEL.72); `None` for no value, or for
/// one that is neither a value nor a label.
fn said_of(values: &[Form]) -> Option<Said> {
    match values {
        [] => None,
        [one] => Some(Said::One(computed(one)?)),
        many => {
            let numbers: Option<Vec<f64>> = many
                .iter()
                .map(|f| match f {
                    Form::Sym(s) => number(s),
                    _ => None,
                })
                .collect();
            Some(match numbers {
                Some(numbers) => Said::Numbers(numbers),
                None => Said::Each(many.iter().map(computed).collect::<Option<_>>()?),
            })
        }
    }
}

/// A `value` row's value: a value, or the label of a formula not computed.
pub(crate) fn computed(f: &Form) -> Option<Computed> {
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

/// A row's `"=<text>"`, without its `=`.
fn formula_said(form: &Form) -> Option<String> {
    match form {
        Form::Str(t) => t
            .strip_prefix('=')
            .filter(|b| !b.is_empty())
            .map(str::to_string),
        _ => None,
    }
}

/// A word of a row, or a list of a word and what follows it.
fn word_of(form: &Form) -> Option<(&str, &[Form])> {
    match form {
        Form::Sym(word) => Some((word, &[])),
        Form::List(list) => match list.items.as_slice() {
            [Form::Sym(word), rest @ ..] => Some((word, rest)),
            _ => None,
        },
        _ => None,
    }
}

/// A stop off its row: `(min)`, `(percent 33)`, `(number "=$A$1")`, `(formula
/// "=$A$1")`, and `greater` after its value.
fn rule_stop(form: &Form) -> Option<Stop> {
    let (word, rest) = word_of(form)?;
    let kind = StopKind::from_word(word)?;
    let (value, rest) = match (kind.says_a_value(), rest) {
        (false, rest) => (String::new(), rest),
        (true, [Form::Sym(n), rest @ ..]) if kind != StopKind::Formula && is_number(n) => {
            (n.clone(), rest)
        }
        (true, [said, rest @ ..]) => (formula_said(said)?, rest),
        _ => return None,
    };
    let greater = match rest {
        [] => false,
        [Form::Sym(word)] if word == "greater" => true,
        _ => return None,
    };
    Some(Stop {
        kind,
        value,
        greater,
    })
}

/// Stops off a row, as many lists as come first, each a stop; and what follows them.
fn rule_stops(items: &[Form]) -> Option<(Vec<Stop>, &[Form])> {
    let count = items
        .iter()
        .take_while(|form| matches!(form, Form::List(_)))
        .count();
    let (stops, rest) = items.split_at(count);
    Some((stops.iter().map(rule_stop).collect::<Option<_>>()?, rest))
}

/// A kind not computed off its row's words, its head taken: `None` for words no such
/// row says.
fn rule_uncomputed(head: &str, items: &[Form]) -> Option<Uncomputed> {
    let color = |form: &Form| match form {
        Form::Str(hex) => Rgb::from_hex(hex),
        _ => None,
    };
    let number = |form: &Form| match form {
        Form::Sym(n) => n.parse::<u32>().ok(),
        _ => None,
    };
    // A word at the end of its row, once; and what stands before it.
    let last_word = |items: &'_ [Form], said: &str| -> (bool, usize) {
        match items.last() {
            Some(Form::Sym(word)) if word == said => (true, items.len() - 1),
            _ => (false, items.len()),
        }
    };
    let held = match head {
        "color-scale" => {
            let (stops, colors) = rule_stops(items)?;
            Uncomputed::ColorScale {
                stops,
                colors: colors.iter().map(color).collect::<Option<_>>()?,
            }
        }
        "data-bar" => {
            let (mut stops, rest) = rule_stops(items)?;
            let (no_value, before) = last_word(rest, "no-value");
            let (bar, lengths) = match &rest[..before] {
                [bar] => (color(bar)?, BAR_LENGTHS),
                [bar, length] => {
                    let (word, sizes) = word_of(length)?;
                    let [low, high] = sizes else {
                        return None;
                    };
                    if word != "length" {
                        return None;
                    }
                    (color(bar)?, (number(low)?, number(high)?))
                }
                _ => return None,
            };
            let high = stops.pop()?;
            let low = stops.pop()?;
            if !stops.is_empty() {
                return None;
            }
            Uncomputed::DataBar {
                low,
                high,
                color: bar,
                lengths,
                shows_value: !no_value,
            }
        }
        "icon-set" => {
            let [Form::Str(set), rest @ ..] = items else {
                return None;
            };
            let (stops, rest) = rule_stops(rest)?;
            let (no_value, before) = last_word(rest, "no-value");
            let (reverse, before) = last_word(&rest[..before], "reverse");
            if before != 0 {
                return None;
            }
            Uncomputed::IconSet {
                set: set.clone(),
                stops,
                reverse,
                shows_value: !no_value,
            }
        }
        "top" | "bottom" => {
            let (percent, before) = last_word(items, "percent");
            let [rank] = &items[..before] else {
                return None;
            };
            Uncomputed::Rank {
                bottom: head == "bottom",
                rank: number(rank)?,
                percent,
            }
        }
        "above-average" | "below-average" => {
            let below = head == "below-average";
            let (equal, deviations) = match items {
                [] => (false, 0),
                [Form::Sym(word)] if word == "equal" => (true, 0),
                [by] => {
                    let (word, count) = word_of(by)?;
                    let [count] = count else {
                        return None;
                    };
                    if word != "deviations" {
                        return None;
                    }
                    (false, number(count).filter(|n| *n > 0)?)
                }
                _ => return None,
            };
            Uncomputed::Average {
                below,
                equal,
                deviations,
            }
        }
        "duplicate" if items.is_empty() => Uncomputed::Duplicate,
        "unique" if items.is_empty() => Uncomputed::Unique,
        _ => return None,
    };
    held.well_said().then_some(held)
}

/// A rule's test off its row: a formula, `"=<text>"`; `(cell <operator>
/// "=<bound>" ["=<bound>"])`, its operator holding it to its bounds; or a kind not
/// computed's words.
fn rule_test(form: &Form) -> Option<Test> {
    let text = formula_said;
    let Form::List(list) = form else {
        return text(form).map(Test::Formula);
    };
    let [Form::Sym(head), after @ ..] = list.items.as_slice() else {
        return None;
    };
    if head != "cell" {
        return rule_uncomputed(head, after).map(Test::Uncomputed);
    }
    let [Form::Sym(word), bounds @ ..] = after else {
        return None;
    };
    let operator = Operator::from_word(word)?;
    let (first, second) = match (bounds, operator.takes_two()) {
        ([first], false) => (text(first)?, None),
        ([first, second], true) => (text(first)?, Some(text(second)?)),
        _ => return None,
    };
    Some(Test::Cell {
        operator,
        first,
        second,
    })
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
