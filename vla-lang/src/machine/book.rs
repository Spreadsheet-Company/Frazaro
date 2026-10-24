//! The machine's grid (KERNEL.49's decisions 1 and 2, its third slice,
//! 2026-10-22): an engine's machine holds its grid in a representation of
//! its own beside `sheet::Workbook`, which stays the file model of the
//! writer, the readers and the doors. A sheet of it is its name and
//! settings, its fills, each one formula written over a rectangle and held
//! as the pieces the writes into it leave, its formulas alone, each a cell's
//! own text, and the few cells' formats; its values stand by place in its
//! store (`super::store`), so that no cell of a fill is ever made, and what a
//! cell holds is found where it stands.
//!
//! A write into part of a fill takes the written rectangle out of each piece
//! it meets and keeps the rest as at most four, each the same fill, its
//! shape, its anchor and its index (decision 2), so that every cell keeps
//! its text as the record moves it; a fill written over whole is gone. A
//! fill whose first cell is written over is given its first cell left as its
//! anchor, its text moved there, or each of its cells its own text where the
//! move would not keep the text whole, as a write has left a fill since
//! KERNEL.34; so a fill's anchor is always its first cell, row by row.

use std::collections::BTreeMap;

use crate::calc::graph::{Area, CellId, Extent};
use crate::grid::{cut, overlap, Formula, Kind, Masters, Model, Rects, RowMajor};
use crate::sheet::{A1Range, Sheet, Styles};

use super::load::{Draft, What};

/// A fill: one formula written over a rectangle, its text without its `=`
/// read at its first cell, its anchor, every other cell's that text moved
/// there; the range it was written over, which a workbook lent names; the
/// rectangle its pieces have stood in, where they are searched for; and how
/// many cells it holds.
#[derive(Clone, Debug, PartialEq)]
pub(crate) struct Fill {
    pub(crate) anchor: (u32, u32),
    pub(crate) text: String,
    pub(crate) written: A1Range,
    pub(crate) reach: A1Range,
    pub(crate) cells: u64,
}

/// A sheet's fills, each by the shared formula's index its sheet gave it,
/// which the plan places it by, and their pieces, which never overlap, each
/// with its fill's index.
#[derive(Clone, Debug, Default, PartialEq)]
pub(crate) struct Fills {
    by: BTreeMap<u32, Fill>,
    pieces: Rects,
}

impl Fills {
    /// A piece of a fill kept, the fill made at its first piece.
    pub(crate) fn add(
        &mut self,
        si: u32,
        range: A1Range,
        anchor: (u32, u32),
        text: &str,
        written: A1Range,
    ) {
        let fill = self.by.entry(si).or_insert_with(|| Fill {
            anchor,
            text: text.to_string(),
            written,
            reach: range,
            cells: 0,
        });
        let r = fill.reach;
        fill.reach = A1Range {
            top: r.top.min(range.top),
            left: r.left.min(range.left),
            bottom: r.bottom.max(range.bottom),
            right: r.right.max(range.right),
        };
        fill.cells += range.cells();
        self.pieces.insert(range, si);
    }

    /// A fill by its index.
    pub(crate) fn fill(&self, si: u32) -> Option<&Fill> {
        self.by.get(&si)
    }

    /// The fill holding a cell, by its index.
    pub(crate) fn at(&self, row: u32, col: u32) -> Option<u32> {
        self.pieces.at(row, col).map(|(_, si)| si)
    }

    /// Whether a piece meets an area.
    pub(crate) fn meets(&self, area: &A1Range) -> bool {
        self.pieces.meeting(*area).next().is_some()
    }

    /// The pieces meeting an area, each the cells it shares with it and its
    /// fill's index.
    pub(crate) fn pieces_in(&self, area: &A1Range) -> Vec<(A1Range, u32)> {
        self.pieces
            .meeting(*area)
            .filter_map(|(piece, si)| Some((overlap(&piece, area)?, si)))
            .collect()
    }

    /// Every piece, with its fill's index: the plan's nodes (KERNEL.49's
    /// decision 7) and the tests'.
    pub(crate) fn every_piece(&self) -> impl Iterator<Item = (A1Range, u32)> + '_ {
        self.pieces.iter()
    }

    /// The fills' masters, by index: each fill's text and its anchor.
    pub(crate) fn masters(&self) -> Masters<'_> {
        #[cfg(test)]
        crate::sheet::work::masters(self.by.len());
        self.by
            .iter()
            .map(|(&si, f)| (si, (f.text.as_str(), f.anchor.0, f.anchor.1)))
            .collect()
    }

    /// A rectangle taken out of every piece it meets, the rest of each kept
    /// as at most four (decision 2) and its fill's count lowered, a fill
    /// emptied gone; the fills it met, by index, each once.
    pub(crate) fn cut(&mut self, out: &A1Range) -> Vec<u32> {
        let met: Vec<(A1Range, u32)> = self.pieces.meeting(*out).collect();
        let mut touched: Vec<u32> = Vec::with_capacity(met.len());
        for (whole, si) in met {
            self.pieces.remove(&whole);
            let (left_over, shared) = cut(whole, out);
            for r in left_over {
                self.pieces.insert(r, si);
            }
            if let (Some(shared), Some(f)) = (shared, self.by.get_mut(&si)) {
                f.cells -= shared.cells();
                if f.cells == 0 {
                    self.by.remove(&si);
                }
            }
            touched.push(si);
        }
        touched.sort_unstable();
        touched.dedup();
        touched
    }

    /// A fill's first cell, row by row: its first piece's, among the pieces
    /// in the rectangle its pieces have stood in.
    pub(crate) fn first_cell(&self, si: u32) -> Option<(u32, u32)> {
        let reach = self.by.get(&si)?.reach;
        self.pieces
            .meeting(reach)
            .filter(|&(_, s)| s == si)
            .map(|(r, _)| (r.top, r.left))
            .min()
    }

    /// A fill's anchor and text made another's (a write that moved its first
    /// cell, KERNEL.34).
    pub(crate) fn anchor_at(&mut self, si: u32, anchor: (u32, u32), text: String) {
        if let Some(f) = self.by.get_mut(&si) {
            f.anchor = anchor;
            f.text = text;
        }
    }

    /// A fill given up whole, its pieces with it; its cells, row by row.
    pub(crate) fn remove(&mut self, si: u32) -> Option<(Fill, Vec<(u32, u32)>)> {
        let fill = self.by.remove(&si)?;
        let mine: Vec<A1Range> = self
            .pieces
            .meeting(fill.reach)
            .filter(|&(_, s)| s == si)
            .map(|(r, _)| r)
            .collect();
        for r in &mine {
            self.pieces.remove(r);
        }
        let cells = RowMajor::new(mine.into_iter().map(|r| (r, 0)).collect())
            .map(|(row, col, _)| (row, col))
            .collect();
        Some((fill, cells))
    }
}

/// One sheet of the machine's grid.
#[derive(Clone, Debug)]
pub(crate) struct Page {
    /// Its name and settings as the file model's sheet holds them, its
    /// sentences and its count of shared formulas among them; no cell.
    pub(crate) sheet: Sheet,
    pub(crate) fills: Fills,
    /// The formulas of their own, a cell each: the text without its `=`,
    /// and its kind.
    pub(crate) alone: BTreeMap<(u32, u32), (String, Kind)>,
    /// Each cell's format, where it is not 0.
    pub(crate) styles: BTreeMap<(u32, u32), u32>,
    /// How many cells it holds, values and formulas.
    pub(crate) cells: usize,
}

impl Page {
    pub(crate) fn new(sheet: Sheet) -> Page {
        Page {
            sheet,
            fills: Fills::default(),
            alone: BTreeMap::new(),
            styles: BTreeMap::new(),
            cells: 0,
        }
    }

    /// What a cell holds where it holds a formula.
    pub(crate) fn formula(&self, row: u32, col: u32) -> Option<Formula<'_>> {
        if let Some((text, kind)) = self.alone.get(&(row, col)) {
            return Some(Formula::Own(text, *kind));
        }
        self.fills.at(row, col).map(Formula::Fill)
    }

    /// A cell's format.
    pub(crate) fn style(&self, row: u32, col: u32) -> u32 {
        self.styles.get(&(row, col)).copied().unwrap_or(0)
    }

    /// Every formula cell inside an area, row by row, with what it holds:
    /// the fills' pieces swept and the formulas alone beside them.
    pub(crate) fn formulas_in(&self, area: &A1Range, f: &mut dyn FnMut((u32, u32), Formula<'_>)) {
        for (at, formula) in self.formulas_iter(area) {
            f(at, formula);
        }
    }

    /// [`Page::formulas_in`]'s walk, as an iterator.
    pub(crate) fn formulas_iter<'a>(
        &'a self,
        area: &A1Range,
    ) -> impl Iterator<Item = ((u32, u32), Formula<'a>)> + 'a {
        let mut fills = RowMajor::new(self.fills.pieces_in(area)).peekable();
        let (left, right) = (area.left, area.right);
        let mut alone = self
            .alone
            .range((area.top, area.left)..=(area.bottom, area.right))
            .filter(move |(&(_, col), _)| col >= left && col <= right)
            .peekable();
        std::iter::from_fn(move || {
            let take_fill = match (fills.peek(), alone.peek()) {
                (None, None) => return None,
                (Some(_), None) => true,
                (None, Some(_)) => false,
                (Some(&(row, col, _)), Some((&at, _))) => (row, col) < at,
            };
            Some(if take_fill {
                let (row, col, si) = fills.next().expect("peeked");
                ((row, col), Formula::Fill(si))
            } else {
                let (&at, (text, kind)) = alone.next().expect("peeked");
                (at, Formula::Own(text.as_str(), *kind))
            })
        })
    }
}

/// The machine's grid: its sheets, the rows' and then their twins', which
/// hold no formula; its format table and its names.
#[derive(Clone, Debug)]
pub(crate) struct Book {
    pub(crate) pages: Vec<Page>,
    pub(crate) styles: Styles,
    pub(crate) defined_names: Vec<(String, String)>,
}

impl Book {
    /// The rows' sheets of a draft, each record of a range never made into
    /// its cells: a formula of one cell a formula alone, a range's a fill's
    /// piece, its values left to the stores; each formula's text moved out
    /// of its record, so that a load holds it once at its peak.
    pub(crate) fn of(draft: &mut Draft) -> Book {
        let mut pages = Vec::with_capacity(draft.sheets.len() * 2);
        for ds in &mut draft.sheets {
            let mut page = Page::new(ds.sheet.clone());
            page.styles = std::mem::take(&mut ds.styles);
            for record in ds.records.iter_mut().flatten() {
                let range = record.range;
                page.cells += range.cells() as usize;
                match &mut record.what {
                    What::Value(_) => {}
                    What::Formula {
                        text,
                        si: None,
                        kind,
                        ..
                    } => {
                        for row in range.top..=range.bottom {
                            for col in range.left..=range.right {
                                page.alone.insert((row, col), (std::mem::take(text), *kind));
                            }
                        }
                    }
                    What::Formula {
                        text,
                        anchor,
                        si: Some(si),
                        written,
                        ..
                    } => page.fills.add(*si, range, *anchor, text, *written),
                }
            }
            pages.push(page);
        }
        Book {
            pages,
            styles: std::mem::take(&mut draft.styles),
            defined_names: std::mem::take(&mut draft.defined_names),
        }
    }

    /// A sheet's index by name, compared as Excel compares names
    /// ([`crate::sheet::Workbook::find_sheet`]).
    pub(crate) fn find_sheet(&self, name: &str) -> Option<usize> {
        self.pages
            .iter()
            .position(|p| p.sheet.name.eq_ignore_ascii_case(name))
    }
}

impl Model for Book {
    fn sheet_count(&self) -> usize {
        self.pages.len()
    }

    fn sheet_name(&self, sheet: usize) -> Option<&str> {
        self.pages.get(sheet).map(|p| p.sheet.name.as_str())
    }

    fn find_sheet(&self, name: &str) -> Option<usize> {
        Book::find_sheet(self, name)
    }

    fn defined_names(&self) -> &[(String, String)] {
        &self.defined_names
    }

    fn formula(&self, (s, row, col): CellId) -> Option<Formula<'_>> {
        self.pages.get(s)?.formula(row, col)
    }

    fn masters(&self, sheet: usize) -> Masters<'_> {
        self.pages
            .get(sheet)
            .map(|p| p.fills.masters())
            .unwrap_or_default()
    }

    fn master(&self, sheet: usize, si: u32) -> Option<(&str, u32, u32)> {
        let f = self.pages.get(sheet)?.fills.fill(si)?;
        Some((f.text.as_str(), f.anchor.0, f.anchor.1))
    }

    fn each_formula(&self, sheet: usize, f: &mut dyn FnMut((u32, u32), Formula<'_>)) {
        if let Some(p) = self.pages.get(sheet) {
            let all = A1Range {
                top: 1,
                left: 1,
                bottom: crate::sheet::MAX_ROW,
                right: crate::sheet::MAX_COLUMN,
            };
            p.formulas_in(&all, f);
        }
    }

    fn formulas_in(&self, area: &Area, extent: Extent, f: &mut dyn FnMut(u32, u32)) {
        let (Some(p), Some(((top, left), (bottom, right)))) = (self.pages.get(area.sheet), extent)
        else {
            return;
        };
        let clamped = A1Range {
            top: area.top.max(top),
            left: area.left.max(left),
            bottom: area.bottom.min(bottom),
            right: area.right.min(right),
        };
        if clamped.top > clamped.bottom || clamped.left > clamped.right {
            return;
        }
        p.formulas_in(&clamped, &mut |(row, col), _| f(row, col));
    }

    /// A sheet's fills as their pieces and its formulas alone, the plan's
    /// nodes (KERNEL.49's decision 7), with no cell of a fill walked.
    fn each_node(&self, sheet: usize, f: &mut dyn FnMut(A1Range, Formula<'_>)) {
        let Some(p) = self.pages.get(sheet) else {
            return;
        };
        for (piece, si) in p.fills.every_piece() {
            f(piece, Formula::Fill(si));
        }
        for (&(row, col), (text, kind)) in &p.alone {
            let cell = A1Range {
                top: row,
                left: col,
                bottom: row,
                right: col,
            };
            f(cell, Formula::Own(text, *kind));
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn r(top: u32, left: u32, bottom: u32, right: u32) -> A1Range {
        A1Range {
            top,
            left,
            bottom,
            right,
        }
    }

    #[test]
    fn a_fill_cut_keeps_its_cells_outside_the_cut_and_finds_each_by_place() {
        // KERNEL.49's decision 2: two fills, one cut through its middle and
        // one at its corner, every cell left found as its fill's and every
        // cell cut found as none; a fill cut whole gone.
        let mut fills = Fills::default();
        fills.add(0, r(1, 1, 6, 6), (1, 1), "A1", r(1, 1, 6, 6));
        fills.add(1, r(1, 8, 3, 9), (1, 8), "B1", r(1, 8, 3, 9));
        assert_eq!(fills.cut(&r(3, 2, 4, 9)), vec![0, 1]);
        for row in 1..=6 {
            for col in 1..=9 {
                let cut = (3..=4).contains(&row) && (2..=9).contains(&col);
                let want = match col {
                    1..=6 if !cut => Some(0),
                    8..=9 if row <= 3 && !cut => Some(1),
                    _ => None,
                };
                assert_eq!(fills.at(row, col), want, "({row}, {col})");
            }
        }
        assert_eq!(fills.fill(0).map(|f| f.cells), Some(36 - 10));
        assert_eq!(fills.fill(1).map(|f| f.cells), Some(4));
        assert_eq!(fills.first_cell(0), Some((1, 1)));
        // The rest of fill 1 cut away: it is gone.
        assert_eq!(fills.cut(&r(1, 8, 2, 9)), vec![1]);
        assert_eq!(fills.fill(1), None);
        assert_eq!(fills.at(1, 8), None);
        // Given up whole, its cells row by row.
        let (fill, cells) = fills.remove(0).expect("fill 0");
        assert_eq!(fill.cells, 26);
        assert_eq!(cells.len(), 26);
        let mut sorted = cells.clone();
        sorted.sort_unstable();
        assert_eq!(cells, sorted, "row by row");
        assert_eq!(fills.every_piece().count(), 0);
    }
}
