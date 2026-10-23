//! A grid as its readers read it (KERNEL.49's decision 1, its third slice,
//! 2026-10-22): the file model's [`Workbook`], which the writer, the readers
//! and the doors hold, a cell an entry of a sheet's tree, or an engine's
//! machine's own grid, which holds a fill as a rectangle of one formula, each
//! value by place in its sheet's store, and no cell of either. The plan, the
//! graph, the reach and the census read either through [`Model`], and the
//! record and the plane through [`Cells`], so that each is planned or printed
//! by the code that planned or printed a workbook before, and a machine's
//! grid and a workbook holding the same cells read alike. Both traits are
//! the crate's own: a door hands a `Workbook` to the functions it always
//! called.

use std::borrow::Cow;
use std::collections::{BTreeMap, HashMap};

use crate::calc::graph::{cells_in, Area, CellId, Extent};
use crate::rows::Visibility;
use crate::sheet::{
    shared_masters, strip_future_prefixes, A1Range, Column, Content, RowSettings, Sheet, Styles,
    Workbook,
};

/// What a formula cell holds, as a reader finds its text: a fill's cell, by
/// the fill's index on its sheet, its text being the fill's at its first
/// cell, its master ([`Model::masters`]), moved by the cell's distance from
/// it; or a formula of its own, its text and its kind.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum Formula<'a> {
    Fill(u32),
    Own(&'a str, Kind),
}

/// A formula of its own's kind (KERNEL.8, slices b and f): a legacy formula,
/// a dynamic-array formula or a legacy array formula, the last two read as
/// arrays. An engine's grid holds the first alone.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum Kind {
    Plain,
    Dynamic,
    Array,
}

/// A sheet's fills' masters, by index: each fill's text and its first cell.
pub(crate) type Masters<'a> = HashMap<u32, (&'a str, u32, u32)>;

/// A grid as the plan reads it: its sheets by name and by index, its names,
/// and its formula cells, each by what it holds.
pub(crate) trait Model {
    /// How many sheets, a machine's twins among them.
    fn sheet_count(&self) -> usize;
    /// A sheet's name, by index.
    fn sheet_name(&self, sheet: usize) -> Option<&str>;
    /// A sheet's index by name, compared as Excel compares names
    /// ([`Workbook::find_sheet`]).
    fn find_sheet(&self, name: &str) -> Option<usize>;
    /// The defined names, each its name as printed and its text.
    fn defined_names(&self) -> &[(String, String)];
    /// What a cell holds where it holds a formula; `None` for a value cell
    /// and for a cell holding nothing.
    fn formula(&self, id: CellId) -> Option<Formula<'_>>;
    /// A sheet's fills' masters, found once by the caller for every cell it
    /// asks of the sheet: a workbook's by a walk of its cells, a machine's
    /// from its fills.
    fn masters(&self, sheet: usize) -> Masters<'_>;
    /// One fill's master, its text and its first cell: found among the
    /// sheet's masters, unless the grid finds it alone, as a machine finds
    /// it by the fill's index, so that a read of one cell's formula never
    /// walks its sheet.
    fn master(&self, sheet: usize, si: u32) -> Option<(&str, u32, u32)> {
        self.masters(sheet).get(&si).copied()
    }
    /// Every formula cell of a sheet, row by row, with what it holds.
    fn each_formula(&self, sheet: usize, f: &mut dyn FnMut((u32, u32), Formula<'_>));
    /// The formula cells inside an area, row by row, the area clamped to the
    /// extent given first, as [`cells_in`] clamps it.
    fn formulas_in(&self, area: &Area, extent: Extent, f: &mut dyn FnMut(u32, u32));
}

/// A formula's text as the formula bar shows it, without its `=`: a fill's
/// cell's the master's text moved by the cell's distance, moved first and
/// stripped after, as the reader does it; `None` for a fill's cell whose
/// master its sheet has lost, which a machine never leaves
/// ([`crate::sheet::formula_text`]'s rule).
pub(crate) fn text_of(
    formula: Formula<'_>,
    (row, col): (u32, u32),
    masters: &Masters<'_>,
) -> Option<String> {
    match formula {
        Formula::Own(text, _) => Some(strip_future_prefixes(text)),
        Formula::Fill(si) => {
            let &(text, m_row, m_col) = masters.get(&si)?;
            // The master itself reads its own text, unmoved.
            if (row, col) == (m_row, m_col) {
                return Some(strip_future_prefixes(text));
            }
            let moved = crate::refers::shift_a1_references(
                text,
                i64::from(row) - i64::from(m_row),
                i64::from(col) - i64::from(m_col),
            );
            Some(strip_future_prefixes(&moved))
        }
    }
}

/// The cells of a rectangle outside another, as at most four rectangles
/// (KERNEL.49's decision 2): the rows above it, the cells to its left and to
/// its right in the rows the two share, and the rows below it; with the
/// cells the two share, `None` where they share none, the rectangle then
/// whole.
pub(crate) fn cut(piece: A1Range, out: &A1Range) -> (Vec<A1Range>, Option<A1Range>) {
    let Some(shared) = overlap(&piece, out) else {
        return (vec![piece], None);
    };
    let mut left = Vec::with_capacity(4);
    if piece.top < shared.top {
        left.push(A1Range {
            bottom: shared.top - 1,
            ..piece
        });
    }
    if piece.left < shared.left {
        left.push(A1Range {
            top: shared.top,
            bottom: shared.bottom,
            right: shared.left - 1,
            ..piece
        });
    }
    if shared.right < piece.right {
        left.push(A1Range {
            top: shared.top,
            bottom: shared.bottom,
            left: shared.right + 1,
            ..piece
        });
    }
    if shared.bottom < piece.bottom {
        left.push(A1Range {
            top: shared.bottom + 1,
            ..piece
        });
    }
    (left, Some(shared))
}

/// The cells two rectangles share, as a rectangle; `None` where they share
/// none.
pub(crate) fn overlap(a: &A1Range, b: &A1Range) -> Option<A1Range> {
    let shared = A1Range {
        top: a.top.max(b.top),
        left: a.left.max(b.left),
        bottom: a.bottom.min(b.bottom),
        right: a.right.min(b.right),
    };
    (shared.top <= shared.bottom && shared.left <= shared.right).then_some(shared)
}

/// Rectangles that never overlap, each with its tag, kept by their first
/// cell in the class of their height, a power of two (KERNEL.49): a search
/// of an area reads, class by class, the rectangles whose first row stands
/// within the class's height above the area, so that a tall rectangle
/// widens its own class's search and no other's. One map searched as far
/// above as its tallest rectangle reached read every rectangle above an area
/// once one was tall: 20,000 ranges of a row each beside a column of 100,000
/// rows loaded in 4.85 s, where a grid of their cells loaded in 0.16 s. A
/// tag is a `u32`, a fill's index or a draft's record's, one type so that the
/// index is compiled once, where generic over its tag it was compiled for
/// each.
#[derive(Clone, Debug, Default)]
pub(crate) struct Rects {
    /// By class `c`, the rectangles of `2^c` to `2^(c+1) - 1` rows.
    classes: Vec<Class>,
}

/// One class's rectangles, each by its first cell: its last row and column,
/// and its tag.
type Class = BTreeMap<(u32, u32), (u32, u32, u32)>;

/// Two sets of rectangles are equal when they hold the same rectangles with
/// the same tags, however many classes each has kept.
impl PartialEq for Rects {
    fn eq(&self, other: &Rects) -> bool {
        self.iter().eq(other.iter())
    }
}

/// The class of a rectangle's height: the power of two at or below its rows.
fn class(r: &A1Range) -> usize {
    let rows = r.bottom - r.top + 1;
    (u32::BITS - 1 - rows.leading_zeros()) as usize
}

impl Rects {
    /// A rectangle kept with its tag.
    pub(crate) fn insert(&mut self, r: A1Range, tag: u32) {
        let c = class(&r);
        if self.classes.len() <= c {
            self.classes.resize_with(c + 1, BTreeMap::new);
        }
        self.classes[c].insert((r.top, r.left), (r.bottom, r.right, tag));
    }

    /// A rectangle given up; its tag.
    pub(crate) fn remove(&mut self, r: &A1Range) -> Option<u32> {
        self.classes
            .get_mut(class(r))?
            .remove(&(r.top, r.left))
            .map(|(_, _, tag)| tag)
    }

    /// The rectangles meeting an area, whole, each with its tag: class by
    /// class, a class holding none passed by, and in each by first cell.
    pub(crate) fn meeting(&self, area: A1Range) -> impl Iterator<Item = (A1Range, u32)> + '_ {
        let held = self
            .classes
            .iter()
            .enumerate()
            .filter(|(_, m)| !m.is_empty());
        held.flat_map(move |(c, m)| {
            // A rectangle of the class ends at most 2^(c+1) - 2 rows below
            // its first.
            let from = area.top.saturating_sub((2u32 << c) - 2);
            m.range((from, 0)..=(area.bottom, u32::MAX))
                .filter(move |(&(_, left), &(bottom, right, _))| {
                    bottom >= area.top && left <= area.right && right >= area.left
                })
                .map(|(&(top, left), &(bottom, right, tag))| {
                    (
                        A1Range {
                            top,
                            left,
                            bottom,
                            right,
                        },
                        tag,
                    )
                })
        })
    }

    /// The rectangle holding a cell, and its tag.
    pub(crate) fn at(&self, row: u32, col: u32) -> Option<(A1Range, u32)> {
        self.meeting(A1Range {
            top: row,
            left: col,
            bottom: row,
            right: col,
        })
        .next()
    }

    /// Every rectangle, with its tag: class by class, and in each by first
    /// cell.
    pub(crate) fn iter(&self) -> impl Iterator<Item = (A1Range, u32)> + '_ {
        self.classes.iter().flat_map(|m| {
            m.iter().map(|(&(top, left), &(bottom, right, tag))| {
                (
                    A1Range {
                        top,
                        left,
                        bottom,
                        right,
                    },
                    tag,
                )
            })
        })
    }
}

/// The cells of rectangles that never overlap, each with its rectangle's
/// tag, in row-major order (KERNEL.49): a sweep down the rows, the
/// rectangles a row crosses kept by their left column, so that a walk of a
/// sheet's fills or a draft's records visits its cells in the order a
/// sheet's tree held them, at the cost of its cells and its rows and with
/// none of them made. A tag is a `u32`, one type so that the sweep is
/// compiled once, where generic over its tag it was compiled for each.
pub(crate) struct RowMajor {
    /// The rectangles, by first cell.
    rects: Vec<(A1Range, u32)>,
    next: usize,
    /// Those the row crosses, by left column.
    active: Vec<(A1Range, u32)>,
    row: u32,
    k: usize,
    col: u32,
    started: bool,
}

impl RowMajor {
    pub(crate) fn new(mut rects: Vec<(A1Range, u32)>) -> RowMajor {
        rects.sort_unstable_by_key(|(r, _)| (r.top, r.left));
        RowMajor {
            rects,
            next: 0,
            active: Vec::new(),
            row: 0,
            k: 0,
            col: 0,
            started: false,
        }
    }
}

impl Iterator for RowMajor {
    type Item = (u32, u32, u32);

    fn next(&mut self) -> Option<(u32, u32, u32)> {
        loop {
            if let Some(&(r, tag)) = self.active.get(self.k) {
                if self.col <= r.right {
                    let col = self.col;
                    self.col += 1;
                    return Some((self.row, col, tag));
                }
                self.k += 1;
                if let Some((r, _)) = self.active.get(self.k) {
                    self.col = r.left;
                }
                continue;
            }
            // The row is done: the next row a rectangle crosses.
            let mut row = match self.started {
                true => self.row.checked_add(1)?,
                false => 0,
            };
            self.active.retain(|(r, _)| r.bottom >= row);
            if self.active.is_empty() {
                row = row.max(self.rects.get(self.next)?.0.top);
            }
            let mut admitted = false;
            while let Some(&(r, tag)) = self.rects.get(self.next) {
                if r.top > row {
                    break;
                }
                self.active.push((r, tag));
                self.next += 1;
                admitted = true;
            }
            if admitted {
                self.active.sort_unstable_by_key(|(r, _)| r.left);
            }
            self.row = row;
            self.started = true;
            self.k = 0;
            self.col = self.active[0].0.left;
        }
    }
}

/// What stands beside a walk of a sheet's places, met in the walk's order
/// (KERNEL.49): the entries of an ordered map of places, or a sweep of a
/// sheet's formulas, each place's found by stepping past those before it, so
/// that a walk of a window's cells finds each one's formula, format and
/// sentence at the cost of the window's entries and never of a lookup.
pub(crate) struct Beside<I: Iterator> {
    inner: std::iter::Peekable<I>,
}

impl<V, I: Iterator<Item = ((u32, u32), V)>> Beside<I> {
    pub(crate) fn new(inner: I) -> Beside<I> {
        Beside {
            inner: inner.peekable(),
        }
    }

    /// What stands at a place, the walk asked at each place in order.
    pub(crate) fn at(&mut self, at: (u32, u32)) -> Option<V> {
        while self.inner.next_if(|(p, _)| *p < at).is_some() {}
        self.inner.next_if(|(p, _)| *p == at).map(|(_, v)| v)
    }
}

/// The entries of an ordered map of places inside a window, row by row.
pub(crate) fn in_window<'a, V: Copy>(
    map: &'a BTreeMap<(u32, u32), V>,
    w: &A1Range,
) -> impl Iterator<Item = ((u32, u32), V)> + 'a {
    let (left, right) = (w.left, w.right);
    map.range((w.top, w.left)..=(w.bottom, w.right))
        .filter(move |(&(_, col), _)| col >= left && col <= right)
        .map(|(&at, &v)| (at, v))
}

/// A value cell's content as the record reads it: lent where the grid keeps
/// it, owned where it makes it, as a machine makes a twin's cell from its
/// sheet's.
#[derive(Clone, Debug, PartialEq)]
pub(crate) enum Constant<'a> {
    Text(Cow<'a, str>),
    Number(f64),
    Bool(bool),
    Error(Cow<'a, str>),
}

impl Constant<'_> {
    /// The content, owned.
    #[cfg(test)]
    pub(crate) fn content(&self) -> Content {
        match self {
            Constant::Text(t) => Content::Text(t.to_string()),
            Constant::Number(n) => Content::Number(*n),
            Constant::Bool(b) => Content::Bool(*b),
            Constant::Error(e) => Content::Error(e.to_string()),
        }
    }

    /// Whether a written content is this value as a write compares them,
    /// a number by its value (KERNEL.34: a value written over the same is no
    /// write).
    pub(crate) fn is(&self, content: &Content) -> bool {
        match (self, content) {
            (Constant::Text(a), Content::Text(b)) => a == b,
            (Constant::Number(a), Content::Number(b)) => a == b,
            (Constant::Bool(a), Content::Bool(b)) => a == b,
            (Constant::Error(a), Content::Error(b)) => a == b,
            _ => false,
        }
    }

    /// A content that is a value, owned.
    pub(crate) fn owned(content: Content) -> Option<Constant<'static>> {
        Some(match content {
            Content::Text(t) => Constant::Text(Cow::Owned(t)),
            Content::Number(n) => Constant::Number(n),
            Content::Bool(b) => Constant::Bool(b),
            Content::Error(e) => Constant::Error(Cow::Owned(e)),
            _ => return None,
        })
    }
}

/// What a cell holds as the record reads it: a value, or a formula.
#[derive(Clone, Debug, PartialEq)]
pub(crate) enum Holds<'a> {
    Constant(Constant<'a>),
    Formula(Formula<'a>),
}

impl<'a> Holds<'a> {
    /// What a workbook's cell holds.
    pub(crate) fn of(content: &'a Content) -> Holds<'a> {
        match content {
            Content::Text(t) => Holds::Constant(Constant::Text(Cow::Borrowed(t))),
            Content::Number(n) => Holds::Constant(Constant::Number(*n)),
            Content::Bool(b) => Holds::Constant(Constant::Bool(*b)),
            Content::Error(e) => Holds::Constant(Constant::Error(Cow::Borrowed(e))),
            Content::Formula(t) => Holds::Formula(Formula::Own(t, Kind::Plain)),
            Content::DynamicFormula(t) => Holds::Formula(Formula::Own(t, Kind::Dynamic)),
            Content::ArrayFormula(t) => Holds::Formula(Formula::Own(t, Kind::Array)),
            Content::SharedMaster { si, .. } | Content::SharedChild { si } => {
                Holds::Formula(Formula::Fill(*si))
            }
        }
    }
}

/// One cell as the record reads it: what it holds, its format, and the row
/// of the sentence that wrote it.
#[derive(Clone, Debug, PartialEq)]
pub(crate) struct Held<'a> {
    pub(crate) holds: Holds<'a>,
    pub(crate) style: u32,
    pub(crate) sentence: Option<u32>,
}

/// A sheet's settings as the record prints them.
#[derive(Clone, Copy, Debug)]
pub(crate) struct Page<'a> {
    pub(crate) name: &'a str,
    pub(crate) visibility: Visibility,
    pub(crate) gridlines: bool,
    pub(crate) columns: &'a [Column],
    pub(crate) rows: &'a BTreeMap<u32, RowSettings>,
    pub(crate) looks: &'a [(String, u32)],
}

impl<'a> Page<'a> {
    /// A workbook's sheet's.
    pub(crate) fn of(sheet: &'a Sheet) -> Page<'a> {
        Page {
            name: &sheet.name,
            visibility: sheet.visibility,
            gridlines: sheet.gridlines,
            columns: &sheet.columns,
            rows: &sheet.rows,
            looks: &sheet.looks,
        }
    }
}

/// A grid as the record and the plane read it: its sheets' settings, its
/// formats, each sheet's extent, and its cells.
pub(crate) trait Cells: Model {
    /// The format table.
    fn styles(&self) -> &Styles;
    /// A sheet's settings.
    fn page(&self, sheet: usize) -> Page<'_>;
    /// A sheet's used rectangle, as the record names it: a workbook's walked,
    /// a machine's as it keeps it.
    fn extent(&self, sheet: usize) -> Extent;
    /// The cells of a sheet inside a window, row by row.
    fn cells<'a>(&'a self, sheet: usize, window: &A1Range, f: &mut dyn FnMut((u32, u32), Held<'a>));
    /// One cell; `None` for one holding nothing.
    fn cell(&self, sheet: usize, row: u32, col: u32) -> Option<Held<'_>>;
}

impl Model for Workbook {
    fn sheet_count(&self) -> usize {
        self.sheets.len()
    }

    fn sheet_name(&self, sheet: usize) -> Option<&str> {
        self.sheets.get(sheet).map(|s| s.name.as_str())
    }

    fn find_sheet(&self, name: &str) -> Option<usize> {
        Workbook::find_sheet(self, name)
    }

    fn defined_names(&self) -> &[(String, String)] {
        &self.defined_names
    }

    fn formula(&self, (s, row, col): CellId) -> Option<Formula<'_>> {
        let cell = self.sheets.get(s)?.cells.get(&(row, col))?;
        match Holds::of(&cell.content) {
            Holds::Formula(f) => Some(f),
            Holds::Constant(_) => None,
        }
    }

    fn masters(&self, sheet: usize) -> Masters<'_> {
        self.sheets
            .get(sheet)
            .map(shared_masters)
            .unwrap_or_default()
    }

    fn each_formula(&self, sheet: usize, f: &mut dyn FnMut((u32, u32), Formula<'_>)) {
        let Some(s) = self.sheets.get(sheet) else {
            return;
        };
        for (&at, cell) in &s.cells {
            if let Holds::Formula(formula) = Holds::of(&cell.content) {
                f(at, formula);
            }
        }
    }

    fn formulas_in(&self, area: &Area, extent: Extent, f: &mut dyn FnMut(u32, u32)) {
        let Some(s) = self.sheets.get(area.sheet) else {
            return;
        };
        for ((row, col), cell) in cells_in(s, extent, area) {
            if matches!(Holds::of(&cell.content), Holds::Formula(_)) {
                f(row, col);
            }
        }
    }
}

impl Cells for Workbook {
    fn styles(&self) -> &Styles {
        &self.styles
    }

    fn page(&self, sheet: usize) -> Page<'_> {
        Page::of(&self.sheets[sheet])
    }

    fn extent(&self, sheet: usize) -> Extent {
        self.sheets.get(sheet).and_then(Sheet::extent)
    }

    fn cells<'a>(
        &'a self,
        sheet: usize,
        window: &A1Range,
        f: &mut dyn FnMut((u32, u32), Held<'a>),
    ) {
        let Some(s) = self.sheets.get(sheet) else {
            return;
        };
        for (&(row, col), cell) in s
            .cells
            .range((window.top, 0)..=(window.bottom, u32::MAX))
            .filter(|((_, col), _)| *col >= window.left && *col <= window.right)
        {
            f(
                (row, col),
                Held {
                    holds: Holds::of(&cell.content),
                    style: cell.style,
                    sentence: s.sentences.get(&(row, col)).copied(),
                },
            );
        }
    }

    fn cell(&self, sheet: usize, row: u32, col: u32) -> Option<Held<'_>> {
        let s = self.sheets.get(sheet)?;
        let cell = s.cells.get(&(row, col))?;
        Some(Held {
            holds: Holds::of(&cell.content),
            style: cell.style,
            sentence: s.sentences.get(&(row, col)).copied(),
        })
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

    fn cells_of(rects: &[A1Range]) -> Vec<(u32, u32)> {
        let mut out = Vec::new();
        for x in rects {
            for row in x.top..=x.bottom {
                for col in x.left..=x.right {
                    out.push((row, col));
                }
            }
        }
        out.sort_unstable();
        out
    }

    #[test]
    fn a_cut_leaves_the_rows_above_the_cells_beside_and_the_rows_below() {
        // KERNEL.49's decision 2: a write into part of a fill takes the
        // written rectangle out and keeps the rest as at most four.
        let (left, shared) = cut(r(1, 1, 5, 5), &r(2, 3, 3, 4));
        assert_eq!(
            left,
            vec![r(1, 1, 1, 5), r(2, 1, 3, 2), r(2, 5, 3, 5), r(4, 1, 5, 5)]
        );
        assert_eq!(shared, Some(r(2, 3, 3, 4)));
        // Over the whole: nothing left. Apart: the piece whole.
        assert_eq!(
            cut(r(2, 2, 3, 3), &r(1, 1, 9, 9)),
            (vec![], Some(r(2, 2, 3, 3)))
        );
        assert_eq!(
            cut(r(2, 2, 3, 3), &r(5, 5, 6, 6)),
            (vec![r(2, 2, 3, 3)], None)
        );
        // Every overlap, held to the cells it leaves: the pieces never
        // overlap, and hold exactly the cells outside.
        let piece = r(3, 3, 6, 7);
        for top in 1..=8 {
            for bottom in top..=8 {
                for left in 1..=9 {
                    for right in left..=9 {
                        let out = r(top, left, bottom, right);
                        let (pieces, _) = cut(piece, &out);
                        assert!(pieces.len() <= 4);
                        let got = cells_of(&pieces);
                        let mut deduped = got.clone();
                        deduped.dedup();
                        assert_eq!(got, deduped, "{out:?}: no cell twice");
                        let want: Vec<(u32, u32)> = cells_of(&[piece])
                            .into_iter()
                            .filter(|&(row, col)| {
                                !(row >= top && row <= bottom && col >= left && col <= right)
                            })
                            .collect();
                        assert_eq!(got, want, "{out:?}");
                    }
                }
            }
        }
    }

    #[test]
    fn rectangles_are_walked_cell_by_cell_in_row_major_order() {
        // KERNEL.49: the sweep visits the cells of rectangles that never
        // overlap in the order a sheet's tree held them: tall ones beside
        // short ones, gaps of rows between them, one cell alone.
        let rects = vec![
            (r(2, 5, 9, 5), 'a'),
            (r(1, 1, 3, 2), 'b'),
            (r(3, 3, 3, 4), 'c'),
            (r(6, 1, 6, 1), 'd'),
            (r(12, 2, 13, 3), 'e'),
            (r(2, 7, 2, 9), 'f'),
            (r(9, 1, 10, 4), 'g'),
        ];
        let mut want: Vec<(u32, u32, char)> = Vec::new();
        for &(x, tag) in &rects {
            for row in x.top..=x.bottom {
                for col in x.left..=x.right {
                    want.push((row, col, tag));
                }
            }
        }
        want.sort_unstable();
        let tagged = rects.iter().map(|&(x, tag)| (x, u32::from(tag))).collect();
        let got: Vec<(u32, u32, char)> = RowMajor::new(tagged)
            .map(|(row, col, tag)| (row, col, char::from_u32(tag).expect("a tag")))
            .collect();
        assert_eq!(got, want);
        assert_eq!(RowMajor::new(Vec::new()).next(), None);
        // The rows past the last: none, even at the sheet's last row.
        let last = RowMajor::new(vec![(r(1_048_576, 1, 1_048_576, 2), 0)]);
        assert_eq!(last.count(), 2);
    }

    #[test]
    fn rectangles_kept_by_height_are_found_by_every_area_they_meet() {
        // KERNEL.49's third slice: each area's rectangles found class by
        // class are those a walk of every rectangle finds, a column of 5,000
        // rows among rectangles of a row and of a few, before and after
        // some are given up.
        let mut seed: u64 = 49;
        let mut next = |n: u32| {
            seed = seed
                .wrapping_mul(6_364_136_223_846_793_005)
                .wrapping_add(1_442_695_040_888_963_407);
            ((seed >> 33) % u64::from(n)) as u32
        };
        let mut rects = Rects::default();
        let mut all: Vec<(A1Range, u32)> = vec![(r(1, 1, 5_000, 1), 0)];
        rects.insert(all[0].0, 0);
        let mut taken: BTreeMap<(u32, u32), ()> = BTreeMap::new();
        for row in 1..=5_000 {
            taken.insert((row, 1), ());
        }
        while all.len() < 600 {
            let (top, left) = (1 + next(6_000), 2 + next(40));
            let (rows, cols) = match next(4) {
                0 => (1, 1 + next(4)),
                1 => (1 + next(8), 1),
                2 => (1 + next(70), 1 + next(3)),
                _ => (1 + next(3), 1 + next(3)),
            };
            let x = r(top, left, top + rows - 1, left + cols - 1);
            let free = (x.top..=x.bottom)
                .all(|row| (x.left..=x.right).all(|col| !taken.contains_key(&(row, col))));
            if !free {
                continue;
            }
            for row in x.top..=x.bottom {
                for col in x.left..=x.right {
                    taken.insert((row, col), ());
                }
            }
            let k = all.len() as u32;
            rects.insert(x, k);
            all.push((x, k));
        }
        let check = |rects: &Rects, all: &[(A1Range, u32)], next: &mut dyn FnMut(u32) -> u32| {
            for _ in 0..400 {
                let (top, left) = (1 + next(6_100), 1 + next(45));
                let area = r(top, left, top + next(90), left + next(6));
                let mut want: Vec<u32> = all
                    .iter()
                    .filter(|(x, _)| overlap(x, &area).is_some())
                    .map(|&(_, k)| k)
                    .collect();
                let mut got: Vec<u32> = rects.meeting(area).map(|(_, k)| k).collect();
                want.sort_unstable();
                got.sort_unstable();
                assert_eq!(got, want, "{area:?}");
                let (row, col) = (area.top, area.left);
                let holder = all
                    .iter()
                    .find(|(x, _)| overlap(x, &r(row, col, row, col)).is_some())
                    .map(|&(x, k)| (x, k));
                assert_eq!(rects.at(row, col), holder, "({row}, {col})");
            }
            let mut every: Vec<u32> = rects.iter().map(|(_, k)| k).collect();
            every.sort_unstable();
            assert_eq!(every, all.iter().map(|&(_, k)| k).collect::<Vec<_>>());
        };
        check(&rects, &all, &mut next);
        // Every third given up, the column among them.
        let kept: Vec<(A1Range, u32)> = all
            .iter()
            .copied()
            .filter(|&(x, k)| k % 3 != 0 || rects.remove(&x) != Some(k))
            .collect();
        check(&rects, &kept, &mut next);
        assert_eq!(rects.remove(&r(1, 1, 5_000, 1)), None, "given up once");
    }
}
