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
//! called. A draft's and a machine's sheet keep their cells' formats as
//! rectangles ([`CellStyles`], KERNEL.67), so that a `style` row over a range
//! holds its rectangle, as a fill holds its formula once.

use std::borrow::Cow;
use std::collections::{BTreeMap, HashMap};

use crate::calc::graph::{cells_in, Area, CellId, Extent};
use crate::rows::Visibility;
use crate::rule::Rule;
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
    /// Every formula cell of a sheet as the plan places it (KERNEL.49's
    /// decision 7): a fill's cells as rectangles that never overlap, each
    /// with the fill's index, and each formula of its own as its cell, in no
    /// order. A workbook's are found from its cells, a fill's of a row
    /// merged where they stand side by side and rows of the same columns
    /// merged where they follow one another; a machine keeps its own.
    fn each_node(&self, sheet: usize, f: &mut dyn FnMut(A1Range, Formula<'_>)) {
        // Each row's runs of one fill, then the runs of one fill and columns
        // in rows that follow one another made one rectangle.
        let mut runs: Vec<(u32, u32, u32, u32)> = Vec::new();
        self.each_formula(sheet, &mut |(row, col), formula| match formula {
            Formula::Fill(si) => match runs.last_mut() {
                Some(last) if last.0 == si && last.1 == row && last.3 + 1 == col => last.3 = col,
                _ => runs.push((si, row, col, col)),
            },
            Formula::Own(..) => f(
                A1Range {
                    top: row,
                    left: col,
                    bottom: row,
                    right: col,
                },
                formula,
            ),
        });
        runs.sort_unstable_by_key(|&(si, row, left, right)| (si, left, right, row));
        let mut k = 0;
        while k < runs.len() {
            let (si, top, left, right) = runs[k];
            let mut bottom = top;
            k += 1;
            while let Some(&(s, row, l, r)) = runs.get(k) {
                if (s, l, r) != (si, left, right) || row != bottom + 1 {
                    break;
                }
                bottom = row;
                k += 1;
            }
            let rect = A1Range {
                top,
                left,
                bottom,
                right,
            };
            f(rect, Formula::Fill(si));
        }
    }
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

/// The cells of a rectangle outside rectangles that never overlap, as
/// rectangles that never overlap (KERNEL.67): each hole cut from every piece
/// it meets, at most four left of each ([`cut`]), the pieces kept by height
/// so that a hole finds the few it meets.
pub(crate) fn outside(range: A1Range, holes: &[A1Range]) -> Vec<A1Range> {
    if holes.is_empty() {
        return vec![range];
    }
    let mut left = Rects::default();
    left.insert(range, 0);
    for hole in holes {
        let met: Vec<A1Range> = left.meeting(*hole).map(|(piece, _)| piece).collect();
        for piece in met {
            left.remove(&piece);
            for rest in cut(piece, hole).0 {
                left.insert(rest, 0);
            }
        }
    }
    left.iter().map(|(piece, _)| piece).collect()
}

/// What of a rectangle pieces that never overlap leave: [`outside`] over the pieces
/// it meets alone, found by their height (KERNEL.58's *Found* 16), where every piece
/// asked of made a rule of many rectangles cost their square.
pub(crate) fn outside_held(range: A1Range, held: &Rects) -> Vec<A1Range> {
    let holes: Vec<A1Range> = held.meeting(range).map(|(piece, _)| piece).collect();
    outside(range, &holes)
}

/// How many cells rectangles hold, a cell two share once (KERNEL.58's *Found* 16): a
/// sweep down their rows over the spans of columns their edges cut, each span keeping
/// how many rectangles stand over it and how many of its columns any covers, so that no
/// rectangle is cut and none is made, whatever they cross.
pub(crate) fn union_cells(ranges: &[A1Range]) -> u64 {
    match ranges {
        [] => return 0,
        [one] => return one.cells(),
        _ => {}
    }
    let mut cuts: Vec<u32> = ranges.iter().flat_map(|r| [r.left, r.right + 1]).collect();
    cuts.sort_unstable();
    cuts.dedup();
    let at = |col: u32| cuts.binary_search(&col).expect("an edge is a cut");
    let mut edges: Vec<(u32, bool, usize, usize)> = Vec::with_capacity(ranges.len() * 2);
    for r in ranges {
        let (from, to) = (at(r.left), at(r.right + 1));
        edges.push((r.top, true, from, to));
        edges.push((r.bottom + 1, false, from, to));
    }
    edges.sort_unstable();
    let mut spans = Spans {
        cuts: &cuts,
        stood: vec![(0, 0); cuts.len() * 4],
    };
    let (mut cells, mut row) = (0, edges[0].0);
    for (next, opens, from, to) in edges {
        cells += u64::from(spans.stood[1].1) * u64::from(next - row);
        row = next;
        spans.stand(1, 0, cuts.len() - 1, from, to, opens);
    }
    cells
}

/// The tree of a sweep's spans of columns: for each node, how many rectangles stand
/// over it whole, and how many of its columns any rectangle covers.
struct Spans<'c> {
    cuts: &'c [u32],
    stood: Vec<(u32, u32)>,
}

impl Spans<'_> {
    /// A rectangle's columns, the cuts `from` to `to`, opened or closed under the
    /// node of the cuts `low` to `high`.
    fn stand(&mut self, node: usize, low: usize, high: usize, from: usize, to: usize, opens: bool) {
        if from <= low && high <= to {
            if opens {
                self.stood[node].0 += 1;
            } else {
                self.stood[node].0 -= 1;
            }
        } else {
            let middle = (low + high) / 2;
            if from < middle {
                self.stand(node * 2, low, middle, from, to, opens);
            }
            if middle < to {
                self.stand(node * 2 + 1, middle, high, from, to, opens);
            }
        }
        self.stood[node].1 = if self.stood[node].0 > 0 {
            self.cuts[high] - self.cuts[low]
        } else if high - low == 1 {
            0
        } else {
            self.stood[node * 2].1 + self.stood[node * 2 + 1].1
        };
    }
}

/// A sheet's cells' formats (KERNEL.67): each format over rectangles of cells
/// that never overlap, kept as a draft keeps its records, a rectangle of more
/// than one cell among the rectangles of its height ([`Rects`]) and a cell
/// alone by its place, so that a `style` row over a range holds its rectangle
/// and never its cells. A cell's format is 0 where none is held. Nothing is
/// made before a sheet's first format: a sheet with none holds a word, where
/// its two maps would hold six.
#[derive(Clone, Debug, Default)]
pub(crate) struct CellStyles(Option<Box<StyleRects>>);

/// The formats of a sheet that has any: its rectangles of more than one cell,
/// and its cells alone.
#[derive(Clone, Debug, Default)]
struct StyleRects {
    ranges: Rects,
    ones: BTreeMap<(u32, u32), u32>,
}

impl StyleRects {
    fn put(&mut self, range: A1Range, format: u32) {
        if range.cells() == 1 {
            self.ones.insert((range.top, range.left), format);
        } else {
            self.ranges.insert(range, format);
        }
    }
}

impl CellStyles {
    /// A cell's format: its own where it stands alone, else its rectangle's;
    /// 0 where none is held.
    pub(crate) fn at(&self, row: u32, col: u32) -> u32 {
        let Some(s) = &self.0 else {
            return 0;
        };
        if let Some(&format) = s.ones.get(&(row, col)) {
            return format;
        }
        s.ranges.at(row, col).map_or(0, |(_, format)| format)
    }

    /// A rectangle given a format, none of its cells holding one.
    pub(crate) fn put(&mut self, range: A1Range, format: u32) {
        self.0.get_or_insert_with(Box::default).put(range, format);
    }

    /// The rectangles meeting an area, whole, each with its format: those of
    /// more than one cell by height, then the cells alone among the area's
    /// rows.
    pub(crate) fn meeting(&self, area: A1Range) -> impl Iterator<Item = (A1Range, u32)> + '_ {
        self.0.iter().flat_map(move |s| {
            let ones = in_window(&s.ones, &area).map(|((row, col), format)| {
                let cell = A1Range {
                    top: row,
                    left: col,
                    bottom: row,
                    right: col,
                };
                (cell, format)
            });
            s.ranges.meeting(area).chain(ones)
        })
    }

    /// One cell's format taken out, the rest of its rectangle kept as at most
    /// four ([`cut`]).
    pub(crate) fn take(&mut self, row: u32, col: u32) {
        let Some(s) = &mut self.0 else {
            return;
        };
        if s.ones.remove(&(row, col)).is_some() {
            return;
        }
        let Some((piece, format)) = s.ranges.at(row, col) else {
            return;
        };
        s.ranges.remove(&piece);
        let cell = A1Range {
            top: row,
            left: col,
            bottom: row,
            right: col,
        };
        for rest in cut(piece, &cell).0 {
            s.put(rest, format);
        }
    }

    /// Each cell inside a window that holds a format, row by row, with its
    /// format: the rectangles meeting the window cut to it and swept
    /// ([`RowMajor`]) beside the cells alone in it, so that a walk of the
    /// window meets each cell's format beside it ([`Beside`]).
    pub(crate) fn in_window(&self, w: &A1Range) -> impl Iterator<Item = ((u32, u32), u32)> + '_ {
        let w = *w;
        let s = self.0.as_deref();
        let pieces: Vec<(A1Range, u32)> = s.map_or_else(Vec::new, |s| {
            s.ranges
                .meeting(w)
                .filter_map(|(piece, format)| Some((overlap(&piece, &w)?, format)))
                .collect()
        });
        let mut ranges = RowMajor::new(pieces).peekable();
        let mut ones = s
            .into_iter()
            .flat_map(move |s| in_window(&s.ones, &w))
            .peekable();
        std::iter::from_fn(move || {
            let take_range = match (ranges.peek(), ones.peek()) {
                (None, None) => return None,
                (Some(_), None) => true,
                (None, Some(_)) => false,
                (Some(&(row, col, _)), Some(&(at, _))) => (row, col) < at,
            };
            if take_range {
                let (row, col, format) = ranges.next().expect("peeked");
                return Some(((row, col), format));
            }
            ones.next()
        })
    }
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
    /// Its conditional formats, the lowest priority first (KERNEL.58).
    pub(crate) rules: &'a [Rule],
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
            rules: sheet.rules(),
        }
    }
}

/// A grid as the record and the plane read it: its sheets' settings, its
/// formats, each sheet's extent, and its cells.
pub(crate) trait Cells: Model {
    /// Whether a formula's row in the record says its kind (KERNEL.70): a
    /// machine's, whose rows can say it and whose save must keep it. A door's
    /// record stays the file's own text, as the reader prints it.
    fn spells_kinds(&self) -> bool {
        false
    }

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
    /// Whether a sheet is one a record lists no row of: a machine's rule's
    /// answer sheet and its twin (KERNEL.58).
    fn unlisted(&self, _sheet: usize) -> bool {
        false
    }
    /// The priorities of the rules that apply at each cell of a window where
    /// any does, in order, up to the first that stops (KERNEL.58): a machine's
    /// answers, a workbook computing none.
    fn lit_in(&self, _sheet: usize, _window: &A1Range) -> BTreeMap<(u32, u32), Vec<u32>> {
        BTreeMap::new()
    }
    /// [`Cells::lit_in`] at one cell, `None` where none applies.
    fn lit_at(&self, _sheet: usize, _row: u32, _col: u32) -> Option<Vec<u32>> {
        None
    }
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

    /// The same 64-bit linear congruence as the test above, seeded.
    fn dealer(mut seed: u64) -> impl FnMut(u32) -> u32 {
        move |n: u32| {
            seed = seed
                .wrapping_mul(6_364_136_223_846_793_005)
                .wrapping_add(1_442_695_040_888_963_407);
            ((seed >> 33) % u64::from(n)) as u32
        }
    }

    /// A rule's cells, counted by a sweep, are the cells of its rectangles cut against
    /// each other, and cut against what they meet is the same cut.
    #[test]
    fn rectangles_count_by_a_sweep_as_they_cut() {
        let mut seed = 0x5EED_u64;
        let mut next = |n: u32| {
            seed = seed
                .wrapping_mul(6364136223846793005)
                .wrapping_add(1442695040888963407);
            ((seed >> 33) as u32) % n
        };
        for round in 0..400 {
            let many = 1 + round % 24;
            let spread = 1 + round % 30;
            let ranges: Vec<A1Range> = (0..many)
                .map(|_| {
                    let (top, left) = (1 + next(spread), 1 + next(spread));
                    r(top, left, top + next(5), left + next(5))
                })
                .collect();
            let mut held = Rects::default();
            let mut before: Vec<A1Range> = Vec::new();
            let mut cut_cells = 0;
            for &range in &ranges {
                let pieces = outside_held(range, &held);
                let same: u64 = outside(range, &before).iter().map(A1Range::cells).sum();
                let cells: u64 = pieces.iter().map(A1Range::cells).sum();
                assert_eq!(cells, same, "{ranges:?}");
                cut_cells += cells;
                for piece in pieces {
                    assert!(
                        held.meeting(piece).next().is_none(),
                        "{ranges:?}: {piece:?}"
                    );
                    held.insert(piece, 0);
                }
                before.push(range);
            }
            let mut every = std::collections::BTreeSet::new();
            for range in &ranges {
                for row in range.top..=range.bottom {
                    for col in range.left..=range.right {
                        every.insert((row, col));
                    }
                }
            }
            assert_eq!(cut_cells, every.len() as u64, "{ranges:?}");
            assert_eq!(union_cells(&ranges), every.len() as u64, "{ranges:?}");
        }
        assert_eq!(union_cells(&[]), 0);
        // What crosses makes more pieces than cells to count, and counts at once: a
        // grid of strips, and the sheet's whole, by the hundred.
        let strips: Vec<A1Range> = (0..4_000)
            .flat_map(|k| {
                [
                    r(k * 2 + 1, 1, k * 2 + 1, 16_384),
                    r(1, k * 2 + 1, 1_048_576, k * 2 + 1),
                ]
            })
            .collect();
        let crossed = 4_000 * 16_384 + 4_000 * 1_048_576 - 4_000 * 4_000;
        assert_eq!(union_cells(&strips), crossed);
        let wholes = vec![r(1, 1, 1_048_576, 16_384); 200];
        assert_eq!(union_cells(&wholes), 1_048_576 * 16_384);
    }

    #[test]
    fn the_cells_outside_rectangles_are_rectangles_that_never_overlap() {
        // KERNEL.67: a range less the rectangles a sheet's formats already
        // hold, each hole cut from the pieces it meets.
        assert_eq!(outside(r(2, 2, 4, 4), &[]), vec![r(2, 2, 4, 4)]);
        assert_eq!(outside(r(2, 2, 4, 4), &[r(1, 1, 9, 9)]), vec![]);
        let mut next = dealer(67);
        for _ in 0..300 {
            let range = r(1 + next(6), 1 + next(6), 7 + next(6), 7 + next(6));
            // Holes that never overlap, some apart from the range.
            let mut holes: Vec<A1Range> = Vec::new();
            for _ in 0..next(9) {
                let (top, left) = (1 + next(14), 1 + next(14));
                let hole = r(top, left, top + next(4), left + next(4));
                if holes.iter().all(|h| overlap(h, &hole).is_none()) {
                    holes.push(hole);
                }
            }
            let got = outside(range, &holes);
            let cells = cells_of(&got);
            let mut deduped = cells.clone();
            deduped.dedup();
            assert_eq!(cells, deduped, "{range:?} less {holes:?}: no cell twice");
            let want: Vec<(u32, u32)> = cells_of(&[range])
                .into_iter()
                .filter(|&(row, col)| {
                    holes
                        .iter()
                        .all(|h| overlap(h, &r(row, col, row, col)).is_none())
                })
                .collect();
            assert_eq!(cells, want, "{range:?} less {holes:?}");
        }
    }

    #[test]
    fn a_sheet_s_formats_held_as_rectangles_read_as_each_cell_s() {
        // KERNEL.67: formats put over rectangles that never overlap, cells
        // alone among them and format 0 among the formats, each read at every
        // cell, walked over windows row by row and found by the areas they
        // meet as a map of each cell's format holds them, before and after
        // cells are taken out of them; and a sheet with none a word.
        assert_eq!(
            std::mem::size_of::<CellStyles>(),
            std::mem::size_of::<usize>()
        );
        let none = CellStyles::default();
        assert_eq!(none.at(1, 1), 0);
        assert_eq!(none.in_window(&r(1, 1, 9, 9)).count(), 0);
        assert_eq!(none.meeting(r(1, 1, 9, 9)).count(), 0);
        let mut next = dealer(1067);
        for round in 0..40 {
            let mut styles = CellStyles::default();
            let mut each: BTreeMap<(u32, u32), u32> = BTreeMap::new();
            for _ in 0..12 {
                let (top, left) = (1 + next(30), 1 + next(20));
                let x = match next(3) {
                    0 => r(top, left, top, left),
                    1 => r(top, left, top + next(12), left + next(2)),
                    _ => r(top, left, top + next(3), left + next(6)),
                };
                if cells_of(&[x]).iter().any(|at| each.contains_key(at)) {
                    continue;
                }
                let format = next(4);
                styles.put(x, format);
                for at in cells_of(&[x]) {
                    each.insert(at, format);
                }
            }
            let check = |styles: &CellStyles,
                         each: &BTreeMap<(u32, u32), u32>,
                         next: &mut dyn FnMut(u32) -> u32| {
                for row in 1..=46 {
                    for col in 1..=30 {
                        let want = each.get(&(row, col)).copied().unwrap_or(0);
                        assert_eq!(styles.at(row, col), want, "round {round}: ({row}, {col})");
                    }
                }
                for _ in 0..20 {
                    let (top, left) = (1 + next(40), 1 + next(26));
                    let w = r(top, left, top + next(12), left + next(8));
                    let want: Vec<((u32, u32), u32)> = each
                        .iter()
                        .filter(|(&(row, col), _)| {
                            row >= w.top && row <= w.bottom && col >= w.left && col <= w.right
                        })
                        .map(|(&at, &format)| (at, format))
                        .collect();
                    let walked: Vec<((u32, u32), u32)> = styles.in_window(&w).collect();
                    assert_eq!(walked, want, "round {round}: {w:?} walked");
                    let mut met: Vec<((u32, u32), u32)> = Vec::new();
                    for (piece, format) in styles.meeting(w) {
                        let shared = overlap(&piece, &w).expect("a piece meeting the area");
                        for at in cells_of(&[shared]) {
                            met.push((at, format));
                        }
                    }
                    met.sort_unstable();
                    assert_eq!(met, want, "round {round}: {w:?} met");
                }
            };
            check(&styles, &each, &mut next);
            // Cells taken out, inside rectangles, alone and where none is.
            for _ in 0..15 {
                let (row, col) = (1 + next(40), 1 + next(26));
                styles.take(row, col);
                each.remove(&(row, col));
            }
            check(&styles, &each, &mut next);
        }
    }
}
