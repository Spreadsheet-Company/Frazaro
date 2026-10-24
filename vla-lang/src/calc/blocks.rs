//! A plan's order kept as blocks (KERNEL.49's decision 7, its fourth slice,
//! 2026-10-22): the cells of one shape on one sheet at one level of the
//! order, row by row, each block the rectangles its cells fill, so that the
//! order is a list of blocks and no list of its cells. A cell's position is
//! the cells of the blocks before its own and its rank in its own, found
//! from where its block's rows and spans begin; the cell at a position is
//! found the other way round.
//!
//! A block whose cells fill one rectangle is that rectangle and nothing
//! more: a fill at one level, a formula alone. Any other is kept as bands,
//! each a run of rows the same spans of columns cross, and each span's
//! cells before it in its rows, so that the cells of a fill cut by a write
//! into pieces side by side, or of formulas alone of one text side by side,
//! are walked row by row across all of them, as the order walks them.

use crate::sheet::A1Range;

use super::graph::CellId;

/// One block of the order: its shape and its sheet, the position of its
/// first cell, its cells' box, and where its bands begin among every
/// block's; a block whose bands are none (the next block's begin where its
/// own do) holds its box whole.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) struct Block {
    pub(crate) shape: u32,
    pub(crate) sheet: u32,
    pub(crate) start: u32,
    pub(crate) rect: A1Range,
    bands: u32,
}

/// A run of a block's rows that the same spans cross: its first and last
/// row, the rank in its block of its first cell, its cells in a row, and
/// where its spans begin among every band's.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct Band {
    top: u32,
    bottom: u32,
    start: u32,
    width: u32,
    spans: u32,
}

/// The columns of a band's rows that its block's cells fill, and how many
/// of a row's cells stand before them.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct Span {
    left: u32,
    right: u32,
    before: u32,
}

/// The order as blocks: each block by the position of its first cell, and
/// every block's bands and spans in one array each, a block's or a band's
/// ending where the next one's begin.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub(crate) struct Blocks {
    blocks: Vec<Block>,
    bands: Vec<Band>,
    spans: Vec<Span>,
    len: u32,
    /// Where the order holds at most [`LIST`] cells, each cell by its
    /// position, its sheet, row, column and shape; else empty.
    list: Vec<[u32; 4]>,
}

/// The most cells an order keeps as a list beside its blocks, 16 bytes a
/// cell: Alonzo's Sand, 189,215, and every bench program, and neither large
/// board, whose bytes are the item's.
const LIST: usize = 1 << 18;

#[cfg(test)]
thread_local! {
    static LIST_AT: std::cell::Cell<usize> = const { std::cell::Cell::new(0) };
}

/// The tests' own bound for the list, none unless a test sets it, so that
/// what the tests hold is the finder; the bound before.
#[cfg(test)]
pub(crate) fn list_at(most: usize) -> usize {
    LIST_AT.with(|c| c.replace(most))
}

#[cfg(not(test))]
fn list_most() -> usize {
    LIST
}

#[cfg(test)]
fn list_most() -> usize {
    let _ = LIST;
    LIST_AT.with(|c| c.get())
}

impl Blocks {
    /// Blocks made from their cells' rectangles, each block's given in the
    /// order of the order, as its shape, its sheet and rectangles that
    /// never overlap, in any order; a block given no cell is passed by.
    #[cfg(test)]
    pub(crate) fn of(groups: impl IntoIterator<Item = (u32, u32, Vec<A1Range>)>) -> Blocks {
        let mut b = Blocks::default();
        for (shape, sheet, rects) in groups {
            b.push(shape, sheet, rects);
        }
        b.shrink();
        b
    }

    /// The blocks made, their arrays held to what they hold, and an order
    /// of at most [`LIST`] cells listed.
    pub(crate) fn shrink(&mut self) {
        self.blocks.shrink_to_fit();
        self.bands.shrink_to_fit();
        self.spans.shrink_to_fit();
        self.list = Vec::new();
        if self.len() <= list_most() {
            let mut list = Vec::with_capacity(self.len());
            for (k, b) in self.blocks.iter().enumerate() {
                let from = b.start as usize;
                for (sheet, row, col) in self.cells(from, from + self.count(k)) {
                    list.push([sheet as u32, row, col, b.shape]);
                }
            }
            self.list = list;
        }
    }

    /// One block added after the others.
    pub(crate) fn push(&mut self, shape: u32, sheet: u32, mut rects: Vec<A1Range>) {
        rects.retain(|r| r.top <= r.bottom && r.left <= r.right);
        let Some(first) = rects.first().copied() else {
            return;
        };
        let start = self.len;
        let bands = self.bands.len() as u32;
        let cells: u64 = rects.iter().map(A1Range::cells).sum();
        let rect = rects.iter().fold(first, |a, r| A1Range {
            top: a.top.min(r.top),
            left: a.left.min(r.left),
            bottom: a.bottom.max(r.bottom),
            right: a.right.max(r.right),
        });
        if rect.cells() != cells {
            self.band(&mut rects, bands as usize);
        }
        self.blocks.push(Block {
            shape,
            sheet,
            start,
            rect,
            bands,
        });
        self.len = start + cells as u32;
    }

    /// A block's rectangles kept as bands, its first at `first`: a sweep
    /// down the rows they cross, each row's spans merged where they meet,
    /// and rows of the same spans one band.
    fn band(&mut self, rects: &mut [A1Range], first_band: usize) {
        rects.sort_unstable_by_key(|r| (r.top, r.left));
        let mut breaks: Vec<u32> = Vec::with_capacity(rects.len() * 2);
        for r in rects.iter() {
            breaks.push(r.top);
            breaks.push(r.bottom + 1);
        }
        breaks.sort_unstable();
        breaks.dedup();
        let mut next = 0;
        let mut active: Vec<A1Range> = Vec::new();
        let mut rank = 0u32;
        let mut row_spans: Vec<(u32, u32)> = Vec::new();
        for w in breaks.windows(2) {
            let (top, end) = (w[0], w[1]);
            active.retain(|r| r.bottom >= top);
            while let Some(r) = rects.get(next).filter(|r| r.top <= top) {
                active.push(*r);
                next += 1;
            }
            if active.is_empty() {
                continue;
            }
            active.sort_unstable_by_key(|r| r.left);
            row_spans.clear();
            for r in &active {
                match row_spans.last_mut() {
                    Some(last) if last.1 + 1 == r.left => last.1 = r.right,
                    _ => row_spans.push((r.left, r.right)),
                }
            }
            let bottom = end - 1;
            // Rows of the same spans as the band above, and below it, grow it.
            let first = self.bands.last().map_or(0, |b| b.spans as usize);
            let same = self.bands.len() > first_band
                && self.bands.last().is_some_and(|b| {
                    b.bottom + 1 == top
                        && self.spans.len() - first == row_spans.len()
                        && self.spans[first..]
                            .iter()
                            .zip(&row_spans)
                            .all(|(s, &(l, r))| s.left == l && s.right == r)
                });
            if same {
                let b = self.bands.last_mut().expect("a band");
                rank += b.width * (bottom - b.bottom);
                b.bottom = bottom;
                continue;
            }
            let spans = self.spans.len() as u32;
            let mut before = 0;
            for &(left, right) in &row_spans {
                self.spans.push(Span {
                    left,
                    right,
                    before,
                });
                before += right - left + 1;
            }
            self.bands.push(Band {
                top,
                bottom,
                start: rank,
                width: before,
                spans,
            });
            rank += before * (bottom - top + 1);
        }
    }

    /// How many cells the list holds (the tests').
    #[cfg(test)]
    pub(crate) fn listed(&self) -> usize {
        self.list.len()
    }

    /// How many cells the order holds.
    pub(crate) fn len(&self) -> usize {
        self.len as usize
    }

    /// Every block, in the order's order.
    pub(crate) fn blocks(&self) -> &[Block] {
        &self.blocks
    }

    /// How many cells a block holds, by its index.
    pub(crate) fn count(&self, k: usize) -> usize {
        let end = self.blocks.get(k + 1).map_or(self.len, |b| b.start);
        (end - self.blocks[k].start) as usize
    }

    /// A block's bands, by its index; none for a block that holds its box
    /// whole.
    fn bands_of(&self, k: usize) -> &[Band] {
        let from = self.blocks[k].bands as usize;
        let to = self
            .blocks
            .get(k + 1)
            .map_or(self.bands.len(), |b| b.bands as usize);
        &self.bands[from..to]
    }

    /// A band's spans, by the band's index among every band's.
    fn spans_of(&self, band: usize) -> &[Span] {
        let from = self.bands[band].spans as usize;
        let to = self
            .bands
            .get(band + 1)
            .map_or(self.spans.len(), |b| b.spans as usize);
        &self.spans[from..to]
    }

    /// The block holding a position, by its index.
    pub(crate) fn block_at(&self, p: usize) -> usize {
        self.blocks.partition_point(|b| b.start as usize <= p) - 1
    }

    /// The cell at a position of the order, and its shape.
    pub(crate) fn cell(&self, p: usize) -> (CellId, usize) {
        let k = self.block_at(p);
        (self.cell_in(k, p), self.blocks[k].shape as usize)
    }

    /// The cell at a position of the order, in the block holding it.
    pub(crate) fn cell_in(&self, k: usize, p: usize) -> CellId {
        let b = &self.blocks[k];
        let r = p as u32 - b.start;
        let bands = self.bands_of(k);
        let (row, col) = if bands.is_empty() {
            let w = b.rect.right - b.rect.left + 1;
            (b.rect.top + r / w, b.rect.left + r % w)
        } else {
            let i = bands.partition_point(|band| band.start <= r) - 1;
            let band = &bands[i];
            let rr = r - band.start;
            let k_in_row = rr % band.width;
            let spans = self.spans_of(b.bands as usize + i);
            let span = &spans[spans.partition_point(|s| s.before <= k_in_row) - 1];
            (
                band.top + rr / band.width,
                span.left + (k_in_row - span.before),
            )
        };
        (b.sheet as usize, row, col)
    }

    /// A cell's rank in a block, the cells of the block before it row by
    /// row; `None` for a cell the block does not hold.
    pub(crate) fn rank(&self, k: usize, row: u32, col: u32) -> Option<u32> {
        let b = &self.blocks[k];
        let r = &b.rect;
        if row < r.top || row > r.bottom || col < r.left || col > r.right {
            return None;
        }
        let bands = self.bands_of(k);
        if bands.is_empty() {
            return Some((row - r.top) * (r.right - r.left + 1) + (col - r.left));
        }
        let i = bands
            .partition_point(|band| band.top <= row)
            .checked_sub(1)?;
        let band = &bands[i];
        if row > band.bottom {
            return None;
        }
        let spans = self.spans_of(b.bands as usize + i);
        let span = &spans[spans.partition_point(|s| s.left <= col).checked_sub(1)?];
        (col <= span.right)
            .then(|| band.start + (row - band.top) * band.width + span.before + (col - span.left))
    }

    /// A cell's position in a block, `None` for a cell it does not hold.
    pub(crate) fn position_in(&self, k: usize, row: u32, col: u32) -> Option<usize> {
        self.rank(k, row, col)
            .map(|r| (self.blocks[k].start + r) as usize)
    }

    /// The cells of a block inside a rectangle, row by row, each with its
    /// position, to `f`.
    pub(crate) fn each_in<F: FnMut(usize, u32, u32)>(&self, k: usize, area: &A1Range, f: &mut F) {
        let b = &self.blocks[k];
        let Some(cut) = crate::grid::overlap(&b.rect, area) else {
            return;
        };
        let bands = self.bands_of(k);
        if bands.is_empty() {
            let w = b.rect.right - b.rect.left + 1;
            for row in cut.top..=cut.bottom {
                let base = b.start + (row - b.rect.top) * w;
                for col in cut.left..=cut.right {
                    f((base + (col - b.rect.left)) as usize, row, col);
                }
            }
            return;
        }
        let from = bands.partition_point(|band| band.bottom < cut.top);
        for (i, band) in bands.iter().enumerate().skip(from) {
            if band.top > cut.bottom {
                break;
            }
            let spans = self.spans_of(b.bands as usize + i);
            let first = spans.partition_point(|s| s.right < cut.left);
            for row in band.top.max(cut.top)..=band.bottom.min(cut.bottom) {
                let base = b.start + band.start + (row - band.top) * band.width;
                for s in &spans[first..] {
                    if s.left > cut.right {
                        break;
                    }
                    for col in s.left.max(cut.left)..=s.right.min(cut.right) {
                        f((base + s.before + (col - s.left)) as usize, row, col);
                    }
                }
            }
        }
    }

    /// A block's cells as rectangles, row by row: its box whole, or each
    /// band's spans.
    pub(crate) fn rects_of(&self, k: usize) -> impl Iterator<Item = A1Range> + '_ {
        let b = self.blocks[k];
        let bands = self.bands_of(k);
        let whole = bands.is_empty().then_some(b.rect);
        let banded = bands.iter().enumerate().flat_map(move |(i, band)| {
            self.spans_of(b.bands as usize + i)
                .iter()
                .map(move |s| A1Range {
                    top: band.top,
                    left: s.left,
                    bottom: band.bottom,
                    right: s.right,
                })
        });
        whole.into_iter().chain(banded)
    }

    /// The cells of the order from one position to another pushed to `out`,
    /// in order, a block's rows and spans at a time: a run's cells laid into
    /// its buffer for an evaluation at once, where each was found by a step.
    pub(crate) fn extend_cells(&self, from: usize, to: usize, out: &mut Vec<CellId>) {
        let to = to.min(self.len());
        if from >= to {
            return;
        }
        let mut k = self.block_at(from);
        let mut p = from;
        while p < to {
            let b = &self.blocks[k];
            let sheet = b.sheet as usize;
            let start = b.start as usize;
            let end = (start + self.count(k)).min(to);
            let bands = self.bands_of(k);
            if bands.is_empty() {
                let (left, right) = (b.rect.left, b.rect.right);
                let r = (p - start) as u32;
                let w = right - left + 1;
                if w == 1 {
                    // A column: its rows at once.
                    let top = b.rect.top + r;
                    let n = (end - p) as u32;
                    out.extend((top..top + n).map(|row| (sheet, row, left)));
                } else {
                    // Row by row, the first from where the stretch begins.
                    let (mut row, mut col) = (b.rect.top + r / w, left + r % w);
                    let mut q = p;
                    while q < end {
                        let n = ((right - col + 1) as usize).min(end - q);
                        out.extend((col..col + n as u32).map(|c| (sheet, row, c)));
                        q += n;
                        row += 1;
                        col = left;
                    }
                }
            } else {
                for (i, band) in bands.iter().enumerate() {
                    let band_start = start + band.start as usize;
                    let band_end =
                        band_start + (band.width * (band.bottom - band.top + 1)) as usize;
                    if band_end <= p || band_start >= end {
                        continue;
                    }
                    let spans = self.spans_of(b.bands as usize + i);
                    for row in band.top..=band.bottom {
                        let row_start = band_start + ((row - band.top) * band.width) as usize;
                        if row_start + band.width as usize <= p || row_start >= end {
                            continue;
                        }
                        for s in spans {
                            let s_start = row_start + s.before as usize;
                            let s_end = s_start + (s.right - s.left + 1) as usize;
                            let (a, z) = (s_start.max(p), s_end.min(end));
                            if a < z {
                                let left = s.left + (a - s_start) as u32;
                                out.extend(
                                    (left..left + (z - a) as u32).map(|col| (sheet, row, col)),
                                );
                            }
                        }
                    }
                }
            }
            p = end;
            k += 1;
        }
    }

    /// The cells of the order from one position to another, in order.
    pub(crate) fn cells(&self, from: usize, to: usize) -> Cells<'_> {
        let to = to.min(self.len());
        let mut at = At::default();
        if from < to {
            at.seek(self, from);
        }
        Cells {
            blocks: self,
            at,
            to,
            from,
        }
    }
}

/// Where a walk of the order stands: the position, its block, the band and
/// the span it is in (`NONE` for a block that holds its box whole), and its
/// cell, so that the next cell is found by a step along the span, the band
/// and the blocks, where it was found by a division and two searches.
#[derive(Clone, Copy, Debug)]
struct At {
    p: usize,
    k: usize,
    band: u32,
    span: u32,
    row: u32,
    col: u32,
    /// The position after the block's last cell, the block's first and last
    /// column, its sheet and its shape, kept so that a step and a cell read
    /// no block.
    end: usize,
    left: u32,
    right: u32,
    sheet: u32,
    shape: u32,
}

impl Default for At {
    fn default() -> At {
        At {
            p: usize::MAX,
            k: 0,
            band: NONE,
            span: NONE,
            row: 0,
            col: 0,
            end: 0,
            left: 0,
            right: 0,
            sheet: 0,
            shape: 0,
        }
    }
}

/// No band: a block holding its box whole.
const NONE: u32 = u32::MAX;

impl At {
    /// The walk put at a position of the order, its block searched for.
    fn seek(&mut self, blocks: &Blocks, p: usize) {
        let k = blocks.block_at(p);
        self.go(blocks, k, p);
    }

    /// The walk put at a position of a block.
    fn go(&mut self, blocks: &Blocks, k: usize, p: usize) {
        let b = &blocks.blocks[k];
        // A block of one cell, a chain's or a netlist's every one: its cell
        // and its end read off it alone.
        if b.rect.top == b.rect.bottom && b.rect.left == b.rect.right {
            self.p = p;
            self.k = k;
            self.end = p + 1;
            (self.left, self.right) = (b.rect.left, b.rect.right);
            (self.sheet, self.shape) = (b.sheet, b.shape);
            (self.band, self.span) = (NONE, NONE);
            (self.row, self.col) = (b.rect.top, b.rect.left);
            return;
        }
        let r = p as u32 - b.start;
        // The next block's entry once: where this one ends, and where its
        // bands end.
        let (end, bands_end) = match blocks.blocks.get(k + 1) {
            Some(n) => (n.start, n.bands),
            None => (blocks.len, blocks.bands.len() as u32),
        };
        self.p = p;
        self.k = k;
        self.end = end as usize;
        (self.left, self.right) = (b.rect.left, b.rect.right);
        (self.sheet, self.shape) = (b.sheet, b.shape);
        if b.bands == bands_end {
            (self.band, self.span) = (NONE, NONE);
            // A block's first cell, a chain's every one, with no division.
            if r == 0 {
                (self.row, self.col) = (b.rect.top, b.rect.left);
                return;
            }
            let w = b.rect.right - b.rect.left + 1;
            (self.row, self.col) = (b.rect.top + r / w, b.rect.left + r % w);
            return;
        }
        let bands = &blocks.bands[b.bands as usize..bands_end as usize];
        let i = bands.partition_point(|band| band.start <= r) - 1;
        let band = &bands[i];
        let rr = r - band.start;
        let in_row = rr % band.width;
        let first = band.spans as usize;
        let spans = blocks.spans_of(b.bands as usize + i);
        let j = spans.partition_point(|s| s.before <= in_row) - 1;
        self.band = b.bands + i as u32;
        self.span = (first + j) as u32;
        self.row = band.top + rr / band.width;
        self.col = spans[j].left + (in_row - spans[j].before);
    }

    /// The walk stepped to the next position, `p + 1`, which the order holds.
    #[inline]
    fn step(&mut self, blocks: &Blocks) {
        let next = self.p + 1;
        if next >= self.end {
            self.go(blocks, self.k + 1, next);
            return;
        }
        self.p = next;
        if self.band == NONE {
            if self.col < self.right {
                self.col += 1;
            } else {
                self.col = self.left;
                self.row += 1;
            }
            return;
        }
        let span = &blocks.spans[self.span as usize];
        if self.col < span.right {
            self.col += 1;
            return;
        }
        let band = &blocks.bands[self.band as usize];
        let spans_end = blocks
            .bands
            .get(self.band as usize + 1)
            .map_or(blocks.spans.len(), |n| n.spans as usize);
        if (self.span as usize) + 1 < spans_end {
            self.span += 1;
            self.col = blocks.spans[self.span as usize].left;
            return;
        }
        if self.row < band.bottom {
            self.row += 1;
            self.span = band.spans;
            self.col = blocks.spans[self.span as usize].left;
            return;
        }
        // The band is done: the next band of the block, which holds the
        // position, the block not being done.
        self.band += 1;
        let band = &blocks.bands[self.band as usize];
        self.row = band.top;
        self.span = band.spans;
        self.col = blocks.spans[self.span as usize].left;
    }

    #[inline]
    fn cell(&self) -> CellId {
        (self.sheet as usize, self.row, self.col)
    }
}

/// The cells of a stretch of the order, block by block (`Blocks::cells`),
/// each by a step from the one before.
pub(crate) struct Cells<'a> {
    blocks: &'a Blocks,
    at: At,
    to: usize,
    from: usize,
}

impl Iterator for Cells<'_> {
    type Item = CellId;

    fn next(&mut self) -> Option<CellId> {
        if self.from >= self.to {
            return None;
        }
        if self.from > self.at.p {
            self.at.step(self.blocks);
        }
        self.from += 1;
        Some(self.at.cell())
    }
}

/// A finder of the cells at positions asked in rising order, which keeps
/// where the last one stood, so that a walk of a frame's marks finds each
/// cell by a step along the order where it is near, and by a search where it
/// is not. Made again for each frame, the order fixed while one is in
/// progress.
#[derive(Clone, Copy, Debug, Default)]
pub(crate) struct Locator {
    at: At,
}

/// The farthest a finder steps to a position before it searches for it.
const NEAR: usize = 32;

impl Locator {
    /// The cell at a position, and its shape.
    #[inline]
    pub(crate) fn cell(&mut self, blocks: &Blocks, p: usize) -> (CellId, usize) {
        // An order small enough to be listed: read off its list.
        if let Some(&[sheet, row, col, shape]) = blocks.list.get(p) {
            return ((sheet as usize, row, col), shape as usize);
        }
        let at = &mut self.at;
        if at.p != usize::MAX && p == at.p + 1 {
            // The next position of a box, the walk's commonest, in place.
            if p < at.end {
                if at.band == NONE {
                    at.p = p;
                    if at.col < at.right {
                        at.col += 1;
                    } else {
                        at.col = at.left;
                        at.row += 1;
                    }
                    return (at.cell(), at.shape as usize);
                }
            } else if let Some(b) = blocks.blocks.get(at.k + 1) {
                // The next block's first cell where the block is that cell
                // alone, a chain's and a netlist's every one, from it alone.
                if b.rect.top == b.rect.bottom && b.rect.left == b.rect.right {
                    at.p = p;
                    at.k += 1;
                    at.end = p + 1;
                    (at.left, at.right) = (b.rect.left, b.rect.right);
                    (at.sheet, at.shape) = (b.sheet, b.shape);
                    (at.band, at.span) = (NONE, NONE);
                    (at.row, at.col) = (b.rect.top, b.rect.left);
                    return (
                        (b.sheet as usize, b.rect.top, b.rect.left),
                        b.shape as usize,
                    );
                }
            }
        }
        self.far(blocks, p)
    }

    /// The cell at a position the walk does not stand next to: where it
    /// stands, again; ahead in its block, along a box's row or down its rows
    /// by a division, along a span, or by a search of the block's bands;
    /// else stepped to where it is near and searched for where it is not.
    #[inline(never)]
    fn far(&mut self, blocks: &Blocks, p: usize) -> (CellId, usize) {
        let at = &mut self.at;
        let placed = at.p != usize::MAX && at.k < blocks.blocks.len();
        if placed && p == at.p {
            return (at.cell(), at.shape as usize);
        }
        if placed && p > at.p && p < at.end {
            let d = (p - at.p) as u32;
            if at.band == NONE {
                let w = at.right - at.left + 1;
                let off = at.col - at.left + d;
                if off < w {
                    at.col = at.left + off;
                } else {
                    at.row += off / w;
                    at.col = at.left + off % w;
                }
                at.p = p;
            } else if at.col + d <= blocks.spans[at.span as usize].right {
                at.col += d;
                at.p = p;
            } else {
                at.go(blocks, at.k, p);
            }
            return (at.cell(), at.shape as usize);
        }
        if !placed || p < at.p || p - at.p > NEAR {
            at.seek(blocks, p);
        } else {
            while at.p < p {
                at.step(blocks);
            }
        }
        (at.cell(), at.shape as usize)
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

    /// Rectangles that never overlap, seeded: cut from a grid of small
    /// cells and runs, some side by side in the same rows.
    fn seeded(seed: &mut u64, n: usize) -> Vec<A1Range> {
        let mut next = |below: u32| {
            *seed = seed
                .wrapping_mul(6_364_136_223_846_793_005)
                .wrapping_add(1_442_695_040_888_963_407);
            ((*seed >> 33) % u64::from(below)) as u32
        };
        let mut taken: Vec<A1Range> = Vec::new();
        for _ in 0..n * 4 {
            if taken.len() == n {
                break;
            }
            let (top, left) = (1 + next(30), 1 + next(30));
            let x = r(top, left, top + next(4), left + next(5));
            if taken.iter().all(|t| crate::grid::overlap(t, &x).is_none()) {
                taken.push(x);
            }
        }
        taken
    }

    fn cells_of(rects: &[A1Range]) -> Vec<(u32, u32)> {
        let mut cells: Vec<(u32, u32)> = rects
            .iter()
            .flat_map(|x| {
                (x.top..=x.bottom)
                    .flat_map(move |row| (x.left..=x.right).map(move |col| (row, col)))
            })
            .collect();
        cells.sort_unstable();
        cells
    }

    #[test]
    fn a_finder_crosses_blocks_of_one_cell_in_place_and_finds_every_position() {
        // Blocks of one cell, a chain's and a netlist's, between blocks of a
        // box and of bands, on two sheets: the finder asked every position
        // in turn, and between them the same one again, one behind, two
        // ahead or one far, each answer the cell the order holds there and
        // its shape, so that the crossing kept in place agrees with a search.
        let mut seed = 0x5EED_0049_0C05;
        for round in 0..120 {
            let mut groups = Vec::new();
            let mut want: Vec<(CellId, usize)> = Vec::new();
            for k in 0..(4 + round % 9) {
                let sheet = (k % 2) as u32;
                let rects = match (k + round) % 4 {
                    // A cell alone, most often.
                    0 | 1 => vec![r(
                        1 + k as u32,
                        1 + (round % 5) as u32,
                        1 + k as u32,
                        1 + (round % 5) as u32,
                    )],
                    // A box.
                    2 => vec![r(40 + k as u32, 2, 41 + k as u32, 4)],
                    // Rectangles cut into bands.
                    _ => seeded(&mut seed, 2 + round % 3),
                };
                for (row, col) in cells_of(&rects) {
                    want.push(((sheet as usize, row, col), k));
                }
                groups.push((k as u32, sheet, rects));
            }
            let b = Blocks::of(groups);
            assert_eq!(b.len(), want.len(), "round {round}");
            let mut at = Locator::default();
            let mut next = |below: usize| {
                seed = seed
                    .wrapping_mul(6_364_136_223_846_793_005)
                    .wrapping_add(1_442_695_040_888_963_407);
                ((seed >> 33) % below as u64) as usize
            };
            for p in 0..want.len() {
                assert_eq!(at.cell(&b, p), want[p], "round {round}, position {p}");
                let q = match next(6) {
                    0 => p,
                    1 => p.saturating_sub(1),
                    2 => (p + 2).min(want.len() - 1),
                    3 => next(want.len()),
                    _ => continue,
                };
                assert_eq!(at.cell(&b, q), want[q], "round {round}, {q} after {p}");
                assert_eq!(
                    at.cell(&b, p),
                    want[p],
                    "round {round}, {p} again after {q}"
                );
            }
        }
    }

    #[test]
    fn a_block_walks_its_cells_row_by_row_and_finds_each_by_rank() {
        // Blocks of seeded rectangles, held to their cells sorted row by row:
        // the cell at each position, each cell's position, a rectangle's
        // cells and their positions, a stretch walked, and positions asked in
        // rising order with gaps.
        let mut seed = 0x5EED_0049_0004;
        for round in 0..200 {
            let mut groups = Vec::new();
            let mut want: Vec<(CellId, usize)> = Vec::new();
            for k in 0..(1 + round % 5) {
                let rects = seeded(&mut seed, 1 + (round % 7) * (k + 1));
                let sheet = (k % 2) as u32;
                for (row, col) in cells_of(&rects) {
                    want.push(((sheet as usize, row, col), k));
                }
                groups.push((k as u32, sheet, rects));
            }
            let b = Blocks::of(groups.clone());
            assert_eq!(b.len(), want.len(), "round {round}");
            for (p, &(cell, shape)) in want.iter().enumerate() {
                assert_eq!(b.cell(p), (cell, shape), "round {round}, position {p}");
                let k = b.block_at(p);
                assert_eq!(b.position_in(k, cell.1, cell.2), Some(p), "round {round}");
            }
            let walked: Vec<CellId> = b.cells(0, b.len()).collect();
            assert_eq!(walked, want.iter().map(|w| w.0).collect::<Vec<_>>());
            let from = want.len() / 3;
            let part: Vec<CellId> = b.cells(from, from + 7).collect();
            assert_eq!(
                part,
                want[from..(from + 7).min(want.len())]
                    .iter()
                    .map(|w| w.0)
                    .collect::<Vec<_>>()
            );
            // The same stretches laid out a block at a time, every stretch of
            // up to nine from every position.
            for start in 0..=want.len() {
                for n in [0, 1, 2, 5, 9] {
                    let mut laid = Vec::new();
                    b.extend_cells(start, start + n, &mut laid);
                    let walked: Vec<CellId> = b.cells(start, start + n).collect();
                    assert_eq!(laid, walked, "round {round}, {start} and {n} more");
                }
            }
            let mut whole = Vec::new();
            b.extend_cells(0, b.len(), &mut whole);
            assert_eq!(whole, want.iter().map(|w| w.0).collect::<Vec<_>>());
            let mut at = Locator::default();
            for p in (0..want.len()).step_by(1 + round % 3) {
                assert_eq!(at.cell(&b, p), want[p], "round {round}");
            }
            for k in 0..b.blocks().len() {
                // Every cell of the box, held or not.
                let x = b.blocks()[k].rect;
                for row in x.top.saturating_sub(1)..=x.bottom + 1 {
                    for col in x.left.saturating_sub(1)..=x.right + 1 {
                        let held = want.iter().position(|&((s, rr, cc), sh)| {
                            s == b.blocks()[k].sheet as usize
                                && rr == row
                                && cc == col
                                && sh == b.blocks()[k].shape as usize
                        });
                        assert_eq!(
                            b.position_in(k, row, col),
                            held,
                            "round {round} {row} {col}"
                        );
                    }
                }
                let area = r(
                    x.top + 1,
                    x.left,
                    x.bottom,
                    x.right.saturating_sub(1).max(x.left),
                );
                let mut got = Vec::new();
                b.each_in(k, &area, &mut |p, row, col| got.push((p, row, col)));
                let wanted: Vec<(usize, u32, u32)> = want
                    .iter()
                    .enumerate()
                    .filter(|(p, _)| b.block_at(*p) == k)
                    .filter(|(_, ((_, row, col), _))| {
                        *row >= area.top
                            && *row <= area.bottom
                            && *col >= area.left
                            && *col <= area.right
                    })
                    .map(|(p, ((_, row, col), _))| (p, *row, *col))
                    .collect();
                assert_eq!(got, wanted, "round {round} block {k}");
                let mut from_rects = cells_of(&b.rects_of(k).collect::<Vec<_>>());
                from_rects.sort_unstable();
                let mut of_block: Vec<(u32, u32)> = want
                    .iter()
                    .enumerate()
                    .filter(|(p, _)| b.block_at(*p) == k)
                    .map(|(_, ((_, row, col), _))| (*row, *col))
                    .collect();
                of_block.sort_unstable();
                assert_eq!(from_rects, of_block, "round {round} block {k}");
            }
        }
    }

    #[test]
    fn cells_filling_a_rectangle_are_kept_as_the_rectangle() {
        // A board's cells given one by one are one rectangle, no band kept;
        // two pieces side by side are one band of two spans, and the rows
        // above and below them a band each.
        let cells: Vec<A1Range> = (1..=4)
            .flat_map(|row| (1..=5).map(move |col| r(row, col, row, col)))
            .collect();
        let b = Blocks::of([(0, 0, cells)]);
        assert_eq!(b.blocks()[0].rect, r(1, 1, 4, 5));
        assert!(b.bands.is_empty() && b.spans.is_empty());
        let cut = vec![r(1, 1, 1, 6), r(2, 1, 3, 2), r(2, 5, 3, 6), r(4, 1, 4, 6)];
        let b = Blocks::of([(0, 0, cut)]);
        assert_eq!(b.bands.len(), 3);
        assert_eq!(b.spans.len(), 4);
        assert_eq!(b.cell(6), ((0, 2, 1), 0));
        assert_eq!(b.cell(8), ((0, 2, 5), 0));
        assert_eq!(b.cell(10), ((0, 3, 1), 0));
        assert_eq!(b.cell(14), ((0, 4, 1), 0));
        assert_eq!(b.rank(0, 2, 3), None);
        assert_eq!(Blocks::of(Vec::<(u32, u32, Vec<A1Range>)>::new()).len(), 0);
        assert_eq!(Blocks::of([(0, 0, Vec::new())]).blocks().len(), 0);
    }
}
