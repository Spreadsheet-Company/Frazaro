//! The plan by fills and levels (KERNEL.49's decision 7, its fourth slice,
//! 2026-10-22): a plan's formula cells placed as nodes, each a fill's piece
//! or a formula alone, and its order found over the nodes and kept as
//! blocks (`super::blocks`), where it was found over every cell and kept as
//! a list of them, every cell's place hashed and a grid of every cell's
//! position beside it.
//!
//! Each of a node's shape's references is swept over the node's offsets
//! once: the rectangle it reads at some offset found, the nodes of that
//! sheet meeting it found by their rectangles, and for each the rectangle of
//! the node's cells that read it, a dimension at a time, each offset at
//! which the reference's numbers, moved, stay on the sheet and meet the
//! other's (`reach::Dim::meeting`); an edge is a pair of nodes and the
//! reference. A reference every `$` holds reads one area at every cell, and
//! the area is one vertex of its own, a hub, which every node reading it
//! reads, its nodes found once, as the plan of cells read a shape's held
//! references once (KERNEL.33, KERNEL.8): twenty thousand formulas alone of
//! one text, each reading forty thousand cells, made an edge for each
//! pair. Tarjan's components are found over the nodes, and each node's
//! cells' levels from its edges and the levels of the cells they read: a
//! level one past the deepest of the cells each rectangle of its cells
//! reads, as a map of rectangles. A node that reads itself, stands in a
//! component of more than one node, reads a defined name, a span of sheets
//! or a text the parse does not read, or whose map would split into many
//! rectangles, is planned cell by cell as every cell was: each cell's reads
//! found at it, Tarjan's components over the cells, a chain's order and a
//! cycle's cells to the cell.
//!
//! The order is by level, then shape, then cell, sheet first, as it was: a
//! block is a shape's cells on a sheet at a level, row by row, its runs and
//! its levels with it, so that the order is the order of cells it was, cell
//! for cell. A run crosses from one block into the next only where the two
//! are one shape on one sheet, at a level's end; it breaks at the first of
//! the later block's cells that reads a cell of it, found by the rectangles
//! of the edges into it, no cell of one level reading another.

use std::collections::HashMap;

use crate::fx::FxHashMap;
use crate::grid::{overlap, Formula, Model, Rects};
use crate::refers;
use crate::sheet::A1Range;

use super::blocks::Blocks;
use super::formula::{Read, RefExpr};
use super::graph::{areas_found, dependencies_in, tarjan_flat, CellId, Extent, Names};
use super::reach::{followed, Dim};
use super::shape::{hub_at, Shape, RUN};

/// No position: a cell on a cycle, which the order does not hold, or one not
/// yet planned.
pub(crate) const NONE: u32 = u32::MAX;

/// A cell's level marked as a cycle's, its level kept for the cells that
/// read it.
const CYCLE: u32 = 1 << 31;

/// The most rectangles one node's cells are raised by before it is planned
/// cell by cell, the most cells of the grid its map is found on, and the
/// most rectangles the map may hold.
const RAISES: usize = 256;
const GRID: usize = 4096;
const PIECES: usize = 64;

/// A formula alone as the plan holds it: its shape, found when it is
/// placed, and its position in the order, found when the order is; `NONE`
/// before that and for a cell on a cycle.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) struct Alone {
    pub(crate) shape: u32,
    pub(crate) pos: u32,
}

/// A run of the order: its cells are the positions from `start` to `end`,
/// of one shape.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) struct Run {
    pub(crate) start: usize,
    pub(crate) end: usize,
    pub(crate) shape: usize,
}

/// A fill's piece as the order holds it: its sheet, its rectangle and its
/// shape, and where its cells' positions are found, its rectangles of one
/// level each with its block (`parts[at..at + len]`), or, for a piece
/// planned cell by cell, each cell's position, row by row (`dense[at..]`).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct Node {
    sheet: u32,
    rect: A1Range,
    shape: u32,
    at: u32,
    len: u32,
    cells: bool,
}

/// Where the plan's formula cells stand in the order: each formula alone by
/// its place, and each fill's piece by its rectangle, so that a cell's
/// position, and a shape's cells inside a rectangle, are found from the
/// pieces and the blocks.
#[derive(Clone, Debug, Default, PartialEq)]
pub(crate) struct Located {
    /// By sheet, the formulas alone, kept from one plan to the next.
    pub(crate) alone: Vec<FxHashMap<(u32, u32), Alone>>,
    nodes: Vec<Node>,
    parts: Vec<(A1Range, u32)>,
    dense: Vec<u32>,
    /// Each shape's pieces, `shape_nodes[shape_firsts[s]..shape_firsts[s + 1]]`.
    shape_firsts: Vec<u32>,
    shape_nodes: Vec<u32>,
    /// By sheet, the pieces by rectangle.
    by_sheet: Vec<Rects>,
    /// By sheet, its formulas alone's shapes and positions over their box,
    /// where they stand close enough together (`None` elsewhere), as the plan
    /// of cells kept every formula cell's position, so that a shape's cells
    /// inside a rectangle and a cell's position are read by place.
    alone_grid: Vec<Option<AloneGrid>>,
    /// Each shape's blocks, `shape_blocks[block_firsts[s]..block_firsts[s + 1]]`.
    block_firsts: Vec<u32>,
    shape_blocks: Vec<u32>,
    /// By shape, how its cells inside a rectangle are found: through its
    /// blocks, few ([`FEW`]); through its pieces, a fill's; or by place, a
    /// shape of formulas alone's.
    how: Vec<u8>,
}

/// A shape's cells found through its blocks, few.
const BY_BLOCKS: u8 = 0;
/// A shape's cells found through its pieces, a fill's of many blocks.
const BY_PIECES: u8 = 1;
/// A shape's cells found by place, formulas alone of many blocks.
const BY_PLACE: u8 = 2;

/// The formulas alone of one sheet over their box: each slot a formula
/// alone's shape and position, `ABSENT` for a slot holding none.
#[derive(Clone, Debug, PartialEq)]
struct AloneGrid {
    top: u32,
    left: u32,
    width: u32,
    height: u32,
    at: Vec<Alone>,
}

/// A slot of the grid holding no formula alone.
const ABSENT: Alone = Alone {
    shape: NONE,
    pos: NONE,
};

impl AloneGrid {
    /// The formulas alone of a sheet in a grid over their box, where it is no
    /// more than four slots a formula and 64.
    fn of(alone: &FxHashMap<(u32, u32), Alone>) -> Option<AloneGrid> {
        let mut cells = alone.keys();
        let &(r0, c0) = cells.next()?;
        let (mut top, mut left, mut bottom, mut right) = (r0, c0, r0, c0);
        for &(row, col) in cells {
            (top, left) = (top.min(row), left.min(col));
            (bottom, right) = (bottom.max(row), right.max(col));
        }
        let (width, height) = (right - left + 1, bottom - top + 1);
        let slots = u64::from(width) * u64::from(height);
        if slots > 4 * alone.len() as u64 + 64 {
            return None;
        }
        let mut at = vec![ABSENT; slots as usize];
        for (&(row, col), &a) in alone {
            at[((row - top) * width + (col - left)) as usize] = a;
        }
        Some(AloneGrid {
            top,
            left,
            width,
            height,
            at,
        })
    }

    /// The formula alone at a place; `None` for a place holding none.
    #[inline]
    fn at(&self, row: u32, col: u32) -> Option<Alone> {
        let (r, c) = (row.checked_sub(self.top)?, col.checked_sub(self.left)?);
        if r >= self.height || c >= self.width {
            return None;
        }
        let a = self.at[(r * self.width + c) as usize];
        (a.shape != NONE).then_some(a)
    }
}

/// The most blocks a shape's cells inside a rectangle are found through,
/// each block's rectangles asked; a shape of more is asked through its
/// pieces or its places.
const FEW: usize = 8;

fn holds(r: &A1Range, row: u32, col: u32) -> bool {
    r.top <= row && row <= r.bottom && r.left <= col && col <= r.right
}

impl Located {
    /// A formula alone's entry.
    fn alone_at(&self, (s, row, col): CellId) -> Option<&Alone> {
        self.alone.get(s)?.get(&(row, col))
    }

    /// The piece holding a cell, by its index.
    fn node_at(&self, sheet: usize, row: u32, col: u32) -> Option<usize> {
        self.by_sheet
            .get(sheet)?
            .at(row, col)
            .map(|(_, n)| n as usize)
    }

    /// Whether a cell holds a placed formula.
    pub(crate) fn holds(&self, id: CellId) -> bool {
        self.alone_at(id).is_some() || self.node_at(id.0, id.1, id.2).is_some()
    }

    /// A placed formula cell's shape.
    pub(crate) fn shape_of(&self, id: CellId) -> Option<usize> {
        if let Some(a) = self.alone_at(id) {
            return Some(a.shape as usize);
        }
        let n = self.node_at(id.0, id.1, id.2)?;
        Some(self.nodes[n].shape as usize)
    }

    /// A formula cell's position in the order; `None` for a cell holding
    /// none and for one on a cycle.
    pub(crate) fn position(&self, blocks: &Blocks, id: CellId) -> Option<usize> {
        let alone = match self.alone_grid.get(id.0).and_then(Option::as_ref) {
            Some(g) => g.at(id.1, id.2),
            None => self.alone_at(id).copied(),
        };
        if let Some(a) = alone {
            return (a.pos != NONE).then_some(a.pos as usize);
        }
        let n = self.node_at(id.0, id.1, id.2)?;
        self.position_in(blocks, n, id.1, id.2)
    }

    fn position_in(&self, blocks: &Blocks, n: usize, row: u32, col: u32) -> Option<usize> {
        let node = &self.nodes[n];
        let at = node.at as usize;
        if node.cells {
            let w = node.rect.right - node.rect.left + 1;
            let k = (row - node.rect.top) * w + (col - node.rect.left);
            let p = self.dense[at + k as usize];
            return (p != NONE).then_some(p as usize);
        }
        let parts = &self.parts[at..at + node.len as usize];
        let &(_, b) = parts.iter().find(|(r, _)| holds(r, row, col))?;
        blocks.position_in(b as usize, row, col)
    }

    /// The positions of a shape's cells on a sheet inside a rectangle, to
    /// `f`, each once: a fill's through its pieces and their blocks, a
    /// formula alone's by its place.
    pub(crate) fn each_in<F: FnMut(usize)>(
        &self,
        blocks: &Blocks,
        shape: usize,
        sheet: usize,
        area: &A1Range,
        f: &mut F,
    ) {
        let Some(&how) = self.how.get(shape) else {
            return;
        };
        if how == BY_BLOCKS {
            let (a, b) = (self.block_firsts[shape], self.block_firsts[shape + 1]);
            for &k in &self.shape_blocks[a as usize..b as usize] {
                if blocks.blocks()[k as usize].sheet as usize == sheet {
                    blocks.each_in(k as usize, area, &mut |p, _, _| f(p));
                }
            }
            return;
        }
        if how == BY_PLACE {
            self.alone_in(shape, sheet, area, f);
            return;
        }
        let (from, to) = (
            self.shape_firsts[shape] as usize,
            self.shape_firsts[shape + 1] as usize,
        );
        for &n in &self.shape_nodes[from..to] {
            let node = &self.nodes[n as usize];
            if node.sheet as usize != sheet {
                continue;
            }
            let Some(cut) = overlap(&node.rect, area) else {
                continue;
            };
            let at = node.at as usize;
            if node.cells {
                let w = node.rect.right - node.rect.left + 1;
                for row in cut.top..=cut.bottom {
                    let base = at + ((row - node.rect.top) * w) as usize;
                    for col in cut.left..=cut.right {
                        let p = self.dense[base + (col - node.rect.left) as usize];
                        if p != NONE {
                            f(p as usize);
                        }
                    }
                }
                continue;
            }
            for &(r, b) in &self.parts[at..at + node.len as usize] {
                if let Some(c) = overlap(&r, &cut) {
                    blocks.each_in(b as usize, &c, &mut |p, _, _| f(p));
                }
            }
        }
    }

    /// The positions of a shape of formulas alone's cells on a sheet inside
    /// a rectangle, to `f`, by place: the sheet's grid of them where it keeps
    /// one, else its map.
    fn alone_in<F: FnMut(usize)>(&self, shape: usize, sheet: usize, area: &A1Range, f: &mut F) {
        if let Some(g) = self.alone_grid.get(sheet).and_then(Option::as_ref) {
            let cut = A1Range {
                top: area.top.max(g.top),
                left: area.left.max(g.left),
                bottom: area.bottom.min(g.top + g.height - 1),
                right: area.right.min(g.left + g.width - 1),
            };
            for row in cut.top..=cut.bottom {
                for col in cut.left..=cut.right {
                    let a = g.at[((row - g.top) * g.width + (col - g.left)) as usize];
                    if a.shape as usize == shape && a.pos != NONE {
                        f(a.pos as usize);
                    }
                }
            }
            return;
        }
        let Some(alone) = self.alone.get(sheet).filter(|m| !m.is_empty()) else {
            return;
        };
        // The cells of the area each asked, or the formulas alone each
        // tried, whichever are fewer.
        if area.cells() <= alone.len() as u64 {
            for row in area.top..=area.bottom {
                for col in area.left..=area.right {
                    if let Some(a) = alone.get(&(row, col)) {
                        if a.shape as usize == shape && a.pos != NONE {
                            f(a.pos as usize);
                        }
                    }
                }
            }
        } else {
            for (&(row, col), a) in alone {
                if a.shape as usize == shape && a.pos != NONE && holds(area, row, col) {
                    f(a.pos as usize);
                }
            }
        }
    }

    /// Every piece, its sheet, its rectangle and its shape, in no order:
    /// the census's.
    pub(crate) fn pieces(&self) -> impl Iterator<Item = (usize, A1Range, usize)> + '_ {
        self.nodes
            .iter()
            .map(|n| (n.sheet as usize, n.rect, n.shape as usize))
    }
}

/// What the plan found: the order as blocks, where its cells stand, its
/// cycles, runs and levels, and each shape's stretches.
#[derive(Debug, Default)]
pub(crate) struct Planned {
    pub(crate) blocks: Blocks,
    pub(crate) located: Located,
    pub(crate) cycles: Vec<Vec<CellId>>,
    pub(crate) runs: Vec<Run>,
    pub(crate) level_ends: Vec<u32>,
    pub(crate) stretches: Vec<(u32, u32)>,
    pub(crate) stretch_firsts: Vec<u32>,
    pub(crate) formulas: usize,
    /// How many hubs the cells planned cell by cell read through, and how
    /// many cells were planned so.
    pub(crate) hubs: usize,
    pub(crate) by_cell: usize,
}

/// A node while the plan is made: a fill's piece or a formula alone.
#[derive(Clone, Copy, Debug)]
struct PNode {
    sheet: u32,
    rect: A1Range,
    shape: u32,
    alone: bool,
}

/// One reference of a shape as its nodes read it: the sheet it reads (the
/// node's own where `None`) and its two dimensions.
#[derive(Clone, Copy, Debug)]
struct NodeRead {
    on: Option<usize>,
    rows: Dim,
    cols: Dim,
}

/// A node's cells' levels: unknown until its component is planned; a map
/// of rectangles, each of one level (`pieces[at..at + len]` of the
/// planner's); or each cell's, row by row, a cell on a cycle marked
/// (`cells[at..]`), with its rectangles of one level and row
/// (`pieces[spans..spans + len]`). Two arrays hold every node's, where a
/// node's own made the lookups' peak 11.8 MB higher, 140,000 formulas alone
/// each allocating its map.
#[derive(Clone, Copy, Debug)]
enum Levels {
    Unknown,
    Pieces { at: u32, len: u32 },
    Cells { at: u32, spans: u32, len: u32 },
}

/// What one cell planned cell by cell reads: an area, one every cell of the
/// grid reads alike by its index among them, or a cell a name's or a text's
/// scan found.
#[derive(Clone, Copy, Debug)]
enum Item {
    Area(u32, A1Range),
    Held(u32),
    Cell(CellId),
}

/// An area every cell of a component reads alike: its cells of the
/// component, the deepest level it reaches outside it, and its hub.
type Group = (Vec<u32>, Option<u32>, Option<u32>);

/// The nodes by sheet, while the plan is made: the pieces by rectangle and
/// the formulas alone by place, sorted, as the nodes are.
#[derive(Default)]
struct Index {
    fills: Vec<Rects>,
    alone: Vec<Vec<((u32, u32), u32)>>,
}

impl Index {
    fn meeting(&self, sheet: usize, area: &A1Range, f: &mut dyn FnMut(u32)) {
        if let Some(r) = self.fills.get(sheet) {
            for (_, n) in r.meeting(*area) {
                f(n);
            }
        }
        let Some(m) = self.alone.get(sheet).filter(|m| !m.is_empty()) else {
            return;
        };
        if area.bottom - area.top < 64 {
            for row in area.top..=area.bottom {
                let from = m.partition_point(|e| e.0 < (row, area.left));
                for &(_, n) in m[from..].iter().take_while(|e| e.0 <= (row, area.right)) {
                    f(n);
                }
            }
        } else {
            let from = m.partition_point(|e| e.0 < (area.top, 0));
            for &((_, col), n) in m[from..].iter().take_while(|e| e.0 .0 <= area.bottom) {
                if area.left <= col && col <= area.right {
                    f(n);
                }
            }
        }
    }

    fn at(&self, (s, row, col): CellId) -> Option<u32> {
        if let Some(m) = self.alone.get(s) {
            if let Ok(k) = m.binary_search_by_key(&(row, col), |e| e.0) {
                return Some(m[k].1);
            }
        }
        self.fills.get(s)?.at(row, col).map(|(_, n)| n)
    }
}

/// The areas every cell reads alike, by index, and each one's index.
#[derive(Default)]
struct HeldAreas {
    areas: Vec<(u32, A1Range)>,
    by: FxHashMap<(u32, u32, u32, u32, u32), u32>,
}

impl HeldAreas {
    fn id(&mut self, sheet: usize, r: A1Range) -> u32 {
        let key = (sheet as u32, r.top, r.left, r.bottom, r.right);
        let next = self.areas.len() as u32;
        *self.by.entry(key).or_insert_with(|| {
            self.areas.push((sheet as u32, r));
            next
        })
    }
}

/// The cells of a node reading any cell of a rectangle of the sheet a read
/// reads, through that read: the rectangle of the offsets at which it does,
/// a dimension at a time, inside the node; `None` where none does.
fn cells_reading(
    rect: &A1Range,
    anchor: (u32, u32),
    read: &NodeRead,
    z: &A1Range,
) -> Option<A1Range> {
    let (a_row, a_col) = (i64::from(anchor.0), i64::from(anchor.1));
    let (r0, r1) = read.rows.meeting(z.top, z.bottom)?;
    let (c0, c1) = read.cols.meeting(z.left, z.right)?;
    let top = (i64::from(rect.top) - a_row).max(r0);
    let bottom = (i64::from(rect.bottom) - a_row).min(r1);
    let left = (i64::from(rect.left) - a_col).max(c0);
    let right = (i64::from(rect.right) - a_col).min(c1);
    (top <= bottom && left <= right).then(|| A1Range {
        top: (a_row + top) as u32,
        left: (a_col + left) as u32,
        bottom: (a_row + bottom) as u32,
        right: (a_col + right) as u32,
    })
}

/// The rectangle a read reads at some offset of a node's cells, on the
/// sheet; `None` where it reads none.
fn sweep(rect: &A1Range, anchor: (u32, u32), read: &NodeRead) -> Option<A1Range> {
    let (a_row, a_col) = (i64::from(anchor.0), i64::from(anchor.1));
    let rows = (i64::from(rect.top) - a_row, i64::from(rect.bottom) - a_row);
    let cols = (i64::from(rect.left) - a_col, i64::from(rect.right) - a_col);
    let (top, bottom) = read.rows.reach(rows)?;
    let (left, right) = read.cols.reach(cols)?;
    Some(A1Range {
        top,
        left,
        bottom,
        right,
    })
}

/// A shape's references as its nodes read them, each once; `None` where
/// none moves within the sheet its placement found or its own.
fn node_reads<M: Model + ?Sized>(model: &M, shape: &Shape) -> Vec<NodeRead> {
    let Ok(e) = &shape.expr else {
        return Vec::new();
    };
    let mut each: Vec<&RefExpr> = Vec::new();
    e.each_read(&mut |r| {
        if let Read::Ref(re) = r {
            if !each.contains(&re) {
                each.push(re);
            }
        }
    });
    let mut reads = Vec::with_capacity(each.len());
    for re in each {
        let on = match (re.found, &re.sheet) {
            (Some((first, _)), _) => Some(usize::from(first)),
            (None, None) => None,
            (None, Some(name)) => match model.find_sheet(name) {
                Some(s) => Some(s),
                None => continue,
            },
        };
        let (Some(rows), Some(cols)) = (Dim::rows(re), Dim::cols(re)) else {
            continue;
        };
        reads.push(NodeRead { on, rows, cols });
    }
    reads
}

/// The level map of a node's rectangle raised by rectangles, each to its
/// level at least, on the grid their edges cut it into, merged into
/// rectangles of one level; `None` where the grid or the map grows past
/// its bound.
fn level_map(rect: A1Range, raises: &[(A1Range, u32)]) -> Option<Vec<(A1Range, u32)>> {
    if raises.is_empty() {
        return Some(vec![(rect, 0)]);
    }
    if raises.iter().all(|(r, _)| *r == rect) {
        let most = raises.iter().map(|&(_, l)| l).max().unwrap_or(0);
        return Some(vec![(rect, most)]);
    }
    let mut rows = vec![rect.top, rect.bottom + 1];
    let mut cols = vec![rect.left, rect.right + 1];
    for (r, _) in raises {
        rows.extend([r.top, r.bottom + 1]);
        cols.extend([r.left, r.right + 1]);
    }
    rows.sort_unstable();
    rows.dedup();
    cols.sort_unstable();
    cols.dedup();
    let (nr, nc) = (rows.len() - 1, cols.len() - 1);
    if nr * nc > GRID {
        return None;
    }
    let mut grid = vec![0u32; nr * nc];
    for &(r, level) in raises {
        let i0 = rows.binary_search(&r.top).expect("a row edge");
        let i1 = rows.binary_search(&(r.bottom + 1)).expect("a row edge");
        let j0 = cols.binary_search(&r.left).expect("a column edge");
        let j1 = cols.binary_search(&(r.right + 1)).expect("a column edge");
        for i in i0..i1 {
            for cell in &mut grid[i * nc + j0..i * nc + j1] {
                *cell = (*cell).max(level);
            }
        }
    }
    // Each grid row's runs of one level, carried down while the rows below
    // hold the same runs.
    let mut pieces: Vec<(A1Range, u32)> = Vec::new();
    let mut open: Vec<(usize, usize, u32, usize)> = Vec::new();
    let mut runs: Vec<(usize, usize, u32)> = Vec::new();
    let close = |pieces: &mut Vec<(A1Range, u32)>,
                 (j0, j1, level, i0): (usize, usize, u32, usize),
                 i: usize| {
        pieces.push((
            A1Range {
                top: rows[i0],
                left: cols[j0],
                bottom: rows[i] - 1,
                right: cols[j1] - 1,
            },
            level,
        ));
    };
    for i in 0..nr {
        runs.clear();
        let row = &grid[i * nc..(i + 1) * nc];
        let mut j = 0;
        while j < nc {
            let mut k = j + 1;
            while k < nc && row[k] == row[j] {
                k += 1;
            }
            runs.push((j, k, row[j]));
            j = k;
        }
        let same = open.len() == runs.len()
            && open
                .iter()
                .zip(&runs)
                .all(|(o, r)| (o.0, o.1, o.2) == (r.0, r.1, r.2));
        if !same {
            for o in open.drain(..) {
                close(&mut pieces, o, i);
            }
            open.extend(runs.iter().map(|&(j0, j1, level)| (j0, j1, level, i)));
        }
        if pieces.len() + open.len() > PIECES {
            return None;
        }
    }
    for o in open.drain(..) {
        close(&mut pieces, o, nr);
    }
    Some(pieces)
}

/// The plan over a grid's nodes: every formula cell placed (its shape found
/// by its fill's index, or as a formula alone by its place) planned, and
/// each formula alone's position written beside its shape.
#[allow(clippy::too_many_arguments)]
pub(crate) fn plan<M: Model + ?Sized>(
    model: &M,
    shapes: &[Shape],
    by_shared: &HashMap<(usize, u32), usize>,
    unread: &FxHashMap<CellId, String>,
    names: &Names,
    extents: &[Extent],
    mut alone: Vec<FxHashMap<(u32, u32), Alone>>,
) -> Planned {
    let mut nodes: Vec<PNode> = Vec::new();
    for si in 0..model.sheet_count() {
        model.each_node(si, &mut |rect, formula| {
            let (shape, is_alone) = match formula {
                Formula::Fill(index) => (by_shared.get(&(si, index)).copied(), false),
                Formula::Own(..) => (
                    alone
                        .get(si)
                        .and_then(|m| m.get(&(rect.top, rect.left)))
                        .map(|a| a.shape as usize),
                    true,
                ),
            };
            if let Some(s) = shape {
                nodes.push(PNode {
                    sheet: si as u32,
                    rect,
                    shape: s as u32,
                    alone: is_alone,
                });
            }
        });
    }
    nodes.sort_unstable_by_key(|n| (n.sheet, n.rect.top, n.rect.left));
    let mut p = Planner {
        model,
        shapes,
        unread,
        names,
        extents,
        index: Index::default(),
        reads: Vec::new(),
        edge_firsts: Vec::new(),
        edges: Vec::new(),
        levels: Vec::new(),
        piece_store: Vec::new(),
        cell_store: Vec::new(),
        hub_firsts: Vec::new(),
        hub_nodes: Vec::new(),
        hub_levels: Vec::new(),
        held: HeldAreas::default(),
        cycles: Vec::new(),
        hubs: 0,
        read_lists: std::cell::RefCell::new(vec![None; shapes.len()]),
        nodes,
    };
    p.index();
    p.reads = shapes
        .iter()
        .map(|s| followed(s).then(|| node_reads(model, s)))
        .collect();
    let (offsets, targets) = p.edges();
    let (members, starts) = tarjan_flat(&offsets, &targets);
    drop((offsets, targets));
    let n = p.nodes.len();
    p.levels = vec![Levels::Unknown; n];
    p.hub_levels = vec![None; p.hub_firsts.len().saturating_sub(1)];
    let mut mine: Vec<u32> = Vec::new();
    for c in starts.windows(2) {
        let comp = &members[c[0]..c[1]];
        mine.clear();
        mine.extend(comp.iter().copied().filter(|&v| (v as usize) < n));
        let planned = match comp {
            [x] if (*x as usize) < n && !p.reads_itself(*x as usize) => p
                .block_levels(*x as usize)
                .map(|map| p.set_pieces(*x as usize, &map)),
            _ => None,
        };
        if planned.is_none() && !mine.is_empty() {
            p.cell_levels(&mine);
        }
        for &h in comp.iter().filter(|&&v| (v as usize) >= n) {
            let h = h as usize - n;
            p.hub_levels[h] = p.hub_level(h);
        }
    }
    let (mut planned, order) = p.order(&mut alone);
    planned.located.alone = alone;
    p.runs(&mut planned, &order);
    let mut cycles = std::mem::take(&mut p.cycles);
    cycles.sort();
    planned.cycles = cycles;
    planned.hubs = p.hubs;
    planned.by_cell = p
        .levels
        .iter()
        .zip(&p.nodes)
        .filter(|(l, _)| matches!(l, Levels::Cells { .. }))
        .map(|(_, n)| n.rect.cells() as usize)
        .sum();
    planned
}

/// What the order's making keeps for the runs: every node's cells of one
/// level by level, shape and sheet, each with its block, and the blocks by
/// level, shape and sheet.
struct Made {
    parts: Vec<(u32, u32, A1Range)>,
    block_of_part: Vec<u32>,
    by_key: FxHashMap<(u32, u32, u32), u32>,
}

struct Planner<'a, M: Model + ?Sized> {
    model: &'a M,
    shapes: &'a [Shape],
    unread: &'a FxHashMap<CellId, String>,
    names: &'a Names,
    extents: &'a [Extent],
    nodes: Vec<PNode>,
    index: Index,
    /// By shape, its references as its nodes read them; `None` for a shape
    /// planned cell by cell.
    reads: Vec<Option<Vec<NodeRead>>>,
    /// Each node's edges, `edges[edge_firsts[x]..edge_firsts[x + 1]]`: the
    /// node read, or an area read alike by its hub's vertex, numbered after
    /// the nodes; and the read, `NONE` for one a cell's scan found.
    edge_firsts: Vec<u32>,
    edges: Vec<(u32, u32)>,
    levels: Vec<Levels>,
    /// Every node's rectangles of one level, and every cell's level of the
    /// nodes planned cell by cell ([`Levels`]).
    piece_store: Vec<(A1Range, u32)>,
    cell_store: Vec<u32>,
    /// Each hub's nodes, `hub_nodes[hub_firsts[h]..hub_firsts[h + 1]]`, and
    /// the deepest level of their cells inside its area, `None` where it
    /// holds no formula or is not yet planned.
    hub_firsts: Vec<u32>,
    hub_nodes: Vec<u32>,
    hub_levels: Vec<Option<u32>>,
    held: HeldAreas,
    cycles: Vec<Vec<CellId>>,
    hubs: usize,
    /// By shape, what its parse reads in the order written, listed the
    /// first time a cell of it is planned cell by cell.
    read_lists: std::cell::RefCell<Vec<Option<std::rc::Rc<[Read<'a>]>>>>,
}

impl<'a, M: Model + ?Sized> Planner<'a, M> {
    fn index(&mut self) {
        let sheets = self.model.sheet_count();
        self.index.fills = vec![Rects::default(); sheets];
        self.index.alone = vec![Default::default(); sheets];
        for (k, n) in self.nodes.iter().enumerate() {
            let s = n.sheet as usize;
            if n.alone {
                self.index.alone[s].push(((n.rect.top, n.rect.left), k as u32));
            } else {
                self.index.fills[s].insert(n.rect, k as u32);
            }
        }
    }

    fn anchor(&self, x: usize) -> (u32, u32) {
        let a = self.shapes[self.nodes[x].shape as usize].anchor;
        (a.1, a.2)
    }

    /// Every node's edges, and the nodes each reads, each once, as
    /// `tarjan_flat` reads them: a followed shape's node's found by its
    /// rectangles, any other's from its cells' reads, found here and kept.
    fn edges(&mut self) -> (Vec<usize>, Vec<u32>) {
        // A hub's vertex is numbered after every node's.
        let n = self.nodes.len() as u32;
        let mut offsets = Vec::with_capacity(self.nodes.len() + 1);
        let mut targets: Vec<u32> = Vec::new();
        offsets.push(0);
        self.edge_firsts.push(0);
        let mut met: Vec<u32> = Vec::new();
        for x in 0..self.nodes.len() {
            let node = self.nodes[x];
            let shape = node.shape as usize;
            met.clear();
            if let Some(reads) = &self.reads[shape] {
                let anchor = self.anchor(x);
                for (k, read) in reads.iter().enumerate() {
                    let on = read.on.unwrap_or(node.sheet as usize);
                    if read.rows.fixed() && read.cols.fixed() {
                        // One area at every cell: its hub's.
                        let ((top, bottom), (left, right)) =
                            (read.rows.bounds(), read.cols.bounds());
                        let area = A1Range {
                            top,
                            left,
                            bottom,
                            right,
                        };
                        let h = self.held.id(on, area);
                        self.edges.push((n + h, k as u32));
                        continue;
                    }
                    let Some(area) = sweep(&node.rect, anchor, read) else {
                        continue;
                    };
                    let nodes = &self.nodes;
                    let edges = &mut self.edges;
                    self.index.meeting(on, &area, &mut |y| {
                        if cells_reading(&node.rect, anchor, read, &nodes[y as usize].rect)
                            .is_some()
                        {
                            edges.push((y, k as u32));
                        }
                    });
                }
            } else {
                let mut held = std::mem::take(&mut self.held);
                let mut items: Vec<Item> = Vec::new();
                let from = *self.edge_firsts.last().expect("a first") as usize;
                for row in node.rect.top..=node.rect.bottom {
                    for col in node.rect.left..=node.rect.right {
                        items.clear();
                        self.reads_of(x, (row, col), Some(&mut held), &mut items);
                        for item in &items {
                            match *item {
                                Item::Area(t, a) => {
                                    let edges = &mut self.edges;
                                    self.index
                                        .meeting(t as usize, &a, &mut |y| edges.push((y, NONE)));
                                }
                                Item::Held(h) => self.edges.push((n + h, NONE)),
                                Item::Cell(c) => {
                                    if let Some(y) = self.index.at(c) {
                                        self.edges.push((y, NONE));
                                    }
                                }
                            }
                        }
                    }
                    // The node's edges each once as they come, row by row.
                    if self.edges.len() - from > 4096 {
                        self.edges[from..].sort_unstable();
                        let mut tail = self.edges.split_off(from);
                        tail.dedup();
                        self.edges.append(&mut tail);
                    }
                }
                self.held = held;
            }
            let from = *self.edge_firsts.last().expect("a first") as usize;
            self.edges[from..].sort_unstable();
            let mut k = from;
            let mut w = from;
            while k < self.edges.len() {
                if w == from || self.edges[w - 1] != self.edges[k] {
                    self.edges[w] = self.edges[k];
                    w += 1;
                }
                k += 1;
            }
            self.edges.truncate(w);
            met.extend(self.edges[from..].iter().map(|&(y, _)| y));
            met.dedup();
            targets.extend_from_slice(&met);
            offsets.push(targets.len());
            self.edge_firsts.push(self.edges.len() as u32);
        }
        // Each hub's nodes, found once.
        self.hub_firsts.push(0);
        for h in 0..self.held.areas.len() {
            let (t, a) = self.held.areas[h];
            let mut mine: Vec<u32> = Vec::new();
            self.index.meeting(t as usize, &a, &mut |y| mine.push(y));
            mine.sort_unstable();
            mine.dedup();
            targets.extend_from_slice(&mine);
            self.hub_nodes.extend_from_slice(&mine);
            offsets.push(targets.len());
            self.hub_firsts.push(self.hub_nodes.len() as u32);
        }
        (offsets, targets)
    }

    /// A hub's nodes.
    fn hub_nodes_of(&self, h: usize) -> &[u32] {
        &self.hub_nodes[self.hub_firsts[h] as usize..self.hub_firsts[h + 1] as usize]
    }

    /// The deepest level of a hub's nodes' cells inside its area, its nodes
    /// planned.
    fn hub_level(&self, h: usize) -> Option<u32> {
        let (_, a) = self.held.areas[h];
        self.hub_nodes_of(h)
            .iter()
            .filter_map(|&y| self.deepest_in(y as usize, &a))
            .max()
    }

    fn edges_of(&self, x: usize) -> &[(u32, u32)] {
        &self.edges[self.edge_firsts[x] as usize..self.edge_firsts[x + 1] as usize]
    }

    fn reads_itself(&self, x: usize) -> bool {
        self.edges_of(x).iter().any(|&(y, _)| y as usize == x)
    }

    /// One cell's reads of a node planned cell by cell, to `out`: each
    /// reference's areas at the cell's offset, an area every cell reads
    /// alike by its index among `held`'s where `held` is given and else as an
    /// area, and the cells a defined name's text or a text the parse does not
    /// read names, as the plan of cells found them. Found where they are
    /// used and never kept: Sand's two sheets, each reading the other's rows,
    /// one component of 125,000 cells, kept 24 bytes a read a cell.
    fn reads_of(
        &self,
        x: usize,
        (row, col): (u32, u32),
        mut held: Option<&mut HeldAreas>,
        out: &mut Vec<Item>,
    ) {
        let node = self.nodes[x];
        let shapes: &'a [Shape] = self.shapes;
        let shape = &shapes[node.shape as usize];
        let sheet = node.sheet as usize;
        let (a_row, a_col) = self.anchor(x);
        let model = self.model;
        let find = |n: &str| model.find_sheet(n);
        let id = (sheet, row, col);
        let (d_row, d_col) = (
            i64::from(row) - i64::from(a_row),
            i64::from(col) - i64::from(a_col),
        );
        let mut deps: Vec<CellId> = Vec::new();
        match &shape.expr {
            Ok(e) if !self.unread.contains_key(&id) => {
                let names = self.names;
                let extents = self.extents;
                let reads = {
                    let mut lists = self.read_lists.borrow_mut();
                    let list = &mut lists[node.shape as usize];
                    if list.is_none() {
                        let mut all: Vec<Read<'a>> = Vec::new();
                        e.each_read(&mut |r| all.push(r));
                        *list = Some(all.into());
                    }
                    list.clone().expect("listed")
                };
                reads.iter().for_each(|&r| match r {
                    Read::Ref(re) => {
                        let Ok(areas) = areas_found(find, sheet, re, d_row, d_col) else {
                            return;
                        };
                        let fixed = re.last_sheet.is_none()
                            && Dim::rows(re).is_some_and(Dim::fixed)
                            && Dim::cols(re).is_some_and(Dim::fixed);
                        for a in &areas {
                            let r = A1Range {
                                top: a.top,
                                left: a.left,
                                bottom: a.bottom,
                                right: a.right,
                            };
                            out.push(match (fixed, held.as_deref_mut()) {
                                (true, Some(held)) => Item::Held(held.id(a.sheet, r)),
                                _ => Item::Area(a.sheet as u32, r),
                            });
                        }
                    }
                    Read::Name {
                        name,
                        sheet: written,
                    } => {
                        let scope = written.and_then(|s| model.find_sheet(s));
                        if let Some(t) = names.resolve(name, scope, sheet) {
                            dependencies_in(
                                model,
                                names,
                                extents,
                                scope.unwrap_or(sheet),
                                t,
                                &mut deps,
                            );
                        }
                    }
                });
            }
            _ => {
                let moved = refers::shift_a1_references(&shape.text, d_row, d_col);
                dependencies_in(model, self.names, self.extents, sheet, &moved, &mut deps);
            }
        }
        out.extend(deps.iter().map(|&c| Item::Cell(c)));
    }

    /// A node's rectangles of one level each, from the rectangles of its
    /// cells each edge raises to one past the level of what it reads;
    /// `None` where they are too many, which plans it cell by cell.
    fn block_levels(&self, x: usize) -> Option<Vec<(A1Range, u32)>> {
        let node = self.nodes[x];
        let reads = self.reads[node.shape as usize].as_ref()?;
        let anchor = self.anchor(x);
        let mut raises: Vec<(A1Range, u32)> = Vec::new();
        let n = self.nodes.len() as u32;
        for &(y, k) in self.edges_of(x) {
            if y >= n {
                // An area every cell reads alike raises every cell.
                if let Some(level) = self.hub_levels[(y - n) as usize] {
                    raises.push((node.rect, level + 1));
                }
                continue;
            }
            let read = &reads[k as usize];
            for (q, level) in self.pieces_of(y as usize) {
                if let Some(r) = cells_reading(&node.rect, anchor, read, q) {
                    raises.push((r, (level & !CYCLE) + 1));
                    if raises.len() > RAISES {
                        return None;
                    }
                }
            }
        }
        level_map(node.rect, &raises)
    }

    /// A node's rectangles of one level each, kept.
    fn set_pieces(&mut self, x: usize, map: &[(A1Range, u32)]) {
        let at = self.piece_store.len() as u32;
        self.piece_store.extend_from_slice(map);
        self.levels[x] = Levels::Pieces {
            at,
            len: map.len() as u32,
        };
    }

    /// A node's rectangles of one level each, a cell planned cell by cell's
    /// a row's run, a cycle's cells marked.
    fn pieces_of(&self, y: usize) -> &[(A1Range, u32)] {
        let (at, len) = match self.levels[y] {
            Levels::Pieces { at, len } => (at, len),
            Levels::Cells { spans, len, .. } => (spans, len),
            Levels::Unknown => unreachable!("a node read is planned before its reader"),
        };
        &self.piece_store[at as usize..(at + len) as usize]
    }

    /// Each cell's level of a node planned cell by cell, row by row; `None`
    /// for one planned by its rectangles.
    fn cells_of(&self, y: usize) -> Option<&[u32]> {
        match self.levels[y] {
            Levels::Cells { at, .. } => {
                let k = self.nodes[y].rect.cells() as usize;
                Some(&self.cell_store[at as usize..at as usize + k])
            }
            _ => None,
        }
    }

    /// The level of a cell of a node already planned.
    fn level_at(&self, y: usize, row: u32, col: u32) -> u32 {
        match self.cells_of(y) {
            Some(cells) => {
                let r = &self.nodes[y].rect;
                cells[((row - r.top) * (r.right - r.left + 1) + (col - r.left)) as usize] & !CYCLE
            }
            None => self
                .pieces_of(y)
                .iter()
                .find(|(r, _)| holds(r, row, col))
                .map_or(0, |&(_, l)| l),
        }
    }

    /// The deepest level of a node's cells already planned inside an area.
    fn deepest_in(&self, y: usize, a: &A1Range) -> Option<u32> {
        self.pieces_of(y)
            .iter()
            .filter(|(r, _)| overlap(r, a).is_some())
            .map(|&(_, l)| l & !CYCLE)
            .max()
    }

    /// The nodes of a component planned cell by cell: each cell's reads,
    /// those of the component's cells as edges between them and those of
    /// any other node by the deepest level they reach; Tarjan's components
    /// over the cells, and each cell's level one past the deepest it reads
    /// outside its own, a cycle's cells marked; an area every cell reads
    /// alike that names many of the component's cells read through a hub.
    fn cell_levels(&mut self, comp: &[u32]) {
        let mut held = std::mem::take(&mut self.held);
        let mut items: Vec<Item> = Vec::new();
        // The component's cells as vertices, node by node, row by row.
        let mut base: FxHashMap<u32, u32> = FxHashMap::default();
        let mut firsts: Vec<u32> = Vec::with_capacity(comp.len());
        let mut cells = 0u32;
        for &x in comp {
            base.insert(x, cells);
            firsts.push(cells);
            cells += self.nodes[x as usize].rect.cells() as u32;
        }
        let n = cells as usize;
        let mut offsets: Vec<usize> = Vec::with_capacity(n + 1);
        let mut targets: Vec<u32> = Vec::new();
        let mut outside: Vec<Option<u32>> = vec![None; n];
        offsets.push(0);
        // Each area read alike: its cells of the component, the deepest level
        // it reaches outside, and its hub.
        let mut groups: FxHashMap<u32, Group> = FxHashMap::default();
        let mut hubs: Vec<u32> = Vec::new();
        let mut edges: Vec<u32> = Vec::new();
        for &x in comp {
            let x = x as usize;
            let rect = self.nodes[x].rect;
            let w = rect.right - rect.left + 1;
            for k in 0..rect.cells() as usize {
                let v = (base[&(x as u32)] as usize) + k;
                let cell = (rect.top + k as u32 / w, rect.left + k as u32 % w);
                items.clear();
                self.reads_of(x, cell, Some(&mut held), &mut items);
                edges.clear();
                let mut deepest: Option<u32> = None;
                for item in &items {
                    match *item {
                        Item::Area(t, a) => {
                            self.area_reads(t as usize, &a, &base, &mut edges, &mut deepest);
                        }
                        Item::Held(h) => {
                            let (inside, out, hub) = &*groups.entry(h).or_insert_with(|| {
                                let (t, a) = held.areas[h as usize];
                                let mut inside = Vec::new();
                                let mut out = None;
                                self.area_reads(t as usize, &a, &base, &mut inside, &mut out);
                                inside.sort_unstable();
                                inside.dedup();
                                let hub = (inside.len() >= hub_at()).then(|| {
                                    hubs.push(h);
                                    (n + hubs.len() - 1) as u32
                                });
                                (inside, out, hub)
                            });
                            match hub {
                                Some(hub) => edges.push(*hub),
                                None => edges.extend_from_slice(inside),
                            }
                            deepest = deepest.max(*out);
                        }
                        Item::Cell(c) => {
                            let Some(y) = self.index.at(c) else {
                                continue;
                            };
                            match base.get(&y) {
                                Some(&b) => {
                                    let r = self.nodes[y as usize].rect;
                                    edges.push(
                                        b + (c.1 - r.top) * (r.right - r.left + 1) + (c.2 - r.left),
                                    );
                                }
                                None => {
                                    deepest = deepest.max(Some(self.level_at(y as usize, c.1, c.2)))
                                }
                            }
                        }
                    }
                }
                outside[v] = deepest;
                edges.sort_unstable();
                edges.dedup();
                targets.extend_from_slice(&edges);
                offsets.push(targets.len());
            }
        }
        for h in &hubs {
            let (inside, out, _) = &groups[h];
            targets.extend_from_slice(inside);
            offsets.push(targets.len());
            outside.push(*out);
        }
        drop(groups);
        self.held = held;
        self.hubs += hubs.len();
        let all = offsets.len() - 1;
        let adj = |v: usize| &targets[offsets[v]..offsets[v + 1]];
        let (members, starts) = tarjan_flat(&offsets, &targets);
        let mut component = vec![0usize; all];
        let mut level = vec![0u32; all];
        let mut cycle = vec![false; n];
        for (c, ends) in starts.windows(2).enumerate() {
            let m = &members[ends[0]..ends[1]];
            for &v in m {
                component[v as usize] = c;
            }
            let deepest = m
                .iter()
                .flat_map(|&v| adj(v as usize))
                .filter(|&&u| component[u as usize] != c)
                .map(|&u| level[u as usize] + 1)
                .chain(m.iter().filter_map(|&v| outside[v as usize].map(|l| l + 1)))
                .max()
                .unwrap_or(0);
            if m.iter().all(|&v| v as usize >= n) {
                for &v in m {
                    level[v as usize] = deepest.saturating_sub(1);
                }
                continue;
            }
            for &v in m {
                level[v as usize] = deepest;
            }
            if !(m.len() == 1 && !adj(m[0] as usize).contains(&m[0])) {
                let mut on: Vec<CellId> = m
                    .iter()
                    .filter(|&&v| (v as usize) < n)
                    .map(|&v| {
                        cycle[v as usize] = true;
                        self.cell_of(comp, &firsts, v)
                    })
                    .collect();
                on.sort_unstable();
                self.cycles.push(on);
            }
        }
        for &x in comp {
            let rect = self.nodes[x as usize].rect;
            let b = base[&x] as usize;
            let k = rect.cells() as usize;
            let at = self.cell_store.len() as u32;
            self.cell_store
                .extend((b..b + k).map(|v| level[v] | if cycle[v] { CYCLE } else { 0 }));
            let spans = self.piece_store.len() as u32;
            let mut runs = spans_of(&rect, &self.cell_store[at as usize..]);
            self.piece_store.append(&mut runs);
            self.levels[x as usize] = Levels::Cells {
                at,
                spans,
                len: self.piece_store.len() as u32 - spans,
            };
        }
    }

    /// A vertex of a component planned cell by cell, as its cell: its node
    /// the last whose first vertex is at most it, `firsts` rising with the
    /// component's nodes.
    fn cell_of(&self, comp: &[u32], firsts: &[u32], v: u32) -> CellId {
        let i = firsts.partition_point(|&f| f <= v) - 1;
        let node = &self.nodes[comp[i] as usize];
        let k = v - firsts[i];
        let w = node.rect.right - node.rect.left + 1;
        (
            node.sheet as usize,
            node.rect.top + k / w,
            node.rect.left + k % w,
        )
    }

    /// An area's reads from a cell planned cell by cell: each cell of the
    /// component inside it as a vertex, and the deepest level of every other
    /// node's cells inside it.
    fn area_reads(
        &self,
        t: usize,
        a: &A1Range,
        base: &FxHashMap<u32, u32>,
        edges: &mut Vec<u32>,
        deepest: &mut Option<u32>,
    ) {
        self.index.meeting(t, a, &mut |y| {
            let r = self.nodes[y as usize].rect;
            let Some(cut) = overlap(&r, a) else {
                return;
            };
            match base.get(&y) {
                Some(&b) => {
                    let w = r.right - r.left + 1;
                    for row in cut.top..=cut.bottom {
                        for col in cut.left..=cut.right {
                            edges.push(b + (row - r.top) * w + (col - r.left));
                        }
                    }
                }
                None => *deepest = (*deepest).max(self.deepest_in(y as usize, &cut)),
            }
        });
    }

    /// The order from the levels: the blocks, each formula's position, the
    /// runs and the stretches, and where each level ends.
    fn order(&self, alone: &mut [FxHashMap<(u32, u32), Alone>]) -> (Planned, Made) {
        // Every node's cells of one level, by level, shape and sheet.
        let mut parts: Vec<(u32, u32, A1Range)> = Vec::new();
        let mut deepest: Option<u32> = None;
        for x in 0..self.nodes.len() {
            for &(r, level) in self.pieces_of(x) {
                deepest = deepest.max(Some(level & !CYCLE));
                if level & CYCLE == 0 {
                    parts.push((level, x as u32, r));
                }
            }
        }
        let nodes = &self.nodes;
        let key = |&(level, x, _): &(u32, u32, A1Range)| {
            let n = &nodes[x as usize];
            (level, n.shape, n.sheet)
        };
        parts.sort_unstable_by_key(|p| (key(p), p.2.top, p.2.left));
        let mut blocks = Blocks::default();
        let mut block_of_part: Vec<u32> = Vec::with_capacity(parts.len());
        let mut block_levels: Vec<u32> = Vec::new();
        let mut by_key: FxHashMap<(u32, u32, u32), u32> = FxHashMap::default();
        let mut k = 0;
        while k < parts.len() {
            let at = key(&parts[k]);
            let mut j = k;
            while j < parts.len() && key(&parts[j]) == at {
                j += 1;
            }
            let b = blocks.blocks().len() as u32;
            blocks.push(at.1, at.2, parts[k..j].iter().map(|p| p.2).collect());
            by_key.insert(at, b);
            block_levels.push(at.0);
            block_of_part.extend(std::iter::repeat(b).take(j - k));
            k = j;
        }
        blocks.shrink();
        let levels = deepest.map_or(0, |d| d as usize + 1);
        let mut level_ends = vec![0u32; levels];
        for (b, &level) in block_levels.iter().enumerate() {
            level_ends[level as usize] += blocks.count(b) as u32;
        }
        for l in 1..levels {
            level_ends[l] += level_ends[l - 1];
        }
        // Where each node's cells stand: a formula alone's position beside its
        // shape, a piece's rectangles with their blocks or each cell's
        // position.
        let mut located = Located::default();
        // Each fill's piece's rectangles with their blocks, counted into place.
        let mut part_firsts = vec![0u32; self.nodes.len() + 1];
        for &(_, x, _) in &parts {
            if !self.nodes[x as usize].alone {
                part_firsts[x as usize + 1] += 1;
            }
        }
        for x in 0..self.nodes.len() {
            part_firsts[x + 1] += part_firsts[x];
        }
        let mut node_parts = vec![
            (
                A1Range {
                    top: 0,
                    left: 0,
                    bottom: 0,
                    right: 0
                },
                0u32
            );
            part_firsts[self.nodes.len()] as usize
        ];
        let mut fill = part_firsts.clone();
        for (p, &(_, x, r)) in parts.iter().enumerate() {
            if !self.nodes[x as usize].alone {
                node_parts[fill[x as usize] as usize] = (r, block_of_part[p]);
                fill[x as usize] += 1;
            }
        }
        drop(fill);
        for (x, node) in self.nodes.iter().enumerate() {
            let position = |row: u32, col: u32, level: u32| -> u32 {
                if level & CYCLE != 0 {
                    return NONE;
                }
                let b = by_key[&(level, node.shape, node.sheet)];
                blocks
                    .position_in(b as usize, row, col)
                    .map_or(NONE, |p| p as u32)
            };
            if node.alone {
                let level = self.pieces_of(x)[0].1;
                if let Some(a) = alone
                    .get_mut(node.sheet as usize)
                    .and_then(|m| m.get_mut(&(node.rect.top, node.rect.left)))
                {
                    a.pos = position(node.rect.top, node.rect.left, level);
                }
                continue;
            }
            match self.cells_of(x) {
                None => {
                    let mine = &node_parts[part_firsts[x] as usize..part_firsts[x + 1] as usize];
                    located.nodes.push(Node {
                        sheet: node.sheet,
                        rect: node.rect,
                        shape: node.shape,
                        at: located.parts.len() as u32,
                        len: mine.len() as u32,
                        cells: false,
                    });
                    located.parts.extend_from_slice(mine);
                }
                Some(lv) => {
                    let at = located.dense.len() as u32;
                    let r = node.rect;
                    let w = r.right - r.left + 1;
                    for (k, &level) in lv.iter().enumerate() {
                        let k = k as u32;
                        located
                            .dense
                            .push(position(r.top + k / w, r.left + k % w, level));
                    }
                    located.nodes.push(Node {
                        sheet: node.sheet,
                        rect: r,
                        shape: node.shape,
                        at,
                        len: lv.len() as u32,
                        cells: true,
                    });
                }
            }
        }
        let shapes = self.shapes.len();
        located.shape_firsts = vec![0u32; shapes + 1];
        for n in &located.nodes {
            located.shape_firsts[n.shape as usize + 1] += 1;
        }
        for s in 0..shapes {
            located.shape_firsts[s + 1] += located.shape_firsts[s];
        }
        let mut at = located.shape_firsts.clone();
        located.shape_nodes = vec![0u32; located.nodes.len()];
        for (k, n) in located.nodes.iter().enumerate() {
            located.shape_nodes[at[n.shape as usize] as usize] = k as u32;
            at[n.shape as usize] += 1;
        }
        // Each shape's blocks, counted into place.
        located.block_firsts = vec![0u32; shapes + 1];
        for b in blocks.blocks() {
            located.block_firsts[b.shape as usize + 1] += 1;
        }
        for s in 0..shapes {
            located.block_firsts[s + 1] += located.block_firsts[s];
        }
        let mut at = located.block_firsts.clone();
        located.shape_blocks = vec![0u32; blocks.blocks().len()];
        for (k, b) in blocks.blocks().iter().enumerate() {
            located.shape_blocks[at[b.shape as usize] as usize] = k as u32;
            at[b.shape as usize] += 1;
        }
        located.by_sheet = vec![Rects::default(); self.model.sheet_count()];
        located.alone_grid = alone.iter().map(AloneGrid::of).collect();
        located.how = (0..shapes)
            .map(|s| {
                let blocks = located.block_firsts[s + 1] - located.block_firsts[s];
                let pieces = located.shape_firsts[s + 1] - located.shape_firsts[s];
                match (blocks as usize <= FEW, pieces > 0) {
                    (true, _) => BY_BLOCKS,
                    (false, true) => BY_PIECES,
                    (false, false) => BY_PLACE,
                }
            })
            .collect();
        for (k, n) in located.nodes.iter().enumerate() {
            located.by_sheet[n.sheet as usize].insert(n.rect, k as u32);
        }
        located.parts.shrink_to_fit();
        located.dense.shrink_to_fit();
        located.nodes.shrink_to_fit();
        let formulas = self.nodes.iter().map(|n| n.rect.cells() as usize).sum();
        let planned = Planned {
            blocks,
            located,
            cycles: Vec::new(),
            runs: Vec::new(),
            level_ends,
            stretches: Vec::new(),
            stretch_firsts: Vec::new(),
            formulas,
            hubs: 0,
            by_cell: 0,
        };
        let made = Made {
            parts,
            block_of_part,
            by_key,
        };
        (planned, made)
    }

    /// The runs and the stretches, block by block: a run carried into the
    /// next block where the two are one shape on one sheet, and broken at
    /// the first of its cells that reads a cell of the run; every run its
    /// shape's stretch, or the last one's grown, and kept to be evaluated at
    /// once when it is long enough and its shape parsed and not risky.
    fn runs(&self, planned: &mut Planned, made: &Made) {
        let blocks = &planned.blocks;
        let count = blocks.blocks().len();
        let mut part_firsts = vec![0usize; count + 1];
        for &b in &made.block_of_part {
            part_firsts[b as usize + 1] += 1;
        }
        for b in 0..count {
            part_firsts[b + 1] += part_firsts[b];
        }
        let mut runs: Vec<Run> = Vec::new();
        let mut stretches: Vec<(u32, u32, u32)> = Vec::new();
        let mut keep = |run: Run| {
            let Some(shape) = self.shapes.get(run.shape) else {
                return;
            };
            let (s, start, end) = (run.shape as u32, run.start as u32, run.end as u32);
            match stretches.last_mut() {
                Some(last) if last.0 == s && last.2 == start => last.2 = end,
                _ => stretches.push((s, start, end)),
            }
            if run.end - run.start >= RUN && !shape.risky && shape.expr.is_ok() {
                runs.push(run);
            }
        };
        let mut run = Run {
            start: 0,
            end: 0,
            shape: usize::MAX,
        };
        let mut run_sheet = u32::MAX;
        for (k, b) in blocks.blocks().iter().enumerate() {
            let start = b.start as usize;
            if b.shape as usize == run.shape && b.sheet == run_sheet {
                let mine = &made.parts[part_firsts[k]..part_firsts[k + 1]];
                if let Some(rank) = self.first_reading(planned, made, k, run.start, mine) {
                    run.end = start + rank;
                    keep(run);
                    run.start = start + rank;
                }
            } else {
                keep(run);
                run = Run {
                    start,
                    end: start,
                    shape: b.shape as usize,
                };
                run_sheet = b.sheet;
            }
            run.end = start + blocks.count(k);
        }
        keep(run);
        // Each shape's stretches together, by a counting sort (KERNEL.54).
        let shapes = self.shapes.len();
        let mut firsts = vec![0u32; shapes + 1];
        for &(s, _, _) in &stretches {
            firsts[s as usize + 1] += 1;
        }
        for s in 0..shapes {
            firsts[s + 1] += firsts[s];
        }
        let mut at = firsts.clone();
        let mut by_shape = vec![(0, 0); stretches.len()];
        for (s, start, end) in stretches {
            by_shape[at[s as usize] as usize] = (start, end);
            at[s as usize] += 1;
        }
        planned.runs = runs;
        planned.stretches = by_shape;
        planned.stretch_firsts = firsts;
    }

    /// The rank in a block of its first cell, row by row, that reads a cell
    /// of the run carried into it, which began at `run_start`; `None` where
    /// none does. A piece's cells reading it are found by the rectangles of
    /// its edges into the run's cells, a cell planned cell by cell by its
    /// reads.
    fn first_reading(
        &self,
        planned: &Planned,
        made: &Made,
        k: usize,
        run_start: usize,
        mine: &[(u32, u32, A1Range)],
    ) -> Option<usize> {
        let blocks = &planned.blocks;
        let b = blocks.blocks()[k];
        let end = b.start as usize;
        let mut best: Option<(u32, u32)> = None;
        for &(_, x, q) in mine {
            let x = x as usize;
            let (shape, sheet) = (self.nodes[x].shape, self.nodes[x].sheet);
            match self.levels[x] {
                Levels::Pieces { .. } => {
                    let reads = self.reads[shape as usize]
                        .as_ref()
                        .expect("a piece's reads");
                    let anchor = self.anchor(x);
                    let rect = self.nodes[x].rect;
                    let n = self.nodes.len() as u32;
                    for &(y, read) in self.edges_of(x) {
                        if y >= n {
                            // An area every cell reads alike: every cell of
                            // the piece reads the run where it meets it.
                            let (t, a) = self.held.areas[(y - n) as usize];
                            let meets = t == sheet
                                && self.hub_nodes_of((y - n) as usize).iter().any(|&z| {
                                    let nz = &self.nodes[z as usize];
                                    nz.shape == shape
                                        && nz.sheet == sheet
                                        && self
                                            .run_region(planned, made, z as usize, run_start, end)
                                            .iter()
                                            .any(|r| overlap(r, &a).is_some())
                                });
                            if meets {
                                let at = (q.top, q.left);
                                best = Some(best.map_or(at, |first| first.min(at)));
                            }
                            continue;
                        }
                        let ny = &self.nodes[y as usize];
                        if ny.shape != shape || ny.sheet != sheet {
                            continue;
                        }
                        for z in self.run_region(planned, made, y as usize, run_start, end) {
                            let Some(r) = cells_reading(&rect, anchor, &reads[read as usize], &z)
                            else {
                                continue;
                            };
                            if let Some(c) = overlap(&r, &q) {
                                let at = (c.top, c.left);
                                best = Some(best.map_or(at, |first| first.min(at)));
                            }
                        }
                    }
                }
                Levels::Cells { .. } => {
                    'cells: for row in q.top..=q.bottom {
                        for col in q.left..=q.right {
                            if best.is_some_and(|first| (row, col) >= first) {
                                break 'cells;
                            }
                            if self.cell_reads(
                                planned,
                                x,
                                (row, col),
                                (shape, sheet),
                                (run_start, end),
                            ) {
                                best = Some((row, col));
                                break 'cells;
                            }
                        }
                    }
                }
                Levels::Unknown => unreachable!("every node is planned"),
            }
        }
        let (row, col) = best?;
        Some(blocks.rank(k, row, col).expect("a block's cell") as usize)
    }

    /// The cells of a node whose positions are from `from` to before `to`,
    /// as rectangles.
    fn run_region(
        &self,
        planned: &Planned,
        made: &Made,
        y: usize,
        from: usize,
        to: usize,
    ) -> Vec<A1Range> {
        let blocks = &planned.blocks;
        let node = self.nodes[y];
        let mut out = Vec::new();
        match self.levels[y] {
            Levels::Pieces { .. } => {
                for &(q, level) in self.pieces_of(y) {
                    let b = made.by_key[&(level, node.shape, node.sheet)] as usize;
                    let bs = blocks.blocks()[b].start as usize;
                    let be = bs + blocks.count(b);
                    if be <= from || bs >= to {
                        continue;
                    }
                    if bs >= from {
                        out.push(q);
                        continue;
                    }
                    // The run began inside the block: the cells from its first.
                    let (_, row0, col0) = blocks.cell_in(b, from);
                    out.extend(from_cell(&q, (row0, col0)));
                }
            }
            Levels::Cells { .. } => {
                let r = node.rect;
                for row in r.top..=r.bottom {
                    for col in r.left..=r.right {
                        let id = (node.sheet as usize, row, col);
                        if let Some(p) = planned.located.position(blocks, id) {
                            if from <= p && p < to {
                                out.push(A1Range {
                                    top: row,
                                    left: col,
                                    bottom: row,
                                    right: col,
                                });
                            }
                        }
                    }
                }
            }
            Levels::Unknown => unreachable!("every node is planned"),
        }
        out
    }

    /// Whether a cell planned cell by cell reads a cell of a shape on a sheet
    /// whose position is from `from` to before `to`.
    fn cell_reads(
        &self,
        planned: &Planned,
        x: usize,
        (row, col): (u32, u32),
        (shape, sheet): (u32, u32),
        (from, to): (usize, usize),
    ) -> bool {
        let mut items: Vec<Item> = Vec::new();
        self.reads_of(x, (row, col), None, &mut items);
        let blocks = &planned.blocks;
        for item in &items {
            let (t, a) = match *item {
                Item::Area(t, a) => (t, a),
                Item::Held(h) => self.held.areas[h as usize],
                Item::Cell(c) => {
                    if planned
                        .located
                        .position(blocks, c)
                        .is_some_and(|p| from <= p && p < to)
                    {
                        return true;
                    }
                    continue;
                }
            };
            if t != sheet {
                continue;
            }
            let mut found = false;
            planned
                .located
                .each_in(blocks, shape as usize, sheet as usize, &a, &mut |p| {
                    found |= from <= p && p < to;
                });
            if found {
                return true;
            }
        }
        false
    }
}

/// A rectangle's cells at a cell or after it, row by row, as at most two
/// rectangles: the rest of the cell's row, and the rows below it.
fn from_cell(q: &A1Range, (row, col): (u32, u32)) -> impl Iterator<Item = A1Range> {
    let rest = (q.top <= row && row <= q.bottom && col.max(q.left) <= q.right).then(|| A1Range {
        top: row,
        bottom: row,
        left: q.left.max(col),
        right: q.right,
    });
    let below = (q.bottom > row).then(|| A1Range {
        top: q.top.max(row + 1),
        ..*q
    });
    rest.into_iter().chain(below)
}

/// A node's cells planned cell by cell as rectangles of one level each, row
/// by row: each row's runs of one level, a cycle's cells marked.
fn spans_of(rect: &A1Range, levels: &[u32]) -> Vec<(A1Range, u32)> {
    let w = (rect.right - rect.left + 1) as usize;
    let mut spans = Vec::new();
    for (i, row) in levels.chunks(w).enumerate() {
        let r = rect.top + i as u32;
        let mut j = 0;
        while j < w {
            let mut k = j + 1;
            while k < w && row[k] == row[j] {
                k += 1;
            }
            spans.push((
                A1Range {
                    top: r,
                    left: rect.left + j as u32,
                    bottom: r,
                    right: rect.left + k as u32 - 1,
                },
                row[j],
            ));
            j = k;
        }
    }
    spans
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn a_rectangle_from_a_cell_on_is_its_cells_from_that_cell_row_by_row() {
        // The run's region where it began inside a block (`Planner::runs`):
        // every rectangle of up to five rows and columns against every cell
        // around it, the rectangles' cells held to the rectangle's own cells
        // counted one by one at or after the cell, each once.
        for top in 2..4 {
            for left in 2..4 {
                for bottom in top..top + 5 {
                    for right in left..left + 5 {
                        let q = A1Range {
                            top,
                            left,
                            bottom,
                            right,
                        };
                        for row in 1..=bottom + 1 {
                            for col in 1..=right + 1 {
                                let mut got: Vec<(u32, u32)> = from_cell(&q, (row, col))
                                    .flat_map(|r| {
                                        (r.top..=r.bottom).flat_map(move |rr| {
                                            (r.left..=r.right).map(move |cc| (rr, cc))
                                        })
                                    })
                                    .collect();
                                got.sort_unstable();
                                let want: Vec<(u32, u32)> = (top..=bottom)
                                    .flat_map(|rr| (left..=right).map(move |cc| (rr, cc)))
                                    .filter(|&cell| cell >= (row, col))
                                    .collect();
                                assert_eq!(got, want, "{q:?} from {row},{col}");
                            }
                        }
                    }
                }
            }
        }
    }
}
