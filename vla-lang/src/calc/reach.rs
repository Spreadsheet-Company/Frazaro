//! What a changed cell reaches (KERNEL.20, slice 4, 2026-10-10): the cells
//! of an engine's order that read it, so that a frame computes those and
//! keeps every other cell's value. It is the second of the structural
//! advantages `KERNEL.20` names: nothing that did not change is recomputed.
//!
//! A cell's value is a function of what it reads and of where it stands. The
//! library's functions are functions of their arguments and their cell
//! (`super::Library` says so of every one), so a cell none of whose reads
//! changed computes the bits it computed before. A frame reads three kinds
//! of cell:
//! - a twin's cell, which the last frame's end changed;
//! - a value cell, which a write changed;
//! - a formula cell of the same frame, which an earlier cell of the order
//!   changed.
//!
//! Each change is followed here to the cells that read it.
//!
//! No index of readers is kept for each cell: each of Life's cells reads
//! ten, so it would hold 636,000 entries. A shape's references are kept
//! instead, each as a pattern of its sheet and its rows and columns as
//! written, each number with its `$`. The cells of a shape that read a
//! changed cell are the offsets at which the moved reference holds it:
//! [`super::formula::RefExpr::moved`]'s rule turned round. They are taken
//! within the offsets the shape's cells stand at, and each is found by its
//! position in the order.
//!
//! Some parts of a formula are the same at every cell: a constant, a
//! reference every `$` holds, and the operators over them, the parts a run
//! computes once (`super::vector`). Such a part reads the same cells at
//! every cell. A change to one of them reaches the whole shape, but first
//! the part is computed again, once, at the frame's start. When it gives the
//! bits it gave before, no cell of the shape can tell, and none is computed
//! for it. Life reads `Clock!$B$1` at every cell, which changes every
//! frame, through `Clock!$B$1=1`, which after the first two frames does not.
//!
//! Two limits keep that sound:
//! - A call is never such a part, since a function may read its cell
//!   (`ROW()`).
//! - A part is computed at the frame's start only when every cell it reads
//!   is final there: a value cell or a twin's.
//!
//! A shape is not followed when its formula reads a defined name, whose
//! text may read its cell, or a span of sheets. Its cells are computed
//! every frame, as every cell was.

use crate::fx::FxHashMap;
use crate::refers::Shape as Part;
use crate::sheet::{Workbook, MAX_COLUMN, MAX_ROW};

use super::formula::{Expr, Read, RefExpr};
use super::graph::CellId;
use super::shape::Shape;

/// One dimension of a reference, its rows or its columns.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Dim {
    /// A whole column's rows, or a whole row's columns, at every offset.
    Whole { max: u32 },
    /// Two numbers as written, each with its `$` (a cell's first again):
    /// the dimension holds what lies between them once moved.
    Span {
        n1: u32,
        abs1: bool,
        n2: u32,
        abs2: bool,
        max: u32,
    },
}

impl Dim {
    /// A reference's rows, as [`RefExpr::moved`] moves them.
    fn rows(r: &RefExpr) -> Option<Dim> {
        let c = &r.corners;
        let span = |n1, abs1, n2, abs2| Dim::Span {
            n1,
            abs1,
            n2,
            abs2,
            max: MAX_ROW,
        };
        match r.shape {
            Part::Cell => Some(span(c.row1, c.row_abs1, c.row1, c.row_abs1)),
            Part::Range | Part::Row => Some(span(c.row1, c.row_abs1, c.row2, c.row_abs2)),
            Part::Column => Some(Dim::Whole { max: MAX_ROW }),
            Part::None => None,
        }
    }

    /// A reference's columns, as [`RefExpr::moved`] moves them.
    fn cols(r: &RefExpr) -> Option<Dim> {
        let c = &r.corners;
        let span = |n1, abs1, n2, abs2| Dim::Span {
            n1,
            abs1,
            n2,
            abs2,
            max: MAX_COLUMN,
        };
        match r.shape {
            Part::Cell => Some(span(c.col1, c.col_abs1, c.col1, c.col_abs1)),
            Part::Range | Part::Column => Some(span(c.col1, c.col_abs1, c.col2, c.col_abs2)),
            Part::Row => Some(Dim::Whole { max: MAX_COLUMN }),
            Part::None => None,
        }
    }

    /// Whether every offset holds the same numbers.
    fn fixed(self) -> bool {
        match self {
            Dim::Whole { .. } => true,
            Dim::Span { abs1, abs2, .. } => abs1 && abs2,
        }
    }

    /// The numbers a fixed dimension holds.
    fn bounds(self) -> (u32, u32) {
        match self {
            Dim::Whole { max } => (1, max),
            Dim::Span { n1, n2, .. } => (n1.min(n2), n1.max(n2)),
        }
    }

    /// The offsets at which the dimension, moved, holds `at`: an interval,
    /// unbounded where any offset does; `None` when no offset does. A
    /// relative number moved off the sheet is `#REF!`, which holds nothing.
    fn offsets(self, at: u32) -> Option<(i64, i64)> {
        let Dim::Span {
            n1,
            abs1,
            n2,
            abs2,
            max,
        } = self
        else {
            return Some((i64::MIN, i64::MAX));
        };
        let (at, n1, n2, max) = (i64::from(at), i64::from(n1), i64::from(n2), i64::from(max));
        let (mut lo, mut hi) = (i64::MIN, i64::MAX);
        for (n, abs) in [(n1, abs1), (n2, abs2)] {
            if !abs {
                lo = lo.max(1 - n);
                hi = hi.min(max - n);
            }
        }
        let (a, b) = match (abs1, abs2) {
            (true, true) if n1.min(n2) <= at && at <= n1.max(n2) => (i64::MIN, i64::MAX),
            (true, true) => return None,
            (false, false) => (at - n1.max(n2), at - n1.min(n2)),
            (true, false) => between(n1, n2, at),
            (false, true) => between(n2, n1, at),
        };
        let (lo, hi) = (lo.max(a), hi.min(b));
        (lo <= hi).then_some((lo, hi))
    }
}

/// Whether a moving reference `q` holds what `p` holds at every offset a
/// shape's cells stand at (`stand`'s): in each dimension, both numbers of
/// each moved by the offset, `q`'s around `p`'s and never off the sheet
/// there, or `q` the whole of it.
fn covers(q: Kind, p: Kind, stand: &Stand) -> bool {
    let (Kind::Moving { rows: qr, cols: qc }, Kind::Moving { rows: pr, cols: pc }) = (q, p) else {
        return false;
    };
    dim_covers(qr, pr, stand.rows) && dim_covers(qc, pc, stand.cols)
}

/// [`covers`] in one dimension, over the offsets `d0` to `d1`.
fn dim_covers(q: Dim, p: Dim, (d0, d1): (i64, i64)) -> bool {
    match (q, p) {
        (Dim::Whole { .. }, _) => true,
        (
            Dim::Span {
                n1: a1,
                abs1: false,
                n2: a2,
                abs2: false,
                max,
            },
            Dim::Span {
                n1: b1,
                abs1: false,
                n2: b2,
                abs2: false,
                ..
            },
        ) => {
            let (q_lo, q_hi) = (i64::from(a1.min(a2)), i64::from(a1.max(a2)));
            let (p_lo, p_hi) = (i64::from(b1.min(b2)), i64::from(b1.max(b2)));
            q_lo <= p_lo && p_hi <= q_hi && 1 - q_lo <= d0 && d1 <= i64::from(max) - q_hi
        }
        _ => false,
    }
}

/// A number its `$` holds and one the offset moves: the offsets at which
/// the numbers between them hold `at`.
fn between(held: i64, moved: i64, at: i64) -> (i64, i64) {
    match at.cmp(&held) {
        std::cmp::Ordering::Equal => (i64::MIN, i64::MAX),
        std::cmp::Ordering::Greater => (at - moved, i64::MAX),
        std::cmp::Ordering::Less => (i64::MIN, at - moved),
    }
}

/// The sheet a reference reads.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum On {
    /// A sheet of the grid, found once.
    Sheet(usize),
    /// The sheet of each cell that reads it.
    Own,
}

/// The sheet a reference to one sheet reads, as `graph::areas_at` finds it;
/// `None` for a sheet the grid does not hold, which reads `#REF!`.
fn on(model: &Workbook, r: &RefExpr) -> Option<On> {
    match (r.found, &r.sheet) {
        (Some((first, _)), _) => Some(On::Sheet(usize::from(first))),
        (None, None) => Some(On::Own),
        (None, Some(name)) => model.find_sheet(name).map(On::Sheet),
    }
}

/// What a change to a cell a shape reads does to the shape.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Kind {
    /// A reference that moves with the cell reading it: the shape's cells
    /// at the offsets that read the changed cell.
    Moving { rows: Dim, cols: Dim },
    /// A rectangle every cell of the shape reads: all of them, when it
    /// holds the changed cell.
    Held {
        top: u32,
        left: u32,
        bottom: u32,
        right: u32,
    },
    /// A cell read only by parts the same at every cell, each computed at
    /// the frame's start: the shape's parts computed again, when it is the
    /// changed cell.
    Parts { row: u32, col: u32 },
}

/// A shape reading a sheet, as the sheet's readers list it.
#[derive(Clone, Copy, Debug)]
struct Reader {
    shape: usize,
    /// Whether the reference is to the reading cell's own sheet, so that a
    /// change reaches the shape's cells on the changed cell's sheet alone.
    own: bool,
    kind: Kind,
}

/// Where a shape's cells stand.
#[derive(Clone, Debug, Default)]
struct Stand {
    /// The sheets its cells are on, sorted; empty for a shape with none.
    sheets: Vec<usize>,
    /// The row and column of the cell the shape was read at, and the least
    /// and most offsets of its cells from it.
    anchor: (u32, u32),
    rows: (i64, i64),
    cols: (i64, i64),
}

/// What a changed cell reaches.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum Reached {
    /// A formula cell, by its position in the order.
    Cell(usize),
    /// Every cell of a shape.
    Shape(usize),
    /// A shape whose parts the same at every cell read the changed cell.
    Parts(usize),
}

/// What each shape of a plan reads and where its cells stand, found when
/// the plan is: the cells a changed cell reaches. See the module's comment.
#[derive(Clone, Debug, Default)]
pub(crate) struct Reach {
    stands: Vec<Stand>,
    /// By sheet, the shapes that read it and how.
    readers: Vec<Vec<Reader>>,
    /// The shapes no change is followed to, computed every frame.
    unfollowed: Vec<usize>,
    /// Every shape's parts the same at every cell that are computed at a
    /// frame's start; a shape's are `parts[firsts[s]..firsts[s + 1]]`.
    parts: Vec<Expr>,
    firsts: Vec<usize>,
    /// Each formula cell's position in the order.
    positions: Positions,
    /// The shape of the cell at each position of the order.
    shape_at: Vec<u32>,
}

impl Reach {
    /// What the shapes read, against the grid they are placed in, and where
    /// their cells stand in it and in the order; `shape_at` is the shape of
    /// each cell of the order, and `formula` whether a cell holds a formula.
    pub(crate) fn of(
        model: &Workbook,
        shapes: &[Shape],
        order: &[CellId],
        shape_at: Vec<u32>,
        formula: &dyn Fn(CellId) -> bool,
    ) -> Reach {
        let mut stands: Vec<Stand> = vec![Stand::default(); shapes.len()];
        let mut seen = vec![false; shapes.len()];
        for (&(sheet, row, col), &s) in order.iter().zip(&shape_at) {
            let s = s as usize;
            let (a_row, a_col) = (shapes[s].anchor.1, shapes[s].anchor.2);
            let (d_row, d_col) = (
                i64::from(row) - i64::from(a_row),
                i64::from(col) - i64::from(a_col),
            );
            let stand = &mut stands[s];
            if !seen[s] {
                seen[s] = true;
                stand.anchor = (a_row, a_col);
                stand.rows = (d_row, d_row);
                stand.cols = (d_col, d_col);
            }
            stand.rows = (stand.rows.0.min(d_row), stand.rows.1.max(d_row));
            stand.cols = (stand.cols.0.min(d_col), stand.cols.1.max(d_col));
            if !stand.sheets.contains(&sheet) {
                stand.sheets.push(sheet);
            }
        }
        let mut readers: Vec<Vec<Reader>> = vec![Vec::new(); model.sheets.len()];
        let mut unfollowed = Vec::new();
        let mut parts = Vec::new();
        let mut firsts = Vec::with_capacity(shapes.len() + 1);
        for (s, shape) in shapes.iter().enumerate() {
            let first = parts.len();
            firsts.push(first);
            if !seen[s] {
                continue;
            }
            let stand = &mut stands[s];
            stand.sheets.sort_unstable();
            let single = (stand.sheets.len() == 1).then(|| stand.sheets[0]);
            let mut found = Vec::new();
            let followed = match &shape.expr {
                Ok(e) if !shape.risky => {
                    let mut at = At {
                        model,
                        single,
                        formula,
                        found: &mut found,
                        parts: &mut parts,
                        first,
                    };
                    at.read(e)
                }
                _ => false,
            };
            if !followed {
                parts.truncate(first);
                unfollowed.push(s);
                continue;
            }
            // Each entry once, the first kept: a shape holds few references.
            let mut k = 0;
            while k < found.len() {
                if found[..k].contains(&found[k]) {
                    found.remove(k);
                } else {
                    k += 1;
                }
            }
            // A moving reference another of the shape's holds at every offset
            // its cells stand at reaches no cell the other does not: Life's
            // `Screen.last!B2` inside its `Screen.last!A1:C3`.
            let inside: Vec<bool> = (0..found.len())
                .map(|i| {
                    (0..found.len()).any(|j| {
                        let (p, q) = (found[i], found[j]);
                        j != i
                            && p.0 == q.0
                            && covers(q.1, p.1, stand)
                            && (j < i || !covers(p.1, q.1, stand))
                    })
                })
                .collect();
            for (k, (on, kind)) in found.into_iter().enumerate() {
                if inside[k] {
                    continue;
                }
                match on {
                    On::Sheet(x) => {
                        if let Some(list) = readers.get_mut(x) {
                            list.push(Reader {
                                shape: s,
                                own: false,
                                kind,
                            });
                        }
                    }
                    On::Own => {
                        for &y in &stand.sheets {
                            readers[y].push(Reader {
                                shape: s,
                                own: true,
                                kind,
                            });
                        }
                    }
                }
            }
        }
        firsts.push(parts.len());
        Reach {
            stands,
            readers,
            unfollowed,
            parts,
            firsts,
            positions: Positions::of(order, model.sheets.len()),
            shape_at,
        }
    }

    /// Everything a changed cell reaches, to `f`; a cell of the order may
    /// come more than once.
    pub(crate) fn each(&self, changed: CellId, f: &mut impl FnMut(Reached)) {
        let (sheet, row, col) = changed;
        let Some(readers) = self.readers.get(sheet) else {
            return;
        };
        for reader in readers {
            let s = reader.shape;
            match reader.kind {
                Kind::Held {
                    top,
                    left,
                    bottom,
                    right,
                } => {
                    if top <= row && row <= bottom && left <= col && col <= right {
                        f(Reached::Shape(s));
                    }
                }
                Kind::Parts { row: r, col: c } => {
                    if (r, c) == (row, col) {
                        f(Reached::Parts(s));
                    }
                }
                Kind::Moving { rows, cols } => {
                    let stand = &self.stands[s];
                    let (Some((r0, r1)), Some((c0, c1))) = (rows.offsets(row), cols.offsets(col))
                    else {
                        continue;
                    };
                    let (r0, r1) = (r0.max(stand.rows.0), r1.min(stand.rows.1));
                    let (c0, c1) = (c0.max(stand.cols.0), c1.min(stand.cols.1));
                    if r0 > r1 || c0 > c1 {
                        continue;
                    }
                    let on = if reader.own {
                        std::slice::from_ref(&sheet)
                    } else {
                        stand.sheets.as_slice()
                    };
                    for &y in on {
                        for d_row in r0..=r1 {
                            // An offset the shape's cells stand at is a row and
                            // a column on the sheet.
                            let r = (i64::from(stand.anchor.0) + d_row) as u32;
                            for d_col in c0..=c1 {
                                let c = (i64::from(stand.anchor.1) + d_col) as u32;
                                if let Some(p) = self.positions.get(y, r, c) {
                                    if self.shape_at[p] as usize == s {
                                        f(Reached::Cell(p));
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    /// A formula cell's position in the order, where an engine's machine
    /// keeps its value (KERNEL.26); `None` for a cell holding no formula.
    pub(crate) fn position(&self, id: CellId) -> Option<usize> {
        self.positions.get(id.0, id.1, id.2)
    }

    /// The shape of the cell at a position of the order.
    pub(crate) fn shape_at(&self, pos: usize) -> usize {
        self.shape_at[pos] as usize
    }

    /// The shapes no change is followed to, computed every frame.
    pub(crate) fn unfollowed(&self) -> &[usize] {
        &self.unfollowed
    }

    /// A shape's parts computed at a frame's start, by their indices.
    pub(crate) fn parts(&self, shape: usize) -> std::ops::Range<usize> {
        match (self.firsts.get(shape), self.firsts.get(shape + 1)) {
            (Some(&a), Some(&b)) => a..b,
            _ => 0..0,
        }
    }

    /// A part computed at a frame's start, by its index.
    pub(crate) fn part(&self, k: usize) -> &Expr {
        &self.parts[k]
    }

    /// How many parts every shape has together.
    pub(crate) fn part_count(&self) -> usize {
        self.parts.len()
    }
}

/// One shape's parse being read.
struct At<'a, 'p> {
    model: &'a Workbook,
    /// The one sheet the shape's cells are on, when they are on one.
    single: Option<usize>,
    formula: &'a dyn Fn(CellId) -> bool,
    found: &'a mut Vec<(On, Kind)>,
    parts: &'p mut Vec<Expr>,
    /// Where the shape's parts begin in `parts`.
    first: usize,
}

impl At<'_, '_> {
    /// The parse read: each reference as what a change to it does, each
    /// part the same at every cell at the top of what is; `false` when a
    /// change cannot be followed to the shape's cells. After its operands,
    /// round a loop, as every walk of a parse keeps its own stack.
    fn read(&mut self, e: &Expr) -> bool {
        let mut followed = true;
        let mut stack: Vec<(&Expr, bool)> = vec![(e, false)];
        // Each node finished, and whether it is the same at every cell.
        let mut done: Vec<(&Expr, bool)> = Vec::new();
        while let Some((e, left)) = stack.pop() {
            if !left {
                stack.push((e, true));
                match e {
                    Expr::Call(_, args) => stack.extend(args.iter().rev().map(|a| (a, false))),
                    Expr::Neg(x) | Expr::Pos(x) | Expr::Percent(x) => stack.push((x, false)),
                    Expr::Binary(_, a, b) => {
                        stack.push((b, false));
                        stack.push((a, false));
                    }
                    _ => {}
                }
                continue;
            }
            let operands = match e {
                Expr::Call(_, args) => args.len(),
                Expr::Neg(_) | Expr::Pos(_) | Expr::Percent(_) => 1,
                Expr::Binary(..) => 2,
                _ => 0,
            };
            let first = done.len() - operands;
            let same = match e {
                Expr::Number(_) | Expr::Text(_) | Expr::Bool(_) | Expr::Error(_) | Expr::Empty => {
                    true
                }
                Expr::Ref(r) => held(r),
                Expr::Name { .. } => {
                    followed = false;
                    false
                }
                Expr::Neg(_) | Expr::Pos(_) | Expr::Percent(_) => done[first].1,
                Expr::Binary(..) => done[first].1 && done[first + 1].1,
                Expr::Call(..) => false,
            };
            if !same {
                for &(operand, same) in &done[first..] {
                    if same {
                        self.part(operand);
                    }
                }
                if let Expr::Ref(r) = e {
                    followed &= self.reference(r);
                }
            }
            done.truncate(first);
            done.push((e, same));
        }
        if let Some(&(root, true)) = done.last() {
            self.part(root);
        }
        followed
    }

    /// A reference that is not one held cell: a pattern that moves, or a
    /// rectangle every cell reads; `false` for a span of sheets.
    fn reference(&mut self, r: &RefExpr) -> bool {
        if r.last_sheet.is_some() {
            return false;
        }
        // A sheet the grid does not hold, or a reference moved off every
        // sheet, reads nothing.
        let (Some(on), Some(rows), Some(cols)) = (on(self.model, r), Dim::rows(r), Dim::cols(r))
        else {
            return true;
        };
        let kind = if rows.fixed() && cols.fixed() {
            let ((top, bottom), (left, right)) = (rows.bounds(), cols.bounds());
            Kind::Held {
                top,
                left,
                bottom,
                right,
            }
        } else {
            Kind::Moving { rows, cols }
        };
        self.found.push((on, kind));
        true
    }

    /// A part the same at every cell: kept to compute at the frame's start
    /// when it computes a value from cells final there; otherwise each cell
    /// it reads reaches the whole shape.
    fn part(&mut self, u: &Expr) {
        let mut cells: Vec<(On, u32, u32)> = Vec::new();
        let model = self.model;
        u.each_read(&mut |read| {
            if let Read::Ref(r) = read {
                if let Some(on) = on(model, r) {
                    cells.push((on, r.top, r.left));
                }
            }
        });
        if cells.is_empty() {
            return;
        }
        let mut bare = u;
        while let Expr::Pos(x) = bare {
            bare = x;
        }
        let final_at_start = |&(on, row, col): &(On, u32, u32)| match on {
            On::Sheet(s) => !(self.formula)((s, row, col)),
            On::Own => self.single.is_some_and(|s| !(self.formula)((s, row, col))),
        };
        let computed = !matches!(bare, Expr::Ref(_)) && cells.iter().all(final_at_start);
        if computed && !self.parts[self.first..].contains(u) {
            self.parts.push(u.clone());
        }
        for (on, row, col) in cells {
            let kind = if computed {
                Kind::Parts { row, col }
            } else {
                Kind::Held {
                    top: row,
                    left: col,
                    bottom: row,
                    right: col,
                }
            };
            self.found.push((on, kind));
        }
    }
}

/// Whether a reference reads one cell every `$` holds, which a run reads
/// once (`super::vector`).
fn held(r: &RefExpr) -> bool {
    matches!(r.shape, Part::Cell)
        && r.last_sheet.is_none()
        && r.corners.row_abs1
        && r.corners.col_abs1
}

/// No position.
const NONE: u32 = u32::MAX;

/// Each formula cell's position in the order: an array over a sheet's
/// formulas where they stand close enough together, the rest by hash.
#[derive(Clone, Debug, Default)]
struct Positions {
    grids: Vec<Option<Grid>>,
    rest: FxHashMap<CellId, u32>,
}

/// The positions of one sheet's formulas over the rectangle they fill.
#[derive(Clone, Debug)]
struct Grid {
    top: u32,
    left: u32,
    width: u32,
    height: u32,
    at: Vec<u32>,
}

/// A sheet's formulas' rectangle, its top, left, bottom and right, and how
/// many they are.
type Bounds = (u32, u32, u32, u32, u64);

impl Positions {
    fn of(order: &[CellId], sheets: usize) -> Positions {
        let mut bounds: Vec<Option<Bounds>> = vec![None; sheets];
        for &(s, row, col) in order {
            if let Some(b) = bounds.get_mut(s) {
                *b = Some(match *b {
                    None => (row, col, row, col, 1),
                    Some((t, l, bt, r, n)) => {
                        (t.min(row), l.min(col), bt.max(row), r.max(col), n + 1)
                    }
                });
            }
        }
        // An array no more than four times its formulas, so that two
        // formulas far apart are two entries and not a million slots.
        let grids = bounds
            .iter()
            .map(|b| {
                let (top, left, bottom, right, n) = (*b)?;
                let (width, height) = (right - left + 1, bottom - top + 1);
                let cells = u64::from(width) * u64::from(height);
                (cells <= 4 * n + 64).then(|| Grid {
                    top,
                    left,
                    width,
                    height,
                    at: vec![NONE; cells as usize],
                })
            })
            .collect();
        let mut p = Positions {
            grids,
            rest: FxHashMap::default(),
        };
        for (i, &(s, row, col)) in order.iter().enumerate() {
            match p.grids.get_mut(s).and_then(Option::as_mut) {
                Some(g) => {
                    let k = (row - g.top) as usize * g.width as usize + (col - g.left) as usize;
                    g.at[k] = i as u32;
                }
                None => {
                    p.rest.insert((s, row, col), i as u32);
                }
            }
        }
        p
    }

    fn get(&self, sheet: usize, row: u32, col: u32) -> Option<usize> {
        match self.grids.get(sheet) {
            Some(Some(g)) => {
                let (r, c) = (row.checked_sub(g.top)?, col.checked_sub(g.left)?);
                if r >= g.height || c >= g.width {
                    return None;
                }
                let at = g.at[r as usize * g.width as usize + c as usize];
                (at != NONE).then_some(at as usize)
            }
            _ => self.rest.get(&(sheet, row, col)).map(|&p| p as usize),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::calc::formula::parse;

    /// The reference of a one-reference formula.
    fn reference(text: &str) -> RefExpr {
        match parse(text) {
            Ok(Expr::Ref(r)) => r,
            other => panic!("{text}: {other:?}"),
        }
    }

    #[test]
    fn the_offsets_a_reference_holds_a_cell_at_are_those_its_move_says() {
        // Each kind of reference, each `$` mark, near the sheet's first rows
        // and columns where a move leaves it: every offset in a window and
        // every cell near the reference, the offsets `Dim::offsets` names
        // for the cell held to the rectangles `RefExpr::moved` gives.
        let texts = [
            "C3",
            "$C3",
            "C$3",
            "$C$3",
            "B2:D4",
            "$B2:D$4",
            "B$2:$D4",
            "$B$2:$D$4",
            "D4:B2",
            "$D4:B2",
            "D$4:B$2",
            "A1:B2",
            "C:D",
            "$C:D",
            "C:$D",
            "3:4",
            "$3:4",
            "3:$4",
            "A$1:A3",
            "A1:A$3",
        ];
        for text in texts {
            let r = reference(text);
            let (rows, cols) = (Dim::rows(&r).unwrap(), Dim::cols(&r).unwrap());
            for at_row in 1..=8u32 {
                for at_col in 1..=8u32 {
                    let (dr, dc) = (rows.offsets(at_row), cols.offsets(at_col));
                    for d_row in -9i64..=9 {
                        for d_col in -9i64..=9 {
                            let holds = r.moved(d_row, d_col).is_some_and(|(t, l, b, rt)| {
                                t <= at_row && at_row <= b && l <= at_col && at_col <= rt
                            });
                            let named = dr.is_some_and(|(a, b)| a <= d_row && d_row <= b)
                                && dc.is_some_and(|(a, b)| a <= d_col && d_col <= b);
                            assert_eq!(
                                holds, named,
                                "{text} moved by ({d_row}, {d_col}) and the cell R{at_row}C{at_col}"
                            );
                        }
                    }
                }
            }
        }
    }

    #[test]
    fn a_reference_inside_another_at_every_offset_is_covered() {
        let moving = |text: &str| {
            let r = reference(text);
            Kind::Moving {
                rows: Dim::rows(&r).unwrap(),
                cols: Dim::cols(&r).unwrap(),
            }
        };
        // Life's board from B2, its cells at offsets 0 to 197 and 0 to 317.
        let board = Stand {
            sheets: vec![0],
            anchor: (2, 2),
            rows: (0, 197),
            cols: (0, 317),
        };
        assert!(covers(moving("A1:C3"), moving("B2"), &board));
        assert!(!covers(moving("B2"), moving("A1:C3"), &board));
        assert!(covers(moving("A:C"), moving("B2:B9"), &board));
        assert!(!covers(moving("A1:C3"), moving("$B2"), &board));
        // A cell above the board's first row moves `A1:C3` off the sheet,
        // where it holds nothing and `B2` still holds a cell.
        let higher = Stand {
            rows: (-1, 197),
            ..board.clone()
        };
        assert!(!covers(moving("A1:C3"), moving("B2"), &higher));
    }

    #[test]
    fn a_fixed_dimension_holds_its_numbers_at_every_offset() {
        for text in ["$B$2:$D$4", "$C:$D", "$3:$4", "$C$3"] {
            let r = reference(text);
            let (rows, cols) = (Dim::rows(&r).unwrap(), Dim::cols(&r).unwrap());
            assert!(rows.fixed() && cols.fixed(), "{text}");
            let ((top, bottom), (left, right)) = (rows.bounds(), cols.bounds());
            assert_eq!(
                (top, left, bottom, right),
                (r.top, r.left, r.bottom, r.right),
                "{text}"
            );
        }
        for text in ["C3", "$C3", "B2:$D$4", "C:D", "3:4"] {
            let r = reference(text);
            let (rows, cols) = (Dim::rows(&r).unwrap(), Dim::cols(&r).unwrap());
            assert!(!(rows.fixed() && cols.fixed()), "{text}");
        }
    }
}
