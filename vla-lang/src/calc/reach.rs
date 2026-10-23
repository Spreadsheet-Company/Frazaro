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
//!
//! A changed cell is checked against the readers of its sheet, each shape's
//! references to it, and a sheet of [`INDEX_FROM`] readers or more keeps
//! them by where a changed cell must stand to reach them, wherever that at
//! least halves what a changed cell is checked against (KERNEL.35,
//! 2026-10-12): written one fill a row, Life's 198 shapes had every frame's
//! start check 2.5 million readers. A reader's footprint is the rows and the
//! columns of the read sheet it can reach from: a held rectangle's, a part's
//! cell, a moving reference's numbers moved over the offsets its shape's
//! cells stand at, each `$`-held number staying where it is. A sheet's
//! readers are filed by the rows or by the columns of their footprints,
//! whichever leaves a changed cell fewer to check, in bands one row or one
//! column wide, and a band that still holds [`INDEX_FROM`] or more is filed
//! again by the other axis. A changed cell is checked against the readers
//! too wide to file and against its band's, each exactly as the scan checks
//! it, so the index leaves out only readers that reach nothing from it.
//!
//! One text is one shape wherever its cells stand (KERNEL.34), so a shape's
//! box can span far more than its cells: a formula pasted where a model
//! needs it, or `CART.9`'s gates shuffled, one in forty sharing a text with
//! a gate far away, whose readers were too wide to file and checked against
//! every changed cell (KERNEL.54, 2026-10-21). A shape whose box spans more
//! than [`WIDE`] rows or columns and whose cells fill less than half of it is
//! cut into pieces by tiles of [`TILE`] rows and columns of each sheet, each
//! piece the box of the cells a tile holds; each moving reference of the
//! shape is filed and checked once a piece, its footprint read over the
//! piece's box and its cells found within it. A piece's cells are its
//! shape's, so what the readers of all its pieces reach is what the shape's
//! one reader reached.
//!
//! A few functions read a sheet's extent itself beside the cells they read
//! (KERNEL.56; the library declares them, `Library::reads_extent`): an
//! approximate match bisects its vector between the extent's first and last
//! places along it, and a database's records end at its sheet's last row. A
//! value written beyond the extent changes no cell such a formula reads and
//! still moves its answer, so every shape calling one is filed by each sheet
//! it reads, in one list of pairs, empty where no shape calls one; a sheet
//! whose extent moved has them computed whole.

use std::collections::hash_map::Entry;

use crate::fx::FxHashMap;
use crate::grid::Model;
use crate::refers::Shape as Part;
use crate::sheet::{MAX_COLUMN, MAX_ROW};

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

    /// The numbers the dimension holds at some offset from `d0` to `d1`, on
    /// the sheet (KERNEL.35); `None` where it holds none. A number its `$`
    /// holds stays where it is and one it does not moves with the offset,
    /// and the dimension holds what lies between its two, so the least and
    /// the most either takes over the offsets bound everything it holds at
    /// any of them.
    fn reach(self, (d0, d1): (i64, i64)) -> Option<Span> {
        let (n1, abs1, n2, abs2, max) = match self {
            Dim::Whole { max } => return Some((1, max)),
            Dim::Span {
                n1,
                abs1,
                n2,
                abs2,
                max,
            } => (n1, abs1, n2, abs2, max),
        };
        let ends = |n: u32, abs: bool| {
            let n = i64::from(n);
            if abs {
                (n, n)
            } else {
                (n + d0, n + d1)
            }
        };
        let ((a0, a1), (b0, b1)) = (ends(n1, abs1), ends(n2, abs2));
        let (lo, hi) = (a0.min(b0).max(1), a1.max(b1).min(i64::from(max)));
        (lo <= hi).then_some((lo as u32, hi as u32))
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
fn on(find: &dyn Fn(&str) -> Option<usize>, r: &RefExpr) -> Option<On> {
    match (r.found, &r.sheet) {
        (Some((first, _)), _) => Some(On::Sheet(usize::from(first))),
        (None, None) => Some(On::Own),
        (None, Some(name)) => find(name).map(On::Sheet),
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
    /// The stand its footprint is read over and its cells found within, by
    /// its index among the stands: its shape's, or one piece's of a shape
    /// cut into pieces (KERNEL.54).
    stand: u32,
}

/// Where a shape's cells stand, or a piece's of them (KERNEL.54).
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

/// The readers a sheet holds from which a changed cell of it is checked
/// against an index of them, and a band's from which its readers are filed
/// again by the other axis (KERNEL.35). Alonzo's cartridges hold at most 15
/// readers a sheet, so none of theirs is indexed.
const INDEX_FROM: usize = 16;

/// The most bands a reader is filed under: one whose footprint is wider is
/// checked against every changed cell of its sheet, as the scan checks it.
const WIDE: u32 = 64;

/// The rows and columns of a sheet's tiles, by which a shape whose cells
/// stand far apart is cut into pieces (KERNEL.54): a piece's box inside a
/// tile, so that a reference reaching 32 cells or fewer either way has a
/// footprint over it inside [`WIDE`].
const TILE: u32 = 32;

/// Whether a shape is cut into pieces (KERNEL.54): its cells' box spans more
/// than [`WIDE`] rows or columns, and its cells fill less than half of it,
/// the line between a fill with holes and cells that only share a text.
fn cut(stand: &Stand, cells: u64) -> bool {
    let (rows, cols) = (
        stand.rows.1 - stand.rows.0 + 1,
        stand.cols.1 - stand.cols.0 + 1,
    );
    let area = u128::from(rows.unsigned_abs())
        * u128::from(cols.unsigned_abs())
        * stand.sheets.len() as u128;
    (rows > i64::from(WIDE) || cols > i64::from(WIDE)) && u128::from(cells) * 2 < area
}

/// The pieces each shape is cut into (KERNEL.54), each a stand of its own
/// added after the shapes', listed by its shape: a shape [`cut`] finds far
/// apart has its cells bucketed by the tiles of each sheet, a piece the box
/// of the cells one tile holds; every other shape none, and so every shape
/// where a test asks the reach unpieced, the list then empty. The pieces
/// are gathered apart and the stands given room for them alone: pushed
/// onto the stands, the netlist's 992 doubled its 19,955 stands' room, 1.2
/// MB held for nothing.
fn pieces(
    stands: &mut Vec<Stand>,
    cells: &[u64],
    order: &[CellId],
    shape_at: &[u32],
) -> Vec<Vec<u32>> {
    let far: Vec<bool> = (0..cells.len())
        .map(|s| cells[s] > 0 && cut(&stands[s], cells[s]))
        .collect();
    if unpieced() || !far.contains(&true) {
        return Vec::new();
    }
    let mut pieces: Vec<Vec<u32>> = vec![Vec::new(); cells.len()];
    let first = stands.len();
    let mut added: Vec<Stand> = Vec::new();
    let mut piece_at: FxHashMap<(u32, usize, u32, u32), u32> = FxHashMap::default();
    for (&(sheet, row, col), &s) in order.iter().zip(shape_at) {
        if !far[s as usize] {
            continue;
        }
        let anchor = stands[s as usize].anchor;
        let (d_row, d_col) = (
            i64::from(row) - i64::from(anchor.0),
            i64::from(col) - i64::from(anchor.1),
        );
        match piece_at.entry((s, sheet, row / TILE, col / TILE)) {
            Entry::Occupied(at) => {
                let piece = &mut added[*at.get() as usize - first];
                piece.rows = (piece.rows.0.min(d_row), piece.rows.1.max(d_row));
                piece.cols = (piece.cols.0.min(d_col), piece.cols.1.max(d_col));
            }
            Entry::Vacant(at) => {
                let k = (first + added.len()) as u32;
                added.push(Stand {
                    sheets: vec![sheet],
                    anchor,
                    rows: (d_row, d_row),
                    cols: (d_col, d_col),
                });
                at.insert(k);
                pieces[s as usize].push(k);
            }
        }
    }
    stands.reserve_exact(added.len());
    stands.extend(added);
    pieces
}

/// Whether a test asks every reach built with no shape cut into pieces;
/// `false` outside a test.
#[cfg(not(test))]
fn unpieced() -> bool {
    false
}

#[cfg(test)]
fn unpieced() -> bool {
    whole::asked()
}

/// Every reach built in a test with no shape cut into pieces, each shape's
/// readers read over its whole stand as before KERNEL.54: the pieced index
/// held to the unpieced scan (KERNEL.54's oracle). On the thread that asked,
/// until the guard drops.
#[cfg(test)]
pub(crate) mod whole {
    use std::cell::Cell;

    thread_local! {
        static ASKED: Cell<bool> = const { Cell::new(false) };
    }

    /// Asked until dropped, then as before.
    pub(crate) struct Guard(bool);

    impl Drop for Guard {
        fn drop(&mut self) {
            ASKED.with(|a| a.set(self.0));
        }
    }

    /// Every reach built on this thread unpieced.
    pub(crate) fn stands() -> Guard {
        Guard(ASKED.with(|a| a.replace(true)))
    }

    pub(super) fn asked() -> bool {
        ASKED.with(Cell::get)
    }
}

/// The least and the most rows, or columns, a footprint holds.
type Span = (u32, u32);

/// The rows and the columns of the sheet a reader reads from which a changed
/// cell reaches a cell of its shape, or more (KERNEL.35); `None` where it
/// holds none. The scan checks a reader at any changed cell and finds what
/// it reaches; outside this, it finds nothing.
fn footprint(reader: &Reader, stand: &Stand) -> Option<(Span, Span)> {
    match reader.kind {
        Kind::Held {
            top,
            left,
            bottom,
            right,
        } => Some(((top, bottom), (left, right))),
        Kind::Parts { row, col } => Some(((row, row), (col, col))),
        Kind::Moving { rows, cols } => Some((rows.reach(stand.rows)?, cols.reach(stand.cols)?)),
    }
}

/// One axis's readers by band, a band one row or one column wide (KERNEL.35):
/// each reader's number filed under every band its footprint holds on the
/// axis, a band's together, in one array with its offsets, as a plan keeps
/// its edges (KERNEL.33); a reader whose footprint spans more than [`WIDE`]
/// bands, or none, checked against every changed cell.
#[derive(Clone, Debug)]
struct Bands {
    /// The readers checked against every changed cell, by number.
    wide: Vec<u32>,
    /// Every band's readers, a band's together, by number: slot `k`'s are
    /// `entries[starts[k]..starts[k + 1]]`.
    entries: Vec<u32>,
    starts: Vec<u32>,
    slots: Slots,
}

/// Which slot each band's readers are at.
#[derive(Clone, Debug)]
enum Slots {
    /// Every band from the least a reader holds to the most, band
    /// `first + k` at slot `k`.
    Flat { first: u32, len: u32 },
    /// The bands readers hold, where they stand too far apart for an array
    /// over every band between them, each at its slot by hash, as
    /// [`Positions`] keeps a sheet's formulas that stand apart: a narrow
    /// reader at row 1,048,576 beside readers at the top would ask the array
    /// for a million slots.
    Hash(FxHashMap<u32, u32>),
}

impl Bands {
    /// The readers `ids` filed by their spans on the axis, `span` giving
    /// each one's.
    fn of(ids: impl Iterator<Item = u32>, span: impl Fn(u32) -> Option<Span>) -> Bands {
        let mut wide = Vec::new();
        let mut filed: Vec<(u32, Span)> = Vec::new();
        let (mut least, mut most, mut held) = (u32::MAX, 0, 0u64);
        for i in ids {
            match span(i) {
                Some((lo, hi)) if hi - lo < WIDE => {
                    filed.push((i, (lo, hi)));
                    least = least.min(lo);
                    most = most.max(hi);
                    held += u64::from(hi - lo + 1);
                }
                _ => wide.push(i),
            }
        }
        // As a sheet's formulas' positions are kept: an array no more than
        // four times what it holds, so that two bands far apart are two
        // slots and not a million.
        let slots = if filed.is_empty() {
            Slots::Flat { first: 0, len: 0 }
        } else if u64::from(most - least) < 4 * held + 64 {
            Slots::Flat {
                first: least,
                len: most - least + 1,
            }
        } else {
            let mut bands: Vec<u32> = filed.iter().flat_map(|&(_, (lo, hi))| lo..=hi).collect();
            bands.sort_unstable();
            bands.dedup();
            Slots::Hash(
                bands
                    .into_iter()
                    .enumerate()
                    .map(|(k, n)| (n, k as u32))
                    .collect(),
            )
        };
        let count = match &slots {
            Slots::Flat { len, .. } => *len as usize,
            Slots::Hash(map) => map.len(),
        };
        let mut b = Bands {
            wide,
            entries: Vec::new(),
            starts: vec![0; count + 1],
            slots,
        };
        for &(_, (lo, hi)) in &filed {
            for n in lo..=hi {
                let k = b.slot(n).expect("a band filed has a slot");
                b.starts[k + 1] += 1;
            }
        }
        for k in 0..count {
            b.starts[k + 1] += b.starts[k];
        }
        let mut at: Vec<u32> = b.starts[..count].to_vec();
        b.entries = vec![0; b.starts[count] as usize];
        for &(i, (lo, hi)) in &filed {
            for n in lo..=hi {
                let k = b.slot(n).expect("a band filed has a slot");
                b.entries[at[k] as usize] = i;
                at[k] += 1;
            }
        }
        b
    }

    /// The slot of the band holding the number `n`, if a reader holds it.
    fn slot(&self, n: u32) -> Option<usize> {
        match &self.slots {
            Slots::Flat { first, len } => {
                let k = n.checked_sub(*first)?;
                (k < *len).then_some(k as usize)
            }
            Slots::Hash(map) => map.get(&n).map(|&k| k as usize),
        }
    }

    /// The readers of the band at slot `k`.
    fn list(&self, k: usize) -> &[u32] {
        &self.entries[self.starts[k] as usize..self.starts[k + 1] as usize]
    }

    /// How many slots there are.
    fn slots(&self) -> usize {
        self.starts.len() - 1
    }

    /// The readers a changed cell is checked against, on average over the
    /// bands that hold one: what the bands are chosen by.
    fn mean(&self) -> f64 {
        let held = (0..self.slots())
            .filter(|&k| !self.list(k).is_empty())
            .count();
        self.wide.len() as f64 + self.entries.len() as f64 / held.max(1) as f64
    }

    /// Each band's number and readers, by number, and the readers checked
    /// always: the index spelled (KERNEL.34's oracle, through `spelled`).
    #[cfg(test)]
    fn spelled(&self) -> String {
        let mut bands: Vec<(u32, usize)> = match &self.slots {
            Slots::Flat { first, len } => (0..*len).map(|k| (first + k, k as usize)).collect(),
            Slots::Hash(map) => map.iter().map(|(&n, &k)| (n, k as usize)).collect(),
        };
        bands.sort_unstable();
        let listed: Vec<(u32, &[u32])> = bands
            .into_iter()
            .map(|(n, k)| (n, self.list(k)))
            .filter(|(_, l)| !l.is_empty())
            .collect();
        format!("wide {:?} bands {listed:?}", self.wide)
    }
}

/// Whether an index whose changed cell meets `mean` readers on average at
/// least halves the `n` the scan checks: an index that does not costs more
/// than it saves, each reader one step further away.
fn halves(mean: f64, n: usize) -> bool {
    mean * 2.0 <= n as f64
}

/// A sheet's readers by where a changed cell stands (KERNEL.35): filed by
/// the axis that leaves a changed cell fewer to check, each band of it that
/// still holds [`INDEX_FROM`] or more filed again by the other.
#[derive(Clone, Debug)]
struct Index {
    /// Whether the first axis is the rows.
    by_rows: bool,
    /// The readers by band of the first axis.
    first: Bands,
    /// For each slot of the first axis, the place in `again` of its readers
    /// filed by the other axis, or [`NONE`]; empty when no band is.
    filed: Vec<u32>,
    again: Vec<Bands>,
}

impl Index {
    /// A sheet's readers indexed, or `None` where they are fewer than
    /// [`INDEX_FROM`] or no index at least halves what a changed cell is
    /// checked against.
    fn of(readers: &[Reader], stands: &[Stand]) -> Option<Index> {
        let force = forced();
        if readers.is_empty() || (force.is_none() && readers.len() < INDEX_FROM) {
            return None;
        }
        let spans: Vec<Option<(Span, Span)>> = readers
            .iter()
            .map(|r| footprint(r, &stands[r.stand as usize]))
            .collect();
        let all = 0..readers.len() as u32;
        let rows = Bands::of(all.clone(), |i| spans[i as usize].map(|s| s.0));
        let cols = Bands::of(all, |i| spans[i as usize].map(|s| s.1));
        let by_rows = match force {
            Some(Some(rows)) => rows,
            _ => rows.mean() <= cols.mean(),
        };
        let first = if by_rows { rows } else { cols };
        if force.is_none() && !halves(first.mean(), readers.len()) {
            return None;
        }
        let other = |i: u32| spans[i as usize].map(|s| if by_rows { s.1 } else { s.0 });
        let (mut filed, mut again) = (Vec::new(), Vec::new());
        for k in 0..first.slots() {
            let list = first.list(k);
            let many = match force {
                Some(_) => list.len() >= 2,
                None => list.len() >= INDEX_FROM,
            };
            if !many {
                continue;
            }
            let by_other = Bands::of(list.iter().copied(), other);
            if force.is_some() || halves(by_other.mean(), list.len()) {
                if filed.is_empty() {
                    filed = vec![NONE; first.slots()];
                }
                filed[k] = again.len() as u32;
                again.push(by_other);
            }
        }
        Some(Index {
            by_rows,
            first,
            filed,
            again,
        })
    }

    /// The readers of the first axis's band at slot `k` filed by the other
    /// axis, if they are.
    fn again(&self, k: usize) -> Option<&Bands> {
        match self.filed.get(k) {
            Some(&j) if j != NONE => Some(&self.again[j as usize]),
            _ => None,
        }
    }

    /// Each reader a cell changed at `row` and `col` is checked against, by
    /// its number among the sheet's readers, to `f`, each once.
    fn each(&self, row: u32, col: u32, f: &mut impl FnMut(u32)) {
        let (at, other) = if self.by_rows { (row, col) } else { (col, row) };
        for &i in &self.first.wide {
            f(i);
        }
        let Some(k) = self.first.slot(at) else {
            return;
        };
        match self.again(k) {
            Some(by_other) => {
                for &i in &by_other.wide {
                    f(i);
                }
                if let Some(m) = by_other.slot(other) {
                    for &i in by_other.list(m) {
                        f(i);
                    }
                }
            }
            None => {
                for &i in self.first.list(k) {
                    f(i);
                }
            }
        }
    }

    /// The most readers any one changed cell is checked against.
    fn most(&self) -> usize {
        let band = |k: usize| match self.again(k) {
            Some(b) => b.wide.len() + (0..b.slots()).map(|m| b.list(m).len()).max().unwrap_or(0),
            None => self.first.list(k).len(),
        };
        self.first.wide.len() + (0..self.first.slots()).map(band).max().unwrap_or(0)
    }

    /// The index spelled, each band's readers by number (KERNEL.34's oracle).
    #[cfg(test)]
    fn spelled(&self) -> String {
        let mut out = format!(
            "by {} {}",
            if self.by_rows { "rows" } else { "columns" },
            self.first.spelled()
        );
        let mut again: Vec<(u32, String)> = (0..self.first.slots())
            .filter_map(|k| {
                let n = match &self.first.slots {
                    Slots::Flat { first, .. } => first + k as u32,
                    Slots::Hash(map) => *map.iter().find(|&(_, &s)| s as usize == k)?.0,
                };
                Some((n, self.again(k)?.spelled()))
            })
            .collect();
        again.sort();
        for (n, b) in again {
            out.push_str(&format!(" again {n} {b}"));
        }
        out
    }
}

/// Whether a test asks every sheet indexed: `Some` with the first axis it
/// asks, `None` for the one an index finds; `None` outside a test.
#[cfg(not(test))]
fn forced() -> Option<Option<bool>> {
    None
}

#[cfg(test)]
fn forced() -> Option<Option<bool>> {
    every::asked()
}

/// Every sheet indexed in a test, whatever its readers' count and whether
/// its index halves them, every band of two readers or more filed again,
/// and the first axis the rows, the columns or the one an index finds:
/// the index held to the scan over grids whose sheets hold few readers
/// (KERNEL.35's oracle). On the thread that asked, until the guard drops.
#[cfg(test)]
pub(crate) mod every {
    use std::cell::Cell;

    thread_local! {
        static ASKED: Cell<Option<Option<bool>>> = const { Cell::new(None) };
    }

    /// Asked until dropped, then as before.
    pub(crate) struct Guard(Option<Option<bool>>);

    impl Drop for Guard {
        fn drop(&mut self) {
            ASKED.with(|a| a.set(self.0));
        }
    }

    /// Every sheet indexed, its first axis the rows (`Some(true)`), the
    /// columns (`Some(false)`) or the one an index finds (`None`).
    pub(crate) fn sheet_indexed(by_rows: Option<bool>) -> Guard {
        Guard(ASKED.with(|a| a.replace(Some(by_rows))))
    }

    pub(super) fn asked() -> Option<Option<bool>> {
        ASKED.with(Cell::get)
    }
}

/// The readers checked against changed cells, counted in a test
/// (KERNEL.35's oracle), as `crate::sheet::work` counts a write's work: each
/// reader of a changed cell's sheet checked, through the index or the scan.
/// Counted on the thread that asked, between [`checked::begin`] and
/// [`checked::end`], and only in a test build.
#[cfg(test)]
pub(crate) mod checked {
    use std::cell::Cell;

    thread_local! {
        static ON: Cell<bool> = const { Cell::new(false) };
        static COUNT: Cell<u64> = const { Cell::new(0) };
    }

    /// Counting begins on this thread, from nothing.
    pub(crate) fn begin() {
        ON.with(|on| on.set(true));
        COUNT.with(|c| c.set(0));
    }

    /// The readers checked since [`begin`]; counting ends.
    pub(crate) fn end() -> u64 {
        ON.with(|on| on.set(false));
        COUNT.with(Cell::get)
    }

    /// One reader checked.
    pub(super) fn one() {
        if ON.with(Cell::get) {
            COUNT.with(|c| c.set(c.get() + 1));
        }
    }
}

/// What each shape of a plan reads and where its cells stand, found when
/// the plan is: the cells a changed cell reaches. See the module's comment.
#[derive(Clone, Debug, Default)]
pub(crate) struct Reach {
    /// Each shape's stand by its index, then the pieces of the shapes cut
    /// into pieces (KERNEL.54).
    stands: Vec<Stand>,
    /// By sheet, the shapes that read it and how.
    readers: Vec<Vec<Reader>>,
    /// By sheet, its readers by where a changed cell stands, where they are
    /// many (KERNEL.35).
    index: Vec<Option<Index>>,
    /// The shapes no change is followed to, computed every frame.
    unfollowed: Vec<usize>,
    /// The shapes that read a sheet's extent itself (KERNEL.56), each by
    /// every sheet it reads, as pairs of the sheet and the shape, sorted and
    /// each once: a shape calling a function the library declares to read
    /// one, whose answer a value written beyond the extent moves where no cell
    /// it reads changed.
    extent_readers: Vec<(u32, u32)>,
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
    pub(crate) fn of<M: Model + ?Sized>(
        model: &M,
        shapes: &[Shape],
        order: &[CellId],
        shape_at: Vec<u32>,
        formula: &dyn Fn(CellId) -> bool,
    ) -> Reach {
        let find: &dyn Fn(&str) -> Option<usize> = &|name| model.find_sheet(name);
        let mut stands: Vec<Stand> = vec![Stand::default(); shapes.len()];
        let mut seen = vec![false; shapes.len()];
        let mut cells = vec![0u64; shapes.len()];
        for (&(sheet, row, col), &s) in order.iter().zip(&shape_at) {
            let s = s as usize;
            cells[s] += 1;
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
        for stand in &mut stands {
            stand.sheets.sort_unstable();
        }
        let pieces = pieces(&mut stands, &cells, order, &shape_at);
        let mut readers: Vec<Vec<Reader>> = vec![Vec::new(); model.sheet_count()];
        let mut extent_readers: Vec<(u32, u32)> = Vec::new();
        let mut unfollowed = Vec::new();
        let mut parts = Vec::new();
        let mut firsts = Vec::with_capacity(shapes.len() + 1);
        for (s, shape) in shapes.iter().enumerate() {
            let first = parts.len();
            firsts.push(first);
            if !seen[s] {
                continue;
            }
            let stand = &stands[s];
            let single = (stand.sheets.len() == 1).then(|| stand.sheets[0]);
            let mut found = Vec::new();
            let followed = match &shape.expr {
                Ok(e) if !shape.risky => {
                    let mut at = At {
                        find,
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
            // A shape calling a function that reads an extent, filed by every
            // sheet it reads (KERNEL.56), its own sheet's as its cells stand.
            if shape.reads_extent {
                for &(on, _) in &found {
                    match on {
                        On::Sheet(x) => extent_readers.push((x as u32, s as u32)),
                        On::Own => {
                            for &y in &stand.sheets {
                                extent_readers.push((y as u32, s as u32));
                            }
                        }
                    }
                }
            }
            // Single cells the shape reads side by side in one row, each
            // moving in both dimensions, read as one rectangle of that row,
            // which reaches the cells they reach (KERNEL.33): a changed cell
            // tries it once where it tried each, Sand's five neighbours of
            // `Screen.last` one pattern where they were five.
            merge_rows(&mut found, stand);
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
            // A moving reference of a shape cut into pieces is filed once a
            // piece (KERNEL.54); a held rectangle's and a part's footprints
            // are read over no cell, so each stays one a shape.
            let whole = [s as u32];
            for (k, (on, kind)) in found.into_iter().enumerate() {
                if inside[k] {
                    continue;
                }
                let filed: &[u32] = match (kind, pieces.get(s)) {
                    (Kind::Moving { .. }, Some(cut)) if !cut.is_empty() => cut,
                    _ => &whole,
                };
                for &at in filed {
                    match on {
                        On::Sheet(x) => {
                            if let Some(list) = readers.get_mut(x) {
                                list.push(Reader {
                                    shape: s,
                                    own: false,
                                    kind,
                                    stand: at,
                                });
                            }
                        }
                        On::Own => {
                            for &y in &stands[at as usize].sheets {
                                readers[y].push(Reader {
                                    shape: s,
                                    own: true,
                                    kind,
                                    stand: at,
                                });
                            }
                        }
                    }
                }
            }
        }
        firsts.push(parts.len());
        extent_readers.sort_unstable();
        extent_readers.dedup();
        let index = readers
            .iter()
            .map(|list| Index::of(list, &stands))
            .collect();
        Reach {
            stands,
            readers,
            index,
            unfollowed,
            extent_readers,
            parts,
            firsts,
            positions: Positions::of(order, model.sheet_count()),
            shape_at,
        }
    }

    /// Everything a changed cell reaches, to `f`; a cell of the order may
    /// come more than once. The readers it is checked against are its
    /// sheet's index's, where the sheet has one, and every reader of the
    /// sheet where it has none (KERNEL.35), each in no order that matters:
    /// what is reached marks a set.
    pub(crate) fn each(&self, changed: CellId, f: &mut impl FnMut(Reached)) {
        let (sheet, row, col) = changed;
        let Some(readers) = self.readers.get(sheet) else {
            return;
        };
        match self.index.get(sheet).and_then(Option::as_ref) {
            Some(index) => {
                index.each(row, col, &mut |i| {
                    self.check(&readers[i as usize], changed, f);
                });
            }
            None => {
                for reader in readers {
                    self.check(reader, changed, f);
                }
            }
        }
    }

    /// Everything a changed cell reaches, every reader of its sheet checked
    /// against it: the reach before KERNEL.35, kept as the index's reference.
    #[cfg(test)]
    pub(crate) fn each_by_scan(&self, changed: CellId, f: &mut impl FnMut(Reached)) {
        if let Some(readers) = self.readers.get(changed.0) {
            for reader in readers {
                self.check(reader, changed, f);
            }
        }
    }

    /// One reader of a changed cell's sheet checked against it: what it
    /// reaches, to `f`.
    fn check(&self, reader: &Reader, changed: CellId, f: &mut impl FnMut(Reached)) {
        #[cfg(test)]
        checked::one();
        let (sheet, row, col) = changed;
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
                // A piece's reader finds its shape's cells within the piece.
                let stand = &self.stands[reader.stand as usize];
                let (Some((r0, r1)), Some((c0, c1))) = (rows.offsets(row), cols.offsets(col))
                else {
                    return;
                };
                let (r0, r1) = (r0.max(stand.rows.0), r1.min(stand.rows.1));
                let (c0, c1) = (c0.max(stand.cols.0), c1.min(stand.cols.1));
                if r0 > r1 || c0 > c1 {
                    return;
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

    /// The shapes that read a sheet's extent itself (KERNEL.56), each once.
    pub(crate) fn extent_readers(&self, sheet: usize) -> impl Iterator<Item = usize> + '_ {
        let x = sheet as u32;
        let from = self.extent_readers.partition_point(|p| p.0 < x);
        self.extent_readers[from..]
            .iter()
            .take_while(move |p| p.0 == x)
            .map(|p| p.1 as usize)
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

    /// The reach spelled but for the positions, which the order makes
    /// (KERNEL.34's oracle, through `Shapes::spelled`), each sheet's index
    /// with it (KERNEL.35).
    #[cfg(test)]
    pub(crate) fn spelled(&self) -> String {
        let index: Vec<String> = self
            .index
            .iter()
            .enumerate()
            .filter_map(|(s, x)| Some(format!("{s} {}", x.as_ref()?.spelled())))
            .collect();
        format!(
            "stands {:?}\nreaders {:?}\nindex {index:?}\nunfollowed {:?}\nextent readers {:?}\nparts {:?}\nfirsts {:?}\nshape at {:?}\n",
            self.stands,
            self.readers,
            self.unfollowed,
            self.extent_readers,
            self.parts,
            self.firsts,
            self.shape_at
        )
    }

    /// How many pieces the shapes cut into pieces make (KERNEL.54).
    #[cfg(test)]
    pub(crate) fn pieces(&self) -> usize {
        self.stands.len() + 1 - self.firsts.len()
    }

    /// The readers of a sheet, each shape's references to it that reach
    /// other cells than another of its own does (KERNEL.30's census prints
    /// how many): what a changed cell of a sheet without an index is checked
    /// against, at a frame's start and at each cell the frame changes. A
    /// shape cut into pieces counts each moving reference once a piece
    /// (KERNEL.54), as a changed cell checks it.
    pub(crate) fn readers_on(&self, sheet: usize) -> usize {
        self.readers.get(sheet).map_or(0, Vec::len)
    }

    /// The most readers one changed cell of a sheet is checked against,
    /// where the sheet's readers are indexed (KERNEL.35, the census's
    /// figure); `None` for a sheet whose every reader is checked.
    pub(crate) fn most_on(&self, sheet: usize) -> Option<usize> {
        self.index.get(sheet)?.as_ref().map(Index::most)
    }
}

/// One shape's parse being read.
struct At<'a, 'p> {
    /// A sheet of the grid by name.
    find: &'a dyn Fn(&str) -> Option<usize>,
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
                    Expr::Range(r) => {
                        stack.push((&r.right, false));
                        stack.push((&r.left, false));
                    }
                    _ => {}
                }
                continue;
            }
            let operands = match e {
                Expr::Call(_, args) => args.len(),
                Expr::Neg(_) | Expr::Pos(_) | Expr::Percent(_) => 1,
                Expr::Binary(..) | Expr::Range(_) => 2,
                _ => 0,
            };
            let first = done.len() - operands;
            let same = match e {
                Expr::Number(_)
                | Expr::Text(_)
                | Expr::Bool(_)
                | Expr::Error(_)
                | Expr::Empty
                | Expr::Array(_) => true,
                Expr::Ref(r) => held(r),
                // A range is read by its reach, the cells any span it
                // answers may name (KERNEL.8, slice f).
                Expr::Range(_) => false,
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
                if let Expr::Range(r) = e {
                    match &r.reach {
                        Some(reach) => followed &= self.reference(reach),
                        None => followed = false,
                    }
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
        let (Some(on), Some(rows), Some(cols)) = (on(self.find, r), Dim::rows(r), Dim::cols(r))
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
        let find = self.find;
        u.each_read(&mut |read| {
            if let Read::Ref(r) = read {
                if let Some(on) = on(find, r) {
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

/// A shape's references to single cells of one sheet that move in both
/// dimensions and stand side by side in one row, two or more, merged into
/// one reference of the row's span (KERNEL.33). A changed cell reaches the
/// offsets at which some single holds it, and the span holds it at those
/// offsets and no other, so long as the span stays on the sheet: moved
/// partly off it, the span is `#REF!` where a single need not be. So the
/// merge is made only where every offset the shape's cells stand at keeps
/// the span's corners on the sheet, and [`Reach::each`] keeps to those
/// offsets.
fn merge_rows(found: &mut Vec<(On, Kind)>, stand: &Stand) {
    let key = |on: On| match on {
        On::Sheet(s) => s,
        On::Own => usize::MAX,
    };
    let mut singles: Vec<(usize, u32, u32, usize)> = Vec::new();
    for (i, &(on, kind)) in found.iter().enumerate() {
        if let Kind::Moving {
            rows:
                Dim::Span {
                    n1: r1,
                    abs1: false,
                    n2: r2,
                    abs2: false,
                    ..
                },
            cols:
                Dim::Span {
                    n1: c1,
                    abs1: false,
                    n2: c2,
                    abs2: false,
                    ..
                },
        } = kind
        {
            if r1 == r2 && c1 == c2 {
                singles.push((key(on), r1, c1, i));
            }
        }
    }
    singles.sort_unstable();
    let mut drop = vec![false; found.len()];
    let mut merged = Vec::new();
    let mut a = 0;
    while a < singles.len() {
        let mut b = a + 1;
        while b < singles.len()
            && singles[b].0 == singles[a].0
            && singles[b].1 == singles[a].1
            && singles[b].2 == singles[b - 1].2 + 1
        {
            b += 1;
        }
        let (row, lo, hi) = (singles[a].1, singles[a].2, singles[b - 1].2);
        let on_sheet = i64::from(lo) + stand.cols.0 >= 1
            && i64::from(hi) + stand.cols.1 <= i64::from(MAX_COLUMN)
            && i64::from(row) + stand.rows.0 >= 1
            && i64::from(row) + stand.rows.1 <= i64::from(MAX_ROW);
        if b - a >= 2 && on_sheet {
            for s in &singles[a..b] {
                drop[s.3] = true;
            }
            let on = found[singles[a].3].0;
            merged.push((
                on,
                Kind::Moving {
                    rows: Dim::Span {
                        n1: row,
                        abs1: false,
                        n2: row,
                        abs2: false,
                        max: MAX_ROW,
                    },
                    cols: Dim::Span {
                        n1: lo,
                        abs1: false,
                        n2: hi,
                        abs2: false,
                        max: MAX_COLUMN,
                    },
                },
            ));
        }
        a = b;
    }
    let mut k = 0;
    found.retain(|_| {
        let keep = !drop[k];
        k += 1;
        keep
    });
    found.extend(merged);
}

/// Whether a reference reads one cell every `$` holds, which a run reads
/// once (`super::vector`).
fn held(r: &RefExpr) -> bool {
    matches!(r.shape, Part::Cell)
        && r.last_sheet.is_none()
        && r.corners.row_abs1
        && r.corners.col_abs1
}

/// No position, and no band filed again.
const NONE: u32 = u32::MAX;

/// Each formula cell's position in the order: an array over a sheet's
/// formulas where they stand close enough together, the rest by hash.
#[derive(Clone, Debug, Default)]
pub(crate) struct Positions {
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
    pub(crate) fn of(order: &[CellId], sheets: usize) -> Positions {
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

    pub(crate) fn get(&self, sheet: usize, row: u32, col: u32) -> Option<usize> {
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

    /// Each position whose cell stands inside a rectangle of a sheet held as
    /// an array, to `f`, row by row; `false` for a sheet whose formulas are
    /// held by hash, which the caller walks instead (KERNEL.33: a plan finds
    /// the formula cells an area reads with it).
    pub(crate) fn each_in(
        &self,
        sheet: usize,
        (top, left, bottom, right): (u32, u32, u32, u32),
        f: &mut impl FnMut(usize),
    ) -> bool {
        let Some(Some(g)) = self.grids.get(sheet) else {
            return false;
        };
        let (t, l) = (top.max(g.top), left.max(g.left));
        let (b, r) = (
            bottom.min(g.top + g.height - 1),
            right.min(g.left + g.width - 1),
        );
        if t > b || l > r {
            return true;
        }
        for row in t..=b {
            let base = (row - g.top) as usize * g.width as usize;
            for col in l..=r {
                let at = g.at[base + (col - g.left) as usize];
                if at != NONE {
                    f(at as usize);
                }
            }
        }
        true
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
    fn a_dimension_reaches_every_number_its_moves_hold() {
        // KERNEL.35: each kind of reference and each `$` mark, over spans of
        // offsets around the sheet's first rows and columns, where a move
        // leaves it: every row and column `RefExpr::moved` gives at an offset
        // of the span lies within what `Dim::reach` names for the span; and
        // where no offset of the span moves it off the sheet, the least and
        // the most it names are held at some offset.
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
            for d0 in -6i64..=3 {
                for d1 in d0..=d0 + 5 {
                    for (dim, by_rows) in [(rows, true), (cols, false)] {
                        let held: Vec<Option<(u32, u32)>> = (d0..=d1)
                            .map(|d| match by_rows {
                                true => r.moved(d, 0).map(|(t, _, b, _)| (t, b)),
                                false => r.moved(0, d).map(|(_, l, _, rt)| (l, rt)),
                            })
                            .collect();
                        let named = dim.reach((d0, d1));
                        for &(lo, hi) in held.iter().flatten() {
                            assert!(
                                named.is_some_and(|(a, b)| a <= lo && hi <= b),
                                "{text} over {d0} to {d1}, {}: {lo} to {hi} outside {named:?}",
                                if by_rows { "rows" } else { "columns" }
                            );
                        }
                        if held.iter().all(Option::is_some) {
                            let least = held.iter().flatten().map(|h| h.0).min();
                            let most = held.iter().flatten().map(|h| h.1).max();
                            assert_eq!(
                                named,
                                least.zip(most),
                                "{text} over {d0} to {d1}, {}",
                                if by_rows { "rows" } else { "columns" }
                            );
                        }
                    }
                }
            }
        }
    }

    #[test]
    fn bands_list_each_reader_whose_span_holds_a_number_lying_close_or_apart() {
        // KERNEL.35: readers of spans of every width, some none and some too
        // wide to file, filed by band: each number's band lists exactly the
        // readers whose span holds it and is narrow enough, in order, each
        // once, and the rest are the ones checked always. Spans lying close
        // are kept in an array over their bands, and spans one of which
        // stands at the sheet's last row in a hash, which lists the same.
        let mut seed = 0x35u64;
        let mut next = |n: u32| {
            seed = seed
                .wrapping_mul(6_364_136_223_846_793_005)
                .wrapping_add(1_442_695_040_888_963_407);
            ((seed >> 33) % u64::from(n)) as u32
        };
        let mut hashed = 0;
        for case in 0..60 {
            let far = case % 3 == 0;
            let spans: Vec<Option<Span>> = (0..40)
                .map(|i| match next(10) {
                    0 => None,
                    1 => Some((1 + next(30), 200 + next(50))),
                    _ if far && i == 7 => Some((MAX_ROW - 1, MAX_ROW)),
                    _ => {
                        let lo = 1 + next(40);
                        Some((lo, lo + next(6)))
                    }
                })
                .collect();
            let filed = |i: usize| spans[i].filter(|&(lo, hi)| hi - lo < WIDE);
            let apart = (0..40).any(|i| filed(i).is_some_and(|(lo, _)| lo > 1_000));
            let bands = Bands::of(0..40, |i| spans[i as usize]);
            assert_eq!(
                matches!(bands.slots, Slots::Hash(_)),
                apart,
                "case {case}: {spans:?}"
            );
            hashed += usize::from(apart);
            let wide: Vec<u32> = (0..40)
                .filter(|&i| filed(i).is_none())
                .map(|i| i as u32)
                .collect();
            assert_eq!(bands.wide, wide, "case {case}");
            let probes = (1..=100).chain([MAX_ROW - 2, MAX_ROW - 1, MAX_ROW]);
            for n in probes {
                let want: Vec<u32> = (0..40)
                    .filter(|&i| filed(i).is_some_and(|(lo, hi)| lo <= n && n <= hi))
                    .map(|i| i as u32)
                    .collect();
                let got = bands.slot(n).map_or(&[][..], |k| bands.list(k));
                assert_eq!(got, want.as_slice(), "case {case}, band {n}");
            }
        }
        assert!(hashed > 5, "{hashed} cases kept their bands by hash");
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
