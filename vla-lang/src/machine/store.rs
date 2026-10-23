//! A rows' sheet's values kept by place (KERNEL.49's decision 3,
//! 2026-10-21): every value its cells hold, a value cell's and a formula's of
//! the frame alike, in one store, so that a read of any of its cells is an
//! index where a formula cell's was a position found in the plan's grid first,
//! and the frame's values are kept once. A store is an array over the sheet's
//! extent ([`Dense`], nine bytes a slot) where the sheet's cells stand close
//! enough together, by the rule the arrays have kept since KERNEL.20
//! (`dense::limit`); and where they do not, a sheet too sparse for one array,
//! each column whose cells stand close enough together down it is an array of
//! its own, a run, and every other cell is kept by place in an ordered map,
//! so that a sheet of two cells far apart is not a million empty slots, and a
//! ledger's columns beside a summary far to their right are read by index
//! still (the criteria bench's sheet, 61,371 cells over 600,030 places, was
//! read 59% slower through a map of every cell and held 2.8 MB more).
//!
//! A slot holds what a read of its cell answers: a value, a formula's reason
//! for not being computed, which a read answers and a walk meets where it
//! stands, or a formula's lack of a value before it is first computed. The
//! machine writes a formula's slot where the frame keeps its value, a value
//! cell's where a write changes it, and makes an array again over a sheet's
//! extent where a write moves it, its slots moved by place.
//!
//! A twin a formula reads keeps a store of its own beside its sheet's
//! (decision 5), its sheet's store's twin as the last frame's end left it,
//! a formula with no value holding nothing: the frame's end copies each
//! changed cell's slot into it where the two stand beside each other, and
//! makes it again from the sheet's store where they do not.

use std::borrow::Cow;
use std::collections::{btree_map, BTreeMap};

use crate::calc::dense::{self, Dense, Slot};
use crate::calc::graph::Extent;
use crate::calc::{Computed, Reason, Value, Walked};
use crate::grid::Constant;
use crate::sheet::{A1Range, Content};

/// The fewest cells a column of a sparse sheet keeps as a run of its own,
/// an array down it: fewer stand loose in the map, a run's own room being
/// worth it only past a few cells.
const RUN: u32 = 8;

/// What a cell kept loose holds, owned: a value, an error a value cell holds
/// by a text its kind does not spell (KERNEL.49's third slice), a formula's
/// reason, or a formula with no value yet.
#[derive(Clone, Debug, PartialEq)]
pub(crate) enum Kept {
    Value(Value),
    Odd(String),
    Reason(Reason),
    Pending,
}

impl Kept {
    fn slot(&self) -> Slot<'_> {
        match self {
            Kept::Value(Value::Number(n)) => Slot::Number(*n),
            Kept::Value(Value::Text(t)) => Slot::Text(t),
            Kept::Value(Value::Bool(b)) => Slot::Bool(*b),
            Kept::Value(Value::Error(k)) => Slot::Error(*k),
            Kept::Value(Value::Empty) => Slot::Empty,
            Kept::Odd(t) => Slot::Error(dense::error_named(t)),
            Kept::Reason(r) => Slot::Reason(r),
            Kept::Pending => Slot::Pending,
        }
    }

    /// The value cell's content it is, exactly; `None` for a formula's reason
    /// or lack of a value.
    fn content(&self) -> Option<Content> {
        match self {
            Kept::Value(Value::Number(n)) => Some(Content::Number(*n)),
            Kept::Value(Value::Text(t)) => Some(Content::Text(t.clone())),
            Kept::Value(Value::Bool(b)) => Some(Content::Bool(*b)),
            Kept::Value(Value::Error(k)) => Some(Content::Error(k.text().to_string())),
            Kept::Odd(t) => Some(Content::Error(t.clone())),
            Kept::Value(Value::Empty) | Kept::Reason(_) | Kept::Pending => None,
        }
    }

    /// What an array's slot at an index holds, owned, an error kept by its
    /// text kept so.
    fn of(d: &Dense, i: usize) -> Kept {
        if let Some(t) = d.odd_text(i) {
            return Kept::Odd(t.to_string());
        }
        match d.slot(i) {
            Slot::Reason(r) => Kept::Reason(r.clone()),
            Slot::Pending => Kept::Pending,
            other => Kept::Value(other.value()),
        }
    }
}

/// A sheet too sparse for one array: its runs, each a column's array over
/// the rows from its first cell to its last, by column; and every other cell
/// by place.
#[derive(Clone, Debug, Default)]
pub(crate) struct Sparse {
    runs: Vec<(u32, Dense)>,
    loose: BTreeMap<(u32, u32), Kept>,
}

impl Sparse {
    /// The runs of a sheet whose cells are these: each column of at least
    /// [`RUN`] cells whose rows from its first to its last are no more than
    /// `dense::SLOTS_A_CELL` a cell, every slot empty.
    fn planned(cells: impl Iterator<Item = (u32, u32)>) -> Sparse {
        let mut columns: BTreeMap<u32, (u32, u32, u32)> = BTreeMap::new();
        for (row, col) in cells {
            let c = columns.entry(col).or_insert((0, row, row));
            c.0 += 1;
            c.1 = c.1.min(row);
            c.2 = c.2.max(row);
        }
        let runs = columns
            .into_iter()
            .filter(|&(_, (n, top, bottom))| {
                n >= RUN && u64::from(bottom - top + 1) <= u64::from(n) * dense::SLOTS_A_CELL
            })
            .filter_map(|(col, (_, top, bottom))| {
                Some((
                    col,
                    Dense::over(Some(((top, col), (bottom, col))), u64::MAX, None)?,
                ))
            })
            .collect();
        Sparse {
            runs,
            loose: BTreeMap::new(),
        }
    }

    /// The run of a column holding a row, and the row's slot in it.
    #[inline]
    fn run(&self, row: u32, col: u32) -> Option<(&Dense, usize)> {
        let k = self.runs.binary_search_by_key(&col, |r| r.0).ok()?;
        let d = &self.runs[k].1;
        Some((d, d.index(row, col)?))
    }

    #[inline]
    fn run_mut(&mut self, row: u32, col: u32) -> Option<(&mut Dense, usize)> {
        let k = self.runs.binary_search_by_key(&col, |r| r.0).ok()?;
        let d = &mut self.runs[k].1;
        let i = d.index(row, col)?;
        Some((d, i))
    }
}

/// The places of a sparse sheet's cells inside an area, row by row,
/// gathering nothing, as every walk is read (KERNEL.25): at each row, the
/// runs of the area's columns that hold a cell of it, by column, merged with
/// the row's loose cells; a row that no run reaches and no loose cell stands
/// in is skipped. A row costs the area's runs, a walk the rows its runs reach
/// and the loose cells of its rows.
pub(crate) struct SparseWalk<'a> {
    /// The runs of the area's columns, by column.
    runs: &'a [(u32, Dense)],
    /// The loose cells of the area's rows, row by row, and the next of them
    /// inside its columns.
    loose: btree_map::Range<'a, (u32, u32), Kept>,
    next_loose: Option<(u32, u32)>,
    left: u32,
    right: u32,
    bottom: u32,
    /// The row walked, and the next of the runs to try at it.
    row: u32,
    k: usize,
}

impl<'a> SparseWalk<'a> {
    fn new(s: &'a Sparse, area: &A1Range) -> SparseWalk<'a> {
        let first = s.runs.partition_point(|r| r.0 < area.left);
        let n = s.runs[first..].partition_point(|r| r.0 <= area.right);
        let mut walk = SparseWalk {
            runs: &s.runs[first..first + n],
            loose: s
                .loose
                .range((area.top, area.left)..=(area.bottom, area.right)),
            next_loose: None,
            left: area.left,
            right: area.right,
            bottom: area.bottom,
            row: area.top,
            k: 0,
        };
        walk.next_loose = walk.loose_inside();
        walk
    }

    /// The next loose cell inside the area's columns.
    fn loose_inside(&mut self) -> Option<(u32, u32)> {
        let (left, right) = (self.left, self.right);
        self.loose
            .by_ref()
            .map(|(&at, _)| at)
            .find(|&(_, col)| col >= left && col <= right)
    }

    /// The first row past the one walked that a run reaches or a loose cell
    /// stands in.
    fn next_row(&self) -> u32 {
        let after = self.row + 1;
        let mut next = self.next_loose.map_or(u32::MAX, |(row, _)| row);
        for (_, d) in self.runs {
            if let Some(((top, _), (bottom, _))) = d.extent() {
                if bottom >= after {
                    next = next.min(top.max(after));
                }
            }
        }
        next
    }
}

impl Iterator for SparseWalk<'_> {
    type Item = (u32, u32);

    fn next(&mut self) -> Option<(u32, u32)> {
        let runs = self.runs;
        while self.row <= self.bottom {
            let loose = self.next_loose.filter(|&(row, _)| row == self.row);
            while let Some((col, d)) = runs.get(self.k) {
                if d.holds(self.row, *col) {
                    if let Some(at) = loose.filter(|&(_, c)| c < *col) {
                        self.next_loose = self.loose_inside();
                        return Some(at);
                    }
                    self.k += 1;
                    return Some((self.row, *col));
                }
                self.k += 1;
            }
            if let Some(at) = loose {
                self.next_loose = self.loose_inside();
                return Some(at);
            }
            self.row = self.next_row();
            self.k = 0;
        }
        None
    }
}

/// A rows' sheet's values. See the module's comment.
#[derive(Clone, Debug)]
pub(crate) enum Store {
    /// Over the sheet's extent, a slot a cell of it.
    Array(Dense),
    /// A sheet too sparse for one array.
    Sparse(Sparse),
}

impl Store {
    /// A store over an extent, every slot empty: an array within `limit`
    /// slots, else runs planned over the cells to come.
    pub(crate) fn over(
        extent: Extent,
        limit: u64,
        cells: impl Iterator<Item = (u32, u32)>,
    ) -> Store {
        match Dense::over(extent, limit, None) {
            Some(d) => Store::Array(d),
            None => Store::Sparse(Sparse::planned(cells)),
        }
    }

    /// The same slots over another extent, an array within `limit` slots, or
    /// runs planned over the cells held: what a write leaves when it moves a
    /// sheet's extent, each slot moved by its place. An array that stands
    /// over the extent, and a sparse sheet that is sparse still, stay as they
    /// are, a cell outside a run's rows kept loose; an array within the limit
    /// over an extent that holds its own is taken over it where it stands,
    /// rows added below it in place (KERNEL.56, [`Dense::regrow`]), where
    /// every slot it held was read out and put again.
    pub(crate) fn moved(mut self, extent: Extent, limit: u64) -> Store {
        let stands = match &mut self {
            Store::Array(d) => {
                d.stands_over(extent) || (dense::fits(extent, limit) && d.regrow(extent))
            }
            Store::Sparse(_) => !dense::fits(extent, limit),
        };
        if stands {
            return self;
        }
        let held = self.held();
        #[cfg(test)]
        crate::sheet::work::relaid(held.len());
        let mut next = Store::over(extent, limit, held.iter().map(|&(at, _)| at));
        for ((row, col), kept) in &held {
            next.put_kept(*row, *col, kept);
        }
        next
    }

    /// Every cell the store holds, by place, owned, each exactly.
    fn held(&self) -> Vec<((u32, u32), Kept)> {
        let mut out = Vec::new();
        let mut array = |d: &Dense| {
            for i in 0..d.slots() {
                let (slot, at) = d.at(i);
                if !slot.is_empty() {
                    out.push((at, Kept::of(d, i)));
                }
            }
        };
        match self {
            Store::Array(d) => array(d),
            Store::Sparse(s) => {
                for (_, d) in &s.runs {
                    array(d);
                }
                for (&at, kept) in &s.loose {
                    out.push((at, kept.clone()));
                }
            }
        }
        out.sort_by_key(|&(at, _)| at);
        out
    }

    /// The store as its sheet's twin keeps it (KERNEL.49's decision 5): the
    /// same slots, each formula with no value yet holding nothing, as its
    /// twin cell does, so that a read of the twin answers nothing there.
    pub(crate) fn twin(&self) -> Store {
        let mut twin = self.clone();
        match &mut twin {
            Store::Array(d) => d.pending_emptied(),
            Store::Sparse(s) => {
                for (_, d) in &mut s.runs {
                    d.pending_emptied();
                }
                s.loose.retain(|_, kept| *kept != Kept::Pending);
            }
        }
        twin
    }

    /// Whether a twin's store holds its slots where its sheet's store holds
    /// them (KERNEL.49), two arrays over one extent or two sparse sheets'
    /// stores, each written by place, so that a frame's end copies a changed
    /// cell's slot into the twin's, where a store over another extent is made
    /// again; a twin's array taken over its sheet's first where a write grew
    /// the sheet's over an extent holding the twin's own (KERNEL.56), its
    /// slots moved to their places and rows added below in place as the
    /// sheet's were, where the twin was made again from the sheet's whole.
    pub(crate) fn follows(&mut self, sheet: &Store) -> bool {
        match (self, sheet) {
            (Store::Array(t), Store::Array(d)) => t.same_extent(d) || t.regrow(d.extent()),
            (Store::Sparse(_), Store::Sparse(_)) => true,
            _ => false,
        }
    }

    /// The array, where the store is one.
    #[inline]
    pub(crate) fn array(&self) -> Option<&Dense> {
        match self {
            Store::Array(d) => Some(d),
            Store::Sparse(_) => None,
        }
    }

    /// A cell's slot where it stands; `None` for a cell outside the array,
    /// or held nowhere in a sparse sheet, which holds nothing.
    #[inline]
    pub(crate) fn slot_of(&self, row: u32, col: u32) -> Option<Slot<'_>> {
        match self {
            Store::Array(d) => d.slot_of(row, col),
            Store::Sparse(s) => match s.run(row, col) {
                Some((d, i)) => Some(d.slot(i)),
                None => s.loose.get(&(row, col)).map(Kept::slot),
            },
        }
    }

    /// A cell's slot written with a number.
    #[inline]
    pub(crate) fn put_number(&mut self, row: u32, col: u32, n: f64) {
        match self {
            Store::Array(d) => {
                if let Some(i) = d.index(row, col) {
                    d.put_number(i, n);
                }
            }
            Store::Sparse(s) => match s.run_mut(row, col) {
                Some((d, i)) => d.put_number(i, n),
                None => {
                    s.loose.insert((row, col), Kept::Value(Value::Number(n)));
                }
            },
        }
    }

    /// A formula cell's slot written with its computed value: its value, its
    /// reason, or none yet.
    pub(crate) fn put_computed(&mut self, row: u32, col: u32, computed: Option<&Computed>) {
        match self {
            Store::Array(d) => {
                if let Some(i) = d.index(row, col) {
                    d.put_computed(i, computed);
                }
            }
            Store::Sparse(s) => match s.run_mut(row, col) {
                Some((d, i)) => d.put_computed(i, computed),
                None => {
                    let kept = match computed {
                        Some(Computed::Value(v)) => Kept::Value(v.clone()),
                        Some(Computed::NotComputed(r)) => Kept::Reason(r.clone()),
                        None => Kept::Pending,
                    };
                    s.loose.insert((row, col), kept);
                }
            },
        }
    }

    /// A value cell's slot written with its content: a value as it stands,
    /// an error by its kind where its text spells it and else by its text,
    /// which a read answers as its kind (KERNEL.49's third slice: the slot is
    /// the cell); nothing for no cell or a content that holds no value.
    pub(crate) fn put_content(&mut self, row: u32, col: u32, content: Option<&Content>) {
        match self {
            Store::Array(d) => {
                if let Some(i) = d.index(row, col) {
                    d.put_content(i, content);
                }
            }
            Store::Sparse(s) => match s.run_mut(row, col) {
                Some((d, i)) => d.put_content(i, content),
                None => {
                    let kept = match content {
                        Some(Content::Number(n)) => Kept::Value(Value::Number(*n)),
                        Some(Content::Text(t)) => Kept::Value(Value::Text(t.clone())),
                        Some(Content::Bool(b)) => Kept::Value(Value::Bool(*b)),
                        Some(Content::Error(e)) if dense::spells_kind(e) => {
                            Kept::Value(Value::Error(dense::error_named(e)))
                        }
                        Some(Content::Error(e)) => Kept::Odd(e.clone()),
                        _ => {
                            s.loose.remove(&(row, col));
                            return;
                        }
                    };
                    s.loose.insert((row, col), kept);
                }
            },
        }
    }

    /// A value cell's slot written with its content where it stands in
    /// another store (KERNEL.49), a text written into the place its slot
    /// names already, so that a frame's end copying a host's text into a
    /// twin allocates nothing where the twin held a text.
    pub(crate) fn put_constant(&mut self, row: u32, col: u32, c: &Constant<'_>) {
        match c {
            Constant::Number(n) => self.put_number(row, col, *n),
            Constant::Bool(b) => self.put_slot(row, col, Slot::Bool(*b)),
            Constant::Error(e) if dense::spells_kind(e) => {
                self.put_slot(row, col, Slot::Error(dense::error_named(e)));
            }
            Constant::Text(t) => match self {
                Store::Array(d) => {
                    if let Some(i) = d.index(row, col) {
                        d.put_text(i, t);
                    }
                }
                Store::Sparse(s) => match s.run_mut(row, col) {
                    Some((d, i)) => d.put_text(i, t),
                    None => {
                        s.loose
                            .insert((row, col), Kept::Value(Value::Text(t.to_string())));
                    }
                },
            },
            Constant::Error(e) => self.put_content(row, col, Some(&Content::Error(e.to_string()))),
        }
    }

    /// A cell written with what another store held there, owned, an error
    /// kept by its text kept so.
    fn put_kept(&mut self, row: u32, col: u32, kept: &Kept) {
        match kept {
            Kept::Odd(t) => self.put_content(row, col, Some(&Content::Error(t.clone()))),
            other => self.put_slot(row, col, other.slot()),
        }
    }

    /// Every cell the store holds inside an area, row by row, its place and
    /// its slot (KERNEL.49's third slice: a rows' sheet's store holds a slot
    /// at every cell of its sheet, a value cell's content, a formula's
    /// value, its reason or its lack of one, and nothing elsewhere).
    pub(crate) fn each_held(&self, area: &A1Range, f: &mut dyn FnMut(u32, u32, Slot<'_>)) {
        match self {
            Store::Array(d) => {
                for (row, col) in d.walk(area.top, area.left, area.bottom, area.right) {
                    if let Some(slot) = d.slot_of(row, col) {
                        f(row, col, slot);
                    }
                }
            }
            Store::Sparse(s) => {
                for (row, col) in SparseWalk::new(s, area) {
                    if let Some(slot) = self.slot_of(row, col) {
                        f(row, col, slot);
                    }
                }
            }
        }
    }

    /// The places of the cells the store holds inside an area, row by row:
    /// what the evaluator's walk of the area visits ([`crate::calc::Env::walk`]).
    pub(crate) fn walk(&self, area: &A1Range) -> Walked<'_> {
        match self {
            Store::Array(d) => Walked::slots(d.walk(area.top, area.left, area.bottom, area.right)),
            Store::Sparse(s) => Walked::sparse(SparseWalk::new(s, area)),
        }
    }

    /// How many cells the store holds inside an area.
    pub(crate) fn count_in(&self, area: &A1Range) -> u64 {
        match self {
            Store::Array(d) => d.walk(area.top, area.left, area.bottom, area.right).count() as u64,
            Store::Sparse(s) => SparseWalk::new(s, area).count() as u64,
        }
    }

    /// A value cell's content as the record reads it, lent where it stands,
    /// an error by the text it was written with; `None` for a cell held
    /// nowhere, a formula's reason and a formula with no value yet.
    pub(crate) fn constant(&self, row: u32, col: u32) -> Option<Constant<'_>> {
        let (slot, odd) = match self {
            Store::Array(d) => {
                let i = d.index(row, col)?;
                (d.slot(i), d.odd_text(i))
            }
            Store::Sparse(s) => match s.run(row, col) {
                Some((d, i)) => (d.slot(i), d.odd_text(i)),
                None => match s.loose.get(&(row, col))? {
                    Kept::Odd(t) => return Some(Constant::Error(Cow::Borrowed(t))),
                    kept => (kept.slot(), None),
                },
            },
        };
        Some(match slot {
            Slot::Number(n) => Constant::Number(n),
            Slot::Text(t) => Constant::Text(Cow::Borrowed(t)),
            Slot::Bool(b) => Constant::Bool(b),
            Slot::Error(k) => Constant::Error(Cow::Borrowed(odd.unwrap_or(k.text()))),
            Slot::Empty | Slot::Reason(_) | Slot::Pending => return None,
        })
    }

    /// The content of a value cell, exactly as it was written (KERNEL.49's
    /// third slice, where the store is the cell); `None` for a cell held
    /// nowhere, a formula's reason and a formula with no value yet. A
    /// formula's value reads back as a value cell of it would hold it.
    pub(crate) fn content(&self, row: u32, col: u32) -> Option<Content> {
        match self {
            Store::Array(d) => d.content(d.index(row, col)?),
            Store::Sparse(s) => match s.run(row, col) {
                Some((d, i)) => d.content(i),
                None => s.loose.get(&(row, col))?.content(),
            },
        }
    }

    /// A cell's slot written with another's.
    pub(crate) fn put_slot(&mut self, row: u32, col: u32, slot: Slot<'_>) {
        match slot {
            Slot::Number(n) => self.put_number(row, col, n),
            Slot::Reason(r) => self.put_computed(row, col, Some(&Computed::NotComputed(r.clone()))),
            Slot::Pending => self.put_computed(row, col, None),
            Slot::Empty => self.put_content(row, col, None),
            other => self.put_computed(row, col, Some(&Computed::Value(other.value()))),
        }
    }
}

#[cfg(test)]
impl Store {
    /// Whether two stores hold their slots by the same places, two arrays
    /// over one extent or two sparse sheets' stores (KERNEL.49): the tests'.
    pub(crate) fn beside(&self, other: &Store) -> bool {
        match (self, other) {
            (Store::Array(a), Store::Array(b)) => a.same_extent(b),
            (Store::Sparse(_), Store::Sparse(_)) => true,
            _ => false,
        }
    }

    /// Whether two stores hold the same slots by place, a number to the bit,
    /// wherever and however each keeps them (KERNEL.49): the tests'.
    pub(crate) fn same(&self, other: &Store) -> bool {
        let (a, b) = (self.held(), other.held());
        a.len() == b.len()
            && a.iter()
                .zip(&b)
                .all(|((p, x), (q, y))| p == q && x.slot().same(y.slot()))
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::calc::ErrorKind;

    #[test]
    fn a_sparse_sheet_keeps_its_dense_columns_as_runs_and_the_rest_loose() {
        // KERNEL.49: a ledger's three columns of twenty rows and a summary
        // far to their right, too sparse for one array: the ledger's columns
        // each a run, the summary's column of three cells loose; every cell
        // read back where it was put, a cell outside a run's rows loose, and
        // the store moved to an array once the cells fill its extent.
        let mut cells: Vec<(u32, u32)> = Vec::new();
        for row in 2..=21 {
            for col in 1..=3 {
                cells.push((row, col));
            }
        }
        cells.extend([(1, 100), (2, 100), (3, 100)]);
        let extent = Some(((1, 1), (21, 100)));
        let limit = dense::limit(cells.len(), 1 << 20);
        let mut store = Store::over(extent, limit, cells.iter().copied());
        let Store::Sparse(s) = &store else {
            panic!("a sparse sheet")
        };
        assert_eq!(
            s.runs.iter().map(|r| r.0).collect::<Vec<_>>(),
            vec![1, 2, 3],
            "the ledger's columns"
        );
        for &(row, col) in &cells {
            store.put_number(row, col, f64::from(row * 1000 + col));
        }
        store.put_computed(30, 2, Some(&Computed::NotComputed(Reason::Cycle)));
        for &(row, col) in &cells {
            assert!(
                matches!(store.slot_of(row, col), Some(Slot::Number(n)) if n == f64::from(row * 1000 + col)),
                "({row}, {col})"
            );
        }
        assert!(
            matches!(store.slot_of(30, 2), Some(Slot::Reason(_))),
            "loose, past its run"
        );
        assert!(store.slot_of(5, 50).is_none(), "nothing held");
        let Store::Sparse(s) = &store else {
            panic!("sparse still")
        };
        assert_eq!(s.loose.len(), 4, "the summary and the cell past the run");
        // Moved over an extent its cells now fill: an array, every slot by
        // its place.
        let filled = store.moved(Some(((1, 1), (30, 100))), 3000);
        let Store::Array(d) = &filled else {
            panic!("an array")
        };
        assert!(matches!(d.slot_of(30, 2), Some(Slot::Reason(_))));
        assert!(matches!(d.slot_of(3, 100), Some(Slot::Number(n)) if n == 3100.0));
        assert!(matches!(d.slot_of(21, 3), Some(Slot::Number(n)) if n == 21003.0));
    }

    #[test]
    fn a_value_cell_reads_back_exactly_loose_in_a_run_and_moved() {
        // KERNEL.49's third slice: a value cell's content comes back as it
        // was written wherever its store keeps it, an error by the text it
        // was written with; and a store moved to an array keeps it so.
        let cells: Vec<(u32, u32)> = (1..=10).map(|row| (row, 1)).chain([(1, 900)]).collect();
        let limit = dense::limit(cells.len(), 1 << 20);
        let mut store = Store::over(Some(((1, 1), (10, 900))), limit, cells.iter().copied());
        assert!(matches!(store, Store::Sparse(_)));
        let odd = Content::Error("#n/a".to_string());
        store.put_content(1, 900, Some(&odd));
        store.put_content(2, 1, Some(&odd));
        store.put_content(3, 1, Some(&Content::Error("#N/A".to_string())));
        store.put_content(4, 1, Some(&Content::Text("t".to_string())));
        assert_eq!(store.content(1, 900), Some(odd.clone()), "loose");
        assert_eq!(store.content(2, 1), Some(odd.clone()), "in a run");
        assert!(matches!(
            store.slot_of(1, 900),
            Some(Slot::Error(ErrorKind::NA))
        ));
        assert_eq!(store.content(5, 1), None, "held nowhere");
        store.put_computed(6, 1, None);
        assert_eq!(store.content(6, 1), None, "a formula with no value");
        let moved = store.moved(Some(((1, 1), (10, 900))), 9000);
        assert!(matches!(moved, Store::Array(_)));
        assert_eq!(moved.content(1, 900), Some(odd.clone()));
        assert_eq!(moved.content(2, 1), Some(odd));
        assert_eq!(
            moved.content(3, 1),
            Some(Content::Error("#N/A".to_string()))
        );
        assert_eq!(moved.content(4, 1), Some(Content::Text("t".to_string())));
        assert!(matches!(moved.slot_of(6, 1), Some(Slot::Pending)));
    }

    #[test]
    fn a_sparse_sheet_is_walked_row_by_row_gathering_nothing() {
        // KERNEL.49's third slice: a sparse sheet's walk merges its runs and
        // its loose cells row by row, gathering nothing (KERNEL.25), held to
        // every place of the area read in order: runs whose rows stand apart,
        // one reaching far below its cells, loose cells between, beside and
        // far from them, and slots emptied inside a run.
        let mut cells: Vec<(u32, u32)> = Vec::new();
        for row in 1..=12 {
            cells.extend([(row, 2), (row, 3)]);
        }
        cells.push((40, 2));
        for row in 20..=31 {
            cells.push((row, 4));
        }
        for row in 5..=16 {
            cells.push((row, 7));
        }
        cells.extend([(3, 1), (3, 5), (18, 5), (6, 9), (25, 100), (1, 50)]);
        let extent = Some(((1, 1), (40, 100)));
        let limit = dense::limit(cells.len(), 1 << 20);
        let mut store = Store::over(extent, limit, cells.iter().copied());
        let Store::Sparse(s) = &store else {
            panic!("a sparse sheet")
        };
        assert_eq!(
            s.runs.iter().map(|r| r.0).collect::<Vec<_>>(),
            vec![2, 3, 4, 7],
            "the runs"
        );
        for &(row, col) in &cells {
            store.put_number(row, col, f64::from(row * 1000 + col));
        }
        store.put_content(6, 2, None);
        store.put_content(20, 4, None);
        store.put_content(3, 5, None);
        let held = |store: &Store, row: u32, col: u32| {
            store.slot_of(row, col).is_some_and(|s| !s.is_empty())
        };
        for (top, left, bottom, right) in [
            (1, 1, 40, 100),
            (1, 1, 1, 1),
            (3, 1, 3, 5),
            (6, 2, 6, 9),
            (1, 2, 12, 2),
            (13, 1, 19, 10),
            (2, 3, 21, 7),
            (17, 1, 40, 4),
            (1, 8, 40, 99),
            (25, 100, 25, 100),
            (1, 1, 40, 6),
            (19, 4, 19, 4),
            (32, 3, 39, 7),
            (40, 1, 40, 100),
        ] {
            let area = A1Range {
                top,
                left,
                bottom,
                right,
            };
            let want: Vec<(u32, u32)> = (top..=bottom)
                .flat_map(|r| (left..=right).map(move |c| (r, c)))
                .filter(|&(r, c)| held(&store, r, c))
                .collect();
            let got: Vec<(u32, u32)> = store.walk(&area).collect();
            assert_eq!(got, want, "{area:?}");
            assert_eq!(store.count_in(&area), want.len() as u64, "{area:?}");
            let mut each = Vec::new();
            store.each_held(&area, &mut |r, c, _| each.push((r, c)));
            assert_eq!(each, want, "{area:?}");
        }
    }
}
