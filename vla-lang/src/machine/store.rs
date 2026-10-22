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

use std::collections::BTreeMap;

use crate::calc::dense::{self, Dense, Slot};
use crate::calc::graph::Extent;
use crate::calc::{Computed, ErrorKind, Reason, Value};
use crate::sheet::Content;

/// The fewest cells a column of a sparse sheet keeps as a run of its own,
/// an array down it: fewer stand loose in the map, a run's own room being
/// worth it only past a few cells.
const RUN: u32 = 8;

/// What a cell kept loose holds, owned.
#[derive(Clone, Debug, PartialEq)]
pub(crate) enum Kept {
    Value(Value),
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
            Kept::Reason(r) => Slot::Reason(r),
            Kept::Pending => Slot::Pending,
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
        for ((row, col), slot) in &held {
            next.put_slot(*row, *col, slot.slot());
        }
        next
    }

    /// Every cell the store holds, by place, owned.
    fn held(&self) -> Vec<((u32, u32), Kept)> {
        let own = |slot: Slot<'_>| match slot {
            Slot::Reason(r) => Kept::Reason(r.clone()),
            Slot::Pending => Kept::Pending,
            other => Kept::Value(other.value()),
        };
        let mut out = Vec::new();
        let mut array = |d: &Dense| {
            for i in 0..d.slots() {
                let (slot, at) = d.at(i);
                if !slot.is_empty() {
                    out.push((at, own(slot)));
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

    /// A value cell's slot written with what a read of its content answers:
    /// a value as it stands, an error text as its kind, nothing for no cell
    /// or a content that holds no value.
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
                    let value = match content {
                        Some(Content::Number(n)) => Value::Number(*n),
                        Some(Content::Text(t)) => Value::Text(t.clone()),
                        Some(Content::Bool(b)) => Value::Bool(*b),
                        Some(Content::Error(e)) => {
                            Value::Error(ErrorKind::parse(e).unwrap_or(ErrorKind::Value))
                        }
                        _ => Value::Empty,
                    };
                    match value {
                        Value::Empty => s.loose.remove(&(row, col)),
                        v => s.loose.insert((row, col), Kept::Value(v)),
                    };
                }
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
}
