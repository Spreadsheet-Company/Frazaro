//! A sheet's values kept as an array over its extent (KERNEL.20, slice 1,
//! 2026-10-10): a read of a cell is an index into it and a range a walk of
//! it, where a sheet's cells are a B-tree looked up one cell at a time.
//!
//! An engine's machine keeps one for each `.last` twin, filled at every
//! frame's end in the walk that fills the twin's cells (`crate::machine`),
//! since most of a frame's reads are of a twin: Life's rule reads 21 cells of
//! `Screen.last` for each cell it computes, 18 of them in two ranges
//! (Alonzo's `CART.1`, its profile). The twin's cells stay a sheet that every
//! other reader reads, the record, the plane and a save; the array is the
//! evaluator's, through [`super::Env::dense`]. A slot holds what a read of
//! its cell answers, a [`Value`], or [`Value::Empty`] for a cell that holds
//! nothing, which a walk skips as a walk of a sheet's cells skips a cell that
//! is not there; so a walk of the array is the walk of the cells, value for
//! value and in the same order, and an aggregate folding over either gives
//! the same bits.
//!
//! The array is never larger than its owner allows (`limit`), so that a sheet
//! of two cells far apart is not a million empty slots; such a sheet is read
//! through its cells.
//!
//! Beside each slot the store keeps its number and its kind, and for each row
//! of the extent how many of its slots hold no number (KERNEL.43), each
//! written wherever a slot is written: so that an evaluation of many cells
//! side by side in one row reads a reference's cells as one slice of numbers,
//! and the slots themselves where any of them holds something else. A store
//! no formula may read so keeps none (its owner says which), so that the end
//! of a spreadsheet's frame, whose twins no formula reads, writes nothing
//! more than its slots; one that comes to be read has them made from its
//! slots.
//!
//! Since KERNEL.31 a machine keeps one for each of a frame's own sheets that
//! its plan reads by a range or by a run's moving cell, each slot written
//! where the frame keeps a value that changed, its numbers beside, so that a
//! range of the frame's own cells is folded and walked by index as a twin's
//! is, and a run's cells read one by slices: side by side in a row, a slice
//! of one row (KERNEL.43), and one below the other in a column, the slots of
//! one column read at the store's width.

use super::graph::{Area, Extent};
use super::value::Value;

/// The most slots a store holds for each cell of its sheet (KERNEL.20): a
/// slot is 24 bytes, its number and kind beside it 9 more (KERNEL.43), and a
/// cell of a sheet's B-tree about 137 (measured 2026-10-10), so a store
/// within this is smaller than the cells it shadows, and a sheet sparser than
/// that, two cells far apart, is read through its cells.
pub(crate) const SLOTS_A_CELL: u64 = 4;

/// The most slots a store of a sheet of `cells` cells holds: [`SLOTS_A_CELL`]
/// a cell, and never more than `most`, a grid's own limit. A twin's rule, and
/// a frame's own sheet's (KERNEL.31), so that a machine and a door's census
/// of the same grid keep the same sheets as arrays.
pub(crate) fn limit(cells: usize, most: u64) -> u64 {
    (cells as u64).saturating_mul(SLOTS_A_CELL).min(most)
}

/// Whether a store over an extent is within a limit, as [`Dense::over`]
/// asks: a sheet that holds a cell, over no more cells than `limit`.
pub(crate) fn fits(extent: Extent, limit: u64) -> bool {
    extent.is_some_and(|((top, left), (bottom, right))| {
        u64::from(right - left + 1) * u64::from(bottom - top + 1) <= limit
    })
}

/// A slot's kind beside its number (KERNEL.43): a number, or anything else (a
/// text, a truth value, an error, nothing), which a read by slice leaves to
/// the slot itself.
const NUMBER: u8 = 0;
const OTHER: u8 = 1;

/// A slot's number, 0 where it holds none, and its kind (KERNEL.43).
#[inline]
fn number_of(v: &Value) -> (f64, u8) {
    match v {
        Value::Number(n) => (*n, NUMBER),
        _ => (0.0, OTHER),
    }
}

/// A sheet's values over its extent, row by row. See the module's comment.
#[derive(Clone, Debug, PartialEq)]
pub struct Dense {
    top: u32,
    left: u32,
    width: usize,
    height: usize,
    values: Vec<Value>,
    /// The numbers beside the slots, for a store a formula may read by
    /// slices (KERNEL.43); `None` for one none may.
    shadow: Option<Shadow>,
}

/// Each slot's number and its kind, and by row of the extent how many of
/// its slots hold no number (KERNEL.43).
#[derive(Clone, Debug, Default, PartialEq)]
struct Shadow {
    numbers: Vec<f64>,
    kinds: Vec<u8>,
    odd: Vec<u32>,
}

impl Dense {
    /// A store over an extent, in an old store's memory when one is given:
    /// the old store as it stands when its extent is the same, for the walk
    /// that keeps it to overwrite slot by slot, so that a text kept from one
    /// frame to the next is not copied again; every slot empty otherwise.
    /// Its numbers kept beside its slots where `numbers` asks (KERNEL.43).
    /// `None` for a sheet with no cell, or for one whose extent holds more
    /// than `limit` cells.
    pub(crate) fn over(
        extent: Extent,
        limit: u64,
        old: Option<Dense>,
        numbers: bool,
    ) -> Option<Dense> {
        let ((top, left), (bottom, right)) = extent?;
        if !fits(extent, limit) {
            return None;
        }
        let width = (right - left + 1) as usize;
        let height = (bottom - top + 1) as usize;
        let (mut values, spare) = match old {
            Some(mut d) if d.stands_over(extent) => {
                d.shade(numbers);
                return Some(d);
            }
            Some(d) => (d.values, d.shadow),
            None => (Vec::new(), None),
        };
        values.clear();
        values.resize(width * height, Value::Empty);
        let mut d = Dense {
            top,
            left,
            width,
            height,
            values,
            shadow: None,
        };
        if numbers {
            d.shadow = Some(d.made(spare.unwrap_or_default()));
        }
        Some(d)
    }

    /// The numbers kept beside the slots, or not (KERNEL.43): made from the
    /// slots when the store comes to keep them, dropped when it stops.
    pub(crate) fn shade(&mut self, numbers: bool) {
        match (numbers, self.shadow.is_some()) {
            (true, false) => self.shadow = Some(self.made(Shadow::default())),
            (false, true) => self.shadow = None,
            _ => {}
        }
    }

    /// Whether the store keeps its numbers beside its slots (KERNEL.43).
    #[cfg(test)]
    pub(crate) fn keeps_numbers(&self) -> bool {
        self.shadow.is_some()
    }

    /// The numbers beside the slots made from the slots, in a shadow's
    /// memory.
    fn made(&self, mut s: Shadow) -> Shadow {
        let n = self.values.len();
        s.numbers.clear();
        s.numbers.reserve(n);
        s.kinds.clear();
        s.kinds.reserve(n);
        for v in &self.values {
            let (x, kind) = number_of(v);
            s.numbers.push(x);
            s.kinds.push(kind);
        }
        s.odd.clear();
        s.odd.extend(
            s.kinds
                .chunks(self.width)
                .map(|row| row.iter().filter(|&&k| k != NUMBER).count() as u32),
        );
        s
    }

    /// Whether the store stands over an extent, as [`Dense::over`] keeps it
    /// whole when it does (KERNEL.36: the walk reads a slot before it
    /// overwrites it only in a store kept so).
    pub(crate) fn stands_over(&self, extent: Extent) -> bool {
        extent.is_some_and(|((top, left), (bottom, right))| {
            let width = (right - left + 1) as usize;
            let height = (bottom - top + 1) as usize;
            (self.top, self.left, self.width, self.height) == (top, left, width, height)
        })
    }

    /// One cell's slot overwritten with a value, its number and kind beside
    /// it; `false` for a cell outside the extent, which a store made over its
    /// sheet's own extent never meets. The tests'; the walk finds a slot's
    /// index first, to read the bit of a twin's cell behind it (KERNEL.36).
    #[cfg(test)]
    pub(crate) fn set(&mut self, row: u32, col: u32, v: Value) -> bool {
        let Some(i) = self.index(row, col) else {
            return false;
        };
        self.write_at(i, |slot| *slot = v);
        true
    }

    /// A cell's slot: inside the extent, what a read of the cell answers,
    /// `Value::Empty` for a cell that holds nothing; `None` outside it.
    pub fn get(&self, row: u32, col: u32) -> Option<&Value> {
        self.index(row, col).map(|i| &self.values[i])
    }

    /// A cell's slot's index, row by row; `None` outside the extent.
    pub(crate) fn index(&self, row: u32, col: u32) -> Option<usize> {
        let r = row.checked_sub(self.top)? as usize;
        let c = col.checked_sub(self.left)? as usize;
        (r < self.height && c < self.width).then(|| r * self.width + c)
    }

    /// How many slots the store holds, one a cell of its extent.
    pub(crate) fn slots(&self) -> usize {
        self.values.len()
    }

    /// The slot at an index as it stands, which a frame's end reads before
    /// it writes it (KERNEL.27).
    pub(crate) fn slot_at(&self, i: usize) -> &Value {
        &self.values[i]
    }

    /// The slot at an index overwritten by `f` (KERNEL.27), and where the
    /// store keeps its numbers, its number and its kind written beside it,
    /// and its row's count of slots holding no number with them (KERNEL.43):
    /// the one way a slot is written.
    #[inline]
    pub(crate) fn write_at(&mut self, i: usize, f: impl FnOnce(&mut Value)) {
        f(&mut self.values[i]);
        let Some(s) = self.shadow.as_mut() else {
            return;
        };
        let (n, kind) = number_of(&self.values[i]);
        if s.kinds[i] != kind {
            let row = i / self.width;
            if kind == NUMBER {
                s.odd[row] -= 1;
            } else {
                s.odd[row] += 1;
            }
            s.kinds[i] = kind;
        }
        s.numbers[i] = n;
    }

    /// Of `n` cells side by side whose reference covers rows `top` to
    /// `bottom` and `width` columns from `col` plus the cell's place, the
    /// places `a..b` of those it keeps inside the extent (KERNEL.43), every
    /// place where it `moves`, or none or all where it stands still: so that
    /// a run's cells whose reference leaves the array at the row's ends are
    /// read alone and the rest by slices.
    pub(crate) fn inside(
        &self,
        (top, bottom): (u32, u32),
        col: u32,
        width: usize,
        n: usize,
        moves: bool,
    ) -> (usize, usize) {
        let (first, last) = (self.top, self.top + (self.height as u32 - 1));
        if top < first || bottom > last {
            return (0, 0);
        }
        let (left, right) = (
            i64::from(self.left),
            i64::from(self.left) + self.width as i64 - 1,
        );
        let (col, width) = (i64::from(col), width as i64);
        if !moves {
            let fits = col >= left && col + width - 1 <= right;
            return if fits { (0, n) } else { (0, 0) };
        }
        let a = (left - col).clamp(0, n as i64);
        let b = (right - (col + width - 1) + 1).clamp(0, n as i64);
        if a >= b {
            (0, 0)
        } else {
            (a as usize, b as usize)
        }
    }

    /// The numbers of `n` slots side by side in one row, from a cell's
    /// (KERNEL.43): what an evaluation of many cells reads a reference's
    /// cells of one row as, when every one is inside the extent and holds a
    /// number; `None` otherwise, or where the store keeps no numbers, its
    /// slots read one at a time.
    #[inline]
    pub(crate) fn numbers_in_row(&self, row: u32, col: u32, n: usize) -> Option<&[f64]> {
        let s = self.shadow.as_ref()?;
        let r = row.checked_sub(self.top)? as usize;
        let c = col.checked_sub(self.left)? as usize;
        if r >= self.height || c + n > self.width {
            return None;
        }
        let i = r * self.width + c;
        if s.odd[r] != 0 && s.kinds[i..i + n].iter().any(|&k| k != NUMBER) {
            return None;
        }
        Some(&s.numbers[i..i + n])
    }

    /// The numbers of `n` slots one below the other in one column, from a
    /// cell's, pushed to `out` (KERNEL.31): what an evaluation of many cells
    /// reads a reference's cells as where its cells stand one below the
    /// other, when every one is inside the extent and holds a number;
    /// `false`, nothing pushed, otherwise, or where the store keeps no
    /// numbers, its slots read one at a time.
    #[inline]
    pub(crate) fn numbers_in_column(
        &self,
        row: u32,
        col: u32,
        n: usize,
        out: &mut Vec<f64>,
    ) -> bool {
        let Some(s) = self.shadow.as_ref() else {
            return false;
        };
        let (Some(r), Some(c)) = (row.checked_sub(self.top), col.checked_sub(self.left)) else {
            return false;
        };
        let (r, c, w) = (r as usize, c as usize, self.width);
        if r + n > self.height || c >= w {
            return false;
        }
        let first = r * w + c;
        if s.odd[r..r + n].iter().any(|&o| o != 0)
            && (0..n).any(|j| s.kinds[first + j * w] != NUMBER)
        {
            return false;
        }
        out.extend((0..n).map(|j| s.numbers[first + j * w]));
        true
    }

    /// The value at an index, and its cell (KERNEL.27).
    pub(crate) fn at(&self, i: usize) -> (&Value, (u32, u32)) {
        let (r, c) = (i / self.width, i % self.width);
        (&self.values[i], (self.top + r as u32, self.left + c as u32))
    }

    /// Every value an area holds, row by row, to `f`, the area clamped to the
    /// extent first, as a walk of the cells clamps it; an empty slot is
    /// skipped. `f` is a type of its own and not a pointer, so that an
    /// aggregate's fold is compiled into the walk (KERNEL.25).
    pub fn each<E, F>(&self, area: &Area, mut f: F) -> Result<(), E>
    where
        F: FnMut(Value) -> Result<(), E>,
    {
        self.each_ref(area, |v| f(v.clone()))
    }

    /// [`Dense::each`]'s walk, each value lent where it stands: what an
    /// evaluation of many cells at once folds an aggregate over, copying
    /// nothing (KERNEL.26).
    pub(crate) fn each_ref<E, F>(&self, area: &Area, mut f: F) -> Result<(), E>
    where
        F: FnMut(&Value) -> Result<(), E>,
    {
        let bottom = self.top + (self.height as u32 - 1);
        let right = self.left + (self.width as u32 - 1);
        let (top, left) = (area.top.max(self.top), area.left.max(self.left));
        let (last_row, last_col) = (area.bottom.min(bottom), area.right.min(right));
        if top > last_row || left > last_col {
            return Ok(());
        }
        let n = (last_col - left + 1) as usize;
        for row in top..=last_row {
            let from = (row - self.top) as usize * self.width + (left - self.left) as usize;
            for v in &self.values[from..from + n] {
                if !matches!(v, Value::Empty) {
                    f(v)?;
                }
            }
        }
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::calc::ErrorKind;
    use crate::sheet::{MAX_COLUMN, MAX_ROW};

    fn rect(top: u32, left: u32, bottom: u32, right: u32) -> Area {
        Area {
            sheet: 0,
            top,
            left,
            bottom,
            right,
        }
    }

    fn walk(d: &Dense, area: Area) -> Vec<Value> {
        let mut out = Vec::new();
        d.each::<(), _>(&area, |v| {
            out.push(v);
            Ok(())
        })
        .unwrap();
        out
    }

    #[test]
    fn a_store_is_its_extent_row_by_row_and_skips_what_is_empty() {
        // B2:D4, every cell but C3 holding a value, one of each kind.
        let mut d = Dense::over(Some(((2, 2), (4, 4))), 9, None, true).unwrap();
        let n = Value::Number;
        let kinds = [
            n(1.0),
            Value::Text("x".to_string()),
            Value::Bool(true),
            n(4.0),
            Value::Error(ErrorKind::NA),
            n(-0.5),
            n(7.0),
            n(8.0),
        ];
        let cells = [
            (2, 2),
            (2, 3),
            (2, 4),
            (3, 2),
            (3, 4),
            (4, 2),
            (4, 3),
            (4, 4),
        ];
        for ((row, col), v) in cells.iter().zip(&kinds) {
            assert!(d.set(*row, *col, v.clone()));
        }
        assert!(!d.set(1, 2, Value::Empty), "above the extent");
        assert!(!d.set(2, 5, Value::Empty), "right of it");
        assert_eq!(d.get(2, 3), Some(&kinds[1]));
        assert_eq!(d.get(3, 3), Some(&Value::Empty), "a cell holding nothing");
        assert_eq!(d.get(5, 2), None, "outside");
        assert_eq!(d.get(2, 1), None);
        // The extent, then areas past it, across it and outside it, each
        // clamped as a walk of the cells clamps it.
        assert_eq!(walk(&d, rect(2, 2, 4, 4)), kinds.to_vec());
        assert_eq!(walk(&d, rect(1, 1, MAX_ROW, MAX_COLUMN)), kinds.to_vec());
        assert_eq!(
            walk(&d, rect(3, 3, 9, 9)),
            vec![kinds[4].clone(), n(7.0), n(8.0)]
        );
        assert_eq!(walk(&d, rect(1, 3, 3, 3)), vec![kinds[1].clone()]);
        assert_eq!(walk(&d, rect(3, 3, 3, 3)), Vec::<Value>::new());
        assert_eq!(walk(&d, rect(5, 1, 9, 9)), Vec::<Value>::new());
        assert_eq!(walk(&d, rect(1, 5, 1, 9)), Vec::<Value>::new());
        // A fold stops at its first error, as a walk of the cells does.
        let mut seen = 0;
        let stopped = d.each(&rect(2, 2, 4, 4), |v| {
            seen += 1;
            if matches!(v, Value::Bool(_)) {
                Err("stop")
            } else {
                Ok(())
            }
        });
        assert_eq!((stopped, seen), (Err("stop"), 3));
    }

    #[test]
    fn a_store_keeps_its_memory_and_refuses_what_passes_its_limit() {
        let extent = Some(((1, 1), (10, 10)));
        let mut d = Dense::over(extent, 100, None, true).unwrap();
        assert!(d.set(10, 10, Value::Text("kept".to_string())));
        let at = d.values.as_ptr();
        let Some(Value::Text(text)) = d.get(10, 10) else {
            panic!("a text")
        };
        let text_at = text.as_ptr();
        // The next frame's store over the same extent is the store as it
        // stands, for the walk to overwrite: a text in it is not copied.
        let same = Dense::over(extent, 100, Some(d), true).unwrap();
        assert_eq!(same.values.as_ptr(), at);
        let Some(Value::Text(text)) = same.get(10, 10) else {
            panic!("the text kept")
        };
        assert_eq!(text.as_ptr(), text_at, "the same text, not a copy");
        // Over another extent, the same memory with every slot empty again.
        let again = Dense::over(Some(((3, 3), (5, 5))), 100, Some(same), true).unwrap();
        assert_eq!(again.values.as_ptr(), at, "the old store's memory");
        assert_eq!(again.values.len(), 9);
        assert!(again.values.iter().all(|v| *v == Value::Empty));
        assert_eq!(again.get(3, 3), Some(&Value::Empty));
        assert!(Dense::over(Some(((1, 1), (10, 10))), 99, None, true).is_none());
        assert!(
            Dense::over(None, 100, None, true).is_none(),
            "a sheet with no cell"
        );
        // The largest extent a sheet can have is refused, never allocated.
        assert!(Dense::over(Some(((1, 1), (MAX_ROW, MAX_COLUMN))), 1 << 20, None, true).is_none());
    }

    /// What a read of slices must answer, from the slots alone: the numbers
    /// of `n` slots of a row from a cell, when each is inside the extent and
    /// holds a number.
    fn slice_of(d: &Dense, row: u32, col: u32, n: usize) -> Option<Vec<f64>> {
        (0..n)
            .map(|j| match d.get(row, col + j as u32)? {
                Value::Number(x) => Some(*x),
                _ => None,
            })
            .collect()
    }

    #[test]
    fn the_cells_a_reference_keeps_inside_are_those_it_leaves_nowhere() {
        // KERNEL.43: over a store of B2:F6, for references of every width
        // moving with cells side by side, and standing still, from columns
        // before, inside and past the extent and rows inside and out, the
        // places `inside` answers are the cells whose reference stays wholly
        // inside the extent, found one by one.
        let d = Dense::over(Some(((2, 2), (6, 6))), 100, None, true).unwrap();
        for (top, bottom) in [(2, 2), (3, 5), (1, 3), (6, 7), (2, 6)] {
            for col in 0..=8u32 {
                for width in 1..=4usize {
                    for n in 1..=7usize {
                        for moves in [true, false] {
                            let fits = |i: usize| {
                                let left = col + if moves { i as u32 } else { 0 };
                                let right = left + width as u32 - 1;
                                top >= 2 && bottom <= 6 && left >= 2 && right <= 6
                            };
                            let want: Vec<usize> = (0..n).filter(|&i| fits(i)).collect();
                            let (a, b) = d.inside((top, bottom), col, width, n, moves);
                            assert_eq!(
                                (a..b).collect::<Vec<_>>(),
                                want,
                                "rows {top}-{bottom}, column {col}, width {width}, {n} cells, moving {moves}"
                            );
                        }
                    }
                }
            }
        }
    }

    /// What a read of a column must answer, from the slots alone: the
    /// numbers of `n` slots one below the other from a cell, when each is
    /// inside the extent and holds a number (KERNEL.31).
    fn column_of(d: &Dense, row: u32, col: u32, n: usize) -> Option<Vec<f64>> {
        (0..n)
            .map(|j| match d.get(row + j as u32, col)? {
                Value::Number(x) => Some(*x),
                _ => None,
            })
            .collect()
    }

    /// A store's numbers spelled by their bits, so that `-0.0` is not `0.0`.
    fn spelled(s: &Shadow) -> (Vec<u64>, &[u8], &[u32]) {
        (
            s.numbers.iter().map(|x| x.to_bits()).collect(),
            &s.kinds,
            &s.odd,
        )
    }

    #[test]
    fn a_store_s_numbers_are_its_slots_after_every_write() {
        // KERNEL.43: slots written at random with values of every kind, the
        // store kept and made again over the same extent and another, its
        // numbers kept, dropped and kept again; after each write, a store
        // keeping its numbers holds the ones its slots make, to the bit, and
        // every slice a read may ask for, inside the extent, across its edges
        // and outside it, is what the slots say, and a store keeping none
        // answers none.
        let mut seed: u64 = 0x5EED;
        let mut next = |n: u64| {
            seed = seed
                .wrapping_mul(6_364_136_223_846_793_005)
                .wrapping_add(1_442_695_040_888_963_407);
            (seed >> 33) % n
        };
        let values = [
            Value::Number(0.0),
            Value::Number(-0.0),
            Value::Number(1.5),
            Value::Number(-3.0),
            Value::Number(1e300),
            Value::Text("x".to_string()),
            Value::Bool(true),
            Value::Error(ErrorKind::NA),
            Value::Empty,
        ];
        let extents = [Some(((2, 3), (6, 9))), Some(((1, 1), (3, 4)))];
        let mut d = Dense::over(extents[0], 100, None, true).unwrap();
        let (mut kept, mut columns) = (0, 0);
        for round in 0..2000 {
            // Every hundredth round the numbers dropped or kept again, every
            // five hundredth the store made again, keeping them but once.
            if round % 500 == 499 {
                d = Dense::over(extents[next(2) as usize], 100, Some(d), round != 999).unwrap();
            } else if round % 100 == 99 {
                d.shade(round / 100 % 2 == 1);
            }
            let (top, left) = (d.top, d.left);
            let (row, col) = (
                top + next(d.height as u64) as u32,
                left + next(d.width as u64) as u32,
            );
            let v = values[next(values.len() as u64) as usize].clone();
            assert!(d.set(row, col, v));
            if let Some(s) = &d.shadow {
                kept += 1;
                for r in 0..d.height {
                    let row = top + r as u32;
                    let odd = (0..d.width as u32)
                        .filter(|&c| !matches!(d.get(row, left + c), Some(Value::Number(_))))
                        .count();
                    assert_eq!(s.odd[r] as usize, odd, "round {round}, row {row}");
                }
                assert_eq!(
                    spelled(s),
                    spelled(&d.made(Shadow::default())),
                    "round {round}"
                );
            }
            for _ in 0..4 {
                let row = (top + next(d.height as u64 + 2) as u32).saturating_sub(1);
                let col = (left + next(d.width as u64 + 2) as u32).saturating_sub(1);
                let n = 1 + next(d.width as u64 + 1) as usize;
                let got = d
                    .numbers_in_row(row, col, n)
                    .map(|s| s.iter().map(|x| x.to_bits()).collect::<Vec<_>>());
                let want = slice_of(&d, row, col, n)
                    .filter(|_| d.shadow.is_some())
                    .map(|s| s.iter().map(|x| x.to_bits()).collect::<Vec<_>>());
                assert_eq!(got, want, "round {round}: {n} slots from ({row}, {col})");
                // And the slots one below the other from it (KERNEL.31),
                // pushed after what the lane holds already, or nothing.
                let n = 1 + next(d.height as u64 + 1) as usize;
                let mut out = vec![7.0];
                let read = d.numbers_in_column(row, col, n, &mut out);
                let want = column_of(&d, row, col, n).filter(|_| d.shadow.is_some());
                let got = read.then(|| out[1..].iter().map(|x| x.to_bits()).collect::<Vec<_>>());
                assert_eq!(
                    got,
                    want.map(|s| s.iter().map(|x| x.to_bits()).collect::<Vec<_>>()),
                    "round {round}: {n} slots down from ({row}, {col})"
                );
                assert_eq!(out.len(), if read { 1 + n } else { 1 }, "round {round}");
                columns += usize::from(read);
            }
        }
        assert!(
            0 < kept && kept < 2000,
            "rounds with numbers and without: {kept}"
        );
        assert!(columns > 0, "a column read whole somewhere");
    }
}
