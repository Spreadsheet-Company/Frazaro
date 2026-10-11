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

use super::graph::{Area, Extent};
use super::value::Value;

/// A sheet's values over its extent, row by row. See the module's comment.
#[derive(Clone, Debug, PartialEq)]
pub struct Dense {
    top: u32,
    left: u32,
    width: usize,
    height: usize,
    values: Vec<Value>,
}

impl Dense {
    /// A store over an extent, in an old store's memory when one is given:
    /// the old store as it stands when its extent is the same, for the walk
    /// that keeps it to overwrite slot by slot, so that a text kept from one
    /// frame to the next is not copied again; every slot empty otherwise.
    /// `None` for a sheet with no cell, or for one whose extent holds more
    /// than `limit` cells.
    pub(crate) fn over(extent: Extent, limit: u64, old: Option<Dense>) -> Option<Dense> {
        let ((top, left), (bottom, right)) = extent?;
        let width = (right - left + 1) as usize;
        let height = (bottom - top + 1) as usize;
        if width as u64 * height as u64 > limit {
            return None;
        }
        if let Some(d) = old.as_ref() {
            if (d.top, d.left, d.width, d.height) == (top, left, width, height) {
                return old;
            }
        }
        let mut values = old.map(|d| d.values).unwrap_or_default();
        values.clear();
        values.resize(width * height, Value::Empty);
        Some(Dense {
            top,
            left,
            width,
            height,
            values,
        })
    }

    /// One cell's slot, to overwrite; `None` for a cell outside the extent,
    /// which a store made over its sheet's own extent never meets.
    pub(crate) fn slot(&mut self, row: u32, col: u32) -> Option<&mut Value> {
        let i = self.index(row, col)?;
        Some(&mut self.values[i])
    }

    /// A cell's slot: inside the extent, what a read of the cell answers,
    /// `Value::Empty` for a cell that holds nothing; `None` outside it.
    pub fn get(&self, row: u32, col: u32) -> Option<&Value> {
        self.index(row, col).map(|i| &self.values[i])
    }

    fn index(&self, row: u32, col: u32) -> Option<usize> {
        let r = row.checked_sub(self.top)? as usize;
        let c = col.checked_sub(self.left)? as usize;
        (r < self.height && c < self.width).then(|| r * self.width + c)
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
        let mut d = Dense::over(Some(((2, 2), (4, 4))), 9, None).unwrap();
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
            *d.slot(*row, *col).unwrap() = v.clone();
        }
        assert!(d.slot(1, 2).is_none(), "above the extent");
        assert!(d.slot(2, 5).is_none(), "right of it");
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
        let mut d = Dense::over(extent, 100, None).unwrap();
        *d.slot(10, 10).unwrap() = Value::Text("kept".to_string());
        let at = d.values.as_ptr();
        let Some(Value::Text(text)) = d.get(10, 10) else {
            panic!("a text")
        };
        let text_at = text.as_ptr();
        // The next frame's store over the same extent is the store as it
        // stands, for the walk to overwrite: a text in it is not copied.
        let same = Dense::over(extent, 100, Some(d)).unwrap();
        assert_eq!(same.values.as_ptr(), at);
        let Some(Value::Text(text)) = same.get(10, 10) else {
            panic!("the text kept")
        };
        assert_eq!(text.as_ptr(), text_at, "the same text, not a copy");
        // Over another extent, the same memory with every slot empty again.
        let again = Dense::over(Some(((3, 3), (5, 5))), 100, Some(same)).unwrap();
        assert_eq!(again.values.as_ptr(), at, "the old store's memory");
        assert_eq!(again.values.len(), 9);
        assert!(again.values.iter().all(|v| *v == Value::Empty));
        assert_eq!(again.get(3, 3), Some(&Value::Empty));
        assert!(Dense::over(Some(((1, 1), (10, 10))), 99, None).is_none());
        assert!(
            Dense::over(None, 100, None).is_none(),
            "a sheet with no cell"
        );
        // The largest extent a sheet can have is refused, never allocated.
        assert!(Dense::over(Some(((1, 1), (MAX_ROW, MAX_COLUMN))), 1 << 20, None).is_none());
    }
}
