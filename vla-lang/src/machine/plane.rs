//! The plane (KERNEL.22, 2026-10-09; `Alonzo/SPEC.md` section 3.1): a
//! window of a sheet as one byte a cell, in row-major order, which a host
//! blits through a palette. A whole number from 0 to 254 is itself; an
//! empty cell, and a formula before its first value, is 0; anything else,
//! a text, a truth value, an error, a fraction, a number past 254, a
//! formula not computed, is 255, which the host paints in the error colour,
//! so a wrong value is seen and never silently black.
//!
//! The plane is a projection by name beside the grid's record, not an
//! implementation of [`crate::projection::Projection`], whose output is
//! lines: a byte a cell is a raster, and whether rasters get a seam of
//! their own waits for a second one (`KERNEL.12`). The record of the same
//! window shows the same cells as values, which is the engine's equality
//! test (`ENGINE.2`).
//!
//! **The kept plane** (KERNEL.24, 2026-10-09; Alonzo's `ENGINE.2`). The
//! walk, [`bytes`], looks up every formula cell's value of the window, the
//! whole cost of a plane view. The end of every frame already looks up the
//! same values to fill the twins, so a sheet an engine views as a plane has
//! its bytes kept there, over its extent, a [`Raster`]; a view of it is a
//! copy. The walk stays the reference: it answers whenever no raster stands
//! (before the first frame's end after a sheet is first viewed, a twin, a
//! sheet past the limit), and it keeps its own spelling of the rule for a
//! cell, so that a test holding the two equal compares two readings and never
//! one twice. A write keeps the raster (KERNEL.56), each cell it changed
//! given its byte, the raster taken over an extent the write grew; it dropped
//! it, the next frame's end making it again from every cell of its sheet.

use crate::calc::dense::Slot;
use crate::calc::graph::Extent;
use crate::calc::{Computed, Value, Values};
use crate::projection::Window;
use crate::sheet::{A1Range, Content, Workbook};

/// A window's bytes: its width times its height, row by row.
pub fn bytes(model: &Workbook, values: &dyn Values, window: &Window) -> Vec<u8> {
    let r = window.range;
    let width = (r.right - r.left + 1) as usize;
    let height = (r.bottom - r.top + 1) as usize;
    let mut out = vec![0u8; width * height];
    let Some(si) = model.find_sheet(&window.sheet) else {
        return out;
    };
    let mut put = |row: u32, col: u32, content: &Content| {
        let at = (row - r.top) as usize * width + (col - r.left) as usize;
        out[at] = match content {
            Content::Number(n) => byte(*n),
            Content::Text(_) | Content::Bool(_) | Content::Error(_) => 255,
            _ => match values.computed((si, row, col)) {
                Some(Computed::Value(Value::Number(n))) => byte(n),
                Some(_) => 255,
                None => 0,
            },
        };
    };
    // The cells the values hold for a sheet that keeps none in the model,
    // an engine's twin (KERNEL.49), or the model's own.
    match values.cells(si, &r) {
        Some(held) => {
            for ((row, col), cell) in &held {
                put(*row, *col, &cell.content);
            }
        }
        None => {
            for (&(row, col), cell) in model.sheets[si]
                .cells
                .range((r.top, 0)..=(r.bottom, u32::MAX))
            {
                if col >= r.left && col <= r.right {
                    put(row, col, &cell.content);
                }
            }
        }
    }
    out
}

/// A number's byte: itself when it is a whole number from 0 to 254, else 255.
pub fn byte(n: f64) -> u8 {
    if (0.0..=254.0).contains(&n) && n.fract() == 0.0 {
        n as u8
    } else {
        255
    }
}

/// A sheet's plane kept over its extent (KERNEL.24): a byte a cell, row by
/// row, a cell of the extent that holds nothing being 0, as the walk reads
/// an empty cell.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Raster {
    top: u32,
    left: u32,
    width: usize,
    height: usize,
    bytes: Vec<u8>,
}

/// Whether a raster can be kept over an extent: one with a cell, holding
/// no more than `limit` cells.
pub(crate) fn fits(extent: Extent, limit: u64) -> bool {
    extent.is_some_and(|((top, left), (bottom, right))| {
        u64::from(right - left + 1) * u64::from(bottom - top + 1) <= limit
    })
}

impl Raster {
    /// A raster over an extent, every byte 0, in an old raster's memory when
    /// one is given; `None` for a sheet with no cell, or for one whose
    /// extent holds more than `limit` cells.
    pub(crate) fn over(extent: Extent, limit: u64, old: Option<Raster>) -> Option<Raster> {
        if !fits(extent, limit) {
            return None;
        }
        let ((top, left), (bottom, right)) = extent?;
        let width = (right - left + 1) as usize;
        let height = (bottom - top + 1) as usize;
        let mut bytes = old.map(|r| r.bytes).unwrap_or_default();
        bytes.clear();
        bytes.resize(width * height, 0);
        Some(Raster {
            top,
            left,
            width,
            height,
            bytes,
        })
    }

    /// One cell's byte; `false` for a cell outside the extent, which a
    /// raster made over its sheet's own extent never meets.
    pub(crate) fn set(&mut self, row: u32, col: u32, byte: u8) -> bool {
        let (Some(r), Some(c)) = (row.checked_sub(self.top), col.checked_sub(self.left)) else {
            return false;
        };
        let (r, c) = (r as usize, c as usize);
        if r >= self.height || c >= self.width {
            return false;
        }
        self.bytes[r * self.width + c] = byte;
        true
    }

    /// The raster over an extent that holds its own, within `limit` cells
    /// (KERNEL.56): as it stands where the extent is its own, else each byte
    /// moved to its cell's place and every new place 0, a raster whose bottom
    /// alone moved extended where it stands; `false`, the raster untouched,
    /// for an extent past the limit or one that does not hold its own.
    pub(crate) fn regrow(&mut self, extent: Extent, limit: u64) -> bool {
        if !fits(extent, limit) {
            return false;
        }
        let Some(((top, left), (bottom, right))) = extent else {
            return false;
        };
        let (old_bottom, old_right) = (
            self.top + (self.height as u32 - 1),
            self.left + (self.width as u32 - 1),
        );
        if top > self.top || left > self.left || bottom < old_bottom || right < old_right {
            return false;
        }
        let width = (right - left + 1) as usize;
        let height = (bottom - top + 1) as usize;
        if (top, left, width) == (self.top, self.left, self.width) {
            self.bytes.resize(width * height, 0);
            self.height = height;
            return true;
        }
        let (dr, dc) = ((self.top - top) as usize, (self.left - left) as usize);
        #[cfg(test)]
        crate::sheet::work::relaid(self.bytes.len());
        let mut bytes = vec![0u8; width * height];
        for r in 0..self.height {
            let (from, to) = (r * self.width, (r + dr) * width + dc);
            bytes[to..to + self.width].copy_from_slice(&self.bytes[from..from + self.width]);
        }
        (self.top, self.left, self.width, self.height) = (top, left, width, height);
        self.bytes = bytes;
        true
    }

    /// A window's bytes, its width times its height, row by row: the kept
    /// bytes where the window meets the extent, 0 everywhere else.
    pub fn window(&self, range: &A1Range) -> Vec<u8> {
        let width = (range.right - range.left + 1) as usize;
        let height = (range.bottom - range.top + 1) as usize;
        let mut out = vec![0u8; width * height];
        let bottom = self.top + (self.height as u32 - 1);
        let right = self.left + (self.width as u32 - 1);
        let (top, left) = (range.top.max(self.top), range.left.max(self.left));
        let (last_row, last_col) = (range.bottom.min(bottom), range.right.min(right));
        if top > last_row || left > last_col {
            return out;
        }
        let n = (last_col - left + 1) as usize;
        for row in top..=last_row {
            let from = (row - self.top) as usize * self.width + (left - self.left) as usize;
            let to = (row - range.top) as usize * width + (left - range.left) as usize;
            out[to..to + n].copy_from_slice(&self.bytes[from..from + n]);
        }
        out
    }
}

/// A cell's byte as the kept plane reads it, from its content and, for a
/// formula, its value of the last complete frame, `None` before the first:
/// the rule [`bytes`] spells for itself, written a second time and naming
/// every kind of content, so that a kind added later is decided here too.
pub(crate) fn kept_byte(content: &Content, value: Option<&Computed>) -> u8 {
    match content {
        Content::Number(n) => byte(*n),
        Content::Text(_) | Content::Bool(_) | Content::Error(_) => 255,
        Content::Formula(_)
        | Content::SharedMaster { .. }
        | Content::SharedChild { .. }
        | Content::DynamicFormula(_)
        | Content::ArrayFormula(_) => formula_byte(value),
    }
}

/// A formula cell's byte from its slot (KERNEL.49): [`formula_byte`]'s rule
/// read off the slot, a formula with no value 0, a number itself, and any
/// other kind, nothing among them, 255.
pub(crate) fn slot_byte(slot: Slot<'_>) -> u8 {
    match slot {
        Slot::Number(n) => byte(n),
        Slot::Pending => 0,
        _ => 255,
    }
}

/// A formula cell's byte, from its value alone.
pub(crate) fn formula_byte(value: Option<&Computed>) -> u8 {
    match value {
        Some(Computed::Value(Value::Number(n))) => byte(*n),
        Some(_) => 255,
        None => 0,
    }
}
