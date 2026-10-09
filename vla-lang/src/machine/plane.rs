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

use crate::calc::{Computed, Value, Values};
use crate::projection::Window;
use crate::sheet::{Content, Workbook};

/// A window's bytes: its width times its height, row by row.
pub fn bytes(model: &Workbook, values: &dyn Values, window: &Window) -> Vec<u8> {
    let r = window.range;
    let width = (r.right - r.left + 1) as usize;
    let height = (r.bottom - r.top + 1) as usize;
    let mut out = vec![0u8; width * height];
    let Some(si) = model.find_sheet(&window.sheet) else {
        return out;
    };
    for (&(row, col), cell) in model.sheets[si]
        .cells
        .range((r.top, 0)..=(r.bottom, u32::MAX))
    {
        if col < r.left || col > r.right {
            continue;
        }
        let at = (row - r.top) as usize * width + (col - r.left) as usize;
        out[at] = match &cell.content {
            Content::Number(n) => byte(*n),
            Content::Text(_) | Content::Bool(_) | Content::Error(_) => 255,
            _ => match values.computed((si, row, col)) {
                Some(Computed::Value(Value::Number(n))) => byte(*n),
                Some(_) => 255,
                None => 0,
            },
        };
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
