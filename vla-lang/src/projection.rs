//! The projections seam: a window onto the grid, and a projection as a pure
//! function from the model and a window to lines (`KERNEL.1`, the seam as
//! data; `KERNEL.4`, the first implementation, [`crate::view::Grid`]). Cut
//! out of `frazaro-core`'s `kernel.rs` with the language (`PORT.12`), since
//! an engine's screen is a projection of the grid and must not depend on
//! the bridges; the seams' registry, the engines trait and the host
//! profiles stay in the core.

use crate::sheet::{A1Range, Workbook};

/// A window onto the model: one sheet by name, one rectangle of it.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Window {
    pub sheet: String,
    pub range: A1Range,
}

/// A projection: a pure function from the model and a window to lines, the
/// shape every door prints (`SD-23`). The grid is the first
/// (`crate::view::Grid`, `KERNEL.4`), the sentence pane and the dependency
/// cone drawn as a diagram the core's (`KERNEL.12`). An engine's plane, a
/// byte a cell, is a projection by name beside this trait
/// (`crate::machine::plane`, `KERNEL.22`), since its output is a raster
/// and not lines; whether rasters get a trait of their own waits for a
/// second one. A projection never edits the model.
pub trait Projection {
    /// The projection's name, as `frazaro view` selects it.
    fn name(&self) -> &str;
    /// The window's lines, in a fixed order, to the sink.
    fn project(&self, model: &Workbook, window: &Window, out: &mut dyn FnMut(&str));
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn a_window_is_one_sheet_and_one_rectangle() {
        let w = Window {
            sheet: "Model".to_string(),
            range: A1Range {
                top: 1,
                left: 1,
                bottom: 20,
                right: 6,
            },
        };
        assert_eq!(w.range.cells(), 120);
        assert_eq!(w.sheet, "Model");
    }
}
