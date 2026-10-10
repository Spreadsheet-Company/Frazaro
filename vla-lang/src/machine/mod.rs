//! The machine (KERNEL.22, 2026-10-09): a grid kept in memory and the four
//! calls an engine makes of it, `load`, `write`, `step` and `view`
//! (`Alonzo/SPEC.md` section 4). This amends `web/CALLOSUM.md` section 7
//! decision 1, the pure function with no model handle, for an engine's door
//! alone, on the measurement that decision asked for: a frame of 64,000
//! cells thirty times a second cannot rebuild its grid from text. Frazaro's
//! own doors keep the pure function.
//!
//! **The grid.** One [`Workbook`]: the rows' sheets, then a twin `X.last`
//! for every sheet `X`, in the same order, hidden, holding `X`'s values as
//! they were when the last step finished (the page's section 2). A twin is
//! an ordinary sheet to every reader, the graph, the evaluator, the record
//! and the plane, so `Screen.last!B2` is a reference like any other and
//! adds no edge, a twin's cells being values. A twin is read-only: only the
//! end of a step writes it.
//!
//! **The step.** Every formula cell is placed on a shape and evaluated at
//! its offset ([`crate::calc::shape`]), in an order that puts every cell
//! after everything it reads, with a budget in cells: a step may stop
//! before the frame is done and be called again to continue it, and a
//! write while a frame is in progress is refused, so a frame is a function
//! of the inputs it began with. Values are kept twice, the last complete
//! frame's and the frame in progress's, swapped when a frame completes, so
//! a view during a yield shows the last complete frame; then every sheet's
//! values are copied into its twin, and in the same walk the plane of each
//! sheet an engine views as one is kept, so that a view of it is a copy
//! ([`plane::Raster`], KERNEL.24). The step computes values and never
//! formulas (`AD-7`): the formula cells after a step are the formula cells
//! before it.
//!
//! **The writes.** `cell`, `formula` and `derived` rows, applied in order;
//! a refused row undoes the rows before it, so a refused write changes
//! nothing. A `derived` row is the host's carriage of the Write sheet: a
//! value that lands only in a value cell, never over a formula.
//!
//! **The refusals** are the language's catalogue's: the `grid` family for
//! the machine's own, `view-*` for the view, `calc-*` for what stops a grid
//! being stepped. An engine adds its own (its manifest, its devices) around
//! these calls and never inside them.

pub mod load;
pub mod plane;

use std::cell::Cell as Slot;
use std::collections::{BTreeMap, BTreeSet, HashMap};

use crate::calc::graph::{is_formula, Extent, Names};
use crate::calc::shape::{spell_cell, Shapes};
use crate::calc::{CellId, Computed, Env, ErrorKind, Library, Reason, Value, Values};
use crate::form::Form;
use crate::fx::FxHashMap;
use crate::messages::{raise, Refusal};
use crate::reader::read_forms;
use crate::refers::shift_a1_references;
use crate::rows::Visibility;
use crate::sheet::{
    formula_text, parse_a1_range, shared_masters, A1Range, Cell, Content, Sheet, Workbook,
};
use crate::view::{view_text_valued, window_of};

pub use load::{read_rows, sheet_name_ok, thousands, twin_base, Draft};

/// The most cells a grid holds beside its twins (KERNEL.22): a stated limit,
/// sixteen Screens of 320 by 200, so that one range row cannot ask for the
/// seventeen billion cells of Excel's sheet.
pub const CELLS: usize = 1_048_576;

/// The most grids live at once in one table (`Alonzo/SPEC.md` section 10).
pub const HANDLES: usize = 16;

/// The rows a write takes, as `grid-row-unknown` lists them.
pub const WRITE_ROWS: &str = "cell, formula and derived";

/// A refusal and the line of the row it stands on, 0 when none.
#[derive(Clone, Debug, PartialEq)]
pub struct LineRefusal {
    pub line: u32,
    pub refusal: Refusal,
}

/// What a load made: the sheets (twins not counted), the cells the rows
/// wrote, and the formula cells among them.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Loaded {
    pub sheets: usize,
    pub cells: usize,
    pub formulas: usize,
}

impl Loaded {
    /// The load's row: `(loaded <handle> <sheets> <cells> <formulas>)`.
    pub fn row(&self, handle: u32) -> String {
        format!(
            "(loaded {handle} {} {} {})",
            self.sheets, self.cells, self.formulas
        )
    }
}

/// How far a step got: the frame being computed, or just completed; the
/// cells evaluated so far in it; the frame's total; whether it is done.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Step {
    pub frame: u64,
    pub evaluated: usize,
    pub of: usize,
    pub done: bool,
}

impl Step {
    /// The step's row: `(step <frame> <evaluated> <of> done|yielded)`.
    pub fn row(&self) -> String {
        format!(
            "(step {} {} {} {})",
            self.frame,
            self.evaluated,
            self.of,
            if self.done { "done" } else { "yielded" }
        )
    }
}

/// A write's row: `(written <n>)`, the cells it changed.
pub fn written_row(n: usize) -> String {
    format!("(written {n})")
}

/// An unload's row: `(unloaded <handle>)`.
pub fn unloaded_row(handle: u32) -> String {
    format!("(unloaded {handle})")
}

/// What a view answers: the record's lines, or the plane's bytes.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum Viewed {
    Text(String),
    Bytes(Vec<u8>),
}

impl Viewed {
    /// The answer as the record's text field holds it.
    pub fn bytes(&self) -> &[u8] {
        match self {
            Viewed::Text(t) => t.as_bytes(),
            Viewed::Bytes(b) => b,
        }
    }
}

/// The grids live in one module, by handle (`Alonzo/SPEC.md` section 4.1):
/// a handle is a whole number from 1, never reused within one table, never
/// 0; at most [`HANDLES`] are live; a call with one that is not live is
/// refused by name. No clock and no randomness: the same calls give the same
/// handles.
#[derive(Debug)]
pub struct Handles<T> {
    next: u32,
    live: BTreeMap<u32, T>,
}

impl<T> Default for Handles<T> {
    fn default() -> Self {
        Handles::new()
    }
}

impl<T> Handles<T> {
    /// A table with nothing in it; the first handle it gives is 1.
    pub const fn new() -> Handles<T> {
        Handles {
            next: 1,
            live: BTreeMap::new(),
        }
    }

    /// Whether one more may be put in.
    pub fn room(&self) -> Result<(), Refusal> {
        if self.live.len() >= HANDLES || self.next == u32::MAX {
            return Err(raise(
                "grid-handle-limit",
                &[("limit", &HANDLES.to_string())],
            ));
        }
        Ok(())
    }

    /// The thing kept under a new handle.
    pub fn insert(&mut self, item: T) -> Result<u32, Refusal> {
        self.room()?;
        let handle = self.next;
        self.next += 1;
        self.live.insert(handle, item);
        Ok(handle)
    }

    /// The thing kept under a handle.
    pub fn get(&self, handle: u32) -> Result<&T, Refusal> {
        self.live.get(&handle).ok_or_else(|| unknown_handle(handle))
    }

    /// The thing kept under a handle, to change.
    pub fn get_mut(&mut self, handle: u32) -> Result<&mut T, Refusal> {
        self.live
            .get_mut(&handle)
            .ok_or_else(|| unknown_handle(handle))
    }

    /// The thing taken out, its handle never given again.
    pub fn remove(&mut self, handle: u32) -> Result<T, Refusal> {
        self.live
            .remove(&handle)
            .ok_or_else(|| unknown_handle(handle))
    }

    /// How many are live.
    pub fn live(&self) -> usize {
        self.live.len()
    }
}

fn unknown_handle(handle: u32) -> Refusal {
    raise("grid-handle-unknown", &[("handle", &handle.to_string())])
}

/// A grid an engine steps. See the module's comment.
pub struct Machine {
    book: Workbook,
    /// The rows' sheets; `book.sheets[sheets..]` are their twins, in order.
    sheets: usize,
    library: Library,
    names: Names,
    extents: Vec<Extent>,
    shapes: Shapes,
    /// The last complete frame's values, by formula cell. This map and the
    /// next two are under the crate's own hash (KERNEL.25); nothing walks
    /// them, so their order is never read.
    current: FxHashMap<CellId, Computed>,
    /// The frame in progress's values; after a frame completes, the stale
    /// values every cell overwrites before anything reads it.
    next: FxHashMap<CellId, Computed>,
    /// A twin cell left empty because its cell was not computed: the
    /// reason, which a formula reading the twin cell inherits.
    twin_reasons: FxHashMap<CellId, Reason>,
    frame: u64,
    /// While a frame is in progress, how far into the order it is.
    progress: Option<usize>,
    /// The cells of the rows' sheets.
    cells: usize,
    here: Slot<CellId>,
    /// By rows' sheet, whether an engine has viewed it as a plane
    /// (KERNEL.24): set by the view, a read, as `here` is set by an
    /// evaluation, so that only a sheet an engine draws keeps a plane.
    wanted: Vec<Slot<bool>>,
    /// By rows' sheet, its plane as the last frame ended, kept for a sheet
    /// in `wanted`; a write drops the planes of the sheets it changed, and
    /// the next frame's end keeps them again.
    planes: Vec<Option<plane::Raster>>,
}

impl std::fmt::Debug for Machine {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("Machine")
            .field("sheets", &self.sheets)
            .field("cells", &self.cells)
            .field("formulas", &self.shapes.formulas())
            .field("shapes", &self.shapes.shapes())
            .field("frame", &self.frame)
            .field("progress", &self.progress)
            .finish()
    }
}

/// What one write changed, to settle or to undo.
#[derive(Default)]
struct Log {
    /// Every cell a row set, with what it held before, in order.
    undo: Vec<(CellId, Option<Cell>)>,
    /// The cells whose content changed, each with the line of the row that
    /// changed it first.
    changed: BTreeMap<CellId, u32>,
    written: usize,
    added: usize,
    sheets: BTreeSet<usize>,
}

impl Machine {
    /// A text of rows loaded under a library: [`read_rows`], then
    /// [`Machine::new`].
    pub fn load(text: &str, library: Library) -> Result<Machine, LineRefusal> {
        let forms = read_forms(text).map_err(|refusal| LineRefusal { line: 0, refusal })?;
        let draft = read_rows(&forms)?;
        Machine::new(draft, library)
    }

    /// The machine over a draft: every formula placed on its shape, then
    /// refused when the grid cannot be stepped (a cycle, a function the
    /// library does not hold, a construct this version does not read, in
    /// that order); then the twins made and filled with the grid's values,
    /// each formula's from its `value` row. Nothing is computed: the frame
    /// is 0 and every formula waits for the first step.
    pub fn new(draft: Draft, library: Library) -> Result<Machine, LineRefusal> {
        let Draft {
            mut workbook,
            values,
            lines,
            cells,
        } = draft;
        let sheets = workbook.sheets.len();
        let names = Names::of(&workbook);
        let mut extents: Vec<Extent> = workbook.sheets.iter().map(Sheet::extent).collect();
        let shapes = Shapes::of(&workbook, sheets, &library, &names, &extents);
        if let Some((cell, refusal)) = shapes.first_refusal(&workbook, None) {
            let line = if refusal.id == "calc-cycle" {
                shapes
                    .cycles()
                    .first()
                    .and_then(|c| c.iter().filter_map(|id| lines.get(id)).min().copied())
                    .unwrap_or(0)
            } else {
                lines.get(&cell).copied().unwrap_or(0)
            };
            return Err(LineRefusal { line, refusal });
        }
        for i in 0..sheets {
            let mut twin = Sheet::new(&format!("{}.last", workbook.sheets[i].name));
            twin.visibility = Visibility::Hidden;
            workbook.sheets.push(twin);
            extents.push(None);
        }
        let current: FxHashMap<CellId, Computed> = values
            .into_iter()
            .filter(|(id, _)| shapes.holds(*id))
            .collect();
        let mut m = Machine {
            book: workbook,
            sheets,
            library,
            names,
            extents,
            shapes,
            current,
            next: FxHashMap::default(),
            twin_reasons: FxHashMap::default(),
            frame: 0,
            progress: None,
            cells,
            here: Slot::new((0, 1, 1)),
            wanted: (0..sheets).map(|_| Slot::new(false)).collect(),
            planes: vec![None; sheets],
        };
        m.snapshot();
        Ok(m)
    }

    /// What the load made, for the load's row.
    pub fn loaded(&self) -> Loaded {
        Loaded {
            sheets: self.sheets,
            cells: self.cells,
            formulas: self.shapes.formulas(),
        }
    }

    /// The grid, its twins included.
    pub fn workbook(&self) -> &Workbook {
        &self.book
    }

    /// How many of the grid's sheets are the rows'; the rest are twins.
    pub fn sheets(&self) -> usize {
        self.sheets
    }

    /// The last complete frame's number: 0 after a load.
    pub fn frame(&self) -> u64 {
        self.frame
    }

    /// Whether a frame is in progress, a step having yielded.
    pub fn in_progress(&self) -> bool {
        self.progress.is_some()
    }

    /// The frame count set, before any step of a grid loaded from a save,
    /// so that its next step is the frame after the one it was saved at:
    /// only an engine knows its Clock, so the engine says. Refused while a
    /// frame is in progress.
    pub fn resume_at(&mut self, frame: u64) -> Result<(), Refusal> {
        if let Some(done) = self.progress {
            return Err(self.during_step(done));
        }
        self.frame = frame;
        Ok(())
    }

    /// How many formula cells and how many shapes they make.
    pub fn formulas(&self) -> (usize, usize) {
        (self.shapes.formulas(), self.shapes.shapes())
    }

    fn during_step(&self, done: usize) -> Refusal {
        raise(
            "grid-write-during-step",
            &[
                ("frame", &thousands(self.frame + 1)),
                ("evaluated", &thousands(done as u64)),
                ("of", &thousands(self.shapes.order().len() as u64)),
            ],
        )
    }

    // ---- the step -------------------------------------------------------

    /// One step: at most `budget` cells evaluated, 0 meaning every one, in
    /// the order; a frame not finished is held, and the next step continues
    /// it; a frame finished swaps the values and fills the twins.
    pub fn step(&mut self, budget: usize) -> Step {
        let of = self.shapes.order().len();
        let start = self.progress.unwrap_or(0);
        let end = if budget == 0 {
            of
        } else {
            start.saturating_add(budget).min(of)
        };
        for i in start..end {
            let id = self.shapes.order()[i];
            let computed = self.evaluate(id);
            self.next.insert(id, computed);
        }
        if end < of {
            self.progress = Some(end);
            return Step {
                frame: self.frame + 1,
                evaluated: end,
                of,
                done: false,
            };
        }
        std::mem::swap(&mut self.current, &mut self.next);
        self.progress = None;
        self.frame += 1;
        self.snapshot();
        Step {
            frame: self.frame,
            evaluated: of,
            of,
            done: true,
        }
    }

    fn evaluate(&self, id: CellId) -> Computed {
        self.here.set(id);
        self.shapes.evaluate(self, &self.library, id)
    }

    /// Every sheet's values copied into its twin: a value cell as it
    /// stands, a formula cell's value of the last complete frame; a formula
    /// with no value leaves its twin cell empty, and one not computed leaves
    /// its reason beside it. A sheet an engine views as a plane has its
    /// plane kept from the same values in the same walk (KERNEL.24).
    fn snapshot(&mut self) {
        self.twin_reasons.clear();
        let sheets = self.sheets;
        let (rows, twins) = self.book.sheets.split_at_mut(sheets);
        for (s, (sheet, twin)) in rows.iter().zip(twins.iter_mut()).enumerate() {
            let t = sheets + s;
            let mut raster = if self.wanted[s].get() {
                plane::Raster::over(self.extents[s], CELLS as u64, self.planes[s].take())
            } else {
                None
            };
            let mut fits = true;
            let mut fresh: Vec<((u32, u32), Content)> = Vec::with_capacity(sheet.cells.len());
            for (&(row, col), cell) in &sheet.cells {
                // A formula's value, looked up once for the twin and the plane.
                let value = is_formula(&cell.content).then(|| self.current.get(&(s, row, col)));
                if let Some(r) = raster.as_mut() {
                    fits &= r.set(row, col, plane::kept_byte(&cell.content, value.flatten()));
                }
                let content = match value {
                    None => cell.content.clone(),
                    Some(Some(Computed::Value(v))) => value_content(v),
                    Some(Some(Computed::NotComputed(r))) => {
                        self.twin_reasons.insert((t, row, col), r.clone());
                        continue;
                    }
                    Some(None) => continue,
                };
                fresh.push(((row, col), content));
            }
            // A cell outside the extent would mean a stale extent: the walk
            // answers rather than a plane that might be wrong.
            self.planes[s] = raster.filter(|_| fits);
            let same_cells = twin.cells.len() == fresh.len()
                && twin.cells.keys().zip(&fresh).all(|(k, (f, _))| k == f);
            if same_cells {
                for (cell, (_, content)) in twin.cells.values_mut().zip(fresh) {
                    cell.content = content;
                }
            } else {
                twin.cells = fresh
                    .into_iter()
                    .map(|(k, content)| (k, Cell { content, style: 0 }))
                    .collect();
                self.extents[t] = twin.extent();
            }
        }
    }

    // ---- the view -------------------------------------------------------

    /// A window of a sheet, a twin's included, through a projection: `grid`
    /// for the record, its `value` rows from the last complete frame, or
    /// `plane` for a byte a cell. `window` `None` is the sheet's extent.
    pub fn view(
        &self,
        projection: &str,
        sheet: &str,
        window: Option<&str>,
    ) -> Result<Viewed, Refusal> {
        match projection {
            "grid" => {
                let w = window_of(&self.book, sheet, window)?;
                Ok(Viewed::Text(view_text_valued(&self.book, &w, self)))
            }
            "plane" => {
                let w = window_of(&self.book, sheet, window)?;
                let cells = w.range.cells();
                if cells > CELLS as u64 {
                    return Err(raise(
                        "view-window-too-large",
                        &[
                            ("window", &w.range.text()),
                            ("cells", &thousands(cells)),
                            ("limit", &thousands(CELLS as u64)),
                        ],
                    ));
                }
                // A rows' sheet viewed as a plane keeps its plane from the
                // next frame's end, and a kept plane answers with a copy
                // (KERNEL.24); a twin's plane is always the walk's.
                let rows_sheet = self.book.find_sheet(&w.sheet).filter(|&s| s < self.sheets);
                if let Some(s) = rows_sheet {
                    self.wanted[s].set(true);
                    if let Some(kept) = &self.planes[s] {
                        return Ok(Viewed::Bytes(kept.window(&w.range)));
                    }
                }
                Ok(Viewed::Bytes(plane::bytes(&self.book, self, &w)))
            }
            other => Err(raise("view-projection-unknown", &[("name", other)])),
        }
    }

    // ---- the write ------------------------------------------------------

    /// Rows written, one a line: `cell`, `formula` or `derived`, each over
    /// a cell or a range, applied in order; the cells changed counted. A
    /// refused row changes nothing, the rows before it undone with it.
    pub fn write(&mut self, text: &str) -> Result<usize, LineRefusal> {
        if let Some(done) = self.progress {
            return Err(LineRefusal {
                line: 0,
                refusal: self.during_step(done),
            });
        }
        let forms = read_forms(text).map_err(|refusal| LineRefusal { line: 0, refusal })?;
        let mut log = Log::default();
        for form in &forms {
            if let Err(e) = self.write_row(form, &mut log) {
                self.undo(&log);
                return Err(e);
            }
        }
        self.settle(log)
    }

    fn write_row(&mut self, form: &Form, log: &mut Log) -> Result<(), LineRefusal> {
        let Some((head, items, line)) = load::row_parts(form) else {
            let line = form.as_list().map(|l| l.line).unwrap_or(0);
            return Err(load::malformed(line, form, ""));
        };
        if !matches!(head, "cell" | "formula" | "derived") {
            return Err(load::unknown(line, head, WRITE_ROWS));
        }
        let [Form::Str(sheet), Form::Str(addr), third] = items else {
            return Err(load::malformed(line, form, head));
        };
        let Some(range) = parse_a1_range(addr) else {
            return Err(load::malformed(line, form, head));
        };
        let what = match head {
            "formula" => match third {
                Form::Str(t) if t.len() > 1 && t.starts_with('=') => None,
                _ => return Err(load::malformed(line, form, head)),
            },
            _ => Some(load::datum(third).ok_or_else(|| load::malformed(line, form, head))?),
        };
        let Some(si) = self.book.find_sheet(sheet) else {
            let sheets: Vec<&str> = self.book.sheets[..self.sheets]
                .iter()
                .map(|s| s.name.as_str())
                .collect();
            return Err(LineRefusal {
                line,
                refusal: raise(
                    "grid-sheet-unknown",
                    &[
                        ("target", &format!("{sheet}!{addr}")),
                        ("name", sheet),
                        ("sheets", &sheets.join(", ")),
                    ],
                ),
            });
        };
        if si >= self.sheets {
            return Err(load::into_twin(line, &self.book.sheets[si].name, addr));
        }
        // The cap: what the range would add to the grid.
        let count = range.cells();
        if count > CELLS as u64 {
            return Err(load::too_many(line, sheet, addr, self.cells as u64 + count));
        }
        let present = self.book.sheets[si]
            .cells
            .range((range.top, 0)..=(range.bottom, u32::MAX))
            .filter(|((_, col), _)| *col >= range.left && *col <= range.right)
            .count() as u64;
        let grows = count - present;
        let now = (self.cells + log.added) as u64;
        if now + grows > CELLS as u64 {
            return Err(load::too_many(line, sheet, addr, now + grows));
        }
        if head == "derived" {
            let formula_cell = self.book.sheets[si]
                .cells
                .range((range.top, 0)..=(range.bottom, u32::MAX))
                .find(|((_, col), c)| {
                    *col >= range.left && *col <= range.right && is_formula(&c.content)
                })
                .map(|(&(row, col), _)| (si, row, col));
            if let Some(id) = formula_cell {
                return Err(LineRefusal {
                    line,
                    refusal: raise(
                        "grid-write-derived",
                        &[
                            ("target", &format!("{sheet}!{addr}")),
                            ("cell", &spell_cell(&self.book, id)),
                        ],
                    ),
                });
            }
        }
        match (what, third) {
            (Some(value), _) => self.put_values(si, range, &value.content(), line, log),
            (None, Form::Str(text)) => self.put_formula(si, range, &text[1..], line, log),
            (None, _) => Err(load::malformed(line, form, head)),
        }
    }

    /// Set one cell, kept in the log to undo.
    fn set_cell(&mut self, id: CellId, cell: Option<Cell>, log: &mut Log) {
        let sheet = &mut self.book.sheets[id.0];
        let old = match cell {
            Some(c) => sheet.cells.insert((id.1, id.2), c),
            None => sheet.cells.remove(&(id.1, id.2)),
        };
        log.undo.push((id, old));
        log.sheets.insert(id.0);
    }

    /// Before a range is written over: every shared formula whose first cell
    /// the range holds gives its other cells outside the range their own
    /// text, since the model reads a child through its master.
    fn release_masters(&mut self, si: usize, range: A1Range, log: &mut Log) {
        let masters: Vec<(u32, u32, u32, String)> = self.book.sheets[si]
            .cells
            .range((range.top, 0)..=(range.bottom, u32::MAX))
            .filter(|((_, col), _)| *col >= range.left && *col <= range.right)
            .filter_map(|(&(row, col), c)| match &c.content {
                Content::SharedMaster {
                    text, si: index, ..
                } => Some((row, col, *index, text.clone())),
                _ => None,
            })
            .collect();
        for (m_row, m_col, index, text) in masters {
            let children: Vec<((u32, u32), u32)> = self.book.sheets[si]
                .cells
                .iter()
                .filter(|((row, col), c)| {
                    matches!(c.content, Content::SharedChild { si } if si == index)
                        && !(*row >= range.top
                            && *row <= range.bottom
                            && *col >= range.left
                            && *col <= range.right)
                })
                .map(|(&k, c)| (k, c.style))
                .collect();
            for ((row, col), style) in children {
                let moved = shift_a1_references(
                    &text,
                    i64::from(row) - i64::from(m_row),
                    i64::from(col) - i64::from(m_col),
                );
                self.set_cell(
                    (si, row, col),
                    Some(Cell {
                        content: Content::Formula(moved),
                        style,
                    }),
                    log,
                );
            }
        }
    }

    fn put_values(
        &mut self,
        si: usize,
        range: A1Range,
        content: &Content,
        line: u32,
        log: &mut Log,
    ) -> Result<(), LineRefusal> {
        self.release_masters(si, range, log);
        for row in range.top..=range.bottom {
            for col in range.left..=range.right {
                let old = self.book.sheets[si].cells.get(&(row, col)).cloned();
                if let Some(o) = &old {
                    if !is_formula(&o.content) && o.content == *content {
                        continue;
                    }
                }
                let style = old.as_ref().map(|c| c.style).unwrap_or(0);
                if old.is_none() {
                    log.added += 1;
                }
                self.set_cell(
                    (si, row, col),
                    Some(Cell {
                        content: content.clone(),
                        style,
                    }),
                    log,
                );
                log.changed.entry((si, row, col)).or_insert(line);
                log.written += 1;
            }
        }
        Ok(())
    }

    fn put_formula(
        &mut self,
        si: usize,
        range: A1Range,
        text: &str,
        line: u32,
        log: &mut Log,
    ) -> Result<(), LineRefusal> {
        // A cell whose formula would read the same is not written.
        let masters_text: HashMap<u32, (String, u32, u32)> = shared_masters(&self.book.sheets[si])
            .into_iter()
            .map(|(k, (t, r, c))| (k, (t.to_string(), r, c)))
            .collect();
        let masters: HashMap<u32, (&str, u32, u32)> = masters_text
            .iter()
            .map(|(k, (t, r, c))| (*k, (t.as_str(), *r, *c)))
            .collect();
        let mut differs: Vec<(u32, u32)> = Vec::new();
        for row in range.top..=range.bottom {
            for col in range.left..=range.right {
                let moved = shift_a1_references(
                    text,
                    i64::from(row) - i64::from(range.top),
                    i64::from(col) - i64::from(range.left),
                );
                let old_text = self.book.sheets[si]
                    .cells
                    .get(&(row, col))
                    .and_then(|c| formula_text(&c.content, row, col, &masters));
                if old_text.as_deref() != Some(crate::sheet::strip_future_prefixes(&moved).as_str())
                {
                    differs.push((row, col));
                }
            }
        }
        if differs.is_empty() {
            return Ok(());
        }
        self.release_masters(si, range, log);
        let mut olds: Vec<((u32, u32), Option<Cell>)> = Vec::with_capacity(range.cells() as usize);
        for row in range.top..=range.bottom {
            for col in range.left..=range.right {
                olds.push((
                    (row, col),
                    self.book.sheets[si].cells.get(&(row, col)).cloned(),
                ));
            }
        }
        let style_of: HashMap<(u32, u32), u32> = olds
            .iter()
            .filter_map(|(k, c)| c.as_ref().map(|c| (*k, c.style)))
            .collect();
        self.book.sheets[si].set_formula(range, text, 0);
        for (k, old) in olds {
            if old.is_none() {
                log.added += 1;
            }
            if let Some(style) = style_of.get(&k) {
                if let Some(c) = self.book.sheets[si].cells.get_mut(&k) {
                    c.style = *style;
                }
            }
            log.undo.push(((si, k.0, k.1), old));
            log.changed.entry((si, k.0, k.1)).or_insert(line);
        }
        log.sheets.insert(si);
        log.written += differs.len();
        Ok(())
    }

    /// Every cell a write set put back as it was, latest first.
    fn undo(&mut self, log: &Log) {
        for (id, old) in log.undo.iter().rev() {
            let sheet = &mut self.book.sheets[id.0];
            match old {
                Some(c) => {
                    sheet.cells.insert((id.1, id.2), c.clone());
                }
                None => {
                    sheet.cells.remove(&(id.1, id.2));
                }
            }
        }
        for &s in &log.sheets {
            self.extents[s] = self.book.sheets[s].extent();
        }
    }

    /// A write's rows all taken: the extents, the cells counted, the plan
    /// read again when a formula changed, and refused (and undone) when the
    /// grid could no longer be stepped; the changed cells' values forgotten.
    fn settle(&mut self, log: Log) -> Result<usize, LineRefusal> {
        if log.changed.is_empty() {
            return Ok(log.written);
        }
        for &s in &log.sheets {
            self.extents[s] = self.book.sheets[s].extent();
        }
        let formula_changed = log.changed.keys().any(|id| {
            self.shapes.holds(*id)
                || self.book.sheets[id.0]
                    .cells
                    .get(&(id.1, id.2))
                    .is_some_and(|c| is_formula(&c.content))
        });
        if formula_changed {
            let saved = self.shapes.clone();
            for id in log.changed.keys() {
                self.shapes.replace(&self.book, *id, &self.library);
            }
            self.shapes.plan(&self.book, &self.names, &self.extents);
            let placed: Vec<CellId> = log
                .changed
                .keys()
                .copied()
                .filter(|id| self.shapes.holds(*id))
                .collect();
            if let Some((cell, refusal)) = self.shapes.first_refusal(&self.book, Some(&placed)) {
                let line = if refusal.id == "calc-cycle" {
                    self.shapes
                        .cycles()
                        .first()
                        .and_then(|c| c.iter().filter_map(|id| log.changed.get(id)).min().copied())
                        .unwrap_or(0)
                } else {
                    log.changed.get(&cell).copied().unwrap_or(0)
                };
                self.shapes = saved;
                self.undo(&log);
                return Err(LineRefusal { line, refusal });
            }
        }
        // A write never empties a cell: it only adds cells or changes them.
        self.cells += log.added;
        for id in log.changed.keys() {
            self.current.remove(id);
            self.next.remove(id);
        }
        // A plane kept for a sheet the write changed is stale: the walk
        // answers until the next frame's end keeps it again (KERNEL.24). A
        // refused write never reaches here, having changed nothing.
        for &s in &log.sheets {
            if let Some(kept) = self.planes.get_mut(s) {
                *kept = None;
            }
        }
        Ok(log.written)
    }
}

impl Values for Machine {
    fn computed(&self, id: CellId) -> Option<&Computed> {
        self.current.get(&id)
    }
}

impl Env for Machine {
    fn model(&self) -> &Workbook {
        &self.book
    }

    fn here(&self) -> CellId {
        self.here.get()
    }

    fn read(&self, id: CellId) -> Result<Value, Reason> {
        // A twin's cell is a value or nothing: no twin cell is ever placed or
        // valued, a twin being filled at a frame's end and refused to every
        // write, so its read goes straight to the twin (KERNEL.25), past two
        // lookups that could only miss. Most of a frame's reads are of a twin.
        if id.0 < self.sheets {
            if let Some(c) = self.next.get(&id) {
                return match c {
                    Computed::Value(v) => Ok(v.clone()),
                    Computed::NotComputed(r) => Err(r.clone()),
                };
            }
            if self.shapes.holds(id) {
                // A formula with no value in this frame can only be read from
                // a cycle, which the plan refuses before it is held.
                return Err(Reason::Cycle);
            }
        }
        let Some(cell) = self
            .book
            .sheets
            .get(id.0)
            .and_then(|s| s.cells.get(&(id.1, id.2)))
        else {
            return match self.twin_reasons.get(&id) {
                Some(r) => Err(r.clone()),
                None => Ok(Value::Empty),
            };
        };
        Ok(match &cell.content {
            Content::Text(t) => Value::Text(t.clone()),
            Content::Number(n) => Value::Number(*n),
            Content::Bool(b) => Value::Bool(*b),
            Content::Error(e) => Value::Error(ErrorKind::parse(e).unwrap_or(ErrorKind::Value)),
            _ => Value::Empty,
        })
    }

    fn extent(&self, sheet: usize) -> Extent {
        self.extents.get(sheet).copied().flatten()
    }

    fn name(&self, name: &str, scope: Option<usize>, home: usize) -> Option<&str> {
        self.names.resolve(name, scope, home)
    }
}

/// A computed value as a twin's cell holds it.
fn value_content(v: &Value) -> Content {
    match v {
        Value::Number(n) => Content::Number(*n),
        Value::Text(t) => Content::Text(t.clone()),
        Value::Bool(b) => Content::Bool(*b),
        Value::Error(k) => Content::Error(k.text().to_string()),
        Value::Empty => Content::Number(0.0),
    }
}

#[cfg(test)]
mod tests {
    //! The four calls' tests, written from `Alonzo/SPEC.md` sections 4 and
    //! 13 and nothing else: each call's row, each refusal by its id, and
    //! the page's free oracles, load the inverse of view, a save resumes, a
    //! step writes no formula, one write per cell, a derived write lands or
    //! is refused, a budget never changes a frame; and Life against a Life
    //! written here, cell for cell.
    use super::*;
    use crate::calc::library;
    use crate::calc::shape::Shapes;
    use crate::projection::Window;
    use crate::reader::read_forms;
    use crate::view::view_text_valued;

    fn load(text: &str) -> Machine {
        Machine::load(text, library::language())
            .unwrap_or_else(|e| panic!("line {}: {}", e.line, e.refusal.text))
    }

    fn refused(text: &str) -> LineRefusal {
        match Machine::load(text, library::language()) {
            Ok(m) => panic!("loaded: {m:?}"),
            Err(e) => e,
        }
    }

    fn grid(m: &Machine, sheet: &str, window: Option<&str>) -> String {
        match m.view("grid", sheet, window) {
            Ok(Viewed::Text(t)) => t,
            other => panic!("{other:?}"),
        }
    }

    fn plane(m: &Machine, sheet: &str, window: &str) -> Vec<u8> {
        match m.view("plane", sheet, Some(window)) {
            Ok(Viewed::Bytes(b)) => b,
            other => panic!("{other:?}"),
        }
    }

    /// A formula cell's computed value, by its sheet and address.
    fn at(m: &Machine, sheet: &str, addr: &str) -> Option<Computed> {
        let i = m.workbook().find_sheet(sheet).unwrap();
        let r = parse_a1_range(addr).unwrap();
        m.computed((i, r.top, r.left)).cloned()
    }

    fn number(m: &Machine, sheet: &str, addr: &str) -> f64 {
        match at(m, sheet, addr) {
            Some(Computed::Value(Value::Number(n))) => n,
            other => panic!("{sheet}!{addr}: {other:?}"),
        }
    }

    /// A value cell's content, a twin's included.
    fn content(m: &Machine, sheet: &str, addr: &str) -> Option<Content> {
        let i = m.workbook().find_sheet(sheet).unwrap();
        let r = parse_a1_range(addr).unwrap();
        m.workbook().sheets[i]
            .cells
            .get(&(r.top, r.left))
            .map(|c| c.content.clone())
    }

    /// The formula rows of every sheet's record, the twins' included.
    fn formula_rows(m: &Machine) -> Vec<String> {
        let mut out = Vec::new();
        for s in &m.workbook().sheets {
            for line in grid(m, &s.name, None).lines() {
                if line.starts_with("(formula ") {
                    out.push(line.to_string());
                }
            }
        }
        out
    }

    /// A save, as `Alonzo/SPEC.md` section 7.5 makes one: every sheet that
    /// is not a twin, viewed whole, one record after another.
    fn save(m: &Machine) -> String {
        let mut out = String::new();
        for s in &m.workbook().sheets[..m.sheets()] {
            out.push_str(&grid(m, &s.name, None));
        }
        out
    }

    /// Values held for a view of a draft.
    #[derive(Debug)]
    struct Held(HashMap<CellId, Computed>);

    impl Values for Held {
        fn computed(&self, id: CellId) -> Option<&Computed> {
            self.0.get(&id)
        }
    }

    const COUNTER: &str = "\
(sheet \"Board\" visible)
(formula \"Board\" \"A1\" \"=Board.last!A1+1\")
(formula \"Board\" \"B1\" \"=A1*10\")
(formula \"Board\" \"C1\" \"=SUM(A1:B1)\")
(cell \"Board\" \"D1\" 5)
(formula \"Board\" \"E1\" \"=D1+C1\")
";

    #[test]
    fn handles_start_at_one_are_never_reused_and_hold_sixteen() {
        let mut h: Handles<u8> = Handles::new();
        for n in 1..=16u32 {
            assert_eq!(h.insert(n as u8), Ok(n));
        }
        let full = h.insert(0).unwrap_err();
        assert_eq!(full.id, "grid-handle-limit");
        assert!(full.text.contains("16 grids"), "{}", full.text);
        assert_eq!(h.remove(3), Ok(3));
        let gone = h.get(3).unwrap_err();
        assert_eq!(gone.id, "grid-handle-unknown");
        assert!(gone.text.contains("handle 3"), "{}", gone.text);
        assert_eq!(h.get(0).unwrap_err().id, "grid-handle-unknown");
        assert_eq!(h.insert(9), Ok(17), "a handle is never given twice");
        assert_eq!(h.live(), 16);
        assert_eq!(h.remove(3).unwrap_err().id, "grid-handle-unknown");
    }

    #[test]
    fn a_load_answers_its_counts_and_computes_nothing() {
        let m = load(COUNTER);
        assert_eq!(
            m.loaded(),
            Loaded {
                sheets: 1,
                cells: 5,
                formulas: 4
            }
        );
        assert_eq!(m.loaded().row(7), "(loaded 7 1 5 4)");
        assert_eq!(m.frame(), 0);
        assert!(at(&m, "Board", "A1").is_none(), "nothing computed at load");
        let record = grid(&m, "Board", None);
        assert!(!record.contains("(value "), "{record}");
        assert!(record.starts_with(
            "(sheet \"Board\" visible)\n(sheet \"Board.last\" hidden)\n(window \"Board\" \"A1:E1\")\n"
        ));
    }

    #[test]
    fn a_step_reads_the_last_frame_from_the_twin_and_this_frame_from_the_grid() {
        let mut m = load(COUNTER);
        let s = m.step(0);
        assert_eq!(s.row(), "(step 1 4 4 done)");
        assert_eq!(number(&m, "Board", "A1"), 1.0, "an empty twin reads as 0");
        assert_eq!(number(&m, "Board", "B1"), 10.0, "this frame's A1");
        assert_eq!(number(&m, "Board", "C1"), 11.0);
        assert_eq!(number(&m, "Board", "E1"), 16.0);
        // The twin holds the frame just computed, values only.
        assert_eq!(content(&m, "Board.last", "A1"), Some(Content::Number(1.0)));
        assert_eq!(content(&m, "Board.last", "D1"), Some(Content::Number(5.0)));
        assert_eq!(m.step(0).row(), "(step 2 4 4 done)");
        assert_eq!(number(&m, "Board", "A1"), 2.0);
        assert_eq!(number(&m, "Board", "B1"), 20.0);
        let twin = grid(&m, "Board.last", None);
        assert!(twin.contains("(cell \"Board.last\" \"B1\" 20)"), "{twin}");
        assert!(!twin.contains("(formula "), "a twin holds values only");
        // A record of the grid shows each formula's value after it.
        let record = grid(&m, "Board", None);
        assert!(record.contains(
            "(formula \"Board\" \"A1\" \"=Board.last!A1+1\")\n(value \"Board\" \"A1\" 2)\n"
        ));
    }

    #[test]
    fn a_budget_yields_resumes_and_never_changes_a_frame() {
        let mut whole = load(COUNTER);
        let mut chunked = load(COUNTER);
        for frame in 1..=3u64 {
            assert!(whole.step(0).done);
            let first = chunked.step(3);
            assert_eq!(first.row(), format!("(step {frame} 3 4 yielded)"));
            // A write while the frame is in progress is refused, and changes
            // nothing.
            let r = chunked.write("(cell \"Board\" \"D1\" 9)").unwrap_err();
            assert_eq!(
                (r.line, r.refusal.id.as_str()),
                (0, "grid-write-during-step")
            );
            assert!(
                r.refusal
                    .text
                    .starts_with(&format!("Frame {frame} is being computed, 3 of 4 cells")),
                "{}",
                r.refusal.text
            );
            // A view meanwhile shows the last complete frame.
            let last = (frame > 1).then(|| Computed::Value(Value::Number((frame - 1) as f64)));
            assert_eq!(at(&chunked, "Board", "A1"), last);
            assert_eq!(chunked.step(3).row(), format!("(step {frame} 4 4 done)"));
            assert_eq!(save(&whole), save(&chunked), "frame {frame}");
        }
        assert_eq!(content(&chunked, "Board", "D1"), Some(Content::Number(5.0)));
    }

    #[test]
    fn a_step_writes_no_formula() {
        let mut m = load(COUNTER);
        let before = formula_rows(&m);
        for _ in 0..3 {
            m.step(0);
        }
        assert_eq!(
            formula_rows(&m),
            before,
            "AD-7: the formula cells are unchanged"
        );
        assert_eq!(before.len(), 4);
    }

    #[test]
    fn one_write_per_cell_names_both_lines() {
        let e = refused("(cell \"S\" \"A1\" 1)\n(cell \"S\" \"B1\" 2)\n(cell \"S\" \"A1\" 1)\n");
        assert_eq!(
            (e.line, e.refusal.id.as_str()),
            (3, "grid-cell-written-twice")
        );
        // The row's own line is the answer's field; the words name the
        // first line, which the answer does not carry (KERNEL.23).
        assert!(
            e.refusal
                .text
                .starts_with("S!A1 was already written by line 1;"),
            "{}",
            e.refusal.text
        );
        let e = refused("(formula \"S\" \"A1:C3\" \"=1\")\n(cell \"S\" \"B2\" 1)\n");
        assert_eq!(
            (e.line, e.refusal.id.as_str()),
            (2, "grid-cell-written-twice")
        );
        assert!(e.refusal.text.contains("S!B2"), "{}", e.refusal.text);
        // A setting given twice is taken when it says the same thing.
        let m = load(
            "(sheet \"S\" visible)\n(sheet \"S\" visible)\n(format 0 none general nowrap)\n(format 0 none general nowrap)\n(cell \"S\" \"A1\" 1)\n",
        );
        assert_eq!(m.loaded().cells, 1);
        let e = refused("(format 1 none general nowrap)\n(format 1 none text nowrap)\n");
        assert_eq!(
            (e.line, e.refusal.id.as_str()),
            (2, "grid-cell-written-twice")
        );
        assert!(e.refusal.text.contains("(format 1)"), "{}", e.refusal.text);
        let e = refused("(cell \"S\" \"A1\" 1)\n(formula \"S\" \"B1\" \"=A1\")\n(value \"S\" \"B1\" 1)\n(value \"S\" \"B1\" 1)\n");
        assert_eq!(
            (e.line, e.refusal.id.as_str()),
            (4, "grid-cell-written-twice")
        );
    }

    #[test]
    fn a_load_refuses_what_it_cannot_hold_by_name() {
        let cases: &[(&str, u32, &str)] = &[
            ("(chart \"S\" \"A1\")\n", 1, "grid-row-unknown"),
            ("(cell \"S\" \"A0\" 1)\n", 1, "grid-row-malformed"),
            ("(cell \"S\" \"A1\" inf)\n", 1, "grid-row-malformed"),
            (
                "(cell \"S\" \"A1\" (date \"2026-10-09\"))\n",
                1,
                "grid-row-malformed",
            ),
            ("(formula \"S\" \"A1\" \"1+1\")\n", 1, "grid-row-malformed"),
            ("(value \"S\" \"A1\" 1)\n", 1, "grid-row-malformed"),
            (
                "(cell \"S\" \"A1\" 1)\n(value \"S\" \"A1\" 1)\n",
                2,
                "grid-row-malformed",
            ),
            ("\n\n(cell \"Board.last\" \"A1\" 1)\n", 3, "grid-write-last"),
            // A setting naming a twin is the name's refusal, as a sheet row
            // naming one alone is: a twin copies values, never settings.
            (
                "(sheet \"Board\" visible)\n(gridlines \"Board.last\" off)\n",
                2,
                "grid-sheet-name-invalid",
            ),
            (
                "(column \"Board.last\" \"A\" 10 shown none)\n",
                1,
                "grid-sheet-name-invalid",
            ),
            (
                "(cell \"Twenty-seven characters, ab\" \"A1\" 1)\n",
                1,
                "grid-sheet-name-invalid",
            ),
            ("(cell \"a:b\" \"A1\" 1)\n", 1, "grid-sheet-name-invalid"),
            ("(cell \".last\" \"A1\" 1)\n", 1, "grid-sheet-name-invalid"),
            (
                "(sheet \"Board.last\" hidden)\n",
                1,
                "grid-sheet-name-invalid",
            ),
            (
                "(cell \"S\" \"A1:XFD1048576\" 0)\n",
                1,
                "grid-too-many-cells",
            ),
            (
                "(formula \"S\" \"A1\" \"=B1\")\n(formula \"S\" \"B1\" \"=A1\")\n",
                1,
                "calc-cycle",
            ),
            (
                "(cell \"S\" \"A1\" 1)\n(formula \"S\" \"B2\" \"=NOW()\")\n",
                2,
                "calc-function-not-computed",
            ),
            (
                "(formula \"S\" \"C3\" \"={1,2}\")\n",
                1,
                "calc-construct-not-read",
            ),
            (
                "(name \"Rate\" \"0.2\")\n(name \"rate\" \"0.3\")\n",
                2,
                "grid-cell-written-twice",
            ),
            ("(cell \"S\" \"A1\" 1)\n(oops\n", 0, "vla-unbalanced-parens"),
        ];
        for (text, line, id) in cases {
            let e = refused(text);
            assert_eq!((e.line, e.refusal.id.as_str()), (*line, *id), "{text}");
            // The line is the answer's field and never the words' opening.
            assert!(!e.refusal.text.starts_with("Line "), "{}", e.refusal.text);
        }
        let e = refused("(cell \"S\" \"A1\" 1)\n(formula \"S\" \"B2\" \"=NOW()\")\n");
        assert!(
            e.refusal.text.contains("S!B2") && e.refusal.text.contains("NOW"),
            "{}",
            e.refusal.text
        );
        let e = refused("(cell \"S\" \"A1\" (date \"2026-10-09\"))\n");
        assert!(
            e.refusal.text.contains("(error \"#N/A\")"),
            "{}",
            e.refusal.text
        );
        // A twin's own declaration is taken beside its sheet, as a save
        // lists the twins in every record.
        let m = load(
            "(sheet \"Board\" visible)\n(sheet \"Board.last\" hidden)\n(cell \"Board\" \"A1\" 1)\n",
        );
        assert_eq!(m.sheets(), 1);
        assert_eq!(m.workbook().sheets.len(), 2);
    }

    #[test]
    fn a_write_changes_cells_counts_them_and_a_refused_write_changes_nothing() {
        let mut m = load(COUNTER);
        m.step(0);
        assert_eq!(
            m.write("(cell \"Board\" \"D1\" 5)"),
            Ok(0),
            "the value it holds"
        );
        assert_eq!(
            m.write("(cell \"Board\" \"D1\" 6)\n(cell \"Board\" \"F1:G2\" \"x\")"),
            Ok(5)
        );
        assert_eq!(m.loaded().cells, 9);
        let before = save(&m);
        // The second row is refused, and the first is undone with it.
        let e = m
            .write("(cell \"Board\" \"D1\" 7)\n(cell \"Nowhere\" \"A1\" 1)")
            .unwrap_err();
        assert_eq!((e.line, e.refusal.id.as_str()), (2, "grid-sheet-unknown"));
        assert!(
            e.refusal.text.contains("Nowhere!A1") && e.refusal.text.contains("it holds Board."),
            "{}",
            e.refusal.text
        );
        assert_eq!(save(&m), before, "a refused write changes nothing");
        let e = m.write("(cell \"Board.last\" \"A1\" 1)").unwrap_err();
        assert_eq!((e.line, e.refusal.id.as_str()), (1, "grid-write-last"));
        // The cell to write instead, not only its sheet (KERNEL.23).
        assert!(
            e.refusal
                .text
                .starts_with("Board.last!A1 is a cell of Board.last,")
                && e.refusal.text.contains("write Board!A1,"),
            "{}",
            e.refusal.text
        );
        let e = m.write("(sheet \"Board\" visible)").unwrap_err();
        assert_eq!(e.refusal.id, "grid-row-unknown");
        assert!(
            e.refusal.text.contains("cell, formula and derived"),
            "{}",
            e.refusal.text
        );
        let e = m.write("(cell \"Board\" \"A1:XFD1048576\" 0)").unwrap_err();
        assert_eq!(e.refusal.id, "grid-too-many-cells");
        assert_eq!(save(&m), before);
        // A value over a formula: the formula goes, and the plan with it.
        assert_eq!(m.write("(cell \"Board\" \"B1\" 100)"), Ok(1));
        m.step(0);
        assert_eq!(number(&m, "Board", "C1"), 102.0, "A1 is 2, B1 holds 100");
        assert_eq!(m.formulas().0, 3);
    }

    #[test]
    fn a_derived_write_lands_in_a_value_cell_or_is_refused_and_the_frame_goes_on() {
        let mut m = load(COUNTER);
        m.step(0);
        assert_eq!(m.write("(derived \"Board\" \"D1\" 7)"), Ok(1));
        let e = m.write("(derived \"Board\" \"C1:D1\" 1)").unwrap_err();
        assert_eq!((e.line, e.refusal.id.as_str()), (1, "grid-write-derived"));
        assert!(
            e.refusal.text.contains("Board!C1:D1")
                && e.refusal.text.contains("Board!C1 holds a formula"),
            "{}",
            e.refusal.text
        );
        assert_eq!(content(&m, "Board", "D1"), Some(Content::Number(7.0)));
        assert_eq!(
            m.write("(derived \"Board.last\" \"D1\" 1)")
                .unwrap_err()
                .refusal
                .id,
            "grid-write-last"
        );
        assert_eq!(
            m.write("(derived \"Nowhere\" \"A1\" 1)")
                .unwrap_err()
                .refusal
                .id,
            "grid-sheet-unknown"
        );
        assert_eq!(m.step(0).row(), "(step 2 4 4 done)");
        assert_eq!(number(&m, "Board", "E1"), 29.0, "D1 7, C1 2 + 20");
    }

    #[test]
    fn a_formula_write_is_planned_again_and_refused_when_it_cannot_be_stepped() {
        let mut m = load(COUNTER);
        m.step(0);
        let before = save(&m);
        let e = m.write("(formula \"Board\" \"D1\" \"=E1\")").unwrap_err();
        assert_eq!((e.line, e.refusal.id.as_str()), (1, "calc-cycle"));
        assert!(
            e.refusal.text.contains("Board!D1, Board!E1"),
            "{}",
            e.refusal.text
        );
        assert_eq!(save(&m), before);
        let e = m
            .write("\n(formula \"Board\" \"F1\" \"=RAND()\")")
            .unwrap_err();
        assert_eq!(
            (e.line, e.refusal.id.as_str()),
            (2, "calc-function-not-computed")
        );
        assert_eq!(save(&m), before);
        assert_eq!(m.write("(formula \"Board\" \"F1:F3\" \"=D1*2\")"), Ok(3));
        assert!(
            at(&m, "Board", "F1").is_none(),
            "no value before the next step"
        );
        assert_eq!(m.formulas(), (7, 5));
        m.step(0);
        assert_eq!(number(&m, "Board", "F1"), 10.0);
        assert_eq!(number(&m, "Board", "F2"), 0.0, "D2 is empty");
        // The same formulas again change nothing.
        assert_eq!(m.write("(formula \"Board\" \"F1:F3\" \"=D1*2\")"), Ok(0));
    }

    /// The fact a twin's read rests on (KERNEL.25): no twin cell is placed
    /// or valued, so a read of one need look nowhere but the twin.
    fn no_twin_cell_is_placed_or_valued(m: &Machine, when: &str) {
        assert_eq!(m.shapes.order().len(), m.shapes.formulas(), "{when}");
        for id in m.shapes.order() {
            assert!(id.0 < m.sheets, "{when}: {id:?} placed on a twin");
        }
        for id in m.current.keys().chain(m.next.keys()) {
            assert!(id.0 < m.sheets, "{when}: {id:?} valued on a twin");
        }
        for (s, sheet) in m.book.sheets.iter().enumerate().skip(m.sheets) {
            for &(row, col) in sheet.cells.keys() {
                assert!(!m.shapes.holds((s, row, col)), "{when}: {s} {row} {col}");
            }
        }
        for id in m.twin_reasons.keys() {
            assert!(id.0 >= m.sheets, "{when}: a reason beside {id:?}");
        }
    }

    #[test]
    fn a_twin_cell_is_never_placed_or_valued() {
        let mut m = load(COUNTER);
        no_twin_cell_is_placed_or_valued(&m, "the load");
        m.step(0);
        no_twin_cell_is_placed_or_valued(&m, "a step");
        for (text, written) in [
            ("(cell \"Board\" \"D1\" 6)", true),
            ("(cell \"Board\" \"F1:G2\" \"x\")", true),
            ("(formula \"Board\" \"H1:H3\" \"=Board.last!A1*2\")", true),
            ("(derived \"Board\" \"D1\" 7)", true),
            ("(cell \"Board\" \"B1\" 100)", true),
            ("(cell \"Board.last\" \"A1\" 1)", false),
            ("(formula \"Board.last\" \"A1\" \"=1\")", false),
            ("(formula \"Board\" \"D1\" \"=E1\")", false),
        ] {
            assert_eq!(m.write(text).is_ok(), written, "{text}");
            no_twin_cell_is_placed_or_valued(&m, text);
            m.step(0);
            no_twin_cell_is_placed_or_valued(&m, text);
        }
        // A frame in progress, half its values in `next`.
        assert!(!m.step(2).done);
        no_twin_cell_is_placed_or_valued(&m, "a yield");
        m.step(0);
        // A cell not computed leaves its reason beside its twin's cell.
        assert_eq!(m.write("(formula \"Board\" \"J1\" \"=F1:G1*2\")"), Ok(1));
        m.step(0);
        assert!(!m.twin_reasons.is_empty());
        no_twin_cell_is_placed_or_valued(&m, "a reason kept");
        let twin = m.book.find_sheet("Board.last").unwrap();
        assert_eq!(
            m.read((twin, 1, 10)),
            Err(Reason::Construct("F1:G1".to_string())),
            "a formula reading the twin's cell inherits the reason"
        );
        assert_eq!(m.read((twin, 1, 4)), Ok(Value::Number(7.0)));
        assert_eq!(m.read((twin, 9, 9)), Ok(Value::Empty));
    }

    #[test]
    fn a_range_where_one_value_is_wanted_is_labelled_as_written() {
        // KERNEL.25: the label is borrowed from the parse, moved with a cell
        // at an offset from its shape's first, and owned when read through
        // a name, whose reference does not move.
        let mut m = load(concat!(
            "(cell \"S\" \"A1:B3\" 1)\n",
            "(name \"Wide\" \"S!$A$1:$B$1\")\n",
            "(formula \"S\" \"D1:D2\" \"=A1:B1*2\")\n",
            "(formula \"S\" \"E1:E2\" \"=Wide*2\")\n",
        ));
        assert_eq!(m.formulas(), (4, 2));
        m.step(0);
        let label = |s: &str| Some(Computed::NotComputed(Reason::Construct(s.to_string())));
        assert_eq!(at(&m, "S", "D1"), label("A1:B1"));
        assert_eq!(at(&m, "S", "D2"), label("A2:B2"), "moved with its cell");
        assert_eq!(at(&m, "S", "E1"), label("S!$A$1:$B$1"));
        assert_eq!(
            at(&m, "S", "E2"),
            label("S!$A$1:$B$1"),
            "a name does not move"
        );
    }

    #[test]
    fn a_write_over_a_shared_formula_s_first_cell_keeps_the_rest() {
        let mut m = load("(cell \"S\" \"A1:A3\" 2)\n(formula \"S\" \"B1:B3\" \"=A1*10\")\n");
        assert_eq!(m.formulas(), (3, 1), "one shape over the range");
        assert_eq!(m.write("(cell \"S\" \"B1\" 5)"), Ok(1));
        let record = grid(&m, "S", None);
        assert!(
            record.contains("(formula \"S\" \"B2\" \"=A2*10\")"),
            "{record}"
        );
        assert!(
            record.contains("(formula \"S\" \"B3\" \"=A3*10\")"),
            "{record}"
        );
        m.step(0);
        assert_eq!(number(&m, "S", "B3"), 20.0);
        assert_eq!(content(&m, "S", "B1"), Some(Content::Number(5.0)));
    }

    #[test]
    fn the_plane_is_a_byte_a_cell_and_agrees_with_the_record() {
        let mut m = load(
            "(cell \"S\" \"A1\" 0)\n(cell \"S\" \"B1\" 1)\n(cell \"S\" \"C1\" 254)\n(cell \"S\" \"D1\" 255)\n(cell \"S\" \"E1\" 1.5)\n(cell \"S\" \"F1\" -1)\n(cell \"S\" \"G1\" \"x\")\n(cell \"S\" \"H1\" true)\n(cell \"S\" \"I1\" (error \"#N/A\"))\n(formula \"S\" \"A2\" \"=B1+2\")\n(formula \"S\" \"B2\" \"=1/0\")\n",
        );
        let row1 = [0u8, 1, 254, 255, 255, 255, 255, 255, 255, 0];
        assert_eq!(plane(&m, "S", "A1:J1"), row1);
        assert_eq!(
            plane(&m, "S", "A2:C2"),
            [0, 0, 0],
            "no value before the first step"
        );
        m.step(0);
        assert_eq!(plane(&m, "S", "A2:C2"), [3, 255, 0]);
        // Past the extent the window reads blank.
        assert_eq!(plane(&m, "S", "J9:K10"), [0, 0, 0, 0]);
        // The record of the same window holds the same values.
        let record = grid(&m, "S", Some("A2:C2"));
        assert!(record.contains("(value \"S\" \"A2\" 3)"), "{record}");
        assert!(
            record.contains("(value \"S\" \"B2\" (error \"#DIV/0!\"))"),
            "{record}"
        );
    }

    /// A sheet of every kind of byte (KERNEL.24): whole numbers in and out
    /// of range, a text, a truth value, an error, a fraction; a counter, a
    /// product that passes 254, a range where one value is wanted (not
    /// computed), a formula turning to a text, an error, a shared formula,
    /// a formula reading an empty cell and one reading a twin's empty cell;
    /// and a second sheet.
    const PLANES: &str = "\
(cell \"Board\" \"B2\" 7)
(cell \"Board\" \"C2\" 254)
(cell \"Board\" \"D2\" \"x\")
(cell \"Board\" \"E2\" true)
(cell \"Board\" \"F2\" (error \"#N/A\"))
(cell \"Board\" \"G2\" 1.5)
(formula \"Board\" \"B3\" \"=Board.last!B3+1\")
(formula \"Board\" \"C3\" \"=B3*100\")
(formula \"Board\" \"D3\" \"=B2:C2\")
(formula \"Board\" \"E3\" \"=IF(B3>1,D2,B3)\")
(formula \"Board\" \"F3\" \"=1/0\")
(formula \"Board\" \"B4:D4\" \"=B3+B2\")
(formula \"Board\" \"F5\" \"=Z99\")
(formula \"Board\" \"G5\" \"=Board.last!D3\")
(cell \"Other\" \"A1\" 3)
(formula \"Other\" \"B1\" \"=A1*2\")
";

    /// Windows inside, across and past `PLANES`'s Board, whose extent is
    /// B2:G5 until a write grows it.
    const BOARD_WINDOWS: &[&str] = &[
        "B2:G5", "C3:E4", "A1:D3", "F4:J7", "L9:M10", "D3", "A1:Z20", "G1:G9",
    ];

    /// Whether a sheet's plane is kept (KERNEL.24).
    fn kept(m: &Machine, sheet: &str) -> bool {
        let s = m.workbook().find_sheet(sheet).unwrap();
        m.planes.get(s).is_some_and(Option::is_some)
    }

    /// The plane a view answers held to the walk, the reference, over each
    /// window; then the sheet's plane kept, or not, as said, so that stale
    /// bytes fail on the bytes and a plane never kept fails on the flag.
    fn walked(m: &Machine, sheet: &str, windows: &[&str], kept_now: bool) {
        let was_kept = kept(m, sheet);
        for w in windows {
            let window = window_of(m.workbook(), sheet, Some(w)).unwrap();
            let walk = plane::bytes(m.workbook(), m, &window);
            assert_eq!(plane(m, sheet, w), walk, "{sheet}!{w}, frame {}", m.frame());
        }
        assert_eq!(was_kept, kept_now, "{sheet} kept, frame {}", m.frame());
    }

    #[test]
    fn a_kept_plane_is_the_walk_after_every_step_and_every_write() {
        let mut m = load(PLANES);
        // A load keeps no plane, no sheet having been viewed as one; the
        // view marks Board, and the next frame's end keeps its plane.
        walked(&m, "Board", BOARD_WINDOWS, false);
        m.step(0);
        walked(&m, "Board", BOARD_WINDOWS, true);
        assert!(!kept(&m, "Other"), "a sheet never viewed keeps no plane");
        // Not computed reads 255, and a number past 254 does too.
        m.step(0);
        assert!(
            matches!(at(&m, "Board", "D3"), Some(Computed::NotComputed(_))),
            "{:?}",
            at(&m, "Board", "D3")
        );
        assert_eq!(plane(&m, "Board", "B3:D3"), [2, 200, 255]);
        m.step(0);
        assert_eq!(plane(&m, "Board", "B3:D3"), [3, 255, 255]);
        walked(&m, "Board", BOARD_WINDOWS, true);
        // Each kind of write drops the plane of the sheet it changes; the
        // walk answers until the next frame's end keeps it again.
        for row in [
            "(cell \"Board\" \"B2\" 9)",
            "(formula \"Board\" \"C2\" \"=B2+1\")",
            "(derived \"Board\" \"G2\" 3)",
            "(cell \"Board\" \"C5:D6\" 2)",
            "(cell \"Board\" \"I1\" 1)",
        ] {
            assert!(m.write(row).is_ok(), "{row}");
            walked(&m, "Board", BOARD_WINDOWS, false);
            m.step(0);
            walked(&m, "Board", BOARD_WINDOWS, true);
        }
        // A refused write changes nothing, so the plane stands; so does a
        // write to another sheet.
        assert!(m
            .write("(cell \"Board\" \"B2\" 1)\n(cell \"Nowhere\" \"A1\" 1)")
            .is_err());
        walked(&m, "Board", BOARD_WINDOWS, true);
        assert_eq!(m.write("(cell \"Other\" \"A1\" 4)"), Ok(1));
        walked(&m, "Board", BOARD_WINDOWS, true);
        // During a yielded frame the plane is the last complete frame's.
        let last = plane(&m, "Board", "A1:Z20");
        assert!(!m.step(2).done);
        walked(&m, "Board", BOARD_WINDOWS, true);
        assert_eq!(plane(&m, "Board", "A1:Z20"), last);
        while !m.step(2).done {}
        walked(&m, "Board", BOARD_WINDOWS, true);
        assert_ne!(plane(&m, "Board", "A1:Z20"), last, "the counter moved");
        // A frame count set for a save changes no byte.
        m.resume_at(100).unwrap();
        walked(&m, "Board", BOARD_WINDOWS, true);
        // A twin's plane is the walk's and is never kept.
        walked(&m, "Board.last", BOARD_WINDOWS, false);
        m.step(0);
        walked(&m, "Board.last", BOARD_WINDOWS, false);
    }

    #[test]
    fn a_kept_plane_from_a_saves_values_is_the_walk() {
        let mut m = load(PLANES);
        for _ in 0..3 {
            m.step(0);
        }
        let mut again = load(&save(&m));
        // The load filled the twins from the save's value rows with no
        // sheet marked; marked now, the same fill keeps the plane from
        // those values, as the end of a frame would.
        walked(&again, "Board", BOARD_WINDOWS, false);
        again.snapshot();
        walked(&again, "Board", BOARD_WINDOWS, true);
    }

    #[test]
    fn lifes_plane_is_kept_through_the_clocks_writes() {
        let windows = [
            "B2:BK40",
            "A1:BL41",
            "A1:C3",
            "BJ39:BM42",
            "CA50:CB51",
            "C3:D5",
        ];
        let mut m = load(LIFE);
        walked(&m, "Screen", &windows, false);
        for frame in 1..=6u64 {
            // The Clock's write leaves the Screen's plane standing.
            assert_eq!(m.write(&format!("(cell \"Clock\" \"B1\" {frame})")), Ok(1));
            walked(&m, "Screen", &windows, frame > 1);
            assert!(m.step(0).done);
            walked(&m, "Screen", &windows, true);
        }
    }

    #[test]
    fn a_sheet_past_a_planes_limit_is_walked() {
        let mut m = load("(cell \"Big\" \"A1\" 1)\n(cell \"Big\" \"XFD1048576\" 2)\n");
        assert_eq!(plane(&m, "Big", "A1:B2"), [1, 0, 0, 0]);
        m.step(0);
        assert!(!kept(&m, "Big"), "its extent passes the limit");
        assert_eq!(plane(&m, "Big", "XFC1048575:XFD1048576"), [0, 0, 0, 2]);
    }

    #[test]
    fn a_view_refuses_what_it_does_not_have_by_name() {
        let m = load(COUNTER);
        let e = m.view("raster", "Board", None).unwrap_err();
        assert_eq!(e.id, "view-projection-unknown");
        assert!(
            e.text.starts_with("raster is not a projection"),
            "{}",
            e.text
        );
        let e = m.view("grid", "Nowhere", None).unwrap_err();
        assert_eq!(e.id, "view-sheet-unknown");
        assert_eq!(
            e.text,
            "The grid holds no sheet named Nowhere; it holds Board, Board.last."
        );
        assert_eq!(
            m.view("plane", "Board", Some("A:A")).unwrap_err().id,
            "view-window-not-a-range"
        );
        let e = m.view("plane", "Board", Some("A1:XFD1048576")).unwrap_err();
        assert_eq!(e.id, "view-window-too-large");
        // The twins are viewed by name; the load filled this one with the
        // grid's values, the one value cell, and no formula has one yet.
        let twin = grid(&m, "board.LAST", None);
        assert!(twin.contains("(window \"Board.last\" \"D1\")"), "{twin}");
        assert!(twin.contains("(cell \"Board.last\" \"D1\" 5)"), "{twin}");
    }

    #[test]
    fn load_is_the_inverse_of_view() {
        // A grid holding every row the record prints.
        let text = "\
(sheet \"Model\" visible)
(sheet \"Notes\" hidden)
(name \"Rate\" \"0.25\")
(gridlines \"Model\" off)
(column \"Model\" \"A\" none hidden none)
(column \"Model\" \"B\" 72 shown 1)
(row \"Model\" 2 30 shown)
(row \"Model\" 4 none hidden)
(format 0 none general nowrap)
(format 1 \"F7F4FC\" text nowrap)
(format 2 \"FF0000\" general wrap)
(look \"Model\" 1 2)
(look \"Model\" \"hot\" 2)
(cell \"Model\" \"A1\" \"say \\\"hi\\\"\")
(style \"Model\" \"A1\" 1)
(sentence \"Model\" \"A1\" 3)
(cell \"Model\" \"B1\" -2.5)
(cell \"Model\" \"C1\" true)
(cell \"Model\" \"D1\" (error \"#N/A\"))
(cell \"Model\" \"A2:B3\" 1)
(formula \"Model\" \"C2:C4\" \"=A2*Rate+Model.last!C2\")
(formula \"Model\" \"D2\" \"=IF(C2>0,\\\"up\\\",\\\"down\\\")\")
(cell \"Notes\" \"A1\" \"the notes\")
";
        let mut m = load(text);
        for _ in 0..2 {
            let saved = save(&m);
            let again = load(&saved);
            assert_eq!(save(&again), saved, "row for row");
            m.step(0);
        }
        let saved = save(&m);
        assert!(saved.contains("(look \"Model\" \"hot\" 2)"), "{saved}");
        assert!(saved.contains("(row \"Model\" 4 none hidden)"), "{saved}");
        assert!(saved.contains("(sheet \"Notes\" hidden)"), "{saved}");
        assert!(saved.contains("(value \"Model\" \"C3\" 0.5)"), "{saved}");
    }

    #[test]
    fn load_is_the_inverse_of_view_on_the_view_goldens() {
        // The goldens of oracle 11 whose window holds the sheet's extent;
        // the sixth, a window clipped inside its sheet, holds the window's
        // cells alone, so its extent row is not the loaded cells'.
        let goldens: &[(&str, &str, &str)] = &[
            (
                include_str!("../../../scripts/view/fixture_frazaro.vla"),
                "Frazaro",
                "B1:C11",
            ),
            (
                include_str!("../../../scripts/view/fixture_output.vla"),
                "Output",
                "A1:F20",
            ),
            (
                include_str!("../../../scripts/view/fixture_data.vla"),
                "data",
                "A1",
            ),
            (
                include_str!("../../../scripts/view/into_output.vla"),
                "Output",
                "A1:B1",
            ),
            (
                include_str!("../../../scripts/view/into_checks.vla"),
                "checks",
                "A1",
            ),
        ];
        for (golden, sheet, window) in goldens {
            let golden = golden.replace("\r\n", "\n");
            let forms = read_forms(&golden).unwrap();
            let draft = read_rows(&forms).unwrap_or_else(|e| panic!("{}", e.refusal.text));
            let w = Window {
                sheet: sheet.to_string(),
                range: parse_a1_range(window).unwrap(),
            };
            let again = view_text_valued(&draft.workbook, &w, &Held(draft.values.clone()));
            assert_eq!(again, golden, "{sheet} {window}");
        }
    }

    /// Life on a 64 by 41 Screen, its interior `B2:BK40` one shared formula:
    /// the first frame a soup from `ROW()` and `COLUMN()`, every later one
    /// Conway's rule over `Screen.last`; the Clock's frame cell written
    /// before each step, as an engine writes it.
    const LIFE: &str = "\
(sheet \"Screen\" visible)
(cell \"Clock\" \"A1\" \"frame\")
(cell \"Clock\" \"B1\" 0)
(formula \"Screen\" \"B2:BK40\" \"=IF(Clock!$B$1=1,IF(MOD(ROW()*7+COLUMN()*13+ROW()*COLUMN(),5)<2,1,0),IF(OR(SUM(Screen.last!A1:C3)-Screen.last!B2=3,AND(Screen.last!B2=1,SUM(Screen.last!A1:C3)-Screen.last!B2=2)),1,0))\")
";
    const W: usize = 64;
    const H: usize = 41;

    /// The same Life written here, every cell of every generation.
    fn life_here(frames: u64) -> Vec<u8> {
        let mut g = vec![0u8; W * H];
        for frame in 1..=frames {
            let mut next = vec![0u8; W * H];
            for r in 2..H {
                for c in 2..W {
                    let (row, col) = (r as u64, c as u64);
                    next[(r - 1) * W + (c - 1)] = if frame == 1 {
                        u8::from((row * 7 + col * 13 + row * col) % 5 < 2)
                    } else {
                        let mut n = 0;
                        for dr in 0..3 {
                            for dc in 0..3 {
                                if dr != 1 || dc != 1 {
                                    n += g[(r - 2 + dr) * W + (c - 2 + dc)];
                                }
                            }
                        }
                        let alive = g[(r - 1) * W + (c - 1)] == 1;
                        u8::from(n == 3 || (alive && n == 2))
                    };
                }
            }
            g = next;
        }
        g
    }

    fn life_step(m: &mut Machine) {
        let frame = m.frame() + 1;
        assert_eq!(m.write(&format!("(cell \"Clock\" \"B1\" {frame})")), Ok(1));
        assert!(m.step(0).done);
    }

    #[test]
    fn life_on_a_small_board_is_the_life_written_here() {
        let mut m = load(LIFE);
        assert_eq!(m.formulas(), (39 * 62, 1), "one shape for the whole rule");
        let before = formula_rows(&m);
        for frame in 1..=12u64 {
            life_step(&mut m);
            assert_eq!(
                plane(&m, "Screen", "A1:BL41"),
                life_here(frame),
                "frame {frame}"
            );
        }
        assert_eq!(formula_rows(&m), before, "AD-7 over twelve frames");
        assert!(life_here(12).contains(&1), "the board is not dead");
    }

    #[test]
    fn a_save_resumes() {
        let mut m = load(LIFE);
        for _ in 0..5 {
            life_step(&mut m);
        }
        let saved = save(&m);
        let mut again = load(&saved);
        assert_eq!(
            again.formulas(),
            (39 * 62, 1),
            "a save's 2,418 rows are one shape"
        );
        again.resume_at(m.frame()).unwrap();
        life_step(&mut m);
        life_step(&mut again);
        assert_eq!(again.frame(), 6);
        assert_eq!(save(&again), save(&m), "cell for cell");
        assert_eq!(plane(&again, "Screen", "A1:BL41"), life_here(6));
    }

    #[test]
    fn a_shape_at_an_offset_is_the_parse_of_its_moved_text() {
        // Every case of the refers fixture that parses, moved by a set of
        // distances to every edge of the sheet: the rectangle evaluated at
        // the offset is the rectangle the moved text parses to, or #REF!
        // where the mover writes #REF!; and where the moved text does not
        // parse (a 3D span moved off the sheet, `Jan:Dec!#REF!`), the shape
        // names the same construct at that offset.
        use crate::calc::formula::{parse, Expr};
        fn same(a: &Expr, b: &Expr, d: (i64, i64)) -> bool {
            match (a, b) {
                (Expr::Ref(x), Expr::Ref(y)) => {
                    x.sheet == y.sheet
                        && x.last_sheet == y.last_sheet
                        && x.moved(d.0, d.1) == Some((y.top, y.left, y.bottom, y.right))
                }
                (Expr::Ref(x), Expr::Error(crate::calc::ErrorKind::Ref)) => {
                    x.moved(d.0, d.1).is_none()
                }
                (Expr::Call(n, xs), Expr::Call(m, ys)) => {
                    n == m && xs.len() == ys.len() && xs.iter().zip(ys).all(|(x, y)| same(x, y, d))
                }
                (Expr::Binary(o, x1, x2), Expr::Binary(p, y1, y2)) => {
                    o == p && same(x1, y1, d) && same(x2, y2, d)
                }
                (Expr::Neg(x), Expr::Neg(y))
                | (Expr::Pos(x), Expr::Pos(y))
                | (Expr::Percent(x), Expr::Percent(y)) => same(x, y, d),
                _ => a == b,
            }
        }
        let fixture = include_str!("../../../scripts/refers.txt").replace("\r\n", "\n");
        let mut formulas: Vec<String> = Vec::new();
        let lines: Vec<&str> = fixture.lines().collect();
        for (i, line) in lines.iter().enumerate() {
            if line.starts_with("=== ") && i + 2 < lines.len() {
                formulas.push(lines[i + 2].to_string());
            }
        }
        formulas.push(
            LIFE.lines()
                .nth(3)
                .unwrap()
                .split('"')
                .nth(5)
                .unwrap()
                .replace("\\\"", "\""),
        );
        let offsets = [
            (1, 0),
            (0, 1),
            (2, 3),
            (-1, 0),
            (0, -1),
            (-3, -2),
            (1_048_575, 0),
            (0, 16_383),
            (-1_048_575, 0),
            (0, -16_383),
        ];
        let mut compared = 0;
        let mut unread = 0;
        for text in &formulas {
            let Ok(e) = parse(text) else { continue };
            let mut shapes = Shapes::new();
            let mut book = Workbook::new();
            let s = book.ensure_sheet("Model");
            book.sheets[s].set(
                1,
                1,
                Cell {
                    content: Content::Formula(text.trim_start_matches('=').to_string()),
                    style: 0,
                },
            );
            shapes.replace(&book, (0, 1, 1), &library::language());
            let shape = shapes.shape(0).unwrap();
            for d in offsets {
                let moved = shift_a1_references(text, d.0, d.1);
                match parse(&moved) {
                    Ok(pm) => {
                        assert_eq!(shape.unread_at(d.0, d.1), None, "{moved}");
                        assert!(same(&e, &pm, d), "{text} moved by {d:?} is {moved}");
                        compared += 1;
                    }
                    Err(construct) => {
                        assert_eq!(
                            shape.unread_at(d.0, d.1),
                            Some(construct),
                            "{text} moved by {d:?} is {moved}"
                        );
                        unread += 1;
                    }
                }
            }
        }
        assert!(
            compared >= 500,
            "{compared} formulas and distances compared"
        );
        assert!(unread > 0, "a 3D span moved off the sheet was reached");
    }

    #[test]
    fn plain_formulas_of_one_r1c1_text_are_one_shape() {
        let m = load("(formula \"S\" \"A2\" \"=A1+1\")\n(formula \"S\" \"A3\" \"=A2+1\")\n(formula \"S\" \"B3\" \"=B2+1\")\n(formula \"S\" \"C3\" \"=$A$1+1\")\n");
        assert_eq!(m.formulas(), (4, 2));
        let book = m.workbook();
        let names = Names::of(book);
        let extents: Vec<Extent> = book.sheets.iter().map(Sheet::extent).collect();
        let s = Shapes::of(book, 1, &library::language(), &names, &extents);
        assert_eq!(
            s.order().first(),
            Some(&(0, 2, 1)),
            "A2 before what reads it"
        );
        assert!(s.cycles().is_empty());
    }
}
