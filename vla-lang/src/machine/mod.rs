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
//! of the inputs it began with. A run of the order, cells of one shape that
//! read none of one another, is evaluated at once, node by node, as each of
//! its cells alone would be (`crate::calc::vector`, KERNEL.20); a cell
//! alone, and a run's few cells, by its shape's parse lowered once to a
//! program, as `Ctx` would evaluate it, and by `Ctx` where the program's
//! guards fail (`crate::calc::program`, KERNEL.44).
//!
//! A frame computes only the cells whose reads changed, and every other
//! keeps its value (KERNEL.20). A change is followed from the cell that
//! changed to the cells that read it (`crate::calc::reach`): from a twin's
//! cell the last frame's end changed, from a cell written since, and from a
//! cell the frame itself computed to a value whose bits differ from its last.
//! The first frame computes every cell, and so does a frame after a write
//! that changed a formula or added a cell, or after a twin gained or lost a
//! cell. The budget counts the cells of the order passed, computed or not.
//!
//! The frame's values are kept once, each overwritten as its cell computes
//! a value that differs; a step that yields keeps each value it overwrote
//! beside, so a view during a yield shows the last complete frame. A sheet
//! of the rows that the plan reads by a range or by a run's moving cell keeps
//! the frame's values as an array too, each slot written where its cell's
//! value is kept, a value cell's at the write that changes it (KERNEL.31): so
//! that a range of the frame's own cells is folded and walked by index, and
//! a run's cells read one of them by slices, as a twin's array is read; the
//! order puts each reader after every cell it reads, so the slots it reads
//! are this frame's. A sheet one of whose formulas holds a reason offers no
//! array until none does, its ranges and cells read through its cells, which
//! answer the reason. When a
//! frame completes, the values that changed are copied into the twins, with
//! the planes and the twins' arrays, however many changed (KERNEL.36); when
//! the cells a copy would write through a twin's own cells, on a sheet too
//! sparse for an array, are more than half the cells, or a twin's cells
//! change which hold a value, every sheet is walked instead.
//! The walk copies every sheet's values into its twin, and keeps in the same
//! walk the plane of each sheet an engine views as one, so that a view of it
//! is a copy ([`plane::Raster`], KERNEL.24), and each twin's values as an
//! array over its sheet's extent, so that the next frame's read of a twin is
//! an index and its range a walk of the array ([`Dense`], KERNEL.20). A
//! frame's end that copies writes a changed formula cell of a twin kept as
//! an array into the array alone, and leaves its twin cell behind its slot,
//! made from it only when a reader of the twin as a sheet asks (KERNEL.27):
//! a view reads such a cell through its slot ([`Values::content`]), the walk
//! makes it where it stands as it passes it (KERNEL.36), and
//! [`Machine::workbook`] makes every such cell first, so that a twin is
//! still an ordinary sheet to every reader. The step computes values
//! and never formulas (`AD-7`): the formula cells after a step are the
//! formula cells before it.
//!
//! **The changes** (KERNEL.42). The twin cells a frame's end changed are
//! kept until the next frame's end, and with them the few whose change a
//! record spells as before, so that a view of a window's changes answers
//! from them alone, the cells whose rows in the window's values the last
//! complete frame changed, at the cost of the changes and not the window;
//! or answers whole where they cannot say: before the first frame, after a
//! frame whose end kept more changes than it held room for or made a twin
//! again, after a write that changed a formula or added a cell, and while a
//! value written since lies in the window.
//!
//! **The count** (KERNEL.48). A frame's end keeps how many formula cells
//! the frame changed the value of, to the bit as the frame judges a change,
//! how many of them stand on each sheet, and how many formula cells it
//! computed ([`Machine::changed`], [`Machine::changed_on`],
//! [`Machine::reached`]), so that a host stepping a grid to a fixed point
//! asks after each frame whether anything still moves. A frame that changed
//! nothing, with no write before it, is a fixed point: the twins at its end
//! are the twins it read, so the next frame with no write before it
//! computes what it computed.
//!
//! **The writes.** `cell`, `formula` and `derived` rows, applied in order;
//! a refused row undoes the rows before it, so a refused write changes
//! nothing. A `derived` row is the host's carriage of the Write sheet: a
//! value that lands only in a value cell, never over a formula. A write
//! costs what its cells cost, a fill rewritten whole in a frame's time
//! rather than the square of its cells, and leaves the plan a load of the
//! same grid makes, numbered and anchored as a load's (KERNEL.34).
//!
//! **The refusals** are the language's catalogue's: the `grid` family for
//! the machine's own, `view-*` for the view, `calc-*` for what stops a grid
//! being stepped. An engine adds its own (its manifest, its devices) around
//! these calls and never inside them.

pub mod load;
pub mod plane;

use std::cell::Cell as Slot;
use std::collections::{btree_map::Entry, BTreeMap, BTreeSet, HashMap};

use crate::calc::dense;
use crate::calc::graph::{is_formula, Extent, Names};
use crate::calc::program::{Answer, Source};
use crate::calc::reach::{Reach, Reached};
use crate::calc::shape::{spell_cell, Shapes, RUN, SHORT};
use crate::calc::vector::{self, Lanes};
use crate::calc::{CellId, Computed, Ctx, Dense, Env, ErrorKind, Library, Reason, Value, Values};
use crate::form::Form;
use crate::fx::FxHashMap;
use crate::messages::{raise, Refusal};
use crate::projection::Window;
use crate::reader::read_forms;
use crate::refers::shift_a1_references;
use crate::rows::Visibility;
use crate::sheet::{
    filled, formula_text, parse_a1_range, shared_masters, strip_future_prefixes, A1Range, Cell,
    Content, Sheet, Workbook,
};
use crate::view::{changes_text, view_text_valued, view_text_values, window_over};

pub use load::{read_rows, sheet_name_ok, thousands, twin_base, Draft};

/// The most cells a grid holds beside its twins (KERNEL.22): a stated limit,
/// sixteen Screens of 320 by 200, so that one range row cannot ask for the
/// seventeen billion cells of Excel's sheet.
pub const CELLS: usize = 1_048_576;

/// The most grids live at once in one table (`Alonzo/SPEC.md` section 10).
pub const HANDLES: usize = 16;

/// The rows a write takes, as `grid-row-unknown` lists them.
pub const WRITE_ROWS: &str = "cell, formula and derived";

/// The value cells a grid keeps room for as written between two frames
/// (KERNEL.20), so that an engine's write of its Clock at every step makes
/// nothing; more are kept all the same.
const WRITTEN: usize = 64;

/// The changes a frame's end keeps of those a record spells as it spelled
/// them before (KERNEL.42): a formula's value turned from 0 to -0, which
/// its `value` row spells 0 both times; a number no cell can hold turned to
/// another, or to the error `#NUM!`, each spelled `#NUM!`; a formula's
/// reason turned to another, which its twin, holding no cell, never shows.
/// Rare, so few: past them, the frame's changes are whole.
const ALIKE: usize = 64;

/// A change the sheet's own record spells as before, and one its twin's
/// does (KERNEL.42).
const SHEET_ALIKE: u8 = 1;
const TWIN_ALIKE: u8 = 2;

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
    /// The frame's values, each formula cell's at its position in the order
    /// (KERNEL.26), `None` for one with no value yet: the last complete
    /// frame's, each overwritten while a frame is in progress as its cell
    /// computes a value that differs (KERNEL.20), so that a cell's answer is
    /// kept and read again by index, and a cell read by its place found in
    /// the plan's own array of positions (`Reach::position`).
    values: Vec<Option<Computed>>,
    /// Each value a frame in progress overwrote while a step yielded, by
    /// position, as the last complete frame left it (`None`, no value), so
    /// that a view before the frame completes reads the last complete frame.
    /// This map and the next two are under the crate's own hash (KERNEL.25);
    /// nothing walks them, so their order is never read.
    before: FxHashMap<u32, Option<Computed>>,
    /// A twin cell left empty because its cell was not computed: the
    /// reason, which a formula reading the twin cell inherits.
    twin_reasons: FxHashMap<CellId, Reason>,
    /// The reasons as the frame before left them, kept for the walk to tell
    /// which changed.
    reasons_before: FxHashMap<CellId, Reason>,
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
    /// By rows' sheet, its twin's values as an array over the sheet's
    /// extent (KERNEL.20), kept with the twin at every frame's end and
    /// changed by nothing else, a twin being refused to every write; `None`
    /// for a sheet too sparse for one, whose twin is read through its cells.
    dense: Vec<Option<Dense>>,
    /// By rows' sheet kept as an array, a bit a slot for each of its twin's
    /// cells whose content the slot holds newer (KERNEL.27): a frame's end
    /// writes a changed formula cell's slot, byte and reason, and the twin
    /// cell is made from its slot only when a reader of the twin as a sheet
    /// asks, a view reading through the slot ([`Values::content`]), the walk
    /// and [`Machine::workbook`] making it. Every bit is clear when no array
    /// is kept.
    behind: Vec<Vec<u64>>,
    /// By rows' sheet, whether any of its twin's cells is behind.
    lagging: Vec<bool>,
    /// By rows' sheet, the frame's own values as an array over the sheet's
    /// extent, its numbers beside (KERNEL.31): for a sheet the plan reads by
    /// a range or by a run's moving cell (`Shapes::read_as_array`) whose
    /// cells stand close enough together for an array, as a twin's do. Each
    /// slot holds what a read of its cell answers now, a formula's value of
    /// this frame as it is kept, a value cell's content as it was written,
    /// nothing for a cell that holds none or a formula with no value, beside
    /// which a reason may stand; made at a load and where a write moves the
    /// sheet's extent or which sheets the plan reads so, and written where a
    /// value changes. `None` for every other sheet.
    own: Vec<Option<Dense>>,
    /// By rows' sheet kept as an array, how many of its formula cells hold
    /// a reason: its array is offered to the evaluator only while none does,
    /// its cells answering a reason where a slot holds nothing.
    own_reasons: Vec<u32>,
    /// Whether any of the frame's own sheets keeps an array, asked first, so
    /// that a plan that keeps none asks of a cell's read and change nothing
    /// it asked before (KERNEL.31).
    owns: bool,
    /// The buffers a run evaluated at once holds, and its cells' answers,
    /// kept from one run to the next so that a frame makes them once
    /// (KERNEL.20).
    lanes: Lanes,
    answers: Vec<Computed>,
    /// A run's answers when every one is a number, as its lane holds them
    /// (KERNEL.43), compared with the values held to the bit.
    answered: Vec<f64>,
    /// The cells of a run a frame computes, gathered a chunk at a time, and
    /// their positions in the order.
    gathered: Vec<CellId>,
    gathered_at: Vec<u32>,
    /// Whether the next frame computes every cell (KERNEL.20): after a load,
    /// a write that changed a formula or added a cell, a twin that gained or
    /// lost a cell, or more changes than were kept.
    every_next: bool,
    /// Whether the frame in progress computes every cell.
    every: bool,
    /// By position in the order, whether the frame in progress computes
    /// the cell there, a bit a cell.
    marked: Vec<u64>,
    /// By shape, whether the frame in progress computes every cell of it,
    /// and whether its parts the same at every cell are computed again at
    /// the frame's start.
    whole: Vec<bool>,
    recheck: Vec<bool>,
    /// Each such part's value as a frame's start last computed it.
    parts: Vec<Option<Result<Value, Reason>>>,
    /// The formula cells whose values the frame in progress changed, by
    /// position in the order, with room for every formula cell (KERNEL.36),
    /// and whether more changed than this holds.
    changed: Vec<u32>,
    changed_over: bool,
    /// What the last complete frame's end kept of the frame (KERNEL.48):
    /// how many formula cells it changed the value of, how many of those
    /// stand on each rows' sheet, and how many formula cells it computed,
    /// counted in `reaching` while it was in progress.
    last_changed: usize,
    changed_sheets: Vec<u32>,
    last_reached: usize,
    reaching: usize,
    /// The value cells written since the last frame completed.
    written: Vec<CellId>,
    /// The twin cells the last frame's end changed, which the next frame
    /// follows to the cells that read them, kept until its end for a view
    /// of the changes (KERNEL.42).
    twin_changed: Vec<CellId>,
    /// Those of them whose change a record spells as before, with the
    /// records that do ([`SHEET_ALIKE`], [`TWIN_ALIKE`]), at most
    /// [`ALIKE`]; and whether there were more (KERNEL.42).
    alike: Vec<(CellId, u8)>,
    alike_over: bool,
    /// Whether the last complete frame's changes are more than the lists
    /// tell (KERNEL.42): before the first frame, or when its end kept more
    /// changes than it held room for, or more spelled alike than
    /// [`ALIKE`], or made a twin again, or a write before it changed a
    /// formula or added a cell, which may change a record's rows where no
    /// twin cell changes.
    changes_whole: bool,
    /// Whether a write since the last frame's end changed a formula or
    /// added a cell (KERNEL.42).
    reshaped: bool,
    /// Whether a run of the order is evaluated at once: always, but in the
    /// tests that hold it to the same cells evaluated one by one.
    by_runs: bool,
    /// Whether a cell evaluated alone, and a run's few cells, run their
    /// shape's program (KERNEL.44): always, but in the tests that hold the
    /// programs to `Ctx` and a run's few cells to the run.
    lowered: bool,
    /// Whether a frame computes only what changed: always, but in the tests
    /// that hold it to every cell computed and every sheet walked.
    incremental: bool,
    /// Whether a frame's end leaves a twin's cells behind their slots, made
    /// when read (KERNEL.27): always, but in the tests that hold it to the
    /// cells copied at every frame's end.
    lazy: bool,
    /// Whether a frame's end copies the cells it changed where it may:
    /// always, but in the tests that hold the copy to the walk, whose every
    /// frame's end walks (KERNEL.36).
    copies: bool,
    /// Whether a frame's own sheets the plan reads so keep their values as
    /// arrays (KERNEL.31): always, but in the tests that hold the arrays to
    /// the same sheets read through their cells.
    arrays: bool,
    /// The cells the frame in progress computed, for the tests to count.
    #[cfg(test)]
    computed: usize,
    /// The cells programs answered since the machine was made, for the tests
    /// to count (KERNEL.44).
    #[cfg(test)]
    ran: usize,
    /// How often a frame's own sheet kept as an array was read through its
    /// cells instead, a formula of it holding a reason, for the tests to
    /// count (KERNEL.31).
    #[cfg(test)]
    aside: Slot<usize>,
    /// How often a write kept a frame's own sheet's array over its extent,
    /// and made one again, for the tests to count (KERNEL.31).
    #[cfg(test)]
    kept_over: usize,
    #[cfg(test)]
    made_again: usize,
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
    /// The cells of a fill a release gave their own text, whose formulas
    /// read as before: placed again when the write settles, their values
    /// kept (KERNEL.34).
    released: BTreeSet<CellId>,
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

    /// The machine over a draft: the twins made, every formula placed on its
    /// shape, its references' sheets found among the grid's and its twins
    /// (KERNEL.20), then refused when the grid cannot be stepped (a cycle, a
    /// function the library does not hold, a construct this version does
    /// not read, in that order); then the twins filled with the grid's
    /// values, each formula's from its `value` row. Nothing is computed: the
    /// frame is 0 and every formula waits for the first step.
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
        for i in 0..sheets {
            let mut twin = Sheet::new(&format!("{}.last", workbook.sheets[i].name));
            twin.visibility = Visibility::Hidden;
            workbook.sheets.push(twin);
            extents.push(None);
        }
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
        Ok(Machine::made(
            (workbook, sheets, cells),
            (library, names, extents),
            shapes,
            values,
        ))
    }

    /// The machine made from its parts, in a frame of its own (KERNEL.20):
    /// the machine is the largest value its load holds, and made in the frame
    /// that parses every formula it would stand on the stack beneath the
    /// deepest parse, which `VLA.1`'s limits pin on a stack of 1 MiB.
    #[inline(never)]
    fn made(
        (book, sheets, cells): (Workbook, usize, usize),
        (library, names, extents): (Library, Names, Vec<Extent>),
        shapes: Shapes,
        loaded: HashMap<CellId, Computed>,
    ) -> Machine {
        // Each formula's value from its `value` row, at its position.
        let mut values: Vec<Option<Computed>> = vec![None; shapes.order().len()];
        for (id, value) in loaded {
            if let Some(p) = shapes.reach().position(id) {
                values[p] = Some(value);
            }
        }
        let mut m = Machine {
            book,
            sheets,
            library,
            names,
            extents,
            shapes,
            values,
            before: FxHashMap::default(),
            twin_reasons: FxHashMap::default(),
            reasons_before: FxHashMap::default(),
            frame: 0,
            progress: None,
            cells,
            here: Slot::new((0, 1, 1)),
            wanted: (0..sheets).map(|_| Slot::new(false)).collect(),
            planes: vec![None; sheets],
            dense: vec![None; sheets],
            behind: vec![Vec::new(); sheets],
            lagging: vec![false; sheets],
            own: vec![None; sheets],
            own_reasons: vec![0; sheets],
            owns: false,
            lanes: Lanes::default(),
            answers: Vec::new(),
            answered: Vec::with_capacity(vector::CHUNK),
            gathered: Vec::with_capacity(vector::CHUNK),
            gathered_at: Vec::with_capacity(vector::CHUNK),
            every_next: true,
            every: true,
            marked: Vec::new(),
            whole: Vec::new(),
            recheck: Vec::new(),
            parts: Vec::new(),
            changed: Vec::new(),
            changed_over: false,
            last_changed: 0,
            changed_sheets: vec![0; sheets],
            last_reached: 0,
            reaching: 0,
            written: Vec::with_capacity(WRITTEN),
            twin_changed: Vec::new(),
            alike: Vec::with_capacity(ALIKE),
            alike_over: false,
            changes_whole: true,
            reshaped: false,
            by_runs: true,
            lowered: true,
            incremental: true,
            lazy: true,
            copies: true,
            arrays: true,
            #[cfg(test)]
            computed: 0,
            #[cfg(test)]
            ran: 0,
            #[cfg(test)]
            aside: Slot::new(0),
            #[cfg(test)]
            kept_over: 0,
            #[cfg(test)]
            made_again: 0,
        };
        m.fit();
        m.walk();
        m.own_made(None);
        m
    }

    /// The frame's own sheets kept as arrays as the plan reads them
    /// (KERNEL.31): at a load, `written` none, and at a write that plans
    /// again or adds a cell, `written` the cells it changed. A sheet the plan
    /// reads by a range or a run's moving cell whose cells stand close enough
    /// together keeps one; one it kept over the same extent stands, its slots
    /// being by place, which a plan does not move, the written cells' slots
    /// alone written again, its reasons counted by the write (`settle`); any
    /// other is made from the frame's values and the sheet's cells, its
    /// reasons counted; every other sheet keeps none.
    fn own_made(&mut self, written: Option<&[CellId]>) {
        for s in 0..self.sheets {
            if !self.arrays || !self.shapes.read_as_array(s) {
                self.own[s] = None;
                self.own_reasons[s] = 0;
                continue;
            }
            if let (Some(written), Some(kept)) = (written, self.own[s].as_ref()) {
                if kept.stands_over(self.extents[s]) {
                    for &id in written.iter().filter(|id| id.0 == s) {
                        let v = self.slot_value(id);
                        if let Some(Some(kept)) = self.own.get_mut(s) {
                            if let Some(i) = kept.index(id.1, id.2) {
                                kept.write_at(i, |slot| *slot = v);
                            }
                        }
                    }
                    #[cfg(test)]
                    {
                        self.kept_over += 1;
                    }
                    continue;
                }
            }
            #[cfg(test)]
            if written.is_some() {
                self.made_again += 1;
            }
            // Made whole, in an old store's memory, its extent having moved
            // (`Dense::over` empties every slot of it first); none for a
            // sheet whose cells stand too far apart for one, as a twin's.
            let limit = dense::limit(self.book.sheets[s].cells.len(), CELLS as u64);
            let old = self.own[s].take();
            let Some(mut d) = Dense::over(self.extents[s], limit, old, true) else {
                self.own_reasons[s] = 0;
                continue;
            };
            let mut reasons = 0;
            for (&(row, col), cell) in &self.book.sheets[s].cells {
                let v = if is_formula(&cell.content) {
                    match value_at(&self.shapes, &self.values, (s, row, col)) {
                        Some(Computed::Value(v)) => v.clone(),
                        Some(Computed::NotComputed(_)) => {
                            reasons += 1;
                            Value::Empty
                        }
                        None => Value::Empty,
                    }
                } else {
                    content_value(&cell.content)
                };
                if let Some(i) = d.index(row, col) {
                    d.write_at(i, |slot| *slot = v);
                }
            }
            self.own[s] = Some(d);
            self.own_reasons[s] = reasons;
        }
        self.owns = self.own.iter().any(Option::is_some);
    }

    /// What a read of a rows' cell answers now, as an array's slot holds it
    /// (KERNEL.31): a formula's value of the frame, nothing for a formula
    /// with no value or a reason, which a read finds beside; a value cell's
    /// content as a read answers it; nothing for a cell holding none.
    fn slot_value(&self, id: CellId) -> Value {
        match self.book.sheets[id.0].cells.get(&(id.1, id.2)) {
            Some(cell) if is_formula(&cell.content) => {
                match value_at(&self.shapes, &self.values, id) {
                    Some(Computed::Value(v)) => v.clone(),
                    _ => Value::Empty,
                }
            }
            Some(cell) => content_value(&cell.content),
            None => Value::Empty,
        }
    }

    /// The value a formula cell's frame changed to, at a position of the
    /// order, into its sheet's array where one is kept (KERNEL.31), and its
    /// sheet's reasons counted, `was` whether it held a reason before.
    #[inline]
    fn own_put(&mut self, p: usize, was: bool) {
        if !self.owns {
            return;
        }
        let id = self.shapes.order()[p];
        let Some(Some(d)) = self.own.get_mut(id.0) else {
            return;
        };
        let Some(i) = d.index(id.1, id.2) else {
            return;
        };
        let now = &self.values[p];
        let reason = matches!(now, Some(Computed::NotComputed(_)));
        d.write_at(i, |slot| match now {
            Some(Computed::Value(Value::Number(n))) => *slot = Value::Number(*n),
            Some(Computed::Value(v)) => *slot = v.clone(),
            _ => *slot = Value::Empty,
        });
        if reason != was {
            let n = &mut self.own_reasons[id.0];
            if reason {
                *n += 1;
            } else {
                *n -= 1;
            }
        }
    }

    /// What follows a frame's changes made to the plan's size, once a plan
    /// is made, so that a frame makes none (KERNEL.20): a bit a cell of the
    /// order, a flag a shape, a value a part, room for every formula cell
    /// among the changes, since a cell is kept as changed at most once a
    /// frame (KERNEL.36: four bytes a formula cell), and room for the twin
    /// cells' changes of half the cells, beyond which the next frame computes
    /// every cell.
    fn fit(&mut self) {
        let of = self.shapes.order().len();
        self.marked.clear();
        self.marked.resize(of.div_ceil(64), 0);
        let shapes = self.shapes.shapes();
        self.whole.clear();
        self.whole.resize(shapes, false);
        self.recheck.clear();
        self.recheck.resize(shapes, false);
        self.parts.clear();
        self.parts.resize(self.shapes.reach().part_count(), None);
        fn room<T>(v: &mut Vec<T>, n: usize) {
            v.reserve(n.saturating_sub(v.len()));
        }
        room(&mut self.changed, of);
        room(&mut self.twin_changed, self.cells / 2 + 64);
    }

    /// What the load made, for the load's row.
    pub fn loaded(&self) -> Loaded {
        Loaded {
            sheets: self.sheets,
            cells: self.cells,
            formulas: self.shapes.formulas(),
        }
    }

    /// The grid, its twins included, each twin's cells made from its array
    /// first where a frame's end left them behind it (KERNEL.27), which is
    /// why the grid is lent only to a caller that may change the machine.
    /// A sheet's name and place are [`Machine::find_sheet`]'s and
    /// [`Machine::sheet_name`]'s, which never need the twins made.
    pub fn workbook(&mut self) -> &Workbook {
        self.make_twins();
        &self.book
    }

    /// A sheet's index by name, compared as Excel compares names, a twin's
    /// included, as [`Workbook::find_sheet`] finds it.
    pub fn find_sheet(&self, name: &str) -> Option<usize> {
        self.book.find_sheet(name)
    }

    /// A sheet's name by index, a twin's included.
    pub fn sheet_name(&self, index: usize) -> Option<&str> {
        self.book.sheets.get(index).map(|s| s.name.as_str())
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

    /// How many formula cells the last complete frame changed the value of
    /// (KERNEL.48): each whose value's bits differ from those it held when
    /// the frame began, as the frame judges a change it keeps, so that a
    /// first value, a 0 turned to -0 and a reason turned to another count,
    /// and a value cell, written or not, never does. 0 before the first
    /// frame since the load; while a frame is in progress, the last complete
    /// one's; a write leaves it as it is.
    ///
    /// A frame that changed nothing, with no write before it, is a fixed
    /// point: the twins at its end are the twins it read, so the next frame
    /// with no write before it computes what it computed and changes
    /// nothing. The frame right after a write is never that proof, since its
    /// end copies the written cells into their twins, which a formula reading
    /// a written cell's twin sees a frame later. So a host stepping a grid to
    /// a fixed point steps again while the last frame changed a formula's
    /// value or had a write before it, a write of its own included.
    pub fn changed(&self) -> usize {
        self.last_changed
    }

    /// Of the cells [`Machine::changed`] counts, how many stand on a sheet,
    /// by its index as [`Machine::sheet_name`] takes one (KERNEL.48): 0 for a
    /// twin, which holds no formula, and `None` past the grid's sheets.
    pub fn changed_on(&self, sheet: usize) -> Option<usize> {
        if sheet >= self.book.sheets.len() {
            return None;
        }
        Some(self.changed_sheets.get(sheet).map_or(0, |&n| n as usize))
    }

    /// How many formula cells the last complete frame computed (KERNEL.48):
    /// every one at a frame that computes every cell, else those its changes
    /// reached, the shapes it computed whole among them. 0 before the first
    /// frame since the load; while a frame is in progress, the last complete
    /// one's.
    pub fn reached(&self) -> usize {
        self.last_reached
    }

    /// The machine's plan printed as rows, its census (KERNEL.30): what each
    /// shape costs a frame, read before a frame runs, as
    /// [`crate::calc::Shapes::census`] prints a plan, but each read of a
    /// twin named by how the machine keeps that twin as its last frame's end
    /// left it, an array read by index or cells read through lookups; and a
    /// frame's own sheet as the plan keeps it (KERNEL.31), as a door's
    /// census names it, a sheet stepping aside for a reason being a frame's
    /// and never the plan's. Computed when asked: a load, a step and a write
    /// keep nothing for it but where each level of the order ends and which
    /// sheets a read takes as an array.
    pub fn census(&self) -> String {
        crate::calc::census::rows(&self.shapes, &self.book, &self.library, &|s| {
            if s < self.sheets {
                self.own[s].is_some()
            } else {
                Env::dense(self, s).is_some()
            }
        })
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

    /// One step: at most `budget` cells of the order passed, 0 meaning every
    /// one, each computed when a cell it reads changed (KERNEL.20); a frame
    /// not finished is held, and the next step continues it; a frame
    /// finished fills the twins. A run of the order, cells of one shape
    /// reading none of one another, is evaluated at once, up to
    /// [`vector::CHUNK`] cells at a time (KERNEL.20); every other cell alone.
    pub fn step(&mut self, budget: usize) -> Step {
        let of = self.shapes.order().len();
        if self.progress.is_none() {
            self.begin();
        }
        let start = self.progress.unwrap_or(0);
        let end = if budget == 0 {
            of
        } else {
            start.saturating_add(budget).min(of)
        };
        // A step that stops before the frame's end keeps each value it
        // overwrites beside, for a view before the frame completes.
        let keep = end < of;
        // The runs walked in order beside the positions, the next found once
        // where a search found the run of each position (KERNEL.44): a grid
        // of many cells alone between its runs searched them at every one.
        let mut i = start;
        let mut k = self.shapes.run_after(start);
        while i < end {
            let next = if self.by_runs {
                self.shapes.run_span(k)
            } else {
                None
            };
            match next {
                Some((from, after, shape)) if from <= i => {
                    let to = after.min(end);
                    self.evaluate_run(shape, i, to, keep);
                    i = to;
                    k += 1;
                }
                _ => {
                    let stop = next.map_or(end, |(from, _, _)| from.min(end));
                    while i < stop {
                        if self.wants(i) {
                            self.evaluate_alone(i, keep);
                        }
                        i += 1;
                    }
                }
            }
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
        self.progress = None;
        self.frame += 1;
        self.before.clear();
        self.snapshot();
        Step {
            frame: self.frame,
            evaluated: of,
            of,
            done: true,
        }
    }

    /// A frame's start (KERNEL.20): every cell to compute, or the cells the
    /// last frame's end and the writes since reach, with every shape not
    /// followed, and each shape whose parts the same at every cell read a
    /// changed cell when those parts compute other bits than they did. The
    /// last frame's changes are kept until this frame's end, which a view
    /// of them during a yield reads (KERNEL.42).
    fn begin(&mut self) {
        self.every = self.every_next || !self.incremental;
        self.every_next = false;
        self.reaching = 0;
        #[cfg(test)]
        {
            self.computed = 0;
        }
        if self.every {
            // Every part computed, for the next frame's to be held to.
            for s in 0..self.shapes.shapes() {
                self.parts_moved(s);
            }
            return;
        }
        let Machine {
            shapes,
            marked,
            whole,
            recheck,
            twin_changed,
            written,
            ..
        } = self;
        let reach = shapes.reach();
        for &s in reach.unfollowed() {
            whole[s] = true;
        }
        for id in twin_changed.iter().chain(written.iter()) {
            reach.each(*id, &mut |r| match r {
                Reached::Cell(p) => marked[p / 64] |= 1 << (p % 64),
                Reached::Shape(s) => whole[s] = true,
                Reached::Parts(s) => recheck[s] = true,
            });
        }
        // Every shape whose parts read a changed cell has them computed, even
        // one computed whole already, so that what is kept is what its cells
        // were computed with: a part kept from an earlier frame would let a
        // later change back to it pass unseen.
        for s in 0..self.shapes.shapes() {
            if self.recheck[s] && self.parts_moved(s) {
                self.whole[s] = true;
            }
        }
    }

    /// A shape's parts the same at every cell computed, each kept for the
    /// next frame's start; whether any computed other bits than it did, or
    /// had not been computed.
    fn parts_moved(&mut self, s: usize) -> bool {
        let mut moved = false;
        for k in self.shapes.reach().parts(s) {
            let anchor = self.shapes.shape(s).map_or((0, 1, 1), |shape| shape.anchor);
            self.here.set(anchor);
            let got = Ctx::at(self, &self.library, 0, 0).scalar(self.shapes.reach().part(k));
            moved |= !self.parts[k]
                .as_ref()
                .is_some_and(|was| same_got(was, &got));
            self.parts[k] = Some(got);
        }
        moved
    }

    /// Whether the frame in progress computes the cell at a position.
    fn wants(&self, pos: usize) -> bool {
        self.every
            || self.whole[self.shapes.reach().shape_at(pos)]
            || self.marked[pos / 64] & (1 << (pos % 64)) != 0
    }

    fn evaluate(&self, id: CellId) -> Computed {
        self.here.set(id);
        self.shapes.evaluate(self, &self.library, id)
    }

    /// The cell at a position of the order evaluated alone, and its value
    /// kept: by its shape's program, at the cell's offset from the shape's
    /// first cell, where the shape has one and no guard fails (KERNEL.44);
    /// else by `Ctx`, as every cell alone was before.
    fn evaluate_alone(&mut self, p: usize, keep: bool) {
        self.reaching += 1;
        let id = self.shapes.order()[p];
        if self.lowered {
            let s = self.shapes.reach().shape_at(p);
            if let Some((program, anchor)) = self.shapes.program(s) {
                self.here.set(id);
                let off = (
                    i64::from(id.1) - i64::from(anchor.1),
                    i64::from(id.2) - i64::from(anchor.2),
                );
                // A cell's offset from its shape's first cell is its place's.
                #[cfg(test)]
                assert_eq!(
                    self.shapes.place(id).map(|q| (q.d_row, q.d_col)),
                    Some(off),
                    "{id:?}'s offset"
                );
                if let Some(answer) = program.run(self, &self.library, id, off) {
                    #[cfg(test)]
                    {
                        self.ran += 1;
                    }
                    return match answer {
                        Answer::Number(x) => self.keep_number(p, x, keep),
                        Answer::Truth(b) => self.keep(p, Computed::Value(Value::Bool(b)), keep),
                    };
                }
            }
        }
        let computed = self.evaluate(id);
        self.keep(p, computed, keep);
    }

    /// The cells of the order from `from` to `to`, of one run, that the
    /// frame computes, evaluated at once up to a chunk at a time (KERNEL.20),
    /// fewer than a run's least one by one, and their values kept; fewer
    /// than [`SHORT`] one by one by their shape's program where it has one
    /// (KERNEL.44).
    fn evaluate_run(&mut self, shape: usize, from: usize, to: usize, keep: bool) {
        let mut lanes = std::mem::take(&mut self.lanes);
        let mut answers = std::mem::take(&mut self.answers);
        let mut answered = std::mem::take(&mut self.answered);
        let mut gathered = std::mem::take(&mut self.gathered);
        let mut gathered_at = std::mem::take(&mut self.gathered_at);
        let every = self.every || self.whole[shape];
        let lowered = self.lowered && self.shapes.program(shape).is_some();
        let few = if lowered { SHORT } else { RUN };
        let mut i = from;
        while i < to {
            if every {
                // Every cell, in chunks of the order as it stands.
                let j = (i + vector::CHUNK).min(to);
                if lowered && j - i < SHORT {
                    for p in i..j {
                        self.evaluate_alone(p, keep);
                    }
                    i = j;
                    continue;
                }
                let cells = &self.shapes.order()[i..j];
                self.shapes.evaluate_run(
                    self,
                    &self.library,
                    shape,
                    cells,
                    &mut lanes,
                    &mut answers,
                    &mut answered,
                );
                if answers.is_empty() {
                    self.reaching += answered.len();
                    for (k, &x) in answered.iter().enumerate() {
                        self.keep_number(i + k, x, keep);
                    }
                    answered.clear();
                } else {
                    self.reaching += answers.len();
                    for (k, computed) in answers.drain(..).enumerate() {
                        self.keep(i + k, computed, keep);
                    }
                }
                i = j;
                continue;
            }
            gathered.clear();
            gathered_at.clear();
            while gathered.len() < vector::CHUNK {
                match next_marked(&self.marked, i, to) {
                    Some(p) => {
                        gathered.push(self.shapes.order()[p]);
                        gathered_at.push(p as u32);
                        i = p + 1;
                    }
                    None => {
                        i = to;
                        break;
                    }
                }
            }
            if gathered.len() < few {
                for &p in &gathered_at {
                    self.evaluate_alone(p as usize, keep);
                }
                continue;
            }
            self.shapes.evaluate_run(
                self,
                &self.library,
                shape,
                &gathered,
                &mut lanes,
                &mut answers,
                &mut answered,
            );
            if answers.is_empty() {
                self.reaching += answered.len();
                for (&p, &x) in gathered_at.iter().zip(&answered) {
                    self.keep_number(p as usize, x, keep);
                }
                answered.clear();
            } else {
                self.reaching += answers.len();
                for (&p, computed) in gathered_at.iter().zip(answers.drain(..)) {
                    self.keep(p as usize, computed, keep);
                }
            }
        }
        self.lanes = lanes;
        self.answers = answers;
        self.answered = answered;
        self.gathered = gathered;
        self.gathered_at = gathered_at;
    }

    /// A number a run computed for the cell at a position of the order, its
    /// cell's answer (KERNEL.43): nothing done where the value held is that
    /// number to the bit, as [`Machine::keep`] would do nothing, and no
    /// value made for it; else kept as `keep` keeps one, compared once.
    #[inline]
    fn keep_number(&mut self, p: usize, x: f64, keep: bool) {
        #[cfg(test)]
        {
            self.computed += 1;
        }
        if let Some(Computed::Value(Value::Number(held))) = &self.values[p] {
            if held.to_bits() == x.to_bits() {
                return;
            }
        }
        self.changed_to(p, Computed::Value(Value::Number(x)), keep);
    }

    /// The value the frame computed for the cell at a position of the order,
    /// kept when its bits differ from the one it holds (KERNEL.20).
    fn keep(&mut self, p: usize, computed: Computed, keep: bool) {
        #[cfg(test)]
        {
            self.computed += 1;
        }
        if let Some(h) = &self.values[p] {
            if same_computed(h, &computed) {
                return;
            }
        }
        self.changed_to(p, computed, keep);
    }

    /// A value that differs from the one its cell holds kept: the old one
    /// beside while a step yields, the new one in its sheet's array where one
    /// is kept (KERNEL.31), the cell counted as changed, and, but in a frame
    /// that computes every cell, the cells of the frame that read it marked.
    fn changed_to(&mut self, p: usize, computed: Computed, keep: bool) {
        let was = self.values[p].replace(computed);
        let reason = matches!(was, Some(Computed::NotComputed(_)));
        if keep {
            self.before.entry(p as u32).or_insert(was);
        }
        self.own_put(p, reason);
        if self.changed.len() < self.changed.capacity() {
            self.changed.push(p as u32);
        } else {
            self.changed_over = true;
        }
        if !self.every {
            let Machine {
                shapes,
                marked,
                whole,
                ..
            } = self;
            let id = shapes.order()[p];
            shapes.reach().each(id, &mut |r| match r {
                Reached::Cell(p) => marked[p / 64] |= 1 << (p % 64),
                Reached::Shape(s) | Reached::Parts(s) => whole[s] = true,
            });
        }
    }

    /// A frame's end (KERNEL.20): the cells it changed and the cells written
    /// since the last copied into their twins, with their planes' bytes and
    /// their arrays' slots, however many changed into twins kept as arrays
    /// (KERNEL.36), when the cells it would copy through twins' own cells
    /// are no more than half the cells and every plane an engine views is
    /// kept; else, or when a twin's cell would gain or lose its value, every
    /// sheet walked. Either way the twin cells whose content or reason
    /// changed are kept for the next frame, and are the frame's changes
    /// unless they are more than the lists tell (KERNEL.42). And what the
    /// frame changed and computed is kept for the counts until the next
    /// frame's end (KERNEL.48): each sheet's share counted by the copy as it
    /// copies each changed cell, or over the list where the end walked.
    fn snapshot(&mut self) {
        self.marked.fill(0);
        self.whole.fill(false);
        self.recheck.fill(false);
        self.twin_changed.clear();
        self.alike.clear();
        self.alike_over = false;
        self.changed_sheets.fill(0);
        if !(self.copies
            && self.incremental
            && !self.changed_over
            && self.few()
            && self.planes_kept()
            && self.copy())
        {
            self.walk();
            self.count_on_sheets();
        }
        // The list holds each formula cell the frame changed once, with room
        // for every one, so its length is the count.
        debug_assert!(!self.changed_over, "the list holds every formula cell");
        self.last_changed = self.changed.len();
        self.last_reached = self.reaching;
        self.changed.clear();
        self.changed_over = false;
        self.written.clear();
        // The frame began with the next frame's flag clear, and no write
        // lands during a frame, so the flag is the end's own.
        self.changes_whole = self.every_next || self.alike_over || self.reshaped;
        self.reshaped = false;
    }

    /// Each rows' sheet's share of the formula cells the frame changed,
    /// counted over the list, for a frame's end that walked (KERNEL.48); a
    /// copy counts each cell as it copies it, its place read already. The
    /// list stands in the order, by level and then shape, so its cells come
    /// a sheet at a time, each run counted in a register and added once: a
    /// count in memory raised at every cell waited on its own last store.
    fn count_on_sheets(&mut self) {
        self.changed_sheets.fill(0);
        let order = self.shapes.order();
        let (mut on, mut run) = (usize::MAX, 0);
        for &p in &self.changed {
            let s = order[p as usize].0;
            if s != on {
                if let Some(n) = self.changed_sheets.get_mut(on) {
                    *n += run;
                }
                (on, run) = (s, 0);
            }
            run += 1;
        }
        if let Some(n) = self.changed_sheets.get_mut(on) {
            *n += run;
        }
    }

    /// Whether the cells a copy writes through their twins' own cells, a
    /// lookup in a twin's B-tree each, are no more than half the cells: the
    /// cells written since the last frame, and the changed formula cells of
    /// the sheets kept with no array (KERNEL.36). Copied into an array, a
    /// cell costs a third of what the walk costs a cell of the grid, and
    /// never counts; through a twin's cells, the copy's cost and the walk's,
    /// which overwrites the twin in one pass, crossed between 40% and 75%
    /// of the cells changed, measured, so the half stands for them. The
    /// changed list is read only where neither bound decides, the formulas
    /// of the sheets with no array few enough to fit in the half, or every
    /// formula one of theirs, and no further than the half: read past half
    /// at every frame of a sparse sheet, it made a frame whose end walks a
    /// tenth dearer than before.
    fn few(&self) -> bool {
        let Some(mut left) = (self.cells / 2).checked_sub(self.written.len()) else {
            return false;
        };
        if self.changed.len() <= left {
            return true;
        }
        let reach = self.shapes.reach();
        let through: usize = (0..self.sheets)
            .filter(|&s| self.dense[s].is_none())
            .map(|s| reach.formulas_on(s))
            .sum();
        if through <= left {
            return true;
        }
        let order = self.shapes.order();
        if through == order.len() {
            return false;
        }
        for &p in &self.changed {
            if self.dense[order[p as usize].0].is_none() {
                let Some(less) = left.checked_sub(1) else {
                    return false;
                };
                left = less;
            }
        }
        true
    }

    /// Whether every sheet an engine views as a plane has one kept, or has
    /// one no raster can hold, so that a frame's end that copies its changed
    /// cells leaves every plane whole.
    fn planes_kept(&self) -> bool {
        (0..self.sheets).all(|s| {
            !self.wanted[s].get()
                || self.planes[s].is_some()
                || !plane::fits(self.extents[s], CELLS as u64)
        })
    }

    /// The cells the frame changed and the cells written since the last
    /// frame's end, each copied into its twin cell, its array's slot, its
    /// plane's byte and its reason; `false` when a twin's cell would gain or
    /// lose a value, which leaves the twin to the walk. A changed formula
    /// cell of a sheet kept as an array leaves its twin cell behind its slot,
    /// made when read (KERNEL.27): the twin's cells are a B-tree of every
    /// cell, and its lookup and its write of each changed cell were most of
    /// the frame's end. Each changed formula cell is counted on its sheet as
    /// it is copied (KERNEL.48), its place read here already: a pass over the
    /// list for it read the order's places a second time, far apart on a
    /// large board, 2% of the frame of Life's board of 1280 by 800
    /// (measured, against nothing measurable counted here).
    fn copy(&mut self) -> bool {
        let Machine {
            book,
            sheets,
            shapes,
            values,
            twin_reasons,
            wanted,
            planes,
            dense,
            behind,
            lagging,
            lazy,
            changed,
            changed_sheets,
            written,
            twin_changed,
            alike,
            alike_over,
            every_next,
            ..
        } = self;
        let (rows, twins) = book.sheets.split_at_mut(*sheets);
        let mut to = Twins {
            sheets: *sheets,
            twins,
            twin_reasons,
            wanted,
            planes,
            dense,
            behind,
            lagging,
            twin_changed,
            alike,
            alike_over,
            every_next,
        };
        // A formula cell the frame changed: its value alone makes its twin
        // cell, its byte and its slot.
        for &p in changed.iter() {
            #[cfg(test)]
            crate::sheet::work::copied();
            let (id, value) = (shapes.order()[p as usize], values[p as usize].as_ref());
            changed_sheets[id.0] += 1;
            let (content, byte) = (formula_twin(value), plane::formula_byte(value));
            let put = match *lazy {
                true => to.put_behind(id, content, value, byte),
                false => to.put(id, content, value, byte),
            };
            if !put {
                return false;
            }
        }
        for &id in written.iter() {
            #[cfg(test)]
            crate::sheet::work::copied();
            let Some(cell) = rows[id.0].cells.get(&(id.1, id.2)) else {
                return false;
            };
            let value = is_formula(&cell.content).then(|| value_at(shapes, values, id));
            let byte = plane::kept_byte(&cell.content, value.flatten());
            if !to.put(id, twin_content(cell, value), value.flatten(), byte) {
                return false;
            }
        }
        true
    }

    /// Every sheet's values copied into its twin: a value cell as it
    /// stands, a formula cell's value of the last complete frame; a formula
    /// with no value leaves its twin cell empty, and one not computed leaves
    /// its reason beside it. The twin's cells are overwritten where they
    /// stand, and made again only when a write changed which cells hold a
    /// value (KERNEL.20). A sheet an engine views as a plane has its plane
    /// kept from the same values in the same walk (KERNEL.24), and every
    /// twin dense enough its array, each slot what a read of the twin's cell
    /// answers (KERNEL.20). The twin cells whose content or reason the walk
    /// changes are kept for the next frame; a twin made again, or more
    /// changes than they hold, and the next frame computes every cell. A
    /// twin's cell left behind its slot (KERNEL.27) is made where the walk
    /// passes it, its slot telling, before the walk overwrites it, whether
    /// the cell changed, so that what the walk finds changed is what a read
    /// of the twin changed (KERNEL.36: each made first was a lookup in the
    /// twin's B-tree, which made the walk after a frame's copy dearer than
    /// the walk); a sheet whose array the walk will not keep over the same
    /// extent has its cells behind made first, as every sheet's were.
    fn walk(&mut self) {
        for s in 0..self.sheets {
            let kept = self.dense[s]
                .as_ref()
                .is_some_and(|k| k.stands_over(self.extents[s]));
            if !kept {
                self.make_twin(s);
            }
        }
        std::mem::swap(&mut self.twin_reasons, &mut self.reasons_before);
        self.twin_reasons.clear();
        self.twin_changed.clear();
        self.alike.clear();
        self.alike_over = false;
        let sheets = self.sheets;
        let reach = self.shapes.reach();
        let (rows, twins) = self.book.sheets.split_at_mut(sheets);
        for (s, (sheet, twin)) in rows.iter().zip(twins.iter_mut()).enumerate() {
            #[cfg(test)]
            crate::sheet::work::swept(sheet.cells.len());
            let t = sheets + s;
            let mut raster = if self.wanted[s].get() {
                plane::Raster::over(self.extents[s], CELLS as u64, self.planes[s].take())
            } else {
                None
            };
            let mut kept = Dense::over(
                self.extents[s],
                dense::limit(sheet.cells.len(), CELLS as u64),
                self.dense[s].take(),
                sliced(reach, t),
            );
            let mut fits = true;
            let mut whole = true;
            // The twin's cells behind their slots, each read where the walk
            // passes it, in the array kept over the same extent.
            let lags = std::mem::take(&mut self.lagging[s]);
            let mut behind = std::mem::take(&mut self.behind[s]);
            // The twin's cells overwritten where they stand, in this walk,
            // while each lies where the sheet's next cell with a value lies,
            // as frame after frame they do (KERNEL.20); the first that does
            // not stops the writes, and the twin is made again below.
            let mut stand = twin.cells.iter_mut();
            let mut in_place = true;
            for (&(row, col), cell) in &sheet.cells {
                // A formula's value, looked up once for the twin and the plane.
                let value = is_formula(&cell.content)
                    .then(|| value_at(&self.shapes, &self.values, (s, row, col)));
                if let Some(r) = raster.as_mut() {
                    fits &= r.set(row, col, plane::kept_byte(&cell.content, value.flatten()));
                }
                let reason = match value {
                    Some(Some(Computed::NotComputed(r))) => Some(r),
                    _ => None,
                };
                if let Some(r) = reason {
                    self.twin_reasons.insert((t, row, col), r.clone());
                }
                let was = self.reasons_before.get(&(t, row, col));
                if was != reason {
                    note(&mut self.twin_changed, &mut self.every_next, (t, row, col));
                    // A reason turned to another, the formula holding no
                    // value either side (any other turn makes the twin
                    // again below).
                    if was.is_some() && reason.is_some() {
                        let bits = alike(true, (Spelled::Nothing, was), (Spelled::Nothing, reason));
                        note_alike(&mut self.alike, &mut self.alike_over, (t, row, col), bits);
                    }
                }
                let content = twin_content(cell, value);
                // For a twin cell behind its slot, whether the slot, which
                // holds the cell's last value exactly, holds its new one.
                let mut same: Option<bool> = None;
                // The records that spell a changed cell as before (KERNEL.42),
                // told before its slot or its cell is overwritten.
                let mut bits = 0;
                // Every cell of the sheet writes its slot, so a store kept
                // over the same extent as the last frame's holds nothing
                // stale: a write only adds cells and changes them, and a cell
                // added since was an empty slot.
                if let Some(k) = kept.as_mut() {
                    match k.index(row, col) {
                        Some(i) => {
                            if lags && behind.get(i / 64).is_some_and(|w| w >> (i % 64) & 1 == 1) {
                                let slot = k.slot_at(i);
                                let held = same_slot(slot, content.as_ref());
                                if !held {
                                    bits = alike(
                                        value.is_some(),
                                        (Spelled::slot(slot), None),
                                        (Spelled::content(content.as_ref()), None),
                                    );
                                }
                                same = Some(held);
                            }
                            k.write_at(i, |slot| refresh(slot, content.as_ref()));
                        }
                        None => whole = false,
                    }
                }
                if let (Some(content), true) = (content, in_place) {
                    match stand.next() {
                        Some((&at, held)) if at == (row, col) => {
                            let changed = match same {
                                // Behind: made from its new content, changed
                                // when its slot held another.
                                Some(same) => {
                                    held.content = content;
                                    !same
                                }
                                None => {
                                    let differs = !same_content(&held.content, &content);
                                    if differs {
                                        bits = alike(
                                            value.is_some(),
                                            (Spelled::content(Some(&held.content)), None),
                                            (Spelled::content(Some(&content)), None),
                                        );
                                        held.content = content;
                                    }
                                    differs
                                }
                            };
                            if changed {
                                note(&mut self.twin_changed, &mut self.every_next, (t, row, col));
                                note_alike(
                                    &mut self.alike,
                                    &mut self.alike_over,
                                    (t, row, col),
                                    bits,
                                );
                            }
                        }
                        _ => in_place = false,
                    }
                }
            }
            let in_place = in_place && stand.next().is_none();
            // A cell outside the extent would mean a stale extent: the walk
            // answers rather than a plane that might be wrong.
            self.planes[s] = raster.filter(|_| fits);
            // And the twin's cells rather than an array that might be.
            self.dense[s] = kept.filter(|_| whole);
            // A bit for each slot, every one clear, the twin's cells made.
            let words = self.dense[s].as_ref().map_or(0, |k| k.slots().div_ceil(64));
            behind.clear();
            behind.resize(words, 0);
            self.behind[s] = behind;
            if !in_place {
                // A write changed which cells hold a value: the twin made
                // again from the sheet's, its extent with it, and the next
                // frame computes every cell.
                self.every_next = true;
                twin.cells = sheet
                    .cells
                    .iter()
                    .filter_map(|(&(row, col), cell)| {
                        let value = is_formula(&cell.content)
                            .then(|| value_at(&self.shapes, &self.values, (s, row, col)));
                        let content = twin_content(cell, value)?;
                        Some(((row, col), Cell { content, style: 0 }))
                    })
                    .collect();
                self.extents[t] = twin.extent();
            }
        }
    }

    /// Every twin cell a frame's end left behind its slot made from it
    /// (KERNEL.27): the twin's cells as a copy at every frame's end would
    /// have left them. A formula's twin cell is its value as a value cell
    /// holds it, which its slot holds exactly; a value cell's twin cell is
    /// copied where it stands at every frame's end and is never behind.
    fn make_twins(&mut self) {
        for s in 0..self.sheets {
            self.make_twin(s);
        }
    }

    /// One rows' sheet's twin cells left behind their slots made from them,
    /// a lookup in the twin's B-tree each.
    fn make_twin(&mut self, s: usize) {
        if !std::mem::take(&mut self.lagging[s]) {
            return;
        }
        let Machine {
            book,
            sheets,
            dense,
            behind,
            ..
        } = self;
        let Some(k) = dense[s].as_ref() else {
            behind[s].fill(0);
            return;
        };
        let twin = &mut book.sheets[*sheets + s];
        for (w, word) in behind[s].iter_mut().enumerate() {
            let mut bits = std::mem::take(word);
            while bits != 0 {
                let (slot, at) = k.at(w * 64 + bits.trailing_zeros() as usize);
                bits &= bits - 1;
                match twin.cells.get_mut(&at) {
                    Some(cell) => cell.content = value_content(slot),
                    None => debug_assert!(false, "a twin cell behind its slot is a cell"),
                }
            }
        }
    }

    // ---- the view -------------------------------------------------------

    /// A window of a sheet, a twin's included, through a projection: `grid`
    /// for the record, its `value` rows from the last complete frame;
    /// `values` for the record less its `formula` rows, what a host draws
    /// (KERNEL.37); `changes` for the rows of the values of the cells the
    /// last complete frame changed, or `whole` where the machine cannot say
    /// which (KERNEL.42); or `plane` for a byte a cell. `window` `None` is
    /// the sheet's extent, as the machine keeps it.
    pub fn view(
        &self,
        projection: &str,
        sheet: &str,
        window: Option<&str>,
    ) -> Result<Viewed, Refusal> {
        let kept = Some(self as &dyn Values);
        match projection {
            "grid" => {
                let w = window_over(&self.book, sheet, window, kept)?;
                Ok(Viewed::Text(view_text_valued(&self.book, &w, self)))
            }
            "values" => {
                let w = window_over(&self.book, sheet, window, kept)?;
                Ok(Viewed::Text(view_text_values(&self.book, &w, self)))
            }
            "changes" => {
                let w = window_over(&self.book, sheet, window, kept)?;
                Ok(Viewed::Text(self.changes(&w)))
            }
            "plane" => {
                let w = window_over(&self.book, sheet, window, kept)?;
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

    /// A window's changes (KERNEL.42): the cells whose rows in the window's
    /// values the last complete frame changed, since the frame before it,
    /// found among the twin cells its end changed, less those whose change
    /// the window's record spells as before; or whole, before the first
    /// frame since the load, a resumed grid's included, after a frame whose
    /// changes are more than the lists tell, after a write since that
    /// changed a formula or added a cell, and while a value written since
    /// lies in the window. A twin is written only at a frame's end, so a
    /// value written since leaves its window still.
    fn changes(&self, w: &Window) -> String {
        let at = self.frame;
        let since = at.saturating_sub(1);
        let i = self.book.find_sheet(&w.sheet).unwrap_or(0);
        let r = w.range;
        let inside =
            |id: &CellId| id.1 >= r.top && id.1 <= r.bottom && id.2 >= r.left && id.2 <= r.right;
        let rows = i < self.sheets;
        let whole = self.changes_whole
            || self.reshaped
            || (rows && self.written.iter().any(|id| id.0 == i && inside(id)));
        if whole {
            return changes_text(&self.book, w, self, (since, at), None);
        }
        let (t, bit) = match rows {
            true => (self.sheets + i, SHEET_ALIKE),
            false => (i, TWIN_ALIKE),
        };
        let mut cells: Vec<(u32, u32)> = self
            .twin_changed
            .iter()
            .filter(|id| id.0 == t && inside(id))
            .filter(|id| !self.alike.iter().any(|(a, b)| a == *id && b & bit != 0))
            .map(|id| (id.1, id.2))
            .collect();
        cells.sort_unstable();
        cells.dedup();
        changes_text(&self.book, w, self, (since, at), Some(&cells))
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
    /// the range holds gives that cell up, since the model reads a child
    /// through its master. Its first cell outside the range, by rows and then
    /// columns, becomes its master, holding the text moved there once, and
    /// its other cells stay its children: a release costs one move whatever
    /// the fill's size, where it gave every cell its own moved text, 62,963
    /// of them for a cell written over Life's rule (KERNEL.34). One walk of
    /// the sheet finds every released fill's first cell. A text moved twice
    /// is the text moved once only where the first move keeps it whole
    /// ([`crate::refers::moves_whole`]); a fill whose text the move would not
    /// keep whole gives each of its cells its own text, as every release
    /// did, and its cells are placed again when the write settles.
    fn release_masters(&mut self, si: usize, range: A1Range, log: &mut Log) {
        let inside = |row: u32, col: u32| {
            row >= range.top && row <= range.bottom && col >= range.left && col <= range.right
        };
        let masters: Vec<(u32, u32, u32, String, String)> = self.book.sheets[si]
            .cells
            .range((range.top, 0)..=(range.bottom, u32::MAX))
            .filter(|((_, col), _)| *col >= range.left && *col <= range.right)
            .filter_map(|(&(row, col), c)| match &c.content {
                Content::SharedMaster {
                    text,
                    range: fill,
                    si: index,
                } => Some((row, col, *index, text.clone(), fill.clone())),
                _ => None,
            })
            .collect();
        if masters.is_empty() {
            return;
        }
        let released: BTreeSet<u32> = masters.iter().map(|m| m.2).collect();
        let sheet = &self.book.sheets[si];
        #[cfg(test)]
        crate::sheet::work::walked(sheet.cells.len());
        let mut firsts: HashMap<u32, ((u32, u32), u32)> = HashMap::new();
        for (&(row, col), c) in &sheet.cells {
            if let Content::SharedChild { si: index } = c.content {
                if !inside(row, col) && released.contains(&index) {
                    firsts.entry(index).or_insert(((row, col), c.style));
                }
            }
        }
        let mut each_own: HashMap<u32, (u32, u32, String)> = HashMap::new();
        for (m_row, m_col, index, text, fill) in masters {
            let Some(&((row, col), style)) = firsts.get(&index) else {
                // No cell of the fill outside the range: nothing to give up.
                continue;
            };
            let at = (
                i64::from(row) - i64::from(m_row),
                i64::from(col) - i64::from(m_col),
            );
            let whole = crate::refers::moves_whole(&text, at, at);
            #[cfg(test)]
            crate::sheet::work::released(whole);
            if whole {
                let moved = shift_a1_references(&text, at.0, at.1);
                self.set_cell(
                    (si, row, col),
                    Some(Cell {
                        content: Content::SharedMaster {
                            text: moved,
                            range: fill,
                            si: index,
                        },
                        style,
                    }),
                    log,
                );
            } else {
                each_own.insert(index, (m_row, m_col, text));
            }
        }
        if each_own.is_empty() {
            return;
        }
        let sheet = &self.book.sheets[si];
        #[cfg(test)]
        crate::sheet::work::walked(sheet.cells.len());
        let children: Vec<((u32, u32), u32, u32)> = sheet
            .cells
            .iter()
            .filter_map(|(&(row, col), c)| match c.content {
                Content::SharedChild { si: index }
                    if !inside(row, col) && each_own.contains_key(&index) =>
                {
                    Some(((row, col), c.style, index))
                }
                _ => None,
            })
            .collect();
        for ((row, col), style, index) in children {
            let (m_row, m_col, text) = &each_own[&index];
            let moved = shift_a1_references(
                text,
                i64::from(row) - i64::from(*m_row),
                i64::from(col) - i64::from(*m_col),
            );
            self.set_cell(
                (si, row, col),
                Some(Cell {
                    content: Content::Formula(moved),
                    style,
                }),
                log,
            );
            log.released.insert((si, row, col));
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
        let differs = self.differing(si, range, text, log).len();
        if differs == 0 {
            return Ok(());
        }
        self.release_masters(si, range, log);
        // Each cell of the range looked up once: what it held kept for the
        // undo, the fill's content put in its place, its format kept
        // (KERNEL.34; the cells were looked up four times).
        let sheet = &mut self.book.sheets[si];
        let index = sheet.next_shared(range);
        for row in range.top..=range.bottom {
            for col in range.left..=range.right {
                let content = filled(range, text, index, row, col);
                let old = match sheet.cells.entry((row, col)) {
                    Entry::Occupied(mut held) => {
                        let cell = held.get_mut();
                        Some(Cell {
                            content: std::mem::replace(&mut cell.content, content),
                            style: cell.style,
                        })
                    }
                    Entry::Vacant(empty) => {
                        empty.insert(Cell { content, style: 0 });
                        log.added += 1;
                        None
                    }
                };
                log.undo.push(((si, row, col), old));
                log.changed.entry((si, row, col)).or_insert(line);
            }
        }
        log.sheets.insert(si);
        log.written += differs;
        Ok(())
    }

    /// The cells of a range whose formula would read otherwise with `text`
    /// written over the range: each cell's formula as the formula bar shows
    /// it, against the written text moved to the cell, as
    /// [`Machine::differing_by_cell`] compares them, found once for each
    /// formula the range covers (KERNEL.34). Comparing every cell moved two
    /// texts a cell, 1.8 s for Life's rule written again with its own text.
    ///
    /// A cell holding no formula differs. A fill's cells are read through
    /// its master: one whose master is the range's first cell and holds the
    /// written text is the same at every cell, nothing moved. Otherwise, so
    /// long as moving either text over the fill's cells keeps it whole
    /// ([`crate::refers::moves_whole`]) and neither holds a future prefix,
    /// two texts moved to a cell read alike there exactly when they make the
    /// same references from it with the same text between, which does not
    /// depend on the cell: one cell where both texts stand moved is compared
    /// and stands for all of them, and the one or two where a text stands as
    /// written, the master and the range's first cell, perhaps spelled
    /// otherwise, are compared alone. A plain formula's cells each stand as
    /// written: where the written text moved over them keeps whole, a shape
    /// whose R1C1 text is not the written text's differs at every cell, as
    /// a cell reads its shape's R1C1 text and the written text moved reads
    /// its own. Every other cell is compared alone: a fill whose texts a
    /// move might not keep whole, a plain formula of the written text's R1C1
    /// text, and a plain formula an earlier row of the same write placed
    /// where the plan has not seen it.
    fn differing(&self, si: usize, range: A1Range, text: &str, log: &Log) -> Vec<(u32, u32)> {
        let sheet = &self.book.sheets[si];
        let masters = shared_masters(sheet);
        let start = (range.top, range.left);
        let same_at = |row: u32, col: u32| {
            let moved = shift_a1_references(
                text,
                i64::from(row) - i64::from(range.top),
                i64::from(col) - i64::from(range.left),
            );
            sheet
                .cells
                .get(&(row, col))
                .and_then(|c| formula_text(&c.content, row, col, &masters))
                .is_some_and(|old| old == strip_future_prefixes(&moved))
        };
        // The offsets of some cells from a cell, nearest and farthest.
        let span = |cells: &[(u32, u32)], (row, col): (u32, u32)| {
            let mut low = (i64::MAX, i64::MAX);
            let mut high = (i64::MIN, i64::MIN);
            for &(r, c) in cells {
                let at = (i64::from(r) - i64::from(row), i64::from(c) - i64::from(col));
                low = (low.0.min(at.0), low.1.min(at.1));
                high = (high.0.max(at.0), high.1.max(at.1));
            }
            (low, high)
        };
        let unprefixed = |t: &str| {
            !t.as_bytes()
                .windows(3)
                .any(|w| w.eq_ignore_ascii_case(b"_xl"))
        };
        let mut differs = Vec::new();
        let mut fills: BTreeMap<u32, Vec<(u32, u32)>> = BTreeMap::new();
        let mut plains: BTreeMap<usize, Vec<(u32, u32)>> = BTreeMap::new();
        let mut alone: Vec<(u32, u32)> = Vec::new();
        for row in range.top..=range.bottom {
            for col in range.left..=range.right {
                match sheet.cells.get(&(row, col)).map(|c| &c.content) {
                    Some(
                        Content::SharedMaster { si: index, .. }
                        | Content::SharedChild { si: index },
                    ) => {
                        fills.entry(*index).or_default().push((row, col));
                    }
                    Some(Content::Formula(_) | Content::DynamicFormula(_)) => {
                        let id = (si, row, col);
                        let fresh = !log.changed.contains_key(&id) && !log.released.contains(&id);
                        match self.shapes.place(id).filter(|_| fresh) {
                            Some(p) => plains.entry(p.shape).or_default().push((row, col)),
                            None => alone.push((row, col)),
                        }
                    }
                    _ => differs.push((row, col)),
                }
            }
        }
        for (index, cells) in &fills {
            let Some(&(own, m_row, m_col)) = masters.get(index) else {
                // A child whose master is gone reads no formula.
                differs.extend(cells);
                continue;
            };
            let master = (m_row, m_col);
            if master == start && own == text {
                continue;
            }
            let (low, high) = span(cells, master);
            let (w_low, w_high) = span(cells, start);
            let whole = unprefixed(own)
                && unprefixed(text)
                && crate::refers::moves_whole(own, low, high)
                && crate::refers::moves_whole(text, w_low, w_high);
            let mut moved_alike: Option<bool> = None;
            for &(row, col) in cells {
                let alike = if !whole || (row, col) == master || (row, col) == start {
                    same_at(row, col)
                } else {
                    *moved_alike.get_or_insert_with(|| same_at(row, col))
                };
                if !alike {
                    differs.push((row, col));
                }
            }
        }
        let mut written_r1c1: Option<String> = None;
        for (shape, cells) in &plains {
            let (w_low, w_high) = span(cells, start);
            let whole = unprefixed(text) && crate::refers::moves_whole(text, w_low, w_high);
            if whole {
                let theirs = self
                    .shapes
                    .shape(*shape)
                    .map(|s| crate::refers::r1c1(&s.text, s.anchor.1, s.anchor.2));
                let ours = written_r1c1
                    .get_or_insert_with(|| crate::refers::r1c1(text, range.top, range.left));
                if theirs.as_ref() != Some(ours) {
                    differs.extend(cells);
                    continue;
                }
            }
            alone.extend(cells);
        }
        differs.extend(alone.into_iter().filter(|&(row, col)| !same_at(row, col)));
        differs
    }

    /// The same-text test as every write made it before KERNEL.34, every
    /// cell's text and the written text moved to it and compared: the
    /// reference [`Machine::differing`] is held to.
    #[cfg(test)]
    fn differing_by_cell(&self, si: usize, range: A1Range, text: &str) -> Vec<(u32, u32)> {
        let masters = shared_masters(&self.book.sheets[si]);
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
                if old_text.as_deref() != Some(strip_future_prefixes(&moved).as_str()) {
                    differs.push((row, col));
                }
            }
        }
        differs
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
            // Every cell the write changed or a release gave its own text,
            // placed again at once, and the plan made a load's (KERNEL.34).
            let again: BTreeSet<CellId> =
                log.changed.keys().chain(&log.released).copied().collect();
            let again: Vec<CellId> = again.into_iter().collect();
            self.shapes.replace_many(&self.book, &again, &self.library);
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
            // A changed cell's reason uncounted from its sheet's array, its
            // value forgotten below (KERNEL.31).
            for id in log.changed.keys() {
                if let Some(p) = saved.reach().position(*id) {
                    let reason = matches!(self.values[p], Some(Computed::NotComputed(_)));
                    if reason && self.own[id.0].is_some() {
                        self.own_reasons[id.0] -= 1;
                    }
                }
            }
            // Each value moved to its cell's new position in the order
            // (KERNEL.26).
            let mut moved: Vec<Option<Computed>> = vec![None; self.shapes.order().len()];
            for (&id, value) in saved.order().iter().zip(std::mem::take(&mut self.values)) {
                if let (Some(q), Some(v)) = (self.shapes.reach().position(id), value) {
                    moved[q] = Some(v);
                }
            }
            self.values = moved;
            // A twin's array a formula now reads by slices keeps its numbers
            // beside its slots from the next frame on, made from its slots;
            // one no formula reads drops them (KERNEL.43).
            let reach = self.shapes.reach();
            for (s, kept) in self.dense.iter_mut().enumerate() {
                if let Some(k) = kept {
                    k.shade(sliced(reach, self.sheets + s));
                }
            }
        }
        // A write never empties a cell: it only adds cells or changes them.
        // A formula changed or a cell added, and the next frame computes
        // every cell; a value cell changed, and the next follows it to the
        // cells that read it (KERNEL.20).
        self.cells += log.added;
        if formula_changed || log.added > 0 {
            self.every_next = true;
            // A cell's kind or a sheet's extent may have changed, which no
            // twin cell's change carries: the changes whole until the next
            // frame's end and at it (KERNEL.42).
            self.reshaped = true;
            self.fit();
        }
        for id in log.changed.keys() {
            if let Some(p) = self.shapes.reach().position(*id) {
                self.values[p] = None;
            }
            if !self.shapes.holds(*id) {
                self.written.push(*id);
            }
        }
        // The frame's own sheets' arrays (KERNEL.31): kept or made again as
        // the plan and the extents now stand, where either may have moved;
        // else each written value cell's slot written, a cell's place and its
        // sheet's extent standing.
        if formula_changed || log.added > 0 {
            let changed: Vec<CellId> = log.changed.keys().copied().collect();
            self.own_made(Some(&changed));
        } else if self.owns {
            for id in log.changed.keys() {
                if self.own[id.0].is_none() {
                    continue;
                }
                let v = self.slot_value(*id);
                if let Some(Some(kept)) = self.own.get_mut(id.0) {
                    if let Some(i) = kept.index(id.1, id.2) {
                        kept.write_at(i, |slot| *slot = v);
                    }
                }
            }
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
    /// The last complete frame's value: while a step yields, a value the
    /// frame in progress overwrote is read from beside it (KERNEL.20).
    fn computed(&self, id: CellId) -> Option<&Computed> {
        let p = self.shapes.reach().position(id)?;
        if !self.before.is_empty() {
            if let Some(was) = self.before.get(&(p as u32)) {
                return was.as_ref();
            }
        }
        self.values[p].as_ref()
    }

    /// A twin whose cells a frame's end left behind its slots (KERNEL.27).
    fn lags(&self, sheet: usize) -> bool {
        sheet
            .checked_sub(self.sheets)
            .and_then(|s| self.lagging.get(s))
            .is_some_and(|&l| l)
    }

    /// A twin cell left behind its slot, made from the slot as
    /// [`Machine::make_twins`] makes it, so that a view of the twin reads
    /// what its cells would hold.
    fn content(&self, id: CellId) -> Option<Content> {
        let s = id.0.checked_sub(self.sheets)?;
        let k = self.dense.get(s)?.as_ref()?;
        let i = k.index(id.1, id.2)?;
        let bit = self.behind.get(s)?.get(i / 64)? >> (i % 64) & 1;
        (bit == 1).then(|| value_content(k.at(i).0))
    }

    /// Every sheet's extent, its twins' included, as the machine keeps it
    /// for the evaluator (KERNEL.37): made at the load, and kept at every
    /// write, every undone write and every frame's end.
    fn kept_extent(&self, sheet: usize) -> Option<Extent> {
        self.extents.get(sheet).copied()
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
            // A frame's own sheet kept as an array read by its index, as a
            // twin's is (KERNEL.31): offered only while none of its formulas
            // holds a reason, its slot holds what a read answers, nothing for
            // a cell holding none, and every cell stands inside its extent.
            if let Some(kept) = Env::dense(self, id.0) {
                return Ok(kept.get(id.1, id.2).cloned().unwrap_or(Value::Empty));
            }
            // The frame's value: computed in this frame by now, the order
            // putting each cell after what it reads, or kept from the last,
            // its reads unchanged (KERNEL.20); found by the cell's position.
            if let Some(p) = self.shapes.reach().position(id) {
                return match &self.values[p] {
                    Some(Computed::Value(v)) => Ok(v.clone()),
                    Some(Computed::NotComputed(r)) => Err(r.clone()),
                    // A formula with no value in this frame can only be read
                    // from a cycle, which the plan refuses before it is held.
                    None => Err(Reason::Cycle),
                };
            }
        } else if let Some(kept) = self.dense(id.0) {
            // A twin kept as an array is read by its index (KERNEL.20): the
            // value its cell holds, or nothing, beside which its reason may
            // stand, as below.
            if let Some(v) = kept.get(id.1, id.2).filter(|v| !matches!(v, Value::Empty)) {
                return Ok(v.clone());
            }
            return match self.twin_reasons.get(&id) {
                Some(r) => Err(r.clone()),
                None => Ok(Value::Empty),
            };
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
        Ok(content_value(&cell.content))
    }

    fn extent(&self, sheet: usize) -> Extent {
        self.extents.get(sheet).copied().flatten()
    }

    fn name(&self, name: &str, scope: Option<usize>, home: usize) -> Option<&str> {
        self.names.resolve(name, scope, home)
    }

    /// A twin's array, as its last frame's end left it (KERNEL.20); a frame's
    /// own sheet's, as the frame keeps it, while none of its formulas holds
    /// a reason (KERNEL.31), so that a range of it answers a reason as a walk
    /// of its cells does.
    fn dense(&self, sheet: usize) -> Option<&Dense> {
        if sheet < self.sheets {
            if !self.owns {
                return None;
            }
            return match self.own_reasons[sheet] {
                0 => self.own[sheet].as_ref(),
                _ => {
                    #[cfg(test)]
                    self.aside.set(self.aside.get() + 1);
                    None
                }
            };
        }
        self.dense.get(sheet - self.sheets)?.as_ref()
    }

    fn enter(&self, id: CellId) {
        self.here.set(id);
    }
}

impl Source for Machine {
    /// [`Env::read`]'s answer as a program's register holds it (KERNEL.44,
    /// `calc::program::number_of`), read where `read` reads it and no value
    /// made: a frame's own sheet's slot where it offers its array, an empty
    /// one nothing (KERNEL.31), a formula's value of the frame by its
    /// position, a twin's slot and the reason beside an empty one, a cell's
    /// content.
    #[inline]
    fn number(&self, id: CellId) -> Option<f64> {
        let reasonless =
            |id: &CellId| self.twin_reasons.is_empty() || !self.twin_reasons.contains_key(id);
        if id.0 < self.sheets {
            if let Some(kept) = Env::dense(self, id.0) {
                return match kept.get(id.1, id.2) {
                    Some(Value::Number(n)) => Some(*n),
                    Some(Value::Empty) | None => Some(0.0),
                    Some(_) => None,
                };
            }
            if let Some(p) = self.shapes.reach().position(id) {
                return match &self.values[p] {
                    Some(Computed::Value(Value::Number(n))) => Some(*n),
                    Some(Computed::Value(Value::Empty)) => Some(0.0),
                    _ => None,
                };
            }
        } else if let Some(kept) = Env::dense(self, id.0) {
            return match kept.get(id.1, id.2) {
                Some(Value::Number(n)) => Some(*n),
                Some(Value::Empty) | None => reasonless(&id).then_some(0.0),
                Some(_) => None,
            };
        }
        match self
            .book
            .sheets
            .get(id.0)
            .and_then(|s| s.cells.get(&(id.1, id.2)))
        {
            Some(cell) => match &cell.content {
                Content::Number(n) => Some(*n),
                Content::Text(_) | Content::Bool(_) | Content::Error(_) => None,
                // `content_value` reads anything else as nothing.
                _ => Some(0.0),
            },
            None => reasonless(&id).then_some(0.0),
        }
    }
}

/// A formula cell's value in the frame's values, kept at its position in the
/// order (KERNEL.26); `None` for a cell holding no formula, or a formula with
/// no value yet.
fn value_at<'v>(
    shapes: &Shapes,
    values: &'v [Option<Computed>],
    id: CellId,
) -> Option<&'v Computed> {
    values.get(shapes.reach().position(id)?)?.as_ref()
}

/// A cell's content as a read of it answers: a value as it stands, an error
/// text as its kind (an error this version does not hold as `#VALUE!`), and
/// nothing for a formula, which a read finds in the frame's values first. The
/// twin's read and its array's slots are spelled here once (KERNEL.20).
fn content_value(content: &Content) -> Value {
    match content {
        Content::Text(t) => Value::Text(t.clone()),
        Content::Number(n) => Value::Number(*n),
        Content::Bool(b) => Value::Bool(*b),
        Content::Error(e) => Value::Error(ErrorKind::parse(e).unwrap_or(ErrorKind::Value)),
        _ => Value::Empty,
    }
}

/// A cell's content in its twin: a value cell's as it stands, a formula's
/// value of the last complete frame as a value cell holds it; nothing for a
/// formula with no value, or one not computed, whose reason stands beside.
fn twin_content(cell: &Cell, value: Option<Option<&Computed>>) -> Option<Content> {
    match value {
        None => Some(cell.content.clone()),
        Some(value) => formula_twin(value),
    }
}

/// A formula cell's content in its twin, from its value alone.
fn formula_twin(value: Option<&Computed>) -> Option<Content> {
    match value {
        Some(Computed::Value(v)) => Some(value_content(v)),
        Some(Computed::NotComputed(_)) | None => None,
    }
}

/// A twin's slot overwritten with what a read of the cell's new content
/// answers, nothing for a cell that keeps no value; a text the slot holds
/// already is kept and never copied again, so that a frame allocates no more
/// than it did before the array (KERNEL.20).
fn refresh(slot: &mut Value, content: Option<&Content>) {
    match (slot, content) {
        (Value::Text(held), Some(Content::Text(t))) => {
            if *held != *t {
                held.clone_from(t);
            }
        }
        (slot, content) => *slot = content.map_or(Value::Empty, content_value),
    }
}

/// Whether a run may read a twin by slices, so that its array keeps its
/// numbers beside its slots (KERNEL.43): a shape the plan follows reads it,
/// or the plan holds a shape it does not follow, whose reads it cannot
/// tell. A spreadsheet's twins, which no formula reads, keep none, and a
/// frame's end that writes their slots writes nothing beside them.
fn sliced(reach: &Reach, twin: usize) -> bool {
    reach.readers_on(twin) > 0 || !reach.unfollowed().is_empty()
}

/// What a frame's end that copies its changed cells writes (KERNEL.20): the
/// twins, their reasons, the kept planes and arrays, the twin cells left
/// behind their slots (KERNEL.27), and the twin cells that changed, with
/// those a record spells as before (KERNEL.42).
struct Twins<'a> {
    sheets: usize,
    twins: &'a mut [Sheet],
    twin_reasons: &'a mut FxHashMap<CellId, Reason>,
    wanted: &'a [Slot<bool>],
    planes: &'a mut [Option<plane::Raster>],
    dense: &'a mut [Option<Dense>],
    behind: &'a mut [Vec<u64>],
    lagging: &'a mut [bool],
    twin_changed: &'a mut Vec<CellId>,
    alike: &'a mut Vec<(CellId, u8)>,
    alike_over: &'a mut bool,
    every_next: &'a mut bool,
}

impl Twins<'_> {
    /// One rows' cell copied into its twin: its twin cell's content, its
    /// slot, its byte in a kept plane, its reason (from its formula's
    /// value); `false` when the twin cell would gain or lose a value. The
    /// twin cell is made where it stands, so it is no longer behind its
    /// slot.
    fn put(
        &mut self,
        id: CellId,
        content: Option<Content>,
        value: Option<&Computed>,
        byte: u8,
    ) -> bool {
        let (s, row, col) = id;
        if let Some(Some(k)) = self.dense.get_mut(s) {
            if let Some(i) = k.index(row, col) {
                k.write_at(i, |slot| refresh(slot, content.as_ref()));
                self.behind[s][i / 64] &= !(1 << (i % 64));
            }
        }
        // The records that spell the change as before (KERNEL.42): told of a
        // reason before it is overwritten, and of a cell before it is.
        let mut bits = match content {
            None => self.reason_alike(id, value),
            Some(_) => 0,
        };
        let mut moved = match (self.twins[s].cells.get_mut(&(row, col)), content) {
            (Some(held), Some(content)) => {
                let differs = !same_content(&held.content, &content);
                if differs {
                    bits = alike(
                        value.is_some(),
                        (Spelled::content(Some(&held.content)), None),
                        (Spelled::content(Some(&content)), None),
                    );
                    held.content = content;
                }
                differs
            }
            (None, None) => false,
            _ => return false,
        };
        moved |= self.byte_and_reason(id, value, byte);
        if moved {
            let t = (self.sheets + s, row, col);
            note(self.twin_changed, self.every_next, t);
            note_alike(self.alike, self.alike_over, t, bits);
        }
        true
    }

    /// A formula cell the frame changed, on a sheet kept as an array: its
    /// slot, its byte and its reason, its twin cell left behind its slot to
    /// be made from it when read (KERNEL.27), which a formula's twin cell can
    /// be, the slot holding its value exactly and holding one exactly where
    /// the twin cell does; `false` when the twin cell would gain or lose a
    /// value. A sheet too sparse for an array, or a cell outside its array,
    /// is copied as [`Twins::put`] copies it.
    fn put_behind(
        &mut self,
        id: CellId,
        content: Option<Content>,
        value: Option<&Computed>,
        byte: u8,
    ) -> bool {
        let (s, row, col) = id;
        let Some(Some(k)) = self.dense.get_mut(s) else {
            return self.put(id, content, value, byte);
        };
        let Some(i) = k.index(row, col) else {
            return self.put(id, content, value, byte);
        };
        let slot = k.slot_at(i);
        if matches!(slot, Value::Empty) == content.is_some() {
            return false;
        }
        let mut moved = !same_slot(slot, content.as_ref());
        let mut bits = 0;
        if moved {
            bits = alike(
                value.is_some(),
                (Spelled::slot(slot), None),
                (Spelled::content(content.as_ref()), None),
            );
            k.write_at(i, |slot| refresh(slot, content.as_ref()));
            self.behind[s][i / 64] |= 1 << (i % 64);
            self.lagging[s] = true;
        }
        if content.is_none() {
            bits = self.reason_alike(id, value);
        }
        moved |= self.byte_and_reason(id, value, byte);
        if moved {
            let t = (self.sheets + s, row, col);
            note(self.twin_changed, self.every_next, t);
            note_alike(self.alike, self.alike_over, t, bits);
        }
        true
    }

    /// The records that spell a formula's reason turned to another as
    /// before (KERNEL.42): the twin's always, holding no cell either side,
    /// and the sheet's when the two are spelled alike; none for a reason
    /// gained or lost, which comes with a value lost or gained.
    fn reason_alike(&self, id: CellId, value: Option<&Computed>) -> u8 {
        let was = self.twin_reasons.get(&(self.sheets + id.0, id.1, id.2));
        let now = match value {
            Some(Computed::NotComputed(r)) => Some(r),
            _ => None,
        };
        match (was, now) {
            (Some(_), Some(_)) => alike(true, (Spelled::Nothing, was), (Spelled::Nothing, now)),
            _ => 0,
        }
    }

    /// A rows' cell's byte written in its sheet's kept plane, and its reason
    /// beside its twin cell (from its formula's value); whether the reason
    /// changed.
    fn byte_and_reason(&mut self, id: CellId, value: Option<&Computed>, byte: u8) -> bool {
        let (s, row, col) = id;
        if self.wanted[s].get() {
            if let Some(r) = self.planes[s].as_mut() {
                r.set(row, col, byte);
            }
        }
        let t = (self.sheets + s, row, col);
        let reason = match value {
            Some(Computed::NotComputed(r)) => Some(r),
            _ => None,
        };
        if self.twin_reasons.get(&t) == reason {
            return false;
        }
        match reason {
            Some(r) => self.twin_reasons.insert(t, r.clone()),
            None => self.twin_reasons.remove(&t),
        };
        true
    }
}

/// Whether a slot holds what a read of a twin cell's new content answers,
/// to the bit: [`refresh`]'s rule, asked before it writes, so that a slot
/// whose cell's value is unchanged is neither written nor its cell behind.
fn same_slot(slot: &Value, content: Option<&Content>) -> bool {
    match (slot, content) {
        (Value::Empty, None) => true,
        (Value::Number(x), Some(Content::Number(y))) => x.to_bits() == y.to_bits(),
        (Value::Text(x), Some(Content::Text(y))) => x == y,
        (Value::Bool(x), Some(Content::Bool(y))) => x == y,
        (Value::Error(k), Some(Content::Error(e))) => {
            *k == ErrorKind::parse(e).unwrap_or(ErrorKind::Value)
        }
        _ => false,
    }
}

/// A twin cell kept as changed for the next frame (KERNEL.20), or, past the
/// room kept for them, the next frame made to compute every cell.
fn note(changed: &mut Vec<CellId>, every_next: &mut bool, id: CellId) {
    if changed.len() < changed.capacity() {
        changed.push(id);
    } else {
        *every_next = true;
    }
}

/// A changed twin cell kept with the records that spell its change as
/// before, where any does (KERNEL.42), or, past [`ALIKE`] of them, the
/// frame's changes made whole.
fn note_alike(alike: &mut Vec<(CellId, u8)>, over: &mut bool, id: CellId, bits: u8) {
    if bits == 0 {
        return;
    }
    if alike.len() < ALIKE {
        alike.push((id, bits));
    } else {
        *over = true;
    }
}

/// A twin cell's content as a record spells it, told apart from another
/// without printing either (KERNEL.42): a number a cell can hold by its
/// bits, its shortest text being one number's alone; a number none can
/// hold as the error `#NUM!`, which the record prints for it; nothing for
/// no cell.
#[derive(Clone, Copy, PartialEq)]
enum Spelled<'a> {
    Nothing,
    Number(u64),
    Text(&'a str),
    Bool(bool),
    Error(&'a str),
}

impl<'a> Spelled<'a> {
    fn content(content: Option<&'a Content>) -> Spelled<'a> {
        match content {
            Some(Content::Number(n)) => Spelled::number(*n),
            Some(Content::Text(t)) => Spelled::Text(t),
            Some(Content::Bool(b)) => Spelled::Bool(*b),
            Some(Content::Error(e)) => Spelled::Error(e),
            _ => Spelled::Nothing,
        }
    }

    /// A twin's slot, which holds its cell's value exactly.
    fn slot(slot: &'a Value) -> Spelled<'a> {
        match slot {
            Value::Number(n) => Spelled::number(*n),
            Value::Text(t) => Spelled::Text(t),
            Value::Bool(b) => Spelled::Bool(*b),
            Value::Error(k) => Spelled::Error(k.text()),
            Value::Empty => Spelled::Nothing,
        }
    }

    fn number(n: f64) -> Spelled<'a> {
        match n.is_finite() {
            true => Spelled::Number(n.to_bits()),
            false => Spelled::Error(ErrorKind::Num.text()),
        }
    }

    /// As a formula's `value` row spells it, which prints 0 for -0.
    fn valued(self) -> Spelled<'a> {
        match self {
            Spelled::Number(bits) if f64::from_bits(bits) == 0.0 => Spelled::Number(0),
            other => other,
        }
    }
}

/// A reason as a `value` row spells it: by its name, or as a cycle.
fn reason_spelled(r: Option<&Reason>) -> Option<Option<&str>> {
    r.map(|r| match r {
        Reason::Function(name) | Reason::Construct(name) => Some(name.as_str()),
        Reason::Cycle => None,
    })
}

/// Which records spell a twin cell's change, from its content and reason
/// to others, as before (KERNEL.42): the twin's, a `cell` row of its
/// content and nothing for none, [`TWIN_ALIKE`]; and its sheet's,
/// [`SHEET_ALIKE`], a value cell's `cell` row as the twin's and a
/// formula's `value` row, its value with 0 for -0 or else its reason.
fn alike(
    formula: bool,
    (was, was_reason): (Spelled, Option<&Reason>),
    (now, now_reason): (Spelled, Option<&Reason>),
) -> u8 {
    let twin = was == now;
    let sheet = match (formula, was, now) {
        (false, _, _) => twin,
        (true, Spelled::Nothing, Spelled::Nothing) => {
            reason_spelled(was_reason) == reason_spelled(now_reason)
        }
        (true, was, now) => was.valued() == now.valued(),
    };
    let mut bits = 0;
    if twin {
        bits |= TWIN_ALIKE;
    }
    if sheet {
        bits |= SHEET_ALIKE;
    }
    bits
}

/// The first position from `from` and before `to` whose bit is set.
fn next_marked(marked: &[u64], from: usize, to: usize) -> Option<usize> {
    let mut w = from / 64;
    let mut word = *marked.get(w)? & (!0u64 << (from % 64));
    loop {
        if word != 0 {
            let p = w * 64 + word.trailing_zeros() as usize;
            return (p < to).then_some(p);
        }
        w += 1;
        if w * 64 >= to {
            return None;
        }
        word = *marked.get(w)?;
    }
}

/// Whether two values are the same to the bit, a number's sign of zero
/// included: what a cell reading either would compute alike (KERNEL.20).
fn same_value(a: &Value, b: &Value) -> bool {
    match (a, b) {
        (Value::Number(x), Value::Number(y)) => x.to_bits() == y.to_bits(),
        _ => a == b,
    }
}

/// [`same_value`] over a formula's answer, a reason the same reason.
fn same_got(a: &Result<Value, Reason>, b: &Result<Value, Reason>) -> bool {
    match (a, b) {
        (Ok(x), Ok(y)) => same_value(x, y),
        (Err(x), Err(y)) => x == y,
        _ => false,
    }
}

/// [`same_value`] over a cell's computed value.
fn same_computed(a: &Computed, b: &Computed) -> bool {
    match (a, b) {
        (Computed::Value(x), Computed::Value(y)) => same_value(x, y),
        (Computed::NotComputed(x), Computed::NotComputed(y)) => x == y,
        _ => false,
    }
}

/// [`same_value`] over a twin cell's content.
fn same_content(a: &Content, b: &Content) -> bool {
    match (a, b) {
        (Content::Number(x), Content::Number(y)) => x.to_bits() == y.to_bits(),
        _ => a == b,
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
    use crate::calc::program;
    use crate::calc::shape::Shapes;
    use crate::projection::Window;
    use crate::reader::read_forms;
    use crate::view::{view_text_valued, view_text_values, whole, window_of};

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

    /// The values of a window (KERNEL.37).
    fn values_view(m: &Machine, sheet: &str, window: Option<&str>) -> String {
        match m.view("values", sheet, window) {
            Ok(Viewed::Text(t)) => t,
            other => panic!("{other:?}"),
        }
    }

    /// A record less its `formula` lines: what the values must be.
    fn less_formulas(record: &str) -> String {
        record
            .lines()
            .filter(|l| !l.starts_with("(formula "))
            .map(|l| format!("{l}\n"))
            .collect()
    }

    fn plane(m: &Machine, sheet: &str, window: &str) -> Vec<u8> {
        match m.view("plane", sheet, Some(window)) {
            Ok(Viewed::Bytes(b)) => b,
            other => panic!("{other:?}"),
        }
    }

    /// A formula cell's computed value, by its sheet and address.
    fn at(m: &Machine, sheet: &str, addr: &str) -> Option<Computed> {
        let i = m.find_sheet(sheet).unwrap();
        let r = parse_a1_range(addr).unwrap();
        m.computed((i, r.top, r.left)).cloned()
    }

    fn number(m: &Machine, sheet: &str, addr: &str) -> f64 {
        match at(m, sheet, addr) {
            Some(Computed::Value(Value::Number(n))) => n,
            other => panic!("{sheet}!{addr}: {other:?}"),
        }
    }

    /// A value cell's content, a twin's included, a twin cell left behind
    /// its slot read as a view reads it (KERNEL.27).
    fn content(m: &Machine, sheet: &str, addr: &str) -> Option<Content> {
        let i = m.find_sheet(sheet).unwrap();
        let r = parse_a1_range(addr).unwrap();
        let held = m.book.sheets[i].cells.get(&(r.top, r.left))?;
        Some(Values::content(m, (i, r.top, r.left)).unwrap_or_else(|| held.content.clone()))
    }

    /// The formula rows of every sheet's record, the twins' included.
    fn formula_rows(m: &Machine) -> Vec<String> {
        let mut out = Vec::new();
        for s in &m.book.sheets {
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
        for s in &m.book.sheets[..m.sheets()] {
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
        assert_eq!(m.book.sheets.len(), 2);
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
        // Every value kept is a formula's of the order, by its position
        // (KERNEL.26), and the order holds no twin's cell.
        assert_eq!(m.values.len(), m.shapes.order().len(), "{when}: the values");
        for &p in m.before.keys() {
            assert!((p as usize) < m.values.len(), "{when}: {p} kept beside");
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

    /// A twin's cell as its cells and its reasons answer: the reference a
    /// twin's array is held to (KERNEL.20), the rule spelled here for itself.
    fn twin_cell(m: &Machine, t: usize, row: u32, col: u32) -> Result<Value, Reason> {
        let held = m.book.sheets[t].cells.get(&(row, col));
        let newer = held.and_then(|_| Values::content(m, (t, row, col)));
        let Some(cell) = newer.or_else(|| held.map(|c| c.content.clone())) else {
            return match m.twin_reasons.get(&(t, row, col)) {
                Some(r) => Err(r.clone()),
                None => Ok(Value::Empty),
            };
        };
        Ok(match &cell {
            Content::Number(n) => Value::Number(*n),
            Content::Text(s) => Value::Text(s.clone()),
            Content::Bool(b) => Value::Bool(*b),
            Content::Error(e) => Value::Error(ErrorKind::parse(e).unwrap_or(ErrorKind::Value)),
            other => panic!("a twin holds {other:?}"),
        })
    }

    /// Whether a sheet's twin is kept as an array (KERNEL.20), asked as the
    /// evaluator asks it, of the twin.
    fn dense_kept(m: &Machine, sheet: &str) -> bool {
        let twin = m.find_sheet(&format!("{sheet}.last")).unwrap();
        m.dense(twin).is_some()
    }

    /// Whether a sheet's twin's array keeps its numbers beside its slots
    /// (KERNEL.43); `None` for a twin kept as no array.
    fn numbers_kept(m: &Machine, sheet: &str) -> Option<bool> {
        let twin = m.find_sheet(&format!("{sheet}.last")).unwrap();
        m.dense(twin).map(Dense::keeps_numbers)
    }

    /// Every twin read and walked by the evaluator held to its cells: each
    /// cell of a window a cell wider than its sheet's extent read both ways
    /// (a twin's own cells and its extent's corners, for an extent too wide
    /// to read whole), and areas inside, across and past the extent walked
    /// both ways, value for value in order; then each sheet named kept as an
    /// array, or not, as said, so that equal answers never hide a twin read
    /// the old way.
    fn dense_is_the_twin(m: &Machine, kept_now: &[(&str, bool)], when: &str) {
        use crate::calc::graph::cells_in;
        use crate::calc::{Area, Ctx};
        use crate::sheet::{MAX_COLUMN, MAX_ROW};
        let ctx = Ctx::new(m, &m.library);
        for s in 0..m.sheets {
            let t = m.sheets + s;
            let name = m.book.sheets[t].name.clone();
            let Some(((top, left), (bottom, right))) = m.extents[s] else {
                continue;
            };
            let (top, left) = (top.saturating_sub(1).max(1), left.saturating_sub(1).max(1));
            let (bottom, right) = ((bottom + 1).min(MAX_ROW), (right + 1).min(MAX_COLUMN));
            let mut cells: Vec<(u32, u32)> = Vec::new();
            if u64::from(bottom - top + 1) * u64::from(right - left + 1) <= 20_000 {
                for row in top..=bottom {
                    cells.extend((left..=right).map(|col| (row, col)));
                }
            } else {
                cells.extend(m.book.sheets[t].cells.keys().copied());
                cells.extend([(top, left), (top, right), (bottom, left), (bottom, right)]);
            }
            for (row, col) in cells {
                assert_eq!(
                    m.read((t, row, col)),
                    twin_cell(m, t, row, col),
                    "{when}: {name} ({row}, {col})"
                );
            }
            let areas = [
                (top, left, bottom, right),
                (1, 1, MAX_ROW, MAX_COLUMN),
                (top + 1, left + 1, top + 1, right),
                (top, right - 1, bottom, right - 1),
                (top + 2, left + 2, bottom + 5, right + 5),
                (bottom, left, bottom.saturating_add(9), left + 9),
                (top + 1, left + 1, top + 1, left + 1),
                (1, left + 1, MAX_ROW, left + 1),
                (top + 1, 1, top + 1, MAX_COLUMN),
            ];
            for (a_top, a_left, a_bottom, a_right) in areas {
                let area = Area {
                    sheet: t,
                    top: a_top.min(MAX_ROW),
                    left: a_left.min(MAX_COLUMN),
                    bottom: a_bottom.min(MAX_ROW),
                    right: a_right.min(MAX_COLUMN),
                };
                let mut walked = Vec::new();
                ctx.each(&[area], |v| {
                    walked.push(v);
                    Ok(())
                })
                .unwrap();
                let want: Vec<Value> = cells_in(&m.book.sheets[t], m.extents[t], &area)
                    .map(|((row, col), _)| twin_cell(m, t, row, col).unwrap())
                    .filter(|v| *v != Value::Empty)
                    .collect();
                assert_eq!(walked, want, "{when}: {name} {area:?}");
            }
        }
        for (sheet, kept) in kept_now {
            assert_eq!(dense_kept(m, sheet), *kept, "{when}: {sheet} kept");
        }
    }

    #[test]
    fn a_twins_array_is_its_cells_after_every_step_and_every_write() {
        // KERNEL.20: each twin's array, read for read and walked for walked,
        // the twin's cells and reasons, over texts, truth values, errors, a
        // cell not computed, a shared formula and a second sheet; after a
        // load, steps, writes of every kind, a refused write, a yield and a
        // frame count set; then over a save loaded again.
        let both = [("Board", true), ("Other", true)];
        let mut m = load(PLANES);
        dense_is_the_twin(&m, &both, "the load");
        for frame in 1..=3 {
            m.step(0);
            dense_is_the_twin(&m, &both, &format!("frame {frame}"));
        }
        for row in [
            "(cell \"Board\" \"B2\" 9)",
            "(formula \"Board\" \"C2\" \"=B2+1\")",
            "(derived \"Board\" \"G2\" 3)",
            // A value cell become a formula not computed: its slot emptied
            // in place, the extent unchanged, its reason beside it.
            "(formula \"Board\" \"E2\" \"=B2:C2\")",
            // A text changed in place, then the extent grown twice.
            "(cell \"Board\" \"D2\" \"a longer text than x\")",
            "(cell \"Board\" \"C5:D6\" 2)",
            "(cell \"Board\" \"I1\" \"far\")",
        ] {
            assert!(m.write(row).is_ok(), "{row}");
            // Until the next frame's end the twin is the last frame's, and
            // so is its array.
            dense_is_the_twin(&m, &both, row);
            m.step(0);
            dense_is_the_twin(&m, &both, row);
        }
        assert!(m
            .write("(cell \"Board\" \"B2\" 1)\n(cell \"Nowhere\" \"A1\" 1)")
            .is_err());
        dense_is_the_twin(&m, &both, "a refused write");
        assert!(!m.step(2).done);
        dense_is_the_twin(&m, &both, "a yield");
        while !m.step(2).done {}
        dense_is_the_twin(&m, &both, "the yielded frame done");
        m.resume_at(50).unwrap();
        m.step(0);
        dense_is_the_twin(&m, &both, "a frame count set");
        let again = load(&save(&m));
        dense_is_the_twin(&again, &both, "a save loaded");
        // Life, whose rule reads its twin in two ranges and three cells.
        let mut life = load(LIFE);
        let kept = [("Screen", true), ("Clock", true)];
        dense_is_the_twin(&life, &kept, "Life's load");
        for frame in 1..=4u64 {
            life_step(&mut life);
            dense_is_the_twin(&life, &kept, &format!("Life's frame {frame}"));
        }
        assert_eq!(plane(&life, "Screen", "A1:BL41"), life_here(4));
    }

    #[test]
    fn a_sheet_too_sparse_for_an_array_is_read_through_its_cells() {
        // KERNEL.20: an array holds at most four slots a cell of its sheet,
        // so two cells far apart keep none, and their twin answers as its
        // cells do; a sheet as dense as that keeps one.
        let mut m = load(concat!(
            "(cell \"Big\" \"A1\" 1)\n",
            "(cell \"Big\" \"XFD1048576\" 2)\n",
            "(formula \"Big\" \"B1\" \"=SUM(Big.last!A:B)+Big.last!XFD1048576\")\n",
            "(cell \"Edge\" \"A1:B1\" 1)\n",
            "(cell \"Edge\" \"D2\" 1)\n",
        ));
        let kept = [("Big", false), ("Edge", true)];
        dense_is_the_twin(&m, &kept, "the load");
        m.step(0);
        assert_eq!(number(&m, "Big", "B1"), 3.0);
        m.step(0);
        assert_eq!(
            number(&m, "Big", "B1"),
            6.0,
            "the twin's B1 read in the range"
        );
        dense_is_the_twin(&m, &kept, "two steps");
        // Edge's extent is A1:D2, eight slots for three cells; a fourth
        // cell far right makes it twenty slots for four, past four a cell.
        assert_eq!(m.write("(cell \"Edge\" \"J2\" 1)"), Ok(1));
        m.step(0);
        dense_is_the_twin(&m, &[("Big", false), ("Edge", false)], "Edge grown");
    }

    #[test]
    fn a_range_over_a_twin_skips_a_cell_not_computed_as_it_did() {
        // Found 2026-10-10 (KERNEL.20), kept bit for bit and named for
        // KERNEL.8: a twin holds no cell for a formula not computed, its
        // reason standing beside, so a read of that cell answers the reason
        // while a range over the twin skips the cell, where a range over the
        // sheet itself answers the reason. The array answers as the cells
        // did, both ways.
        let mut m = load(concat!(
            "(cell \"S\" \"A1:A3\" 1)\n",
            "(cell \"S\" \"B1\" 5)\n",
            "(formula \"S\" \"B2\" \"=A1:A2\")\n",
            "(cell \"S\" \"B3\" 7)\n",
            "(formula \"S\" \"C1\" \"=SUM(S.last!B1:B3)\")\n",
            "(formula \"S\" \"C2\" \"=S.last!B2\")\n",
            "(formula \"S\" \"C3\" \"=SUM(B1:B3)\")\n",
        ));
        m.step(0);
        m.step(0);
        assert!(dense_kept(&m, "S"));
        let reason = Some(Computed::NotComputed(Reason::Construct(
            "A1:A2".to_string(),
        )));
        assert_eq!(at(&m, "S", "B2"), reason);
        assert_eq!(number(&m, "S", "C1"), 12.0, "the twin's range skips B2");
        assert_eq!(at(&m, "S", "C2"), reason, "a read of B2's twin answers");
        assert_eq!(at(&m, "S", "C3"), reason, "a range over the sheet answers");
    }

    #[test]
    fn a_placed_reference_reads_the_sheets_found_once_as_the_names_find_them() {
        // KERNEL.20, slice 2: a placed shape's references to named sheets
        // found at placement, the twins among them, and a frame computed
        // through them equal to `Calc` over the same workbook, which finds
        // every sheet by name: a span, a sheet's own cells, a twin, a sheet
        // the grid does not hold, and one formula text on two sheets.
        use crate::calc::Calc;
        let mut m = load(concat!(
            "(cell \"Data\" \"A1:C3\" 2)\n",
            "(cell \"Jan\" \"A1\" 1)\n",
            "(cell \"Feb\" \"A1\" 10)\n",
            "(cell \"Mar\" \"A1\" 100)\n",
            "(formula \"Sum\" \"A1\" \"=SUM(Jan:Mar!A1)\")\n",
            "(formula \"Sum\" \"A2\" \"=SUM(Data!A1:C3)*Data!B2\")\n",
            "(formula \"Sum\" \"A3\" \"=Nowhere!A1\")\n",
            "(formula \"Sum\" \"B1:B3\" \"=A1+Sum.last!B1\")\n",
            "(formula \"Jan\" \"B1\" \"=Data!A1*2+A1\")\n",
            "(formula \"Feb\" \"B1\" \"=Data!A1*2+A1\")\n",
        ));
        let (cells, shapes) = m.formulas();
        assert_eq!((cells, shapes), (8, 5), "Jan's and Feb's B1 one shape");
        // Every reference naming a sheet the grid holds was found, and to
        // the sheets the names find; one naming none, or a sheet it does not
        // hold, was not.
        let (mut found, mut unfound) = (0, 0);
        for i in 0..shapes {
            let Ok(e) = &m.shapes.shape(i).unwrap().expr else {
                panic!("shape {i} parses")
            };
            e.each_read(&mut |read| {
                let crate::calc::formula::Read::Ref(r) = read else {
                    return;
                };
                let first = r.sheet.as_deref().and_then(|s| m.book.find_sheet(s));
                let last = match &r.last_sheet {
                    Some(s) => m.book.find_sheet(s),
                    None => first,
                };
                let want = first
                    .zip(last)
                    .map(|(a, b)| (u8::try_from(a).unwrap(), u8::try_from(b).unwrap()));
                assert_eq!(r.found, want, "{}", r.written);
                if want.is_some() {
                    found += 1;
                } else {
                    unfound += 1;
                }
            });
        }
        assert_eq!(
            (found, unfound),
            (5, 3),
            "Jan:Mar, Data twice, Sum.last, Data"
        );
        let before = Calc::run(&m.book, &m.library);
        let mut cells: Vec<CellId> = m.shapes.order().to_vec();
        cells.sort_unstable();
        let want: Vec<Option<Computed>> = cells
            .iter()
            .map(|id| before.computed(*id).cloned())
            .collect();
        m.step(0);
        let got: Vec<Option<Computed>> = cells.iter().map(|id| m.computed(*id).cloned()).collect();
        assert_eq!(got, want, "the machine's frame and Calc's, cell for cell");
        assert_eq!(number(&m, "Sum", "A1"), 111.0, "the span");
        assert_eq!(number(&m, "Sum", "A2"), 36.0);
        assert_eq!(
            at(&m, "Sum", "A3"),
            Some(Computed::Value(Value::Error(ErrorKind::Ref))),
            "a sheet the grid does not hold"
        );
        assert_eq!(
            number(&m, "Feb", "B1"),
            14.0,
            "the shared shape on its own sheet"
        );
    }

    #[test]
    fn a_sheet_past_the_256th_is_found_by_name_as_before() {
        // KERNEL.20, slice 2: a found sheet is a byte, so a grid of 130
        // sheets, whose twins stand from the 131st to the 260th, finds the
        // first 256 at placement and the rest by name at every evaluation,
        // the same values either way.
        let mut text = String::new();
        for i in 1..=130 {
            text.push_str(&format!("(cell \"S{i}\" \"A1\" {i})\n"));
        }
        text.push_str("(formula \"S1\" \"B1\" \"=S130!A1+S126.last!A1+S127.last!A1\")\n");
        let mut m = load(&text);
        let Ok(e) = &m.shapes.shape(0).unwrap().expr else {
            panic!("parses")
        };
        let mut found = Vec::new();
        e.each_read(&mut |read| {
            if let crate::calc::formula::Read::Ref(r) = read {
                found.push(r.found);
            }
        });
        assert_eq!(
            found,
            vec![Some((129, 129)), Some((255, 255)), None],
            "S130, S126.last the 256th, S127.last past it"
        );
        m.step(0);
        m.step(0);
        assert_eq!(number(&m, "S1", "B1"), 130.0 + 126.0 + 127.0);
    }

    #[test]
    fn a_strict_call_s_arguments_are_evaluated_in_order_at_every_count() {
        // KERNEL.20, slice 2: the arguments on the stack for up to four and
        // in a vector past them, each evaluated left to right, a call
        // answering the reason of its first argument not computed.
        let mut m = load(concat!(
            "(cell \"S\" \"A1:A2\" 1)\n",
            "(formula \"S\" \"B1\" \"=SUM()\")\n",
            "(formula \"S\" \"B2\" \"=SUM(1)\")\n",
            "(formula \"S\" \"B3\" \"=SUM(1,2)\")\n",
            "(formula \"S\" \"B4\" \"=SUM(1,2,3)\")\n",
            "(formula \"S\" \"B5\" \"=SUM(1,2,3,4)\")\n",
            "(formula \"S\" \"B6\" \"=SUM(1,2,3,4,5)\")\n",
            "(formula \"S\" \"C1\" \"=MOD(7,3)\")\n",
            "(formula \"S\" \"C2\" \"=SUM(A1:A2*1,Z1:Z2*1)\")\n",
            "(formula \"S\" \"C3\" \"=SUM(1,A1:A2*1,Z1:Z2*1)\")\n",
            "(formula \"S\" \"C4\" \"=SUM(1,2,A1:A2*1,Z1:Z2*1)\")\n",
            "(formula \"S\" \"C5\" \"=SUM(1,2,3,4,A1:A2*1,Z1:Z2*1)\")\n",
        ));
        m.step(0);
        for (addr, want) in [
            ("B1", 0.0),
            ("B2", 1.0),
            ("B3", 3.0),
            ("B4", 6.0),
            ("B5", 10.0),
            ("B6", 15.0),
            ("C1", 1.0),
        ] {
            assert_eq!(number(&m, "S", addr), want, "{addr}");
        }
        let first = Some(Computed::NotComputed(Reason::Construct(
            "A1:A2".to_string(),
        )));
        for addr in ["C2", "C3", "C4", "C5"] {
            assert_eq!(at(&m, "S", addr), first, "{addr}: the first not computed");
        }
    }

    /// A twin's cells as its sheet's values make them, spelled here for
    /// itself (KERNEL.20, slice 2): a value cell's content, a formula's value
    /// of the last complete frame as a value cell holds it, nothing for a
    /// formula with no value or one not computed. The last complete frame's
    /// value is what a view reads, the frame in progress's being beside it
    /// while a step yields (slice 4).
    fn twin_from(m: &Machine, s: usize) -> BTreeMap<(u32, u32), Content> {
        m.book.sheets[s]
            .cells
            .iter()
            .filter_map(|(&k, c)| {
                if !is_formula(&c.content) {
                    return Some((k, c.content.clone()));
                }
                let content = match m.computed((s, k.0, k.1)) {
                    Some(Computed::Value(Value::Number(n))) => Content::Number(*n),
                    Some(Computed::Value(Value::Text(t))) => Content::Text(t.clone()),
                    Some(Computed::Value(Value::Bool(b))) => Content::Bool(*b),
                    Some(Computed::Value(Value::Error(e))) => Content::Error(e.text().to_string()),
                    Some(Computed::Value(Value::Empty)) => Content::Number(0.0),
                    Some(Computed::NotComputed(_)) | None => return None,
                };
                Some((k, content))
            })
            .collect()
    }

    /// A twin's cells as a reader of the twin as a sheet reads them, a cell
    /// a frame's end left behind its slot through its slot (KERNEL.27), as a
    /// view reads it; nothing is made.
    fn twin_cells(m: &Machine, t: usize) -> BTreeMap<(u32, u32), Content> {
        m.book.sheets[t]
            .cells
            .iter()
            .map(|(&k, c)| {
                let content = Values::content(m, (t, k.0, k.1));
                (k, content.unwrap_or_else(|| c.content.clone()))
            })
            .collect()
    }

    /// Every twin's cells, and its extent, as its sheet's values make them:
    /// held at a frame's end or a load, before any write.
    fn twins_are_their_sheets(m: &Machine, when: &str) {
        for s in 0..m.sheets {
            let t = m.sheets + s;
            let twin = &m.book.sheets[t];
            assert_eq!(twin_cells(m, t), twin_from(m, s), "{when}: {}", twin.name);
            assert_eq!(
                m.extents[t],
                twin.extent(),
                "{when}: {}'s extent",
                twin.name
            );
            assert!(twin.cells.values().all(|c| c.style == 0), "{when}");
        }
    }

    #[test]
    fn a_twin_overwritten_in_place_is_its_sheet_after_every_step() {
        // KERNEL.20, slice 2: the twin's cells overwritten in the snapshot's
        // walk while they stand where the sheet's valued cells stand, and
        // made again when a write changed which cells hold a value: a value
        // turned to a formula not computed, a text changed in place, cells
        // added, the last valued cell losing its value, the extent grown.
        let mut m = load(PLANES);
        twins_are_their_sheets(&m, "the load");
        for frame in 1..=3 {
            m.step(0);
            twins_are_their_sheets(&m, &format!("frame {frame}"));
        }
        for row in [
            "(cell \"Board\" \"B2\" 9)",
            "(formula \"Board\" \"E2\" \"=B2:C2\")",
            "(cell \"Board\" \"D2\" \"a longer text than x\")",
            "(cell \"Board\" \"C5:D6\" 2)",
            "(formula \"Board\" \"D6\" \"=B2:C2\")",
            // One cell losing its value and another gaining one in the same
            // write: as many valued cells as before, standing elsewhere.
            "(formula \"Board\" \"F2\" \"=B2:C2\")\n(cell \"Board\" \"H3\" 5)",
            "(cell \"Board\" \"I1\" \"far\")",
            "(cell \"Other\" \"A1\" 4)",
        ] {
            assert!(m.write(row).is_ok(), "{row}");
            m.step(0);
            twins_are_their_sheets(&m, row);
        }
        assert!(m
            .write("(cell \"Board\" \"B2\" 1)\n(cell \"Nowhere\" \"A1\" 1)")
            .is_err());
        assert!(!m.step(2).done);
        twins_are_their_sheets(&m, "a yield");
        while !m.step(2).done {}
        twins_are_their_sheets(&m, "the yielded frame done");
        let mut life = load(LIFE);
        for frame in 1..=3u64 {
            life_step(&mut life);
            twins_are_their_sheets(&life, &format!("Life's frame {frame}"));
        }
    }

    /// splitmix64, the generated grids' only randomness, seeded so that every
    /// run writes the same grids, as `tests/laws.rs` writes its own.
    struct Dice(u64);

    impl Dice {
        fn next(&mut self) -> u64 {
            self.0 = self.0.wrapping_add(0x9E37_79B9_7F4A_7C15);
            let mut z = self.0;
            z = (z ^ (z >> 30)).wrapping_mul(0xBF58_476D_1CE4_E5B9);
            z = (z ^ (z >> 27)).wrapping_mul(0x94D0_49BB_1331_11EB);
            z ^ (z >> 31)
        }

        fn below(&mut self, n: usize) -> usize {
            (self.next() % n as u64) as usize
        }

        fn pick<'a>(&mut self, xs: &[&'a str]) -> &'a str {
            xs[self.below(xs.len())]
        }
    }

    /// Where a generated formula stands: the sheet before its own, of which
    /// it may read any cell, the first row of its band and the rows the
    /// formulas fill, and the column left of its first.
    struct Band<'a> {
        before: Option<&'a str>,
        top: u32,
        rows: u32,
        left: &'a str,
    }

    const LETTERS: [&str; 11] = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K"];

    /// A cell with `$` marks of any kind.
    fn marked(d: &mut Dice, col: &str, row: u32) -> String {
        match d.below(4) {
            0 => format!("${col}${row}"),
            1 => format!("${col}{row}"),
            2 => format!("{col}${row}"),
            _ => format!("{col}{row}"),
        }
    }

    /// A reference to one cell a formula of the band may read without a
    /// cycle, every one pointing before the cell in reading order or off its
    /// sheet's formulas: a value cell, its column held; any cell of a twin or
    /// of the sheet before; a formula cell above the band, which a cell lower
    /// in the band reads in the band itself; the cell to its left, so that a
    /// run reads itself and the plan splits it; a cell on the sheet's last
    /// column or row, which a cell to its right or below moves off the sheet;
    /// a span of both sheets, which makes its shape risky and leaves its
    /// cells alone (a span moved off the sheet is refused at the load, so
    /// none is written).
    fn reference(d: &mut Dice, at: &Band) -> String {
        let (col, row) = (d.pick(&LETTERS), 1 + d.below(at.rows as usize) as u32);
        match d.below(9) {
            0 => {
                let held = if d.below(2) == 0 { "$" } else { "" };
                format!("${}{held}{}", d.pick(&["A", "B"]), 1 + d.below(20))
            }
            1 | 2 => {
                let twin = d.pick(&["A", "B"]);
                format!("{twin}.last!{}", marked(d, col, row))
            }
            3 if at.top > 1 => {
                let above = 1 + d.below(at.top as usize - 1) as u32;
                let col = d.pick(&LETTERS[2..]);
                marked(d, col, above)
            }
            4 => format!("{}{}", at.left, at.top),
            5 => d.pick(&["XFD1", "A1048576", "A.last!XFD2"]).to_string(),
            6 if d.below(8) == 0 => "A:B!$A3".to_string(),
            7 => held(d, at),
            _ => match at.before {
                Some(s) => format!("{s}!{}", marked(d, col, row)),
                None => format!("$A{}", 1 + d.below(20)),
            },
        }
    }

    /// A reference every `$` holds, to a cell whose formula may not have
    /// been computed: a twin's, or a formula cell above the band, so that a
    /// reason may be the same at every cell of a run.
    fn held(d: &mut Dice, at: &Band) -> String {
        let col = d.pick(&LETTERS[2..]);
        if at.top > 1 && d.below(2) == 0 {
            format!("${col}${}", 1 + d.below(at.top as usize - 1))
        } else {
            let twin = d.pick(&["A", "B"]);
            format!("{twin}.last!${col}${}", 1 + d.below(at.rows as usize))
        }
    }

    /// A range the band may read: a twin's, the value columns, formula cells
    /// above the band, or the sheet before's cells, its formulas among them,
    /// each corner held any way (KERNEL.33: a plan finds a range every `$`
    /// holds once for the shape, and one held only in part at each cell).
    fn range(d: &mut Dice, at: &Band) -> String {
        let (r1, c1) = (1 + d.below(at.rows as usize) as u32, d.below(9));
        let (h, w) = (d.below(3) as u32, d.below(3));
        match d.below(4) {
            0 => format!("$A{r1}:$B{}", r1 + h),
            1 if at.top > 2 => {
                let r = 1 + d.below(at.top as usize - 2) as u32;
                format!("C{r}:{}{}", LETTERS[2 + w], (r + h).min(at.top - 1))
            }
            3 if at.before.is_some() => {
                let first = marked(d, LETTERS[c1], r1);
                let last = marked(d, LETTERS[c1 + w], r1 + h);
                format!("{}!{first}:{last}", at.before.unwrap_or("A"))
            }
            _ => {
                let twin = d.pick(&["A", "B"]);
                let first = marked(d, LETTERS[c1], r1);
                format!("{twin}.last!{first}:{}{}", LETTERS[c1 + w], r1 + h)
            }
        }
    }

    fn leaf(d: &mut Dice, at: &Band) -> String {
        match d.below(10) {
            0 => d
                .pick(&["0", "1", "2", "3", "0.5", "2.5", "10", "1E+300"])
                .to_string(),
            1 => d
                .pick(&["\"x\"", "\"5\"", "\"TRUE\"", "\"\"", "\" 2 \""])
                .to_string(),
            2 => d.pick(&["TRUE", "FALSE", "#DIV/0!", "#N/A"]).to_string(),
            // A name over formula cells is read from the sheet after theirs
            // alone, so that no grid holds a cycle (KERNEL.33).
            3 if at.before.is_some() => d.pick(&["Rate", "Box", "Prev", "Band"]).to_string(),
            3 => d.pick(&["Rate", "Box", "Prev"]).to_string(),
            _ => reference(d, at),
        }
    }

    /// A formula of the language's functions over every kind of node the
    /// evaluation of a run computes, and of those it leaves to each cell: a
    /// name, a range where one value is wanted, a call whose operand may be
    /// an area; a part written more than once, which a run computes once
    /// where its places stand in one branch (KERNEL.27); and the edges of the
    /// functions a run computes by their meanings (KERNEL.28): `CHOOSE`'s
    /// index at each end of its count, `MOD` by 0 and of each sign, `ROW` and
    /// `COLUMN` of a reference across sheets or off the sheet, `CHOOSE`
    /// beside and inside the run's other calls.
    fn expr(d: &mut Dice, at: &Band, depth: u32) -> String {
        if depth >= 3 || d.below(4) == 0 {
            return leaf(d, at);
        }
        let e = |d: &mut Dice| expr(d, at, depth + 1);
        let operands = |d: &mut Dice, n: usize| {
            let parts: Vec<String> = (0..n)
                .map(|_| {
                    if d.below(3) == 0 {
                        range(d, at)
                    } else {
                        expr(d, at, depth + 1)
                    }
                })
                .collect();
            parts.join(",")
        };
        match d.below(24) {
            0 => format!("-({})", e(d)),
            1 => format!("({})%", e(d)),
            2 => format!("+{}", reference(d, at)),
            3 => {
                let ops = [
                    "+", "-", "*", "/", "^", "&", "=", "<>", "<", ">", "<=", ">=",
                ];
                let op = d.pick(&ops);
                format!("({}){op}({})", e(d), e(d))
            }
            4 => format!("({})+({})-({})*({})", e(d), e(d), e(d), e(d)),
            5 => format!("IF({},{},{})", e(d), e(d), e(d)),
            6 => match d.below(3) {
                0 => format!("IF({},{})", e(d), e(d)),
                1 => format!("IF({},,{})", e(d), e(d)),
                _ => format!("IF({},{},)", e(d), e(d)),
            },
            7 => {
                let f = d.pick(&["AND", "OR"]);
                let n = 1 + d.below(3);
                format!("{f}({})", operands(d, n))
            }
            8 => format!("NOT({})", e(d)),
            9 => {
                let f = d.pick(&["SUM", "MIN", "MAX"]);
                let n = d.below(5);
                format!("{f}({})", operands(d, n))
            }
            10 => format!("{}({})", d.pick(&["ABS", "INT", "SIN"]), e(d)),
            11 => format!("MOD({},{})", e(d), e(d)),
            12 => match d.below(3) {
                0 => d.pick(&["ROW()", "COLUMN()"]).to_string(),
                1 => format!("ROW({})", reference(d, at)),
                _ => format!("COLUMN({})", range(d, at)),
            },
            13 => format!("CHOOSE({},{},{},{})", e(d), e(d), e(d), e(d)),
            14 => format!("SUM(IF({},{},{}))", e(d), range(d, at), range(d, at)),
            15 => format!("({})+1", range(d, at)),
            16 => {
                let op = d.pick(&["+", "&", "<"]);
                format!("{}{op}({})", held(d, at), e(d))
            }
            // A part written more than once: beside itself in a chain, as a
            // call's operands, in a condition and its branches, in both
            // branches of an `IF`, and inside a part itself written twice.
            18 => {
                let p = e(d);
                match d.below(5) {
                    0 => format!("({p})+({p})*({p})"),
                    1 => format!("{}(({p}),({p}))", d.pick(&["SUM", "MAX", "AND", "MOD"])),
                    2 => format!("IF({p},{p},({p})-1)"),
                    3 => format!("IF({},({p})*2,({p})&\"x\")", e(d)),
                    _ => format!("NOT({p})+({p})"),
                }
            }
            19 => {
                let q = e(d);
                let p = format!("({q})-({q})*2");
                format!("({p})+({p})")
            }
            // CHOOSE of two to six values, its index at each end of its count,
            // of each kind, or the row's, whose cells choose every value and
            // one past the last (KERNEL.28).
            20 => {
                let n = 2 + d.below(5);
                let index = match d.below(4) {
                    0 => d
                        .pick(&["0.5", "1", "-1", "\"2\"", "TRUE", "1E+300"])
                        .to_string(),
                    1 => format!("{n}{}", d.pick(&["", ".9", "+1"])),
                    2 => format!("MOD(ROW(),{})+1", n + 1),
                    _ => e(d),
                };
                let values: Vec<String> = (0..n).map(|_| e(d)).collect();
                format!("CHOOSE({index},{})", values.join(","))
            }
            // MOD by 0, and of each sign.
            21 => match d.below(3) {
                0 => format!("MOD({},0)", e(d)),
                1 => format!("MOD(-({}),{})", e(d), e(d)),
                _ => format!(
                    "MOD(ROW()-{},{})",
                    1 + d.below(5),
                    d.pick(&["-3", "2", "-0.5", "ROW()-3"])
                ),
            },
            // ROW and COLUMN of a reference to another sheet, a twin, a span,
            // a sheet the grid does not hold, or moved off the sheet.
            22 => {
                let r = match d.below(3) {
                    0 => reference(d, at),
                    1 => range(d, at),
                    _ => d
                        .pick(&[
                            "A1048576",
                            "XFD1",
                            "A:B!$A3",
                            "Gone!A1",
                            "$A:$B",
                            "A.last!3:3",
                        ])
                        .to_string(),
                };
                format!("{}({r})", d.pick(&["ROW", "COLUMN"]))
            }
            // CHOOSE an operand of the calls the run computes, inside IF, and
            // around IF.
            23 => match d.below(3) {
                0 => format!(
                    "{}(CHOOSE(MOD(ROW(),2)+1,{},{}),{})",
                    d.pick(&["MOD", "SUM", "MAX", "ABS"]),
                    e(d),
                    e(d),
                    e(d)
                ),
                1 => format!(
                    "IF({},CHOOSE(MOD(ROW(),3)+1,{},{},{}),{})",
                    e(d),
                    e(d),
                    e(d),
                    e(d),
                    e(d)
                ),
                _ => format!(
                    "CHOOSE(MOD(ROW(),2)+1,IF({},{},{}),{})",
                    e(d),
                    e(d),
                    e(d),
                    e(d)
                ),
            },
            _ => format!("{}*{}", reference(d, at), reference(d, at)),
        }
    }

    /// A value cell's datum, or nothing for an empty cell.
    fn datum(d: &mut Dice) -> Option<&'static str> {
        let all = [
            "0",
            "1",
            "-2.5",
            "3",
            "\"5\"",
            "\"x\"",
            "\"TRUE\"",
            "true",
            "(error \"#N/A\")",
        ];
        (d.below(4) != 0).then(|| d.pick(&all))
    }

    /// A grid of two sheets, each with value cells in A1:B20 and formulas
    /// over C1:H`rows` in bands, each band one shape; in some grids a column
    /// of single formulas beside, which interleaves the order and leaves each
    /// run a row of a band; and four defined names, one over the first
    /// sheet's formulas (KERNEL.33).
    fn generated_grid(d: &mut Dice) -> String {
        grid_of(d, false)
    }

    /// The grid `generated_grid` writes from the same throws, three bands in
    /// four and every other single formula adding the Clock's `B1`, which a
    /// test writes before every frame, so that most of its cells change at
    /// every frame (KERNEL.36): a number changing, a text or an error not.
    fn churning_grid(d: &mut Dice) -> String {
        grid_of(d, true)
    }

    /// `generated_grid`, or `churning_grid` when `churn`.
    fn grid_of(d: &mut Dice, churn: bool) -> String {
        let rows = 4 + d.below(60) as u32;
        let singles = d.below(3) == 0;
        let mut out = String::new();
        let mut band = 0;
        let churned = |f: String, yes: bool| {
            if churn && yes {
                format!("({f})+Clock!$B$1")
            } else {
                f
            }
        };
        for (sheet, before) in [("A", None), ("B", Some("A"))] {
            out.push_str(&format!("(sheet \"{sheet}\" visible)\n"));
            for row in 1..=20 {
                for col in ["A", "B"] {
                    if let Some(v) = datum(d) {
                        out.push_str(&format!("(cell \"{sheet}\" \"{col}{row}\" {v})\n"));
                    }
                }
            }
            let mut top = 1;
            while top <= rows {
                let bottom = (top + d.below(rows as usize) as u32).min(rows);
                let at = Band {
                    before,
                    top,
                    rows,
                    left: "B",
                };
                band += 1;
                let f = churned(expr(d, &at, 0), band % 4 != 0);
                let f = crate::rows::quoted(&format!("={f}"));
                out.push_str(&format!("(formula \"{sheet}\" \"C{top}:H{bottom}\" {f})\n"));
                top = bottom + 1;
            }
            for row in (1..=rows).filter(|_| singles) {
                let at = Band {
                    before,
                    top: row,
                    rows,
                    left: "J",
                };
                let f = churned(expr(d, &at, 0), row % 2 == 1);
                let f = crate::rows::quoted(&format!("={f}"));
                out.push_str(&format!("(formula \"{sheet}\" \"K{row}\" {f})\n"));
            }
        }
        if churn {
            out.push_str("(sheet \"Clock\" visible)\n(cell \"Clock\" \"B1\" 0)\n");
        }
        out.push_str("(name \"Rate\" \"A!$A$2\")\n");
        out.push_str("(name \"Box\" \"A!$A$1:$B$3\")\n");
        out.push_str("(name \"Prev\" \"A.last!$C$1\")\n");
        out.push_str("(name \"Band\" \"A!$C$1:$D$2\")\n");
        out
    }

    /// Every formula cell's value of the last complete frame, spelled to the
    /// bit (a negative zero, too).
    fn values(m: &Machine) -> BTreeMap<CellId, String> {
        m.shapes
            .order()
            .iter()
            .zip(&m.values)
            .filter_map(|(id, c)| Some((*id, format!("{:?}", c.as_ref()?))))
            .collect()
    }

    #[test]
    fn a_run_evaluated_at_once_is_its_cells_evaluated_one_by_one() {
        // KERNEL.20, slice 3: grids of generated formulas, every kind of node,
        // stepped by runs in slices of any budget, against the same grids
        // with every cell evaluated alone, frame for frame and bit for bit;
        // writes between frames change values, break a band's shape and make
        // the runs again.
        let (mut in_runs, mut by_cell) = (0, 0);
        for seed in 1..=80u64 {
            let mut d = Dice(seed);
            let rows = generated_grid(&mut d);
            let mut runs = Machine::load(&rows, library::language()).unwrap_or_else(|e| {
                panic!("seed {seed}: line {}: {}\n{rows}", e.line, e.refusal.text)
            });
            let mut alone = load(&rows);
            alone.by_runs = false;
            // The first frame also as `Calc` computes it, which plans its own
            // order and parses every cell's own text: before it no twin holds
            // a reason, which `Calc`, reading a twin as a sheet, would not see.
            let calc: BTreeMap<CellId, String> = {
                let calc = crate::calc::Calc::run(&runs.book, &runs.library);
                let spell = |id: &CellId| format!("{:?}", calc.computed(*id).expect("computed"));
                runs.shapes
                    .order()
                    .iter()
                    .map(|id| (*id, spell(id)))
                    .collect()
            };
            for frame in 1..=4 {
                in_runs += runs.shapes.cells_in_runs();
                by_cell += runs.shapes.order().len();
                while !runs.step(1 + d.below(300)).done {}
                assert!(alone.step(0).done);
                assert_eq!(
                    values(&runs),
                    values(&alone),
                    "seed {seed}, frame {frame}\n{rows}"
                );
                if frame == 1 {
                    assert_eq!(values(&runs), calc, "seed {seed}: Calc\n{rows}");
                }
                let write = match d.below(4) {
                    0 => format!("(cell \"A\" \"A{}\" 7)", 1 + d.below(20)),
                    1 => format!("(cell \"B\" \"B{}\" \"y\")", 1 + d.below(20)),
                    2 => format!(
                        "(cell \"A\" \"{}{}\" 2)",
                        d.pick(&LETTERS[2..8]),
                        1 + d.below(4)
                    ),
                    _ => continue,
                };
                let (a, b) = (runs.write(&write), alone.write(&write));
                assert!(a.is_ok() && b.is_ok(), "seed {seed}: {write}: {a:?} {b:?}");
            }
        }
        assert!(
            in_runs * 2 > by_cell,
            "{in_runs} of {by_cell} cells were in runs: too few for the law to hold over"
        );
    }

    // ---- KERNEL.43: a run's lanes typed end to end ----

    /// A cell of one of the typed grids' twins, up to two rows and four
    /// columns from the formula's own, its column held one time in five and
    /// its row one in three.
    fn typed_ref(d: &mut Dice) -> String {
        let twin = d.pick(&["S.last", "T.last"]);
        let col = if d.below(5) == 0 { "$" } else { "" };
        let row = if d.below(3) == 0 { "$" } else { "" };
        let letter = d.pick(&["A", "B", "C", "D", "E"]);
        format!("{twin}!{col}{letter}{row}{}", 1 + d.below(5))
    }

    /// A range of a typed grid's twin, up to three rows by four columns,
    /// moving with the cell at both corners, held at both, or held at one.
    fn typed_range(d: &mut Dice) -> String {
        let twin = d.pick(&["S.last", "T.last"]);
        let (c1, r1) = (d.below(5), 1 + d.below(5) as u32);
        let (h, w) = (d.below(3) as u32, d.below(4));
        let (a, b) = match d.below(6) {
            0 => ("$", "$"),
            1 => ("$", ""),
            2 => ("", "$"),
            _ => ("", ""),
        };
        format!(
            "{twin}!{a}{}{r1}:{b}{}{}",
            LETTERS[c1],
            LETTERS[c1 + w],
            r1 + h
        )
    }

    /// A comparison over the typed grids' nodes: a lane of truth values.
    fn typed_test(d: &mut Dice, depth: u32) -> String {
        let op = d.pick(&["=", "<>", "<", ">", "<=", ">="]);
        format!(
            "({}){op}({})",
            typed_expr(d, depth + 1),
            typed_expr(d, depth + 1)
        )
    }

    /// A formula of the nodes a run computes by typed lanes (KERNEL.43): the
    /// twins' cells and ranges read by slices or, held, once a row; the
    /// aggregates and the truth functions over values, comparisons and
    /// ranges; `NOT`; `IF` between literals of one kind or of two, and
    /// between any branches.
    fn typed_expr(d: &mut Dice, depth: u32) -> String {
        if depth >= 3 || d.below(3) == 0 {
            return match d.below(6) {
                0 => d.pick(&["0", "1", "2", "3", "0.5", "-1"]).to_string(),
                _ => typed_ref(d),
            };
        }
        let e = |d: &mut Dice| typed_expr(d, depth + 1);
        let operands = |d: &mut Dice, n: usize| {
            let parts: Vec<String> = (0..n)
                .map(|_| match d.below(3) {
                    0 => typed_range(d),
                    1 => typed_test(d, depth + 1),
                    _ => typed_expr(d, depth + 1),
                })
                .collect();
            parts.join(",")
        };
        match d.below(11) {
            0 => format!("({})+({})", e(d), e(d)),
            1 => format!("({})-({})*({})", e(d), e(d), e(d)),
            2 => {
                let f = d.pick(&["SUM", "MIN", "MAX"]);
                let n = 1 + d.below(3);
                format!("{f}({})", operands(d, n))
            }
            3 => format!("SUM({})-({})", typed_range(d), typed_ref(d)),
            4 => {
                let f = d.pick(&["AND", "OR"]);
                let n = 1 + d.below(3);
                format!("{f}({})", operands(d, n))
            }
            5 => format!("NOT({})", typed_test(d, depth)),
            6 => {
                let (a, b) = match d.below(3) {
                    0 => ("1", "0"),
                    1 => ("TRUE", "FALSE"),
                    _ => ("2.5", "TRUE"),
                };
                format!("IF({},{a},{b})", typed_test(d, depth))
            }
            7 => format!("IF({},{},{})", typed_test(d, depth), e(d), e(d)),
            8 => {
                let (a, b) = (d.pick(&["1", "7"]), d.pick(&["0", "-2"]));
                format!("IF({},{a},{b})", e(d))
            }
            9 => match d.below(2) {
                0 => typed_test(d, depth),
                _ => format!("NOT({})", e(d)),
            },
            _ => format!("MOD({},5)", e(d)),
        }
    }

    /// A grid of KERNEL.43's law: on `S`, value cells in `A1:B<rows>` and
    /// bands of formulas from `C1` to a column 20 to 300 wide, each band one
    /// shape, every formula folded into a number by `MOD`, so that its twin
    /// holds numbers but where a formula answers an error; one band in four
    /// with a hole, a cell of its own formula, so that its runs stop beside
    /// it; on `T`, numbers over the same columns, a few texts, truth values,
    /// errors and empty cells among them. The rows and the last column.
    fn typed_grid(d: &mut Dice) -> (String, u32, u32) {
        let rows = 2 + d.below(10) as u32;
        let last = 3 + 20 + d.below(281) as u32;
        let right = crate::sheet::column_letters(last);
        let mut out = String::from("(sheet \"S\" visible)\n(sheet \"T\" visible)\n");
        for row in 1..=rows {
            for col in ["A", "B"] {
                let v = d.pick(&["0", "1", "2", "-3", "0.5", "1", "2"]);
                out.push_str(&format!("(cell \"S\" \"{col}{row}\" {v})\n"));
            }
        }
        // T's cells, each written once: a number; and in one row of T in
        // four, now and then another kind, or nothing, so that a lane read
        // over those rows meets a value that is not a number and a lane over
        // the others does not.
        for row in 1..=rows + 2 {
            let odd = d.below(4) == 0;
            for col in 1..=last {
                let v = match d.below(if odd { 10 } else { 1000 }) {
                    0 => "\"x\"",
                    1 => "true",
                    2 => "(error \"#N/A\")",
                    3 => continue,
                    _ => d.pick(&["1", "2", "-1", "0.25", "3", "0"]),
                };
                let at = crate::sheet::cell_ref(row, col);
                out.push_str(&format!("(cell \"T\" \"{at}\" {v})\n"));
            }
        }
        let column = crate::sheet::column_letters;
        let mut top = 1;
        while top <= rows {
            let bottom = (top + d.below(rows as usize) as u32).min(rows);
            let f = crate::rows::quoted(&format!("=MOD({},97)", typed_expr(d, 0)));
            if d.below(4) == 0 {
                // A hole in the band's first row, a formula of its own.
                let hole = 4 + d.below(last as usize - 4) as u32;
                let (before, after) = (column(hole - 1), column(hole + 1));
                out.push_str(&format!("(formula \"S\" \"C{top}:{before}{top}\" {f})\n"));
                out.push_str(&format!(
                    "(formula \"S\" \"{}{top}\" \"=S.last!A1+1\")\n",
                    column(hole)
                ));
                out.push_str(&format!(
                    "(formula \"S\" \"{after}{top}:{right}{top}\" {f})\n"
                ));
                if bottom > top {
                    out.push_str(&format!(
                        "(formula \"S\" \"C{}:{right}{bottom}\" {f})\n",
                        top + 1
                    ));
                }
            } else {
                out.push_str(&format!("(formula \"S\" \"C{top}:{right}{bottom}\" {f})\n"));
            }
            top = bottom + 1;
        }
        out.push_str("(sheet \"Clock\" visible)\n(cell \"Clock\" \"B1\" 0)\n");
        (out, rows, last)
    }

    #[test]
    fn cells_that_read_past_their_twin_s_array_read_inside_it_by_a_slice() {
        // KERNEL.43: of two fills over a twin's array, columns C to Z, the
        // one whose last two cells read right of it reads a slice to its
        // edge and those two one by one; the one whose first two read left
        // of it reads every cell one by one, a cell outside the array being
        // empty, which a lane of numbers never holds. Both are the cells
        // evaluated one by one.
        let rows = concat!(
            "(sheet \"U\" visible)\n",
            "(cell \"U\" \"C1:Z1\" 2)\n",
            "(formula \"U\" \"C2:Z2\" \"=U.last!A1+COLUMN()\")\n",
            "(formula \"U\" \"C3:Z3\" \"=U.last!E1+COLUMN()\")\n",
        );
        let (mut runs, mut alone) = (load(rows), load(rows));
        alone.by_runs = false;
        vector::sliced::take();
        assert!(runs.step(0).done && alone.step(0).done);
        assert!(dense_kept(&runs, "U"), "the twin is kept as an array");
        assert_eq!(values(&runs), values(&alone));
        assert_eq!(
            vector::sliced::take()[vector::sliced::READ],
            1,
            "the right edge's fill alone reads a slice"
        );
    }

    #[test]
    fn a_twin_no_formula_reads_keeps_no_numbers_beside_its_slots() {
        // KERNEL.43: a spreadsheet's sheet, whose twin no formula reads,
        // keeps its twin's array without the numbers beside its slots; a
        // write that makes a formula read the twin makes them at once, from
        // the slots, and the frames after it read them by slices; a write
        // that leaves no formula reading it drops them. Every frame is the
        // cells evaluated one by one.
        let rows = concat!(
            "(sheet \"U\" visible)\n",
            "(cell \"U\" \"C1:Z1\" 2)\n",
            "(formula \"U\" \"C2:Z2\" \"=U!C1+COLUMN()\")\n",
        );
        let (mut runs, mut alone) = (load(rows), load(rows));
        alone.by_runs = false;
        assert!(runs.step(0).done && alone.step(0).done);
        assert_eq!(
            numbers_kept(&runs, "U"),
            Some(false),
            "no formula reads U.last"
        );
        for m in [&mut runs, &mut alone] {
            assert_eq!(
                m.write("(formula \"U\" \"C3:Z3\" \"=U.last!C2*2\")"),
                Ok(24)
            );
        }
        assert_eq!(numbers_kept(&runs, "U"), Some(true), "made at the write");
        vector::sliced::take();
        for frame in 2..=3 {
            assert!(runs.step(0).done && alone.step(0).done);
            assert_eq!(values(&runs), values(&alone), "frame {frame}");
            assert_eq!(numbers_kept(&runs, "U"), Some(true), "frame {frame}");
        }
        assert!(
            vector::sliced::take()[vector::sliced::READ] > 0,
            "the twin read by slices"
        );
        for m in [&mut runs, &mut alone] {
            assert_eq!(m.write("(formula \"U\" \"C3:Z3\" \"=U!C2*2\")"), Ok(24));
        }
        assert_eq!(
            numbers_kept(&runs, "U"),
            Some(false),
            "dropped at the write"
        );
        assert!(runs.step(0).done && alone.step(0).done);
        assert_eq!(values(&runs), values(&alone));
        assert_eq!(numbers_kept(&runs, "U"), Some(false));
        // A shape whose reads the plan cannot follow, a name's, might read
        // any twin: while one stands, every twin keeps its numbers.
        let named = concat!(
            "(sheet \"U\" visible)\n",
            "(name \"Rate\" \"U!$C$1\")\n",
            "(cell \"U\" \"C1:Z1\" 2)\n",
            "(formula \"U\" \"C2:Z2\" \"=U!C1+COLUMN()\")\n",
            "(formula \"U\" \"C3\" \"=Rate*2\")\n",
        );
        let mut m = load(named);
        assert!(m.step(0).done);
        assert_eq!(numbers_kept(&m, "U"), Some(true), "a name's shape stands");
    }

    #[test]
    fn a_run_read_by_slices_is_its_cells_evaluated_one_by_one() {
        // KERNEL.43: grids of wide bands whose formulas read their twins by
        // slices and fold their ranges by slices, stepped by runs in slices of
        // any budget, against the same grids with every cell evaluated alone
        // and against `Calc` at the first frame, frame for frame and bit for
        // bit; between frames, writes into the twins' sheets change a value,
        // turn a number into a text and back, and empty a cell, so that a
        // frame reaches some cells and not others, its stretches broken where
        // a cell is not reached, and a slice meets a slot that is not a
        // number. Every path the item adds is taken.
        let mut taken = [0usize; 4];
        for seed in 1..=40u64 {
            let mut d = Dice(seed);
            let (rows, height, last) = typed_grid(&mut d);
            let mut runs = load(&rows);
            let mut alone = load(&rows);
            alone.by_runs = false;
            let calc: BTreeMap<CellId, String> = {
                let calc = crate::calc::Calc::run(&runs.book, &runs.library);
                let spell = |id: &CellId| format!("{:?}", calc.computed(*id).expect("computed"));
                runs.shapes
                    .order()
                    .iter()
                    .map(|id| (*id, spell(id)))
                    .collect()
            };
            vector::sliced::take();
            for frame in 1..=6 {
                while !runs.step(1 + d.below(2000)).done {}
                assert!(alone.step(0).done);
                assert_eq!(
                    values(&runs),
                    values(&alone),
                    "seed {seed}, frame {frame}\n{rows}"
                );
                if frame == 1 {
                    assert_eq!(values(&runs), calc, "seed {seed}: Calc\n{rows}");
                }
                let at = |d: &mut Dice| {
                    crate::sheet::cell_ref(
                        1 + d.below(height as usize + 2) as u32,
                        1 + d.below(last as usize) as u32,
                    )
                };
                let write = match d.below(5) {
                    0 => format!(
                        "(cell \"T\" \"{}\" {})",
                        at(&mut d),
                        d.pick(&["4", "-2", "0"])
                    ),
                    1 => format!("(cell \"T\" \"{}\" \"y\")", at(&mut d)),
                    2 => format!(
                        "(cell \"S\" \"A{}\" {})",
                        1 + d.below(height as usize),
                        d.pick(&["3", "\"z\""])
                    ),
                    3 => format!("(cell \"Clock\" \"B1\" {frame})"),
                    _ => continue,
                };
                let (a, b) = (runs.write(&write), alone.write(&write));
                assert!(a.is_ok() && b.is_ok(), "seed {seed}: {write}: {a:?} {b:?}");
            }
            let counts = vector::sliced::take();
            for (t, c) in taken.iter_mut().zip(counts) {
                *t += c;
            }
        }
        assert!(
            taken.iter().all(|&c| c > 0),
            "every path taken somewhere: reads by slices, held columns, folds, choices {taken:?}"
        );
    }

    // ---- KERNEL.30: the census held to the plan walked again, and to the
    // paths runs take ----

    /// A machine's census, its twins' arrays as it keeps them and its own
    /// sheets' as the plan keeps them (KERNEL.31), as `Machine::census`.
    fn census_of(m: &Machine) -> crate::calc::census::Census {
        crate::calc::census::of(&m.shapes, &m.book, &m.library, &|s| {
            if s < m.sheets {
                m.own[s].is_some()
            } else {
                Env::dense(m, s).is_some()
            }
        })
    }

    /// The same census with no sheet of the rows kept as an array: what a
    /// run reads while a sheet steps its array aside, one of its formulas
    /// holding a reason (KERNEL.31).
    fn census_unkept(m: &Machine) -> crate::calc::census::Census {
        crate::calc::census::of(&m.shapes, &m.book, &m.library, &|s| {
            s >= m.sheets && Env::dense(m, s).is_some()
        })
    }

    /// Each formula cell's reads and level found again as `Calc` would find
    /// them, none of the plan's code used: its text read off the grid (a
    /// shared formula's child as its master's text moved), its reads by
    /// `graph::dependencies`, the order by `graph::plan`; a cell's level 0
    /// when it reads no formula cell, else one past the deepest it reads.
    fn walked_again(m: &Machine) -> (BTreeMap<CellId, usize>, HashMap<CellId, Vec<CellId>>) {
        let book = &m.book;
        let names = Names::of(book);
        let extents: Vec<Extent> = book.sheets.iter().map(Sheet::extent).collect();
        let mut texts = HashMap::new();
        for (si, sheet) in book.sheets.iter().enumerate() {
            let masters = shared_masters(sheet);
            for (&(row, col), cell) in &sheet.cells {
                if let Some(text) = formula_text(&cell.content, row, col, &masters) {
                    texts.insert((si, row, col), text);
                }
            }
        }
        let plan = crate::calc::graph::plan(book, &names, &extents, &texts);
        assert!(plan.cycles.is_empty(), "a grid that loaded holds no cycle");
        let (mut level, mut reads) = (BTreeMap::new(), HashMap::new());
        for id in &plan.order {
            let mut deps = Vec::new();
            crate::calc::graph::dependencies(book, &names, &extents, id.0, &texts[id], &mut deps);
            let deepest = deps.iter().filter_map(|d| level.get(d)).max();
            level.insert(*id, deepest.map_or(0, |l: &usize| l + 1));
            reads.insert(*id, deps);
        }
        (level, reads)
    }

    /// Each shape's runs found again by the rule restated over the plan's
    /// order: a stretch of one shape on one sheet, none of whose cells reads
    /// another of it by the reads found again, as long as the next cell of
    /// the order can join it, kept when at least `RUN` long and its shape
    /// parsed and not risky; as its count, its cells, the fewest and the
    /// most in one.
    fn runs_again(
        m: &Machine,
        reads: &HashMap<CellId, Vec<CellId>>,
    ) -> BTreeMap<usize, (usize, usize, usize, usize)> {
        let order = m.shapes.order();
        let shape_of = |id: CellId| m.shapes.place(id).expect("placed").shape;
        let mut out = BTreeMap::new();
        let mut i = 0;
        while i < order.len() {
            let (first, s) = (order[i], shape_of(order[i]));
            let mut stretch = std::collections::HashSet::from([first]);
            let mut j = i + 1;
            while j < order.len()
                && shape_of(order[j]) == s
                && order[j].0 == first.0
                && !reads[&order[j]].iter().any(|r| stretch.contains(r))
            {
                stretch.insert(order[j]);
                j += 1;
            }
            let shape = m.shapes.shape(s).expect("a shape");
            if j - i >= RUN && shape.expr.is_ok() && !shape.risky {
                let r = out.entry(s).or_insert((0, 0, usize::MAX, 0));
                (r.0, r.1, r.2, r.3) = (r.0 + 1, r.1 + j - i, r.2.min(j - i), r.3.max(j - i));
            }
            i = j;
        }
        out
    }

    #[test]
    fn the_census_is_the_plan_walked_again() {
        // KERNEL.30: over the generated grids, at the load and after writes
        // that plan the order again, every count the census prints found
        // again by a walk of its own: each shape's cells from the cells, the
        // levels from the reads `Calc` reads, the runs by the rule restated;
        // and every count summed back to the plan's.
        for seed in 1..=80u64 {
            let mut d = Dice(seed);
            let rows = generated_grid(&mut d);
            let mut m = load(&rows);
            for round in 0..3 {
                let when = format!("seed {seed}, round {round}");
                let census = census_of(&m);
                let (level, reads) = walked_again(&m);
                let runs = runs_again(&m, &reads);
                let mut cells: BTreeMap<usize, usize> = BTreeMap::new();
                let mut at: BTreeMap<usize, BTreeMap<usize, usize>> = BTreeMap::new();
                for (si, sheet) in m.book.sheets.iter().enumerate() {
                    for (&(row, col), cell) in &sheet.cells {
                        if !is_formula(&cell.content) {
                            continue;
                        }
                        let id = (si, row, col);
                        let s = m.shapes.place(id).expect("placed").shape;
                        *cells.entry(s).or_default() += 1;
                        *at.entry(s).or_default().entry(level[&id]).or_default() += 1;
                    }
                }
                assert_eq!(census.formulas, cells.values().sum(), "{when}: the cells");
                assert_eq!(census.cycles, 0, "{when}");
                let held: BTreeSet<usize> = level.values().copied().collect();
                assert_eq!(census.levels, held.len(), "{when}: the levels");
                assert_eq!(census.shapes.len(), cells.len(), "{when}: the shapes");
                for f in &census.shapes {
                    let s = f.shape;
                    assert_eq!(f.cells, cells[&s], "{when}: shape {s}'s cells");
                    let by = &at[&s];
                    let first = *by.keys().next().expect("a level");
                    let last = *by.keys().last().expect("a level");
                    let widest = *by.values().max().expect("a level");
                    assert_eq!(
                        f.levels,
                        Some((first, last, by.len(), widest)),
                        "{when}: shape {s}'s levels"
                    );
                    assert_eq!(f.in_order, f.cells, "{when}: shape {s}, no cycle");
                    let want = runs.get(&s).copied().unwrap_or((0, 0, 0, 0));
                    assert_eq!(
                        (f.runs, f.lanes, f.narrowest, f.widest),
                        want,
                        "{when}: shape {s}'s runs"
                    );
                    assert_eq!(f.alone() + f.lanes, f.cells, "{when}: shape {s}");
                }
                // A write, as the law writes, or a fill planting a formula
                // over cells of the bands, so that the order is planned again.
                let write = match d.below(3) {
                    0 => format!("(cell \"A\" \"A{}\" 7)", 1 + d.below(20)),
                    1 => format!(
                        "(cell \"A\" \"{}{}\" 2)",
                        d.pick(&LETTERS[2..8]),
                        1 + d.below(4)
                    ),
                    _ => "(formula \"B\" \"C1:H2\" \"=$A1+1\")".to_string(),
                };
                assert!(m.write(&write).is_ok(), "{when}: {write}");
                while !m.step(0).done {}
            }
        }
    }

    /// Whether the census names a path a run took.
    fn named(f: &crate::calc::census::Facts, path: &vector::taken::Path) -> bool {
        use crate::calc::census::Kind;
        use vector::taken::Path;
        match path {
            Path::Stride => f.reads.contains(&Kind::Stride),
            Path::Slot => f.reads.contains(&Kind::Slot),
            Path::Array => f.reads.contains(&Kind::Array),
            Path::Lookup => f.reads.contains(&Kind::Lookup),
            Path::Cells => f.reads.contains(&Kind::Cells),
            Path::Once => f.reads.contains(&Kind::Once),
            Path::Each => f.reads.contains(&Kind::Each),
            Path::Call(name) => f.calls.contains(name),
        }
    }

    #[test]
    fn the_census_names_every_path_a_run_takes() {
        // KERNEL.30: over the generated grids, stepped in slices of any
        // budget with writes between frames, every path a run takes, noted
        // as it takes it, is a kind of read or a call the census names for
        // the run's shape, the census taken at the frame's start, when the
        // twins' arrays are the ones the run reads, or, for a frame's own
        // sheet that steps its array aside while a formula of it holds a
        // reason, the census with no array of the rows (KERNEL.31); and every
        // kind of path is taken somewhere, so that the law holds over each.
        let mut kinds: BTreeSet<&str> = BTreeSet::new();
        for seed in 1..=80u64 {
            let mut d = Dice(seed);
            let rows = generated_grid(&mut d);
            let mut m = load(&rows);
            for frame in 1..=4 {
                let census = census_of(&m);
                let unkept = census_unkept(&m);
                let facts: BTreeMap<usize, &crate::calc::census::Facts> =
                    census.shapes.iter().map(|f| (f.shape, f)).collect();
                let aside: BTreeMap<usize, &crate::calc::census::Facts> =
                    unkept.shapes.iter().map(|f| (f.shape, f)).collect();
                let stepped = m.aside.get();
                vector::taken::begin();
                while !m.step(1 + d.below(300)).done {}
                let stepped = m.aside.get() > stepped;
                for (s, path) in vector::taken::end() {
                    let Some(f) = facts.get(&s) else {
                        panic!(
                            "seed {seed}, frame {frame}: a run of shape {s}, which no row names"
                        );
                    };
                    assert!(
                        named(f, &path) || (stepped && named(aside[&s], &path)),
                        "seed {seed}, frame {frame}: shape {s} took {path:?}; its facts {f:?}\n{rows}"
                    );
                    kinds.insert(match path {
                        vector::taken::Path::Stride => "stride",
                        vector::taken::Path::Slot => "slot",
                        vector::taken::Path::Array => "array",
                        vector::taken::Path::Lookup => "lookup",
                        vector::taken::Path::Cells => "cells",
                        vector::taken::Path::Once => "once",
                        vector::taken::Path::Each => "each",
                        vector::taken::Path::Call(_) => "call",
                    });
                }
                let write = match d.below(4) {
                    0 => format!("(cell \"A\" \"A{}\" 7)", 1 + d.below(20)),
                    1 => format!("(cell \"B\" \"B{}\" \"y\")", 1 + d.below(20)),
                    2 => format!(
                        "(cell \"A\" \"{}{}\" 2)",
                        d.pick(&LETTERS[2..8]),
                        1 + d.below(4)
                    ),
                    _ => continue,
                };
                assert!(m.write(&write).is_ok(), "seed {seed}: {write}");
            }
        }
        assert_eq!(
            kinds.into_iter().collect::<Vec<_>>(),
            ["array", "call", "cells", "each", "lookup", "once", "slot", "stride"],
            "every kind of path taken somewhere"
        );
    }

    #[test]
    fn each_kind_is_named_by_its_own_formula() {
        // KERNEL.30: a formula for each kind of read, each filling eight
        // cells, a run, whose census names that kind and no other; and the
        // paths its run takes are the kinds named, but the parts compared at
        // the frame's start, which no run takes, and the reads under a node
        // the evaluator computes at each cell, which are its own. `T`, whose
        // own cells and range its runs read, keeps the frame's values as an
        // array (KERNEL.31), and holds no cell not computed, as `S` does at
        // `M`'s, a range where one value is wanted; `W`, widened past four
        // slots a cell by a cell far from the rest, keeps none, and is read
        // through its cells.
        let mut m = load(concat!(
            "(sheet \"V\" visible)\n",
            "(cell \"V\" \"A1\" 1)\n",
            "(sheet \"S\" visible)\n",
            "(cell \"S\" \"A1:B8\" 1)\n",
            "(sheet \"W\" visible)\n",
            "(cell \"W\" \"A1:B9\" 1)\n",
            "(cell \"W\" \"XFA1\" 1)\n",
            "(sheet \"T\" visible)\n",
            "(cell \"T\" \"A1:P8\" 1)\n",
            "(name \"Rate\" \"S!$A$1\")\n",
            "(formula \"S\" \"C1:C8\" \"=S.last!A1\")\n",
            "(formula \"S\" \"D1:D8\" \"=SUM(S.last!A1:B2)\")\n",
            "(formula \"S\" \"E1:E8\" \"=W!A1\")\n",
            "(formula \"S\" \"F1:F8\" \"=SUM(W!A1:B1)\")\n",
            "(formula \"S\" \"G1:G8\" \"=$A$1+1\")\n",
            "(formula \"S\" \"H1:H8\" \"=IF(V!$A$1=1,1,0)\")\n",
            "(formula \"S\" \"I1:I8\" \"=Rate+1\")\n",
            "(formula \"S\" \"J1:J8\" \"=MOD(S.last!A1,2)\")\n",
            "(formula \"S\" \"K1:K8\" \"=ROW()\")\n",
            "(formula \"S\" \"L1:L8\" \"=SUM(IF(W!A1>0,W!A1:A2,W!B1:B2))\")\n",
            "(formula \"S\" \"M1:M8\" \"=ABS(S.last!A1:A2)\")\n",
            "(formula \"S\" \"N1:N8\" \"=S.last!$A1\")\n",
            "(formula \"S\" \"O1:O8\" \"=SUM(S.last!$A1:$B2)\")\n",
            "(formula \"S\" \"P1\" \"=S.last!A1\")\n",
            "(formula \"T\" \"Q1:Q8\" \"=A1\")\n",
            "(formula \"T\" \"R1:R8\" \"=SUM(A1:B1)\")\n",
        ));
        let want: [(&str, &str, &[&str]); 16] = [
            ("C1", "stride", &[]),
            ("D1", "stride", &[]),
            ("E1", "lookup", &[]),
            ("F1", "cells", &["SUM"]),
            ("G1", "once start", &[]),
            ("H1", "once start", &[]),
            ("I1", "each", &[]),
            ("J1", "stride", &[]),
            ("K1", "", &[]),
            ("L1", "lookup cells each", &["IF", "SUM"]),
            // A function of numbers given a range is the library's at each
            // cell, and the range its walk of the array.
            ("M1", "array", &["ABS"]),
            // A reference whose column a `$` holds, one cell or a range at both
            // corners, is read by index, never by slices (KERNEL.43); and so
            // is one a cell alone reads, no run having cells beside it.
            ("N1", "slot", &[]),
            ("O1", "array", &[]),
            ("P1", "slot", &[]),
            // The frame's own cells and range, its sheet kept as an array
            // (KERNEL.31): read and folded by the run as a twin's are.
            ("Q1", "stride", &[]),
            ("R1", "stride", &[]),
        ];
        // Runs of eight, fewer than `SHORT`, are evaluated cell by cell by
        // their programs (KERNEL.44): the paths held here are the run's.
        m.lowered = false;
        let census = census_of(&m);
        vector::taken::begin();
        assert!(m.step(0).done);
        let taken = vector::taken::end();
        assert_eq!(census.shapes.len(), want.len());
        for (f, (first, reads, calls)) in census.shapes.iter().zip(want) {
            let words: Vec<&str> = f.reads.iter().map(|k| k.word()).collect();
            assert_eq!(
                (f.first.as_str(), words.join(" "), f.calls.clone()),
                (
                    first,
                    reads.to_string(),
                    calls.iter().map(|c| c.to_string()).collect()
                ),
                "{f:?}"
            );
            let paths: BTreeSet<&vector::taken::Path> = taken
                .iter()
                .filter(|(s, _)| *s == f.shape)
                .map(|(_, p)| p)
                .collect();
            let words_taken: Vec<String> = paths
                .iter()
                .map(|p| match p {
                    vector::taken::Path::Call(name) => format!("call {name}"),
                    other => format!("{other:?}").to_ascii_lowercase(),
                })
                .collect();
            let each = f.reads.contains(&crate::calc::census::Kind::Each);
            let expected: Vec<String> = if f.runs == 0 {
                // A cell alone is evaluated by `Ctx`, which notes no path.
                Vec::new()
            } else if each {
                vec!["each".to_string()]
            } else {
                let mut w: Vec<String> = f
                    .reads
                    .iter()
                    .filter(|k| **k != crate::calc::census::Kind::Start)
                    .map(|k| k.word().to_string())
                    .collect();
                w.extend(f.calls.iter().map(|c| format!("call {c}")));
                w
            };
            assert_eq!(words_taken, expected, "{first}: the paths taken");
        }
        let every: Vec<&str> = census
            .shapes
            .iter()
            .filter(|f| f.every_frame.is_some())
            .map(|f| f.first.as_str())
            .collect();
        assert_eq!(
            every,
            ["I1"],
            "the shape reading a name computed every frame"
        );
    }

    #[test]
    fn a_cell_in_a_run_answers_the_first_reason_its_evaluation_meets() {
        // KERNEL.20, slice 3: every operand a reason, in every place a run
        // computes: both sides of an operator, each the same at every cell or
        // each cell's own; a strict call's operands at every count; a chain;
        // a condition. Each cell's answer is the first reason met left to
        // right, its label moved with the cell, as alone.
        let grid = concat!(
            "(cell \"S\" \"A1:B8\" 1)\n",
            "(formula \"S\" \"X1\" \"=A1:A2\")\n",
            "(formula \"S\" \"C1:C8\" \"=(A1:A2)+(B1:B2)\")\n",
            "(formula \"S\" \"D1:D8\" \"=$X$1+(B1:B2)\")\n",
            "(formula \"S\" \"E1:E8\" \"=(B1:B2)+$X$1\")\n",
            "(formula \"S\" \"F1:F8\" \"=MOD((A1:A2)+1,(B1:B2)+1)\")\n",
            "(formula \"S\" \"G1:G8\" \"=SUM((A1:A2)+1,(B1:B2)+1,(A3:A4)+1)\")\n",
            "(formula \"S\" \"H1:H8\" \"=ABS((B1:B2)*2)\")\n",
            "(formula \"S\" \"I1:I8\" \"=1+(A1:A2)-(B1:B2)\")\n",
            "(formula \"S\" \"J1:J8\" \"=IF((A1:A2)=1,(B1:B2),0)\")\n",
            "(formula \"S\" \"K1:K8\" \"=$X$1&$X$1\")\n",
        );
        let mut runs = load(grid);
        let mut alone = load(grid);
        alone.by_runs = false;
        assert_eq!(runs.shapes.cells_in_runs(), 72, "nine runs of eight");
        runs.step(0);
        alone.step(0);
        assert_eq!(values(&runs), values(&alone));
        let first = |col: &str, row: u32| match at(&runs, "S", &format!("{col}{row}")) {
            Some(Computed::NotComputed(Reason::Construct(c))) => c,
            other => panic!("{col}{row}: {other:?}"),
        };
        for (col, wants) in [
            ("C", "A3:A4"),
            ("D", "A1:A2"),
            ("E", "B3:B4"),
            ("F", "A3:A4"),
            ("G", "A3:A4"),
            ("H", "B3:B4"),
            ("I", "A3:A4"),
            ("J", "A3:A4"),
            ("K", "A1:A2"),
        ] {
            assert_eq!(first(col, 3), wants, "{col}3");
        }
    }

    #[test]
    fn a_run_makes_every_buffer_it_holds_in_its_first_frame() {
        // KERNEL.20, slice 3: a shape whose first frame takes a branch with
        // nothing in it and whose later frames take one of many calls, as
        // Life's and the gun's do: the buffers its later frames hold are all
        // made in the first, and none holds more than a chunk of cells; with
        // them, the lanes a later frame keeps for the parts Life's rule writes
        // twice (KERNEL.27), and the stretches' buffer, which only the later
        // frames' reads of the twin's array fill (KERNEL.43).
        let mut m = load(concat!(
            "(cell \"S\" \"A1\" 1)\n",
            "(formula \"S\" \"B1:B300\" \"=IF(S.last!$A$1=1,0,IF(OR(SUM(S.last!B1:B3)=3,",
            "AND(S.last!B2=1,MOD(SUM(S.last!A1:C3),4)=2)),1,0))\")\n",
            "(formula \"S\" \"C1:C300\" \"=IF(S.last!$A$1=1,0,IF(OR(SUM(S.last!B1:D3)-S.last!C2=3,",
            "AND(S.last!C2=1,SUM(S.last!B1:D3)-S.last!C2=2)),1,0))\")\n",
        ));
        assert_eq!(m.shapes.cells_in_runs(), 600);
        m.step(0);
        let made = m.lanes.made();
        assert_eq!(m.write("(cell \"S\" \"A1\" 2)"), Ok(1));
        for frame in 2..=4 {
            m.step(0);
            assert_eq!(m.lanes.made(), made, "frame {frame}");
        }
        assert_eq!(number(&m, "S", "B2"), 0.0);
        assert!(
            made.0 > 3 && made.4 == vector::CHUNK && made.5 == vector::CHUNK,
            "{made:?}"
        );
        assert!(
            made.1 == made.0 && made.2 == made.0,
            "each kind of lane: {made:?}"
        );
    }

    #[test]
    fn a_run_holds_its_parts_kept_lanes_within_its_shape_s_bound() {
        // KERNEL.27: a part's lane kept beside the lanes the rest of its
        // formula holds, which `holds` alone does not count: `C1` read and a
        // copy kept, `D1*2` beside them, and the copy taken for the last `C1`.
        // The run makes the buffers its shape's bound says, no more.
        let mut m = load(concat!(
            "(cell \"S\" \"C1:D300\" 2)\n",
            "(formula \"S\" \"B1:B300\" \"=S.last!C1+(S.last!D1*2)+S.last!C1\")\n",
        ));
        assert_eq!(m.shapes.cells_in_runs(), 300);
        m.step(0);
        let (held, _) = m.shapes.shape(0).expect("the shape").holds();
        assert_eq!(held, 3, "two lanes, and one kept for the part");
        let made = m.lanes.made();
        assert_eq!((made.0, made.1, made.2), (held, held, held), "{made:?}");
        assert_eq!(number(&m, "S", "B7"), 8.0);
    }

    thread_local! {
        /// How many times `TICK` has been called on this thread.
        static TICKS: Slot<usize> = const { Slot::new(0) };
    }

    /// A function that counts its calls, once an evaluation of its cell.
    fn tick(_: &crate::calc::Ctx<'_>, _: &[crate::calc::Arg<'_>]) -> Result<Value, Reason> {
        TICKS.with(|t| t.set(t.get() + 1));
        Ok(Value::Number(1.0))
    }

    #[test]
    fn a_run_evaluates_no_cell_past_a_step_s_budget() {
        // KERNEL.20, slice 3: a run of 600 cells, longer than the cells it
        // evaluates at once, stepped in budgets that end inside those, at
        // their end and past it: each step evaluates exactly the cells its
        // budget allows, counted by a function every evaluation calls once.
        let mut lib = library::language();
        lib.register("TICK", tick);
        let mut m = Machine::load("(formula \"S\" \"A1:F100\" \"=TICK()+S.last!A1\")\n", lib)
            .unwrap_or_else(|e| panic!("{}", e.refusal.text));
        assert_eq!(m.shapes.cells_in_runs(), 600);
        for budget in [1, 7, 255, 256, 257, 300, 0] {
            let mut done = 0;
            loop {
                let before = TICKS.with(Slot::get);
                let step = m.step(budget);
                let ticks = TICKS.with(Slot::get) - before;
                assert_eq!(ticks, step.evaluated - done, "budget {budget} after {done}");
                if step.done {
                    break;
                }
                done = step.evaluated;
            }
        }
        assert_eq!(number(&m, "S", "F100"), 7.0, "seven frames");
    }

    /// A grid's formulas, one a column of `rows` rows from column D, each
    /// column one shape, after its value rows (KERNEL.26).
    fn columns(values: &str, rows: u32, formulas: &[String]) -> String {
        let mut out = values.to_string();
        for (k, f) in formulas.iter().enumerate() {
            let col = 4 + k as u32;
            let range = format!(
                "{}:{}",
                crate::sheet::cell_ref(1, col),
                crate::sheet::cell_ref(rows, col)
            );
            let text = crate::rows::quoted(&format!("={f}"));
            out.push_str(&format!("(formula \"S\" \"{range}\" {text})\n"));
        }
        out
    }

    /// The same grid stepped by runs and with every cell evaluated alone,
    /// frame for frame and bit for bit, its first frame `Calc`'s, a write
    /// after each frame and two frames after the last, so that the last's
    /// value reaches a twin and is read; the grid stepped by runs, after its
    /// last frame.
    fn by_runs_and_alone(rows: &str, writes: &[&str]) -> Machine {
        let mut runs = load(rows);
        let mut alone = load(rows);
        alone.by_runs = false;
        assert_eq!(
            runs.shapes.cells_in_runs(),
            runs.shapes.order().len(),
            "every cell in a run"
        );
        let calc: BTreeMap<CellId, String> = {
            let calc = crate::calc::Calc::run(&runs.book, &runs.library);
            let spell = |id: &CellId| format!("{:?}", calc.computed(*id).expect("computed"));
            runs.shapes
                .order()
                .iter()
                .map(|id| (*id, spell(id)))
                .collect()
        };
        for frame in 1..=writes.len() + 2 {
            assert!(runs.step(0).done && alone.step(0).done);
            assert_eq!(values(&runs), values(&alone), "frame {frame}");
            if frame == 1 {
                assert_eq!(values(&runs), calc, "Calc");
            }
            if let Some(w) = writes.get(frame - 1) {
                assert!(runs.write(w).is_ok() && alone.write(w).is_ok(), "{w}");
            }
        }
        runs
    }

    #[test]
    fn the_aggregates_a_run_computes_are_the_library_s_at_every_cell() {
        // KERNEL.26: SUM, MIN, MAX, AND, OR and NOT computed by the run with
        // the library's own rule and fold, at every count of operands and over
        // every kind of value, given directly and read out of a range, against
        // the same grid with every cell evaluated alone: a range of a twin
        // kept as an array, of a twin too sparse for one and of the rows' own
        // sheet, both walked through their cells; operands of numbers and of
        // truth values; a sheet the grid does not hold; a range moved off the
        // sheet; more operands than the run computes itself.
        let values_rows = concat!(
            "(cell \"S\" \"A1\" 1)\n",
            "(cell \"S\" \"A2\" -2.5)\n",
            "(cell \"S\" \"A3\" \"5\")\n",
            "(cell \"S\" \"A4\" \"x\")\n",
            "(cell \"S\" \"A5\" true)\n",
            "(cell \"S\" \"A6\" (error \"#N/A\"))\n",
            "(cell \"S\" \"B1\" 0)\n",
            "(cell \"S\" \"B2\" false)\n",
            "(cell \"S\" \"B3\" 3)\n",
            "(cell \"S\" \"B5\" \"TRUE\")\n",
            "(cell \"Far\" \"A1\" 2)\n",
            "(cell \"Far\" \"Z900\" 3)\n",
        );
        let operands = [
            "S.last!A1:B3",
            "S.last!A4:B6",
            "$A$1:$B$3",
            "Far.last!A1:B3",
            "S.last!A1,\"5\",TRUE,S.last!B2",
            "Nope!A1:B2",
            "S.last!A1048575:A1048576",
            "",
            "1,2,0,4,5,6,7,8",
            "1,2,0,4,5,6,7,8,S.last!A1",
            "S.last!A1>0,S.last!B1<1",
            "S.last!A1+1,S.last!B3*2",
            "-S.last!A2,S.last!A1:A3,S.last!B1=0",
        ];
        let mut formulas = Vec::new();
        for f in ["SUM", "MIN", "MAX", "AND", "OR"] {
            for o in operands {
                formulas.push(format!("{f}({o})"));
            }
        }
        for o in [
            "S.last!A1>0",
            "S.last!A1",
            "S.last!A1+0",
            "\"TRUE\"",
            "S.last!A3&\"\"",
            "S.last!A1:A2",
            "",
            "1,2",
        ] {
            formulas.push(format!("NOT({o})"));
        }
        let grid = columns(values_rows, 8, &formulas);
        let m = by_runs_and_alone(
            &grid,
            &[
                "(cell \"S\" \"A1\" 4)",
                "(cell \"S\" \"B3\" \"y\")",
                "(cell \"Far\" \"A1\" (error \"#DIV/0!\"))",
            ],
        );
        // The first, SUM over the twin's A1:B3 moved to each row: at D1, after
        // the writes of A1 and B3, 4 + 0 - 2.5.
        assert_eq!(number(&m, "S", "D1"), 1.5);
    }

    /// A meaning of its own for a name the language's library holds.
    fn forty_two(_: &crate::calc::Ctx<'_>, _: &[crate::calc::Arg<'_>]) -> Result<Value, Reason> {
        Ok(Value::Number(42.0))
    }

    #[test]
    fn a_meaning_registered_over_the_language_s_own_is_the_one_a_run_computes() {
        // KERNEL.26: a run computes the language's aggregates and `NOT` by
        // their rule and fold only while the library holds the language's
        // own; a crate that registers a meaning of its own under the name has
        // its meaning computed, at every cell of the run. KERNEL.28: so for
        // the language's other functions, `CHOOSE` among them, a form the
        // registration replaces with a strict function.
        for name in [
            "SUM", "MIN", "MAX", "AND", "OR", "NOT", "ABS", "INT", "SIN", "MOD", "ROW", "COLUMN",
            "CHOOSE",
        ] {
            let mut lib = library::language();
            lib.register(name, forty_two);
            let rows = format!(
                "(cell \"S\" \"A1:A8\" 1)\n(formula \"S\" \"B1:B8\" \"={name}(S.last!A1>0)\")\n"
            );
            let mut m = Machine::load(&rows, lib).unwrap_or_else(|e| panic!("{}", e.refusal.text));
            assert_eq!(m.shapes.cells_in_runs(), 8, "{name}");
            m.step(0);
            for row in 1..=8 {
                assert_eq!(number(&m, "S", &format!("B{row}")), 42.0, "{name}");
            }
        }
    }

    /// The value cells the tests of the language's other functions read
    /// (KERNEL.28): numbers of each sign and zero, a text that reads as a
    /// number and one that does not, truth values, an error and empty cells
    /// down columns A and B; another sheet's number; and down column C a
    /// reason at each row, a range where one value is wanted, its label moved
    /// with the row.
    const EVERY_KIND: &str = concat!(
        "(cell \"S\" \"A1\" 1)\n",
        "(cell \"S\" \"A2\" -2.5)\n",
        "(cell \"S\" \"A3\" \"5\")\n",
        "(cell \"S\" \"A4\" \"x\")\n",
        "(cell \"S\" \"A5\" true)\n",
        "(cell \"S\" \"A6\" (error \"#N/A\"))\n",
        "(cell \"S\" \"A8\" 0)\n",
        "(cell \"S\" \"B1\" 0)\n",
        "(cell \"S\" \"B2\" false)\n",
        "(cell \"S\" \"B3\" 3)\n",
        "(cell \"S\" \"B4\" -2)\n",
        "(cell \"S\" \"B5\" \"TRUE\")\n",
        "(cell \"S\" \"B8\" 7.5)\n",
        "(cell \"Far\" \"C3\" 2)\n",
        "(formula \"S\" \"C1:C8\" \"=XX1:XX2\")\n",
    );

    /// The value a formula of a grid `columns` wrote computes at a row.
    fn formula_at(m: &Machine, formulas: &[String], text: &str, row: u32) -> Option<Computed> {
        let k = formulas
            .iter()
            .position(|f| f == text)
            .unwrap_or_else(|| panic!("no formula {text}"));
        at(m, "S", &crate::sheet::cell_ref(row, 4 + k as u32))
    }

    #[test]
    fn the_language_s_other_functions_a_run_computes_are_the_library_s_at_every_cell() {
        // KERNEL.28: ABS, INT, SIN and MOD computed by the run at the cells
        // whose operands are all numbers and by the library at any other, and
        // ROW and COLUMN from the run's own cells, against the same grid with
        // every cell evaluated alone: operands of every kind, given directly,
        // computed, read from the rows' sheet, its twin and another sheet's,
        // the same at every cell and each cell's own; a reason the same at
        // every cell and each cell's own; a range where one value is wanted; a
        // sheet the grid does not hold; a reference moved off the sheet; every
        // count of operands.
        let operands = [
            "S.last!A1",
            "A1",
            "B1",
            "S.last!A1+0",
            "$B$3",
            "$A$4",
            "-2.5",
            "\"7\"",
            "TRUE",
            "C1",
            "$C$1",
            "C1+1",
            "A1:A2",
            "S.last!A1>0",
            "ROW()*0.5-2",
            "Far.last!C3",
            "Nope!A1",
            "S.last!A1048576",
        ];
        let mut formulas = Vec::new();
        for f in ["ABS", "INT", "SIN"] {
            for o in operands {
                formulas.push(format!("{f}({o})"));
            }
            formulas.push(format!("{f}()"));
            formulas.push(format!("{f}(1,2)"));
        }
        for (x, y) in [
            ("S.last!A1", "S.last!B1"),
            ("A1", "B1"),
            ("S.last!A1", "S.last!B1+2"),
            ("ROW()*3-13", "ROW()-4"),
            ("-7", "S.last!B1+2"),
            ("S.last!B8", "-2"),
            ("$A$1", "3"),
            ("1E+300", "7"),
            ("S.last!A1", "$C$1"),
            ("C1", "2"),
            ("2", "C1"),
            ("A1:A2", "2"),
            ("\"9\"", "S.last!B3"),
            ("S.last!A1>0", "2"),
            ("Nope!A1", "2"),
        ] {
            formulas.push(format!("MOD({x},{y})"));
        }
        formulas.push("MOD(1)".to_string());
        formulas.push("MOD(1,2,3)".to_string());
        // The functions over references read at the cells a branch selects,
        // some of a run's cells and not the first of them.
        for f in [
            "IF(MOD(ROW(),2)=0,MOD(S.last!A1,3),ABS(S.last!B1))",
            "CHOOSE(MOD(ROW(),3)+1,INT(S.last!A1),SIN(A1),MOD(B1,4))",
            "IF(ROW()>2,ROW(S.last!A1)+COLUMN(),0)",
        ] {
            formulas.push(f.to_string());
        }
        for f in ["ROW", "COLUMN"] {
            for o in [
                "",
                "A1",
                "B3",
                "$A$2",
                "S.last!A1",
                "Far!C3",
                "Far.last!B2:C9",
                "Nope!A1",
                "A1048576",
                "A1048570:B1048576",
                "B:B",
                "S.last!3:3",
                "C1",
                "5",
                "A1+0",
                "A1,B1",
            ] {
                formulas.push(format!("{f}({o})"));
            }
        }
        let grid = columns(EVERY_KIND, 8, &formulas);
        let m = by_runs_and_alone(
            &grid,
            &[
                "(cell \"S\" \"A1\" -4)",
                "(cell \"S\" \"B1\" \"y\")",
                "(cell \"S\" \"A7\" 9)",
            ],
        );
        // The meanings themselves, the library's: MOD takes its divisor's sign
        // and is #DIV/0! by 0; a row a cell's own, or a reference's moved; a
        // reference moved off the sheet, or to a sheet the grid does not hold,
        // places no cell.
        let value = |text: &str, row: u32| formula_at(&m, &formulas, text, row);
        let number = |n: f64| Some(Computed::Value(Value::Number(n)));
        let error = |k: ErrorKind| Some(Computed::Value(Value::Error(k)));
        assert_eq!(value("MOD(ROW()*3-13,ROW()-4)", 1), number(-1.0));
        assert_eq!(value("MOD(ROW()*3-13,ROW()-4)", 4), error(ErrorKind::Div0));
        assert_eq!(value("MOD(ROW()*3-13,ROW()-4)", 8), number(3.0));
        // Operands of opposite signs: MOD(-7,5) is 3 and MOD(7.5,-2) is -0.5,
        // the divisor's sign, where a quotient cut toward 0 would give -2 and
        // 1.5.
        assert_eq!(value("MOD(-7,S.last!B1+2)", 3), number(3.0));
        assert_eq!(value("MOD(S.last!B8,-2)", 1), number(-0.5));
        assert_eq!(value("INT(ROW()*0.5-2)", 1), number(-2.0));
        assert_eq!(value("ABS(ROW()*0.5-2)", 1), number(1.5));
        assert_eq!(value("ROW()", 5), number(5.0));
        let column = 4 + formulas
            .iter()
            .position(|f| f == "COLUMN()")
            .expect("COLUMN()");
        assert_eq!(value("COLUMN()", 5), number(column as f64));
        assert_eq!(value("ROW(B3)", 5), number(7.0));
        assert_eq!(value("ROW(A1048576)", 1), number(1_048_576.0));
        assert_eq!(value("ROW(A1048576)", 2), error(ErrorKind::Value));
        assert_eq!(value("ROW(Nope!A1)", 2), error(ErrorKind::Value));
        assert_eq!(value("COLUMN(Far.last!B2:C9)", 3), number(2.0));
        assert_eq!(value("ROW(C1)", 6), number(6.0));
    }

    #[test]
    fn a_call_of_two_operands_meets_their_reasons_and_errors_in_the_library_s_order() {
        // KERNEL.28: the order a run keeps by leaving to the library every
        // cell whose operands are not all numbers. Every operand is evaluated
        // first, left to right, a reason the answer; a reference written as
        // one is not read until the function reads it, after the operands
        // before it are values. So in MOD(A1,C1+1) the reason of C1+1 comes
        // before the error A1 holds, and in MOD(A1+1,C1) the error of A1+1
        // before the reason C1 holds; with both written as references, the
        // first read first.
        let rows = concat!(
            "(cell \"S\" \"A1:A8\" (error \"#N/A\"))\n",
            "(formula \"S\" \"C1:C8\" \"=XX1:XX2\")\n",
        );
        let formulas: Vec<String> = ["MOD(A1,C1+1)", "MOD(A1+1,C1)", "MOD(C1,A1)", "MOD(A1,C1)"]
            .iter()
            .map(|f| f.to_string())
            .collect();
        let m = by_runs_and_alone(&columns(rows, 8, &formulas), &[]);
        let reason = Some(Computed::NotComputed(Reason::Construct(
            "XX3:XX4".to_string(),
        )));
        let na = Some(Computed::Value(Value::Error(ErrorKind::NA)));
        assert_eq!(formula_at(&m, &formulas, "MOD(A1,C1+1)", 3), reason);
        assert_eq!(formula_at(&m, &formulas, "MOD(A1+1,C1)", 3), na);
        assert_eq!(formula_at(&m, &formulas, "MOD(C1,A1)", 3), reason);
        assert_eq!(formula_at(&m, &formulas, "MOD(A1,C1)", 3), na);
    }

    #[test]
    fn choose_in_a_run_is_the_library_s_at_every_cell() {
        // KERNEL.28: CHOOSE computed by the run as IF is, its index at every
        // cell and each value at the cells that choose it, against the same
        // grid with every cell evaluated alone: an index at each edge of its
        // count and of every kind, the same at every cell and each cell's own,
        // a reason among them; values of every kind, a reference read as one
        // value and a range where one value is wanted among them; a value left
        // out; many values; CHOOSE inside IF and IF inside CHOOSE; CHOOSE an
        // operand of the calls the run computes, and of one it leaves to the
        // library at each cell, a value of it a range; CHOOSE written twice,
        // one part.
        let mut formulas: Vec<String> = [
            "0.5",
            "1",
            "2.9",
            "3",
            "3.9",
            "4",
            "-1",
            "\"2\"",
            "TRUE",
            "S.last!A7",
            "#N/A",
            "C1",
            "$C$1",
            "ROW()",
            "MOD(ROW(),3)+1",
            "ROW()-4",
            "S.last!A1:A2",
            "1E+300",
            "S.last!A1+1",
        ]
        .iter()
        .map(|i| format!("CHOOSE({i},10,20,30)"))
        .collect();
        formulas.extend(
            [
                "CHOOSE(MOD(ROW(),4)+1,S.last!A1,A1:A2,\"t\",C1)",
                "CHOOSE(MOD(ROW(),3)+1,,5,S.last!B1)",
                "CHOOSE(1)",
                "CHOOSE(ROW(),1,2,3,4,5,6,7,8,9)",
                "CHOOSE(MOD(ROW(),2)+1,IF(S.last!A1>0,1,2),CHOOSE(MOD(ROW(),3)+1,7,8,9))",
                "IF(S.last!B1>0,CHOOSE(MOD(ROW(),2)+1,S.last!A1,2),0)",
                "MOD(CHOOSE(MOD(ROW(),2)+1,5,7),3)",
                "SUM(CHOOSE(MOD(ROW(),2)+1,1,S.last!A1),3)",
                "ABS(CHOOSE(MOD(ROW(),2)+1,-1,\"x\"))",
                "SUM(CHOOSE(MOD(ROW(),2)+1,S.last!A1:A2,3))",
                "CHOOSE(MOD(ROW(),2)+1,S.last!A1,1)*CHOOSE(MOD(ROW(),2)+1,S.last!A1,1)",
            ]
            .iter()
            .map(|f| f.to_string()),
        );
        let m = by_runs_and_alone(
            &columns(EVERY_KIND, 8, &formulas),
            &[
                "(cell \"S\" \"A1\" -4)",
                "(cell \"S\" \"B1\" \"y\")",
                "(cell \"S\" \"A7\" 2)",
            ],
        );
        // The index's rule, the library's: read down to a whole number, of a
        // text that reads as one or a truth value, and #VALUE! outside 1 to
        // the count; a value chosen at each row.
        let value = |text: &str, row: u32| formula_at(&m, &formulas, text, row);
        let number = |n: f64| Some(Computed::Value(Value::Number(n)));
        let error = |k: ErrorKind| Some(Computed::Value(Value::Error(k)));
        assert_eq!(value("CHOOSE(2.9,10,20,30)", 1), number(20.0));
        assert_eq!(value("CHOOSE(3.9,10,20,30)", 1), number(30.0));
        assert_eq!(value("CHOOSE(4,10,20,30)", 1), error(ErrorKind::Value));
        assert_eq!(value("CHOOSE(0.5,10,20,30)", 1), error(ErrorKind::Value));
        assert_eq!(value("CHOOSE(\"2\",10,20,30)", 1), number(20.0));
        assert_eq!(value("CHOOSE(TRUE,10,20,30)", 1), number(10.0));
        assert_eq!(value("CHOOSE(#N/A,10,20,30)", 1), error(ErrorKind::NA));
        assert_eq!(value("CHOOSE(ROW(),10,20,30)", 3), number(30.0));
        assert_eq!(value("CHOOSE(ROW(),10,20,30)", 4), error(ErrorKind::Value));
        assert_eq!(value("CHOOSE(ROW(),1,2,3,4,5,6,7,8,9)", 8), number(8.0));
        assert_eq!(value("CHOOSE(MOD(ROW(),3)+1,,5,S.last!B1)", 3), number(0.0));
        assert_eq!(value("MOD(CHOOSE(MOD(ROW(),2)+1,5,7),3)", 1), number(1.0));
    }

    #[test]
    fn a_run_makes_the_buffers_choose_holds_in_its_first_frame() {
        // KERNEL.28: a shape whose first frame takes a branch with nothing in
        // it and whose later frames choose among values with calls of their
        // own: the buffers its later frames hold, the lists that gather each
        // value's cells among them, are all made in the first.
        let mut m = load(concat!(
            "(cell \"S\" \"A1\" 1)\n",
            "(cell \"S\" \"B1:B300\" 2)\n",
            "(formula \"S\" \"C1:C300\" \"=IF(S.last!$A$1=1,0,CHOOSE(MOD(ROW()+S.last!$A$1,3)+1,",
            "S.last!B1*2,IF(S.last!B1>0,MOD(S.last!B1,7),1),ABS(S.last!B1-ROW())))\")\n",
        ));
        assert_eq!(m.shapes.cells_in_runs(), 300);
        m.step(0);
        let made = m.lanes.made();
        assert_eq!(m.write("(cell \"S\" \"A1\" 2)"), Ok(1));
        for frame in 2..=4 {
            m.step(0);
            assert_eq!(m.lanes.made(), made, "frame {frame}");
        }
        // Row 2 in the frames after: MOD(2+2,3)+1 is 2, MOD(2,7).
        assert_eq!(number(&m, "S", "C2"), 2.0);
        assert!(
            made.3 >= 3,
            "the lists that gather a value's cells: {made:?}"
        );
        // Alone in its grid, so that no other shape's bound covers it: a
        // CHOOSE whose index is the same at every cell in the first frame,
        // which so computes one value and gathers nothing, and differs by
        // cell after, when the cells choosing an IF need the lists of both
        // beside each other.
        let mut m = load(concat!(
            "(cell \"S\" \"A1\" 1)\n",
            "(cell \"S\" \"B1:B300\" 2)\n",
            "(formula \"S\" \"D1:D300\" \"=CHOOSE(IF(S.last!$A$1=1,1,MOD(ROW(),3)+1),",
            "S.last!B1*2,IF(MOD(ROW(),2)=0,MOD(S.last!B1,7),1),5)\")\n",
        ));
        m.step(0);
        let made = m.lanes.made();
        assert_eq!(m.write("(cell \"S\" \"A1\" 2)"), Ok(1));
        for frame in 2..=4 {
            m.step(0);
            assert_eq!(m.lanes.made(), made, "alone, frame {frame}");
        }
        // Row 4 in the frames after: MOD(4,3)+1 is 2, an even row, MOD(2,7).
        assert_eq!(number(&m, "S", "D4"), 2.0);
    }

    #[test]
    fn a_lane_of_numbers_is_values_from_the_first_answer_of_another_kind() {
        // KERNEL.26: lanes of numbers and of truth values, each made a lane
        // of values where an answer of another kind first appears, against the
        // same grid with every cell evaluated alone: a read meeting a text, an
        // empty cell and an error; arithmetic meeting `#DIV/0!` and `#NUM!`; a
        // comparison of numbers read as numbers; `IF` over numbers and truth
        // values with a branch of texts; `NOT` over numbers and over truth
        // values; a function's answers meeting an error; signs, percents and
        // `&` over numbers.
        let values_rows = concat!(
            "(cell \"S\" \"A1\" 1)\n",
            "(cell \"S\" \"A2\" 2)\n",
            "(cell \"S\" \"A3\" 0)\n",
            "(cell \"S\" \"A4\" \"t\")\n",
            "(cell \"S\" \"A6\" (error \"#N/A\"))\n",
            "(cell \"S\" \"A7\" 4)\n",
            "(cell \"S\" \"A8\" 1E+300)\n",
            "(cell \"S\" \"B1:B2\" 1)\n",
            "(cell \"S\" \"B3\" 0)\n",
            "(cell \"S\" \"B4:B8\" 1)\n",
        );
        let formulas: Vec<String> = [
            "S.last!A1",
            "S.last!A1*2",
            "1/S.last!B1",
            "S.last!A1*1E+300",
            "(S.last!A1>1)+0",
            "IF(S.last!A1>1,\"big\",S.last!A1)",
            "IF(S.last!B1,S.last!A1*2,-1)",
            "IF(S.last!B1>0,S.last!B1,\"no\")",
            "NOT(S.last!B1+0)",
            "NOT(S.last!B1>0)",
            "MOD(S.last!A1,S.last!B1)",
            "-S.last!B1",
            "S.last!B1%",
            "S.last!B1&\"x\"",
            "S.last!A1=S.last!B1",
            "S.last!B1=S.last!B2",
            // Two lanes of numbers compared, every cell either side read
            // holding a number until a write makes one a text.
            "S.last!B1<S.last!B1*2",
            "S.last!B1*3>=S.last!B1+1",
            "S.last!B1+1>S.last!B1",
            "2^S.last!B1",
            "S.last!B1-S.last!B2*3",
            "SUM(S.last!A1:A2)/S.last!B1",
            "ABS(S.last!B1-2)",
        ]
        .iter()
        .map(|f| f.to_string())
        .collect();
        let grid = columns(values_rows, 8, &formulas);
        let m = by_runs_and_alone(
            &grid,
            &[
                "(cell \"S\" \"B5\" \"x\")",
                "(cell \"S\" \"A4\" 3)",
                "(cell \"S\" \"B3\" 2)",
            ],
        );
        // `1/S.last!B1` over B, 2 at B3 by the last write.
        assert_eq!(number(&m, "S", "F3"), 0.5);
    }

    #[test]
    fn a_value_follows_its_cell_when_a_formula_s_write_plans_the_order_again() {
        // KERNEL.26: each formula's value kept at its cell's position in the
        // order, moved with its cell when a formula's write plans the order
        // again: every cell the write leaves keeps its value until the next
        // frame computes it, the cell written has none, and the next frame is
        // the frame of the same grid loaded and stepped.
        let rows = concat!(
            "(cell \"S\" \"A1\" 2)\n",
            "(formula \"S\" \"B1:B8\" \"=S.last!$A$1*ROW()\")\n",
            "(formula \"S\" \"D1\" \"=B8+1\")\n",
        );
        let mut m = load(rows);
        m.step(0);
        let kept = values(&m);
        let at_b5 = m.shapes.reach().position((0, 5, 2));
        assert_eq!(m.write("(formula \"S\" \"B4\" \"=D1*2\")"), Ok(1));
        assert_ne!(m.shapes.reach().position((0, 5, 2)), at_b5, "B5 moved");
        for (id, v) in &kept {
            let now = m.computed(*id).map(|c| format!("{c:?}"));
            if *id == (0, 4, 2) {
                assert_eq!(now, None, "B4, written");
            } else {
                assert_eq!(now.as_ref(), Some(v), "{id:?}");
            }
        }
        m.step(0);
        let mut fresh = load(concat!(
            "(cell \"S\" \"A1\" 2)\n",
            "(formula \"S\" \"B1:B3\" \"=S.last!$A$1*ROW()\")\n",
            "(formula \"S\" \"B4\" \"=D1*2\")\n",
            "(formula \"S\" \"B5:B8\" \"=S.last!$A$1*ROW()\")\n",
            "(formula \"S\" \"D1\" \"=B8+1\")\n",
        ));
        fresh.step(0);
        assert_eq!(values(&m), values(&fresh), "the frame after");
        assert_eq!(number(&m, "S", "B4"), 34.0, "(16 + 1) * 2");
    }

    #[test]
    fn a_part_written_more_than_once_is_computed_once_a_run() {
        // KERNEL.27: a part a formula writes more than once in one branch,
        // counted by `TICK` where it is computed: by a run once a cell, beside
        // itself in a chain, as a call's operand and inside a larger part
        // written twice; at each cell alone, once a place; in a condition and
        // the branches of an `IF`, which take other cells, once a place by
        // either. The values the same. Runs of twelve, fewer than `SHORT`,
        // are evaluated cell by cell by their programs (KERNEL.44), so the
        // runs counted here are held to the run's own path.
        let rows = concat!(
            "(cell \"S\" \"A1:A12\" 3)\n",
            "(formula \"S\" \"B1:B12\" \"=(TICK()+A1)*2+(TICK()+A1)-MAX(TICK()+A1,0)\")\n",
            "(formula \"S\" \"C1:C12\" \"=IF(TICK()>0,TICK()+1,TICK()+2)\")\n",
            "(formula \"S\" \"D1:D12\" ",
            "\"=((TICK()+A1)*(TICK()+A1))-(TICK()+A1)*(TICK()+A1)+B1\")\n",
        );
        let frame = |by_runs: bool| {
            let mut lib = library::language();
            lib.register("TICK", tick);
            let mut m = Machine::load(rows, lib).unwrap_or_else(|e| panic!("{}", e.refusal.text));
            assert_eq!(m.shapes.cells_in_runs(), 36, "every cell in a run");
            m.by_runs = by_runs;
            m.lowered = false;
            let before = TICKS.with(Slot::get);
            assert!(m.step(0).done);
            (TICKS.with(Slot::get) - before, values(&m))
        };
        let ((runs, by_run), (alone, by_cell)) = (frame(true), frame(false));
        assert_eq!(by_run, by_cell, "the values");
        assert_eq!(number_in(&by_run, (0, 2, 4)), "Value(Number(8.0))");
        assert_eq!(
            (runs, alone),
            (12 + 24 + 12, 36 + 24 + 48),
            "TICK's calls in a frame, by runs and alone"
        );
    }

    /// A formula cell's value as [`values`] spells it.
    fn number_in(values: &BTreeMap<CellId, String>, id: CellId) -> &str {
        values.get(&id).map_or("none", String::as_str)
    }

    #[test]
    fn the_parts_a_run_computes_once_are_those_written_in_one_branch() {
        // KERNEL.27: the places a run takes a part's lane at, found once a
        // shape, counted as parts and places: a part not a literal, which the
        // run reaches (not a reference a strict call takes as one, nor inside
        // a call computed at each cell), in the same branch of every `IF`
        // around it; a place inside a later place of a larger part left out.
        let lib = library::language();
        let life = "IF(Clock!$B$1=1,IF(MOD(MOD(MOD(Clock!$B$3*1103515+ROW()*1597334+\
                    COLUMN()*1830421+12345,67108859)*MOD(Clock!$B$3*2012345+ROW()*1209487+\
                    COLUMN()*1949213+54321,67108859),67108859)*MOD(MOD(Clock!$B$3*1103515+\
                    ROW()*1597334+COLUMN()*1830421+12345,67108859)*MOD(Clock!$B$3*2012345+\
                    ROW()*1209487+COLUMN()*1949213+54321,67108859),67108859),67108859)\
                    <25165822,1,0),IF(OR(SUM(Screen.last!A1:C3)-Screen.last!B2=3,\
                    AND(Screen.last!B2=1,SUM(Screen.last!A1:C3)-Screen.last!B2=2)),1,0))";
        for (text, wants) in [
            ("A1+A1", (1, 2)),
            ("A1*2+A1-A1", (1, 3)),
            ("1+1", (0, 0)),
            ("\"x\"&\"x\"", (0, 0)),
            ("SUM(A1:B2)+SUM(A1:B2)", (1, 2)),
            ("SUM(A1,A1)", (0, 0)),
            ("ABS(A1)+ABS(A1)", (1, 2)),
            ("Rate+Rate", (1, 2)),
            ("IF(A1>0,A1,A1)", (0, 0)),
            ("IF(B1,A1+A1,0)", (1, 2)),
            ("IF(B1,A1,0)+IF(B1,A1,0)", (1, 2)),
            ("CHOOSE(1,A1+A1,A1)", (0, 0)),
            ("CHOOSE(1,A1,B1)*CHOOSE(1,A1,B1)", (1, 2)),
            ("(A1+B1)*(A1+B1)+A1", (2, 4)),
            ("(A1+B1)*(A1-B1)", (2, 4)),
            (
                "IF(OR(SUM(S.last!A1:C3)-S.last!B2=3,AND(S.last!B2=1,\
                 SUM(S.last!A1:C3)-S.last!B2=2)),1,0)",
                (2, 4),
            ),
            (life, (6, 12)),
        ] {
            let e = crate::calc::formula::parse(text).unwrap_or_else(|c| panic!("{text}: {c}"));
            let r = vector::Repeats::of(&lib, &e);
            assert_eq!((r.parts(), r.places()), wants, "{text}");
        }
    }

    /// Every twin's record and plane, whole, as a view reads them.
    fn twin_views(m: &Machine) -> Vec<(String, Vec<u8>)> {
        (m.sheets..m.book.sheets.len())
            .map(|t| {
                let name = m.sheet_name(t).expect("a twin");
                let record = grid(m, name, None);
                // A twin holds no formula, so its values are its record,
                // its cells left behind their slots read alike (KERNEL.37).
                assert_eq!(values_view(m, name, None), record, "{name}'s values");
                (record, plane(m, name, &whole(&m.book.sheets[t]).text()))
            })
            .collect()
    }

    /// A write a generated grid takes between two frames, or none: value
    /// cells changed, two at once to one of two values, a cell of the other
    /// sheet, a value over a formula, a cell added beside (KERNEL.20, slice
    /// 4); a formula over a part of the first bands (KERNEL.34).
    fn write_between(d: &mut Dice, frame: u32) -> Option<String> {
        Some(match d.below(9) {
            0..=2 => format!(
                "(cell \"A\" \"{}{}\" {})",
                d.pick(&["A", "B"]),
                1 + d.below(20),
                d.pick(&["0", "1", "2", "7", "\"x\"", "true"])
            ),
            // Two of a few cells at once, each to one of two values, so that
            // a part the same at every cell changes and changes back, alone
            // and beside another change.
            6 | 7 => format!(
                "(cell \"A\" \"A{}\" {})\n(cell \"A\" \"B{}\" {})",
                1 + d.below(4),
                d.pick(&["0", "7"]),
                1 + d.below(4),
                d.pick(&["0", "7"])
            ),
            3 => format!(
                "(cell \"B\" \"A{}\" {})",
                1 + d.below(20),
                d.pick(&["1", "3", "\"y\""])
            ),
            // A value over a formula, which plans the grid again.
            4 if frame % 3 == 0 => format!(
                "(cell \"A\" \"{}{}\" 2)",
                d.pick(&LETTERS[2..8]),
                1 + d.below(4)
            ),
            // A cell added beside, where references may point.
            5 if frame % 4 == 0 => format!("(cell \"A\" \"J{}\" 4)", 1 + d.below(20)),
            // A formula over a part of the first bands, a fill's first cell
            // among it at times, which plans the grid again and may give the
            // fill's first cell up (KERNEL.34).
            8 if frame % 2 == 1 => {
                let top = 1 + d.below(4);
                format!(
                    "(formula \"A\" \"{}{top}:{}{}\" \"=A.last!C1*2+$A$1\")",
                    d.pick(&LETTERS[2..5]),
                    d.pick(&LETTERS[5..8]),
                    top + d.below(3)
                )
            }
            _ => return None,
        })
    }

    #[test]
    fn a_twin_s_cells_made_when_read_are_the_cells_a_frame_s_end_copies() {
        // KERNEL.27: grids whose frames' ends leave the twins' cells behind
        // their slots, against the same grids copying them at every frame's
        // end, frame for frame: every twin's record and plane as a view reads
        // them, while a step yields and after; every value, twin cell,
        // reason, array and plane; the twins' cells made where they stand at
        // times, by the grid lent whole and by a frame's end that walks every
        // sheet, and the frames after; writes between frames. Over grids of
        // generated formulas, most of whose cells change every frame, and over
        // Life from soups of other seeds, whose frames' ends copy some cells
        // and not others, the Clock written before every step. Frames that
        // left a twin cell behind, twins lent whole made, and walks after such
        // a frame.
        let mut seen = (0, 0, 0);
        for seed in 1..=40u64 {
            let mut d = Dice(seed ^ 0x7A1C);
            let rows = generated_grid(&mut d);
            lazy_against_eager(&rows, &mut d, 8, &mut seen, write_between);
        }
        for seed in 1..=12u64 {
            let mut d = Dice(seed ^ 0x11FE);
            let soup = format!("ROW()*{}+COLUMN()*{}", 3 + d.below(20), 5 + d.below(20));
            let rows = LIFE.replace("ROW()*7+COLUMN()*13", &soup);
            lazy_against_eager(&rows, &mut d, 16, &mut seen, |_, frame| {
                Some(format!("(cell \"Clock\" \"B1\" {frame})"))
            });
        }
        let (behind, lent, walked) = seen;
        assert!(
            behind > 100 && lent > 20 && walked > 20,
            "{behind} frames left a twin cell behind, {lent} twins lent whole made and \
             {walked} walked: too few for the law to hold over"
        );
    }

    /// One grid stepped with its twins' cells left behind their slots and
    /// copied at every frame's end, `write` before each frame, held to each
    /// other as `a_twin_s_cells_made_when_read_are_the_cells_a_frame_s_end_copies`
    /// says; `seen` counts its frames that left a twin cell behind, its twins
    /// lent whole made, and its walks after such a frame.
    fn lazy_against_eager(
        rows: &str,
        d: &mut Dice,
        frames: u32,
        seen: &mut (usize, usize, usize),
        mut write: impl FnMut(&mut Dice, u32) -> Option<String>,
    ) {
        let mut lazy = load(rows);
        let mut eager = load(rows);
        eager.lazy = false;
        let first = lazy.sheet_name(0).expect("a sheet").to_string();
        for m in [&lazy, &eager] {
            m.view("plane", &first, None)
                .expect("a plane of the first sheet");
        }
        for frame in 1..=frames {
            let when = format!("frame {frame}\n{rows}");
            if let Some(w) = write(d, frame) {
                let (a, b) = (lazy.write(&w), eager.write(&w));
                assert_eq!(a.is_ok(), b.is_ok(), "{when}: {w}: {a:?} {b:?}");
            }
            let budget = 1 + d.below(200);
            if !lazy.step(budget).done {
                let views = twin_views(&lazy);
                assert_eq!(views, twin_views(&eager), "{when}: a view during a yield");
                while !lazy.step(budget).done {}
            }
            assert!(eager.step(0).done);
            let lagging = lazy.lagging.iter().any(|&l| l);
            seen.0 += usize::from(lagging);
            assert_eq!(
                twin_views(&lazy),
                twin_views(&eager),
                "{when}: the twins' views"
            );
            same_state(&lazy, &eager, &when);
            // A frame's end that walks every sheet, the twins' cells made
            // first; or the twins made by the grid lent whole.
            let walk = d.below(4) == 0;
            for m in [&mut lazy, &mut eager] {
                m.incremental = !walk;
            }
            seen.2 += usize::from(walk && lagging);
            if d.below(4) == 0 {
                seen.1 += usize::from(lagging);
                let rows_sheets = lazy.sheets;
                let made: Vec<String> = lazy.workbook().sheets[rows_sheets..]
                    .iter()
                    .map(|s| format!("{:?}", s.cells))
                    .collect();
                assert!(lazy.lagging.iter().all(|&l| !l), "{when}: every twin made");
                assert!(lazy.behind.iter().flatten().all(|&w| w == 0), "{when}");
                let copied: Vec<String> = eager.book.sheets[eager.sheets..]
                    .iter()
                    .map(|s| format!("{:?}", s.cells))
                    .collect();
                assert_eq!(made, copied, "{when}: the twins made");
            }
        }
    }
    #[test]
    fn a_value_cell_s_twin_is_copied_as_written_beside_formulas_made_from_their_slots() {
        // KERNEL.27: only a formula's twin cell is left behind its slot, which
        // holds its value exactly; a value cell written between frames is
        // copied into its twin where it stands, an error the language does
        // not name spelled as written, which its slot reads as `#VALUE!`. A
        // sheet's place and name asked of the machine, which never makes the
        // twins.
        let mut m = load(concat!(
            "(cell \"S\" \"A1:A8\" 1)\n",
            "(cell \"S\" \"C1:C20\" 0)\n",
            "(formula \"S\" \"B1:B8\" \"=S.last!B1+A1\")\n",
        ));
        for _ in 0..3 {
            m.step(0);
        }
        assert_eq!(m.write("(cell \"S\" \"A3\" (error \"#FOO!\"))"), Ok(1));
        m.step(0);
        assert!(m.lagging[0], "B's twin cells behind their slots");
        assert_eq!((m.find_sheet("s.LAST"), m.find_sheet("T")), (Some(1), None));
        assert_eq!((m.sheet_name(1), m.sheet_name(2)), (Some("S.last"), None));
        let wants = [
            ("A3", Content::Error("#FOO!".to_string())),
            ("B2", Content::Number(4.0)),
            ("B3", Content::Error("#VALUE!".to_string())),
        ];
        let record = grid(&m, "S.last", None);
        assert!(
            record.contains("(cell \"S.last\" \"A3\" (error \"#FOO!\"))"),
            "{record}"
        );
        for (at, wanted) in &wants {
            assert_eq!(content(&m, "S.last", at).as_ref(), Some(wanted), "{at}");
        }
        let book = m.workbook();
        for (at, wanted) in wants {
            let r = parse_a1_range(at).unwrap();
            assert_eq!(
                book.sheets[1].cells[&(r.top, r.left)].content,
                wanted,
                "{at}, made"
            );
        }
        assert!(!m.lagging[0]);
        assert_eq!(grid(&m, "S.last", None), record, "the record, made");
        // A formula whose twin cell is behind its slot, written over with a
        // value: copied into its twin as written at the next frame's end, and
        // no longer behind, while the cells beside it are.
        m.step(0);
        assert!(m.lagging[0], "B's twin cells behind again");
        assert_eq!(m.write("(cell \"S\" \"B5\" (error \"#BAR!\"))"), Ok(1));
        m.step(0);
        assert!(m.lagging[0], "B's other twin cells behind");
        let bar = Content::Error("#BAR!".to_string());
        assert_eq!(content(&m, "S.last", "B5").as_ref(), Some(&bar), "B5");
        assert_eq!(
            m.workbook().sheets[1].cells[&(5, 2)].content,
            bar,
            "B5, made"
        );
    }

    /// Two grids' kept state the same to the bit: every formula's value of
    /// the last complete frame, every twin's cells, extent and reasons, every
    /// twin's array, and every kept plane (KERNEL.20, slice 4).
    fn same_state(a: &Machine, b: &Machine, when: &str) {
        same_state_but_arrays(a, b, when);
        assert_eq!(
            format!("{:?}", a.dense),
            format!("{:?}", b.dense),
            "{when}: the arrays"
        );
    }

    /// `same_state` but for the twins' arrays.
    fn same_state_but_arrays(a: &Machine, b: &Machine, when: &str) {
        assert_eq!(values(a), values(b), "{when}: the values");
        for t in a.sheets..a.book.sheets.len() {
            let cells = |m: &Machine| format!("{:?}", twin_cells(m, t));
            let name = &a.book.sheets[t].name;
            assert_eq!(cells(a), cells(b), "{when}: {name}");
            assert_eq!(a.extents[t], b.extents[t], "{when}: {name}'s extent");
        }
        let reasons = |m: &Machine| -> BTreeMap<CellId, Reason> {
            m.twin_reasons
                .iter()
                .map(|(k, r)| (*k, r.clone()))
                .collect()
        };
        assert_eq!(reasons(a), reasons(b), "{when}: the reasons");
        assert_eq!(a.planes, b.planes, "{when}: the planes");
    }

    #[test]
    fn a_frame_that_computes_what_changed_is_the_frame_computed_whole() {
        // KERNEL.20, slice 4: grids of generated formulas, every kind of node
        // and reference, reading the last frame through their twins, stepped
        // computing only what changed, against the same grids computing every
        // cell and walking every sheet, frame for frame and bit for bit; a view
        // while a step yields showing the last complete frame; writes between
        // frames changing value cells, adding cells and changing formulas.
        // Frames that followed what changed and computed some cells and not
        // others, and the cells such frames computed of those they passed.
        computes_what_changed(1..=60);
    }

    #[test]
    fn a_frame_that_computes_what_changed_is_the_frame_computed_whole_every_sheet_indexed() {
        // KERNEL.35: the law above with every sheet's readers indexed, by
        // rows, by columns and by the axis an index finds, whatever their
        // count, so that every change is followed through an index.
        for (by_rows, seeds) in [
            (None, 61..=80),
            (Some(true), 81..=100),
            (Some(false), 101..=120),
        ] {
            let _every = crate::calc::reach::every::sheet_indexed(by_rows);
            computes_what_changed(seeds);
        }
    }

    /// `a_frame_that_computes_what_changed_is_the_frame_computed_whole` over
    /// the grids of some seeds.
    fn computes_what_changed(seeds: std::ops::RangeInclusive<u64>) {
        let (mut followed, mut frames, mut part, mut computed, mut passed) = (0, 0, 0, 0, 0);
        for seed in seeds {
            let mut d = Dice(seed ^ 0x5EED);
            let rows = generated_grid(&mut d);
            let mut inc = load(&rows);
            let mut whole = load(&rows);
            whole.incremental = false;
            for m in [&inc, &whole] {
                m.view("plane", "A", None).expect("a plane of A");
            }
            for frame in 1..=8 {
                let when = format!("seed {seed}, frame {frame}");
                let last = grid(&whole, "A", None) + &grid(&whole, "B", None);
                let budget = 1 + d.below(200);
                if !inc.step(budget).done {
                    let now = grid(&inc, "A", None) + &grid(&inc, "B", None);
                    assert_eq!(now, last, "{when}: a view during a yield\n{rows}");
                    while !inc.step(budget).done {}
                }
                assert!(whole.step(0).done);
                same_state(&inc, &whole, &format!("{when}\n{rows}"));
                frames += 1;
                if !inc.every {
                    let of = inc.shapes.order().len();
                    followed += 1;
                    part += usize::from(inc.computed > 0 && inc.computed < of);
                    computed += inc.computed;
                    passed += of;
                }
                let Some(write) = write_between(&mut d, frame) else {
                    continue;
                };
                let (a, b) = (inc.write(&write), whole.write(&write));
                assert_eq!(a.is_ok(), b.is_ok(), "{when}: {write}: {a:?} {b:?}");
            }
        }
        // The first two frames of every grid compute every cell, as do those
        // after a write that adds a cell or changes a formula and after a twin
        // gains or loses a cell, which a generated formula's reason often
        // makes.
        assert!(
            followed * 3 > frames && part * 3 > followed && computed * 2 < passed,
            "{followed} of {frames} frames followed what changed, {part} of them computing \
             some cells and not others, {computed} cells of {passed}: too few for the law \
             to hold over"
        );
    }

    /// The cells a frame computed, counted by `TICK`.
    fn ticked(m: &mut Machine) -> usize {
        let before = TICKS.with(Slot::get);
        assert!(m.step(0).done);
        TICKS.with(Slot::get) - before
    }

    #[test]
    fn a_frame_computes_the_cells_a_change_reaches_and_no_other() {
        // KERNEL.20, slice 4: each kind of read, a cell whose reference moves
        // with it, a part every `$` holds compared before its shape is
        // computed, a twin's cell a frame late, a defined name, a rectangle
        // every cell reads; the cells each frame computes counted exactly.
        let mut lib = library::language();
        lib.register("TICK", tick);
        let mut m = Machine::load(
            concat!(
                "(cell \"S\" \"A1:A10\" 1)\n",
                "(name \"Rate\" \"S!$A$1\")\n",
                "(formula \"S\" \"B1:B10\" \"=TICK()+A1\")\n",
                "(formula \"S\" \"C1:C10\" \"=TICK()+IF($A$1>100,1,0)\")\n",
                "(formula \"S\" \"D1:D10\" \"=TICK()+S.last!B1\")\n",
                "(formula \"S\" \"E1\" \"=TICK()+Rate\")\n",
                "(formula \"S\" \"F1:F10\" \"=TICK()+SUM($A$1:$A$3)\")\n",
            ),
            lib,
        )
        .unwrap_or_else(|e| panic!("{}", e.refusal.text));
        assert_eq!(ticked(&mut m), 41, "the first frame computes every cell");
        assert_eq!(
            ticked(&mut m),
            41,
            "and the second, its twins having gained cells"
        );
        assert_eq!(ticked(&mut m), 1, "then the name's alone");
        assert_eq!(m.write("(cell \"S\" \"A5\" 2)"), Ok(1));
        assert_eq!(ticked(&mut m), 2, "B5, and the name's");
        assert_eq!(ticked(&mut m), 2, "D5 a frame late, and the name's");
        assert_eq!(ticked(&mut m), 1);
        assert_eq!(m.write("(cell \"S\" \"A1\" 7)"), Ok(1));
        assert_eq!(
            ticked(&mut m),
            12,
            "B1, every F, the name's: C's part is FALSE still"
        );
        assert_eq!(ticked(&mut m), 2, "D1 a frame late, and the name's");
        assert_eq!(m.write("(cell \"S\" \"A1\" 500)"), Ok(1));
        assert_eq!(ticked(&mut m), 22, "B1, every C, every F, the name's");
        assert_eq!(number(&m, "S", "C4"), 2.0);
        assert_eq!(number(&m, "S", "F4"), 503.0);
        assert_eq!(number(&m, "S", "E1"), 501.0);
        assert_eq!(m.write("(cell \"S\" \"A1\" 500)"), Ok(0));
        assert_eq!(
            ticked(&mut m),
            2,
            "D1 a frame late, and the name's: an equal write writes nothing"
        );
        assert_eq!(ticked(&mut m), 1);
    }

    /// The formula cells whose references, moved to where each cell stands,
    /// read a cell: what a change to it must reach, found by moving every
    /// reference of every cell's parse (`graph::areas_at`) and none of the
    /// reach's code.
    fn readers_of(m: &Machine, changed: CellId) -> BTreeSet<CellId> {
        let (sheet, row, col) = changed;
        let mut out = BTreeSet::new();
        for &id in m.shapes.order() {
            let p = m.shapes.place(id).expect("placed");
            let Ok(e) = &m.shapes.shape(p.shape).expect("a shape").expr else {
                continue;
            };
            e.each_read(&mut |r| {
                let crate::calc::formula::Read::Ref(re) = r else {
                    return;
                };
                let Ok(areas) = crate::calc::graph::areas_at(&m.book, id.0, re, p.d_row, p.d_col)
                else {
                    return;
                };
                if areas.iter().any(|a| {
                    a.sheet == sheet
                        && a.top <= row
                        && row <= a.bottom
                        && a.left <= col
                        && col <= a.right
                }) {
                    out.insert(id);
                }
            });
        }
        out
    }

    /// The formula cells the plan's reach reaches from a changed cell.
    fn reached(m: &Machine, changed: CellId) -> BTreeSet<CellId> {
        use crate::calc::reach::Reached;
        let order = m.shapes.order();
        let mut out = BTreeSet::new();
        m.shapes.reach().each(changed, &mut |r| match r {
            Reached::Cell(pos) => {
                out.insert(order[pos]);
            }
            Reached::Shape(s) | Reached::Parts(s) => {
                out.extend(
                    order
                        .iter()
                        .filter(|&&id| m.shapes.place(id).map(|p| p.shape) == Some(s)),
                );
            }
        });
        out
    }

    #[test]
    fn a_row_of_single_cells_reaches_what_each_single_reached() {
        // KERNEL.33: a shape's single cells side by side in one row are read
        // as one span of the row where every cell of the shape keeps the span
        // on the sheet: Sand's five neighbours; the row from the sheet's first
        // column; the row at its last column, where a cell to the right moves
        // one single off the sheet and a span would move off it whole, so
        // that no span is made; a row with a gap, whose sides are spans apart;
        // and single cells down a diagonal, of no row. The value sheet's
        // readers are the patterns a changed cell of it tries, the spans made
        // counted there; and every cell of the sheet's window, changed,
        // reaches exactly the formula cells whose moved references read it.
        for (rows, readers, (r0, r1, c0, c1)) in [
            (
                "(formula \"S\" \"C3:H9\" \"=T!B2+T!C2+T!D2+T!E2+T!F2\")",
                1,
                (1, 12, 1, 12),
            ),
            (
                "(formula \"S\" \"A3:E6\" \"=T!A2*T!B2-T!C2\")",
                1,
                (1, 8, 1, 9),
            ),
            (
                "(formula \"S\" \"A3:C4\" \"=T!XFB2+T!XFC2+T!XFD2\")",
                3,
                (1, 5, 16_378, 16_384),
            ),
            (
                "(formula \"S\" \"B4:F7\" \"=T!B2+T!D2+T!E2-T!G2\")",
                3,
                (1, 10, 1, 14),
            ),
            (
                "(formula \"S\" \"B4:E6\" \"=T!C2+T!D3+T!E4\")",
                3,
                (1, 10, 1, 10),
            ),
        ] {
            let m = load(&format!("(cell \"T\" \"A1\" 1)\n{rows}\n"));
            let t = m.book.find_sheet("T").expect("the value sheet");
            assert_eq!(m.shapes.reach().readers_on(t), readers, "{rows}");
            for row in r0..=r1 {
                for col in c0..=c1 {
                    let changed = (t, row, col);
                    assert_eq!(
                        reached(&m, changed),
                        readers_of(&m, changed),
                        "{rows}: {changed:?}"
                    );
                }
            }
        }
    }

    #[test]
    fn a_range_held_at_one_corner_moves_its_other_at_each_cell() {
        // KERNEL.33: a plan finds a reference every `$` holds once for its
        // whole shape, and moves one held at a corner alone at each cell. A
        // running sum of a chain whose cells stand at rising levels reads a
        // range held at its first corner: each of its cells is put after the
        // last cell its own range reads, so the first frame sums each prefix.
        let mut m = load(concat!(
            "(sheet \"A\" visible)\n",
            "(sheet \"B\" visible)\n",
            "(formula \"A\" \"A1\" \"=1\")\n",
            "(formula \"A\" \"A2:A6\" \"=A1+1\")\n",
            "(formula \"B\" \"A1:A5\" \"=SUM(A!$A$1:A1)\")\n",
        ));
        while !m.step(usize::MAX).done {}
        for (row, sum) in [(1, 1.0), (2, 3.0), (3, 6.0), (4, 10.0), (5, 15.0)] {
            assert_eq!(number(&m, "B", &format!("A{row}")), sum, "B!A{row}");
        }
    }

    #[test]
    fn a_range_over_formulas_that_stand_apart_reads_them_all() {
        // KERNEL.33: a plan finds the formula cells an area holds through an
        // array where a sheet's formulas stand close together, and through
        // the sheet's cells where they stand apart, as here, at its corners
        // and between. The cell of the first sheet that sums them is put
        // after them in the order, so its first frame reads all three.
        let mut m = load(concat!(
            "(sheet \"A\" visible)\n",
            "(sheet \"B\" visible)\n",
            "(formula \"A\" \"A1\" \"=SUM(B!A1:XFD1048576)\")\n",
            "(formula \"B\" \"A1\" \"=1\")\n",
            "(formula \"B\" \"M500\" \"=4\")\n",
            "(formula \"B\" \"XFD1048576\" \"=2\")\n",
        ));
        while !m.step(usize::MAX).done {}
        assert_eq!(number(&m, "A", "A1"), 7.0);
    }

    #[test]
    fn a_part_is_computed_whenever_a_cell_it_reads_changes() {
        // KERNEL.20, slice 4, found by Alonzo's input card: a shape computed
        // whole because a cell it reads every `$` held changed, in the frame a
        // part of it changed too; then the part changing back alone. The part
        // kept is the one its cells were computed with, so the change back is
        // seen.
        let mut m = load(concat!(
            "(cell \"S\" \"A1\" 1)\n",
            "(cell \"S\" \"A2\" 0)\n",
            "(formula \"S\" \"C1:C4\" \"=IF(ROW()=$A$1,IF($A$2=1,12,15),0)\")\n",
        ));
        m.step(0);
        m.step(0);
        assert_eq!(number(&m, "S", "C1"), 15.0);
        // The row and the press in one write: the shape computed whole.
        assert_eq!(
            m.write("(cell \"S\" \"A1\" 3)\n(cell \"S\" \"A2\" 1)"),
            Ok(2)
        );
        m.step(0);
        assert_eq!(number(&m, "S", "C3"), 12.0);
        m.step(0);
        // The press let go alone: `$A$2=1` FALSE again, as it was before.
        assert_eq!(m.write("(cell \"S\" \"A2\" 0)"), Ok(1));
        m.step(0);
        assert_eq!(number(&m, "S", "C3"), 15.0, "the press let go is seen");
        assert_eq!(number(&m, "S", "C1"), 0.0);
    }

    #[test]
    fn a_part_reading_a_formula_cell_reaches_its_whole_shape() {
        // KERNEL.20, slice 4: a part the same at every cell that reads a
        // formula cell, whose value is final only once the frame has computed
        // it, is never computed at the frame's start: a change to any cell it
        // reads computes the whole shape. Computed there, it would keep the
        // formula's last value, and a later change of its value cell landing
        // the part back on what was kept would pass unseen.
        // Its cells named on the formula's own sheet, and on a sheet named.
        let mut m = load(concat!(
            "(cell \"S\" \"A1\" 1)\n",
            "(cell \"S\" \"B1\" 1)\n",
            "(formula \"S\" \"X1\" \"=B1\")\n",
            "(formula \"S\" \"C1:C4\" \"=($X$1+$A$1)*10+ROW()\")\n",
            "(formula \"S\" \"D1:D4\" \"=(S!$X$1+S!$A$1)*10+ROW()\")\n",
        ));
        m.step(0);
        m.step(0);
        for col in ["C", "D"] {
            assert_eq!(number(&m, "S", &format!("{col}2")), 22.0, "{col}");
        }
        assert_eq!(m.write("(cell \"S\" \"B1\" 2)"), Ok(1));
        m.step(0);
        for col in ["C", "D"] {
            let at = format!("{col}2");
            assert_eq!(number(&m, "S", &at), 32.0, "{col}: X1 changed in the frame");
        }
        assert_eq!(m.write("(cell \"S\" \"A1\" 0)"), Ok(1));
        m.step(0);
        for col in ["C", "D"] {
            let at = format!("{col}2");
            assert_eq!(number(&m, "S", &at), 22.0, "{col}: the part back on 2 + 0");
        }
    }

    #[test]
    fn a_twin_s_reason_changed_reaches_its_readers_however_the_twins_are_filled() {
        // KERNEL.20, slice 4: a formula not computed, whose reason changes
        // and whose twin cell stays empty, reaches the cell reading that twin
        // cell the next frame, whether the frame's end walks every sheet (most
        // cells changed) or copies the cells that changed (a few did).
        for value_cells in [0, 20] {
            let grid = format!(
                "(cell \"S\" \"A1\" 1)\n{}(formula \"S\" \"C1\" \"=IF(A1=1,D1:D2,E1:E2)\")\n\
                 (formula \"S\" \"F1\" \"=S.last!C1\")\n",
                (1..=value_cells)
                    .map(|r| format!("(cell \"S\" \"H{r}\" {r})\n"))
                    .collect::<String>()
            );
            let mut m = load(&grid);
            for _ in 0..3 {
                m.step(0);
            }
            let reason = |m: &Machine, addr: &str| match at(m, "S", addr) {
                Some(Computed::NotComputed(Reason::Construct(c))) => c,
                other => panic!("{addr}: {other:?}"),
            };
            assert_eq!(reason(&m, "F1"), "D1:D2", "{value_cells} value cells");
            assert_eq!(m.write("(cell \"S\" \"A1\" 2)"), Ok(1));
            m.step(0);
            assert_eq!(reason(&m, "C1"), "E1:E2", "{value_cells} value cells");
            m.step(0);
            assert_eq!(reason(&m, "F1"), "E1:E2", "{value_cells} value cells");
        }
    }

    #[test]
    fn lifes_rule_under_a_frame_counter_computes_near_what_changed() {
        // KERNEL.20, slice 4: Life's rule as the cartridges write it, its seed
        // read while a frame counter every cell reads holds 1, over a board
        // with a blinker. The counter changes every frame and the part it is
        // read through does not after the second, so each frame from the
        // third computes the 21 cells around the blinker's four changes.
        let mut lib = library::language();
        lib.register("TICK", tick);
        let mut m = Machine::load(
            concat!(
                "(cell \"S\" \"A1\" 0)\n",
                "(cell \"Seed\" \"E4:E6\" 1)\n",
                "(formula \"Board\" \"B2:O15\" \"=IF(S!$A$1=1,Seed!B2,TICK()*0+",
                "IF(OR(SUM(Board.last!A1:C3)-Board.last!B2=3,AND(Board.last!B2=1,",
                "SUM(Board.last!A1:C3)-Board.last!B2=2)),1,0))\")\n",
            ),
            lib,
        )
        .unwrap_or_else(|e| panic!("{}", e.refusal.text));
        let mut computed = Vec::new();
        for frame in 1..=8u32 {
            assert_eq!(m.write(&format!("(cell \"S\" \"A1\" {frame})")), Ok(1));
            computed.push(ticked(&mut m));
            let upright = frame % 2 == 1;
            for (addr, alive) in [
                ("E4", upright),
                ("E6", upright),
                ("D5", !upright),
                ("F5", !upright),
            ] {
                assert_eq!(
                    number(&m, "Board", addr),
                    f64::from(u8::from(alive)),
                    "frame {frame}: {addr}"
                );
            }
            assert_eq!(number(&m, "Board", "E5"), 1.0, "frame {frame}");
        }
        assert_eq!(computed, [0, 196, 21, 21, 21, 21, 21, 21]);
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
        let s = m.find_sheet(sheet).unwrap();
        m.planes.get(s).is_some_and(Option::is_some)
    }

    /// The plane a view answers held to the walk, the reference, over each
    /// window; then the sheet's plane kept, or not, as said, so that stale
    /// bytes fail on the bytes and a plane never kept fails on the flag.
    fn walked(m: &Machine, sheet: &str, windows: &[&str], kept_now: bool) {
        let was_kept = kept(m, sheet);
        for w in windows {
            let window = window_of(&m.book, sheet, Some(w)).unwrap();
            let walk = plane::bytes(&m.book, m, &window);
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
        again.walk();
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
        assert!(
            e.text
                .contains(" values, those rows without the formulas' texts, "),
            "{}",
            e.text
        );
        assert!(
            e.text
                .contains(" changes, the values' rows of the cells the last frame changed, "),
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
            m.view("values", "Nowhere", None).unwrap_err().id,
            "view-sheet-unknown"
        );
        assert_eq!(
            m.view("plane", "Board", Some("A:A")).unwrap_err().id,
            "view-window-not-a-range"
        );
        assert_eq!(
            m.view("values", "Board", Some("A:A")).unwrap_err().id,
            "view-window-not-a-range"
        );
        // KERNEL.42: the changes refused as the values are.
        assert_eq!(
            m.view("changes", "Nowhere", None).unwrap_err().id,
            "view-sheet-unknown"
        );
        assert_eq!(
            m.view("changes", "Board", Some("A:A")).unwrap_err().id,
            "view-window-not-a-range"
        );
        let e = m.view("plane", "Board", Some("A1:XFD1048576")).unwrap_err();
        assert_eq!(e.id, "view-window-too-large");
        // The twins are viewed by name; the load filled this one with the
        // grid's values, the one value cell, and no formula has one yet.
        let twin = grid(&m, "board.LAST", None);
        assert!(twin.contains("(window \"Board.last\" \"D1\")"), "{twin}");
        assert!(twin.contains("(cell \"Board.last\" \"D1\" 5)"), "{twin}");
        assert_eq!(values_view(&m, "board.LAST", None), twin);
    }

    #[test]
    fn a_steady_frame_s_values_walk_no_sheet() {
        // KERNEL.37: a running grid's values read a fill's cell through its
        // value, and the sheet's extent as the machine keeps it, so a steady
        // frame's view walks no sheet. Before the first frame, no cell having
        // a value yet, the masters are walked once; the grid walks them at
        // every view, to move their texts. The values are the grid less its
        // formula lines throughout, over a window of the fill's cells and
        // over the whole sheet.
        let mut m = load(LIFE);
        let screen = m.find_sheet("Screen").expect("the Screen");
        let cells = m.book.sheets[screen].cells.len() as u64;
        let counted = |m: &Machine, projection: &str, window: Option<&str>| {
            crate::sheet::work::begin();
            let text = match m.view(projection, "Screen", window) {
                Ok(Viewed::Text(t)) => t,
                other => panic!("{other:?}"),
            };
            (text, crate::sheet::work::end())
        };
        for frame in 0..=3u32 {
            if frame > 0 {
                m.write(&format!("(cell \"Clock\" \"B1\" {frame})"))
                    .expect("the Clock");
                assert!(m.step(0).done);
            }
            for window in [Some("C3:J10"), None] {
                let (values, v) = counted(&m, "values", window);
                let (grid, g) = counted(&m, "grid", window);
                let when = format!("frame {frame}, {window:?}");
                assert_eq!(values, less_formulas(&grid), "{when}");
                let walked = if frame == 0 { cells } else { 0 };
                assert_eq!((v.walked, v.extents), (walked, 0), "{when}: the values");
                assert_eq!((g.walked, g.extents), (cells, 0), "{when}: the grid");
            }
        }
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
        // A record whose formats hold no code prints no text row (KERNEL.47).
        assert!(!saved.contains("(text "), "{saved}");
    }

    /// A grid whose formats hold codes (KERNEL.47): a number's, a sign's
    /// glyph, a text's fourth section, a look's code over its fill.
    const CODED: &str = "\
(sheet \"P\" visible)
(format 1 none \"0.00\" nowrap)
(format 2 none \"▲0.00;▼0.00;=0.00\" nowrap)
(format 3 none \"0;-0;0;\\\"<\\\"@\\\">\\\"\" nowrap)
(format 4 \"F7F4FC\" \"#,##0\" nowrap)
(look \"P\" 7 4)
(look \"P\" -0 2)
(cell \"P\" \"A1\" 1)
(style \"P\" \"A1\" 1)
(cell \"P\" \"A2\" 1.005)
(style \"P\" \"A2\" 1)
(formula \"P\" \"B1\" \"=A1-A2\")
(style \"P\" \"B1\" 2)
(cell \"P\" \"C1\" \"abc\")
(style \"P\" \"C1\" 3)
(cell \"P\" \"C2\" true)
(style \"P\" \"C2\" 1)
(cell \"P\" \"D1\" 7)
(cell \"P\" \"D2\" 12)
(cell \"P\" \"D3\" 0)
(style \"P\" \"D3\" 1)
";

    #[test]
    fn a_code_draws_a_number_s_text_beside_its_value() {
        let mut m = load(CODED);
        let g = grid(&m, "P", None);
        for rows in [
            "(format 1 none \"0.00\" nowrap)",
            "(format 2 none \"▲0.00;▼0.00;=0.00\" nowrap)",
            "(format 3 none \"0;-0;0;\\\"<\\\"@\\\">\\\"\" nowrap)",
            "(format 4 \"F7F4FC\" \"#,##0\" nowrap)",
            // After the cell's value and before its style.
            "(cell \"P\" \"A1\" 1)\n(text \"P\" \"A1\" \"1.00\")\n(style \"P\" \"A1\" 1)\n",
            // Excel's rounding: fifteen digits, then half away from zero.
            "(text \"P\" \"A2\" \"1.01\")",
            // A text under a code's fourth section.
            "(text \"P\" \"C1\" \"<abc>\")",
            // A look's format in place of the cell's own, 0 meeting -0.
            "(cell \"P\" \"D1\" 7)\n(text \"P\" \"D1\" \"7\")\n",
            "(text \"P\" \"D3\" \"=0.00\")",
        ] {
            assert!(g.contains(rows), "{rows}\n{g}");
        }
        // No text for a formula before its first value, a truth value, or
        // a number under general.
        for addr in ["B1", "C2", "D2"] {
            assert!(
                !g.contains(&format!("(text \"P\" \"{addr}\"")),
                "{addr}\n{g}"
            );
        }
        assert!(m.step(0).done);
        let g = grid(&m, "P", None);
        // 1 - 1.005 is -0.004999999999999893: the negative's section, its
        // fifteen digits rounding to 0.00.
        assert!(g.contains("(text \"P\" \"B1\" \"▼0.00\")"), "{g}");
        assert_eq!(values_view(&m, "P", None), less_formulas(&g));
        // The load skips a text row and the record prints it again.
        let saved = save(&m);
        assert_eq!(save(&load(&saved)), saved, "row for row");
    }

    #[test]
    fn a_text_row_is_the_code_s_and_never_the_rows() {
        // A text row the rows wrote is skipped; the code makes the text.
        let m = load(
            "(format 1 none \"0.00\" nowrap)\n(cell \"P\" \"A1\" 2)\n(text \"P\" \"A1\" \"nonsense\")\n(style \"P\" \"A1\" 1)\n",
        );
        let g = grid(&m, "P", None);
        assert!(g.contains("(text \"P\" \"A1\" \"2.00\")"), "{g}");
        assert!(!g.contains("nonsense"), "{g}");
    }

    #[test]
    fn a_code_this_version_does_not_draw_is_refused_by_name() {
        let e = refused("(sheet \"P\" visible)\n(format 1 none \"[Red]0.00\" nowrap)\n");
        assert_eq!(e.refusal.id, "grid-format-code");
        assert_eq!(e.line, 2);
        assert!(
            e.refusal.text.contains("\"[Red]0.00\""),
            "{}",
            e.refusal.text
        );
        assert!(e.refusal.text.contains("a color"), "{}", e.refusal.text);
        let e = refused("(format 1 none \"General\" nowrap)\n");
        assert_eq!(e.refusal.id, "grid-format-code");
        assert!(
            e.refusal.text.contains("write it general"),
            "{}",
            e.refusal.text
        );
        let e = refused("(format 1 none \"yyyy-mm-dd\" nowrap)\n");
        assert_eq!(e.refusal.id, "grid-format-code");
        // A bare symbol but the two is no code: the row's shape says one.
        let e = refused("(format 1 none dollars nowrap)\n");
        assert_eq!(e.refusal.id, "grid-row-malformed");
        assert!(e.refusal.text.contains("\"<code>\""), "{}", e.refusal.text);
    }

    #[test]
    fn load_is_the_inverse_of_view_on_the_view_goldens() {
        // The goldens of oracle 11 whose window holds the sheet's extent;
        // the sixth, a window clipped inside its sheet, holds the window's
        // cells alone, so its extent row is not the loaded cells'. Each
        // golden less its formula lines is the values of the same window
        // (KERNEL.37), the values' goldens without a copy of their own.
        let goldens: &[(&str, &str, &str)] = &[
            (
                include_str!("../../testdata/view/fixture_frazaro.vla"),
                "Frazaro",
                "B1:C11",
            ),
            (
                include_str!("../../testdata/view/fixture_output.vla"),
                "Output",
                "A1:F20",
            ),
            (
                include_str!("../../testdata/view/fixture_data.vla"),
                "data",
                "A1",
            ),
            (
                include_str!("../../testdata/view/into_output.vla"),
                "Output",
                "A1:B1",
            ),
            (
                include_str!("../../testdata/view/into_checks.vla"),
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
            let values = view_text_values(&draft.workbook, &w, &Held(draft.values.clone()));
            assert_eq!(
                values,
                less_formulas(&golden),
                "{sheet} {window}: the values"
            );
            assert!(!values.contains("(formula "), "{sheet} {window}");
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
        let fixture = include_str!("../../testdata/refers.txt").replace("\r\n", "\n");
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
        let book = &m.book;
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

    // ---- KERNEL.34: a formula's write in time with its cells, and the grid
    // a write leaves the grid a load of it makes ----

    /// The plan a load of `m`'s grid makes, spelled whole.
    fn plan_of_a_load(m: &Machine) -> String {
        Shapes::of(&m.book, m.sheets, &m.library, &m.names, &m.extents).spelled()
    }

    /// A machine loaded from the grid `m` holds, its rows' sheets copied
    /// whole, fills and all, and made by `Machine::new`, as a load of the
    /// same rows makes one.
    fn loaded_again(m: &Machine) -> Machine {
        let mut workbook = m.book.clone();
        workbook.sheets.truncate(m.sheets);
        let draft = Draft {
            workbook,
            values: HashMap::new(),
            lines: HashMap::new(),
            cells: m.cells,
        };
        Machine::new(draft, library::language())
            .unwrap_or_else(|e| panic!("line {}: {}", e.line, e.refusal.text))
    }

    /// Every formula cell of the rows' sheets with its text as the formula
    /// bar shows it, a fill's cells through their master.
    fn formula_texts(m: &Machine) -> BTreeMap<CellId, String> {
        let mut out = BTreeMap::new();
        for (s, sheet) in m.book.sheets[..m.sheets].iter().enumerate() {
            let masters = shared_masters(sheet);
            for (&(row, col), cell) in &sheet.cells {
                if let Some(text) = formula_text(&cell.content, row, col, &masters) {
                    out.insert((s, row, col), text);
                }
            }
        }
        out
    }

    /// Every cell of the rows' sheets with its format.
    fn styles(m: &Machine) -> BTreeMap<CellId, u32> {
        let mut out = BTreeMap::new();
        for (s, sheet) in m.book.sheets[..m.sheets].iter().enumerate() {
            for (&(row, col), cell) in &sheet.cells {
                out.insert((s, row, col), cell.style);
            }
        }
        out
    }

    /// The cells a write's rows name.
    fn named_cells(m: &Machine, write: &str) -> BTreeSet<CellId> {
        let mut out = BTreeSet::new();
        for form in read_forms(write).unwrap_or_default() {
            let Some((_, [Form::Str(sheet), Form::Str(addr), _], _)) = load::row_parts(&form)
            else {
                continue;
            };
            let (Some(s), Some(r)) = (m.find_sheet(sheet), parse_a1_range(addr)) else {
                continue;
            };
            for row in r.top..=r.bottom {
                for col in r.left..=r.right {
                    out.insert((s, row, col));
                }
            }
        }
        out
    }

    /// One write held to a load (KERNEL.34's oracle): every cell it does not
    /// name reads as it did, and every cell keeps its format; the plan it
    /// leaves is the plan a load of its grid makes, part for part; the frame
    /// after it is `Calc`'s over the grid as the frame began; and its census
    /// is the census of the same grid loaded and stepped. `Calc` reads a twin as a
    /// sheet and sees no reason a cell not computed left beside its twin's
    /// cell, which a formula reading that cell inherits, so the frame is held
    /// to it only when no twin holds one, which the second answer says. A
    /// refused write leaves the record, the plan and the census as they
    /// were, and steps no frame. The write's answer.
    fn as_a_load(m: &mut Machine, write: &str, when: &str) -> (Result<usize, LineRefusal>, bool) {
        let before = (save(m), m.shapes.spelled(), m.census());
        let (texts, styles) = (formula_texts(m), styles(m));
        let answer = m.write(write);
        if answer.is_err() {
            let after = (save(m), m.shapes.spelled(), m.census());
            assert_eq!(after, before, "{when}: {write}: a refused write");
            return (answer, false);
        }
        // Every cell the write does not name reads as it did, and every cell
        // that was keeps its format.
        let named = named_cells(m, write);
        let others = |t: BTreeMap<CellId, String>| -> BTreeMap<CellId, String> {
            t.into_iter()
                .filter(|(id, _)| !named.contains(id))
                .collect()
        };
        assert_eq!(
            others(formula_texts(m)),
            others(texts),
            "{when}: {write}: every other cell's formula"
        );
        let now = self::styles(m);
        for (id, style) in &styles {
            assert_eq!(now.get(id), Some(style), "{when}: {write}: {id:?}'s format");
        }
        assert_eq!(
            m.shapes.spelled(),
            plan_of_a_load(m),
            "{when}: {write}: the plan"
        );
        let book = m.workbook().clone();
        let held = m.twin_reasons.is_empty();
        assert!(m.step(0).done);
        if held {
            let calc = crate::calc::Calc::run(&book, &m.library);
            let by_calc: BTreeMap<CellId, String> = m
                .shapes
                .order()
                .iter()
                .map(|id| (*id, format!("{:?}", calc.computed(*id).expect("computed"))))
                .collect();
            assert_eq!(values(m), by_calc, "{when}: {write}: the values");
        }
        let mut again = loaded_again(m);
        assert!(again.step(0).done);
        assert_eq!(m.census(), again.census(), "{when}: {write}: the census");
        (answer, held)
    }

    #[test]
    fn a_write_leaves_the_grid_a_load_of_it_makes() {
        // KERNEL.34: each kind of write held to a load of the grid it leaves
        // (`as_a_load`), with the cells it counts: a row inside a fill; the
        // fill's first row, then its first column, which give its first cell
        // up twice; a cell inside it; a value over its first cell; a fill's
        // own text again, which writes nothing, and another; plain formulas
        // of one R1C1 text written as a fill, a shape's first cell spelled
        // otherwise, then the fill in both spellings over plain cells and over
        // a fill; a fill whose text a release would move off the sheet; the
        // second sheet's fill, which reads the first; a plain cell an earlier
        // row of the same write placed; a write refused for a cycle, one
        // refused after giving a fill's first cell up, and one refused at its
        // second row. A cell of the fill keeps its format.
        let mut m = load(concat!(
            "(format 1 \"F7F4FC\" text nowrap)\n",
            "(cell \"S\" \"A1:A12\" 2)\n",
            "(formula \"S\" \"B1:D8\" \"=A1*10+S.last!B1\")\n",
            "(style \"S\" \"C1\" 1)\n",
            "(formula \"S\" \"E1:E8\" \"=SUM(B1:D1)\")\n",
            "(formula \"S\" \"F1\" \"=A1+1\")\n",
            "(formula \"S\" \"F2\" \"=A2+1\")\n",
            "(formula \"S\" \"H1:I3\" \"=XFD1\")\n",
            "(cell \"T\" \"A1:B4\" 3)\n",
            "(formula \"T\" \"C1:D4\" \"=S!B1+A1\")\n",
        ));
        assert!(m.step(0).done);
        assert!(m.step(0).done);
        let writes: [(&str, Result<usize, &str>); 19] = [
            ("(formula \"S\" \"B3:D3\" \"=A3*10+S.last!B3+1\")", Ok(3)),
            ("(formula \"S\" \"B1:D1\" \"=A1*20\")", Ok(3)),
            ("(formula \"S\" \"B1:B8\" \"=A1*30\")", Ok(8)),
            ("(formula \"S\" \"C4\" \"=1\")", Ok(1)),
            ("(cell \"S\" \"C2\" 5)", Ok(1)),
            ("(formula \"S\" \"E1:E8\" \"=SUM(B1:D1)\")", Ok(0)),
            ("(formula \"S\" \"E1:E8\" \"=SUM(B1:D1)+0\")", Ok(8)),
            ("(formula \"S\" \"F1:F2\" \"=A1+1\")", Ok(0)),
            ("(formula \"S\" \"F1\" \"=a1+1\")", Ok(1)),
            ("(formula \"S\" \"F1:F2\" \"=a1+1\")", Ok(0)),
            ("(formula \"S\" \"F1:F2\" \"=A1+1\")", Ok(1)),
            ("(formula \"S\" \"F1:F2\" \"=a1+1\")", Ok(1)),
            ("(cell \"S\" \"H1\" 0)", Ok(1)),
            ("(formula \"T\" \"C1:C4\" \"=S!B1*2\")", Ok(4)),
            (
                "(formula \"S\" \"K1\" \"=A1\")\n(formula \"S\" \"K1:K2\" \"=A1\")",
                Ok(2),
            ),
            ("(formula \"S\" \"A1\" \"=B1\")", Err("calc-cycle")),
            ("(formula \"T\" \"D1\" \"=D1\")", Err("calc-cycle")),
            (
                "(formula \"S\" \"J1\" \"=1\")\n(cell \"Nowhere\" \"A1\" 1)",
                Err("grid-sheet-unknown"),
            ),
            ("(formula \"T\" \"D2:D4\" \"=S!C2+B2\")", Ok(0)),
        ];
        for (write, wants) in writes {
            let (got, held) = as_a_load(&mut m, write, "S and T");
            assert!(held || got.is_err(), "{write}: the frame held to Calc's");
            assert_eq!(
                got.map_err(|e| e.refusal.id),
                wants.map_err(str::to_string),
                "{write}"
            );
        }
        assert_eq!(
            m.book.sheets[0].cells[&(1, 3)].style,
            1,
            "C1 keeps its format"
        );
    }

    #[test]
    fn the_same_text_test_compares_each_cell_where_a_moved_text_leaves_the_sheet() {
        // KERNEL.34: a reference that moves off the sheet is `#REF!` at every
        // cell past the edge, whatever it was, so two texts whose references
        // differ by a column read alike from the second cell past the edge
        // on and differently before: one cell cannot stand for the rest, and
        // the test compares each. A plain formula `#REF!` reads as the written
        // text moved off the sheet reads, though its R1C1 text is not the
        // written text's. Spelled otherwise, a fill's first cell differs where
        // its other cells do not.
        let m = load(concat!(
            "(formula \"S\" \"C1:H1\" \"=XFD1\")\n",
            "(formula \"S\" \"C2:E2\" \"=A2+1\")\n",
            "(formula \"S\" \"C3\" \"=XFD3\")\n",
            "(formula \"S\" \"D3\" \"=#REF!\")\n",
        ));
        let range = |a: &str| parse_a1_range(a).expect("a range");
        assert_eq!(same_text(&m, "S", range("C1:H1"), "=XFC1", "the edge"), 2);
        assert_eq!(same_text(&m, "S", range("C1:H1"), "=XFD1", "the edge"), 0);
        assert_eq!(same_text(&m, "S", range("D1:H1"), "=#REF!", "the edge"), 0);
        assert_eq!(same_text(&m, "S", range("C3:D3"), "=XFD3", "plain"), 0);
        assert_eq!(same_text(&m, "S", range("C2:E2"), "=a2+1", "spelled"), 1);
        assert_eq!(same_text(&m, "S", range("D2:E2"), "=b2+1", "spelled"), 1);
    }

    /// A generated grid's formula rows: each one's sheet, range and text,
    /// its `=` kept.
    fn formula_rows_of(rows: &str) -> Vec<(String, A1Range, String)> {
        read_forms(rows)
            .expect("the rows read")
            .iter()
            .filter_map(|form| match load::row_parts(form)? {
                ("formula", [Form::Str(sheet), Form::Str(addr), Form::Str(text)], _) => {
                    Some((sheet.clone(), parse_a1_range(addr)?, text.clone()))
                }
                _ => None,
            })
            .collect()
    }

    /// The cells the same-text test finds before one write, held to the cells
    /// comparing every cell finds (KERNEL.34): how many.
    fn same_text(m: &Machine, sheet: &str, range: A1Range, text: &str, when: &str) -> usize {
        let si = m.find_sheet(sheet).expect("the sheet");
        let body = text.strip_prefix('=').expect("a formula");
        let mut found = m.differing(si, range, body, &Log::default());
        let mut every = m.differing_by_cell(si, range, body);
        found.sort_unstable();
        every.sort_unstable();
        assert_eq!(found, every, "{when}: {sheet}!{} {text}", range.text());
        found.len()
    }

    #[test]
    fn the_same_text_test_finds_the_cells_comparing_every_cell_finds() {
        // KERNEL.34: the test reading each formula once, against every cell's
        // text and the written text moved to it and compared, over generated
        // grids: each formula row's own text over its range (the same), over
        // its range from a cell inside, moved there (the same but where a text
        // stands as written), with `+0`, in lower case, with a prefix, from
        // the value columns, over the next row's range and over both; the
        // plain formulas beside the bands among them. Between frames a value
        // over a band's first cell gives the band up, so that fills whose
        // master is not the range's first cell are met. Writes alike at every
        // cell, differing at every cell, and some of each.
        let (mut alike, mut differ, mut some) = (0, 0, 0);
        for seed in 1..=60u64 {
            let mut d = Dice(seed ^ 0x34E5);
            let rows = generated_grid(&mut d);
            let bands = formula_rows_of(&rows);
            let mut m = load(&rows);
            for frame in 1..=3 {
                assert!(m.step(0).done);
                let when = format!("seed {seed}, frame {frame}\n{rows}");
                for (k, (sheet, range, text)) in bands.iter().enumerate() {
                    let mut tries: Vec<(A1Range, String)> = vec![
                        (*range, text.clone()),
                        (*range, format!("{text}+0")),
                        (*range, text.to_lowercase()),
                        (*range, text.replace("SUM(", "_xlfn.SUM(")),
                        (A1Range { left: 1, ..*range }, text.clone()),
                    ];
                    if range.top < range.bottom && range.left < range.right {
                        let inside = A1Range {
                            top: range.top + 1,
                            left: range.left + 1,
                            ..*range
                        };
                        let moved = shift_a1_references(&text[1..], 1, 1);
                        tries.push((inside, format!("={moved}")));
                    }
                    if let Some((_, next, other)) = bands.get(k + 1).filter(|b| b.0 == *sheet) {
                        tries.push((*next, text.clone()));
                        let both = A1Range {
                            bottom: next.bottom.max(range.bottom),
                            right: next.right.max(range.right),
                            ..*range
                        };
                        tries.push((both, other.clone()));
                    }
                    for (r, t) in tries {
                        match same_text(&m, sheet, r, &t, &when) {
                            0 => alike += 1,
                            n if n as u64 == r.cells() => differ += 1,
                            _ => some += 1,
                        }
                    }
                }
                let (sheet, range, _) = &bands[d.below(bands.len())];
                let cell = crate::sheet::cell_ref(range.top, range.left);
                let w = format!("(cell \"{sheet}\" \"{cell}\" 1)");
                assert!(m.write(&w).is_ok(), "{when}: {w}");
            }
        }
        assert!(
            alike > 200 && differ > 200 && some > 200,
            "{alike} writes alike at every cell, {differ} differing at every cell and {some} \
             of each: too few for the law to hold over"
        );
    }

    #[test]
    fn the_rows_of_one_write_count_what_they_count_written_one_by_one() {
        // KERNEL.34: in one write a later row meets cells an earlier row wrote
        // or a release gave their own text, whose places the plan has not yet
        // seen; each row counts the cells comparing every cell counts, as the
        // same rows written one write each count them. A plain cell written
        // with another formula and then within a fill of it; a fill's cells
        // given their own text, one of them `#REF!`, then written with that.
        let rows = concat!(
            "(cell \"S\" \"A1:A8\" 1)\n",
            "(formula \"S\" \"B1:C4\" \"=A1+1\")\n",
            "(formula \"S\" \"D1\" \"=A1*2\")\n",
            "(formula \"S\" \"D2\" \"=A2*2\")\n",
            "(formula \"S\" \"E1:F2\" \"=XFD1\")\n",
        );
        for write in [
            "(formula \"S\" \"D1\" \"=A1*3\")\n(formula \"S\" \"D1:D2\" \"=A1*3\")",
            "(formula \"S\" \"G1\" \"=A1\")\n(formula \"S\" \"G1:G2\" \"=A1\")",
            "(cell \"S\" \"E1\" 0)\n(formula \"S\" \"F1:F2\" \"=#REF!\")",
            "(cell \"S\" \"E1\" 0)\n(formula \"S\" \"E2:F2\" \"=XFD2\")",
            "(cell \"S\" \"B1\" 7)\n(formula \"S\" \"B2:C4\" \"=A2+1\")",
            "(formula \"S\" \"B1:C2\" \"=A1-1\")\n(formula \"S\" \"B1:C4\" \"=A1-1\")",
        ] {
            let mut together = load(rows);
            let mut apart = load(rows);
            assert!(together.step(0).done && apart.step(0).done);
            let all = together.write(write).expect("written");
            assert_eq!(
                together.shapes.spelled(),
                plan_of_a_load(&together),
                "{write}"
            );
            let mut counted = 0;
            for row in write.lines() {
                let one = formula_rows_of(row);
                let want = one
                    .first()
                    .map(|(sheet, range, text)| same_text(&apart, sheet, *range, text, write));
                let got = apart.write(row).expect("written");
                if let Some(want) = want {
                    assert_eq!(got, want, "{write}: {row}");
                }
                counted += got;
            }
            assert_eq!(all, counted, "{write}");
            assert_eq!(save(&together), save(&apart), "{write}: the records");
        }
    }

    #[test]
    fn a_release_makes_the_fill_s_first_cell_left_its_master_and_moves_one_text() {
        // KERNEL.34: a cell written over a fill's first cell leaves the rest
        // of the fill one shared formula, its master the first cell left by
        // rows and then columns, holding the text moved there, and every other
        // cell its child, unchanged: one text moved, where every cell was
        // given its own, and every cell reads as it did.
        let mut m = load(concat!(
            "(cell \"S\" \"A1:A3\" 2)\n",
            "(formula \"S\" \"B1:C3\" \"=A1*10+S.last!B1\")\n",
        ));
        assert!(m.step(0).done);
        let others = |m: &Machine| -> Vec<String> {
            formula_rows(m)
                .into_iter()
                .filter(|r| !r.contains("\"B1\""))
                .collect()
        };
        let before = others(&m);
        crate::sheet::work::begin();
        assert_eq!(m.write("(cell \"S\" \"B1\" 5)"), Ok(1));
        let work = crate::sheet::work::end();
        assert_eq!((work.moved, work.whole, work.split), (1, 1, 0));
        assert_eq!(
            content(&m, "S", "C1"),
            Some(Content::SharedMaster {
                text: "B1*10+S.last!C1".to_string(),
                range: "B1:C3".to_string(),
                si: 0,
            })
        );
        for cell in ["B2", "C2", "B3", "C3"] {
            assert_eq!(
                content(&m, "S", cell),
                Some(Content::SharedChild { si: 0 }),
                "{cell}"
            );
        }
        assert_eq!(others(&m), before, "every other cell's formula as it was");
        assert_eq!(m.shapes.spelled(), plan_of_a_load(&m));
        // The new master given up in turn, by a range over it and a cell
        // below it: the first cell left is the next row's first.
        assert_eq!(
            as_a_load(
                &mut m,
                "(formula \"S\" \"B1:C1\" \"=7\")",
                "the master given up again"
            ),
            (Ok(2), true)
        );
        assert_eq!(
            content(&m, "S", "B2"),
            Some(Content::SharedMaster {
                text: "A2*10+S.last!B2".to_string(),
                range: "B1:C3".to_string(),
                si: 0,
            })
        );
        assert_eq!(others(&m)[1..], before[1..], "the rows below as they were");
    }

    #[test]
    fn a_fill_whose_text_a_release_would_move_off_the_sheet_gives_each_cell_its_own() {
        // KERNEL.34: a reference to the sheet's last column moved one column
        // right is `#REF!` and stays so, so the text moved to the first cell
        // left of a fill and moved again to the cell below and left of that
        // would not be the text moved there at once: the release gives each
        // cell its own text, as every release did, and the grid is a load's.
        let mut m = load("(formula \"S\" \"A1:B3\" \"=XFD1\")\n");
        assert!(m.step(0).done);
        let before: Vec<String> = formula_rows(&m)
            .into_iter()
            .filter(|r| !r.contains("\"A1\""))
            .collect();
        crate::sheet::work::begin();
        assert_eq!(
            as_a_load(&mut m, "(cell \"S\" \"A1\" 0)", "the edge"),
            (Ok(1), true)
        );
        let work = crate::sheet::work::end();
        assert_eq!((work.whole, work.split), (0, 1));
        assert_eq!(
            content(&m, "S", "B1"),
            Some(Content::Formula("#REF!".to_string()))
        );
        assert_eq!(
            content(&m, "S", "A2"),
            Some(Content::Formula("XFD2".to_string()))
        );
        let after: Vec<String> = formula_rows(&m)
            .into_iter()
            .filter(|r| !r.contains("\"A1\""))
            .collect();
        assert_eq!(after, before, "every other cell's formula as it was");
    }

    /// What one write did, counted (KERNEL.34).
    fn work_of(m: &mut Machine, write: &str) -> crate::sheet::work::Work {
        crate::sheet::work::begin();
        let answer = m.write(write);
        let work = crate::sheet::work::end();
        assert!(answer.is_ok(), "{write}: {answer:?}");
        work
    }

    #[test]
    fn a_fill_written_into_a_running_grid_costs_its_cells() {
        // KERNEL.34: the cells a write walks and the texts it moves, counted
        // and not timed, for a fill of 20 rows and of 40, eight cells a row:
        // a fill twice the size costs at most about twice the work, where
        // every cell placed walked the sheet, a fill's text written whole
        // moves two texts whatever its size, and a cell written over a fill's
        // first cell moves one. Each kind of write: the fill whole, a row
        // inside it, a column inside it, a value over its first cell, and a
        // column over its first cell.
        let grid = |rows: u32| {
            format!(
                "(cell \"S\" \"A1:A{rows}\" 1)\n(formula \"S\" \"B1:H{rows}\" \"=A1+S.last!B1\")\n"
            )
        };
        let writes = |rows: u32| {
            [
                format!("(formula \"S\" \"B1:H{rows}\" \"=A1+S.last!B1+0\")"),
                "(formula \"S\" \"B3:H3\" \"=A3*2\")".to_string(),
                format!("(formula \"S\" \"C1:C{rows}\" \"=A1*3\")"),
                "(cell \"S\" \"B1\" 5)".to_string(),
                format!("(formula \"S\" \"B1:B{rows}\" \"=A1*4\")"),
            ]
        };
        let cost = |rows: u32| -> Vec<crate::sheet::work::Work> {
            writes(rows)
                .iter()
                .map(|w| {
                    let mut m = load(&grid(rows));
                    assert!(m.step(0).done);
                    work_of(&mut m, w)
                })
                .collect()
        };
        let (small, large) = (cost(20), cost(40));
        for (k, (s, l)) in small.iter().zip(&large).enumerate() {
            assert!(
                l.walked <= 2 * s.walked + 64,
                "write {k}: {} cells walked for 20 rows, {} for 40",
                s.walked,
                l.walked
            );
            assert!(
                l.moved == s.moved && s.moved <= 3,
                "write {k}: {} texts moved for 20 rows, {} for 40",
                s.moved,
                l.moved
            );
        }
        assert_eq!(large[0].moved, 2, "a fill written whole: two texts");
        assert_eq!(large[3].moved, 1, "a cell over the fill's first: one text");
    }

    /// A formula write a generated grid takes between two frames, over one
    /// of its bands (KERNEL.34): another text over the band whole, over a
    /// row, a column or a cell of it, from its first cell or inside; its
    /// own text again; its own text moved to a cell inside; a value over its
    /// first cell; another text from the value columns, or over the band and
    /// the rows below; two rows, the second the band's own; or a cell that
    /// reads itself, which is refused.
    fn formula_between(d: &mut Dice, bands: &[(String, A1Range, String)]) -> String {
        let (sheet, range, text) = &bands[d.below(bands.len())];
        let texts = [
            "=A.last!C1*2+$A$1",
            "=ROW()+COLUMN()",
            "=SUM(B.last!C1:D2)",
            "=$B$1&\"x\"",
            "=IF(A.last!D1>1,1,0)",
        ];
        let other = d.pick(&texts).to_string();
        let row = range.top + d.below((range.bottom - range.top + 1) as usize) as u32;
        let col = range.left + d.below((range.right - range.left + 1) as usize) as u32;
        let (row, col) = if d.below(3) == 0 {
            (range.top, range.left)
        } else {
            (row, col)
        };
        let part = match d.below(4) {
            0 => *range,
            1 => A1Range {
                top: row,
                bottom: row,
                ..*range
            },
            2 => A1Range {
                left: col,
                right: col,
                ..*range
            },
            _ => A1Range {
                top: row,
                bottom: row,
                left: col,
                right: col,
            },
        };
        let over = |r: A1Range, t: &str| {
            format!(
                "(formula \"{sheet}\" \"{}\" {})",
                r.text(),
                crate::rows::quoted(t)
            )
        };
        match d.below(10) {
            0..=2 => over(part, &other),
            3 => over(*range, text),
            4 => {
                let at = (
                    i64::from(part.top) - i64::from(range.top),
                    i64::from(part.left) - i64::from(range.left),
                );
                let moved = shift_a1_references(&text[1..], at.0, at.1);
                over(part, &format!("={moved}"))
            }
            5 => format!(
                "(cell \"{sheet}\" \"{}\" 1)",
                crate::sheet::cell_ref(range.top, range.left)
            ),
            6 => over(A1Range { left: 1, ..part }, &other),
            7 => over(
                A1Range {
                    bottom: range.bottom + 2,
                    ..*range
                },
                &other,
            ),
            8 => format!("{}\n{}", over(part, &other), over(*range, text)),
            _ => {
                let cell = crate::sheet::cell_ref(row, col);
                over(
                    A1Range {
                        top: row,
                        bottom: row,
                        left: col,
                        right: col,
                    },
                    &format!("={cell}+1"),
                )
            }
        }
    }

    #[test]
    fn a_write_between_frames_is_a_load_of_the_grid_it_leaves() {
        // KERNEL.34: grids of generated formulas, each frame preceded by a
        // formula write over one of their bands (`formula_between`), each held
        // to a load of the grid it leaves (`as_a_load`), and a write of one
        // formula row counting the cells comparing every cell counts. Writes
        // that gave a fill up whole and cell by cell, wrote nothing and were
        // refused, and frames held to `Calc`'s.
        writes_as_loads(1..=40, (21, 3, 11, 6, 61));
    }

    #[test]
    fn a_write_between_frames_is_a_load_of_the_grid_it_leaves_every_sheet_indexed() {
        // KERNEL.35: the law above with every sheet's readers indexed, by the
        // axis an index finds, so that the plan a write leaves is held to a
        // load's index and all (`Shapes::spelled` spells it).
        let _every = crate::calc::reach::every::sheet_indexed(None);
        writes_as_loads(41..=60, (6, 0, 4, 3, 21));
    }

    /// `a_write_between_frames_is_a_load_of_the_grid_it_leaves` over the
    /// grids of some seeds, and the fewest of each kind of write and frame
    /// they must hold: fills given up whole and cell by cell, writes of
    /// nothing, refusals, frames held to `Calc`'s.
    fn writes_as_loads(
        seeds: std::ops::RangeInclusive<u64>,
        least: (u64, u64, usize, usize, usize),
    ) {
        let (mut nothing, mut refused, mut held) = (0, 0, 0);
        crate::sheet::work::begin();
        for seed in seeds {
            let mut d = Dice(seed ^ 0x3400);
            let rows = generated_grid(&mut d);
            let bands = formula_rows_of(&rows);
            let mut m = load(&rows);
            assert!(m.step(0).done);
            for frame in 1..=6 {
                let when = format!("seed {seed}, frame {frame}\n{rows}");
                let write = formula_between(&mut d, &bands);
                let one = formula_rows_of(&write);
                let want = match one.as_slice() {
                    [(sheet, range, text)] => Some(same_text(&m, sheet, *range, text, &when)),
                    _ => None,
                };
                let (answer, by_calc) = as_a_load(&mut m, &write, &when);
                held += usize::from(by_calc);
                match answer {
                    Ok(n) => {
                        if let Some(want) = want {
                            assert_eq!(n, want, "{when}: {write}: the cells counted");
                        }
                        nothing += usize::from(n == 0);
                    }
                    Err(_) => {
                        refused += 1;
                        assert!(m.step(0).done);
                    }
                }
            }
        }
        let work = crate::sheet::work::end();
        let (whole, split, nothings, refusals, helds) = least;
        assert!(
            work.whole >= whole
                && work.split >= split
                && nothing >= nothings
                && refused >= refusals
                && held >= helds,
            "{} fills given up whole, {} cell by cell, {nothing} writes of nothing, \
             {refused} refused and {held} frames held to Calc's: too few for the law to \
             hold over",
            work.whole,
            work.split
        );
    }

    // ---- KERNEL.35: a sheet's readers by where a changed cell stands ----

    /// How a grid's formulas over more than one cell are written: as one
    /// fill each, one fill a row, one fill a column, or their first rows one
    /// formula a cell and the rest one fill.
    #[derive(Clone, Copy, Debug)]
    enum Spelling {
        One,
        Rows,
        Cols,
        Cells(u32),
    }

    /// A grid's rows with each formula over more than one cell written as
    /// `how` says (KERNEL.35), each part's text the fill's moved to the
    /// part's first cell, so that each cell reads as the fill read there; a
    /// formula a cell is given a text of its own that computes the same
    /// value, `(...)+0*n`. Every other row as it was.
    fn spelled(rows: &str, how: Spelling) -> String {
        if let Spelling::One = how {
            return rows.to_string();
        }
        let mut out = String::new();
        let mut n = 0;
        for line in rows.lines() {
            let fill =
                read_forms(line)
                    .ok()
                    .and_then(|forms| match load::row_parts(forms.first()?)? {
                        ("formula", [Form::Str(sheet), Form::Str(addr), Form::Str(text)], _) => {
                            Some((sheet.clone(), parse_a1_range(addr)?, text.clone()))
                        }
                        _ => None,
                    });
            let Some((sheet, r, text)) = fill.filter(|f| f.1.cells() > 1) else {
                out.push_str(line);
                out.push('\n');
                continue;
            };
            let moved = |row: u32, col: u32| {
                let d = (i64::from(row - r.top), i64::from(col - r.left));
                shift_a1_references(&text[1..], d.0, d.1)
            };
            let mut put = |range: A1Range, text: String| {
                out.push_str(&format!(
                    "(formula {} {} {})\n",
                    crate::rows::quoted(&sheet),
                    crate::rows::quoted(&range.text()),
                    crate::rows::quoted(&text)
                ));
            };
            match how {
                Spelling::Rows => {
                    for row in r.top..=r.bottom {
                        let part = A1Range {
                            top: row,
                            bottom: row,
                            ..r
                        };
                        put(part, format!("={}", moved(row, r.left)));
                    }
                }
                Spelling::Cols => {
                    for col in r.left..=r.right {
                        let part = A1Range {
                            left: col,
                            right: col,
                            ..r
                        };
                        put(part, format!("={}", moved(r.top, col)));
                    }
                }
                Spelling::Cells(k) => {
                    let last = r.bottom.min(r.top + k.max(1) - 1);
                    for row in r.top..=last {
                        for col in r.left..=r.right {
                            n += 1;
                            let cell = A1Range {
                                top: row,
                                left: col,
                                bottom: row,
                                right: col,
                            };
                            put(cell, format!("=({})+0*{n}", moved(row, col)));
                        }
                    }
                    if last < r.bottom {
                        let rest = A1Range { top: last + 1, ..r };
                        put(rest, format!("={}", moved(last + 1, r.left)));
                    }
                }
                Spelling::One => {}
            }
        }
        out
    }

    /// A grid of the readers the generated grids hold few of (KERNEL.35): a
    /// value sheet `V` read by fills of rows and of columns and by single
    /// cells of `S`, through moving windows, held rectangles and a part
    /// computed at a frame's start; held rectangles too wide to file by
    /// both axes, by rows alone and by columns alone, which an index checks
    /// against every changed cell, or every one of a band's; a fill reading
    /// its own sheet; one shape standing on `S` and on `T`; fills at the
    /// sheet's last columns, whose references move off it; and `n` single
    /// cells of `T` reading the sheet `W`, which holds `n` readers.
    fn index_grid(n: u32) -> String {
        let mut out = String::from(concat!(
            "(cell \"V\" \"A1:G30\" 1)\n",
            "(cell \"V\" \"H1:L30\" 2)\n",
            "(cell \"V\" \"XFA1:XFD6\" 3)\n",
            "(cell \"W\" \"A1:A20\" 4)\n",
        ));
        for r in 1..=8 {
            out.push_str(&format!(
                "(formula \"S\" \"A{r}:F{r}\" \"=SUM(V!A{r}:B{}){}\")\n",
                r + 1,
                "+IF(V!$H$1=2,ROW(),0)+SUM(V!$J$2:$K$3)"
            ));
        }
        for col in ["H", "I", "J", "K", "L", "M"] {
            out.push_str(&format!(
                "(formula \"S\" \"{col}1:{col}8\" \"=V!{col}1*2+V!$A1\")\n"
            ));
        }
        out.push_str(concat!(
            "(formula \"S\" \"A10\" \"=V!C5+V!$D$6\")\n",
            "(formula \"S\" \"B10\" \"=V!C6*3\")\n",
            "(formula \"S\" \"A16\" \"=SUM(V!$A$1:$CZ$200)\")\n",
            "(formula \"S\" \"B16\" \"=SUM(V!$E$1:$E$150)\")\n",
            "(formula \"S\" \"C16\" \"=SUM(V!$A$5:$CV$5)\")\n",
            "(formula \"S\" \"A12:F14\" \"=A11+V!A1\")\n",
            "(formula \"S\" \"P1\" \"=A1+V!B2\")\n",
            "(formula \"T\" \"P1\" \"=A1+V!B2\")\n",
            "(formula \"S\" \"XFB1:XFD2\" \"=V!XFC1\")\n",
            "(formula \"T\" \"XFB5:XFD5\" \"=XFD1\")\n",
        ));
        for k in 1..=n {
            out.push_str(&format!("(formula \"T\" \"H{k}\" \"=W!A{k}+{k}\")\n"));
        }
        out
    }

    /// The last row any row of a grid names, for the window its cells are
    /// changed over: the grid's own rows, and one more.
    fn last_row(rows: &str) -> u32 {
        read_forms(rows)
            .unwrap_or_default()
            .iter()
            .filter_map(|form| match load::row_parts(form)? {
                (_, [Form::Str(_), Form::Str(addr), _], _) => Some(parse_a1_range(addr)?.bottom),
                _ => None,
            })
            .filter(|&row| row < 1_000)
            .max()
            .unwrap_or(1)
            + 1
    }

    /// Every cell of every sheet, the twins' among them, over the first
    /// `rows` rows and `cols` columns and at the sheet's far edges, changed,
    /// reaching through its sheet's index exactly what every reader checked
    /// reaches, and checking no more readers than the sheet holds
    /// (KERNEL.35); how many sheets were indexed.
    fn reaches_as_the_scan_does(m: &Machine, (rows, cols): (u32, u32), when: &str) -> usize {
        use crate::calc::reach::{checked, Reached};
        use crate::sheet::{MAX_COLUMN, MAX_ROW};
        let reach = m.shapes.reach();
        let key = |r: Reached| match r {
            Reached::Cell(p) => (0, p),
            Reached::Shape(s) => (1, s),
            Reached::Parts(s) => (2, s),
        };
        let edges = (1..=6)
            .flat_map(|row| (MAX_COLUMN - 4..=MAX_COLUMN).map(move |col| (row, col)))
            .chain([(MAX_ROW, 1), (MAX_ROW - 2, 3), (MAX_ROW, MAX_COLUMN)]);
        let mut indexed = 0;
        for sheet in 0..m.book.sheets.len() {
            indexed += usize::from(reach.most_on(sheet).is_some());
            let readers = reach.readers_on(sheet) as u64;
            let name = &m.book.sheets[sheet].name;
            let window = (1..=rows).flat_map(|row| (1..=cols).map(move |col| (row, col)));
            for (row, col) in window.chain(edges.clone()) {
                let id = (sheet, row, col);
                let (mut by_index, mut by_scan) = (Vec::new(), Vec::new());
                checked::begin();
                reach.each(id, &mut |r| by_index.push(key(r)));
                let checks = checked::end();
                reach.each_by_scan(id, &mut |r| by_scan.push(key(r)));
                let at = format!("{when}: {name}!{}", crate::sheet::cell_ref(row, col));
                assert!(
                    checks <= readers,
                    "{at}: {checks} of {readers} readers checked"
                );
                by_index.sort_unstable();
                by_index.dedup();
                by_scan.sort_unstable();
                by_scan.dedup();
                assert_eq!(by_index, by_scan, "{at}");
            }
        }
        indexed
    }

    #[test]
    fn the_index_reaches_what_the_scan_reaches() {
        // KERNEL.35: generated grids, and grids of the readers they hold few
        // of with a sheet of 15, 16 and 17 readers, each written one fill, one
        // fill a row, one fill a column and its first rows a cell each;
        // loaded as a load indexes them, and with every sheet indexed, first
        // by rows, by columns and by the axis an index finds. At the load,
        // after a frame and a formula's write, which plans the grid again, and
        // after a value over a fill's first cell, which gives it up: every
        // cell around the grid's and at the sheet's far edges, changed,
        // reaches through the index what every reader checked reaches.
        let mut grids: Vec<(String, String, Vec<String>)> = Vec::new();
        for seed in 1..=8u64 {
            let mut d = Dice(seed ^ 0x3535);
            let rows = generated_grid(&mut d);
            let bands = formula_rows_of(&rows);
            let writes = (0..2).map(|_| formula_between(&mut d, &bands)).collect();
            grids.push((format!("seed {seed}"), rows, writes));
        }
        for n in [15, 16, 17] {
            let writes = vec![
                "(formula \"S\" \"A3:F4\" \"=V!A1*7\")".to_string(),
                "(cell \"S\" \"H1\" 5)".to_string(),
            ];
            grids.push((format!("{n} readers of W"), index_grid(n), writes));
        }
        let (mut indexed, mut cases) = (0, 0);
        for (name, rows, writes) in &grids {
            let around = (last_row(rows) + 1, 18);
            for how in [
                Spelling::One,
                Spelling::Rows,
                Spelling::Cols,
                Spelling::Cells(2),
            ] {
                let text = spelled(rows, how);
                for way in [None, Some(None), Some(Some(true)), Some(Some(false))] {
                    let _every = way.map(crate::calc::reach::every::sheet_indexed);
                    let when = format!("{name}, {how:?}, every sheet indexed {way:?}");
                    let mut m = load(&text);
                    indexed += reaches_as_the_scan_does(&m, around, &when);
                    for write in writes {
                        assert!(m.step(0).done);
                        if m.write(write).is_ok() {
                            let when = format!("{when}, after {write}");
                            indexed += reaches_as_the_scan_does(&m, around, &when);
                        }
                    }
                    cases += 1;
                }
            }
        }
        assert!(
            cases == grids.len() * 16 && indexed > cases * 2,
            "{indexed} sheets indexed over {cases} cases: too few for the law to hold over"
        );
    }

    #[test]
    fn a_sheet_is_indexed_from_sixteen_readers_and_its_census_says_the_most_it_checks() {
        // KERNEL.35: a sheet of 15 readers is checked reader by reader and its
        // census row is as it was; one of 16 or 17, each reading one cell of
        // its own, is indexed, and a changed cell of it meets one reader, as
        // its row says, `(most 1)`; the value sheet that fills of rows, of
        // columns and single cells read is indexed too.
        for (n, row) in [
            (15, "(readers \"W\" 15)\n"),
            (16, "(readers \"W\" 16 (most 1))\n"),
            (17, "(readers \"W\" 17 (most 1))\n"),
        ] {
            let m = load(&index_grid(n));
            let census = m.census();
            assert!(census.contains(row), "{n}: {census}");
            let w = m.find_sheet("W").expect("W");
            assert_eq!(m.shapes.reach().most_on(w).is_some(), n >= 16, "{n}");
            let v = m.find_sheet("V").expect("V");
            let most = m.shapes.reach().most_on(v).expect("V indexed");
            assert!(
                most < m.shapes.reach().readers_on(v),
                "{n}: V's most, {most}"
            );
        }
    }

    /// A board of `n` rows and columns of value cells from `V!A1`, read by a
    /// formula at each of `n` by `n` cells from `F!B2`, each summing the
    /// three by three value cells around its own, written as `how` says
    /// (KERNEL.35).
    fn board(n: u32, how: Spelling) -> String {
        let (v, f) = (
            crate::sheet::cell_ref(n + 2, n + 2),
            crate::sheet::cell_ref(n + 1, n + 1),
        );
        spelled(
            &format!("(cell \"V\" \"A1:{v}\" 1)\n(formula \"F\" \"B2:{f}\" \"=SUM(V!A1:C3)\")\n"),
            how,
        )
    }

    #[test]
    fn a_board_written_any_way_checks_the_readers_around_a_changed_cell() {
        // KERNEL.35, counted and not timed: the readers a frame checks after
        // three value cells are written. Written one fill a row or a column, a
        // changed cell meets the three rows or columns of formulas around it,
        // three times one fill's one reader, and written a cell each the nine
        // formulas around it, at 16 rows and at 32 alike; every reader checked,
        // as before the index, a changed cell meets every row, column or
        // cell of the board.
        use crate::calc::reach::checked;
        let writes = [(5, 5), (9, 8), (12, 10)];
        for n in [16u32, 32] {
            for (how, per, all) in [
                (Spelling::One, 1, 1),
                (Spelling::Rows, 3, n),
                (Spelling::Cols, 3, n),
                (Spelling::Cells(n), 9, n * n),
            ] {
                let mut m = load(&board(n, how));
                assert!(m.step(0).done && m.step(0).done);
                let v = m.find_sheet("V").expect("V");
                for &(row, col) in &writes {
                    let cell = crate::sheet::cell_ref(row, col);
                    assert_eq!(m.write(&format!("(cell \"V\" \"{cell}\" 2)")), Ok(1));
                }
                checked::begin();
                assert!(m.step(0).done);
                let frame = checked::end();
                checked::begin();
                for &(row, col) in &writes {
                    m.shapes.reach().each_by_scan((v, row, col), &mut |_| {});
                }
                let scan = checked::end();
                assert_eq!(
                    (frame, scan),
                    (3 * per, 3 * u64::from(all)),
                    "{how:?} at {n} rows: the readers checked, through the index and by the scan"
                );
            }
        }
    }

    // ---- KERNEL.36: a frame's end at any share ----

    /// Two grids' frames' ends the same to the bit, one copying where it may
    /// and one walking at every frame's end: `same_state`; every twin's cells
    /// as the twin's B-tree holds them, where no slot stands ahead of them;
    /// whether the next frame computes every cell; and, where it does not,
    /// the twin cells kept as changed, as a set.
    fn same_end(copied: &Machine, walked: &Machine, when: &str) {
        same_state_but_arrays(copied, walked, when);
        // Each twin's array the walk's, or, where the copying machine's still
        // stands over the extent its sheet had at the last walk, a formula
        // with no value added past it since, every slot a read of the twin
        // finds the same: the twin gained no cell, and an array answers
        // nothing for a cell outside it, as the twin's cells do.
        for s in 0..copied.sheets {
            let (a, b) = (copied.dense[s].as_ref(), walked.dense[s].as_ref());
            let extent = walked.extents[s];
            match (a, b) {
                (Some(a), Some(b)) if !a.stands_over(extent) && b.stands_over(extent) => {
                    let ((top, left), (bottom, right)) = extent.expect("an array's extent");
                    for row in top..=bottom {
                        for col in left..=right {
                            let at = |k: &Dense| {
                                format!(
                                    "{:?}",
                                    k.get(row, col).filter(|v| !matches!(v, Value::Empty))
                                )
                            };
                            assert_eq!(
                                at(a),
                                at(b),
                                "{when}: the array of sheet {s} at ({row}, {col})"
                            );
                        }
                    }
                }
                _ => assert_eq!(
                    format!("{a:?}"),
                    format!("{b:?}"),
                    "{when}: the array of sheet {s}"
                ),
            }
        }
        assert_eq!(
            copied.every_next, walked.every_next,
            "{when}: the next frame computing every cell"
        );
        if !copied.every_next {
            let set = |m: &Machine| m.twin_changed.iter().copied().collect::<BTreeSet<CellId>>();
            assert_eq!(set(copied), set(walked), "{when}: the twin cells changed");
        }
        for s in 0..copied.sheets {
            let t = copied.sheets + s;
            let (held, walked_held) = (&copied.book.sheets[t].cells, &walked.book.sheets[t].cells);
            let name = &copied.book.sheets[t].name;
            assert!(held.keys().eq(walked_held.keys()), "{when}: {name}'s cells");
            let behind = |row: u32, col: u32| {
                copied.dense[s]
                    .as_ref()
                    .and_then(|k| k.index(row, col))
                    .is_some_and(|i| {
                        copied.behind[s]
                            .get(i / 64)
                            .is_some_and(|w| w >> (i % 64) & 1 == 1)
                    })
            };
            for (&(row, col), cell) in held.iter().filter(|(&(row, col), _)| !behind(row, col)) {
                assert_eq!(
                    cell.content,
                    walked_held[&(row, col)].content,
                    "{when}: {name}!{} as its B-tree holds it",
                    crate::sheet::cell_ref(row, col)
                );
            }
        }
    }

    /// The cells a frame's step and end counted (`sheet::work`), the frame
    /// done.
    fn ended(m: &mut Machine, budget: usize) -> crate::sheet::work::Work {
        crate::sheet::work::begin();
        while !m.step(budget).done {}
        crate::sheet::work::end()
    }

    #[test]
    fn a_frame_s_end_copied_is_the_walk_bit_for_bit() {
        // KERNEL.36: grids of generated formulas most of whose cells change
        // at every frame, the Clock written before every frame, copied at a
        // frame's end however many changed, against the same grids walking at
        // every frame's end, frame for frame and bit for bit (`same_end`); the
        // values against the same grids computing every cell (KERNEL.20's
        // reference). Writes between frames change value cells, add cells and
        // change formulas; a sheet is viewed as a plane, the first, which most
        // writes change, or the second, so that a write to it drops its plane
        // and the next frame's end walks over twin cells copied behind their
        // slots; steps yield. Frames whose ends copied more than half the
        // formula cells, and walks after a frame's end that left twin cells
        // behind, each seen.
        let (mut past_half, mut after_copies) = (0, 0);
        for seed in 1..=60u64 {
            let mut d = Dice(seed ^ 0x36C0);
            let rows = churning_grid(&mut d);
            let mut copied = load(&rows);
            let mut walked = load(&rows);
            walked.copies = false;
            let mut whole = load(&rows);
            whole.incremental = false;
            let viewed = if seed % 2 == 1 { "A" } else { "B" };
            for m in [&copied, &walked, &whole] {
                m.view("plane", viewed, None).expect("a plane");
            }
            for frame in 1..=10u32 {
                let when = format!("seed {seed}, frame {frame}\n{rows}");
                let mut writes = vec![format!("(cell \"Clock\" \"B1\" {frame})")];
                writes.extend(write_between(&mut d, frame));
                for w in &writes {
                    let (a, b, c) = (copied.write(w), walked.write(w), whole.write(w));
                    assert!(
                        a.is_ok() == b.is_ok() && b.is_ok() == c.is_ok(),
                        "{when}: {w}: {a:?} {b:?} {c:?}"
                    );
                }
                let lagging = copied.lagging.iter().any(|&l| l);
                let of = copied.shapes.order().len();
                let work = ended(&mut copied, 1 + d.below(300));
                assert!(walked.step(0).done && whole.step(0).done);
                same_end(&copied, &walked, &when);
                assert_eq!(values(&copied), values(&whole), "{when}: computed whole");
                past_half += usize::from(work.swept == 0 && work.copied as usize * 2 > of);
                after_copies += usize::from(lagging && work.swept > 0);
            }
        }
        assert!(
            past_half >= 80 && after_copies >= 60,
            "{past_half} frames' ends copied past half and {after_copies} walked over twin \
             cells behind their slots: too few for the law to hold over"
        );
    }

    /// A sheet `Grid` of `n` formulas in column B, the first `k` changing at
    /// every frame and the rest computed again to the same value, and column
    /// A the last frame's B doubled, through B's twin; the Clock's `B1`
    /// beside; a value cell far off when `sparse`, so that the sheet is too
    /// sparse for an array (KERNEL.36). From the third frame on, `2k` formula
    /// cells change at every frame. A stands first in the order, so a frame's
    /// end notes B's twin changes, the ones A reads, last: past the room kept
    /// for them, the next frame must compute every cell, or A misses them.
    fn share_grid(n: u32, k: u32, sparse: bool) -> String {
        let far = if sparse {
            "(cell \"Grid\" \"Z1000\" 1)\n"
        } else {
            ""
        };
        format!(
            "(sheet \"Clock\" visible)\n(sheet \"Grid\" visible)\n(cell \"Clock\" \"B1\" 0)\n\
             (formula \"Grid\" \"A1:A{n}\" \"=Grid.last!B1*2\")\n\
             (formula \"Grid\" \"B1:B{n}\" \"=IF(ROW()<={k},MOD(Clock!$B$1+ROW(),16),0)\")\n{far}"
        )
    }

    #[test]
    fn a_frame_s_end_copies_into_arrays_at_any_share_and_through_cells_to_half() {
        // KERNEL.36: 400 formulas of which 198, 200, 202 or all change at
        // every frame, the Clock written before each. On a sheet kept as an
        // array every frame's end copies, past half and at all, the cells it
        // visits the changed and the Clock; on a sheet too sparse for an
        // array, a frame's end copies while the cells it copies through the
        // twin's own cells, the changed and the Clock, are no more than half
        // the cells (exactly half at 200), and walks past that. Each held to
        // the walk's (`same_end`), its values to every cell computed, and past
        // a twin change beyond the room kept for them, which makes the next
        // frame compute every cell.
        let mut overflowed = 0;
        for sparse in [false, true] {
            for k in [99u32, 100, 101, 200] {
                let rows = share_grid(200, k, sparse);
                let mut copied = load(&rows);
                let mut walked = load(&rows);
                walked.copies = false;
                let mut whole = load(&rows);
                whole.incremental = false;
                for frame in 1..=6u64 {
                    let when = format!("{k} of 200 changing, sparse {sparse}, frame {frame}");
                    let w = format!("(cell \"Clock\" \"B1\" {frame})");
                    for m in [&mut copied, &mut walked, &mut whole] {
                        assert_eq!(m.write(&w), Ok(1), "{when}");
                    }
                    let work = ended(&mut copied, 0);
                    assert!(walked.step(0).done && whole.step(0).done);
                    same_end(&copied, &walked, &when);
                    assert_eq!(values(&copied), values(&whole), "{when}: computed whole");
                    let grid = copied.find_sheet("Grid").expect("Grid");
                    assert_eq!(copied.dense[grid].is_none(), sparse, "{when}: Grid's array");
                    if frame < 3 {
                        continue;
                    }
                    let changed = 2 * u64::from(k);
                    let through = if sparse { changed + 1 } else { 1 };
                    if through * 2 <= copied.cells as u64 {
                        assert_eq!(
                            (work.copied, work.swept),
                            (changed + 1, 0),
                            "{when}: copied"
                        );
                    } else {
                        assert_eq!(work.swept, copied.cells as u64, "{when}: walked");
                    }
                    overflowed += usize::from(copied.every_next);
                }
            }
        }
        assert!(
            overflowed >= 6,
            "{overflowed} frames' twin changes past their room"
        );
    }

    #[test]
    fn a_frame_s_end_counts_only_the_cells_it_copies_through_twins_own_cells() {
        // KERNEL.36: a sheet kept as an array, all 400 of its formulas
        // changing, beside a sheet too sparse for one, 500 or 501 of its 600
        // changing, of 1,002 cells: the cells a copy writes through twins'
        // own cells, the sparse sheet's changed and the Clock, are exactly
        // half at 500, and the frame's end copies all 901; at 501 it walks.
        // Neither bound decides, so the changed list is counted. Each frame
        // held to the walk's (`same_end`) and its values to every cell
        // computed.
        for j in [500u32, 501] {
            let rows = format!(
                "{}(sheet \"Far\" visible)\n\
                 (formula \"Far\" \"A1:A600\" \"=IF(ROW()<={j},MOD(Clock!$B$1+ROW(),16),0)\")\n\
                 (cell \"Far\" \"Z5000\" 1)\n",
                share_grid(200, 200, false)
            );
            let mut copied = load(&rows);
            let mut walked = load(&rows);
            walked.copies = false;
            let mut whole = load(&rows);
            whole.incremental = false;
            for frame in 1..=5u64 {
                let when = format!("{j} of Far's 600 changing, frame {frame}");
                let w = format!("(cell \"Clock\" \"B1\" {frame})");
                for m in [&mut copied, &mut walked, &mut whole] {
                    assert_eq!(m.write(&w), Ok(1), "{when}");
                }
                let work = ended(&mut copied, 0);
                assert!(walked.step(0).done && whole.step(0).done);
                same_end(&copied, &walked, &when);
                assert_eq!(values(&copied), values(&whole), "{when}: computed whole");
                let (grid, far) = (
                    copied.find_sheet("Grid").expect("Grid"),
                    copied.find_sheet("Far").expect("Far"),
                );
                assert!(
                    copied.dense[grid].is_some() && copied.dense[far].is_none(),
                    "{when}"
                );
                assert_eq!(copied.cells, 1_002, "{when}");
                if frame < 3 {
                    continue;
                }
                if j == 500 {
                    assert_eq!((work.copied, work.swept), (901, 0), "{when}: copied");
                } else {
                    assert_eq!(work.swept, 1_002, "{when}: walked");
                }
            }
        }
    }

    /// The number its argument holds while it is even, a reason while it is
    /// odd: a cell whose twin cell gains a value at one frame and loses it at
    /// the next (KERNEL.36).
    fn flip(ctx: &crate::calc::Ctx<'_>, args: &[crate::calc::Arg<'_>]) -> Result<Value, Reason> {
        match library::number_arg(ctx, args, 0)? {
            Ok(n) if n % 2.0 == 0.0 => Ok(Value::Number(n)),
            Ok(_) => Err(Reason::Function("FLIP".to_string())),
            Err(e) => Ok(Value::Error(e)),
        }
    }

    #[test]
    fn a_twin_cell_gaining_or_losing_a_value_ends_its_frame_by_the_walk() {
        // KERNEL.36: a cell whose value is a number at one frame and a reason
        // at the next, beside a column most of whose cells change and a cell
        // reading its twin: its twin cell gains a value and loses it at
        // alternate frames, which a copy refuses, so every frame's end walks,
        // held to the walk's (`same_end`) and its values to every cell
        // computed.
        let rows = concat!(
            "(sheet \"Clock\" visible)\n(sheet \"Grid\" visible)\n(cell \"Clock\" \"B1\" 0)\n",
            "(formula \"Grid\" \"A1:A100\" \"=MOD(Clock!$B$1+ROW(),16)\")\n",
            "(formula \"Grid\" \"B1\" \"=FLIP(Clock!$B$1)\")\n",
            "(formula \"Grid\" \"C1\" \"=Grid.last!B1\")\n",
        );
        let flipping = || {
            let mut lib = library::language();
            lib.register("FLIP", flip);
            Machine::load(rows, lib).unwrap_or_else(|e| panic!("{}", e.refusal.text))
        };
        let (mut copied, mut walked, mut whole) = (flipping(), flipping(), flipping());
        walked.copies = false;
        whole.incremental = false;
        for frame in 1..=8u64 {
            let when = format!("frame {frame}");
            let w = format!("(cell \"Clock\" \"B1\" {frame})");
            for m in [&mut copied, &mut walked, &mut whole] {
                assert_eq!(m.write(&w), Ok(1), "{when}");
            }
            let work = ended(&mut copied, 0);
            assert!(walked.step(0).done && whole.step(0).done);
            same_end(&copied, &walked, &when);
            assert_eq!(values(&copied), values(&whole), "{when}: computed whole");
            let twin = copied.find_sheet("Grid.last").expect("Grid's twin");
            assert_eq!(
                copied.twin_reasons.contains_key(&(twin, 1, 2)),
                frame % 2 == 1,
                "{when}: B1's twin a reason at odd frames"
            );
            assert!(work.swept > 0, "{when}: walked");
        }
    }

    #[test]
    fn a_frame_s_end_visits_the_cells_it_changed_and_not_the_grid() {
        // KERNEL.36, counted and not timed: the cells a frame's end visits.
        // Copying, the cells that changed and the cell written, at 1,000
        // formulas a column and at 4,000 alike when as many change, and nine
        // tenths of them when nine tenths do; walking at every frame's end,
        // every cell of the grid. The changed list holds every formula cell
        // from the load and never grows.
        for n in [1_000u32, 4_000] {
            for k in [100u32, n * 9 / 10] {
                let rows = share_grid(n, k, false);
                let mut copied = load(&rows);
                let mut walked = load(&rows);
                walked.copies = false;
                let room = copied.changed.capacity();
                assert!(room >= copied.shapes.order().len(), "{n}: the list's room");
                for frame in 1..=5u64 {
                    let when = format!("{k} of {n} changing, frame {frame}");
                    let w = format!("(cell \"Clock\" \"B1\" {frame})");
                    assert!(copied.write(&w).is_ok() && walked.write(&w).is_ok());
                    let (c, w) = (ended(&mut copied, 0), ended(&mut walked, 0));
                    if frame >= 3 {
                        let changed = 2 * u64::from(k);
                        assert_eq!((c.copied, c.swept), (changed + 1, 0), "{when}: copied");
                        assert_eq!(
                            (w.copied, w.swept),
                            (0, walked.cells as u64),
                            "{when}: walked"
                        );
                    }
                }
                assert_eq!(copied.changed.capacity(), room, "{n}: the list grew");
            }
        }
    }

    #[test]
    fn a_walk_after_copied_frames_makes_each_cell_behind_its_slot_where_it_passes() {
        // KERNEL.36: frames whose ends copy, leaving the twin cells they
        // change behind their slots, then a frame's end that walks: after a
        // write into the sheet viewed as a plane, which drops its plane; and
        // after a write that adds a cell past the sheet's extent, whose array
        // the walk makes again over the new extent, its cells behind made
        // first. After each walk the frames' ends copy again, the Clock
        // written into its own sheet before every frame. Each frame held to
        // the walk's (`same_end`), the twin's cells as its B-tree holds them,
        // behind their slots no more after a walk.
        let rows = concat!(
            "(sheet \"Clock\" visible)\n(sheet \"Grid\" visible)\n(cell \"Clock\" \"B1\" 0)\n",
            "(formula \"Grid\" \"A1:A300\" \"=MOD(Clock!$B$1+ROW(),16)\")\n",
            "(formula \"Grid\" \"B1:B300\" \"=Grid.last!A1*2+Grid.last!C1\")\n",
            "(cell \"Grid\" \"C1:C300\" 0)\n",
        );
        let mut copied = load(rows);
        let mut walked = load(rows);
        walked.copies = false;
        for m in [&copied, &walked] {
            m.view("plane", "Grid", None).expect("a plane of Grid");
        }
        let grid = copied.find_sheet("Grid").expect("Grid");
        let mut walks_after = 0;
        for frame in 1..=12u64 {
            let when = format!("frame {frame}");
            let mut w = format!("(cell \"Clock\" \"B1\" {frame})");
            match frame {
                5 => w.push_str("\n(cell \"Grid\" \"C7\" 5)"),
                8 => w.push_str("\n(cell \"Grid\" \"D400\" 1)"),
                _ => {}
            }
            assert!(
                copied.write(&w).is_ok() && walked.write(&w).is_ok(),
                "{when}"
            );
            let lagging = copied.lagging[grid];
            let work = ended(&mut copied, 0);
            assert!(walked.step(0).done);
            same_end(&copied, &walked, &when);
            match frame {
                5 | 8 => {
                    assert!(
                        lagging && work.swept > 0,
                        "{when}: a walk over cells behind"
                    );
                    assert!(!copied.lagging[grid], "{when}: behind no more");
                    assert!(copied.behind[grid].iter().all(|&w| w == 0), "{when}");
                    walks_after += 1;
                }
                2..=4 | 6 | 7 | 9..=12 => {
                    assert_eq!(work.swept, 0, "{when}: copied");
                    assert!(copied.lagging[grid], "{when}: Grid's twin cells behind");
                }
                _ => {}
            }
        }
        assert_eq!(walks_after, 2);
    }

    #[test]
    fn a_walk_over_an_array_made_again_makes_its_cells_behind_first() {
        // KERNEL.36: a formula written past a sheet's extent whose value is a
        // reason, so that the extent grows and the twin gains no cell: the
        // frames' ends go on copying into the array kept over the old extent,
        // leaving cells behind their slots, until a frame's end walks, which
        // makes the array again over the new extent, its slots read by other
        // indices. The cells behind are made first, from the old array; read
        // through the new one's indices, a cell whose value was the one its
        // twin cell held when last made, two frames back, would pass as
        // unchanged, as values that alternate at every frame make some.
        // Held to the walk's at every frame (`same_end`) and its values to
        // every cell computed.
        let rows = concat!(
            "(sheet \"Clock\" visible)\n(sheet \"Grid\" visible)\n(cell \"Clock\" \"B1\" 0)\n",
            "(formula \"Grid\" \"A1:A100\" \"=MOD(Clock!$B$1+ROW(),2)\")\n",
            "(formula \"Grid\" \"B1:B100\" \"=Grid.last!A1*2\")\n",
            "(cell \"Grid\" \"C1:C100\" 0)\n",
        );
        let mut copied = load(rows);
        let mut walked = load(rows);
        walked.copies = false;
        let mut whole = load(rows);
        whole.incremental = false;
        let grid = copied.find_sheet("Grid").expect("Grid");
        for frame in 1..=9u64 {
            let when = format!("frame {frame}");
            let mut w = format!("(cell \"Clock\" \"B1\" {frame})");
            if frame == 4 {
                w.push_str("\n(formula \"Grid\" \"E200\" \"=A1:A2\")");
            }
            for m in [&mut copied, &mut walked, &mut whole] {
                assert!(m.write(&w).is_ok(), "{when}");
            }
            let lagging = copied.lagging[grid];
            let work = ended(&mut copied, 0);
            assert!(walked.step(0).done && whole.step(0).done);
            same_end(&copied, &walked, &when);
            assert_eq!(values(&copied), values(&whole), "{when}: computed whole");
            if (2..=6).contains(&frame) {
                assert_eq!(work.swept, 0, "{when}: copied");
            }
            if frame == 7 {
                assert!(
                    lagging && work.swept > 0,
                    "{when}: a walk over cells behind"
                );
            }
            // A plane asked for after the sixth frame: its first kept at the
            // seventh frame's end, which walks.
            if frame == 6 {
                for m in [&copied, &walked, &whole] {
                    m.view("plane", "Grid", None).expect("a plane of Grid");
                }
                let extent = copied.extents[grid];
                assert!(
                    copied.dense[grid]
                        .as_ref()
                        .is_some_and(|k| !k.stands_over(extent)),
                    "{when}: the array kept over the old extent"
                );
            }
        }
    }

    // ---- KERNEL.42: a window's changes ----

    /// A window's changes, as the machine answers them.
    fn changes_of(m: &Machine, sheet: &str, window: Option<&str>) -> String {
        match m.view("changes", sheet, window) {
            Ok(Viewed::Text(t)) => t,
            other => panic!("{other:?}"),
        }
    }

    /// The cell a record's row belongs to: a `cell`, `value`, `style`,
    /// `sentence` or `empty` row's.
    fn cell_of(line: &str) -> Option<(u32, u32)> {
        let forms = read_forms(line).ok()?;
        let Some(Form::List(l)) = forms.first() else {
            return None;
        };
        match l.items.as_slice() {
            [Form::Sym(head), Form::Str(_), Form::Str(addr), ..]
                if matches!(
                    head.as_str(),
                    "cell" | "value" | "style" | "sentence" | "empty"
                ) =>
            {
                let r = parse_a1_range(addr)?;
                Some((r.top, r.left))
            }
            _ => None,
        }
    }

    /// A record's rows before its cells', and each cell's rows.
    fn rows_by_cell(record: &str) -> (String, BTreeMap<(u32, u32), String>) {
        let mut head = String::new();
        let mut cells: BTreeMap<(u32, u32), String> = BTreeMap::new();
        for line in record.lines() {
            match cell_of(line) {
                Some(cell) => cells
                    .entry(cell)
                    .or_default()
                    .push_str(&format!("{line}\n")),
                None => head.push_str(&format!("{line}\n")),
            }
        }
        (head, cells)
    }

    /// The rows a changes answer gives each cell it names, none for an
    /// `empty` row's; `None` for changes answered whole.
    fn given(changes: &str) -> Option<BTreeMap<(u32, u32), String>> {
        let mut lines = changes.lines();
        if lines.next()?.ends_with(" whole)") {
            return None;
        }
        let mut given: BTreeMap<(u32, u32), String> = BTreeMap::new();
        for line in lines {
            let cell =
                cell_of(line).unwrap_or_else(|| panic!("a changes row names a cell: {line}"));
            let rows = given.entry(cell).or_default();
            if !line.starts_with("(empty ") {
                rows.push_str(&format!("{line}\n"));
            }
        }
        Some(given)
    }

    /// A values record patched by a window's changes, as a host patches
    /// the record it holds: each cell the changes name has its rows
    /// replaced with theirs, an `empty` row taking them away, every cell in
    /// the record's order. `None` for changes answered whole.
    fn patched(record: &str, changes: &str) -> Option<String> {
        let (head, mut cells) = rows_by_cell(record);
        for (cell, rows) in given(changes)? {
            match rows.is_empty() {
                true => cells.remove(&cell),
                false => cells.insert(cell, rows),
            };
        }
        Some(head + &cells.into_values().collect::<String>())
    }

    /// The number of cells a changes answer names, `None` for whole.
    fn changed_cells(changes: &str) -> Option<usize> {
        let head = changes.lines().next().expect("a changes row");
        let last = head.trim_end_matches(')').rsplit(' ').next()?;
        last.parse().ok()
    }

    /// The window's values at the last frame patched by its changes at
    /// this one, against its values now: equal, the changes naming no cell
    /// exactly when the two records are, or answered whole. Whether whole.
    fn held_to(m: &Machine, sheet: &str, window: Option<&str>, before: &str, when: &str) -> bool {
        let now = values_view(m, sheet, window);
        let changes = changes_of(m, sheet, window);
        let Some(got) = patched(before, &changes) else {
            return true;
        };
        assert_eq!(got, now, "{when}: {sheet} {window:?} patched by\n{changes}");
        assert_eq!(
            changed_cells(&changes) == Some(0),
            before == now,
            "{when}: {sheet} {window:?}: no cell named exactly when still\n{changes}"
        );
        // And a cell named exactly when its rows changed.
        let (_, was) = rows_by_cell(before);
        let (_, is) = rows_by_cell(&now);
        for cell in given(&changes).expect("not whole").keys() {
            assert_ne!(
                was.get(cell),
                is.get(cell),
                "{when}: {sheet} {window:?}: {cell:?} named, its rows as they were\n{changes}"
            );
        }
        false
    }

    /// A still card in small: a day of 30 frames read from the Clock, and a
    /// Screen of 400 formulas of which the first hundred read the day.
    const STILL: &str = concat!(
        "(sheet \"Clock\" visible)\n(cell \"Clock\" \"B1\" 0)\n",
        "(sheet \"Day\" visible)\n(formula \"Day\" \"B1\" \"=INT(Clock!$B$1/30)\")\n",
        "(sheet \"Screen\" visible)\n",
        "(formula \"Screen\" \"A1:J10\" \"=Day!$B$1*10000+ROW()*100+COLUMN()\")\n",
        "(formula \"Screen\" \"K1:T10\" \"=ROW()*100+COLUMN()\")\n",
        "(formula \"Screen\" \"A11:T20\" \"=ROW()*100+COLUMN()\")\n",
        "(cell \"Screen\" \"U1\" 7)\n",
    );

    /// What `f` did, counted (`sheet::work`).
    fn counted(f: impl FnOnce()) -> crate::sheet::work::Work {
        crate::sheet::work::begin();
        f();
        crate::sheet::work::end()
    }

    /// One frame of a card: its Clock written, then stepped whole.
    fn next_frame(m: &mut Machine) {
        let next = m.frame() + 1;
        assert!(m.write(&format!("(cell \"Clock\" \"B1\" {next})")).is_ok());
        assert!(m.step(0).done);
    }

    #[test]
    fn a_still_window_s_changes_visit_none_of_its_cells() {
        // KERNEL.42: the still card, its Screen still but at the day's turn.
        // At a still frame the changes name no cell and visit none, where the
        // values visit every cell of the window; at the turn they name the
        // hundred cells that read the day, each spelled as the values spell
        // it, visit those alone, and patch the last frame's values into this
        // frame's.
        let mut m = load(STILL);
        next_frame(&mut m);
        assert_eq!(
            changes_of(&m, "Screen", None),
            "(changes \"Screen\" \"A1:U20\" 0 1 whole)\n",
            "the first frame: every formula gained its first value"
        );
        let mut before = values_view(&m, "Screen", None);
        for frame in 2..=61u64 {
            next_frame(&mut m);
            let when = format!("frame {frame}");
            let work = counted(|| {
                changes_of(&m, "Screen", None);
            });
            let changes = changes_of(&m, "Screen", None);
            assert!(
                !held_to(&m, "Screen", None, &before, &when),
                "{when}: whole"
            );
            if frame % 30 == 0 {
                let j10 = format!("(value \"Screen\" \"J10\" {})\n", frame / 30 * 10000 + 1010);
                assert_eq!(changed_cells(&changes), Some(100), "{when}\n{changes}");
                assert!(changes.contains(&j10), "{when}\n{changes}");
                assert_eq!(work.viewed, 100, "{when}: the changed cells alone");
            } else {
                assert_eq!(
                    changes,
                    format!("(changes \"Screen\" \"A1:U20\" {} {frame} 0)\n", frame - 1),
                    "{when}"
                );
                assert_eq!(work.viewed, 0, "{when}: no cell visited");
            }
            assert_eq!(
                (work.walked, work.extents),
                (0, 0),
                "{when}: no sheet walked"
            );
            let viewed = counted(|| {
                values_view(&m, "Screen", None);
            });
            assert_eq!(viewed.viewed, 401, "{when}: the values visit the window");
            before = values_view(&m, "Screen", None);
        }
    }

    /// A window's changes now, and the window's values now as the record a
    /// host holds.
    fn asked(m: &Machine, window: &str) -> (String, String) {
        (
            changes_of(m, "Screen", Some(window)),
            values_view(m, "Screen", Some(window)),
        )
    }

    #[test]
    fn the_changes_answer_whole_where_the_lists_cannot_say() {
        // KERNEL.42: before the first frame; at the first, every formula
        // gaining a value; after a write that changes a formula or adds a
        // cell, until the next frame's end and at it, the frame after it
        // answering again; while a value written since lies in the window,
        // and not when it lies outside or on another sheet; and at a frame
        // that changes more cells than the machine keeps room for.
        let mut m = load(STILL);
        let whole =
            |m: &Machine, w: &str| changed_cells(&changes_of(m, "Screen", Some(w))).is_none();
        assert_eq!(
            changes_of(&m, "Screen", Some("A1:T20")),
            "(changes \"Screen\" \"A1:T20\" 0 0 whole)\n"
        );
        // A grid resumed at its saved frame, before its first step.
        let mut resumed = load(STILL);
        assert!(resumed.resume_at(41).is_ok());
        assert_eq!(
            changes_of(&resumed, "Screen", Some("A1:T20")),
            "(changes \"Screen\" \"A1:T20\" 40 41 whole)\n"
        );
        next_frame(&mut m);
        assert!(whole(&m, "A1:T20"), "frame 1");
        next_frame(&mut m);
        assert!(!whole(&m, "A1:T20"), "frame 2");
        // A formula written over: its kind and its value may change where no
        // twin cell does.
        assert!(m.write("(formula \"Screen\" \"T20\" \"=1+1\")").is_ok());
        assert!(
            whole(&m, "A1:T20") && whole(&m, "A1:B2"),
            "after a formula's write"
        );
        next_frame(&mut m);
        assert!(whole(&m, "A1:T20"), "the frame after a formula's write");
        let (_, before) = asked(&m, "A1:U20");
        next_frame(&mut m);
        assert!(!whole(&m, "A1:T20"), "the frame after that");
        assert!(!held_to(&m, "Screen", Some("A1:U20"), &before, "frame 4"));
        // A cell added: the extent may change.
        assert!(m.write("(cell \"Screen\" \"V30\" 1)").is_ok());
        assert!(whole(&m, "A1:B2"), "after a cell added");
        next_frame(&mut m);
        assert!(whole(&m, "A1:B2"), "the frame after a cell added");
        next_frame(&mut m);
        assert!(!whole(&m, "A1:B2"), "the frame after that");
        // A value written into the window, and outside it.
        let (changes, before) = asked(&m, "A1:U20");
        assert!(m.write("(cell \"Screen\" \"U1\" 8)").is_ok());
        assert!(whole(&m, "A1:U20"), "a value written into the window");
        assert!(!whole(&m, "A1:T20"), "a value written outside the window");
        assert!(
            m.write("(cell \"Clock\" \"B1\" 1000)").is_ok() && !whole(&m, "A1:T20"),
            "a value written on another sheet"
        );
        assert!(!changes.contains("U1"));
        assert!(m.step(0).done);
        let at = changes_of(&m, "Screen", Some("A1:U20"));
        assert!(at.contains("(cell \"Screen\" \"U1\" 8)\n"), "{at}");
        assert!(!held_to(
            &m,
            "Screen",
            Some("A1:U20"),
            &before,
            "the written frame"
        ));
        // More changes than the room kept, half the cells and 64: a card
        // whose every Screen cell reads the day, loaded so that no write has
        // grown the room, its day turning at frame 30.
        let mut m = load(&STILL.replace("ROW()*100+COLUMN()\")", "ROW()*100+COLUMN()+Day!$B$1\")"));
        for frame in 1..=31u64 {
            next_frame(&mut m);
            if frame > 1 {
                let when = format!("frame {frame}");
                assert_eq!(
                    m.every_next,
                    frame == 30,
                    "{when}: more changed than the room kept"
                );
                assert_eq!(whole(&m, "A1:T20"), frame == 30, "{when}");
            }
        }
    }

    #[test]
    fn a_view_during_a_yield_answers_the_last_frame_s_changes() {
        // KERNEL.42: a step that yields keeps the last frame's changes for a
        // view before its frame completes, the values showing that frame
        // too; a value written into the window before the step makes them
        // whole until the frame's end.
        let mut m = load(STILL);
        for _ in 0..30 {
            next_frame(&mut m);
        }
        let turned = changes_of(&m, "Screen", None);
        assert_eq!(changed_cells(&turned), Some(100), "{turned}");
        assert!(m.write("(cell \"Clock\" \"B1\" 31)").is_ok());
        assert!(!m.step(50).done);
        assert_eq!(changes_of(&m, "Screen", None), turned, "during the yield");
        while !m.step(50).done {}
        assert_eq!(changed_cells(&changes_of(&m, "Screen", None)), Some(0));
        assert!(m
            .write("(cell \"Clock\" \"B1\" 32)\n(cell \"Screen\" \"U1\" 9)")
            .is_ok());
        assert!(!m.step(50).done);
        assert!(
            changed_cells(&changes_of(&m, "Screen", None)).is_none(),
            "written into"
        );
        assert_eq!(
            changed_cells(&changes_of(&m, "Screen", Some("A1:T20"))),
            Some(0)
        );
        while !m.step(50).done {}
        let after = changes_of(&m, "Screen", None);
        assert_eq!(
            after,
            "(changes \"Screen\" \"A1:U20\" 31 32 1)\n(cell \"Screen\" \"U1\" 9)\n"
        );
    }

    /// A grid whose cells turn at every frame between values a record
    /// spells alike: `B1` 1 and 0 by turns, `C1:C{n}` 0 and -0 as `B1`
    /// reads, and `D1` a reason and another. A -0 comes only from a cell
    /// holding one, arithmetic making it 0. `E1`'s value stretches the
    /// twin's extent past `D1`, which holds no twin cell, so that a view of
    /// the twin takes it in.
    fn alike_rows(n: u32) -> String {
        format!(
            concat!(
                "(sheet \"S\" visible)\n(cell \"S\" \"A1\" 0)\n(cell \"S\" \"A2\" -0)\n",
                "(cell \"S\" \"E1\" 1)\n",
                "(formula \"S\" \"B1\" \"=IF(S.last!B1=1,0,1)\")\n",
                "(formula \"S\" \"D1\" \"=IF(B1=1,X1:X2,Y1:Y2)\")\n",
                "(formula \"S\" \"C1:C{}\" \"=IF($B$1=1,$A$1,$A$2)\")\n",
            ),
            n
        )
    }

    #[test]
    fn a_change_its_record_spells_as_before_is_no_change_of_it() {
        // KERNEL.42: a formula's value turned from 0 to -0 is a change of the
        // twin, whose cell row spells -0, and none of the sheet, whose value
        // row spells 0 both times; a reason turned to another, a change of
        // the sheet's value row and none of the twin, which holds no cell
        // for it. Up to 64 such changes a frame are kept; past them the
        // frame's changes are whole.
        let mut m = load(&alike_rows(3));
        assert!(m.step(0).done);
        let reason = |m: &Machine| match at(m, "S", "D1") {
            Some(Computed::NotComputed(Reason::Construct(c))) => c,
            other => panic!("D1: {other:?}"),
        };
        let mut before = (values_view(&m, "S", None), values_view(&m, "S.last", None));
        for frame in 2..=6u64 {
            let was = reason(&m);
            assert!(m.step(0).done);
            let when = format!("frame {frame}");
            assert_ne!(reason(&m), was, "{when}: D1's reason turned");
            assert_eq!(
                content(&m, "S.last", "C1")
                    .map(|c| matches!(c, Content::Number(n) if n.is_sign_negative())),
                Some(frame % 2 == 0),
                "{when}: C1's twin 0 and -0 by turns"
            );
            let sheet = changes_of(&m, "S", None);
            let twin = changes_of(&m, "S.last", None);
            let names = |text: &str, addr: &str| text.contains(&format!(" \"{addr}\" "));
            assert!(
                names(&sheet, "B1") && names(&sheet, "D1") && !names(&sheet, "C1"),
                "{when}\n{sheet}"
            );
            assert!(
                names(&twin, "B1") && names(&twin, "C1") && !names(&twin, "D1"),
                "{when}\n{twin}"
            );
            assert_eq!(changed_cells(&sheet), Some(2), "{when}\n{sheet}");
            assert_eq!(changed_cells(&twin), Some(4), "{when}\n{twin}");
            assert!(!held_to(&m, "S", None, &before.0, &when));
            assert!(!held_to(&m, "S.last", None, &before.1, &when));
            before = (values_view(&m, "S", None), values_view(&m, "S.last", None));
        }
        // 64 kept a frame: 63 of C and D1's reason; 65, and whole.
        for (n, whole) in [(63, false), (64, true)] {
            let mut m = load(&alike_rows(n));
            for _ in 0..3 {
                assert!(m.step(0).done);
            }
            assert_eq!(
                changed_cells(&changes_of(&m, "S", None)).is_none(),
                whole,
                "{n} cells"
            );
        }
    }

    #[test]
    fn the_changes_are_alike_copied_walked_and_computed_whole() {
        // KERNEL.42: the changes of every window, sheet and twin, whole and
        // through a window, alike from a machine whose frame's end copies
        // (its formulas' twin cells left behind their slots), one that
        // copies them at once, one that walks at every frame's end, and one
        // that computes every cell, over a grid of values written between
        // frames, a twin read, a day, a value turning between 0 and -0 and a
        // reason turning, with a plane asked for midway, so that a walk
        // passes cells left behind their slots.
        let rows = format!(
            "{}{}{}",
            STILL,
            alike_rows(4),
            concat!(
                "(sheet \"G\" visible)\n(cell \"G\" \"A1:A12\" 3)\n",
                "(formula \"G\" \"B1:D12\" \"=MOD(Clock!$B$1*3+ROW()*7+COLUMN(),5)+G.last!$A1\")\n",
            ),
        );
        let mut machines = [load(&rows), load(&rows), load(&rows), load(&rows)];
        machines[1].lazy = false;
        machines[2].copies = false;
        machines[3].incremental = false;
        let windows: Vec<(String, Option<&str>)> = ["Screen", "S", "G", "Clock", "Day"]
            .iter()
            .flat_map(|s| [s.to_string(), format!("{s}.last")])
            .flat_map(|s| [(s.clone(), None), (s, Some("B2:C11"))])
            .collect();
        let mut before: Vec<String> = Vec::new();
        for frame in 1..=40u64 {
            let mut write = format!("(cell \"Clock\" \"B1\" {frame})");
            if frame % 7 == 3 {
                write.push_str(&format!(
                    "\n(cell \"G\" \"A{}\" {})",
                    1 + frame % 12,
                    frame % 4
                ));
            }
            for m in &mut machines {
                assert!(m.write(&write).is_ok());
                assert!(m.step(0).done);
                if frame == 20 {
                    m.view("plane", "S", None).expect("a plane of S");
                }
            }
            let when = format!("frame {frame}");
            let now: Vec<String> = windows
                .iter()
                .map(|(s, w)| changes_of(&machines[0], s, *w))
                .collect();
            for m in &machines[1..] {
                let theirs: Vec<String> =
                    windows.iter().map(|(s, w)| changes_of(m, s, *w)).collect();
                assert_eq!(theirs, now, "{when}");
            }
            if frame > 1 {
                for ((s, w), was) in windows.iter().zip(&before) {
                    held_to(&machines[0], s, *w, was, &when);
                }
            }
            before = windows
                .iter()
                .map(|(s, w)| values_view(&machines[0], s, *w))
                .collect();
        }
        assert!(
            machines[0].lagging.iter().any(|&l| l),
            "cells left behind their slots"
        );
    }

    // ---- KERNEL.44: a cell alone lowered once to a register program ----

    /// Twice a number its operand reads as, as `ABS` reads one: a strict
    /// function no run knows, which a program calls through its pointer.
    fn twice(ctx: &crate::calc::Ctx<'_>, args: &[crate::calc::Arg<'_>]) -> Result<Value, Reason> {
        if args.len() != 1 {
            return Ok(Value::Error(ErrorKind::Value));
        }
        Ok(match library::number_arg(ctx, args, 0)? {
            Ok(n) => crate::calc::eval::finite(n * 2.0),
            Err(k) => Value::Error(k),
        })
    }

    /// 1 for an operand that is nothing, handed as a value or read from the
    /// one cell it names, else 0: a function that tells nothing from 0, as
    /// `ISBLANK` does.
    fn empties(ctx: &crate::calc::Ctx<'_>, args: &[crate::calc::Arg<'_>]) -> Result<Value, Reason> {
        Ok(Value::Number(match args {
            [crate::calc::Arg::Scalar(Value::Empty)] => 1.0,
            [a @ crate::calc::Arg::Area { .. }] => match ctx.one(a)? {
                Value::Empty => 1.0,
                _ => 0.0,
            },
            _ => 0.0,
        }))
    }

    /// 1 for an operand handed as cells, 0 for one handed as a value.
    fn area(_: &crate::calc::Ctx<'_>, args: &[crate::calc::Arg<'_>]) -> Result<Value, Reason> {
        Ok(Value::Number(match args {
            [crate::calc::Arg::Area { .. }] => 1.0,
            _ => 0.0,
        }))
    }

    /// Its first argument as `Ctx` evaluates it: a form no program holds.
    fn first<'e>(
        ctx: &crate::calc::Ctx<'_>,
        args: &'e [crate::calc::Expr],
    ) -> Result<crate::calc::Arg<'e>, Reason> {
        match args.first() {
            Some(e) => ctx.eval(e),
            None => Ok(crate::calc::Arg::Scalar(Value::Error(ErrorKind::Value))),
        }
    }

    /// The language's library with the functions above.
    fn program_library() -> Library {
        let mut lib = library::language();
        lib.register("TWICE", twice);
        lib.register("EMPTIES", empties);
        lib.register("AREA", area);
        lib.register_form("FIRST", first);
        lib
    }

    /// A leaf of `lowered_expr`: a number, a truth value, a reference
    /// (`reference`'s: a value cell of any kind, a twin, a formula cell above,
    /// the sheet before, a cell moved off the sheet), and one time in eighty
    /// a node a program leaves.
    fn lowered_leaf(d: &mut Dice, at: &Band) -> String {
        match d.below(16) {
            0 | 1 => d
                .pick(&["0", "1", "2", "-3", "0.5", "1E+300", "-0"])
                .to_string(),
            2 => d.pick(&["TRUE", "FALSE"]).to_string(),
            3 if d.below(5) == 0 => d
                .pick(&["\"x\"", "#N/A", "Rate", "A1:A2", "FIRST(1)"])
                .to_string(),
            _ => reference(d, at),
        }
    }

    /// A formula over every node a program holds, now and then one it leaves
    /// (a text, `&`, a comparison of two kinds, an `IF` of one argument, a
    /// call of an omitted `CHOOSE` value): the operators, a sign, a percent,
    /// `IF` with branches omitted or missing, `CHOOSE` with an index of every
    /// kind and omitted values, the aggregates over values and ranges, `NOT`,
    /// `ABS`, `INT`, `SIN`, `MOD`, `ROW` and `COLUMN` of the cell and of a
    /// reference, and strict functions through their pointers handed values,
    /// ranges and references.
    fn lowered_expr(d: &mut Dice, at: &Band, depth: u32) -> String {
        if depth >= 4 || d.below(4) == 0 {
            return lowered_leaf(d, at);
        }
        let e = |d: &mut Dice| lowered_expr(d, at, depth + 1);
        let operands = |d: &mut Dice, n: usize| {
            let parts: Vec<String> = (0..n)
                .map(|_| match d.below(3) {
                    0 => range(d, at),
                    1 => reference(d, at),
                    _ => lowered_expr(d, at, depth + 1),
                })
                .collect();
            parts.join(",")
        };
        match d.below(22) {
            0 => format!("-({})", e(d)),
            1 => format!("({})%", e(d)),
            2 => format!("+{}", reference(d, at)),
            3 | 4 => {
                let op = d.pick(&["+", "-", "*", "/", "^"]);
                format!("({}){op}({})", e(d), e(d))
            }
            5 => {
                let op = d.pick(&["=", "<>", "<", ">", "<=", ">=", "&"]);
                format!("({}){op}({})", e(d), e(d))
            }
            6 => format!("({})+({})-({})*({})", e(d), e(d), e(d), e(d)),
            7 => format!("IF({},{},{})", e(d), e(d), e(d)),
            8 => match d.below(5) {
                0 => format!("IF({},{})", e(d), e(d)),
                1 => format!("IF({},,{})", e(d), e(d)),
                2 => format!("IF({},{},)", e(d), e(d)),
                3 => format!("IF({})", e(d)),
                _ => format!("IF(({})>({}),{},{})", e(d), e(d), e(d), e(d)),
            },
            9 => {
                let f = d.pick(&["AND", "OR"]);
                let n = 1 + d.below(3);
                format!("{f}({})", operands(d, n))
            }
            10 => format!("NOT({})", e(d)),
            11 => {
                let f = d.pick(&["SUM", "MIN", "MAX"]);
                let n = d.below(4);
                format!("{f}({})", operands(d, n))
            }
            12 => format!("{}({})", d.pick(&["ABS", "INT", "SIN"]), e(d)),
            13 => format!("MOD({},{})", e(d), e(d)),
            14 => match d.below(4) {
                0 => d.pick(&["ROW()", "COLUMN()"]).to_string(),
                1 => format!("ROW({})", reference(d, at)),
                2 => format!("COLUMN({})", range(d, at)),
                _ => format!("ROW({})+COLUMN()", d.pick(&["A1048576", "XFD1", "$A$1"])),
            },
            15 => {
                let n = 2 + d.below(4);
                let index = match d.below(3) {
                    0 => d
                        .pick(&["0.5", "1", "-1", "TRUE", "1E+300", "3"])
                        .to_string(),
                    1 => format!("MOD(ROW(),{})+1", n + 1),
                    _ => e(d),
                };
                let values: Vec<String> = (0..n)
                    .map(|_| if d.below(8) == 0 { String::new() } else { e(d) })
                    .collect();
                format!("CHOOSE({index},{})", values.join(","))
            }
            16 => format!("TWICE({})", e(d)),
            17 => format!("TWICE({})", operands(d, 1)),
            18 => match d.below(3) {
                0 => format!("EMPTIES(CHOOSE({},{},))", e(d), e(d)),
                1 => format!("EMPTIES({})", reference(d, at)),
                _ => format!("EMPTIES({})", e(d)),
            },
            19 => format!("AREA({})", operands(d, 1)),
            _ => format!("{}*{}", reference(d, at), reference(d, at)),
        }
    }

    /// A grid of two sheets, each with value cells in A1:B20 of every kind,
    /// formulas over C1:H`rows` in bands of one to three rows, each a shape
    /// whose runs are fewer cells than `SHORT`, and a column of single
    /// formulas beside, each its own shape; the names `grid_of` defines.
    fn program_grid(d: &mut Dice) -> String {
        let rows = 4 + d.below(24) as u32;
        let mut out = String::new();
        for (sheet, before) in [("A", None), ("B", Some("A"))] {
            out.push_str(&format!("(sheet \"{sheet}\" visible)\n"));
            for row in 1..=20 {
                for col in ["A", "B"] {
                    if let Some(v) = datum(d) {
                        out.push_str(&format!("(cell \"{sheet}\" \"{col}{row}\" {v})\n"));
                    }
                }
            }
            let mut top = 1;
            while top <= rows {
                let bottom = (top + d.below(3) as u32).min(rows);
                let at = Band {
                    before,
                    top,
                    rows,
                    left: "B",
                };
                let f = crate::rows::quoted(&format!("={}", lowered_expr(d, &at, 0)));
                out.push_str(&format!("(formula \"{sheet}\" \"C{top}:H{bottom}\" {f})\n"));
                top = bottom + 1;
            }
            for row in 1..=rows {
                let at = Band {
                    before,
                    top: row,
                    rows,
                    left: "J",
                };
                let f = crate::rows::quoted(&format!("={}", lowered_expr(d, &at, 0)));
                out.push_str(&format!("(formula \"{sheet}\" \"K{row}\" {f})\n"));
            }
        }
        out.push_str("(name \"Rate\" \"A!$A$2\")\n");
        out.push_str("(name \"Box\" \"A!$A$1:$B$3\")\n");
        out.push_str("(name \"Prev\" \"A.last!$C$1\")\n");
        out.push_str("(name \"Band\" \"A!$C$1:$D$2\")\n");
        out
    }

    /// A machine over rows under `program_library`.
    fn load_programs(rows: &str, when: &str) -> Machine {
        Machine::load(rows, program_library())
            .unwrap_or_else(|e| panic!("{when}: line {}: {}\n{rows}", e.line, e.refusal.text))
    }

    /// Every formula cell's value as `Calc` computes it, spelled as `values`
    /// spells a machine's.
    fn calc_values(m: &Machine) -> BTreeMap<CellId, String> {
        let calc = crate::calc::Calc::run(&m.book, &m.library);
        m.shapes
            .order()
            .iter()
            .map(|id| (*id, format!("{:?}", calc.computed(*id).expect("computed"))))
            .collect()
    }

    #[test]
    fn a_cell_alone_by_its_program_is_its_evaluation_by_ctx() {
        // KERNEL.44: grids of short bands and single formulas over every node
        // a program holds and now and then one it leaves, reading value cells
        // of every kind, twins holding reasons, formula cells above and cells
        // moved off the sheet; stepped as a machine steps them (its cells
        // alone and its runs' few by their programs, in steps of any budget),
        // with every cell alone by its program, and with every cell alone by
        // `Ctx`, frame for frame and bit for bit, and against `Calc` at the
        // first frame. Writes between frames turn a value cell's number into a
        // text, a truth value or an error and back, and write a formula, which
        // plans again and lowers the shape it makes. Every guard fails
        // somewhere, and programs answer more than a third of the cells
        // alone, the rest left to `Ctx` by a guard (the grids read texts and
        // errors on purpose) or by a node a program does not hold.
        let mut guards = [0usize; program::guarded::GUARDS];
        let (mut ran, mut alone) = (0, 0);
        for seed in 1..=60u64 {
            let mut d = Dice(seed ^ 0x4444);
            let rows = program_grid(&mut d);
            let when = format!("seed {seed}");
            let mut m = load_programs(&rows, &when);
            let mut programs = load_programs(&rows, &when);
            programs.by_runs = false;
            let mut ctx = load_programs(&rows, &when);
            ctx.by_runs = false;
            ctx.lowered = false;
            let calc = calc_values(&m);
            program::guarded::take();
            for frame in 1..=6 {
                while !m.step(1 + d.below(40)).done {}
                assert!(programs.step(0).done);
                assert!(ctx.step(0).done);
                alone += programs.computed;
                let want = values(&ctx);
                assert_eq!(
                    values(&programs),
                    want,
                    "{when}, frame {frame}: every cell by its program\n{rows}"
                );
                assert_eq!(values(&m), want, "{when}, frame {frame}: by runs\n{rows}");
                if frame == 1 {
                    assert_eq!(want, calc, "{when}: Calc\n{rows}");
                }
                let write = match d.below(5) {
                    0 => format!(
                        "(cell \"A\" \"A{}\" {})",
                        1 + d.below(20),
                        d.pick(&["7", "-0.5", "0"])
                    ),
                    1 => format!(
                        "(cell \"B\" \"{}{}\" {})",
                        d.pick(&["A", "B"]),
                        1 + d.below(20),
                        d.pick(&["\"y\"", "\"3\"", "true", "(error \"#DIV/0!\")", "4"])
                    ),
                    2 => format!(
                        "(cell \"A\" \"{}{}\" 2)",
                        d.pick(&LETTERS[2..8]),
                        1 + d.below(4)
                    ),
                    3 => {
                        let row = 1 + d.below(4) as u32;
                        let at = Band {
                            before: None,
                            top: row,
                            rows: 4,
                            left: "J",
                        };
                        let f = crate::rows::quoted(&format!("={}", lowered_expr(&mut d, &at, 0)));
                        format!("(formula \"A\" \"K{row}\" {f})")
                    }
                    _ => continue,
                };
                let got: Vec<bool> = [&mut m, &mut programs, &mut ctx]
                    .into_iter()
                    .map(|x| x.write(&write).is_ok())
                    .collect();
                assert!(got.iter().all(|&g| g == got[0]), "{when}: {write}: {got:?}");
            }
            ran += programs.ran;
            let counts = program::guarded::take();
            for (g, c) in guards.iter_mut().zip(counts) {
                *g += c;
            }
        }
        assert!(
            guards.iter().all(|&g| g > 0),
            "every guard failed somewhere: reads, moves, answers, ranges, indices, calls {guards:?}"
        );
        assert!(
            ran * 3 > alone,
            "{ran} of {alone} cells alone answered by their programs: too few for the law to hold over"
        );
    }

    #[test]
    fn a_shape_left_to_ctx_is_named_by_the_first_node_a_program_does_not_hold() {
        // KERNEL.44: a single formula of each kind of node a program leaves,
        // placed in a grid, keeps the reason its lowering met first, one of
        // each holds a program, and every cell computes what `Ctx` computes.
        // Deeper than a program nests and within the 64 levels a parse does.
        let deep = format!("={}1", "-".repeat(52));
        let wide = format!("=SUM({}1)", "S!A1+0,".repeat(40));
        let cases: Vec<(String, Option<&str>)> = vec![
            ("=1+S!A1*2".into(), None),
            ("=S!A1>2".into(), None),
            ("=IF(S!A1>1,S!A2,-1)".into(), None),
            ("=CHOOSE(S!A1,1,,3)".into(), None),
            ("=TWICE(S!B1)+AREA(S!A1)".into(), None),
            ("=\"x\"".into(), Some("text")),
            ("=S!A1&S!A2".into(), Some("text")),
            ("=#N/A".into(), Some("error")),
            ("=MOD(S!A1)".into(), Some("error")),
            ("=IF(S!A1)".into(), Some("error")),
            ("=ROW(1)".into(), Some("error")),
            ("=Rate+1".into(), Some("name")),
            ("=S!A1:A2+1".into(), Some("range")),
            ("=SUM(IF(S!A1>0,S!A1:A2,S!B1))".into(), Some("range")),
            ("=FIRST(1)".into(), Some("form")),
            ("=S!A1=TRUE".into(), Some("kinds")),
            ("=IF(S!A1,1,TRUE)".into(), Some("kinds")),
            ("=Gone!A1+1".into(), Some("sheet")),
            ("=EMPTIES(CHOOSE(1,,2))".into(), Some("omitted")),
            ("=SUM(1,,2)".into(), Some("omitted")),
            (deep, Some("deep")),
            (wide, Some("deep")),
        ];
        let mut rows = String::from(
            "(sheet \"S\" visible)\n(cell \"S\" \"A1\" 2)\n(cell \"S\" \"A2\" 3)\n\
             (cell \"S\" \"B1\" \"x\")\n(sheet \"T\" visible)\n(name \"Rate\" \"S!$A$1\")\n",
        );
        for (k, (text, _)) in cases.iter().enumerate() {
            rows.push_str(&format!(
                "(formula \"T\" \"A{}\" {})\n",
                k + 1,
                crate::rows::quoted(text)
            ));
        }
        let mut m = load_programs(&rows, "the cases");
        let mut ctx = load_programs(&rows, "the cases");
        ctx.lowered = false;
        for (k, (text, want)) in cases.iter().enumerate() {
            let id = (1, k as u32 + 1, 1);
            let s = m.shapes.place(id).expect("placed").shape;
            let shape = m.shapes.shape(s).expect("a shape");
            let got = shape.unlowered().map(|why| why.word());
            assert_eq!(got, *want, "{text}");
            assert_eq!(m.shapes.program(s).is_some(), want.is_none(), "{text}");
        }
        for frame in 1..=2 {
            assert!(m.step(0).done && ctx.step(0).done);
            assert_eq!(values(&m), values(&ctx), "frame {frame}");
        }
        assert!(m.ran > 0, "the programs answered");
    }

    #[test]
    fn each_guard_leaves_its_cell_to_ctx() {
        // KERNEL.44: a cell whose program meets a value no register holds,
        // one guard a case, is computed by `Ctx`, its guard counted and no
        // other, its value `Ctx`'s: a read of a text, a reference moved off
        // the sheet (the fill's second cell), an operation answering an
        // error, a fold meeting a text in a range, `CHOOSE`'s index past its
        // values, a call answering an error.
        use program::guarded;
        // Each case's cells, its formula, its guard, and the cells its
        // program answers (the fill's first, whose reference stays on).
        let cases = [
            ("A1", "=S!B1+1", guarded::READ, 0),
            ("B1:B2", "=S!A1048576*0", guarded::OFF, 1),
            ("C1", "=1/S!A3", guarded::ANSWER, 0),
            ("D1", "=SUM(S!A1:B1)", guarded::RANGE, 0),
            ("E1", "=CHOOSE(S!A1+3,1,2)", guarded::INDEX, 0),
            ("F1", "=TWICE(S!B1)", guarded::CALL, 0),
        ];
        for (cell, text, guard, answered) in cases {
            let rows = format!(
                "(sheet \"S\" visible)\n(cell \"S\" \"A1\" 2)\n(cell \"S\" \"A2\" 3)\n\
                 (cell \"S\" \"B1\" \"x\")\n(sheet \"T\" visible)\n(formula \"T\" \"{cell}\" {})\n",
                crate::rows::quoted(text)
            );
            let mut m = load_programs(&rows, text);
            let mut ctx = load_programs(&rows, text);
            ctx.lowered = false;
            guarded::take();
            assert!(m.step(0).done && ctx.step(0).done);
            let counts = guarded::take();
            assert!(counts[guard] > 0, "{text}: {counts:?}");
            assert_eq!(
                counts.iter().sum::<usize>(),
                counts[guard],
                "{text}: no other guard {counts:?}"
            );
            assert_eq!(values(&m), values(&ctx), "{text}");
            assert_eq!(m.ran, answered, "{text}: the cells its program answered");
        }
    }

    #[test]
    fn a_run_s_few_cells_are_evaluated_by_its_program() {
        // KERNEL.44: runs of twenty, ten and sixteen cells of shapes with a
        // program, and one of twenty of a shape left to `Ctx`. A frame that
        // computes every cell computes the run of ten by its program at each
        // cell, fewer than `SHORT`, and the others as runs; one that reaches
        // one cell of each, the three by their programs and the fourth alone
        // by `Ctx`; eight of each, the three by their programs and the fourth,
        // with no program, as a run; sixteen of each, the run of ten by its
        // program and the others as runs. The values those of `Ctx` and of
        // the runs alone.
        let rows = concat!(
            "(sheet \"S\" visible)\n",
            "(cell \"S\" \"A1:A40\" 1)\n",
            "(sheet \"T\" visible)\n",
            "(formula \"T\" \"B1:B20\" \"=S.last!A1*2\")\n",
            "(formula \"T\" \"C1:C10\" \"=S.last!A1*3\")\n",
            "(formula \"T\" \"D1:D16\" \"=S.last!A1*5\")\n",
            "(formula \"T\" \"E1:E20\" \"=S.last!A1&\\\"\\\"\")\n",
        );
        let mut m = load(rows);
        let mut runs = load(rows);
        runs.lowered = false;
        let mut ctx = load(rows);
        ctx.by_runs = false;
        ctx.lowered = false;
        assert_eq!(m.shapes.cells_in_runs(), 66, "runs of 20, 10, 16 and 20");
        let shape = |m: &Machine, col: u32| m.shapes.place((1, 1, col)).expect("placed").shape;
        let (b, c, d, e) = (shape(&m, 2), shape(&m, 3), shape(&m, 4), shape(&m, 5));
        assert!(m.shapes.program(e).is_none() && m.shapes.program(b).is_some());
        let writes = [
            None,
            None,
            Some("(cell \"S\" \"A3\" 5)"),
            None,
            Some("(cell \"S\" \"A1:A8\" 7)"),
            None,
            Some("(cell \"S\" \"A1:A16\" 9)"),
            None,
        ];
        let mut ran = Vec::new();
        let mut as_runs = Vec::new();
        for (frame, write) in writes.iter().enumerate() {
            for x in [&mut m, &mut runs, &mut ctx] {
                if let Some(w) = write {
                    assert!(x.write(w).is_ok(), "{w}");
                }
            }
            let before = m.ran;
            vector::taken::begin();
            assert!(m.step(0).done);
            let shapes: BTreeSet<usize> =
                vector::taken::end().into_iter().map(|(s, _)| s).collect();
            assert!(runs.step(0).done && ctx.step(0).done);
            let want = values(&ctx);
            assert_eq!(values(&m), want, "frame {}", frame + 1);
            assert_eq!(values(&runs), want, "frame {}, runs", frame + 1);
            ran.push(m.ran - before);
            as_runs.push(shapes);
        }
        // The first two frames compute every cell, the twins gaining their
        // cells at the first's end; a written value reaches its twin at its
        // frame's end, and the cells that read the twin at the next frame.
        assert_eq!(
            ran,
            [10, 10, 0, 3, 0, 24, 0, 10],
            "the cells programs answered"
        );
        let set = |s: &[usize]| s.iter().copied().collect::<BTreeSet<usize>>();
        assert_eq!(
            as_runs[0],
            set(&[b, d, e]),
            "every cell: the shapes as runs"
        );
        assert_eq!(as_runs[3], set(&[]), "one cell of each: no run");
        assert_eq!(
            as_runs[5],
            set(&[e]),
            "eight of each: the shape with no program"
        );
        assert_eq!(
            as_runs[7],
            set(&[b, d, e]),
            "sixteen of each: all but the ten"
        );
        assert!(
            as_runs.iter().all(|s| !s.contains(&c)),
            "the run of ten is never a run"
        );
        assert_eq!(runs.ran, 0, "no program without them");
    }

    #[test]
    fn a_step_walks_the_runs_beside_the_positions() {
        // KERNEL.44: an order whose runs stand between cells alone at every
        // level, stepped in steps of every budget from one cell to all of
        // them, frame after frame with writes between, is the same frames
        // stepped whole; by runs and every cell alone. In the two frames that
        // compute every cell in one step, the runs of twenty are computed as
        // runs and the run of five by its program; the order's runs being
        // twenty, twenty, five and twenty, a run the walk skipped leaves the
        // cells of a run of twenty alone.
        let rows = concat!(
            "(sheet \"S\" visible)\n",
            "(cell \"S\" \"A1:A30\" 1)\n",
            "(sheet \"T\" visible)\n",
            "(formula \"T\" \"A1:A20\" \"=S.last!A1+1\")\n",
            "(formula \"T\" \"B1\" \"=S.last!A1*7\")\n",
            "(formula \"T\" \"C1:C20\" \"=S.last!A1-2\")\n",
            "(formula \"T\" \"D1\" \"=S!A1+S.last!A2\")\n",
            "(formula \"T\" \"E1:E20\" \"=A1*2\")\n",
            "(formula \"T\" \"F1\" \"=E1+B1\")\n",
            "(formula \"T\" \"G1:G5\" \"=S.last!A1*3\")\n",
        );
        let of = load(rows).shapes.order().len();
        let twenties: BTreeSet<usize> = {
            let m = load(rows);
            [(1, 1, 1), (1, 1, 3), (1, 1, 5)]
                .iter()
                .map(|id| m.shapes.place(*id).expect("placed").shape)
                .collect()
        };
        for by_runs in [true, false] {
            for budget in 1..=of {
                let mut stepped = load(rows);
                stepped.by_runs = by_runs;
                let mut whole = load(rows);
                for frame in 1..=3u32 {
                    vector::taken::begin();
                    while !stepped.step(budget).done {}
                    let as_runs: BTreeSet<usize> =
                        vector::taken::end().into_iter().map(|(s, _)| s).collect();
                    // A step's piece of a run shorter than `SHORT` is its
                    // program's, so the runs are whole in a frame of one step.
                    if frame <= 2 && budget == of {
                        let want = if by_runs {
                            twenties.clone()
                        } else {
                            BTreeSet::new()
                        };
                        assert_eq!(as_runs, want, "by runs {by_runs}, frame {frame}");
                    }
                    assert!(whole.step(0).done);
                    assert_eq!(
                        values(&stepped),
                        values(&whole),
                        "budget {budget}, by runs {by_runs}, frame {frame}"
                    );
                    let w = format!("(cell \"S\" \"A1:A30\" {})", frame + 1);
                    assert!(stepped.write(&w).is_ok() && whole.write(&w).is_ok());
                }
            }
        }
    }

    #[test]
    fn the_census_names_a_shape_left_to_ctx() {
        // KERNEL.44: a shape with cells alone whose parse a program does not
        // hold says why in its row; a shape lowered, a shape whose cells all
        // stand in runs, a shape computed every frame and one calling a
        // function the library lacks, each of which says why already, say
        // nothing of it.
        let f = |cell: &str, text: &str| {
            format!("(formula \"S\" \"{cell}\" {})\n", crate::rows::quoted(text))
        };
        let rows = [
            "(sheet \"S\" visible)\n(cell \"S\" \"A1:A30\" 1)\n(name \"Rate\" \"S!$A$1\")\n"
                .to_string(),
            f("B1", "=A1*2"),
            f("C1", "=A1&\"x\""),
            f("D1", "=A1:A2+1"),
            f("E1", "=IF(A1,1,TRUE)"),
            f("F1", "=Rate+1"),
            f("G1:G30", "=A1&\"y\""),
            f("H1", "=EMPTIES(CHOOSE(1,,2))"),
        ]
        .concat();
        let m = load_programs(&rows, "the census");
        let census = m.census();
        for (cell, want) in [
            ("B1", None),
            ("C1", Some("text")),
            ("D1", Some("range")),
            ("E1", Some("kinds")),
            ("F1", None),
            ("G1", None),
            ("H1", Some("omitted")),
        ] {
            let row = census
                .lines()
                .find(|l| l.starts_with("(shape") && l.contains(&format!("\"S\" \"{cell}\"")))
                .unwrap_or_else(|| panic!("{cell}'s row in\n{census}"));
            match want {
                Some(w) => assert!(row.ends_with(&format!(" (unlowered {w}))")), "{row}"),
                None => assert!(!row.contains("(unlowered"), "{row}"),
            }
        }
        // The plan's totals unchanged by the item: no row of a census whose
        // shapes are all lowered moves.
        assert!(census.starts_with("(plan (shapes 7)"), "{census}");
    }

    #[test]
    fn a_program_reads_a_cell_as_the_machine_reads_it() {
        // KERNEL.44: the machine's read of a cell's number, made without its
        // value, is `Env::read`'s answer as a register holds it at every cell
        // of every sheet and twin, past their extents too, after every frame
        // and every write: a number, 0 for nothing, `None` for anything else.
        for seed in 1..=20u64 {
            let mut d = Dice(seed ^ 0x5151);
            let rows = program_grid(&mut d);
            let mut m = load_programs(&rows, &format!("seed {seed}"));
            for frame in 1..=3 {
                assert!(m.step(0).done);
                for sheet in 0..m.book.sheets.len() {
                    for row in 1..=32u32 {
                        for col in 1..=13u32 {
                            let id = (sheet, row, col);
                            assert_eq!(
                                Source::number(&m, id),
                                program::number_of(Env::read(&m, id)),
                                "seed {seed}, frame {frame}, {id:?}"
                            );
                        }
                    }
                }
                let w = format!(
                    "(cell \"A\" \"A{}\" {})",
                    1 + d.below(20),
                    d.pick(&["\"t\"", "true", "(error \"#N/A\")", "8"])
                );
                assert!(m.write(&w).is_ok(), "{w}");
            }
        }
    }

    // ---- KERNEL.31: a frame's own sheets kept as arrays ----

    /// Each sheet of the rows a machine keeps as an array holds at every slot
    /// what a read of its cell answers now, to the bit, over its sheet's
    /// extent, and counts the reasons its formula cells hold; each kept where
    /// the plan reads it so and its cells stand close enough together, and no
    /// other (KERNEL.31).
    fn own_is_the_frame(m: &Machine, when: &str) {
        for s in 0..m.sheets {
            let sheet = &m.book.sheets[s];
            let fits = dense::fits(m.extents[s], dense::limit(sheet.cells.len(), CELLS as u64));
            assert_eq!(
                m.own[s].is_some(),
                m.arrays && m.shapes.read_as_array(s) && fits,
                "{when}: sheet {s} kept as an array"
            );
            let Some(d) = &m.own[s] else {
                assert_eq!(m.own_reasons[s], 0, "{when}: sheet {s}, no array, no count");
                continue;
            };
            assert!(
                d.stands_over(m.extents[s]),
                "{when}: sheet {s}'s array over its extent"
            );
            assert!(
                d.keeps_numbers(),
                "{when}: sheet {s}'s numbers beside its slots"
            );
            for i in 0..d.slots() {
                let (slot, (row, col)) = d.at(i);
                let want = m.slot_value((s, row, col));
                assert!(
                    same_value(slot, &want),
                    "{when}: sheet {s}, row {row}, column {col}: {slot:?} for {want:?}"
                );
            }
            let reasons = m
                .shapes
                .order()
                .iter()
                .zip(&m.values)
                .filter(|(id, v)| id.0 == s && matches!(v, Some(Computed::NotComputed(_))))
                .count();
            assert_eq!(
                m.own_reasons[s] as usize, reasons,
                "{when}: sheet {s}'s reasons"
            );
        }
    }

    /// A grid of two sheets whose formulas stand in bands down a column,
    /// each band one shape whose cells, one below the other, read the cell
    /// beside each, the cells above the band, the value columns' ranges and
    /// the sheet before's (KERNEL.31): `grid_of`'s bands one column wide,
    /// every reference pointing before its cell in reading order, so that no
    /// grid holds a cycle, and a run's cells read a column of numbers.
    fn stacked_grid(d: &mut Dice) -> String {
        let rows = 4 + d.below(40) as u32;
        let mut out = String::new();
        for (sheet, before) in [("A", None), ("B", Some("A"))] {
            out.push_str(&format!("(sheet \"{sheet}\" visible)\n"));
            for row in 1..=20 {
                for col in ["A", "B"] {
                    if let Some(v) = datum(d) {
                        out.push_str(&format!("(cell \"{sheet}\" \"{col}{row}\" {v})\n"));
                    }
                }
            }
            for c in 2..8 {
                let (col, left) = (LETTERS[c], LETTERS[c - 1]);
                let mut top = 1;
                while top <= rows {
                    let bottom = (top + d.below(rows as usize) as u32).min(rows);
                    let at = Band {
                        before,
                        top,
                        rows,
                        left,
                    };
                    let f = expr(d, &at, 0);
                    let f = match d.below(3) {
                        0 => format!("{left}{top}+({f})"),
                        1 => format!("{left}{top}*2"),
                        _ => f,
                    };
                    let f = crate::rows::quoted(&format!("={f}"));
                    out.push_str(&format!(
                        "(formula \"{sheet}\" \"{col}{top}:{col}{bottom}\" {f})\n"
                    ));
                    top = bottom + 1;
                }
            }
        }
        out.push_str("(name \"Rate\" \"A!$A$2\")\n");
        out.push_str("(name \"Box\" \"A!$A$1:$B$3\")\n");
        out.push_str("(name \"Prev\" \"A.last!$C$1\")\n");
        out.push_str("(name \"Band\" \"A!$C$1:$D$2\")\n");
        out
    }

    /// The same rows loaded twice: one machine keeping its frame's own
    /// sheets as arrays, as every machine does, and one keeping none, its
    /// sheets read through their cells as before KERNEL.31.
    fn kept_and_not(rows: &str) -> (Machine, Machine) {
        let kept = load(rows);
        let mut cells = load(rows);
        cells.arrays = false;
        cells.own_made(None);
        (kept, cells)
    }

    #[test]
    fn a_frame_s_own_sheets_read_as_arrays_are_their_cells_read_one_by_one() {
        // KERNEL.31: grids of generated formulas, in bands across a row and
        // down a column, reading ranges and cells of their own sheet and of
        // another, stepped in slices of any budget, against the same grids
        // whose frames keep no arrays, frame for frame and bit for bit, and
        // the first frame against `Calc`; every array held at every yield and
        // after every write to what a read of each of its cells answers, and
        // a view at every yield the last complete frame, never the array's;
        // writes between frames of values, of texts, of a formula leaving its
        // cell not computed, which steps its sheet's array aside, and of a
        // cell past a sheet's extent, which makes its array again, each write
        // planning again keeping the arrays over the extents that stand.
        vector::sliced::take();
        let (mut aside, mut kept_over, mut made_again, mut kept) = (0, 0, 0, 0);
        for seed in 1..=60u64 {
            for stacked in [false, true] {
                let mut d = Dice(seed ^ if stacked { 0x3131 } else { 0 });
                let rows = if stacked {
                    stacked_grid(&mut d)
                } else {
                    generated_grid(&mut d)
                };
                let (mut m, mut cells) = kept_and_not(&rows);
                let calc: BTreeMap<CellId, String> = {
                    let calc = crate::calc::Calc::run(&m.book, &m.library);
                    let spell =
                        |id: &CellId| format!("{:?}", calc.computed(*id).expect("computed"));
                    m.shapes.order().iter().map(|id| (*id, spell(id))).collect()
                };
                let shown = |m: &Machine| -> Vec<String> {
                    let order = m.shapes.order().iter();
                    order.map(|&id| format!("{:?}", m.computed(id))).collect()
                };
                for frame in 1..=5 {
                    let when = format!("seed {seed}, stacked {stacked}, frame {frame}");
                    let last = shown(&m);
                    loop {
                        let s = m.step(1 + d.below(300));
                        own_is_the_frame(&m, &when);
                        if s.done {
                            break;
                        }
                        assert_eq!(shown(&m), last, "{when}: a view while the step yields");
                    }
                    assert!(cells.step(0).done);
                    assert_eq!(values(&m), values(&cells), "{when}\n{rows}");
                    if frame == 1 {
                        assert_eq!(values(&m), calc, "{when}: Calc\n{rows}");
                    }
                    kept += m.own.iter().filter(|o| o.is_some()).count();
                    let write = match d.below(6) {
                        0 => format!("(cell \"A\" \"A{}\" 7)", 1 + d.below(20)),
                        1 => format!("(cell \"B\" \"B{}\" \"y\")", 1 + d.below(20)),
                        2 => format!(
                            "(cell \"A\" \"{}{}\" 2)",
                            d.pick(&LETTERS[2..8]),
                            1 + d.below(4)
                        ),
                        3 => format!(
                            "(formula \"A\" \"{}{}\" \"=$A$1:$A$2\")",
                            d.pick(&LETTERS[2..8]),
                            1 + d.below(4)
                        ),
                        4 => format!("(cell \"B\" \"J{}\" 1)", 30 + d.below(40)),
                        _ => continue,
                    };
                    let (a, b) = (m.write(&write), cells.write(&write));
                    assert_eq!(a.is_ok(), b.is_ok(), "{when}: {write}: {a:?} {b:?}");
                    own_is_the_frame(&m, &format!("{when}, after {write}"));
                }
                aside += m.aside.get();
                kept_over += m.kept_over;
                made_again += m.made_again;
            }
        }
        let down = vector::sliced::take()[vector::sliced::DOWN];
        assert!(kept > 0, "a frame's own sheet kept as an array");
        assert!(down > 0, "a column of a frame's numbers read whole");
        assert!(aside > 0, "a sheet's array stepped aside for a reason");
        assert!(kept_over > 0, "an array kept over a write that plans again");
        assert!(made_again > 0, "an array made again past its extent");
    }

    #[test]
    fn a_range_over_a_frame_s_sheet_answers_a_cell_not_computed() {
        // KERNEL.31: a range of a frame's own sheet holding a cell not
        // computed answers the cell's reason, as a walk of its cells does, a
        // run's fold and a cell alone's walk alike: the sheet steps its array
        // aside while a reason stands, and offers it again once none does.
        // Beside them a twin's range skips the cell, as
        // `a_range_over_a_twin_skips_a_cell_not_computed_as_it_did` pins it.
        let rows = concat!(
            "(sheet \"S\" visible)\n",
            "(cell \"S\" \"B1:B2\" 1)\n",
            "(formula \"S\" \"A1\" \"=B1:B2\")\n",
            "(formula \"S\" \"C1:C8\" \"=SUM(A$1:B$1)+ROW()\")\n",
            "(formula \"S\" \"D1\" \"=SUM(A1:B2)\")\n",
            "(formula \"S\" \"E1:E8\" \"=SUM(S.last!A$1:B$1)\")\n",
        );
        let (mut m, mut cells) = kept_and_not(rows);
        assert!(m.own[0].is_some(), "the sheet kept as an array");
        for frame in 1..=2 {
            assert!(m.step(0).done && cells.step(0).done);
            assert_eq!(values(&m), values(&cells), "frame {frame}");
            own_is_the_frame(&m, &format!("frame {frame}"));
        }
        for addr in ["A1", "C1", "C8", "D1"] {
            assert!(
                matches!(at(&m, "S", addr), Some(Computed::NotComputed(_))),
                "{addr}: {:?}",
                at(&m, "S", addr)
            );
        }
        assert_eq!(at(&m, "S", "E1"), Some(Computed::Value(Value::Number(1.0))));
        // `A1`'s reason, and the nine cells that answer it.
        assert_eq!(m.own_reasons[0], 10);
        assert!(m.aside.get() > 0, "the array stepped aside");
        // A1 computed: no reason stands, and the array is read again.
        let w = "(formula \"S\" \"A1\" \"=B1+B2\")";
        assert!(m.write(w).is_ok() && cells.write(w).is_ok());
        assert!(m.step(0).done && cells.step(0).done);
        let aside = m.aside.get();
        assert!(m.step(0).done && cells.step(0).done);
        assert_eq!(values(&m), values(&cells));
        assert_eq!(m.own_reasons[0], 0);
        assert_eq!(
            m.aside.get(),
            aside,
            "no array stepped aside once none stands"
        );
        assert_eq!(at(&m, "S", "C1"), Some(Computed::Value(Value::Number(4.0))));
        assert_eq!(at(&m, "S", "D1"), Some(Computed::Value(Value::Number(4.0))));
        own_is_the_frame(&m, "after the reason");
    }

    #[test]
    fn a_frame_s_array_keeps_a_cell_not_reached_and_takes_each_write() {
        // KERNEL.31: a frame computes what changed, so a slot whose cell it
        // does not reach keeps the last frame's value; a value written lands
        // in its slot at the write, before any frame; a formula written
        // inside the sheet's extent plans again and keeps the array, its slot
        // nothing until its first frame; a cell written past the extent makes
        // the array again over the new one.
        let rows = concat!(
            "(sheet \"S\" visible)\n",
            "(cell \"S\" \"A1:A8\" 1)\n",
            "(formula \"S\" \"B1:B8\" \"=A1*2\")\n",
            "(formula \"S\" \"C1:C8\" \"=B1+1\")\n",
        );
        let mut m = load(rows);
        let slot = |m: &Machine, row: u32, col: u32| {
            m.own[0].as_ref().and_then(|d| d.get(row, col).cloned())
        };
        assert!(m.step(0).done);
        own_is_the_frame(&m, "frame 1");
        assert!(m.write("(cell \"S\" \"A3\" 5)").is_ok());
        assert_eq!(
            slot(&m, 3, 1),
            Some(Value::Number(5.0)),
            "a value at its write"
        );
        own_is_the_frame(&m, "after a value");
        assert!(m.step(0).done);
        assert_eq!(slot(&m, 3, 2), Some(Value::Number(10.0)));
        assert_eq!(slot(&m, 3, 3), Some(Value::Number(11.0)));
        assert_eq!(
            slot(&m, 1, 2),
            Some(Value::Number(2.0)),
            "a cell not reached"
        );
        own_is_the_frame(&m, "frame 2");
        let first = |m: &Machine| m.own[0].as_ref().map(|d| d.slot_at(0) as *const Value);
        let was = first(&m);
        assert!(m.write("(formula \"S\" \"B4\" \"=A4*3\")").is_ok());
        assert_eq!(first(&m), was, "the array kept, in its own memory");
        assert_eq!((m.kept_over, m.made_again), (1, 0));
        assert_eq!(
            slot(&m, 4, 2),
            Some(Value::Empty),
            "a formula with no value yet"
        );
        own_is_the_frame(&m, "after a formula");
        assert!(m.step(0).done);
        assert_eq!(slot(&m, 4, 2), Some(Value::Number(3.0)));
        assert_eq!(slot(&m, 4, 3), Some(Value::Number(4.0)));
        own_is_the_frame(&m, "frame 3");
        assert!(m.write("(cell \"S\" \"E12\" 1)").is_ok());
        assert_eq!((m.kept_over, m.made_again), (1, 1));
        assert_eq!(
            slot(&m, 12, 5),
            Some(Value::Number(1.0)),
            "over the new extent"
        );
        own_is_the_frame(&m, "after a cell past the extent");
        assert!(m.step(0).done);
        own_is_the_frame(&m, "frame 4");
    }

    #[test]
    fn a_sheet_no_read_takes_keeps_no_array() {
        // KERNEL.31: a sheet whose cells only cells alone read, a running
        // balance's, keeps no array, nor one read only by cells every `$`
        // holds, nor one a range reads that is too sparse for an array; and
        // a machine whose frames keep none keeps none where one would be.
        for rows in [
            concat!(
                "(sheet \"S\" visible)\n",
                "(cell \"S\" \"A1\" 1)\n",
                "(formula \"S\" \"A2:A20\" \"=MOD(ROW()*7,13)\")\n",
                "(formula \"S\" \"B1\" \"=A1\")\n",
                "(formula \"S\" \"B2:B20\" \"=B1+A2\")\n",
            ),
            concat!(
                "(sheet \"S\" visible)\n",
                "(cell \"S\" \"A1\" 1)\n",
                "(formula \"S\" \"B1:B8\" \"=$A$1+ROW()\")\n",
            ),
            concat!(
                "(sheet \"S\" visible)\n",
                "(cell \"S\" \"A1:A8\" 1)\n",
                "(cell \"S\" \"XFA1\" 1)\n",
                "(formula \"S\" \"B1:B8\" \"=SUM(A1:A2)\")\n",
            ),
        ] {
            let mut m = load(rows);
            assert!(m.step(0).done);
            assert!(m.own[0].is_none(), "{rows}");
            own_is_the_frame(&m, rows);
        }
        let (kept, cells) = kept_and_not(concat!(
            "(sheet \"S\" visible)\n",
            "(cell \"S\" \"A1:A8\" 1)\n",
            "(formula \"S\" \"B1:B8\" \"=SUM(A1:A2)\")\n",
        ));
        assert!(kept.own[0].is_some() && cells.own[0].is_none());
    }

    #[test]
    fn cells_one_below_the_other_read_a_column_of_numbers() {
        // KERNEL.31: a run's cells one below the other read the cell beside
        // each, a reference no `$` holds, as one column of the sheet's
        // numbers; a text in the column, and a reference moved past the
        // array's last row or off the sheet's first, read cell by cell, the
        // lane turning to values at the first cell not a number; a reference
        // whose row a `$` holds, the same cell at each, never as a column;
        // each frame against the same cells read one by one.
        let rows = concat!(
            "(sheet \"S\" visible)\n",
            "(cell \"S\" \"A1:A6\" 2)\n",
            "(cell \"S\" \"A7\" \"x\")\n",
            "(cell \"S\" \"A8:A20\" 2)\n",
            "(cell \"S\" \"G1:G20\" 3)\n",
            "(formula \"S\" \"B1:B20\" \"=A1*2\")\n",
            "(formula \"S\" \"C1:C20\" \"=A2+1\")\n",
            "(formula \"S\" \"D1:D20\" \"=A1\")\n",
            "(formula \"S\" \"E1:E20\" \"=SUM(D1:D20)+B1\")\n",
            "(formula \"S\" \"H1:H20\" \"=G1+1\")\n",
            "(formula \"S\" \"I1:I20\" \"=G2*H1\")\n",
            "(formula \"S\" \"K1:K20\" \"=ROW()\")\n",
            "(formula \"S\" \"J1:J20\" \"=K$1*10\")\n",
        );
        let (mut m, mut cells) = kept_and_not(rows);
        vector::sliced::take();
        let mut down = 0;
        for frame in 1..=3 {
            assert!(m.step(0).done && cells.step(0).done);
            down += vector::sliced::take()[vector::sliced::DOWN];
            assert_eq!(values(&m), values(&cells), "frame {frame}");
            own_is_the_frame(&m, &format!("frame {frame}"));
            let w = format!("(cell \"S\" \"A{}\" {})", 1 + frame, frame);
            assert!(m.write(&w).is_ok() && cells.write(&w).is_ok());
        }
        assert!(down > 0, "a column of numbers read whole");
        // `G21`, past the array's last row, is nothing.
        assert_eq!(
            at(&m, "S", "I20"),
            Some(Computed::Value(Value::Number(0.0)))
        );
        assert_eq!(
            at(&m, "S", "C20"),
            Some(Computed::Value(Value::Number(1.0)))
        );
        assert!(matches!(
            at(&m, "S", "B7"),
            Some(Computed::Value(Value::Error(_)))
        ));
        assert_eq!(
            at(&m, "S", "J20"),
            Some(Computed::Value(Value::Number(10.0)))
        );
    }

    /// The formula cells at which two spellings of a frame's values differ,
    /// by sheet: a value gained, lost or spelled otherwise, to the bit
    /// (KERNEL.48).
    fn differing(
        before: &BTreeMap<CellId, String>,
        after: &BTreeMap<CellId, String>,
    ) -> BTreeMap<usize, usize> {
        let mut by_sheet = BTreeMap::new();
        let ids: BTreeSet<&CellId> = before.keys().chain(after.keys()).collect();
        for id in ids {
            if before.get(id) != after.get(id) {
                *by_sheet.entry(id.0).or_insert(0) += 1;
            }
        }
        by_sheet
    }

    /// A machine's counts after a frame held to the cells it changed, by
    /// sheet: the count, each sheet's share, a twin's none, none past the
    /// grid's sheets, and the cells computed the frame's own count of them.
    fn counts_held(m: &Machine, by_sheet: &BTreeMap<usize, usize>, when: &str) {
        let total: usize = by_sheet.values().sum();
        assert_eq!(m.changed(), total, "{when}: the count");
        for s in 0..m.book.sheets.len() {
            let share = by_sheet.get(&s).copied().unwrap_or(0);
            assert_eq!(m.changed_on(s), Some(share), "{when}: sheet {s}'s count");
        }
        let past = m.book.sheets.len();
        assert_eq!(m.changed_on(past), None, "{when}: past the sheets");
        assert_eq!(m.reached(), m.computed, "{when}: the cells computed");
    }

    #[test]
    fn a_frame_s_counts_are_the_cells_it_changed_however_it_is_computed() {
        // KERNEL.48: after every done step the count is the formula cells
        // whose values differ to the bit across the frame, each sheet's its
        // share, and the cells computed the frame's own count of them; the
        // same count with each of the machine's paths turned off, the first
        // machine stepped under a budget, its counts during a yield the
        // last complete frame's.
        let (mut changing, mut still, mut yielded, mut followed) = (0, 0, 0, 0);
        for seed in 1..=40u64 {
            let mut d = Dice(seed ^ 0x48);
            let churn = seed % 2 == 0;
            let rows = grid_of(&mut d, churn);
            let mut machines: Vec<Machine> = (0..6).map(|_| load(&rows)).collect();
            machines[1].incremental = false;
            machines[2].by_runs = false;
            machines[3].lowered = false;
            machines[4].copies = false;
            machines[5].arrays = false;
            machines[5].own_made(None);
            for m in &machines {
                counts_held(m, &BTreeMap::new(), &format!("seed {seed}: loaded"));
            }
            for frame in 1..=8u32 {
                let when = format!("seed {seed}, frame {frame}");
                if churn {
                    let clock = format!("(cell \"Clock\" \"B1\" {frame})");
                    for m in &mut machines {
                        assert!(m.write(&clock).is_ok(), "{when}: the Clock");
                    }
                }
                let before = values(&machines[0]);
                let last = (machines[0].changed(), machines[0].reached());
                let budget = 1 + d.below(64);
                if !machines[0].step(budget).done {
                    yielded += 1;
                    let now = (machines[0].changed(), machines[0].reached());
                    assert_eq!(now, last, "{when}: the counts during a yield");
                    while !machines[0].step(budget).done {}
                }
                for m in &mut machines[1..] {
                    assert!(m.step(0).done, "{when}");
                }
                let by_sheet = differing(&before, &values(&machines[0]));
                for (k, m) in machines.iter().enumerate() {
                    counts_held(m, &by_sheet, &format!("{when}, machine {k}\n{rows}"));
                }
                let of = machines[1].shapes.order().len();
                assert_eq!(machines[1].reached(), of, "{when}: every cell computed");
                followed += usize::from(machines[0].reached() < of);
                match machines[0].changed() {
                    0 => still += 1,
                    _ => changing += 1,
                }
                let Some(write) = write_between(&mut d, frame) else {
                    continue;
                };
                let first = machines[0].write(&write).is_ok();
                for m in &mut machines[1..] {
                    assert_eq!(m.write(&write).is_ok(), first, "{when}: {write}");
                }
            }
        }
        // Measured 2026-10-18: of 320 frames, 226 changing, 94 still, 315
        // yielding and 104 computing what changed.
        assert!(
            changing >= 110 && still >= 45 && yielded >= 150 && followed >= 50,
            "{changing} frames changing, {still} still, {yielded} yielding, {followed} \
             computing what changed: too few of a kind for the counts to be held over"
        );
    }

    #[test]
    fn a_frame_s_count_names_each_kind_of_change() {
        // KERNEL.48, the named cases: a grid with no value rows counts every
        // formula at its first frame; a frame that changes nothing counts
        // none and, still, computes none; a value write counts its readers
        // on each sheet and never the cell written; a reason turned to
        // another counts; a cell whose one reader reads its twin counts
        // nothing at the write's frame and its reader at the next, so that
        // the frame right after a write is no fixed point; a 0 turned to -0
        // counts, though a window's changes name no row of it; and a
        // formula written counts as a first value, its bits the old ones.
        let mut m = load(concat!(
            "(sheet \"S\" visible)\n",
            "(cell \"S\" \"A1\" 0)\n",
            "(cell \"S\" \"A2\" 1)\n",
            "(cell \"S\" \"A3\" 1)\n",
            "(cell \"S\" \"A5\" -0)\n",
            "(cell \"S\" \"A6\" 1)\n",
            "(formula \"S\" \"B1\" \"=IF(A6=1,A1,A5)\")\n",
            "(formula \"S\" \"B2\" \"=A2*2\")\n",
            "(formula \"S\" \"B3\" \"=S.last!A3*2\")\n",
            "(formula \"S\" \"B4\" \"=IF(A2>5,A1:A2,A2:A3)\")\n",
            "(sheet \"T\" visible)\n",
            "(formula \"T\" \"A1\" \"=S!B2+1\")\n",
        ));
        let counts = |m: &Machine| (m.changed(), m.changed_on(0), m.changed_on(1));
        let step = |m: &mut Machine| assert!(m.step(0).done);
        let reason = |m: &Machine| match at(m, "S", "B4") {
            Some(Computed::NotComputed(Reason::Construct(c))) => c,
            other => panic!("B4: {other:?}"),
        };
        assert_eq!(counts(&m), (0, Some(0), Some(0)), "before the first frame");
        assert_eq!((m.changed_on(2), m.changed_on(4)), (Some(0), None));
        step(&mut m);
        assert_eq!(
            counts(&m),
            (5, Some(4), Some(1)),
            "every formula's first value"
        );
        assert_eq!(m.reached(), 5);
        step(&mut m);
        assert_eq!(m.changed(), 0, "every cell computed again, alike");
        step(&mut m);
        assert_eq!(
            (m.changed(), m.reached()),
            (0, 0),
            "still, and nothing computed"
        );
        // A value write: its readers on both sheets, never the cell.
        assert_eq!(m.write("(cell \"S\" \"A2\" 3)"), Ok(1));
        step(&mut m);
        assert_eq!(counts(&m), (2, Some(1), Some(1)), "B2 and T!A1");
        assert_eq!(reason(&m), "A2:A3");
        // A reason turned to another; the write before it leaves the last
        // frame's counts as they were.
        assert_eq!(m.write("(cell \"S\" \"A2\" 7)"), Ok(1));
        assert_eq!(counts(&m), (2, Some(1), Some(1)), "a value written");
        step(&mut m);
        assert_eq!(reason(&m), "A1:A2");
        assert_eq!(
            counts(&m),
            (3, Some(2), Some(1)),
            "B2, B4's reason and T!A1"
        );
        // A cell read through its twin alone: nothing at the write's frame,
        // its reader at the next.
        assert_eq!(m.write("(cell \"S\" \"A3\" 5)"), Ok(1));
        step(&mut m);
        assert_eq!(m.changed(), 0, "the write's frame, B3 reading A3's twin");
        step(&mut m);
        assert_eq!(counts(&m), (1, Some(1), Some(0)), "B3, a frame later");
        assert_eq!(number(&m, "S", "B3"), 10.0);
        step(&mut m);
        assert_eq!(m.changed(), 0, "and then still");
        // A 0 turned to -0, which no record row of B1 spells otherwise: B1
        // reads A5's -0 in place of A1's 0. (A write of -0 over a cell
        // holding 0 is no change to the write, which keeps the 0.)
        assert_eq!(m.write("(cell \"S\" \"A6\" 2)"), Ok(1));
        step(&mut m);
        assert_eq!(counts(&m), (1, Some(1), Some(0)), "B1, 0 to -0");
        assert!(number(&m, "S", "B1").is_sign_negative());
        let changes = changes_of(&m, "S", None);
        assert!(
            changes.contains("\"A6\"") && !changes.contains("\"B1\""),
            "the changes name A6 alone:\n{changes}"
        );
        // A formula written: its first value, the others where they moved,
        // the last frame's counts as they were until the next frame ends.
        assert_eq!(m.write("(formula \"S\" \"B2\" \"=A2+A2\")"), Ok(1));
        assert_eq!(counts(&m), (1, Some(1), Some(0)), "a formula written");
        step(&mut m);
        assert_eq!(
            counts(&m),
            (1, Some(1), Some(0)),
            "B2's first value, 14 again"
        );
        assert_eq!(m.reached(), 5, "every cell computed");
    }

    #[test]
    fn a_save_s_first_frame_counts_what_its_value_rows_did_not_hold() {
        // KERNEL.48: a fresh grid's first frame counts every formula, each a
        // first value; a save's, whose value rows hold what its frame
        // computes again, none, though it computes every cell; and a save
        // resumed at its frame answers 0 before its first, as a grid just
        // loaded does.
        let rows = concat!(
            "(sheet \"S\" visible)\n",
            "(cell \"S\" \"A1\" 3)\n",
            "(formula \"S\" \"B1\" \"=A1*2\")\n",
            "(formula \"S\" \"C1\" \"=B1+S.last!A1\")\n",
        );
        let mut m = load(rows);
        assert!(m.step(0).done);
        assert_eq!((m.changed(), m.reached()), (2, 2));
        assert!(m.step(0).done);
        assert_eq!(m.changed(), 0);
        let mut again = load(&save(&m));
        again.resume_at(m.frame()).expect("no frame in progress");
        assert_eq!((again.changed(), again.reached(), again.frame()), (0, 0, 2));
        assert!(again.step(0).done);
        assert_eq!(
            (again.changed(), again.reached()),
            (0, 2),
            "its values held"
        );
    }
}
