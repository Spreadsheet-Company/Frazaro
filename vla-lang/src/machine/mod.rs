//! The machine (KERNEL.22, 2026-10-09): a grid kept in memory and the four
//! calls an engine makes of it, `load`, `write`, `step` and `view`
//! (`Alonzo/SPEC.md` section 4). This amends `web/CALLOSUM.md` section 7
//! decision 1, the pure function with no model handle, for an engine's door
//! alone, on the measurement that decision asked for: a frame of 64,000
//! cells thirty times a second cannot rebuild its grid from text. Frazaro's
//! own doors keep the pure function.
//!
//! **The grid** (KERNEL.49): the machine's own, beside the file model's
//! [`Workbook`], which it lends when asked ([`Machine::workbook`]): the rows'
//! sheets, then a twin `X.last` for every sheet `X`, in the same order,
//! hidden, holding `X`'s values as they were when the last step finished
//! (the page's section 2). A sheet is its name and settings, its fills, each
//! one formula written over a rectangle and held as the pieces the writes
//! into it leave (`book`), its formulas alone, each a cell's own text, and
//! the few cells' formats; every value it holds stands by place in its store
//! (`store`), so that no cell of a fill is ever made, and a load makes the
//! grid from the draft's records ([`load`]) and no cell of a range. Every
//! reader, the plan, the evaluator, the record and the plane, reads it as it
//! reads a workbook, through the crate's one reading of a grid
//! (`crate::grid`), so that each record is printed by the code that prints a
//! workbook's. A twin is an ordinary sheet to every reader, so
//! `Screen.last!B2` is a reference like any other and adds no edge, a twin's
//! cells being values. A twin is read-only: only the end of a step writes it.
//! A twin keeps no cell of its own: where a formula reads it, its values stand
//! in a store beside its sheet's, and any other twin is made from its sheet's
//! store as the last frame's end left it, wherever it is read.
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
//! that changed a formula, or after a twin lost a cell. A value cell a write
//! added, and a twin cell a frame's end gained, are followed as a changed one
//! is (KERNEL.56): an added cell moves no formula's place and is found by its
//! place, and a twin's extent grows by the cells it gained. A few functions
//! read a sheet's extent itself, an approximate match's bisection and a
//! database's last row (the library declares them), so a sheet whose extent
//! moved, at a write or a twin at a frame's end, has every formula calling
//! one over it computed at the next frame. The budget counts the cells of
//! the order passed, computed or not.
//!
//! The frame's values are kept once, by place (KERNEL.49): each rows'
//! sheet keeps every value its cells hold in one store ([`store`]), an array
//! over its extent where its cells stand close enough together ([`Dense`],
//! nine bytes a slot), a value cell's written at the write that changes it
//! and a formula's overwritten as its cell computes a value that differs, a
//! formula's reason and a formula with no value yet a slot's kinds; a step
//! that yields keeps each value it overwrote beside, so a view during a
//! yield shows the last complete frame. A range of the frame's own cells is
//! folded and walked by index, and a run's cells read one of them by slices
//! (KERNEL.31; KERNEL.43), as a twin's array is read; the order puts each
//! reader after every cell it reads, so the slots it reads are this frame's.
//! When a frame completes, each cell it changed and each written since is
//! told against its twin as the last frame's end left it, and copied into
//! its twin's store where a formula reads the twin, and into its sheet's
//! plane where an engine views one, so that a view of it is a copy
//! ([`plane::Raster`], KERNEL.24), however many changed (KERNEL.36): the
//! next frame's read of a twin is an index and its range a walk of the array
//! (KERNEL.20). A write that grows a sheet's extent takes its store and its
//! plane over the new one where they stand, rows added in place where the
//! bottom alone moved, and the frame's end the twin's store the same way
//! (KERNEL.56); a twin's store that cannot follow its sheet's, and a plane
//! not kept, are made again from the sheet. The step computes values and
//! never formulas (`AD-7`): the formula cells after a step are the formula
//! cells before it.
//!
//! **The changes** (KERNEL.42). The twin cells a frame's end changed are
//! kept until the next frame's end, a bit each (KERNEL.49), and with them the
//! few whose change a record spells as before, so that a view of a window's
//! changes answers from them alone, the cells whose rows in the window's
//! values the last complete frame changed, at the cost of the changes and not
//! the window; or answers whole where they cannot say: before the first
//! frame, after a frame whose end changed more twin cells than half the
//! cells, or made a twin's extent again, a twin cell losing a value, or
//! grew it (KERNEL.56), after a write that changed a formula or moved a
//! sheet's extent, which every window's record names in its head, and while
//! a value written since lies in the window.
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
//! same grid makes, numbered and anchored as a load's (KERNEL.34). A write
//! into part of a fill cuts it, the rest kept as at most four rectangles of
//! the same fill, its first cell moved where the write takes it (KERNEL.49's
//! decision 2); the write's values reach the stores when it settles, and a
//! refused write puts the formulas back as it found them.
//!
//! **The refusals** are the language's catalogue's: the `grid` family for
//! the machine's own, `view-*` for the view, `calc-*` for what stops a grid
//! being stepped. An engine adds its own (its manifest, its devices) around
//! these calls and never inside them.

mod bits;
mod book;
pub mod load;
pub mod plane;
mod store;

pub(crate) use store::SparseWalk;

use std::cell::Cell as Slot;
use std::collections::{BTreeMap, BTreeSet, HashMap};

use crate::calc::blocks::Locator;
use crate::calc::dense::{self, Slot as Stored};
use crate::calc::graph::{Area, Extent, Names};
use crate::calc::program::{Answer, Source};
use crate::calc::reach::Reached;
use crate::calc::shape::{spell_at, Shapes, RUN, SHORT};
use crate::calc::vector::{self, Lanes};
use crate::calc::{
    CellId, Computed, Ctx, Dense, Env, ErrorKind, Indexes, Library, Reason, Value, Values, Walked,
    Written,
};
use crate::form::Form;
use crate::fx::FxHashMap;
use crate::grid::{self, text_of, Cells, Constant, Formula, Holds, Kind, Masters, Model};
use crate::messages::{raise, Refusal};
use crate::projection::Window;
use crate::reader::read_forms;
use crate::refers::shift_a1_references;
use crate::rows::Visibility;
use crate::sheet::{
    parse_a1_range, strip_future_prefixes, A1Range, Cell, Content, RowSettings, Sheet, Styles,
    Workbook, MAX_COLUMN, MAX_ROW,
};
use crate::view::{changes_text, record_text, window_over};

use bits::Bits;
use book::{Book, Fills, Page};
use load::What;
pub use load::{read_rows, sheet_name_ok, thousands, twin_base, Draft, Drafted};
use store::Store;

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
    /// The grid, the machine's own (KERNEL.49's decision 1): the rows'
    /// sheets' names, settings, fills, formulas alone and formats, their
    /// values standing in `stores`; `book.pages[sheets..]` are their twins,
    /// in order, which hold no cell.
    book: Book,
    /// The rows' sheets.
    sheets: usize,
    library: Library,
    names: Names,
    extents: Vec<Extent>,
    shapes: Shapes,
    /// By rows' sheet, every value its cells hold, by place (KERNEL.49's
    /// decision 3, [`store`]): a value cell's content as written, and each
    /// formula's value of the frame, the last complete frame's, each
    /// overwritten while a frame is in progress as its cell computes a value
    /// that differs (KERNEL.20), its reason where it was not computed, or
    /// none yet. An array over the sheet's extent where its cells stand
    /// close enough together, read by the evaluator by index, by ranges and
    /// by slices (KERNEL.31; KERNEL.43), a reason among its slots; else a
    /// map of its cells.
    stores: Vec<Store>,
    /// Each value a frame in progress overwrote while a step yielded, by
    /// position, as the last complete frame left it (`None`, no value), so
    /// that a view before the frame completes reads the last complete frame.
    /// This map and the next are under the crate's own hash (KERNEL.25);
    /// nothing walks them, so their order is never read.
    before: FxHashMap<u32, Option<Computed>>,
    /// Each cell a write changed since the last frame's end, as its twin
    /// held it there (KERNEL.49): what a twin kept as no store is made from
    /// while a write since stands, and what the next frame's end tells the
    /// twin's changes against; room kept for [`WRITTEN`].
    written_before: FxHashMap<CellId, Held>,
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
    /// By rows' sheet, its plane as the last frame ended and every write
    /// since, kept for a sheet in `wanted`: a write gives each cell it changed
    /// its byte, the plane taken over an extent it grew (KERNEL.56), where it
    /// dropped the planes of the sheets it changed until the next frame's end
    /// made them again.
    planes: Vec<Option<plane::Raster>>,
    /// By rows' sheet, its twin's values as the last frame's end left them
    /// (KERNEL.49's decision 5), kept as a store where a formula reads the
    /// twin, and changed by nothing but a frame's end, a twin being refused
    /// to every write: an array over the sheet's extent where the sheet's
    /// store is one, read by index, by ranges and by slices (KERNEL.20), else
    /// a sparse sheet's store; a formula's reason among its slots, and a
    /// formula with no value nothing. `None` for a twin no formula reads,
    /// which a view, the grid lent and any other read make from its sheet's
    /// store as the last frame's end left it.
    twins: Vec<Option<Store>>,
    /// By rows' sheet, whether a twin cell lost its value in the frame in
    /// progress, its twin's extent made again at the frame's end and the next
    /// frame computing every cell (KERNEL.20); the rectangle of the twin
    /// cells that gained one, by which that end grows its twin's extent, the
    /// next frame following them as any change (KERNEL.56), where a gain was
    /// a loss's; and what that end copies the sheet's changes into, a bit for
    /// its twin's store and a bit for its plane, each made again from the
    /// sheet where it is not.
    lost: Vec<bool>,
    gained: Vec<Extent>,
    copied_into: Vec<u8>,
    /// The sheets whose extent moved since the last frame began, a rows'
    /// sheet's at the write that grew it and a twin's at the frame's end that
    /// grew it, each once (KERNEL.56): the next frame computes whole every
    /// shape that reads its extent itself.
    extent_moved: Vec<usize>,
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
    /// The block of the order the last cell a step found by its position
    /// stood in, so that a step passing from one position to a later one
    /// finds each cell by a step past the blocks between (KERNEL.49's
    /// decision 7).
    locator: Locator,
    /// Whether the next frame computes every cell (KERNEL.20): after a load,
    /// a write that changed a formula, a twin that lost a cell, or more
    /// changes than were kept; a value cell added and a twin cell gained are
    /// followed as any change is (KERNEL.56).
    every_next: bool,
    /// Whether the frame in progress computes every cell.
    every: bool,
    /// By position in the order, whether the frame in progress computes
    /// the cell there, a bit a cell and a summary bit a word ([`Bits`],
    /// KERNEL.49's decision 6).
    marked: Bits,
    /// By shape, whether the frame in progress computes every cell of it,
    /// and whether its parts the same at every cell are computed again at
    /// the frame's start.
    whole: Vec<bool>,
    recheck: Vec<bool>,
    /// The shapes made whole and the shapes rechecked by the frame in
    /// progress, each listed as it is first set (KERNEL.54): the frame's
    /// start asks the shapes rechecked alone for their parts, and its end
    /// clears what it set by the lists, and the marks by their summary, where
    /// it went over every shape and every word.
    wholes: Vec<u32>,
    rechecked: Vec<u32>,
    /// Each such part's value as a frame's start last computed it.
    parts: Vec<Option<Result<Value, Reason>>>,
    /// The formula cells whose values the frame in progress changed, a bit a
    /// position in the order (KERNEL.49's decision 6), where a list held room
    /// for every formula cell (KERNEL.36).
    changed: Bits,
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
    /// of the changes (KERNEL.42): a formula's twin cell by its formula's
    /// position, a bit each (KERNEL.49's decision 6), and a value cell's,
    /// one written since, few, by its place; where a list held room for half
    /// the cells.
    twin_changed: Bits,
    twin_written: Vec<CellId>,
    /// How many twin cells the last frame's end changed, and how many it may
    /// change, half the cells, past which the next frame computes every cell
    /// rather than follow each (KERNEL.20), as when the list's room ran out.
    twin_count: usize,
    room: usize,
    /// Those of them whose change a record spells as before, with the
    /// records that do ([`SHEET_ALIKE`], [`TWIN_ALIKE`]), at most
    /// [`ALIKE`]; and whether there were more (KERNEL.42).
    alike: Vec<(CellId, u8)>,
    alike_over: bool,
    /// The same of the frame in progress, kept as it changes each formula's
    /// value and made the frame's at its end (KERNEL.49).
    alike_next: Vec<(CellId, u8)>,
    alike_next_over: bool,
    /// Whether the last complete frame's changes are more than the lists
    /// tell (KERNEL.42): before the first frame, or when its end kept more
    /// changes than it held room for, or more spelled alike than [`ALIKE`],
    /// or made a twin's extent again or grew it, or a write before it changed
    /// a formula or moved a sheet's extent, which may change a record's rows
    /// or its head where no twin cell changes (KERNEL.56).
    changes_whole: bool,
    /// Whether a write since the last frame's end changed a formula or
    /// moved a sheet's extent (KERNEL.42; KERNEL.56).
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
    /// Whether a frame's end copies the cells it changed into the twins'
    /// stores and the planes where it may: always, but in the tests that hold
    /// the copies to each made again at every frame's end (KERNEL.36).
    copies: bool,
    /// Whether every twin keeps a store, read or not: never, but in the tests
    /// that hold a twin made from its sheet to the same twin kept (KERNEL.49).
    every_twin: bool,
    /// Whether a rows' sheet's store is offered to the evaluator as an array
    /// (KERNEL.31; KERNEL.49): always, but in the tests that hold the arrays
    /// to the same sheets read cell by cell.
    arrays: bool,
    /// The vectors the lookups and the criteria search, each indexed at a
    /// second probe in a frame and kept until a cell of it changes, which
    /// every change of a value tells them (KERNEL.8).
    indexes: Indexes,
    /// Whether the lookups and the criteria search through those indexes:
    /// always, but in the tests that hold the indexes to the scan.
    indexed: bool,
    /// The answer of the call at a cell's program's root that is no number,
    /// handed over as the program gives the cell up (KERNEL.8).
    root_answer: Slot<Option<Value>>,
    /// The cells the frame in progress computed, for the tests to count.
    #[cfg(test)]
    computed: usize,
    /// The cells programs answered since the machine was made, for the tests
    /// to count (KERNEL.44).
    #[cfg(test)]
    ran: usize,
    /// How often a write that planned again or added a cell kept a rows'
    /// sheet's store over its extent, and how often one moved a store to
    /// another extent, taken over it where it stands or re-laid, for the
    /// tests to count (KERNEL.31; KERNEL.49; KERNEL.56).
    #[cfg(test)]
    stores_kept: usize,
    #[cfg(test)]
    stores_moved: usize,
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

/// What one write changed, to settle or to undo (KERNEL.49's third slice):
/// each row's formulas set in the grid as it lands, its values kept here
/// until the settle writes them into the stores, which no row reads before
/// it; and the formulas as the write found them, which an undo puts back.
#[derive(Default)]
struct Log {
    /// The cells whose content changed, each with the line of the row that
    /// changed it first.
    changed: BTreeMap<CellId, u32>,
    /// What each changed value cell now holds, which the settle writes into
    /// its sheet's store; a changed cell missing here holds a formula now.
    contents: BTreeMap<CellId, Content>,
    /// The cells of a fill a release gave their own text, whose formulas
    /// read as before: placed again when the write settles, their values
    /// kept (KERNEL.34).
    released: BTreeSet<CellId>,
    written: usize,
    added: usize,
    /// The cells added on each sheet, which its count takes at the settle.
    added_on: BTreeMap<usize, usize>,
    /// Each sheet the write set a cell of, with its extent as the write
    /// found it, which an undo puts back (KERNEL.54): kept beside the sheet,
    /// where a list of its own cost a write an allocation more, and a host
    /// writing its devices every frame a frame's.
    sheets: BTreeMap<usize, Extent>,
    /// Each sheet's fills and its count of shared formulas as the write
    /// found them, kept at the first change of either.
    fills: BTreeMap<usize, (Fills, u32)>,
    /// Each formula alone the write set or gave up, with what its cell held
    /// as one before, in order.
    alone: Vec<(CellId, Option<(String, Kind)>)>,
}

impl Machine {
    /// A text of rows loaded under a library: [`read_rows`], then
    /// [`Machine::new`].
    pub fn load(text: &str, library: Library) -> Result<Machine, LineRefusal> {
        let forms = read_forms(text).map_err(|refusal| LineRefusal { line: 0, refusal })?;
        let draft = read_rows(&forms)?;
        Machine::new(draft, library)
    }

    /// The machine over a draft: its grid made from the draft's records, no
    /// cell of a range made (KERNEL.49's decision 8), and the twins; every
    /// formula placed on its shape, its references' sheets found among the
    /// grid's and its twins (KERNEL.20), then refused when the grid cannot be
    /// stepped (a cycle, a function the library does not hold, a construct
    /// this version does not read, in that order); then each sheet's values
    /// put in its store and the twins filled from them, each formula's from
    /// its `value` row. Nothing is computed: the frame is 0 and every formula
    /// waits for the first step.
    pub fn new(mut draft: Draft, library: Library) -> Result<Machine, LineRefusal> {
        let values = std::mem::take(&mut draft.values);
        let cells = draft.cells;
        let mut extents: Vec<Extent> = draft.sheets.iter().map(|s| s.extent()).collect();
        let mut book = Book::of(&mut draft);
        // Each record's rectangle, line and value, moved out of the draft,
        // which is given up before the plan, the load's peak; the stores made
        // after it.
        let laid: Vec<Vec<Laid>> = draft
            .sheets
            .into_iter()
            .map(|s| s.records.into_iter().flatten().map(Laid::of).collect())
            .collect();
        let line = |&(s, row, col): &CellId| {
            laid.get(s)?.iter().find_map(|l| {
                let r = l.range;
                (row >= r.top && row <= r.bottom && col >= r.left && col <= r.right)
                    .then_some(l.line)
            })
        };
        let sheets = book.pages.len();
        let names = Names::of_model(&book);
        for i in 0..sheets {
            let mut twin = Sheet::new(&format!("{}.last", book.pages[i].sheet.name));
            twin.visibility = Visibility::Hidden;
            book.pages.push(Page::new(twin));
            extents.push(None);
        }
        let shapes = Shapes::of_model(&book, sheets, &library, &names, &extents);
        if let Some((cell, refusal)) = shapes.first_refusal_in(&book, None) {
            let line = if refusal.id == "calc-cycle" {
                shapes
                    .cycles()
                    .first()
                    .and_then(|c| c.iter().filter_map(line).min())
                    .unwrap_or(0)
            } else {
                line(&cell).unwrap_or(0)
            };
            return Err(LineRefusal { line, refusal });
        }
        let stores = stores_of(&laid, &book, &extents, &values);
        drop(laid);
        drop(values);
        Ok(Machine::made(
            (book, sheets, cells),
            (library, names, extents),
            shapes,
            stores,
        ))
    }

    /// The machine made from its parts, in a frame of its own (KERNEL.20):
    /// the machine is the largest value its load holds, and made in the frame
    /// that parses every formula it would stand on the stack beneath the
    /// deepest parse, which `VLA.1`'s limits pin on a stack of 1 MiB.
    #[inline(never)]
    fn made(
        (book, sheets, cells): (Book, usize, usize),
        (library, names, extents): (Library, Names, Vec<Extent>),
        shapes: Shapes,
        stores: Vec<Store>,
    ) -> Machine {
        let mut m = Machine {
            book,
            sheets,
            library,
            names,
            extents,
            shapes,
            stores,
            before: FxHashMap::default(),
            written_before: FxHashMap::with_capacity_and_hasher(WRITTEN, Default::default()),
            frame: 0,
            progress: None,
            cells,
            here: Slot::new((0, 1, 1)),
            wanted: (0..sheets).map(|_| Slot::new(false)).collect(),
            planes: vec![None; sheets],
            twins: (0..sheets).map(|_| None).collect(),
            lost: vec![false; sheets],
            gained: vec![None; sheets],
            copied_into: vec![0; sheets],
            extent_moved: Vec::new(),
            lanes: Lanes::default(),
            answers: Vec::new(),
            answered: Vec::with_capacity(vector::CHUNK),
            gathered: Vec::with_capacity(vector::CHUNK),
            gathered_at: Vec::with_capacity(vector::CHUNK),
            locator: Locator::default(),
            every_next: true,
            every: true,
            marked: Bits::default(),
            whole: Vec::new(),
            recheck: Vec::new(),
            wholes: Vec::new(),
            rechecked: Vec::new(),
            parts: Vec::new(),
            changed: Bits::default(),
            last_changed: 0,
            changed_sheets: vec![0; sheets],
            last_reached: 0,
            reaching: 0,
            written: Vec::with_capacity(WRITTEN),
            twin_changed: Bits::default(),
            twin_written: Vec::with_capacity(WRITTEN),
            twin_count: 0,
            room: 0,
            alike: Vec::with_capacity(ALIKE),
            alike_over: false,
            alike_next: Vec::with_capacity(ALIKE),
            alike_next_over: false,
            changes_whole: true,
            reshaped: false,
            by_runs: true,
            lowered: true,
            incremental: true,
            copies: true,
            every_twin: false,
            arrays: true,
            indexes: Indexes::new(),
            indexed: true,
            root_answer: Slot::new(None),
            #[cfg(test)]
            computed: 0,
            #[cfg(test)]
            ran: 0,
            #[cfg(test)]
            stores_kept: 0,
            #[cfg(test)]
            stores_moved: 0,
        };
        m.fit();
        // Each twin as the load leaves it, each formula's from its `value`
        // row (KERNEL.49): its extent, and its store where a formula reads it.
        for s in 0..sheets {
            m.extents[sheets + s] = m.twin_extent(s);
        }
        m.keep_twins();
        m
    }

    /// What follows a frame's changes made to the plan's size, once a plan
    /// is made, so that a frame makes none (KERNEL.20): a bit a cell of the
    /// order for the cells a frame computes, for those it changes and for
    /// those whose twin its end changes (KERNEL.49's decision 6), a flag a
    /// shape, a value a part; room in the lists of what a frame sets, each
    /// entry set at most once a frame (KERNEL.54), every shape made whole and
    /// every shape with parts rechecked; and the count of the twin cells'
    /// changes, half the cells, past which the next frame computes every
    /// cell. A plan made again follows no change of the frame before: the
    /// next frame computes every cell.
    fn fit(&mut self) {
        let of = self.shapes.len();
        self.marked.sized(of);
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
        self.changed.sized(of);
        self.twin_changed.sized(of);
        self.twin_written.clear();
        self.twin_count = 0;
        self.room = self.cells / 2 + 64;
        let reach = self.shapes.reach();
        let parted = (0..shapes).filter(|&s| !reach.parts(s).is_empty()).count();
        self.wholes.clear();
        room(&mut self.wholes, shapes);
        self.rechecked.clear();
        room(&mut self.rechecked, parted);
    }

    /// What the load made, for the load's row.
    pub fn loaded(&self) -> Loaded {
        Loaded {
            sheets: self.sheets,
            cells: self.cells,
            formulas: self.shapes.formulas(),
        }
    }

    /// The grid as a workbook, its twins included (KERNEL.49): each rows'
    /// sheet's cells made from its grid, a fill a shared formula from its
    /// anchor naming the range it was written over, a formula alone its own,
    /// a value cell's content as written; each twin's made from its sheet as
    /// the last frame's end left it, a value cell's content and a formula's
    /// value as a value cell holds it, in the default format. The machine
    /// keeps no cell, so the grid is lent as a copy, made when asked, as a
    /// door and a test ask. A sheet's name and place are
    /// [`Machine::find_sheet`]'s and [`Machine::sheet_name`]'s, which need no
    /// copy.
    pub fn workbook(&self) -> Workbook {
        let mut wb = Workbook::new();
        wb.styles = self.book.styles.clone();
        wb.defined_names = self.book.defined_names.clone();
        let all = A1Range {
            top: 1,
            left: 1,
            bottom: MAX_ROW,
            right: MAX_COLUMN,
        };
        for (s, page) in self.book.pages.iter().enumerate() {
            let mut sheet = page.sheet.clone();
            if s < self.sheets {
                let store = &self.stores[s];
                let mut formulas = grid::Beside::new(page.formulas_iter(&all));
                let mut styles = grid::Beside::new(grid::in_window(&page.styles, &all));
                // The cells in their order, the sheet's tree made from them at
                // once rather than a cell at a time.
                let mut cells: Vec<((u32, u32), Cell)> = Vec::with_capacity(page.cells);
                for (row, col) in store.walk(&all) {
                    let content = match formulas.at((row, col)) {
                        Some(Formula::Own(text, Kind::Plain)) => Content::Formula(text.to_string()),
                        Some(Formula::Own(text, Kind::Dynamic)) => {
                            Content::DynamicFormula(text.to_string())
                        }
                        Some(Formula::Own(text, Kind::Array)) => {
                            Content::ArrayFormula(text.to_string())
                        }
                        Some(Formula::Fill(si)) => match page.fills.fill(si) {
                            Some(f) if f.anchor == (row, col) => Content::SharedMaster {
                                text: f.text.clone(),
                                range: f.written.text(),
                                si,
                            },
                            _ => Content::SharedChild { si },
                        },
                        None => match store.content(row, col) {
                            Some(content) => content,
                            None => continue,
                        },
                    };
                    let style = styles.at((row, col)).unwrap_or(0);
                    cells.push(((row, col), Cell { content, style }));
                }
                sheet.cells = cells.into_iter().collect();
            } else {
                sheet.cells = self.twin_cells(s - self.sheets, &all).into_iter().collect();
            }
            wb.sheets.push(sheet);
        }
        wb
    }

    /// A sheet's index by name, compared as Excel compares names, a twin's
    /// included, as [`Workbook::find_sheet`] finds it.
    pub fn find_sheet(&self, name: &str) -> Option<usize> {
        self.book.find_sheet(name)
    }

    /// A sheet's name by index, a twin's included.
    pub fn sheet_name(&self, index: usize) -> Option<&str> {
        self.book.pages.get(index).map(|p| p.sheet.name.as_str())
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
        if sheet >= self.book.pages.len() {
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
    /// rows' sheet as its store keeps it (KERNEL.49), an array where its
    /// cells stand close enough together, as a door's census names it, a
    /// reason being a slot's kind and never a sheet's. Computed when asked:
    /// a load, a step and a write keep nothing for it but where each level
    /// of the order ends.
    pub fn census(&self) -> String {
        crate::calc::census::rows(&self.shapes, &self.book, &self.library, &|s| {
            if s < self.sheets {
                self.arrays && self.stores[s].array().is_some()
            } else {
                self.twin_store(s).is_some_and(|t| t.array().is_some())
            }
        })
    }

    /// A twin's store, where a formula reads the twin (KERNEL.49); `None`
    /// for a rows' sheet and a twin no formula reads.
    fn twin_store(&self, sheet: usize) -> Option<&Store> {
        self.twins.get(sheet.checked_sub(self.sheets)?)?.as_ref()
    }

    fn during_step(&self, done: usize) -> Refusal {
        raise(
            "grid-write-during-step",
            &[
                ("frame", &thousands(self.frame + 1)),
                ("evaluated", &thousands(done as u64)),
                ("of", &thousands(self.shapes.len() as u64)),
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
        let of = self.shapes.len();
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
        if self.every {
            self.walk_every(start, end, keep);
        } else {
            self.walk_marks(start, end, keep);
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

    /// The positions from `start` to `end` of a frame that computes every
    /// cell, every one computed: the runs walked in order beside the
    /// positions, the next found once where a search found the run of each
    /// position (KERNEL.44), a grid of many cells alone between its runs
    /// having searched them at every one.
    fn walk_every(&mut self, start: usize, end: usize, keep: bool) {
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
                    #[cfg(test)]
                    walk::run();
                    self.evaluate_run(shape, i, to, keep);
                    i = to;
                    k += 1;
                }
                _ => {
                    let stop = next.map_or(end, |(from, _, _)| from.min(end));
                    #[cfg(test)]
                    walk::visited((stop - i) as u64);
                    while i < stop {
                        self.evaluate_alone(i, keep);
                        i += 1;
                    }
                }
            }
        }
    }

    /// The positions from `start` to `end` of a frame that computes what
    /// changed, passed from one marked position to the next (KERNEL.54): a
    /// run's cells gathered from the first marked among them, and the cells
    /// alone up to the next run each at its mark, the next run found once
    /// for them all; a shape made whole having marked every cell of it
    /// through its stretches. A cell a computed cell's change reaches stands
    /// later in the order and outside its run, so the search finds its mark.
    /// The walk visited every run and asked every cell alone whether it was
    /// marked or its shape whole: an idle frame of 391,168 cells alone cost
    /// 430 µs, and costs 1.3, the search a word for 64 positions.
    fn walk_marks(&mut self, start: usize, end: usize, keep: bool) {
        let mut i = start;
        while let Some(p) = self.marked.next(i, end) {
            let run = if self.by_runs {
                self.shapes.run_span(self.shapes.run_after(p))
            } else {
                None
            };
            match run {
                Some((from, after, shape)) if from <= p => {
                    let to = after.min(end);
                    #[cfg(test)]
                    walk::run();
                    self.evaluate_run(shape, p, to, keep);
                    i = to;
                }
                next => {
                    let stop = next.map_or(end, |(from, _, _)| from.min(end));
                    let mut q = p;
                    while let Some(m) = self.marked.next(q, stop) {
                        #[cfg(test)]
                        walk::visited(1);
                        self.evaluate_alone(m, keep);
                        q = m + 1;
                    }
                    i = stop;
                }
            }
        }
    }

    /// A frame's start (KERNEL.20): every cell to compute, or the cells the
    /// last frame's end and the writes since reach, with every shape not
    /// followed, and each shape whose parts the same at every cell read a
    /// changed cell when those parts compute other bits than they did. The
    /// last frame's changes are kept until this frame's end, which a view
    /// of them during a yield reads (KERNEL.42).
    fn begin(&mut self) {
        // A vector probed once in the frame before is probed once again here
        // before its index is built (KERNEL.8).
        self.indexes.begin();
        // The order may have been planned again since the last frame.
        self.locator = Locator::default();
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
            self.extent_moved.clear();
            return;
        }
        let Machine {
            shapes,
            sheets,
            marked,
            whole,
            recheck,
            wholes,
            rechecked,
            twin_changed,
            twin_written,
            written,
            extent_moved,
            ..
        } = self;
        let reach = shapes.reach();
        for &s in reach.unfollowed() {
            make_whole(shapes, whole, wholes, marked, s);
        }
        // A sheet whose extent moved, a rows' sheet's at a write and a twin's
        // at the last frame's end: every shape reading its extent itself
        // computed whole (KERNEL.56), the extent moved where no cell it reads
        // changed.
        for &x in extent_moved.iter() {
            for s in reach.extent_readers(x) {
                make_whole(shapes, whole, wholes, marked, s);
            }
        }
        extent_moved.clear();
        // The twin cells the last frame's end changed, a formula's by its
        // position in the order, which no write since moved (a write that
        // plans again makes this frame compute every cell); and the cells
        // written since.
        let blocks = shapes.blocks();
        let mut at = Locator::default();
        let twins = twin_changed.iter().map(|p| {
            let ((s, row, col), _) = at.cell(blocks, p);
            (*sheets + s, row, col)
        });
        let others = twin_written.iter().chain(written.iter()).copied();
        for id in twins.chain(others) {
            shapes.reached(id, &mut |r| match r {
                Reached::Cell(p) => marked.set(p),
                Reached::Shape(s) => make_whole(shapes, whole, wholes, marked, s),
                Reached::Parts(s) => {
                    if !recheck[s] {
                        recheck[s] = true;
                        rechecked.push(s as u32);
                    }
                }
            });
        }
        // Every shape whose parts read a changed cell has them computed, even
        // one computed whole already, so that what is kept is what its cells
        // were computed with: a part kept from an earlier frame would let a
        // later change back to it pass unseen. The shapes on the list alone,
        // where every shape was asked (KERNEL.54).
        #[cfg(test)]
        walk::shapes(self.rechecked.len() as u64);
        for k in 0..self.rechecked.len() {
            let s = self.rechecked[k] as usize;
            if self.parts_moved(s) {
                let Machine {
                    shapes,
                    marked,
                    whole,
                    wholes,
                    ..
                } = self;
                make_whole(shapes, whole, wholes, marked, s);
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

    fn evaluate(&self, id: CellId, shape: usize) -> Computed {
        self.here.set(id);
        self.shapes.evaluate_shaped(self, &self.library, shape, id)
    }

    /// The cell at a position of the order evaluated alone, and its value
    /// kept: by its shape's program, at the cell's offset from the shape's
    /// first cell, where the shape has one and no guard fails (KERNEL.44);
    /// else by `Ctx`, as every cell alone was before.
    fn evaluate_alone(&mut self, p: usize, keep: bool) {
        let (id, s) = self.locator.cell(self.shapes.blocks(), p);
        self.evaluate_cell(p, id, s, keep);
    }

    /// The cell at a position of the order, its place and its shape known,
    /// evaluated alone and its value kept (`evaluate_alone`).
    fn evaluate_cell(&mut self, p: usize, id: CellId, s: usize, keep: bool) {
        self.reaching += 1;
        if self.lowered {
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
                        Answer::Number(x) => self.keep_number(p, id, x, keep),
                        Answer::Truth(b) => self.keep(p, id, Computed::Value(Value::Bool(b)), keep),
                    };
                }
                // The call at the root answered no number: its answer the
                // cell's all the same (KERNEL.8).
                if let Some(v) = self.root_answer.take() {
                    #[cfg(test)]
                    {
                        self.ran += 1;
                    }
                    return self.keep(p, id, Computed::Value(v), keep);
                }
            }
        }
        let computed = self.evaluate(id, s);
        self.keep(p, id, computed, keep);
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
                gathered.clear();
                self.shapes.blocks().extend_cells(i, j, &mut gathered);
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
                    for (k, (&x, &id)) in answered.iter().zip(&gathered).enumerate() {
                        self.keep_number(i + k, id, x, keep);
                    }
                    answered.clear();
                } else {
                    self.reaching += answers.len();
                    for (k, (computed, &id)) in answers.drain(..).zip(&gathered).enumerate() {
                        self.keep(i + k, id, computed, keep);
                    }
                }
                i = j;
                continue;
            }
            gathered.clear();
            gathered_at.clear();
            let blocks = self.shapes.blocks();
            let mut finder = self.locator;
            while gathered.len() < vector::CHUNK {
                match self.marked.next(i, to) {
                    Some(p) => {
                        gathered.push(finder.cell(blocks, p).0);
                        gathered_at.push(p as u32);
                        i = p + 1;
                    }
                    None => {
                        i = to;
                        break;
                    }
                }
            }
            self.locator = finder;
            if gathered.len() < few {
                for (&p, &id) in gathered_at.iter().zip(&gathered) {
                    self.evaluate_cell(p as usize, id, shape, keep);
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
                for ((&p, &id), &x) in gathered_at.iter().zip(&gathered).zip(&answered) {
                    self.keep_number(p as usize, id, x, keep);
                }
                answered.clear();
            } else {
                self.reaching += answers.len();
                let cells = gathered_at.iter().zip(&gathered);
                for ((&p, &id), computed) in cells.zip(answers.drain(..)) {
                    self.keep(p as usize, id, computed, keep);
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
    fn keep_number(&mut self, p: usize, id: CellId, x: f64, keep: bool) {
        #[cfg(test)]
        {
            self.computed += 1;
        }
        let (s, row, col) = id;
        if let Some(Stored::Number(held)) = self.stores[s].slot_of(row, col) {
            if held.to_bits() == x.to_bits() {
                return;
            }
        }
        self.changed_to(p, id, Computed::Value(Value::Number(x)), keep);
    }

    /// The value the frame computed for the cell at a position of the order,
    /// kept when its bits differ from the one its slot holds (KERNEL.20).
    fn keep(&mut self, p: usize, id: CellId, computed: Computed, keep: bool) {
        #[cfg(test)]
        {
            self.computed += 1;
        }
        let (s, row, col) = id;
        if self.stores[s]
            .slot_of(row, col)
            .is_some_and(|held| same_held(held, &computed))
        {
            return;
        }
        self.changed_to(p, id, computed, keep);
    }

    /// A value that differs from the one its cell holds kept in its slot
    /// (KERNEL.49): the old one beside while a step yields, the cell counted
    /// as changed, and, but in a frame that computes every cell, the cells of
    /// the frame that read it marked. The twin's change is told before the
    /// slot is overwritten, the slot holding the twin's value as the last
    /// frame's end left it: a twin cell lost, which makes the twin's extent
    /// again at the frame's end, or gained, which grows it by the cell there
    /// (KERNEL.56), or the records that spell the change as before
    /// (KERNEL.42). A formula a write changed since, its slot holding no
    /// value, is told at the frame's end, against its twin as the write found
    /// it.
    fn changed_to(&mut self, p: usize, id: CellId, computed: Computed, keep: bool) {
        debug_assert!(
            !matches!(computed, Computed::Value(Value::Empty)),
            "a formula's result is settled"
        );
        debug_assert_eq!(self.shapes.at(p).0, id, "the cell at its position");
        // An index whose vector holds the cell is built again at its next
        // probe (KERNEL.8); nothing asked where none is built.
        self.indexes.changed(id);
        let was = self.stores[id.0].slot_of(id.1, id.2);
        if keep {
            if let std::collections::hash_map::Entry::Vacant(e) = self.before.entry(p as u32) {
                e.insert(was.and_then(Stored::computed));
            }
        }
        let written = matches!(was, Some(Stored::Pending))
            && !self.written_before.is_empty()
            && self.written_before.contains_key(&id);
        if !written {
            let cell = matches!(computed, Computed::Value(_));
            if holds_cell(was) && !cell {
                self.lost[id.0] = true;
            } else if !holds_cell(was) && cell {
                self.gained[id.0] = grown(self.gained[id.0], id.1, id.2);
            } else {
                let twin = (self.sheets + id.0, id.1, id.2);
                let bits = changed_alike(was, &computed);
                note_alike(&mut self.alike_next, &mut self.alike_next_over, twin, bits);
            }
        }
        self.stores[id.0].put_computed(id.1, id.2, Some(&computed));
        self.changed.set(p);
        if !self.every {
            let Machine {
                shapes,
                marked,
                whole,
                wholes,
                ..
            } = self;
            shapes.reached(id, &mut |r| match r {
                Reached::Cell(q) => marked.set(q),
                Reached::Shape(s) | Reached::Parts(s) => {
                    make_whole(shapes, whole, wholes, marked, s);
                }
            });
        }
    }

    /// A frame's end (KERNEL.20): the cells it changed and the cells written
    /// since the last told against their twins and copied into them, the
    /// twins' stores and the planes kept (KERNEL.49, [`Machine::end_twins`]);
    /// the twin cells that changed kept for the next frame, and the frame's
    /// changes unless they are more than the lists tell (KERNEL.42); and what
    /// the frame changed and computed kept for the counts until the next
    /// frame's end (KERNEL.48), each sheet's share counted as each changed
    /// cell is copied.
    fn snapshot(&mut self) {
        // What the frame set cleared by its lists, and its marks by their
        // summary (KERNEL.54; KERNEL.49), where every shape's two flags and
        // every word of the marks were.
        #[cfg(test)]
        walk::cleared((self.wholes.len() + self.rechecked.len()) as u64);
        for &s in &self.wholes {
            self.whole[s as usize] = false;
        }
        self.wholes.clear();
        for &s in &self.rechecked {
            self.recheck[s as usize] = false;
        }
        self.rechecked.clear();
        let _words = self.marked.clear();
        #[cfg(test)]
        walk::cleared(_words as u64);
        self.twin_changed.clear();
        self.twin_written.clear();
        self.twin_count = 0;
        // The changes the frame spelled as before, the frame's now (KERNEL.42).
        std::mem::swap(&mut self.alike, &mut self.alike_next);
        self.alike_next.clear();
        self.alike_over = std::mem::take(&mut self.alike_next_over);
        self.changed_sheets.fill(0);
        let grew = self.end_twins();
        // Each sheet's share counted as its changed cells were copied.
        self.last_changed = self.changed_sheets.iter().map(|&n| n as usize).sum();
        self.last_reached = self.reaching;
        // A twin's indexes (KERNEL.8): each whose vector holds a twin cell
        // the end changed built again at its next probe, or every twin's where
        // the end made a twin's extent again or changed more cells than it
        // kept; the cells not walked where no index is built.
        if self.indexes.any() && self.every_next {
            for t in self.sheets..self.book.pages.len() {
                self.indexes.sheet_changed(t);
            }
        } else if self.indexes.any() {
            let Machine {
                shapes,
                sheets,
                indexes,
                twin_changed,
                twin_written,
                ..
            } = self;
            let mut at = Locator::default();
            for p in twin_changed.iter() {
                let ((s, row, col), _) = at.cell(shapes.blocks(), p);
                indexes.changed((*sheets + s, row, col));
            }
            for &id in twin_written.iter() {
                indexes.changed(id);
            }
        }
        self.changed.clear();
        self.written.clear();
        // The frame began with the next frame's flag clear, and no write
        // lands during a frame, so the flag is the end's own. A twin's extent
        // the end grew is named in every window's record of the twin, which
        // no cell's change carries (KERNEL.56).
        self.changes_whole = self.every_next || self.alike_over || self.reshaped || grew;
        self.reshaped = false;
    }

    /// The twins at a frame's end (KERNEL.49): each formula cell the frame
    /// changed and each cell written since the last end told against its
    /// twin as that end left it, and copied into its twin's store where one
    /// is kept and into its sheet's plane where one is kept, each changed
    /// formula counted on its sheet as it goes (KERNEL.48), its place read
    /// here already. A formula the frame changed with no write since changed
    /// its twin cell, its bits differing, its gain or loss and its spelling
    /// told as the frame changed it; a cell written since is told here,
    /// against its twin as the write found it. A twin cell lost makes its
    /// twin's extent again and the next frame compute every cell (KERNEL.20),
    /// as more changes than the room kept for them do; the cells a twin
    /// gained grow its extent by their rectangle, the next frame following
    /// them as any change, and a twin whose extent grew is named for the
    /// next frame to compute whole the shapes that read its extent itself
    /// (KERNEL.56); whether one grew, for the changes. A twin's array whose
    /// sheet's array a write grew over an extent holding its own is taken
    /// over it the same way, its slots moved and rows added in place
    /// (KERNEL.56); a twin's store whose sheet's store moved otherwise under
    /// it, and a plane a write gave up, are made again from the sheet, as
    /// each is at every end when the tests hold the copies to that.
    fn end_twins(&mut self) -> bool {
        for s in 0..self.sheets {
            let twin = self.copies
                && match self.twins[s].as_mut() {
                    Some(t) => t.follows(&self.stores[s]),
                    None => false,
                };
            let plane = self.copies && self.wanted[s].get() && self.planes[s].is_some();
            self.copied_into[s] = u8::from(twin) | u8::from(plane) << 1;
        }
        let Machine {
            book,
            sheets,
            shapes,
            stores,
            twins,
            written_before,
            planes,
            changed,
            changed_sheets,
            written,
            twin_changed,
            twin_written,
            twin_count,
            room,
            alike,
            alike_over,
            every_next,
            lost,
            gained,
            copied_into,
            ..
        } = self;
        let sheets = *sheets;
        let mut at = Locator::default();
        for p in changed.iter() {
            #[cfg(test)]
            crate::sheet::work::copied();
            let ((s, row, col), _) = at.cell(shapes.blocks(), p);
            changed_sheets[s] += 1;
            let slot = stores[s].slot_of(row, col);
            let t = (sheets + s, row, col);
            let was = match written_before.is_empty() {
                true => None,
                false => written_before.remove(&(s, row, col)),
            };
            match was {
                Some(was) => {
                    let (moved, bits, grew) = told(true, &was, &formula_held(slot));
                    if moved {
                        note(twin_count, *room, every_next);
                        twin_changed.set(p);
                        note_alike(alike, alike_over, t, bits);
                    }
                    grew.tell(&mut lost[s], &mut gained[s], row, col);
                }
                None => {
                    note(twin_count, *room, every_next);
                    twin_changed.set(p);
                }
            }
            if copied_into[s] & 1 != 0 {
                if let Some(twin) = twins[s].as_mut() {
                    twin.put_slot(row, col, twin_slot(slot));
                }
            }
            if copied_into[s] & 2 != 0 {
                if let Some(r) = planes[s].as_mut() {
                    r.set(row, col, plane::slot_byte(slot.unwrap_or(Stored::Pending)));
                }
            }
        }
        for &id in written.iter() {
            if written_before.is_empty() {
                break;
            }
            // A cell written twice since is told once.
            let Some(was) = written_before.remove(&id) else {
                continue;
            };
            #[cfg(test)]
            crate::sheet::work::copied();
            let (s, row, col) = id;
            // A write empties no cell: the store holds a slot at every cell
            // of its sheet (KERNEL.49).
            let Some(slot) = stores[s].slot_of(row, col).filter(|x| !x.is_empty()) else {
                continue;
            };
            let formula = book.pages[s].formula(row, col).is_some();
            let t = (sheets + s, row, col);
            // A value cell's content told where it stands, uncopied.
            let constant = (!formula).then(|| stores[s].constant(row, col)).flatten();
            let (moved, bits, grew) = match &constant {
                None => told(true, &was, &formula_held(Some(slot))),
                Some(now) => told_cell(&was, now),
            };
            if moved {
                note(twin_count, *room, every_next);
                twin_written.push(t);
                note_alike(alike, alike_over, t, bits);
            }
            grew.tell(&mut lost[s], &mut gained[s], row, col);
            if copied_into[s] & 1 != 0 {
                if let Some(twin) = twins[s].as_mut() {
                    match &constant {
                        None => twin.put_slot(row, col, twin_slot(Some(slot))),
                        Some(now) => twin.put_constant(row, col, now),
                    }
                }
            }
            if copied_into[s] & 2 != 0 {
                if let Some(r) = planes[s].as_mut() {
                    // A value cell's slot draws the byte its content does.
                    r.set(row, col, plane::slot_byte(slot));
                }
            }
        }
        // Every changed cell was in one list or the other; the room kept
        // for a host's writes of its devices, and no more after a large write.
        written_before.clear();
        if written_before.capacity() > 4 * WRITTEN {
            written_before.shrink_to(WRITTEN);
        }
        for s in 0..sheets {
            if copied_into[s] & 1 == 0 {
                if let Some(twin) = twins[s].as_mut() {
                    #[cfg(test)]
                    crate::sheet::work::swept(book.pages[s].cells);
                    *twin = stores[s].twin();
                }
            }
        }
        let mut grew = false;
        for s in 0..sheets {
            if self.wanted[s].get() && self.copied_into[s] & 2 == 0 {
                self.plane_made(s);
            }
            let gained = self.gained[s].take();
            if std::mem::take(&mut self.lost[s]) {
                self.every_next = true;
                self.extents[sheets + s] = self.twin_extent(s);
            } else if let Some(((top, left), (bottom, right))) = gained {
                let t = sheets + s;
                let was = self.extents[t];
                let now = grown(grown(was, top, left), bottom, right);
                if now != was {
                    self.extents[t] = now;
                    self.note_extent_moved(t);
                    grew = true;
                }
            }
        }
        grew
    }

    /// A sheet whose extent moved named for the next frame's start, once
    /// (KERNEL.56).
    fn note_extent_moved(&mut self, sheet: usize) {
        if !self.extent_moved.contains(&sheet) {
            self.extent_moved.push(sheet);
        }
    }

    /// A sheet's plane made again over its extent from its store (KERNEL.24;
    /// KERNEL.49), where a write gave it up, or every frame's end in the
    /// tests that hold the copies to it: a byte a cell, a value cell's slot
    /// drawing the byte its content does. None for a sheet past a plane's
    /// limit, or with a cell outside its extent, which would mean a stale
    /// extent: the walk answers rather than a plane that might be wrong.
    fn plane_made(&mut self, s: usize) {
        #[cfg(test)]
        crate::sheet::work::swept(self.book.pages[s].cells);
        let old = self.planes[s].take();
        let Some(mut raster) = plane::Raster::over(self.extents[s], CELLS as u64, old) else {
            return;
        };
        let mut fits = true;
        self.stores[s].each_held(&EVERY, &mut |row, col, slot| {
            fits &= raster.set(row, col, plane::slot_byte(slot));
        });
        self.planes[s] = fits.then_some(raster);
    }

    /// A twin's extent at a load or a frame's end (KERNEL.49): the rectangle
    /// of its cells, its sheet's that hold a value there, a value cell or a
    /// formula with one, as the record names it and a copy of the grid lent
    /// walks it; a twin holding no cell for a reason or for a formula with no
    /// value.
    fn twin_extent(&self, s: usize) -> Extent {
        debug_assert!(self.written_before.is_empty() && self.before.is_empty());
        #[cfg(test)]
        crate::sheet::work::swept(self.book.pages[s].cells);
        let mut extent: Extent = None;
        self.stores[s].each_held(&EVERY, &mut |row, col, slot| {
            if holds_cell(Some(slot)) {
                extent = grown(extent, row, col);
            }
        });
        extent
    }

    /// Each twin's store kept where a formula reads the twin (KERNEL.49's
    /// decision 5), as the plan says: a twin a shape reads, or every twin
    /// while some shape is not followed and may read any; made from its sheet
    /// as the last frame's end left it where the plan first reads it, and
    /// given up where the plan no longer does.
    fn keep_twins(&mut self) {
        let any = self.every_twin || !self.shapes.reach().unfollowed().is_empty();
        for s in 0..self.sheets {
            let read = any || self.shapes.reach().readers_on(self.sheets + s) > 0;
            match (read, self.twins[s].is_some()) {
                (true, false) => self.twins[s] = Some(self.twin_made(s)),
                (false, true) => self.twins[s] = None,
                _ => {}
            }
        }
    }

    /// A twin's store made from its sheet as the last frame's end left it
    /// (KERNEL.49): its sheet's store's twin while no write since stands,
    /// and otherwise each cell's slot what its twin holds, over the sheet's
    /// extent by the rule its sheet's store keeps.
    fn twin_made(&self, s: usize) -> Store {
        if self.written_before.is_empty() && self.before.is_empty() {
            return self.stores[s].twin();
        }
        let limit = dense::limit(self.book.pages[s].cells, CELLS as u64);
        let mut places: Vec<(u32, u32)> = Vec::new();
        self.stores[s].each_held(&EVERY, &mut |row, col, _| places.push((row, col)));
        let mut store = Store::over(self.extents[s], limit, places.iter().copied());
        for &(row, col) in &places {
            match self.held(s, row, col) {
                Held::Cell(content) => store.put_content(row, col, Some(&content)),
                Held::Reason(r) => {
                    store.put_computed(row, col, Some(&Computed::NotComputed(r)));
                }
                Held::Nothing => {}
            }
        }
        store
    }

    /// What a rows' cell's twin holds as the last frame's end left it
    /// (KERNEL.49): a cell a write changed since, as the write found it; a
    /// formula's value of the last complete frame, one a frame in progress
    /// overwrote read from beside it (KERNEL.20), as a value cell holds it,
    /// an empty result 0; else its slot, a value cell's content as written,
    /// a formula's value, its reason, or nothing before its first value.
    fn held(&self, s: usize, row: u32, col: u32) -> Held {
        if !self.written_before.is_empty() {
            if let Some(was) = self.written_before.get(&(s, row, col)) {
                return was.clone();
            }
        }
        if !self.before.is_empty() {
            let p = self.shapes.position((s, row, col));
            if let Some(was) = p.and_then(|p| self.before.get(&(p as u32))) {
                return match was {
                    None => Held::Nothing,
                    Some(Computed::NotComputed(r)) => Held::Reason(r.clone()),
                    Some(Computed::Value(v)) => Held::Cell(value_content(v)),
                };
            }
        }
        store_held(&self.stores[s], row, col)
    }

    /// What a twin holds at a cell of its sheet's (KERNEL.49), nothing where
    /// its sheet holds none, a write emptying no cell.
    fn twin_held(&self, s: usize, row: u32, col: u32) -> Held {
        self.held(s, row, col)
    }

    /// A twin's cells inside a window, row by row (KERNEL.49): each cell of
    /// its sheet there whose twin holds a value, its content as a value cell
    /// holds it, in the default format.
    fn twin_cells(&self, s: usize, window: &A1Range) -> Vec<((u32, u32), Cell)> {
        let mut out = Vec::new();
        self.stores[s].each_held(window, &mut |row, col, _| {
            if let Held::Cell(content) = self.held(s, row, col) {
                out.push(((row, col), Cell { content, style: 0 }));
            }
        });
        out
    }

    /// Each cell a write changed, as its twin held it at the last frame's
    /// end (KERNEL.49), where no write since kept it already: read off its
    /// store's slot, which the settle has yet to write, a value cell's
    /// content as written, a formula's value, its reason, or nothing for a
    /// cell the write added.
    fn remember(&mut self, changed: &BTreeMap<CellId, u32>) {
        for &id in changed.keys() {
            if self.written_before.contains_key(&id) {
                continue;
            }
            let held = store_held(&self.stores[id.0], id.1, id.2);
            self.written_before.insert(id, held);
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
                let w = window_over(self, sheet, window, kept)?;
                Ok(Viewed::Text(record_text(self, &w, kept, true)))
            }
            "values" => {
                let w = window_over(self, sheet, window, kept)?;
                Ok(Viewed::Text(record_text(self, &w, kept, false)))
            }
            "changes" => {
                let w = window_over(self, sheet, window, kept)?;
                Ok(Viewed::Text(self.changes(&w)))
            }
            "plane" => {
                let w = window_over(self, sheet, window, kept)?;
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
                Ok(Viewed::Bytes(plane::bytes_in(self, self, &w)))
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
    /// changed a formula or moved a sheet's extent (KERNEL.56), and while a
    /// value written since lies in the window. A twin is written only at a
    /// frame's end, so a value written since leaves its window still.
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
            return changes_text(self, w, self, (since, at), None);
        }
        let (t, bit) = match rows {
            true => (self.sheets + i, SHEET_ALIKE),
            false => (i, TWIN_ALIKE),
        };
        // The twin cells of the sheet the end changed: a formula's by its
        // position, which no write since moved (one that plans again makes
        // the changes whole), and a value cell's by its place.
        let s = t - self.sheets;
        let blocks = self.shapes.blocks();
        let mut finder = Locator::default();
        let formulas = self.twin_changed.iter().filter_map(|p| {
            let ((on, row, col), _) = finder.cell(blocks, p);
            (on == s).then_some((t, row, col))
        });
        let written = self.twin_written.iter().copied().filter(|id| id.0 == t);
        let mut cells: Vec<(u32, u32)> = formulas
            .chain(written)
            .filter(|id| inside(id))
            .filter(|id| !self.alike.iter().any(|(a, b)| a == id && b & bit != 0))
            .map(|id| (id.1, id.2))
            .collect();
        cells.sort_unstable();
        cells.dedup();
        changes_text(self, w, self, (since, at), Some(&cells))
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
            let sheets: Vec<&str> = self.book.pages[..self.sheets]
                .iter()
                .map(|p| p.sheet.name.as_str())
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
            return Err(load::into_twin(line, &self.book.pages[si].sheet.name, addr));
        }
        // The cap: what the range would add to the grid.
        let count = range.cells();
        if count > CELLS as u64 {
            return Err(load::too_many(line, sheet, addr, self.cells as u64 + count));
        }
        let present = self.present(si, range, log);
        let grows = count - present;
        let now = (self.cells + log.added) as u64;
        if now + grows > CELLS as u64 {
            return Err(load::too_many(line, sheet, addr, now + grows));
        }
        if head == "derived" {
            let mut formula_cell: Option<CellId> = None;
            self.book.pages[si].formulas_in(&range, &mut |(row, col), _| {
                formula_cell.get_or_insert((si, row, col));
            });
            if let Some(id) = formula_cell {
                return Err(LineRefusal {
                    line,
                    refusal: raise(
                        "grid-write-derived",
                        &[
                            ("target", &format!("{sheet}!{addr}")),
                            ("cell", &spell_at(&self.book, id)),
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

    /// How many cells a range holds as the write in progress leaves it: those
    /// its sheet's store holds, a slot at every cell (KERNEL.49), and those an
    /// earlier row of the write added.
    fn present(&self, si: usize, range: A1Range, log: &Log) -> u64 {
        let store = &self.stores[si];
        let added = log
            .changed
            .range((si, range.top, range.left)..=(si, range.bottom, range.right))
            .filter(|(&(_, _, col), _)| col >= range.left && col <= range.right)
            .filter(|(&(_, row, col), _)| !store.slot_of(row, col).is_some_and(|x| !x.is_empty()))
            .count();
        store.count_in(&range) + added as u64
    }

    /// Whether a cell holds anything as the write in progress leaves it.
    fn holds(&self, (si, row, col): CellId, log: &Log) -> bool {
        log.changed.contains_key(&(si, row, col))
            || self.stores[si]
                .slot_of(row, col)
                .is_some_and(|x| !x.is_empty())
    }

    /// Before a range is written over (KERNEL.49's decision 2): every formula
    /// alone it holds given up, and every fill it meets cut, the written
    /// rectangle taken out of each of its pieces and the rest kept as at
    /// most four, a fill written over whole gone. A fill whose first cell the
    /// range holds gives that cell up: its first cell left, row by row,
    /// becomes its anchor, holding the text moved there once, where the move
    /// keeps the text whole ([`crate::refers::moves_whole`]), so that a
    /// release costs one move whatever the fill's size (KERNEL.34); else
    /// each of its cells is given its own text, a formula alone, and placed
    /// again when the write settles. The sheet's fills and formulas as the
    /// write found them are kept for an undo.
    fn release(&mut self, si: usize, range: A1Range, log: &mut Log) {
        let inside = |row: u32, col: u32| {
            row >= range.top && row <= range.bottom && col >= range.left && col <= range.right
        };
        let page = &mut self.book.pages[si];
        let alone: Vec<(u32, u32)> = page
            .alone
            .range((range.top, range.left)..=(range.bottom, range.right))
            .map(|(&at, _)| at)
            .filter(|&(row, col)| inside(row, col))
            .collect();
        if !alone.is_empty() {
            log.sheets.entry(si).or_insert(self.extents[si]);
        }
        for at in alone {
            let old = page.alone.remove(&at);
            log.alone.push(((si, at.0, at.1), old));
        }
        if !page.fills.meets(&range) {
            return;
        }
        log.sheets.entry(si).or_insert(self.extents[si]);
        log.fills
            .entry(si)
            .or_insert_with(|| (page.fills.clone(), page.sheet.shared_formulas));
        for f in page.fills.cut(&range) {
            let Some(fill) = page.fills.fill(f) else {
                // Written over whole: gone.
                continue;
            };
            let anchor = fill.anchor;
            if !inside(anchor.0, anchor.1) {
                continue;
            }
            let first = page.fills.first_cell(f).expect("a fill holds a cell");
            let at = (
                i64::from(first.0) - i64::from(anchor.0),
                i64::from(first.1) - i64::from(anchor.1),
            );
            let whole = crate::refers::moves_whole(&fill.text, at, at);
            #[cfg(test)]
            crate::sheet::work::released(whole);
            if whole {
                let moved = shift_a1_references(&fill.text, at.0, at.1);
                page.fills.anchor_at(f, first, moved);
                continue;
            }
            let (fill, cells) = page.fills.remove(f).expect("a fill");
            for (row, col) in cells {
                let own = shift_a1_references(
                    &fill.text,
                    i64::from(row) - i64::from(anchor.0),
                    i64::from(col) - i64::from(anchor.1),
                );
                let old = page.alone.insert((row, col), (own, Kind::Plain));
                log.alone.push(((si, row, col), old));
                log.released.insert((si, row, col));
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
        // The range's formula cells as the write finds them, each written
        // over; a value cell holding the written value already is not.
        let mut formulas: BTreeSet<(u32, u32)> = BTreeSet::new();
        self.book.pages[si].formulas_in(&range, &mut |at, _| {
            formulas.insert(at);
        });
        self.release(si, range, log);
        for row in range.top..=range.bottom {
            for col in range.left..=range.right {
                let id = (si, row, col);
                let formula = formulas.contains(&(row, col));
                if !formula {
                    let same = match log.contents.get(&id) {
                        Some(held) => held == content,
                        None => self.stores[si]
                            .constant(row, col)
                            .is_some_and(|held| held.is(content)),
                    };
                    if same {
                        continue;
                    }
                }
                if !formula && !self.holds(id, log) {
                    log.added += 1;
                    *log.added_on.entry(si).or_default() += 1;
                }
                log.contents.insert(id, content.clone());
                log.sheets.entry(si).or_insert(self.extents[si]);
                log.changed.entry(id).or_insert(line);
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
        // The range's cells counted before its formulas are given up, a cell
        // the write adds counted once (KERNEL.34; the cells were looked up
        // four times).
        for row in range.top..=range.bottom {
            for col in range.left..=range.right {
                if !self.holds((si, row, col), log) {
                    log.added += 1;
                    *log.added_on.entry(si).or_default() += 1;
                }
            }
        }
        self.release(si, range, log);
        let page = &mut self.book.pages[si];
        log.fills
            .entry(si)
            .or_insert_with(|| (page.fills.clone(), page.sheet.shared_formulas));
        // A range of one cell a formula alone, a larger one a fill under the
        // sheet's next shared index, as `Sheet::set_formula` numbers them.
        match page.sheet.next_shared(range) {
            Some(index) => page
                .fills
                .add(index, range, (range.top, range.left), text, range),
            None => {
                let at = (range.top, range.left);
                let old = page.alone.insert(at, (text.to_string(), Kind::Plain));
                log.alone.push(((si, at.0, at.1), old));
            }
        }
        for row in range.top..=range.bottom {
            for col in range.left..=range.right {
                let id = (si, row, col);
                log.contents.remove(&id);
                log.changed.entry(id).or_insert(line);
            }
        }
        log.sheets.entry(si).or_insert(self.extents[si]);
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
        let page = &self.book.pages[si];
        let masters = page.fills.masters();
        let start = (range.top, range.left);
        let same_at = |row: u32, col: u32| {
            let moved = shift_a1_references(
                text,
                i64::from(row) - i64::from(range.top),
                i64::from(col) - i64::from(range.left),
            );
            page.formula(row, col)
                .and_then(|f| text_of(f, (row, col), &masters))
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
                match page.formula(row, col) {
                    Some(Formula::Fill(index)) => {
                        fills.entry(index).or_default().push((row, col));
                    }
                    Some(Formula::Own(_, Kind::Plain | Kind::Dynamic)) => {
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
        let page = &self.book.pages[si];
        let masters = page.fills.masters();
        let mut differs: Vec<(u32, u32)> = Vec::new();
        for row in range.top..=range.bottom {
            for col in range.left..=range.right {
                let moved = shift_a1_references(
                    text,
                    i64::from(row) - i64::from(range.top),
                    i64::from(col) - i64::from(range.left),
                );
                let old_text = page
                    .formula(row, col)
                    .and_then(|f| text_of(f, (row, col), &masters));
                if old_text.as_deref() != Some(strip_future_prefixes(&moved).as_str()) {
                    differs.push((row, col));
                }
            }
        }
        differs
    }

    /// A write's rows put back (KERNEL.49): every formula alone it set or
    /// gave up as it was, latest first, and each sheet's fills it changed as
    /// the write found them; the stores, which only the settle writes, as
    /// they stand.
    fn undo(&mut self, log: &Log) {
        for ((s, row, col), old) in log.alone.iter().rev() {
            let alone = &mut self.book.pages[*s].alone;
            match old {
                Some(held) => {
                    alone.insert((*row, *col), held.clone());
                }
                None => {
                    alone.remove(&(*row, *col));
                }
            }
        }
        for (&s, (fills, shared)) in &log.fills {
            let page = &mut self.book.pages[s];
            page.fills = fills.clone();
            page.sheet.shared_formulas = *shared;
        }
        // Each written sheet's extent as the write found it (KERNEL.54),
        // where every cell of each was walked again.
        for (&s, &extent) in &log.sheets {
            self.extents[s] = extent;
        }
    }

    /// A write's rows all taken: the extents, the cells counted, the plan
    /// read again when a formula changed, and refused (and undone) when the
    /// grid could no longer be stepped; each written sheet's store kept over
    /// its extent, a changed value cell's slot written with what it now holds
    /// and a changed formula's value forgotten (KERNEL.49).
    fn settle(&mut self, log: Log) -> Result<usize, LineRefusal> {
        if log.changed.is_empty() {
            return Ok(log.written);
        }
        // No write empties a cell (below), so each written sheet's extent is
        // its extent grown by the cells the write changed (KERNEL.54), where
        // every cell of the sheet was walked for it: one value written into
        // a sheet of 458,753 cells cost 5.2 ms.
        for &(s, row, col) in log.changed.keys() {
            self.extents[s] = grown(self.extents[s], row, col);
        }
        let formula_changed = log
            .changed
            .keys()
            .any(|id| self.shapes.holds(*id) || self.book.formula(*id).is_some());
        if formula_changed {
            let saved = self.shapes.clone();
            // Every cell the write changed or a release gave its own text,
            // placed again at once, and the plan made a load's (KERNEL.34).
            let again: BTreeSet<CellId> =
                log.changed.keys().chain(&log.released).copied().collect();
            let again: Vec<CellId> = again.into_iter().collect();
            self.shapes.replace_many(&self.book, &again, &self.library);
            self.shapes
                .plan_model(&self.book, &self.names, &self.extents);
            let placed: Vec<CellId> = log
                .changed
                .keys()
                .copied()
                .filter(|id| self.shapes.holds(*id))
                .collect();
            if let Some((cell, refusal)) = self.shapes.first_refusal_in(&self.book, Some(&placed)) {
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
            // The values stand where their cells stand, by place (KERNEL.49):
            // a plan made again moves none, where it moved each to its cell's
            // new position in the order (KERNEL.26).
        }
        // Each changed cell's twin as the last frame's end left it, kept
        // until the next end (KERNEL.49), read off its slot before the slot
        // is written.
        self.remember(&log.changed);
        // The indexes (KERNEL.8): every one dropped with a plan made again,
        // else each whose vector holds a changed cell built again at its next
        // probe.
        if formula_changed {
            self.indexes.clear();
        } else if self.indexes.any() {
            for id in log.changed.keys() {
                self.indexes.changed(*id);
            }
        }
        // A write never empties a cell: it only adds cells or changes them.
        // A formula changed, and the next frame computes every cell; a value
        // cell changed or added, and the next follows it to the cells that
        // read it (KERNEL.20; KERNEL.56): a value cell added moves no
        // formula's place, and its readers are found by its place as a
        // changed one's are, whether or not it held a value. The room kept
        // for the twin cells' changes grows with the cells, as `fit` keeps
        // it, which a write that refitted the plan for an added cell emptied
        // the shapes' parts to find.
        self.cells += log.added;
        for (&s, &n) in &log.added_on {
            self.book.pages[s].cells += n;
        }
        if formula_changed {
            self.every_next = true;
            // A cell's kind may have changed, which no twin cell's change
            // carries: the changes whole until the next frame's end and at
            // it (KERNEL.42).
            self.reshaped = true;
            self.fit();
        } else if log.added > 0 {
            self.room = self.cells / 2 + 64;
        }
        // A sheet whose extent the write moved (KERNEL.56): every window's
        // record of it names its extent in its head, which no cell's change
        // carries, so the changes are whole until the next frame's end and at
        // it; and the next frame computes whole every shape that reads its
        // extent itself.
        for (&s, &was) in &log.sheets {
            if self.extents[s] != was {
                self.reshaped = true;
                self.note_extent_moved(s);
            }
        }
        // Each written sheet's store over its extent as it now stands
        // (KERNEL.49): kept where its extent stands, else its slots moved by
        // place, an array or a map by the rule at the sheet's cells now, an
        // array grown where it stands (KERNEL.56).
        for &s in log.sheets.keys() {
            if let Store::Array(d) = &self.stores[s] {
                if d.stands_over(self.extents[s]) {
                    #[cfg(test)]
                    if formula_changed || log.added > 0 {
                        self.stores_kept += 1;
                    }
                    continue;
                }
            }
            let limit = dense::limit(self.book.pages[s].cells, CELLS as u64);
            #[cfg(test)]
            let was_sparse = matches!(self.stores[s], Store::Sparse(_));
            let store = std::mem::replace(
                &mut self.stores[s],
                Store::over(None, 0, std::iter::empty()),
            );
            self.stores[s] = store.moved(self.extents[s], limit);
            #[cfg(test)]
            if !(was_sparse && matches!(self.stores[s], Store::Sparse(_))) {
                self.stores_moved += 1;
            }
        }
        // Each changed cell's slot: a value cell's content as written, a
        // formula's value forgotten until the next frame computes it.
        for id in log.changed.keys() {
            match log.contents.get(id) {
                Some(content) => self.stores[id.0].put_content(id.1, id.2, Some(content)),
                None => self.stores[id.0].put_computed(id.1, id.2, None),
            }
            if !self.shapes.holds(*id) {
                self.written.push(*id);
            }
        }
        // A plane kept for a sheet the write changed takes each changed
        // cell's byte where it stands, the plane taken over an extent the
        // write grew (KERNEL.56), where it was dropped and the walk answered
        // until the next frame's end made it again from every cell of its
        // sheet (KERNEL.24); one past the limit is dropped. A refused write
        // never reaches here, having changed nothing.
        for &s in log.sheets.keys() {
            let Some(raster) = self.planes[s].as_mut() else {
                continue;
            };
            if !raster.regrow(self.extents[s], CELLS as u64) {
                self.planes[s] = None;
                continue;
            }
            for (&(_, row, col), _) in log.changed.range((s, 0, 0)..(s + 1, 0, 0)) {
                // A value cell's slot draws the byte its content does, and a
                // formula's with no value 0.
                let slot = self.stores[s].slot_of(row, col).unwrap_or(Stored::Pending);
                raster.set(row, col, plane::slot_byte(slot));
            }
        }
        // A plan made again may read other twins: each it reads kept as a
        // store, each it no longer reads given up (KERNEL.49).
        if formula_changed {
            self.keep_twins();
        }
        Ok(log.written)
    }
}

impl Values for Machine {
    /// The last complete frame's value of a formula cell, read off its slot
    /// (KERNEL.49): while a step yields, a value the frame in progress
    /// overwrote is read from beside it (KERNEL.20); `None` for a value cell
    /// and for a formula with no value yet.
    fn computed(&self, id: CellId) -> Option<Computed> {
        let p = self.shapes.position(id)?;
        if !self.before.is_empty() {
            if let Some(was) = self.before.get(&(p as u32)) {
                return was.clone();
            }
        }
        self.stores.get(id.0)?.slot_of(id.1, id.2)?.computed()
    }

    /// Every sheet's extent, its twins' included, as the machine keeps it
    /// for the evaluator (KERNEL.37): made at the load, and kept at every
    /// write, every undone write and every frame's end.
    fn kept_extent(&self, sheet: usize) -> Option<Extent> {
        self.extents.get(sheet).copied()
    }
}

impl Env for Machine {
    fn find_sheet(&self, name: &str) -> Option<usize> {
        self.book.find_sheet(name)
    }

    fn sheet_name(&self, sheet: usize) -> Option<&str> {
        self.book.pages.get(sheet).map(|p| p.sheet.name.as_str())
    }

    fn sheets(&self) -> usize {
        self.book.pages.len()
    }

    fn rows(&self, sheet: usize) -> Option<&BTreeMap<u32, RowSettings>> {
        self.book.pages.get(sheet).map(|p| &p.sheet.rows)
    }

    /// A rows' sheet's cells inside an area, and a twin's by its sheet's
    /// (KERNEL.49): the places its store holds a slot at, a value, a reason
    /// or a formula with no value yet, every walk clamped to the extent the
    /// evaluator reads, a twin's the rectangle of its own cells; at every
    /// place of its sheet's the twin holds a value, a reason or nothing,
    /// each read through [`Env::read`].
    fn walk(&self, area: &Area) -> Walked<'_> {
        let s = match area.sheet.checked_sub(self.sheets) {
            Some(s) if s < self.sheets => s,
            Some(_) => return Walked::nothing(),
            None => area.sheet,
        };
        let Some(((top, left), (bottom, right))) = Env::extent(self, area.sheet) else {
            return Walked::nothing();
        };
        let clamped = A1Range {
            top: area.top.max(top),
            left: area.left.max(left),
            bottom: area.bottom.min(bottom),
            right: area.right.min(right),
        };
        if clamped.top > clamped.bottom || clamped.left > clamped.right {
            return Walked::nothing();
        }
        self.stores[s].walk(&clamped)
    }

    fn written(&self, id: CellId) -> Option<Written<'_>> {
        crate::calc::eval::written_in(&self.book, id)
    }

    fn here(&self) -> CellId {
        self.here.get()
    }

    fn read(&self, id: CellId) -> Result<Value, Reason> {
        // A rows' cell read by its place (KERNEL.49): its value, a formula's
        // of the frame, computed in this frame by now, the order putting each
        // cell after what it reads, or kept from the last, its reads
        // unchanged (KERNEL.20); a formula's reason; nothing for a cell
        // holding none.
        if id.0 < self.sheets {
            return match self.stores[id.0].slot_of(id.1, id.2) {
                Some(slot) => slot.read(),
                None => Ok(Value::Empty),
            };
        }
        // A twin's cell (KERNEL.49): its store's slot where a formula reads
        // the twin, a value, a formula's reason or nothing, by its index where
        // the store is an array (KERNEL.20); else made from its sheet's cell
        // as the last frame's end left it. A twin is written only at a
        // frame's end, so its read goes straight to the twin (KERNEL.25). Most
        // of a frame's reads are of a twin.
        match self.twins.get(id.0 - self.sheets) {
            Some(Some(twin)) => match twin.slot_of(id.1, id.2) {
                Some(slot) => slot.read(),
                None => Ok(Value::Empty),
            },
            Some(None) => match self.twin_held(id.0 - self.sheets, id.1, id.2) {
                Held::Cell(content) => Ok(content_value(&content)),
                Held::Reason(r) => Err(r),
                Held::Nothing => Ok(Value::Empty),
            },
            None => Ok(Value::Empty),
        }
    }

    /// A sheet's extent as the evaluator clamps a walk of it and bisects a
    /// range over it (KERNEL.37): a twin's the rectangle of its cells, as
    /// `Calc` reads the twin as a sheet (KERNEL.49), whatever keeps the twin,
    /// its array over its sheet's extent clamped to it.
    fn extent(&self, sheet: usize) -> Extent {
        self.extents.get(sheet).copied().flatten()
    }

    fn name(&self, name: &str, scope: Option<usize>, home: usize) -> Option<&str> {
        self.names.resolve(name, scope, home)
    }

    /// A rows' sheet's store where it is an array (KERNEL.49), and a twin's
    /// store kept as one (KERNEL.20), a reason among the slots of either
    /// meeting a walk where it stands, as a walk of the sheet meets it.
    fn dense(&self, sheet: usize) -> Option<&Dense> {
        if sheet < self.sheets {
            return self.stores[sheet].array().filter(|_| self.arrays);
        }
        self.twin_store(sheet)?.array()
    }

    /// Whether the cell being computed holds a dynamic-array formula, as its
    /// grid says (KERNEL.8, slice b): a grid loaded from rows holds none, a
    /// `formula` row making a plain or a shared formula. A fill is a plain
    /// formula, so only a formula alone is asked, one lookup at every
    /// evaluation that reads a range where one value is wanted, where asking
    /// the grid searched its fills too.
    fn dynamic(&self) -> bool {
        let (s, row, col) = self.here.get();
        self.book
            .pages
            .get(s)
            .and_then(|p| p.alone.get(&(row, col)))
            .is_some_and(|&(_, kind)| kind == Kind::Dynamic)
    }

    fn enter(&self, id: CellId) {
        self.here.set(id);
    }

    /// The indexes of the vectors the lookups and the criteria search
    /// (KERNEL.8), told every cell whose value changes: at a write, at a
    /// frame's computed value and at a twin's copy.
    fn indexes(&self) -> Option<&Indexes> {
        self.indexed.then_some(&self.indexes)
    }
}

impl Model for Machine {
    fn sheet_count(&self) -> usize {
        self.book.pages.len()
    }

    fn sheet_name(&self, sheet: usize) -> Option<&str> {
        Model::sheet_name(&self.book, sheet)
    }

    fn find_sheet(&self, name: &str) -> Option<usize> {
        self.book.find_sheet(name)
    }

    fn defined_names(&self) -> &[(String, String)] {
        &self.book.defined_names
    }

    fn formula(&self, id: CellId) -> Option<Formula<'_>> {
        self.book.formula(id)
    }

    fn masters(&self, sheet: usize) -> Masters<'_> {
        self.book.masters(sheet)
    }

    fn master(&self, sheet: usize, si: u32) -> Option<(&str, u32, u32)> {
        self.book.master(sheet, si)
    }

    fn each_formula(&self, sheet: usize, f: &mut dyn FnMut((u32, u32), Formula<'_>)) {
        self.book.each_formula(sheet, f);
    }

    fn formulas_in(&self, area: &Area, extent: Extent, f: &mut dyn FnMut(u32, u32)) {
        self.book.formulas_in(area, extent, f);
    }
}

impl Cells for Machine {
    fn styles(&self) -> &Styles {
        &self.book.styles
    }

    fn page(&self, sheet: usize) -> grid::Page<'_> {
        grid::Page::of(&self.book.pages[sheet].sheet)
    }

    fn extent(&self, sheet: usize) -> Extent {
        self.extents.get(sheet).copied().flatten()
    }

    /// A rows' sheet's cells as its grid and its store hold them, row by
    /// row: each place its store holds a slot at, a formula's where the grid
    /// holds one and a value cell's content as written elsewhere, lent where
    /// it stands (KERNEL.49); a twin's made from its sheet.
    fn cells<'a>(
        &'a self,
        sheet: usize,
        window: &A1Range,
        f: &mut dyn FnMut((u32, u32), grid::Held<'a>),
    ) {
        let Some(s) = sheet.checked_sub(self.sheets) else {
            let page = &self.book.pages[sheet];
            let store = &self.stores[sheet];
            // Each cell's formula, format and sentence met beside the store's
            // walk, in its order, and never looked up (KERNEL.49).
            let mut formulas = grid::Beside::new(page.formulas_iter(window));
            let mut styles = grid::Beside::new(grid::in_window(&page.styles, window));
            let mut sentences = grid::Beside::new(grid::in_window(&page.sheet.sentences, window));
            for (row, col) in store.walk(window) {
                let holds = match formulas.at((row, col)) {
                    Some(formula) => Holds::Formula(formula),
                    None => match store.constant(row, col) {
                        Some(c) => Holds::Constant(c),
                        None => continue,
                    },
                };
                f(
                    (row, col),
                    grid::Held {
                        holds,
                        style: styles.at((row, col)).unwrap_or(0),
                        sentence: sentences.at((row, col)),
                    },
                );
            }
            return;
        };
        for (at, cell) in self.twin_cells(s, window) {
            if let Some(c) = Constant::owned(cell.content) {
                f(
                    at,
                    grid::Held {
                        holds: Holds::Constant(c),
                        style: 0,
                        sentence: None,
                    },
                );
            }
        }
    }

    fn cell(&self, sheet: usize, row: u32, col: u32) -> Option<grid::Held<'_>> {
        let Some(s) = sheet.checked_sub(self.sheets) else {
            let page = self.book.pages.get(sheet)?;
            let holds = match page.formula(row, col) {
                Some(formula) => Holds::Formula(formula),
                None => Holds::Constant(self.stores[sheet].constant(row, col)?),
            };
            return Some(grid::Held {
                holds,
                style: page.style(row, col),
                sentence: page.sheet.sentences.get(&(row, col)).copied(),
            });
        };
        match self.twin_held(s, row, col) {
            Held::Cell(content) => Some(grid::Held {
                holds: Holds::Constant(Constant::owned(content)?),
                style: 0,
                sentence: None,
            }),
            Held::Reason(_) | Held::Nothing => None,
        }
    }
}

impl Source for Machine {
    /// The call at a program's root answering no number, its answer kept
    /// for the cell (KERNEL.8).
    fn answered(&self, value: Value) {
        self.root_answer.set(Some(value));
    }

    /// [`Env::read`]'s answer as a program's register holds it (KERNEL.44,
    /// `calc::program::number_of`), read where `read` reads it and no value
    /// made: a rows' cell's slot or a twin's (KERNEL.49), a number, nothing
    /// as 0, any other kind none, a reason among them; a twin kept as no
    /// store read as `read` reads it.
    #[inline]
    fn number(&self, id: CellId) -> Option<f64> {
        if id.0 < self.sheets {
            return match self.stores[id.0].slot_of(id.1, id.2) {
                Some(Stored::Number(n)) => Some(n),
                Some(Stored::Empty) | None => Some(0.0),
                Some(_) => None,
            };
        }
        match self.twin_store(id.0) {
            Some(twin) => match twin.slot_of(id.1, id.2) {
                Some(Stored::Number(n)) => Some(n),
                Some(Stored::Empty) | None => Some(0.0),
                Some(_) => None,
            },
            None => match self.read(id) {
                Ok(Value::Number(n)) => Some(n),
                Ok(Value::Empty) => Some(0.0),
                _ => None,
            },
        }
    }
}

/// Every cell of a sheet: the area a walk of a whole sheet takes.
const EVERY: A1Range = A1Range {
    top: 1,
    left: 1,
    bottom: MAX_ROW,
    right: MAX_COLUMN,
};

/// A draft's record as the load keeps it past the plan (KERNEL.49): its
/// rectangle, its line, and its value, `None` for a formula's.
struct Laid {
    range: A1Range,
    line: u32,
    value: Option<Content>,
}

impl Laid {
    fn of(record: load::Record) -> Laid {
        Laid {
            range: record.range,
            line: record.line,
            value: match record.what {
                What::Value(content) => Some(content),
                What::Formula { .. } => None,
            },
        }
    }
}

/// Each rows' sheet's store from a draft's records (KERNEL.49's decision 8):
/// over its extent, an array or a sparse sheet's by the rule at its cells,
/// each value cell's content as written and each formula's value from its
/// `value` row, none for a formula without one.
fn stores_of(
    laid: &[Vec<Laid>],
    book: &Book,
    extents: &[Extent],
    values: &HashMap<CellId, Computed>,
) -> Vec<Store> {
    let mut stores = Vec::with_capacity(laid.len());
    for (s, records) in laid.iter().enumerate() {
        let places = records.iter().flat_map(|r| {
            let x = r.range;
            (x.top..=x.bottom).flat_map(move |row| (x.left..=x.right).map(move |col| (row, col)))
        });
        let limit = dense::limit(book.pages[s].cells, CELLS as u64);
        let mut store = Store::over(extents[s], limit, places);
        for r in records {
            let x = r.range;
            for row in x.top..=x.bottom {
                for col in x.left..=x.right {
                    match &r.value {
                        Some(content) => store.put_content(row, col, Some(content)),
                        None if values.is_empty() => store.put_computed(row, col, None),
                        None => store.put_computed(row, col, values.get(&(s, row, col))),
                    }
                }
            }
        }
        stores.push(store);
    }
    stores
}

/// What a twin holds at a cell as its sheet's store holds it (KERNEL.49): a
/// value cell's content as written, a formula's value as a value cell holds
/// it; a formula's reason, with no cell; nothing for a cell held nowhere and
/// a formula with no value yet.
fn store_held(store: &Store, row: u32, col: u32) -> Held {
    match store.slot_of(row, col) {
        None | Some(Stored::Empty | Stored::Pending) => Held::Nothing,
        Some(Stored::Reason(r)) => Held::Reason(r.clone()),
        Some(_) => store.content(row, col).map_or(Held::Nothing, Held::Cell),
    }
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

/// What a twin holds at a cell (KERNEL.49): a cell, its content as a value
/// cell holds it; a formula's reason, with no cell; or nothing.
#[derive(Clone, Debug, PartialEq)]
enum Held {
    Cell(Content),
    Reason(Reason),
    Nothing,
}

/// A formula's twin from its slot (KERNEL.49): its value as a value cell
/// holds it, an empty result 0; its reason; nothing before its first value.
fn formula_held(slot: Option<Stored<'_>>) -> Held {
    match slot {
        None | Some(Stored::Pending) => Held::Nothing,
        Some(Stored::Reason(r)) => Held::Reason(r.clone()),
        Some(slot) => Held::Cell(slot_content(slot)),
    }
}

/// Whether a formula's slot makes a twin cell: a value does, an empty
/// result 0; a reason and no value yet do not.
fn holds_cell(slot: Option<Stored<'_>>) -> bool {
    !matches!(slot, None | Some(Stored::Pending | Stored::Reason(_)))
}

/// A formula's slot as its twin's store holds it (KERNEL.49): nothing
/// before its first value, every other slot as it stands.
fn twin_slot(slot: Option<Stored<'_>>) -> Stored<'_> {
    match slot {
        None | Some(Stored::Pending) => Stored::Empty,
        Some(slot) => slot,
    }
}

/// A formula's value as its twin cell holds it, a value cell's content: an
/// error by its text, an empty result 0.
fn value_content(v: &Value) -> Content {
    match v {
        Value::Number(n) => Content::Number(*n),
        Value::Text(t) => Content::Text(t.clone()),
        Value::Bool(b) => Content::Bool(*b),
        Value::Error(k) => Content::Error(k.text().to_string()),
        Value::Empty => Content::Number(0.0),
    }
}

/// The records that spell a formula's twin cell's change as before
/// (KERNEL.42), from its slot to the value the frame computed, the twin
/// neither gaining nor losing the cell: for a value turned to another, as
/// [`alike`] tells them; for a reason turned to another, the twin's always,
/// holding no cell either side, and the sheet's where the two are spelled
/// alike; for a reason a formula with no value meets, the twin's alone,
/// which spells nothing either side (KERNEL.56: named among a twin's
/// changes, it was found once a load's first frame's changes were no longer
/// whole).
fn changed_alike(was: Option<Stored<'_>>, now: &Computed) -> u8 {
    match (was, now) {
        (Some(Stored::Reason(a)), Computed::NotComputed(b)) => alike(
            true,
            (Spelled::Nothing, Some(a)),
            (Spelled::Nothing, Some(b)),
        ),
        (None | Some(Stored::Pending), Computed::NotComputed(b)) => {
            alike(true, (Spelled::Nothing, None), (Spelled::Nothing, Some(b)))
        }
        (Some(slot), Computed::Value(v)) => {
            alike(true, (Spelled::slot(slot), None), (Spelled::value(v), None))
        }
        _ => 0,
    }
}

/// A written cell's twin told at a frame's end against the twin as the
/// write found it (KERNEL.49): whether it changed, a number to the bit; the
/// records that spell the change as before (KERNEL.42), for a cell turned to
/// another as [`alike`] tells them, `formula` whether the cell is a formula
/// now, and for a reason turned to another as [`changed_alike`] tells them;
/// and whether the twin gained or lost the cell.
fn told(formula: bool, was: &Held, now: &Held) -> (bool, u8, Grew) {
    match (was, now) {
        (Held::Cell(a), Held::Cell(b)) => match same_content(a, b) {
            true => (false, 0, Grew::Neither),
            false => (
                true,
                alike(
                    formula,
                    (Spelled::content(Some(a)), None),
                    (Spelled::content(Some(b)), None),
                ),
                Grew::Neither,
            ),
        },
        (Held::Cell(_), _) => (true, 0, Grew::Lost),
        (_, Held::Cell(_)) => (true, 0, Grew::Gained),
        (Held::Reason(a), Held::Reason(b)) => match a == b {
            true => (false, 0, Grew::Neither),
            false => (
                true,
                alike(
                    true,
                    (Spelled::Nothing, Some(a)),
                    (Spelled::Nothing, Some(b)),
                ),
                Grew::Neither,
            ),
        },
        (Held::Nothing, Held::Nothing) => (false, 0, Grew::Neither),
        // A reason gained or lost, with no cell either side.
        _ => (true, 0, Grew::Neither),
    }
}

/// A written value cell's twin told at a frame's end against the twin as the
/// write found it, as [`told`] tells it, the cell's content read where it
/// stands: a cell gained where the twin held a reason or nothing.
fn told_cell(was: &Held, now: &Constant<'_>) -> (bool, u8, Grew) {
    match was {
        Held::Cell(a) if same_constant(a, now) => (false, 0, Grew::Neither),
        Held::Cell(a) => (
            true,
            alike(
                false,
                (Spelled::content(Some(a)), None),
                (Spelled::constant(now), None),
            ),
            Grew::Neither,
        ),
        Held::Reason(_) | Held::Nothing => (true, 0, Grew::Gained),
    }
}

/// Whether a twin gained a cell at a frame's end, lost one, or neither
/// (KERNEL.56): a cell gained grows its twin's extent by its place, and one
/// lost makes the extent again and the next frame compute every cell.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Grew {
    Neither,
    Gained,
    Lost,
}

impl Grew {
    /// The change told to its sheet's flag of a cell lost and its rectangle
    /// of the cells gained.
    fn tell(self, lost: &mut bool, gained: &mut Extent, row: u32, col: u32) {
        match self {
            Grew::Neither => {}
            Grew::Gained => *gained = grown(*gained, row, col),
            Grew::Lost => *lost = true,
        }
    }
}

/// An extent grown by a cell (KERNEL.54; KERNEL.56): the cell alone where
/// there was none.
fn grown(extent: Extent, row: u32, col: u32) -> Extent {
    Some(match extent {
        None => ((row, col), (row, col)),
        Some(((top, left), (bottom, right))) => (
            (top.min(row), left.min(col)),
            (bottom.max(row), right.max(col)),
        ),
    })
}

/// A twin cell kept as changed for the next frame (KERNEL.20), counted;
/// past the room kept for them, half the cells, the next frame made to
/// compute every cell.
fn note(count: &mut usize, room: usize, every_next: &mut bool) {
    *count += 1;
    if *count > room {
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

    /// A value cell's content where it stands (KERNEL.49).
    fn constant(c: &'a Constant<'_>) -> Spelled<'a> {
        match c {
            Constant::Number(n) => Spelled::number(*n),
            Constant::Text(t) => Spelled::Text(t),
            Constant::Bool(b) => Spelled::Bool(*b),
            Constant::Error(e) => Spelled::Error(e),
        }
    }

    /// A formula's slot as its twin holds it, which the twin holds exactly:
    /// an empty result 0, nothing before its first value or for a reason.
    fn slot(slot: Stored<'a>) -> Spelled<'a> {
        match slot {
            Stored::Number(n) => Spelled::number(n),
            Stored::Text(t) => Spelled::Text(t),
            Stored::Bool(b) => Spelled::Bool(b),
            Stored::Error(k) => Spelled::Error(k.text()),
            Stored::Empty => Spelled::number(0.0),
            Stored::Reason(_) | Stored::Pending => Spelled::Nothing,
        }
    }

    /// A formula's value as its twin holds it, an empty result 0.
    fn value(v: &'a Value) -> Spelled<'a> {
        match v {
            Value::Number(n) => Spelled::number(*n),
            Value::Text(t) => Spelled::Text(t),
            Value::Bool(b) => Spelled::Bool(*b),
            Value::Error(k) => Spelled::Error(k.text()),
            Value::Empty => Spelled::number(0.0),
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

/// A shape made whole for the frame in progress, listed as it first is,
/// and every cell of it marked through its stretches a word at a time
/// (KERNEL.54), so that a step passing from mark to mark computes them.
fn make_whole(
    shapes: &Shapes,
    whole: &mut [bool],
    wholes: &mut Vec<u32>,
    marked: &mut Bits,
    s: usize,
) {
    if whole[s] {
        return;
    }
    whole[s] = true;
    wholes.push(s as u32);
    for &(start, end) in shapes.stretches(s) {
        marked.set_range(start as usize, end as usize);
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

/// Whether a slot holds a computed value to the bit (KERNEL.20; KERNEL.49),
/// a number's sign of zero included, a reason the same reason; a formula
/// with no value holds none.
fn same_held(slot: Stored<'_>, computed: &Computed) -> bool {
    match (slot, computed) {
        (Stored::Number(x), Computed::Value(Value::Number(y))) => x.to_bits() == y.to_bits(),
        (Stored::Text(x), Computed::Value(Value::Text(y))) => x == y,
        (Stored::Bool(x), Computed::Value(Value::Bool(y))) => x == *y,
        (Stored::Error(x), Computed::Value(Value::Error(y))) => x == *y,
        (Stored::Empty, Computed::Value(Value::Empty)) => true,
        (Stored::Reason(x), Computed::NotComputed(y)) => x == y,
        _ => false,
    }
}

/// [`same_value`] over a twin cell's content and a value cell's where it
/// stands (KERNEL.49).
fn same_constant(a: &Content, b: &Constant<'_>) -> bool {
    match (a, b) {
        (Content::Number(x), Constant::Number(y)) => x.to_bits() == y.to_bits(),
        (Content::Text(x), Constant::Text(y)) => x == y,
        (Content::Bool(x), Constant::Bool(y)) => x == y,
        (Content::Error(x), Constant::Error(y)) => x == y,
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

/// A twin's slot as its cell holds it (KERNEL.27), which the slot holds
/// exactly.
fn slot_content(slot: Stored<'_>) -> Content {
    match slot {
        Stored::Number(n) => Content::Number(n),
        Stored::Text(t) => Content::Text(t.to_string()),
        Stored::Bool(b) => Content::Bool(b),
        Stored::Error(k) => Content::Error(k.text().to_string()),
        Stored::Empty | Stored::Reason(_) | Stored::Pending => Content::Number(0.0),
    }
}

/// What a frame touched beside the cells it computed, counted in a test
/// (KERNEL.54's oracle), as `crate::calc::reach::checked` counts the readers
/// checked: the positions of cells alone the step visited, the runs it
/// visited, the shapes a frame's start asked for their parts, and the flags
/// and words a frame's end cleared. Counted on the thread that asked, between
/// [`walk::begin`] and [`walk::end`], and only in a test build.
#[cfg(test)]
pub(crate) mod walk {
    use std::cell::Cell;

    /// What the frames since [`begin`] touched.
    #[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
    pub(crate) struct Walk {
        /// The positions of cells alone the step visited: each marked one a
        /// frame that computes what changed computes, and each a frame that
        /// computes every cell passes.
        pub(crate) visited: u64,
        /// The runs the step visited.
        pub(crate) runs: u64,
        /// The shapes a frame's start asked for their parts.
        pub(crate) shapes: u64,
        /// The flags and words a frame's end cleared.
        pub(crate) cleared: u64,
    }

    thread_local! {
        static ON: Cell<bool> = const { Cell::new(false) };
        static DONE: Cell<Walk> = const {
            Cell::new(Walk {
                visited: 0,
                runs: 0,
                shapes: 0,
                cleared: 0,
            })
        };
    }

    /// Counting begins on this thread, from nothing.
    pub(crate) fn begin() {
        ON.with(|on| on.set(true));
        DONE.with(|done| done.set(Walk::default()));
    }

    /// What was counted since [`begin`]; counting ends.
    pub(crate) fn end() -> Walk {
        ON.with(|on| on.set(false));
        DONE.with(Cell::get)
    }

    fn add(f: impl FnOnce(&mut Walk)) {
        if ON.with(Cell::get) {
            DONE.with(|done| {
                let mut walk = done.get();
                f(&mut walk);
                done.set(walk);
            });
        }
    }

    /// Positions of cells alone visited.
    pub(crate) fn visited(n: u64) {
        add(|w| w.visited += n);
    }

    /// One run visited.
    pub(crate) fn run() {
        add(|w| w.runs += 1);
    }

    /// Shapes asked for their parts.
    pub(crate) fn shapes(n: u64) {
        add(|w| w.shapes += n);
    }

    /// Flags and words cleared.
    pub(crate) fn cleared(n: u64) {
        add(|w| w.cleared += n);
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
    use crate::calc::graph::is_formula;
    use crate::calc::library;
    use crate::calc::program;
    use crate::calc::shape::Shapes;
    use crate::projection::Window;
    use crate::reader::read_forms;
    use crate::sheet::{cell_ref, formula_text, shared_masters};
    use crate::view::{view_text_valued, view_text_values};

    fn load(text: &str) -> Machine {
        Machine::load(text, labelling())
            .unwrap_or_else(|e| panic!("line {}: {}", e.line, e.refusal.text))
    }

    /// The language's library with [`label`] beside it.
    fn labelling() -> Library {
        let mut lib = library::language();
        lib.register_form("LABEL", label);
        lib
    }

    /// The reason a range where one value is wanted held before KERNEL.8's
    /// second slice met it at the cell: `Ctx::one`'s label, the reference as
    /// written moved with its cell, owned through a name; one cell its
    /// value. A form of the tests' library alone, so that the tests of a
    /// reason at a step, a twin's and a plane's among them, keep one to
    /// stand on, a legacy formula's range now computing.
    fn label<'e>(
        ctx: &crate::calc::Ctx<'_>,
        args: &'e [crate::calc::Expr],
    ) -> Result<crate::calc::Arg<'e>, Reason> {
        use crate::calc::Arg;
        let arg = ctx.eval(&args[0])?;
        if let Arg::Area {
            areas,
            written,
            moved,
        } = &arg
        {
            if !matches!(&areas[..], [a] if a.is_cell()) {
                return Err(Reason::Construct(if *moved == (0, 0) {
                    written.to_string()
                } else {
                    crate::refers::shift_a1_references(written, moved.0, moved.1)
                }));
            }
        }
        Ok(Arg::Scalar(ctx.one(&arg)?))
    }

    fn refused(text: &str) -> LineRefusal {
        match Machine::load(text, labelling()) {
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
        m.computed((i, r.top, r.left))
    }

    fn number(m: &Machine, sheet: &str, addr: &str) -> f64 {
        match at(m, sheet, addr) {
            Some(Computed::Value(Value::Number(n))) => n,
            other => panic!("{sheet}!{addr}: {other:?}"),
        }
    }

    /// A cell's content, a twin's included, a twin's made from its sheet as
    /// a view makes it (KERNEL.49).
    fn content(m: &Machine, sheet: &str, addr: &str) -> Option<Content> {
        let i = m.find_sheet(sheet).unwrap();
        let r = parse_a1_range(addr).unwrap();
        if i < m.sheets {
            return m.workbook().sheets[i]
                .cells
                .get(&(r.top, r.left))
                .map(|c| c.content.clone());
        }
        match Cells::cell(m, i, r.top, r.left)?.holds {
            Holds::Constant(c) => Some(c.content()),
            Holds::Formula(_) => panic!("a twin holds no formula"),
        }
    }

    /// The formula rows of every sheet's record, the twins' included.
    fn formula_rows(m: &Machine) -> Vec<String> {
        let mut out = Vec::new();
        for s in m.book.pages.iter().map(|p| &p.sheet) {
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
        for s in m.book.pages[..m.sheets()].iter().map(|p| &p.sheet) {
            out.push_str(&grid(m, &s.name, None));
        }
        out
    }

    /// Values held for a view of a draft.
    #[derive(Debug)]
    struct Drafted(HashMap<CellId, Computed>);

    impl Values for Drafted {
        fn computed(&self, id: CellId) -> Option<Computed> {
            self.0.get(&id).cloned()
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
                "(formula \"S\" \"C3\" \"=A1#\")\n",
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
        assert_eq!(m.book.pages.len(), 2);
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
        assert_eq!(m.shapes.len(), m.shapes.formulas(), "{when}");
        for id in m.shapes.order_cells() {
            assert!(id.0 < m.sheets, "{when}: {id:?} placed on a twin");
        }
        // A store for every rows' sheet, by place (KERNEL.49), each value
        // kept beside a formula's of the order, and the order holds no
        // twin's cell.
        assert_eq!(m.stores.len(), m.sheets, "{when}: the stores");
        for &p in m.before.keys() {
            assert!((p as usize) < m.shapes.len(), "{when}: {p} kept beside");
        }
        // A twin keeps no cell (KERNEL.49): its values stand in a store of
        // its own, or are made from its sheet.
        for page in &m.book.pages[m.sheets..] {
            let none = page.cells == 0
                && page.alone.is_empty()
                && page.fills.every_piece().next().is_none();
            assert!(none, "{when}: {}'s cells", page.sheet.name);
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
        assert_eq!(
            m.write("(formula \"Board\" \"J1\" \"=LABEL(F1:G1)*2\")"),
            Ok(1)
        );
        m.step(0);
        let board = m.find_sheet("Board").unwrap();
        assert!(
            matches!(m.twin_held(board, 1, 10), Held::Reason(_)),
            "J1's reason kept"
        );
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

    /// Each rows' sheet's values as a frame's end or a load left them, by
    /// place (KERNEL.49): the reference its twin is held to until the next
    /// frame's end, spelled here from the cells and the last complete
    /// frame's values, a value cell's value, a formula's, an empty result 0,
    /// or its reason, and nothing for a formula with no value. Taken at a
    /// frame's end or a load, or while a step yields, before any write since.
    type Left = Vec<BTreeMap<(u32, u32), Result<Value, Reason>>>;

    fn sheets_left(m: &Machine) -> Left {
        let book = m.workbook();
        (0..m.sheets)
            .map(|s| {
                book.sheets[s]
                    .cells
                    .iter()
                    .filter_map(|(&at, cell)| {
                        if !is_formula(&cell.content) {
                            return Some((at, Ok(content_value(&cell.content))));
                        }
                        let held = match m.computed((s, at.0, at.1))? {
                            Computed::Value(v) => Ok(v.settled()),
                            Computed::NotComputed(r) => Err(r),
                        };
                        Some((at, held))
                    })
                    .collect()
            })
            .collect()
    }

    /// The rectangle of the values a reference holds of a sheet, its reasons
    /// outside it: a twin's extent, the rectangle of its cells (KERNEL.49).
    fn extent_of(held: &BTreeMap<(u32, u32), Result<Value, Reason>>) -> Extent {
        let values = held.iter().filter(|(_, v)| v.is_ok()).map(|(at, _)| at);
        values.fold(None, |extent, &(row, col)| {
            Some(match extent {
                None => ((row, col), (row, col)),
                Some(((top, left), (bottom, right))) => (
                    (top.min(row), left.min(col)),
                    (bottom.max(row), right.max(col)),
                ),
            })
        })
    }

    /// A twin's cell as the reference holds it, `s` its rows' sheet: nothing
    /// where its sheet's cell held nothing.
    fn twin_cell(last: &Left, s: usize, row: u32, col: u32) -> Result<Value, Reason> {
        last[s]
            .get(&(row, col))
            .cloned()
            .unwrap_or(Ok(Value::Empty))
    }

    /// How many reasons the twins hold, as a view and a read make them.
    fn twin_reasons(m: &Machine) -> usize {
        let book = m.workbook();
        (0..m.sheets)
            .map(|s| {
                book.sheets[s]
                    .cells
                    .keys()
                    .filter(|&&(row, col)| matches!(m.held(s, row, col), Held::Reason(_)))
                    .count()
            })
            .sum()
    }

    /// Whether a sheet's twin is kept as an array (KERNEL.20), asked as the
    /// evaluator asks it, of the twin.
    fn dense_kept(m: &Machine, sheet: &str) -> bool {
        let twin = m.find_sheet(&format!("{sheet}.last")).unwrap();
        m.dense(twin).is_some()
    }

    /// Whether a sheet's twin keeps a store (KERNEL.49), and whether it is an
    /// array; `None` for a twin kept as no store, made from its sheet.
    fn store_kept(m: &Machine, sheet: &str) -> Option<bool> {
        let s = m.find_sheet(sheet).unwrap();
        m.twins[s].as_ref().map(|t| t.array().is_some())
    }

    /// Every twin read and walked by the evaluator held to the reference
    /// (`sheets_left`), its sheet as the last frame's end left it: each cell
    /// of a window a cell wider than its sheet's extent read (the reference's
    /// own cells and the extent's corners, for an extent too wide to read
    /// whole), and areas inside, across and past the extent walked, value
    /// for value in order, a walk answering the first reason it meets
    /// (KERNEL.8, slice b) wherever it stands; the twin's extent the
    /// rectangle of the reference's values (KERNEL.49); then each
    /// sheet named kept as an array, or not, as said, so that equal answers
    /// never hide a twin read another way.
    fn dense_is_the_twin(m: &Machine, last: &Left, kept_now: &[(&str, bool)], when: &str) {
        use crate::calc::{Area, Ctx};
        let ctx = Ctx::new(m, &m.library);
        for s in 0..m.sheets {
            let t = m.sheets + s;
            let name = m.book.pages[t].sheet.name.clone();
            assert_eq!(m.extents[t], extent_of(&last[s]), "{when}: {name}'s extent");
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
                cells.extend(last[s].keys().copied());
                cells.extend([(top, left), (top, right), (bottom, left), (bottom, right)]);
            }
            for (row, col) in cells {
                assert_eq!(
                    m.read((t, row, col)),
                    twin_cell(last, s, row, col),
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
                let walk = ctx.each(&[area], |v| {
                    walked.push(v);
                    Ok(())
                });
                let mut want = Vec::new();
                let mut stop = None;
                for (&(_, col), held) in last[s].range((area.top, 0)..=(area.bottom, u32::MAX)) {
                    if col < area.left || col > area.right {
                        continue;
                    }
                    match held {
                        Ok(Value::Empty) => {}
                        Ok(v) => want.push(v.clone()),
                        Err(r) => {
                            stop = Some(r.clone());
                            break;
                        }
                    }
                }
                assert_eq!(
                    (walked, walk.err()),
                    (want, stop),
                    "{when}: {name} {area:?}"
                );
            }
        }
        for (sheet, kept) in kept_now {
            assert_eq!(dense_kept(m, sheet), *kept, "{when}: {sheet} kept");
        }
    }

    #[test]
    fn a_twins_store_is_its_sheet_as_the_last_frame_left_it() {
        // KERNEL.49 (KERNEL.20's law): each twin, read for read and walked
        // for walked, its sheet's values as the last frame's end or the load
        // left them, held to the reference taken from the sheet at each end
        // (`sheets_left`), over texts, truth values, errors, a cell not
        // computed, a shared formula and a second sheet; after a load, steps,
        // writes of every kind, a refused write, a yield and a frame count
        // set; then over a save loaded again. A machine keeping a store for
        // the twins its formulas read alone, the others made from their
        // sheets, and one keeping every twin's.
        for every in [false, true] {
            let kept = [("Board", true), ("Other", every)];
            let mut m = load(PLANES);
            if every {
                m.every_twin = true;
                m.keep_twins();
            }
            let mut last = sheets_left(&m);
            dense_is_the_twin(&m, &last, &kept, "the load");
            for frame in 1..=3 {
                m.step(0);
                last = sheets_left(&m);
                dense_is_the_twin(&m, &last, &kept, &format!("frame {frame}"));
            }
            for row in [
                "(cell \"Board\" \"B2\" 9)",
                "(formula \"Board\" \"C2\" \"=B2+1\")",
                "(derived \"Board\" \"G2\" 3)",
                // A value cell become a formula not computed: its slot a
                // reason, the extent unchanged.
                "(formula \"Board\" \"E2\" \"=LABEL(B2:C2)\")",
                // A text changed in place, then the extent grown twice.
                "(cell \"Board\" \"D2\" \"a longer text than x\")",
                "(cell \"Board\" \"C5:D6\" 2)",
                "(cell \"Board\" \"I1\" \"far\")",
            ] {
                assert!(m.write(row).is_ok(), "{row}");
                // Until the next frame's end the twin is the last frame's.
                dense_is_the_twin(&m, &last, &kept, row);
                m.step(0);
                last = sheets_left(&m);
                dense_is_the_twin(&m, &last, &kept, row);
            }
            assert!(m
                .write("(cell \"Board\" \"B2\" 1)\n(cell \"Nowhere\" \"A1\" 1)")
                .is_err());
            dense_is_the_twin(&m, &last, &kept, "a refused write");
            assert!(!m.step(2).done);
            dense_is_the_twin(&m, &last, &kept, "a yield");
            while !m.step(2).done {}
            last = sheets_left(&m);
            dense_is_the_twin(&m, &last, &kept, "the yielded frame done");
            m.resume_at(50).unwrap();
            m.step(0);
            last = sheets_left(&m);
            dense_is_the_twin(&m, &last, &kept, "a frame count set");
            let again = load(&save(&m));
            let read = [("Board", true), ("Other", false)];
            dense_is_the_twin(&again, &sheets_left(&again), &read, "a save loaded");
        }
        // Life, whose rule reads its twin in two ranges and three cells, and
        // no formula the Clock's.
        let mut life = load(LIFE);
        let kept = [("Screen", true), ("Clock", false)];
        dense_is_the_twin(&life, &sheets_left(&life), &kept, "Life's load");
        for frame in 1..=4u64 {
            life_step(&mut life);
            let when = format!("Life's frame {frame}");
            dense_is_the_twin(&life, &sheets_left(&life), &kept, &when);
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
        // Edge's twin no formula reads, kept here all the same (KERNEL.49).
        m.every_twin = true;
        m.keep_twins();
        let kept = [("Big", false), ("Edge", true)];
        dense_is_the_twin(&m, &sheets_left(&m), &kept, "the load");
        m.step(0);
        assert_eq!(number(&m, "Big", "B1"), 3.0);
        m.step(0);
        assert_eq!(
            number(&m, "Big", "B1"),
            6.0,
            "the twin's B1 read in the range"
        );
        dense_is_the_twin(&m, &sheets_left(&m), &kept, "two steps");
        // Edge's extent is A1:D2, eight slots for three cells; a fourth
        // cell far right makes it twenty slots for four, past four a cell.
        let last = sheets_left(&m);
        assert_eq!(m.write("(cell \"Edge\" \"J2\" 1)"), Ok(1));
        dense_is_the_twin(&m, &last, &kept, "J2 written");
        m.step(0);
        let grown = [("Big", false), ("Edge", false)];
        dense_is_the_twin(&m, &sheets_left(&m), &grown, "Edge grown");
    }

    #[test]
    fn a_range_over_a_twin_meets_a_reason_where_the_sheet_does() {
        // The defect found 2026-10-10 (KERNEL.20), removed by KERNEL.8's
        // second slice: a twin holds no cell for a formula not computed, and
        // a range over the twin answers its reason as a range over the sheet
        // does: on the walk, in a run, in a program, and in the lookups' and
        // the criteria's walks; since KERNEL.49 the reason a slot's kind in
        // the twin's array, which a walk meets where it stands. The range
        // where one value is wanted that made the reason before now meets
        // its cell; a reason comes from a label of the first slice.
        let mut m = load(concat!(
            "(cell \"S\" \"A1:A3\" 1)\n",
            "(cell \"S\" \"B1\" 5)\n",
            "(formula \"S\" \"B2\" \"=XMATCH(1,A1:A3,3)\")\n",
            "(cell \"S\" \"B3\" 7)\n",
            "(formula \"S\" \"C1\" \"=SUM(S.last!B1:B3)\")\n",
            "(formula \"S\" \"C2\" \"=S.last!B2\")\n",
            "(formula \"S\" \"C3\" \"=SUM(B1:B3)\")\n",
            "(formula \"S\" \"C4\" \"=MATCH(7,S.last!B1:B3,0)\")\n",
            "(formula \"S\" \"C5\" \"=COUNTIF(S.last!B1:B3,7)\")\n",
            "(formula \"S\" \"C6\" \"=SUM(S.last!B1:B3)+1\")\n",
            "(formula \"S\" \"E1:E8\" \"=SUM(S.last!$B$1:$B$3)\")\n",
            "(formula \"S\" \"D1\" \"=A1:A3\")\n",
            "(formula \"S\" \"D2\" \"=A1:A3*2\")\n",
            "(formula \"S\" \"D4\" \"=A1:A3\")\n",
        ));
        m.step(0);
        m.step(0);
        assert!(dense_kept(&m, "S"), "the twin is kept as an array");
        let twin = m.find_sheet("S.last").unwrap();
        assert!(
            m.dense(twin).is_some(),
            "and offers it while a reason stands among its slots"
        );
        let reason = Some(Computed::NotComputed(Reason::Construct(
            "match mode 3".to_string(),
        )));
        assert_eq!(at(&m, "S", "B2"), reason);
        for addr in ["C1", "C2", "C3", "C4", "C5", "C6", "E1", "E8"] {
            assert_eq!(at(&m, "S", addr), reason, "{addr}");
        }
        assert_eq!(number(&m, "S", "D1"), 1.0, "A1:A3 met at row 1");
        assert_eq!(number(&m, "S", "D2"), 2.0, "and at row 2");
        assert_eq!(
            at(&m, "S", "D4"),
            Some(Computed::Value(Value::Error(ErrorKind::Value))),
            "no cell of A1:A3 in row 4"
        );
        // B2 computed: no reason stands, and the array is offered again.
        assert!(m.write("(formula \"S\" \"B2\" \"=8\")").is_ok());
        m.step(0);
        m.step(0);
        assert!(m.dense(twin).is_some());
        assert_eq!(number(&m, "S", "C1"), 20.0);
        assert_eq!(number(&m, "S", "C4"), 3.0);
        assert_eq!(number(&m, "S", "E8"), 20.0);
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
        let book = m.workbook();
        let before = Calc::run(&book, &m.library);
        let mut cells: Vec<CellId> = m.shapes.order_cells().to_vec();
        cells.sort_unstable();
        let want: Vec<Option<Computed>> = cells
            .iter()
            .map(|id| before.computed(*id).cloned())
            .collect();
        m.step(0);
        let got: Vec<Option<Computed>> = cells.iter().map(|id| m.computed(*id)).collect();
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
            "(formula \"S\" \"C2\" \"=SUM(LABEL(A1:A2)*1,LABEL(Z1:Z2)*1)\")\n",
            "(formula \"S\" \"C3\" \"=SUM(1,LABEL(A1:A2)*1,LABEL(Z1:Z2)*1)\")\n",
            "(formula \"S\" \"C4\" \"=SUM(1,2,LABEL(A1:A2)*1,LABEL(Z1:Z2)*1)\")\n",
            "(formula \"S\" \"C5\" \"=SUM(1,2,3,4,LABEL(A1:A2)*1,LABEL(Z1:Z2)*1)\")\n",
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
        let book = m.workbook();
        book.sheets[s]
            .cells
            .iter()
            .filter_map(|(&k, c)| {
                if !is_formula(&c.content) {
                    return Some((k, c.content.clone()));
                }
                let content = match m.computed((s, k.0, k.1)) {
                    Some(Computed::Value(Value::Number(n))) => Content::Number(n),
                    Some(Computed::Value(Value::Text(t))) => Content::Text(t),
                    Some(Computed::Value(Value::Bool(b))) => Content::Bool(b),
                    Some(Computed::Value(Value::Error(e))) => Content::Error(e.text().to_string()),
                    Some(Computed::Value(Value::Empty)) => Content::Number(0.0),
                    Some(Computed::NotComputed(_)) | None => return None,
                };
                Some((k, content))
            })
            .collect()
    }

    /// A twin's cells as a reader of the twin as a sheet reads them, made
    /// from its sheet (KERNEL.49), as a view makes them.
    fn twin_cells(m: &Machine, t: usize) -> BTreeMap<(u32, u32), Content> {
        let all = A1Range {
            top: 1,
            left: 1,
            bottom: MAX_ROW,
            right: MAX_COLUMN,
        };
        let mut out = BTreeMap::new();
        Cells::cells(m, t, &all, &mut |at, cell| match cell.holds {
            Holds::Constant(c) => {
                out.insert(at, c.content());
            }
            Holds::Formula(_) => panic!("a twin holds no formula"),
        });
        out
    }

    /// Every twin's cells, and its extent, as its sheet's values make them:
    /// held at a frame's end or a load, or while a step yields, before any
    /// write; the extent the rectangle of the cells (KERNEL.49).
    fn twins_are_their_sheets(m: &Machine, when: &str) {
        for (s, held) in sheets_left(m).iter().enumerate() {
            let t = m.sheets + s;
            let name = &m.book.pages[t].sheet.name;
            assert_eq!(twin_cells(m, t), twin_from(m, s), "{when}: {name}");
            assert_eq!(m.extents[t], extent_of(held), "{when}: {name}'s extent");
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
            "(formula \"Board\" \"E2\" \"=LABEL(B2:C2)\")",
            "(cell \"Board\" \"D2\" \"a longer text than x\")",
            "(cell \"Board\" \"C5:D6\" 2)",
            "(formula \"Board\" \"D6\" \"=LABEL(B2:C2)\")",
            // One cell losing its value and another gaining one in the same
            // write: as many valued cells as before, standing elsewhere.
            "(formula \"Board\" \"F2\" \"=LABEL(B2:C2)\")\n(cell \"Board\" \"H3\" 5)",
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

    /// A vector a lookup searches (KERNEL.8), one row or one column, and one
    /// time in eight any range `range` writes: a value column held whole,
    /// whose index every cell shares, or moving with the cell; a twin's
    /// column or row held whole, which the frames change; the sheet before's
    /// formulas held whole, which its frame changes before the band reads
    /// them; or the formula cells above the band.
    fn vector_of(d: &mut Dice, at: &Band) -> String {
        let col = d.pick(&["A", "B"]);
        let twin = d.pick(&["A", "B"]);
        let f = d.pick(&LETTERS[2..8]);
        let n = 1 + d.below(20) as u32;
        match d.below(8) {
            0 | 1 => format!("${col}$1:${col}${n}"),
            2 => {
                let r = 1 + d.below(at.rows as usize) as u32;
                format!("${col}{r}:${col}{}", r + d.below(4) as u32)
            }
            3 => format!("{twin}.last!${f}$1:${f}${}", at.rows),
            4 => format!("{twin}.last!$C${n}:$H${n}"),
            5 if at.before.is_some() => {
                format!("{}!${f}$1:${f}${}", at.before.unwrap_or("A"), at.rows)
            }
            6 if at.top > 1 => format!("${f}$1:${f}${}", at.top - 1),
            _ => range(d, at),
        }
    }

    /// A table `VLOOKUP` searches down its first column (KERNEL.8): the value
    /// columns, a twin's formulas, the sheet before's, or any range.
    fn table_of(d: &mut Dice, at: &Band) -> String {
        let twin = d.pick(&["A", "B"]);
        match d.below(4) {
            0 => "$A$1:$B$20".to_string(),
            1 => format!("{twin}.last!$C$1:$E${}", at.rows),
            2 if at.before.is_some() => format!("{}!$C$1:$F${}", at.before.unwrap_or("A"), at.rows),
            _ => range(d, at),
        }
    }

    /// What a lookup searches for: a value of each kind the value cells
    /// hold, a text in another case, a reference, or any formula.
    fn probe_of(d: &mut Dice, at: &Band, depth: u32) -> String {
        match d.below(4) {
            0 => d
                .pick(&[
                    "0", "1", "3", "-2.5", "\"5\"", "\"x\"", "\"X\"", "TRUE", "\"TRUE\"", "\"\"",
                    "\"x*\"", "\"?\"",
                ])
                .to_string(),
            1 => reference(d, at),
            _ => expr(d, at, depth + 1),
        }
    }

    /// A criterion of `COUNTIF` and `SUMIF`: by size, `<>`, a pattern, the
    /// blank, an error, or any probe, a cell's value among them.
    fn criterion_of(d: &mut Dice, at: &Band, depth: u32) -> String {
        match d.below(3) {
            0 => d
                .pick(&[
                    "\">1\"",
                    "\"<>x\"",
                    "\"x*\"",
                    "\"\"",
                    "\"<>\"",
                    "\"=5\"",
                    "\"<3\"",
                    "\"?\"",
                    "\"=TRUE\"",
                    "\"#N/A\"",
                    "\"<>3\"",
                ])
                .to_string(),
            _ => probe_of(d, at, depth),
        }
    }

    /// The lookups, the criteria and the error tests (KERNEL.8), each of
    /// every mode and search, over vectors held whole, whose index a band's
    /// cells share, or moving with the cell.
    fn lookup_expr(d: &mut Dice, at: &Band, depth: u32) -> String {
        let e = |d: &mut Dice| expr(d, at, depth + 1);
        let twin = d.pick(&["A", "B"]);
        match d.below(12) {
            0 => {
                let t = d.pick(&["", ",0", ",1", ",-1", ","]);
                format!("MATCH({},{}{t})", probe_of(d, at, depth), vector_of(d, at))
            }
            1 => {
                let mode = d.pick(&["", ",0", ",-1", ",1", ",2", ",3"]);
                let search = match mode {
                    "" => "",
                    _ => d.pick(&["", ",1", ",-1", ",2", ",-2"]),
                };
                format!(
                    "XMATCH({},{}{mode}{search})",
                    probe_of(d, at, depth),
                    vector_of(d, at)
                )
            }
            2 => {
                let col = d.pick(&["1", "2", "3", "0"]);
                let approx = d.pick(&["", ",FALSE", ",TRUE", ",0", ","]);
                format!(
                    "VLOOKUP({},{},{col}{approx})",
                    probe_of(d, at, depth),
                    table_of(d, at)
                )
            }
            3 => {
                let row = d.pick(&["1", "2", "3"]);
                let approx = d.pick(&["", ",FALSE", ",TRUE"]);
                format!(
                    "HLOOKUP({},{twin}.last!$C$1:$H$3,{row}{approx})",
                    probe_of(d, at, depth)
                )
            }
            4 => {
                let results = d.pick(&["", ",$B$1:$B$20", ",A.last!$D$1:$D$20", ",$B$1"]);
                format!(
                    "LOOKUP({},{}{results})",
                    probe_of(d, at, depth),
                    vector_of(d, at)
                )
            }
            5 => {
                let (vector, results) = match d.below(4) {
                    0 => ("$A$1:$A$20".to_string(), "$B$1:$B$20".to_string()),
                    1 => (
                        format!("{twin}.last!$C$1:$C${}", at.rows),
                        format!("{twin}.last!$D$1:$D${}", at.rows),
                    ),
                    2 => (
                        format!("{twin}.last!$C$1:$H$1"),
                        format!("{twin}.last!$C$2:$H$2"),
                    ),
                    _ => (vector_of(d, at), "$B$1:$B$20".to_string()),
                };
                let missing = d.pick(&["", ",\"none\"", ","]);
                let mode = match missing {
                    "" => "",
                    _ => d.pick(&["", ",0", ",-1", ",1", ",2"]),
                };
                let search = match mode {
                    "" => "",
                    _ => d.pick(&["", ",1", ",-1", ",2", ",-2"]),
                };
                let p = probe_of(d, at, depth);
                format!("XLOOKUP({p},{vector},{results}{missing}{mode}{search})")
            }
            // A row of a constant mostly, so that few cells turn from a value
            // to a pick of many and back as the twins change.
            6 => {
                let col = d.pick(&["", ",1", ",2", ",0", ",3"]);
                let row = match d.below(4) {
                    0 => e(d),
                    _ => d.pick(&["1", "2", "3", "0", "25"]).to_string(),
                };
                format!("INDEX({},{row}{col})", table_of(d, at))
            }
            7 => match d.below(3) {
                0 => format!("SUM(INDEX({},0,{}))", table_of(d, at), d.pick(&["1", "2"])),
                1 => format!("INDEX({},1)", reference(d, at)),
                _ => format!(
                    "({})+INDEX({},{},{})",
                    e(d),
                    range(d, at),
                    d.pick(&["1", "2", "0"]),
                    d.pick(&["1", "2", "MOD(ROW(),3)"])
                ),
            },
            8 => format!("COUNTIF({},{})", range(d, at), criterion_of(d, at, depth)),
            9 => {
                let sums = d.pick(&["", ",$B$1:$B$20", ",A.last!$C$1", ",$A$1:$A$3"]);
                format!(
                    "SUMIF({},{}{sums})",
                    range(d, at),
                    criterion_of(d, at, depth)
                )
            }
            10 => format!("{}({},{})", d.pick(&["IFERROR", "IFNA"]), e(d), e(d)),
            _ => format!("{}({})", d.pick(&["ISERROR", "ISNA"]), e(d)),
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
    /// beside and inside the run's other calls; and the lookups, the
    /// criteria and the error tests (KERNEL.8).
    fn expr(d: &mut Dice, at: &Band, depth: u32) -> String {
        if depth >= 3 || d.below(4) == 0 {
            return leaf(d, at);
        }
        if d.below(10) == 0 {
            return lookup_expr(d, at, depth);
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

    /// Every probe an index answers held to its scan while it lives, on this
    /// thread (KERNEL.8).
    struct Checked;

    impl Checked {
        fn on() -> Checked {
            crate::calc::index::check(true);
            Checked
        }
    }

    impl Drop for Checked {
        fn drop(&mut self) {
            crate::calc::index::check(false);
        }
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

    /// Every formula cell's value as its slot holds it (KERNEL.49), the last
    /// complete frame's after a done step, spelled to the bit (a negative
    /// zero, too).
    fn values(m: &Machine) -> BTreeMap<CellId, String> {
        m.shapes
            .order_cells()
            .iter()
            .filter_map(|&id| {
                let held = m.stores[id.0].slot_of(id.1, id.2)?.computed()?;
                Some((id, format!("{held:?}")))
            })
            .collect()
    }

    #[test]
    fn a_frame_is_the_same_with_its_order_listed_as_found_by_the_finder() {
        // KERNEL.49's fourth slice: an order of at most 262,144 cells kept
        // also as a list (`blocks::LIST`) is read off it, and the finder
        // walks the blocks of every larger one, and in the tests unless a
        // test asks for the list. Churning grids loaded twice, listed and
        // not, written and stepped alike in steps that yield, a value or a
        // formula written between frames, the formula's plan made again
        // listed and not: every step's report and every frame's values the
        // same.
        use crate::calc::blocks::list_at;
        let mut listed_cells = 0;
        for seed in 1..=60u64 {
            let mut d = Dice(seed ^ 0x4904_115D);
            let rows = churning_grid(&mut d);
            let was = list_at(usize::MAX);
            let mut listed = load(&rows);
            list_at(0);
            let mut found = load(&rows);
            list_at(was);
            assert_eq!(
                found.shapes.blocks().listed(),
                0,
                "seed {seed}: no list asked"
            );
            listed_cells += listed.shapes.blocks().listed();
            for frame in 1..=4 {
                let tick = format!("(cell \"Clock\" \"B1\" {frame})");
                assert!(listed.write(&tick).is_ok() && found.write(&tick).is_ok());
                loop {
                    let budget = 1 + d.below(200);
                    let (a, b) = (listed.step(budget), found.step(budget));
                    assert_eq!(a, b, "seed {seed}, frame {frame}");
                    if a.done {
                        break;
                    }
                }
                assert_eq!(
                    values(&listed),
                    values(&found),
                    "seed {seed}, frame {frame}\n{rows}"
                );
                let write = match d.below(3) {
                    0 => format!("(cell \"A\" \"A{}\" 7)", 1 + d.below(20)),
                    1 => format!(
                        "(formula \"A\" \"{}{}\" \"=1\")",
                        d.pick(&LETTERS[2..8]),
                        1 + d.below(4)
                    ),
                    _ => continue,
                };
                let was = list_at(usize::MAX);
                let a = listed.write(&write);
                list_at(0);
                let b = found.write(&write);
                list_at(was);
                assert_eq!(a.is_ok(), b.is_ok(), "seed {seed}: {write}");
                assert_eq!(
                    listed.shapes.blocks().listed(),
                    listed.shapes.len(),
                    "seed {seed}"
                );
            }
        }
        assert!(
            listed_cells > 1_000,
            "{listed_cells} cells listed: too few for the law to hold over"
        );
    }

    #[test]
    fn a_run_evaluated_at_once_is_its_cells_evaluated_one_by_one() {
        // KERNEL.20, slice 3: grids of generated formulas, every kind of node,
        // stepped by runs in slices of any budget, against the same grids
        // with every cell evaluated alone, frame for frame and bit for bit;
        // writes between frames change values, break a band's shape and make
        // the runs again.
        let _checked = Checked::on();
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
                let book = runs.workbook();
                let calc = crate::calc::Calc::run(&book, &runs.library);
                let spell = |id: &CellId| format!("{:?}", calc.computed(*id).expect("computed"));
                runs.shapes
                    .order_cells()
                    .iter()
                    .map(|id| (*id, spell(id)))
                    .collect()
            };
            for frame in 1..=4 {
                in_runs += runs.shapes.cells_in_runs();
                by_cell += runs.shapes.len();
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
    fn a_twin_no_formula_reads_keeps_no_store() {
        // KERNEL.49's decision 5 (KERNEL.43 kept such a twin's array without
        // the numbers beside its slots): a twin no formula reads keeps no
        // store, made from its sheet where it is read. A write that makes a
        // formula read the twin makes its store, an array read by slices in
        // the frames after it; a write that leaves no formula reading it gives
        // the store up. Every frame is the cells evaluated one by one.
        let rows = concat!(
            "(sheet \"U\" visible)\n",
            "(cell \"U\" \"C1:Z1\" 2)\n",
            "(formula \"U\" \"C2:Z2\" \"=U!C1+COLUMN()\")\n",
        );
        let (mut runs, mut alone) = (load(rows), load(rows));
        alone.by_runs = false;
        assert!(runs.step(0).done && alone.step(0).done);
        assert_eq!(store_kept(&runs, "U"), None, "no formula reads U.last");
        for m in [&mut runs, &mut alone] {
            assert_eq!(
                m.write("(formula \"U\" \"C3:Z3\" \"=U.last!C2*2\")"),
                Ok(24)
            );
        }
        assert_eq!(store_kept(&runs, "U"), Some(true), "made at the write");
        vector::sliced::take();
        for frame in 2..=3 {
            assert!(runs.step(0).done && alone.step(0).done);
            assert_eq!(values(&runs), values(&alone), "frame {frame}");
            assert_eq!(store_kept(&runs, "U"), Some(true), "frame {frame}");
        }
        assert!(
            vector::sliced::take()[vector::sliced::READ] > 0,
            "the twin read by slices"
        );
        for m in [&mut runs, &mut alone] {
            assert_eq!(m.write("(formula \"U\" \"C3:Z3\" \"=U!C2*2\")"), Ok(24));
        }
        assert_eq!(store_kept(&runs, "U"), None, "given up at the write");
        assert!(runs.step(0).done && alone.step(0).done);
        assert_eq!(values(&runs), values(&alone));
        assert_eq!(store_kept(&runs, "U"), None);
        // A shape whose reads the plan cannot follow, a name's, might read
        // any twin: while one stands, every twin keeps its store.
        let named = concat!(
            "(sheet \"U\" visible)\n",
            "(name \"Rate\" \"U!$C$1\")\n",
            "(cell \"U\" \"C1:Z1\" 2)\n",
            "(formula \"U\" \"C2:Z2\" \"=U!C1+COLUMN()\")\n",
            "(formula \"U\" \"C3\" \"=Rate*2\")\n",
        );
        let mut m = load(named);
        assert!(m.step(0).done);
        assert_eq!(store_kept(&m, "U"), Some(true), "a name's shape stands");
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
                let book = runs.workbook();
                let calc = crate::calc::Calc::run(&book, &runs.library);
                let spell = |id: &CellId| format!("{:?}", calc.computed(*id).expect("computed"));
                runs.shapes
                    .order_cells()
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

    /// A machine's census, its twins' arrays as it keeps them and its rows'
    /// sheets' stores as it offers them (KERNEL.31; KERNEL.49), as
    /// `Machine::census`.
    fn census_of(m: &Machine) -> crate::calc::census::Census {
        crate::calc::census::of(&m.shapes, &m.book, &m.library, &|s| {
            Env::dense(m, s).is_some()
        })
    }

    /// Each formula cell's reads and level found again as `Calc` would find
    /// them, none of the plan's code used: its text read off the grid (a
    /// shared formula's child as its master's text moved), its reads by
    /// `graph::dependencies`, the order by `graph::plan`; a cell's level 0
    /// when it reads no formula cell, else one past the deepest it reads.
    fn walked_again(m: &Machine) -> (BTreeMap<CellId, usize>, HashMap<CellId, Vec<CellId>>) {
        let book = &m.workbook();
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
        let order = m.shapes.order_cells();
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
                let book = m.workbook();
                for (si, sheet) in book.sheets.iter().enumerate() {
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
            Path::Keyed => f.reads.contains(&Kind::Keyed),
            Path::Call(name) => f.calls.contains(name),
        }
    }

    #[test]
    fn the_census_names_every_path_a_run_takes() {
        // KERNEL.30: over the generated grids, stepped in slices of any
        // budget with writes between frames, every path a run takes, noted
        // as it takes it, is a kind of read or a call the census names for
        // the run's shape, the census taken at the frame's start, when the
        // twins' arrays are the ones the run reads, a store keeping its
        // array while a formula of it holds a reason, a slot's kind
        // (KERNEL.49); and every kind of path is taken somewhere, so that the
        // law holds over each: every other grid's second sheet made too
        // sparse for an array by a cell far off, its cells and its twin's
        // read through lookups and its ranges through its cells.
        let mut kinds: BTreeSet<&str> = BTreeSet::new();
        for seed in 1..=80u64 {
            let mut d = Dice(seed);
            let mut rows = generated_grid(&mut d);
            if seed % 2 == 0 {
                rows.push_str("(cell \"B\" \"Z5000\" 1)\n");
            }
            let mut m = load(&rows);
            for frame in 1..=4 {
                let census = census_of(&m);
                let facts: BTreeMap<usize, &crate::calc::census::Facts> =
                    census.shapes.iter().map(|f| (f.shape, f)).collect();
                vector::taken::begin();
                while !m.step(1 + d.below(300)).done {}
                for (s, path) in vector::taken::end() {
                    let Some(f) = facts.get(&s) else {
                        panic!(
                            "seed {seed}, frame {frame}: a run of shape {s}, which no row names"
                        );
                    };
                    assert!(
                        named(f, &path),
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
                        vector::taken::Path::Keyed => "keyed",
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
            ["array", "call", "cells", "each", "keyed", "lookup", "once", "slot", "stride"],
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
        // through its cells; and `T`'s `S`, a vector a lookup searches, every
        // `$` holding it (KERNEL.8).
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
            "(formula \"T\" \"S1:S8\" \"=MATCH(1,$A$1:$A$8,0)\")\n",
        ));
        let want: [(&str, &str, &[&str]); 17] = [
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
            // A vector every `$` holds, which the run's cells search by the
            // index they share (KERNEL.8).
            ("S1", "keyed", &["MATCH"]),
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
            "(formula \"S\" \"X1\" \"=LABEL(A1:A2)\")\n",
            "(formula \"S\" \"C1:C8\" \"=LABEL(A1:A2)+LABEL(B1:B2)\")\n",
            "(formula \"S\" \"D1:D8\" \"=$X$1+LABEL(B1:B2)\")\n",
            "(formula \"S\" \"E1:E8\" \"=LABEL(B1:B2)+$X$1\")\n",
            "(formula \"S\" \"F1:F8\" \"=MOD(LABEL(A1:A2)+1,LABEL(B1:B2)+1)\")\n",
            "(formula \"S\" \"G1:G8\" \"=SUM(LABEL(A1:A2)+1,LABEL(B1:B2)+1,LABEL(A3:A4)+1)\")\n",
            "(formula \"S\" \"H1:H8\" \"=ABS(LABEL(B1:B2)*2)\")\n",
            "(formula \"S\" \"I1:I8\" \"=1+LABEL(A1:A2)-LABEL(B1:B2)\")\n",
            "(formula \"S\" \"J1:J8\" \"=IF(LABEL(A1:A2)=1,LABEL(B1:B2),0)\")\n",
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
            runs.shapes.len(),
            "every cell in a run"
        );
        let calc: BTreeMap<CellId, String> = {
            let book = runs.workbook();
            let calc = crate::calc::Calc::run(&book, &runs.library);
            let spell = |id: &CellId| format!("{:?}", calc.computed(*id).expect("computed"));
            runs.shapes
                .order_cells()
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

    #[test]
    fn a_run_meets_a_range_where_one_value_is_wanted_at_each_cell() {
        // KERNEL.8, slice b: implicit intersection by runs, by every cell
        // alone and by `Calc`, frame for frame: a column met at each cell's
        // row, inside it and past it, held and moving; a row met at each
        // cell's column and a block at both, on another sheet; a twin's
        // column; a name's range; `SINGLE`; `INDEX`'s pick of many cells; and
        // each as the operand of a node a run computes, a condition, a call,
        // a sign, `&` and `CHOOSE`'s value among them.
        let rows = concat!(
            "(cell \"S\" \"A1:A6\" 3)\n",
            "(cell \"S\" \"B1:C8\" 2)\n",
            "(cell \"T\" \"A1:Z1\" 5)\n",
            "(cell \"T\" \"A2:Z3\" 6)\n",
            "(name \"Col\" \"S!$A$1:$A$6\")\n",
        );
        let formulas: Vec<String> = [
            "$A$1:$A$6*2",
            "A1:A2+1",
            "T!$A$1:$Z$1",
            "T!$A$1:$Z$3",
            "S.last!$A$1:$A$6",
            "Col+ROW()",
            "SINGLE($A$1:$A$6)",
            "ROW(SINGLE($A$1:$A$8))",
            "INDEX($A$1:$C$8,0,1)",
            "IF($A$1:$A$6>2,$B$1:$B$8,0)",
            "SUM($A$1:$A$6,1)",
            "MOD($A$1:$A$6,2)",
            "-$A$1:$A$6",
            "$A$1:$A$6&\"x\"",
            "CHOOSE(1,$A$1:$A$6)",
        ]
        .iter()
        .map(|f| f.to_string())
        .collect();
        let m = by_runs_and_alone(&columns(rows, 8, &formulas), &[]);
        assert_eq!(number(&m, "S", "D3"), 6.0, "A3 doubled");
        assert_eq!(
            at(&m, "S", "D7"),
            Some(Computed::Value(Value::Error(ErrorKind::Value))),
            "row 7 is not in A1:A6"
        );
        assert_eq!(
            number(&m, "S", "E8"),
            1.0,
            "A8:A9 met at row 8, empty, plus 1"
        );
        assert_eq!(number(&m, "S", "F1"), 5.0, "T's row met at column F");
        assert_eq!(number(&m, "S", "K8"), 8.0, "ROW of SINGLE's cell");
    }

    #[test]
    fn the_date_functions_a_run_calls_are_the_library_s_at_every_cell() {
        // KERNEL.8, slice d: the date functions, which a run calls at each
        // cell, by runs, by every cell alone and by `Calc`, frame for frame,
        // over serials, a time, a truth value and a date a locale reads; and
        // arithmetic and its unary minus labelling that text at a step.
        let rows = concat!(
            "(cell \"S\" \"A1\" 46023)\n",
            "(cell \"S\" \"A2\" 46023.75)\n",
            "(cell \"S\" \"A3\" 60)\n",
            "(cell \"S\" \"A4\" true)\n",
            "(cell \"S\" \"A5\" \"1/1/2026\")\n",
            "(cell \"S\" \"A6\" 45351)\n",
            "(cell \"S\" \"B1:B6\" 2)\n",
        );
        let formulas: Vec<String> = [
            "YEAR(A1)",
            "MONTH(A1)",
            "DAY(A1)",
            "WEEKDAY(A1,$B1)",
            "DATE(2026,B1,ROW())",
            "EOMONTH(A1,B1)",
            "EDATE(A1,-B1)",
            "DAYS(A1,A6)",
            "NETWORKDAYS(A6,A1)",
            "WORKDAY(A1,B1*ROW())",
            "HOUR(A2)",
            "TIME(ROW(),B1,0)",
            "DATEDIF(A6,A1,\"M\")",
            "YEARFRAC(A6,A1,B1)",
            "WEEKNUM(A1,B1)",
            "A1+1",
            "-A1",
        ]
        .iter()
        .map(|f| f.to_string())
        .collect();
        let m = by_runs_and_alone(&columns(rows, 6, &formulas), &[]);
        assert_eq!(number(&m, "S", "D1"), 2026.0, "YEAR");
        assert_eq!(number(&m, "S", "F3"), 29.0, "DAY of the day that never was");
        assert_eq!(number(&m, "S", "N1"), 18.0, "HOUR of a time");
        let locale = Some(Computed::NotComputed(Reason::Construct(
            "a locale's number".to_string(),
        )));
        assert_eq!(at(&m, "S", "D5"), locale, "YEAR of a locale's date");
        assert_eq!(at(&m, "S", "S5"), locale, "+ over it");
        assert_eq!(at(&m, "S", "T5"), locale, "- over it");
        assert_eq!(number(&m, "S", "S4"), 2.0, "TRUE is 1 to +");
    }

    #[test]
    fn the_aggregates_by_criteria_a_run_calls_are_the_library_s_at_every_cell() {
        // KERNEL.8, slice e: the aggregates, `SUBTOTAL`, the criteria's
        // family and the database functions, which a run calls at each
        // cell, by runs, by every cell alone and by `Calc`, frame for frame,
        // over a table under names whose fourth row a `row` row hides; and
        // a `SUBTOTAL` over a column of `SUBTOTAL`s, which a formula written
        // over a range makes a master and its shared children, passing each
        // by.
        let rows = concat!(
            "(cell \"S\" \"A1\" \"Account\")\n",
            "(cell \"S\" \"B1\" \"Amount\")\n",
            "(cell \"S\" \"C1\" \"Month\")\n",
            "(cell \"S\" \"A2\" \"Cash\")\n",
            "(cell \"S\" \"A3\" \"Bank\")\n",
            "(cell \"S\" \"A4\" \"Cash\")\n",
            "(cell \"S\" \"A5\" \"Card\")\n",
            "(cell \"S\" \"A6\" \"Bank\")\n",
            "(cell \"S\" \"B2\" 10)\n",
            "(cell \"S\" \"B3\" 20)\n",
            "(cell \"S\" \"B4\" 30)\n",
            "(cell \"S\" \"B5\" 40)\n",
            "(cell \"S\" \"B6\" 50)\n",
            "(cell \"S\" \"C2:C3\" 1)\n",
            "(cell \"S\" \"C4:C5\" 2)\n",
            "(cell \"S\" \"C6\" 3)\n",
            "(cell \"S\" \"Z1\" \"Account\")\n",
            "(cell \"S\" \"Z2\" \"Cash\")\n",
            "(row \"S\" 4 none hidden)\n",
        );
        let formulas: Vec<String> = [
            "SUMIFS($B$1:$B$6,$A$1:$A$6,A1)",
            "COUNTIFS($A$1:$A$6,A1,$C$1:$C$6,\">=\"&C1)",
            "AVERAGEIF($A$1:$A$6,A1,$B$1:$B$6)",
            "AVERAGEIFS($B$1:$B$6,$A$1:$A$6,A1)",
            "MAXIFS($B$1:$B$6,$A$1:$A$6,A1)",
            "MINIFS($B$1:$B$6,$C$1:$C$6,C1)",
            "COUNT($B$1:B1)",
            "COUNTA($A$1:A1)",
            "PRODUCT(C1,ROW())",
            "AVERAGE($B$1:B1)",
            "STDEV($B$1:B1)",
            "VAR.P($B$1:B1)",
            "SUBTOTAL(109,$B$1:B1)",
            "SUBTOTAL(9,$B$1:B1)",
            "DSUM($A$1:$C$6,\"Amount\",$Z$1:$Z$2)",
            "DGET($A$1:$C$6,2,$Z$1:$Z$2)",
            "SUBTOTAL(109,$P$1:P1)",
        ]
        .iter()
        .map(|f| f.to_string())
        .collect();
        let m = by_runs_and_alone(&columns(rows, 6, &formulas), &[]);
        assert_eq!(number(&m, "S", "D2"), 40.0, "SUMIFS of Cash");
        assert_eq!(number(&m, "S", "P4"), 30.0, "SUBTOTAL 109 passes row 4 by");
        assert_eq!(number(&m, "S", "Q3"), 30.0, "SUBTOTAL 9 above row 4");
        let filtered = Some(Computed::NotComputed(Reason::Construct(
            "a row a filter may hide".to_string(),
        )));
        assert_eq!(at(&m, "S", "Q4"), filtered, "SUBTOTAL 9 over row 4");
        assert_eq!(number(&m, "S", "R2"), 40.0, "DSUM counts a hidden row");
        assert_eq!(
            at(&m, "S", "S2"),
            Some(Computed::Value(Value::Error(ErrorKind::Num))),
            "DGET of two records"
        );
        assert_eq!(
            number(&m, "S", "T6"),
            0.0,
            "SUBTOTAL passes a SUBTOTAL's shared children by"
        );
    }

    #[test]
    fn the_text_functions_a_run_calls_are_the_library_s_at_every_cell() {
        // KERNEL.8, slice c: the text functions, which a run calls at each
        // cell, by runs, by every cell alone and by `Calc`, frame for frame,
        // over texts, a number, a truth value and an emoji, counted by
        // characters; `&` cut to 32,767 units, and through a pair labelled
        // at a step as at a load; and a comparison in Excel's order of
        // texts, labelled where a character beyond its table decides.
        let rows = concat!(
            "(cell \"S\" \"A1\" \"Hello\")\n",
            "(cell \"S\" \"A2\" \"a\u{1F600}b\")\n",
            "(cell \"S\" \"A3\" 12.5)\n",
            "(cell \"S\" \"A4\" true)\n",
            "(cell \"S\" \"A5\" \"  x  y \")\n",
            "(cell \"S\" \"A6\" \"\u{C9}CLAIR\")\n",
            "(cell \"S\" \"B1:B6\" 2)\n",
        );
        let formulas: Vec<String> = [
            "LEN(A1)",
            "LEFT(A1,$B1)",
            "MID(A1,2,B1)",
            "UPPER(A1)&\"!\"",
            "TRIM(A1)",
            "SUBSTITUTE(A1,\"l\",\"L\")",
            "FIND(\"l\",A1&\"l\")",
            "TEXT(B1+ROW(),\"0.00\")",
            "TEXTJOIN(\",\",TRUE,A1:B1)",
            "CONCATENATE(A1,B1)",
            "REPT(\"x\",B1)",
            "VALUE(B1&\"5\")",
            "EXACT(A1,\"Hello\")",
            "LOWER(A1)",
            "LEN(REPT(\"x\",32765)&A2)",
            "A1<\"b\"",
            "A1>\"\u{D8}\"",
        ]
        .iter()
        .map(|f| f.to_string())
        .collect();
        let m = by_runs_and_alone(&columns(rows, 6, &formulas), &[]);
        assert_eq!(number(&m, "S", "D1"), 5.0, "LEN of Hello");
        assert_eq!(number(&m, "S", "D2"), 3.0, "an emoji is one character");
        assert_eq!(
            at(&m, "S", "E2"),
            Some(Computed::Value(Value::Text("a\u{1F600}".to_string()))),
            "LEFT by characters"
        );
        assert_eq!(
            at(&m, "S", "K1"),
            Some(Computed::Value(Value::Text("3.00".to_string()))),
            "TEXT of 2 and the row"
        );
        assert_eq!(
            at(&m, "S", "R1"),
            Some(Computed::NotComputed(Reason::Construct(
                "half a surrogate pair".to_string()
            ))),
            "& through the pair at the 32,767th unit"
        );
        assert_eq!(number(&m, "S", "R2"), 32_767.0, "& cut to 32,767 units");
        let truth = |cell: &str| match at(&m, "S", cell) {
            Some(Computed::Value(Value::Bool(b))) => b,
            other => panic!("{cell}: {other:?}"),
        };
        assert!(!truth("S1"), "Hello is not below b");
        assert!(
            truth("S2"),
            "a text's first letter decides before its emoji"
        );
        assert_eq!(
            at(&m, "S", "T1"),
            Some(Computed::NotComputed(Reason::Construct(
                "a character's order".to_string()
            ))),
            "Ø's place is not known"
        );
        assert!(!truth("T3"), "a number is below every text");
        assert!(truth("T4"), "a truth value is above every text");
    }

    #[test]
    fn a_grid_reading_a_name_whose_text_is_not_read_is_refused() {
        // KERNEL.8, slice b: such a name, refused at the load with the
        // constructs, so that no cell holds its reason at a step; a name read,
        // and one no formula reads, refused by nothing.
        let e = refused(concat!(
            "(cell \"S\" \"A1\" 1)\n",
            "(name \"Pair\" \"A1#\")\n",
            "(formula \"S\" \"B1\" \"=Pair*2\")\n",
        ));
        assert_eq!(e.refusal.id, "calc-construct-not-read");
        assert!(e.refusal.text.contains("S!B1"), "{}", e.refusal.text);
        let m = load(concat!(
            "(cell \"S\" \"A1\" 1)\n",
            "(name \"Pair\" \"A1#\")\n",
            "(name \"Rate\" \"S!$A$1\")\n",
            "(formula \"S\" \"B1\" \"=Rate*2\")\n",
        ));
        assert_eq!(m.formulas(), (1, 1));
    }

    #[test]
    fn a_draft_s_dynamic_array_formula_keeps_its_range_an_array() {
        // KERNEL.8, slice b: a draft made by hand, as from a reader's
        // workbook, may hold a dynamic-array formula: the machine asks the
        // kind of the cell it computes, as `Calc` does, so that the formula's
        // range where one value is wanted is an array, its answer of many a
        // spill, labelled at its cell (slice f; slice b labelled the range as
        // written), each kind on a shape of its own; and the same text
        // written legacy is met at its cell; `SINGLE` meets the range in
        // either kind.
        let text = concat!(
            "(cell \"S\" \"A1:A6\" 3)\n",
            "(formula \"S\" \"B3\" \"=$A$1:$A$6*2\")\n",
            "(formula \"S\" \"B7\" \"=$A$1:$A$6*2\")\n",
            "(formula \"S\" \"C3\" \"=$A$1:$A$6*2\")\n",
            "(formula \"S\" \"C7\" \"=$A$1:$A$6*2\")\n",
            "(formula \"S\" \"D3\" \"=SINGLE($A$1:$A$6)*2\")\n",
            "(formula \"S\" \"E3\" \"=SUM(ABS($A$1:$A$6))\")\n",
            "(formula \"S\" \"F3\" \"=INDEX($A$1:$B$6,2,0)\")\n",
        );
        let mut draft = read_rows(&read_forms(text).unwrap()).unwrap();
        let s = draft.find_sheet("S").unwrap();
        for (row, col) in [(3, 3), (7, 3), (3, 4), (3, 5), (3, 6)] {
            draft.set_kind(s, row, col, crate::grid::Kind::Dynamic);
        }
        let cells = ["B3", "B7", "C3", "C7", "D3", "E3", "F3"];
        let library = labelling();
        let book = draft.workbook();
        let calc: Vec<String> = {
            let calc = crate::calc::Calc::run(&book, &library);
            cells
                .iter()
                .map(|addr| {
                    let r = parse_a1_range(addr).unwrap();
                    format!("{:?}", calc.computed((s, r.top, r.left)).cloned())
                })
                .collect()
        };
        let mut m = Machine::new(draft, labelling())
            .unwrap_or_else(|e| panic!("line {}: {}", e.line, e.refusal.text));
        assert!(m.step(0).done);
        assert_eq!(number(&m, "S", "B3"), 6.0, "legacy: A3 doubled");
        assert_eq!(
            at(&m, "S", "B7"),
            Some(Computed::Value(Value::Error(ErrorKind::Value))),
            "legacy: row 7 is not in A1:A6"
        );
        for addr in ["C3", "C7"] {
            assert_eq!(
                at(&m, "S", addr),
                Some(Computed::NotComputed(Reason::Construct(
                    "a spill".to_string()
                ))),
                "dynamic {addr}: an array, a spill labelled"
            );
        }
        assert_eq!(number(&m, "S", "D3"), 6.0, "SINGLE in a dynamic formula");
        // A function of one value over a range in a dynamic formula lifts
        // over the range read whole, where a program lowered for a legacy
        // formula would meet it at its cell (KERNEL.8, slice f).
        assert_eq!(number(&m, "S", "E3"), 18.0, "dynamic: ABS over A1:A6");
        // `INDEX`'s whole row read as a dynamic formula answers a range of
        // two cells, labelled by its text as slice b labels one (a program
        // lowered for a legacy formula leaves it to `Ctx` by its guard).
        assert_eq!(
            at(&m, "S", "F3"),
            Some(Computed::NotComputed(Reason::Construct(
                "$A$1:$B$6".to_string()
            ))),
            "dynamic: INDEX's row 2 of A1:B6"
        );
        let machine: Vec<String> = cells
            .iter()
            .map(|addr| format!("{:?}", at(&m, "S", addr)))
            .collect();
        assert_eq!(machine, calc, "the machine is Calc, cell for cell");
    }

    #[test]
    fn a_grid_holding_arrays_computes_each_cell_by_ctx_as_calc_does() {
        // KERNEL.8, slice f: a shape whose parse holds an array (a constant,
        // `:` between expressions, an array function, an array parameter
        // given an expression) is computed at each of its cells by `Ctx`,
        // a run's among them, its program never lowered; every cell as
        // `Calc` computes it, at the load and after a write a range's reach
        // and an array parameter's references follow.
        let text = concat!(
            "(cell \"S\" \"A1\" 1)\n",
            "(cell \"S\" \"A2\" 2)\n",
            "(cell \"S\" \"A3\" 3)\n",
            "(cell \"S\" \"A4\" 4)\n",
            "(cell \"S\" \"A5\" 5)\n",
            "(cell \"S\" \"A6\" 6)\n",
            "(formula \"S\" \"B1:B6\" \"=SUMPRODUCT($A$1:$A$6*(ROW($A$1:$A$6)<=ROW()))\")\n",
            "(formula \"S\" \"C1\" \"=SUM({1,2;3,4})\")\n",
            "(formula \"S\" \"C2\" \"=SUM(A1:INDEX(A1:A6,3))\")\n",
            "(formula \"S\" \"C3\" \"=INDEX(SORT(A1:A6,1,-1),1)\")\n",
            "(formula \"S\" \"C4\" \"=ROWS(FILTER(A1:A6,A1:A6>2))\")\n",
            "(formula \"S\" \"C5\" \"={10,20}+C1\")\n",
            "(formula \"S\" \"C6\" \"=SUM(INDEX(A1:A6,2):A4)*2\")\n",
            // Runs whose cells each hold an array, a constant's and an
            // array function's, which no lane holds whole.
            "(formula \"S\" \"D1:D6\" \"=SUM({1,2}*A1)\")\n",
            "(formula \"S\" \"E1:E6\" \"=SUM(SEQUENCE(2))*A1\")\n",
        );
        let cells = [
            "B1", "B2", "B3", "B4", "B5", "B6", "C1", "C2", "C3", "C4", "C5", "C6", "D1", "D6",
            "E1", "E6",
        ];
        let mut m = load(text);
        let held = |m: &mut Machine| {
            assert!(m.step(0).done);
            let library = labelling();
            let book = m.workbook();
            let calc = crate::calc::Calc::run(&book, &library);
            let s = book.find_sheet("S").unwrap();
            for addr in cells {
                let r = parse_a1_range(addr).unwrap();
                assert_eq!(
                    at(m, "S", addr),
                    calc.computed((s, r.top, r.left)).cloned(),
                    "{addr}: the machine is Calc"
                );
            }
        };
        held(&mut m);
        assert_eq!(number(&m, "S", "B6"), 21.0, "a running sum by arrays");
        assert_eq!(number(&m, "S", "C1"), 10.0);
        assert_eq!(number(&m, "S", "C2"), 6.0, "A1:A3");
        assert_eq!(number(&m, "S", "C3"), 6.0);
        assert_eq!(number(&m, "S", "C4"), 4.0);
        assert_eq!(number(&m, "S", "C5"), 20.0, "the first value of {{20,30}}");
        assert_eq!(number(&m, "S", "C6"), 18.0, "A2:A4 doubled");
        assert_eq!(number(&m, "S", "D6"), 18.0, "{{1,2}} times A6, summed");
        assert_eq!(number(&m, "S", "E6"), 18.0, "SEQUENCE(2) summed, times A6");
        m.write("(cell \"S\" \"A3\" 30)").unwrap();
        held(&mut m);
        assert_eq!(number(&m, "S", "B6"), 48.0);
        assert_eq!(number(&m, "S", "C2"), 33.0, "the range's reach follows A3");
        assert_eq!(number(&m, "S", "C3"), 30.0);
        assert_eq!(number(&m, "S", "C6"), 72.0);
        // A range whose ends move apart, one held by `$` and one not, is
        // bounded at its shape's first cell alone; its other cells are
        // labelled, since their reach is not the first one's moved.
        let m = load(concat!(
            "(cell \"S\" \"A1:C9\" 1)\n",
            "(formula \"S\" \"E1:E3\" \"=SUM($A$1:INDEX(C1:C3,2))\")\n",
        ));
        let mut m = m;
        assert!(m.step(0).done);
        assert_eq!(number(&m, "S", "E1"), 6.0, "A1:C2 at the first cell");
        for addr in ["E2", "E3"] {
            assert_eq!(
                at(&m, "S", addr),
                Some(Computed::NotComputed(Reason::Construct(
                    "a range whose ends move apart".to_string()
                ))),
                "{addr}"
            );
        }
    }

    #[test]
    fn a_run_compares_numbers_to_fifteen_digits_as_a_cell_alone_does() {
        // KERNEL.8, Excel's save of the lookups' fixture: a run's lanes of
        // numbers compared with one number, on either side of the operator,
        // and with another lane, by every operator, over numbers equal to
        // fifteen digits and not to the bit, just past the fifteenth digit
        // and far apart, of both signs, both zeros, the smallest and the
        // largest: each answer the one a cell evaluated alone gives, and
        // `Calc`, which compare two numbers one pair at a time.
        let numbers = [
            0.3,
            0.1 + 0.2,
            0.3 + 4e-16,
            0.3 - 4e-16,
            0.3 + 6e-16,
            0.3 - 1e-15,
            0.300_000_000_000_1,
            3.0,
            3.0 + 4e-15,
            3.0 - 4e-15,
            2.999_999_999_99,
            -0.3,
            -(0.1 + 0.2),
            0.0,
            -0.0,
            1e15,
            1e15 + 1.0,
            1e300,
            5e-324,
            1e-309,
            2.0,
        ];
        let mut values = String::new();
        for (k, n) in numbers.iter().enumerate() {
            let m = numbers[(k + 1) % numbers.len()];
            values.push_str(&format!("(cell \"S\" \"A{}\" {n})\n", k + 1));
            values.push_str(&format!("(cell \"S\" \"B{}\" {m})\n", k + 1));
        }
        let mut formulas = Vec::new();
        for op in ["=", "<>", "<", ">", "<=", ">="] {
            for c in ["0.3", "3", "-0.3", "0", "1000000000000000"] {
                formulas.push(format!("S.last!A1{op}{c}"));
                formulas.push(format!("{c}{op}S.last!A1"));
            }
            formulas.push(format!("S.last!A1{op}S.last!B1"));
            formulas.push(format!("S.last!A1{op}(0.1+0.2)"));
        }
        let rows = numbers.len() as u32;
        let m = by_runs_and_alone(&columns(&values, rows, &formulas), &[]);
        // `S.last!A1=0.3`, the first column, at the row holding 0.1+0.2, and
        // `S.last!A1<0.3`, the twenty-fifth, at the rows holding 0.3 less
        // 4E-16, equal to fifteen digits, and 0.3 less 1E-15, not.
        let truth = |addr: &str| match at(&m, "S", addr) {
            Some(Computed::Value(Value::Bool(b))) => b,
            other => panic!("S!{addr}: {other:?}"),
        };
        assert!(truth("D2"));
        assert!(!truth("AB4"));
        assert!(truth("AB6"));
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
            "(formula \"S\" \"C1:C8\" \"=LABEL(XX1:XX2)\")\n",
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
        let at_b5 = m.shapes.position((0, 5, 2));
        assert_eq!(m.write("(formula \"S\" \"B4\" \"=D1*2\")"), Ok(1));
        assert_ne!(m.shapes.position((0, 5, 2)), at_b5, "B5 moved");
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
        (m.sheets..m.book.pages.len())
            .map(|t| {
                let name = m.sheet_name(t).expect("a twin");
                let record = grid(m, name, None);
                // A twin holds no formula, so its values are its record,
                // its cells left behind their slots read alike (KERNEL.37).
                assert_eq!(values_view(m, name, None), record, "{name}'s values");
                let window = match m.extents[t] {
                    Some(((top, left), (bottom, right))) => A1Range {
                        top,
                        left,
                        bottom,
                        right,
                    }
                    .text(),
                    None => "A1".to_string(),
                };
                (record, plane(m, name, &window))
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
    fn a_twin_made_from_its_sheet_is_the_twin_kept() {
        // KERNEL.49's decision 5: grids keeping a store for the twins their
        // formulas read alone, every other twin made from its sheet where it
        // is read, against the same grids keeping every twin's store, frame
        // for frame: every twin's record and plane as a view reads them,
        // after a write and while a step yields as well as at a frame's end;
        // every value, twin cell, reason and extent, kept store and plane; the
        // grids lent whole; frames whose ends make every store and plane
        // again. Over grids of generated formulas, most of whose cells change
        // every frame, beside a sheet of values, of formulas reading the
        // first sheet and of reasons, which no formula reads through its twin
        // unless one of the grid's shapes is not followed, and which writes
        // change between frames; and over Life from soups of other seeds, the
        // Clock written before every step, whose Clock no formula reads
        // through its twin. Frames with a twin made from its sheet, and views
        // of one after a write and during a yield, each seen in both.
        let beside = concat!(
            "(sheet \"C\" visible)\n(cell \"C\" \"A1:A4\" 1)\n",
            "(formula \"C\" \"B1:B4\" \"=A!C1+A1\")\n",
            "(formula \"C\" \"C1\" \"=SUM(A!C1:H4)*2\")\n",
            "(formula \"C\" \"D1:D2\" \"=LABEL(A1:A2)\")\n",
        );
        let mut seen = [(0, 0, 0); 2];
        for seed in 1..=80u64 {
            let mut d = Dice(seed ^ 0x7A1C);
            let rows = generated_grid(&mut d) + beside;
            made_against_kept(&rows, &mut d, 8, &mut seen[0], |d, frame| {
                let mut w = write_between(d, frame).map_or(String::new(), |w| w + "\n");
                w.push_str(&match frame {
                    5 => "(formula \"C\" \"E1:E3\" \"=A1*3\")".to_string(),
                    _ => format!("(cell \"C\" \"A{}\" {frame})", 1 + d.below(4)),
                });
                Some(w)
            });
        }
        for seed in 1..=12u64 {
            let mut d = Dice(seed ^ 0x11FE);
            let soup = format!("ROW()*{}+COLUMN()*{}", 3 + d.below(20), 5 + d.below(20));
            let rows = LIFE.replace("ROW()*7+COLUMN()*13", &soup);
            made_against_kept(&rows, &mut d, 16, &mut seen[1], |_, frame| {
                Some(format!("(cell \"Clock\" \"B1\" {frame})"))
            });
        }
        for ((made, written, yielded), least) in
            seen.into_iter().zip([(80, 80, 60), (150, 150, 150)])
        {
            assert!(
                made >= least.0 && written >= least.1 && yielded >= least.2,
                "{made} frames with a twin made from its sheet, {written} views of one after a \
                 write and {yielded} during a yield: too few for the law to hold over"
            );
        }
    }

    /// One grid stepped keeping a store for the twins its formulas read and
    /// keeping every twin's, `write` before each frame, held to each other as
    /// `a_twin_made_from_its_sheet_is_the_twin_kept` says; `seen` counts its
    /// frames with a twin made from its sheet, and views of one after a write
    /// and during a yield.
    fn made_against_kept(
        rows: &str,
        d: &mut Dice,
        frames: u32,
        seen: &mut (usize, usize, usize),
        mut write: impl FnMut(&mut Dice, u32) -> Option<String>,
    ) {
        let mut made = load(rows);
        let mut kept = load(rows);
        kept.every_twin = true;
        kept.keep_twins();
        let first = made.sheet_name(0).expect("a sheet").to_string();
        for m in [&made, &kept] {
            m.view("plane", &first, None)
                .expect("a plane of the first sheet");
        }
        for frame in 1..=frames {
            let when = format!("frame {frame}\n{rows}");
            let unkept = made.twins.iter().any(Option::is_none);
            if let Some(w) = write(d, frame) {
                let (a, b) = (made.write(&w), kept.write(&w));
                assert_eq!(a.is_ok(), b.is_ok(), "{when}: {w}: {a:?} {b:?}");
                if a.is_ok() && unkept {
                    let views = twin_views(&made);
                    assert_eq!(views, twin_views(&kept), "{when}: a view after {w}");
                    same_state(&made, &kept, &format!("{when}: after {w}"));
                    seen.1 += 1;
                }
            }
            let budget = 1 + d.below(200);
            if !made.step(budget).done {
                let views = twin_views(&made);
                assert_eq!(views, twin_views(&kept), "{when}: a view during a yield");
                seen.2 += usize::from(unkept);
                while !made.step(budget).done {}
            }
            assert!(kept.step(0).done);
            seen.0 += usize::from(made.twins.iter().any(Option::is_none));
            assert_eq!(
                twin_views(&made),
                twin_views(&kept),
                "{when}: the twins' views"
            );
            same_state(&made, &kept, &when);
            // A frame's end that makes every store and plane again; or the
            // grids lent whole.
            let again = d.below(4) == 0;
            for m in [&mut made, &mut kept] {
                m.copies = !again;
            }
            if d.below(4) == 0 {
                assert_eq!(
                    format!("{:?}", made.workbook().sheets),
                    format!("{:?}", kept.workbook().sheets),
                    "{when}: the grids lent"
                );
            }
        }
    }
    #[test]
    fn a_value_cell_s_twin_is_made_as_written_beside_formulas_made_from_their_slots() {
        // KERNEL.27; KERNEL.49: a twin's cell is made from its sheet where it
        // is read, a formula's from its slot, which holds its value exactly,
        // and a value cell's as written, an error the language does not name
        // spelled as written, which the twin's store reads as `#VALUE!`; the
        // same in the record, in a view of one cell and in the grid lent. A
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
        let twin = m.find_sheet("S.last").unwrap();
        assert_eq!(store_kept(&m, "S"), Some(true), "B reads S.last");
        assert_eq!(
            m.read((twin, 3, 1)),
            Ok(Value::Error(ErrorKind::Value)),
            "the store reads its kind"
        );
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
        assert_eq!(grid(&m, "S.last", None), record, "the record, again");
        // A formula written over with a value: its twin made as written once
        // the next frame's end copies it.
        m.step(0);
        assert_eq!(m.write("(cell \"S\" \"B5\" (error \"#BAR!\"))"), Ok(1));
        assert_eq!(
            content(&m, "S.last", "B5"),
            Some(Content::Number(5.0)),
            "B5 as the last frame left it"
        );
        m.step(0);
        let bar = Content::Error("#BAR!".to_string());
        assert_eq!(content(&m, "S.last", "B5").as_ref(), Some(&bar), "B5");
        assert_eq!(
            m.workbook().sheets[1].cells[&(5, 2)].content,
            bar,
            "B5, made"
        );
    }

    /// Two grids' kept state the same to the bit: every formula's value of
    /// the last complete frame, every twin's cells, reasons and extent as
    /// they are made from its sheet, every twin's store that both keep, slot
    /// for slot, and every kept plane (KERNEL.20, slice 4; KERNEL.49); and
    /// every twin's store either keeps its twin made from its sheet.
    fn same_state(a: &Machine, b: &Machine, when: &str) {
        assert_eq!(values(a), values(b), "{when}: the values");
        for s in 0..a.sheets {
            let t = a.sheets + s;
            let name = &a.book.pages[t].sheet.name;
            assert_eq!(twin_holds(a, s), twin_holds(b, s), "{when}: {name}");
            assert_eq!(a.extents[t], b.extents[t], "{when}: {name}'s extent");
            if let (Some(x), Some(y)) = (&a.twins[s], &b.twins[s]) {
                assert!(x.same(y), "{when}: {name}'s store");
            }
        }
        kept_is_made(a, when);
        kept_is_made(b, when);
        assert_eq!(a.planes, b.planes, "{when}: the planes");
    }

    /// What each twin holds at each cell of its sheet, as its views and its
    /// reads make it (KERNEL.49), spelled to the bit.
    fn twin_holds(m: &Machine, s: usize) -> String {
        let book = m.workbook();
        let held: Vec<((u32, u32), Held)> = book.sheets[s]
            .cells
            .keys()
            .map(|&(row, col)| ((row, col), m.held(s, row, col)))
            .collect();
        format!("{held:?}")
    }

    /// Every twin's store a machine keeps, slot for slot its twin made from
    /// its sheet as the last frame's end left it (KERNEL.49): the stores a
    /// frame's end copies into, held to the store made again.
    fn kept_is_made(m: &Machine, when: &str) {
        for s in 0..m.sheets {
            if let Some(kept) = &m.twins[s] {
                let name = &m.book.pages[m.sheets + s].sheet.name;
                assert!(kept.same(&m.twin_made(s)), "{when}: {name}'s store made");
            }
        }
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
        let _checked = Checked::on();
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
                    let of = inc.shapes.len();
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
        // The first frame of every grid computes every cell, as do those
        // after a write that changes a formula and after a twin loses a
        // cell, which a generated formula's reason often makes; a cell added
        // and a twin cell gained are followed (KERNEL.56).
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
        // The twin cells the first frame's values gained are followed as any
        // change is (KERNEL.56), where a twin's gain made the second frame
        // compute every cell.
        assert_eq!(
            ticked(&mut m),
            11,
            "the second, each D reading its B's twin cell gained, and the name's"
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
        for id in m.shapes.order_cells() {
            let p = m.shapes.place(id).expect("placed");
            let Ok(e) = &m.shapes.shape(p.shape).expect("a shape").expr else {
                continue;
            };
            e.each_read(&mut |r| {
                let crate::calc::formula::Read::Ref(re) = r else {
                    return;
                };
                let Ok(areas) = crate::calc::graph::areas_found(
                    |n| m.book.find_sheet(n),
                    id.0,
                    re,
                    p.d_row,
                    p.d_col,
                ) else {
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
        let order = m.shapes.order_cells();
        let mut out = BTreeSet::new();
        m.shapes.reached(changed, &mut |r| match r {
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
                "(cell \"S\" \"A1\" 1)\n{}(formula \"S\" \"C1\" \"=IF(A1=1,LABEL(D1:D2),LABEL(E1:E2))\")\n\
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
    fn a_reason_s_label_is_moved_with_its_cell_and_owned_through_a_name() {
        // KERNEL.25: the label is borrowed from the parse, moved with a cell
        // at an offset from its shape's first, and owned when read through
        // a name, whose reference does not move; made by the tests' `LABEL`
        // since KERNEL.8's second slice met a range where one value is wanted
        // at its cell.
        let mut m = load(concat!(
            "(cell \"S\" \"A1:B3\" 1)\n",
            "(name \"Wide\" \"S!$A$1:$B$1\")\n",
            "(formula \"S\" \"D1:D2\" \"=LABEL(A1:B1)*2\")\n",
            "(formula \"S\" \"E1:E2\" \"=LABEL(Wide)*2\")\n",
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
(formula \"Board\" \"D3\" \"=LABEL(B2:C2)\")
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
            let window = window_over(m, sheet, Some(w), None).unwrap();
            let walk = plane::bytes_in(m, m, &window);
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
        // Each kind of write keeps the plane of the sheet it changes, each
        // cell it changed given its byte and the plane taken over an extent
        // it grew, below and to the right and above (KERNEL.56), every byte
        // the walk's, where each dropped it until the next frame's end.
        for row in [
            "(cell \"Board\" \"B2\" 9)",
            "(formula \"Board\" \"C2\" \"=B2+1\")",
            "(derived \"Board\" \"G2\" 3)",
            "(cell \"Board\" \"C5:D6\" 2)",
            "(cell \"Board\" \"I1\" 1)",
            "(cell \"Board\" \"B9\" 4)",
        ] {
            assert!(m.write(row).is_ok(), "{row}");
            walked(&m, "Board", BOARD_WINDOWS, true);
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
        again.end_twins();
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
        // frame's view walks no sheet. Since KERNEL.49 a fill's master is its
        // fill's own, which the machine keeps, so no view walks a sheet for
        // it either, before the first frame, no cell having a value yet, or
        // in the grid, which moves a fill's text to each cell from its fill.
        // The values are the grid less its formula lines throughout, over a
        // window of the fill's cells and over the whole sheet.
        let mut m = load(LIFE);
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
                assert_eq!((v.walked, v.extents), (0, 0), "{when}: the values");
                assert_eq!((g.walked, g.extents), (0, 0), "{when}: the grid");
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
        // A date's code draws since KERNEL.8's third slice; a locale's
        // bracket before it is still refused.
        let e = refused("(format 1 none \"[$-409]yyyy-mm-dd\" nowrap)\n");
        assert_eq!(e.refusal.id, "grid-format-code");
        assert!(e.refusal.text.contains("a locale"), "{}", e.refusal.text);
        load("(format 1 none \"yyyy-mm-dd\" nowrap)\n");
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
            let book = draft.workbook();
            let again = view_text_valued(&book, &w, &Drafted(draft.values.clone()));
            assert_eq!(again, golden, "{sheet} {window}");
            let values = view_text_values(&book, &w, &Drafted(draft.values.clone()));
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
                // A range's ends moved as references are, and its reach with
                // them where it moves as they do (KERNEL.8, slice f).
                (Expr::Range(x), Expr::Range(y)) => {
                    same(&x.left, &y.left, d)
                        && same(&x.right, &y.right, d)
                        && match (&x.reach, &y.reach) {
                            (Some(p), Some(q)) if x.rigid => {
                                p.moved(d.0, d.1) == Some((q.top, q.left, q.bottom, q.right))
                            }
                            _ => true,
                        }
                }
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
        let book = &m.workbook();
        let names = Names::of(book);
        let extents: Vec<Extent> = book.sheets.iter().map(Sheet::extent).collect();
        let s = Shapes::of(book, 1, &library::language(), &names, &extents);
        assert_eq!(
            s.order_cells().first(),
            Some(&(0, 2, 1)),
            "A2 before what reads it"
        );
        assert!(s.cycles().is_empty());
    }

    // ---- KERNEL.34: a formula's write in time with its cells, and the grid
    // a write leaves the grid a load of it makes ----

    /// The plan a load of `m`'s grid makes, spelled whole.
    fn plan_of_a_load(m: &Machine) -> String {
        Shapes::of_model(&m.book, m.sheets, &m.library, &m.names, &m.extents).spelled()
    }

    /// A machine loaded from the grid `m` holds, its rows' sheets copied
    /// whole, fills and all, and made by `Machine::new`, as a load of the
    /// same rows makes one.
    fn loaded_again(m: &Machine) -> Machine {
        let mut workbook = m.workbook();
        workbook.sheets.truncate(m.sheets);
        Machine::new(Draft::of_workbook(&workbook), library::language())
            .unwrap_or_else(|e| panic!("line {}: {}", e.line, e.refusal.text))
    }

    /// Every formula cell of the rows' sheets with its text as the formula
    /// bar shows it, a fill's cells through their master.
    fn formula_texts(m: &Machine) -> BTreeMap<CellId, String> {
        let mut out = BTreeMap::new();
        let book = m.workbook();
        for (s, sheet) in book.sheets[..m.sheets].iter().enumerate() {
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
        let book = m.workbook();
        for (s, sheet) in book.sheets[..m.sheets].iter().enumerate() {
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
        let book = m.workbook();
        let held = twin_reasons(m) == 0;
        assert!(m.step(0).done);
        if held {
            let calc = crate::calc::Calc::run(&book, &m.library);
            let by_calc: BTreeMap<CellId, String> = m
                .shapes
                .order_cells()
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
            m.workbook().sheets[0].cells[&(1, 3)].style,
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
        let _checked = Checked::on();
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
    /// (KERNEL.35); how many sheets were indexed. The readers checked are
    /// the same grid's planned with no shape cut into pieces, so that the
    /// pieced index is held to the unpieced scan (KERNEL.54).
    fn reaches_as_the_scan_does(m: &Machine, (rows, cols): (u32, u32), when: &str) -> usize {
        use crate::calc::reach::{checked, Reached};
        use crate::sheet::{MAX_COLUMN, MAX_ROW};
        let reach = m.shapes.reach();
        let unpieced = {
            let _whole = crate::calc::reach::whole::stands();
            Shapes::of_model(&m.book, m.sheets, &m.library, &m.names, &m.extents)
        };
        assert_eq!(
            unpieced.order_cells(),
            m.shapes.order_cells(),
            "{when}: the orders"
        );
        assert_eq!(unpieced.reach().pieces(), 0, "{when}: unpieced");
        let key = |r: Reached| match r {
            Reached::Cell(p) => (0, p),
            Reached::Shape(s) => (1, s),
            Reached::Parts(s) => (2, s),
        };
        let edges = (1..=6)
            .flat_map(|row| (MAX_COLUMN - 4..=MAX_COLUMN).map(move |col| (row, col)))
            .chain([(MAX_ROW, 1), (MAX_ROW - 2, 3), (MAX_ROW, MAX_COLUMN)]);
        let mut indexed = 0;
        for sheet in 0..m.book.pages.len() {
            indexed += usize::from(reach.most_on(sheet).is_some());
            let readers = reach.readers_on(sheet) as u64;
            let name = &m.book.pages[sheet].sheet.name;
            let window = (1..=rows).flat_map(|row| (1..=cols).map(move |col| (row, col)));
            for (row, col) in window.chain(edges.clone()) {
                let id = (sheet, row, col);
                let (mut by_index, mut by_scan) = (Vec::new(), Vec::new());
                checked::begin();
                m.shapes.reached(id, &mut |r| by_index.push(key(r)));
                let checks = checked::end();
                unpieced.reached_by_scan(id, &mut |r| by_scan.push(key(r)));
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

    /// A text written at `C1` moved to a cell, as a fill reads there: the
    /// copies of one text one shape wherever they stand (KERNEL.34).
    fn moved_to(text: &str, row: u32, col: u32) -> String {
        shift_a1_references(text, i64::from(row) - 1, i64::from(col) - 3)
    }

    /// A grid of a few texts each moved to scattered cells of sheets S and T
    /// over 200 rows (KERNEL.54), each one shape whose cells stand far apart:
    /// references moving over a value sheet V and over the cell's own sheet's
    /// values, a range moving with them, held rectangles and parts; one text
    /// copied twice to a tile's neighbouring cells, and one also as a fill,
    /// which is a shape of its own. A reference to a cell's own sheet reads
    /// a column left of it, so the grid holds no cycle. With two writes: a
    /// value of V, and one more copy of a text, which plans the grid again.
    fn scattered_grid(d: &mut Dice) -> (String, Vec<String>) {
        const TEXTS: [&str; 6] = [
            "=V!A1*2",
            "=SUM(V!A1:B2)+1",
            "=A2+B1",
            "=V!$A$1+V!A1",
            "=SUM(V!$A$1:$B$3)*V!B2",
            "=IF(V!$C$2>3,A1,V!A1)",
        ];
        let mut out = String::new();
        for sheet in ["S", "T", "V"] {
            out.push_str(&format!("(sheet \"{sheet}\" visible)\n"));
        }
        for row in 1..=40 {
            for col in ["A", "B", "C", "D"] {
                if d.below(3) != 0 {
                    out.push_str(&format!("(cell \"V\" \"{col}{row}\" {})\n", d.below(9)));
                }
            }
        }
        out.push_str("(cell \"S\" \"A1:A200\" 2)\n(cell \"T\" \"A1:A200\" 3)\n");
        let mut taken = BTreeSet::new();
        let mut copies = Vec::new();
        let first = d.below(TEXTS.len());
        for k in 0..3 + d.below(3) {
            let text = TEXTS[(first + k) % TEXTS.len()];
            for _ in 0..2 + d.below(9) {
                let at = (
                    d.pick(&["S", "T"]),
                    1 + d.below(200) as u32,
                    3 + d.below(8) as u32,
                );
                if taken.insert(at) {
                    copies.push((text, at));
                }
            }
        }
        // A copy beside another, in its tile unless at its edge.
        if let Some(&(text, (sheet, row, col))) = copies.first() {
            let at = (sheet, row + 1, col);
            if taken.insert(at) {
                copies.push((text, at));
            }
        }
        for (text, (sheet, row, col)) in &copies {
            let cell = crate::sheet::cell_ref(*row, *col);
            let f = crate::rows::quoted(&format!("={}", moved_to(&text[1..], *row, *col)));
            out.push_str(&format!("(formula \"{sheet}\" \"{cell}\" {f})\n"));
        }
        let fill = moved_to(&TEXTS[first][1..], 150, 12);
        out.push_str(&format!("(formula \"S\" \"L150:N155\" \"={fill}\")\n"));
        let (text, _) = copies[0];
        let (row, col) = (190 + d.below(10) as u32, 3 + d.below(8) as u32);
        let cell = crate::sheet::cell_ref(row, col);
        let again = crate::rows::quoted(&format!("={}", moved_to(&text[1..], row, col)));
        let writes = vec![
            format!("(cell \"V\" \"B{}\" {})", 1 + d.below(40), d.below(9)),
            format!("(formula \"T\" \"{cell}\" {again})"),
        ];
        (out, writes)
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
        // reaches through the index what every reader checked reaches. And
        // grids of far-apart cells of one text, each written as it is
        // (KERNEL.54): their shapes cut into pieces, the pieced index reaches
        // what the unpieced scan reaches.
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
        let regular = grids.len();
        for seed in 1..=6u64 {
            let mut d = Dice(seed ^ 0x5454);
            let (rows, writes) = scattered_grid(&mut d);
            grids.push((format!("far-apart seed {seed}"), rows, writes));
        }
        grids.push((
            "12 far-apart pairs".to_string(),
            far_apart(12),
            vec![
                "(cell \"S\" \"A4\" 9)".to_string(),
                "(formula \"S\" \"C700\" \"=A703\")".to_string(),
            ],
        ));
        let (mut indexed, mut cases, mut pieced) = (0, 0, 0);
        for (k, (name, rows, writes)) in grids.iter().enumerate() {
            let around = (last_row(rows) + 1, 18);
            let hows: &[Spelling] = if k < regular {
                &[
                    Spelling::One,
                    Spelling::Rows,
                    Spelling::Cols,
                    Spelling::Cells(2),
                ]
            } else {
                &[Spelling::One]
            };
            for &how in hows {
                let text = spelled(rows, how);
                for way in [None, Some(None), Some(Some(true)), Some(Some(false))] {
                    let _every = way.map(crate::calc::reach::every::sheet_indexed);
                    let when = format!("{name}, {how:?}, every sheet indexed {way:?}");
                    let mut m = load(&text);
                    indexed += reaches_as_the_scan_does(&m, around, &when);
                    pieced += usize::from(m.shapes.reach().pieces() > 0);
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
        let far = grids.len() - regular;
        assert!(
            cases == regular * 16 + far * 4 && indexed > cases * 2 && pieced == far * 4,
            "{indexed} sheets indexed and {pieced} grids pieced over {cases} cases: too few for the law to hold over"
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
                    m.shapes.reached_by_scan((v, row, col), &mut |_| {});
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

    // ---- KERNEL.54: a change's cost set by the change ----

    /// `p` pairs of cells of one text standing far apart (KERNEL.54): on sheet
    /// S, values in column A, and for each `i` from 1 to `p` a formula reading
    /// A `i` rows below its own cell, written at `C{i}` and again 1,000 rows
    /// lower, so that each pair is one shape (`KERNEL.34`'s rule, one text one
    /// shape) whose cells' box is 1,001 rows tall.
    fn far_apart(p: u32) -> String {
        far_apart_as(p, false)
    }

    /// [`far_apart`], each copy given a text of its own (`+0*k`,
    /// KERNEL.35's spelling) where `own`, so that no shape is cut.
    fn far_apart_as(p: u32, own: bool) -> String {
        let gap = 1000;
        let rows = 2 * p + gap + 1;
        let mut s = format!("(cell \"S\" \"A1:A{rows}\" 1)\n");
        for i in 1..=p {
            for at in [i, i + gap] {
                let own = if own {
                    format!("+0*{at}")
                } else {
                    String::new()
                };
                s.push_str(&format!(
                    "(formula \"S\" \"C{at}\" \"=A{}{own}\")\n",
                    at + i
                ));
            }
        }
        s
    }

    #[test]
    fn far_apart_cells_of_one_text_are_checked_where_a_change_reaches_them() {
        // KERNEL.54, counted and not timed: three value cells written, each
        // read by the first cell of one pair, and the frame's six changed
        // cells, those three and the three formulas reading them. Each pair's
        // shape is cut into two pieces, its reader filed once a piece over
        // the piece's one cell, so the sheet's readers, two a pair, are
        // indexed by rows, a band each: A2, A4 and A6 meet one reader each,
        // C2 the reader of row 2, and C1 and C3 none, four at 32 pairs and at
        // 64, as many as the same grid with every copy its own shape. They
        // were each checked against every pair's reader, its footprint read
        // over its shape's box of 1,001 rows, too wide to file: 192 and 384.
        use crate::calc::reach::checked;
        for p in [32u32, 64] {
            let mut checks = Vec::new();
            for own in [false, true] {
                let mut m = load(&far_apart_as(p, own));
                assert!(m.step(0).done && m.step(0).done);
                for i in [1u32, 2, 3] {
                    assert_eq!(m.write(&format!("(cell \"S\" \"A{}\" 2)", 2 * i)), Ok(1));
                }
                checked::begin();
                assert!(m.step(0).done);
                checks.push(checked::end());
                assert_eq!(m.changed(), 3, "{p} pairs: the formulas the frame changed");
                let census = m.census();
                assert!(
                    census.contains(&format!("(readers \"S\" {} (most 1))\n", 2 * p)),
                    "{p} pairs, each copy its own text {own}: {census}"
                );
            }
            assert_eq!(
                checks,
                [4, 4],
                "{p} pairs: the readers checked, pieced and each copy its own shape"
            );
        }
    }

    #[test]
    fn a_far_apart_shape_is_cut_by_tiles_its_moving_references_filed_a_piece() {
        // KERNEL.54: one text at C31, C32, C1000, C1005 and C2000, its box
        // 1,970 rows by one column, its five cells a 394th of it, cut into
        // pieces by tiles of 32 rows and columns: C31 and C32 either side of
        // a tile's edge, C1000 and C1005 one piece; its moving reference
        // filed once a piece, its held rectangle and its part once. Another
        // text at AE3000, AF3000 and AE5000, its columns either side of a
        // tile's edge, three pieces; and a fill of 200 rows, its every cell
        // in its box, one shape and one reader as before. A change reaches
        // each piece's cells: C1005, the second cell of its piece.
        let mut rows = String::from(concat!(
            "(cell \"S\" \"A1:A200\" 1)\n(cell \"S\" \"B1:B3\" 1)\n(cell \"S\" \"D1\" 10)\n",
            "(formula \"S\" \"E1:E200\" \"=A1+1\")\n",
        ));
        for r in [31, 32, 1000, 1005, 2000] {
            if r > 200 {
                rows.push_str(&format!("(cell \"S\" \"A{r}\" 1)\n"));
            }
            rows.push_str(&format!(
                "(formula \"S\" \"C{r}\" \"=A{r}*2+SUM($B$1:$B$3)+$D$1\")\n"
            ));
        }
        for (at, read) in [
            ("AE3000", "A3000"),
            ("AF3000", "B3000"),
            ("AE5000", "A5000"),
        ] {
            rows.push_str(&format!("(formula \"S\" \"{at}\" \"={read}*3\")\n"));
        }
        let mut m = load(&rows);
        let census = m.census();
        assert!(
            census.contains("(readers \"S\" 10)\n"),
            "four pieces' and one held rectangle's and one part's readers, three pieces', one fill's: {census}"
        );
        assert!(m.step(0).done);
        assert_eq!(m.write("(cell \"S\" \"A1005\" 7)"), Ok(1));
        assert!(m.step(0).done);
        assert_eq!(m.changed(), 1, "C1005 alone changed");
        assert_eq!(
            m.computed((0, 1005, 3)),
            Some(Computed::Value(Value::Number(27.0))),
            "7 * 2 + 3 + 10"
        );
    }

    #[test]
    fn a_value_written_walks_no_cell_for_its_extent() {
        // KERNEL.54, counted: one value written into a sheet of `2n` cells,
        // at 100 rows and at 1,000, walks none of them for the sheet's
        // extent, grown by the cells the write changed, where it walked every
        // one of them (`Sheet::extent`, from `settle`): 200 and 2,000. The
        // extent kept is the sheet's walked, after that write and after one
        // past the extent's corner.
        for n in [100u32, 1000] {
            let mut m = load(&format!(
                "(cell \"S\" \"A1:A{n}\" 1)\n(formula \"S\" \"B1:B{n}\" \"=A1*2\")\n"
            ));
            assert!(m.step(0).done);
            crate::sheet::work::begin();
            assert_eq!(m.write("(cell \"S\" \"A5\" 3)"), Ok(1));
            let work = crate::sheet::work::end();
            assert_eq!(work.extents, 0, "{n} rows: the cells walked for the extent");
            assert_eq!(m.extents[0], m.workbook().sheets[0].extent(), "{n} rows");
            let far = format!("(cell \"S\" \"{}\" 3)", crate::sheet::cell_ref(n + 7, 4));
            assert_eq!(m.write(&far), Ok(1));
            assert_eq!(m.extents[0], Some(((1, 1), (n + 7, 4))), "{n} rows: {far}");
            assert_eq!(
                m.extents[0],
                m.workbook().sheets[0].extent(),
                "{n} rows: {far}"
            );
        }
    }

    #[test]
    fn a_refused_write_leaves_the_extents_it_found() {
        // KERNEL.54: a write whose first row lands past the extent's corner
        // and whose second makes a cycle is refused when it settles, after
        // its extent grew by the first; the undo puts back the extent the
        // write found, where it walked the sheet again.
        let mut m = load("(cell \"S\" \"A1:A10\" 1)\n(formula \"S\" \"B1:B10\" \"=A1*2\")\n");
        assert!(m.step(0).done);
        let found = m.extents.clone();
        let refused = m.write("(cell \"S\" \"Z500\" 1)\n(formula \"S\" \"C1\" \"=C1+1\")\n");
        assert_eq!(
            refused.map_err(|e| e.refusal.id),
            Err("calc-cycle".to_string())
        );
        assert_eq!(m.extents, found, "the extents the write found");
        assert_eq!(m.extents[0], m.workbook().sheets[0].extent());
    }

    /// `n` cells alone, each a shape of its own reading the value beside it
    /// (`=A{r}+{r}`), and a fill of `n` cells reading the same values, one run
    /// (KERNEL.54).
    fn alone_and_a_run(n: u32) -> String {
        let mut s = format!("(cell \"S\" \"A1:A{n}\" 1)\n(formula \"S\" \"C1:C{n}\" \"=A1*2\")\n");
        for r in 1..=n {
            s.push_str(&format!("(formula \"S\" \"B{r}\" \"=A{r}+{r}\")\n"));
        }
        s
    }

    #[test]
    fn an_idle_frame_touches_nothing_and_a_cone_its_marks() {
        // KERNEL.54, counted: at 64 cells alone and at 128 beside a run, a
        // frame that computes nothing visits no position, no run and no
        // shape and clears nothing, where it asked each of its `n` cells
        // alone, visited its run, asked each of its `n + 1` shapes for its
        // parts and cleared every shape's two flags and every word of its
        // marks; after one value written, its frame visits its two marks,
        // the cell alone and the run's cell, and clears their two words; and
        // the frame after it, nothing again.
        let idle = walk::Walk::default();
        let cone = walk::Walk {
            visited: 1,
            runs: 1,
            shapes: 0,
            cleared: 2,
        };
        for n in [64u32, 128] {
            let mut m = load(&alone_and_a_run(n));
            for _ in 0..3 {
                assert!(m.step(0).done);
            }
            for (write, want, reached) in [(None, idle, 0), (Some(5), cone, 2), (None, idle, 0)] {
                if let Some(v) = write {
                    assert_eq!(m.write(&format!("(cell \"S\" \"A5\" {v})")), Ok(1));
                }
                walk::begin();
                assert!(m.step(0).done);
                let w = walk::end();
                assert_eq!(
                    m.reached(),
                    reached,
                    "{n} cells alone: what the frame computed"
                );
                assert_eq!(w, want, "{n} cells alone: what the frame touched");
            }
        }
    }

    // ---- KERNEL.36: a frame's end at any share ----

    /// The twin cells the last frame's end changed, by place (KERNEL.42;
    /// KERNEL.49): a formula's twin cell by its position, and a value cell's.
    fn twin_changes(m: &Machine) -> BTreeSet<CellId> {
        let order = m.shapes.order_cells();
        m.twin_changed
            .iter()
            .map(|p| {
                let (s, row, col) = order[p];
                (m.sheets + s, row, col)
            })
            .chain(m.twin_written.iter().copied())
            .collect()
    }

    /// Two grids' frames' ends the same to the bit, one copying into its
    /// twins' stores and its planes where it may and one making each again
    /// at every frame's end (`same_state`); whether the next frame computes
    /// every cell; and, where it does not, the twin cells kept as changed,
    /// as a set, those whose twin the end changed, against the references
    /// taken at the two ends (`sheets_left`, KERNEL.49).
    fn same_end(copied: &Machine, walked: &Machine, before: &Left, when: &str) {
        same_state(copied, walked, when);
        assert_eq!(
            copied.every_next, walked.every_next,
            "{when}: the next frame computing every cell"
        );
        if !copied.every_next {
            let set = twin_changes;
            assert_eq!(set(copied), set(walked), "{when}: the twin cells changed");
            let after = sheets_left(copied);
            let mut changed = BTreeSet::new();
            for s in 0..copied.sheets {
                let places: BTreeSet<(u32, u32)> =
                    before[s].keys().chain(after[s].keys()).copied().collect();
                for at in places {
                    let same = match (before[s].get(&at), after[s].get(&at)) {
                        (Some(x), Some(y)) => same_got(x, y),
                        (x, y) => x.is_none() && y.is_none(),
                    };
                    if !same {
                        changed.insert((copied.sheets + s, at.0, at.1));
                    }
                }
            }
            assert_eq!(
                set(copied),
                changed,
                "{when}: the twin cells the end changed"
            );
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
    fn a_frame_s_end_copied_is_the_end_made_again_bit_for_bit() {
        // KERNEL.36; KERNEL.49: grids of generated formulas most of whose
        // cells change at every frame, the Clock written before every frame,
        // a frame's end copying into the twins' stores and the planes however
        // many changed, against the same grids making each store and plane
        // again at every frame's end, frame for frame and bit for bit
        // (`same_end`), the twin cells changed held to the references taken
        // at each end; the values against the same grids computing every cell
        // (KERNEL.20's reference). Writes between frames change value cells,
        // add cells and change formulas; a sheet is viewed as a plane, the
        // first, which most writes change, or the second, its plane kept
        // through a write to it (KERNEL.56); steps yield. Frames whose ends
        // copied more than half the formula cells, and frames after the
        // first whose ends made a twin's extent or store again, each seen:
        // a twin cell lost, since a write that added a cell or changed the
        // viewed sheet no longer makes one again (KERNEL.56).
        let _checked = Checked::on();
        let (mut past_half, mut made_again) = (0, 0);
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
            let mut before = sheets_left(&copied);
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
                let of = copied.shapes.len();
                let work = ended(&mut copied, 1 + d.below(300));
                assert!(walked.step(0).done && whole.step(0).done);
                same_end(&copied, &walked, &before, &when);
                before = sheets_left(&copied);
                assert_eq!(values(&copied), values(&whole), "{when}: computed whole");
                past_half += usize::from(work.swept == 0 && work.copied as usize * 2 > of);
                made_again += usize::from(work.swept > 0 && frame > 1);
            }
        }
        assert!(
            past_half >= 80 && made_again >= 8,
            "{past_half} frames' ends copied past half and {made_again} made something \
             again: too few for the law to hold over"
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
    fn a_frame_s_end_copies_at_any_share_into_an_array_or_a_sparse_store() {
        // KERNEL.36; KERNEL.49: 400 formulas of which 198, 200, 202 or all
        // change at every frame, the Clock written before each. Into a twin
        // kept as an array, and into one too sparse for an array, every
        // frame's end copies the cells it changed and the Clock and walks no
        // sheet, where a sparse sheet's twin's own cells, a B-tree, were
        // walked past half the cells. Each held to the same grids making each
        // store again at every frame's end (`same_end`), its values to every
        // cell computed, and past a twin change beyond the room kept for
        // them, which makes the next frame compute every cell.
        let mut overflowed = 0;
        for sparse in [false, true] {
            for k in [99u32, 100, 101, 200] {
                let rows = share_grid(200, k, sparse);
                let mut copied = load(&rows);
                let mut walked = load(&rows);
                walked.copies = false;
                let mut whole = load(&rows);
                whole.incremental = false;
                let mut before = sheets_left(&copied);
                for frame in 1..=6u64 {
                    let when = format!("{k} of 200 changing, sparse {sparse}, frame {frame}");
                    let w = format!("(cell \"Clock\" \"B1\" {frame})");
                    for m in [&mut copied, &mut walked, &mut whole] {
                        assert_eq!(m.write(&w), Ok(1), "{when}");
                    }
                    let work = ended(&mut copied, 0);
                    assert!(walked.step(0).done && whole.step(0).done);
                    same_end(&copied, &walked, &before, &when);
                    before = sheets_left(&copied);
                    assert_eq!(values(&copied), values(&whole), "{when}: computed whole");
                    let grid = copied.find_sheet("Grid").expect("Grid");
                    assert_eq!(
                        store_kept(&copied, "Grid"),
                        Some(!sparse),
                        "{when}: Grid's twin"
                    );
                    assert_eq!(
                        copied.dense(copied.sheets + grid).is_none(),
                        sparse,
                        "{when}"
                    );
                    if frame < 3 {
                        continue;
                    }
                    let changed = 2 * u64::from(k);
                    assert_eq!(
                        (work.copied, work.swept),
                        (changed + 1, 0),
                        "{when}: copied"
                    );
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
    fn a_frame_s_end_copies_beside_a_twin_no_formula_reads() {
        // KERNEL.36; KERNEL.49: a sheet whose twin is kept as an array, all
        // 400 of its formulas changing, beside a sheet too sparse for one,
        // 500 or 501 of its 600 changing, whose twin no formula reads and so
        // keeps no store: a frame's end copies the first's changes into its
        // twin and tells the second's, with the Clock 901 or 902 cells, and
        // walks no sheet, where the share of the cells it would copy through
        // twins' own cells decided between a copy and a walk. Each frame held
        // to the same grids making each store again at every frame's end
        // (`same_end`) and its values to every cell computed.
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
            let mut before = sheets_left(&copied);
            for frame in 1..=5u64 {
                let when = format!("{j} of Far's 600 changing, frame {frame}");
                let w = format!("(cell \"Clock\" \"B1\" {frame})");
                for m in [&mut copied, &mut walked, &mut whole] {
                    assert_eq!(m.write(&w), Ok(1), "{when}");
                }
                let work = ended(&mut copied, 0);
                assert!(walked.step(0).done && whole.step(0).done);
                same_end(&copied, &walked, &before, &when);
                before = sheets_left(&copied);
                assert_eq!(values(&copied), values(&whole), "{when}: computed whole");
                assert_eq!(
                    (store_kept(&copied, "Grid"), store_kept(&copied, "Far")),
                    (Some(true), None),
                    "{when}"
                );
                assert_eq!(copied.cells, 1_002, "{when}");
                if frame < 3 {
                    continue;
                }
                assert_eq!(
                    (work.copied, work.swept),
                    (401 + u64::from(j), 0),
                    "{when}: copied"
                );
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
    fn a_twin_cell_lost_makes_its_extent_again_and_one_gained_grows_it() {
        // KERNEL.36; KERNEL.49; KERNEL.56: a cell whose value is a number
        // for two frames and a reason for the next two, beside a column most
        // of whose cells change and a cell reading its twin, which follows it
        // a frame late: twin cells gained alone at the first, fifth and sixth
        // frames, the first frame's values among them, and lost at the
        // third, fourth, seventh and eighth. A loss makes the twin's extent
        // again from its sheet at the frame's end and the next frame compute
        // every cell; a gain grows the extent by the cell, no cell walked,
        // and the next frame follows it as any change, where every gain was a
        // loss's. Held to the same grids making each store again
        // (`same_end`), the twin's extent to the rectangle of its cells, and
        // its values to every cell computed.
        let rows = concat!(
            "(sheet \"Clock\" visible)\n(sheet \"Grid\" visible)\n(cell \"Clock\" \"B1\" 0)\n",
            "(formula \"Grid\" \"A1:A100\" \"=MOD(Clock!$B$1+ROW(),16)\")\n",
            "(formula \"Grid\" \"B1\" \"=FLIP(INT((Clock!$B$1-1)/2))\")\n",
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
        let mut before = sheets_left(&copied);
        for frame in 1..=8u64 {
            let when = format!("frame {frame}");
            let w = format!("(cell \"Clock\" \"B1\" {frame})");
            for m in [&mut copied, &mut walked, &mut whole] {
                assert_eq!(m.write(&w), Ok(1), "{when}");
            }
            let work = ended(&mut copied, 0);
            assert!(walked.step(0).done && whole.step(0).done);
            same_end(&copied, &walked, &before, &when);
            before = sheets_left(&copied);
            assert_eq!(values(&copied), values(&whole), "{when}: computed whole");
            let grid = copied.find_sheet("Grid").expect("Grid");
            let reason = matches!(frame, 3 | 4 | 7 | 8);
            assert_eq!(
                matches!(copied.twin_held(grid, 1, 2), Held::Reason(_)),
                reason,
                "{when}: B1's twin a reason for two frames of four"
            );
            let twin = copied.sheets + grid;
            assert_eq!(
                copied.extents[twin],
                copied.twin_extent(grid),
                "{when}: the twin's extent the rectangle of its cells"
            );
            // B1's twin cell lost at the third and seventh frames' ends, and
            // C1's a frame later.
            assert_eq!(
                (work.swept > 0, copied.every_next),
                (reason, reason),
                "{when}: the extent made again and the next frame whole for a loss alone"
            );
        }
    }

    #[test]
    fn a_frame_s_end_visits_the_cells_it_changed_and_not_the_grid() {
        // KERNEL.36, counted and not timed: the cells a frame's end visits.
        // Copying, the cells that changed and the cell written, at 1,000
        // formulas a column and at 4,000 alike when as many change, and nine
        // tenths of them when nine tenths do; making each twin's store again
        // at every frame's end, the same cells told and every cell of the
        // sheet whose twin a formula reads, the Clock's read by none
        // (KERNEL.49). The set of the cells changed holds a bit for every
        // formula cell from the load and never grows.
        for n in [1_000u32, 4_000] {
            for k in [100u32, n * 9 / 10] {
                let rows = share_grid(n, k, false);
                let mut copied = load(&rows);
                let mut walked = load(&rows);
                walked.copies = false;
                let room = copied.changed.words();
                assert_eq!(
                    room,
                    copied.shapes.len().div_ceil(64),
                    "{n}: the set's room"
                );
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
                            (changed + 1, 2 * u64::from(n)),
                            "{when}: made again"
                        );
                    }
                }
                assert_eq!(copied.changed.words(), room, "{n}: the set grew");
            }
        }
    }

    #[test]
    fn a_frame_s_end_takes_what_a_write_grew_and_makes_again_what_it_turned() {
        // KERNEL.36; KERNEL.49; KERNEL.56: frames whose ends copy, through a
        // write into the sheet viewed as a plane, which keeps its plane, each
        // cell it changed given its byte; and through a write that adds a
        // cell past the sheet's extent, below it and to its right, which takes
        // the sheet's store and its plane over the new extent where they
        // stand, the frame's end taking the twin's store over it the same way
        // and growing the twin's extent by the cell, no cell walked, where
        // each was made again from the sheet; then a cell written so far out
        // that the sheet's store turns sparse, which the twin's array cannot
        // follow, the frame's end making it again from the sheet's. After
        // each the frames' ends copy, the Clock written into its own sheet
        // before every frame. Each frame held to the same grid making each
        // again at every frame's end (`same_end`).
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
        let mut made_again = 0;
        let mut before = sheets_left(&copied);
        for frame in 1..=12u64 {
            let when = format!("frame {frame}");
            let mut w = format!("(cell \"Clock\" \"B1\" {frame})");
            match frame {
                5 => w.push_str("\n(cell \"Grid\" \"C7\" 5)"),
                8 => w.push_str("\n(cell \"Grid\" \"D400\" 1)"),
                10 => w.push_str("\n(cell \"Grid\" \"Z5000\" 1)"),
                _ => {}
            }
            assert!(
                copied.write(&w).is_ok() && walked.write(&w).is_ok(),
                "{when}"
            );
            if frame >= 5 {
                assert!(
                    kept(&copied, "Grid"),
                    "{when}: the plane kept through the write"
                );
            }
            let work = ended(&mut copied, 0);
            assert!(walked.step(0).done);
            same_end(&copied, &walked, &before, &when);
            before = sheets_left(&copied);
            match frame {
                10 => {
                    assert!(work.swept > 0, "{when}: made again");
                    made_again += 1;
                }
                2..=9 | 11 | 12 => assert_eq!(work.swept, 0, "{when}: copied"),
                _ => {}
            }
            // The twin's array of 300 rows by three columns taken over the
            // sheet's of 400 by four, its 900 slots moved.
            let relaid = if frame == 8 { 900 } else { 0 };
            assert_eq!(work.relaid, relaid, "{when}: the twin's slots moved");
            assert!(
                copied.twins[grid]
                    .as_ref()
                    .is_some_and(|t| t.beside(&copied.stores[grid])),
                "{when}: Grid's twin's store over its sheet's extent"
            );
        }
        assert_eq!(made_again, 1);
    }

    #[test]
    fn a_twin_s_store_follows_its_sheet_s_over_an_extent_a_reason_grew() {
        // KERNEL.36; KERNEL.49; KERNEL.56: a formula written past a sheet's
        // extent whose value is a reason, so that the extent grows and the
        // twin gains no cell: the write takes the sheet's store over the new
        // extent, and the frame's end the twin's the same way, its 300 slots
        // moved where it was made again from the sheet, its own extent the
        // rectangle of its cells still, a walk of it meeting the reason past
        // them; the frames' ends copy before and after, values alternating at
        // every frame among them. Held to the same grid making each store
        // again at every frame's end (`same_end`) and its values to every
        // cell computed; a plane asked for after the sixth frame, first kept
        // at the seventh frame's end.
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
        let mut before = sheets_left(&copied);
        for frame in 1..=9u64 {
            let when = format!("frame {frame}");
            let mut w = format!("(cell \"Clock\" \"B1\" {frame})");
            if frame == 4 {
                w.push_str("\n(formula \"Grid\" \"E200\" \"=LABEL(A1:A2)\")");
            }
            for m in [&mut copied, &mut walked, &mut whole] {
                assert!(m.write(&w).is_ok(), "{when}");
            }
            let work = ended(&mut copied, 0);
            assert!(walked.step(0).done && whole.step(0).done);
            same_end(&copied, &walked, &before, &when);
            before = sheets_left(&copied);
            assert_eq!(values(&copied), values(&whole), "{when}: computed whole");
            match frame {
                2..=6 | 8 | 9 => assert_eq!(work.swept, 0, "{when}: copied"),
                7 => assert!(work.swept > 0, "{when}: made again"),
                _ => {}
            }
            let relaid = if frame == 4 { 300 } else { 0 };
            assert_eq!(work.relaid, relaid, "{when}: the twin's slots moved");
            if frame == 4 {
                assert!(
                    copied.twins[grid]
                        .as_ref()
                        .is_some_and(|t| t.beside(&copied.stores[grid])),
                    "{when}: the twin's store over the new extent"
                );
                let twin = copied.sheets + grid;
                assert_eq!(
                    copied.extents[twin],
                    Some(((1, 1), (100, 3))),
                    "{when}: the twin's extent its cells'"
                );
                assert_eq!(
                    copied.read((twin, 200, 5)),
                    Err(Reason::Construct("A1:A2".to_string())),
                    "{when}: the reason read past them"
                );
            }
            // A plane asked for after the sixth frame: its first kept at the
            // seventh frame's end, which makes it.
            if frame == 6 {
                for m in [&copied, &walked, &whole] {
                    m.view("plane", "Grid", None).expect("a plane of Grid");
                }
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
                "(formula \"S\" \"D1\" \"=IF(B1=1,LABEL(X1:X2),LABEL(Y1:Y2))\")\n",
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
        // into the twins its formulas read and makes the others from their
        // sheets (KERNEL.49), one that keeps every twin's store, one that
        // makes each store and plane again at every frame's end, and one
        // that computes every cell, over a grid of values written between
        // frames, a twin read, a day, a value turning between 0 and -0 and a
        // reason turning, with a plane asked for midway.
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
        machines[1].every_twin = true;
        machines[1].keep_twins();
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
            machines[0].twins.iter().any(Option::is_none),
            "a twin made from its sheet"
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
    /// reference, strict functions through their pointers handed values,
    /// ranges and references, and `INDEX` (KERNEL.8).
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
            // `INDEX`, which a program picks by itself, of a range or of a
            // cell moved off the sheet, its row and column of every kind;
            // and a lookup or a criterion through its pointer, at the root
            // or inside (KERNEL.8).
            20 => match d.below(5) {
                0 => {
                    let col = if d.below(2) == 0 {
                        format!(",{}", e(d))
                    } else {
                        String::new()
                    };
                    format!("INDEX({},{}{col})", range(d, at), e(d))
                }
                1 => format!(
                    "INDEX({},{})",
                    reference(d, at),
                    d.pick(&["1", "0", "2", "TRUE"])
                ),
                2 => format!("({})+INDEX({},{},{})", e(d), range(d, at), e(d), e(d)),
                3 => format!(
                    "MATCH({},{},{})",
                    e(d),
                    vector_of(d, at),
                    d.pick(&["0", "1", "-1"])
                ),
                _ => format!("COUNTIF({},{})", range(d, at), e(d)),
            },
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
        let book = m.workbook();
        let calc = crate::calc::Calc::run(&book, &m.library);
        m.shapes
            .order_cells()
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
        let _checked = Checked::on();
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
        // values, a call answering an error inside the formula (at its root,
        // since KERNEL.8, a call's answer is the cell's, whatever its kind:
        // `a_call_at_a_program_s_root_answers_whatever_its_kind`); and
        // since KERNEL.8 `INDEX`, a call through its pointer, picking a text
        // and picking past its area.
        use program::guarded;
        // Each case's cells, its formula, its guard, and the cells its
        // program answers (the fill's first, whose reference stays on).
        let cases = [
            ("A1", "=S!B1+1", guarded::READ, 0),
            ("B1:B2", "=S!A1048576*0", guarded::OFF, 1),
            ("C1", "=1/S!A3", guarded::ANSWER, 0),
            ("D1", "=SUM(S!A1:B1)", guarded::RANGE, 0),
            ("E1", "=CHOOSE(S!A1+3,1,2)", guarded::INDEX, 0),
            ("F1", "=TWICE(S!B1)+0", guarded::CALL, 0),
            ("G1", "=INDEX(S!A1:B1,1,2)+0", guarded::CALL, 0),
            ("H1", "=INDEX(S!A1:B1,1,3)+0", guarded::CALL, 0),
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
    fn a_call_at_a_program_s_root_answers_whatever_its_kind() {
        // KERNEL.8: a call that is the cell's whole formula answers what the
        // call answers, a text, a truth value or an error among them, with no
        // guard and no evaluation again from the start, as `Ctx` would answer
        // it; the same call inside the formula keeps its guard. And `INDEX`
        // lowered picks a number where it stands, the cell's own program.
        use program::guarded;
        let cases = [
            ("A1", "=TWICE(S!B1)", 1),
            ("A2", "=TWICE(S!A1)", 1),
            ("A3", "=MATCH(9,S!A1:A2,0)", 1),
            ("A4", "=VLOOKUP(2,S!A1:B2,2,FALSE)", 1),
            ("A5", "=INDEX(S!A1:A2,2)*2", 1),
            ("A6", "=COUNTIF(S!A1:B2,3)", 1),
        ];
        for (cell, text, answered) in cases {
            let rows = format!(
                "(sheet \"S\" visible)\n(cell \"S\" \"A1\" 2)\n(cell \"S\" \"A2\" 3)\n\
                 (cell \"S\" \"B1\" \"x\")\n(cell \"S\" \"B2\" \"y\")\n(sheet \"T\" visible)\n\
                 (formula \"T\" \"{cell}\" {})\n",
                crate::rows::quoted(text)
            );
            let mut m = load_programs(&rows, text);
            let mut ctx = load_programs(&rows, text);
            ctx.lowered = false;
            guarded::take();
            assert!(m.step(0).done && ctx.step(0).done);
            let counts = guarded::take();
            assert_eq!(
                counts.iter().sum::<usize>(),
                0,
                "{text}: no guard {counts:?}"
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
        // The first frame computes every cell, and the second none, the
        // twin cells its values gained read by no cell (KERNEL.56); a
        // written value reaches its twin at its frame's end, and the cells
        // that read the twin at the next frame.
        assert_eq!(
            ran,
            [10, 0, 0, 3, 0, 24, 0, 10],
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
        // stepped whole; by runs and every cell alone. In the frames that
        // compute every cell of the runs in one step, the first and the third
        // (the second computes `D1` alone, which reads the sheet written and
        // no twin, the twin cells the first gained read by no cell,
        // KERNEL.56), the runs of twenty are computed as runs and the run of
        // five by its program; the order's runs being twenty, twenty, five and
        // twenty, a run the walk skipped leaves the cells of a run of twenty
        // alone.
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
        let of = load(rows).shapes.len();
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
                    if frame != 2 && budget == of {
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
                for sheet in 0..m.book.pages.len() {
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

    /// Each rows' sheet's store holds what a read of each of its cells
    /// answers now (KERNEL.49): a value cell its content, to the bit, a
    /// formula a formula's kind, its value, its reason or none yet, and a
    /// place no cell holds nothing; an array over the sheet's extent where its
    /// cells stand close enough together, and a map of its cells where they do
    /// not (KERNEL.31's arrays, kept now for every sheet of the rows).
    fn store_is_the_frame(m: &Machine, when: &str) {
        let book = m.workbook();
        for s in 0..m.sheets {
            let sheet = &book.sheets[s];
            let fits = dense::fits(m.extents[s], dense::limit(sheet.cells.len(), CELLS as u64));
            assert_eq!(
                m.stores[s].array().is_some(),
                fits,
                "{when}: sheet {s} kept as an array"
            );
            if let Some(d) = m.stores[s].array() {
                assert!(
                    d.stands_over(m.extents[s]),
                    "{when}: sheet {s}'s array over its extent"
                );
                for i in 0..d.slots() {
                    let (slot, at) = d.at(i);
                    if !sheet.cells.contains_key(&at) {
                        assert!(slot.is_empty(), "{when}: sheet {s}, {at:?} holds no cell");
                    }
                }
            }
            for (&(row, col), cell) in &sheet.cells {
                let slot = m.stores[s].slot_of(row, col);
                if is_formula(&cell.content) {
                    assert!(
                        slot.is_some_and(|slot| !slot.is_empty() || m.shapes.holds((s, row, col))),
                        "{when}: sheet {s}, row {row}, column {col}: a formula's slot {slot:?}"
                    );
                    continue;
                }
                let got = slot.map_or(Value::Empty, |slot| slot.value());
                let want = content_value(&cell.content);
                assert!(
                    same_value(&got, &want),
                    "{when}: sheet {s}, row {row}, column {col}: {got:?} for {want:?}"
                );
            }
        }
    }

    /// The reasons a rows' sheet's store holds (KERNEL.49).
    fn reasons_in(m: &Machine, s: usize) -> usize {
        let book = m.workbook();
        let sheet = &book.sheets[s];
        sheet
            .cells
            .keys()
            .filter(|&&(row, col)| matches!(m.stores[s].slot_of(row, col), Some(Stored::Reason(_))))
            .count()
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

    /// The same rows loaded twice: one machine offering its rows' sheets'
    /// stores to the evaluator as arrays, as every machine does, and one
    /// offering none, its sheets read cell by cell as before KERNEL.31.
    fn kept_and_not(rows: &str) -> (Machine, Machine) {
        let kept = load(rows);
        let mut cells = load(rows);
        cells.arrays = false;
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
        // cell not computed, which its sheet's array holds among its slots
        // (KERNEL.49, where it stepped the array aside), and of a cell past a
        // sheet's extent, which moves its array to the new one, each write
        // planning again keeping the arrays over the extents that stand.
        vector::sliced::take();
        let (mut kept_over, mut made_again, mut kept) = (0, 0, 0);
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
                    let book = m.workbook();
                    let calc = crate::calc::Calc::run(&book, &m.library);
                    let spell =
                        |id: &CellId| format!("{:?}", calc.computed(*id).expect("computed"));
                    m.shapes
                        .order_cells()
                        .iter()
                        .map(|id| (*id, spell(id)))
                        .collect()
                };
                let shown = |m: &Machine| -> Vec<String> {
                    let order = m.shapes.order_cells().into_iter();
                    order.map(|id| format!("{:?}", m.computed(id))).collect()
                };
                for frame in 1..=5 {
                    let when = format!("seed {seed}, stacked {stacked}, frame {frame}");
                    let last = shown(&m);
                    loop {
                        let s = m.step(1 + d.below(300));
                        store_is_the_frame(&m, &when);
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
                    kept += m.stores.iter().filter(|s| s.array().is_some()).count();
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
                    store_is_the_frame(&m, &format!("{when}, after {write}"));
                }
                kept_over += m.stores_kept;
                made_again += m.stores_moved;
            }
        }
        let down = vector::sliced::take()[vector::sliced::DOWN];
        assert!(kept > 0, "a frame's own sheet kept as an array");
        assert!(down > 0, "a column of a frame's numbers read whole");
        assert!(kept_over > 0, "an array kept over a write that plans again");
        assert!(made_again > 0, "an array moved past its extent");
    }

    #[test]
    fn a_range_over_a_frame_s_sheet_answers_a_cell_not_computed() {
        // KERNEL.31: a range of a frame's own sheet holding a cell not
        // computed answers the cell's reason, as a walk of its cells does, a
        // run's fold and a cell alone's walk alike: since KERNEL.49 the reason
        // is a kind among its sheet's slots, the array offered all the while,
        // where the sheet stepped its array aside while a reason stood.
        // Beside them a twin's range skips the cell, which stands outside
        // the rectangle of the twin's cells, its extent, to which a walk of
        // the twin is clamped as `Calc`'s walk of a sheet is (KERNEL.49).
        let rows = concat!(
            "(sheet \"S\" visible)\n",
            "(cell \"S\" \"B1:B2\" 1)\n",
            "(formula \"S\" \"A1\" \"=LABEL(B1:B2)\")\n",
            "(formula \"S\" \"C1:C8\" \"=SUM(A$1:B$1)+ROW()\")\n",
            "(formula \"S\" \"D1\" \"=SUM(A1:B2)\")\n",
            "(formula \"S\" \"E1:E8\" \"=SUM(S.last!A$1:B$1)\")\n",
        );
        let (mut m, mut cells) = kept_and_not(rows);
        assert!(Env::dense(&m, 0).is_some(), "the sheet kept as an array");
        for frame in 1..=2 {
            assert!(m.step(0).done && cells.step(0).done);
            assert_eq!(values(&m), values(&cells), "frame {frame}");
            store_is_the_frame(&m, &format!("frame {frame}"));
        }
        for addr in ["A1", "C1", "C8", "D1"] {
            assert!(
                matches!(at(&m, "S", addr), Some(Computed::NotComputed(_))),
                "{addr}: {:?}",
                at(&m, "S", addr)
            );
        }
        assert_eq!(at(&m, "S", "E1"), Some(Computed::Value(Value::Number(1.0))));
        // `A1`'s reason, and the nine cells that answer it, among the slots
        // of an array offered all the while.
        assert_eq!(reasons_in(&m, 0), 10);
        assert!(Env::dense(&m, 0).is_some(), "the array offered");
        // A1 computed: no reason stands.
        let w = "(formula \"S\" \"A1\" \"=B1+B2\")";
        assert!(m.write(w).is_ok() && cells.write(w).is_ok());
        assert!(m.step(0).done && cells.step(0).done);
        assert!(m.step(0).done && cells.step(0).done);
        assert_eq!(values(&m), values(&cells));
        assert_eq!(reasons_in(&m, 0), 0);
        assert_eq!(at(&m, "S", "C1"), Some(Computed::Value(Value::Number(4.0))));
        assert_eq!(at(&m, "S", "D1"), Some(Computed::Value(Value::Number(4.0))));
        store_is_the_frame(&m, "after the reason");
    }

    #[test]
    fn a_twin_s_extent_is_the_rectangle_of_its_cells_to_every_reader() {
        // KERNEL.49, found by KERNEL.56's law probe (2026-10-22): an
        // approximate MATCH bisects a twin's whole column over the twin's
        // extent, the rectangle of its cells, as `Calc` bisects the twin read
        // as a sheet; a write that grows the sheet grows its twin's extent at
        // the next frame's end and not before; and a formula's reason past the
        // twin's cells, the twin holding no cell for it, is skipped by every
        // walk of a range over the twin, clamped to its extent as `Calc`'s
        // walk of a sheet is: a fill's range moving with its cells folded at
        // each by a run, and one held computed once, a cell alone's program,
        // the criteria's and the lookups' walks, and every cell computed by
        // `Ctx`. Each frame held to `Calc` over the grid lent as the frame
        // began, its twin made from its sheet, the twin kept as an array and,
        // its sheet made sparse by a cell far below, as a sparse sheet's store,
        // the reason standing past the twin's cells in both: a reason inside
        // them a range meets as its sheet's range does (KERNEL.8's second
        // slice), where `Calc`, reading the twin as a sheet, meets none.
        let formulas = [
            ("B2", "=MATCH(4,S.last!A:A,1)"),
            ("B3", "=COUNT(S.last!A:A)+MATCH(1,S.last!1:1,1)"),
            ("B4", "=SUM(S.last!A1:F10)"),
            ("B5", "=COUNTIF(S.last!A1:F10,\">0\")"),
            ("B6", "=MATCH(3,S.last!A1:A10,0)"),
            ("D1:D4", "=SUM(S.last!$A$1:$F$10)+ROW()"),
            ("C1:C6", "=SUM(S.last!A1:E10)"),
        ];
        for far in ["", "(cell \"S\" \"A900\" 1)\n"] {
            // As the machine computes them, by programs and runs; by runs alone,
            // so that a run of fewer cells than a program takes is folded; and
            // every cell by `Ctx`.
            for (lowered, runs) in [(true, true), (false, true), (false, false)] {
                let mut rows = String::from(concat!(
                    "(cell \"S\" \"A1\" 1)\n(cell \"S\" \"A2\" 5)\n",
                    "(cell \"S\" \"A3\" 2)\n(cell \"S\" \"A4\" 3)\n",
                ));
                for (at, f) in formulas {
                    let text = crate::rows::quoted(f);
                    rows.push_str(&format!("(formula \"S\" \"{at}\" {text})\n"));
                }
                rows.push_str(far);
                let mut m = load(&rows);
                (m.lowered, m.by_runs) = (lowered, runs);
                assert_eq!(store_kept(&m, "S"), Some(far.is_empty()), "{far}");
                for frame in 1..=8 {
                    match frame {
                        4 => assert_eq!(m.write("(cell \"S\" \"C8\" 1)"), Ok(1)),
                        6 => assert_eq!(m.write("(formula \"S\" \"E9\" \"=LABEL(A1:A2)\")"), Ok(1)),
                        _ => {}
                    }
                    let (book, lib) = (m.workbook(), labelling());
                    let calc = crate::calc::Calc::run(&book, &lib);
                    assert!(m.step(0).done);
                    for id in m.shapes.order_cells() {
                        assert_eq!(
                            m.computed(id),
                            calc.computed(id).cloned(),
                            "frame {frame}, programs {lowered}, runs {runs}, {id:?}\n{rows}"
                        );
                    }
                }
            }
        }
    }

    #[test]
    fn a_frame_s_array_keeps_a_cell_not_reached_and_takes_each_write() {
        // KERNEL.31: a frame computes what changed, so a slot whose cell it
        // does not reach keeps the last frame's value; a value written lands
        // in its slot at the write, before any frame; a formula written
        // inside the sheet's extent plans again and keeps the array, its slot
        // nothing until its first frame; a cell written past the extent takes
        // the array over the new one, its 24 slots moved to their places, and
        // a cell below it rows added in place, nothing moved (KERNEL.56),
        // where each made the array again.
        let rows = concat!(
            "(sheet \"S\" visible)\n",
            "(cell \"S\" \"A1:A8\" 1)\n",
            "(formula \"S\" \"B1:B8\" \"=A1*2\")\n",
            "(formula \"S\" \"C1:C8\" \"=B1+1\")\n",
        );
        let mut m = load(rows);
        let slot =
            |m: &Machine, row: u32, col: u32| m.stores[0].array().and_then(|d| d.get(row, col));
        assert!(m.step(0).done);
        store_is_the_frame(&m, "frame 1");
        assert!(m.write("(cell \"S\" \"A3\" 5)").is_ok());
        assert_eq!(
            slot(&m, 3, 1),
            Some(Value::Number(5.0)),
            "a value at its write"
        );
        store_is_the_frame(&m, "after a value");
        assert!(m.step(0).done);
        assert_eq!(slot(&m, 3, 2), Some(Value::Number(10.0)));
        assert_eq!(slot(&m, 3, 3), Some(Value::Number(11.0)));
        assert_eq!(
            slot(&m, 1, 2),
            Some(Value::Number(2.0)),
            "a cell not reached"
        );
        store_is_the_frame(&m, "frame 2");
        let first = |m: &Machine| m.stores[0].array().map(Dense::memory);
        let was = first(&m);
        assert!(m.write("(formula \"S\" \"B4\" \"=A4*3\")").is_ok());
        assert_eq!(first(&m), was, "the array kept, in its own memory");
        assert_eq!((m.stores_kept, m.stores_moved), (1, 0));
        assert_eq!(
            slot(&m, 4, 2),
            Some(Value::Empty),
            "a formula with no value yet"
        );
        store_is_the_frame(&m, "after a formula");
        assert!(m.step(0).done);
        assert_eq!(slot(&m, 4, 2), Some(Value::Number(3.0)));
        assert_eq!(slot(&m, 4, 3), Some(Value::Number(4.0)));
        store_is_the_frame(&m, "frame 3");
        let work = work_of(&mut m, "(cell \"S\" \"E12\" 1)");
        assert_eq!((m.stores_kept, m.stores_moved), (1, 1));
        assert_eq!(work.relaid, 24, "eight rows of three moved");
        assert_eq!(
            slot(&m, 12, 5),
            Some(Value::Number(1.0)),
            "over the new extent"
        );
        store_is_the_frame(&m, "after a cell past the extent");
        assert!(m.step(0).done);
        store_is_the_frame(&m, "frame 4");
        let work = work_of(&mut m, "(cell \"S\" \"A13\" 2)");
        assert_eq!((m.stores_kept, m.stores_moved, work.relaid), (1, 2, 0));
        assert_eq!(slot(&m, 13, 1), Some(Value::Number(2.0)), "a row below");
        store_is_the_frame(&m, "after a cell below the extent");
        assert!(m.step(0).done);
        store_is_the_frame(&m, "frame 5");
    }

    #[test]
    fn a_sheet_keeps_an_array_where_its_cells_stand_close_enough() {
        // KERNEL.49: a sheet whose cells only cells alone read, a running
        // balance's, keeps its values as an array all the same, and so does
        // one read only by cells every `$` holds, where KERNEL.31 kept an
        // array only for a sheet the plan read by a range or a run's moving
        // cell; one too sparse for an array keeps a map of its cells; and a
        // machine offering no arrays reads the same store cell by cell.
        for (rows, array) in [
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
        ]
        .into_iter()
        .zip([true, true, false])
        {
            let mut m = load(rows);
            assert!(m.step(0).done);
            assert_eq!(m.stores[0].array().is_some(), array, "{rows}");
            store_is_the_frame(&m, rows);
        }
        let (kept, cells) = kept_and_not(concat!(
            "(sheet \"S\" visible)\n",
            "(cell \"S\" \"A1:A8\" 1)\n",
            "(formula \"S\" \"B1:B8\" \"=SUM(A1:A2)\")\n",
        ));
        assert!(Env::dense(&kept, 0).is_some() && Env::dense(&cells, 0).is_none());
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
            store_is_the_frame(&m, &format!("frame {frame}"));
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
        for s in 0..m.book.pages.len() {
            let share = by_sheet.get(&s).copied().unwrap_or(0);
            assert_eq!(m.changed_on(s), Some(share), "{when}: sheet {s}'s count");
        }
        let past = m.book.pages.len();
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
                let of = machines[1].shapes.len();
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
            "(formula \"S\" \"B4\" \"=IF(A2>5,LABEL(A1:A2),LABEL(A2:A3))\")\n",
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

    // ---- KERNEL.8: the vectors' indexes ----

    #[test]
    fn an_index_answers_as_its_scan_frame_after_frame() {
        // KERNEL.8: the generated grids, lookups and criteria among their
        // formulas, every other one churning with the Clock written before
        // every frame, stepped in slices of any budget with writes between
        // frames to the value cells the vectors hold and to the bands, against
        // the same grids searching by the scan alone, frame for frame and bit
        // for bit, every probe an index answers held to its scan as it
        // answers; indexes built, and reached by changes, many times over.
        let _checked = Checked::on();
        let (mut builds, mut reached, mut answered) = (0, 0, 0);
        for seed in 1..=60u64 {
            let mut d = Dice(seed ^ 0x0808);
            let churns = seed % 2 == 0;
            let mut rows = if churns {
                churning_grid(&mut d)
            } else {
                generated_grid(&mut d)
            };
            // Bands searching the first sheet's value columns and formulas,
            // every vector held whole, so that each grid's indexes are built
            // and reached; since slice e the criteria's family among them,
            // a sum range on the bands' own sheet, a second equality's range
            // and a test by size after an equality, which the group's places
            // are tested by.
            rows.push_str(concat!(
                "(formula \"B\" \"M1:M20\" \"=MATCH(A!$A1,A!$A$1:$A$20,0)\")\n",
                "(formula \"B\" \"N1:N20\" \"=COUNTIF(A!$B$1:$B$20,$A1)\")\n",
                "(formula \"B\" \"O1:O20\" \"=SUMIF(A!$A$1:$A$20,$B1,A!$B$1:$B$20)\")\n",
                "(formula \"B\" \"P1:P20\" \"=XLOOKUP($A1,A!$C$1:$C$20,A!$D$1:$D$20,\\\"none\\\",0,-1)\")\n",
                "(formula \"B\" \"Q1:Q20\" \"=IFERROR(VLOOKUP(A!$B1,A!$C$1:$E$20,2,FALSE),-1)\")\n",
                "(formula \"B\" \"R1:R20\" \"=SUMIFS(B!$B$1:$B$20,A!$A$1:$A$20,$A1,A!$C$1:$C$20,$B1)\")\n",
                "(formula \"B\" \"S1:S20\" \"=COUNTIFS(A!$B$1:$B$20,$B1,A!$A$1:$A$20,$A1)\")\n",
                "(formula \"B\" \"T1:T20\" \"=AVERAGEIFS(A!$A$1:$A$20,A!$B$1:$B$20,$A1,A!$C$1:$C$20,\\\">1\\\")\")\n",
            ));
            let mut keyed = load(&rows);
            let mut scanned = load(&rows);
            scanned.indexed = false;
            for frame in 1..=6 {
                let when = format!("seed {seed}, frame {frame}");
                if churns {
                    let w = format!("(cell \"Clock\" \"B1\" {frame})");
                    assert!(keyed.write(&w).is_ok() && scanned.write(&w).is_ok());
                }
                while !keyed.step(1 + d.below(300)).done {}
                assert!(scanned.step(0).done);
                assert_eq!(values(&keyed), values(&scanned), "{when}\n{rows}");
                let write = match d.below(5) {
                    0 | 1 => format!(
                        "(cell \"{}\" \"{}{}\" {})",
                        d.pick(&["A", "B"]),
                        d.pick(&["A", "B"]),
                        1 + d.below(20),
                        d.pick(&[
                            "7",
                            "\"x\"",
                            "\"X\"",
                            "true",
                            "0",
                            "\"5\"",
                            "(error \"#N/A\")"
                        ])
                    ),
                    2 => format!(
                        "(cell \"A\" \"{}{}\" 2)",
                        d.pick(&LETTERS[2..8]),
                        1 + d.below(4)
                    ),
                    _ => continue,
                };
                let (a, b) = (keyed.write(&write), scanned.write(&write));
                assert!(a.is_ok() == b.is_ok(), "{when}: {write}: {a:?} {b:?}");
            }
            let (b, r, a) = keyed.indexes.counted();
            builds += b;
            reached += r;
            answered += a;
            assert_eq!(
                scanned.indexes.counted(),
                (0, 0, 0),
                "seed {seed}: no index unasked"
            );
        }
        assert!(
            builds >= 600 && reached >= 200 && answered > builds * 20,
            "{builds} indexes built, {reached} reached by a change and {answered} probes \
             answered: too few for the law to hold over"
        );
    }

    #[test]
    fn an_index_is_built_at_a_second_probe_and_kept_until_its_vector_changes() {
        // KERNEL.8: a vector probed by one cell is never built; by two, once,
        // at the second probe; kept from frame to frame while nothing of its
        // vector changes, a write beside it among them; reached by a write to
        // a cell of it and by a change of the formula cell it holds, and
        // built again at its next probe; and dropped with every other when a
        // formula is written.
        let mut m = load(concat!(
            "(sheet \"S\" visible)\n",
            "(cell \"S\" \"A1:A2\" 1)\n",
            "(cell \"S\" \"A3\" 5)\n",
            "(cell \"S\" \"A4:A6\" 1)\n",
            "(cell \"S\" \"B1\" 5)\n",
            "(cell \"S\" \"D1\" 0)\n",
            "(formula \"S\" \"A7\" \"=D1*2\")\n",
            "(formula \"S\" \"C1\" \"=MATCH(5,$A$1:$A$7,0)\")\n",
            "(formula \"S\" \"C2\" \"=MATCH(B1,$A$1:$A$7,0)\")\n",
            "(formula \"S\" \"C3\" \"=COUNTIF($A$1:$A$9,1)\")\n",
            "(formula \"S\" \"C4\" \"=MATCH(1,$A$1:$A$7,1)\")\n",
            "(formula \"S\" \"E1\" \"=D1+1\")\n",
        ));
        let counted = |m: &Machine| {
            let (builds, reached, _) = m.indexes.counted();
            (m.indexes.built(), builds, reached)
        };
        let answered = |m: &Machine| m.indexes.counted().2;
        // `$A$1:$A$7` searched exactly twice, built; the count's range once,
        // never; the bisection keys nothing.
        assert!(m.step(0).done);
        assert_eq!(counted(&m), (1, 1, 0), "the first frame");
        assert_eq!(at(&m, "S", "C1"), Some(Computed::Value(Value::Number(3.0))));
        assert_eq!(at(&m, "S", "C3"), Some(Computed::Value(Value::Number(5.0))));
        assert!(m.step(0).done);
        assert_eq!(counted(&m), (1, 1, 0), "kept while nothing changes");
        // A value beside the vector, which a probe reads.
        let before = answered(&m);
        assert!(m.write("(cell \"S\" \"B1\" 1)").is_ok());
        assert!(m.step(0).done);
        assert_eq!(counted(&m), (1, 1, 0), "kept past a write beside it");
        assert!(answered(&m) > before, "the probe answered by the index");
        assert_eq!(at(&m, "S", "C2"), Some(Computed::Value(Value::Number(1.0))));
        // A value of the vector: reached at the write, built again at its
        // next probe.
        assert!(m.write("(cell \"S\" \"A1\" 5)").is_ok());
        assert_eq!(counted(&m), (0, 1, 1), "reached at the write");
        assert!(m.step(0).done);
        assert_eq!(counted(&m), (1, 2, 1), "built again at its next probe");
        assert_eq!(at(&m, "S", "C1"), Some(Computed::Value(Value::Number(1.0))));
        // The formula cell of the vector, its value changed by a write it
        // reads.
        assert!(m.write("(cell \"S\" \"D1\" 1)").is_ok());
        assert!(m.step(0).done);
        assert_eq!(counted(&m), (1, 3, 2), "reached by a computed value");
        assert_eq!(at(&m, "S", "C3"), Some(Computed::Value(Value::Number(4.0))));
        // A formula written: every index dropped with the plan.
        assert!(m.write("(formula \"S\" \"F1\" \"=1\")").is_ok());
        assert_eq!(counted(&m).0, 0, "dropped with the plan");
        assert!(m.step(0).done);
        assert_eq!(counted(&m), (1, 4, 2), "built again at the second probe");
    }

    #[test]
    fn a_plan_through_hubs_is_the_plan_of_every_edge() {
        // KERNEL.8: the generated grids, and the bands the index's law adds,
        // planned with a hub for every shape's held references and with none,
        // at the load and after every write that plans again: the two plans
        // spelled whole alike (each shape, the order, its cycles, runs and
        // levels, the reach), and their frames bit for bit. Since KERNEL.49's
        // decision 7 a hub stands in the plan of cells, every plan of the
        // tests' reference (`shape::reference`), and among the cells a plan by
        // nodes plans cell by cell: two fills reading each other's rows, one
        // of them a held cell of the other, are one component planned cell by
        // cell, its held read through a hub. Each plan is held to the plan of
        // cells, and both count their hubs.
        use crate::calc::shape::{hubs_at, hubs_made};
        let planned = |rows: &str, hubs: usize| {
            hubs_at(hubs);
            let m = load(rows);
            hubs_at(64);
            m
        };
        let written = |m: &mut Machine, w: &str, hubs: usize| {
            hubs_at(hubs);
            let ok = m.write(w).is_ok();
            hubs_at(64);
            ok
        };
        hubs_made();
        let mut made = 0;
        for seed in 1..=60u64 {
            let mut d = Dice(seed ^ 0x4B8);
            let mut rows = generated_grid(&mut d);
            rows.push_str(concat!(
                "(formula \"B\" \"M1:M20\" \"=MATCH(A!$A1,A!$C$1:$C$20,0)\")\n",
                "(formula \"B\" \"N1:N20\" \"=SUM(A!$C$1:$H$4)+$A1\")\n",
                "(formula \"B\" \"P2:P30\" \"=Q1+$Q$1\")\n",
                "(formula \"B\" \"Q1:Q30\" \"=P1+1\")\n",
            ));
            let mut with = planned(&rows, 1);
            made += hubs_made();
            let mut without = planned(&rows, usize::MAX);
            assert_eq!(hubs_made(), 0, "seed {seed}: no hub at usize::MAX");
            for frame in 1..=4 {
                let when = format!("seed {seed}, frame {frame}");
                assert_eq!(
                    with.shapes.spelled(),
                    without.shapes.spelled(),
                    "{when}: the plans\n{rows}"
                );
                assert!(with.step(0).done && without.step(0).done);
                assert_eq!(values(&with), values(&without), "{when}\n{rows}");
                let w = match d.below(3) {
                    0 => format!("(cell \"A\" \"A{}\" 7)", 1 + d.below(20)),
                    1 => format!(
                        "(cell \"A\" \"{}{}\" 2)",
                        d.pick(&LETTERS[2..8]),
                        1 + d.below(4)
                    ),
                    _ => "(formula \"B\" \"C1:H2\" \"=$A1+1\")".to_string(),
                };
                let (a, b) = (
                    written(&mut with, &w, 1),
                    written(&mut without, &w, usize::MAX),
                );
                assert_eq!(a, b, "{when}: {w}");
                made += hubs_made();
            }
        }
        assert!(
            made > 200,
            "{made} hubs made: too few for the law to hold over"
        );
    }

    #[test]
    fn a_plan_by_nodes_plans_cell_by_cell_what_it_must_and_nothing_else() {
        // KERNEL.49's decision 7, by name: a fill reading itself (a chain, a
        // running balance), two fills reading each other's rows (one
        // component of two nodes), a fill reading a defined name and one
        // whose text the parse does not read are planned cell by cell, a
        // chain's order and a cycle's cells and its refusal the plan of
        // cells' to the cell (every plan held to it, `shape::reference`);
        // fills reading other fills, value cells and held cells, and
        // formulas alone of one text side by side, are planned by their
        // rectangles, no cell planned alone. The held read of a component
        // planned cell by cell is read through a hub when hubs stand from one
        // cell, the plan the same as with none.
        use crate::calc::shape::{hubs_at, hubs_made, planned_by_cell};
        planned_by_cell();
        let by_rects = load(concat!(
            "(cell \"S\" \"A1:A20\" 2)\n",
            "(formula \"S\" \"B1:D20\" \"=$A1*2+$A$1\")\n",
            "(formula \"S\" \"E1:E20\" \"=SUM(B1:D1)+B$1\")\n",
            "(formula \"S\" \"F1\" \"=E1+1\")\n",
            "(formula \"S\" \"G1\" \"=F1+1\")\n",
            "(formula \"S\" \"F2\" \"=E2+1\")\n",
            "(formula \"S\" \"G2\" \"=F2+1\")\n",
        ));
        assert_eq!(
            planned_by_cell(),
            0,
            "fills and formulas alone by rectangles"
        );
        assert_eq!(by_rects.shapes.formulas(), 84);
        let chain = load(concat!(
            "(cell \"S\" \"A1\" 1)\n",
            "(formula \"S\" \"A2:A40\" \"=MOD(ROW()*7,13)\")\n",
            "(formula \"S\" \"B1\" \"=A1\")\n",
            "(formula \"S\" \"B2:B40\" \"=B1+A2\")\n",
        ));
        assert_eq!(planned_by_cell(), 39, "the running balance cell by cell");
        assert_eq!(chain.shapes.level_ends().len(), 40, "a level a balance");
        let named = load(concat!(
            "(name \"Rate\" \"S!$A$1\")\n",
            "(cell \"S\" \"A1\" 3)\n",
            "(formula \"S\" \"B1:B5\" \"=Rate*ROW()\")\n",
            "(formula \"S\" \"C1:C5\" \"=B1+1\")\n",
        ));
        assert_eq!(planned_by_cell(), 5, "the name's readers cell by cell");
        assert_eq!(named.shapes.formulas(), 10);
        let rows = concat!(
            "(formula \"S\" \"P2:P30\" \"=Q1+$Q$1\")\n",
            "(formula \"S\" \"Q1:Q30\" \"=P1+1\")\n",
        );
        hubs_made();
        let (made, through) = {
            let _unheld = crate::calc::shape::reference::unheld();
            hubs_at(1);
            let m = load(rows);
            hubs_at(64);
            (hubs_made(), m)
        };
        assert_eq!(planned_by_cell(), 59, "one component of two nodes");
        assert!(made >= 1, "the held cell read through a hub");
        let mut direct = load(rows);
        assert_eq!(hubs_made(), 0, "no hub of one cell at 64");
        assert_eq!(planned_by_cell(), 59);
        assert_eq!(through.shapes.spelled(), direct.shapes.spelled());
        assert!(direct.step(0).done);
        assert_eq!(
            direct.computed((0, 30, 17)),
            Some(Computed::Value(Value::Number(59.0)))
        );
        let cycle = Machine::load(
            concat!(
                "(formula \"S\" \"A1:A3\" \"=B1+1\")\n",
                "(formula \"S\" \"B1:B3\" \"=A1\")\n",
                "(formula \"S\" \"C1:C3\" \"=A1\")\n",
            ),
            labelling(),
        )
        .expect_err("a cycle refused");
        assert_eq!(planned_by_cell(), 6, "the cycle's two nodes cell by cell");
        assert_eq!(cycle.refusal.id, "calc-cycle");
        assert!(
            cycle.refusal.text.contains("S!A1, S!B1"),
            "{}",
            cycle.refusal.text
        );
        let unread = Machine::load("(formula \"S\" \"A1:A3\" \"=1+\")\n", labelling())
            .expect_err("a text not read refused");
        assert_eq!(planned_by_cell(), 3, "a text not read cell by cell");
        assert_eq!(unread.refusal.id, "calc-construct-not-read");
        // A span of sheets read at a fill's first cell and moved off the
        // sheet at its second: the second's text not read, the load refused
        // there, where the first reads.
        let span = Machine::load(
            concat!(
                "(sheet \"Jan\" visible)\n(sheet \"Dec\" visible)\n",
                "(formula \"S\" \"A1:A2\" \"=SUM(Jan:Dec!B1048576)\")\n",
            ),
            labelling(),
        )
        .expect_err("a span moved off the sheet refused");
        assert_eq!(span.refusal.id, "calc-construct-not-read");
        assert!(span.refusal.text.contains("S!A2"), "{}", span.refusal.text);
        // Two nodes calling functions the library lacks: the first named.
        let two = Machine::load(
            concat!(
                "(formula \"S\" \"C3:D4\" \"=NOSUCHTWO(1)\")\n",
                "(formula \"S\" \"B2\" \"=NOSUCHONE(1)\")\n",
            ),
            labelling(),
        )
        .expect_err("a function not held refused");
        assert_eq!(two.refusal.id, "calc-function-not-computed");
        assert!(two.refusal.text.contains("S!B2"), "{}", two.refusal.text);
    }

    #[test]
    fn runs_carried_across_a_shape_s_levels_are_the_plan_of_cells_runs() {
        // KERNEL.49's decision 7: seeded grids where one fill's pieces stand
        // at several levels of one sheet, so that the blocks of its shape
        // follow one another in the order and a run is carried from one into
        // the next and broken inside it: a fill reading up, left or up and
        // left, cut into pieces by values, constants and formulas alone of its
        // own text written into it, a column of values and constants beside
        // it. Each grid is planned by nodes and held to the plan of cells
        // (`shape::reference`), runs and all, at its load and after every
        // write; its frames stepped.
        let texts = ["=A1+1", "=B1*2", "=A2+1", "=$A2+B1", "=A1+B1"];
        let mut carried = 0;
        for seed in 1..=300u64 {
            let mut d = Dice(seed ^ 0x4904_0004);
            let mut rows = String::new();
            for r in 1..=9 {
                match d.below(3) {
                    0 => rows.push_str(&format!("(cell \"S\" \"A{r}\" {})\n", d.below(5))),
                    1 => rows.push_str(&format!("(formula \"S\" \"A{r}\" \"=1\")\n")),
                    _ => {}
                }
            }
            let text = d.pick(&texts);
            let (bottom, right) = (4 + d.below(6), 3 + d.below(5));
            rows.push_str(&format!(
                "(formula \"S\" \"B2:{}{bottom}\" \"{text}\")\n",
                LETTERS[right - 1]
            ));
            let mut m = load(&rows);
            for _ in 0..8 {
                let (r, c) = (2 + d.below(bottom - 1), 2 + d.below(right - 1));
                let at = format!("{}{r}", LETTERS[c - 1]);
                let w = match d.below(4) {
                    0 => format!("(cell \"S\" \"{at}\" {})", d.below(9)),
                    1 => format!("(formula \"S\" \"{at}\" \"=1\")"),
                    2 => {
                        let moved = shift_a1_references(text, r as i64 - 2, c as i64 - 2);
                        format!("(formula \"S\" \"{at}\" \"{moved}\")")
                    }
                    _ => {
                        let r2 = (r + d.below(3)).min(bottom);
                        format!("(cell \"S\" \"{at}:{}{r2}\" 0)", LETTERS[c - 1])
                    }
                };
                if m.write(&w).is_ok() {
                    assert!(m.step(0).done, "seed {seed}: {w}");
                }
                carried += m.shapes.run_spans().count();
            }
        }
        assert!(
            carried > 1_000,
            "{carried} runs: too few for the law to hold over"
        );
    }

    // ---- KERNEL.56: a value written into an empty cell ----

    /// Each formula cell's value as `Calc` computes it over the grid lent as
    /// the frame begins, its twins as sheets (KERNEL.34's oracle): what the
    /// frame computes, where no twin holds a reason, which `Calc` reading a
    /// twin as a sheet would not see.
    fn calc_of(m: &Machine) -> BTreeMap<CellId, String> {
        let book = m.workbook();
        let calc = crate::calc::Calc::run(&book, &m.library);
        m.shapes
            .order_cells()
            .iter()
            .map(|&id| (id, format!("{:?}", calc.computed(id).expect("computed"))))
            .collect()
    }

    /// One frame stepped whole: the cells it computed, and what it and its
    /// end did.
    fn frame_of(m: &mut Machine) -> (usize, crate::sheet::work::Work) {
        let work = ended(m, 0);
        (m.computed, work)
    }

    #[test]
    fn the_extent_s_readers_compute_again_where_a_write_moved_the_extent() {
        // KERNEL.56: an approximate match bisects its vector between the
        // first and last places of its sheet's extent along it, and a
        // database's records end at its sheet's last row, so a value written
        // beyond the extent, which no such formula reads by place, moves
        // their answers: over a column of 1, 5, 2 and 3, `MATCH(4,A:A,1)` 1
        // to 4 and `LOOKUP(4,A:A)` 1 to 3; over a database of a name and two
        // records, its criteria a name and a blank row, `DCOUNTA` 2 to 7;
        // and the same over a twin's column and database a frame later, as
        // the twin's extent grows at the frame's end; `SUM(C:C)` reads the
        // cell by its place. The first frame after the write computes the
        // sheets' readers of their extents and the cell's reader, five
        // cells (the twin's database's among them, reading its criteria on
        // the sheet), and the second the twins' readers, two; a value written
        // inside the extent moves no reader of it. Every frame held to
        // `Calc` over the grid as it began and to the grid computing every
        // cell. The extent's readers were computed again only by the whole
        // frame a cell added made; no test held them.
        let rows = concat!(
            "(cell \"S\" \"A1\" 1)\n(cell \"S\" \"A2\" 5)\n",
            "(cell \"S\" \"A3\" 2)\n(cell \"S\" \"A4\" 3)\n",
            "(formula \"S\" \"B1\" \"=MATCH(4,A:A,1)\")\n",
            "(formula \"S\" \"B2\" \"=MATCH(4,S.last!A:A,1)\")\n",
            "(formula \"S\" \"B3\" \"=SUM(C:C)\")\n",
            "(formula \"S\" \"B4\" \"=LOOKUP(4,A:A)\")\n",
            "(cell \"T\" \"A1\" \"h\")\n(cell \"T\" \"A2\" 1)\n",
            "(cell \"T\" \"A3\" 2)\n(cell \"T\" \"D1\" \"h\")\n",
            "(formula \"T\" \"E1\" \"=DCOUNTA(A:A,,D1:D2)\")\n",
            "(formula \"T\" \"E2\" \"=DCOUNTA(T.last!A:A,,D1:D2)\")\n",
        );
        let mut m = load(rows);
        let mut whole = load(rows);
        whole.incremental = false;
        let cells = ["S!B1", "S!B2", "S!B3", "S!B4", "T!E1", "T!E2"];
        let shown = |m: &Machine| -> Vec<f64> {
            cells
                .iter()
                .map(|c| c.split_once('!').expect("a sheet's cell"))
                .map(|(s, a)| match at(m, s, a) {
                    Some(Computed::Value(Value::Number(n))) => n,
                    other => panic!("{s}!{a}: {other:?}"),
                })
                .collect()
        };
        // A frame held to `Calc` and to every cell computed; the cells it
        // computed.
        fn step(m: &mut Machine, whole: &mut Machine, when: &str) -> usize {
            let calc = calc_of(m);
            let (computed, _) = frame_of(m);
            assert!(whole.step(0).done);
            assert_eq!(values(m), calc, "{when}: Calc");
            assert_eq!(values(m), values(whole), "{when}: computed whole");
            computed
        }
        for frame in 1..=3 {
            step(&mut m, &mut whole, &format!("frame {frame}"));
        }
        assert_eq!(
            shown(&m),
            [1.0, 1.0, 0.0, 1.0, 2.0, 2.0],
            "before the write"
        );
        let w = "(cell \"S\" \"C8\" 1)\n(cell \"T\" \"C8\" 1)";
        assert!(m.write(w).is_ok() && whole.write(w).is_ok());
        assert_eq!(
            step(&mut m, &mut whole, "after the write"),
            5,
            "B1, B3, B4, E1 and E2"
        );
        assert_eq!(
            shown(&m),
            [4.0, 1.0, 1.0, 3.0, 7.0, 2.0],
            "the sheets' readers"
        );
        assert_eq!(step(&mut m, &mut whole, "a frame later"), 2, "B2 and E2");
        assert_eq!(
            shown(&m),
            [4.0, 4.0, 1.0, 3.0, 7.0, 7.0],
            "the twins' readers"
        );
        assert_eq!(step(&mut m, &mut whole, "still"), 0);
        // A hole filled: its reader by place alone, at either frame.
        let w = "(cell \"S\" \"C2\" 2)";
        assert!(m.write(w).is_ok() && whole.write(w).is_ok());
        assert_eq!(step(&mut m, &mut whole, "after a hole"), 1, "B3");
        assert_eq!(step(&mut m, &mut whole, "a frame after a hole"), 0);
        assert_eq!(shown(&m), [4.0, 4.0, 3.0, 3.0, 7.0, 7.0], "SUM(C:C)");
    }

    #[test]
    fn a_value_written_into_an_empty_cell_costs_its_readers_and_not_the_grid() {
        // KERNEL.56, counted and not timed: a value written into an empty
        // cell at each placement, a hole inside the extent, a row below it
        // and a column to its right, of a sheet kept as an array and viewed
        // as a plane and of a sparse one, each sheet's twin read by a
        // formula; one formula reading every cell written, and on each sheet
        // and its twin a reader of the extent. The next two frames compute
        // the cell's reader and the readers of the extents its growth moved,
        // the sheet's at the first frame and the twin's at the second, and
        // nothing else; no write walks a sheet and no frame's end sweeps
        // one; a store and a plane whose bottom alone moved take rows in
        // place, nothing moved, and where a column was added to the right
        // the array's slots and the plane's bytes are moved at the write, the
        // twin's slots at the frame's end. A write that added a cell made
        // both frames compute every cell, refitted the plan, re-laid the
        // sheet's store, dropped its plane, and at the frame's end made the
        // twin's store again, its extent from a walk of its sheet.
        let rows = concat!(
            "(cell \"D\" \"A1:A100\" 1)\n",
            "(formula \"D\" \"B1:B100\" \"=A1*2\")\n",
            "(formula \"D\" \"C1:C100\" \"=D.last!B1+1\")\n",
            "(cell \"D\" \"E1:E100\" 2)\n",
            "(cell \"P\" \"A1:A20\" 3)\n",
            "(formula \"P\" \"B1:B20\" \"=A1+1\")\n",
            "(cell \"P\" \"Z5000\" 1)\n",
            "(formula \"R\" \"A1\" \"=SUM(D!D5,D!A101,D!F1,P!C10,P!A5001,P!AA1)\")\n",
            "(formula \"R\" \"A2\" \"=MATCH(1E+9,D!A:A,1)\")\n",
            "(formula \"R\" \"A3\" \"=MATCH(1E+9,D.last!A:A,1)\")\n",
            "(formula \"R\" \"A4\" \"=MATCH(1E+9,P!A:A,1)\")\n",
            "(formula \"R\" \"A5\" \"=MATCH(1E+9,P.last!A:A,1)\")\n",
        );
        let mut m = load(rows);
        let mut whole = load(rows);
        whole.incremental = false;
        for x in [&m, &whole] {
            x.view("plane", "D", None).expect("a plane of D");
        }
        for _ in 0..3 {
            assert!(m.step(0).done && whole.step(0).done);
        }
        let (d, p) = (m.find_sheet("D").unwrap(), m.find_sheet("P").unwrap());
        assert!(m.stores[d].array().is_some(), "D an array");
        assert!(m.stores[p].array().is_none(), "P sparse");
        assert!(
            m.twins[d].is_some() && m.twins[p].is_some(),
            "both twins read"
        );
        assert!(kept(&m, "D"), "D's plane kept");
        // Each placement: the cell, then what its write moved, the cells the
        // next two frames computed, and what the first frame's end moved.
        let placements = [
            ("D", "D5", 0, (1, 0), 0),
            ("D", "A101", 0, (2, 1), 0),
            ("D", "F1", 2 * 505, (2, 1), 505),
            ("P", "C10", 0, (1, 0), 0),
            ("P", "A5001", 0, (2, 1), 0),
            ("P", "AA1", 0, (2, 1), 0),
        ];
        for (sheet, cell, relaid, (first, second), relaid_end) in placements {
            let when = format!("{sheet}!{cell}");
            let w = format!("(cell \"{sheet}\" \"{cell}\" 7)");
            let work = work_of(&mut m, &w);
            assert!(whole.write(&w).is_ok());
            assert_eq!(
                (work.walked, work.relaid),
                (0, relaid),
                "{when}: what the write walked and moved"
            );
            assert!(kept(&m, "D"), "{when}: D's plane kept");
            let (computed, work) = frame_of(&mut m);
            assert!(whole.step(0).done);
            same_state(&m, &whole, &when);
            assert_eq!(
                (work.swept, work.relaid),
                (0, relaid_end),
                "{when}: the first frame's end"
            );
            let (again, work) = frame_of(&mut m);
            assert!(whole.step(0).done);
            same_state(&m, &whole, &when);
            assert_eq!(
                (computed, again),
                (first, second),
                "{when}: the cells the two frames computed"
            );
            assert_eq!(
                (work.swept, work.relaid),
                (0, 0),
                "{when}: the second frame's end"
            );
        }
    }

    #[test]
    fn a_frame_after_cells_written_into_empty_places_is_the_frame_computed_whole() {
        // KERNEL.56: the law above (KERNEL.20, slice 4) over the same
        // generated grids beside a sheet of the extent's readers: approximate
        // lookups and database functions over whole columns of the grids'
        // value cells, of their twins and of a database of its own, each
        // reading a sheet's extent itself. Between frames, beside the law's
        // writes, values written into empty places, a row below a sheet's
        // extent and a column to its right, which grow the extent where no
        // reader reads the cell by place, and holes inside it. Every frame
        // held to the same grid computing every cell, and the extent's
        // readers, which read value cells alone, to `Calc` over the grid as
        // the frame began, twins as sheets.
        let readers = concat!(
            "(sheet \"X\" visible)\n",
            "(cell \"X\" \"F1\" \"h\")\n(cell \"X\" \"F2\" 1)\n(cell \"X\" \"F3\" 5)\n",
            "(cell \"X\" \"F4\" 2)\n(cell \"X\" \"F5\" 3)\n(cell \"X\" \"H1\" \"h\")\n",
            "(formula \"X\" \"A1\" \"=MATCH(3,A!$A:$A,1)\")\n",
            "(formula \"X\" \"A2\" \"=LOOKUP(2,A!A:A)\")\n",
            "(formula \"X\" \"A3\" \"=VLOOKUP(3,A!$A:$B,2)\")\n",
            "(formula \"X\" \"A4\" \"=XMATCH(3,B!$B:$B,-1,2)\")\n",
            "(formula \"X\" \"A5\" \"=XLOOKUP(3,A!$A:$A,A!$B:$B,\\\"none\\\",-1,2)\")\n",
            "(formula \"X\" \"A6\" \"=MATCH(3,A.last!$A:$A,-1)\")\n",
            "(formula \"X\" \"A7\" \"=LOOKUP(2,B.last!$B:$B)\")\n",
            "(formula \"X\" \"A8\" \"=DCOUNTA(F:F,,H1:H2)\")\n",
            "(formula \"X\" \"A9\" \"=DCOUNTA(X.last!F:F,,X.last!H1:H2)\")\n",
            "(formula \"X\" \"A10\" \"=MATCH(4,F:F,1)\")\n",
            "(formula \"X\" \"A11\" \"=DGET(F:F,1,H1:H2)\")\n",
            "(formula \"X\" \"A12\" \"=DSUM(F:F,1,H1:H2)\")\n",
        );
        let x_cells = |m: &Machine| -> Vec<CellId> {
            let x = m.find_sheet("X").expect("X");
            m.shapes
                .order_cells()
                .iter()
                .copied()
                .filter(|id| id.0 == x)
                .collect()
        };
        let (mut frames, mut followed, mut placed, mut moved) = (0, 0, 0, 0);
        for seed in 1..=60u64 {
            let mut d = Dice(seed ^ 0x56E5);
            let rows = generated_grid(&mut d) + readers;
            let mut inc = load(&rows);
            let mut whole = load(&rows);
            whole.incremental = false;
            for m in [&inc, &whole] {
                m.view("plane", "A", None).expect("a plane of A");
            }
            let x = x_cells(&inc);
            for frame in 1..=10u32 {
                let when = format!("seed {seed}, frame {frame}");
                let calc = calc_of(&inc);
                let before: Vec<String> = x
                    .iter()
                    .map(|id| format!("{:?}", inc.computed(*id)))
                    .collect();
                let budget = 1 + d.below(200);
                while !inc.step(budget).done {}
                assert!(whole.step(0).done);
                same_state(&inc, &whole, &format!("{when}\n{rows}"));
                for (k, id) in x.iter().enumerate() {
                    let now = format!("{:?}", inc.computed(*id));
                    let want = calc.get(id).map(|c| format!("Some({c})"));
                    assert_eq!(
                        Some(&now),
                        want.as_ref(),
                        "{when}: {id:?} against Calc\n{rows}"
                    );
                    moved += usize::from(!inc.every && now != before[k]);
                }
                frames += 1;
                followed += usize::from(!inc.every);
                // A value into an empty place, or the law's own write.
                let write = match d.below(4) {
                    0 => {
                        placed += 1;
                        let sheet = d.pick(&["A", "B", "X"]);
                        match d.below(3) {
                            0 => format!(
                                "(cell \"{sheet}\" \"A{}\" {})",
                                70 + d.below(30),
                                d.pick(&["1", "3", "4"])
                            ),
                            1 => format!(
                                "(cell \"{sheet}\" \"{}{}\" 2)",
                                d.pick(&["M", "P", "T"]),
                                1 + d.below(20)
                            ),
                            _ => format!("(cell \"{sheet}\" \"I{}\" 5)", 1 + d.below(20)),
                        }
                    }
                    _ => match write_between(&mut d, frame) {
                        Some(w) => w,
                        None => continue,
                    },
                };
                let (a, b) = (inc.write(&write), whole.write(&write));
                assert_eq!(a.is_ok(), b.is_ok(), "{when}: {write}: {a:?} {b:?}");
            }
        }
        assert!(
            followed * 2 > frames && placed > 100 && moved > 100,
            "{followed} of {frames} frames followed what changed, {placed} cells placed, \
             {moved} readers of an extent moved in frames that followed: too few for the law \
             to hold over"
        );
    }

    // ---- KERNEL.49's third slice: the machine's own grid ----

    /// A grid of cells the tests keep as their reference (KERNEL.49's
    /// decision 2): each cell what the last row writing it wrote, a value,
    /// or a formula's text as the formula bar shows it, the row's text moved
    /// to the cell, with no fill anywhere.
    fn reference_after(cells: &mut BTreeMap<(u32, u32), String>, row: &str) {
        let forms = read_forms(row).unwrap();
        for form in &forms {
            let Some((head, items, _)) = load::row_parts(form) else {
                continue;
            };
            let [Form::Str(_), Form::Str(addr), third] = items else {
                continue;
            };
            let r = parse_a1_range(addr).unwrap();
            for row in r.top..=r.bottom {
                for col in r.left..=r.right {
                    let held = match (head, third) {
                        ("formula", Form::Str(text)) => {
                            let moved = shift_a1_references(
                                &text[1..],
                                i64::from(row) - i64::from(r.top),
                                i64::from(col) - i64::from(r.left),
                            );
                            format!("={}", strip_future_prefixes(&moved))
                        }
                        (_, value) => format!("{value:?}"),
                    };
                    cells.insert((row, col), held);
                }
            }
        }
    }

    /// The machine's rows' sheet as the reference spells it: each cell a
    /// formula's text or a value, through the grid it lends.
    fn spelled_cells(m: &Machine) -> BTreeMap<(u32, u32), String> {
        let book = m.workbook();
        let sheet = &book.sheets[0];
        let masters = shared_masters(sheet);
        sheet
            .cells
            .iter()
            .map(|(&(row, col), cell)| {
                let held = match formula_text(&cell.content, row, col, &masters) {
                    Some(text) => format!("={text}"),
                    None => {
                        let value = match &cell.content {
                            Content::Number(n) => Form::Sym(crate::sheet::number_text(*n).unwrap()),
                            Content::Text(t) => Form::Str(t.clone()),
                            other => panic!("({row}, {col}): {other:?}"),
                        };
                        format!("{value:?}")
                    }
                };
                ((row, col), held)
            })
            .collect()
    }

    #[test]
    fn fills_cut_by_writes_are_the_cells_the_same_writes_leave() {
        // KERNEL.49's decision 2: generated grids of fills, written over by
        // values and formulas of every size, each meeting a fill anywhere,
        // its first cell among them, so that a fill is cut into pieces, its
        // first cell moved where the move keeps its text whole and each cell
        // given its own text where it does not (a reference near the sheet's
        // first column, moved left of it). After every write the grid holds
        // the cells a grid of cells holds after the same writes, every
        // formula's text as the bar shows it, every value as written; and the
        // plan is a load's of the grid it leaves (KERNEL.34).
        // Every text reads cells to its left alone, so that no write makes
        // a cycle wherever it lands; the last reads the sheet's last row, so
        // that a cell of its fill below its first row reads past it, and a
        // move down does not keep the text whole.
        let texts = [
            "=A1+$B$1",
            "=B2*2+A$1",
            "=SUM($A$1:A3)",
            "=A2-B1*2",
            "=IF(A1>2,B1,$B$2)",
            "=A1048576+1",
        ];
        let (mut cuts, mut moved, mut own) = (0, 0, 0);
        for seed in 1..=60u64 {
            let mut d = Dice(seed ^ 0x4949);
            let mut rows = String::from("(cell \"S\" \"A1:B12\" 1)\n");
            let fills = [
                ("C1:F6", 0),
                ("G1:L4", 1),
                ("C7:L12", 2),
                ("G5:L6", 3),
                ("M1:M12", 5),
            ];
            for (range, k) in fills {
                rows.push_str(&format!(
                    "(formula \"S\" \"{range}\" \"{}\")\n",
                    texts[(k + d.below(texts.len())) % texts.len()]
                ));
            }
            let mut reference = BTreeMap::new();
            reference_after(&mut reference, &rows);
            let mut m = load(&rows);
            assert_eq!(spelled_cells(&m), reference, "seed {seed}: the load");
            // The same grid lent and loaded again, each fill's cells pieces of
            // one cell, takes every write as the grid itself.
            let mut again = loaded_again(&m);
            assert_eq!(
                spelled_cells(&again),
                reference,
                "seed {seed}: loaded again"
            );
            for k in 0..8 {
                // A third of the writes at a fill's first cell as it stands.
                let mut anchors: Vec<(u32, u32)> = m.book.pages[0]
                    .fills
                    .masters()
                    .values()
                    .map(|&(_, row, col)| (row, col))
                    .collect();
                anchors.sort_unstable();
                let (top, left) = match d.below(3) {
                    0 if !anchors.is_empty() => anchors[d.below(anchors.len())],
                    _ => (1 + d.below(12) as u32, 3 + d.below(11) as u32),
                };
                let bottom = (top + d.below(4) as u32).min(12);
                let right = (left + d.below(4) as u32).min(13);
                let range = A1Range {
                    top,
                    left,
                    bottom,
                    right,
                };
                let write = match d.below(3) {
                    0 => format!("(cell \"S\" \"{}\" {})", range.text(), d.below(9)),
                    _ => format!(
                        "(formula \"S\" \"{}\" \"{}\")",
                        range.text(),
                        texts[d.below(texts.len())]
                    ),
                };
                crate::sheet::work::begin();
                let answer = m.write(&write);
                let work = crate::sheet::work::end();
                let when = format!("seed {seed}, write {k}: {write}");
                assert!(answer.is_ok(), "{when}: {answer:?}");
                reference_after(&mut reference, &write);
                assert_eq!(spelled_cells(&m), reference, "{when}");
                assert_eq!(m.shapes.spelled(), plan_of_a_load(&m), "{when}: the plan");
                assert_eq!(again.write(&write), answer, "{when}: loaded again");
                assert_eq!(spelled_cells(&again), reference, "{when}: loaded again");
                assert_eq!(
                    again.shapes.spelled(),
                    plan_of_a_load(&again),
                    "{when}: loaded again, the plan"
                );
                cuts += 1;
                moved += work.whole;
                own += work.split;
            }
            assert!(m.step(0).done, "seed {seed}");
            assert!(again.step(0).done, "seed {seed}: loaded again");
        }
        assert!(
            moved >= 20 && own >= 5,
            "{cuts} writes, {moved} first cells moved, {own} fills given their own texts: \
             too few for the law to hold over"
        );
    }

    #[test]
    fn one_write_per_cell_is_refused_at_the_first_cell_written_twice() {
        // KERNEL.49's decision 8: rows of values and formulas over ranges and
        // cells, held as records by their rectangles and never made into
        // cells, refused where a row writes a cell an earlier row wrote, at
        // its first such cell in row-major order, naming the earlier row's
        // line, as a grid of cells refused it: a reference keeping each
        // cell's line. Where none is refused, every cell's line is its row's,
        // the cells are counted, and a walk of the sheet meets them in order.
        for seed in 1..=300u64 {
            let mut d = Dice(seed ^ 0x8888);
            let mut rows = String::new();
            let mut lines: BTreeMap<(u32, u32), u32> = BTreeMap::new();
            let mut want: Option<(u32, String, u32)> = None;
            for line in 1..=12u32 {
                let (top, left) = (1 + d.below(12) as u32, 1 + d.below(12) as u32);
                let (bottom, right) = match d.below(3) {
                    0 => (top, left),
                    _ => (top + d.below(6) as u32, left + d.below(6) as u32),
                };
                let r = A1Range {
                    top,
                    left,
                    bottom,
                    right,
                };
                rows.push_str(&match d.below(2) {
                    0 => format!("(cell \"S\" \"{}\" 1)\n", r.text()),
                    _ => format!("(formula \"S\" \"{}\" \"=1\")\n", r.text()),
                });
                if want.is_some() {
                    continue;
                }
                let first = (r.top..=r.bottom)
                    .flat_map(|row| (r.left..=r.right).map(move |col| (row, col)))
                    .find(|at| lines.contains_key(at));
                match first {
                    Some((row, col)) => {
                        let what = format!("S!{}", cell_ref(row, col));
                        want = Some((line, what, lines[&(row, col)]));
                    }
                    None => {
                        for row in r.top..=r.bottom {
                            for col in r.left..=r.right {
                                lines.insert((row, col), line);
                            }
                        }
                    }
                }
            }
            let got = read_rows(&read_forms(&rows).unwrap());
            let when = format!("seed {seed}\n{rows}");
            match want {
                Some((line, what, first)) => {
                    let e = got.expect_err(&when);
                    let w = load::twice(line, &what, first);
                    assert_eq!((e.line, e.refusal.text), (w.line, w.refusal.text), "{when}");
                }
                None => {
                    let draft = got.unwrap_or_else(|e| panic!("{when}: {}", e.refusal.text));
                    for (&(row, col), &line) in &lines {
                        assert_eq!(
                            draft.line(0, row, col),
                            Some(line),
                            "{when}: ({row}, {col})"
                        );
                    }
                    assert_eq!(draft.count(), lines.len(), "{when}: counted");
                    let walked: Vec<(u32, u32)> = draft.cells(0).map(|(at, _)| at).collect();
                    let cells: Vec<(u32, u32)> = lines.keys().copied().collect();
                    assert_eq!(walked, cells, "{when}: walked in order");
                }
            }
        }
    }

    #[test]
    fn a_draft_s_workbook_is_the_one_its_rows_wrote_cell_by_cell() {
        // KERNEL.49's decision 8: a draft keeps its rows as records and makes
        // a door's workbook from them whole; the workbook is the one a reader
        // making every cell as its row came made (`Sheet::set` for a value,
        // `Sheet::set_formula` for a formula, a shared formula over a range,
        // each cell's format and sentence set after), cell for cell.
        let texts = ["=A1+1", "=$B$2*C3", "=SUM(A1:A4)"];
        for seed in 1..=120u64 {
            let mut d = Dice(seed ^ 0x5d5d);
            let mut rows = String::new();
            let mut sheet = Sheet::new("S");
            let mut written: BTreeSet<(u32, u32)> = BTreeSet::new();
            for _ in 0..14 {
                let (top, left) = (1 + d.below(14) as u32, 1 + d.below(14) as u32);
                let (bottom, right) = match d.below(3) {
                    0 => (top, left),
                    _ => (top + d.below(4) as u32, left + d.below(4) as u32),
                };
                let r = A1Range {
                    top,
                    left,
                    bottom,
                    right,
                };
                let cells: Vec<(u32, u32)> = (r.top..=r.bottom)
                    .flat_map(|row| (r.left..=r.right).map(move |col| (row, col)))
                    .collect();
                if cells.iter().any(|at| written.contains(at)) {
                    continue;
                }
                written.extend(cells.iter().copied());
                match d.below(3) {
                    0 => {
                        let text = texts[d.below(texts.len())];
                        rows.push_str(&format!("(formula \"S\" \"{}\" \"{text}\")\n", r.text()));
                        sheet.set_formula(r, &text[1..], 0);
                    }
                    1 => {
                        rows.push_str(&format!("(cell \"S\" \"{}\" \"t{seed}\")\n", r.text()));
                        for &(row, col) in &cells {
                            let content = Content::Text(format!("t{seed}"));
                            sheet.set(row, col, Cell { content, style: 0 });
                        }
                    }
                    _ => {
                        let n = d.below(50);
                        rows.push_str(&format!("(cell \"S\" \"{}\" {n})\n", r.text()));
                        for &(row, col) in &cells {
                            let content = Content::Number(n as f64);
                            sheet.set(row, col, Cell { content, style: 0 });
                        }
                    }
                }
            }
            for &(row, col) in written.iter().step_by(3) {
                let index = 1 + d.below(3) as u32;
                rows.push_str(&format!(
                    "(style \"S\" \"{}\" {index})\n",
                    cell_ref(row, col)
                ));
                sheet.cells.get_mut(&(row, col)).unwrap().style = index;
            }
            for &(row, col) in written.iter().step_by(5) {
                rows.push_str(&format!(
                    "(sentence \"S\" \"{}\" {})\n",
                    cell_ref(row, col),
                    row
                ));
                sheet.sentences.insert((row, col), row);
            }
            let draft = read_rows(&read_forms(&rows).unwrap())
                .unwrap_or_else(|e| panic!("seed {seed}: {}", e.refusal.text));
            let book = draft.workbook();
            assert_eq!(book.sheets.len(), 1, "seed {seed}");
            assert_eq!(book.sheets[0], sheet, "seed {seed}\n{rows}");
            assert_eq!(draft.count(), written.len(), "seed {seed}: counted");
        }
    }

    #[test]
    fn a_cell_put_into_a_draft_or_taken_out_is_cut_from_its_record_and_counted() {
        // KERNEL.49's decision 8: an engine's own cells put into a draft and
        // taken out of it (Alonzo's Clock and File cells), each cut from the
        // record holding it, the rest of the record kept, and counted; a
        // fill's first cell taken gives its first cell left the fill, its
        // text moved there, as a machine's write does (KERNEL.34).
        let rows = "(cell \"S\" \"A1:C3\" 5)\n(formula \"S\" \"D1:F3\" \"=A1*2\")\n";
        let mut draft = read_rows(&read_forms(rows).unwrap()).unwrap();
        let texts = |d: &Draft| -> BTreeMap<(u32, u32), String> {
            let book = d.workbook();
            let sheet = &book.sheets[0];
            let masters = shared_masters(sheet);
            sheet
                .cells
                .iter()
                .filter_map(|(&(row, col), c)| {
                    Some(((row, col), formula_text(&c.content, row, col, &masters)?))
                })
                .collect()
        };
        let before = texts(&draft);
        assert_eq!(draft.count(), 18);
        assert!(draft.take(0, 2, 2), "B2 out of the values' record");
        assert_eq!(draft.count(), 17);
        assert_eq!(draft.content(0, 2, 2), None);
        assert_eq!(
            draft.content(0, 2, 1),
            Some(load::Drafted::Value(&Content::Number(5.0)))
        );
        assert_eq!(draft.line(0, 3, 3), Some(1), "C3 left in the record");
        draft.put(0, 2, 2, Content::Text("x".to_string()));
        assert_eq!((draft.count(), draft.line(0, 2, 2)), (18, Some(0)));
        // The fill's first cell taken: its first cell left, E1, its anchor.
        assert!(draft.take(0, 1, 4));
        let book = draft.workbook();
        assert!(
            matches!(
                &book.sheets[0].cells[&(1, 5)].content,
                Content::SharedMaster { text, .. } if text == "B1*2"
            ),
            "E1 the fill's first cell: {:?}",
            book.sheets[0].cells[&(1, 5)]
        );
        let mut want = before.clone();
        want.remove(&(1, 4));
        assert_eq!(texts(&draft), want, "every other cell's formula as it was");
        assert!(!draft.take(0, 9, 9), "a cell no row wrote");
        assert_eq!(draft.count(), 17);
        // A value put over a formula: cut from its fill, the rest as it was.
        draft.put(0, 2, 5, Content::Number(1.0));
        assert_eq!(draft.count(), 17);
        assert_eq!(
            draft.content(0, 2, 5),
            Some(load::Drafted::Value(&Content::Number(1.0)))
        );
        want.remove(&(2, 5));
        assert_eq!(texts(&draft), want);
        // Every formula once a record, by its first cell.
        let formulas: Vec<CellId> = draft.formulas().map(|(id, _, _)| id).collect();
        assert!(formulas.windows(2).all(|w| w[0] < w[1]), "{formulas:?}");
    }

    #[test]
    fn a_value_cell_s_error_views_by_the_text_it_was_written_with() {
        // KERNEL.49's third slice: the store is the cell, so an error a value
        // cell holds by a text its kind does not spell views as written, in
        // the record and in a twin's, the twin's own store kept where a
        // formula reads it, across a frame and a write, read as its kind by a
        // formula; an array's slot and a sparse sheet's loose cell alike.
        let rows = "(cell \"S\" \"A1:A3\" 1)\n\
                    (cell \"S\" \"B1\" (error \"#n/a\"))\n\
                    (cell \"S\" \"B2\" (error \"#SPILL!\"))\n\
                    (cell \"S\" \"B3\" (error \"#N/A\"))\n\
                    (cell \"S\" \"Z900\" (error \"#ref!\"))\n\
                    (formula \"S\" \"C1:C2\" \"=ISNA(B1)\")\n\
                    (formula \"S\" \"C3\" \"=ISNA(S.last!B1)\")\n";
        let mut m = load(rows);
        let row = |m: &Machine, sheet: &str, addr: &str| -> String {
            let text = grid(m, sheet, Some(addr));
            text.lines()
                .find(|l| l.starts_with("(cell "))
                .unwrap_or("")
                .to_string()
        };
        for frame in 0..=2 {
            if frame > 0 {
                assert!(m.step(0).done);
            }
            assert_eq!(row(&m, "S", "B1"), "(cell \"S\" \"B1\" (error \"#n/a\"))");
            assert_eq!(
                row(&m, "S", "B2"),
                "(cell \"S\" \"B2\" (error \"#SPILL!\"))"
            );
            assert_eq!(row(&m, "S", "B3"), "(cell \"S\" \"B3\" (error \"#N/A\"))");
            assert_eq!(
                row(&m, "S", "Z900"),
                "(cell \"S\" \"Z900\" (error \"#ref!\"))"
            );
            if frame > 0 {
                assert_eq!(
                    row(&m, "S.last", "B1"),
                    "(cell \"S.last\" \"B1\" (error \"#n/a\"))"
                );
                assert_eq!(
                    at(&m, "S", "C1"),
                    Some(Computed::Value(Value::Bool(true))),
                    "#n/a read as #N/A"
                );
                assert_eq!(
                    at(&m, "S", "C2"),
                    Some(Computed::Value(Value::Bool(false))),
                    "#SPILL! read as #VALUE!"
                );
                assert_eq!(
                    at(&m, "S", "C3"),
                    Some(Computed::Value(Value::Bool(true))),
                    "the twin's #n/a read as #N/A"
                );
            }
        }
        assert_eq!(m.write("(cell \"S\" \"B2\" (error \"#Div/0!\"))"), Ok(1));
        assert!(m.step(0).done);
        assert_eq!(
            row(&m, "S", "B2"),
            "(cell \"S\" \"B2\" (error \"#Div/0!\"))"
        );
        assert_eq!(
            content(&m, "S.last", "B2"),
            Some(Content::Error("#Div/0!".to_string()))
        );
        // The grid lent holds them as written, and loads again the same.
        let again = loaded_again(&m);
        assert_eq!(
            row(&again, "S", "B1"),
            "(cell \"S\" \"B1\" (error \"#n/a\"))"
        );
        assert_eq!(
            row(&again, "S", "Z900"),
            "(cell \"S\" \"Z900\" (error \"#ref!\"))"
        );
    }

    #[test]
    fn a_fill_s_cell_s_formula_is_read_by_its_fill_and_never_by_the_sheet_s_masters() {
        // KERNEL.49's third slice: SUBTOTAL asks each cell it walks whether
        // its formula calls SUBTOTAL, and FORMULATEXT asks a cell for its
        // text; a fill's cell's formula is read by its fill's index, so
        // neither makes the sheet's masters, however many fills it holds,
        // where each read made them, 5,050 fills read in a frame.
        let mut rows = String::new();
        for k in 0..50u32 {
            let col = 3 + k;
            rows.push_str(&format!(
                "(formula \"S\" \"{}:{}\" \"=ROW()*{}\")\n",
                cell_ref(1, col),
                cell_ref(100, col),
                k + 1
            ));
        }
        rows.push_str("(formula \"S\" \"A1\" \"=SUBTOTAL(9,C1:C100)\")\n");
        rows.push_str("(formula \"S\" \"A2\" \"=FORMULATEXT(D7)\")\n");
        let mut m = load(&rows);
        crate::sheet::work::begin();
        assert!(m.step(0).done);
        let work = crate::sheet::work::end();
        assert_eq!(work.masters, 0, "the sheet's masters made");
        assert_eq!(
            at(&m, "S", "A1"),
            Some(Computed::Value(Value::Number(5050.0)))
        );
        assert_eq!(
            at(&m, "S", "A2"),
            Some(Computed::Value(Value::Text("=ROW()*2".to_string())))
        );
    }

    #[test]
    fn a_value_written_over_a_value_cell_is_a_write_unless_it_is_the_same() {
        // KERNEL.49's third slice: the store is the cell, so a write compares
        // the value it writes with the value cell's content where it stands,
        // a text by its text, an error by the text it was written with, a
        // number by its value, a truth value by itself, on an array's store
        // and on a sparse sheet's: the same is no write (KERNEL.34), and
        // anything else is, viewed as written.
        let mut m = load(
            "(cell \"D\" \"A1\" \"a\")\n\
             (cell \"D\" \"A2\" (error \"#n/a\"))\n\
             (cell \"D\" \"A3\" 1)\n\
             (cell \"D\" \"A4\" true)\n\
             (cell \"S\" \"A1\" \"a\")\n\
             (cell \"S\" \"A2\" (error \"#n/a\"))\n\
             (cell \"S\" \"A3\" 1)\n\
             (cell \"S\" \"A4\" true)\n\
             (cell \"S\" \"Z900\" \"far\")\n",
        );
        assert!(m.step(0).done);
        let writes = [
            ("A1", "\"a\"", 0),
            ("A1", "\"b\"", 1),
            ("A1", "\"b\"", 0),
            ("A2", "(error \"#n/a\")", 0),
            ("A2", "(error \"#N/A\")", 1),
            ("A2", "(error \"#N/A\")", 0),
            ("A2", "(error \"#n/a\")", 1),
            ("A3", "1", 0),
            ("A3", "2", 1),
            ("A4", "true", 0),
            ("A4", "false", 1),
            ("Z900", "\"far\"", 0),
            ("Z900", "\"near\"", 1),
        ];
        for sheet in ["D", "S"] {
            for (addr, value, n) in writes {
                if sheet == "D" && addr == "Z900" {
                    continue;
                }
                let row = format!("(cell \"{sheet}\" \"{addr}\" {value})");
                assert_eq!(m.write(&row), Ok(n), "{row}");
                assert!(m.step(0).done, "{row}");
                let text = grid(&m, sheet, Some(addr));
                let held = text.lines().find(|l| l.starts_with("(cell ")).unwrap_or("");
                assert_eq!(held, row, "{row}: viewed");
            }
        }
    }

    #[test]
    fn a_fill_s_cell_reads_its_formula_moved_from_its_first_cell_left() {
        // KERNEL.49's third slice: a fill whose first cell a write took gives
        // its first cell left the fill, its text moved there (KERNEL.34); a
        // cell of it reads its formula from there, by its fill's index, as
        // the record prints it.
        let mut m = load(
            "(cell \"S\" \"A1:A9\" 1)\n\
             (formula \"S\" \"B1:B9\" \"=A1*2\")\n\
             (formula \"S\" \"C1\" \"=FORMULATEXT(B7)\")\n\
             (formula \"S\" \"C2\" \"=ISFORMULA(B1)\")\n",
        );
        assert!(m.step(0).done);
        let text = |m: &Machine| at(m, "S", "C1");
        let b7 = Some(Computed::Value(Value::Text("=A7*2".to_string())));
        assert_eq!(text(&m), b7);
        assert_eq!(m.write("(cell \"S\" \"B1\" 5)"), Ok(1));
        assert!(m.step(0).done);
        assert_eq!(text(&m), b7, "read from B2, the fill's first cell left");
        assert_eq!(at(&m, "S", "C2"), Some(Computed::Value(Value::Bool(false))));
        let record = grid(&m, "S", Some("B7"));
        assert!(
            record.contains("(formula \"S\" \"B7\" \"=A7*2\")"),
            "{record}"
        );
    }

    #[test]
    fn a_write_s_rows_count_the_cells_its_earlier_rows_added_against_the_cap() {
        // KERNEL.49's third slice: a write's values reach the store when it
        // settles, so the cells an earlier row of the same write added are
        // counted from the write's own log against the grid's cap: a row
        // writing them again adds none, and a row writing others adds its own.
        let mut m = load("(cell \"S\" \"A1:A1048000\" 1)\n");
        assert_eq!(
            m.write("(cell \"S\" \"B1:B400\" 2)\n(cell \"S\" \"B1:B400\" 3)"),
            Ok(800)
        );
        assert_eq!(m.cells, 1_048_400);
        let e = m
            .write("(cell \"S\" \"C1:C100\" 4)\n(cell \"S\" \"D1:D100\" 5)")
            .unwrap_err();
        assert_eq!((e.line, e.refusal.id.as_str()), (2, "grid-too-many-cells"));
        assert_eq!(m.cells, 1_048_400, "the refused write undone");
    }

    #[test]
    fn a_sentence_whose_cell_a_draft_gave_up_leaves_the_rest_theirs() {
        // KERNEL.49's third slice: a view meets each cell's format and
        // sentence beside the store's walk, passing by one whose cell is
        // gone: a cell taken out of a draft leaves its sheet's sentence, as
        // the file model's sheet kept it, and the cells after it inside the
        // window print theirs.
        let rows = "(cell \"S\" \"A1:C1\" 1)\n\
                    (sentence \"S\" \"A1\" 1)\n\
                    (sentence \"S\" \"B1\" 2)\n\
                    (sentence \"S\" \"C1\" 3)\n";
        let mut draft = read_rows(&read_forms(rows).unwrap()).unwrap();
        assert!(draft.take(0, 1, 2));
        let m = Machine::new(draft, library::language()).unwrap();
        let text = grid(&m, "S", None);
        assert!(text.contains("(sentence \"S\" \"A1\" 1)"), "{text}");
        assert!(!text.contains("\"B1\""), "{text}");
        assert!(text.contains("(sentence \"S\" \"C1\" 3)"), "{text}");
    }
}
