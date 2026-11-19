//! Recalculation's mechanism (KERNEL.7, slice 1, 2026-10-08): the formula
//! parser over the reference scanner, the dependency graph from `refers`,
//! evaluation in topological order, a cycle refused by name, Excel's error
//! values as a value kind, and a registry seam for functions, with the
//! language's own functions registered.
//!
//! It lands here, in the language, and not in `frazaro-core`, by the
//! three-crate order `Alonzo/CHARTER.md` section 4 and the roadmap's dated
//! note on `KERNEL.7` decide: an evaluator written into the bridges with
//! Excel's functions tangled in would make every game depend on the
//! bridges. So the mechanism and the day-one functions are the language's
//! ([`library::language`]); Excel's wider library registers through
//! [`Library`] from `frazaro-core`, by measurement and with a fixture per
//! function, and `frazaro calc` compares what it computes against the
//! values a host saved.
//!
//! An engine's grid is stepped by `crate::machine` (KERNEL.22,
//! 2026-10-09), over the same evaluator and the plan of shapes in
//! [`shape`]: every formula parsed once per shape and evaluated at each
//! cell's offset, a step with a budget in cells that yields and resumes
//! (`Alonzo/SPEC.md` section 4.4), the previous frame's `.last` twins being
//! ordinary sheets to it, each kept beside as an array over its extent that
//! a read indexes and a range walks ([`dense`], KERNEL.20), a run of a
//! shape's cells that read none of one another evaluated at once, node by
//! node (`vector`, KERNEL.20), and a changed cell followed to the cells that
//! read it, so that a frame computes those and keeps every other cell's
//! value (`reach`, KERNEL.20). [`Calc`] keeps its plan of texts and its
//! parse at evaluation, since it evaluates each formula once and an engine
//! every formula every frame. The plan of shapes prints itself on request,
//! its census (`census`, KERNEL.30): each shape's cells, levels and runs,
//! and the path each of its reads takes, so that a grid's cost is read
//! before a frame runs.
//!
//! The lookups compute since `KERNEL.8`'s first slice (2026-10-18), with the
//! error tests a lookup's miss is read by and `COUNTIF` and `SUMIF` moved
//! down from Excel's library ([`criteria`]), each search by a vector's
//! index where the environment keeps one ([`index`]), so that a lookup
//! asked again of one range costs a probe and not a scan. Excel's save of
//! the slice's fixture set its comparisons: two numbers equal to fifteen
//! significant digits are equal to an operator and a criterion, a lookup
//! matching a number to the bit, and two texts are equal without case
//! beyond ASCII too (`eval::compare_numbers`, `value::fold_char`).
//!
//! A legacy formula reads a range where one value is wanted at its own row
//! or column, Excel's implicit intersection, since the second slice
//! (2026-10-19, `eval::intersection`, `SINGLE` the form of `@`); and the
//! text functions compute since the third (2026-10-20, [`texts`]), over
//! Excel's order of texts as its save ranks them (`order`): two texts
//! ordered by their characters' places, then their accents, case aside, the
//! hyphen and the apostrophe last, `ß` and `ss` one, by `<` and `>`, a
//! lookup's sorted search and a criterion by size alike, and a comparison a
//! character beyond the save's table decides labelled.
//!
//! The date functions compute since the fourth slice (2026-10-20,
//! [`dates`]), over Excel's 1900 calendar, which `numfmt` draws by, a text
//! holding a digit no rule reads labelled in arithmetic as in `VALUE`, a
//! locale's number.
//!
//! A formula reads another cell's formula since KERNEL.53 (2026-10-21,
//! `formulas`): `ISFORMULA`, and `FORMULATEXT`, the text the view record
//! prints, labelled where Excel's formula bar spells it otherwise.
//!
//! A cell's address is a text since KERNEL.60 (2026-10-30, [`texts`]):
//! `ADDRESS`, an address in R1C1 labelled.
//!
//! A mask reads by one call since KERNEL.71 (2026-10-30, [`library`]): `BITAND`,
//! `BITOR` and `BITXOR`, over whole numbers below 2^48, exactly.
//!
//! A function is written in formulas since KERNEL.29's first slice (2026-11-15):
//! `LET` and `LAMBDA`, the names they bind resolved by the parse, a named
//! function a defined name whose text is a `LAMBDA`, and `MAP`, `REDUCE`,
//! `SCAN`, `BYROW`, `BYCOL`, `MAKEARRAY` and `ISOMITTED` (`helpers`). Each call
//! a formula makes is written into it at a shape's placement and by the door
//! before it evaluates (`lambda`), so that a run, a program and the plan meet
//! the formula written out; a recursion runs as calls (`calls`), each counted,
//! to a depth of 4,096 on a stack of the language's own.
//!
//! A dynamic formula spills since KERNEL.55's third slice (2026-11-18,
//! `spill`): the door plans each by the most rows and columns its parse can
//! answer, or the extent its file saved, a reader of any cell it may fill
//! ordered after its anchor; computes its anchor once and lays the answer into
//! the cells beside it, `#SPILL!` where they meet a value, a formula, a merged
//! range, a legacy array formula's block or the sheet's edge, as Excel's save
//! of the item's probe blocks one; and labels two spills that may meet, which
//! Excel settles by the order it computes them in, and one no plan bounds. A
//! spilled cell is a cell [`Calc`] computes, read by its address, by a range
//! and by `A1#` though the sheet holds none, and its value is what the writer
//! saves there, the anchor's `ref` its extent.
//!
//! What is not here, by name, each a slice of its own: spills in an engine's
//! machine (`KERNEL.55`'s fifth slice, until which `Calc::unspilled` reads a
//! grid as the machine does), a `LET`'s value computed once across `IF`'s
//! branches in a run and a program
//! and sheet-defined functions (`KERNEL.29`'s later slices); the rest of
//! speed, content-addressed evaluation among them (`KERNEL.20`). A formula
//! that reaches any of them is not computed,
//! and says so naming the function or the construct ([`Reason`]), which is
//! the honest static label of `HORIZON.md` section 12.6; a cell that reads
//! such a cell inherits the reason.

pub(crate) mod aggregates;
pub mod array;
pub(crate) mod arrays;
pub(crate) mod blocks;
pub(crate) mod calls;
pub(crate) mod census;
pub mod criteria;
pub(crate) mod database;
pub mod dates;
pub mod dense;
pub mod eval;
pub mod formula;
pub(crate) mod formulas;
pub mod graph;
pub(crate) mod helpers;
pub mod index;
pub(crate) mod information;
pub(crate) mod lambda;
pub mod library;
pub(crate) mod lookups;
pub(crate) mod nodes;
pub(crate) mod order;
pub(crate) mod program;
pub(crate) mod reach;
pub(crate) mod room;
pub mod shape;
pub(crate) mod spill;
pub mod stored;
pub mod texts;
pub mod value;
pub(crate) mod vector;

use std::cell::Cell as Slot;
use std::collections::{BTreeMap, HashMap};

use crate::grid::Rects;
use crate::messages::{raise, Refusal};
use crate::refers::quote_sheet;
use crate::sheet::{
    cell_ref, formula_text, shared_masters, A1Range, Content, RowSettings, Workbook, MAX_COLUMN,
    MAX_ROW,
};

pub use array::Array;
pub use dense::Dense;
pub use eval::{
    Arg, Arrayed, Classes, Computed, Ctx, Env, Form, Library, Param, Reason, Strict, Written,
};
pub use formula::{calls, parse, Expr};
pub use graph::{Area, Areas, CellId, Extent, Names, Plan, Walked};
pub use index::Indexes;
pub use shape::{Place, Shape, Shapes};
pub use value::{agrees, ErrorKind, Value};

/// How far a step got: the cells evaluated so far in this frame, the
/// frame's total, and whether the frame is complete.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Progress {
    pub evaluated: usize,
    pub of: usize,
    pub done: bool,
}

/// One recalculation of one model under one library: the plan made once,
/// the cells evaluated in order, by steps or all at once.
pub struct Calc<'m> {
    model: &'m Workbook,
    library: &'m Library,
    names: Names,
    extents: Vec<Extent>,
    texts: HashMap<CellId, String>,
    plan: Plan,
    values: HashMap<CellId, Computed>,
    next: usize,
    here: Slot<CellId>,
    /// The vectors its lookups and criteria search, indexed at a second
    /// probe (KERNEL.8): its one pass changes none of their cells.
    indexes: index::Indexes,
    /// The workbook's named functions (KERNEL.29), which every formula's
    /// inlining writes in as an engine's placement does, so that the door
    /// and an engine read a formula alike.
    functions: lambda::Functions,
    /// Each block a legacy array formula fills (KERNEL.55), by its anchor, the
    /// range's first cell, holding a formula.
    blocks: HashMap<CellId, crate::sheet::A1Range>,
    /// The cells of the blocks the recalculation computes beside the formula
    /// cells, each with its anchor: every cell of a block the sheet holds but
    /// the anchor and a cell holding a formula of its own.
    places: graph::Places,
    /// Each block's anchor's answer, as much of it as its block lays, kept
    /// from its computation for its places', which follow it in the order; a
    /// reason where it was not computed.
    laid: HashMap<CellId, Result<Array, Reason>>,
    /// Each dynamic formula's spill the plan bounds (KERNEL.55), by its
    /// anchor, with the box of the cells it may fill.
    spills: spill::Spills,
    /// The extent each anchor's answer spilled into, past its own cell.
    spilled: HashMap<CellId, A1Range>,
    /// Each cell a spill's computation wrote, of its extent and of the extent
    /// its file saved, with its anchor: the cells the recalculation computes
    /// beside the formula cells and the blocks'.
    poured: HashMap<CellId, CellId>,
    /// By sheet, the extents holding a cell the sheet does not, which a walk
    /// visits beside the sheet's own.
    unheld: Vec<Rects>,
    /// Whether it computes spills: every recalculation but [`Calc::unspilled`].
    spilling: bool,
}

impl<'m> Calc<'m> {
    /// The plan over the model: every formula cell's text read off the
    /// grid (a shared formula's child as its master's text moved), the
    /// graph built from `refers`, the order found and every cycle named.
    /// Nothing is evaluated yet.
    pub fn new(model: &'m Workbook, library: &'m Library) -> Calc<'m> {
        Calc::planned(model, library, true)
    }

    /// The plan over the model as an engine's machine computes it until
    /// KERNEL.55's fifth slice: no spill planned, a dynamic formula answering
    /// many values labelled a spill at its cell, as every one was before the
    /// item's third slice, and `A1#` of one labelled; what the laws holding a
    /// machine to `Calc` take until the machine computes spills.
    #[doc(hidden)]
    pub fn unspilled(model: &'m Workbook, library: &'m Library) -> Calc<'m> {
        let mut calc = Calc::planned(model, library, false);
        calc.step(0);
        calc
    }

    fn planned(model: &'m Workbook, library: &'m Library, spilling: bool) -> Calc<'m> {
        let names = Names::of(model);
        let extents: Vec<Extent> = model.sheets.iter().map(|s| s.extent()).collect();
        let mut texts = HashMap::new();
        for (si, sheet) in model.sheets.iter().enumerate() {
            let masters = shared_masters(sheet);
            for (&(row, col), cell) in &sheet.cells {
                if let Some(text) = formula_text(&cell.content, row, col, &masters) {
                    texts.insert((si, row, col), text);
                }
            }
        }
        let (blocks, places) = blocks_of(model, &texts);
        let listed: Vec<(usize, crate::sheet::A1Range)> =
            blocks.iter().map(|(id, range)| (id.0, *range)).collect();
        let functions = lambda::Functions::of_model(model);
        let spills = match spilling {
            true => spill::Spills::of(
                &spill::dynamics_of(model, &texts),
                &|name| model.find_sheet(name),
                &names,
                &functions,
                library,
            ),
            false => spill::Spills::default(),
        };
        let plan = graph::plan_placed(model, &names, &extents, &texts, (&places, &listed), &spills);
        let mut values = HashMap::with_capacity(texts.len() + places.len());
        for cycle in &plan.cycles {
            for id in cycle {
                values.insert(*id, Computed::NotComputed(Reason::Cycle));
            }
        }
        Calc {
            model,
            library,
            names,
            extents,
            texts,
            plan,
            values,
            next: 0,
            here: Slot::new((0, 1, 1)),
            indexes: index::Indexes::new(),
            functions,
            blocks,
            places,
            laid: HashMap::new(),
            spills,
            spilled: HashMap::new(),
            poured: HashMap::new(),
            unheld: Vec::new(),
            spilling,
        }
    }

    /// The plan made and every cell evaluated.
    pub fn run(model: &'m Workbook, library: &'m Library) -> Calc<'m> {
        let mut calc = Calc::new(model, library);
        calc.step(0);
        calc
    }

    /// The model.
    pub fn model(&self) -> &'m Workbook {
        self.model
    }

    /// The library.
    pub fn library(&self) -> &'m Library {
        self.library
    }

    /// The plan: the order and the cycles.
    pub fn plan(&self) -> &Plan {
        &self.plan
    }

    /// Every cycle, each its cells sorted, the cycles by their first cell.
    pub fn cycles(&self) -> &[Vec<CellId>] {
        &self.plan.cycles
    }

    /// How many formula cells the model holds.
    pub fn formulas(&self) -> usize {
        self.texts.len()
    }

    /// A formula cell's text without its `=`, a block's cell's its anchor's
    /// (KERNEL.55), as Excel holds a legacy array formula at every cell of its
    /// block, and a spilled cell's its anchor's, as its formula bar shows it;
    /// `None` for a value cell.
    pub fn formula(&self, id: CellId) -> Option<&str> {
        let id = self
            .places
            .get(&id)
            .or_else(|| self.poured.get(&id))
            .copied()
            .unwrap_or(id);
        self.texts.get(&id).map(String::as_str)
    }

    /// Every formula cell, in sheet then row-major order.
    pub fn formula_cells(&self) -> Vec<CellId> {
        let mut ids: Vec<CellId> = self.texts.keys().copied().collect();
        ids.sort_unstable();
        ids
    }

    /// Every cell the recalculation computes, in sheet then row-major order:
    /// each formula cell, each cell of a block a legacy array formula fills,
    /// which its anchor's array lays, and, once its anchor is computed, each
    /// cell a dynamic formula spills into and each its file saved it spilling
    /// into (KERNEL.55).
    pub fn cells(&self) -> Vec<CellId> {
        let mut ids: Vec<CellId> = self
            .texts
            .keys()
            .chain(self.places.keys())
            .chain(self.poured.keys())
            .copied()
            .collect();
        ids.sort_unstable();
        ids
    }

    /// What a formula cell holds, once evaluated, a block's cell, its place of
    /// its anchor's array, and a spilled cell, its place of its anchor's, or
    /// nothing where its file saved it spilled and it no longer does
    /// (KERNEL.55); `None` for a value cell or a cell the steps have not
    /// reached.
    pub fn computed(&self, id: CellId) -> Option<&Computed> {
        self.values.get(&id)
    }

    /// Evaluate the next cells of the order, at most `budget` of them, 0
    /// meaning every one: `SD-34`'s step, a budget in cells and never in
    /// time, so that the frame's values are the same whatever the host's
    /// chunking. The cells in a cycle count as done from the start.
    pub fn step(&mut self, budget: usize) -> Progress {
        let order_len = self.plan.order.len();
        let end = if budget == 0 {
            order_len
        } else {
            self.next.saturating_add(budget).min(order_len)
        };
        for i in self.next..end {
            let id = self.plan.order[i];
            let computed = match self.places.get(&id) {
                Some(anchor) => self.place(id, *anchor),
                None => match self.blocks.get(&id) {
                    Some(block) => {
                        let (computed, laid) = self.evaluated(id, Some(*block));
                        if let Some(laid) = laid {
                            self.laid.insert(id, laid);
                        }
                        computed
                    }
                    None => match self.spills.get(id) {
                        Some(spill) => {
                            let (computed, laid) = self.spill(id, &spill);
                            self.pour(id, &spill, &computed, laid);
                            computed
                        }
                        None => self.evaluate(id),
                    },
                },
            };
            self.values.insert(id, computed);
        }
        self.next = end;
        self.progress()
    }

    /// A block's cell's value, its place of its anchor's array as Excel lays
    /// one ([`laid_at`]), settled as a formula's answer is; the anchor's
    /// reason where the anchor was not computed (KERNEL.55).
    fn place(&self, (_, row, col): CellId, anchor: CellId) -> Computed {
        match self.laid.get(&anchor) {
            Some(Ok(array)) => {
                Computed::Value(laid_at(array, row - anchor.1, col - anchor.2).settled())
            }
            Some(Err(reason)) => Computed::NotComputed(reason.clone()),
            // An anchor in a cycle is never laid.
            None => match self.values.get(&anchor) {
                Some(Computed::NotComputed(reason)) => Computed::NotComputed(reason.clone()),
                _ => Computed::NotComputed(Reason::Cycle),
            },
        }
    }

    /// Where the evaluation stands.
    pub fn progress(&self) -> Progress {
        let in_cycles = self.plan.cycles.iter().map(Vec::len).sum::<usize>();
        Progress {
            evaluated: self.next + in_cycles,
            of: self.plan.formulas(),
            done: self.next == self.plan.order.len(),
        }
    }

    /// A spill's anchor computed (KERNEL.55): its answer's first value at its
    /// cell, and the extent it spills into with the array it lays there; an
    /// answer of one value, no extent; `#SPILL!` where the extent passes the
    /// sheet's edge or meets what blocks it ([`Calc::blocked`]); labelled where
    /// its plan is, or where its answer passes the bound it was planned by.
    fn spill(&self, id: CellId, spill: &spill::Spill) -> (Computed, Option<(A1Range, Array)>) {
        let not = |reason: Reason| (Computed::NotComputed(reason), None);
        if let Some(label) = spill.label {
            return not(Reason::Construct(label.to_string()));
        }
        let Some(text) = self.texts.get(&id) else {
            return not(Reason::Construct(String::new()));
        };
        self.here.set(id);
        let parsed = match lambda::parsed(text, &self.functions, self.library) {
            Ok(e) => e,
            Err(construct) => return not(Reason::Construct(construct)),
        };
        let ctx = Ctx::new(self, self.library);
        if parsed.has_lambda() {
            ctx.nests(lambda::kept_depth(&parsed));
        }
        let arg = match ctx.eval(&parsed) {
            Ok(arg) => arg,
            Err(reason) => return not(reason),
        };
        let past = |rows: u32, columns: u32| rows > spill.dims.0 || columns > spill.dims.1;
        let unplanned = || not(Reason::Construct(spill::NO_PLAN.to_string()));
        let array = match arg {
            Arg::Array(a) => a,
            Arg::Area { ref areas, .. } if matches!(&areas[..], [a] if !a.is_cell()) => {
                let a = areas[0];
                if past(a.bottom - a.top + 1, a.right - a.left + 1) {
                    return unplanned();
                }
                match ctx.area_array(&a) {
                    Ok(array) => array,
                    Err(reason) => return not(reason),
                }
            }
            arg => {
                let computed = match ctx.result(&arg) {
                    Ok(v) => Computed::Value(v.settled()),
                    Err(reason) => Computed::NotComputed(reason),
                };
                return (computed, None);
            }
        };
        let first = array.get(0, 0).clone().settled();
        if array.is_single() {
            return (Computed::Value(first), None);
        }
        let (rows, columns) = (array.rows(), array.columns());
        if past(rows, columns) {
            return unplanned();
        }
        let bottom = u64::from(id.1) + u64::from(rows) - 1;
        let right = u64::from(id.2) + u64::from(columns) - 1;
        let spilt = Computed::Value(Value::Error(ErrorKind::Spill));
        if bottom > u64::from(MAX_ROW) || right > u64::from(MAX_COLUMN) {
            return (spilt, None);
        }
        let extent = A1Range {
            top: id.1,
            left: id.2,
            bottom: bottom as u32,
            right: right as u32,
        };
        if self.blocked(id, spill, extent) {
            return (spilt, None);
        }
        (Computed::Value(first), Some((extent, array)))
    }

    /// Whether what a spill's extent meets blocks it, as Excel's save of the
    /// item's probe shows (KERNEL.55): a cell holding a formula of its own, or
    /// a value but where the extent its file saved holds its last; a cell of a
    /// legacy array formula's block; or a merged range.
    fn blocked(&self, anchor: CellId, spill: &spill::Spill, extent: A1Range) -> bool {
        let Some(sheet) = self.model.sheets.get(anchor.0) else {
            return false;
        };
        if sheet.merges().meeting(extent).next().is_some() {
            return true;
        }
        let met = |r: &A1Range| crate::grid::overlap(r, &extent).is_some();
        if self.blocks.iter().any(|(b, r)| b.0 == anchor.0 && met(r)) {
            return true;
        }
        let cells = sheet
            .cells
            .range((extent.top, extent.left)..=(extent.bottom, extent.right));
        for (&(row, col), cell) in cells {
            if col < extent.left || col > extent.right || (row, col) == (anchor.1, anchor.2) {
                continue;
            }
            let last = spill.file.is_some_and(|f| spill::inside(&f, row, col));
            if graph::is_formula(&cell.content) || !last {
                return true;
            }
        }
        false
    }

    /// A spill's computation written into its cells (KERNEL.55): each of its
    /// extent its place of its anchor's array, as a lift reads it, settled as a
    /// formula's answer is; each the extent its file saved holds past it, but a
    /// formula's, empty, or the anchor's reason where it was not computed.
    fn pour(
        &mut self,
        anchor: CellId,
        spill: &spill::Spill,
        computed: &Computed,
        laid: Option<(A1Range, Array)>,
    ) {
        let s = anchor.0;
        let mut extent = None;
        if let Some((e, array)) = laid {
            let mut unheld = false;
            for row in e.top..=e.bottom {
                for col in e.left..=e.right {
                    if (row, col) == (anchor.1, anchor.2) {
                        continue;
                    }
                    let v = laid_at(&array, row - e.top, col - e.left).settled();
                    unheld |= !self.model.sheets[s].cells.contains_key(&(row, col));
                    self.values.insert((s, row, col), Computed::Value(v));
                    self.poured.insert((s, row, col), anchor);
                }
            }
            // The sheet's extent holds it, as a range of whole columns reads.
            self.extents[s] = grown(self.extents[s], e);
            if unheld {
                if self.unheld.len() <= s {
                    self.unheld.resize_with(s + 1, Rects::new);
                }
                self.unheld[s].insert(e, 0);
            }
            self.spilled.insert(anchor, e);
            extent = Some(e);
        }
        let Some(file) = spill.file else {
            return;
        };
        let rest = match computed {
            Computed::NotComputed(reason) => Computed::NotComputed(reason.clone()),
            Computed::Value(_) => Computed::Value(Value::Empty),
        };
        for row in file.top..=file.bottom {
            for col in file.left..=file.right {
                let id = (s, row, col);
                let own = self.texts.contains_key(&id) || self.places.contains_key(&id);
                if id == anchor || own || extent.is_some_and(|e| spill::inside(&e, row, col)) {
                    continue;
                }
                self.values.insert(id, rest.clone());
                self.poured.insert(id, anchor);
            }
        }
    }

    /// Start the order again, every value forgotten but the cycles'.
    pub fn restart(&mut self) {
        self.values
            .retain(|_, v| matches!(v, Computed::NotComputed(Reason::Cycle)));
        self.next = 0;
        self.indexes.clear();
        self.laid.clear();
        self.spilled.clear();
        self.poured.clear();
        self.unheld.clear();
        self.extents = self.model.sheets.iter().map(|s| s.extent()).collect();
    }

    fn evaluate(&self, id: CellId) -> Computed {
        self.evaluated(id, None).0
    }

    /// A formula cell computed; and where it anchors `block`, its answer as
    /// much as the block lays of it (KERNEL.55, [`laid_of`]), or the reason it
    /// was not computed.
    fn evaluated(
        &self,
        id: CellId,
        block: Option<crate::sheet::A1Range>,
    ) -> (Computed, Option<Result<Array, Reason>>) {
        let not = |reason: Reason| {
            let laid = block.map(|_| Err(reason.clone()));
            (Computed::NotComputed(reason), laid)
        };
        let Some(text) = self.texts.get(&id) else {
            return not(Reason::Construct(String::new()));
        };
        self.here.set(id);
        let parsed = match lambda::parsed(text, &self.functions, self.library) {
            Ok(e) => e,
            Err(construct) => return not(Reason::Construct(construct)),
        };
        let ctx = Ctx::new(self, self.library);
        // How deep the formula nests, so that a call it keeps is entered
        // where its body fits the stack (KERNEL.29).
        if parsed.has_lambda() {
            ctx.nests(lambda::kept_depth(&parsed));
        }
        let arg = match ctx.eval(&parsed) {
            Ok(arg) => arg,
            Err(reason) => return not(reason),
        };
        let laid = block.map(|b| laid_of(&ctx, &arg, b));
        let computed = match ctx.result(&arg) {
            Ok(v) => Computed::Value(v.settled()),
            // A dynamic formula planned at its cell alone, which spills
            // (KERNEL.55).
            Err(Reason::Construct(c)) if self.spilling && c == eval::SPILL => {
                Computed::NotComputed(Reason::Construct(spill::NO_PLAN.to_string()))
            }
            Err(reason) => Computed::NotComputed(reason),
        };
        (computed, laid)
    }

    /// The functions the formulas call that the library does not hold,
    /// with the first cell calling each, in sheet then row-major order:
    /// what the engine's `load` refuses by name before the first step.
    pub fn missing_functions(&self) -> Vec<(CellId, String)> {
        let mut seen: HashMap<String, CellId> = HashMap::new();
        for id in self.formula_cells() {
            // The calls the parse makes, written in, as an engine's placement
            // reads them (KERNEL.29); a text not read, its tokens'.
            let text = &self.texts[&id];
            let calls = match lambda::parsed(text, &self.functions, self.library) {
                Ok(e) => lambda::calls_of(&e),
                Err(_) => formula::calls(text),
            };
            for name in calls {
                if !self.library.has(&name) {
                    seen.entry(name).or_insert(id);
                }
            }
        }
        let mut out: Vec<(CellId, String)> = seen.into_iter().map(|(n, id)| (id, n)).collect();
        out.sort();
        out
    }

    /// `Sheet!A1` for a cell, the sheet quoted as a reference quotes it.
    pub fn spell_cell(&self, id: CellId) -> String {
        let sheet = self
            .model
            .sheets
            .get(id.0)
            .map(|s| s.name.as_str())
            .unwrap_or("");
        format!("{}!{}", quote_sheet(sheet), cell_ref(id.1, id.2))
    }

    /// The first thing that stops the whole grid from being stepped, as a
    /// refusal from the catalogue: a cycle, naming the cells that close it
    /// (`calc-cycle`), else the first function the library does not hold,
    /// naming the cell that calls it (`calc-function-not-computed`), else
    /// the first formula whose text this version does not read, naming the
    /// construct (`calc-construct-not-read`, KERNEL.22: the "unreadable
    /// reference" KERNEL.7's fifth decision named for an engine's load);
    /// `None` when every formula can be computed. An engine's `load` asks
    /// this once; `frazaro calc` never does, since a workbook is read as it
    /// is and each cell says what it can.
    pub fn first_refusal(&self) -> Option<Refusal> {
        if let Some(cycle) = self.plan.cycles.first() {
            let cells: Vec<String> = cycle.iter().map(|id| self.spell_cell(*id)).collect();
            return Some(raise("calc-cycle", &[("cells", &cells.join(", "))]));
        }
        if let Some((id, name)) = self.missing_functions().into_iter().next() {
            return Some(raise(
                "calc-function-not-computed",
                &[("cell", &self.spell_cell(id)), ("function", &name)],
            ));
        }
        // A construct the parse does not read, or a defined name the
        // formula reads whose text it does not (KERNEL.8, slice b), the first
        // cell holding either.
        let named = self.names.any_unread();
        let (id, construct) =
            self.formula_cells().into_iter().find_map(|id| {
                match lambda::parsed(&self.texts[&id], &self.functions, self.library) {
                    Err(construct) => Some((id, construct)),
                    Ok(parsed) if named => {
                        eval::unread_name(self.model, &self.names, &parsed, id.0).map(|c| (id, c))
                    }
                    Ok(_) => None,
                }
            })?;
        Some(formula::unread_refusal(&self.spell_cell(id), &construct))
    }
}

/// What a projection asks of a recalculation (KERNEL.22): a formula cell's
/// computed value, or `None` for a value cell or a formula with no value
/// yet. [`Calc`] answers it, and so does an engine's machine, whose values
/// outlive any one recalculation; the view record prints its `value` rows
/// from either. The value is the answer's own (KERNEL.49): a machine keeps a
/// formula's value as nine bytes in its sheet's store and holds no
/// [`Computed`] to lend.
pub trait Values: std::fmt::Debug {
    fn computed(&self, id: CellId) -> Option<Computed>;

    /// A sheet's extent where the values keep it (KERNEL.37), so that a view
    /// reads its `extent` row without walking every cell of the sheet: the
    /// extent itself, `Some(None)` for a sheet with no cell, or `None` to
    /// have the view walk the sheet. [`Calc`] keeps every sheet's from its
    /// making, and an engine's machine from its load and every write and
    /// frame's end, for the evaluator ([`Env::extent`]).
    fn kept_extent(&self, _sheet: usize) -> Option<Extent> {
        None
    }

    /// The extent an array formula's anchor fills past its own cell, where the
    /// values computed it (KERNEL.55): a dynamic formula's spill, what a writer
    /// saves as its `ref`, each of its other cells holding its value, as Excel
    /// saves a spill, and a legacy array formula's block; what the record
    /// prints as the anchor's `spill` row. `None`, unless the values say
    /// otherwise.
    fn spill(&self, _anchor: CellId) -> Option<A1Range> {
        None
    }

    /// A sheet's spills and blocks, each its anchor and the box of the cells
    /// its rows cover past its cell (KERNEL.55): its extent, and the cells its
    /// file saved it spilling into that its answer no longer reaches, each a
    /// `value` row with nothing before it in the record, or none for a cell
    /// left empty. None, unless the values say otherwise.
    fn spills(&self, _sheet: usize) -> Vec<(CellId, A1Range)> {
        Vec::new()
    }
}

impl Values for Calc<'_> {
    fn computed(&self, id: CellId) -> Option<Computed> {
        Calc::computed(self, id).cloned()
    }

    fn kept_extent(&self, sheet: usize) -> Option<Extent> {
        self.extents.get(sheet).copied()
    }

    fn spill(&self, anchor: CellId) -> Option<A1Range> {
        self.spilled
            .get(&anchor)
            .or_else(|| self.blocks.get(&anchor))
            .copied()
    }

    /// Each spill on the sheet, its extent with each cell it poured into, and
    /// each block, its range.
    fn spills(&self, sheet: usize) -> Vec<(CellId, A1Range)> {
        let mut boxes: BTreeMap<CellId, A1Range> = self
            .spilled
            .iter()
            .chain(self.blocks.iter())
            .filter(|(anchor, _)| anchor.0 == sheet)
            .map(|(anchor, e)| (*anchor, *e))
            .collect();
        for (&(s, row, col), &anchor) in &self.poured {
            if s != sheet {
                continue;
            }
            let cell = A1Range {
                top: row,
                left: col,
                bottom: row,
                right: col,
            };
            let start = A1Range {
                top: anchor.1,
                left: anchor.2,
                bottom: anchor.1,
                right: anchor.2,
            };
            let b = boxes.entry(anchor).or_insert(start);
            *b = A1Range {
                top: b.top.min(cell.top),
                left: b.left.min(cell.left),
                bottom: b.bottom.max(cell.bottom),
                right: b.right.max(cell.right),
            };
        }
        boxes.into_iter().collect()
    }
}

impl std::fmt::Debug for Calc<'_> {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("Calc")
            .field("formulas", &self.texts.len())
            .field("cycles", &self.plan.cycles.len())
            .field("progress", &self.progress())
            .finish()
    }
}

impl Env for Calc<'_> {
    fn find_sheet(&self, name: &str) -> Option<usize> {
        self.model.find_sheet(name)
    }

    fn sheet_name(&self, sheet: usize) -> Option<&str> {
        self.model.sheets.get(sheet).map(|s| s.name.as_str())
    }

    fn sheets(&self) -> usize {
        self.model.sheets.len()
    }

    fn rows(&self, sheet: usize) -> Option<&BTreeMap<u32, RowSettings>> {
        self.model.sheets.get(sheet).map(|s| &s.rows)
    }

    /// The cells a sheet holds inside an area, and with them each cell a spill
    /// filled that the sheet does not hold, a model's built in memory, so that
    /// a range over a spill reads it whole (KERNEL.55).
    fn walk(&self, area: &Area) -> Walked<'_> {
        let Some(sheet) = self.model.sheets.get(area.sheet) else {
            return Walked::nothing();
        };
        let extent = self.extent(area.sheet);
        let met = A1Range {
            top: area.top,
            left: area.left,
            bottom: area.bottom,
            right: area.right,
        };
        let mut unheld = self
            .unheld
            .get(area.sheet)
            .into_iter()
            .flat_map(|r| r.meeting(met))
            .filter_map(|(e, _)| crate::grid::overlap(&e, &met))
            .peekable();
        if unheld.peek().is_none() {
            return Walked::cells(sheet, extent, area);
        }
        let mut at: Vec<(u32, u32)> = Walked::cells(sheet, extent, area).collect();
        for e in unheld {
            for row in e.top..=e.bottom {
                at.extend((e.left..=e.right).map(|col| (row, col)));
            }
        }
        at.sort_unstable();
        at.dedup();
        Walked::listed(at)
    }

    /// A cell's formula, its text the one the plan read off the grid, moved
    /// and its prefixes dropped, so that no read walks a sheet for a fill's
    /// master (KERNEL.49).
    fn written(&self, id: CellId) -> Option<Written<'_>> {
        let kind = match crate::grid::Model::formula(self.model, id)? {
            crate::grid::Formula::Own(_, kind) => kind,
            crate::grid::Formula::Fill(_) => crate::grid::Kind::Plain,
        };
        Some(Written {
            text: self.texts.get(&id).map(String::as_str),
            moved: (0, 0),
            kind,
        })
    }

    fn here(&self) -> CellId {
        self.here.get()
    }

    fn read(&self, id: CellId) -> Result<Value, Reason> {
        // A formula cell, a block's cell, its place of its anchor's array, or a
        // spilled cell, its place of its anchor's (KERNEL.55), read as computed
        // and never as the value the file gave it.
        let spills = !self.spills.is_empty();
        if self.texts.contains_key(&id)
            || self.places.contains_key(&id)
            || (spills && self.poured.contains_key(&id))
        {
            return match self.values.get(&id) {
                Some(Computed::Value(v)) => Ok(v.clone()),
                Some(Computed::NotComputed(r)) => Err(r.clone()),
                // Every dependency precedes its reader in the order; a
                // formula with no value yet can only be read from a cycle.
                None => Err(Reason::Cycle),
            };
        }
        // A cell a spill may fill that its anchors' answers left empty, which
        // the sheet holds nothing at, or the first's reason not computed.
        if spills {
            for anchor in self.spills.owners(self.model, id) {
                match self.values.get(&anchor) {
                    Some(Computed::Value(_)) => {}
                    Some(Computed::NotComputed(r)) => return Err(r.clone()),
                    None => return Err(Reason::Cycle),
                }
            }
        }
        let Some(cell) = self
            .model
            .sheets
            .get(id.0)
            .and_then(|s| s.cells.get(&(id.1, id.2)))
        else {
            return Ok(Value::Empty);
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

    fn function(&self, name: &str, home: usize) -> Option<std::rc::Rc<formula::Lambda>> {
        self.functions.resolve(name, home)
    }

    fn indexes(&self) -> Option<&index::Indexes> {
        Some(&self.indexes)
    }

    /// Whether the cell being computed holds a dynamic-array formula, as its
    /// content says (KERNEL.8, slice b).
    fn dynamic(&self) -> bool {
        let (s, row, col) = self.here.get();
        is_dynamic(self.model, (s, row, col))
    }

    /// Whether the cell being computed holds a legacy array formula, as its
    /// content says (KERNEL.8, slice f).
    fn array_formula(&self) -> bool {
        is_array_formula(self.model, self.here.get())
    }

    /// The cells a dynamic formula's anchor spilled into (KERNEL.55): its
    /// extent, or its own cell for an answer of one value; the reason it was
    /// not computed.
    fn spilled(&self, anchor: CellId) -> Option<Result<Area, Reason>> {
        if !self.spilling {
            return None;
        }
        let e = match self.values.get(&anchor) {
            Some(Computed::Value(_)) => self.spilled.get(&anchor).copied().unwrap_or(A1Range {
                top: anchor.1,
                left: anchor.2,
                bottom: anchor.1,
                right: anchor.2,
            }),
            Some(Computed::NotComputed(r)) => return Some(Err(r.clone())),
            None => return Some(Err(Reason::Cycle)),
        };
        Some(Ok(Area {
            sheet: anchor.0,
            top: e.top,
            left: e.left,
            bottom: e.bottom,
            right: e.right,
        }))
    }
}

/// A sheet's extent grown to hold a spill's (KERNEL.55).
fn grown(extent: Extent, e: A1Range) -> Extent {
    Some(match extent {
        None => ((e.top, e.left), (e.bottom, e.right)),
        Some(((top, left), (bottom, right))) => (
            (top.min(e.top), left.min(e.left)),
            (bottom.max(e.bottom), right.max(e.right)),
        ),
    })
}

/// Whether a cell of a model holds a legacy array formula (KERNEL.8, slice
/// f): [`Env::array_formula`]'s answer for `Calc`.
pub(crate) fn is_array_formula(model: &Workbook, (s, row, col): CellId) -> bool {
    matches!(
        model
            .sheets
            .get(s)
            .and_then(|sheet| sheet.cells.get(&(row, col)))
            .map(|c| &c.content),
        Some(Content::ArrayFormula(_) | Content::ArrayBlock { .. })
    )
}

/// Each block a legacy array formula fills in a model (KERNEL.55), by its
/// anchor, a formula cell at its range's first; and its places: every cell of
/// the block the sheet holds but the anchor and a cell holding a formula of
/// its own, a cell two blocks share its first's. The cells the sheet holds
/// bound the places, as a file of Excel's holds a value at each.
fn blocks_of(
    model: &Workbook,
    texts: &HashMap<CellId, String>,
) -> (HashMap<CellId, crate::sheet::A1Range>, graph::Places) {
    let mut blocks = HashMap::new();
    let mut places = graph::Places::new();
    for (s, sheet) in model.sheets.iter().enumerate() {
        for (&(row, col), cell) in &sheet.cells {
            let Content::ArrayBlock { range, .. } = &cell.content else {
                continue;
            };
            if (range.top, range.left) != (row, col) || !texts.contains_key(&(s, row, col)) {
                continue;
            }
            blocks.insert((s, row, col), *range);
        }
    }
    let mut anchors: Vec<(&CellId, &crate::sheet::A1Range)> = blocks.iter().collect();
    anchors.sort_unstable_by_key(|(id, _)| **id);
    for (&(s, row, col), range) in anchors {
        let held = model.sheets[s]
            .cells
            .range((range.top, range.left)..=(range.bottom, range.right));
        for (&(r, c), _) in held {
            let id = (s, r, c);
            if c < range.left || c > range.right || (r, c) == (row, col) || texts.contains_key(&id)
            {
                continue;
            }
            places.entry(id).or_insert((s, row, col));
        }
    }
    (blocks, places)
}

/// A block's anchor's answer as much as its block lays of it (KERNEL.55): an
/// array computed, as it is; a reference's values, read no further than the
/// block reaches, which lays them as it would the whole; one value, an array
/// of one; several areas `#REF!`, as its cell's answer is.
fn laid_of(ctx: &Ctx<'_>, arg: &Arg<'_>, block: crate::sheet::A1Range) -> Result<Array, Reason> {
    match arg {
        Arg::Array(a) => Ok(a.clone()),
        Arg::Area { areas, .. } => match &areas[..] {
            [a] => {
                let rows = (a.bottom - a.top).min(block.bottom - block.top);
                let columns = (a.right - a.left).min(block.right - block.left);
                ctx.area_array(&graph::Area {
                    sheet: a.sheet,
                    top: a.top,
                    left: a.left,
                    bottom: a.top + rows,
                    right: a.left + columns,
                })
            }
            _ => Ok(Array::single(Value::Error(ErrorKind::Ref))),
        },
        Arg::Scalar(v) => Ok(Array::single(v.clone())),
        // A function's value, as a cell's answer is (KERNEL.29).
        Arg::Function(_) => Ok(Array::single(Value::Error(ErrorKind::Calc))),
    }
}

/// The value an array lays at a place of a block `row` rows and `column`
/// columns from its anchor, as Excel lays a legacy array formula's (KERNEL.55,
/// the save of its probe): its place as a lift reads it ([`Array::at`]), an
/// array of one row repeated down the block, one of one column across it, one
/// value everywhere, and `#N/A` past the array.
fn laid_at(a: &Array, row: u32, column: u32) -> Value {
    a.at(row, column)
        .cloned()
        .unwrap_or(Value::Error(ErrorKind::NA))
}

/// Whether a cell of a model holds a dynamic-array formula (KERNEL.8, slice
/// b): [`Env::dynamic`]'s answer for `Calc` and an engine's machine alike.
pub(crate) fn is_dynamic(model: &Workbook, (s, row, col): CellId) -> bool {
    matches!(
        model
            .sheets
            .get(s)
            .and_then(|sheet| sheet.cells.get(&(row, col)))
            .map(|c| &c.content),
        Some(Content::DynamicFormula(_) | Content::DynamicSpill { .. })
    )
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::sheet::{parse_a1_range, Cell, Sheet};

    /// A model from rows of `(sheet, address, content)`, a formula's text
    /// beginning with `=`.
    fn model(rows: &[(&str, &str, &str)]) -> Workbook {
        let mut wb = Workbook::new();
        for (sheet, addr, text) in rows {
            let i = wb.ensure_sheet(sheet);
            let range = parse_a1_range(addr).unwrap();
            let content = if let Some(f) = text.strip_prefix('=') {
                Content::Formula(f.to_string())
            } else if let Ok(n) = text.parse::<f64>() {
                Content::Number(n)
            } else if *text == "TRUE" || *text == "FALSE" {
                Content::Bool(*text == "TRUE")
            } else if let Some(e) = ErrorKind::parse(text) {
                Content::Error(e.text().to_string())
            } else {
                Content::Text(text.to_string())
            };
            wb.sheets[i].set(range.top, range.left, Cell { content, style: 0 });
        }
        wb
    }

    fn at(calc: &Calc<'_>, sheet: &str, addr: &str) -> Computed {
        let i = calc.model().find_sheet(sheet).unwrap();
        let r = parse_a1_range(addr).unwrap();
        calc.computed((i, r.top, r.left)).cloned().unwrap()
    }

    fn number(calc: &Calc<'_>, sheet: &str, addr: &str) -> f64 {
        match at(calc, sheet, addr) {
            Computed::Value(Value::Number(n)) => n,
            other => panic!("{sheet}!{addr}: {other:?}"),
        }
    }

    #[test]
    fn arithmetic_runs_in_dependency_order_whatever_the_cell_order() {
        let wb = model(&[
            ("S", "A3", "=A2*2"),
            ("S", "A2", "=A1+1"),
            ("S", "A1", "5"),
            ("S", "B1", "=A3-A2"),
        ]);
        let lib = library::language();
        let calc = Calc::run(&wb, &lib);
        assert_eq!(number(&calc, "S", "A2"), 6.0);
        assert_eq!(number(&calc, "S", "A3"), 12.0);
        assert_eq!(number(&calc, "S", "B1"), 6.0);
        assert_eq!(calc.formulas(), 3);
        assert!(calc.cycles().is_empty());
        assert_eq!(
            calc.progress(),
            Progress {
                evaluated: 3,
                of: 3,
                done: true
            }
        );
    }

    #[test]
    fn a_budget_never_changes_a_value() {
        let wb = model(&[
            ("S", "A1", "1"),
            ("S", "A2", "=A1+1"),
            ("S", "A3", "=A2+1"),
            ("S", "A4", "=A3+1"),
            ("S", "A5", "=SUM(A1:A4)"),
        ]);
        let lib = library::language();
        let whole = Calc::run(&wb, &lib);
        let mut chunked = Calc::new(&wb, &lib);
        let p1 = chunked.step(2);
        assert_eq!(
            p1,
            Progress {
                evaluated: 2,
                of: 4,
                done: false
            }
        );
        assert!(
            chunked.computed((0, 4, 1)).is_none(),
            "A4 waits for the next step"
        );
        let p2 = chunked.step(1);
        assert!(!p2.done);
        let p3 = chunked.step(5);
        assert!(p3.done && p3.evaluated == 4);
        for id in whole.formula_cells() {
            assert_eq!(whole.computed(id), chunked.computed(id), "{id:?}");
        }
        assert_eq!(number(&whole, "S", "A5"), 10.0);
        chunked.restart();
        assert!(chunked.computed((0, 5, 1)).is_none());
        chunked.step(0);
        assert_eq!(number(&chunked, "S", "A5"), 10.0);
    }

    #[test]
    fn a_cycle_is_named_and_what_reads_it_is_not_computed() {
        let wb = model(&[
            ("S", "A1", "=B1+1"),
            ("S", "B1", "=A1+1"),
            ("S", "C1", "=A1*2"),
            ("S", "D1", "=D1"),
            ("S", "E1", "=1+1"),
        ]);
        let lib = library::language();
        let calc = Calc::run(&wb, &lib);
        assert_eq!(
            calc.cycles(),
            &[vec![(0, 1, 1), (0, 1, 2)], vec![(0, 1, 4)]]
        );
        assert_eq!(at(&calc, "S", "A1"), Computed::NotComputed(Reason::Cycle));
        assert_eq!(at(&calc, "S", "C1"), Computed::NotComputed(Reason::Cycle));
        assert_eq!(at(&calc, "S", "D1"), Computed::NotComputed(Reason::Cycle));
        assert_eq!(number(&calc, "S", "E1"), 2.0);
        let refusal = calc.first_refusal().unwrap();
        assert_eq!(refusal.id, "calc-cycle");
        assert!(refusal.text.contains("S!A1, S!B1"), "{}", refusal.text);
        assert_eq!(
            calc.progress(),
            Progress {
                evaluated: 5,
                of: 5,
                done: true
            }
        );
    }

    #[test]
    fn a_cell_that_reads_itself_through_another_sheet_s_twin_is_no_cycle() {
        // The engine's previous frame: Screen.last holds values only.
        let wb = model(&[
            ("Screen", "B2", "=Screen.last!B2+1"),
            ("Screen.last", "B2", "41"),
        ]);
        let lib = library::language();
        let calc = Calc::run(&wb, &lib);
        assert!(calc.cycles().is_empty());
        assert_eq!(number(&calc, "Screen", "B2"), 42.0);
    }

    #[test]
    fn ranges_names_and_spans_are_read_as_excel_reads_them() {
        let mut wb = model(&[
            ("Model", "B1", "1200"),
            ("Model", "B2", "800"),
            ("Model", "B3", "=B1-B2"),
            ("Model", "C1", "x"),
            ("Model", "C2", "TRUE"),
            ("Model", "D1", "=SUM(B:B)"),
            ("Model", "D2", "=SUM(B1:C2)"),
            ("Model", "D3", "=SUM(B1,\"5\",TRUE,C2)"),
            ("Model", "D4", "=B3*Rate"),
            ("Model", "D5", "=Local+1"),
            ("Model", "D6", "=SUM(Model:Scratch!B1)"),
            ("Model", "D7", "=Nothing"),
            ("Model", "D8", "=Missing!A1"),
            ("Model", "D9", "=MAX(B1:B3,C1)"),
            ("Model", "D10", "=MIN(C1:C2)"),
            ("Scratch", "B1", "7"),
            ("Scratch", "B2", "=Model!D1*2"),
        ]);
        wb.defined_names
            .push(("Rate".to_string(), "0.2".to_string()));
        wb.defined_names
            .push(("Model!Local".to_string(), "Model!$B$2".to_string()));
        let lib = library::language();
        let calc = Calc::run(&wb, &lib);
        assert_eq!(number(&calc, "Model", "B3"), 400.0);
        assert_eq!(
            number(&calc, "Model", "D1"),
            2400.0,
            "the formula in B3 counts"
        );
        assert_eq!(
            number(&calc, "Model", "D2"),
            2000.0,
            "a text and a truth value in a range are skipped"
        );
        assert_eq!(
            number(&calc, "Model", "D3"),
            1206.0,
            "given directly they are coerced; C2 is read through a reference and skipped"
        );
        assert_eq!(number(&calc, "Model", "D4"), 80.0);
        assert_eq!(number(&calc, "Model", "D5"), 801.0);
        assert_eq!(
            number(&calc, "Model", "D6"),
            1207.0,
            "a span sums both sheets"
        );
        assert_eq!(
            at(&calc, "Model", "D7"),
            Computed::Value(Value::Error(ErrorKind::Name))
        );
        assert_eq!(
            at(&calc, "Model", "D8"),
            Computed::Value(Value::Error(ErrorKind::Ref))
        );
        assert_eq!(number(&calc, "Model", "D9"), 1200.0);
        assert_eq!(
            number(&calc, "Model", "D10"),
            0.0,
            "no number in the range is 0"
        );
        assert_eq!(
            number(&calc, "Scratch", "B2"),
            4800.0,
            "a sheet reads another's formula after it is computed"
        );
    }

    #[test]
    fn an_aggregate_folds_in_the_walk_s_order_and_keeps_the_first_answer() {
        // KERNEL.25: SUM, MIN, MAX, AND and OR fold as they walk instead of
        // collecting. The order is a range's walk, row by row: these four
        // cells sum to 0 by rows and to 1 by columns, as doubles.
        let wb = model(&[
            ("S", "A1", "1E16"),
            ("S", "B1", "1"),
            ("S", "A2", "-1E16"),
            ("S", "B2", "0"),
            ("S", "C1", "#N/A"),
            ("S", "C2", "=FACT(3)"),
            ("S", "D1", "=SUM(A1:B2)"),
            ("S", "D2", "=SUM(C1:C2)"),
            ("S", "D3", "=SUM(C1+0,C2:C2)"),
            ("S", "D4", "=AND(FALSE,C1)"),
            ("S", "D5", "=OR(TRUE,C1)"),
            ("S", "D6", "=MIN(Z1:Z3)"),
            ("S", "D7", "=MAX(Z1:Z3)"),
            ("S", "D8", "=MIN(A1:B2)"),
            ("S", "D9", "=MAX(A1:B2,-5)"),
            ("S", "D10", "=OR(A2:B2)"),
            ("S", "D11", "=AND(A1:B1,TRUE)"),
        ]);
        let lib = library::language();
        let calc = Calc::run(&wb, &lib);
        let err = |k| Computed::Value(Value::Error(k));
        let by_rows = [1e16, 1.0, -1e16, 0.0].iter().fold(-0.0, |a, b| a + b);
        let by_columns = [1e16, -1e16, 1.0, 0.0].iter().fold(-0.0, |a, b| a + b);
        assert_ne!(
            by_rows, by_columns,
            "the fixture tells the two orders apart"
        );
        assert_eq!(number(&calc, "S", "D1"), by_rows);
        assert_eq!(number(&calc, "S", "D1"), 0.0);
        assert_eq!(
            at(&calc, "S", "D2"),
            Computed::NotComputed(Reason::Function("FACT".to_string())),
            "a range is walked to its end: the cell not computed answers after the error"
        );
        assert_eq!(
            at(&calc, "S", "D3"),
            err(ErrorKind::NA),
            "an error given directly answers before a later range is walked"
        );
        assert_eq!(at(&calc, "S", "D4"), err(ErrorKind::NA), "AND is strict");
        assert_eq!(at(&calc, "S", "D5"), err(ErrorKind::NA), "OR is strict");
        assert_eq!(number(&calc, "S", "D6"), 0.0, "MIN of no number");
        assert_eq!(number(&calc, "S", "D7"), 0.0, "MAX of no number");
        assert_eq!(number(&calc, "S", "D8"), -1e16);
        assert_eq!(number(&calc, "S", "D9"), 1e16);
        assert_eq!(
            at(&calc, "S", "D10"),
            Computed::Value(Value::Bool(true)),
            "-1E16 is true, 0 false"
        );
        assert_eq!(at(&calc, "S", "D11"), Computed::Value(Value::Bool(true)));
    }

    #[test]
    fn errors_are_values_and_propagate_as_excel_propagates_them() {
        let wb = model(&[
            ("S", "A1", "=1/0"),
            ("S", "A2", "=A1+1"),
            ("S", "A3", "=IF(A1>0,1,2)"),
            ("S", "A4", "=IF(TRUE,1,A1)"),
            ("S", "A5", "=SUM(A1:A2)"),
            ("S", "A6", "#N/A"),
            ("S", "A7", "=A6"),
            ("S", "A8", "=\"x\"+1"),
            ("S", "A9", "=-\"5\""),
            ("S", "A10", "=50%"),
            ("S", "A11", "=1&2"),
            ("S", "A12", "=2^10"),
            ("S", "A13", "=B99"),
            ("S", "A14", "=#REF!+1"),
        ]);
        let lib = library::language();
        let calc = Calc::run(&wb, &lib);
        let err = |k| Computed::Value(Value::Error(k));
        assert_eq!(at(&calc, "S", "A1"), err(ErrorKind::Div0));
        assert_eq!(at(&calc, "S", "A2"), err(ErrorKind::Div0));
        assert_eq!(
            at(&calc, "S", "A3"),
            err(ErrorKind::Div0),
            "the condition's error"
        );
        assert_eq!(
            number(&calc, "S", "A4"),
            1.0,
            "the branch not taken is not computed"
        );
        assert_eq!(at(&calc, "S", "A5"), err(ErrorKind::Div0));
        assert_eq!(
            at(&calc, "S", "A7"),
            err(ErrorKind::NA),
            "an error typed into a cell"
        );
        assert_eq!(at(&calc, "S", "A8"), err(ErrorKind::Value));
        assert_eq!(number(&calc, "S", "A9"), -5.0);
        assert_eq!(number(&calc, "S", "A10"), 0.5);
        assert_eq!(
            at(&calc, "S", "A11"),
            Computed::Value(Value::Text("12".to_string()))
        );
        assert_eq!(number(&calc, "S", "A12"), 1024.0);
        assert_eq!(
            number(&calc, "S", "A13"),
            0.0,
            "an empty cell read is 0 in the cell"
        );
        assert_eq!(at(&calc, "S", "A14"), err(ErrorKind::Ref));
    }

    #[test]
    fn a_range_where_one_value_is_wanted_meets_the_cell_in_a_legacy_formula() {
        // KERNEL.8, slice b: implicit intersection. A legacy formula reads an
        // area where one value is wanted at the cell: the area's cell in its
        // row for one column, in its column for one row, in both for any
        // other, on any sheet, `#VALUE!` where the area holds none there and
        // `#REF!` where it spans sheets, as Excel's save of the fixture holds
        // them; `SINGLE`, Excel's `@`, the same in a formula of
        // either kind, answering a reference; a dynamic-array formula's range
        // an array, which spills (KERNEL.55), blocked by the formula below it,
        // and `INDEX`'s pick of many cells spilled into the empty cells below.
        let mut wb = model(&[
            ("S", "A1", "1"),
            ("S", "A2", "2"),
            ("S", "A3", "3"),
            ("S", "A4", "4"),
            ("S", "A5", "5"),
            ("S", "B1", "10"),
            ("S", "C1", "20"),
            ("S", "D1", "30"),
            ("S", "E1", "40"),
            ("T", "A1", "100"),
            ("T", "A2", "200"),
            ("T", "A3", "300"),
            ("T", "B3", "310"),
            ("T", "C3", "320"),
            ("U", "A1", "1000"),
            ("S", "G3", "=A1:A5"),
            ("S", "G7", "=A1:A5"),
            ("S", "D2", "=B1:E1"),
            ("S", "H2", "=B1:E1"),
            ("S", "B3", "=T!A1:C5"),
            ("S", "C9", "=T!A1:C5"),
            ("S", "G2", "=T!A1:A3*2"),
            ("S", "G4", "=IF(A1:A5>3,\"y\",\"n\")"),
            ("S", "G5", "=A:A+1"),
            ("S", "G1", "=SUM(T:U!A1)+T:U!A1"),
            ("S", "I2", "=Column*10"),
            ("S", "H4", "=SINGLE(A1:A5)"),
            ("S", "H5", "=ROW(SINGLE(A1:A5))"),
            ("S", "H6", "=SINGLE(7)"),
            ("S", "H9", "=SINGLE(A1:A5)"),
            ("S", "H7", "=SINGLE(T:U!A1)"),
            ("S", "I3", "=INDEX(A1:B5,0,1)"),
            ("S", "J3", "=A1:A5"),
            ("S", "J4", "=SINGLE(A1:A5)"),
            ("S", "J5", "=INDEX(A1:B5,0,1)"),
        ]);
        wb.defined_names
            .push(("Column".to_string(), "S!$A$1:$A$5".to_string()));
        let s = wb.find_sheet("S").unwrap();
        for (row, col) in [(3, 10), (4, 10), (5, 10)] {
            let text = match &wb.sheets[s].cells[&(row, col)].content {
                Content::Formula(t) => t.clone(),
                other => panic!("{other:?}"),
            };
            wb.sheets[s].set(
                row,
                col,
                Cell {
                    content: Content::DynamicFormula(text),
                    style: 0,
                },
            );
        }
        let lib = library::language();
        let calc = Calc::run(&wb, &lib);
        let value_error = Computed::Value(Value::Error(ErrorKind::Value));
        assert_eq!(number(&calc, "S", "G3"), 3.0, "a column, met at row 3");
        assert_eq!(at(&calc, "S", "G7"), value_error, "row 7 is not in A1:A5");
        assert_eq!(number(&calc, "S", "D2"), 30.0, "a row, met at column D");
        assert_eq!(at(&calc, "S", "H2"), value_error, "column H is not in B:E");
        assert_eq!(number(&calc, "S", "B3"), 310.0, "a block, at both, on T");
        assert_eq!(at(&calc, "S", "C9"), value_error, "row 9 is not in it");
        assert_eq!(number(&calc, "S", "G2"), 400.0, "met, then doubled");
        assert_eq!(
            at(&calc, "S", "G4"),
            Computed::Value(Value::Text("y".to_string())),
            "a condition met at row 4"
        );
        assert_eq!(number(&calc, "S", "G5"), 6.0, "a whole column, at row 5");
        assert_eq!(
            at(&calc, "S", "G1"),
            Computed::Value(Value::Error(ErrorKind::Ref)),
            "a span of sheets"
        );
        assert_eq!(
            number(&calc, "S", "I2"),
            20.0,
            "a name's range, met at row 2"
        );
        assert_eq!(number(&calc, "S", "H4"), 4.0);
        assert_eq!(number(&calc, "S", "H5"), 5.0, "SINGLE answers a reference");
        assert_eq!(number(&calc, "S", "H6"), 7.0);
        assert_eq!(at(&calc, "S", "H9"), value_error);
        assert_eq!(
            at(&calc, "S", "H7"),
            Computed::Value(Value::Error(ErrorKind::Ref)),
            "SINGLE over a span of sheets"
        );
        assert_eq!(number(&calc, "S", "I3"), 3.0, "INDEX's column, met");
        assert_eq!(
            at(&calc, "S", "J3"),
            Computed::Value(Value::Error(ErrorKind::Spill)),
            "a dynamic formula's range is an array, its spill blocked (KERNEL.55)"
        );
        assert_eq!(number(&calc, "S", "J4"), 4.0, "SINGLE in a dynamic formula");
        for (addr, want) in [("J5", 1.0), ("J6", 2.0), ("J9", 5.0)] {
            assert_eq!(
                number(&calc, "S", addr),
                want,
                "INDEX's column in a dynamic formula, spilled"
            );
        }
    }

    #[test]
    fn a_name_whose_text_is_not_read_is_refused_at_the_load() {
        // KERNEL.8, slice b: a defined name a formula reads, or a name such a
        // name reads, whose text the parse does not read, refused with the
        // constructs, so that no cell of an engine's grid holds the reason;
        // a name read, or none the workbook defines, refused by nothing.
        let refusal = |names: &[(&str, &str)], formula: &str| {
            let mut wb = model(&[("S", "A1", "1"), ("S", "B1", formula)]);
            for (n, t) in names {
                wb.defined_names.push((n.to_string(), t.to_string()));
            }
            let lib = library::language();
            let calc = Calc::new(&wb, &lib);
            calc.first_refusal()
        };
        let read = refusal(&[("Rate", "S!$A$1")], "=Rate*2");
        assert!(read.is_none(), "{read:?}");
        assert!(refusal(&[("Rate", "S!$A$1")], "=Missing*2").is_none());
        let unread = refusal(&[("Pair", "Odd#")], "=Pair*2").unwrap();
        assert_eq!(unread.id, "calc-construct-not-read");
        assert!(
            unread.text.contains("S!B1") && unread.text.contains("Odd#"),
            "{}",
            unread.text
        );
        let nested = refusal(&[("Outer", "Inner+1"), ("Inner", "Odd#")], "=Outer").unwrap();
        assert_eq!(nested.id, "calc-construct-not-read");
        let unused = refusal(&[("Pair", "Odd#")], "=A1*2");
        assert!(unused.is_none(), "a name no formula reads: {unused:?}");
    }

    #[test]
    fn the_day_one_functions_compute_exactly() {
        let wb = model(&[
            ("S", "A1", "=IF(1>0,\"yes\",\"no\")"),
            ("S", "A2", "=IF(0)"),
            ("S", "A3", "=IF(FALSE,1)"),
            ("S", "A4", "=IF(TRUE,,1)"),
            ("S", "A5", "=AND(TRUE,1,\"TRUE\")"),
            ("S", "A6", "=AND(TRUE,0)"),
            ("S", "A7", "=OR(FALSE,0,2)"),
            ("S", "A8", "=NOT(0)"),
            ("S", "A9", "=AND(\"yes\")"),
            ("S", "A10", "=ABS(-3.5)"),
            ("S", "A11", "=INT(-1.5)"),
            ("S", "A12", "=MOD(-7,3)"),
            ("S", "A13", "=MOD(7,-3)"),
            ("S", "A14", "=MOD(7,0)"),
            ("S", "A15", "=ROW()"),
            ("S", "A16", "=COLUMN()"),
            ("S", "A17", "=ROW(C9)"),
            ("S", "A18", "=COLUMN(C9:D10)"),
            ("S", "A19", "=CHOOSE(2,\"a\",\"b\",\"c\")"),
            ("S", "A20", "=CHOOSE(4,\"a\",\"b\",\"c\")"),
            ("S", "A21", "=SIN(0)"),
            ("S", "A22", "=MOD(5.5,2)"),
            ("S", "A23", "=CHOOSE(1.9,\"a\",\"b\")"),
            ("S", "A24", "=AND()"),
            ("S", "B1", "=FACT(3)"),
            ("S", "B2", "=B1+1"),
            ("S", "B3", "=SUM(A1:A2 B1)"),
            ("S", "B4", "={1,2}"),
            ("S", "B5", "=A1:A2*2"),
        ]);
        let lib = library::language();
        let calc = Calc::run(&wb, &lib);
        let text = |s: &str| Computed::Value(Value::Text(s.to_string()));
        let truth = |b: bool| Computed::Value(Value::Bool(b));
        let err = |k| Computed::Value(Value::Error(k));
        assert_eq!(at(&calc, "S", "A1"), text("yes"));
        assert_eq!(
            at(&calc, "S", "A2"),
            err(ErrorKind::Value),
            "IF wants a branch"
        );
        assert_eq!(at(&calc, "S", "A3"), truth(false), "no else is FALSE");
        assert_eq!(number(&calc, "S", "A4"), 0.0, "a branch left out is 0");
        assert_eq!(at(&calc, "S", "A5"), truth(true));
        assert_eq!(at(&calc, "S", "A6"), truth(false));
        assert_eq!(at(&calc, "S", "A7"), truth(true));
        assert_eq!(at(&calc, "S", "A8"), truth(true));
        assert_eq!(at(&calc, "S", "A9"), err(ErrorKind::Value));
        assert_eq!(number(&calc, "S", "A10"), 3.5);
        assert_eq!(number(&calc, "S", "A11"), -2.0);
        assert_eq!(number(&calc, "S", "A12"), 2.0, "the divisor's sign");
        assert_eq!(number(&calc, "S", "A13"), -2.0);
        assert_eq!(at(&calc, "S", "A14"), err(ErrorKind::Div0));
        assert_eq!(number(&calc, "S", "A15"), 15.0);
        assert_eq!(number(&calc, "S", "A16"), 1.0);
        assert_eq!(number(&calc, "S", "A17"), 9.0);
        assert_eq!(number(&calc, "S", "A18"), 3.0);
        assert_eq!(at(&calc, "S", "A19"), text("b"));
        assert_eq!(at(&calc, "S", "A20"), err(ErrorKind::Value));
        assert_eq!(number(&calc, "S", "A21"), 0.0);
        assert_eq!(number(&calc, "S", "A22"), 1.5);
        assert_eq!(
            at(&calc, "S", "A23"),
            text("a"),
            "the index is read down to a whole number"
        );
        assert_eq!(
            at(&calc, "S", "A24"),
            err(ErrorKind::Value),
            "nothing to AND"
        );
        assert_eq!(
            at(&calc, "S", "B1"),
            Computed::NotComputed(Reason::Function("FACT".to_string()))
        );
        assert_eq!(
            at(&calc, "S", "B2"),
            Computed::NotComputed(Reason::Function("FACT".to_string())),
            "a reader inherits the reason"
        );
        assert_eq!(
            at(&calc, "S", "B3"),
            Computed::NotComputed(Reason::Construct("intersection".to_string()))
        );
        assert_eq!(
            at(&calc, "S", "B4"),
            Computed::Value(Value::Number(1.0)),
            "an array constant's first value, in a legacy formula (KERNEL.8, slice f)"
        );
        assert_eq!(
            at(&calc, "S", "B5"),
            err(ErrorKind::Value),
            "a range where one value is wanted met at the cell, row 5 outside A1:A2"
        );
        assert_eq!(
            calc.missing_functions(),
            vec![((0, 1, 2), "FACT".to_string())]
        );
        let refusal = calc.first_refusal().unwrap();
        assert_eq!(refusal.id, "calc-function-not-computed");
        assert!(
            refusal.text.contains("FACT") && refusal.text.contains("S!B1"),
            "{}",
            refusal.text
        );
        assert_eq!(at(&calc, "S", "A1").spell(), "\"yes\"");
        assert_eq!(at(&calc, "S", "B1").spell(), "(not-computed \"FACT\")");
    }

    #[test]
    fn a_construct_not_read_is_the_third_thing_a_load_refuses() {
        // KERNEL.22: after the cycle and the function, a formula whose text
        // this version does not read, naming the construct.
        let wb = model(&[
            ("S", "A1", "=1+1"),
            ("S", "B2", "=Odd#"),
            ("S", "C3", "=@A1"),
        ]);
        let lib = library::language();
        let calc = Calc::new(&wb, &lib);
        let r = calc.first_refusal().unwrap();
        assert_eq!(r.id, "calc-construct-not-read");
        assert!(
            r.text.contains("S!B2") && r.text.contains("reaches Odd#"),
            "{}",
            r.text
        );
        // A function the library does not hold comes first.
        let wb = model(&[("S", "B2", "=Odd#"), ("S", "C3", "=NOW()")]);
        let calc = Calc::new(&wb, &lib);
        assert_eq!(
            calc.first_refusal().unwrap().id,
            "calc-function-not-computed"
        );
        let wb = model(&[("S", "A1", "=1+1")]);
        assert!(Calc::new(&wb, &lib).first_refusal().is_none());
    }

    #[test]
    fn a_chain_of_operators_evaluates_left_to_right_however_long() {
        // VLA.1: a chain goes round a loop instead of recursing once a
        // link, read as Excel reads it, left to right: the values the
        // recursion gave, at every length Excel holds.
        let long = format!("={}1", "1+".repeat(4095));
        let wb = model(&[
            ("S", "A1", "=10-1-1-1"),
            ("S", "A2", "=2^3^2"),
            ("S", "A3", "=1+2*3-4/2+5"),
            ("S", "A4", "=\"a\"&\"b\"&\"c\""),
            ("S", "A5", long.as_str()),
            ("S", "A6", "=1/0+1+1"),
            ("S", "A7", "=1+1+1+1+1+1+1+1+1+1+1+1+1+1+1+1+1+1+1+1"),
        ]);
        let lib = library::language();
        let calc = Calc::run(&wb, &lib);
        assert_eq!(number(&calc, "S", "A1"), 7.0);
        assert_eq!(number(&calc, "S", "A2"), 64.0);
        assert_eq!(number(&calc, "S", "A3"), 10.0);
        assert_eq!(
            at(&calc, "S", "A4"),
            Computed::Value(Value::Text("abc".to_string()))
        );
        assert_eq!(number(&calc, "S", "A5"), 4096.0);
        assert_eq!(
            at(&calc, "S", "A6"),
            Computed::Value(Value::Error(ErrorKind::Div0))
        );
        // Past the sixteen links held in place, the rest on the heap.
        assert_eq!(number(&calc, "S", "A7"), 20.0);
    }

    #[test]
    fn excel_s_two_limits_are_refused_by_their_own_ids() {
        // VLA.1: past Excel's nesting or length limit an engine's load is
        // refused by the limit's own id, and frazaro calc, which never
        // refuses, shows the label in the cell.
        let deep = format!("={}1{}", "(".repeat(65), ")".repeat(65));
        let long = format!("={}1", "1+".repeat(4096));
        let lib = library::language();
        let wb = model(&[("S", "A1", deep.as_str())]);
        let r = Calc::new(&wb, &lib).first_refusal().unwrap();
        assert_eq!(r.id, "calc-formula-too-deep");
        assert_eq!(
            r.text,
            "The formula in S!A1 nests more than 64 levels deep, more than Excel itself allows, so the formula cannot be computed."
        );
        let wb = model(&[("S", "A1", long.as_str())]);
        let r = Calc::new(&wb, &lib).first_refusal().unwrap();
        assert_eq!(r.id, "calc-formula-too-long");
        assert_eq!(
            r.text,
            "The formula in S!A1 is longer than 8,192 characters, more than Excel itself allows, so the formula cannot be computed."
        );
        let calc = Calc::run(&wb, &lib);
        assert_eq!(
            at(&calc, "S", "A1"),
            Computed::NotComputed(Reason::Construct(formula::TOO_LONG.to_string()))
        );
    }

    #[test]
    fn a_logical_test_coerces_a_text_read_through_a_reference() {
        // Measured: Excel's save of the subset fixture
        // (scripts/recalc/subset_saved.xlsx, logic!C3, C6 and C13): IF
        // coerces the text TRUE found in a cell as it coerces a typed one and
        // refuses another text; AND reads the reference as a range and skips
        // the text. NOT follows IF's rule, which scripts/recalc/edges.txt
        // holds for an Excel save to witness.
        let mut wb = model(&[
            ("S", "C1", "=IF(A1,1,2)"),
            ("S", "C2", "=IF(A2,1,2)"),
            ("S", "C3", "=AND(A1)"),
            ("S", "C4", "=NOT(A1)"),
        ]);
        for (row, text) in [(1, "TRUE"), (2, "x")] {
            wb.sheets[0].set(
                row,
                1,
                Cell {
                    content: Content::Text(text.to_string()),
                    style: 0,
                },
            );
        }
        let lib = library::language();
        let calc = Calc::run(&wb, &lib);
        let err = |k| Computed::Value(Value::Error(k));
        assert_eq!(number(&calc, "S", "C1"), 1.0);
        assert_eq!(at(&calc, "S", "C2"), err(ErrorKind::Value));
        assert_eq!(at(&calc, "S", "C3"), err(ErrorKind::Value));
        assert_eq!(at(&calc, "S", "C4"), Computed::Value(Value::Bool(false)));
    }

    #[test]
    fn shared_and_dynamic_formulas_are_read_through_their_masters() {
        let mut wb = Workbook::new();
        let i = wb.ensure_sheet("Output");
        let sheet: &mut Sheet = &mut wb.sheets[i];
        sheet.set(
            2,
            2,
            Cell {
                content: Content::Number(5.0),
                style: 0,
            },
        );
        sheet.set(
            3,
            2,
            Cell {
                content: Content::Formula("B2*2".to_string()),
                style: 0,
            },
        );
        sheet.set_formula(parse_a1_range("C2:C4").unwrap(), "B2+B3", 0);
        sheet.set_formula_dynamic(parse_a1_range("D2:D3").unwrap(), "B2*10", 0);
        let lib = library::language();
        let calc = Calc::run(&wb, &lib);
        assert_eq!(
            calc.formula((0, 3, 3)),
            Some("B3+B4"),
            "the child's text is the master's moved"
        );
        assert_eq!(number(&calc, "Output", "C2"), 15.0);
        assert_eq!(number(&calc, "Output", "C3"), 10.0);
        assert_eq!(number(&calc, "Output", "C4"), 0.0);
        assert_eq!(number(&calc, "Output", "D2"), 50.0);
        assert_eq!(number(&calc, "Output", "D3"), 100.0);
        assert_eq!(calc.spell_cell((0, 3, 3)), "Output!C3");
    }

    /// A model of rows and, over it, legacy array formulas each entered over a
    /// block (KERNEL.55), `(sheet, range, formula)`: the anchor the range's
    /// first cell, every other a 0 until it is laid, as a save holds a value at
    /// each.
    fn blocked(rows: &[(&str, &str, &str)], blocks: &[(&str, &str, &str)]) -> Workbook {
        let mut wb = model(rows);
        for (sheet, range, formula) in blocks {
            let i = wb.ensure_sheet(sheet);
            let range = parse_a1_range(range).unwrap();
            for row in range.top..=range.bottom {
                for col in range.left..=range.right {
                    let content = Content::Number(0.0);
                    wb.sheets[i].set(row, col, Cell { content, style: 0 });
                }
            }
            let text = formula.strip_prefix('=').unwrap().to_string();
            let content = Content::ArrayBlock { text, range };
            wb.sheets[i].set(range.top, range.left, Cell { content, style: 0 });
        }
        wb
    }

    /// A cell's answer as the door's row spells it, a reason by its debug.
    fn shown(calc: &Calc<'_>, sheet: &str, addr: &str) -> String {
        match at(calc, sheet, addr) {
            Computed::Value(v) => v.spell(),
            Computed::NotComputed(r) => format!("{r:?}"),
        }
    }

    /// The probe's blocks in small (KERNEL.55's second slice), and readers of
    /// them by a cell and by a range.
    fn the_blocks() -> Workbook {
        blocked(
            &[
                ("D", "A1", "3"),
                ("D", "A2", "1"),
                ("D", "A3", "4"),
                ("D", "B1", "1"),
                ("D", "C1", "2"),
                ("D", "D1", "3"),
                ("D", "B2", "4"),
                ("D", "C2", "5"),
                ("D", "D2", "6"),
                ("S", "A7", "=SUM(A1:A3)"),
                ("S", "A8", "=A4"),
                ("S", "A9", "=SUM(B1:D2)"),
                ("S", "A10", "=J3"),
                ("S", "A11", "=ISNA(K3)"),
            ],
            &[
                ("S", "A1:A5", "=D!A1:A3*2"),
                ("S", "B1:D2", "=D!B1:D1*2"),
                ("S", "E1:G3", "=D!A1:A3*10"),
                ("S", "H1:I2", "=5"),
                ("S", "J1:K3", "=D!B1:D2"),
                ("S", "L1:L2", "=1/0"),
            ],
        )
    }

    #[test]
    fn a_block_is_computed_once_and_laid_into_its_cells_as_excel_lays_one() {
        // KERNEL.55's second slice, as Excel's save of its probe lays one: a
        // column longer than its block cut short and one shorter #N/A past
        // it; a row repeated down a block and a column across one; one value
        // everywhere; a reference's 2-by-3 over a 3-by-2, #N/A past it; an
        // error at every place; and each read as computed, by its address
        // and in a range, the file's 0 never.
        let wb = the_blocks();
        let lib = library::language();
        let calc = Calc::run(&wb, &lib);
        let na = "(error \"#N/A\")";
        let div = "(error \"#DIV/0!\")";
        for (addr, want) in [
            ("A1", "6"),
            ("A2", "2"),
            ("A3", "8"),
            ("A4", na),
            ("A5", na),
            ("B1", "2"),
            ("C1", "4"),
            ("D1", "6"),
            ("B2", "2"),
            ("C2", "4"),
            ("D2", "6"),
            ("E1", "30"),
            ("F1", "30"),
            ("G1", "30"),
            ("E2", "10"),
            ("F2", "10"),
            ("G3", "40"),
            ("H1", "5"),
            ("I2", "5"),
            ("J1", "1"),
            ("K1", "2"),
            ("J2", "4"),
            ("K2", "5"),
            ("J3", na),
            ("K3", na),
            ("L1", div),
            ("L2", div),
            ("A7", "16"),
            ("A8", na),
            ("A9", "24"),
            ("A10", na),
            ("A11", "true"),
        ] {
            assert_eq!(shown(&calc, "S", addr), want, "S!{addr}");
        }
        // The cells computed: eleven formulas and twenty-six places, each
        // place's formula its anchor's, as Excel holds one at every cell.
        let s = calc.model().find_sheet("S").unwrap();
        assert_eq!(calc.formulas(), 11);
        assert_eq!(calc.cells().len(), 37);
        assert_eq!(calc.progress().of, 37);
        assert!(calc.cells().windows(2).all(|w| w[0] < w[1]));
        assert_eq!(calc.formula((s, 2, 1)), Some("D!A1:A3*2"));
        assert_eq!(calc.formula((s, 3, 11)), Some("D!B1:D2"));
        assert!(calc.cycles().is_empty());
    }

    #[test]
    fn a_budget_never_changes_a_block_s_value() {
        // KERNEL.55's second slice: an anchor's array kept from its step for
        // its places', whatever steps between, and every place after it.
        let wb = the_blocks();
        let lib = library::language();
        let whole = Calc::run(&wb, &lib);
        for budget in [1, 3, 7] {
            let mut calc = Calc::new(&wb, &lib);
            while !calc.step(budget).done {}
            for id in whole.cells() {
                assert_eq!(calc.computed(id), whole.computed(id), "{id:?} at {budget}");
            }
            calc.restart();
            calc.step(0);
            for id in whole.cells() {
                assert_eq!(calc.computed(id), whole.computed(id), "{id:?} again");
            }
        }
    }

    #[test]
    fn a_block_s_places_are_the_cells_its_sheet_holds_but_a_formula_s_own() {
        // KERNEL.55's second slice: a block's cell the sheet does not hold is
        // no place, read as nothing, and one holding a formula of its own
        // keeps it; a reference past the block is read no further than it.
        let mut wb = model(&[("S", "A3", "=99"), ("S", "A4", "0"), ("S", "C1", "=A2")]);
        let range = parse_a1_range("A1:A4").unwrap();
        let text = "SEQUENCE(4)".to_string();
        let content = Content::ArrayBlock { text, range };
        wb.sheets[0].set(1, 1, Cell { content, style: 0 });
        let lib = library::language();
        let calc = Calc::run(&wb, &lib);
        assert_eq!(shown(&calc, "S", "A1"), "1");
        assert_eq!(calc.computed((0, 2, 1)), None, "A2 no place");
        assert_eq!(shown(&calc, "S", "A3"), "99");
        assert_eq!(shown(&calc, "S", "A4"), "4");
        assert_eq!(shown(&calc, "S", "C1"), "0");
        assert_eq!(calc.cells().len(), 4);
    }

    #[test]
    fn a_block_reading_its_own_cells_is_a_cycle_and_an_anchor_s_reason_its_places() {
        // KERNEL.55's second slice: a block's cells are read after its anchor,
        // so that one reading them is a cycle, as Excel's circular reference
        // is; and an anchor not computed gives each place its reason.
        let wb = blocked(&[], &[("S", "A1:A3", "=SUM(A1:A3)")]);
        let lib = library::language();
        let calc = Calc::run(&wb, &lib);
        assert_eq!(calc.cycles(), [vec![(0, 1, 1), (0, 2, 1), (0, 3, 1)]]);
        assert_eq!(at(&calc, "S", "A3"), Computed::NotComputed(Reason::Cycle));
        let wb = blocked(&[], &[("S", "A1:B1", "=NOSUCH(1)")]);
        let calc = Calc::run(&wb, &lib);
        let reason = Computed::NotComputed(Reason::Function("NOSUCH".to_string()));
        assert_eq!(at(&calc, "S", "A1"), reason);
        assert_eq!(at(&calc, "S", "B1"), reason);
    }

    #[test]
    fn a_spill_read_through_its_hash_is_a_dynamic_anchor_s_and_no_other_cell_s() {
        // KERNEL.55's second slice: `A1#`, read as the file spells it,
        // `ANCHORARRAY(A1)`. A value's and a legacy formula's, #REF!, as
        // Excel's save of the probe answers a value's and a spilled cell's; a
        // dynamic formula's its spill (the third slice), here over cells the
        // model does not hold; a legacy array formula's, which no save shows, labelled by
        // the function's name, as is a reference to no one cell; a spill whose
        // anchor moved off the sheet #REF!; and `#SPILL!`, a value of the
        // language, ERROR.TYPE 9.
        let mut wb = model(&[
            ("S", "A1", "5"),
            ("S", "B1", "=1+1"),
            ("S", "D1", "=A1#"),
            ("S", "D2", "=B1#"),
            ("S", "D3", "=SUM(C1#)"),
            ("S", "D4", "=ROWS(E1#)"),
            ("S", "D5", "=#REF!#"),
            ("S", "D6", "=ANCHORARRAY(A1:A2)"),
            ("S", "D7", "=ISERROR(A1#)"),
            ("S", "F1", "#SPILL!"),
            ("S", "D8", "=ERROR.TYPE(F1)"),
            ("S", "D9", "=ISERROR(F1)"),
        ]);
        let dynamic = Content::DynamicFormula("SEQUENCE(3)".to_string());
        wb.sheets[0].set(
            1,
            3,
            Cell {
                content: dynamic,
                style: 0,
            },
        );
        let legacy = Content::ArrayFormula("SEQUENCE(2)".to_string());
        wb.sheets[0].set(
            1,
            5,
            Cell {
                content: legacy,
                style: 0,
            },
        );
        let lib = library::language();
        let calc = Calc::run(&wb, &lib);
        let anchorarray = format!("{:?}", Reason::Function("ANCHORARRAY".to_string()));
        for (addr, want) in [
            ("D1", "(error \"#REF!\")"),
            ("D2", "(error \"#REF!\")"),
            ("D3", "6"),
            ("D4", &anchorarray),
            ("D5", "(error \"#REF!\")"),
            ("D6", &anchorarray),
            ("D7", "true"),
            ("D8", "9"),
            ("D9", "true"),
        ] {
            assert_eq!(shown(&calc, "S", addr), want, "S!{addr}");
        }
        // The plan reads `B1#` as a read of B1, so that two cells reading each
        // other so are a cycle.
        let wb = model(&[("S", "A1", "=B1#"), ("S", "B1", "=A1")]);
        let calc = Calc::run(&wb, &lib);
        assert_eq!(calc.cycles(), [vec![(0, 1, 1), (0, 1, 2)]]);
    }

    /// A model whose `dynamic` formulas are dynamic-array formulas, each with
    /// the extent its file saved where one is given (KERNEL.55's third slice),
    /// and whose `merged` ranges are merged.
    fn spilling(
        rows: &[(&str, &str, &str)],
        dynamic: &[(&str, &str, &str, Option<&str>)],
        merged: &[(&str, &str)],
    ) -> Workbook {
        let mut wb = model(rows);
        for (sheet, addr, formula, file) in dynamic {
            let i = wb.ensure_sheet(sheet);
            let at = parse_a1_range(addr).unwrap();
            let text = formula.strip_prefix('=').unwrap().to_string();
            let content = match file {
                Some(range) => Content::DynamicSpill {
                    text,
                    range: parse_a1_range(range).unwrap(),
                },
                None => Content::DynamicFormula(text),
            };
            wb.sheets[i].set(at.top, at.left, Cell { content, style: 0 });
        }
        for (sheet, range) in merged {
            let i = wb.ensure_sheet(sheet);
            wb.sheets[i]
                .put_merge(parse_a1_range(range).unwrap())
                .unwrap();
        }
        wb
    }

    /// The probe's spills in small, none of their cells held by the sheet: an
    /// extent fixed, one a filter sets, a spill of each, and readers of them
    /// through `#`, by their addresses, by a range and by a whole column.
    fn the_spills() -> Workbook {
        spilling(
            &[
                ("D", "A1", "3"),
                ("D", "A2", "1"),
                ("D", "A3", "4"),
                ("D", "A4", "1"),
                ("D", "A5", "5"),
                ("D", "A6", "9"),
                ("D", "K1", "1"),
                // On the sheet before the spills', so that only the plan's
                // edges put them after their anchors.
                ("D", "M1", "=SUM(S!B:B)"),
                ("D", "M2", "=S!B3"),
                ("S", "A15", "=SUM(A1#)"),
                ("S", "A16", "=ROWS(A1#)"),
                ("S", "A17", "=COLUMNS(A1#)"),
                ("S", "A18", "=INDEX(A1#,3,2)"),
                ("S", "A19", "=B3"),
                ("S", "A20", "=SUM(A1:B3)"),
                ("S", "A21", "=ROWS(D1#)"),
                ("S", "A22", "=SUM(N1:O3)"),
                ("S", "A23", "=SUM(B:B)"),
                ("S", "A24", "=D5"),
                ("S", "A25", "=ISBLANK(D5)"),
                ("S", "A26", "=ROWS(X1#)"),
                ("S", "A27", "=SUM(X1#)*2"),
                // A sorted search of a column, bounded by its sheet's extent,
                // which holds the spill on a sheet holding its anchor alone.
                ("S", "A28", "=MATCH(3,E!A:A,1)"),
            ],
            &[
                ("S", "A1", "=SEQUENCE(3,2)", None),
                ("S", "X1", "=SEQUENCE(1)", None),
                ("E", "A1", "=SEQUENCE(3)", None),
                ("S", "D1", "=FILTER(D!A1:A6,D!A1:A6>D!K1)", None),
                ("S", "N1", "=A1#*10", None),
                ("S", "Q1", "=SORT(D1#,1,-1)", None),
            ],
            &[],
        )
    }

    #[test]
    fn a_dynamic_formula_spills_its_answer_into_the_cells_beside_it() {
        // KERNEL.55's third slice: the door plans each spill by its parse's
        // bound and lays its answer into its extent, each cell its place of
        // the array, as Excel's save of the item's probe lays one; `A1#` the
        // extent as a reference; a spilled cell read by its address, by a range
        // and by a whole column, though the sheet holds none of them.
        let wb = the_spills();
        let lib = library::language();
        let calc = Calc::run(&wb, &lib);
        let held = |cells: &[(&str, f64)]| {
            for (addr, want) in cells {
                assert_eq!(number(&calc, "S", addr), *want, "S!{addr}");
            }
        };
        held(&[("A1", 1.0), ("B1", 2.0), ("A2", 3.0), ("B3", 6.0)]);
        held(&[("D1", 3.0), ("D2", 4.0), ("D3", 5.0), ("D4", 9.0)]);
        held(&[("N1", 10.0), ("O1", 20.0), ("O3", 60.0)]);
        held(&[("Q1", 9.0), ("Q2", 5.0), ("Q3", 4.0), ("Q4", 3.0)]);
        held(&[("A15", 21.0), ("A16", 3.0), ("A17", 2.0), ("A18", 6.0)]);
        held(&[("A19", 6.0), ("A20", 21.0), ("A21", 4.0), ("A22", 210.0)]);
        held(&[
            ("A23", 12.0),
            ("A24", 0.0),
            ("A26", 1.0),
            ("A27", 2.0),
            ("A28", 3.0),
        ]);
        assert_eq!(number(&calc, "D", "M1"), 12.0, "a column read before");
        assert_eq!(number(&calc, "D", "M2"), 6.0, "a cell read before");
        assert_eq!(at(&calc, "S", "A25"), Computed::Value(Value::Bool(true)));
        // A cell the filter's bound reaches and its answer does not holds
        // nothing, and is no cell the recalculation computes.
        let s = wb.find_sheet("S").unwrap();
        assert_eq!(calc.computed((s, 5, 4)), None);
        assert!(!calc.cells().contains(&(s, 5, 4)));
        // Each spilled cell a cell the door computes, its anchor's text at it.
        assert!(calc.cells().contains(&(s, 3, 2)));
        assert_eq!(calc.formula((s, 3, 2)), Some("SEQUENCE(3,2)"));
        assert_eq!(calc.cells().len(), calc.formulas() + 5 + 3 + 5 + 3 + 2);
        assert_eq!(Values::spill(&calc, (s, 1, 1)), parse_a1_range("A1:B3"));
        assert_eq!(Values::spill(&calc, (s, 1, 4)), parse_a1_range("D1:D4"));
        // Computed again, the same.
        let mut again = Calc::run(&wb, &lib);
        again.restart();
        again.step(0);
        for id in calc.cells() {
            assert_eq!(again.computed(id), calc.computed(id), "{id:?}");
        }
    }

    #[test]
    fn a_budget_never_changes_a_spill_s_value() {
        let wb = the_spills();
        let lib = library::language();
        let whole = Calc::run(&wb, &lib);
        for budget in [1, 3, 7] {
            let mut calc = Calc::new(&wb, &lib);
            while !calc.step(budget).done {}
            assert_eq!(calc.cells(), whole.cells(), "budget {budget}");
            for id in whole.cells() {
                assert_eq!(calc.computed(id), whole.computed(id), "{id:?}");
            }
        }
    }

    #[test]
    fn a_spill_is_blocked_by_what_its_extent_meets_as_excel_blocks_one() {
        // As Excel's save of the probe blocks one: by a value, a formula, a
        // merged range, a legacy array formula's block, and the sheet's last
        // row and column, its anchor `#SPILL!` and its would-be cells empty;
        // and a spill reaching exactly to each edge.
        let wb = {
            let mut wb = spilling(
                &[
                    ("S", "F3", "x"),
                    ("S", "G3", "=1+1"),
                    ("S", "A24", "=SUM(F1#)"),
                    ("S", "A25", "=F1"),
                    ("S", "A26", "=ISERROR(F1)"),
                    ("S", "A27", "=ERROR.TYPE(F1)"),
                    ("S", "A28", "=IFERROR(F1,\"caught\")"),
                    ("S", "A29", "=F2"),
                    ("S", "A30", "=ISBLANK(F2)"),
                    ("S", "C1048575", "=B1048576*10"),
                    ("S", "XFB5", "=XFD5*10"),
                    ("S", "A31", "=F3"),
                    ("S", "A32", "=ROWS(F1#)"),
                    ("S", "E5", "=MATCH(1,B:B,0)"),
                    ("S", "E6", "=MATCH(2,B:B,0)"),
                ],
                &[
                    ("S", "F1", "=SEQUENCE(4)", None),
                    ("S", "G1", "=SEQUENCE(4)", None),
                    ("S", "H1", "=SEQUENCE(4)", None),
                    ("S", "L1", "=SEQUENCE(4)", None),
                    ("S", "A1048575", "=SEQUENCE(3)", None),
                    ("S", "B1048575", "=SEQUENCE(2)", None),
                    ("S", "XFC1", "=SEQUENCE(1,3)", None),
                    ("S", "XFC5", "=SEQUENCE(1,2)", None),
                    ("S", "Q1", "=SEQUENCE(4)", None),
                ],
                &[("S", "H3:I3")],
            );
            let s = wb.find_sheet("S").unwrap();
            let range = parse_a1_range("L3:L4").unwrap();
            wb.sheets[s].set(
                4,
                12,
                Cell {
                    content: Content::Number(0.0),
                    style: 0,
                },
            );
            wb.sheets[s].set(
                3,
                12,
                Cell {
                    content: Content::ArrayBlock {
                        text: "5".to_string(),
                        range,
                    },
                    style: 0,
                },
            );
            // A block whose other cells the sheet does not hold, a model's
            // built in memory, its range blocking all the same.
            let range = parse_a1_range("P3:Q4").unwrap();
            wb.sheets[s].set(
                3,
                16,
                Cell {
                    content: Content::ArrayBlock {
                        text: "6".to_string(),
                        range,
                    },
                    style: 0,
                },
            );
            wb
        };
        let lib = library::language();
        let calc = Calc::run(&wb, &lib);
        let spill = Computed::Value(Value::Error(ErrorKind::Spill));
        for addr in [
            "F1", "G1", "H1", "L1", "Q1", "A1048575", "XFC1", "A24", "A25",
        ] {
            assert_eq!(at(&calc, "S", addr), spill, "S!{addr}");
        }
        assert_eq!(at(&calc, "S", "A26"), Computed::Value(Value::Bool(true)));
        assert_eq!(number(&calc, "S", "A27"), 9.0);
        assert_eq!(
            at(&calc, "S", "A28"),
            Computed::Value(Value::Text("caught".to_string()))
        );
        assert_eq!(number(&calc, "S", "A29"), 0.0);
        assert_eq!(at(&calc, "S", "A30"), Computed::Value(Value::Bool(true)));
        assert_eq!(number(&calc, "S", "L4"), 5.0, "the block's own");
        assert_eq!(
            number(&calc, "S", "C1048575"),
            20.0,
            "the last row, exactly"
        );
        assert_eq!(number(&calc, "S", "XFB5"), 20.0, "the last column, exactly");
        // What blocks a spill is read as itself; the blocked anchor's `#` is
        // `#SPILL!`, the reference itself, as the save answers `SUM`'s; and a
        // whole column reads a spill past the cells the model holds, scanned
        // and then searched by its index, which the sheet's extent bounds.
        assert_eq!(
            at(&calc, "S", "A31"),
            Computed::Value(Value::Text("x".to_string()))
        );
        assert_eq!(at(&calc, "S", "A32"), spill);
        assert_eq!(number(&calc, "S", "E5"), 1_048_575.0);
        assert_eq!(number(&calc, "S", "E6"), 1_048_576.0);
    }

    #[test]
    fn two_spills_that_may_meet_are_labelled_and_one_in_another_s_box_is_not() {
        // Two whose cells may meet, neither anchor in the other's box, which
        // Excel settles by its order: labelled, and what reads them. One
        // anchor in another's box, which no order settles: the outer blocked
        // where its answer reaches the inner's anchor, a formula, and each
        // spilled where it does not.
        let wb = spilling(
            &[
                ("D", "A1", "3"),
                ("D", "A2", "1"),
                ("D", "A3", "4"),
                ("D", "A4", "1"),
                ("D", "A5", "5"),
                ("D", "A6", "9"),
                ("S", "A31", "=IFERROR(K1,\"blocked\")"),
                ("S", "T1", "=SUM(P1:P4)"),
                ("S", "T2", "=SUM(R3:R4)"),
                ("S", "T3", "=R2"),
                ("S", "T4", "=K2"),
                ("S", "T5", "=SUM(K1#)"),
            ],
            &[
                ("S", "K1", "=SEQUENCE(4)", None),
                ("S", "J3", "=SEQUENCE(1,3)", None),
                ("S", "P1", "=FILTER(D!A1:A6,D!A1:A6>4)", None),
                ("S", "P3", "=SEQUENCE(2)", None),
                ("S", "R1", "=FILTER(D!A1:A6,D!A1:A6>2)", None),
                ("S", "R3", "=SEQUENCE(2)", None),
            ],
            &[],
        );
        let lib = library::language();
        let calc = Calc::run(&wb, &lib);
        let meets = Computed::NotComputed(Reason::Construct(spill::MEETS.to_string()));
        for addr in ["K1", "J3", "A31", "T4", "T5"] {
            assert_eq!(at(&calc, "S", addr), meets, "S!{addr}");
        }
        for (addr, want) in [("P1", 5.0), ("P2", 9.0), ("P3", 1.0), ("P4", 2.0)] {
            assert_eq!(number(&calc, "S", addr), want, "S!{addr}");
        }
        assert_eq!(
            at(&calc, "S", "R1"),
            Computed::Value(Value::Error(ErrorKind::Spill))
        );
        for (addr, want) in [("R3", 1.0), ("R4", 2.0), ("T1", 17.0), ("T2", 3.0)] {
            assert_eq!(number(&calc, "S", addr), want, "S!{addr}");
        }
        assert_eq!(number(&calc, "S", "T3"), 0.0);
    }

    #[test]
    fn a_spill_no_plan_bounds_is_labelled_and_a_saved_extent_plans_one() {
        // A count read from a cell gives no bound: one value computes, many
        // are labelled; the extent its file saved plans it, and cells it saved
        // that the answer no longer reaches hold nothing, never their last
        // values; a spill reading its own extent is a cycle; and a recalculation
        // as an engine's machine reads one labels each spill as before.
        let wb = spilling(
            &[
                ("D", "K3", "3"),
                ("D", "K4", "1"),
                ("S", "C2", "0"),
                ("S", "C3", "0"),
                ("S", "E2", "7"),
                ("S", "E3", "7"),
                ("S", "E4", "7"),
                ("S", "E5", "7"),
                ("S", "H1", "=SUM(E4:E5)"),
                ("S", "N2", "=1"),
            ],
            &[
                ("S", "A1", "=SEQUENCE(D!K3)", None),
                ("S", "B1", "=SEQUENCE(D!K4)", None),
                ("S", "C1", "=SEQUENCE(D!K3)", Some("C1:C3")),
                ("S", "E1", "=SEQUENCE(3)", Some("E1:E5")),
                ("S", "G1", "=SUM(G2:G3)+SEQUENCE(3)", None),
                ("S", "J1", "=SEQUENCE(D!K3)", Some("J1:J2")),
                ("S", "L1", "=M1#", None),
                ("S", "M1", "=L1#", None),
                ("S", "N1", "=SEQUENCE(3)", Some("N1:N3")),
            ],
            &[],
        );
        let lib = library::language();
        let calc = Calc::run(&wb, &lib);
        let s = wb.find_sheet("S").unwrap();
        assert_eq!(
            at(&calc, "S", "A1"),
            Computed::NotComputed(Reason::Construct(spill::NO_PLAN.to_string()))
        );
        assert_eq!(number(&calc, "S", "B1"), 1.0);
        for (addr, want) in [("C1", 1.0), ("C2", 2.0), ("C3", 3.0), ("E3", 3.0)] {
            assert_eq!(number(&calc, "S", addr), want, "S!{addr}");
        }
        assert_eq!(at(&calc, "S", "E4"), Computed::Value(Value::Empty));
        assert_eq!(at(&calc, "S", "E5"), Computed::Value(Value::Empty));
        assert!(
            calc.cells().contains(&(s, 5, 5)),
            "a saved cell is computed"
        );
        assert_eq!(number(&calc, "S", "H1"), 0.0);
        // Planned by its file and answering more, labelled; a formula where
        // its file saved its spill blocks it.
        assert_eq!(
            at(&calc, "S", "J1"),
            Computed::NotComputed(Reason::Construct(spill::NO_PLAN.to_string()))
        );
        assert_eq!(
            at(&calc, "S", "N1"),
            Computed::Value(Value::Error(ErrorKind::Spill))
        );
        // A spill reading its own extent, and two reading each other's, cycles.
        assert_eq!(
            calc.cycles(),
            [vec![(s, 1, 7)], vec![(s, 1, 12), (s, 1, 13)]]
        );
        // Labelled as an engine's machine labels it until the item's fifth
        // slice: no spill planned, `A1#` labelled.
        let wb = the_spills();
        let calc = Calc::unspilled(&wb, &lib);
        let spill = Computed::NotComputed(Reason::Construct(eval::SPILL.to_string()));
        for addr in ["A1", "A15", "N1", "A26"] {
            assert_eq!(at(&calc, "S", addr), spill, "S!{addr}");
        }
        assert_eq!(calc.cells().len(), calc.formulas());
    }

    #[test]
    fn the_record_prints_a_spill_s_cells_as_values_and_its_extent_after_its_anchor() {
        // KERNEL.55's third slice, decision 2: each record says a formula's
        // kind; after a spill's anchor's `value` row its `spill` row, each of
        // its other cells a `value` row with nothing before it, the sheet
        // holding it or not; a blocked anchor's `#SPILL!` and no `spill` row; a
        // block the same, its anchor's word `array`; a cell its file saved its
        // spill in that the answer no longer reaches, no row; and the values,
        // the grid without its `formula` rows.
        let mut wb = spilling(
            &[
                ("S", "D3", "x"),
                ("S", "G2", "7"),
                ("S", "G3", "7"),
                ("S", "H1", "=SUM(A1#)"),
            ],
            &[
                ("S", "A1", "=SEQUENCE(2,2)", None),
                ("S", "D1", "=SEQUENCE(4)", None),
                ("S", "G1", "=SEQUENCE(2)", Some("G1:G3")),
            ],
            &[],
        );
        let s = wb.find_sheet("S").unwrap();
        let range = parse_a1_range("E1:E2").unwrap();
        wb.sheets[s].set(
            1,
            5,
            Cell {
                content: Content::ArrayBlock {
                    text: "SEQUENCE(2)*10".to_string(),
                    range,
                },
                style: 0,
            },
        );
        // A block's other cell, as a file holds one.
        wb.sheets[s].set(
            2,
            5,
            Cell {
                content: Content::Number(0.0),
                style: 0,
            },
        );
        let lib = library::language();
        let calc = Calc::run(&wb, &lib);
        let window = crate::view::window_of(&wb, "S", None).unwrap();
        let grid = crate::view::view_text_valued(&wb, &window, &calc);
        let rows: Vec<&str> = grid
            .lines()
            .filter(|l| {
                ["(formula ", "(value ", "(spill ", "(cell "]
                    .iter()
                    .any(|k| l.starts_with(k))
            })
            .collect();
        assert_eq!(
            rows,
            [
                "(formula \"S\" \"A1\" \"=SEQUENCE(2,2)\" dynamic)",
                "(value \"S\" \"A1\" 1)",
                "(spill \"S\" \"A1\" \"A1:B2\")",
                "(value \"S\" \"B1\" 2)",
                "(formula \"S\" \"D1\" \"=SEQUENCE(4)\" dynamic)",
                "(value \"S\" \"D1\" (error \"#SPILL!\"))",
                "(formula \"S\" \"E1\" \"=SEQUENCE(2)*10\" array)",
                "(value \"S\" \"E1\" 10)",
                "(spill \"S\" \"E1\" \"E1:E2\")",
                "(formula \"S\" \"G1\" \"=SEQUENCE(2)\" dynamic)",
                "(value \"S\" \"G1\" 1)",
                "(spill \"S\" \"G1\" \"G1:G2\")",
                "(formula \"S\" \"H1\" \"=SUM(A1#)\")",
                "(value \"S\" \"H1\" 10)",
                "(value \"S\" \"A2\" 3)",
                "(value \"S\" \"B2\" 4)",
                "(value \"S\" \"E2\" 20)",
                "(value \"S\" \"G2\" 2)",
                "(cell \"S\" \"D3\" \"x\")",
            ],
            "{grid}"
        );
        let values = crate::view::view_text_values(&wb, &window, &calc);
        let less: String = grid
            .lines()
            .filter(|l| !l.starts_with("(formula "))
            .map(|l| format!("{l}\n"))
            .collect();
        assert_eq!(values, less);
        // A record with no values prints the formulas alone, their words said.
        let bare = crate::view::view_text(&wb, &window);
        assert!(bare.contains("(formula \"S\" \"E1\" \"=SEQUENCE(2)*10\" array)"));
        assert!(!bare.contains("(spill ") && !bare.contains("(value "));
    }

    #[test]
    fn a_machine_loads_a_door_s_record_of_spills_and_sets_their_cells_aside() {
        // Until the item's fifth slice an engine computes no spill: its load
        // takes a door's record whole, the words, each `spill` row naming an
        // anchor of a dynamic or a legacy array formula, and the `value` and
        // `style` rows of the cells it fills, which it sets aside; a `spill`
        // row anywhere else, and a `value` row at a cell no formula and no
        // spill fills, are refused as a malformed row.
        let mut wb = spilling(
            &[("S", "B2", "4")],
            &[("S", "A1", "=SEQUENCE(2,2)", Some("A1:B2"))],
            &[],
        );
        // A spilled cell the file holds with a format of its own.
        let s = wb.find_sheet("S").unwrap();
        wb.sheets[s].cells.get_mut(&(2, 2)).unwrap().style = 1;
        let lib = library::language();
        let calc = Calc::run(&wb, &lib);
        let window = crate::view::window_of(&wb, "S", None).unwrap();
        let record = crate::view::view_text_valued(&wb, &window, &calc);
        assert!(record.contains("(style \"S\" \"B2\" 1)"), "{record}");
        let mut m = crate::machine::Machine::load(&record, library::language())
            .unwrap_or_else(|e| panic!("line {}: {}\n{record}", e.line, e.refusal.text));
        assert!(m.step(0).done);
        let text =
            format!("{record}(formula \"S\" \"D1\" \"=1+1\" array)\n(value \"S\" \"D1\" 2)\n");
        assert!(crate::machine::Machine::load(&text, library::language()).is_ok());
        for wrong in [
            "(spill \"S\" \"C1\" \"C1:C2\")",
            "(spill \"S\" \"A1\" \"B1:B2\")",
            "(spill \"S\" \"A1\" \"A1\")",
            "(value \"S\" \"C3\" 5)",
            "(formula \"S\" \"D1:D2\" \"=1+1\" array)",
            // The anchor's value said again: its first taken as its formula's
            // last value, never set aside as a spilled cell's.
            "(value \"S\" \"A1\" 5)",
        ] {
            let text = format!("{record}{wrong}\n");
            let e = crate::machine::Machine::load(&text, library::language()).expect_err(wrong);
            assert!(
                e.refusal.text.contains("row"),
                "{wrong}: {}",
                e.refusal.text
            );
        }
    }
}
