//! The plan of shapes (KERNEL.22, 2026-10-09): every formula of a grid
//! parsed once per shape, each formula cell evaluated at its offset from
//! the shape's first cell, and the order of evaluation read off the parse.
//!
//! An engine evaluates every formula every frame, so a parse at each
//! evaluation, as [`super::Calc`] makes, costs most of a frame (Life's
//! 62,964 formulas: about two seconds of a frame's 2.5, measured 2026-10-09),
//! and a parse of every cell held once costs memory no browser holds (8.9 KB
//! for one of Life's formulas, 562 MB for all of them). A shape is the
//! answer: a shared formula's text is read once at its master and its other
//! cells are the same formula moved, and two plain formula cells whose R1C1
//! text is the same (`refers::r1c1`, the reference's own normal form) are
//! the same formula moved too, so a save of Life, 62,964 plain formula rows,
//! loads back to one shape. A cell at an offset from its shape's first cell
//! is evaluated with every reference moved by that offset, which is what the
//! text `refers::shift_a1_references` moves would parse to
//! ([`super::formula::RefExpr::moved`]); the references inside a defined
//! name's text never move.
//!
//! [`super::Calc`] keeps its plan of texts: it evaluates each formula once,
//! so a parse held would buy it nothing. The plan here is an engine's, kept
//! for the life of a grid and read again when a write changes a formula; a
//! shape's references find the sheets they name once, as it is placed,
//! among the grid's and its twins, which are fixed for the grid's life
//! (KERNEL.20), where a reference found its sheet by name at every
//! evaluation.
//!
//! The order is by level, a cell's being one past the deepest it reads, then
//! by shape and cell, which computes the frame any order after what each cell
//! reads computes; its runs are the cells of a shape that stand together in
//! it, on one sheet, reading none of one another, and an engine evaluates
//! each run at once, node by node (`super::vector`, KERNEL.20), and a cell
//! alone, or a run's few, by its shape's parse lowered once to a program
//! (`super::program`, KERNEL.44). What each
//! shape reads is found with the order, so that an engine follows a changed
//! cell to the cells of the order that read it and computes those alone
//! (`super::reach`, KERNEL.20); and so are the sheets a read takes as an
//! array, a range of them or a run's moving cell, which an engine's machine
//! keeps a frame's values of as arrays (KERNEL.31, [`Shapes::read_as_array`]).
//!
//! The plan is printed as rows on request ([`Shapes::census`], KERNEL.30,
//! `super::census`): each shape's cells, the levels of the order it spans,
//! its runs and the paths its reads take, so that a grid's cost is read
//! before a frame runs. Of all that, the plan keeps for it only where each
//! level of the order ends, the array its counting sort already makes.

use std::collections::HashMap;

use crate::fx::FxHashMap;
use crate::messages::{raise, Refusal};
use crate::refers::{self, quote_sheet};
use crate::sheet::{cell_ref, shared_masters, strip_future_prefixes, Content, Workbook};

use super::eval::{unread_name, Computed, Ctx, Env, Library, Reason};
use super::formula::{self, Expr, Read, RefExpr};
use super::graph::{
    self, areas_at, cells_in, dependencies, is_formula, Area, CellId, Extent, Names,
};
use super::program::{Program, Why};
use super::reach::{Positions, Reach};
use super::vector::{self, Lanes, Repeats};

/// One formula, parsed once, and the cell it was read at.
#[derive(Clone, Debug)]
pub struct Shape {
    /// The text at the shape's first cell, without its `=`, as the formula
    /// bar shows it.
    pub text: String,
    /// The parse, or the construct it stopped at.
    pub expr: Result<Expr, String>,
    /// The cell the text was read at; every other cell of the shape is at
    /// an offset from it.
    pub anchor: CellId,
    /// The functions the parse calls that the library does not hold,
    /// sorted, each once: what a load refuses, read once a shape.
    missing: Vec<String>,
    /// Whether a cell of the shape may stand for a text the parse does not
    /// read: the text holds a 3D span (`Jan:Dec!A1`) or a broken reference
    /// with a cell after it (`#REF!A1`), each of which the mover writes as
    /// a construct when it moves off the sheet (`Jan:Dec!#REF!`,
    /// `#REF!#REF!`). Found 2026-10-09 by holding the offset to the moved
    /// text's parse over the refers fixture.
    pub(crate) risky: bool,
    /// The most buffers an evaluation of the parse at many cells holds at
    /// once, found once (KERNEL.20, `vector::holds`), a lane for each of its
    /// repeated parts with them (KERNEL.27).
    holds: (usize, usize),
    /// The parts the parse writes more than once that a run computes once
    /// (KERNEL.27, `vector::Repeats`).
    repeats: Repeats,
    /// The index of the shared formula the shape is, on its first cell's
    /// sheet, whose master is its first cell; `None` for the shape of plain
    /// formulas of one R1C1 text (KERNEL.34).
    fill: Option<u32>,
    /// The parse lowered once, for a cell evaluated alone and a run's few
    /// cells (KERNEL.44, `super::program`), or why it is left to `Ctx`;
    /// `None` for a text not read and a risky shape, whose cells `Ctx`
    /// evaluates as it always has.
    lowered: Option<Result<Program, Why>>,
    /// Whether the parse holds an array anywhere (KERNEL.8, slice f,
    /// `vector::arrays_in`): its cells each computed by `Ctx` whole, a run's
    /// among them, since a lane holds one value a cell; and a shape of dynamic
    /// formulas or legacy array formulas, read as arrays.
    arrayed: bool,
}

impl Shape {
    /// The most buffers of each kind a run of the shape holds at once.
    #[cfg(test)]
    pub(crate) fn holds(&self) -> (usize, usize) {
        self.holds
    }

    /// The functions the parse calls that the library does not hold, sorted,
    /// each once (KERNEL.30's census prints them).
    pub(crate) fn missing(&self) -> &[String] {
        &self.missing
    }

    /// Why the parse is left to `Ctx` (KERNEL.44's census prints it); `None`
    /// for a parse lowered, a text not read and a risky shape.
    pub(crate) fn unlowered(&self) -> Option<Why> {
        match &self.lowered {
            Some(Err(why)) => Some(*why),
            _ => None,
        }
    }

    /// The construct a cell at this offset stands for that the shape's
    /// first cell does not, as the moved text's parse names it; `None` when
    /// the moved text reads as the shape does, which for a shape that is not
    /// risky it always does (the refers fixture moved to every edge of the
    /// sheet holds that).
    pub fn unread_at(&self, d_row: i64, d_col: i64) -> Option<String> {
        if !self.risky || (d_row == 0 && d_col == 0) || self.expr.is_err() {
            return None;
        }
        formula::parse(&refers::shift_a1_references(&self.text, d_row, d_col)).err()
    }
}

/// Where a formula cell's formula comes from: its shape, and the cell's
/// offset from the shape's first cell.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Place {
    pub shape: usize,
    pub d_row: i64,
    pub d_col: i64,
}

/// The plan: every formula cell placed on a shape, the order, the cycles.
#[derive(Clone, Debug, Default)]
pub struct Shapes {
    shapes: Vec<Shape>,
    /// A plain formula's shape, by its R1C1 text.
    by_r1c1: HashMap<String, usize>,
    /// A shared formula's shape, by its sheet and its index.
    by_shared: HashMap<(usize, u32), usize>,
    /// Every formula cell's place, looked up at each of its evaluations,
    /// under the crate's own hash (KERNEL.25).
    places: FxHashMap<CellId, Place>,
    /// A cell of a risky shape whose own text the parse does not read: the
    /// construct it names, found once when the cell is placed.
    unread: FxHashMap<CellId, String>,
    order: Vec<CellId>,
    cycles: Vec<Vec<CellId>>,
    /// The runs of the order an engine evaluates at once (KERNEL.20), by
    /// their first position: each of at least [`RUN`] cells of one parsed
    /// shape that is not risky, on one sheet, consecutive in the order, none
    /// reading another.
    runs: Vec<Run>,
    /// What each shape reads, found with the order (KERNEL.20).
    reach: Reach,
    /// Where each level of the order ends, by level (KERNEL.30): the
    /// position after its last cell, the order's own counting sort kept
    /// rather than dropped, so that the census finds each cell's level
    /// without planning again.
    level_ends: Vec<u32>,
    /// By sheet of the grid, whether a read takes it as an array (KERNEL.31,
    /// [`Shapes::read_as_array`]), found with the order.
    arrays: Vec<bool>,
}

/// A run of the order: its cells are `order[start..end]`, of one shape.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct Run {
    start: usize,
    end: usize,
    shape: usize,
}

/// The fewest cells a run is evaluated at once for; fewer are evaluated one
/// by one, as every cell was before (KERNEL.20).
pub(crate) const RUN: usize = 4;

/// The fewest cells of a run evaluated at once when its shape has a program
/// (KERNEL.44): fewer are evaluated one by one by the program, which costs
/// them less than a run's start costs them. Measured over 4 to 64: Kairos's
/// runs of about 14 cells a frame selects cost 143 ns a cell against a
/// program's 40, and Life's rule breaks even near 20 (its program 108 ns a
/// cell, a run's 32 and a start of about 1.5 µs), 64 costing the gun.
pub(crate) const SHORT: usize = 16;

impl Shapes {
    /// A plan with nothing in it.
    pub fn new() -> Shapes {
        Shapes::default()
    }

    /// Every formula cell of the sheets `0..sheets` placed, the shapes
    /// parsed, and the order found. The plan is made in a box: every formula
    /// is parsed beneath this frame, the deepest a load reaches, which
    /// `VLA.1`'s limits pin on a stack of 1 MiB, and the plan held here by
    /// value would stand on it, 480 bytes since it kept its level ends
    /// (KERNEL.30).
    pub fn of(
        model: &Workbook,
        sheets: usize,
        library: &Library,
        names: &Names,
        extents: &[Extent],
    ) -> Shapes {
        let mut s: Box<Shapes> = Box::default();
        for (si, sheet) in model.sheets.iter().enumerate().take(sheets) {
            let masters = shared_masters(sheet);
            for (&(row, col), cell) in &sheet.cells {
                s.place_with(model, (si, row, col), &cell.content, &masters, library);
            }
        }
        s.plan(model, names, extents);
        *s
    }

    /// Place one cell after a write: its old place forgotten, and a new one
    /// found when it holds a formula. The order waits for [`Shapes::plan`].
    pub fn replace(&mut self, model: &Workbook, id: CellId, library: &Library) {
        self.places.remove(&id);
        self.unread.remove(&id);
        let Some(sheet) = model.sheets.get(id.0) else {
            return;
        };
        let Some(cell) = sheet.cells.get(&(id.1, id.2)) else {
            return;
        };
        if !is_formula(&cell.content) {
            return;
        }
        let masters = shared_masters(sheet);
        self.place_with(model, id, &cell.content, &masters, library);
    }

    /// Many cells placed after a write, as [`Shapes::replace`] places one,
    /// each sheet's shared formulas found once for all its cells, where
    /// `replace` walks the sheet for each (KERNEL.34: Life's rule written
    /// whole walked its sheet 62,964 times, 3.96 billion cells visited, in
    /// 13.3 s); then the plan made the one a load of the same grid makes
    /// ([`Shapes::as_loaded`]). The order waits for [`Shapes::plan`].
    pub(crate) fn replace_many(&mut self, model: &Workbook, ids: &[CellId], library: &Library) {
        for id in ids {
            self.places.remove(id);
            self.unread.remove(id);
        }
        let mut sheets: Vec<usize> = ids.iter().map(|id| id.0).collect();
        sheets.sort_unstable();
        sheets.dedup();
        for s in sheets {
            let Some(sheet) = model.sheets.get(s) else {
                continue;
            };
            let masters = shared_masters(sheet);
            for &id in ids.iter().filter(|id| id.0 == s) {
                let Some(cell) = sheet.cells.get(&(id.1, id.2)) else {
                    continue;
                };
                if is_formula(&cell.content) {
                    self.place_with(model, id, &cell.content, &masters, library);
                }
            }
        }
        self.as_loaded(model, library);
    }

    /// The plan made the one [`Shapes::of`] makes over the same grid, once a
    /// write has placed some of its cells again (KERNEL.34): the shapes in
    /// the order of their first cells in the grid's order, sheet, row and
    /// column; a fill's shape anchored at its master and a plain formula's
    /// at its first cell, read again from that cell's text where the write
    /// moved or respelled it; a shape no cell holds dropped. Every cell's
    /// place and every construct a cell leaves unread are then a load's, and
    /// so, once planned, are the order, its runs and the reach, and the
    /// census prints a load's. A write left a fill's other cells on the
    /// shape of its first cell even when that cell had been written over,
    /// and kept a shape no cell held, so that the census after a write was
    /// numbered and anchored otherwise than a load's.
    fn as_loaded(&mut self, model: &Workbook, library: &Library) {
        // Each shape's first cell, the least of its places, found in one pass:
        // sorting the places to find them took a write into Sand 14 ms.
        let mut first: Vec<Option<CellId>> = vec![None; self.shapes.len()];
        for (&id, p) in &self.places {
            match first[p.shape] {
                Some(f) if f <= id => {}
                _ => first[p.shape] = Some(id),
            }
        }
        let mut order: Vec<usize> = (0..self.shapes.len())
            .filter(|&s| first[s].is_some())
            .collect();
        order.sort_unstable_by_key(|&s| first[s]);
        // Each shape's first cell as a load anchors it, with its text there.
        let mut masters: HashMap<usize, HashMap<u32, (&str, u32, u32)>> = HashMap::new();
        let mut again: Vec<(usize, CellId, String)> = Vec::new();
        for &s in &order {
            let f = first[s].expect("a shape in the order has a first cell");
            let shape = &self.shapes[s];
            let (anchor, raw) = match shape.fill {
                Some(index) => {
                    let sheet = shape.anchor.0;
                    let found = masters
                        .entry(sheet)
                        .or_insert_with(|| shared_masters(&model.sheets[sheet]))
                        .get(&index);
                    match found {
                        Some(&(text, row, col)) => ((sheet, row, col), text),
                        None => continue,
                    }
                }
                None => match model.sheets[f.0].cells.get(&(f.1, f.2)).map(|c| &c.content) {
                    Some(
                        Content::Formula(t) | Content::DynamicFormula(t) | Content::ArrayFormula(t),
                    ) => (f, t.as_str()),
                    _ => continue,
                },
            };
            if anchor != shape.anchor
                || (raw != shape.text && strip_future_prefixes(raw) != shape.text)
            {
                again.push((s, anchor, strip_future_prefixes(raw)));
            }
        }
        let numbered = order.iter().enumerate().all(|(k, &s)| k == s);
        if again.is_empty() && numbered && order.len() == self.shapes.len() {
            return;
        }
        // The shapes read again from their first cells, then every shape in
        // the order of its first cell, each place renumbered and, for a shape
        // read again, moved from its new first cell.
        let mut renewed: HashMap<usize, Shape> = HashMap::new();
        for (s, anchor, text) in again {
            let fill = self.shapes[s].fill;
            self.add(model, text, anchor, library);
            let mut shape = self.shapes.pop().expect("the shape just made");
            shape.fill = fill;
            renewed.insert(s, shape);
        }
        let mut number: Vec<Option<usize>> = vec![None; self.shapes.len()];
        let mut moved: Vec<bool> = Vec::with_capacity(order.len());
        let mut old: Vec<Option<Shape>> = std::mem::take(&mut self.shapes)
            .into_iter()
            .map(Some)
            .collect();
        for (k, &s) in order.iter().enumerate() {
            number[s] = Some(k);
            moved.push(renewed.contains_key(&s));
            let shape = renewed
                .remove(&s)
                .or_else(|| old[s].take())
                .expect("each shape once");
            self.shapes.push(shape);
        }
        let shapes = &self.shapes;
        for (id, p) in self.places.iter_mut() {
            p.shape = number[p.shape].expect("a placed cell's shape is kept");
            if !moved[p.shape] {
                continue;
            }
            let shape = &shapes[p.shape];
            p.d_row = i64::from(id.1) - i64::from(shape.anchor.1);
            p.d_col = i64::from(id.2) - i64::from(shape.anchor.2);
            match shape.unread_at(p.d_row, p.d_col) {
                Some(construct) => self.unread.insert(*id, construct),
                None => self.unread.remove(id),
            };
        }
        self.by_shared.retain(|_, s| number[*s].is_some());
        self.by_r1c1.retain(|_, s| number[*s].is_some());
        for s in self.by_shared.values_mut().chain(self.by_r1c1.values_mut()) {
            *s = number[*s].expect("kept");
        }
    }

    fn place_with(
        &mut self,
        model: &Workbook,
        id: CellId,
        content: &Content,
        masters: &HashMap<u32, (&str, u32, u32)>,
        library: &Library,
    ) {
        let (si, row, col) = id;
        let (shape, anchor) = match content {
            Content::SharedMaster { si: index, .. } | Content::SharedChild { si: index } => {
                let Some(&(text, m_row, m_col)) = masters.get(index) else {
                    // A child whose master is gone has no text; the machine
                    // never leaves one (a write over a master makes the
                    // fill's first cell left its master, or gives each of its
                    // cells its own text, first).
                    return;
                };
                let anchor = (si, m_row, m_col);
                let key = (si, *index);
                let shape = match self.by_shared.get(&key) {
                    Some(&s) => s,
                    None => {
                        let s = self.add(model, strip_future_prefixes(text), anchor, library);
                        self.shapes[s].fill = Some(*index);
                        self.by_shared.insert(key, s);
                        s
                    }
                };
                (shape, anchor)
            }
            Content::Formula(t) | Content::DynamicFormula(t) | Content::ArrayFormula(t) => {
                let text = strip_future_prefixes(t);
                let mut key = refers::r1c1(&text, row, col);
                // A dynamic formula and a legacy array formula are read as
                // arrays, a shape of their own beside a legacy formula of the
                // same text (KERNEL.8, slice f).
                match content {
                    Content::DynamicFormula(_) => key.insert(0, '\u{1}'),
                    Content::ArrayFormula(_) => key.insert(0, '\u{2}'),
                    _ => {}
                }
                match self.by_r1c1.get(&key) {
                    Some(&s) => (s, self.shapes[s].anchor),
                    None => {
                        let s = self.add(model, text, id, library);
                        self.by_r1c1.insert(key, s);
                        (s, id)
                    }
                }
            }
            _ => return,
        };
        let (d_row, d_col) = (
            i64::from(row) - i64::from(anchor.1),
            i64::from(col) - i64::from(anchor.2),
        );
        self.places.insert(
            id,
            Place {
                shape,
                d_row,
                d_col,
            },
        );
        match self.shapes[shape].unread_at(d_row, d_col) {
            Some(construct) => self.unread.insert(id, construct),
            None => self.unread.remove(&id),
        };
    }

    fn add(&mut self, model: &Workbook, text: String, anchor: CellId, library: &Library) -> usize {
        let mut expr = formula::parse(&text);
        if let Ok(e) = expr.as_mut() {
            find_sheets(e, model);
        }
        // The calls as the text's tokens hold them, parsed or not, as
        // `Calc::missing_functions` reads them.
        let mut missing: Vec<String> = formula::calls(&text)
            .into_iter()
            .filter(|name| !library.has(name))
            .collect();
        let risky = refers::scan(&text).iter().any(|r| {
            r.kind == refers::Kind::ThreeD
                || (r.kind == refers::Kind::Broken && r.shape != refers::Shape::None)
        });
        missing.sort_unstable();
        missing.dedup();
        // A shape of dynamic formulas or legacy array formulas, read as
        // arrays, is computed at each cell by `Ctx` (KERNEL.8, slice f).
        let kind_reads_arrays =
            super::is_dynamic(model, anchor) || super::is_array_formula(model, anchor);
        let s = self.push(text, expr, anchor, missing, risky, library);
        if kind_reads_arrays {
            // A program lowered for such a shape would read its ranges as a
            // legacy formula does; one already left to `Ctx` keeps its reason.
            let shape = &mut self.shapes[s];
            shape.arrayed = true;
            if let Some(Ok(_)) = shape.lowered {
                shape.lowered = Some(Err(Why::Array));
            }
        }
        s
    }

    /// A shape kept, with what an evaluation of it at many cells holds found
    /// once (KERNEL.20, `vector::holds`), the parts it writes more than once
    /// (KERNEL.27, `vector::Repeats`) and its parse lowered (KERNEL.44,
    /// `super::program`; a risky shape keeps none, its cells read as it does
    /// only in a grid that loads), in a frame of its own: `add`'s frame
    /// stands beneath the parse, the deepest a load reaches, which `VLA.1`'s
    /// limits pin on a stack of 1 MiB.
    #[inline(never)]
    fn push(
        &mut self,
        text: String,
        expr: Result<Expr, String>,
        anchor: CellId,
        missing: Vec<String>,
        risky: bool,
        library: &Library,
    ) -> usize {
        let (holds, repeats) = match &expr {
            Ok(e) => {
                let repeats = Repeats::of(library, e);
                let (values, positions) = vector::holds(library, e);
                ((values + repeats.parts(), positions), repeats)
            }
            Err(_) => ((0, 0), Repeats::default()),
        };
        let arrayed = expr.as_ref().is_ok_and(|e| vector::arrays_in(library, e));
        let lowered = match &expr {
            Ok(e) if !risky => Some(Program::lower(library, e)),
            _ => None,
        };
        self.shapes.push(Shape {
            text,
            expr,
            anchor,
            missing,
            risky,
            holds,
            repeats,
            fill: None,
            lowered,
            arrayed,
        });
        self.shapes.len() - 1
    }

    /// The order found again from the places: every formula cell's edges
    /// read off its shape's parse at its offset (a defined name's text read
    /// as `graph::dependencies` reads it), then Tarjan's components, as
    /// `graph::plan` finds them for a plan of texts.
    ///
    /// Each shape's references are read off its parse once (KERNEL.33,
    /// asked by Alonzo's playtest pass of Sand, whose load spent 341 ms of
    /// 491 finding 2.37 million edges a cell at a time): each reference once,
    /// one to a sheet holding no formula left out, one every `$` holds found
    /// once for the whole shape; each cell then moves only the rest. A cell's
    /// vertex is found as reach finds a position, by an array over a sheet
    /// whose formulas stand close together, and the edges are one flat array.
    /// Every cell's edges are the ones a walk of its parse at its offset
    /// finds, sorted and each once, so the order, its runs and its cycles
    /// are the same. A shape whose held references name many formula cells
    /// has them as one vertex of its own, a hub after the cells, each of its
    /// cells one edge to it (KERNEL.8): twenty thousand cells searching one
    /// table of twenty thousand formulas held whole made four hundred million
    /// edges, the same cells in every one. A hub stands for its cells and no
    /// more, which [`Shapes::order_and_runs`] reads through.
    pub fn plan(&mut self, model: &Workbook, names: &Names, extents: &[Extent]) {
        let mut ids: Vec<CellId> = self.places.keys().copied().collect();
        ids.sort_unstable();
        let index = Positions::of(&ids, model.sheets.len());
        let mut holds = vec![false; model.sheets.len()];
        for id in &ids {
            holds[id.0] = true;
        }
        let vertices = Vertices {
            model,
            extents,
            index: &index,
            holds: &holds,
        };
        let reads: Vec<ShapeReads> = self
            .shapes
            .iter()
            .map(|shape| vertices.reads_of(shape))
            .collect();
        // Every cell's edges in one array, cell `v`'s at
        // `targets[offsets[v]..offsets[v + 1]]`, where a vector of edges was
        // made for each cell.
        let n = ids.len();
        let mut offsets: Vec<usize> = Vec::with_capacity(n + 1);
        let mut targets: Vec<u32> = Vec::with_capacity(n * 8);
        offsets.push(0);
        let mut shape_of: Vec<usize> = Vec::with_capacity(n);
        let mut deps: Vec<CellId> = Vec::new();
        let mut edges: Vec<u32> = Vec::new();
        // The vertices each shape's held references name, by the shape and,
        // where one of them names the formula's own sheet, the cell's sheet,
        // found once (KERNEL.8: a range of the own sheet was found again at
        // every cell); with the hub of each that names enough of them,
        // numbered after the cells.
        let mut held: FxHashMap<(usize, usize), (Vec<u32>, Option<u32>)> = FxHashMap::default();
        let mut hubs: Vec<(usize, usize)> = Vec::new();
        for id in &ids {
            edges.clear();
            deps.clear();
            let p = self.places[id];
            shape_of.push(p.shape);
            let shape = &self.shapes[p.shape];
            let read = !self.unread.contains_key(id);
            match &shape.expr {
                Ok(e) if read => {
                    let r = &reads[p.shape];
                    // A shape holding nothing of its own sheet and too few
                    // cells for a hub reads its held cells as the shape found
                    // them, no entry made for it.
                    if r.own.is_empty() && r.fixed.len() < hub_at() {
                        edges.extend_from_slice(&r.fixed);
                    } else {
                        let key = (p.shape, if r.own.is_empty() { usize::MAX } else { id.0 });
                        let (fixed, hub) = held.entry(key).or_insert_with(|| {
                            let mut fixed = r.fixed.clone();
                            for re in &r.own {
                                if let Ok(areas) = areas_at(model, id.0, re, 0, 0) {
                                    for a in &areas {
                                        vertices.formulas_in(a, &mut fixed);
                                    }
                                }
                            }
                            fixed.sort_unstable();
                            fixed.dedup();
                            let hub = (fixed.len() >= hub_at()).then(|| {
                                hubs.push(key);
                                vertex(n + hubs.len() - 1)
                            });
                            (fixed, hub)
                        });
                        match hub {
                            Some(hub) => edges.push(*hub),
                            None => edges.extend_from_slice(fixed),
                        }
                    }
                    for re in &r.moving {
                        let Ok(areas) = areas_at(model, id.0, re, p.d_row, p.d_col) else {
                            continue;
                        };
                        for a in &areas {
                            vertices.formulas_in(a, &mut edges);
                        }
                    }
                    if r.named {
                        e.each_read(&mut |r| {
                            if let Read::Name { name, sheet } = r {
                                let scope = sheet.and_then(|s| model.find_sheet(s));
                                if let Some(t) = names.resolve(name, scope, id.0) {
                                    dependencies(
                                        model,
                                        names,
                                        extents,
                                        scope.unwrap_or(id.0),
                                        t,
                                        &mut deps,
                                    );
                                }
                            }
                        });
                    }
                }
                // A text this version does not parse, here or at this
                // offset: its edges as the text scan reads them, so that a
                // cycle through it is still found.
                _ => {
                    let moved = refers::shift_a1_references(&shape.text, p.d_row, p.d_col);
                    dependencies(model, names, extents, id.0, &moved, &mut deps);
                }
            }
            for d in &deps {
                if let Some(v) = index.get(d.0, d.1, d.2) {
                    edges.push(vertex(v));
                }
            }
            edges.sort_unstable();
            edges.dedup();
            targets.extend_from_slice(&edges);
            offsets.push(targets.len());
        }
        for key in &hubs {
            targets.extend_from_slice(&held[key].0);
            offsets.push(targets.len());
        }
        #[cfg(test)]
        HUBS_MADE.with(|h| h.set(h.get() + hubs.len()));
        let (members, starts) = graph::tarjan_flat(&offsets, &targets);
        let shape_at =
            self.order_and_runs(&ids, (&offsets, &targets), &shape_of, (&members, &starts));
        self.arrays = read_as_arrays(model, &self.shapes, (&ids, &shape_of), &self.runs);
        let places = &self.places;
        let formula = |id: CellId| places.contains_key(&id);
        self.reach = Reach::of(model, &self.shapes, &self.order, shape_at, &formula);
    }

    /// The order, the cycles and the runs, from the cells, each one's edges
    /// (one flat array, as [`graph::tarjan_flat`] reads them) and shape, and
    /// the components, which come each after everything it reads; the shape
    /// of each cell of the order answered.
    ///
    /// Each cell's level is 0 when it reads no formula of the plan, else one
    /// past the deepest it reads, its component's. The order is by level,
    /// then shape, then cell, sheet first: any order that puts a cell after
    /// what it reads computes the same frame, and in this one the cells of a
    /// shape that read none of one another stand together, whether the shape
    /// fills rows or a column beside others (KERNEL.20).
    fn order_and_runs(
        &mut self,
        ids: &[CellId],
        (offsets, targets): (&[usize], &[u32]),
        shape_of: &[usize],
        (members, starts): (&[u32], &[usize]),
    ) -> Vec<u32> {
        self.order.clear();
        self.cycles.clear();
        self.runs.clear();
        let adj = |v: usize| &targets[offsets[v]..offsets[v + 1]];
        // The cells, then the hubs (`Shapes::plan`): a hub's level is its
        // deepest cell's, so that a cell reading through it stands one past
        // that cell, as an edge to each cell would stand it.
        let n = ids.len();
        let all = offsets.len() - 1;
        let mut component = vec![0; all];
        let mut level: Vec<usize> = vec![0; all];
        let mut alone = vec![false; n];
        for (c, ends) in starts.windows(2).enumerate() {
            let comp = &members[ends[0]..ends[1]];
            for &v in comp {
                component[v as usize] = c;
            }
            let deepest = comp
                .iter()
                .flat_map(|&v| adj(v as usize))
                .filter(|&&u| component[u as usize] != c)
                .map(|&u| level[u as usize] + 1)
                .max()
                .unwrap_or(0);
            if comp.iter().all(|&v| v as usize >= n) {
                // A hub alone: its cells' deepest, one fewer than a reader's.
                for &v in comp {
                    level[v as usize] = deepest.saturating_sub(1);
                }
                continue;
            }
            for &v in comp {
                level[v as usize] = deepest;
            }
            if comp.len() == 1 && !adj(comp[0] as usize).contains(&comp[0]) {
                alone[comp[0] as usize] = true;
            } else {
                let mut cells: Vec<CellId> = comp
                    .iter()
                    .filter(|&&i| (i as usize) < n)
                    .map(|&i| ids[i as usize])
                    .collect();
                cells.sort_unstable();
                self.cycles.push(cells);
            }
        }
        self.cycles.sort();
        // By shape and then by level, each pass keeping the order before it,
        // so by level, shape and cell, a cell's sheet first among its own;
        // the second pass's counts are where each level ends (KERNEL.30).
        let cells: Vec<usize> = (0..n).filter(|&v| alone[v]).collect();
        let (cells, _) = by_key(&cells, self.shapes.len(), |v| shape_of[v]);
        let levels = level[..n].iter().max().map_or(0, |deepest| deepest + 1);
        let (order, ends) = by_key(&cells, levels, |v| level[v]);
        self.level_ends = ends;
        // The run each cell of the order joined, by the run's first position:
        // a cell joins the run before it when it is of the run's shape and
        // sheet and reads no cell of the run, so that the run's cells may be
        // computed in any order, all at once among them (KERNEL.20). A hub
        // names no cell of a run: its cells are those its shape's every cell
        // reads, so that one of its own shape would read itself, a cycle, and
        // a run is one shape's cells.
        let mut joined = vec![usize::MAX; n];
        let mut run = Run {
            start: 0,
            end: 0,
            shape: usize::MAX,
        };
        let mut shape_at = Vec::with_capacity(order.len());
        for v in order {
            let (id, shape) = (ids[v], shape_of[v]);
            let pos = self.order.len();
            let joins = shape == run.shape
                && id.0 == self.order[run.start].0
                && adj(v)
                    .iter()
                    .all(|&u| u as usize >= n || joined[u as usize] != run.start);
            if !joins {
                self.keep(run);
                run = Run {
                    start: pos,
                    end: pos,
                    shape,
                };
            }
            joined[v] = run.start;
            run.end = pos + 1;
            self.order.push(id);
            shape_at.push(shape as u32);
        }
        self.keep(run);
        shape_at
    }

    /// A run of the order kept to be evaluated at once when it is long
    /// enough and its shape parsed and not risky, so that every cell of it
    /// reads as the shape does (KERNEL.20). (A risky shape's cells all read
    /// as it does in a grid that loaded, every other being refused; the
    /// shape is evaluated cell by cell all the same.)
    fn keep(&mut self, run: Run) {
        let Some(shape) = self.shapes.get(run.shape) else {
            return;
        };
        if run.end - run.start >= RUN && !shape.risky && shape.expr.is_ok() {
            self.runs.push(run);
        }
    }

    /// How many cells of the order are in runs.
    #[cfg(test)]
    pub(crate) fn cells_in_runs(&self) -> usize {
        self.runs.iter().map(|r| r.end - r.start).sum()
    }

    /// The plan spelled whole, each map in order: every shape with its first
    /// cell, its fill, its text and its parse; every cell's place and the
    /// construct it leaves unread; the two maps a cell finds its shape by;
    /// the order, its cycles, runs and levels; the reach. KERNEL.34's oracle,
    /// a write leaving the plan a load of its grid makes, part for part.
    #[cfg(test)]
    pub(crate) fn spelled(&self) -> String {
        use std::fmt::Write;
        let mut out = String::new();
        for (k, s) in self.shapes.iter().enumerate() {
            let _ = writeln!(
                out,
                "shape {k} {:?} {:?} {:?} {:?} {:?} {:?} {:?} {:?} {:?}",
                s.anchor, s.fill, s.text, s.expr, s.missing, s.risky, s.holds, s.repeats, s.lowered
            );
        }
        let mut places: Vec<(&CellId, &Place)> = self.places.iter().collect();
        places.sort_by_key(|(id, _)| **id);
        let mut unread: Vec<(&CellId, &String)> = self.unread.iter().collect();
        unread.sort();
        let mut shared: Vec<(&(usize, u32), &usize)> = self.by_shared.iter().collect();
        shared.sort();
        let mut r1c1: Vec<(&String, &usize)> = self.by_r1c1.iter().collect();
        r1c1.sort();
        let _ = writeln!(out, "places {places:?}\nunread {unread:?}");
        let _ = writeln!(out, "by shared {shared:?}\nby r1c1 {r1c1:?}");
        let _ = writeln!(
            out,
            "order {:?}\ncycles {:?}\nruns {:?}\nlevels {:?}\narrays {:?}",
            self.order, self.cycles, self.runs, self.level_ends, self.arrays
        );
        out.push_str(&self.reach.spelled());
        out
    }

    /// Whether a read takes a sheet as an array (KERNEL.31): a shape reads a
    /// range of it, or a shape with runs reads a cell of it by a reference
    /// no `$` holds at both corners, which its runs read by slices of the
    /// sheet's array where one is kept. A sheet whose cells only cells alone
    /// read, or only references every `$` holds, a run reading such a cell
    /// once, is taken as none: a cell's read by an array's slot costs what
    /// its read by its position costs (measured in the item's scoping).
    pub(crate) fn read_as_array(&self, sheet: usize) -> bool {
        self.arrays.get(sheet).is_some_and(|&b| b)
    }

    /// Whether a sheet of the grid the plan was made in keeps its frame's
    /// values as an array (KERNEL.31): a read takes it as one, and its cells
    /// stand close enough together, by the rule a machine keeps it by
    /// (`dense::limit`), so that a door's census and a machine's name the
    /// same reads alike.
    pub(crate) fn keeps_array(&self, model: &Workbook, sheet: usize) -> bool {
        self.read_as_array(sheet)
            && model.sheets.get(sheet).is_some_and(|s| {
                let most = crate::machine::CELLS as u64;
                super::dense::fits(s.extent(), super::dense::limit(s.cells.len(), most))
            })
    }

    /// The first run of the order that ends after a position, by its index
    /// among the runs: the one the position is in, or the next (KERNEL.44, a
    /// step walking the runs beside the positions).
    pub(crate) fn run_after(&self, pos: usize) -> usize {
        self.runs.partition_point(|r| r.end <= pos)
    }

    /// A run by its index among the runs: its first position, the position
    /// after its last cell, and its shape; `None` past the last.
    #[inline]
    pub(crate) fn run_span(&self, k: usize) -> Option<(usize, usize, usize)> {
        self.runs.get(k).map(|r| (r.start, r.end, r.shape))
    }

    /// A shape's program and its first cell, from which each of its cells'
    /// offset is, by its index (KERNEL.44); `None` for a shape left to `Ctx`.
    #[inline]
    pub(crate) fn program(&self, index: usize) -> Option<(&Program, CellId)> {
        let s = self.shapes.get(index)?;
        match &s.lowered {
            Some(Ok(p)) => Some((p, s.anchor)),
            _ => None,
        }
    }

    /// Cells of one run, a part of it or the whole, evaluated at once
    /// against the grid as `env` shows it (KERNEL.20, `super::vector`), each
    /// cell's answer pushed to `out` in order: what [`Shapes::evaluate`]
    /// answers for each, the environment told each cell as it is computed;
    /// or, where every answer is a number, `out` left empty and the numbers
    /// in `numbers` (KERNEL.43).
    #[allow(clippy::too_many_arguments)]
    pub(crate) fn evaluate_run(
        &self,
        env: &dyn Env,
        library: &Library,
        shape: usize,
        cells: &[CellId],
        lanes: &mut Lanes,
        out: &mut Vec<Computed>,
        numbers: &mut Vec<f64>,
    ) {
        let s = &self.shapes[shape];
        // The paths the run takes are noted against its shape, in a test
        // (KERNEL.30's oracle: the census names each of them).
        #[cfg(test)]
        vector::taken::shape(shape);
        match &s.expr {
            // A parse holding an array: each cell by `Ctx` whole (KERNEL.8,
            // slice f).
            Ok(_) if s.arrayed => out.extend(cells.iter().map(|&id| {
                env.enter(id);
                self.evaluate(env, library, id)
            })),
            Ok(e) => {
                let held = s.holds;
                vector::evaluate(
                    env, library, e, held, &s.repeats, s.anchor, cells, lanes, out, numbers,
                );
            }
            // A run holds a parsed shape; were it not, each cell alone.
            Err(_) => out.extend(cells.iter().map(|&id| {
                env.enter(id);
                self.evaluate(env, library, id)
            })),
        }
    }

    /// The formula cells in the order of evaluation, each after everything
    /// it reads; the cells of a cycle are not in it.
    pub fn order(&self) -> &[CellId] {
        &self.order
    }

    /// What each shape reads, and where the cells of the order stand
    /// (KERNEL.20).
    pub(crate) fn reach(&self) -> &Reach {
        &self.reach
    }

    /// Every cycle, each its cells sorted, the cycles by their first cell.
    pub fn cycles(&self) -> &[Vec<CellId>] {
        &self.cycles
    }

    /// How many formula cells are placed.
    pub fn formulas(&self) -> usize {
        self.places.len()
    }

    /// How many shapes the formulas make.
    pub fn shapes(&self) -> usize {
        self.shapes.len()
    }

    /// Whether a cell holds a placed formula.
    pub fn holds(&self, id: CellId) -> bool {
        self.places.contains_key(&id)
    }

    /// A formula cell's place.
    pub fn place(&self, id: CellId) -> Option<Place> {
        self.places.get(&id).copied()
    }

    /// A shape, by its index.
    pub fn shape(&self, index: usize) -> Option<&Shape> {
        self.shapes.get(index)
    }

    /// The plan printed as rows, the census (KERNEL.30): a `(plan ...)` row
    /// of totals, a `(readers ...)` row for each sheet whose changed cells
    /// are checked against readers, then a `(shape ...)` row for each shape
    /// a cell holds, each read named by the path an engine's machine takes
    /// for it (`super::census`). Every sheet is read as a frame's own, as a
    /// workbook's are: this is the census of a plan [`Shapes::of`] made over
    /// a grid with no `.last` twin, each sheet kept as an array where a
    /// machine of the same grid keeps it (KERNEL.31, [`Shapes::keeps_array`]).
    /// A machine's, its twins read as it keeps them, is `Machine::census`.
    /// Computed when asked from the plan as it stands, which keeps nothing
    /// for it but where each level ends and which sheets a read takes as an
    /// array.
    pub fn census(&self, model: &Workbook, library: &Library) -> String {
        super::census::rows(self, model, library, &|s| self.keeps_array(model, s))
    }

    /// Where each level of the order ends, by level (KERNEL.30).
    pub(crate) fn level_ends(&self) -> &[u32] {
        &self.level_ends
    }

    /// Each run of the order as its first position, the position after its
    /// last cell and its shape, by first position (KERNEL.30).
    pub(crate) fn run_spans(&self) -> impl Iterator<Item = (usize, usize, usize)> + '_ {
        self.runs.iter().map(|r| (r.start, r.end, r.shape))
    }

    /// Every formula cell and its place, in no order (KERNEL.30).
    pub(crate) fn places(&self) -> impl Iterator<Item = (CellId, Place)> + '_ {
        self.places.iter().map(|(&id, &p)| (id, p))
    }

    /// Every shape the plan holds, a shape whose cells a write has since
    /// replaced among them (KERNEL.30).
    pub(crate) fn every_shape(&self) -> &[Shape] {
        &self.shapes
    }

    /// One formula cell evaluated against the grid as `env` shows it, the
    /// cell being computed being `env.here()`.
    pub fn evaluate(&self, env: &dyn Env, library: &Library, id: CellId) -> Computed {
        let Some(p) = self.places.get(&id) else {
            return Computed::NotComputed(Reason::Construct(String::new()));
        };
        let shape = &self.shapes[p.shape];
        if let Some(construct) = self.unread.get(&id) {
            return Computed::NotComputed(Reason::Construct(construct.clone()));
        }
        match &shape.expr {
            Err(construct) => Computed::NotComputed(Reason::Construct(construct.clone())),
            Ok(e) => {
                let ctx = Ctx::at(env, library, p.d_row, p.d_col);
                match ctx.eval(e).and_then(|arg| ctx.result(&arg)) {
                    Ok(v) => Computed::Value(v.settled()),
                    Err(reason) => Computed::NotComputed(reason),
                }
            }
        }
    }

    /// The first thing that stops the grid from being stepped, in the order
    /// [`super::Calc::first_refusal`] asks it: a cycle (`calc-cycle`), the
    /// first cell calling a function the library does not hold
    /// (`calc-function-not-computed`), the first cell whose text this
    /// version does not read (`calc-construct-not-read`); with the cell the
    /// refusal stands on, for its line. `only` limits the second and third
    /// to the cells a write placed.
    pub fn first_refusal(
        &self,
        model: &Workbook,
        only: Option<&[CellId]>,
    ) -> Option<(CellId, Refusal)> {
        if let Some(cycle) = self.cycles.first() {
            let cells: Vec<String> = cycle.iter().map(|id| spell_cell(model, *id)).collect();
            return Some((
                cycle[0],
                raise("calc-cycle", &[("cells", &cells.join(", "))]),
            ));
        }
        let mut ids: Vec<CellId> = match only {
            Some(cells) => cells
                .iter()
                .copied()
                .filter(|id| self.places.contains_key(id))
                .collect(),
            None => self.places.keys().copied().collect(),
        };
        // A defined name a cell's formula reads whose text the parse does
        // not read, which a load refuses with the constructs (KERNEL.8, slice
        // b): asked only of a workbook that defines a name whose text is
        // not read, and once a shape and a sheet, a name being resolved by
        // the sheet that reads it.
        let mut named: FxHashMap<(usize, usize), String> = FxHashMap::default();
        let names = (!model.defined_names.is_empty()).then(|| Names::of(model));
        if let Some(names) = names.filter(Names::any_unread) {
            let mut asked: FxHashMap<(usize, usize), ()> = FxHashMap::default();
            for id in &ids {
                let s = self.places[id].shape;
                if asked.insert((s, id.0), ()).is_some() {
                    continue;
                }
                if let Ok(e) = &self.shapes[s].expr {
                    if let Some(construct) = unread_name(model, &names, e, id.0) {
                        named.insert((s, id.0), construct);
                    }
                }
            }
        }
        let unread_named =
            |id: &CellId| -> Option<String> { named.get(&(self.places[id].shape, id.0)).cloned() };
        // A grid is refused nowhere far more often than somewhere: its cells
        // are sorted, to find the first refused, only when one is (KERNEL.34).
        let refused = |id: &CellId| {
            let shape = &self.shapes[self.places[id].shape];
            !shape.missing.is_empty()
                || shape.expr.is_err()
                || self.unread.contains_key(id)
                || unread_named(id).is_some()
        };
        if !ids.iter().any(refused) {
            return None;
        }
        ids.sort_unstable();
        for id in &ids {
            let shape = &self.shapes[self.places[id].shape];
            if let Some(name) = shape.missing.first() {
                return Some((
                    *id,
                    raise(
                        "calc-function-not-computed",
                        &[("cell", &spell_cell(model, *id)), ("function", name)],
                    ),
                ));
            }
        }
        for id in &ids {
            let p = self.places[id];
            let shape = &self.shapes[p.shape];
            let label = match (&shape.expr, self.unread.get(id).cloned()) {
                (_, Some(unread)) => Some(unread),
                // The construct as the cell's own text spells it.
                (Err(construct), None) if p.d_row == 0 && p.d_col == 0 => Some(construct.clone()),
                (Err(construct), None) => {
                    let moved = refers::shift_a1_references(&shape.text, p.d_row, p.d_col);
                    Some(
                        formula::parse(&moved)
                            .err()
                            .unwrap_or_else(|| construct.clone()),
                    )
                }
                (Ok(_), None) => unread_named(id),
            };
            if let Some(label) = label {
                return Some((
                    *id,
                    formula::unread_refusal(&spell_cell(model, *id), &label),
                ));
            }
        }
        None
    }
}

/// Cells ordered by a key below `keys`, cells of one key in the order they
/// came in: a counting sort, a pass over the cells and one over the keys,
/// for the order of a plan (KERNEL.20); and, by key, where its cells end in
/// that order, which the sort's own counts hold once it is done (KERNEL.30).
/// A plan's cells are positions of an order a `u32` holds, as its reach's
/// are.
fn by_key(cells: &[usize], keys: usize, key: impl Fn(usize) -> usize) -> (Vec<usize>, Vec<u32>) {
    let mut start = vec![0u32; keys + 1];
    for &v in cells {
        start[key(v) + 1] += 1;
    }
    for k in 0..keys {
        start[k + 1] += start[k];
    }
    let mut out = vec![0; cells.len()];
    for &v in cells {
        let k = key(v);
        out[start[k] as usize] = v;
        start[k] += 1;
    }
    start.truncate(keys);
    (out, start)
}

/// By sheet of the grid, whether a read takes it as an array (KERNEL.31,
/// [`Shapes::read_as_array`]): a sheet a shape reads a range of, or a cell of
/// by a reference no `$` holds at both corners in a shape with runs. A
/// reference that names no sheet reads each sheet its shape's cells stand
/// on; a span of sheets is read through its cells, and so is a defined
/// name's text. One pass over each shape's reads, and one over the plan's
/// cells, `ids` beside `shape_of`, for the shapes that read their own
/// sheets: two arrays read in order, where a walk of the places would look
/// each cell up (`plan`).
fn read_as_arrays(
    model: &Workbook,
    shapes: &[Shape],
    (ids, shape_of): (&[CellId], &[usize]),
    runs: &[Run],
) -> Vec<bool> {
    let mut with_runs = vec![false; shapes.len()];
    for r in runs {
        with_runs[r.shape] = true;
    }
    let mut read = vec![false; model.sheets.len()];
    // By shape, whether it reads its own sheets so.
    let mut own = vec![false; shapes.len()];
    for (k, shape) in shapes.iter().enumerate() {
        let Ok(e) = &shape.expr else {
            continue;
        };
        e.each_read(&mut |r| {
            let Read::Ref(r) = r else {
                return;
            };
            if r.last_sheet.is_some() {
                return;
            }
            let held = r.corners.row_abs1 && r.corners.col_abs1;
            let range = !matches!(r.shape, refers::Shape::Cell);
            if !range && (held || !with_runs[k]) {
                return;
            }
            let mut mark = |s: usize| {
                if let Some(b) = read.get_mut(s) {
                    *b = true;
                }
            };
            match (r.found, &r.sheet) {
                (Some((first, _)), _) => mark(usize::from(first)),
                (None, None) => own[k] = true,
                (None, Some(name)) => {
                    if let Some(s) = model.find_sheet(name) {
                        mark(s);
                    }
                }
            }
        });
    }
    for (id, &k) in ids.iter().zip(shape_of) {
        if own[k] {
            read[id.0] = true;
        }
    }
    read
}

/// A formula cell's vertex, its position among a plan's sorted cells, as an
/// edge holds it: a plan's cells are positions a `u32` holds, as its
/// order's are.
fn vertex(position: usize) -> u32 {
    position as u32
}

/// The fewest formula cells a shape's held references name for its cells to
/// read them through one hub (`Shapes::plan`, KERNEL.8).
const HUB: usize = 64;

#[cfg(test)]
thread_local! {
    /// The test switch: hubs from this many cells on instead.
    static HUB_AT: std::cell::Cell<usize> = const { std::cell::Cell::new(HUB) };
    /// The hubs the plans on this thread made, for a test to count.
    static HUBS_MADE: std::cell::Cell<usize> = const { std::cell::Cell::new(0) };
}

/// The fewest cells a hub stands for: [`HUB`], or the test switch's.
fn hub_at() -> usize {
    #[cfg(test)]
    {
        HUB_AT.with(std::cell::Cell::get)
    }
    #[cfg(not(test))]
    {
        HUB
    }
}

/// A plan's hubs made from this many cells on, on this thread, for a test.
#[cfg(test)]
pub(crate) fn hubs_at(cells: usize) {
    HUB_AT.with(|h| h.set(cells));
}

/// The hubs the plans on this thread made since the last asking.
#[cfg(test)]
pub(crate) fn hubs_made() -> usize {
    HUBS_MADE.with(|h| h.replace(0))
}

/// A plan's formula cells as the vertices of its edges (KERNEL.33): each
/// cell's position among the plan's sorted cells, and whether each sheet
/// holds any formula.
struct Vertices<'a> {
    model: &'a Workbook,
    extents: &'a [Extent],
    index: &'a Positions,
    holds: &'a [bool],
}

/// A shape's references as its plan reads them, once a shape (KERNEL.33):
/// the vertices those every `$` holds name, the same at every cell, sorted
/// and each once; those every `$` holds on the formula's own sheet, the
/// same at every cell of one sheet (KERNEL.8); the references each cell
/// moves, each once; and whether the parse reads a name, which each cell
/// resolves from its own sheet.
struct ShapeReads<'e> {
    fixed: Vec<u32>,
    own: Vec<&'e RefExpr>,
    moving: Vec<&'e RefExpr>,
    named: bool,
}

impl Vertices<'_> {
    /// The formula cells inside an area, each by its vertex, to `out`: by
    /// the array where the sheet's formulas stand close enough together to
    /// be held as one, else by the sheet's tree; none on a sheet holding no
    /// formula, from which no edge comes.
    fn formulas_in(&self, a: &Area, out: &mut Vec<u32>) {
        if !self.holds.get(a.sheet).copied().unwrap_or(false) {
            return;
        }
        let rect = (a.top, a.left, a.bottom, a.right);
        if self
            .index
            .each_in(a.sheet, rect, &mut |v| out.push(vertex(v)))
        {
            return;
        }
        let sheet = &self.model.sheets[a.sheet];
        let extent = self.extents.get(a.sheet).copied().flatten();
        for ((row, col), cell) in cells_in(sheet, extent, a) {
            if is_formula(&cell.content) {
                if let Some(v) = self.index.get(a.sheet, row, col) {
                    out.push(vertex(v));
                }
            }
        }
    }

    /// A shape's references read off its parse once: each once, however
    /// often the parse repeats it; one to the single sheet its placement
    /// found, when that sheet holds no formula, left out; one every `$`
    /// holds there, found now, since it names the same cells at every
    /// offset and `graph::areas_at` reads a found sheet and not the cell's;
    /// every other kept for each cell to move.
    fn reads_of<'e>(&self, shape: &'e Shape) -> ShapeReads<'e> {
        let mut reads = ShapeReads {
            fixed: Vec::new(),
            own: Vec::new(),
            moving: Vec::new(),
            named: false,
        };
        let Ok(e) = &shape.expr else {
            return reads;
        };
        let mut each: Vec<&RefExpr> = Vec::new();
        e.each_read(&mut |r| match r {
            Read::Ref(re) => {
                if !each.contains(&re) {
                    each.push(re);
                }
            }
            Read::Name { .. } => reads.named = true,
        });
        for re in each {
            let found_once = match (re.found, &re.last_sheet) {
                (Some((first, last)), None) if first == last => Some(usize::from(first)),
                _ => None,
            };
            if let Some(sheet) = found_once {
                if !self.holds.get(sheet).copied().unwrap_or(false) {
                    continue;
                }
                if every_dollar(re) {
                    if let Ok(areas) = areas_at(self.model, sheet, re, 0, 0) {
                        for a in &areas {
                            self.formulas_in(a, &mut reads.fixed);
                        }
                    }
                    continue;
                }
            }
            if re.sheet.is_none() && re.last_sheet.is_none() && every_dollar(re) {
                reads.own.push(re);
                continue;
            }
            reads.moving.push(re);
        }
        reads.fixed.sort_unstable();
        reads.fixed.dedup();
        reads
    }
}

/// Whether every `$` holds a reference, so that it names the same cells at
/// every offset (KERNEL.33): a cell's two, a range's four, a whole column's
/// columns, a whole row's rows.
fn every_dollar(r: &RefExpr) -> bool {
    let c = &r.corners;
    match r.shape {
        refers::Shape::Cell => c.row_abs1 && c.col_abs1,
        refers::Shape::Range => c.row_abs1 && c.col_abs1 && c.row_abs2 && c.col_abs2,
        refers::Shape::Column => c.col_abs1 && c.col_abs2,
        refers::Shape::Row => c.row_abs1 && c.row_abs2,
        refers::Shape::None => false,
    }
}

/// Every reference of a parse that names its sheets, those sheets found once
/// against the grid it is placed in (KERNEL.20): a machine's sheets, its
/// twins among them, are fixed for its life, so the index a name finds at
/// placement is the one it would find at every evaluation. A reference to
/// the formula's own sheet, and one naming a sheet the grid does not hold,
/// stay unfound and are found by name at each evaluation, as before. The
/// walk keeps its own stack, so that a chain of operators costs it no depth.
fn find_sheets(expr: &mut Expr, model: &Workbook) {
    let mut stack = vec![expr];
    while let Some(e) = stack.pop() {
        match e {
            Expr::Ref(r) => r.found = sheets_of(model, r),
            // A range's reach is found as a reference is, and its ends
            // (KERNEL.8, slice f).
            Expr::Range(r) => {
                let crate::calc::formula::RangeExpr {
                    left, right, reach, ..
                } = &mut **r;
                if let Some(reach) = reach {
                    reach.found = sheets_of(model, reach);
                }
                stack.push(left);
                stack.push(right);
            }
            Expr::Call(_, args) => stack.extend(args.iter_mut()),
            Expr::Neg(x) | Expr::Pos(x) | Expr::Percent(x) => stack.push(x),
            Expr::Binary(_, a, b) => {
                stack.push(a);
                stack.push(b);
            }
            Expr::Name { .. }
            | Expr::Number(_)
            | Expr::Text(_)
            | Expr::Bool(_)
            | Expr::Error(_)
            | Expr::Empty
            | Expr::Array(_) => {}
        }
    }
}

/// A reference's named sheets as `graph::areas_at` finds them by name: the
/// first, and the last of a span or the first again; `None` when it names
/// none, one the grid does not hold, or one past the 256th, which a byte
/// cannot hold and the name finds as before.
fn sheets_of(model: &Workbook, r: &RefExpr) -> Option<(u8, u8)> {
    let first = model.find_sheet(r.sheet.as_deref()?)?;
    let last = match &r.last_sheet {
        None => first,
        Some(name) => model.find_sheet(name)?,
    };
    Some((u8::try_from(first).ok()?, u8::try_from(last).ok()?))
}

/// `Sheet!A1` for a cell, the sheet quoted as a reference quotes it.
pub fn spell_cell(model: &Workbook, id: CellId) -> String {
    let sheet = model
        .sheets
        .get(id.0)
        .map(|s| s.name.as_str())
        .unwrap_or("");
    format!("{}!{}", quote_sheet(sheet), cell_ref(id.1, id.2))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn every_dollar_holds_a_reference_whose_every_number_is_held() {
        // KERNEL.33: a plan finds a reference once for its whole shape only
        // when every `$` holds it, so that it names the same cells at every
        // offset: a cell's two numbers, a range's four, a whole column's
        // columns, a whole row's rows. One unheld number moves it.
        for (text, held) in [
            ("$C$2", true),
            ("$C2", false),
            ("C$2", false),
            ("C2", false),
            ("$C$2:$D$4", true),
            ("$C$2:D4", false),
            ("$C$2:$D4", false),
            ("$C$2:D$4", false),
            ("C2:$D$4", false),
            ("$C:$D", true),
            ("$C:D", false),
            ("C:D", false),
            ("$2:$4", true),
            ("$2:4", false),
            ("2:4", false),
            ("S!$C$2:$D$4", true),
            ("S!C2:$D$4", false),
        ] {
            let Ok(Expr::Ref(r)) = formula::parse(text) else {
                panic!("{text} reads as one reference")
            };
            assert_eq!(every_dollar(&r), held, "{text}");
        }
    }
}
