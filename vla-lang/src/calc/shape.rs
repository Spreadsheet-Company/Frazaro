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
//! (`super::reach`, KERNEL.20).
//!
//! The plan is printed as rows on request ([`Shapes::census`], KERNEL.30,
//! `super::census`): each shape's cells, the levels of the order it spans,
//! its runs and the paths its reads take, so that a grid's cost is read
//! before a frame runs. Of all that, the plan keeps for it only where each
//! level of the order ends.
//!
//! Since KERNEL.49's decision 7 (its fourth slice, 2026-10-22) the plan
//! places nodes and not cells, each a fill's piece or a formula alone, and
//! keeps the order as blocks (`super::nodes`, `super::blocks`): no map of
//! every cell's place, no list of every cell of the order and no grid of
//! every cell's position, a cell's shape and position found from its node
//! and its block. The plan of every cell it was stays as the tests'
//! reference (`reference`), the order by blocks held to it cell for cell.

use std::collections::HashMap;

use crate::fx::FxHashMap;
use crate::grid::{Formula, Kind, Masters, Model};
use crate::messages::{raise, Refusal};
use crate::refers::{self, quote_sheet};
use crate::sheet::{cell_ref, strip_future_prefixes, A1Range, Workbook};

use super::blocks::Blocks;
use super::eval::{unread_name, Computed, Ctx, Env, Library, Reason};
use super::formula::{self, Expr, RefExpr};
use super::graph::{CellId, Extent, Names};
use super::nodes::{self, Alone, Located, Run};
use super::program::{Program, Why};
use super::reach::{Reach, Reached};
use super::vector::{self, Lanes, Repeats};

#[cfg(test)]
pub(crate) mod reference;

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
    /// Whether the text calls a function the library declares to read a
    /// sheet's extent itself (KERNEL.56, [`Library::reads_extent`]), read
    /// once a shape as `missing` is: the reach files the shape by every sheet
    /// it reads, for a moved extent to have it computed whole.
    pub(crate) reads_extent: bool,
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
    /// A cell of a risky shape whose own text the parse does not read: the
    /// construct it names, found once when the cell is placed.
    unread: FxHashMap<CellId, String>,
    /// The order as blocks (KERNEL.49's decision 7): a shape's cells on a
    /// sheet at a level each, row by row, the order a list of them.
    blocks: Blocks,
    /// Where each formula cell stands (KERNEL.49's decision 7): each
    /// formula alone by its place, with its shape, kept from one plan to the
    /// next, and each fill's piece by its rectangle, where a map of every
    /// formula cell's place stood.
    located: Located,
    /// How many formula cells are placed.
    formulas: usize,
    cycles: Vec<Vec<CellId>>,
    /// The runs of the order an engine evaluates at once (KERNEL.20), by
    /// their first position: each of at least [`RUN`] cells of one parsed
    /// shape that is not risky, on one sheet, consecutive in the order, none
    /// reading another.
    runs: Vec<Run>,
    /// What each shape reads, found with the order (KERNEL.20).
    reach: Reach,
    /// Where each level of the order ends, by level (KERNEL.30): the
    /// position after its last cell, so that the census finds each cell's
    /// level without planning again.
    level_ends: Vec<u32>,
    /// Each shape's stretches of the order (KERNEL.54), each its first
    /// position and the position after its last: the longest spans of the
    /// order its cells fill with no other shape's between, runs and cells
    /// alone, found with the order, so that a machine marks every cell of a
    /// shape a word at a time. A shape's are
    /// `stretches[stretch_firsts[s]..stretch_firsts[s + 1]]`.
    stretches: Vec<(u32, u32)>,
    stretch_firsts: Vec<u32>,
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
    /// value would stand on it, 600 bytes since it kept each shape's
    /// stretches (KERNEL.54).
    pub fn of(
        model: &Workbook,
        sheets: usize,
        library: &Library,
        names: &Names,
        extents: &[Extent],
    ) -> Shapes {
        Shapes::of_model(model, sheets, library, names, extents)
    }

    /// [`Shapes::of`] over a grid, a workbook's or an engine's machine's
    /// (KERNEL.49), its formula cells placed node by node (decision 7), each
    /// fill's piece and each formula alone in the order of its first cell,
    /// sheet by sheet, so that the shapes are numbered as a walk of every
    /// cell numbered them.
    pub(crate) fn of_model<M: Model + ?Sized>(
        model: &M,
        sheets: usize,
        library: &Library,
        names: &Names,
        extents: &[Extent],
    ) -> Shapes {
        let mut s: Box<Shapes> = Box::default();
        s.located.alone = vec![FxHashMap::default(); model.sheet_count()];
        let mut nodes: Vec<(usize, A1Range, Option<u32>)> = Vec::new();
        for si in 0..sheets.min(model.sheet_count()) {
            model.each_node(si, &mut |rect, f| {
                let fill = match f {
                    Formula::Fill(index) => Some(index),
                    Formula::Own(..) => None,
                };
                nodes.push((si, rect, fill));
            });
        }
        nodes.sort_unstable_by_key(|&(si, r, _)| (si, r.top, r.left));
        let mut masters: Option<(usize, Masters<'_>)> = None;
        for (si, rect, fill) in nodes {
            if masters.as_ref().map(|m| m.0) != Some(si) {
                masters = Some((si, model.masters(si)));
            }
            let formula = match fill {
                Some(index) => Formula::Fill(index),
                None => match model.formula((si, rect.top, rect.left)) {
                    Some(f) => f,
                    None => continue,
                },
            };
            let m = &masters.as_ref().expect("this sheet's masters").1;
            s.place_node(model, si, rect, formula, m, library);
        }
        s.plan_model(model, names, extents);
        *s
    }

    /// Place one cell after a write: its old place forgotten, and a new one
    /// found when it holds a formula. The order waits for the plan.
    #[cfg(test)]
    pub(crate) fn replace(&mut self, model: &Workbook, id: CellId, library: &Library) {
        if let Some(m) = self.located.alone.get_mut(id.0) {
            m.remove(&(id.1, id.2));
        }
        self.unread.remove(&id);
        let Some(formula) = model.formula(id) else {
            return;
        };
        let masters = model.masters(id.0);
        self.place_node(model, id.0, cell_range(id), formula, &masters, library);
    }

    /// Many cells placed after a write, each sheet's shared formulas found
    /// once for all its cells (KERNEL.34: Life's rule written whole walked
    /// its sheet 62,964 times, 3.96 billion cells visited, in 13.3 s); then
    /// the plan made the one a load of the same grid makes
    /// ([`Shapes::as_loaded`]). The order waits for the plan.
    pub(crate) fn replace_many<M: Model + ?Sized>(
        &mut self,
        model: &M,
        ids: &[CellId],
        library: &Library,
    ) {
        for id in ids {
            if let Some(m) = self.located.alone.get_mut(id.0) {
                m.remove(&(id.1, id.2));
            }
            self.unread.remove(id);
        }
        let mut sheets: Vec<usize> = ids.iter().map(|id| id.0).collect();
        sheets.sort_unstable();
        sheets.dedup();
        for s in sheets {
            if s >= model.sheet_count() {
                continue;
            }
            let masters = model.masters(s);
            for &id in ids.iter().filter(|id| id.0 == s) {
                if let Some(formula) = model.formula(id) {
                    self.place_node(model, s, cell_range(id), formula, &masters, library);
                }
            }
        }
        self.as_loaded(model, library);
    }

    /// A placed cell's shape as the grid now holds it: a fill's by its
    /// index, a formula alone's by its place.
    fn shape_in<M: Model + ?Sized>(&self, model: &M, id: CellId) -> Option<usize> {
        match model.formula(id)? {
            Formula::Fill(index) => self.by_shared.get(&(id.0, index)).copied(),
            Formula::Own(..) => self
                .located
                .alone
                .get(id.0)?
                .get(&(id.1, id.2))
                .map(|a| a.shape as usize),
        }
    }

    /// The plan made the one [`Shapes::of`] makes over the same grid, once a
    /// write has placed some of its cells again (KERNEL.34): the shapes in
    /// the order of their first cells in the grid's order, sheet, row and
    /// column; a fill's shape anchored at its master and a plain formula's
    /// at its first cell, read again from that cell's text where the write
    /// moved or respelled it; a shape no cell holds dropped. Every cell's
    /// shape and every construct a cell leaves unread are then a load's, and
    /// so, once planned, are the order, its runs and the reach, and the
    /// census prints a load's. A write left a fill's other cells on the
    /// shape of its first cell even when that cell had been written over,
    /// and kept a shape no cell held, so that the census after a write was
    /// numbered and anchored otherwise than a load's.
    fn as_loaded<M: Model + ?Sized>(&mut self, model: &M, library: &Library) {
        // Each shape's first cell, the least of its cells, found in one pass
        // over its nodes: a fill's pieces and the formulas alone.
        let mut first: Vec<Option<CellId>> = vec![None; self.shapes.len()];
        let least = |first: &mut Vec<Option<CellId>>, s: usize, id: CellId| match first[s] {
            Some(f) if f <= id => {}
            _ => first[s] = Some(id),
        };
        for si in 0..model.sheet_count() {
            model.each_node(si, &mut |rect, f| {
                if let Formula::Fill(index) = f {
                    if let Some(&s) = self.by_shared.get(&(si, index)) {
                        least(&mut first, s, (si, rect.top, rect.left));
                    }
                }
            });
        }
        for (si, m) in self.located.alone.iter().enumerate() {
            for (&(row, col), a) in m {
                least(&mut first, a.shape as usize, (si, row, col));
            }
        }
        let mut order: Vec<usize> = (0..self.shapes.len())
            .filter(|&s| first[s].is_some())
            .collect();
        order.sort_unstable_by_key(|&s| first[s]);
        // Each shape's first cell as a load anchors it, with its text there.
        let mut masters: HashMap<usize, Masters<'_>> = HashMap::new();
        let mut again: Vec<(usize, CellId, String)> = Vec::new();
        for &s in &order {
            let f = first[s].expect("a shape in the order has a first cell");
            let shape = &self.shapes[s];
            let (anchor, raw) = match shape.fill {
                Some(index) => {
                    let sheet = shape.anchor.0;
                    let found = masters
                        .entry(sheet)
                        .or_insert_with(|| model.masters(sheet))
                        .get(&index);
                    match found {
                        Some(&(text, row, col)) => ((sheet, row, col), text),
                        None => continue,
                    }
                }
                None => match model.formula(f) {
                    Some(Formula::Own(t, _)) => (f, t),
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
        // the order of its first cell, each formula alone's renumbered and,
        // for a shape read again, each of its cells' constructs found again
        // from its new first cell.
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
        for m in &mut self.located.alone {
            for a in m.values_mut() {
                a.shape = number[a.shape as usize].expect("a placed cell's shape is kept") as u32;
            }
        }
        self.by_shared.retain(|_, s| number[*s].is_some());
        self.by_r1c1.retain(|_, s| number[*s].is_some());
        for s in self.by_shared.values_mut().chain(self.by_r1c1.values_mut()) {
            *s = number[*s].expect("kept");
        }
        if !moved.contains(&true) {
            return;
        }
        // The constructs of the cells of a shape read again, found again at
        // their new offsets: those it had dropped, and its cells' found
        // where it is risky.
        let unread = std::mem::take(&mut self.unread);
        self.unread = unread
            .into_iter()
            .filter(|(id, _)| self.shape_in(model, *id).map_or(true, |s| !moved[s]))
            .collect();
        let mut cells: Vec<(CellId, usize)> = Vec::new();
        for si in 0..model.sheet_count() {
            model.each_node(si, &mut |rect, f| {
                let s = match f {
                    Formula::Fill(index) => self.by_shared.get(&(si, index)).copied(),
                    Formula::Own(..) => self
                        .located
                        .alone
                        .get(si)
                        .and_then(|m| m.get(&(rect.top, rect.left)))
                        .map(|a| a.shape as usize),
                };
                if let Some(s) = s.filter(|&s| moved[s] && self.shapes[s].risky) {
                    for row in rect.top..=rect.bottom {
                        for col in rect.left..=rect.right {
                            cells.push(((si, row, col), s));
                        }
                    }
                }
            });
        }
        for (id, s) in cells {
            let shape = &self.shapes[s];
            let (d_row, d_col) = (
                i64::from(id.1) - i64::from(shape.anchor.1),
                i64::from(id.2) - i64::from(shape.anchor.2),
            );
            if let Some(construct) = shape.unread_at(d_row, d_col) {
                self.unread.insert(id, construct);
            }
        }
    }

    /// One node placed (KERNEL.49's decision 7): a fill's piece on its
    /// fill's shape, made at its first piece and anchored at its master; a
    /// formula alone on the shape of its R1C1 text, kept by its place; and
    /// each cell of a risky shape whose own text the parse does not read
    /// with the construct it names.
    fn place_node<M: Model + ?Sized>(
        &mut self,
        model: &M,
        si: usize,
        rect: A1Range,
        formula: Formula<'_>,
        masters: &Masters<'_>,
        library: &Library,
    ) {
        let shape = match formula {
            Formula::Fill(index) => {
                let Some(&(text, m_row, m_col)) = masters.get(&index) else {
                    // A child whose master is gone has no text; the machine
                    // never leaves one (a write over a master makes the
                    // fill's first cell left its master, or gives each of its
                    // cells its own text, first).
                    return;
                };
                let key = (si, index);
                match self.by_shared.get(&key) {
                    Some(&s) => s,
                    None => {
                        let anchor = (si, m_row, m_col);
                        let s = self.add(model, strip_future_prefixes(text), anchor, library);
                        self.shapes[s].fill = Some(index);
                        self.by_shared.insert(key, s);
                        s
                    }
                }
            }
            Formula::Own(t, kind) => {
                let (row, col) = (rect.top, rect.left);
                let text = strip_future_prefixes(t);
                let mut key = refers::r1c1(&text, row, col);
                // A dynamic formula and a legacy array formula are read as
                // arrays, a shape of their own beside a legacy formula of the
                // same text (KERNEL.8, slice f).
                match kind {
                    Kind::Dynamic => key.insert(0, '\u{1}'),
                    Kind::Array => key.insert(0, '\u{2}'),
                    Kind::Plain => {}
                }
                let s = match self.by_r1c1.get(&key) {
                    Some(&s) => s,
                    None => {
                        let s = self.add(model, text, (si, row, col), library);
                        self.by_r1c1.insert(key, s);
                        s
                    }
                };
                if self.located.alone.len() <= si {
                    self.located.alone.resize_with(si + 1, FxHashMap::default);
                }
                self.located.alone[si].insert(
                    (row, col),
                    Alone {
                        shape: s as u32,
                        pos: nodes::NONE,
                    },
                );
                s
            }
        };
        let shape = &self.shapes[shape];
        if !shape.risky || shape.expr.is_err() {
            return;
        }
        for row in rect.top..=rect.bottom {
            for col in rect.left..=rect.right {
                let (d_row, d_col) = (
                    i64::from(row) - i64::from(shape.anchor.1),
                    i64::from(col) - i64::from(shape.anchor.2),
                );
                if let Some(construct) = shape.unread_at(d_row, d_col) {
                    self.unread.insert((si, row, col), construct);
                }
            }
        }
    }

    fn add<M: Model + ?Sized>(
        &mut self,
        model: &M,
        text: String,
        anchor: CellId,
        library: &Library,
    ) -> usize {
        let mut expr = formula::parse(&text);
        if let Ok(e) = expr.as_mut() {
            find_sheets(e, model);
        }
        // The calls as the text's tokens hold them, parsed or not, as
        // `Calc::missing_functions` reads them; and whether one reads an
        // extent (KERNEL.56).
        let calls = formula::calls(&text);
        let reads_extent = calls.iter().any(|name| library.reads_extent(name));
        let mut missing: Vec<String> = calls
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
        let kind_reads_arrays = matches!(
            model.formula(anchor),
            Some(Formula::Own(_, Kind::Dynamic | Kind::Array))
        );
        let s = self.push(text, expr, anchor, missing, risky, library);
        self.shapes[s].reads_extent = reads_extent;
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
            reads_extent: false,
        });
        self.shapes.len() - 1
    }

    /// The order found again from the placed nodes (KERNEL.49's decision 7,
    /// `super::nodes`): each node's edges by the rectangles its shape's
    /// references read, Tarjan's components over the nodes, each node's
    /// cells' levels a map of rectangles, a node that reads itself, stands
    /// in a component of more than one or reads what the reach does not
    /// follow planned cell by cell as every cell was (KERNEL.33's edges,
    /// KERNEL.8's hubs), the order kept as blocks; then the reach. The order,
    /// its runs, levels and cycles are those a plan of every cell finds,
    /// cell for cell (the tests' reference, `reference::Reference`).
    pub(crate) fn plan_model<M: Model + ?Sized>(
        &mut self,
        model: &M,
        names: &Names,
        extents: &[Extent],
    ) {
        let alone = std::mem::take(&mut self.located.alone);
        let planned = nodes::plan(
            model,
            &self.shapes,
            &self.by_shared,
            &self.unread,
            names,
            extents,
            alone,
        );
        #[cfg(test)]
        HUBS_MADE.with(|h| h.set(h.get() + planned.hubs));
        #[cfg(test)]
        BY_CELL.with(|h| h.set(h.get() + planned.by_cell));
        self.blocks = planned.blocks;
        self.located = planned.located;
        self.cycles = planned.cycles;
        self.runs = planned.runs;
        self.level_ends = planned.level_ends;
        self.stretches = planned.stretches;
        self.stretch_firsts = planned.stretch_firsts;
        self.formulas = planned.formulas;
        self.reach = Reach::of(model, &self.shapes, &self.blocks, &self.located);
        #[cfg(test)]
        reference::hold(self, model, names, extents);
    }

    /// A shape's stretches of the order, each its first position and the
    /// position after its last (KERNEL.54): every position its cells stand
    /// at, and none other.
    pub(crate) fn stretches(&self, shape: usize) -> &[(u32, u32)] {
        match (
            self.stretch_firsts.get(shape),
            self.stretch_firsts.get(shape + 1),
        ) {
            (Some(&a), Some(&b)) => &self.stretches[a as usize..b as usize],
            _ => &[],
        }
    }

    /// The order walked whole from its blocks, the tests' list of it.
    #[cfg(test)]
    pub(crate) fn order_cells(&self) -> Vec<CellId> {
        self.order().collect()
    }

    /// Everything a changed cell reaches, every reader of its sheet checked
    /// (`Reach::each_by_scan`), the tests' reference.
    #[cfg(test)]
    pub(crate) fn reached_by_scan(&self, changed: CellId, f: &mut impl FnMut(Reached)) {
        self.reach
            .each_by_scan(&self.blocks, &self.located, changed, f);
    }

    /// How many cells of the order are in runs.
    #[cfg(test)]
    pub(crate) fn cells_in_runs(&self) -> usize {
        self.runs.iter().map(|r| r.end - r.start).sum()
    }

    /// The plan spelled whole, each map in order: every shape with its first
    /// cell, its fill, its text and its parse; every formula alone with its
    /// shape and position, every fill's piece, the constructs cells leave
    /// unread and the two maps a cell finds its shape by; the order as its
    /// blocks, its cycles, runs and levels, and each shape's stretches
    /// (KERNEL.54); the reach, its pieces among its stands. KERNEL.34's
    /// oracle, a write leaving the plan a load of its grid makes, part for
    /// part.
    #[cfg(test)]
    pub(crate) fn spelled(&self) -> String {
        use std::fmt::Write;
        let mut out = String::new();
        for (k, s) in self.shapes.iter().enumerate() {
            let _ = writeln!(
                out,
                "shape {k} {:?} {:?} {:?} {:?} {:?} {:?} {:?} {:?} {:?} {:?}",
                s.anchor,
                s.fill,
                s.text,
                s.expr,
                s.missing,
                s.risky,
                s.holds,
                s.repeats,
                s.lowered,
                s.reads_extent
            );
        }
        let mut alone: Vec<(usize, (u32, u32), Alone)> = self
            .located
            .alone
            .iter()
            .enumerate()
            .flat_map(|(s, m)| m.iter().map(move |(&at, &a)| (s, at, a)))
            .collect();
        alone.sort_by_key(|&(s, at, _)| (s, at));
        let mut pieces: Vec<_> = self
            .located
            .pieces()
            .map(|(s, r, shape)| (s, (r.top, r.left, r.bottom, r.right), shape))
            .collect();
        pieces.sort_unstable();
        let mut unread: Vec<(&CellId, &String)> = self.unread.iter().collect();
        unread.sort();
        let mut shared: Vec<(&(usize, u32), &usize)> = self.by_shared.iter().collect();
        shared.sort();
        let mut r1c1: Vec<(&String, &usize)> = self.by_r1c1.iter().collect();
        r1c1.sort();
        let _ = writeln!(out, "alone {alone:?}\npieces {pieces:?}\nunread {unread:?}");
        let _ = writeln!(out, "by shared {shared:?}\nby r1c1 {r1c1:?}");
        let _ = writeln!(
            out,
            "blocks {:?}\nformulas {}\ncycles {:?}\nruns {:?}\nlevels {:?}",
            self.blocks, self.formulas, self.cycles, self.runs, self.level_ends
        );
        let _ = writeln!(
            out,
            "stretches {:?}\nstretch firsts {:?}",
            self.stretches, self.stretch_firsts
        );
        out.push_str(&self.reach.spelled());
        out
    }

    /// Whether a sheet of the grid the plan was made in keeps its values as
    /// an array (KERNEL.31; KERNEL.49): its cells stand close enough
    /// together, by the rule a machine keeps every sheet's store by
    /// (`dense::limit`), so that a door's census and a machine's name the
    /// same reads alike. Since KERNEL.49 a machine keeps every sheet of its
    /// rows in a store by place, where KERNEL.31 kept an array only for a
    /// sheet the plan read by a range or a run's moving cell.
    pub(crate) fn keeps_array(model: &Workbook, sheet: usize) -> bool {
        model.sheets.get(sheet).is_some_and(|s| {
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
                self.evaluate_shaped(env, library, shape, id)
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
                self.evaluate_shaped(env, library, shape, id)
            })),
        }
    }

    /// The formula cells in the order of evaluation, each after everything
    /// it reads; the cells of a cycle are not in it. Walked from the order's
    /// blocks (KERNEL.49's decision 7), where it was a list of its cells.
    pub fn order(&self) -> impl Iterator<Item = CellId> + '_ {
        self.blocks.cells(0, self.blocks.len())
    }

    /// How many cells the order holds.
    pub(crate) fn len(&self) -> usize {
        self.blocks.len()
    }

    /// The cell at a position of the order, and its shape.
    pub(crate) fn at(&self, p: usize) -> (CellId, usize) {
        self.blocks.cell(p)
    }

    /// The order's blocks.
    pub(crate) fn blocks(&self) -> &Blocks {
        &self.blocks
    }

    /// A formula cell's position in the order, where an engine's machine
    /// marks it (KERNEL.26); `None` for a cell holding no formula and for a
    /// cell on a cycle. Found from its node and its block (KERNEL.49's
    /// decision 7), where a grid of every cell's position was asked.
    pub(crate) fn position(&self, id: CellId) -> Option<usize> {
        self.located.position(&self.blocks, id)
    }

    /// Everything a changed cell reaches, to `f` ([`Reach::each`]), the
    /// cells of a shape found through its nodes and blocks.
    pub(crate) fn reached(&self, changed: CellId, f: &mut impl FnMut(Reached)) {
        self.reach.each(&self.blocks, &self.located, changed, f);
    }

    /// What each shape reads (KERNEL.20).
    pub(crate) fn reach(&self) -> &Reach {
        &self.reach
    }

    /// Every cycle, each its cells sorted, the cycles by their first cell.
    pub fn cycles(&self) -> &[Vec<CellId>] {
        &self.cycles
    }

    /// How many formula cells are placed.
    pub fn formulas(&self) -> usize {
        self.formulas
    }

    /// How many shapes the formulas make.
    pub fn shapes(&self) -> usize {
        self.shapes.len()
    }

    /// Whether a cell holds a placed formula.
    pub fn holds(&self, id: CellId) -> bool {
        self.located.holds(id)
    }

    /// A formula cell's place: its shape, found from its node (KERNEL.49's
    /// decision 7), and its offset from the shape's first cell.
    pub fn place(&self, id: CellId) -> Option<Place> {
        let shape = self.located.shape_of(id)?;
        let a = self.shapes[shape].anchor;
        Some(Place {
            shape,
            d_row: i64::from(id.1) - i64::from(a.1),
            d_col: i64::from(id.2) - i64::from(a.2),
        })
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
    /// machine of the same grid keeps it (KERNEL.31; KERNEL.49,
    /// [`Shapes::keeps_array`]). A machine's, its twins read as it keeps them,
    /// is `Machine::census`. Computed when asked from the plan as it stands,
    /// which keeps nothing for it but where each level ends.
    pub fn census(&self, model: &Workbook, library: &Library) -> String {
        super::census::rows(self, model, library, &|s| Shapes::keeps_array(model, s))
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

    /// Every node, its sheet, its rectangle and its shape, in no order: each
    /// fill's piece and each formula alone (KERNEL.30's census; KERNEL.49's
    /// decision 7, where it walked every cell's place).
    pub(crate) fn each_node(&self, f: &mut dyn FnMut(usize, A1Range, usize)) {
        for (sheet, rect, shape) in self.located.pieces() {
            f(sheet, rect, shape);
        }
        for (sheet, m) in self.located.alone.iter().enumerate() {
            for (&(row, col), a) in m {
                f(sheet, cell_range((sheet, row, col)), a.shape as usize);
            }
        }
    }

    /// Every shape the plan holds, a shape whose cells a write has since
    /// replaced among them (KERNEL.30).
    pub(crate) fn every_shape(&self) -> &[Shape] {
        &self.shapes
    }

    /// One formula cell evaluated against the grid as `env` shows it, the
    /// cell being computed being `env.here()`.
    pub fn evaluate(&self, env: &dyn Env, library: &Library, id: CellId) -> Computed {
        let Some(shape) = self.located.shape_of(id) else {
            return Computed::NotComputed(Reason::Construct(String::new()));
        };
        self.evaluate_shaped(env, library, shape, id)
    }

    /// One formula cell of a shape evaluated, its shape known (KERNEL.49's
    /// decision 7: a cell of the order's by its block's).
    pub(crate) fn evaluate_shaped(
        &self,
        env: &dyn Env,
        library: &Library,
        shape: usize,
        id: CellId,
    ) -> Computed {
        let s = &self.shapes[shape];
        if let Some(construct) = self.unread.get(&id) {
            return Computed::NotComputed(Reason::Construct(construct.clone()));
        }
        match &s.expr {
            Err(construct) => Computed::NotComputed(Reason::Construct(construct.clone())),
            Ok(e) => {
                let d_row = i64::from(id.1) - i64::from(s.anchor.1);
                let d_col = i64::from(id.2) - i64::from(s.anchor.2);
                let ctx = Ctx::at(env, library, d_row, d_col);
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
        self.first_refusal_in(model, only)
    }

    /// [`Shapes::first_refusal`] over a grid, a workbook's or an engine's
    /// machine's (KERNEL.49). Every cell of a node refuses alike but for a
    /// construct one cell leaves unread, so the first refused cell is found
    /// node by node (decision 7), each node's first cell standing for it,
    /// where every placed cell was sorted.
    pub(crate) fn first_refusal_in<M: Model + ?Sized>(
        &self,
        model: &M,
        only: Option<&[CellId]>,
    ) -> Option<(CellId, Refusal)> {
        if let Some(cycle) = self.cycles.first() {
            let cells: Vec<String> = cycle.iter().map(|id| spell_at(model, *id)).collect();
            return Some((
                cycle[0],
                raise("calc-cycle", &[("cells", &cells.join(", "))]),
            ));
        }
        // A defined name a cell's formula reads whose text the parse does
        // not read, which a load refuses with the constructs (KERNEL.8, slice
        // b): asked only of a workbook that defines a name whose text is
        // not read, and once a shape and a sheet, a name being resolved by
        // the sheet that reads it.
        let mut named: FxHashMap<(usize, usize), String> = FxHashMap::default();
        let names = (!model.defined_names().is_empty()).then(|| Names::of_model(model));
        let names = names.filter(Names::any_unread);
        let mut asked: FxHashMap<(usize, usize), ()> = FxHashMap::default();
        let mut ask = |s: usize, sheet: usize, named: &mut FxHashMap<(usize, usize), String>| {
            let Some(names) = &names else {
                return;
            };
            if asked.insert((s, sheet), ()).is_some() {
                return;
            }
            if let Ok(e) = &self.shapes[s].expr {
                if let Some(construct) = unread_name(model, names, e, sheet) {
                    named.insert((s, sheet), construct);
                }
            }
        };
        // The first cell calling a function the library lacks, and the first
        // standing for a text not read.
        let (missing, labelled) = match only {
            Some(cells) => {
                let mut ids: Vec<(CellId, usize)> = cells
                    .iter()
                    .filter_map(|&id| Some((id, self.located.shape_of(id)?)))
                    .collect();
                ids.sort_unstable();
                for &(id, s) in &ids {
                    ask(s, id.0, &mut named);
                }
                let missing = ids
                    .iter()
                    .find(|(_, s)| !self.shapes[*s].missing.is_empty())
                    .copied();
                let labelled = ids
                    .iter()
                    .find(|(id, s)| {
                        self.unread.contains_key(id)
                            || self.shapes[*s].expr.is_err()
                            || named.contains_key(&(*s, id.0))
                    })
                    .copied();
                (missing, labelled)
            }
            None => {
                let mut nodes: Vec<(CellId, usize)> = Vec::new();
                self.each_node(&mut |sheet, rect, s| nodes.push(((sheet, rect.top, rect.left), s)));
                for &(id, s) in &nodes {
                    ask(s, id.0, &mut named);
                }
                let least = |a: Option<(CellId, usize)>, b: (CellId, usize)| match a {
                    Some(x) if x.0 <= b.0 => Some(x),
                    _ => Some(b),
                };
                let mut missing = None;
                let mut labelled = None;
                for &(id, s) in &nodes {
                    let shape = &self.shapes[s];
                    if !shape.missing.is_empty() {
                        missing = least(missing, (id, s));
                    }
                    if shape.expr.is_err() || named.contains_key(&(s, id.0)) {
                        labelled = least(labelled, (id, s));
                    }
                }
                for &id in self.unread.keys() {
                    if let Some(s) = self.located.shape_of(id) {
                        labelled = least(labelled, (id, s));
                    }
                }
                (missing, labelled)
            }
        };
        if let Some((id, s)) = missing {
            let name = self.shapes[s].missing.first().expect("a missing function");
            return Some((
                id,
                raise(
                    "calc-function-not-computed",
                    &[("cell", &spell_at(model, id)), ("function", name)],
                ),
            ));
        }
        let (id, s) = labelled?;
        let shape = &self.shapes[s];
        let (d_row, d_col) = (
            i64::from(id.1) - i64::from(shape.anchor.1),
            i64::from(id.2) - i64::from(shape.anchor.2),
        );
        let label = match (&shape.expr, self.unread.get(&id).cloned()) {
            (_, Some(unread)) => unread,
            // The construct as the cell's own text spells it.
            (Err(construct), None) if d_row == 0 && d_col == 0 => construct.clone(),
            (Err(construct), None) => {
                let moved = refers::shift_a1_references(&shape.text, d_row, d_col);
                formula::parse(&moved)
                    .err()
                    .unwrap_or_else(|| construct.clone())
            }
            (Ok(_), None) => named.get(&(s, id.0)).cloned()?,
        };
        Some((id, formula::unread_refusal(&spell_at(model, id), &label)))
    }
}

/// A cell as a rectangle of one.
fn cell_range((_, row, col): CellId) -> A1Range {
    A1Range {
        top: row,
        left: col,
        bottom: row,
        right: col,
    }
}

/// The fewest formula cells a shape's held references name for its cells to
/// read them through one hub (KERNEL.8; since KERNEL.49's decision 7 the
/// cells planned cell by cell, `super::nodes`).
const HUB: usize = 64;

#[cfg(test)]
thread_local! {
    /// The test switch: hubs from this many cells on instead.
    static HUB_AT: std::cell::Cell<usize> = const { std::cell::Cell::new(HUB) };
    /// The hubs the plans on this thread made, for a test to count.
    static HUBS_MADE: std::cell::Cell<usize> = const { std::cell::Cell::new(0) };
    /// The cells the plans on this thread planned cell by cell (KERNEL.49's
    /// decision 7), for a test to count.
    static BY_CELL: std::cell::Cell<usize> = const { std::cell::Cell::new(0) };
}

/// The fewest cells a hub stands for: [`HUB`], or the test switch's.
pub(crate) fn hub_at() -> usize {
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

/// The cells the plans on this thread planned cell by cell since the last
/// asking.
#[cfg(test)]
pub(crate) fn planned_by_cell() -> usize {
    BY_CELL.with(|h| h.replace(0))
}

/// Every reference of a parse that names its sheets, those sheets found once
/// against the grid it is placed in (KERNEL.20): a machine's sheets, its
/// twins among them, are fixed for its life, so the index a name finds at
/// placement is the one it would find at every evaluation. A reference to
/// the formula's own sheet, and one naming a sheet the grid does not hold,
/// stay unfound and are found by name at each evaluation, as before. The
/// walk keeps its own stack, so that a chain of operators costs it no depth.
fn find_sheets<M: Model + ?Sized>(expr: &mut Expr, model: &M) {
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
fn sheets_of<M: Model + ?Sized>(model: &M, r: &RefExpr) -> Option<(u8, u8)> {
    let first = model.find_sheet(r.sheet.as_deref()?)?;
    let last = match &r.last_sheet {
        None => first,
        Some(name) => model.find_sheet(name)?,
    };
    Some((u8::try_from(first).ok()?, u8::try_from(last).ok()?))
}

/// `Sheet!A1` for a cell, the sheet quoted as a reference quotes it.
pub fn spell_cell(model: &Workbook, id: CellId) -> String {
    spell_at(model, id)
}

/// [`spell_cell`] in a grid, a workbook's or an engine's machine's
/// (KERNEL.49).
pub(crate) fn spell_at<M: Model + ?Sized>(model: &M, id: CellId) -> String {
    let sheet = model.sheet_name(id.0).unwrap_or("");
    format!("{}!{}", quote_sheet(sheet), cell_ref(id.1, id.2))
}
