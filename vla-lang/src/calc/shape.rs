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
//! each run at once, node by node (`super::vector`, KERNEL.20). What each
//! shape reads is found with the order, so that an engine follows a changed
//! cell to the cells of the order that read it and computes those alone
//! (`super::reach`, KERNEL.20).

use std::collections::HashMap;

use crate::fx::FxHashMap;
use crate::messages::{raise, Refusal};
use crate::refers::{self, quote_sheet};
use crate::sheet::{cell_ref, shared_masters, strip_future_prefixes, Content, Workbook};

use super::eval::{Computed, Ctx, Env, Library, Reason};
use super::formula::{self, Expr, Read, RefExpr};
use super::graph::{self, areas_at, cells_in, dependencies, is_formula, CellId, Extent, Names};
use super::reach::Reach;
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
}

impl Shape {
    /// The most buffers of each kind a run of the shape holds at once.
    #[cfg(test)]
    pub(crate) fn holds(&self) -> (usize, usize) {
        self.holds
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

impl Shapes {
    /// A plan with nothing in it.
    pub fn new() -> Shapes {
        Shapes::default()
    }

    /// Every formula cell of the sheets `0..sheets` placed, the shapes
    /// parsed, and the order found.
    pub fn of(
        model: &Workbook,
        sheets: usize,
        library: &Library,
        names: &Names,
        extents: &[Extent],
    ) -> Shapes {
        let mut s = Shapes::new();
        for (si, sheet) in model.sheets.iter().enumerate().take(sheets) {
            let masters = shared_masters(sheet);
            for (&(row, col), cell) in &sheet.cells {
                s.place_with(model, (si, row, col), &cell.content, &masters, library);
            }
        }
        s.plan(model, names, extents);
        s
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
                    // never leaves one (a write over a master gives its
                    // other cells their own text first).
                    return;
                };
                let anchor = (si, m_row, m_col);
                let key = (si, *index);
                let shape = match self.by_shared.get(&key) {
                    Some(&s) => s,
                    None => {
                        let s = self.add(model, strip_future_prefixes(text), anchor, library);
                        self.by_shared.insert(key, s);
                        s
                    }
                };
                (shape, anchor)
            }
            Content::Formula(t) | Content::DynamicFormula(t) => {
                let text = strip_future_prefixes(t);
                let key = refers::r1c1(&text, row, col);
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
        self.push(text, expr, anchor, missing, risky, library)
    }

    /// A shape kept, with what an evaluation of it at many cells holds found
    /// once (KERNEL.20, `vector::holds`) and the parts it writes more than
    /// once (KERNEL.27, `vector::Repeats`), in a frame of its own: `add`'s
    /// frame stands beneath the parse, the deepest a load reaches, which
    /// `VLA.1`'s limits pin on a stack of 1 MiB.
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
        self.shapes.push(Shape {
            text,
            expr,
            anchor,
            missing,
            risky,
            holds,
            repeats,
        });
        self.shapes.len() - 1
    }

    /// The order found again from the places: every formula cell's edges
    /// read off its shape's parse at its offset (a defined name's text read
    /// as `graph::dependencies` reads it), then Tarjan's components, as
    /// `graph::plan` finds them for a plan of texts.
    pub fn plan(&mut self, model: &Workbook, names: &Names, extents: &[Extent]) {
        let mut ids: Vec<CellId> = self.places.keys().copied().collect();
        ids.sort_unstable();
        let index: HashMap<CellId, usize> =
            ids.iter().enumerate().map(|(i, id)| (*id, i)).collect();
        let mut holds = vec![false; model.sheets.len()];
        for id in &ids {
            holds[id.0] = true;
        }
        let mut adj: Vec<Vec<usize>> = vec![Vec::new(); ids.len()];
        let mut shape_of: Vec<usize> = Vec::with_capacity(ids.len());
        let mut deps: Vec<CellId> = Vec::new();
        for (vi, id) in ids.iter().enumerate() {
            deps.clear();
            let p = self.places[id];
            shape_of.push(p.shape);
            let shape = &self.shapes[p.shape];
            let read = !self.unread.contains_key(id);
            match &shape.expr {
                Ok(e) if read => e.each_read(&mut |r| match r {
                    Read::Ref(re) => {
                        let Ok(areas) = areas_at(model, id.0, re, p.d_row, p.d_col) else {
                            return;
                        };
                        for a in &areas {
                            if !holds.get(a.sheet).copied().unwrap_or(false) {
                                continue;
                            }
                            let sheet = &model.sheets[a.sheet];
                            let extent = extents.get(a.sheet).copied().flatten();
                            for ((row, col), cell) in cells_in(sheet, extent, a) {
                                if is_formula(&cell.content) {
                                    deps.push((a.sheet, row, col));
                                }
                            }
                        }
                    }
                    Read::Name { name, sheet } => {
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
                }),
                // A text this version does not parse, here or at this
                // offset: its edges as the text scan reads them, so that a
                // cycle through it is still found.
                _ => {
                    let moved = refers::shift_a1_references(&shape.text, p.d_row, p.d_col);
                    dependencies(model, names, extents, id.0, &moved, &mut deps);
                }
            }
            let edges = &mut adj[vi];
            for d in &deps {
                if let Some(&di) = index.get(d) {
                    edges.push(di);
                }
            }
            edges.sort_unstable();
            edges.dedup();
        }
        let components = graph::tarjan(&adj);
        let shape_at = self.order_and_runs(&ids, &adj, &shape_of, &components);
        let places = &self.places;
        let formula = |id: CellId| places.contains_key(&id);
        self.reach = Reach::of(model, &self.shapes, &self.order, shape_at, &formula);
    }

    /// The order, the cycles and the runs, from the cells, each one's edges
    /// and shape, and the components, which come each after everything it
    /// reads; the shape of each cell of the order answered.
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
        adj: &[Vec<usize>],
        shape_of: &[usize],
        components: &[Vec<usize>],
    ) -> Vec<u32> {
        self.order.clear();
        self.cycles.clear();
        self.runs.clear();
        let mut component = vec![0; ids.len()];
        let mut level = vec![0; ids.len()];
        let mut alone = vec![false; ids.len()];
        for (c, comp) in components.iter().enumerate() {
            for &v in comp {
                component[v] = c;
            }
            let deepest = comp
                .iter()
                .flat_map(|&v| &adj[v])
                .filter(|&&u| component[u] != c)
                .map(|&u| level[u] + 1)
                .max()
                .unwrap_or(0);
            for &v in comp {
                level[v] = deepest;
            }
            if comp.len() == 1 && !adj[comp[0]].contains(&comp[0]) {
                alone[comp[0]] = true;
            } else {
                let mut cells: Vec<CellId> = comp.iter().map(|&i| ids[i]).collect();
                cells.sort_unstable();
                self.cycles.push(cells);
            }
        }
        self.cycles.sort();
        // By shape and then by level, each pass keeping the order before it,
        // so by level, shape and cell, a cell's sheet first among its own.
        let cells: Vec<usize> = (0..ids.len()).filter(|&v| alone[v]).collect();
        let cells = by_key(&cells, self.shapes.len(), |v| shape_of[v]);
        let levels = level.iter().max().map_or(0, |deepest| deepest + 1);
        let order = by_key(&cells, levels, |v| level[v]);
        // The run each cell of the order joined, by the run's first position:
        // a cell joins the run before it when it is of the run's shape and
        // sheet and reads no cell of the run, so that the run's cells may be
        // computed in any order, all at once among them (KERNEL.20).
        let mut joined = vec![usize::MAX; ids.len()];
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
                && adj[v].iter().all(|&u| joined[u] != run.start);
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

    /// The run a position of the order is in, as its shape and the position
    /// after its last cell; `None` for a cell evaluated alone.
    pub(crate) fn run_at(&self, pos: usize) -> Option<(usize, usize)> {
        let after = self.runs.partition_point(|r| r.start <= pos);
        let run = self.runs.get(after.checked_sub(1)?)?;
        (pos < run.end).then_some((run.shape, run.end))
    }

    /// Cells of one run, a part of it or the whole, evaluated at once
    /// against the grid as `env` shows it (KERNEL.20, `super::vector`), each
    /// cell's answer pushed to `out` in order: what [`Shapes::evaluate`]
    /// answers for each, the environment told each cell as it is computed.
    pub(crate) fn evaluate_run(
        &self,
        env: &dyn Env,
        library: &Library,
        shape: usize,
        cells: &[CellId],
        lanes: &mut Lanes,
        out: &mut Vec<Computed>,
    ) {
        let s = &self.shapes[shape];
        match &s.expr {
            Ok(e) => {
                let held = s.holds;
                vector::evaluate(
                    env, library, e, held, &s.repeats, s.anchor, cells, lanes, out,
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
                match ctx.eval(e).and_then(|arg| ctx.one(&arg)) {
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
                (Ok(_), None) => None,
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
/// for the order of a plan (KERNEL.20).
fn by_key(cells: &[usize], keys: usize, key: impl Fn(usize) -> usize) -> Vec<usize> {
    let mut start = vec![0; keys + 1];
    for &v in cells {
        start[key(v) + 1] += 1;
    }
    for k in 0..keys {
        start[k + 1] += start[k];
    }
    let mut out = vec![0; cells.len()];
    for &v in cells {
        let k = key(v);
        out[start[k]] = v;
        start[k] += 1;
    }
    out
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
            | Expr::Empty => {}
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
