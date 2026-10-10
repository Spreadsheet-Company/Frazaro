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
//! for the life of a grid and read again when a write changes a formula.

use std::collections::HashMap;

use crate::fx::FxHashMap;
use crate::messages::{raise, Refusal};
use crate::refers::{self, quote_sheet};
use crate::sheet::{cell_ref, shared_masters, strip_future_prefixes, Content, Workbook};

use super::eval::{Computed, Ctx, Env, Library, Reason};
use super::formula::{self, Expr, Read};
use super::graph::{self, areas_at, cells_in, dependencies, is_formula, CellId, Extent, Names};

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
    risky: bool,
}

impl Shape {
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
}

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
                s.place_with((si, row, col), &cell.content, &masters, library);
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
        self.place_with(id, &cell.content, &masters, library);
    }

    fn place_with(
        &mut self,
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
                        let s = self.add(strip_future_prefixes(text), anchor, library);
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
                        let s = self.add(text, id, library);
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

    fn add(&mut self, text: String, anchor: CellId, library: &Library) -> usize {
        let expr = formula::parse(&text);
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
        self.shapes.push(Shape {
            text,
            expr,
            anchor,
            missing,
            risky,
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
        let mut deps: Vec<CellId> = Vec::new();
        for (vi, id) in ids.iter().enumerate() {
            deps.clear();
            let p = self.places[id];
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
        self.order.clear();
        self.cycles.clear();
        for comp in components {
            if comp.len() == 1 && !adj[comp[0]].contains(&comp[0]) {
                self.order.push(ids[comp[0]]);
            } else {
                let mut cells: Vec<CellId> = comp.iter().map(|&i| ids[i]).collect();
                cells.sort_unstable();
                self.cycles.push(cells);
            }
        }
        self.cycles.sort();
    }

    /// The formula cells in the order of evaluation, each after everything
    /// it reads; the cells of a cycle are not in it.
    pub fn order(&self) -> &[CellId] {
        &self.order
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
                    raise(
                        "calc-construct-not-read",
                        &[("cell", &spell_cell(model, *id)), ("construct", &label)],
                    ),
                ));
            }
        }
        None
    }
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
