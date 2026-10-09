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
//! ordinary sheets to it. [`Calc`] keeps its plan of texts and its parse at
//! evaluation, since it evaluates each formula once and an engine every
//! formula every frame.
//!
//! What is not here, by name, each a slice of its own: dates, text
//! functions, lookups, dynamic arrays and spills, implicit intersection
//! and `@`, `LET` and `LAMBDA` (`KERNEL.8`); speed, the vectorized shared
//! formula, incremental recomputation and content-addressed evaluation
//! (`KERNEL.20`). A formula that reaches any of them is not computed, and
//! says so naming the function or the construct ([`Reason`]), which is the
//! honest static label of `HORIZON.md` section 12.6; a cell that reads such
//! a cell inherits the reason.

pub mod eval;
pub mod formula;
pub mod graph;
pub mod library;
pub mod shape;
pub mod value;

use std::cell::Cell as Slot;
use std::collections::HashMap;

use crate::messages::{raise, Refusal};
use crate::refers::quote_sheet;
use crate::sheet::{cell_ref, formula_text, shared_masters, Content, Workbook};

pub use eval::{Arg, Computed, Ctx, Env, Form, Library, Reason, Strict};
pub use formula::{calls, parse, Expr};
pub use graph::{Area, CellId, Extent, Names, Plan};
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
}

impl<'m> Calc<'m> {
    /// The plan over the model: every formula cell's text read off the
    /// grid (a shared formula's child as its master's text moved), the
    /// graph built from `refers`, the order found and every cycle named.
    /// Nothing is evaluated yet.
    pub fn new(model: &'m Workbook, library: &'m Library) -> Calc<'m> {
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
        let plan = graph::plan(model, &names, &extents, &texts);
        let mut values = HashMap::with_capacity(texts.len());
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

    /// A formula cell's text without its `=`, or `None` for a value cell.
    pub fn formula(&self, id: CellId) -> Option<&str> {
        self.texts.get(&id).map(String::as_str)
    }

    /// Every formula cell, in sheet then row-major order.
    pub fn formula_cells(&self) -> Vec<CellId> {
        let mut ids: Vec<CellId> = self.texts.keys().copied().collect();
        ids.sort_unstable();
        ids
    }

    /// What a formula cell holds, once evaluated; `None` for a value cell
    /// or a cell the steps have not reached.
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
            let computed = self.evaluate(id);
            self.values.insert(id, computed);
        }
        self.next = end;
        self.progress()
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

    /// Start the order again, every value forgotten but the cycles'.
    pub fn restart(&mut self) {
        self.values
            .retain(|_, v| matches!(v, Computed::NotComputed(Reason::Cycle)));
        self.next = 0;
    }

    fn evaluate(&self, id: CellId) -> Computed {
        let Some(text) = self.texts.get(&id) else {
            return Computed::NotComputed(Reason::Construct(String::new()));
        };
        self.here.set(id);
        let parsed = match formula::parse(text) {
            Ok(e) => e,
            Err(construct) => return Computed::NotComputed(Reason::Construct(construct)),
        };
        let ctx = Ctx::new(self, self.library);
        match ctx.eval(&parsed).and_then(|arg| ctx.one(&arg)) {
            Ok(v) => Computed::Value(v.settled()),
            Err(reason) => Computed::NotComputed(reason),
        }
    }

    /// The functions the formulas call that the library does not hold,
    /// with the first cell calling each, in sheet then row-major order:
    /// what the engine's `load` refuses by name before the first step.
    pub fn missing_functions(&self) -> Vec<(CellId, String)> {
        let mut seen: HashMap<String, CellId> = HashMap::new();
        for id in self.formula_cells() {
            for name in formula::calls(&self.texts[&id]) {
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
        let (id, construct) = self.formula_cells().into_iter().find_map(|id| {
            formula::parse(&self.texts[&id])
                .err()
                .map(|construct| (id, construct))
        })?;
        Some(raise(
            "calc-construct-not-read",
            &[("cell", &self.spell_cell(id)), ("construct", &construct)],
        ))
    }
}

/// What a projection asks of a recalculation (KERNEL.22): a formula cell's
/// computed value, or `None` for a value cell or a formula with no value
/// yet. [`Calc`] answers it, and so does an engine's machine, whose values
/// outlive any one recalculation; the view record prints its `value` rows
/// from either.
pub trait Values: std::fmt::Debug {
    fn computed(&self, id: CellId) -> Option<&Computed>;
}

impl Values for Calc<'_> {
    fn computed(&self, id: CellId) -> Option<&Computed> {
        Calc::computed(self, id)
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
    fn model(&self) -> &Workbook {
        self.model
    }

    fn here(&self) -> CellId {
        self.here.get()
    }

    fn read(&self, id: CellId) -> Result<Value, Reason> {
        if self.texts.contains_key(&id) {
            return match self.values.get(&id) {
                Some(Computed::Value(v)) => Ok(v.clone()),
                Some(Computed::NotComputed(r)) => Err(r.clone()),
                // Every dependency precedes its reader in the order; a
                // formula with no value yet can only be read from a cycle.
                None => Err(Reason::Cycle),
            };
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
            ("S", "B1", "=VLOOKUP(1,A1:A2,2)"),
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
            Computed::NotComputed(Reason::Function("VLOOKUP".to_string()))
        );
        assert_eq!(
            at(&calc, "S", "B2"),
            Computed::NotComputed(Reason::Function("VLOOKUP".to_string())),
            "a reader inherits the reason"
        );
        assert_eq!(
            at(&calc, "S", "B3"),
            Computed::NotComputed(Reason::Construct("intersection".to_string()))
        );
        assert_eq!(
            at(&calc, "S", "B4"),
            Computed::NotComputed(Reason::Construct("{".to_string()))
        );
        assert_eq!(
            at(&calc, "S", "B5"),
            Computed::NotComputed(Reason::Construct("A1:A2".to_string())),
            "a range where one value is wanted waits for KERNEL.8"
        );
        assert_eq!(
            calc.missing_functions(),
            vec![((0, 1, 2), "VLOOKUP".to_string())]
        );
        let refusal = calc.first_refusal().unwrap();
        assert_eq!(refusal.id, "calc-function-not-computed");
        assert!(
            refusal.text.contains("VLOOKUP") && refusal.text.contains("S!B1"),
            "{}",
            refusal.text
        );
        assert_eq!(at(&calc, "S", "A1").spell(), "\"yes\"");
        assert_eq!(at(&calc, "S", "B1").spell(), "(not-computed \"VLOOKUP\")");
    }

    #[test]
    fn a_construct_not_read_is_the_third_thing_a_load_refuses() {
        // KERNEL.22: after the cycle and the function, a formula whose text
        // this version does not read, naming the construct.
        let wb = model(&[
            ("S", "A1", "=1+1"),
            ("S", "B2", "={1,2}"),
            ("S", "C3", "=@A1"),
        ]);
        let lib = library::language();
        let calc = Calc::new(&wb, &lib);
        let r = calc.first_refusal().unwrap();
        assert_eq!(r.id, "calc-construct-not-read");
        assert!(
            r.text.contains("S!B2") && r.text.contains("reaches {"),
            "{}",
            r.text
        );
        // A function the library does not hold comes first.
        let wb = model(&[("S", "B2", "={1,2}"), ("S", "C3", "=NOW()")]);
        let calc = Calc::new(&wb, &lib);
        assert_eq!(
            calc.first_refusal().unwrap().id,
            "calc-function-not-computed"
        );
        let wb = model(&[("S", "A1", "=1+1")]);
        assert!(Calc::new(&wb, &lib).first_refusal().is_none());
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
}
