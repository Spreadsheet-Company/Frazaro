//! `frazaro calc` (KERNEL.7, 2026-10-08): a workbook's formulas computed by
//! the language's evaluator under Excel's library, each against the value
//! the host saved into the file.
//!
//! The reader streams the file's rows into a grid, `vla-lang`'s sheet
//! model, the inverse of what `reflect` prints: a `cell` row a constant, a
//! `formula` row the formula with the `cell` row before it kept aside as the
//! host's cached value, a `name` row a defined name. The evaluator
//! (`vla_lang::calc::Calc`) then computes every formula in dependency order,
//! and the door prints one `calc` row per formula cell, sheets in tab order
//! and cells in row-major order:
//!
//! ```text
//! (cycle "Model!A1" "Model!B1")                        every cycle first, its cells
//! (calc "Output" "C2" "=B2+B3" 15 15 agree)            computed, cached, the verdict
//! (calc "Output" "E2" "=VLOOKUP(1,A:B,2)" (not-computed "VLOOKUP") 3 unchecked)
//! (calc "Checks" "A1" "=Model!B3*2" 800 none unchecked)   a file with no cached value
//! ```
//!
//! The verdict is `agree` when the computed value and the cached one are
//! the same under the one stated tolerance (`vla_lang::calc::agrees`:
//! fifteen significant digits for a number, Excel's documented precision,
//! at most one unit of the last apart as each spells them),
//! `differ` when both exist and are not, `unchecked` when the file holds no
//! cached value, the formula was not computed, or the cached value is one
//! this version does not compare. This is the treaty's twelfth oracle: the
//! goldens under `scripts/recalc/` are this door's own output over the
//! fixtures, and `scripts/reflect/saved.xlsx`, saved by Excel 365, is the
//! first whose every row must `agree`. A cycle, an unreadable reference and
//! a function outside the subset are each named in their cell's row and
//! refuse nothing here: a workbook is read as it is, which is what an
//! auditor needs; the engine's `load` is where they refuse
//! (`vla_lang::calc::Calc::first_refusal`).
//!
//! With `--shapes` the door computes nothing and prints the plan an
//! engine's machine makes of the same grid, its census (KERNEL.30,
//! [`census_of`]): a `(plan ...)` row of totals, a `(readers ...)` row a
//! sheet and a `(shape ...)` row a shape, each read named by the path the
//! machine takes for it, so that a model's cost is read before anything
//! computes it.

use std::borrow::Cow;
use std::collections::BTreeMap;

use vla_lang::calc::{agrees, Calc, CellId, Computed, Extent, Library, Names, Reason, Shapes};
use vla_lang::rows::{datum, quoted, ArrayKind, Row, Sink, Value};
use vla_lang::sheet::{cell_ref, parse_a1_range, Cell, Content, RowSettings, Sheet, Workbook};

use super::{open, Discard};
use crate::messages::Refusal;

/// A workbook read into the grid, with the host's cached values beside it.
#[derive(Debug, Default)]
pub struct Collected {
    pub model: Workbook,
    /// The value the file holds for each formula cell, where it holds one.
    pub cached: BTreeMap<CellId, Value>,
}

impl Collected {
    fn place(&mut self, sheet: &str, addr: &str) -> Option<CellId> {
        let i = self.model.ensure_sheet(sheet);
        let r = parse_a1_range(addr)?;
        Some((i, r.top, r.left))
    }
}

impl Sink for Collected {
    fn row(&mut self, row: &Row<'_>) {
        match row {
            Row::Sheet { name, .. } => {
                self.model.ensure_sheet(name);
            }
            Row::Name { name, refers_to } => {
                self.model
                    .defined_names
                    .push((name.to_string(), refers_to.to_string()));
            }
            Row::Table { .. } | Row::Refers { .. } => {}
            Row::Cell { sheet, addr, value } => {
                let Some(id) = self.place(sheet, addr) else {
                    return;
                };
                let content = match value {
                    Value::Number(text) => match text.trim().parse::<f64>() {
                        Ok(n) if n.is_finite() => Content::Number(n),
                        _ => Content::Text(text.clone()),
                    },
                    Value::Text(t) => Content::Text(t.clone()),
                    Value::Bool(b) => Content::Bool(*b),
                    Value::Error(e) => Content::Error(e.clone()),
                    Value::Date(d) => Content::Text(d.clone()),
                };
                self.model.sheets[id.0].set(id.1, id.2, Cell { content, style: 0 });
                self.cached.insert(id, (*value).clone());
            }
            Row::Formula { sheet, addr, text } => {
                let Some(id) = self.place(sheet, addr) else {
                    return;
                };
                // The cell row before this one, if any, was the host's
                // cached value, which stays in `cached` alone.
                let body = text.strip_prefix('=').unwrap_or(text);
                self.model.sheets[id.0].set(
                    id.1,
                    id.2,
                    Cell {
                        content: Content::Formula(body.to_string()),
                        style: 0,
                    },
                );
            }
        }
    }

    /// An array formula's anchor (KERNEL.8, slices b and f): a dynamic-array
    /// formula loaded as one, in which a range where one value is wanted is
    /// read whole as an array and an answer of many is a spill, labelled at
    /// its cell; a legacy array formula, entered with Ctrl+Shift+Enter, and
    /// an OpenDocument matrix formula as a legacy array formula, read as
    /// arrays the same way, its cell holding the first value of an answer of
    /// many, the other cells of its range the values the file gives them.
    fn array(&mut self, sheet: &str, addr: &str, kind: ArrayKind) {
        let Some(id) = self.place(sheet, addr) else {
            return;
        };
        if let Some(cell) = self.model.sheets[id.0].cells.get(&(id.1, id.2)) {
            if let Content::Formula(body) = &cell.content {
                let content = match kind {
                    ArrayKind::Dynamic => Content::DynamicFormula(body.clone()),
                    ArrayKind::Legacy => Content::ArrayFormula(body.clone()),
                };
                self.model.sheets[id.0].set(id.1, id.2, Cell { content, style: 0 });
            }
        }
    }

    /// A hidden row (KERNEL.8, slice e): the model's row settings, as an
    /// engine's load row sets them, which `SUBTOTAL` reads.
    fn hidden_row(&mut self, sheet: &str, row: u32) {
        let i = self.model.ensure_sheet(sheet);
        self.model.sheets[i].rows.insert(
            row,
            RowSettings {
                height: None,
                hidden: true,
            },
        );
    }
}

/// The counts `--counts` prints.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct Counts {
    pub formulas: u64,
    pub computed: u64,
    pub not_computed: u64,
    pub cycles: u64,
    pub agree: u64,
    pub differ: u64,
    pub unchecked: u64,
}

impl Counts {
    /// The counts as one line's words.
    pub fn line(&self) -> String {
        format!(
            "formulas {} computed {} not-computed {} cycles {} agree {} differ {} unchecked {}",
            self.formulas,
            self.computed,
            self.not_computed,
            self.cycles,
            self.agree,
            self.differ,
            self.unchecked
        )
    }
}

/// The file read into the grid: every sheet walked, the constants kept in
/// the model and the cached values beside it; a constant-only workbook has
/// nothing to compute and prints nothing.
pub fn collect(bytes: &[u8], label: &str) -> Result<Collected, Refusal> {
    Ok(collect_dated(bytes, label)?.0)
}

/// [`collect`], and whether the workbook counts its dates from 1904, as a
/// Mac's may (KERNEL.8, slice d), which [`dated`] makes its library of.
pub fn collect_dated(bytes: &[u8], label: &str) -> Result<(Collected, bool), Refusal> {
    let package = open(bytes, label)?;
    let mut collected = Collected::default();
    package.header_rows(&mut collected);
    // Names arrive before cells, in the header; the walk adds the cells.
    for i in 0..package.sheets().len() {
        package.walk_sheet(i, &mut collected)?;
    }
    let _ = Discard;
    Ok((collected, package.date1904()))
}

/// The library a workbook computes under (KERNEL.8, slice d): the one
/// given, or, for a workbook counting its dates from 1904, the same with
/// every function whose answer turns on the calendar labelled, `the 1904
/// date system` (`vla_lang::calc::dates::label_1904`), until a workbook
/// asks for that system computed.
pub fn dated(library: &Library, date1904: bool) -> Cow<'_, Library> {
    if date1904 {
        let mut labelled = library.clone();
        vla_lang::calc::dates::label_1904(&mut labelled);
        Cow::Owned(labelled)
    } else {
        Cow::Borrowed(library)
    }
}

/// A collected workbook's census (KERNEL.30, `frazaro calc --shapes`): the
/// grid planned as an engine's machine plans one, by `Shapes::of` under the
/// library with the grid's names and extents, every sheet read as the
/// frame's own, a workbook holding no `.last` twin, and nothing computed. A
/// cycle, a function the library does not hold and a construct not read
/// are printed as facts of the plan, where a machine's load would refuse
/// the grid. The rows end in a newline each.
pub fn census_of(collected: &Collected, library: &Library) -> String {
    let model = &collected.model;
    let extents: Vec<Extent> = model.sheets.iter().map(Sheet::extent).collect();
    let plan = Shapes::of(
        model,
        model.sheets.len(),
        library,
        &Names::of(model),
        &extents,
    );
    plan.census(model, library)
}

/// The rows and the counts of one recalculation of a collected workbook.
pub fn rows_of(collected: &Collected, library: &Library) -> (String, Counts) {
    let calc = Calc::run(&collected.model, library);
    let mut out = String::new();
    let mut counts = Counts::default();
    for cycle in calc.cycles() {
        counts.cycles += 1;
        let cells: Vec<String> = cycle
            .iter()
            .map(|id| quoted(&calc.spell_cell(*id)))
            .collect();
        out.push_str(&format!("(cycle {})\n", cells.join(" ")));
    }
    for id in calc.formula_cells() {
        counts.formulas += 1;
        let sheet = &collected.model.sheets[id.0].name;
        let addr = cell_ref(id.1, id.2);
        let formula = format!("={}", calc.formula(id).unwrap_or(""));
        let computed = calc
            .computed(id)
            .cloned()
            .unwrap_or(Computed::NotComputed(Reason::Cycle));
        let cached = collected.cached.get(&id);
        let verdict = match (&computed, cached) {
            (Computed::Value(v), Some(c)) => match agrees(v, c) {
                Some(true) => "agree",
                Some(false) => "differ",
                None => "unchecked",
            },
            _ => "unchecked",
        };
        match &computed {
            Computed::Value(_) => counts.computed += 1,
            Computed::NotComputed(_) => counts.not_computed += 1,
        }
        match verdict {
            "agree" => counts.agree += 1,
            "differ" => counts.differ += 1,
            _ => counts.unchecked += 1,
        }
        out.push_str(&format!(
            "(calc {} {} {} {} {} {})\n",
            quoted(sheet),
            quoted(&addr),
            quoted(&formula),
            computed.spell(),
            cached.map(datum).unwrap_or_else(|| "none".to_string()),
            verdict
        ));
    }
    (out, counts)
}

/// The whole of a workbook's recalculation as text, one row a line in the
/// fixed order: the API's surface and the goldens'.
pub fn calc_text(bytes: &[u8], label: &str, library: &Library) -> Result<String, Refusal> {
    let (collected, date1904) = collect_dated(bytes, label)?;
    Ok(rows_of(&collected, &dated(library, date1904)).0)
}

/// The counts alone.
pub fn calc_counts(bytes: &[u8], label: &str, library: &Library) -> Result<Counts, Refusal> {
    let (collected, date1904) = collect_dated(bytes, label)?;
    Ok(rows_of(&collected, &dated(library, date1904)).1)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::excel;

    const SAVED: &[u8] = include_bytes!("../../../scripts/reflect/saved.xlsx");
    const FIXTURE: &[u8] = include_bytes!("../../../scripts/reflect/fixture.xlsx");
    const BUILT: &[u8] = include_bytes!("../../../scripts/build/fixture_golden.xlsx");

    #[test]
    fn the_file_excel_saved_agrees_on_every_formula() {
        let lib = excel::library();
        let text = calc_text(SAVED, "saved.xlsx", &lib).unwrap();
        let rows: Vec<&str> = text.lines().collect();
        assert_eq!(rows.len(), 5, "{text}");
        for row in &rows {
            assert!(row.ends_with(" agree)"), "{row}");
        }
        assert!(
            rows.contains(&"(calc \"Output\" \"C2\" \"=B2+B3\" 15 15 agree)"),
            "{text}"
        );
        assert!(
            rows.contains(&"(calc \"Output\" \"D2\" \"=IFS(B2>3,\\\"big\\\",TRUE,\\\"small\\\")\" \"big\" \"big\" agree)"),
            "{text}"
        );
        let counts = calc_counts(SAVED, "saved.xlsx", &lib).unwrap();
        assert_eq!(
            counts,
            Counts {
                formulas: 5,
                computed: 5,
                not_computed: 0,
                cycles: 0,
                agree: 5,
                differ: 0,
                unchecked: 0
            }
        );
    }

    /// KERNEL.20's fifth slice: the core's own build holds the value of each
    /// formula it computes, so every row agrees with the value it folded.
    #[test]
    fn the_core_s_own_build_holds_the_values_it_computes_and_agrees() {
        let lib = excel::library();
        let text = calc_text(BUILT, "fixture_golden.xlsx", &lib).unwrap();
        for row in text.lines() {
            assert!(row.ends_with(" agree)"), "{row}");
        }
        assert!(
            text.contains("(calc \"Output\" \"B3\" \"=B2*2\" 10 10 agree)"),
            "{text}"
        );
    }

    #[test]
    fn the_reflect_fixture_names_what_this_slice_does_not_compute() {
        let lib = excel::library();
        let text = calc_text(FIXTURE, "fixture.xlsx", &lib).unwrap();
        let has = |s: &str| assert!(text.contains(s), "missing {s} in\n{text}");
        has("(calc \"Model\" \"C1\" \"=B1*2\" 2400 2400 agree)");
        has("(calc \"Model\" \"D1\" \"=SUM(B:B)\"");
        has("(not-computed \"Sales[Amount]\")");
        // A dynamic formula answering many values, a spill (KERNEL.8, slice f).
        has("(not-computed \"a spill\")");
        has("(not-computed \"INDIRECT\")");
        has("(not-computed \"OFFSET\")");
        has("(not-computed \"[Rates.xlsx]Sheet1!A1\")");
        // A legacy array formula, its first value (KERNEL.8, slice f).
        has("(calc \"Model\" \"E3\" \"=B1:B2*2\" 2400 2400 agree)");
        has("(calc \"Model\" \"E5\" \"=1/0\" (error \"#DIV/0!\") (error \"#DIV/0!\") agree)");
        has("(calc \"Model\" \"B4\" \"=B3*Rate\" 80 80 agree)");
        has("(calc \"Model\" \"G2\" \"=Local\" 800 800 agree)");
        has("(calc \"Model\" \"D4\" \"=SUM(Model:Scratch!B1)\"");
        let counts = calc_counts(FIXTURE, "fixture.xlsx", &lib).unwrap();
        assert_eq!(counts.differ, 0, "{text}");
        assert_eq!(counts.cycles, 0);
        assert!(counts.agree >= 15, "{}", counts.line());
    }

    /// Excel 365's save of the date functions' fixture (KERNEL.8, slice d).
    const DATES_SAVED: &[u8] = include_bytes!("../../../scripts/recalc/dates_saved.xlsx");

    /// KERNEL.8's fourth slice: every day from day 0 to 31 December 2101 is
    /// the day Excel's save makes it. The save's Calendar sheet holds a row
    /// a year from 1900, the year's first day's serial (A), that day's
    /// weekday a hundred times over (C) and the day before's date (D); the
    /// days between follow by the months' fixed lengths, February's the
    /// days between two years' first days less 337, 1900's 29 among them,
    /// and 2101's by its `EOMONTH` (B). Each day's `YEAR`, `MONTH`, `DAY`
    /// and `WEEKDAY`, as the language computes them down a column of every
    /// serial, is held to the day so made.
    #[test]
    fn every_day_from_1900_to_2101_is_the_day_excel_s_save_makes_it() {
        let saved = collect(DATES_SAVED, "dates_saved.xlsx").unwrap();
        let sheet = saved
            .model
            .sheets
            .iter()
            .position(|s| s.name == "calendar")
            .unwrap();
        let cell = |row: u32, col: u32| match saved.cached.get(&(sheet, row, col)) {
            Some(Value::Number(n)) => n.parse::<i64>().unwrap(),
            other => panic!("calendar R{row}C{col}: {other:?}"),
        };
        // Day 0 first, 1900-01-00 as the day before serial 1 (D1), a
        // Saturday; then each year's days, its first the serial Excel's
        // DATE gave it.
        let mut days = vec![(cell(1, 4), 7)];
        for row in 1..=202 {
            let year = 1899 + i64::from(row);
            let first = cell(row, 1);
            assert_eq!(first, days.len() as i64, "{year}'s first day");
            assert_eq!(days[days.len() - 1].0, cell(row, 4), "{year}'s day before");
            let february = if row < 202 {
                cell(row + 1, 1) - first - 337
            } else {
                cell(row, 2)
            };
            let lengths = [31, february, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
            let mut weekday = cell(row, 3) / 100;
            for (month, length) in (1..).zip(lengths) {
                for day in 1..=length {
                    days.push((year * 10_000 + month * 100 + day, weekday));
                    weekday = weekday % 7 + 1;
                }
            }
        }
        assert_eq!(days.len(), 73_781, "day 0 to 31 December 2101");
        let mut wb = Workbook::new();
        let t = wb.ensure_sheet("T");
        let put = |content| Cell { content, style: 0 };
        for s in 0..days.len() {
            let row = s as u32 + 1;
            wb.sheets[t].set(row, 1, put(Content::Number(s as f64)));
            let formula =
                format!("((YEAR(A{row})*100+MONTH(A{row}))*100+DAY(A{row}))*10+WEEKDAY(A{row})");
            wb.sheets[t].set(row, 2, put(Content::Formula(formula)));
        }
        let lib = excel::library();
        let calc = Calc::run(&wb, &lib);
        for (s, (date, weekday)) in days.iter().enumerate() {
            let got = calc.computed((t, s as u32 + 1, 2)).unwrap().spell();
            assert_eq!(got, (date * 10 + weekday).to_string(), "serial {s}");
        }
    }
}
