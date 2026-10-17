//! The census (KERNEL.30, 2026-10-12): a plan of shapes printed as rows, so
//! that what a grid costs a frame is read before a frame runs.
//!
//! An engine's machine spends a frame by two properties of each shape it
//! plans (`super::shape`): how many of its cells stand at one level of the
//! order, since the cells of a shape at one level are evaluated at once as a
//! run (`super::vector`), and what each of its reads costs, which is the
//! path the machine takes for it and not the way it is written. The plan
//! held both and showed neither: a cartridge written one fill a row was 398
//! shapes where one fill a formula was four, and its frame's start 369 ms of
//! a 387 ms frame (Alonzo's `CART.2`), which these rows say before a frame
//! runs. The rows, in the record's notation, one fact a keyed form:
//!
//! - `(plan (shapes n) (texts n) (formulas n) (levels n) (runs n) (lanes n)
//!   (alone n) (every-frame n) (cycles n))`: the shapes a cell holds, their
//!   distinct texts in R1C1, the formula cells, the levels of the order that
//!   hold a cell, the runs and the cells in them, the cells evaluated one at
//!   a time, the cells of the shapes computed every frame, and the cells on
//!   cycles, which only a plan a load would refuse holds;
//! - `(readers "<sheet>" n)`, for each sheet a change to which is followed:
//!   its readers, each shape's references to it, which a changed cell of it
//!   is checked against at a frame's start and at each cell a frame changes
//!   (`super::reach`); `(readers "<sheet>" n (most k))` where the sheet's
//!   readers are indexed by where a changed cell stands (KERNEL.35), `k` the
//!   most of them any one changed cell of it is checked against;
//! - `(shape k "<sheet>" "<first cell>" "=<R1C1>" (cells n) (levels first
//!   last count widest) (runs count lanes narrowest widest) (alone n)
//!   (reads ...))`, for each shape a cell holds, by the plan's own number,
//!   then, only where they hold, `(calls "NAME" ...)`, the functions called
//!   at each cell; `(missing "NAME" ...)`, those the library does not hold;
//!   `(sheets n)`, a shape on more than one sheet; `(same k)`, the first
//!   shape of the same text; `(every-frame name|span|broken|unread)`, a shape
//!   no change is followed to, and why; `(unlowered <why>)`, a shape with
//!   cells alone whose parse a program does not hold, so that `Ctx` walks it
//!   at each of them (KERNEL.44, `super::program::Why`'s words: `text`,
//!   `error`, `name`, `range`, `form`, `kinds`, `span`, `sheet`, `omitted`,
//!   `deep`).
//!
//! The kinds of read, each the cost class of the path a run takes:
//! - `slot`: one cell of a sheet kept as an array, a `.last` twin dense
//!   enough, read by its index;
//! - `array`: a range of a sheet kept as an array, walked by index, by the
//!   run's fold or by the library's walk at each cell;
//! - `lookup`: one cell through the machine's lookups, a formula's value of
//!   the frame by its position, a value cell or a twin too sparse for an
//!   array by its cells;
//! - `cells`: a range walked through its cells at each cell, by the
//!   library: a range of a frame's sheet, or of a twin too sparse;
//! - `once`: a cell every `$` holds, read once a run, a change to which
//!   reaches every cell of the shape;
//! - `start`: a part the same at every cell, reading value cells and twins
//!   alone, computed once at a frame's start, which reaches the shape only
//!   when its bits change;
//! - `each`: a node the evaluator computes at each cell, everything under
//!   it with it: a defined name, a form other than `IF` and `CHOOSE`, a range
//!   where one value is wanted, a call one of whose operands may answer a
//!   reference.
//!
//! A shape's kinds are the paths its run takes, read off its parse by a walk
//! beside the run's own dispatch (`vector::paths`), and held to it by a test
//! that notes the paths runs take; a cell evaluated alone, which `(alone n)`
//! counts, runs its shape's program, its reads of the same classes, or where
//! `(unlowered <why>)` says the shape has none, the evaluator's path for
//! every node (KERNEL.44). The facts are
//! computed when asked, from the plan as it stands, which keeps nothing for
//! them but where each level of its order ends; they are facts alone, no
//! flag and no threshold, the judgment being the reader's (Alonzo's
//! `SPEC.md` section 2, what a rule costs).

use std::collections::{BTreeSet, HashMap};

use crate::refers;
use crate::rows::quoted;
use crate::sheet::{cell_ref, Workbook};

use super::eval::Library;
use super::formula::Read;
use super::program;
use super::shape::Shapes;
use super::vector;

/// A kind of read, the cost class of the path the machine takes for it, in
/// the order a row prints them.
#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord)]
pub(crate) enum Kind {
    Stride,
    Slot,
    Array,
    Lookup,
    Cells,
    Once,
    Start,
    Each,
}

impl Kind {
    /// The kind as a row spells it.
    pub(crate) fn word(self) -> &'static str {
        match self {
            Kind::Stride => "stride",
            Kind::Slot => "slot",
            Kind::Array => "array",
            Kind::Lookup => "lookup",
            Kind::Cells => "cells",
            Kind::Once => "once",
            Kind::Start => "start",
            Kind::Each => "each",
        }
    }
}

/// Why no change is followed to a shape, so that its cells are computed
/// every frame (`super::reach`).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum Why {
    /// It reads a defined name, whose text may read its cell.
    Name,
    /// It reads a span of sheets.
    Span,
    /// It holds a broken reference with a cell after it.
    Broken,
    /// Its text is one this version does not read.
    Unread,
}

impl Why {
    fn word(self) -> &'static str {
        match self {
            Why::Name => "name",
            Why::Span => "span",
            Why::Broken => "broken",
            Why::Unread => "unread",
        }
    }
}

/// One shape's facts, a row of the census.
#[derive(Clone, Debug, Default, PartialEq)]
pub(crate) struct Facts {
    /// The plan's number for the shape.
    pub(crate) shape: usize,
    /// Its first cell's sheet and address, and its text in R1C1 there,
    /// with its `=`.
    pub(crate) sheet: String,
    pub(crate) first: String,
    pub(crate) r1c1: String,
    /// The formula cells that hold it, and the sheets they stand on.
    pub(crate) cells: usize,
    pub(crate) sheets: usize,
    /// Its cells in the order, and the levels they stand at: the first, the
    /// last, how many, and the most cells at one; `None` with no cell there.
    pub(crate) in_order: usize,
    pub(crate) levels: Option<(usize, usize, usize, usize)>,
    /// Its runs: how many, their cells, the fewest and the most in one.
    pub(crate) runs: usize,
    pub(crate) lanes: usize,
    pub(crate) narrowest: usize,
    pub(crate) widest: usize,
    pub(crate) reads: BTreeSet<Kind>,
    /// The functions called at each cell, and those the library does not
    /// hold, upper case, sorted, each once.
    pub(crate) calls: Vec<String>,
    pub(crate) missing: Vec<String>,
    /// The first shape of the same text, when this is not it.
    pub(crate) same: Option<usize>,
    pub(crate) every_frame: Option<Why>,
    /// Why its cells alone are left to `Ctx`, where it has cells alone, no
    /// program, no function the library lacks and no reason to be computed
    /// every frame, each of which says why already (KERNEL.44).
    pub(crate) unlowered: Option<program::Why>,
}

impl Facts {
    /// Its cells of the order evaluated one at a time.
    pub(crate) fn alone(&self) -> usize {
        self.in_order - self.lanes
    }
}

/// A plan's census: its shapes' facts, the readers of each sheet and, where
/// they are indexed, the most a changed cell of it is checked against, and
/// what the plan alone knows.
#[derive(Clone, Debug, Default, PartialEq)]
pub(crate) struct Census {
    pub(crate) shapes: Vec<Facts>,
    pub(crate) readers: Vec<(String, usize, Option<usize>)>,
    pub(crate) formulas: usize,
    /// The levels of the order that hold a cell.
    pub(crate) levels: usize,
    /// The cells on cycles.
    pub(crate) cycles: usize,
}

/// A plan's census over the grid it was made in: `library` the one the plan
/// placed its shapes under, `arrays` whether the environment keeps a sheet
/// as an array (a machine's twins, as its last frame's end left them).
pub(crate) fn of(
    plan: &Shapes,
    model: &Workbook,
    library: &Library,
    arrays: &dyn Fn(usize) -> bool,
) -> Census {
    let all = plan.every_shape();
    let n = all.len();
    let reach = plan.reach();
    let mut f: Vec<Facts> = (0..n)
        .map(|shape| Facts {
            shape,
            ..Facts::default()
        })
        .collect();
    let mut on: Vec<Vec<usize>> = vec![Vec::new(); n];
    for (id, p) in plan.places() {
        f[p.shape].cells += 1;
        if !on[p.shape].contains(&id.0) {
            on[p.shape].push(id.0);
        }
    }
    // The levels, one pass over the order, a stretch of one level and one
    // shape at a time, the order being by level, then shape, then cell.
    let order = plan.order();
    let ends = plan.level_ends();
    let (mut k, mut p, mut levels) = (0usize, 0usize, 0usize);
    let mut counted: Option<usize> = None;
    while p < order.len() {
        while k < ends.len() && ends[k] as usize <= p {
            k += 1;
        }
        let end = ends.get(k).map_or(order.len(), |&e| e as usize);
        let s = reach.shape_at(p);
        let mut q = p + 1;
        while q < end && reach.shape_at(q) == s {
            q += 1;
        }
        let wide = q - p;
        let x = &mut f[s];
        x.in_order += wide;
        x.levels = Some(match x.levels {
            None => (k, k, 1, wide),
            Some((first, _, count, most)) => (first, k, count + 1, most.max(wide)),
        });
        if counted != Some(k) {
            counted = Some(k);
            levels += 1;
        }
        p = q;
    }
    for (start, end, s) in plan.run_spans() {
        let x = &mut f[s];
        let wide = end - start;
        x.narrowest = if x.runs == 0 {
            wide
        } else {
            x.narrowest.min(wide)
        };
        x.runs += 1;
        x.lanes += wide;
        x.widest = x.widest.max(wide);
    }
    let mut unfollowed = vec![false; n];
    for &s in reach.unfollowed() {
        unfollowed[s] = true;
    }
    let mut first_of: HashMap<String, usize> = HashMap::new();
    let mut shapes = Vec::new();
    for (s, mut x) in f.into_iter().enumerate() {
        if x.cells == 0 {
            continue;
        }
        let shape = &all[s];
        let (si, row, col) = shape.anchor;
        x.sheet = model
            .sheets
            .get(si)
            .map_or_else(String::new, |sh| sh.name.clone());
        x.first = cell_ref(row, col);
        x.r1c1 = format!("={}", refers::r1c1(&shape.text, row, col));
        x.sheets = on[s].len();
        x.same = match first_of.get(&x.r1c1) {
            Some(&t) => Some(t),
            None => {
                first_of.insert(x.r1c1.clone(), s);
                None
            }
        };
        if let Ok(e) = &shape.expr {
            let paths = vector::paths(library, model, e, arrays, x.runs > 0);
            for (taken, kind) in [
                (paths.stride, Kind::Stride),
                (paths.slot, Kind::Slot),
                (paths.array, Kind::Array),
                (paths.lookup, Kind::Lookup),
                (paths.cells, Kind::Cells),
                (paths.once, Kind::Once),
                (!reach.parts(s).is_empty(), Kind::Start),
                (paths.each, Kind::Each),
            ] {
                if taken {
                    x.reads.insert(kind);
                }
            }
            x.calls = paths.calls;
        }
        let mut missing: Vec<String> = shape
            .missing()
            .iter()
            .map(|m| m.to_ascii_uppercase())
            .collect();
        missing.sort_unstable();
        missing.dedup();
        x.missing = missing;
        if unfollowed[s] {
            x.every_frame = Some(why(shape));
        } else if x.alone() > 0 && x.missing.is_empty() {
            x.unlowered = shape.unlowered();
        }
        shapes.push(x);
    }
    let readers = model
        .sheets
        .iter()
        .enumerate()
        .filter_map(|(si, sheet)| {
            let n = reach.readers_on(si);
            (n > 0).then(|| (sheet.name.clone(), n, reach.most_on(si)))
        })
        .collect();
    Census {
        shapes,
        readers,
        formulas: plan.formulas(),
        levels,
        cycles: plan.cycles().iter().map(Vec::len).sum(),
    }
}

/// Why no change is followed to a shape the reach computes every frame: a
/// text not read, a risky one (a span of sheets, or a broken reference with
/// a cell after it), or else a defined name it reads.
fn why(shape: &super::shape::Shape) -> Why {
    let Ok(e) = &shape.expr else {
        return Why::Unread;
    };
    if shape.risky {
        let span = refers::scan(&shape.text)
            .iter()
            .any(|r| r.kind == refers::Kind::ThreeD);
        return if span { Why::Span } else { Why::Broken };
    }
    let mut name = false;
    e.each_read(&mut |r| {
        if let Read::Name { .. } = r {
            name = true;
        }
    });
    if name {
        Why::Name
    } else {
        Why::Span
    }
}

/// The census as rows, each a line.
pub(crate) fn spell(c: &Census) -> String {
    let texts = c.shapes.iter().filter(|x| x.same.is_none()).count();
    let runs: usize = c.shapes.iter().map(|x| x.runs).sum();
    let lanes: usize = c.shapes.iter().map(|x| x.lanes).sum();
    let alone: usize = c.shapes.iter().map(Facts::alone).sum();
    let every: usize = c
        .shapes
        .iter()
        .filter(|x| x.every_frame.is_some())
        .map(|x| x.cells)
        .sum();
    let mut out = format!(
        "(plan (shapes {}) (texts {texts}) (formulas {}) (levels {}) (runs {runs}) \
         (lanes {lanes}) (alone {alone}) (every-frame {every}) (cycles {}))\n",
        c.shapes.len(),
        c.formulas,
        c.levels,
        c.cycles
    );
    for (sheet, n, most) in &c.readers {
        match most {
            Some(k) => out.push_str(&format!("(readers {} {n} (most {k}))\n", quoted(sheet))),
            None => out.push_str(&format!("(readers {} {n})\n", quoted(sheet))),
        }
    }
    for x in &c.shapes {
        out.push_str(&format!(
            "(shape {} {} {} {} (cells {})",
            x.shape,
            quoted(&x.sheet),
            quoted(&x.first),
            quoted(&x.r1c1),
            x.cells
        ));
        match x.levels {
            Some((first, last, count, most)) => {
                out.push_str(&format!(" (levels {first} {last} {count} {most})"))
            }
            None => out.push_str(" (levels)"),
        }
        out.push_str(&format!(
            " (runs {} {} {} {}) (alone {}) (reads",
            x.runs,
            x.lanes,
            x.narrowest,
            x.widest,
            x.alone()
        ));
        for kind in &x.reads {
            out.push(' ');
            out.push_str(kind.word());
        }
        out.push(')');
        for (head, names) in [("calls", &x.calls), ("missing", &x.missing)] {
            if names.is_empty() {
                continue;
            }
            out.push_str(&format!(" ({head}"));
            for name in names {
                out.push(' ');
                out.push_str(&quoted(name));
            }
            out.push(')');
        }
        if x.sheets > 1 {
            out.push_str(&format!(" (sheets {})", x.sheets));
        }
        if let Some(k) = x.same {
            out.push_str(&format!(" (same {k})"));
        }
        if let Some(why) = x.every_frame {
            out.push_str(&format!(" (every-frame {})", why.word()));
        }
        if let Some(why) = x.unlowered {
            out.push_str(&format!(" (unlowered {})", why.word()));
        }
        out.push_str(")\n");
    }
    out
}

/// A plan's census as rows: [`of`], then [`spell`].
pub(crate) fn rows(
    plan: &Shapes,
    model: &Workbook,
    library: &Library,
    arrays: &dyn Fn(usize) -> bool,
) -> String {
    spell(&of(plan, model, library, arrays))
}

#[cfg(test)]
mod tests {
    //! The census of grids whose rows are stated whole, each fact derived by
    //! hand from the grid; a plan a load would refuse, through the door's
    //! route; a twin dropped from its array; a shape written over; the
    //! census pure of the frames. The laws over generated grids, the census
    //! held to the plan walked again and to the paths runs take, are the
    //! machine's tests.
    use super::*;
    use crate::calc::graph::Names;
    use crate::calc::library;
    use crate::calc::{Arg, Ctx, Extent, Reason, Value};
    use crate::machine::{read_rows, Machine};
    use crate::reader::read_forms;
    use crate::sheet::Sheet;

    fn machine(rows: &str, lib: Library) -> Machine {
        Machine::load(rows, lib).unwrap_or_else(|e| panic!("line {}: {}", e.line, e.refusal.text))
    }

    /// A strict function no run knows, standing for Excel's `ROUND`.
    fn zero(_: &Ctx<'_>, _: &[Arg<'_>]) -> Result<Value, Reason> {
        Ok(Value::Number(0.0))
    }

    /// The language's library with `ROUND` beside it.
    fn with_round() -> Library {
        let mut lib = library::language();
        lib.register("ROUND", zero);
        lib
    }

    /// The census of a grid's rows planned as a door plans a workbook: the
    /// rows read into a grid, no twin made, the plan made by `Shapes::of`.
    fn door(rows: &str, lib: &Library) -> String {
        let forms = read_forms(rows).expect("reads");
        let draft = read_rows(&forms).unwrap_or_else(|e| panic!("{}", e.refusal.text));
        let wb = &draft.workbook;
        let extents: Vec<Extent> = wb.sheets.iter().map(Sheet::extent).collect();
        let plan = Shapes::of(wb, wb.sheets.len(), lib, &Names::of(wb), &extents);
        plan.census(wb, lib)
    }

    #[test]
    fn a_column_s_census_is_its_rows_whole() {
        // Three columns of twenty, each reading its own row: three shapes at
        // three levels, each a run of twenty; the middle one reads its row
        // and a held cell, and the readers of the sheet are the two moving
        // references and the held one.
        let m = machine(
            concat!(
                "(sheet \"Output\" visible)\n",
                "(cell \"Output\" \"E1\" 3)\n",
                "(formula \"Output\" \"A1:A20\" \"=ROW()\")\n",
                "(formula \"Output\" \"B1:B20\" \"=A1*$E$1+1\")\n",
                "(formula \"Output\" \"C1:C20\" \"=IF(B1>30000,B1-30000,B1)\")\n",
            ),
            library::language(),
        );
        assert_eq!(
            m.census(),
            concat!(
                "(plan (shapes 3) (texts 3) (formulas 60) (levels 3) (runs 3) (lanes 60) (alone 0) (every-frame 0) (cycles 0))\n",
                "(readers \"Output\" 3)\n",
                "(shape 0 \"Output\" \"A1\" \"=ROW()\" (cells 20) (levels 0 0 1 20) (runs 1 20 20 20) (alone 0) (reads))\n",
                "(shape 1 \"Output\" \"B1\" \"=RC[-1]*R1C5+1\" (cells 20) (levels 1 1 1 20) (runs 1 20 20 20) (alone 0) (reads lookup once))\n",
                "(shape 2 \"Output\" \"C1\" \"=IF(RC[-1]>30000,RC[-1]-30000,RC[-1])\" (cells 20) (levels 2 2 1 20) (runs 1 20 20 20) (alone 0) (reads lookup))\n",
            )
        );
    }

    #[test]
    fn a_running_balance_s_census_is_one_cell_a_level() {
        // Each balance reads the one above it: nineteen levels of one cell,
        // every one alone; the entries beside them one run at the first
        // level, with the first balance, a shape of its own.
        let m = machine(
            concat!(
                "(sheet \"Output\" visible)\n",
                "(cell \"Output\" \"A1\" 1)\n",
                "(formula \"Output\" \"A2:A20\" \"=MOD(ROW()*7,13)\")\n",
                "(formula \"Output\" \"B1\" \"=A1\")\n",
                "(formula \"Output\" \"B2:B20\" \"=B1+A2\")\n",
            ),
            library::language(),
        );
        assert_eq!(
            m.census(),
            concat!(
                "(plan (shapes 3) (texts 3) (formulas 39) (levels 20) (runs 1) (lanes 19) (alone 20) (every-frame 0) (cycles 0))\n",
                "(readers \"Output\" 3)\n",
                "(shape 0 \"Output\" \"B1\" \"=RC[-1]\" (cells 1) (levels 0 0 1 1) (runs 0 0 0 0) (alone 1) (reads lookup))\n",
                "(shape 1 \"Output\" \"A2\" \"=MOD(ROW()*7,13)\" (cells 19) (levels 0 0 1 19) (runs 1 19 19 19) (alone 0) (reads))\n",
                "(shape 2 \"Output\" \"B2\" \"=R[-1]C+RC[-1]\" (cells 19) (levels 1 19 19 1) (runs 0 0 0 0) (alone 19) (reads lookup))\n",
            )
        );
    }

    #[test]
    fn a_range_of_the_frame_s_own_sheet_is_walked_through_its_cells() {
        // Three months of twelve over a held assumption, each row summed and
        // rounded, each column totalled: the row sums read a range of their
        // own sheet, which the library walks at each cell; `ROUND`, which no
        // run knows, is called at each cell over what the run computed; the
        // totals stand at three levels, a run of twelve and two cells alone,
        // and the three rows' sums and roundings, three cells a level, are
        // each too few for a run.
        let m = machine(
            concat!(
                "(sheet \"Output\" visible)\n",
                "(cell \"Output\" \"P1\" 100)\n",
                "(formula \"Output\" \"B2:M4\" \"=MOD(ROW()*31+COLUMN()*17,97)+$P$1\")\n",
                "(formula \"Output\" \"N2:N4\" \"=SUM(B2:M2)\")\n",
                "(formula \"Output\" \"O2:O4\" \"=ROUND(N2/12,0)\")\n",
                "(formula \"Output\" \"B5:O5\" \"=SUM(B2:B4)\")\n",
            ),
            with_round(),
        );
        assert_eq!(
            m.census(),
            concat!(
                "(plan (shapes 4) (texts 4) (formulas 56) (levels 4) (runs 2) (lanes 48) (alone 8) (every-frame 0) (cycles 0))\n",
                "(readers \"Output\" 4)\n",
                "(shape 0 \"Output\" \"B2\" \"=MOD(ROW()*31+COLUMN()*17,97)+R1C16\" (cells 36) (levels 0 0 1 36) (runs 1 36 36 36) (alone 0) (reads once))\n",
                "(shape 1 \"Output\" \"N2\" \"=SUM(RC[-12]:RC[-1])\" (cells 3) (levels 1 1 1 3) (runs 0 0 0 0) (alone 3) (reads cells) (calls \"SUM\"))\n",
                "(shape 2 \"Output\" \"O2\" \"=ROUND(RC[-1]/12,0)\" (cells 3) (levels 2 2 1 3) (runs 0 0 0 0) (alone 3) (reads lookup) (calls \"ROUND\"))\n",
                "(shape 3 \"Output\" \"B5\" \"=SUM(R[-3]C:R[-1]C)\" (cells 14) (levels 1 3 3 12) (runs 1 12 12 12) (alone 2) (reads cells) (calls \"SUM\"))\n",
            )
        );
    }

    /// Life's rule over a small board, its first frame a pattern of the
    /// cell's place, its twin read by index and its range walked as an array.
    const LIFE: &str = concat!(
        "(sheet \"Clock\" visible)\n",
        "(cell \"Clock\" \"B1\" 1)\n",
        "(cell \"Clock\" \"B3\" 1)\n",
        "(sheet \"Screen\" visible)\n",
        "(formula \"Screen\" \"B2:F4\" \"=IF(Clock!$B$1=1,MOD(ROW()+COLUMN()+Clock!$B$3,2),",
        "IF(OR(SUM(Screen.last!A1:C3)-Screen.last!B2=3,AND(Screen.last!B2=1,",
        "SUM(Screen.last!A1:C3)-Screen.last!B2=2)),1,0))\")\n",
    );

    /// Life's one shape's row, its reads given and what follows them.
    fn life_row(reads: &str, tail: &str) -> String {
        format!(
            "(shape 0 \"Screen\" \"B2\" \"=IF(Clock!R1C2=1,MOD(ROW()+COLUMN()+Clock!R3C2,2),\
             IF(OR(SUM(Screen.last!R[-1]C[-1]:R[1]C[1])-Screen.last!RC=3,AND(Screen.last!RC=1,\
             SUM(Screen.last!R[-1]C[-1]:R[1]C[1])-Screen.last!RC=2)),1,0))\" (cells 15) \
             (levels 0 0 1 15) (runs 1 15 15 15) (alone 0) (reads {reads}){tail})\n"
        )
    }

    #[test]
    fn life_s_census_reads_its_twin_and_its_range_by_slices() {
        // One shape at one level, one run; the Clock's frame compared at the
        // frame's start and its seed read once a run; the twin's cell and its
        // range, which move with the cell's column, read by slices of the
        // array across a row's cells (KERNEL.43).
        // The Clock's readers are the part and the held seed; the twin's, the
        // range, inside which its cell is.
        let mut m = machine(LIFE, library::language());
        let census = format!(
            "{}{}{}{}",
            "(plan (shapes 1) (texts 1) (formulas 15) (levels 1) (runs 1) (lanes 15) (alone 0) (every-frame 0) (cycles 0))\n",
            "(readers \"Clock\" 2)\n",
            "(readers \"Screen.last\" 1)\n",
            life_row("stride once start", "")
        );
        assert_eq!(m.census(), census);
        // Pure of the frames: the plan and the twins' arrays as they were.
        for _ in 0..3 {
            assert!(m.step(0).done);
        }
        assert_eq!(m.census(), census, "after three frames");
    }

    #[test]
    fn a_twin_too_sparse_for_an_array_is_read_through_its_cells() {
        // A cell written far from the rest widens the sheet past four slots a
        // cell, so its twin's array is dropped at the next frame's end
        // (`KERNEL.20`'s finding, which only a frame's time had shown): the
        // twin's cell is then read through lookups, and its range, which no
        // run walks without the array, by the library at each cell.
        let mut m = machine(LIFE, library::language());
        assert!(m.step(0).done);
        m.write("(cell \"Screen\" \"XFA1\" 1)").expect("writes");
        assert!(m.step(0).done);
        let census = m.census();
        assert!(
            census.ends_with(&life_row("lookup cells once start", " (calls \"SUM\")")),
            "{census}"
        );
    }

    #[test]
    fn the_census_counts_the_shapes_a_cell_holds() {
        // A fill written over with another formula leaves no shape in the
        // plan that holds no cell, and the census numbers the shapes as a
        // load of the same grid would (KERNEL.34; a write kept the old shape
        // and numbered the new one after it).
        let mut m = machine(
            "(sheet \"S\" visible)\n(formula \"S\" \"A1:A8\" \"=ROW()+1\")\n",
            library::language(),
        );
        m.write("(formula \"S\" \"A1:A8\" \"=ROW()*2\")")
            .expect("writes");
        assert_eq!(m.formulas(), (8, 1));
        assert_eq!(
            m.census(),
            concat!(
                "(plan (shapes 1) (texts 1) (formulas 8) (levels 1) (runs 1) (lanes 8) (alone 0) (every-frame 0) (cycles 0))\n",
                "(shape 0 \"S\" \"A1\" \"=ROW()*2\" (cells 8) (levels 0 0 1 8) (runs 1 8 8 8) (alone 0) (reads))\n",
            )
        );
    }

    #[test]
    fn a_door_reads_another_sheet_through_lookups_as_the_machine_does() {
        // A workbook keeps no sheet as an array: a door's plan reads another
        // sheet's cell through lookups and its range through its cells at
        // each cell, the library summing it, as a machine reads a sheet that
        // is not a twin; so a grid that reads no twin is planned alike by
        // the door and the machine. The range holds the cell at every offset,
        // so the sheet's one reader is the range.
        let rows = concat!(
            "(sheet \"S\" visible)\n",
            "(sheet \"T\" visible)\n",
            "(cell \"T\" \"A1:A5\" 1)\n",
            "(formula \"S\" \"A1:A4\" \"=T!A1+SUM(T!A1:A2)\")\n",
        );
        let census = door(rows, &library::language());
        assert_eq!(
            census,
            concat!(
                "(plan (shapes 1) (texts 1) (formulas 4) (levels 1) (runs 1) (lanes 4) (alone 0) (every-frame 0) (cycles 0))\n",
                "(readers \"T\" 1)\n",
                "(shape 0 \"S\" \"A1\" \"=T!RC+SUM(T!RC:R[1]C)\" (cells 4) (levels 0 0 1 4) (runs 1 4 4 4) (alone 0) (reads lookup cells) (calls \"SUM\"))\n",
            )
        );
        assert_eq!(machine(rows, library::language()).census(), census);
    }

    #[test]
    fn a_level_only_a_cycle_stands_at_holds_no_cell_of_the_order() {
        // A cycle reading a formula stands a level above it, and a cell
        // reading the cycle a level above that; the cycle's cells are not in
        // the order, so its level holds none of it, and the plan's levels are
        // the two that hold a cell. The cycle's shapes have no level, no run
        // and no cell alone; no change is followed to a shape not in the
        // order, so it lists no reader.
        let census = door(
            concat!(
                "(sheet \"S\" visible)\n",
                "(formula \"S\" \"X1\" \"=1\")\n",
                "(formula \"S\" \"B1\" \"=X1+C1\")\n",
                "(formula \"S\" \"C1\" \"=B1\")\n",
                "(formula \"S\" \"D1\" \"=B1\")\n",
            ),
            &library::language(),
        );
        assert_eq!(
            census,
            concat!(
                "(plan (shapes 4) (texts 4) (formulas 4) (levels 2) (runs 0) (lanes 0) (alone 2) (every-frame 0) (cycles 2))\n",
                "(readers \"S\" 1)\n",
                "(shape 0 \"S\" \"B1\" \"=RC[22]+RC[1]\" (cells 1) (levels) (runs 0 0 0 0) (alone 0) (reads lookup))\n",
                "(shape 1 \"S\" \"C1\" \"=RC[-1]\" (cells 1) (levels) (runs 0 0 0 0) (alone 0) (reads lookup))\n",
                "(shape 2 \"S\" \"D1\" \"=RC[-2]\" (cells 1) (levels 2 2 1 1) (runs 0 0 0 0) (alone 1) (reads lookup))\n",
                "(shape 3 \"S\" \"X1\" \"=1\" (cells 1) (levels 0 0 1 1) (runs 0 0 0 0) (alone 1) (reads))\n",
            )
        );
    }

    #[test]
    fn a_door_s_plan_holds_what_a_load_refuses() {
        // A workbook a machine would refuse to load, planned as a door plans
        // it: a cycle, a function the library does not hold, a construct the
        // parser does not read, a span of sheets and a defined name, each
        // printed as a fact; two fills of one formula, two shapes of one
        // text; one formula on two sheets, one shape.
        let rows = concat!(
            "(sheet \"Model\" visible)\n",
            "(sheet \"Jan\" visible)\n",
            "(sheet \"Feb\" visible)\n",
            "(cell \"Model\" \"A1:A4\" 2)\n",
            "(cell \"Jan\" \"A1\" 5)\n",
            "(cell \"Feb\" \"A1\" 7)\n",
            "(name \"Rate\" \"Model!$A$1\")\n",
            "(formula \"Model\" \"B1\" \"=C1*3\")\n",
            "(formula \"Model\" \"C1\" \"=B1*3\")\n",
            "(formula \"Model\" \"D1:D2\" \"=A1*2\")\n",
            "(formula \"Model\" \"D3:D4\" \"=A3*2\")\n",
            "(formula \"Model\" \"E1\" \"=Rate*2\")\n",
            "(formula \"Model\" \"F1\" \"=VLOOKUP(1,A1:A4,1)\")\n",
            "(formula \"Model\" \"G1\" \"=A1 A2\")\n",
            "(formula \"Model\" \"H1\" \"=SUM(Jan:Feb!A1)\")\n",
            "(formula \"Jan\" \"B1\" \"=A1+1\")\n",
            "(formula \"Feb\" \"B1\" \"=A1+1\")\n",
        );
        let census = door(rows, &library::language());
        let lines: Vec<&str> = census.lines().collect();
        assert_eq!(
            lines[0],
            "(plan (shapes 9) (texts 8) (formulas 12) (levels 1) (runs 0) (lanes 0) (alone 10) (every-frame 3) (cycles 2))",
            "{census}"
        );
        // The readers: the moving references of the three shapes followed on
        // Model, and the one of the shape on two sheets on each of them.
        assert_eq!(
            &lines[1..4],
            [
                "(readers \"Model\" 3)",
                "(readers \"Jan\" 1)",
                "(readers \"Feb\" 1)"
            ],
            "{census}"
        );
        for row in [
            "(shape 0 \"Model\" \"B1\" \"=RC[1]*3\" (cells 1) (levels) (runs 0 0 0 0) (alone 0) (reads lookup))",
            "(shape 1 \"Model\" \"C1\" \"=RC[-1]*3\" (cells 1) (levels) (runs 0 0 0 0) (alone 0) (reads lookup))",
            "(shape 2 \"Model\" \"D1\" \"=RC[-3]*2\" (cells 2) (levels 0 0 1 2) (runs 0 0 0 0) (alone 2) (reads lookup))",
            "(shape 3 \"Model\" \"E1\" \"=Rate*2\" (cells 1) (levels 0 0 1 1) (runs 0 0 0 0) (alone 1) (reads each) (every-frame name))",
            "(shape 4 \"Model\" \"F1\" \"=VLOOKUP(1,RC[-5]:R[3]C[-5],1)\" (cells 1) (levels 0 0 1 1) (runs 0 0 0 0) (alone 1) (reads) (missing \"VLOOKUP\"))",
            "(shape 5 \"Model\" \"G1\" \"=RC[-6] R[1]C[-6]\" (cells 1) (levels 0 0 1 1) (runs 0 0 0 0) (alone 1) (reads) (every-frame unread))",
            "(shape 6 \"Model\" \"H1\" \"=SUM(Jan:Feb!RC[-7])\" (cells 1) (levels 0 0 1 1) (runs 0 0 0 0) (alone 1) (reads lookup) (calls \"SUM\") (every-frame span))",
            "(shape 7 \"Model\" \"D3\" \"=RC[-3]*2\" (cells 2) (levels 0 0 1 2) (runs 0 0 0 0) (alone 2) (reads lookup) (same 2))",
            "(shape 8 \"Jan\" \"B1\" \"=RC[-1]+1\" (cells 2) (levels 0 0 1 2) (runs 0 0 0 0) (alone 2) (reads lookup) (sheets 2))",
        ] {
            assert!(lines.contains(&row), "no row {row} in\n{census}");
        }
        assert_eq!(lines.len(), 1 + 3 + 9, "{census}");
    }
}
