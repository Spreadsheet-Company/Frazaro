//! The native runner of the recalculation floors (`KERNEL.20`, its sixth
//! slice; `docs/BETA_REARVIEW.md`, the slice's scoping): each fixture program
//! of `scripts/recalc/bench/` built through the core's own surface as `frazaro
//! build` builds it, written to `target/bench/` for Excel's half to open, read
//! back through the reader into rows, and loaded into the machine
//! (`vla_lang::machine`) under Excel's library: the engine `KERNEL.20`'s
//! slices made faster, computing the file Excel computes.
//!
//! ```text
//! cargo run --release -p frazaro-core --example recalc               # every fixture, the baseline rows
//! cargo run --release -p frazaro-core --example recalc -- --check    # nothing timed: the fixtures and their values held
//! cargo run --release -p frazaro-core --example recalc -- --netlist  # the netlist program written by its generator
//! ```
//!
//! The netlist program, `scripts/recalc/bench/netlist.txt`, is the one
//! fixture no person writes: its generator is `bench/netlist.rs`, `--netlist`
//! writes the file, and `--check` holds the file to it, line for line.
//!
//! **What is timed**, each one `step` with no budget, timed alone: a whole
//! frame, every formula computed, as Excel's `CalculateFull` computes every
//! formula (a formula written into the `Frazaro` sheet before each, which no
//! formula reads, makes the next frame whole); and the frame after the
//! fixture's edit, as Excel's `Calculate` after the cell is written. For
//! Life, the edit is the frame: the clock advanced, the previous frame being
//! the machine's twin where Excel copies it into a sheet of its own. And for
//! a fixture whose comments name a type (`KERNEL.56`), a value typed into a
//! fresh cell, one the grid has never held, down a column from the cell the
//! comment names, a cell further down each time, timed with its write and
//! the frames to the grid's stillness, as Excel's half times the same value
//! written into the same cell and its `Calculate`: a frame after a write
//! proves nothing, so the frames go on until one with no write before it
//! changes nothing (`KERNEL.48`). Beside them, printed and not floored: the
//! load, the first frame, and `Calc`, the reference evaluator the doors
//! compute with (`view`, `build`, `calc`).
//!
//! **What is held before any number is printed**, since a fast wrong answer
//! is no floor: every formula's value in the first frame equal to the value
//! the build folded into its cell (`KERNEL.20`'s fifth slice, `Calc`'s own);
//! after the edits, every formula equal to `Calc` over the machine's own
//! grid, and with `--check` after every edit, since alternating values could
//! hide a machine that stopped following them; and every frame of Life equal
//! to a Life written here by hand, the soup of the rule's own arithmetic and
//! then Conway's rule with the border dead (the third Life outside the engine, after Alonzo's runner's and its
//! page's); and every probe of the lookups' fixture and its four answers,
//! after its first frame and its edits, equal to the lookups written here by
//! hand, a map of the table's keys made from the program's own arithmetic
//! (`KERNEL.8`, a third lineage beside `Calc` and the machine, as Life's);
//! and every cell of the criteria fixture's summary, its `SUMIFS` and its
//! `COUNTIFS`, equal to the summary written here by hand, the ledger added
//! up by account and month in an array (`KERNEL.8`, its fifth slice); and
//! every gate of the netlist fixture equal to the adder evaluated by hand,
//! gate by gate in the order they were made (`KERNEL.54`, `bench/netlist.rs`).
//! Each fixture's checksum, the sum of a range its program names,
//! is printed as built and as edited for Excel's half to equal: every value
//! the fixtures compute is a whole number, so the sums are exact.
//!
//! **The clock.** Beside every timed step and every run of `Calc` the
//! hand-written Life runs a fixed workload and is timed too, the processor's
//! clock seen from a computation that never changes (Alonzo's `CART.1`,
//! measured there): when it moves more than 10% across a fixture's steps, or
//! from the run's start, the run is not a baseline and prints no rows.
//! `tools/bench_recalc.ps1` starts this runner on the performance cores, at
//! high priority and opted out of Windows' power throttling, which slows a
//! background process by about 40% three seconds into a run. Each fixture is
//! measured in three rounds, each a fresh machine, and its numbers are the
//! median of the rounds' medians: one process's memory, placed one way, moved
//! the cell-by-cell path's frames by up to a third (measured 2026-10-10).
//!
//! **The flag.** The runner also writes the build golden's fixture three
//! ways, `calcpr_built.xlsx` as built, `calcpr_calcid.xlsx` with Excel 365's
//! `calcId="191029"` in place of `fullCalcOnLoad`, and `calcpr_none.xlsx`
//! with neither, each with `Output!C2`'s folded 15 planted as 999, so that
//! Excel's half can read, before any calculation of its own, whether Excel
//! computed the file on opening it (the fifth slice's fourth decision).
//!
//! The runner is an example: no member of the workspace, nothing in the lock,
//! and out of the crate's package.

use std::collections::BTreeMap;
use std::hint::black_box;
use std::path::{Path, PathBuf};
use std::process::ExitCode;
use std::time::Instant;

use frazaro_core::reflect::calc::collect;
use frazaro_core::sheet::zip;
use frazaro_core::{api, excel, sha256};
use vla_lang::calc::{Calc, CellId, Computed, Value, Values};
use vla_lang::machine::Machine;
use vla_lang::sheet::{parse_a1_range, A1Range, Content};

/// The fixtures in the order the block prints them.
const FIXTURES: [&str; 7] = [
    "column", "chain", "sums", "life", "lookups", "criteria", "netlist",
];

/// The steps timed: whole frames, edits, types, and the reference's runs.
const WHOLE: usize = 11;
const EDITS: usize = 21;
const TYPES: usize = 21;
const CALCS: usize = 3;

/// The most frames a type's grid takes to be still, past which the runner
/// stops: two where nothing reads the typed cell, a frame after the write
/// and one that proves it.
const STILL_WITHIN: usize = 8;

/// The rounds a fixture is measured in, each a fresh machine.
const ROUNDS: usize = 3;

// The clock, the hand-written Life it times, the statistics and the rows the
// machine loads, shared with the memory runner (`KERNEL.49`): `bench/common.rs`.
#[path = "bench/common.rs"]
mod common;
use common::{
    drift, generation, median, range, rows_for_machine, soup, spread, thousands, Probe,
    CLOCK_SPREAD_MAX, H, START_READINGS, W,
};

// The netlist fixture's generator and its hand (`KERNEL.54`).
#[path = "bench/netlist.rs"]
mod netlist;

// ---- the reference: the lookups by hand ------------------------------------

/// The lookups fixture's table, `A2:A20001`'s keys and `B2:B20001`'s values,
/// and its probes beside it in `D2:D20001`, a row each.
const LOOKUP_ROWS: u64 = 20_000;

/// Each probe row's probe and four answers by hand, for the assumption `p1`,
/// as the program's own arithmetic writes them: a key is its row times 7,919
/// modulo the prime 20,011, a value its row times 31 modulo 97, and a probe
/// the key of the row its own row and `p1` name; then the probe found in a
/// map of the keys to their rows, written here and not the engine's, its
/// value (`VLOOKUP`), its position from 1 (`MATCH`), the value there
/// (`INDEX`) and its count, 1, the keys being distinct (`COUNTIF`).
fn lookups_by_hand(p1: u64) -> Vec<[f64; 5]> {
    let key = |r: u64| (r * 7_919) % 20_011;
    let value = |r: u64| (r * 31) % 97;
    let rows: std::collections::HashMap<u64, u64> =
        (2..=LOOKUP_ROWS + 1).map(|r| (key(r), r)).collect();
    (2..=LOOKUP_ROWS + 1)
        .map(|r| {
            let probe = key((r + p1) % LOOKUP_ROWS + 2);
            let found = rows[&probe];
            let v = value(found) as f64;
            [probe as f64, v, (found - 1) as f64, v, 1.0]
        })
        .collect()
}

/// The machine's probes and answers, `D2:H20001`, equal to the lookups by
/// hand for the assumption `p1`.
fn held_lookups(m: &Machine, p1: &str) -> Result<usize, String> {
    let p1: u64 = p1
        .parse()
        .map_err(|_| format!("the assumption {p1} is no whole number"))?;
    let si = m.find_sheet("Output").ok_or("no sheet Output")?;
    for (k, want) in lookups_by_hand(p1).iter().enumerate() {
        let row = k as u32 + 2;
        for (j, w) in want.iter().enumerate() {
            let col = 4 + j as u32;
            match m.computed((si, row, col)) {
                Some(Computed::Value(Value::Number(n))) if n == *w => {}
                got => {
                    return Err(format!(
                        "Output!R{row}C{col}, with P1 {p1}: the machine holds {got:?}, the lookups by hand {w}"
                    ))
                }
            }
        }
    }
    Ok(LOOKUP_ROWS as usize * 5)
}

// ---- the reference: the criteria by hand -----------------------------------

/// The criteria fixture's ledger, `A2:C20001`, and the accounts of its
/// summary, `R2:AD98`, a row each (KERNEL.8, its fifth slice).
const LEDGER_ROWS: u64 = 20_000;
const ACCOUNTS: usize = 97;

/// The summary by hand for the assumption `p1`, as the program's own
/// arithmetic writes the ledger: an account its row times 7,919 modulo 97, a
/// month its row times 31 modulo 12 and 1, an amount its row times 13
/// modulo 1,000 and `p1`; each account's amounts added by month in an array
/// here, not the engine's, and its rows counted.
fn criteria_by_hand(p1: u64) -> Vec<[u64; 13]> {
    let mut summary = vec![[0u64; 13]; ACCOUNTS];
    for r in 2..=LEDGER_ROWS + 1 {
        let account = ((r * 7_919) % 97) as usize;
        let month = ((r * 31) % 12) as usize;
        summary[account][month] += (r * 13) % 1_000 + p1;
        summary[account][12] += 1;
    }
    summary
}

/// The machine's summary, its twelve `SUMIFS` and its `COUNTIFS` a row,
/// `R2:AD98`, equal to the criteria by hand for the assumption `p1`.
fn held_criteria(m: &Machine, p1: &str) -> Result<usize, String> {
    let p1: u64 = p1
        .parse()
        .map_err(|_| format!("the assumption {p1} is no whole number"))?;
    let si = m.find_sheet("Output").ok_or("no sheet Output")?;
    for (a, want) in criteria_by_hand(p1).iter().enumerate() {
        let row = a as u32 + 2;
        for (j, w) in want.iter().enumerate() {
            let col = 18 + j as u32;
            match m.computed((si, row, col)) {
                Some(Computed::Value(Value::Number(n))) if n == *w as f64 => {}
                got => {
                    return Err(format!(
                        "Output!R{row}C{col}, with P1 {p1}: the machine holds {got:?}, the criteria by hand {w}"
                    ))
                }
            }
        }
    }
    Ok(ACCOUNTS * 13)
}

// ---- the reference: the netlist by hand -------------------------------------

/// Every gate of the netlist fixture equal to the adder evaluated by hand for
/// the carry-in `cin` (`KERNEL.54`, `bench/netlist.rs`).
fn held_netlist(m: &Machine, cin: &str) -> Result<usize, String> {
    let cin: u8 = cin
        .parse()
        .ok()
        .filter(|&c| c <= 1)
        .ok_or_else(|| format!("the carry-in {cin} is neither 0 nor 1"))?;
    let si = m.find_sheet("Output").ok_or("no sheet Output")?;
    let gates = netlist::by_hand(cin);
    for &((row, col), v) in &gates {
        match m.computed((si, row, col)) {
            Some(Computed::Value(Value::Number(n))) if n == f64::from(v) => {}
            got => {
                return Err(format!(
                    "Output!R{row}C{col}, with the carry-in {cin}: the machine holds {got:?}, the adder by hand {v}"
                ))
            }
        }
    }
    Ok(gates.len())
}

// ---- a fixture: its program and what its comments say ---------------------

/// A sheet and a cell or a range, as `Sheet!A1` writes it.
struct At {
    sheet: String,
    area: A1Range,
    text: String,
}

fn at(word: &str) -> Option<At> {
    let (sheet, addr) = word.split_once('!')?;
    Some(At {
        sheet: sheet.to_string(),
        area: parse_a1_range(addr)?,
        text: addr.to_string(),
    })
}

/// How a fixture is measured, read off its `# bench:` comments.
enum Change {
    /// A value cell written with its other value, and back.
    Edit {
        cell: At,
        built: String,
        other: String,
    },
    /// The frame cell advanced, the previous frame copied into a sheet of
    /// its own by Excel and kept as its twin by the machine.
    Frames {
        cell: At,
        count: usize,
        twin: String,
        seed: At,
    },
}

struct Fixture {
    name: String,
    program: String,
    change: Change,
    sum: At,
    /// The first cell typed into, a fresh one each time down its column
    /// (`KERNEL.56`); none for a fixture whose comments name no type.
    typed: Option<At>,
}

fn read_fixture(root: &Path, name: &str) -> Result<Fixture, String> {
    let path = root
        .join("scripts/recalc/bench")
        .join(format!("{name}.txt"));
    let program = std::fs::read_to_string(&path)
        .map_err(|e| format!("{}: cannot be read: {e}", path.display()))?;
    let mut change = None;
    let mut sum = None;
    let mut typed = None;
    for line in program.lines() {
        let Some(rest) = line.trim().strip_prefix("# bench:") else {
            continue;
        };
        let words: Vec<&str> = rest.split_whitespace().collect();
        let bad = || format!("{name}: a bench comment this runner does not read: {line}");
        match words.as_slice() {
            ["edit", cell, built, other] => {
                change = Some(Change::Edit {
                    cell: at(cell).ok_or_else(bad)?,
                    built: built.to_string(),
                    other: other.to_string(),
                })
            }
            ["frames", cell, count, "twin", twin, "seed", seed] => {
                change = Some(Change::Frames {
                    cell: at(cell).ok_or_else(bad)?,
                    count: count.parse().map_err(|_| bad())?,
                    twin: twin.to_string(),
                    seed: at(seed).ok_or_else(bad)?,
                })
            }
            ["sum", area] => sum = Some(at(area).ok_or_else(bad)?),
            ["type", cell] => typed = Some(at(cell).ok_or_else(bad)?),
            _ => return Err(bad()),
        }
    }
    Ok(Fixture {
        name: name.to_string(),
        program,
        change: change.ok_or_else(|| format!("{name}: no bench comment says what changes"))?,
        sum: sum.ok_or_else(|| format!("{name}: no bench comment names the sum"))?,
        typed,
    })
}

fn same(a: &Value, b: &Value) -> bool {
    match (a, b) {
        (Value::Number(x), Value::Number(y)) => x == y,
        _ => a == b,
    }
}

/// The cells of one sheet within an area, by place, each with the number it
/// holds as a value cell, read from the grid the machine lends, a copy whose
/// twins it makes from their sheets (`KERNEL.49`).
struct Cells {
    sheet: usize,
    places: Vec<((u32, u32), Option<f64>)>,
}

fn cells_within(m: &Machine, sheet: &str, a: &A1Range) -> Result<Cells, String> {
    let si = m
        .find_sheet(sheet)
        .ok_or_else(|| format!("no sheet {sheet}"))?;
    let places = m.workbook().sheets[si]
        .cells
        .range((a.top, 0)..=(a.bottom, u32::MAX))
        .filter(|(place, _)| place.1 >= a.left && place.1 <= a.right)
        .map(|(&place, c)| {
            let own = match c.content {
                Content::Number(n) => Some(n),
                _ => None,
            };
            (place, own)
        })
        .collect();
    Ok(Cells { sheet: si, places })
}

/// A number a cell holds in the machine: a formula's last complete value, or
/// the number it holds as a value cell, `own`.
fn number_at(m: &Machine, id: CellId, own: Option<f64>) -> Option<f64> {
    if let Some(Computed::Value(Value::Number(n))) = m.computed(id) {
        return Some(n);
    }
    own
}

/// The sum of the numbers a range holds, as Excel's `SUM` reads it.
fn checksum(m: &mut Machine, area: &At) -> Result<f64, String> {
    let cells = cells_within(m, &area.sheet, &area.area).map_err(|e| format!("{e} for the sum"))?;
    let mut total = 0.0;
    for &((row, col), own) in &cells.places {
        if let Some(n) = number_at(m, (cells.sheet, row, col), own) {
            total += n;
        }
    }
    Ok(total)
}

/// Every formula's value in the machine equal to the one the build folded
/// into its file, by the sheet's name: the first frame against `Calc`.
fn held_to_file(
    m: &Machine,
    file: &frazaro_core::reflect::calc::Collected,
) -> Result<usize, String> {
    let mut held = 0;
    for (&(si, row, col), want) in &file.cached {
        let formula = matches!(
            file.model.sheets[si]
                .cells
                .get(&(row, col))
                .map(|c| &c.content),
            Some(
                Content::Formula(_)
                    | Content::SharedMaster { .. }
                    | Content::SharedChild { .. }
                    | Content::DynamicFormula(_)
            )
        );
        if !formula {
            continue;
        }
        let name = &file.model.sheets[si].name;
        let Some(mi) = m.find_sheet(name) else {
            continue;
        };
        if mi >= m.sheets() {
            continue;
        }
        let Some(want) = Value::from_file(want) else {
            return Err(format!(
                "{name}!R{row}C{col}: the file holds {want:?}, no value to compare"
            ));
        };
        match m.computed((mi, row, col)) {
            Some(Computed::Value(got)) if same(&got, &want) => held += 1,
            got => {
                return Err(format!(
                    "{name}!R{row}C{col}: the machine holds {got:?}, the file {want:?}"
                ))
            }
        }
    }
    Ok(held)
}

/// Every formula of the machine's grid equal to `Calc` over the same grid,
/// the copy the machine lends, its twins made from their sheets (`KERNEL.49`).
fn held_to_calc(m: &Machine, lib: &vla_lang::calc::Library) -> Result<usize, String> {
    let wants: Vec<(CellId, Option<Computed>)> = {
        let book = m.workbook();
        let calc = Calc::run(&book, lib);
        calc.formula_cells()
            .into_iter()
            .map(|id| (id, calc.computed(id).cloned()))
            .collect()
    };
    let mut held = 0;
    for (id, want) in wants {
        if id.0 >= m.sheets() {
            continue;
        }
        match (m.computed(id), want.as_ref()) {
            (Some(Computed::Value(a)), Some(Computed::Value(b))) if same(&a, b) => held += 1,
            (got, want) => {
                return Err(format!(
                    "{}!R{}C{}: the machine holds {got:?}, Calc {want:?}",
                    m.sheet_name(id.0).unwrap_or("?"),
                    id.1,
                    id.2
                ))
            }
        }
    }
    Ok(held)
}

/// The machine's plane of Life's sheet, a byte a cell, the border dead, over
/// the sheet's cells as a round read them before its first frame: Life's
/// frames write none of them, and reading the grid at every frame would make
/// the twin's cells, which the engine's own frames never do.
fn plane(m: &Machine, screen: &Cells) -> Vec<u8> {
    let mut g = vec![0u8; W * H];
    for &((r, c), own) in &screen.places {
        if let Some(n) = number_at(m, (screen.sheet, r, c), own) {
            g[(r as usize - 1) * W + (c as usize - 1)] = u8::from(n == 1.0);
        }
    }
    g
}

fn cell_name(i: usize) -> String {
    let (row, mut col) = (i / W + 1, i % W + 1);
    let mut letters = Vec::new();
    while col > 0 {
        letters.push(b'A' + ((col - 1) % 26) as u8);
        col = (col - 1) / 26;
    }
    letters.reverse();
    format!("{}{row}", String::from_utf8(letters).unwrap_or_default())
}

// ---- one fixture: prepared once, measured in rounds -----------------------

/// A fixture built, written for Excel's half and read back: what every
/// round loads into a fresh machine.
struct Prepared {
    sha: String,
    bytes: usize,
    rows: String,
    read: frazaro_core::reflect::calc::Collected,
}

/// What a fixture's rounds measured: each round's median of its whole
/// frames and of its edits, every step's time for the range, and the clock
/// beside every timed step.
struct Run {
    name: String,
    sha: String,
    bytes: usize,
    formulas: usize,
    loads: Vec<f64>,
    firsts: Vec<f64>,
    whole_rounds: Vec<f64>,
    edit_rounds: Vec<f64>,
    whole: Vec<f64>,
    edits: Vec<f64>,
    /// Each round's median of its types and every type's time (KERNEL.56).
    type_rounds: Vec<f64>,
    types: Vec<f64>,
    calcs: Vec<f64>,
    clocks: Vec<f64>,
    sums: Option<(f64, f64)>,
    held: usize,
}

fn timed<T>(f: impl FnOnce() -> T) -> (T, f64) {
    let t = Instant::now();
    let out = f();
    (out, t.elapsed().as_secs_f64() * 1e3)
}

fn prepare(out: &Path, texts: (&str, &str), fx: &Fixture) -> Result<Prepared, String> {
    let name = &fx.name;
    let (prelude, english) = texts;
    let bytes = api::english_build_xlsx(&fx.program, prelude, &[english]).map_err(|e| {
        format!(
            "{name}: the build refused at line {}: {}",
            e.line, e.refusal.text
        )
    })?;
    let file = out.join(format!("{name}.xlsx"));
    std::fs::write(&file, &bytes).map_err(|e| format!("{}: {e}", file.display()))?;
    let rows = api::reflect_relations(&bytes, &format!("{name}.xlsx"))
        .map_err(|r| format!("{name}: the reader refused: {}", r.text))?;
    let twin = match &fx.change {
        Change::Frames { twin, .. } => Some(twin.as_str()),
        Change::Edit { .. } => None,
    };
    let read = collect(&bytes, "bench.xlsx").map_err(|r| format!("{name}: {}", r.text))?;
    Ok(Prepared {
        sha: sha256::sha256_hex(&bytes),
        bytes: bytes.len(),
        rows: rows_for_machine(&rows, twin),
        read,
    })
}

/// One round: a fresh machine loaded, its first frame held to the file, the
/// change (the edits, or Life's frames, each held), then whole frames, a
/// formula no formula reads written before each, Life's held as well, since
/// every frame of Life is a generation.
fn round(
    prep: &Prepared,
    fx: &Fixture,
    check: bool,
    probe: &mut Probe,
    r: &mut Run,
) -> Result<(), String> {
    let name = &fx.name;
    let (loaded, load_ms) = timed(|| Machine::load(&prep.rows, excel::library()));
    let mut m = loaded.map_err(|e| {
        format!(
            "{name}: the machine refused line {}: {}",
            e.line, e.refusal.text
        )
    })?;
    let formulas = m.formulas().0;
    r.formulas = formulas;
    let (s, first_ms) = timed(|| m.step(0));
    if !s.done {
        return Err(format!("{name}: the first step answered {}", s.row()));
    }
    let held = held_to_file(&m, &prep.read).map_err(|e| format!("{name}: {e}"))?;
    if held != formulas {
        return Err(format!(
            "{name}: {held} formulas held to the file's values of {formulas}"
        ));
    }
    r.held += held;
    // The lookups' fixture held to its lookups by hand as well, the
    // criteria's to its summary by hand (KERNEL.8), and the netlist's to its
    // adder by hand (KERNEL.54).
    let by_hand = |m: &Machine, p1: &str| -> Result<usize, String> {
        match name.as_str() {
            "lookups" => held_lookups(m, p1).map_err(|e| format!("{name}: {e}")),
            "criteria" => held_criteria(m, p1).map_err(|e| format!("{name}: {e}")),
            "netlist" => held_netlist(m, p1).map_err(|e| format!("{name}: {e}")),
            _ => Ok(0),
        }
    };
    if let Change::Edit { built, .. } = &fx.change {
        r.held += by_hand(&m, built)?;
    }
    r.loads.push(load_ms);
    r.firsts.push(first_ms);
    let sum_built = checksum(&mut m, &fx.sum).map_err(|e| format!("{name}: {e}"))?;
    let (whole_n, edit_n) = if check { (2, 3) } else { (WHOLE, EDITS) };
    let mut whole = Vec::with_capacity(whole_n);
    let mut edits = Vec::with_capacity(edit_n);

    let mut life: Option<(Vec<u8>, Cells)> = None;
    match &fx.change {
        Change::Edit { cell, built, other } => {
            // An odd count, alternating, the last the other value.
            for k in 0..edit_n {
                let v = if k % 2 == 0 { other } else { built };
                m.write(&format!("(cell \"{}\" \"{}\" {v})", cell.sheet, cell.text))
                    .map_err(|e| format!("{name}: {}", e.refusal.text))?;
                if !check {
                    r.clocks.push(probe.read());
                }
                let (s, ms) = timed(|| m.step(0));
                if !s.done {
                    return Err(format!("{name}: an edit's frame answered {}", s.row()));
                }
                edits.push(ms);
                // Checking, every edit held, since alternating values could
                // hide a machine that stopped following them: its last frame
                // computed would equal the state the edits came back to.
                if check {
                    r.held += held_to_calc(&m, &excel::library())
                        .map_err(|e| format!("{name}: after edit {}: {e}", k + 1))?;
                    r.held += by_hand(&m, v)?;
                }
            }
            if !check {
                r.held +=
                    held_to_calc(&m, &excel::library()).map_err(|e| format!("{name}: {e}"))?;
                r.held += by_hand(&m, other)?;
            }
        }
        Change::Frames {
            cell, count, seed, ..
        } => {
            let seeds = cells_within(&m, &seed.sheet, &seed.area)
                .map_err(|e| format!("{name}: {e} for the seed"))?;
            let seed = seeds
                .places
                .first()
                .and_then(|&((row, col), own)| number_at(&m, (seeds.sheet, row, col), own))
                .ok_or_else(|| format!("{name}: no seed in {}!{}", seed.sheet, seed.text))?;
            let mut want = soup(seed as u64);
            let extent = A1Range {
                top: 1,
                left: 1,
                bottom: H as u32,
                right: W as u32,
            };
            let screen =
                cells_within(&m, &fx.sum.sheet, &extent).map_err(|e| format!("{name}: {e}"))?;
            held_plane(&m, &screen, &want, 1).map_err(|e| format!("{name}: {e}"))?;
            let last = *count;
            for f in 2..=last {
                want = generation(&want);
                m.write(&format!("(cell \"{}\" \"{}\" {f})", cell.sheet, cell.text))
                    .map_err(|e| format!("{name}: {}", e.refusal.text))?;
                if !check {
                    r.clocks.push(probe.read());
                }
                let (s, ms) = timed(|| m.step(0));
                if !s.done {
                    return Err(format!("{name}: frame {f} answered {}", s.row()));
                }
                edits.push(ms);
                held_plane(&m, &screen, &want, f).map_err(|e| format!("{name}: {e}"))?;
                r.held += formulas;
            }
            life = Some((want, screen));
        }
    }
    let sum_edited = checksum(&mut m, &fx.sum).map_err(|e| format!("{name}: {e}"))?;
    match r.sums {
        None => r.sums = Some((sum_built, sum_edited)),
        Some(was) if was == (sum_built, sum_edited) => {}
        Some(was) => {
            return Err(format!(
                "{name}: a round summed {sum_built} and {sum_edited}, an earlier one {} and {}",
                was.0, was.1
            ))
        }
    }
    // The type (KERNEL.56): 1 written into a fresh cell, a row further down
    // each time, and the frames to the grid's stillness, timed with the
    // write; held to `Calc` after each type in the check.
    if let Some(t) = &fx.typed {
        let type_n = if check { 3 } else { TYPES };
        let mut types = Vec::with_capacity(type_n);
        for k in 0..type_n {
            let addr = vla_lang::sheet::cell_ref(t.area.top + k as u32, t.area.left);
            let w = format!("(cell \"{}\" \"{addr}\" 1)", t.sheet);
            if !check {
                r.clocks.push(probe.read());
            }
            let (still, ms) = timed(|| typed(&mut m, &w));
            still.map_err(|e| format!("{name}: typing {}!{addr}: {e}", t.sheet))?;
            types.push(ms);
            if check {
                r.held += held_to_calc(&m, &excel::library())
                    .map_err(|e| format!("{name}: after typing {}!{addr}: {e}", t.sheet))?;
            }
        }
        r.type_rounds.push(median(&types));
        r.types.extend(types);
    }
    for k in 0..whole_n {
        m.write(&format!("(formula \"Frazaro\" \"E1\" \"={}\")", k % 2))
            .map_err(|e| format!("{name}: {}", e.refusal.text))?;
        if !check {
            r.clocks.push(probe.read());
        }
        let (s, ms) = timed(|| m.step(0));
        if !s.done {
            return Err(format!("{name}: a whole frame answered {}", s.row()));
        }
        whole.push(ms);
        if let Some((want, screen)) = &mut life {
            *want = generation(want);
            held_plane(&m, screen, want, s.frame as usize).map_err(|e| format!("{name}: {e}"))?;
            r.held += formulas;
        }
    }
    r.whole_rounds.push(median(&whole));
    r.edit_rounds.push(median(&edits));
    r.whole.extend(whole);
    r.edits.extend(edits);
    Ok(())
}

/// A type's write and its frames to the grid's stillness (KERNEL.56): the
/// frame after a write proves nothing, a write before it, so the frames go on
/// until one changes nothing, within [`STILL_WITHIN`].
fn typed(m: &mut Machine, write: &str) -> Result<usize, String> {
    m.write(write).map_err(|e| e.refusal.text)?;
    for frame in 1..=STILL_WITHIN {
        let s = m.step(0);
        if !s.done {
            return Err(format!("frame {frame} answered {}", s.row()));
        }
        if frame > 1 && m.changed() == 0 {
            return Ok(frame);
        }
    }
    Err(format!("the grid not still within {STILL_WITHIN} frames"))
}

/// The machine's plane of Life's sheet equal to the hand-written Life's.
fn held_plane(m: &Machine, screen: &Cells, want: &[u8], frame: usize) -> Result<(), String> {
    let got = plane(m, screen);
    match got.iter().zip(want).position(|(a, b)| a != b) {
        None => Ok(()),
        Some(i) => Err(format!(
            "frame {frame}'s cell {} is {} in the machine and {} by hand",
            cell_name(i),
            got[i],
            want[i]
        )),
    }
}

// ---- the flag: the build golden's fixture written three ways --------------

/// The build golden's fixture built, `Output!C2`'s folded 15 planted as 999,
/// and its `calcPr` as built, as Excel 365 writes its own, and left out.
fn write_flag_variants(
    out: &Path,
    root: &Path,
    texts: (&str, &str),
) -> Result<Vec<String>, String> {
    let program = std::fs::read_to_string(root.join("scripts/build/fixture.txt"))
        .map_err(|e| format!("the build golden's fixture: {e}"))?;
    let bytes = api::english_build_xlsx(&program, texts.0, &[texts.1])
        .map_err(|e| format!("the build golden's fixture: {}", e.refusal.text))?;
    let entries = zip::entries(&bytes).ok_or("the build golden's fixture is no archive")?;
    let mut parts: Vec<(String, Vec<u8>)> = Vec::new();
    for e in &entries {
        let data = zip::stored_data(&bytes, e).ok_or("an entry not stored")?;
        parts.push((e.name.clone(), data.to_vec()));
    }
    let planted = parts
        .iter_mut()
        .filter(|(n, _)| n.starts_with("xl/worksheets/"))
        .find_map(|(_, data)| {
            let text = String::from_utf8(data.clone()).ok()?;
            let start = text.find("<c r=\"C2\"")?;
            let end = start + text[start..].find("</c>")?;
            let cell = &text[start..end];
            if !cell.contains("<v>15</v>") {
                return None;
            }
            let replaced = format!(
                "{}{}{}",
                &text[..start],
                cell.replacen("<v>15</v>", "<v>999</v>", 1),
                &text[end..]
            );
            *data = replaced.into_bytes();
            Some(())
        });
    if planted.is_none() {
        return Err("Output!C2 holds no folded 15 to plant".to_string());
    }
    let mut written = Vec::new();
    for (file, calc_pr) in [
        ("calcpr_built.xlsx", Some("<calcPr fullCalcOnLoad=\"1\"/>")),
        ("calcpr_calcid.xlsx", Some("<calcPr calcId=\"191029\"/>")),
        ("calcpr_none.xlsx", None),
    ] {
        let mut these = parts.clone();
        let wb = these
            .iter_mut()
            .find(|(n, _)| n == "xl/workbook.xml")
            .ok_or("no xl/workbook.xml")?;
        let text = String::from_utf8(wb.1.clone()).map_err(|_| "xl/workbook.xml is not UTF-8")?;
        let built = "<calcPr fullCalcOnLoad=\"1\"/>";
        if text.matches(built).count() != 1 {
            return Err("xl/workbook.xml holds no single calcPr as built".to_string());
        }
        wb.1 = text.replace(built, calc_pr.unwrap_or("")).into_bytes();
        let bytes = zip::write_stored(&these).ok_or("the variant would not write")?;
        std::fs::write(out.join(file), bytes).map_err(|e| format!("{file}: {e}"))?;
        written.push(file.to_string());
    }
    Ok(written)
}

// ---- the runner -------------------------------------------------------------

fn main() -> ExitCode {
    let args: Vec<String> = std::env::args().skip(1).collect();
    let check = args.iter().any(|a| a == "--check");
    let write_netlist = args.iter().any(|a| a == "--netlist");
    if let Some(a) = args
        .iter()
        .find(|a| !matches!(a.as_str(), "--check" | "--netlist"))
    {
        eprintln!("recalc: an unknown argument, {a}");
        return ExitCode::from(2);
    }
    let root: PathBuf = Path::new(env!("CARGO_MANIFEST_DIR")).join("..");
    // The netlist program, written by its generator or held to it (KERNEL.54).
    let netlist_path = root.join("scripts/recalc/bench/netlist.txt");
    if write_netlist {
        let text = netlist::program();
        if let Err(e) = std::fs::write(&netlist_path, &text) {
            eprintln!("recalc: {}: {e}", netlist_path.display());
            return ExitCode::FAILURE;
        }
        println!(
            "recalc: wrote {}: {} bytes, {} lines",
            netlist_path.display(),
            thousands(text.len() as u64),
            thousands(text.lines().count() as u64)
        );
        return ExitCode::SUCCESS;
    }
    if check {
        match std::fs::read_to_string(&netlist_path) {
            Ok(have) if have.replace("\r\n", "\n") == netlist::program() => {}
            Ok(_) => {
                println!("(recalc-differs \"netlist: scripts/recalc/bench/netlist.txt is not its generator's program; write it again with --netlist\")");
                return ExitCode::FAILURE;
            }
            Err(e) => {
                println!(
                    "(recalc-differs \"netlist: {}: {e}\")",
                    netlist_path.display()
                );
                return ExitCode::FAILURE;
            }
        }
    }
    let out = root.join("target").join("bench");
    if let Err(e) = std::fs::create_dir_all(&out) {
        eprintln!("recalc: {}: {e}", out.display());
        return ExitCode::FAILURE;
    }
    let read = |p: &str| std::fs::read_to_string(root.join(p));
    let (prelude, english) = match (
        read("scripts/prelude.vla"),
        read("scripts/polyglotta/english.vla"),
    ) {
        (Ok(p), Ok(e)) => (p, e),
        _ => {
            eprintln!("recalc: the prelude or the phrasebook cannot be read");
            return ExitCode::FAILURE;
        }
    };
    let mut fixtures = Vec::new();
    for name in FIXTURES {
        match read_fixture(&root, name) {
            Ok(f) => fixtures.push(f),
            Err(e) => {
                println!("{e}");
                return ExitCode::FAILURE;
            }
        }
    }

    // The clock at the start, before anything heavy has run.
    let mut probe = Probe::new(soup(1));
    let start_clock = if check {
        0.0
    } else {
        let readings: Vec<f64> = (0..START_READINGS).map(|_| probe.read()).collect();
        median(&readings)
    };
    let mut prepared = Vec::new();
    let mut runs = Vec::new();
    for fx in &fixtures {
        match prepare(&out, (&prelude, &english), fx) {
            Ok(p) => {
                runs.push(Run {
                    name: fx.name.clone(),
                    sha: p.sha.clone(),
                    bytes: p.bytes,
                    formulas: 0,
                    loads: Vec::new(),
                    firsts: Vec::new(),
                    whole_rounds: Vec::new(),
                    edit_rounds: Vec::new(),
                    whole: Vec::new(),
                    edits: Vec::new(),
                    type_rounds: Vec::new(),
                    types: Vec::new(),
                    calcs: Vec::new(),
                    clocks: Vec::new(),
                    sums: None,
                    held: 0,
                });
                prepared.push(p);
            }
            Err(e) => {
                println!("(recalc-differs {e:?})");
                eprintln!("recalc: FAILED, {e}");
                return ExitCode::FAILURE;
            }
        }
    }
    // The rounds, the fixtures in turn, each round a fresh machine: one
    // process's memory placed one way can move a frame of the cell-by-cell
    // path by up to a third (measured 2026-10-10), and a round's machine is
    // placed anew.
    let rounds = if check { 1 } else { ROUNDS };
    for _ in 0..rounds {
        for ((fx, p), r) in fixtures.iter().zip(&prepared).zip(runs.iter_mut()) {
            if let Err(e) = round(p, fx, check, &mut probe, r) {
                println!("(recalc-differs {e:?})");
                eprintln!("recalc: FAILED, {e}");
                return ExitCode::FAILURE;
            }
        }
    }
    // The reference the doors compute with, over each file read back.
    let lib = excel::library();
    for (p, r) in prepared.iter().zip(runs.iter_mut()) {
        for _ in 0..(if check { 1 } else { CALCS }) {
            // The clock beside the reference too: a run disturbed here is
            // as much not a baseline as one disturbed beside a frame.
            if !check {
                r.clocks.push(probe.read());
            }
            let (c, ms) = timed(|| Calc::run(&p.read.model, &lib));
            black_box(c.formulas());
            r.calcs.push(ms);
        }
    }
    let flags = match write_flag_variants(&out, &root, (&prelude, &english)) {
        Ok(f) => f,
        Err(e) => {
            println!("(recalc-flag-failed {e:?})");
            return ExitCode::FAILURE;
        }
    };

    if check {
        for r in &runs {
            let (built, edited) = r.sums.unwrap_or_default();
            println!(
                "(recalc-check \"{}\" (sha256 \"{}\") (formulas {}) (held {}) (sum-built {built}) (sum-edited {edited}))",
                r.name, r.sha, r.formulas, r.held
            );
        }
        println!("(recalc-check-flags {})", flags.len());
        println!("(recalc-check-done {})", runs.len());
        return ExitCode::SUCCESS;
    }

    println!(
        "recalc: the native runner, {} on {}, {} fixtures in {rounds} rounds, every value held before any number; the files in target/bench",
        std::env::consts::ARCH,
        std::env::consts::OS,
        runs.len()
    );
    let mut steady_all = true;
    let mut clocks_of = BTreeMap::new();
    for r in &runs {
        let whole = median(&r.whole_rounds);
        let (wlo, whi) = range(&r.whole);
        let edit = median(&r.edit_rounds);
        let (elo, ehi) = range(&r.edits);
        let clock = median(&r.clocks);
        let sp = spread(&r.clocks);
        let dr = drift(clock, start_clock);
        clocks_of.insert(r.name.clone(), (clock, sp, dr));
        let (built, edited) = r.sums.unwrap_or_default();
        let rounds_text = |v: &[f64]| {
            v.iter()
                .map(|x| format!("{x:.3}"))
                .collect::<Vec<_>>()
                .join(", ")
        };
        println!(
            "{}: {} formulas, {} bytes; load {:.1} ms; first frame {:.2} ms; whole frame {:.3} ms, the median of the rounds' {} (every frame {:.3} to {:.3}), {} formulas a second; the edit's frame {:.3} ms, of {} ({:.3} to {:.3}); Calc {:.1} ms; sums {built} as built, {edited} edited",
            r.name,
            thousands(r.formulas as u64),
            thousands(r.bytes as u64),
            median(&r.loads),
            median(&r.firsts),
            whole,
            rounds_text(&r.whole_rounds),
            wlo,
            whi,
            thousands((r.formulas as f64 / (whole / 1e3)).round() as u64),
            edit,
            rounds_text(&r.edit_rounds),
            elo,
            ehi,
            median(&r.calcs)
        );
        if !r.type_rounds.is_empty() {
            let typed = median(&r.type_rounds);
            let (tlo, thi) = range(&r.types);
            println!(
                "  the type: a value written into a fresh cell and the frames to the grid's stillness, {typed:.3} ms, the median of the rounds' {} ({tlo:.3} to {thi:.3}), {:.1} a second",
                rounds_text(&r.type_rounds),
                1e3 / typed
            );
        }
        println!(
            "  the clock: the hand-written Life {clock:.3} ms a generation beside every timed frame, spread {sp:.1}%, {dr:+.1}% from {start_clock:.3} at the start"
        );
        if sp > CLOCK_SPREAD_MAX || dr.abs() > CLOCK_SPREAD_MAX {
            steady_all = false;
            println!(
                "  the clock moved during the run, more than {CLOCK_SPREAD_MAX}%: power throttling or a thermal limit; these frames are not a baseline"
            );
        }
    }
    println!(
        "the flag's three files for Excel's half: {}",
        flags.join(", ")
    );
    if !steady_all {
        println!("no baseline rows: run again through tools/bench_recalc.ps1, which opts the runner out of Windows' power throttling");
        return ExitCode::from(3);
    }
    println!("the baseline rows, for tools/check_recalc_floors.ps1:");
    for r in &runs {
        let (clock, sp, dr) = clocks_of[&r.name];
        let whole = median(&r.whole_rounds);
        let edit = median(&r.edit_rounds);
        let (built, edited) = r.sums.unwrap_or_default();
        // A typed fixture's type and its rate (KERNEL.56), beside its edit's.
        let typed = match r.type_rounds.is_empty() {
            true => String::new(),
            false => {
                let t = median(&r.type_rounds);
                format!(" Type = {t:.3}; TypesPerSecond = {:.1};", 1e3 / t)
            }
        };
        println!(
            "    @{{ Half = 'native'; Fixture = '{}'; Formulas = {}; Sha256 = '{}'; Rounds = {rounds}; Load = {:.1}; First = {:.2}; Whole = {:.3}; Edit = {:.3}; Calc = {:.1}; WholePerSecond = {}; EditsPerSecond = {:.1};{typed} SumBuilt = {built}; SumEdited = {edited}; ClockMs = {:.3}; ClockSpread = {:.1}; ClockDrift = {:.1} }}",
            r.name,
            r.formulas,
            r.sha,
            median(&r.loads),
            median(&r.firsts),
            whole,
            edit,
            median(&r.calcs),
            (r.formulas as f64 / (whole / 1e3)).round() as u64,
            1e3 / edit,
            clock,
            sp,
            dr
        );
    }
    ExitCode::SUCCESS
}
