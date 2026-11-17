//! The memory floors' and the load floors' runner (`KERNEL.49`, its first
//! slice; `docs/BETA_REARVIEW.md`, the item's scoping, decision 12 and its
//! addendum): every bench program of `scripts/recalc/bench/` built through the
//! core's own surface and read back into rows, as the recalc runner builds it
//! (`recalc.rs`), and the fixtures of rows beside them, Superposition's board
//! and Life on the size ladder's sixteenth rung, which the machine loads as
//! they stand, and since `KERNEL.67` the board with one style row over its
//! Field, loaded after the board's own rows; each loaded into the machine
//! under Excel's library inside a counting allocator, and stepped twice; and
//! since `KERNEL.72` its save printed, every sheet's fills as a host prints
//! them and since `KERNEL.29` the grid's names, and loaded again.
//!
//! ```text
//! cargo run --release -p frazaro-core --example memory               # every fixture, the baseline rows
//! cargo run --release -p frazaro-core --example memory -- --check    # nothing timed: the bytes alone
//! ```
//!
//! **What is counted**, in a pass of its own for each fixture: the bytes the
//! allocator was asked for and not given back, from just before the load (the
//! rows' text read and the library made) to just after the second frame, which
//! are the load's parse, its draft and its plan, the machine itself and what
//! two frames add; and the most of them at any moment between, which a
//! WebAssembly module, whose memory never shrinks, pays whole. Then, a pass of
//! their own after the frames, a save's: its text's bytes, the most the printing
//! of it held over what the machine holds, the bytes a load from it held after
//! the load, and the most at any moment of that load, the text the caller's and
//! outside the count. Bytes asked are bytes asked: the same on every run of one
//! build, and on every 64-bit machine, so the check holds them exactly, on every
//! push, as ceilings that never rise.
//!
//! **What is timed:** the load alone, `Machine::load` from the rows' text, the
//! median of three rounds, each a fresh machine, beside the processor's clock
//! (`bench/common.rs`, the recalc runner's): when it moves more than 10% across
//! the run, or from its start, the run prints no rows. The count stays on while
//! a load is timed, an atomic add a block, the same in every run.
//!
//! The runner is an example: no member of the workspace, nothing in the lock,
//! and out of the crate's package.

use std::alloc::{GlobalAlloc, Layout, System};
use std::collections::BTreeMap;
use std::hint::black_box;
use std::path::{Path, PathBuf};
use std::process::ExitCode;
use std::sync::atomic::{AtomicUsize, Ordering};
use std::time::Instant;

use frazaro_core::{api, excel, sha256};
use vla_lang::machine::{Machine, Viewed};

#[path = "bench/common.rs"]
mod common;
use common::{
    build_program, drift, median, rows_for_machine, soup, spread, thousands, Probe,
    CLOCK_SPREAD_MAX, START_READINGS,
};

// ---- the count ---------------------------------------------------------------

/// The allocator, counting: the bytes asked and not given back, and the most
/// of them since the count was last reset.
struct Counting;

static LIVE: AtomicUsize = AtomicUsize::new(0);
static PEAK: AtomicUsize = AtomicUsize::new(0);

fn grew(n: usize) {
    let now = LIVE.fetch_add(n, Ordering::Relaxed) + n;
    PEAK.fetch_max(now, Ordering::Relaxed);
}

// SAFETY: every call is `System`'s, the counts beside it.
unsafe impl GlobalAlloc for Counting {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        // SAFETY: the caller's layout, passed on.
        let p = unsafe { System.alloc(layout) };
        if !p.is_null() {
            grew(layout.size());
        }
        p
    }

    unsafe fn alloc_zeroed(&self, layout: Layout) -> *mut u8 {
        // SAFETY: the caller's layout, passed on.
        let p = unsafe { System.alloc_zeroed(layout) };
        if !p.is_null() {
            grew(layout.size());
        }
        p
    }

    unsafe fn dealloc(&self, p: *mut u8, layout: Layout) {
        LIVE.fetch_sub(layout.size(), Ordering::Relaxed);
        // SAFETY: the caller's block, given back.
        unsafe { System.dealloc(p, layout) }
    }

    unsafe fn realloc(&self, p: *mut u8, layout: Layout, size: usize) -> *mut u8 {
        // SAFETY: the caller's block and size, passed on.
        let q = unsafe { System.realloc(p, layout, size) };
        if !q.is_null() {
            LIVE.fetch_sub(layout.size(), Ordering::Relaxed);
            grew(size);
        }
        q
    }
}

#[global_allocator]
static ALLOCATOR: Counting = Counting;

fn live() -> usize {
    LIVE.load(Ordering::Relaxed)
}

/// The most held from now on counted from what is held now.
fn reset() {
    PEAK.store(live(), Ordering::Relaxed);
}

fn peak() -> usize {
    PEAK.load(Ordering::Relaxed)
}

// ---- the fixtures ------------------------------------------------------------

/// The bench programs, built through the door, in the order the recalc runner
/// prints them; and the fixtures of rows after them.
const PROGRAMS: [&str; 9] = [
    "column",
    "chain",
    "sums",
    "life",
    "lookups",
    "criteria",
    "netlist",
    "adder",
    "adder-named",
];
const ROWS: [&str; 3] = ["superposition", "life-x16", "superposition-styled"];

/// A fixture of rows that adds to another's, its twin, and is loaded after
/// it (KERNEL.67): Superposition's board with one style row over its Field,
/// held by the floors' check within a constant of the board's own bytes.
const AFTER: [(&str, &str); 1] = [("superposition-styled", "superposition")];

/// The rounds a load is timed in, each a fresh machine.
const ROUNDS: usize = 3;

/// The frames stepped after the load, each whole.
const FRAMES: usize = 2;

/// A fixture as the machine loads it: its name, its source's SHA-256 over LF
/// bytes (the program's or the rows' own file), and the rows.
struct Prepared {
    name: String,
    sha: String,
    rows: String,
}

fn lf_sha(text: &str) -> String {
    sha256::sha256_hex(text.replace("\r\n", "\n").as_bytes())
}

/// The sheet a program keeps its previous frame in, by its `# bench: frames`
/// comment, which the machine holds as its own twin (`recalc.rs`'s reading).
fn twin_of(program: &str) -> Option<String> {
    program.lines().find_map(|line| {
        let words: Vec<&str> = line
            .trim()
            .strip_prefix("# bench:")?
            .split_whitespace()
            .collect();
        match words.as_slice() {
            ["frames", _, _, "twin", twin, "seed", _] => Some(twin.to_string()),
            _ => None,
        }
    })
}

/// A source's text as every platform reads it: its lines ended by LF, so that
/// a checkout's line ends (CRLF on Windows, LF elsewhere) move no byte the
/// machine allocates.
fn read_lf(path: &Path) -> Result<String, String> {
    std::fs::read_to_string(path)
        .map(|t| t.replace("\r\n", "\n"))
        .map_err(|e| format!("{}: cannot be read: {e}", path.display()))
}

/// A bench program built as `frazaro build` builds it, its names added where
/// it defines any (`bench/common.rs`), and read back.
fn program(root: &Path, name: &str, (prelude, english): (&str, &str)) -> Result<Prepared, String> {
    let path = root
        .join("scripts/recalc/bench")
        .join(format!("{name}.txt"));
    let text = read_lf(&path)?;
    let bytes = build_program(&text, prelude, english).map_err(|e| format!("{name}: {e}"))?;
    let rows = api::reflect_relations(&bytes, &format!("{name}.xlsx"))
        .map_err(|r| format!("{name}: the reader refused: {}", r.text))?;
    Ok(Prepared {
        name: name.to_string(),
        sha: lf_sha(&text),
        rows: rows_for_machine(&rows, twin_of(&text).as_deref()).replace("\r\n", "\n"),
    })
}

/// A fixture of rows, as it stands, its lines ended by LF; one that adds to a
/// twin's loaded after the twin's rows, its SHA-256 its own file's, the
/// twin's pinned by the twin's row.
fn rows_fixture(root: &Path, name: &str) -> Result<Prepared, String> {
    let bench = root.join("scripts/recalc/bench");
    let text = read_lf(&bench.join(format!("{name}.vla")))?;
    let rows = match AFTER.iter().find(|(own, _)| *own == name) {
        Some((_, twin)) => read_lf(&bench.join(format!("{twin}.vla")))? + &text,
        None => text.clone(),
    };
    Ok(Prepared {
        name: name.to_string(),
        sha: lf_sha(&text),
        rows,
    })
}

// ---- one fixture: weighed once, its load timed in rounds -------------------------

/// What a fixture's machine held: after its load, after its frames, and at
/// the most between; and its save.
struct Weighed {
    formulas: usize,
    load_held: usize,
    held: usize,
    peak: usize,
    save: Saved,
}

/// What a machine's save costs (`KERNEL.72`): its text's bytes, the most its
/// printing held over the machine, and a load from it, held after the load and
/// at the most, one load's time.
struct Saved {
    bytes: usize,
    printing: usize,
    held: usize,
    peak: usize,
    ms: f64,
}

/// A machine's save as a host prints it, weighed: the grid's names, its
/// `names` view (`KERNEL.29`'s third slice), then every sheet's fills; and a
/// load from it under Excel's library.
fn weigh_save(m: &Machine, name: &str) -> Result<Saved, String> {
    let held = live();
    reset();
    let mut save = match m.view("names", "", None) {
        Ok(Viewed::Text(t)) => t,
        Ok(_) => return Err(format!("{name}: the names are no text")),
        Err(r) => return Err(format!("{name}: the names: {}", r.text)),
    };
    for i in 0..m.sheets() {
        let Some(sheet) = m.sheet_name(i) else {
            continue;
        };
        match m.view("fills", sheet, None) {
            Ok(Viewed::Text(t)) => save.push_str(&t),
            Ok(_) => return Err(format!("{name}: the fills of {sheet} are no text")),
            Err(r) => return Err(format!("{name}: the fills of {sheet}: {}", r.text)),
        }
    }
    let printing = peak() - held;
    let library = excel::library();
    let base = live();
    reset();
    let t = Instant::now();
    let again = Machine::load(&save, library).map_err(|e| {
        format!(
            "{name}: its save refused at line {}: {}",
            e.line, e.refusal.text
        )
    })?;
    let ms = t.elapsed().as_secs_f64() * 1e3;
    let saved = Saved {
        bytes: save.len(),
        printing,
        held: live() - base,
        peak: peak() - base,
        ms,
    };
    drop(again);
    Ok(saved)
}

fn weigh(p: &Prepared) -> Result<Weighed, String> {
    let name = &p.name;
    let library = excel::library();
    let base = live();
    reset();
    let mut m = Machine::load(&p.rows, library).map_err(|e| {
        format!(
            "{name}: the machine refused line {}: {}",
            e.line, e.refusal.text
        )
    })?;
    let load_held = live() - base;
    for frame in 1..=FRAMES {
        let s = m.step(0);
        if !s.done {
            return Err(format!("{name}: frame {frame} answered {}", s.row()));
        }
    }
    let formulas = m.formulas().0;
    let held = live() - base;
    let peak = peak() - base;
    let save = weigh_save(&m, name)?;
    drop(m);
    if formulas == 0 {
        return Err(format!("{name}: no formula loaded"));
    }
    Ok(Weighed {
        formulas,
        load_held,
        held,
        peak,
        save,
    })
}

/// The load timed in rounds, each a fresh machine, the clock read before each.
fn time_loads(p: &Prepared, probe: &mut Probe, clocks: &mut Vec<f64>) -> Result<Vec<f64>, String> {
    let mut loads = Vec::with_capacity(ROUNDS);
    for _ in 0..ROUNDS {
        clocks.push(probe.read());
        let t = Instant::now();
        let m = Machine::load(&p.rows, excel::library());
        let ms = t.elapsed().as_secs_f64() * 1e3;
        let m = m.map_err(|e| format!("{}: the machine refused line {}", p.name, e.line))?;
        black_box(m.formulas());
        loads.push(ms);
    }
    Ok(loads)
}

fn bytes_a_formula(bytes: usize, formulas: usize) -> f64 {
    bytes as f64 / formulas as f64
}

// ---- the runner ----------------------------------------------------------------

fn failed(e: &str) -> ExitCode {
    println!("(memory-failed {e:?})");
    eprintln!("memory: FAILED, {e}");
    ExitCode::FAILURE
}

fn main() -> ExitCode {
    let args: Vec<String> = std::env::args().skip(1).collect();
    let check = args.iter().any(|a| a == "--check");
    if let Some(a) = args.iter().find(|a| a.as_str() != "--check") {
        eprintln!("memory: an unknown argument, {a}");
        return ExitCode::from(2);
    }
    let root: PathBuf = Path::new(env!("CARGO_MANIFEST_DIR")).join("..");
    let read = |p: &str| std::fs::read_to_string(root.join(p));
    let (prelude, english) = match (
        read("scripts/prelude.vla"),
        read("scripts/polyglotta/english.vla"),
    ) {
        (Ok(p), Ok(e)) => (p, e),
        _ => return failed("the prelude or the phrasebook cannot be read"),
    };
    let mut fixtures = Vec::new();
    for name in PROGRAMS {
        match program(&root, name, (&prelude, &english)) {
            Ok(p) => fixtures.push(p),
            Err(e) => return failed(&e),
        }
    }
    for name in ROWS {
        match rows_fixture(&root, name) {
            Ok(p) => fixtures.push(p),
            Err(e) => return failed(&e),
        }
    }

    // The bytes, each fixture weighed once: exact, so once is enough.
    let mut weighed = Vec::new();
    for p in &fixtures {
        match weigh(p) {
            Ok(w) => weighed.push(w),
            Err(e) => return failed(&e),
        }
    }
    if check {
        for (p, w) in fixtures.iter().zip(&weighed) {
            println!(
                "(memory-check \"{}\" (sha256 \"{}\") (formulas {}) (load-held {}) (held {}) (peak {}) (save {}) (save-peak {}) (saved-held {}) (saved-peak {}))",
                p.name, p.sha, w.formulas, w.load_held, w.held, w.peak, w.save.bytes, w.save.printing, w.save.held, w.save.peak
            );
        }
        println!("(memory-check-done {})", fixtures.len());
        return ExitCode::SUCCESS;
    }

    // The loads, timed beside the clock.
    let mut probe = Probe::new(soup(1));
    let readings: Vec<f64> = (0..START_READINGS).map(|_| probe.read()).collect();
    let start_clock = median(&readings);
    let mut loads = Vec::new();
    let mut clocks_of = BTreeMap::new();
    for p in &fixtures {
        let mut clocks = Vec::new();
        match time_loads(p, &mut probe, &mut clocks) {
            Ok(l) => loads.push(l),
            Err(e) => return failed(&e),
        }
        let clock = median(&clocks);
        clocks_of.insert(
            p.name.clone(),
            (clock, spread(&clocks), drift(clock, start_clock)),
        );
    }

    println!(
        "memory: the native runner, {} on {}, {} fixtures, each weighed once and its load timed in {ROUNDS} rounds",
        std::env::consts::ARCH,
        std::env::consts::OS,
        fixtures.len()
    );
    let mut steady_all = true;
    for ((p, w), l) in fixtures.iter().zip(&weighed).zip(&loads) {
        let load = median(l);
        let (clock, sp, dr) = clocks_of[&p.name];
        println!(
            "{}: {} formulas; the load {load:.1} ms, the median of {} ({}), {} formulas a second; {} bytes held after the load ({:.1} a formula), {} after {FRAMES} frames ({:.1}), {} at the peak ({:.1})",
            p.name,
            thousands(w.formulas as u64),
            l.len(),
            l.iter().map(|x| format!("{x:.1}")).collect::<Vec<_>>().join(", "),
            thousands((w.formulas as f64 / (load / 1e3)).round() as u64),
            thousands(w.load_held as u64),
            bytes_a_formula(w.load_held, w.formulas),
            thousands(w.held as u64),
            bytes_a_formula(w.held, w.formulas),
            thousands(w.peak as u64),
            bytes_a_formula(w.peak, w.formulas),
        );
        println!(
            "  its save: {} bytes, {} over the machine to print; a load from it {:.1} ms, {} bytes held, {} at the peak",
            thousands(w.save.bytes as u64),
            thousands(w.save.printing as u64),
            w.save.ms,
            thousands(w.save.held as u64),
            thousands(w.save.peak as u64),
        );
        println!(
            "  the clock: the hand-written Life {clock:.3} ms a generation beside every timed load, spread {sp:.1}%, {dr:+.1}% from {start_clock:.3} at the start"
        );
        if sp > CLOCK_SPREAD_MAX || dr.abs() > CLOCK_SPREAD_MAX {
            steady_all = false;
            println!(
                "  the clock moved during the run, more than {CLOCK_SPREAD_MAX}%: power throttling or a thermal limit; these loads are not a baseline"
            );
        }
    }
    if !steady_all {
        println!("no baseline rows: run again through tools/bench_recalc.ps1, which opts the runner out of Windows' power throttling");
        return ExitCode::from(3);
    }
    println!("the baseline rows, for tools/check_recalc_floors.ps1:");
    for ((p, w), l) in fixtures.iter().zip(&weighed).zip(&loads) {
        let load = median(l);
        let (clock, sp, dr) = clocks_of[&p.name];
        println!(
            "    @{{ Half = 'memory'; Fixture = '{}'; Formulas = {}; Sha256 = '{}'; Rounds = {}; Load = {load:.1}; LoadPerSecond = {}; LoadHeld = {}; Held = {}; Peak = {}; Save = {}; SavePeak = {}; SavedHeld = {}; SavedPeak = {}; ClockMs = {clock:.3}; ClockSpread = {sp:.1}; ClockDrift = {dr:.1} }}",
            p.name,
            w.formulas,
            p.sha,
            l.len(),
            (w.formulas as f64 / (load / 1e3)).round() as u64,
            w.load_held,
            w.held,
            w.peak,
            w.save.bytes,
            w.save.printing,
            w.save.held,
            w.save.peak,
        );
    }
    ExitCode::SUCCESS
}
