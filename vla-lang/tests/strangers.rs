//! Text a stranger wrote (VLA.1): every door of the language that reads
//! text answers it or refuses it by name, and never panics, overflows its
//! stack or hangs. A cartridge, a program, a phrasebook or a workbook's
//! formulas may come from anyone; an engine reads them at its door. The
//! doors: the reader, the formula parser, the reference scanner, a
//! machine's load, a machine's write, and the expander. A grid a load or a
//! write leaves is printed by its census as well, and a grid the machine
//! refuses is planned as a door plans a workbook and its census printed
//! (KERNEL.30).
//!
//! The inputs are generated and mutated from the crate's own fixtures by a
//! seeded generator, the same on every run, and every case runs on a thread
//! of 1 MiB, the tightest stack measured. A case that panics fails with its
//! door, its seed and its input; one that overflows ends the test process,
//! which is the loud failure meant (`VLA_STRANGERS_TRACE=1` prints each case
//! before it runs, so the last one printed is the culprit); one that runs
//! past `SLOW` fails as a hang. Each failure found is shrunk and kept under
//! `testdata/strangers/` as a case of its own, which `the_named_strangers`
//! replays: the first were found before this file was written, by a probe
//! that overflowed the reader, the parser, the evaluator, a machine's plan
//! and the expander, and hung the expander on a macro that doubles.
//!
//! The suite's run is short; `cargo test --release -p vla-lang --test
//! strangers -- --ignored` is the long one.

use std::panic::{catch_unwind, AssertUnwindSafe};
use std::time::{Duration, Instant};

use vla_lang::calc::{formula, library, Extent, Names, Shapes};
use vla_lang::expand::expand_text;
use vla_lang::machine::{read_rows, Machine};
use vla_lang::reader::read_forms;
use vla_lang::refers;
use vla_lang::sheet::Sheet;

const PRELUDE: &str = include_str!("../testdata/prelude.vla");
const GOLDEN: &str = include_str!("../testdata/instructions_golden.vla");
const REFERS: &str = include_str!("../testdata/refers.txt");
const VIEWS: [&str; 5] = [
    include_str!("../testdata/view/fixture_frazaro.vla"),
    include_str!("../testdata/view/fixture_output.vla"),
    include_str!("../testdata/view/fixture_data.vla"),
    include_str!("../testdata/view/into_output.vla"),
    include_str!("../testdata/view/into_checks.vla"),
];

/// The grid a write is made into.
const BASE: &str = "(sheet \"G\" visible)\n(cell \"G\" \"A1\" 1)\n(formula \"G\" \"B1\" \"=A1+1\")\n(formula \"G\" \"C1\" \"=G.last!C1+B1\")\n";

/// How long one case may take before it counts as a hang, generous enough
/// for a debug build on a slow machine.
const SLOW: Duration = Duration::from_secs(5);

/// splitmix64, seeded: every run knocks with the same texts.
struct Rng(u64);

impl Rng {
    fn new(seed: u64) -> Rng {
        Rng(seed)
    }

    fn next(&mut self) -> u64 {
        self.0 = self.0.wrapping_add(0x9E37_79B9_7F4A_7C15);
        let mut z = self.0;
        z = (z ^ (z >> 30)).wrapping_mul(0xBF58_476D_1CE4_E5B9);
        z = (z ^ (z >> 27)).wrapping_mul(0x94D0_49BB_1331_11EB);
        z ^ (z >> 31)
    }

    fn below(&mut self, n: usize) -> usize {
        (self.next() % n.max(1) as u64) as usize
    }
}

/// The doors.
#[derive(Clone, Copy, Debug, PartialEq)]
enum Door {
    Read,
    Formula,
    Refers,
    Load,
    Write,
    Expand,
}

const DOORS: [Door; 6] = [
    Door::Read,
    Door::Formula,
    Door::Refers,
    Door::Load,
    Door::Write,
    Door::Expand,
];

impl Door {
    fn by_name(name: &str) -> Option<Door> {
        DOORS
            .iter()
            .copied()
            .find(|d| format!("{d:?}").eq_ignore_ascii_case(name))
    }
}

/// A refusal is an answer when it has a name and words.
fn named(id: &str, text: &str) -> Result<(), String> {
    if id.is_empty() || text.is_empty() {
        Err(format!(
            "a refusal without a name or words: {id:?} {text:?}"
        ))
    } else {
        Ok(())
    }
}

/// A grid's rows planned as a door plans a workbook, with no twin made, and
/// its census printed (KERNEL.30): what a load refuses, a cycle, a function
/// the library does not hold, a construct not read, a door still prints.
fn door_census(text: &str) {
    let Ok(forms) = read_forms(text) else { return };
    let Ok(draft) = read_rows(&forms) else { return };
    let wb = &draft.workbook();
    let lib = library::language();
    let extents: Vec<Extent> = wb.sheets.iter().map(Sheet::extent).collect();
    let plan = Shapes::of(wb, wb.sheets.len(), &lib, &Names::of(wb), &extents);
    let _ = plan.census(wb, &lib);
}

/// One knock at a door: its answer, or why it is no answer. A machine that
/// loads or takes a write is stepped through a frame and its census printed
/// as well; a grid its load refuses is planned and printed as a door would.
fn knock(door: Door, text: &str) -> Result<(), String> {
    match door {
        Door::Read => match read_forms(text) {
            Ok(_) => Ok(()),
            Err(r) => named(&r.id, &r.text),
        },
        // A label for a construct not read is an answer too.
        Door::Formula => formula::parse(text).map(|_| ()).or(Ok(())),
        Door::Refers => {
            let _ = refers::scan(text);
            Ok(())
        }
        Door::Load => match Machine::load(text, library::language()) {
            Ok(mut m) => {
                m.step(0);
                let _ = m.census();
                Ok(())
            }
            Err(e) => {
                door_census(text);
                named(&e.refusal.id, &e.refusal.text)
            }
        },
        Door::Write => {
            let mut m = Machine::load(BASE, library::language()).expect("the base grid loads");
            m.step(0);
            match m.write(text) {
                Ok(_) => {
                    m.step(0);
                    let _ = m.census();
                    Ok(())
                }
                Err(e) => named(&e.refusal.id, &e.refusal.text),
            }
        }
        Door::Expand => match expand_text(text, "", true) {
            Ok(_) => Ok(()),
            Err(r) => named(&r.id, &r.text),
        },
    }
}

/// Knock with every case on a thread of 1 MiB; a panic or a slow answer is
/// a failure naming the case, an overflow the end of the process.
fn knock_all(cases: Vec<(Door, String, String)>) {
    let trace = std::env::var_os("VLA_STRANGERS_TRACE").is_some();
    std::thread::Builder::new()
        .stack_size(1 << 20)
        .spawn(move || {
            let mut failures = Vec::new();
            for (door, what, text) in &cases {
                if trace {
                    eprintln!("{door:?} {what}: {text:?}");
                }
                let started = Instant::now();
                let answer = catch_unwind(AssertUnwindSafe(|| knock(*door, text)));
                let took = started.elapsed();
                let failure = match answer {
                    Err(_) => Some("panicked".to_string()),
                    Ok(Err(why)) => Some(why),
                    Ok(Ok(())) if took > SLOW => Some(format!("took {took:?}, a hang")),
                    Ok(Ok(())) => None,
                };
                if let Some(why) = failure {
                    let shown: String = text.chars().take(300).collect();
                    failures.push(format!("{door:?} {what}: {why}\n  input: {shown:?}"));
                }
            }
            assert!(failures.is_empty(), "{}", failures.join("\n"));
        })
        .expect("a thread")
        .join()
        .expect("every stranger answered or refused by name");
}

// ---- the generator ------------------------------------------------------

/// What an insertion is made of: the delimiters, the escapes, the
/// operators, the reference marks, and a few characters nothing expects.
const ALPHABET: &[char] = &[
    '(', ')', '"', '\\', ' ', '\n', '\r', '\t', ';', '=', '!', ':', '$', '#', '%', '&', '+', '-',
    '*', '/', '^', '<', '>', ',', '.', '{', '}', '[', ']', '@', '\'', 'A', 'z', '1', '9', '0', 'é',
    'λ', '\u{0}', '\u{feff}',
];

/// What a deep run is made of.
const RUNS: &[&str] = &["(", ")", "-", "%", "ABS(", "(+ 1 ", "\"", "1+"];

/// A window of a fixture, where it starts chosen at random.
fn excerpt(r: &mut Rng, text: &str, most: usize) -> String {
    let chars: Vec<char> = text.chars().collect();
    let start = r.below(chars.len());
    let len = 1 + r.below(most);
    chars[start..(start + len).min(chars.len())]
        .iter()
        .collect()
}

/// One to four mutations of a text.
fn mutate(r: &mut Rng, text: &str) -> String {
    let mut c: Vec<char> = text.chars().collect();
    for _ in 0..1 + r.below(4) {
        let at = r.below(c.len() + 1);
        match r.below(7) {
            0 if !c.is_empty() => {
                let to = (at + 1 + r.below(20)).min(c.len());
                c.drain(at.min(c.len())..to);
            }
            1 => {
                for _ in 0..1 + r.below(10) {
                    c.insert(at.min(c.len()), ALPHABET[r.below(ALPHABET.len())]);
                }
            }
            2 if !c.is_empty() => {
                let from = r.below(c.len());
                let to = (from + 1 + r.below(50)).min(c.len());
                let span: Vec<char> = c[from..to].to_vec();
                let at = at.min(c.len());
                c.splice(at..at, span);
            }
            3 if !c.is_empty() => {
                let i = r.below(c.len());
                c[i] = ALPHABET[r.below(ALPHABET.len())];
            }
            4 => {
                let run = RUNS[r.below(RUNS.len())].repeat(1 + r.below(200));
                let at = at.min(c.len());
                c.splice(at..at, run.chars());
            }
            5 if !c.is_empty() => c.truncate(at.min(c.len())),
            _ if c.len() > 1 => {
                let (i, j) = (r.below(c.len()), r.below(c.len()));
                c.swap(i, j);
            }
            _ => {}
        }
    }
    c.into_iter().collect()
}

/// The formulas the refers fixture holds, one a line beginning `=`.
fn formulas() -> Vec<&'static str> {
    REFERS.lines().filter(|l| l.starts_with('=')).collect()
}

/// The programs the expander is knocked with before they are mutated.
const PROGRAMS: &[&str] = &[
    "(defmacro (twice x) (begin x x)) (twice (debug-print 1))",
    "(defmacro (m a & rest) (list a rest)) (m 1 2 3)",
    "(defmacro (deep10) (a (a (a (a (a (a (a (a (a (a 1))))))))))) (((((deep10)))))",
    "(defmacro (walk lst) (quote-if (null? lst) (begin) (begin (debug-print (car lst)) (walk (cdr lst))))) (walk (quote (1 2 3)))",
    "(sub main () (dim x Long) (set! x (+ 1 2)) (debug-print x))",
    // VLA.3: the expand-time object model, a picture and an address's arithmetic.
    "(defmacro (pic r t) (pic-from r 1 (characters-expand t))) (defmacro (pic-from r c m) (quote-if (null? m) (begin) (begin (quote-if (eq? (car m) \"#\") (cell \"S\" (address-expand r c) 1) (begin)) (pic-from r (+expand c 1) (cdr m))))) (pic 2 \".#\") (&expand (row-expand \"B3\") (column-letters-expand (mod-expand 30 26)) (quotient-expand 9 (-expand 4 (*expand 1 2))) (column-expand \"$C$4\"))",
    // VLA.5: the loop in a template and out of one, a count, a loop inside a
    // loop, a loop where a value is read, and a piece found by its head.
    "(defmacro (cells sheet & entries) (each (entry entries) (cell sheet (car entry) (car (cdr entry))))) (cells \"S\" (\"A1\" 1) (\"B1\" 2)) (each-index (i x (quote (a b)) 0) (each (y (cdr (quote (p q r)))) (f i x y))) (+ (each (c (quote (A1 B1))) c)) (car (each (z (quote (u v))) (g z))) (assoc-expand (quote b) (quote ((a 1) (b 2))))",
    // VLA.6: defines of a reference and of a ref, a define a macro writes, a
    // twin of each, and a row at a name.
    "(defmacro (name-cell word sheet address) (define word (ref sheet address))) (define stop State!B12) (define press (ref \"State\" \"B9\")) (formula-expand (if (and (= (last stop) 1) (= press 1)) 1 0)) (formula (sheet-expand stop) (range-expand stop) (formula-expand (+ stop 1))) (name-cell level \"State\" \"B3\") (formula-expand (last level))",
];

/// The rows a write is made of before they are mutated.
const WRITES: &[&str] = &[
    "(cell \"G\" \"A1\" 5)",
    "(formula \"G\" \"D1\" \"=A1*2+SUM(A1:C1)\")",
    "(cell \"G\" \"E5\" \"text\")",
    "(formula \"G\" \"F1\" \"=IF(G.last!F1>3,0,G.last!F1+1)\")",
];

/// `n` cases at one door from one seed.
fn cases_at(door: Door, seed: u64, n: usize) -> Vec<(Door, String, String)> {
    let mut r = Rng::new(seed);
    let refers = formulas();
    (0..n)
        .map(|i| {
            let seed_text = match door {
                Door::Read => match r.below(3) {
                    0 => excerpt(&mut r, GOLDEN, 400),
                    1 => excerpt(&mut r, PRELUDE, 400),
                    _ => VIEWS[r.below(VIEWS.len())].to_string(),
                },
                Door::Formula | Door::Refers => refers[r.below(refers.len())].to_string(),
                Door::Load => VIEWS[r.below(VIEWS.len())].to_string(),
                Door::Write => WRITES[r.below(WRITES.len())].to_string(),
                Door::Expand => PROGRAMS[r.below(PROGRAMS.len())].to_string(),
            };
            let text = mutate(&mut r, &seed_text);
            (door, format!("seed {seed} case {i}"), text)
        })
        .collect()
}

#[test]
fn strangers_at_every_door_are_answered_or_refused_by_name() {
    let mut cases = Vec::new();
    for door in DOORS {
        let n = if door == Door::Load || door == Door::Write {
            150
        } else {
            400
        };
        cases.extend(cases_at(door, 1, n));
    }
    knock_all(cases);
}

#[test]
#[ignore = "the long run: cargo test --release -p vla-lang --test strangers -- --ignored"]
fn strangers_at_length() {
    for seed in 2..=11u64 {
        let mut cases = Vec::new();
        for door in DOORS {
            cases.extend(cases_at(door, seed, 5_000));
        }
        knock_all(cases);
    }
}

// ---- the named strangers ------------------------------------------------

/// Cases chosen by hand: each limit's edges, the inputs nothing expects,
/// and the shapes the probe crashed on, then every case kept under
/// `testdata/strangers/`, one a file, its first line `door: <name>` and
/// the rest the input.
#[test]
fn the_named_strangers() {
    let deep = |n: usize, open: &str, close: &str| format!("{}{}", open.repeat(n), close.repeat(n));
    let mut cases: Vec<(Door, String, String)> = vec![
        (Door::Read, "empty".into(), String::new()),
        (Door::Read, "a lone quote".into(), "\"".into()),
        (Door::Read, "an escape at the end".into(), "\"abc\\".into()),
        (Door::Read, "a hundred thousand closings".into(), ")".repeat(100_000)),
        (Door::Read, "a flat list of a hundred thousand".into(), format!("({})", "1 ".repeat(100_000))),
        (Door::Read, "NUL and BOM".into(), "(a\u{0}b \u{feff})".into()),
        (Door::Formula, "only the sign".into(), "=".into()),
        (Door::Formula, "an opening alone".into(), "=(".into()),
        (Door::Formula, "percents past the limit".into(), format!("=1{}", "%".repeat(8000))),
        (Door::Formula, "signs past the limit".into(), format!("={}1", "-".repeat(8000))),
        (Door::Formula, "calls past the limit".into(), format!("={}1{}", "ABS(".repeat(1000), ")".repeat(1000))),
        (Door::Refers, "a long formula".into(), format!("={}", "A1+".repeat(20_000))),
        (Door::Refers, "a quote never closed".into(), "='Sheet 1!A1".into()),
        (Door::Load, "rows nested past the limit".into(), deep(100_000, "(", ")")),
        (
            Door::Load,
            "a sum over the whole sheet with two cells".into(),
            "(cell \"G\" \"A1\" 1)\n(cell \"G\" \"XFD1048576\" 2)\n(formula \"G\" \"B1\" \"=SUM(A1:XFD1048576)\")\n".into(),
        ),
        (Door::Write, "a formula past the limit".into(), format!("(formula \"G\" \"D1\" \"={}\")", deep(65, "(", ")"))),
        (Door::Write, "a formula past the length".into(), format!("(formula \"G\" \"D1\" \"={}1\")", "1+".repeat(5000))),
        (Door::Expand, "a macro that calls itself unchanged".into(), "(defmacro (spin x) (spin x)) (spin 1)".into()),
        (Door::Expand, "a macro with no template".into(), "(defmacro (m))".into()),
        // VLA.5: a malformed loop is refused by name, and a long one answers.
        (Door::Expand, "a loop alone".into(), "(each)".into()),
        (Door::Expand, "a loop whose binding is an atom".into(), "(each x)".into()),
        (Door::Expand, "a loop whose binding is empty".into(), "(each () (f))".into()),
        (Door::Expand, "a loop of a hundred thousand".into(), format!("(each (x (quote ({}))) x)", "1 ".repeat(100_000))),
        (
            Door::Expand,
            "a loop past a million forms in one application".into(),
            format!("(defmacro (m & xs) (each (x xs) (f x x x x))) (m {})", "1 ".repeat(200_000)),
        ),
        // VLA.6: a malformed define and a malformed last are refused by name,
        // and a chain of words past 5,000 links meets the chain's limit.
        (Door::Expand, "a define alone".into(), "(define)".into()),
        (Door::Expand, "a define whose word is a list".into(), "(define (a) 1)".into()),
        (Door::Expand, "a last alone".into(), "(formula-expand (last))".into()),
        (Door::Expand, "a last of an empty ref".into(), "(formula-expand (last (ref)))".into()),
        (Door::Expand, "a reference's sheet alone".into(), "(sheet-expand)".into()),
        (
            Door::Expand,
            "a chain of six thousand words".into(),
            format!(
                "{} (f link-1)",
                (1..=6000)
                    .map(|i| format!("(define link-{i} link-{})", i + 1))
                    .collect::<Vec<_>>()
                    .join(" ")
            ),
        ),
    ];
    let dir = concat!(env!("CARGO_MANIFEST_DIR"), "/testdata/strangers");
    let mut kept: Vec<_> = std::fs::read_dir(dir)
        .expect("testdata/strangers")
        .map(|e| e.expect("an entry").path())
        .collect();
    kept.sort();
    assert!(
        kept.len() >= 8,
        "{} cases kept, fewer than were found",
        kept.len()
    );
    for path in kept {
        let text = std::fs::read_to_string(&path).expect("a kept case");
        let (first, input) = text.split_once('\n').expect("a door line, then the input");
        let name = first
            .trim_end_matches('\r')
            .strip_prefix("door: ")
            .expect("the first line names the door");
        let door = Door::by_name(name).unwrap_or_else(|| panic!("{path:?}: no door {name}"));
        let what = path.file_name().unwrap().to_string_lossy().into_owned();
        cases.push((door, what, input.to_string()));
    }
    knock_all(cases);
}
