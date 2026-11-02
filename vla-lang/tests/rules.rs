//! A conditional format's laws (KERNEL.58, its third slice), over generated
//! grids as the language's own laws are (`tests/laws.rs`), and its rows.
//!
//! 1. A rule moves no value: a grid's values at every frame are the same
//!    grid's without its `conditional` rows, every sheet's and twin's, which
//!    is `SD-32`: a look computes nothing and the logic depends on no look.
//! 2. Each `lit` row is the answers of each rule's test written as ordinary
//!    formulas on ordinary sheets of a twin grid, walked by priority up to the
//!    first that stops, at every frame, through value and formula writes.
//! 3. A window's values at the frame before the last complete one, patched by
//!    its changes, are its values now, as `KERNEL.42`'s law holds them, a
//!    cell's `lit` row among its rows.
//! 4. A grid viewed whole and loaded again views the same, its rules and its
//!    answers kept; and resumed and stepped, steps as the original does.
//! 5. A frame stepped in slices of any budget is the frame stepped whole, a
//!    view during a yield the last complete frame's.
//!
//! The grids hold merged ranges (the item's sixth slice), whose rows are among
//! those law 1 takes away, so that a merge moves no value and no plan, and whose
//! cells each draw by their first in law 2, and through laws 3 to 5 with it.

use std::collections::BTreeMap;

use vla_lang::calc::library;
use vla_lang::calc::{Computed, Names, Value, Values};
use vla_lang::form::Form;
use vla_lang::machine::{read_rows, Machine, Viewed};
use vla_lang::reader::read_forms;
use vla_lang::refers::shift_a1_references;
use vla_lang::sheet::{cell_ref, parse_a1_range, A1Range};
use vla_lang::view::{view_text, window_of};

/// splitmix64, seeded, so every run writes the same grids.
struct Rng(u64);

impl Rng {
    fn next(&mut self) -> u64 {
        self.0 = self.0.wrapping_add(0x9E37_79B9_7F4A_7C15);
        let mut z = self.0;
        z = (z ^ (z >> 30)).wrapping_mul(0xBF58_476D_1CE4_E5B9);
        z = (z ^ (z >> 27)).wrapping_mul(0x94D0_49BB_1331_11EB);
        z ^ (z >> 31)
    }

    fn below(&mut self, n: usize) -> usize {
        (self.next() % n as u64) as usize
    }

    fn pick<'a, T>(&mut self, xs: &'a [T]) -> &'a T {
        &xs[self.below(xs.len())]
    }
}

fn load(rows: &str, what: &str) -> Machine {
    Machine::load(rows, library::language())
        .unwrap_or_else(|e| panic!("{what}: line {}: {}\n{rows}", e.line, e.refusal.text))
}

fn record(m: &Machine, projection: &str, sheet: &str, window: Option<&str>) -> String {
    match m.view(projection, sheet, window) {
        Ok(Viewed::Text(t)) => t,
        other => panic!("{projection} of {sheet}: {other:?}"),
    }
}

/// The formats a generated rule draws, and the one the cells hold.
const FORMATS: &str = concat!(
    "(format 1 (fill \"FFE699\"))\n",
    "(format 2 (font bold))\n",
    "(format 3 (fill none) (font \"800020\"))\n",
    "(format 4 (number \"0.00\"))\n",
    "(format 5 \"CCFFCC\" \"0.0\" nowrap)\n",
);

/// A rule's test as a generated grid writes it, `@` where a reference to the
/// rule's own sheet names none in its row and `G!` in the twin grid's formula.
/// A reference to `H`, to a twin, and none at all, are the same in both.
const EXPRESSIONS: &[&str] = &[
    "@A1>2",
    "@$B1=1",
    "@A$1<3",
    "@$C$2>1",
    "H!A1=1",
    "H!$B$2>1",
    "G.last!D1>2",
    "MOD(ROW(),2)=0",
    "MOD(COLUMN(),2)=1",
    "@D1>@A1",
    "SUM(@A1:B2)>5",
    "COUNTIF(@$A$1:$C$5,@A1)>1",
    "@A1",
    "1/@E1>0",
    "TRUE",
    "0",
];

/// A generated rule: its row's parts, and what its row and the twin grid's
/// formula are written from.
struct Generated {
    priority: u32,
    rectangles: Vec<A1Range>,
    /// The test's text, read at the corner, as the twin grid's formula holds
    /// it; and how the rule's row spells the test.
    formula: String,
    row_test: String,
    format: String,
    stop: bool,
}

impl Generated {
    fn row(&self) -> String {
        let range: Vec<String> = self.rectangles.iter().map(A1Range::text).collect();
        format!(
            "(conditional \"G\" {} \"{}\" {} {}{})\n",
            self.priority,
            range.join(" "),
            self.row_test,
            self.format,
            if self.stop { " stop" } else { "" }
        )
    }

    fn corner(&self) -> (u32, u32) {
        (
            self.rectangles.iter().map(|r| r.top).min().unwrap(),
            self.rectangles.iter().map(|r| r.left).min().unwrap(),
        )
    }
}

fn rectangle(top: u32, left: u32, bottom: u32, right: u32) -> A1Range {
    A1Range {
        top,
        left,
        bottom,
        right,
    }
}

/// One rule over a rectangle or two that do not meet, its second above or
/// beside its first as often as below.
fn generated(r: &mut Rng, priority: u32) -> Generated {
    let top = 1 + r.below(6) as u32;
    let left = 1 + r.below(6) as u32;
    let first = rectangle(top, left, top + r.below(3) as u32, left + r.below(2) as u32);
    let mut rectangles = vec![first];
    if r.below(3) == 0 {
        // A rectangle clear of the first, in the rows past it.
        let row = first.bottom + 1 + r.below(2) as u32;
        let col = 1 + r.below(8) as u32;
        let second = rectangle(row, col, row + r.below(2) as u32, col);
        match r.below(2) {
            0 => rectangles.push(second),
            _ => rectangles.insert(0, second),
        }
    }
    let (row_test, formula) = match r.below(5) {
        0 => {
            let (word, op) = *r.pick(&[("less", "<"), ("greater", ">"), ("equal", "=")]);
            let bound = 1 + r.below(4);
            let corner = {
                let t = rectangles.iter().map(|x| x.top).min().unwrap();
                let l = rectangles.iter().map(|x| x.left).min().unwrap();
                cell_ref(t, l)
            };
            (
                format!("(cell {word} \"={bound}\")"),
                format!("G!{corner}{op}{bound}"),
            )
        }
        1 => {
            let corner = {
                let t = rectangles.iter().map(|x| x.top).min().unwrap();
                let l = rectangles.iter().map(|x| x.left).min().unwrap();
                cell_ref(t, l)
            };
            let (low, high) = *r.pick(&[(1, 3), (3, 1), (2, 2)]);
            let (a, b) = (low.min(high), low.max(high));
            (
                format!("(cell between \"={low}\" \"={high}\")"),
                format!("AND(G!{corner}>={a},G!{corner}<={b})"),
            )
        }
        _ => {
            let text = *r.pick(EXPRESSIONS);
            (
                format!("\"={}\"", text.replace('@', "")),
                text.replace('@', "G!"),
            )
        }
    };
    let format = match r.below(6) {
        0 => "none".to_string(),
        n => n.min(4).to_string(),
    };
    Generated {
        priority,
        rectangles,
        formula,
        row_test,
        format,
        stop: r.below(4) == 0,
    }
}

/// The merged ranges a generated grid may hold, no two overlapping: over its
/// values, its formulas and its empty cells, each inside [`ALL`].
const MERGES: [&str; 6] = ["A1:B2", "C3:D3", "F2:F4", "H6:J7", "E9:G9", "B8:C10"];

/// A generated grid's rows of cells and formulas, its rules and its merged ranges.
fn generated_grid(r: &mut Rng) -> (String, Vec<Generated>, Vec<A1Range>) {
    let mut rows = String::from("(sheet \"G\" visible)\n");
    rows.push_str(FORMATS);
    for row in 1..=5 {
        for col in ["A", "B", "C"] {
            rows.push_str(&format!("(cell \"G\" \"{col}{row}\" {})\n", r.below(6)));
        }
    }
    rows.push_str("(formula \"G\" \"D1:D5\" \"=MOD(G.last!D1+A1,5)\")\n");
    rows.push_str("(formula \"G\" \"E1:E5\" \"=MOD(G.last!E1+1,4)\")\n");
    rows.push_str("(style \"G\" \"A1:C1\" 5)\n");
    rows.push_str("(sheet \"H\" visible)\n");
    rows.push_str(&format!("(cell \"H\" \"A1\" {})\n", r.below(3)));
    rows.push_str(&format!("(cell \"H\" \"B2\" {})\n", r.below(4)));
    let count = 1 + r.below(5);
    let rules = (0..count)
        .map(|k| generated(r, 1 + 2 * k as u32))
        .collect::<Vec<_>>();
    let merges = MERGES
        .iter()
        .filter(|_| r.below(3) == 0)
        .map(|range| parse_a1_range(range).expect("a range"))
        .collect();
    (rows, rules, merges)
}

/// The grid with its rules' rows and its merged ranges'.
fn with_rules(rows: &str, rules: &[Generated], merges: &[A1Range]) -> String {
    let mut out = rows.to_string();
    for rule in rules {
        out.push_str(&rule.row());
    }
    for merge in merges {
        out.push_str(&format!("(merge \"G\" \"{}\")\n", merge.text()));
    }
    out
}

/// The twin grid: each rule's test written as ordinary formulas over its
/// rectangles, on an ordinary sheet of its own at the same places, moved as
/// a fill moves one from the corner.
fn with_twin_sheets(rows: &str, rules: &[Generated]) -> String {
    let mut out = rows.to_string();
    for rule in rules {
        let (top, left) = rule.corner();
        for rect in &rule.rectangles {
            let moved = shift_a1_references(
                &rule.formula,
                i64::from(rect.top) - i64::from(top),
                i64::from(rect.left) - i64::from(left),
            );
            out.push_str(&format!(
                "(formula \"R{}\" \"{}\" \"=IFERROR(IF({moved},1,0),0)\")\n",
                rule.priority,
                rect.text()
            ));
        }
    }
    out
}

/// The writes between two frames, the same to both grids: a value at most
/// frames, and a formula now and then, which plans the grid again.
fn writes(r: &mut Rng, frame: u64) -> Vec<String> {
    let mut out = Vec::new();
    if r.below(4) != 0 {
        let at = cell_ref(1 + r.below(5) as u32, 1 + r.below(3) as u32);
        out.push(format!("(cell \"G\" \"{at}\" {})", r.below(6)));
    }
    if r.below(3) == 0 {
        out.push(format!("(cell \"H\" \"A1\" {})", r.below(3)));
    }
    if frame % 4 == 3 {
        out.push(format!("(formula \"G\" \"B3\" \"=A3+{}\")", r.below(3)));
    }
    out
}

/// A record less the rows a rule may add, its formats among them, and its
/// `text` rows, which a rule's number format moves and no value does.
fn less_the_rules(record: &str) -> String {
    record
        .lines()
        .filter(|l| {
            !l.starts_with("(conditional ")
                && !l.starts_with("(merge ")
                && !l.starts_with("(lit ")
                && !l.starts_with("(format ")
                && !l.starts_with("(text ")
        })
        .map(|l| format!("{l}\n"))
        .collect()
}

/// The `lit` rows of a record, by cell.
fn lit_rows(record: &str) -> BTreeMap<(u32, u32), Vec<u32>> {
    let mut lit = BTreeMap::new();
    for line in record.lines().filter(|l| l.starts_with("(lit ")) {
        let forms = read_forms(line).expect("a row");
        let Some(Form::List(l)) = forms.first() else {
            panic!("{line}");
        };
        let [_, _, Form::Str(addr), priorities @ ..] = l.items.as_slice() else {
            panic!("{line}");
        };
        let at = parse_a1_range(addr).expect("a cell");
        let listed: Vec<u32> = priorities
            .iter()
            .map(|p| match p {
                Form::Sym(n) => n.parse().expect("a priority"),
                _ => panic!("{line}"),
            })
            .collect();
        lit.insert((at.top, at.left), listed);
    }
    lit
}

/// What each cell's rules draw as the twin grid computes them: by priority,
/// each true rule listed, up to the first that stops, and a merged range's cells
/// each its first's.
fn lit_by_hand(
    twin: &Machine,
    rules: &[Generated],
    merges: &[A1Range],
) -> BTreeMap<(u32, u32), Vec<u32>> {
    let mut lit: BTreeMap<(u32, u32), Vec<u32>> = BTreeMap::new();
    let mut stopped: Vec<(u32, u32)> = Vec::new();
    for rule in rules {
        let sheet = twin
            .find_sheet(&format!("R{}", rule.priority))
            .expect("a rule's ordinary sheet");
        for rect in &rule.rectangles {
            for row in rect.top..=rect.bottom {
                for col in rect.left..=rect.right {
                    let answer = twin.computed((sheet, row, col));
                    let on = matches!(
                        answer,
                        Some(Computed::Value(Value::Number(n))) if n != 0.0
                    );
                    if on && !stopped.contains(&(row, col)) {
                        lit.entry((row, col)).or_default().push(rule.priority);
                        if rule.stop {
                            stopped.push((row, col));
                        }
                    }
                }
            }
        }
    }
    for merge in merges {
        let first = lit.get(&(merge.top, merge.left)).cloned();
        for row in merge.top..=merge.bottom {
            for col in merge.left..=merge.right {
                match &first {
                    Some(first) => lit.insert((row, col), first.clone()),
                    None => lit.remove(&(row, col)),
                };
            }
        }
    }
    lit
}

/// The window every generated rule stands in.
const ALL: &str = "A1:J12";

#[test]
fn a_rule_moves_no_value_and_each_lit_row_is_its_test_s_answer() {
    let mut lit_rows_held = 0;
    let mut stops_met = 0;
    let mut drawn_by_a_first = 0;
    for seed in 1..=60u64 {
        let mut r = Rng(seed);
        let (rows, rules, merges) = generated_grid(&mut r);
        let ruled_rows = with_rules(&rows, &rules, &merges);
        let what = format!("seed {seed}");
        let mut ruled = load(&ruled_rows, &what);
        let mut plain = load(&rows, &what);
        let mut twin = load(&with_twin_sheets(&rows, &rules), &what);
        assert_eq!(ruled.sheets(), plain.sheets(), "{what}: the rows' sheets");
        assert_eq!(ruled.loaded(), plain.loaded(), "{what}: the load's counts");
        assert_eq!(ruled.formulas(), plain.formulas(), "{what}: the rows' own");
        for frame in 1..=6u64 {
            for w in writes(&mut r, frame) {
                for m in [&mut ruled, &mut plain, &mut twin] {
                    m.write(&w)
                        .unwrap_or_else(|e| panic!("{what}: {w}: {}", e.refusal.text));
                }
            }
            for m in [&mut ruled, &mut plain, &mut twin] {
                assert!(m.step(0).done);
            }
            // Law 1, every sheet and twin of the rows, in every projection.
            for sheet in ["G", "H", "G.last", "H.last"] {
                for projection in ["grid", "values", "fills"] {
                    assert_eq!(
                        less_the_rules(&record(&ruled, projection, sheet, None)),
                        less_the_rules(&record(&plain, projection, sheet, None)),
                        "{what}, frame {frame}: {projection} of {sheet}\n{ruled_rows}"
                    );
                }
            }
            // Law 2.
            let lit = lit_rows(&record(&ruled, "values", "G", Some(ALL)));
            let want = lit_by_hand(&twin, &rules, &merges);
            assert_eq!(lit, want, "{what}, frame {frame}\n{ruled_rows}");
            lit_rows_held += lit.len();
            let own = lit_by_hand(&twin, &rules, &[]);
            drawn_by_a_first += want.iter().filter(|(at, l)| own.get(at) != Some(l)).count()
                + own.keys().filter(|at| !want.contains_key(at)).count();
            stops_met += lit
                .values()
                .filter(|l| {
                    l.iter()
                        .any(|p| rules.iter().any(|x| x.priority == *p && x.stop))
                })
                .count();
        }
    }
    assert!(lit_rows_held >= 1_000, "too few lit rows: {lit_rows_held}");
    assert!(stops_met >= 300, "too few stops met: {stops_met}");
    assert!(
        drawn_by_a_first >= 100,
        "too few cells drawn otherwise than they answer: {drawn_by_a_first}"
    );
}

/// The cell a record's row belongs to, a `lit` row among them.
fn cell_of(line: &str) -> Option<(u32, u32)> {
    let forms = read_forms(line).ok()?;
    let Some(Form::List(l)) = forms.first() else {
        return None;
    };
    match l.items.as_slice() {
        [Form::Sym(head), Form::Str(_), Form::Str(addr), ..]
            if matches!(
                head.as_str(),
                "cell" | "value" | "text" | "style" | "lit" | "sentence" | "empty"
            ) =>
        {
            let r = parse_a1_range(addr)?;
            Some((r.top, r.left))
        }
        _ => None,
    }
}

/// A values record patched by a window's changes, as a host patches the record
/// it holds; `None` for changes answered whole.
fn patched(record: &str, changes: &str) -> Option<String> {
    let mut head = String::new();
    let mut cells: BTreeMap<(u32, u32), String> = BTreeMap::new();
    for line in record.lines() {
        match cell_of(line) {
            Some(cell) => cells
                .entry(cell)
                .or_default()
                .push_str(&format!("{line}\n")),
            None => head.push_str(&format!("{line}\n")),
        }
    }
    let mut lines = changes.lines();
    if lines.next()?.ends_with(" whole)") {
        return None;
    }
    let mut given: BTreeMap<(u32, u32), String> = BTreeMap::new();
    for line in lines {
        let cell = cell_of(line).unwrap_or_else(|| panic!("a changes row names a cell: {line}"));
        let rows = given.entry(cell).or_default();
        if !line.starts_with("(empty ") {
            rows.push_str(&format!("{line}\n"));
        }
    }
    for (cell, rows) in given {
        match rows.is_empty() {
            true => cells.remove(&cell),
            false => cells.insert(cell, rows),
        };
    }
    Some(head + &cells.into_values().collect::<String>())
}

#[test]
fn a_window_s_values_patched_by_its_changes_are_its_values_now_under_rules() {
    let windows: &[(&str, Option<&str>)] = &[
        ("G", None),
        ("G", Some(ALL)),
        ("G", Some("B2:D4")),
        ("G", Some("A3:F3")),
        ("G", Some("F6:J12")),
        ("G.last", None),
        ("H", None),
    ];
    let mut held = 0;
    let mut lit_given = 0;
    for seed in 1..=40u64 {
        let mut r = Rng(1000 + seed);
        let (rows, rules, merges) = generated_grid(&mut r);
        let ruled_rows = with_rules(&rows, &rules, &merges);
        let what = format!("seed {seed}");
        let mut m = load(&ruled_rows, &what);
        let budget = [0, 1, 7][seed as usize % 3];
        let mut base: Option<Vec<String>> = None;
        let values = |m: &Machine| -> Vec<String> {
            windows
                .iter()
                .map(|(s, w)| record(m, "values", s, *w))
                .collect()
        };
        let mut at = values(&m);
        let hold = |m: &Machine, was: &[String], when: &str, held: &mut usize, lit: &mut usize| {
            for ((sheet, window), was) in windows.iter().zip(was) {
                let changes = record(m, "changes", sheet, *window);
                let now = record(m, "values", sheet, *window);
                if let Some(got) = patched(was, &changes) {
                    assert_eq!(
                        got, now,
                        "{what}, {when}: {sheet} {window:?} patched by\n{changes}\n{ruled_rows}"
                    );
                    *held += 1;
                    *lit += changes.lines().filter(|l| l.starts_with("(lit ")).count();
                }
            }
        };
        for frame in 1..=8u64 {
            // A value alone: a formula written makes the changes whole.
            let moved = cell_ref(1 + r.below(5) as u32, 1 + r.below(3) as u32);
            m.write(&format!("(cell \"G\" \"{moved}\" {})", r.below(6)))
                .unwrap();
            if let Some(b) = &base {
                hold(&m, b, "after a write", &mut held, &mut lit_given);
            }
            if !m.step(budget).done {
                if let Some(b) = &base {
                    hold(&m, b, "during a frame", &mut held, &mut lit_given);
                }
                while !m.step(budget).done {}
            }
            let was = std::mem::replace(&mut at, values(&m));
            hold(
                &m,
                &was,
                &format!("frame {frame}"),
                &mut held,
                &mut lit_given,
            );
            base = Some(was);
        }
    }
    assert!(held >= 500, "too few windows patched: {held}");
    assert!(lit_given >= 300, "too few lit rows given: {lit_given}");
}

/// A machine's save: the grid view of every sheet its rows name.
fn save(m: &Machine) -> String {
    (0..m.sheets())
        .filter_map(|i| m.sheet_name(i))
        .map(|name| record(m, "grid", name, None))
        .collect()
}

#[test]
fn a_grid_viewed_whole_and_loaded_again_views_the_same_and_steps_the_same() {
    let mut rules_saved = 0;
    let mut lit_saved = 0;
    for seed in 1..=40u64 {
        let mut r = Rng(2000 + seed);
        let (rows, rules, merges) = generated_grid(&mut r);
        let ruled_rows = with_rules(&rows, &rules, &merges);
        let what = format!("seed {seed}");
        let mut m = load(&ruled_rows, &what);
        for frame in 0..=3u64 {
            let saved = save(&m);
            let mut again = load(&saved, &format!("{what}, frame {frame}'s save"));
            assert_eq!(save(&again), saved, "{what}, frame {frame}\n{ruled_rows}");
            rules_saved += saved
                .lines()
                .filter(|l| l.starts_with("(conditional "))
                .count();
            lit_saved += saved.lines().filter(|l| l.starts_with("(lit ")).count();
            again.resume_at(frame).unwrap();
            assert!(again.step(0).done);
            assert!(m.step(0).done);
            assert_eq!(
                save(&again),
                save(&m),
                "{what}, frame {frame} resumed and stepped\n{ruled_rows}"
            );
        }
    }
    assert!(rules_saved >= 100, "too few rules saved: {rules_saved}");
    assert!(lit_saved >= 200, "too few lit rows saved: {lit_saved}");
}

#[test]
fn a_frame_stepped_in_slices_is_the_frame_stepped_whole_and_a_yield_shows_the_last() {
    for seed in 1..=30u64 {
        let mut r = Rng(3000 + seed);
        let (rows, rules, merges) = generated_grid(&mut r);
        let ruled_rows = with_rules(&rows, &rules, &merges);
        let what = format!("seed {seed}");
        let mut whole = load(&ruled_rows, &what);
        let mut sliced = load(&ruled_rows, &what);
        for frame in 1..=4u64 {
            let moved = format!("(cell \"G\" \"A{}\" {})", 1 + r.below(5), r.below(6));
            whole.write(&moved).unwrap();
            sliced.write(&moved).unwrap();
            let shown = record(&sliced, "values", "G", Some(ALL));
            assert!(whole.step(0).done);
            let mut yields = 0;
            while !sliced.step(1).done {
                yields += 1;
                assert_eq!(
                    record(&sliced, "values", "G", Some(ALL)),
                    shown,
                    "{what}, frame {frame}, yield {yields}: the last complete frame's\n{ruled_rows}"
                );
            }
            assert!(yields > 0, "{what}: a frame of no yield");
            assert_eq!(
                record(&sliced, "grid", "G", Some(ALL)),
                record(&whole, "grid", "G", Some(ALL)),
                "{what}, frame {frame}\n{ruled_rows}"
            );
            assert_eq!(sliced.changed(), whole.changed(), "{what}, frame {frame}");
        }
    }
}

// ---- the rows and what a grid says of them ------------------------------

/// A grid of still values whose rule reads a sheet a write moves.
const STILL: &str = concat!(
    "(sheet \"Board\" visible)\n",
    "(format 64 (fill \"FFE699\"))\n",
    "(cell \"Board\" \"A1\" 1)\n",
    "(cell \"Board\" \"B1\" 2)\n",
    "(cell \"Board\" \"C1\" 3)\n",
    "(cell \"Paint\" \"A1\" 0)\n",
    "(conditional \"Board\" 1 \"A1:C1\" \"=Paint!$A$1=1\" 64)\n",
);

#[test]
fn a_rule_s_answer_moves_where_no_value_does() {
    let mut m = load(STILL, "the board");
    let board = m.find_sheet("Board").unwrap();
    let paint = m.find_sheet("Paint").unwrap();
    assert!(m.step(0).done);
    let still = record(&m, "values", "Board", None);
    assert!(!still.contains("(lit "), "{still}");

    // The next frame: nothing moved, nothing named.
    assert!(m.step(0).done);
    assert_eq!(m.changed(), 0);
    let quiet = record(&m, "changes", "Board", None);
    assert!(quiet.ends_with(" 0)\n"), "{quiet}");

    // A write moves a sheet the rule reads, and no value on the Board.
    m.write("(cell \"Paint\" \"A1\" 1)").unwrap();
    assert!(m.step(0).done);
    let lit = record(&m, "values", "Board", None);
    assert!(
        lit.contains("(cell \"Board\" \"A1\" 1)\n(lit \"Board\" \"A1\" 1)\n"),
        "{lit}"
    );
    assert_eq!(m.changed(), 3, "the three answers");
    assert_eq!(m.changed_on(board), Some(3), "counted on the rule's sheet");
    assert_eq!(
        m.changed_on(paint),
        Some(0),
        "and not on the sheet that moved"
    );
    assert!(
        record(&m, "changes", "Board.last", None).ends_with(" 0)\n"),
        "an answer that moved is no change of the sheet's twin"
    );
    let changes = record(&m, "changes", "Board", None);
    assert_eq!(
        changes,
        concat!(
            "(changes \"Board\" \"A1:C1\" 2 3 3)\n",
            "(cell \"Board\" \"A1\" 1)\n(lit \"Board\" \"A1\" 1)\n",
            "(cell \"Board\" \"B1\" 2)\n(lit \"Board\" \"B1\" 1)\n",
            "(cell \"Board\" \"C1\" 3)\n(lit \"Board\" \"C1\" 1)\n",
        )
    );

    // The next frame still: the answers held, nothing to do.
    assert!(m.step(0).done);
    assert_eq!(m.changed(), 0);
    assert!(record(&m, "changes", "Board", None).ends_with(" 0)\n"));
    assert_eq!(record(&m, "values", "Board", None), lit);

    // And back.
    m.write("(cell \"Paint\" \"A1\" 0)").unwrap();
    assert!(m.step(0).done);
    assert_eq!(m.changed_on(board), Some(3));
    assert_eq!(record(&m, "values", "Board", None), still);
}

#[test]
fn a_lit_cell_holding_nothing_prints_alone_and_empties_when_it_goes_out() {
    let rows = concat!(
        "(cell \"S\" \"B2\" 1)\n",
        "(cell \"S\" \"A5\" 0)\n",
        "(format 7 (font bold))\n",
        "(conditional \"S\" 1 \"A1:B3\" \"=$A$5=1\" 7)\n",
    );
    let mut m = load(rows, "the sheet");
    assert!(m.step(0).done);
    m.write("(cell \"S\" \"A5\" 1)").unwrap();
    assert!(m.step(0).done);
    let lit = record(&m, "values", "S", Some("A1:B3"));
    let cells: Vec<&str> = lit
        .lines()
        .filter(|l| l.starts_with("(lit ") || l.starts_with("(cell "))
        .collect();
    assert_eq!(
        cells,
        [
            "(lit \"S\" \"A1\" 1)",
            "(lit \"S\" \"B1\" 1)",
            "(lit \"S\" \"A2\" 1)",
            "(cell \"S\" \"B2\" 1)",
            "(lit \"S\" \"B2\" 1)",
            "(lit \"S\" \"A3\" 1)",
            "(lit \"S\" \"B3\" 1)",
        ]
    );
    m.write("(cell \"S\" \"A5\" 0)").unwrap();
    assert!(m.step(0).done);
    let changes = record(&m, "changes", "S", Some("A1:B3"));
    assert_eq!(
        changes,
        concat!(
            "(changes \"S\" \"A1:B3\" 2 3 6)\n",
            "(empty \"S\" \"A1\")\n",
            "(empty \"S\" \"B1\")\n",
            "(empty \"S\" \"A2\")\n",
            "(cell \"S\" \"B2\" 1)\n",
            "(empty \"S\" \"A3\")\n",
            "(empty \"S\" \"B3\")\n",
        )
    );
}

#[test]
fn a_record_prints_a_rule_after_its_looks_and_its_format_among_its_formats() {
    let rows = concat!(
        "(sheet \"D\" visible)\n",
        "(format 66 (font bold))\n",
        "(format 67 (fill \"FFE699\"))\n",
        "(look \"D\" 9 67)\n",
        "(cell \"D\" \"A1\" 6)\n",
        // Written out of order, and past the sheet's extent.
        "(conditional \"D\" 4 \"A1:A3 C1:C3\" (cell between \"=1\" \"=5\") none stop)\n",
        "(conditional \"D\" 3 \"A1:A9\" (cell greater \"=5\") 66)\n",
    );
    let mut m = load(rows, "the sheet");
    assert!(m.step(0).done);
    let got = record(&m, "grid", "D", Some("A1"));
    let lines: Vec<&str> = got.lines().collect();
    let at = lines.iter().position(|l| l.starts_with("(look ")).unwrap();
    assert_eq!(
        lines[at..],
        [
            "(look \"D\" 9 67)",
            "(conditional \"D\" 3 \"A1:A9\" (cell greater \"=5\") 66)",
            "(conditional \"D\" 4 \"A1:A3 C1:C3\" (cell between \"=1\" \"=5\") none stop)",
            "(cell \"D\" \"A1\" 6)",
            "(lit \"D\" \"A1\" 3)",
        ]
    );
    assert!(got.contains("(format 66 (font bold))\n"), "{got}");
    // A rule's answer sheet is listed by no record.
    assert!(!got.contains("D[3]"), "{got}");
    assert!(got.contains("(sheet \"D.last\" hidden)\n"), "{got}");
    assert_eq!(m.sheets(), 1);
}

#[test]
fn a_rule_s_number_format_draws_a_cell_s_text_over_a_look_and_its_own() {
    let rows = concat!(
        "(format 1 \"FFFFFF\" \"0.0\" nowrap)\n",
        "(format 2 (number \"0.000\"))\n",
        "(format 3 (number \"0.00\"))\n",
        "(format 4 (fill \"FFE699\"))\n",
        "(cell \"N\" \"A1\" 12.25)\n",
        "(style \"N\" \"A1\" 1)\n",
        "(cell \"N\" \"A2\" 12.25)\n",
        "(style \"N\" \"A2\" 1)\n",
        "(cell \"N\" \"A3\" 12.25)\n",
        "(style \"N\" \"A3\" 1)\n",
        "(cell \"N\" \"B1\" 0)\n",
        "(look \"N\" 12.25 3)\n",
    );
    // Under no rule, the look's code; a rule's over it; a rule saying none
    // leaves the next that does.
    let ruled = format!(
        "{rows}{}",
        concat!(
            "(conditional \"N\" 1 \"A2\" \"=TRUE\" 4)\n",
            "(conditional \"N\" 2 \"A2\" \"=TRUE\" 2)\n",
            "(conditional \"N\" 3 \"A3\" \"=FALSE\" 2)\n",
        )
    );
    let mut m = load(&ruled, "the sheet");
    assert!(m.step(0).done);
    let got = record(&m, "values", "N", None);
    assert!(got.contains("(text \"N\" \"A1\" \"12.25\")\n"), "{got}");
    assert!(got.contains("(text \"N\" \"A2\" \"12.250\")\n"), "{got}");
    assert!(got.contains("(text \"N\" \"A3\" \"12.25\")\n"), "{got}");
}

#[test]
fn a_stop_ends_every_later_rule_over_the_cell_whatever_its_range() {
    // The owner's look at Excel's own drawing of the `Order` sheet, where its
    // object model's `DisplayFormat` said otherwise: `C` stops over `C:E`, and
    // `C:E` over `C`.
    let rows = concat!(
        "(format 1 (fill \"9DC3E6\"))\n",
        "(format 2 (font bold))\n",
        "(cell \"O\" \"E9\" 0)\n",
        "(conditional \"O\" 1 \"C1\" \"=1\" 1 stop)\n",
        "(conditional \"O\" 2 \"C1:E1\" \"=1\" 2)\n",
        "(conditional \"O\" 3 \"C2:E2\" \"=1\" 1 stop)\n",
        "(conditional \"O\" 4 \"C2\" \"=1\" 2)\n",
        "(conditional \"O\" 5 \"C3\" \"=1\" none stop)\n",
        "(conditional \"O\" 6 \"C3:E3\" \"=1\" 1)\n",
        // A false rule's stop stops nothing.
        "(conditional \"O\" 7 \"C4\" \"=0\" 1 stop)\n",
        "(conditional \"O\" 8 \"C4\" \"=1\" 2)\n",
    );
    let mut m = load(rows, "the sheet");
    assert!(m.step(0).done);
    let lit = lit_rows(&record(&m, "values", "O", Some("A1:F5")));
    let want: BTreeMap<(u32, u32), Vec<u32>> = [
        ((1, 3), vec![1]),
        ((1, 4), vec![2]),
        ((1, 5), vec![2]),
        ((2, 3), vec![3]),
        ((2, 4), vec![3]),
        ((2, 5), vec![3]),
        ((3, 3), vec![5]),
        ((3, 4), vec![6]),
        ((3, 5), vec![6]),
        ((4, 3), vec![8]),
    ]
    .into_iter()
    .collect();
    assert_eq!(lit, want);
}

#[test]
fn a_formula_moves_from_the_corner_of_its_range_whatever_the_order() {
    // Excel's `Bases` sheet: the lower left rectangle first, where the corner,
    // `D10`, is no cell of the range.
    let rows = concat!(
        "(cell \"B\" \"H14\" 1)\n",
        "(cell \"B\" \"H15\" 1)\n",
        "(cell \"B\" \"J12\" 1)\n",
        "(cell \"B\" \"J13\" 1)\n",
        // What the first rectangle's first cell, `D12`, would light and not.
        "(cell \"B\" \"H12\" 1)\n",
        "(cell \"B\" \"H13\" 0)\n",
        "(cell \"B\" \"J10\" 0)\n",
        "(cell \"B\" \"J11\" 1)\n",
        "(conditional \"B\" 1 \"D12:D13 F10:F11\" \"=H12=1\" none)\n",
    );
    let mut m = load(rows, "the sheet");
    assert!(m.step(0).done);
    let lit = lit_rows(&record(&m, "values", "B", Some("A1:K20")));
    let cells: Vec<(u32, u32)> = lit.keys().copied().collect();
    assert_eq!(cells, [(10, 6), (11, 6), (12, 4), (13, 4)]);
}

#[test]
fn a_bound_reading_a_cell_holding_nothing_is_the_number_0() {
    let rows = concat!(
        "(cell \"C\" \"A1\" 0)\n",
        "(cell \"C\" \"B1\" false)\n",
        "(cell \"C\" \"C1\" true)\n",
        "(cell \"C\" \"D1\" \"abc\")\n",
        "(formula \"C\" \"E1\" \"=\\\"\\\"\")\n",
        "(cell \"C\" \"G1\" 7)\n",
        "(conditional \"C\" 1 \"A1:F1\" (cell equal \"=$Z$9\") none)\n",
        "(conditional \"C\" 2 \"A2:F2\" (cell equal \"=$G$1\") none)\n",
    );
    let mut m = load(rows, "the sheet");
    assert!(m.step(0).done);
    let lit = lit_rows(&record(&m, "values", "C", Some("A1:G2")));
    // The number 0 and the cell holding nothing equal it; `FALSE`, a text and
    // the empty text do not, as a comparison reading the reference would say.
    let cells: Vec<(u32, u32)> = lit.keys().copied().collect();
    assert_eq!(cells, [(1, 1), (1, 6)], "{lit:?}");
}

#[test]
fn a_rule_reading_where_it_stands_reads_its_own_row_and_column() {
    let rows = concat!(
        "(cell \"W\" \"A9\" 0)\n",
        "(conditional \"W\" 1 \"A2:B3\" \"=MOD(ROW(),2)=0\" none)\n",
        "(conditional \"W\" 2 \"D2:E2\" \"=MOD(COLUMN(),2)=1\" none)\n",
    );
    let mut m = load(rows, "the sheet");
    assert!(m.step(0).done);
    let lit = lit_rows(&record(&m, "values", "W", Some("A1:F9")));
    let cells: Vec<(u32, u32)> = lit.keys().copied().collect();
    assert_eq!(cells, [(2, 1), (2, 2), (2, 5)]);
}

#[test]
fn a_rule_finds_a_name_from_its_own_sheet() {
    let rows = concat!(
        "(cell \"E\" \"A1\" 1)\n",
        "(cell \"E\" \"A2\" 0)\n",
        "(cell \"N\" \"B9\" 0)\n",
        "(name \"pick\" \"E!$A$2\")\n",
        "(name \"N!pick\" \"E!$A$1\")\n",
        "(conditional \"N\" 1 \"A1\" \"=pick=1\" none)\n",
        "(conditional \"E\" 1 \"C1\" \"=pick=1\" none)\n",
    );
    let mut m = load(rows, "the sheet");
    assert!(m.step(0).done);
    assert_eq!(lit_rows(&record(&m, "values", "N", Some("A1"))).len(), 1);
    assert_eq!(lit_rows(&record(&m, "values", "E", Some("C1"))).len(), 0);
}

#[test]
fn a_rule_s_rows_are_the_rows_own_counts_and_its_cells_a_frame_s_work() {
    let rows = concat!(
        "(cell \"G\" \"A1\" 1)\n",
        "(formula \"G\" \"B1\" \"=A1+1\")\n",
        "(conditional \"G\" 1 \"A1:C3\" \"=A1>0\" none)\n",
        "(conditional \"G\" 2 \"E1 E3\" \"=A1>0\" none)\n",
    );
    let mut m = load(rows, "the sheet");
    assert_eq!(m.sheets(), 1);
    assert_eq!(m.loaded().sheets, 1);
    assert_eq!(m.loaded().cells, 2);
    assert_eq!(m.loaded().formulas, 1);
    assert_eq!(m.formulas().0, 1);
    assert_eq!(m.formulas().1, 1, "the rows' own shape");
    let step = m.step(0);
    assert_eq!(step.of, 1 + 9 + 2, "a frame's work holds the rules' cells");
    assert_eq!(m.reached(), 12);
    // The answers' sheets are found by no name, read by no write and viewed by
    // no window.
    assert_eq!(m.find_sheet("G[1]"), None);
    assert_eq!(m.find_sheet("G[1].last"), None);
    assert!(m.view("values", "G[1]", None).is_err());
    let asked = m.view("values", "Nope", None).expect_err("no such sheet");
    assert!(!asked.text.contains("G[1]"), "{}", asked.text);
    assert!(asked.text.contains("G, G.last"), "{}", asked.text);
    let unknown = m.write("(cell \"G[1]\" \"A1\" 1)").unwrap_err().refusal;
    assert_eq!(unknown.id, "grid-sheet-unknown");
    assert!(!unknown.text.contains("G[1],"), "{}", unknown.text);
    // The grid lent holds the rule on its sheet and no answer sheet.
    let wb = m.workbook();
    let names: Vec<&str> = wb.sheets.iter().map(|s| s.name.as_str()).collect();
    assert_eq!(names, ["G", "G.last"]);
    assert_eq!(wb.sheets[0].rules().len(), 2);
}

#[test]
fn a_door_prints_a_rule_s_row_and_no_answer() {
    let rows = "(cell \"G\" \"A1\" 1)\n(conditional \"G\" 1 \"A1\" \"=A1>0\" none)\n";
    let forms = read_forms(rows).unwrap();
    let wb = read_rows(&forms).unwrap().workbook();
    let window = window_of(&wb, "G", None).unwrap();
    let got = view_text(&wb, &window);
    assert!(
        got.contains("(conditional \"G\" 1 \"A1\" \"=A1>0\" none)\n"),
        "{got}"
    );
    assert!(!got.contains("(lit "), "{got}");
}

#[test]
fn a_draft_lists_a_rule_s_formulas_among_its_own() {
    let rows = concat!(
        "(formula \"G\" \"A1\" \"=1\")\n",
        "(conditional \"G\" 2 \"B2\" (cell between \"=TODAY()\" \"=2\") none)\n",
        "(conditional \"G\" 3 \"C3 B4\" \"=NOW()\" none)\n",
    );
    let forms = read_forms(rows).unwrap();
    let draft = read_rows(&forms).unwrap();
    let texts: Vec<(&str, u32)> = draft.formulas().map(|(_, t, line)| (t, line)).collect();
    assert_eq!(
        texts,
        [("1", 1), ("TODAY()", 2), ("2", 2), ("NOW()", 3)],
        "by sheet and first cell"
    );
}

#[test]
fn a_rule_row_is_refused_by_what_it_gets_wrong() {
    let bad = |row: &str, id: &str| {
        let refusal = Machine::load(row, library::language())
            .err()
            .unwrap_or_else(|| panic!("{row} was taken"));
        assert_eq!(refusal.refusal.id, id, "{row}: {}", refusal.refusal.text);
    };
    let malformed = "grid-row-malformed";
    for row in [
        "(conditional \"S\" 1 \"A1\" \"=1\")",
        "(conditional \"S\" one \"A1\" \"=1\" none)",
        "(conditional \"S\" 1 \"A1,B1\" \"=1\" none)",
        "(conditional \"S\" 1 \"\" \"=1\" none)",
        "(conditional \"S\" 1 \"A1\" \"1\" none)",
        "(conditional \"S\" 1 \"A1\" \"=\" none)",
        "(conditional \"S\" 1 \"A1\" \"=1\" bold)",
        "(conditional \"S\" 1 \"A1\" \"=1\" 65536)",
        "(conditional \"S\" 1 \"A1\" \"=1\" none go)",
        "(conditional \"S\" 1 \"A1\" (cell above \"=1\") none)",
        "(conditional \"S\" 1 \"A1\" (cell equal) none)",
        "(conditional \"S\" 1 \"A1\" (cell equal \"=1\" \"=2\") none)",
        "(conditional \"S\" 1 \"A1\" (cell between \"=1\") none)",
        "(conditional \"S\" 1 \"A1\" (formula \"=1\") none)",
    ] {
        bad(row, malformed);
    }
    // A twin holds no rule.
    bad(
        "(cell \"S\" \"A1\" 1)\n(conditional \"S.last\" 1 \"A1\" \"=1\" none)",
        "grid-sheet-name-invalid",
    );
    // A name no sheet may hold, an answer sheet's among them.
    bad(
        "(conditional \"S[1]\" 1 \"A1\" \"=1\" none)",
        "grid-sheet-name-invalid",
    );
    // A second row of one priority taken alike and refused unlike.
    let one = "(conditional \"S\" 1 \"A1\" \"=1\" none)\n";
    assert!(Machine::load(&format!("{one}{one}"), library::language()).is_ok());
    let refused = Machine::load(
        &format!("{one}(conditional \"S\" 1 \"A1\" \"=2\" none)"),
        library::language(),
    )
    .expect_err("refused");
    assert_eq!(refused.refusal.id, "grid-cell-written-twice");
    assert_eq!(refused.line, 2);
    // A `lit` row: of a rule held, at a cell it answers for.
    for row in [
        "(lit \"S\" \"A1\" 1)",
        "(cell \"S\" \"A1\" 1)\n(lit \"S\" \"A1\" 1)",
        "(conditional \"S\" 1 \"B1\" \"=1\" none)\n(lit \"S\" \"A1\" 1)",
        "(conditional \"S\" 1 \"A1\" \"=1\" none)\n(lit \"S\" \"A1\")",
        "(conditional \"S\" 1 \"A1\" \"=1\" none)\n(lit \"S\" \"A1:B1\" 1)",
        "(conditional \"S\" 1 \"A1\" \"=1\" none)\n(lit \"S\" \"A1\" one)",
    ] {
        bad(row, malformed);
    }
    bad(
        concat!(
            "(conditional \"S\" 1 \"A1\" \"=1\" none)\n(conditional \"S\" 2 \"A1\" \"=1\" none)\n",
            "(lit \"S\" \"A1\" 1)\n(lit \"S\" \"A1\" 2)",
        ),
        "grid-cell-written-twice",
    );
    // A function the library does not hold, as in a cell, at the rule's line.
    let refused = Machine::load(
        "(cell \"S\" \"A1\" 1)\n(conditional \"S\" 1 \"A1\" \"=NOSUCH(A1)\" none)",
        library::language(),
    )
    .expect_err("refused");
    assert_eq!(refused.line, 2, "{}", refused.refusal.text);
}

#[test]
fn a_rule_s_cells_stand_beside_the_grid_s_at_most_the_cap() {
    // The grid's cap and the rules' are two.
    let fits = concat!(
        "(cell \"S\" \"A1:B524288\" 1)\n",
        "(conditional \"S\" 1 \"A1:B524288\" \"=1\" none)\n",
    );
    let m = load(fits, "the sheet");
    assert_eq!(m.loaded().cells, 1_048_576);
    // Past the room each answers inside its sheet's extent; past it there too, by name.
    let refused = Machine::load(
        concat!(
            "(cell \"S\" \"A1:C349525\" 1)\n",
            "(conditional \"S\" 1 \"A1:C349525\" \"=1\" none)\n",
            "(conditional \"S\" 2 \"A1:A2\" \"=1\" none)\n",
        ),
        library::language(),
    )
    .expect_err("refused");
    assert_eq!(refused.refusal.id, "grid-rule-cells");
    assert_eq!(refused.line, 3);
    assert!(
        refused.refusal.text.contains("S!A1:A2"),
        "{}",
        refused.refusal.text
    );
    assert!(
        refused.refusal.text.contains("1,048,577"),
        "{}",
        refused.refusal.text
    );
    assert!(
        refused.refusal.text.contains("narrower"),
        "{}",
        refused.refusal.text
    );
}

#[test]
fn many_answers_moving_at_once_are_named_where_the_grid_holds_few_cells() {
    // A frame's end keeps room for half the cells, a rule's answers among them,
    // past which it answers whole: a hundred answers moving of five hundred, in a
    // grid of one cell, are named.
    let rows = concat!(
        "(cell \"S\" \"A1\" 0)\n",
        "(conditional \"S\" 1 \"B1:B100\" \"=$A$1=1\" none)\n",
        "(conditional \"S\" 2 \"D1:D400\" \"=FALSE\" none)\n",
    );
    let mut m = load(rows, "the sheet");
    assert!(m.step(0).done);
    assert!(m.step(0).done);
    m.write("(cell \"S\" \"A1\" 1)").unwrap();
    assert!(m.step(0).done);
    let changes = record(&m, "changes", "S", Some("A1:B100"));
    let head = changes.lines().next().unwrap();
    assert_eq!(head, "(changes \"S\" \"A1:B100\" 2 3 101)", "{changes}");
    assert_eq!(changes.lines().count(), 1 + 1 + 100);
}

#[test]
fn a_window_s_changes_print_what_a_stop_leaves() {
    // Rule 1 stops over rule 2, each turning with the same cell: the cell's
    // changes print what its values do.
    let rows = concat!(
        "(cell \"S\" \"A1\" 0)\n",
        "(cell \"S\" \"A2\" 0)\n",
        "(conditional \"S\" 1 \"A1:B1\" \"=$A$2=1\" none stop)\n",
        "(conditional \"S\" 2 \"A1:B1\" \"=$A$2=1\" none)\n",
    );
    let mut m = load(rows, "the sheet");
    assert!(m.step(0).done);
    m.write("(cell \"S\" \"A2\" 1)").unwrap();
    assert!(m.step(0).done);
    assert!(m.step(0).done);
    m.write("(cell \"S\" \"A2\" 0)").unwrap();
    assert!(m.step(0).done);
    m.write("(cell \"S\" \"A2\" 1)").unwrap();
    assert!(m.step(0).done);
    let changes = record(&m, "changes", "S", Some("A1:B1"));
    assert_eq!(
        changes,
        concat!(
            "(changes \"S\" \"A1:B1\" 4 5 2)\n",
            "(cell \"S\" \"A1\" 0)\n(lit \"S\" \"A1\" 1)\n",
            "(lit \"S\" \"B1\" 1)\n",
        )
    );
}

#[test]
fn a_rectangle_overlapping_one_before_it_answers_once() {
    let rows = concat!(
        "(cell \"S\" \"A9\" 1)\n",
        "(conditional \"S\" 1 \"A1:B2 B2:C3\" \"=$A$9=1\" none)\n",
    );
    let mut m = load(rows, "the sheet");
    let step = m.step(0);
    assert_eq!(step.of, 7, "the cell they share computed once");
    let lit = lit_rows(&record(&m, "values", "S", Some("A1:D4")));
    assert_eq!(lit.len(), 7);
}

/// What a reader asks of a rule before it prints it is what a load answers of it:
/// a function the library lacks, in a formula or a bound; a text the parse does
/// not read, alone or inside the answer a load makes of it; a defined name the
/// parse does not read, at any remove, by its scope; and a reference moved off the
/// sheet at any rectangle's cell. Each case is held to the load itself.
#[test]
fn a_rule_is_computable_exactly_where_a_load_takes_it() {
    let months = "(sheet \"Jan\" visible)\n(sheet \"Dec\" visible)\n";
    // Under the parse's 8,192 characters alone, and past them inside the answer.
    let long = format!("\"A1\" \"=1{}\"", "+1".repeat(4_092));
    // At the parse's 64 levels alone, and past them inside the answer.
    let deep = format!("\"A1\" \"={}1{}\"", "ABS(".repeat(64), ")".repeat(64));
    let deep_enough = format!("\"A1\" \"={}1{}\"", "ABS(".repeat(62), ")".repeat(62));
    let cases: Vec<(&str, &str, bool)> = vec![
        ("", "\"A1\" \"=A1=1\"", true),
        ("", "\"A1\" \"=IFERROR(A1,0)=1\"", true),
        ("", "\"A1\" \"=_xlfn.IFNA(A1,0)=1\"", true),
        ("", "\"A1\" (cell greater \"=1\")", true),
        ("", "\"A1\" (cell between \"=1\" \"=A1\")", true),
        ("", "\"A1\" \"=NOSUCH(A1)=1\"", false),
        ("", "\"A1\" \"=FLOOR(A1,1)=TODAY()\"", false),
        ("", "\"A1\" \"=IF(A1,NOW(),0)\"", false),
        ("", "\"A1\" (cell greater \"=NOW()\")", false),
        ("", "\"A1\" (cell between \"=1\" \"=NOW()\")", false),
        ("", "\"A1\" \"=A1#\"", false),
        ("", "\"A1\" \"=[Rates.xlsx]Sheet1!A1=1\"", false),
        ("", "\"A1\" \"=[1]Sheet1!A1=1\"", false),
        ("", "\"A1\" \"=#REF!B1048576=1\"", true),
        ("", "\"A1:A2\" \"=#REF!B1048576=1\"", false),
        ("", "\"A1\" (cell greater \"=A1#\")", false),
        ("", &long, false),
        ("", &deep, false),
        ("", &deep_enough, true),
        ("(name \"Fine\" \"A1\")\n", "\"A1\" \"=Fine=1\"", true),
        ("(name \"Odd\" \"A1#\")\n", "\"A1\" \"=A1=1\"", true),
        ("(name \"Odd\" \"A1#\")\n", "\"A1\" \"=Odd=1\"", false),
        (
            "(name \"Far\" \"Odd\")\n(name \"Odd\" \"A1#\")\n",
            "\"A1\" \"=Far=1\"",
            false,
        ),
        (
            "(sheet \"S\" visible)\n(sheet \"T\" visible)\n(name \"S!Odd\" \"A1#\")\n",
            "\"A1\" \"=Odd=1\"",
            false,
        ),
        (
            "(sheet \"S\" visible)\n(sheet \"T\" visible)\n(name \"T!Odd\" \"A1#\")\n",
            "\"A1\" \"=Odd=1\"",
            true,
        ),
        (months, "\"A1\" \"=SUM(Jan:Dec!B1048576)>0\"", true),
        (months, "\"A1:A2\" \"=SUM(Jan:Dec!B1048576)>0\"", false),
        (months, "\"A1 A3\" \"=SUM(Jan:Dec!B1048576)>0\"", false),
        (months, "\"A3 A1\" \"=SUM(Jan:Dec!B1048575)>0\"", false),
        (months, "\"A1:A2\" \"=SUM(Jan:Dec!B1048575)>0\"", true),
        (months, "\"A1:A2\" \"=SUM(Jan:Dec!$B$1048576)>0\"", true),
        (months, "\"A1:XFD1\" \"=SUM(Jan:Dec!A1)>0\"", true),
        (months, "\"A1:XFD1\" \"=SUM(Jan:Dec!B1)>0\"", false),
    ];
    for (before, rule, taken) in cases {
        let rows = format!("{before}(conditional \"S\" 1 {rule} none)\n");
        let what = &rows[..rows.len().min(160)];
        let load = Machine::load(&rows, library::language());
        assert_eq!(
            load.as_ref().err().map(|e| e.refusal.id.as_str()).is_none(),
            taken,
            "{what}: {:?}",
            load.as_ref().err().map(|e| e.refusal.id.as_str())
        );
        if let Err(refused) = &load {
            let id = refused.refusal.id.as_str();
            assert!(
                [
                    "calc-function-not-computed",
                    "calc-construct-not-read",
                    "calc-formula-too-long",
                    "calc-formula-too-deep",
                ]
                .contains(&id),
                "{what}: {id}"
            );
        }
        let wb = read_rows(&read_forms(&rows).unwrap()).unwrap().workbook();
        let home = wb.sheets.iter().position(|s| s.name == "S").unwrap();
        let names = Names::of(&wb);
        let rule = &wb.sheets[home].rules()[0];
        assert_eq!(
            rule.computable(&wb, &names, home, &library::language()),
            taken,
            "{what}"
        );
        assert!(
            !rule.computable(&wb, &names, wb.sheets.len(), &library::language()),
            "{what}: a sheet the workbook does not hold"
        );
    }
}

/// Rules over a whole column, a third of the room's: the grid's cells and the rules' rows.
const PAST_THE_ROOM: &str = concat!(
    "(cell \"S\" \"A1\" 1)\n",
    "(cell \"S\" \"A2\" 2)\n",
    "(cell \"S\" \"A3\" 1)\n",
    "(cell \"S\" \"B2\" 2)\n",
    "(conditional \"S\" 1 \"A1:A1048576\" \"=A1=1\" none)\n",
    "(conditional \"S\" 2 \"B1:B1048576\" \"=A1=2\" none)\n",
);

/// The rules a window's record lights, each its cell and its priorities.
fn lit_in(m: &Machine, window: &str) -> Vec<String> {
    record(m, "values", "S", Some(window))
        .lines()
        .filter(|l| l.starts_with("(lit "))
        .map(str::to_string)
        .collect()
}

/// Rules asking more than the room are each planned only inside their sheet's extent:
/// so that two whole columns load, cost what the sheet holds and draw as they do there,
/// and the cells past it, which no record of the sheet whole shows, answer nothing.
#[test]
fn rules_past_the_room_answer_inside_their_sheet_s_extent() {
    let mut m = load(PAST_THE_ROOM, "two whole columns");
    let step = m.step(0);
    assert!(step.done);
    assert_eq!(step.of, 3 + 3, "each rule's cells of A1:B3, the extent");
    assert_eq!(
        lit_in(&m, "A1:B9"),
        [
            "(lit \"S\" \"A1\" 1)",
            "(lit \"S\" \"B2\" 2)",
            "(lit \"S\" \"A3\" 1)",
        ]
    );
    // A cell the plan left out holds no answer a `lit` row could say.
    let said = format!("{PAST_THE_ROOM}(lit \"S\" \"A9\" 1)\n");
    let held = load(&said, "a lit row past the extent");
    assert_eq!(lit_in(&held, "A9:A9"), Vec::<String>::new());
    // What the sheet's whole record prints is the rules' rows as written.
    let whole = record(&m, "values", "S", None);
    assert!(whole.contains("\"A1:A1048576\""), "{whole}");
    assert!(whole.contains("\"B1:B1048576\""), "{whole}");
    // One column alone fits the room, and answers for every cell of it as it did: a rule
    // true of an empty cell lights it past the extent, which past the room it does not.
    let empty =
        "(cell \"S\" \"A1\" 1)\n(conditional \"S\" 1 \"A1:A1048576\" \"=A1=\\\"\\\"\" none)\n";
    let mut whole = load(empty, "one whole column");
    assert_eq!(whole.step(0).of, 1_048_576);
    assert_eq!(lit_in(&whole, "A1:A3").len(), 2, "its two empty cells");
    let both = format!("{empty}(conditional \"S\" 2 \"B1:B1048576\" \"=1\" none)\n");
    let mut past = load(&both, "two");
    assert_eq!(past.step(0).of, 1);
    assert_eq!(lit_in(&past, "A1:B3"), Vec::<String>::new());
}

/// A rule planned inside its sheet's extent follows it as writes grow it: a value typed
/// below the data is answered for, on every rule of its sheet and of no other, and a
/// refused write takes the plan back with it.
#[test]
fn a_rule_past_the_room_follows_its_sheet_s_extent() {
    let mut m = load(PAST_THE_ROOM, "two whole columns");
    assert!(m.step(0).done);
    m.write("(cell \"S\" \"A6\" 1)").unwrap();
    assert!(m.step(0).done);
    assert_eq!(
        lit_in(&m, "A1:B9"),
        [
            "(lit \"S\" \"A1\" 1)",
            "(lit \"S\" \"B2\" 2)",
            "(lit \"S\" \"A3\" 1)",
            "(lit \"S\" \"A6\" 1)",
        ]
    );
    m.write("(cell \"S\" \"A7\" 2)").unwrap();
    assert!(m.step(0).done);
    assert!(lit_in(&m, "A1:B9").contains(&"(lit \"S\" \"B7\" 2)".to_string()));
    // A formula's write grows it too, and its answer reads the formula's value.
    m.write("(formula \"S\" \"A8\" \"=A6\")").unwrap();
    while !m.step(0).done {}
    while !m.step(0).done {}
    assert!(lit_in(&m, "A1:B9").contains(&"(lit \"S\" \"A8\" 1)".to_string()));
    // A rule that planned nothing, its sheet's cells nowhere near its column, grows
    // from its own corner: B5 reads A5 as it would have.
    let far = concat!(
        "(cell \"S\" \"A5\" 2)\n",
        "(conditional \"S\" 1 \"A1:A1048576\" \"=A1=9\" none)\n",
        "(conditional \"S\" 2 \"B1:B1048576\" \"=A1=2\" none)\n",
    );
    let mut grown = load(far, "a rule of no cell yet");
    assert_eq!(grown.step(0).of, 1);
    grown.write("(cell \"S\" \"B6\" 0)").unwrap();
    assert!(grown.step(0).done);
    assert_eq!(lit_in(&grown, "A1:B9"), ["(lit \"S\" \"B5\" 2)"]);
    // A cell on another sheet grows nothing of this one.
    let other = format!("{PAST_THE_ROOM}(cell \"T\" \"A1\" 1)\n");
    let mut m = load(&other, "two sheets");
    let of = m.step(0).of;
    m.write("(cell \"T\" \"A9\" 1)").unwrap();
    assert_eq!(m.step(0).of, of, "the plan stands");
    // A write refused takes the plan back: the same grid as before it, at every frame.
    let mut refused = load(PAST_THE_ROOM, "refused");
    let mut stood = load(PAST_THE_ROOM, "stood");
    assert!(refused.step(0).done && stood.step(0).done);
    let e = refused
        .write("(formula \"S\" \"A9\" \"=A9\")")
        .expect_err("a cycle");
    assert_eq!(e.refusal.id, "calc-cycle");
    assert_eq!(
        record(&refused, "values", "S", None),
        record(&stood, "values", "S", None)
    );
    assert_eq!(refused.loaded().formulas, stood.loaded().formulas);
    refused.write("(cell \"S\" \"A5\" 1)").unwrap();
    stood.write("(cell \"S\" \"A5\" 1)").unwrap();
    assert_eq!(refused.step(0).of, stood.step(0).of);
    assert_eq!(lit_in(&refused, "A1:B9"), lit_in(&stood, "A1:B9"));
}

/// What a machine past the room holds is what a load of its record holds, through writes
/// that grow its sheets: its answers, its changes and what it steps from there.
#[test]
fn a_grid_past_the_room_loaded_again_views_the_same() {
    for seed in 0..12 {
        let mut rng = Rng(0x5EED_0000 + seed);
        let mut rows = String::from("(cell \"S\" \"A1\" 1)\n(cell \"T\" \"A1\" 1)\n");
        for (k, column) in ["A", "B", "C"].iter().enumerate() {
            let test = rng.pick(&[
                "=A1=1",
                "=$A1>1",
                "=A1=T!$A$1",
                "=B2=1",
                "=ISBLANK(A1)",
                "=1",
            ]);
            rows.push_str(&format!(
                "(conditional \"S\" {} \"{column}1:{column}1048576\" \"{test}\" none {})\n",
                k + 1,
                *rng.pick(&["", "stop"]),
            ));
        }
        rows = rows.replace(" none )", " none)");
        let mut m = load(&rows, "a grid past the room");
        assert!(m.step(0).done);
        let mut frame = 1;
        for _ in 0..8 {
            let (row, col) = (1 + rng.below(12) as u32, 1 + rng.below(4) as u32);
            let cell = cell_ref(row, col);
            let write = match rng.below(3) {
                0 => format!("(cell \"S\" \"{cell}\" {})", rng.below(3)),
                1 => format!("(formula \"S\" \"{cell}\" \"=1+{}\")", rng.below(3)),
                _ => format!("(cell \"T\" \"{cell}\" {})", rng.below(3)),
            };
            m.write(&write).unwrap();
            assert!(m.step(0).done && m.step(0).done);
            frame += 2;
            let saved = save(&m);
            let mut again = load(&saved, "its save");
            assert_eq!(save(&again), saved, "seed {seed}");
            again.resume_at(frame).unwrap();
            assert!(again.step(0).done && m.step(0).done);
            frame += 1;
            assert_eq!(save(&again), save(&m), "seed {seed}: a frame from there");
        }
    }
}

/// Past the room and past it inside the extents too is refused by name, at the load and
/// at the write that would grow to it, the write taken back.
#[test]
fn what_still_passes_the_room_is_refused() {
    let wide = "(cell \"S\" \"A1:B262144\" 1)\n";
    let rule = |k: u32, column: &str| {
        format!("(conditional \"S\" {k} \"A1:{column}1048576\" \"=1\" none)\n")
    };
    let rows = format!("{wide}{}{}{}", rule(1, "B"), rule(2, "B"), rule(3, "B"));
    let e = Machine::load(&rows, library::language()).expect_err("three of half a million");
    assert_eq!(e.refusal.id, "grid-rule-cells", "{}", e.refusal.text);
    assert_eq!(e.line, 4);
    let narrow = "(cell \"S\" \"A1:B262143\" 1)\n";
    let rows = format!("{narrow}{}{}", rule(1, "XFD"), rule(2, "XFD"));
    let mut m = load(&rows, "two at the room's edge");
    assert!(m.step(0).done);
    let e = m
        .write("(cell \"S\" \"A262146\" 1)")
        .expect_err("grown past the room");
    assert_eq!(e.refusal.id, "grid-rule-cells", "{}", e.refusal.text);
    // Taken back, the grid steps as it stood.
    m.write("(cell \"S\" \"A1\" 2)").unwrap();
    assert!(m.step(0).done);
}

/// A merged range and a rule over it, the first cell's answer `G1`'s and the
/// covered cells' their own, `H1`'s and `I1`'s.
const MERGED: &str = concat!(
    "(cell \"S\" \"C1\" \"Aa\")\n",
    "(cell \"S\" \"G1\" 1)\n",
    "(cell \"S\" \"H1\" 0)\n",
    "(cell \"S\" \"I1\" 0)\n",
    "(format 7 (font bold))\n",
    "(conditional \"S\" 1 \"C1:E1\" \"=G1=1\" 7)\n",
    "(merge \"S\" \"C1:E1\")\n",
);

/// A sheet's merged ranges print by their first cells after its rules, unclipped, in
/// every record of it and no other's, and a load and a workbook keep them.
#[test]
fn a_merge_prints_after_the_rules_in_every_record_of_its_sheet() {
    let rows = concat!(
        "(cell \"S\" \"A1\" 1)\n",
        "(merge \"S\" \"H9:I9\")\n",
        "(merge \"S\" \"D2:E2\")\n",
        "(conditional \"S\" 1 \"A1\" \"=A1>0\" none)\n",
        "(merge \"S\" \"C2:C4\")\n",
        "(cell \"T\" \"A1\" 1)\n",
    );
    let m = load(rows, "the sheet");
    let want = [
        "(conditional \"S\" 1 \"A1\" \"=A1>0\" none)",
        "(merge \"S\" \"C2:C4\")",
        "(merge \"S\" \"D2:E2\")",
        "(merge \"S\" \"H9:I9\")",
    ];
    for projection in ["grid", "values", "fills"] {
        for window in [None, Some("A1"), Some("Z100")] {
            let text = record(&m, projection, "S", window);
            let held: Vec<&str> = text
                .lines()
                .filter(|l| l.starts_with("(conditional ") || l.starts_with("(merge "))
                .collect();
            assert_eq!(held, want, "{projection} {window:?}\n{text}");
        }
        assert!(!record(&m, projection, "T", None).contains("(merge "));
    }
    let saved = save(&m);
    assert_eq!(save(&load(&saved, "the save")), saved);
    let ranges: Vec<String> = m.workbook().sheets[0]
        .merges()
        .ranges()
        .iter()
        .map(A1Range::text)
        .collect();
    assert_eq!(ranges, ["C2:C4", "D2:E2", "H9:I9"]);
}

/// A covered cell holds what its rows wrote and a formula reads it as it stands.
#[test]
fn a_merge_moves_no_value_and_a_covered_cell_is_read_as_any_cell() {
    let rows = concat!(
        "(cell \"S\" \"C1\" 5)\n",
        "(cell \"S\" \"D1\" 7)\n",
        "(formula \"S\" \"F1\" \"=D1+1\")\n",
        "(merge \"S\" \"C1:E1\")\n",
    );
    let mut m = load(rows, "the sheet");
    assert!(m.step(0).done);
    let at = record(&m, "values", "S", None);
    assert!(at.contains("(cell \"S\" \"D1\" 7)\n"), "{at}");
    assert!(at.contains("(value \"S\" \"F1\" 8)\n"), "{at}");
    m.write("(cell \"S\" \"D1\" 9)").unwrap();
    assert!(m.step(0).done);
    let at = record(&m, "values", "S", None);
    assert!(at.contains("(cell \"S\" \"D1\" 9)\n"), "{at}");
    assert!(at.contains("(value \"S\" \"F1\" 10)\n"), "{at}");
}

/// Every cell of a merged range draws what its first answers, whatever it answers
/// itself, in a window of any of them.
#[test]
fn a_merged_range_draws_by_its_first_cell_s_answers_at_every_cell() {
    let mut m = load(MERGED, "the sheet");
    assert!(m.step(0).done);
    assert_eq!(
        lit_in(&m, "A1:J1"),
        [
            "(lit \"S\" \"C1\" 1)",
            "(lit \"S\" \"D1\" 1)",
            "(lit \"S\" \"E1\" 1)"
        ]
    );
    assert_eq!(lit_in(&m, "D1"), ["(lit \"S\" \"D1\" 1)"]);
    assert_eq!(
        lit_in(&m, "D1:J9"),
        ["(lit \"S\" \"D1\" 1)", "(lit \"S\" \"E1\" 1)"]
    );
    m.write("(cell \"S\" \"G1\" 0)\n(cell \"S\" \"H1\" 1)\n(cell \"S\" \"I1\" 1)")
        .unwrap();
    assert!(m.step(0).done);
    assert_eq!(lit_in(&m, "A1:J1"), Vec::<String>::new());
    assert_eq!(lit_in(&m, "D1"), Vec::<String>::new());
}

/// A rule over the first cell alone draws over the range, one over a covered cell
/// alone over none of it, a stop ends the walk at every cell as at the first, and a
/// cell beside it draws as its own.
#[test]
fn a_rule_draws_over_a_merged_range_as_its_first_cell_answers_it() {
    let rows = concat!(
        "(cell \"S\" \"A5\" 1)\n",
        "(format 7 (font bold))\n",
        "(conditional \"S\" 1 \"C1\" \"=$A$5=1\" 7)\n",
        "(merge \"S\" \"C1:E1\")\n",
        "(conditional \"S\" 2 \"D2\" \"=$A$5=1\" 7)\n",
        "(merge \"S\" \"C2:E2\")\n",
        "(conditional \"S\" 3 \"C3\" \"=$A$5=1\" 7 stop)\n",
        "(conditional \"S\" 4 \"C3:E3\" \"=$A$5=1\" 7)\n",
        "(merge \"S\" \"C3:E3\")\n",
        "(conditional \"S\" 5 \"G1\" \"=$A$5=1\" 7)\n",
        "(merge \"S\" \"G1:H2\")\n",
        "(conditional \"S\" 6 \"F1:F3\" \"=$A$5=1\" 7)\n",
    );
    let mut m = load(rows, "the sheet");
    assert!(m.step(0).done);
    assert_eq!(
        lit_in(&m, "A1:J3"),
        [
            "(lit \"S\" \"C1\" 1)",
            "(lit \"S\" \"D1\" 1)",
            "(lit \"S\" \"E1\" 1)",
            "(lit \"S\" \"F1\" 6)",
            "(lit \"S\" \"G1\" 5)",
            "(lit \"S\" \"H1\" 5)",
            "(lit \"S\" \"F2\" 6)",
            "(lit \"S\" \"G2\" 5)",
            "(lit \"S\" \"H2\" 5)",
            "(lit \"S\" \"C3\" 3)",
            "(lit \"S\" \"D3\" 3)",
            "(lit \"S\" \"E3\" 3)",
            "(lit \"S\" \"F3\" 6)",
        ]
    );
    assert_eq!(lit_in(&m, "H2"), ["(lit \"S\" \"H2\" 5)"]);
    assert_eq!(lit_in(&m, "D2"), Vec::<String>::new());
    // A save of either shape holds them, a covered cell's row over a rule of its first's.
    let drawn = lit_in(&m, "A1:J3");
    for projection in ["grid", "fills"] {
        let saved = record(&m, projection, "S", Some("A1:J5"));
        assert!(saved.contains("(lit \"S\" \"D1\" 1)\n"), "{saved}");
        let mut again = load(&saved, projection);
        assert_eq!(lit_in(&again, "A1:J3"), drawn, "{projection} at rest");
        assert!(again.step(0).done);
        assert_eq!(lit_in(&again, "A1:J3"), drawn, "{projection} stepped");
    }
}

/// A first cell's answer moving names the range's cells in the window, its first
/// inside it or not; a covered cell's own moving names none.
#[test]
fn a_first_cell_s_answer_moving_names_the_range_s_cells_in_the_window() {
    let mut m = load(MERGED, "the sheet");
    assert!(m.step(0).done);
    m.write("(cell \"S\" \"G1\" 0)").unwrap();
    assert!(m.step(0).done);
    assert_eq!(
        record(&m, "changes", "S", Some("C1:E1")),
        concat!(
            "(changes \"S\" \"C1:E1\" 1 2 3)\n",
            "(cell \"S\" \"C1\" \"Aa\")\n",
            "(empty \"S\" \"D1\")\n",
            "(empty \"S\" \"E1\")\n",
        )
    );
    assert_eq!(
        record(&m, "changes", "S", Some("D1")),
        "(changes \"S\" \"D1\" 1 2 1)\n(empty \"S\" \"D1\")\n"
    );
    assert_eq!(
        record(&m, "changes", "S", Some("A1:B1")),
        "(changes \"S\" \"A1:B1\" 1 2 0)\n"
    );
    m.write("(cell \"S\" \"H1\" 1)").unwrap();
    assert!(m.step(0).done);
    assert_eq!(
        record(&m, "changes", "S", Some("C1:E1")),
        "(changes \"S\" \"C1:E1\" 2 3 0)\n"
    );
    m.write("(cell \"S\" \"G1\" 1)").unwrap();
    assert!(m.step(0).done);
    assert_eq!(
        record(&m, "changes", "S", Some("D1:E1")),
        "(changes \"S\" \"D1:E1\" 3 4 2)\n(lit \"S\" \"D1\" 1)\n(lit \"S\" \"E1\" 1)\n"
    );
}

/// On a sheet whose merged cells by its rules pass what rules may answer for, a range's
/// rules print at its first cell alone, so that a merge of the whole sheet costs a record
/// no row a cell, and its covered cells draw nothing of their own either way.
#[test]
fn a_merged_range_past_the_room_prints_at_its_first_cell_alone() {
    let rows = concat!(
        "(cell \"S\" \"A1\" 1)\n",
        "(formula \"S\" \"C1\" \"=$XFD$1048576\")\n",
        "(formula \"S\" \"A5\" \"=$XFD$1048576\")\n",
        "(formula \"S\" \"C5\" \"=$XFD$1048576\")\n",
        "(cell \"S\" \"XFD1048576\" 1)\n",
        "(conditional \"S\" 1 \"A1:A3\" \"=$XFD$1048576>0\" none)\n",
        "(merge \"S\" \"A1:XFD1048576\")\n",
    );
    let mut m = load(rows, "the whole sheet");
    assert!(m.step(0).done);
    assert_eq!(
        lit_in(&m, "A1:C3"),
        ["(lit \"S\" \"A1\" 1)"],
        "a covered cell's own and its first's, past the room"
    );
    assert_eq!(lit_in(&m, "A2"), Vec::<String>::new());
    let whole = record(&m, "values", "S", None);
    let lit: Vec<&str> = whole.lines().filter(|l| l.starts_with("(lit ")).collect();
    assert_eq!(lit, ["(lit \"S\" \"A1\" 1)"]);
    // A covered cell a value names prints no rule, its first's or its own.
    m.write("(cell \"S\" \"XFD1048576\" 2)").unwrap();
    assert!(m.step(0).done);
    for at in ["C1", "A5", "C5"] {
        assert_eq!(
            record(&m, "changes", "S", Some(at)),
            format!("(changes \"S\" \"{at}\" 1 2 1)\n(value \"S\" \"{at}\" 2)\n")
        );
    }
    m.write("(cell \"S\" \"XFD1048576\" 0)").unwrap();
    assert!(m.step(0).done);
    assert_eq!(
        record(&m, "changes", "S", Some("A1:C3")),
        "(changes \"S\" \"A1:C3\" 2 3 2)\n(cell \"S\" \"A1\" 1)\n(value \"S\" \"C1\" 0)\n"
    );
    assert_eq!(
        record(&m, "changes", "S", Some("A2")),
        "(changes \"S\" \"A2\" 2 3 0)\n"
    );
    // At the room, every cell; a cell past it, on the sheet and whichever range, the first;
    // and a rule more counts each merged cell again.
    let at_the_room = concat!(
        "(cell \"S\" \"A1\" 1)\n",
        "(conditional \"S\" 1 \"A1\" \"=1=1\" none)\n",
        "(merge \"S\" \"A1:A1048576\")\n",
    );
    let mut m = load(at_the_room, "at the room");
    assert!(m.step(0).done);
    assert_eq!(lit_in(&m, "A1:A3").len(), 3);
    let mut m = load(
        &format!("{at_the_room}(merge \"S\" \"B1:C1\")\n"),
        "past it",
    );
    assert!(m.step(0).done);
    assert_eq!(lit_in(&m, "A1:A3"), ["(lit \"S\" \"A1\" 1)"]);
    let half = concat!(
        "(cell \"S\" \"A1\" 1)\n",
        "(conditional \"S\" 1 \"A1\" \"=1=1\" none)\n",
        "(conditional \"S\" 2 \"A1\" \"=1=1\" none)\n",
        "(merge \"S\" \"A1:A524288\")\n",
    );
    let mut m = load(half, "half the room, twice");
    assert!(m.step(0).done);
    assert_eq!(lit_in(&m, "A1:A3").len(), 3);
    let mut m = load(
        &format!("{half}(conditional \"S\" 3 \"Z1\" \"=1=1\" none)\n"),
        "half the room, three times",
    );
    assert!(m.step(0).done);
    assert_eq!(lit_in(&m, "A1:A3"), ["(lit \"S\" \"A1\" 1 2)"]);
    // The rules counted where a covered cell's value names it and where a first's answer moves.
    let moving = concat!(
        "(cell \"S\" \"A1\" 1)\n",
        "(cell \"S\" \"Z9\" 1)\n",
        "(formula \"S\" \"A5\" \"=$Z$9\")\n",
        "(conditional \"S\" 1 \"A1\" \"=$Z$9>0\" none)\n",
        "(conditional \"S\" 2 \"A1\" \"=$Z$9>0\" none)\n",
        "(conditional \"S\" 3 \"Z1\" \"=1=1\" none)\n",
        "(merge \"S\" \"A1:A524288\")\n",
    );
    let mut m = load(moving, "half the room, three times, moving");
    assert!(m.step(0).done);
    m.write("(cell \"S\" \"Z9\" 2)").unwrap();
    assert!(m.step(0).done);
    assert_eq!(
        record(&m, "changes", "S", Some("A5")),
        "(changes \"S\" \"A5\" 1 2 1)\n(value \"S\" \"A5\" 2)\n"
    );
    m.write("(cell \"S\" \"Z9\" 0)").unwrap();
    assert!(m.step(0).done);
    assert_eq!(
        record(&m, "changes", "S", Some("A1:A3")),
        "(changes \"S\" \"A1:A3\" 2 3 1)\n(cell \"S\" \"A1\" 1)\n"
    );
    // Another sheet's cells are another sheet's.
    let mut m = load(
        &format!("{at_the_room}(merge \"T\" \"B1:C1\")\n"),
        "beside it",
    );
    assert!(m.step(0).done);
    assert_eq!(lit_in(&m, "A1:A3").len(), 3);
}

#[test]
fn a_merge_row_is_refused_by_what_it_gets_wrong() {
    let bad = |rows: &str, id: &str| {
        let refusal = Machine::load(rows, library::language())
            .err()
            .unwrap_or_else(|| panic!("{rows} was taken"));
        assert_eq!(refusal.refusal.id, id, "{rows}: {}", refusal.refusal.text);
        refusal
    };
    for row in [
        "(merge \"S\")",
        "(merge \"S\" \"C6\")",
        "(merge \"S\" \"C6:C6\")",
        "(merge \"S\" \"nope\")",
        "(merge \"S\" \"C6:E6\" 1)",
        "(merge S \"C6:E6\")",
        "(merge \"S\" C6:E6)",
    ] {
        let refused = bad(row, "grid-row-malformed");
        assert!(refused.refusal.text.contains("more than one cell"), "{row}");
    }
    bad("(merge \"S.last\" \"A1:B1\")", "grid-sheet-name-invalid");
    for (first, second) in [
        ("C6:E6", "E6:F7"),
        ("E6:F7", "C6:E6"),
        ("A1:D4", "B2:C3"),
        ("B2:C3", "A1:D4"),
        ("C6:E6", "C6:D6"),
    ] {
        let rows =
            format!("(merge \"S\" \"{first}\")\n(cell \"S\" \"A9\" 1)\n(merge \"S\" \"{second}\")");
        let refused = bad(&rows, "grid-merge-overlap");
        assert_eq!(refused.line, 3);
        let text = &refused.refusal.text;
        assert!(
            text.contains(&format!("S!{second}")) && text.contains(first),
            "{text}"
        );
    }
    // A second alike is the first, however it is spelled.
    let m = load(
        "(merge \"S\" \"C6:E6\")\n(merge \"S\" \"C6:E6\")\n(merge \"S\" \"E6:C6\")\n",
        "alike",
    );
    assert_eq!(record(&m, "grid", "S", None).matches("(merge ").count(), 1);
    // Beside it, and on another sheet, is none of it.
    let beside = concat!(
        "(merge \"S\" \"C6:E6\")\n",
        "(merge \"S\" \"F6:G6\")\n",
        "(merge \"S\" \"C7:E7\")\n",
        "(merge \"T\" \"C6:E6\")\n",
    );
    let m = load(beside, "beside");
    assert_eq!(record(&m, "grid", "S", None).matches("(merge ").count(), 3);
}

/// A covered cell's `lit` row says its first cell's answers, so that it names a rule
/// over that cell, and a load draws by the first cell's own row alone.
#[test]
fn a_covered_cell_s_lit_row_is_held_to_its_range_s_first_cell() {
    let head = concat!(
        "(cell \"S\" \"A5\" 1)\n",
        "(conditional \"S\" 1 \"C1\" \"=$A$5=1\" none)\n",
        "(conditional \"S\" 2 \"D1\" \"=$A$5=1\" none)\n",
        "(merge \"S\" \"C1:E1\")\n",
    );
    load(&format!("{head}(lit \"S\" \"D1\" 1)\n"), "over its first");
    for row in [
        "(lit \"S\" \"D1\" 2)",
        "(lit \"S\" \"C1\" 2)",
        "(lit \"S\" \"D1\" 1 2)",
    ] {
        let refused = Machine::load(&format!("{head}{row}\n"), library::language())
            .err()
            .unwrap_or_else(|| panic!("{row} was taken"));
        assert_eq!(refused.refusal.id, "grid-row-malformed", "{row}");
    }
    let m = load(&format!("{head}(lit \"S\" \"C1\" 1)\n"), "the first's");
    assert_eq!(
        lit_in(&m, "C1:E1"),
        [
            "(lit \"S\" \"C1\" 1)",
            "(lit \"S\" \"D1\" 1)",
            "(lit \"S\" \"E1\" 1)"
        ]
    );
    // A row at a covered cell alone says nothing of its first.
    let m = load(&format!("{head}(lit \"S\" \"D1\" 1)\n"), "the covered's");
    assert_eq!(lit_in(&m, "C1:E1"), Vec::<String>::new());
    // A rule of many rectangles, read by its pieces.
    let many: Vec<String> = (1..=20).map(|row| format!("C{row}")).collect();
    let rows = format!(
        "(conditional \"S\" 1 \"{}\" \"=1=1\" none)\n(conditional \"S\" 2 \"{}\" \"=1=1\" none)\n(merge \"S\" \"C5:E5\")\n",
        many.join(" "),
        many.join(" ").replace('C', "F")
    );
    load(&format!("{rows}(lit \"S\" \"D5\" 1)\n"), "by its pieces");
    let refused = Machine::load(
        &format!("{rows}(lit \"S\" \"D5\" 2)\n"),
        library::language(),
    )
    .expect_err("a rule over no first cell");
    assert_eq!(refused.refusal.id, "grid-row-malformed");
    // One of many rectangles past the room, whose pieces no load makes.
    let columns: Vec<String> = (0..17)
        .map(|k| {
            let at = cell_ref(1, 3 + 3 * k);
            let column = at.strip_suffix('1').expect("a first row");
            format!("{column}1:{column}1048576")
        })
        .collect();
    let rows = format!(
        "(cell \"S\" \"A1\" 1)\n(conditional \"S\" 1 \"{}\" \"=1=1\" none)\n(conditional \"S\" 2 \"D1:D1048576\" \"=1=1\" none)\n(merge \"S\" \"C5:D5\")\n",
        columns.join(" "),
    );
    load(&format!("{rows}(lit \"S\" \"D5\" 1)\n"), "past the room");
    let refused = Machine::load(
        &format!("{rows}(lit \"S\" \"D5\" 2)\n"),
        library::language(),
    )
    .expect_err("a rule over no first cell, past the room");
    assert_eq!(refused.refusal.id, "grid-row-malformed");
}

#[test]
fn a_door_prints_a_merge_s_row() {
    let rows = "(cell \"G\" \"A1\" 1)\n(merge \"G\" \"C1:D1\")\n";
    let forms = read_forms(rows).unwrap();
    let wb = read_rows(&forms).unwrap().workbook();
    let window = window_of(&wb, "G", None).unwrap();
    let got = view_text(&wb, &window);
    assert!(got.contains("(merge \"G\" \"C1:D1\")\n"), "{got}");
    assert_eq!(wb.sheets[0].merges().len(), 1);
}
