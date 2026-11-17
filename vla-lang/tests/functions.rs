//! A grid's functions of cells (KERNEL.29, its third slice's second part):
//! a `function` row, a room of cells on a sheet compiled to the `LAMBDA` a
//! person would write from it; called as that `LAMBDA` is; printed by the
//! `names` projection as its row, which a save prints beside every sheet's
//! `fills`; compiled again at a write to its row or to a formula of its room,
//! every shape calling it planned again; and `#CALC!` where its body does not
//! compile, the census saying why.

use vla_lang::calc::library;
use vla_lang::calc::{Computed, Value, Values};
use vla_lang::machine::{Machine, Viewed};

/// The seeded generator the laws use (`tests/laws.rs`), the same cases on
/// every run.
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
}

fn load(rows: &str) -> Machine {
    Machine::load(rows, library::language())
        .unwrap_or_else(|e| panic!("line {}: {}\n{rows}", e.line, e.refusal.text))
}

fn text(m: &Machine, projection: &str, sheet: &str) -> String {
    match m.view(projection, sheet, None) {
        Ok(Viewed::Text(t)) => t,
        other => panic!("{projection} of {sheet}: {other:?}"),
    }
}

fn names(m: &Machine) -> String {
    text(m, "names", "")
}

/// Every sheet's values, the twins aside.
fn values(m: &Machine) -> String {
    (0..m.sheets())
        .filter_map(|i| m.sheet_name(i))
        .map(|s| text(m, "values", s))
        .collect()
}

/// A save as a host prints one: every sheet's fills, then the names, each
/// function's row among them.
fn save(m: &Machine) -> String {
    let mut out: String = (0..m.sheets())
        .filter_map(|i| m.sheet_name(i))
        .map(|s| text(m, "fills", s))
        .collect();
    out.push_str(&names(m));
    out
}

fn computed(m: &Machine, sheet: &str, row: u32, col: u32) -> Computed {
    let si = m.find_sheet(sheet).expect("the sheet");
    m.computed((si, row, col))
        .unwrap_or_else(|| panic!("{sheet}!R{row}C{col} computes nothing"))
}

fn number(m: &Machine, sheet: &str, row: u32, col: u32) -> f64 {
    match computed(m, sheet, row, col) {
        Computed::Value(Value::Number(n)) => n,
        other => panic!("{sheet}!R{row}C{col}: {other:?}"),
    }
}

fn is_calc(m: &Machine, sheet: &str, row: u32, col: u32) -> bool {
    matches!(
        computed(m, sheet, row, col),
        Computed::Value(Value::Error(k)) if k.text() == "#CALC!"
    )
}

/// Naught's VIII SUMMA, a full adder (Alonzo's `CART.9`): its switches
/// `A3:A5`, its gates, and as chips the sum, the carry, and both as one
/// array.
const SUMMA: &str = "(cell \"VIII SUMMA\" \"A3:A5\" 0)\n\
                     (formula \"VIII SUMMA\" \"C3\" \"=MOD(A3+A4,2)\")\n\
                     (formula \"VIII SUMMA\" \"C4\" \"=A3*A4\")\n\
                     (formula \"VIII SUMMA\" \"C5\" \"=C3*A5\")\n\
                     (formula \"VIII SUMMA\" \"J3\" \"=MOD(C3+A5,2)\")\n\
                     (formula \"VIII SUMMA\" \"J4\" \"=MAX(C4,C5)\")\n\
                     (function \"SUMMA\" \"VIII SUMMA\" \"A3:A5\" \"J3\")\n\
                     (function \"SUMMACARRY\" \"VIII SUMMA\" \"A3:A5\" \"J4\")\n\
                     (function \"SUMMAS\" \"VIII SUMMA\" \"A3:A5\" \"J3:J4\")\n";

/// Every input of three bits on a sheet of calls, a row each, and each chip
/// called at it: the sum, the carry, and the array's two picked by `INDEX`.
fn calls() -> String {
    let mut s = String::new();
    for k in 0..8u32 {
        let row = k + 1;
        let (a, b, c) = ((k >> 2) & 1, (k >> 1) & 1, k & 1);
        s.push_str(&format!(
            "(cell \"Calls\" \"A{row}\" {a})\n(cell \"Calls\" \"B{row}\" {b})\n(cell \"Calls\" \"C{row}\" {c})\n"
        ));
    }
    s.push_str("(formula \"Calls\" \"D1:D8\" \"=SUMMA(A1,B1,C1)\")\n");
    s.push_str("(formula \"Calls\" \"E1:E8\" \"=SUMMACARRY(A1,B1,C1)\")\n");
    s.push_str("(formula \"Calls\" \"F1:F8\" \"=INDEX(SUMMAS(A1,B1,C1),1)\")\n");
    s.push_str("(formula \"Calls\" \"G1:G8\" \"=INDEX(SUMMAS(A1,B1,C1),2)\")\n");
    s
}

#[test]
fn viii_summa_as_a_chip_is_a_full_adder_at_every_input() {
    let mut m = load(&format!("{SUMMA}{}", calls()));
    assert!(m.step(0).done);
    for k in 0..8u32 {
        let (a, b, c) = ((k >> 2) & 1, (k >> 1) & 1, k & 1);
        let sum = f64::from((a + b + c) % 2);
        let carry = f64::from(u32::from(a + b + c >= 2));
        let row = k + 1;
        assert_eq!(number(&m, "Calls", row, 4), sum, "SUMMA({a},{b},{c})");
        assert_eq!(
            number(&m, "Calls", row, 5),
            carry,
            "SUMMACARRY({a},{b},{c})"
        );
        assert_eq!(
            number(&m, "Calls", row, 6),
            sum,
            "SUMMAS({a},{b},{c})'s first"
        );
        assert_eq!(
            number(&m, "Calls", row, 7),
            carry,
            "SUMMAS({a},{b},{c})'s second"
        );
    }
    // The room computes at its own switches, its function at the calls'.
    assert_eq!(number(&m, "VIII SUMMA", 3, 10), 0.0);
    assert_eq!(number(&m, "VIII SUMMA", 4, 10), 0.0);
    // A switch of the room written moves the room alone: its values never
    // move the function's LAMBDA, which reads none of them.
    assert!(m.step(0).done);
    assert_eq!(m.write("(cell \"VIII SUMMA\" \"A3\" 1)"), Ok(1));
    assert!(m.step(0).done);
    // C3, C4, C5 and J3, which read the switch or a gate it moved; J4's
    // gates kept their values, so it is not reached; and no call.
    assert_eq!(
        m.reached(),
        4,
        "the room's gates the switch reaches and no call"
    );
    assert_eq!(number(&m, "VIII SUMMA", 3, 10), 1.0);
    assert_eq!(number(&m, "Calls", 8, 4), 1.0);
}

#[test]
fn the_names_view_prints_each_function_as_its_row_and_a_save_loads_it() {
    let mut m = load(&format!("(name \"rate\" \"0.5\")\n{SUMMA}{}", calls()));
    let printed = names(&m);
    assert_eq!(
        printed,
        "(name \"rate\" \"0.5\")\n\
         (function \"SUMMA\" \"VIII SUMMA\" \"A3:A5\" \"J3\")\n\
         (function \"SUMMACARRY\" \"VIII SUMMA\" \"A3:A5\" \"J4\")\n\
         (function \"SUMMAS\" \"VIII SUMMA\" \"A3:A5\" \"J3:J4\")\n"
    );
    assert!(m.step(0).done);
    let mut back = load(&save(&m));
    assert!(back.step(0).done);
    assert_eq!(names(&back), printed);
    assert_eq!(values(&back), values(&m));
    assert_eq!(back.census(), m.census());
    // A function row and the sheet it names may come in either order, and a
    // row said again alike is the one row.
    let m = load(
        "(function \"F\" \"S\" \"A1\" \"B1\")\n\
         (function \"f\" \"s\" \"a1\" \"B1\")\n\
         (cell \"S\" \"A1\" 2)\n\
         (formula \"S\" \"B1\" \"=A1*3\")\n",
    );
    assert_eq!(names(&m), "(function \"F\" \"S\" \"A1\" \"B1\")\n");
}

/// A room generated with the `LAMBDA` a person writes from it by hand, by
/// substitution and without the compile: its inputs `A1:A<n>` as the
/// parameters `x_1` to `x_n`; its gates `C1:C<g>`, each reading the inputs,
/// the gates before it, a value of the room (`B1`), a cell of another sheet
/// or a number, each written twice, as the gate's formula and as the
/// expression the hand writes for it; and its output, its last gate or its
/// last two as an array.
struct Generated {
    inputs: usize,
    room: String,
    hand: String,
    output: String,
    pair: bool,
}

/// A gate's operand, as the room's formula and as the hand writes it: a gate
/// the hand substitutes in parentheses, and only where that stays short, so
/// that the hand's text stays within a formula's length.
fn leaf(r: &mut Rng, inputs: usize, gates: &[(String, String)]) -> (String, String) {
    match r.below(6) {
        0 | 1 => {
            let i = 1 + r.below(inputs);
            (format!("A{i}"), format!("x_{i}"))
        }
        2 | 3 if !gates.is_empty() => {
            let g = r.below(gates.len());
            if gates[g].1.len() < 240 {
                (format!("C{}", g + 1), format!("({})", gates[g].1))
            } else {
                let i = 1 + r.below(inputs);
                (format!("A{i}"), format!("x_{i}"))
            }
        }
        4 => ("B1".to_string(), "Room!$B$1".to_string()),
        5 => ("Other!A1".to_string(), "Other!A1".to_string()),
        _ => {
            let n = r.below(4).to_string();
            (n.clone(), n)
        }
    }
}

/// A gate's expression, written alike as the formula and by the hand.
fn expr(r: &mut Rng, depth: u32, inputs: usize, gates: &[(String, String)]) -> (String, String) {
    if depth == 0 || r.below(3) == 0 {
        return leaf(r, inputs, gates);
    }
    let a = expr(r, depth - 1, inputs, gates);
    let b = expr(r, depth - 1, inputs, gates);
    let both = |f: &dyn Fn(&str, &str) -> String| (f(&a.0, &b.0), f(&a.1, &b.1));
    match r.below(7) {
        0 => both(&|x, y| format!("{x}+{y}")),
        1 => both(&|x, y| format!("{x}-{y}")),
        2 => both(&|x, y| format!("{x}*{y}")),
        3 => both(&|x, _| format!("MOD({x},3)")),
        4 => both(&|x, y| format!("MAX({x},{y})")),
        5 => both(&|x, y| format!("IF({x}>{y},{x},{y})")),
        _ => both(&|x, _| format!("ABS({x})")),
    }
}

fn generated(r: &mut Rng) -> Generated {
    let inputs = 1 + r.below(4);
    let count = 2 + r.below(5);
    let mut room = String::new();
    for i in 1..=inputs {
        room.push_str(&format!("(cell \"Room\" \"A{i}\" {})\n", r.below(5)));
    }
    room.push_str(&format!("(cell \"Room\" \"B1\" {})\n", 1 + r.below(4)));
    room.push_str(&format!("(cell \"Other\" \"A1\" {})\n", r.below(7)));
    let mut gates: Vec<(String, String)> = Vec::new();
    for g in 1..=count {
        let e = expr(r, 2, inputs, &gates);
        room.push_str(&format!("(formula \"Room\" \"C{g}\" \"={}\")\n", e.0));
        gates.push(e);
    }
    let pair = r.below(4) == 0;
    let last = &gates[count - 1].1;
    let (hand, output) = if pair {
        let before = &gates[count - 2].1;
        (
            format!("VSTACK(({before}),({last}))"),
            format!("C{}:C{count}", count - 1),
        )
    } else {
        (last.clone(), format!("C{count}"))
    };
    Generated {
        inputs,
        room,
        hand,
        output,
        pair,
    }
}

/// A sheet of calls of `F`: six rows of inputs, their calls by fill, each
/// output of a pair picked by `INDEX`, and a call of numbers alone.
fn called(r: &mut Rng, g: &Generated) -> String {
    let mut s = String::new();
    let cols = ["A", "B", "C", "D"];
    for row in 1..=6 {
        for col in &cols[..g.inputs] {
            s.push_str(&format!("(cell \"Calls\" \"{col}{row}\" {})\n", r.below(6)));
        }
    }
    let args: Vec<String> = cols[..g.inputs].iter().map(|c| format!("{c}1")).collect();
    let call = format!("F({})", args.join(","));
    let numbers: Vec<String> = (0..g.inputs).map(|_| r.below(5).to_string()).collect();
    let alone = format!("F({})", numbers.join(","));
    if g.pair {
        s.push_str(&format!(
            "(formula \"Calls\" \"E1:E6\" \"=INDEX({call},1)\")\n"
        ));
        s.push_str(&format!(
            "(formula \"Calls\" \"F1:F6\" \"=INDEX({call},2)\")\n"
        ));
        s.push_str(&format!(
            "(formula \"Calls\" \"G1\" \"=INDEX({alone},2)\")\n"
        ));
    } else {
        s.push_str(&format!("(formula \"Calls\" \"E1:E6\" \"={call}+1\")\n"));
        s.push_str(&format!("(formula \"Calls\" \"G1\" \"={alone}\")\n"));
    }
    s
}

#[test]
fn a_function_of_cells_computes_as_its_lambda_written_by_hand() {
    for seed in 1..=60u64 {
        let mut r = Rng(seed);
        let g = generated(&mut r);
        let calls = called(&mut r, &g);
        let params: Vec<String> = (1..=g.inputs).map(|i| format!("x_{i}")).collect();
        let by_room = format!(
            "{}(function \"F\" \"Room\" \"A1:A{}\" \"{}\")\n{calls}",
            g.room, g.inputs, g.output
        );
        let by_hand = format!(
            "{}(name \"F\" \"=LAMBDA({},{})\")\n{calls}",
            g.room,
            params.join(","),
            g.hand
        );
        let when = format!("seed {seed}\n{by_room}\n{by_hand}");
        let mut room = load(&by_room);
        let mut hand = load(&by_hand);
        assert!(room.step(0).done, "{when}");
        assert!(hand.step(0).done, "{when}");
        assert_eq!(values(&room), values(&hand), "{when}");
        assert_eq!(room.census(), hand.census(), "{when}");
        // And again after the calls' inputs move.
        let write = format!("(cell \"Calls\" \"A1:A6\" {})", r.below(6));
        assert_eq!(room.write(&write), hand.write(&write), "{when}");
        assert!(room.step(0).done && hand.step(0).done, "{when}");
        assert_eq!(values(&room), values(&hand), "{when}");
    }
}

/// A write between two frames: a gate's formula replaced, reading the
/// inputs and the gates before it; a gate made a value, or the room's value
/// a formula; an input or a call's input written; the function's row
/// rewritten, its output another gate or a pair; or a second function
/// defined, or redefined, and called.
fn room_write(r: &mut Rng, inputs: usize, count: usize) -> String {
    let gate = 1 + r.below(count);
    let earlier: Vec<(String, String)> = (1..gate)
        .map(|k| (format!("C{k}"), String::new()))
        .collect();
    let cols = ["A", "B", "C", "D"];
    let args: Vec<String> = cols[..inputs].iter().map(|c| format!("{c}1")).collect();
    match r.below(8) {
        0 | 1 => {
            let e = expr(r, 2, inputs, &earlier);
            format!("(formula \"Room\" \"C{gate}\" \"={}\")", e.0)
        }
        2 => format!("(cell \"Room\" \"C{gate}\" {})", r.below(5)),
        3 => format!("(formula \"Room\" \"B1\" \"=A1+{}\")", r.below(3)),
        4 => format!(
            "(cell \"Room\" \"A{}\" {})\n(cell \"Calls\" \"A{}\" {})",
            1 + r.below(inputs),
            r.below(5),
            1 + r.below(6),
            r.below(5)
        ),
        5 => {
            let out = if gate > 1 && r.below(3) == 0 {
                format!("C{}:C{gate}", gate - 1)
            } else {
                format!("C{gate}")
            };
            format!("(function \"F\" \"Room\" \"A1:A{inputs}\" \"{out}\")")
        }
        _ => format!(
            "(function \"G\" \"Room\" \"A1:A{inputs}\" \"C{gate}\")\n\
             (formula \"Calls\" \"H1\" \"=G({})*2\")",
            args.join(",")
        ),
    }
}

#[test]
fn a_write_to_a_body_or_a_row_computes_as_a_load_of_the_written_grid() {
    let mut written = 0;
    for seed in 1..=40u64 {
        let mut r = Rng(1000 + seed);
        let g = generated(&mut r);
        let count = g.room.matches("(formula \"Room\" \"C").count();
        let calls = called(&mut r, &g);
        let rows = format!(
            "{}(function \"F\" \"Room\" \"A1:A{}\" \"{}\")\n{calls}",
            g.room, g.inputs, g.output
        );
        let mut m = load(&rows);
        for frame in 1..=6 {
            assert!(m.step(0).done);
            let write = room_write(&mut r, g.inputs, count);
            let when = format!("seed {seed}, frame {frame}: {write}\n{rows}");
            m.write(&write)
                .unwrap_or_else(|e| panic!("{when}: {}", e.refusal.text));
            written += 1;
            assert!(m.step(0).done);
            let mut fresh = load(&save(&m));
            assert!(fresh.step(0).done);
            assert_eq!(m.census(), fresh.census(), "{when}");
            assert_eq!(values(&m), values(&fresh), "{when}");
            assert_eq!(names(&m), names(&fresh), "{when}");
        }
    }
    assert_eq!(written, 240);
}

#[test]
fn a_recursive_function_of_cells_computes_as_its_lambda_written_by_hand() {
    // The room's own gate calls the function it is the body of: a call kept
    // as a call, which a cycle of named functions holds, at the room's
    // switch and at every call.
    let room = "(cell \"Fact\" \"A1\" 4)\n\
                (formula \"Fact\" \"B1\" \"=IF(A1<=1,1,A1*FACT(A1-1))\")\n";
    let mut calls = String::new();
    for k in 1..=5 {
        calls.push_str(&format!("(cell \"Calls\" \"A{k}\" {k})\n"));
    }
    calls.push_str("(formula \"Calls\" \"B1:B5\" \"=FACT(A1)\")\n");
    calls.push_str("(formula \"Calls\" \"C1\" \"=SUM(B1:B5)\")\n");
    let mut by_room = load(&format!(
        "{room}(function \"FACT\" \"Fact\" \"A1\" \"B1\")\n{calls}"
    ));
    let mut by_hand = load(&format!(
        "{room}(name \"FACT\" \"=LAMBDA(x_1,IF(x_1<=1,1,x_1*FACT(x_1-1)))\")\n{calls}"
    ));
    assert!(by_room.step(0).done && by_hand.step(0).done);
    assert_eq!(number(&by_room, "Fact", 1, 2), 24.0);
    assert_eq!(number(&by_room, "Calls", 5, 2), 120.0);
    assert_eq!(number(&by_room, "Calls", 1, 3), 153.0);
    assert_eq!(values(&by_room), values(&by_hand));
    assert_eq!(by_room.census(), by_hand.census());
    let write = "(cell \"Calls\" \"A1:A5\" 3)";
    assert_eq!(by_room.write(write), by_hand.write(write));
    assert!(by_room.step(0).done && by_hand.step(0).done);
    assert_eq!(number(&by_room, "Calls", 1, 3), 30.0);
    assert_eq!(values(&by_room), values(&by_hand));
}

#[test]
fn a_write_to_a_room_that_leaves_the_grid_unstepped_puts_the_function_back() {
    let mut m = load(&format!("{SUMMA}{}", calls()));
    assert!(m.step(0).done);
    let (at, census, seen) = (save(&m), m.census(), values(&m));
    // The sum's first gate made to read the output it feeds: the room's own
    // cells close a cycle, which compiles each chip again to a cycle, and the
    // write is refused, the room and every chip's LAMBDA put back.
    let e = m
        .write("(formula \"VIII SUMMA\" \"C3\" \"=J3\")")
        .unwrap_err();
    assert_eq!(e.refusal.id, "calc-cycle", "{}", e.refusal.text);
    assert_eq!(save(&m), at);
    assert_eq!(m.census(), census);
    assert!(m.step(0).done);
    assert_eq!(values(&m), seen);
    // And the room still takes a gate it can step.
    assert_eq!(
        m.write("(formula \"VIII SUMMA\" \"C3\" \"=MAX(A3,A4)\")"),
        Ok(1)
    );
    assert!(m.step(0).done);
    assert_eq!(
        number(&m, "Calls", 8, 4),
        0.0,
        "SUMMA(1,1,1) reads MAX(1,1) now"
    );
}

#[test]
fn a_chip_of_chips_follows_a_write_to_the_chip_inside_it() {
    let mut m = load(
        "(cell \"Half\" \"A1:A2\" 0)\n\
         (formula \"Half\" \"B1\" \"=MOD(A1+A2,2)\")\n\
         (function \"HALF\" \"Half\" \"A1:A2\" \"B1\")\n\
         (cell \"Full\" \"A1:A3\" 0)\n\
         (formula \"Full\" \"B1\" \"=HALF(HALF(A1,A2),A3)\")\n\
         (function \"FULL\" \"Full\" \"A1:A3\" \"B1\")\n\
         (cell \"S\" \"A1:C1\" 1)\n\
         (formula \"S\" \"D1\" \"=FULL(A1,B1,C1)\")\n",
    );
    assert!(m.step(0).done);
    assert_eq!(number(&m, "S", 1, 4), 1.0);
    // The half adder's gate made a sum: FULL's own LAMBDA calls HALF by
    // name and is the same, and every call of FULL is planned again.
    assert_eq!(m.write("(formula \"Half\" \"B1\" \"=A1+A2\")"), Ok(1));
    assert!(m.step(0).done);
    assert_eq!(number(&m, "S", 1, 4), 3.0);
    let mut fresh = load(&save(&m));
    assert!(fresh.step(0).done);
    assert_eq!(m.census(), fresh.census());
    assert_eq!(values(&m), values(&fresh));
}

#[test]
fn a_body_no_lambda_writes_answers_calc_and_the_census_says_why() {
    let mut m = load(
        "(cell \"Room\" \"A1:A3\" 1)\n\
         (formula \"Room\" \"B1\" \"=SUM(A1:A3)\")\n\
         (function \"F\" \"Room\" \"A1:A3\" \"B1\")\n\
         (cell \"S\" \"A1\" 2)\n\
         (formula \"S\" \"B1\" \"=F(A1,A1,A1)\")\n\
         (formula \"S\" \"C1\" \"=F(A1,A1,A1)+1\")\n",
    );
    assert!(m.step(0).done);
    // The room computes; its function, a range over its inputs, does not.
    assert_eq!(number(&m, "Room", 1, 2), 3.0);
    assert!(is_calc(&m, "S", 1, 2) && is_calc(&m, "S", 1, 3));
    assert!(
        m.census().contains("(uncompiled \"F\" range \"A1:A3\")\n"),
        "{}",
        m.census()
    );
    // The gate written to read its inputs one at a time: the calls compute.
    assert_eq!(m.write("(formula \"Room\" \"B1\" \"=A1+A2+A3\")"), Ok(1));
    assert!(m.step(0).done);
    assert_eq!(number(&m, "S", 1, 2), 6.0);
    assert_eq!(number(&m, "S", 1, 3), 7.0);
    assert!(!m.census().contains("uncompiled"), "{}", m.census());
    // A function of more inputs than a LAMBDA takes, which no call can call
    // with its parameters: a call of it is an error, as a call of a LAMBDA
    // with too few arguments is.
    assert_eq!(
        m.write("(function \"G\" \"Room\" \"A1:A300\" \"B1\")\n(formula \"S\" \"D1\" \"=G(1)\")"),
        Ok(2)
    );
    assert!(m.step(0).done);
    assert!(matches!(
        computed(&m, "S", 1, 4),
        Computed::Value(Value::Error(_))
    ));
    assert!(
        m.census().contains("(uncompiled \"G\" inputs 300)\n"),
        "{}",
        m.census()
    );
    let mut fresh = load(&save(&m));
    assert!(fresh.step(0).done);
    assert_eq!(m.census(), fresh.census());
    assert_eq!(values(&m), values(&fresh));
}

#[test]
fn a_load_refuses_a_function_row_out_of_its_shape_by_name() {
    let refused = |rows: &str| {
        Machine::load(rows, library::language())
            .err()
            .unwrap_or_else(|| panic!("loaded: {rows}"))
    };
    for bad in [
        "(function \"A1\" \"S\" \"A1\" \"B1\")",
        "(function \"R1C1\" \"S\" \"A1\" \"B1\")",
        "(function \"F\" \"S\" \"A1:\" \"B1\")",
        "(function \"F\" \"S\" \"A1\")",
        "(function \"F\" \"\" \"A1\" \"B1\")",
        "(function \"2F\" \"S\" \"A1\" \"B1\")",
    ] {
        let e = refused(&format!("(cell \"S\" \"A1\" 1)\n{bad}\n"));
        assert_eq!(
            (e.line, e.refusal.id.as_str()),
            (2, "grid-row-malformed"),
            "{bad}: {}",
            e.refusal.text
        );
    }
    let e = refused("(cell \"S\" \"A1\" 1)\n(function \"F\" \"Nowhere\" \"A1\" \"B1\")\n");
    assert_eq!((e.line, e.refusal.id.as_str()), (2, "grid-function-sheet"));
    assert_eq!(
        e.refusal.text,
        "The function F reads its body from the sheet Nowhere, which the grid does not hold; it holds S."
    );
    let e = refused(
        "(name \"F\" \"1\")\n(function \"f\" \"S\" \"A1\" \"B1\")\n(cell \"S\" \"A1\" 1)\n",
    );
    assert_eq!((e.line, e.refusal.id.as_str()), (2, "grid-name-taken"));
    assert_eq!(
        e.refusal.text,
        "f is a name of this grid already; a formula calls a function and a name alike, so a function row cannot take the name."
    );
    let e = refused(
        "(function \"F\" \"S\" \"A1\" \"B1\")\n(name \"F\" \"1\")\n(cell \"S\" \"A1\" 1)\n",
    );
    assert_eq!((e.line, e.refusal.id.as_str()), (2, "grid-name-taken"));
    assert!(
        e.refusal.text.starts_with("F is a function of this grid"),
        "{}",
        e.refusal.text
    );
    let e = refused("(function \"F\" \"S\" \"A1\" \"B1\")\n(function \"F\" \"S\" \"A1\" \"B2\")\n(cell \"S\" \"A1\" 1)\n");
    assert_eq!(
        (e.line, e.refusal.id.as_str()),
        (2, "grid-cell-written-twice")
    );
    assert!(e
        .refusal
        .text
        .starts_with("(function \"F\") was already written by line 1"));
}

#[test]
fn a_write_takes_a_function_row_as_a_load_reads_one() {
    let mut m = load(&format!("(name \"rate\" \"2\")\n{SUMMA}{}", calls()));
    assert!(m.step(0).done);
    let before = save(&m);
    // The row the grid holds writes nothing; spelled without case, it is the
    // same function, its place kept and its spelling the write's.
    assert_eq!(
        m.write("(function \"SUMMA\" \"VIII SUMMA\" \"A3:A5\" \"J3\")"),
        Ok(0)
    );
    assert_eq!(
        m.write("(function \"summa\" \"VIII SUMMA\" \"A3:A5\" \"J3\")"),
        Ok(1)
    );
    assert!(names(&m).contains("(function \"summa\" \"VIII SUMMA\" \"A3:A5\" \"J3\")\n"));
    // Rewritten to the carry: every call of it computes the carry.
    assert_eq!(
        m.write("(function \"SUMMA\" \"VIII SUMMA\" \"A3:A5\" \"J4\")"),
        Ok(1)
    );
    assert!(m.step(0).done);
    assert_eq!(number(&m, "Calls", 4, 4), 1.0, "SUMMA(0,1,1) the carry now");
    let at = save(&m);
    for (bad, id) in [
        (
            "(function \"F\" \"VIII SUMMA\" \"A3\")",
            "grid-row-malformed",
        ),
        (
            "(function \"F\" \"Nowhere\" \"A3\" \"J3\")",
            "grid-function-sheet",
        ),
        (
            "(function \"F\" \"Calls.last\" \"A3\" \"J3\")",
            "grid-function-sheet",
        ),
        (
            "(function \"rate\" \"VIII SUMMA\" \"A3\" \"J3\")",
            "grid-name-taken",
        ),
        ("(name \"SUMMAS\" \"1\")", "grid-name-taken"),
        // A later row refused undoes the function before it.
        (
            "(function \"NEW\" \"VIII SUMMA\" \"A3:A5\" \"J3\")\n(cell \"Nowhere\" \"A1\" 1)",
            "grid-sheet-unknown",
        ),
    ] {
        let e = m.write(bad).unwrap_err();
        assert_eq!(e.refusal.id, id, "{bad}: {}", e.refusal.text);
        assert_eq!(save(&m), at, "{bad}: a refused write changes nothing");
    }
    assert_ne!(before, at);
    // A function defined by a write and called in the same write.
    assert_eq!(
        m.write("(function \"NEW\" \"VIII SUMMA\" \"A3:A5\" \"C4\")\n(formula \"Calls\" \"H1:H8\" \"=NEW(A1,B1,C1)\")"),
        Ok(9)
    );
    assert!(m.step(0).done);
    assert_eq!(number(&m, "Calls", 8, 8), 1.0, "NEW(1,1,1) is A AND B");
    assert_eq!(number(&m, "Calls", 6, 8), 0.0, "NEW(1,0,1)");
}
