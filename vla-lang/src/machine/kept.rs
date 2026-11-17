//! A call kept (KERNEL.29's third slice's third part, 2026-11-17): every
//! cell of a grid whose short shapes keep their calls of named functions held
//! to the same grid with every call written in, as the third slice's second
//! part wrote them (`shape::inlined`), at every frame and after each write;
//! a write that moves a shape's form across a run's least cells held to a
//! load of the written grid; and a call kept followed by the reach, its edit
//! as narrow as the call written in, a recursion's among them, computed
//! every frame nowhere.

use super::{Machine, Viewed};
use crate::calc::library;
use crate::calc::shape::{inlined, Form};

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

/// The grid loaded with every call written in.
fn load_inlined(rows: &str) -> Machine {
    let _always = inlined::always();
    load(rows)
}

fn text(m: &Machine, projection: &str, sheet: &str) -> String {
    match m.view(projection, sheet, None) {
        Ok(Viewed::Text(t)) => t,
        other => panic!("{projection} of {sheet}: {other:?}"),
    }
}

/// Every sheet's values, the twins aside.
fn values(m: &Machine) -> String {
    (0..m.sheets())
        .filter_map(|i| m.sheet_name(i))
        .map(|s| text(m, "values", s))
        .collect()
}

/// A save as a host prints one: every sheet's fills, then the names.
fn save(m: &Machine) -> String {
    let mut out: String = (0..m.sheets())
        .filter_map(|i| m.sheet_name(i))
        .map(|s| text(m, "fills", s))
        .collect();
    out.push_str(&text(m, "names", ""));
    out
}

/// The cells of the shapes computed every frame, as the census's plan row
/// counts them.
fn every_frame(m: &Machine) -> usize {
    let census = m.census();
    let plan = census.lines().next().unwrap_or_default();
    let at = plan.find("(every-frame ").expect("the plan row") + "(every-frame ".len();
    plan[at..]
        .split(')')
        .next()
        .and_then(|n| n.parse().ok())
        .expect("a count")
}

/// A gate's operand: an input, an earlier gate, the room's value cell, a
/// cell of another sheet, or a number.
fn leaf(r: &mut Rng, inputs: usize, gates: usize) -> String {
    match r.below(7) {
        0 | 1 => format!("A{}", 1 + r.below(inputs)),
        2 | 3 if gates > 0 => format!("C{}", 1 + r.below(gates)),
        4 => "B1".to_string(),
        5 => "Other!A1".to_string(),
        _ => r.below(4).to_string(),
    }
}

fn gate(r: &mut Rng, depth: u32, inputs: usize, gates: usize) -> String {
    if depth == 0 || r.below(3) == 0 {
        return leaf(r, inputs, gates);
    }
    let a = gate(r, depth - 1, inputs, gates);
    let b = gate(r, depth - 1, inputs, gates);
    match r.below(6) {
        0 => format!("{a}+{b}"),
        1 => format!("{a}-{b}"),
        2 => format!("{a}*{b}"),
        3 => format!("MOD({a},3)"),
        4 => format!("MAX({a},{b})"),
        _ => format!("IF({a}>{b},{a},{b})"),
    }
}

/// A grid of three functions and their calls: `F`, a room of gates on its
/// sheet (`function`), reading its inputs, earlier gates, its value cell
/// `B1` and `Other!A1`; `G`, a `name` row calling `F` and reading another
/// sheet's cells and its cell's row; `H`, a recursion reading another sheet.
/// Calls at places of their own, each reading its own row, a short fill and
/// a fill of more cells than a run's least.
struct Grid {
    rows: String,
    inputs: usize,
    gates: usize,
}

fn grid(r: &mut Rng) -> Grid {
    let inputs = 1 + r.below(3);
    let gates = 2 + r.below(4);
    let mut rows = String::new();
    for i in 1..=inputs {
        rows.push_str(&format!("(cell \"Room\" \"A{i}\" {})\n", r.below(5)));
    }
    rows.push_str(&format!("(cell \"Room\" \"B1\" {})\n", 1 + r.below(4)));
    for g in 1..=gates {
        let e = gate(r, 2, inputs, g - 1);
        rows.push_str(&format!("(formula \"Room\" \"C{g}\" \"={e}\")\n"));
    }
    rows.push_str(&format!(
        "(function \"F\" \"Room\" \"A1:A{inputs}\" \"C{gates}\")\n"
    ));
    for row in 1..=3 {
        rows.push_str(&format!("(cell \"Other\" \"A{row}\" {})\n", r.below(6)));
        rows.push_str(&format!("(cell \"Other\" \"B{row}\" {})\n", r.below(6)));
    }
    rows.push_str(&g_row(r, inputs));
    rows.push_str("(name \"H\" \"=LAMBDA(n,IF(n<=0,Other!$A$3,n+H(n-1)))\")\n");
    for row in 1..=24 {
        for col in ["A", "B", "C"] {
            rows.push_str(&format!("(cell \"Calls\" \"{col}{row}\" {})\n", r.below(6)));
        }
    }
    // Calls at places of their own, each reading a row of its own, so that
    // each is a shape of one cell or of a few.
    for at in 1..=24 {
        let row = 1 + r.below(24);
        rows.push_str(&format!(
            "(formula \"Calls\" \"E{at}\" \"={}\")\n",
            call(r, inputs, row)
        ));
    }
    rows.push_str("(formula \"Calls\" \"F1:F6\" \"=G(A1,B1)*2\")\n");
    let args: Vec<String> = ["A1", "B1", "C1"][..inputs]
        .iter()
        .map(|s| s.to_string())
        .collect();
    rows.push_str(&format!(
        "(formula \"Calls\" \"H1:H20\" \"=F({})+G(B1,C1)\")\n",
        args.join(",")
    ));
    Grid {
        rows,
        inputs,
        gates,
    }
}

/// `G`'s row: a `LET`, a comparison, a call of `F` at its parameters, its
/// cell's row and a range of another sheet.
fn g_row(r: &mut Rng, inputs: usize) -> String {
    let args: Vec<&str> = ["x", "y", "t"][..inputs].to_vec();
    let read = match r.below(3) {
        0 => "Other!$A$2",
        1 => "SUM(Other!$B$1:$B$3)",
        _ => "ROW()",
    };
    format!(
        "(name \"G\" \"=LAMBDA(x,y,LET(t,x*2+{read},IF(t>y,t-y,F({})+1)))\")\n",
        args.join(",")
    )
}

/// A call reading `row`'s inputs.
fn call(r: &mut Rng, inputs: usize, row: usize) -> String {
    let args: Vec<String> = ["A", "B", "C"][..inputs]
        .iter()
        .map(|c| format!("{c}{row}"))
        .collect();
    let f = format!("F({})", args.join(","));
    let g = format!("G(A{row},B{row})");
    let h = format!("H(MOD(C{row},4))");
    match r.below(6) {
        0 => f,
        1 => g,
        2 => h,
        3 => format!("IF(A{row}>2,{f},{g})"),
        4 => format!("{f}+{g}*{h}"),
        _ => format!("F({})", vec![g.clone(); inputs].join(",")),
    }
}

/// A write between two frames: an input of the calls, a cell a body reads,
/// a gate rewritten, `G` rewritten, a call rewritten, or the short fill
/// grown past a run's least and cut back.
fn write(r: &mut Rng, g: &Grid) -> String {
    match r.below(9) {
        0 | 1 => format!(
            "(cell \"Calls\" \"{}{}\" {})",
            ["A", "B", "C"][r.below(3)],
            1 + r.below(24),
            r.below(6)
        ),
        2 => format!("(cell \"Other\" \"A{}\" {})", 1 + r.below(3), r.below(6)),
        3 => format!("(cell \"Other\" \"B{}\" {})", 1 + r.below(3), r.below(6)),
        4 => format!("(cell \"Room\" \"B1\" {})", 1 + r.below(4)),
        5 => {
            let at = 1 + r.below(g.gates);
            format!(
                "(formula \"Room\" \"C{at}\" \"={}\")",
                gate(r, 2, g.inputs, at - 1)
            )
        }
        6 => g_row(r, g.inputs).trim_end().to_string(),
        7 => {
            let at = 1 + r.below(24);
            let row = 1 + r.below(24);
            format!(
                "(formula \"Calls\" \"E{at}\" \"={}\")",
                call(r, g.inputs, row)
            )
        }
        _ => {
            if r.below(2) == 0 {
                "(formula \"Calls\" \"F1:F20\" \"=G(A1,B1)*2\")".to_string()
            } else {
                "(cell \"Calls\" \"F7:F20\" 0)".to_string()
            }
        }
    }
}

#[test]
fn a_call_kept_computes_as_the_call_written_in_at_every_frame_and_write() {
    // Each call kept by its program where its function's body lowers, held
    // to the same calls kept and computed by `Ctx`, and both to the calls
    // written in (the third slice's second part).
    for seed in 1..=40u64 {
        let mut r = Rng(seed);
        let g = grid(&mut r);
        let mut kept = load(&g.rows);
        let mut by_ctx = load(&g.rows);
        by_ctx.lowered = false;
        let mut written = load_inlined(&g.rows);
        let when = |k: usize| format!("seed {seed}, write {k}\n{}", g.rows);
        // The calls at places of their own and the short fill keep their
        // calls; the long fill writes them in.
        let forms = |m: &Machine, form: Form| {
            m.shapes
                .every_shape()
                .iter()
                .filter(|s| s.form == form)
                .count()
        };
        assert!(forms(&kept, Form::Kept) > 0, "{}", when(0));
        assert!(forms(&kept, Form::Inlined) > 0, "{}", when(0));
        assert_eq!(forms(&written, Form::Kept), 0, "{}", when(0));
        assert!(kept.step(0).done, "{}", when(0));
        assert!(by_ctx.step(0).done, "{}", when(0));
        assert!(written.step(0).done, "{}", when(0));
        assert_eq!(values(&kept), values(&by_ctx), "{}", when(0));
        assert_eq!(values(&kept), values(&written), "{}", when(0));
        // The short fill's six calls of `G` at least by programs, `F`'s
        // body run inside each; the recursion's left to `Ctx`.
        assert!(kept.ran >= 6, "{}: {} by programs", when(0), kept.ran);
        assert_eq!(every_frame(&kept), 0, "{}\n{}", when(0), kept.census());
        for k in 1..=12 {
            let w = write(&mut r, &g);
            let a = kept.write(&w);
            assert_eq!(by_ctx.write(&w).is_ok(), a.is_ok(), "{}: {w}", when(k));
            let b = {
                let _always = inlined::always();
                written.write(&w)
            };
            assert_eq!(a.is_ok(), b.is_ok(), "{}: {w}\n{a:?}\n{b:?}", when(k));
            assert!(kept.step(0).done, "{}: {w}", when(k));
            assert!(by_ctx.step(0).done, "{}: {w}", when(k));
            assert!(written.step(0).done, "{}: {w}", when(k));
            assert_eq!(values(&kept), values(&by_ctx), "{}: {w}", when(k));
            assert_eq!(values(&kept), values(&written), "{}: {w}", when(k));
            assert_eq!(every_frame(&kept), 0, "{}: {w}\n{}", when(k), kept.census());
            // The write's plan is a load's of the written grid (the third
            // slice's law), the forms chosen again by the cells.
            let mut again = load(&save(&kept));
            assert!(again.step(0).done, "{}: {w}", when(k));
            assert_eq!(values(&again), values(&kept), "{}: {w}", when(k));
            assert_eq!(again.census(), kept.census(), "{}: {w}", when(k));
        }
    }
}

/// A cell's value, by programs where they answer and by `Ctx` alone, the
/// two held alike, and the cell's shape's census row.
fn both(rows: &str, sheet: &str, row: u32, col: u32) -> (crate::calc::Computed, String) {
    use crate::calc::Values;
    let mut m = load(rows);
    let mut by_ctx = load(rows);
    by_ctx.lowered = false;
    assert!(m.step(0).done && by_ctx.step(0).done);
    let s = m.find_sheet(sheet).expect("the sheet");
    let got = m.computed((s, row, col)).expect("a value");
    assert_eq!(Some(got.clone()), by_ctx.computed((s, row, col)), "{rows}");
    let census = m.census();
    let shape = census
        .lines()
        .find(|l| {
            let at = crate::sheet::cell_ref(row, col);
            l.starts_with("(shape") && l.contains(&format!("\"{sheet}\" \"{at}\""))
        })
        .unwrap_or_default()
        .to_string();
    (got, shape)
}

#[test]
fn a_body_is_one_program_its_bindings_registers_given_back_at_any_length() {
    // A chain of gates in series through `LET`, as a room's gates written
    // out are: sixty gates, each read once by the next, lowered in a program
    // of 32 registers, each binding's register given back after its read,
    // where written into the call it nested past a program's depth.
    let mut body = String::from("LET(g_1,MOD(a+b,7)");
    for j in 2..=60 {
        body.push_str(&format!(",g_{j},MOD(g_{}*a+b,7)", j - 1));
    }
    body.push_str(",g_60)");
    let rows = format!(
        "(name \"CHAIN\" \"=LAMBDA(a,b,{body})\")\n(cell \"S\" \"A1\" 3)\n(cell \"S\" \"B1\" 4)\n\
         (formula \"S\" \"C1\" \"=CHAIN(A1,B1)\")\n"
    );
    let (got, shape) = both(&rows, "S", 1, 3);
    assert!(!shape.contains("unlowered"), "{shape}");
    // The chain computed by hand.
    let (a, b) = (3.0f64, 4.0f64);
    let mut g = (a + b) % 7.0;
    for _ in 2..=60 {
        g = (g * a + b) % 7.0;
    }
    assert_eq!(
        got,
        crate::calc::Computed::Value(crate::calc::Value::Number(g))
    );
}

#[test]
fn a_parameter_folded_is_handed_a_cell_s_number_and_nothing_is_ctx_s() {
    // `MAX` folds a cell's reference as a range of one cell, which skips a
    // cell holding nothing, and a value given, which counts it 0: a call
    // handing a parameter an aggregate folds reads its cell strictly, so that
    // a cell holding nothing leaves the call to `Ctx`. And a function whose
    // answer is its parameter, a reference where a cell's is handed it,
    // folded by its caller as a range.
    for (a, b, want) in [
        ("-5", "-2", -2.0),
        ("", "-2", -2.0),
        ("3", "", 3.0),
        ("", "", 0.0),
    ] {
        let cell = |c: &str, v: &str| {
            if v.is_empty() {
                String::new()
            } else {
                format!("(cell \"S\" \"{c}1\" {v})\n")
            }
        };
        let rows = format!(
            "(name \"BIG\" \"=LAMBDA(x,y,MAX(x,y))\")\n(name \"SAME\" \"=LAMBDA(x,x)\")\n\
             {}{}(formula \"S\" \"C1\" \"=BIG(A1,B1)\")\n(formula \"S\" \"D1\" \"=MAX(SAME(A1),SAME(B1))\")\n",
            cell("A", a),
            cell("B", b)
        );
        let n = crate::calc::Computed::Value(crate::calc::Value::Number(want));
        let (got, shape) = both(&rows, "S", 1, 3);
        assert_eq!(got, n, "BIG({a},{b})");
        assert!(!shape.contains("unlowered"), "{shape}");
        let (got, _) = both(&rows, "S", 1, 4);
        assert_eq!(got, n, "MAX(SAME({a}),SAME({b}))");
    }
}

#[test]
fn calls_deeper_than_a_program_nests_are_ctx_s() {
    // Twenty functions, each calling the next: a call of the first nests
    // past `program::NEST` and is left to `Ctx`; a call of the fifteenth
    // nests within it and is a program.
    let mut rows = String::from("(cell \"S\" \"A1\" 2)\n");
    for k in 1..20 {
        rows.push_str(&format!("(name \"F{k}\" \"=LAMBDA(x,F{}(x)+1)\")\n", k + 1));
    }
    rows.push_str("(name \"F20\" \"=LAMBDA(x,x*2)\")\n");
    rows.push_str("(formula \"S\" \"B1\" \"=F1(A1)\")\n(formula \"S\" \"C1\" \"=F15(A1)\")\n");
    let (deep, deep_row) = both(&rows, "S", 1, 2);
    let (shallow, shallow_row) = both(&rows, "S", 1, 3);
    let n = |x: f64| crate::calc::Computed::Value(crate::calc::Value::Number(x));
    assert_eq!(deep, n(4.0 + 19.0));
    assert_eq!(shallow, n(4.0 + 5.0));
    assert!(deep_row.contains("(unlowered deep)"), "{deep_row}");
    assert!(!shallow_row.contains("unlowered"), "{shallow_row}");
}

#[test]
fn a_call_kept_reads_and_calls_what_its_body_reads_and_calls() {
    use crate::calc::{Computed, Value, Values};
    let n = |x: f64| Some(Computed::Value(Value::Number(x)));
    // The plan: a body reading a formula cell of a later sheet orders the
    // call after it, so that the first frame reads it computed; and a write
    // to what that cell reads reaches the call through it.
    let rows = "(sheet \"Calls\" visible)\n(sheet \"Other\" visible)\n\
                (name \"RD\" \"=LAMBDA(x,x+Other!$C$1)\")\n(cell \"Calls\" \"A1\" 1)\n\
                (formula \"Calls\" \"B1\" \"=RD(A1)\")\n(cell \"Other\" \"A1\" 5)\n\
                (formula \"Other\" \"C1\" \"=A1*2\")\n";
    let mut m = load(rows);
    assert!(m.step(0).done);
    assert_eq!(m.computed((0, 1, 2)), n(11.0));
    m.write("(cell \"Other\" \"A1\" 7)").unwrap();
    assert!(m.step(0).done);
    assert_eq!(m.computed((0, 1, 2)), n(15.0));
    assert_eq!(every_frame(&m), 0, "{}", m.census());
    // The refusals: a body calling a function the library does not hold is
    // refused at a call kept as at a call written in.
    let missing = "(name \"BAD\" \"=LAMBDA(x,NOSUCH(x))\")\n(cell \"S\" \"A1\" 1)\n\
                   (formula \"S\" \"B1\" \"=BAD(A1)\")\n";
    for always in [false, true] {
        let refused = {
            let _always = always.then(inlined::always);
            Machine::load(missing, library::language()).map(|_| ())
        };
        let refusal = refused.expect_err("refused");
        assert_eq!(refusal.refusal.id, "calc-function-not-computed", "{always}");
    }
    // The extent: a body whose approximate match bisects its column between
    // the first and last places of its sheet's extent (KERNEL.56), so that a
    // value written below the last row, in another column, which the call
    // reads at no place, moves its answer, 1 to 4 over a column of 1, 5, 2
    // and 3, as it moves the call written in.
    let extent = "(sheet \"S\" visible)
(sheet \"D\" visible)
                  (name \"AT\" \"=LAMBDA(x,MATCH(x,D!$A:$A,1))\")
                  (cell \"D\" \"A1\" 1)
(cell \"D\" \"A2\" 5)
                  (cell \"D\" \"A3\" 2)
(cell \"D\" \"A4\" 3)
                  (formula \"S\" \"B1\" \"=AT(4)\")
";
    let mut kept = load(extent);
    let mut written = load_inlined(extent);
    assert!(kept.step(0).done && written.step(0).done);
    assert_eq!(kept.computed((0, 1, 2)), n(1.0));
    let w = "(cell \"D\" \"C8\" 1)";
    kept.write(w).unwrap();
    {
        let _always = inlined::always();
        written.write(w).unwrap();
    }
    assert!(kept.step(0).done && written.step(0).done);
    assert_eq!(kept.computed((0, 1, 2)), n(4.0));
    assert_eq!(values(&kept), values(&written));
}

#[test]
fn a_body_making_more_calls_than_a_cell_may_is_ctx_s() {
    // Eleven functions each calling the next four times: a call of the first
    // makes more than `calls::CALLS` calls, which `Ctx` counts to `#NUM!`, so
    // its program is never lowered; the ninth's, 4^3 calls, is.
    let mut rows = String::from("(cell \"S\" \"A1\" 1)\n");
    for k in 1..=11 {
        let next = format!("F{}(x)", k + 1);
        rows.push_str(&format!(
            "(name \"F{k}\" \"=LAMBDA(x,{next}+{next}+{next}+{next})\")\n"
        ));
    }
    rows.push_str("(name \"F12\" \"=LAMBDA(x,x)\")\n");
    rows.push_str("(formula \"S\" \"B1\" \"=F1(A1)\")\n(formula \"S\" \"C1\" \"=F9(A1)\")\n");
    let m = load(&rows);
    let census = m.census();
    let row = |at: &str| {
        census
            .lines()
            .find(|l| l.contains(&format!("\"S\" \"{at}\"")))
            .unwrap_or_default()
            .to_string()
    };
    assert!(row("B1").contains("(unlowered deep)"), "{census}");
    assert!(!row("C1").contains("unlowered"), "{census}");
}

#[test]
fn a_call_of_a_function_a_sheet_scopes_is_computed_every_frame() {
    // A function some sheet scopes, whose meaning a shape on several sheets
    // cannot fix: its call followed by no reach, computed every frame, and a
    // write to what its body reads seen at the next.
    use crate::calc::{Computed, Value, Values};
    let rows = "(sheet \"S\" visible)\n(sheet \"Other\" visible)\n\
                (name \"S!K\" \"=LAMBDA(x,x+Other!$A$1)\")\n(cell \"S\" \"A1\" 1)\n\
                (formula \"S\" \"B1\" \"=K(A1)\")\n(cell \"Other\" \"A1\" 2)\n";
    let mut m = load(rows);
    assert!(m.step(0).done);
    assert_eq!(
        m.computed((0, 1, 2)),
        Some(Computed::Value(Value::Number(3.0)))
    );
    assert_eq!(every_frame(&m), 1, "{}", m.census());
    m.write("(cell \"Other\" \"A1\" 5)").unwrap();
    assert!(m.step(0).done);
    assert_eq!(
        m.computed((0, 1, 2)),
        Some(Computed::Value(Value::Number(6.0)))
    );
}

#[test]
fn a_body_answers_its_call_and_binds_a_range_as_its_reference() {
    // A body whose root is a call answering a text answers its call, which
    // its caller reads, never the cell as a cell's whole formula's root does;
    // and a binding of a range, which no register holds, is taken as its
    // reference by the aggregate reading it.
    let rows = "(sheet \"S\" visible)\n(sheet \"D\" visible)\n\
                (name \"LETTER\" \"=LAMBDA(k,CHAR(64+k))\")\n\
                (name \"TOTAL\" \"=LAMBDA(x,LET(r,D!$B$1:$B$3,SUM(r)+x))\")\n\
                (cell \"D\" \"B1\" 1)\n(cell \"D\" \"B3\" 4)\n(cell \"S\" \"A1\" 10)\n\
                (formula \"S\" \"B1\" \"=LEN(LETTER(2))\")\n(formula \"S\" \"C1\" \"=TOTAL(A1)\")\n";
    let n = |x: f64| crate::calc::Computed::Value(crate::calc::Value::Number(x));
    let (len, len_row) = both(rows, "S", 1, 2);
    assert_eq!(len, n(1.0));
    assert!(!len_row.contains("unlowered"), "{len_row}");
    let (total, total_row) = both(rows, "S", 1, 3);
    assert_eq!(total, n(15.0));
    assert!(!total_row.contains("unlowered"), "{total_row}");
}

#[test]
fn a_function_calling_a_function_s_value_it_is_handed_is_written_in() {
    // `TWICE` calls the function it is handed, which no reach follows: a
    // call of it at a cell alone is written in, the function it is handed
    // reduced where the call stands, and computed when what it reads moves,
    // never every frame; a call of `SQ` beside it kept, in a cell of its own
    // and in the formula calling both.
    use crate::calc::{Computed, Value, Values};
    let rows = "(name \"TWICE\" \"=LAMBDA(g,x,g(g(x)))\")\n(name \"SQ\" \"=LAMBDA(x,x*x)\")\n\
                (cell \"S\" \"A1\" 3)\n(formula \"S\" \"B1\" \"=TWICE(SQ,A1)\")\n\
                (formula \"S\" \"C1\" \"=TWICE(LAMBDA(x,x+1),A1)\")\n(formula \"S\" \"D1\" \"=SQ(A1)\")\n\
                (formula \"S\" \"E1\" \"=TWICE(SQ,A1)+SQ(A1)\")\n";
    let mut m = load(rows);
    assert!(m.step(0).done);
    let n = |x: f64| Some(Computed::Value(Value::Number(x)));
    let row = |m: &Machine| -> Vec<_> { (2..=5).map(|col| m.computed((0, 1, col))).collect() };
    assert_eq!(row(&m), [n(81.0), n(5.0), n(9.0), n(90.0)]);
    assert_eq!(every_frame(&m), 0, "{}", m.census());
    let forms: Vec<Form> = m.shapes.every_shape().iter().map(|s| s.form).collect();
    assert_eq!(forms, [Form::Plain, Form::Plain, Form::Kept, Form::Kept]);
    m.write("(cell \"S\" \"A1\" 2)").unwrap();
    assert!(m.step(0).done);
    assert_eq!(row(&m), [n(16.0), n(4.0), n(4.0), n(20.0)]);
}
