//! The formula syntax's laws (VLA.4), through the public API a tool has.
//!
//! 1. The reading fixture (`testdata/formula_syntax.txt`) reads as its
//!    golden records: the canonical expression and the text written of it,
//!    or a refusal by id. `VLA_BLESS=1` writes the golden again; a person
//!    reads it before it is committed, the reader having no VBA twin.
//! 2. Over every formula of the fixture and of the refers fixture (the
//!    reference's own 85, every kind of reference), a text read is written
//!    in the canonical spelling, that spelling reads back as the same
//!    expression and writes as the same text (both canonical forms are
//!    fixed points), and the evaluator's parser reads the canonical text
//!    as the tree it reads the original as; a text not read is refused by
//!    one of the reader's two ids, never anything else.
//! 3. Over expressions a seeded generator writes, the same ones on every
//!    run: the evaluator's parser reads the text written as exactly the
//!    tree the expression means, its operators folded from the left and
//!    its precedence Excel's, and reading that text and writing it again
//!    gives it back byte for byte.

use vla_lang::calc::formula::{parse, BinOp, Expr};
use vla_lang::form::Form;
use vla_lang::formula_syntax::{read, write};
use vla_lang::printer::write_datum;

const CASES: &str = include_str!("../testdata/formula_syntax.txt");
const GOLDEN: &str = include_str!("../testdata/formula_syntax_golden.txt");
const REFERS: &str = include_str!("../testdata/refers.txt");

/// The fixture's cases: each header and the formula's text on the line
/// after it.
fn cases(text: &str) -> Vec<(String, String)> {
    let mut out = Vec::new();
    let lines: Vec<&str> = text.split('\n').map(|l| l.trim_end_matches('\r')).collect();
    for (i, ln) in lines.iter().enumerate() {
        if let Some(name) = ln.strip_prefix("=== ") {
            out.push((
                name.to_string(),
                lines.get(i + 1).unwrap_or(&"").to_string(),
            ));
        }
    }
    out
}

fn record(formula: &str) -> String {
    match read(formula) {
        Ok(v) => match write(&v) {
            Ok(t) => format!("VLA\t{}\nTEXT\t{t}\n", write_datum(&v)),
            Err(e) => format!(
                "VLA\t{}\nUNWRITTEN\t{}\t{}\n",
                write_datum(&v),
                e.id,
                e.text
            ),
        },
        Err(e) => format!("REFUSED\t{}\t{}\n", e.id, e.text),
    }
}

#[test]
fn the_reading_fixture_is_its_golden() {
    let mut got = String::new();
    let all = cases(CASES);
    assert!(
        all.len() >= 70,
        "only {} cases in the reading fixture",
        all.len()
    );
    for (name, formula) in &all {
        got.push_str(&format!("=== {name}\n"));
        got.push_str(&record(formula));
    }
    if std::env::var("VLA_BLESS").as_deref() == Ok("1") {
        let path = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/testdata/formula_syntax_golden.txt"
        );
        std::fs::write(path, got.replace('\n', "\r\n")).unwrap();
        return;
    }
    let want = GOLDEN.replace("\r\n", "\n");
    if got != want {
        let (g, w): (Vec<&str>, Vec<&str>) = (got.lines().collect(), want.lines().collect());
        let at = (0..g.len().max(w.len()))
            .find(|&i| g.get(i) != w.get(i))
            .unwrap_or(0);
        panic!(
            "the reading fixture's record differs at line {}:\n  got  {:?}\n  want {:?}",
            at + 1,
            g.get(at),
            w.get(at)
        );
    }
    assert!(
        !got.contains("UNWRITTEN"),
        "an expression read is one the writer refuses"
    );
}

/// The evaluator's tree as a text: references by their sheets, rectangle
/// and `$` marks, numbers by value, texts and names as they stand.
fn shape(e: &Expr) -> String {
    match e {
        Expr::Number(v) => format!("{v}"),
        Expr::Text(s) => format!("{s:?}"),
        Expr::Bool(b) => format!("{b}"),
        Expr::Error(k) => k.text().to_string(),
        Expr::Empty => "_".to_string(),
        Expr::Ref(r) => format!(
            "[{:?}:{:?} {},{},{},{} {:?}]",
            r.sheet,
            r.last_sheet,
            r.top,
            r.left,
            r.bottom,
            r.right,
            (
                r.corners.row_abs1,
                r.corners.col_abs1,
                r.corners.row_abs2,
                r.corners.col_abs2
            )
        ),
        Expr::Name { name, sheet } => format!("<{sheet:?}!{name}>"),
        Expr::Call(name, args) => format!(
            "({name} {})",
            args.iter().map(shape).collect::<Vec<_>>().join(" ")
        ),
        Expr::Neg(x) => format!("(neg {})", shape(x)),
        Expr::Pos(x) => format!("(pos {})", shape(x)),
        Expr::Percent(x) => format!("(pct {})", shape(x)),
        Expr::Binary(op, a, b) => format!("({} {} {})", op_word(*op), shape(a), shape(b)),
    }
}

fn op_word(op: BinOp) -> &'static str {
    match op {
        BinOp::Add => "+",
        BinOp::Sub => "-",
        BinOp::Mul => "*",
        BinOp::Div => "/",
        BinOp::Pow => "^",
        BinOp::Concat => "&",
        BinOp::Eq => "=",
        BinOp::Ne => "<>",
        BinOp::Lt => "<",
        BinOp::Gt => ">",
        BinOp::Le => "<=",
        BinOp::Ge => ">=",
    }
}

/// Law 2 over one formula's text.
fn assert_laws_on_text(formula: &str, what: &str) -> bool {
    let v = match read(formula) {
        Ok(v) => v,
        Err(e) => {
            assert!(
                e.id == "vla-formula-read-malformed" || e.id == "vla-formula-read-unsupported",
                "{what}: {formula} refused as {}",
                e.id
            );
            return false;
        }
    };
    let t = write(&v).unwrap_or_else(|e| {
        panic!(
            "{what}: {formula} read as {} and not written: {}",
            write_datum(&v),
            e.text
        )
    });
    let v2 = read(&t).unwrap_or_else(|e| panic!("{what}: {t} written and not read: {}", e.text));
    assert_eq!(v2, v, "{what}: {formula} -> {t} reads back otherwise");
    assert_eq!(write(&v2).unwrap(), t, "{what}: {t} writes otherwise");
    let original = parse(formula)
        .unwrap_or_else(|c| panic!("{what}: {formula} read here and not by the evaluator: {c}"));
    let canonical =
        parse(&t).unwrap_or_else(|c| panic!("{what}: {t} not read by the evaluator: {c}"));
    assert_eq!(
        shape(&canonical),
        shape(&original),
        "{what}: {formula} -> {t}"
    );
    true
}

#[test]
fn the_laws_hold_over_the_fixture_and_the_reference_s_formulas() {
    let mut read_count = 0;
    for (name, formula) in cases(CASES) {
        if assert_laws_on_text(&formula, &name) {
            read_count += 1;
        }
    }
    assert!(
        read_count >= 45,
        "only {read_count} of the fixture's formulas read"
    );
    // The refers fixture: a case's first line is the cell, the next the
    // formula.
    let mut refers_read = 0;
    let mut refers_all = 0;
    for (name, _) in cases(REFERS) {
        let block: Vec<&str> = REFERS
            .split('\n')
            .map(|l| l.trim_end_matches('\r'))
            .skip_while(|l| *l != format!("=== {name}"))
            .skip(1)
            .take_while(|l| !l.starts_with("=== "))
            .collect();
        let Some(formula) = block.get(1) else {
            continue;
        };
        refers_all += 1;
        if assert_laws_on_text(formula, &name) {
            refers_read += 1;
        }
    }
    assert!(
        refers_all >= 80,
        "only {refers_all} formulas in the refers fixture"
    );
    assert!(
        refers_read >= 50,
        "only {refers_read} of the refers fixture's formulas read"
    );
}

// ---- 3. generated expressions -----------------------------------------

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

/// An atom and the shape the evaluator's parser gives it.
fn atom(r: &mut Rng) -> (Form, String) {
    let numeral = |t: &str| -> (Form, String) {
        let v: f64 = t.parse().unwrap();
        (Form::sym(t), format!("{v}"))
    };
    match r.below(14) {
        0 => numeral(["0", "7", "42", "1000000"][r.below(4)]),
        1 => numeral(["1.50", ".5", "2.25", "1e3", "1E+05"][r.below(5)]),
        2 => {
            let t = ["7", "1.5", ".25"][r.below(3)];
            let v: f64 = t.parse().unwrap();
            (Form::sym(&format!("-{t}")), format!("(neg {v})"))
        }
        3 => {
            let t = ["", "say \"hi\"", "a;b", "▲", "x y"][r.below(5)];
            (Form::string(t), format!("{t:?}"))
        }
        4 => {
            let b = r.below(2) == 0;
            (Form::sym(if b { "TRUE" } else { "false" }), format!("{b}"))
        }
        5 => {
            let (t, canon) = [
                ("#n/a", "#N/A"),
                ("#DIV/0!", "#DIV/0!"),
                ("#value!", "#VALUE!"),
            ][r.below(3)];
            (
                Form::list(vec![Form::sym("error"), Form::string(t)]),
                canon.to_string(),
            )
        }
        6 => (
            Form::sym("A1"),
            "[None:None 1,1,1,1 (false, false, false, false)]".to_string(),
        ),
        7 => (
            Form::sym("$b$2"),
            "[None:None 2,2,2,2 (true, true, false, false)]".to_string(),
        ),
        8 => (
            Form::sym("Data!C3:D4"),
            "[Some(\"Data\"):None 3,3,4,4 (false, false, false, false)]".to_string(),
        ),
        9 => (
            Form::list(vec![
                Form::sym("ref"),
                Form::string("Q1 Data"),
                Form::string("E5"),
            ]),
            "[Some(\"Q1 Data\"):None 5,5,5,5 (false, false, false, false)]".to_string(),
        ),
        10 => (Form::sym("Rate"), "<None!Rate>".to_string()),
        11 => (
            Form::sym("Model!Local"),
            "<Some(\"Model\")!Local>".to_string(),
        ),
        12 => (Form::sym("x"), "<None!x>".to_string()),
        _ => (
            Form::sym("Tally.last!B16"),
            "[Some(\"Tally.last\"):None 16,2,16,2 (false, false, false, false)]".to_string(),
        ),
    }
}

/// An expression of depth at most `depth`, and the evaluator's shape of it.
fn expr(r: &mut Rng, depth: u32) -> (Form, String) {
    if depth == 0 || r.below(4) == 0 {
        return atom(r);
    }
    const BINARY: &[(&str, &str)] = &[
        ("+", "+"),
        ("-", "-"),
        ("*", "*"),
        ("/", "/"),
        ("^", "^"),
        ("&", "&"),
        ("=", "="),
        ("<>", "<>"),
        ("<", "<"),
        (">", ">"),
        ("<=", "<="),
        (">=", ">="),
    ];
    match r.below(10) {
        0..=4 => {
            let (op, word) = BINARY[r.below(BINARY.len())];
            let n = 2 + r.below(2);
            let mut items = vec![Form::sym(op)];
            let (first, mut acc) = expr(r, depth - 1);
            items.push(first);
            for _ in 1..n {
                let (f, s) = expr(r, depth - 1);
                items.push(f);
                acc = format!("({word} {acc} {s})");
            }
            (Form::list(items), acc)
        }
        5 => {
            let (f, s) = expr(r, depth - 1);
            (Form::list(vec![Form::sym("-"), f]), format!("(neg {s})"))
        }
        6 => {
            let (f, s) = expr(r, depth - 1);
            (Form::list(vec![Form::sym("+"), f]), format!("(pos {s})"))
        }
        7 => {
            let (f, s) = expr(r, depth - 1);
            (Form::list(vec![Form::sym("%"), f]), format!("(pct {s})"))
        }
        8 => {
            let (c, cs) = expr(r, depth - 1);
            let (a, as_) = expr(r, depth - 1);
            let (b, bs) = expr(r, depth - 1);
            (
                Form::list(vec![Form::sym("if"), c, a, b]),
                format!("(IF {cs} {as_} {bs})"),
            )
        }
        _ => {
            let name = ["sum", "MAX", "Stdev.S"][r.below(3)];
            let n = r.below(3);
            let mut items = vec![Form::sym(name)];
            let mut shapes = Vec::new();
            for _ in 0..n {
                let (f, s) = expr(r, depth - 1);
                items.push(f);
                shapes.push(s);
            }
            (
                Form::list(items),
                format!("({} {})", name.to_ascii_uppercase(), shapes.join(" ")),
            )
        }
    }
}

#[test]
fn a_written_expression_reads_as_the_tree_it_means() {
    let mut written = 0;
    for seed in 0..3000u64 {
        let mut r = Rng(seed);
        let (v, want) = expr(&mut r, 4);
        let t = write(&v)
            .unwrap_or_else(|e| panic!("seed {seed}: {} not written: {}", write_datum(&v), e.text));
        let e =
            parse(&t).unwrap_or_else(|c| panic!("seed {seed}: {t} not read by the evaluator: {c}"));
        assert_eq!(shape(&e), want, "seed {seed}: {} -> {t}", write_datum(&v));
        let v2 = read(&t).unwrap_or_else(|e| panic!("seed {seed}: {t} not read back: {}", e.text));
        assert_eq!(
            write(&v2).unwrap(),
            t,
            "seed {seed}: {t} -> {}",
            write_datum(&v2)
        );
        assert_eq!(read(&write(&v2).unwrap()).unwrap(), v2, "seed {seed}: {t}");
        written += 1;
    }
    assert_eq!(written, 3000);
}
