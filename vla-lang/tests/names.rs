//! A grid's defined names viewed and written (KERNEL.29, its third slice's
//! first part): the `names` projection, each name as the `name` row a load
//! reads, which a save prints beside every sheet's `fills`; and a `name` row
//! in a write, the name defined or its text replaced, every shape reading it
//! planned again and computed at the next frame, as a load of the written
//! grid would plan and compute it.

use vla_lang::calc::library;
use vla_lang::calc::Values;
use vla_lang::machine::{Draft, Machine, Viewed};

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

/// A save as a host prints one: every sheet's fills, then the names.
fn save(m: &Machine) -> String {
    let mut out: String = (0..m.sheets())
        .filter_map(|i| m.sheet_name(i))
        .map(|s| text(m, "fills", s))
        .collect();
    out.push_str(&names(m));
    out
}

/// A fresh machine over the grid a machine holds, its names among it, as
/// law 7 makes one (`tests/laws.rs`).
fn again(m: &Machine) -> Machine {
    let mut workbook = m.workbook();
    workbook.sheets.truncate(m.sheets());
    Machine::new(Draft::of_workbook(&workbook), library::language())
        .unwrap_or_else(|e| panic!("line {}: {}", e.line, e.refusal.text))
}

fn number(m: &Machine, sheet: &str, row: u32, col: u32) -> f64 {
    let si = m.find_sheet(sheet).expect("the sheet");
    match m.computed((si, row, col)) {
        Some(vla_lang::calc::Computed::Value(vla_lang::calc::Value::Number(n))) => n,
        other => panic!("{sheet}!R{row}C{col}: {other:?}"),
    }
}

#[test]
fn the_names_view_prints_each_name_as_a_load_reads_it() {
    let rows = "(sheet \"S\" visible)\n\
                (cell \"S\" \"A1\" 2)\n\
                (name \"rate\" \"0.25\")\n\
                (name \"TWICE\" \"=LAMBDA(x,x*2)\")\n\
                (name \"S!Local\" \"S!$A$1\")\n\
                (name \"said\" \"\\\"a \\\"\\\"quoted\\\"\\\" text\\\"\")\n\
                (formula \"S\" \"B1\" \"=TWICE(Local)+rate\")\n";
    let m = load(rows);
    let printed = names(&m);
    assert_eq!(
        printed,
        "(name \"rate\" \"0.25\")\n\
         (name \"TWICE\" \"=LAMBDA(x,x*2)\")\n\
         (name \"S!Local\" \"S!$A$1\")\n\
         (name \"said\" \"\\\"a \\\"\\\"quoted\\\"\\\" text\\\"\")\n"
    );
    // The sheet and the window asked are not read.
    assert_eq!(
        m.view("names", "S", Some("A1")),
        Ok(Viewed::Text(printed.clone()))
    );
    // A grid with no name prints none.
    assert_eq!(names(&load("(cell \"S\" \"A1\" 1)\n")), "");
    // A save of the grid loads the names it printed, and computes alike.
    let mut m = m;
    assert!(m.step(0).done);
    let mut back = load(&save(&m));
    assert!(back.step(0).done);
    assert_eq!(names(&back), printed);
    assert_eq!(values(&back), values(&m));
    assert_eq!(number(&m, "S", 1, 2), 4.25);
}

#[test]
fn a_name_row_written_is_read_as_a_load_reads_one() {
    let mut m = load(
        "(cell \"S\" \"A1\" 1)\n\
         (name \"F\" \"=LAMBDA(x,x+1)\")\n\
         (formula \"S\" \"B1\" \"=F(A1)\")\n\
         (formula \"S\" \"C1\" \"=K+1\")\n",
    );
    assert!(m.step(0).done);
    assert_eq!(number(&m, "S", 1, 2), 2.0);
    // The text it holds already writes nothing; spelled without case, it is
    // the same name.
    assert_eq!(m.write("(name \"F\" \"=LAMBDA(x,x+1)\")"), Ok(0));
    assert_eq!(m.write("(name \"f\" \"=LAMBDA(x,x+10)\")"), Ok(1));
    assert!(m.step(0).done);
    assert_eq!(number(&m, "S", 1, 2), 11.0);
    assert_eq!(
        names(&m),
        "(name \"f\" \"=LAMBDA(x,x+10)\")\n",
        "the name's place kept, its spelling the write's"
    );
    // A name the grid lacked, defined: the formula reading it computes.
    assert_eq!(m.write("(name \"K\" \"41\")"), Ok(1));
    assert!(m.step(0).done);
    assert_eq!(number(&m, "S", 1, 3), 42.0);
    // A row the load would refuse, refused alike, and nothing changed.
    let before = names(&m);
    for bad in [
        "(name \"F\")",
        "(name \"\" \"1\")",
        "(name \"F\" \"\")",
        "(name F \"1\")",
    ] {
        let e = m.write(bad).unwrap_err();
        assert_eq!(
            e.refusal.id, "grid-row-malformed",
            "{bad}: {}",
            e.refusal.text
        );
        assert_eq!(names(&m), before, "{bad}");
    }
    // A later row refused undoes the name before it.
    let e = m
        .write("(name \"K\" \"1\")\n(cell \"Nowhere\" \"A1\" 1)")
        .unwrap_err();
    assert_eq!(e.refusal.id, "grid-sheet-unknown");
    assert_eq!(names(&m), before);
    assert!(m.step(0).done);
    assert_eq!(number(&m, "S", 1, 3), 42.0);
}

#[test]
fn a_name_written_computes_its_readers_alone() {
    let mut rows = String::from("(name \"F\" \"=LAMBDA(x,x+1)\")\n");
    for r in 1..=100 {
        rows.push_str(&format!("(cell \"S\" \"B{r}\" {r})\n"));
    }
    rows.push_str("(formula \"S\" \"A1:A100\" \"=F(B1)\")\n");
    rows.push_str("(formula \"S\" \"C1:C100\" \"=B1*2\")\n");
    rows.push_str("(formula \"S\" \"D1\" \"=SUM(A1:A100)\")\n");
    let mut m = load(&rows);
    assert!(m.step(0).done);
    assert_eq!(m.reached(), 201);
    assert_eq!(number(&m, "S", 1, 4), 5150.0);
    // A frame that changes nothing, so that the first frame's changes, every
    // formula's, are no longer carried: past the room a write's next frame
    // is whole, whatever it wrote.
    assert!(m.step(0).done);
    let census = m.census();
    assert_eq!(m.write("(name \"F\" \"=LAMBDA(x,x+2)\")"), Ok(1));
    assert!(m.step(0).done);
    // The hundred cells calling F and the sum reading them, and not the
    // hundred that read neither.
    assert_eq!(m.reached(), 101);
    assert_eq!(number(&m, "S", 1, 1), 3.0);
    assert_eq!(number(&m, "S", 1, 4), 5250.0);
    // The plan is the one it was: the call written in again, no shape more.
    assert_eq!(m.census(), census);
    assert_eq!(m.census(), again(&m).census());
}

#[test]
fn a_name_written_that_leaves_the_grid_unstepped_is_refused_and_changes_nothing() {
    let mut m = load(
        "(cell \"S\" \"B1\" 1)\n\
         (name \"R\" \"S!$B$1\")\n\
         (name \"F\" \"=LAMBDA(x,x+1)\")\n\
         (formula \"S\" \"A1\" \"=R+1\")\n\
         (formula \"S\" \"C1\" \"=F(B1)\")\n",
    );
    assert!(m.step(0).done);
    let (before, census, seen) = (names(&m), m.census(), values(&m));
    // A cell reading itself through its name.
    let e = m.write("(name \"R\" \"S!$A$1\")").unwrap_err();
    assert_eq!(e.refusal.id, "calc-cycle", "{}", e.refusal.text);
    assert_eq!(e.line, 1);
    // A function calling one no library holds.
    let e = m
        .write("(name \"F\" \"=LAMBDA(x,NOSUCH(x))\")")
        .unwrap_err();
    assert_eq!(
        e.refusal.id, "calc-function-not-computed",
        "{}",
        e.refusal.text
    );
    for (what, now) in [("names", names(&m)), ("census", m.census())] {
        assert_eq!(
            now,
            if what == "names" {
                before.clone()
            } else {
                census.clone()
            },
            "{what}"
        );
    }
    assert!(m.step(0).done);
    assert_eq!(values(&m), seen);
    // And the grid still takes a name it can step.
    assert_eq!(m.write("(name \"R\" \"S!$B$1*10\")"), Ok(1));
    assert!(m.step(0).done);
    assert_eq!(number(&m, "S", 1, 1), 11.0);
}

/// A grid of value cells, names and the formulas reading them: `A1:A6`
/// numbers; three named functions, each calling only those before it, and
/// two names of cells; and formulas over `B1:F6`, by fill and alone, each
/// reading the values, calling a function or reading a name. No formula
/// reads a twin, so that a frame's values are the grid's alone.
fn named_rows(r: &mut Rng) -> String {
    let mut s = String::from("(sheet \"G\" visible)\n");
    for row in 1..=6 {
        s.push_str(&format!("(cell \"G\" \"A{row}\" {})\n", r.below(9)));
    }
    s.push_str("(name \"F1\" \"=LAMBDA(x,x*2+1)\")\n");
    s.push_str("(name \"F2\" \"=LAMBDA(x,y,F1(x)+y)\")\n");
    s.push_str("(name \"F3\" \"=LAMBDA(x,IF(x>3,F2(x,1),F1(x)))\")\n");
    s.push_str("(name \"R1\" \"G!$A$1\")\n");
    s.push_str("(name \"R2\" \"G!$A$2:$A$4\")\n");
    let operand = |r: &mut Rng, row: u32| -> String {
        match r.below(6) {
            0 => format!("F1(A{row})"),
            1 => format!("F2(A{row},R1)"),
            2 => format!("F3(A{row})"),
            3 => "SUM(R2)".to_string(),
            4 => "R1".to_string(),
            _ => format!("A{row}"),
        }
    };
    for (k, col) in ["B", "C", "D", "E", "F"].iter().enumerate() {
        if r.below(2) == 0 {
            // A fill down the column, its rows reading their own.
            let f = format!("{}+{}", operand(r, 1), operand(r, 1));
            s.push_str(&format!("(formula \"G\" \"{col}1:{col}6\" \"={f}\")\n"));
            continue;
        }
        for row in 1..=6u32 {
            if r.below(3) == 0 {
                continue;
            }
            let f = format!("{}*{}+{k}", operand(r, row), operand(r, row));
            s.push_str(&format!("(formula \"G\" \"{col}{row}\" \"={f}\")\n"));
        }
    }
    s
}

/// A write between two frames: a name's text replaced, a function's or a
/// cell's, a name the grid holds written with its own text, or a value cell,
/// alone or beside the name.
fn name_write(r: &mut Rng) -> String {
    let function = |r: &mut Rng| match r.below(4) {
        0 => format!(
            "(name \"F1\" \"=LAMBDA(x,x*{}+{})\")",
            1 + r.below(3),
            r.below(5)
        ),
        1 => format!("(name \"F2\" \"=LAMBDA(x,y,F1(y)-x*{})\")", r.below(4)),
        2 => format!("(name \"F3\" \"=LAMBDA(x,F2(x,x)+{})\")", r.below(3)),
        _ => "(name \"F1\" \"=LAMBDA(x,x*2+1)\")".to_string(),
    };
    match r.below(5) {
        0 | 1 => function(r),
        2 => format!("(name \"R1\" \"G!$A${}\")", 1 + r.below(6)),
        3 => format!(
            "(name \"R2\" \"G!$A${}:$A$6\")\n(cell \"G\" \"A{}\" {})",
            1 + r.below(5),
            1 + r.below(6),
            r.below(9)
        ),
        _ => format!("{}\n{}", function(r), function(r)),
    }
}

#[test]
fn a_name_written_is_planned_and_computed_as_a_load_of_the_written_grid() {
    let mut written = 0;
    for seed in 1..=60u64 {
        let mut r = Rng(seed);
        let rows = named_rows(&mut r);
        let mut m = load(&rows);
        for frame in 1..=4 {
            assert!(m.step(0).done);
            let write = name_write(&mut r);
            let when = format!("seed {seed}, frame {frame}: {write}\n{rows}");
            m.write(&write)
                .unwrap_or_else(|e| panic!("{when}: {}", e.refusal.text));
            written += 1;
            assert!(m.step(0).done);
            let mut fresh = again(&m);
            assert!(fresh.step(0).done);
            assert_eq!(m.census(), fresh.census(), "{when}");
            assert_eq!(values(&m), values(&fresh), "{when}");
            assert_eq!(names(&m), names(&fresh), "{when}");
            // And a save of it loads the grid it printed.
            let mut back = load(&save(&m));
            assert!(back.step(0).done);
            assert_eq!(values(&back), values(&m), "{when}: the save");
        }
    }
    assert_eq!(written, 240);
}
