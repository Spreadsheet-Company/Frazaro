//! `LET` and `LAMBDA` (KERNEL.29): a function written in formulas computed by
//! the door's evaluator (`Calc`, every formula by `Ctx`) and by an engine's
//! machine (its shapes inlined at their placement), each held to the value
//! the formula written out computes, and the two to each other.

use vla_lang::calc::library;
use vla_lang::calc::{Calc, Computed, ErrorKind, Value};
use vla_lang::machine::Machine;
use vla_lang::sheet::{parse_a1_range, Cell, Content, Workbook};

/// A workbook from rows of `(sheet, address, content)`, a formula's text
/// beginning with `=`, and defined names of `(name, text)`.
fn book(rows: &[(&str, &str, &str)], names: &[(&str, &str)]) -> Workbook {
    let mut wb = Workbook::new();
    for (sheet, addr, text) in rows {
        let i = wb.ensure_sheet(sheet);
        let range = parse_a1_range(addr).unwrap();
        let content = if let Some(f) = text.strip_prefix('=') {
            Content::Formula(f.to_string())
        } else if let Ok(n) = text.parse::<f64>() {
            Content::Number(n)
        } else {
            Content::Text(text.to_string())
        };
        wb.sheets[i].set(range.top, range.left, Cell { content, style: 0 });
    }
    for (name, text) in names {
        wb.defined_names.push((name.to_string(), text.to_string()));
    }
    wb
}

/// Each formula cell's value as the door computes it, by its address on the
/// first sheet.
fn door(rows: &[(&str, &str, &str)], names: &[(&str, &str)], addr: &str) -> Computed {
    let wb = book(rows, names);
    let lib = library::language();
    let calc = Calc::run(&wb, &lib);
    let r = parse_a1_range(addr).unwrap();
    calc.computed((0, r.top, r.left)).cloned().unwrap()
}

fn num(n: f64) -> Computed {
    Computed::Value(Value::Number(n))
}

fn err(k: ErrorKind) -> Computed {
    Computed::Value(Value::Error(k))
}

/// One formula in `S!A1` over `S!B1:B3` holding 1, 2 and 3.
fn one(formula: &str, names: &[(&str, &str)]) -> Computed {
    door(
        &[
            ("S", "A1", formula),
            ("S", "B1", "1"),
            ("S", "B2", "2"),
            ("S", "B3", "3"),
        ],
        names,
        "A1",
    )
}

#[test]
fn a_let_binds_its_names_in_order_and_its_body_reads_them() {
    assert_eq!(one("=LET(x,2,x*3)", &[]), num(6.0));
    assert_eq!(one("=LET(x,B3,y,x+1,y*2)", &[]), num(8.0));
    // A name shadows a defined name of its spelling, and a later binding an
    // earlier one.
    assert_eq!(one("=LET(rate,5,rate+1)", &[("rate", "=S!$B$1")]), num(6.0));
    assert_eq!(one("=LET(x,1,LET(x,x+10,x))", &[]), num(11.0));
    // A binding nothing reaches is not computed.
    assert_eq!(one("=LET(x,1/0,5)", &[]), num(5.0));
    // A reference bound is a reference.
    assert_eq!(one("=LET(r,B1:B3,SUM(r)+ROWS(r))", &[]), num(9.0));
}

#[test]
fn a_malformed_binding_is_value() {
    assert_eq!(one("=LET(x,1)", &[]), err(ErrorKind::Value));
    assert_eq!(one("=LET(x,1,y,2)", &[]), err(ErrorKind::Value));
    assert_eq!(one("=LET(B1,1,2)", &[]), err(ErrorKind::Value));
    assert_eq!(one("=LAMBDA(x,x,x)(1,2)", &[]), err(ErrorKind::Value));
}

#[test]
fn a_lambda_is_called_in_place_bound_and_by_its_name() {
    assert_eq!(one("=LAMBDA(x,x+1)(5)", &[]), num(6.0));
    assert_eq!(one("=LAMBDA(x,LAMBDA(y,x*y))(3)(4)", &[]), num(12.0));
    assert_eq!(one("=LET(f,LAMBDA(x,x*x),f(3)+f(4))", &[]), num(25.0));
    assert_eq!(one("=SQ(7)", &[("SQ", "=LAMBDA(x,x*x)")]), num(49.0));
    // A named function's references are where the name says.
    assert_eq!(
        one("=PLUS1(B2)", &[("PLUS1", "=LAMBDA(x,x+S!$B$1)")]),
        num(3.0)
    );
    // A function of no parameters.
    assert_eq!(one("=LAMBDA(42)()", &[]), num(42.0));
}

#[test]
fn a_function_s_value_is_no_cell_s() {
    // A cell's answer that is a function's is `#CALC!`, and a function's
    // value read where one value is wanted `#VALUE!`, as Excel's save of the
    // fixture answers each (`scripts/recalc/lambda_saved.xlsx`).
    assert_eq!(one("=LAMBDA(x,x)", &[]), err(ErrorKind::Calc));
    assert_eq!(one("=IF(TRUE,LAMBDA(x,x))", &[]), err(ErrorKind::Calc));
    assert_eq!(
        one("=SQ", &[("SQ", "=LAMBDA(x,x*x)")]),
        err(ErrorKind::Calc)
    );
    assert_eq!(one("=LET(f,LAMBDA(x,x),f+1)", &[]), err(ErrorKind::Value));
    assert_eq!(one("=SUM(LAMBDA(x,x))", &[]), err(ErrorKind::Value));
    assert_eq!(one("=-LAMBDA(x,x)", &[]), err(ErrorKind::Value));
    assert_eq!(one("=ABS(LAMBDA(x,x))", &[]), err(ErrorKind::Value));
    assert_eq!(one("=LAMBDA(x,x)&\"a\"", &[]), err(ErrorKind::Value));
}

#[test]
fn the_information_functions_read_a_function_as_a_kind_of_its_own() {
    // Excel's third save: neither an error, a number nor a text, as `TYPE`'s
    // 128 says, and `ERROR.TYPE` of what is no error `#N/A`.
    let sq = [("SQ", "=LAMBDA(x,x*x)")];
    let truth = |b| Computed::Value(Value::Bool(b));
    assert_eq!(one("=ISERROR(SQ)", &sq), truth(false));
    assert_eq!(one("=ISNUMBER(SQ)", &sq), truth(false));
    assert_eq!(one("=ISNONTEXT(SQ)", &sq), truth(true));
    assert_eq!(one("=ERROR.TYPE(SQ)", &sq), err(ErrorKind::NA));
    assert_eq!(one("=TYPE(SQ)", &sq), num(128.0));
    // A function of values still reads it as `#VALUE!`.
    assert_eq!(one("=ABS(SQ)", &sq), err(ErrorKind::Value));
}

#[test]
fn a_count_of_arguments_not_its_parameters_is_value() {
    assert_eq!(one("=LAMBDA(x,y,x+y)(1)", &[]), err(ErrorKind::Value));
    assert_eq!(one("=LAMBDA(x,x)(1,2)", &[]), err(ErrorKind::Value));
    // A value called that is no function's.
    assert_eq!(one("=LET(f,3,f(1))", &[]), err(ErrorKind::Value));
    // A defined name called: its value called, a number's `#VALUE!` (Excel's
    // save of `RATEN(1)`), a reference's `#REF!`, a cell's, a range's and a
    // local's alike (its third save), a function's applied.
    assert_eq!(
        one("=RATEN(1)", &[("RATEN", "=0.2")]),
        err(ErrorKind::Value)
    );
    assert_eq!(one("=SPOT(1)", &[("SPOT", "=S!$B$2")]), err(ErrorKind::Ref));
    assert_eq!(
        one("=BLOCK(1)", &[("BLOCK", "=S!$B$1:$B$3")]),
        err(ErrorKind::Ref)
    );
    assert_eq!(one("=LET(r,B2,r(1))", &[]), err(ErrorKind::Ref));
    assert_eq!(
        one("=ALIAS(3)", &[("SQ", "=LAMBDA(x,x*x)"), ("ALIAS", "=SQ")]),
        num(9.0)
    );
}

#[test]
fn a_makearray_count_not_whole_is_value() {
    // Excel's save: `MAKEARRAY(2.7,2,...)` is `#VALUE!`, where `SEQUENCE`
    // reads its count down.
    assert_eq!(
        one("=SUM(MAKEARRAY(2.7,2,LAMBDA(r,c,1)))", &[]),
        err(ErrorKind::Value)
    );
    assert_eq!(
        one("=SUM(MAKEARRAY(2,2.5,LAMBDA(r,c,1)))", &[]),
        err(ErrorKind::Value)
    );
    assert_eq!(one("=SUM(MAKEARRAY(4/2,2,LAMBDA(r,c,1)))", &[]), num(4.0));
}

#[test]
fn recursion_runs_to_its_depth_and_answers_num_past_it() {
    let fact = [("FACT", "=LAMBDA(n,IF(n<2,1,n*FACT(n-1)))")];
    assert_eq!(one("=FACT(5)", &fact), num(120.0));
    let cnt = [("CNT", "=LAMBDA(n,IF(n=0,0,1+CNT(n-1)))")];
    assert_eq!(one("=CNT(100)", &cnt), num(100.0));
    // Self-application, Excel's idiom for a recursion with no name.
    assert_eq!(
        one("=LET(f,LAMBDA(g,n,IF(n<2,1,n*g(g,n-1))),f(f,5))", &[]),
        num(120.0)
    );
}

#[test]
fn recursion_as_deep_as_its_limit_runs_on_a_small_stack() {
    // VLA.1's thread: the calls past the stack's bound are computed on a
    // fresh stack and remembered (KERNEL.29).
    std::thread::Builder::new()
        .stack_size(1 << 20)
        .spawn(|| {
            let cnt = [("CNT", "=LAMBDA(n,IF(n=0,0,1+CNT(n-1)))")];
            assert_eq!(one("=CNT(4095)", &cnt), num(4095.0));
            assert_eq!(one("=CNT(4096)", &cnt), err(ErrorKind::Num));
            let fact = [("FACT", "=LAMBDA(n,IF(n<2,1,n*FACT(n-1)))")];
            let Computed::Value(Value::Number(f)) = one("=FACT(170)", &fact) else {
                panic!("170!");
            };
            assert!(f > 7.25e306 && f < 7.26e306, "{f}");
        })
        .unwrap()
        .join()
        .unwrap();
}

#[test]
fn the_helpers_call_their_function_at_each_place() {
    assert_eq!(one("=SUM(MAP(B1:B3,LAMBDA(c,c*2)))", &[]), num(12.0));
    assert_eq!(one("=SUM(MAP(B1:B3,LAMBDA(c,ROW(c))))", &[]), num(6.0));
    assert_eq!(one("=REDUCE(0,B1:B3,LAMBDA(a,v,a+v))", &[]), num(6.0));
    assert_eq!(one("=REDUCE(,B1:B3,LAMBDA(a,v,a*10+v))", &[]), num(123.0));
    assert_eq!(one("=SUM(SCAN(0,B1:B3,LAMBDA(a,v,a+v)))", &[]), num(10.0));
    assert_eq!(one("=SUM(MAKEARRAY(2,3,LAMBDA(r,c,r*c)))", &[]), num(18.0));
    assert_eq!(one("=SUM(BYROW(B1:B3,LAMBDA(r,SUM(r)*2)))", &[]), num(12.0));
    assert_eq!(one("=SUM(BYCOL(B1:B3,LAMBDA(c,SUM(c))))", &[]), num(6.0));
    assert_eq!(
        one("=SUM(MAP(B1:B3,SQ))", &[("SQ", "=LAMBDA(x,x*x)")]),
        num(14.0)
    );
    assert_eq!(one("=MAP(B1:B3,LAMBDA(a,b,a))", &[]), err(ErrorKind::Value));
}

#[test]
fn isomitted_tells_an_argument_left_out() {
    assert_eq!(
        one("=LAMBDA(x,y,IF(ISOMITTED(y),\"m\",x+y))(1,)", &[]),
        Computed::Value(Value::Text("m".to_string()))
    );
    assert_eq!(
        one("=LAMBDA(x,y,IF(ISOMITTED(y),\"m\",x+y))(1,2)", &[]),
        num(3.0)
    );
}

/// The rows of a cartridge: `S!B1:B3` holding 1, 2 and 3, the names, and one
/// formula in `S!A1`.
fn rows(formula: &str, names: &[(&str, &str)]) -> String {
    let mut out =
        String::from("(cell \"S\" \"B1\" 1)\n(cell \"S\" \"B2\" 2)\n(cell \"S\" \"B3\" 3)\n");
    for (name, text) in names {
        out.push_str(&format!(
            "(name \"{name}\" \"{}\")\n",
            text.replace('"', "\\\"")
        ));
    }
    out.push_str(&format!(
        "(formula \"S\" \"A1\" \"{}\")\n",
        formula.replace('"', "\\\"")
    ));
    out
}

/// One formula's value in an engine's machine, its shape inlined at its
/// placement, stepped to the end of its first frame.
fn engine(formula: &str, names: &[(&str, &str)]) -> Computed {
    use vla_lang::calc::Values;
    let mut m = Machine::load(&rows(formula, names), library::language()).expect("loads");
    let step = m.step(usize::MAX);
    assert!(step.done, "{step:?}");
    let s = m.find_sheet("S").expect("the sheet");
    m.computed((s, 1, 1)).expect("a value")
}

/// Every formula of the cases above, computed by an engine as the door
/// computes it.
const CASES: &[(&str, &[(&str, &str)])] = &[
    ("=LET(x,2,x*3)", &[]),
    ("=LET(x,B3,y,x+1,y*2)", &[]),
    ("=LET(rate,5,rate+1)", &[("rate", "=S!$B$1")]),
    ("=LET(x,1,LET(x,x+10,x))", &[]),
    ("=LET(x,1/0,5)", &[]),
    ("=LET(r,B1:B3,SUM(r)+ROWS(r))", &[]),
    ("=LAMBDA(x,x+1)(5)", &[]),
    ("=LAMBDA(x,LAMBDA(y,x*y))(3)(4)", &[]),
    ("=LET(f,LAMBDA(x,x*x),f(3)+f(4))", &[]),
    ("=SQ(7)", &[("SQ", "=LAMBDA(x,x*x)")]),
    ("=PLUS1(B2)", &[("PLUS1", "=LAMBDA(x,x+S!$B$1)")]),
    ("=PLUS1(B2)", &[("PLUS1", "=LAMBDA(x,x+S!B1)")]),
    ("=LAMBDA(42)()", &[]),
    ("=LAMBDA(x,x)", &[]),
    ("=LET(f,LAMBDA(x,x),f+1)", &[]),
    ("=SUM(LAMBDA(x,x))", &[]),
    ("=LAMBDA(x,y,x+y)(1)", &[]),
    ("=LET(f,3,f(1))", &[]),
    ("=RATEN(1)", &[("RATEN", "=0.2")]),
    ("=SPOT(1)", &[("SPOT", "=S!$B$2")]),
    ("=LET(r,B2,r(1))", &[]),
    ("=ISERROR(SQ)", &[("SQ", "=LAMBDA(x,x*x)")]),
    ("=ERROR.TYPE(SQ)", &[("SQ", "=LAMBDA(x,x*x)")]),
    ("=ALIAS(3)", &[("SQ", "=LAMBDA(x,x*x)"), ("ALIAS", "=SQ")]),
    (
        "=ALIAS(B3)",
        &[("PLUS1", "=LAMBDA(x,x+S!$B$1)"), ("ALIAS", "=PLUS1")],
    ),
    ("=SUM(MAKEARRAY(2.7,2,LAMBDA(r,c,1)))", &[]),
    ("=FACT(5)", &[("FACT", "=LAMBDA(n,IF(n<2,1,n*FACT(n-1)))")]),
    ("=LET(f,LAMBDA(g,n,IF(n<2,1,n*g(g,n-1))),f(f,5))", &[]),
    ("=SUM(MAP(B1:B3,LAMBDA(c,c*2)))", &[]),
    ("=REDUCE(,B1:B3,LAMBDA(a,v,a*10+v))", &[]),
    ("=SUM(MAKEARRAY(2,3,LAMBDA(r,c,r*c)))", &[]),
    ("=LAMBDA(x,y,IF(ISOMITTED(y),\"m\",x+y))(1,)", &[]),
    (
        "=TWICE(LAMBDA(x,x+1),5)",
        &[("TWICE", "=LAMBDA(g,x,g(g(x)))")],
    ),
    (
        "=OUTER(B3)",
        &[
            ("OUTER", "=LAMBDA(x,INNER(x)*2)"),
            ("INNER", "=LAMBDA(y,y+B1)"),
        ],
    ),
];

#[test]
fn an_engine_computes_each_case_as_the_door_does() {
    for (formula, names) in CASES {
        assert_eq!(engine(formula, names), one(formula, names), "{formula}");
    }
}

#[test]
fn a_shape_inlined_is_the_formula_written_out() {
    // Its census's rows but its text are the formula written out's: lowered
    // to a program, read by the reach, no call left (KERNEL.29).
    // The census with each shape's text, which it prints in R1C1, cut out.
    let census = |formula: &str, names: &[(&str, &str)]| {
        let m = Machine::load(&rows(formula, names), library::language()).expect("loads");
        m.census()
            .lines()
            .map(
                |line| match (line.find("\"A1\" \""), line.find("\" (cells")) {
                    (Some(a), Some(b)) => format!("{}{}", &line[..a + 5], &line[b..]),
                    _ => line.to_string(),
                },
            )
            .collect::<Vec<_>>()
            .join("\n")
    };
    let written = census("=B1*B1+B2", &[]);
    for (formula, names) in [
        ("=LET(x,B1,x*x+B2)", &[][..]),
        ("=SQ(B1)+B2", &[("SQ", "=LAMBDA(x,x*x)")][..]),
        ("=LAMBDA(a,b,a*a+b)(B1,B2)", &[][..]),
    ] {
        let inlined = census(formula, names);
        assert!(!inlined.contains("unlowered"), "{formula}: {inlined}");
        assert_eq!(inlined, written, "{formula}");
    }
}

#[test]
fn a_call_through_a_name_follows_the_cells_its_function_reads() {
    // `ALIAS` names `PLUS1`, which reads `S!$B$1`: the call is kept, its
    // shape computed every frame, so that a write to `B1` reaches it.
    use vla_lang::calc::Values;
    let names = [("PLUS1", "=LAMBDA(x,x+S!$B$1)"), ("ALIAS", "=PLUS1")];
    let mut m = Machine::load(&rows("=ALIAS(B3)", &names), library::language()).expect("loads");
    assert!(m.step(usize::MAX).done);
    let s = m.find_sheet("S").expect("the sheet");
    assert_eq!(m.computed((s, 1, 1)), Some(num(4.0)));
    m.write("(cell \"S\" \"B1\" 10)").unwrap();
    assert!(m.step(usize::MAX).done);
    assert_eq!(m.computed((s, 1, 1)), Some(num(13.0)));
}

/// A cell's place, its row and its column; and a bit's, its sum's and its
/// carry's.
type Place = (u32, u32);
type Bit = (Place, Place);

/// An adder of `bits` bits, a sum and a carry a bit, each written out or
/// through named functions, filled down two columns or each cell at a
/// shuffled place of D:K; its switches a all 1 and b all 0, the carry-in at
/// `G1`. The rows, and by bit the places of its sum and carry.
fn adder(bits: usize, named: bool, shuffled: bool) -> (String, Vec<Bit>) {
    fn column(n: u32) -> String {
        let (mut n, mut out) = (n, String::new());
        while n > 0 {
            out.insert(0, (b'A' + ((n - 1) % 26) as u8) as char);
            n = (n - 1) / 26;
        }
        out
    }
    let texts = |a: &str, b: &str, c: &str| {
        if named {
            (
                format!("=SUMBIT({a},{b},{c})"),
                format!("=CARRYBIT({a},{b},{c})"),
            )
        } else {
            (
                format!("=MOD(MOD({a}+{b},2)+{c},2)"),
                format!("=MAX({a}*{b},MOD({a}+{b},2)*{c})"),
            )
        }
    };
    let mut out = String::from("(cell \"S\" \"G1\" 0)\n");
    for r in 2..bits + 2 {
        out.push_str(&format!(
            "(cell \"S\" \"A{r}\" 1)\n(cell \"S\" \"B{r}\" 0)\n"
        ));
    }
    if named {
        out.push_str("(name \"SUMBIT\" \"=LAMBDA(a,b,c,MOD(MOD(a+b,2)+c,2))\")\n");
        out.push_str("(name \"CARRYBIT\" \"=LAMBDA(a,b,c,MAX(a*b,MOD(a+b,2)*c))\")\n");
    }
    let mut at = Vec::with_capacity(bits);
    if !shuffled {
        let (s, c) = texts("A2", "B2", "G1");
        let last = bits + 1;
        out.push_str(&format!("(formula \"S\" \"D2:D{last}\" \"{s}\")\n"));
        out.push_str(&format!("(formula \"S\" \"G2:G{last}\" \"{c}\")\n"));
        at.extend((2..=last as u32).map(|r| ((r, 4), (r, 7))));
        return (out, at);
    }
    // A fixed shuffle of the slots of D2:K(bits+1), two cells a bit.
    let mut slots: Vec<u32> = (0..8 * bits as u32).collect();
    let mut x: u64 = 29;
    for i in (1..slots.len()).rev() {
        x = x
            .wrapping_mul(6364136223846793005)
            .wrapping_add(1442695040888963407);
        slots.swap(i, (x >> 33) as usize % (i + 1));
    }
    let place = |k: usize| (slots[k] / 8 + 2, slots[k] % 8 + 4);
    for i in 0..bits {
        let (sum, carry) = (place(2 * i), place(2 * i + 1));
        let cin = match i {
            0 => "G1".to_string(),
            _ => {
                let (r, c) = place(2 * i - 1);
                format!("{}{r}", column(c))
            }
        };
        let r = i + 2;
        let (s, c) = texts(&format!("A{r}"), &format!("B{r}"), &cin);
        out.push_str(&format!(
            "(formula \"S\" \"{}{}\" \"{s}\")\n",
            column(sum.1),
            sum.0
        ));
        out.push_str(&format!(
            "(formula \"S\" \"{}{}\" \"{c}\")\n",
            column(carry.1),
            carry.0
        ));
        at.push((sum, carry));
    }
    (out, at)
}

/// A census with each shape's text cut out, which names a call where its
/// twin writes the formula out.
fn census_but_texts(m: &Machine) -> String {
    m.census()
        .lines()
        .map(|line| {
            let Some(end) = line.find("\" (cells") else {
                return line.to_string();
            };
            let mut seps = line.match_indices("\" \"").map(|(i, _)| i);
            match (seps.next(), seps.next()) {
                (Some(_), Some(text)) if text < end => {
                    format!("{}{}", &line[..text + 3], &line[end..])
                }
                _ => line.to_string(),
            }
        })
        .collect::<Vec<_>>()
        .join("\n")
}

#[test]
fn a_netlist_through_named_functions_plans_and_computes_as_written_out() {
    // Slice 2's parity (KERNEL.29): an adder whose sum and carry are named
    // functions is the adder written out, filled or each cell at its own
    // place, row for row of the census but its texts and value for value
    // before and after the carry-in ripples through every bit; and Life's
    // rule through a named function plans as Life's rule written out.
    use vla_lang::calc::Values;
    for shuffled in [false, true] {
        let (written, at) = adder(64, false, shuffled);
        let (named, _) = adder(64, true, shuffled);
        let lib = || library::language();
        let mut a = Machine::load(&written, lib()).expect("loads");
        let mut b = Machine::load(&named, lib()).expect("loads");
        assert_eq!(
            census_but_texts(&a),
            census_but_texts(&b),
            "shuffled {shuffled}"
        );
        let s = a.find_sheet("S").expect("the sheet");
        for edit in [None, Some("(cell \"S\" \"G1\" 1)")] {
            if let Some(row) = edit {
                a.write(row).unwrap();
                b.write(row).unwrap();
            }
            assert!(a.step(usize::MAX).done && b.step(usize::MAX).done);
            for &((sr, sc), (cr, cc)) in &at {
                for (r, c) in [(sr, sc), (cr, cc)] {
                    assert_eq!(a.computed((s, r, c)), b.computed((s, r, c)), "{r},{c}");
                }
            }
        }
        // The carry-in reached the last bit.
        let (_, last) = *at.last().unwrap();
        assert_eq!(b.computed((s, last.0, last.1)), Some(num(1.0)));
    }
    let rule = |text: &str, names: &[(&str, &str)]| {
        let mut rows = String::from("(cell \"T\" \"B2\" 1)\n(cell \"S\" \"B1\" 0)\n");
        for (name, t) in names {
            rows.push_str(&format!("(name \"{name}\" \"{t}\")\n"));
        }
        rows.push_str(&format!("(formula \"S\" \"D1:F3\" \"{text}\")\n"));
        census_but_texts(&Machine::load(&rows, library::language()).expect("loads"))
    };
    assert_eq!(
        rule(
            "=IF($B$1=1,0,IF(OR(SUM(T!A1:C3)-T!B2=3,AND(T!B2=1,SUM(T!A1:C3)-T!B2=2)),1,0))",
            &[]
        ),
        rule(
            "=IF($B$1=1,0,NEXTCELL(SUM(T!A1:C3)-T!B2,T!B2))",
            &[("NEXTCELL", "=LAMBDA(n,a,IF(OR(n=3,AND(a=1,n=2)),1,0))")]
        )
    );
}
