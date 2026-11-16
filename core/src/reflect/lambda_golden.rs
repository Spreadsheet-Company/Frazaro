//! `KERNEL.29`'s fixture, `scripts/recalc/lambda.xlsx`, written by the core
//! (`FRAZARO_BLESS=1` writes it again): `LET`, `LAMBDA` and the functions that
//! take a function, one formula a probe on a Functions sheet over a Data
//! sheet, and the named functions it calls as the workbook's defined names,
//! which no sentence defines. Each formula is written as the build writes one
//! (`xlfn::as_filed`, dynamic where it may answer an array) and each holds
//! the value the door computes, so that `check_recalc_golden.ps1` holds it at
//! `built`; the owner's save in Excel 365 beside it, `lambda_saved.xlsx`, at
//! `no-differ`.
//!
//! Every probe is a formula Excel reads, since one it cannot read makes it
//! repair the file (`KERNEL.8`'s fifth slice); the malformed bindings are the
//! language's tests'. The recursion's probes climb a ladder of depths, so
//! that the one save says where Excel's own limit stands.

use vla_lang::sheet::{parse_a1_range, Cell, Content, Workbook};

use crate::sheet::xlfn::{as_filed, can_return_array};

const FIXTURE: &str = concat!(env!("CARGO_MANIFEST_DIR"), "/../scripts/recalc/lambda.xlsx");

/// The Data sheet: `A1:A6` 3, 1, 4, 1, 5, 9; `B1` a text; `C1:C3` 10, 20,
/// 30; `E1:F2` 1, 2 over 3, 4.
const DATA: &[(&str, &str)] = &[
    ("A1", "3"),
    ("A2", "1"),
    ("A3", "4"),
    ("A4", "1"),
    ("A5", "5"),
    ("A6", "9"),
    ("B1", "\"abc"),
    ("C1", "10"),
    ("C2", "20"),
    ("C3", "30"),
    ("E1", "1"),
    ("F1", "2"),
    ("E2", "3"),
    ("F2", "4"),
];

/// The named functions, and four names a formula reads otherwise: `rate`, a
/// name a `LET` shadows; `RATEN` and `SPOT`, a number and a reference called
/// as functions; and `ALIAS`, a name whose value is a named function's.
pub(crate) const NAMES: &[(&str, &str)] = &[
    ("SQ", "LAMBDA(x,x*x)"),
    ("PLUS1", "LAMBDA(x,x+Data!$A$1)"),
    ("TWICE", "LAMBDA(g,x,g(g(x)))"),
    ("FACT", "LAMBDA(n,IF(n<2,1,n*FACT(n-1)))"),
    ("CNT", "LAMBDA(n,IF(n=0,0,1+CNT(n-1)))"),
    ("CNTT", "LAMBDA(n,acc,IF(n=0,acc,CNTT(n-1,acc+1)))"),
    ("OUTER", "LAMBDA(x,INNER(x)*2)"),
    ("INNER", "LAMBDA(y,y+Data!$A$2)"),
    ("ISEV", "LAMBDA(n,IF(n=0,TRUE,ISOD(n-1)))"),
    ("ISOD", "LAMBDA(n,IF(n=0,FALSE,ISEV(n-1)))"),
    (
        "REV",
        "LAMBDA(s,IF(LEN(s)<2,s,REV(MID(s,2,LEN(s)))&LEFT(s,1)))",
    ),
    ("SUMSQ", "LAMBDA(r,SUM(MAP(r,LAMBDA(x,x*x))))"),
    ("GCDR", "LAMBDA(a,b,IF(b=0,a,GCDR(b,MOD(a,b))))"),
    ("rate", "Data!$A$1"),
    ("RATEN", "0.2"),
    ("SPOT", "Data!$A$2"),
    ("ALIAS", "SQ"),
];

/// The probes, one formula a row of the Functions sheet from row 1, by
/// family: `LET`; a `LAMBDA` called in place and bound; a function's value
/// read as a value; named functions and recursion, its ladder of depths; the
/// functions that take a function.
pub(crate) const PROBES: &[&str] = &[
    // LET
    "=LET(x,2,x*3)",
    "=LET(x,Data!A1,y,x+1,y*2)",
    "=LET(r,Data!A1:A3,SUM(r))",
    "=LET(r,Data!A1:A3,ROWS(r))",
    "=LET(x,1,LET(x,x+10,x))",
    "=LET(x,1/0,5)",
    "=LET(x,1/0,x+1)",
    "=LET(rate,5,rate+1)",
    "=rate",
    "=LET(x,Data!A1:A3*10,SUM(x))",
    "=LET(x,\"a\",y,x&\"b\",y&y)",
    "=LET(s,Data!B1,LEN(s)&UPPER(s))",
    "=LET(a,Data!A1,b,Data!A2,c,a*b,IF(c>2,c,-c))",
    // LAMBDA in place and bound
    "=LAMBDA(x,x+1)(5)",
    "=LAMBDA(x,y,x*y)(3,4)",
    "=LAMBDA(x,LAMBDA(y,x*y))(3)(4)",
    "=LAMBDA(42)()",
    "=LET(f,LAMBDA(x,x*x),f(3)+f(4))",
    "=LET(k,10,f,LAMBDA(x,x+k),f(1))",
    "=LET(f,LAMBDA(x,x*2),g,LAMBDA(y,f(y)+1),g(5))",
    "=LAMBDA(x,SUM(x))(Data!A1:A3)",
    "=LAMBDA(x,ROWS(x))(Data!A1:A3)",
    "=LAMBDA(x,x/0)(1)",
    "=IFERROR(LAMBDA(x,x/0)(1),\"e\")",
    "=LAMBDA(x,y,IF(ISOMITTED(y),\"m\",x+y))(1,)",
    "=LAMBDA(x,y,IF(ISOMITTED(y),\"m\",x+y))(1,2)",
    "=LAMBDA(x,y,x+y)(1,)",
    // a function's value where a value is wanted, and a count not its own
    "=LAMBDA(x,x)",
    "=LET(f,LAMBDA(x,x),f+1)",
    "=SUM(LAMBDA(x,x))",
    "=TYPE(LAMBDA(x,x))",
    "=LET(f,3,f(1))",
    "=LAMBDA(x,y,x+y)(1)",
    "=LAMBDA(x,x)(1,2)",
    "=LAMBDA(x,ISOMITTED(x))()",
    // named functions
    "=SQ(7)",
    "=SQ",
    "=PLUS1(Data!A2)",
    "=TWICE(LAMBDA(x,x+1),5)",
    "=TWICE(SQ,3)",
    "=OUTER(Data!A3)",
    "=SUMSQ(Data!A1:A3)",
    "=RATEN(1)",
    // recursion, and its ladder of depths
    "=FACT(5)",
    "=FACT(20)",
    "=FACT(170)",
    "=ISEV(10)",
    "=ISOD(7)",
    "=REV(\"abcdef\")",
    "=GCDR(1071,462)",
    "=LET(f,LAMBDA(g,n,IF(n<2,1,n*g(g,n-1))),f(f,10))",
    "=CNT(250)",
    "=CNT(500)",
    "=CNT(1000)",
    "=CNT(2000)",
    "=CNT(4095)",
    "=CNT(4096)",
    "=CNTT(300,0)",
    "=CNTT(1000,0)",
    "=CNTT(4095,0)",
    // the functions that take a function
    "=SUM(MAP(Data!A1:A3,LAMBDA(c,c*2)))",
    "=SUM(MAP(Data!A1:A3,LAMBDA(c,ROW(c))))",
    "=SUM(MAP(Data!A1:A3,Data!C1:C3,LAMBDA(a,b,a*b)))",
    "=SUM(MAP(Data!A1:A3,Data!C1:C2,LAMBDA(a,b,a+b)))",
    "=MAP(Data!A1:A3,LAMBDA(a,b,a))",
    "=SUM(MAP(SEQUENCE(4),LAMBDA(x,x^2)))",
    "=SUM(MAP(Data!A1:A3,SQ))",
    "=LET(sq,LAMBDA(x,x*x),SUM(MAP(Data!A1:A3,sq)))",
    "=MAP(Data!A1:A3,LAMBDA(c,c*2))",
    "=REDUCE(0,Data!A1:A3,LAMBDA(a,v,a+v))",
    "=REDUCE(,Data!A1:A3,LAMBDA(a,v,a*10+v))",
    "=REDUCE(1,Data!A1:A6,LAMBDA(a,v,a*v))",
    "=REDUCE(0,Data!A1:A6,LAMBDA(a,v,IF(v>a,v,a)))",
    "=SUM(SCAN(0,Data!A1:A3,LAMBDA(a,v,a+v)))",
    "=INDEX(SCAN(0,Data!A1:A3,LAMBDA(a,v,a+v)),3)",
    "=INDEX(SCAN(,Data!A1:A3,LAMBDA(a,v,a+v)),1)",
    "=SUM(BYROW(Data!E1:F2,LAMBDA(r,SUM(r))))",
    "=INDEX(BYROW(Data!E1:F2,LAMBDA(r,SUM(r))),2)",
    "=SUM(BYCOL(Data!E1:F2,LAMBDA(c,MAX(c))))",
    "=INDEX(BYROW(Data!E1:F2,LAMBDA(r,r)),1)",
    "=SUM(MAKEARRAY(2,3,LAMBDA(r,c,r*c)))",
    "=ROWS(MAKEARRAY(3,2,LAMBDA(r,c,r)))",
    "=MAKEARRAY(0,1,LAMBDA(r,c,1))",
    "=SUM(MAKEARRAY(2.7,2,LAMBDA(r,c,1)))",
    "=INDEX(MAKEARRAY(2,2,LAMBDA(r,c,r*10+c)),2,1)",
    // The second save's, confirming what the first corrected (the stop
    // rule): a function's value read as a value `#VALUE!` and a cell's answer
    // `#CALC!`; a `MAKEARRAY` count not whole `#VALUE!`; a defined name called
    // its value called.
    "=-LAMBDA(x,x)",
    "=ABS(LAMBDA(x,x))",
    "=LAMBDA(x,x)&\"a\"",
    "=IF(TRUE,LAMBDA(x,x))",
    "=SUM(MAKEARRAY(2,2.5,LAMBDA(r,c,1)))",
    "=SUM(MAKEARRAY(4/2,2,LAMBDA(r,c,1)))",
    "=SPOT(1)",
    "=ALIAS(3)",
];

/// The fixture's workbook, each formula as the build writes it.
pub(crate) fn workbook() -> Workbook {
    let mut wb = Workbook::new();
    let data = wb.ensure_sheet("Data");
    for (addr, text) in DATA {
        let r = parse_a1_range(addr).expect("an address");
        let content = match text.strip_prefix('"') {
            Some(t) => Content::Text(t.to_string()),
            None => Content::Number(text.parse().expect("a number")),
        };
        wb.sheets[data].set(r.top, r.left, Cell { content, style: 0 });
    }
    let functions = wb.ensure_sheet("Functions");
    for (k, formula) in PROBES.iter().enumerate() {
        let body = formula.strip_prefix('=').expect("a formula");
        let text = as_filed(body);
        let r = parse_a1_range(&format!("A{}", k + 1)).expect("an address");
        if can_return_array(&text) {
            wb.sheets[functions].set_formula_dynamic(r, &text, 0);
        } else {
            wb.sheets[functions].set_formula(r, &text, 0);
        }
    }
    for (name, text) in NAMES {
        wb.defined_names.push((name.to_string(), text.to_string()));
    }
    wb
}

/// The fixture's bytes: the workbook, each formula holding the door's value.
pub(crate) fn bytes() -> Vec<u8> {
    let wb = workbook();
    let library = crate::excel::library();
    let values = vla_lang::calc::Calc::run(&wb, &library);
    crate::sheet::ooxml::workbook_bytes_valued(&wb, &values).expect("a workbook that fits")
}

#[test]
fn the_lambda_fixture_is_its_builder_s() {
    let bytes = bytes();
    crate::reflect::reflect_text(&bytes, FIXTURE).unwrap_or_else(|r| panic!("{}", r.text));
    if std::env::var("FRAZARO_BLESS").as_deref() == Ok("1") {
        std::fs::write(FIXTURE, &bytes).expect("the fixture written");
        return;
    }
    let held = std::fs::read(FIXTURE)
        .unwrap_or_else(|_| panic!("{FIXTURE}: FRAZARO_BLESS=1 and the core's tests write it"));
    assert!(
        held == bytes,
        "{FIXTURE} is not what its builder writes today: FRAZARO_BLESS=1 writes it again"
    );
}

#[test]
fn the_fixture_reads_back_to_its_formulas_and_names_and_computes_alike() {
    // Read back by the door as Excel's save will be read: each formula and
    // each named function the language reads as written, every value the
    // door caches its own computation again.
    let text = crate::reflect::calc::calc_text(&bytes(), FIXTURE, &crate::excel::library())
        .unwrap_or_else(|r| panic!("{}", r.text));
    let rows: Vec<&str> = text.lines().filter(|l| l.starts_with("(calc ")).collect();
    assert_eq!(rows.len(), PROBES.len(), "{text}");
    for (row, formula) in rows.iter().zip(PROBES) {
        assert!(
            row.contains(&format!("\"{}\"", formula.replace('"', "\\\""))),
            "{row} for {formula}"
        );
        assert!(
            row.ends_with(" agree)")
                || row.contains("(not-computed")
                || row.ends_with(" none unchecked)"),
            "{row}"
        );
    }
}
