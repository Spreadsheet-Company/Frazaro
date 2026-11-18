//! `KERNEL.55`'s probe, `scripts/recalc/spills.xlsx`, written by the core
//! (`FRAZARO_BLESS=1` writes it again), test-only: every case the item's oracle
//! names, each formula held as its case needs, a dynamic formula over its own
//! cell, a legacy formula or a legacy array formula over the block it names,
//! which no sentence writes today. The owner opens it in Excel 365, writes the
//! Typed sheet's texts in beside them through `Range.Formula2`, sets `Data!K1`
//! to 1 and saves it as `spills_saved.xlsx`: Excel's answer for each case, the
//! extent it saves for each spill, and the kind its own writer gives each typed
//! text, the evidence every slice of the item is held to (the scoping's stop
//! rule: two saves at most, the second only to confirm what the first
//! corrected).
//!
//! Five sheets. *Data*, the values read. *Spills*, dynamic formulas, each with
//! room beside it but where its case blocks it: an extent fixed and one moving
//! with `Data!K1`; blocked by a value, a formula, a merged range and another
//! spill; meeting the sheet's edges; read through `A1#`, `_xlfn.ANCHORARRAY` in
//! the file, by its address and by a range; a spill of a spill. *Blocks*,
//! legacy array formulas over many cells, and of many values in one. *Legacy*,
//! `KERNEL.8`'s unwitnessed edges, legacy formulas beside their dynamic twins
//! and their legacy array twins of one cell. *Typed*, texts, every tenth row,
//! for the owner's `Formula2`. No formula holds a cached value: Excel computes
//! each as it opens the file. Beside each probe, the answer the language's
//! reading predicts, the save's to tell.

use vla_lang::sheet::parse_a1_range;

use super::looks_golden::{Book, Part, Styles, CALIBRI};
use crate::sheet::xlfn::as_filed;
use crate::sheet::xml;

const FIXTURE: &str = concat!(env!("CARGO_MANIFEST_DIR"), "/../scripts/recalc/spills.xlsx");

/// How a probe's formula is held.
#[derive(Clone, Copy)]
enum Held {
    /// A legacy formula.
    Plain,
    /// A dynamic-array formula over its own cell, as Excel 365 stores one typed
    /// and its save writes its spill's extent.
    Dynamic,
    /// A legacy array formula, entered with Ctrl+Shift+Enter over the block it
    /// names, the anchor its first cell.
    Block(&'static str),
}

use Held::{Block, Dynamic, Plain};

/// The sheets, in tab order.
const SHEETS: [&str; 5] = ["Data", "Spills", "Blocks", "Legacy", "Typed"];

/// The values, a number, or a text after `"`.
const VALUES: &[(&str, &str, &str)] = &[
    // Numbers, A1:A6; texts, B1:B6; numbers with a 0, C1:C6; a column with D3
    // empty, D1:D5; a block, E1:G2; a row of keys over their values, H1:J2.
    ("Data", "A1", "3"),
    ("Data", "A2", "1"),
    ("Data", "A3", "4"),
    ("Data", "A4", "1"),
    ("Data", "A5", "5"),
    ("Data", "A6", "9"),
    ("Data", "B1", "\"a"),
    ("Data", "B2", "\"bb"),
    ("Data", "B3", "\"ccc"),
    ("Data", "B4", "\"dddd"),
    ("Data", "B5", "\"e"),
    ("Data", "B6", "\"ff"),
    ("Data", "C1", "10"),
    ("Data", "C2", "0"),
    ("Data", "C3", "30"),
    ("Data", "C4", "40"),
    ("Data", "C5", "50"),
    ("Data", "C6", "60"),
    ("Data", "D1", "1"),
    ("Data", "D2", "2"),
    ("Data", "D4", "4"),
    ("Data", "D5", "5"),
    ("Data", "E1", "1"),
    ("Data", "F1", "2"),
    ("Data", "G1", "3"),
    ("Data", "E2", "4"),
    ("Data", "F2", "5"),
    ("Data", "G2", "6"),
    ("Data", "H1", "\"a"),
    ("Data", "I1", "\"b"),
    ("Data", "J1", "\"c"),
    ("Data", "H2", "10"),
    ("Data", "I2", "20"),
    ("Data", "J2", "30"),
    // The edit's cell: 3 as built, 1 in the save, so that FILTER keeps a row
    // more.
    ("Data", "K1", "3"),
    // Sorted keys and their values, M1:N5.
    ("Data", "M1", "1"),
    ("Data", "M2", "3"),
    ("Data", "M3", "4"),
    ("Data", "M4", "5"),
    ("Data", "M5", "9"),
    ("Data", "N1", "10"),
    ("Data", "N2", "20"),
    ("Data", "N3", "30"),
    ("Data", "N4", "40"),
    ("Data", "N5", "50"),
    // INDEX's row and column at entry: one value and a row; a range of rows.
    ("Data", "X1", "3"),
    ("Data", "X2", "0"),
    ("Data", "X3", "2"),
    ("Data", "X4", "5"),
    // A value in a spill's way.
    ("Spills", "F3", "\"x"),
];

/// The probes, each its sheet, its cell, how it is held and its formula as a
/// person types it, `A1#` written `ANCHORARRAY(A1)`; the predicted answer in
/// the comment before it, a spill's cells in order.
const PROBES: &[(&str, &str, Held, &str)] = &[
    // -- Spills. An extent fixed: A1:B3, 1 to 6.
    ("Spills", "A1", Dynamic, "=SEQUENCE(3,2)"),
    // An extent moving at an edit: D1:D3 4, 5, 9 as built, D1:D4 3, 4, 5, 9
    // in the save.
    (
        "Spills",
        "D1",
        Dynamic,
        "=FILTER(Data!A1:A6,Data!A1:A6>Data!K1)",
    ),
    // Blocked by a value (F3), by a formula (G3) and by a merged range
    // (H3:I3): #SPILL! each.
    ("Spills", "F1", Dynamic, "=SEQUENCE(4)"),
    ("Spills", "G1", Dynamic, "=SEQUENCE(4)"),
    ("Spills", "G3", Plain, "=1+1"),
    ("Spills", "H1", Dynamic, "=SEQUENCE(4)"),
    // Two spills meeting at K3, K1:K4 and J3:L3: one #SPILL!, the order
    // Excel computes them in saying which.
    ("Spills", "K1", Dynamic, "=SEQUENCE(4)"),
    ("Spills", "J3", Dynamic, "=SEQUENCE(1,3)"),
    // A spill of a spill, N1:O3 10 to 60; of the moving one sorted down, Q1:Q4
    // 9, 5, 4, 3 in the save.
    ("Spills", "N1", Dynamic, "=ANCHORARRAY(A1)*10"),
    ("Spills", "Q1", Dynamic, "=SORT(ANCHORARRAY(D1),1,-1)"),
    // A range holding an empty cell, S1:S5 1, 2, 0, 4, 5; errors, U1:U3 0.1,
    // #DIV/0!, 1/30; texts, W1:W3 a!, bb!, ccc!; truth values, Y1:Y3 TRUE,
    // FALSE, TRUE.
    ("Spills", "S1", Dynamic, "=Data!D1:D5"),
    ("Spills", "U1", Dynamic, "=1/Data!C1:C3"),
    ("Spills", "W1", Dynamic, "=Data!B1:B3&\"!\""),
    ("Spills", "Y1", Dynamic, "=Data!A1:A3>2"),
    // An array of one value, 1 and no spill; of none, #CALC!.
    ("Spills", "AA1", Dynamic, "=SEQUENCE(1)"),
    (
        "Spills",
        "AC1",
        Dynamic,
        "=FILTER(Data!A1:A6,Data!A1:A6>100)",
    ),
    // The functions that take a function: AE1:AE3 6, 2, 8; AG1:AG2 6, 15;
    // AI1:AK2 1, 2, 3 over 2, 4, 6.
    ("Spills", "AE1", Dynamic, "=MAP(Data!A1:A3,LAMBDA(c,c*2))"),
    (
        "Spills",
        "AG1",
        Dynamic,
        "=BYROW(Data!E1:G2,LAMBDA(r,SUM(r)))",
    ),
    ("Spills", "AI1", Dynamic, "=MAKEARRAY(2,3,LAMBDA(r,c,r*c))"),
    // INDEX at a 0 row, AM1:AM2 2, 5; at a 0 column, AO1:AQ1 4, 5, 6;
    // TRANSPOSE, AS1:AU1 3, 1, 4.
    ("Spills", "AM1", Dynamic, "=INDEX(Data!E1:G2,0,2)"),
    ("Spills", "AO1", Dynamic, "=INDEX(Data!E1:G2,2,0)"),
    ("Spills", "AS1", Dynamic, "=TRANSPOSE(Data!A1:A3)"),
    // The sheet's edges: past its last row and past its last column #SPILL!;
    // to each exactly, B1048575:B1048576 1, 2 and XFC5:XFD5 1, 2, each read,
    // 20.
    ("Spills", "A1048575", Dynamic, "=SEQUENCE(3)"),
    ("Spills", "B1048575", Dynamic, "=SEQUENCE(2)"),
    ("Spills", "C1048575", Plain, "=B1048576*10"),
    ("Spills", "XFC1", Dynamic, "=SEQUENCE(1,3)"),
    ("Spills", "XFC5", Dynamic, "=SEQUENCE(1,2)"),
    ("Spills", "XFB5", Plain, "=XFD5*10"),
    // A1# read: 21, 3, 2, 6.
    ("Spills", "A15", Dynamic, "=SUM(ANCHORARRAY(A1))"),
    ("Spills", "A16", Dynamic, "=ROWS(ANCHORARRAY(A1))"),
    ("Spills", "A17", Dynamic, "=COLUMNS(ANCHORARRAY(A1))"),
    ("Spills", "A18", Dynamic, "=INDEX(ANCHORARRAY(A1),3,2)"),
    // A spilled cell read by its address, 6, and by a range, 21; the moving
    // extent's rows, 4 in the save.
    ("Spills", "A19", Plain, "=B3"),
    ("Spills", "A20", Plain, "=SUM(A1:B3)"),
    ("Spills", "A21", Dynamic, "=ROWS(ANCHORARRAY(D1))"),
    // A1# of a spilled cell and of a value, #REF! each; of a blocked anchor,
    // #SPILL! or #REF!.
    ("Spills", "A22", Dynamic, "=SUM(ANCHORARRAY(B2))"),
    ("Spills", "A23", Dynamic, "=SUM(ANCHORARRAY(Data!A1))"),
    ("Spills", "A24", Dynamic, "=SUM(ANCHORARRAY(F1))"),
    // A blocked anchor read: #SPILL!, TRUE, 9 and caught; a cell of its
    // would-be extent, 0 and TRUE.
    ("Spills", "A25", Plain, "=F1"),
    ("Spills", "A26", Plain, "=ISERROR(F1)"),
    ("Spills", "A27", Plain, "=ERROR.TYPE(F1)"),
    ("Spills", "A28", Plain, "=IFERROR(F1,\"caught\")"),
    ("Spills", "A29", Plain, "=F2"),
    ("Spills", "A30", Plain, "=ISBLANK(F2)"),
    // The two that meet, each read: one 1, the other blocked.
    ("Spills", "A31", Plain, "=IFERROR(K1,\"blocked\")"),
    ("Spills", "A32", Plain, "=IFERROR(J3,\"blocked\")"),
    // A spill of a spill read by a range, 210; the empty cell's place, 0 and
    // FALSE; an error's place, #DIV/0! and TRUE.
    ("Spills", "A33", Plain, "=SUM(N1:O3)"),
    ("Spills", "A34", Plain, "=S3"),
    ("Spills", "A35", Plain, "=ISBLANK(S3)"),
    ("Spills", "A36", Plain, "=U2"),
    ("Spills", "A37", Plain, "=ISERROR(U2)"),
    // Blocked by a formula and by a merged range, 9 each; past the sheet's
    // edges, 9 each; an array of none, 14.
    ("Spills", "A38", Plain, "=ERROR.TYPE(G1)"),
    ("Spills", "A39", Plain, "=ERROR.TYPE(H1)"),
    ("Spills", "A40", Plain, "=ERROR.TYPE(A1048575)"),
    ("Spills", "A41", Plain, "=ERROR.TYPE(XFC1)"),
    ("Spills", "A42", Plain, "=ERROR.TYPE(AC1)"),
    // A spill of texts counted through its anchor, 3.
    ("Spills", "A43", Dynamic, "=COUNTA(ANCHORARRAY(W1))"),
    // -- Blocks, each cell its place of the array. Longer than its block,
    // A1:A3 6, 2, 8; shorter, B1:B8 6, 2, 8, 2, 10, 18, #N/A, #N/A.
    ("Blocks", "A1", Block("A1:A3"), "=Data!A1:A6*2"),
    ("Blocks", "B1", Block("B1:B8"), "=Data!A1:A6*2"),
    // A row over three rows, C1:E3 2, 4, 6 in each; a column over three
    // columns, F1:H3 30, 10, 40 down each.
    ("Blocks", "C1", Block("C1:E3"), "=Data!E1:G1*2"),
    ("Blocks", "F1", Block("F1:H3"), "=Data!A1:A3*10"),
    // One value over a block, 5 at each of I1:J2; SUM's, 23 at K1 and K2.
    ("Blocks", "I1", Block("I1:J2"), "=5"),
    ("Blocks", "K1", Block("K1:K2"), "=SUM(Data!A1:A6)"),
    // TRANSPOSE, L1:N1 3, 1, 4; ROW of a range, O1:O6 1 to 6; INDEX at a 0
    // row, P1:P2 2, 5.
    ("Blocks", "L1", Block("L1:N1"), "=TRANSPOSE(Data!A1:A3)"),
    ("Blocks", "O1", Block("O1:O6"), "=ROW(Data!A1:A6)"),
    ("Blocks", "P1", Block("P1:P2"), "=INDEX(Data!E1:G2,0,2)"),
    // A 2-by-3 array over a 3-by-2 block, Q1:R3 1, 2 over 4, 5 over #N/A,
    // #N/A.
    ("Blocks", "Q1", Block("Q1:R3"), "=Data!E1:G2"),
    // Many values in a single cell, its first: 6; SEQUENCE's, 1.
    ("Blocks", "S1", Block("S1"), "=Data!A1:A6*2"),
    ("Blocks", "T1", Block("T1"), "=SEQUENCE(3)"),
    // A block read: 16; past its array, #N/A; a repeated row's place, 2; the
    // repeated block, 36; TRUE.
    ("Blocks", "A12", Plain, "=SUM(A1:A3)"),
    ("Blocks", "A13", Plain, "=B7"),
    ("Blocks", "A14", Plain, "=C3"),
    ("Blocks", "A15", Plain, "=SUM(C1:E3)"),
    ("Blocks", "A16", Plain, "=ISNA(B8)"),
    // -- Legacy: a legacy formula in column A, its dynamic twin in B and its
    // legacy array twin of one cell in C, each row's place the intersection's.
    // HLOOKUP's value sought given an array: legacy 20, the first forced;
    // dynamic and array 50.
    (
        "Legacy",
        "A1",
        Plain,
        "=SUM(HLOOKUP({\"b\",\"c\"},Data!H1:J2,2,FALSE))",
    ),
    (
        "Legacy",
        "B1",
        Dynamic,
        "=SUM(HLOOKUP({\"b\",\"c\"},Data!H1:J2,2,FALSE))",
    ),
    (
        "Legacy",
        "C1",
        Block("C1"),
        "=SUM(HLOOKUP({\"b\",\"c\"},Data!H1:J2,2,FALSE))",
    ),
    // VLOOKUP's match kind given an array: legacy 30, the first forced;
    // dynamic and array 60.
    (
        "Legacy",
        "A2",
        Plain,
        "=SUM(VLOOKUP(4,Data!M1:N5,2,{FALSE,TRUE}))",
    ),
    (
        "Legacy",
        "B2",
        Dynamic,
        "=SUM(VLOOKUP(4,Data!M1:N5,2,{FALSE,TRUE}))",
    ),
    (
        "Legacy",
        "C2",
        Block("C2"),
        "=SUM(VLOOKUP(4,Data!M1:N5,2,{FALSE,TRUE}))",
    ),
    // INDEX's row given a range, met at row 3 (2): legacy 1; dynamic and
    // array 6.
    ("Legacy", "A3", Plain, "=SUM(INDEX(Data!A1:A6,Data!X3:X4))"),
    (
        "Legacy",
        "B3",
        Dynamic,
        "=SUM(INDEX(Data!A1:A6,Data!X3:X4))",
    ),
    (
        "Legacy",
        "C3",
        Block("C3"),
        "=SUM(INDEX(Data!A1:A6,Data!X3:X4))",
    ),
    // INDEX at a 0 row, CHOOSE of a range and LOOKUP of a range, met at their
    // rows: legacy 1, 5 and 50; their arrays of one cell their first, 3, 3 and
    // 20.
    ("Legacy", "A4", Plain, "=INDEX(Data!A1:A6,0)"),
    ("Legacy", "C4", Block("C4"), "=INDEX(Data!A1:A6,0)"),
    ("Legacy", "A5", Plain, "=CHOOSE(1,Data!A1:A6)"),
    ("Legacy", "C5", Block("C5"), "=CHOOSE(1,Data!A1:A6)"),
    (
        "Legacy",
        "A6",
        Plain,
        "=LOOKUP(Data!A1:A6,Data!M1:M5,Data!N1:N5)",
    ),
    (
        "Legacy",
        "C6",
        Block("C6"),
        "=LOOKUP(Data!A1:A6,Data!M1:M5,Data!N1:N5)",
    ),
    // HLOOKUP's match kind given an array: legacy 20; dynamic and array 40.
    (
        "Legacy",
        "A7",
        Plain,
        "=SUM(HLOOKUP(\"b\",Data!H1:J2,2,{FALSE,TRUE}))",
    ),
    (
        "Legacy",
        "B7",
        Dynamic,
        "=SUM(HLOOKUP(\"b\",Data!H1:J2,2,{FALSE,TRUE}))",
    ),
    (
        "Legacy",
        "C7",
        Block("C7"),
        "=SUM(HLOOKUP(\"b\",Data!H1:J2,2,{FALSE,TRUE}))",
    ),
    // INDEX's row and its column, CHOOSE's index and LOOKUP's value sought
    // given arrays: dynamic and array 8, 4, 4 and 40; legacy, Excel's to tell,
    // the language's 8, 4, 4 and 40 (lifted).
    ("Legacy", "A8", Plain, "=SUM(INDEX(Data!A1:A6,{1,2,3}))"),
    ("Legacy", "B8", Dynamic, "=SUM(INDEX(Data!A1:A6,{1,2,3}))"),
    (
        "Legacy",
        "C8",
        Block("C8"),
        "=SUM(INDEX(Data!A1:A6,{1,2,3}))",
    ),
    ("Legacy", "A9", Plain, "=SUM(INDEX(Data!E1:G2,1,{1,3}))"),
    ("Legacy", "B9", Dynamic, "=SUM(INDEX(Data!E1:G2,1,{1,3}))"),
    (
        "Legacy",
        "C9",
        Block("C9"),
        "=SUM(INDEX(Data!E1:G2,1,{1,3}))",
    ),
    (
        "Legacy",
        "A10",
        Plain,
        "=SUM(CHOOSE({1,2},Data!A1,Data!A2))",
    ),
    (
        "Legacy",
        "B10",
        Dynamic,
        "=SUM(CHOOSE({1,2},Data!A1,Data!A2))",
    ),
    (
        "Legacy",
        "C10",
        Block("C10"),
        "=SUM(CHOOSE({1,2},Data!A1,Data!A2))",
    ),
    (
        "Legacy",
        "A11",
        Plain,
        "=SUM(LOOKUP({1,4},Data!M1:M5,Data!N1:N5))",
    ),
    (
        "Legacy",
        "B11",
        Dynamic,
        "=SUM(LOOKUP({1,4},Data!M1:M5,Data!N1:N5))",
    ),
    (
        "Legacy",
        "C11",
        Block("C11"),
        "=SUM(LOOKUP({1,4},Data!M1:M5,Data!N1:N5))",
    ),
];

/// The Typed sheet's texts, each at a tenth row from row 1 (1, 11, 21, ...),
/// which the owner writes into column B through `Range.Formula2`, as Excel
/// 365 takes a person's typing; the value each is predicted, and in the
/// comment the kind the writer's rule as `typed_saved` witnesses it would give,
/// where its text alone decides. Whether Excel's writer decides by the text or
/// by the answer at entry is what the rows marked "entry" tell.
pub(crate) const TYPED: &[&str] = &[
    // big; plain by the text, IFS answering one value here.
    "=IFS(Data!A1>1,\"big\",TRUE,\"small\")",
    // three; plain.
    "=SWITCH(Data!A1,3,\"three\",\"other\")",
    // 1; plain.
    "=CHOOSE(2,Data!A1,Data!A2)",
    // 4; plain (entry: an index from a cell).
    "=CHOOSE(Data!X1,Data!A1,Data!A2,Data!A3)",
    // 30; plain (a cell of one column).
    "=XLOOKUP(4,Data!A1:A6,Data!C1:C6)",
    // 4, ccc, 30 spilled right; dynamic (a row of a block).
    "=XLOOKUP(4,Data!A1:A6,Data!A1:C6)",
    // 30; plain.
    "=LOOKUP(4,Data!M1:M5,Data!N1:N5)",
    // 4; plain (entry: one value, the row from a cell).
    "=INDEX(Data!A1:A6,Data!X1)",
    // 4, 5, 6 spilled right; dynamic (entry: a row, the column from a cell
    // holding 0).
    "=INDEX(Data!E1:G2,2,Data!X2)",
    // 6; plain.
    "=INDEX(Data!E1:G2,2,3)",
    // 1; plain (OFFSET answers a reference).
    "=OFFSET(Data!A1,1,0)",
    // 3; plain.
    "=INDIRECT(\"Data!A1\")",
    // 0; dynamic by the text, a range in a branch (entry: one value).
    "=IF(Data!A1>100,Data!A1:A3,0)",
    // 3, 1, 4 spilled down; dynamic.
    "=IF(Data!A1>0,Data!A1:A3,0)",
    // big, small, big spilled down; dynamic.
    "=IFS(Data!A1:A3>1,\"big\",TRUE,\"small\")",
    // 10; plain (SEQUENCE handed whole to SUM).
    "=SUM(SEQUENCE(4))",
    // 6; plain.
    "=SUMPRODUCT(SEQUENCE(3))",
    // 6; plain.
    "=LET(x,2,x*3)",
    // 8; plain (the range handed whole to SUM).
    "=LET(x,Data!A1:A3,SUM(x))",
    // 8; dynamic (IFERROR's value forced to one in a legacy formula).
    "=SUMPRODUCT(IFERROR(Data!A1:A3,0))",
    // 2; dynamic (a range written with :, of one cell).
    "=ROW(Data!A2:A2)",
    // 23; dynamic.
    "=SUM(Data!A:A*1)",
    // 21.
    "=SUM(Spills!A1#)",
    // 4, after the edit.
    "=ROWS(Spills!D1#)",
    // 21; plain.
    "=SUM(Spills!A1:B3)",
    // 3; dynamic by the text (an array function), one value at entry.
    "=TRANSPOSE(Data!A1)",
    // 1; dynamic by the text, one value at entry.
    "=SEQUENCE(1)",
    // 3; dynamic by the text, one value at entry.
    "=UNIQUE(Data!A1)",
    // 980; plain (an array parameter).
    "=SUMPRODUCT(Data!A1:A6*Data!C1:C6)",
    // 980; dynamic, the brief's own case.
    "=SUM(Data!A1:A6*Data!C1:C6)",
    // 9; plain, a control.
    "=MAX(Data!A1:A6)",
    // 1; plain, a control.
    "=COUNTIF(Data!A1:A6,Data!A1)",
    // #VALUE!, row 321 outside A1:A6; plain without the @, as typed_saved
    // witnessed.
    "=@Data!A1:A6",
    // #VALUE!; plain without the @, as typed_saved witnessed.
    "=SUM(@Data!A1:A6*2)",
];

/// A formula's text as the file holds it: its newer functions under their
/// prefixes, as the build writes them (`xlfn::as_filed`), and `ANCHORARRAY`
/// under `_xlfn.`, which the writer's list does not hold yet (the item's second
/// slice adds it; a file holding it bare reads in Excel as an add-in's).
fn filed(formula: &str) -> String {
    let body = formula.strip_prefix('=').expect("a formula");
    as_filed(body).replace("ANCHORARRAY(", "_xlfn.ANCHORARRAY(")
}

fn place(addr: &str) -> (u32, u32) {
    let r = parse_a1_range(addr).expect("an address");
    (r.top, r.left)
}

fn sheet(parts: &[Part], name: &str) -> usize {
    parts
        .iter()
        .position(|p| p.name == name)
        .unwrap_or_else(|| panic!("a sheet {name}"))
}

/// The probe's workbook.
fn book() -> Book {
    let styles = Styles::new("11", CALIBRI, "<color theme=\"1\"/>");
    let mut parts: Vec<Part> = SHEETS.into_iter().map(Part::bare).collect();
    for (name, addr, value) in VALUES {
        let s = sheet(&parts, name);
        let (row, col) = place(addr);
        match value.strip_prefix('"') {
            Some(text) => parts[s].text(row, col, text, 0),
            None => parts[s].number(row, col, value, 0),
        }
    }
    for (name, addr, held, formula) in PROBES {
        let s = sheet(&parts, name);
        let (row, col) = place(addr);
        let text = xml::text(&filed(formula));
        match held {
            Plain => parts[s].put(row, col, 0, "", &format!("<f>{text}</f>")),
            Dynamic => parts[s].put(
                row,
                col,
                0,
                " cm=\"1\"",
                &format!("<f t=\"array\" ref=\"{addr}\">{text}</f>"),
            ),
            Block(range) => {
                // Every cell of the block holds a value, as Excel's save of one
                // does; the anchor holds the formula over the block.
                let r = parse_a1_range(range).expect("a block");
                for br in r.top..=r.bottom {
                    for bc in r.left..=r.right {
                        parts[s].put(br, bc, 0, "", "<v>0</v>");
                    }
                }
                parts[s].put(
                    row,
                    col,
                    0,
                    "",
                    &format!("<f t=\"array\" ref=\"{range}\">{text}</f><v>0</v>"),
                );
            }
        }
    }
    let typed = sheet(&parts, "Typed");
    for (k, text) in TYPED.iter().enumerate() {
        parts[typed].text(10 * k as u32 + 1, 1, text, 0);
    }
    // A merged range in a spill's way.
    let spills = sheet(&parts, "Spills");
    parts[spills].merges.push("H3:I3".to_string());
    Book {
        styles,
        parts,
        strings: Vec::new(),
        theme: true,
        names: Vec::new(),
    }
}

/// The probe's bytes.
fn bytes() -> Vec<u8> {
    book().bytes_dynamic()
}

#[test]
fn the_spills_probe_is_its_builder_s() {
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
fn the_probe_reads_back_with_every_formula_where_its_builder_put_it() {
    // Read back by the door as Excel's save will be read: each probe a
    // formula at its cell, its text as written, the blocks' other cells and
    // the typed texts values; 104 formulas and 34 texts, so that a probe
    // added or lost is seen.
    let text = crate::reflect::calc::calc_text(&bytes(), FIXTURE, &crate::excel::library())
        .unwrap_or_else(|r| panic!("{}", r.text));
    let rows: Vec<&str> = text.lines().filter(|l| l.starts_with("(calc ")).collect();
    assert_eq!(rows.len(), PROBES.len(), "{text}");
    assert_eq!(PROBES.len(), 104);
    assert_eq!(TYPED.len(), 34);
    for (name, addr, _, formula) in PROBES {
        let at = format!("(calc \"{name}\" \"{addr}\" ");
        let row = rows
            .iter()
            .find(|r| r.starts_with(&at))
            .unwrap_or_else(|| panic!("no row at {name}!{addr}: {text}"));
        let spelled = format!("\"{}\"", formula.replace('"', "\\\""));
        assert!(row.contains(&spelled), "{row} for {formula}");
    }
    let relations =
        crate::reflect::reflect_text(&bytes(), FIXTURE).unwrap_or_else(|r| panic!("{}", r.text));
    for (k, typed) in TYPED.iter().enumerate() {
        let cell = format!(
            "(cell \"Typed\" \"A{}\" \"{}\")",
            10 * k + 1,
            typed.replace('"', "\\\"")
        );
        assert!(relations.contains(&cell), "{cell}");
    }
    let looks = crate::reflect::reflect_looks_text(&bytes(), FIXTURE)
        .unwrap_or_else(|r| panic!("{}", r.text));
    assert!(looks.contains("(merge \"Spills\" \"H3:I3\")"), "{looks}");
}
