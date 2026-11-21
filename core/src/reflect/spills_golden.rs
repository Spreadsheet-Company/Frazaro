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
//! The second save, after the item's last slice, is made the same way over the
//! first's file: every case of the first again, to confirm it, and the cases the
//! slices left to it, each question at two cells or more so that a rule it shows
//! rests on more than one: `ISFORMULA` and `FORMULATEXT` of a spilled cell and of
//! a block's (slice (b)'s and (c)'s *Found*); `INDEX`'s array form and the
//! lookups' index given an array in each kind of formula, and `A1#` of a legacy
//! array formula's anchor (slice (b)'s); two spills whose boxes meet though their
//! answers never do (slice (c)'s); and the texts the writer's rule leans dynamic
//! on where no save typed one, an `@` before a call, a span of sheets and an
//! information function of an array constant with room to spill (slice (d)'s and
//! slice 0's). Never a third save.
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

use vla_lang::calc::{agrees, Calc, Computed, Reason};
use vla_lang::sheet::{cell_ref, parse_a1_range, Content};

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
    // -- The second save's. ISFORMULA and FORMULATEXT of a spilled cell, of
    // two spills, the language's FALSE and #N/A, as of a cell of its own value;
    // of their anchor, TRUE and its text.
    ("Spills", "A44", Plain, "=ISFORMULA(B2)"),
    ("Spills", "A45", Plain, "=ISFORMULA(N2)"),
    ("Spills", "A46", Plain, "=FORMULATEXT(B2)"),
    ("Spills", "A47", Plain, "=FORMULATEXT(N2)"),
    ("Spills", "A48", Plain, "=ISFORMULA(A1)"),
    ("Spills", "A49", Plain, "=FORMULATEXT(A1)"),
    // A1# of a legacy array formula's anchor, of two blocks: labelled.
    ("Spills", "A50", Dynamic, "=SUM(ANCHORARRAY(Blocks!B1))"),
    ("Spills", "A51", Dynamic, "=ROWS(ANCHORARRAY(Blocks!C1))"),
    // Two spills whose boxes meet at AW3, AW1:AW6 and AV3:BA3, though each
    // answers one value, 9: labelled, a box being a bound's; and each read.
    ("Spills", "AW1", Dynamic, "=FILTER(Data!A1:A6,Data!A1:A6>8)"),
    (
        "Spills",
        "AV3",
        Dynamic,
        "=TRANSPOSE(FILTER(Data!A1:A6,Data!A1:A6>8))",
    ),
    ("Spills", "A52", Plain, "=IFERROR(AW1,\"blocked\")"),
    ("Spills", "A53", Plain, "=IFERROR(AV3,\"blocked\")"),
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
    // -- The second save's. ISFORMULA and FORMULATEXT of a block's cell, of two
    // blocks, the language's FALSE and #N/A, where Excel's bar shows the block's
    // formula at every cell; of an anchor, TRUE and its text, which the language
    // spells in braces, {=Data!A1:A6*2}.
    ("Blocks", "A17", Plain, "=ISFORMULA(B2)"),
    ("Blocks", "A18", Plain, "=ISFORMULA(D2)"),
    ("Blocks", "A19", Plain, "=FORMULATEXT(B2)"),
    ("Blocks", "A20", Plain, "=FORMULATEXT(D2)"),
    ("Blocks", "A21", Plain, "=ISFORMULA(B1)"),
    ("Blocks", "A22", Plain, "=FORMULATEXT(B1)"),
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
    // -- The second save's. INDEX's array form given an array, of a row and of
    // two rows: dynamic 40 and 4; legacy and array, Excel's to tell, the
    // language's 40 and 4 (lifted).
    ("Legacy", "A12", Plain, "=SUM(INDEX({10,20,30},{1,3}))"),
    ("Legacy", "B12", Dynamic, "=SUM(INDEX({10,20,30},{1,3}))"),
    (
        "Legacy",
        "C12",
        Block("C12"),
        "=SUM(INDEX({10,20,30},{1,3}))",
    ),
    ("Legacy", "A13", Plain, "=SUM(INDEX({1,2;3,4},{1,2},1))"),
    ("Legacy", "B13", Dynamic, "=SUM(INDEX({1,2;3,4},{1,2},1))"),
    (
        "Legacy",
        "C13",
        Block("C13"),
        "=SUM(INDEX({1,2;3,4},{1,2},1))",
    ),
    // VLOOKUP's column and HLOOKUP's row given an array: dynamic 34 and 7;
    // legacy, the first forced, 4 and 2; array, Excel's to tell, the
    // language's 34 and 7 (lifted).
    (
        "Legacy",
        "A14",
        Plain,
        "=SUM(VLOOKUP(4,Data!M1:N5,{1,2},FALSE))",
    ),
    (
        "Legacy",
        "B14",
        Dynamic,
        "=SUM(VLOOKUP(4,Data!M1:N5,{1,2},FALSE))",
    ),
    (
        "Legacy",
        "C14",
        Block("C14"),
        "=SUM(VLOOKUP(4,Data!M1:N5,{1,2},FALSE))",
    ),
    (
        "Legacy",
        "A15",
        Plain,
        "=SUM(HLOOKUP(2,Data!E1:G2,{1,2},FALSE))",
    ),
    (
        "Legacy",
        "B15",
        Dynamic,
        "=SUM(HLOOKUP(2,Data!E1:G2,{1,2},FALSE))",
    ),
    (
        "Legacy",
        "C15",
        Block("C15"),
        "=SUM(HLOOKUP(2,Data!E1:G2,{1,2},FALSE))",
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
    // -- The second save's: the texts the writer's rule leans dynamic on where
    // no save typed one (slice (d)'s *Found*), two of each kind, the rule
    // storing each dynamic, Excel's writer to tell. SINGLE of a reference
    // INDEX answers, met at row 341, #VALUE!; of an array, its first, 6.
    "=SINGLE(INDEX(Data!E1:G2,0,1))",
    "=SINGLE(Data!A1:A3*2)",
    // TYPE of an array constant and of an array function's answer, 64 each.
    "=TYPE({1,2})",
    "=TYPE(SEQUENCE(2))",
    // INDEX of one column at a constant row, 4 and 0.
    "=INDEX(Data!A1:A6,3)",
    "=INDEX(Data!C1:C6,2)",
    // A LAMBDA called where one value is wanted, 4 and 6.
    "=LAMBDA(x,x+1)(Data!A1)",
    "=LAMBDA(x,y,x*y)(2,3)",
    // An @ before a call, which the language's parse does not read (slice
    // (d)'s *Found*), so that the rule stores none: its text as Excel stores
    // it, and its value, met at rows 421 and 431, #VALUE! each.
    "=@INDEX(Data!E1:G2,0,1)",
    "=@INDEX(Data!A1:A6,0)",
    // A span of sheets, which slice 0's probe left out: 4 and 8, plain.
    "=SUM(Data:Spills!A1)",
    "=SUM(Data:Spills!A1:A2)",
    // An information function of an array constant, with room to spill
    // (slice (d)'s *Found*): TRUE, FALSE, FALSE and FALSE, TRUE spilled right,
    // dynamic.
    "=ISNUMBER({1,\"a\",TRUE})",
    "=ISTEXT({1,\"a\"})",
];

/// A formula's text as the file holds it: its newer functions under their
/// prefixes, as the build writes them (`xlfn::as_filed`), `ANCHORARRAY` among
/// them since the item's second slice (a file holding it bare reads in Excel as
/// an add-in's).
fn filed(formula: &str) -> String {
    let body = formula.strip_prefix('=').expect("a formula");
    as_filed(body)
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
    // formula at its cell, its text as written, each block's other cells its
    // anchor's text (the item's second slice), each cell a spill fills its
    // anchor's (the third), and the typed texts values; 134 formulas, the
    // second save's 30 among them, 42 cells of blocks, 39 of spills as built,
    // `Data!K1` 3, and 48 texts, so that a probe added or lost is seen.
    let text = crate::reflect::calc::calc_text(&bytes(), FIXTURE, &crate::excel::library())
        .unwrap_or_else(|r| panic!("{}", r.text));
    let all: Vec<&str> = text.lines().filter(|l| l.starts_with("(calc ")).collect();
    // A spilled cell's row, its anchor's formula at a cell no probe holds.
    let probed = |r: &&str| {
        PROBES
            .iter()
            .any(|(name, addr, _, _)| r.starts_with(&format!("(calc \"{name}\" \"{addr}\" ")))
    };
    let dynamic = |r: &&str| {
        PROBES.iter().any(|(name, _, held, formula)| {
            matches!(held, Dynamic)
                && r.starts_with(&format!("(calc \"{name}\" "))
                && r.contains(&format!(" \"{}\" ", formula.replace('"', "\\\"")))
        })
    };
    let spilled: Vec<&str> = all
        .iter()
        .filter(|r| !probed(r) && dynamic(r))
        .copied()
        .collect();
    assert_eq!(spilled.len(), 39, "{spilled:#?}");
    let rows: Vec<&str> = all
        .iter()
        .filter(|r| !spilled.contains(r))
        .copied()
        .collect();
    let mut cells: Vec<(&str, String, &str)> = Vec::new();
    for (name, addr, held, formula) in PROBES {
        cells.push((name, addr.to_string(), formula));
        if let Block(range) = held {
            let r = parse_a1_range(range).expect("a block");
            for row in r.top..=r.bottom {
                for col in r.left..=r.right {
                    if (row, col) != (r.top, r.left) {
                        cells.push((name, vla_lang::sheet::cell_ref(row, col), formula));
                    }
                }
            }
        }
    }
    assert_eq!(rows.len(), cells.len(), "{text}");
    assert_eq!(PROBES.len(), 134);
    assert_eq!(cells.len() - PROBES.len(), 42);
    assert_eq!(TYPED.len(), 48);
    for (name, addr, formula) in &cells {
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

/// Excel 365's save of the probe: the owner's second, of 2026-11-19, made over
/// the first's file of 2026-11-17, every case of the first in it.
const SAVED: &str = concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../scripts/recalc/spills_saved.xlsx"
);

#[test]
fn excel_s_save_of_the_probe_agrees_wherever_the_door_computes_it() {
    // The item, held to Excel's second save of its probe, every row of the
    // first as the first read it: every cell of the Blocks sheet, each block's
    // laid, and of the Legacy sheet computed and agreeing, the folds of the
    // first save's two rules and of the second's among them (`INDEX`'s row and
    // column of an array forced in a legacy formula, a block's every cell
    // holding the block's formula); every spill of the Spills sheet and of the
    // Typed sheet computed, each cell it fills agreeing with the value Excel
    // saved there, blocked where Excel blocked it, by a value, a formula, a
    // merged range and the sheet's edges; `A1#` of a spilled cell and of a
    // value #REF! as Excel's, of a blocked anchor #SPILL!, and of a block's
    // anchor its block; no row differing; and two spills whose boxes meet, and
    // what reads them, labelled, as the scoping's stop rule holds an edge Excel
    // settles by its order, whether their answers meet (K1 and J3) or not (AW1
    // and AV3, each 9 in Excel's save).
    let saved = std::fs::read(SAVED).expect("Excel's save of the probe");
    let text = crate::reflect::calc::calc_text(&saved, SAVED, &crate::excel::library())
        .unwrap_or_else(|r| panic!("{}", r.text));
    let rows: Vec<&str> = text.lines().collect();
    assert_eq!(rows.len(), 279, "{text}");
    let on = |sheet: &str| {
        let at = format!("(calc \"{sheet}\" ");
        rows.iter()
            .filter(|r| r.starts_with(&at))
            .copied()
            .collect::<Vec<_>>()
    };
    for row in &rows {
        assert!(!row.ends_with(" differ)"), "{row}");
    }
    for (sheet, count) in [("Blocks", 65), ("Legacy", 42)] {
        let held = on(sheet);
        assert_eq!(held.len(), count, "{sheet}");
        for row in held {
            assert!(row.ends_with(" agree)"), "{row}");
        }
    }
    let row_at = |sheet: &str, addr: &str| {
        let head = format!("(calc \"{sheet}\" \"{addr}\" ");
        rows.iter()
            .find(|r| r.starts_with(&head))
            .copied()
            .unwrap_or_else(|| panic!("{sheet}!{addr}"))
    };
    for addr in ["A22", "A23"] {
        assert!(
            row_at("Spills", addr).ends_with("(error \"#REF!\") (error \"#REF!\") agree)"),
            "{addr}"
        );
    }
    // The second save's: a block's cells hold its formula, and `INDEX`'s row
    // of an array is forced in a legacy formula and a legacy array formula.
    for (addr, tail) in [
        ("A17", " true true agree)"),
        ("A18", " true true agree)"),
        ("A19", " \"{=Data!A1:A6*2}\" \"{=Data!A1:A6*2}\" agree)"),
        ("A20", " \"{=Data!E1:G1*2}\" \"{=Data!E1:G1*2}\" agree)"),
    ] {
        assert!(row_at("Blocks", addr).ends_with(tail), "Blocks!{addr}");
    }
    for (addr, tail) in [
        ("A12", " 10 10 agree)"),
        ("B12", " 40 40 agree)"),
        ("C12", " 10 10 agree)"),
        ("A13", " 1 1 agree)"),
        ("B13", " 4 4 agree)"),
        ("C13", " 1 1 agree)"),
    ] {
        assert!(row_at("Legacy", addr).ends_with(tail), "Legacy!{addr}");
    }
    // The Spills sheet: every row agreeing but the two pairs of spills whose
    // boxes meet, at K3 and at AW3, and the four reading them.
    let meeting = [
        "K1", "K2", "K3", "K4", "J3", "A31", "A32", "AW1", "AV3", "A52", "A53",
    ];
    let spills = on("Spills");
    assert_eq!(spills.len(), 57 + 44 + 12, "{spills:#?}");
    for row in &spills {
        let met = meeting
            .iter()
            .any(|a| row.starts_with(&format!("(calc \"Spills\" \"{a}\" ")));
        if met {
            assert!(
                row.contains("(not-computed \"a spill meeting another\")"),
                "{row}"
            );
        } else {
            assert!(row.ends_with(" agree)"), "{row}");
        }
    }
    let at = |addr: &str| row_at("Spills", addr);
    // A blocked anchor's `A1#`, #SPILL! (A24); the sheet's last row filled
    // exactly and read (C1048575); a spill of a spill read by a range (A33);
    // a block read through its anchor's `A1#`, `#N/A` past its array among it
    // (A50, A51); and a spilled cell's formula, none, and its anchor's (A44 to
    // A49).
    assert!(at("A24").contains("(error \"#SPILL!\") (error \"#SPILL!\") agree"));
    assert!(at("C1048575").ends_with(" 20 20 agree)"));
    assert!(at("A33").ends_with(" 210 210 agree)"));
    assert!(at("A50").ends_with(" (error \"#N/A\") (error \"#N/A\") agree)"));
    assert!(at("A51").ends_with(" 3 3 agree)"));
    for (addr, tail) in [
        ("A44", " false false agree)"),
        ("A45", " false false agree)"),
        ("A46", " (error \"#N/A\") (error \"#N/A\") agree)"),
        ("A47", " (error \"#N/A\") (error \"#N/A\") agree)"),
        ("A48", " true true agree)"),
        ("A49", " \"=SEQUENCE(3,2)\" \"=SEQUENCE(3,2)\" agree)"),
    ] {
        assert!(at(addr).ends_with(tail), "Spills!{addr}");
    }
    for addr in ["F1", "G1", "H1", "A1048575", "XFC1"] {
        assert!(
            at(addr).contains("(error \"#SPILL!\") (error \"#SPILL!\") agree"),
            "{addr}"
        );
    }
    // The Typed sheet: every spill computed, an XLOOKUP's row, an IF's range
    // and two information functions of an array constant among them, and
    // `Spills!A1#` read from another sheet; the six left labelled are the
    // language's own, three functions it does not hold and an IFS of an
    // array, its spill's two cells with it.
    let typed = on("Typed");
    let labelled: Vec<&str> = typed
        .iter()
        .filter(|r| !r.ends_with(" agree)"))
        .copied()
        .collect();
    assert_eq!(labelled.len(), 6, "{labelled:#?}");
    for (addr, tail) in [
        ("D51", " 30 30 agree)"),
        ("B133", " 4 4 agree)"),
        ("B221", " 21 21 agree)"),
        ("D461", " false false agree)"),
        ("C471", " true true agree)"),
    ] {
        assert!(row_at("Typed", addr).ends_with(tail), "Typed!{addr}");
    }
    assert_eq!(rows.iter().filter(|r| r.ends_with(" agree)")).count(), 262);
}

#[test]
fn an_engine_computes_excel_s_save_of_the_probe_as_excel_saved_it() {
    // The item's fifth slice, held to Excel's second save of its probe: the
    // rows the reader writes of the workbook loaded into an engine's machine
    // under Excel's library, each spill and each block Excel saved laid again
    // where the plan bounds it, and its first frame computed; at every
    // formula, every cell a spill fills and every cell of a block where the
    // door computes the value Excel saved, 262, the machine computes it too,
    // its legacy array formulas among them since the fourth part. A formula
    // calling a function the language does not hold, which a machine refuses
    // at its load and the door labels, is the value Excel saved there.
    use vla_lang::calc::Values;
    use vla_lang::machine::Machine;
    let saved = std::fs::read(SAVED).expect("Excel's save of the probe");
    let read =
        crate::reflect::calc::collect(&saved, SAVED).unwrap_or_else(|r| panic!("{}", r.text));
    let lib = crate::excel::library();
    let door = Calc::run(&read.model, &lib);
    let unheld: Vec<(String, String)> = door
        .cells()
        .into_iter()
        .filter(|id| {
            matches!(
                door.computed(*id),
                Some(Computed::NotComputed(Reason::Function(_)))
            )
        })
        .map(|id| (read.model.sheets[id.0].name.clone(), cell_ref(id.1, id.2)))
        .collect();
    assert_eq!(unheld.len(), 3, "{unheld:?}");
    let (looks, _) =
        crate::reflect::reflect_looks(&saved, SAVED).unwrap_or_else(|r| panic!("{}", r.text));
    let mut rows = String::with_capacity(looks.len());
    for line in looks.lines() {
        let of = |kind: &str, sheet: &str, addr: &str| {
            line.starts_with(&format!("({kind} \"{sheet}\" \"{addr}\" "))
        };
        let refers =
            |sheet: &str, addr: &str| line.starts_with(&format!("(refers \"{sheet}!{addr}\" "));
        if unheld
            .iter()
            .any(|(s, a)| of("formula", s, a) || refers(s, a))
        {
            continue;
        }
        if let Some(rest) = line
            .strip_prefix("(value ")
            .filter(|_| unheld.iter().any(|(s, a)| of("value", s, a)))
        {
            rows.push_str(&format!("(cell {rest}\n"));
            continue;
        }
        rows.push_str(line);
        rows.push('\n');
    }
    let mut m = Machine::load(&rows, crate::excel::library())
        .unwrap_or_else(|e| panic!("line {}: {}", e.line, e.refusal.text));
    assert!(m.step(0).done);
    let (mut held, mut wrong) = (0, Vec::new());
    for id in door.cells() {
        let Some(Computed::Value(value)) = door.computed(id) else {
            continue;
        };
        let Some(excel) = read.cached.get(&id) else {
            continue;
        };
        if agrees(value, excel) != Some(true) {
            continue;
        }
        let sheet = read.model.sheets[id.0].name.as_str();
        let s = m.find_sheet(sheet).expect(sheet);
        let ours = m.computed((s, id.1, id.2));
        match &ours {
            Some(Computed::Value(v)) if agrees(v, excel) == Some(true) => held += 1,
            _ => wrong.push(format!(
                "{sheet}!{}: {} where Excel saved {}",
                cell_ref(id.1, id.2),
                ours.map_or("nothing".to_string(), |c| c.spell()),
                vla_lang::rows::datum(excel)
            )),
        }
    }
    assert!(
        wrong.is_empty(),
        "{}",
        wrong.join(
            "
"
        )
    );
    assert_eq!(held, 262);
}

/// Excel 365's save of `scripts/recalc/typed.txt`'s texts, each typed through
/// `Range.Formula2` (KERNEL.70).
const TYPED_SAVED: &str = concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../scripts/recalc/typed_saved.xlsx"
);

/// The texts the language's parse does not read, an `@` before a call (the
/// item's fourth slice's *Found* 1), each with the text Excel stored for it,
/// legacy and without the `@`.
const UNREAD: [(&str, &str); 2] = [
    ("=@INDEX(Data!E1:G2,0,1)", "INDEX(Data!E1:G2,0,1)"),
    ("=@INDEX(Data!A1:A6,0)", "INDEX(Data!A1:A6,0)"),
];

#[test]
fn the_writer_stores_each_typed_text_as_excel_s_writer_did() {
    // The item's fourth slice, decision 3: each of the 104 texts the
    // three saves typed through `Range.Formula2`, the typed fixture's 56 and
    // the probe's 48, stored by the writer's rule as Excel's writer stored it,
    // dynamic or legacy, and as its text, a typed `@` written as Excel writes
    // it; but the two the parse does not read, named.
    let lib = crate::excel::library();
    let (mut seen, mut unread) = (0, 0);
    let mut wrong: Vec<String> = Vec::new();
    for (path, sheet) in [(TYPED_SAVED, "typed"), (SAVED, "Typed")] {
        let bytes = std::fs::read(path).expect("a save");
        let read =
            crate::reflect::calc::collect(&bytes, path).unwrap_or_else(|r| panic!("{}", r.text));
        let s = read.model.find_sheet(sheet).expect(sheet);
        let cells = &read.model.sheets[s].cells;
        for (&(row, col), cell) in cells {
            if col != 2 {
                continue;
            }
            let (excel_dynamic, excel_text) = match &cell.content {
                Content::DynamicFormula(t) | Content::DynamicSpill { text: t, .. } => (true, t),
                Content::Formula(t) => (false, t),
                _ => continue,
            };
            let Some(Content::Text(typed)) = cells.get(&(row, 1)).map(|c| &c.content) else {
                panic!("{sheet}!A{row}: no typed text");
            };
            seen += 1;
            if let Some((_, stored)) = UNREAD.iter().find(|(t, _)| t == typed) {
                assert_eq!(vla_lang::calc::stored::stored(typed, &lib), None, "{typed}");
                assert!(
                    !excel_dynamic && excel_text == stored,
                    "{typed}: {excel_text}"
                );
                unread += 1;
                continue;
            }
            // The text as the file spells it and the reader reads it back.
            let filed = |text: &str| crate::sheet::xlfn::strip_future_prefixes(&as_filed(text));
            match vla_lang::calc::stored::stored(typed, &lib) {
                Some((dynamic, text))
                    if dynamic == excel_dynamic
                        && filed(&text).eq_ignore_ascii_case(excel_text) => {}
                other => wrong.push(format!(
                    "{sheet}!B{row}: {typed} -> {other:?}, Excel {} {excel_text}",
                    if excel_dynamic { "dynamic" } else { "plain" }
                )),
            }
        }
    }
    assert_eq!((seen, unread), (104, 2));
    assert!(
        wrong.is_empty(),
        "{} of 104:\n{}",
        wrong.len(),
        wrong.join("\n")
    );
}
