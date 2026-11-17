//! The adder bench programs' generator and their hand (`KERNEL.29`, its second
//! slice; `docs/BETA_REARVIEW.md`, the slice's second part): a ripple-carry
//! adder of [`BITS`] bits, a full adder a bit, its sum and its carry each a
//! fill, written out in `adder.txt` and through two named functions, `SUMBIT`
//! and `CARRYBIT`, in `adder-named.txt`. No sentence defines a name, so the
//! named program's `# bench: name` comments define them and the runners add
//! each to the built workbook before its values are folded (`bench/common.rs`,
//! `build_program`). A load writes each call's body into its fill's formula, so
//! the two plan alike but for their texts, which the recalc runner holds, and
//! the floors hold that a function costs what its formula written out costs.
//!
//! Bit `i` stands in row `i + 2`: its switches a and b in columns A and B, all
//! 1 and all 0, so that the carry-in at `D1`, the edit, ripples through every
//! bit; its sum in column C and its carry in column D, each column one fill
//! reading the carry above it. Filled, and not shuffled as the netlist is: the
//! adder at shuffled places, each cell a shape of its own, moved its frames by
//! up to a fifth between processes, by where the heap placed its machine, more
//! than the margin the two are held within, where a fill's frames move as the
//! chain's do (the slice's record). The sum `C2:C4097` reads every bit's sum:
//! 4,096 as built, and 0 with the carry in.
//!
//! The recalc runner writes the two programs (`-- --adder`) and holds the
//! files to them (`-- --check`); and holds every sum and carry the machine
//! computes to the adder evaluated here, bit by bit, a lineage of its own
//! beside the machine and `Calc`.

use super::netlist::a1;

/// The adder's width, the netlist's.
pub const BITS: usize = 4096;

/// The named functions, each text as a workbook holds a defined name's, with
/// no `=`: a bit's sum and its carry over its two switches and the carry in.
pub const NAMES: [(&str, &str); 2] = [
    ("SUMBIT", "LAMBDA(a,b,c,MOD(MOD(a+b,2)+c,2))"),
    ("CARRYBIT", "LAMBDA(a,b,c,MAX(a*b,MOD(a+b,2)*c))"),
];

/// The carry-in's cell, `D1`, above the carries.
const CARRY_IN: (u32, u32) = (1, 4);

/// A program, its lines LF: the comments, the names where it is `named`, the
/// switches and the carry-in, then the sums' fill and the carries'.
pub fn program(named: bool) -> String {
    let last = BITS + 1;
    let mut s = String::new();
    let mut lines = vec![
        "# The recalculation floors' adder fixtures (KERNEL.29, its second slice): a ripple-carry adder of",
        "# 4,096 bits, a full adder a bit, its sum and its carry each a fill, written out in adder.txt and",
        "# through the named functions SUMBIT and CARRYBIT in adder-named.txt, so that the floors hold what a",
        "# function costs against its formula written out. Bit i stands in row i + 2: its switches a,",
        "# A2:A4097, all 1, and b, B2:B4097, all 0; its sum in C and its carry in D, each reading the carry",
        "# above it. The edit is the carry-in D1, from 0 to 1, which ripples through every bit. The sum,",
        "# C2:C4097, is every bit's sum: 4,096 as built and 0 with the carry in.",
    ];
    lines.extend(match named {
        false => &["# This one writes each formula out."][..],
        true => &[
            "# This one calls the two names its name comments define: no sentence defines a name, so the",
            "# runners add each to the built workbook before its values are folded (bench/common.rs).",
        ][..],
    });
    lines.extend([
        "# Written by core/examples/recalc.rs --adder (bench/adder.rs) and held to it by --check: never",
        "# edit it by hand.",
    ]);
    for line in lines {
        s.push_str(line);
        s.push('\n');
    }
    if named {
        for (name, text) in NAMES {
            s.push_str(&format!("# bench: name {name} {text}\n"));
        }
    }
    let cin = a1(CARRY_IN);
    s.push_str(&format!("# bench: edit Output!{cin} 0 1\n"));
    s.push_str(&format!("# bench: sum Output!C2:C{last}\n\n"));
    s.push_str(&format!("Put 1 into range A2:A{last}.\n"));
    s.push_str(&format!("Put 0 into range B2:B{last}.\n"));
    s.push_str(&format!("Put 0 into cell {cin}.\n"));
    let (sum, carry) = match named {
        true => (
            format!("=SUMBIT(A2,B2,{cin})"),
            format!("=CARRYBIT(A2,B2,{cin})"),
        ),
        false => (
            format!("=MOD(MOD(A2+B2,2)+{cin},2)"),
            format!("=MAX(A2*B2,MOD(A2+B2,2)*{cin})"),
        ),
    };
    s.push_str(&format!("Put formula \"{sum}\" into range C2:C{last}.\n"));
    s.push_str(&format!("Put formula \"{carry}\" into range D2:D{last}.\n"));
    s
}

/// Every sum's and carry's cell and value by hand, the switches a all 1 and b
/// all 0 and the carry-in `cin`: the adder evaluated bit by bit.
pub fn by_hand(cin: u8) -> Vec<((u32, u32), u8)> {
    let (a, b) = (1u8, 0u8);
    let mut c = cin;
    let mut out = Vec::with_capacity(2 * BITS);
    for i in 0..BITS {
        let r = i as u32 + 2;
        let total = a + b + c;
        c = total / 2;
        out.push(((r, 3), total % 2));
        out.push(((r, 4), c));
    }
    out
}
