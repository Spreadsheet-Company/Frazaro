//! A formula read by a formula in an engine's machine (KERNEL.53), through
//! the public API an outside user has: a cell's formula changes only by a
//! write, and the frame after a write that changes one computes `ISFORMULA`
//! and `FORMULATEXT` again where no value changed, a formula rewritten to
//! the value it had and a value written over a formula that gave it; a
//! twin holds values, so its cells hold no formula; and a write of a cell
//! reading its own formula is refused as a cycle, as a `ROW` of its own cell
//! is (the owner's decision of 2026-10-21).

use vla_lang::calc::library;
use vla_lang::machine::{Machine, Viewed};

const GRID: &str = "(sheet \"S\" visible)\n\
(cell \"S\" \"A1\" 2)\n\
(formula \"S\" \"A2\" \"=1+1\")\n\
(formula \"S\" \"B1\" \"=FORMULATEXT(A2)\")\n\
(formula \"S\" \"B2\" \"=ISFORMULA(A2)\")\n\
(formula \"S\" \"B3\" \"=ISFORMULA(A1)\")\n\
(formula \"S\" \"C1:C3\" \"=A1*10\")\n\
(formula \"S\" \"B4\" \"=FORMULATEXT(C3)\")\n\
(formula \"S\" \"B5\" \"=ISFORMULA(S.last!A2)\")\n\
(formula \"S\" \"B6\" \"=FORMULATEXT(S.last!A2)\")\n\
(formula \"S\" \"B7\" \"=SUMPRODUCT(--ISFORMULA(A1:C3))\")\n";

fn stepped(m: &mut Machine) {
    assert!(m.step(0).done);
}

/// The value the machine's record shows for a cell of `S`.
fn value(m: &Machine, addr: &str) -> String {
    let record = match m.view("values", "S", None) {
        Ok(Viewed::Text(t)) => t,
        other => panic!("{other:?}"),
    };
    let head = format!("(value \"S\" \"{addr}\" ");
    record
        .lines()
        .find_map(|l| l.strip_prefix(&head)?.strip_suffix(')'))
        .unwrap_or_else(|| panic!("no value for {addr} in\n{record}"))
        .to_string()
}

#[test]
fn a_formula_s_text_and_kind_are_read_at_every_frame_after_a_write_changes_them() {
    let mut m = Machine::load(GRID, library::language())
        .unwrap_or_else(|e| panic!("line {}: {}", e.line, e.refusal.text));
    stepped(&mut m);
    assert_eq!(value(&m, "B1"), "\"=1+1\"");
    assert_eq!(value(&m, "B2"), "true");
    assert_eq!(value(&m, "B3"), "false");
    assert_eq!(
        value(&m, "B4"),
        "\"=A3*10\"",
        "a fill's cell, its master's text moved"
    );
    assert_eq!(value(&m, "B5"), "false", "a twin holds values");
    assert_eq!(value(&m, "B6"), "(error \"#N/A\")");
    assert_eq!(value(&m, "B7"), "7", "A2, B1:B3 and the fill");

    // The same value, another formula: no value changed, and the text did.
    m.write("(formula \"S\" \"A2\" \"=4-2\")").unwrap();
    stepped(&mut m);
    assert_eq!(value(&m, "A2"), "2");
    assert_eq!(value(&m, "B1"), "\"=4-2\"");

    // A value written over the formula that gave it.
    m.write("(cell \"S\" \"A2\" 2)").unwrap();
    stepped(&mut m);
    assert_eq!(value(&m, "B1"), "(error \"#N/A\")");
    assert_eq!(value(&m, "B2"), "false");
    assert_eq!(value(&m, "B7"), "6");

    // A formula written over a value of the same value.
    m.write("(formula \"S\" \"A1\" \"=2\")").unwrap();
    stepped(&mut m);
    assert_eq!(value(&m, "B3"), "true");
    assert_eq!(value(&m, "B7"), "7");
    assert_eq!(
        value(&m, "C1"),
        "20",
        "the formula written computes, its value unchanged"
    );

    // A fill's cell written over: the cell beside it reads the fill's text.
    m.write("(formula \"S\" \"C2\" \"=A1*10\")").unwrap();
    stepped(&mut m);
    assert_eq!(value(&m, "B4"), "\"=A3*10\"");
}

#[test]
fn a_write_of_a_cell_reading_its_own_formula_is_refused_as_a_cycle() {
    let mut m = Machine::load(GRID, library::language()).unwrap();
    stepped(&mut m);
    for (addr, text) in [
        ("D1", "=ISFORMULA(D1)"),
        ("D2", "=FORMULATEXT(D2)"),
        ("D3", "=ROW(D3)"),
    ] {
        let refused = m
            .write(&format!("(formula \"S\" \"{addr}\" \"{text}\")"))
            .expect_err(text);
        assert_eq!(refused.refusal.id, "calc-cycle", "{text}");
    }
}
