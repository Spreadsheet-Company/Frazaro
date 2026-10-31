//! A cell's address as a text in an engine's machine (KERNEL.60), through the
//! public API an outside user has: the load no longer refuses `ADDRESS`, a
//! fill's every cell writes its own place, a hint names a cell by its numbers,
//! a Write sheet row's address follows a state cell, and an address in R1C1 is
//! held labelled at its cell and its readers, as a door labels it.

use vla_lang::calc::library;
use vla_lang::machine::{Machine, Viewed};

const GRID: &str = "(sheet \"S\" visible)\n\
(cell \"S\" \"A1\" 3)\n\
(cell \"S\" \"A2\" 2)\n\
(formula \"S\" \"B1\" \"=ADDRESS(A1,A2,4)\")\n\
(formula \"S\" \"B2\" \"=ADDRESS(A1,A2)\")\n\
(formula \"S\" \"B3\" \"=ADDRESS(A1+1,A2+1,4)&\\\" is 5: the only place left for it in its row.\\\"\")\n\
(formula \"S\" \"B4\" \"=ADDRESS(1,1,4,TRUE,\\\"My Sheet\\\")\")\n\
(formula \"S\" \"B5\" \"=SUMPRODUCT(LEN(ADDRESS(A1:A2,1)))\")\n\
(formula \"S\" \"D1:E3\" \"=ADDRESS(ROW(),COLUMN(),4)\")\n";

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
fn an_address_follows_the_numbers_it_reads_and_a_fill_writes_its_own_places() {
    let mut m = Machine::load(GRID, library::language())
        .unwrap_or_else(|e| panic!("line {}: {}", e.line, e.refusal.text));
    stepped(&mut m);
    assert_eq!(value(&m, "B1"), "\"B3\"");
    assert_eq!(value(&m, "B2"), "\"$B$3\"");
    assert_eq!(
        value(&m, "B3"),
        "\"C4 is 5: the only place left for it in its row.\""
    );
    assert_eq!(value(&m, "B4"), "\"'My Sheet'!A1\"");
    assert_eq!(value(&m, "B5"), "8", "$A$3 and $A$2");
    for (cell, place) in [
        ("D1", "\"D1\""),
        ("E1", "\"E1\""),
        ("D2", "\"D2\""),
        ("E3", "\"E3\""),
    ] {
        assert_eq!(value(&m, cell), place);
    }

    m.write("(cell \"S\" \"A1\" 12)").unwrap();
    stepped(&mut m);
    assert_eq!(value(&m, "B1"), "\"B12\"");
    assert_eq!(value(&m, "B3").get(..4), Some("\"C13"));

    // Past the last row: a value, and the frame goes on.
    m.write("(cell \"S\" \"A1\" 1048577)").unwrap();
    stepped(&mut m);
    assert_eq!(value(&m, "B1"), "(error \"#VALUE!\")");
}

#[test]
fn an_address_in_r1c1_is_held_labelled_at_its_cell_and_its_readers() {
    let grid = "(sheet \"S\" visible)\n\
(formula \"S\" \"A1\" \"=ADDRESS(3,2,1,FALSE)\")\n\
(formula \"S\" \"A2\" \"=LEN(A1)\")\n\
(formula \"S\" \"A3\" \"=ADDRESS(3,2,4)\")\n";
    let mut m = Machine::load(grid, library::language())
        .unwrap_or_else(|e| panic!("line {}: {}", e.line, e.refusal.text));
    stepped(&mut m);
    let label = "(not-computed \"an address in R1C1\")";
    assert_eq!(value(&m, "A1"), label);
    assert_eq!(value(&m, "A2"), label);
    assert_eq!(value(&m, "A3"), "\"B3\"");
}
