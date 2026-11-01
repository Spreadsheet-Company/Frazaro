//! A mask read by one call in an engine's machine (KERNEL.71), through the public
//! API an outside user has: the load no longer refuses `BITAND`, `BITOR` and
//! `BITXOR`, a mask's bits follow a write, a fill reads each bit at its own place,
//! a mark toggled over a twin keeps its state from frame to frame, a popcount reads
//! the mask as arrays, and a number out of range is a value and the frame goes on.

use vla_lang::calc::library;
use vla_lang::machine::{Machine, Viewed};

const GRID: &str = "(sheet \"S\" visible)\n\
(cell \"S\" \"A1\" 341)\n\
(cell \"S\" \"A2\" 3)\n\
(cell \"S\" \"A3\" 0)\n\
(formula \"S\" \"B1\" \"=BITAND(A1,2^(A2-1))>0\")\n\
(formula \"S\" \"B2\" \"=BITOR(A1,2^(2-1))\")\n\
(formula \"S\" \"B3\" \"=BITAND(A1,511-2^(A2-1))\")\n\
(formula \"S\" \"B4\" \"=BITXOR(A1,511)\")\n\
(formula \"S\" \"B5\" \"=SUMPRODUCT(--(BITAND(A1,2^{0,1,2,3,4,5,6,7,8})>0))\")\n\
(formula \"S\" \"B6\" \"=IF(BITAND(A1,16)>0,\\\"5\\\",\\\" \\\")\")\n\
(formula \"S\" \"C1\" \"=BITXOR(S.last!C1,A3)\")\n\
(formula \"S\" \"D1:L2\" \"=BITAND($A$1,2^(COLUMN()-4))\")\n";

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
fn a_mask_read_by_one_call_follows_each_write() {
    let mut m = Machine::load(GRID, library::language())
        .unwrap_or_else(|e| panic!("line {}: {}", e.line, e.refusal.text));
    stepped(&mut m);
    assert_eq!(
        value(&m, "B1"),
        "true",
        "the mark 3 of the marks 1, 3, 5, 7 and 9"
    );
    assert_eq!(value(&m, "B2"), "343");
    assert_eq!(value(&m, "B3"), "337");
    assert_eq!(value(&m, "B4"), "170");
    assert_eq!(value(&m, "B5"), "5");
    assert_eq!(value(&m, "B6"), "\"5\"");
    let bits = ["1", "0", "4", "0", "16", "0", "64", "0", "256"];
    let columns = ["D", "E", "F", "G", "H", "I", "J", "K", "L"];
    for (column, bit) in columns.iter().zip(bits) {
        assert_eq!(value(&m, &format!("{column}1")), bit);
        assert_eq!(value(&m, &format!("{column}2")), bit);
    }

    m.write("(cell \"S\" \"A2\" 2)").unwrap();
    stepped(&mut m);
    assert_eq!(value(&m, "B1"), "false");
    assert_eq!(value(&m, "B3"), "341");

    m.write("(cell \"S\" \"A1\" 2)").unwrap();
    stepped(&mut m);
    assert_eq!(value(&m, "B1"), "true");
    assert_eq!(value(&m, "B4"), "509");
    assert_eq!(value(&m, "B5"), "1");
    assert_eq!(value(&m, "B6"), "\" \"");
    assert_eq!(value(&m, "E1"), "2");
    assert_eq!(value(&m, "D2"), "0");
}

#[test]
fn a_mark_toggled_over_its_twin_keeps_its_state_from_frame_to_frame() {
    let mut m = Machine::load(GRID, library::language())
        .unwrap_or_else(|e| panic!("line {}: {}", e.line, e.refusal.text));
    stepped(&mut m);
    assert_eq!(value(&m, "C1"), "0");
    for (act, want) in [(4, "4"), (0, "4"), (16, "20"), (4, "16"), (16, "0")] {
        m.write(&format!("(cell \"S\" \"A3\" {act})")).unwrap();
        stepped(&mut m);
        assert_eq!(value(&m, "C1"), want, "the act {act}");
    }
}

#[test]
fn a_number_out_of_range_is_a_value_and_the_frame_goes_on() {
    let mut m = Machine::load(GRID, library::language())
        .unwrap_or_else(|e| panic!("line {}: {}", e.line, e.refusal.text));
    stepped(&mut m);
    let num = "(error \"#NUM!\")";
    for bad in ["-1", "0.5", "281474976710656"] {
        m.write(&format!("(cell \"S\" \"A1\" {bad})")).unwrap();
        stepped(&mut m);
        assert_eq!(value(&m, "B2"), num, "{bad}");
        assert_eq!(value(&m, "B4"), num, "{bad}");
        assert_eq!(value(&m, "D2"), num, "{bad}");
    }
    m.write("(cell \"S\" \"A1\" \"x\")").unwrap();
    stepped(&mut m);
    assert_eq!(value(&m, "B2"), "(error \"#VALUE!\")");
    assert_eq!(value(&m, "D2"), "(error \"#VALUE!\")");

    // Back in range, every reader computes again.
    m.write("(cell \"S\" \"A1\" 281474976710655)").unwrap();
    stepped(&mut m);
    assert_eq!(value(&m, "B4"), "281474976710144");
    assert_eq!(value(&m, "L2"), "256");
}

#[test]
fn a_machine_refuses_none_of_the_three_and_a_program_calls_each_as_a_cell_alone_does() {
    // One shape in a run of many cells and the same text alone, so that a run's
    // call, a program's and `Ctx`'s are held to one answer.
    let grid = "(sheet \"S\" visible)\n\
(formula \"S\" \"A1:A64\" \"=BITOR(ROW(),BITXOR(ROW(),3))\")\n\
(formula \"S\" \"B1:B64\" \"=BITAND(A1,ROW())\")\n\
(formula \"S\" \"D1\" \"=BITOR(ROW(),BITXOR(ROW(),3))\")\n\
(formula \"S\" \"D2\" \"=BITAND(D1,ROW())\")\n";
    let mut m = Machine::load(grid, library::language())
        .unwrap_or_else(|e| panic!("line {}: {}", e.line, e.refusal.text));
    stepped(&mut m);
    for row in 1u64..=64 {
        let or = row | (row ^ 3);
        assert_eq!(value(&m, &format!("A{row}")), or.to_string(), "row {row}");
        assert_eq!(
            value(&m, &format!("B{row}")),
            (or & row).to_string(),
            "row {row}"
        );
    }
    assert_eq!(value(&m, "D1"), "3");
    assert_eq!(value(&m, "D2"), "2");
}
