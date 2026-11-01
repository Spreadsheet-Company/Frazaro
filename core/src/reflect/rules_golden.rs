//! Excel's golden of a conditional format (`KERNEL.58`, the third slice),
//! test-only.
//!
//! Three workbooks, a case a row, written here and held byte for byte
//! (`FRAZARO_BLESS=1` writes them again), hold what a rule may say, so that
//! Excel's reading of each is the oracle of the language's rules:
//!
//! - `scripts/reflect/rules.xlsx`: what a rule's answer takes for true, each
//!   of `cellIs`' operators over each kind of value, how a reference moves
//!   from a rule's range, the order of several rules, a format that says a
//!   part of a look over a cell's own, what a rule's format says and Excel
//!   does not draw, the kinds with a formula and those without, and a rule
//!   reading another sheet through a defined name.
//! - `scripts/reflect/rules_plain.xlsx`: a rule reading another sheet, visible
//!   and hidden, in the sheet's own part, apart since a repair would taint the
//!   rest.
//! - `scripts/reflect/rules_x14.xlsx`: the same in the sheet's extension, as
//!   Excel's save of such a rule is remembered to write.
//!
//! A case's name is in column A and its reads in column B: the targets
//! `VlaWriteRulesGolden` (VLA_Tests.bas) reads through Excel's object model
//! into `scripts/reflect/rules_golden.txt`, case by case, before it saves each
//! workbook as Excel's own, `<name>_saved.xlsx` and `<name>_saved.ods`. A
//! target is `shown:<A1>`, the look a cell shows, its rules applied, by
//! `Range.DisplayFormat`; `text:<A1>`, the text it shows, by `Range.Text`; and
//! `rule:<A1>`, each rule over the cell, by `Range.FormatConditions`. A case's
//! cells read no other case's.
//!
//! The builder keeps each rule's row beside its element, as a `Part`'s
//! `rule_rows`, and each format's words beside its XML ([`format_words`]), what
//! a reader of the file is to say of them. A test loads the rows the reader
//! prints of the cells, and those, into the language's machine under Excel's
//! library, and holds what each cell draws to Excel's reading, case by case,
//! each difference named in [`GAPS`].

use std::collections::{BTreeSet, HashMap};

use vla_lang::form::Form;
use vla_lang::machine::{Machine, Viewed};
use vla_lang::rows::{quoted, record_line, Record};
use vla_lang::sheet::{parse_a1_range, Style, Workbook};

use super::looks_golden::{drawn, place, solid, Book, Drawn, Part, Styles, Xf, CALIBRI};
use crate::sheet::{cell_ref, column_letters};

const RULES: &str = concat!(env!("CARGO_MANIFEST_DIR"), "/../scripts/reflect/rules.xlsx");
const PLAIN: &str = concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../scripts/reflect/rules_plain.xlsx"
);
const X14: &str = concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../scripts/reflect/rules_x14.xlsx"
);
const GOLDEN: &str = concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../scripts/reflect/rules_golden.txt"
);

/// The fill most cases light a cell with, and the second where two rules
/// must be told apart.
const LIT: &str = "<fill><patternFill><bgColor rgb=\"FFFFE699\"/></patternFill></fill>";
const OTHER: &str = "<fill><patternFill><bgColor rgb=\"FF9DC3E6\"/></patternFill></fill>";

/// A cell a rule lights and reads nothing of.
const PLAIN_TEXT: &str = "Aa 123";

/// A value a case's cell holds.
#[derive(Clone, Copy)]
enum Held {
    Number(&'static str),
    Text(&'static str),
    Truth(bool),
    Error(&'static str),
    Empty,
    /// A text of spaces alone, which a file must say to keep.
    Spaces,
    /// The empty text, which only a formula holds.
    EmptyText,
}

fn hold(p: &mut Part, row: u32, col: u32, value: Held, s: u32) {
    match value {
        Held::Number(n) => p.number(row, col, n, s),
        Held::Text(t) => p.text(row, col, t, s),
        Held::Truth(b) => p.put(row, col, s, " t=\"b\"", &format!("<v>{}</v>", u8::from(b))),
        Held::Error(e) => p.put(row, col, s, " t=\"e\"", &format!("<v>{e}</v>")),
        Held::Empty => {
            if s != 0 {
                p.empty(row, col, s);
            }
        }
        Held::Spaces => p.put(
            row,
            col,
            s,
            " t=\"inlineStr\"",
            "<is><t xml:space=\"preserve\">   </t></is>",
        ),
        Held::EmptyText => p.put(row, col, s, " t=\"str\"", "<f>\"\"</f><v></v>"),
    }
}

/// A sheet of rules being written: its part, the workbook's styles, and the
/// priority its next rule takes, each once on its sheet as Excel writes them.
struct Sheet<'s> {
    part: Part,
    styles: &'s mut Styles,
    last: u32,
    /// The rules made and not yet placed, each element with its row's parts.
    made: Vec<(String, Made)>,
    /// How many rows each priority has taken on the sheet, a tie's among them.
    taken: HashMap<u32, u32>,
}

/// A rule's row's parts, but its range.
struct Made {
    priority: u32,
    test: Option<String>,
    format: String,
    stop: bool,
}

/// A rule's test as its row spells it; `None` for a kind the language does
/// not compute, and for `timePeriod`, whose `TODAY` a load refuses as it does
/// in a cell.
fn row_test(kind: &str, extra: &str, formulas: &[&str]) -> Option<String> {
    let spelled = |f: &str| quoted(&format!("={f}"));
    match kind {
        "cellIs" => {
            let word = extra.split('"').nth(1)?;
            let word = match word {
                "lessThan" => "less",
                "lessThanOrEqual" => "less-or-equal",
                "equal" => "equal",
                "notEqual" => "not-equal",
                "greaterThanOrEqual" => "greater-or-equal",
                "greaterThan" => "greater",
                "between" => "between",
                "notBetween" => "not-between",
                other => panic!("an operator of no word: {other}"),
            };
            let bounds: Vec<String> = formulas.iter().map(|f| spelled(f)).collect();
            Some(format!("(cell {word} {})", bounds.join(" ")))
        }
        "timePeriod" => None,
        // Excel lights an error as not blank, where this kind's own formula
        // answers the error: a reader says so, and every other value as the
        // formula does (its `is not blank` and `is blank` cases).
        "notContainsBlanks" => formulas
            .first()
            .map(|f| spelled(&format!("IFERROR({f},TRUE)"))),
        _ => formulas.first().map(|f| spelled(f)),
    }
}

/// Each differential format's XML and what a reader is to say of it: the
/// words of a `format` row, or none where the language draws none of what it
/// says, each of those named in [`GAPS`] where Excel's reading shows it.
fn format_words() -> Vec<(String, &'static str)> {
    let edge = |name: &str, color: &str, body: &str| {
        format!(
            "<border>{}</border>",
            body.replace(
                "{e}",
                &format!("<{name} style=\"thin\"><color rgb=\"{color}\"/></{name}>")
            )
        )
    };
    vec![
        (LIT.to_string(), "(fill \"FFE699\")"),
        (OTHER.to_string(), "(fill \"9DC3E6\")"),
        (BOLD.to_string(), "(font bold)"),
        (ITALIC.to_string(), "(font italic)"),
        // A fill.
        ("<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FF800020\"/></patternFill></fill>".to_string(), ""),
        ("<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FF800020\"/><bgColor rgb=\"FFFFE699\"/></patternFill></fill>".to_string(), "(fill \"FFE699\")"),
        ("<fill><patternFill patternType=\"none\"/></fill>".to_string(), "(fill none)"),
        ("<fill><patternFill patternType=\"none\"><bgColor auto=\"1\"/></patternFill></fill>".to_string(), "(fill none)"),
        ("<fill><patternFill patternType=\"gray125\"><fgColor rgb=\"FF800020\"/><bgColor rgb=\"FFFFE699\"/></patternFill></fill>".to_string(), ""),
        ("<fill><patternFill><bgColor theme=\"4\" tint=\"0.39997558519241921\"/></patternFill></fill>".to_string(), "(fill \"8EA9DB\")"),
        ("<fill><patternFill><bgColor indexed=\"10\"/></patternFill></fill>".to_string(), "(fill \"FF0000\")"),
        // A font.
        ("<font><b val=\"0\"/></font>".to_string(), "(font not-bold)"),
        ("<font><i val=\"0\"/></font>".to_string(), "(font not-italic)"),
        ("<font><color rgb=\"FF800020\"/></font>".to_string(), "(font \"800020\")"),
        ("<font><color rgb=\"FF0000FF\"/></font>".to_string(), "(font \"0000FF\")"),
        ("<font><color rgb=\"FF800020\"/></font><fill><patternFill><bgColor rgb=\"FFFFE699\"/></patternFill></fill>".to_string(), "(font \"800020\") (fill \"FFE699\")"),
        ("<font><color theme=\"4\"/></font>".to_string(), "(font \"4472C4\")"),
        ("<font><color auto=\"1\"/></font>".to_string(), ""),
        ("<font><u/></font>".to_string(), "(font underline)"),
        ("<font><u val=\"double\"/></font>".to_string(), "(font double-underline)"),
        ("<font><u val=\"singleAccounting\"/></font>".to_string(), "(font accounting-underline)"),
        ("<font><u val=\"doubleAccounting\"/></font>".to_string(), "(font double-accounting-underline)"),
        ("<font><u val=\"none\"/></font>".to_string(), "(font not-underline)"),
        ("<font><strike/></font>".to_string(), "(font strike)"),
        ("<font><strike val=\"0\"/></font>".to_string(), "(font not-strike)"),
        // A number format, by the text a cell shows.
        ("<numFmt numFmtId=\"200\" formatCode=\"0.000\"/>".to_string(), "(number \"0.000\")"),
        ("<numFmt numFmtId=\"2\" formatCode=\"0.00\"/>".to_string(), "(number \"0.00\")"),
        ("<numFmt numFmtId=\"0\" formatCode=\"General\"/>".to_string(), "(number general)"),
        ("<numFmt numFmtId=\"49\" formatCode=\"@\"/>".to_string(), "(number text)"),
        ("<numFmt numFmtId=\"201\" formatCode=\"0.00&quot; kg&quot;\"/>".to_string(), "(number \"0.00\\\" kg\\\"\")"),
        ("<numFmt numFmtId=\"202\" formatCode=\"[Red]0.00\"/>".to_string(), ""),
        // What a format holds and no dialog of Excel's writes.
        ("<font><sz val=\"20\"/></font>".to_string(), "(font 20)"),
        ("<font><name val=\"Consolas\"/></font>".to_string(), ""),
        ("<font><vertAlign val=\"superscript\"/></font>".to_string(), ""),
        ("<alignment horizontal=\"center\"/>".to_string(), "(align center)"),
        ("<alignment vertical=\"top\"/>".to_string(), ""),
        ("<alignment wrapText=\"1\"/>".to_string(), "(wrap wrap)"),
        ("<font><sz val=\"20\"/></font><fill><patternFill><bgColor rgb=\"FFFFE699\"/></patternFill></fill>".to_string(), "(font 20) (fill \"FFE699\")"),
        // An edge, an empty side saying no line and an absent one nothing.
        (edge("left", "FF800020", "{e}<right/><top/><bottom/>"), "(border (left thin \"800020\") (right none) (top none) (bottom none))"),
        (edge("left", "FF800020", "{e}"), "(border (left thin \"800020\"))"),
        (edge("right", "FF0000FF", "{e}"), "(border (right thin \"0000FF\"))"),
        ("<border><left style=\"thin\"><color rgb=\"FF0000FF\"/></left><right style=\"thin\"><color rgb=\"FF0000FF\"/></right><top/><bottom/></border>".to_string(), "(border (left thin \"0000FF\") (right thin \"0000FF\") (top none) (bottom none))"),
        (edge("right", "FF800020", "<left/>{e}<top/><bottom/>"), "(border (left none) (right thin \"800020\") (top none) (bottom none))"),
        (edge("top", "FF800020", "<left/><right/>{e}<bottom/>"), "(border (left none) (right none) (top thin \"800020\") (bottom none))"),
        (edge("bottom", "FF800020", "<left/><right/><top/>{e}"), "(border (left none) (right none) (top none) (bottom thin \"800020\"))"),
        (left("medium", "rgb=\"FF800020\""), "(border (left medium \"800020\") (right none) (top none) (bottom none))"),
        (left("thick", "rgb=\"FF800020\""), "(border (left thick \"800020\") (right none) (top none) (bottom none))"),
        (left("double", "rgb=\"FF800020\""), "(border (left double \"800020\") (right none) (top none) (bottom none))"),
        (left("dashed", "rgb=\"FF800020\""), "(border (left dashed \"800020\") (right none) (top none) (bottom none))"),
        (left("thin", "auto=\"1\""), "(border (left thin) (right none) (top none) (bottom none))"),
        ("<border><left/><right/><top/><bottom/></border>".to_string(), "(border (left none) (right none) (top none) (bottom none))"),
        ("<border><left style=\"none\"/></border>".to_string(), "(border (left none))"),
    ]
}

/// The number a format's XML is loaded as, past the reader's own: a rule with
/// nothing the language draws loads with no format.
fn format_number(xml: &str) -> String {
    let words = format_words();
    let at = words
        .iter()
        .position(|(x, _)| x == xml)
        .unwrap_or_else(|| panic!("a format of no words: {xml}"));
    match words[at].1 {
        "" => "none".to_string(),
        _ => (1000 + at).to_string(),
    }
}

impl<'s> Sheet<'s> {
    fn new(name: &'static str, styles: &'s mut Styles) -> Sheet<'s> {
        Sheet {
            part: Part::new(name),
            styles,
            last: 0,
            made: Vec::new(),
            taken: HashMap::new(),
        }
    }

    fn priority(&mut self) -> u32 {
        self.last += 1;
        self.last
    }

    /// A differential format's index, written once.
    fn format(&mut self, body: &str) -> usize {
        place(&mut self.styles.dxfs, format!("<dxf>{body}</dxf>"))
    }

    /// A rule's element at a priority: its kind, its format, the attributes
    /// after its priority and its formulas.
    fn rule_at(
        &mut self,
        priority: u32,
        kind: &str,
        format: Option<&str>,
        extra: &str,
        formulas: &[&str],
    ) -> String {
        let dxf = format.map_or(String::new(), |f| format!(" dxfId=\"{}\"", self.format(f)));
        let open = format!("<cfRule type=\"{kind}\"{dxf} priority=\"{priority}\"{extra}");
        let bodies: String = formulas
            .iter()
            .map(|f| format!("<formula>{}</formula>", crate::sheet::xml::text(f)))
            .collect();
        let element = match formulas.is_empty() {
            true => format!("{open}/>"),
            false => format!("{open}>{bodies}</cfRule>"),
        };
        self.made.push((
            element.clone(),
            Made {
                priority,
                test: row_test(kind, extra, formulas),
                format: format.map_or("none".to_string(), format_number),
                stop: extra.contains("stopIfTrue"),
            },
        ));
        element
    }

    /// A rule's element at the sheet's next priority.
    fn rule(&mut self, kind: &str, format: Option<&str>, extra: &str, formulas: &[&str]) -> String {
        let priority = self.priority();
        self.rule_at(priority, kind, format, extra, formulas)
    }

    /// An `expression` at the sheet's next priority.
    fn test(&mut self, format: Option<&str>, extra: &str, test: &str) -> String {
        self.rule("expression", format, extra, &[test])
    }

    /// One `conditionalFormatting` element, and its rules' rows. A file may
    /// tie a priority, as no row may: in one element the later rule leads, and
    /// of two elements of one range the first, as Excel read them, so that a
    /// row's priority is the file's a hundred times over, and a tie's its place
    /// among them.
    fn block(&mut self, sqref: &str, rules: &[String]) {
        self.part.conditional.push(format!(
            "<conditionalFormatting sqref=\"{sqref}\">{}</conditionalFormatting>",
            rules.concat()
        ));
        // A kind with no formula and no format of the language's is made here
        // by its element alone, and has no row.
        let made: Vec<Made> = rules
            .iter()
            .filter_map(|rule| {
                let at = self.made.iter().rposition(|(element, _)| element == rule)?;
                Some(self.made.remove(at).1)
            })
            .collect();
        for (k, rule) in made.iter().enumerate() {
            let after = made[k + 1..]
                .iter()
                .filter(|m| m.priority == rule.priority)
                .count();
            let base = self.taken.get(&rule.priority).copied().unwrap_or(0);
            let priority = 100 * rule.priority + base + after as u32;
            if let Some(test) = &rule.test {
                self.part.rule_rows.push(format!(
                    "(conditional {} {priority} {} {test} {}{})\n",
                    quoted(self.part.name),
                    quoted(sqref),
                    rule.format,
                    if rule.stop { " stop" } else { "" }
                ));
            }
        }
        for rule in &made {
            *self.taken.entry(rule.priority).or_insert(0) += 1;
        }
    }

    /// One rule of an `expression` over a range.
    fn over(&mut self, sqref: &str, format: &str, test: &str) {
        let rule = self.test(Some(format), "", test);
        self.block(sqref, &[rule]);
    }
}

/// `shown:` and `rule:` of the cells named, each of the first and only the
/// `shown:` of the rest, where every cell's rule reads alike.
fn reads_of(cells: &[String]) -> String {
    let mut reads: Vec<String> = cells.iter().map(|c| format!("shown:{c}")).collect();
    if let Some(first) = cells.first() {
        reads.push(format!("rule:{first}"));
    }
    reads.join(" ")
}

/// What a rule's answer takes for true (the probes' A): each case's rule over
/// its `C`, a literal or a reference to the `E` beside it.
fn truth(styles: &mut Styles) -> Part {
    let mut s = Sheet::new("Truth", styles);
    for (name, test) in [
        ("a test of the number 1", "1"),
        ("a test of the number 2.5", "2.5"),
        ("a test of the number -1", "-1"),
        ("a test of the number 0", "0"),
        ("a test of TRUE", "TRUE"),
        ("a test of FALSE", "FALSE"),
        ("a test of a text", "\"abc\""),
        ("a test of the text TRUE", "\"TRUE\""),
        ("a test of the text FALSE", "\"FALSE\""),
        ("a test of the text 1", "\"1\""),
        ("a test of the empty text", "\"\""),
        ("a test of a division by zero", "1/0"),
        ("a test of NA", "NA()"),
    ] {
        let row = s.part.next_row();
        s.part.case(name, &reads_of(&[format!("C{row}")]));
        s.part.text(row, 3, PLAIN_TEXT, 0);
        s.over(&format!("C{row}"), LIT, test);
    }
    for (name, value) in [
        ("a reference to the number 1", Held::Number("1")),
        ("a reference to the number 2.5", Held::Number("2.5")),
        ("a reference to the number 0", Held::Number("0")),
        ("a reference to TRUE", Held::Truth(true)),
        ("a reference to FALSE", Held::Truth(false)),
        ("a reference to a text", Held::Text("abc")),
        ("a reference to the text TRUE", Held::Text("TRUE")),
        ("a reference to the text FALSE", Held::Text("FALSE")),
        ("a reference to the text 1", Held::Text("1")),
        ("a reference to an empty cell", Held::Empty),
        ("a reference to a division by zero", Held::Error("#DIV/0!")),
        ("a reference to NA", Held::Error("#N/A")),
    ] {
        let row = s.part.next_row();
        s.part.case(name, &reads_of(&[format!("C{row}")]));
        s.part.text(row, 3, PLAIN_TEXT, 0);
        hold(&mut s.part, row, 5, value, 0);
        s.over(&format!("C{row}"), LIT, &format!("E{row}"));
    }
    for (name, value) in [
        ("a comparison answering true", Held::Number("1")),
        ("a comparison answering false", Held::Number("2")),
        ("a comparison of a text", Held::Text("abc")),
        ("a comparison of an error", Held::Error("#N/A")),
    ] {
        let row = s.part.next_row();
        s.part.case(name, &reads_of(&[format!("C{row}")]));
        s.part.text(row, 3, PLAIN_TEXT, 0);
        hold(&mut s.part, row, 5, value, 0);
        s.over(&format!("C{row}"), LIT, &format!("E{row}=1"));
    }
    for (name, first, last) in [
        ("a range whose first answer is true", "1", "0"),
        ("a range whose first answer is false", "0", "1"),
    ] {
        let row = s.part.next_row();
        s.part.case(name, &reads_of(&[format!("C{row}")]));
        s.part.text(row, 3, PLAIN_TEXT, 0);
        s.part.number(row, 5, first, 0);
        s.part.number(row, 6, last, 0);
        s.over(&format!("C{row}"), LIT, &format!("E{row}:F{row}"));
    }
    s.part
}

/// The values each of `cellIs`' cases is read over, a column apart: below, at
/// and between the bounds and above them, a text, a number's text, an empty
/// cell, the truth values, an error and the empty text.
const OPERANDS: [Held; 12] = [
    Held::Number("3"),
    Held::Number("5"),
    Held::Number("7"),
    Held::Number("9"),
    Held::Number("11"),
    Held::Text("abc"),
    Held::Text("5"),
    Held::Empty,
    Held::Truth(true),
    Held::Truth(false),
    Held::Error("#N/A"),
    Held::EmptyText,
];

/// One of `cellIs`' cases at a row: its rule over the row's rectangle, read at
/// the operands.
fn operands_case(s: &mut Sheet, row: u32, name: &str, operator: &str, bounds: &[&str]) {
    let last = 3 + 2 * (OPERANDS.len() as u32 - 1);
    let cells: Vec<String> = (0..OPERANDS.len() as u32)
        .map(|i| cell_ref(row, 3 + 2 * i))
        .collect();
    s.part.case(name, &reads_of(&cells));
    for (i, value) in OPERANDS.iter().enumerate() {
        hold(&mut s.part, row, 3 + 2 * i as u32, *value, 0);
    }
    let extra = format!(" operator=\"{operator}\"");
    let rule = s.rule("cellIs", Some(LIT), &extra, bounds);
    s.block(&format!("C{row}:{}", cell_ref(row, last)), &[rule]);
}

/// `cellIs` (the probes' B).
fn cell_is(styles: &mut Styles) -> Part {
    let mut s = Sheet::new("Cell", styles);
    for (name, operator, bounds) in [
        ("less than 5", "lessThan", &["5"][..]),
        ("less than or equal to 5", "lessThanOrEqual", &["5"][..]),
        ("equal to 5", "equal", &["5"][..]),
        ("not equal to 5", "notEqual", &["5"][..]),
        (
            "greater than or equal to 5",
            "greaterThanOrEqual",
            &["5"][..],
        ),
        ("greater than 5", "greaterThan", &["5"][..]),
        ("between 5 and 9", "between", &["5", "9"][..]),
        ("not between 5 and 9", "notBetween", &["5", "9"][..]),
        ("between 9 and 5", "between", &["9", "5"][..]),
        ("not between 9 and 5", "notBetween", &["9", "5"][..]),
        ("equal to 0", "equal", &["0"][..]),
        ("equal to the empty text", "equal", &["\"\""][..]),
        ("equal to the text ABC", "equal", &["\"ABC\""][..]),
        ("less than the text abc", "lessThan", &["\"abc\""][..]),
        ("equal to TRUE", "equal", &["TRUE"][..]),
        ("equal to NA", "equal", &["NA()"][..]),
        ("not equal to NA", "notEqual", &["NA()"][..]),
    ] {
        let row = s.part.next_row();
        operands_case(&mut s, row, name, operator, bounds);
    }

    // A bound that moves with its cell: `C` reads `AE`, `E` reads `AG`, and on
    // two columns to a column of the operands.
    let row = s.part.next_row();
    s.part.number(row, 31, "3", 0);
    s.part.number(row, 33, "6", 0);
    s.part.number(row, 35, "7", 0);
    let bound = format!("AE{row}");
    operands_case(
        &mut s,
        row,
        "equal to a bound that moves",
        "equal",
        &[&bound],
    );

    // A bound reading a cell that holds nothing, past the operands: the first
    // run read one as the number 0, whatever the cell holds.
    for (name, operator, second) in [
        ("equal to an empty cell", "equal", None),
        ("not equal to an empty cell", "notEqual", None),
        ("less than an empty cell", "lessThan", None),
        (
            "less than or equal to an empty cell",
            "lessThanOrEqual",
            None,
        ),
        ("greater than an empty cell", "greaterThan", None),
        (
            "greater than or equal to an empty cell",
            "greaterThanOrEqual",
            None,
        ),
        ("between an empty cell and 9", "between", Some("9")),
        ("not between an empty cell and 9", "notBetween", Some("9")),
    ] {
        let row = s.part.next_row();
        let bound = format!("$AA{row}");
        match second {
            Some(b) => operands_case(&mut s, row, name, operator, &[&bound, b]),
            None => operands_case(&mut s, row, name, operator, &[&bound]),
        }
    }
    s.part
}

/// A case of three columns and two rows `C:E`, the values it reads at `G:I`.
fn moved_case(s: &mut Sheet, name: &str, test: &dyn Fn(u32) -> String, values: [[&str; 3]; 2]) {
    let row = s.part.next_row();
    let reads = format!(
        "shown:C{row} shown:D{row} shown:E{row} shown:C{0} shown:D{0} shown:E{0} rule:C{row}",
        row + 1
    );
    s.part.case(name, &reads);
    for (dy, line) in values.iter().enumerate() {
        for (dx, value) in line.iter().enumerate() {
            s.part.number(row + dy as u32, 7 + dx as u32, value, 0);
        }
    }
    s.over(&format!("C{row}:E{}", row + 1), LIT, &test(row));
}

/// How a reference moves (the probes' C): each mark's way over a rectangle.
fn references(styles: &mut Styles) -> Part {
    let mut s = Sheet::new("Refs", styles);
    let mixed = [["1", "0", "1"], ["0", "1", "0"]];
    moved_case(
        &mut s,
        "a relative reference",
        &|r| format!("G{r}=1"),
        mixed,
    );
    moved_case(
        &mut s,
        "a reference of a fixed column",
        &|r| format!("$G{r}=1"),
        [["1", "0", "0"], ["0", "1", "1"]],
    );
    moved_case(
        &mut s,
        "a reference of a fixed row",
        &|r| format!("G${r}=1"),
        mixed,
    );
    moved_case(
        &mut s,
        "a reference of a fixed cell",
        &|r| format!("$G${r}=1"),
        [["1", "0", "0"], ["0", "0", "0"]],
    );

    // A range's reference: an array in Excel, an intersection in a cell's
    // reading.
    for (name, test) in [
        (
            "a range under AND",
            &(|r: u32| format!("AND($G{r}:$I{r}=1)")) as &dyn Fn(u32) -> String,
        ),
        ("a range under SUMPRODUCT", &|r: u32| {
            format!("SUMPRODUCT(($G{r}:$I{r}=1)*1)=3")
        }),
        ("a range under COUNTIF", &|r: u32| {
            format!("COUNTIF($G{r}:$I{r},1)=3")
        }),
    ] {
        let row = s.part.next_row();
        s.part.case(name, &reads_of(&[format!("C{row}")]));
        s.part.text(row, 3, PLAIN_TEXT, 0);
        for col in 7..=9 {
            s.part.number(row, col, "1", 0);
        }
        s.over(&format!("C{row}"), LIT, &test(row));
    }

    // A rule reading where it stands, and a column whole, as real rules do.
    let row = s.part.next_row();
    s.part.case(
        "a band by the row",
        &format!("shown:C{row} shown:C{} rule:C{row}", row + 1),
    );
    s.over(&format!("C{row}:E{}", row + 1), LIT, "MOD(ROW(),2)=0");
    let row = s.part.next_row();
    s.part.case(
        "a band by the column",
        &format!("shown:C{row} shown:D{row} shown:E{row} rule:C{row}"),
    );
    s.over(&format!("C{row}:E{row}"), LIT, "MOD(COLUMN(),2)=0");
    let row = s.part.next_row();
    s.part.case(
        "a column read whole",
        &format!("shown:C{row} shown:E{row} rule:C{row}"),
    );
    s.part.number(row, 3, "5", 0);
    s.part.number(row, 5, "6", 0);
    s.part.number(row, 11, "5", 0);
    s.part.number(row + 1, 11, "5", 0);
    s.over(
        &format!("C{row} E{row}"),
        LIT,
        &format!("COUNTIF($K:$K,C{row})>1"),
    );

    // A rule over a merged range, whose cells answer apart.
    for (name, values) in [
        (
            "a merged range whose first cell answers false",
            ["0", "1", "1"],
        ),
        (
            "a merged range whose first cell answers true",
            ["1", "0", "0"],
        ),
    ] {
        let row = s.part.next_row();
        s.part.case(
            name,
            &format!("shown:C{row} shown:D{row} shown:E{row} rule:C{row}"),
        );
        s.part.text(row, 3, PLAIN_TEXT, 0);
        s.part.merges.push(format!("C{row}:E{row}"));
        for (dx, value) in values.iter().enumerate() {
            s.part.number(row, 7 + dx as u32, value, 0);
        }
        s.over(&format!("C{row}:E{row}"), LIT, &format!("G{row}=1"));
    }
    s.part
}

/// A case of several rectangles (the probes' C): its rows are its own, the
/// values at the columns each candidate for the formula's base reads, chosen
/// so that each base lights another set of its cells.
fn based_case(
    s: &mut Sheet,
    row: u32,
    name: &str,
    sqref: &str,
    test: &str,
    values: &[(&str, &str)],
) {
    let cells: Vec<&str> = sqref
        .split(' ')
        .flat_map(|r| r.split(':'))
        .collect::<Vec<_>>();
    let shown: Vec<String> = cells.iter().map(|c| format!("shown:{c}")).collect();
    s.part
        .case_at(row, name, &format!("{} rule:{}", shown.join(" "), cells[0]));
    for (at, value) in values {
        let (col, line) = split_ref(at);
        s.part.number(line, col, value, 0);
    }
    s.over(sqref, LIT, test);
}

/// A cell's address as its column's number and its row's.
fn split_ref(at: &str) -> (u32, u32) {
    let digits = at
        .find(|c: char| c.is_ascii_digit())
        .expect("an address has a row");
    let col = at[..digits]
        .bytes()
        .fold(0u32, |n, b| 26 * n + u32::from(b - b'A') + 1);
    let row: u32 = at[digits..].parse().expect("an address's row");
    assert_eq!(column_letters(col), at[..digits], "{at}");
    (col, row)
}

/// Where a formula's base stands in a range of several rectangles.
fn bases(styles: &mut Styles) -> Part {
    let mut s = Sheet::new("Bases", styles);
    // First a rectangle below and to the right of the second. The first's
    // first cell, `F4`: `D2` reads `H2`. The corner, `D2`: `D2` reads `J4`.
    based_case(
        &mut s,
        2,
        "the lower rectangle first",
        "F4:F5 D2:D3",
        "J4=1",
        &[
            ("H2", "1"),
            ("H3", "0"),
            ("J4", "0"),
            ("J5", "1"),
            ("L6", "1"),
            ("L7", "0"),
        ],
    );
    // A rectangle below and to the left of the second, where the corner is
    // no cell of the range. Its first's first cell, `D12`; the first cell by
    // rows, `F10`; the corner, `D10`.
    based_case(
        &mut s,
        10,
        "the lower left rectangle first",
        "D12:D13 F10:F11",
        "H12=1",
        &[
            ("H12", "1"),
            ("H13", "0"),
            ("J10", "0"),
            ("J11", "1"),
            ("F14", "0"),
            ("F15", "1"),
            ("H14", "1"),
            ("H15", "1"),
            ("J12", "1"),
            ("J13", "1"),
        ],
    );
    // The first rectangle the top left one: every candidate agrees.
    based_case(
        &mut s,
        18,
        "the upper left rectangle first",
        "D18:D19 F20:F21",
        "H18=1",
        &[("H18", "1"), ("H19", "0"), ("J20", "0"), ("J21", "1")],
    );
    // Side by side, the right one first.
    based_case(
        &mut s,
        26,
        "the right rectangle first",
        "F26:F27 D26:D27",
        "J26=1",
        &[
            ("H26", "1"),
            ("H27", "0"),
            ("J26", "0"),
            ("J27", "1"),
            ("L26", "1"),
            ("L27", "0"),
        ],
    );
    // A bound moves from the same place as a formula. The corner, `D34`: `D34`
    // reads `J36` and `F37` reads `L39`.
    s.part.case_at(
        34,
        "a bound over two rectangles, the lower first",
        "shown:F36 shown:F37 shown:D34 shown:D35 rule:F36",
    );
    for (at, value) in [
        ("D34", "1"),
        ("D35", "2"),
        ("F36", "3"),
        ("F37", "4"),
        ("J36", "1"),
        ("J37", "0"),
        ("L38", "3"),
        ("L39", "0"),
    ] {
        let (col, line) = split_ref(at);
        s.part.number(line, col, value, 0);
    }
    let rule = s.rule("cellIs", Some(LIT), " operator=\"equal\"", &["J36"]);
    s.block("F36:F37 D34:D35", &[rule]);
    s.part
}

/// A reference moved past the sheet's edge (the probes' C).
fn edges(styles: &mut Styles) -> Part {
    let mut s = Sheet::new("Edges", styles);
    // A rule moves from its range's top left, down and right, so a reference
    // moves past the bottom or the right and never the top. Past the bottom:
    // from `C3` the rule's `E1048576` is a row below the last, `E1`.
    s.part.case_at(
        2,
        "a reference moved past the bottom",
        "shown:C2 shown:C3 rule:C2",
    );
    for (at, value) in [("E1", "1"), ("E1048576", "0")] {
        let (col, row) = split_ref(at);
        s.part.number(row, col, value, 0);
    }
    s.over("C2:C3", LIT, "E1048576=1");
    // Past the right: from `E6` the rule's `XFD7` is two columns past the
    // last, `B7`.
    s.part.case_at(
        6,
        "a reference moved past the right",
        "shown:C6 shown:E6 rule:C6",
    );
    for (at, value) in [("XFD7", "0"), ("B7", "1")] {
        let (col, row) = split_ref(at);
        s.part.number(row, col, value, 0);
    }
    s.over("C6:E6", LIT, "XFD7=1");
    s.part
}

const BOLD: &str = "<font><b/></font>";
const ITALIC: &str = "<font><i/></font>";

/// A case of the rules over `C`, its cell holding a text.
fn ordered(s: &mut Sheet, name: &str, reads: &[&str]) -> u32 {
    let row = s.part.next_row();
    let cells: Vec<String> = reads.iter().map(|c| format!("{c}{row}")).collect();
    let mut all: Vec<String> = cells.iter().map(|c| format!("shown:{c}")).collect();
    all.extend(cells.iter().map(|c| format!("rule:{c}")));
    s.part.case(name, &all.join(" "));
    for c in reads {
        let (col, _) = split_ref(&format!("{c}1"));
        s.part.text(row, col, PLAIN_TEXT, 0);
    }
    row
}

/// The order of several rules (the probes' D).
fn order(styles: &mut Styles) -> Part {
    let mut s = Sheet::new("Order", styles);
    let row = ordered(&mut s, "two true rules saying one property", &["C"]);
    let rules = [s.test(Some(LIT), "", "1"), s.test(Some(OTHER), "", "1")];
    s.block(&format!("C{row}"), &rules);

    let row = ordered(&mut s, "a false rule over a true one", &["C"]);
    let rules = [s.test(Some(LIT), "", "0"), s.test(Some(OTHER), "", "1")];
    s.block(&format!("C{row}"), &rules);

    let row = ordered(&mut s, "two true rules saying different properties", &["C"]);
    let rules = [s.test(Some(LIT), "", "1"), s.test(Some(BOLD), "", "1")];
    s.block(&format!("C{row}"), &rules);

    // A file's order against its priorities: the element first in the file
    // has the higher number.
    let row = ordered(&mut s, "a file's order against its priorities", &["C"]);
    let lower = s.priority();
    let higher = s.priority();
    let first = s.rule_at(higher, "expression", Some(OTHER), "", &["1"]);
    let second = s.rule_at(lower, "expression", Some(LIT), "", &["1"]);
    s.block(&format!("C{row}"), &[first, second]);

    let row = ordered(&mut s, "two elements against their priorities", &["C", "D"]);
    let lower = s.priority();
    let higher = s.priority();
    let first = s.rule_at(higher, "expression", Some(OTHER), "", &["1"]);
    let second = s.rule_at(lower, "expression", Some(LIT), "", &["1"]);
    s.block(&format!("C{row}:D{row}"), &[first]);
    s.block(&format!("C{row}"), &[second]);

    let row = ordered(&mut s, "a true rule that stops", &["C"]);
    let rules = [
        s.test(Some(LIT), " stopIfTrue=\"1\"", "1"),
        s.test(Some(BOLD), "", "1"),
    ];
    s.block(&format!("C{row}"), &rules);

    let row = ordered(&mut s, "a second rule that stops", &["C"]);
    let rules = [
        s.test(Some(BOLD), "", "1"),
        s.test(Some(LIT), " stopIfTrue=\"1\"", "1"),
        s.test(Some(ITALIC), "", "1"),
    ];
    s.block(&format!("C{row}"), &rules);

    let row = ordered(&mut s, "a false rule that stops", &["C"]);
    let rules = [
        s.test(Some(OTHER), " stopIfTrue=\"1\"", "0"),
        s.test(Some(LIT), "", "1"),
    ];
    s.block(&format!("C{row}"), &rules);

    let row = ordered(&mut s, "a rule of no format that stops", &["C"]);
    let rules = [
        s.test(None, " stopIfTrue=\"1\"", "1"),
        s.test(Some(LIT), "", "1"),
    ];
    s.block(&format!("C{row}"), &rules);

    let row = ordered(&mut s, "a rule of no format that goes on", &["C"]);
    let rules = [s.test(None, "", "1"), s.test(Some(LIT), "", "1")];
    s.block(&format!("C{row}"), &rules);

    let row = ordered(&mut s, "no fill said over a fill", &["C"]);
    let rules = [
        s.test(
            Some("<fill><patternFill patternType=\"none\"/></fill>"),
            "",
            "1",
        ),
        s.test(Some(LIT), "", "1"),
    ];
    s.block(&format!("C{row}"), &rules);

    let row = ordered(&mut s, "not bold said over bold", &["C"]);
    let rules = [
        s.test(Some("<font><b val=\"0\"/></font>"), "", "1"),
        s.test(Some(BOLD), "", "1"),
    ];
    s.block(&format!("C{row}"), &rules);

    // Edge by edge, each format of Excel's spelling, every side named.
    let row = ordered(&mut s, "two rules saying two edges", &["C"]);
    let rules = [
        s.test(
            Some("<border><left style=\"thin\"><color rgb=\"FF800020\"/></left><right/><top/><bottom/></border>"),
            "",
            "1",
        ),
        s.test(
            Some("<border><left style=\"thin\"><color rgb=\"FF0000FF\"/></left><right style=\"thin\"><color rgb=\"FF0000FF\"/></right><top/><bottom/></border>"),
            "",
            "1",
        ),
    ];
    s.block(&format!("C{row}"), &rules);

    let row = ordered(&mut s, "two rules saying a side alone", &["C"]);
    let rules = [
        s.test(
            Some("<border><left style=\"thin\"><color rgb=\"FF800020\"/></left></border>"),
            "",
            "1",
        ),
        s.test(
            Some("<border><right style=\"thin\"><color rgb=\"FF0000FF\"/></right></border>"),
            "",
            "1",
        ),
    ];
    s.block(&format!("C{row}"), &rules);

    // A font's words, each the first true rule's to say.
    let row = ordered(&mut s, "a color, a weight and a second color", &["C"]);
    let rules = [
        s.test(Some("<font><color rgb=\"FF800020\"/></font>"), "", "1"),
        s.test(Some(BOLD), "", "1"),
        s.test(Some("<font><color rgb=\"FF0000FF\"/></font>"), "", "1"),
    ];
    s.block(&format!("C{row}"), &rules);

    // Ranges that meet in one cell.
    let row = ordered(&mut s, "two ranges meeting in a cell", &["C", "E"]);
    let first = s.test(Some(OTHER), "", "1");
    let second = s.test(Some(LIT), "", "1");
    s.block(&format!("C{row}"), &[first]);
    s.block(&format!("C{row}:E{row}"), &[second]);

    let row = ordered(&mut s, "a stop in the cell two ranges meet in", &["C", "E"]);
    let first = s.test(Some(OTHER), " stopIfTrue=\"1\"", "1");
    let second = s.test(Some(BOLD), "", "1");
    s.block(&format!("C{row}"), &[first]);
    s.block(&format!("C{row} E{row}"), &[second]);

    // How far a stop reaches: the first run's did not reach a rule of another
    // range.
    let stop = " stopIfTrue=\"1\"";
    let row = ordered(&mut s, "a stop and a second element of its range", &["C"]);
    let first = s.test(Some(OTHER), stop, "1");
    let second = s.test(Some(BOLD), "", "1");
    s.block(&format!("C{row}"), &[first]);
    s.block(&format!("C{row}"), &[second]);

    let row = ordered(&mut s, "a stop over a wider range", &["C", "E"]);
    let first = s.test(Some(OTHER), stop, "1");
    let second = s.test(Some(BOLD), "", "1");
    s.block(&format!("C{row}:E{row}"), &[first]);
    s.block(&format!("C{row}"), &[second]);

    let row = ordered(&mut s, "a stop under a wider range", &["C", "E"]);
    let first = s.test(Some(OTHER), stop, "1");
    let second = s.test(Some(BOLD), "", "1");
    s.block(&format!("C{row}"), &[first]);
    s.block(&format!("C{row}:E{row}"), &[second]);

    let row = ordered(
        &mut s,
        "a stop and its range spelled the other way",
        &["C", "E"],
    );
    let first = s.test(Some(OTHER), stop, "1");
    let second = s.test(Some(BOLD), "", "1");
    s.block(&format!("C{row} E{row}"), &[first]);
    s.block(&format!("E{row} C{row}"), &[second]);

    let row = ordered(&mut s, "a stop between two ranges", &["C", "E"]);
    let bold = s.test(Some(BOLD), "", "1");
    let stopping = s.test(Some(OTHER), stop, "1");
    let italic = s.test(Some(ITALIC), "", "1");
    let color = s.test(Some("<font><color rgb=\"FF800020\"/></font>"), "", "1");
    s.block(&format!("C{row}"), &[bold, color]);
    s.block(&format!("C{row}:E{row}"), &[stopping, italic]);

    let row = ordered(&mut s, "a stop and a later rule of each range", &["C", "E"]);
    let narrow = [
        s.test(Some(OTHER), stop, "1"),
        s.test(Some(ITALIC), "", "1"),
    ];
    let wide = s.test(Some(BOLD), "", "1");
    s.block(&format!("C{row}"), &narrow);
    s.block(&format!("C{row}:E{row}"), &[wide]);

    let row = ordered(
        &mut s,
        "a stop of no format over a wider range",
        &["C", "E"],
    );
    let first = s.test(None, stop, "1");
    let second = s.test(Some(LIT), "", "1");
    s.block(&format!("C{row}"), &[first]);
    s.block(&format!("C{row}:E{row}"), &[second]);
    s.part
}

/// Two rules of one priority on a sheet of its own, where a file may hold
/// them and no dialog of Excel's writes them.
fn alike(styles: &mut Styles) -> Part {
    let mut s = Sheet::new("Alike", styles);
    let row = ordered(&mut s, "one priority twice, this fill first", &["C"]);
    let rules = [
        s.rule_at(1, "expression", Some(LIT), "", &["1"]),
        s.rule_at(1, "expression", Some(OTHER), "", &["1"]),
    ];
    s.block(&format!("C{row}"), &rules);
    let row = ordered(&mut s, "one priority twice, that fill first", &["C"]);
    let rules = [
        s.rule_at(2, "expression", Some(OTHER), "", &["1"]),
        s.rule_at(2, "expression", Some(LIT), "", &["1"]),
    ];
    s.block(&format!("C{row}"), &rules);

    // The first run's later rule led: each of these two properties says which
    // of the three lead it.
    let row = ordered(&mut s, "one priority three times", &["C"]);
    let rules = [
        s.rule_at(
            3,
            "expression",
            Some("<font><color rgb=\"FF800020\"/></font><fill><patternFill><bgColor rgb=\"FFFFE699\"/></patternFill></fill>"),
            "",
            &["1"],
        ),
        s.rule_at(3, "expression", Some(OTHER), "", &["1"]),
        s.rule_at(
            3,
            "expression",
            Some("<font><color rgb=\"FF0000FF\"/></font>"),
            "",
            &["1"],
        ),
    ];
    s.block(&format!("C{row}"), &rules);

    let row = ordered(&mut s, "one priority in two elements", &["C"]);
    let first = s.rule_at(4, "expression", Some(LIT), "", &["1"]);
    let second = s.rule_at(4, "expression", Some(OTHER), "", &["1"]);
    s.block(&format!("C{row}"), &[first]);
    s.block(&format!("C{row}"), &[second]);
    s.part
}

/// A value every number format of `partial` draws, and a width that shows it.
const PARTIAL_VALUE: &str = "12.25";

/// A left edge in a line and a color, Excel's other sides named empty.
fn left(line: &str, color: &str) -> String {
    format!(
        "<border><left style=\"{line}\"><color {color}/></left><right/><top/><bottom/></border>"
    )
}

/// A format's parts over a cell's own (the probes' E): each case's rule true
/// over a cell of a full look of its own, `C`, and a cell of none, `E`.
fn partial(styles: &mut Styles) -> Part {
    let mut own = Xf {
        num_fmt: styles.code("0.0"),
        font: Some(styles.font("<b/><i/><strike/><u/>", "<color rgb=\"FF0000FF\"/>")),
        fill: Some(solid("rgb=\"FFCCFFCC\"")),
        ..Xf::default()
    };
    own.border = Some(
        "<border><left style=\"thin\"><color rgb=\"FF000000\"/></left><right style=\"thin\"><color rgb=\"FF000000\"/></right><top style=\"thin\"><color rgb=\"FF000000\"/></top><bottom style=\"thin\"><color rgb=\"FF000000\"/></bottom><diagonal/></border>"
            .to_string(),
    );
    let own = styles.xf(own);
    let mut s = Sheet::new("Partial", styles);
    s.part
        .cols
        .push("<col min=\"3\" max=\"5\" width=\"20\" customWidth=\"1\"/>".to_string());

    let case = |s: &mut Sheet, name: &str| -> u32 {
        let row = s.part.next_row();
        s.part.case(
            name,
            &format!("shown:C{row} text:C{row} shown:E{row} text:E{row} rule:C{row}"),
        );
        s.part.number(row, 3, PARTIAL_VALUE, own);
        s.part.number(row, 5, PARTIAL_VALUE, 0);
        row
    };

    // The cell's own with no rule, and under a rule that is false.
    case(&mut s, "no rule");
    let row = case(&mut s, "a false rule");
    s.over(&format!("C{row} E{row}"), LIT, "0");

    let edges = [
        (
            "a right edge",
            "<border><left/><right style=\"thin\"><color rgb=\"FF800020\"/></right><top/><bottom/></border>"
                .to_string(),
        ),
        (
            "a top edge",
            "<border><left/><right/><top style=\"thin\"><color rgb=\"FF800020\"/></top><bottom/></border>"
                .to_string(),
        ),
        (
            "a bottom edge",
            "<border><left/><right/><top/><bottom style=\"thin\"><color rgb=\"FF800020\"/></bottom></border>"
                .to_string(),
        ),
        ("a thin left edge", left("thin", "rgb=\"FF800020\"")),
        ("a medium left edge", left("medium", "rgb=\"FF800020\"")),
        ("a thick left edge", left("thick", "rgb=\"FF800020\"")),
        ("a double left edge", left("double", "rgb=\"FF800020\"")),
        ("a dashed left edge", left("dashed", "rgb=\"FF800020\"")),
        ("a left edge of an automatic color", left("thin", "auto=\"1\"")),
        (
            "a left edge alone",
            "<border><left style=\"thin\"><color rgb=\"FF800020\"/></left></border>".to_string(),
        ),
        (
            "every side said empty",
            "<border><left/><right/><top/><bottom/></border>".to_string(),
        ),
        (
            "a left edge of no style",
            "<border><left style=\"none\"/></border>".to_string(),
        ),
    ];
    let mut formats: Vec<(&str, String)> = [
        // A fill.
        ("a fill's bgColor", LIT),
        (
            "a fill's fgColor under solid",
            "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FF800020\"/></patternFill></fill>",
        ),
        (
            "a fill's fgColor and bgColor under solid",
            "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FF800020\"/><bgColor rgb=\"FFFFE699\"/></patternFill></fill>",
        ),
        (
            "a fill's pattern of none",
            "<fill><patternFill patternType=\"none\"/></fill>",
        ),
        (
            "a fill of no color as Excel writes it",
            "<fill><patternFill patternType=\"none\"><bgColor auto=\"1\"/></patternFill></fill>",
        ),
        (
            "a fill's pattern",
            "<fill><patternFill patternType=\"gray125\"><fgColor rgb=\"FF800020\"/><bgColor rgb=\"FFFFE699\"/></patternFill></fill>",
        ),
        (
            "a fill's theme color",
            "<fill><patternFill><bgColor theme=\"4\" tint=\"0.39997558519241921\"/></patternFill></fill>",
        ),
        (
            "a fill's indexed color",
            "<fill><patternFill><bgColor indexed=\"10\"/></patternFill></fill>",
        ),
        // A font.
        ("not bold", "<font><b val=\"0\"/></font>"),
        ("not italic", "<font><i val=\"0\"/></font>"),
        ("bold", BOLD),
        ("italic", ITALIC),
        ("a font's color", "<font><color rgb=\"FF800020\"/></font>"),
        ("a font's theme color", "<font><color theme=\"4\"/></font>"),
        ("a font's automatic color", "<font><color auto=\"1\"/></font>"),
        ("an underline", "<font><u/></font>"),
        ("a double underline", "<font><u val=\"double\"/></font>"),
        (
            "an accounting underline",
            "<font><u val=\"singleAccounting\"/></font>",
        ),
        (
            "a double accounting underline",
            "<font><u val=\"doubleAccounting\"/></font>",
        ),
        ("no underline", "<font><u val=\"none\"/></font>"),
        ("a strike", "<font><strike/></font>"),
        ("no strike", "<font><strike val=\"0\"/></font>"),
        // A number format, by the text a cell shows.
        (
            "a number format of its own",
            "<numFmt numFmtId=\"200\" formatCode=\"0.000\"/>",
        ),
        (
            "a built-in number format",
            "<numFmt numFmtId=\"2\" formatCode=\"0.00\"/>",
        ),
        (
            "the general number format",
            "<numFmt numFmtId=\"0\" formatCode=\"General\"/>",
        ),
        (
            "the text number format",
            "<numFmt numFmtId=\"49\" formatCode=\"@\"/>",
        ),
        (
            "a number format with a literal",
            "<numFmt numFmtId=\"201\" formatCode=\"0.00&quot; kg&quot;\"/>",
        ),
        (
            "a number format with a color",
            "<numFmt numFmtId=\"202\" formatCode=\"[Red]0.00\"/>",
        ),
        // What a format holds and no dialog of Excel's writes.
        ("a size", "<font><sz val=\"20\"/></font>"),
        ("a name", "<font><name val=\"Consolas\"/></font>"),
        (
            "a raised text",
            "<font><vertAlign val=\"superscript\"/></font>",
        ),
        ("a horizontal alignment", "<alignment horizontal=\"center\"/>"),
        ("a vertical alignment", "<alignment vertical=\"top\"/>"),
        ("a wrap", "<alignment wrapText=\"1\"/>"),
        (
            "a size beside a fill",
            "<font><sz val=\"20\"/></font><fill><patternFill><bgColor rgb=\"FFFFE699\"/></patternFill></fill>",
        ),
    ]
    .iter()
    .map(|(name, body)| (*name, (*body).to_string()))
    .collect();
    formats.extend(edges);
    for (name, body) in &formats {
        let row = case(&mut s, name);
        s.over(&format!("C{row} E{row}"), body, "1");
    }
    s.part
}

/// The kinds with a formula of their own and those without (the probes' G).
fn kinds(styles: &mut Styles) -> Part {
    let mut s = Sheet::new("Kinds", styles);
    let held = [
        Held::Text("xaby"),
        Held::Text("zzz"),
        Held::Number("5"),
        Held::Empty,
        Held::Error("#N/A"),
        Held::Spaces,
        Held::EmptyText,
    ];
    for (name, kind, extra, test) in [
        (
            "contains a text",
            "containsText",
            " operator=\"containsText\" text=\"ab\"",
            "NOT(ISERROR(SEARCH(\"ab\",C{r})))",
        ),
        (
            "contains no such text",
            "notContainsText",
            " operator=\"notContains\" text=\"ab\"",
            "ISERROR(SEARCH(\"ab\",C{r}))",
        ),
        (
            "begins with a text",
            "beginsWith",
            " operator=\"beginsWith\" text=\"xa\"",
            "LEFT(C{r},LEN(\"xa\"))=\"xa\"",
        ),
        (
            "ends with a text",
            "endsWith",
            " operator=\"endsWith\" text=\"by\"",
            "RIGHT(C{r},LEN(\"by\"))=\"by\"",
        ),
        ("is blank", "containsBlanks", "", "LEN(TRIM(C{r}))=0"),
        ("is not blank", "notContainsBlanks", "", "LEN(TRIM(C{r}))>0"),
        ("is an error", "containsErrors", "", "ISERROR(C{r})"),
        (
            "is not an error",
            "notContainsErrors",
            "",
            "NOT(ISERROR(C{r}))",
        ),
    ] {
        let row = s.part.next_row();
        let cells: Vec<String> = (0..held.len() as u32)
            .map(|i| cell_ref(row, 3 + 2 * i))
            .collect();
        s.part.case(name, &reads_of(&cells));
        for (i, value) in held.iter().enumerate() {
            hold(&mut s.part, row, 3 + 2 * i as u32, *value, 0);
        }
        let test = test.replace("{r}", &row.to_string());
        let rule = s.rule(kind, Some(LIT), extra, &[&test]);
        s.block(&format!("C{row}:O{row}"), &[rule]);
    }

    // A date of long ago, so that the run's own day answers alike.
    let row = s.part.next_row();
    s.part.case("is today", &reads_of(&[format!("C{row}")]));
    s.part.number(row, 3, "36526", 0);
    let test = format!("FLOOR(C{row},1)=TODAY()");
    let rule = s.rule("timePeriod", Some(LIT), " timePeriod=\"today\"", &[&test]);
    s.block(&format!("C{row}"), &[rule]);

    // The kinds with no formula, over 1, 5 and 9 and, for a duplicate, two
    // cells alike.
    let scales = [
        (
            "a color scale",
            "colorScale",
            "<colorScale><cfvo type=\"min\"/><cfvo type=\"max\"/><color rgb=\"FFF8696B\"/><color rgb=\"FF63BE7B\"/></colorScale>",
        ),
        (
            "a data bar",
            "dataBar",
            "<dataBar><cfvo type=\"min\"/><cfvo type=\"max\"/><color rgb=\"FF638EC6\"/></dataBar>",
        ),
        (
            "an icon set",
            "iconSet",
            "<iconSet><cfvo type=\"percent\" val=\"0\"/><cfvo type=\"percent\" val=\"33\"/><cfvo type=\"percent\" val=\"67\"/></iconSet>",
        ),
    ];
    for (name, kind, body) in scales {
        let row = s.part.next_row();
        let cells = [format!("C{row}"), format!("E{row}"), format!("G{row}")];
        s.part.case(name, &reads_of(&cells));
        for (col, value) in [(3, "1"), (5, "5"), (7, "9")] {
            s.part.number(row, col, value, 0);
        }
        let priority = s.priority();
        let rule = format!("<cfRule type=\"{kind}\" priority=\"{priority}\">{body}</cfRule>");
        s.block(&format!("C{row} E{row} G{row}"), &[rule]);
    }
    for (name, kind, extra, values) in [
        ("the top value", "top10", " rank=\"1\"", ["1", "5", "9"]),
        ("above the average", "aboveAverage", "", ["1", "5", "9"]),
        ("a duplicate", "duplicateValues", "", ["1", "1", "9"]),
        ("a value alone", "uniqueValues", "", ["1", "1", "9"]),
    ] {
        let row = s.part.next_row();
        let cells = [format!("C{row}"), format!("E{row}"), format!("G{row}")];
        s.part.case(name, &reads_of(&cells));
        for (col, value) in [3, 5, 7].into_iter().zip(values) {
            s.part.number(row, col, value, 0);
        }
        let rule = s.rule(kind, Some(LIT), extra, &[]);
        s.block(&format!("C{row} E{row} G{row}"), &[rule]);
    }
    s.part
}

/// A rule reading another sheet through a defined name (the probes' F), and a
/// name a sheet holds over the workbook's own of one spelling.
fn names(styles: &mut Styles) -> (Part, Part) {
    let mut s = Sheet::new("Names", styles);
    for (name, test) in [
        ("a name of a cell holding 1", "seen_one=1"),
        ("a name of a cell holding 0", "seen_zero=1"),
        ("a sheet's own name over the workbook's", "pick=1"),
    ] {
        let row = s.part.next_row();
        s.part.case(name, &reads_of(&[format!("C{row}")]));
        s.part.text(row, 3, PLAIN_TEXT, 0);
        s.over(&format!("C{row}"), LIT, test);
    }
    let names = s.part;
    let mut s = Sheet::new("Elsewhere", styles);
    let row = s.part.next_row();
    s.part.case(
        "the workbook's name, no name of its own",
        &reads_of(&[format!("C{row}")]),
    );
    s.part.text(row, 3, PLAIN_TEXT, 0);
    s.over(&format!("C{row}"), LIT, "pick=1");
    s.part.number(3, 3, "1", 0);
    s.part.number(5, 3, "0", 0);
    (names, s.part)
}

/// `rules.xlsx`, its sheets in the probes' order.
fn rules() -> Book {
    let mut styles = Styles::new("11", CALIBRI, "<color theme=\"1\"/>");
    let (names_part, elsewhere) = names(&mut styles);
    let mut parts = vec![
        truth(&mut styles),
        cell_is(&mut styles),
        references(&mut styles),
        bases(&mut styles),
        edges(&mut styles),
        order(&mut styles),
        alike(&mut styles),
        partial(&mut styles),
        kinds(&mut styles),
    ];
    let local = parts.len();
    parts.push(names_part);
    parts.push(elsewhere);
    Book {
        styles,
        parts,
        strings: Vec::new(),
        theme: true,
        names: vec![
            format!(
                "<definedName name=\"pick\" localSheetId=\"{local}\">Elsewhere!$C$3</definedName>"
            ),
            "<definedName name=\"pick\">Elsewhere!$C$5</definedName>".to_string(),
            "<definedName name=\"seen_one\">Elsewhere!$C$3</definedName>".to_string(),
            "<definedName name=\"seen_zero\">Elsewhere!$C$5</definedName>".to_string(),
        ],
    }
}

/// A rule reading another sheet, in the part or in the extension.
struct Elsewhere<'s> {
    sheet: Sheet<'s>,
    /// Each rule's element where the rule goes to the extension.
    extended: Option<Vec<String>>,
}

impl Elsewhere<'_> {
    /// A rule of a kind over a range, where the workbook's form writes it.
    fn rule(&mut self, sqref: &str, kind: &str, extra: &str, priority: u32, test: &str) {
        let raw = test.to_string();
        let Some(extended) = &mut self.extended else {
            let rule = self
                .sheet
                .rule_at(priority, kind, Some(LIT), extra, &[test]);
            self.sheet.block(sqref, &[rule]);
            return;
        };
        let test = row_test(kind, extra, &[test]).expect("a test of a formula");
        let stop = if extra.contains("stopIfTrue") {
            " stop"
        } else {
            ""
        };
        self.sheet.part.rule_rows.push(format!(
            "(conditional {} {} {} {test} {}{stop})\n",
            quoted(self.sheet.part.name),
            100 * priority,
            quoted(sqref),
            format_number(LIT),
        ));
        let test = &raw;
        let id = extended.len() + 1;
        extended.push(format!(
            "<x14:conditionalFormatting xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\"><x14:cfRule type=\"{kind}\" priority=\"{priority}\"{extra} id=\"{{00000000-0000-4000-8000-{id:012X}}}\"><xm:f>{}</xm:f><x14:dxf>{LIT}</x14:dxf></x14:cfRule><xm:sqref>{sqref}</xm:sqref></x14:conditionalFormatting>",
            crate::sheet::xml::text(test)
        ));
    }
}

/// A workbook of rules reading another sheet (the probes' F): `Rules`, and the
/// values it reads, on `Seen` and on the hidden `Unseen`.
fn elsewhere(extension: bool) -> Book {
    let mut styles = Styles::new("11", CALIBRI, "<color theme=\"1\"/>");
    let mut seen = Part::new("Seen");
    let mut unseen = Part::new("Unseen");
    unseen.hidden = true;
    let mut e = Elsewhere {
        sheet: Sheet::new("Rules", &mut styles),
        extended: extension.then(Vec::new),
    };
    let mut priority = 0;
    for (name, sheet) in [("visible", "Seen"), ("hidden", "Unseen")] {
        let values = if sheet == "Seen" {
            &mut seen
        } else {
            &mut unseen
        };

        priority += 1;
        let row = e.sheet.part.next_row();
        e.sheet.part.case(
            &format!("a fixed cell on a {name} sheet holding 1"),
            &reads_of(&[format!("C{row}")]),
        );
        e.sheet.part.text(row, 3, PLAIN_TEXT, 0);
        values.number(row, 3, "1", 0);
        e.rule(
            &format!("C{row}"),
            "expression",
            "",
            priority,
            &format!("{sheet}!$C${row}=1"),
        );

        priority += 1;
        let row = e.sheet.part.next_row();
        e.sheet.part.case(
            &format!("a fixed cell on a {name} sheet holding 0"),
            &reads_of(&[format!("C{row}")]),
        );
        e.sheet.part.text(row, 3, PLAIN_TEXT, 0);
        values.number(row, 3, "0", 0);
        e.rule(
            &format!("C{row}"),
            "expression",
            "",
            priority,
            &format!("{sheet}!$C${row}=1"),
        );

        // `C` reads the sheet's `C`, and `D` and `E` move with it.
        priority += 1;
        let row = e.sheet.part.next_row();
        e.sheet.part.case(
            &format!("a relative cell on a {name} sheet"),
            &format!("shown:C{row} shown:D{row} shown:E{row} rule:C{row}"),
        );
        for (col, value) in [(3, "1"), (4, "0"), (5, "1")] {
            e.sheet.part.text(row, col, PLAIN_TEXT, 0);
            values.number(row, col, value, 0);
        }
        e.rule(
            &format!("C{row}:E{row}"),
            "expression",
            "",
            priority,
            &format!("{sheet}!C{row}=1"),
        );

        priority += 1;
        let row = e.sheet.part.next_row();
        e.sheet.part.case(
            &format!("a cell's value against a {name} sheet"),
            &reads_of(&[format!("C{row}")]),
        );
        e.sheet.part.number(row, 3, "7", 0);
        values.number(row, 3, "7", 0);
        e.rule(
            &format!("C{row}"),
            "cellIs",
            " operator=\"equal\"",
            priority,
            &format!("{sheet}!$C${row}"),
        );
    }

    // A rule of the workbook's form against one in the sheet's own part, by
    // their priorities: in the extension's workbook these are the two parts'.
    let row = e.sheet.part.next_row();
    e.sheet.part.case(
        "a rule elsewhere over one of the sheet's own",
        &reads_of(&[format!("C{row}")]),
    );
    e.sheet.part.text(row, 3, PLAIN_TEXT, 0);
    seen.number(row, 3, "1", 0);
    priority += 2;
    e.rule(
        &format!("C{row}"),
        "expression",
        "",
        priority - 1,
        &format!("Seen!$C${row}=1"),
    );
    let rule = e
        .sheet
        .rule_at(priority, "expression", Some(OTHER), "", &["1"]);
    e.sheet.block(&format!("C{row}"), &[rule]);

    let row = e.sheet.part.next_row();
    e.sheet.part.case(
        "a rule of the sheet's own over one elsewhere",
        &reads_of(&[format!("C{row}")]),
    );
    e.sheet.part.text(row, 3, PLAIN_TEXT, 0);
    seen.number(row, 3, "1", 0);
    priority += 2;
    let rule = e
        .sheet
        .rule_at(priority - 1, "expression", Some(OTHER), "", &["1"]);
    e.sheet.block(&format!("C{row}"), &[rule]);
    e.rule(
        &format!("C{row}"),
        "expression",
        "",
        priority,
        &format!("Seen!$C${row}=1"),
    );

    // A stop in the workbook's form over a rule of the sheet's own on one
    // range: in the extension's workbook, across the two parts.
    let row = e.sheet.part.next_row();
    e.sheet.part.case(
        "a stop elsewhere over a rule of the sheet's own",
        &reads_of(&[format!("C{row}")]),
    );
    e.sheet.part.text(row, 3, PLAIN_TEXT, 0);
    seen.number(row, 3, "1", 0);
    priority += 2;
    e.rule(
        &format!("C{row}"),
        "expression",
        " stopIfTrue=\"1\"",
        priority - 1,
        &format!("Seen!$C${row}=1"),
    );
    let rule = e
        .sheet
        .rule_at(priority, "expression", Some(BOLD), "", &["1"]);
    e.sheet.block(&format!("C{row}"), &[rule]);

    let mut rules_part = e.sheet.part;
    if let Some(extended) = e.extended {
        rules_part.extension = format!(
            "<ext xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" uri=\"{{78C0D931-6437-407d-A8EE-F0AAD7539E65}}\"><x14:conditionalFormattings>{}</x14:conditionalFormattings></ext>",
            extended.concat()
        );
    }
    Book {
        styles,
        parts: vec![rules_part, seen, unseen],
        strings: Vec::new(),
        theme: true,
        names: Vec::new(),
    }
}

/// The three workbooks, each with its path and its file's name in the golden.
fn books() -> [(&'static str, &'static str, Book); 3] {
    [
        (RULES, "rules.xlsx", rules()),
        (PLAIN, "rules_plain.xlsx", elsewhere(false)),
        (X14, "rules_x14.xlsx", elsewhere(true)),
    ]
}

#[test]
fn the_rules_fixtures_are_their_builder_s() {
    let bless = std::env::var("FRAZARO_BLESS").as_deref() == Ok("1");
    for (path, _, book) in books() {
        let bytes = book.bytes();
        // Each reads, under the looks and without them.
        crate::reflect::reflect_looks(&bytes, path)
            .unwrap_or_else(|r| panic!("{path}: {}", r.text));
        crate::reflect::reflect_text(&bytes, path).unwrap_or_else(|r| panic!("{path}: {}", r.text));
        if bless {
            std::fs::write(path, &bytes).expect("the fixture written");
            continue;
        }
        let held = std::fs::read(path)
            .unwrap_or_else(|_| panic!("{path}: FRAZARO_BLESS=1 and the core's tests write it"));
        assert!(
            held == bytes,
            "{path} is not what its builder writes today: FRAZARO_BLESS=1 writes it again"
        );
    }
}

/// The golden's cases, in order: a header and the targets read, each once
/// however many rules a `rule:` target found.
fn golden_cases() -> Vec<(String, Vec<String>)> {
    let text = std::fs::read_to_string(GOLDEN)
        .unwrap_or_else(|_| panic!("{GOLDEN}: VlaWriteRulesGolden writes it"));
    let mut cases: Vec<(String, Vec<String>)> = Vec::new();
    for line in text.lines().skip(1) {
        if let Some(header) = line.strip_prefix("=== ") {
            cases.push((header.to_string(), Vec::new()));
            continue;
        }
        let (target, _) = line.split_once('\t').expect("a target and its reading");
        let (_, targets) = cases.last_mut().expect("a target under a case");
        if targets.last().map(String::as_str) != Some(target) {
            targets.push(target.to_string());
        }
    }
    cases
}

/// One case of the golden: its workbook, its sheet, its name and its targets.
struct GoldenCase {
    file: String,
    sheet: String,
    name: String,
    targets: Vec<(String, String)>,
}

fn golden_read() -> Vec<GoldenCase> {
    let text = std::fs::read_to_string(GOLDEN).expect("the golden");
    let mut cases: Vec<GoldenCase> = Vec::new();
    for line in text.lines().skip(1) {
        if let Some(header) = line.strip_prefix("=== ") {
            let mut parts = header.splitn(3, '!').map(str::to_string);
            let mut next = || parts.next().expect("a header's three parts");
            cases.push(GoldenCase {
                file: next(),
                sheet: next(),
                name: next(),
                targets: Vec::new(),
            });
        } else {
            let (target, reading) = line.split_once('\t').expect("a target and its reading");
            let case = cases.last_mut().expect("a target under a case");
            case.targets.push((target.to_string(), reading.to_string()));
        }
    }
    cases
}

/// A window of one cell's values record.
fn one_cell(m: &Machine, sheet: &str, place: &str) -> String {
    match m.view("values", sheet, Some(place)) {
        Ok(Viewed::Text(t)) => t,
        other => panic!("{sheet}!{place}: {other:?}"),
    }
}

/// The rules that apply at a cell, as its record's `lit` row names them.
fn lit_at(record: &str) -> Vec<u32> {
    let Some(line) = record.lines().find(|l| l.starts_with("(lit ")) else {
        return Vec::new();
    };
    let forms = vla_lang::reader::read_forms(line).expect("a row");
    let items = &forms[0].as_list().expect("a row is a list").items;
    items[3..]
        .iter()
        .map(|p| match p {
            Form::Sym(n) => n.parse().expect("a priority"),
            other => panic!("{other:?}"),
        })
        .collect()
}

/// What a cell draws as the language folds it: each rule that applies over the
/// next, the last over the cell's own (`Style::over`). DisplayFormat keeps the
/// cell's own number format wherever a rule says another, so that part of the
/// reading is the cell's own, and the text is held by `text:`.
fn folded(wb: &Workbook, sheet: &str, (row, col): (u32, u32), lit: &[u32]) -> Drawn {
    let page = wb
        .sheets
        .iter()
        .find(|s| s.name == sheet)
        .expect("the sheet");
    let index = page.cells.get(&(row, col)).map_or(0, |c| c.style);
    let default = Style::default();
    let own = wb.styles.get(index).unwrap_or(&default);
    let mut fold = own.clone();
    for priority in lit.iter().rev() {
        let rule = page
            .rules()
            .iter()
            .find(|r| r.priority == *priority)
            .expect("a rule");
        if let Some(style) = rule.format.and_then(|f| wb.styles.get(f)) {
            fold = style.over(&fold);
        }
    }
    fold.num_fmt = own.num_fmt.clone();
    let line = record_line(&Record::Format {
        index: 0,
        style: Some(&fold),
        code: None,
    });
    let body = line
        .strip_prefix("(format 0 ")
        .and_then(|l| l.strip_suffix(')'))
        .expect("a format's row");
    let mut d = drawn(body, true);
    // A word that says a default draws as nothing said.
    d.font.retain(|w| !w.starts_with("not-"));
    d.edges.retain(|e| !e.ends_with(" none"));
    d
}

/// Excel's reading of a cell's look, what the record has no word for left out.
fn read_look(reading: &str) -> Drawn {
    drawn(reading.split(" !").next().expect("a reading"), true)
}

/// The text a cell shows in its record: its `text` row, or its value as a
/// `cell` or `value` row spells it.
fn shown_text(record: &str) -> String {
    for head in ["(text ", "(value ", "(cell "] {
        if let Some(line) = record.lines().find(|l| l.starts_with(head)) {
            let forms = vla_lang::reader::read_forms(line).expect("a row");
            let items = &forms[0].as_list().expect("a row is a list").items;
            return match &items[3] {
                Form::Str(s) | Form::Sym(s) => format!("\"{s}\""),
                other => panic!("{other:?}"),
            };
        }
    }
    "\"\"".to_string()
}

/// Every place the language draws otherwise than Excel's reading, by its file,
/// sheet, case and target, and why: a gap the scoping named, a word the record
/// has none for, a rule a reader rewrites, and what `DisplayFormat` reads and
/// Excel's screen does not draw (the owner's look, `Partial` and `Order`).
const GAPS: &[(&str, &str, &str, &[&str], &str)] = &[
    // A range read as an array, `KERNEL.55`'s.
    (
        "rules.xlsx",
        "Truth",
        "a range whose first answer is true",
        &["shown:C60"],
        "Excel answers by a range's first value; a cell's reading intersects it",
    ),
    (
        "rules.xlsx",
        "Refs",
        "a range under AND",
        &["shown:C10"],
        "Excel reads the comparison as an array; a cell's reading intersects it",
    ),
    // A merge, the item's sixth slice.
    (
        "rules.xlsx",
        "Refs",
        "a merged range whose first cell answers false",
        &["shown:D22", "shown:E22"],
        "a merged range draws by its first cell's answer",
    ),
    (
        "rules.xlsx",
        "Refs",
        "a merged range whose first cell answers true",
        &["shown:D24", "shown:E24"],
        "a merged range draws by its first cell's answer",
    ),
    // A reference moved past the sheet's edge.
    (
        "rules.xlsx",
        "Edges",
        "a reference moved past the bottom",
        &["shown:C3"],
        "Excel wraps it to the top; a fill's mover makes it #REF!, false",
    ),
    (
        "rules.xlsx",
        "Edges",
        "a reference moved past the right",
        &["shown:E6"],
        "Excel wraps it to the left; a fill's mover makes it #REF!, false",
    ),
    // What DisplayFormat reads and Excel's screen does not draw.
    (
        "rules.xlsx",
        "Order",
        "a stop in the cell two ranges meet in",
        &["shown:C34"],
        "the screen draws no bold: a stop ends every later rule, whatever its range",
    ),
    (
        "rules.xlsx",
        "Order",
        "a stop over a wider range",
        &["shown:C38"],
        "the screen draws no bold: a stop ends every later rule, whatever its range",
    ),
    (
        "rules.xlsx",
        "Order",
        "a stop under a wider range",
        &["shown:C40"],
        "the screen draws no bold: a stop ends every later rule, whatever its range",
    ),
    (
        "rules.xlsx",
        "Order",
        "a stop of no format over a wider range",
        &["shown:C48"],
        "the screen draws no fill: a stop ends every later rule, whatever its range",
    ),
    (
        "rules.xlsx",
        "Partial",
        "a name",
        &["shown:C64", "shown:E64"],
        "the screen draws the cell's own face: a reader leaves a rule's name out",
    ),
    (
        "rules.xlsx",
        "Partial",
        "a vertical alignment",
        &["shown:C70", "shown:E70"],
        "the screen draws it where it was: a reader leaves a rule's vertical out",
    ),
    (
        "rules.xlsx",
        "Partial",
        "a horizontal alignment",
        &["shown:C68", "shown:E68"],
        "drawn centered; the top DisplayFormat adds is not drawn",
    ),
    (
        "rules.xlsx",
        "Partial",
        "a wrap",
        &["shown:C72", "shown:E72"],
        "the top DisplayFormat adds is not drawn",
    ),
    // What the record has no word for.
    (
        "rules.xlsx",
        "Partial",
        "a fill's fgColor under solid",
        &["shown:E8"],
        "a pattern said with no color draws white over no fill and nothing over a fill",
    ),
    (
        "rules.xlsx",
        "Partial",
        "a fill's pattern",
        &["shown:C16"],
        "a pattern, which a format has no word for, as a cell's",
    ),
    (
        "rules.xlsx",
        "Partial",
        "a font's automatic color",
        &["shown:C34"],
        "the automatic color said over a cell's own, which a format has no word for",
    ),
    (
        "rules.xlsx",
        "Partial",
        "a number format with a color",
        &["shown:C60", "text:C60", "shown:E60"],
        "a code's color, which a format's code refuses by name",
    ),
    // The kinds the language does not compute, the item's last slice.
    (
        "rules.xlsx",
        "Kinds",
        "a color scale",
        &["shown:C20", "shown:E20", "shown:G20"],
        "a color scale computes a color no row says",
    ),
    (
        "rules.xlsx",
        "Kinds",
        "the top value",
        &["shown:G26"],
        "not computed",
    ),
    (
        "rules.xlsx",
        "Kinds",
        "above the average",
        &["shown:G28"],
        "not computed",
    ),
    (
        "rules.xlsx",
        "Kinds",
        "a duplicate",
        &["shown:C30", "shown:E30"],
        "not computed",
    ),
    (
        "rules.xlsx",
        "Kinds",
        "a value alone",
        &["shown:G32"],
        "not computed",
    ),
];

fn gap(file: &str, sheet: &str, case: &str, target: &str) -> Option<usize> {
    GAPS.iter().position(|(f, s, c, targets, _)| {
        *f == file && *s == sheet && *c == case && targets.contains(&target)
    })
}

#[test]
fn the_machine_draws_each_case_as_excel_draws_it() {
    let golden = golden_read();
    let mut seen = vec![false; GAPS.len()];
    let mut wrong: Vec<String> = Vec::new();
    let mut held = 0;
    for (_, file, book) in books() {
        let bytes = book.bytes();
        let (cells, _) = crate::reflect::reflect_looks(&bytes, file)
            .unwrap_or_else(|r| panic!("{file}: {}", r.text));
        let mut rows = cells;
        for (k, (_, words)) in format_words().iter().enumerate() {
            if !words.is_empty() {
                rows.push_str(&format!("(format {} {words})\n", 1000 + k));
            }
        }
        for part in &book.parts {
            rows.extend(part.rule_rows.iter().cloned());
        }
        let mut m = Machine::load(&rows, crate::excel::library())
            .unwrap_or_else(|e| panic!("{file}: line {}: {}\n{rows}", e.line, e.refusal.text));
        assert!(m.step(0).done);
        let wb = m.workbook();
        for case in golden.iter().filter(|c| c.file == file) {
            for (target, reading) in &case.targets {
                let (kind, place) = target.split_once(':').expect("a kind and a place");
                if kind == "rule" {
                    continue;
                }
                let record = one_cell(&m, &case.sheet, place);
                let at = parse_a1_range(place).expect("a cell");
                let agrees = match kind {
                    "shown" => {
                        let lit = lit_at(&record);
                        folded(&wb, &case.sheet, (at.top, at.left), &lit) == read_look(reading)
                    }
                    "text" => shown_text(&record) == *reading,
                    other => panic!("a target of no kind: {other}"),
                };
                held += 1;
                match (agrees, gap(file, &case.sheet, &case.name, target)) {
                    (true, None) => {}
                    (true, Some(k)) => {
                        wrong.push(format!(
                            "{file}!{}!{}: {target} agrees, and is named a gap: {}",
                            case.sheet, case.name, GAPS[k].4
                        ));
                    }
                    (false, Some(k)) => seen[k] = true,
                    (false, None) => {
                        let ours: Vec<&str> = record
                            .lines()
                            .filter(|l| {
                                l.starts_with("(lit ")
                                    || l.starts_with("(text ")
                                    || l.starts_with("(style ")
                            })
                            .collect();
                        wrong.push(format!(
                            "{file}!{}!{}: {target}: Excel reads {reading}; the machine's rows {ours:?}",
                            case.sheet, case.name
                        ));
                    }
                }
            }
        }
    }
    for (k, (f, s, c, _, why)) in GAPS.iter().enumerate() {
        if !seen[k] {
            wrong.push(format!("{f}!{s}!{c}: a gap no target met: {why}"));
        }
    }
    assert!(
        wrong.is_empty(),
        "{} of {held} targets:\n{}",
        wrong.len(),
        wrong.join("\n")
    );
}

#[test]
fn the_rules_golden_reads_every_case_of_the_builder() {
    let mut want: Vec<(String, Vec<String>)> = Vec::new();
    for (_, file, book) in books() {
        for part in &book.parts {
            for (name, reads) in &part.cases {
                want.push((
                    format!("{file}!{}!{name}", part.name),
                    reads.split(' ').map(str::to_string).collect(),
                ));
            }
        }
    }
    let got = golden_cases();
    let names: BTreeSet<&String> = want.iter().map(|(name, _)| name).collect();
    assert_eq!(names.len(), want.len(), "two cases of one name");
    assert_eq!(
        got.len(),
        want.len(),
        "the golden's cases, against the builder's"
    );
    for ((got_name, got_targets), (want_name, want_targets)) in got.iter().zip(&want) {
        assert_eq!(got_name, want_name);
        assert_eq!(got_targets, want_targets, "{got_name}");
    }
}
