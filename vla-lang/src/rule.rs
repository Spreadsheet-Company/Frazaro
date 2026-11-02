//! A conditional format (KERNEL.58): a rule as a sheet holds it, its row
//! `(conditional "<sheet>" <priority> "<range>" <test> <format>|none [stop])`,
//! and the formula that answers it.
//!
//! A rule's test is a formula, or Excel's `cellIs`, a cell's value against one
//! bound or two, `(cell greater "=5")`. Both move from the corner of the
//! rule's range, the top left of the rectangle around all its rectangles,
//! whatever order they are written in and whether or not a rectangle holds that
//! cell, as Excel moved them in every case of its two saves
//! (`scripts/reflect/rules_golden.txt`).
//!
//! A machine computes a rule as a fill of [`Rule::answer_text`] on an answer
//! sheet of the rule's own, at the places of the cells it answers for
//! (`crate::machine`). The text makes an answer 1 or 0, an error and a text
//! false and the text `TRUE` true, as `IF` reads each, which is what Excel
//! answered; it qualifies every reference that names no sheet by the rule's,
//! a defined name found from there, so that a rule reads its own sheet from a
//! sheet of its own. What it leaves as Excel's alone is named in the core's
//! test of that golden: a reference moved past the bottom or the right of the
//! sheet, which Excel wraps and a fill's mover makes `#REF!`, and a range read
//! as an array.
//!
//! Which rules draw at a cell is the record's `lit` row: those whose answer is
//! true, in priority order, up to the first that stops. A stop ends every later
//! rule over the cell, as Excel draws one and as its object model's
//! `DisplayFormat` does not say across two ranges (the owner's look at the
//! `Order` sheet).

use crate::refers;
use crate::sheet::{cell_ref, A1Range};

/// Excel's eight comparisons of a cell's value, in the row's words.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Operator {
    Less,
    LessOrEqual,
    Equal,
    NotEqual,
    GreaterOrEqual,
    Greater,
    Between,
    NotBetween,
}

impl Operator {
    const WORDS: [(Operator, &'static str); 8] = [
        (Operator::Less, "less"),
        (Operator::LessOrEqual, "less-or-equal"),
        (Operator::Equal, "equal"),
        (Operator::NotEqual, "not-equal"),
        (Operator::GreaterOrEqual, "greater-or-equal"),
        (Operator::Greater, "greater"),
        (Operator::Between, "between"),
        (Operator::NotBetween, "not-between"),
    ];

    /// The word a row spells it by.
    pub fn word(self) -> &'static str {
        Self::WORDS
            .iter()
            .find(|(o, _)| *o == self)
            .map_or("", |(_, w)| w)
    }

    /// The operator a row's word spells.
    pub fn from_word(word: &str) -> Option<Operator> {
        Self::WORDS
            .iter()
            .find(|(_, w)| *w == word)
            .map(|(o, _)| *o)
    }

    /// Its name in a file, a `cfRule`'s `operator`.
    pub fn excel(self) -> &'static str {
        match self {
            Operator::Less => "lessThan",
            Operator::LessOrEqual => "lessThanOrEqual",
            Operator::Equal => "equal",
            Operator::NotEqual => "notEqual",
            Operator::GreaterOrEqual => "greaterThanOrEqual",
            Operator::Greater => "greaterThan",
            Operator::Between => "between",
            Operator::NotBetween => "notBetween",
        }
    }

    /// The operator a file's name spells; `None` for the names of no
    /// comparison, `containsText` and its kin, which hold a formula of their own.
    pub fn from_excel(name: &str) -> Option<Operator> {
        Self::WORDS
            .iter()
            .map(|(o, _)| *o)
            .find(|o| o.excel() == name)
    }

    /// Whether it takes two bounds.
    pub fn takes_two(self) -> bool {
        matches!(self, Operator::Between | Operator::NotBetween)
    }
}

/// What a rule tests: each formula's text without its `=`, as a `formula` row
/// holds one.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum Test {
    /// A formula, read for its truth at the range's corner.
    Formula(String),
    /// A cell's value against its bound, or against both of an operator that
    /// takes two.
    Cell {
        operator: Operator,
        first: String,
        second: Option<String>,
    },
}

/// One conditional format, as a sheet holds it and its row spells it.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Rule {
    /// Its place among its sheet's rules, the lowest first.
    pub priority: u32,
    /// The cells it answers for, as a file's `sqref` names them: a rectangle
    /// or several, in the order written, at least one.
    pub ranges: Vec<A1Range>,
    pub test: Test,
    /// The format it draws, as a `style` row names one; `None` for a rule
    /// that draws nothing, whose purpose can only be to stop.
    pub format: Option<u32>,
    /// Whether, true, it ends the rules after it over the cell.
    pub stop: bool,
}

impl Rule {
    /// The range as a row spells it: `A1:A3 C1:C3`.
    pub fn range_text(&self) -> String {
        let spelled: Vec<String> = self.ranges.iter().map(A1Range::text).collect();
        spelled.join(" ")
    }

    /// The top left of the rectangle around its rectangles, which its formulas
    /// are read at and move from.
    pub fn corner(&self) -> (u32, u32) {
        let top = self.ranges.iter().map(|r| r.top).min().unwrap_or(1);
        let left = self.ranges.iter().map(|r| r.left).min().unwrap_or(1);
        (top, left)
    }

    /// The rectangle around its rectangles.
    pub fn around(&self) -> A1Range {
        let (top, left) = self.corner();
        A1Range {
            top,
            left,
            bottom: self.ranges.iter().map(|r| r.bottom).max().unwrap_or(top),
            right: self.ranges.iter().map(|r| r.right).max().unwrap_or(left),
        }
    }

    /// Whether it answers for a cell.
    pub fn covers(&self, row: u32, col: u32) -> bool {
        self.ranges
            .iter()
            .any(|r| row >= r.top && row <= r.bottom && col >= r.left && col <= r.right)
    }

    /// How many cells it answers for, a cell two rectangles share once: what a
    /// load counts of it beside the grid's cells.
    pub fn cells(&self) -> u64 {
        let mut before: Vec<A1Range> = Vec::new();
        let mut cells = 0;
        for &range in &self.ranges {
            cells += crate::grid::outside(range, &before)
                .iter()
                .map(A1Range::cells)
                .sum::<u64>();
            before.push(range);
        }
        cells
    }

    /// Its formulas' texts, each without its `=`.
    pub fn formulas(&self) -> Vec<&str> {
        match &self.test {
            Test::Formula(text) => vec![text],
            Test::Cell { first, second, .. } => {
                let mut texts = vec![first.as_str()];
                texts.extend(second.as_deref());
                texts
            }
        }
    }

    /// The formula an answer sheet holds for it, without its `=`, read at
    /// [`Rule::corner`]: 1 where the rule is true there and 0 elsewhere,
    /// reading the rule's sheet, `sheet`.
    pub(crate) fn answer_text(&self, sheet: &str) -> String {
        let test = match &self.test {
            Test::Formula(text) => refers::qualified(text, sheet),
            Test::Cell {
                operator,
                first,
                second,
            } => {
                let (row, col) = self.corner();
                let cell = refers::qualified(&cell_ref(row, col), sheet);
                let first = bound(first, sheet);
                let second = second.as_deref().map(|b| bound(b, sheet));
                compared(*operator, &cell, &first, second.as_deref())
            }
        };
        format!("IFERROR(IF({test},1,0),0)")
    }
}

/// A `cellIs` bound as its comparison reads it: Excel computes a bound before
/// it compares, so a reference to a cell holding nothing is the number 0 under
/// every operator, whatever the cell compared holds, where a comparison reading
/// the reference itself takes it as `FALSE` beside a truth value and as the
/// empty text beside a text. A bound reading nothing cannot be that, and is
/// its own text in parentheses, which may hold any operator.
fn bound(text: &str, sheet: &str) -> String {
    if refers::scan(text).is_empty() {
        return format!("({text})");
    }
    let read = refers::qualified(text, sheet);
    format!("IF(COUNTA({read})=0,0,{read})")
}

/// The test of a cell against its bounds. Excel takes a `between`'s two bounds
/// in either order.
fn compared(operator: Operator, cell: &str, first: &str, second: Option<&str>) -> String {
    let inside = |second: &str| {
        format!(
            "IF({first}<={second},AND({cell}>={first},{cell}<={second}),AND({cell}>={second},{cell}<={first}))"
        )
    };
    match (operator, second) {
        (Operator::Less, _) => format!("{cell}<{first}"),
        (Operator::LessOrEqual, _) => format!("{cell}<={first}"),
        (Operator::Equal, _) => format!("{cell}={first}"),
        (Operator::NotEqual, _) => format!("{cell}<>{first}"),
        (Operator::GreaterOrEqual, _) => format!("{cell}>={first}"),
        (Operator::Greater, _) => format!("{cell}>{first}"),
        (Operator::Between, Some(second)) => inside(second),
        (Operator::NotBetween, Some(second)) => format!("NOT({})", inside(second)),
        // A row holds an operator to its bounds, so this is a reader's own
        // rule with a bound missing: nothing lights it.
        (Operator::Between | Operator::NotBetween, None) => "FALSE".to_string(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::sheet::parse_a1_range;

    fn over(ranges: &[&str], test: Test) -> Rule {
        Rule {
            priority: 1,
            ranges: ranges.iter().map(|r| parse_a1_range(r).unwrap()).collect(),
            test,
            format: Some(64),
            stop: false,
        }
    }

    fn cell(operator: Operator, first: &str, second: Option<&str>) -> Test {
        Test::Cell {
            operator,
            first: first.to_string(),
            second: second.map(str::to_string),
        }
    }

    #[test]
    fn a_rule_s_corner_is_the_top_left_of_its_rectangles_in_any_order() {
        let rule = over(&["D12:D13", "F10:F11"], Test::Formula("1".to_string()));
        assert_eq!(rule.corner(), (10, 4), "no cell of the range");
        assert_eq!(rule.around().text(), "D10:F13");
        assert_eq!(rule.range_text(), "D12:D13 F10:F11");
        assert!(rule.covers(12, 4) && rule.covers(11, 6));
        assert!(!rule.covers(10, 4), "the corner answers for nothing");
        assert_eq!(rule.cells(), 4);
        let rule = over(&["A1:B2", "B2:C3"], Test::Formula("1".to_string()));
        assert_eq!(rule.cells(), 7, "a cell two rectangles share is one");
    }

    #[test]
    fn a_rule_s_answer_reads_its_own_sheet_and_answers_1_or_0() {
        let text = over(&["B2:J10"], Test::Formula("$A2=1".to_string())).answer_text("Board");
        assert_eq!(text, "IFERROR(IF(Board!$A2=1,1,0),0)");
        // A sheet it names stands; a name is found from the rule's sheet; a
        // function, a truth value and a text are no reference; a quoted name.
        let test = "AND(Other!A1=TRUE,pick=\"A1\",SUM(B1:B3)>0)".to_string();
        let text = over(&["B2"], Test::Formula(test)).answer_text("My Sheet");
        assert_eq!(
            text,
            "IFERROR(IF(AND(Other!A1=TRUE,'My Sheet'!pick=\"A1\",SUM('My Sheet'!B1:B3)>0),1,0),0)"
        );
    }

    #[test]
    fn a_bound_reading_a_cell_reads_nothing_as_the_number_0() {
        let text = over(&["C2:E2"], cell(Operator::Equal, "$G2", None)).answer_text("S");
        assert_eq!(text, "IFERROR(IF(S!C2=IF(COUNTA(S!$G2)=0,0,S!$G2),1,0),0)");
        let text = over(&["C2:E2"], cell(Operator::Less, "5", None)).answer_text("S");
        assert_eq!(text, "IFERROR(IF(S!C2<(5),1,0),0)");
    }

    #[test]
    fn a_between_takes_its_bounds_in_either_order() {
        let text = over(&["C2"], cell(Operator::Between, "5", Some("9"))).answer_text("S");
        assert_eq!(
            text,
            "IFERROR(IF(IF((5)<=(9),AND(S!C2>=(5),S!C2<=(9)),AND(S!C2>=(9),S!C2<=(5))),1,0),0)"
        );
        let text = over(&["C2"], cell(Operator::NotBetween, "5", Some("9"))).answer_text("S");
        assert!(text.starts_with("IFERROR(IF(NOT(IF((5)<=(9),"), "{text}");
    }

    #[test]
    fn each_operator_is_its_word() {
        for (operator, word) in Operator::WORDS {
            assert_eq!(operator.word(), word);
            assert_eq!(Operator::from_word(word), Some(operator));
            assert_eq!(Operator::from_excel(operator.excel()), Some(operator));
        }
        assert_eq!(Operator::from_word("greater-than"), None);
        assert_eq!(Operator::from_excel("greater"), None);
        assert_eq!(Operator::from_excel("containsText"), None);
        assert!(Operator::Between.takes_two() && !Operator::Equal.takes_two());
    }
}
