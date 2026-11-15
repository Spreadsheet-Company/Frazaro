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
//!
//! A load refuses a rule as it refuses a cell's formula, for a function its
//! library lacks or a text its parse does not read. [`Rule::computable`] asks
//! the same before a load does, for a reader of a file, which leaves such a rule
//! out so that one rule does not keep a whole workbook out.
//!
//! The kinds a file holds that say no formula of their own, a color scale, a data
//! bar, a set of icons, a rank, an average, a duplicate and a value alone, are
//! [`Uncomputed`]: a row of words a machine plans nothing for and draws nothing
//! by, printed back and written as it was read, so that a workbook keeps them.

use crate::calc::eval::{unread_name, Library};
use crate::calc::formula;
use crate::calc::shape::moves_unread;
use crate::calc::Names;
use crate::refers;
use crate::rows::quoted;
use crate::sheet::{cell_ref, strip_future_prefixes, A1Range, Rgb, Workbook};

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

/// What a point of a color scale, a data bar or a set of icons is measured by.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum StopKind {
    Min,
    Max,
    Number,
    Percent,
    Percentile,
    Formula,
}

impl StopKind {
    const WORDS: [(StopKind, &'static str, &'static str); 6] = [
        (StopKind::Min, "min", "min"),
        (StopKind::Max, "max", "max"),
        (StopKind::Number, "number", "num"),
        (StopKind::Percent, "percent", "percent"),
        (StopKind::Percentile, "percentile", "percentile"),
        (StopKind::Formula, "formula", "formula"),
    ];

    /// As a row spells it.
    pub fn word(self) -> &'static str {
        Self::WORDS[self as usize].1
    }

    /// From a row's word.
    pub fn from_word(word: &str) -> Option<StopKind> {
        Self::WORDS.iter().find(|w| w.1 == word).map(|w| w.0)
    }

    /// As a file's `cfvo` names it.
    pub fn excel(self) -> &'static str {
        Self::WORDS[self as usize].2
    }

    /// From a file's `cfvo`.
    pub fn from_excel(name: &str) -> Option<StopKind> {
        Self::WORDS.iter().find(|w| w.2 == name).map(|w| w.0)
    }

    /// Whether it says a value: all but the lowest and the highest.
    pub fn says_a_value(self) -> bool {
        !matches!(self, StopKind::Min | StopKind::Max)
    }
}

/// A point of a color scale, a data bar or a set of icons, a file's `cfvo`, as a
/// row spells it: `(min)`, `(percent 33)`, `(formula "=$A$1")`.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Stop {
    pub kind: StopKind,
    /// A number as it was written, or a formula's text without its `=`; empty
    /// where its kind says none.
    pub value: String,
    /// Whether, at an icon's, a value must pass it where it may otherwise equal it.
    pub greater: bool,
}

impl Stop {
    /// Its value's formula: its kind's own, or a number's place held by one.
    pub fn formula(&self) -> Option<&str> {
        let numbered = self.kind != StopKind::Formula && is_number(&self.value);
        (self.kind.says_a_value() && !numbered).then_some(self.value.as_str())
    }

    fn words(&self) -> String {
        let mut s = format!("({}", self.kind.word());
        if self.kind.says_a_value() {
            s.push(' ');
            match self.formula() {
                Some(text) => s.push_str(&quoted(&format!("={text}"))),
                None => s.push_str(&self.value),
            }
        }
        if self.greater {
            s.push_str(" greater");
        }
        s.push(')');
        s
    }
}

/// Whether a text is a number a stop may say as it stands.
pub fn is_number(text: &str) -> bool {
    text.parse::<f64>().is_ok_and(f64::is_finite)
}

/// The sets of icons a file's own `iconSet` names, each by its count of icons first.
pub const ICON_SETS: [&str; 17] = [
    "3Arrows",
    "3ArrowsGray",
    "3Flags",
    "3TrafficLights1",
    "3TrafficLights2",
    "3Signs",
    "3Symbols",
    "3Symbols2",
    "4Arrows",
    "4ArrowsGray",
    "4RedToBlack",
    "4Rating",
    "4TrafficLights",
    "5Arrows",
    "5ArrowsGray",
    "5Rating",
    "5Quarters",
];

/// A data bar's shortest and longest as percents of its cell, when a file says neither.
pub const BAR_LENGTHS: (u32, u32) = (10, 90);

/// A kind of rule the language does not compute: a machine plans nothing for it and
/// draws nothing by it, and it is held so that a workbook keeps it. What each draws
/// is computed for each cell or over its whole range, which no `lit` row says.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum Uncomputed {
    /// Two stops or three, and a color for each.
    ColorScale { stops: Vec<Stop>, colors: Vec<Rgb> },
    /// Its lowest and its highest, its color, its shortest and longest as percents
    /// of its cell, and whether the cell's value is drawn beside it.
    DataBar {
        low: Stop,
        high: Stop,
        color: Rgb,
        lengths: (u32, u32),
        shows_value: bool,
    },
    /// One of [`ICON_SETS`], a stop for each icon, whether they run the other way,
    /// and whether the cell's value is drawn beside them.
    IconSet {
        set: String,
        stops: Vec<Stop>,
        reverse: bool,
        shows_value: bool,
    },
    /// The highest values, or the lowest, by count or by percent.
    Rank {
        bottom: bool,
        rank: u32,
        percent: bool,
    },
    /// Above the range's average or below it, or equal to it too, or by a count of
    /// standard deviations.
    Average {
        below: bool,
        equal: bool,
        deviations: u32,
    },
    /// A value more than one cell of the range holds.
    Duplicate,
    /// A value one cell of the range holds.
    Unique,
}

impl Uncomputed {
    /// Its stops.
    pub fn stops(&self) -> Vec<&Stop> {
        match self {
            Uncomputed::ColorScale { stops, .. } | Uncomputed::IconSet { stops, .. } => {
                stops.iter().collect()
            }
            Uncomputed::DataBar { low, high, .. } => vec![low, high],
            _ => Vec::new(),
        }
    }

    /// Whether it draws by a format and may stop a rule after it. A scale, a bar
    /// and a set of icons draw by their own words and neither.
    pub fn takes_a_format(&self) -> bool {
        !matches!(
            self,
            Uncomputed::ColorScale { .. } | Uncomputed::DataBar { .. } | Uncomputed::IconSet { .. }
        )
    }

    /// The `type` of a file's element.
    pub fn excel(&self) -> &'static str {
        match self {
            Uncomputed::ColorScale { .. } => "colorScale",
            Uncomputed::DataBar { .. } => "dataBar",
            Uncomputed::IconSet { .. } => "iconSet",
            Uncomputed::Rank { .. } => "top10",
            Uncomputed::Average { .. } => "aboveAverage",
            Uncomputed::Duplicate => "duplicateValues",
            Uncomputed::Unique => "uniqueValues",
        }
    }

    /// Whether it says only what a file can: what Excel's dialogs let a rule say,
    /// `greater` at an icon's alone, and each stop's formula one the parse reads, though
    /// nothing runs it, so that a function no library holds is no matter. A load refuses
    /// the rest, and a reader leaves it out.
    pub fn well_said(&self) -> bool {
        let stops = self.stops();
        let greater =
            matches!(self, Uncomputed::IconSet { .. }) || stops.iter().all(|stop| !stop.greater);
        let own = match self {
            Uncomputed::ColorScale { stops, colors } => {
                (2..=3).contains(&stops.len()) && colors.len() == stops.len()
            }
            Uncomputed::DataBar { lengths, .. } => lengths.0 <= lengths.1 && lengths.1 <= 100,
            Uncomputed::IconSet { set, stops, .. } => {
                ICON_SETS.contains(&set.as_str())
                    && set
                        .chars()
                        .next()
                        .and_then(|icons| icons.to_digit(10))
                        .is_some_and(|icons| icons as usize == stops.len())
            }
            Uncomputed::Rank { rank, percent, .. } => {
                let most = if *percent { 100 } else { 1000 };
                (1..=most).contains(rank)
            }
            Uncomputed::Average {
                equal, deviations, ..
            } => *deviations <= 3 && !(*equal && *deviations > 0),
            Uncomputed::Duplicate | Uncomputed::Unique => true,
        };
        let formulas = stops
            .iter()
            .filter_map(|stop| stop.formula())
            .all(|text| formula::parse(text).is_ok());
        greater && own && formulas
    }

    fn words(&self) -> String {
        let stops = |stops: &[&Stop]| -> String {
            stops
                .iter()
                .map(|stop| format!(" {}", stop.words()))
                .collect()
        };
        let color = |color: &Rgb| format!(" {}", quoted(&color.hex()));
        match self {
            Uncomputed::ColorScale { colors, .. } => {
                let colors: String = colors.iter().map(color).collect();
                format!("(color-scale{}{colors})", stops(&self.stops()))
            }
            Uncomputed::DataBar {
                color: bar,
                lengths,
                shows_value,
                ..
            } => {
                let length = if *lengths == BAR_LENGTHS {
                    String::new()
                } else {
                    format!(" (length {} {})", lengths.0, lengths.1)
                };
                let hidden = if *shows_value { "" } else { " no-value" };
                format!(
                    "(data-bar{}{}{length}{hidden})",
                    stops(&self.stops()),
                    color(bar)
                )
            }
            Uncomputed::IconSet {
                set,
                reverse,
                shows_value,
                ..
            } => format!(
                "(icon-set {}{}{}{})",
                quoted(set),
                stops(&self.stops()),
                if *reverse { " reverse" } else { "" },
                if *shows_value { "" } else { " no-value" }
            ),
            Uncomputed::Rank {
                bottom,
                rank,
                percent,
            } => format!(
                "({} {rank}{})",
                if *bottom { "bottom" } else { "top" },
                if *percent { " percent" } else { "" }
            ),
            Uncomputed::Average {
                below,
                equal,
                deviations,
            } => {
                let by = if *deviations > 0 {
                    format!(" (deviations {deviations})")
                } else {
                    String::new()
                };
                format!(
                    "({}-average{}{by})",
                    if *below { "below" } else { "above" },
                    if *equal { " equal" } else { "" }
                )
            }
            Uncomputed::Duplicate => "(duplicate)".to_string(),
            Uncomputed::Unique => "(unique)".to_string(),
        }
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
    /// A kind the language does not compute.
    Uncomputed(Uncomputed),
}

impl Test {
    /// As its row spells it.
    pub fn words(&self) -> String {
        match self {
            Test::Formula(text) => quoted(&format!("={text}")),
            Test::Cell {
                operator,
                first,
                second,
            } => {
                let second = second
                    .as_ref()
                    .map_or(String::new(), |b| format!(" {}", quoted(&format!("={b}"))));
                format!(
                    "(cell {} {}{second})",
                    operator.word(),
                    quoted(&format!("={first}"))
                )
            }
            Test::Uncomputed(held) => held.words(),
        }
    }
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
        #[cfg(test)]
        crate::sheet::work::rectangles(self.ranges.len());
        let top = self.ranges.iter().map(|r| r.top).min().unwrap_or(1);
        let left = self.ranges.iter().map(|r| r.left).min().unwrap_or(1);
        (top, left)
    }

    /// The rectangle around its rectangles.
    pub fn around(&self) -> A1Range {
        #[cfg(test)]
        crate::sheet::work::rectangles(self.ranges.len());
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
        #[cfg(test)]
        crate::sheet::work::rectangles(self.ranges.len());
        self.ranges
            .iter()
            .any(|r| row >= r.top && row <= r.bottom && col >= r.left && col <= r.right)
    }

    /// How many cells it answers for, a cell two rectangles share once: what a
    /// load counts of it beside the grid's cells.
    pub fn cells(&self) -> u64 {
        crate::grid::union_cells(&self.ranges)
    }

    /// Its formulas' texts, each without its `=`: a kind not computed's its stops'.
    pub fn formulas(&self) -> Vec<&str> {
        match &self.test {
            Test::Formula(text) => vec![text],
            Test::Cell { first, second, .. } => {
                let mut texts = vec![first.as_str()];
                texts.extend(second.as_deref());
                texts
            }
            Test::Uncomputed(held) => held.stops().into_iter().filter_map(Stop::formula).collect(),
        }
    }

    /// Whether a machine answers it: not one of the [`Uncomputed`] kinds.
    pub fn computed(&self) -> bool {
        !matches!(self.test, Test::Uncomputed(_))
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
            // No machine plans it.
            Test::Uncomputed(_) => "0".to_string(),
        };
        format!("IFERROR(IF({test},1,0),0)")
    }

    /// Whether a load computes it on the sheet at `home` of `model`, which names
    /// a workbook's sheets and defined names (`names`): what a load asks of its
    /// answer, that `library` holds each function it calls, that the parse reads
    /// its text and each name it reads, and, of a reference a move may spoil, its
    /// text at the far corner of each rectangle, a cell there moved furthest. Of a kind
    /// it does not compute, whether a load takes it: [`Uncomputed::well_said`].
    pub fn computable(
        &self,
        model: &Workbook,
        names: &Names,
        home: usize,
        library: &Library,
    ) -> bool {
        let Some(sheet) = model.sheets.get(home) else {
            return false;
        };
        if let Test::Uncomputed(held) = &self.test {
            return held.well_said();
        }
        let text = strip_future_prefixes(&self.answer_text(&sheet.name));
        if formula::calls(&text).iter().any(|name| !library.has(name)) {
            return false;
        }
        let Ok(parsed) = formula::parse(&text) else {
            return false;
        };
        if unread_name(model, names, &parsed, home).is_some() {
            return false;
        }
        if !moves_unread(&text) {
            return true;
        }
        let (top, left) = self.corner();
        self.ranges.iter().all(|range| {
            let (d_row, d_col) = (range.bottom - top, range.right - left);
            let moved = refers::shift_a1_references(&text, i64::from(d_row), i64::from(d_col));
            formula::parse(&moved).is_ok()
        })
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

    /// A kind not computed is computable where it is well said, on a sheet the
    /// workbook holds, a stop calling what no library holds among it.
    #[test]
    fn a_kind_not_computed_is_computable_where_it_is_well_said() {
        let mut wb = Workbook::new();
        wb.sheets.push(crate::sheet::Sheet::new("S"));
        let names = Names::of(&wb);
        let library = crate::calc::library::language();
        let stop = |kind, value: &str| Stop {
            kind,
            value: value.to_string(),
            greater: false,
        };
        let scale = |first: Stop| {
            over(
                &["A1"],
                Test::Uncomputed(Uncomputed::ColorScale {
                    stops: vec![first, stop(StopKind::Max, "")],
                    colors: vec![Rgb(1, 2, 3), Rgb(4, 5, 6)],
                }),
            )
        };
        let well = scale(stop(StopKind::Formula, "NOSUCH(A1)"));
        assert!(well.computable(&wb, &names, 0, &library));
        assert!(!well.computable(&wb, &names, 1, &library));
        assert!(!scale(stop(StopKind::Formula, "1+")).computable(&wb, &names, 0, &library));
        assert!(!Uncomputed::Average {
            below: false,
            equal: true,
            deviations: 1,
        }
        .well_said());
    }

    /// A stop's kinds, in a row's words and a file's.
    #[test]
    fn a_stop_s_kinds_are_a_row_s_words_and_a_file_s() {
        let kinds = [
            (StopKind::Min, "min", "min"),
            (StopKind::Max, "max", "max"),
            (StopKind::Number, "number", "num"),
            (StopKind::Percent, "percent", "percent"),
            (StopKind::Percentile, "percentile", "percentile"),
            (StopKind::Formula, "formula", "formula"),
        ];
        for (kind, word, excel) in kinds {
            assert_eq!((kind.word(), kind.excel()), (word, excel));
            assert_eq!(StopKind::from_word(word), Some(kind));
            assert_eq!(StopKind::from_excel(excel), Some(kind));
        }
        assert_eq!(StopKind::from_excel("number"), None);
        assert_eq!(StopKind::from_word("num"), None);
    }

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
