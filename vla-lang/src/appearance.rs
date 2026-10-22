//! A format's appearance (KERNEL.50, 2026-10-21): what a cell's format says
//! of its text's font, of its alignment and of its borders, beside the fill,
//! the number format and the wrap a format held before
//! ([`crate::sheet::Style`]). Each is Excel's own, the subset its first
//! consumers asked for (Alonzo's `CART.14`, Buffet's buttons and glyphs): a
//! weight, a slant, a size and a color; a horizontal and a vertical
//! alignment; and a cell's four edges, each a line of Excel's fourteen and a
//! color.
//!
//! Each property is said or not, `None` where not: a format says what its
//! row spells. A look draws a cell with what its format says and, for each
//! property of a font, an alignment or a border it does not say, the cell's
//! own ([`crate::sheet::Style::over`]); its fill, number format and wrap,
//! which every format row spells, stand in place of the cell's own as they
//! did. A property said is said whatever its value, so that `not-bold`,
//! `general`, `bottom` and an edge's `none` take a cell's own back under a
//! look; as a cell's own format, a said default draws as the default.
//!
//! The record spells them as groups after a format's wrap, each group and
//! each property at most once, printed in the order below and read in any:
//!
//! ```text
//! (format 1 "6B4FC8" general nowrap (font bold "FFFFFF") (align center))
//!
//! (font bold italic 22 "FFFFFF")        a weight, bold or not-bold; a slant,
//!                                       italic or not-italic; a size in points
//!                                       from 1 to 409.55 by twentieths, as
//!                                       Office keeps one; a color, six hex digits
//! (align center middle)                 general, left, center or right; top,
//!                                       middle or bottom (Excel's Middle Align)
//! (border (left thin) (top medium "6B4FC8"))
//!                                       an edge said whole: its line and its
//!                                       color, Excel's automatic where none
//! ```
//!
//! A format that says none of them prints as a format printed before them,
//! byte for byte, and a viewport that reads a format row's first four items
//! passes over its groups.

use crate::form::Form;
use crate::printer::write_datum;
use crate::rows::quoted;
use crate::sheet::Rgb;

/// What a format says of its text's font. A list that may grow
/// (`#[non_exhaustive]`): an underline or a family would join it.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
#[non_exhaustive]
pub struct Font {
    /// The weight: bold, or not.
    pub bold: Option<bool>,
    /// The slant: italic, or not.
    pub italic: Option<bool>,
    /// The size; unsaid, the workbook's default, which the writer's font
    /// makes 11 points.
    pub size: Option<FontSize>,
    /// The text's color; unsaid, the host's own ink.
    pub color: Option<Rgb>,
}

impl Font {
    /// This font's said properties, and `own`'s where this one says none.
    pub fn over(&self, own: &Font) -> Font {
        Font {
            bold: self.bold.or(own.bold),
            italic: self.italic.or(own.italic),
            size: self.size.or(own.size),
            color: self.color.or(own.color),
        }
    }
}

/// A font's size: a whole number of twentieths of a point from 20 to 8,191,
/// 1 to 409.55 points, the unit and the bounds Office keeps a size in
/// ([MS-OI29500] 2.1.677, on ECMA-376 Part 1 §18.4.11, `sz`).
#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct FontSize(u16);

impl FontSize {
    /// The least size, one point.
    pub const LEAST: u16 = 20;
    /// The most, 409.55 points.
    pub const MOST: u16 = 8_191;

    /// A size of `n` twentieths of a point, within the bounds.
    pub fn from_twentieths(n: u16) -> Option<FontSize> {
        (Self::LEAST..=Self::MOST)
            .contains(&n)
            .then_some(FontSize(n))
    }

    /// Its twentieths of a point.
    pub fn twentieths(self) -> u16 {
        self.0
    }

    /// A size as the record writes it: digits, then a point and one or two
    /// more, read exactly in hundredths, so that a size is taken only where
    /// it is a whole number of twentieths (`10.25` and `10.3` are, `10.33` is not) and
    /// within the bounds; no sign, no exponent.
    pub fn parse(text: &str) -> Option<FontSize> {
        let (whole, part) = match text.split_once('.') {
            Some((w, p)) => (w, p),
            None => (text, ""),
        };
        let digits = |s: &str| s.bytes().all(|b| b.is_ascii_digit());
        if whole.is_empty() || whole.len() > 3 || !digits(whole) || !digits(part) {
            return None;
        }
        if text.contains('.') && (part.is_empty() || part.len() > 2) {
            return None;
        }
        let mut hundredths: u32 = whole.parse::<u32>().ok()? * 100;
        match part.len() {
            1 => hundredths += part.parse::<u32>().ok()? * 10,
            2 => hundredths += part.parse::<u32>().ok()?,
            _ => {}
        }
        if hundredths % 5 != 0 {
            return None;
        }
        u16::try_from(hundredths / 5)
            .ok()
            .and_then(Self::from_twentieths)
    }

    /// The size in points at its shortest: `22`, `10.5`, `10.25`, `409.55`.
    pub fn spell(self) -> String {
        let whole = self.0 / 20;
        let hundredths = (self.0 % 20) * 5;
        match hundredths {
            0 => whole.to_string(),
            h if h % 10 == 0 => format!("{whole}.{}", h / 10),
            h => format!("{whole}.{h:02}"),
        }
    }
}

/// What a format says of where its text sits in its cell. A list that may
/// grow (`#[non_exhaustive]`): an indent or a rotation would join it.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
#[non_exhaustive]
pub struct Align {
    pub horizontal: Option<Horizontal>,
    pub vertical: Option<Vertical>,
}

impl Align {
    /// This alignment's said properties, and `own`'s where it says none.
    pub fn over(&self, own: &Align) -> Align {
        Align {
            horizontal: self.horizontal.or(own.horizontal),
            vertical: self.vertical.or(own.vertical),
        }
    }
}

/// A horizontal alignment: `general`, Excel's default, a number right, a
/// text left and a truth value or an error centered; or one of the three.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum Horizontal {
    General,
    Left,
    Center,
    Right,
}

const HORIZONTALS: [(Horizontal, &str); 4] = [
    (Horizontal::General, "general"),
    (Horizontal::Left, "left"),
    (Horizontal::Center, "center"),
    (Horizontal::Right, "right"),
];

impl Horizontal {
    /// The record's word, which is Excel's own name for it in a file.
    pub fn word(self) -> &'static str {
        HORIZONTALS
            .iter()
            .find(|(h, _)| *h == self)
            .map_or("", |p| p.1)
    }

    /// Excel's name for it in a file's `<alignment horizontal>`.
    pub fn excel(self) -> &'static str {
        self.word()
    }

    /// The alignment a record's word names.
    pub fn from_word(word: &str) -> Option<Horizontal> {
        HORIZONTALS.iter().find(|(_, w)| *w == word).map(|p| p.0)
    }
}

/// A vertical alignment: `bottom`, Excel's default; `middle`, the vertical
/// center, as the ribbon's Middle Align and the reference's `align-middle`
/// name it (a file and the object model say `center`); or `top`.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum Vertical {
    Top,
    Middle,
    Bottom,
}

const VERTICALS: [(Vertical, &str, &str); 3] = [
    (Vertical::Top, "top", "top"),
    (Vertical::Middle, "middle", "center"),
    (Vertical::Bottom, "bottom", "bottom"),
];

impl Vertical {
    /// The record's word.
    pub fn word(self) -> &'static str {
        VERTICALS.iter().find(|v| v.0 == self).map_or("", |v| v.1)
    }

    /// Excel's name for it in a file's `<alignment vertical>`.
    pub fn excel(self) -> &'static str {
        VERTICALS.iter().find(|v| v.0 == self).map_or("", |v| v.2)
    }

    /// The alignment a record's word names.
    pub fn from_word(word: &str) -> Option<Vertical> {
        VERTICALS.iter().find(|v| v.1 == word).map(|v| v.0)
    }
}

/// What a format says of a cell's four edges, each said whole or not. A
/// list that may grow (`#[non_exhaustive]`): the diagonal would join it.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
#[non_exhaustive]
pub struct Border {
    pub left: Option<Edge>,
    pub right: Option<Edge>,
    pub top: Option<Edge>,
    pub bottom: Option<Edge>,
}

impl Border {
    /// This border's said edges, and `own`'s where it says none.
    pub fn over(&self, own: &Border) -> Border {
        Border {
            left: self.left.or(own.left),
            right: self.right.or(own.right),
            top: self.top.or(own.top),
            bottom: self.bottom.or(own.bottom),
        }
    }

    /// The edges in the order a record prints them and a file writes them,
    /// each with its word.
    pub fn edges(&self) -> [(&'static str, Option<Edge>); 4] {
        [
            ("left", self.left),
            ("right", self.right),
            ("top", self.top),
            ("bottom", self.bottom),
        ]
    }
}

/// One edge: its line, and its color, Excel's automatic where none is said.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub struct Edge {
    pub line: Line,
    pub color: Option<Rgb>,
}

impl Edge {
    pub fn new(line: Line, color: Option<Rgb>) -> Edge {
        Edge { line, color }
    }
}

/// An edge's line: `none`, or one of the thirteen Excel draws
/// (ECMA-376 Part 1, `ST_BorderStyle`), a closed list.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Line {
    None,
    Thin,
    Medium,
    Thick,
    Double,
    Hair,
    Dotted,
    Dashed,
    DashDot,
    DashDotDot,
    MediumDashed,
    MediumDashDot,
    MediumDashDotDot,
    SlantDashDot,
}

/// Each line, the record's word and Excel's name in a file.
const LINES: [(Line, &str, &str); 14] = [
    (Line::None, "none", "none"),
    (Line::Thin, "thin", "thin"),
    (Line::Medium, "medium", "medium"),
    (Line::Thick, "thick", "thick"),
    (Line::Double, "double", "double"),
    (Line::Hair, "hair", "hair"),
    (Line::Dotted, "dotted", "dotted"),
    (Line::Dashed, "dashed", "dashed"),
    (Line::DashDot, "dash-dot", "dashDot"),
    (Line::DashDotDot, "dash-dot-dot", "dashDotDot"),
    (Line::MediumDashed, "medium-dashed", "mediumDashed"),
    (Line::MediumDashDot, "medium-dash-dot", "mediumDashDot"),
    (
        Line::MediumDashDotDot,
        "medium-dash-dot-dot",
        "mediumDashDotDot",
    ),
    (Line::SlantDashDot, "slant-dash-dot", "slantDashDot"),
];

impl Line {
    /// Every line, in the order of Excel's list.
    pub fn all() -> [Line; 14] {
        LINES.map(|l| l.0)
    }

    /// The record's word.
    pub fn word(self) -> &'static str {
        LINES.iter().find(|l| l.0 == self).map_or("", |l| l.1)
    }

    /// Excel's name in a file's `<left style>` and its kin.
    pub fn excel(self) -> &'static str {
        LINES.iter().find(|l| l.0 == self).map_or("", |l| l.2)
    }

    /// The line a record's word names.
    pub fn from_word(word: &str) -> Option<Line> {
        LINES.iter().find(|l| l.1 == word).map(|l| l.0)
    }

    /// The line Excel's name in a file names.
    pub fn from_excel(name: &str) -> Option<Line> {
        LINES.iter().find(|l| l.2 == name).map(|l| l.0)
    }
}

/// What a format row's groups said.
#[derive(Debug, Default)]
pub(crate) struct Said {
    pub font: Font,
    pub align: Align,
    pub border: Border,
}

/// Why a format row's groups said nothing a format takes.
#[derive(Debug)]
pub(crate) enum Wrong {
    /// Not in the row's shape: an item that is no group, a group whose head
    /// is none of `font`, `align` and `border`, or a group given twice.
    Shape,
    /// A group holding what a format does not take: the group, what it held,
    /// and what the group takes.
    Property {
        group: &'static str,
        what: String,
        fix: &'static str,
    },
}

const FONT_TAKES: &str = "a font takes bold or not-bold, italic or not-italic, a size in points from 1 to 409.55 in steps of 0.05, and a color of six hex digits in quotes, each at most once";
const ALIGN_TAKES: &str = "an alignment takes one of general, left, center and right, and one of top, middle and bottom, each at most once";
const BORDER_TAKES: &str = "a border takes each edge at most once, (left <line>), (right <line>), (top <line>) or (bottom <line>), a color of six hex digits in quotes after the line where wanted, the line one of none, thin, medium, thick, double, hair, dotted, dashed, dash-dot, dash-dot-dot, medium-dashed, medium-dash-dot, medium-dash-dot-dot and slant-dash-dot, and none taking no color";

/// The groups after a format row's wrap, read in any order.
pub(crate) fn read_groups(items: &[Form]) -> Result<Said, Wrong> {
    let mut said = Said::default();
    let mut seen = [false; 3];
    for item in items {
        let Form::List(l) = item else {
            return Err(Wrong::Shape);
        };
        let Some(Form::Sym(head)) = l.items.first() else {
            return Err(Wrong::Shape);
        };
        let rest = &l.items[1..];
        let k = match head.as_str() {
            "font" => 0,
            "align" => 1,
            "border" => 2,
            _ => return Err(Wrong::Shape),
        };
        if seen[k] {
            return Err(Wrong::Shape);
        }
        seen[k] = true;
        match k {
            0 => said.font = read_font(rest)?,
            1 => said.align = read_align(rest)?,
            _ => said.border = read_border(rest)?,
        }
    }
    Ok(said)
}

fn property(group: &'static str, what: String, fix: &'static str) -> Wrong {
    Wrong::Property { group, what, fix }
}

fn read_font(items: &[Form]) -> Result<Font, Wrong> {
    let wrong = |what: String| property("font", what, FONT_TAKES);
    if items.is_empty() {
        return Err(wrong("nothing".to_string()));
    }
    let mut font = Font::default();
    for item in items {
        match item {
            Form::Sym(w) if w == "bold" || w == "not-bold" => {
                if font.bold.is_some() {
                    return Err(wrong("a second weight".to_string()));
                }
                font.bold = Some(w == "bold");
            }
            Form::Sym(w) if w == "italic" || w == "not-italic" => {
                if font.italic.is_some() {
                    return Err(wrong("a second slant".to_string()));
                }
                font.italic = Some(w == "italic");
            }
            Form::Sym(n) if n.starts_with(|c: char| c.is_ascii_digit() || c == '.') => {
                let size = FontSize::parse(n).ok_or_else(|| wrong(n.clone()))?;
                if font.size.is_some() {
                    return Err(wrong("a second size".to_string()));
                }
                font.size = Some(size);
            }
            Form::Str(hex) => {
                let color = Rgb::from_hex(hex).ok_or_else(|| wrong(quoted(hex)))?;
                if font.color.is_some() {
                    return Err(wrong("a second color".to_string()));
                }
                font.color = Some(color);
            }
            other => return Err(wrong(write_datum(other))),
        }
    }
    Ok(font)
}

fn read_align(items: &[Form]) -> Result<Align, Wrong> {
    let wrong = |what: String| property("align", what, ALIGN_TAKES);
    if items.is_empty() {
        return Err(wrong("nothing".to_string()));
    }
    let mut align = Align::default();
    for item in items {
        let Form::Sym(w) = item else {
            return Err(wrong(write_datum(item)));
        };
        if let Some(h) = Horizontal::from_word(w) {
            if align.horizontal.is_some() {
                return Err(wrong("a second horizontal alignment".to_string()));
            }
            align.horizontal = Some(h);
        } else if let Some(v) = Vertical::from_word(w) {
            if align.vertical.is_some() {
                return Err(wrong("a second vertical alignment".to_string()));
            }
            align.vertical = Some(v);
        } else {
            return Err(wrong(w.clone()));
        }
    }
    Ok(align)
}

fn read_border(items: &[Form]) -> Result<Border, Wrong> {
    let wrong = |what: String| property("border", what, BORDER_TAKES);
    if items.is_empty() {
        return Err(wrong("nothing".to_string()));
    }
    let mut border = Border::default();
    for item in items {
        let Form::List(l) = item else {
            return Err(wrong(write_datum(item)));
        };
        let (side, line, color) = match l.items.as_slice() {
            [Form::Sym(side), Form::Sym(line)] => (side, line, None),
            [Form::Sym(side), Form::Sym(line), Form::Str(color)] => (side, line, Some(color)),
            _ => return Err(wrong(write_datum(item))),
        };
        let line = Line::from_word(line).ok_or_else(|| wrong(write_datum(item)))?;
        let color = match color {
            Some(_) if line == Line::None => return Err(wrong(write_datum(item))),
            Some(hex) => Some(Rgb::from_hex(hex).ok_or_else(|| wrong(write_datum(item)))?),
            None => None,
        };
        let slot = match side.as_str() {
            "left" => &mut border.left,
            "right" => &mut border.right,
            "top" => &mut border.top,
            "bottom" => &mut border.bottom,
            _ => return Err(wrong(write_datum(item))),
        };
        if slot.is_some() {
            return Err(wrong(format!("the {side} edge twice")));
        }
        *slot = Some(Edge::new(line, color));
    }
    Ok(border)
}

/// The groups a format says, as its row prints them after its wrap, each
/// with the space before it: the font, the alignment and the border it
/// says anything of, in that order, and nothing for a format that says
/// none of them.
pub(crate) fn groups_text(font: &Font, align: &Align, border: &Border) -> String {
    let mut out = String::new();
    // A format that says none of them, every format of a record before
    // KERNEL.50, asks nothing more.
    if *font == Font::default() && *align == Align::default() && *border == Border::default() {
        return out;
    }
    let mut parts: Vec<String> = Vec::new();
    if let Some(b) = font.bold {
        parts.push(if b { "bold" } else { "not-bold" }.to_string());
    }
    if let Some(i) = font.italic {
        parts.push(if i { "italic" } else { "not-italic" }.to_string());
    }
    if let Some(size) = font.size {
        parts.push(size.spell());
    }
    if let Some(color) = font.color {
        parts.push(quoted(&color.hex()));
    }
    if !parts.is_empty() {
        out.push_str(&format!(" (font {})", parts.join(" ")));
    }
    let words: Vec<&str> = [
        align.horizontal.map(Horizontal::word),
        align.vertical.map(Vertical::word),
    ]
    .into_iter()
    .flatten()
    .collect();
    if !words.is_empty() {
        out.push_str(&format!(" (align {})", words.join(" ")));
    }
    let edges: Vec<String> = border
        .edges()
        .into_iter()
        .filter_map(|(side, edge)| {
            edge.map(|e| match e.color {
                Some(c) => format!("({side} {} {})", e.line.word(), quoted(&c.hex())),
                None => format!("({side} {})", e.line.word()),
            })
        })
        .collect();
    if !edges.is_empty() {
        out.push_str(&format!(" (border {})", edges.join(" ")));
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::read_forms;

    fn groups(text: &str) -> Result<Said, Wrong> {
        let forms = read_forms(text).expect("forms");
        read_groups(&forms)
    }

    fn printed(text: &str) -> String {
        let s = groups(text).unwrap_or_else(|e| panic!("{text}: {e:?}"));
        groups_text(&s.font, &s.align, &s.border)
    }

    fn refused(text: &str) -> (&'static str, String) {
        match groups(text) {
            Err(Wrong::Property { group, what, .. }) => (group, what),
            other => panic!("{text}: {other:?}"),
        }
    }

    #[test]
    fn a_size_is_a_whole_number_of_twentieths_within_office_s_bounds() {
        let size = |t: &str| FontSize::parse(t).map(FontSize::twentieths);
        assert_eq!(size("22"), Some(440));
        assert_eq!(size("10.5"), Some(210));
        assert_eq!(size("10.50"), Some(210));
        assert_eq!(size("10.25"), Some(205));
        assert_eq!(size("10.05"), Some(201));
        assert_eq!(size("10.3"), Some(206));
        assert_eq!(size("1"), Some(20));
        assert_eq!(size("409.55"), Some(8_191));
        for text in [
            "0.95", "409.6", "410", "10.33", "10.02", "10.125", "-3", "1e1", "", ".5", "5.", "1,5",
            "0022.0", "x",
        ] {
            assert_eq!(size(text), None, "{text}");
        }
        for (n, spelled) in [
            (440, "22"),
            (210, "10.5"),
            (205, "10.25"),
            (201, "10.05"),
            (20, "1"),
            (8_191, "409.55"),
        ] {
            let s = FontSize::from_twentieths(n).unwrap();
            assert_eq!(s.spell(), spelled);
            assert_eq!(FontSize::parse(spelled), Some(s), "{spelled}");
        }
        assert_eq!(FontSize::from_twentieths(19), None);
        assert_eq!(FontSize::from_twentieths(8_192), None);
    }

    #[test]
    fn the_groups_print_in_one_order_whatever_order_they_were_read_in() {
        assert_eq!(
            printed(
                "(border (top thin) (left medium \"6b4fc8\")) (align middle center) (font \"ffffff\" 22 italic bold)"
            ),
            " (font bold italic 22 \"FFFFFF\") (align center middle) (border (left medium \"6B4FC8\") (top thin))"
        );
        assert_eq!(printed(""), "");
        assert_eq!(
            printed("(font not-bold not-italic)"),
            " (font not-bold not-italic)"
        );
        assert_eq!(printed("(align general bottom)"), " (align general bottom)");
        assert_eq!(printed("(align bottom)"), " (align bottom)");
        assert_eq!(
            printed("(border (bottom none) (right double \"000000\"))"),
            " (border (right double \"000000\") (bottom none))"
        );
        // Every line of Excel's list, by its word, read and printed back.
        for line in Line::all() {
            let text = format!("(border (left {}))", line.word());
            assert_eq!(printed(&text), format!(" {text}"));
            assert_eq!(Line::from_excel(line.excel()), Some(line));
        }
        assert_eq!(Vertical::Middle.excel(), "center");
        assert_eq!(Horizontal::Center.excel(), "center");
    }

    #[test]
    fn a_group_holding_what_a_format_does_not_take_is_refused_naming_it() {
        let cases: &[(&str, &str, &str)] = &[
            ("(font)", "font", "nothing"),
            ("(font heavy)", "font", "heavy"),
            ("(font bold not-bold)", "font", "a second weight"),
            ("(font italic italic)", "font", "a second slant"),
            ("(font 11 12)", "font", "a second size"),
            ("(font 10.33)", "font", "10.33"),
            ("(font 410)", "font", "410"),
            ("(font \"FFF\")", "font", "\"FFF\""),
            ("(font \"FFFFFF\" \"000000\")", "font", "a second color"),
            ("(font (bold))", "font", "(bold)"),
            ("(align)", "align", "nothing"),
            ("(align centre)", "align", "centre"),
            (
                "(align left right)",
                "align",
                "a second horizontal alignment",
            ),
            ("(align top middle)", "align", "a second vertical alignment"),
            ("(align \"center\")", "align", "\"center\""),
            ("(border)", "border", "nothing"),
            ("(border left thin)", "border", "left"),
            ("(border (left))", "border", "(left)"),
            ("(border (left heavy))", "border", "(left heavy)"),
            ("(border (middle thin))", "border", "(middle thin)"),
            (
                "(border (left thin \"12345\"))",
                "border",
                "(left thin \"12345\")",
            ),
            (
                "(border (left none \"000000\"))",
                "border",
                "(left none \"000000\")",
            ),
            (
                "(border (left thin) (left medium))",
                "border",
                "the left edge twice",
            ),
            (
                "(border (top thin \"000000\" x))",
                "border",
                "(top thin \"000000\" x)",
            ),
        ];
        for (text, group, what) in cases {
            assert_eq!(refused(text), (*group, what.to_string()), "{text}");
        }
        // Not a group at all, a head that is none of the three, or a group
        // twice: the row's shape.
        for text in [
            "bold",
            "(color \"FFFFFF\")",
            "(font bold) (font 22)",
            "(\"font\" bold)",
            "()",
        ] {
            assert!(matches!(groups(text), Err(Wrong::Shape)), "{text}");
        }
    }

    #[test]
    fn a_look_says_what_its_format_says_and_the_cell_s_own_the_rest() {
        let said = |text: &str| groups(text).unwrap();
        let own = said(
            "(font bold 11 \"1D1D1B\") (align center middle) (border (left medium) (top medium \"6B4FC8\"))",
        );
        let look = said("(font 22) (border (top none))");
        let font = look.font.over(&own.font);
        assert_eq!(font.bold, Some(true));
        assert_eq!(font.size.map(FontSize::twentieths), Some(440));
        assert_eq!(font.color, own.font.color);
        assert_eq!(look.align.over(&own.align), own.align);
        let border = look.border.over(&own.border);
        assert_eq!(border.left, own.border.left);
        assert_eq!(border.top, Some(Edge::new(Line::None, None)));
        assert_eq!(border.right, None);
        // A said default takes the cell's own back.
        let back = said("(font not-bold not-italic) (align general bottom)");
        assert_eq!(back.font.over(&own.font).bold, Some(false));
        let align = back.align.over(&own.align);
        assert_eq!(
            (align.horizontal, align.vertical),
            (Some(Horizontal::General), Some(Vertical::Bottom))
        );
        // A look that says nothing leaves the cell's own whole, and a cell
        // that says nothing takes the look's.
        let none = Said::default();
        assert_eq!(none.font.over(&own.font), own.font);
        assert_eq!(none.border.over(&own.border), own.border);
        assert_eq!(own.font.over(&none.font), own.font);
    }
}
