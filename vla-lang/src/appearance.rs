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
//!                                       Office keeps one; a color, six hex digits;
//!                                       and since KERNEL.58 an underline, a strike
//!                                       and a family, (family "Consolas" mono)
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
//!
//! A format of groups alone (KERNEL.58) spells neither a fill, a number
//! format nor a wrap in its row's places, and says what its groups say and no
//! more, three more groups saying those three, so that a look of it draws the
//! cell's own for the rest ([`Unsaid`]):
//!
//! ```text
//! (format 64 (fill "FFE699") (font "800020"))
//!
//! (fill "FFE699")                       a fill, six hex digits, or (fill none)
//! (number "0.00")                       general, text, or a code of Excel's
//! (wrap nowrap)                         wrap or nowrap
//! ```
//!
//! A format whose groups say all three prints as a format that spells them. A
//! format of groups alone is no record's before this item, and one whose row
//! says nothing is refused.

use crate::form::Form;
use crate::numfmt::{Code, Unsupported};
use crate::printer::write_datum;
use crate::rows::quoted;
use crate::sheet::{NumFmt, Rgb};

const UNSAID_FILL: u8 = 1;
const UNSAID_NUMBER: u8 = 2;
const UNSAID_WRAP: u8 = 4;

/// What a format leaves unsaid of its fill, its number format and its wrap
/// (KERNEL.58): nothing, as every format whose row spells the three; or, in a
/// format of groups alone, what its groups do not say, so that a look of it
/// draws the cell's own for each ([`crate::sheet::Style::over`]). A format
/// holds the default of what it leaves unsaid, which is what it draws as a
/// cell's own format.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct Unsaid(u8);

impl Unsaid {
    /// Nothing unsaid, as every format whose row spells the three.
    pub const NOTHING: Unsaid = Unsaid(0);
    /// All three unsaid.
    pub const ALL: Unsaid = Unsaid(UNSAID_FILL | UNSAID_NUMBER | UNSAID_WRAP);

    /// Whether the fill is unsaid.
    pub fn fill(self) -> bool {
        self.0 & UNSAID_FILL != 0
    }

    /// Whether the number format is unsaid.
    pub fn number(self) -> bool {
        self.0 & UNSAID_NUMBER != 0
    }

    /// Whether the wrap is unsaid.
    pub fn wrap(self) -> bool {
        self.0 & UNSAID_WRAP != 0
    }

    /// What both leave unsaid, which is what a format drawn of the two leaves.
    pub(crate) fn both(self, other: Unsaid) -> Unsaid {
        Unsaid(self.0 & other.0)
    }

    pub(crate) fn said_fill(&mut self) {
        self.0 &= !UNSAID_FILL;
    }

    pub(crate) fn said_number(&mut self) {
        self.0 &= !UNSAID_NUMBER;
    }

    pub(crate) fn said_wrap(&mut self) {
        self.0 &= !UNSAID_WRAP;
    }
}

/// What a format says of its text's font. A list that may grow
/// (`#[non_exhaustive]`). A family, which is a name of any length, is the
/// format's own ([`crate::sheet::Style::family`]), so that a font stays
/// `Copy` and a format's table slot a pointer more.
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
    /// The underline, one of Excel's or none (KERNEL.58).
    pub underline: Option<Underline>,
    /// The strike: struck through, or not (KERNEL.58).
    pub strike: Option<bool>,
}

impl Font {
    /// This font's said properties, and `own`'s where this one says none.
    pub fn over(&self, own: &Font) -> Font {
        Font {
            bold: self.bold.or(own.bold),
            italic: self.italic.or(own.italic),
            size: self.size.or(own.size),
            color: self.color.or(own.color),
            underline: self.underline.or(own.underline),
            strike: self.strike.or(own.strike),
        }
    }
}

/// A text's underline: none, or one of Excel's four (ECMA-376 Part 1,
/// `ST_UnderlineValues`), a closed list.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Underline {
    None,
    Single,
    Double,
    SingleAccounting,
    DoubleAccounting,
}

/// Each underline, the record's word and Excel's name in a file's `<u val>`.
const UNDERLINES: [(Underline, &str, &str); 5] = [
    (Underline::None, "not-underline", "none"),
    (Underline::Single, "underline", "single"),
    (Underline::Double, "double-underline", "double"),
    (
        Underline::SingleAccounting,
        "accounting-underline",
        "singleAccounting",
    ),
    (
        Underline::DoubleAccounting,
        "double-accounting-underline",
        "doubleAccounting",
    ),
];

impl Underline {
    /// The record's word.
    pub fn word(self) -> &'static str {
        UNDERLINES.iter().find(|u| u.0 == self).map_or("", |u| u.1)
    }

    /// Excel's name in a file's `<u val>`, `single` where it writes none.
    pub fn excel(self) -> &'static str {
        UNDERLINES.iter().find(|u| u.0 == self).map_or("", |u| u.2)
    }

    /// The underline a record's word names.
    pub fn from_word(word: &str) -> Option<Underline> {
        UNDERLINES.iter().find(|u| u.1 == word).map(|u| u.0)
    }

    /// The underline Excel's name in a file names.
    pub fn from_excel(name: &str) -> Option<Underline> {
        UNDERLINES.iter().find(|u| u.2 == name).map(|u| u.0)
    }
}

/// A font's family (KERNEL.58): its name as a file holds it, and the generic
/// it belongs to, which is what draws where a reading machine lacks the name.
/// A list that may grow (`#[non_exhaustive]`).
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub struct Family {
    pub name: String,
    pub generic: Option<Generic>,
}

impl Family {
    /// The most characters a name holds, Excel's own.
    pub const MOST: usize = 31;

    /// A family of a name of 1 to [`Family::MOST`] characters.
    pub fn new(name: &str, generic: Option<Generic>) -> Option<Family> {
        let n = name.chars().count();
        (1..=Self::MOST).contains(&n).then(|| Family {
            name: name.to_string(),
            generic,
        })
    }
}

/// A family's generic, a file's `<family val>` and an OpenDocument file's
/// `style:font-family-generic`, a closed list.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Generic {
    Serif,
    Sans,
    Mono,
    Script,
    Decorative,
}

/// Each generic, the record's word, its number in a file's `<family val>`
/// and its word in an OpenDocument file.
const GENERICS: [(Generic, &str, u32, &str); 5] = [
    (Generic::Serif, "serif", 1, "roman"),
    (Generic::Sans, "sans", 2, "swiss"),
    (Generic::Mono, "mono", 3, "modern"),
    (Generic::Script, "script", 4, "script"),
    (Generic::Decorative, "decorative", 5, "decorative"),
];

impl Generic {
    /// The record's word.
    pub fn word(self) -> &'static str {
        GENERICS.iter().find(|g| g.0 == self).map_or("", |g| g.1)
    }

    /// Its number in a file's `<family val>`.
    pub fn excel(self) -> u32 {
        GENERICS.iter().find(|g| g.0 == self).map_or(0, |g| g.2)
    }

    /// The generic a record's word names.
    pub fn from_word(word: &str) -> Option<Generic> {
        GENERICS.iter().find(|g| g.1 == word).map(|g| g.0)
    }

    /// The generic a file's `<family val>` names; `None` for 0, no family.
    pub fn from_excel(n: u32) -> Option<Generic> {
        GENERICS.iter().find(|g| g.2 == n).map(|g| g.0)
    }

    /// The generic an OpenDocument file's `style:font-family-generic` names;
    /// `None` for `system`, the machine's own, which is no family.
    pub fn from_odf(word: &str) -> Option<Generic> {
        GENERICS.iter().find(|g| g.3 == word).map(|g| g.0)
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
    /// The family its font group says (KERNEL.58).
    pub family: Option<Family>,
    pub align: Align,
    pub border: Border,
    /// A format of groups alone's fill, number format and wrap, each where
    /// its group says one (KERNEL.58).
    pub fill: Option<Option<Rgb>>,
    pub number: Option<NumFmt>,
    pub wrap: Option<bool>,
}

impl Said {
    /// Whether it said a fill, a number format or a wrap, which a format row
    /// spelling the three in its places says there alone.
    pub fn says_the_places(&self) -> bool {
        self.fill.is_some() || self.number.is_some() || self.wrap.is_some()
    }
}

/// Why a format row's groups said nothing a format takes.
#[derive(Debug)]
pub(crate) enum Wrong {
    /// Not in the row's shape: an item that is no group, a group whose head
    /// is none of `font`, `align`, `border`, `fill`, `number` and `wrap`, or
    /// a group given twice.
    Shape,
    /// A group holding what a format does not take: the group, what it held,
    /// and what the group takes.
    Property {
        group: &'static str,
        what: String,
        fix: &'static str,
    },
    /// A `number` group's code this version does not draw, refused as a code
    /// in a format's place is (KERNEL.47).
    Code { code: String, why: Unsupported },
}

const FONT_TAKES: &str = "a font takes bold or not-bold, italic or not-italic, a size in points from 1 to 409.55 in steps of 0.05, a color of six hex digits in quotes, underline, double-underline, accounting-underline, double-accounting-underline or not-underline, strike or not-strike, and (family \"<name>\" [serif|sans|mono|script|decorative]), a name of 1 to 31 characters, each at most once";
const ALIGN_TAKES: &str = "an alignment takes one of general, left, center and right, and one of top, middle and bottom, each at most once";
const BORDER_TAKES: &str = "a border takes each edge at most once, (left <line>), (right <line>), (top <line>) or (bottom <line>), a color of six hex digits in quotes after the line where wanted, the line one of none, thin, medium, thick, double, hair, dotted, dashed, dash-dot, dash-dot-dot, medium-dashed, medium-dash-dot, medium-dash-dot-dot and slant-dash-dot, and none taking no color";
const FILL_TAKES: &str = "a fill takes one color of six hex digits in quotes, or none";
const NUMBER_TAKES: &str =
    "a number format takes one of general, text and a code of Excel's in quotes";
const WRAP_TAKES: &str = "a wrap takes one of wrap and nowrap";

/// The groups after a format row's wrap, or all of a format of groups alone,
/// read in any order.
pub(crate) fn read_groups(items: &[Form]) -> Result<Said, Wrong> {
    let mut said = Said::default();
    let mut seen = [false; 6];
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
            "fill" => 3,
            "number" => 4,
            "wrap" => 5,
            _ => return Err(Wrong::Shape),
        };
        if seen[k] {
            return Err(Wrong::Shape);
        }
        seen[k] = true;
        match k {
            0 => (said.font, said.family) = read_font(rest)?,
            1 => said.align = read_align(rest)?,
            2 => said.border = read_border(rest)?,
            3 => said.fill = Some(read_fill(rest)?),
            4 => said.number = Some(read_number(rest)?),
            _ => said.wrap = Some(read_wrap(rest)?),
        }
    }
    Ok(said)
}

fn property(group: &'static str, what: String, fix: &'static str) -> Wrong {
    Wrong::Property { group, what, fix }
}

/// A group that holds one item: that item; what it holds else, refused.
fn one<'f>(items: &'f [Form], group: &'static str, fix: &'static str) -> Result<&'f Form, Wrong> {
    match items {
        [item] => Ok(item),
        [] => Err(property(group, "nothing".to_string(), fix)),
        [_, second, ..] => Err(property(group, write_datum(second), fix)),
    }
}

fn read_fill(items: &[Form]) -> Result<Option<Rgb>, Wrong> {
    let item = one(items, "fill", FILL_TAKES)?;
    match item {
        Form::Sym(w) if w == "none" => Ok(None),
        Form::Str(hex) => Rgb::from_hex(hex)
            .map(Some)
            .ok_or_else(|| property("fill", quoted(hex), FILL_TAKES)),
        other => Err(property("fill", write_datum(other), FILL_TAKES)),
    }
}

fn read_number(items: &[Form]) -> Result<NumFmt, Wrong> {
    let item = one(items, "number", NUMBER_TAKES)?;
    match item {
        Form::Sym(w) if w == "general" => Ok(NumFmt::General),
        Form::Sym(w) if w == "text" => Ok(NumFmt::Text),
        Form::Str(code) => match Code::parse(code) {
            Ok(code) => Ok(NumFmt::Code(code)),
            Err(why) => Err(Wrong::Code {
                code: code.clone(),
                why,
            }),
        },
        other => Err(property("number", write_datum(other), NUMBER_TAKES)),
    }
}

fn read_wrap(items: &[Form]) -> Result<bool, Wrong> {
    let item = one(items, "wrap", WRAP_TAKES)?;
    match item {
        Form::Sym(w) if w == "wrap" => Ok(true),
        Form::Sym(w) if w == "nowrap" => Ok(false),
        other => Err(property("wrap", write_datum(other), WRAP_TAKES)),
    }
}

fn read_font(items: &[Form]) -> Result<(Font, Option<Family>), Wrong> {
    let wrong = |what: String| property("font", what, FONT_TAKES);
    if items.is_empty() {
        return Err(wrong("nothing".to_string()));
    }
    let mut font = Font::default();
    let mut family = None;
    for item in items {
        match item {
            Form::Sym(w) if Underline::from_word(w).is_some() => {
                if font.underline.is_some() {
                    return Err(wrong("a second underline".to_string()));
                }
                font.underline = Underline::from_word(w);
            }
            Form::Sym(w) if w == "strike" || w == "not-strike" => {
                if font.strike.is_some() {
                    return Err(wrong("a second strike".to_string()));
                }
                font.strike = Some(w == "strike");
            }
            Form::List(l) if matches!(l.items.first(), Some(Form::Sym(h)) if h == "family") => {
                if family.is_some() {
                    return Err(wrong("a second family".to_string()));
                }
                family = Some(read_family(&l.items[1..]).ok_or_else(|| wrong(write_datum(item)))?);
            }
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
    Ok((font, family))
}

/// A `(family ...)` list's items: its name, and its generic where one is said.
fn read_family(items: &[Form]) -> Option<Family> {
    match items {
        [Form::Str(name)] => Family::new(name, None),
        [Form::Str(name), Form::Sym(generic)] => {
            Family::new(name, Some(Generic::from_word(generic)?))
        }
        _ => None,
    }
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
pub(crate) fn groups_text(
    font: &Font,
    family: Option<&Family>,
    align: &Align,
    border: &Border,
) -> String {
    let mut out = String::new();
    // A format that says none of them, every format of a record before
    // KERNEL.50, asks nothing more.
    if *font == Font::default()
        && family.is_none()
        && *align == Align::default()
        && *border == Border::default()
    {
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
    if let Some(underline) = font.underline {
        parts.push(underline.word().to_string());
    }
    if let Some(s) = font.strike {
        parts.push(if s { "strike" } else { "not-strike" }.to_string());
    }
    if let Some(f) = family {
        parts.push(match f.generic {
            Some(g) => format!("(family {} {})", quoted(&f.name), g.word()),
            None => format!("(family {})", quoted(&f.name)),
        });
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
        groups_text(&s.font, s.family.as_ref(), &s.align, &s.border)
    }

    #[test]
    fn an_underline_a_strike_and_a_family_are_words_of_a_font() {
        assert_eq!(
            printed("(font (family \"Consolas\" mono) strike \"800020\" underline 22 bold)"),
            " (font bold 22 \"800020\" underline strike (family \"Consolas\" mono))"
        );
        assert_eq!(
            printed("(font not-strike not-underline)"),
            " (font not-underline not-strike)"
        );
        assert_eq!(
            printed("(font (family \"Arial\"))"),
            " (font (family \"Arial\"))"
        );
        for (word, underline, excel) in [
            ("not-underline", Underline::None, "none"),
            ("underline", Underline::Single, "single"),
            ("double-underline", Underline::Double, "double"),
            (
                "accounting-underline",
                Underline::SingleAccounting,
                "singleAccounting",
            ),
            (
                "double-accounting-underline",
                Underline::DoubleAccounting,
                "doubleAccounting",
            ),
        ] {
            let s = groups(&format!("(font {word})")).unwrap();
            assert_eq!(s.font.underline, Some(underline), "{word}");
            assert_eq!(underline.word(), word);
            assert_eq!(underline.excel(), excel);
            assert_eq!(Underline::from_excel(excel), Some(underline));
            assert_eq!(
                printed(&format!("(font {word})")),
                format!(" (font {word})")
            );
        }
        for (word, generic, n, odf) in [
            ("serif", Generic::Serif, 1, "roman"),
            ("sans", Generic::Sans, 2, "swiss"),
            ("mono", Generic::Mono, 3, "modern"),
            ("script", Generic::Script, 4, "script"),
            ("decorative", Generic::Decorative, 5, "decorative"),
        ] {
            let s = groups(&format!("(font (family \"F\" {word}))")).unwrap();
            assert_eq!(s.family, Some(Family::new("F", Some(generic)).unwrap()));
            assert_eq!(generic.word(), word);
            assert_eq!(generic.excel(), n);
            assert_eq!(Generic::from_excel(n), Some(generic));
            assert_eq!(Generic::from_odf(odf), Some(generic));
        }
        assert_eq!(Generic::from_excel(0), None);
        assert_eq!(Generic::from_excel(6), None);
        assert_eq!(Generic::from_odf("system"), None);
        // A name of 1 to 31 characters, counted as characters.
        assert!(Family::new("", None).is_none());
        assert!(Family::new(&"x".repeat(31), None).is_some());
        assert!(Family::new(&"x".repeat(32), None).is_none());
        assert!(Family::new(&"é".repeat(31), None).is_some());
        // What a font does not take, by what it holds.
        let long = "x".repeat(32);
        let cases: &[(String, String)] = &[
            (
                "(font underline strike underline)".to_string(),
                "a second underline".to_string(),
            ),
            (
                "(font underline not-underline)".to_string(),
                "a second underline".to_string(),
            ),
            (
                "(font strike not-strike)".to_string(),
                "a second strike".to_string(),
            ),
            (
                "(font (family \"A\") (family \"B\"))".to_string(),
                "a second family".to_string(),
            ),
            ("(font (family))".to_string(), "(family)".to_string()),
            (
                "(font (family Arial))".to_string(),
                "(family Arial)".to_string(),
            ),
            (
                "(font (family \"\"))".to_string(),
                "(family \"\")".to_string(),
            ),
            (
                "(font (family \"Arial\" comic))".to_string(),
                "(family \"Arial\" comic)".to_string(),
            ),
            (
                "(font (family \"Arial\" sans mono))".to_string(),
                "(family \"Arial\" sans mono)".to_string(),
            ),
            (
                format!("(font (family \"{long}\"))"),
                format!("(family \"{long}\")"),
            ),
            (
                "(font (colour \"Arial\"))".to_string(),
                "(colour \"Arial\")".to_string(),
            ),
            ("(font single)".to_string(), "single".to_string()),
        ];
        for (text, what) in cases {
            assert_eq!(refused(text), ("font", what.clone()), "{text}");
        }
    }

    #[test]
    fn an_underline_and_a_strike_fall_over_a_cell_s_own_as_the_rest_of_a_font() {
        let own = groups("(font underline strike)").unwrap().font;
        let none = Font::default();
        assert_eq!(none.over(&own), own);
        assert_eq!(own.over(&none), own);
        let back = groups("(font not-underline not-strike)").unwrap().font;
        let drawn = back.over(&own);
        assert_eq!(
            (drawn.underline, drawn.strike),
            (Some(Underline::None), Some(false))
        );
        let double = groups("(font double-underline)").unwrap().font;
        let drawn = double.over(&own);
        assert_eq!(
            (drawn.underline, drawn.strike),
            (Some(Underline::Double), Some(true))
        );
    }

    fn refused(text: &str) -> (&'static str, String) {
        match groups(text) {
            Err(Wrong::Property { group, what, .. }) => (group, what),
            other => panic!("{text}: {other:?}"),
        }
    }

    #[test]
    fn a_fill_a_number_format_and_a_wrap_are_groups_a_format_of_groups_alone_says() {
        let s = groups("(wrap wrap) (number \"0.00\") (fill \"ffe699\")").unwrap();
        assert_eq!(s.fill, Some(Some(Rgb(0xFF, 0xE6, 0x99))));
        assert!(matches!(s.number, Some(NumFmt::Code(_))));
        assert_eq!(s.wrap, Some(true));
        assert!(s.says_the_places());
        let s = groups("(fill none) (number general) (wrap nowrap)").unwrap();
        assert_eq!(
            (s.fill, s.number, s.wrap),
            (Some(None), Some(NumFmt::General), Some(false))
        );
        assert_eq!(groups("(number text)").unwrap().number, Some(NumFmt::Text));
        // The other three say none of them.
        assert!(!groups("(font bold) (align left) (border (top thin))")
            .unwrap()
            .says_the_places());
        assert!(!groups("").unwrap().says_the_places());
        let cases: &[(&str, &str, &str)] = &[
            ("(fill)", "fill", "nothing"),
            ("(fill red)", "fill", "red"),
            ("(fill \"FFF\")", "fill", "\"FFF\""),
            ("(fill 5)", "fill", "5"),
            ("(fill none none)", "fill", "none"),
            ("(fill \"FFFFFF\" \"000000\")", "fill", "\"000000\""),
            ("(number)", "number", "nothing"),
            ("(number wide)", "number", "wide"),
            ("(number 5)", "number", "5"),
            ("(number general text)", "number", "text"),
            ("(wrap)", "wrap", "nothing"),
            ("(wrap maybe)", "wrap", "maybe"),
            ("(wrap \"wrap\")", "wrap", "\"wrap\""),
            ("(wrap wrap nowrap)", "wrap", "nowrap"),
        ];
        for (text, group, what) in cases {
            assert_eq!(refused(text), (*group, what.to_string()), "{text}");
        }
        // Twice is the row's shape, as for the other three.
        for text in [
            "(fill none) (fill none)",
            "(number text) (number text)",
            "(wrap wrap) (wrap wrap)",
        ] {
            assert!(matches!(groups(text), Err(Wrong::Shape)), "{text}");
        }
        // A code this version does not draw, by its own refusal.
        match groups("(number \"0.00E+00\")") {
            Err(Wrong::Code { code, .. }) => assert_eq!(code, "0.00E+00"),
            other => panic!("{other:?}"),
        }
    }

    #[test]
    fn what_a_format_leaves_unsaid_is_a_mask_and_a_drawn_format_leaves_what_both_do() {
        assert!(!Unsaid::NOTHING.fill() && !Unsaid::NOTHING.number() && !Unsaid::NOTHING.wrap());
        assert!(Unsaid::ALL.fill() && Unsaid::ALL.number() && Unsaid::ALL.wrap());
        assert_eq!(Unsaid::default(), Unsaid::NOTHING);
        let mut u = Unsaid::ALL;
        u.said_fill();
        assert!(!u.fill() && u.number() && u.wrap());
        u.said_number();
        assert!(!u.fill() && !u.number() && u.wrap());
        u.said_wrap();
        assert_eq!(u, Unsaid::NOTHING);
        let mut a = Unsaid::ALL;
        a.said_fill();
        let mut b = Unsaid::ALL;
        b.said_wrap();
        let both = a.both(b);
        assert!(!both.fill() && both.number() && !both.wrap());
        assert_eq!(a.both(Unsaid::NOTHING), Unsaid::NOTHING);
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
