//! An OpenDocument file's looks read as the record's rows (KERNEL.57, the
//! third slice, decision 11): what `odf.rs`'s walk asks of a spreadsheet's
//! styles to print each sheet as the record prints it whole, the same rows
//! in the same spelling as a package's (`PORT.8`'s 8e rule).
//!
//! What is read, and from where. From `content.xml` and `styles.xml`, each
//! part's font faces, its cell styles (the automatic ones of `content.xml`
//! and the named ones of `styles.xml`, `Default` among them) and the
//! family's default style, its data styles, and `content.xml`'s column and
//! row styles; from `settings.xml`, where a file has one, each sheet's
//! `ShowGrid`, else the document's, else gridlines on.
//!
//! A cell style is resolved up its chain: the family's default style, then
//! the `Default` cell style, what a workbook's default is, then each parent
//! down to the style itself, a property said nearer winning; a data style the
//! nearest one named. Format 0 is the `Default` style, and the others are
//! numbered in the order the document declares them, `content.xml`'s
//! automatic styles first and then `styles.xml`'s named ones (decision 4); a
//! style that draws as format 0 reads as format 0, as every cell of an Excel
//! save names a style of its own that says nothing.
//!
//! A format is said as a package's is said against its Normal style: a fill,
//! `fo:background-color`; a weight, a slant, a size and an ink; a horizontal
//! (`start` left, `end` right, a value's own alignment unsaid) and a vertical
//! (`automatic` bottom); the wrap; the four edges, a width, a style and a
//! color each, as Excel's lines; and a number format, the data style as a
//! code in Excel's syntax. What the record has no word for is counted, the
//! rest said: a font's family, an underline, a strike, a raised or lowered
//! text, an outline, a shadow, a text repeated to fill its cell, a justified
//! text, an indent, a rotation, a stacked text, shrinking to fit, a
//! right-to-left order and a diagonal.
//!
//! LibreOffice's own extensions of a data style are read where they keep
//! what Excel's codes say: a padding of a character's width, `_)`, as
//! `loext:blank-width-char`, and a digit's space, `?`, as
//! `loext:max-blank-integer-digits`; an empty currency symbol is the
//! currency of the style's country, else the document's. The family's
//! default wrap is passed over: LibreOffice 26.8 writes it on where its own
//! drawing of the same cells, its save's thumbnail and its fitted heights,
//! wraps none (the third slice's save of `looks.xlsx`).

use std::collections::HashMap;

use super::cursor::{attr, attributes_full, skip_element, Cursor, Event};
use super::looks::{font_size, Look};
use super::ooxml::{step, xml_refused};
use crate::messages::Refusal;
use crate::sheet::{Edge, Horizontal, Line, Rgb, Vertical};

/// A style's properties as its part writes them: every attribute of its
/// property elements, by its name as written.
type Props = HashMap<String, String>;

/// A cell style as a part declares it.
#[derive(Clone, Debug, Default)]
struct CellStyle {
    parent: Option<String>,
    data: Option<String>,
    props: Props,
    /// Its conditions, `style:map`, each a conditional format's rule.
    rules: u32,
}

/// A part of a data style, in the order it is written.
#[derive(Clone, Debug)]
enum Part {
    Number {
        decimals: Option<u32>,
        min_decimals: Option<u32>,
        min_int: u32,
        grouping: bool,
        /// Each text set among the digits, by the digits after it.
        embedded: Vec<(u32, String)>,
        /// Its scale, a thousand for each comma.
        factor: Option<f64>,
        /// The leading places of its integer that show a space for a zero,
        /// `?`, LibreOffice's `loext:max-blank-integer-digits`.
        blank_int: u32,
    },
    Scientific {
        decimals: u32,
        min_decimals: Option<u32>,
        min_int: u32,
        min_exponent: u32,
        /// The exponent's multiple, three for engineering's `##0.0E+0`.
        interval: u32,
        /// Whether a positive exponent shows its sign, `E+`, or not, `E-`.
        forced: bool,
    },
    Fraction {
        min_int: u32,
        numerator: u32,
        denominator: u32,
        value: Option<u32>,
    },
    /// A text, and the character of it that pads the width of another,
    /// `_)`, by its place in the text (`loext:blank-width-char`).
    Text(String, Option<(char, usize)>),
    /// A currency's symbol, empty for its country's own, and the country the
    /// symbol names, where it names one.
    Currency(String, Option<String>),
    Day(bool),
    Month {
        long: bool,
        textual: bool,
    },
    Year(bool),
    DayOfWeek(bool),
    Hours(bool),
    Minutes(bool),
    Seconds {
        long: bool,
        decimals: u32,
    },
    AmPm,
    TextContent,
    Boolean,
    Fill(String),
}

/// A data style as a part declares it: its kind, the local name of its
/// element (`number-style`, `date-style`), its parts, its ink, its
/// conditions, each the style it applies, and its country.
#[derive(Clone, Debug, Default)]
struct DataStyle {
    kind: String,
    parts: Vec<Part>,
    color: Option<String>,
    maps: Vec<(String, String)>,
    country: Option<String>,
    /// A time's hours past a day, `number:truncate-on-overflow="false"`.
    elapsed: bool,
}

/// What an OpenDocument file's styles say of its looks, read once at the
/// first walk under the looks.
#[derive(Debug, Default)]
pub(crate) struct OdsLooks {
    /// Each format as the record says it, format 0 the `Default` style.
    pub(crate) looks: Vec<Look>,
    /// A cell style's format by its name.
    index: HashMap<String, u32>,
    /// Each format's conditions, a conditional format's rules.
    rules: Vec<u32>,
    /// A column style's width in the record's characters, `None` for
    /// Excel's default.
    columns: HashMap<String, Option<f64>>,
    /// A row style's height in points, where a person set it.
    rows: HashMap<String, Option<f64>>,
    /// Each sheet's gridlines by its name, `settings.xml`'s.
    gridlines: HashMap<String, bool>,
    /// The document's gridlines, where `settings.xml` says them for no one
    /// sheet; on where it has none.
    document_gridlines: Option<bool>,
}

impl OdsLooks {
    /// A cell or a column naming a style: its format, 0 for a style that
    /// draws as format 0 or none the file declares.
    pub(crate) fn format(&self, style: Option<&str>) -> u32 {
        let Some(i) = style.and_then(|s| self.index.get(s)).copied() else {
            return 0;
        };
        if self.looks[i as usize] == self.looks[0] {
            0
        } else {
            i
        }
    }

    /// The conditional rules of a cell style by its name, `style:map`.
    pub(crate) fn rules_of(&self, style: &str) -> u32 {
        self.index
            .get(style)
            .and_then(|i| self.rules.get(*i as usize))
            .copied()
            .unwrap_or(0)
    }

    /// A column style's width, as the record says one.
    pub(crate) fn column_width(&self, style: Option<&str>) -> Option<f64> {
        style.and_then(|s| self.columns.get(s)).copied().flatten()
    }

    /// A row style's height where a person set it.
    pub(crate) fn row_height(&self, style: Option<&str>) -> Option<f64> {
        style.and_then(|s| self.rows.get(s)).copied().flatten()
    }

    /// A sheet's gridlines.
    pub(crate) fn gridlines(&self, sheet: &str) -> bool {
        self.gridlines
            .get(sheet)
            .copied()
            .or(self.document_gridlines)
            .unwrap_or(true)
    }
}

/// What a part's styles hold, read in one pass.
#[derive(Default)]
struct Declared {
    faces: HashMap<String, String>,
    default_style: Props,
    /// Cell styles in the order the part declares them.
    cells: Vec<(String, CellStyle)>,
    data: HashMap<String, DataStyle>,
    columns: HashMap<String, Option<f64>>,
    rows: HashMap<String, Option<f64>>,
}

/// A length a part writes, `0.74pt`, `1.69333333333333cm`, `2mm`, `1in`,
/// `12px`, `1pc`, in points.
fn points(text: &str) -> Option<f64> {
    let text = text.trim();
    let split = text
        .find(|c: char| c.is_ascii_alphabetic() || c == '%')
        .unwrap_or(text.len());
    let (number, unit) = text.split_at(split);
    let n: f64 = number.trim().parse().ok()?;
    Some(match unit {
        "pt" | "" => n,
        "cm" => n * 72.0 / 2.54,
        "mm" => n * 7.2 / 2.54,
        "in" | "inch" => n * 72.0,
        "px" => n * 0.75,
        "pc" => n * 12.0,
        _ => return None,
    })
}

/// A column's width as the record says it: the length to whole pixels at 96
/// to the inch, then to characters in hundredths as Excel reads a width
/// (decision 10), Excel's default of 64 pixels `None`.
fn column_width(text: &str) -> Option<f64> {
    let px = (points(text)? * 96.0 / 72.0).round().max(0.0);
    if px == 64.0 {
        return None;
    }
    let chars = if px < 12.0 {
        px / 12.0
    } else {
        (px - 5.0) / 7.0
    };
    Some((chars * 100.0 + 0.5).trunc() / 100.0)
}

/// A row's height a person set, in points to the twentieth.
fn row_height(props: &Props) -> Option<f64> {
    if props
        .get("style:use-optimal-row-height")
        .is_some_and(|v| v == "true")
    {
        return None;
    }
    let pt = points(props.get("style:row-height")?)?;
    Some((pt * 20.0).round() / 20.0)
}

/// The attributes of a property element into a style's properties.
fn take_props(props: &mut Props, attrs: &str) {
    for (name, value) in attributes_full(attrs) {
        props.insert(
            name.to_string(),
            super::cursor::unescape_entities(value).into_owned(),
        );
    }
}

/// A number attribute of a data style's part.
fn number_attr(attrs: &str, name: &str) -> Option<u32> {
    attr(attrs, name).and_then(|v| v.trim().parse().ok())
}

/// One part's styles read: its font faces, its default and its cell styles,
/// its data styles, and its column and row styles.
fn declared(text: &str, label: &str, part: &str, out: &mut Declared) -> Result<(), Refusal> {
    let mut cursor = Cursor::new(text);
    // The style being read: a cell style's name, a column's or a row's, or
    // the default; and the data style being read.
    let mut cell: Option<(String, CellStyle)> = None;
    let mut in_default = false;
    let mut column: Option<String> = None;
    let mut row: Option<String> = None;
    let mut data: Option<(String, DataStyle)> = None;
    // A data style's number part, open for its embedded texts, and the
    // place of the embedded text being read.
    let mut embedded_at: Option<usize> = None;
    let mut embedded_position: u32 = 0;
    let mut text_open: Option<String> = None;
    // The open text's padding character and its place, and the open
    // currency symbol's country.
    let mut text_blank: Option<(char, usize)> = None;
    let mut symbol_country: Option<String> = None;
    while let Some(event) = step(&mut cursor, label, part)? {
        match event {
            Event::Start { name, attrs, empty } => match name {
                // The document's body holds no style; a part's body is passed
                // over whole.
                "body" | "master-styles" if !empty => {
                    skip_element(&mut cursor).map_err(|why| xml_refused(label, part, why))?
                }
                "font-face" => {
                    if let Some(face) = attr(attrs, "name") {
                        let family = attr(attrs, "font-family")
                            .map(|f| f.trim_matches(['\'', '"']).to_string())
                            .unwrap_or_else(|| face.to_string());
                        out.faces.insert(face.into_owned(), family);
                    }
                }
                "default-style" => {
                    in_default = attr(attrs, "family").as_deref() == Some("table-cell") && !empty;
                }
                "style" => {
                    let family = attr(attrs, "family").unwrap_or_default();
                    let Some(style) = attr(attrs, "name").map(|n| n.into_owned()) else {
                        continue;
                    };
                    match family.as_ref() {
                        "table-cell" => {
                            let declared = CellStyle {
                                parent: attr(attrs, "parent-style-name").map(|p| p.into_owned()),
                                data: attr(attrs, "data-style-name").map(|d| d.into_owned()),
                                ..CellStyle::default()
                            };
                            if empty {
                                out.cells.push((style, declared));
                            } else {
                                cell = Some((style, declared));
                            }
                        }
                        "table-column" if !empty => column = Some(style),
                        "table-row" if !empty => row = Some(style),
                        _ => {}
                    }
                }
                "table-cell-properties" | "paragraph-properties" | "text-properties" => {
                    if let Some((_, c)) = cell.as_mut() {
                        take_props(&mut c.props, attrs);
                    } else if in_default {
                        take_props(&mut out.default_style, attrs);
                    } else if let Some((_, d)) = data.as_mut() {
                        if name == "text-properties" {
                            d.color = attr(attrs, "color").map(|c| c.into_owned());
                        }
                    }
                }
                "table-column-properties" => {
                    if let Some(c) = &column {
                        let width = attr(attrs, "column-width").and_then(|w| column_width(&w));
                        out.columns.insert(c.clone(), width);
                    }
                }
                "table-row-properties" => {
                    if let Some(r) = &row {
                        let mut props = Props::new();
                        take_props(&mut props, attrs);
                        out.rows.insert(r.clone(), row_height(&props));
                    }
                }
                "map" => {
                    if let Some((_, c)) = cell.as_mut() {
                        c.rules += 1;
                    } else if let Some((_, d)) = data.as_mut() {
                        if let (Some(condition), Some(applied)) =
                            (attr(attrs, "condition"), attr(attrs, "apply-style-name"))
                        {
                            d.maps.push((condition.into_owned(), applied.into_owned()));
                        }
                    }
                }
                kind @ ("number-style" | "percentage-style" | "currency-style" | "date-style"
                | "time-style" | "boolean-style" | "text-style") => {
                    let Some(style) = attr(attrs, "name").map(|n| n.into_owned()) else {
                        continue;
                    };
                    let d = DataStyle {
                        kind: kind.to_string(),
                        elapsed: attr(attrs, "truncate-on-overflow").as_deref() == Some("false"),
                        country: attr(attrs, "country").map(|c| c.into_owned()),
                        ..DataStyle::default()
                    };
                    if empty {
                        out.data.insert(style, d);
                    } else {
                        data = Some((style, d));
                    }
                }
                part_name => {
                    let Some((_, d)) = data.as_mut() else {
                        continue;
                    };
                    let long = attr(attrs, "style").as_deref() == Some("long");
                    let pushed = match part_name {
                        "number" => {
                            embedded_at = (!empty).then_some(d.parts.len());
                            Some(Part::Number {
                                decimals: number_attr(attrs, "decimal-places"),
                                min_decimals: number_attr(attrs, "min-decimal-places"),
                                min_int: number_attr(attrs, "min-integer-digits").unwrap_or(0),
                                grouping: attr(attrs, "grouping").as_deref() == Some("true"),
                                embedded: Vec::new(),
                                factor: attr(attrs, "display-factor")
                                    .and_then(|f| f.trim().parse().ok()),
                                blank_int: number_attr(attrs, "max-blank-integer-digits")
                                    .unwrap_or(0),
                            })
                        }
                        "scientific-number" => Some(Part::Scientific {
                            decimals: number_attr(attrs, "decimal-places").unwrap_or(0),
                            min_decimals: number_attr(attrs, "min-decimal-places"),
                            min_int: number_attr(attrs, "min-integer-digits").unwrap_or(1),
                            min_exponent: number_attr(attrs, "min-exponent-digits").unwrap_or(1),
                            interval: number_attr(attrs, "exponent-interval").unwrap_or(1),
                            forced: attr(attrs, "forced-exponent-sign").as_deref() != Some("false"),
                        }),
                        "fraction" => Some(Part::Fraction {
                            min_int: number_attr(attrs, "min-integer-digits").unwrap_or(0),
                            numerator: number_attr(attrs, "min-numerator-digits").unwrap_or(1),
                            denominator: number_attr(attrs, "min-denominator-digits").unwrap_or(1),
                            value: number_attr(attrs, "denominator-value"),
                        }),
                        "embedded-text" => {
                            text_open = Some(String::new());
                            embedded_position = number_attr(attrs, "position").unwrap_or(0);
                            None
                        }
                        "text" | "currency-symbol" | "fill-character" if !empty => {
                            text_open = Some(String::new());
                            text_blank = attr(attrs, "blank-width-char").and_then(|b| {
                                let mut chars = b.chars();
                                let c = chars.next()?;
                                Some((c, chars.as_str().trim().parse().unwrap_or(0)))
                            });
                            symbol_country = attr(attrs, "country").map(|c| c.into_owned());
                            None
                        }
                        // An empty symbol, its country's own currency.
                        "currency-symbol" => Some(Part::Currency(
                            String::new(),
                            attr(attrs, "country").map(|c| c.into_owned()),
                        )),
                        "day" => Some(Part::Day(long)),
                        "month" => Some(Part::Month {
                            long,
                            textual: attr(attrs, "textual").as_deref() == Some("true"),
                        }),
                        "year" => Some(Part::Year(long)),
                        "day-of-week" => Some(Part::DayOfWeek(long)),
                        "hours" => Some(Part::Hours(long)),
                        "minutes" => Some(Part::Minutes(long)),
                        "seconds" => Some(Part::Seconds {
                            long,
                            decimals: number_attr(attrs, "decimal-places").unwrap_or(0),
                        }),
                        "am-pm" => Some(Part::AmPm),
                        "text-content" => Some(Part::TextContent),
                        "boolean" => Some(Part::Boolean),
                        _ => None,
                    };
                    if let Some(p) = pushed {
                        d.parts.push(p);
                    }
                }
            },
            Event::Text { text, raw } => {
                if let Some(t) = text_open.as_mut() {
                    if raw {
                        t.push_str(text);
                    } else {
                        t.push_str(&super::cursor::unescape_entities(text));
                    }
                }
            }
            Event::End { name } => match name {
                "style" => {
                    if let Some(c) = cell.take() {
                        out.cells.push(c);
                    }
                    column = None;
                    row = None;
                }
                "default-style" => in_default = false,
                "embedded-text" => {
                    let text = text_open.take().unwrap_or_default();
                    if let (Some(at), Some((_, d))) = (embedded_at, data.as_mut()) {
                        if let Some(Part::Number { embedded, .. }) = d.parts.get_mut(at) {
                            embedded.push((embedded_position, text));
                        }
                    }
                }
                "number" => embedded_at = None,
                "text" | "currency-symbol" | "fill-character" => {
                    if let (Some(t), Some((_, d))) = (text_open.take(), data.as_mut()) {
                        d.parts.push(match name {
                            "text" => Part::Text(t, text_blank.take()),
                            "currency-symbol" => Part::Currency(t, symbol_country.take()),
                            _ => Part::Fill(t),
                        });
                    }
                }
                "number-style" | "percentage-style" | "currency-style" | "date-style"
                | "time-style" | "boolean-style" | "text-style" => {
                    if let Some((n, d)) = data.take() {
                        out.data.insert(n, d);
                    }
                }
                _ => {}
            },
        }
    }
    Ok(())
}

/// `settings.xml`'s gridlines: each sheet's `ShowGrid` under the `Tables`
/// map by its name, and the view's own where it says one outside it.
fn gridlines(text: &str, label: &str, looks: &mut OdsLooks) -> Result<(), Refusal> {
    let part = "settings.xml";
    let mut cursor = Cursor::new(text);
    // The named maps open, `Tables` among them, and the entry being read.
    let mut maps: Vec<String> = Vec::new();
    let mut entry: Option<String> = None;
    let mut item: Option<String> = None;
    while let Some(event) = step(&mut cursor, label, part)? {
        match event {
            Event::Start {
                name: "config-item-map-named",
                attrs,
                empty: false,
            } => maps.push(attr(attrs, "name").unwrap_or_default().into_owned()),
            Event::End {
                name: "config-item-map-named",
            } => {
                maps.pop();
            }
            Event::Start {
                name: "config-item-map-entry",
                attrs,
                empty: false,
            } if maps.last().is_some_and(|m| m == "Tables") => {
                entry = attr(attrs, "name").map(|n| n.into_owned());
            }
            Event::End {
                name: "config-item-map-entry",
            } => {
                if maps.last().is_some_and(|m| m == "Tables") {
                    entry = None;
                }
            }
            Event::Start {
                name: "config-item",
                attrs,
                empty: false,
            } => item = attr(attrs, "name").map(|n| n.into_owned()),
            Event::End {
                name: "config-item",
            } => item = None,
            Event::Text { text, .. } if item.as_deref() == Some("ShowGrid") => {
                let on = text.trim() != "false";
                match &entry {
                    Some(sheet) if maps.last().is_some_and(|m| m == "Tables") => {
                        looks.gridlines.insert(sheet.clone(), on);
                    }
                    _ => {
                        looks.document_gridlines.get_or_insert(on);
                    }
                }
            }
            _ => {}
        }
    }
    Ok(())
}

/// A color a style writes, `#RRGGBB`; `None` for `transparent` and for none.
fn color(props: &Props, key: &str) -> Option<Rgb> {
    props
        .get(key)
        .and_then(|v| v.strip_prefix('#'))
        .and_then(Rgb::from_hex)
}

fn bold(props: &Props) -> bool {
    props
        .get("fo:font-weight")
        .is_some_and(|w| w == "bold" || w.parse::<u32>().is_ok_and(|n| n >= 600))
}

fn italic(props: &Props) -> bool {
    props
        .get("fo:font-style")
        .is_some_and(|s| s == "italic" || s == "oblique")
}

fn size(props: &Props) -> Option<f64> {
    props.get("fo:font-size").and_then(|s| points(s))
}

/// A font's family: its face's, by the name the part declares, or its own.
fn family<'p>(props: &'p Props, faces: &'p HashMap<String, String>) -> Option<&'p str> {
    if let Some(face) = props.get("style:font-name") {
        return Some(faces.get(face).map_or(face.as_str(), String::as_str));
    }
    props.get("fo:font-family").map(String::as_str)
}

/// An edge as a style writes one, `thin solid #000000`: Excel's line by its
/// style and its width, and its color; `None` for no line. The widths of
/// Excel's lines as an OpenDocument file writes them: a keyword, or a length
/// to `hair` below a quarter point, `thin` to a point and a quarter,
/// `medium` to two and a quarter, `thick` past them (*prediction*, held by
/// LibreOffice's save).
fn edge(spec: &str) -> Option<(Line, Option<Rgb>)> {
    let mut width: Option<f64> = None;
    let mut style = "solid";
    let mut rgb = None;
    for token in spec.split_whitespace() {
        match token {
            "none" | "hidden" => return None,
            "thin" => width = Some(0.75),
            "medium" => width = Some(1.75),
            "thick" => width = Some(2.5),
            "solid" | "double" | "dotted" | "dashed" | "dash-dot" | "dash-dot-dot"
            | "fine-dashed" | "dot-dash" | "dot-dot-dash" | "groove" | "ridge" | "inset"
            | "outset" | "double-thin" | "wave" => style = token,
            t if t.starts_with('#') => rgb = Rgb::from_hex(&t[1..]),
            t => {
                if let Some(pt) = points(t) {
                    width = Some(pt);
                }
            }
        }
    }
    let w = width.unwrap_or(0.75);
    if w <= 0.0 {
        return None;
    }
    let heavy = w > 1.25;
    let line = match style {
        "double" | "double-thin" => Line::Double,
        "dotted" => Line::Dotted,
        "fine-dashed" => Line::Hair,
        "dashed" if heavy => Line::MediumDashed,
        "dashed" => Line::Dashed,
        "dash-dot" | "dot-dash" if heavy => Line::MediumDashDot,
        "dash-dot" | "dot-dash" => Line::DashDot,
        "dash-dot-dot" | "dot-dot-dash" if heavy => Line::MediumDashDotDot,
        "dash-dot-dot" | "dot-dot-dash" => Line::DashDotDot,
        _ if w < 0.25 => Line::Hair,
        _ if w <= 1.25 => Line::Thin,
        _ if w <= 2.25 => Line::Medium,
        _ => Line::Thick,
    };
    Some((line, rgb))
}

/// What the record has no word for, said where a style says it.
fn unheld(props: &Props, default: &Props, faces: &HashMap<String, String>) -> bool {
    let said =
        |key: &str, none: &[&str]| props.get(key).is_some_and(|v| !none.contains(&v.as_str()));
    family(props, faces) != family(default, faces)
        || said("style:text-underline-style", &["none"])
        || said("style:text-line-through-style", &["none"])
        || said("style:text-position", &["0%", "0% 100%"])
        || said("style:text-outline", &["false"])
        || said("fo:text-shadow", &["none"])
        || said("style:repeat-content", &["false"])
        || props.get("fo:text-align").is_some_and(|a| a == "justify")
        || said(
            "style:vertical-align",
            &["top", "middle", "bottom", "automatic"],
        )
        || props
            .get("fo:margin-left")
            .and_then(|m| points(m))
            .is_some_and(|m| m > 0.0)
        || said("style:rotation-angle", &["0"])
        || said("style:direction", &["ltr"])
        || said("style:shrink-to-fit", &["false"])
        || props
            .get("style:writing-mode")
            .is_some_and(|w| w.starts_with("rl"))
        || said("style:diagonal-bl-tr", &["none"])
        || said("style:diagonal-tl-br", &["none"])
}

/// A resolved style as the record says it, against the default's resolved
/// properties; its code from its data style.
fn look_of(
    props: &Props,
    default: &Props,
    code: Option<&str>,
    faces: &HashMap<String, String>,
) -> Look {
    let mut look = Look::default();
    if let Some(code) = code {
        look.set_code(code);
    }
    look.style.fill = color(props, "fo:background-color");
    if bold(props) != bold(default) {
        look.style.font.bold = Some(bold(props));
    }
    if italic(props) != italic(default) {
        look.style.font.italic = Some(italic(props));
    }
    if size(props) != size(default) {
        if let Some(points) = size(props) {
            match font_size(points) {
                Some(s) => look.style.font.size = Some(s),
                None => look.unheld = true,
            }
        }
    }
    let window = props
        .get("style:use-window-font-color")
        .is_some_and(|w| w == "true");
    let ink = color(props, "fo:color");
    if !window && ink != color(default, "fo:color") {
        look.style.font.color = ink;
    }
    let repeat = props
        .get("style:repeat-content")
        .is_some_and(|r| r == "true");
    let value_type = props
        .get("style:text-align-source")
        .is_some_and(|s| s == "value-type");
    if !repeat && !value_type {
        look.style.align.horizontal = match props.get("fo:text-align").map(String::as_str) {
            Some("start" | "left") => Some(Horizontal::Left),
            Some("center") => Some(Horizontal::Center),
            Some("end" | "right") => Some(Horizontal::Right),
            _ => None,
        };
    }
    look.style.align.vertical = match props.get("style:vertical-align").map(String::as_str) {
        Some("top") => Some(Vertical::Top),
        Some("middle") => Some(Vertical::Middle),
        _ => None,
    };
    look.style.wrap = props.get("fo:wrap-option").is_some_and(|w| w == "wrap");
    for (i, side) in ["left", "right", "top", "bottom"].iter().enumerate() {
        let spec = props
            .get(&format!("fo:border-{side}"))
            .or_else(|| props.get("fo:border"));
        let Some((line, rgb)) = spec.and_then(|s| edge(s)) else {
            continue;
        };
        let e = Some(Edge::new(line, rgb));
        match i {
            0 => look.style.border.left = e,
            1 => look.style.border.right = e,
            2 => look.style.border.top = e,
            _ => look.style.border.bottom = e,
        }
    }
    look.unheld |= unheld(props, default, faces);
    look
}

// ---- A data style as Excel's code ------------------------------------------

/// The characters a code draws as themselves outside quotes.
const BARE: &str = "$-+():!^&'~{}<>= ";

/// A literal of a code's section, bare where every character draws as
/// itself, quoted otherwise; in a date's or a time's section `/`, `,` and
/// `.` bare too.
fn literal(text: &str, date: bool) -> String {
    let bare = |c: char| BARE.contains(c) || !c.is_ascii() || (date && "/,.".contains(c));
    if !text.is_empty() && text.chars().all(bare) {
        text.to_string()
    } else {
        format!("\"{}\"", text.replace('"', "\\\""))
    }
}

/// Excel's name for a section's ink, where it has one.
fn ink_name(color: &str) -> Option<&'static str> {
    Some(
        match color.trim_start_matches('#').to_ascii_uppercase().as_str() {
            "000000" => "Black",
            "0000FF" => "Blue",
            "00FFFF" => "Cyan",
            "00FF00" => "Green",
            "FF00FF" => "Magenta",
            "FF0000" => "Red",
            "FFFFFF" => "White",
            "FFFF00" => "Yellow",
            _ => return None,
        },
    )
}

/// A number's digits as a code spells them.
fn digits(part: &Part) -> String {
    let Part::Number {
        decimals,
        min_decimals,
        min_int,
        grouping,
        embedded,
        factor,
        blank_int,
    } = part
    else {
        return String::new();
    };
    // The integer's places, the least of them `0`s and the leading of
    // those `?`s, a group's `#,##` before.
    let mut int: Vec<String> = Vec::new();
    let shown = (*min_int).max(1);
    let blank_from = min_int.saturating_sub(*blank_int);
    for k in 0..shown.max(if *grouping { 4 } else { 1 }) {
        int.push(
            if k < blank_from {
                "0"
            } else if k < *min_int {
                "?"
            } else {
                "#"
            }
            .to_string(),
        );
    }
    int.reverse();
    if *grouping {
        let n = int.len();
        int.insert(n - 3, ",".to_string());
    }
    // Each embedded text before the digit it stands before, counted from the
    // last.
    let mut placed: Vec<String> = Vec::new();
    let mut place = 0u32;
    for token in int.iter().rev() {
        if token != "," {
            for (at, text) in embedded {
                if *at == place && place > 0 {
                    placed.push(literal(text, false));
                }
            }
            place += 1;
        }
        placed.push(token.clone());
    }
    placed.reverse();
    let mut s: String = placed.concat();
    for (at, text) in embedded {
        if *at == 0 {
            s.push_str(&literal(text, false));
        }
    }
    let decimals = decimals.unwrap_or(0);
    if decimals > 0 {
        let least = min_decimals.unwrap_or(decimals).min(decimals);
        s.push('.');
        s.push_str(&"0".repeat(least as usize));
        s.push_str(&"#".repeat((decimals - least) as usize));
    }
    if let Some(f) = factor {
        let mut f = *f;
        while f >= 1000.0 {
            s.push(',');
            f /= 1000.0;
        }
    }
    s
}

/// The currency symbol a country's locale names, for an empty symbol: the
/// currencies of the larger economies, and the generic sign `¤` for
/// another country.
fn currency_of(country: Option<&str>) -> &'static str {
    match country.unwrap_or("US").to_ascii_uppercase().as_str() {
        "US" | "CA" | "AU" | "NZ" | "MX" | "SG" => "$",
        "GB" => "£",
        "AT" | "BE" | "CY" | "DE" | "EE" | "ES" | "FI" | "FR" | "GR" | "HR" | "IE" | "IT"
        | "LT" | "LU" | "LV" | "MT" | "NL" | "PT" | "SI" | "SK" => "€",
        "JP" | "CN" => "¥",
        "IN" => "₹",
        "KR" => "₩",
        "BR" => "R$",
        "CH" => "CHF",
        _ => "¤",
    }
}

/// A text as a code spells it, its padding character, where it has one,
/// as Excel's `_` and the character.
fn padded(text: &str, blank: Option<(char, usize)>, date: bool) -> String {
    let Some((c, at)) = blank else {
        return literal(text, date);
    };
    let chars: Vec<char> = text.chars().collect();
    let at = at.min(chars.len().saturating_sub(1));
    let before: String = chars[..at].iter().collect();
    let after: String = chars
        .get(at + 1..)
        .map(|r| r.iter().collect())
        .unwrap_or_default();
    let mut s = String::new();
    if !before.is_empty() {
        s.push_str(&literal(&before, date));
    }
    s.push('_');
    s.push(c);
    if !after.is_empty() {
        s.push_str(&literal(&after, date));
    }
    s
}

/// One data style's own section as a code: its parts in order, its ink
/// before them; `country` the document's, for a currency's empty symbol.
fn section(style: &DataStyle, country: Option<&str>) -> String {
    let date = matches!(style.kind.as_str(), "date-style" | "time-style");
    let mut s = String::new();
    if let Some(name) = style.color.as_deref().and_then(ink_name) {
        s.push_str(&format!("[{name}]"));
    }
    let mut first_hours = true;
    for part in &style.parts {
        match part {
            Part::Number { .. } => s.push_str(&digits(part)),
            Part::Scientific {
                decimals,
                min_decimals,
                min_int,
                min_exponent,
                interval,
                forced,
            } => {
                // The integer's places as many as the exponent's interval,
                // the least of them `0`s.
                let zeros = (*min_int).max(1);
                s.push_str(&"#".repeat((*interval).saturating_sub(zeros) as usize));
                s.push_str(&"0".repeat(zeros as usize));
                if *decimals > 0 {
                    let least = min_decimals.unwrap_or(*decimals).min(*decimals);
                    s.push('.');
                    s.push_str(&"0".repeat(least as usize));
                    s.push_str(&"#".repeat((decimals - least) as usize));
                }
                s.push_str(if *forced { "E+" } else { "E-" });
                s.push_str(&"0".repeat((*min_exponent).max(1) as usize));
            }
            Part::Fraction {
                min_int,
                numerator,
                denominator,
                value,
            } => {
                s.push_str(if *min_int > 0 { "0 " } else { "# " });
                s.push_str(&"?".repeat((*numerator).max(1) as usize));
                s.push('/');
                match value {
                    Some(v) => s.push_str(&v.to_string()),
                    None => s.push_str(&"?".repeat((*denominator).max(1) as usize)),
                }
            }
            Part::Text(t, None) if style.kind == "percentage-style" && t == "%" => s.push('%'),
            Part::Text(t, blank) => s.push_str(&padded(t, *blank, date)),
            Part::Currency(c, own) if c.trim().is_empty() => {
                let named = own.as_deref().or(style.country.as_deref()).or(country);
                s.push_str(&literal(currency_of(named), false));
            }
            Part::Currency(c, _) => s.push_str(&literal(c, false)),
            Part::Day(long) => s.push_str(if *long { "dd" } else { "d" }),
            Part::Month { long, textual } => s.push_str(match (textual, long) {
                (true, true) => "mmmm",
                (true, false) => "mmm",
                (false, true) => "mm",
                (false, false) => "m",
            }),
            Part::Year(long) => s.push_str(if *long { "yyyy" } else { "yy" }),
            Part::DayOfWeek(long) => s.push_str(if *long { "dddd" } else { "ddd" }),
            Part::Hours(long) => {
                let h = if *long { "hh" } else { "h" };
                if style.elapsed && first_hours {
                    s.push_str(&format!("[{h}]"));
                } else {
                    s.push_str(h);
                }
                first_hours = false;
            }
            Part::Minutes(long) => s.push_str(if *long { "mm" } else { "m" }),
            Part::Seconds { long, decimals } => {
                s.push_str(if *long { "ss" } else { "s" });
                if *decimals > 0 {
                    s.push('.');
                    s.push_str(&"0".repeat(*decimals as usize));
                }
            }
            Part::AmPm => s.push_str("AM/PM"),
            Part::TextContent => s.push('@'),
            Part::Boolean => {}
            Part::Fill(c) => {
                s.push('*');
                s.push_str(c);
            }
        }
    }
    s
}

/// Whether a data style is General: one number with no decimals said, no
/// grouping, no scale, no text and no ink.
fn is_general(style: &DataStyle) -> bool {
    if style.kind == "boolean-style" {
        return true;
    }
    style.maps.is_empty()
        && style.color.is_none()
        && matches!(
            style.parts.as_slice(),
            [Part::Number {
                decimals: None,
                grouping: false,
                factor: None,
                embedded,
                min_int: 0 | 1,
                ..
            }] if embedded.is_empty()
        )
}

/// A data style as a code in Excel's syntax, its conditions as Excel's
/// sections: a positive's and zero's section before the style's own,
/// `value()>=0`; a positive's, a negative's and zero's before a text's own,
/// `value()>0`, `value()<0` and `value()=0`; another condition in brackets.
/// `None` for General, and for a style the file does not define; `country`
/// the document's.
fn code(styles: &HashMap<String, DataStyle>, name: &str, country: Option<&str>) -> Option<String> {
    let style = styles.get(name)?;
    if is_general(style) {
        return None;
    }
    let own = section(style, country);
    if style.maps.is_empty() {
        return Some(own);
    }
    let mut sections: Vec<String> = Vec::new();
    for (condition, applied) in &style.maps {
        let applied = styles
            .get(applied)
            .map(|s| section(s, country))
            .unwrap_or_default();
        let condition = condition.replace(' ', "");
        let condition = condition.strip_prefix("value()").unwrap_or(&condition);
        let standard = matches!(
            (sections.len(), condition),
            (0, ">=0") | (0, ">0") | (1, "<0") | (2, "=0")
        );
        if standard {
            sections.push(applied);
        } else {
            sections.push(format!("[{condition}]{applied}"));
        }
    }
    sections.push(own);
    Some(sections.join(";"))
}

/// An OpenDocument file's looks, read from its parts: `content.xml`, held
/// whole, `styles.xml` and `settings.xml` where it has them.
pub(crate) fn read(
    content: &str,
    styles: Option<&str>,
    settings: Option<&str>,
    label: &str,
) -> Result<OdsLooks, Refusal> {
    let mut named = Declared::default();
    if let Some(s) = styles {
        declared(s, label, "styles.xml", &mut named)?;
    }
    let mut automatic = Declared::default();
    declared(content, label, "content.xml", &mut automatic)?;
    // The faces and the data styles of both parts, content's nearer.
    let mut faces = named.faces.clone();
    faces.extend(automatic.faces.clone());
    let mut data = named.data.clone();
    data.extend(automatic.data.clone());
    // Every cell style by its name, and the formats in their order: Default
    // first, the automatic styles, then the named ones.
    let mut by_name: HashMap<String, CellStyle> = HashMap::new();
    for (n, s) in named.cells.iter().chain(automatic.cells.iter()) {
        by_name.insert(n.clone(), s.clone());
    }
    let mut order: Vec<String> = vec!["Default".to_string()];
    for (n, _) in automatic.cells.iter().chain(named.cells.iter()) {
        if n != "Default" && !order.contains(n) {
            order.push(n.clone());
        }
    }
    let mut default_style = named.default_style.clone();
    default_style.extend(automatic.default_style.clone());
    // LibreOffice 26.8 writes its family's default wrap on where its own
    // drawing wraps none (its save of `looks.xlsx`: its thumbnail and its
    // fitted heights); the wrap a style says is its own.
    default_style.remove("fo:wrap-option");
    let country = default_style.get("fo:country").cloned();
    // A style's properties and data style, up its chain to the default.
    let resolve = |name: &str| -> (Props, Option<String>) {
        let mut chain: Vec<&CellStyle> = Vec::new();
        let mut at = Some(name.to_string());
        while let Some(n) = at {
            let Some(s) = by_name.get(&n) else { break };
            if chain.len() > 64 {
                break;
            }
            chain.push(s);
            at = s.parent.clone();
        }
        let mut props = default_style.clone();
        let mut data_style: Option<String> = None;
        if name != "Default" {
            if let Some(d) = by_name.get("Default") {
                props.extend(d.props.clone());
                data_style = d.data.clone();
            }
        }
        for s in chain.iter().rev() {
            props.extend(s.props.clone());
            if s.data.is_some() {
                data_style = s.data.clone();
            }
        }
        (props, data_style)
    };
    let (default_props, _) = resolve("Default");
    let mut looks = OdsLooks {
        columns: automatic.columns,
        rows: automatic.rows,
        ..OdsLooks::default()
    };
    for (i, n) in order.iter().enumerate() {
        let (props, data_style) = resolve(n);
        let code = data_style
            .as_deref()
            .and_then(|d| code(&data, d, country.as_deref()));
        let mut look = look_of(&props, &default_props, code.as_deref(), &faces);
        if i == 0 {
            // The default says nothing of itself.
            look = Look {
                style: Default::default(),
                ..look
            };
            if let Some(code) = &code {
                look.set_code(code);
            }
        }
        looks.looks.push(look);
        looks.index.insert(n.clone(), i as u32);
        looks.rules.push(by_name.get(n).map_or(0, |s| s.rules));
    }
    if let Some(s) = settings {
        gridlines(s, label, &mut looks)?;
    }
    Ok(looks)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn data(xml: &str) -> HashMap<String, DataStyle> {
        let mut d = Declared::default();
        declared(
            &format!("<office:automatic-styles>{xml}</office:automatic-styles>"),
            "t.ods",
            "content.xml",
            &mut d,
        )
        .unwrap();
        d.data
    }

    #[test]
    fn a_data_style_is_a_code_in_excel_s_syntax() {
        // Excel's own .ods save's data styles (the item's fixture) and
        // LibreOffice's spellings beside them.
        let styles = data(concat!(
            "<number:number-style style:name=\"N0\"><number:number number:min-integer-digits=\"1\"/></number:number-style>",
            "<number:number-style style:name=\"N2\"><number:number number:decimal-places=\"2\" number:min-decimal-places=\"2\" number:min-integer-digits=\"1\"/></number:number-style>",
            "<number:number-style style:name=\"N3\"><number:number number:decimal-places=\"0\" number:min-integer-digits=\"1\" number:grouping=\"true\"/></number:number-style>",
            "<number:percentage-style style:name=\"N36\"><number:number number:decimal-places=\"1\" number:min-decimal-places=\"1\" number:min-integer-digits=\"1\"/><number:text>%</number:text></number:percentage-style>",
            "<number:currency-style style:name=\"N37\"><number:currency-symbol>$</number:currency-symbol><number:number number:decimal-places=\"2\" number:min-decimal-places=\"2\" number:min-integer-digits=\"1\" number:grouping=\"true\"/></number:currency-style>",
            "<number:date-style style:name=\"N19\"><number:month/><number:text>/</number:text><number:day/><number:text>/</number:text><number:year number:style=\"long\"/></number:date-style>",
            "<number:date-style style:name=\"N39\"><number:month number:style=\"long\" number:textual=\"true\"/><number:text> </number:text><number:day/><number:text>, </number:text><number:year number:style=\"long\"/></number:date-style>",
            "<number:time-style style:name=\"N40\" number:truncate-on-overflow=\"false\"><number:hours/><number:text>:</number:text><number:minutes number:style=\"long\"/></number:time-style>",
            "<number:number-style style:name=\"N41P0\"><number:number number:decimal-places=\"0\" number:min-integer-digits=\"1\"/></number:number-style>",
            "<number:number-style style:name=\"N41P1\"><number:text>-</number:text><number:number number:decimal-places=\"0\" number:min-integer-digits=\"1\"/></number:number-style>",
            "<number:number-style style:name=\"N41P2\"/>",
            "<number:text-style style:name=\"N41\"><number:text-content/><style:map style:condition=\"value()&gt;0\" style:apply-style-name=\"N41P0\"/><style:map style:condition=\"value()&lt;0\" style:apply-style-name=\"N41P1\"/><style:map style:condition=\"value()=0\" style:apply-style-name=\"N41P2\"/></number:text-style>",
            "<number:number-style style:name=\"N42\"><number:text>&lt;</number:text><number:number number:decimal-places=\"0\" number:min-decimal-places=\"0\" number:min-integer-digits=\"1\"><number:embedded-text number:position=\"0\">&gt;</number:embedded-text></number:number></number:number-style>",
            "<number:number-style style:name=\"N43\"><style:text-properties fo:color=\"#FF0000\"/><number:number number:decimal-places=\"2\" number:min-decimal-places=\"2\" number:min-integer-digits=\"1\"/></number:number-style>",
            "<number:number-style style:name=\"N44P0\"><number:number number:decimal-places=\"0\" number:min-integer-digits=\"1\"/></number:number-style>",
            "<number:number-style style:name=\"N44\"><number:number number:decimal-places=\"2\" number:min-decimal-places=\"2\" number:min-integer-digits=\"1\"/><style:map style:condition=\"value()&gt;=100\" style:apply-style-name=\"N44P0\"/></number:number-style>",
            "<number:number-style style:name=\"N15\"><number:scientific-number number:decimal-places=\"2\" number:min-decimal-places=\"2\" number:min-integer-digits=\"1\" number:min-exponent-digits=\"2\"/></number:number-style>",
            "<number:number-style style:name=\"N16\"><number:scientific-number number:decimal-places=\"1\" number:min-decimal-places=\"1\" number:min-integer-digits=\"1\" number:min-exponent-digits=\"1\" number:exponent-interval=\"3\" number:forced-exponent-sign=\"true\"/></number:number-style>",
            "<number:number-style style:name=\"N48\"><number:scientific-number number:decimal-places=\"3\" number:min-decimal-places=\"1\" number:min-integer-digits=\"1\" number:min-exponent-digits=\"2\" number:forced-exponent-sign=\"false\"/></number:number-style>",
            "<number:number-style style:name=\"N17\"><number:fraction number:min-integer-digits=\"0\" number:min-numerator-digits=\"1\" number:min-denominator-digits=\"1\"/></number:number-style>",
            "<number:number-style style:name=\"N47\"><number:number number:decimal-places=\"0\" number:min-integer-digits=\"1\"/><number:fill-character>-</number:fill-character></number:number-style>",
            "<number:text-style style:name=\"N30\"><number:text-content/></number:text-style>",
            "<number:date-style style:name=\"N38\"><number:year number:style=\"long\"/><number:text>-</number:text><number:month number:style=\"long\"/><number:text>-</number:text><number:day number:style=\"long\"/><number:text> </number:text><number:hours number:style=\"long\"/><number:text>:</number:text><number:minutes number:style=\"long\"/><number:text>:</number:text><number:seconds number:style=\"long\"/></number:date-style>",
            "<number:boolean-style style:name=\"N99\"><number:boolean/></number:boolean-style>",
            // LibreOffice's own save of the fixture: a padding of a
            // character's width, a digit's space and an empty currency
            // symbol, the locale's.
            "<number:currency-style style:name=\"N129P0\" style:volatile=\"true\"><number:currency-symbol/><number:number number:decimal-places=\"0\" number:min-decimal-places=\"0\" number:min-integer-digits=\"1\" number:grouping=\"true\"/><number:text loext:blank-width-char=\")\"> </number:text></number:currency-style>",
            "<number:currency-style style:name=\"N129\"><number:text>(</number:text><number:currency-symbol/><number:number number:decimal-places=\"0\" number:min-decimal-places=\"0\" number:min-integer-digits=\"1\" number:grouping=\"true\"/><number:text>)</number:text><style:map style:condition=\"value()&gt;=0\" style:apply-style-name=\"N129P0\"/></number:currency-style>",
            "<number:number-style style:name=\"N146P2\" style:volatile=\"true\"><number:text loext:blank-width-char=\"(\"> </number:text><number:fill-character> </number:fill-character><number:text>-</number:text><number:number number:decimal-places=\"0\" number:min-decimal-places=\"0\" number:min-integer-digits=\"2\" loext:max-blank-integer-digits=\"2\"/><number:text loext:blank-width-char=\")\"> </number:text></number:number-style>",
            "<number:text-style style:name=\"N144P2\"><number:text loext:blank-width-char=\")1\">- </number:text></number:text-style>",
            "<number:currency-style style:name=\"N150\" number:language=\"de\" number:country=\"DE\"><number:number number:decimal-places=\"2\" number:min-decimal-places=\"2\" number:min-integer-digits=\"1\" number:grouping=\"true\"/><number:text> </number:text><number:currency-symbol/></number:currency-style>",
        ));
        for (name, want) in [
            ("N0", None),
            ("N2", Some("0.00")),
            ("N3", Some("#,##0")),
            ("N36", Some("0.0%")),
            ("N37", Some("$#,##0.00")),
            ("N19", Some("m/d/yyyy")),
            ("N39", Some("mmmm d, yyyy")),
            ("N40", Some("[h]:mm")),
            ("N41", Some("0;-0;;@")),
            ("N42", Some("<0>")),
            ("N43", Some("[Red]0.00")),
            ("N44", Some("[>=100]0;0.00")),
            ("N15", Some("0.00E+00")),
            ("N16", Some("##0.0E+0")),
            ("N48", Some("0.0##E-00")),
            ("N17", Some("# ?/?")),
            ("N47", Some("0*-")),
            ("N30", Some("@")),
            ("N38", Some("yyyy-mm-dd hh:mm:ss")),
            ("N99", None),
            ("N129P0", Some("$#,##0_)")),
            ("N129", Some("$#,##0_);($#,##0)")),
            ("N146P2", Some("_(* -??_)")),
            ("N144P2", Some("-_)")),
            ("N150", Some("#,##0.00 €")),
        ] {
            assert_eq!(code(&styles, name, Some("US")).as_deref(), want, "{name}");
        }
        assert_eq!(code(&styles, "N404", None), None, "a style the file lacks");
    }

    #[test]
    fn a_length_and_a_width_read_as_the_record_says_them() {
        assert_eq!(points("0.74pt"), Some(0.74));
        assert!((points("2.54cm").unwrap() - 72.0).abs() < 1e-9);
        assert_eq!(points("1in"), Some(72.0));
        assert_eq!(points("12px"), Some(9.0));
        assert_eq!(column_width("1.69333333333333cm"), None, "64 pixels");
        assert_eq!(column_width("2.06375cm"), Some(10.43), "78 pixels");
        assert_eq!(column_width("0cm"), Some(0.0));
        let mut props = Props::new();
        props.insert("style:row-height".into(), "1.0583cm".into());
        assert_eq!(row_height(&props), Some(30.0));
        props.insert("style:use-optimal-row-height".into(), "true".into());
        assert_eq!(row_height(&props), None, "a height fitted");
    }

    #[test]
    fn an_edge_reads_as_excel_s_line() {
        assert_eq!(
            edge("thin solid #000000"),
            Some((Line::Thin, Rgb::from_hex("000000")))
        );
        assert_eq!(edge("2pt solid #FF0000").map(|e| e.0), Some(Line::Medium));
        assert_eq!(edge("0.74pt solid #000000").map(|e| e.0), Some(Line::Thin));
        assert_eq!(edge("2.49pt solid #000000").map(|e| e.0), Some(Line::Thick));
        assert_eq!(
            edge("1.1pt double #000000").map(|e| e.0),
            Some(Line::Double)
        );
        assert_eq!(
            edge("0.74pt dashed #000000").map(|e| e.0),
            Some(Line::Dashed)
        );
        assert_eq!(
            edge("1.76pt dashed #000000").map(|e| e.0),
            Some(Line::MediumDashed)
        );
        assert_eq!(edge("none"), None);
        assert_eq!(edge("0pt solid #000000"), None);
    }
}
