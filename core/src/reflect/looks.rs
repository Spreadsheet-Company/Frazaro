//! A workbook's looks read as the record's rows (KERNEL.57): what a reader
//! shares when `frazaro reflect --looks` prints a workbook as the record
//! prints a sheet whole (`vla_lang::view`, oracle 11). A color a file names,
//! resolved to the record's six hex digits through the workbook's theme and
//! palette; Excel's built-in number formats, read as the owner's English
//! Excel reads them; a cell format as the record says it ([`Look`]); a
//! column's width as Excel reads one; an ISO date as Excel's serial; and
//! what the record cannot hold, counted and said ([`LooksStats`]).
//!
//! A format prints what it changes: a property the workbook's default font
//! has is unsaid, and so is Excel's automatic color. What the record has no
//! word for, a font's family, an underline, a pattern, a rotation, a
//! diagonal, is left out of the row and counted, the rest of the format
//! printed. A number format the language does not draw is printed as the
//! file holds it ([`Look::undrawn`]), so that a load refuses it by name
//! (`grid-format-code`) and a later language that draws more reads the same
//! rows further.
//!
//! Each rule here is held to Excel's reading of two hand-made workbooks
//! (the item's second slice, `looks_golden.rs`): a theme's first two pairs
//! of slots counted swapped from the order a theme part lists them
//! ([`THEME_SLOTS`]); a tint by ECMA-376's rule as Excel computes it, in
//! Windows' HLS ([`tinted`]); Excel's default palette ([`PALETTE`]), the
//! system's paper white; the built-ins as the owner's English Excel reads
//! them, a locale's id as its twin ([`builtin_code`], [`builtin_twin`]); a
//! width as Excel's rule counts it at a screen's own pixels; and a solid
//! fill of Excel's automatic color drawn white.

use vla_lang::numfmt::Code;
use vla_lang::sheet::{NumFmt, Rgb, Style};

/// How a file names a color, before the workbook's theme and palette
/// resolve it.
#[derive(Clone, Copy, Debug, PartialEq)]
pub enum Color {
    /// Six hex digits, an alpha written before them passed over.
    Rgb(Rgb),
    /// A slot of the workbook's theme, by Excel's count ([`THEME_SLOTS`]).
    Theme(u32),
    /// An index of the workbook's palette, Excel's own ([`PALETTE`]) where
    /// it has none of its own.
    Indexed(u32),
    /// Excel's automatic color.
    Auto,
}

/// A color as a file names it: its kind and its tint, from -1, darkened
/// whole, to 1, lightened whole, 0 for none.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct ColorRef {
    pub color: Color,
    pub tint: f64,
}

/// A color resolved: six hex digits, Excel's automatic, or a reference the
/// workbook cannot resolve (a theme's slot in a workbook with no theme, an
/// index past its palette).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Resolved {
    Rgb(Rgb),
    Automatic,
    Unresolved,
}

/// A theme's slots in the order Excel counts them, a `theme` attribute's
/// index: the light and the dark of the first two pairs swapped from the
/// order a theme part lists them, `dk1` first (Excel's reading: a file's
/// slots 0 to 3 drawn white, black, `lt2` and `dk2`). Excel's object model
/// numbers them from 1 in the same order, its names, `xlThemeColorDark1`
/// for the file's slot 0, the theme part's order.
pub const THEME_SLOTS: [&str; 12] = [
    "lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
    "hlink", "folHlink",
];

/// Excel's default palette of 64 colors by index (ECMA-376 Part 1,
/// §18.8.27), the colors an `indexed` attribute names where a workbook has
/// no palette of its own; 64 and 65, past it, are the system's ink, Excel's
/// automatic, and its paper, which Excel draws white.
pub const PALETTE: [u32; 64] = [
    0x000000, 0xFFFFFF, 0xFF0000, 0x00FF00, 0x0000FF, 0xFFFF00, 0xFF00FF, 0x00FFFF, //
    0x000000, 0xFFFFFF, 0xFF0000, 0x00FF00, 0x0000FF, 0xFFFF00, 0xFF00FF, 0x00FFFF, //
    0x800000, 0x008000, 0x000080, 0x808000, 0x800080, 0x008080, 0xC0C0C0, 0x808080, //
    0x9999FF, 0x993366, 0xFFFFCC, 0xCCFFFF, 0x660066, 0xFF8080, 0x0066CC, 0xCCCCFF, //
    0x000080, 0xFF00FF, 0xFFFF00, 0x00FFFF, 0x800080, 0x800000, 0x008080, 0x0000FF, //
    0x00CCFF, 0xCCFFFF, 0xCCFFCC, 0xFFFF99, 0x99CCFF, 0xFF99CC, 0xCC99FF, 0xFFCC99, //
    0x3366FF, 0x33CCCC, 0x99CC00, 0xFFCC00, 0xFF9900, 0xFF6600, 0x666699, 0x969696, //
    0x003366, 0x339966, 0x003300, 0x333300, 0x993300, 0x993366, 0x333399, 0x333333, //
];

fn rgb_of(word: u32) -> Rgb {
    Rgb((word >> 16) as u8, (word >> 8) as u8, word as u8)
}

/// A color's six or eight hex digits, an alpha before them passed over.
pub fn rgb_from_hex(text: &str) -> Option<Rgb> {
    let text = text.trim();
    match text.len() {
        6 => Rgb::from_hex(text),
        8 => Rgb::from_hex(&text[2..]),
        _ => None,
    }
}

/// What resolves a workbook's colors: its theme's twelve slots in Excel's
/// count, where it has a theme, and its palette.
#[derive(Clone, Debug)]
pub struct Colors {
    theme: Option<[Option<Rgb>; 12]>,
    palette: Vec<Rgb>,
}

impl Colors {
    /// A workbook's colors: its theme's slots by Excel's count, if it has a
    /// theme, and its own palette, if it has one, Excel's otherwise.
    pub fn new(theme: Option<[Option<Rgb>; 12]>, palette: Option<Vec<Rgb>>) -> Colors {
        Colors {
            theme,
            palette: palette.unwrap_or_else(|| PALETTE.iter().map(|w| rgb_of(*w)).collect()),
        }
    }

    /// A color resolved, its tint applied.
    pub fn resolve(&self, c: &ColorRef) -> Resolved {
        let base = match c.color {
            Color::Auto => return Resolved::Automatic,
            Color::Rgb(rgb) => rgb,
            Color::Theme(slot) => match self
                .theme
                .as_ref()
                .and_then(|t| t.get(slot as usize).copied().flatten())
            {
                Some(rgb) => rgb,
                None => return Resolved::Unresolved,
            },
            Color::Indexed(64) => return Resolved::Automatic,
            Color::Indexed(65) => Rgb(255, 255, 255),
            Color::Indexed(i) => match self.palette.get(i as usize) {
                Some(rgb) => *rgb,
                None => return Resolved::Unresolved,
            },
        };
        Resolved::Rgb(tinted(base, c.tint))
    }
}

/// The range of a hue, a luminance and a saturation in Windows' HLS, which
/// Excel's tint is applied in.
const HLS_MAX: i32 = 240;

/// A color with its luminance moved by a tint, as Excel draws it: the tint
/// kept in 32767ths, as Excel keeps it (its save writes 0.1249 back as
/// 4092/32767), and ECMA-376's rule (Part 1, §18.8.19) applied in Windows'
/// HLS, each to 240 (`ColorRGBToHLS`): a darkening to `L·(1 + tint)`, a
/// lightening to `L·(1 − tint) + (240 − 240·(1 − tint))`, each product
/// truncated to a whole luminance, and back (`ColorHLSToRGB`), a gray to its
/// nearest channel. Held to Excel's reading of 90 tints to the channel,
/// every slot of a theme at its palette's tints and at others (the item's
/// golden), where the rule in exact HLS rounded at the end gives 32, and
/// truncated once at the end 70.
pub fn tinted(rgb: Rgb, tint: f64) -> Rgb {
    if !tint.is_finite() {
        return rgb;
    }
    let kept = (tint.clamp(-1.0, 1.0) * 32767.0).trunc() / 32767.0;
    if kept == 0.0 {
        return rgb;
    }
    let (h, l, s) = to_hls(rgb);
    let (lum, most) = (f64::from(l), f64::from(HLS_MAX));
    let l = if kept < 0.0 {
        (lum * (1.0 + kept)).trunc() as i32
    } else {
        (lum * (1.0 - kept)).trunc() as i32 + HLS_MAX - (most * (1.0 - kept)).trunc() as i32
    };
    from_hls(h, l.clamp(0, HLS_MAX), s)
}

/// A color in Windows' HLS, each to 240, as its `ColorRGBToHLS` counts one
/// (the algorithm of Microsoft's KB29240, integers throughout).
fn to_hls(Rgb(r, g, b): Rgb) -> (i32, i32, i32) {
    let (r, g, b) = (i32::from(r), i32::from(g), i32::from(b));
    let (max, min) = (r.max(g).max(b), r.min(g).min(b));
    let l = ((max + min) * HLS_MAX + 255) / (2 * 255);
    if max == min {
        return (0, l, 0);
    }
    let (sum, range) = (max + min, max - min);
    let s = if l <= HLS_MAX / 2 {
        (range * HLS_MAX + sum / 2) / sum
    } else {
        (range * HLS_MAX + (2 * 255 - sum) / 2) / (2 * 255 - sum)
    };
    let part = |c: i32| ((max - c) * (HLS_MAX / 6) + range / 2) / range;
    let mut h = if r == max {
        part(b) - part(g)
    } else if g == max {
        HLS_MAX / 3 + part(r) - part(b)
    } else {
        2 * HLS_MAX / 3 + part(g) - part(r)
    };
    if h < 0 {
        h += HLS_MAX;
    }
    if h > HLS_MAX {
        h -= HLS_MAX;
    }
    (h, l, s)
}

/// A color of Windows' HLS back to its channels, as its `ColorHLSToRGB`
/// draws one, a gray rounded to its channel (Excel's reading).
fn from_hls(h: i32, l: i32, s: i32) -> Rgb {
    let channel = |v: i32| v.clamp(0, 255) as u8;
    if s == 0 {
        let v = channel((l * 255 + HLS_MAX / 2) / HLS_MAX);
        return Rgb(v, v, v);
    }
    let m2 = if l <= HLS_MAX / 2 {
        (l * (HLS_MAX + s) + HLS_MAX / 2) / HLS_MAX
    } else {
        l + s - (l * s + HLS_MAX / 2) / HLS_MAX
    };
    let m1 = 2 * l - m2;
    let part = |hue: i32| {
        let hue = if hue < 0 {
            hue + HLS_MAX
        } else if hue > HLS_MAX {
            hue - HLS_MAX
        } else {
            hue
        };
        let v = if hue < HLS_MAX / 6 {
            m1 + ((m2 - m1) * hue + HLS_MAX / 12) / (HLS_MAX / 6)
        } else if hue < HLS_MAX / 2 {
            m2
        } else if hue < 2 * HLS_MAX / 3 {
            m1 + ((m2 - m1) * (2 * HLS_MAX / 3 - hue) + HLS_MAX / 12) / (HLS_MAX / 6)
        } else {
            m1
        };
        channel((v * 255 + HLS_MAX / 2) / HLS_MAX)
    };
    Rgb(part(h + HLS_MAX / 3), part(h), part(h - HLS_MAX / 3))
}

/// A built-in number format's code by its id, for a `numFmtId` a styles
/// part does not define, as the owner's English Excel reads it (Excel's
/// reading of every id from 0 to 81, the item's golden): ECMA-376's list
/// (Part 1, §18.8.30) for the ids Excel draws alike everywhere; Excel's
/// `en-US` codes for 37 to 40, where ECMA's are a space short of Excel's
/// `_)`; for 5 to 8 and 41 to 44, the locale's own, the codes Excel's save
/// writes for them. `None` for an id with no code of its own, a locale's
/// id among them, which [`builtin_twin`] reads first.
pub fn builtin_code(id: u32) -> Option<&'static str> {
    Some(match id {
        0 => "General",
        1 => "0",
        2 => "0.00",
        3 => "#,##0",
        4 => "#,##0.00",
        5 => "\"$\"#,##0_);\\(\"$\"#,##0\\)",
        6 => "\"$\"#,##0_);[Red]\\(\"$\"#,##0\\)",
        7 => "\"$\"#,##0.00_);\\(\"$\"#,##0.00\\)",
        8 => "\"$\"#,##0.00_);[Red]\\(\"$\"#,##0.00\\)",
        9 => "0%",
        10 => "0.00%",
        11 => "0.00E+00",
        12 => "# ?/?",
        13 => "# ??/??",
        14 => "m/d/yyyy",
        15 => "d-mmm-yy",
        16 => "d-mmm",
        17 => "mmm-yy",
        18 => "h:mm AM/PM",
        19 => "h:mm:ss AM/PM",
        20 => "h:mm",
        21 => "h:mm:ss",
        22 => "m/d/yyyy h:mm",
        37 => "#,##0_);(#,##0)",
        38 => "#,##0_);[Red](#,##0)",
        39 => "#,##0.00_);(#,##0.00)",
        40 => "#,##0.00_);[Red](#,##0.00)",
        41 => "_(* #,##0_);_(* \\(#,##0\\);_(* \"-\"_);_(@_)",
        42 => "_(\"$\"* #,##0_);_(\"$\"* \\(#,##0\\);_(\"$\"* \"-\"_);_(@_)",
        43 => "_(* #,##0.00_);_(* \\(#,##0.00\\);_(* \"-\"??_);_(@_)",
        44 => "_(\"$\"* #,##0.00_);_(\"$\"* \\(#,##0.00\\);_(\"$\"* \"-\"??_);_(@_)",
        45 => "mm:ss",
        46 => "[h]:mm:ss",
        47 => "mm:ss.0",
        48 => "##0.0E+0",
        49 => "@",
        _ => return None,
    })
}

/// The id Excel reads a built-in id as, where a styles part defines it
/// not: an id Excel draws alike everywhere itself; a locale's id, the
/// East Asian dates and times from 27 to 36 and 50 to 58 and the Thai ids
/// from 59 to 81, its twin among those (14's date, 21's time, the numbers,
/// the currencies, the percents, the fractions and the clocks), as Excel's
/// save writes each; and any other, 23 to 26 and past 81, General, 0.
pub fn builtin_twin(id: u32) -> u32 {
    match id {
        0..=22 | 37..=49 => id,
        27..=31 | 36 | 50..=58 | 71 | 72 => 14,
        32..=35 | 77 => 21,
        59..=68 | 73..=75 => id - 58,
        69 | 70 => id - 57,
        76 => 20,
        78 => 22,
        79..=81 => id - 34,
        _ => 0,
    }
}

/// A cell format as the record says it, and what its file says beside it
/// that the record cannot.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct Look {
    /// The format as the record spells it, what the workbook's defaults say
    /// left unsaid.
    pub style: Style,
    /// A number format the file holds that the language does not draw,
    /// printed as written in place of the style's own.
    pub undrawn: Option<String>,
    /// Whether the file's format says what the record has no word for.
    pub unheld: bool,
}

impl Look {
    /// The number format a file's code is: `general` and `text` for
    /// Excel's names of them, `General` and `@`, which the record spells
    /// as those words; a code the language draws; or, where it draws none,
    /// the code as written.
    pub fn set_code(&mut self, text: &str) {
        if text.eq_ignore_ascii_case("General") {
            self.style.num_fmt = NumFmt::General;
        } else if text == "@" {
            self.style.num_fmt = NumFmt::Text;
        } else {
            match Code::parse(text) {
                Ok(code) => self.style.num_fmt = NumFmt::Code(code),
                Err(_) => {
                    self.style.num_fmt = NumFmt::General;
                    self.undrawn = Some(text.to_string());
                }
            }
        }
    }
}

/// A font's size in points as a file holds it, as the record's twentieths;
/// `None` for one that is no whole twentieth or out of the record's bounds,
/// 1 to 409.55 points. A save's `9.3000000000000007` is 9.3.
pub fn font_size(points: f64) -> Option<vla_lang::sheet::FontSize> {
    let twentieths = (points * 20.0).round();
    if !points.is_finite() || (points * 20.0 - twentieths).abs() > 1e-6 {
        return None;
    }
    u16::try_from(twentieths as i64)
        .ok()
        .and_then(vla_lang::sheet::FontSize::from_twentieths)
}

/// The pixels a column of a file's width draws, by Excel's rule at the
/// record's 7 pixels a character (ECMA-376 Part 1, §18.3.1.13):
/// `Truncate(((256·w + Truncate(128/7))/256)·7)`.
pub fn width_pixels(width: f64) -> f64 {
    ((256.0 * width + 18.0) / 256.0 * 7.0).trunc()
}

/// A column's width as a file holds it, Excel's characters with Excel's
/// padding, read as Excel reads one: to whole pixels, then to characters in
/// hundredths, `Truncate((px − 5)/7·100 + 0.5)/100`, and a column narrower
/// than one character, 12 pixels, a twelfth a pixel (*prediction*); `None`
/// for a width that draws Excel's default 64 pixels, 8.43 characters, the
/// record's `none`. The writer's 72.7109375 and Excel 365's 72.7265625 are
/// both 509 pixels, 72.
pub fn width_chars(width: f64) -> Option<f64> {
    let px = width_pixels(width).max(0.0);
    if px == 64.0 {
        return None;
    }
    let chars = if px < 12.0 {
        (px / 12.0 * 100.0 + 0.5).trunc() / 100.0
    } else {
        ((px - 5.0) / 7.0 * 100.0 + 0.5).trunc() / 100.0
    };
    Some(chars)
}

/// An ISO 8601 date, `2026-10-03`, with or without a time after a `T`, as
/// Excel's serial in its 1900 calendar, or its 1904 one where the workbook
/// counts from 1904: the language's calendar (`calc::dates::serial_of`), the
/// time its fraction of a day. `None` for a text that is no such date or a
/// date before the calendar's first year or after 9999.
pub fn iso_serial(text: &str, date1904: bool) -> Option<f64> {
    let text = text.trim();
    let (date, time) = match text.split_once('T') {
        Some((d, t)) => (d, Some(t)),
        None => (text, None),
    };
    let mut parts = date.split('-');
    let (y, m, d) = (parts.next()?, parts.next()?, parts.next()?);
    if parts.next().is_some() || y.len() != 4 || m.len() != 2 || d.len() != 2 {
        return None;
    }
    let digits = |s: &str| s.bytes().all(|b| b.is_ascii_digit());
    if !digits(y) || !digits(m) || !digits(d) {
        return None;
    }
    let (y, m, d): (i64, i64, i64) = (y.parse().ok()?, m.parse().ok()?, d.parse().ok()?);
    let first = if date1904 { 1904 } else { 1900 };
    if !(first..=9999).contains(&y) || !(1..=12).contains(&m) || !(1..=31).contains(&d) {
        return None;
    }
    let mut serial = vla_lang::calc::dates::serial_of(y, m, d) as f64;
    if date1904 {
        serial -= 1462.0;
    }
    if let Some(time) = time {
        let time = time.trim_end_matches('Z');
        let mut hms = time.split(':');
        let h: f64 = hms.next()?.parse().ok()?;
        let mi: f64 = hms.next().unwrap_or("0").parse().ok()?;
        let s: f64 = hms.next().unwrap_or("0").parse().ok()?;
        if hms.next().is_some() || !(0.0..24.0).contains(&h) || !(0.0..60.0).contains(&mi) {
            return None;
        }
        serial += (h * 3600.0 + mi * 60.0 + s) / 86_400.0;
    }
    Some(serial)
}

/// What a sheet's looks hold, counted (`--looks --counts`): the formats its
/// rows print, its cells with a format other than 0, and what the record
/// cannot hold, each by its kind. Counts alone, AXM.1's discipline.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct LooksStats {
    /// The `format` rows the sheet's record prints, format 0 among them.
    pub formats: u64,
    /// The cells printed with a format other than 0.
    pub styled: u64,
    /// The cells that keep a format and no value, which the record cannot
    /// hold, left out.
    pub formatted_empty: u64,
    /// The cells whose format or text says what the record has no word for,
    /// and each default of the sheet's own, its width or height, the record
    /// cannot say: left out of their rows.
    pub unheld: u64,
    /// The cells whose number format the language does not draw, printed as
    /// the file holds it, which a load refuses by name.
    pub undrawn: u64,
    /// The merged ranges, which no row says.
    pub merged: u64,
    /// The conditional formats' rules, which no row says.
    pub conditional: u64,
    /// The rows formatted whole, a row's own format, which no row says.
    pub row_formats: u64,
}

impl LooksStats {
    /// The sum of two counts.
    pub fn add(&mut self, other: &LooksStats) {
        self.formats += other.formats;
        self.styled += other.styled;
        self.formatted_empty += other.formatted_empty;
        self.unheld += other.unheld;
        self.undrawn += other.undrawn;
        self.merged += other.merged;
        self.conditional += other.conditional;
        self.row_formats += other.row_formats;
    }

    /// The counts as `--looks --counts` prints them beside a sheet's.
    pub fn line(&self) -> String {
        format!(
            "looks formats {} styled {} formatted-empty {} unheld {} undrawn {} merged {} conditional {} row-formats {}",
            self.formats,
            self.styled,
            self.formatted_empty,
            self.unheld,
            self.undrawn,
            self.merged,
            self.conditional,
            self.row_formats
        )
    }

    /// What a reading of these looks left out, as a door says it beside the
    /// rows (`SD-30`); `None` where it left nothing out and no code is one a
    /// load refuses.
    pub fn note(&self) -> Option<String> {
        let count =
            |n: u64, one: &str, many: &str| format!("{n} {}", if n == 1 { one } else { many });
        let mut left: Vec<String> = Vec::new();
        if self.formatted_empty > 0 {
            left.push(count(
                self.formatted_empty,
                "cell that keeps a format and no value",
                "cells that keep a format and no value",
            ));
        }
        if self.unheld > 0 {
            left.push(count(
                self.unheld,
                "look the record has no word for",
                "looks the record has no word for",
            ));
        }
        if self.merged > 0 {
            left.push(count(self.merged, "merged range", "merged ranges"));
        }
        if self.conditional > 0 {
            left.push(count(
                self.conditional,
                "conditional format",
                "conditional formats",
            ));
        }
        if self.row_formats > 0 {
            left.push(count(
                self.row_formats,
                "row formatted whole",
                "rows formatted whole",
            ));
        }
        let mut note = String::new();
        if !left.is_empty() {
            note.push_str(&format!(
                "the looks leave out what the record cannot hold: {}",
                left.join(", ")
            ));
        }
        if self.undrawn > 0 {
            if !note.is_empty() {
                note.push_str("; and ");
            }
            note.push_str(&format!(
                "{} printed as the file holds {}, which a load refuses by name",
                count(
                    self.undrawn,
                    "cell's number format this version does not draw is",
                    "cells' number formats this version does not draw are"
                ),
                if self.undrawn == 1 { "it" } else { "them" }
            ));
        }
        if note.is_empty() {
            return None;
        }
        note.push_str("; --looks --counts counts each sheet's.");
        Some(note)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// A draft's own `value` rows, as a grid's values for the writer and
    /// the view.
    #[derive(Debug)]
    struct Drafted(std::collections::HashMap<vla_lang::calc::CellId, vla_lang::calc::Computed>);

    impl vla_lang::calc::Values for Drafted {
        fn computed(&self, id: vla_lang::calc::CellId) -> Option<vla_lang::calc::Computed> {
            self.0.get(&id).cloned()
        }
    }

    /// Each sheet's record, viewed whole, of a text of rows loaded.
    fn records(rows: &str) -> Vec<String> {
        let forms = vla_lang::reader::read_forms(rows).unwrap();
        let draft = vla_lang::machine::read_rows(&forms)
            .unwrap_or_else(|r| panic!("line {}: {}", r.line, r.refusal));
        let held = Drafted(draft.values.clone());
        let book = &draft.workbook;
        book.sheets
            .iter()
            .map(|sheet| {
                crate::view::view_text_valued(
                    book,
                    &crate::kernel::Window {
                        sheet: sheet.name.clone(),
                        range: crate::view::whole(sheet),
                    },
                    &held,
                )
            })
            .collect()
    }

    /// The round trip (decision 12): a record holding every look the writer
    /// writes, built into a workbook by the writer with its values, read
    /// under the looks and loaded, is each sheet's record again, byte for
    /// byte: gridlines, columns with widths on whole pixels, hidden and with
    /// formats, rows with heights, hidden, and of settings alone, formats of
    /// fills, codes and every group, styled cells, formulas with their
    /// values, a number under a date's code, a hidden sheet, a sheet with
    /// no cell.
    #[test]
    fn a_record_written_and_read_with_its_looks_is_itself() {
        let rows = r##"
(sheet "Board" visible)
(sheet "Notes" hidden)
(sheet "Empty" visible)
(gridlines "Board" off)
(column "Board" "A" 20 shown 1)
(column "Board" "B" none hidden none)
(column "Board" "C" 9.14 shown none)
(column "Board" "D" none shown 2)
(row "Board" 2 30 shown)
(row "Board" 3 none hidden)
(row "Board" 5 22.5 shown)
(format 0 none general nowrap)
(format 1 "F7F4FC" text nowrap (font bold "6B4FC8") (align center middle))
(format 2 none "0.00" wrap (font italic 22) (border (left thin) (top medium "6B4FC8")))
(format 3 none "m/d/yyyy" nowrap (align right top) (border (right double "800020") (bottom hair)))
(format 4 "DDEBDD" general nowrap (font 9.3 "000000") (border (left dash-dot) (right medium-dash-dot-dot) (top slant-dash-dot) (bottom thick)))
(cell "Board" "A1" "Title")
(style "Board" "A1" 1)
(cell "Board" "B1" 1.5)
(style "Board" "B1" 2)
(cell "Board" "C2" 45930)
(style "Board" "C2" 3)
(formula "Board" "D2" "=B1*2")
(value "Board" "D2" 3)
(style "Board" "D2" 2)
(cell "Board" "A3" true)
(cell "Board" "B3" (error "#N/A"))
(style "Board" "B3" 4)
(formula "Board" "C3" "=A1&\"!\"")
(value "Board" "C3" "Title!")
(cell "Board" "D6" "end")
(cell "Notes" "A1" "a note")
(style "Notes" "A1" 4)
(row "Empty" 1 40 shown)
"##;
        let before = records(rows);
        let forms = vla_lang::reader::read_forms(rows).unwrap();
        let draft = vla_lang::machine::read_rows(&forms).unwrap();
        let held = Drafted(draft.values.clone());
        let bytes = crate::sheet::ooxml::workbook_bytes_valued(&draft.workbook, &held).unwrap();
        let (read, left) = crate::reflect::reflect_looks(&bytes, "law.xlsx").unwrap();
        assert_eq!(left.note(), None, "{left:?}");
        let after = records(&read);
        assert_eq!(before.len(), after.len());
        for (b, a) in before.iter().zip(&after) {
            assert_eq!(b, a, "read back:\n{read}");
        }
        // The rows the law reaches, so that a reader printing too little
        // cannot pass.
        for row in [
            "(gridlines \"Board\" off)",
            "(column \"Board\" \"A\" 20 shown 1)",
            "(column \"Board\" \"B\" none hidden none)",
            "(column \"Board\" \"C\" 9.14 shown none)",
            "(column \"Board\" \"D\" none shown 2)",
            "(row \"Board\" 5 22.5 shown)",
            "(row \"Empty\" 1 40 shown)",
            "(value \"Board\" \"D2\" 3)",
            "(sheet \"Notes\" hidden)",
        ] {
            assert!(read.lines().any(|l| l == row), "{row} is not in:\n{read}");
        }
    }

    #[test]
    fn a_color_resolves_through_the_theme_and_the_palette() {
        let mut slots = [None; 12];
        slots[0] = Some(Rgb(0xFF, 0xFF, 0xFF));
        slots[1] = Some(Rgb(0, 0, 0));
        slots[4] = Some(Rgb(0x15, 0x60, 0x82));
        let colors = Colors::new(Some(slots), None);
        let of = |color: Color, tint: f64| colors.resolve(&ColorRef { color, tint });
        assert_eq!(
            of(Color::Rgb(Rgb(0x6B, 0x4F, 0xC8)), 0.0),
            Resolved::Rgb(Rgb(0x6B, 0x4F, 0xC8))
        );
        assert_eq!(of(Color::Theme(1), 0.0), Resolved::Rgb(Rgb(0, 0, 0)));
        assert_eq!(of(Color::Theme(0), 0.0), Resolved::Rgb(Rgb(255, 255, 255)));
        assert_eq!(
            of(Color::Theme(4), 0.0),
            Resolved::Rgb(Rgb(0x15, 0x60, 0x82))
        );
        assert_eq!(of(Color::Theme(5), 0.0), Resolved::Unresolved);
        assert_eq!(of(Color::Theme(12), 0.0), Resolved::Unresolved);
        assert_eq!(of(Color::Indexed(2), 0.0), Resolved::Rgb(Rgb(255, 0, 0)));
        assert_eq!(of(Color::Indexed(10), 0.0), Resolved::Rgb(Rgb(255, 0, 0)));
        assert_eq!(
            of(Color::Indexed(22), 0.0),
            Resolved::Rgb(Rgb(0xC0, 0xC0, 0xC0))
        );
        assert_eq!(of(Color::Indexed(64), 0.0), Resolved::Automatic);
        assert_eq!(
            of(Color::Indexed(65), 0.0),
            Resolved::Rgb(Rgb(255, 255, 255)),
            "the system's paper, which Excel draws white"
        );
        assert_eq!(of(Color::Indexed(66), 0.0), Resolved::Unresolved);
        assert_eq!(of(Color::Auto, 0.0), Resolved::Automatic);
        let none = Colors::new(None, Some(vec![Rgb(1, 2, 3)]));
        assert_eq!(
            none.resolve(&ColorRef {
                color: Color::Theme(1),
                tint: 0.0
            }),
            Resolved::Unresolved
        );
        assert_eq!(
            none.resolve(&ColorRef {
                color: Color::Indexed(0),
                tint: 0.0
            }),
            Resolved::Rgb(Rgb(1, 2, 3))
        );
        assert_eq!(
            none.resolve(&ColorRef {
                color: Color::Indexed(1),
                tint: 0.0
            }),
            Resolved::Unresolved
        );
    }

    #[test]
    fn a_tint_moves_a_color_s_luminance_as_excel_draws_it() {
        // Excel's own colors (the item's golden): the grays a theme's white
        // and black make at Excel's palette tints, as Excel's saves write
        // the tints, each a gray to its nearest channel.
        let tint = |text: &str| text.parse::<f64>().unwrap();
        let white = Rgb(255, 255, 255);
        let black = Rgb(0, 0, 0);
        for (base, t, want) in [
            (white, "-4.9989318521683403E-2", "F2F2F2"),
            (white, "-0.14999847407452621", "D9D9D9"),
            (white, "-0.249977111117893", "BFBFBF"),
            (white, "-0.34998626667073579", "A6A6A6"),
            (white, "-0.499984740745262", "808080"),
            (black, "0.499984740745262", "808080"),
            (black, "0.34998626667073579", "595959"),
            (black, "0.249977111117893", "404040"),
            (black, "0.14999847407452621", "262626"),
            (black, "4.9989318521683403E-2", "0D0D0D"),
            // The accents, lightened and darkened, where exact HLS is a
            // unit or two off a channel (DAE3F3 for D9E1F2) and a tint of
            // each channel further (773F19 for 833C0C).
            (Rgb(0x44, 0x72, 0xC4), "0.79998168889431442", "D9E1F2"),
            (Rgb(0x44, 0x72, 0xC4), "-0.499984740745262", "203764"),
            (Rgb(0xED, 0x7D, 0x31), "0.39997558519241921", "F4B084"),
            (Rgb(0xED, 0x7D, 0x31), "-0.499984740745262", "833C0C"),
            (Rgb(0xE7, 0xE6, 0xE6), "-0.89999084444715716", "161616"),
            (Rgb(0x44, 0x54, 0x6A), "-0.249977111117893", "333F4F"),
            // A tint no palette makes, kept in 32767ths as Excel keeps it.
            (Rgb(0x44, 0x72, 0xC4), "0.3", "7A9AD6"),
            (Rgb(0xED, 0x7D, 0x31), "-0.6", "69300A"),
            // Tints between two thousandths, each product truncated: a tint
            // in thousandths would draw the white DFDFDF, in hundredths the
            // black 1E1E1E.
            (black, "0.1249", "202020"),
            (white, "-0.1251", "DEDEDE"),
            (Rgb(0x44, 0x72, 0xC4), "0.1249", "5A82CB"),
            (Rgb(0xED, 0x7D, 0x31), "-0.1251", "E76814"),
        ] {
            assert_eq!(tinted(base, tint(t)).hex(), want, "{} at {t}", base.hex());
        }
        assert_eq!(tinted(Rgb(0x44, 0x72, 0xC4), 0.0), Rgb(0x44, 0x72, 0xC4));
        assert_eq!(tinted(Rgb(0x44, 0x72, 0xC4), 1.0), white);
        assert_eq!(tinted(Rgb(0x44, 0x72, 0xC4), -1.0), black);
    }

    #[test]
    fn a_built_in_number_format_reads_as_excel_reads_it() {
        assert_eq!(builtin_code(2), Some("0.00"));
        assert_eq!(builtin_code(14), Some("m/d/yyyy"));
        assert_eq!(builtin_code(37), Some("#,##0_);(#,##0)"));
        assert_eq!(builtin_code(49), Some("@"));
        assert_eq!(builtin_code(23), None);
        // A locale's id is its twin, as Excel's save writes it.
        for (id, twin) in [
            (27, 14),
            (32, 21),
            (36, 14),
            (50, 14),
            (59, 1),
            (63, 5),
            (69, 12),
            (73, 15),
            (76, 20),
            (77, 21),
            (78, 22),
            (81, 47),
            (23, 0),
            (163, 0),
            (5, 5),
            (44, 44),
        ] {
            assert_eq!(builtin_twin(id), twin, "id {id}");
        }
        let mut look = Look::default();
        look.set_code("0.00");
        assert!(matches!(look.style.num_fmt, NumFmt::Code(_)));
        let mut look = Look::default();
        look.set_code("General");
        assert_eq!(look.style.num_fmt, NumFmt::General);
        assert_eq!(look.undrawn, None);
        let mut look = Look::default();
        look.set_code("@");
        assert_eq!(look.style.num_fmt, NumFmt::Text);
        // A code the language does not draw is kept as written.
        let mut look = Look::default();
        look.set_code(builtin_code(44).unwrap());
        assert_eq!(look.style.num_fmt, NumFmt::General);
        assert_eq!(look.undrawn.as_deref(), builtin_code(44));
        let mut look = Look::default();
        look.set_code(builtin_code(14).unwrap());
        assert!(
            matches!(look.style.num_fmt, NumFmt::Code(_)),
            "a date draws"
        );
    }

    #[test]
    fn a_width_reads_as_excel_reads_it() {
        assert_eq!(width_chars(72.7109375), Some(72.0));
        assert_eq!(width_chars(72.7265625), Some(72.0));
        assert_eq!(width_chars(60.7265625), Some(60.0));
        assert_eq!(width_chars(60.7109375), Some(60.0));
        assert_eq!(width_chars(11.2109375), Some(10.43));
        assert_eq!(width_chars(9.140625), None, "Excel's default, 64 pixels");
        // A width a writer wrote without Excel's padding draws narrower: 59
        // pixels, 7.71 characters.
        assert_eq!(width_chars(8.43), Some(7.71));
        assert_eq!(width_pixels(0.0), 0.0);
    }

    #[test]
    fn a_font_s_size_is_its_whole_twentieths() {
        assert_eq!(font_size(22.0).map(|s| s.spell()), Some("22".to_string()));
        let saved: f64 = "9.3000000000000007".parse().unwrap();
        assert_eq!(font_size(saved).map(|s| s.spell()), Some("9.3".to_string()));
        assert_eq!(
            font_size(409.55).map(|s| s.spell()),
            Some("409.55".to_string())
        );
        assert_eq!(font_size(10.33), None);
        assert_eq!(font_size(0.5), None);
        assert_eq!(font_size(409.6), None);
    }

    #[test]
    fn an_iso_date_is_excel_s_serial() {
        assert_eq!(iso_serial("2025-09-30", false), Some(45930.0));
        assert_eq!(iso_serial("1900-01-01", false), Some(1.0));
        assert_eq!(iso_serial("1900-03-01", false), Some(61.0));
        assert_eq!(iso_serial("2025-09-30T12:00:00", false), Some(45930.5));
        assert_eq!(iso_serial("2025-09-30T00:00:00Z", false), Some(45930.0));
        assert_eq!(iso_serial("2025-09-30", true), Some(44468.0));
        assert_eq!(iso_serial("1899-12-31", false), None);
        assert_eq!(iso_serial("2025-13-01", false), None);
        assert_eq!(iso_serial("30/09/2025", false), None);
    }

    #[test]
    fn the_note_says_what_was_left_out() {
        assert_eq!(LooksStats::default().note(), None);
        let stats = LooksStats {
            formats: 3,
            styled: 9,
            ..LooksStats::default()
        };
        assert_eq!(stats.note(), None, "formats and styled cells are held");
        let stats = LooksStats {
            formatted_empty: 2,
            merged: 1,
            undrawn: 1,
            ..LooksStats::default()
        };
        assert_eq!(
            stats.note().unwrap(),
            "the looks leave out what the record cannot hold: 2 cells that keep a format and no value, 1 merged range; and 1 cell's number format this version does not draw is printed as the file holds it, which a load refuses by name; --looks --counts counts each sheet's."
        );
        assert_eq!(
            stats.line(),
            "looks formats 0 styled 0 formatted-empty 2 unheld 0 undrawn 1 merged 1 conditional 0 row-formats 0"
        );
    }
}
