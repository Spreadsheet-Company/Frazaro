//! The build (PORT.7): a program's sentences to a workbook's bytes, through
//! the sheet model and the OOXML writer. Slice 7a writes the room's first
//! sheet, `Frazaro`: the program's lines in column B from row 1 exactly as
//! written, blank lines included, and OK in green in column C beside every
//! line that holds text, which is what the add-in's Check marks (`DoCheck`,
//! VLA_IDE.bas: once the program translates, a row whose column B is not
//! blank is marked OK). The widths, the fills, the hidden column A and the
//! gridlines are `BuildWorkspace`'s (VLA_IDE.bas), so the room looks the same
//! through this door as through the add-in's. The program's VLA is read here
//! and, in this slice, nothing beyond that sheet is rendered from it: slice
//! 7b walks it into cells and formulas.
//!
//! The build is a pure function of its texts: the same sentences and the
//! same core give the same bytes on every machine, which the build golden
//! (`scripts/build/fixture_golden.xlsx`, the treaty's amendment of
//! 2026-10-03) holds it to.

use crate::messages::{raise, Refusal};
use crate::reader::read_forms;
use crate::sheet::{ooxml, Cell, Column, Content, NumFmt, Rgb, Sheet, Style, Styles, Workbook};

/// The sheet that holds the program, named as the add-in names its
/// workspace (`IDE_SHEET`).
pub const FRAZARO_SHEET: &str = "Frazaro";
/// `BuildWorkspace`'s whisper of the logo's lavender on the sentence column.
pub const LAVENDER: Rgb = Rgb(247, 244, 252);
/// The result column's gray, under the marks.
pub const RESULT_GRAY: Rgb = Rgb(242, 242, 242);
/// `MarkOK`'s green.
pub const OK_GREEN: Rgb = Rgb(221, 235, 221);
/// `BuildWorkspace`'s column widths, in characters.
pub const SENTENCE_COLUMN_WIDTH: f64 = 72.0;
pub const RESULT_COLUMN_WIDTH: f64 = 60.0;

/// The `Frazaro` sheet of a program whose text translated: each line of the
/// text in column B of its own row, and OK in column C beside a line that
/// is not blank. Blank is `Len(Trim$(...)) = 0`, and VBA's `Trim$` trims
/// spaces only, so a line of tabs is marked and a line of spaces is not.
pub fn frazaro_sheet(program_text: &str, styles: &mut Styles) -> Sheet {
    let sentence = styles.id(Style {
        fill: Some(LAVENDER),
        num_fmt: NumFmt::Text,
        wrap: false,
    });
    let result = styles.id(Style {
        fill: Some(RESULT_GRAY),
        num_fmt: NumFmt::General,
        wrap: true,
    });
    let ok = styles.id(Style {
        fill: Some(OK_GREEN),
        num_fmt: NumFmt::General,
        wrap: true,
    });
    let mut sheet = Sheet::new(FRAZARO_SHEET);
    sheet.gridlines = false;
    sheet.active_cell = (1, 2);
    sheet.columns = vec![
        Column {
            index: 1,
            width: None,
            hidden: true,
            style: None,
        },
        Column {
            index: 2,
            width: Some(SENTENCE_COLUMN_WIDTH),
            hidden: false,
            style: Some(sentence),
        },
        Column {
            index: 3,
            width: Some(RESULT_COLUMN_WIDTH),
            hidden: false,
            style: Some(result),
        },
    ];
    for (i, line) in program_text.lines().enumerate() {
        let row = i as u32 + 1;
        if line.is_empty() {
            continue;
        }
        sheet.set(
            row,
            2,
            Cell {
                content: Content::Text(line.to_string()),
                style: sentence,
            },
        );
        if !line.trim_matches(' ').is_empty() {
            sheet.set(
                row,
                3,
                Cell {
                    content: Content::Text("OK".to_string()),
                    style: ok,
                },
            );
        }
    }
    sheet
}

/// The workbook a translated program builds: in this slice, the `Frazaro`
/// sheet alone. `vla` is the program's VLA, read so that what slice 7b walks
/// is in hand; a VLA that does not read is the core's own fault and comes
/// back as the reader's refusal.
pub fn build_workbook(program_text: &str, vla: &str) -> Result<Workbook, Refusal> {
    let _forms = read_forms(vla)?;
    let mut wb = Workbook::new();
    let sheet = frazaro_sheet(program_text, &mut wb.styles);
    wb.sheets.push(sheet);
    wb.active_sheet = 0;
    Ok(wb)
}

/// The workbook's bytes, or the catalogue's refusal when they would not fit
/// the file format.
pub fn build_xlsx(program_text: &str, vla: &str) -> Result<Vec<u8>, Refusal> {
    let wb = build_workbook(program_text, vla)?;
    ooxml::workbook_bytes(&wb).ok_or_else(|| raise("build-workbook-too-large", &[]))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::sheet::zip;

    const PRELUDE: &str = include_str!("../../scripts/prelude.vla");
    const ENGLISH: &str = include_str!("../../scripts/polyglotta/english.vla");
    const FIXTURE: &str = include_str!("../../scripts/build/fixture.txt");
    const GOLDEN: &[u8] = include_bytes!("../../scripts/build/fixture_golden.xlsx");

    fn text_of(cell: &Cell) -> &str {
        match &cell.content {
            Content::Text(t) => t,
            other => panic!("not a text: {other:?}"),
        }
    }

    #[test]
    fn the_frazaro_sheet_is_the_room() {
        let mut styles = Styles::new();
        let sh = frazaro_sheet(
            "Put 5 into cell B2.\n\n# a comment\n   \n\tindented\nlast",
            &mut styles,
        );
        assert_eq!(sh.name, "Frazaro");
        assert!(!sh.gridlines);
        assert_eq!(sh.active_cell, (1, 2));
        assert!(sh.columns[0].hidden);
        assert_eq!(sh.columns[1].width, Some(72.0));
        assert_eq!(sh.columns[2].width, Some(60.0));
        // Row 1: the sentence and its OK. Row 2, blank: nothing. Row 3, a
        // comment: marked, as Check marks any row with text. Row 4, spaces
        // only: the text kept, no mark. Row 5, a tab: marked, as Trim$ would
        // leave the tab. Row 6: the last line, without a line break after it.
        assert_eq!(text_of(&sh.cells[&(1, 2)]), "Put 5 into cell B2.");
        assert_eq!(text_of(&sh.cells[&(1, 3)]), "OK");
        assert!(!sh.cells.contains_key(&(2, 2)));
        assert_eq!(text_of(&sh.cells[&(3, 2)]), "# a comment");
        assert_eq!(text_of(&sh.cells[&(3, 3)]), "OK");
        assert_eq!(text_of(&sh.cells[&(4, 2)]), "   ");
        assert!(!sh.cells.contains_key(&(4, 3)));
        assert_eq!(text_of(&sh.cells[&(5, 3)]), "OK");
        assert_eq!(text_of(&sh.cells[&(6, 2)]), "last");
        assert_eq!(sh.cells.len(), 9);
        // The three formats, in order of first use, after the default.
        let xfs = styles.xfs();
        assert_eq!(xfs.len(), 4);
        assert_eq!(xfs[1].fill, Some(LAVENDER));
        assert_eq!(xfs[1].num_fmt, NumFmt::Text);
        assert_eq!(xfs[2].fill, Some(RESULT_GRAY));
        assert_eq!(xfs[3].fill, Some(OK_GREEN));
        assert!(xfs[3].wrap);
    }

    #[test]
    fn the_same_text_gives_the_same_bytes() {
        let a = build_xlsx("Log 1.\n", "(sub main (log 1))").unwrap();
        let b = build_xlsx("Log 1.\n", "(sub main (log 1))").unwrap();
        assert_eq!(a, b);
        assert!(build_xlsx("Log 1.\n", "(sub main").is_err());
    }

    #[test]
    fn the_refusals_come_from_the_catalogue() {
        let r = raise("build-output-exists", &[("path", "out.xlsx")]);
        assert_eq!(r.source, "VLA-Build");
        assert!(r.text.contains("out.xlsx"));
        assert!(r.text.contains("--replace"));
        let r = raise("build-workbook-too-large", &[]);
        assert_eq!(r.source, "VLA-Build");
    }

    /// The build golden: the fixture, built with the prelude and english.vla,
    /// is `scripts/build/fixture_golden.xlsx` byte for byte. When this fails
    /// the writer changed: regenerate with `frazaro build
    /// scripts/build/fixture.txt --prelude scripts/prelude.vla --phrasebook
    /// scripts/polyglotta/english.vla --out scripts/build/fixture_golden.xlsx
    /// --replace`, open it in Excel (the owner's live pass), and raise
    /// tools/check_build_golden.ps1's floor if it grew.
    #[test]
    fn the_build_golden_is_reproduced() {
        let got = crate::api::english_build_xlsx(FIXTURE, PRELUDE, &[ENGLISH])
            .unwrap_or_else(|r| panic!("the fixture did not build: {}", r.refusal));
        if got != GOLDEN {
            let at = got
                .iter()
                .zip(GOLDEN.iter())
                .position(|(a, b)| a != b)
                .unwrap_or(got.len().min(GOLDEN.len()));
            let part = zip::entries(GOLDEN)
                .unwrap_or_default()
                .into_iter()
                .find(|e| at >= e.data_offset && at < e.data_offset + e.size as usize)
                .map(|e| format!("{} at offset {}", e.name, at - e.data_offset))
                .unwrap_or_else(|| "a header".to_string());
            panic!(
                "the build differs from scripts/build/fixture_golden.xlsx at byte {at} of {} ({part}); got {} bytes",
                GOLDEN.len(),
                got.len()
            );
        }
    }
}
