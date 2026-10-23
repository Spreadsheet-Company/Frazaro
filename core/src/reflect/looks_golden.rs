//! Excel's golden of the looks (KERNEL.57, the second slice), test-only.
//!
//! Two workbooks, a case a row, written here and held byte for byte
//! (`FRAZARO_BLESS=1` writes them again), hold what neither Excel's writer
//! nor the core's writes, so that Excel's reading of each is the oracle of
//! the package reader's rules (decision 12):
//!
//! - `scripts/reflect/looks.xlsx`: every kind of color a package names,
//!   each of the theme's twelve slots untinted, at each tint of Excel's
//!   palette and at tints no palette makes, each index of the default
//!   palette, an alpha and Excel's automatic, as a fill, a font's color and
//!   an edge's; every built-in number format and codes the language does
//!   not draw; the `apply` flags off and a cell style that says otherwise;
//!   what the record has no word for; widths and heights; a row's format, a
//!   formatted empty cell, a merge and a conditional format; a sheet's own
//!   defaults.
//! - `scripts/reflect/looks_book.xlsx`: what a workbook says once, a palette
//!   of its own, no theme part and a default font of Arial 10.
//!
//! A case's name is in column A and its reads in column B: the targets
//! `VlaWriteLooksGolden` (VLA_Tests.bas) reads through Excel's object model
//! into `scripts/reflect/looks_golden.txt`, case by case, before it saves
//! each workbook as Excel's own beside it, `looks_saved.xlsx` and
//! `looks_book_saved.xlsx`. A target is `cell:<A1>`, a cell and its value;
//! `empty:<A1>`, a cell that keeps a format and no value; `blank:<A1>`, no
//! cell, drawn with its column's or its row's format; `theme:<A1>`, a
//! fill's theme color and tint as Excel names them; `col:<letters>`,
//! `row:<n>`, `merge:<A1>`, `cf:<A1>`, `grid` or `std`. The cases sit two
//! rows apart and a case's cells two columns apart, so that no two cells
//! share an edge, which the object model reads as either cell's.

use std::collections::{BTreeMap, HashMap, HashSet};

use vla_lang::form::Form;

use super::looks::{LooksStats, THEME_SLOTS};
use crate::sheet::{cell_ref, column_letters, number_text, xml, zip};

const LOOKS: &str = concat!(env!("CARGO_MANIFEST_DIR"), "/../scripts/reflect/looks.xlsx");
const BOOK: &str = concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../scripts/reflect/looks_book.xlsx"
);

const HEAD: &str = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
const NS_MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const NS_REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_PKG: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const CT: &str = "application/vnd.openxmlformats-officedocument";

/// The workbook's font, Calibri 11 in the theme's ink, as Excel 2016 to
/// 2022 wrote a new workbook's.
const CALIBRI: &str = "<name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/>";
/// The second workbook's font, Arial 10 with no color of its own.
const ARIAL: &str = "<name val=\"Arial\"/><family val=\"2\"/>";

/// A number every number format draws: 2025-01-21 at 9:22:30.
const VALUE: &str = "45678.390625";

/// A theme part with the twelve colors of Office's theme of 2016 to 2022,
/// each slot its own color, and the least of the rest a theme part holds.
const THEME: &str = concat!(
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
    "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Office Theme\">",
    "<a:themeElements><a:clrScheme name=\"Office\">",
    "<a:dk1><a:sysClr val=\"windowText\" lastClr=\"000000\"/></a:dk1>",
    "<a:lt1><a:sysClr val=\"window\" lastClr=\"FFFFFF\"/></a:lt1>",
    "<a:dk2><a:srgbClr val=\"44546A\"/></a:dk2><a:lt2><a:srgbClr val=\"E7E6E6\"/></a:lt2>",
    "<a:accent1><a:srgbClr val=\"4472C4\"/></a:accent1><a:accent2><a:srgbClr val=\"ED7D31\"/></a:accent2>",
    "<a:accent3><a:srgbClr val=\"A5A5A5\"/></a:accent3><a:accent4><a:srgbClr val=\"FFC000\"/></a:accent4>",
    "<a:accent5><a:srgbClr val=\"5B9BD5\"/></a:accent5><a:accent6><a:srgbClr val=\"70AD47\"/></a:accent6>",
    "<a:hlink><a:srgbClr val=\"0563C1\"/></a:hlink><a:folHlink><a:srgbClr val=\"954F72\"/></a:folHlink>",
    "</a:clrScheme>",
    "<a:fontScheme name=\"Office\">",
    "<a:majorFont><a:latin typeface=\"Calibri Light\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:majorFont>",
    "<a:minorFont><a:latin typeface=\"Calibri\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:minorFont>",
    "</a:fontScheme>",
    "<a:fmtScheme name=\"Office\">",
    "<a:fillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></a:fillStyleLst>",
    "<a:lnStyleLst><a:ln w=\"6350\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></a:ln><a:ln w=\"12700\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></a:ln><a:ln w=\"19050\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></a:ln></a:lnStyleLst>",
    "<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>",
    "<a:bgFillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></a:bgFillStyleLst>",
    "</a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>\n",
);

/// Excel's palette tints for a theme slot by its place in Excel's count,
/// as its saves write them: the light slots darkened, the dark lightened,
/// the accents and the links both.
fn palette_tints(slot: usize) -> [&'static str; 5] {
    match slot {
        0 => [
            "-4.9989318521683403E-2",
            "-0.14999847407452621",
            "-0.249977111117893",
            "-0.34998626667073579",
            "-0.499984740745262",
        ],
        1 => [
            "0.499984740745262",
            "0.34998626667073579",
            "0.249977111117893",
            "0.14999847407452621",
            "4.9989318521683403E-2",
        ],
        2 => [
            "-9.9978637043366805E-2",
            "-0.249977111117893",
            "-0.499984740745262",
            "-0.74999237037263116",
            "-0.89999084444715716",
        ],
        _ => [
            "0.79998168889431442",
            "0.59999389629810485",
            "0.39997558519241921",
            "-0.249977111117893",
            "-0.499984740745262",
        ],
    }
}

/// An element's place in its list, appended where the list has it not.
fn place(list: &mut Vec<String>, element: String) -> usize {
    match list.iter().position(|e| *e == element) {
        Some(i) => i,
        None => {
            list.push(element);
            list.len() - 1
        }
    }
}

/// A solid fill of a color, its attributes as a file writes them.
fn solid(color: &str) -> String {
    format!("<fill><patternFill patternType=\"solid\"><fgColor {color}/><bgColor indexed=\"64\"/></patternFill></fill>")
}

/// A border of one edge, the cell's left.
fn left_edge(line: &str, color: &str) -> String {
    format!("<border><left style=\"{line}\"><color {color}/></left><right/><top/><bottom/><diagonal/></border>")
}

/// A cell format as the fixture writes it.
#[derive(Clone, Default)]
struct Xf {
    num_fmt: u32,
    font: Option<String>,
    fill: Option<String>,
    border: Option<String>,
    /// The alignment element's attributes.
    alignment: Option<String>,
    /// The protection element's attributes.
    protection: Option<String>,
    /// The cell style's format, its `xfId`.
    style: u32,
    /// The `apply` flags as written; `None` writes Excel's, a flag for each
    /// property the format sets.
    apply: Option<String>,
}

/// The `apply` flags of a cell format's five properties.
const FLAGS: [&str; 5] = [
    "applyNumberFormat",
    "applyFont",
    "applyFill",
    "applyBorder",
    "applyAlignment",
];

/// The five `apply` flags written, each on but those named.
fn flags(off: &[&str]) -> String {
    FLAGS
        .iter()
        .map(|f| format!(" {f}=\"{}\"", if off.contains(f) { 0 } else { 1 }))
        .collect()
}

/// A format of one look: a font, a fill, an alignment or a border.
fn with_font(font: String) -> Xf {
    Xf {
        font: Some(font),
        ..Xf::default()
    }
}

fn with_fill(fill: String) -> Xf {
    Xf {
        fill: Some(fill),
        ..Xf::default()
    }
}

fn aligned(alignment: &str) -> Xf {
    Xf {
        alignment: Some(alignment.to_string()),
        ..Xf::default()
    }
}

fn bordered(border: String) -> Xf {
    Xf {
        border: Some(border),
        ..Xf::default()
    }
}

/// A styles part being written, each list's elements once, in order.
struct Styles {
    /// The default font's size and face, and its color element.
    size: &'static str,
    face: &'static str,
    ink: &'static str,
    codes: Vec<(u32, String)>,
    fonts: Vec<String>,
    fills: Vec<String>,
    borders: Vec<String>,
    style_xfs: Vec<String>,
    xfs: Vec<String>,
    cell_styles: Vec<String>,
    dxfs: Vec<String>,
    palette: Vec<String>,
}

impl Styles {
    fn new(size: &'static str, face: &'static str, ink: &'static str) -> Styles {
        let mut s = Styles {
            size,
            face,
            ink,
            codes: Vec::new(),
            fonts: Vec::new(),
            fills: vec![
                "<fill><patternFill patternType=\"none\"/></fill>".to_string(),
                "<fill><patternFill patternType=\"gray125\"/></fill>".to_string(),
            ],
            borders: vec!["<border><left/><right/><top/><bottom/><diagonal/></border>".to_string()],
            style_xfs: vec![
                "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>".to_string(),
            ],
            xfs: Vec::new(),
            cell_styles: vec!["<cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/>".to_string()],
            dxfs: Vec::new(),
            palette: Vec::new(),
        };
        let default = s.font("", ink);
        s.fonts.push(default);
        s.xfs.push(
            "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>".to_string(),
        );
        s
    }

    /// A font of the workbook's default face and size, `before` its flags
    /// and `color` its color element.
    fn font(&self, before: &str, color: &str) -> String {
        format!(
            "<font>{before}<sz val=\"{}\"/>{color}{}</font>",
            self.size, self.face
        )
    }

    /// The id of a code the part defines, from 164.
    fn code(&mut self, code: &str) -> u32 {
        if let Some((id, _)) = self.codes.iter().find(|(_, c)| c == code) {
            return *id;
        }
        let id = 164 + self.codes.len() as u32;
        self.codes.push((id, code.to_string()));
        id
    }

    /// A cell format's index, the format written once.
    fn xf(&mut self, x: Xf) -> u32 {
        let font = x.font.map_or(0, |f| place(&mut self.fonts, f));
        let fill = x.fill.map_or(0, |f| place(&mut self.fills, f));
        let border = x.border.map_or(0, |b| place(&mut self.borders, b));
        let apply = x.apply.unwrap_or_else(|| {
            let mut a = String::new();
            for (on, name) in [
                (x.num_fmt != 0, "applyNumberFormat"),
                (font != 0, "applyFont"),
                (fill != 0, "applyFill"),
                (border != 0, "applyBorder"),
                (x.alignment.is_some(), "applyAlignment"),
                (x.protection.is_some(), "applyProtection"),
            ] {
                if on {
                    a.push_str(&format!(" {name}=\"1\""));
                }
            }
            a
        });
        let mut children = String::new();
        if let Some(a) = &x.alignment {
            children.push_str(&format!("<alignment {a}/>"));
        }
        if let Some(p) = &x.protection {
            children.push_str(&format!("<protection {p}/>"));
        }
        let open = format!(
            "<xf numFmtId=\"{}\" fontId=\"{font}\" fillId=\"{fill}\" borderId=\"{border}\" xfId=\"{}\"{apply}",
            x.num_fmt, x.style
        );
        let element = if children.is_empty() {
            format!("{open}/>")
        } else {
            format!("{open}>{children}</xf>")
        };
        place(&mut self.xfs, element) as u32
    }

    fn part(&self) -> String {
        let list = |name: &str, items: &[String]| {
            format!(
                "<{name} count=\"{}\">{}</{name}>\n",
                items.len(),
                items.concat()
            )
        };
        let mut s = format!("{HEAD}<styleSheet xmlns=\"{NS_MAIN}\">\n");
        if !self.codes.is_empty() {
            let codes: Vec<String> = self
                .codes
                .iter()
                .map(|(id, code)| {
                    format!(
                        "<numFmt numFmtId=\"{id}\" formatCode=\"{}\"/>",
                        xml::attr(code)
                    )
                })
                .collect();
            s.push_str(&list("numFmts", &codes));
        }
        s.push_str(&list("fonts", &self.fonts));
        s.push_str(&list("fills", &self.fills));
        s.push_str(&list("borders", &self.borders));
        s.push_str(&list("cellStyleXfs", &self.style_xfs));
        s.push_str(&list("cellXfs", &self.xfs));
        s.push_str(&list("cellStyles", &self.cell_styles));
        s.push_str(&list("dxfs", &self.dxfs));
        if !self.palette.is_empty() {
            let colors: Vec<String> = self
                .palette
                .iter()
                .map(|c| format!("<rgbColor rgb=\"{c}\"/>"))
                .collect();
            s.push_str(&format!(
                "<colors><indexedColors>{}</indexedColors></colors>\n",
                colors.concat()
            ));
        }
        s.push_str("</styleSheet>\n");
        s
    }
}

/// A sheet being written: its view, its defaults, its columns, its rows of
/// cells, its merges and its conditional formats.
struct Part {
    name: &'static str,
    gridlines: bool,
    format_pr: String,
    cols: Vec<String>,
    /// Each row's attributes and its cells by column.
    rows: BTreeMap<u32, (String, BTreeMap<u32, String>)>,
    merges: Vec<String>,
    conditional: Vec<String>,
    /// Each case's name and its reads, in order.
    cases: Vec<(String, String)>,
}

impl Part {
    fn new(name: &'static str) -> Part {
        let mut p = Part {
            name,
            gridlines: true,
            format_pr: "<sheetFormatPr defaultRowHeight=\"15\"/>".to_string(),
            cols: Vec::new(),
            rows: BTreeMap::new(),
            merges: Vec::new(),
            conditional: Vec::new(),
            cases: Vec::new(),
        };
        p.text(1, 1, "case", 0);
        p.text(1, 2, "reads", 0);
        p
    }

    /// The row the next case takes.
    fn next_row(&self) -> u32 {
        2 * (self.cases.len() as u32 + 1)
    }

    /// A case, its name in column A and its reads in column B: its row.
    fn case(&mut self, name: &str, reads: &str) -> u32 {
        let row = self.next_row();
        self.cases.push((name.to_string(), reads.to_string()));
        self.text(row, 1, name, 0);
        self.text(row, 2, reads, 0);
        row
    }

    fn put(&mut self, row: u32, col: u32, s: u32, kind: &str, body: &str) {
        let at = cell_ref(row, col);
        let s = if s == 0 {
            String::new()
        } else {
            format!(" s=\"{s}\"")
        };
        let element = if body.is_empty() {
            format!("<c r=\"{at}\"{s}{kind}/>")
        } else {
            format!("<c r=\"{at}\"{s}{kind}>{body}</c>")
        };
        self.rows.entry(row).or_default().1.insert(col, element);
    }

    fn text(&mut self, row: u32, col: u32, text: &str, s: u32) {
        let body = format!("<is><t>{}</t></is>", xml::text(text));
        self.put(row, col, s, " t=\"inlineStr\"", &body);
    }

    fn runs(&mut self, row: u32, col: u32, runs: &str, s: u32) {
        self.put(row, col, s, " t=\"inlineStr\"", &format!("<is>{runs}</is>"));
    }

    fn number(&mut self, row: u32, col: u32, value: &str, s: u32) {
        self.put(row, col, s, "", &format!("<v>{value}</v>"));
    }

    fn shared(&mut self, row: u32, col: u32, index: usize, s: u32) {
        self.put(row, col, s, " t=\"s\"", &format!("<v>{index}</v>"));
    }

    fn empty(&mut self, row: u32, col: u32, s: u32) {
        self.put(row, col, s, "", "");
    }

    fn row_attrs(&mut self, row: u32, attrs: &str) {
        self.rows.entry(row).or_default().0 = attrs.to_string();
    }

    fn xml(&self, selected: bool) -> String {
        let bottom = self.rows.keys().next_back().copied().unwrap_or(1);
        let right = self
            .rows
            .values()
            .filter_map(|(_, cells)| cells.keys().next_back().copied())
            .max()
            .unwrap_or(1);
        let mut s = format!("{HEAD}<worksheet xmlns=\"{NS_MAIN}\" xmlns:r=\"{NS_REL}\">\n");
        s.push_str(&format!(
            "<dimension ref=\"A1:{}\"/>\n",
            cell_ref(bottom, right)
        ));
        s.push_str(&format!(
            "<sheetViews><sheetView{}{} workbookViewId=\"0\"/></sheetViews>\n",
            if self.gridlines {
                ""
            } else {
                " showGridLines=\"0\""
            },
            if selected { " tabSelected=\"1\"" } else { "" }
        ));
        s.push_str(&self.format_pr);
        s.push('\n');
        if !self.cols.is_empty() {
            s.push_str(&format!("<cols>{}</cols>\n", self.cols.concat()));
        }
        s.push_str("<sheetData>\n");
        for (number, (attrs, cells)) in &self.rows {
            if cells.is_empty() {
                s.push_str(&format!("<row r=\"{number}\"{attrs}/>\n"));
            } else {
                let cells: String = cells.values().map(String::as_str).collect();
                s.push_str(&format!("<row r=\"{number}\"{attrs}>{cells}</row>\n"));
            }
        }
        s.push_str("</sheetData>\n");
        if !self.merges.is_empty() {
            let merges: String = self
                .merges
                .iter()
                .map(|m| format!("<mergeCell ref=\"{m}\"/>"))
                .collect();
            s.push_str(&format!(
                "<mergeCells count=\"{}\">{merges}</mergeCells>\n",
                self.merges.len()
            ));
        }
        for c in &self.conditional {
            s.push_str(c);
            s.push('\n');
        }
        s.push_str("<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>\n");
        s.push_str("</worksheet>\n");
        s
    }
}

/// A workbook being written: its styles, its sheets, its shared strings and
/// whether it has a theme part.
struct Book {
    styles: Styles,
    parts: Vec<Part>,
    strings: Vec<String>,
    theme: bool,
}

impl Book {
    fn bytes(&self) -> Vec<u8> {
        let n = self.parts.len();
        let mut types = format!(
            "{HEAD}<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n\
             <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n\
             <Default Extension=\"xml\" ContentType=\"application/xml\"/>\n\
             <Override PartName=\"/xl/workbook.xml\" ContentType=\"{CT}.spreadsheetml.sheet.main+xml\"/>\n"
        );
        for i in 1..=n {
            types.push_str(&format!(
                "<Override PartName=\"/xl/worksheets/sheet{i}.xml\" ContentType=\"{CT}.spreadsheetml.worksheet+xml\"/>\n"
            ));
        }
        if self.theme {
            types.push_str(&format!(
                "<Override PartName=\"/xl/theme/theme1.xml\" ContentType=\"{CT}.theme+xml\"/>\n"
            ));
        }
        types.push_str(&format!(
            "<Override PartName=\"/xl/styles.xml\" ContentType=\"{CT}.spreadsheetml.styles+xml\"/>\n"
        ));
        if !self.strings.is_empty() {
            types.push_str(&format!(
                "<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"{CT}.spreadsheetml.sharedStrings+xml\"/>\n"
            ));
        }
        types.push_str(&format!(
            "<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>\n\
             <Override PartName=\"/docProps/app.xml\" ContentType=\"{CT}.extended-properties+xml\"/>\n\
             </Types>\n"
        ));
        let package_rels = format!(
            "{HEAD}<Relationships xmlns=\"{NS_PKG}\">\n\
             <Relationship Id=\"rId1\" Type=\"{REL}/officeDocument\" Target=\"xl/workbook.xml\"/>\n\
             <Relationship Id=\"rId2\" Type=\"{NS_PKG}/metadata/core-properties\" Target=\"docProps/core.xml\"/>\n\
             <Relationship Id=\"rId3\" Type=\"{REL}/extended-properties\" Target=\"docProps/app.xml\"/>\n\
             </Relationships>\n"
        );
        let mut workbook = format!(
            "{HEAD}<workbook xmlns=\"{NS_MAIN}\" xmlns:r=\"{NS_REL}\">\n<bookViews><workbookView activeTab=\"0\"/></bookViews>\n<sheets>\n"
        );
        let mut rels = format!("{HEAD}<Relationships xmlns=\"{NS_PKG}\">\n");
        for (i, p) in self.parts.iter().enumerate() {
            workbook.push_str(&format!(
                "<sheet name=\"{}\" sheetId=\"{}\" r:id=\"rId{}\"/>\n",
                p.name,
                i + 1,
                i + 1
            ));
            rels.push_str(&format!(
                "<Relationship Id=\"rId{}\" Type=\"{REL}/worksheet\" Target=\"worksheets/sheet{}.xml\"/>\n",
                i + 1,
                i + 1
            ));
        }
        workbook.push_str("</sheets>\n</workbook>\n");
        rels.push_str(&format!(
            "<Relationship Id=\"rId{}\" Type=\"{REL}/styles\" Target=\"styles.xml\"/>\n",
            n + 1
        ));
        if self.theme {
            rels.push_str(&format!(
                "<Relationship Id=\"rId{}\" Type=\"{REL}/theme\" Target=\"theme/theme1.xml\"/>\n",
                n + 2
            ));
        }
        if !self.strings.is_empty() {
            rels.push_str(&format!(
                "<Relationship Id=\"rId{}\" Type=\"{REL}/sharedStrings\" Target=\"sharedStrings.xml\"/>\n",
                n + 3
            ));
        }
        rels.push_str("</Relationships>\n");
        let mut parts: Vec<(String, Vec<u8>)> = vec![
            ("[Content_Types].xml".to_string(), types.into_bytes()),
            ("_rels/.rels".to_string(), package_rels.into_bytes()),
            (
                "docProps/app.xml".to_string(),
                format!("{HEAD}<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\"><Application>Frazaro</Application></Properties>\n").into_bytes(),
            ),
            (
                "docProps/core.xml".to_string(),
                format!("{HEAD}<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\"/>\n").into_bytes(),
            ),
            ("xl/workbook.xml".to_string(), workbook.into_bytes()),
            ("xl/_rels/workbook.xml.rels".to_string(), rels.into_bytes()),
            ("xl/styles.xml".to_string(), self.styles.part().into_bytes()),
        ];
        if self.theme {
            parts.push(("xl/theme/theme1.xml".to_string(), THEME.as_bytes().to_vec()));
        }
        if !self.strings.is_empty() {
            let items: String = self
                .strings
                .iter()
                .map(|s| format!("<si>{s}</si>"))
                .collect();
            parts.push((
                "xl/sharedStrings.xml".to_string(),
                format!(
                    "{HEAD}<sst xmlns=\"{NS_MAIN}\" count=\"{0}\" uniqueCount=\"{0}\">{items}</sst>\n",
                    self.strings.len()
                )
                .into_bytes(),
            ));
        }
        for (i, p) in self.parts.iter().enumerate() {
            parts.push((
                format!("xl/worksheets/sheet{}.xml", i + 1),
                p.xml(i == 0).into_bytes(),
            ));
        }
        zip::write_stored(&parts).expect("the fixture fits a zip")
    }
}

/// One color as a fill (column C), a font's color (E) and an edge's (G),
/// and the fill's theme as Excel names it where the color is a theme's.
fn color_case(part: &mut Part, styles: &mut Styles, name: &str, color: &str) {
    let row = part.next_row();
    let mut reads = format!("cell:C{row} cell:E{row} cell:G{row}");
    if color.starts_with("theme=") {
        reads.push_str(&format!(" theme:C{row}"));
    }
    part.case(name, &reads);
    let fill = styles.xf(Xf {
        fill: Some(solid(color)),
        ..Xf::default()
    });
    let font = styles.font("", &format!("<color {color}/>"));
    let font = styles.xf(Xf {
        font: Some(font),
        ..Xf::default()
    });
    let edge = styles.xf(Xf {
        border: Some(left_edge("thin", color)),
        ..Xf::default()
    });
    part.text(row, 3, "Aa 123", fill);
    part.text(row, 5, "Aa 123", font);
    part.text(row, 7, "Aa 123", edge);
}

/// A case of one cell, `Aa 123` at column C in a format of its own.
fn cell_case(part: &mut Part, styles: &mut Styles, name: &str, xf: Xf) {
    let row = part.next_row();
    part.case(name, &format!("cell:C{row}"));
    let s = styles.xf(xf);
    part.text(row, 3, "Aa 123", s);
}

/// A case of one column's settings, its value at its own column.
fn column_case(part: &mut Part, name: &str, col: u32, attrs: &str, extra: &str) -> u32 {
    let row = part.next_row();
    let letters = column_letters(col);
    let reads = if extra.is_empty() {
        format!("col:{letters}")
    } else {
        format!("col:{letters} {extra}")
    };
    part.case(name, &reads);
    part.cols
        .push(format!("<col min=\"{col}\" max=\"{col}\" {attrs}/>"));
    part.number(row, col, "1", 0);
    row
}

/// The first workbook.
fn looks() -> Book {
    let mut styles = Styles::new("11", CALIBRI, "<color theme=\"1\"/>");
    let mut parts = Vec::new();
    let mut strings = Vec::new();

    // Colors: each kind a package names, as a fill, an ink and an edge.
    let mut p = Part::new("Colors");
    for (slot, name) in THEME_SLOTS.iter().enumerate() {
        color_case(
            &mut p,
            &mut styles,
            &format!("theme {slot} ({name})"),
            &format!("theme=\"{slot}\""),
        );
    }
    for (slot, name) in THEME_SLOTS.iter().enumerate() {
        for tint in palette_tints(slot) {
            color_case(
                &mut p,
                &mut styles,
                &format!("theme {slot} ({name}) tint {tint}"),
                &format!("theme=\"{slot}\" tint=\"{tint}\""),
            );
        }
    }
    for (slot, name) in THEME_SLOTS.iter().enumerate() {
        for tint in ["0.3", "-0.6"] {
            color_case(
                &mut p,
                &mut styles,
                &format!("theme {slot} ({name}) tint {tint}, no palette's"),
                &format!("theme=\"{slot}\" tint=\"{tint}\""),
            );
        }
    }
    for tint in ["1", "-1"] {
        color_case(
            &mut p,
            &mut styles,
            &format!("theme 4 (accent1) tint {tint}, the most"),
            &format!("theme=\"4\" tint=\"{tint}\""),
        );
    }
    for i in 0..66 {
        color_case(
            &mut p,
            &mut styles,
            &format!("indexed {i}"),
            &format!("indexed=\"{i}\""),
        );
    }
    color_case(
        &mut p,
        &mut styles,
        "rgb with an alpha of 80",
        "rgb=\"80FF0000\"",
    );
    color_case(
        &mut p,
        &mut styles,
        "rgb with an alpha of 00",
        "rgb=\"00FF0000\"",
    );
    color_case(&mut p, &mut styles, "rgb black", "rgb=\"FF000000\"");
    color_case(&mut p, &mut styles, "auto", "auto=\"1\"");
    // A tint's quantum, the fold's second run: tints a tint in hundredths,
    // in thousandths and in ten-thousandths each draw apart, a gray of each
    // kind and an accent of each.
    for (slot, tint) in [(1, "0.1249"), (0, "-0.1251"), (4, "0.1249"), (5, "-0.1251")] {
        color_case(
            &mut p,
            &mut styles,
            &format!(
                "theme {slot} ({}) tint {tint}, a tint's quantum",
                THEME_SLOTS[slot]
            ),
            &format!("theme=\"{slot}\" tint=\"{tint}\""),
        );
    }
    parts.push(p);

    // Number formats: every built-in id with no definition, two ids with
    // none at all, and codes the part defines, some the language does not
    // draw.
    let mut p = Part::new("Formats");
    for id in (0..=81).chain([163, 200]) {
        let row = p.next_row();
        p.case(&format!("built-in {id}"), &format!("cell:C{row}"));
        let s = styles.xf(Xf {
            num_fmt: id,
            ..Xf::default()
        });
        p.number(row, 3, VALUE, s);
    }
    for code in [
        "0.0%",
        "\"$\"#,##0.00",
        "yyyy-mm-dd hh:mm:ss",
        "mmmm d, yyyy",
        "[h]:mm",
        "0;-0;;@",
        "\"<\"0\">\"",
        "@",
        "General",
        "[Red]0.00",
        "[>=100]0;0.00",
        "[$-409]m/d/yyyy",
        "0.00E+00",
        "# ?/?",
        "_(0_)",
        "0*-",
    ] {
        let row = p.next_row();
        p.case(&format!("code {code}"), &format!("cell:C{row}"));
        let id = styles.code(code);
        let s = styles.xf(Xf {
            num_fmt: id,
            ..Xf::default()
        });
        p.number(row, 3, VALUE, s);
    }
    parts.push(p);

    // The apply flags and a cell style: Probe says a number format, a font,
    // a fill, a border and an alignment, and each cell's own entry says
    // others.
    let mut p = Part::new("Flags");
    let italic = styles.font("<i/>", "<color rgb=\"FF00B050\"/>");
    let probe_font = place(&mut styles.fonts, italic);
    let probe_fill = place(&mut styles.fills, solid("rgb=\"FFFFFF00\""));
    let probe_border = place(&mut styles.borders, left_edge("medium", "rgb=\"FFFF0000\""));
    styles.style_xfs.push(format!(
        "<xf numFmtId=\"10\" fontId=\"{probe_font}\" fillId=\"{probe_fill}\" borderId=\"{probe_border}\"{}><alignment horizontal=\"right\"/></xf>",
        flags(&[])
    ));
    styles
        .cell_styles
        .push("<cellStyle name=\"Probe\" xfId=\"1\"/>".to_string());
    let own = Xf {
        num_fmt: 2,
        font: Some(styles.font("<b/>", styles.ink)),
        fill: Some(solid("rgb=\"FF5B9BD5\"")),
        border: Some(left_edge("thin", "auto=\"1\"")),
        alignment: Some("horizontal=\"center\"".to_string()),
        style: 1,
        ..Xf::default()
    };
    let probe = Xf {
        num_fmt: 10,
        font: Some(styles.font("<i/>", "<color rgb=\"FF00B050\"/>")),
        fill: Some(solid("rgb=\"FFFFFF00\"")),
        border: Some(left_edge("medium", "rgb=\"FFFF0000\"")),
        alignment: Some("horizontal=\"right\"".to_string()),
        style: 1,
        apply: Some(flags(&[])),
        ..Xf::default()
    };
    let own_with = |apply: String, style: u32| Xf {
        apply: Some(apply),
        style,
        ..own.clone()
    };
    let mut flag_cases = vec![(
        "the entry's own, every flag on".to_string(),
        own_with(flags(&[]), 1),
    )];
    for flag in FLAGS {
        flag_cases.push((format!("{flag} off"), own_with(flags(&[flag]), 1)));
    }
    flag_cases.push(("every flag off".to_string(), own_with(flags(&FLAGS), 1)));
    flag_cases.push(("no flag written".to_string(), own_with(String::new(), 1)));
    flag_cases.push(("the style's own".to_string(), probe));
    let beside = "the entry's own beside Normal, no flag written".to_string();
    flag_cases.push((beside, own_with(String::new(), 0)));
    for (name, xf) in flag_cases {
        let name = name.as_str();
        let row = p.next_row();
        p.case(name, &format!("cell:C{row}"));
        let s = styles.xf(xf);
        p.number(row, 3, "0.5", s);
    }
    parts.push(p);

    // What the record has no word for, each alone.
    let mut p = Part::new("Unheld");
    let ink = styles.ink;
    let arial =
        "<font><sz val=\"11\"/><color theme=\"1\"/><name val=\"Arial\"/><family val=\"2\"/></font>";
    let mut unheld = vec![("family Arial", with_font(arial.to_string()))];
    for (name, flag) in [
        ("underline single", "<u/>"),
        ("underline double", "<u val=\"double\"/>"),
        (
            "underline single accounting",
            "<u val=\"singleAccounting\"/>",
        ),
        ("strike", "<strike/>"),
        ("superscript", "<vertAlign val=\"superscript\"/>"),
        ("subscript", "<vertAlign val=\"subscript\"/>"),
        ("outline", "<outline/>"),
        ("shadow", "<shadow/>"),
        ("condense", "<condense/>"),
        ("extend", "<extend/>"),
    ] {
        unheld.push((name, with_font(styles.font(flag, ink))));
    }
    let pattern = |kind: &str| {
        format!("<fill><patternFill patternType=\"{kind}\"><fgColor rgb=\"FFFF0000\"/><bgColor rgb=\"FF0000FF\"/></patternFill></fill>")
    };
    unheld.push(("pattern gray125", with_fill(pattern("gray125"))));
    unheld.push(("pattern darkGrid", with_fill(pattern("darkGrid"))));
    let gradient = "<fill><gradientFill degree=\"90\"><stop position=\"0\"><color theme=\"0\"/></stop><stop position=\"1\"><color theme=\"4\"/></stop></gradientFill></fill>";
    unheld.push(("gradient", with_fill(gradient.to_string())));
    for (name, alignment) in [
        ("horizontal fill", "horizontal=\"fill\""),
        ("horizontal justify", "horizontal=\"justify\""),
        (
            "horizontal centerContinuous",
            "horizontal=\"centerContinuous\"",
        ),
        ("horizontal distributed", "horizontal=\"distributed\""),
        ("vertical justify", "vertical=\"justify\""),
        ("vertical distributed", "vertical=\"distributed\""),
        ("indent 2", "horizontal=\"left\" indent=\"2\""),
        ("rotation 45", "textRotation=\"45\""),
        ("rotation 255, stacked", "textRotation=\"255\""),
        ("shrink to fit", "shrinkToFit=\"1\""),
        ("reading order right to left", "readingOrder=\"2\""),
    ] {
        unheld.push((name, aligned(alignment)));
    }
    let diagonal = |way: &str| {
        format!("<border diagonal{way}=\"1\"><left/><right/><top/><bottom/><diagonal style=\"thin\"><color auto=\"1\"/></diagonal></border>")
    };
    unheld.push(("diagonal up", bordered(diagonal("Up"))));
    unheld.push(("diagonal down", bordered(diagonal("Down"))));
    for (name, xf) in unheld {
        cell_case(&mut p, &mut styles, name, xf);
    }
    let bold_run = "<r><rPr><b/><sz val=\"11\"/><color theme=\"1\"/><rFont val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></rPr><t>Aa</t></r><r><t xml:space=\"preserve\"> 123</t></r>";
    let row = p.next_row();
    p.case("rich text, inline", &format!("cell:C{row}"));
    p.runs(row, 3, bold_run, 0);
    let row = p.next_row();
    p.case("rich text, shared", &format!("cell:C{row}"));
    strings.push(bold_run.to_string());
    p.shared(row, 3, 0, 0);
    cell_case(
        &mut p,
        &mut styles,
        "protection unlocked, no look",
        Xf {
            protection: Some("locked=\"0\"".to_string()),
            ..Xf::default()
        },
    );
    parts.push(p);

    // Widths, each a column of its own from D, its value at its column.
    let mut p = Part::new("Columns");
    let mut col = 4;
    let custom = " customWidth=\"1\"";
    let hidden = " hidden=\"1\" customWidth=\"1\"";
    let best_fit = " bestFit=\"1\" customWidth=\"1\"";
    for (name, width, rest) in [
        ("11.2109375, the writer's 10.5", "11.2109375", custom),
        ("10, off a whole pixel", "10", custom),
        (
            "8.43, the dialog's number written as a file's",
            "8.43",
            custom,
        ),
        ("9.140625, Excel's default", "9.140625", custom),
        ("72.7109375, the writer's 72", "72.7109375", custom),
        ("72.7265625, Excel's 72 saved at 150%", "72.7265625", custom),
        ("20.7109375, Excel's own 20", "20.7109375", custom),
        ("0.5, below a character", "0.5", custom),
        ("1", "1", custom),
        ("1.7109375, a character", "1.7109375", custom),
        ("255, the most", "255", custom),
        ("20, hidden", "20", hidden),
        ("0 wide", "0", custom),
        ("15, its custom width unsaid", "15", ""),
        ("12, best fit", "12", best_fit),
    ] {
        column_case(&mut p, name, col, &format!("width=\"{width}\"{rest}"), "");
        col += 1;
    }
    // A span of three, one element.
    let row = p.next_row();
    let (a, b, c) = (
        column_letters(col),
        column_letters(col + 1),
        column_letters(col + 2),
    );
    p.case("15, a span of three", &format!("col:{a} col:{b} col:{c}"));
    p.cols.push(format!(
        "<col min=\"{col}\" max=\"{}\" width=\"15\" customWidth=\"1\"/>",
        col + 2
    ));
    for k in 0..3 {
        p.number(row, col + k, "1", 0);
    }
    col += 3;
    // A column's format, drawn on a cell the file does not hold and not on
    // the cell at format 0 it does.
    let amber = styles.xf(Xf {
        font: Some(styles.font("<b/>", styles.ink)),
        fill: Some(solid("rgb=\"FFFFC000\"")),
        ..Xf::default()
    });
    let letters = column_letters(col);
    let row = p.next_row();
    column_case(
        &mut p,
        "a column's format",
        col,
        &format!("width=\"9.140625\" style=\"{amber}\""),
        &format!("cell:{letters}{row} blank:{letters}{}", row + 1),
    );
    col += 1;
    let green = styles.xf(Xf {
        fill: Some(solid("rgb=\"FF92D050\"")),
        ..Xf::default()
    });
    let letters = column_letters(col);
    let row = p.next_row();
    column_case(
        &mut p,
        "a column's format and its width 15",
        col,
        &format!("width=\"15\" style=\"{green}\" customWidth=\"1\""),
        &format!("blank:{letters}{}", row + 1),
    );
    parts.push(p);

    // Heights, each a row of its own, its text at column C.
    let mut p = Part::new("Rows");
    let lilac = styles.xf(Xf {
        fill: Some(solid("rgb=\"FFCCC0DA\"")),
        ..Xf::default()
    });
    let custom = |ht: &str| format!(" ht=\"{ht}\" customHeight=\"1\"");
    for (name, attrs, format) in [
        ("custom 30", custom("30"), false),
        ("custom 21", custom("21"), false),
        ("custom 13.5", custom("13.5"), false),
        ("custom 6", custom("6"), false),
        ("custom 408, near the most", custom("408"), false),
        ("custom 409, the most", custom("409"), false),
        ("custom 20, off a pixel at 96 dots", custom("20"), false),
        (
            "custom 20.25, off a pixel at 144 dots",
            custom("20.25"),
            false,
        ),
        ("fitted 45, one line", " ht=\"45\"".to_string(), false),
        ("fitted 10, one line", " ht=\"10\"".to_string(), false),
        ("no height", String::new(), false),
        ("hidden", " hidden=\"1\"".to_string(), false),
        (
            "hidden, custom 30",
            " ht=\"30\" hidden=\"1\" customHeight=\"1\"".to_string(),
            false,
        ),
        (
            "a row's format",
            format!(" s=\"{lilac}\" customFormat=\"1\""),
            true,
        ),
        (
            "a row's format, its customFormat unsaid",
            format!(" s=\"{lilac}\""),
            true,
        ),
    ] {
        let row = p.next_row();
        let reads = if format {
            format!("row:{row} cell:C{row} blank:E{row}")
        } else {
            format!("row:{row}")
        };
        p.case(name, &reads);
        p.row_attrs(row, &attrs);
        p.text(row, 3, "Aa 123", 0);
    }
    p.case("the sheet's standard", "std");
    parts.push(p);

    // A sheet with no gridlines: a cell that keeps a format and no value, a
    // merge and conditional formats.
    let mut p = Part::new("Cells");
    p.gridlines = false;
    p.case("gridlines off", "grid");
    let yellow = styles.xf(Xf {
        fill: Some(solid("rgb=\"FFFFFF00\"")),
        ..Xf::default()
    });
    let row = p.next_row();
    p.case(
        "a cell that keeps a format and no value",
        &format!("empty:C{row}"),
    );
    p.empty(row, 3, yellow);
    let row = p.next_row();
    p.case("a merged range", &format!("merge:C{row} cell:C{row}"));
    p.text(row, 3, "Aa 123", yellow);
    p.merges.push(format!("C{row}:E{row}"));
    styles.dxfs.push(
        "<dxf><fill><patternFill><bgColor rgb=\"FFFF0000\"/></patternFill></fill></dxf>"
            .to_string(),
    );
    styles.dxfs.push("<dxf><font><b/></font></dxf>".to_string());
    let row = p.next_row();
    p.case("a conditional format", &format!("cf:C{row} cell:C{row}"));
    p.number(row, 3, "1", 0);
    p.conditional.push(format!(
        "<conditionalFormatting sqref=\"C{row}\"><cfRule type=\"cellIs\" dxfId=\"0\" priority=\"1\" operator=\"equal\"><formula>1</formula></cfRule></conditionalFormatting>"
    ));
    let row = p.next_row();
    p.case(
        "a conditional format of two rules",
        &format!("cf:C{row} cell:C{row}"),
    );
    p.number(row, 3, "2", 0);
    p.conditional.push(format!(
        "<conditionalFormatting sqref=\"C{row}\"><cfRule type=\"cellIs\" dxfId=\"0\" priority=\"2\" operator=\"equal\"><formula>2</formula></cfRule><cfRule type=\"expression\" dxfId=\"1\" priority=\"3\"><formula>TRUE</formula></cfRule></conditionalFormatting>"
    ));
    parts.push(p);

    // A sheet's own defaults: a base width and a default height a person
    // set, and on the next sheet a default width.
    let mut p = Part::new("Defaults");
    p.format_pr = "<sheetFormatPr baseColWidth=\"10\" defaultRowHeight=\"20\" customHeight=\"1\"/>"
        .to_string();
    let row = p.next_row();
    p.case(
        "a base width of 10 and a default height of 20",
        &format!("std grid col:D row:{row}"),
    );
    p.number(row, 4, "1", 0);
    parts.push(p);
    let mut p = Part::new("Defaults2");
    p.format_pr =
        "<sheetFormatPr defaultColWidth=\"12.7109375\" defaultRowHeight=\"15\"/>".to_string();
    let row = p.next_row();
    p.case("a default width of 12", "std col:D");
    p.number(row, 4, "1", 0);
    parts.push(p);

    Book {
        styles,
        parts,
        strings,
        theme: true,
    }
}

/// The second workbook's own palette: an index's color, none of them the
/// default palette's at the same index.
fn own_palette(i: u32) -> String {
    format!(
        "FF{:02X}{:02X}{:02X}",
        (i * 4) % 256,
        255 - i * 3,
        (i * 37) % 256
    )
}

/// The second workbook: its own palette, no theme part, Arial 10.
fn looks_book() -> Book {
    let mut styles = Styles::new("10", ARIAL, "");
    styles.palette = (0..64).map(own_palette).collect();
    let mut parts = Vec::new();

    let mut p = Part::new("Palette");
    for i in [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 16, 30, 40, 63, 64, 65] {
        color_case(
            &mut p,
            &mut styles,
            &format!("indexed {i}"),
            &format!("indexed=\"{i}\""),
        );
    }
    parts.push(p);

    let mut p = Part::new("Theme");
    for (slot, name) in THEME_SLOTS.iter().enumerate() {
        color_case(
            &mut p,
            &mut styles,
            &format!("theme {slot} ({name}) with no theme part"),
            &format!("theme=\"{slot}\""),
        );
    }
    color_case(
        &mut p,
        &mut styles,
        "theme 4 (accent1) tint 0.39997558519241921 with no theme part",
        "theme=\"4\" tint=\"0.39997558519241921\"",
    );
    parts.push(p);

    let mut p = Part::new("Fonts");
    let row = p.next_row();
    p.case("the default font, Arial 10", &format!("cell:C{row}"));
    p.text(row, 3, "Aa 123", 0);
    for (name, font) in [
        (
            "Arial 10, its color said automatic",
            styles.font("", "<color auto=\"1\"/>"),
        ),
        (
            "Arial 11",
            "<font><sz val=\"11\"/><name val=\"Arial\"/><family val=\"2\"/></font>".to_string(),
        ),
        (
            "Calibri 10",
            "<font><sz val=\"10\"/><name val=\"Calibri\"/><family val=\"2\"/></font>".to_string(),
        ),
        ("bold", styles.font("<b/>", "")),
        (
            "the theme's ink with no theme part",
            styles.font("", "<color theme=\"1\"/>"),
        ),
    ] {
        cell_case(
            &mut p,
            &mut styles,
            name,
            Xf {
                font: Some(font),
                ..Xf::default()
            },
        );
    }
    parts.push(p);

    let mut p = Part::new("Widths");
    for (k, (name, width)) in [
        ("11.2109375 in Arial 10", "11.2109375"),
        ("72.7109375 in Arial 10", "72.7109375"),
        ("9.140625 in Arial 10", "9.140625"),
    ]
    .into_iter()
    .enumerate()
    {
        column_case(
            &mut p,
            name,
            4 + k as u32,
            &format!("width=\"{width}\" customWidth=\"1\""),
            "",
        );
    }
    p.case("the sheet's standard", "std");
    parts.push(p);

    Book {
        styles,
        parts,
        strings: Vec::new(),
        theme: false,
    }
}

#[test]
fn the_looks_fixtures_are_their_builder_s() {
    let bless = std::env::var("FRAZARO_BLESS").as_deref() == Ok("1");
    for (path, book) in [(LOOKS, looks()), (BOOK, looks_book())] {
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

// ---- Excel's reading of the fixtures: the golden ---------------------

const GOLDEN: &str = concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../scripts/reflect/looks_golden.txt"
);
const SAVED: &str = concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../scripts/reflect/looks_saved.xlsx"
);
const BOOK_SAVED: &str = concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../scripts/reflect/looks_book_saved.xlsx"
);

/// The object model's number for a file's theme slot, `XlThemeColor`: its
/// index from 1 in the file's own order (Excel's reading; its names, the
/// theme part's order, call the file's slot 0, drawn `lt1`,
/// `xlThemeColorDark1`).
fn object_model_slot(file_slot: usize) -> u32 {
    file_slot as u32 + 1
}

/// How near Excel's tint is to the file's: Excel keeps a tint in 32767ths,
/// truncated (0.1249 as 4092/32767).
const TINT_KEPT: f64 = 1.0 / 32767.0 + 1e-12;

/// A case as the golden holds it: its workbook, its sheet, its name and
/// each target with Excel's reading of it.
#[derive(Debug)]
struct Read {
    file: String,
    sheet: String,
    name: String,
    targets: Vec<(String, String)>,
}

/// The golden: its source, Excel's (`excel <version> <build>`) or the
/// prediction written before Excel's run, and its cases in order.
fn golden() -> (String, Vec<Read>) {
    let text = std::fs::read_to_string(GOLDEN)
        .expect("scripts/reflect/looks_golden.txt: VlaWriteLooksGolden (VLA_Tests.bas) writes it in Excel")
        .replace("\r\n", "\n");
    let mut lines = text.lines();
    let source = lines
        .next()
        .and_then(|l| l.strip_prefix("source\t"))
        .expect("the golden's first line names its source")
        .to_string();
    let mut cases: Vec<Read> = Vec::new();
    for line in lines {
        if let Some(header) = line.strip_prefix("=== ") {
            let mut parts = header.splitn(3, '!');
            let mut next = || {
                parts
                    .next()
                    .expect("a header's file, sheet and name")
                    .to_string()
            };
            let (file, sheet, name) = (next(), next(), next());
            cases.push(Read {
                file,
                sheet,
                name,
                targets: Vec::new(),
            });
        } else if let Some((target, reading)) = line.split_once('\t') {
            cases
                .last_mut()
                .expect("a target after a case's header")
                .targets
                .push((target.to_string(), reading.to_string()));
        } else {
            panic!("the golden's line of no shape: {line}");
        }
    }
    (source, cases)
}

/// What a workbook's reading under the looks says, sheet by sheet.
#[derive(Default)]
struct Reading {
    formats: HashMap<u32, String>,
    styles: HashMap<(String, String), u32>,
    printed: HashSet<(String, String)>,
    /// A column's width, `shown` or `hidden`, and its format.
    columns: HashMap<(String, String), [String; 3]>,
    /// A row's height and `shown` or `hidden`.
    rows: HashMap<(String, u32), [String; 2]>,
    gridlines: HashMap<String, String>,
    stats: LooksStats,
}

/// An atom as the record spells it: a word, or a string in its quotes.
fn atom(form: &Form) -> String {
    match form {
        Form::Sym(s) => s.clone(),
        Form::Str(s) => format!("\"{s}\""),
        Form::List(_) => "(...)".to_string(),
    }
}

/// A string's content or a word's text.
fn content(form: &Form) -> String {
    match form {
        Form::Sym(s) | Form::Str(s) => s.clone(),
        Form::List(_) => "(...)".to_string(),
    }
}

fn reading(bytes: &[u8], label: &str) -> Reading {
    let (text, stats) = crate::reflect::reflect_looks(bytes, label)
        .unwrap_or_else(|r| panic!("{label}: {}", r.text));
    let mut r = Reading {
        stats,
        ..Reading::default()
    };
    for line in text.lines() {
        if let Some(rest) = line.strip_prefix("(format ") {
            let (index, body) = rest.split_once(' ').expect("a format's index");
            let body = body.strip_suffix(')').expect("a format's close");
            r.formats
                .insert(index.parse().expect("an index"), body.to_string());
            continue;
        }
        let forms = vla_lang::reader::read_forms(line).expect("a row of the record");
        let items = &forms[0].as_list().expect("a row is a list").items;
        let word = |i: usize| content(&items[i]);
        match word(0).as_str() {
            "cell" | "formula" => {
                r.printed.insert((word(1), word(2)));
            }
            "style" => {
                r.styles
                    .insert((word(1), word(2)), word(3).parse().expect("an index"));
            }
            "column" => {
                r.columns
                    .insert((word(1), word(2)), [word(3), word(4), word(5)]);
            }
            "row" => {
                r.rows.insert(
                    (word(1), word(2).parse().expect("a row's number")),
                    [word(3), word(4)],
                );
            }
            "gridlines" => {
                r.gridlines.insert(word(1), word(2));
            }
            _ => {}
        }
    }
    r
}

/// A format's text as what it draws, part by part: its fill, its number
/// format, its wrap, its font's words, its alignment's and its edges. The
/// defaults' ink, black in both workbooks, is unsaid, a said black and an
/// unsaid one drawing alike; with `colors` false every color is left out,
/// as a workbook with no theme part leaves a theme's color unsaid.
#[derive(Debug, PartialEq, Default)]
struct Drawn {
    fill: String,
    num: String,
    wrap: String,
    font: Vec<String>,
    align: Vec<String>,
    edges: Vec<String>,
}

fn drawn(text: &str, colors: bool) -> Drawn {
    let forms = vla_lang::reader::read_forms(&format!("(f {text})"))
        .unwrap_or_else(|_| panic!("a format of the record's spelling: {text}"));
    let items = &forms[0].as_list().expect("a list").items;
    let color = |form: &Form| -> Option<String> {
        match form {
            Form::Str(c) if colors => Some(format!("\"{c}\"")),
            _ => None,
        }
    };
    let num = match &items[2] {
        Form::Str(code) => format!("\"{}\"", as_excel_shows(code)),
        other => atom(other),
    };
    let mut d = Drawn {
        fill: color(&items[1]).unwrap_or_else(|| "none".to_string()),
        num,
        wrap: atom(&items[3]),
        ..Drawn::default()
    };
    for group in &items[4..] {
        let group = &group.as_list().expect("a group").items;
        match content(&group[0]).as_str() {
            "font" => {
                for item in &group[1..] {
                    match item {
                        Form::Str(c) if c == "000000" => {}
                        Form::Str(_) => d.font.extend(color(item)),
                        _ => d.font.push(atom(item)),
                    }
                }
            }
            "align" => d.align = group[1..].iter().map(atom).collect(),
            _ => {
                for edge in &group[1..] {
                    let edge = &edge.as_list().expect("an edge").items;
                    let mut e = format!("{} {}", atom(&edge[0]), atom(&edge[1]));
                    if let Some(c) = edge.get(2).and_then(color) {
                        e.push(' ');
                        e.push_str(&c);
                    }
                    d.edges.push(e);
                }
            }
        }
    }
    d
}

/// The characters a number format draws as themselves, which a file may
/// quote alone or escape and Excel's object model shows bare.
const BARE: &str = "$-+/():!^&'~{}<>= ,";

/// A code as Excel's object model shows it beside a file's spelling of it:
/// a character drawn as itself, quoted alone or escaped by a backslash,
/// bare (Excel's save writes `yyyy\-mm\-dd` and `"$"#,##0.00`, its object
/// model shows `yyyy-mm-dd` and `$#,##0.00`), so that the reader's code, the
/// file's, and Excel's reading of it compare as what they draw. Both sides
/// are read so; an unescaped comma scales a number, which no code here
/// writes where an escaped one is.
fn as_excel_shows(code: &str) -> String {
    let chars: Vec<char> = code.chars().collect();
    let mut out = String::new();
    let mut i = 0;
    while i < chars.len() {
        let bare = |c: char| BARE.contains(c);
        if chars[i] == '\\' && chars.get(i + 1).is_some_and(|c| bare(*c)) {
            out.push(chars[i + 1]);
            i += 2;
        } else if chars[i] == '"' && chars.get(i + 2) == Some(&'"') && bare(chars[i + 1]) {
            out.push(chars[i + 1]);
            i += 3;
        } else {
            out.push(chars[i]);
            i += 1;
        }
    }
    out
}

/// The code a format's text names, as written, where it names one.
fn code_of(text: &str) -> Option<String> {
    let forms = vla_lang::reader::read_forms(&format!("(f {text})")).ok()?;
    match forms[0].as_list()?.items.get(2)? {
        Form::Str(code) => Some(code.clone()),
        _ => None,
    }
}

/// A column's width in characters at a screen's pixels a digit and its
/// padding, by Excel's rule (ECMA-376 Part 1, §18.3.1.13): the file's
/// width to whole pixels, then to characters in hundredths, a column
/// narrower than one character counted in its own pixels.
fn width_at(width: f64, digit: f64, pad: f64) -> f64 {
    let px = ((256.0 * width + (128.0 / digit).trunc()) / 256.0 * digit).trunc();
    let chars = if px < digit + pad {
        px / (digit + pad)
    } else {
        (px - pad) / digit
    };
    (chars * 100.0 + 0.5).trunc() / 100.0
}

/// A part of a package as text, stored as the builder stores it or
/// deflated as Excel saves it.
fn part_text(bytes: &[u8], name: &str) -> Option<String> {
    let entries = zip::entries(bytes)?;
    let e = entries.iter().find(|e| e.name == name)?;
    let raw = zip::raw_data(bytes, e)?;
    let data = match e.method {
        0 => raw.to_vec(),
        8 => crate::sheet::inflate::inflate(raw, e.size as usize, 1 << 26).ok()?,
        _ => return None,
    };
    String::from_utf8(data).ok()
}

/// An attribute's value in an element's opening tag.
fn attr<'t>(element: &'t str, name: &str) -> Option<&'t str> {
    let key = format!(" {name}=\"");
    let at = element.find(&key)? + key.len();
    element[at..].split('"').next()
}

/// Each sheet's columns as the package writes them: a sheet's name and a
/// column's index, its width and whether it is hidden.
fn col_specs(bytes: &[u8]) -> HashMap<(String, u32), (f64, bool)> {
    let workbook = part_text(bytes, "xl/workbook.xml").expect("a workbook part");
    let rels = part_text(bytes, "xl/_rels/workbook.xml.rels").expect("its relationships");
    let mut specs = HashMap::new();
    for (at, _) in workbook.match_indices("<sheet ") {
        let element = &workbook[at..at + workbook[at..].find('>').expect("a tag")];
        let (Some(name), Some(id)) = (attr(element, "name"), attr(element, "r:id")) else {
            continue;
        };
        let target = rels
            .match_indices("<Relationship ")
            .map(|(i, _)| &rels[i..i + rels[i..].find('>').expect("a tag")])
            .find(|r| attr(r, "Id") == Some(id))
            .and_then(|r| attr(r, "Target"))
            .expect("the sheet's relationship");
        let target = target.trim_start_matches('/').trim_start_matches("xl/");
        let sheet = part_text(bytes, &format!("xl/{target}")).expect("the sheet's part");
        for (i, _) in sheet.match_indices("<col ") {
            let col = &sheet[i..i + sheet[i..].find('>').expect("a tag")];
            let number = |n: &str| attr(col, n).and_then(|v| v.parse::<u32>().ok());
            let (Some(min), Some(max)) = (number("min"), number("max")) else {
                continue;
            };
            let width: f64 = attr(col, "width")
                .and_then(|v| v.parse().ok())
                .unwrap_or(0.0);
            let hidden = matches!(attr(col, "hidden"), Some("1" | "true"));
            for c in min..=max.min(16_384) {
                specs.insert((name.to_string(), c), (width, hidden));
            }
        }
    }
    specs
}

/// A column's index from its letters.
fn column_number(letters: &str) -> u32 {
    letters
        .bytes()
        .fold(0, |n, b| n * 26 + u32::from(b - b'A' + 1))
}

/// What a workbook's builder says the reader counts that no target of the
/// golden shows: the rows formatted whole, and each sheet's own default
/// width or height, by the sheet's name.
fn builder_counts(book: &Book) -> (u64, HashMap<&'static str, u64>) {
    let mut row_formats = 0;
    let mut defaults = HashMap::new();
    for p in &book.parts {
        row_formats += p
            .rows
            .values()
            .filter(|(attrs, _)| attrs.contains("customFormat=\"1\""))
            .count() as u64;
        let pr = &p.format_pr;
        let own = u64::from(pr.contains("defaultColWidth="))
            + u64::from(pr.contains("baseColWidth=") && !pr.contains("baseColWidth=\"8\""))
            + u64::from(pr.contains("customHeight=\"1\""));
        defaults.insert(p.name, own);
    }
    (row_formats, defaults)
}

/// Whether the reader counts a case's cell as a look the record cannot
/// hold where Excel's object model shows nothing of it: a condensed or an
/// extended face, for which the object model has no property (labelled:
/// Excel draws neither on Windows).
fn counted_unseen(name: &str) -> bool {
    name == "condense" || name == "extend"
}

/// Every difference between a workbook's reading under the looks and
/// Excel's reading of it, the golden's cases of that workbook; `saved` for
/// Excel's save of it, which writes a merge's covered cells in its format.
fn differences(
    label: &str,
    bytes: &[u8],
    cases: &[&Read],
    book: &Book,
    saved: bool,
) -> Vec<String> {
    let r = reading(bytes, label);
    let specs = col_specs(bytes);
    let (row_formats, defaults) = builder_counts(book);
    // A workbook with no theme part leaves a theme's color unsaid; Excel's
    // save of one writes its own theme, which the reader then reads.
    let themed = part_text(bytes, "xl/_rels/workbook.xml.rels")
        .is_some_and(|rels| rels.contains("relationships/theme\""));
    let mut found = Vec::new();
    let mut want = LooksStats {
        row_formats,
        ..LooksStats::default()
    };
    // A sheet's standard width and height, the measure of its columns and
    // rows of no setting.
    let mut standard: HashMap<String, (f64, f64)> = HashMap::new();
    for case in cases {
        for (target, excel) in &case.targets {
            if target == "std" {
                let (w, h) = excel.split_once(' ').expect("a width and a height");
                standard.insert(
                    case.sheet.clone(),
                    (w.parse().expect("a width"), h.parse().expect("a height")),
                );
                want.unheld += defaults.get(case.sheet.as_str()).copied().unwrap_or(0);
            }
        }
    }
    // The widths Excel read, beside the file's: the screen's rule is found
    // from them below.
    let mut widths: Vec<(f64, f64, String)> = Vec::new();
    for case in cases {
        let no_theme = case.name.contains("no theme part") && !themed;
        for (target, excel) in &case.targets {
            let at = format!("{label}: {}!{}, {target}", case.sheet, case.name);
            let (kind, place) = target.split_once(':').unwrap_or((target.as_str(), ""));
            let key = (case.sheet.clone(), place.to_string());
            match kind {
                "cell" | "empty" | "blank" => {
                    let (held, words) = match excel.split_once(" !") {
                        Some((h, w)) => (h, Some(w)),
                        None => (excel.as_str(), None),
                    };
                    let printed = r.printed.contains(&key);
                    if kind != "cell" {
                        if printed {
                            found.push(format!("{at}: a cell row, where the file holds no value"));
                        }
                        if kind == "empty" {
                            want.formatted_empty += 1;
                            continue;
                        }
                        // No cell: Excel draws its column's format, where
                        // the record's column row names one; a row's own
                        // format the record cannot say.
                        let letters = place.trim_end_matches(|c: char| c.is_ascii_digit());
                        let format = r
                            .columns
                            .get(&(case.sheet.clone(), letters.to_string()))
                            .map(|c| c[2].clone())
                            .filter(|f| f != "none");
                        if let Some(index) = format {
                            let ours = &r.formats[&index.parse::<u32>().expect("an index")];
                            if drawn(ours, true) != drawn(held, true) {
                                found.push(format!("{at}: the column's format {index}, {ours}, where Excel draws {held}"));
                            }
                        }
                        continue;
                    }
                    if !printed {
                        found.push(format!("{at}: no cell row"));
                        continue;
                    }
                    let index = r.styles.get(&key).copied().unwrap_or(0);
                    let ours = &r.formats[&index];
                    if drawn(ours, !no_theme) != drawn(held, !no_theme) {
                        found.push(format!(
                            "{at}: the reader's format {index}, {ours}, where Excel reads {held}"
                        ));
                    }
                    if words.is_some() || no_theme || counted_unseen(&case.name) {
                        want.unheld += 1;
                    }
                    // A code the language does not draw, the reader's own.
                    if code_of(ours).is_some_and(|c| vla_lang::numfmt::Code::parse(&c).is_err()) {
                        want.undrawn += 1;
                    }
                }
                "theme" => {
                    let (slot, tint) = excel.split_once(' ').unwrap_or((excel.as_str(), "0"));
                    let file_slot: usize = case
                        .name
                        .strip_prefix("theme ")
                        .and_then(|s| s.split(' ').next())
                        .and_then(|s| s.parse().ok())
                        .expect("a theme case's slot");
                    if slot.parse::<u32>().ok() != Some(object_model_slot(file_slot)) {
                        found.push(format!(
                            "{at}: Excel names the slot {slot}, where the count predicts {}",
                            object_model_slot(file_slot)
                        ));
                    }
                    let file_tint: f64 = case
                        .name
                        .split(" tint ")
                        .nth(1)
                        .and_then(|t| t.split([',', ' ']).next())
                        .map_or(0.0, |t| t.parse().expect("a tint"));
                    let read: f64 = tint.parse().expect("Excel's tint");
                    if (read - file_tint).abs() > TINT_KEPT {
                        found.push(format!(
                            "{at}: Excel's tint {read}, where the file's is {file_tint}"
                        ));
                    }
                }
                "col" => {
                    let (w, shown) = excel.split_once(' ').expect("a width and its shown");
                    let excel_width: f64 = w.parse().expect("Excel's width");
                    let hidden = shown == "hidden";
                    let ours = r
                        .columns
                        .get(&key)
                        .cloned()
                        .unwrap_or_else(|| ["none".into(), "shown".into(), "none".into()]);
                    if (ours[1] == "hidden") != hidden {
                        found.push(format!(
                            "{at}: the reader's column {}, where Excel's is {shown}",
                            ours[1]
                        ));
                    }
                    match specs.get(&(case.sheet.clone(), column_number(place))) {
                        Some((width, file_hidden)) if !hidden && !file_hidden => {
                            widths.push((*width, excel_width, at.clone()));
                            // The record's width is the rule at its own 7
                            // pixels a character, Excel's at 96 dots.
                            let px = super::looks::width_pixels(*width);
                            let want_width = if px == 64.0 {
                                "none".to_string()
                            } else {
                                number_text(width_at(*width, 7.0, 5.0)).expect("a width")
                            };
                            if ours[0] != want_width {
                                found.push(format!("{at}: the reader's width {}, where the rule at 96 dots gives {want_width}", ours[0]));
                            }
                        }
                        Some(_) => {}
                        None => {
                            if ours[0] != "none" {
                                found.push(format!("{at}: the reader's width {} for a column the file says nothing of", ours[0]));
                            }
                            if let Some((sw, _)) = standard.get(&case.sheet) {
                                if (excel_width - sw).abs() > 0.005 {
                                    found.push(format!("{at}: Excel's width {excel_width}, not its sheet's standard {sw}"));
                                }
                            }
                        }
                    }
                }
                "row" => {
                    let (h, shown) = excel.split_once(' ').expect("a height and its shown");
                    let excel_height: f64 = h.parse().expect("Excel's height");
                    let hidden = shown == "hidden";
                    let number: u32 = place.parse().expect("a row's number");
                    let ours = r
                        .rows
                        .get(&(case.sheet.clone(), number))
                        .cloned()
                        .unwrap_or_else(|| ["none".into(), "shown".into()]);
                    if (ours[1] == "hidden") != hidden {
                        found.push(format!(
                            "{at}: the reader's row {}, where Excel's is {shown}",
                            ours[1]
                        ));
                    }
                    if hidden {
                        continue;
                    }
                    match ours[0].parse::<f64>() {
                        // A height a person set, as the screen's pixels hold
                        // it: within one, three quarters of a point at 96
                        // dots, the coarsest screen.
                        Ok(height) => {
                            if (height - excel_height).abs() > 0.76 {
                                found.push(format!("{at}: the reader's height {height}, where Excel's is {excel_height}"));
                            }
                        }
                        // No height: the host fits the row, as Excel fits a
                        // row of one line to its sheet's standard at the
                        // open, a height the file says it fitted among them.
                        Err(_) => {
                            if let Some((_, sh)) = standard.get(&case.sheet) {
                                if (excel_height - sh).abs() > 0.001 {
                                    found.push(format!("{at}: Excel's height {excel_height}, not its sheet's standard {sh}"));
                                }
                            }
                        }
                    }
                }
                "merge" => {
                    if let Some((first, last)) = excel.split_once(':') {
                        want.merged += 1;
                        // Excel's save writes the cells a merge covers, each
                        // in its first cell's format and with no value.
                        if saved {
                            let area = crate::sheet::parse_a1_range(&format!("{first}:{last}"))
                                .expect("a merge's range");
                            let cells = u64::from(area.bottom - area.top + 1)
                                * u64::from(area.right - area.left + 1);
                            want.formatted_empty += cells - 1;
                        }
                    }
                }
                "cf" => {
                    let rules: u64 = excel
                        .split(' ')
                        .next()
                        .and_then(|n| n.parse().ok())
                        .expect("the rules counted");
                    want.conditional += rules;
                }
                "grid" => {
                    let ours = r.gridlines.get(&case.sheet).map(String::as_str);
                    if ours != Some(excel.as_str()) {
                        found.push(format!(
                            "{at}: the reader's gridlines {ours:?}, where Excel's are {excel}"
                        ));
                    }
                }
                "std" => {}
                _ => found.push(format!("{at}: a target of no kind this test reads")),
            }
        }
    }
    // The screen's rule: a digit's pixels and the padding that give every
    // width Excel read from the file's.
    if !widths.is_empty() {
        let fits = (5..=40u32)
            .flat_map(|d| (3..=20u32).map(move |p| (d, p)))
            .filter(|(d, p)| {
                widths
                    .iter()
                    .all(|(w, e, _)| (width_at(*w, f64::from(*d), f64::from(*p)) - e).abs() < 0.005)
            })
            .count();
        if fits == 0 {
            found.push(format!(
                "{label}: no digit and padding give every width Excel read: {}",
                widths
                    .iter()
                    .map(|(w, e, at)| format!("{at} {w} read {e}"))
                    .collect::<Vec<_>>()
                    .join("; ")
            ));
        }
    }
    let got = r.stats;
    for (what, ours, excel) in [
        ("formatted-empty", got.formatted_empty, want.formatted_empty),
        ("unheld", got.unheld, want.unheld),
        ("undrawn", got.undrawn, want.undrawn),
        ("merged", got.merged, want.merged),
        ("conditional", got.conditional, want.conditional),
        ("row-formats", got.row_formats, want.row_formats),
    ] {
        if ours != excel {
            found.push(format!(
                "{label}: the reader counts {what} {ours}, where Excel's reading makes it {excel}"
            ));
        }
    }
    found
}

/// The reader's reading of each fixture, and of Excel's save of each once
/// the golden is Excel's, held to Excel's object model's reading of every
/// case, target by target (decision 12).
#[test]
fn the_looks_read_each_case_as_excel_reads_it() {
    let (source, cases) = golden();
    let books = [("looks.xlsx", looks()), ("looks_book.xlsx", looks_book())];
    // The golden's cases are the fixtures', in order, each with its reads.
    let fixture: Vec<(String, String, String, Vec<String>)> = books
        .iter()
        .flat_map(|(file, book)| {
            book.parts.iter().flat_map(move |p| {
                p.cases.iter().map(move |(name, reads)| {
                    (
                        file.to_string(),
                        p.name.to_string(),
                        name.clone(),
                        reads.split(' ').map(str::to_string).collect(),
                    )
                })
            })
        })
        .collect();
    assert_eq!(
        cases.len(),
        fixture.len(),
        "the golden's cases and the fixtures'"
    );
    for (g, f) in cases.iter().zip(&fixture) {
        let targets: Vec<String> = g.targets.iter().map(|(t, _)| t.clone()).collect();
        assert_eq!(
            (&g.file, &g.sheet, &g.name, &targets),
            (&f.0, &f.1, &f.2, &f.3),
            "the golden's case beside the fixture's"
        );
    }
    let excel = source.starts_with("excel ");
    let mut found = Vec::new();
    for (file, book) in &books {
        let mine: Vec<&Read> = cases.iter().filter(|c| c.file == *file).collect();
        found.extend(differences(file, &book.bytes(), &mine, book, false));
        if excel {
            let path = if *file == "looks.xlsx" {
                SAVED
            } else {
                BOOK_SAVED
            };
            let saved = std::fs::read(path).unwrap_or_else(|_| {
                panic!("{path}: VlaWriteLooksGolden saves it beside the golden")
            });
            found.extend(differences(
                &format!("{file}, Excel's save"),
                &saved,
                &mine,
                book,
                true,
            ));
        }
    }
    assert!(
        found.is_empty(),
        "{} difference(s) from Excel's reading ({source}):\n{}",
        found.len(),
        found.join("\n")
    );
}
