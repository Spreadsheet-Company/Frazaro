//! The OOXML parts of a workbook (ISO/IEC 29500, SpreadsheetML), rendered
//! from the sheet model as text in a fixed order, then zipped stored
//! (PORT.7, slice 7a).
//!
//! New ground, said plainly: the reference writes through Excel's object
//! model and never spells a part. What is here is the smallest set of parts
//! Excel, Sheets, Calc and Numbers open: the content types, the package
//! relationships, two property parts that name Frazaro and carry no date and
//! no person, the workbook with its `calcPr` (`fullCalcOnLoad`: the host
//! computes every formula when it opens the file, since the core computes
//! nothing), the workbook's relationships, one styles part, and one part per
//! sheet. Cells hold inline strings, so there is no shared-string table, and
//! (slice 7d) adding a sheet to an existing workbook edits no part the
//! model's own cells depend on. tools/build_examples.ps1 writes the same
//! minimal shape in PowerShell for the onboarding samples, and Excel opens
//! those.
//!
//! Every name here is fixed and every order is fixed, so that the same model
//! gives the same bytes: the parts in the order [`parts`] lists them, a
//! sheet's rows and cells in row-then-column order (the model's map), the
//! fills and formats in the order the style table holds them.

use super::xml;
use super::zip;
use super::{cell_ref, column_width_attr, Content, Sheet, Workbook};

const XML_HEAD: &str = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
const NS_MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const NS_REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_PKG_REL: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const NS_CT: &str = "http://schemas.openxmlformats.org/package/2006/content-types";
const NS_CORE: &str = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
const NS_APP: &str = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
const REL_OFFICE_DOCUMENT: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const REL_CORE_PROPS: &str =
    "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
const REL_APP_PROPS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
const REL_WORKSHEET: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
const REL_STYLES: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
const CT_WORKBOOK: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
const CT_WORKSHEET: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
const CT_STYLES: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
const CT_CORE: &str = "application/vnd.openxmlformats-package.core-properties+xml";
const CT_APP: &str = "application/vnd.openxmlformats-officedocument.extended-properties+xml";
const CT_RELS: &str = "application/vnd.openxmlformats-package.relationships+xml";
const CT_XML: &str = "application/xml";

/// The relationship id of the styles part, named so that it collides with
/// nothing Excel writes (`rId1`, `rId2`, ...).
const STYLES_RID: &str = "rIdFrazaroStyles";

/// The part name of the n-th sheet, 1-based, chosen never to collide with a
/// workbook saved by Excel, whose sheets are `sheetN.xml`.
pub fn sheet_part(n: usize) -> String {
    format!("xl/worksheets/frazaro_{n}.xml")
}

fn sheet_rid(n: usize) -> String {
    format!("rIdFrazaro{n}")
}

/// Every part of the workbook, name and bytes, in the order they go into the
/// archive.
pub fn parts(wb: &Workbook) -> Vec<(String, Vec<u8>)> {
    let mut list: Vec<(String, Vec<u8>)> = vec![
        (
            "[Content_Types].xml".to_string(),
            content_types(wb).into_bytes(),
        ),
        ("_rels/.rels".to_string(), package_rels().into_bytes()),
        ("docProps/app.xml".to_string(), app_props().into_bytes()),
        ("docProps/core.xml".to_string(), core_props().into_bytes()),
        ("xl/workbook.xml".to_string(), workbook_xml(wb).into_bytes()),
        (
            "xl/_rels/workbook.xml.rels".to_string(),
            workbook_rels(wb).into_bytes(),
        ),
        ("xl/styles.xml".to_string(), styles_xml(wb).into_bytes()),
    ];
    for (i, sheet) in wb.sheets.iter().enumerate() {
        list.push((
            sheet_part(i + 1),
            sheet_xml(sheet, i == wb.active_sheet).into_bytes(),
        ));
    }
    list
}

/// The workbook as a file's bytes; `None` when it would not fit the zip
/// format's sizes.
pub fn workbook_bytes(wb: &Workbook) -> Option<Vec<u8>> {
    zip::write_stored(&parts(wb))
}

fn content_types(wb: &Workbook) -> String {
    let mut s = String::from(XML_HEAD);
    s.push_str(&format!("<Types xmlns=\"{NS_CT}\">\n"));
    s.push_str(&format!(
        "<Default Extension=\"rels\" ContentType=\"{CT_RELS}\"/>\n"
    ));
    s.push_str(&format!(
        "<Default Extension=\"xml\" ContentType=\"{CT_XML}\"/>\n"
    ));
    s.push_str(&format!(
        "<Override PartName=\"/xl/workbook.xml\" ContentType=\"{CT_WORKBOOK}\"/>\n"
    ));
    for i in 0..wb.sheets.len() {
        s.push_str(&format!(
            "<Override PartName=\"/{}\" ContentType=\"{CT_WORKSHEET}\"/>\n",
            sheet_part(i + 1)
        ));
    }
    s.push_str(&format!(
        "<Override PartName=\"/xl/styles.xml\" ContentType=\"{CT_STYLES}\"/>\n"
    ));
    s.push_str(&format!(
        "<Override PartName=\"/docProps/core.xml\" ContentType=\"{CT_CORE}\"/>\n"
    ));
    s.push_str(&format!(
        "<Override PartName=\"/docProps/app.xml\" ContentType=\"{CT_APP}\"/>\n"
    ));
    s.push_str("</Types>\n");
    s
}

fn package_rels() -> String {
    format!(
        "{XML_HEAD}<Relationships xmlns=\"{NS_PKG_REL}\">\n\
         <Relationship Id=\"rId1\" Type=\"{REL_OFFICE_DOCUMENT}\" Target=\"xl/workbook.xml\"/>\n\
         <Relationship Id=\"rId2\" Type=\"{REL_CORE_PROPS}\" Target=\"docProps/core.xml\"/>\n\
         <Relationship Id=\"rId3\" Type=\"{REL_APP_PROPS}\" Target=\"docProps/app.xml\"/>\n\
         </Relationships>\n"
    )
}

/// The application's name and nothing else: no version string a host would
/// read as its own, no person, no date.
fn app_props() -> String {
    format!("{XML_HEAD}<Properties xmlns=\"{NS_APP}\"><Application>Frazaro</Application></Properties>\n")
}

/// The core properties part, present and empty: no author, no dates, so the
/// bytes carry no clock and no name.
fn core_props() -> String {
    format!("{XML_HEAD}<cp:coreProperties xmlns:cp=\"{NS_CORE}\"/>\n")
}

fn workbook_xml(wb: &Workbook) -> String {
    let mut s = String::from(XML_HEAD);
    s.push_str(&format!(
        "<workbook xmlns=\"{NS_MAIN}\" xmlns:r=\"{NS_REL}\">\n"
    ));
    s.push_str(&format!(
        "<bookViews><workbookView activeTab=\"{}\"/></bookViews>\n",
        wb.active_sheet
    ));
    s.push_str("<sheets>\n");
    for (i, sheet) in wb.sheets.iter().enumerate() {
        s.push_str(&format!(
            "<sheet name=\"{}\" sheetId=\"{}\" r:id=\"{}\"/>\n",
            xml::attr(&sheet.name),
            i + 1,
            sheet_rid(i + 1)
        ));
    }
    s.push_str("</sheets>\n");
    // The host computes every formula when it opens the file (HORIZON.md
    // section 12: the core has no calc engine, on purpose).
    s.push_str("<calcPr fullCalcOnLoad=\"1\"/>\n");
    s.push_str("</workbook>\n");
    s
}

fn workbook_rels(wb: &Workbook) -> String {
    let mut s = String::from(XML_HEAD);
    s.push_str(&format!("<Relationships xmlns=\"{NS_PKG_REL}\">\n"));
    for i in 0..wb.sheets.len() {
        s.push_str(&format!(
            "<Relationship Id=\"{}\" Type=\"{REL_WORKSHEET}\" Target=\"worksheets/frazaro_{}.xml\"/>\n",
            sheet_rid(i + 1),
            i + 1
        ));
    }
    s.push_str(&format!(
        "<Relationship Id=\"{STYLES_RID}\" Type=\"{REL_STYLES}\" Target=\"styles.xml\"/>\n"
    ));
    s.push_str("</Relationships>\n");
    s
}

/// The styles part: one font, the two fills the format requires first (none
/// and gray125) then each fill the model uses, one empty border, and one
/// cell format per entry of the style table, in its order.
fn styles_xml(wb: &Workbook) -> String {
    let fills = wb.styles.fills();
    let mut s = String::from(XML_HEAD);
    s.push_str(&format!("<styleSheet xmlns=\"{NS_MAIN}\">\n"));
    s.push_str(
        "<fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/><family val=\"2\"/></font></fonts>\n",
    );
    s.push_str(&format!("<fills count=\"{}\">\n", fills.len() + 2));
    s.push_str("<fill><patternFill patternType=\"none\"/></fill>\n");
    s.push_str("<fill><patternFill patternType=\"gray125\"/></fill>\n");
    for f in &fills {
        s.push_str(&format!(
            "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"{}\"/><bgColor indexed=\"64\"/></patternFill></fill>\n",
            f.argb_hex()
        ));
    }
    s.push_str("</fills>\n");
    s.push_str(
        "<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>\n",
    );
    s.push_str(
        "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>\n",
    );
    s.push_str(&format!("<cellXfs count=\"{}\">\n", wb.styles.xfs().len()));
    for xf in wb.styles.xfs() {
        let fill_id = match xf.fill {
            Some(f) => fills.iter().position(|g| *g == f).unwrap_or(0) + 2,
            None => 0,
        };
        s.push_str(&format!(
            "<xf numFmtId=\"{}\" fontId=\"0\" fillId=\"{}\" borderId=\"0\" xfId=\"0\"",
            xf.num_fmt.id(),
            fill_id
        ));
        if xf.num_fmt.id() != 0 {
            s.push_str(" applyNumberFormat=\"1\"");
        }
        if xf.fill.is_some() {
            s.push_str(" applyFill=\"1\"");
        }
        if xf.wrap {
            s.push_str(" applyAlignment=\"1\"><alignment wrapText=\"1\"/></xf>\n");
        } else {
            s.push_str("/>\n");
        }
    }
    s.push_str("</cellXfs>\n");
    s.push_str(
        "<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>\n",
    );
    s.push_str("</styleSheet>\n");
    s
}

/// A number as the file holds it. Rust's shortest round-trip spelling,
/// which never uses an exponent, is one Excel reads; a value Excel cannot
/// hold (not finite) is written as its `#NUM!` error.
fn number_text(v: f64) -> Option<String> {
    v.is_finite().then(|| format!("{v}"))
}

fn sheet_xml(sheet: &Sheet, selected: bool) -> String {
    let mut s = String::from(XML_HEAD);
    s.push_str(&format!(
        "<worksheet xmlns=\"{NS_MAIN}\" xmlns:r=\"{NS_REL}\">\n"
    ));
    if let Some(((top, left), (bottom, right))) = sheet.extent() {
        s.push_str(&format!(
            "<dimension ref=\"{}:{}\"/>\n",
            cell_ref(top, left),
            cell_ref(bottom, right)
        ));
    }
    let active = cell_ref(sheet.active_cell.0, sheet.active_cell.1);
    s.push_str("<sheetViews><sheetView");
    if !sheet.gridlines {
        s.push_str(" showGridLines=\"0\"");
    }
    if selected {
        s.push_str(" tabSelected=\"1\"");
    }
    s.push_str(&format!(
        " workbookViewId=\"0\"><selection activeCell=\"{active}\" sqref=\"{active}\"/></sheetView></sheetViews>\n"
    ));
    s.push_str("<sheetFormatPr defaultRowHeight=\"15\"/>\n");
    if !sheet.columns.is_empty() {
        s.push_str("<cols>\n");
        for c in &sheet.columns {
            s.push_str(&format!("<col min=\"{}\" max=\"{}\"", c.index, c.index));
            if c.hidden {
                s.push_str(" width=\"0\" hidden=\"1\" customWidth=\"1\"");
            } else if let Some(w) = c.width {
                s.push_str(&format!(
                    " width=\"{}\" customWidth=\"1\"",
                    column_width_attr(w)
                ));
            }
            if let Some(st) = c.style {
                s.push_str(&format!(" style=\"{st}\""));
            }
            s.push_str("/>\n");
        }
        s.push_str("</cols>\n");
    }
    s.push_str("<sheetData>\n");
    let mut current_row: Option<u32> = None;
    for (&(row, col), cell) in &sheet.cells {
        if current_row != Some(row) {
            if current_row.is_some() {
                s.push_str("</row>\n");
            }
            s.push_str(&format!("<row r=\"{row}\">"));
            current_row = Some(row);
        }
        let r = cell_ref(row, col);
        let st = cell.style;
        match &cell.content {
            Content::Text(t) => s.push_str(&format!(
                "<c r=\"{r}\" s=\"{st}\" t=\"inlineStr\"><is><t xml:space=\"preserve\">{}</t></is></c>",
                xml::text(t)
            )),
            Content::Number(v) => match number_text(*v) {
                Some(n) => s.push_str(&format!("<c r=\"{r}\" s=\"{st}\"><v>{n}</v></c>")),
                None => s.push_str(&format!(
                    "<c r=\"{r}\" s=\"{st}\" t=\"e\"><v>#NUM!</v></c>"
                )),
            },
            Content::Bool(b) => s.push_str(&format!(
                "<c r=\"{r}\" s=\"{st}\" t=\"b\"><v>{}</v></c>",
                u8::from(*b)
            )),
            Content::Formula(f) => s.push_str(&format!(
                "<c r=\"{r}\" s=\"{st}\"><f>{}</f></c>",
                xml::text(f)
            )),
        }
    }
    if current_row.is_some() {
        s.push_str("</row>\n");
    }
    s.push_str("</sheetData>\n");
    s.push_str("</worksheet>\n");
    s
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::sheet::{Cell, Column, NumFmt, Rgb, Style};

    fn a_workbook() -> Workbook {
        let mut wb = Workbook::new();
        let lavender = wb.styles.id(Style {
            fill: Some(Rgb(247, 244, 252)),
            num_fmt: NumFmt::Text,
            wrap: false,
        });
        let mut sh = Sheet::new("Fra & zaro");
        sh.gridlines = false;
        sh.active_cell = (1, 2);
        sh.columns.push(Column {
            index: 1,
            width: None,
            hidden: true,
            style: None,
        });
        sh.columns.push(Column {
            index: 2,
            width: Some(72.0),
            hidden: false,
            style: Some(lavender),
        });
        sh.set(
            1,
            2,
            Cell {
                content: Content::Text("Put 5 into cell B2.".to_string()),
                style: lavender,
            },
        );
        sh.set(
            3,
            3,
            Cell {
                content: Content::Number(2.5),
                style: 0,
            },
        );
        sh.set(
            3,
            4,
            Cell {
                content: Content::Formula("B1*2".to_string()),
                style: 0,
            },
        );
        sh.set(
            4,
            2,
            Cell {
                content: Content::Bool(true),
                style: 0,
            },
        );
        wb.sheets.push(sh);
        wb
    }

    #[test]
    fn the_parts_in_their_fixed_order() {
        let names: Vec<String> = parts(&a_workbook()).into_iter().map(|(n, _)| n).collect();
        assert_eq!(
            names,
            vec![
                "[Content_Types].xml",
                "_rels/.rels",
                "docProps/app.xml",
                "docProps/core.xml",
                "xl/workbook.xml",
                "xl/_rels/workbook.xml.rels",
                "xl/styles.xml",
                "xl/worksheets/frazaro_1.xml",
            ]
        );
    }

    #[test]
    fn the_workbook_names_its_sheet_and_asks_the_host_to_compute() {
        let wb = a_workbook();
        let w = workbook_xml(&wb);
        assert!(w.contains("<sheet name=\"Fra &amp; zaro\" sheetId=\"1\" r:id=\"rIdFrazaro1\"/>"));
        assert!(w.contains("<calcPr fullCalcOnLoad=\"1\"/>"));
        let ct = content_types(&wb);
        assert_eq!(ct.matches("<Override").count(), 5);
        assert!(ct.contains("/xl/worksheets/frazaro_1.xml"));
        let rels = workbook_rels(&wb);
        assert!(rels.contains("Target=\"worksheets/frazaro_1.xml\""));
        assert!(rels.contains("Id=\"rIdFrazaroStyles\""));
    }

    #[test]
    fn the_styles_part_keeps_the_two_required_fills_first() {
        let st = styles_xml(&a_workbook());
        assert!(st.contains("<fills count=\"3\">\n<fill><patternFill patternType=\"none\"/></fill>\n<fill><patternFill patternType=\"gray125\"/></fill>\n<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFF7F4FC\"/>"));
        assert!(st.contains("<cellXfs count=\"2\">"));
        assert!(st.contains("<xf numFmtId=\"49\" fontId=\"0\" fillId=\"2\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\" applyFill=\"1\"/>"));
    }

    #[test]
    fn the_sheet_part_holds_each_kind_of_cell() {
        let wb = a_workbook();
        let sx = sheet_xml(&wb.sheets[0], true);
        assert!(sx.contains("<dimension ref=\"B1:D4\"/>"));
        assert!(sx.contains("<sheetView showGridLines=\"0\" tabSelected=\"1\" workbookViewId=\"0\"><selection activeCell=\"B1\" sqref=\"B1\"/>"));
        assert!(
            sx.contains("<col min=\"1\" max=\"1\" width=\"0\" hidden=\"1\" customWidth=\"1\"/>")
        );
        assert!(sx.contains(
            "<col min=\"2\" max=\"2\" width=\"72.7109375\" customWidth=\"1\" style=\"1\"/>"
        ));
        assert!(sx.contains("<row r=\"1\"><c r=\"B1\" s=\"1\" t=\"inlineStr\"><is><t xml:space=\"preserve\">Put 5 into cell B2.</t></is></c></row>"));
        assert!(sx.contains("<row r=\"3\"><c r=\"C3\" s=\"0\"><v>2.5</v></c><c r=\"D3\" s=\"0\"><f>B1*2</f></c></row>"));
        assert!(sx.contains("<row r=\"4\"><c r=\"B4\" s=\"0\" t=\"b\"><v>1</v></c></row>"));
        let unselected = sheet_xml(&wb.sheets[0], false);
        assert!(!unselected.contains("tabSelected"));
    }

    #[test]
    fn numbers_as_the_file_holds_them() {
        assert_eq!(number_text(5.0).as_deref(), Some("5"));
        assert_eq!(number_text(0.1).as_deref(), Some("0.1"));
        assert_eq!(number_text(-2.5).as_deref(), Some("-2.5"));
        assert_eq!(number_text(f64::NAN), None);
        assert_eq!(number_text(f64::INFINITY), None);
    }

    #[test]
    fn the_bytes_are_a_stored_archive_of_the_parts() {
        let wb = a_workbook();
        let bytes = workbook_bytes(&wb).expect("fits");
        let entries = zip::entries(&bytes).expect("a zip");
        let expected = parts(&wb);
        assert_eq!(entries.len(), expected.len());
        for (e, (name, data)) in entries.iter().zip(expected.iter()) {
            assert_eq!(&e.name, name);
            assert_eq!(zip::stored_data(&bytes, e).unwrap(), data.as_slice());
        }
        assert_eq!(workbook_bytes(&wb), Some(bytes));
    }
}
