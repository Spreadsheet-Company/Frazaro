//! The OOXML parts of a workbook (ISO/IEC 29500, SpreadsheetML), rendered
//! from the sheet model as text in a fixed order, then zipped stored
//! (PORT.7, slice 7a).
//!
//! New ground, said plainly: the reference writes through Excel's object
//! model and never spells a part. What is here is the smallest set of parts
//! Excel, Sheets, Calc and Numbers open: the content types, the package
//! relationships, two property parts that name Frazaro and carry no date and
//! no person, the workbook with its `calcPr` (`fullCalcOnLoad`: the host
//! computes every formula when it opens the file), the workbook's
//! relationships, one styles part, and one part per sheet. Cells hold inline
//! strings, so there is no shared-string table, and (slice 7d) adding a sheet
//! to an existing workbook edits no part the model's own cells depend on.
//! tools/build_examples.ps1 writes the same minimal shape in PowerShell for
//! the onboarding samples, and Excel opens those.
//!
//! Since `KERNEL.20`'s fifth slice (2026-10-10) a fresh build's formula cells
//! hold the values the build computed (`parts_valued`), as a host's save
//! holds the values it computed: a number, a text (`t="str"`), a truth value
//! (`t="b"`) or an error (`t="e"`), each spelled as Excel spells it, so that
//! a reader that does not calculate shows them. A formula the build did not
//! compute holds none. The host still computes every formula when it opens
//! the file.
//!
//! A filled formula is written as one shared formula, as Excel writes one,
//! unless it holds what a shared formula may not (`PORT.13`, 2026-10-10): a
//! reference to another sheet or another workbook, an array constant, a
//! union, an intersection, a range built by the range operator, a `#REF!`, a
//! table's column, a spill, a newer function. Excel writes such a fill one
//! formula a cell, and refuses a file that shares one (the owner's run of
//! `KERNEL.20`'s sixth slice: "Unable to get the Open property of the
//! Workbooks class" for a Life whose rule reads another sheet), so each of
//! its cells here holds its own formula, the master's with its references
//! moved. The model is unchanged: it still holds the fill as one formula, and
//! what a reader prints of the file is what it printed before.
//!
//! Since `KERNEL.47` (2026-10-17) a format's number format may be a code of
//! Excel's (`vla_lang::numfmt`): one that is a built-in Excel draws alike in
//! every locale, `0.00` for one, is written by its id as Excel writes it,
//! and any other is defined in the styles part's `<numFmts>`, ids from 164
//! as Excel numbers its own, so that a cell holding 1 under `0.00` draws
//! `1.00` in any host.
//!
//! Every name here is fixed and every order is fixed, so that the same model
//! gives the same bytes: the parts in the order [`parts`] lists them, a
//! sheet's rows and cells in row-then-column order (the model's map), the
//! fills, formats and codes in the order the style table holds them.

use std::collections::HashSet;

use super::xml;
use super::zip;
use super::{
    cell_ref, column_width_attr, number_text, shared_masters, Content, NumFmt, Sheet, Workbook,
};
use crate::calc::{Computed, Value, Values};
use crate::refers::{self, Kind};
use vla_lang::numfmt::Code;

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
pub(super) const REL_WORKSHEET: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
pub(super) const REL_STYLES: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
const CT_WORKBOOK: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
pub(super) const CT_WORKSHEET: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
const CT_STYLES: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
const CT_CORE: &str = "application/vnd.openxmlformats-package.core-properties+xml";
const CT_APP: &str = "application/vnd.openxmlformats-officedocument.extended-properties+xml";
const CT_RELS: &str = "application/vnd.openxmlformats-package.relationships+xml";
const CT_XML: &str = "application/xml";
pub(super) const CT_METADATA: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml";
pub(super) const REL_METADATA: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata";
/// The relationship id of the cell-metadata part, present only when a
/// dynamic-array formula is.
pub(super) const METADATA_RID: &str = "rIdFrazaroMetadata";

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
    parts_with(wb, None)
}

/// Every part, each formula cell holding the value `values` computed for it
/// (KERNEL.20): what a fresh build writes.
pub fn parts_valued(wb: &Workbook, values: &dyn Values) -> Vec<(String, Vec<u8>)> {
    parts_with(wb, Some(values))
}

fn parts_with(wb: &Workbook, values: Option<&dyn Values>) -> Vec<(String, Vec<u8>)> {
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
    if wb.has_dynamic_formulas() {
        list.push(("xl/metadata.xml".to_string(), metadata_xml().into_bytes()));
    }
    for (i, sheet) in wb.sheets.iter().enumerate() {
        let selected = i == wb.active_sheet;
        list.push((
            sheet_part(i + 1),
            sheet_text(sheet, selected, &Render::FRESH, values.map(|v| (v, i))).into_bytes(),
        ));
    }
    list
}

/// The workbook as a file's bytes; `None` when it would not fit the zip
/// format's sizes.
pub fn workbook_bytes(wb: &Workbook) -> Option<Vec<u8>> {
    zip::write_stored(&parts(wb))
}

/// The workbook as a file's bytes, each formula cell holding the value
/// `values` computed for it ([`parts_valued`]).
pub fn workbook_bytes_valued(wb: &Workbook, values: &dyn Values) -> Option<Vec<u8>> {
    zip::write_stored(&parts_valued(wb, values))
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
    if wb.has_dynamic_formulas() {
        s.push_str(&format!(
            "<Override PartName=\"/xl/metadata.xml\" ContentType=\"{CT_METADATA}\"/>\n"
        ));
    }
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
    if !wb.defined_names.is_empty() {
        s.push_str("<definedNames>\n");
        for (name, formula) in &wb.defined_names {
            s.push_str(&format!(
                "<definedName name=\"{}\">{}</definedName>\n",
                xml::attr(name),
                xml::text(formula)
            ));
        }
        s.push_str("</definedNames>\n");
    }
    // The host computes every formula when it opens the file, whatever
    // value a build folded into it (KERNEL.20).
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
    if wb.has_dynamic_formulas() {
        s.push_str(&format!(
            "<Relationship Id=\"{METADATA_RID}\" Type=\"{REL_METADATA}\" Target=\"metadata.xml\"/>\n"
        ));
    }
    s.push_str("</Relationships>\n");
    s
}

/// The id Excel gives the first number format a workbook defines; the ids
/// below it are its built-in formats.
pub const FIRST_CODE_ID: u32 = 164;

/// The ids a styles part gives the codes of a style table that are no
/// built-in Excel draws alike in every locale (KERNEL.47): each such code
/// once, in the order the table first uses it, numbered from `first`.
pub fn code_ids(styles: &super::Styles, first: u32) -> Vec<(Code, u32)> {
    let mut ids: Vec<(Code, u32)> = Vec::new();
    for xf in styles.xfs() {
        if let NumFmt::Code(code) = &xf.num_fmt {
            if code.builtin().is_none() && !ids.iter().any(|(c, _)| c == code) {
                let id = first + ids.len() as u32;
                ids.push((code.clone(), id));
            }
        }
    }
    ids
}

/// A format's `numFmtId`: its built-in's, or the id `codes` gives its code.
fn num_fmt_id(num_fmt: &NumFmt, codes: &[(Code, u32)]) -> u32 {
    if let Some(id) = num_fmt.builtin() {
        return id;
    }
    match num_fmt {
        NumFmt::Code(code) => codes
            .iter()
            .find(|(c, _)| c == code)
            .map_or(0, |(_, id)| *id),
        _ => 0,
    }
}

/// Codes and their ids as `<numFmt>` elements, for a `<numFmts>` list. A
/// code is the attribute's text as written, escaped as every `ST_Xstring`
/// is (`xml::attr`).
pub fn numfmts_fragment(codes: &[(Code, u32)]) -> String {
    let mut s = String::new();
    for (code, id) in codes {
        s.push_str(&format!(
            "<numFmt numFmtId=\"{id}\" formatCode=\"{}\"/>",
            xml::attr(code.as_str())
        ));
    }
    s
}

/// The styles part: the number formats the style table defines (KERNEL.47),
/// one font, the two fills the format requires first (none and gray125)
/// then each fill the model uses, one empty border, and one cell format per
/// entry of the style table, in its order.
fn styles_xml(wb: &Workbook) -> String {
    let fills = wb.styles.fills();
    let codes = code_ids(&wb.styles, FIRST_CODE_ID);
    let mut s = String::from(XML_HEAD);
    s.push_str(&format!("<styleSheet xmlns=\"{NS_MAIN}\">\n"));
    if !codes.is_empty() {
        s.push_str(&format!(
            "<numFmts count=\"{}\">{}</numFmts>\n",
            codes.len(),
            numfmts_fragment(&codes)
        ));
    }
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
        let num_fmt_id = num_fmt_id(&xf.num_fmt, &codes);
        s.push_str(&format!(
            "<xf numFmtId=\"{num_fmt_id}\" fontId=\"0\" fillId=\"{fill_id}\" borderId=\"0\" xfId=\"0\""
        ));
        if num_fmt_id != 0 {
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

/// This build's solid fills as `<fill>` elements, in the style table's
/// order, for appending to a workbook's `<fills>`.
pub fn fills_fragment(styles: &super::Styles) -> String {
    let mut s = String::new();
    for f in styles.fills() {
        s.push_str(&format!(
            "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"{}\"/><bgColor indexed=\"64\"/></patternFill></fill>",
            f.argb_hex()
        ));
    }
    s
}

/// This build's cell formats after the default, as `<xf>` elements whose
/// fills start at `fill_base` (the number of fills the workbook has before
/// this build's) and whose codes take the ids `codes` gives them in that
/// workbook (KERNEL.47), for appending to a workbook's `<cellXfs>`.
pub fn cellxfs_fragment(styles: &super::Styles, fill_base: u32, codes: &[(Code, u32)]) -> String {
    let fills = styles.fills();
    let mut s = String::new();
    for xf in styles.xfs().iter().skip(1) {
        let fill_id = match xf.fill {
            Some(f) => fills.iter().position(|g| *g == f).unwrap_or(0) as u32 + fill_base,
            None => 0,
        };
        let num_fmt_id = num_fmt_id(&xf.num_fmt, codes);
        s.push_str(&format!(
            "<xf numFmtId=\"{num_fmt_id}\" fontId=\"0\" fillId=\"{fill_id}\" borderId=\"0\" xfId=\"0\""
        ));
        if num_fmt_id != 0 {
            s.push_str(" applyNumberFormat=\"1\"");
        }
        if xf.fill.is_some() {
            s.push_str(" applyFill=\"1\"");
        }
        if xf.wrap {
            s.push_str(" applyAlignment=\"1\"><alignment wrapText=\"1\"/></xf>");
        } else {
            s.push_str("/>");
        }
    }
    s
}

/// The cell metadata that `cm="1"` on a cell points at, as Excel writes it:
/// one metadata type, XLDAPR, carrying the dynamic-array properties, so
/// that the host reads the cell's formula as a dynamic-array formula.
pub fn metadata_xml() -> String {
    format!(
        "{XML_HEAD}<metadata xmlns=\"{NS_MAIN}\" xmlns:xda=\"http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray\">\n\
         <metadataTypes count=\"1\"><metadataType name=\"XLDAPR\" minSupportedVersion=\"120000\" copy=\"1\" pasteAll=\"1\" pasteValues=\"1\" merge=\"1\" splitFirst=\"1\" rowColShift=\"1\" clearFormats=\"1\" clearComments=\"1\" assign=\"1\" coerce=\"1\" cellMeta=\"1\"/></metadataTypes>\n\
         <futureMetadata name=\"XLDAPR\" count=\"1\"><bk><extLst><ext uri=\"{{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}}\"><xda:dynamicArrayProperties fDynamic=\"1\" fCollapsed=\"0\"/></ext></extLst></bk></futureMetadata>\n\
         <cellMetadata count=\"1\"><bk><rc t=\"1\" v=\"0\"/></bk></cellMetadata>\n\
         </metadata>\n"
    )
}

/// How a sheet's cells name the workbook's shared tables: the number of
/// cell formats the workbook already has before this build's (0 for a
/// fresh workbook, where this build's table is the whole table), and the
/// index of the dynamic-array cell metadata (1 in a fresh workbook).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Render {
    pub style_base: u32,
    pub cm: u32,
}

impl Render {
    pub const FRESH: Render = Render {
        style_base: 0,
        cm: 1,
    };

    /// A cell's `s` attribute: this build's format 0 is the workbook's
    /// Normal; its others follow the workbook's own.
    pub fn style(&self, st: u32) -> u32 {
        if st == 0 || self.style_base == 0 {
            st
        } else {
            self.style_base + st - 1
        }
    }
}

/// One sheet's part. `selected` marks the tab the workbook opens on, never
/// set when adding to a workbook that has its own selected tab. Its formula
/// cells hold no value, as a build added to a workbook writes them: their
/// formulas may read the workbook's own cells, which the build never read.
pub fn sheet_xml(sheet: &Sheet, selected: bool, render: &Render) -> String {
    sheet_text(sheet, selected, render, None)
}

/// One sheet's part, the `index`-th of the workbook `values` computed, each
/// formula cell holding the value computed for it (KERNEL.20).
pub fn sheet_xml_valued(
    sheet: &Sheet,
    index: usize,
    selected: bool,
    render: &Render,
    values: &dyn Values,
) -> String {
    sheet_text(sheet, selected, render, Some((values, index)))
}

/// Whether Excel can hold a filled formula, its first cell's text `text`, as
/// one shared formula (`PORT.13`). A shared formula holds none of what
/// [MS-XLS] `SharedParsedFormula`, the formula of a `ShrFmla` record
/// (2.4.260), forbids: a reference to another
/// sheet or a span of sheets (`PtgRef3d`, `PtgArea3d`), to another workbook
/// or its names (`PtgNameX`), a broken one (`PtgRefErr`), an array constant
/// (`PtgArray`), a union, an intersection or a range built by the range
/// operator (`PtgUnion`, `PtgIsect`, `PtgRange`). Read on the safe side, since
/// a formula a cell holds alone is never refused and a shared one may be: a
/// table's column or a spill, which the specification predates; a newer
/// function, which Excel's binary formats hold as an external name; any
/// text the evaluator's parser does not read, which is how a union and an
/// intersection arrive; and a range operator over anything but two corners,
/// which the parser reads as a range between expressions (KERNEL.8, slice f).
fn shares(text: &str) -> bool {
    if text.contains("_xlfn.") || has_array_constant(text) {
        return false;
    }
    let refs = refers::scan(text);
    if refs.iter().any(|r| {
        !r.sheet_name.is_empty()
            || !r.book.is_empty()
            || matches!(
                r.kind,
                Kind::ThreeD | Kind::External | Kind::Broken | Kind::Structured | Kind::Spill
            )
    }) {
        return false;
    }
    crate::calc::parse(text).is_ok_and(|e| !e.has_range())
}

/// Whether a formula's text holds an array constant: a `{` outside its
/// quoted texts.
fn has_array_constant(text: &str) -> bool {
    let mut quoted = false;
    for ch in text.chars() {
        match ch {
            '"' => quoted = !quoted,
            '{' if !quoted => return true,
            _ => {}
        }
    }
    false
}

/// A formula cell's value as a host's save holds one: the `t` attribute its
/// kind takes and its `<v>` element, a number in the writer's spelling (zero
/// as `0`, since no cell holds a negative zero) and one no cell can hold as
/// `#NUM!`, as a value cell spells it; nothing at all for a formula with no
/// value, which the host computes when it opens the file, and for one whose
/// value is an error with no cell error value of its own, `#CALC!`, which
/// the file format holds only as a rich value: written bare, Excel opened
/// the arrays' fixture by repairing every such cell (KERNEL.8, slice f).
fn cached(value: Option<&Value>) -> (&'static str, String) {
    match value {
        None => ("", String::new()),
        Some(Value::Error(e)) if !e.is_literal() => ("", String::new()),
        Some(Value::Number(n)) if *n == 0.0 => ("", "<v>0</v>".to_string()),
        Some(Value::Number(n)) => match number_text(*n) {
            Some(t) => ("", format!("<v>{t}</v>")),
            None => (" t=\"e\"", "<v>#NUM!</v>".to_string()),
        },
        Some(Value::Text(t)) => (" t=\"str\"", format!("<v>{}</v>", xml::text(t))),
        Some(Value::Bool(b)) => (" t=\"b\"", format!("<v>{}</v>", u8::from(*b))),
        Some(Value::Error(e)) => (" t=\"e\"", format!("<v>{}</v>", e.text())),
        Some(Value::Empty) => ("", "<v>0</v>".to_string()),
    }
}

fn sheet_text(
    sheet: &Sheet,
    selected: bool,
    render: &Render,
    values: Option<(&dyn Values, usize)>,
) -> String {
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
                s.push_str(&format!(" style=\"{}\"", render.style(st)));
            }
            s.push_str("/>\n");
        }
        s.push_str("</cols>\n");
    }
    s.push_str("<sheetData>\n");
    // The fills Excel cannot hold as one shared formula, by index, each of
    // whose cells is written with its own formula (PORT.13).
    let masters = shared_masters(sheet);
    let unshared: HashSet<u32> = masters
        .iter()
        .filter(|(_, (text, _, _))| !shares(text))
        .map(|(si, _)| *si)
        .collect();
    // The fills still shared numbered from 0 in their order, as Excel numbers
    // its own, so that a fill written cell by cell leaves no gap: the model's
    // numbers whenever none is.
    let mut kept: Vec<u32> = masters
        .keys()
        .filter(|si| !unshared.contains(*si))
        .copied()
        .collect();
    kept.sort_unstable();
    let file_si: std::collections::HashMap<u32, u32> =
        kept.iter().zip(0u32..).map(|(si, k)| (*si, k)).collect();
    let numbered = |si: &u32| file_si.get(si).copied().unwrap_or(*si);
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
        let st = render.style(cell.style);
        let value = values.and_then(|(v, i)| match v.computed((i, row, col)) {
            Some(Computed::Value(x)) => Some(x),
            _ => None,
        });
        let (kind, held) = cached(value);
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
            Content::Error(e) => s.push_str(&format!(
                "<c r=\"{r}\" s=\"{st}\" t=\"e\"><v>{}</v></c>",
                xml::text(e)
            )),
            Content::Formula(f) => s.push_str(&format!(
                "<c r=\"{r}\" s=\"{st}\"{kind}><f>{}</f>{held}</c>",
                xml::text(f)
            )),
            Content::SharedMaster { text, si, .. } if unshared.contains(si) => s.push_str(
                &format!("<c r=\"{r}\" s=\"{st}\"{kind}><f>{}</f>{held}</c>", xml::text(text)),
            ),
            Content::SharedChild { si } if unshared.contains(si) => {
                // The master's text moved by the cell's distance from it, its
                // newer functions' prefixes kept, as the file spells them.
                let (text, m_row, m_col) = masters[si];
                let moved = refers::shift_a1_references(
                    text,
                    i64::from(row) - i64::from(m_row),
                    i64::from(col) - i64::from(m_col),
                );
                s.push_str(&format!(
                    "<c r=\"{r}\" s=\"{st}\"{kind}><f>{}</f>{held}</c>",
                    xml::text(&moved)
                ));
            }
            Content::SharedMaster { text, range, si } => s.push_str(&format!(
                "<c r=\"{r}\" s=\"{st}\"{kind}><f t=\"shared\" ref=\"{range}\" si=\"{}\">{}</f>{held}</c>",
                numbered(si),
                xml::text(text)
            )),
            Content::SharedChild { si } => s.push_str(&format!(
                "<c r=\"{r}\" s=\"{st}\"{kind}><f t=\"shared\" si=\"{}\"/>{held}</c>",
                numbered(si)
            )),
            Content::DynamicFormula(f) => s.push_str(&format!(
                "<c r=\"{r}\" s=\"{st}\"{kind} cm=\"{}\"><f t=\"array\" ref=\"{r}\">{}</f>{held}</c>",
                render.cm,
                xml::text(f)
            )),
            // A legacy array formula over its own cell, Ctrl+Shift+Enter's,
            // with no dynamic-array metadata (KERNEL.8, slice f).
            Content::ArrayFormula(f) => s.push_str(&format!(
                "<c r=\"{r}\" s=\"{st}\"{kind}><f t=\"array\" ref=\"{r}\">{}</f>{held}</c>",
                xml::text(f)
            )),
            // A kind of content this writer does not know writes no cell.
            _ => {}
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
        let lavender = wb
            .styles
            .id(Style::new(Some(Rgb(247, 244, 252)), NumFmt::Text, false));
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
        sh.set_formula(crate::sheet::parse_a1_range("e2:e3").unwrap(), "D3+1", 0);
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
    fn a_code_is_written_as_its_built_in_or_defined_from_164() {
        // KERNEL.47: a built-in's code by its id, as Excel writes one; any
        // other once, from 164, in the order the style table first uses it.
        let code = |text: &str| NumFmt::Code(Code::parse(text).unwrap());
        let mut wb = a_workbook();
        let two = wb.styles.id(Style::new(None, code("0.00"), false));
        let glyph = wb
            .styles
            .id(Style::new(None, code("▲0.00;▼0.00;=0.00"), false));
        let kg = wb.styles.id(Style::new(
            Some(Rgb(1, 2, 3)),
            code("0.00\" kg & <g>\""),
            false,
        ));
        let again = wb
            .styles
            .id(Style::new(None, code("▲0.00;▼0.00;=0.00"), true));
        assert_eq!((two, glyph, kg, again), (2, 3, 4, 5));
        let st = styles_xml(&wb);
        assert!(st.contains("<numFmts count=\"2\"><numFmt numFmtId=\"164\" formatCode=\"▲0.00;▼0.00;=0.00\"/><numFmt numFmtId=\"165\" formatCode=\"0.00&quot; kg &amp; &lt;g&gt;&quot;\"/></numFmts>\n<fonts"), "{st}");
        assert!(st.contains("<xf numFmtId=\"2\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>"), "{st}");
        assert!(st.contains("<xf numFmtId=\"164\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>"), "{st}");
        assert!(st.contains("<xf numFmtId=\"165\" fontId=\"0\" fillId=\"3\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\" applyFill=\"1\"/>"), "{st}");
        assert!(st.contains("<xf numFmtId=\"164\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\" applyAlignment=\"1\"><alignment wrapText=\"1\"/></xf>"), "{st}");
        // A workbook whose formats hold no code writes no list, as before.
        assert!(!styles_xml(&a_workbook()).contains("numFmts"));
    }

    /// The codes and values of the workbook the owner opens in Excel and in
    /// LibreOffice Calc at the hand test (KERNEL.47).
    const DRAWN: &[(&str, f64)] = &[
        ("0.00", 1.0),
        ("0.00", 1.005),
        ("0.00", 2.675),
        ("0.00", 0.125),
        ("0.00", -0.001),
        ("#,##0.00", 1234.5),
        ("#,##0", 1234567.0),
        ("0%", 0.5),
        ("0.00%", 0.12345),
        ("#,##0,\"K\"", 2500.0),
        ("\\$#,##0.00", 1234.5),
        ("#,##0 ;(#,##0)", -1234.0),
        ("#,##0.00;(#,##0.00)", -0.001),
        ("▲0.00;▼0.00;=0.00", 1.5),
        ("▲0.00;▼0.00;=0.00", -1.5),
        ("▲0.00;▼0.00;=0.00", 0.0),
        ("000-00-0000", 123456789.0),
        ("0.0,,", 12200000.0),
        ("#.##", 0.5),
        ("0.#", 1.0),
        ("0", 123456789012345678.0),
        ("0.00\" kg\"", 2.0),
    ];

    #[test]
    fn a_workbook_of_codes_holds_what_a_host_draws_beside_what_the_language_does() {
        // Each row the code in A, a value under it in B, the text the
        // language draws in C, Excel's TEXT of B by A in D, and in E whether
        // C and D are alike: the host computes D and E when it opens the
        // file. Written where FRAZARO_WRITE_FORMATS names a path, for the
        // owner's open of it in Excel and in Calc.
        let mut wb = Workbook::new();
        let mut sh = Sheet::new("Formats");
        for (i, (text, x)) in DRAWN.iter().enumerate() {
            let row = i as u32 + 1;
            let code = Code::parse(text).unwrap();
            let drawn = code.text(&Value::Number(*x)).unwrap();
            let style = wb.styles.id(Style::new(None, NumFmt::Code(code), false));
            let cells = [
                (1, Content::Text(text.to_string()), 0),
                (2, Content::Number(*x), style),
                (3, Content::Text(drawn), 0),
                (4, Content::Formula(format!("TEXT(B{row},A{row})")), 0),
                (5, Content::Formula(format!("EXACT(C{row},D{row})")), 0),
            ];
            for (col, content, style) in cells {
                sh.set(row, col, Cell { content, style });
            }
        }
        wb.sheets.push(sh);
        let st = styles_xml(&wb);
        // The built-ins among them by their ids, the rest defined.
        assert!(st.contains("<numFmts count=\"8\">"), "{st}");
        assert!(st.contains("<xf numFmtId=\"2\" "), "{st}");
        assert!(st.contains("<xf numFmtId=\"37\" "), "{st}");
        let bytes = workbook_bytes(&wb).expect("the workbook's bytes");
        if let Ok(path) = std::env::var("FRAZARO_WRITE_FORMATS") {
            std::fs::write(&path, &bytes).expect("the workbook written");
        }
    }

    #[test]
    fn the_sheet_part_holds_each_kind_of_cell() {
        let wb = a_workbook();
        let sx = sheet_xml(&wb.sheets[0], true, &Render::FRESH);
        assert!(sx.contains("<dimension ref=\"B1:E4\"/>"));
        assert!(sx.contains("<sheetView showGridLines=\"0\" tabSelected=\"1\" workbookViewId=\"0\"><selection activeCell=\"B1\" sqref=\"B1\"/>"));
        assert!(
            sx.contains("<col min=\"1\" max=\"1\" width=\"0\" hidden=\"1\" customWidth=\"1\"/>")
        );
        assert!(sx.contains(
            "<col min=\"2\" max=\"2\" width=\"72.7109375\" customWidth=\"1\" style=\"1\"/>"
        ));
        assert!(sx.contains("<row r=\"1\"><c r=\"B1\" s=\"1\" t=\"inlineStr\"><is><t xml:space=\"preserve\">Put 5 into cell B2.</t></is></c></row>"));
        assert!(sx.contains("<row r=\"2\"><c r=\"E2\" s=\"0\"><f t=\"shared\" ref=\"E2:E3\" si=\"0\">D3+1</f></c></row>"));
        assert!(sx.contains("<row r=\"3\"><c r=\"C3\" s=\"0\"><v>2.5</v></c><c r=\"D3\" s=\"0\"><f>B1*2</f></c><c r=\"E3\" s=\"0\"><f t=\"shared\" si=\"0\"/></c></row>"));
        assert!(sx.contains("<row r=\"4\"><c r=\"B4\" s=\"0\" t=\"b\"><v>1</v></c></row>"));
        assert!(sx.contains("<dimension ref=\"B1:E4\"/>"));
        let unselected = sheet_xml(&wb.sheets[0], false, &Render::FRESH);
        assert!(!unselected.contains("tabSelected"));
        // Into a workbook with 7 cell formats and 3 fills, whose dynamic-array
        // metadata is its second cell-metadata block: this build's format 1
        // is the workbook's 7, the column styled 1 says 7, and cm is 2.
        let hosted = Render {
            style_base: 7,
            cm: 2,
        };
        assert_eq!(hosted.style(0), 0);
        assert_eq!(hosted.style(1), 7);
        assert_eq!(hosted.style(2), 8);
        let hx = sheet_xml(&wb.sheets[0], false, &hosted);
        assert!(hx.contains(
            "<col min=\"2\" max=\"2\" width=\"72.7109375\" customWidth=\"1\" style=\"7\"/>"
        ));
        assert!(hx.contains("<c r=\"B1\" s=\"7\" t=\"inlineStr\">"));
        assert!(hx.contains("<c r=\"C3\" s=\"0\"><v>2.5</v></c>"));
        assert_eq!(
            fills_fragment(&wb.styles),
            "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFF7F4FC\"/><bgColor indexed=\"64\"/></patternFill></fill>"
        );
        assert_eq!(
            cellxfs_fragment(&wb.styles, 3, &[]),
            "<xf numFmtId=\"49\" fontId=\"0\" fillId=\"3\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\" applyFill=\"1\"/>"
        );
    }

    #[test]
    fn a_dynamic_formula_brings_the_metadata_part_and_a_name_its_element() {
        let plain = a_workbook();
        assert!(!parts(&plain).iter().any(|(n, _)| n == "xl/metadata.xml"));
        assert!(!content_types(&plain).contains("metadata"));
        assert!(!workbook_rels(&plain).contains("metadata"));
        assert!(!workbook_xml(&plain).contains("definedNames"));

        let mut wb = a_workbook();
        wb.sheets[0].set_formula_dynamic(
            crate::sheet::parse_a1_range("f2").unwrap(),
            "_xlfn.IFS(B2>3,\"big\",TRUE,\"small\")",
            0,
        );
        wb.defined_names.push((
            "Frazaro.Build".to_string(),
            "\"Frazaro 0.7.1; a & b\"".to_string(),
        ));
        let names: Vec<String> = parts(&wb).into_iter().map(|(n, _)| n).collect();
        assert_eq!(names[6], "xl/styles.xml");
        assert_eq!(names[7], "xl/metadata.xml");
        assert_eq!(names[8], "xl/worksheets/frazaro_1.xml");
        assert!(content_types(&wb).contains(
            "<Override PartName=\"/xl/metadata.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml\"/>"
        ));
        assert!(workbook_rels(&wb).contains(
            "<Relationship Id=\"rIdFrazaroMetadata\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata\" Target=\"metadata.xml\"/>"
        ));
        let w = workbook_xml(&wb);
        assert!(w.contains(
            "</sheets>\n<definedNames>\n<definedName name=\"Frazaro.Build\">\"Frazaro 0.7.1; a &amp; b\"</definedName>\n</definedNames>\n<calcPr"
        ));
        let sx = sheet_xml(&wb.sheets[0], true, &Render::FRESH);
        assert!(sx.contains(
            "<c r=\"F2\" s=\"0\" cm=\"1\"><f t=\"array\" ref=\"F2\">_xlfn.IFS(B2&gt;3,\"big\",TRUE,\"small\")</f></c>"
        ));
        let hosted = sheet_xml(
            &wb.sheets[0],
            false,
            &Render {
                style_base: 4,
                cm: 3,
            },
        );
        assert!(hosted.contains("<c r=\"F2\" s=\"0\" cm=\"3\"><f t=\"array\" ref=\"F2\">"));
        let meta = metadata_xml();
        assert!(meta.contains("<metadataType name=\"XLDAPR\""));
        assert!(meta.contains("<xda:dynamicArrayProperties fDynamic=\"1\" fCollapsed=\"0\"/>"));
        assert!(meta
            .contains("<cellMetadata count=\"1\"><bk><rc t=\"1\" v=\"0\"/></bk></cellMetadata>"));
    }

    #[test]
    fn a_legacy_array_formula_is_written_over_its_cell_and_brings_no_metadata() {
        // Ctrl+Shift+Enter's formula (KERNEL.8, slice f): `t="array"` over its
        // own cell and no `cm`, which a reader tells from a dynamic one, and
        // no metadata part, which a dynamic formula alone brings.
        let mut wb = a_workbook();
        wb.sheets[0].set(
            2,
            6,
            Cell {
                content: Content::ArrayFormula("SUM(C3:C4*2)".to_string()),
                style: 0,
            },
        );
        assert!(!parts(&wb).iter().any(|(n, _)| n == "xl/metadata.xml"));
        let sx = sheet_xml(&wb.sheets[0], true, &Render::FRESH);
        assert!(sx.contains("<c r=\"F2\" s=\"0\"><f t=\"array\" ref=\"F2\">SUM(C3:C4*2)</f></c>"));
    }

    /// Values held as a recalculation would hold them, by cell.
    #[derive(Debug)]
    struct Held(std::collections::HashMap<crate::calc::CellId, Computed>);

    impl Values for Held {
        fn computed(&self, id: crate::calc::CellId) -> Option<&Computed> {
            self.0.get(&id)
        }
    }

    #[test]
    fn a_formula_cell_holds_the_value_the_build_computed_as_a_host_saves_one() {
        use crate::calc::{ErrorKind, Reason};
        let mut wb = a_workbook();
        let sh = &mut wb.sheets[0];
        sh.set_formula_dynamic(crate::sheet::parse_a1_range("f2").unwrap(), "1/0", 0);
        let column_g = [
            "VLOOKUP(1,A1:B2,2)",
            "G1+1",
            "0.1+0.2",
            "-0",
            "1E308*10",
            "B2",
            "SEQUENCE(0)",
        ];
        for (row, f) in (1..).zip(column_g) {
            sh.set(
                row,
                7,
                Cell {
                    content: Content::Formula(f.to_string()),
                    style: 0,
                },
            );
        }
        let held = Held(
            [
                ((0, 3, 4), Computed::Value(Value::Number(5.0))),
                (
                    (0, 2, 5),
                    Computed::Value(Value::Text("a & <b> _x0041_".into())),
                ),
                ((0, 3, 5), Computed::Value(Value::Bool(true))),
                ((0, 2, 6), Computed::Value(Value::Error(ErrorKind::Div0))),
                (
                    (0, 1, 7),
                    Computed::NotComputed(Reason::Function("VLOOKUP".into())),
                ),
                ((0, 3, 7), Computed::Value(Value::Number(0.1 + 0.2))),
                ((0, 4, 7), Computed::Value(Value::Number(-0.0))),
                ((0, 5, 7), Computed::Value(Value::Number(f64::INFINITY))),
                ((0, 6, 7), Computed::Value(Value::Empty)),
                ((0, 7, 7), Computed::Value(Value::Error(ErrorKind::Calc))),
                // A value cell is never asked for its value.
                ((0, 3, 3), Computed::Value(Value::Number(99.0))),
            ]
            .into_iter()
            .collect(),
        );
        let sx = sheet_xml_valued(&wb.sheets[0], 0, true, &Render::FRESH, &held);
        assert!(sx.contains("<c r=\"D3\" s=\"0\"><f>B1*2</f><v>5</v></c>"));
        assert!(sx.contains("<c r=\"E2\" s=\"0\" t=\"str\"><f t=\"shared\" ref=\"E2:E3\" si=\"0\">D3+1</f><v>a &amp; &lt;b&gt; _x005F_x0041_</v></c>"));
        assert!(sx.contains("<c r=\"E3\" s=\"0\" t=\"b\"><f t=\"shared\" si=\"0\"/><v>1</v></c>"));
        assert!(sx.contains(
            "<c r=\"F2\" s=\"0\" t=\"e\" cm=\"1\"><f t=\"array\" ref=\"F2\">1/0</f><v>#DIV/0!</v></c>"
        ));
        // Not computed, and not asked of the values at all: no value.
        assert!(sx.contains("<c r=\"G1\" s=\"0\"><f>VLOOKUP(1,A1:B2,2)</f></c>"));
        assert!(sx.contains("<c r=\"G2\" s=\"0\"><f>G1+1</f></c>"));
        assert!(sx.contains("<c r=\"G3\" s=\"0\"><f>0.1+0.2</f><v>0.30000000000000004</v></c>"));
        assert!(sx.contains("<c r=\"G4\" s=\"0\"><f>-0</f><v>0</v></c>"));
        assert!(sx.contains("<c r=\"G5\" s=\"0\" t=\"e\"><f>1E308*10</f><v>#NUM!</v></c>"));
        assert!(sx.contains("<c r=\"G6\" s=\"0\"><f>B2</f><v>0</v></c>"));
        // `#CALC!` has no cell error value: Excel repaired a bare one, so the
        // formula holds no value and the host computes it (KERNEL.8, slice f).
        assert!(sx.contains("<c r=\"G7\" s=\"0\"><f>SEQUENCE(0)</f></c>"));
        assert!(sx.contains("<c r=\"C3\" s=\"0\"><v>2.5</v></c>"));
        // The same sheet unvalued, as a build added to a workbook writes it,
        // and valued from nothing, is the part as it always was.
        let bare = sheet_xml(&wb.sheets[0], true, &Render::FRESH);
        assert!(!bare.contains("</f><v>"));
        let nothing = Held(std::collections::HashMap::new());
        assert_eq!(
            sheet_xml_valued(&wb.sheets[0], 0, true, &Render::FRESH, &nothing),
            bare
        );
        assert_eq!(parts_valued(&wb, &nothing), parts(&wb));
        // Another sheet's values are its own.
        let other = sheet_xml_valued(&wb.sheets[0], 1, true, &Render::FRESH, &held);
        assert_eq!(other, bare);
    }

    #[test]
    fn a_fill_excel_cannot_share_is_written_one_formula_a_cell() {
        // PORT.13: what a shared formula may not hold, each alone...
        for text in [
            "Data!A1*2",
            "Output!A1+B1",
            "'My Data'!A1",
            "Jan:Dec!A1",
            "[1]Sheet1!A1",
            "SUM({1,2})+A1",
            "#REF!+1",
            "SUM((A1,B1))",
            "A1:B2 B1:C2",
            "SUM(A1:INDEX(B1:B9,2))",
            "Table1[Amount]",
            "A1#",
            "_xlfn.XOR(A1>1,B1>1)",
        ] {
            assert!(!shares(text), "{text} must be written one formula a cell");
        }
        // ...and what it may, as Excel writes it.
        for text in [
            "B1*2",
            "$A$1+B1",
            "SUM(A1:A3)",
            "Rate*B1",
            "\"{x}\"&B1",
            "IF(B1>1,\"big\",B1)",
            "SUM(A:A)",
            "ROW()*COLUMN()",
        ] {
            assert!(shares(text), "{text} shares");
        }
        // Rendered: a fill reading another sheet one formula a cell, each
        // the first cell's text with its references moved, the value folded
        // into each as before, and no shared index for it; the fill on its
        // own sheet shared as before, its index the first.
        let mut wb = Workbook::new();
        let mut sh = Sheet::new("Output");
        sh.set_formula(
            crate::sheet::parse_a1_range("b1:c2").unwrap(),
            "Data!A1*2+$D$1",
            0,
        );
        sh.set_formula(crate::sheet::parse_a1_range("e1:e2").unwrap(), "B1+1", 0);
        sh.set_formula(
            crate::sheet::parse_a1_range("f1:f2").unwrap(),
            "_xlfn.XOR(A1>1,B1>1)",
            0,
        );
        wb.sheets.push(sh);
        let held = Held(
            [
                ((0, 2, 3), Computed::Value(Value::Number(8.0))),
                ((0, 1, 5), Computed::Value(Value::Number(3.0))),
            ]
            .into_iter()
            .collect(),
        );
        let sx = sheet_xml_valued(&wb.sheets[0], 0, true, &Render::FRESH, &held);
        for (cell, formula) in [
            ("B1", "Data!A1*2+$D$1"),
            ("C1", "Data!B1*2+$D$1"),
            ("B2", "Data!A2*2+$D$1"),
        ] {
            assert!(
                sx.contains(&format!("<c r=\"{cell}\" s=\"0\"><f>{formula}</f></c>")),
                "{cell} in {sx}"
            );
        }
        assert!(sx.contains("<c r=\"C2\" s=\"0\"><f>Data!B2*2+$D$1</f><v>8</v></c>"));
        assert!(sx.contains(
            "<c r=\"E1\" s=\"0\"><f t=\"shared\" ref=\"E1:E2\" si=\"0\">B1+1</f><v>3</v></c>"
        ));
        assert!(sx.contains("<c r=\"E2\" s=\"0\"><f t=\"shared\" si=\"0\"/></c>"));
        assert!(!sx.contains("si=\"1\""), "{sx}");
        assert!(!sx.contains("si=\"2\""), "{sx}");
        // A newer function's cells keep its prefix, as the file spells it.
        assert!(
            sx.contains("<c r=\"F2\" s=\"0\"><f>_xlfn.XOR(A2&gt;1,B2&gt;1)</f></c>"),
            "{sx}"
        );
        // What a reader prints of the file is each cell's formula, as it
        // printed the shared one's.
        let bytes = workbook_bytes(&wb).unwrap();
        let rows = crate::api::reflect_relations(&bytes, "port13.xlsx").unwrap();
        assert!(
            rows.contains("(formula \"Output\" \"C2\" \"=Data!B2*2+$D$1\")"),
            "{rows}"
        );
        assert!(
            rows.contains("(formula \"Output\" \"E2\" \"=B2+1\")"),
            "{rows}"
        );
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
