//! Adding this build's sheets to a workbook someone else made (PORT.7,
//! slice 7d): `frazaro build --into model.xlsx`.
//!
//! The promise is that the model is untouched: every part of it is copied
//! as the compressed bytes it already is, byte for byte, and only the parts
//! that must know about the new sheets are edited, by text, at the one
//! place each needs: the workbook (its sheet list, the defined names, the
//! calcPr that asks the host to compute on open), the workbook's
//! relationships, the content types, and the styles, whose fill and format
//! lists grow at their ends so that every index the model's own cells use
//! stays what it was. A model that already carries dynamic-array cell
//! metadata lends this build its index; one that carries none gets this
//! writer's part. The model's own cells, strings, theme, calculation
//! chain and properties are never read beyond the central directory.
//!
//! A workbook an earlier build added sheets to is a model like any other:
//! its Frazaro parts (`xl/worksheets/frazaro_N.xml`, the `Frazaro.Build`
//! name, this writer's own metadata part) are dropped before this build's
//! go in, so building into one's own output replaces, never duplicates.
//! The cell formats an earlier build appended stay, unreferenced.
//!
//! New ground, said plainly: the reference opens a workbook in Excel and
//! adds sheets through the object model. What is here is the package
//! surgery that gives the same result in the file, with a model that is
//! someone else's file bounded at every step: a part inflates to at most
//! [`MAX_PART`] bytes, a checksum that does not match is refused, and no
//! entry name is ever taken for a path.

use super::inflate::inflate;
use super::ooxml::{self, Render};
use super::xml;
use super::zip::{self, Entry, RawEntry};
use super::Workbook;
use crate::messages::{raise, Refusal};
use crate::sha256::sha256_hex;

/// The most bytes one part of a model may inflate to.
pub const MAX_PART: usize = 256 << 20;

const WORKBOOK: &str = "xl/workbook.xml";
const WORKBOOK_RELS: &str = "xl/_rels/workbook.xml.rels";
const CONTENT_TYPES: &str = "[Content_Types].xml";
const STYLES: &str = "xl/styles.xml";
const METADATA: &str = "xl/metadata.xml";
const OUR_SHEET_PREFIX: &str = "xl/worksheets/frazaro_";
const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/// What the walk and the stamp need to know about the workbook this build
/// adds to.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct HostInfo {
    /// What a refusal calls the model: the path the door was given.
    pub label: String,
    /// The model's own sheets, which a program may name but never write to.
    pub sheet_names: Vec<String>,
    /// The SHA-256 of the model's bytes, as the stamp records it.
    pub model_hex: String,
    /// How many cell formats the model has: this build's first format
    /// after the default takes the next index.
    pub style_base: u32,
    /// The cell-metadata index a dynamic-array formula points at: the
    /// model's own dynamic-array block when it has one, 1 when this writer
    /// adds its part, `None` when the model's metadata part has no block
    /// to point at (a dynamic-array formula is then refused).
    pub cm: Option<u32>,
}

/// A workbook read far enough to add sheets to it.
#[derive(Debug)]
pub struct Model {
    label: String,
    bytes: Vec<u8>,
    entries: Vec<Entry>,
    workbook: String,
    rels: String,
    content_types: String,
    styles: String,
    metadata: Option<String>,
    sheet_names: Vec<String>,
    max_sheet_id: u32,
    rel_prefix: String,
    rel_ns_declared: bool,
    fills_count: u32,
    cellxfs_count: u32,
    dynamic_cm: Option<u32>,
    stripped: Vec<String>,
}

fn not_a_workbook(label: &str, why: &str) -> Refusal {
    raise(
        "build-into-not-a-workbook",
        &[("path", label), ("why", why)],
    )
}

fn unsupported(label: &str, part: &str, why: &str) -> Refusal {
    raise(
        "build-into-unsupported",
        &[("path", label), ("part", part), ("why", why)],
    )
}

/// An attribute's value in a start tag's text, unescaped.
fn attr(tag: &str, name: &str) -> Option<String> {
    let key = format!(" {name}=\"");
    let i = tag.find(&key)? + key.len();
    let j = tag[i..].find('"')?;
    Some(xml::unescape(&tag[i..i + j]))
}

/// The value of the attribute whose name ends with `suffix` (`r:id`, or
/// another prefix's `id`).
fn attr_ending(tag: &str, suffix: &str) -> Option<String> {
    let key = format!("{suffix}=\"");
    let i = tag.find(&key)? + key.len();
    let j = tag[i..].find('"')?;
    Some(xml::unescape(&tag[i..i + j]))
}

/// Every start tag of the element `name`: (start, end exclusive) spans of
/// `<name ...>` or `<name .../>`.
fn tags(text: &str, name: &str) -> Vec<(usize, usize)> {
    let open = format!("<{name}");
    let mut out = Vec::new();
    for (i, _) in text.match_indices(&open) {
        let after = text[i + open.len()..].chars().next();
        if !matches!(
            after,
            Some(' ') | Some('/') | Some('>') | Some('\n') | Some('\r') | Some('\t')
        ) {
            continue;
        }
        if let Some(j) = text[i..].find('>') {
            out.push((i, i + j + 1));
        }
    }
    out
}

/// The text with `insertion` put before the first `marker`.
fn insert_before(text: &str, marker: &str, insertion: &str) -> Option<String> {
    let i = text.find(marker)?;
    Some(format!("{}{}{}", &text[..i], insertion, &text[i..]))
}

/// The text with `insertion` put after the first `marker`.
fn insert_after(text: &str, marker: &str, insertion: &str) -> Option<String> {
    let i = text.find(marker)? + marker.len();
    Some(format!("{}{}{}", &text[..i], insertion, &text[i..]))
}

/// The `count` attribute of the first `<element ...>` tag.
fn count_of(text: &str, element: &str) -> Option<u32> {
    let (s, e) = *tags(text, element).first()?;
    attr(&text[s..e], "count")?.parse().ok()
}

/// The text with the first `<element ...>` tag's `count` raised by `add`.
fn bump_count(text: &str, element: &str, add: u32) -> Option<String> {
    let (s, e) = *tags(text, element).first()?;
    let tag = &text[s..e];
    let key = " count=\"";
    let i = tag.find(key)? + key.len();
    let j = tag[i..].find('"')?;
    let n: u32 = tag[i..i + j].parse().ok()?;
    let new_tag = format!("{}{}{}", &tag[..i], n + add, &tag[i + j..]);
    Some(format!("{}{}{}", &text[..s], new_tag, &text[e..]))
}

/// The text without the first self-closing tag of `name` that `keep`
/// rejects.
fn remove_tag(text: &str, name: &str, reject: impl Fn(&str) -> bool) -> String {
    for (s, e) in tags(text, name) {
        if reject(&text[s..e]) {
            let mut end = e;
            // Swallow one line break after the tag, so the part stays tidy.
            if text[end..].starts_with("\r\n") {
                end += 2;
            } else if text[end..].starts_with('\n') {
                end += 1;
            }
            return format!("{}{}", &text[..s], &text[end..]);
        }
    }
    text.to_string()
}

/// The text without the element whose start tag begins with `start` and
/// which ends with `end`.
fn remove_element(text: &str, start: &str, end: &str) -> String {
    let Some(s) = text.find(start) else {
        return text.to_string();
    };
    let Some(rel) = text[s..].find(end) else {
        return text.to_string();
    };
    let mut e = s + rel + end.len();
    if text[e..].starts_with("\r\n") {
        e += 2;
    } else if text[e..].starts_with('\n') {
        e += 1;
    }
    format!("{}{}", &text[..s], &text[e..])
}

/// A part's text: stored or inflated, its checksum and size checked, its
/// byte-order mark dropped.
fn part_text(
    bytes: &[u8],
    entries: &[Entry],
    name: &str,
    label: &str,
) -> Result<Option<String>, Refusal> {
    let Some(e) = entries.iter().find(|e| e.name == name) else {
        return Ok(None);
    };
    let raw = zip::raw_data(bytes, e).ok_or_else(|| {
        not_a_workbook(
            label,
            &format!("its part {name} reaches past the end of the file"),
        )
    })?;
    let data: Vec<u8> = match e.method {
        zip::STORED => raw.to_vec(),
        8 => inflate(raw, e.size as usize, MAX_PART).map_err(|why| {
            not_a_workbook(label, &format!("its part {name} does not inflate ({why})"))
        })?,
        m => {
            return Err(not_a_workbook(
                label,
                &format!(
                    "its part {name} uses compression method {m}, which this version does not read"
                ),
            ))
        }
    };
    if data.len() != e.size as usize {
        return Err(not_a_workbook(
            label,
            &format!("its part {name} is not the size its directory says"),
        ));
    }
    if zip::crc32(&data) != e.crc {
        return Err(not_a_workbook(
            label,
            &format!("its part {name} is damaged: its checksum does not match"),
        ));
    }
    let data = data
        .strip_prefix(b"\xEF\xBB\xBF")
        .map(<[u8]>::to_vec)
        .unwrap_or(data);
    String::from_utf8(data)
        .map(Some)
        .map_err(|_| not_a_workbook(label, &format!("its part {name} is not UTF-8")))
}

impl Model {
    /// A model read from its bytes; `label` is what a refusal calls it.
    pub fn read(bytes: Vec<u8>, label: &str) -> Result<Model, Refusal> {
        let entries = zip::entries(&bytes).ok_or_else(|| {
            not_a_workbook(
                label,
                zip::why_not_an_archive(&bytes)
                    .unwrap_or("it is not a zip archive this reader knows"),
            )
        })?;
        let Some(mut workbook) = part_text(&bytes, &entries, WORKBOOK, label)? else {
            return Err(not_a_workbook(
                label,
                "it has no xl/workbook.xml part, so it is not a spreadsheet package",
            ));
        };
        let mut rels = part_text(&bytes, &entries, WORKBOOK_RELS, label)?
            .ok_or_else(|| not_a_workbook(label, "it has no xl/_rels/workbook.xml.rels part"))?;
        let mut content_types = part_text(&bytes, &entries, CONTENT_TYPES, label)?
            .ok_or_else(|| not_a_workbook(label, "it has no [Content_Types].xml part"))?;
        let styles = part_text(&bytes, &entries, STYLES, label)?.ok_or_else(|| {
            unsupported(
                label,
                STYLES,
                "it is missing, and this version adds no styles part",
            )
        })?;
        let mut metadata = part_text(&bytes, &entries, METADATA, label)?;

        // An earlier build's parts come out first.
        let mut stripped: Vec<String> = Vec::new();
        for e in &entries {
            if !e.name.starts_with(OUR_SHEET_PREFIX) {
                continue;
            }
            let target = e.name.trim_start_matches("xl/").to_string();
            let mut rid: Option<String> = None;
            for (s, en) in tags(&rels, "Relationship") {
                if attr(&rels[s..en], "Target").as_deref() == Some(&target) {
                    rid = attr(&rels[s..en], "Id");
                }
            }
            rels = remove_tag(&rels, "Relationship", |t| {
                attr(t, "Target").as_deref() == Some(&target)
            });
            if let Some(rid) = rid {
                workbook = remove_tag(&workbook, "sheet", |t| {
                    attr_ending(t, ":id").as_deref() == Some(&rid)
                });
            }
            let part_name = format!("/{}", e.name);
            content_types = remove_tag(&content_types, "Override", |t| {
                attr(t, "PartName").as_deref() == Some(&part_name)
            });
            stripped.push(e.name.clone());
        }
        workbook = remove_element(
            &workbook,
            &format!("<definedName name=\"{}\"", super::super::build::STAMP_NAME),
            "</definedName>",
        );
        if metadata.as_deref() == Some(ooxml::metadata_xml().as_str()) {
            rels = remove_tag(&rels, "Relationship", |t| {
                attr(t, "Target").as_deref() == Some("metadata.xml")
            });
            content_types = remove_tag(&content_types, "Override", |t| {
                attr(t, "PartName").as_deref() == Some("/xl/metadata.xml")
            });
            metadata = None;
            stripped.push(METADATA.to_string());
        }

        // The model's sheets, and the ids this build's must not reuse.
        let mut sheet_names = Vec::new();
        let mut max_sheet_id = 0u32;
        for (s, e) in tags(&workbook, "sheet") {
            let tag = &workbook[s..e];
            if let Some(name) = attr(tag, "name") {
                sheet_names.push(name);
            }
            if let Some(id) = attr(tag, "sheetId").and_then(|v| v.parse::<u32>().ok()) {
                max_sheet_id = max_sheet_id.max(id);
            }
        }
        if !workbook.contains("</sheets>") {
            return Err(unsupported(
                label,
                WORKBOOK,
                "it has no sheet list to add to",
            ));
        }
        // The prefix the workbook binds to the relationships namespace.
        let (rel_prefix, rel_ns_declared) = match tags(&workbook, "workbook").first() {
            Some(&(s, e)) => {
                let root = &workbook[s..e];
                let key = format!("=\"{REL_NS}\"");
                match root.find(&key) {
                    Some(i) => {
                        let before = &root[..i];
                        let p = before.rfind("xmlns:").map(|k| before[k + 6..].to_string());
                        (p.unwrap_or_else(|| "r".to_string()), true)
                    }
                    None => ("r".to_string(), false),
                }
            }
            None => return Err(unsupported(label, WORKBOOK, "it has no workbook element")),
        };
        let fills_count = count_of(&styles, "fills")
            .ok_or_else(|| unsupported(label, STYLES, "it has no fills list"))?;
        let cellxfs_count = count_of(&styles, "cellXfs")
            .ok_or_else(|| unsupported(label, STYLES, "it has no cellXfs list"))?;
        if !styles.contains("</fills>") || !styles.contains("</cellXfs>") {
            return Err(unsupported(
                label,
                STYLES,
                "its fills or cellXfs list is self-closing",
            ));
        }
        let dynamic_cm = match &metadata {
            None => Some(1),
            Some(m) => dynamic_cm_of(m),
        };
        Ok(Model {
            label: label.to_string(),
            bytes,
            entries,
            workbook,
            rels,
            content_types,
            styles,
            metadata,
            sheet_names,
            max_sheet_id,
            rel_prefix,
            rel_ns_declared,
            fills_count,
            cellxfs_count,
            dynamic_cm,
            stripped,
        })
    }

    /// What the walk and the stamp need: the sheets, the fingerprint, the
    /// style base and the metadata index.
    pub fn host_info(&self) -> HostInfo {
        HostInfo {
            label: self.label.clone(),
            sheet_names: self.sheet_names.clone(),
            model_hex: sha256_hex(&self.bytes),
            style_base: self.cellxfs_count,
            cm: self.dynamic_cm,
        }
    }

    /// The model's sheet names, as it spells them.
    pub fn sheet_names(&self) -> &[String] {
        &self.sheet_names
    }

    /// The entries of an earlier build dropped on reading.
    pub fn stripped(&self) -> &[String] {
        &self.stripped
    }
}

/// The sheets of a workbook that are not this writer's, read from its
/// workbook part and relationships: those whose part is not a
/// `frazaro_N.xml`. For `frazaro rebuild` on a workbook a build was added
/// to, where the model is not at hand but its sheet names must be known
/// to walk the program again.
pub(crate) fn host_sheets_of(workbook: &str, rels: &str) -> Vec<String> {
    let mut ours: Vec<String> = Vec::new();
    for (s, e) in tags(rels, "Relationship") {
        let tag = &rels[s..e];
        if attr(tag, "Target").is_some_and(|t| t.starts_with("worksheets/frazaro_")) {
            if let Some(id) = attr(tag, "Id") {
                ours.push(id);
            }
        }
    }
    let mut names = Vec::new();
    for (s, e) in tags(workbook, "sheet") {
        let tag = &workbook[s..e];
        let rid = attr_ending(tag, ":id").unwrap_or_default();
        if ours.contains(&rid) {
            continue;
        }
        if let Some(name) = attr(tag, "name") {
            names.push(name);
        }
    }
    names
}

/// The 1-based index of the cell-metadata block that marks a dynamic-array
/// formula in a model's metadata part, when the part has one: the block
/// whose record names the XLDAPR type.
fn dynamic_cm_of(metadata: &str) -> Option<u32> {
    let mut xldapr: Option<usize> = None;
    for (i, (s, e)) in tags(metadata, "metadataType").iter().enumerate() {
        if attr(&metadata[*s..*e], "name").as_deref() == Some("XLDAPR") {
            xldapr = Some(i + 1);
        }
    }
    let t = xldapr?;
    let start = metadata.find("<cellMetadata")?;
    let end = start + metadata[start..].find("</cellMetadata>")?;
    let section = &metadata[start..end];
    for (i, (s, e)) in tags(section, "rc").iter().enumerate() {
        if attr(&section[*s..*e], "t").and_then(|v| v.parse::<usize>().ok()) == Some(t) {
            return u32::try_from(i + 1).ok();
        }
    }
    None
}

/// The model with this build's sheets added: every part of the model as it
/// was, the four it must edit edited, this build's parts appended. `cm` is
/// the metadata index the build's dynamic-array formulas were given.
pub fn write(model: &Model, ours: &Workbook, cm: u32) -> Result<Vec<u8>, Refusal> {
    let render = Render {
        style_base: model.cellxfs_count,
        cm,
    };
    let add_metadata = ours.has_dynamic_formulas() && model.metadata.is_none();
    let too_large = || raise("build-workbook-too-large", &[]);

    // The workbook: its sheet list, its names, its calcPr.
    let mut workbook = model.workbook.clone();
    let mut sheet_elements = String::new();
    for (i, sheet) in ours.sheets.iter().enumerate() {
        sheet_elements.push_str(&format!(
            "<sheet name=\"{}\" sheetId=\"{}\" {}:id=\"rIdFrazaro{}\"/>",
            xml::attr(&sheet.name),
            model.max_sheet_id + i as u32 + 1,
            model.rel_prefix,
            i + 1
        ));
    }
    workbook = insert_before(&workbook, "</sheets>", &sheet_elements).ok_or_else(too_large)?;
    if !model.rel_ns_declared {
        workbook = insert_after(
            &workbook,
            "<workbook",
            &format!(" xmlns:{}=\"{REL_NS}\"", model.rel_prefix),
        )
        .ok_or_else(too_large)?;
    }
    if !ours.defined_names.is_empty() {
        let mut names = String::new();
        for (name, formula) in &ours.defined_names {
            names.push_str(&format!(
                "<definedName name=\"{}\">{}</definedName>",
                xml::attr(name),
                xml::text(formula)
            ));
        }
        workbook = if workbook.contains("</definedNames>") {
            insert_before(&workbook, "</definedNames>", &names)
        } else if workbook.contains("<definedNames/>") {
            Some(workbook.replacen(
                "<definedNames/>",
                &format!("<definedNames>{names}</definedNames>"),
                1,
            ))
        } else {
            let wrapped = format!("<definedNames>{names}</definedNames>");
            let anchor = if workbook.contains("</externalReferences>") {
                "</externalReferences>"
            } else {
                "</sheets>"
            };
            insert_after(&workbook, anchor, &wrapped)
        }
        .ok_or_else(too_large)?;
    }
    workbook = match tags(&workbook, "calcPr").first() {
        Some(&(s, e)) => {
            let tag = &workbook[s..e];
            let new_tag = if let Some(i) = tag.find(" fullCalcOnLoad=\"") {
                let v = i + " fullCalcOnLoad=\"".len();
                let j = tag[v..].find('"').unwrap_or(0);
                format!("{}1{}", &tag[..v], &tag[v + j..])
            } else {
                insert_after(tag, "<calcPr", " fullCalcOnLoad=\"1\"").ok_or_else(too_large)?
            };
            format!("{}{}{}", &workbook[..s], new_tag, &workbook[e..])
        }
        None => {
            let anchor = if workbook.contains("</definedNames>") {
                "</definedNames>"
            } else if workbook.contains("</externalReferences>") {
                "</externalReferences>"
            } else {
                "</sheets>"
            };
            insert_after(&workbook, anchor, "<calcPr fullCalcOnLoad=\"1\"/>")
                .ok_or_else(too_large)?
        }
    };

    // The relationships and the content types.
    let mut rel_elements = String::new();
    let mut overrides = String::new();
    for i in 1..=ours.sheets.len() {
        rel_elements.push_str(&format!(
            "<Relationship Id=\"rIdFrazaro{i}\" Type=\"{}\" Target=\"worksheets/frazaro_{i}.xml\"/>",
            ooxml::REL_WORKSHEET
        ));
        overrides.push_str(&format!(
            "<Override PartName=\"/{}\" ContentType=\"{}\"/>",
            ooxml::sheet_part(i),
            ooxml::CT_WORKSHEET
        ));
    }
    if add_metadata {
        rel_elements.push_str(&format!(
            "<Relationship Id=\"{}\" Type=\"{}\" Target=\"metadata.xml\"/>",
            ooxml::METADATA_RID,
            ooxml::REL_METADATA
        ));
        overrides.push_str(&format!(
            "<Override PartName=\"/xl/metadata.xml\" ContentType=\"{}\"/>",
            ooxml::CT_METADATA
        ));
    }
    let rels =
        insert_before(&model.rels, "</Relationships>", &rel_elements).ok_or_else(too_large)?;
    let content_types =
        insert_before(&model.content_types, "</Types>", &overrides).ok_or_else(too_large)?;

    // The styles: this build's fills and formats at the ends of the lists.
    let fills = ours.styles.fills();
    let own_xfs = ours.styles.xfs().len().saturating_sub(1) as u32;
    let mut styles = model.styles.clone();
    if !fills.is_empty() {
        styles = insert_before(&styles, "</fills>", &ooxml::fills_fragment(&ours.styles))
            .ok_or_else(too_large)?;
        styles = bump_count(&styles, "fills", fills.len() as u32).ok_or_else(too_large)?;
    }
    if own_xfs > 0 {
        styles = insert_before(
            &styles,
            "</cellXfs>",
            &ooxml::cellxfs_fragment(&ours.styles, model.fills_count),
        )
        .ok_or_else(too_large)?;
        styles = bump_count(&styles, "cellXfs", own_xfs).ok_or_else(too_large)?;
    }

    // The entries: the model's in its order, ours after.
    let edited: Vec<(&str, &str)> = vec![
        (WORKBOOK, &workbook),
        (WORKBOOK_RELS, &rels),
        (CONTENT_TYPES, &content_types),
        (STYLES, &styles),
    ];
    let mut raws: Vec<RawEntry> = Vec::with_capacity(model.entries.len() + ours.sheets.len() + 1);
    for e in &model.entries {
        if model.stripped.contains(&e.name) {
            continue;
        }
        if let Some((_, text)) = edited.iter().find(|(n, _)| *n == e.name) {
            raws.push(RawEntry {
                name: &e.name,
                method: zip::STORED,
                crc: zip::crc32(text.as_bytes()),
                size: u32::try_from(text.len()).map_err(|_| too_large())?,
                bytes: text.as_bytes(),
            });
            continue;
        }
        raws.push(RawEntry {
            name: &e.name,
            method: e.method,
            crc: e.crc,
            size: e.size,
            bytes: zip::raw_data(&model.bytes, e).ok_or_else(too_large)?,
        });
    }
    let metadata_text = ooxml::metadata_xml();
    if add_metadata {
        raws.push(RawEntry {
            name: METADATA,
            method: zip::STORED,
            crc: zip::crc32(metadata_text.as_bytes()),
            size: u32::try_from(metadata_text.len()).map_err(|_| too_large())?,
            bytes: metadata_text.as_bytes(),
        });
    }
    let sheet_names: Vec<String> = (1..=ours.sheets.len()).map(ooxml::sheet_part).collect();
    let sheet_texts: Vec<String> = ours
        .sheets
        .iter()
        .map(|s| ooxml::sheet_xml(s, false, &render))
        .collect();
    for (name, text) in sheet_names.iter().zip(sheet_texts.iter()) {
        raws.push(RawEntry {
            name,
            method: zip::STORED,
            crc: zip::crc32(text.as_bytes()),
            size: u32::try_from(text.len()).map_err(|_| too_large())?,
            bytes: text.as_bytes(),
        });
    }
    zip::write_entries(&raws).ok_or_else(too_large)
}

/// A model of the fixture's shape, stored rather than deflated, for tests
/// that need a workbook with a sheet list, names, styles and a string
/// table without reading one from disk.
#[cfg(test)]
pub fn test_model(with_metadata: Option<&str>) -> Vec<u8> {
    let head = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n";
    let mut parts: Vec<(String, Vec<u8>)> = vec![
        (CONTENT_TYPES.to_string(), format!("{head}<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/><Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/><Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>{}</Types>", if with_metadata.is_some() { "<Override PartName=\"/xl/metadata.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml\"/>" } else { "" }).into_bytes()),
        ("_rels/.rels".to_string(), format!("{head}<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>").into_bytes()),
        (WORKBOOK.to_string(), format!("{head}<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"{REL_NS}\"><bookViews><workbookView activeTab=\"0\"/></bookViews><sheets><sheet name=\"Model\" sheetId=\"1\" r:id=\"rId1\"/><sheet name=\"Notes\" sheetId=\"4\" r:id=\"rId2\"/></sheets><definedNames><definedName name=\"Rate\">0.2</definedName></definedNames><calcPr calcId=\"191029\"/></workbook>").into_bytes()),
        (WORKBOOK_RELS.to_string(), format!("{head}<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"{}\" Target=\"worksheets/sheet1.xml\"/><Relationship Id=\"rId2\" Type=\"{}\" Target=\"worksheets/sheet2.xml\"/><Relationship Id=\"rId3\" Type=\"{}\" Target=\"styles.xml\"/>{}</Relationships>", ooxml::REL_WORKSHEET, ooxml::REL_WORKSHEET, ooxml::REL_STYLES, if with_metadata.is_some() { format!("<Relationship Id=\"rId4\" Type=\"{}\" Target=\"metadata.xml\"/>", ooxml::REL_METADATA) } else { String::new() }).into_bytes()),
        ("xl/worksheets/sheet1.xml".to_string(), format!("{head}<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\"/></sheetViews><sheetData><row r=\"3\"><c r=\"B3\"><f>B1-B2</f><v>400</v></c></row></sheetData></worksheet>").into_bytes()),
        ("xl/worksheets/sheet2.xml".to_string(), format!("{head}<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData/></worksheet>").into_bytes()),
        (STYLES.to_string(), format!("{head}<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><fonts count=\"1\"><font><sz val=\"11\"/></font></fonts><fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills><borders count=\"1\"><border/></borders><cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs><cellXfs count=\"3\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/></cellXfs><cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles></styleSheet>").into_bytes()),
    ];
    if let Some(m) = with_metadata {
        parts.push((METADATA.to_string(), format!("{head}{m}").into_bytes()));
    }
    zip::write_stored(&parts).unwrap()
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::sheet::{parse_a1_range, Cell, Content, Sheet};

    const EXCEL_METADATA: &str = "<metadata xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:xlrd=\"http://schemas.microsoft.com/office/spreadsheetml/2017/richdata\" xmlns:xda=\"http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray\"><metadataTypes count=\"2\"><metadataType name=\"XLRICHVALUE\" minSupportedVersion=\"120000\" copy=\"1\" pasteAll=\"1\" pasteValues=\"1\" merge=\"1\" splitFirst=\"1\" rowColShift=\"1\" clearFormats=\"1\" clearComments=\"1\" assign=\"1\" coerce=\"1\"/><metadataType name=\"XLDAPR\" minSupportedVersion=\"120000\" copy=\"1\" pasteAll=\"1\" pasteValues=\"1\" merge=\"1\" splitFirst=\"1\" rowColShift=\"1\" clearFormats=\"1\" clearComments=\"1\" assign=\"1\" coerce=\"1\" cellMeta=\"1\"/></metadataTypes><futureMetadata name=\"XLDAPR\" count=\"1\"><bk><extLst><ext uri=\"{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}\"><xda:dynamicArrayProperties fDynamic=\"1\" fCollapsed=\"0\"/></ext></extLst></bk></futureMetadata><cellMetadata count=\"2\"><bk><rc t=\"1\" v=\"0\"/></bk><bk><rc t=\"2\" v=\"0\"/></bk></cellMetadata></metadata>";

    fn ours() -> Workbook {
        let mut wb = Workbook::new();
        let st = wb.styles.id(crate::sheet::Style {
            fill: Some(crate::sheet::Rgb(1, 2, 3)),
            num_fmt: crate::sheet::NumFmt::Text,
            wrap: false,
        });
        let mut a = Sheet::new("Frazaro");
        a.set(
            1,
            2,
            Cell {
                content: Content::Text("Put 1 into cell A1.".to_string()),
                style: st,
            },
        );
        let mut b = Sheet::new("Output");
        b.set_formula_dynamic(parse_a1_range("a1").unwrap(), "_xlfn.IFS(1,2)", 0);
        wb.sheets.push(a);
        wb.sheets.push(b);
        wb.defined_names
            .push(("Frazaro.Build".to_string(), "\"stamp\"".to_string()));
        wb
    }

    #[test]
    fn a_model_is_read_far_enough_and_no_further() {
        let m = Model::read(test_model(None), "model.xlsx").unwrap();
        assert_eq!(m.sheet_names(), ["Model", "Notes"]);
        assert_eq!(m.max_sheet_id, 4);
        assert_eq!(m.rel_prefix, "r");
        assert!(m.rel_ns_declared);
        assert_eq!((m.fills_count, m.cellxfs_count), (2, 3));
        let info = m.host_info();
        assert_eq!(info.style_base, 3);
        assert_eq!(info.cm, Some(1));
        assert_eq!(info.model_hex.len(), 64);
        assert!(m.stripped().is_empty());
        // Excel's own metadata lends its dynamic-array block's index.
        let m2 = Model::read(test_model(Some(EXCEL_METADATA)), "m").unwrap();
        assert_eq!(m2.host_info().cm, Some(2));
        let no_block = EXCEL_METADATA.replace("XLDAPR", "XLOTHER");
        let m3 = Model::read(test_model(Some(&no_block)), "m").unwrap();
        assert_eq!(m3.host_info().cm, None);
    }

    #[test]
    fn what_is_not_a_model_is_refused_with_why() {
        let r = Model::read(b"nope".to_vec(), "x.xlsx").unwrap_err();
        assert_eq!(r.id, "build-into-not-a-workbook");
        assert!(r.text.contains("x.xlsx is not a workbook"), "{}", r.text);
        assert!(r.text.contains("not a zip archive"), "{}", r.text);
        let no_wb = zip::write_stored(&[("a.txt".to_string(), b"hi".to_vec())]).unwrap();
        let r = Model::read(no_wb, "x.zip").unwrap_err();
        assert!(r.text.contains("no xl/workbook.xml"), "{}", r.text);
        let ole = vec![0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, 0, 0, 0, 0];
        let r = Model::read(ole, "old.xls").unwrap_err();
        assert!(r.text.contains("older .xls or an encrypted"), "{}", r.text);
    }

    #[test]
    fn the_merge_edits_four_parts_and_copies_the_rest() {
        let model_bytes = test_model(None);
        let model = Model::read(model_bytes.clone(), "model.xlsx").unwrap();
        let out = write(&model, &ours(), 1).unwrap();
        let before = zip::entries(&model_bytes).unwrap();
        let after = zip::entries(&out).unwrap();
        let names: Vec<&str> = after.iter().map(|e| e.name.as_str()).collect();
        assert_eq!(
            names,
            [
                "[Content_Types].xml",
                "_rels/.rels",
                "xl/workbook.xml",
                "xl/_rels/workbook.xml.rels",
                "xl/worksheets/sheet1.xml",
                "xl/worksheets/sheet2.xml",
                "xl/styles.xml",
                "xl/metadata.xml",
                "xl/worksheets/frazaro_1.xml",
                "xl/worksheets/frazaro_2.xml",
            ]
        );
        // The untouched parts are byte for byte the model's.
        for name in [
            "_rels/.rels",
            "xl/worksheets/sheet1.xml",
            "xl/worksheets/sheet2.xml",
        ] {
            let a = before.iter().find(|e| e.name == name).unwrap();
            let b = after.iter().find(|e| e.name == name).unwrap();
            assert_eq!(
                zip::raw_data(&model_bytes, a),
                zip::raw_data(&out, b),
                "{name}"
            );
            assert_eq!(a.crc, b.crc);
        }
        let text = |name: &str| {
            let e = after.iter().find(|e| e.name == name).unwrap();
            String::from_utf8(zip::stored_data(&out, e).unwrap().to_vec()).unwrap()
        };
        let wb = text("xl/workbook.xml");
        assert!(wb.contains("<sheet name=\"Notes\" sheetId=\"4\" r:id=\"rId2\"/><sheet name=\"Frazaro\" sheetId=\"5\" r:id=\"rIdFrazaro1\"/><sheet name=\"Output\" sheetId=\"6\" r:id=\"rIdFrazaro2\"/></sheets>"), "{wb}");
        assert!(wb.contains("<definedName name=\"Rate\">0.2</definedName><definedName name=\"Frazaro.Build\">\"stamp\"</definedName></definedNames>"), "{wb}");
        assert!(
            wb.contains("<calcPr fullCalcOnLoad=\"1\" calcId=\"191029\"/>"),
            "{wb}"
        );
        let rels = text("xl/_rels/workbook.xml.rels");
        assert!(
            rels.contains("Target=\"worksheets/frazaro_1.xml\"/><Relationship Id=\"rIdFrazaro2\"")
        );
        assert!(rels.contains("Id=\"rIdFrazaroMetadata\""));
        let ct = text("[Content_Types].xml");
        assert!(ct.contains("<Override PartName=\"/xl/worksheets/frazaro_1.xml\""));
        assert!(ct.contains("<Override PartName=\"/xl/metadata.xml\""));
        let st = text("xl/styles.xml");
        assert!(st.contains("<fills count=\"3\">"), "{st}");
        assert!(st.contains("<fgColor rgb=\"FF010203\"/>"));
        assert!(st.contains("<cellXfs count=\"4\">"), "{st}");
        assert!(st.contains("<xf numFmtId=\"49\" fontId=\"0\" fillId=\"2\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\" applyFill=\"1\"/></cellXfs>"), "{st}");
        // Our sheets name the model's formats and metadata: style 1 is the
        // model's 3, cm is 1, and no tab is selected twice.
        let s1 = text("xl/worksheets/frazaro_1.xml");
        assert!(s1.contains("<c r=\"B1\" s=\"3\" t=\"inlineStr\">"), "{s1}");
        assert!(!s1.contains("tabSelected"));
        let s2 = text("xl/worksheets/frazaro_2.xml");
        assert!(
            s2.contains(
                "<c r=\"A1\" s=\"0\" cm=\"1\"><f t=\"array\" ref=\"A1\">_xlfn.IFS(1,2)</f></c>"
            ),
            "{s2}"
        );
        // Determinism, and the same bytes from the same inputs.
        assert_eq!(write(&model, &ours(), 1).unwrap(), out);
    }

    #[test]
    fn building_into_ones_own_output_replaces_the_earlier_build() {
        let model = Model::read(test_model(None), "model.xlsx").unwrap();
        let first = write(&model, &ours(), 1).unwrap();
        let again = Model::read(first.clone(), "first.xlsx").unwrap();
        assert_eq!(again.sheet_names(), ["Model", "Notes"]);
        assert_eq!(
            again.stripped(),
            [
                "xl/worksheets/frazaro_1.xml",
                "xl/worksheets/frazaro_2.xml",
                "xl/metadata.xml"
            ]
        );
        assert_eq!(again.host_info().cm, Some(1));
        assert!(!again.workbook.contains("Frazaro.Build"));
        assert!(!again.rels.contains("rIdFrazaro"));
        assert!(!again.content_types.contains("frazaro_"));
        // The earlier build's formats stay, unreferenced: the base moves up.
        assert_eq!(again.host_info().style_base, 4);
        let second = write(&again, &ours(), 1).unwrap();
        let names: Vec<String> = zip::entries(&second)
            .unwrap()
            .into_iter()
            .map(|e| e.name)
            .collect();
        assert_eq!(
            names
                .iter()
                .filter(|n| n.starts_with(OUR_SHEET_PREFIX))
                .count(),
            2
        );
        assert_eq!(names.iter().filter(|n| *n == METADATA).count(), 1);
    }

    #[test]
    fn a_model_with_excel_metadata_lends_its_index_and_gets_no_part() {
        let model = Model::read(test_model(Some(EXCEL_METADATA)), "m.xlsx").unwrap();
        let cm = model.host_info().cm.unwrap();
        let out = write(&model, &ours(), cm).unwrap();
        let after = zip::entries(&out).unwrap();
        assert_eq!(after.iter().filter(|e| e.name == METADATA).count(), 1);
        let e = after.iter().find(|e| e.name == METADATA).unwrap();
        assert_eq!(e.method, zip::STORED); // the model's own, copied as it was
        let s2 = after
            .iter()
            .find(|e| e.name == "xl/worksheets/frazaro_2.xml")
            .unwrap();
        let text = String::from_utf8(zip::stored_data(&out, s2).unwrap().to_vec()).unwrap();
        assert!(text.contains("cm=\"2\""), "{text}");
    }

    /// The fixture model, deflated by .NET (tools/build_model_fixture.ps1):
    /// the inflate decoder on real streams, and the reader on a package of
    /// Excel's shape with a string table and a calculation chain.
    #[test]
    fn the_fixture_model_reads() {
        const MODEL: &[u8] = include_bytes!("../../../scripts/build/model.xlsx");
        let m = Model::read(MODEL.to_vec(), "scripts/build/model.xlsx").unwrap();
        assert_eq!(m.sheet_names(), ["Model", "Notes"]);
        assert_eq!(m.max_sheet_id, 2);
        assert_eq!((m.fills_count, m.cellxfs_count), (2, 2));
        assert_eq!(m.host_info().cm, Some(1));
        assert_eq!(m.entries.len(), 11);
        assert!(
            m.entries.iter().all(|e| e.method == 8),
            "the fixture is deflated throughout"
        );
        assert!(m
            .workbook
            .contains("<definedName name=\"Rate\">0.2</definedName>"));
        assert!(m.workbook.contains("<calcPr calcId=\"191029\"/>"));
        let out = write(&m, &ours(), 1).unwrap();
        let after = zip::entries(&out).unwrap();
        assert_eq!(after.len(), 11 + 3);
        // The model's deflated parts travel as they were.
        let before = zip::entries(MODEL).unwrap();
        for name in [
            "xl/sharedStrings.xml",
            "xl/calcChain.xml",
            "xl/worksheets/sheet1.xml",
            "docProps/core.xml",
        ] {
            let a = before.iter().find(|e| e.name == name).unwrap();
            let b = after.iter().find(|e| e.name == name).unwrap();
            assert_eq!(b.method, 8);
            assert_eq!(zip::raw_data(MODEL, a), zip::raw_data(&out, b), "{name}");
        }
        assert_eq!(
            host_sheets_of(
                "<sheets><sheet name=\"Model\" sheetId=\"1\" r:id=\"rId1\"/><sheet name=\"Frazaro\" sheetId=\"3\" r:id=\"rIdFrazaro1\"/></sheets>",
                "<Relationships><Relationship Id=\"rId1\" Type=\"t\" Target=\"worksheets/sheet1.xml\"/><Relationship Id=\"rIdFrazaro1\" Type=\"t\" Target=\"worksheets/frazaro_1.xml\"/></Relationships>"
            ),
            ["Model"]
        );
    }

    #[test]
    fn the_text_helpers() {
        assert_eq!(
            count_of("<a><fills count=\"7\"><x/></fills></a>", "fills"),
            Some(7)
        );
        assert_eq!(
            bump_count("<fills count=\"7\">", "fills", 2).unwrap(),
            "<fills count=\"9\">"
        );
        assert_eq!(
            remove_tag(
                "<s><sheet name=\"A\" r:id=\"r1\"/>\n<sheet name=\"B\" r:id=\"r2\"/></s>",
                "sheet",
                |t| attr_ending(t, ":id").as_deref() == Some("r1")
            ),
            "<s><sheet name=\"B\" r:id=\"r2\"/></s>"
        );
        assert_eq!(
            remove_element("<n><definedName name=\"X\">1</definedName><definedName name=\"Y\">2</definedName></n>", "<definedName name=\"X\"", "</definedName>"),
            "<n><definedName name=\"Y\">2</definedName></n>"
        );
        assert_eq!(
            attr("<sheet name=\"a &amp; b\" sheetId=\"3\"/>", "name").as_deref(),
            Some("a & b")
        );
        assert_eq!(
            tags("<sheets><sheet a=\"1\"/><sheetView/></sheets>", "sheet").len(),
            1
        );
    }
}
