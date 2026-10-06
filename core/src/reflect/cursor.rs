//! A cursor over a part's XML (PORT.8, slice 8a): the subset sheet XML
//! needs and no more.
//!
//! New ground, said plainly: the reference reads a workbook through Excel's
//! object model and never parses a part. What is here is a forward scan over
//! a part's text that yields start tags, end tags and the text between them,
//! with comments and processing instructions skipped, a CDATA section read
//! as text, entity references undone, and element and attribute names
//! matched by their local part, so that a writer that prefixes the main
//! namespace reads the same. It builds no tree: the walkers in `ooxml.rs`
//! take the events as they come and hold only what a relation row needs,
//! which is what lets a sheet part be walked once and dropped.
//!
//! Refused outright: a `<!DOCTYPE` or `<!ENTITY` declaration anywhere, since
//! no entity is ever expanded here (docs/THREAT_MODEL.md, section 5), and a
//! tag or section that never closes. Everything the cursor does not
//! understand is passed over unread.

use std::borrow::Cow;

use crate::sheet::xml;

/// Why a declaration is refused, in the words the catalogue's `{why}` takes.
pub const DOCTYPE_REFUSED: &str =
    "it declares a DOCTYPE or an entity, and no entity is ever expanded here";

/// One event of the scan. Names are local names, without their prefix.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Event<'a> {
    /// `<name attrs>` or, with `empty`, `<name attrs/>`.
    Start {
        name: &'a str,
        attrs: &'a str,
        empty: bool,
    },
    /// `</name>`.
    End { name: &'a str },
    /// Character data between tags, as written; `raw` for a CDATA section,
    /// in which nothing is an entity reference.
    Text { text: &'a str, raw: bool },
}

/// The scan's position in a part's text.
pub struct Cursor<'a> {
    text: &'a str,
    pos: usize,
}

/// The name after its prefix: `c` for `x:c`.
fn local(name: &str) -> &str {
    match name.rfind(':') {
        Some(i) => &name[i + 1..],
        None => name,
    }
}

impl<'a> Cursor<'a> {
    pub fn new(text: &'a str) -> Cursor<'a> {
        Cursor { text, pos: 0 }
    }

    /// Where the scan stands: the offset just past the last event, so that
    /// a reader can remember where an element began (the offset before the
    /// call that returned its start tag) and where it ended.
    pub fn position(&self) -> usize {
        self.pos
    }

    /// The next event, or `None` at the end of the text. An error names
    /// what the markup does that this reader refuses.
    pub fn next_event(&mut self) -> Result<Option<Event<'a>>, &'static str> {
        let text: &'a str = self.text;
        loop {
            let rest = &text[self.pos..];
            if rest.is_empty() {
                return Ok(None);
            }
            if !rest.starts_with('<') {
                let end = rest.find('<').unwrap_or(rest.len());
                self.pos += end;
                return Ok(Some(Event::Text {
                    text: &rest[..end],
                    raw: false,
                }));
            }
            if let Some(after) = rest.strip_prefix("<!--") {
                let end = after.find("-->").ok_or("a comment never closes")?;
                self.pos += 4 + end + 3;
                continue;
            }
            if let Some(after) = rest.strip_prefix("<![CDATA[") {
                let end = after.find("]]>").ok_or("a CDATA section never closes")?;
                self.pos += 9 + end + 3;
                return Ok(Some(Event::Text {
                    text: &after[..end],
                    raw: true,
                }));
            }
            if rest.starts_with("<!") {
                return Err(DOCTYPE_REFUSED);
            }
            if let Some(after) = rest.strip_prefix("<?") {
                let end = after
                    .find("?>")
                    .ok_or("a processing instruction never closes")?;
                self.pos += 2 + end + 2;
                continue;
            }
            if let Some(after) = rest.strip_prefix("</") {
                let end = after.find('>').ok_or("an end tag never closes")?;
                self.pos += 2 + end + 1;
                return Ok(Some(Event::End {
                    name: local(after[..end].trim_end()),
                }));
            }
            // A start tag: the name, then the attributes to the `>` that
            // is outside every quoted value.
            let after = &rest[1..];
            let name_end = after
                .find(|c: char| c.is_whitespace() || c == '/' || c == '>')
                .ok_or("a tag never closes")?;
            if name_end == 0 {
                return Err("a tag has no name");
            }
            let bytes = after.as_bytes();
            let mut quote: Option<u8> = None;
            let mut j = name_end;
            let close = loop {
                let Some(&b) = bytes.get(j) else {
                    return Err("a tag never closes");
                };
                match quote {
                    Some(q) => {
                        if b == q {
                            quote = None;
                        }
                    }
                    None => match b {
                        b'"' | b'\'' => quote = Some(b),
                        b'>' => break j,
                        _ => {}
                    },
                }
                j += 1;
            };
            let empty = close > name_end && bytes[close - 1] == b'/';
            let attrs_end = if empty { close - 1 } else { close };
            self.pos += 1 + close + 1;
            return Ok(Some(Event::Start {
                name: local(&after[..name_end]),
                attrs: &after[name_end..attrs_end],
                empty,
            }));
        }
    }
}

/// The attributes of a start tag, each as its local name and its value as
/// written, entities not yet undone.
pub fn attributes(attrs: &str) -> Attributes<'_> {
    Attributes {
        rest: attrs,
        full: false,
    }
}

/// The attributes of a start tag with their names as written, prefix and
/// all, for the one case two namespaces share a local name (OpenDocument's
/// `office:value-type` beside `calcext:value-type`).
pub fn attributes_full(attrs: &str) -> Attributes<'_> {
    Attributes {
        rest: attrs,
        full: true,
    }
}

pub struct Attributes<'a> {
    rest: &'a str,
    full: bool,
}

impl<'a> Iterator for Attributes<'a> {
    type Item = (&'a str, &'a str);

    fn next(&mut self) -> Option<Self::Item> {
        let s = self.rest.trim_start();
        let eq = s.find('=')?;
        let name = s[..eq].trim_end();
        let after = s[eq + 1..].trim_start();
        let q = *after.as_bytes().first()?;
        if q != b'"' && q != b'\'' {
            self.rest = "";
            return None;
        }
        let end = after[1..].find(q as char)?;
        let value = &after[1..1 + end];
        self.rest = &after[1 + end + 1..];
        Some((if self.full { name } else { local(name) }, value))
    }
}

/// An attribute's value by its local name, entities undone.
pub fn attr<'a>(attrs: &'a str, name: &str) -> Option<Cow<'a, str>> {
    attributes(attrs)
        .find(|(n, _)| *n == name)
        .map(|(_, v)| unescape_entities(v))
}

/// An attribute's value by its name as written, prefix and all, entities
/// undone.
pub fn attr_full<'a>(attrs: &'a str, name: &str) -> Option<Cow<'a, str>> {
    attributes_full(attrs)
        .find(|(n, _)| *n == name)
        .map(|(_, v)| unescape_entities(v))
}

/// Entity references undone: the five named ones and numeric character
/// references; a reference that does not parse is kept as written. No
/// allocation when there is none.
pub fn unescape_entities(s: &str) -> Cow<'_, str> {
    if !s.contains('&') {
        return Cow::Borrowed(s);
    }
    let mut out = String::with_capacity(s.len());
    let mut rest = s;
    while let Some(i) = rest.find('&') {
        out.push_str(&rest[..i]);
        let tail = &rest[i..];
        let decoded = match tail.find(';') {
            Some(end) if end <= 12 => {
                let entity = &tail[1..end];
                let c = match entity {
                    "amp" => Some('&'),
                    "lt" => Some('<'),
                    "gt" => Some('>'),
                    "quot" => Some('"'),
                    "apos" => Some('\''),
                    e => e
                        .strip_prefix('#')
                        .and_then(|n| match n.strip_prefix('x') {
                            Some(h) => u32::from_str_radix(h, 16).ok(),
                            None => n.parse::<u32>().ok(),
                        })
                        .and_then(char::from_u32),
                };
                c.map(|c| (c, end + 1))
            }
            _ => None,
        };
        match decoded {
            Some((c, len)) => {
                out.push(c);
                rest = &tail[len..];
            }
            None => {
                out.push('&');
                rest = &tail[1..];
            }
        }
    }
    out.push_str(rest);
    Cow::Owned(out)
}

/// A string's or a cell's text: entities and Excel's `_xHHHH_` spelling
/// undone (`sheet::xml::unescape`); no allocation when there is neither.
pub fn unescape_text(s: &str) -> Cow<'_, str> {
    if !s.contains('&') && !s.contains("_x") {
        return Cow::Borrowed(s);
    }
    Cow::Owned(xml::unescape(s))
}

/// The text of the element a start event (not an empty one) just opened,
/// to its end tag: nested elements' text included, except under an element
/// named in `skip`; entities undone, and Excel's `_xHHHH_` too when
/// `excel_escapes` says the text is a string of the kind Excel spells that
/// way (a shared or inline string), not a formula or a value.
pub fn element_text(
    cursor: &mut Cursor,
    skip: &[&str],
    excel_escapes: bool,
) -> Result<String, &'static str> {
    let mut out = String::new();
    let mut depth = 0u32;
    loop {
        match cursor.next_event()? {
            None => return Err("an element never closes"),
            Some(Event::Start { name, empty, .. }) => {
                if empty {
                    continue;
                }
                if skip.contains(&name) {
                    skip_element(cursor)?;
                } else {
                    depth += 1;
                }
            }
            Some(Event::End { .. }) => {
                if depth == 0 {
                    return Ok(out);
                }
                depth -= 1;
            }
            Some(Event::Text { text, raw }) => {
                if raw {
                    out.push_str(text);
                } else if excel_escapes {
                    out.push_str(&unescape_text(text));
                } else {
                    out.push_str(&unescape_entities(text));
                }
            }
        }
    }
}

/// Past the element a start event (not an empty one) just opened, to its
/// end tag, whatever is nested inside.
pub fn skip_element(cursor: &mut Cursor) -> Result<(), &'static str> {
    let mut depth = 1u32;
    loop {
        match cursor.next_event()? {
            None => return Err("an element never closes"),
            Some(Event::Start { empty: false, .. }) => depth += 1,
            Some(Event::End { .. }) => {
                depth -= 1;
                if depth == 0 {
                    return Ok(());
                }
            }
            _ => {}
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn events(text: &str) -> Vec<String> {
        let mut cur = Cursor::new(text);
        let mut out = Vec::new();
        while let Some(ev) = cur.next_event().unwrap() {
            out.push(match ev {
                Event::Start { name, attrs, empty } => {
                    format!("<{name}|{attrs}|{}>", if empty { "/" } else { "" })
                }
                Event::End { name } => format!("</{name}>"),
                Event::Text { text, raw } => format!("{}{text}", if raw { "!" } else { "" }),
            });
        }
        out
    }

    #[test]
    fn the_events_of_a_part_come_in_order_with_local_names() {
        assert_eq!(
            events(
                "<?xml version=\"1.0\"?>\n<x:worksheet xmlns:x=\"u\"><!-- c --><sheetData><row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c><c r='B1'/></row></sheetData></x:worksheet>"
            ),
            [
                "\n",
                "<worksheet| xmlns:x=\"u\"|>",
                "<sheetData||>",
                "<row| r=\"1\"|>",
                "<c| r=\"A1\" t=\"s\"|>",
                "<v||>",
                "0",
                "</v>",
                "</c>",
                "<c| r='B1'|/>",
                "</row>",
                "</sheetData>",
                "</worksheet>",
            ]
        );
        assert_eq!(
            events("<a b=\"x>y\"><![CDATA[1 < 2]]></a>"),
            ["<a| b=\"x>y\"|>", "!1 < 2", "</a>"]
        );
    }

    #[test]
    fn attributes_read_by_local_name_in_either_quote() {
        let attrs = " name=\"Q1 &amp; Q2\" sheetId='3' r:id=\"rId7\" state=\"hidden\"";
        assert_eq!(attr(attrs, "name").as_deref(), Some("Q1 & Q2"));
        assert_eq!(attr(attrs, "sheetId").as_deref(), Some("3"));
        assert_eq!(attr(attrs, "id").as_deref(), Some("rId7"));
        assert_eq!(attr(attrs, "state").as_deref(), Some("hidden"));
        assert_eq!(attr(attrs, "missing"), None);
        assert_eq!(attributes(attrs).count(), 4);
        assert_eq!(attributes("").count(), 0);
    }

    #[test]
    fn what_is_refused_or_malformed_is_named() {
        let mut cur = Cursor::new("<!DOCTYPE x [<!ENTITY a \"b\">]><a/>");
        assert_eq!(cur.next_event(), Err(DOCTYPE_REFUSED));
        let mut cur = Cursor::new("<a><b attr=\"1\"");
        assert_eq!(
            cur.next_event(),
            Ok(Some(Event::Start {
                name: "a",
                attrs: "",
                empty: false
            }))
        );
        assert_eq!(cur.next_event(), Err("a tag never closes"));
        let mut cur = Cursor::new("<!-- open");
        assert_eq!(cur.next_event(), Err("a comment never closes"));
        let mut cur = Cursor::new("<a>");
        cur.next_event().unwrap();
        assert_eq!(
            element_text(&mut cur, &[], false),
            Err("an element never closes")
        );
    }

    #[test]
    fn element_text_gathers_runs_and_skips_phonetics() {
        let text = "<si><r><rPr><b/></rPr><t>Rate </t></r><r><t xml:space=\"preserve\">applies &amp; _x000D_</t></r><rPh sb=\"0\" eb=\"1\"><t>ignored</t></rPh><phoneticPr fontId=\"1\"/></si>";
        let mut cur = Cursor::new(text);
        cur.next_event().unwrap();
        assert_eq!(
            element_text(&mut cur, &["rPh"], true).unwrap(),
            "Rate applies & \r"
        );
        assert_eq!(cur.next_event(), Ok(None));
        // A formula's text undoes entities and nothing else.
        let mut cur = Cursor::new("<f>A1&amp;\"_x0041_\"&lt;B1</f>");
        cur.next_event().unwrap();
        assert_eq!(
            element_text(&mut cur, &[], false).unwrap(),
            "A1&\"_x0041_\"<B1"
        );
    }

    #[test]
    fn entities_undone_and_the_rest_kept() {
        assert_eq!(unescape_entities("plain"), "plain");
        assert!(matches!(unescape_entities("plain"), Cow::Borrowed(_)));
        assert_eq!(
            unescape_entities(
                "a &amp; b &lt; c &gt; d &quot;e&quot; &apos;f&apos; &#65;&#x42; &unknown; & g"
            ),
            "a & b < c > d \"e\" 'f' AB &unknown; & g"
        );
        assert_eq!(unescape_text("x_x000A_y"), "x\ny");
        assert!(matches!(unescape_text("plain"), Cow::Borrowed(_)));
    }

    #[test]
    fn skipping_an_element_passes_what_is_nested() {
        let mut cur = Cursor::new("<a><b><c/><d>t</d></b><e/></a>");
        assert!(matches!(
            cur.next_event(),
            Ok(Some(Event::Start { name: "a", .. }))
        ));
        assert!(matches!(
            cur.next_event(),
            Ok(Some(Event::Start { name: "b", .. }))
        ));
        skip_element(&mut cur).unwrap();
        assert!(matches!(
            cur.next_event(),
            Ok(Some(Event::Start {
                name: "e",
                empty: true,
                ..
            }))
        ));
    }
}
