//! XML text as the OOXML parts need it (PORT.7): the markup characters
//! escaped, and the characters XML 1.0 cannot carry spelled as Excel spells
//! them.
//!
//! New ground, said plainly: the reference writes cells through Excel's
//! object model and never spells a part by hand, so there is no VBA arm to
//! port. Excel's own convention for a character the format cannot hold is
//! `_xHHHH_`, the code unit in four upper-case hex digits; a text that
//! happens to contain that spelling has its underscore written `_x005F_`,
//! so that it reads back as written.

/// A text as element content: `&`, `<` and `>` escaped, `"` left alone, a
/// tab and a line break kept.
pub fn text(s: &str) -> String {
    escape(s, false)
}

/// A text as an attribute value between double quotes: as [`text`], with
/// `"` escaped too and a line break as a character reference, so that it
/// survives attribute-value normalization.
pub fn attr(s: &str) -> String {
    escape(s, true)
}

fn escape(s: &str, in_attr: bool) -> String {
    let chars: Vec<char> = s.chars().collect();
    let mut out = String::with_capacity(s.len() + 8);
    for (i, &c) in chars.iter().enumerate() {
        match c {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' if in_attr => out.push_str("&quot;"),
            '\n' if in_attr => out.push_str("&#10;"),
            '\r' if in_attr => out.push_str("&#13;"),
            '\t' | '\n' | '\r' => out.push(c),
            '_' if is_escape_spelling(&chars[i..]) => out.push_str("_x005F_"),
            c if (c as u32) < 0x20 || c == '\u{FFFE}' || c == '\u{FFFF}' => {
                out.push_str(&format!("_x{:04X}_", c as u32));
            }
            c => out.push(c),
        }
    }
    out
}

/// `_xHHHH_` at the start of the slice: Excel's escape spelling, which a
/// text must not be mistaken for.
fn is_escape_spelling(rest: &[char]) -> bool {
    rest.len() >= 7
        && rest[0] == '_'
        && rest[1] == 'x'
        && rest[2..6].iter().all(|c| c.is_ascii_hexdigit())
        && rest[6] == '_'
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn the_markup_characters_are_escaped() {
        assert_eq!(text("a & b < c > d \"e\""), "a &amp; b &lt; c &gt; d \"e\"");
        assert_eq!(attr("say \"hi\"\nnow"), "say &quot;hi&quot;&#10;now");
        assert_eq!(text("tab\there"), "tab\there");
    }

    #[test]
    fn a_character_xml_cannot_hold_is_spelled_as_excel_spells_it() {
        assert_eq!(text("a\u{1}b"), "a_x0001_b");
        assert_eq!(text("\u{FFFE}"), "_xFFFE_");
        // A text that looks like the spelling keeps its underscore escaped.
        assert_eq!(text("_x0041_"), "_x005F_x0041_");
        assert_eq!(text("_xZZZZ_"), "_xZZZZ_");
        assert_eq!(text("_x004"), "_x004");
    }
}
