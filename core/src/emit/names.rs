//! Names: VLA.bas's `SymName`, `TransliterateToAscii`, `TransliterateChar`,
//! `ClickHandlerSlug` and `ToVbaString`.

use crate::intrinsics::is_numeric;
use crate::messages::{raise, Refusal};

/// VBA `SymName`: a numeric token passes through; non-ASCII letters
/// transliterate to their closest ASCII (LX.6); hyphens and colons become
/// underscores; dots stay, so `worksheet.cells` works unmodified. Refuses
/// a name that transliteration leaves empty.
pub fn sym_name(s: &str) -> Result<String, Refusal> {
    if is_numeric(s) {
        return Ok(s.to_string());
    }
    let r = transliterate_to_ascii(s).replace(['-', ':'], "_");
    if r.is_empty() {
        return Err(raise("vla-name-no-plain-letters", &[("value", s)]));
    }
    Ok(r)
}

/// VBA `TransliterateToAscii`: ASCII stays; everything else goes through
/// the table, which drops what it does not know rather than guessing.
pub fn transliterate_to_ascii(s: &str) -> String {
    let mut buf = String::with_capacity(s.len());
    for c in s.chars() {
        if c.is_ascii() {
            buf.push(c);
        } else {
            buf.push_str(transliterate_char(c as u32));
        }
    }
    buf
}

/// VBA `TransliterateChar`: one non-ASCII code point to its closest plain
/// ASCII letter(s), over Latin-1 Supplement and the common Latin Extended-A
/// letters; "" where the table has no mapping.
pub fn transliterate_char(code: u32) -> &'static str {
    match code {
        // Multi-letter expansions, checked before the 1:1 table.
        0xC6 | 0xE6 => "ae",
        0x152 | 0x153 => "oe",
        0xDE | 0xFE => "th",
        0xDF => "ss",

        0xC0 | 0xC1 | 0xC2 | 0xC3 | 0xC4 | 0xC5 | 0x100 | 0x102 | 0x104 => "A",
        0xE0 | 0xE1 | 0xE2 | 0xE3 | 0xE4 | 0xE5 | 0x101 | 0x103 | 0x105 => "a",
        0xC7 | 0x106 | 0x108 | 0x10A | 0x10C => "C",
        0xE7 | 0x107 | 0x109 | 0x10B | 0x10D => "c",
        0x10E | 0x110 => "D",
        0x10F | 0x111 => "d",
        0xC8 | 0xC9 | 0xCA | 0xCB | 0x112 | 0x114 | 0x116 | 0x118 | 0x11A => "E",
        0xE8 | 0xE9 | 0xEA | 0xEB | 0x113 | 0x115 | 0x117 | 0x119 | 0x11B => "e",
        0x11C | 0x11E | 0x120 | 0x122 => "G",
        0x11D | 0x11F | 0x121 | 0x123 => "g",
        0x124 | 0x126 => "H",
        0x125 | 0x127 => "h",
        0xCC | 0xCD | 0xCE | 0xCF | 0x128 | 0x12A | 0x12C | 0x12E | 0x130 => "I",
        0xEC | 0xED | 0xEE | 0xEF | 0x129 | 0x12B | 0x12D | 0x12F | 0x131 => "i",
        0x134 => "J",
        0x135 => "j",
        0x136 => "K",
        0x137 => "k",
        0x139 | 0x13B | 0x13D | 0x13F | 0x141 => "L",
        0x13A | 0x13C | 0x13E | 0x140 | 0x142 => "l",
        0xD1 | 0x143 | 0x145 | 0x147 => "N",
        0xF1 | 0x144 | 0x146 | 0x148 => "n",
        0xD2 | 0xD3 | 0xD4 | 0xD5 | 0xD6 | 0xD8 | 0x14C | 0x14E | 0x150 => "O",
        0xF2 | 0xF3 | 0xF4 | 0xF5 | 0xF6 | 0xF8 | 0x14D | 0x14F | 0x151 => "o",
        0x154 | 0x156 | 0x158 => "R",
        0x155 | 0x157 | 0x159 => "r",
        0x15A | 0x15C | 0x15E | 0x160 => "S",
        0x15B | 0x15D | 0x15F | 0x161 => "s",
        0x162 | 0x164 | 0x166 => "T",
        0x163 | 0x165 | 0x167 => "t",
        0xD9 | 0xDA | 0xDB | 0xDC | 0x168 | 0x16A | 0x16C | 0x16E | 0x170 | 0x172 => "U",
        0xF9 | 0xFA | 0xFB | 0xFC | 0x169 | 0x16B | 0x16D | 0x16F | 0x171 | 0x173 => "u",
        0x174 => "W",
        0x175 => "w",
        0xDD | 0x176 | 0x178 => "Y",
        0xFD | 0xFF | 0x177 => "y",
        0x179 | 0x17B | 0x17D => "Z",
        0x17A | 0x17C | 0x17E => "z",
        _ => "",
    }
}

/// VBA `ClickHandlerSlug`: letters, digits and `_` stay; every other
/// character becomes `_`.
pub fn click_handler_slug(caption: &str) -> String {
    caption
        .chars()
        .map(|c| {
            if c.is_ascii_alphanumeric() || c == '_' {
                c
            } else {
                '_'
            }
        })
        .collect()
}

/// VBA `ToVbaString` over a string literal's content: quoted, with each
/// inner quote mark doubled.
pub fn to_vba_string(content: &str) -> String {
    format!("\"{}\"", content.replace('"', "\"\""))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn sym_name_mangles_as_the_vba_does() {
        assert_eq!(sym_name("hot-pink").unwrap(), "hot_pink");
        assert_eq!(sym_name("on:click:go").unwrap(), "on_click_go");
        assert_eq!(sym_name("worksheet.cells").unwrap(), "worksheet.cells");
        assert_eq!(sym_name("0.08").unwrap(), "0.08");
        assert_eq!(sym_name("-5").unwrap(), "-5");
        assert_eq!(sym_name("1e-5").unwrap(), "1e-5");
        assert_eq!(sym_name("café").unwrap(), "cafe");
        assert_eq!(sym_name("Straße").unwrap(), "Strasse");
        assert_eq!(sym_name("Œuvre-Ñ").unwrap(), "oeuvre_N");
        let r = sym_name("Привет").unwrap_err();
        assert_eq!(r.id, "vla-name-no-plain-letters");
    }

    #[test]
    fn slugs_and_strings() {
        assert_eq!(click_handler_slug("Go now!"), "Go_now_");
        assert_eq!(click_handler_slug("a_b9"), "a_b9");
        assert_eq!(to_vba_string("say \"hi\""), "\"say \"\"hi\"\"\"");
        assert_eq!(to_vba_string(""), "\"\"");
    }
}
