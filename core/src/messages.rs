//! The message catalogue and the refusal: INTRINSICS.md's sixth entry.
//!
//! Every refusal in the reference is `VLA_Messages.RaiseMsg "<id>", slot,
//! value, ...` (SD-2): the id names a catalogue entry holding the VBA error
//! number, the `Err.Source` tag and an English template whose `{slot}`s the
//! call fills. The treaty's fifth oracle holds every implementation to the
//! same id in the same situation, so this core refuses from the same
//! catalogue, read as data: `core/data/messages.vla`, exported once from
//! `src/VLA_Messages.bas` by `tools/export_messages.ps1` and held to it by
//! `tools/check_data_exports.ps1`. The VBA is the source until PORT.6 hands
//! the catalogue to the file.
//!
//! A refusal here is a value, not a raise: [`Refusal`], carried in a
//! `Result`. [`raise`] builds one exactly as `RaiseMsg` builds its text,
//! `SubstituteSlots` included. The two internal errors `RaiseMsg` can meet,
//! an unknown id and a slot left unfilled, are refusals too, with the VBA's
//! own words, so that nothing on this path can panic.

use std::sync::OnceLock;

use crate::form::Form;
use crate::reader::read_forms;

/// One refusal, finished: the id that named it, the VBA error number and
/// `Err.Source` the catalogue gives it, and the text with its slots filled.
#[derive(Clone, Debug, PartialEq)]
pub struct Refusal {
    pub id: String,
    pub number: i64,
    pub source: String,
    pub text: String,
}

impl Refusal {
    /// The reference appends to a caught refusal's text on its way out
    /// (`VlaTranspile`'s `emitfail`: ` (near vla line N)`); this is that.
    pub fn with_suffix(mut self, suffix: &str) -> Refusal {
        self.text.push_str(suffix);
        self
    }
}

impl std::fmt::Display for Refusal {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.write_str(&self.text)
    }
}

struct Entry {
    id: String,
    number: i64,
    source: String,
    template: String,
}

/// The exported catalogue, embedded at build time: no file is read when the
/// core runs (the wasm import section stays empty).
const CATALOGUE_TEXT: &str = include_str!("../data/messages.vla");

fn parse_catalogue() -> Vec<Entry> {
    // The file is CRLF on a Windows checkout and LF elsewhere; a template
    // that holds a line break must read the same on both.
    let text = CATALOGUE_TEXT.replace("\r\n", "\n");
    let forms = match read_forms(&text) {
        Ok(forms) => forms,
        Err(_) => return Vec::new(),
    };
    let mut entries = Vec::with_capacity(forms.len());
    for form in &forms {
        let Form::List(l) = form else { continue };
        if !l.head_is("message") || l.items.len() != 5 {
            continue;
        }
        let (Form::Sym(id), Form::Sym(number), Form::Str(source), Form::Str(template)) =
            (&l.items[1], &l.items[2], &l.items[3], &l.items[4])
        else {
            continue;
        };
        let Ok(number) = number.parse::<i64>() else {
            continue;
        };
        entries.push(Entry {
            id: id.clone(),
            number,
            source: source.clone(),
            template: template.clone(),
        });
    }
    entries
}

fn catalogue() -> &'static [Entry] {
    static CATALOGUE: OnceLock<Vec<Entry>> = OnceLock::new();
    CATALOGUE.get_or_init(parse_catalogue)
}

/// How many entries the catalogue holds.
pub fn count() -> usize {
    catalogue().len()
}

/// The template of an id, unfilled, or `None` for an id the catalogue
/// does not know.
pub fn template(id: &str) -> Option<&'static str> {
    catalogue()
        .iter()
        .find(|e| e.id == id)
        .map(|e| e.template.as_str())
}

/// Every id, in the order the VBA registers them.
pub fn ids() -> Vec<&'static str> {
    catalogue().iter().map(|e| e.id.as_str()).collect()
}

/// VBA `RaiseMsg`: the entry's number and source, and its template with
/// each `{slot}` replaced by the value given for that name.
pub fn raise(id: &str, slots: &[(&str, &str)]) -> Refusal {
    let Some(entry) = catalogue().iter().find(|e| e.id == id) else {
        return internal(id, &format!("RaiseMsg: unknown message id '{id}'"));
    };
    match substitute_slots(&entry.template, slots) {
        Ok(text) => Refusal {
            id: id.to_string(),
            number: entry.number,
            source: entry.source.clone(),
            text,
        },
        Err(slot) => internal(
            id,
            &format!("RaiseMsg: no value supplied for slot '{{{slot}}}'"),
        ),
    }
}

/// The reference's two internal errors are `Err.Raise 5, "VLA-Messages",
/// ...`; here they keep the id asked for, so a test can still see it.
fn internal(id: &str, text: &str) -> Refusal {
    Refusal {
        id: id.to_string(),
        number: 5,
        source: "VLA-Messages".to_string(),
        text: text.to_string(),
    }
}

/// VBA `SubstituteSlots`: scan for `{`, take the name up to the next `}`,
/// substitute its value; a `{` with no `}` after it ends the scan with the
/// rest copied as it stands. A name with no value is the error.
fn substitute_slots(template: &str, slots: &[(&str, &str)]) -> Result<String, String> {
    let mut out = String::with_capacity(template.len());
    let mut rest = template;
    loop {
        let Some(open) = rest.find('{') else {
            out.push_str(rest);
            return Ok(out);
        };
        let Some(close_rel) = rest[open..].find('}') else {
            out.push_str(rest);
            return Ok(out);
        };
        let close = open + close_rel;
        out.push_str(&rest[..open]);
        let name = &rest[open + 1..close];
        match slots.iter().find(|(k, _)| *k == name) {
            Some((_, v)) => out.push_str(v),
            None => return Err(name.to_string()),
        }
        rest = &rest[close + 1..];
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn the_catalogue_is_read_whole() {
        // 574 entries at the export of 2026-10-01 (VLA_Messages.bas, LX.14);
        // tools/check_data_exports.ps1 holds the same floor on the file.
        assert!(count() >= 574, "{} entries read", count());
        assert_eq!(ids()[0], "vla-source-not-found");
        assert_eq!(
            template("vla-unbalanced-parens").unwrap(),
            "unbalanced parentheses: missing ')'{loc}"
        );
        // A template exported with a line break inside it reads as one
        // string, with LF on every platform.
        assert!(template("english-test-failed")
            .unwrap()
            .contains("test FAILED\n  sentence: {sentence}"));
        // A named error number, resolved by the exporter from VLA.bas.
        let r = raise(
            "vla-interpreter-only-handler",
            &[("name", "on:sheet-change")],
        );
        assert_eq!(r.number, -2147221504 + 7001);
        assert_eq!(r.source, "VLA");
    }

    #[test]
    fn raise_fills_slots_as_raisemsg_does() {
        let r = raise("vla-form-missing-element", &[("n", "3")]);
        assert_eq!(r.id, "vla-form-missing-element");
        assert_eq!(r.number, 5);
        assert_eq!(r.source, "VLA");
        assert_eq!(r.text, "form is missing required element 3");
        let r = raise("vla-unbalanced-parens", &[("loc", "")]);
        assert_eq!(r.text, "unbalanced parentheses: missing ')'");
        let r = raise(
            "vla-expected-list-at-position",
            &[("n", "2"), ("value", "x")],
        );
        assert_eq!(r.text, "expected a list at position 2, got 'x'");
    }

    #[test]
    fn the_two_internal_errors_are_refusals_too() {
        let r = raise("no-such-id", &[]);
        assert_eq!(r.number, 5);
        assert_eq!(r.source, "VLA-Messages");
        assert_eq!(r.text, "RaiseMsg: unknown message id 'no-such-id'");
        let r = raise("vla-form-missing-element", &[]);
        assert_eq!(r.text, "RaiseMsg: no value supplied for slot '{n}'");
    }

    #[test]
    fn substitute_slots_copies_an_unclosed_brace() {
        assert_eq!(
            substitute_slots("a {x} b {", &[("x", "1")]).unwrap(),
            "a 1 b {"
        );
        assert_eq!(substitute_slots("no slots", &[]).unwrap(), "no slots");
        assert_eq!(
            substitute_slots("{a}{b}", &[("a", "1"), ("b", "2")]).unwrap(),
            "12"
        );
    }
}
