//! The datum: what the reader produces and every later stage walks.
//!
//! VLA.bas keeps a form as a `Variant`: a `String` for an atom, with a
//! leading quote character marking a string literal, or a `Collection` for
//! a list, whose source line is remembered beside it (`TagLine`/`FormLine`).
//! Here that is one enum. A string literal holds its content, not the
//! marker; [`Form::atom_text`] gives the VBA's own view back, the marker
//! included, for the few places that read it verbatim.
//!
//! Equality is structural and ignores a list's line, as the VBA's own
//! `DeepEqual` does: a line is where a form was read, not what it is.
//!
//! The helpers below are VLA.bas's datum helpers with their refusals: `Nth`,
//! `NthList`, `SymText`, `HeadSym`, `ListHeadIs`, `StrLitContent`, each
//! raising the same message id in the same situation.

use crate::intrinsics::fold;
use crate::messages::{raise, Refusal};

/// One form: a symbol (a number is a symbol until someone asks), a string
/// literal's content, or a list.
#[derive(Clone, Debug)]
pub enum Form {
    Sym(String),
    Str(String),
    List(List),
}

/// A list, with the source line of its opening parenthesis: 0 when it was
/// born in the prelude or made by a macro (S3.1's explicit zero).
#[derive(Clone, Debug, Default)]
pub struct List {
    pub items: Vec<Form>,
    pub line: u32,
}

impl PartialEq for Form {
    fn eq(&self, other: &Form) -> bool {
        match (self, other) {
            (Form::Sym(a), Form::Sym(b)) => a == b,
            (Form::Str(a), Form::Str(b)) => a == b,
            (Form::List(a), Form::List(b)) => a == b,
            _ => false,
        }
    }
}

impl PartialEq for List {
    fn eq(&self, other: &List) -> bool {
        self.items == other.items
    }
}

impl Form {
    pub fn sym(text: &str) -> Form {
        Form::Sym(text.to_string())
    }

    pub fn string(content: &str) -> Form {
        Form::Str(content.to_string())
    }

    /// A list with no line, as every list the macro system makes has none.
    pub fn list(items: Vec<Form>) -> Form {
        Form::List(List { items, line: 0 })
    }

    pub fn is_list(&self) -> bool {
        matches!(self, Form::List(_))
    }

    /// VBA `IsSym`: an atom that is not a string literal.
    pub fn is_sym(&self) -> bool {
        matches!(self, Form::Sym(_))
    }

    /// VBA `IsKeywordArg`: a symbol that starts with `:` and is more than
    /// the colon.
    pub fn is_keyword_arg(&self) -> bool {
        match self {
            Form::Sym(s) => s.starts_with(':') && s.len() > 1,
            _ => false,
        }
    }

    /// The VBA's `CStr(v)` of an atom: a symbol's text, or a string
    /// literal's content behind its quote marker. `None` for a list.
    pub fn atom_text(&self) -> Option<String> {
        match self {
            Form::Sym(s) => Some(s.clone()),
            Form::Str(s) => Some(format!("\"{s}")),
            Form::List(_) => None,
        }
    }

    pub fn as_list(&self) -> Option<&List> {
        match self {
            Form::List(l) => Some(l),
            _ => None,
        }
    }

    /// VBA `ListHeadIs`: a list whose first element is a symbol folding
    /// equal to `name`.
    pub fn head_is(&self, name: &str) -> bool {
        match self {
            Form::List(l) => l.head_is(name),
            _ => false,
        }
    }
}

impl List {
    pub fn new(items: Vec<Form>, line: u32) -> List {
        List { items, line }
    }

    pub fn count(&self) -> usize {
        self.items.len()
    }

    /// VBA `Nth`: element `i`, 1-based, refusing past the end.
    pub fn nth(&self, i: usize) -> Result<&Form, Refusal> {
        if i == 0 || i > self.items.len() {
            return Err(raise("vla-form-missing-element", &[("n", &i.to_string())]));
        }
        Ok(&self.items[i - 1])
    }

    /// VBA `NthList`: element `i` as a list, refusing an atom there.
    pub fn nth_list(&self, i: usize) -> Result<&List, Refusal> {
        match self.nth(i)? {
            Form::List(l) => Ok(l),
            other => Err(raise(
                "vla-expected-list-at-position",
                &[
                    ("n", &i.to_string()),
                    ("value", &other.atom_text().unwrap_or_default()),
                ],
            )),
        }
    }

    /// VBA `HeadSym`: the first element as a symbol's text.
    pub fn head_sym(&self) -> Result<String, Refusal> {
        if self.items.is_empty() {
            return Err(raise("vla-empty-form", &[]));
        }
        sym_text(&self.items[0])
    }

    /// VBA `ListHeadIs` over this list.
    pub fn head_is(&self, name: &str) -> bool {
        match self.items.first() {
            Some(Form::Sym(s)) => fold(s) == fold(name),
            _ => false,
        }
    }

    /// The list from its `i`th element on, 1-based: VBA's `VlaSlice` (cdr)
    /// as a copy, since nothing on this path is long enough to share.
    pub fn tail_from(&self, i: usize) -> List {
        let start = i.saturating_sub(1).min(self.items.len());
        List {
            items: self.items[start..].to_vec(),
            line: 0,
        }
    }
}

/// VBA `SymText`: the text of a symbol, refusing a list or a string.
pub fn sym_text(f: &Form) -> Result<String, Refusal> {
    match f {
        Form::Sym(s) => Ok(s.clone()),
        Form::Str(_) => Err(raise("vla-expected-symbol-got-string", &[])),
        Form::List(_) => Err(raise("vla-expected-symbol-got-list", &[])),
    }
}

/// VBA `StrLitContent`: the content of a string literal, refusing a list
/// or a symbol.
pub fn str_lit_content(f: &Form) -> Result<String, Refusal> {
    match f {
        Form::Str(s) => Ok(s.clone()),
        Form::Sym(s) => Err(raise("vla-expected-string-got-other", &[("value", s)])),
        Form::List(_) => Err(raise("vla-expected-string-got-list", &[])),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn equality_is_structural_and_ignores_lines() {
        let a = Form::List(List::new(vec![Form::sym("a"), Form::string("b")], 7));
        let b = Form::List(List::new(vec![Form::sym("a"), Form::string("b")], 0));
        assert_eq!(a, b);
        assert_ne!(Form::sym("x"), Form::string("x"));
        assert_ne!(Form::sym("x"), Form::list(vec![Form::sym("x")]));
    }

    #[test]
    fn atom_text_is_the_vba_view() {
        assert_eq!(Form::sym("x").atom_text().unwrap(), "x");
        assert_eq!(Form::string("x").atom_text().unwrap(), "\"x");
        assert!(Form::list(vec![]).atom_text().is_none());
        assert!(Form::sym(":key").is_keyword_arg());
        assert!(!Form::sym(":").is_keyword_arg());
        assert!(!Form::string(":key").is_keyword_arg());
    }

    #[test]
    fn helpers_refuse_with_the_catalogue_ids() {
        let l = List::new(vec![Form::sym("f"), Form::string("s")], 1);
        assert_eq!(l.head_sym().unwrap(), "f");
        assert!(l.head_is("F"));
        assert_eq!(l.nth(3).unwrap_err().id, "vla-form-missing-element");
        assert_eq!(
            l.nth(3).unwrap_err().text,
            "form is missing required element 3"
        );
        assert_eq!(
            l.nth_list(2).unwrap_err().id,
            "vla-expected-list-at-position"
        );
        assert_eq!(
            sym_text(&Form::string("s")).unwrap_err().id,
            "vla-expected-symbol-got-string"
        );
        assert_eq!(
            sym_text(&Form::list(vec![])).unwrap_err().id,
            "vla-expected-symbol-got-list"
        );
        assert_eq!(
            str_lit_content(&Form::sym("s")).unwrap_err().id,
            "vla-expected-string-got-other"
        );
        assert_eq!(List::default().head_sym().unwrap_err().id, "vla-empty-form");
        assert_eq!(l.tail_from(2).items, vec![Form::string("s")]);
        assert_eq!(l.tail_from(5).items, vec![]);
    }
}
