//! The word tables and name lists of the English engine, read as data
//! (PORT.6, slice 6a).
//!
//! `VLA_English.bas` holds nine word tables in code and
//! `VLA_SentenceEngine.bas` six name lists; `tools/export_words.ps1` and
//! `tools/export_names.ps1` write them to `core/data/words.vla` and
//! `core/data/names.vla`, one form per entry in the VBA's own order, and
//! `tools/check_data_exports.ps1` fails when either has drifted from its
//! VBA. The VBA stays the source (SD-18); this module reads the files at
//! build time, so nothing is typed twice and the wasm import section stays
//! empty.
//!
//! One function per VBA table, named as the VBA names it, with the VBA's
//! own answer for a word that is not in it: `NumberWord` gives the word
//! back, `OrdinalWord` gives nothing, `SlotDesc` gives the category back,
//! `StrayCharHint` gives its `Case Else`. Comparison is exact, as a
//! `Select Case` on strings is; `IsReservedName` and `IsEngineCallName`
//! fold their argument first, as the VBA does.

use std::sync::OnceLock;

use crate::form::Form;
use crate::intrinsics::fold;
use crate::reader::read_forms;

const WORDS_TEXT: &str = include_str!("../../data/words.vla");
const NAMES_TEXT: &str = include_str!("../../data/names.vla");

/// One built-in function word: `RegisterBuiltinFuncWords`'s `AddFnEntry`.
/// `takes_of` is the `mFnOf` table (`length of x`); otherwise the word is
/// a value by itself (`mFnNullary`: `today`).
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct FunctionWord {
    pub word: String,
    pub target: String,
    pub takes_of: bool,
}

#[derive(Default)]
struct Tables {
    expr_ops: Vec<(String, String)>,
    expr_op_words: Vec<String>,
    noise_words: Vec<String>,
    dropped_words: Vec<String>,
    stray_hints: Vec<(String, String)>,
    stray_default: String,
    number_words: Vec<(String, String)>,
    ordinal_words: Vec<(String, String)>,
    color_words: Vec<String>,
    slot_descs: Vec<(String, String)>,
    reserved_names: Vec<String>,
    engine_call_names: Vec<String>,
    function_words: Vec<FunctionWord>,
    function_word_display: Vec<String>,
    conditions_grammar_words: Vec<String>,
    set_verbs: Vec<String>,
    after_value_words: Vec<String>,
    /// How many forms each file gave, for the floors the tests hold.
    words_count: usize,
    names_count: usize,
}

/// The string arguments of a form after its head, or `None` when any
/// argument is not a string literal.
fn strings(items: &[Form]) -> Option<Vec<String>> {
    items
        .iter()
        .skip(1)
        .map(|f| match f {
            Form::Str(s) => Some(s.clone()),
            _ => None,
        })
        .collect()
}

fn pair(items: &[Form]) -> Option<(String, String)> {
    let s = strings(items)?;
    if s.len() != 2 {
        return None;
    }
    Some((s[0].clone(), s[1].clone()))
}

fn one(items: &[Form]) -> Option<String> {
    let s = strings(items)?;
    if s.len() != 1 {
        return None;
    }
    Some(s[0].clone())
}

fn parse() -> Tables {
    let mut t = Tables::default();
    for (text, is_words) in [(WORDS_TEXT, true), (NAMES_TEXT, false)] {
        let text = text.replace("\r\n", "\n");
        let forms = match read_forms(&text) {
            Ok(forms) => forms,
            Err(_) => continue,
        };
        for form in &forms {
            let Form::List(l) = form else { continue };
            let Some(Form::Sym(head)) = l.items.first() else {
                continue;
            };
            let taken = match head.as_str() {
                "expr-op" => pair(&l.items).map(|p| t.expr_ops.push(p)),
                "expr-op-word" => one(&l.items).map(|w| t.expr_op_words.push(w)),
                "noise-word" => one(&l.items).map(|w| t.noise_words.push(w)),
                "dropped-word" => one(&l.items).map(|w| t.dropped_words.push(w)),
                "stray-char-hint" => pair(&l.items).map(|p| t.stray_hints.push(p)),
                "stray-char-default" => one(&l.items).map(|w| t.stray_default = w),
                "number-word" => pair(&l.items).map(|p| t.number_words.push(p)),
                "ordinal-word" => pair(&l.items).map(|p| t.ordinal_words.push(p)),
                "color-word" => one(&l.items).map(|w| t.color_words.push(w)),
                "slot-desc" => pair(&l.items).map(|p| t.slot_descs.push(p)),
                "reserved-name" => one(&l.items).map(|w| t.reserved_names.push(w)),
                "engine-call-name" => one(&l.items).map(|w| t.engine_call_names.push(w)),
                "function-word" => {
                    // (function-word of|nullary "word" "target")
                    let kind = match l.items.get(1) {
                        Some(Form::Sym(k)) => k.as_str(),
                        _ => continue,
                    };
                    let takes_of = match kind {
                        "of" => true,
                        "nullary" => false,
                        _ => continue,
                    };
                    pair(&l.items[1..]).map(|(word, target)| {
                        t.function_words.push(FunctionWord {
                            word,
                            target,
                            takes_of,
                        })
                    })
                }
                "function-word-display" => one(&l.items).map(|w| t.function_word_display.push(w)),
                "conditions-grammar-word" => {
                    one(&l.items).map(|w| t.conditions_grammar_words.push(w))
                }
                "set-verb" => one(&l.items).map(|w| t.set_verbs.push(w)),
                "after-value-word" => one(&l.items).map(|w| t.after_value_words.push(w)),
                _ => None,
            };
            if taken.is_some() {
                if is_words {
                    t.words_count += 1;
                } else {
                    t.names_count += 1;
                }
            }
        }
    }
    t
}

fn tables() -> &'static Tables {
    static TABLES: OnceLock<Tables> = OnceLock::new();
    TABLES.get_or_init(parse)
}

/// How many entries `core/data/words.vla` gave (the floor the tests hold).
pub fn words_count() -> usize {
    tables().words_count
}

/// How many entries `core/data/names.vla` gave.
pub fn names_count() -> usize {
    tables().names_count
}

/// VBA `ExprOpWord`: an operator symbol's words, or nothing.
pub fn expr_op_word(op: &str) -> String {
    tables()
        .expr_ops
        .iter()
        .find(|(k, _)| k == op)
        .map(|(_, v)| v.clone())
        .unwrap_or_default()
}

/// VBA `IsExprOpWord`: a word the expression parser treats as an operator.
pub fn is_expr_op_word(w: &str) -> bool {
    tables().expr_op_words.iter().any(|x| x == w)
}

/// VBA `IsNoiseWord`: stripped from patterns at registration.
pub fn is_noise_word(w: &str) -> bool {
    tables().noise_words.iter().any(|x| x == w)
}

/// VBA `IsDroppedWord`: dropped from input at tokenize time.
pub fn is_dropped_word(w: &str) -> bool {
    tables().dropped_words.iter().any(|x| x == w)
}

/// VBA `StrayCharHint`: ` - ` and the teaching half of the stray-character
/// refusal, its `Case Else` for a character the table does not name.
pub fn stray_char_hint(c: &str) -> String {
    let t = tables();
    let h = t
        .stray_hints
        .iter()
        .find(|(k, _)| k == c)
        .map(|(_, v)| v.as_str())
        .unwrap_or(t.stray_default.as_str());
    format!(" - {h}")
}

/// VBA `NumberWord`: `one` to `1`; any other word unchanged.
pub fn number_word(w: &str) -> String {
    tables()
        .number_words
        .iter()
        .find(|(k, _)| k == w)
        .map(|(_, v)| v.clone())
        .unwrap_or_else(|| w.to_string())
}

/// VBA `OrdinalWord`: `first` to `1`; nothing for any other word.
pub fn ordinal_word(w: &str) -> String {
    tables()
        .ordinal_words
        .iter()
        .find(|(k, _)| k == w)
        .map(|(_, v)| v.clone())
        .unwrap_or_default()
}

/// VBA `IsColorWord`.
pub fn is_color_word(s: &str) -> bool {
    tables().color_words.iter().any(|x| x == s)
}

/// VBA `SlotDesc`: what a slot category asks for, in words; the category
/// itself when the table has no entry.
pub fn slot_desc(cat: &str) -> String {
    tables()
        .slot_descs
        .iter()
        .find(|(k, _)| k == cat)
        .map(|(_, v)| v.clone())
        .unwrap_or_else(|| cat.to_string())
}

/// VBA `IsReservedName`: a VBA keyword, compared folded.
pub fn is_reserved_name(n: &str) -> bool {
    tables().reserved_names.contains(&fold(n))
}

/// VBA `IsEngineCallName` (U.30): a name the generated code calls by name,
/// compared folded.
pub fn is_engine_call_name(n: &str) -> bool {
    tables().engine_call_names.contains(&fold(n))
}

/// VBA `RegisterBuiltinFuncWords`: the engine's own function words, in the
/// order it registers them.
pub fn builtin_function_words() -> &'static [FunctionWord] {
    &tables().function_words
}

/// Its display strings, for the "What can I say?" listing.
pub fn function_word_display() -> &'static [String] {
    &tables().function_word_display
}

/// VBA `IsConditionsGrammarWord` (G-PROLOG): a word of the conditions
/// grammar, refused as a role or relation name.
pub fn is_conditions_grammar_word(nm: &str) -> bool {
    tables().conditions_grammar_words.iter().any(|x| x == nm)
}

/// VBA `IsSetVerb`.
pub fn is_set_verb(t: &str) -> bool {
    tables().set_verbs.iter().any(|x| x == t)
}

/// VBA `IsAfterValueWord` (LX.14): an operator word or one of the words
/// that may follow a value, where a phrase's hole ends.
pub fn is_after_value_word(w: &str) -> bool {
    is_expr_op_word(w) || tables().after_value_words.iter().any(|x| x == w)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn the_two_files_are_read_whole() {
        // 113 and 211 at the export of 2026-10-02 (VLA_English.bas at LX.14,
        // VLA_SentenceEngine.bas at LE.11), 218 names once U.31's seven
        // reserved words joined the same day; tools/check_data_exports.ps1
        // holds the same floors on the files.
        assert!(words_count() >= 113, "{} word entries read", words_count());
        assert!(names_count() >= 218, "{} name entries read", names_count());
        let t = tables();
        assert_eq!(t.expr_ops.len(), 11);
        assert_eq!(t.expr_op_words.len(), 7);
        assert_eq!(t.noise_words.len(), 4);
        assert_eq!(t.dropped_words.len(), 2);
        assert_eq!(t.stray_hints.len(), 16);
        assert_eq!(t.number_words.len(), 21);
        assert_eq!(t.ordinal_words.len(), 20);
        assert_eq!(t.color_words.len(), 8);
        assert_eq!(t.slot_descs.len(), 23);
        assert_eq!(t.engine_call_names.len(), 42);
        assert_eq!(t.function_words.len(), 16);
        assert_eq!(t.function_word_display.len(), 17);
        assert_eq!(t.conditions_grammar_words.len(), 8);
        assert_eq!(t.set_verbs.len(), 2);
        assert_eq!(t.after_value_words.len(), 12);
        assert!(t.reserved_names.len() >= 121);
        // U.31's seven, confirmed red in the VBA editor 2026-10-01.
        for w in [
            "return", "gosub", "global", "scale", "circle", "decimal", "longlong",
        ] {
            assert!(is_reserved_name(w), "{w} is not reserved");
        }
    }

    #[test]
    fn each_table_answers_as_the_vba_does() {
        assert_eq!(expr_op_word("<>"), "does not equal");
        assert_eq!(expr_op_word("^"), "");
        assert!(is_expr_op_word("divided") && !is_expr_op_word("by"));
        assert!(is_noise_word("a") && !is_dropped_word("a"));
        assert!(is_dropped_word("the") && is_dropped_word("please"));
        assert_eq!(stray_char_hint("+"), " - write 'plus' for addition");
        assert_eq!(
            stray_char_hint("\u{2014}"),
            " - use a plain hyphen - (Word may have autocorrected it)"
        );
        assert_eq!(
            stray_char_hint("\\"),
            " - if this is a file path, put the WHOLE path in \"quotes\" (a period ends a sentence otherwise)"
        );
        assert_eq!(
            stray_char_hint("@"),
            " - remove it, or put it inside \"quotes\" if it is part of a text value"
        );
        assert_eq!(number_word("twenty"), "20");
        assert_eq!(number_word("hundred"), "hundred");
        assert_eq!(number_word("One"), "One"); // the tokenizer folds first
        assert_eq!(ordinal_word("twelfth"), "12");
        assert_eq!(ordinal_word("zeroth"), "");
        assert!(is_color_word("magenta") && !is_color_word("pink"));
        assert_eq!(slot_desc("var"), "a name (one word, like total)");
        assert_eq!(slot_desc("nothing-of-the-kind"), "nothing-of-the-kind");
        assert!(is_reserved_name("Dim") && is_reserved_name("cvar"));
        assert!(!is_reserved_name("total"));
        assert!(is_engine_call_name("Range") && is_engine_call_name("make-button"));
        assert!(!is_engine_call_name("my-range"));
        let fw = builtin_function_words();
        assert_eq!(fw[0].word, "length");
        assert_eq!(fw[0].target, "len");
        assert!(fw[0].takes_of);
        assert_eq!(fw[14].word, "today");
        assert_eq!(fw[14].target, "date");
        assert!(!fw[14].takes_of);
        assert_eq!(function_word_display()[12], "item <n> of <list>");
        assert!(is_conditions_grammar_word("lists") && !is_conditions_grammar_word("has"));
        assert!(is_set_verb("are") && !is_set_verb("were"));
        assert!(is_after_value_word("percent") && is_after_value_word("plus"));
        assert!(!is_after_value_word("total"));
    }
}
