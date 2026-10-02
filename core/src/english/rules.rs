//! The rule store's registration grammar: VLA_SentenceEngine.bas's
//! `IsSlotTok`, `IsOptTok`, `IsAltCat`, the surface specs (`SurfaceMatch`,
//! `SurfaceForms`, `BareSurfaces`, `BareAltMatch`, `ValidateSurfaceSpec`),
//! `ValidateRuleItems`, `ExpandedSignatures` and `JoinSigs` (PORT.6, slice
//! 6c). Everything here is a pure function of a pattern's items; the store
//! itself, `AddPhraseRule` and `LintRule`, is `grammar.rs`.
//!
//! A pattern is split on spaces, each word folded, noise words dropped and
//! number words rewritten as the tokenizer rewrites them; what remains is
//! the rule's items: literals, slots (`{name}`, `{name:category}`,
//! `{name:category=default}`), optionals (`[word]`, `[a|b]`) and bare
//! alternations (`into|in`, `center/ed`). A category holding `|` or `/` is
//! an alternation of surfaces; a surface is a word or `stem/suffix`, which
//! matches the stem or the stem with its suffix and binds the stem.

use crate::intrinsics::fold;
use crate::messages::{raise, Refusal};

/// A slot's parts, as `IsSlotTok` hands them back.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Slot {
    pub name: String,
    pub cat: String,
    pub has_default: bool,
    pub default: String,
}

/// VBA `IsSlotTok`: `Some` for a well-formed slot, `None` for a literal, a
/// refusal with directions for a malformed one.
pub fn slot_tok(t: &str) -> Result<Option<Slot>, Refusal> {
    if !t.starts_with('{') {
        if let Some(rest) = t.strip_prefix('?') {
            return Err(raise(
                "english-old-slot-spelling",
                &[
                    ("t", t),
                    ("suggested", &format!("{{{rest}}}")),
                    ("example", "{name} (no category)"),
                ],
            ));
        }
        return Ok(None);
    }
    if !t.ends_with('}') || t.chars().count() < 3 {
        return Err(raise(
            "english-slot-not-closed",
            &[("t", t), ("ex1", "{name}"), ("ex2", "{name:category}")],
        ));
    }
    let inner = &t[1..t.len() - 1];
    let (name, mut cat) = match inner.find(':') {
        None => (inner.to_string(), "name".to_string()),
        Some(cp) => (inner[..cp].to_string(), inner[cp + 1..].to_string()),
    };
    if name.is_empty() || (inner.contains(':') && cat.is_empty()) {
        return Err(raise(
            "english-malformed-slot",
            &[("t", t), ("ex1", "{name}"), ("ex2", "{name:category}")],
        ));
    }
    // G7: an optional default, {name:category=text}; the last '=' wins.
    let mut has_default = false;
    let mut default = String::new();
    if let Some(eq) = cat.rfind('=') {
        has_default = true;
        default = cat[eq + 1..].to_string();
        cat = cat[..eq].to_string();
        if cat.is_empty() || default.is_empty() {
            return Err(raise(
                "english-malformed-default",
                &[("t", t), ("ex", "{name:category=text}")],
            ));
        }
    }
    Ok(Some(Slot {
        name,
        cat,
        has_default,
        default,
    }))
}

/// VBA `IsOptTok`: `[word]` or `[a|b]`, the inside folded; `None` for a
/// token that does not begin with `[`.
pub fn opt_tok(t: &str) -> Result<Option<String>, Refusal> {
    if !t.starts_with('[') {
        return Ok(None);
    }
    if !t.ends_with(']') || t.chars().count() < 3 {
        return Err(raise("english-optional-not-closed", &[("token", t)]));
    }
    Ok(Some(fold(&t[1..t.len() - 1])))
}

/// VBA `IsAltCat`: a category that is an alternation of surfaces.
pub fn is_alt_cat(cat: &str) -> bool {
    cat.contains('|') || cat.contains('/')
}

/// VBA `SurfaceMatch`: does a token match a surface spec, and what is its
/// stem? `center/ed` matches `center` and `centered`, binding `center`.
pub fn surface_match(spec: &str, tok: &str) -> (bool, String) {
    match spec.find('/') {
        None => (tok == spec, spec.to_string()),
        Some(sp) => {
            let stem = &spec[..sp];
            let suffix = &spec[sp + 1..];
            (
                tok == stem || tok == format!("{stem}{suffix}"),
                stem.to_string(),
            )
        }
    }
}

/// VBA `SurfaceForms`: the one or two surfaces a spec presents.
pub fn surface_forms(spec: &str) -> Vec<String> {
    match spec.find('/') {
        None => vec![spec.to_string()],
        Some(sp) => vec![
            spec[..sp].to_string(),
            format!("{}{}", &spec[..sp], &spec[sp + 1..]),
        ],
    }
}

/// VBA `BareSurfaces`: every surface of every branch of a bare spec.
pub fn bare_surfaces(spec: &str) -> Vec<String> {
    spec.split('|')
        .flat_map(|branch| surface_forms(&fold(branch)))
        .collect()
}

/// VBA `BareAltMatch`: a token against a bare surface spec, binding nothing.
pub fn bare_alt_match(spec: &str, tok: &str) -> bool {
    spec.split('|')
        .any(|branch| surface_match(&fold(branch), tok).0)
}

/// VBA `ValidateSurfaceSpec`: at most one slash, both sides non-empty.
pub fn validate_surface_spec(pattern: &str, where_: &str, spec: &str) -> Result<(), Refusal> {
    let Some(first) = spec.find('/') else {
        return Ok(());
    };
    if spec[first + 1..].contains('/') {
        return Err(raise(
            "english-surface-spec-multi-slash",
            &[("pattern", pattern), ("where", where_), ("spec", spec)],
        ));
    }
    if first == 0 || first == spec.len() - 1 {
        return Err(raise(
            "english-surface-spec-empty-side",
            &[("pattern", pattern), ("where", where_), ("spec", spec)],
        ));
    }
    Ok(())
}

/// VBA `ValidateRuleItems`: every slot's category is known or a
/// well-formed alternation, every optional and bare alternation has
/// non-empty branches, every surface spec stands.
pub fn validate_rule_items(pattern: &str, items: &[String]) -> Result<(), Refusal> {
    for t in items {
        if let Some(slot) = slot_tok(t)? {
            if is_alt_cat(&slot.cat) {
                for branch in slot.cat.split('|') {
                    if branch.trim_matches(' ').is_empty() {
                        return Err(raise(
                            "english-alternation-empty-branch",
                            &[
                                ("pattern", pattern),
                                ("token", t),
                                ("example", "{name:word|word}"),
                            ],
                        ));
                    }
                    validate_surface_spec(pattern, "branch", &fold(branch))?;
                }
            } else {
                match slot.cat.as_str() {
                    "name" | "var" | "text" | "expr" | "cond" | "conditions" | "role"
                    | "relation" | "clause" | "question" | "range" | "cell" | "column"
                    | "sheet" | "color" | "path" => {}
                    // G6: list-valued slots.
                    "text-list" | "range-list" | "cell-list" | "column-list" | "sheet-list"
                    | "color-list" => {}
                    other => {
                        return Err(raise(
                            "english-unknown-slot-category",
                            &[
                                ("pattern", pattern),
                                ("cat", other),
                                ("example", "{name:word|word}"),
                            ],
                        ))
                    }
                }
            }
        } else if let Some(ow) = opt_tok(t)? {
            for branch in ow.split('|') {
                if branch.trim_matches(' ').is_empty() {
                    return Err(raise(
                        "english-optional-empty-branch",
                        &[("pattern", pattern), ("token", t)],
                    ));
                }
                validate_surface_spec(pattern, "optional literal", &fold(branch))?;
            }
        } else if t.contains('|') || t.contains('/') {
            // G10: a bare surface token validates branch by branch.
            for branch in t.split('|') {
                if branch.trim_matches(' ').is_empty() {
                    return Err(raise(
                        "english-bare-alternation-empty-branch",
                        &[("pattern", pattern), ("token", t)],
                    ));
                }
                validate_surface_spec(pattern, "branch", &fold(branch))?;
            }
        }
    }
    Ok(())
}

/// VBA `ExpandedSignatures`: every shape a rule can present, one per
/// alternation branch and optional include/omit combination, literals as
/// themselves and category slots as `{category}`, pieces joined by `|`.
pub fn expanded_signatures(items: &[String]) -> Result<Vec<String>, Refusal> {
    let mut sigs: Vec<String> = vec![String::new()];
    for t in items {
        let mut pieces: Vec<String> = Vec::new();
        if let Some(slot) = slot_tok(t)? {
            if is_alt_cat(&slot.cat) {
                for branch in slot.cat.split('|') {
                    // G1.1: a stem/suffix branch is two shapes.
                    pieces.extend(surface_forms(&fold(branch)));
                }
            } else {
                pieces.push(format!("{{{}}}", slot.cat));
            }
            // G7: a default makes the slot omissible.
            if slot.has_default {
                pieces.push(String::new());
            }
        } else if let Some(ow) = opt_tok(t)? {
            pieces.push(String::new());
            pieces.extend(bare_surfaces(&ow));
        } else if t.contains('|') || t.contains('/') {
            pieces.extend(bare_surfaces(t));
        } else {
            pieces.push(t.clone());
        }
        let mut grown = Vec::with_capacity(sigs.len() * pieces.len());
        for s in &sigs {
            for pc in &pieces {
                if pc.is_empty() {
                    grown.push(s.clone());
                } else if s.is_empty() {
                    grown.push(pc.clone());
                } else {
                    grown.push(format!("{s}|{pc}"));
                }
            }
        }
        sigs = grown;
    }
    Ok(sigs)
}

/// VBA `JoinSigs`: the signatures as one text, one per line (LF).
pub fn join_sigs(sigs: &[String]) -> String {
    sigs.join("\n")
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn slots_parse_as_isslottok_does() {
        let s = slot_tok("{v:var}").unwrap().unwrap();
        assert_eq!(
            (s.name.as_str(), s.cat.as_str(), s.has_default),
            ("v", "var", false)
        );
        let s = slot_tok("{n}").unwrap().unwrap();
        assert_eq!((s.name.as_str(), s.cat.as_str()), ("n", "name"));
        let s = slot_tok("{k:sample|population}").unwrap().unwrap();
        assert!(is_alt_cat(&s.cat));
        let s = slot_tok("{r:text=\"a1\"}").unwrap().unwrap();
        assert_eq!(
            (s.cat.as_str(), s.default.as_str(), s.has_default),
            ("text", "\"a1\"", true)
        );
        assert!(slot_tok("plain").unwrap().is_none());
        assert_eq!(slot_tok("?x").unwrap_err().id, "english-old-slot-spelling");
        assert_eq!(slot_tok("{x").unwrap_err().id, "english-slot-not-closed");
        assert_eq!(slot_tok("{}").unwrap_err().id, "english-slot-not-closed");
        assert_eq!(
            slot_tok("{:expr}").unwrap_err().id,
            "english-malformed-slot"
        );
        assert_eq!(slot_tok("{x:}").unwrap_err().id, "english-malformed-slot");
        assert_eq!(
            slot_tok("{x:text=}").unwrap_err().id,
            "english-malformed-default"
        );
        assert_eq!(
            slot_tok("{x:=5}").unwrap_err().id,
            "english-malformed-default"
        );
        assert_eq!(opt_tok("[key]").unwrap().unwrap(), "key");
        assert_eq!(opt_tok("[A|B]").unwrap().unwrap(), "a|b");
        assert!(opt_tok("key").unwrap().is_none());
        assert_eq!(
            opt_tok("[key").unwrap_err().id,
            "english-optional-not-closed"
        );
    }

    #[test]
    fn surfaces_match_stems_and_suffixes() {
        assert_eq!(
            surface_match("center/ed", "centered"),
            (true, "center".to_string())
        );
        assert_eq!(
            surface_match("center/ed", "center"),
            (true, "center".to_string())
        );
        assert_eq!(
            surface_match("center/ed", "centre"),
            (false, "center".to_string())
        );
        assert_eq!(surface_match("left", "left"), (true, "left".to_string()));
        assert_eq!(surface_forms("sort/ed"), vec!["sort", "sorted"]);
        assert_eq!(
            bare_surfaces("left|center/ed"),
            vec!["left", "center", "centered"]
        );
        assert!(bare_alt_match("into|in", "in") && !bare_alt_match("into|in", "on"));
        assert!(validate_surface_spec("p", "branch", "a/b").is_ok());
        assert_eq!(
            validate_surface_spec("p", "branch", "a/b/c")
                .unwrap_err()
                .id,
            "english-surface-spec-multi-slash"
        );
        assert_eq!(
            validate_surface_spec("p", "branch", "/b").unwrap_err().id,
            "english-surface-spec-empty-side"
        );
        assert_eq!(
            validate_surface_spec("p", "branch", "a/").unwrap_err().id,
            "english-surface-spec-empty-side"
        );
    }

    fn items(s: &str) -> Vec<String> {
        s.split(' ').map(|w| w.to_string()).collect()
    }

    #[test]
    fn items_validate_and_expand_into_signatures() {
        assert!(validate_rule_items("p", &items("warm cell {r:cell}")).is_ok());
        assert_eq!(
            validate_rule_items("p", &items("warm {r:banana}"))
                .unwrap_err()
                .id,
            "english-unknown-slot-category"
        );
        assert_eq!(
            validate_rule_items("p", &items("x {d:left||right}"))
                .unwrap_err()
                .id,
            "english-alternation-empty-branch"
        );
        assert_eq!(
            validate_rule_items("p", &items("x [a|]")).unwrap_err().id,
            "english-optional-empty-branch"
        );
        assert_eq!(
            validate_rule_items("p", &items("x into|")).unwrap_err().id,
            "english-bare-alternation-empty-branch"
        );
        assert_eq!(
            expanded_signatures(&items("fix {w:cell|range} {r:text}")).unwrap(),
            vec!["fix|cell|{text}", "fix|range|{text}"]
        );
        assert_eq!(
            expanded_signatures(&items("warm cell {r:text} [up]")).unwrap(),
            vec!["warm|cell|{text}", "warm|cell|{text}|up"]
        );
        assert_eq!(
            expanded_signatures(&items("align {d:left|center/ed}")).unwrap(),
            vec!["align|left", "align|center", "align|centered"]
        );
        assert_eq!(
            expanded_signatures(&items("store {e:expr} into|in {r:text=\"a1\"}")).unwrap(),
            vec![
                "store|{expr}|into|{text}",
                "store|{expr}|into",
                "store|{expr}|in|{text}",
                "store|{expr}|in"
            ]
        );
        assert_eq!(join_sigs(&["a|b".to_string(), "c".to_string()]), "a|b\nc");
    }
}
