//! G-PROLOG's three built-in sub-grammars (PORT.6, slice 6e), ported from
//! VLA_SentenceEngine.bas arm for arm: `conditions` (`ParseConditions`,
//! `ParseOneRoleCondition`, `ParseOneTableRow`, the operands and constants),
//! `clause` (`ParseClause`, `ParseOtherwiseBranches`) and `question`
//! (`ParseQuestion` with the how-many, none, every, alone and list shapes).
//! SD-16 holds `conditions` regular: five fixed shapes joined by `and`, no
//! nesting, single terms as operands. Role nouns are the variables
//! (`RoleToVarName`), constants stand where roles do, and every generated
//! name is derivable by hand (`EscapeNamePart`, the `vla-` prefixes).
//!
//! A syntax failure returns `Ok(None)` and moves nothing, so the slot falls
//! through to the near-miss machinery; a semantic failure refuses by name.

use super::matcher::{is_str_tok, vla_string_lit, Parser};
use super::words::{is_conditions_grammar_word, is_set_verb};
use crate::intrinsics::fold;
use crate::messages::{raise, Refusal};

/// What one role condition says of itself, for the closure rules, the
/// left-recursion check and the first relation.
struct RoleCond {
    text: String,
    kind: &'static str,
    name: String,
    closure: bool,
}

/// One Table row condition, with what a negated row's projection needs.
#[derive(Default)]
struct TableRow {
    text: String,
    table: String,
    headers: String,
    args: String,
    proj_vars: String,
    proj_body: String,
}

/// VBA `RoleToVarName`: the first letter raised by code point, never a
/// locale's case (SD-8).
pub fn role_to_var_name(w: &str) -> String {
    let mut chars = w.chars();
    match chars.next() {
        Some(c) if c.is_ascii_lowercase() => c.to_ascii_uppercase().to_string() + chars.as_str(),
        Some(_) => w.to_string(),
        None => String::new(),
    }
}

/// VBA `EscapeNamePart`: a `-` inside a part doubled, so a single `-` only
/// ever separates parts.
pub fn escape_name_part(part: &str) -> String {
    part.replace('-', "--")
}

/// VBA `IsDigitRun`.
fn is_digit_run(s: &str) -> bool {
    !s.is_empty() && s.bytes().all(|b| b.is_ascii_digit())
}

/// VBA `IsInvariantNumeral`: an optional minus, digits, at most one point,
/// at least one digit.
pub fn is_invariant_numeral(s: &str) -> bool {
    let body = s.strip_prefix('-').unwrap_or(s);
    if body.is_empty() {
        return false;
    }
    let mut saw_digit = false;
    let mut saw_dot = false;
    for c in body.chars() {
        if c.is_ascii_digit() {
            saw_digit = true;
        } else if c == '.' && !saw_dot {
            saw_dot = true;
        } else {
            return false;
        }
    }
    saw_digit
}

/// VBA `IsCanonicalNumeral`: a numeral as a number would print it, no
/// leading zero and no trailing zero after the point.
pub fn is_canonical_numeral(s: &str) -> bool {
    let body = s.strip_prefix('-').unwrap_or(s);
    if body.is_empty() {
        return false;
    }
    let (int_part, frac_part) = match body.find('.') {
        Some(i) => (&body[..i], Some(&body[i + 1..])),
        None => (body, None),
    };
    if let Some(frac) = frac_part {
        if !is_digit_run(frac) || frac.ends_with('0') {
            return false;
        }
    }
    if !is_digit_run(int_part) {
        return false;
    }
    !(int_part.len() > 1 && int_part.starts_with('0'))
}

/// VBA `IsConstantTok`.
fn is_constant_tok(t: &str) -> bool {
    is_str_tok(t) || is_invariant_numeral(t)
}

/// VBA `ConstantAsWritten`.
fn constant_as_written(t: &str) -> String {
    if is_str_tok(t) {
        format!("\"{}\"", &t[1..])
    } else {
        t.to_string()
    }
}

fn note_role_bound(pos_roles: &mut Vec<String>, v: &str) {
    if !pos_roles.iter().any(|r| r == v) {
        pos_roles.push(v.to_string());
    }
}

fn note_name(names: &mut Vec<String>, nm: &str) {
    if !names.iter().any(|n| n == nm) {
        names.push(nm.to_string());
    }
}

/// VBA `RefuseGeneratedPrefix`: nothing a writer types may start with
/// `vla-`, where the generated rules live.
pub fn refuse_generated_prefix(nm: &str) -> Result<(), Refusal> {
    let Some(rest) = nm.strip_prefix("vla-") else {
        return Ok(());
    };
    let suggest = if rest.is_empty() { "can-cover" } else { rest };
    Err(raise(
        "english-conditions-reserved-prefix",
        &[("name", nm), ("suggest", suggest)],
    ))
}

/// VBA `RefuseGrammarWordAsName`.
fn refuse_grammar_word_as_name(nm: &str) -> Result<(), Refusal> {
    if is_conditions_grammar_word(nm) {
        return Err(raise("english-conditions-reserved-name", &[("name", nm)]));
    }
    refuse_generated_prefix(nm)
}

fn refuse_unbound_head_role(v: &str, bound: &[String]) -> Result<(), Refusal> {
    if bound.iter().any(|b| b == v) {
        return Ok(());
    }
    Err(raise(
        "english-clause-unbound-head-role",
        &[("role", &fold(v))],
    ))
}

/// VBA `ClosureName` / `ClosureRules`: the transitive closure of a relation
/// under `directly or not`.
fn closure_name(rel: &str) -> String {
    format!("vla-any-{}", escape_name_part(rel))
}

fn closure_rules(rel: &str) -> String {
    let nm = closure_name(rel);
    format!("(rule ({nm} X Y) ({rel} X Y)) (rule ({nm} X Y) ({rel} X Z) ({nm} Z Y)) ")
}

/// VBA `ShapeName`: a generated rule's name from its own content.
fn shape_name(prefix: &str, part1: &str, part2: &str) -> String {
    let mut nm = format!("vla-{prefix}-{}", escape_name_part(part1));
    if !part2.is_empty() {
        nm.push('-');
        nm.push_str(&escape_name_part(part2));
    }
    nm
}

fn first_guard_name(rel: &str, branch: usize) -> String {
    shape_name("first", rel, &branch.to_string())
}

fn refuse_list_shape(shape: &str) -> Refusal {
    raise("english-question-list-shape", &[("shape", shape)])
}

fn list_tail(rel: &str, v: &str, list_goal: &str) -> String {
    let nm = shape_name("list", rel, "");
    format!("(headless) (rule ({nm} {v}) (textjoin {v} \", \" {list_goal})) (query {nm})")
}

impl<'g> Parser<'g> {
    /// VBA `TakeConstant`: a quoted text (never a numeral in quotes) or an
    /// invariant numeral, as VLA text.
    fn take_constant(&self, p: &mut usize) -> Result<Option<String>, Refusal> {
        let t = self.tok(*p);
        if is_str_tok(t) {
            let content = &t[1..];
            if is_canonical_numeral(content) {
                return Err(raise(
                    "english-constant-quoted-number",
                    &[("text", content)],
                ));
            }
            let r = vla_string_lit(content);
            *p += 1;
            return Ok(Some(r));
        }
        if is_invariant_numeral(t) {
            let r = t.to_string();
            *p += 1;
            return Ok(Some(r));
        }
        Ok(None)
    }

    /// VBA `TakeRoleVar`: a bare word is a role, noted among all roles.
    fn take_role_var(&self, p: &mut usize, all_roles: &mut Vec<String>) -> Option<String> {
        let w = self.word_at(*p);
        if w.is_empty() {
            return None;
        }
        let v = role_to_var_name(&w);
        if !all_roles.iter().any(|r| r == &v) {
            all_roles.push(v.clone());
        }
        *p += 1;
        Some(v)
    }

    /// VBA `TakeOperand`: a role or a constant; the flag says which.
    fn take_operand(
        &self,
        p: &mut usize,
        all_roles: &mut Vec<String>,
    ) -> Result<Option<(String, bool)>, Refusal> {
        if let Some(v) = self.take_role_var(p, all_roles) {
            return Ok(Some((v, true)));
        }
        Ok(self.take_constant(p)?.map(|c| (c, false)))
    }

    /// VBA `TakeHeadOperand`: articles skipped, a role or a constant.
    fn take_head_operand(&self, p: &mut usize) -> Result<Option<(String, bool)>, Refusal> {
        self.skip_articles(p);
        let w = self.word_at(*p);
        if !w.is_empty() {
            *p += 1;
            return Ok(Some((role_to_var_name(&w), true)));
        }
        Ok(self.take_constant(p)?.map(|c| (c, false)))
    }

    fn at_directly_or_not(&self, p: usize) -> bool {
        self.tok(p) == "directly" && self.tok(p + 1) == "or" && self.tok(p + 2) == "not"
    }

    /// VBA `OtherwiseWidth`: `otherwise` or `, otherwise`, in tokens.
    fn otherwise_width(&self, p: usize) -> usize {
        if self.tok(p) == "otherwise" {
            1
        } else if self.tok(p) == "," && self.tok(p + 1) == "otherwise" {
            2
        } else {
            0
        }
    }

    /// VBA `TextTestAt`: `contains`, `starts with`, `ends with`.
    fn text_test_at(&self, p: usize) -> (&'static str, usize) {
        let t = self.tok(p);
        if t == "contains" {
            ("text-contains", 1)
        } else if self.tok(p + 1) == "with" {
            match t {
                "starts" => ("text-starts-with", 2),
                "ends" => ("text-ends-with", 2),
                _ => ("", 0),
            }
        } else {
            ("", 0)
        }
    }

    /// VBA `ParseTextTestRest`.
    fn parse_text_test_rest(
        &self,
        p: &mut usize,
        all_roles: &mut Vec<String>,
        test_name: &str,
        v1: &str,
    ) -> Result<Option<String>, Refusal> {
        let t = self.tok(*p);
        let v2 = if is_str_tok(t) {
            let r = vla_string_lit(&t[1..]);
            *p += 1;
            r
        } else {
            match self.take_operand(p, all_roles)? {
                Some((v2, _)) => v2,
                None => return Ok(None),
            }
        };
        if self.at_directly_or_not(*p) {
            return Err(raise(
                "english-conditions-closure-needs-relation",
                &[("shape", "a text test")],
            ));
        }
        Ok(Some(format!("({test_name} {v1} {v2})")))
    }

    /// VBA `ParseOneRoleCondition`: a text test, a relation, a comparison
    /// or a set membership, decided by the condition's second token.
    fn parse_one_role_condition(
        &self,
        p: &mut usize,
        negated: bool,
        pos_roles: &mut Vec<String>,
        all_roles: &mut Vec<String>,
        names_out: &mut Vec<String>,
    ) -> Result<Option<RoleCond>, Refusal> {
        let Some((v1, v1_is_role)) = self.take_operand(p, all_roles)? else {
            return Ok(None);
        };
        let (tt, tt_width) = self.text_test_at(*p);
        if !tt.is_empty() {
            if !v1_is_role {
                return Ok(None);
            }
            *p += tt_width;
            return Ok(self
                .parse_text_test_rest(p, all_roles, tt, &v1)?
                .map(|text| RoleCond {
                    text,
                    kind: "text",
                    name: String::new(),
                    closure: false,
                }));
        }
        if self.tok(*p) != "is" {
            let nm = self.word_at(*p);
            if nm.is_empty() {
                return Ok(None);
            }
            refuse_grammar_word_as_name(&nm)?;
            *p += 1;
            let Some((v2, v2_is_role)) = self.take_operand(p, all_roles)? else {
                return Ok(None);
            };
            if !negated {
                if v1_is_role {
                    note_role_bound(pos_roles, &v1);
                }
                if v2_is_role {
                    note_role_bound(pos_roles, &v2);
                }
            }
            note_name(names_out, &nm);
            if self.at_directly_or_not(*p) {
                *p += 3;
                return Ok(Some(RoleCond {
                    text: format!("({} {v1} {v2})", closure_name(&nm)),
                    kind: "relation",
                    name: nm,
                    closure: true,
                }));
            }
            return Ok(Some(RoleCond {
                text: format!("({nm} {v1} {v2})"),
                kind: "relation",
                name: nm,
                closure: false,
            }));
        }
        if !v1_is_role {
            return Ok(None);
        }
        *p += 1; // "is"
        let a = self.tok(*p).to_string();
        let b = self.tok(*p + 1).to_string();
        let mut op = "";
        let mut swapped = false;
        if a == "at" && b == "least" {
            op = ">=";
            *p += 2;
        } else if a == "at" && b == "most" {
            op = ">=";
            swapped = true;
            *p += 2;
        } else if a == "greater" && b == "than" {
            op = ">";
            *p += 2;
        } else if a == "less" && b == "than" {
            op = "<";
            *p += 2;
        }
        if !op.is_empty() {
            let t = self.tok(*p).to_string();
            if is_str_tok(&t) && !is_canonical_numeral(&t[1..]) {
                return Err(raise(
                    "english-conditions-compare-needs-number",
                    &[("role", &fold(&v1)), ("text", &t[1..])],
                ));
            }
            let Some((v2, _)) = self.take_operand(p, all_roles)? else {
                return Ok(None);
            };
            if self.at_directly_or_not(*p) {
                return Err(raise(
                    "english-conditions-closure-needs-relation",
                    &[("shape", "a comparison")],
                ));
            }
            let text = if swapped {
                format!("({op} {v2} {v1})")
            } else {
                format!("({op} {v1} {v2})")
            };
            return Ok(Some(RoleCond {
                text,
                kind: "comparison",
                name: String::new(),
                closure: false,
            }));
        }
        let t = self.tok(*p).to_string();
        if is_constant_tok(&t) {
            return Err(raise(
                "english-conditions-equality-constant",
                &[("role", &fold(&v1)), ("value", &constant_as_written(&t))],
            ));
        }
        let nm = self.word_at(*p);
        if nm.is_empty() {
            return Ok(None);
        }
        refuse_grammar_word_as_name(&nm)?;
        *p += 1;
        if self.at_directly_or_not(*p) {
            return Err(raise(
                "english-conditions-closure-needs-relation",
                &[("shape", "a set")],
            ));
        }
        if !negated {
            note_role_bound(pos_roles, &v1);
        }
        note_name(names_out, &nm);
        Ok(Some(RoleCond {
            text: format!("({nm} {v1})"),
            kind: "set",
            name: nm,
            closure: false,
        }))
    }

    /// VBA `ParseOneTableRow`: `<table> lists <operand> as <header>, ...`.
    fn parse_one_table_row(
        &self,
        p: &mut usize,
        negated: bool,
        pos_roles: &mut Vec<String>,
        all_roles: &mut Vec<String>,
    ) -> Result<Option<TableRow>, Refusal> {
        let tbl = self.word_at(*p);
        if tbl.is_empty() {
            return Ok(None);
        }
        *p += 1;
        if self.tok(*p) != "lists" {
            return Ok(None);
        }
        *p += 1;
        let mut pairs = String::new();
        let mut headers = String::new();
        let mut args = String::new();
        let mut header_vars = String::new();
        let mut header_pairs = String::new();
        let mut any_constant = false;
        loop {
            let Some((term, is_role)) = self.take_operand(p, all_roles)? else {
                return Ok(None);
            };
            if self.tok(*p) != "as" {
                return Ok(None);
            }
            *p += 1;
            let hdr = self.word_at(*p);
            if hdr.is_empty() {
                return Ok(None);
            }
            *p += 1;
            push_sep(&mut pairs, ' ', &format!("({hdr} {term})"));
            push_sep(&mut headers, '-', &escape_name_part(&hdr));
            push_sep(&mut args, ' ', &term);
            push_sep(&mut header_vars, ' ', &role_to_var_name(&hdr));
            push_sep(
                &mut header_pairs,
                ' ',
                &format!("({hdr} {})", role_to_var_name(&hdr)),
            );
            if is_role {
                if !negated {
                    note_role_bound(pos_roles, &term);
                }
            } else {
                any_constant = true;
            }
            let mut q = *p;
            if self.tok(q) == "," {
                q += 1;
            }
            if self.tok(q) == "and" {
                q += 1;
            }
            if q == *p {
                break;
            }
            let window_ok = (!self.word_at(q).is_empty() || is_constant_tok(self.tok(q)))
                && self.tok(q + 1) == "as";
            if !window_ok {
                break;
            }
            *p = q;
        }
        let text = format!("({tbl} {pairs})");
        let (proj_vars, proj_body) = if any_constant {
            (header_vars, format!("({tbl} {header_pairs})"))
        } else {
            (args.clone(), text.clone())
        };
        Ok(Some(TableRow {
            text,
            table: tbl,
            headers,
            args,
            proj_vars,
            proj_body,
        }))
    }

    /// VBA `ParseConditions`: the goals, with the projection and closure
    /// rules they need in `pre_rules`; the relation and set names read, the
    /// roles bound positively, the tables read, and the first condition's
    /// relation when it is positive.
    #[allow(clippy::too_many_arguments)]
    pub(super) fn parse_conditions(
        &self,
        pos: &mut usize,
        pre_rules: &mut String,
        names_out: &mut Vec<String>,
        bound_out: &mut Vec<String>,
        tables_out: &mut Vec<String>,
        first_rel: &mut String,
        pre_seen: &mut Vec<String>,
    ) -> Result<Option<String>, Refusal> {
        let mut goals = String::new();
        let mut pre = String::new();
        let mut pos_roles: Vec<String> = Vec::new();
        let mut all_roles: Vec<String> = Vec::new();
        let mut p = *pos;
        let mut cond_index = 0;
        pre_rules.clear();
        first_rel.clear();
        loop {
            cond_index += 1;
            let mut negated = false;
            if self.tok(p) == "not" {
                negated = true;
                p += 1;
            }
            let mut goal;
            let mut row: Option<TableRow> = None;
            if self.tok(p + 1) == "lists" {
                let Some(r) =
                    self.parse_one_table_row(&mut p, negated, &mut pos_roles, &mut all_roles)?
                else {
                    return Ok(None);
                };
                goal = r.text.clone();
                if self.at_directly_or_not(p) {
                    return Err(raise(
                        "english-conditions-closure-needs-relation",
                        &[("shape", "a table row")],
                    ));
                }
                note_name(tables_out, &r.table);
                row = Some(r);
            } else {
                let Some(c) = self.parse_one_role_condition(
                    &mut p,
                    negated,
                    &mut pos_roles,
                    &mut all_roles,
                    names_out,
                )?
                else {
                    return Ok(None);
                };
                goal = c.text;
                if c.closure {
                    let cn = closure_name(&c.name);
                    if !pre_seen.contains(&cn) {
                        pre_seen.push(cn);
                        pre.push_str(&closure_rules(&c.name));
                    }
                }
                if cond_index == 1 && !negated && (c.kind == "relation" || c.kind == "set") {
                    *first_rel = c.name.clone();
                }
            }
            if negated {
                if let Some(r) = &row {
                    let proj_name = format!("vla-not-{}-{}", escape_name_part(&r.table), r.headers);
                    if !pre_seen.contains(&proj_name) {
                        pre_seen.push(proj_name.clone());
                        pre.push_str(&format!(
                            "(rule ({proj_name} {}) {}) ",
                            r.proj_vars, r.proj_body
                        ));
                    }
                    goal = format!("({proj_name} {})", r.args);
                }
                goal = format!("(not {goal})");
            }
            push_sep(&mut goals, ' ', &goal);
            if self.otherwise_width(p) > 0 {
                break;
            }
            let mut q = p;
            let mut saw_sep = false;
            if self.tok(q) == "," {
                q += 1;
                saw_sep = true;
            }
            if self.tok(q) == "and" {
                q += 1;
                saw_sep = true;
            }
            if !saw_sep {
                break;
            }
            p = q;
        }
        for rv in &all_roles {
            if !pos_roles.contains(rv) {
                return Err(raise(
                    "english-conditions-unbound-role",
                    &[("role", &fold(rv))],
                ));
            }
        }
        for rv in &pos_roles {
            if !bound_out.contains(rv) {
                bound_out.push(rv.clone());
            }
        }
        *pos = p;
        *pre_rules = pre;
        Ok(Some(goals))
    }

    /// VBA `ParseClause`: a rule or a fact, whole; `first_match` says the
    /// rule had otherwise branches.
    pub(super) fn parse_clause(
        &self,
        pos: &mut usize,
        head_relation: &mut String,
        body_names: &mut Vec<String>,
        body_tables: &mut Vec<String>,
        first_match: &mut bool,
    ) -> Result<Option<String>, Refusal> {
        let mut p = *pos;
        *first_match = false;
        let Some((subj, subj_role)) = self.take_head_operand(&mut p)? else {
            return Ok(None);
        };
        let rel;
        let head_text;
        let mut obj = String::new();
        let mut obj_role = false;
        let is_set_head;
        if self.tok(p) == "is" {
            p += 1;
            self.skip_articles(&mut p);
            rel = self.word_at(p);
            if rel.is_empty() {
                return Ok(None);
            }
            refuse_grammar_word_as_name(&rel)?;
            p += 1;
            head_text = format!("({rel} {subj})");
            is_set_head = true;
        } else {
            rel = self.word_at(p);
            if rel.is_empty() {
                return Ok(None);
            }
            refuse_grammar_word_as_name(&rel)?;
            p += 1;
            let Some((o, o_role)) = self.take_head_operand(&mut p)? else {
                return Ok(None);
            };
            obj = o;
            obj_role = o_role;
            head_text = format!("({rel} {subj} {obj})");
            is_set_head = false;
        }
        if self.at_directly_or_not(p) {
            return Err(raise(
                "english-clause-closure-in-head",
                &[("relation", &rel)],
            ));
        }
        let text;
        if self.tok(p) == "if" {
            p += 1;
            let mut bound: Vec<String> = Vec::new();
            let mut pre_seen: Vec<String> = Vec::new();
            let mut pre = String::new();
            let mut first_rel = String::new();
            let Some(goals) = self.parse_conditions(
                &mut p,
                &mut pre,
                body_names,
                &mut bound,
                body_tables,
                &mut first_rel,
                &mut pre_seen,
            )?
            else {
                return Ok(None);
            };
            if subj_role {
                refuse_unbound_head_role(&subj, &bound)?;
            }
            if obj_role {
                refuse_unbound_head_role(&obj, &bound)?;
            }
            if body_tables.contains(&rel) {
                return Err(raise(
                    "english-clause-relation-names-table",
                    &[("relation", &rel)],
                ));
            }
            if self.otherwise_width(p) > 0 {
                let Some(t) = self.parse_otherwise_branches(
                    &mut p,
                    &rel,
                    &subj,
                    subj_role,
                    is_set_head,
                    &obj,
                    &goals,
                    &pre,
                    &mut pre_seen,
                    body_names,
                    body_tables,
                )?
                else {
                    return Ok(None);
                };
                text = t;
                *first_match = true;
            } else {
                if first_rel == rel {
                    return Err(raise(
                        "english-clause-left-recursion",
                        &[("relation", &rel)],
                    ));
                }
                text = format!("{pre}(rule {head_text} {goals})");
            }
        } else {
            if subj_role {
                return Err(raise(
                    "english-clause-fact-needs-values",
                    &[("role", &fold(&subj))],
                ));
            }
            if obj_role {
                return Err(raise(
                    "english-clause-fact-needs-values",
                    &[("role", &fold(&obj))],
                ));
            }
            text = format!("(fact {head_text})");
        }
        *head_relation = rel;
        *pos = p;
        Ok(Some(text))
    }

    /// VBA `ParseOtherwiseBranches`: first-match branches, each guarded by
    /// the negation of every earlier branch's guard rule.
    #[allow(clippy::too_many_arguments)]
    fn parse_otherwise_branches(
        &self,
        p: &mut usize,
        rel: &str,
        subj: &str,
        subj_role: bool,
        is_set_head: bool,
        first_obj: &str,
        first_goals: &str,
        pre_in: &str,
        pre_seen: &mut Vec<String>,
        body_names: &mut Vec<String>,
        body_tables: &mut Vec<String>,
    ) -> Result<Option<String>, Refusal> {
        if is_set_head || !subj_role {
            return Err(raise(
                "english-clause-otherwise-shape",
                &[("relation", rel)],
            ));
        }
        let mut pre = pre_in.to_string();
        let mut objs: Vec<String> = vec![first_obj.to_string()];
        let mut goals_of: Vec<String> = vec![first_goals.to_string()];
        while self.otherwise_width(*p) > 0 {
            *p += self.otherwise_width(*p);
            let Some((obj, obj_role)) = self.take_head_operand(p)? else {
                return Ok(None);
            };
            if self.tok(*p) != "if" {
                return Err(raise(
                    "english-clause-otherwise-needs-if",
                    &[("relation", rel)],
                ));
            }
            *p += 1;
            let mut bound: Vec<String> = Vec::new();
            let mut bpre = String::new();
            let mut first_rel = String::new();
            let Some(goals) = self.parse_conditions(
                p,
                &mut bpre,
                body_names,
                &mut bound,
                body_tables,
                &mut first_rel,
                pre_seen,
            )?
            else {
                return Ok(None);
            };
            pre.push_str(&bpre);
            refuse_unbound_head_role(subj, &bound)?;
            if obj_role {
                refuse_unbound_head_role(&obj, &bound)?;
            }
            objs.push(obj);
            goals_of.push(goals);
        }
        if body_names.iter().any(|n| n == rel) {
            return Err(raise(
                "english-clause-otherwise-reads-itself",
                &[("relation", rel)],
            ));
        }
        if body_tables.iter().any(|t| t == rel) {
            return Err(raise(
                "english-clause-relation-names-table",
                &[("relation", rel)],
            ));
        }
        let mut guards = String::new();
        let mut rules = String::new();
        for (k0, goals) in goals_of.iter().enumerate() {
            let k = k0 + 1;
            if k < goals_of.len() {
                guards.push_str(&format!(
                    "(rule ({} {subj}) {goals}) ",
                    first_guard_name(rel, k)
                ));
            }
            let mut body = goals.clone();
            for j in 1..k {
                body.push_str(&format!(" (not ({} {subj}))", first_guard_name(rel, j)));
            }
            if !rules.is_empty() {
                rules.push(' ');
            }
            rules.push_str(&format!("(rule ({rel} {subj} {}) {body})", objs[k0]));
        }
        Ok(Some(format!("{pre}{guards}{rules}")))
    }

    /// VBA `TakeUnknown`: `who`, `what`, or `which <noun>`.
    fn take_unknown(&self, p: &mut usize) -> Option<String> {
        match self.tok(*p) {
            "who" => {
                *p += 1;
                Some("Who".to_string())
            }
            "what" => {
                *p += 1;
                Some("What".to_string())
            }
            "which" => {
                let noun = self.word_at(*p + 1);
                if noun.is_empty() || is_conditions_grammar_word(&noun) {
                    return None;
                }
                *p += 2;
                Some(role_to_var_name(&noun))
            }
            _ => None,
        }
    }

    fn at_alone_shape(&self, p: usize) -> bool {
        self.tok(p) == "alone" && !self.word_at(p + 1).is_empty()
    }

    fn take_shape_noun(&self, p: &mut usize) -> Option<String> {
        let w = self.word_at(*p);
        if w.is_empty() || is_conditions_grammar_word(&w) {
            return None;
        }
        *p += 1;
        Some(w)
    }

    fn take_set_after_verb(&self, p: &mut usize) -> Result<Option<String>, Refusal> {
        if !is_set_verb(self.tok(*p)) {
            return Ok(None);
        }
        let mut q = *p + 1;
        self.skip_articles(&mut q);
        let nm = self.word_at(q);
        if nm.is_empty() {
            return Ok(None);
        }
        refuse_grammar_word_as_name(&nm)?;
        *p = q + 1;
        Ok(Some(nm))
    }

    fn refuse_set_closure(&self, p: usize) -> Result<(), Refusal> {
        if self.at_directly_or_not(p) {
            return Err(raise(
                "english-conditions-closure-needs-relation",
                &[("shape", "a set")],
            ));
        }
        Ok(())
    }

    fn at_one_list_shape(&self, p: usize) -> bool {
        self.tok(p) == "as" && self.tok(p + 1) == "1" && self.tok(p + 2) == "list"
    }

    fn list_shape_kind(&self, p: usize) -> &'static str {
        if self.tok(p) == "how" {
            "a how many question"
        } else if self.tok(p) == "whether" && self.tok(p + 1) == "every" {
            "a whether every question"
        } else if self.tok(p) == "which" && self.tok(p + 2) == "that" {
            "a which ... is not question"
        } else {
            "an alone question"
        }
    }

    /// VBA `ParseHowManyQuestion` (G-PROLOG slice 4).
    fn parse_how_many_question(
        &self,
        p: &mut usize,
        relation: &mut String,
        relation2: &mut String,
    ) -> Result<Option<String>, Refusal> {
        let Some(noun) = self.take_shape_noun(p) else {
            return Ok(None);
        };
        let nv = role_to_var_name(&noun);
        if is_constant_tok(self.tok(*p)) {
            let Some(c) = self.take_constant(p)? else {
                return Ok(None);
            };
            let rel = self.word_at(*p);
            if rel.is_empty() {
                return Ok(None);
            }
            refuse_grammar_word_as_name(&rel)?;
            *p += 1;
            let mut gp = rel.clone();
            let mut pre = String::new();
            if self.at_directly_or_not(*p) {
                *p += 3;
                gp = closure_name(&rel);
                pre = closure_rules(&rel);
            }
            let nm = shape_name("count", &rel, "");
            *relation = rel;
            return Ok(Some(format!(
                "{pre}(headless) (rule ({nm} {nv}) (count {nv} ({gp} {c} VlaCounted))) (query {nm})"
            )));
        }
        if is_set_verb(self.tok(*p)) {
            return Ok(None);
        }
        let rel = self.word_at(*p);
        if rel.is_empty() {
            return Ok(None);
        }
        refuse_grammar_word_as_name(&rel)?;
        *p += 1;
        if self.tok(*p) == "each" {
            *p += 1;
            let Some(noun2) = self.take_shape_noun(p) else {
                return Ok(None);
            };
            let v2 = role_to_var_name(&noun2);
            if self.tok(*p) != "that" {
                return Ok(None);
            }
            *p += 1;
            let Some(set_name) = self.take_set_after_verb(p)? else {
                return Ok(None);
            };
            self.refuse_set_closure(*p)?;
            if fold(&v2) == fold(&nv) {
                return Err(raise(
                    "english-question-same-unknown",
                    &[("unknown", &fold(&noun))],
                ));
            }
            let nm = shape_name("each", &rel, "");
            *relation = rel.clone();
            *relation2 = set_name.clone();
            return Ok(Some(format!(
                "(rule ({nm} {v2} {nv}) ({set_name} {v2}) (count {nv} ({rel} VlaCounted {v2}))) (query {nm})"
            )));
        }
        let Some(c) = self.take_constant(p)? else {
            return Ok(None);
        };
        let mut gp = rel.clone();
        let mut pre = String::new();
        if self.at_directly_or_not(*p) {
            *p += 3;
            gp = closure_name(&rel);
            pre = closure_rules(&rel);
        }
        let nm = shape_name("count", &rel, "");
        *relation = rel;
        Ok(Some(format!(
            "{pre}(headless) (rule ({nm} {nv}) (count {nv} ({gp} VlaCounted {c}))) (query {nm})"
        )))
    }

    /// VBA `ParseNoneQuestion`: `which <noun> that is <set> is not <set>`.
    fn parse_none_question(
        &self,
        p: &mut usize,
        relation: &mut String,
        relation2: &mut String,
    ) -> Result<Option<String>, Refusal> {
        let Some(noun) = self.take_shape_noun(p) else {
            return Ok(None);
        };
        let v = role_to_var_name(&noun);
        if self.tok(*p) != "that" {
            return Ok(None);
        }
        *p += 1;
        let Some(set1) = self.take_set_after_verb(p)? else {
            return Ok(None);
        };
        self.refuse_set_closure(*p)?;
        if !is_set_verb(self.tok(*p)) || self.tok(*p + 1) != "not" {
            return Ok(None);
        }
        *p += 2;
        self.skip_articles(p);
        let set2 = self.word_at(*p);
        if set2.is_empty() {
            return Ok(None);
        }
        refuse_grammar_word_as_name(&set2)?;
        *p += 1;
        self.refuse_set_closure(*p)?;
        let nm = shape_name("none", &set1, &set2);
        *relation = set1.clone();
        *relation2 = set2.clone();
        Ok(Some(format!(
            "(rule ({nm} {v}) ({set1} {v}) (not ({set2} {v}))) (query {nm})"
        )))
    }

    /// VBA `ParseEveryQuestion`: `whether every <noun> that is <set> is <set>`.
    fn parse_every_question(
        &self,
        p: &mut usize,
        relation: &mut String,
        relation2: &mut String,
    ) -> Result<Option<String>, Refusal> {
        let Some(noun) = self.take_shape_noun(p) else {
            return Ok(None);
        };
        let v = role_to_var_name(&noun);
        if self.tok(*p) != "that" {
            return Ok(None);
        }
        *p += 1;
        let Some(set1) = self.take_set_after_verb(p)? else {
            return Ok(None);
        };
        self.refuse_set_closure(*p)?;
        if is_set_verb(self.tok(*p)) && self.tok(*p + 1) == "not" {
            return Ok(None);
        }
        let Some(set2) = self.take_set_after_verb(p)? else {
            return Ok(None);
        };
        self.refuse_set_closure(*p)?;
        let nm = shape_name("none", &set1, &set2);
        *relation = set1.clone();
        *relation2 = set2.clone();
        Ok(Some(format!(
            "(rule ({nm} {v}) ({set1} {v}) (not ({set2} {v}))) (query (not ({nm} {v})))"
        )))
    }

    /// VBA `ParseAloneQuestion`: `who alone <rel> <constant>` or `whether
    /// <constant> alone <rel> <constant>`.
    fn parse_alone_question(
        &self,
        p: &mut usize,
        head_var: &str,
        ask_const: &str,
        relation: &mut String,
    ) -> Result<Option<String>, Refusal> {
        let rel = self.word_at(*p);
        if rel.is_empty() {
            return Ok(None);
        }
        refuse_grammar_word_as_name(&rel)?;
        *p += 1;
        let Some(c) = self.take_constant(p)? else {
            return Ok(None);
        };
        let mut gp = rel.clone();
        let mut pre = String::new();
        if self.at_directly_or_not(*p) {
            *p += 3;
            gp = closure_name(&rel);
            pre = closure_rules(&rel);
        }
        let nm = shape_name("alone", &rel, "");
        let (hv, ask_text) = if !ask_const.is_empty() {
            ("Who".to_string(), format!("(query ({nm} {ask_const}))"))
        } else {
            (head_var.to_string(), format!("(query {nm})"))
        };
        *relation = rel;
        Ok(Some(format!(
            "{pre}(rule ({nm} {hv}) ({gp} {hv} {c}) (count VlaCount ({gp} VlaCounted {c})) (= VlaCount 1)) {ask_text}"
        )))
    }

    /// VBA `ParseEachListQuestion`: `who <rel> each <noun> that is <set> as
    /// one list`.
    fn parse_each_list_question(
        &self,
        p: &mut usize,
        rel: &str,
        head_var: &str,
        relation2: &mut String,
    ) -> Result<Option<String>, Refusal> {
        let Some(noun2) = self.take_shape_noun(p) else {
            return Ok(None);
        };
        let v2 = role_to_var_name(&noun2);
        if self.tok(*p) != "that" {
            return Ok(None);
        }
        *p += 1;
        let Some(set_name) = self.take_set_after_verb(p)? else {
            return Ok(None);
        };
        self.refuse_set_closure(*p)?;
        if !self.at_one_list_shape(*p) {
            return Ok(None);
        }
        *p += 3;
        if fold(&v2) == fold(head_var) {
            return Err(raise(
                "english-question-same-unknown",
                &[("unknown", &fold(&noun2))],
            ));
        }
        let nm = shape_name("list", rel, "");
        *relation2 = set_name.clone();
        Ok(Some(format!(
            "(rule ({nm} {v2} {head_var}) ({set_name} {v2}) (textjoin {head_var} \", \" ({rel} VlaListed {v2}))) (query {nm})"
        )))
    }

    /// VBA `ParseQuestion` (G-PROLOG slice 2 and after): the program tail a
    /// question compiles to, its quotes doubled for the formula text, and
    /// the engine and the relation(s) it reads.
    pub(super) fn parse_question(
        &self,
        pos: &mut usize,
        engine: &mut String,
        relation: &mut String,
        relation2: &mut String,
    ) -> Result<Option<String>, Refusal> {
        let mut p = *pos;
        let mut a1 = String::new();
        let a2: String;
        let mut u1 = false;
        let mut u2 = false;
        let mut is_whether = false;
        let mut rel = String::new();
        let mut tail = String::new();
        let mut shaped = false;
        relation2.clear();
        if self.tok(p) == "how" && self.tok(p + 1) == "many" {
            p += 2;
            match self.parse_how_many_question(&mut p, &mut rel, relation2)? {
                Some(t) => tail = t,
                None => return Ok(None),
            }
            shaped = true;
        } else if self.tok(p) == "whether" && self.tok(p + 1) == "every" {
            p += 2;
            match self.parse_every_question(&mut p, &mut rel, relation2)? {
                Some(t) => tail = t,
                None => return Ok(None),
            }
            shaped = true;
        } else if self.tok(p) == "which"
            && self.tok(p + 2) == "that"
            && is_set_verb(self.tok(p + 3))
        {
            p += 1;
            match self.parse_none_question(&mut p, &mut rel, relation2)? {
                Some(t) => tail = t,
                None => return Ok(None),
            }
            shaped = true;
        } else if self.tok(p) == "whether" {
            is_whether = true;
            p += 1;
            let Some(c) = self.take_constant(&mut p)? else {
                return Ok(None);
            };
            a1 = c;
            if self.at_alone_shape(p) {
                p += 1;
                match self.parse_alone_question(&mut p, "", &a1, &mut rel)? {
                    Some(t) => tail = t,
                    None => return Ok(None),
                }
                shaped = true;
            }
        } else {
            let Some(u) = self.take_unknown(&mut p) else {
                return Ok(None);
            };
            a1 = u;
            u1 = true;
            if self.at_alone_shape(p) {
                p += 1;
                match self.parse_alone_question(&mut p, &a1, "", &mut rel)? {
                    Some(t) => tail = t,
                    None => return Ok(None),
                }
                shaped = true;
            }
        }
        if shaped {
            if self.at_one_list_shape(p) {
                return Err(refuse_list_shape(self.list_shape_kind(*pos)));
            }
            *engine = "DATALOG".to_string();
            *relation = rel;
            *pos = p;
            return Ok(Some(tail.replace('"', "\"\"")));
        }
        let goal;
        let list_goal;
        let mut head_vars = String::new();
        let mut closure = false;
        if self.tok(p) == "is" {
            p += 1;
            self.skip_articles(&mut p);
            rel = self.word_at(p);
            if rel.is_empty() {
                return Ok(None);
            }
            refuse_grammar_word_as_name(&rel)?;
            p += 1;
            if self.at_directly_or_not(p) {
                return Err(raise(
                    "english-conditions-closure-needs-relation",
                    &[("shape", "a set")],
                ));
            }
            goal = format!("({rel} {a1})");
            list_goal = format!("({rel} VlaListed)");
            if u1 {
                head_vars = a1.clone();
            }
        } else if u1 && is_constant_tok(self.tok(p)) {
            let Some(c) = self.take_constant(&mut p)? else {
                return Ok(None);
            };
            a2 = c;
            rel = self.word_at(p);
            if rel.is_empty() {
                return Ok(None);
            }
            refuse_grammar_word_as_name(&rel)?;
            p += 1;
            let mut goal_pred = rel.clone();
            if self.at_directly_or_not(p) {
                p += 3;
                closure = true;
                goal_pred = closure_name(&rel);
            }
            goal = format!("({goal_pred} {a2} {a1})");
            list_goal = format!("({goal_pred} {a2} VlaListed)");
            head_vars = a1.clone();
        } else {
            rel = self.word_at(p);
            if rel.is_empty() {
                return Ok(None);
            }
            refuse_grammar_word_as_name(&rel)?;
            p += 1;
            if u1 && self.tok(p) == "each" {
                p += 1;
                let Some(t) = self.parse_each_list_question(&mut p, &rel, &a1, relation2)? else {
                    return Ok(None);
                };
                *engine = "DATALOG".to_string();
                *relation = rel;
                *pos = p;
                return Ok(Some(t.replace('"', "\"\"")));
            }
            match self.take_constant(&mut p)? {
                Some(c) => a2 = c,
                None => {
                    let Some(u) = self.take_unknown(&mut p) else {
                        return Ok(None);
                    };
                    a2 = u;
                    u2 = true;
                }
            }
            let mut goal_pred = rel.clone();
            if self.at_directly_or_not(p) {
                p += 3;
                closure = true;
                goal_pred = closure_name(&rel);
            }
            if is_whether && u2 {
                return Err(raise(
                    "english-question-whether-unknown",
                    &[("unknown", &fold(&a2))],
                ));
            }
            if u1 && u2 && fold(&a1) == fold(&a2) {
                return Err(raise(
                    "english-question-same-unknown",
                    &[("unknown", &fold(&a1))],
                ));
            }
            goal = format!("({goal_pred} {a1} {a2})");
            list_goal = format!("({goal_pred} VlaListed {a2})");
            if u1 {
                head_vars = a1.clone();
            }
            if u2 {
                if !head_vars.is_empty() {
                    head_vars.push(' ');
                }
                head_vars.push_str(&a2);
            }
        }
        let pre = if closure {
            closure_rules(&rel)
        } else {
            String::new()
        };
        *engine = "DATALOG".to_string();
        let tail = if self.at_one_list_shape(p) {
            if is_whether {
                return Err(refuse_list_shape("a whether question"));
            }
            if u2 {
                return Err(refuse_list_shape("a question with two unknowns"));
            }
            p += 3;
            format!("{pre}{}", list_tail(&rel, &a1, &list_goal))
        } else if is_whether {
            format!("{pre}(query {goal})")
        } else {
            format!("{pre}(rule (vla-ask-{rel} {head_vars}) {goal}) (query vla-ask-{rel})")
        };
        *relation = rel;
        *pos = p;
        Ok(Some(tail.replace('"', "\"\"")))
    }
}

/// Append `piece` to `s`, with `sep` between non-empty parts.
fn push_sep(s: &mut String, sep: char, piece: &str) {
    if !s.is_empty() {
        s.push(sep);
    }
    s.push_str(piece);
}

/// VBA `UnquoteRefLit`: a reference literal's content, lowercased.
pub fn unquote_ref_lit(lit: &str) -> String {
    let inner = if lit.len() >= 2 && lit.starts_with('"') && lit.ends_with('"') {
        &lit[1..lit.len() - 1]
    } else {
        lit
    };
    inner.to_lowercase()
}

/// VBA `SplitA1`: a column number and a row number from an A1 reference.
pub fn split_a1(s: &str) -> Option<(i64, i64)> {
    let s = s.to_lowercase();
    let mut letters = String::new();
    let mut digits = String::new();
    for ch in s.chars() {
        if ch.is_ascii_lowercase() && digits.is_empty() {
            letters.push(ch);
        } else if ch.is_ascii_digit() && !letters.is_empty() {
            digits.push(ch);
        } else {
            return None;
        }
    }
    if letters.is_empty() || letters.len() > 3 || digits.is_empty() || digits.len() > 7 {
        return None;
    }
    let mut col = 0i64;
    for b in letters.bytes() {
        col = col * 26 + (b as i64 - 96);
    }
    let row: i64 = digits.parse().ok()?;
    if row < 1 {
        return None;
    }
    Some((col, row))
}

/// VBA `RefInRange`: 1 inside, 0 outside, -1 when either reference cannot
/// be placed (a sheet qualifier, a shape that is not A1).
pub fn ref_in_range(cell_ref: &str, range_ref: &str) -> i64 {
    if cell_ref.contains('!') || range_ref.contains('!') {
        return -1;
    }
    let Some((cc, cr)) = split_a1(cell_ref) else {
        return -1;
    };
    let parts: Vec<&str> = range_ref.split(':').collect();
    let (mut c1, mut r1, mut c2, mut r2);
    match parts.len() {
        1 => {
            let Some((c, r)) = split_a1(parts[0]) else {
                return -1;
            };
            c1 = c;
            r1 = r;
            c2 = c;
            r2 = r;
        }
        2 => {
            let (Some((ca, ra)), Some((cb, rb))) = (split_a1(parts[0]), split_a1(parts[1])) else {
                return -1;
            };
            c1 = ca;
            r1 = ra;
            c2 = cb;
            r2 = rb;
        }
        _ => return -1,
    }
    if c1 > c2 {
        std::mem::swap(&mut c1, &mut c2);
    }
    if r1 > r2 {
        std::mem::swap(&mut r1, &mut r2);
    }
    if cc >= c1 && cc <= c2 && cr >= r1 && cr <= r2 {
        1
    } else {
        0
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn the_numerals_and_names_are_the_vba_s() {
        assert!(is_invariant_numeral("-3.5"));
        assert!(!is_invariant_numeral("3.5.1"));
        assert!(!is_invariant_numeral("."));
        assert!(is_canonical_numeral("10"));
        assert!(!is_canonical_numeral("010"));
        assert!(!is_canonical_numeral("1.50"));
        assert!(is_canonical_numeral("-1.5"));
        assert_eq!(role_to_var_name("person"), "Person");
        assert_eq!(role_to_var_name("Person"), "Person");
        assert_eq!(escape_name_part("needs-by"), "needs--by");
        assert_eq!(closure_name("reports-to"), "vla-any-reports--to");
        assert_eq!(shape_name("none", "a", "b"), "vla-none-a-b");
        assert_eq!(unquote_ref_lit("\"H2\""), "h2");
        assert_eq!(split_a1("AA10"), Some((27, 10)));
        assert_eq!(split_a1("a"), None);
        assert_eq!(ref_in_range("h3", "h2:h4"), 1);
        assert_eq!(ref_in_range("h5", "h4:h2"), 0);
        assert_eq!(ref_in_range("data!h3", "h2:h4"), -1);
    }
}
