//! The macro system: VLA.bas's `DefineMacro`, `ExpandMacros`, `ExpandOne`,
//! `Substitute` with the QUASIQUOTE bundle, the seventeen LISTOPS
//! primitives, `ExpandOnePass`, and `VlaExpandText` (PORT.5, slice 5d).
//!
//! Template-substitution macros: `(defmacro (name p1 p2 & rest) template
//! ...)`. A parameter is replaced by the form passed at the call site; the
//! rest parameter splices its forms into whatever list it stands in; several
//! template forms become one `(begin ...)`. No gensym, no compile-time
//! computation beyond the expand-time primitives, by design.
//!
//! The shape is the VBA's, function for function. `ExpandMacros` is the
//! hand-made trampoline: a form whose head is a macro is substituted and
//! looked at again, in a loop, and a `(begin ...)` a macro made (its line
//! is 0) is unwrapped around its last element, so a self-recursive walker's
//! chain length never becomes stack depth; the forms it set aside are
//! combined back into one `begin` at the end. Everything the macro system
//! makes carries line 0 (S3.1's explicit zero); a user's argument forms
//! keep their lines as they are copied in.
//!
//! The refusals are the catalogue's, by id, in the same situations.

use std::collections::HashMap;

use crate::form::{sym_text, Form, List};
use crate::intrinsics::{fold, is_numeric_literal_text, str_, val};
use crate::messages::{raise, Refusal};
use crate::printer::write_pretty;
use crate::reader::{count_lf, parse_all, tokenize};

/// One macro, as `DefineMacro` records it.
#[derive(Clone, Debug)]
pub struct Macro {
    /// The parameter names, as written.
    pub params: Vec<String>,
    /// The rest parameter's name as written, or empty.
    pub rest_name: String,
    pub template: Form,
    /// The docstring (L17), or empty.
    pub doc: String,
    /// The name as written.
    pub name: String,
}

/// The expand-time names a macro may not take, each with the description
/// its refusal gives (`DefineMacro`'s list, in its order).
const RESERVED: &[(&str, &str)] = &[
    ("quote", "data-literal form"),
    ("quasiquote", "template-shielding form"),
    ("unquote", "quasiquote escape form"),
    ("unquote-splicing", "quasiquote splice form"),
    ("symbol", "identifier-fusion primitive"),
    ("car", "list-head primitive"),
    ("cdr", "list-tail primitive"),
    ("cddr", "list-tail primitive"),
    ("cons", "list-construction primitive"),
    ("list", "list-construction primitive"),
    ("null?", "empty-list primitive"),
    ("eq?", "atom-comparison primitive"),
    ("equal?", "structural-comparison primitive"),
    ("quote-if", "expand-time conditional"),
    (
        "cond",
        "multi-clause conditional (folds at expand time when possible, defers to runtime otherwise)",
    ),
    ("+expand", "expand-time addition primitive"),
    ("=expand", "expand-time equality primitive"),
    ("<>expand", "expand-time inequality primitive"),
    (">expand", "expand-time comparison primitive"),
    ("<expand", "expand-time comparison primitive"),
    (">=expand", "expand-time comparison primitive"),
    ("<=expand", "expand-time comparison primitive"),
];

/// The seventeen LISTOPS primitives `Substitute` and `ExpandMacros` both
/// recognise at a form's head.
const LISTOPS: &[&str] = &[
    "car", "cdr", "cddr", "cons", "list", "null?", "eq?", "equal?", "quote-if", "cond", "+expand",
    "=expand", "<>expand", ">expand", "<expand", ">=expand", "<=expand",
];

fn is_listops(head_fold: &str) -> bool {
    LISTOPS.contains(&head_fold)
}

/// The substitution environment of one macro application: `Substitute`'s
/// `bindings`, `restName` (folded) and `restItems`. Empty for a primitive
/// met outside any template.
#[derive(Default)]
struct Env {
    bindings: HashMap<String, Form>,
    rest_name: String,
    rest_items: Vec<Form>,
}

/// The macro table and the application counters: `mMacros`, `mBudgetOn`,
/// `mExpandBudget` and `mExpandFired`.
#[derive(Default)]
pub struct Expander {
    macros: HashMap<String, Macro>,
    /// L15: when on, applications beyond the budget do not fire.
    pub budget_on: bool,
    pub expand_budget: u64,
    /// L1: macro applications since the last reset.
    pub expand_fired: u64,
}

impl Expander {
    pub fn new() -> Expander {
        Expander::default()
    }

    /// VBA `GetMacro`: by name, folded.
    pub fn get_macro(&self, name: &str) -> Option<&Macro> {
        self.macros.get(&fold(name))
    }

    /// VBA `VlaMacroDoc`: a macro's docstring, or empty.
    pub fn macro_doc(&self, name: &str) -> String {
        self.get_macro(name)
            .map(|m| m.doc.clone())
            .unwrap_or_default()
    }

    /// VBA `DefineMacro`: `(defmacro (name params... [& rest]) ["doc"]
    /// template...)`.
    pub fn define_macro(&mut self, f: &List) -> Result<(), Refusal> {
        let sig = f.nth_list(2)?;
        let mname = sig.head_sym()?;
        let mfold = fold(&mname);
        if let Some((word, desc)) = RESERVED.iter().find(|(w, _)| *w == mfold) {
            return Err(raise(
                "vla-defmacro-reserved-name",
                &[("word", word), ("desc", desc)],
            ));
        }

        let mut params = Vec::new();
        let mut rest_name = String::new();
        let mut j = 2;
        while j <= sig.count() {
            let t = sym_text(sig.nth(j)?)?;
            if t == "&" {
                if j + 1 > sig.count() {
                    return Err(raise(
                        "vla-defmacro-rest-param-missing",
                        &[("name", &mname)],
                    ));
                }
                rest_name = sym_text(sig.nth(j + 1)?)?;
                if j + 1 != sig.count() {
                    return Err(raise(
                        "vla-defmacro-rest-param-not-last",
                        &[("name", &mname)],
                    ));
                }
                break;
            }
            params.push(t);
            j += 1;
        }

        // L17: a string at form 3 is documentation only when at least one
        // more form follows; a lone string is the template.
        let mut doc = String::new();
        let mut doc_idx = 3;
        if f.count() >= 4 {
            if let Form::Str(s) = f.nth(3)? {
                doc = s.clone();
                doc_idx = 4;
            }
        }

        let template = if f.count() == doc_idx {
            f.nth(doc_idx)?.clone()
        } else if f.count() > doc_idx {
            let mut items = vec![Form::sym("begin")];
            items.extend(f.items[doc_idx - 1..].iter().cloned());
            Form::list(items)
        } else {
            return Err(raise("vla-defmacro-missing-template", &[("name", &mname)]));
        };

        // A redefinition overwrites, as the Dictionary's keyed assignment does.
        self.macros.insert(
            mfold,
            Macro {
                params,
                rest_name,
                template,
                doc,
                name: mname,
            },
        );
        Ok(())
    }

    /// VBA `ExpandMacros`: a form to its fixpoint.
    pub fn expand(&mut self, d: &Form, depth: u32) -> Result<Form, Refusal> {
        if depth > 200 {
            return Err(raise("vla-macro-expansion-too-deep", &[]));
        }
        let mut cur = d.clone();
        let mut chain_len: u32 = 0;
        let mut pending_head: Vec<Form> = Vec::new();
        loop {
            let Form::List(lst) = &cur else {
                return Ok(cur);
            };
            if lst.items.is_empty() {
                break;
            }
            let Form::Sym(head) = &lst.items[0] else {
                break;
            };
            let head_fold = fold(head);
            // P.L5: quote is a data context; expansion never enters.
            if head_fold == "quote" {
                return Ok(cur);
            }
            // QUASIQUOTE: (symbol ...) needs no ambient bindings.
            if head_fold == "symbol" {
                return fuse_symbol(lst, &Env::default());
            }
            if matches!(
                head_fold.as_str(),
                "quasiquote" | "unquote" | "unquote-splicing"
            ) {
                return Err(raise(
                    "vla-quasiquote-outside-template",
                    &[("head", &head_fold)],
                ));
            }
            // LISTOPS: the seventeen primitives, standalone.
            if is_listops(&head_fold) {
                return eval_listops_prim(&head_fold, lst, &Env::default());
            }
            if head_fold == "begin" && lst.count() > 1 && lst.line == 0 {
                // TABLESPEC-SCALE: a begin a macro made is a tail-position
                // wrapper; its other elements wait, its last is looked at
                // again.
                pending_head.extend(lst.items[1..lst.count() - 1].iter().cloned());
                chain_len += 1;
                if chain_len > 5000 {
                    return Err(raise("vla-macro-self-expansion-limit", &[]));
                }
                let last = lst.items[lst.count() - 1].clone();
                cur = last;
            } else if head_fold == "begin" {
                break;
            } else {
                let Some(rec) = self.get_macro(head).cloned() else {
                    break;
                };
                if self.budget_on && self.expand_fired >= self.expand_budget {
                    break; // L15
                }
                chain_len += 1;
                if chain_len > 5000 {
                    return Err(raise("vla-macro-self-expansion-limit", &[]));
                }
                let r = expand_one(lst, &rec)?;
                self.expand_fired += 1; // L1: one per application
                cur = r;
            }
        }
        // No macro at the head: rebuild with expanded elements, keeping the
        // list's own line (C2).
        let Form::List(cur_list) = cur else {
            unreachable!("the loop leaves only on a list");
        };
        let mut out = Vec::with_capacity(cur_list.count());
        for e in &cur_list.items {
            out.push(self.expand(e, depth)?);
        }
        let out = Form::List(List::new(out, cur_list.line));
        if pending_head.is_empty() {
            return Ok(out);
        }
        let mut final_out = vec![Form::sym("begin")];
        for e in &pending_head {
            final_out.push(self.expand(e, depth)?);
        }
        final_out.push(out);
        Ok(Form::list(final_out))
    }

    /// VBA `ExpandOnePass`: top-down; a list whose head is a macro is
    /// replaced by its substitution and not walked further.
    pub fn expand_one_pass(&mut self, d: &Form) -> Result<Form, Refusal> {
        let Form::List(lst) = d else {
            return Ok(d.clone());
        };
        if let Some(Form::Sym(head)) = lst.items.first() {
            if fold(head) == "quote" {
                return Ok(d.clone());
            }
            if let Some(rec) = self.get_macro(head).cloned() {
                if !(self.budget_on && self.expand_fired >= self.expand_budget) {
                    let r = expand_one(lst, &rec)?;
                    self.expand_fired += 1;
                    return Ok(r);
                }
            }
        }
        let mut out = Vec::with_capacity(lst.count());
        for e in &lst.items {
            out.push(self.expand_one_pass(e)?);
        }
        Ok(Form::List(List::new(out, lst.line)))
    }

    /// `VlaTranspile`'s and `VlaExpandText`'s pass 1 over the prelude and
    /// the source read together: every `defmacro` is defined, in order, and
    /// the other forms come back as the body.
    pub fn collect(&mut self, forms: Vec<Form>) -> Result<Vec<Form>, Refusal> {
        let mut body = Vec::with_capacity(forms.len());
        for f in forms {
            if f.head_is("defmacro") {
                let Form::List(l) = &f else { unreachable!() };
                self.define_macro(l)?;
            } else {
                body.push(f);
            }
        }
        Ok(body)
    }
}

/// VBA `VlaExpandText`: the prelude and the source read together, every
/// `defmacro` defined, every other form expanded (to its fixpoint, or one
/// pass) and written pretty, one per line. Returns the text and the number
/// of applications that fired.
pub fn expand_text(
    source: &str,
    prelude: &str,
    to_fixpoint: bool,
) -> Result<(String, u64), Refusal> {
    let line_offset = count_lf(prelude) as u32 + 1;
    let text = format!("{prelude}\r\n{source}");
    let forms = parse_all(&tokenize(&text), line_offset)?;
    let mut ex = Expander::new();
    let body = ex.collect(forms)?;
    let mut out = String::new();
    for f in &body {
        let expanded = if to_fixpoint {
            ex.expand(f, 0)?
        } else {
            ex.expand_one_pass(f)?
        };
        out.push_str(&write_pretty(&expanded, 0));
        out.push_str("\r\n");
    }
    Ok((out, ex.expand_fired))
}

/// VBA `ExpandOne`: one application of a macro to a call form.
fn expand_one(call: &List, rec: &Macro) -> Result<Form, Refusal> {
    let argc = call.count() - 1;
    let name = call.items[0].atom_text().unwrap_or_default();
    if rec.rest_name.is_empty() {
        if argc != rec.params.len() {
            return Err(raise(
                "vla-macro-arity",
                &[
                    ("name", &name),
                    ("n", &rec.params.len().to_string()),
                    ("given", &argc.to_string()),
                ],
            ));
        }
    } else if argc < rec.params.len() {
        return Err(raise(
            "vla-macro-arity-min",
            &[
                ("name", &name),
                ("n", &rec.params.len().to_string()),
                ("given", &argc.to_string()),
            ],
        ));
    }
    let mut env = Env {
        bindings: HashMap::new(),
        rest_name: fold(&rec.rest_name),
        rest_items: Vec::new(),
    };
    for (j, p) in rec.params.iter().enumerate() {
        env.bindings.insert(fold(p), call.items[j + 1].clone());
    }
    env.rest_items
        .extend(call.items[rec.params.len() + 1..].iter().cloned());
    substitute(&rec.template, &env, false)
}

/// VBA `Substitute`: a template with its parameters replaced. Inside a
/// quasiquote shield (`in_quasi`) a bare symbol is data, and only an
/// explicit `unquote` or `unquote-splicing` reaches the bindings.
fn substitute(t: &Form, env: &Env, in_quasi: bool) -> Result<Form, Refusal> {
    let lst = match t {
        Form::List(l) => l,
        Form::Sym(s) if !in_quasi => {
            let key = fold(s);
            if !env.rest_name.is_empty() && key == env.rest_name {
                // The rest parameter standalone: the whole list.
                return Ok(Form::list(env.rest_items.clone()));
            }
            if let Some(v) = env.bindings.get(&key) {
                return Ok(v.clone());
            }
            return Ok(t.clone());
        }
        _ => return Ok(t.clone()),
    };

    if let Some(Form::Sym(head)) = lst.items.first() {
        let h = fold(head);
        match h.as_str() {
            "symbol" => return fuse_symbol(lst, env),
            "quasiquote" => {
                if in_quasi {
                    return Err(raise("vla-quasiquote-nested", &[]));
                }
                if lst.count() != 2 {
                    return Err(raise(
                        "vla-quasiquote-arity",
                        &[("n", &(lst.count() - 1).to_string())],
                    ));
                }
                return substitute(lst.nth(2)?, env, true);
            }
            "unquote" => {
                if !in_quasi {
                    return Err(raise("vla-unquote-outside-quasiquote", &[]));
                }
                if lst.count() != 2 {
                    return Err(raise(
                        "vla-unquote-arity",
                        &[("n", &(lst.count() - 1).to_string())],
                    ));
                }
                return substitute(lst.nth(2)?, env, false);
            }
            "unquote-splicing" => {
                // A splice in its place is taken by the element walk below;
                // reaching here it is a bare value, or outside a shield.
                if !in_quasi {
                    return Err(raise("vla-unquote-splicing-outside-quasiquote", &[]));
                }
                return Err(raise("vla-unquote-splicing-not-element", &[]));
            }
            _ if is_listops(&h) && !in_quasi => {
                return eval_listops_prim(&h, lst, env);
            }
            _ => {}
        }
    }

    let mut out = Vec::with_capacity(lst.count());
    for e in &lst.items {
        let mut spliced = false;
        if in_quasi && e.head_is("unquote-splicing") {
            let Form::List(spl) = e else { unreachable!() };
            if spl.count() != 2 {
                return Err(raise(
                    "vla-unquote-splicing-arity",
                    &[("n", &(spl.count() - 1).to_string())],
                ));
            }
            let val = substitute(spl.nth(2)?, env, false)?;
            let Form::List(val) = val else {
                return Err(raise("vla-unquote-splicing-not-list", &[]));
            };
            out.extend(val.items.iter().cloned());
            spliced = true;
        } else if !in_quasi {
            if let Form::Sym(s) = e {
                if !env.rest_name.is_empty() && fold(s) == env.rest_name {
                    // The rest parameter as an element: its forms spliced in.
                    out.extend(env.rest_items.iter().cloned());
                    spliced = true;
                }
            }
        }
        if !spliced {
            out.push(substitute(e, env, in_quasi)?);
        }
    }
    // S3.1: a template copy is explicitly untagged; the user's own forms
    // spliced in above keep their lines.
    Ok(Form::list(out))
}

/// VBA `FuseSymbol`: `(symbol a b ...)` fuses each argument's resolved
/// text into one new atom. A fused text that begins with a quote mark is
/// a string literal, as the VBA's leading-quote marker makes it.
fn fuse_symbol(lst: &List, env: &Env) -> Result<Form, Refusal> {
    if lst.count() < 2 {
        return Err(raise("vla-symbol-arity", &[]));
    }
    let mut fused = String::new();
    for k in 2..=lst.count() {
        let piece = substitute(lst.nth(k)?, env, false)?;
        match piece {
            Form::List(_) => {
                return Err(raise(
                    "vla-symbol-arg-is-list",
                    &[("n", &(k - 1).to_string())],
                ));
            }
            Form::Str(s) => fused.push_str(&s),
            Form::Sym(s) => fused.push_str(&s),
        }
    }
    Ok(match fused.strip_prefix('"') {
        Some(rest) => Form::Str(rest.to_string()),
        None => Form::Sym(fused),
    })
}

/// VBA `EvalListopsPrim`: the dispatcher the two recognition sites share.
fn eval_listops_prim(head_fold: &str, lst: &List, env: &Env) -> Result<Form, Refusal> {
    match head_fold {
        "car" => eval_car(lst, env),
        "cdr" => eval_cdr(lst, env),
        "cddr" => eval_cddr(lst, env),
        "cons" => eval_cons(lst, env),
        "list" => eval_list(lst, env),
        "null?" => eval_nullq(lst, env),
        "eq?" => eval_eq(lst, env),
        "equal?" => eval_equal(lst, env),
        "quote-if" => eval_quote_if(lst, env),
        "cond" => eval_cond_from(lst, 2, env),
        _ => eval_arith_expand(head_fold, lst, env),
    }
}

/// VBA `ResolveListopsArg`: the argument through `Substitute`, then one
/// top-level `(quote X)` unwrapped to its data.
fn resolve_listops_arg(arg: &Form, env: &Env) -> Result<Form, Refusal> {
    let v = substitute(arg, env, false)?;
    if v.head_is("quote") {
        let Form::List(qc) = &v else { unreachable!() };
        if qc.count() != 2 {
            return Err(raise("vla-quote-arity", &[]));
        }
        return Ok(qc.items[1].clone());
    }
    Ok(v)
}

fn arity_n(lst: &List) -> String {
    (lst.count() - 1).to_string()
}

fn eval_car(lst: &List, env: &Env) -> Result<Form, Refusal> {
    if lst.count() != 2 {
        return Err(raise("vla-car-arity", &[("n", &arity_n(lst))]));
    }
    let v = resolve_listops_arg(lst.nth(2)?, env)?;
    let Form::List(c) = &v else {
        return Err(raise(
            "vla-car-expected-list",
            &[("value", &v.atom_text().unwrap_or_default())],
        ));
    };
    if c.items.is_empty() {
        return Err(raise("vla-car-empty-list", &[]));
    }
    Ok(c.items[0].clone())
}

/// VBA `ListTail`: the items from position `n + 1` on, untagged.
fn list_tail(c: &List, n: usize, caller: &str) -> Result<Form, Refusal> {
    if c.count() < n {
        return Err(raise(
            "vla-list-too-short",
            &[
                ("caller", caller),
                ("n", &n.to_string()),
                ("count", &c.count().to_string()),
            ],
        ));
    }
    Ok(Form::list(c.items[n..].to_vec()))
}

fn eval_cdr(lst: &List, env: &Env) -> Result<Form, Refusal> {
    if lst.count() != 2 {
        return Err(raise("vla-cdr-arity", &[("n", &arity_n(lst))]));
    }
    let v = resolve_listops_arg(lst.nth(2)?, env)?;
    let Form::List(c) = &v else {
        return Err(raise(
            "vla-cdr-expected-list",
            &[("value", &v.atom_text().unwrap_or_default())],
        ));
    };
    list_tail(c, 1, "cdr")
}

fn eval_cddr(lst: &List, env: &Env) -> Result<Form, Refusal> {
    if lst.count() != 2 {
        return Err(raise("vla-cddr-arity", &[("n", &arity_n(lst))]));
    }
    let v = resolve_listops_arg(lst.nth(2)?, env)?;
    let Form::List(c) = &v else {
        return Err(raise(
            "vla-cddr-expected-list",
            &[("value", &v.atom_text().unwrap_or_default())],
        ));
    };
    list_tail(c, 2, "cddr")
}

fn eval_cons(lst: &List, env: &Env) -> Result<Form, Refusal> {
    if lst.count() != 3 {
        return Err(raise("vla-cons-arity", &[("n", &arity_n(lst))]));
    }
    let item = resolve_listops_arg(lst.nth(2)?, env)?;
    let tail = resolve_listops_arg(lst.nth(3)?, env)?;
    let Form::List(tail) = &tail else {
        return Err(raise(
            "vla-cons-second-not-list",
            &[("value", &tail.atom_text().unwrap_or_default())],
        ));
    };
    let mut out = Vec::with_capacity(tail.count() + 1);
    out.push(item);
    out.extend(tail.items.iter().cloned());
    Ok(Form::list(out))
}

fn eval_list(lst: &List, env: &Env) -> Result<Form, Refusal> {
    let mut out = Vec::with_capacity(lst.count());
    for k in 2..=lst.count() {
        out.push(resolve_listops_arg(lst.nth(k)?, env)?);
    }
    Ok(Form::list(out))
}

fn bool_sym(b: bool) -> Form {
    Form::sym(if b { "true" } else { "false" })
}

/// VBA `AtomEqual`: symbols folded, strings exact, never across the kinds.
fn atom_equal(a: &Form, b: &Form) -> bool {
    match (a, b) {
        (Form::Sym(x), Form::Sym(y)) => fold(x) == fold(y),
        (Form::Str(x), Form::Str(y)) => x == y,
        _ => false,
    }
}

/// VBA `DeepEqual`.
fn deep_equal(a: &Form, b: &Form) -> bool {
    match (a, b) {
        (Form::List(x), Form::List(y)) => {
            x.count() == y.count() && x.items.iter().zip(&y.items).all(|(p, q)| deep_equal(p, q))
        }
        (Form::List(_), _) | (_, Form::List(_)) => false,
        _ => atom_equal(a, b),
    }
}

fn eval_nullq(lst: &List, env: &Env) -> Result<Form, Refusal> {
    if lst.count() != 2 {
        return Err(raise("vla-nullq-arity", &[("n", &arity_n(lst))]));
    }
    let v = resolve_listops_arg(lst.nth(2)?, env)?;
    Ok(match &v {
        Form::List(c) => bool_sym(c.items.is_empty()),
        _ => bool_sym(false),
    })
}

fn eval_eq(lst: &List, env: &Env) -> Result<Form, Refusal> {
    if lst.count() != 3 {
        return Err(raise("vla-eq-arity", &[("n", &arity_n(lst))]));
    }
    let a = resolve_listops_arg(lst.nth(2)?, env)?;
    let b = resolve_listops_arg(lst.nth(3)?, env)?;
    if a.is_list() || b.is_list() {
        return Err(raise("vla-eq-expected-atom", &[]));
    }
    Ok(bool_sym(atom_equal(&a, &b)))
}

fn eval_equal(lst: &List, env: &Env) -> Result<Form, Refusal> {
    if lst.count() != 3 {
        return Err(raise("vla-equal-arity", &[("n", &arity_n(lst))]));
    }
    let a = resolve_listops_arg(lst.nth(2)?, env)?;
    let b = resolve_listops_arg(lst.nth(3)?, env)?;
    Ok(bool_sym(deep_equal(&a, &b)))
}

/// `(quote-if test then else)`: only the selected branch is ever resolved.
fn eval_quote_if(lst: &List, env: &Env) -> Result<Form, Refusal> {
    if lst.count() != 4 {
        return Err(raise("vla-quote-if-arity", &[("n", &arity_n(lst))]));
    }
    let t = resolve_listops_arg(lst.nth(2)?, env)?;
    let tf = match &t {
        Form::List(_) => return Err(raise("vla-quote-if-test-is-list", &[])),
        Form::Str(_) => {
            return Err(raise(
                "vla-quote-if-test-invalid",
                &[("value", &t.atom_text().unwrap_or_default())],
            ))
        }
        Form::Sym(s) => fold(s),
    };
    let branch = match tf.as_str() {
        "true" => lst.nth(3)?,
        "false" => lst.nth(4)?,
        _ => {
            return Err(raise(
                "vla-quote-if-test-invalid",
                &[("value", &t.atom_text().unwrap_or_default())],
            ))
        }
    };
    substitute(branch, env, false)
}

/// LISTOPS-EXPAND: `(+expand a b)` and the six comparisons, over literal
/// numeric operands only, through `Val` and `Str$`.
fn eval_arith_expand(op: &str, lst: &List, env: &Env) -> Result<Form, Refusal> {
    if lst.count() != 3 {
        return Err(raise(
            "vla-arith-expand-arity",
            &[("op", op), ("n", &arity_n(lst))],
        ));
    }
    let a = resolve_listops_arg(lst.nth(2)?, env)?;
    let b = resolve_listops_arg(lst.nth(3)?, env)?;
    let mut texts = Vec::with_capacity(2);
    for v in [&a, &b] {
        let Some(text) = v.atom_text() else {
            return Err(raise("vla-arith-expand-not-number-list", &[("op", op)]));
        };
        texts.push(text);
    }
    for text in &texts {
        if !is_numeric_literal_text(text) {
            return Err(raise(
                "vla-arith-expand-not-number-value",
                &[("op", op), ("value", text)],
            ));
        }
    }
    let na = val(&texts[0]);
    let nb = val(&texts[1]);
    Ok(match op {
        "+expand" => Form::Sym(str_(na + nb).trim().to_string()),
        "=expand" => bool_sym(na == nb),
        "<>expand" => bool_sym(na != nb),
        ">expand" => bool_sym(na > nb),
        "<expand" => bool_sym(na < nb),
        ">=expand" => bool_sym(na >= nb),
        _ => bool_sym(na <= nb),
    })
}

/// VBA `EvalCondFrom`: `(cond (test form) ... (else form))`, one clause at
/// a time from `start_idx`; a test that folds to true or false is taken or
/// skipped at expand time, any other test defers to a native `if`.
fn eval_cond_from(lst: &List, start_idx: usize, env: &Env) -> Result<Form, Refusal> {
    if start_idx > lst.count() {
        return Ok(Form::list(vec![Form::sym("begin")]));
    }
    let cl = lst.nth_list(start_idx)?;
    if cl.count() != 2 {
        return Err(raise(
            "vla-cond-clause-shape",
            &[("n", &cl.count().to_string())],
        ));
    }
    if let Form::Sym(s) = cl.nth(1)? {
        if fold(s) == "else" {
            if start_idx != lst.count() {
                return Err(raise("vla-cond-else-not-last", &[]));
            }
            return substitute(cl.nth(2)?, env, false);
        }
    }
    let t = resolve_listops_arg(cl.nth(1)?, env)?;
    let tf = match &t {
        Form::Sym(s) => fold(s),
        _ => String::new(),
    };
    if tf == "true" {
        return substitute(cl.nth(2)?, env, false);
    }
    if tf == "false" {
        return eval_cond_from(lst, start_idx + 1, env);
    }
    let raw_form = substitute(cl.nth(2)?, env, false)?;
    let rest_form = eval_cond_from(lst, start_idx + 1, env)?;
    Ok(Form::list(vec![
        Form::sym("if"),
        t,
        Form::list(vec![Form::sym("then"), raw_form]),
        Form::list(vec![Form::sym("else"), rest_form]),
    ]))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::read_forms;

    const PRELUDE: &str = include_str!("../../scripts/prelude.vla");

    /// `VlaExpandText(src, True, fired)` under the real prelude, normalised
    /// to LF for the pins; the VBA pins compare fragments of this text.
    fn expand(src: &str) -> String {
        let (text, _) = expand_text(src, PRELUDE, true).unwrap_or_else(|e| panic!("{src}: {e}"));
        text.replace("\r\n", "\n")
    }

    /// The pins' `CheckExpandErr`: the refusal's text holds the fragment.
    fn expand_err(src: &str) -> Refusal {
        match expand_text(src, PRELUDE, true) {
            Ok((text, _)) => panic!("{src} expanded instead of refusing: {text}"),
            Err(e) => e,
        }
    }

    fn one(src: &str) -> String {
        expand(src).trim_end().to_string()
    }

    #[test]
    fn ordinary_substitution_and_the_rest_parameter() {
        assert_eq!(
            one("(defmacro (make-adder n) (begin (dim total Long) (set! total (+ total n)))) (make-adder 5)"),
            "(begin (dim total Long) (set! total (+ total 5)))"
        );
        assert_eq!(
            one("(defmacro (all & xs) (list-of xs)) (all a b c)"),
            "(list-of a b c)"
        );
        // In element position the rest parameter splices; only a template
        // that is the bare name gives the whole list.
        assert_eq!(
            one("(defmacro (whole & xs) (f xs 1)) (whole a b)"),
            "(f a b 1)"
        );
        assert_eq!(one("(defmacro (bare & xs) xs) (bare a b)"), "(a b)");
        // Two template forms become one begin; a docstring is not one.
        assert_eq!(
            one("(defmacro (two x) \"the doc\" (a x) (b x)) (two 1)"),
            "(begin (a 1) (b 1))"
        );
        assert_eq!(
            one("(defmacro (lone) \"just this\") (lone)"),
            "\"just this\""
        );
        // A nested macro call inside a substitution expands too.
        assert_eq!(
            one("(defmacro (inner x) (set! x 1)) (defmacro (outer y) (inner y)) (outer q)"),
            "(set! q 1)"
        );
        assert_eq!(one("(quote (when unless))"), "(quote (when unless))");
    }

    #[test]
    fn the_prelude_macros_fire_and_the_count_is_reported() {
        let (text, fired) = expand_text("(when c a b)", PRELUDE, true).unwrap();
        assert_eq!(text.trim_end(), "(if c (then a b))");
        assert_eq!(fired, 1);
        let (text, fired) = expand_text("(when c a b)", PRELUDE, false).unwrap();
        assert_eq!(text.trim_end(), "(if c (then a b))");
        assert_eq!(fired, 1);
        let (_, fired) = expand_text("(set! x 1)", PRELUDE, true).unwrap();
        assert_eq!(fired, 0);
    }

    #[test]
    fn arity_refusals_name_the_macro() {
        let r = expand_err("(defmacro (two a b) (f a b)) (two 1)");
        assert_eq!(r.id, "vla-macro-arity");
        assert_eq!(r.text, "macro 'two' expects 2 argument(s), got 1");
        let r = expand_err("(defmacro (atleast a & r) (f a r)) (atleast)");
        assert_eq!(r.id, "vla-macro-arity-min");
        assert_eq!(
            r.text,
            "macro 'atleast' expects at least 1 argument(s), got 0"
        );
        assert_eq!(
            expand_err("(defmacro (m a &) a)").id,
            "vla-defmacro-rest-param-missing"
        );
        assert_eq!(
            expand_err("(defmacro (m & r x) r)").id,
            "vla-defmacro-rest-param-not-last"
        );
        assert_eq!(
            expand_err("(defmacro (m x))").id,
            "vla-defmacro-missing-template"
        );
    }

    // TestListops (VLA_Tests.bas), the standalone primitives.
    #[test]
    fn listops_car_cdr_cddr_cons_list() {
        assert_eq!(one("(car (quote (1 2 3)))"), "1");
        assert_eq!(one("(cdr (quote (1 2 3)))"), "(2 3)");
        assert_eq!(one("(cddr (quote (1 2 3 4)))"), "(3 4)");
        assert!(expand_err("(car (quote ()))")
            .text
            .contains("car: expected a non-empty list"));
        assert!(expand_err("(cdr (quote ()))")
            .text
            .contains("cdr: expected a list of at least 1 element(s), got 0"));
        assert!(expand_err("(cddr (quote (1)))")
            .text
            .contains("cddr: expected a list of at least 2 element(s), got 1"));
        assert!(expand_err("(car 5)")
            .text
            .contains("car: expected a list, got '5'"));
        assert!(expand_err("(car (quote (1 2)) (quote (3 4)))")
            .text
            .contains("car expects exactly 1 argument, got 2"));
        assert_eq!(one("(cons 1 (quote (2 3)))"), "(1 2 3)");
        assert!(expand_err("(cons 1 2)")
            .text
            .contains("cons: second argument must be a list, got '2'"));
        assert_eq!(one("(list 1 2 3)"), "(1 2 3)");
        assert_eq!(one("(list)"), "()");
    }

    #[test]
    fn listops_predicates_and_quote_if() {
        assert_eq!(one("(null? (quote ()))"), "true");
        assert_eq!(one("(null? (quote (1)))"), "false");
        assert_eq!(one("(null? 5)"), "false");
        assert_eq!(one("(eq? (quote a) (quote a))"), "true");
        assert_eq!(one("(eq? (quote a) (quote b))"), "false");
        assert_eq!(one("(eq? (quote A) (quote a))"), "true");
        assert_eq!(one("(eq? \"a\" (quote a))"), "false");
        assert!(expand_err("(eq? (quote (1)) (quote (1)))")
            .text
            .contains("eq?: expected an atom, got a list"));
        assert_eq!(
            one("(equal? (quote (1 2 (3 4))) (quote (1 2 (3 4))))"),
            "true"
        );
        assert_eq!(
            one("(equal? (quote (1 2 (3 4))) (quote (1 2 (3 5))))"),
            "false"
        );
        assert_eq!(
            one("(quote-if (null? (quote ())) (debug-print \"base\") (car (quote ())))"),
            "(debug-print \"base\")"
        );
        assert_eq!(
            one("(quote-if (null? (quote (1))) (car (quote ())) (debug-print \"recurse\"))"),
            "(debug-print \"recurse\")"
        );
        assert!(
            expand_err("(quote-if 5 (debug-print \"a\") (debug-print \"b\"))")
                .text
                .contains("quote-if: test must resolve to true or false")
        );
        assert!(
            expand_err("(quote-if (null? (quote ())) (debug-print \"a\"))")
                .text
                .contains("quote-if expects exactly 3 arguments")
        );
        for (word, frag) in [
            (
                "car",
                "'car' is the list-head primitive and cannot be a macro name",
            ),
            ("cond", "'cond' is the multi-clause conditional"),
            ("quote", "'quote' is the data-literal form"),
            ("+expand", "'+expand' is the expand-time addition primitive"),
        ] {
            let r = expand_err(&format!("(defmacro ({word} x) x)"));
            assert_eq!(r.id, "vla-defmacro-reserved-name");
            assert!(r.text.contains(frag), "{word}: {}", r.text);
        }
        assert_eq!(
            one("(defmacro (first-of xs) (car xs)) (first-of (quote (7 8 9)))"),
            "7"
        );
        assert_eq!(
            one("(defmacro (demo xs) (quasiquote (unquote (car xs)))) (demo (quote (1 2 3)))"),
            "1"
        );
    }

    #[test]
    fn a_self_terminating_walker_emits_one_statement_per_row() {
        let walker = "(defmacro (walk-rows lst) (quote-if (null? lst) (begin) (begin (debug-print (car lst)) (walk-rows (cdr lst)))))";
        let text = one(&format!("{walker} (walk-rows (quote (10 20 30)))"));
        // The trampoline flattens the chain into one begin, as the VBA's
        // own pendingHead combine does.
        assert_eq!(
            text,
            "(begin (debug-print 10) (debug-print 20) (debug-print 30) (begin))"
        );
    }

    // TestListopsExpand: the seven arithmetic primitives.
    #[test]
    fn expand_time_arithmetic() {
        assert_eq!(one("(+expand 2 3)"), "5");
        assert_eq!(one("(+expand -2 3)"), "1");
        assert_eq!(one("(+expand 2.5 1.5)"), "4");
        assert_eq!(one("(+expand 0.1 0.2)"), ".3");
        assert_eq!(one("(=expand 5 5)"), "true");
        assert_eq!(one("(=expand 5 6)"), "false");
        assert_eq!(one("(<>expand 5 6)"), "true");
        assert_eq!(one("(>expand 5 3)"), "true");
        assert_eq!(one("(>expand 3 5)"), "false");
        assert_eq!(one("(<expand 3 5)"), "true");
        assert_eq!(one("(>=expand 5 5)"), "true");
        assert_eq!(one("(<=expand 5 5)"), "true");
        assert!(expand_err("(+expand 2 (quote x))")
            .text
            .contains("+expand: expected a number, got 'x'"));
        assert!(expand_err("(>expand (quote (1)) 2)")
            .text
            .contains(">expand: expected a number, got a list"));
        assert!(expand_err("(+expand (quote 1,000) 1)")
            .text
            .contains("+expand: expected a number, got '1,000'"));
        assert!(expand_err("(+expand 2)")
            .text
            .contains("+expand expects exactly 2 arguments, got 1"));
        assert_eq!(
            one("(quote-if (>expand 5 3) (debug-print \"yes\") (debug-print \"no\"))"),
            "(debug-print \"yes\")"
        );
        assert_eq!(
            one("(defmacro (bigger x) (quote-if (>expand x 10) (quote yes) (quote no))) (bigger 20)"),
            "(quote yes)"
        );
    }

    // TestCond.
    #[test]
    fn cond_folds_defers_and_refuses() {
        assert_eq!(
            one("(cond (true (debug-print \"a\")) (false (car (quote ()))))"),
            "(debug-print \"a\")"
        );
        assert_eq!(
            one("(cond (false (car (quote ()))) (true (debug-print \"b\")))"),
            "(debug-print \"b\")"
        );
        assert_eq!(
            one("(cond ((null? (quote ())) (debug-print \"yes\")) (else (debug-print \"no\")))"),
            "(debug-print \"yes\")"
        );
        assert_eq!(
            one("(cond ((null? (quote (1))) (debug-print \"wrong\")) (else (debug-print \"right\")))"),
            "(debug-print \"right\")"
        );
        // A runtime test defers to a native if, the rest in its else.
        assert_eq!(
            one("(cond ((> x 0) (debug-print \"pos\")) (else (debug-print \"nonpos\")))"),
            "(if (> x 0) (then (debug-print \"pos\")) (else (debug-print \"nonpos\")))"
        );
        assert_eq!(
            one("(cond ((> x 0) (debug-print \"pos\")) (true (debug-print \"fallback\")))"),
            "(if (> x 0) (then (debug-print \"pos\")) (else (debug-print \"fallback\")))"
        );
        assert_eq!(one("(cond (false (debug-print \"a\")))"), "(begin)");
        assert_eq!(one("(cond)"), "(begin)");
        assert!(
            expand_err("(cond (else (debug-print \"a\")) (true (debug-print \"b\")))")
                .text
                .contains("cond: 'else' must be the last clause")
        );
        assert!(expand_err("(cond ((> 1 0)))")
            .text
            .contains("cond: each clause must be (test form) or (else form), got 1 element(s)"));
        assert_eq!(
            one("(defmacro (branch x) (cond ((null? x) (quote empty)) (else (quote nonempty)))) (branch (quote ()))"),
            "(quote empty)"
        );
    }

    // TestQuasiquote.
    #[test]
    fn quasiquote_shields_and_unquote_punches_through() {
        assert_eq!(one("(symbol \"make-\" \"bold\")"), "make-bold");
        assert_eq!(one("(symbol \"bold\")"), "bold");
        assert!(expand_err("(symbol)")
            .text
            .contains("symbol expects at least 1 argument, got 0"));
        assert_eq!(
            one("(defmacro (make-getter suffix) (function (symbol \"get-\" suffix) () String (return \"hi\"))) (make-getter widget)"),
            "(function get-widget () String (return \"hi\"))"
        );
        for (word, frag) in [
            ("symbol", "'symbol' is the identifier-fusion primitive"),
            ("quasiquote", "'quasiquote' is the template-shielding form"),
            ("unquote", "'unquote' is the quasiquote escape form"),
            (
                "unquote-splicing",
                "'unquote-splicing' is the quasiquote splice form",
            ),
        ] {
            assert!(expand_err(&format!("(defmacro ({word} x) x)"))
                .text
                .contains(frag));
        }
        assert_eq!(
            one("(defmacro (demo x) (quasiquote (x (unquote x)))) (demo hello)"),
            "(x hello)"
        );
        assert_eq!(
            one("(defmacro (demo & rest) (quasiquote (literal (unquote-splicing rest) end))) (demo a b c)"),
            "(literal a b c end)"
        );
        assert!(expand_err(
            "(defmacro (demo x) (quasiquote (a (quasiquote (unquote x))))) (demo hello)"
        )
        .text
        .contains("nested quasiquote is not supported"));
        assert!(expand_err("(unquote x)")
            .text
            .contains("'unquote' only has meaning inside a defmacro's own template"));
        assert!(expand_err("(defmacro (demo x) (unquote x)) (demo hello)")
            .text
            .contains("unquote used outside quasiquote"));
        assert!(
            expand_err("(defmacro (demo x) (unquote-splicing x)) (demo hello)")
                .text
                .contains("unquote-splicing used outside quasiquote")
        );
        assert!(
            expand_err("(defmacro (demo x) (quasiquote (unquote-splicing x))) (demo hello)")
                .text
                .contains("unquote-splicing must appear as a list element, not as a value")
        );
        assert!(expand_err(
            "(defmacro (demo x) (quasiquote (a (unquote-splicing x) b))) (demo hello)"
        )
        .text
        .contains("unquote-splicing: operand did not resolve to a list"));
        assert!(
            expand_err("(defmacro (demo x) (quasiquote a b)) (demo hello)")
                .text
                .contains("quasiquote expects exactly 1 argument, got 2")
        );
        assert!(
            expand_err("(defmacro (demo x) (quasiquote (unquote x y))) (demo hello)")
                .text
                .contains("unquote expects exactly 1 argument, got 2")
        );
        assert!(expand_err(
            "(defmacro (demo x) (quasiquote (a (unquote-splicing x y) b))) (demo hello)"
        )
        .text
        .contains("unquote-splicing expects exactly 1 argument, got 2"));
    }

    #[test]
    fn lines_survive_where_the_vba_keeps_them() {
        // A user's argument keeps its line through substitution; the
        // template's own lists carry none.
        let mut ex = Expander::new();
        let defs = read_forms("(defmacro (wrap x) (begin x))").unwrap();
        let body = ex.collect(defs).unwrap();
        assert!(body.is_empty());
        let call = read_forms("\n\n(wrap (f 1))").unwrap().remove(0);
        // A one-element begin a macro made is unwrapped by the trampoline;
        // the user's own form comes through with its line.
        let out = ex.expand(&call, 0).unwrap();
        assert_eq!(out, read_forms("(f 1)").unwrap()[0]);
        assert_eq!(out.as_list().unwrap().line, 3);
        let defs = read_forms("(defmacro (wrap2 x) (begin (g 0) x))").unwrap();
        ex.collect(defs).unwrap();
        let call = read_forms("\n\n(wrap2 (f 1))").unwrap().remove(0);
        let out = ex.expand(&call, 0).unwrap();
        let Form::List(l) = &out else { panic!() };
        assert_eq!(l.line, 0);
        assert_eq!(l.items[1].as_list().unwrap().line, 0);
        assert_eq!(l.items[2].as_list().unwrap().line, 3);
        // A hand-written begin keeps its line through the rebuild.
        let hand = read_forms("\n(begin (f 1) (g 2))").unwrap().remove(0);
        let out = ex.expand(&hand, 0).unwrap();
        assert_eq!(out.as_list().unwrap().line, 2);
    }

    #[test]
    fn every_golden_form_expands_under_the_prelude() {
        // The corpus golden carries its own defmacros at the end; every
        // other form expands without a refusal (the compile oracle holds
        // the result to the byte in slice 5e).
        const GOLDEN: &str = include_str!("../../scripts/instructions_golden.vla");
        let (text, fired) = expand_text(GOLDEN, PRELUDE, true).unwrap();
        assert!(fired > 100, "{fired} applications");
        assert!(text.starts_with("(const hot-pink \"#FF69B4\")"));
    }
}
