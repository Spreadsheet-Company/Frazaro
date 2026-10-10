//! The macro system: VLA.bas's `DefineMacro`, `ExpandMacros`, `ExpandOne`,
//! `Substitute` with the QUASIQUOTE bundle, the twenty-seven LISTOPS
//! primitives, `ExpandOnePass`, and `VlaExpandText` (PORT.5, slice 5d; the
//! expand-time object model's ten, VLA.3).
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
use crate::reader::{count_lf, parse_all, src_line_tag, tokenize, FORM_DEPTH_LIMIT};
use crate::refers::letters_to_col;
use crate::sheet::{column_letters, MAX_COLUMN, MAX_ROW};

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
    // VLA.3: the expand-time object model's ten.
    ("-expand", "expand-time subtraction primitive"),
    ("*expand", "expand-time multiplication primitive"),
    ("quotient-expand", "expand-time integer quotient primitive"),
    ("mod-expand", "expand-time remainder primitive"),
    ("&expand", "expand-time text-joining primitive"),
    ("characters-expand", "expand-time text-splitting primitive"),
    ("address-expand", "expand-time address primitive"),
    ("row-expand", "expand-time row-number primitive"),
    ("column-expand", "expand-time column-number primitive"),
    ("column-letters-expand", "expand-time column-letters primitive"),
];

/// The twenty-seven LISTOPS primitives `Substitute` and `ExpandMacros` both
/// recognise at a form's head: the original nine, LISTOPS-EXPAND's seven,
/// `cond`, and the expand-time object model's ten (VLA.3).
const LISTOPS: &[&str] = &[
    "car",
    "cdr",
    "cddr",
    "cons",
    "list",
    "null?",
    "eq?",
    "equal?",
    "quote-if",
    "cond",
    "+expand",
    "=expand",
    "<>expand",
    ">expand",
    "<expand",
    ">=expand",
    "<=expand",
    "-expand",
    "*expand",
    "quotient-expand",
    "mod-expand",
    "&expand",
    "characters-expand",
    "address-expand",
    "row-expand",
    "column-expand",
    "column-letters-expand",
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

/// The most forms one macro application's result may hold (VLA.1). The
/// port copies a macro's arguments into its template, so a macro that
/// doubles its argument on every application would run the memory out long
/// before the chain's 5,000 applications. The reference shares them by
/// reference, builds no copies, and meets the chain limit instead. One
/// application's result, not the program's sum: a walk over a long list
/// writes the rest of the list into every result, so a sum grows as the
/// square of the list. The largest real expanded form is 20,144 forms; the
/// table generator at the chain limit's 5,000 rows writes about 540,000 in
/// its first application.
pub const EXPANSION_SIZE_LIMIT: usize = 1_000_000;

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
        self.expand_at(d, depth, 1)
    }

    /// [`Expander::expand`] at `walk`, the depth of the list `d` would be
    /// within its top-level form (VLA.1): a list deeper than the reader's
    /// limit is refused, whether a text or a macro nested it so.
    fn expand_at(&mut self, d: &Form, depth: u32, walk: u32) -> Result<Form, Refusal> {
        if let Form::List(lst) = d {
            if walk > FORM_DEPTH_LIMIT {
                return Err(too_deep(lst));
            }
        }
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
            // LISTOPS: the twenty-seven primitives, standalone.
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
                self.weigh(&r)?;
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
            out.push(self.expand_at(e, depth, walk + 1)?);
        }
        let out = Form::List(List::new(out, cur_list.line));
        if pending_head.is_empty() {
            return Ok(out);
        }
        let mut final_out = vec![Form::sym("begin")];
        for e in &pending_head {
            final_out.push(self.expand_at(e, depth, walk + 1)?);
        }
        final_out.push(out);
        Ok(Form::list(final_out))
    }

    /// VBA `ExpandOnePass`: top-down; a list whose head is a macro is
    /// replaced by its substitution and not walked further.
    pub fn expand_one_pass(&mut self, d: &Form) -> Result<Form, Refusal> {
        self.expand_one_pass_at(d, 1)
    }

    /// [`Expander::expand_one_pass`] at `walk`, held to the reader's limit
    /// as [`Expander::expand`]'s walk is (VLA.1).
    fn expand_one_pass_at(&mut self, d: &Form, walk: u32) -> Result<Form, Refusal> {
        let Form::List(lst) = d else {
            return Ok(d.clone());
        };
        if walk > FORM_DEPTH_LIMIT {
            return Err(too_deep(lst));
        }
        if let Some(Form::Sym(head)) = lst.items.first() {
            if fold(head) == "quote" {
                return Ok(d.clone());
            }
            if let Some(rec) = self.get_macro(head).cloned() {
                if !(self.budget_on && self.expand_fired >= self.expand_budget) {
                    let r = expand_one(lst, &rec)?;
                    self.weigh(&r)?;
                    self.expand_fired += 1;
                    return Ok(r);
                }
            }
        }
        let mut out = Vec::with_capacity(lst.count());
        for e in &lst.items {
            out.push(self.expand_one_pass_at(e, walk + 1)?);
        }
        Ok(Form::List(List::new(out, lst.line)))
    }

    /// VLA.1: one application's result, weighed before it is used: no list
    /// in it nested past the reader's limit, and no more than
    /// [`EXPANSION_SIZE_LIMIT`] forms in it. So a macro that deepens its
    /// argument on every application is refused at the depth the walk would
    /// refuse it, before copying a form thousands deep runs the stack out,
    /// and one that doubles it is refused within twenty applications. The
    /// walk costs what the copy that built the result cost. A stack of its
    /// own, since the form is what may be deep.
    fn weigh(&self, r: &Form) -> Result<(), Refusal> {
        let mut forms = 0usize;
        let mut stack = vec![(r, 1u32)];
        while let Some((f, depth)) = stack.pop() {
            forms += 1;
            if forms > EXPANSION_SIZE_LIMIT {
                return Err(raise(
                    "vla-macro-expansion-too-large",
                    &[("limit", "1,000,000")],
                ));
            }
            if let Form::List(l) = f {
                if depth > FORM_DEPTH_LIMIT {
                    return Err(too_deep(l));
                }
                stack.extend(l.items.iter().map(|item| (item, depth + 1)));
            }
        }
        Ok(())
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

/// The walk's refusal of a list past the reader's limit, at the list's own
/// line, or at none for a list a macro made (VLA.1).
fn too_deep(lst: &List) -> Refusal {
    raise(
        "vla-form-too-deep",
        &[
            ("limit", &FORM_DEPTH_LIMIT.to_string()),
            ("loc", &src_line_tag(lst.line, 0)),
        ],
    )
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
        "&expand" => eval_concat_expand(lst, env),
        "characters-expand" => eval_characters_expand(lst, env),
        "address-expand" => eval_address_expand(lst, env),
        "row-expand" | "column-expand" => eval_address_part_expand(head_fold, lst, env),
        "column-letters-expand" => eval_column_letters_expand(lst, env),
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
/// numeric operands only, through `Val` and `Str$`. VLA.3 adds `-expand`
/// and `*expand`, and `quotient-expand` and `mod-expand`: Excel's
/// `QUOTIENT` (truncated toward zero) and `MOD` (the divisor's sign,
/// `a - b * floor(a / b)`), over whole numbers only, a zero divisor
/// refused.
fn eval_arith_expand(op: &str, lst: &List, env: &Env) -> Result<Form, Refusal> {
    if lst.count() != 3 {
        return Err(raise(
            "vla-arith-expand-arity",
            &[("op", op), ("n", &arity_n(lst))],
        ));
    }
    let a = resolve_listops_arg(lst.nth(2)?, env)?;
    let b = resolve_listops_arg(lst.nth(3)?, env)?;
    let [ta, tb] = number_texts(op, [&a, &b])?;
    let na = val(&ta);
    let nb = val(&tb);
    let number = |x: f64| Form::Sym(str_(x).trim().to_string());
    Ok(match op {
        "+expand" => number(na + nb),
        "=expand" => bool_sym(na == nb),
        "<>expand" => bool_sym(na != nb),
        ">expand" => bool_sym(na > nb),
        "<expand" => bool_sym(na < nb),
        ">=expand" => bool_sym(na >= nb),
        "<=expand" => bool_sym(na <= nb),
        "-expand" => number(na - nb),
        "*expand" => number(na * nb),
        _ => {
            for (n, text) in [(na, &ta), (nb, &tb)] {
                if n.fract() != 0.0 {
                    return Err(raise(
                        "vla-arith-expand-not-whole",
                        &[("op", op), ("value", text)],
                    ));
                }
            }
            if nb == 0.0 {
                return Err(raise("vla-arith-expand-divide-by-zero", &[("op", op)]));
            }
            if op == "quotient-expand" {
                number((na / nb).trunc())
            } else {
                number(na - nb * (na / nb).floor())
            }
        }
    })
}

/// The seven's gate over two operands, in the reference's order: both
/// checked as lists, then both as numeric literals. Their texts.
fn number_texts(op: &str, operands: [&Form; 2]) -> Result<[String; 2], Refusal> {
    let mut texts = [String::new(), String::new()];
    for (k, v) in operands.iter().enumerate() {
        let Some(text) = v.atom_text() else {
            return Err(raise("vla-arith-expand-not-number-list", &[("op", op)]));
        };
        texts[k] = text;
    }
    for text in &texts {
        if !is_numeric_literal_text(text) {
            return Err(raise(
                "vla-arith-expand-not-number-value",
                &[("op", op), ("value", text)],
            ));
        }
    }
    Ok(texts)
}

/// The same gate over one operand.
fn number_text(op: &str, v: &Form) -> Result<String, Refusal> {
    let Some(text) = v.atom_text() else {
        return Err(raise("vla-arith-expand-not-number-list", &[("op", op)]));
    };
    if !is_numeric_literal_text(&text) {
        return Err(raise(
            "vla-arith-expand-not-number-value",
            &[("op", op), ("value", &text)],
        ));
    }
    Ok(text)
}

/// VLA.3's arity refusal for a primitive taking `want` arguments.
fn expand_arity(op: &str, want: usize, lst: &List) -> Refusal {
    raise(
        "vla-expand-arity",
        &[
            ("op", op),
            ("want", &want.to_string()),
            ("n", &arity_n(lst)),
        ],
    )
}

/// VLA.3: an atom's text as the object model's primitives read it, a
/// string literal's content or a symbol's (a number's) own text.
fn atom_text_expand(v: &Form) -> String {
    match v {
        Form::Str(s) | Form::Sym(s) => s.clone(),
        Form::List(_) => String::new(),
    }
}

/// VLA.3: `(&expand a b ...)`, each argument resolved as every LISTOPS
/// argument is and its text joined. Always a string literal, never a
/// symbol, so a macro writes an address or a formula's text with no
/// leading-quote trick; `symbol` keeps that trick for identifiers. At
/// least one argument, as `symbol` takes.
fn eval_concat_expand(lst: &List, env: &Env) -> Result<Form, Refusal> {
    if lst.count() < 2 {
        return Err(raise("vla-concat-expand-arity", &[]));
    }
    let mut joined = String::new();
    for k in 2..=lst.count() {
        let piece = resolve_listops_arg(lst.nth(k)?, env)?;
        if piece.is_list() {
            return Err(raise(
                "vla-expand-arg-is-list",
                &[("op", "&expand"), ("n", &(k - 1).to_string())],
            ));
        }
        joined.push_str(&atom_text_expand(&piece));
    }
    Ok(Form::Str(joined))
}

/// VLA.3: `(characters-expand text)`, the text split into one-character
/// string literals in order. A character is a Unicode scalar value, as
/// the reference counts one by joining a UTF-16 surrogate pair. The empty
/// text gives the empty list.
fn eval_characters_expand(lst: &List, env: &Env) -> Result<Form, Refusal> {
    if lst.count() != 2 {
        return Err(expand_arity("characters-expand", 1, lst));
    }
    let v = resolve_listops_arg(lst.nth(2)?, env)?;
    if v.is_list() {
        return Err(raise(
            "vla-expand-arg-is-list",
            &[("op", "characters-expand"), ("n", "1")],
        ));
    }
    let items = atom_text_expand(&v)
        .chars()
        .map(|c| Form::Str(c.to_string()))
        .collect();
    Ok(Form::list(items))
}

/// VLA.3: `(address-expand row column)`, a cell's address, relative, as
/// Excel's `ADDRESS(row, column, 4)` writes it: `B3`. Each argument a
/// number literal under the seven's gate, then a whole number within
/// Excel's bounds. A string literal.
fn eval_address_expand(lst: &List, env: &Env) -> Result<Form, Refusal> {
    const OP: &str = "address-expand";
    if lst.count() != 3 {
        return Err(expand_arity(OP, 2, lst));
    }
    let a = resolve_listops_arg(lst.nth(2)?, env)?;
    let b = resolve_listops_arg(lst.nth(3)?, env)?;
    let [ta, tb] = number_texts(OP, [&a, &b])?;
    let (row, col) = (val(&ta), val(&tb));
    if row.fract() != 0.0 || !(1.0..=f64::from(MAX_ROW)).contains(&row) {
        return Err(raise(
            "vla-address-expand-row",
            &[("op", OP), ("value", &ta)],
        ));
    }
    if col.fract() != 0.0 || !(1.0..=f64::from(MAX_COLUMN)).contains(&col) {
        return Err(raise(
            "vla-address-expand-column",
            &[("op", OP), ("value", &tb)],
        ));
    }
    Ok(Form::Str(format!(
        "{}{}",
        column_letters(col as u32),
        str_(row).trim()
    )))
}

/// VLA.3: `(row-expand address)` and `(column-expand address)`, an
/// address's row and column numbers, as Excel's `ROW` and `COLUMN` of the
/// reference; and `(column-expand letters)`, a column's letters alone, the
/// reverse of `column-letters-expand`.
fn eval_address_part_expand(op: &str, lst: &List, env: &Env) -> Result<Form, Refusal> {
    if lst.count() != 2 {
        return Err(expand_arity(op, 1, lst));
    }
    let v = resolve_listops_arg(lst.nth(2)?, env)?;
    if v.is_list() {
        return Err(raise("vla-expand-arg-is-list", &[("op", op), ("n", "1")]));
    }
    let text = atom_text_expand(&v);
    let Some((row, col)) = parse_address_text(&text, op == "column-expand") else {
        return Err(raise(
            "vla-address-expand-not-address",
            &[("op", op), ("value", &text)],
        ));
    };
    let n = if op == "row-expand" { row } else { col };
    Ok(Form::Sym(str_(f64::from(n)).trim().to_string()))
}

/// VLA.3: one grammar for an address, the reference's `ParseAddressText`
/// rule for rule: an optional `$`, one to three ASCII letters in either
/// case, an optional `$`, then one to seven ASCII digits; the column at
/// most XFD, the row 1 to 1,048,576. With `letters_alone`, the letters with
/// their optional `$` and nothing after them are a column, the row 0.
/// Anything else, a sheet's name before it, a range, a space or a `$`
/// anywhere else, is not an address. (`sheet.rs`'s `parse_a1_cell` strips
/// a `$` wherever it stands, so it cannot be the rule both implementations
/// keep.)
fn parse_address_text(s: &str, letters_alone: bool) -> Option<(u32, u32)> {
    let b = s.as_bytes();
    let n = b.len();
    let i = usize::from(n >= 1 && b[0] == b'$');
    let mut j = i;
    while j < n && b[j].is_ascii_alphabetic() {
        j += 1;
    }
    if j == i || j - i > 3 {
        return None;
    }
    let col = letters_to_col(&s[i..j])?;
    if j == n {
        return letters_alone.then_some((0, col));
    }
    if b[j] == b'$' {
        j += 1;
    }
    let mut k = j;
    while k < n && b[k].is_ascii_digit() {
        k += 1;
    }
    if k < n || k == j || k - j > 7 {
        return None;
    }
    let row: u32 = s[j..k].parse().ok()?;
    if row == 0 || row > MAX_ROW {
        return None;
    }
    Some((row, col))
}

/// VLA.3: `(column-letters-expand column)`, a column's letters as a string
/// literal, 1 = A ... 16384 = XFD; the reverse of `column-expand`.
fn eval_column_letters_expand(lst: &List, env: &Env) -> Result<Form, Refusal> {
    const OP: &str = "column-letters-expand";
    if lst.count() != 2 {
        return Err(expand_arity(OP, 1, lst));
    }
    let v = resolve_listops_arg(lst.nth(2)?, env)?;
    let text = number_text(OP, &v)?;
    let col = val(&text);
    if col.fract() != 0.0 || !(1.0..=f64::from(MAX_COLUMN)).contains(&col) {
        return Err(raise(
            "vla-address-expand-column",
            &[("op", OP), ("value", &text)],
        ));
    }
    Ok(Form::Str(column_letters(col as u32)))
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

    const PRELUDE: &str = include_str!("../testdata/prelude.vla");

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

    // TestListopsObjectModel (VLA.3), pin for pin.
    #[test]
    fn the_expand_time_object_model_folds_and_refuses() {
        let refused = |src: &str, id: &str, text: &str| {
            let e = expand_err(src);
            assert_eq!((e.id.as_str(), e.text.as_str()), (id, text), "{src}");
        };
        // Arithmetic: -expand and *expand as +expand; quotient-expand and
        // mod-expand Excel's QUOTIENT and MOD over whole numbers.
        assert_eq!(one("(-expand 5 2)"), "3");
        assert_eq!(one("(-expand 2 5)"), "-3");
        assert_eq!(one("(-expand 2.5 1)"), "1.5");
        assert_eq!(one("(*expand 6 7)"), "42");
        assert_eq!(one("(*expand 1.5 2)"), "3");
        assert_eq!(one("(quotient-expand 7 2)"), "3");
        assert_eq!(one("(quotient-expand -7 2)"), "-3");
        assert_eq!(one("(mod-expand 7 3)"), "1");
        assert_eq!(one("(mod-expand -7 2)"), "1");
        assert_eq!(one("(mod-expand 7 -2)"), "-1");
        refused(
            "(quotient-expand 7.5 2)",
            "vla-arith-expand-not-whole",
            "quotient-expand: expected a whole number, got '7.5'",
        );
        refused(
            "(mod-expand 7 0)",
            "vla-arith-expand-divide-by-zero",
            "mod-expand: cannot divide by zero",
        );
        refused(
            "(-expand 1)",
            "vla-arith-expand-arity",
            "-expand expects exactly 2 arguments, got 1",
        );
        refused(
            "(*expand 2 \"3\")",
            "vla-arith-expand-not-number-value",
            "*expand: expected a number, got '\"3'",
        );
        // Text: &expand joins atoms into a string; characters-expand splits.
        assert_eq!(one("(&expand \"A\" 1)"), "\"A1\"");
        assert_eq!(one("(&expand x)"), "\"x\"");
        assert_eq!(
            one("(&expand \"=SUM(\" (address-expand 1 2) \":\" (address-expand 9 2) \")\")"),
            "\"=SUM(B1:B9)\""
        );
        refused(
            "(&expand)",
            "vla-concat-expand-arity",
            "&expand expects at least 1 argument, got 0",
        );
        refused(
            "(&expand \"a\" (quote (b)))",
            "vla-expand-arg-is-list",
            "&expand: argument 2 is a list, expected an atom or string literal",
        );
        assert_eq!(one("(characters-expand \"ab#\")"), "(\"a\" \"b\" \"#\")");
        assert_eq!(one("(characters-expand \"\")"), "()");
        assert_eq!(one("(car (cdr (characters-expand \"xyz\")))"), "\"y\"");
        assert_eq!(
            one("(car (cdr (characters-expand \"a\u{1F600}b\")))"),
            "\"\u{1F600}\""
        );
        assert_eq!(
            one("(null? (cdr (cdr (cdr (characters-expand \"a\u{1F600}b\")))))"),
            "true"
        );
        refused(
            "(characters-expand (quote (a)))",
            "vla-expand-arg-is-list",
            "characters-expand: argument 1 is a list, expected an atom or string literal",
        );
        refused(
            "(characters-expand \"a\" \"b\")",
            "vla-expand-arity",
            "characters-expand expects exactly 1 argument(s), got 2",
        );
        // Addresses, by one grammar, within Excel's bounds.
        assert_eq!(one("(address-expand 3 2)"), "\"B3\"");
        assert_eq!(one("(address-expand 1 16384)"), "\"XFD1\"");
        assert_eq!(one("(address-expand 1048576 27)"), "\"AA1048576\"");
        refused(
            "(address-expand 0 1)",
            "vla-address-expand-row",
            "address-expand: a row is a whole number from 1 to 1048576, got '0'",
        );
        refused(
            "(address-expand 1 16385)",
            "vla-address-expand-column",
            "address-expand: a column is a whole number from 1 to 16384, got '16385'",
        );
        refused(
            "(address-expand 1.5 1)",
            "vla-address-expand-row",
            "address-expand: a row is a whole number from 1 to 1048576, got '1.5'",
        );
        assert_eq!(one("(row-expand \"B3\")"), "3");
        assert_eq!(one("(column-expand \"$AA$10\")"), "27");
        assert_eq!(one("(column-expand \"xfd1\")"), "16384");
        assert_eq!(one("(column-expand \"AA\")"), "27");
        for (src, value) in [
            ("(row-expand \"AA\")", "AA"),
            ("(row-expand \"Sheet1!B3\")", "Sheet1!B3"),
            ("(column-expand \"B3:C4\")", "B3:C4"),
            ("(row-expand \"B0\")", "B0"),
        ] {
            let op = &src[1..src.find(' ').unwrap()];
            refused(
                src,
                "vla-address-expand-not-address",
                &format!("{op}: expected an address such as B3, got '{value}'"),
            );
        }
        assert_eq!(one("(column-letters-expand 28)"), "\"AB\"");
        refused(
            "(column-letters-expand 0)",
            "vla-address-expand-column",
            "column-letters-expand: a column is a whole number from 1 to 16384, got '0'",
        );
        // Inside a macro's template, the other recognition site.
        assert_eq!(
            one("(defmacro (below a) (address-expand (+expand (row-expand a) 1) (column-expand a))) (below \"C7\")"),
            "\"C8\""
        );
        assert_eq!(
            one(concat!(
                "(defmacro (picture r txt) (picture-from r 1 (characters-expand txt))) ",
                "(defmacro (picture-from r c m) (quote-if (null? m) (begin) (begin ",
                "(quote-if (eq? (car m) \"#\") (cell \"S\" (address-expand r c) 1) (begin)) ",
                "(picture-from r (+expand c 1) (cdr m))))) (picture 2 \".#.#\")"
            )),
            "(begin (begin) (cell \"S\" \"B2\" 1) (begin) (cell \"S\" \"D2\" 1) (begin))"
        );
        // Reserved: none of the ten can be a macro's name.
        for (word, desc) in [
            ("-expand", "subtraction"),
            ("address-expand", "address"),
            ("characters-expand", "text-splitting"),
        ] {
            refused(
                &format!("(defmacro ({word} x) x)"),
                "vla-defmacro-reserved-name",
                &format!("defmacro '{word}': '{word}' is the expand-time {desc} primitive and cannot be a macro name"),
            );
        }
    }

    /// Every name the reference reserves is a primitive both recognition
    /// sites know, and the other way round, the quote bundle aside.
    #[test]
    fn every_listops_name_is_reserved() {
        let reserved: Vec<&str> = RESERVED.iter().map(|(w, _)| *w).collect();
        let quote_bundle = [
            "quote",
            "quasiquote",
            "unquote",
            "unquote-splicing",
            "symbol",
        ];
        for name in LISTOPS {
            assert!(
                reserved.contains(name),
                "{name} is recognised but not reserved"
            );
        }
        for name in reserved {
            assert!(
                LISTOPS.contains(&name) || quote_bundle.contains(&name),
                "{name} is reserved but not recognised"
            );
        }
        assert_eq!(LISTOPS.len(), 27);
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
        const GOLDEN: &str = include_str!("../testdata/instructions_golden.vla");
        let (text, fired) = expand_text(GOLDEN, PRELUDE, true).unwrap();
        assert!(fired > 100, "{fired} applications");
        assert!(text.starts_with("(const hot-pink \"#FF69B4\")"));
    }

    #[test]
    fn a_macro_nesting_a_call_past_the_depth_limit_is_refused_by_the_walk() {
        // VLA.1, TestFormTooDeep's second pair: a ten-deep macro called
        // inside 54 lists expands to exactly 64; inside 55 it would reach
        // 65, which the walk refuses though the text is 56 deep. The list
        // past the limit is the macro's, so the refusal names no line.
        let deep10 = "(defmacro (deep10) (a (a (a (a (a (a (a (a (a (a 1)))))))))))\r\n";
        let around = |n: usize| format!("{deep10}{}(deep10){}", "(".repeat(n), ")".repeat(n));
        let (_, fired) = expand_text(&around(54), "", true).unwrap();
        assert_eq!(fired, 1);
        let r = expand_text(&around(55), "", true).unwrap_err();
        assert_eq!(r.id, "vla-form-too-deep");
        assert_eq!(
            r.text,
            "lists nested more than 64 deep - no program needs that many"
        );
        // The one-pass walk holds the same limit over the lists it walks.
        let mut ex = Expander::new();
        let deep = read_forms(&format!("{}{}", "(".repeat(64), ")".repeat(64))).unwrap();
        let wrapped = Form::list(vec![deep[0].clone()]);
        assert_eq!(
            ex.expand_one_pass(&wrapped).unwrap_err().id,
            "vla-form-too-deep"
        );
        assert_eq!(ex.expand(&wrapped, 0).unwrap_err().id, "vla-form-too-deep");
    }

    #[test]
    fn a_macro_that_deepens_or_doubles_its_argument_is_refused_by_name() {
        // VLA.1: each application's result is weighed, so a macro that
        // deepens its argument every time is refused once it nests past 64,
        // and one that doubles it once one application writes a million
        // forms. In the port, which copies arguments, the first ran the
        // stack out and the second the memory, before the 5000 chain limit
        // could stop either; the reference shares arguments, builds no
        // copies, and meets the chain limit (TestFormTooDeep's last pair).
        let deeper = expand_text("(defmacro (deeper x) (deeper (x))) (deeper 1)", "", true);
        assert_eq!(deeper.unwrap_err().id, "vla-form-too-deep");
        let dup = expand_text("(defmacro (dup x) (dup (x x))) (dup 1)", "", true).unwrap_err();
        assert_eq!(dup.id, "vla-macro-expansion-too-large");
        assert_eq!(
            dup.text,
            "one macro application wrote more than 1,000,000 forms (a macro that doubles its argument?)"
        );
        // One pass weighs the result of each application it makes too.
        let mut ex = Expander::new();
        let forms = read_forms("(defmacro (wrap x) (w (w (w x))))").unwrap();
        ex.collect(forms).unwrap();
        let deep = read_forms(&format!("(wrap {}1{})", "(".repeat(62), ")".repeat(62))).unwrap();
        assert_eq!(
            ex.expand_one_pass(&deep[0]).unwrap_err().id,
            "vla-form-too-deep"
        );
    }

    #[test]
    fn a_long_walk_writes_millions_of_forms_but_none_too_many_at_once() {
        // VLA.1, found by the owner's run of the reference's scale pins
        // (2026-10-10): a walk over a long list writes the rest of the list
        // into every application's result, so the results summed grow as
        // the square of the list. A budget over the program's sum refused
        // the 2,000-row walk TABLESPEC-SCALE expands; the limit is one
        // application's result, which a walk keeps to the list's length.
        let walker = "(defmacro (walk-rows lst) (quote-if (null? lst) (begin) \
                      (begin (debug-print (car lst)) (walk-rows (cdr lst)))))\r\n";
        let rows: Vec<String> = (1..=2000).map(|i| i.to_string()).collect();
        let program = format!(
            "{walker}(sub t () (walk-rows (quote ({}))))",
            rows.join(" ")
        );
        let (text, fired) = expand_text(&program, "", true).unwrap();
        assert_eq!(fired, 2001, "one application a row, and the empty end");
        assert!(text.contains("(debug-print 2000)"));
    }
}
