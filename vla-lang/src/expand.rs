//! The macro system: VLA.bas's `DefineMacro`, `ExpandMacros`, `ExpandOne`,
//! `Substitute` with the QUASIQUOTE bundle, the twenty-nine LISTOPS
//! primitives, `ExpandOnePass`, and `VlaExpandText` (PORT.5, slice 5d; the
//! expand-time object model's ten, VLA.3; `assoc-expand`, VLA.5; `list?`, a
//! word told from a list, VLA.15); and the
//! walk's two kinds of form: `formula-expand`, the formula syntax at expand
//! time, which expands its argument before writing it (`ExpandFormulaCall`,
//! VLA.4), and the loop, `each` and `each-index`, a template substituted
//! once for each element of a list (`EachFormResults`, VLA.5). And, since
//! VLA.6, a name at expand time, `(define name form)`: the walk passes a
//! define by as written, and once the form of the source it stands in has
//! expanded, [`Expander::expand`] takes the defines from that form's
//! statements, checks them and registers them (`DefinesTake`), so that a
//! word binds from the next form of the source on, and the walk puts a
//! define's form where its word stands as an expression; with a reference's
//! two parts, `sheet-expand` and `range-expand`, forms of the walk beside
//! `formula-expand` (`ExpandReferencePart`), and since VLA.15 its first
//! cell, `first-cell-expand`, the third.
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
//! keep their lines as they are carried in.
//!
//! The forms are the VBA's too (VLA.2). A [`Form`] is an owned tree, so a
//! copy of a list copies everything under it; VLA.bas passes a list by
//! reference and makes `cdr` a view (`VlaSlice`, since P-NTH), so a macro
//! that walks a list of N rows pays for N rows, not N squared. The expander
//! holds the forms it works on as nodes of its own, which share as the
//! reference shares: a parameter's value at each use, a quoted argument's
//! content, `car`, `cdr` and `cddr`, a call's arguments and the elements a
//! trampoline sets aside. What the reference copies, the items of `cons`,
//! `list`, the rest parameter and `unquote-splicing`, are copied here too,
//! as references to the same items. A text is read into nodes once and the
//! expansion written out as forms once.
//!
//! The refusals are the catalogue's, by id, in the same situations.

use std::collections::HashMap;
use std::sync::{Arc, OnceLock};

use crate::form::{sym_text, Form, List};
use crate::formula_syntax;
use crate::headtable;
use crate::intrinsics::{fold, is_numeric, is_numeric_literal_text, str_, val};
use crate::messages::{raise, Refusal};
use crate::printer::{write_datum, write_pretty};
use crate::reader::{count_lf, parse_all, src_line_tag, tokenize, FORM_DEPTH_LIMIT};
use crate::refers::{self, letters_to_col, Kind};
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
    // VLA.4: the formula syntax at expand time, an operation of the walk.
    ("formula-expand", "expand-time formula-syntax operation"),
    // VLA.5: a piece by its head, and the loop.
    ("assoc-expand", "expand-time entry-lookup primitive"),
    ("each", "expand-time loop form"),
    ("each-index", "expand-time counting loop form"),
    // VLA.6: a name at expand time, and a reference's two parts.
    ("define", "expand-time naming form"),
    ("sheet-expand", "expand-time reference-sheet operation"),
    ("range-expand", "expand-time reference-range operation"),
    // VLA.15: a word told from a list, and a reference's first cell.
    ("list?", "list-test primitive"),
    ("first-cell-expand", "expand-time reference-first-cell operation"),
];

/// The twenty-nine LISTOPS primitives `Substitute` and `ExpandMacros` both
/// recognise at a form's head: the original nine, LISTOPS-EXPAND's seven,
/// `cond`, the expand-time object model's ten (VLA.3), `assoc-expand`
/// (VLA.5) and `list?` (VLA.15).
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
    "assoc-expand",
    "list?",
];

fn is_listops(head_fold: &str) -> bool {
    LISTOPS.contains(&head_fold)
}

/// VLA.5, VBA `IsEachForm`: a loop, a list headed by `each` or
/// `each-index`. Asked of every list a walk meets, so it folds nothing:
/// `fold` maps A to Z alone, which an ASCII comparison without case is.
fn is_each_form(n: &Node) -> bool {
    match n {
        Node::List(l) => matches!(
            l.items().first(),
            Some(Node::Sym(s)) if s.eq_ignore_ascii_case("each") || s.eq_ignore_ascii_case("each-index")
        ),
        _ => false,
    }
}

/// VLA.6, VBA `IsDefineForm`: a define, a list headed by `define`, over a
/// form the walk has written out, folding nothing, as [`is_each_form`]
/// folds nothing.
fn is_define(f: &Form) -> bool {
    match f {
        Form::List(l) => {
            matches!(l.items.first(), Some(Form::Sym(s)) if s.eq_ignore_ascii_case("define"))
        }
        _ => false,
    }
}

/// A form as the expander holds it (VLA.2): VLA.bas's own model, where an
/// atom is a text and a list a `Collection` passed by reference. Cloning a
/// node clones a reference, never the tree under it.
#[derive(Clone, Debug)]
enum Node {
    Sym(Arc<str>),
    Str(Arc<str>),
    List(Seq),
}

/// A list: the items of a shared [`Run`] from `start` on, under a line of
/// its own. At `start` 0 it is the reference's `Collection`; past it, its
/// `VlaSlice`, the view `cdr` makes. Its weight is kept beside it, made
/// when it is made: `forms`, the forms in its tree as the printer would
/// write them, itself among them (a list shared twice is counted twice,
/// saturating), and `depth`, its nesting, itself 1.
#[derive(Clone, Debug)]
struct Seq {
    run: Arc<Run>,
    start: usize,
    line: u32,
    forms: u32,
    depth: u32,
}

/// The items a list was made with, shared by every view of them.
#[derive(Debug)]
struct Run {
    items: Vec<Node>,
    /// For each `k`, the deepest list among `items[k..]` (0 for none),
    /// made the first time a view of this run is made, so that a view's
    /// depth is read and not walked.
    deepest: OnceLock<Vec<u32>>,
}

impl Run {
    fn deepest_from(&self, k: usize) -> u32 {
        let deepest = self.deepest.get_or_init(|| {
            let mut v = vec![0; self.items.len() + 1];
            for i in (0..self.items.len()).rev() {
                v[i] = v[i + 1].max(self.items[i].weight().1);
            }
            v
        });
        deepest[k]
    }
}

impl Seq {
    fn new(items: Vec<Node>, line: u32) -> Seq {
        let (mut forms, mut deepest) = (1u32, 0u32);
        for item in &items {
            let (f, d) = item.weight();
            forms = forms.saturating_add(f);
            deepest = deepest.max(d);
        }
        Seq {
            run: Arc::new(Run {
                items,
                deepest: OnceLock::new(),
            }),
            start: 0,
            line,
            forms,
            depth: deepest.saturating_add(1),
        }
    }

    fn items(&self) -> &[Node] {
        &self.run.items[self.start..]
    }

    fn count(&self) -> usize {
        self.run.items.len() - self.start
    }

    /// VBA `ListTail`'s view: the items from position `n + 1` on, the same
    /// items, untagged. Its forms are this list's less the `n` it passes;
    /// a count that saturated says only "at least", so it stays saturated
    /// and the walk in [`Expander::weigh`] decides.
    fn tail(&self, n: usize) -> Seq {
        let passed = self.items()[..n]
            .iter()
            .fold(0u32, |sum, item| sum.saturating_add(item.weight().0));
        let start = self.start + n;
        Seq {
            run: Arc::clone(&self.run),
            start,
            line: 0,
            forms: if self.forms == u32::MAX {
                u32::MAX
            } else {
                self.forms - passed
            },
            depth: self.run.deepest_from(start).saturating_add(1),
        }
    }
}

impl Node {
    fn sym(text: &str) -> Node {
        Node::Sym(Arc::from(text))
    }

    /// A list with no line, as every list the macro system makes has none.
    fn list(items: Vec<Node>) -> Node {
        Node::List(Seq::new(items, 0))
    }

    /// A form read into nodes, every list keeping its line.
    fn of(f: &Form) -> Node {
        match f {
            Form::Sym(s) => Node::Sym(Arc::from(s.as_str())),
            Form::Str(s) => Node::Str(Arc::from(s.as_str())),
            Form::List(l) => Node::List(Seq::new(l.items.iter().map(Node::of).collect(), l.line)),
        }
    }

    /// The node written out as a form: a list shared in two places is two
    /// lists of the form, as the reference's printer writes it twice.
    fn form(&self) -> Form {
        match self {
            Node::Sym(s) => Form::Sym(s.to_string()),
            Node::Str(s) => Form::Str(s.to_string()),
            Node::List(l) => Form::List(List::new(
                l.items().iter().map(Node::form).collect(),
                l.line,
            )),
        }
    }

    /// The forms in the node's tree and its nesting: an atom is one form
    /// and no list.
    fn weight(&self) -> (u32, u32) {
        match self {
            Node::List(l) => (l.forms, l.depth),
            _ => (1, 0),
        }
    }

    fn is_list(&self) -> bool {
        matches!(self, Node::List(_))
    }

    /// [`Form::atom_text`]: a symbol's text, or a string literal's content
    /// behind its quote marker. `None` for a list.
    fn atom_text(&self) -> Option<String> {
        match self {
            Node::Sym(s) => Some(s.to_string()),
            Node::Str(s) => Some(format!("\"{s}")),
            Node::List(_) => None,
        }
    }

    /// VBA `ListHeadIs`: a list whose first element is a symbol folding
    /// equal to `name`.
    fn head_is(&self, name: &str) -> bool {
        match self {
            Node::List(l) => {
                matches!(l.items().first(), Some(Node::Sym(s)) if fold(s) == fold(name))
            }
            _ => false,
        }
    }
}

/// VBA `Nth` over a list's items: element `i`, 1-based, refusing past the
/// end ([`List::nth`]'s refusal).
fn nth(items: &[Node], i: usize) -> Result<&Node, Refusal> {
    if i == 0 || i > items.len() {
        return Err(raise("vla-form-missing-element", &[("n", &i.to_string())]));
    }
    Ok(&items[i - 1])
}

/// VBA `NthList`: element `i` as a list, refusing an atom there
/// ([`List::nth_list`]'s refusal).
fn nth_list(items: &[Node], i: usize) -> Result<&Seq, Refusal> {
    match nth(items, i)? {
        Node::List(l) => Ok(l),
        other => Err(raise(
            "vla-expected-list-at-position",
            &[
                ("n", &i.to_string()),
                ("value", &other.atom_text().unwrap_or_default()),
            ],
        )),
    }
}

/// The substitution environment of one macro application: `Substitute`'s
/// `bindings`, `restName` (folded) and `restItems`, the call's own items
/// borrowed, so that a loop's scope (VLA.5) copies no list; and the
/// expander's macros, which the reference's `GetMacro` reads from its
/// module, so that a loop refuses a list written as a macro's call. Empty
/// for a primitive met outside any template, but for the macros.
#[derive(Default)]
struct Env<'r> {
    bindings: HashMap<&'r str, Node>,
    rest_name: &'r str,
    rest_items: &'r [Node],
    macros: Option<&'r HashMap<String, Rule>>,
}

/// A macro as the expander applies it (VLA.2): its parameters folded and
/// its template read into nodes once, when it is defined.
struct Rule {
    params: Vec<String>,
    rest_name: String,
    template: Node,
}

/// The macro table and the application counters: `mMacros`, `mBudgetOn`,
/// `mExpandBudget` and `mExpandFired`.
#[derive(Default)]
pub struct Expander {
    macros: HashMap<String, Macro>,
    /// Each macro of `macros` as it is applied, under the same name.
    rules: HashMap<String, Rule>,
    /// L15: when on, applications beyond the budget do not fire.
    pub budget_on: bool,
    pub expand_budget: u64,
    /// L1: macro applications since the last reset.
    pub expand_fired: u64,
    /// VLA.4: whether the walk stands inside a `formula-expand`'s argument,
    /// where another one is refused.
    formula_open: bool,
    /// VLA.6, `mDefines`: the words a source's defines bind, folded, each its
    /// form as written, registered once the form of the source each stands
    /// in has expanded, so that a word binds from the next form on. Empty
    /// for a source with none, which the walk's atom case asks first.
    defines: HashMap<String, Node>,
    /// VLA.6, `mDefineMet`: the walk passed a define by in the form
    /// [`Expander::expand`] expands.
    define_met: bool,
    /// VLA.6, `mFormulaTexts`: the texts formula-expand wrote in that form,
    /// read for a word used in the form that defines it.
    formula_texts: Vec<String>,
}

/// The most forms one macro application's result may hold (VLA.1), counted
/// as the printer would write them, so a value its template uses twice
/// counts twice. A macro that doubles its argument at every application is
/// refused at its nineteenth, long before the chain's 5,000: its value is
/// shared and cheap to hold (VLA.2), but its expansion is written out as a
/// tree. One application's result, not the program's sum: a walk over a
/// long list holds the rest of the list in every result, so a sum grows as
/// the square of the list. The largest form the corpus expands to is 20,144
/// forms; the table generator's first application holds 30,041 at 2,499
/// rows, the most its chain allows at two links a row.
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
        self.rules.insert(
            mfold.clone(),
            Rule {
                params: params.iter().map(|p| fold(p)).collect(),
                rest_name: fold(&rest_name),
                template: Node::of(&template),
            },
        );
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

    /// VBA `ExpandSourceForm`: a form of the source to its fixpoint
    /// (`ExpandMacros`), and then, when the walk passed a define by, the
    /// form's defines taken from its statements and registered for the forms
    /// after it (VLA.6). Every caller expands a source's forms through this,
    /// once a form and in order, so a word binds from the next form of the
    /// source on, whatever order the trampoline expanded this one in.
    pub fn expand(&mut self, d: &Form, depth: u32) -> Result<Form, Refusal> {
        self.define_met = false;
        self.formula_texts.clear();
        let out = self.expand_at(&Node::of(d), depth, 1)?;
        if self.define_met {
            return self.take_defines(out);
        }
        Ok(out)
    }

    /// [`Expander::expand`] at `walk`, the depth of the list `d` would be
    /// within its top-level form (VLA.1): a list deeper than the reader's
    /// limit is refused, whether a text or a macro nested it so.
    fn expand_at(&mut self, d: &Node, depth: u32, walk: u32) -> Result<Form, Refusal> {
        if let Node::List(lst) = d {
            if walk > FORM_DEPTH_LIMIT {
                return Err(too_deep(lst.line));
            }
        }
        if depth > 200 {
            return Err(raise("vla-macro-expansion-too-deep", &[]));
        }
        let mut cur = d.clone();
        let mut chain_len: u32 = 0;
        let mut pending_head: Vec<Node> = Vec::new();
        loop {
            // VLA.6: a word a define binds, its form put where the word stands
            // and looked at again, one link of the chain. A source with no
            // define pays one `is_empty` and folds nothing.
            if let Node::Sym(s) = &cur {
                if !self.defines.is_empty() {
                    if let Some(v) = self.defines.get(&fold(s)) {
                        chain_len += 1;
                        if chain_len > 5000 {
                            return Err(raise("vla-macro-self-expansion-limit", &[]));
                        }
                        cur = v.clone();
                        continue;
                    }
                }
            }
            let Node::List(lst) = &cur else {
                // VLA.5: an atom with forms set aside before it is the last
                // form of a begin a macro made, combined with them below.
                if pending_head.is_empty() {
                    return Ok(cur.form());
                }
                break;
            };
            let items = lst.items();
            let Some(Node::Sym(head)) = items.first() else {
                break;
            };
            let head_fold = fold(head);
            // P.L5: quote is a data context; expansion never enters. VLA.6:
            // every return of a form of its own, this one and those below,
            // goes by `finish`, which keeps the forms set aside before it.
            if head_fold == "quote" {
                let quoted = cur.form();
                return self.finish(&pending_head, quoted, depth, walk);
            }
            // QUASIQUOTE: (symbol ...) needs no ambient bindings.
            if head_fold == "symbol" {
                let fused = fuse_symbol(items, &self.bare_env())?.form();
                return self.finish(&pending_head, fused, depth, walk);
            }
            // VLA.4: the formula syntax, its argument expanded first.
            if head_fold == "formula-expand" {
                let text = self.formula_expand(items, depth, walk)?;
                return self.finish(&pending_head, text, depth, walk);
            }
            // VLA.6: a define, a form of the walk that writes nothing here:
            // passed by as written, its word and its form never walked, for
            // `expand` to take once the form of the source it stands in has
            // expanded; refused inside a formula's expression.
            if head_fold == "define" {
                self.define_passed_by(items)?;
                let passed = cur.form();
                return self.finish(&pending_head, passed, depth, walk);
            }
            // VLA.6: a reference's sheet and its cell or range as texts, forms
            // of the walk beside formula-expand, their argument expanded first
            // so that a word a define binds is replaced.
            if head_fold == "sheet-expand" || head_fold == "range-expand" {
                let part = self.reference_part(&head_fold, items, depth, walk)?;
                return self.finish(&pending_head, part, depth, walk);
            }
            // VLA.15: a reference's first cell as a text, a third form of the
            // walk beside the two parts and read as they are.
            if head_fold == "first-cell-expand" {
                let part = self.reference_part(&head_fold, items, depth, walk)?;
                return self.finish(&pending_head, part, depth, walk);
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
            // LISTOPS: the twenty-eight primitives, standalone.
            if is_listops(&head_fold) {
                let result = eval_listops_prim(&head_fold, items, &self.bare_env())?.form();
                self.define_note_chosen(&result)?;
                return self.finish(&pending_head, result, depth, walk);
            }
            // VLA.5: a loop at a form's head, outside any template, is a
            // statement: its results as one begin, handed back to this loop
            // so that its rows expand, where a primitive's result above is
            // returned as it is. One link of the chain, as a begin unwrapped
            // is.
            if head_fold == "each" || head_fold == "each-index" {
                chain_len += 1;
                if chain_len > 5000 {
                    return Err(raise("vla-macro-self-expansion-limit", &[]));
                }
                let next = each_begin(items, &self.bare_env())?;
                cur = next;
                continue;
            }
            if head_fold == "begin" && items.len() > 1 && lst.line == 0 {
                // TABLESPEC-SCALE: a begin a macro made is a tail-position
                // wrapper; its other elements wait, its last is looked at
                // again.
                pending_head.extend(items[1..items.len() - 1].iter().cloned());
                chain_len += 1;
                if chain_len > 5000 {
                    return Err(raise("vla-macro-self-expansion-limit", &[]));
                }
                let last = items[items.len() - 1].clone();
                cur = last;
            } else if head_fold == "begin" {
                break;
            } else {
                let Some(rule) = self.rules.get(&head_fold) else {
                    break;
                };
                if self.budget_on && self.expand_fired >= self.expand_budget {
                    break; // L15
                }
                chain_len += 1;
                if chain_len > 5000 {
                    return Err(raise("vla-macro-self-expansion-limit", &[]));
                }
                let r = expand_one(items, rule, &self.rules)?;
                self.weigh(&r)?;
                self.expand_fired += 1; // L1: one per application
                cur = r;
            }
        }
        // VLA.5, found while building the loop and fixed: a begin a macro
        // made whose last form is an atom, `(defmacro (m) (debug-print 1)
        // x)` or a loop's results ending in one, reaches here with its other
        // forms set aside, where the atom alone was returned and they were
        // dropped. They are kept, in order, the atom last.
        let Node::List(cur_list) = &cur else {
            return self.pending_begin(&pending_head, cur.form(), depth, walk);
        };
        // No macro at the head: rebuild with expanded elements, keeping the
        // list's own line (C2); a loop among them spliced (VLA.5). VLA.6: a
        // list's head is never a word a define binds, so an atom there is
        // passed by, as the walk would have returned it.
        let mut out = Vec::with_capacity(cur_list.count());
        let mut elements = cur_list.items();
        if let Some(head) = elements.first().filter(|head| !head.is_list()) {
            out.push(head.form());
            elements = &elements[1..];
        }
        for e in elements {
            match e {
                Node::List(l) if is_each_form(e) => {
                    self.expand_each_into(&mut out, l.items(), depth, walk + 1)?;
                }
                _ => out.push(self.expand_at(e, depth, walk + 1)?),
            }
        }
        let out = Form::List(List::new(out, cur_list.line));
        if pending_head.is_empty() {
            return Ok(out);
        }
        self.pending_begin(&pending_head, out, depth, walk)
    }

    /// TABLESPEC-SCALE and VLA.5, VBA `PendingBegin`: the forms the
    /// trampoline set aside, each expanded as a chain of its own (a loop
    /// among them spliced), then the last form, already expanded: one
    /// `begin`, with no line.
    fn pending_begin(
        &mut self,
        pending_head: &[Node],
        last: Form,
        depth: u32,
        walk: u32,
    ) -> Result<Form, Refusal> {
        let mut final_out = vec![Form::sym("begin")];
        for e in pending_head {
            match e {
                Node::List(l) if is_each_form(e) => {
                    self.expand_each_into(&mut final_out, l.items(), depth, walk + 1)?;
                }
                _ => final_out.push(self.expand_at(e, depth, walk + 1)?),
            }
        }
        final_out.push(last);
        Ok(Form::list(final_out))
    }

    /// VLA.6, `ExpandMacros`' `finish`: a form of its own the trampoline
    /// returns as it is (a quote, `symbol`'s atom, formula-expand's text, a
    /// primitive's result, a define passed by, a reference's part), after
    /// the forms it set aside before it, as an atom's are since VLA.5. Found
    /// while scoping and fixed: `(defmacro (m) (debug-print 1) (quote (a
    /// b)))` expanded `(m)` to `(quote (a b))` alone.
    fn finish(
        &mut self,
        pending_head: &[Node],
        result: Form,
        depth: u32,
        walk: u32,
    ) -> Result<Form, Refusal> {
        if pending_head.is_empty() {
            return Ok(result);
        }
        self.pending_begin(pending_head, result, depth, walk)
    }

    /// VLA.6, VBA `DefinePassedBy`: a define the walk meets at a form's
    /// head, passed by as written for [`Expander::expand`] to take. Inside a
    /// formula's expression it would be written as a function's call, so it
    /// is refused there.
    fn define_passed_by(&mut self, items: &[Node]) -> Result<(), Refusal> {
        if self.formula_open {
            let word = items
                .get(1)
                .map(|w| write_datum(&w.form()))
                .unwrap_or_default();
            return Err(raise("vla-define-not-a-statement", &[("word", &word)]));
        }
        self.define_met = true;
        Ok(())
    }

    /// VLA.6, VBA `DefineNoteChosen`: a standalone primitive's chosen branch
    /// is not expanded again (VLA.5, left as found), so a define it chooses
    /// is met at no head: noted here, as [`Expander::define_passed_by`]
    /// notes one, and refused inside a formula's expression.
    fn define_note_chosen(&mut self, result: &Form) -> Result<(), Refusal> {
        let Some(held) = define_held_in(result) else {
            return Ok(());
        };
        if self.formula_open {
            return Err(raise(
                "vla-define-not-a-statement",
                &[("word", &define_word_of(held))],
            ));
        }
        self.define_met = true;
        Ok(())
    }

    /// VLA.6, VBA `ExpandReferencePart`: `(sheet-expand reference)` and
    /// `(range-expand reference)`, forms of the walk beside formula-expand:
    /// the argument expanded to its fixpoint, so that a word a define binds
    /// is replaced first, then read as one reference on a sheet, and its
    /// sheet's name or its cell or range answered as a text, the two texts
    /// `ref` takes; a string, as formula-expand's is. VLA.15: and
    /// `(first-cell-expand reference)`, the third, its first cell as the
    /// text `ref` takes (VBA `FsFirstCellText`).
    fn reference_part(
        &mut self,
        op: &str,
        items: &[Node],
        depth: u32,
        walk: u32,
    ) -> Result<Form, Refusal> {
        if items.len() != 2 {
            return Err(expand_arity(op, 1, items));
        }
        let read = self.expand_at(&items[1], depth, walk + 1)?;
        let r = formula_syntax::sheet_reference(&read, op)?;
        Ok(Form::Str(match op {
            "sheet-expand" => r.sheet_name,
            "first-cell-expand" => formula_syntax::first_cell_text(&r, &read, op)?,
            _ => formula_syntax::part_text(&r),
        }))
    }

    /// VLA.5, VBA `ExpandEachInto`: a loop met among a list's elements by
    /// the walk, outside any template: its results, the loop evaluated with
    /// no bindings but its own names, pushed in order, each expanded at the
    /// depth the loop stood at.
    fn expand_each_into(
        &mut self,
        out: &mut Vec<Form>,
        each_form: &[Node],
        depth: u32,
        walk: u32,
    ) -> Result<(), Refusal> {
        let results = each_results(each_form, &self.bare_env())?;
        for r in &results {
            out.push(self.expand_at(r, depth, walk)?);
        }
        Ok(())
    }

    /// The environment of a form met outside any template: no bindings, but
    /// the macros, which a loop asks of its list (VLA.5).
    fn bare_env(&self) -> Env<'_> {
        Env {
            macros: Some(&self.rules),
            ..Env::default()
        }
    }

    /// VLA.4: `(formula-expand <expression>)`, VBA `ExpandFormulaCall`: the
    /// expression expanded to its fixpoint, then written by the formula
    /// syntax as one text, the formula's with its `=`. Not one of the
    /// LISTOPS primitives, which take their arguments as data and are
    /// computed at substitution too: `Substitute` leaves this form alone, so
    /// a template's parameters are filled in inside it, and the walk expands
    /// its argument before writing it, so that an author's macros inside the
    /// expression fire and a macro's call is never written as a function's.
    /// One inside another's argument is refused, since its text would be a
    /// text there and not the formula it wrote.
    fn formula_expand(&mut self, items: &[Node], depth: u32, walk: u32) -> Result<Form, Refusal> {
        if items.len() != 2 {
            return Err(raise(
                "vla-formula-expand-arity",
                &[("n", &(items.len() - 1).to_string())],
            ));
        }
        if self.formula_open {
            return Err(raise("vla-formula-expand-nested", &[]));
        }
        self.formula_open = true;
        let expr = self.expand_at(&items[1], depth, walk + 1);
        self.formula_open = false;
        let text = formula_syntax::write(&expr?)?;
        // VLA.6: kept for the form `expand` expands, so that a word a define
        // of that form binds, written here as one of Excel's names, is
        // refused (`DefinesRefuseFormulaUse`).
        self.formula_texts.push(text.clone());
        Ok(Form::Str(text))
    }

    /// VBA `ExpandOnePass`: top-down; a list whose head is a macro is
    /// replaced by its substitution and not walked further.
    pub fn expand_one_pass(&mut self, d: &Form) -> Result<Form, Refusal> {
        self.expand_one_pass_at(&Node::of(d), 1)
    }

    /// [`Expander::expand_one_pass`] at `walk`, held to the reader's limit
    /// as [`Expander::expand`]'s walk is (VLA.1).
    fn expand_one_pass_at(&mut self, d: &Node, walk: u32) -> Result<Form, Refusal> {
        let Node::List(lst) = d else {
            return Ok(d.form());
        };
        if walk > FORM_DEPTH_LIMIT {
            return Err(too_deep(lst.line));
        }
        let items = lst.items();
        if let Some(Node::Sym(head)) = items.first() {
            let head_fold = fold(head);
            // VLA.6: a define is left whole as written, as a quote is, since it
            // belongs to the fixpoint walk; one pass replaces no word.
            if head_fold == "quote" || head_fold == "define" {
                return Ok(d.form());
            }
            if let Some(rule) = self.rules.get(&head_fold) {
                if !(self.budget_on && self.expand_fired >= self.expand_budget) {
                    let r = expand_one(items, rule, &self.rules)?;
                    self.weigh(&r)?;
                    self.expand_fired += 1;
                    return Ok(r.form());
                }
            }
        }
        let mut out = Vec::with_capacity(items.len());
        for e in items {
            out.push(self.expand_one_pass_at(e, walk + 1)?);
        }
        Ok(Form::List(List::new(out, lst.line)))
    }

    /// VLA.1: one application's result, weighed before it is used: no list
    /// in it nested past the reader's limit, and no more than
    /// [`EXPANSION_SIZE_LIMIT`] forms in it. So a macro that deepens its
    /// argument on every application is refused at the depth the walk would
    /// refuse it, before a form thousands deep runs the stack out, and one
    /// that doubles it is refused within twenty applications. The result's
    /// weight was made with it (VLA.2), so a walk that holds the rest of
    /// its list in every result does not walk that rest at every step; only
    /// a result past a limit is walked, a stack of its own since the form
    /// is what may be deep, and in the order the copies were walked, so
    /// the refusal and the list it names are the same.
    fn weigh(&self, r: &Node) -> Result<(), Refusal> {
        let (forms, depth) = r.weight();
        if forms as usize <= EXPANSION_SIZE_LIMIT && depth <= FORM_DEPTH_LIMIT {
            return Ok(());
        }
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
            if let Node::List(l) = f {
                if depth > FORM_DEPTH_LIMIT {
                    return Err(too_deep(l.line));
                }
                stack.extend(l.items().iter().map(|item| (item, depth + 1)));
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

    /// VLA.6, VBA `DefinesTake`: a form of the source, expanded, with its
    /// defines taken from it. Those among its statements, the form itself or
    /// a begin's elements at any depth of begins, are taken in order, each
    /// leaving nothing where it stood; one anywhere else is refused. Each is
    /// checked; then a word of theirs used in this same form, where it is
    /// not yet bound, is refused, found in the result or in a text
    /// formula-expand wrote, and so is a ring of words; and they bind the
    /// forms of the source after this one.
    fn take_defines(&mut self, out: Form) -> Result<Form, Refusal> {
        let mut found = Vec::new();
        let kept = defines_from_statement(out, &mut found);
        refuse_stray_define(&kept)?;
        let mut taken = TakenDefines::default();
        for d in &found {
            self.take_one_define(d, &mut taken)?;
        }
        refuse_define_use(&kept, &taken)?;
        self.refuse_formula_use(&taken)?;
        for (word, _) in &taken.words {
            self.refuse_define_ring(word, &taken)?;
        }
        for (word, form) in taken.words {
            self.defines.insert(word, Node::of(&form));
        }
        Ok(kept)
    }

    /// VLA.6, VBA `DefineTakeOne`: one define taken from its form: its
    /// shape, `(define name form)`, the name a symbol and the form holding
    /// no define; its name held to the rules; then the word, folded, and its
    /// form as written into `taken`.
    fn take_one_define(&self, d: &Form, taken: &mut TakenDefines) -> Result<(), Refusal> {
        let Form::List(l) = d else {
            unreachable!("a define is a list headed by define")
        };
        let shape = || raise("vla-define-shape", &[("form", &write_datum(d))]);
        if l.items.len() != 3 {
            return Err(shape());
        }
        let Form::Sym(word) = &l.items[1] else {
            return Err(shape());
        };
        refuse_stray_define(&l.items[2])?;
        self.check_define_name(word, taken)?;
        taken.push(fold(word), l.items[2].clone());
        Ok(())
    }

    /// VLA.6, VBA `DefineCheckName`: a define's word, refused by name when
    /// the reference scanner reads it whole as a reference or a name on a
    /// sheet, or it has the R1C1 shape Excel refuses a name; when it is a
    /// number, a truth value, a keyword, a form of the head table, an
    /// expand-time name or a form of the formula syntax's own; and when it
    /// is a macro's name or defined already, in an earlier form or this one.
    fn check_define_name(&self, word: &str, taken: &TakenDefines) -> Result<(), Refusal> {
        if let Some(what) = define_reads_as(word) {
            return Err(raise(
                "vla-define-name-reference",
                &[("word", word), ("what", what)],
            ));
        }
        if let Some(desc) = define_reserved_desc(word) {
            return Err(raise(
                "vla-define-name-reserved",
                &[("word", word), ("desc", &desc)],
            ));
        }
        if self.get_macro(word).is_some() {
            return Err(raise(
                "vla-define-name-taken",
                &[("word", word), ("what", "a macro's name")],
            ));
        }
        if self.define_bound(&fold(word), taken) {
            return Err(raise(
                "vla-define-name-taken",
                &[("word", word), ("what", "defined already")],
            ));
        }
        Ok(())
    }

    /// VLA.6, VBA `DefineBound`: whether a folded word is bound by a define,
    /// of this form or of an earlier one.
    fn define_bound(&self, word_fold: &str, taken: &TakenDefines) -> bool {
        taken.has(word_fold) || self.defines.contains_key(word_fold)
    }

    /// VLA.6, VBA `DefinesRefuseFormulaUse`: a word a define of this form
    /// binds, written by formula-expand into one of this form's formulas as
    /// one of Excel's names, since it was not yet bound there: each text
    /// formula-expand wrote in the form read again by the scanner, and a name
    /// with no sheet that is such a word refused.
    fn refuse_formula_use(&self, taken: &TakenDefines) -> Result<(), Refusal> {
        for text in &self.formula_texts {
            for r in refers::scan(text) {
                if r.kind == Kind::Name && r.sheet_name.is_empty() && taken.has(&fold(&r.part)) {
                    return Err(raise("vla-define-used-in-its-form", &[("word", &r.part)]));
                }
            }
        }
        Ok(())
    }

    /// VLA.6, VBA `DefinesRefuseRing`: a ring of words, a define whose form
    /// comes back to its own word through the words its forms stand for,
    /// refused when its last define registers, the ring named. The words a
    /// define's form would have replaced are followed once each, breadth
    /// first, through every word a define binds, in an earlier form or this
    /// one; the ring is the way back.
    fn refuse_define_ring(&self, start: &str, taken: &TakenDefines) -> Result<(), Refusal> {
        let mut came_from: HashMap<String, String> = HashMap::new();
        let mut queue = vec![start.to_string()];
        let mut next = 0;
        while next < queue.len() {
            let at_name = queue[next].clone();
            next += 1;
            let mut words = Vec::new();
            match taken.get(&at_name) {
                Some(form) => define_words_in(form, &mut words),
                None => define_words_in(&self.defines[&at_name].form(), &mut words),
            }
            for word in words {
                if word == start {
                    return Err(raise(
                        "vla-define-cycle",
                        &[
                            ("word", start),
                            ("ring", &ring_text(start, &at_name, &came_from)),
                        ],
                    ));
                }
                if self.define_bound(&word, taken) && !came_from.contains_key(&word) {
                    came_from.insert(word.clone(), at_name.clone());
                    queue.push(word);
                }
            }
        }
        Ok(())
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
fn too_deep(line: u32) -> Refusal {
    raise(
        "vla-form-too-deep",
        &[
            ("limit", &FORM_DEPTH_LIMIT.to_string()),
            ("loc", &src_line_tag(line, 0)),
        ],
    )
}

/// VBA `ExpandOne`: one application of a macro to a call form's items. A
/// parameter is bound to its argument itself, not a copy of it, and the
/// rest parameter to the call's own items after the parameters'.
fn expand_one(call: &[Node], rule: &Rule, rules: &HashMap<String, Rule>) -> Result<Node, Refusal> {
    let argc = call.len() - 1;
    let name = call[0].atom_text().unwrap_or_default();
    if rule.rest_name.is_empty() {
        if argc != rule.params.len() {
            return Err(raise(
                "vla-macro-arity",
                &[
                    ("name", &name),
                    ("n", &rule.params.len().to_string()),
                    ("given", &argc.to_string()),
                ],
            ));
        }
    } else if argc < rule.params.len() {
        return Err(raise(
            "vla-macro-arity-min",
            &[
                ("name", &name),
                ("n", &rule.params.len().to_string()),
                ("given", &argc.to_string()),
            ],
        ));
    }
    let mut env = Env {
        bindings: HashMap::new(),
        rest_name: &rule.rest_name,
        rest_items: &call[rule.params.len() + 1..],
        macros: Some(rules),
    };
    for (j, p) in rule.params.iter().enumerate() {
        env.bindings.insert(p, call[j + 1].clone());
    }
    substitute(&rule.template, &env, false)
}

/// VBA `Substitute`: a template with its parameters replaced, each by its
/// value itself (`Set Substitute = v`). Inside a quasiquote shield
/// (`in_quasi`) a bare symbol is data, and only an explicit `unquote` or
/// `unquote-splicing` reaches the bindings.
fn substitute(t: &Node, env: &Env, in_quasi: bool) -> Result<Node, Refusal> {
    let lst = match t {
        Node::List(l) => l,
        Node::Sym(s) if !in_quasi => {
            let key = fold(s);
            if !env.rest_name.is_empty() && key == env.rest_name {
                // The rest parameter standalone: the whole list.
                return Ok(Node::list(env.rest_items.to_vec()));
            }
            if let Some(v) = env.bindings.get(key.as_str()) {
                return Ok(v.clone());
            }
            return Ok(t.clone());
        }
        _ => return Ok(t.clone()),
    };
    let items = lst.items();

    if let Some(Node::Sym(head)) = items.first() {
        let h = fold(head);
        match h.as_str() {
            "symbol" => return fuse_symbol(items, env),
            "quasiquote" => {
                if in_quasi {
                    return Err(raise("vla-quasiquote-nested", &[]));
                }
                if items.len() != 2 {
                    return Err(raise(
                        "vla-quasiquote-arity",
                        &[("n", &(items.len() - 1).to_string())],
                    ));
                }
                return substitute(nth(items, 2)?, env, true);
            }
            "unquote" => {
                if !in_quasi {
                    return Err(raise("vla-unquote-outside-quasiquote", &[]));
                }
                if items.len() != 2 {
                    return Err(raise(
                        "vla-unquote-arity",
                        &[("n", &(items.len() - 1).to_string())],
                    ));
                }
                // VLA.5: a value is read here, so a loop is its results'
                // list, as the rest parameter is its whole list here.
                return substitute_value(nth(items, 2)?, env);
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
                return eval_listops_prim(&h, items, env);
            }
            // VLA.5: a loop standing alone is a statement, its results one
            // begin, taken at its head before the element walk below, where
            // the rest parameter in its binding would be spliced. Shielded
            // data inside a quasiquote, as the primitives are.
            "each" | "each-index" if !in_quasi => {
                return each_begin(items, env);
            }
            _ => {}
        }
    }

    let mut out = Vec::with_capacity(items.len());
    for e in items {
        let mut spliced = false;
        if in_quasi && e.head_is("unquote-splicing") {
            let Node::List(spl) = e else { unreachable!() };
            let spl = spl.items();
            if spl.len() != 2 {
                return Err(raise(
                    "vla-unquote-splicing-arity",
                    &[("n", &(spl.len() - 1).to_string())],
                ));
            }
            // VLA.5: a value is read here too, so a loop splices its results.
            let val = substitute_value(nth(spl, 2)?, env)?;
            let Node::List(val) = val else {
                return Err(raise("vla-unquote-splicing-not-list", &[]));
            };
            out.extend(val.items().iter().cloned());
            spliced = true;
        } else if !in_quasi {
            match e {
                Node::Sym(s) if !env.rest_name.is_empty() && fold(s) == env.rest_name => {
                    // The rest parameter as an element: its forms spliced in.
                    out.extend(env.rest_items.iter().cloned());
                    spliced = true;
                }
                Node::List(l) if is_each_form(e) => {
                    // VLA.5: a loop as an element, its results spliced in.
                    out.extend(each_results(l.items(), env)?);
                    spliced = true;
                }
                _ => {}
            }
        }
        if !spliced {
            out.push(substitute(e, env, in_quasi)?);
        }
    }
    // S3.1: a template copy is explicitly untagged; the user's own forms
    // spliced in above keep their lines.
    Ok(Node::list(out))
}

/// VBA `FuseSymbol`: `(symbol a b ...)` fuses each argument's resolved
/// text into one new atom. A fused text that begins with a quote mark is
/// a string literal, as the VBA's leading-quote marker makes it.
fn fuse_symbol(lst: &[Node], env: &Env) -> Result<Node, Refusal> {
    if lst.len() < 2 {
        return Err(raise("vla-symbol-arity", &[]));
    }
    let mut fused = String::new();
    for k in 2..=lst.len() {
        let piece = substitute(nth(lst, k)?, env, false)?;
        match piece {
            Node::List(_) => {
                return Err(raise(
                    "vla-symbol-arg-is-list",
                    &[("n", &(k - 1).to_string())],
                ));
            }
            Node::Str(s) | Node::Sym(s) => fused.push_str(&s),
        }
    }
    Ok(match fused.strip_prefix('"') {
        Some(rest) => Node::Str(Arc::from(rest)),
        None => Node::Sym(Arc::from(fused.as_str())),
    })
}

// VLA.5: the loop at expand time, `(each (name list) form ...)` and
// `(each-index (index name list start) form ...)`. A form of the walks,
// beside `formula-expand`, and no LISTOPS primitive: its forms are a
// template substituted once for each element of the list, the name bound
// to the element, and its results take the rest parameter's own rule:
// spliced where the loop is a list's element, the list of them where a
// value is read, and one `begin` where it stands alone as a statement. No
// gensym and no eval: the names are the author's, and the loop reads
// nothing but the source.

/// VBA `EachFormBegin`: a loop standing alone, a statement (a macro's
/// whole template, a branch of `quote-if` or `cond`, a form of the source):
/// its results as one `begin`, with no line, so the trampoline unwraps it.
fn each_begin(each_form: &[Node], env: &Env) -> Result<Node, Refusal> {
    let mut out = vec![Node::sym("begin")];
    out.extend(each_results(each_form, env)?);
    Ok(Node::list(out))
}

/// VBA `SubstituteValue`: a form where a value is read, `unquote`'s or
/// `unquote-splicing`'s: a loop there is the list of its results, any other
/// form as `substitute` makes it.
fn substitute_value(t: &Node, env: &Env) -> Result<Node, Refusal> {
    match t {
        Node::List(l) if is_each_form(t) => Ok(Node::list(each_results(l.items(), env)?)),
        _ => substitute(t, env, false),
    }
}

/// VBA `EachFormResults`: a loop's results. The binding read by its shape,
/// each name it binds checked, the list resolved as a primitive's argument
/// is, each-index's start a whole number; then, for each element in order,
/// every form of the loop substituted as a list's element is (a loop among
/// them spliced, the rest parameter spliced), under the template's
/// bindings with the name bound to the element, and the index to its
/// count, the start up by one. The reference builds that scope from the
/// bindings its forms read, since a Collection cannot list its keys; the
/// port clones its map, which answers every form the same.
fn each_results(each_form: &[Node], env: &Env) -> Result<Vec<Node>, Refusal> {
    let op = match &each_form[0] {
        Node::Sym(s) => fold(s),
        _ => unreachable!("a loop is headed by its name"),
    };
    let binding = each_binding(each_form, &op)?;
    let indexed = op == "each-index";
    let (index_name, var_name, list_at) = if indexed {
        (Some(symbol_text(&binding[0])), symbol_text(&binding[1]), 2)
    } else {
        (None, symbol_text(&binding[0]), 1)
    };
    if let Some(index_name) = &index_name {
        each_check_name(&op, index_name, env)?;
    }
    each_check_name(&op, &var_name, env)?;
    let var_fold = fold(&var_name);
    let index_fold = index_name.as_deref().map(fold);
    if index_fold.as_deref() == Some(var_fold.as_str()) {
        return Err(raise(
            "vla-each-name-bound",
            &[("op", &op), ("word", &var_name)],
        ));
    }
    let walked = each_list(&op, &binding[list_at], env)?;
    let start = if indexed {
        each_start_value(&op, &binding[3], env)?
    } else {
        0.0
    };
    let mut scope = Env {
        bindings: env.bindings.clone(),
        rest_name: env.rest_name,
        rest_items: env.rest_items,
        macros: env.macros,
    };
    let mut results = Vec::new();
    for (k, element) in walked.items().iter().enumerate() {
        scope.bindings.insert(var_fold.as_str(), element.clone());
        if let Some(index_fold) = &index_fold {
            let count = start + k as f64;
            scope
                .bindings
                .insert(index_fold.as_str(), Node::sym(str_(count).trim()));
        }
        for f in &each_form[2..] {
            each_body_form_into(&mut results, f, &scope)?;
        }
    }
    Ok(results)
}

/// A name the loop's binding holds, which `each_binding` saw is a symbol.
fn symbol_text(n: &Node) -> String {
    match n {
        Node::Sym(s) => s.to_string(),
        _ => unreachable!("the binding's names are symbols"),
    }
}

/// VBA `EachBinding`: the loop's binding, or the refusal that shows its
/// shape: `(name list)` for `each`, `(index name list start)` for
/// `each-index`, each name a symbol, and at least one form after it.
fn each_binding<'a>(each_form: &'a [Node], op: &str) -> Result<&'a [Node], Refusal> {
    let want = if op == "each-index" { 4 } else { 2 };
    let names = want / 2;
    if each_form.len() >= 3 {
        if let Node::List(b) = &each_form[1] {
            let items = b.items();
            if items.len() == want && items[..names].iter().all(|n| matches!(n, Node::Sym(_))) {
                return Ok(items);
            }
        }
    }
    let shape = if want == 4 {
        "(index name list start)"
    } else {
        "(name list)"
    };
    Err(raise("vla-each-shape", &[("op", op), ("binding", shape)]))
}

/// VBA `EachCheckName`: a name a loop binds, refused when it is an
/// expand-time name (the table `defmacro` refuses a name by) or one the
/// template already binds, the macro's parameter, its rest, an enclosing
/// loop's name.
fn each_check_name(op: &str, word: &str, env: &Env) -> Result<(), Refusal> {
    let word_fold = fold(word);
    if let Some((reserved, desc)) = RESERVED.iter().find(|(w, _)| *w == word_fold) {
        return Err(raise(
            "vla-each-name-reserved",
            &[("op", op), ("word", reserved), ("desc", desc)],
        ));
    }
    let bound = env.bindings.contains_key(word_fold.as_str())
        || (!env.rest_name.is_empty() && word_fold == env.rest_name);
    if bound {
        return Err(raise("vla-each-name-bound", &[("op", op), ("word", word)]));
    }
    Ok(())
}

/// VBA `EachList`: the list a loop walks, resolved as a primitive's
/// argument is: a parameter, the rest parameter, a quoted list, a
/// primitive's result or another loop's results. A list written as a
/// macro's call is refused by name, since a primitive's argument is never
/// expanded and the call would be walked as its symbols; an atom is refused
/// as no list.
fn each_list(op: &str, list_form: &Node, env: &Env) -> Result<Seq, Refusal> {
    if let Node::List(l) = list_form {
        if let Some(Node::Sym(head)) = l.items().first() {
            if env.macros.is_some_and(|m| m.contains_key(&fold(head))) {
                return Err(raise("vla-each-list-a-call", &[("op", op), ("head", head)]));
            }
        }
    }
    match resolve_listops_arg(list_form, env)? {
        Node::List(walked) => Ok(walked),
        other => Err(raise(
            "vla-expand-expected-list",
            &[
                ("op", op),
                ("value", &other.atom_text().unwrap_or_default()),
            ],
        )),
    }
}

/// VBA `EachStartValue`: each-index's start, resolved as a primitive's
/// argument and held to the arithmetic's gate, a number literal and a
/// whole one, so that the count is exact.
fn each_start_value(op: &str, start_form: &Node, env: &Env) -> Result<f64, Refusal> {
    let v = resolve_listops_arg(start_form, env)?;
    let text = number_text(op, &v)?;
    let n = val(&text);
    if n.fract() != 0.0 {
        return Err(raise(
            "vla-arith-expand-not-whole",
            &[("op", op), ("value", &text)],
        ));
    }
    Ok(n)
}

/// VBA `EachBodyFormInto`: one of a loop's forms for one element, as a
/// list's element is substituted: a loop spliced, the rest parameter
/// spliced, any other form substituted under the element's scope.
fn each_body_form_into(out: &mut Vec<Node>, f: &Node, scope: &Env) -> Result<(), Refusal> {
    match f {
        Node::List(l) if is_each_form(f) => out.extend(each_results(l.items(), scope)?),
        Node::Sym(s) if !scope.rest_name.is_empty() && fold(s) == scope.rest_name => {
            out.extend(scope.rest_items.iter().cloned());
        }
        _ => out.push(substitute(f, scope, false)?),
    }
    Ok(())
}

// VLA.6: a name at expand time, `(define name form)`. A form of the walk that
// writes nothing where it stands: the walk passes it by as written, and once
// the form of the source it stands in has expanded, `Expander::expand` takes
// the defines from that form's statements, so that a word binds from the
// next form of the source on, whatever order the trampoline expanded the
// form in. The walk puts a define's form where its word stands as an
// expression, never at a list's head, never inside a quote, a text or a
// primitive's argument. No gensym and no eval: the word is the author's and
// the form the source's own, both gone at the pass.

/// VLA.6: the defines of one form of the source, in the order the form holds
/// them, each word folded with its form as written: `DefinesTake`'s `taken`
/// Dictionary, whose order is the order they register in.
#[derive(Default)]
struct TakenDefines {
    words: Vec<(String, Form)>,
    index: HashMap<String, usize>,
}

impl TakenDefines {
    fn has(&self, word_fold: &str) -> bool {
        self.index.contains_key(word_fold)
    }

    fn get(&self, word_fold: &str) -> Option<&Form> {
        self.index.get(word_fold).map(|&i| &self.words[i].1)
    }

    fn push(&mut self, word_fold: String, form: Form) {
        self.index.insert(word_fold.clone(), self.words.len());
        self.words.push((word_fold, form));
    }
}

/// VLA.6, VBA `DefinesFromStatement`: a statement with the defines among it
/// taken out, in order, into `found`: a define itself leaves an empty begin;
/// a begin is rebuilt without them, a begin among its elements the same,
/// keeping its line; any other form is kept whole.
fn defines_from_statement(v: Form, found: &mut Vec<Form>) -> Form {
    if is_define(&v) {
        found.push(v);
        return Form::list(vec![Form::sym("begin")]);
    }
    if !v.head_is("begin") {
        return v;
    }
    let Form::List(l) = v else {
        unreachable!("a begin is a list")
    };
    let mut items = l.items.into_iter();
    let mut out: Vec<Form> = items.next().into_iter().collect();
    for e in items {
        if is_define(&e) {
            found.push(e);
        } else {
            out.push(defines_from_statement(e, found));
        }
    }
    Form::List(List::new(out, l.line))
}

/// VLA.6, VBA `DefinesRefuseStray`: a define left anywhere in a form the
/// statements' walk kept whole, a row's argument, a call's, a procedure's
/// body, refused by name, since a define stands as a form of the source or a
/// statement of a begin standing as one; a quote's data passed over.
fn refuse_stray_define(v: &Form) -> Result<(), Refusal> {
    let Form::List(l) = v else {
        return Ok(());
    };
    if v.head_is("quote") {
        return Ok(());
    }
    if is_define(v) {
        return Err(raise(
            "vla-define-not-a-statement",
            &[("word", &define_word_of(l))],
        ));
    }
    for e in &l.items {
        refuse_stray_define(e)?;
    }
    Ok(())
}

/// VLA.6, VBA `DefineWordOf`: a define's word as a refusal names it: the
/// word, or what stands where the word belongs, as written.
fn define_word_of(l: &List) -> String {
    l.items.get(1).map(write_datum).unwrap_or_default()
}

/// VLA.6, VBA `DefineHeldIn`: the first define a form holds, itself or at
/// any depth, a quote's data passed over.
fn define_held_in(v: &Form) -> Option<&List> {
    let Form::List(l) = v else {
        return None;
    };
    if v.head_is("quote") {
        return None;
    }
    if is_define(v) {
        return Some(l);
    }
    l.items.iter().find_map(define_held_in)
}

/// VLA.6, VBA `DefinesRefuseUse`: a word a define of this form binds,
/// standing in the form's result where the walk would have put its form,
/// the result itself or a list's element after an atom head, never a
/// quote's, refused: in the form that defines it a word is not yet bound.
fn refuse_define_use(v: &Form, taken: &TakenDefines) -> Result<(), Refusal> {
    match v {
        Form::Sym(s) if taken.has(&fold(s)) => {
            Err(raise("vla-define-used-in-its-form", &[("word", s)]))
        }
        Form::List(l) if !v.head_is("quote") => {
            for (i, e) in l.items.iter().enumerate() {
                if i > 0 || matches!(e, Form::List(_)) {
                    refuse_define_use(e, taken)?;
                }
            }
            Ok(())
        }
        _ => Ok(()),
    }
}

/// VLA.6, VBA `DefineWordsIn`: the words a define's form would have replaced
/// where the walk meets them, the form itself when a symbol, a list's
/// elements after an atom head, never a quote's, folded, in order.
fn define_words_in(v: &Form, words: &mut Vec<String>) {
    match v {
        Form::Sym(s) => words.push(fold(s)),
        Form::List(l) if !v.head_is("quote") => {
            for (i, e) in l.items.iter().enumerate() {
                if i > 0 || matches!(e, Form::List(_)) {
                    define_words_in(e, words);
                }
            }
        }
        _ => {}
    }
}

/// VLA.6, VBA `DefineRingText`: a ring as its refusal names it, from its
/// first word around to it again, `start -> ... -> last -> start`, by the
/// word each was reached from.
fn ring_text(start: &str, last: &str, came_from: &HashMap<String, String>) -> String {
    let mut ring = format!("{last} -> {start}");
    let mut at = last.to_string();
    while at != start {
        at = came_from[&at].clone();
        ring = format!("{at} -> {ring}");
    }
    ring
}

/// VLA.6, VBA `DefineReadsAs`: what the reference scanner reads a define's
/// word as, or `None` for a word of the source's own: a cell, a range, a
/// column, a row, a span of sheets, a name on a sheet or any other reference
/// it reads whole; or a reference in R1C1 notation, `R`, `C`, `RC` or `R1C1`,
/// which Excel refuses as a name.
fn define_reads_as(word: &str) -> Option<&'static str> {
    let read = formula_syntax::whole_reference(word).and_then(|r| match r.kind {
        Kind::Cell => Some("a cell"),
        Kind::Range => Some("a range"),
        Kind::Column => Some("a column"),
        Kind::Row => Some("a row"),
        Kind::ThreeD => Some("a span of sheets"),
        Kind::Name if r.sheet_name.is_empty() => None,
        Kind::Name => Some("a name on a sheet"),
        _ => Some("a reference"),
    });
    read.or_else(|| refers::is_r1c1_shaped(word).then_some("a reference in R1C1 notation"))
}

/// VLA.6, VBA `DefineReservedDesc`: what a define's word already is in the
/// language, or `None` for a word of the source's own: a number, a truth
/// value, a keyword argument, a form of the head table or one of its
/// aliases, or `defmacro`; an expand-time name ([`RESERVED`]); or a form of
/// the formula syntax's own, the three marks among them since VLA.15.
fn define_reserved_desc(word: &str) -> Option<String> {
    let word_fold = fold(word);
    let desc = if formula_syntax::is_numeral(word)
        || formula_syntax::is_negative_numeral(word)
        || is_numeric(word)
    {
        "a number".to_string()
    } else if word_fold == "true" || word_fold == "false" {
        "a truth value".to_string()
    } else if word.starts_with(':') && word.chars().count() > 1 {
        "a keyword argument".to_string()
    } else if headtable::row(&word_fold).is_some() || word_fold == "defmacro" {
        "a form of the language".to_string()
    } else if let Some((_, desc)) = RESERVED.iter().find(|(w, _)| *w == word_fold) {
        format!("the {desc}")
    } else if matches!(
        word_fold.as_str(),
        "error" | "ref" | "last" | "absolute" | "absolute-row" | "absolute-column"
    ) {
        "a form of the formula syntax's own".to_string()
    } else {
        return None;
    };
    Some(desc)
}

/// VBA `EvalListopsPrim`: the dispatcher the two recognition sites share,
/// over the primitive's call form.
fn eval_listops_prim(head_fold: &str, lst: &[Node], env: &Env) -> Result<Node, Refusal> {
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
        "assoc-expand" => eval_assoc_expand(lst, env),
        "list?" => eval_listq(lst, env),
        _ => eval_arith_expand(head_fold, lst, env),
    }
}

/// VBA `ResolveListopsArg`: the argument through `Substitute`, then one
/// top-level `(quote X)` unwrapped to its data, X itself. A loop there is
/// the list of its results (VLA.5), as the rest parameter is its whole list.
fn resolve_listops_arg(arg: &Node, env: &Env) -> Result<Node, Refusal> {
    let v = match arg {
        Node::List(l) if is_each_form(arg) => Node::list(each_results(l.items(), env)?),
        _ => substitute(arg, env, false)?,
    };
    if v.head_is("quote") {
        let Node::List(qc) = &v else { unreachable!() };
        let qc = qc.items();
        if qc.len() != 2 {
            return Err(raise("vla-quote-arity", &[]));
        }
        return Ok(qc[1].clone());
    }
    Ok(v)
}

fn arity_n(lst: &[Node]) -> String {
    (lst.len() - 1).to_string()
}

fn eval_car(lst: &[Node], env: &Env) -> Result<Node, Refusal> {
    if lst.len() != 2 {
        return Err(raise("vla-car-arity", &[("n", &arity_n(lst))]));
    }
    let v = resolve_listops_arg(nth(lst, 2)?, env)?;
    let Node::List(c) = &v else {
        return Err(raise(
            "vla-car-expected-list",
            &[("value", &v.atom_text().unwrap_or_default())],
        ));
    };
    let Some(first) = c.items().first() else {
        return Err(raise("vla-car-empty-list", &[]));
    };
    Ok(first.clone())
}

/// VBA `ListTail`: the items from position `n + 1` on, untagged, a view
/// of the same items (`VlaSlice`).
fn list_tail(c: &Seq, n: usize, caller: &str) -> Result<Node, Refusal> {
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
    Ok(Node::List(c.tail(n)))
}

fn eval_cdr(lst: &[Node], env: &Env) -> Result<Node, Refusal> {
    if lst.len() != 2 {
        return Err(raise("vla-cdr-arity", &[("n", &arity_n(lst))]));
    }
    let v = resolve_listops_arg(nth(lst, 2)?, env)?;
    let Node::List(c) = &v else {
        return Err(raise(
            "vla-cdr-expected-list",
            &[("value", &v.atom_text().unwrap_or_default())],
        ));
    };
    list_tail(c, 1, "cdr")
}

fn eval_cddr(lst: &[Node], env: &Env) -> Result<Node, Refusal> {
    if lst.len() != 2 {
        return Err(raise("vla-cddr-arity", &[("n", &arity_n(lst))]));
    }
    let v = resolve_listops_arg(nth(lst, 2)?, env)?;
    let Node::List(c) = &v else {
        return Err(raise(
            "vla-cddr-expected-list",
            &[("value", &v.atom_text().unwrap_or_default())],
        ));
    };
    list_tail(c, 2, "cddr")
}

/// `(cons item lst)`: a new list of the item and the tail's items, as
/// VBA's `EvalCons` makes it, which shares no tail (P-NTH: a flat view
/// cannot hold one element ahead of a shared tail). The items themselves
/// are shared.
fn eval_cons(lst: &[Node], env: &Env) -> Result<Node, Refusal> {
    if lst.len() != 3 {
        return Err(raise("vla-cons-arity", &[("n", &arity_n(lst))]));
    }
    let item = resolve_listops_arg(nth(lst, 2)?, env)?;
    let tail = resolve_listops_arg(nth(lst, 3)?, env)?;
    let Node::List(tail) = &tail else {
        return Err(raise(
            "vla-cons-second-not-list",
            &[("value", &tail.atom_text().unwrap_or_default())],
        ));
    };
    let mut out = Vec::with_capacity(tail.count() + 1);
    out.push(item);
    out.extend(tail.items().iter().cloned());
    Ok(Node::list(out))
}

fn eval_list(lst: &[Node], env: &Env) -> Result<Node, Refusal> {
    let mut out = Vec::with_capacity(lst.len());
    for k in 2..=lst.len() {
        out.push(resolve_listops_arg(nth(lst, k)?, env)?);
    }
    Ok(Node::list(out))
}

fn bool_sym(b: bool) -> Node {
    Node::sym(if b { "true" } else { "false" })
}

/// VBA `AtomEqual`: symbols folded, strings exact, never across the kinds.
fn atom_equal(a: &Node, b: &Node) -> bool {
    match (a, b) {
        (Node::Sym(x), Node::Sym(y)) => fold(x) == fold(y),
        (Node::Str(x), Node::Str(y)) => x == y,
        _ => false,
    }
}

/// VBA `DeepEqual`.
fn deep_equal(a: &Node, b: &Node) -> bool {
    match (a, b) {
        (Node::List(x), Node::List(y)) => {
            x.count() == y.count()
                && x.items()
                    .iter()
                    .zip(y.items())
                    .all(|(p, q)| deep_equal(p, q))
        }
        (Node::List(_), _) | (_, Node::List(_)) => false,
        _ => atom_equal(a, b),
    }
}

fn eval_nullq(lst: &[Node], env: &Env) -> Result<Node, Refusal> {
    if lst.len() != 2 {
        return Err(raise("vla-nullq-arity", &[("n", &arity_n(lst))]));
    }
    let v = resolve_listops_arg(nth(lst, 2)?, env)?;
    Ok(match &v {
        Node::List(c) => bool_sym(c.count() == 0),
        _ => bool_sym(false),
    })
}

/// VLA.15, VBA `EvalListQ`: `(list? x)`, `true` when x resolves to a list,
/// the empty list among them and a view `cdr` made, and `false` for a
/// symbol or a text: the one test of a form's kind that refuses neither,
/// where `null?` is false for a word and for a list with items alike, `car`
/// refuses a word and `eq?` a list. Scheme's own `list?`.
fn eval_listq(lst: &[Node], env: &Env) -> Result<Node, Refusal> {
    if lst.len() != 2 {
        return Err(expand_arity("list?", 1, lst));
    }
    let v = resolve_listops_arg(nth(lst, 2)?, env)?;
    Ok(bool_sym(v.is_list()))
}

fn eval_eq(lst: &[Node], env: &Env) -> Result<Node, Refusal> {
    if lst.len() != 3 {
        return Err(raise("vla-eq-arity", &[("n", &arity_n(lst))]));
    }
    let a = resolve_listops_arg(nth(lst, 2)?, env)?;
    let b = resolve_listops_arg(nth(lst, 3)?, env)?;
    if a.is_list() || b.is_list() {
        return Err(raise("vla-eq-expected-atom", &[]));
    }
    Ok(bool_sym(atom_equal(&a, &b)))
}

fn eval_equal(lst: &[Node], env: &Env) -> Result<Node, Refusal> {
    if lst.len() != 3 {
        return Err(raise("vla-equal-arity", &[("n", &arity_n(lst))]));
    }
    let a = resolve_listops_arg(nth(lst, 2)?, env)?;
    let b = resolve_listops_arg(nth(lst, 3)?, env)?;
    Ok(bool_sym(deep_equal(&a, &b)))
}

/// `(quote-if test then else)`: only the selected branch is ever resolved.
fn eval_quote_if(lst: &[Node], env: &Env) -> Result<Node, Refusal> {
    if lst.len() != 4 {
        return Err(raise("vla-quote-if-arity", &[("n", &arity_n(lst))]));
    }
    let t = resolve_listops_arg(nth(lst, 2)?, env)?;
    let tf = match &t {
        Node::List(_) => return Err(raise("vla-quote-if-test-is-list", &[])),
        Node::Str(_) => {
            return Err(raise(
                "vla-quote-if-test-invalid",
                &[("value", &t.atom_text().unwrap_or_default())],
            ))
        }
        Node::Sym(s) => fold(s),
    };
    let branch = match tf.as_str() {
        "true" => nth(lst, 3)?,
        "false" => nth(lst, 4)?,
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
fn eval_arith_expand(op: &str, lst: &[Node], env: &Env) -> Result<Node, Refusal> {
    if lst.len() != 3 {
        return Err(raise(
            "vla-arith-expand-arity",
            &[("op", op), ("n", &arity_n(lst))],
        ));
    }
    let a = resolve_listops_arg(nth(lst, 2)?, env)?;
    let b = resolve_listops_arg(nth(lst, 3)?, env)?;
    let [ta, tb] = number_texts(op, [&a, &b])?;
    let na = val(&ta);
    let nb = val(&tb);
    let number = |x: f64| Node::sym(str_(x).trim());
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
fn number_texts(op: &str, operands: [&Node; 2]) -> Result<[String; 2], Refusal> {
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
fn number_text(op: &str, v: &Node) -> Result<String, Refusal> {
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
fn expand_arity(op: &str, want: usize, lst: &[Node]) -> Refusal {
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
fn atom_text_expand(v: &Node) -> String {
    match v {
        Node::Str(s) | Node::Sym(s) => s.to_string(),
        Node::List(_) => String::new(),
    }
}

/// VLA.3: `(&expand a b ...)`, each argument resolved as every LISTOPS
/// argument is and its text joined. Always a string literal, never a
/// symbol, so a macro writes an address or a formula's text with no
/// leading-quote trick; `symbol` keeps that trick for identifiers. At
/// least one argument, as `symbol` takes.
fn eval_concat_expand(lst: &[Node], env: &Env) -> Result<Node, Refusal> {
    if lst.len() < 2 {
        return Err(raise("vla-concat-expand-arity", &[]));
    }
    let mut joined = String::new();
    for k in 2..=lst.len() {
        let piece = resolve_listops_arg(nth(lst, k)?, env)?;
        if piece.is_list() {
            return Err(raise(
                "vla-expand-arg-is-list",
                &[("op", "&expand"), ("n", &(k - 1).to_string())],
            ));
        }
        joined.push_str(&atom_text_expand(&piece));
    }
    Ok(Node::Str(Arc::from(joined.as_str())))
}

/// VLA.3: `(characters-expand text)`, the text split into one-character
/// string literals in order. A character is a Unicode scalar value, as
/// the reference counts one by joining a UTF-16 surrogate pair. The empty
/// text gives the empty list.
fn eval_characters_expand(lst: &[Node], env: &Env) -> Result<Node, Refusal> {
    if lst.len() != 2 {
        return Err(expand_arity("characters-expand", 1, lst));
    }
    let v = resolve_listops_arg(nth(lst, 2)?, env)?;
    if v.is_list() {
        return Err(raise(
            "vla-expand-arg-is-list",
            &[("op", "characters-expand"), ("n", "1")],
        ));
    }
    let items = atom_text_expand(&v)
        .chars()
        .map(|c| Node::Str(Arc::from(c.to_string().as_str())))
        .collect();
    Ok(Node::list(items))
}

/// VLA.3: `(address-expand row column)`, a cell's address, relative, as
/// Excel's `ADDRESS(row, column, 4)` writes it: `B3`. Each argument a
/// number literal under the seven's gate, then a whole number within
/// Excel's bounds. A string literal.
fn eval_address_expand(lst: &[Node], env: &Env) -> Result<Node, Refusal> {
    const OP: &str = "address-expand";
    if lst.len() != 3 {
        return Err(expand_arity(OP, 2, lst));
    }
    let a = resolve_listops_arg(nth(lst, 2)?, env)?;
    let b = resolve_listops_arg(nth(lst, 3)?, env)?;
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
    let address = format!("{}{}", column_letters(col as u32), str_(row).trim());
    Ok(Node::Str(Arc::from(address.as_str())))
}

/// VLA.3: `(row-expand address)` and `(column-expand address)`, an
/// address's row and column numbers, as Excel's `ROW` and `COLUMN` of the
/// reference; and `(column-expand letters)`, a column's letters alone, the
/// reverse of `column-letters-expand`.
fn eval_address_part_expand(op: &str, lst: &[Node], env: &Env) -> Result<Node, Refusal> {
    if lst.len() != 2 {
        return Err(expand_arity(op, 1, lst));
    }
    let v = resolve_listops_arg(nth(lst, 2)?, env)?;
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
    Ok(Node::sym(str_(f64::from(n)).trim()))
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
fn eval_column_letters_expand(lst: &[Node], env: &Env) -> Result<Node, Refusal> {
    const OP: &str = "column-letters-expand";
    if lst.len() != 2 {
        return Err(expand_arity(OP, 1, lst));
    }
    let v = resolve_listops_arg(nth(lst, 2)?, env)?;
    let text = number_text(OP, &v)?;
    let col = val(&text);
    if col.fract() != 0.0 || !(1.0..=f64::from(MAX_COLUMN)).contains(&col) {
        return Err(raise(
            "vla-address-expand-column",
            &[("op", OP), ("value", &text)],
        ));
    }
    Ok(Node::Str(Arc::from(column_letters(col as u32).as_str())))
}

/// VLA.5: `(assoc-expand head entries)`, the first entry of the list whose
/// own head is `head`, compared as `equal?` compares (an atom as `eq?`
/// compares it, a symbol by its folded text and a string exactly, never
/// one to the other, and a list by its structure), the entry itself; or
/// the empty list when none is: Scheme's `assoc`. An entry is a list with
/// a head, and one that is not is refused when the search reaches it; an
/// entry past the one found is never looked at.
fn eval_assoc_expand(lst: &[Node], env: &Env) -> Result<Node, Refusal> {
    const OP: &str = "assoc-expand";
    if lst.len() != 3 {
        return Err(expand_arity(OP, 2, lst));
    }
    let head = resolve_listops_arg(nth(lst, 2)?, env)?;
    let listed = resolve_listops_arg(nth(lst, 3)?, env)?;
    let Node::List(entries) = &listed else {
        return Err(raise(
            "vla-expand-expected-list",
            &[
                ("op", OP),
                ("value", &listed.atom_text().unwrap_or_default()),
            ],
        ));
    };
    for entry in entries.items() {
        let Node::List(e) = entry else {
            return Err(raise(
                "vla-assoc-expand-entry",
                &[("value", &entry.atom_text().unwrap_or_default())],
            ));
        };
        let Some(first) = e.items().first() else {
            return Err(raise("vla-assoc-expand-entry", &[("value", "()")]));
        };
        if deep_equal(&head, first) {
            return Ok(entry.clone());
        }
    }
    Ok(Node::list(Vec::new()))
}

/// VBA `EvalCondFrom`: `(cond (test form) ... (else form))`, one clause at
/// a time from `start_idx`; a test that folds to true or false is taken or
/// skipped at expand time, any other test defers to a native `if`.
fn eval_cond_from(lst: &[Node], start_idx: usize, env: &Env) -> Result<Node, Refusal> {
    if start_idx > lst.len() {
        return Ok(Node::list(vec![Node::sym("begin")]));
    }
    let cl = nth_list(lst, start_idx)?.items();
    if cl.len() != 2 {
        return Err(raise(
            "vla-cond-clause-shape",
            &[("n", &cl.len().to_string())],
        ));
    }
    if let Node::Sym(s) = nth(cl, 1)? {
        if fold(s) == "else" {
            if start_idx != lst.len() {
                return Err(raise("vla-cond-else-not-last", &[]));
            }
            return substitute(nth(cl, 2)?, env, false);
        }
    }
    let t = resolve_listops_arg(nth(cl, 1)?, env)?;
    let tf = match &t {
        Node::Sym(s) => fold(s),
        _ => String::new(),
    };
    if tf == "true" {
        return substitute(nth(cl, 2)?, env, false);
    }
    if tf == "false" {
        return eval_cond_from(lst, start_idx + 1, env);
    }
    let raw_form = substitute(nth(cl, 2)?, env, false)?;
    let rest_form = eval_cond_from(lst, start_idx + 1, env)?;
    Ok(Node::list(vec![
        Node::sym("if"),
        t,
        Node::list(vec![Node::sym("then"), raw_form]),
        Node::list(vec![Node::sym("else"), rest_form]),
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
        // VLA.4: an operation of the walk, recognized by `expand_at` alone;
        // VLA.5: the loop, a form of the walk and of `substitute`; VLA.6:
        // define and a reference's two parts, forms of the walk alone;
        // VLA.15: its first cell, the third.
        let walk_operations = [
            "formula-expand",
            "each",
            "each-index",
            "define",
            "sheet-expand",
            "range-expand",
            "first-cell-expand",
        ];
        for name in LISTOPS {
            assert!(
                reserved.contains(name),
                "{name} is recognised but not reserved"
            );
        }
        for name in reserved {
            assert!(
                LISTOPS.contains(&name)
                    || quote_bundle.contains(&name)
                    || walk_operations.contains(&name),
                "{name} is reserved but not recognised"
            );
        }
        assert_eq!(LISTOPS.len(), 29);
        assert_eq!(one("(formula-expand (+ 1 2))"), "\"=1+2\"");
    }

    // VLA.4: formula-expand, the formula syntax at expand time.
    #[test]
    fn formula_expand_expands_its_argument_first() {
        // A macro inside the expression fires before the expression is
        // written, so it is never written as a function's call.
        assert_eq!(
            one("(defmacro (edge cell twin) (and (= cell 1) (= twin 0)))\n\
                 (formula-expand (if (edge Input!B5 Input.last!B5) 1 0))"),
            "\"=IF(AND(Input!B5=1,Input.last!B5=0),1,0)\""
        );
        // A template's parameters are filled in inside it, and the row it
        // stands in is written with the formula's text.
        assert_eq!(
            one(
                "(defmacro (rule sheet range expr) (formula sheet range (formula-expand expr)))\n\
                 (rule \"Board\" \"A1:B2\" (+ A1 1))"
            ),
            "(formula \"Board\" \"A1:B2\" \"=A1+1\")"
        );
        // An expand-time primitive inside it computes, a reference built
        // by `ref` from the address it computes.
        assert_eq!(
            one("(formula-expand (ref \"Input\" (address-expand 5 2)))"),
            "\"=Input!B5\""
        );
        // A text stays a text.
        assert_eq!(one("(formula-expand \"=A1+1\")"), "\"=\\\"=A1+1\\\"\"");
        assert_eq!(
            expand_err("(formula-expand (+ (formula-expand A1) 1))").id,
            "vla-formula-expand-nested"
        );
        assert_eq!(
            expand_err("(formula-expand)").id,
            "vla-formula-expand-arity"
        );
        assert_eq!(
            expand_err("(formula-expand 1 2)").id,
            "vla-formula-expand-arity"
        );
        assert_eq!(
            expand_err("(&expand (formula-expand A1))").id,
            "vla-expand-arg-is-list"
        );
        assert_eq!(
            expand_err("(defmacro (formula-expand x) x)").id,
            "vla-defmacro-reserved-name"
        );
        // After a refusal the walk is outside every formula-expand again.
        let mut ex = Expander::new();
        let bad = read_forms("(formula-expand (+ (formula-expand A1) 1))").unwrap();
        assert!(ex.expand(&bad[0], 0).is_err());
        let good = read_forms("(formula-expand A1)").unwrap();
        assert_eq!(ex.expand(&good[0], 0).unwrap(), Form::string("=A1"));
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
        // forms. In the port, which copied arguments until VLA.2, the first
        // ran the stack out and the second the memory, before the 5000
        // chain limit could stop either; the reference shares arguments,
        // builds no copies, and meets the chain limit (TestFormTooDeep's
        // last pair). The port shares them now too, and still writes its
        // expansion out as a tree, so both stay refused here by name.
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

    /// A node's weight found by walking its tree, as `weigh` walked it.
    fn walked(n: &Node) -> (u32, u32) {
        match n {
            Node::List(l) => {
                let (mut forms, mut deepest) = (1, 0);
                for item in l.items() {
                    let (f, d) = walked(item);
                    forms += f;
                    deepest = deepest.max(d);
                }
                (forms, deepest + 1)
            }
            _ => (1, 0),
        }
    }

    #[test]
    fn a_list_keeps_the_weight_its_tree_has_and_a_view_its_own() {
        // VLA.2: the forms and the depth a list keeps are what a walk of
        // its tree finds, for a list as read, every view `cdr` and `cddr`
        // make of it from every start, a view of a view, and a list that
        // holds views beside the list itself. Every view is of the one run.
        for text in [
            "(a (b (c (d))) \"s\" (e f) ((g)) h)",
            "((((1))) 2 (3 (4 (5 (6)))) () 7)",
            "(1 2 3)",
            "()",
            "(quote (a quote b))",
        ] {
            let node = Node::of(&read_forms(text).unwrap()[0]);
            let Node::List(seq) = &node else { panic!() };
            assert_eq!(node.weight(), walked(&node), "{text}");
            for n in 1..=seq.count() {
                let Node::List(view) = list_tail(seq, n, "cdr").unwrap() else {
                    panic!()
                };
                assert!(Arc::ptr_eq(&view.run, &seq.run), "{text} from {n}");
                let once = Node::List(view.clone());
                assert_eq!(once.weight(), walked(&once), "{text} from {n}");
                for m in 1..=view.count() {
                    let twice = Node::List(view.tail(m));
                    assert_eq!(twice.weight(), walked(&twice), "{text} from {n}, {m}");
                }
                let holder = Node::list(vec![once.clone(), node.clone(), once]);
                assert_eq!(holder.weight(), walked(&holder), "{text} held from {n}");
            }
            assert_eq!(
                list_tail(seq, seq.count() + 1, "cddr").unwrap_err().id,
                "vla-list-too-short"
            );
        }
    }

    #[test]
    fn a_parameter_s_value_is_the_argument_itself_at_every_use() {
        // VLA.2, the reference's `Set Substitute = v`: a value used twice is
        // the argument's own list twice, a quoted argument's content is that
        // content, and `cdr` of it a view of its items. A list the template
        // writes is made anew, untagged.
        let mut ex = Expander::new();
        ex.collect(read_forms("(defmacro (m x y) (f x x (cdr y) (car y)))").unwrap())
            .unwrap();
        let call = Node::of(&read_forms("(m (g 1) (quote ((h 2) 3)))").unwrap()[0]);
        let Node::List(call) = &call else { panic!() };
        let r = expand_one(call.items(), &ex.rules["m"], &ex.rules).unwrap();
        let Node::List(r) = &r else { panic!() };
        let shared = |n: &Node, with: &Node| match (n, with) {
            (Node::List(a), Node::List(b)) => Arc::ptr_eq(&a.run, &b.run),
            _ => false,
        };
        let [_, x, y] = call.items() else { panic!() };
        let Node::List(quoted) = y else { panic!() };
        let content = &quoted.items()[1];
        let Node::List(rows) = content else { panic!() };
        let [_, x1, x2, tail, head] = r.items() else {
            panic!()
        };
        assert!(shared(x1, x) && shared(x2, x));
        assert!(shared(tail, content));
        let Node::List(tail) = tail else { panic!() };
        assert_eq!((tail.start, tail.line), (1, 0));
        assert!(shared(head, &rows.items()[0]));
        assert_eq!(r.line, 0);
        assert_eq!(
            Node::List(r.clone()).form(),
            read_forms("(f (g 1) (g 1) (3) (h 2))").unwrap()[0]
        );
    }

    #[test]
    fn a_result_past_a_limit_is_refused_as_the_walk_refused_it() {
        // VLA.1's weighing reads a result's two numbers (VLA.2) and walks it
        // only past a limit, in the copies' order, so the refusal is the
        // walk's: a value doubled refused at its nineteenth doubling, past a
        // million forms, a list nested past 64 named at its line, none for a
        // list a macro made.
        let small = Node::of(&read_forms("(a (b) c)").unwrap()[0]);
        assert_eq!(small.weight(), (5, 2));
        let ex = Expander::new();
        let mut doubled = Node::sym("1");
        for k in 1..=19 {
            doubled = Node::list(vec![doubled.clone(), doubled.clone()]);
            let refused = ex.weigh(&doubled);
            assert_eq!(refused.is_err(), k == 19, "{k} doublings");
        }
        assert_eq!(doubled.weight().0, (1 << 20) - 1);
        assert_eq!(
            ex.weigh(&doubled).unwrap_err().id,
            "vla-macro-expansion-too-large"
        );
        let mut deep = Node::sym("1");
        for k in 1..=65 {
            deep = Node::list(vec![deep]);
            assert_eq!(ex.weigh(&deep).is_err(), k == 65, "{k} deep");
        }
        let r = ex.weigh(&deep).unwrap_err();
        assert_eq!(r.id, "vla-form-too-deep");
        assert_eq!(r.text, too_deep(0).text);
        // A count past what a u32 holds saturates and stays saturated in
        // every view, so the walk decides.
        let mut huge = Node::sym("1");
        for _ in 0..40 {
            huge = Node::list(vec![huge.clone(), huge.clone(), huge]);
        }
        let Node::List(seq) = &huge else { panic!() };
        assert_eq!(seq.forms, u32::MAX);
        assert_eq!(seq.tail(1).forms, u32::MAX);
        assert_eq!(
            ex.weigh(&huge).unwrap_err().id,
            "vla-macro-expansion-too-large"
        );
    }

    #[test]
    fn a_quoted_list_and_a_pass_s_arguments_keep_their_lines() {
        // VLA.2: a form read into nodes and written out again keeps every
        // list's line where the walk does not rebuild it: a quoted list,
        // returned as it stands, and a user's argument inside the result of
        // one pass. Found by a mutant that wrote every list at line 0.
        let mut ex = Expander::new();
        ex.collect(read_forms("(defmacro (wrap2 x) (begin (g 0) x))").unwrap())
            .unwrap();
        let quoted = read_forms("\n\n(quote (a\n(b)))").unwrap().remove(0);
        let out = ex.expand(&quoted, 0).unwrap();
        let Form::List(q) = &out else { panic!() };
        let Form::List(data) = &q.items[1] else {
            panic!()
        };
        let Form::List(b) = &data.items[1] else {
            panic!()
        };
        assert_eq!((q.line, data.line, b.line), (3, 3, 4));
        let call = read_forms("\n\n(wrap2\n(f 1))").unwrap().remove(0);
        let out = ex.expand_one_pass(&call).unwrap();
        let Form::List(r) = &out else { panic!() };
        assert_eq!(r.line, 0);
        assert_eq!(r.items[2].as_list().unwrap().line, 4);
    }

    // VLA.5: the loop and assoc-expand, TestEachForms (VLA_Tests.bas) pin
    // for pin.
    #[test]
    fn a_loop_writes_its_forms_once_an_element() {
        assert_eq!(one("(each (x (quote (a b))) (f x))"), "(begin (f a) (f b))");
        assert_eq!(
            one("(each (x (quote (a b))) (when x (f)))"),
            "(begin (if a (then (f))) (if b (then (f))))"
        );
        assert_eq!(
            one("(defmacro (cells sheet & entries) (each (entry entries) (cell sheet (car entry) (car (cdr entry))))) \
                 (cells \"S\" (\"A1\" 1) (\"B1\" 2))"),
            "(begin (cell \"S\" \"A1\" 1) (cell \"S\" \"B1\" 2))"
        );
        assert_eq!(one("(each (x (quote ())) (f x))"), "(begin)");
        assert_eq!(
            one("(each (x (quote (a b))) (f x) (g x))"),
            "(begin (f a) (g a) (f b) (g b))"
        );
        assert_eq!(
            one("(defmacro (seq & xs) (each (x xs) x)) (seq 1 2 3)"),
            "(begin 1 2 3)"
        );
    }

    #[test]
    fn each_index_counts_from_the_start_its_binding_names() {
        assert_eq!(
            one("(each-index (i x (quote (a b c)) 0) (f i x))"),
            "(begin (f 0 a) (f 1 b) (f 2 c))"
        );
        assert_eq!(
            one("(each-index (i x (quote (a b)) 1) (f i x))"),
            "(begin (f 1 a) (f 2 b))"
        );
        assert_eq!(
            one("(each-index (row x (quote (a b)) (row-expand \"C7\")) (f row x))"),
            "(begin (f 7 a) (f 8 b))"
        );
        assert_eq!(
            one("(each-index (i x (quote (a b)) -1) (f i))"),
            "(begin (f -1) (f 0))"
        );
    }

    #[test]
    fn a_loop_splices_as_an_element_and_is_its_list_where_a_value_is_read() {
        assert_eq!(one("(+ (each (c (quote (A1 B1 C1))) c))"), "(+ A1 B1 C1)");
        assert_eq!(
            one("(defmacro (total & cells) (+ (each (c cells) c))) (total A1 B1)"),
            "(+ A1 B1)"
        );
        assert_eq!(
            one("(formula-expand (+ (each (c (quote (A1 B1))) c)))"),
            "\"=A1+B1\""
        );
        assert_eq!(
            one("(each (r (quote ((a b) (c)))) (each (x r) (f x)))"),
            "(begin (f a) (f b) (f c))"
        );
        assert_eq!(one("(car (cdr (each (x (quote (a b c))) (f x))))"), "(f b)");
        assert_eq!(
            one("(each (y (each (x (quote (a b))) (f x))) (g y))"),
            "(begin (g (f a)) (g (f b)))"
        );
        assert_eq!(
            one("(defmacro (m & xs) (quasiquote (f (unquote-splicing (each (x xs) (g x)))))) (m a b)"),
            "(f (g a) (g b))"
        );
        assert_eq!(
            one("(defmacro (m) (quasiquote (each (x (quote (a b))) x))) (m)"),
            "(begin a b)"
        );
    }

    #[test]
    fn a_malformed_loop_is_refused_by_name() {
        for (src, id, frag) in [
            (
                "(each (x (quote (a))))",
                "vla-each-shape",
                "each: write (each (name list) form ...)",
            ),
            (
                "(each x (quote (a)) (f x))",
                "vla-each-shape",
                "each: write (each (name list) form ...)",
            ),
            (
                "(each (\"x\" (quote (a))) (f))",
                "vla-each-shape",
                "each: write (each (name list) form ...)",
            ),
            (
                "(each-index (i x (quote (a))) (f x))",
                "vla-each-shape",
                "each-index: write (each-index (index name list start) form ...)",
            ),
            (
                "(each (car (quote (a))) (f car))",
                "vla-each-name-reserved",
                "each 'car': 'car' is the list-head primitive and cannot be a name each binds",
            ),
            (
                "(defmacro (m xs) (each (xs xs) (f xs))) (m (quote (a)))",
                "vla-each-name-bound",
                "each: 'xs' is already a name in this template",
            ),
            (
                "(defmacro (m & xs) (each (xs xs) (f xs))) (m a)",
                "vla-each-name-bound",
                "each: 'xs' is already a name in this template",
            ),
            (
                "(each (x (quote ((a)))) (each (x x) (f x)))",
                "vla-each-name-bound",
                "each: 'x' is already a name in this template",
            ),
            (
                "(each-index (x x (quote (a)) 0) (f x))",
                "vla-each-name-bound",
                "each-index: 'x' is already a name in this template",
            ),
            (
                "(each (x y) (f x))",
                "vla-expand-expected-list",
                "each: expected a list, got 'y'",
            ),
            (
                "(each (x (when a b)) (f x))",
                "vla-each-list-a-call",
                "each: the list to walk is a call of the macro 'when'",
            ),
            (
                "(each-index (i x (quote (a)) one) (f i))",
                "vla-arith-expand-not-number-value",
                "each-index: expected a number, got 'one'",
            ),
            (
                "(each-index (i x (quote (a)) (quote (1))) (f i))",
                "vla-arith-expand-not-number-list",
                "each-index: expected a number, got a list",
            ),
            (
                "(each-index (i x (quote (a)) 1.5) (f i))",
                "vla-arith-expand-not-whole",
                "each-index: expected a whole number, got '1.5'",
            ),
        ] {
            let r = expand_err(src);
            assert_eq!(r.id, id, "{src}");
            assert!(r.text.contains(frag), "{src}: {}", r.text);
        }
    }

    #[test]
    fn assoc_expand_answers_the_first_entry_by_its_head_or_the_empty_list() {
        assert_eq!(
            one("(assoc-expand (quote b) (quote ((a 1) (b 2) (b 3))))"),
            "(b 2)"
        );
        assert_eq!(one("(assoc-expand (quote z) (quote ((a 1))))"), "()");
        assert_eq!(
            one("(assoc-expand (quote SHEET) (quote ((sheet \"Tower\"))))"),
            "(sheet \"Tower\")"
        );
        assert_eq!(one("(assoc-expand \"a\" (quote ((a 1))))"), "()");
        assert_eq!(
            one("(assoc-expand (quote (1 2)) (quote (((1 2) x))))"),
            "((1 2) x)"
        );
        assert_eq!(
            one("(assoc-expand (quote a) (quote ((a 1) junk)))"),
            "(a 1)"
        );
        assert_eq!(
            one(
                "(defmacro (put at & pieces) (cell (car (cdr (assoc-expand sheet pieces))) at)) \
                 (put \"A2\" (value 1) (sheet \"Tower\"))"
            ),
            "(cell \"Tower\" \"A2\")"
        );
        for (src, id, frag) in [
            (
                "(assoc-expand (quote a))",
                "vla-expand-arity",
                "assoc-expand expects exactly 2 argument(s), got 1",
            ),
            (
                "(assoc-expand (quote a) b)",
                "vla-expand-expected-list",
                "assoc-expand: expected a list, got 'b'",
            ),
            (
                "(assoc-expand (quote a) (quote (x (a 1))))",
                "vla-assoc-expand-entry",
                "assoc-expand: an entry is a list with a head, got 'x'",
            ),
            (
                "(assoc-expand (quote a) (quote (())))",
                "vla-assoc-expand-entry",
                "assoc-expand: an entry is a list with a head, got '()'",
            ),
        ] {
            let r = expand_err(src);
            assert_eq!(r.id, id, "{src}");
            assert!(r.text.contains(frag), "{src}: {}", r.text);
        }
    }

    #[test]
    fn the_loop_s_names_are_reserved_and_the_one_table_still_refuses() {
        for (name, desc) in [
            ("each", "expand-time loop form"),
            ("each-index", "expand-time counting loop form"),
            ("assoc-expand", "expand-time entry-lookup primitive"),
        ] {
            let r = expand_err(&format!("(defmacro ({name} x) x)"));
            assert_eq!(r.id, "vla-defmacro-reserved-name");
            assert_eq!(
                r.text,
                format!("defmacro '{name}': '{name}' is the {desc} and cannot be a macro name")
            );
        }
        assert_eq!(
            expand_err("(defmacro (QUOTE x) x)").text,
            "defmacro 'quote': 'quote' is the data-literal form and cannot be a macro name"
        );
    }

    #[test]
    fn a_begin_a_macro_made_keeps_its_forms_before_an_atom() {
        // Found while building the loop: the trampoline returned the atom
        // alone and dropped the forms it had set aside before it.
        assert_eq!(
            one("(defmacro (m) (debug-print 1) x) (m)"),
            "(begin (debug-print 1) x)"
        );
        assert_eq!(
            one("(defmacro (m & xs) (begin xs)) (m 1 2 3)"),
            "(begin 1 2 3)"
        );
    }

    #[test]
    fn a_loop_of_2500_rows_is_one_application_past_the_walker_s_cap() {
        let entries: Vec<String> = (1..=2500).map(|i| format!("(\"A{i}\" {i})")).collect();
        let src = format!(
            "(defmacro (cells sheet & entries) (each (entry entries) (cell sheet (car entry) (car (cdr entry))))) \
             (cells \"S\" {})",
            entries.join(" ")
        );
        let (text, fired) = expand_text(&src, PRELUDE, true).unwrap();
        assert_eq!(fired, 1);
        assert!(text.contains("(cell \"S\" \"A2500\" 2500)"));
    }

    // VLA.6: a name at expand time, (define name form), and a reference's
    // twin, (last X), with a reference's two parts, TestDefineForms' pins.
    const NAMES_MACROS: &str = "(defmacro (names sheet top & words) \
         (each-index (row word words top) (label-and-name sheet row word))) \
         (defmacro (label-and-name sheet row word) \
         (cell sheet (address-expand row 1) (&expand word)) \
         (define word (ref sheet (address-expand row 2)))) ";

    #[test]
    fn a_define_binds_the_forms_after_its_own_and_is_replaced_as_an_expression() {
        assert_eq!(
            one("(define stop State!B12) (f stop)"),
            "(begin)\n(f State!B12)"
        );
        assert_eq!(
            one("(define stop State!B12) (formula-expand (if (= stop 1) 1 0))"),
            "(begin)\n\"=IF(State!B12=1,1,0)\""
        );
        assert_eq!(
            one("(define left State!B5) (formula-expand (left left 1))"),
            "(begin)\n\"=LEFT(State!B5,1)\""
        );
        assert_eq!(
            one("(define stop State!B12) (f (quote stop) \"stop\" stop)"),
            "(begin)\n(f (quote stop) \"stop\" State!B12)"
        );
        assert_eq!(
            one("(define stop State!B12) (eq? stop (quote stop))"),
            "(begin)\ntrue"
        );
        assert_eq!(
            one("(defmacro (twice stop) (+ stop stop)) (define stop State!B12) (twice 5) (twice stop)"),
            "(begin)\n(+ 5 5)\n(+ State!B12 State!B12)"
        );
        assert_eq!(
            one("(define last-column State!B24) (define right-end last-column) (f right-end)"),
            "(begin)\n(begin)\n(f State!B24)"
        );
        assert_eq!(
            one("(define right-end last-column) (define last-column State!B24) (f right-end)"),
            "(begin)\n(begin)\n(f State!B24)"
        );
        assert_eq!(
            one(
                "(defmacro (pressed-now button) (and (= button 1) (= (last button) 0))) \
                 (define press-edge (pressed-now Input!B5)) (formula-expand (if press-edge 1 0))"
            ),
            "(begin)\n\"=IF(AND(Input!B5=1,Input.last!B5=0),1,0)\""
        );
        assert_eq!(
            one(
                "(defmacro (name-cell word sheet address) (define word (ref sheet address))) \
                 (name-cell stop \"State\" \"B12\") (formula-expand (+ stop 1))"
            ),
            "(begin)\n\"=State!B12+1\""
        );
        assert_eq!(
            one("(begin (define stop State!B12) (f 1)) (f stop)"),
            "(begin (f 1))\n(f State!B12)"
        );
        let (_, fired) = expand_text("(define stop State!B12) (f stop)", PRELUDE, true).unwrap();
        assert_eq!(fired, 0, "a word replaced is no macro's application");
        let (one_pass, _) =
            expand_text("(define stop (when a b)) (f stop)", PRELUDE, false).unwrap();
        assert_eq!(one_pass, "(define stop (when a b))\r\n(f stop)\r\n");
    }

    #[test]
    fn a_loop_writes_a_block_s_labels_and_names_and_the_stop_rule_reads_them() {
        assert_eq!(
            one(&format!(
                "{NAMES_MACROS}(names \"State\" 1 phase since) (formula-expand (if (= phase 1) since 0))"
            )),
            "(begin (begin (cell \"State\" \"A1\" \"phase\")) (cell \"State\" \"A2\" \"since\"))\n\
             \"=IF(State!B1=1,State!B2,0)\""
        );
        let by_names = one(&format!(
            "{NAMES_MACROS}(defmacro (rule name expression) \
             (formula (sheet-expand name) (range-expand name) (formula-expand expression))) \
             (define sliding 1) (names \"State\" 1 phase) (names \"State\" 9 press take start stop) \
             (rule stop (if (and (= (last phase) sliding) (= press 1)) 1 0))"
        ));
        let kairos = "(formula \"State\" \"B12\" \"=IF(AND(State.last!B1=1,State!B9=1),1,0)\")";
        assert_eq!(by_names.lines().last(), Some(kairos));
        assert_eq!(
            one("(defmacro (rule sheet range expression) (formula sheet range (formula-expand expression))) \
                 (rule \"State\" \"B12\" (if (and (= State.last!B1 1) (= State!B9 1)) 1 0))"),
            kairos
        );
        assert_eq!(
            one("(each (w (quote (p q))) (cell (&expand w)) (define w (ref \"S\" \"A1\")))"),
            "(begin (cell \"p\") (cell \"q\"))"
        );
    }

    #[test]
    fn a_malformed_define_is_refused_by_name() {
        for (src, id, frag) in [
            ("(define stop)", "vla-define-shape", "- got (define stop)"),
            ("(define stop A1 B1)", "vla-define-shape", "- got (define stop A1 B1)"),
            (
                "(define \"stop\" A1)",
                "vla-define-shape",
                "define: write (define name form), a name and the one form it stands for - got (define \"stop\" A1)",
            ),
            (
                "(define B12 1)",
                "vla-define-name-reference",
                "define 'B12': 'B12' reads as a cell, and a define gives a reference a word of its own",
            ),
            ("(define A:A 1)", "vla-define-name-reference", "'A:A' reads as a column"),
            (
                "(define Sheet1!stop 1)",
                "vla-define-name-reference",
                "'Sheet1!stop' reads as a name on a sheet",
            ),
            (
                "(define R1C1 1)",
                "vla-define-name-reference",
                "'R1C1' reads as a reference in R1C1 notation",
            ),
            (
                "(define 12 A1)",
                "vla-define-name-reserved",
                "define '12': '12' is a number, and a define binds a word of the source's own",
            ),
            ("(define true 1)", "vla-define-name-reserved", "'true' is a truth value"),
            ("(define :stop 1)", "vla-define-name-reserved", "':stop' is a keyword argument"),
            ("(define if A1)", "vla-define-name-reserved", "'if' is a form of the language"),
            ("(define si A1)", "vla-define-name-reserved", "'si' is a form of the language"),
            ("(define car A1)", "vla-define-name-reserved", "'car' is the list-head primitive"),
            (
                "(define last State!B24)",
                "vla-define-name-reserved",
                "'last' is a form of the formula syntax's own",
            ),
            (
                "(define define A1)",
                "vla-define-name-reserved",
                "'define' is the expand-time naming form",
            ),
            (
                "(defmacro (since event) event) (define since State!B2)",
                "vla-define-name-taken",
                "define 'since': 'since' is a macro's name in this source - a word names one thing",
            ),
            (
                "(define stop A1) (define stop B1)",
                "vla-define-name-taken",
                "'stop' is defined already in this source",
            ),
            (
                "(begin (define stop A1) (define stop B1))",
                "vla-define-name-taken",
                "'stop' is defined already in this source",
            ),
            (
                "(cell \"S\" \"A1\" (define stop A1))",
                "vla-define-not-a-statement",
                "define 'stop': a define stands as a form of the source",
            ),
            (
                "(formula-expand (+ 1 (define stop A1)))",
                "vla-define-not-a-statement",
                "define 'stop'",
            ),
            (
                "(define stop (define press A1))",
                "vla-define-not-a-statement",
                "define 'press'",
            ),
            (
                "(begin (define stop State!B12) (f stop))",
                "vla-define-used-in-its-form",
                "'stop' is used in the form of the source that defines it, where it is not yet a name",
            ),
            (
                "(begin (define stop State!B12) (formula-expand (+ stop 1)))",
                "vla-define-used-in-its-form",
                "'stop' is used",
            ),
            (
                "(defmacro (use-then-define) (f stop) (define stop A1)) (use-then-define)",
                "vla-define-used-in-its-form",
                "'stop' is used",
            ),
            (
                "(define start-row end-row) (define end-row start-row)",
                "vla-define-cycle",
                "define 'end-row': the words end-row -> start-row -> end-row stand for one another in a ring",
            ),
            (
                "(define total (+ total 1))",
                "vla-define-cycle",
                "the words total -> total stand for one another",
            ),
            (
                "(defmacro (define x) x)",
                "vla-defmacro-reserved-name",
                "'define' is the expand-time naming form and cannot be a macro name",
            ),
            (
                "(each (define (quote (a))) (f define))",
                "vla-each-name-reserved",
                "each 'define': 'define' is the expand-time naming form and cannot be a name each binds",
            ),
        ] {
            let e = expand_err(src);
            assert_eq!(e.id, id, "{src}: {}", e.text);
            assert!(e.text.contains(frag), "{src}: {}", e.text);
        }
    }

    #[test]
    fn last_writes_a_reference_s_twin_and_refuses_what_has_none() {
        assert_eq!(
            one("(formula-expand (+ (last Input!B13) 1))"),
            "\"=Input.last!B13+1\""
        );
        assert_eq!(
            one("(formula-expand (sum (last Tally!B1:E10)))"),
            "\"=SUM(Tally.last!B1:E10)\""
        );
        assert_eq!(
            one("(formula-expand (last (ref \"Q1 Data\" \"A1\")))"),
            "\"='Q1 Data.last'!A1\""
        );
        assert_eq!(
            one("(formula-expand (last State!$B$3))"),
            "\"=State.last!$B$3\""
        );
        assert_eq!(
            one("(formula-expand (sum (last Screen!A:A)))"),
            "\"=SUM(Screen.last!A:A)\""
        );
        assert_eq!(
            one("(define level State!B3) (formula-expand (+ (last level) 1))"),
            "(begin)\n\"=State.last!B3+1\""
        );
        for (src, id, frag) in [
            (
                "(formula-expand (last B13))",
                "vla-reference-no-sheet",
                "last: 'B13' names no sheet, and a formula does not know the sheet it will stand on",
            ),
            (
                "(formula-expand (last Input.last!B13))",
                "vla-formula-last-twin",
                "last: 'Input.last!B13' is a twin already, and a twin has none",
            ),
            (
                "(formula-expand (last stop))",
                "vla-reference-expected",
                "got 'stop', a name",
            ),
            (
                "(formula-expand (last Jan:Dec!A1))",
                "vla-reference-expected",
                "got 'Jan:Dec!A1', a span of sheets",
            ),
            (
                "(formula-expand (last 5))",
                "vla-reference-expected",
                "got '5', no reference",
            ),
            ("(formula-expand (last A1 B1))", "vla-formula-arity", "(last ...)"),
            (
                "(formula-expand (last (ref \"A1\")))",
                "vla-reference-no-sheet",
                "'(ref \"A1\")' names no sheet",
            ),
        ] {
            let e = expand_err(src);
            assert_eq!(e.id, id, "{src}: {}", e.text);
            assert!(e.text.contains(frag), "{src}: {}", e.text);
        }
    }

    #[test]
    fn a_reference_s_two_parts_write_a_row_at_a_name() {
        assert_eq!(
            one("(define stop State!B12) (formula (sheet-expand stop) (range-expand stop) (formula-expand (+ stop 1)))"),
            "(begin)\n(formula \"State\" \"B12\" \"=State!B12+1\")"
        );
        assert_eq!(one("(range-expand Tally!B1:E10)"), "\"B1:E10\"");
        assert_eq!(
            one("(sheet-expand (ref \"Q1 Data\" \"A1\"))"),
            "\"Q1 Data\""
        );
        for (src, id, frag) in [
            ("(sheet-expand B12)", "vla-reference-no-sheet", "sheet-expand: 'B12' names no sheet"),
            (
                "(range-expand stop)",
                "vla-reference-expected",
                "range-expand: expected one reference on a sheet",
            ),
            (
                "(sheet-expand A1 B1)",
                "vla-expand-arity",
                "sheet-expand expects exactly 1 argument(s), got 2",
            ),
            (
                "(defmacro (sheet-expand x) x)",
                "vla-defmacro-reserved-name",
                "'sheet-expand' is the expand-time reference-sheet operation and cannot be a macro name",
            ),
            (
                "(defmacro (range-expand x) x)",
                "vla-defmacro-reserved-name",
                "'range-expand' is the expand-time reference-range operation and cannot be a macro name",
            ),
        ] {
            let e = expand_err(src);
            assert_eq!(e.id, id, "{src}: {}", e.text);
            assert!(e.text.contains(frag), "{src}: {}", e.text);
        }
    }

    #[test]
    fn a_begin_a_macro_made_keeps_its_forms_before_every_return_of_its_own() {
        // Found while scoping VLA.6: the quote, the formula's text and the
        // primitive's result were returned alone, the debug-print dropped.
        assert_eq!(
            one("(defmacro (m) (debug-print 1) (quote (a b))) (m)"),
            "(begin (debug-print 1) (quote (a b)))"
        );
        assert_eq!(
            one("(defmacro (m) (debug-print 1) (formula-expand (+ A1 1))) (m)"),
            "(begin (debug-print 1) \"=A1+1\")"
        );
        assert_eq!(
            one("(defmacro (m) (quasiquote (begin (debug-print 1) (car (quote (a b)))))) (m)"),
            "(begin (debug-print 1) a)"
        );
    }

    /// The pins' last words: the refusal's id, and its text holding the
    /// fragment.
    fn assert_refusals(cases: &[(&str, &str, &str)]) {
        for (src, id, frag) in cases {
            let e = expand_err(src);
            assert_eq!(e.id, *id, "{src}: {}", e.text);
            assert!(e.text.contains(frag), "{src}: {}", e.text);
        }
    }

    #[test]
    fn list_tells_a_word_from_a_list_and_refuses_neither() {
        // VLA.15, `TestAddressForms`' first part.
        for (src, want) in [
            ("(list? (quote (a b)))", "true"),
            ("(list? (quote ()))", "true"),
            ("(list? (quote a))", "false"),
            ("(list? \"a\")", "false"),
            ("(list? 5)", "false"),
            ("(list? (each (x (quote (a b))) x))", "true"),
            ("(list? (cdr (quote (a b))))", "true"),
            ("(define stop State!B12) (list? stop)", "(begin)\nfalse"),
        ] {
            assert_eq!(one(src), want, "{src}");
        }
        assert_eq!(
            one("(defmacro (entry-label entry) (quote-if (list? entry) (car entry) (&expand entry))) \
                 (cell \"State\" \"A1\" (entry-label phase)) \
                 (cell \"State\" \"A2\" (entry-label (\"stopped at\" stopped-at)))"),
            "(cell \"State\" \"A1\" \"phase\")\n(cell \"State\" \"A2\" \"stopped at\")"
        );
        let names = "(defmacro (names sheet top & entries) (each-index (row entry entries top) (name-entry sheet row entry))) \
                     (defmacro (name-entry sheet row entry) (quote-if (list? entry) \
                     (label-and-name sheet row (car entry) (car (cdr entry))) (label-and-name sheet row (&expand entry) entry))) \
                     (defmacro (label-and-name sheet row label word) (cell sheet (address-expand row 1) label) \
                     (define word (ref sheet (address-expand row 2)))) ";
        assert_eq!(
            one(&format!(
                "{names}(names \"State\" 1 phase (\"stopped at\" stopped-at)) (formula-expand (if (= phase 1) stopped-at 0))"
            )),
            "(begin (begin (cell \"State\" \"A1\" \"phase\")) (cell \"State\" \"A2\" \"stopped at\"))\n\"=IF(State!B1=1,State!B2,0)\""
        );
        assert_refusals(&[
            (
                "(list? (quote a) (quote b))",
                "vla-expand-arity",
                "list? expects exactly 1 argument(s), got 2",
            ),
            ("(list?)", "vla-expand-arity", "got 0"),
            (
                "(defmacro (list? x) x)",
                "vla-defmacro-reserved-name",
                "'list?' is the list-test primitive and cannot be a macro name",
            ),
            (
                "(define list? A1)",
                "vla-define-name-reserved",
                "'list?' is the list-test primitive",
            ),
        ]);
    }

    #[test]
    fn the_three_marks_anchor_a_reference_and_commute_with_last() {
        // VLA.15, `TestAddressForms`' second part: each head adds its marks
        // and takes none away.
        for (src, want) in [
            ("(formula-expand (+ (absolute State!B12) 1))", "\"=State!$B$12+1\""),
            ("(formula-expand (absolute-row Gate!A4))", "\"=Gate!A$4\""),
            ("(formula-expand (absolute-column Gate!D1))", "\"=Gate!$D1\""),
            ("(formula-expand (sum (absolute Tally!B1:E10)))", "\"=SUM(Tally!$B$1:$E$10)\""),
            ("(formula-expand (sum (absolute-row Tally!B1:E10)))", "\"=SUM(Tally!B$1:E$10)\""),
            ("(formula-expand (sum (absolute-column Tally!B1:E10)))", "\"=SUM(Tally!$B1:$E10)\""),
            ("(formula-expand (sum (absolute Screen!A:A)))", "\"=SUM(Screen!$A:$A)\""),
            ("(formula-expand (sum (absolute-column Screen!A:C)))", "\"=SUM(Screen!$A:$C)\""),
            ("(formula-expand (sum (absolute 3:3)))", "\"=SUM($3:$3)\""),
            ("(formula-expand (sum (absolute-row Screen!3:5)))", "\"=SUM(Screen!$3:$5)\""),
            ("(formula-expand (+ (absolute B6) 1))", "\"=$B$6+1\""),
            ("(formula-expand (absolute-row $B12))", "\"=$B$12\""),
            ("(formula-expand (absolute (ref \"Q1 Data\" \"A1\")))", "\"='Q1 Data'!$A$1\""),
            ("(formula-expand (absolute Input.last!B5))", "\"=Input.last!$B$5\""),
            (
                "(define stop State!B12) (formula-expand (+ (absolute stop) (last (absolute stop))))",
                "(begin)\n\"=State!$B$12+State.last!$B$12\"",
            ),
            (
                "(define stop State!B12) (formula-expand (= (last (absolute stop)) (absolute (last stop))))",
                "(begin)\n\"=State.last!$B$12=State.last!$B$12\"",
            ),
            ("(formula-expand (absolute-row (absolute-column Input!B5)))", "\"=Input!$B$5\""),
            ("(formula-expand (absolute (absolute Input!B5)))", "\"=Input!$B$5\""),
        ] {
            assert_eq!(one(src), want, "{src}");
        }
        let fill = one("(defmacro (rule sheet range expression) (formula sheet range (formula-expand expression))) \
                        (define stop State!B12) \
                        (rule \"Gate\" \"A1:O1\" (if (= (absolute stop) 1) (last (absolute stop)) 0))");
        assert_eq!(
            fill.lines().last(),
            Some("(formula \"Gate\" \"A1:O1\" \"=IF(State!$B$12=1,State.last!$B$12,0)\")")
        );
        assert_refusals(&[
            ("(formula-expand (absolute A1 B1))", "vla-formula-arity", "(absolute ...)"),
            (
                "(formula-expand (absolute stop))",
                "vla-reference-expected",
                "absolute: expected a cell, a range, a column or a row, such as Input!B5",
            ),
            ("(formula-expand (absolute 5))", "vla-reference-expected", "got '5', no reference"),
            (
                "(formula-expand (absolute Jan:Dec!A1))",
                "vla-reference-expected",
                "got 'Jan:Dec!A1', a span of sheets",
            ),
            ("(formula-expand (absolute \"B12\"))", "vla-reference-expected", "', a text"),
            (
                "(formula-expand (absolute-row Screen!A:A))",
                "vla-reference-whole-line",
                "absolute-row: 'Screen!A:A' is a whole column, and absolute-row takes a reference with a row to mark",
            ),
            (
                "(formula-expand (absolute-column 3:3))",
                "vla-reference-whole-line",
                "absolute-column: '3:3' is a whole row, and absolute-column takes a reference with a column to mark",
            ),
            (
                "(formula-expand (last (absolute B12)))",
                "vla-reference-no-sheet",
                "last: '(absolute B12)' names no sheet",
            ),
            (
                "(formula-expand (absolute (last Input.last!B5)))",
                "vla-formula-last-twin",
                "last: 'Input.last!B5' is a twin already",
            ),
            (
                "(formula-expand (last (last Input!B5)))",
                "vla-formula-last-twin",
                "last: '(last Input!B5)' is a twin already",
            ),
            (
                "(define absolute A1)",
                "vla-define-name-reserved",
                "'absolute' is a form of the formula syntax's own",
            ),
            (
                "(define absolute-row A1)",
                "vla-define-name-reserved",
                "'absolute-row' is a form of the formula syntax's own",
            ),
            (
                "(define absolute-column A1)",
                "vla-define-name-reserved",
                "'absolute-column' is a form of the formula syntax's own",
            ),
        ]);
    }

    #[test]
    fn the_first_cell_is_a_third_form_of_the_walk() {
        // VLA.15, `TestAddressForms`' third part: a range's top-left as the
        // grid reads a row's range, the marks written on each lesser.
        for (src, want) in [
            ("(first-cell-expand Tally!B1:E10)", "\"B1\""),
            ("(first-cell-expand State!B12)", "\"B12\""),
            ("(first-cell-expand Tally!$B$1:$E$10)", "\"$B$1\""),
            ("(first-cell-expand Tally!E10:B1)", "\"B1\""),
            ("(first-cell-expand Tally!E$10:$B1)", "\"$B1\""),
            ("(define tally Tally!B1:E10) (first-cell-expand tally)", "(begin)\n\"B1\""),
            (
                "(define tally Tally!B1:E10) (first-cell-expand (ref (sheet-expand tally) (range-expand tally)))",
                "(begin)\n\"B1\"",
            ),
            ("(define tally Tally!B1:E10) (first-cell-expand (absolute tally))", "(begin)\n\"$B$1\""),
        ] {
            assert_eq!(one(src), want, "{src}");
        }
        let count = "(defmacro (count sheet range event) (formula sheet range \
                     (formula-expand (+ (last (ref sheet (first-cell-expand (ref sheet range)))) event)))) ";
        let at_a_word = one(&format!(
            "{count}(defmacro (at word call) (quasiquote ((unquote (car call)) (sheet-expand (unquote word)) \
             (range-expand (unquote word)) (unquote-splicing (cdr call))))) \
             (define tally Tally!B1:E10) (at tally (count Input!B16))"
        ));
        let row = "(formula \"Tally\" \"B1:E10\" \"=Tally.last!B1+Input!B16\")";
        assert_eq!(at_a_word.lines().last(), Some(row));
        assert_eq!(
            one(&format!("{count}(count \"Tally\" \"B1:E10\" Input!B16)")),
            row
        );
        assert_refusals(&[
            (
                "(first-cell-expand Screen!A:A)",
                "vla-reference-whole-line",
                "first-cell-expand: 'Screen!A:A' is a whole column, and first-cell-expand takes a cell or a range, such as Tally!B1:E10",
            ),
            ("(first-cell-expand Screen!3:3)", "vla-reference-whole-line", "'Screen!3:3' is a whole row"),
            (
                "(first-cell-expand B1:E10)",
                "vla-reference-no-sheet",
                "first-cell-expand: 'B1:E10' names no sheet",
            ),
            (
                "(first-cell-expand tally)",
                "vla-reference-expected",
                "first-cell-expand: expected one reference on a sheet",
            ),
            ("(first-cell-expand \"B1:E10\")", "vla-reference-expected", "', a text"),
            (
                "(first-cell-expand A1 B1)",
                "vla-expand-arity",
                "first-cell-expand expects exactly 1 argument(s), got 2",
            ),
            (
                "(&expand (first-cell-expand Tally!B1:E10))",
                "vla-expand-arg-is-list",
                "&expand: argument 1 is a list",
            ),
            (
                "(defmacro (first-cell-expand x) x)",
                "vla-defmacro-reserved-name",
                "'first-cell-expand' is the expand-time reference-first-cell operation and cannot be a macro name",
            ),
        ]);
    }

    #[test]
    fn the_walk_s_forms_read_what_a_head_writes() {
        // VLA.15, `TestAddressForms`' last part: one reading of a reference,
        // and the trampoline's begin kept before the first cell.
        assert_eq!(
            one("(define stop State!B12) (range-expand (absolute stop))"),
            "(begin)\n\"$B$12\""
        );
        assert_eq!(
            one("(define stop State!B12) (sheet-expand (last stop))"),
            "(begin)\n\"State.last\""
        );
        assert_eq!(
            one("(defmacro (m) (debug-print 1) (first-cell-expand Tally!B1:E10)) (m)"),
            "(begin (debug-print 1) \"B1\")"
        );
    }

    #[test]
    fn the_expander_crosses_threads_as_it_did() {
        // VLA.2 keeps a macro's template as nodes in the expander; a
        // published type's auto traits are part of its interface, so the
        // nodes count their references atomically and the expander stays
        // Send, Sync and unwind-safe, as `Macro` does.
        fn interface<T: Send + Sync + std::panic::UnwindSafe + std::panic::RefUnwindSafe>() {}
        interface::<Expander>();
        interface::<Macro>();
    }
}
