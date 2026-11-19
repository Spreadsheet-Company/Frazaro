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
//! cell, `first-cell-expand`, the third. And, since VLA.10, the place a verb
//! writes: `(bind-expand ((self reference)) form ...)` binds the word `self`
//! to one cell on a sheet for every form the walk meets in its body
//! (`ExpandBindCall`), and the walk alone reads the word, where it stands as
//! an expression, as a define's word is; `substitute` does not change, so a
//! macro's result is still a function of its template and its arguments. With
//! it `offset-expand`, Excel's `OFFSET` at expand time, a reference moved, a
//! form of the walk (`ExpandOffsetCall`), so that the words over `self` read
//! any reference. And, since VLA.8, functions at expand time: `(defexpand
//! (name parameter ... [& rest]) ["doc"] form)`, a pure function of its
//! arguments, its body a template of one form (`DefineFunction`), whose call
//! computes wherever a primitive's would, its arguments values and its value
//! standing where the call stood (`CallFunction`, `FunctionWorkOut`), a call
//! in a tail place looping and every other nesting to
//! [`FUNCTION_DEPTH_LIMIT`], the calls of one working-out counted to
//! [`FUNCTION_FUEL`]; `(let-expand ((name value) ...) form ...)`, a value bound
//! for its forms (`EvalLetExpand`); and a value read the same whoever wrote
//! it, a define's word its form and a call a word of the template holds
//! computed where a primitive reads its argument (`ValueRead`).
//!
//! Template-substitution macros: `(defmacro (name p1 p2 & rest) template
//! ...)`. A parameter is replaced by the form passed at the call site; the
//! rest parameter splices its forms into whatever list it stands in; several
//! template forms become one `(begin ...)`. No gensym and no eval.
//! Compile-time computation is the expand-time primitives and the functions
//! a source defines over them (VLA.8): each a pure function of its arguments
//! and the source's literals, never reading a file, a clock or anything two
//! compiles could see differently; every call counted against the fuel and
//! refused by name past it; and what it computes lands in the rows a reader
//! sees - by design. Since VLA.9 a
//! list in the signature, after the positional parameters, declares a named
//! piece, `(word default)` or `(word)`, which a call gives as `(word value)`
//! after its positional arguments, in any order; a piece left out takes its
//! default, substituted under the parameters bound before it.
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

use std::collections::{HashMap, HashSet};
use std::sync::atomic::{AtomicBool, AtomicU32, Ordering};
use std::sync::{Arc, Mutex, MutexGuard, OnceLock, PoisonError};

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
#[non_exhaustive]
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
    /// The named pieces, after `params` and in the declaration's order
    /// (VLA.9).
    pub pieces: Vec<Piece>,
}

/// A macro's named piece (VLA.9), as `DefineMacro` records it: `(word
/// default)`, or `(word)` for a piece the caller must give.
#[derive(Clone, Debug)]
#[non_exhaustive]
pub struct Piece {
    /// The word, as written.
    pub word: String,
    /// The default as written, or `None` for a piece the caller must give.
    pub default: Option<Form>,
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
    // VLA.10: the word the walk reads, the form that binds it, and a reference
    // moved.
    ("self", "place a verb writes"),
    ("bind-expand", "expand-time place-binding form"),
    ("offset-expand", "expand-time reference-moving operation"),
    // VLA.8: the form that binds a value for one form, and the form that
    // defines a function, which every entry point takes before any form
    // expands, so a macro of the name could never be called.
    ("let-expand", "expand-time value-binding form"),
    ("defexpand", "expand-time function-defining form"),
];

/// VLA.8: the most calls of functions one working-out nests, a call in a
/// tail place looping and nesting none (`VLA_FUNCTION_DEPTH_LIMIT`): half the
/// one the reference's stack gives out at where a call stands deepest, at the
/// form depth's limit, where Excel ran out 29 calls deep. On 1 MiB in the
/// package's unoptimized build a level of nested calls costs about 18 KB of
/// frames, and 22 held there and 23 did not, so 14 holds in both with
/// headroom, which `tests/limits.rs` holds.
pub const FUNCTION_DEPTH_LIMIT: u32 = 14;

/// VLA.8: the most calls one working-out makes, the chain's own number
/// (`VLA_FUNCTION_FUEL`).
pub const FUNCTION_FUEL: u32 = 5000;

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
/// VLA.8: and the expander's functions, its defines and the working-out under
/// way, which the reference reads from its module (`mFunctions`, `mDefines`,
/// `mFnDepth`), so that a template's primitive reads a define's word and calls
/// a function; `None` for a node a test builds bare.
#[derive(Default)]
struct Env<'r> {
    bindings: HashMap<&'r str, Node>,
    rest_name: &'r str,
    rest_items: &'r [Node],
    macros: Option<&'r HashMap<String, Rule>>,
    functions: Option<&'r HashMap<String, Arc<Function>>>,
    defines: Option<&'r HashMap<String, Node>>,
    calls: Option<&'r Calls>,
}

impl<'r> Env<'r> {
    /// The reference's empty scope, `New Collection, "", New Collection`:
    /// no bindings and no rest, the expander's tables kept.
    fn bare(&self) -> Env<'r> {
        Env {
            macros: self.macros,
            functions: self.functions,
            defines: self.defines,
            calls: self.calls,
            ..Env::default()
        }
    }
}

/// VLA.8: a function as the expander applies it, as `DefineFunction` records
/// it: its parameters folded, its rest's word folded or empty, its body one
/// form read into nodes once, and its name as written.
#[derive(Debug)]
struct Function {
    params: Vec<String>,
    rest_name: String,
    body: Node,
    name: String,
}

/// VLA.8: the working-out under way, the reference's `mFnDepth`, `mFnFuel`,
/// `mFnTrail` and `mFnTrailWritten`: the calls nested, the calls made, each
/// nested call's name and arguments outermost first, and whether a refusal
/// under way has its trail already. Empty at every entry and exit, since
/// [`function_work_out`] restores it on every way out. Atomics and a mutex,
/// never contended, so that [`Expander`] stays `Send`, `Sync` and
/// unwind-safe, as its interface promises.
#[derive(Default)]
struct Calls {
    depth: AtomicU32,
    fuel: AtomicU32,
    trail: Mutex<Vec<(String, Vec<Node>)>>,
    trail_written: AtomicBool,
}

impl Calls {
    fn depth(&self) -> u32 {
        self.depth.load(Ordering::Relaxed)
    }

    fn set_depth(&self, n: u32) {
        self.depth.store(n, Ordering::Relaxed);
    }

    /// The trail, a poisoned lock read as it stands: nothing panics holding
    /// it, and the trail is rebuilt at every outermost working-out.
    fn trail(&self) -> MutexGuard<'_, Vec<(String, Vec<Node>)>> {
        self.trail.lock().unwrap_or_else(PoisonError::into_inner)
    }
}

/// VLA.8: one call of a working-out: its function and its arguments' values,
/// the parameters' in order and the rest's after them.
struct CallState {
    function: Arc<Function>,
    values: Vec<Node>,
    rest: Vec<Node>,
}

impl CallState {
    /// The call as the trail keeps it: its name as written and every value.
    fn trail_entry(&self) -> (String, Vec<Node>) {
        let args = self.values.iter().chain(&self.rest).cloned().collect();
        (self.function.name.clone(), args)
    }
}

/// VLA.8: what one step of a working-out comes to: the value, or a call in a
/// tail place, which replaces the call under way.
enum Step {
    Value(Node),
    Tail(CallState),
}

/// A macro as the expander applies it (VLA.2): its parameters folded and
/// its template read into nodes once, when it is defined.
struct Rule {
    params: Vec<String>,
    /// VLA.9: the parameters as written, for a refusal's words.
    param_words: Vec<String>,
    rest_name: String,
    template: Node,
    /// VLA.9: the named pieces in the declaration's order, none for most.
    pieces: Vec<RulePiece>,
    /// VLA.9: where each piece stands in `pieces`, by its word folded, so
    /// that a stranger's call of thousands of pieces stays linear.
    piece_index: HashMap<String, usize>,
}

/// A named piece as the expander applies it (VLA.9): its word folded, and as
/// written for a refusal, and its default read into nodes once.
struct RulePiece {
    word: String,
    word_fold: String,
    default: Option<Node>,
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
    /// VLA.10, `mSelfStack`: the places `bind-expand` holds open, innermost
    /// last, each a form the walk has finished with. Empty at every entry and
    /// exit of [`Expander::expand`], a refusal's exit among them, since each
    /// is popped before the `?` that would leave it.
    self_stack: Vec<Node>,
    /// VLA.8, `mFunctions`: the functions a source defines, by the folded
    /// name, taken before any form expands as the macros are.
    functions: HashMap<String, Arc<Function>>,
    /// VLA.8: the working-out under way, empty at every entry and exit.
    calls: Calls,
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
        // VLA.8: a function of the name, defined before it, whose calls the
        // walks compute before a macro's would apply: a word names one thing.
        if self.functions.contains_key(&mfold) {
            return Err(raise("vla-defmacro-name-taken", &[("name", &mname)]));
        }

        let mut params = Vec::new();
        let mut pieces = Vec::new();
        let mut rest_name = String::new();
        // VLA.9: every word declared so far, folded, the reference's scan as
        // a set, since a stranger's signature may hold a hundred thousand.
        let mut declared = HashSet::new();
        let mut j = 2;
        while j <= sig.count() {
            if let Form::List(decl) = sig.nth(j)? {
                pieces.push(declare_piece(&mname, decl, &mut declared)?);
                j += 1;
                continue;
            }
            let t = sym_text(sig.nth(j)?)?;
            if t == "&" {
                if !pieces.is_empty() {
                    return Err(raise("vla-defmacro-piece-with-rest", &[("name", &mname)]));
                }
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
                refuse_word_twice(&mname, &rest_name, &mut declared)?;
                break;
            }
            if !pieces.is_empty() {
                return Err(raise(
                    "vla-defmacro-param-after-piece",
                    &[("name", &mname), ("word", &t)],
                ));
            }
            refuse_word_twice(&mname, &t, &mut declared)?;
            params.push(t);
            j += 1;
        }
        let piece_index: HashMap<String, usize> = pieces
            .iter()
            .enumerate()
            .map(|(k, piece): (usize, &Piece)| (fold(&piece.word), k))
            .collect();
        refuse_defaults_unbound(&mname, &pieces, &piece_index)?;

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
                param_words: params.clone(),
                rest_name: fold(&rest_name),
                template: Node::of(&template),
                pieces: pieces.iter().map(RulePiece::of).collect(),
                piece_index,
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
                pieces,
            },
        );
        Ok(())
    }

    /// VLA.8: a definition taken, a macro's or a function's, as every entry
    /// point takes one before any form expands: `true` when the form was a
    /// definition, `false` for any other form, which is the caller's to
    /// expand. A client that loops over a source's forms calls this, so that
    /// a definition the language adds later asks it to change nothing.
    pub fn define(&mut self, f: &Form) -> Result<bool, Refusal> {
        let Form::List(l) = f else {
            return Ok(false);
        };
        if f.head_is("defmacro") {
            self.define_macro(l)?;
            return Ok(true);
        }
        if f.head_is("defexpand") {
            self.define_function(l)?;
            return Ok(true);
        }
        Ok(false)
    }

    /// VLA.8, VBA `DefineFunction`: `(defexpand (name parameter ... [&
    /// rest]) ["doc"] form)`. In one order, so that the first fault is the
    /// one named: the shape; the name, a word of the source's own by a
    /// define's rules, neither a macro's nor a function's already; the
    /// parameters, positional and a rest, no named piece, none twice; the
    /// docstring, L17's rule; and one form, the body.
    fn define_function(&mut self, f: &List) -> Result<(), Refusal> {
        let shape = || {
            raise(
                "vla-defexpand-shape",
                &[("form", &write_datum(&Form::List(f.clone())))],
            )
        };
        if f.count() < 3 {
            return Err(shape());
        }
        let Form::List(sig) = f.nth(2)? else {
            return Err(shape());
        };
        let Some(Form::Sym(fname)) = sig.items.first() else {
            return Err(shape());
        };
        let fname = fname.clone();
        if let Some(what) = define_reads_as(&fname) {
            return Err(raise(
                "vla-defexpand-name-reference",
                &[("word", &fname), ("what", what)],
            ));
        }
        if let Some(desc) = define_reserved_desc(&fname) {
            return Err(raise(
                "vla-defexpand-name-reserved",
                &[("word", &fname), ("desc", &desc)],
            ));
        }
        if self.get_macro(&fname).is_some() {
            return Err(raise(
                "vla-defexpand-name-taken",
                &[("word", &fname), ("what", "a macro's name")],
            ));
        }
        if self.functions.contains_key(&fold(&fname)) {
            return Err(raise(
                "vla-defexpand-name-taken",
                &[("word", &fname), ("what", "defined already")],
            ));
        }

        let mut params: Vec<String> = Vec::new();
        let mut rest_name = String::new();
        let mut j = 2;
        while j <= sig.count() {
            if let Form::List(_) = sig.nth(j)? {
                return Err(raise(
                    "vla-defexpand-piece",
                    &[("name", &fname), ("value", &write_datum(sig.nth(j)?))],
                ));
            }
            let t = sym_text(sig.nth(j)?)?;
            if t == "&" {
                if j + 1 != sig.count() {
                    return Err(raise("vla-defexpand-rest", &[("name", &fname)]));
                }
                if let Form::List(_) = sig.nth(j + 1)? {
                    return Err(raise("vla-defexpand-rest", &[("name", &fname)]));
                }
                let rest = sym_text(sig.nth(j + 1)?)?;
                if params.contains(&fold(&rest)) {
                    return Err(raise(
                        "vla-defexpand-param-twice",
                        &[("name", &fname), ("word", &rest)],
                    ));
                }
                rest_name = fold(&rest);
                break;
            }
            if params.contains(&fold(&t)) {
                return Err(raise(
                    "vla-defexpand-param-twice",
                    &[("name", &fname), ("word", &t)],
                ));
            }
            params.push(fold(&t));
            j += 1;
        }

        // L17: a string at form 3 is the docstring only when a form follows it.
        let mut doc_idx = 3;
        if f.count() >= 4 {
            if let Form::Str(_) = f.nth(3)? {
                doc_idx = 4;
            }
        }
        let forms = f.count() + 1 - doc_idx;
        if forms != 1 {
            return Err(raise(
                "vla-defexpand-body",
                &[("name", &fname), ("n", &forms.to_string())],
            ));
        }
        self.functions.insert(
            fold(&fname),
            Arc::new(Function {
                params,
                rest_name,
                body: Node::of(f.nth(doc_idx)?),
                name: fname,
            }),
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
                // VLA.10: the word self inside a binding, the place it holds
                // put where the word stands and looked at again, one link of
                // the chain, as a define's word is. With no binding open the
                // word stays the word it is in a program.
                if is_self_word(s) {
                    if let Some(v) = self.self_stack.last() {
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
            // VLA.10: a place bound for the forms of its body, and a reference
            // moved, forms of the walk.
            if head_fold == "bind-expand" {
                let body = self.bind_expand(items, depth, walk)?;
                return self.finish(&pending_head, body, depth, walk);
            }
            if head_fold == "offset-expand" {
                let moved = self.offset_expand(items, depth, walk)?;
                return self.finish(&pending_head, moved, depth, walk);
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
            // VLA.8: a value bound for one form outside any template, its forms
            // handed back to this loop so that its rows expand, one link of the
            // chain as a loop's begin is; and a function's call the walk meets,
            // its value looked at again, one application, standing at a spent
            // budget as a macro's call does ([`Expander::walk_computes`]).
            if head_fold == "let-expand" || self.functions.contains_key(&head_fold) {
                match self.walk_step(&head_fold, items, &mut chain_len)? {
                    Some(next) => {
                        cur = next;
                        continue;
                    }
                    None => break, // L15
                }
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
                let r = expand_one(items, rule, &self.bare_env(), lst.line)?;
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
    /// text `ref` takes (VBA `FsFirstCellText`). VLA.10: at a spent budget it
    /// stands ([`Expander::budget_spent`]).
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
        if self.budget_spent() {
            return Ok(form_of_the_walk_standing(&items[0], read));
        }
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
    /// the macros, which a loop asks of its list (VLA.5), and since VLA.8 the
    /// functions, the defines and the working-out, which a primitive's
    /// argument reads.
    fn bare_env(&self) -> Env<'_> {
        Env {
            macros: Some(&self.rules),
            functions: Some(&self.functions),
            defines: Some(&self.defines),
            calls: Some(&self.calls),
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
    /// text there and not the formula it wrote. VLA.10: at a spent budget it
    /// stands ([`Expander::budget_spent`]); and the word `self` left in its
    /// expression is refused ([`holds_self`]), since the walk replaced every
    /// one it read, so that none is written as a name.
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
        let expr = expr?;
        if self.budget_spent() {
            return Ok(form_of_the_walk_standing(&items[0], expr));
        }
        if holds_self(&expr) {
            return Err(raise("vla-self-free", &[]));
        }
        let text = formula_syntax::write(&expr)?;
        // VLA.6: kept for the form `expand` expands, so that a word a define
        // of that form binds, written here as one of Excel's names, is
        // refused (`DefinesRefuseFormulaUse`).
        self.formula_texts.push(text.clone());
        Ok(Form::Str(text))
    }

    /// L15 and VLA.10, VBA `ExpandBudgetSpent`: whether the stepper's budget
    /// is spent. No application fires after it, so a form of the walk that
    /// read its argument now would read a macro not yet fired, and be refused,
    /// or write a macro's call as a function's: found while scoping VLA.10, a
    /// verb's every frame refused since VLA.4. Each form stands instead, and
    /// computes on a later frame or at the fixpoint.
    fn budget_spent(&self) -> bool {
        self.budget_on && self.expand_fired >= self.expand_budget
    }

    /// VLA.10, VBA `ExpandBindCall`: `(bind-expand ((self reference)) form
    /// ...)`, the place a verb writes, bound for every form the walk meets in
    /// its body. One order, so that the first fault a form holds is the one
    /// named: its shape and its word ([`bind_value`]); its value, expanded to
    /// its fixpoint under the binding outside it, `let`'s rule, so that no read
    /// can loop; at a spent budget the form stands ([`bind_standing`]); its
    /// value read as one cell on a sheet ([`formula_syntax::bound_cell`]). Then
    /// the body to its fixpoint inside the binding, and never handed back to
    /// the trampoline, whose set-aside forms would expand after the pop. The
    /// stack is restored before the `?` that would leave it, on every way out.
    fn bind_expand(&mut self, items: &[Node], depth: u32, walk: u32) -> Result<Form, Refusal> {
        let value = bind_value(items)?;
        let bound = self.expand_at(value, depth, walk + 1)?;
        if self.budget_spent() {
            return Ok(bind_standing(items, bound));
        }
        formula_syntax::bound_cell(&bound)?;
        self.self_stack.push(Node::of(&bound));
        let body = self.bind_body(&items[2..], depth, walk);
        self.self_stack.pop();
        body
    }

    /// VLA.10, VBA `ExpandBindBody`: a binding's body, each form to its
    /// fixpoint in written order: one form its own result, and several one
    /// `begin` with no line, as a template's several forms are. A chain inside
    /// a form sets its own forms aside and expands them before
    /// [`Expander::expand_at`] returns, so all of it stands inside the binding.
    fn bind_body(&mut self, forms: &[Node], depth: u32, walk: u32) -> Result<Form, Refusal> {
        if let [one] = forms {
            return self.expand_at(one, depth, walk + 1);
        }
        let mut out = Vec::with_capacity(forms.len() + 1);
        out.push(Form::sym("begin"));
        for f in forms {
            out.push(self.expand_at(f, depth, walk + 1)?);
        }
        Ok(Form::list(out))
    }

    /// VLA.10, VBA `ExpandOffsetCall`: `(offset-expand reference rows columns)`
    /// and `(offset-expand reference rows columns height width)`, Excel's
    /// `OFFSET` at expand time, in its argument order. One order, so that the
    /// first fault a form holds is the one named: the count; every argument
    /// expanded to its fixpoint, left to right, so that `self` or a define's
    /// word is replaced first; at a spent budget the form stands; each number
    /// through the arithmetic's gate, left to right
    /// ([`offset_whole_number`]); and the reference moved
    /// ([`formula_syntax::offset_form`]). No primitive, since `self` and a
    /// define's word reach none.
    fn offset_expand(&mut self, items: &[Node], depth: u32, walk: u32) -> Result<Form, Refusal> {
        if items.len() != 4 && items.len() != 6 {
            return Err(raise(
                "vla-offset-expand-arity",
                &[("n", &(items.len() - 1).to_string())],
            ));
        }
        let mut read = Vec::with_capacity(items.len());
        read.push(items[0].form());
        for item in &items[1..] {
            read.push(self.expand_at(item, depth, walk + 1)?);
        }
        if self.budget_spent() {
            return Ok(Form::list(read));
        }
        let mut numbers = Vec::with_capacity(4);
        for v in &read[2..] {
            numbers.push(offset_whole_number(v)?);
        }
        formula_syntax::offset_form(&read[1], &numbers)
    }

    /// VLA.8: the walk's step at a let-expand or a function's call: at a spent
    /// budget a function's call stands (`None`), as a macro's does; else one
    /// link of the chain, refused past 5,000, then what
    /// [`Expander::walk_computes`] computes. A method of its own, so that the
    /// walk's recursive frame holds none of it.
    #[inline(never)]
    fn walk_step(
        &mut self,
        head_fold: &str,
        items: &[Node],
        chain_len: &mut u32,
    ) -> Result<Option<Node>, Refusal> {
        if head_fold != "let-expand" && self.budget_spent() {
            return Ok(None);
        }
        *chain_len += 1;
        if *chain_len > 5000 {
            return Err(raise("vla-macro-self-expansion-limit", &[]));
        }
        self.walk_computes(head_fold, items).map(Some)
    }

    /// VLA.8: what the walk computes at a form's head outside any template: a
    /// let-expand's forms under its binding (`EvalLetExpand`), or a function's
    /// value, one application, weighed as a macro's result is
    /// (`CallFunction`). A method of its own, so that the walk's recursive
    /// frame holds none of it.
    fn walk_computes(&mut self, head_fold: &str, items: &[Node]) -> Result<Node, Refusal> {
        if head_fold == "let-expand" {
            return eval_let_expand(items, &self.bare_env());
        }
        let value = call_function(head_fold, items, &self.bare_env())?;
        self.weigh(&value)?;
        self.expand_fired += 1; // L1: one per application
        Ok(value)
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
                    let r = expand_one(items, rule, &self.bare_env(), lst.line)?;
                    self.weigh(&r)?;
                    self.expand_fired += 1;
                    return Ok(r.form());
                }
            }
            // VLA.8: a function's call, replaced by its value in one step, as a
            // macro's call is by its template.
            if self.functions.contains_key(&head_fold) && !self.budget_spent() {
                let r = call_function(&head_fold, items, &self.bare_env())?;
                self.weigh(&r)?;
                self.expand_fired += 1;
                return Ok(r.form());
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
    /// since VLA.8 every `defexpand` ([`Expander::define`]), and the other
    /// forms come back as the body.
    pub fn collect(&mut self, forms: Vec<Form>) -> Result<Vec<Form>, Refusal> {
        let mut body = Vec::with_capacity(forms.len());
        for f in forms {
            if !self.define(&f)? {
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
        // VLA.8: a function's name, which the walks call.
        if self.functions.contains_key(&fold(word)) {
            return Err(raise(
                "vla-define-name-taken",
                &[("word", word), ("what", "a function's name")],
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

/// VLA.8: a value held past one step of the expander's computing, a
/// function's argument, the value a working-out answers and a let-expand's
/// name, weighed by its nesting, as the walk weighs what it reads (VLA.1):
/// past the reader's limit, refused by [`too_deep`]. A call that wraps its
/// argument nests it a level deeper, so a loop within the fuel could nest a
/// value thousands deep and more, which no walk reads and whose freeing alone
/// overflows the stack. The depth was made with the value, so the test reads
/// it. The reference shares a value by reference and holds it to no depth,
/// meeting the fuel instead, as `VLA.1`'s growing macros meet the chain's
/// limit there.
#[inline]
fn held_depth(v: &Node) -> Result<(), Refusal> {
    match v {
        Node::List(l) if l.depth > FORM_DEPTH_LIMIT => Err(held_too_deep(l.line)),
        _ => Ok(()),
    }
}

/// [`held_depth`]'s refusal, made off the path of the calls it guards, so
/// that their frames hold none of it.
#[cold]
#[inline(never)]
fn held_too_deep(line: u32) -> Refusal {
    too_deep(line)
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

/// VBA `DeclarePiece`: a list in a macro's signature, a named piece's
/// declaration (VLA.9). Refused by name: any shape but one word, or a word
/// and one form, the rest's mark being no piece's word; a word declared
/// before it; and a word the walks read as a form of their own, since a
/// template holding a call would read such a piece as that form before the
/// call fires.
fn declare_piece(
    mname: &str,
    decl: &List,
    declared: &mut HashSet<String>,
) -> Result<Piece, Refusal> {
    let word = match (decl.count(), decl.items.first()) {
        (1 | 2, Some(Form::Sym(word))) if word != "&" => word.clone(),
        _ => {
            return Err(raise(
                "vla-defmacro-piece-shape",
                &[
                    ("name", mname),
                    ("value", &write_datum(&Form::List(decl.clone()))),
                ],
            ))
        }
    };
    refuse_word_twice(mname, &word, declared)?;
    let word_fold = fold(&word);
    if let Some((_, desc)) = RESERVED.iter().find(|(w, _)| *w == word_fold) {
        return Err(raise(
            "vla-defmacro-piece-reserved",
            &[("name", mname), ("word", &word), ("desc", desc)],
        ));
    }
    Ok(Piece {
        word,
        default: decl.items.get(1).cloned(),
    })
}

/// VBA `RefuseWordTwice`: a word a signature has declared already,
/// positional, piece or rest, compared as a call binds them, by the fold
/// (VLA.9).
fn refuse_word_twice(
    mname: &str,
    word: &str,
    declared: &mut HashSet<String>,
) -> Result<(), Refusal> {
    if !declared.insert(fold(word)) {
        return Err(raise(
            "vla-defmacro-param-twice",
            &[("name", mname), ("word", word)],
        ));
    }
    Ok(())
}

/// VBA `RefuseDefaultsUnbound`: a default may name only the parameters
/// declared before it, so that the binding is one pass left to right and no
/// ring can form: its own word or a later piece's is refused at the
/// definition, the first such word its walk meets (VLA.9).
fn refuse_defaults_unbound(
    mname: &str,
    pieces: &[Piece],
    piece_index: &HashMap<String, usize>,
) -> Result<(), Refusal> {
    for (k, piece) in pieces.iter().enumerate() {
        let Some(default) = &piece.default else {
            continue;
        };
        if let Some(other) = default_names_from(default, k, piece_index) {
            return Err(raise(
                "vla-defmacro-piece-default-unbound",
                &[
                    ("name", mname),
                    ("word", &piece.word),
                    ("other", &pieces[other].word),
                ],
            ));
        }
    }
    Ok(())
}

/// VBA `DefaultNamesFrom`: the first piece from `from` on whose word a
/// default holds. Every word of the form counts, a head's and a quoted one's
/// among them, as [`substitute`] replaces them all. Bounded by the reader's
/// depth (VLA.1).
fn default_names_from(
    d: &Form,
    from: usize,
    piece_index: &HashMap<String, usize>,
) -> Option<usize> {
    match d {
        Form::Sym(s) => piece_index.get(&fold(s)).copied().filter(|&k| k >= from),
        Form::Str(_) => None,
        Form::List(l) => l
            .items
            .iter()
            .find_map(|e| default_names_from(e, from, piece_index)),
    }
}

impl RulePiece {
    fn of(piece: &Piece) -> RulePiece {
        RulePiece {
            word: piece.word.clone(),
            word_fold: fold(&piece.word),
            default: piece.default.as_ref().map(Node::of),
        }
    }
}

/// VBA `PieceWrittenAs`: which of a macro's pieces a form is written as, the
/// piece whose word heads it, by the fold; none for an atom, the empty list
/// and a list headed by anything else (VLA.9).
fn piece_written_as(v: &Node, piece_index: &HashMap<String, usize>) -> Option<usize> {
    let Node::List(lst) = v else {
        return None;
    };
    let Some(Node::Sym(head)) = lst.items().first() else {
        return None;
    };
    piece_index.get(&fold(head)).copied()
}

/// VBA `RefusePieceForParam`: a positional argument written as one of the
/// macro's own pieces is refused, left to right, so that a call the language
/// takes reads the same by position and by its pieces' words, and a caller
/// who leaves out a positional argument hears so, where the first piece would
/// be bound to it (VLA.9).
fn refuse_piece_for_param(call: &[Node], rule: &Rule, line: u32) -> Result<(), Refusal> {
    for (j, param) in rule.param_words.iter().enumerate() {
        if let Some(k) = piece_written_as(&call[j + 1], &rule.piece_index) {
            return Err(raise(
                "vla-macro-piece-for-param",
                &[
                    ("name", &call[0].atom_text().unwrap_or_default()),
                    ("value", &write_datum(&call[j + 1].form())),
                    ("param", param),
                    ("word", &rule.pieces[k].word),
                    ("loc", &src_line_tag(line, 0)),
                ],
            ));
        }
    }
    Ok(())
}

/// VBA `BindPieces`: the arguments after a macro's positional ones, each one
/// of its pieces (VLA.9). In one order, so that the first fault a call holds
/// is the one named: the arguments left to right; then the pieces without a
/// default, in the declaration's order; and only then every piece bound, in
/// the declaration's order and never the call's: a given one to its value
/// itself, its line with it, and a left-out one to its default substituted
/// under the parameters bound before it, let*'s reading, so one call written
/// in two orders writes the same rows.
fn bind_pieces<'r>(
    call: &'r [Node],
    rule: &'r Rule,
    line: u32,
    env: &mut Env<'r>,
) -> Result<(), Refusal> {
    let name = call[0].atom_text().unwrap_or_default();
    let mut given: Vec<Option<&Node>> = vec![None; rule.pieces.len()];
    for arg in &call[rule.params.len() + 1..] {
        let Some(k) = piece_written_as(arg, &rule.piece_index) else {
            let pieces: Vec<&str> = rule.pieces.iter().map(|p| p.word.as_str()).collect();
            return Err(raise(
                "vla-macro-piece-unknown",
                &[
                    ("name", &name),
                    ("value", &write_datum(&arg.form())),
                    ("pieces", &pieces.join(", ")),
                    ("loc", &src_line_tag(line, 0)),
                ],
            ));
        };
        let Node::List(piece) = arg else {
            unreachable!("piece_written_as answers only a list")
        };
        if piece.count() != 2 {
            return Err(raise(
                "vla-macro-piece-arity",
                &[
                    ("name", &name),
                    ("word", &rule.pieces[k].word),
                    ("n", &(piece.count() - 1).to_string()),
                    ("loc", &src_line_tag(line, 0)),
                ],
            ));
        }
        if given[k].is_some() {
            return Err(raise(
                "vla-macro-piece-twice",
                &[
                    ("name", &name),
                    ("word", &rule.pieces[k].word),
                    ("loc", &src_line_tag(line, 0)),
                ],
            ));
        }
        given[k] = Some(&piece.items()[1]);
    }
    for (piece, value) in rule.pieces.iter().zip(&given) {
        if value.is_none() && piece.default.is_none() {
            return Err(raise(
                "vla-macro-piece-missing",
                &[
                    ("name", &name),
                    ("word", &piece.word),
                    ("loc", &src_line_tag(line, 0)),
                ],
            ));
        }
    }
    for (piece, value) in rule.pieces.iter().zip(&given) {
        let bound = match (value, &piece.default) {
            (Some(v), _) => (*v).clone(),
            (None, Some(default)) => substitute(default, env, false)?,
            (None, None) => unreachable!("a piece with neither was refused above"),
        };
        env.bindings.insert(&piece.word_fold, bound);
    }
    Ok(())
}

/// VLA.10, VBA `IsSelfWord`: whether a word is `self`, by the fold, folding
/// nothing: `fold` maps A to Z alone, which an ASCII comparison without case
/// is. Asked of every word the walk meets inside a binding.
fn is_self_word(s: &str) -> bool {
    s.eq_ignore_ascii_case("self")
}

/// VLA.10, VBA `FormOfTheWalkStanding`: a form of the walk standing at a spent
/// budget, its head as written and its one argument as far as the budget took
/// it.
fn form_of_the_walk_standing(head: &Node, arg: Form) -> Form {
    Form::list(vec![head.form(), arg])
}

/// VLA.10, VBA `FormHoldsSelf`: whether a formula's finished expression holds
/// the word `self` where a value stands: never a list's head, whose atom is a
/// function's name, and never inside a quote, an array's data. The walk
/// replaced every one it read, so one left here is free, or stands in a
/// primitive's chosen branch, which the walk does not expand (VLA.5). At most
/// the form's own nesting.
fn holds_self(f: &Form) -> bool {
    match f {
        Form::Sym(s) => is_self_word(s),
        Form::Str(_) => false,
        Form::List(l) => {
            if f.head_is("quote") {
                return false;
            }
            let values = match l.items.first() {
                Some(Form::Sym(_) | Form::Str(_)) => &l.items[1..],
                _ => &l.items[..],
            };
            values.iter().any(holds_self)
        }
    }
}

/// VLA.10, VBA `RefuseSelfText`: a primitive that reads a word's text refuses
/// the word `self`, free or bound, since the walk alone reads it and a
/// primitive takes its arguments as data; so a later reading there moves no
/// source. A text so spelled is a text.
fn refuse_self_text(op: &str, v: &Node) -> Result<(), Refusal> {
    match v {
        Node::Sym(s) if is_self_word(s) => Err(raise("vla-self-in-a-primitive", &[("op", op)])),
        _ => Ok(()),
    }
}

/// VLA.10, VBA `BindValueForm`: a binding's shape and its word, and the form it
/// binds. A list of bindings, so that a later item can open a declared word
/// with no source read again, each `(word form)`; every word but `self`
/// refused by name; then one binding, and at least one form after it.
fn bind_value(items: &[Node]) -> Result<&Node, Refusal> {
    let shape = || {
        raise(
            "vla-bind-expand-shape",
            &[("form", &write_datum(&Node::list(items.to_vec()).form()))],
        )
    };
    let [_, Node::List(bindings), _, ..] = items else {
        return Err(shape());
    };
    if bindings.count() == 0 {
        return Err(shape());
    }
    for binding in bindings.items() {
        let Node::List(binding) = binding else {
            return Err(shape());
        };
        let [Node::Sym(word), _] = binding.items() else {
            return Err(shape());
        };
        if !is_self_word(word) {
            return Err(raise("vla-bind-expand-word", &[("word", word)]));
        }
    }
    let [Node::List(binding)] = bindings.items() else {
        return Err(shape());
    };
    Ok(&binding.items()[1])
}

/// VLA.10, VBA `BindStanding`: a `bind-expand` standing at a spent budget: its
/// word as written, its value as far as the budget took it, and its body as
/// written. [`bind_value`] has read its shape.
fn bind_standing(items: &[Node], bound: Form) -> Form {
    let word = match &items[1] {
        Node::List(bindings) => match &bindings.items()[0] {
            Node::List(binding) => binding.items()[0].form(),
            other => other.form(),
        },
        other => other.form(),
    };
    let binding = Form::list(vec![word, bound]);
    let mut out = Vec::with_capacity(items.len());
    out.push(items[0].form());
    out.push(Form::list(vec![binding]));
    out.extend(items[2..].iter().map(Node::form));
    Form::list(out)
}

/// VLA.10, VBA `OffsetWholeNumber`: one number of `offset-expand`, through the
/// arithmetic's gate: an atom, a number literal, and a whole one.
fn offset_whole_number(v: &Form) -> Result<f64, Refusal> {
    const OP: &str = "offset-expand";
    let text = number_text(OP, &Node::of(v))?;
    let n = val(&text);
    if n.fract() != 0.0 {
        return Err(raise(
            "vla-arith-expand-not-whole",
            &[("op", OP), ("value", &text)],
        ));
    }
    Ok(n)
}

/// VBA `ExpandOne`: one application of a macro to a call form's items. A
/// parameter is bound to its argument itself, not a copy of it, and the
/// rest parameter to the call's own items after the parameters'; the
/// arguments after the positional ones are its named pieces, for a macro
/// that declares any (VLA.9). `line` is the call's own, 0 for one a macro
/// made, which the refusals of its pieces name.
fn expand_one(call: &[Node], rule: &Rule, tables: &Env, line: u32) -> Result<Node, Refusal> {
    let argc = call.len() - 1;
    let name = call[0].atom_text().unwrap_or_default();
    if rule.rest_name.is_empty() && rule.pieces.is_empty() {
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
    let named = !rule.pieces.is_empty();
    if named {
        refuse_piece_for_param(call, rule, line)?;
    }
    let mut env = Env {
        bindings: HashMap::new(),
        rest_name: &rule.rest_name,
        rest_items: if named {
            &[]
        } else {
            &call[rule.params.len() + 1..]
        },
        macros: tables.macros,
        functions: tables.functions,
        defines: tables.defines,
        calls: tables.calls,
    };
    for (j, p) in rule.params.iter().enumerate() {
        env.bindings.insert(p, call[j + 1].clone());
    }
    if named {
        bind_pieces(call, rule, line, &mut env)?;
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
            // VLA.8: a value bound for one form, and a function's call, its head
            // the function's name or a word of the template naming one; shielded
            // data inside a quasiquote, as the primitives are. One arm and one
            // call, so that this recursive frame grows the least it can.
            _ if !in_quasi
                && (h == "let-expand" || function_head_name(&items[0], env).is_some()) =>
            {
                return computes_in_template(&h, items, env);
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
        refuse_self_text("symbol", &piece)?; // VLA.10
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
//
// VLA.16: a list of names where the name stands, `(each ((address value)
// entries) form ...)` and `(each-index (index (color name) entries 0) form
// ...)`, binds each name to its part of the element, so that a template
// names an entry's parts by what they are and never as `(car (cdr entry))`.
// The element a list of as many parts as there are names, or refused by
// name with the entry shown; no rest and no list among the names.

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
/// port clones its map, which answers every form the same. VLA.16: where a
/// list of names stands for the name, each name is checked in turn, none
/// alike another or the index, and each is bound to its part of the
/// element ([`each_entry_parts`]).
fn each_results(each_form: &[Node], env: &Env) -> Result<Vec<Node>, Refusal> {
    let op = match &each_form[0] {
        Node::Sym(s) => fold(s),
        _ => unreachable!("a loop is headed by its name"),
    };
    let binding = each_binding(each_form, &op)?;
    let indexed = op == "each-index";
    let (index_name, names_at, list_at) = if indexed {
        (Some(symbol_text(&binding[0])), 1, 2)
    } else {
        (None, 0, 1)
    };
    let names_form = &binding[names_at];
    let (names, parted) = match names_form {
        Node::List(l) => (l.items().iter().map(symbol_text).collect(), true),
        _ => (vec![symbol_text(names_form)], false),
    };
    if let Some(index_name) = &index_name {
        each_check_name(&op, index_name, env)?;
    }
    let index_fold = index_name.as_deref().map(fold);
    let mut folds: Vec<String> = Vec::with_capacity(names.len());
    for name in &names {
        each_check_name(&op, name, env)?;
        let name_fold = fold(name);
        if index_fold.as_deref() == Some(name_fold.as_str()) || folds.contains(&name_fold) {
            return Err(raise("vla-each-name-bound", &[("op", &op), ("word", name)]));
        }
        folds.push(name_fold);
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
        functions: env.functions,
        defines: env.defines,
        calls: env.calls,
    };
    let mut results = Vec::new();
    for (k, element) in walked.items().iter().enumerate() {
        if parted {
            let parts = each_entry_parts(&op, names_form, element)?;
            for (name_fold, part) in folds.iter().zip(parts) {
                scope.bindings.insert(name_fold.as_str(), part.clone());
            }
        } else {
            scope.bindings.insert(folds[0].as_str(), element.clone());
        }
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

/// VLA.16, VBA `EachEntryParts`: an element read by a loop's list of
/// names, a list of as many parts as there are names, its parts in order;
/// or the refusal that shows the names, the element and what it is.
fn each_entry_parts<'a>(op: &str, names: &Node, entry: &'a Node) -> Result<&'a [Node], Refusal> {
    let Node::List(want) = names else {
        unreachable!("a loop reads its element's parts by a list of names")
    };
    let held = match entry {
        Node::List(parts) if parts.count() == want.count() => return Ok(parts.items()),
        Node::List(parts) => format!("a list of {}", parts.count()),
        _ => "no list".to_string(),
    };
    Err(raise(
        "vla-each-entry-shape",
        &[
            ("op", op),
            ("names", &write_datum(&names.form())),
            ("n", &want.count().to_string()),
            ("entry", &write_datum(&entry.form())),
            ("held", &held),
        ],
    ))
}

/// VBA `EachBinding`: the loop's binding, or the refusal that shows its
/// shape: `(name list)` for `each`, `(index name list start)` for
/// `each-index`, each name a symbol, and at least one form after it.
/// VLA.16: the name, `each`'s first and `each-index`'s second, may be a
/// list of names, at least one and each a symbol ([`each_names_ok`]).
fn each_binding<'a>(each_form: &'a [Node], op: &str) -> Result<&'a [Node], Refusal> {
    let want = if op == "each-index" { 4 } else { 2 };
    if each_form.len() >= 3 {
        if let Node::List(b) = &each_form[1] {
            let items = b.items();
            if items.len() == want
                && (want == 2 || matches!(items[0], Node::Sym(_)))
                && each_names_ok(&items[want / 2 - 1])
            {
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

/// VLA.16: where a loop's name stands, a symbol, or a list of names, at
/// least one and each a symbol: no rest and no list among them.
fn each_names_ok(n: &Node) -> bool {
    match n {
        Node::Sym(_) => true,
        Node::List(l) => l.count() > 0 && l.items().iter().all(|p| matches!(p, Node::Sym(_))),
        _ => false,
    }
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

// VLA.8: functions at expand time, and let-expand. `(defexpand (name
// parameter ... [& rest]) ["doc"] form)` defines a pure function of its
// arguments, its body a template of one form; its call computes wherever a
// primitive's would, its arguments read as a primitive reads its own, and its
// value stands where the call stood. `(let-expand ((name value) ...) form
// ...)` binds each value, read in order under the names before it, for its
// forms. And where a primitive reads its argument, a define's word is its
// form and a call a word of the template holds is computed ([`value_read`]).
// No gensym and no eval: every name is the author's, a function reads its
// arguments and the source's literals alone, and every call counts against
// the fuel ([`FUNCTION_FUEL`]) and every nesting against the depth
// ([`FUNCTION_DEPTH_LIMIT`]), each refused by name. The reference keeps a
// call under way in variables it replaces on a tail call; the port keeps it
// as owned state ([`CallState`]) and reads a body's tail by recursion over its
// written nesting ([`tail_of`]), the same steps in the same order.

/// VLA.8, VBA `FunctionHeadName`: the folded name of the function a call's
/// head names: the head itself when it is a function's name and no word of
/// the template, or the word's value when a word of the template stands
/// there bound to a function's name, so that a template handed a function
/// calls it as `(f x)`. A word of the template hides a function of its own
/// spelling, as it hides a define's word.
fn function_head_name(head: &Node, env: &Env) -> Option<String> {
    let functions = env.functions?;
    if functions.is_empty() {
        return None;
    }
    let Node::Sym(h) = head else {
        return None;
    };
    let mut hf = fold(h);
    if let Some(v) = env.bindings.get(hf.as_str()) {
        let Node::Sym(s) = v else {
            return None;
        };
        hf = fold(s);
    }
    functions.contains_key(&hf).then_some(hf)
}

/// VLA.8: [`substitute`]'s arm, VBA `Substitute`'s `let-expand` and `Case
/// Else`: a let-expand's value, or a function's call computed in the
/// template's scope, its arguments bound and worked out with no frame
/// between, since a call nested in a call's argument stacks every frame of
/// this path once a level (VLA.1's 1 MiB).
fn computes_in_template(h: &str, items: &[Node], env: &Env) -> Result<Node, Refusal> {
    if h == "let-expand" {
        return eval_let_expand(items, env);
    }
    match function_head_name(&items[0], env) {
        Some(name) => function_work_out(function_bind_args(&name, items, env)?, env),
        None => unreachable!("the arm asked for a function's name first"),
    }
}

/// VLA.8, VBA `CallFunction`: a function's call, its arguments read in the
/// caller's scope ([`function_bind_args`]), then worked out
/// ([`function_work_out`]).
fn call_function(name: &str, call: &[Node], env: &Env) -> Result<Node, Refusal> {
    function_work_out(function_bind_args(name, call, env)?, env)
}

/// VLA.8, VBA `FunctionBindArgs`: a call's arguments bound to the function's
/// parameters, each read as a primitive reads its own, left to right, in the
/// caller's scope; those past the parameters the rest's values, a function
/// with no rest refusing them by its arity.
fn function_bind_args(name: &str, call: &[Node], env: &Env) -> Result<CallState, Refusal> {
    let function = match env.functions.and_then(|f| f.get(name)) {
        Some(f) => Arc::clone(f),
        None => unreachable!("a call is bound only for a function's name"),
    };
    let argc = call.len() - 1;
    let want = function.params.len();
    if function.rest_name.is_empty() {
        if argc != want {
            return Err(raise(
                "vla-function-arity",
                &[
                    ("name", &function.name),
                    ("n", &want.to_string()),
                    ("given", &argc.to_string()),
                ],
            ));
        }
    } else if argc < want {
        return Err(raise(
            "vla-function-arity-min",
            &[
                ("name", &function.name),
                ("n", &want.to_string()),
                ("given", &argc.to_string()),
            ],
        ));
    }
    let mut values = Vec::with_capacity(want);
    let mut rest = Vec::new();
    for (j, arg) in call[1..].iter().enumerate() {
        let v = resolve_listops_arg(arg, env)?;
        held_depth(&v)?;
        if j < want {
            values.push(v);
        } else {
            rest.push(v);
        }
    }
    Ok(CallState {
        function,
        values,
        rest,
    })
}

/// VLA.8, VBA `FunctionWorkOut`: a working-out. Each step is one call's body
/// in its scope ([`function_step`]): a call in a tail place replaces the call
/// under way and loops, nesting nothing; any other form is the value. A
/// working-out nested inside another is one level of
/// [`FUNCTION_DEPTH_LIMIT`], the outermost one starting the fuel and the
/// trail ([`work_out_enter`]), and every way out restores the depth and gives
/// a refusal the trail of calls it stands in ([`work_out_leave`]). One frame
/// for the whole loop, its refusals and its trail made by helpers off the
/// path, so that a level of nested calls costs the stack the least it can.
fn function_work_out(mut state: CallState, env: &Env) -> Result<Node, Refusal> {
    let local;
    let calls = match env.calls {
        Some(c) => c,
        None => {
            local = Calls::default();
            &local
        }
    };
    work_out_enter(calls, &state)?;
    let result = loop {
        if let Err(e) = work_out_fuel(calls, &state) {
            break Err(e);
        }
        match function_step(&state, env) {
            Ok(Step::Value(v)) => break Ok(v),
            Ok(Step::Tail(next)) => {
                state = next;
                work_out_retrail(calls, &state);
            }
            Err(e) => break Err(e),
        }
    };
    work_out_leave(calls, result)
}

/// VLA.8: a working-out's entry: the fuel and the trail started by the
/// outermost one, the depth held to [`FUNCTION_DEPTH_LIMIT`], then a level
/// more and the call on the trail.
#[inline(never)]
fn work_out_enter(calls: &Calls, state: &CallState) -> Result<(), Refusal> {
    if calls.depth() == 0 {
        calls.fuel.store(0, Ordering::Relaxed);
        calls.trail_written.store(false, Ordering::Relaxed);
        calls.trail().clear();
    }
    if calls.depth() >= FUNCTION_DEPTH_LIMIT {
        return Err(raise(
            "vla-function-depth",
            &[
                ("name", &state.function.name),
                ("limit", &FUNCTION_DEPTH_LIMIT.to_string()),
            ],
        ));
    }
    calls.set_depth(calls.depth() + 1);
    calls.trail().push(state.trail_entry());
    Ok(())
}

/// VLA.8: one call's fuel, refused past [`FUNCTION_FUEL`].
#[inline(never)]
fn work_out_fuel(calls: &Calls, state: &CallState) -> Result<(), Refusal> {
    let fuel = calls.fuel.fetch_add(1, Ordering::Relaxed) + 1;
    if fuel > FUNCTION_FUEL {
        return Err(raise(
            "vla-function-fuel",
            &[
                ("name", &state.function.name),
                ("limit", &FUNCTION_FUEL.to_string()),
            ],
        ));
    }
    Ok(())
}

/// VLA.8: a tail call's entry on the trail, in place of the call it replaces.
#[inline(never)]
fn work_out_retrail(calls: &Calls, state: &CallState) {
    let mut trail = calls.trail();
    trail.pop();
    trail.push(state.trail_entry());
}

/// VLA.8: a working-out's way out: its value held to the reader's depth
/// ([`held_depth`]), a refusal given the trail by the innermost call it
/// passes, then the call off the trail and a level less, on every way out.
#[inline(never)]
fn work_out_leave(calls: &Calls, result: Result<Node, Refusal>) -> Result<Node, Refusal> {
    let result = match result.and_then(|v| held_depth(&v).map(|()| v)) {
        Ok(v) => Ok(v),
        Err(mut e) => {
            if !calls.trail_written.load(Ordering::Relaxed) {
                let text = trail_text(&calls.trail());
                e.text.push_str(&text);
                calls.trail_written.store(true, Ordering::Relaxed);
            }
            Err(e)
        }
    };
    calls.trail().pop();
    calls.set_depth(calls.depth() - 1);
    result
}

/// VLA.8, VBA `FunctionStep`: one step of a working-out: the call's scope, its
/// parameters bound to their values and its rest to the rest's, and the body
/// read down its tail ([`tail_of`]).
fn function_step(state: &CallState, env: &Env) -> Result<Step, Refusal> {
    let function = &state.function;
    let mut scope = env.bare();
    scope.rest_name = &function.rest_name;
    scope.rest_items = &state.rest;
    for (param, value) in function.params.iter().zip(&state.values) {
        scope.bindings.insert(param.as_str(), value.clone());
    }
    tail_of(&function.body, &scope)
}

/// VLA.8: what a form of a function's tail is: a quote-if or a cond, whose
/// chosen form the tail reads on; a let-expand; a call of a function, by its
/// folded name; or a value.
enum Tail {
    QuoteIf,
    Cond,
    Let,
    Call(String),
    Value,
}

/// VLA.8: a form of a tail told by its head, a function's call by
/// [`function_head_name`] in the tail's scope.
#[inline(never)]
fn tail_kind(body: &Node, scope: &Env) -> Tail {
    let Node::List(lst) = body else {
        return Tail::Value;
    };
    let Some(Node::Sym(head)) = lst.items().first() else {
        return Tail::Value;
    };
    match fold(head).as_str() {
        "quote-if" => Tail::QuoteIf,
        "cond" => Tail::Cond,
        "let-expand" => Tail::Let,
        _ => match function_head_name(&lst.items()[0], scope) {
            Some(name) => Tail::Call(name),
            None => Tail::Value,
        },
    }
}

/// VLA.8: a list's items, none for an atom.
fn list_items(n: &Node) -> &[Node] {
    match n {
        Node::List(l) => l.items(),
        _ => &[],
    }
}

/// VLA.8: a body read down its tail - a quote-if's chosen branch, a cond's
/// chosen clause while its tests fold, a let-expand's one form under its
/// binding ([`tail_of_let`]) - to a call of a function, a tail call bound in
/// this scope, or to a value, read in it as a value is read
/// ([`substitute_value`], so that a loop there is the list of its results and
/// a map is a loop). A loop over the chosen forms, a let-expand alone
/// recursing, by the written nesting of lets, so that the tail costs the stack
/// one frame however deep its choices nest.
fn tail_of(body: &Node, scope: &Env) -> Result<Step, Refusal> {
    let mut body = body.clone();
    loop {
        let next = match tail_kind(&body, scope) {
            Tail::QuoteIf => quote_if_branch(list_items(&body), scope)?.clone(),
            Tail::Cond => match cond_tail_form(list_items(&body), scope)? {
                Some(form) => form,
                None => break,
            },
            Tail::Let => return tail_of_let(list_items(&body), scope),
            Tail::Call(name) => {
                return Ok(Step::Tail(function_bind_args(
                    &name,
                    list_items(&body),
                    scope,
                )?))
            }
            Tail::Value => break,
        };
        body = next;
    }
    Ok(Step::Value(substitute_value(&body, scope)?))
}

/// VLA.8: a let-expand in a function's tail: its scope, built and checked as
/// `LetExpandScope` builds it, then its one form, read on in that scope.
#[inline(never)]
fn tail_of_let(items: &[Node], scope: &Env) -> Result<Step, Refusal> {
    let bindings = let_expand_bindings(items)?;
    let folds = let_expand_folds(bindings);
    let let_scope = let_expand_scope(bindings, &folds, scope)?;
    if items.len() != 3 {
        return Err(raise(
            "vla-let-expand-value-forms",
            &[("n", &(items.len() - 2).to_string())],
        ));
    }
    tail_of(&items[2], &let_scope)
}

/// VLA.8, VBA `CondTailForm`: the chosen clause's form of a cond in a
/// function's tail while its tests fold to true or false, as
/// [`eval_cond_from`] chooses it (the same refusals in the same order), or
/// the empty begin when no clause holds; `None` for a test that waits on the
/// running program, whose cond is then substituted whole, as a template's is.
fn cond_tail_form(lst: &[Node], scope: &Env) -> Result<Option<Node>, Refusal> {
    for k in 2..=lst.len() {
        let cl = nth_list(lst, k)?.items();
        if cl.len() != 2 {
            return Err(raise(
                "vla-cond-clause-shape",
                &[("n", &cl.len().to_string())],
            ));
        }
        if let Node::Sym(s) = nth(cl, 1)? {
            if fold(s) == "else" {
                if k != lst.len() {
                    return Err(raise("vla-cond-else-not-last", &[]));
                }
                return Ok(Some(nth(cl, 2)?.clone()));
            }
        }
        let t = resolve_listops_arg(nth(cl, 1)?, scope)?;
        let tf = match &t {
            Node::Sym(s) => fold(s),
            _ => String::new(),
        };
        if tf == "true" {
            return Ok(Some(nth(cl, 2)?.clone()));
        }
        if tf != "false" {
            return Ok(None);
        }
    }
    Ok(Some(Node::list(vec![Node::sym("begin")])))
}

/// VLA.8, VBA `FunctionTrailText`: the calls a refusal stands in, outermost
/// first, as they were called: ` (computing (f 1) > (g 2))`. Four at most:
/// past four, the outermost, then `...`, then the three innermost.
fn trail_text(trail: &[(String, Vec<Node>)]) -> String {
    let n = trail.len();
    if n == 0 {
        return String::new();
    }
    let mut s = String::new();
    if n <= 4 {
        for (k, entry) in trail.iter().enumerate() {
            if k > 0 {
                s.push_str(" > ");
            }
            s.push_str(&call_text(entry));
        }
    } else {
        s.push_str(&call_text(&trail[0]));
        s.push_str(" > ...");
        for entry in &trail[n - 3..] {
            s.push_str(" > ");
            s.push_str(&call_text(entry));
        }
    }
    format!(" (computing {s})")
}

/// VLA.8, VBA `FunctionCallText`: one call of the trail, `(name value ...)`,
/// each value as the printer writes it, or `...` for one of more than 60
/// forms, its weight read and never walked (`TrailCountForms`), so that a
/// value shared into a vast tree is never written; past 60 UTF-16 units cut to
/// 57 and `...` added, never between the two halves of a surrogate pair, as
/// the reference's `Left$` counts and cuts.
fn call_text((name, args): &(String, Vec<Node>)) -> String {
    let mut s = format!("({name}");
    for v in args {
        s.push(' ');
        if v.weight().0 > 60 {
            s.push_str("...");
        } else {
            s.push_str(&write_datum(&v.form()));
        }
    }
    s.push(')');
    let units: Vec<u16> = s.encode_utf16().collect();
    if units.len() <= 60 {
        return s;
    }
    let mut cut = 57;
    if (0xD800..=0xDBFF).contains(&units[cut - 1]) {
        cut -= 1;
    }
    String::from_utf16_lossy(&units[..cut]) + "..."
}

/// VLA.8, VBA `EvalLetExpand`: `(let-expand ((name value) ...) form ...)`, a
/// value bound for its forms: the scope ([`let_expand_scope`]), then the forms
/// substituted in it, one form its own result and several one begin, as a
/// template's are. Where a value is read it holds one form ([`value_read`]).
fn eval_let_expand(items: &[Node], env: &Env) -> Result<Node, Refusal> {
    let bindings = let_expand_bindings(items)?;
    let folds = let_expand_folds(bindings);
    let scope = let_expand_scope(bindings, &folds, env)?;
    if items.len() == 3 {
        return substitute(&items[2], &scope, false);
    }
    let mut out = Vec::with_capacity(items.len() - 1);
    out.push(Node::sym("begin"));
    for f in &items[2..] {
        out.push(substitute(f, &scope, false)?);
    }
    Ok(Node::list(out))
}

/// VLA.8: a let-expand's bindings, or the refusal of its shape:
/// `((name value) ...)`, each a name and its value, and at least one form
/// after them, every binding's shape held before any name is read, as
/// `LetExpandScope` holds it.
fn let_expand_bindings(items: &[Node]) -> Result<&[Node], Refusal> {
    let shape = || raise("vla-let-expand-shape", &[]);
    if items.len() < 3 {
        return Err(shape());
    }
    let Node::List(bl) = &items[1] else {
        return Err(shape());
    };
    for b in bl.items() {
        let Node::List(pair) = b else {
            return Err(shape());
        };
        if pair.count() != 2 || !matches!(pair.items()[0], Node::Sym(_)) {
            return Err(shape());
        }
    }
    Ok(bl.items())
}

/// VLA.8: each binding's name, folded, in order, which the scope borrows.
fn let_expand_folds(bindings: &[Node]) -> Vec<String> {
    bindings
        .iter()
        .map(|b| match b {
            Node::List(pair) => fold(&symbol_text(&pair.items()[0])),
            _ => unreachable!("a binding's shape was held first"),
        })
        .collect()
}

/// VLA.8, VBA `LetExpandScope`: the scope a let-expand's forms read: the
/// template's bindings, and each name, checked as a loop's is
/// ([`each_check_name`]) and none alike another, bound in order to its value,
/// read as a primitive reads its argument under the names before it.
fn let_expand_scope<'a>(
    bindings: &'a [Node],
    folds: &'a [String],
    env: &Env<'a>,
) -> Result<Env<'a>, Refusal> {
    let mut scope = Env {
        bindings: env.bindings.clone(),
        rest_name: env.rest_name,
        rest_items: env.rest_items,
        macros: env.macros,
        functions: env.functions,
        defines: env.defines,
        calls: env.calls,
    };
    for (k, b) in bindings.iter().enumerate() {
        let Node::List(pair) = b else {
            unreachable!("a binding's shape was held first")
        };
        let pair = pair.items();
        let word = symbol_text(&pair[0]);
        each_check_name("let-expand", &word, env)?;
        if folds[..k].contains(&folds[k]) {
            return Err(raise(
                "vla-each-name-bound",
                &[("op", "let-expand"), ("word", &word)],
            ));
        }
        let value = resolve_listops_arg(&pair[1], &scope)?;
        held_depth(&value)?;
        scope.bindings.insert(folds[k].as_str(), value);
    }
    Ok(scope)
}

/// VLA.8: an argument substituted and read as a value ([`value_read`]), a
/// helper so that [`resolve_listops_arg`], on every primitive's path, holds
/// neither.
#[inline(never)]
fn substitute_read(arg: &Node, env: &Env) -> Result<Node, Refusal> {
    let v = substitute(arg, env, false)?;
    value_read(arg, v, env)
}

/// VLA.8, VBA `ValueRead`: a value read the same whoever wrote it,
/// [`resolve_listops_arg`]'s last step before a quote is unwrapped. A symbol
/// that is a define's word is its form, read as a value in an empty scope, so
/// that `(+expand side 1)` reads a define of `side`. A list a word of the
/// template held, a parameter, a loop's name or a let-expand's name, never the
/// rest's whole list, whose head names a primitive or a function, is that
/// call, computed in an empty scope, so that a call handed in reads as one
/// written in place. Any other value as it is. And a let-expand written where
/// the value is read holds one form, its value.
fn value_read(arg: &Node, v: Node, env: &Env) -> Result<Node, Refusal> {
    if let Node::List(l) = arg {
        if arg.head_is("let-expand") && l.count() > 3 {
            return Err(raise(
                "vla-let-expand-value-forms",
                &[("n", &(l.count() - 2).to_string())],
            ));
        }
    }
    match &v {
        Node::Sym(s) => {
            let form = env
                .defines
                .filter(|d| !d.is_empty())
                .and_then(|d| d.get(&fold(s)));
            match form {
                Some(form) => resolve_listops_arg(form, &env.bare()),
                None => Ok(v),
            }
        }
        Node::Str(_) => Ok(v),
        Node::List(held) => {
            let Node::Sym(word) = arg else {
                return Ok(v);
            };
            if fold(word) == env.rest_name {
                return Ok(v);
            }
            let Some(Node::Sym(h)) = held.items().first() else {
                return Ok(v);
            };
            let hf = fold(h);
            if is_listops(&hf) {
                return eval_listops_prim(&hf, held.items(), &env.bare());
            }
            if env.functions.is_some_and(|f| f.contains_key(&hf)) {
                return call_function(&hf, held.items(), &env.bare());
            }
            Ok(v)
        }
    }
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
        // VLA.8: a value read the same whoever wrote it ([`substitute_read`]).
        _ => substitute_read(arg, env)?,
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
    substitute(quote_if_branch(lst, env)?, env, false)
}

/// VLA.8, VBA `QuoteIfBranch`: quote-if's chosen branch, as written: the
/// arity, then the test resolved as a primitive's argument and held to true or
/// false, as [`eval_quote_if`] has always held it. A function's tail takes the
/// branch from here without substituting it, so that a call there loops.
fn quote_if_branch<'a>(lst: &'a [Node], env: &Env) -> Result<&'a Node, Refusal> {
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
    match tf.as_str() {
        "true" => nth(lst, 3),
        "false" => nth(lst, 4),
        _ => Err(raise(
            "vla-quote-if-test-invalid",
            &[("value", &t.atom_text().unwrap_or_default())],
        )),
    }
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
        refuse_self_text("&expand", &piece)?; // VLA.10
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
    refuse_self_text("characters-expand", &v)?; // VLA.10
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
    refuse_self_text(op, &v)?; // VLA.10
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
        // VLA.15: its first cell, the third; VLA.10: the word the walk reads,
        // the form that binds it, and a reference moved.
        let walk_operations = [
            "formula-expand",
            "each",
            "each-index",
            "define",
            "sheet-expand",
            "range-expand",
            "first-cell-expand",
            "self",
            "bind-expand",
            "offset-expand",
            // VLA.8: the value-binding form and the function-defining form.
            "let-expand",
            "defexpand",
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
        let r = expand_one(call.items(), &ex.rules["m"], &ex.bare_env(), 0).unwrap();
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

    // VLA.16: a list of names where the name stands, each bound to its part.
    #[test]
    fn a_loop_names_each_part_of_an_entry() {
        assert_eq!(
            one("(each ((a b) (quote ((1 2) (3 4)))) (f a b))"),
            "(begin (f 1 2) (f 3 4))"
        );
        assert_eq!(
            one(concat!(
                "(defmacro (cells sheet & entries) ",
                "(each ((address value) entries) (cell sheet address value))) ",
                "(cells \"S\" (\"A1\" 1) (\"C1:C3\" 0))"
            )),
            "(begin (cell \"S\" \"A1\" 1) (cell \"S\" \"C1:C3\" 0))"
        );
        assert_eq!(
            one("(each ((a) (quote ((x) (y)))) (f a))"),
            "(begin (f x) (f y))"
        );
        assert_eq!(
            one("(each ((at fill) (quote ((A1 (1 2)) (B1 ())))) (f at fill))"),
            "(begin (f A1 (1 2)) (f B1 ()))"
        );
        assert_eq!(one("(each ((a b) (quote ())) (f a b))"), "(begin)");
        assert_eq!(
            one("(each-index (i (color name) (quote ((c0 black) (c1 blue))) 0) (f i color name))"),
            "(begin (f 0 c0 black) (f 1 c1 blue))"
        );
        assert_eq!(
            one("(each ((row xs) (quote ((1 (a b)) (2 (c))))) (each ((x) (quote ((p)))) (f row x xs)))"),
            "(begin (f 1 p (a b)) (f 2 p (c)))"
        );
        assert_eq!(
            one("(+ (each ((c k) (quote ((A1 2) (B1 3)))) (* c k)))"),
            "(+ (* A1 2) (* B1 3))"
        );
    }

    #[test]
    fn a_loop_s_names_are_refused_by_name_where_they_cannot_read_an_entry() {
        for (src, id, frag) in [
            (
                "(each ((a b) (quote ((1 2) (3)))) (f a b))",
                "vla-each-entry-shape",
                "each: the names (a b) read each entry as a list of 2, a part for each name, and the entry (3) is a list of 1",
            ),
            (
                "(each ((a b) (quote ((1 2 3)))) (f a b))",
                "vla-each-entry-shape",
                "and the entry (1 2 3) is a list of 3",
            ),
            (
                "(each ((a b) (quote (x))) (f a b))",
                "vla-each-entry-shape",
                "and the entry x is no list",
            ),
            (
                "(each-index (i (a b) (quote (\"t\")) 0) (f a b))",
                "vla-each-entry-shape",
                "each-index: the names (a b) read each entry as a list of 2, a part for each name, and the entry \"t\" is no list",
            ),
            (
                "(each (() (quote ((1)))) (f))",
                "vla-each-shape",
                "each: write (each (name list) form ...)",
            ),
            (
                "(each ((a \"b\") (quote ((1 2)))) (f a))",
                "vla-each-shape",
                "each: write (each (name list) form ...)",
            ),
            (
                "(each ((a (b c)) (quote ((1 (2 3))))) (f a))",
                "vla-each-shape",
                "each: write (each (name list) form ...)",
            ),
            (
                "(each-index (i (a \"b\") (quote ((1 2))) 0) (f a))",
                "vla-each-shape",
                "each-index: write (each-index (index name list start) form ...)",
            ),
            (
                "(each ((a A) (quote ((1 2)))) (f a))",
                "vla-each-name-bound",
                "each: 'A' is already a name in this template",
            ),
            (
                "(each-index (i (x i) (quote ((1 2))) 0) (f x))",
                "vla-each-name-bound",
                "each-index: 'i' is already a name in this template",
            ),
            (
                "(each ((at car) (quote ((1 2)))) (f at))",
                "vla-each-name-reserved",
                "each 'car': 'car' is the list-head primitive and cannot be a name each binds",
            ),
            (
                "(defmacro (m sheet & xs) (each ((sheet v) xs) (f v))) (m \"S\" (1 2))",
                "vla-each-name-bound",
                "each: 'sheet' is already a name in this template",
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
        // VLA.8, decision 5, a changed meaning: a primitive reads a define's
        // word as its form, where it read the word as data.
        assert_eq!(
            one("(define stop State!B12) (eq? stop (quote stop))"),
            "(begin)\nfalse"
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

    /// VLA.9's macro of the fixture's first cases: a positional parameter
    /// and two pieces, each with a default.
    const RULE: &str =
        "(defmacro (rule range (sheet \"Screen\") (text \"=1\")) (formula sheet range text)) ";

    #[test]
    fn a_piece_takes_its_value_or_its_default_in_any_order() {
        let with = |call: &str| one(&format!("{RULE}{call}"));
        assert_eq!(with("(rule \"B2\")"), "(formula \"Screen\" \"B2\" \"=1\")");
        assert_eq!(
            with("(rule \"B2\" (sheet \"Tower\"))"),
            "(formula \"Tower\" \"B2\" \"=1\")"
        );
        assert_eq!(
            with("(rule \"B2\" (text \"=2\"))"),
            "(formula \"Screen\" \"B2\" \"=2\")"
        );
        let want = "(formula \"Tower\" \"B2\" \"=2\")";
        assert_eq!(with("(rule \"B2\" (sheet \"Tower\") (text \"=2\"))"), want);
        assert_eq!(with("(rule \"B2\" (text \"=2\") (sheet \"Tower\"))"), want);
        assert_eq!(with("(RULE \"B2\" (TEXT \"=2\") (Sheet \"Tower\"))"), want);
        assert_eq!(
            one("(defmacro (mark (at)) (cell \"S\" at 1)) (mark (at \"A1\"))"),
            "(cell \"S\" \"A1\" 1)"
        );
        assert_eq!(
            one("(defmacro (mark (at \"A1\")) (cell \"S\" at 1)) (mark)"),
            "(cell \"S\" \"A1\" 1)"
        );
    }

    #[test]
    fn a_default_is_substituted_under_the_parameters_bound_before_it() {
        let frame = "(defmacro (frame sheet (top 1) (bottom top)) (box sheet top bottom)) ";
        let with = |call: &str| one(&format!("{frame}{call}"));
        assert_eq!(with("(frame \"S\")"), "(box \"S\" 1 1)");
        assert_eq!(with("(frame \"S\" (top 3))"), "(box \"S\" 3 3)");
        assert_eq!(with("(frame \"S\" (bottom 9) (top 3))"), "(box \"S\" 3 9)");
        // A positional parameter, a primitive reading a piece, and a default
        // leaning on a default.
        assert_eq!(
            one("(defmacro (titled sheet (title sheet)) (cell sheet \"A1\" title)) (titled \"T\")"),
            "(cell \"T\" \"A1\" \"T\")"
        );
        assert_eq!(
            one("(defmacro (span (from 1) (to (+expand from 9))) (f from to)) (span (from 3))"),
            "(f 3 12)"
        );
        assert_eq!(
            one("(defmacro (chain (a 1) (b a) (c b)) (f a b c)) (chain (a 7))"),
            "(f 7 7 7)"
        );
        // A default stands where its piece's word stands and expands there,
        // so a word reads the defines the call's form sees.
        assert_eq!(
            one("(defmacro (m (at home)) (f at)) (define home \"B2\") (m)"),
            "(begin)\n(f \"B2\")"
        );
        assert_eq!(
            one("(defmacro (m (at home)) (f at)) (m) (define home \"B2\")"),
            "(f home)\n(begin)"
        );
        // A default left out of nothing is never read.
        assert_eq!(
            one("(defmacro (m (a (car \"t\"))) (f a)) (m (a 1))"),
            "(f 1)"
        );
        assert_eq!(
            expand_err("(defmacro (m (a (car \"t\"))) (f a)) (m)").id,
            "vla-car-expected-list"
        );
    }

    #[test]
    fn a_macro_records_its_pieces_as_written() {
        let mut ex = Expander::new();
        ex.collect(
            read_forms("(defmacro (Rule range (Sheet \"Screen\") (look)) \"the doc\" range)")
                .unwrap(),
        )
        .unwrap();
        let m = ex.get_macro("rule").unwrap();
        assert_eq!(m.params, ["range"]);
        assert_eq!(m.rest_name, "");
        assert_eq!(m.doc, "the doc");
        assert_eq!(m.name, "Rule");
        assert_eq!(m.pieces.len(), 2);
        assert_eq!(m.pieces[0].word, "Sheet");
        assert_eq!(m.pieces[0].default, Some(Form::string("Screen")));
        assert_eq!(m.pieces[1].word, "look");
        assert_eq!(m.pieces[1].default, None);
        assert_eq!(ex.macro_doc("RULE"), "the doc");
        // A macro with none records none.
        ex.collect(read_forms("(defmacro (plain x & more) x)").unwrap())
            .unwrap();
        assert!(ex.get_macro("plain").unwrap().pieces.is_empty());
    }

    #[test]
    fn a_malformed_declaration_is_refused_by_name() {
        for (src, id, frag) in [
            (
                "(defmacro (m ()) 1)",
                "vla-defmacro-piece-shape",
                "got '()'",
            ),
            (
                "(defmacro (m (a 1 2)) 1)",
                "vla-defmacro-piece-shape",
                "got '(a 1 2)'",
            ),
            (
                "(defmacro (m (\"a\" 1)) 1)",
                "vla-defmacro-piece-shape",
                "got '(\"a\" 1)'",
            ),
            (
                "(defmacro (m ((a) 1)) 1)",
                "vla-defmacro-piece-shape",
                "got '((a) 1)'",
            ),
            (
                "(defmacro (m (& rest)) 1)",
                "vla-defmacro-piece-shape",
                "got '(& rest)'",
            ),
            (
                "(defmacro (m (a 1) b) 1)",
                "vla-defmacro-param-after-piece",
                "the parameter 'b'",
            ),
            (
                "(defmacro (m (a 1) & b) 1)",
                "vla-defmacro-piece-with-rest",
                "defmacro 'm'",
            ),
            (
                "(defmacro (m (a 1) &) 1)",
                "vla-defmacro-piece-with-rest",
                "defmacro 'm'",
            ),
            (
                "(defmacro (m & b (a 1)) 1)",
                "vla-defmacro-rest-param-not-last",
                "must be last",
            ),
            (
                "(defmacro (m & (a 1)) 1)",
                "vla-expected-symbol-got-list",
                "got a list",
            ),
            (
                "(defmacro (m a a) 1)",
                "vla-defmacro-param-twice",
                "'a' is declared twice",
            ),
            (
                "(defmacro (m a A) 1)",
                "vla-defmacro-param-twice",
                "'A' is declared twice",
            ),
            (
                "(defmacro (m a (a 1)) 1)",
                "vla-defmacro-param-twice",
                "'a' is declared twice",
            ),
            (
                "(defmacro (m (a 1) (A 2)) 1)",
                "vla-defmacro-param-twice",
                "'A' is declared twice",
            ),
            (
                "(defmacro (m a & a) 1)",
                "vla-defmacro-param-twice",
                "'a' is declared twice",
            ),
            (
                "(defmacro (m (each 1)) 1)",
                "vla-defmacro-piece-reserved",
                "expand-time loop form",
            ),
            (
                "(defmacro (m (CAR)) 1)",
                "vla-defmacro-piece-reserved",
                "the piece 'CAR'",
            ),
            (
                "(defmacro (m (a a)) 1)",
                "vla-defmacro-piece-default-unbound",
                "names 'a'",
            ),
            (
                "(defmacro (m (a b) (b 1)) 1)",
                "vla-defmacro-piece-default-unbound",
                "names 'b'",
            ),
            (
                "(defmacro (m (a (g (B))) (b 1)) 1)",
                "vla-defmacro-piece-default-unbound",
                "names 'b'",
            ),
            (
                "(defmacro (m (a (quote b)) (b 1)) 1)",
                "vla-defmacro-piece-default-unbound",
                "names 'b'",
            ),
            (
                "(defmacro (m (a 1)))",
                "vla-defmacro-missing-template",
                "missing template",
            ),
        ] {
            let e = expand_err(src);
            assert_eq!(e.id, id, "{src}");
            assert!(e.text.contains(frag), "{src}: {}", e.text);
        }
        // A text naming a later piece's word is no name.
        assert_eq!(
            one("(defmacro (m (a \"b\") (b 1)) (f a b)) (m)"),
            "(f \"b\" 1)"
        );
    }

    #[test]
    fn a_call_s_fault_is_refused_by_name_in_one_order() {
        let both = "(defmacro (two p q (a) (b)) (f p q a b)) ";
        for (call, id, frag) in [
            // The count, before any piece is read.
            (
                "(two (a 1))",
                "vla-macro-arity-min",
                "at least 2 argument(s), got 1",
            ),
            // The positional arguments, left to right.
            (
                "(two (a 1) (b 2))",
                "vla-macro-piece-for-param",
                "the positional argument 'p'",
            ),
            (
                "(two 1 (B 2) (x 3))",
                "vla-macro-piece-for-param",
                "the positional argument 'q'",
            ),
            // The arguments after them, left to right.
            (
                "(two 1 2 (x 3) (a))",
                "vla-macro-piece-unknown",
                "'(x 3)' stands after",
            ),
            (
                "(two 1 2 (a) (x 3))",
                "vla-macro-piece-arity",
                "holds 0 form(s)",
            ),
            (
                "(two 1 2 (a 1 2))",
                "vla-macro-piece-arity",
                "holds 2 form(s)",
            ),
            (
                "(two 1 2 (a 1) (A 2))",
                "vla-macro-piece-twice",
                "the piece 'a' is given twice",
            ),
            ("(two 1 2 x)", "vla-macro-piece-unknown", "'x' stands after"),
            (
                "(two 1 2 \"x\")",
                "vla-macro-piece-unknown",
                "'\"x\"' stands after",
            ),
            (
                "(two 1 2 ())",
                "vla-macro-piece-unknown",
                "'()' stands after",
            ),
            (
                "(two 1 2 ((a) 1))",
                "vla-macro-piece-unknown",
                "'((a) 1)' stands after",
            ),
            (
                "(two 1 2 (\"a\" 1))",
                "vla-macro-piece-unknown",
                "'(\"a\" 1)' stands after",
            ),
            // Then the pieces left out, in the declaration's order.
            ("(two 1 2)", "vla-macro-piece-missing", "the piece 'a'"),
            (
                "(two 1 2 (a 1))",
                "vla-macro-piece-missing",
                "the piece 'b'",
            ),
        ] {
            let e = expand_err(&format!("{both}{call}"));
            assert_eq!(e.id, id, "{call}");
            assert!(e.text.contains(frag), "{call}: {}", e.text);
        }
        // An unknown piece's refusal names the pieces, as written.
        let e = expand_err("(defmacro (m (Sheet 1) (look 2)) 1) (M (x 3))");
        assert!(e.text.starts_with("macro 'M': "), "{}", e.text);
        assert!(
            e.text
                .ends_with("only its pieces go: Sheet, look (vla line 1)"),
            "{}",
            e.text
        );
        // A default is read only once no piece is missing.
        assert_eq!(
            expand_err("(defmacro (m (a (car \"t\")) (b)) (f a b)) (m)").id,
            "vla-macro-piece-missing"
        );
        // A macro with no piece reads as it did.
        assert_eq!(one("(defmacro (m x) (f x)) (m (x 1))"), "(f (x 1))");
        assert_eq!(
            expand_err("(defmacro (m x) x) (m 1 (x 1))").id,
            "vla-macro-arity"
        );
    }

    #[test]
    fn a_refusal_names_the_call_s_line_or_none() {
        let rule = "(defmacro (rule (sheet)) (f sheet))\n";
        let at_line = |src: &str| expand_err(&format!("{rule}{src}")).text;
        assert!(at_line("\n(rule)").ends_with(" (vla line 3)"));
        assert!(at_line("(begin\n  (rule (x 1)))").ends_with(" (vla line 3)"));
        assert!(at_line("(define b (rule))\n\nb").ends_with(" (vla line 2)"));
        // A call a template writes, or one a macro rebuilds, carries none.
        for src in [
            "(defmacro (t) (rule))\n(t)",
            "(defmacro (t) (quasiquote (rule)))\n(t)",
            "(defmacro (t & c) (cons (car (car c)) (cdr (car c))))\n(t (rule))",
            "(each (x (quote (1))) (rule))",
        ] {
            let text = at_line(src);
            assert!(text.ends_with("does not give it"), "{src}: {text}");
        }
        // A given piece's value keeps its line; a default's is made afresh.
        let outer = "(defmacro (outer (inner (rule))) inner)\n";
        let given = expand_err(&format!("{rule}{outer}(outer\n  (inner (rule)))")).text;
        assert!(given.ends_with(" (vla line 4)"), "{given}");
        let fresh = expand_err(&format!("{rule}{outer}(outer)")).text;
        assert!(fresh.ends_with("does not give it"), "{fresh}");
    }

    #[test]
    fn one_pass_binds_a_piece_and_walks_no_default() {
        let src = "(defmacro (m (a (g))) (h a)) (defmacro (g) 1) (m) (m (a (g)))";
        let (text, fired) = expand_text(src, PRELUDE, false).unwrap();
        assert_eq!(text.replace("\r\n", "\n"), "(h (g))\n(h (g))\n");
        assert_eq!(fired, 2);
        let e = expand_text("(defmacro (m (a)) a)\n(m)", PRELUDE, false).unwrap_err();
        assert_eq!(e.id, "vla-macro-piece-missing");
        assert!(e.text.ends_with(" (vla line 2)"), "{}", e.text);
    }

    #[test]
    fn a_given_piece_is_the_argument_itself_and_a_default_is_made_afresh() {
        // The reference's `Set Substitute = v` for a given piece, and a
        // default substituted as a template's forms are, at line 0 (S3.1),
        // though its macro was read at a line of its own.
        let mut ex = Expander::new();
        ex.collect(read_forms("\n\n(defmacro (m (x (g 1)) (y (g 2))) (f x y))").unwrap())
            .unwrap();
        let call = Node::of(&read_forms("\n\n\n(m (x (h 3)))").unwrap()[0]);
        let Node::List(call) = &call else { panic!() };
        let r = expand_one(call.items(), &ex.rules["m"], &ex.bare_env(), call.line).unwrap();
        let Node::List(r) = &r else { panic!() };
        let [_, x, y] = r.items() else { panic!() };
        let Node::List(piece) = &call.items()[1] else {
            panic!()
        };
        let (Node::List(x), Node::List(given)) = (x, &piece.items()[1]) else {
            panic!()
        };
        assert!(Arc::ptr_eq(&x.run, &given.run));
        assert_eq!(x.line, 4);
        let Node::List(y) = y else { panic!() };
        assert_eq!(y.line, 0);
    }

    #[test]
    fn the_rest_s_mark_and_the_joining_operator_never_meet() {
        // In a signature `&` is an atom among the parameters, the mark before
        // a rest; anywhere else `(& a b)` is a list with `&` at its head.
        assert_eq!(
            one("(defmacro (m x & more) (f x more)) (m (& \"a\" \"b\") (& \"c\" \"d\"))"),
            "(f (& \"a\" \"b\") (& \"c\" \"d\"))"
        );
        let with = |call: &str| one(&format!("{RULE}{call}"));
        assert_eq!(
            with("(rule (& \"B\" 2))"),
            "(formula \"Screen\" (& \"B\" 2) \"=1\")"
        );
        assert_eq!(
            with("(rule \"B2\" (sheet (& \"T\" 1)))"),
            "(formula (& \"T\" 1) \"B2\" \"=1\")"
        );
        assert_eq!(
            one("(defmacro (m (top 1) (label (& \"row \" top))) (f label)) (m (top 3))"),
            "(f (& \"row \" 3))"
        );
        assert_eq!(
            one("(defmacro (m a (b \"x\")) (formula-expand (& a b))) (m A1)"),
            "\"=A1&\\\"x\\\"\""
        );
        // Where a piece goes, it is no piece; and it can declare none.
        assert_eq!(
            expand_err(&format!("{RULE}(rule \"B2\" (& \"T\" 1))")).id,
            "vla-macro-piece-unknown"
        );
        assert_eq!(
            expand_err("(defmacro (m (& rest)) 1)").id,
            "vla-defmacro-piece-shape"
        );
        // A rest's word spelled as its mark, taken since the first macro, is
        // a parameter's capture, VLA.11's.
        assert_eq!(
            one("(defmacro (m & &) (& \"a\" \"b\")) (m 1 2)"),
            "(1 2 \"a\" \"b\")"
        );
    }

    /// VLA.10: a verb that binds its range's first cell.
    const BOUND_RULE: &str = "(defmacro (rule sheet range expression) (bind-expand ((self (ref sheet (first-cell-expand (ref sheet range))))) (formula sheet range (formula-expand expression)))) ";

    #[test]
    fn a_binding_is_read_by_the_walk_alone() {
        let bound = |form: &str| one(&format!("(bind-expand ((self Screen!B2)) {form})"));
        assert_eq!(bound("(formula-expand (last self))"), "\"=Screen.last!B2\"");
        assert_eq!(bound("(formula-expand self)"), "\"=Screen!B2\"");
        assert_eq!(bound("(sheet-expand self)"), "\"Screen\"");
        assert_eq!(bound("(range-expand self)"), "\"B2\"");
        assert_eq!(bound("(first-cell-expand self)"), "\"B2\"");
        assert_eq!(bound("(f SELF)"), "(f Screen!B2)");
        // A head, a text and a quote are no places the walk reads a word.
        assert_eq!(bound("(formula-expand (self 1))"), "\"=SELF(1)\"");
        assert_eq!(bound("(formula-expand \"self\")"), "\"=\\\"self\\\"\"");
        assert_eq!(bound("(quote self)"), "(quote self)");
        // A read is a link of the chain and no application.
        let (_, fired) = expand_text(
            "(bind-expand ((self Screen!B2)) (f self self self))",
            PRELUDE,
            true,
        )
        .unwrap();
        assert_eq!(fired, 0);
    }

    #[test]
    fn a_verb_binds_its_range_s_first_cell() {
        let with = |call: &str| one(&format!("{BOUND_RULE}{call}"));
        assert_eq!(
            with("(rule \"Tally\" \"B1:E10\" (+ (last self) Input!B1))"),
            "(formula \"Tally\" \"B1:E10\" \"=Tally.last!B1+Input!B1\")"
        );
        assert_eq!(
            with("(rule \"Q 1\" \"E10:B1\" (last self))"),
            "(formula \"Q 1\" \"E10:B1\" \"='Q 1.last'!B1\")"
        );
        assert_eq!(
            expand_err(&format!("{BOUND_RULE}(rule \"S\" \"A:A\" 1)")).id,
            "vla-reference-whole-line"
        );
        // The nest stays refused, and says where such a verb goes.
        let nested = expand_err(&format!(
            "{BOUND_RULE}(rule \"S\" \"B1\" (rule \"S\" \"B2\" 1))"
        ));
        assert_eq!(nested.id, "vla-formula-expand-nested");
        assert!(
            nested.text.contains("an expression macro over self"),
            "{}",
            nested.text
        );
    }

    #[test]
    fn the_stack_is_empty_after_every_refusal() {
        // The wall's one condition (the determinism wall, 2026-10-29): one
        // expander, as a caller holds one across forms, its stack read after
        // each way out, and the next form reading no stale place.
        let mut ex = Expander::new();
        let form = |text: &str| read_forms(text).unwrap().remove(0);
        for refused in [
            "(bind-expand ((self A!B2)) (formula-expand (no-such-function-name 1)))",
            "(bind-expand ((self A!B2)) (cell \"A\" \"A1\" 1) (range-expand \"x\"))",
            "(bind-expand ((self A!B2)) (bind-expand ((self C!D4)) (bind-expand ((self B7)) 1)))",
            "(bind-expand ((self A!B2)) (row-expand self))",
            "(bind-expand ((self A!B2)) (bind-expand ((self (no-sheet-here B7))) (range-expand B7)))",
            "(formula-expand (bind-expand ((self A!B2)) (formula-expand 1)))",
        ] {
            assert!(ex.expand(&form(refused), 0).is_err(), "{refused}");
            assert!(ex.self_stack.is_empty(), "{refused}");
            assert!(!ex.formula_open, "{refused}");
            assert_eq!(
                ex.expand(&form("(formula-expand (last self))"), 0).unwrap_err().id,
                "vla-self-free",
                "{refused}"
            );
        }
        // And on the way out that succeeds, at any depth.
        ex.expand(
            &form("(bind-expand ((self A!B2)) (bind-expand ((self C!D4)) 1))"),
            0,
        )
        .unwrap();
        assert!(ex.self_stack.is_empty());
    }

    #[test]
    fn siblings_and_a_recursive_hand_off_each_keep_their_binding() {
        // LISTOPS-CONFLUENCE's three, as `TestListopsConfluence` proves them:
        // two chains beside each other, in either order, and a value threaded
        // through one chain handing off into a second, each ending in self.
        let mut chains = String::new();
        for i in 0..49 {
            chains.push_str(&format!("(defmacro (step-c{i}) (step-c{}))\n", i + 1));
            chains.push_str(&format!("(defmacro (step-d{i}) (step-d{}))\n", i + 1));
        }
        chains.push_str("(defmacro (step-c49) (formula-expand (last self)))\n");
        chains.push_str("(defmacro (step-d49) (formula-expand (last self)))\n");
        chains.push_str("(defmacro (both a b) (begin a b))\n");
        let c = "(bind-expand ((self A!B2)) (step-c0))";
        let d = "(bind-expand ((self C!D4)) (step-d0))";
        assert_eq!(
            one(&format!("{chains}(both {c} {d})")),
            "(begin \"=A.last!B2\" \"=C.last!D4\")"
        );
        assert_eq!(
            one(&format!("{chains}(both {d} {c})")),
            "(begin \"=C.last!D4\" \"=A.last!B2\")"
        );
        let mut handing = String::new();
        for i in 0..24 {
            handing.push_str(&format!(
                "(defmacro (step-a{i} tag) (step-a{} tag))\n",
                i + 1
            ));
            handing.push_str(&format!(
                "(defmacro (step-b{i} tag) (step-b{} tag))\n",
                i + 1
            ));
        }
        handing.push_str("(defmacro (step-a24 tag) (step-b0 tag))\n");
        handing.push_str("(defmacro (step-b24 tag) (f tag (formula-expand (last self))))\n");
        assert_eq!(
            one(&format!(
                "{handing}(bind-expand ((self A!B2)) (step-a0 \"PAYLOAD\"))"
            )),
            "(f \"PAYLOAD\" \"=A.last!B2\")"
        );
    }

    #[test]
    fn a_free_word_is_refused_where_it_would_be_read_and_nowhere_else() {
        for read in [
            "(formula-expand self)",
            "(formula-expand (+ self 1))",
            "(formula-expand (last self))",
            "(formula-expand (lambda (self) (+ self 1)))",
            "(formula-expand (quote-if (null? (quote ())) self 1))",
            "(bind-expand ((self A!B2)) (formula-expand (quote-if (null? (quote ())) self 1)))",
            "(sheet-expand self)",
            "(range-expand (last self))",
            "(first-cell-expand self)",
            "(offset-expand self 1 0)",
            "(bind-expand ((self self)) 1)",
        ] {
            assert_eq!(expand_err(read).id, "vla-self-free", "{read}");
        }
        assert_eq!(one("(set! self 1)"), "(set! self 1)");
        assert_eq!(one("(sub m () (dim self))"), "(sub m () (dim self))");
        assert_eq!(one("(cell \"S\" \"A1\" self)"), "(cell \"S\" \"A1\" self)");
        assert_eq!(one("(formula-expand (ref \"self\"))"), "\"=self\"");
        assert_eq!(one("(formula-expand (quote (1)))"), "\"={1}\"");
        for primitive in [
            "(&expand self)",
            "(characters-expand self)",
            "(row-expand self)",
            "(column-expand self)",
            "(symbol self)",
        ] {
            assert_eq!(
                expand_err(primitive).id,
                "vla-self-in-a-primitive",
                "{primitive}"
            );
        }
        assert_eq!(one("(car (quote (self)))"), "self");
        assert_eq!(one("(list? self)"), "false");
        assert_eq!(one("(&expand \"self\")"), "\"self\"");
    }

    #[test]
    fn a_word_means_one_thing_and_a_parameter_hides_it() {
        for (form, id) in [
            ("(defmacro (self) 1)", "vla-defmacro-reserved-name"),
            ("(defmacro (bind-expand) 1)", "vla-defmacro-reserved-name"),
            ("(defmacro (offset-expand) 1)", "vla-defmacro-reserved-name"),
            ("(defmacro (m (self 1)) 1)", "vla-defmacro-piece-reserved"),
            ("(each (self (quote (1))) 1)", "vla-each-name-reserved"),
            ("(define self A!B1)", "vla-define-name-reserved"),
            ("(define bind-expand A!B1)", "vla-define-name-reserved"),
        ] {
            assert_eq!(expand_err(form).id, id, "{form}");
        }
        assert_eq!(
            one("(defmacro (m self) (f self)) (bind-expand ((self A!B2)) (m 1))"),
            "(f 1)"
        );
        assert_eq!(one("(defmacro (m & self) (f self)) (m 1 2)"), "(f 1 2)");
    }

    #[test]
    fn a_reference_is_moved_as_excel_s_offset_moves_it() {
        let moved = |form: &str| one(&format!("(formula-expand {form})"));
        assert_eq!(
            moved("(offset-expand Screen!B2 -1 -1 3 3)"),
            "\"=Screen!A1:C3\""
        );
        assert_eq!(
            moved("(offset-expand Tally!B2:E10 1 1)"),
            "\"=Tally!C3:F11\""
        );
        assert_eq!(
            moved("(offset-expand Tally!E10:B2 0 0 2 2)"),
            "\"=Tally!B2:C3\""
        );
        assert_eq!(moved("(offset-expand Tally!$B2 1 1)"), "\"=Tally!$C3\"");
        assert_eq!(
            moved("(offset-expand Tally!$B$2 0 0 2 2)"),
            "\"=Tally!$B$2:$C$3\""
        );
        assert_eq!(moved("(offset-expand B2 1 0)"), "\"=B3\"");
        assert_eq!(moved("(offset-expand Tally!B2 0 0 1 1)"), "\"=Tally!B2\"");
        assert_eq!(
            moved("(last (offset-expand Tally!B2 1 0))"),
            "\"=Tally.last!B3\""
        );
        assert_eq!(
            moved("(offset-expand (last Tally!B2) 1 0)"),
            "\"=Tally.last!B3\""
        );
        assert_eq!(
            one("(offset-expand (ref \"Q 1\" \"B2\") 1 0)"),
            "(ref \"Q 1\" \"B3\")"
        );
        for (form, id) in [
            ("(offset-expand)", "vla-offset-expand-arity"),
            ("(offset-expand A!B2 1 1 1)", "vla-offset-expand-arity"),
            ("(offset-expand A!B2 -2 0)", "vla-address-expand-row"),
            ("(offset-expand A!B2 0 -2)", "vla-address-expand-column"),
            (
                "(offset-expand A!A1048576 0 0 4294967296 1)",
                "vla-address-expand-row",
            ),
            (
                "(offset-expand A!B2 99999999999999999999999 0)",
                "vla-address-expand-row",
            ),
            ("(offset-expand A!A:A 0 1)", "vla-reference-whole-line"),
            ("(offset-expand A!B2 0 0 0 1)", "vla-offset-expand-size"),
            ("(offset-expand A!B2 1.5 0)", "vla-arith-expand-not-whole"),
            (
                "(offset-expand A!B2 \"1\" 0)",
                "vla-arith-expand-not-number-value",
            ),
            (
                "(offset-expand A!B2 (f) 0)",
                "vla-arith-expand-not-number-list",
            ),
        ] {
            assert_eq!(expand_err(form).id, id, "{form}");
        }
    }

    /// One frame of the stepper: `VlaExpandStepText` under a budget of `spent`
    /// applications, the forms of the source written out a line each.
    fn frame(src: &str, spent: u64) -> String {
        let text = format!("{PRELUDE}\r\n{src}");
        let forms = parse_all(&tokenize(&text), count_lf(PRELUDE) as u32 + 1).unwrap();
        let mut ex = Expander::new();
        let body = ex.collect(forms).unwrap();
        ex.budget_on = true;
        ex.expand_budget = spent;
        body.iter()
            .map(|f| write_datum(&ex.expand(f, 0).unwrap()))
            .collect::<Vec<_>>()
            .join("\n")
    }

    #[test]
    fn the_walk_s_forms_stand_at_a_spent_budget() {
        // Found while scoping VLA.10: each of these refused its frame, or
        // wrote a macro's call as a function's, since VLA.4.
        assert_eq!(
            frame("(defmacro (my-m) (+ 1 2)) (formula-expand (my-m))", 0),
            "(formula-expand (my-m))"
        );
        assert_eq!(
            frame("(defmacro (m) Sheet1!B2) (range-expand (m))", 0),
            "(range-expand (m))"
        );
        assert_eq!(
            frame("(defmacro (m) Sheet1!B2) (first-cell-expand (m))", 0),
            "(first-cell-expand (m))"
        );
        assert_eq!(
            frame("(defmacro (m) 1) (offset-expand A!B2 (m) 0)", 0),
            "(offset-expand A!B2 (m) 0)"
        );
        assert_eq!(
            frame(
                "(defmacro (m) A!B2) (bind-expand ((self (m))) (formula-expand (last self)))",
                0
            ),
            "(bind-expand ((self (m))) (formula-expand (last self)))"
        );
        let verb = format!("{BOUND_RULE}(defmacro (running e) (+ (last self) e)) (rule \"Tally\" \"B1\" (running Input!B1))");
        assert_eq!(
            frame(&verb, 0),
            "(rule \"Tally\" \"B1\" (running Input!B1))"
        );
        assert_eq!(
            frame(&verb, 1),
            "(bind-expand ((self (ref \"Tally\" (first-cell-expand (ref \"Tally\" \"B1\"))))) (formula \"Tally\" \"B1\" (formula-expand (running Input!B1))))"
        );
        assert_eq!(
            expand(&verb).trim_end(),
            "(formula \"Tally\" \"B1\" \"=Tally.last!B1+Input!B1\")"
        );
        // A binding's read inside a macro left unfired.
        assert_eq!(
            frame("(defmacro (a) 1) (defmacro (m x) x) (bind-expand ((self A!B2)) (formula-expand (+ (a) (m self))))", 1),
            "(formula-expand (+ 1 (m A!B2)))"
        );
        // And a binding reached after the last application stands whole.
        assert_eq!(
            frame("(defmacro (m x) x) (defmacro (n) (bind-expand ((self A!B2)) (formula-expand (m self)))) (n)", 1),
            "(bind-expand ((self A!B2)) (formula-expand (m self)))"
        );
    }

    #[test]
    fn one_pass_walks_a_binding_as_a_list_and_replaces_no_word() {
        let (text, fired) = expand_text(
            "(defmacro (m) (last self)) (bind-expand ((self (m))) (formula-expand (m)))",
            PRELUDE,
            false,
        )
        .unwrap();
        assert_eq!(fired, 2);
        assert_eq!(
            text.trim_end(),
            "(bind-expand ((self (last self))) (formula-expand (last self)))"
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
