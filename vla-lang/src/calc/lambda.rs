//! `LET` and `LAMBDA` (KERNEL.29, 2026-11-15): the names they bind resolved
//! in the parse, and a formula's calls written into it at a shape's
//! placement, so that a function written in formulas costs what the formula
//! it stands for costs.
//!
//! **The names.** A name a `LET` binds or a `LAMBDA` takes is resolved once
//! the parse has read the whole text ([`resolve`]): a [`Local`] where it is
//! read, its binder's slot its own within the formula, and a call of such a
//! name, `f(1)`, an [`Apply`] of it. A name no binder in scope holds stays
//! what it was, a defined name or a function's call. A name stands for its
//! expression, substituted where it is read (the owner's decision 2,
//! 2026-11-15): substituting and binding a value compute the same value but
//! where a binding nothing reaches is one the language cannot compute, and
//! where a legacy formula's binding reads a range where one value is wanted
//! and is used inside an array parameter. A malformed binding is `#VALUE!`:
//! a `LET` of fewer than three arguments, of an even count, of more than
//! [`LET_PAIRS`] pairs, or binding something no name; a `LAMBDA` of no body,
//! of more than [`LAMBDA_PARAMS`] parameters, of a parameter no name, or of a
//! name twice.
//!
//! **The inlining** ([`inline`]), at a shape's placement and never at a step
//! (Alonzo's `AD-7`): every `LET` substituted, every application of a
//! `LAMBDA` and every call of a named function written in, its parameters
//! replaced by its arguments, each copy's own binders given fresh slots so
//! that no substitution captures a name, and a named function's references
//! made absolute, since a name's references never move with the cell that
//! reads it (`Ctx::named`'s rule). What is left is the formula written out,
//! which a run and a program compute as they compute any formula and whose
//! plan sees every cell it reads. What is not written in is kept a call,
//! which `Ctx` computes with its arguments' values (the owner's decision 1):
//! a function applied inside its own application, a named function some
//! cycle of named functions holds, and a named function some sheet scopes,
//! whose meaning a shape of cells on several sheets cannot fix at its
//! placement. A formula whose inlining would write more than [`NODES`] nodes,
//! or nest deeper than a formula may, is kept as written, which `Ctx`
//! computes by the same reading, each binding once a reading; one that would
//! nest deeper than a formula may even so is refused as such a formula is.
//!
//! **Named functions** ([`Functions`]): each defined name whose text is a
//! `LAMBDA`, read once a grid, and the cycles among them, found by
//! `graph::tarjan` over the calls their bodies make by name.

use std::rc::Rc;

use crate::grid::Model;
use crate::intrinsics::fold;
use crate::refers::Shape as Part;
use crate::sheet::column_letters;

use super::eval::Library;
use super::formula::{self, Apply, Expr, Lambda, Let, Local, RangeExpr, RefExpr, NESTING_LIMIT};
use super::graph::{split_scope, tarjan};
use super::value::ErrorKind;

/// The most pairs one `LET` binds, Excel's ("supports up to 126").
pub const LET_PAIRS: usize = 126;

/// The most parameters one `LAMBDA` takes, Excel's ("up to 253
/// parameters"); past it, `#VALUE!`.
pub const LAMBDA_PARAMS: usize = 253;

/// The most nodes a formula's inlining writes, the copies it makes as it
/// substitutes among them (KERNEL.29): enough for Excel's longest `LET`, 126
/// names each its predecessor and one, whose substitution copies about 16,000
/// nodes to write 250, so that a formula whose functions would write more, a
/// binding doubled at every step among them, is computed by `Ctx` as written,
/// each binding once a reading. A few megabytes at the most, made and dropped
/// at a shape's placement.
pub const NODES: usize = 1 << 16;

/// How deep the inlining itself recurses before it keeps a formula as
/// written: a chain of named functions each calling the next is as deep as
/// it is long, and the inlining recurses once a level of what it writes.
const RECURSION: u32 = 4 * NESTING_LIMIT;

/// The first slot the inlining gives a copy's binder: above every slot a
/// parse holds, a formula's or a named function's, each numbered from 0 by
/// its own resolution, so that a fresh slot never meets an original one.
const FRESH: u32 = 1 << 31;

/// A name as a binding keys it: the file's `_xlpm.` dropped, as the reader
/// drops it, and folded as a defined name is (`intrinsics::fold`).
fn key(name: &str) -> String {
    match name.get(..6) {
        Some(p) if p.eq_ignore_ascii_case("_xlpm.") => fold(&name[6..]),
        _ => fold(name),
    }
}

// ---- the names -----------------------------------------------------------

/// A parse with the names its `LET`s and `LAMBDA`s bind resolved: each a
/// [`Local`] where it is read, each call of one an [`Apply`], each binder's
/// slot its own within the formula.
pub(crate) fn resolve(e: Expr) -> Expr {
    Resolver {
        scope: Vec::new(),
        depths: Vec::new(),
        next: 0,
    }
    .expr(e)
}

struct Resolver {
    /// The names in scope, the innermost last: each keyed, and its slot.
    scope: Vec<(String, u32)>,
    /// How deep each `LET`'s value in scope nests ([`kept_depth`]), by its
    /// slot, so that a function's body reading one counts it as its own.
    depths: Vec<(u32, u32)>,
    next: u32,
}

impl Resolver {
    fn bound(&self, name: &str) -> Option<u32> {
        let k = key(name);
        self.scope
            .iter()
            .rev()
            .find(|(n, _)| *n == k)
            .map(|&(_, slot)| slot)
    }

    fn fresh(&mut self) -> u32 {
        let slot = self.next;
        self.next += 1;
        slot
    }

    fn all(&mut self, xs: Vec<Expr>) -> Vec<Expr> {
        xs.into_iter().map(|x| self.expr(x)).collect()
    }

    fn expr(&mut self, e: Expr) -> Expr {
        match e {
            Expr::Call(name, args) if name == "LET" => self.let_(args),
            Expr::Call(name, args) if name == "LAMBDA" => self.lambda(args),
            Expr::Call(name, args) => {
                let args = self.all(args);
                match self.bound(&name) {
                    // A name a binding holds, called: the binding's value
                    // called, whatever function of the library it spells.
                    Some(slot) => Expr::Apply(Box::new(Apply {
                        callee: Expr::Local(Local { slot, name }),
                        args,
                    })),
                    None => Expr::Call(name, args),
                }
            }
            Expr::Name { name, sheet: None } => match self.bound(&name) {
                Some(slot) => Expr::Local(Local { slot, name }),
                None => Expr::Name { name, sheet: None },
            },
            Expr::Apply(a) => {
                let Apply { callee, args } = *a;
                let callee = self.expr(callee);
                let args = self.all(args);
                Expr::Apply(Box::new(Apply { callee, args }))
            }
            Expr::Neg(x) => Expr::Neg(Box::new(self.expr(*x))),
            Expr::Pos(x) => Expr::Pos(Box::new(self.expr(*x))),
            Expr::Percent(x) => Expr::Percent(Box::new(self.expr(*x))),
            Expr::Binary(..) => self.chain(e),
            // A range's ends resolved and its reach found again, since an end
            // that is a name bound is no defined name (KERNEL.8, slice f).
            Expr::Range(r) => {
                let RangeExpr { left, right, .. } = *r;
                let left = self.expr(left);
                let right = self.expr(right);
                Expr::Range(Box::new(RangeExpr::new(left, right)))
            }
            other => other,
        }
    }

    /// A chain of operators, its left spine round a loop, as every walk of a
    /// chain goes (VLA.1).
    fn chain(&mut self, e: Expr) -> Expr {
        let mut links = Vec::new();
        let mut left = e;
        while let Expr::Binary(op, l, r) = left {
            links.push((op, *r));
            left = *l;
        }
        let mut acc = self.expr(left);
        for (op, r) in links.into_iter().rev() {
            let r = self.expr(r);
            acc = Expr::Binary(op, Box::new(acc), Box::new(r));
        }
        acc
    }

    /// `LET(name, value, ..., body)`: each name in scope for the values after
    /// it and for the body.
    fn let_(&mut self, args: Vec<Expr>) -> Expr {
        let n = args.len();
        if n < 3 || n % 2 == 0 || (n - 1) / 2 > LET_PAIRS {
            return Expr::Error(ErrorKind::Value);
        }
        let mark = self.scope.len();
        let deep = self.depths.len();
        let mut binds = Vec::with_capacity((n - 1) / 2);
        let mut it = args.into_iter();
        for _ in 0..(n - 1) / 2 {
            let (Some(name), Some(value)) = (it.next(), it.next()) else {
                break;
            };
            let Expr::Name { name, sheet: None } = name else {
                self.scope.truncate(mark);
                self.depths.truncate(deep);
                return Expr::Error(ErrorKind::Value);
            };
            let value = self.expr(value);
            let slot = self.fresh();
            let depth = nested(&value, &mut self.depths.clone());
            self.depths.push((slot, depth));
            self.scope.push((key(&name), slot));
            binds.push((Local { slot, name }, Rc::new(value)));
        }
        let body = match it.next() {
            Some(body) => self.expr(body),
            None => Expr::Error(ErrorKind::Value),
        };
        self.scope.truncate(mark);
        self.depths.truncate(deep);
        Expr::Let(Box::new(Let { binds, body }))
    }

    /// `LAMBDA(parameter, ..., body)`: the parameters in scope for the body.
    fn lambda(&mut self, args: Vec<Expr>) -> Expr {
        let n = args.len();
        if n == 0 || n - 1 > LAMBDA_PARAMS {
            return Expr::Error(ErrorKind::Value);
        }
        let mark = self.scope.len();
        let mut params = Vec::with_capacity(n - 1);
        let mut it = args.into_iter();
        for _ in 0..n - 1 {
            let Some(Expr::Name { name, sheet: None }) = it.next() else {
                self.scope.truncate(mark);
                return Expr::Error(ErrorKind::Value);
            };
            let k = key(&name);
            if self.scope[mark..].iter().any(|(n, _)| *n == k) {
                self.scope.truncate(mark);
                return Expr::Error(ErrorKind::Value);
            }
            let slot = self.fresh();
            self.scope.push((k, slot));
            params.push(Local { slot, name });
        }
        let body = match it.next() {
            Some(body) => self.expr(body),
            None => Expr::Error(ErrorKind::Value),
        };
        self.scope.truncate(mark);
        // Its body's nesting, a `LET`'s value it reads from around it counted
        // as its own, since it may be computed first inside the body.
        let depth = nested(&body, &mut self.depths.clone());
        Expr::Lambda(Rc::new(Lambda {
            params,
            body,
            depth,
        }))
    }
}

// ---- how deep an evaluation nests ------------------------------------------

/// How deep `Ctx` nests evaluating a parse as written (KERNEL.29), a level
/// for every call's arguments, sign, percent and range's end as `VLA.1`
/// counts a formula's, and a level more for every binding a name reads and
/// every binding's body, which `Ctx` enters as it enters a call's arguments.
/// A function's value is no level here: its body is evaluated where it is
/// called, which a call counts on its own ([`Lambda::depth`]).
pub(crate) fn kept_depth(e: &Expr) -> u32 {
    let mut bound: Vec<(u32, u32)> = Vec::new();
    nested(e, &mut bound)
}

fn nested(e: &Expr, bound: &mut Vec<(u32, u32)>) -> u32 {
    match e {
        Expr::Local(l) => bound
            .iter()
            .rev()
            .find(|(slot, _)| *slot == l.slot)
            .map_or(0, |&(_, d)| d + 1),
        Expr::Let(l) => {
            let mark = bound.len();
            for (local, value) in &l.binds {
                let d = nested(value, bound);
                bound.push((local.slot, d));
            }
            let d = nested(&l.body, bound) + 1;
            bound.truncate(mark);
            d
        }
        Expr::Lambda(_) => 0,
        Expr::Apply(a) => {
            let callee = nested(&a.callee, bound);
            let args = a.args.iter().map(|x| nested(x, bound)).max().unwrap_or(0);
            callee.max(args) + 1
        }
        Expr::Call(_, args) => args.iter().map(|x| nested(x, bound)).max().unwrap_or(0) + 1,
        Expr::Neg(x) | Expr::Pos(x) | Expr::Percent(x) => nested(x, bound) + 1,
        Expr::Binary(..) => {
            let mut most = 0;
            let mut left = e;
            while let Expr::Binary(_, l, r) = left {
                most = most.max(nested(r, bound));
                left = l;
            }
            most.max(nested(left, bound))
        }
        Expr::Range(r) => nested(&r.left, bound).max(nested(&r.right, bound)) + 1,
        _ => 0,
    }
}

// ---- named functions -------------------------------------------------------

/// The named functions a grid defines (KERNEL.29): each defined name whose
/// text is a `LAMBDA`, its references made absolute, and whether a cycle of
/// named functions holds it, by the calls the bodies make by name; and the
/// keys of its other names, a call of one its value called (`RATEN(1)` is
/// `#VALUE!`, as Excel's save says, `LET(f,3,f(1))`'s answer).
#[derive(Clone, Debug, Default)]
pub(crate) struct Functions {
    entries: Vec<Entry>,
    others: Vec<String>,
}

#[derive(Clone, Debug)]
struct Entry {
    key: String,
    scope: Option<usize>,
    lambda: Rc<Lambda>,
    cyclic: bool,
}

impl Functions {
    /// A grid's named functions, a workbook's or an engine's machine's.
    pub(crate) fn of_model<M: Model + ?Sized>(model: &M) -> Functions {
        let mut entries: Vec<Entry> = Vec::new();
        let mut others: Vec<String> = Vec::new();
        for (printed, text) in model.defined_names() {
            let text = crate::sheet::strip_future_prefixes(text);
            let (scope, bare) = split_scope(model, printed);
            let Ok(Expr::Lambda(l)) = formula::parse(&text) else {
                let k = fold(bare);
                if !others.contains(&k) {
                    others.push(k);
                }
                continue;
            };
            entries.push(Entry {
                key: fold(bare),
                scope,
                lambda: Rc::new(absolute_lambda(&l)),
                cyclic: false,
            });
        }
        // The functions of cells (KERNEL.29's third slice), the workbook's
        // each, as a `name` row's `LAMBDA` is.
        for f in model.functions() {
            entries.push(Entry {
                key: fold(&f.name),
                scope: None,
                lambda: Rc::new(of_cells(f)),
                cyclic: false,
            });
        }
        if entries.is_empty() {
            return Functions { entries, others };
        }
        // An edge from a function to every function of each name its body
        // calls, in any scope, so that a cycle is found wherever the calls
        // resolve; a call of a name of the library is the library's.
        let adj: Vec<Vec<usize>> = entries
            .iter()
            .map(|e| {
                let mut to = Vec::new();
                e.lambda.body.each_call(&mut |name| {
                    let k = fold(name);
                    for (j, other) in entries.iter().enumerate() {
                        if other.key == k && !to.contains(&j) {
                            to.push(j);
                        }
                    }
                });
                to
            })
            .collect();
        for component in tarjan(&adj) {
            let cyclic = component.len() > 1 || adj[component[0]].contains(&component[0]);
            if cyclic {
                for i in component {
                    entries[i].cyclic = true;
                }
            }
        }
        Functions { entries, others }
    }

    /// Whether the grid defines no name a call can call: no named function
    /// and no other name.
    pub(crate) fn is_empty(&self) -> bool {
        self.entries.is_empty() && self.others.is_empty()
    }

    /// Whether a call by this name calls a defined name whose text is no
    /// `LAMBDA` and no named function's, in any scope: its value called.
    pub(crate) fn calls_value(&self, name: &str) -> bool {
        let k = fold(name);
        !self.holds(name) && self.others.contains(&k)
    }

    /// The named function a name calls from a sheet, as `Names::resolve`
    /// resolves a name: the one the sheet scopes, else the workbook's.
    pub(crate) fn resolve(&self, name: &str, home: usize) -> Option<Rc<Lambda>> {
        let k = fold(name);
        let of = |scope: Option<usize>| {
            self.entries
                .iter()
                .find(|e| e.key == k && e.scope == scope)
                .map(|e| Rc::clone(&e.lambda))
        };
        of(Some(home)).or_else(|| of(None))
    }

    /// Whether a call by this name calls a named function, in any scope.
    pub(crate) fn holds(&self, name: &str) -> bool {
        let k = fold(name);
        self.entries.iter().any(|e| e.key == k)
    }

    /// The function a call by this name writes in wherever it stands: one no
    /// sheet scopes, whose meaning is the same at every cell, and no cycle
    /// holds.
    fn inlined(&self, name: &str) -> Option<&Rc<Lambda>> {
        let k = fold(name);
        let mut found = None;
        for e in self.entries.iter().filter(|e| e.key == k) {
            if e.scope.is_some() || e.cyclic {
                return None;
            }
            found = Some(&e.lambda);
        }
        found
    }
}

/// A function of cells' `LAMBDA` as a call writes it in (KERNEL.29's third
/// slice): its compiled text's, its references made absolute; where its body
/// does not compile, one taking as many parameters and answering `#CALC!`,
/// which every call of it answers.
fn of_cells(f: &super::room::Compiled) -> Lambda {
    if let Ok(text) = &f.lambda {
        if let Ok(Expr::Lambda(l)) = formula::parse(text) {
            return absolute_lambda(&l);
        }
    }
    let params: String = (1..=f.params).map(|k| format!("_{k},")).collect();
    match formula::parse(&format!("=LAMBDA({params}0)")) {
        Ok(Expr::Lambda(l)) => Lambda {
            params: l.params.clone(),
            body: Expr::Error(ErrorKind::Calc),
            depth: l.depth,
        },
        _ => Lambda {
            params: Vec::new(),
            body: Expr::Error(ErrorKind::Calc),
            depth: 0,
        },
    }
}

/// A named function's `LAMBDA` with every reference its body writes made
/// absolute, since a name's references do not move with the cell that reads
/// it, whatever the calling formula's offset.
fn absolute_lambda(l: &Lambda) -> Lambda {
    Lambda {
        params: l.params.clone(),
        body: absolute(&l.body),
        depth: l.depth,
    }
}

fn absolute(e: &Expr) -> Expr {
    match e {
        Expr::Ref(r) => Expr::Ref(absolute_ref(r)),
        Expr::Range(r) => Expr::Range(Box::new(RangeExpr::new(
            absolute(&r.left),
            absolute(&r.right),
        ))),
        Expr::Call(name, args) => Expr::Call(name.clone(), args.iter().map(absolute).collect()),
        Expr::Neg(x) => Expr::Neg(Box::new(absolute(x))),
        Expr::Pos(x) => Expr::Pos(Box::new(absolute(x))),
        Expr::Percent(x) => Expr::Percent(Box::new(absolute(x))),
        Expr::Binary(..) => {
            let mut links = Vec::new();
            let mut left = e;
            while let Expr::Binary(op, l, r) = left {
                links.push((*op, absolute(r)));
                left = l;
            }
            let mut acc = absolute(left);
            for (op, r) in links.into_iter().rev() {
                acc = Expr::Binary(op, Box::new(acc), Box::new(r));
            }
            acc
        }
        Expr::Let(l) => Expr::Let(Box::new(Let {
            binds: l
                .binds
                .iter()
                .map(|(local, v)| (local.clone(), Rc::new(absolute(v))))
                .collect(),
            body: absolute(&l.body),
        })),
        Expr::Lambda(l) => Expr::Lambda(Rc::new(absolute_lambda(l))),
        Expr::Apply(a) => Expr::Apply(Box::new(Apply {
            callee: absolute(&a.callee),
            args: a.args.iter().map(absolute).collect(),
        })),
        other => other.clone(),
    }
}

/// A reference held by a `$` at every corner, its text spelled so.
fn absolute_ref(r: &RefExpr) -> RefExpr {
    let mut a = r.clone();
    let c = &mut a.corners;
    c.row_abs1 = true;
    c.col_abs1 = true;
    c.row_abs2 = true;
    c.col_abs2 = true;
    let cell = |row: u32, col: u32| format!("${}${row}", column_letters(col));
    let text = match r.shape {
        Part::Cell => cell(r.top, r.left),
        Part::Range => format!("{}:{}", cell(r.top, r.left), cell(r.bottom, r.right)),
        Part::Column => format!("${}:${}", column_letters(r.left), column_letters(r.right)),
        Part::Row => format!("${}:${}", r.top, r.bottom),
        Part::None => return a,
    };
    a.written = match (&r.sheet, &r.last_sheet) {
        (Some(first), Some(last)) => format!(
            "{}:{}!{text}",
            crate::refers::quote_sheet(first),
            crate::refers::quote_sheet(last)
        ),
        (Some(sheet), None) => format!("{}!{text}", crate::rows::sheet_prefix(sheet)),
        _ => text,
    };
    a
}

// ---- the inlining ----------------------------------------------------------

/// What the inlining makes of a parse.
#[derive(Debug, PartialEq)]
pub(crate) enum Inlined {
    /// Nothing to write in, or a formula kept as written: the parse as it is.
    Same,
    /// The parse with its bindings substituted and its calls written in.
    To(Expr),
    /// A formula that nests deeper than a formula may, written in or not:
    /// refused as such a formula is (`formula::TOO_DEEP`).
    TooDeep,
}

/// A parse's bindings substituted and its calls written in, at a shape's
/// placement (KERNEL.29). A parse holding no binding, function or call of a
/// named function is the same.
pub(crate) fn inline(e: &Expr, functions: &Functions, library: &Library) -> Inlined {
    if !e.has_lambda() && !calls_named(e, functions, library) {
        return Inlined::Same;
    }
    let mut i = Inliner {
        functions,
        library,
        next: FRESH,
        nodes: 0,
        level: 0,
        env: Vec::new(),
        applying: Vec::new(),
    };
    match i.expr(e) {
        Ok(n) if n.depth <= NESTING_LIMIT => Inlined::To(n.e),
        // Written in, too deep, or too large: kept as written where that
        // nests no deeper than a formula may.
        _ if kept_depth(e) <= NESTING_LIMIT => Inlined::Same,
        _ => Inlined::TooDeep,
    }
}

/// The functions a parse calls by name, upper-cased, in the order they are
/// written: [`formula::calls`] over a text that parses, a name a binding
/// holds and a named function the inlining kept a call now an [`Apply`], and
/// so none of them.
pub(crate) fn calls_of(e: &Expr) -> Vec<String> {
    let mut out = Vec::new();
    e.each_call(&mut |name| out.push(name.to_string()));
    out
}

/// A formula's text parsed and inlined, as a shape is placed and as the door
/// evaluates it (KERNEL.29): the construct it stops at, `formula::TOO_DEEP`
/// for one too deep written in or not.
pub(crate) fn parsed(text: &str, functions: &Functions, library: &Library) -> Result<Expr, String> {
    let e = formula::parse(text)?;
    match inline(&e, functions, library) {
        Inlined::Same => Ok(e),
        Inlined::To(inlined) => Ok(inlined),
        Inlined::TooDeep => Err(formula::TOO_DEEP.to_string()),
    }
}

/// Whether a parse calls or reads a named function the library does not
/// hold, or calls another defined name.
fn calls_named(e: &Expr, functions: &Functions, library: &Library) -> bool {
    if functions.is_empty() {
        return false;
    }
    let mut found = false;
    e.each_call(&mut |name| {
        found |= !library.has(name) && (functions.holds(name) || functions.calls_value(name));
    });
    e.each_read(&mut |r| {
        if let formula::Read::Name { name, sheet: None } = r {
            found |= functions.holds(name);
        }
    });
    found
}

/// An expression the inlining wrote, how deep it nests as `VLA.1` counts a
/// formula's, and how many nodes it holds.
#[derive(Clone, Debug)]
struct Node {
    e: Expr,
    depth: u32,
    size: usize,
}

impl Node {
    fn leaf(e: Expr) -> Node {
        Node {
            e,
            depth: 0,
            size: 1,
        }
    }
}

/// Why an inlining stopped: it wrote too much or recursed too deep.
struct Over;

struct Inliner<'a> {
    functions: &'a Functions,
    library: &'a Library,
    /// The next fresh slot, above every slot the parse holds.
    next: u32,
    /// The nodes written so far.
    nodes: usize,
    /// How deep the inlining recurses where it stands.
    level: u32,
    /// The substitutions in force, the innermost last, by slot.
    env: Vec<(u32, Node)>,
    /// The functions being applied, the outermost first: a function applied
    /// inside its own application is kept a call.
    applying: Vec<*const Lambda>,
}

impl Inliner<'_> {
    fn grow(&mut self, n: usize) -> Result<(), Over> {
        self.nodes += n;
        if self.nodes > NODES {
            Err(Over)
        } else {
            Ok(())
        }
    }

    fn fresh(&mut self, l: &Local) -> Local {
        let slot = self.next;
        self.next += 1;
        Local {
            slot,
            name: l.name.clone(),
        }
    }

    fn all(&mut self, xs: &[Expr]) -> Result<Vec<Node>, Over> {
        xs.iter().map(|x| self.expr(x)).collect()
    }

    /// A node made of what was written for its operands, a level above the
    /// deepest of them.
    fn over(e: Expr, operands: &[&Node]) -> Node {
        Node {
            e,
            depth: operands.iter().map(|n| n.depth).max().unwrap_or(0) + 1,
            size: operands.iter().map(|n| n.size).sum::<usize>() + 1,
        }
    }

    fn expr(&mut self, e: &Expr) -> Result<Node, Over> {
        if self.level >= RECURSION {
            return Err(Over);
        }
        self.level += 1;
        let n = self.node(e);
        self.level -= 1;
        n
    }

    fn node(&mut self, e: &Expr) -> Result<Node, Over> {
        match e {
            Expr::Local(l) => match self.env.iter().rev().find(|(slot, _)| *slot == l.slot) {
                Some((_, n)) => {
                    let n = n.clone();
                    self.grow(n.size)?;
                    Ok(n)
                }
                None => {
                    self.grow(1)?;
                    Ok(Node::leaf(e.clone()))
                }
            },
            Expr::Let(l) => {
                let mark = self.env.len();
                let body = self.let_body(l);
                self.env.truncate(mark);
                body
            }
            Expr::Lambda(l) => self.lambda(l),
            // A named function read as a value, `TWICE(SQ,3)`, written in as
            // its `LAMBDA` is, so that a call of it is written in too.
            Expr::Name { name, sheet: None } if !self.library.has(name) => {
                match self.functions.inlined(name) {
                    Some(l) => {
                        let l = Rc::clone(l);
                        self.named(&l)
                    }
                    None => {
                        self.grow(1)?;
                        Ok(Node::leaf(e.clone()))
                    }
                }
            }
            Expr::Apply(a) => {
                let callee = self.expr(&a.callee)?;
                let args = self.all(&a.args)?;
                self.beta(callee, args)
            }
            Expr::Call(name, args) => {
                let args = self.all(args)?;
                if !self.library.has(name) && self.functions.holds(name) {
                    if let Some(l) = self.functions.inlined(name) {
                        let l = Rc::clone(l);
                        let f = self.named(&l)?;
                        return self.beta(f, args);
                    }
                    // A named function a cycle holds or a sheet scopes, kept
                    // a call of its name.
                    self.grow(2)?;
                    let callee = Node::leaf(Expr::Name {
                        name: name.clone(),
                        sheet: None,
                    });
                    return Ok(self.kept(callee, args));
                }
                if !self.library.has(name) && self.functions.calls_value(name) {
                    // A defined name whose text is no `LAMBDA`, called: its
                    // value called where the cell stands, `#VALUE!` where it is
                    // no function's (Excel's save of `RATEN(1)`), and a
                    // function it names applied.
                    self.grow(2)?;
                    let callee = Node::leaf(Expr::Name {
                        name: name.clone(),
                        sheet: None,
                    });
                    return Ok(self.kept(callee, args));
                }
                self.grow(1)?;
                let operands: Vec<&Node> = args.iter().collect();
                let mut n = Self::over(Expr::Empty, &operands);
                n.e = Expr::Call(name.clone(), args.into_iter().map(|a| a.e).collect());
                Ok(n)
            }
            Expr::Neg(x) => {
                let x = self.expr(x)?;
                self.grow(1)?;
                let n = Self::over(Expr::Empty, &[&x]);
                Ok(Node {
                    e: Expr::Neg(Box::new(x.e)),
                    ..n
                })
            }
            Expr::Pos(x) => {
                let x = self.expr(x)?;
                self.grow(1)?;
                let n = Self::over(Expr::Empty, &[&x]);
                Ok(Node {
                    e: Expr::Pos(Box::new(x.e)),
                    ..n
                })
            }
            Expr::Percent(x) => {
                let x = self.expr(x)?;
                self.grow(1)?;
                let n = Self::over(Expr::Empty, &[&x]);
                Ok(Node {
                    e: Expr::Percent(Box::new(x.e)),
                    ..n
                })
            }
            Expr::Binary(..) => self.chain(e),
            Expr::Range(r) => {
                let left = self.expr(&r.left)?;
                let right = self.expr(&r.right)?;
                self.grow(1)?;
                let n = Self::over(Expr::Empty, &[&left, &right]);
                Ok(Node {
                    e: Expr::Range(Box::new(RangeExpr::new(left.e, right.e))),
                    ..n
                })
            }
            other => {
                self.grow(1)?;
                Ok(Node::leaf(other.clone()))
            }
        }
    }

    /// A `LET`'s body with each of its names standing for its value as
    /// written, in order; the caller drops the substitutions after it.
    fn let_body(&mut self, l: &Let) -> Result<Node, Over> {
        for (local, value) in &l.binds {
            let v = self.expr(value)?;
            self.env.push((local.slot, v));
        }
        self.expr(&l.body)
    }

    /// A chain of operators, its left spine round a loop as `Ctx::chain`
    /// evaluates it, each operand a level deeper than the chain.
    fn chain(&mut self, e: &Expr) -> Result<Node, Over> {
        let mut links = Vec::new();
        let mut left = e;
        while let Expr::Binary(op, l, r) = left {
            links.push((*op, &**r));
            left = l;
        }
        let first = self.expr(left)?;
        let mut depth = first.depth;
        let mut size = first.size;
        let mut acc = first.e;
        for (op, r) in links.into_iter().rev() {
            let r = self.expr(r)?;
            self.grow(1)?;
            depth = depth.max(r.depth);
            size += r.size + 1;
            acc = Expr::Binary(op, Box::new(acc), Box::new(r.e));
        }
        Ok(Node {
            e: acc,
            depth,
            size,
        })
    }

    /// A function's value, its parameters given fresh slots and its body
    /// written in, a level deeper than where it is called.
    fn lambda(&mut self, l: &Lambda) -> Result<Node, Over> {
        let mark = self.env.len();
        let mut params = Vec::with_capacity(l.params.len());
        for p in &l.params {
            let fresh = self.fresh(p);
            self.env
                .push((p.slot, Node::leaf(Expr::Local(fresh.clone()))));
            params.push(fresh);
        }
        let body = self.expr(&l.body);
        self.env.truncate(mark);
        let body = body?;
        self.grow(1)?;
        Ok(Node {
            depth: body.depth + 1,
            size: body.size + 1,
            e: Expr::Lambda(Rc::new(Lambda {
                params,
                depth: body.depth,
                body: body.e,
            })),
        })
    }

    /// A named function's `LAMBDA` as a value, written in away from the
    /// calling formula's substitutions, which its body cannot see and whose
    /// slots its own numbering shares.
    fn named(&mut self, l: &Lambda) -> Result<Node, Over> {
        let env = std::mem::take(&mut self.env);
        let n = self.lambda(l);
        self.env = env;
        n
    }

    /// A value called with what was written for its arguments: a function's
    /// value written in, its parameters standing for its arguments; `#VALUE!`
    /// where the count of arguments is not its parameters', as Excel answers;
    /// and anything else kept a call.
    fn beta(&mut self, callee: Node, args: Vec<Node>) -> Result<Node, Over> {
        if let Expr::Lambda(l) = &callee.e {
            if args.len() != l.params.len() {
                self.grow(1)?;
                return Ok(Node::leaf(Expr::Error(ErrorKind::Value)));
            }
            let id = Rc::as_ptr(l);
            if !self.applying.contains(&id) {
                let l = Rc::clone(l);
                let mark = self.env.len();
                for (p, a) in l.params.iter().zip(args) {
                    self.env.push((p.slot, a));
                }
                self.applying.push(id);
                let body = self.expr(&l.body);
                self.applying.pop();
                self.env.truncate(mark);
                return body;
            }
        }
        self.grow(1)?;
        Ok(self.kept(callee, args))
    }

    /// A call kept as a call, `Ctx`'s to compute with its arguments' values.
    fn kept(&self, callee: Node, args: Vec<Node>) -> Node {
        let mut operands: Vec<&Node> = args.iter().collect();
        operands.push(&callee);
        let n = Self::over(Expr::Empty, &operands);
        Node {
            e: Expr::Apply(Box::new(Apply {
                callee: callee.e,
                args: args.into_iter().map(|a| a.e).collect(),
            })),
            ..n
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::calc::library::language;
    use crate::calc::{Calc, Ctx, Value};
    use crate::sheet::{Cell, Content, Workbook};

    /// The named functions the law's formulas call, one of them recursive.
    const NAMED: &[(&str, &str)] = &[
        ("SQ", "=LAMBDA(x,x*x)"),
        ("ADD", "=LAMBDA(a,b,a+b)"),
        ("TWICE", "=LAMBDA(g,x,g(g(x)))"),
        ("SHIFT", "=LAMBDA(x,x+Data!$A$4)"),
        ("DOWN", "=LAMBDA(n,IF(n<=0,0,1+DOWN(n-1)))"),
    ];

    /// A formula's text the law writes, from a seed, over `Data!A1:A5` and
    /// the names in scope: literals, references, operators, `IF`, `LET`, a
    /// `LAMBDA` called in place, bound and handed to a named function, the
    /// named functions, and `MAP` over a range.
    struct Writer {
        seed: u64,
        next: u32,
    }

    impl Writer {
        fn roll(&mut self, n: u64) -> u64 {
            self.seed = self
                .seed
                .wrapping_mul(6_364_136_223_846_793_005)
                .wrapping_add(1_442_695_040_888_963_407);
            (self.seed >> 33) % n
        }

        fn name(&mut self) -> String {
            self.next += 1;
            format!("n_{}", self.next)
        }

        fn expr(&mut self, depth: u32, scope: &mut Vec<String>) -> String {
            let leaf = depth == 0 || self.roll(4) == 0;
            if leaf {
                return match self.roll(3) {
                    0 => format!("{}", self.roll(7)),
                    1 => format!("Data!A{}", 1 + self.roll(5)),
                    _ if !scope.is_empty() => {
                        let k = self.roll(scope.len() as u64) as usize;
                        scope[k].clone()
                    }
                    _ => format!("Data!A{}", 1 + self.roll(5)),
                };
            }
            let d = depth - 1;
            match self.roll(11) {
                0 => format!("({}+{})", self.expr(d, scope), self.expr(d, scope)),
                1 => format!("({}*{})", self.expr(d, scope), self.expr(d, scope)),
                2 => format!("({}-{})", self.expr(d, scope), self.expr(d, scope)),
                3 => format!(
                    "IF({}>{},{},{})",
                    self.expr(d, scope),
                    self.expr(d, scope),
                    self.expr(d, scope),
                    self.expr(d, scope)
                ),
                4 | 5 => {
                    let x = self.name();
                    let value = self.expr(d, scope);
                    scope.push(x.clone());
                    let y = self.name();
                    let second = self.expr(d, scope);
                    scope.push(y.clone());
                    let body = self.expr(d, scope);
                    scope.truncate(scope.len() - 2);
                    format!("LET({x},{value},{y},{second},{body})")
                }
                6 => {
                    let p = self.name();
                    scope.push(p.clone());
                    let body = self.expr(d, scope);
                    scope.pop();
                    format!("LAMBDA({p},{body})({})", self.expr(d, scope))
                }
                7 => format!("SQ({})", self.expr(d, scope)),
                8 => format!(
                    "ADD({},SHIFT({}))",
                    self.expr(d, scope),
                    self.expr(d, scope)
                ),
                9 => {
                    let p = self.name();
                    scope.push(p.clone());
                    let body = self.expr(d, scope);
                    scope.pop();
                    format!("TWICE(LAMBDA({p},{body}),{})", self.expr(d, scope))
                }
                _ => {
                    let p = self.name();
                    scope.push(p.clone());
                    let body = self.expr(d, scope);
                    scope.pop();
                    match self.roll(2) {
                        0 => format!("SUM(MAP(Data!A1:A5,LAMBDA({p},{body})))"),
                        _ => format!(
                            "DOWN(MOD({},5))+{body_free}",
                            self.expr(d, scope),
                            body_free = self.expr(d, scope)
                        ),
                    }
                }
            }
        }
    }

    fn book() -> Workbook {
        let mut wb = Workbook::new();
        let s = wb.ensure_sheet("Data");
        for (row, v) in [(1, 3.0), (2, -1.0), (3, 4.0), (4, 1.5), (5, 0.0)] {
            wb.sheets[s].set(
                row,
                1,
                Cell {
                    content: Content::Number(v),
                    style: 0,
                },
            );
        }
        for (n, t) in NAMED {
            wb.defined_names.push((n.to_string(), t.to_string()));
        }
        wb
    }

    #[test]
    fn a_formula_written_in_computes_what_it_computes_as_written() {
        // The law (KERNEL.29): over generated formulas, the parse with its
        // bindings substituted and its calls written in, and the parse as
        // written, which `Ctx` computes by its bindings and calls, answer
        // alike; and every formula of them the inlining writes in whole is
        // an ordinary formula.
        let wb = book();
        let lib = language();
        let calc = Calc::new(&wb, &lib);
        let functions = Functions::of_model(&wb);
        let mut w = Writer { seed: 29, next: 0 };
        let (mut whole, mut kept) = (0, 0);
        for _ in 0..400 {
            let mut scope = Vec::new();
            let text = w.expr(4, &mut scope);
            let written = formula::parse(&text).expect(&text);
            let inlined = match inline(&written, &functions, &lib) {
                Inlined::To(e) => e,
                Inlined::Same => written.clone(),
                Inlined::TooDeep => panic!("{text}: too deep"),
            };
            if inlined.has_lambda() {
                kept += 1;
            } else {
                whole += 1;
            }
            let answer = |e: &Expr| {
                let ctx = Ctx::new(&calc, &lib);
                ctx.eval(e).and_then(|a| ctx.result(&a)).map(Value::settled)
            };
            assert_eq!(answer(&inlined), answer(&written), "{text}");
        }
        // The law met both readings: most formulas written in whole, the
        // recursion's and the helper's kept.
        assert!(whole > 150 && kept > 50, "{whole} whole, {kept} kept");
    }

    #[test]
    fn a_name_bound_is_read_where_it_is_written_and_never_as_a_defined_name() {
        let e = formula::parse("LET(rate,2,rate*LAMBDA(rate,rate+1)(rate))").unwrap();
        // Three slots: the LET's, and the LAMBDA's parameter; every read a
        // local, none a defined name.
        let mut names = 0;
        e.each_read(&mut |r| {
            if let formula::Read::Name { .. } = r {
                names += 1;
            }
        });
        assert_eq!(names, 0);
        assert!(e.has_lambda());
        // Inlined, an ordinary formula: 2*(2+1).
        let wb = book();
        let lib = language();
        match inline(&e, &Functions::of_model(&wb), &lib) {
            Inlined::To(i) => assert!(!i.has_lambda(), "{i:?}"),
            other => panic!("{other:?}"),
        }
    }

    #[test]
    fn a_named_function_on_a_cycle_or_scoped_by_a_sheet_is_kept_a_call() {
        let wb = book();
        let functions = Functions::of_model(&wb);
        let lib = language();
        assert!(functions.inlined("DOWN").is_none(), "a cycle holds DOWN");
        assert!(functions.inlined("SQ").is_some());
        let e = formula::parse("DOWN(3)+SQ(2)").unwrap();
        let Inlined::To(i) = inline(&e, &functions, &lib) else {
            panic!("inlined");
        };
        // DOWN kept a call of its name, SQ written in.
        assert_eq!(calls_of(&i), Vec::<String>::new());
        assert!(i.has_lambda());
    }
}
