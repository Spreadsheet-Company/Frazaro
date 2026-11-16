//! A function's value and its calls, computed (KERNEL.29, 2026-11-15): what
//! `Ctx` keeps of `LET` and `LAMBDA` where the inlining has not written them
//! in (`super::lambda`).
//!
//! **The bindings.** A `LET`'s value is kept as written where it was written
//! ([`Thunk`]) and computed where a name reads it, by the reading there, the
//! name standing for its expression (the owner's decision 2), once a reading
//! and kept; a call's argument is its value, computed where the call stands
//! (decision 1). A frame holds one binding, under those made before it, so
//! that a function's value keeps the frames it was made among and nothing
//! refers to itself.
//!
//! **The calls**, each counted against two limits for a cell's evaluation,
//! [`DEPTH`] calls open at once and [`CALLS`] made, `#NUM!` past either as
//! Excel answers a recursion too deep for it, so that every frame ends. The
//! stack: a call stands on frames of the language's own evaluation, and a
//! recursion as deep as Excel's would stand on more than `VLA.1` pins a
//! thread's stack at, a formula's 64 levels of nesting in a build of no
//! optimization. So a body is entered only where its nesting fits within
//! those levels below where it stands, counted as `VLA.1` counts a formula's
//! (`super::lambda::kept_depth`) and on the safe side, each formula and body
//! taken to stand as deep as it nests anywhere; a call that would not fit is
//! asked of the top of the cell's evaluation instead ([`Reason::Deep`]),
//! which computes it alone on a fresh stack, remembers its answer by the
//! function and its arguments' values, and evaluates the cell again, which
//! meets the call answered; a deeper call asked for in its turn waits on a
//! list of the language's own. Exact, since every function the library holds
//! is a function of its arguments and of the cell alone (`KERNEL.20`'s
//! premise); about twice the work of the same recursion on a stack deep
//! enough.

use std::cell::OnceCell;
use std::rc::Rc;

use crate::intrinsics::fold;

use super::eval::{Arg, Ctx, Reason};
use super::formula::{self, Apply, Expr, Lambda, Let, Local};
use super::value::{ErrorKind, Value};

/// The deepest a call nests, the calls open at once in a cell's evaluation
/// (the owner's decision 1): past it, `#NUM!`.
pub const DEPTH: u32 = 4096;

/// The most calls a cell's evaluation makes (the owner's decision 1), the
/// applications a helper makes of its function aside, which its array
/// bounds: past it, `#NUM!` for each call after.
pub const CALLS: u64 = 1_000_000;

/// The levels of nesting an evaluation may stand on, a formula's own and the
/// bodies it has entered: `VLA.1`'s limit, which its tests hold on a stack of
/// 1 MiB in every build, a build of no optimization among them.
const LEVELS: u32 = super::formula::NESTING_LIMIT;

/// The levels a call's own frames stand for, beside its body's nesting.
const ENTRY: u32 = 1;

/// A function's value (KERNEL.29): a `LAMBDA`, the bindings it was made
/// among, and the offset its references are read at; a value a formula hands
/// about and calls, never a cell's, which `#CALC!` stands for as a cell's
/// answer and `#VALUE!` where any other value is wanted (Excel's save).
#[derive(Clone)]
pub struct Function(Rc<Closure>);

struct Closure {
    lambda: Rc<Lambda>,
    frame: Option<Rc<Frame>>,
    offset: (i64, i64),
}

impl std::fmt::Debug for Function {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "Function({} parameters)", self.0.lambda.params.len())
    }
}

/// Two functions' values are one when they are the same value.
impl PartialEq for Function {
    fn eq(&self, other: &Function) -> bool {
        Rc::ptr_eq(&self.0, &other.0)
    }
}

impl Function {
    /// How many parameters the function takes.
    pub fn parameters(&self) -> usize {
        self.0.lambda.params.len()
    }

    /// Whether two functions' values compute alike: the same `LAMBDA`, read
    /// at the same offset among bindings alike, as a cell's evaluation begun
    /// again makes them anew (a call remembered, [`Ctx::call_function`]).
    fn same(&self, other: &Function) -> bool {
        let (a, b) = (&*self.0, &*other.0);
        Rc::ptr_eq(&self.0, &other.0)
            || (Rc::ptr_eq(&a.lambda, &b.lambda)
                && a.offset == b.offset
                && frames_same(a.frame.as_deref(), b.frame.as_deref()))
    }
}

/// One binding in scope, under those made before it.
pub(crate) struct Frame {
    slot: u32,
    bound: Bound,
    parent: Option<Rc<Frame>>,
}

enum Bound {
    /// A `LET`'s value, computed where it is read.
    Thunk(Thunk),
    /// A call's argument's value, or the element a helper hands its function.
    Value(Arg<'static>),
    /// An argument left out, read as nothing and told by `ISOMITTED`.
    Omitted,
}

/// A `LET`'s value as written and where it was written, computed once a
/// reading (`Ctx::arrays`, the one thing about where it is read that changes
/// what it computes) and kept.
struct Thunk {
    expr: Rc<Expr>,
    frame: Option<Rc<Frame>>,
    offset: (i64, i64),
    memo: [OnceCell<Result<Arg<'static>, Reason>>; 2],
}

fn frames_same(mut a: Option<&Frame>, mut b: Option<&Frame>) -> bool {
    loop {
        match (a, b) {
            (None, None) => return true,
            (Some(x), Some(y)) if std::ptr::eq(x, y) => return true,
            (Some(x), Some(y)) => {
                if x.slot != y.slot || !bound_same(&x.bound, &y.bound) {
                    return false;
                }
                a = x.parent.as_deref();
                b = y.parent.as_deref();
            }
            _ => return false,
        }
    }
}

fn bound_same(a: &Bound, b: &Bound) -> bool {
    match (a, b) {
        (Bound::Value(x), Bound::Value(y)) => arg_same(x, y),
        (Bound::Omitted, Bound::Omitted) => true,
        (Bound::Thunk(x), Bound::Thunk(y)) => {
            Rc::ptr_eq(&x.expr, &y.expr)
                && x.offset == y.offset
                && frames_same(x.frame.as_deref(), y.frame.as_deref())
        }
        _ => false,
    }
}

fn arg_same(a: &Arg<'_>, b: &Arg<'_>) -> bool {
    match (a, b) {
        (Arg::Function(x), Arg::Function(y)) => x.same(y),
        _ => a == b,
    }
}

fn args_same(a: &[Option<Arg<'static>>], b: &[Option<Arg<'static>>]) -> bool {
    a.len() == b.len()
        && a.iter().zip(b).all(|(x, y)| match (x, y) {
            (None, None) => true,
            (Some(x), Some(y)) => arg_same(x, y),
            _ => false,
        })
}

/// A call asked of the outermost call, to be computed on a fresh stack: the
/// function, its arguments' values, the calls open where it stood and how
/// it was read there.
struct Request {
    f: Function,
    args: Vec<Option<Arg<'static>>>,
    open: u32,
    inside: u32,
}

/// A call computed on a fresh stack: the function, its arguments' values,
/// and its answer.
type Remembered = (
    Function,
    Vec<Option<Arg<'static>>>,
    Result<Arg<'static>, Reason>,
);

/// What an evaluation of `LET` and `LAMBDA` keeps for one cell.
pub(crate) struct Calls {
    /// The bindings in scope where the evaluation stands.
    pub(super) frame: Option<Rc<Frame>>,
    /// The calls made, counted against [`CALLS`].
    made: u64,
    /// The calls open, counted against [`DEPTH`].
    open: u32,
    /// The level of nesting the formula or the body being evaluated began at,
    /// and how deep it nests, so that anything in it stands no deeper than
    /// their sum ([`LEVELS`]). A formula's own nesting is the cell's
    /// evaluation's to say ([`Ctx::nests`]); until it does, its most.
    base: u32,
    span: u32,
    /// Each call computed on a fresh stack, with its answer.
    remembered: Vec<Remembered>,
    /// The call too deep for the stack, asked of the outermost.
    request: Option<Request>,
    /// Each named function read, by its name folded and the sheet it was
    /// read from: one `LAMBDA` however often it is called.
    named: Vec<(String, usize, Option<Rc<Lambda>>)>,
}

impl Default for Calls {
    fn default() -> Calls {
        Calls {
            frame: None,
            made: 0,
            open: 0,
            base: 0,
            span: LEVELS,
            remembered: Vec::new(),
            request: None,
            named: Vec::new(),
        }
    }
}

impl Ctx<'_> {
    /// How deep the formula about to be evaluated nests as `VLA.1` counts
    /// (`super::lambda::kept_depth`), which a cell's evaluation tells its
    /// context when it knows it, so that a call in it is entered where it
    /// fits; a context told nothing takes the formula to nest its most.
    pub(crate) fn nests(&self, depth: u32) {
        self.calls.borrow_mut().span = depth.min(LEVELS);
    }

    /// The top of a cell's evaluation (KERNEL.29): the expression evaluated,
    /// and again after each call too deep for the stack where it stood is
    /// computed alone here and remembered ([`Ctx::call_function`]).
    #[inline(never)]
    pub(super) fn top<'e>(&self, e: &'e Expr) -> Result<Arg<'e>, Reason> {
        self.running.set(true);
        loop {
            let r = self.node(e);
            if matches!(r, Err(Reason::Deep)) && self.calls.borrow().request.is_some() {
                self.settle();
                continue;
            }
            self.running.set(false);
            return r;
        }
    }
    /// A name a `LET` binds or a `LAMBDA` takes, read where it is written:
    /// its binding's value, a `LET`'s computed once a reading.
    pub(super) fn local(&self, l: &Local) -> Result<Arg<'static>, Reason> {
        let frame = self.calls.borrow().frame.clone();
        let mut at = frame.as_deref();
        while let Some(f) = at {
            if f.slot == l.slot {
                return match &f.bound {
                    Bound::Value(v) => Ok(v.clone()),
                    Bound::Omitted => Ok(Arg::Scalar(Value::Empty)),
                    Bound::Thunk(t) => self.force(t),
                };
            }
            at = f.parent.as_deref();
        }
        // A name no binding in scope holds, which no parse leaves.
        Ok(Arg::Scalar(Value::Error(ErrorKind::Name)))
    }

    /// A `LET`'s value, computed as written among the bindings it was written
    /// among, by the reading where it is read (the owner's decision 2), once
    /// a reading.
    fn force(&self, t: &Thunk) -> Result<Arg<'static>, Reason> {
        let k = usize::from(self.arrays());
        if let Some(v) = t.memo[k].get() {
            return v.clone();
        }
        let frame = self.swap_frame(t.frame.clone());
        let offset = self.offset.replace(t.offset);
        let r = self.eval(&t.expr).map(Arg::into_owned);
        self.offset.set(offset);
        self.swap_frame(frame);
        // A step back is no value: the binding is computed again when the
        // cell's call is.
        if !matches!(r, Err(Reason::Deep)) {
            let _ = t.memo[k].set(r.clone());
        }
        r
    }

    /// `LET`: its body among its bindings, each value kept as written until a
    /// name reads it.
    pub(super) fn let_in<'e>(&self, l: &'e Let) -> Result<Arg<'e>, Reason> {
        let saved = self.calls.borrow().frame.clone();
        let mut frame = saved.clone();
        let offset = self.offset.get();
        for (local, value) in &l.binds {
            let thunk = Thunk {
                expr: Rc::clone(value),
                frame: frame.clone(),
                offset,
                memo: [OnceCell::new(), OnceCell::new()],
            };
            frame = Some(Rc::new(Frame {
                slot: local.slot,
                bound: Bound::Thunk(thunk),
                parent: frame,
            }));
        }
        self.swap_frame(frame);
        let r = self.eval(&l.body);
        self.swap_frame(saved);
        r
    }

    /// `LAMBDA`: a function's value, among the bindings in scope.
    pub(super) fn function_value(&self, l: &Rc<Lambda>) -> Arg<'static> {
        Arg::Function(Function(Rc::new(Closure {
            lambda: Rc::clone(l),
            frame: self.calls.borrow().frame.clone(),
            offset: self.offset.get(),
        })))
    }

    /// A value called: a function's, with its arguments' values (the owner's
    /// decision 1), each read where the call stands; `#REF!` where it is a
    /// reference, a cell's or a range's, and `#VALUE!` where it is any other
    /// value, as where its count of arguments is not its parameters' (Excel's
    /// saves: `SPOT(1)`, `LET(r,Data!A2,r(1))` and `BLOCK(1)`; `RATEN(1)`).
    pub(super) fn applied<'e>(&self, a: &'e Apply) -> Result<Arg<'e>, Reason> {
        let callee = self.eval(&a.callee)?;
        let f = match callee {
            Arg::Function(f) => f,
            Arg::Area { .. } => return Ok(Arg::Scalar(Value::Error(ErrorKind::Ref))),
            _ => return Ok(Arg::Scalar(Value::Error(ErrorKind::Value))),
        };
        let args = self.arguments(&a.args)?;
        self.call_function(&f, args, true)
    }

    /// A named function called by its name, with its arguments' values.
    pub(super) fn call_named<'e>(&self, f: &Function, args: &'e [Expr]) -> Result<Arg<'e>, Reason> {
        let args = self.arguments(args)?;
        self.call_function(f, args, true)
    }

    /// A call's arguments' values, each read where the call stands, one left
    /// out none.
    fn arguments(&self, args: &[Expr]) -> Result<Vec<Option<Arg<'static>>>, Reason> {
        let mut out = Vec::with_capacity(args.len());
        for x in args {
            out.push(match x {
                Expr::Empty => None,
                _ => Some(self.eval(x)?.into_owned()),
            });
        }
        Ok(out)
    }

    /// Whether a name a `LAMBDA` takes was given no argument (`ISOMITTED`).
    pub(crate) fn omitted(&self, l: &Local) -> bool {
        let frame = self.calls.borrow().frame.clone();
        let mut at = frame.as_deref();
        while let Some(f) = at {
            if f.slot == l.slot {
                return matches!(f.bound, Bound::Omitted);
            }
            at = f.parent.as_deref();
        }
        false
    }

    /// The named function a name calls from the cell being computed, a
    /// defined name whose text is a `LAMBDA`, read once a cell's evaluation.
    pub(super) fn named_function(&self, name: &str) -> Option<Function> {
        let home = self.env.here().0;
        let k = fold(name);
        let known = self
            .calls
            .borrow()
            .named
            .iter()
            .find(|(n, h, _)| *n == k && *h == home)
            .map(|(_, _, l)| l.clone());
        let lambda = match known {
            Some(l) => l,
            None => {
                // As the grid read it once, where the environment keeps its
                // named functions, so that no text is parsed as deep as the
                // call stands; else its text parsed here.
                let l = self.env.function(name, home).or_else(|| {
                    self.env
                        .name(name, None, home)
                        .and_then(|text| match formula::parse(text) {
                            Ok(Expr::Lambda(l)) => Some(l),
                            _ => None,
                        })
                });
                self.calls.borrow_mut().named.push((k, home, l.clone()));
                l
            }
        }?;
        // A name's references are where the name says, whatever cell reads
        // it, and no binding of the cell's is in its scope.
        Some(Function(Rc::new(Closure {
            lambda,
            frame: None,
            offset: (0, 0),
        })))
    }

    /// A function's value called with its arguments' values (the owner's
    /// decision 1): `#VALUE!` where their count is not its parameters', and
    /// `#NUM!` past [`DEPTH`] calls open or [`CALLS`] made. A call whose body
    /// would not fit the levels below where it stands ([`LEVELS`]) is asked
    /// of the top of the cell's evaluation ([`Reason::Deep`], [`Ctx::top`]),
    /// which computes it there, remembers its answer and evaluates the cell
    /// again, meeting it answered. `counted` is false for a helper's
    /// application of its function, which its array bounds.
    pub(crate) fn call_function(
        &self,
        f: &Function,
        args: Vec<Option<Arg<'static>>>,
        counted: bool,
    ) -> Result<Arg<'static>, Reason> {
        let lambda = &f.0.lambda;
        if args.len() != lambda.params.len() {
            return Ok(Arg::Scalar(Value::Error(ErrorKind::Value)));
        }
        {
            let mut s = self.calls.borrow_mut();
            if counted {
                s.made += 1;
            }
            if s.made > CALLS || s.open >= DEPTH {
                return Ok(Arg::Scalar(Value::Error(ErrorKind::Num)));
            }
            // The latest remembered first, which a cell evaluated again meets
            // first.
            if let Some((_, _, answer)) = s
                .remembered
                .iter()
                .rev()
                .find(|(g, a, _)| g.same(f) && args_same(a, &args))
            {
                return answer.clone();
            }
            // Where it stands, on the safe side: as deep as the formula or
            // the body it is written in nests anywhere. At the top of a
            // computation of one call alone it stands on nothing, and goes in.
            let at = s.base + s.span;
            if at > 0 && at + ENTRY + lambda.depth > LEVELS {
                s.request = Some(Request {
                    f: f.clone(),
                    args,
                    open: s.open,
                    inside: self.inside.get(),
                });
                return Err(Reason::Deep);
            }
        }
        self.enter(f, args)
    }

    /// A function's body computed among its parameters' bindings, at its
    /// offset, a call more open, standing a level below where it was called.
    fn enter(&self, f: &Function, args: Vec<Option<Arg<'static>>>) -> Result<Arg<'static>, Reason> {
        let c = &*f.0;
        let mut frame = c.frame.clone();
        for (p, a) in c.lambda.params.iter().zip(args) {
            frame = Some(Rc::new(Frame {
                slot: p.slot,
                bound: match a {
                    Some(v) => Bound::Value(v),
                    None => Bound::Omitted,
                },
                parent: frame,
            }));
        }
        let (saved, base, span) = {
            let mut s = self.calls.borrow_mut();
            s.open += 1;
            let at = s.base + s.span;
            let base = std::mem::replace(&mut s.base, at + ENTRY);
            let span = std::mem::replace(&mut s.span, c.lambda.depth);
            (std::mem::replace(&mut s.frame, frame), base, span)
        };
        let offset = self.offset.replace(c.offset);
        let r = self.eval(&c.lambda.body).map(Arg::into_owned);
        self.offset.set(offset);
        let mut s = self.calls.borrow_mut();
        s.open -= 1;
        s.base = base;
        s.span = span;
        s.frame = saved;
        r
    }

    /// Each call asked of the top computed there alone, on the stack the
    /// cell's evaluation began on, the deepest first, each remembered with its
    /// answer: a call that asks for a deeper one in its turn waits for that
    /// one.
    fn settle(&self) {
        let Some(first) = self.calls.borrow_mut().request.take() else {
            return;
        };
        let mut pending = vec![first];
        while let Some(top) = pending.last() {
            let (base, span, open) = {
                let mut s = self.calls.borrow_mut();
                (
                    std::mem::replace(&mut s.base, 0),
                    std::mem::replace(&mut s.span, 0),
                    std::mem::replace(&mut s.open, top.open),
                )
            };
            let inside = self.inside.replace(top.inside);
            let r = self.enter(&top.f, top.args.clone());
            self.inside.set(inside);
            {
                let mut s = self.calls.borrow_mut();
                s.base = base;
                s.span = span;
                s.open = open;
            }
            match r {
                Err(Reason::Deep) => match self.calls.borrow_mut().request.take() {
                    Some(deeper) => pending.push(deeper),
                    None => break,
                },
                r => {
                    let done = pending.pop().expect("the call computed");
                    self.calls
                        .borrow_mut()
                        .remembered
                        .push((done.f, done.args, r));
                }
            }
        }
    }
}
