//! How Excel 365 stores a formula typed into a cell through `Range.Formula2`
//! (KERNEL.55, its decision 3, the writer's rule, as the saves of the item's
//! probe and of `typed.txt` show it): as a dynamic-array formula wherever its
//! legacy reading could differ, and as a legacy one, a typed `@` written as no
//! more than the legacy reading it is, everywhere else; so that the door's
//! writer writes a sentence's formula as the reference's `Formula2` does and as
//! Excel stores the same text typed.
//!
//! A formula is dynamic where its parse holds, where one value is wanted and
//! no array parameter stands between (its top, a parameter of one value, an
//! operator's operand, an `IF`'s condition, a `CHOOSE`'s index and an
//! `IFERROR`'s value, the last three wherever they stand): a reference written
//! with `:`, `A2:A2` among them, a whole column or row, a range between
//! expressions, an array constant, an answer of many by its bound, or a
//! function answering an array, at any size; an operator whose operand is any
//! of those, wherever it stands; `ROW`, `COLUMN`, `ISFORMULA`, `FORMULATEXT`
//! or `TYPE` outside an array parameter of what would make the formula
//! dynamic at its top, which a legacy formula reads at the cell or at its
//! first row or column; `INDEX` where one value is wanted but at a row and a
//! column each a number past 0; `IFS`, `SWITCH` and `INDIRECT` by their names,
//! and `OFFSET` given a height or a width or moving a range where one value is
//! wanted; and a function called in place (KERNEL.29). A form reads its own
//! arguments, and a name, a `LET`'s binding and a function not held say
//! nothing, but a `SINGLE` inside, which a legacy formula cannot say. Where
//! no save typed a text of the kind, the rule's reading leans dynamic, which
//! reads the text as typed.

use super::eval::{Library, Param};
use super::formula::{parse, Expr, RefExpr};
use super::lambda::{inline, Functions, Inlined};
use super::spill::bound;
use crate::refers::{self, Kind, Shape};

/// Where a part of a formula stands: where one value is wanted, in a
/// parameter handed a reference or an array whole (`SUM`'s), or in an array
/// parameter (`SUMPRODUCT`'s).
#[derive(Clone, Copy, PartialEq, Eq)]
enum Want {
    One,
    Range,
    Array,
}

/// How Excel 365 stores a formula typed with this text (its `=` first or
/// not): whether as a dynamic-array formula, and the text it stores, a typed
/// `@` before a reference written as `SINGLE` in a dynamic formula and left
/// out of a legacy one, as `SINGLE` of a reference alone is; `None` for a
/// text the parse does not read.
pub fn stored(text: &str, library: &Library) -> Option<(bool, String)> {
    let body = text.strip_prefix('=').unwrap_or(text);
    let parsed = parse(&format!("={body}")).ok()?;
    // A `LET`'s names read as what they stand for, where they stand, as an
    // evaluation inlines them (KERNEL.29), beside the parse as written, which
    // says a function called in place.
    let inlined = match inline(&parsed, &Functions::default(), library, false) {
        Inlined::To(e) => Some(e),
        Inlined::Same | Inlined::TooDeep => None,
    };
    let dynamic =
        walk(&parsed, Want::One, library) || inlined.is_some_and(|e| walk(&e, Want::One, library));
    Some((dynamic, written(body, dynamic)))
}

/// Whether a part of a formula makes it dynamic, standing where `want` says.
fn walk(e: &Expr, want: Want, lib: &Library) -> bool {
    match e {
        Expr::Number(_)
        | Expr::Text(_)
        | Expr::Bool(_)
        | Expr::Error(_)
        | Expr::Empty
        | Expr::Local(_)
        | Expr::Name { .. } => false,
        Expr::Array(_) | Expr::Range(_) => want == Want::One,
        Expr::Ref(r) => want == Want::One && spans(r),
        // An operator's operand stands where one value is wanted, wherever the
        // operator stands; what answers many there, a reference written with
        // `:` or any answer its bound gives more than one value, is dynamic.
        Expr::Neg(x) | Expr::Pos(x) | Expr::Percent(x) => walk(x, Want::One, lib),
        Expr::Binary(_, a, b) => walk(a, Want::One, lib) || walk(b, Want::One, lib),
        Expr::Call(name, args) => call(e, name, args, want, lib),
        Expr::Let(l) => {
            l.binds.iter().any(|(_, v)| walk(v, Want::Array, lib)) || walk(&l.body, want, lib)
        }
        Expr::Lambda(_) => want == Want::One,
        Expr::Apply(_) => true,
    }
}

/// Whether a parse calls `SINGLE`, Excel's `@`, anywhere in it.
fn holds_single(e: &Expr) -> bool {
    match e {
        Expr::Call(name, args) => name == "SINGLE" || args.iter().any(holds_single),
        Expr::Neg(x) | Expr::Pos(x) | Expr::Percent(x) => holds_single(x),
        Expr::Binary(_, a, b) => holds_single(a) || holds_single(b),
        Expr::Let(l) => l.binds.iter().any(|(_, v)| holds_single(v)) || holds_single(&l.body),
        Expr::Lambda(l) => holds_single(&l.body),
        _ => false,
    }
}

/// Whether a reference spans more than its cell as written: written with
/// `:`, `A2:A2` among them, or a whole column or row. A span of sheets at one
/// cell does not: no save typed one, and the legacy formula the writer wrote
/// of it before, which Excel loaded and saved as it was, answers `#REF!` as
/// Excel's does, under its `@` too (`intersect.txt`'s `B5` and `B6`).
fn spans(r: &RefExpr) -> bool {
    r.shape != Shape::Cell
}

/// Whether an operand holds many values: a reference written with `:`, a
/// range between expressions, an array constant, or an answer its bound
/// gives more than one value.
fn many(e: &Expr, lib: &Library) -> bool {
    match e {
        Expr::Array(_) | Expr::Range(_) => true,
        Expr::Ref(r) => spans(r),
        _ => bound(e, lib, &mut |_| None).is_some_and(|(rows, columns)| rows > 1 || columns > 1),
    }
}

/// A call's part in the rule.
fn call(e: &Expr, name: &str, args: &[Expr], want: Want, lib: &Library) -> bool {
    let one = |i: usize| args.get(i).is_some_and(|a| walk(a, Want::One, lib));
    let rest = |from: usize, w: Want| args.iter().skip(from).any(|a| walk(a, w, lib));
    match name {
        // Excel's `@`: a reference it reads at the cell, as a legacy formula
        // reads one where one value is wanted; anything else it reads at its
        // answer's first value, where a legacy formula reads a range inside
        // it at the cell; and anywhere but where one value is wanted a legacy
        // formula cannot say it, and a dynamic one keeps it.
        "SINGLE" => {
            want != Want::One
                || args.first().is_some_and(|a| {
                    !matches!(a, Expr::Ref(_) | Expr::Range(_) | Expr::Name { .. })
                        && walk(a, Want::One, lib)
                })
        }
        "IFS" | "SWITCH" | "INDIRECT" => true,
        // Its reference read where the call stands, so that a range it moves
        // answers many where one value is wanted; a height or a width, any.
        "OFFSET" => {
            let moved = if want == Want::One {
                Want::One
            } else {
                Want::Range
            };
            args.len() > 3
                || args.first().is_some_and(|a| walk(a, moved, lib))
                || rest(1, Want::One)
        }
        "IF" | "CHOOSE" | "IFERROR" | "IFNA" => one(0) || rest(1, want),
        "INDEX" => {
            let picked = |i: usize| match args.get(i) {
                Some(Expr::Number(n)) => *n >= 1.0,
                Some(Expr::Neg(x)) => matches!(&**x, Expr::Number(_)),
                _ => false,
            };
            (want == Want::One && !(picked(1) && picked(2)))
                || args.first().is_some_and(|a| walk(a, Want::Range, lib))
                || rest(1, Want::One)
        }
        "XLOOKUP" => (want == Want::One && many(e, lib)) || one(0) || rest(1, Want::Array),
        // These read whole what a legacy formula reads at the cell or at its
        // first row or column: their argument makes the formula dynamic
        // wherever it would at the formula's top, `INDEX`'s column and
        // `OFFSET`'s range among them, but in an array parameter, which a
        // legacy formula reads whole too.
        "ROW" | "COLUMN" | "ISFORMULA" | "FORMULATEXT" | "TYPE" => {
            (want != Want::Array && one(0)) || rest(0, Want::Range)
        }
        _ => {
            let arrayed = lib.arrayed(name).is_some();
            if arrayed && want == Want::One {
                return true;
            }
            let Some(classes) = lib.classes(name) else {
                // A form, which reads its own arguments, or a function the
                // library does not hold: nothing, but an `@` inside, which no
                // legacy formula can be trusted to read alike.
                return args.iter().any(holds_single);
            };
            // Its answer is many where a parameter of one value is handed
            // many, which the walk of that argument says.
            args.iter().enumerate().any(|(i, a)| {
                let w = match classes(i) {
                    Param::Array => Want::Array,
                    Param::Range => Want::Range,
                    _ => Want::One,
                };
                walk(a, w, lib)
            })
        }
    }
}

/// The text Excel stores for a typed text (without its `=`): a typed `@`
/// before a reference written `SINGLE(...)` in a dynamic formula; and in a
/// legacy one left out, as `SINGLE` of a reference alone is, since a legacy
/// formula reads each at its cell already.
fn written(body: &str, dynamic: bool) -> String {
    let chars: Vec<char> = body.chars().collect();
    let refs: Vec<refers::FormulaRef> = refers::scan(body)
        .into_iter()
        .filter(|r| {
            matches!(
                r.kind,
                Kind::Cell | Kind::Range | Kind::Column | Kind::Row | Kind::ThreeD | Kind::Broken
            )
        })
        .collect();
    let at_ref = |i: usize| refs.iter().find(|r| r.pos == i);
    // A text's strings hold no reference the scan finds, so neither an `@`
    // nor a `SINGLE(` inside one is followed by one.
    let mut out = String::with_capacity(body.len() + 8);
    let mut i = 0;
    while i < chars.len() {
        let c = chars[i];
        // `@` before a reference.
        if c == '@' {
            if let Some(r) = at_ref(i + 1) {
                let end = i + 1 + r.span.max(1);
                let reference: String = chars[i + 1..end].iter().collect();
                if dynamic {
                    out.push_str(&format!("SINGLE({reference})"));
                } else {
                    out.push_str(&reference);
                }
                i = end;
                continue;
            }
        }
        // `SINGLE(` of a reference alone, in a legacy formula.
        if !dynamic && starts_single(&chars, i) {
            let open = i + chars[i..].iter().position(|&ch| ch == '(').unwrap_or(0);
            if let Some(r) = at_ref(open + 1) {
                let end = open + 1 + r.span.max(1);
                if chars.get(end) == Some(&')') {
                    out.extend(&chars[open + 1..end]);
                    i = end + 1;
                    continue;
                }
            }
        }
        out.push(c);
        i += 1;
    }
    out
}

/// Whether `SINGLE(`, its `_xlfn.` prefix or not and in any case, begins at a
/// name's first character.
fn starts_single(chars: &[char], i: usize) -> bool {
    if i > 0 && (chars[i - 1].is_alphanumeric() || matches!(chars[i - 1], '_' | '.')) {
        return false;
    }
    let rest: String = chars[i..chars.len().min(i + 13)].iter().collect();
    let upper = rest.to_ascii_uppercase();
    upper.starts_with("SINGLE(") || upper.starts_with("_XLFN.SINGLE(")
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::calc::library;

    #[test]
    fn a_typed_at_is_written_as_excel_stores_it() {
        let lib = library::language();
        for (text, want) in [
            ("=@Data!A1:A6", (false, "Data!A1:A6")),
            ("=SUM(@Data!A1:A6*2)", (false, "SUM(Data!A1:A6*2)")),
            ("=SINGLE(A1:A5)+1", (false, "A1:A5+1")),
            ("=_xlfn.SINGLE(A1:A5)", (false, "A1:A5")),
            ("=SEQUENCE(3)+@A1:A3", (true, "SEQUENCE(3)+SINGLE(A1:A3)")),
            ("=SINGLE(A1+B1)", (false, "SINGLE(A1+B1)")),
            ("=\"@A1\"&SINGLE(A1)", (false, "\"@A1\"&A1")),
            // A reference a move broke, as Excel's mover writes it.
            ("=@#REF!", (false, "#REF!")),
            ("=SEQUENCE(2)+@#REF!", (true, "SEQUENCE(2)+SINGLE(#REF!)")),
            // Where a legacy formula would not read it at the cell, it stays,
            // and the formula is dynamic.
            ("=SUM(@A1:A8)", (true, "SUM(SINGLE(A1:A8))")),
            ("=ROW(SINGLE(A1:A8))", (true, "ROW(SINGLE(A1:A8))")),
            ("=ISEVEN(SINGLE(A1:A3))", (true, "ISEVEN(SINGLE(A1:A3))")),
            // Of an answer of many, its first value, where a legacy formula
            // reads the range at the cell; of a reference alone, its value at
            // the cell, as a legacy formula reads it.
            ("=SINGLE(A1:A3*2)", (true, "SINGLE(A1:A3*2)")),
            ("=1+A1:A3", (true, "1+A1:A3")),
            ("=SINGLE(A1:A3)*2", (false, "A1:A3*2")),
            // No save typed `@` before a call: where the rule cannot show the
            // legacy reading alike, dynamic, which reads the text as typed.
            (
                "=SINGLE(INDEX(A1:C3,0,1))",
                (true, "SINGLE(INDEX(A1:C3,0,1))"),
            ),
            // A reference a call answers, read whole as a reference written
            // with `:` is: a legacy formula meets it at the cell, `#VALUE!`
            // past it, as Excel's save of `info.txt`'s `rest!A59` shows.
            ("=TYPE(INDEX(C1:C3,0))", (true, "TYPE(INDEX(C1:C3,0))")),
            ("=ROW(OFFSET(A1:A3,1,0))", (true, "ROW(OFFSET(A1:A3,1,0))")),
            ("=OFFSET(A1:A3,1,0)", (true, "OFFSET(A1:A3,1,0)")),
            ("=SUM(OFFSET(A1:A3,1,0))", (false, "SUM(OFFSET(A1:A3,1,0))")),
            ("=TYPE(INDEX(C1:C3,2,1))", (false, "TYPE(INDEX(C1:C3,2,1))")),
            // A `LET`'s names read as what they stand for: Excel saved
            // `LET(x,2,x*3)` and a range summed legacy; a range answered, which
            // a legacy formula meets at the cell, is dynamic.
            ("=LET(x,2,x*3)", (false, "LET(x,2,x*3)")),
            ("=LET(x,A1:A3,SUM(x))", (false, "LET(x,A1:A3,SUM(x))")),
            ("=LET(x,A1:A3,x)", (true, "LET(x,A1:A3,x)")),
            ("=LET(x,1,A1:A3)", (true, "LET(x,1,A1:A3)")),
            // A function's value, and a call in place, which Excel repairs in
            // a legacy formula (KERNEL.29's third save).
            ("=LAMBDA(x,x+1)", (true, "LAMBDA(x,x+1)")),
            ("=LAMBDA(x,x+1)(5)", (true, "LAMBDA(x,x+1)(5)")),
            // A height given; a row picked by a negative number, `#VALUE!`
            // either way; a row read from a range where one value is wanted;
            // and an `IF`'s values handed whole where the `IF` is.
            ("=OFFSET(A1,1,0,2)", (true, "OFFSET(A1,1,0,2)")),
            ("=INDEX(A1:C3,-1,1)", (false, "INDEX(A1:C3,-1,1)")),
            (
                "=SUM(INDEX(A1:C3,B1:B2,1))",
                (true, "SUM(INDEX(A1:C3,B1:B2,1))"),
            ),
            ("=SUM(IF(A1>0,B1:B3,0))", (false, "SUM(IF(A1>0,B1:B3,0))")),
            // A name that ends in `SINGLE`, and a string that holds it.
            ("=XSINGLE(A1)", (false, "XSINGLE(A1)")),
            ("=\"SINGLE(A1)\"&A1", (false, "\"SINGLE(A1)\"&A1")),
            // A span of sheets at one cell, as Excel saved `@data:other!A1`.
            ("=@Data:Other!A1", (false, "Data:Other!A1")),
            ("=Data:Other!A1", (false, "Data:Other!A1")),
            ("=SUM(Data:Other!A1)", (false, "SUM(Data:Other!A1)")),
        ] {
            let (dynamic, text_out) = stored(text, &lib).expect(text);
            assert_eq!((dynamic, text_out.as_str()), want, "{text}");
        }
        assert_eq!(stored("=Sales[Amount]", &lib), None);
    }
}
