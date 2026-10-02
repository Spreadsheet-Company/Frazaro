//! The printer: VLA.bas's `WriteDatum` and `WritePretty`.
//!
//! `WriteDatum` writes a form flat: a symbol as it stands, a string literal
//! quoted with `\` and `"` escaped (backslashes first), a list in
//! parentheses with one space between elements, and `()` for an empty one.
//! `WritePretty` is `VlaExpandText`'s and `VlaFormat`'s layout: a form that
//! fits within 90 columns at its indent prints flat; a wider list opens with
//! its head and gives every remaining element its own line four columns in,
//! closing on the last. Width is measured as VBA's `Len` measures it, in
//! UTF-16 code units.

use crate::form::Form;

/// VBA `WriteDatum`.
pub fn write_datum(f: &Form) -> String {
    let mut out = String::new();
    write_into(f, &mut out);
    out
}

fn write_into(f: &Form, out: &mut String) {
    match f {
        Form::Sym(s) => out.push_str(s),
        Form::Str(s) => {
            out.push('"');
            out.push_str(&s.replace('\\', "\\\\").replace('"', "\\\""));
            out.push('"');
        }
        Form::List(l) => {
            out.push('(');
            for (i, item) in l.items.iter().enumerate() {
                if i > 0 {
                    out.push(' ');
                }
                write_into(item, out);
            }
            out.push(')');
        }
    }
}

/// VBA `Len`: the width of a text in UTF-16 code units.
fn vba_len(s: &str) -> usize {
    s.encode_utf16().count()
}

/// VBA `WritePretty`, with lines joined by CRLF as the VBA joins them.
pub fn write_pretty(f: &Form, indent: usize) -> String {
    let pad = " ".repeat(indent);
    let flat = write_datum(f);
    let Form::List(l) = f else {
        return format!("{pad}{flat}");
    };
    if indent + vba_len(&flat) <= 90 {
        return format!("{pad}{flat}");
    }
    if l.items.is_empty() {
        return format!("{pad}()");
    }
    let mut out = format!("{pad}({}", write_datum(&l.items[0]));
    for item in &l.items[1..] {
        out.push_str("\r\n");
        out.push_str(&write_pretty(item, indent + 4));
    }
    out.push(')');
    out
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::form::List;

    #[test]
    fn write_datum_escapes_as_the_reader_reads() {
        assert_eq!(write_datum(&Form::sym("set!")), "set!");
        assert_eq!(
            write_datum(&Form::string("a \"q\" \\ b")),
            "\"a \\\"q\\\" \\\\ b\""
        );
        assert_eq!(write_datum(&Form::list(vec![])), "()");
        let f = Form::list(vec![
            Form::sym("a"),
            Form::list(vec![Form::sym("b"), Form::sym("c")]),
            Form::string("d"),
        ]);
        assert_eq!(write_datum(&f), "(a (b c) \"d\")");
    }

    #[test]
    fn write_pretty_breaks_a_wide_list_after_its_head() {
        let short = Form::list(vec![Form::sym("a"), Form::sym("b")]);
        assert_eq!(write_pretty(&short, 4), "    (a b)");
        let wide = Form::List(List::new(
            vec![
                Form::sym("begin"),
                Form::list(vec![
                    Form::sym("set!"),
                    Form::sym("x"),
                    Form::string(&"y".repeat(60)),
                ]),
                Form::list(vec![
                    Form::sym("set!"),
                    Form::sym("z"),
                    Form::string(&"w".repeat(60)),
                ]),
            ],
            0,
        ));
        let text = write_pretty(&wide, 0);
        let lines: Vec<&str> = text.split("\r\n").collect();
        assert_eq!(lines.len(), 3);
        assert_eq!(lines[0], "(begin");
        assert!(lines[1].starts_with("    (set! x \"yyy"));
        assert!(lines[2].ends_with("\"))"));
        assert_eq!(write_pretty(&Form::sym("x"), 2), "  x");
    }
}
