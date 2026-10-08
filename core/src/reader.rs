//! The reader: VLA.bas's `Tokenize` and `ParseAll`/`ParseForm`, with the
//! source lines they remember and the two refusals they raise.
//!
//! The shape is the VBA's. `Tokenize` walks the text once, by character:
//! a line feed counts a line; a space, a tab or a carriage return is
//! skipped; `(` and `)` are tokens; `;` runs a comment to the end of its
//! line; `"` opens a string literal whose only escapes are `\"` and `\\`
//! (any other backslash stands); everything else runs a symbol to the next
//! delimiter. A string token carries its content behind a leading quote
//! mark, the VBA's own marker, until the parser makes a [`Form::Str`] of it.
//! An unterminated string ends at the end of the text, as the VBA's does.
//!
//! Lines: every token remembers the raw line it began on, and a list
//! remembers the line of its `(` less the caller's offset, stored as 0 when
//! that is below 1 (the prelude's territory; S3.1's explicit zero). A
//! refusal's location is spelled by `SrcLineTag` and `OpenedAtTag`: ` (vla
//! line N)` and ` for the list opened at vla line N`, or nothing at all when
//! the line is the prelude's. The reference also maps lines spliced in by
//! `(include ...)`; nothing here splices, so every line is the main file's.

use crate::form::{Form, List};
use crate::messages::{raise, Refusal};

/// One token: its text, with a string literal's content behind a leading
/// quote mark, and the raw line it began on (1-based).
#[derive(Clone, Debug, PartialEq)]
pub struct Token {
    pub text: String,
    pub line: u32,
}

/// VBA `Tokenize`.
pub fn tokenize(s: &str) -> Vec<Token> {
    let chars: Vec<char> = s.chars().collect();
    let n = chars.len();
    let mut out = Vec::new();
    let mut i = 0;
    let mut line: u32 = 1;
    while i < n {
        let c = chars[i];
        match c {
            '\n' => {
                line += 1;
                i += 1;
            }
            ' ' | '\t' | '\r' => {
                i += 1;
            }
            '(' | ')' => {
                out.push(Token {
                    text: c.to_string(),
                    line,
                });
                i += 1;
            }
            ';' => {
                // A comment runs to the end of its line; the line break is
                // left for the arm above, which counts it.
                while i < n && chars[i] != '\r' && chars[i] != '\n' {
                    i += 1;
                }
            }
            '"' => {
                let start_line = line;
                i += 1;
                let mut buf = String::from("\"");
                while i < n {
                    let c = chars[i];
                    if c == '\\' {
                        match chars.get(i + 1) {
                            Some('"') => {
                                buf.push('"');
                                i += 2;
                            }
                            Some('\\') => {
                                buf.push('\\');
                                i += 2;
                            }
                            _ => {
                                buf.push(c);
                                i += 1;
                            }
                        }
                    } else if c == '"' {
                        i += 1;
                        break;
                    } else {
                        if c == '\n' {
                            line += 1;
                        }
                        buf.push(c);
                        i += 1;
                    }
                }
                out.push(Token {
                    text: buf,
                    line: start_line,
                });
            }
            _ => {
                let start = i;
                while i < n {
                    let c = chars[i];
                    if matches!(c, '(' | ')' | ' ' | '\t' | '\r' | '\n' | ';' | '"') {
                        break;
                    }
                    i += 1;
                }
                out.push(Token {
                    text: chars[start..i].iter().collect(),
                    line,
                });
            }
        }
    }
    out
}

/// VBA `CountLf`: how many line feeds a text holds.
pub fn count_lf(s: &str) -> usize {
    s.matches('\n').count()
}

/// VBA `LabeledLine` with no include map: the line as the user counts it,
/// or `None` for the prelude's territory.
fn user_line(raw_line: u32, line_offset: u32) -> Option<u32> {
    if raw_line > line_offset {
        Some(raw_line - line_offset)
    } else {
        None
    }
}

/// VBA `SrcLineTag`: ` (vla line N)`, or nothing for the prelude.
pub fn src_line_tag(raw_line: u32, line_offset: u32) -> String {
    match user_line(raw_line, line_offset) {
        Some(n) => format!(" (vla line {n})"),
        None => String::new(),
    }
}

/// VBA `OpenedAtTag`: ` for the list opened at vla line N`, or nothing.
pub fn opened_at_tag(raw_line: u32, line_offset: u32) -> String {
    match user_line(raw_line, line_offset) {
        Some(n) => format!(" for the list opened at vla line {n}"),
        None => String::new(),
    }
}

/// VBA `ParseAll`: every top-level form of a token list. `line_offset` is
/// the number of raw lines before the user's line 1 (`mLineOffset`): 0 for
/// a bare read, the prelude's line count plus one for a compile.
pub fn parse_all(toks: &[Token], line_offset: u32) -> Result<Vec<Form>, Refusal> {
    let mut pos = 0;
    let mut out = Vec::new();
    while pos < toks.len() {
        out.push(parse_form(toks, &mut pos, line_offset)?);
    }
    Ok(out)
}

/// VBA `ParseForm`.
fn parse_form(toks: &[Token], pos: &mut usize, line_offset: u32) -> Result<Form, Refusal> {
    let tok = &toks[*pos];
    let raw_line = tok.line;
    *pos += 1;
    if tok.text == "(" {
        // A list is born on the line of its opening paren; below the
        // offset that is the prelude, stored as 0.
        let line = user_line(raw_line, line_offset).unwrap_or(0);
        let mut items = Vec::new();
        while *pos < toks.len() {
            if toks[*pos].text == ")" {
                *pos += 1;
                return Ok(Form::List(List::new(items, line)));
            }
            items.push(parse_form(toks, pos, line_offset)?);
        }
        Err(raise(
            "vla-unbalanced-parens",
            &[("loc", &opened_at_tag(raw_line, line_offset))],
        ))
    } else if tok.text == ")" {
        Err(raise(
            "vla-unexpected-close-paren",
            &[("loc", &src_line_tag(raw_line, line_offset))],
        ))
    } else if let Some(content) = tok.text.strip_prefix('"') {
        Ok(Form::Str(content.to_string()))
    } else {
        Ok(Form::Sym(tok.text.clone()))
    }
}

/// VBA `VlaReadForms`: the forms of a text on its own, every list tagged
/// with its line.
pub fn read_forms(text: &str) -> Result<Vec<Form>, Refusal> {
    parse_all(&tokenize(text), 0)
}

/// VBA `VlaReadFormsWithLines`: the forms, each with its own source line
/// (0 for an atom at top level, which carries none).
pub fn read_forms_with_lines(text: &str) -> Result<Vec<(Form, u32)>, Refusal> {
    Ok(read_forms(text)?
        .into_iter()
        .map(|f| {
            let line = f.as_list().map(|l| l.line).unwrap_or(0);
            (f, line)
        })
        .collect())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::printer::{write_datum, write_pretty};

    fn texts(toks: &[Token]) -> Vec<&str> {
        toks.iter().map(|t| t.text.as_str()).collect()
    }

    #[test]
    fn tokenize_follows_the_vba_character_classes() {
        let toks = tokenize("(set! x \"a \\\"q\\\" b\\n\")  ; a comment\r\n(y)");
        assert_eq!(
            texts(&toks),
            vec!["(", "set!", "x", "\"a \"q\" b\\n", ")", "(", "y", ")"]
        );
        assert_eq!(toks[0].line, 1);
        assert_eq!(toks[5].line, 2);
        // A quote ends a symbol; a tab and a CR are whitespace.
        assert_eq!(
            texts(&tokenize("abc\"d\"\te\rf")),
            vec!["abc", "\"d", "e", "f"]
        );
        // A line feed inside a string counts a line and stays in the text.
        let toks = tokenize("\"two\nlines\" after");
        assert_eq!(toks[0].text, "\"two\nlines");
        assert_eq!(toks[0].line, 1);
        assert_eq!(toks[1].line, 2);
        // An unterminated string ends at the end of the text.
        assert_eq!(texts(&tokenize("\"open")), vec!["\"open"]);
        // Nothing to read reads nothing (TER-8).
        assert!(tokenize("").is_empty());
        assert!(tokenize("  ; only a comment\n").is_empty());
    }

    #[test]
    fn parse_tags_lists_with_their_lines_less_the_offset() {
        let forms = read_forms("(a\n  (b) c)\n\n(d)").unwrap();
        assert_eq!(forms.len(), 2);
        let a = forms[0].as_list().unwrap();
        assert_eq!(a.line, 1);
        assert_eq!(a.items[1].as_list().unwrap().line, 2);
        assert_eq!(forms[1].as_list().unwrap().line, 4);
        // With an offset, the first lines are the prelude's: stored as 0.
        let forms = parse_all(&tokenize("(p)\n(q)\n(r)"), 2).unwrap();
        assert_eq!(forms[0].as_list().unwrap().line, 0);
        assert_eq!(forms[1].as_list().unwrap().line, 0);
        assert_eq!(forms[2].as_list().unwrap().line, 1);
        let with_lines = read_forms_with_lines("x\n(y)").unwrap();
        assert_eq!(with_lines[0], (Form::sym("x"), 0));
        assert_eq!(with_lines[1].1, 2);
    }

    #[test]
    fn the_two_reader_refusals_name_their_lines() {
        let r = read_forms("(a\n(b").unwrap_err();
        assert_eq!(r.id, "vla-unbalanced-parens");
        assert_eq!(
            r.text,
            "unbalanced parentheses: missing ')' for the list opened at vla line 2"
        );
        let r = read_forms("(a)\n)").unwrap_err();
        assert_eq!(r.id, "vla-unexpected-close-paren");
        assert_eq!(r.text, "unexpected ')' (vla line 2)");
        // In the prelude's territory the tag is silent.
        let r = parse_all(&tokenize("(a"), 5).unwrap_err();
        assert_eq!(r.text, "unbalanced parentheses: missing ')'");
    }

    const GOLDEN: &str = include_str!("../../scripts/instructions_golden.vla");
    const PRELUDE: &str = include_str!("../../scripts/prelude.vla");

    #[test]
    fn the_golden_and_the_prelude_read_print_and_read_again() {
        // The counts are what VlaReadForms(...).Count gives in Excel (the
        // owner's step 0 reading confirms them).
        let golden = read_forms(GOLDEN).unwrap();
        assert_eq!(golden.len(), 248);
        let prelude = read_forms(PRELUDE).unwrap();
        assert_eq!(prelude.len(), 46);
        for forms in [&golden, &prelude] {
            let flat: String = forms.iter().map(write_datum).collect::<Vec<_>>().join("\n");
            assert_eq!(&read_forms(&flat).unwrap(), forms);
            let pretty: String = forms
                .iter()
                .map(|f| write_pretty(f, 0))
                .collect::<Vec<_>>()
                .join("\n");
            assert_eq!(&read_forms(&pretty).unwrap(), forms);
        }
        assert!(golden[0].head_is("const"));
        assert!(prelude.iter().all(|f| f.head_is("defmacro")));
    }

    #[test]
    fn count_lf_counts_line_feeds_only() {
        assert_eq!(count_lf("a\r\nb\nc"), 2);
        assert_eq!(count_lf(""), 0);
    }
}
