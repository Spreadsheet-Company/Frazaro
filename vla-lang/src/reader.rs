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
//!
//! Beside them, [`each_form`] (KERNEL.72): the same forms of a whole text one
//! at a time, no token kept, what a machine's load reads a save by, where
//! [`read_forms`] holds the text four bytes a character, its tokens and its
//! forms at once, many times the text at its peak. Nothing in the reference is
//! its twin, and `each_form_reads_what_read_forms_reads` holds the two one.

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
        out.push(parse_form(toks, &mut pos, line_offset, 1)?);
    }
    Ok(out)
}

/// The deepest a list may nest, the top-level list being 1 (VLA.1): the
/// reader refuses a list opened deeper, and so does the expander's walk,
/// so that no stage after them recurses past it. Six times the deepest
/// real program (the corpus golden nests 10 deep, `english.vla` 9, the
/// prelude 6); every stage of the language holds well past it, and the
/// core's emitter, the shallowest, holds 80 in a debug build.
pub const FORM_DEPTH_LIMIT: u32 = 64;

/// VBA `ParseForm`, `depth` the depth of the list a `(` here would open.
fn parse_form(
    toks: &[Token],
    pos: &mut usize,
    line_offset: u32,
    depth: u32,
) -> Result<Form, Refusal> {
    let tok = &toks[*pos];
    let raw_line = tok.line;
    *pos += 1;
    if tok.text == "(" {
        if depth > FORM_DEPTH_LIMIT {
            return Err(raise(
                "vla-form-too-deep",
                &[
                    ("limit", &FORM_DEPTH_LIMIT.to_string()),
                    ("loc", &src_line_tag(raw_line, line_offset)),
                ],
            ));
        }
        // A list is born on the line of its opening paren; below the
        // offset that is the prelude, stored as 0.
        let line = user_line(raw_line, line_offset).unwrap_or(0);
        let mut items = Vec::new();
        while *pos < toks.len() {
            if toks[*pos].text == ")" {
                *pos += 1;
                return Ok(Form::List(List::new(items, line)));
            }
            items.push(parse_form(toks, pos, line_offset, depth + 1)?);
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

/// One token as [`EachForm`] meets it: the text is lent, so that no token is
/// kept and none is made for a form nobody asked for.
enum Met<'a> {
    Open,
    Close,
    /// A string literal's content as written, its escapes in it.
    Literal(&'a str),
    Word(&'a str),
}

/// The walk of a text, by byte (KERNEL.72): every delimiter is ASCII, which
/// no byte of a longer UTF-8 character can be, so that it meets the tokens
/// [`tokenize`] meets, by character, on the lines [`tokenize`] counts.
#[derive(Clone, Debug)]
struct Walk<'a> {
    text: &'a str,
    at: usize,
    line: u32,
}

impl<'a> Walk<'a> {
    /// The next token and the line it began on; `None` at the end of the text.
    fn met(&mut self) -> Option<(Met<'a>, u32)> {
        let b = self.text.as_bytes();
        loop {
            let c = *b.get(self.at)?;
            match c {
                b'\n' => {
                    self.line += 1;
                    self.at += 1;
                }
                b' ' | b'\t' | b'\r' => self.at += 1,
                b';' => {
                    while !matches!(b.get(self.at), None | Some(b'\r' | b'\n')) {
                        self.at += 1;
                    }
                }
                b'(' => {
                    self.at += 1;
                    return Some((Met::Open, self.line));
                }
                b')' => {
                    self.at += 1;
                    return Some((Met::Close, self.line));
                }
                b'"' => {
                    let line = self.line;
                    self.at += 1;
                    let start = self.at;
                    while let Some(&c) = b.get(self.at) {
                        match c {
                            b'"' => break,
                            b'\\' if matches!(b.get(self.at + 1), Some(b'"' | b'\\')) => {
                                self.at += 2;
                            }
                            b'\n' => {
                                self.line += 1;
                                self.at += 1;
                            }
                            _ => self.at += 1,
                        }
                    }
                    let content = &self.text[start..self.at];
                    // An unterminated string ends at the end of the text.
                    self.at = (self.at + 1).min(b.len());
                    return Some((Met::Literal(content), line));
                }
                _ => {
                    let start = self.at;
                    while !matches!(
                        b.get(self.at),
                        None | Some(b'(' | b')' | b' ' | b'\t' | b'\r' | b'\n' | b';' | b'"')
                    ) {
                        self.at += 1;
                    }
                    return Some((Met::Word(&self.text[start..self.at]), self.line));
                }
            }
        }
    }

    /// The form that begins at a token already met, in a text [`each_form`]
    /// has read through: nothing is left to refuse, and its lists nest no
    /// deeper than the limit, so that this recursion is as shallow as
    /// [`parse_form`]'s.
    fn form(&mut self, met: Met<'_>, line: u32) -> Form {
        match met {
            Met::Open => {
                let mut items = Vec::new();
                while let Some((met, at)) = self.met() {
                    if matches!(met, Met::Close) {
                        break;
                    }
                    items.push(self.form(met, at));
                }
                Form::List(List::new(items, line))
            }
            Met::Literal(content) => Form::Str(unescaped(content)),
            Met::Word(word) => Form::Sym(word.to_string()),
            // Never met at a form's place in a text read through.
            Met::Close => Form::List(List::new(Vec::new(), line)),
        }
    }
}

/// A string literal's content without its escapes: `\"` and `\\`, any other
/// backslash standing.
fn unescaped(content: &str) -> String {
    if !content.contains('\\') {
        return content.to_string();
    }
    let mut out = String::with_capacity(content.len());
    let mut rest = content;
    while let Some(k) = rest.find('\\') {
        out.push_str(&rest[..k]);
        let after = &rest[k + 1..];
        if after.starts_with('"') || after.starts_with('\\') {
            out.push_str(&after[..1]);
            rest = &after[1..];
        } else {
            out.push('\\');
            rest = after;
        }
    }
    out.push_str(rest);
    out
}

/// The top-level forms of a text one at a time, each made when it is asked
/// for and the caller's to give up (KERNEL.72): [`read_forms`]'s forms in its
/// order, every list on its line, where that makes every form and every
/// token of the text before the first is read, a text of a million rows many
/// times its own size at the read's peak.
#[derive(Clone, Debug)]
pub struct EachForm<'a> {
    walk: Walk<'a>,
}

/// A text's forms one at a time ([`EachForm`]), once the whole text has been
/// read through and found to read: the refusal, if any, is [`read_forms`]'
/// own, and no form or token is made to find it, so that a caller meets a
/// refusal of the reading before anything a form says, as it does over
/// [`read_forms`].
pub fn each_form(text: &str) -> Result<EachForm<'_>, Refusal> {
    let start = Walk {
        text,
        at: 0,
        line: 1,
    };
    let mut walk = start.clone();
    let mut open = [0u32; FORM_DEPTH_LIMIT as usize];
    let mut depth = 0usize;
    while let Some((met, line)) = walk.met() {
        match met {
            Met::Open => {
                if depth == open.len() {
                    return Err(raise(
                        "vla-form-too-deep",
                        &[
                            ("limit", &FORM_DEPTH_LIMIT.to_string()),
                            ("loc", &src_line_tag(line, 0)),
                        ],
                    ));
                }
                open[depth] = line;
                depth += 1;
            }
            Met::Close => {
                if depth == 0 {
                    return Err(raise(
                        "vla-unexpected-close-paren",
                        &[("loc", &src_line_tag(line, 0))],
                    ));
                }
                depth -= 1;
            }
            Met::Literal(_) | Met::Word(_) => {}
        }
    }
    if depth > 0 {
        return Err(raise(
            "vla-unbalanced-parens",
            &[("loc", &opened_at_tag(open[depth - 1], 0))],
        ));
    }
    Ok(EachForm { walk: start })
}

impl<'a> EachForm<'a> {
    /// Each form still to come that is a list headed by a word: its head and
    /// the line it opened on, read from the text and no form made, so that a
    /// caller who looks for a form by its head before reading the rest makes
    /// no form twice.
    pub fn heads(&self) -> impl Iterator<Item = (&'a str, u32)> {
        let mut walk = self.walk.clone();
        let mut depth = 0usize;
        let mut opened = None;
        std::iter::from_fn(move || {
            while let Some((met, line)) = walk.met() {
                let word = match met {
                    Met::Open => {
                        depth += 1;
                        opened = (depth == 1).then_some(line);
                        continue;
                    }
                    Met::Close => {
                        depth = depth.saturating_sub(1);
                        opened = None;
                        continue;
                    }
                    Met::Word(word) => Some(word),
                    Met::Literal(_) => None,
                };
                if let (Some(line), Some(word)) = (opened.take(), word) {
                    return Some((word, line));
                }
            }
            None
        })
    }
}

impl Iterator for EachForm<'_> {
    type Item = Form;

    fn next(&mut self) -> Option<Form> {
        let (met, line) = self.walk.met()?;
        Some(self.walk.form(met, line))
    }
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

    #[test]
    fn a_list_past_the_depth_limit_is_refused_by_name() {
        // VLA.1, TestFormTooDeep's first pair: 64 nested lists read, and the
        // 65th is refused at its line, where the reader used to recurse
        // until the stack gave out; far past it is the same refusal.
        let nest = |n: usize| format!("{}{}", "(".repeat(n), ")".repeat(n));
        assert_eq!(read_forms(&nest(64)).unwrap().len(), 1);
        let r = read_forms(&nest(65)).unwrap_err();
        assert_eq!(r.id, "vla-form-too-deep");
        assert_eq!(
            r.text,
            "lists nested more than 64 deep (vla line 1) - no program needs that many"
        );
        assert_eq!(
            read_forms(&nest(100_000)).unwrap_err().id,
            "vla-form-too-deep"
        );
    }

    #[test]
    fn each_form_refuses_a_text_before_its_first_form_as_read_forms_does() {
        // KERNEL.72: the three refusals, their words and their lines, the first
        // of them in the text; and no form is made before the whole text has
        // been read through, so that a caller meets the last line's first.
        for text in [
            "(a\n(b",
            "(a)\n)",
            "(a)\n)\n(",
            "(a\n\"open",
            "(a) ; a comment\n(b ; a comment\n",
        ] {
            let e = each_form(text).unwrap_err();
            let want = read_forms(text).unwrap_err();
            assert_eq!((e.id, e.text), (want.id, want.text), "{text:?}");
        }
        let nest = |n: usize| format!("{}{}", "(".repeat(n), ")".repeat(n));
        assert_eq!(each_form(&nest(64)).unwrap().count(), 1);
        assert_eq!(each_form(&nest(65)).unwrap_err().id, "vla-form-too-deep");
        assert_eq!(
            each_form(&nest(100_000)).unwrap_err().id,
            "vla-form-too-deep"
        );
    }

    #[test]
    fn each_form_gives_each_form_on_its_line_and_each_head_without_one() {
        // KERNEL.72: a text of every token's kinds, its forms and their lists'
        // lines read one at a time; what `heads` answers from any place in the
        // walk, a list's first word and the line it opened on, and nothing for
        // an atom, a list with no word first, or a list's own lists.
        let text = "x\n(a\n \"two\nlines\" (b ; c)\n ) \"d\\\"e\"\n)\n( \"s\" t)\r\n(\nu (v))()";
        let forms: Vec<Form> = each_form(text).unwrap().collect();
        assert_eq!(forms, read_forms(text).unwrap());
        let lines: Vec<u32> = forms
            .iter()
            .filter_map(|f| Some(f.as_list()?.line))
            .collect();
        assert_eq!(lines, [2, 7, 8, 9]);
        let inner = forms[1].as_list().unwrap().items[2].as_list().unwrap().line;
        assert_eq!(
            inner, 4,
            "a list inside a list, after a line feed in a string"
        );
        let mut each = each_form(text).unwrap();
        let heads: Vec<(&str, u32)> = each.heads().collect();
        assert_eq!(heads, [("a", 2), ("u", 8)]);
        assert_eq!(each.next(), Some(Form::sym("x")));
        assert_eq!(each.heads().next(), Some(("a", 2)));
        each.next();
        assert_eq!(
            each.heads().next(),
            Some(("u", 8)),
            "from where the walk stands"
        );
        assert_eq!(each_form("").unwrap().count(), 0);
        assert_eq!(each_form(" ; only a comment\n").unwrap().count(), 0);
    }

    const GOLDEN: &str = include_str!("../testdata/instructions_golden.vla");
    const PRELUDE: &str = include_str!("../testdata/prelude.vla");

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
