//! The sentence entry and the proof runners (PORT.6, slice 6d): the part of
//! VLA_SentenceEngine.bas's `ParseStmt` a phrasebook's proofs reach - the
//! raw-form and blank-line gates, B7.4's percent form of Increase, Decrease
//! and Add, the rule walk, the action-call fallback and the parse error -
//! and `RunVocabTest`, `RunVocabFailTest` and `NormalizeWs`, which run the
//! proofs a load collected against the grammar as the file left it.
//!
//! The built-in statement forms (`if`, `repeat`, `count`, `stop`, `while`,
//! `for`, `get`, `give`, `try`, `when`, `create`, `to`, `define`), a raw
//! VLA form's probe and `EnglishToVla`'s program frame are slice 6e's; a
//! sentence that reaches one refuses by name until then, and a proof that
//! does fails for that reason and no other.

use super::grammar::{Grammar, Proof, ProofFailure, ProofKind, ProofMode};
use super::matcher::{is_num_tok, port_pending, Parser};
use super::tokenize::{tokenize, PARA_TOK};
use super::vocab::prov_loc;
use crate::intrinsics::fold;
use crate::messages::{raise, Refusal};

/// VBA `NormalizeWs`: all whitespace runs collapsed to single spaces, for
/// the proof comparison.
pub fn normalize_ws(s: &str) -> String {
    let mut s = s.replace("\r\n", " ").replace(['\n', '\t'], " ");
    while s.contains("  ") {
        s = s.replace("  ", " ");
    }
    s.trim_matches(' ').to_string()
}

/// VBA `StrConv(w, vbProperCase)` on one word.
fn proper_case(w: &str) -> String {
    let mut chars = w.chars();
    match chars.next() {
        Some(c) => c.to_ascii_uppercase().to_string() + chars.as_str(),
        None => String::new(),
    }
}

impl<'g> Parser<'g> {
    /// VBA `ParseStmt`: one statement from pos, as VLA text, indented `ind`
    /// levels. Slice 6d reads what a phrasebook's proofs reach.
    pub fn parse_stmt(&mut self, pos: &mut usize, ind: usize) -> Result<String, Refusal> {
        let pad = " ".repeat(ind * 2);
        // LX.14: where this sentence began, for a function phrase that stops
        // partway.
        if *pos >= 1 && matches!(self.tok(*pos - 1), "" | "." | ":" | PARA_TOK) {
            self.sentence_start = *pos;
        }
        if self.tok(*pos) == PARA_TOK {
            return Err(raise("english-expected-sentence-blank-line", &[]));
        }
        // L0: a token beginning with "(" is a raw VLA form the tokenizer
        // captured whole; its probe-transpile (ValidateRawVla) is 6e's.
        if self.tok(*pos).starts_with('(') {
            self.claim("a raw VLA form");
            return Err(port_pending("a raw VLA form in a sentence"));
        }

        let w = self.peek_word(*pos);
        match w.as_str() {
            "if" | "repeat" | "count" | "while" | "for" | "get" | "give" | "try" | "when"
            | "create" | "to" | "define" => {
                return Err(port_pending(&format!("the statement form '{w}'")));
            }
            "stop" => {
                if self.tok(*pos + 1) == "loop" {
                    self.claim("the Stop form");
                    if self.current_loop().is_empty() {
                        let context = self.sentence_context(*pos);
                        return Err(raise(
                            "english-stop-loop-outside-loop",
                            &[("context", &context)],
                        ));
                    }
                    *pos += 2;
                    self.expect_tok(pos, ".", "'.' after 'Stop the loop'")?;
                    let exit = if self.current_loop() == "do" {
                        "(exit-do)"
                    } else {
                        "(exit-for)"
                    };
                    return Ok(format!("{pad}{exit}"));
                } else if self.tok(*pos + 1) == "." {
                    // B4: inside a value-returning action this must be
                    // Exit Function - Exit Sub there is a compile error.
                    self.claim("the Stop form");
                    *pos += 2;
                    let exit = if self.in_func_def {
                        "(exit-function)"
                    } else {
                        "(exit-sub)"
                    };
                    return Ok(format!("{pad}{exit}"));
                }
                // anything else starting with "stop" tries the phrase rules
            }
            "increase" | "decrease" | "add" => {
                if let Some(r) = self.percent_share_form(&w, pos, &pad)? {
                    return Ok(r);
                }
            }
            "done" | "otherwise" => {
                self.claim("the block structure (Done/Otherwise placement)");
                return Err(raise("english-unexpected-token-block", &[("tok", &w)]));
            }
            _ => {}
        }

        // Phrase rules (DCG productions), in order; first match wins.
        if let Some(txt) = self.try_rules(pos)? {
            return Ok(format!("{pad}{txt}"));
        }

        // Fallback: a name is a call to a defined action - bare ("Greet.")
        // or with named arguments ("Stamp with row of 2 and value of "x".").
        // The call is recorded against the program's definitions (6e).
        if !w.is_empty() {
            if self.tok(*pos + 1) == "." {
                self.claim(&format!("a call to the action '{w}'"));
                *pos += 2;
                return Ok(format!("{pad}({w})"));
            } else if self.tok(*pos + 1) == "with" {
                self.claim(&format!("a call to the action '{w}' with arguments"));
                *pos += 2;
                let mut inner = String::new();
                loop {
                    let v = self.expect_word(pos, "a parameter name after 'with'")?;
                    self.expect_word_is(pos, "of")?;
                    let e = self.parse_expr_req(pos)?;
                    inner.push_str(&format!(" :{v} {e}"));
                    if self.tok(*pos) == "and" {
                        *pos += 1;
                    } else {
                        break;
                    }
                }
                self.expect_tok(pos, ".", "'.' at the end of the sentence")?;
                return Ok(format!("{pad}({w}{inner})"));
            }
        }

        let msg = self.build_parse_error(*pos);
        Err(raise("english-parse-error", &[("msg", &msg)]))
    }

    /// B7.4: "Increase total by 10%." and "Add 10% to total." mean Grow and
    /// Shrink; any other amount carrying % or percent refuses with
    /// directions; an amount without % falls through to the rules.
    fn percent_share_form(
        &mut self,
        w: &str,
        pos: &mut usize,
        pad: &str,
    ) -> Result<Option<String>, Refusal> {
        let p = *pos;
        let is_pct = |t: &str| t == "%" || t == "percent";
        let sh_stop = if w == "add" {
            // Add <N><%|percent> to <name>.
            let sh_n = self.tok(p + 1).to_string();
            if is_pct(self.tok(p + 2))
                && self.tok(p + 3) == "to"
                && !self.word_at(p + 4).is_empty()
                && self.tok(p + 5) == "."
                && (is_num_tok(&sh_n) || !self.word_at(p + 1).is_empty())
            {
                let tgt = self.word_at(p + 4);
                self.claim("the percent form of Increase/Decrease/Add");
                self.mark_assigned(&tgt)?;
                *pos = p + 6;
                return Ok(Some(format!(
                    "{pad}(set! {tgt} (* {tgt} (+ 1 (/ {sh_n} 100))))"
                )));
            }
            "to"
        } else {
            // Increase/Decrease <name> by <N><%|percent>.
            let tgt = self.word_at(p + 1);
            let sh_n = self.tok(p + 3).to_string();
            if !tgt.is_empty()
                && self.tok(p + 2) == "by"
                && is_pct(self.tok(p + 4))
                && self.tok(p + 5) == "."
                && (is_num_tok(&sh_n) || !self.word_at(p + 3).is_empty())
            {
                let op = if w == "increase" { "+" } else { "-" };
                self.claim("the percent form of Increase/Decrease/Add");
                self.mark_assigned(&tgt)?;
                *pos = p + 6;
                return Ok(Some(format!(
                    "{pad}(set! {tgt} (* {tgt} ({op} 1 (/ {sh_n} 100))))"
                )));
            }
            "."
        };
        // Not the simple share: % or percent anywhere in the AMOUNT refuses
        // rather than let either reading win silently.
        let mut scan = if w == "add" {
            p + 1
        } else if self.tok(p + 2) == "by" {
            p + 3
        } else {
            p + 1
        };
        let mut hit = false;
        while scan <= self.toks.count() {
            let t = self.tok(scan);
            if t == "." || t == sh_stop || t == PARA_TOK {
                break;
            }
            if is_pct(t) {
                hit = true;
                break;
            }
            scan += 1;
        }
        if hit {
            self.claim("the percent guard on Increase/Decrease/Add");
            let suggestion = if w == "add" {
                "'Add step to <name>.'".to_string()
            } else {
                format!("'{} <name> by step.'", proper_case(w))
            };
            let sentence = self.render_sentence_at(p);
            let loc = self.line_tag(p);
            return Err(raise(
                "english-percent-mixed-amount",
                &[
                    ("sentence", &sentence),
                    ("suggestion", &suggestion),
                    ("loc", &loc),
                ],
            ));
        }
        Ok(None)
    }

    /// The translation both proof runners make: `EnTokenize`, the alias
    /// rewrite, one `ParseStmt` at position 1, LX.13's read check, and
    /// nothing but paragraph marks allowed after the statement.
    pub fn translate_proof_sentence(&mut self, sentence: &str) -> Result<String, Refusal> {
        let toks = tokenize(sentence).map_err(|e| e.refusal)?;
        self.set_tokens(toks);
        self.cur_line = 0; // LX.13: a proof has no program line
        let mut pos = 1;
        let got = self.parse_stmt(&mut pos, 0)?;
        self.refuse_cell_shaped_read(&got, 0)?;
        while self.tok(pos) == PARA_TOK {
            pos += 1;
        }
        if pos <= self.toks.count() {
            return Err(raise("english-extra-words-after-statement", &[]));
        }
        Ok(got)
    }

    /// VBA `RunVocabTest`: the sentence translates, and reads as the
    /// expected VLA with whitespace normalized.
    pub fn run_vocab_test(&mut self, proof: &Proof) -> Result<(), Refusal> {
        let loc = prov_loc(&proof.source, proof.line, &proof.row_tag);
        let got = match self.translate_proof_sentence(&proof.sentence) {
            Ok(got) => got,
            Err(d) => {
                return Err(raise(
                    "english-test-failed-to-translate",
                    &[
                        ("loc", &loc),
                        ("sentence", &proof.sentence),
                        ("error", &d.text),
                    ],
                ))
            }
        };
        if normalize_ws(&got) != normalize_ws(&proof.expected) {
            return Err(raise(
                "english-test-failed",
                &[
                    ("loc", &loc),
                    ("sentence", &proof.sentence),
                    ("expected", &normalize_ws(&proof.expected)),
                    ("got", &normalize_ws(&got)),
                ],
            ));
        }
        Ok(())
    }

    /// VBA `RunVocabFailTest` (G4): the sentence must refuse, and the
    /// refusal must contain the fragment, whitespace-normalized and
    /// case-blind.
    pub fn run_vocab_fail_test(&mut self, proof: &Proof) -> Result<(), Refusal> {
        let loc = prov_loc(&proof.source, proof.line, &proof.row_tag);
        match self.translate_proof_sentence(&proof.sentence) {
            Ok(got) => Err(raise(
                "english-failtest-translated",
                &[
                    ("loc", &loc),
                    ("sentence", &proof.sentence),
                    ("became", &normalize_ws(&got)),
                ],
            )),
            Err(d) => {
                let message = fold(&normalize_ws(&d.text));
                let wanted = fold(&normalize_ws(&proof.expected));
                if !message.contains(&wanted) {
                    return Err(raise(
                        "english-failtest-message-drifted",
                        &[
                            ("loc", &loc),
                            ("sentence", &proof.sentence),
                            ("wanted", &proof.expected),
                            ("message", &d.text),
                        ],
                    ));
                }
                Ok(())
            }
        }
    }
}

impl Grammar {
    /// `EnglishLoadVocabularyText`'s second loop: the file's proofs run
    /// against the fully loaded grammar, in file order, positives and
    /// negatives alike; each is counted before it runs. Under
    /// [`ProofMode::Raise`] the first failure refuses the load, as the
    /// reference does; under [`ProofMode::Collect`] every failure is kept
    /// and the load goes on, which is what `frazaro prove` scores.
    pub(super) fn run_file_proofs(&mut self, proofs: Vec<Proof>) -> Result<(), Refusal> {
        let mut tests = 0u64;
        let mut fails = 0u64;
        let mut failures: Vec<ProofFailure> = Vec::new();
        let mut first_err: Option<Refusal> = None;
        {
            let mut parser = Parser::new(self);
            for proof in &proofs {
                let r = match proof.kind {
                    ProofKind::Fail => {
                        fails += 1;
                        parser.run_vocab_fail_test(proof)
                    }
                    ProofKind::Success => {
                        tests += 1;
                        parser.run_vocab_test(proof)
                    }
                };
                if let Err(e) = r {
                    match self.proof_mode {
                        ProofMode::Raise => {
                            first_err = Some(e);
                            break;
                        }
                        ProofMode::Collect => failures.push(ProofFailure {
                            proof: proof.clone(),
                            refusal: e,
                        }),
                    }
                }
            }
        }
        self.tests_run += tests;
        self.fails_run += fails;
        self.proofs.extend(proofs);
        self.proof_failures.extend(failures);
        match first_err {
            Some(e) => Err(e),
            None => Ok(()),
        }
    }

    /// One sentence through the grammar as a proof's sentence goes: the
    /// VLA it translates to, or the refusal.
    pub fn translate_sentence(&self, sentence: &str) -> Result<String, Refusal> {
        let mut parser = Parser::new(self);
        parser.translate_proof_sentence(sentence)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const PRELUDE: &str = include_str!("../../../scripts/prelude.vla");
    const ENGLISH: &str = include_str!("../../../scripts/polyglotta/english.vla");
    const DANSK: &str = include_str!("../../../scripts/polyglotta/dansk.vla");
    const DEUTSCHE: &str = include_str!("../../../scripts/polyglotta/deutsche.vla");
    const ESPANOL: &str = include_str!("../../../scripts/polyglotta/espanol.vla");
    const ESPERANTO: &str = include_str!("../../../scripts/polyglotta/esperanto.vla");
    const FRANCAIS: &str = include_str!("../../../scripts/polyglotta/francais.vla");
    const LATIN: &str = include_str!("../../../scripts/polyglotta/latin.vla");
    const PIRATE: &str = include_str!("../../../scripts/polyglotta/pirate.vla");

    fn collecting() -> Grammar {
        let mut g = Grammar::new(PRELUDE);
        g.set_proof_mode(ProofMode::Collect);
        g
    }

    fn translate(rules: &str, sentence: &str) -> Result<String, Refusal> {
        let mut g = Grammar::new(PRELUDE);
        g.load_vocabulary_text(rules, "t.vla").unwrap();
        g.translate_sentence(sentence)
    }

    const SHOW: &str = "(x-vla \"peek {e:expr}\" (debug-print {e}))";

    #[test]
    fn expressions_read_as_the_reference_reads_them() {
        let t = |s: &str| translate(SHOW, s).unwrap();
        assert_eq!(t("Peek 2 plus 3 times 4."), "(debug-print (+ 2 (* 3 4)))");
        assert_eq!(
            t("Peek total divided by 2 minus 1."),
            "(debug-print (- (/ total 2) 1))"
        );
        assert_eq!(t("Peek 15%."), "(debug-print (/ 15 100))");
        assert_eq!(t("Peek 10 percent."), "(debug-print (/ 10 100))");
        assert_eq!(
            t("Peek \"a\" joined with \"b\" followed by c."),
            "(debug-print (& (& \"a\" \"b\") c))"
        );
        assert_eq!(t("Peek cell B2."), "(debug-print (range \"b2\"))");
        assert_eq!(
            t("Peek cell in column C row 5."),
            "(debug-print (cells 5 \"c\"))"
        );
        assert_eq!(
            t("Peek value in column number 2 row k."),
            "(debug-print (cells k 2))"
        );
        assert_eq!(t("Peek value in cell D4."), "(debug-print (range \"d4\"))");
        assert_eq!(
            t("Peek item 2 of found-items."),
            "(debug-print (vlaitem found-items 2))"
        );
        assert_eq!(t("Peek the third."), "(debug-print 3)");
        assert_eq!(
            t("Peek length of cell B2."),
            "(debug-print (len (range \"b2\")))"
        );
        assert_eq!(t("Peek today."), "(debug-print (date))");
        assert_eq!(t("Peek length."), "(debug-print length)");
        assert_eq!(t("Peek cell."), "(debug-print cell)");
        assert_eq!(t("Peek item."), "(debug-print item)");
    }

    #[test]
    fn a_bare_range_is_no_value_and_a_cell_shaped_word_is_refused() {
        let e = translate(SHOW, "Peek a1:b2.").unwrap_err();
        assert_eq!(e.id, "english-parse-error");
        assert!(
            e.text
                .contains("I understood 'peek' - then I expected a value (like 5, "),
            "{}",
            e.text
        );
        let e = translate(SHOW, "Peek a2.").unwrap_err();
        assert_eq!(e.id, "english-cell-shaped-name");
        assert!(e.text.starts_with("'a2' is shaped like a cell"));
    }

    #[test]
    fn conditions_read_as_the_reference_reads_them() {
        const IF: &str = "(x-vla \"check {c:cond}\" (if {c} 1))";
        let t = |s: &str| translate(IF, s).unwrap();
        assert_eq!(t("Check x is 5."), "(if (= x 5) 1)");
        assert_eq!(
            t("Check x is at least 5 and y is not empty."),
            "(if (and (>= x 5) (not (blank? y))) 1)"
        );
        assert_eq!(
            t("Check name contains \"z\" or name starts with \"a\"."),
            "(if (or (positive? (instr 1 name \"z\" vbtextcompare)) (= (instr 1 name \"a\" vbtextcompare) 1)) 1)"
        );
        assert_eq!(t("Check n is divisible by 3."), "(if (zero? (mod n 3)) 1)");
        assert_eq!(
            t("Check range A1:D50 contains \"x\"."),
            "(if (positive? (vlafindtext (range \"a1:d50\") \"x\" \"row\")) 1)"
        );
        assert_eq!(
            t("Check column C does not contain \"x\"."),
            "(if (zero? (vlafindtext (columns \"c\") \"x\" \"row\")) 1)"
        );
        assert_eq!(
            t("Check name ends with \"s\"."),
            "(if (= (lcase (right name (len \"s\"))) (lcase \"s\")) 1)"
        );
    }

    #[test]
    fn the_percent_form_and_its_guard() {
        let g = Grammar::new(PRELUDE);
        assert_eq!(
            g.translate_sentence("Increase total by 10%.").unwrap(),
            "(set! total (* total (+ 1 (/ 10 100))))"
        );
        assert_eq!(
            g.translate_sentence("Decrease total by rate percent.")
                .unwrap(),
            "(set! total (* total (- 1 (/ rate 100))))"
        );
        assert_eq!(
            g.translate_sentence("Add 5% to total.").unwrap(),
            "(set! total (* total (+ 1 (/ 5 100))))"
        );
        let e = g.translate_sentence("Add 5 plus 5% to total.").unwrap_err();
        assert_eq!(e.id, "english-percent-mixed-amount");
        assert!(e.text.contains("'Add step to <name>.'"), "{}", e.text);
        let e = g
            .translate_sentence("Increase total by 10% times 2.")
            .unwrap_err();
        assert!(e.text.contains("'Increase <name> by step.'"), "{}", e.text);
    }

    #[test]
    fn the_parse_error_is_the_reference_s() {
        let mut g = Grammar::new(PRELUDE);
        g.load_vocabulary_text(SHOW, "t.vla").unwrap();
        let e = g.translate_sentence("Frobnicate the widget.").unwrap_err();
        assert_eq!(e.id, "english-parse-error");
        assert!(
            e.text.starts_with("Don't understand: 'frobnicate widget.' No loaded sentence starts with 'frobnicate' - press What can I say? on the Frazaro tab to see all 12 sentences this program understands. (line 1)"),
            "{}",
            e.text
        );
        let e = g.translate_sentence("Peek 5 6.").unwrap_err();
        assert!(
            e.text.starts_with("I understood 'peek 5' - then I expected '.' to end the sentence but found '6'. Did you mean: 'peek {e:expr}' (line 1)"),
            "{}",
            e.text
        );
        // A bare name is an action call, with or without arguments.
        assert_eq!(g.translate_sentence("Greet.").unwrap(), "(greet)");
        assert_eq!(
            g.translate_sentence("Stamp with row of 2 and value of \"x\".")
                .unwrap(),
            "(stamp :row 2 :value \"x\")"
        );
        let e = g.translate_sentence("Done.").unwrap_err();
        assert_eq!(e.id, "english-unexpected-token-block");
        let e = g.translate_sentence("If x is 5, show x.").unwrap_err();
        assert_eq!(e.id, "english-port-pending");
    }

    #[test]
    fn a_failing_proof_refuses_the_load_as_the_reference_does() {
        let mut g = Grammar::new(PRELUDE);
        let e = g
            .load_vocabulary_text(
                "(x-vla \"warm cell {r:cell}\" (debug-print {r}))\n(test-success \"Warm cell B2.\" (debug-print \"b3\"))",
                "t.vla",
            )
            .unwrap_err();
        assert_eq!(e.id, "english-test-failed");
        assert!(
            e.text.starts_with("t.vla line 2: test FAILED"),
            "{}",
            e.text
        );
        assert!(
            e.text.contains("expected: (debug-print \"b3\")")
                && e.text.contains("(debug-print \"b2\")"),
            "{}",
            e.text
        );
        // A refused load leaves its rules registered, as in Excel, so the
        // next load starts from a reset grammar (EnglishResetGrammar).
        g.reset();
        let e = g
            .load_vocabulary_text(
                "(x-vla \"warm cell {r:cell}\" (debug-print {r}))\n(test-fail \"Warm cell B2.\" \"anything\")",
                "t.vla",
            )
            .unwrap_err();
        assert_eq!(e.id, "english-failtest-translated");
        g.reset();
        let e = g
            .load_vocabulary_text(
                "(x-vla \"warm cell {r:cell}\" (debug-print {r}))\n(test-fail \"Warm cell banana.\" \"a unicorn\")",
                "t.vla",
            )
            .unwrap_err();
        assert_eq!(e.id, "english-failtest-message-drifted");
        g.reset();
        g.load_vocabulary_text(
            "(x-vla \"warm cell {r:cell}\" (debug-print {r}))\n(test-success \"Warm cell B2.\" (debug-print \"b2\"))\n(test-fail \"Warm cell banana.\" \"I expected a cell\")",
            "t.vla",
        )
        .unwrap();
        assert_eq!(
            g.vocab_stats(),
            "loaded: 1 rule, 0 macros, 1 test (1 expected fail) from t.vla"
        );
    }

    #[test]
    fn every_dialect_s_proofs_pass_whole() {
        for (text, name) in [
            (DANSK, "dansk.vla"),
            (DEUTSCHE, "deutsche.vla"),
            (ESPERANTO, "esperanto.vla"),
            (FRANCAIS, "francais.vla"),
            (LATIN, "latin.vla"),
            (PIRATE, "pirate.vla"),
        ] {
            // Raise mode: a failing proof would refuse the load, as in Excel.
            let mut g = Grammar::new(PRELUDE);
            g.load_vocabulary_text(text, name)
                .unwrap_or_else(|e| panic!("{name}: {}", e.text));
            assert_eq!(
                g.vocab_stats(),
                format!("loaded: 19 rules, 0 macros, 20 tests (0 expected fails) from {name}")
            );
        }
    }

    /// What a failure under 6d is allowed to be: a shape slice 6e ports.
    fn is_pending(f: &ProofFailure) -> bool {
        f.refusal
            .text
            .contains("is not yet in the core (PORT.6, slice 6e)")
    }

    #[test]
    fn espanol_s_proofs_fail_only_where_a_statement_form_is_pending() {
        let mut g = collecting();
        g.load_vocabulary_text(ESPANOL, "espanol.vla").unwrap();
        assert_eq!(g.proofs().len(), 30);
        let failures = g.proof_failures();
        for f in failures {
            assert!(is_pending(f), "{}", f.refusal.text);
        }
        assert_eq!(
            30 - failures.len(),
            26,
            "{:?}",
            failures.iter().map(|f| f.proof.line).collect::<Vec<_>>()
        );
    }

    #[test]
    fn english_s_proofs_fail_only_where_a_shape_is_pending() {
        let mut g = collecting();
        g.load_vocabulary_text(ENGLISH, "english.vla").unwrap();
        assert_eq!(g.proofs().len(), 482);
        let failures = g.proof_failures();
        for f in failures {
            assert!(is_pending(f), "{}", f.refusal.text);
        }
        // The floor check_prove_floors.ps1 holds english.vla to: the 49
        // pending are 27 sentences that begin with a built-in statement form
        // (If, Create, Repeat, For each) and 22 that bind a G-PROLOG
        // `clause` or `question` slot.
        assert_eq!(
            482 - failures.len(),
            433,
            "{:?}",
            failures.iter().map(|f| f.proof.line).collect::<Vec<_>>()
        );
    }
}
