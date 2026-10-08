//! The distro (KERNEL.2, 2026-10-07): the folder the build tools take, read
//! from its manifest. The design is `web/CALLOSUM.md` §14.1: a distro is a
//! folder of phrasebooks, libraries, examples, a palette, an edition's
//! chrome, a README and the lines a new program opens with, baked by the
//! existing build tools into one page, one add-in and one command-line
//! door; the registry is a repository the user clones, never fetched
//! (`SD-13`), and the build stamp's hashes are the lockfile.
//!
//! The manifest is `distro.vla` in the folder, one `(distro "name" ...)`
//! form of directives, each a path relative to the folder with forward
//! slashes, so that the first distro, `distros/english`, names the corpus
//! files where they are and nothing moves. This module reads the text into
//! a [`Distro`] and refuses a malformed one by name
//! (`distro-manifest-invalid`, what is wrong in the `why` slot); it reads no
//! file and resolves no path, which is a door's business. The page builder
//! (`tools/build_web.ps1`) and the add-in builder (`src/VLA_Build.bas`)
//! read the same file one directive a line, which `tools/check_distro.ps1`
//! holds every manifest to.
//!
//! ```text
//! (distro "english"
//!   (title "Frazaro")                       the page's title, the heading's first word
//!   (tagline "a phrasebook for your workbook")
//!   (opens-with "Set total to 5.")          the sentence a new program's first row offers
//!   (palette (lavender "#e2d8f6") (whisper "#f7f4fc") (deep "#6b4fc8"))
//!   (prelude "../../scripts/prelude.vla")
//!   (phrasebook "english" "../../scripts/polyglotta/english.vla")   the base, loaded first
//!   (overlay "espanol" "...")                what an edition loads after the base; at most one
//!   (dialect "pirate" "...")                 the language picker's alternatives, each over the base
//!   (library "alien" "...")                  a macro library a program includes; listed, not proved
//!   (examples "../../examples")
//!   (addin "Frazaro_English.xlam")           the add-in's file name
//!   (readme "README.md"))
//! ```

use crate::form::Form;
use crate::messages::{raise, Refusal};
use crate::reader::read_forms;

/// A phrasebook or library the manifest names: its name, and its path as
/// written, relative to the folder.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Book {
    pub name: String,
    pub path: String,
}

/// A distro, as its manifest describes it.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Distro {
    pub name: String,
    pub title: String,
    pub tagline: String,
    pub opens_with: String,
    /// The three shades, in the order the manifest gives them: `lavender`,
    /// `whisper` and `deep`, each `#` and six hex digits.
    pub palette: Vec<(String, String)>,
    pub prelude: String,
    pub base: Book,
    pub overlay: Option<Book>,
    pub dialects: Vec<Book>,
    pub libraries: Vec<Book>,
    pub examples: Option<String>,
    pub addin: Option<String>,
    pub readme: Option<String>,
}

impl Distro {
    /// The phrasebooks an edition loads, in order: the base, then the
    /// overlay when there is one. What the add-in embeds and what a door
    /// loads with no `--phrasebook`.
    pub fn chain(&self) -> Vec<&Book> {
        let mut chain = vec![&self.base];
        if let Some(o) = &self.overlay {
            chain.push(o);
        }
        chain
    }

    /// A shade of the palette by its name.
    pub fn shade(&self, name: &str) -> Option<&str> {
        self.palette
            .iter()
            .find(|(n, _)| n == name)
            .map(|(_, v)| v.as_str())
    }
}

/// The names the palette must give, in order.
pub const SHADES: [&str; 3] = ["lavender", "whisper", "deep"];

fn invalid(label: &str, why: &str) -> Refusal {
    raise("distro-manifest-invalid", &[("path", label), ("why", why)])
}

/// A directive's items after its head, each a string, or the refusal
/// naming the directive and the count it wanted.
fn strings<'a>(
    items: &'a [Form],
    head: &str,
    want: usize,
    label: &str,
) -> Result<Vec<&'a str>, Refusal> {
    let mut out = Vec::with_capacity(want);
    for f in items {
        match f {
            Form::Str(s) => out.push(s.as_str()),
            _ => {
                return Err(invalid(
                    label,
                    &format!("({head} ...) holds something that is not a quoted string"),
                ))
            }
        }
    }
    if out.len() != want {
        return Err(invalid(
            label,
            &format!(
                "({head} ...) wants {want} quoted string{}, not {}",
                if want == 1 { "" } else { "s" },
                out.len()
            ),
        ));
    }
    Ok(out)
}

/// A path as a manifest writes one: relative to the folder, forward
/// slashes, not empty.
fn checked_path(path: &str, head: &str, label: &str) -> Result<String, Refusal> {
    if path.is_empty() {
        return Err(invalid(label, &format!("({head} ...) names an empty path")));
    }
    if path.contains('\\') {
        return Err(invalid(
            label,
            &format!("({head} ...) writes a path with a backslash; a manifest's paths use forward slashes"),
        ));
    }
    if path.starts_with('/') || path.chars().nth(1) == Some(':') {
        return Err(invalid(
            label,
            &format!("({head} ...) names an absolute path; a manifest's paths are relative to its folder"),
        ));
    }
    Ok(path.to_string())
}

fn is_name(s: &str) -> bool {
    !s.is_empty()
        && s.chars()
            .all(|c| c.is_ascii_lowercase() || c.is_ascii_digit() || c == '-')
}

fn is_shade(s: &str) -> bool {
    s.len() == 7 && s.starts_with('#') && s[1..].chars().all(|c| c.is_ascii_hexdigit())
}

/// The manifest's text as a [`Distro`], or the refusal that names what is
/// wrong with it; `label` is what the refusal calls the file.
pub fn parse(text: &str, label: &str) -> Result<Distro, Refusal> {
    let forms = read_forms(text).map_err(|r| invalid(label, &r.text))?;
    let mut tops = forms.iter().filter(|f| matches!(f, Form::List(_)));
    let (Some(Form::List(top)), None) = (tops.next(), tops.next()) else {
        return Err(invalid(
            label,
            "it must hold one (distro \"name\" ...) form and nothing else at the top level",
        ));
    };
    if !top.head_is("distro") {
        return Err(invalid(
            label,
            "its top-level form does not begin with distro",
        ));
    }
    let Some(Form::Str(name)) = top.items.get(1) else {
        return Err(invalid(
            label,
            "(distro ...) names no distro: a quoted name comes first",
        ));
    };
    if !is_name(name) {
        return Err(invalid(
            label,
            "a distro's name is lowercase letters, digits and hyphens, the folder's own name",
        ));
    }
    let mut title = None;
    let mut tagline = None;
    let mut opens_with = None;
    let mut palette: Option<Vec<(String, String)>> = None;
    let mut prelude = None;
    let mut base = None;
    let mut overlay = None;
    let mut dialects: Vec<Book> = Vec::new();
    let mut libraries: Vec<Book> = Vec::new();
    let mut examples = None;
    let mut addin = None;
    let mut readme = None;
    let once = |given: bool, head: &str| -> Result<(), Refusal> {
        if given {
            return Err(invalid(label, &format!("({head} ...) is given twice")));
        }
        Ok(())
    };
    for item in &top.items[2..] {
        let Form::List(l) = item else {
            return Err(invalid(
                label,
                "every directive inside (distro ...) is a list",
            ));
        };
        let Ok(head) = l.head_sym() else {
            return Err(invalid(label, "a directive has no head"));
        };
        let rest = &l.items[1..];
        match head.as_str() {
            "title" => {
                once(title.is_some(), "title")?;
                title = Some(strings(rest, "title", 1, label)?[0].to_string());
            }
            "tagline" => {
                once(tagline.is_some(), "tagline")?;
                tagline = Some(strings(rest, "tagline", 1, label)?[0].to_string());
            }
            "opens-with" => {
                once(opens_with.is_some(), "opens-with")?;
                opens_with = Some(strings(rest, "opens-with", 1, label)?[0].to_string());
            }
            "palette" => {
                once(palette.is_some(), "palette")?;
                let mut shades = Vec::with_capacity(SHADES.len());
                for (i, want) in SHADES.iter().enumerate() {
                    let Some(Form::List(p)) = rest.get(i) else {
                        return Err(invalid(
                            label,
                            &format!(
                                "(palette ...) wants ({want} \"#rrggbb\") in position {}",
                                i + 1
                            ),
                        ));
                    };
                    let Ok(n) = p.head_sym() else {
                        return Err(invalid(label, "a shade of the palette has no name"));
                    };
                    if n != *want {
                        return Err(invalid(
                            label,
                            &format!("(palette ...) names {n} where {want} belongs; the shades are lavender, whisper and deep, in that order"),
                        ));
                    }
                    let v = strings(&p.items[1..], want, 1, label)?[0];
                    if !is_shade(v) {
                        return Err(invalid(
                            label,
                            &format!("the shade {want} is {v}, not # and six hex digits"),
                        ));
                    }
                    shades.push((n, v.to_string()));
                }
                if rest.len() != SHADES.len() {
                    return Err(invalid(
                        label,
                        "(palette ...) holds more than its three shades",
                    ));
                }
                palette = Some(shades);
            }
            "prelude" => {
                once(prelude.is_some(), "prelude")?;
                let p = strings(rest, "prelude", 1, label)?[0];
                prelude = Some(checked_path(p, "prelude", label)?);
            }
            "phrasebook" | "overlay" | "dialect" | "library" => {
                let s = strings(rest, &head, 2, label)?;
                if !is_name(s[0]) {
                    return Err(invalid(
                        label,
                        &format!("({head} ...) names {}, and a book's name is lowercase letters, digits and hyphens", s[0]),
                    ));
                }
                let book = Book {
                    name: s[0].to_string(),
                    path: checked_path(s[1], &head, label)?,
                };
                match head.as_str() {
                    "phrasebook" => {
                        once(base.is_some(), "phrasebook")?;
                        base = Some(book);
                    }
                    "overlay" => {
                        once(overlay.is_some(), "overlay")?;
                        overlay = Some(book);
                    }
                    "dialect" => {
                        if dialects.iter().any(|d| d.name == book.name) {
                            return Err(invalid(
                                label,
                                &format!("(dialect \"{}\" ...) is given twice", book.name),
                            ));
                        }
                        dialects.push(book);
                    }
                    _ => {
                        if libraries.iter().any(|d| d.name == book.name) {
                            return Err(invalid(
                                label,
                                &format!("(library \"{}\" ...) is given twice", book.name),
                            ));
                        }
                        libraries.push(book);
                    }
                }
            }
            "examples" => {
                once(examples.is_some(), "examples")?;
                let p = strings(rest, "examples", 1, label)?[0];
                examples = Some(checked_path(p, "examples", label)?);
            }
            "addin" => {
                once(addin.is_some(), "addin")?;
                let a = strings(rest, "addin", 1, label)?[0];
                if a.is_empty() || a.contains('/') || a.contains('\\') {
                    return Err(invalid(
                        label,
                        "(addin ...) is a bare file name, the add-in's own",
                    ));
                }
                addin = Some(a.to_string());
            }
            "readme" => {
                once(readme.is_some(), "readme")?;
                let p = strings(rest, "readme", 1, label)?[0];
                readme = Some(checked_path(p, "readme", label)?);
            }
            other => {
                return Err(invalid(
                    label,
                    &format!("({other} ...) is not a directive this version knows"),
                ));
            }
        }
    }
    let need = |what: &str| invalid(label, &format!("it names no ({what} ...)"));
    let Some(title) = title else {
        return Err(need("title"));
    };
    let Some(tagline) = tagline else {
        return Err(need("tagline"));
    };
    let Some(opens_with) = opens_with else {
        return Err(need("opens-with"));
    };
    let Some(palette) = palette else {
        return Err(need("palette"));
    };
    let Some(prelude) = prelude else {
        return Err(need("prelude"));
    };
    let Some(base) = base else {
        return Err(need("phrasebook"));
    };
    Ok(Distro {
        name: name.clone(),
        title,
        tagline,
        opens_with,
        palette,
        prelude,
        base,
        overlay,
        dialects,
        libraries,
        examples,
        addin,
        readme,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    const ENGLISH: &str = include_str!("../../distros/english/distro.vla");
    const ESPANOL: &str = include_str!("../../distros/espanol/distro.vla");

    #[test]
    fn the_english_distro_reads_as_the_doors_ship_it() {
        let d = parse(ENGLISH, "distros/english/distro.vla").unwrap();
        assert_eq!(d.name, "english");
        assert_eq!(d.title, "Frazaro");
        assert_eq!(d.tagline, "a phrasebook for your workbook");
        assert_eq!(d.opens_with, "Set total to 5.");
        assert_eq!(d.shade("lavender"), Some("#e2d8f6"));
        assert_eq!(d.shade("whisper"), Some("#f7f4fc"));
        assert_eq!(d.shade("deep"), Some("#6b4fc8"));
        assert_eq!(d.shade("mauve"), None);
        assert_eq!(d.prelude, "../../scripts/prelude.vla");
        assert_eq!(d.base.name, "english");
        assert_eq!(d.base.path, "../../scripts/polyglotta/english.vla");
        assert!(d.overlay.is_none());
        let names: Vec<&str> = d.dialects.iter().map(|b| b.name.as_str()).collect();
        assert_eq!(
            names,
            [
                "dansk",
                "deutsche",
                "espanol",
                "esperanto",
                "francais",
                "latin",
                "pirate"
            ]
        );
        assert!(d
            .dialects
            .iter()
            .all(|b| b.path == format!("../../scripts/polyglotta/{}.vla", b.name)));
        assert_eq!(d.libraries.len(), 1);
        assert_eq!(d.libraries[0].name, "alien");
        assert_eq!(d.examples.as_deref(), Some("../../examples"));
        assert_eq!(d.addin.as_deref(), Some("Frazaro_English.xlam"));
        assert_eq!(d.readme.as_deref(), Some("README.md"));
        assert_eq!(d.chain().len(), 1);
    }

    #[test]
    fn the_espanol_distro_is_the_add_in_s_chain_english_then_the_overlay() {
        let d = parse(ESPANOL, "distros/espanol/distro.vla").unwrap();
        assert_eq!(d.name, "espanol");
        let chain: Vec<&str> = d.chain().iter().map(|b| b.path.as_str()).collect();
        assert_eq!(
            chain,
            [
                "../../scripts/polyglotta/english.vla",
                "../../scripts/polyglotta/espanol.vla"
            ]
        );
        assert_eq!(d.addin.as_deref(), Some("Frazaro_Espanol.xlam"));
        assert_eq!(d.dialects.len(), 7);
    }

    fn why(text: &str) -> String {
        let r = parse(text, "x/distro.vla").unwrap_err();
        assert_eq!(r.id, "distro-manifest-invalid");
        assert_eq!(r.source, "VLA-Distro");
        assert!(r.text.starts_with("x/distro.vla "), "{}", r.text);
        r.text
    }

    const MINIMAL: &str = "(distro \"mine\"\n  (title \"T\")\n  (tagline \"t\")\n  (opens-with \"Say hi.\")\n  (palette (lavender \"#e2d8f6\") (whisper \"#f7f4fc\") (deep \"#6b4fc8\"))\n  (prelude \"prelude.vla\")\n  (phrasebook \"mine\" \"mine.vla\"))\n";

    #[test]
    fn a_minimal_manifest_parses_and_each_defect_is_named() {
        let d = parse(MINIMAL, "x/distro.vla").unwrap();
        assert_eq!(d.name, "mine");
        assert!(d.dialects.is_empty() && d.libraries.is_empty());
        assert!(d.examples.is_none() && d.addin.is_none() && d.readme.is_none());
        // Comments are skipped, as everywhere in the language.
        assert!(parse(&format!("; a note\n{MINIMAL}"), "x/distro.vla").is_ok());
        for (text, fragment) in [
            ("", "one (distro \"name\" ...) form"),
            (
                &format!("{MINIMAL}(distro \"two\")"),
                "one (distro \"name\" ...) form",
            ),
            ("(edition \"mine\")", "does not begin with distro"),
            ("(distro mine)", "a quoted name comes first"),
            ("(distro \"Mine\")", "lowercase letters, digits and hyphens"),
            (
                &MINIMAL.replace("(title \"T\")", "(title \"T\")\n  (title \"U\")"),
                "(title ...) is given twice",
            ),
            (
                &MINIMAL.replace("(title \"T\")", "(title T)"),
                "not a quoted string",
            ),
            (
                &MINIMAL.replace("(tagline \"t\")", "(tagline \"t\" \"u\")"),
                "(tagline ...) wants 1 quoted string, not 2",
            ),
            (
                &MINIMAL.replace("(title \"T\")\n", ""),
                "it names no (title ...)",
            ),
            (
                &MINIMAL.replace("(prelude \"prelude.vla\")\n", ""),
                "it names no (prelude ...)",
            ),
            (
                &MINIMAL.replace("  (phrasebook \"mine\" \"mine.vla\")", ""),
                "it names no (phrasebook ...)",
            ),
            (
                &MINIMAL.replace("\"mine.vla\"", "\"..\\\\mine.vla\""),
                "forward slashes",
            ),
            (
                &MINIMAL.replace("\"mine.vla\"", "\"/srv/mine.vla\""),
                "absolute path",
            ),
            (
                &MINIMAL.replace("\"mine.vla\"", "\"C:/mine.vla\""),
                "absolute path",
            ),
            (&MINIMAL.replace("\"mine.vla\"", "\"\""), "empty path"),
            (
                &MINIMAL.replace("(phrasebook \"mine\"", "(phrasebook \"Mine\""),
                "a book's name is lowercase",
            ),
            (
                &MINIMAL.replace("(deep \"#6b4fc8\")", "(deep \"6b4fc8\")"),
                "not # and six hex digits",
            ),
            (
                &MINIMAL.replace("(whisper \"#f7f4fc\")", "(pale \"#f7f4fc\")"),
                "names pale where whisper belongs",
            ),
            (
                &MINIMAL.replace(" (deep \"#6b4fc8\")", ""),
                "(palette ...) wants (deep \"#rrggbb\") in position 3",
            ),
            (
                &MINIMAL.replace("(deep \"#6b4fc8\")", "(deep \"#6b4fc8\") (ink \"#000000\")"),
                "more than its three shades",
            ),
            (
                &MINIMAL.replace("(prelude \"prelude.vla\")", "(prelude \"prelude.vla\")\n  (dialect \"a\" \"a.vla\")\n  (dialect \"a\" \"b.vla\")"),
                "(dialect \"a\" ...) is given twice",
            ),
            (
                &MINIMAL.replace("(prelude \"prelude.vla\")", "(prelude \"prelude.vla\")\n  (addin \"out/x.xlam\")"),
                "a bare file name",
            ),
            (
                &MINIMAL.replace("(prelude \"prelude.vla\")", "(prelude \"prelude.vla\")\n  (colour \"red\")"),
                "(colour ...) is not a directive this version knows",
            ),
            (
                &MINIMAL.replace("(prelude \"prelude.vla\")", "(prelude \"prelude.vla\")\n  \"loose\""),
                "every directive inside (distro ...) is a list",
            ),
            ("(distro \"mine\" (title \"T\"", "x/distro.vla "),
        ] {
            let text = why(text);
            assert!(text.contains(fragment), "{text:?} should name {fragment:?}");
        }
    }
}
