//! VLA.bas's version predicates (CO.4): `VlaVersionParse`, `VlaVersionCompare`
//! and `VlaVersionAtLeast`, which a phrasebook's `(requires-version "...")`
//! is checked with. The version compared against is this crate's, which
//! `tools/check_version_twin.ps1` holds equal to `VLA_RELEASE_VERSION`.

use std::cmp::Ordering;

use crate::messages::{raise, Refusal};

/// VBA `VlaVersionParse`: `MAJOR.MINOR.PATCH`, each part one to nine
/// digits, spaces around the whole allowed; `None` for anything else.
pub fn parse(s: &str) -> Option<(u32, u32, u32)> {
    let t = s.trim_matches(' ');
    if t.is_empty() {
        return None;
    }
    let parts: Vec<&str> = t.split('.').collect();
    if parts.len() != 3 {
        return None;
    }
    let mut n = [0u32; 3];
    for (i, p) in parts.iter().enumerate() {
        if p.is_empty() || p.len() > 9 || !p.chars().all(|c| c.is_ascii_digit()) {
            return None;
        }
        n[i] = p.parse().ok()?;
    }
    Some((n[0], n[1], n[2]))
}

/// VBA `VlaVersionCompare`: the order of two versions, refusing a
/// malformed one by name.
pub fn compare(a: &str, b: &str) -> Result<Ordering, Refusal> {
    let pa = parse(a).ok_or_else(|| raise("vla-version-malformed", &[("value", a)]))?;
    let pb = parse(b).ok_or_else(|| raise("vla-version-malformed", &[("value", b)]))?;
    Ok(pa.cmp(&pb))
}

/// VBA `VlaVersionAtLeast`: is this build's version at least `minimum`?
pub fn at_least(minimum: &str) -> Result<bool, Refusal> {
    Ok(compare(crate::VERSION, minimum)? != Ordering::Less)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_takes_three_numeric_parts_and_nothing_else() {
        assert_eq!(parse("0.5.2"), Some((0, 5, 2)));
        assert_eq!(parse(" 10.0.1 "), Some((10, 0, 1)));
        assert_eq!(parse("banana"), None);
        assert_eq!(parse("1.2"), None);
        assert_eq!(parse("1.2.3.4"), None);
        assert_eq!(parse("1..3"), None);
        assert_eq!(parse("1.2.x"), None);
        assert_eq!(parse(""), None);
        assert_eq!(parse("1234567890.0.0"), None);
    }

    #[test]
    fn compare_orders_by_part_and_refuses_a_malformed_one() {
        assert_eq!(compare("0.5.2", "0.5.10").unwrap(), Ordering::Less);
        assert_eq!(compare("1.0.0", "0.9.9").unwrap(), Ordering::Greater);
        assert_eq!(compare("0.7.1", "0.7.1").unwrap(), Ordering::Equal);
        let r = compare("x", "0.1.0").unwrap_err();
        assert_eq!(r.id, "vla-version-malformed");
        assert!(at_least("0.0.0").unwrap());
        assert!(!at_least("999.0.0").unwrap());
        assert!(at_least(crate::VERSION).unwrap());
    }
}
