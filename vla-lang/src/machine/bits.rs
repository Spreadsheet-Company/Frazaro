//! A set of the order's positions, a bit each (KERNEL.49's decision 6,
//! 2026-10-22): the cells a frame computes, the formula cells it changed and
//! the twin cells its end changed, each an eighth of a byte a formula cell
//! where the lists of the two changes held room for twelve bytes a cell. A
//! summary bit stands beside each word of 64 positions, set while the word
//! holds any (`KERNEL.54`'s addendum), so that a search for the next position
//! passes 4,096 empty positions a word, where it passed 64 and an idle frame
//! of 391,168 cells alone cost 1.3 µs, and a set is cleared by the words it
//! set, where a list kept each word as it was first set. Sized once a plan
//! is made, so that a frame makes nothing.

/// The order's positions, a bit each, with a summary bit a word. See the
/// module's comment.
#[derive(Clone, Debug, Default)]
pub(crate) struct Bits {
    words: Vec<u64>,
    summary: Vec<u64>,
}

impl Bits {
    /// Every bit clear, room for `n` positions.
    pub(crate) fn sized(&mut self, n: usize) {
        let words = n.div_ceil(64);
        self.words.clear();
        self.words.resize(words, 0);
        self.summary.clear();
        self.summary.resize(words.div_ceil(64), 0);
    }

    /// A position's bit set.
    #[inline]
    pub(crate) fn set(&mut self, p: usize) {
        let w = p / 64;
        self.words[w] |= 1 << (p % 64);
        self.summary[w / 64] |= 1 << (w % 64);
    }

    /// Whether a position's bit is set.
    #[cfg(test)]
    pub(crate) fn get(&self, p: usize) -> bool {
        self.words
            .get(p / 64)
            .is_some_and(|w| w >> (p % 64) & 1 == 1)
    }

    /// Every position from `start` to `end` set, a word at a time
    /// (KERNEL.54: a shape made whole marks its stretches so).
    pub(crate) fn set_range(&mut self, start: usize, end: usize) {
        let mut p = start;
        while p < end {
            // The word's bits from `p` to the range's end or the word's.
            let w = p / 64;
            let (lo, hi) = (p % 64, (end - w * 64).min(64));
            let bits = if hi - lo == 64 {
                !0
            } else {
                ((1u64 << (hi - lo)) - 1) << lo
            };
            self.words[w] |= bits;
            self.summary[w / 64] |= 1 << (w % 64);
            p = w * 64 + hi;
        }
    }

    /// The first position from `from` and before `to` whose bit is set: the
    /// rest of `from`'s word, then the next word holding any, found through
    /// the summary.
    #[inline]
    pub(crate) fn next(&self, from: usize, to: usize) -> Option<usize> {
        if from >= to {
            return None;
        }
        let w = from / 64;
        let word = *self.words.get(w)? & (!0u64 << (from % 64));
        if word != 0 {
            let p = w * 64 + word.trailing_zeros() as usize;
            return (p < to).then_some(p);
        }
        let last = (to - 1) / 64;
        let mut w = w + 1;
        while w <= last {
            let s = w / 64;
            let held = *self.summary.get(s)? & (!0u64 << (w % 64));
            if held == 0 {
                w = (s + 1) * 64;
                continue;
            }
            w = s * 64 + held.trailing_zeros() as usize;
            if w > last {
                return None;
            }
            // A word the summary names holds a bit: nothing clears one bit.
            let p = w * 64 + self.words[w].trailing_zeros() as usize;
            return (p < to).then_some(p);
        }
        None
    }

    /// Every bit clear, by the words the summary names; how many words were
    /// cleared.
    pub(crate) fn clear(&mut self) -> usize {
        let mut cleared = 0;
        for (s, held) in self.summary.iter_mut().enumerate() {
            let mut bits = std::mem::take(held);
            while bits != 0 {
                self.words[s * 64 + bits.trailing_zeros() as usize] = 0;
                bits &= bits - 1;
                cleared += 1;
            }
        }
        cleared
    }

    /// How many positions are set.
    #[cfg(test)]
    pub(crate) fn count(&self) -> usize {
        let mut n = 0;
        for (s, &held) in self.summary.iter().enumerate() {
            let mut bits = held;
            while bits != 0 {
                n += self.words[s * 64 + bits.trailing_zeros() as usize].count_ones() as usize;
                bits &= bits - 1;
            }
        }
        n
    }

    /// The positions set, in order.
    pub(crate) fn iter(&self) -> impl Iterator<Item = usize> + '_ {
        let end = self.words.len() * 64;
        let mut at = 0;
        std::iter::from_fn(move || {
            let p = self.next(at, end)?;
            at = p + 1;
            Some(p)
        })
    }

    /// Whether no position is set.
    #[cfg(test)]
    pub(crate) fn is_empty(&self) -> bool {
        self.summary.iter().all(|&s| s == 0)
    }

    /// The words held, for the tests that hold a set's room to its plan.
    #[cfg(test)]
    pub(crate) fn words(&self) -> usize {
        self.words.len()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn a_set_s_search_finds_each_bit_in_order_past_empty_words() {
        // KERNEL.49: positions far apart, in one word and across a summary
        // word's 4,096, a range over three words, the search bounded below
        // a set bit, and the set cleared by the words it set.
        let mut b = Bits::default();
        b.sized(20_000);
        let set = [3, 64, 65, 4095, 4096, 12_000, 19_999];
        for &p in &set {
            b.set(p);
        }
        assert_eq!(b.iter().collect::<Vec<_>>(), set);
        assert_eq!(b.count(), set.len());
        assert_eq!(b.next(4, 64), None, "bounded before the next bit");
        assert_eq!(b.next(66, 20_000), Some(4095));
        assert_eq!(b.next(4097, 12_000), None);
        assert_eq!(b.next(4097, 12_001), Some(12_000));
        assert!(b.get(4096) && !b.get(4097));
        assert_eq!(b.clear(), 6, "the six words set");
        assert!(b.is_empty() && b.next(0, 20_000).is_none());
        b.set_range(60, 200);
        assert_eq!(b.count(), 140);
        assert_eq!(b.next(0, 20_000), Some(60));
        assert_eq!(b.next(130, 20_000), Some(130));
        assert_eq!(b.iter().last(), Some(199));
        assert_eq!(b.clear(), 4, "the four words of 60 to 199");
        // Every position against a plain search, over ranges of every kind.
        let mut d = 0x9E37_79B9_7F4A_7C15u64;
        let mut roll = |n: u64| {
            d ^= d << 13;
            d ^= d >> 7;
            d ^= d << 17;
            d % n
        };
        for _ in 0..200 {
            b.clear();
            let mut plain = vec![false; 20_000];
            for _ in 0..roll(40) {
                let start = roll(20_000) as usize;
                let end = (start + roll(300) as usize).min(20_000);
                b.set_range(start, end);
                plain[start..end].fill(true);
            }
            let from = roll(20_000) as usize;
            let to = from + roll((20_000 - from) as u64 + 1) as usize;
            let want = (from..to).find(|&p| plain[p]);
            assert_eq!(b.next(from, to), want, "{from}..{to}");
            assert_eq!(b.count(), plain.iter().filter(|&&x| x).count());
        }
    }
}
