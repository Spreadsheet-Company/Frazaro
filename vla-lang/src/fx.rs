//! The machine's hash (KERNEL.25, 2026-10-09): Firefox's and rustc's word
//! at a time multiplicative hash, in `rustc-hash` 2's form, written here so
//! that the crate keeps no dependency.
//!
//! The standard library's maps hash with SipHash under keys drawn at random
//! from the system, which resists keys chosen to collide and costs most of
//! a lookup of three integers. An engine's step makes some fifty such
//! lookups for every formula cell in every frame: Alonzo's `CART.1`
//! measured a step of Life at 273 ms with SipHash and 243.6 ms with an Fx
//! hash in its place, nothing else changed. The keys hashed here are a
//! grid's own cells, under the machine's cap of cells and never a network's
//! (`SD-13`), so the worst a crafted grid can do with a hash that is not
//! keyed is slow its own step. No output reads a map's order: nothing walks
//! the machine's maps of values, and every walk of the plan's places sorts
//! its keys first. The hash moves no row and no byte, only the time, and it
//! draws nothing at random.

use std::collections::HashMap;
use std::hash::{BuildHasherDefault, Hasher};

/// The odd constant each word is multiplied by, for the pointer's width.
#[cfg(target_pointer_width = "64")]
const MULTIPLIER: usize = 0xf135_7aea_2e62_a9c5;
#[cfg(not(target_pointer_width = "64"))]
const MULTIPLIER: usize = 0x93d7_65dd;

/// How far the state turns at the end: a product's high bits are its best
/// mixed, and a table takes its index from the low ones.
#[cfg(target_pointer_width = "64")]
const ROTATE: u32 = 26;
#[cfg(not(target_pointer_width = "64"))]
const ROTATE: u32 = 15;

/// The hasher: each word added to the state, and the sum multiplied.
#[derive(Clone, Copy, Default)]
pub struct FxHasher {
    hash: usize,
}

impl FxHasher {
    fn add(&mut self, word: usize) {
        self.hash = self.hash.wrapping_add(word).wrapping_mul(MULTIPLIER);
    }
}

impl Hasher for FxHasher {
    fn write(&mut self, bytes: &[u8]) {
        const WIDTH: usize = std::mem::size_of::<usize>();
        let mut words = bytes.chunks_exact(WIDTH);
        for w in &mut words {
            let mut word = [0u8; WIDTH];
            word.copy_from_slice(w);
            self.add(usize::from_le_bytes(word));
        }
        let rest = words.remainder();
        if !rest.is_empty() {
            let mut word = [0u8; WIDTH];
            word[..rest.len()].copy_from_slice(rest);
            self.add(usize::from_le_bytes(word));
        }
    }

    fn write_u8(&mut self, i: u8) {
        self.add(usize::from(i));
    }

    fn write_u16(&mut self, i: u16) {
        self.add(usize::from(i));
    }

    fn write_u32(&mut self, i: u32) {
        self.add(i as usize);
    }

    fn write_u64(&mut self, i: u64) {
        self.add(i as usize);
        if usize::BITS < 64 {
            self.add((i >> 32) as usize);
        }
    }

    fn write_usize(&mut self, i: usize) {
        self.add(i);
    }

    fn finish(&self) -> u64 {
        self.hash.rotate_left(ROTATE) as u64
    }
}

/// A map under [`FxHasher`].
pub type FxHashMap<K, V> = HashMap<K, V, BuildHasherDefault<FxHasher>>;

#[cfg(test)]
mod tests {
    use super::*;
    use std::collections::HashSet;
    use std::hash::BuildHasher;

    fn hash(cell: (usize, u32, u32)) -> u64 {
        BuildHasherDefault::<FxHasher>::default().hash_one(cell)
    }

    #[test]
    fn a_cell_hashes_the_same_every_time_and_life_s_cells_apart() {
        assert_eq!(hash((5, 2, 3)), hash((5, 2, 3)));
        // Not a proof of a good hash, a check against a broken one: every
        // cell of a 320 by 200 Screen, on its sheet and on its twin's, hashes
        // to a value of its own.
        let mut seen = HashSet::new();
        for sheet in [0usize, 5] {
            for row in 1..=200u32 {
                for col in 1..=320u32 {
                    assert!(seen.insert(hash((sheet, row, col))), "{sheet} {row} {col}");
                }
            }
        }
        assert_eq!(seen.len(), 128_000);
    }

    #[test]
    fn a_map_under_it_is_a_map() {
        let mut m: FxHashMap<(usize, u32, u32), u32> = FxHashMap::default();
        for i in 0..1000u32 {
            m.insert((0, i, i % 7), i);
        }
        assert_eq!(m.len(), 1000);
        assert_eq!(m.get(&(0, 700, 0)), Some(&700));
        assert_eq!(m.remove(&(0, 701, 1)), Some(701));
        assert_eq!(m.get(&(0, 701, 1)), None);
    }
}
