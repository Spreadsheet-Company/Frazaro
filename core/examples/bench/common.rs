//! What the two runners of the recalculation floors share (`KERNEL.49`, its
//! first slice): the processor's clock, a computation that never changes timed
//! beside every timed step, the hand-written Life it times, the statistics their
//! rows print, and the rows of a built file as the machine loads them. Moved
//! from `recalc.rs` whole, its items made visible to the runner that includes
//! them: `recalc.rs` and `memory.rs` each name this file by its path, an
//! example's module and no example of its own (Cargo makes an example of a
//! file under `examples/` or of a folder's `main.rs`, and of nothing else).
//! Each runner uses part of it, so the part one leaves is no warning.
#![allow(dead_code)]

use std::hint::black_box;
use std::time::Instant;

/// The most the hand-written Life's time may move across a fixture, in percent.
pub(crate) const CLOCK_SPREAD_MAX: f64 = 10.0;
pub(crate) const PROBE_GENERATIONS: usize = 20;
pub(crate) const PROBE_BATCHES: usize = 3;
pub(crate) const START_READINGS: usize = 5;

/// The plane of the Life fixture, `A1:LH200`.
pub(crate) const W: usize = 320;
pub(crate) const H: usize = 200;

// ---- the reference: Life by hand -----------------------------------------

/// The soup of the rule's first frame, its own arithmetic: two linear forms
/// of the seed, the row and the column modulo a prime, their product `h`, and
/// a cell alive when `h * h` modulo the prime is under three eighths of it.
pub(crate) fn soup(seed: u64) -> Vec<u8> {
    let p: u64 = 67_108_859;
    let t: u64 = 25_165_822;
    let mut g = vec![0u8; W * H];
    for r in 2..H as u64 {
        for c in 2..W as u64 {
            let a = (seed * 1_103_515 + r * 1_597_334 + c * 1_830_421 + 12_345) % p;
            let b = (seed * 2_012_345 + r * 1_209_487 + c * 1_949_213 + 54_321) % p;
            let h = (a * b) % p;
            g[(r as usize - 1) * W + (c as usize - 1)] = u8::from((h * h) % p < t);
        }
    }
    g
}

/// One generation of Conway's rule over the plane into `n`, the border dead.
pub(crate) fn generation_into(g: &[u8], n: &mut [u8]) {
    for r in 1..H - 1 {
        for c in 1..W - 1 {
            let i = r * W + c;
            let others = g[i - W - 1]
                + g[i - W]
                + g[i - W + 1]
                + g[i - 1]
                + g[i + 1]
                + g[i + W - 1]
                + g[i + W]
                + g[i + W + 1];
            n[i] = u8::from(others == 3 || (g[i] == 1 && others == 2));
        }
    }
}

pub(crate) fn generation(g: &[u8]) -> Vec<u8> {
    let mut n = vec![0u8; W * H];
    generation_into(g, &mut n);
    n
}

/// The hand-written Life over a fixed workload, timed: milliseconds a
/// generation, the best of a few batches, allocating nothing.
pub(crate) struct Probe {
    fixed: Vec<u8>,
    a: Vec<u8>,
    b: Vec<u8>,
}

impl Probe {
    pub(crate) fn new(fixed: Vec<u8>) -> Probe {
        let a = fixed.clone();
        let b = vec![0u8; fixed.len()];
        Probe { fixed, a, b }
    }

    pub(crate) fn read(&mut self) -> f64 {
        let mut best = f64::INFINITY;
        for _ in 0..PROBE_BATCHES {
            self.a.copy_from_slice(&self.fixed);
            generation_into(&self.a, &mut self.b);
            let t = Instant::now();
            for _ in 0..PROBE_GENERATIONS / 2 {
                generation_into(black_box(&self.a), &mut self.b);
                generation_into(black_box(&self.b), &mut self.a);
            }
            black_box(&self.a);
            best = best.min(t.elapsed().as_secs_f64() * 1e3);
        }
        best / PROBE_GENERATIONS as f64
    }
}

pub(crate) fn median(v: &[f64]) -> f64 {
    let mut s = v.to_vec();
    s.sort_by(f64::total_cmp);
    let n = s.len();
    if n == 0 {
        0.0
    } else if n % 2 == 1 {
        s[n / 2]
    } else {
        (s[n / 2 - 1] + s[n / 2]) / 2.0
    }
}

pub(crate) fn range(v: &[f64]) -> (f64, f64) {
    let lo = v.iter().copied().fold(f64::INFINITY, f64::min);
    let hi = v.iter().copied().fold(0.0, f64::max);
    (lo, hi)
}

/// How far the clock moved: its readings' 10th to 90th percentile over
/// their median, in percent, so that one interrupted reading is not a moved
/// clock and a throttled tenth of the run is.
pub(crate) fn spread(clocks: &[f64]) -> f64 {
    let m = median(clocks);
    if m <= 0.0 || clocks.is_empty() {
        return 0.0;
    }
    let mut c = clocks.to_vec();
    c.sort_by(f64::total_cmp);
    let at = |q: f64| c[((c.len() - 1) as f64 * q).round() as usize];
    (at(0.9) - at(0.1)) / m * 100.0
}

pub(crate) fn drift(now: f64, start: f64) -> f64 {
    if start <= 0.0 {
        0.0
    } else {
        (now / start - 1.0) * 100.0
    }
}

pub(crate) fn thousands(n: u64) -> String {
    let s = n.to_string();
    let mut out = String::new();
    for (i, ch) in s.chars().enumerate() {
        if i > 0 && (s.len() - i) % 3 == 0 {
            out.push(',');
        }
        out.push(ch);
    }
    out
}

// ---- the file read back into the machine ----------------------------------

/// The reader's rows of the built file as the machine loads them: a cell's
/// cached value, which the reader prints as a `cell` row before the cell's
/// `formula` row and the machine would read as a second write of the cell,
/// left out, the machine computing its own and being held to it; and the
/// rows of the sheet Excel keeps the previous frame in, which is the
/// machine's own twin.
pub(crate) fn rows_for_machine(rows: &str, twin: Option<&str>) -> String {
    let lines: Vec<&str> = rows.lines().collect();
    let key = |s: &str, head: &str| -> Option<String> {
        let rest = s.strip_prefix(head)?;
        let mut words = rest.splitn(3, ' ');
        Some(format!("{} {}", words.next()?, words.next()?))
    };
    let twin_word = twin.map(|t| format!("\"{t}\""));
    let mut out = String::with_capacity(rows.len());
    for (i, line) in lines.iter().enumerate() {
        if twin_word.is_some() && line.split(' ').nth(1) == twin_word.as_deref() {
            continue;
        }
        if line.starts_with("(cell ") {
            if let (Some(a), Some(b)) = (
                key(line, "(cell "),
                lines.get(i + 1).and_then(|n| key(n, "(formula ")),
            ) {
                if a == b {
                    continue;
                }
            }
        }
        out.push_str(line);
        out.push('\n');
    }
    out
}
