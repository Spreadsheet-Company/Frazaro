//! The netlist bench program's generator and its hand (`KERNEL.54`, its first
//! slice; `docs/BETA_REARVIEW.md`, the item's scoping, decision 6): Alonzo's
//! `CART.9` adder of [`BITS`] bits, five gates a bit (`AUT =MOD(a+b,2)`,
//! `ET =a*b`, `VEL =MAX(a,b)`), its gates shuffled among their own cells
//! `C2:G4097` by a fixed seed with the coincidences kept: each gate's formula a
//! text of its own but for the few that happen to share a far-off gate's, as
//! formulas pasted where they are needed do. Its switches a in `A2:A4097` are
//! all 1 and b in `B2:B4097` all 0, so that the carry-in at `F1`, the edit,
//! ripples through every bit.
//!
//! The recalc runner writes the program into `scripts/recalc/bench/netlist.txt`
//! (`-- --netlist`) and holds the file to it (`-- --check`); and holds every
//! gate the machine computes to the netlist evaluated here, in the order its
//! gates were made, a lineage of its own beside the machine and `Calc`.

/// The adder's width.
pub const BITS: usize = 4096;

/// A gate's function over 0 and 1.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
enum Kind {
    And,
    Or,
    Xor,
}

/// What a gate reads: a switch, by its place among the switches, or a gate.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
enum Node {
    In(usize),
    Gate(usize),
}

#[derive(Clone, Copy, Debug)]
struct Gate {
    kind: Kind,
    a: Node,
    b: Node,
}

/// The adder: its switches' cells, its gates in the order they were made and
/// each one's cell before the shuffle, by bit a row and by role a column.
struct Net {
    inputs: Vec<(u32, u32)>,
    gates: Vec<Gate>,
    home: Vec<(u32, u32)>,
}

/// The adder of `n` bits: bit `i` in row `i + 2`, its switches in columns A
/// and B, its gates `x = AUT(a,b)`, `g = ET(a,b)`, `p = ET(x,c)`,
/// `c' = VEL(g,p)` and `s = AUT(x,c)` in columns C to G; the carry-in at `F1`.
fn build(n: usize) -> Net {
    let mut net = Net {
        inputs: Vec::new(),
        gates: Vec::new(),
        home: Vec::new(),
    };
    let input = |net: &mut Net, at: (u32, u32)| {
        net.inputs.push(at);
        Node::In(net.inputs.len() - 1)
    };
    let a: Vec<Node> = (0..n).map(|i| input(&mut net, (i as u32 + 2, 1))).collect();
    let b: Vec<Node> = (0..n).map(|i| input(&mut net, (i as u32 + 2, 2))).collect();
    let mut c = input(&mut net, (1, 6));
    let put = |net: &mut Net, kind: Kind, a: Node, b: Node, at: (u32, u32)| {
        net.gates.push(Gate { kind, a, b });
        net.home.push(at);
        Node::Gate(net.gates.len() - 1)
    };
    for i in 0..n {
        let r = i as u32 + 2;
        let x = put(&mut net, Kind::Xor, a[i], b[i], (r, 3));
        let g = put(&mut net, Kind::And, a[i], b[i], (r, 4));
        let p = put(&mut net, Kind::And, x, c, (r, 5));
        let carry = put(&mut net, Kind::Or, g, p, (r, 6));
        put(&mut net, Kind::Xor, x, c, (r, 7));
        c = carry;
    }
    net
}

/// `CART.9`'s dice: xorshift with a multiplier.
struct Dice(u64);

impl Dice {
    fn below(&mut self, n: usize) -> usize {
        let mut x = self.0;
        x ^= x >> 12;
        x ^= x << 25;
        x ^= x >> 27;
        self.0 = x;
        (x.wrapping_mul(0x2545_F491_4F6C_DD1D) % n as u64) as usize
    }
}

/// Each gate's cell, the gates permuted among their own cells by `CART.9`'s
/// seed, the coincidences kept.
fn cells(net: &Net) -> Vec<(u32, u32)> {
    let mut perm: Vec<usize> = (0..net.gates.len()).collect();
    let mut d = Dice(0x9E37_79B9_7F4A_7C15 ^ net.gates.len() as u64);
    for i in (1..perm.len()).rev() {
        let j = d.below(i + 1);
        perm.swap(i, j);
    }
    perm.iter().map(|&p| net.home[p]).collect()
}

fn letters(mut col: u32) -> String {
    let mut s = Vec::new();
    while col > 0 {
        s.push(b'A' + ((col - 1) % 26) as u8);
        col = (col - 1) / 26;
    }
    s.reverse();
    String::from_utf8(s).unwrap_or_default()
}

fn a1((row, col): (u32, u32)) -> String {
    format!("{}{row}", letters(col))
}

/// The program, its lines LF: the comments, the switches, then a gate a
/// sentence in the order of its cell.
pub fn program() -> String {
    let net = build(BITS);
    let at = cells(&net);
    let place = |node: Node| match node {
        Node::In(k) => a1(net.inputs[k]),
        Node::Gate(k) => a1(at[k]),
    };
    let last = BITS + 1;
    let mut s = String::new();
    for line in [
        "# The recalculation floors' netlist fixture (KERNEL.54, its first slice): Alonzo's CART.9 adder of",
        "# 4,096 bits, five gates a bit (AUT =MOD(a+b,2), ET =a*b, VEL =MAX(a,b)), its gates shuffled among",
        "# their own cells C2:G4097 by a fixed seed, so that each gate's formula is a text of its own but for",
        "# the few that happen to share a far-off gate's, as formulas pasted where they are needed do. The",
        "# switches a, A2:A4097, are all 1 and b, B2:B4097, all 0; the edit is the carry-in F1, from 0 to 1,",
        "# which ripples through every bit, the deepest cone. The type (KERNEL.56) is a value written into",
        "# A4098, a cell no formula reads below the switches, and into the cell below it each time after.",
        "# Written by core/examples/recalc.rs --netlist (bench/netlist.rs) and held to it by --check: never",
        "# edit it by hand.",
    ] {
        s.push_str(line);
        s.push('\n');
    }
    s.push_str("# bench: edit Output!F1 0 1\n");
    s.push_str(&format!("# bench: type Output!A{}\n", last + 1));
    s.push_str(&format!("# bench: sum Output!C2:G{last}\n\n"));
    s.push_str(&format!("Put 1 into range A2:A{last}.\n"));
    s.push_str(&format!("Put 0 into range B2:B{last}.\n"));
    s.push_str("Put 0 into cell F1.\n");
    let mut order: Vec<usize> = (0..net.gates.len()).collect();
    order.sort_unstable_by_key(|&k| at[k]);
    for k in order {
        let g = net.gates[k];
        let (x, y) = (place(g.a), place(g.b));
        let text = match g.kind {
            Kind::And => format!("={x}*{y}"),
            Kind::Or => format!("=MAX({x},{y})"),
            Kind::Xor => format!("=MOD({x}+{y},2)"),
        };
        s.push_str(&format!(
            "Put formula \"{text}\" into cell {}.\n",
            a1(at[k])
        ));
    }
    s
}

/// Every gate's cell and value by hand, the switches a all 1, b all 0 and the
/// carry-in `cin`: the netlist evaluated in the order its gates were made.
pub fn by_hand(cin: u8) -> Vec<((u32, u32), u8)> {
    let net = build(BITS);
    let at = cells(&net);
    let switch = |k: usize| -> u8 {
        if k < BITS {
            1
        } else if k < 2 * BITS {
            0
        } else {
            cin
        }
    };
    let mut values = vec![0u8; net.gates.len()];
    for (k, g) in net.gates.iter().enumerate() {
        let read = |node: Node, values: &[u8]| match node {
            Node::In(j) => switch(j),
            Node::Gate(j) => values[j],
        };
        let (x, y) = (read(g.a, &values), read(g.b, &values));
        values[k] = match g.kind {
            Kind::And => x & y,
            Kind::Or => x | y,
            Kind::Xor => x ^ y,
        };
    }
    at.into_iter().zip(values).collect()
}
