# DATALOG.13 - live steps: the guard clause, before and after

DATALOG.13 replaces `TypeName(d) = "Dictionary"` with
`d Is Nothing Or TypeOf d Is Collection` in the five `VlaDict` wrappers
(`VLA_Runtime.bas`), and with `TypeOf idx Is Collection` in the join index
(`VLA_Relation.bas`). Answers must not change. Only the time should.

The harness is `tools/VLA_Diag13.bas` (new: it times the real wrappers) and
`tools/VLA_Diag12.bas` (DATALOG.12's ladder, now with an optional extra rung).
The eight programs are unchanged. They are in `tools/datalog12_live_steps.md`,
steps 1 to 8.

## Before you start

1. Save the workbook.
2. **Do not reload yet.** Step 1 below must run on the OLD runtime.
3. Import `tools/VLA_Diag13.bas`. If a `VLA_Diag12` module is still in the
   project, remove it and import `tools/VLA_Diag12.bas` again: it has a new
   optional argument. Open the Immediate window (Ctrl+G).

---

## Step 1 - the real wrappers, on the OLD code

**Type:** `D13Wrappers`

**Expect:**
- the version line reads `PF4B.0`;
- the three Dictionary lines (`VlaDictGet`, `VlaDictHas`, `VlaDictSet`) at
  about **0.15 ms each**, the guard's cost plus a little;
- the three Collection lines at a few thousandths of a ms;
- `control: TypeName(d)` about 0.148 ms and `control: TypeOf` 0.000 ms,
  repeating `D12bCost` sections 6 and 9.

This is the first time the wrappers themselves have been timed. Pass 3 timed
a model of them.

**Send back:** the whole output.

## Step 2 - reload, and confirm the new runtime is the one loaded

**Type:** `VlaDevReload`, then Debug > Compile VBAProject, then
`VlaDiagnostics`.

**Expect:** it compiles clean, and the diagnostics line reads
`VLA_Runtime:   DATALOG13.0`. If the reload removed `VLA_Diag13` or
`VLA_Diag12`, import them again.

## Step 3 - Nothing still misses loudly, in the real interpreter

Run this **before any English program runs** after the reload, so the
interpreter's module frame is still empty (Nothing).

**Type:** `eval "zzq-unbound"`

**Expect:** a runtime error saying **`there is nothing stored at key
'zzq-unbound'`**, exactly as before the fix. **Not** error 91, "Object
variable or With block variable not set". That would mean Nothing reached the
dictionary branch.

## Step 4 - the real wrappers, on the NEW code

**Type:** `D13Wrappers`

**Expect:**
- the version line reads `DATALOG13.0`;
- the three Dictionary lines **fall from about 0.15 ms to under 0.01 ms**;
- the three Collection lines about where step 1 left them;
- both controls where step 1 left them. If the controls moved, the machine
  moved, not the code.

**Send back:** the whole output.

## Step 5 - the pure suite

**Type:** `?VlaSelfTest`

**Expect:** your last pure count **+ 6**, 0 failed, with these six new PASS
lines:
- `VlaDict sense: a Dictionary takes the native branch`
- `VlaDict sense: Has reads a Dictionary`
- `VlaDict sense: a Collection takes the fallback branch`
- `VlaDict sense: Has reads the fallback`
- `VlaDict sense: Has on Nothing answers False, raising nothing`
- `VlaDict sense: Get on Nothing misses loudly, naming the key`

The existing `VlaDict fallback: ...` lines (the forced fallback) must still
pass.

## Step 6 - the host suite

**Type:** `?VlaSelfTestHost`

**Expect:** the same count as before this change, 0 failed.

## Step 7 - the DSL suite

**Type:** `?TestDSLs`

**Expect:** the same count as before this change, 0 failed. The join
index's change runs inside every DATALOG and SQL join these pins make.

## Step 8 - the reports

**Type:** `?VerifyReports`

**Expect:** unchanged, on both backends.

---

## Steps 9 to 16 - DATALOG.12's ladder again, after the fix

For each ladder step: delete sheet `D12T<n>` (and any `D12T<n> (2)`) if it
exists, Interpret that step's program from `tools/datalog12_live_steps.md`,
check its cell shows what that file says, then type the line below. The
expected answer rows are the ones that file gives. The harness checks them
and stops on a wrong one.

The prediction was written before this pass. A row should fall from about
1.4 ms to about 0.5 ms, roughly threefold.

| step | type | recorded (DATALOG.12) | predicted now |
|---|---|---|---|
| 9 | `D12Ladder 1, , 10000` | 100: 0.273s, 1,000: 2.254s, 10,000: 25.547s | 1,000: **~0.8s** (0.5-1.0 meets it), 10,000: **~8s** |
| 10 | `D12Ladder 2` | 3,000: 11.867s | ~4s |
| 11 | `D12Ladder 3` | 3,000: 7.6s | ~2.5-3s |
| 12 | `D12Ladder 4` | 3,000: 7.0s | ~2.5s |
| 13 | `D12Ladder 5` | 3,000: 8.1s | ~2.5-3s |
| 14 | `D12Ladder 6` | 300: 4.0s (1,000 skipped) | 300: ~1.3s; 1,000 may now run, at ~7s |
| 15 | `D12Ladder 7` | 300: 5.484s (1,000 skipped) | 300: ~1.8s; 1,000 may now run, at ~8s |
| 16 | `D12Ladder 8` (run last) | 100: 21.648s | 100: ~7s; 250 still skipped by projection |

Step 9's third argument adds a 10,000-row rung after 3,000. The 15s guard
still applies to it: if the fix falls short, it is skipped by projection, and
that skip is itself the answer. A 1,000-row scan above about 1.2s would say
the guard clause was not the 85% of a row DATALOG.12 attributed to it.

**Send back:** every `D12|` line and each step's `formula:` line, and for any
size over about 5s, whether the title bar said "(Not Responding)".

**Afterwards:** delete sheets D12T1 to D12T8 and remove `VLA_Diag12` and
`VLA_Diag13`.
