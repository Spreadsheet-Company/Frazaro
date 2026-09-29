# The inbox

*Findings from cloud sessions, waiting for the owner. Not a roadmap, not a
queue, and not governed by `ID_REGISTRY.md`: nothing in this folder has an
ID until the owner gives it one.*

---

## Why this folder exists

`BETA_REARVIEW.md`, `BETA_ROADMAP.md` and `RELEASES.md` change in nearly
every commit, so a second session editing them in parallel collides with
the first. The roadmap therefore keeps one writer: the local session, at
the owner's direction. Sessions that work elsewhere (a cloud session
auditing the tree, say) write here instead. A new file in this folder
cannot conflict with anything.

## The protocol

1. **A cloud session writes one new file per run**, named
   `YYYY-MM-DD-<topic>.md`, and never edits an older one or any roadmap.
   Candidates inside it are labelled `C1`, `C2`, … — local to that file,
   deliberately not in any roadmap family's form, so `check_id_registry.ps1`
   has nothing to govern and nothing can be cited as if it were filed.
2. **Every candidate carries its own evidence**: file and line, the code,
   a failure scenario, a verdict (*confirmed by reading* or *plausible,
   needs a live repro*), the roadmap search that shows it is not already
   filed, and a fix direction. Nothing in this folder has been run in
   Excel; a *confirmed* verdict means the path was traced end to end by
   reading, not observed.
3. **The owner triages locally**: file a candidate as a `TER-` item (or
   whichever family fits) in `BETA_ROADMAP.md` with a real ID, fold it
   into an existing item, or reject it. Record the outcome in the
   candidate's file with one line (`C3 → TER-12`, `C5 → rejected: …`), so
   the next audit does not raise it again.
4. **A file whose every candidate has an outcome is deleted** in the same
   commit that records the last outcome. The trail survives in git and in
   the roadmap items themselves.

## Standing audits

[`AUDIT_CLASSES.md`](AUDIT_CLASSES.md) lists the classes of defect a cloud
session re-scans the whole tree for, each with its recipe and its current
count, so a class that is already clean stays clean and a class that is
not is shrinking. It is the only file here that is edited in place, and
only by the auditing session.
