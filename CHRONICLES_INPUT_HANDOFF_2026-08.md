# Handoff → The Trump Chronicles (CoWork): input changes, Aug 2026

**Purpose.** The Democracy Clock Automation repo (Step 3 + Publish) changed the artifacts
The Trump Chronicles ingests. This brief is the cross-project notice — the two projects
don't share state, so carry this into the CoWork project and reconcile each item against
whatever the Chronicles actually reads.

**First question to answer on the CoWork side:** *which* of our artifacts does the
Chronicles consume?
- the **appendix JSON** (`events_appendix_week{N}.json`), or
- our **rendered outputs** (Substack markdown / WordPress HTML / Scrivener plaintext), or
- the **step-5 narrative** file, or
- **raw Step-3 archive data** upstream of all of the above.

Your answer determines which changes below matter. If the Chronicles builds from **raw
Step-3 data**, none of the annotation/marker changes reach it automatically and it would
need its own verification pass; if it reads our appendix/rendered outputs, it inherits the
work for free but must tolerate the new fields/markers.

---

## 1. Appendix JSON now carries a `corroboration` object on EVERY event

`events_appendix_week{N}.json` is now soft-annotated in place. Two things changed:

**a) A `corroboration` object is added to every event.** Shape:

```json
"corroboration": {
  "verification_tier": "tier1_primary | multi_source_corroborated | single_source_unverified | unverified",
  "verification_basis": "native_tier1 | newsdesks_ge2 | web_search_samefact | not_in_corroboration_set | ...",
  "archive_status":     "keep_tier1 | keep_corroborated | retained_unverified | ...",
  "archive_grade":      true,          // bool — publication-worthy on its own
  "archive_date":       "…",
  "corroborated_by":    [ … ]          // independent confirmers, for citing in prose
}
```

- **Every** event is stamped. Events not in the corroboration set get an explicit
  `verification_tier: "unverified"` (with `verification_basis: "not_in_corroboration_set"`,
  `archive_grade: false`) — never a missing field. So "no `corroboration` object" no longer
  occurs; treat its absence as a pipeline error, not as "unknown."

**b) Top-level markers** on the appendix root:

```json
"corroboration_applied": true,
"corroboration_applied_at": "2026-…Z",
"corroboration_counts": { "tier1_primary": N, "multi_source_corroborated": N, ... }
```

**Consumer guidance (how we intend these to be used):**
- **Lead** on `archive_grade: true` and `tier1_primary` / `multi_source_corroborated` events.
- Treat `single_source_unverified` / `unverified` as **supporting** material, not headline claims.
- **Never surface the tier labels in prose** — they're editorial gating metadata, not reader-facing text.
- Cite `corroborated_by` when attributing.

## 2. Rendered outputs now carry an inline "single source" marker

Unverified events get a visible marker in each rendered format. If the Chronicles reads any
of these, expect the literal strings:

| Format | Marker string |
|---|---|
| Substack **markdown** | `*[single source]*` |
| WordPress **HTML** | `<em class="dc-unverified">[single source]</em>` |
| Scrivener **plaintext** | ` [single source]` (leading space) |

If the Chronicles re-parses these files, it must **tolerate or strip** these markers.

## 3. Canonical narrative file changed → `_publish`

The step-5 chain now writes, in order: `..._final.txt` (step5_v4) then `..._publish.txt`
(step5b, style-constrained) **last**. The **`_publish` file is now canonical.**

- **Old:** consumers took `step5_narrative_week{NN}_final.txt` (or a `_draft3`).
- **New:** take `step5_narrative_week{NN}_publish.txt`; fall back to `_final` only if step5b
  didn't run. **`_draft3` is no longer produced** — remove any reference to it.

The `_publish` narrative is written to **lead with verified/high-grade events** and use
unverified ones as support (matches §1's guidance).

## 4. A week re-run is not byte-identical to a pre-Aug-2026 run

**CORRECTION (post-handoff).** This section originally implied corroboration *removes* events
("which events survive to the appendix shifted"). That was wrong and is retracted. **Corroboration
is non-subtractive** — it only *labels*. Every event is stamped; events absent from the
corroboration set get an explicit `verification_tier: "unverified"`; the builder retains all
(`archive_grade` is an ingest filter flag, not a deletion). So gating **cannot** strip an event —
including single-source comparators.

What *does* make a re-run differ from a pre-Aug-2026 run is unrelated to gating:
- The **provider migration to Claude** (Steps 3–9) changes LLM-driven allocator/narrative output.
- The **writer-prompt changes** (verified-first) change emphasis and ordering.
- The clock **scoring math is unchanged**; any clock delta is downstream of the above, not of a
  math change.

So a re-run is **not byte-identical**, but the event *set* is not thinned by corroboration. If the
Chronicles cached rendered/narrative output, re-pull; if it cached only the event *list*,
corroboration alone did not change its membership.

## 5. Ordering guarantee: annotation happens BEFORE any writing

The controller runs the corroboration + annotation phase (step_pre8 → step-9 resolver →
build_corroboration → annotate) **before** the writing steps. So by the time any narrative or
appendix artifact exists, §1's `corroboration` object is present. Best-effort fallback: if step 9
fails, the appendix is left un-annotated and a loud warning is logged (the book is never halted)
— in that case `corroboration_applied` is absent, which a consumer can test for.

**CORRECTION (post-handoff).** This ordering guarantee originally held for the **base** appendix
(`events_appendix_week{N}.json`) only. The **enriched** appendix
(`events_appendix_enriched_week{N}.json`) is a *rebuild* that step_pre8 produces **upstream** of
annotation, so it did NOT carry the `corroboration` object. That gap is now closed: `annotate`
stamps the enriched appendix too, in a **second pass, behind a base↔enriched event-ID parity gate**
that fails loudly and leaves the enriched file un-stamped on any mismatch. Two consequences for a
consumer of the enriched file: (a) a missing `corroboration` object there means annotation did not
complete (pipeline error), never "unknown"; (b) a **standalone step_pre8 re-run after annotate
silently wipes the enriched stamp** — re-run annotate afterward, and re-copy the enriched file.

---

## Action items for the CoWork side
1. Confirm which artifact the Chronicles ingests (see top).
2. If it reads the **appendix JSON** — start honoring `corroboration.archive_grade` /
   `verification_tier` for lead-vs-support selection; tolerate the new fields.
3. If it reads **rendered outputs** — tolerate/strip the §2 markers.
4. If it reads a **narrative file by name** — switch to `_publish`, drop `_draft3`.
5. Re-pull per-week event sets (§4); don't trust pre-Aug-2026 caches.
6. If it builds from **raw Step-3 data** (bypassing our appendix) — it does **not** inherit
   any of the above; decide whether to point it at our annotated appendix instead, or
   replicate the verification pass on the Chronicles side.
