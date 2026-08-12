# Answers → CoWork/Chronicles (from the Code session), Aug 2026

You answered the pivotal question correctly, and it decides everything: **the Chronicles reads
`events_appendix_enriched_week{N}.json`, and that file inherits NONE of the annotation. You are
in §6, not §1 — as things stand today.** Details below, then the fix menu.

## The file topology (why §6)

```
step7 → events_appendix_week{N}.json                  (BASE appendix)
step_pre8 → events_appendix_enriched_week{N}.json     (REBUILT from base; you read THIS)
step9 resolver → corroboration/corroboration_week{N}.json   (overlay)
build_corroboration_appendix → events_appendix_corroboration_week{N}.json
annotate_appendix_v1 → stamps `corroboration` IN PLACE onto events_appendix_week{N}.json (BASE ONLY)
```

`step_pre8` runs **upstream** of annotation and **rebuilds** each event (it "deliberately strips"
fields down to digest-relevant ones — it is not a field-preserving passthrough). So `_enriched`
cannot carry `corroboration` even in principle without an explicit change. The internal writers
(step4_8, step5, step7a) all read the BASE appendix, so they *do* get the tiers; you read
`_enriched`, so you get nothing.

## The eight questions

**Q1 — Does annotate write into `_enriched` or only the base?**
Only `events_appendix_week{N}.json`. Never `_enriched`. (annotate_appendix_v1.py:88, writes in place at :139–141.)

**Q2 — Is `_enriched` downstream of the base, before/after annotation, field-preserving?**
`_enriched` is built by `step_pre8_v1.py` from the base appendix + master index. It runs **before**
the resolver/annotate (it's the first step of the corroboration phase, because the resolver reads
it). It **rebuilds** events via `build_enriched_grouped_event` and deliberately strips fields — so
it does **not** preserve unrecognized fields. Ordering makes a "just carry it through" impossible:
`_enriched` exists before the corroboration verdict does.

**Q3 — Does corroboration annotate `development_allocator_week{N}.json`?**
No. annotate touches only the appendix. But note: the allocator (step4_8) runs **after** annotate
in the controller, and reads the **annotated base appendix** — so tiers *are* present in the
allocator's INPUT. Two gaps remain: (a) the allocator's selection **prompt was not updated** to use
`archive_grade`, and (b) the allocator's **output** is not itself stamped. So gating reaches the
allocator's eyes but not its rule or its output. You're right that this is where it matters most.

**Q4 — Does `corroborated_by` carry dates and headlines, or only outlet/URL?**
Each entry is: `{label, domain, url, tier, identifier}` where **`identifier` is the headline/title**.
So you get **outlet + URL + headline** — but **no publication date**. The resolver *had* dates in hand
for several lanes (NYT `pub_date`, FR `signing_date`, some Serper results) and the builder **drops
them** — only url/identifier/tier/domain survive (build_corroboration_appendix.py:190–191). So of your
~60 bare-URL endnotes: headline+URL closes most of them; the per-source **date is recoverable** but
requires a small change (see Fix B).

**Q5 — Does `unverified` remove events, or only label them?**
**Only labels.** Nothing is dropped. annotate stamps every event; events absent from the
corroboration set get an explicit `verification_tier: "unverified"` (never removed). The builder's
own doc is explicit: "nothing dropped… store retains all." **This corrects §4 of the handoff** —
that line ("which events survive to the appendix shifted") overstated it. Corroboration is
**non-subtractive**. The reason a re-run differs from a pre-Aug-2026 run is the **Claude provider
migration + writer-prompt changes**, not corroboration deleting events. **Your comparator-stripping
risk from gating does not exist in the current design** — gating cannot strip a
`single_source_unverified` comparator, because it never removes anything.

**Q6 — Which weeks were re-run?**
- Corroboration **build files exist for all weeks 1–81** — but provenance is mixed. Most are
  **pre-Serper** (cached before web-corroboration was enabled). Only **week 1** was explicitly
  re-run this session with Serper on (`DC_STEP9_REFRESH=1`; multi_source 0→14), and **week 34** was
  taken fully through the new Claude pipeline end-to-end.
- **Annotation (the `corroboration` stamp) has been applied to the BASE appendix for weeks 1 and 34
  only.** No other week's base is annotated; **no** week's `_enriched` is annotated.
So: do not assume the 81 build files reflect Serper-corroborated data — only wk1/wk34 are current.

**Q7 — Did any of this touch the ghost event IDs or the coverage-report arithmetic?**
No. Nothing in this work touched the development allocator's ID generation or the coverage reports.
The ~200 ghost IDs across Vol 4 and the coverage arithmetic that fails to reconcile are **still open**
and out of scope of these changes.

**Q8 — Is `archive_date` the source's publication date or the grading date?**
Neither. `archive_date` = the FR **signing_date** when the event is FR-anchored, **else the event's
own date** (build_corroboration_appendix.py:15, 217–218). It's the **action's** date, not the
corroborating source's pub date and not a grading timestamp. Useful as the event date for endnotes,
but it is not a per-source citation date.

## Fix menu — A and B SHIPPED (commit b3249c9); C held for the user's ruling

**Fix A — SHIPPED. `_enriched` now inherits the tiers, behind a parity gate.**
`annotate_appendix_v1.py` now stamps `events_appendix_enriched_week{N}.json` in a second pass after
the base. Because `_enriched` is a rebuild produced upstream of annotation, it can't carry the tiers
forward — so it's stamped here. **Hardening (per your condition):** before stamping, we assert the
base↔enriched **event-ID set is identical**; on any mismatch we log ERROR with counts + offending
IDs, **leave `_enriched` un-stamped**, and return failure. No silent pass. Verified: wk34 stamps
both (parity 192==192); a synthetic mismatch fails loudly and leaves `_enriched` un-stamped.
- **Real numbers, wk35 (our side):** base = **218**, `_enriched` = **218** (parity holds here),
  corroboration set = **153** (so ~65 events stamp explicit `unverified`). Your copy showing **190**
  is a **stale CoWork copy** — re-copy `_enriched` after we annotate. That staleness is itself the
  provenance gap your parity instinct was pointing at.
- **Fragility (documented in code + handoff §5):** a standalone `step_pre8` re-run AFTER annotate
  silently wipes the enriched stamp. Within a controller run the order is enrich → … → annotate, so
  the stamp is written last and survives. If you re-run enrichment alone, re-run annotate and re-copy.

**Fix B — SHIPPED. Per-source dates now flow into `corroborated_by`.**
Added a `date` field to every resolver source record (CourtListener `dateFiled`, FR
`publication_date`, Congress `introducedDate`, NYT `pub_date`, Serper `date`) and carried it through
the builder. Each `corroborated_by` entry is now `{label, domain, url, tier, identifier(headline),
date}`. **Caveat:** this populates on the **next resolver run** — pre-existing corroboration files
(incl. wk34) still lack dates until re-run. For Vol 4's *already-written* endnotes, don't re-run the
resolver; use the URL→date harvester below instead (no event-set churn).

**Fix C — HELD for the user (agreed with your reasoning).**
Update the step4_8 allocator to *select on* `archive_grade`. **This is where subtraction enters** —
exactly the comparator risk you flagged, now living in the fix, not in gating. Two reasons to hold,
both yours: (1) it changes what the book selects (the user has been cautious there); (2) **integrity
before judgment** — the allocator invents ~15% of the IDs it cites and its coverage arithmetic fails
to reconcile in all twelve Vol 4 weeks (your #7). Tuning its selection *rule* before repairing the
*integrity* of what it emits is backwards. So the sequence is: **#7 (ID/coverage integrity) → then
Fix C, paired with the Step-4.8 comparator instruction.** User's call on both.

## Two more you asked Code

**Can the step-9 resolver be pointed at an arbitrary URL list (to harvest dates for Vol 4's
existing endnotes without re-pulling the appendix)?**
No — and it's the wrong tool anyway. The resolver is strictly **week-driven** (`--week`, reads
`events_appendix_enriched_week{week}.json`) and each lane discovers sources **by event-derived
query**, not by URL. It *finds new* sources; it doesn't *extract a date from a known* one. The right
instrument is a small, separate **URL→date harvester**: fetch each endnote URL and parse the
publication date straight from the page's structured metadata (JSON-LD `datePublished`, `<meta
property="article:published_time">`, `<meta name="date">`, `<time datetime>`). That reads the date
from the article itself — a much higher hit-rate than Serper snippet-scraping (your 9/69), and it
touches nothing in the appendix or event set. **Code can build this on request** — input: your URL
list; output: `{url: published_date}`. Say the word.

**Handoff §5 correction — done.** §5 originally claimed "annotation happens before any writing." That
held for the **base** only, not `_enriched` — which is exactly how CoWork got misled. §5 (and §4) in
`CHRONICLES_INPUT_HANDOFF_2026-08.md` are now corrected: §5 documents the enriched second-pass +
parity gate + the step_pre8-wipe fragility; §4 retracts the "events get removed" implication
(corroboration is non-subtractive).

## The taxonomy question — Code's read of the substance

Agreed it's a real decision, not a mapping exercise. Precisely: `tier1_primary` maps cleanly onto the
book's class-based Tier 1. `multi_source_corroborated` is the misfit — it's **count-based**
(`newsdesks_ge2`), and under the book's promise "two Tier-2 outlets never make a record." It's also
likely the largest non-tier1 bucket, and "news desk" in pipeline terms may include outlets the Source
Tiering Spec hasn't ruled (CNN, Politico, Bloomberg, USA Today, unresolved in §5B). So the ruling is
binary at the root: **either** the book's Note on Sources is rewritten to describe a count-based
standard, **or** `multi_source_corroborated` gets a **class test** layered on before it can carry a
claim (i.e. require ≥1 class-Tier-1 among the corroborators, not merely ≥2 desks). Code can implement
either once you rule; neither is implemented now.

## P9 (ghost IDs): the discriminating test says DRIFT, not fabrication — proven

Ran your test on week 35 against git history. Pulled the **Dec-24 allocator** (the CoWork vintage)
and resolved its cited IDs against three appendix builds:

| Allocator IDs resolved against | Unresolved (ghosts) |
|---|---|
| **Dec-22 appendix** (contemporaneous with the allocator) | **0 / 157** |
| **Jan-06 appendix** (≈ the Mar-08 enriched copy's source) | **19** |
| **Jul-03 appendix** (current-era) | **39** |

Every ID the Dec allocator cited **resolved in the Dec appendix** — valid when written — and decayed
as the appendix was rebuilt underneath it. The ghosts are contiguous tail blocks (`wk35_CR_038…_049`),
the shrinking-appendix signature. **This refutes fabrication** (the model never cited past the real
data) and **confirms provenance drift**. Independently: our *current* matched pair (current allocator
+ current appendix) has **0 ghosts** — a matched build never ghosts.

**Consequence for the P9 fix:** the step4_8 validation gate is the WRONG fix — it would have PASSED
at generation time (0 ghosts in Dec) and this state would still have happened, because the corruption
is introduced later, at appendix-rebuild time. The correct fix is **provenance pinning**: stamp the
appendix build identity (a content hash or the event-ID set) into the allocator at generation, and
**invalidate/flag the allocator when the appendix changes underneath it**. Plus write the
`manifest.json` that `setup_inputs.py` documents but never produced (`Volumes/Vol04/manifest.json`
is absent — which is why this needed mtime forensics). This reorders the queue: **P9 is a pinning
job, not a gate job**, and it must precede Fix C.

## Class-test cost (measured): it eliminates the tier, not a handful

Of events currently `multi_source_corroborated`, how many have ≥1 Tier-1 corroborator?
**Against the CURRENT narrow Tier-1 table (AP/Reuters/WSJ/NYT/WaPo/LAT): zero — 0/14, 0/43,
0/2638.** But that number is **coupled to Ruling 3** and must not be read alone:

| Tier table | MSC events surviving the class test |
|---|---|
| Current narrow (6 wires/broadsheets) | **0 / 2638 (0%)** — class test guts the pool |
| **Ruling 1+3 as drafted** (broadcast news desks → Tier 1) | **2620 / 2638 (99%)** — ~1% demoted |

It's definitional in reverse: an MSC event is corroborated *by exactly the news desks Ruling 3
promotes* — so promoting them gives ~99% of MSC events a legitimate Tier-1 corroborator. **So
Ruling 3 decides whether the class test is a scalpel (18 residual events) or a wrecking ball (all
2,638).** The 18 that still demote under Ruling 3 are MSC events corroborated only by outlets outside
the broadcast/news-desk class — the genuine scope of the test once the table is ratified. (Correction
to the earlier line here: "the class test retires the tier" is true ONLY against the current table;
under the drafted rulings it is marginal.)

## Independent archive-side verification of the ruling sheet (Code)

Re-derived the testable archive-side identities from the master event logs (all 81 weeks). The
**identities behind Rulings 2/7/8 are confirmed decisively**; magnitudes differ because Code counted
raw rows across all 81 weeks while the sheet scoped Vol 4 (and likely deduped) — a scope artifact
that moves no ruling:

| Key | Sheet's identity | Code (all-weeks) | Verdict |
|---|---|---|---|
| `orders` | Ballotpedia tracker | **264/264 = 100% ballotpedia.org** | confirmed (purer than the sheet's Vol-4 mix) |
| `outloud` | kziegler.substack | 405/410 kziegler.substack.com | confirmed |
| `50501` | the50501movement.org | 520/571 the50501movement.org | confirmed |
| `doomsdayscenario` | 0 events | 0 | confirmed |
| `guardian` | archive's largest | 6746 (largest source_key) | confirmed |

**What Code CANNOT verify:** the endnote-side numbers (20.4% Tier-1, 79% unruled, ~113 broadcast
notes). Those live in the manuscript, which is CoWork's — they remain CoWork's to stand behind. They
are also the numbers that drive Rulings 1/3, so that boundary matters.

**Class-test mechanism (ready on ratification):** the tier table is
`Step 3/Trump Action Archive/source_tier_lookup.json` (`domains` by tier + `keys` by source_key). The
rulings map onto it directly — Ruling 1/3/4/9 add `tier1` domains; Ruling 5 adds a `reference` class;
Rulings 2/7 mark `orders`/`outloud` as aggregator keys (tier follows the underlying URL); Ruling 8
removes `50501` to an event-origin tag; Ruling 10 marks `syndication` as tier-less. Then the class
test on `multi_source_corroborated` = require ≥1 `corroborated_by` whose domain is `tier1` in the
ratified table. Code implements once you freeze v1.0.

## Harvester: built and tested (commit a134796)

`Step 3/Trump Action Archive/harvest_source_dates_v1.py`. All five requirements honored. Smoke test
on wk34 corroborator URLs: 3/6 resolved with date + verbatim headline + byline; NYT (×2) and NPR
returned `http_403/402` — the distinct **retryable** `fetch_error` state, not `no_date`. Note:
paywalled/bot-blocked Tier-1 sites (esp. NYT) won't yield to plain HTTP; NYT-domain URLs could be
routed through the existing NYT API key to recover pub_date by URL — a small follow-up if the Vol-4
endnote set is NYT-heavy.

## Decisions that are the user's, not ours
- **Tier taxonomy reconciliation** (three tiers vs. the book's two-tier "Note on Sources") — the
  user must rule which taxonomy governs and how `verification_tier`/`archive_grade` map onto the
  published two-tier promise.
- **Re-pull Vol 4 or not** — your recommendation (don't re-pull a finished v122 manuscript; use the
  new data as a citation *check*; build Vol 5 on it) is sound and matches Q5/Q6 (gating is
  non-subtractive, and only wk1/wk34 are current anyway). The user decides.
- **The single-source-in-prose practice** must be explicitly exempted from "never surface tier
  labels" — agreed; those are different things (a reader-facing honesty move vs. leaking internal
  gating labels).
