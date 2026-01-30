# build_cover_images_prompts_v1.py
#
# Democracy Clock – Book Cover Image Prompts (v1)
#
# PROMPT-FIRST MODULE
# -------------------
# Defines canonical prompts as string constants plus lightweight helper(s).
#
# CURRENT SCOPE (INTENTIONALLY LIMITED)
# -----------------------------------
# v1 supports THREE image types:
#   1) Democracy Clock (main book) front-cover BACKGROUND image (portrait 6x9)
#   2) Democracy Clock Events front-cover BACKGROUND image (portrait 7x10)
#   3) Democracy Clock / Democracy Clock Events eBook front-cover BACKGROUND image (portrait, square-safe)
#
# IMPORTANT
# ---------
# - This prompt is for the IMAGE ONLY (background artwork).
# - The generated image must contain NO typography, NO logos, NO signage, NO words.
# - Title/subtitle/author text will be added later in Canva.

from __future__ import annotations

from textwrap import dedent


COVER_PALETTE_TONE_CLAUSE = dedent(
    """
    PALETTE + TONE (CANONICAL)
    - archival artifact aesthetic; institutional, documentary, historically grounded
    - richer, disciplined tones (not pale tan): deep slate blue, smoky steel, warm parchment, restrained amber highlights
    - low saturation overall but with deeper midtones; moderate contrast; avoid both crushed blacks and washed-out beige
    - soft diffuse light with gentle directional falloff; no theatrical spotlighting; no neon
    - subtle archival grain / paper texture / film noise (NOT painterly brushwork; NOT flat vector illustration)
    - mood: serious, authoritative, evidentiary; calm gravity rather than melodrama
    """
).strip()


COVER_GLOBAL_CONSTRAINTS = dedent(
    """
    GLOBAL CONSTRAINTS
    - Aim for documentary realism / archival-photographic feel; avoid cartoonish, cute, or graphic-illustration styles
    - Avoid flat iconography, vector shapes, or simplified clip-art geometry
    - NO text, NO letters, NO numbers, NO words, NO logos, NO seals, NO signage, NO watermarks
    - Avoid plaques, banners, etched stone, carved inscriptions, newspapers with readable headlines
    - NO modern campaign graphics; NO partisan iconography
    - Do not depict identifiable real people; crowds may be implied/silhouetted
    - Avoid gore, violence, weapons, or explicit confrontation
    - Composition must read at thumbnail size
    - Maintain negative space for later Canva title/subtitle overlay
    """
).strip()


COVER_PROMPT_DEMOCRACY_CLOCK_FRONT_PORTRAIT_6X9 = dedent(
    f"""
    Democracy Clock — book cover background artwork (front cover only).
    ORIENTATION: portrait, designed for a 6x9 book cover.

    {COVER_PALETTE_TONE_CLAUSE}

    SUBJECT + MOTIF
    - dominant analog clock motif with Roman numerals; hands near midnight (sense of proximity, not spectacle)
    - archival-record overlays: faint docket lines, filings, timestamp textures, aged paper (NON-readable, abstract)
    - civic process is implied, not illustrated: soft silhouettes or a quiet queue as atmospheric presence (no literal booths)
    - optional, extremely faint institutional backdrop: courthouse columns / Capitol-like architecture, out of focus and secondary
    - overall emphasis: documentary record, institutional accountability, time pressure

    COMPOSITION
    - Preserve a clean title field: generous negative space in the upper 30% (smooth gradient, low texture).
    - Prefer asymmetry and editorial restraint: clock may be partially cropped or off-center to avoid “poster symbolism.”
    - Keep primary visual mass in the mid-to-lower half; avoid competing focal points.
    - Depth via atmospheric haze and layered archival textures; keep textures subtle in the title area.
    - Edges calm and low-frequency; avoid noisy borders.
    - Thumbnail test: the clock motif must still read clearly at small sizes.

    {COVER_GLOBAL_CONSTRAINTS}
    """
).strip()


COVER_PROMPT_DEMOCRACY_CLOCK_EVENTS_FRONT_PORTRAIT_7X10 = dedent(
    f"""
    Democracy Clock Events — book cover background artwork (front cover only).
    ORIENTATION: portrait, designed for a 7x10 book cover.

    {COVER_PALETTE_TONE_CLAUSE}

    SUBJECT + MOTIF
    - archival evidence stacks: binders, folders, case files, ledgers, docket stacks (photographic, documentary feel)
    - visual sense of accumulation and chronology; repetition with variation
    - subtle temporal traces ONLY as paperwork artifacts: ruled lines, date-stamp textures, pagination edges (NON-readable, abstract)
    - institutional environment implied: archive table, shelving, file room depth-of-field (soft, realistic)
    - no heroic symbolism; no staged “still life” perfection
    - emphasis: evidentiary weight, volume, and recordkeeping

    EXPLICIT EXCLUSIONS (IMPORTANT)
    - NO clocks, NO clock faces, NO timepieces, NO hour/minute hands
    - NO allegorical symbols (scales of justice, flags as focal symbol, eagles, torches, etc.)
    - do not turn the records into a “poster”; keep it like an archival photograph

    COMPOSITION
    - Strong vertical composition suitable for 7x10 format.
    - Clean title field: generous negative space in the upper 30% (smooth, low texture).
    - Visual weight concentrated mid-to-lower frame; edges calm and low-contrast.
    - Convey scale through layered depth and repetition, not clutter.
    - Realistic lighting and depth-of-field; avoid graphic illustration cues.
    - Thumbnail test: the stacked-record motif must read as “archives / evidence” at small sizes.

    {COVER_GLOBAL_CONSTRAINTS}
    """
).strip()


COVER_PROMPT_DEMOCRACY_CLOCK_EBOOK_FRONT = dedent(
    f"""
    Democracy Clock — eBook cover background artwork (front cover only).
    ORIENTATION: portrait, optimized for eBook thumbnails and square-safe cropping.

    {COVER_PALETTE_TONE_CLAUSE}

    SUBJECT + MOTIF
    - archival artifact aesthetic tuned for small thumbnails: simple, authoritative, documentary
    - a single cohesive motif (choose one): partial clock rim OR institutional architectural silhouette OR archival paper layers
    - visual continuity cue with the print covers: implied circular falloff, partial clock rim shadow, or curved edge suggested through light and texture only (no full clock face)
    - cooler slate-blue and smoky steel midtones as the dominant base, with restrained warm parchment highlights layered secondarily; deeper midtone density to visually match the print covers
    - do not combine multiple motifs; keep it singular and thumbnail-strong
    - faint archival textures: paper grain, filing lines, ledger patterns (NON-readable, abstract)
    - avoid flat geometric iconography; avoid “logo-like” minimalism

    COMPOSITION
    - Central, stable composition with strong visual coherence.
    - Guarantee square-safe cropping: key motif stays within the central 60%.
    - Maintain generous negative space in upper and central areas for Canva text overlay.
    - Keep detail simple and bold enough to survive very small sizes.
    - Avoid heavy vignettes or edge-darkening that harms thumbnail clarity.
    - Tonal balance must visually align with the 6x9 and 7x10 print covers when viewed side-by-side; avoid warmer overall temperature or flatter contrast than the print editions
    - Documentary/archival-photographic feel; not cartoonish; not vector.

    {COVER_GLOBAL_CONSTRAINTS}
    """
).strip()


def get_democracy_clock_front_cover_prompt_portrait_6x9() -> str:
    """Return the canonical prompt for the Democracy Clock (main book) front cover background image."""
    return COVER_PROMPT_DEMOCRACY_CLOCK_FRONT_PORTRAIT_6X9


def get_democracy_clock_events_front_cover_prompt_portrait_7x10() -> str:
    """Return the canonical prompt for the Democracy Clock Events front cover background image."""
    return COVER_PROMPT_DEMOCRACY_CLOCK_EVENTS_FRONT_PORTRAIT_7X10


def get_democracy_clock_ebook_front_cover_prompt() -> str:
    """Return the canonical prompt for the Democracy Clock eBook front cover background image."""
    return COVER_PROMPT_DEMOCRACY_CLOCK_EBOOK_FRONT


if __name__ == "__main__":
    print("=== Democracy Clock (6x9) ===")
    print(get_democracy_clock_front_cover_prompt_portrait_6x9())
    print()
    print("=== Democracy Clock Events (7x10) ===")
    print(get_democracy_clock_events_front_cover_prompt_portrait_7x10())
    print()
    print("=== Democracy Clock eBook ===")
    print(get_democracy_clock_ebook_front_cover_prompt())
