#!/usr/bin/env python3
# build_cover_images_v1.py — Democracy Clock
# Build Cover Images (v1)
#
# CURRENT SCOPE (INTENTIONALLY LIMITED)
# -----------------------------------
# Generates a cover image *set* per run (unless constrained by --only/--skip):
#   1) Democracy Clock (main book) front cover background — portrait 6x9
#   2) Democracy Clock Events front cover background — portrait 7x10
#   3) Democracy Clock eBook front cover background — portrait, square-safe
#
# The cover image must contain NO text; typography is applied later in Canva.

from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path
from typing import Optional, cast
from typing_extensions import Literal


import logging
import traceback

def setup_logger_local(name: str, level: str = "INFO") -> logging.Logger:
    lvl = getattr(logging, level.upper(), logging.INFO)
    logger = logging.getLogger(name)
    logger.setLevel(lvl)

    if not any(isinstance(h, logging.StreamHandler) for h in logger.handlers):
        h = logging.StreamHandler(stream=sys.stdout)
        h.setLevel(lvl)
        h.setFormatter(logging.Formatter("%(asctime)s | %(levelname)s | %(name)s | %(message)s"))
        logger.addHandler(h)

    logger.propagate = False
    return logger

# Prompt module (prompt-first rule)
from build_cover_images_prompts_v1 import (
    get_democracy_clock_front_cover_prompt_portrait_6x9,
    get_democracy_clock_events_front_cover_prompt_portrait_7x10,
    get_democracy_clock_ebook_front_cover_prompt,
)


# NOTE: must match OpenAI Images API Literal size values exactly
DEFAULT_SIZE = "1024x1536"  # portrait-friendly, Pylance-safe literal
DEFAULT_IMAGE_MODEL = "gpt-image-1"

CoverType = Literal["clock", "events", "ebook"]
ALL_COVERS: tuple[CoverType, ...] = ("clock", "events", "ebook")
DEFAULT_COVER: CoverType = "clock"


def default_output_dir() -> Path:
    """Return the default output directory for cover images.

    We deliberately separate book cover assets from weekly post assets.
    """
    publish_dir = Path(__file__).resolve().parent
    return publish_dir / "Output" / "Covers" / "Democracy_Clock" / "v1"


def build_output_paths(out_dir: Path, *, cover: CoverType) -> tuple[Path, Path]:
    """Return (image_path, prompt_path) for the requested cover artifact."""
    if cover == "clock":
        img_name = "democracy_clock_front_6x9_portrait.png"
        prompt_name = "democracy_clock_front_6x9_portrait.prompt.txt"
    elif cover == "events":
        img_name = "democracy_clock_events_front_7x10_portrait.png"
        prompt_name = "democracy_clock_events_front_7x10_portrait.prompt.txt"
    elif cover == "ebook":
        img_name = "democracy_clock_ebook_front_portrait.png"
        prompt_name = "democracy_clock_ebook_front_portrait.prompt.txt"
    else:
        raise ValueError(f"Unknown cover type: {cover}")

    img_path = out_dir / img_name
    prompt_path = out_dir / prompt_name
    return img_path, prompt_path


ImageSizeLiteral = Literal[
    "auto",
    "1024x1024",
    "1536x1024",
    "1024x1536",
    "256x256",
    "512x512",
    "1792x1024",
    "1024x1792",
]

# Self-contained OpenAI image generation helper (inlined)
def generate_image_via_openai(prompt: str, *, model: str, size: str, logger: logging.Logger) -> bytes:
    """Generate an image via OpenAI Images API and return raw PNG bytes.

    This is an inlined, self‑contained adaptation of the Step 6 logic.
    No external modules are assumed.
    """
    try:
        from openai import OpenAI
    except ImportError as e:
        raise RuntimeError("openai package not installed in this environment") from e

    client = OpenAI()

    logger.info("Submitting image generation request to OpenAI")

    # Pylance requires size to be a known Literal, not an arbitrary str
    size_literal = cast(ImageSizeLiteral, size)

    result = client.images.generate(
        model=model,
        prompt=prompt,
        size=size_literal,
    )

    if not result or not result.data or not result.data[0].b64_json:
        raise RuntimeError("Image API returned no image data")

    import base64

    return base64.b64decode(result.data[0].b64_json)


def get_cover_spec(*, cover: CoverType) -> tuple[str, str, str]:
    """Return (label, prompt, default_size) for the requested cover type."""
    if cover == "clock":
        label = "Democracy Clock front cover (portrait 6x9)"
        prompt = get_democracy_clock_front_cover_prompt_portrait_6x9().strip() + "\n"
        default_size = "1024x1536"  # portrait-friendly
        return label, prompt, default_size

    if cover == "events":
        label = "Democracy Clock Events front cover (portrait 7x10)"
        prompt = get_democracy_clock_events_front_cover_prompt_portrait_7x10().strip() + "\n"
        default_size = "1024x1536"  # still portrait; Canva/print crop handled later
        return label, prompt, default_size

    if cover == "ebook":
        label = "Democracy Clock eBook front cover (portrait, square-safe)"
        prompt = get_democracy_clock_ebook_front_cover_prompt().strip() + "\n"
        default_size = "1024x1536"  # portrait-friendly; prompt is square-safe
        return label, prompt, default_size

    raise ValueError(f"Unknown cover type: {cover}")


def run(
    *,
    cover: CoverType = DEFAULT_COVER,
    level: str = "INFO",
    out_dir: Optional[Path] = None,
    force: bool = False,
    dry_run: bool = False,
    image_model: str = DEFAULT_IMAGE_MODEL,
    size: str = "",
) -> int:
    logger = setup_logger_local("build_cover_images_v1", level)

    logger.info("=== BUILD COVER IMAGES v1: START ===")

    out_dir = out_dir or default_output_dir()
    out_dir.mkdir(parents=True, exist_ok=True)

    label, prompt, default_size_for_cover = get_cover_spec(cover=cover)
    if not size:
        size = default_size_for_cover

    img_path, prompt_path = build_output_paths(out_dir, cover=cover)

    logger.info("Cover image v1: %s", label)
    logger.info("Output dir: %s", out_dir)
    logger.info("Image path: %s", img_path)

    # Always write the prompt sidecar (useful for provenance)
    if (not prompt_path.exists()) or force:
        prompt_path.write_text(prompt, encoding="utf-8")
        logger.info("Wrote prompt sidecar: %s", prompt_path)
    else:
        logger.info("Prompt sidecar exists; skipping (use --force to overwrite): %s", prompt_path)

    logger.info("Prompt loaded (%d chars). Preview: %s", len(prompt), prompt[:200].replace("\n", " "))

    # NOTE: Existence of one cover image is not a fatal condition.
    # This script is designed to skip existing artifacts and proceed.
    if img_path.exists() and not force:
        logger.info("Image already exists; skipping (use --force to overwrite): %s", img_path)
        logger.info("Continuing without regenerating existing image.")
        return 0

    if dry_run:
        logger.info("Dry-run mode enabled; prompt written, no image API call made.")
        logger.info("=== BUILD COVER IMAGES v1: DRY-RUN COMPLETE (no API call) ===")
        return 0

    if not os.getenv("OPENAI_API_KEY"):
        msg = "OPENAI_API_KEY not set in environment; cannot call image API."
        logger.error(msg)
        print(msg, file=sys.stderr)
        logger.info("=== BUILD COVER IMAGES v1: FAILED (missing OPENAI_API_KEY) ===")
        return 2

    try:
        logger.info("Calling image API (cover=%s, model=%s, size=%s)...", cover, image_model, size)
        img_bytes = generate_image_via_openai(
            prompt,
            model=image_model,
            size=size,
            logger=logger,
        )
    except Exception as e:
        logger.error("Image generation failed: %s", e)
        logger.debug("Traceback:\n%s", traceback.format_exc())
        print(f"Image generation failed: {e}", file=sys.stderr)
        return 40

    try:
        img_path.write_bytes(img_bytes)
    except Exception as e:
        logger.error("Failed to write image file at %s: %s", img_path, e)
        return 41

    logger.info("Cover image generated: %s (%d bytes)", img_path, len(img_bytes))
    logger.info("=== BUILD COVER IMAGES v1: COMPLETE ===")
    return 0


def main(argv: Optional[list[str]] = None) -> int:
    ap = argparse.ArgumentParser(description="Democracy Clock — Build Cover Images (v1)")
    ap.add_argument("--level", default="INFO", help="Logging level (e.g., INFO or DEBUG)")
    ap.add_argument("--out-dir", default="", help="Optional output directory (defaults to Publish/Output/Covers/...) ")
    ap.add_argument("--force", action="store_true", help="Overwrite existing prompt/image files")
    ap.add_argument("--dry-run", action="store_true", help="Write prompt sidecar only; do not call the image API")
    ap.add_argument("--image-model", default=DEFAULT_IMAGE_MODEL, help="Image model identifier (default: gpt-image-1)")
    ap.add_argument(
        "--size",
        default="",
        choices=[
            "1024x1024",
            "1024x1536",
            "1536x1024",
            "1792x1024",
            "1024x1792",
        ],
        help="Image size parameter (OpenAI Images API literal sizes); overrides cover default",
    )
    ap.add_argument(
        "--only",
        default="",
        help="Comma-separated list of covers to build exclusively (e.g. clock,events,ebook)",
    )
    ap.add_argument(
        "--skip",
        default="",
        help="Comma-separated list of covers to skip (e.g. events,ebook)",
    )
    args = ap.parse_args(argv)

    requested: list[CoverType] = list(ALL_COVERS)

    if args.only:
        requested = [cast(CoverType, c.strip()) for c in args.only.split(",") if c.strip()]
    elif args.skip:
        skips = {c.strip() for c in args.skip.split(",") if c.strip()}
        requested = [c for c in requested if c not in skips]

    if not requested:
        print("No covers selected after applying --only/--skip", file=sys.stderr)
        return 3

    out_dir = Path(args.out_dir).expanduser().resolve() if args.out_dir.strip() else None

    final_rc = 0
    for cover in requested:
        rc = run(
            cover=cover,
            level=args.level,
            out_dir=out_dir,
            force=args.force,
            dry_run=args.dry_run,
            image_model=args.image_model,
            size=args.size,
        )
        if rc != 0:
            final_rc = rc

    if final_rc == 0:
        print("build_cover_images_v1: OK")
    else:
        print(f"build_cover_images_v1: FAILED (rc={final_rc})", file=sys.stderr)
    return final_rc


if __name__ == "__main__":  # pragma: no cover
    sys.exit(main())