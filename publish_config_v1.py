#!/usr/bin/env python3
# publish_config_v1.py — config for Democracy Clock publish/post automation

from __future__ import annotations
from pathlib import Path

# Root of the Publish workspace (this file lives in /Programs)
PUBLISH_ROOT: Path = Path(__file__).resolve().parent.parent

# Where publish automation scripts live
PUBLISH_PROGRAMS_DIR: Path = PUBLISH_ROOT / "Programs"

# Read-only access to Step 3 outputs
STEP3_ROOT: Path = PUBLISH_ROOT.parent / "Step 3"
STEP3_WEEKS_DIR: Path = STEP3_ROOT / "Weeks"

# Outputs we generate
SUBSTACK_OUTPUT_DIR: Path = PUBLISH_ROOT / "Output" / "Substack"
SCRIVENER_OUTPUT_DIR: Path = PUBLISH_ROOT / "Output" / "Scrivener"
PUBLISH_LOGS_DIR: Path = PUBLISH_ROOT / "Logs"
PUBLISH_IMAGES_DIR: Path = PUBLISH_ROOT / "Images"   # ← NEW
# Timezone for publish timestamps
TZ_DEFAULT: str = "Australia/Brisbane"