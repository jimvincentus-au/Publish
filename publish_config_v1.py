#!/usr/bin/env python3
# publish_config_v1.py — config for Democracy Clock publish/post automation

from __future__ import annotations
from pathlib import Path

"""
Directory layout reminder (from the external drive root):

/Volumes/PRINTIFY24/
    Democracy Clock Automation/
        Step 2/
        Step 3/
            Weeks/
        Publish/
            publish_config_v1.py
            build_publish_week_v1.py
            Output/
            Logs/
            Images/        <- we create/use this
"""

# This file lives in: .../Democracy Clock Automation/Publish/publish_config_v1.py
# So:
#   PUBLISH_ROOT         = .../Democracy Clock Automation/Publish
#   PROJECT_ROOT         = .../Democracy Clock Automation
#   STEP3_ROOT           = .../Democracy Clock Automation/Step 3
#   STEP3_WEEKS_DIR      = .../Democracy Clock Automation/Step 3/Weeks

# Root of the Publish workspace (folder that contains this config file)
PUBLISH_ROOT: Path = Path(__file__).resolve().parent

# One level up: the Democracy Clock Automation project root
PROJECT_ROOT: Path = PUBLISH_ROOT.parent

# Where publish automation scripts live
PUBLISH_PROGRAMS_DIR: Path = PUBLISH_ROOT / "Programs"

# Read-only access to Step 3 outputs
STEP3_ROOT: Path = PROJECT_ROOT / "Step 3"
STEP3_WEEKS_DIR: Path = STEP3_ROOT / "Weeks"

# Outputs we generate
SUBSTACK_OUTPUT_DIR: Path = PUBLISH_ROOT / "Output" / "Substack"
SCRIVENER_OUTPUT_DIR: Path = PUBLISH_ROOT / "Output" / "Scrivener"

# Where we log things
PUBLISH_LOGS_DIR: Path = PUBLISH_ROOT / "Logs"

# Where we stage images for Substack/Scrivener
PUBLISH_IMAGES_DIR: Path = PUBLISH_ROOT / "Images"

# Timezone for publish timestamps
TZ_DEFAULT: str = "Australia/Brisbane"