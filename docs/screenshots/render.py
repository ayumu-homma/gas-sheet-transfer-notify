#!/usr/bin/env python3
"""説明用 HTML を PNG にする。Google の実画面は撮影しない。"""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path

CHROME = "google-chrome"
ROOT = Path(__file__).resolve().parent
MOCKS = ROOT / "mocks"

JOBS = [
    ("intake-before.html", "intake-before.png", 1100, 560),
    ("done-after.html", "done-after.png", 1100, 480),
    ("email-example.html", "email-example.png", 1100, 700),
    ("trigger-settings.html", "trigger-settings.png", 1100, 520),
]


def render(html: Path, png: Path, w: int, h: int) -> None:
    cmd = [
        CHROME,
        "--headless=new",
        "--disable-gpu",
        "--no-sandbox",
        "--hide-scrollbars",
        "--force-device-scale-factor=1",
        f"--window-size={w},{h}",
        f"--screenshot={png}",
        html.resolve().as_uri(),
    ]
    subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    print(f"wrote {png}")


def main() -> int:
    for name, out, w, h in JOBS:
        render(MOCKS / name, ROOT / out, w, h)
    return 0


if __name__ == "__main__":
    sys.exit(main())
