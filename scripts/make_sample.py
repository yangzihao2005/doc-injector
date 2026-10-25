"""Regenerate templates/sample_dsp.docx and examples/ demo assets."""

from __future__ import annotations

import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from src.demo_assets import build_sample_payload, build_sample_template, draw_waveform  # noqa: E402


def main() -> None:
    template = build_sample_template(REPO_ROOT / "templates" / "sample_dsp.docx")
    image = draw_waveform(REPO_ROOT / "examples" / "waveform.png")
    payload = build_sample_payload(image, REPO_ROOT / "examples" / "sample_payload.json")
    print(f"template: {template}")
    print(f"image:    {image}")
    print(f"payload:  {payload}")


if __name__ == "__main__":
    main()
