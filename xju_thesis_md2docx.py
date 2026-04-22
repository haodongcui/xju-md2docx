#!/usr/bin/env python3
from __future__ import annotations

import runpy
from pathlib import Path


def main() -> None:
    tool_main = Path(__file__).resolve().parent / "xju_thesis_md2docx" / "main.py"
    runpy.run_path(str(tool_main), run_name="__main__")


if __name__ == "__main__":
    main()
