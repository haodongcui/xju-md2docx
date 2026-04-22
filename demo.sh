#!/usr/bin/env bash
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

python3 "$ROOT/xju_thesis_md2docx.py" \
  "$ROOT/example/thesis-demo.md" \
  "$ROOT/example/thesis-demo.generated.docx"
 
echo "Generated: $ROOT/example/thesis-demo.generated.docx"
