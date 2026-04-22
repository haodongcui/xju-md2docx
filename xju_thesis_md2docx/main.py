#!/usr/bin/env python3
from __future__ import annotations

import argparse
import datetime as dt
import html
import json
import re
import subprocess
import sys
import zipfile
from dataclasses import dataclass
from pathlib import Path
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

try:
    from PIL import Image
except ImportError:  # pragma: no cover - optional dependency in some environments
    Image = None

TOOL_ROOT = Path(__file__).resolve().parent
DEFAULT_TEMPLATE_PATH = TOOL_ROOT / "resources" / "xju-template.docx"
DEFAULT_COVER_ASSETS_DIR = TOOL_ROOT / "resources"
DEFAULT_LOCAL_COVER_ASSETS_REL = Path("img/cover-assets")

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
CP_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
DC_NS = "http://purl.org/dc/elements/1.1/"
DCTERMS_NS = "http://purl.org/dc/terms/"
XSI_NS = "http://www.w3.org/2001/XMLSchema-instance"
VT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"

INLINE_MATH_PATTERN = re.compile(r"(?<!\\)\$(?!\$)(.+?)(?<!\\)\$(?!\$)")
INLINE_CITATION_PATTERN = re.compile(r"\[(\d+(?:\s*(?:[-,，]\s*\d+)*)+)\]")
IMAGE_PATTERN = re.compile(r"^!\[(?P<alt>[^\]]*)\]\((?P<target>[^)]+)\)$")
FIGURE_ROW_START_PATTERN = re.compile(r"^:::\s*figure-row\s*$")
FIGURE_ROW_END_PATTERN = re.compile(r"^:::\s*$")
CAPTION_PATTERN = re.compile(r"^[图表]\s*\d+(?:[-.]\d+)*(?:\([a-zA-Z]\))?\s+")
WORD_MATH_DIR = TOOL_ROOT / "world-math"
WORD_MATH_SCRIPT = WORD_MATH_DIR / "convert.js"
WORD_MATH_REQUIRED_MODULES = (
    WORD_MATH_DIR / "node_modules" / "temml",
    WORD_MATH_DIR / "node_modules" / "@hungknguyen" / "mathml2omml",
)
OMML_TEXT_PATTERN = re.compile(r"(<(?:m|w):t\b[^>]*>)(.*?)(</(?:m|w):t>)", re.DOTALL)
COVER_EMBLEM_NAME = "xju-emblem.jpeg"
COVER_WORDMARK_NAME = "xju-wordmark.png"

IMAGE_CONTENT_TYPES = {
    "png": "image/png",
    "jpg": "image/jpeg",
    "jpeg": "image/jpeg",
    "gif": "image/gif",
    "bmp": "image/bmp",
}
IMAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
EMU_PER_INCH = 914400
DEFAULT_DPI = 96
MAX_IMAGE_WIDTH_IN = 5.8
MAX_IMAGE_HEIGHT_IN = 8.0
FIGURE_ROW_MAX_WIDTH_IN = 2.75
FIGURE_ROW_MAX_HEIGHT_IN = 3.2


def xml_text(text: str) -> str:
    if text == "":
        return '<w:t xml:space="preserve"></w:t>'
    text = escape(text)
    if text.startswith(" ") or text.endswith(" ") or "  " in text:
        return f'<w:t xml:space="preserve">{text}</w:t>'
    return f"<w:t>{text}</w:t>"


def run_text_xml(
    text: str,
    *,
    bold: bool = False,
    italic: bool = False,
    font_ascii: str | None = None,
    font_hansi: str | None = None,
    font_eastasia: str | None = None,
    size: int | None = None,
) -> str:
    rpr: list[str] = []
    fonts: list[str] = []
    if font_ascii:
        fonts.append(f'w:ascii="{escape(font_ascii)}"')
    if font_hansi:
        fonts.append(f'w:hAnsi="{escape(font_hansi)}"')
    if font_eastasia:
        fonts.append(f'w:eastAsia="{escape(font_eastasia)}"')
    if fonts:
        rpr.append(f"<w:rFonts {' '.join(fonts)}/>")
    if bold:
        rpr.append("<w:b/><w:bCs/>")
    if italic:
        rpr.append("<w:i/><w:iCs/>")
    if size is not None:
        rpr.append(f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>')
    rpr_xml = f"<w:rPr>{''.join(rpr)}</w:rPr>" if rpr else ""
    return f"<w:r>{rpr_xml}{xml_text(text)}</w:r>"


def break_run_xml() -> str:
    return "<w:r><w:br/></w:r>"


def field_char_run_xml(kind: str, *, dirty: bool = False) -> str:
    dirty_attr = ' w:dirty="true"' if dirty else ""
    return f'<w:r><w:fldChar w:fldCharType="{kind}"{dirty_attr}/></w:r>'


def instr_text_run_xml(text: str) -> str:
    return f'<w:r><w:instrText xml:space="preserve">{escape(text)}</w:instrText></w:r>'


def spacing_xml(
    *,
    line: int | None = None,
    before: int | None = None,
    after: int | None = None,
    before_lines: int | None = None,
    after_lines: int | None = None,
    line_rule: str = "auto",
) -> str:
    attrs: list[str] = []
    if before_lines is not None:
        attrs.append(f'w:beforeLines="{before_lines}"')
    if before is not None:
        attrs.append(f'w:before="{before}"')
    if after_lines is not None:
        attrs.append(f'w:afterLines="{after_lines}"')
    if after is not None:
        attrs.append(f'w:after="{after}"')
    if line is not None:
        attrs.append(f'w:line="{line}"')
        attrs.append(f'w:lineRule="{line_rule}"')
    if not attrs:
        return ""
    return f"<w:spacing {' '.join(attrs)}/>"


def indent_xml(
    *,
    first_line_chars: int | None = None,
    first_line: int | None = None,
    left: int | None = None,
    hanging: int | None = None,
) -> str:
    attrs: list[str] = []
    if first_line_chars is not None:
        attrs.append(f'w:firstLineChars="{first_line_chars}"')
    if first_line is not None:
        attrs.append(f'w:firstLine="{first_line}"')
    if left is not None:
        attrs.append(f'w:left="{left}"')
    if hanging is not None:
        attrs.append(f'w:hanging="{hanging}"')
    if not attrs:
        return ""
    return f"<w:ind {' '.join(attrs)}/>"


def text_runs(text: str, run_kwargs: dict[str, object] | None = None, preserve_breaks: bool = False) -> list[str]:
    run_kwargs = run_kwargs or {}
    if preserve_breaks and "\n" in text:
        parts = text.split("\n")
        runs: list[str] = []
        for idx, part in enumerate(parts):
            runs.append(run_text_xml(part, **run_kwargs))
            if idx != len(parts) - 1:
                runs.append(break_run_xml())
        return runs
    return [run_text_xml(text, **run_kwargs)]


def split_inline_code(text: str) -> list[tuple[str, str]]:
    parts: list[tuple[str, str]] = []
    i = 0
    last = 0
    while i < len(text):
        if text[i] != "`":
            i += 1
            continue

        tick_count = 1
        while i + tick_count < len(text) and text[i + tick_count] == "`":
            tick_count += 1

        marker = "`" * tick_count
        closing = text.find(marker, i + tick_count)
        if closing == -1:
            i += tick_count
            continue

        if i > last:
            parts.append(("text", text[last:i]))
        parts.append(("code", text[i + tick_count : closing]))
        i = closing + tick_count
        last = i

    if last < len(text):
        parts.append(("text", text[last:]))

    return parts if parts else [("text", text)]


def split_inline_math(text: str) -> list[tuple[str, str]]:
    parts: list[tuple[str, str]] = []
    last = 0
    for match in INLINE_MATH_PATTERN.finditer(text):
        if match.start() > last:
            parts.append(("text", text[last:match.start()]))
        latex = match.group(1).strip()
        if latex:
            parts.append(("math", latex))
        else:
            parts.append(("text", "$$"))
        last = match.end()
    if last < len(text):
        parts.append(("text", text[last:]))
    return [(kind, value.replace(r"\$", "$")) for kind, value in parts if value]


def inline_code_run_xml(text: str, *, size: int | None = None) -> str:
    return run_text_xml(
        text,
        font_ascii="Courier New",
        font_hansi="Courier New",
        font_eastasia="等线",
        size=size,
    )


def reference_bookmark_name(ref_id: str) -> str:
    return f"ref_{ref_id}"


def reference_bookmark_id(ref_id: str) -> int:
    return 1000 + int(ref_id)


def extract_reference_anchors(text: str) -> dict[str, str]:
    anchors: dict[str, str] = {}
    for ref_id in re.findall(r"^\[(\d+)\]\s", text, re.MULTILINE):
        anchors.setdefault(ref_id, reference_bookmark_name(ref_id))
    return anchors


def hyperlink_run_xml(text: str, anchor: str, *, run_kwargs: dict[str, object] | None = None) -> str:
    run_kwargs = dict(run_kwargs or {})
    run_kwargs.pop("bold", None)
    run_kwargs.pop("italic", None)
    rpr: list[str] = []
    fonts: list[str] = []
    if font_ascii := run_kwargs.get("font_ascii"):
        fonts.append(f'w:ascii="{escape(str(font_ascii))}"')
    if font_hansi := run_kwargs.get("font_hansi"):
        fonts.append(f'w:hAnsi="{escape(str(font_hansi))}"')
    if font_eastasia := run_kwargs.get("font_eastasia"):
        fonts.append(f'w:eastAsia="{escape(str(font_eastasia))}"')
    if fonts:
        rpr.append(f"<w:rFonts {' '.join(fonts)}/>")
    if size := run_kwargs.get("size"):
        rpr.append(f'<w:sz w:val="{int(size)}"/><w:szCs w:val="{int(size)}"/>')
    rpr_xml = f"<w:rPr>{''.join(rpr)}</w:rPr>" if rpr else ""
    return f'<w:hyperlink w:anchor="{escape(anchor)}" w:history="1"><w:r>{rpr_xml}{xml_text(text)}</w:r></w:hyperlink>'


def citation_text_runs(
    text: str,
    *,
    run_kwargs: dict[str, object] | None = None,
    reference_anchors: dict[str, str] | None = None,
) -> list[str]:
    if not reference_anchors:
        return text_runs(text, run_kwargs=run_kwargs)

    runs: list[str] = []
    last = 0
    for match in INLINE_CITATION_PATTERN.finditer(text):
        if match.start() > last:
            runs.extend(text_runs(text[last:match.start()], run_kwargs=run_kwargs))

        ref_ids = re.findall(r"\d+", match.group(1))
        anchor = reference_anchors.get(ref_ids[0]) if ref_ids else None
        if anchor:
            runs.append(hyperlink_run_xml(match.group(0), anchor, run_kwargs=run_kwargs))
        else:
            runs.extend(text_runs(match.group(0), run_kwargs=run_kwargs))
        last = match.end()

    if last < len(text):
        runs.extend(text_runs(text[last:], run_kwargs=run_kwargs))
    return runs


def paragraph_with_inline_math_xml(
    text: str,
    *,
    style: str | None = None,
    align: str | None = None,
    ppr_extra: str = "",
    first_line_chars: int | None = None,
    first_line: int | None = None,
    run_kwargs: dict[str, object] | None = None,
    math_converter: "MathConverter | None" = None,
    reference_anchors: dict[str, str] | None = None,
) -> str:
    code_segments = split_inline_code(text)
    has_code = any(kind == "code" for kind, _ in code_segments)
    has_math = any(
        kind == "math"
        for segment_kind, segment_text in code_segments
        if segment_kind == "text"
        for kind, _ in split_inline_math(segment_text)
    )
    has_citation = any(
        bool(reference_anchors) and bool(INLINE_CITATION_PATTERN.search(segment_text))
        for segment_kind, segment_text in code_segments
        if segment_kind == "text"
    )
    if not has_code and not has_math and not has_citation:
        return formatted_paragraph_xml(
            text,
            style=style,
            align=align,
            ppr_extra=ppr_extra,
            first_line_chars=first_line_chars,
            first_line=first_line,
            run_kwargs=run_kwargs,
        )

    run_kwargs = run_kwargs or {}
    runs: list[str] = []
    code_size = int(run_kwargs.get("size")) if run_kwargs.get("size") else None
    for segment_kind, segment_text in code_segments:
        if segment_kind == "code":
            runs.append(inline_code_run_xml(segment_text, size=code_size))
            continue
        for kind, value in split_inline_math(segment_text):
            if kind == "text":
                runs.extend(citation_text_runs(value, run_kwargs=run_kwargs, reference_anchors=reference_anchors))
                continue
            omml = math_converter.get(value, display_mode=False) if math_converter else None
            if omml:
                runs.append(omml)
            else:
                runs.append(run_text_xml(f"${value}$", **run_kwargs))

    return paragraph_xml(
        style=style,
        align=align,
        runs=runs,
        ppr_extra=ppr_extra,
        first_line_chars=first_line_chars,
        first_line=first_line,
    )


def math_paragraph_xml(
    latex: str,
    *,
    style: str | None = None,
    align: str | None = "center",
    math_converter: "MathConverter | None" = None,
) -> str:
    if math_converter:
        omml = math_converter.get(latex, display_mode=True)
        if omml:
            return paragraph_xml(style=style, align=align, runs=[omml])
    return paragraph_xml(latex, style=style, align=align)


def collect_math_items(text: str) -> list[tuple[str, bool]]:
    items: list[tuple[str, bool]] = []
    seen: set[tuple[str, bool]] = set()
    lines = text.splitlines()
    in_code = False
    in_math = False
    math_lines: list[str] = []

    def remember(latex: str, display_mode: bool) -> None:
        normalized = latex.strip()
        if not normalized:
            return
        key = (normalized, display_mode)
        if key not in seen:
            seen.add(key)
            items.append(key)

    for line in lines:
        stripped = line.strip()

        if in_code:
            if stripped.startswith("```"):
                in_code = False
            continue

        if in_math:
            if stripped == "$$":
                remember("\n".join(math_lines).strip("\n"), True)
                in_math = False
                math_lines = []
            else:
                math_lines.append(line.rstrip("\n"))
            continue

        if stripped.startswith("```"):
            in_code = True
            continue

        if stripped == "$$":
            in_math = True
            math_lines = []
            continue

        for segment_kind, segment_text in split_inline_code(line):
            if segment_kind != "text":
                continue
            for kind, value in split_inline_math(segment_text):
                if kind == "math":
                    remember(value, False)

    if in_math and math_lines:
        remember("\n".join(math_lines).strip("\n"), True)

    return items


class MathConverter:
    def __init__(self) -> None:
        self.cache: dict[tuple[str, bool], str | None] = {}
        self.ready = False
        self.failed = False
        self.failed_reason: str | None = None
        self.fallback_items: set[tuple[str, bool]] = set()
        self.item_errors: list[str] = []
        self.warning_reported = False

    def _remember_failure(self, reason: str) -> None:
        self.failed = True
        if self.failed_reason is None:
            self.failed_reason = reason

    def _remember_item_error(self, message: str) -> None:
        cleaned = message.strip()
        if cleaned and cleaned not in self.item_errors and len(self.item_errors) < 3:
            self.item_errors.append(cleaned)

    def ensure_ready(self) -> bool:
        if self.failed:
            return False
        if self.ready:
            return True
        if not WORD_MATH_SCRIPT.exists():
            self._remember_failure(f"missing converter script: {WORD_MATH_SCRIPT}")
            return False
        missing_modules = [str(path) for path in WORD_MATH_REQUIRED_MODULES if not path.exists()]
        if missing_modules:
            self._remember_failure(
                "formula converter dependencies are not installed"
            )
            return False
        self.ready = True
        return True

    def convert_many(self, items: list[tuple[str, bool]]) -> None:
        pending = []
        for latex, display_mode in items:
            key = (latex.strip(), display_mode)
            if key[0] and key not in self.cache:
                pending.append(key)
        if not pending:
            return
        if not self.ensure_ready():
            for key in pending:
                self.cache[key] = None
                self.fallback_items.add(key)
            return

        payload = {
            "items": [
                {"id": str(idx), "latex": latex, "displayMode": display_mode}
                for idx, (latex, display_mode) in enumerate(pending)
            ]
        }
        try:
            result = subprocess.run(
                ["node", str(WORD_MATH_SCRIPT)],
                cwd=WORD_MATH_DIR,
                input=json.dumps(payload, ensure_ascii=False),
                capture_output=True,
                text=True,
                check=True,
            )
            data = json.loads(result.stdout or "{}")
        except FileNotFoundError:
            self._remember_failure("node is not available, so formulas cannot be converted into Word equations")
            for key in pending:
                self.cache[key] = None
                self.fallback_items.add(key)
            return
        except subprocess.CalledProcessError as exc:
            detail = (exc.stderr or exc.stdout or "").strip().splitlines()
            reason = "the formula converter failed while invoking node"
            if detail:
                reason += f": {detail[0]}"
            self._remember_failure(reason)
            for key in pending:
                self.cache[key] = None
                self.fallback_items.add(key)
            return
        except json.JSONDecodeError:
            self._remember_failure("the formula converter returned invalid output")
            for key in pending:
                self.cache[key] = None
                self.fallback_items.add(key)
            return

        results = {str(item.get("id")): item for item in data.get("results", []) if isinstance(item, dict)}
        for idx, key in enumerate(pending):
            item = results.get(str(idx), {})
            omml = item.get("omml") if item.get("ok") else None
            sanitized = self.sanitize_omml(omml) if isinstance(omml, str) else None
            self.cache[key] = sanitized
            if sanitized is None:
                self.fallback_items.add(key)
                error_message = item.get("error") if isinstance(item, dict) else None
                if isinstance(error_message, str):
                    self._remember_item_error(error_message)
                elif isinstance(omml, str):
                    self._remember_item_error("converter returned invalid OMML")

    def preload_from_markdown(self, text: str) -> None:
        self.convert_many(collect_math_items(text))

    def get(self, latex: str, *, display_mode: bool) -> str | None:
        key = (latex.strip(), display_mode)
        if key[0] and key not in self.cache:
            self.convert_many([key])
        return self.cache.get(key)

    @staticmethod
    def sanitize_omml(omml: str) -> str | None:
        def repl(match: re.Match[str]) -> str:
            raw = match.group(2)
            cleaned = escape(html.unescape(raw))
            return f"{match.group(1)}{cleaned}{match.group(3)}"

        sanitized = OMML_TEXT_PATTERN.sub(repl, omml)
        try:
            ET.fromstring(sanitized)
        except ET.ParseError:
            return None
        return sanitized

    def emit_warning(self) -> None:
        if self.warning_reported:
            return
        self.warning_reported = True

        fallback_count = len(self.fallback_items)
        if fallback_count == 0:
            return

        if self.failed_reason:
            install_dir = str(WORD_MATH_DIR.resolve())
            print(
                (
                    "[warning] Word formula conversion is unavailable: "
                    f"{self.failed_reason}. {fallback_count} formula(s) were kept as raw LaTeX.\n"
                    "          To enable Word equations, install the helper dependencies with:\n"
                    f"          cd {install_dir}\n"
                    "          npm install"
                ),
                file=sys.stderr,
            )
            return

        detail = f" Example converter error: {self.item_errors[0]}" if self.item_errors else ""
        print(
            (
                f"[warning] {fallback_count} formula(s) could not be converted to Word equations "
                f"and were kept as raw LaTeX.{detail}"
            ),
            file=sys.stderr,
        )


@dataclass
class MediaImage:
    source_path: Path
    filename: str
    part_name: str
    rel_id: str
    content_type: str
    width_emu: int
    height_emu: int


class MediaManager:
    def __init__(self, *, starting_rid: int = 2, starting_image_index: int = 1) -> None:
        self.starting_rid = starting_rid
        self.next_rid = starting_rid
        self.next_image_index = starting_image_index
        self.next_docpr_id = 1
        self.images: list[MediaImage] = []
        self.by_path: dict[Path, MediaImage] = {}

    def register_image(self, source_path: Path) -> MediaImage | None:
        resolved = source_path.resolve()
        if resolved in self.by_path:
            return self.by_path[resolved]
        if not resolved.exists() or not resolved.is_file():
            return None

        suffix = resolved.suffix.lower().lstrip(".")
        content_type = IMAGE_CONTENT_TYPES.get(suffix)
        if not content_type:
            return None

        width_emu, height_emu = image_extent_emu(resolved)
        rel_id = f"rId{self.next_rid}"
        self.next_rid += 1
        filename = f"image{self.next_image_index}{resolved.suffix.lower()}"
        self.next_image_index += 1
        item = MediaImage(
            source_path=resolved,
            filename=filename,
            part_name=f"media/{filename}",
            rel_id=rel_id,
            content_type=content_type,
            width_emu=width_emu,
            height_emu=height_emu,
        )
        self.images.append(item)
        self.by_path[resolved] = item
        return item

    def next_drawing_id(self) -> int:
        current = self.next_docpr_id
        self.next_docpr_id += 1
        return current

    def image_extensions(self) -> set[str]:
        return {item.filename.rsplit(".", 1)[-1].lower() for item in self.images if "." in item.filename}


def relationship_id_number(rel_id: str) -> int | None:
    match = re.fullmatch(r"rId(\d+)", rel_id.strip())
    return int(match.group(1)) if match else None


def next_relationship_id(rels_data: bytes) -> int:
    try:
        root = ET.fromstring(rels_data)
    except ET.ParseError:
        return 2
    max_id = 1
    for rel in root.findall(f"{{{PKG_REL_NS}}}Relationship"):
        rel_id = rel.get("Id", "")
        number = relationship_id_number(rel_id)
        if number is not None:
            max_id = max(max_id, number)
    return max_id + 1


def next_image_index_from_template(template_path: Path) -> int:
    max_index = 0
    try:
        with zipfile.ZipFile(template_path) as zf:
            for name in zf.namelist():
                match = re.match(r"word/media/image(\d+)\.[A-Za-z0-9]+$", name)
                if match:
                    max_index = max(max_index, int(match.group(1)))
    except (FileNotFoundError, zipfile.BadZipFile):
        return 1
    return max_index + 1


def image_extent_emu(path: Path) -> tuple[int, int]:
    default_width = int(MAX_IMAGE_WIDTH_IN * EMU_PER_INCH)
    default_height = int(3.8 * EMU_PER_INCH)
    if Image is None:
        return default_width, default_height

    try:
        with Image.open(path) as img:
            width_px, height_px = img.size
            dpi_info = img.info.get("dpi", (DEFAULT_DPI, DEFAULT_DPI))
    except Exception:
        return default_width, default_height

    if width_px <= 0 or height_px <= 0:
        return default_width, default_height

    try:
        dpi_x = float(dpi_info[0]) if dpi_info and dpi_info[0] else DEFAULT_DPI
        dpi_y = float(dpi_info[1]) if dpi_info and len(dpi_info) > 1 and dpi_info[1] else dpi_x
    except (TypeError, ValueError):
        dpi_x = dpi_y = DEFAULT_DPI

    dpi_x = dpi_x if dpi_x > 1 else DEFAULT_DPI
    dpi_y = dpi_y if dpi_y > 1 else DEFAULT_DPI

    width_emu = int(width_px / dpi_x * EMU_PER_INCH)
    height_emu = int(height_px / dpi_y * EMU_PER_INCH)

    max_width_emu = int(MAX_IMAGE_WIDTH_IN * EMU_PER_INCH)
    max_height_emu = int(MAX_IMAGE_HEIGHT_IN * EMU_PER_INCH)
    scale = min(
        1.0,
        max_width_emu / width_emu if width_emu else 1.0,
        max_height_emu / height_emu if height_emu else 1.0,
    )
    width_emu = max(1, int(width_emu * scale))
    height_emu = max(1, int(height_emu * scale))
    return width_emu, height_emu


def fit_extent_emu(
    width_emu: int,
    height_emu: int,
    *,
    max_width_emu: int,
    max_height_emu: int,
) -> tuple[int, int]:
    if width_emu <= 0 or height_emu <= 0:
        return max_width_emu, max_height_emu
    scale = min(
        1.0,
        max_width_emu / width_emu if width_emu else 1.0,
        max_height_emu / height_emu if height_emu else 1.0,
    )
    return max(1, int(width_emu * scale)), max(1, int(height_emu * scale))


def image_run_xml(
    item: MediaImage,
    *,
    docpr_id: int,
    alt_text: str = "",
    width_emu: int | None = None,
    height_emu: int | None = None,
) -> str:
    width_emu = width_emu or item.width_emu
    height_emu = height_emu or item.height_emu
    descr = escape(alt_text or item.filename)
    name = escape(item.filename)
    return (
        "<w:r><w:drawing>"
        '<wp:inline distT="0" distB="0" distL="0" distR="0">'
        f'<wp:extent cx="{width_emu}" cy="{height_emu}"/>'
        '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        f'<wp:docPr id="{docpr_id}" name="{name}" descr="{descr}"/>'
        '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
        "<a:graphic>"
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        "<pic:pic>"
        "<pic:nvPicPr>"
        f'<pic:cNvPr id="{docpr_id}" name="{name}"/>'
        "<pic:cNvPicPr/>"
        "</pic:nvPicPr>"
        "<pic:blipFill>"
        f'<a:blip r:embed="{item.rel_id}"/>'
        "<a:stretch><a:fillRect/></a:stretch>"
        "</pic:blipFill>"
        "<pic:spPr>"
        '<a:xfrm><a:off x="0" y="0"/>'
        f'<a:ext cx="{width_emu}" cy="{height_emu}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        "</pic:spPr>"
        "</pic:pic>"
        "</a:graphicData>"
        "</a:graphic>"
        "</wp:inline>"
        "</w:drawing></w:r>"
    )


def image_paragraph_xml(item: MediaImage, media_manager: MediaManager, *, alt_text: str = "") -> str:
    return paragraph_xml(
        align="center",
        runs=[image_run_xml(item, docpr_id=media_manager.next_drawing_id(), alt_text=alt_text)],
        ppr_extra=spacing_xml(after=120),
    )


def figure_row_xml(
    items: list[tuple[MediaImage | None, str]],
    media_manager: MediaManager,
) -> str:
    if not items:
        return ""

    col_count = len(items)
    col_width = max(1800, 9000 // col_count)
    max_width_emu = int(FIGURE_ROW_MAX_WIDTH_IN * EMU_PER_INCH)
    max_height_emu = int(FIGURE_ROW_MAX_HEIGHT_IN * EMU_PER_INCH)
    common_height_emu = max_height_emu
    for item, _ in items:
        if item is None or item.width_emu <= 0 or item.height_emu <= 0:
            continue
        height_limit_by_width = int(max_width_emu * item.height_emu / item.width_emu)
        common_height_emu = min(common_height_emu, max(1, height_limit_by_width))
    common_height_emu = max(1, min(common_height_emu, max_height_emu))
    tbl_pr = (
        "<w:tblPr>"
        '<w:tblW w:w="9000" w:type="dxa"/>'
        '<w:jc w:val="center"/>'
        "<w:tblBorders>"
        '<w:top w:val="nil"/>'
        '<w:left w:val="nil"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="nil"/>'
        '<w:insideH w:val="nil"/>'
        '<w:insideV w:val="nil"/>'
        "</w:tblBorders>"
        "</w:tblPr>"
    )
    tbl_grid = "<w:tblGrid>" + "".join(f'<w:gridCol w:w="{col_width}"/>' for _ in range(col_count)) + "</w:tblGrid>"

    cells: list[str] = []
    for item, alt_text in items:
        body: list[str] = []
        tc_pr = f'<w:tcPr><w:tcW w:w="{col_width}" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>'
        if item is None:
            body.append(
                formatted_paragraph_xml(
                    "图片待补充",
                    align="center",
                    ppr_extra=spacing_xml(after=60),
                    run_kwargs={"italic": True},
                )
            )
        else:
            width_emu = max(1, int(item.width_emu * common_height_emu / item.height_emu))
            height_emu = common_height_emu
            if width_emu > max_width_emu:
                width_emu, height_emu = fit_extent_emu(
                    item.width_emu,
                    item.height_emu,
                    max_width_emu=max_width_emu,
                    max_height_emu=max_height_emu,
                )
            body.append(
                paragraph_xml(
                    align="center",
                    runs=[
                        image_run_xml(
                            item,
                            docpr_id=media_manager.next_drawing_id(),
                            alt_text=alt_text,
                            width_emu=width_emu,
                            height_emu=height_emu,
                        )
                    ],
                    ppr_extra=spacing_xml(after=80),
                )
            )
        if alt_text:
            body.append(paragraph_xml(alt_text, align="center", ppr_extra=spacing_xml(after=0)))
        cells.append(f"<w:tc>{tc_pr}{''.join(body)}</w:tc>")

    return f"<w:tbl>{tbl_pr}{tbl_grid}<w:tr>{''.join(cells)}</w:tr></w:tbl>"


def is_caption_paragraph(text: str) -> bool:
    return bool(CAPTION_PATTERN.match(text.strip()))


def paragraph_xml(
    text: str | None = None,
    *,
    style: str | None = None,
    align: str | None = None,
    preserve_breaks: bool = False,
    runs: list[str] | None = None,
    ppr_extra: str = "",
    first_line_chars: int | None = None,
    first_line: int | None = None,
) -> str:
    ppr: list[str] = []
    if style:
        ppr.append(f'<w:pStyle w:val="{style}"/>')
    if align:
        ppr.append(f'<w:jc w:val="{align}"/>')
    indent = indent_xml(first_line_chars=first_line_chars, first_line=first_line)
    if indent:
        ppr.append(indent)
    if ppr_extra:
        ppr.append(ppr_extra)
    ppr_xml = f"<w:pPr>{''.join(ppr)}</w:pPr>" if ppr else ""

    if runs is None:
        value = text or ""
        if preserve_breaks and "\n" in value:
            body = "".join(text_runs(value, preserve_breaks=True))
        else:
            body = f"<w:r>{xml_text(value)}</w:r>"
    else:
        body = "".join(runs)
    return f"<w:p>{ppr_xml}{body}</w:p>"


def formatted_paragraph_xml(
    text: str,
    *,
    style: str | None = None,
    align: str | None = None,
    ppr_extra: str = "",
    first_line_chars: int | None = None,
    first_line: int | None = None,
    run_kwargs: dict[str, object] | None = None,
    preserve_breaks: bool = False,
) -> str:
    runs = text_runs(text, run_kwargs=run_kwargs, preserve_breaks=preserve_breaks)
    return paragraph_xml(
        style=style,
        align=align,
        runs=runs,
        ppr_extra=ppr_extra,
        first_line_chars=first_line_chars,
        first_line=first_line,
    )


def page_break_xml() -> str:
    return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'


def section_break_paragraph_xml(sect_pr: str) -> str:
    return f"<w:p><w:pPr>{sect_pr}</w:pPr></w:p>"


def toc_field_paragraph_xml() -> str:
    runs = [
        field_char_run_xml("begin", dirty=True),
        # Restrict the TOC to heading styles only. The school template marks some
        # non-heading styles (for example the code block style) with outline levels,
        # and the `\\u` switch would pull those paragraphs into the TOC.
        instr_text_run_xml(' TOC \\o "1-3" \\h \\z '),
        field_char_run_xml("separate"),
        run_text_xml(" ", size=24),
        field_char_run_xml("end"),
    ]
    return paragraph_xml(
        runs=runs,
        style="10",
        ppr_extra=spacing_xml(line=288),
    )


def split_markdown_row(line: str) -> list[str]:
    raw = line.strip()
    if raw.startswith("|"):
        raw = raw[1:]
    if raw.endswith("|"):
        raw = raw[:-1]
    return [cell.strip() for cell in raw.split("|")]


def is_table_separator(line: str) -> bool:
    cells = split_markdown_row(line)
    if not cells:
        return False
    return all(re.fullmatch(r":?-{3,}:?", cell) for cell in cells)


def table_xml(
    rows: list[list[str]],
    cell_style: str = "TableText",
    *,
    math_converter: MathConverter | None = None,
    reference_anchors: dict[str, str] | None = None,
) -> str:
    col_count = max(len(rows[0]), 1)
    col_width = max(1200, 9000 // col_count)
    tbl_pr = (
        "<w:tblPr>"
        '<w:tblW w:w="9000" w:type="dxa"/>'
        '<w:jc w:val="center"/>'
        "<w:tblBorders>"
        '<w:top w:val="single" w:sz="12" w:space="0" w:color="auto"/>'
        '<w:left w:val="nil"/>'
        '<w:bottom w:val="single" w:sz="12" w:space="0" w:color="auto"/>'
        '<w:right w:val="nil"/>'
        '<w:insideH w:val="nil"/>'
        '<w:insideV w:val="nil"/>'
        "</w:tblBorders>"
        "</w:tblPr>"
    )
    tbl_grid = "<w:tblGrid>" + "".join(f'<w:gridCol w:w="{col_width}"/>' for _ in range(col_count)) + "</w:tblGrid>"

    trs = []
    for r_idx, row in enumerate(rows):
        cells = []
        for cell in row:
            cell_text = cell.strip()
            p = paragraph_with_inline_math_xml(
                cell_text,
                style=cell_style,
                align="center",
                ppr_extra=spacing_xml(after=0),
                run_kwargs={"bold": r_idx == 0},
                math_converter=math_converter,
                reference_anchors=reference_anchors,
            )
            cell_pr = ""
            if r_idx == 0:
                cell_pr = (
                    "<w:tcPr><w:tcBorders>"
                    '<w:bottom w:val="single" w:sz="8" w:space="0" w:color="auto"/>'
                    "</w:tcBorders></w:tcPr>"
                )
            cells.append(f"<w:tc>{cell_pr}{p}</w:tc>")
        trs.append(f"<w:tr>{''.join(cells)}</w:tr>")
    return f"<w:tbl>{tbl_pr}{tbl_grid}{''.join(trs)}</w:tbl>"


def split_plain_paragraphs(text: str) -> list[str]:
    paragraphs: list[str] = []
    buffer: list[str] = []
    for line in text.splitlines():
        stripped = line.strip()
        if not stripped:
            if buffer:
                paragraphs.append(" ".join(part.strip() for part in buffer if part.strip()))
                buffer = []
            continue
        if stripped.startswith(">"):
            stripped = stripped[1:].strip()
        buffer.append(stripped)
    if buffer:
        paragraphs.append(" ".join(part.strip() for part in buffer if part.strip()))
    return paragraphs


def parse_markdown_document(text: str) -> tuple[str, dict[str, str], str]:
    lines = text.splitlines()
    title = ""
    front_sections: dict[str, str] = {}
    current_section: str | None = None
    buffer: list[str] = []
    body_start = len(lines)

    for idx, line in enumerate(lines):
        if not title:
            match = re.match(r"^#\s+(.*)$", line)
            if match:
                title = match.group(1).strip()
                continue

        if re.match(r"^#\s+\d+\b", line):
            body_start = idx
            break

        section_match = re.match(r"^##\s+(.*)$", line)
        if section_match:
            if current_section is not None:
                front_sections[current_section] = "\n".join(buffer).strip()
            current_section = section_match.group(1).strip()
            buffer = []
            continue

        if re.fullmatch(r"-{3,}|\*{3,}", line.strip()):
            if current_section is not None:
                front_sections[current_section] = "\n".join(buffer).strip()
                current_section = None
                buffer = []
            continue

        if current_section is not None:
            buffer.append(line)

    if current_section is not None:
        front_sections[current_section] = "\n".join(buffer).strip()

    body_text = "\n".join(lines[body_start:]).strip()
    return title, front_sections, body_text


def parse_cover_info(text: str) -> dict[str, str]:
    info: dict[str, str] = {}
    for line in text.splitlines():
        stripped = line.strip()
        if not stripped or stripped.startswith(">"):
            continue
        if "：" in stripped:
            key, value = stripped.split("：", 1)
        elif ":" in stripped:
            key, value = stripped.split(":", 1)
        else:
            continue
        info[key.strip()] = value.strip()
    return info


def extract_abstract_and_keywords(text: str, keyword_prefix: str) -> tuple[list[str], str]:
    paragraphs = split_plain_paragraphs(text)
    body: list[str] = []
    keywords = ""
    for paragraph in paragraphs:
        if paragraph.startswith(keyword_prefix):
            keywords = paragraph[len(keyword_prefix):].strip()
        else:
            body.append(paragraph)
    return body, keywords


def split_cover_title_lines(title: str) -> list[str]:
    compact = re.sub(r"\s+", "", title.strip())
    if not compact:
        return [""]
    if len(compact) <= 14:
        return [compact]
    if len(compact) <= 28:
        split_at = (len(compact) + 1) // 2
        return [compact[:split_at], compact[split_at:]]

    lines: list[str] = []
    chunk = 14
    for idx in range(0, len(compact), chunk):
        lines.append(compact[idx : idx + chunk])
    return lines


def resolve_cover_assets_dir(markdown_path: Path, assets_dir: Path | None, *, use_cover_assets: bool) -> Path | None:
    if not use_cover_assets:
        return None

    candidates: list[Path] = []
    if assets_dir is not None:
        candidates.append(assets_dir)
    local_assets_dir = markdown_path.parent / DEFAULT_LOCAL_COVER_ASSETS_REL
    if local_assets_dir not in candidates:
        candidates.append(local_assets_dir)

    for candidate in candidates:
        if (candidate / COVER_EMBLEM_NAME).exists() or (candidate / COVER_WORDMARK_NAME).exists():
            return candidate

    return candidates[0] if candidates else None


def cover_logo_table_xml(
    emblem_item: MediaImage | None,
    wordmark_item: MediaImage | None,
    media_manager: MediaManager | None,
) -> str:
    if media_manager is None or (emblem_item is None and wordmark_item is None):
        return ""

    tbl_pr = (
        "<w:tblPr>"
        '<w:tblW w:w="6200" w:type="dxa"/>'
        '<w:jc w:val="center"/>'
        '<w:tblLayout w:type="fixed"/>'
        "<w:tblBorders>"
        '<w:top w:val="nil"/>'
        '<w:left w:val="nil"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="nil"/>'
        '<w:insideH w:val="nil"/>'
        '<w:insideV w:val="nil"/>'
        "</w:tblBorders>"
        "</w:tblPr>"
    )
    tbl_grid = '<w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="4700"/></w:tblGrid>'

    def cover_logo_cell(item: MediaImage | None, *, max_width_in: float, max_height_in: float) -> str:
        if item is None or media_manager is None:
            body = paragraph_xml(" ", align="center", ppr_extra=spacing_xml(after=0))
        else:
            width_emu, height_emu = fit_extent_emu(
                item.width_emu,
                item.height_emu,
                max_width_emu=int(max_width_in * EMU_PER_INCH),
                max_height_emu=int(max_height_in * EMU_PER_INCH),
            )
            body = paragraph_xml(
                align="center",
                runs=[
                    image_run_xml(
                        item,
                        docpr_id=media_manager.next_drawing_id(),
                        alt_text=item.filename,
                        width_emu=width_emu,
                        height_emu=height_emu,
                    )
                ],
                ppr_extra=spacing_xml(after=0),
            )
        return "<w:tc><w:tcPr><w:vAlign w:val=\"center\"/></w:tcPr>" + body + "</w:tc>"

    row = (
        "<w:tr>"
        '<w:trPr><w:trHeight w:val="900" w:hRule="atLeast"/></w:trPr>'
        + cover_logo_cell(emblem_item, max_width_in=1.05, max_height_in=1.05)
        + cover_logo_cell(wordmark_item, max_width_in=4.55, max_height_in=1.55)
        + "</w:tr>"
    )
    return f"<w:tbl>{tbl_pr}{tbl_grid}{row}</w:tbl>"


def cover_info_table_xml(title: str, cover_info: dict[str, str]) -> str:
    title_lines = split_cover_title_lines(title)
    info_rows: list[tuple[str, str, bool]] = []

    if title_lines:
        info_rows.append(("论文题目:", title_lines[0], False))
        for extra_line in title_lines[1:]:
            info_rows.append(("", extra_line, True))

    ordered_fields = [
        ("学生姓名", "学生姓名:"),
        ("学号", "学    号:"),
        ("所属院系", "所属院系:"),
        ("专业", "专    业:"),
        ("班级", "班    级:"),
        ("指导教师", "指导教师:"),
        ("日期", "日    期:"),
    ]
    for source_key, display_label in ordered_fields:
        value = cover_info.get(source_key)
        if value:
            info_rows.append((display_label, value, True))

    tbl_pr = (
        "<w:tblPr>"
        '<w:tblW w:w="6943" w:type="dxa"/>'
        '<w:jc w:val="center"/>'
        '<w:tblLayout w:type="fixed"/>'
        "<w:tblCellMar>"
        '<w:top w:w="0" w:type="dxa"/>'
        '<w:left w:w="108" w:type="dxa"/>'
        '<w:bottom w:w="0" w:type="dxa"/>'
        '<w:right w:w="108" w:type="dxa"/>'
        "</w:tblCellMar>"
        "<w:tblBorders>"
        '<w:top w:val="nil"/>'
        '<w:left w:val="nil"/>'
        '<w:bottom w:val="nil"/>'
        '<w:right w:val="nil"/>'
        '<w:insideH w:val="nil"/>'
        '<w:insideV w:val="nil"/>'
        "</w:tblBorders>"
        "</w:tblPr>"
    )
    tbl_grid = '<w:tblGrid><w:gridCol w:w="1948"/><w:gridCol w:w="4995"/></w:tblGrid>'

    label_run = {
        "font_ascii": "Times New Roman",
        "font_hansi": "Times New Roman",
        "font_eastasia": "楷体_GB2312",
        "bold": True,
        "size": 32,
    }
    value_run = {
        "font_ascii": "Times New Roman",
        "font_hansi": "Times New Roman",
        "font_eastasia": "楷体_GB2312",
        "bold": True,
        "size": 32,
    }

    rows_xml: list[str] = []
    for idx, (label, value, draw_top) in enumerate(info_rows):
        label_para = formatted_paragraph_xml(
            label,
            align="center",
            ppr_extra=spacing_xml(before=240, after=120),
            run_kwargs=label_run,
        )
        value_para = formatted_paragraph_xml(
            value,
            align="center",
            ppr_extra=spacing_xml(before=240, after=120),
            run_kwargs=value_run,
        )
        value_borders = ["<w:tcBorders>"]
        if draw_top and idx > 0:
            value_borders.append('<w:top w:val="single" w:color="auto" w:sz="4" w:space="0"/>')
        value_borders.append('<w:bottom w:val="single" w:color="auto" w:sz="4" w:space="0"/>')
        value_borders.append("</w:tcBorders>")

        rows_xml.append(
            "<w:tr>"
            '<w:trPr><w:trHeight w:val="680" w:hRule="atLeast"/></w:trPr>'
            '<w:tc><w:tcPr><w:tcW w:w="1948" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>'
            + label_para
            + "</w:tc>"
            + '<w:tc><w:tcPr><w:tcW w:w="4995" w:type="dxa"/>'
            + "".join(value_borders)
            + '<w:vAlign w:val="center"/></w:tcPr>'
            + value_para
            + "</w:tc>"
            + "</w:tr>"
        )

    return f"<w:tbl>{tbl_pr}{tbl_grid}{''.join(rows_xml)}</w:tbl>"


def build_cover_elements(
    title: str,
    cover_info: dict[str, str],
    *,
    cover_assets_dir: Path | None = None,
    media_manager: MediaManager | None = None,
) -> list[str]:
    elements: list[str] = []
    title_run = {
        "font_ascii": "Times New Roman",
        "font_hansi": "Times New Roman",
        "font_eastasia": "黑体",
    }

    elements.append(
        formatted_paragraph_xml(
            "新疆大学本科毕业论文(设计)",
            align="center",
            ppr_extra=spacing_xml(before=240, after=120, line=600),
            run_kwargs={**title_run, "bold": True, "size": 52},
        )
    )

    elements.append(paragraph_xml(" ", ppr_extra=spacing_xml(after=60)))

    emblem_item = None
    wordmark_item = None
    if cover_assets_dir is not None and media_manager is not None:
        emblem_item = media_manager.register_image(cover_assets_dir / COVER_EMBLEM_NAME)
        wordmark_item = media_manager.register_image(cover_assets_dir / COVER_WORDMARK_NAME)

    logo_tbl = cover_logo_table_xml(emblem_item, wordmark_item, media_manager)
    if logo_tbl:
        elements.append(logo_tbl)

    for _ in range(4):
        elements.append(paragraph_xml(" ", ppr_extra=spacing_xml(after=40)))

    elements.append(cover_info_table_xml(title, cover_info))

    return elements


def build_front_heading(text: str, *, english: bool = False, toc: bool = False) -> str:
    if toc:
        return paragraph_xml("目  录", style="afa")

    if english:
        run_kwargs = {
            "font_ascii": "Times New Roman",
            "font_hansi": "Times New Roman",
            "font_eastasia": "Times New Roman",
            "size": 32,
        }
        ppr_extra = spacing_xml(before_lines=300, before=720, after_lines=200, after=480)
    else:
        run_kwargs = {
            "font_ascii": "黑体",
            "font_hansi": "黑体",
            "font_eastasia": "黑体",
            "size": 32,
        }
        ppr_extra = '<w:snapToGrid w:val="0"/>' + spacing_xml(before_lines=300, before=720, after_lines=200, after=480)

    return formatted_paragraph_xml(
        text,
        align="center",
        ppr_extra=ppr_extra,
        run_kwargs=run_kwargs,
    )


def build_body_paragraph(
    text: str,
    *,
    english: bool = False,
    math_converter: MathConverter | None = None,
    reference_anchors: dict[str, str] | None = None,
) -> str:
    run_kwargs = {
        "font_ascii": "Times New Roman",
        "font_hansi": "Times New Roman",
        "font_eastasia": "Times New Roman" if english else "宋体",
        "size": 24,
    }
    return paragraph_with_inline_math_xml(
        text,
        style="a0",
        ppr_extra=spacing_xml(line=360),
        first_line_chars=200,
        first_line=480,
        run_kwargs=run_kwargs,
        math_converter=math_converter,
        reference_anchors=reference_anchors,
    )


def build_caption_paragraph(
    text: str,
    *,
    style: str | None = None,
    english: bool = False,
    math_converter: MathConverter | None = None,
    reference_anchors: dict[str, str] | None = None,
) -> str:
    run_kwargs = {
        "font_ascii": "Times New Roman",
        "font_hansi": "Times New Roman",
        "font_eastasia": "Times New Roman" if english else "宋体",
        "size": 21,
    }
    return paragraph_with_inline_math_xml(
        text,
        style=style,
        align="center",
        ppr_extra=spacing_xml(after=120),
        run_kwargs=run_kwargs,
        math_converter=math_converter,
        reference_anchors=reference_anchors,
    )


def build_keyword_paragraph(keywords: str, *, english: bool = False) -> str | None:
    if not keywords:
        return None
    if english:
        runs = [
            run_text_xml(
                "KEY WORDS: ",
                bold=True,
                font_ascii="Times New Roman",
                font_hansi="Times New Roman",
                font_eastasia="Times New Roman",
                size=24,
            ),
            run_text_xml(
                keywords,
                font_ascii="Times New Roman",
                font_hansi="Times New Roman",
                font_eastasia="Times New Roman",
                size=24,
            ),
        ]
    else:
        runs = [
            run_text_xml(
                "关 键 词：",
                bold=True,
                font_ascii="Times New Roman",
                font_hansi="Times New Roman",
                font_eastasia="宋体",
                size=24,
            ),
            run_text_xml(
                keywords,
                font_ascii="Times New Roman",
                font_hansi="Times New Roman",
                font_eastasia="宋体",
                size=24,
            ),
        ]
    return paragraph_xml(runs=runs, style="a0", ppr_extra=spacing_xml(line=360))


def build_reference_paragraph(text: str, reference_anchors: dict[str, str] | None = None) -> str:
    run_kwargs = {
        "font_ascii": "Times New Roman",
        "font_hansi": "Times New Roman",
        "font_eastasia": "宋体",
        "size": 21,
    }
    match = re.match(r"^\[(\d+)\]\s*(.*)$", text)
    if not match:
        return formatted_paragraph_xml(
            text,
            style="a0",
            ppr_extra=spacing_xml(line=288) + indent_xml(left=420, hanging=420),
            run_kwargs=run_kwargs,
        )

    ref_id, rest = match.groups()
    anchor = reference_anchors.get(ref_id, reference_bookmark_name(ref_id)) if reference_anchors else reference_bookmark_name(ref_id)
    bookmark_id = reference_bookmark_id(ref_id)
    runs = [
        f'<w:bookmarkStart w:id="{bookmark_id}" w:name="{escape(anchor)}"/>',
        run_text_xml(f"[{ref_id}] ", **run_kwargs),
        f'<w:bookmarkEnd w:id="{bookmark_id}"/>',
    ]
    if rest:
        runs.extend(text_runs(rest, run_kwargs=run_kwargs))
    return paragraph_xml(
        style="a0",
        runs=runs,
        ppr_extra=spacing_xml(line=288) + indent_xml(left=420, hanging=420),
    )


def strip_heading_prefix(text: str) -> str:
    stripped = re.sub(r"^\d+(?:\.\d+)*\s+", "", text).strip()
    return stripped or text.strip()


def heading_paragraph_xml(text: str, level: int, profile: dict[str, object], *, numbered: bool = True) -> str:
    if level == 1:
        style = profile.get("heading1")
    elif level == 2:
        style = profile.get("heading2")
    else:
        style = profile.get("heading3")

    if numbered:
        heading_text = strip_heading_prefix(text) if profile.get("strip_heading_numbers") else text.strip()
        return paragraph_xml(heading_text, style=str(style) if style else None)

    ppr_extra = f'<w:numPr><w:ilvl w:val="{level - 1}"/><w:numId w:val="0"/></w:numPr>'
    if level == 1:
        ppr_extra += spacing_xml(line=240)
    return paragraph_xml(text.strip(), style=str(style) if style else None, ppr_extra=ppr_extra)


def build_document_elements(
    text: str,
    profile: dict[str, object] | None = None,
    *,
    treat_first_heading_as_title: bool = True,
    math_converter: MathConverter | None = None,
    reference_anchors: dict[str, str] | None = None,
    markdown_dir: Path | None = None,
    media_manager: MediaManager | None = None,
) -> list[str]:
    lines = text.splitlines()
    elements: list[str] = []
    paragraph_buffer: list[str] = []
    i = 0
    in_code = False
    code_lines: list[str] = []
    in_math = False
    math_lines: list[str] = []
    current_top_heading = ""
    in_appendix = False

    profile = profile or {
        "title": "Heading1",
        "heading1": "Heading1",
        "heading2": "Heading2",
        "heading3": "Heading3",
        "normal": None,
        "quote": "Quote",
        "code": "CodeBlock",
        "code_ppr_extra": '<w:outlineLvl w:val="9"/>',
        "math": "MathBlock",
        "table": "TableText",
        "strip_heading_numbers": False,
    }

    def flush_paragraph() -> None:
        nonlocal paragraph_buffer
        if not paragraph_buffer:
            return
        paragraph = " ".join(line.strip() for line in paragraph_buffer).strip()
        paragraph_buffer = []
        if not paragraph:
            return

        if current_top_heading == "参考文献":
            if profile.get("skip_reference_notes") and paragraph.startswith("说明："):
                return
            if re.match(r"^\[\d+\]", paragraph):
                elements.append(build_reference_paragraph(paragraph, reference_anchors=reference_anchors))
                return

        if is_caption_paragraph(paragraph):
            elements.append(
                build_caption_paragraph(
                    paragraph,
                    style=str(profile.get("normal")) if profile.get("normal") else None,
                    math_converter=math_converter,
                    reference_anchors=reference_anchors,
                )
            )
            return

        normal_run = profile.get("normal_run")
        ppr_extra = str(profile.get("normal_ppr_extra", ""))
        if normal_run:
            elements.append(
                paragraph_with_inline_math_xml(
                    paragraph,
                    style=str(profile.get("normal")) if profile.get("normal") else None,
                    ppr_extra=ppr_extra,
                    first_line_chars=int(profile.get("normal_first_line_chars", 0) or 0) or None,
                    first_line=int(profile.get("normal_first_line", 0) or 0) or None,
                    run_kwargs=dict(normal_run),
                    math_converter=math_converter,
                    reference_anchors=reference_anchors,
                )
            )
        else:
            elements.append(
                paragraph_with_inline_math_xml(
                    paragraph,
                    style=str(profile.get("normal")) if profile.get("normal") else None,
                    first_line_chars=int(profile.get("normal_first_line_chars", 0) or 0) or None,
                    first_line=int(profile.get("normal_first_line", 0) or 0) or None,
                    ppr_extra=ppr_extra,
                    math_converter=math_converter,
                    reference_anchors=reference_anchors,
                )
            )

    def resolve_image(target: str) -> MediaImage | None:
        image_path = markdown_dir / target if markdown_dir else Path(target)
        return media_manager.register_image(image_path) if media_manager else None

    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        if in_code:
            if stripped.startswith("```"):
                code_text = "\n".join(code_lines).rstrip("\n")
                if code_text:
                    elements.append(
                        paragraph_xml(
                            code_text,
                            style=str(profile.get("code")) if profile.get("code") else None,
                            preserve_breaks=True,
                            ppr_extra=str(profile.get("code_ppr_extra", "")),
                        )
                    )
                in_code = False
                code_lines = []
            else:
                code_lines.append(line.rstrip("\n"))
            i += 1
            continue

        if in_math:
            if stripped == "$$":
                math_text = "\n".join(math_lines).strip("\n")
                if math_text:
                    elements.append(math_paragraph_xml(math_text, style=str(profile.get("math")) if profile.get("math") else None, math_converter=math_converter))
                in_math = False
                math_lines = []
            else:
                math_lines.append(line.rstrip("\n"))
            i += 1
            continue

        if stripped.startswith("```"):
            flush_paragraph()
            in_code = True
            code_lines = []
            i += 1
            continue

        if stripped == "$$":
            flush_paragraph()
            in_math = True
            math_lines = []
            i += 1
            continue

        if not stripped:
            flush_paragraph()
            i += 1
            continue

        if FIGURE_ROW_START_PATTERN.match(stripped):
            flush_paragraph()
            i += 1
            figure_items: list[tuple[MediaImage | None, str]] = []
            raw_block: list[str] = [line]
            while i < len(lines):
                candidate = lines[i]
                candidate_stripped = candidate.strip()
                raw_block.append(candidate)
                if FIGURE_ROW_END_PATTERN.match(candidate_stripped):
                    break
                if candidate_stripped:
                    image_match = IMAGE_PATTERN.match(candidate_stripped)
                    if image_match:
                        alt_text = image_match.group("alt").strip()
                        target = image_match.group("target").strip()
                        figure_items.append((resolve_image(target), alt_text))
                i += 1
            if figure_items and media_manager is not None:
                figure_xml = figure_row_xml(figure_items, media_manager)
                if figure_xml:
                    elements.append(figure_xml)
            else:
                paragraph_buffer.extend(raw_block)
            i += 1
            continue

        image_match = IMAGE_PATTERN.match(stripped)
        if image_match:
            flush_paragraph()
            target = image_match.group("target").strip()
            alt_text = image_match.group("alt").strip()
            item = resolve_image(target)
            if item is not None:
                elements.append(image_paragraph_xml(item, media_manager, alt_text=alt_text))
            else:
                paragraph_buffer.append(line)
            i += 1
            continue

        if re.fullmatch(r"-{3,}|\*{3,}", stripped):
            flush_paragraph()
            elements.append(page_break_xml())
            i += 1
            continue

        heading_match = re.match(r"^(#{1,6})\s+(.*)$", line)
        if heading_match:
            flush_paragraph()
            level = min(len(heading_match.group(1)), 3)
            heading_text = heading_match.group(2).strip()

            if len(heading_match.group(1)) == 1:
                current_top_heading = heading_text
                in_appendix = heading_text == "附录"
            elif current_top_heading == "附录":
                in_appendix = True

            if len(elements) == 0 and treat_first_heading_as_title:
                elements.append(paragraph_xml(heading_text, style=str(profile.get("title")) if profile.get("title") else None, align="center"))
                i += 1
                continue

            is_unnumbered = False
            if heading_text in {"参考文献", "致谢", "附录"}:
                is_unnumbered = True
            elif in_appendix and level >= 2:
                is_unnumbered = True

            elements.append(heading_paragraph_xml(heading_text, level, profile, numbered=not is_unnumbered))
            i += 1
            continue

        if stripped.startswith(">"):
            flush_paragraph()
            quote = stripped[1:].strip()
            elements.append(
                paragraph_xml(
                    quote,
                    style=str(profile.get("quote")) if profile.get("quote") else None,
                )
            )
            i += 1
            continue

        if "|" in line and i + 1 < len(lines) and is_table_separator(lines[i + 1]):
            flush_paragraph()
            rows = [split_markdown_row(line)]
            i += 2
            while i < len(lines):
                candidate = lines[i].strip()
                if not candidate or "|" not in candidate:
                    break
                rows.append(split_markdown_row(lines[i]))
                i += 1
            if rows:
                width = len(rows[0])
                normalized = [row[:width] + [""] * max(0, width - len(row)) for row in rows]
                elements.append(
                    table_xml(
                        normalized,
                        cell_style=str(profile.get("table", "TableText")),
                        math_converter=math_converter,
                        reference_anchors=reference_anchors,
                    )
                )
            continue

        paragraph_buffer.append(line)
        i += 1

    flush_paragraph()

    if in_code and code_lines:
        elements.append(
            paragraph_xml(
                "\n".join(code_lines),
                style=str(profile.get("code")) if profile.get("code") else None,
                preserve_breaks=True,
                ppr_extra=str(profile.get("code_ppr_extra", "")),
            )
        )
    if in_math and math_lines:
        elements.append(
            math_paragraph_xml(
                "\n".join(math_lines),
                style=str(profile.get("math")) if profile.get("math") else None,
                math_converter=math_converter,
            )
        )

    return elements


def default_sect_pr_xml() -> str:
    return (
        "<w:sectPr>"
        '<w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
        "</w:sectPr>"
    )


def document_xml(elements: list[str], sect_pr: str | None = None) -> str:
    sect_pr = sect_pr or default_sect_pr_xml()
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{W_NS}" xmlns:r="{R_NS}" xmlns:m="{M_NS}" xmlns:wp="{WP_NS}" xmlns:a="{A_NS}" xmlns:pic="{PIC_NS}">'
        f"<w:body>{''.join(elements)}{sect_pr}</w:body>"
        "</w:document>"
    )


def styles_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:styles xmlns:w="{W_NS}">'
        "<w:docDefaults>"
        '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体"/>'
        '<w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:rPrDefault>'
        '<w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="360" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
        "</w:docDefaults>"
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
        '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:qFormat/>'
        '<w:pPr><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="0"/></w:pPr>'
        '<w:rPr><w:b/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:qFormat/>'
        '<w:pPr><w:spacing w:before="200" w:after="100"/><w:outlineLvl w:val="1"/></w:pPr>'
        '<w:rPr><w:b/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:qFormat/>'
        '<w:pPr><w:spacing w:before="160" w:after="80"/><w:outlineLvl w:val="2"/></w:pPr>'
        '<w:rPr><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/>'
        '<w:pPr><w:ind w:left="720"/><w:spacing w:after="120"/></w:pPr><w:rPr><w:i/></w:rPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="CodeBlock"><w:name w:val="CodeBlock"/><w:basedOn w:val="Normal"/>'
        '<w:pPr><w:spacing w:after="120"/><w:shd w:val="clear" w:fill="F5F5F5"/></w:pPr>'
        '<w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:eastAsia="等线"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="MathBlock"><w:name w:val="MathBlock"/><w:basedOn w:val="Normal"/>'
        '<w:pPr><w:spacing w:after="120"/></w:pPr>'
        '<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math" w:eastAsia="Cambria Math"/></w:rPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="TableText"><w:name w:val="TableText"/><w:basedOn w:val="Normal"/>'
        '<w:pPr><w:spacing w:after="0"/></w:pPr></w:style>'
        "</w:styles>"
    )


def content_types_xml(image_extensions: set[str] | None = None) -> str:
    defaults = [
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
        '<Default Extension="xml" ContentType="application/xml"/>',
    ]
    for ext in sorted(image_extensions or set()):
        content_type = IMAGE_CONTENT_TYPES.get(ext)
        if content_type:
            defaults.append(f'<Default Extension="{ext}" ContentType="{content_type}"/>')
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        f'{"".join(defaults)}'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
        '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
        '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
        "</Types>"
    )


def rels_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
        "</Relationships>"
    )


def document_rels_xml(media_manager: MediaManager | None = None) -> str:
    relationships = [
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    ]
    if media_manager:
        for item in media_manager.images:
            relationships.append(
                f'<Relationship Id="{item.rel_id}" Type="{IMAGE_REL_TYPE}" Target="{item.part_name}"/>'
            )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'{"".join(relationships)}'
        "</Relationships>"
    )


def core_xml(title: str) -> str:
    created = dt.datetime.now(dt.timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<cp:coreProperties xmlns:cp="{CP_NS}" xmlns:dc="{DC_NS}" xmlns:dcterms="{DCTERMS_NS}" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="{XSI_NS}">'
        f"<dc:title>{escape(title)}</dc:title>"
        "<dc:creator>Codex</dc:creator>"
        "<cp:lastModifiedBy>Codex</cp:lastModifiedBy>"
        f'<dcterms:created xsi:type="dcterms:W3CDTF">{created}</dcterms:created>'
        f'<dcterms:modified xsi:type="dcterms:W3CDTF">{created}</dcterms:modified>'
        "</cp:coreProperties>"
    )


def app_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="{VT_NS}">'
        "<Application>Codex</Application>"
        "</Properties>"
    )


def extract_template_section_properties(template_path: Path) -> list[str]:
    with zipfile.ZipFile(template_path) as zf:
        data = zf.read("word/document.xml")
    root = ET.fromstring(data)
    ns = {"w": W_NS}
    body = root.find("w:body", ns)
    if body is None:
        return []

    sections: list[str] = []
    for child in body:
        if child.tag == f"{{{W_NS}}}p":
            ppr = child.find("w:pPr", ns)
            if ppr is not None:
                sect = ppr.find("w:sectPr", ns)
                if sect is not None:
                    sections.append(ET.tostring(sect, encoding="unicode"))
        elif child.tag == f"{{{W_NS}}}sectPr":
            sections.append(ET.tostring(child, encoding="unicode"))
    return sections


def ensure_update_fields_xml(settings_data: bytes) -> bytes:
    root = ET.fromstring(settings_data)
    update_node = root.find(f"{{{W_NS}}}updateFields")
    if update_node is None:
        update_node = ET.Element(f"{{{W_NS}}}updateFields")
        update_node.set(f"{{{W_NS}}}val", "true")
        root.append(update_node)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def ensure_image_content_types_xml(content_types_data: bytes, image_extensions: set[str]) -> bytes:
    if not image_extensions:
        return content_types_data
    ns = "http://schemas.openxmlformats.org/package/2006/content-types"
    root = ET.fromstring(content_types_data)
    existing = {node.get("Extension", "").lower() for node in root.findall(f"{{{ns}}}Default")}
    for ext in sorted(image_extensions):
        content_type = IMAGE_CONTENT_TYPES.get(ext)
        if not content_type or ext in existing:
            continue
        node = ET.Element(f"{{{ns}}}Default")
        node.set("Extension", ext)
        node.set("ContentType", content_type)
        root.append(node)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def append_image_relationships_xml(rels_data: bytes, media_manager: MediaManager | None) -> bytes:
    if media_manager is None or not media_manager.images:
        return rels_data
    root = ET.fromstring(rels_data)
    for item in media_manager.images:
        rel = ET.Element(f"{{{PKG_REL_NS}}}Relationship")
        rel.set("Id", item.rel_id)
        rel.set("Type", IMAGE_REL_TYPE)
        rel.set("Target", item.part_name)
        root.append(rel)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def sanitize_template_styles_xml(styles_data: bytes) -> bytes:
    root = ET.fromstring(styles_data)
    ns = {"w": W_NS}

    # Some school templates assign the code block style an outline level of 0,
    # which makes Word include code paragraphs in the TOC. Force it out of the
    # collected outline range.
    code_style = root.find('w:style[@w:styleId="a9"]', ns)
    if code_style is not None:
        ppr = code_style.find("w:pPr", ns)
        if ppr is None:
            ppr = ET.SubElement(code_style, f"{{{W_NS}}}pPr")
        outline = ppr.find("w:outlineLvl", ns)
        if outline is None:
            outline = ET.SubElement(ppr, f"{{{W_NS}}}outlineLvl")
        outline.set(f"{{{W_NS}}}val", "9")

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def build_template_document(
    text: str,
    profile: dict[str, object],
    template_path: Path,
    *,
    math_converter: MathConverter | None = None,
    reference_anchors: dict[str, str] | None = None,
    markdown_dir: Path | None = None,
    cover_assets_dir: Path | None = None,
    media_manager: MediaManager | None = None,
) -> tuple[list[str], str, str]:
    markdown_title, front_sections, body_text = parse_markdown_document(text)
    cover_info = parse_cover_info(front_sections.get("封面信息", ""))
    thesis_title = cover_info.get("论文题目") or markdown_title
    sections = extract_template_section_properties(template_path)
    cover_sect = sections[0] if len(sections) >= 1 else default_sect_pr_xml()
    front_sect = sections[1] if len(sections) >= 2 else default_sect_pr_xml()
    body_sect = sections[2] if len(sections) >= 3 else (sections[-1] if sections else default_sect_pr_xml())

    elements: list[str] = []
    elements.extend(
        build_cover_elements(
            thesis_title,
            cover_info,
            cover_assets_dir=cover_assets_dir,
            media_manager=media_manager,
        )
    )
    elements.append(section_break_paragraph_xml(cover_sect))

    declaration = front_sections.get("声明", "").strip()
    if declaration:
        elements.append(build_front_heading("声  明"))
        for paragraph in split_plain_paragraphs(declaration):
            elements.append(build_body_paragraph(paragraph, math_converter=math_converter, reference_anchors=reference_anchors))
        elements.append(page_break_xml())

    cn_abstract, cn_keywords = extract_abstract_and_keywords(front_sections.get("摘要", ""), "关键词：")
    if cn_abstract or cn_keywords:
        elements.append(build_front_heading("摘  要"))
        for paragraph in cn_abstract:
            elements.append(build_body_paragraph(paragraph, math_converter=math_converter, reference_anchors=reference_anchors))
        keyword_paragraph = build_keyword_paragraph(cn_keywords)
        if keyword_paragraph:
            elements.append(keyword_paragraph)
        elements.append(page_break_xml())

    en_abstract, en_keywords = extract_abstract_and_keywords(front_sections.get("ABSTRACT", ""), "KEY WORDS:")
    if en_abstract or en_keywords:
        elements.append(build_front_heading("ABSTRACT", english=True))
        for paragraph in en_abstract:
            elements.append(
                build_body_paragraph(
                    paragraph,
                    english=True,
                    math_converter=math_converter,
                    reference_anchors=reference_anchors,
                )
            )
        keyword_paragraph = build_keyword_paragraph(en_keywords, english=True)
        if keyword_paragraph:
            elements.append(keyword_paragraph)
        elements.append(page_break_xml())

    elements.append(build_front_heading("目  录", toc=True))
    elements.append(toc_field_paragraph_xml())
    elements.append(section_break_paragraph_xml(front_sect))

    body_elements = build_document_elements(
        body_text,
        profile=profile,
        treat_first_heading_as_title=False,
        math_converter=math_converter,
        reference_anchors=reference_anchors,
        markdown_dir=markdown_dir,
        media_manager=media_manager,
    )
    elements.extend(body_elements)
    return elements, body_sect, thesis_title


def write_docx(
    markdown_path: Path,
    output_path: Path,
    *,
    template_path: Path | None = None,
    cover_assets_dir: Path | None = None,
    use_cover_assets: bool = True,
    enable_formula_conversion: bool = True,
) -> None:
    text = markdown_path.read_text(encoding="utf-8")
    resolved_cover_assets_dir = resolve_cover_assets_dir(markdown_path, cover_assets_dir, use_cover_assets=use_cover_assets)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    math_converter = MathConverter() if enable_formula_conversion else None
    if math_converter is not None:
        math_converter.preload_from_markdown(text)
    reference_anchors = extract_reference_anchors(text)
    if template_path:
        with zipfile.ZipFile(template_path) as template_zip:
            try:
                template_rels_data = template_zip.read("word/_rels/document.xml.rels")
            except KeyError:
                template_rels_data = b""
            try:
                template_content_types = template_zip.read("[Content_Types].xml")
            except KeyError:
                template_content_types = b""
        media_manager = MediaManager(
            starting_rid=next_relationship_id(template_rels_data) if template_rels_data else 2,
            starting_image_index=next_image_index_from_template(template_path),
        )
        profile: dict[str, object] = {
            "title": "af2",
            "heading1": "1",
            "heading2": "2",
            "heading3": "3",
            "normal": "a0",
            "quote": "a0",
            "code": "a9",
            "code_ppr_extra": '<w:outlineLvl w:val="9"/>',
            "math": "af8",
            "table": "a0",
            "normal_first_line_chars": 200,
            "normal_first_line": 480,
            "normal_ppr_extra": spacing_xml(line=360),
            "normal_run": {
                "font_ascii": "Times New Roman",
                "font_hansi": "Times New Roman",
                "font_eastasia": "宋体",
                "size": 24,
            },
            "skip_reference_notes": True,
            "strip_heading_numbers": True,
        }
        elements, sect_pr, doc_title = build_template_document(
            text,
            profile,
            template_path,
            math_converter=math_converter,
            reference_anchors=reference_anchors,
            markdown_dir=markdown_path.parent,
            cover_assets_dir=resolved_cover_assets_dir,
            media_manager=media_manager,
        )
        with zipfile.ZipFile(template_path) as src, zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as dst:
            for item in src.infolist():
                data = src.read(item.filename)
                if item.filename == "word/document.xml":
                    data = document_xml(elements, sect_pr=sect_pr).encode("utf-8")
                elif item.filename == "docProps/core.xml":
                    data = core_xml(doc_title).encode("utf-8")
                elif item.filename == "word/settings.xml":
                    data = ensure_update_fields_xml(data)
                elif item.filename == "word/styles.xml":
                    data = sanitize_template_styles_xml(data)
                elif item.filename == "word/_rels/document.xml.rels":
                    data = append_image_relationships_xml(data, media_manager)
                elif item.filename == "[Content_Types].xml":
                    data = ensure_image_content_types_xml(data, media_manager.image_extensions())
                dst.writestr(item, data)
            if not template_rels_data:
                dst.writestr("word/_rels/document.xml.rels", document_rels_xml(media_manager).encode("utf-8"))
            if not template_content_types:
                dst.writestr("[Content_Types].xml", content_types_xml(media_manager.image_extensions()))
            for image in media_manager.images:
                dst.writestr(f"word/{image.part_name}", image.source_path.read_bytes())
    else:
        media_manager = MediaManager(starting_rid=2)
        elements = build_document_elements(
            text,
            math_converter=math_converter,
            reference_anchors=reference_anchors,
            markdown_dir=markdown_path.parent,
            media_manager=media_manager,
        )
        with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("[Content_Types].xml", content_types_xml(media_manager.image_extensions()))
            zf.writestr("_rels/.rels", rels_xml())
            zf.writestr("docProps/core.xml", core_xml(markdown_path.stem))
            zf.writestr("docProps/app.xml", app_xml())
            zf.writestr("word/document.xml", document_xml(elements))
            zf.writestr("word/styles.xml", styles_xml())
            zf.writestr("word/_rels/document.xml.rels", document_rels_xml(media_manager))
            for image in media_manager.images:
                zf.writestr(f"word/{image.part_name}", image.source_path.read_bytes())

    if math_converter is not None:
        math_converter.emit_warning()


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Convert a Xinjiang University thesis-style Markdown document to DOCX."
    )
    parser.add_argument("input", type=Path)
    parser.add_argument("output", nargs="?", type=Path, default=None)
    parser.add_argument(
        "--template",
        type=Path,
        default=DEFAULT_TEMPLATE_PATH if DEFAULT_TEMPLATE_PATH.exists() else None,
        help="DOCX template used for section properties and Word styles.",
    )
    parser.add_argument(
        "--no-template",
        action="store_true",
        help="Generate a plain DOCX without inheriting styles from a DOCX template.",
    )
    parser.add_argument(
        "--assets-dir",
        type=Path,
        default=DEFAULT_COVER_ASSETS_DIR if DEFAULT_COVER_ASSETS_DIR.exists() else None,
        help="Directory containing cover assets such as xju-emblem.jpeg and xju-wordmark.png.",
    )
    parser.add_argument(
        "--no-cover-assets",
        action="store_true",
        help="Disable cover logos even if asset files are available.",
    )
    parser.add_argument(
        "--no-formula-conversion",
        action="store_true",
        help="Disable LaTeX-to-OMML conversion and keep formulas as plain LaTeX text.",
    )
    args = parser.parse_args()
    output_path = args.output or args.input.with_suffix(".docx")
    template_path = None if args.no_template else args.template
    write_docx(
        args.input,
        output_path,
        template_path=template_path,
        cover_assets_dir=args.assets_dir,
        use_cover_assets=not args.no_cover_assets,
        enable_formula_conversion=not args.no_formula_conversion,
    )


if __name__ == "__main__":
    main()
