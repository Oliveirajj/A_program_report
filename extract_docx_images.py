"""
Extract images embedded in the DOCX report template and save them with filenames
that reflect the chapter hierarchy.
"""
from __future__ import annotations

import json
import re
import unicodedata
from collections import defaultdict
from itertools import count
from pathlib import Path
from typing import Dict, Iterable, List, Tuple, Union

from docx import Document as load_document
from docx.document import Document as DocumentType
from docx.oxml.ns import qn
from docx.table import Table, _Cell
from docx.text.paragraph import Paragraph

ROOT_DIR = Path(__file__).resolve().parent
DOCX_PATH = ROOT_DIR / "A项目资料补充_251111" / "附件6_报告模板.docx"
OUTPUT_DIR = ROOT_DIR / "docx_exports" / "report_template"
MANIFEST_PATH = OUTPUT_DIR / "image_manifest.json"

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"


def iter_block_items(parent: Union[DocumentType, _Cell]) -> Iterable[Union[Paragraph, Table]]:
    """Yield paragraphs and tables in document order (recursive for cells)."""
    if isinstance(parent, DocumentType):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc  # type: ignore[attr-defined]
    else:
        raise TypeError(f"Unsupported parent type: {type(parent)!r}")

    for child in parent_elm.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, parent)  # type: ignore[arg-type]
        elif child.tag == qn("w:tbl"):
            yield Table(child, parent)  # type: ignore[arg-type]


def detect_heading_level(paragraph: Paragraph) -> Tuple[int, str] | None:
    """
    Return heading level (1-based) and cleaned text if the paragraph is a heading.

    The detection relies on style names such as "Heading 1", "标题 1", etc.
    """
    text = paragraph.text.strip()
    if not text:
        return None

    style_name = paragraph.style.name if paragraph.style else ""
    if not style_name:
        style_name = ""

    if "Heading" in style_name:
        match = re.search(r"(\d+)", style_name)
        level = int(match.group(1)) if match else 1
        return level, text
    if "标题" in style_name:
        match = re.search(r"(\d+)", style_name)
        level = int(match.group(1)) if match else 1
        return level, text

    # Fallback to textual heuristics
    if re.match(r"^第[一二三四五六七八九十百零]+章", text):
        return 1, text

    numeric_path = re.match(r"^(\d+(?:\.\d+)+)", text)
    if numeric_path:
        level = numeric_path.group(1).count(".") + 1
        return level, text

    if re.match(r"^[\(\[（【]?\d+[\)\]）】]?[．.、-]", text):
        return 2, text

    if re.match(r"^[一二三四五六七八九十]+[、.．]", text):
        return 2, text

    return None


def sanitize_segment(text: str, max_length: int = 30) -> str:
    """
    Produce a filesystem-safe slug that still keeps recognizable characters.

    Chinese characters are kept; other invalid characters are removed. If the
    slug would be empty, a fallback hash is returned.
    """
    normalized = unicodedata.normalize("NFKC", text)
    cleaned = re.sub(r'[\\/:*?"<>|]+', "_", normalized).strip("_")
    if len(cleaned) > max_length:
        cleaned = cleaned[: max_length].rstrip("_")
    if not cleaned:
        cleaned = f"segment_{abs(hash(text)) & 0xFFFF:04x}"
    return cleaned


def current_heading_code(counts: Dict[int, int]) -> str:
    parts = []
    for level in sorted(counts):
        value = counts[level]
        if value <= 0:
            continue
        parts.append(f"{value:02d}")
    return "ch" + "-".join(parts) if parts else "ch00"


def export_docx_images(doc_path: Path, output_dir: Path, manifest_path: Path) -> None:
    if not doc_path.exists():
        raise FileNotFoundError(f"DOCX template not found: {doc_path}")

    output_dir.mkdir(parents=True, exist_ok=True)

    document = load_document(doc_path)

    heading_counts: Dict[int, int] = defaultdict(int)
    heading_titles: Dict[int, str] = {}

    manifest_entries: List[Dict[str, object]] = []

    occurrence_counter = defaultdict(int)
    global_counter = count(1)

    def export_picture(r_id: str, chapter_code: str, chapter_title: str) -> Path:
        part = document.part.related_parts[r_id]
        suffix = Path(part.partname).suffix or ".png"
        occurrence_counter[r_id] += 1
        idx = next(global_counter)
        base_name = sanitize_segment(chapter_title) if chapter_title else "document"
        file_name = f"{idx:03d}_{chapter_code}_{base_name}_{occurrence_counter[r_id]:02d}{suffix}"
        dest_path = output_dir / file_name
        dest_path.write_bytes(part.blob)
        return dest_path

    def process_paragraph(paragraph: Paragraph, chapter_code: str, chapter_title: str) -> None:
        for drawing in paragraph._p.iter():
            if drawing.tag != f"{{{PIC_NS}}}pic":
                continue
            blip = list(drawing.iterfind(f".//{{{A_NS}}}blip"))
            if not blip:
                continue
            r_id = blip[0].get(qn("r:embed"))
            if not r_id:
                continue
            output_path = export_picture(r_id, chapter_code, chapter_title)
            manifest_entries.append(
                {
                    "file": output_path.name,
                    "chapter_code": chapter_code,
                    "chapter_title": chapter_title,
                    "paragraph_text": paragraph.text.strip(),
                }
            )

    for block in iter_block_items(document):
        if isinstance(block, Paragraph):
            heading_info = detect_heading_level(block)
            if heading_info:
                level, title = heading_info
                heading_counts[level] += 1
                for higher_level in list(heading_counts.keys()):
                    if higher_level > level:
                        heading_counts[higher_level] = 0
                        heading_titles.pop(higher_level, None)
                heading_titles[level] = title
                continue

            chapter_title = " / ".join(heading_titles[level] for level in sorted(heading_titles))
            chapter_code = current_heading_code(heading_counts)
            process_paragraph(block, chapter_code, chapter_title)
        elif isinstance(block, Table):
            chapter_title = " / ".join(heading_titles[level] for level in sorted(heading_titles))
            chapter_code = current_heading_code(heading_counts)
            for row in block.rows:
                for cell in row.cells:
                    for cell_block in iter_block_items(cell):
                        if isinstance(cell_block, Paragraph):
                            process_paragraph(cell_block, chapter_code, chapter_title)

    if manifest_entries:
        manifest_path.write_text(json.dumps(manifest_entries, ensure_ascii=False, indent=2), encoding="utf-8")


if __name__ == "__main__":
    export_docx_images(DOCX_PATH, OUTPUT_DIR, MANIFEST_PATH)

