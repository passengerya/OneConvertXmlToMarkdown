#!/usr/bin/env python3
from __future__ import annotations

import argparse
import base64
import hashlib
import html
import os
import re
import shutil
import sys
from xml.sax.saxutils import escape as xml_escape
from dataclasses import dataclass, field
from html.parser import HTMLParser
from pathlib import Path
from typing import Iterable, Iterator
import xml.etree.ElementTree as ET


NS = {"one": "http://schemas.microsoft.com/office/onenote/2013/onenote"}
ONE_NS = "{http://schemas.microsoft.com/office/onenote/2013/onenote}"


def one(tag: str) -> str:
    return f"{ONE_NS}{tag}"


def tag_name(elem: ET.Element) -> str:
    if elem.tag.startswith("{"):
        return elem.tag.split("}", 1)[1]
    return elem.tag


def normalize_space(text: str) -> str:
    if not text:
        return ""
    text = text.replace("\xa0", " ")
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    lines = [re.sub(r"[ \t]+", " ", line).strip() for line in text.split("\n")]
    # collapse excessive blank lines
    out: list[str] = []
    blank = False
    for line in lines:
        if line == "":
            if not blank:
                out.append("")
            blank = True
        else:
            out.append(line)
            blank = False
    return "\n".join(out).strip()


def strip_markup(text: str) -> str:
    text = re.sub(r"<font[^>]*>", "", text, flags=re.I)
    text = re.sub(r"</font>", "", text, flags=re.I)
    text = text.replace("<br>", "\n")
    return normalize_space(text)


def looks_short_title(text: str) -> bool:
    if not text:
        return False
    t = strip_markup(text)
    if not t:
        return False
    if re.search(r"!\[[^\]]*\]\(", t):
        return False
    return len(t) <= 32 and "\n" not in t


def looks_heading_label(text: str) -> bool:
    if not text:
        return False
    t = strip_markup(text).replace("\n", " ").strip()
    if re.search(r"!\[[^\]]*\]\(", t):
        return False
    return bool(t) and len(t) <= 100


PROPERTY_LABEL_KEYWORDS = (
    "特点",
    "概念",
    "组成",
    "比例",
    "原理",
    "应用",
    "用途",
    "优点",
    "缺点",
    "分类",
    "影响",
    "评价",
    "质量控制",
    "注意事项",
    "参数说明",
    "染色效果",
    "染色成分",
    "储存",
    "pH值",
)


def is_property_like_label(text: str) -> bool:
    t = strip_markup(text).replace("\n", " ").strip("：: ").strip()
    if not t:
        return False
    if any(k in t for k in PROPERTY_LABEL_KEYWORDS):
        return True
    # Short generic field labels are more likely KV keys than section titles.
    if len(t) <= 8 and t in {"检查项目", "采血部位", "操作步骤", "采血顺序", "发生溶血的原因", "溶血对结果的影响"}:
        return True
    return False


def escape_md_table_cell(text: str) -> str:
    text = text.replace("|", "\\|")
    return text


def strip_leading_colored_dash_markup(text: str) -> str:
    # OneNote sometimes exports visual bullets as a colored "-" span before actual content.
    text = re.sub(r"^(?:<font[^>]*>\s*-\s*</font>\s*)+", "", text, flags=re.I | re.S)
    text = re.sub(r"^\s*-\s*</font>\s*", "", text, flags=re.I | re.S)
    text = re.sub(r"^(?:</font>\s*)+", "", text, flags=re.I | re.S)
    return text


def chinese_numeral(n: int) -> str:
    nums = "零一二三四五六七八九十"
    if 0 <= n <= 10:
        return nums[n]
    if n < 20:
        return "十" + nums[n - 10]
    if n < 100:
        tens, ones = divmod(n, 10)
        if ones == 0:
            return nums[tens] + "十"
        return nums[tens] + "十" + nums[ones]
    return str(n)


def chinese_text_to_int(text: str) -> int | None:
    text = text.strip()
    if not text:
        return None
    digit_map = {
        "零": 0,
        "〇": 0,
        "一": 1,
        "二": 2,
        "两": 2,
        "三": 3,
        "四": 4,
        "五": 5,
        "六": 6,
        "七": 7,
        "八": 8,
        "九": 9,
    }
    if text == "十":
        return 10
    if "十" in text:
        left, right = text.split("十", 1)
        tens = 1 if left == "" else digit_map.get(left)
        ones = 0 if right == "" else digit_map.get(right)
        if tens is None or ones is None:
            return None
        return tens * 10 + ones
    if len(text) == 1 and text in digit_map:
        return digit_map[text]
    return None


def section_sort_key(path: Path) -> tuple[int, int, str]:
    stem = path.stem
    m = re.search(r"第([零〇一二两三四五六七八九十]+)节", stem)
    if m:
        n = chinese_text_to_int(m.group(1))
        if n is not None:
            return (0, n, stem)
    return (1, 0, stem)


def split_circled_number_lines(text: str) -> list[str] | None:
    t = normalize_space(text)
    if not t:
        return None
    if not re.search(r"[①②③④⑤⑥⑦⑧⑨⑩]", t):
        return None
    parts = re.split(r"(?=[①②③④⑤⑥⑦⑧⑨⑩])", t)
    parts = [p.strip() for p in parts if p.strip()]
    if len(parts) <= 1:
        return None
    if not all(re.match(r"^[①②③④⑤⑥⑦⑧⑨⑩]", p) for p in parts):
        return None
    return parts


FORMAT_TAGS = {"b", "i", "u", "s", "strong", "em", "code", "mark"}

class RichTextHTMLParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.parts: list[str] = []
        self.tag_stack: list[str] = []      # open inline tags for proper nesting
        self.color_stack: list[str | None] = [None]

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        t = tag.lower()
        if t == "br":
            self.parts.append("\n")
            return
        if t in FORMAT_TAGS:
            self.parts.append(f"<{t}>")
            self.tag_stack.append(t)
            self.color_stack.append(self.color_stack[-1])
            return
        if t == "span":
            color = self.color_stack[-1]
            style_map = {}
            for k, v in attrs:
                if k and v and k.lower() == "style":
                    for item in v.split(";"):
                        if ":" in item:
                            kk, vv = item.split(":", 1)
                            style_map[kk.strip().lower()] = vv.strip()
            if "color" in style_map:
                color = style_map["color"]
            if "background-color" in style_map:
                bg = style_map["background-color"]
                self.parts.append(f'<span style="background-color:{bg}">')
                self.tag_stack.append("span")
            self.color_stack.append(color)
            return
        self.color_stack.append(self.color_stack[-1])

    def handle_endtag(self, tag: str) -> None:
        t = tag.lower()
        if t in FORMAT_TAGS and self.tag_stack and self.tag_stack[-1] == t:
            self.parts.append(f"</{t}>")
            self.tag_stack.pop()
            if self.color_stack:
                self.color_stack.pop()
        elif t == "span" and self.tag_stack and self.tag_stack[-1] == "span":
            self.parts.append("</span>")
            self.tag_stack.pop()
            if self.color_stack:
                self.color_stack.pop()
        else:
            if self.color_stack:
                self.color_stack.pop()
        if not self.color_stack:
            self.color_stack = [None]

    def handle_data(self, data: str) -> None:
        if not data:
            return
        color = normalize_color(self.color_stack[-1])
        text = data.replace("\xa0", " ")
        if not text:
            return
        if color and color not in {"#ffffff", "#fff", "white", "automatic", "#000000", "#000", "black"}:
            md_color = map_color(color)
            self.parts.append(f'<font color="{md_color}">{text}</font>')
        else:
            self.parts.append(text)

    def get_text(self) -> str:
        text = "".join(self.parts)
        text = text.replace("<br/>", "\n")
        text = text.replace("<br />", "\n")
        return normalize_space(text)


def normalize_color(color: str | None) -> str | None:
    if color is None:
        return None
    c = color.strip().strip("'\"").lower()
    if not c:
        return None
    if c.startswith("#") and len(c) == 7:
        return c
    return c


def map_color(color: str) -> str:
    c = normalize_color(color) or color
    if c == "yellow":
        return "#ff0000"
    return c.lower()


def html_fragment_to_markdown_text(fragment: str) -> str:
    if fragment is None:
        return ""
    wrapped = f"<root>{fragment}</root>"
    parser = RichTextHTMLParser()
    try:
        parser.feed(wrapped)
        parser.close()
        return parser.get_text()
    except Exception:
        # Fallback: keep text content only if HTML is malformed.
        text = fragment
        text = re.sub(r"(?i)<br\s*/?>", "\n", text)
        text = re.sub(r"<[^>]+>", "", text)
        text = html.unescape(text)
        return normalize_space(text)


@dataclass
class ImageRef:
    filename: str
    link_path: str
    width: float | None = None
    height: float | None = None
    callback_id: str | None = None
    is_placeholder: bool = False


class SequentialImageResolver:
    def __init__(
        self,
        attachment_dir: Path | None,
        sort_key: str = "name",
        attachment_link_prefix: str | None = None,
        copy_attachment_dir: Path | None = None,
        placeholder_dir: Path | None = None,
        placeholder_link_prefix: str = "_generated_images",
        semantic_image_names: bool = True,
    ) -> None:
        self.attachment_dir = attachment_dir
        self.sort_key = sort_key
        self.attachment_link_prefix = (attachment_link_prefix or "").replace("\\", "/").rstrip("/")
        self.copy_attachment_dir = copy_attachment_dir
        self.placeholder_dir = placeholder_dir
        self.placeholder_link_prefix = placeholder_link_prefix.replace("\\", "/").rstrip("/")
        self.semantic_image_names = semantic_image_names
        self.files: list[Path] = []
        self.index = 0
        self.by_callback: dict[str, tuple[str, str, bool]] = {}
        # Content-hash dedupe for XML-embedded images: same bytes -> same exported file.
        self.by_embedded_digest: dict[str, tuple[str, str]] = {}
        self.placeholder_count = 0
        self.copied_attachment_count = 0
        self.extracted_from_xml_count = 0
        self._used_output_names: set[str] = set()
        if self.placeholder_dir is not None:
            self.placeholder_dir.mkdir(parents=True, exist_ok=True)
        if self.copy_attachment_dir is not None:
            self.copy_attachment_dir.mkdir(parents=True, exist_ok=True)
        if attachment_dir and attachment_dir.exists():
            files = [p for p in attachment_dir.iterdir() if p.is_file()]
            if sort_key == "mtime":
                files.sort(key=lambda p: (p.stat().st_mtime, p.name))
            else:
                files.sort(key=lambda p: p.name)
            self.files = files

    def resolve(self, image_el: ET.Element) -> ImageRef:
        size_el = image_el.find("one:Size", NS)
        width = float(size_el.get("width")) if size_el is not None and size_el.get("width") else None
        height = float(size_el.get("height")) if size_el is not None and size_el.get("height") else None
        cb = image_el.find("one:CallbackID", NS)
        callback_id = cb.get("callbackID") if cb is not None else None
        if callback_id and callback_id in self.by_callback:
            name, link_path, is_placeholder = self.by_callback[callback_id]
            return ImageRef(name, link_path, width, height, callback_id, is_placeholder=is_placeholder)
        embedded = self._extract_embedded_image(image_el)
        if embedded is not None:
            data, ext = embedded
            digest_key = f"{ext}:{hashlib.sha256(data).hexdigest()}"
            if digest_key in self.by_embedded_digest:
                name, link_path = self.by_embedded_digest[digest_key]
                if callback_id:
                    self.by_callback[callback_id] = (name, link_path, False)
                return ImageRef(name, link_path, width, height, callback_id, is_placeholder=False)
            name = self._embedded_output_name(image_el, ext)
            link_path = self._save_generated_binary(name, data)
            self.by_embedded_digest[digest_key] = (name, link_path)
            if callback_id:
                self.by_callback[callback_id] = (name, link_path, False)
            return ImageRef(name, link_path, width, height, callback_id, is_placeholder=False)
        if self.index < len(self.files):
            src = self.files[self.index]
            name = src.name
            self.index += 1
            if self.copy_attachment_dir is not None:
                name = self._attachment_output_name(src, image_el)
                self._ensure_attachment_copied(src, name)
            if callback_id:
                self.by_callback[callback_id] = (name, self._link_for_attachment(name), False)
            return ImageRef(name, self._link_for_attachment(name), width, height, callback_id, is_placeholder=False)
        placeholder = self._placeholder_output_name(image_el)
        self.index += 1
        self.placeholder_count += 1
        self._ensure_placeholder_svg(placeholder, image_el, width, height, callback_id)
        if callback_id:
            self.by_callback[callback_id] = (placeholder, self._link_for_placeholder(placeholder), True)
        return ImageRef(placeholder, self._link_for_placeholder(placeholder), width, height, callback_id, is_placeholder=True)

    def _make_placeholder_name(self) -> str:
        return f"image_{self.index + 1}.svg"

    def _link_for_attachment(self, name: str) -> str:
        if self.attachment_link_prefix:
            return f"{self.attachment_link_prefix}/{name}".replace("\\", "/")
        return name

    def _link_for_placeholder(self, name: str) -> str:
        if self.placeholder_link_prefix:
            return f"{self.placeholder_link_prefix}/{name}".replace("\\", "/")
        return name

    def _link_for_generated(self, name: str) -> str:
        return self._link_for_placeholder(name)

    def _extract_ocr_text(self, image_el: ET.Element) -> str:
        ocr_text = image_el.find("one:OCRData/one:OCRText", NS)
        if ocr_text is not None and ocr_text.text:
            return normalize_space(ocr_text.text)
        return ""

    def _mime_to_ext(self, mime: str) -> str | None:
        m = mime.lower().strip()
        mapping = {
            "image/png": ".png",
            "image/jpeg": ".jpg",
            "image/jpg": ".jpg",
            "image/gif": ".gif",
            "image/bmp": ".bmp",
            "image/tiff": ".tif",
            "image/webp": ".webp",
            "image/svg+xml": ".svg",
        }
        return mapping.get(m)

    def _sniff_image_ext(self, data: bytes) -> str:
        if data.startswith(b"\x89PNG\r\n\x1a\n"):
            return ".png"
        if data.startswith(b"\xff\xd8\xff"):
            return ".jpg"
        if data.startswith((b"GIF87a", b"GIF89a")):
            return ".gif"
        if data.startswith(b"BM"):
            return ".bmp"
        if data.startswith(b"RIFF") and b"WEBP" in data[:16]:
            return ".webp"
        if data.startswith((b"II*\x00", b"MM\x00*")):
            return ".tif"
        if data.lstrip().startswith(b"<svg"):
            return ".svg"
        return ".bin"

    def _extract_embedded_image(self, image_el: ET.Element) -> tuple[bytes, str] | None:
        # Some OneNote XML exports embed image bytes; this dataset mostly uses callback references only.
        data_el = image_el.find("one:Data", NS)
        if data_el is None:
            data_el = image_el.find("one:ImageData", NS)
        if data_el is None:
            data_el = image_el.find("one:BinaryData", NS)
        if data_el is None or not data_el.text or not data_el.text.strip():
            return None
        raw_text = re.sub(r"\s+", "", data_el.text)
        try:
            data = base64.b64decode(raw_text, validate=False)
        except Exception:
            return None
        if not data:
            return None
        mime_el = image_el.find("one:MimeType", NS)
        if mime_el is None:
            mime_el = image_el.find("one:ContentType", NS)
        ext = None
        if mime_el is not None and mime_el.text:
            ext = self._mime_to_ext(mime_el.text)
        if not ext:
            fmt = image_el.get("format") or image_el.get("Format") or ""
            if fmt:
                fmt = fmt.strip().lower().lstrip(".")
                if fmt in {"png", "jpg", "jpeg", "gif", "bmp", "tif", "tiff", "webp", "svg"}:
                    ext = "." + ("jpg" if fmt == "jpeg" else "tif" if fmt == "tiff" else fmt)
        if not ext:
            ext = self._sniff_image_ext(data)
        return data, ext

    def _normalize_filename_stem(self, text: str) -> str:
        text = normalize_space(text).replace("\n", " ").strip()
        # Keep Chinese/letters/digits, map everything else to hyphen.
        text = re.sub(r"[^\u4e00-\u9fffA-Za-z0-9]+", "-", text)
        text = re.sub(r"-{2,}", "-", text).strip("-._ ")
        if len(text) > 36:
            text = text[:36].rstrip("-._ ")
        return text

    def _is_meaningful_theme_stem(self, stem: str) -> bool:
        if not stem or len(stem) < 2:
            return False
        letters = re.findall(r"[A-Za-z]", stem)
        chinese = re.findall(r"[\u4e00-\u9fff]", stem)
        digits = re.findall(r"\d", stem)
        core_chars = re.findall(r"[\u4e00-\u9fffA-Za-z0-9]", stem)
        if not core_chars:
            return False
        if len(chinese) + len(letters) < 2:
            return False
        if not chinese and len(letters) < 4:
            return False
        if len(chinese) + len(letters) <= 2 and len(digits) >= max(1, len(core_chars) // 2):
            return False
        if len(digits) == len(core_chars):
            return False
        if stem.count("口") >= 2 and stem.count("口") >= max(2, len(core_chars) // 2):
            return False
        compressed = stem.replace("-", "")
        if compressed and len(set(compressed)) <= 2 and len(compressed) >= 4:
            return False
        return True

    def _theme_stem_from_image(self, image_el: ET.Element, fallback_prefix: str) -> str:
        ocr = self._extract_ocr_text(image_el)
        candidate_lines = [ln.strip() for ln in ocr.split("\n") if ln.strip()]
        # Prefer the first non-trivial line.
        candidate = ""
        for ln in candidate_lines:
            norm = self._normalize_filename_stem(ln)
            if self._is_meaningful_theme_stem(norm):
                candidate = norm
                break
        if not candidate and ocr:
            merged = self._normalize_filename_stem(ocr)
            if self._is_meaningful_theme_stem(merged):
                candidate = merged
        if not candidate:
            candidate = f"{fallback_prefix}-{self.index + 1}"
        return candidate

    def _unique_output_name(self, stem: str, ext: str) -> str:
        stem = self._normalize_filename_stem(stem) or f"image-{self.index + 1}"
        ext = ext if ext.startswith(".") else f".{ext}"
        base = f"{stem}{ext}"
        key = base.lower()
        if key not in self._used_output_names:
            self._used_output_names.add(key)
            return base
        n = 2
        while True:
            cand = f"{stem}-{n}{ext}"
            k = cand.lower()
            if k not in self._used_output_names:
                self._used_output_names.add(k)
                return cand
            n += 1

    def _attachment_output_name(self, src: Path, image_el: ET.Element) -> str:
        if not self.semantic_image_names:
            return src.name
        stem = self._theme_stem_from_image(image_el, "image")
        return self._unique_output_name(stem, src.suffix or ".png")

    def _embedded_output_name(self, image_el: ET.Element, ext: str) -> str:
        if not self.semantic_image_names:
            return self._unique_output_name(f"image-{self.index + 1}", ext)
        stem = self._theme_stem_from_image(image_el, "image")
        return self._unique_output_name(stem, ext)

    def _placeholder_output_name(self, image_el: ET.Element) -> str:
        if not self.semantic_image_names:
            return self._unique_output_name(f"image-{self.index + 1}", ".svg")
        stem = self._theme_stem_from_image(image_el, "image")
        return self._unique_output_name(stem, ".svg")

    def _ensure_attachment_copied(self, src: Path, dst_name: str) -> None:
        if self.copy_attachment_dir is None:
            return
        dst = self.copy_attachment_dir / dst_name
        if dst.exists():
            return
        shutil.copy2(src, dst)
        self.copied_attachment_count += 1

    def _save_generated_binary(self, name: str, data: bytes) -> str:
        # Store XML-embedded image bytes alongside generated placeholders so markdown always has a file target.
        if self.placeholder_dir is None:
            raise RuntimeError("No generated asset directory configured for XML-embedded images")
        path = self.placeholder_dir / name
        if not path.exists():
            path.write_bytes(data)
            self.extracted_from_xml_count += 1
        return self._link_for_generated(name)

    def _wrap_text(self, text: str, max_len: int = 22, max_lines: int = 6) -> list[str]:
        if not text:
            return []
        lines: list[str] = []
        for raw in text.split("\n"):
            raw = raw.strip()
            if not raw:
                continue
            while len(raw) > max_len:
                lines.append(raw[:max_len])
                raw = raw[max_len:]
                if len(lines) >= max_lines:
                    return lines[:max_lines]
            if raw:
                lines.append(raw)
                if len(lines) >= max_lines:
                    return lines[:max_lines]
        return lines[:max_lines]

    def _ensure_placeholder_svg(
        self,
        name: str,
        image_el: ET.Element,
        width: float | None,
        height: float | None,
        callback_id: str | None,
    ) -> None:
        if self.placeholder_dir is None:
            return
        path = self.placeholder_dir / name
        if path.exists():
            return
        w = max(140, int(round(width or 260)))
        h = max(90, int(round(height or 160)))
        ocr = self._extract_ocr_text(image_el)
        lines = self._wrap_text(ocr) or ["Image not available"]
        if callback_id:
            cid = callback_id[-10:] if len(callback_id) > 10 else callback_id
            lines.append(f"id: {cid}")
        text_y = 24
        svg_lines = []
        for line in lines:
            svg_lines.append(
                f'<text x="12" y="{text_y}" fill="#334155" font-size="12" font-family="Segoe UI, Microsoft YaHei, sans-serif">{xml_escape(line)}</text>'
            )
            text_y += 18
            if text_y > h - 8:
                break
        svg = (
            f'<svg xmlns="http://www.w3.org/2000/svg" width="{w}" height="{h}" viewBox="0 0 {w} {h}">'
            f'<rect width="{w}" height="{h}" fill="#f8fafc" stroke="#cbd5e1"/>'
            f'<rect x="0" y="0" width="{w}" height="26" fill="#e2e8f0" />'
            f'<text x="10" y="17" fill="#0f172a" font-size="12" font-family="Segoe UI, Microsoft YaHei, sans-serif">OneNote image placeholder</text>'
            + "".join(svg_lines)
            + "</svg>"
        )
        path.write_text(svg, encoding="utf-8")


@dataclass
class RenderContext:
    table_depth: int = 0
    section_depth: int = 0
    list_indent: int = 0
    numbering_counter: list[int] = field(default_factory=list)

    def child_for_table(self) -> "RenderContext":
        return RenderContext(
            table_depth=self.table_depth + 1,
            section_depth=self.section_depth,
            list_indent=self.list_indent,
            numbering_counter=list(self.numbering_counter),
        )

    def with_section_depth(self, depth: int) -> "RenderContext":
        return RenderContext(
            table_depth=self.table_depth,
            section_depth=depth,
            list_indent=self.list_indent,
            numbering_counter=list(self.numbering_counter),
        )

    def with_list_indent(self, indent: int) -> "RenderContext":
        return RenderContext(
            table_depth=self.table_depth,
            section_depth=self.section_depth,
            list_indent=indent,
            numbering_counter=list(self.numbering_counter),
        )


class OneNoteXMLToMarkdownConverter:
    def __init__(self, image_resolver: SequentialImageResolver, image_syntax: str = "markdown") -> None:
        self.image_resolver = image_resolver
        self.image_syntax = image_syntax

    def convert_file(self, xml_path: Path) -> str:
        root = ET.parse(xml_path).getroot()
        page_title = root.get("name") or ""
        blocks: list[str] = []
        outlines = root.findall("one:Outline", NS)
        for outline in outlines:
            oe_children = outline.find("one:OEChildren", NS)
            if oe_children is None:
                continue
            outline_blocks = self.render_oe_children(oe_children, RenderContext())
            blocks.extend(outline_blocks)

        blocks = self.postprocess_blocks(blocks)
        # Normalize prose to bullet-list style used by this project.
        blocks = self.bulletize_body_blocks(blocks)
        if not any(b.strip() for b in blocks):
            title_text = normalize_space(page_title)
            if title_text:
                blocks = [f"# {title_text}"]
        md_text = "\n\n".join(b for b in blocks if b.strip())
        # Final formatting cleanup: remove invalid blank lines and normalize spacing.
        md_text = self.cleanup_markdown_text(md_text)
        return md_text + "\n"

    def postprocess_blocks(self, blocks: list[str]) -> list[str]:
        out: list[str] = []
        top_ol_index = 1
        i = 0
        while i < len(blocks):
            raw = blocks[i]
            b = raw.strip()
            if not b:
                i += 1
                continue
            if self.is_plain_single_line(b):
                # If next block is a table/list block, treat this as numbered item title.
                if i + 1 < len(blocks) and self.is_structured_block(blocks[i + 1]):
                    out.append(f"{top_ol_index}. {b}")
                    out.append(blocks[i + 1].strip("\n"))
                    top_ol_index += 1
                    i += 2
                    continue
            out.append(raw.strip("\n"))
            i += 1
        return out

    def bulletize_body_blocks(self, blocks: list[str]) -> list[str]:
        out: list[str] = []
        for raw in blocks:
            block = raw.strip("\n")
            if not block.strip():
                continue
            if not self.is_body_bullet_candidate_block(block):
                out.append(block)
                continue
            lines_out: list[str] = []
            for line in block.split("\n"):
                t = line.strip()
                if not t:
                    continue
                if self.is_list_item_line(t):
                    lines_out.append(t)
                else:
                    lines_out.append(f"- {t}")
            if lines_out:
                out.append("\n".join(lines_out))
        return out

    @staticmethod
    def is_body_bullet_candidate_block(text: str) -> bool:
        t = text.strip()
        if not t:
            return False
        if "\n\n" in t:
            return False
        first = t.lstrip()
        if first.startswith(("#", "|", "![", "> ", "```", "$$", "<table", "<img", "<svg")):
            return False
        lines = [ln.strip() for ln in t.split("\n") if ln.strip()]
        if not lines:
            return False
        for ln in lines:
            if ln.startswith(("#", "|", "![", "> ", "```", "$$")):
                return False
            if re.match(r"^\d+\. ", ln):
                return False
            if re.match(r"^[-*+] ", ln):
                return False
        return True

    @staticmethod
    def cleanup_markdown_text(md_text: str) -> str:
        text = md_text.replace("\r\n", "\n").replace("\r", "\n")
        text = re.sub(r"[ \t]+\n", "\n", text)

        # Final pass: convert remaining plain body lines to unordered list items.
        lines = text.split("\n")
        out_lines: list[str] = []
        in_fence = False
        for raw in lines:
            line = raw.rstrip()
            stripped = line.strip()
            lstripped = line.lstrip()

            if lstripped.startswith("```"):
                out_lines.append(line)
                in_fence = not in_fence
                continue
            if in_fence:
                out_lines.append(line)
                continue
            if not stripped:
                out_lines.append("")
                continue

            if (
                stripped.startswith(("#", "|", "![", ">", "```", "$$"))
                or re.match(r"^[-*+]\s+", stripped)
                or re.match(r"^\d+\.\s+", stripped)
                or re.match(r"^(?:---|\*\*\*|___)$", stripped)
            ):
                out_lines.append(line)
                continue

            out_lines.append(f"- {stripped}")

        text = "\n".join(out_lines)

        # Remove blank lines between adjacent list items and table rows.
        text = re.sub(r"(?m)(^[-*+] .*)\n\n(?=^[-*+] )", r"\1\n", text)
        text = re.sub(r"(?m)(^\d+\. .*)\n\n(?=^\d+\. )", r"\1\n", text)
        text = re.sub(r"(?m)(^\|.*\|)\n\n(?=^\|)", r"\1\n", text)
        # Remove blank line between a heading and a following bullet item.
        text = re.sub(r"(?m)(^#{1,6} .*)\n\n(?=^[-*+] )", r"\1\n", text)
        # Collapse excessive blank lines globally.
        text = re.sub(r"\n{3,}", "\n\n", text)
        return text.strip("\n")

    @staticmethod
    def is_plain_single_line(text: str) -> bool:
        t = text.strip()
        if "\n" in t:
            return False
        if t.startswith(("#", "-", "1.", "|", "!", ">")):
            return False
        return len(t) <= 40

    @staticmethod
    def is_structured_block(text: str) -> bool:
        t = text.lstrip()
        return t.startswith("|") or t.startswith("- ") or t.startswith("1.")

    def render_oe_children(self, oe_children: ET.Element, ctx: RenderContext) -> list[str]:
        blocks: list[str] = []
        for oe in oe_children.findall("one:OE", NS):
            blocks.extend(self.render_oe(oe, ctx))
        blocks = [b for b in self.compact_blocks(blocks) if b.strip()]
        blocks = self.postprocess_local_blocks(blocks, ctx)
        return blocks

    def postprocess_local_blocks(self, blocks: list[str], ctx: RenderContext) -> list[str]:
        # Promote standalone subsection titles inside a top-level section body:
        # e.g. "草酸盐抗凝剂" + following KV/list blocks -> "## 1.草酸盐抗凝剂".
        if ctx.section_depth != 1 or not blocks:
            return blocks
        out: list[str] = []
        sub_idx = 1
        i = 0
        while i < len(blocks):
            raw = blocks[i]
            b = raw.strip()
            if i + 1 < len(blocks):
                inline_sub = self.split_inline_subsection_lead(b, blocks[i + 1])
                if inline_sub is not None:
                    sub_title, first_body = inline_sub
                    out.append(f"## {sub_idx}.{sub_title}")
                    sub_idx += 1
                    if first_body:
                        out.append(first_body)
                    i += 1
                    continue
            if self.is_promotable_subsection_title(b) and i + 1 < len(blocks):
                nxt = blocks[i + 1].lstrip()
                if not nxt.startswith("#") and not nxt.startswith("!["):
                    out.append(f"## {sub_idx}.{b}")
                    sub_idx += 1
                    i += 1
                    continue
            out.append(raw)
            i += 1
        # Normalize ## numbering inside a top-level section body to keep sequence stable
        # after heuristic promotions (e.g. 第三节).
        renum: list[str] = []
        h2_idx = 1
        for block in out:
            lines_out: list[str] = []
            for line in block.split("\n"):
                m = re.match(r"^(##\s+)(\d+)\.(.*)$", line.strip())
                if m:
                    prefix, _, rest = m.groups()
                    lines_out.append(f"{prefix}{h2_idx}.{rest}".strip())
                    h2_idx += 1
                else:
                    lines_out.append(line)
            renum.append("\n".join(lines_out))
        return renum

    def split_inline_subsection_lead(self, block: str, next_block: str) -> tuple[str, str] | None:
        # Sample markdown often rewrites "标题：正文" + following structured block into
        # a subsection heading plus a bullet body item.
        t = block.strip()
        if not t or "\n" in t:
            return None
        if t.startswith(("#", "-", "1.", "|", "!", ">")):
            return None
        nxt = next_block.lstrip()
        if not (nxt.startswith("## ") or nxt.startswith("|") or nxt.startswith("![")):
            return None
        sep_idx = -1
        for sep in ("：", ":"):
            idx = t.find(sep)
            if idx != -1 and (sep_idx == -1 or idx < sep_idx):
                sep_idx = idx
        if sep_idx <= 0:
            return None
        raw_left = t[:sep_idx]
        raw_right = t[sep_idx + 1 :]
        title = strip_markup(raw_left).strip("：: ").strip()
        if not title or len(title) > 20:
            return None
        if re.search(r"!\[[^\]]*\]\(", title):
            return None
        if not looks_heading_label(title):
            return None
        body = raw_right.strip()
        if not body:
            return (title, "")
        if self.is_list_item_line(body) or body.startswith(("#", "|", "![", ">")):
            return (title, body)
        return (title, f"- {body}")

    @staticmethod
    def is_promotable_subsection_title(text: str) -> bool:
        if "<font" in text.lower():
            return False
        t = strip_markup(text).strip()
        if not t or "\n" in t:
            return False
        if len(t) > 24:
            return False
        if t.startswith(("#", "-", "1.", "|", "!", ">")):
            return False
        if re.search(r"[。；;，,:：]$", t):
            return False
        if re.search(r"[，。；;:：()（）↑↓]", t):
            return False
        # Titles are usually noun phrases, not long sentences.
        if any(ch in t for ch in [" ", "\t"]) and len(t) > 18:
            return False
        return True

    def compact_blocks(self, blocks: list[str]) -> list[str]:
        out: list[str] = []
        for b in blocks:
            b = b.replace("\r\n", "\n").replace("\r", "\n")
            cleaned_lines: list[str] = []
            for line in b.split("\n"):
                if line.strip() and not strip_markup(line).strip():
                    continue
                cleaned_lines.append(line)
            b = "\n".join(cleaned_lines)
            b = re.sub(r"\n{3,}", "\n\n", b)
            b = b.strip("\n")
            if not b.strip():
                continue
            if self.is_noise_block(b):
                continue
            out.append(b)
        return out

    @staticmethod
    def is_noise_block(text: str) -> bool:
        t = text.strip()
        plain = strip_markup(t).strip()
        if not plain:
            return True
        if plain in {"-", "•", "·"}:
            return True
        return False

    def render_oe(self, oe: ET.Element, ctx: RenderContext) -> list[str]:
        blocks: list[str] = []
        rendered_any = False
        for child in list(oe):
            name = tag_name(child)
            if name in {"Meta", "List", "OutlookTask", "Tag"}:
                continue
            if name == "T":
                text = html_fragment_to_markdown_text(child.text or "")
                if text:
                    text = self.apply_list_prefix_from_oe(text, oe, ctx)
                    blocks.extend(self.render_text_block(text, ctx))
                    rendered_any = True
            elif name == "Image":
                blocks.append(self.render_image(child))
                rendered_any = True
            elif name == "Table":
                blocks.extend(self.render_table(child, ctx))
                rendered_any = True
            elif name == "OEChildren":
                blocks.extend(self.render_oe_children(child, ctx))
                rendered_any = True
            elif name in {"Position", "Size", "InsertedFile"}:
                continue
        if not rendered_any:
            # Some OE nodes contain content only in nested OEChildren under unexpected wrappers.
            nested = oe.find("one:OEChildren", NS)
            if nested is not None:
                blocks.extend(self.render_oe_children(nested, ctx))
        return blocks

    def apply_list_prefix_from_oe(self, text: str, oe: ET.Element, ctx: RenderContext) -> str:
        list_el = oe.find("one:List", NS)
        if list_el is None:
            return text
        bullet = list_el.find("one:Bullet", NS)
        number = list_el.find("one:Number", NS)
        if bullet is not None:
            lines = [ln for ln in text.split("\n") if ln.strip()]
            if not lines:
                return text
            return "\n".join(f"- {ln.strip()}" for ln in lines)
        if number is not None:
            lines = [ln for ln in text.split("\n") if ln.strip()]
            return "\n".join(f"1. {ln.strip()}" for ln in lines)
        return text

    def render_text_block(self, text: str, ctx: RenderContext) -> list[str]:
        t = text.strip()
        t = strip_leading_colored_dash_markup(t).strip()
        if not t:
            return []
        circled = split_circled_number_lines(t)
        if circled:
            lines = []
            for idx, part in enumerate(circled, 1):
                cleaned = re.sub(r"^[①②③④⑤⑥⑦⑧⑨⑩]\s*", "", part).strip()
                lines.append(f"{idx}. {cleaned}")
            return ["\n".join(lines)]
        if re.match(r"^[①②③④⑤⑥⑦⑧⑨⑩]\s*", t):
            cleaned = re.sub(r"^[①②③④⑤⑥⑦⑧⑨⑩]\s*", "", t).strip()
            return [f"1. {cleaned}"]
        return [t]

    def render_image(self, image_el: ET.Element) -> str:
        ref = self.image_resolver.resolve(image_el)
        link = ref.link_path.replace("\\", "/")
        if self.image_syntax == "obsidian":
            return f"![[{link}]]"
        alt = ref.filename
        return f"![{alt}](<{link}>)"

    def render_table(self, table_el: ET.Element, ctx: RenderContext) -> list[str]:
        rows = table_el.findall("one:Row", NS)
        if not rows:
            return []
        has_header = (table_el.get("hasHeaderRow") or "").lower() == "true"
        row_cells: list[list[ET.Element]] = []
        max_cols = 0
        for row in rows:
            cells = row.findall("one:Cell", NS)
            row_cells.append(cells)
            max_cols = max(max_cols, len(cells))

        if max_cols == 1:
            return self.render_single_column_table(row_cells, ctx.child_for_table())

        if self.is_wrapper_section_table(row_cells, ctx):
            return self.render_wrapper_section_table(row_cells, ctx)

        if max_cols == 2 and self.is_hierarchical_table(row_cells, ctx):
            return self.render_hierarchical_table(row_cells, ctx)

        if max_cols == 2 and self.is_kv_table(row_cells):
            return [self.render_kv_table_as_list(row_cells, ctx.child_for_table())]

        return [self.render_markdown_table(row_cells, ctx.child_for_table(), has_header=has_header)]

    def render_single_column_table(self, row_cells: list[list[ET.Element]], ctx: RenderContext) -> list[str]:
        blocks: list[str] = []
        for cells in row_cells:
            if not cells:
                continue
            blocks.extend(self.render_cell_blocks(cells[0], ctx))
        return self.compact_blocks(blocks)

    def render_hierarchical_table(self, row_cells: list[list[ET.Element]], ctx: RenderContext) -> list[str]:
        blocks: list[str] = []
        level = ctx.section_depth + 1
        for idx, cells in enumerate(row_cells, 1):
            if len(cells) < 2:
                continue
            title = strip_markup(self.cell_to_inline_text(cells[0], ctx)).strip("：: ")
            if not title:
                continue
            heading = self.make_section_heading(title, idx, level)
            blocks.append(heading)
            body_blocks = self.render_cell_blocks(cells[1], ctx.with_section_depth(level))
            if body_blocks:
                blocks.extend(body_blocks)
        return self.compact_blocks(blocks)

    def is_hierarchical_table(self, row_cells: list[list[ET.Element]], ctx: RenderContext) -> bool:
        if ctx.section_depth >= 2:
            return False
        if len(row_cells) == 0:
            return False
        titles = []
        complex_count = 0
        property_like_count = 0
        for cells in row_cells:
            if len(cells) < 2:
                return False
            t = self.cell_to_inline_text(cells[0], ctx)
            if not looks_heading_label(t):
                return False
            titles.append(t)
            if is_property_like_label(t):
                property_like_count += 1
            if self.cell_has_nested_structure(cells[1]):
                complex_count += 1
        if titles and property_like_count >= max(1, len(titles) // 2):
            return False
        return complex_count >= 1 and len(titles) >= 1

    def analyze_cell(self, cell_el: ET.Element, ctx: RenderContext) -> dict[str, object]:
        plain = strip_markup(self.cell_to_inline_text(cell_el, ctx)).replace("\n", " ").strip()
        return {
            "cell": cell_el,
            "plain": plain,
            "nested": self.cell_has_nested_structure(cell_el),
            "len": len(plain),
        }

    def is_wrapper_section_table(self, row_cells: list[list[ET.Element]], ctx: RenderContext) -> bool:
        if ctx.section_depth > 1:
            return False
        if not row_cells:
            return False
        meaningful_rows = 0
        matched_rows = 0
        property_like_title_rows = 0
        all_rows_complex_2col = True
        for cells in row_cells:
            infos = [self.analyze_cell(c, ctx) for c in cells]
            nonempty = [info for info in infos if info["plain"]]
            if not nonempty:
                continue
            meaningful_rows += 1
            has_complex = any(bool(info["nested"]) for info in nonempty)
            if len(nonempty) != 2 or sum(bool(info["nested"]) for info in nonempty) != 1:
                all_rows_complex_2col = False
            # Wrapper rows usually have one short title cell and one complex body cell.
            has_title = any(looks_heading_label(str(info["plain"])) for info in nonempty)
            title_like_candidates = [
                info for info in nonempty if (not bool(info["nested"])) and int(info["len"]) <= 40
            ]
            if title_like_candidates and all(is_property_like_label(str(info["plain"])) for info in title_like_candidates):
                property_like_title_rows += 1
            if has_complex and has_title:
                matched_rows += 1
        if meaningful_rows == 0:
            return False
        if (
            property_like_title_rows >= max(1, meaningful_rows // 2)
            and not (all_rows_complex_2col and matched_rows == meaningful_rows and meaningful_rows >= 2)
        ):
            return False
        return matched_rows >= 1 and matched_rows >= meaningful_rows // 2

    def render_wrapper_section_table(self, row_cells: list[list[ET.Element]], ctx: RenderContext) -> list[str]:
        blocks: list[str] = []
        level = ctx.section_depth + 1
        section_idx = 1
        for cells in row_cells:
            infos = [self.analyze_cell(c, ctx) for c in cells]
            nonempty = [info for info in infos if info["plain"]]
            if not nonempty:
                continue

            complex_infos = [info for info in nonempty if bool(info["nested"])]
            if complex_infos:
                content_info = max(complex_infos, key=lambda x: int(x["len"]))
            else:
                content_info = max(nonempty, key=lambda x: int(x["len"]))

            title_candidates = [
                info for info in nonempty if info is not content_info and looks_heading_label(str(info["plain"]))
            ]
            title_info = min(title_candidates, key=lambda x: int(x["len"])) if title_candidates else None

            if title_info is not None:
                title = str(title_info["plain"]).strip("：: ")
                if title:
                    blocks.append(self.make_section_heading(title, section_idx, level))
                    section_idx += 1

            body_blocks = self.render_cell_blocks(content_info["cell"], ctx.with_section_depth(level))
            blocks.extend(body_blocks)

            for info in nonempty:
                if info is content_info or info is title_info:
                    continue
                extra_blocks = self.render_cell_blocks(info["cell"], ctx.with_section_depth(level))
                blocks.extend(extra_blocks)

        return self.compact_blocks(blocks)

    def make_section_heading(self, title: str, index: int, level: int) -> str:
        title = title.replace("\n", " ").strip()
        if level == 1:
            return f"# {chinese_numeral(index)}、{title}"
        if level == 2:
            return f"## {index}.{title}"
        return f"{'#' * min(level, 6)} {title}"

    def is_kv_table(self, row_cells: list[list[ET.Element]]) -> bool:
        if not row_cells:
            return False
        good = 0
        for cells in row_cells:
            if len(cells) != 2:
                return False
            left = strip_markup(self.cell_to_inline_text(cells[0], RenderContext()))
            if not looks_short_title(left):
                return False
            good += 1
        return good == len(row_cells)

    def render_kv_table_as_list(self, row_cells: list[list[ET.Element]], ctx: RenderContext) -> str:
        if self.should_render_kv_rows_as_sections(row_cells, ctx):
            return self.render_kv_rows_as_sections(row_cells, ctx)
        lines: list[str] = []
        for cells in row_cells:
            label = strip_markup(self.cell_to_inline_text(cells[0], ctx)).strip("：: ")
            body = self.render_cell_blocks(cells[1], ctx)
            body = [b for b in body if not self.is_noise_block(b)]
            body = self.normalize_kv_body_blocks(body)
            if not body:
                if label:
                    lines.append(f"- {label}")
                continue
            if len(body) == 1 and self.can_inline_kv_value(body[0]):
                value = body[0].replace("\n", "<br>")
                punct = "：" if self.needs_fullwidth_colon(label) else ": "
                lines.append(f"- {label}{punct}{value}".rstrip())
                continue

            lines.append(f"- {label}")
            nested_lines = self.indent_blocks(body, spaces=4)
            lines.extend(nested_lines)
        return "\n".join(lines)

    def should_render_kv_rows_as_sections(self, row_cells: list[list[ET.Element]], ctx: RenderContext) -> bool:
        if ctx.section_depth != 1:
            return False
        if not (2 <= len(row_cells) <= 4):
            return False
        labels: list[str] = []
        property_like = 0
        for cells in row_cells:
            if len(cells) != 2:
                return False
            label = strip_markup(self.cell_to_inline_text(cells[0], ctx)).strip("：: ")
            if not looks_heading_label(label):
                return False
            labels.append(label)
            if is_property_like_label(label):
                property_like += 1
        if property_like > 0:
            return False
        suffixes: dict[str, int] = {}
        for label in labels:
            plain = strip_markup(label).strip()
            if len(plain) >= 2:
                suf = plain[-2:]
                suffixes[suf] = suffixes.get(suf, 0) + 1
        if any(cnt >= 2 for cnt in suffixes.values()):
            return False
        # Avoid turning short “问题-原因” lookup tables into headings.
        if len(row_cells) >= 5:
            return False
        return True

    def render_kv_rows_as_sections(self, row_cells: list[list[ET.Element]], ctx: RenderContext) -> str:
        blocks: list[str] = []
        for idx, cells in enumerate(row_cells, 1):
            label = strip_markup(self.cell_to_inline_text(cells[0], ctx)).strip("：: ")
            if not label:
                continue
            blocks.append(f"## {idx}.{label}")
            body_blocks = self.render_cell_blocks(cells[1], ctx.with_section_depth(ctx.section_depth + 1))
            body_blocks = [b for b in body_blocks if not self.is_noise_block(b)]
            body_blocks = self.normalize_kv_body_blocks(body_blocks)
            blocks.extend(body_blocks)
        return "\n\n".join(blocks)

    def normalize_kv_body_blocks(self, blocks: list[str]) -> list[str]:
        if len(blocks) <= 1:
            return blocks
        has_media_or_table = any(
            (b.lstrip().startswith("![") or b.lstrip().startswith("|"))
            for b in blocks
            if b.strip()
        )
        if not has_media_or_table:
            return blocks

        out: list[str] = []
        for block in blocks:
            t = block.strip("\n")
            s = t.lstrip()
            if not s:
                continue
            if s.startswith(("#", "|", "![", "> ", "```", "$$")):
                out.append(t)
                continue

            lines = [ln.strip() for ln in t.split("\n") if ln.strip()]
            if not lines:
                continue
            if any(
                ln.startswith(("#", "|", "![", "> ", "```", "$$")) or self.is_list_item_line(ln)
                for ln in lines
            ):
                out.append(t)
                continue

            # Align with sample markdown style: prose notes in a KV value that also
            # contains media/table blocks are better represented as nested bullet items.
            out.append("\n".join(f"- {ln}" for ln in lines))
        return out

    @staticmethod
    def can_inline_kv_value(block: str) -> bool:
        b = block.strip()
        if any(b.startswith(prefix) for prefix in ("# ", "## ", "|", "> ")):
            return False
        if re.search(r"(^|\n)(- |\d+\. )", b):
            return False
        if "<br>- " in b or re.search(r"<br>\d+\. ", b):
            return False
        return True

    @staticmethod
    def needs_fullwidth_colon(label: str) -> bool:
        return bool(re.search(r"[\u4e00-\u9fff]", label))

    def render_markdown_table(self, row_cells: list[list[ET.Element]], ctx: RenderContext, has_header: bool = False) -> str:
        ncols = max(len(cells) for cells in row_cells) if row_cells else 0
        if ncols == 0:
            return ""
        data_rows = row_cells
        if has_header and row_cells:
            first_row = row_cells[0]
            header = []
            for ci in range(ncols):
                if ci < len(first_row):
                    header.append(self.cell_to_table_cell(first_row[ci], ctx))
                else:
                    header.append(" ")
            data_rows = row_cells[1:]
        else:
            header = [" " for _ in range(ncols)]
        sep = ["---" for _ in range(ncols)]
        lines = [
            "| " + " | ".join(escape_md_table_cell(h) for h in header) + " |",
            "| " + " | ".join(sep) + " |",
        ]
        for cells in data_rows:
            rendered_cells: list[str] = []
            for ci in range(ncols):
                if ci < len(cells):
                    inline = self.cell_to_table_cell(cells[ci], ctx)
                else:
                    inline = ""
                rendered_cells.append(escape_md_table_cell(inline))
            lines.append("| " + " | ".join(rendered_cells) + " |")
        return "\n".join(lines)

    def cell_to_table_cell(self, cell_el: ET.Element, ctx: RenderContext) -> str:
        blocks = self.render_cell_blocks(cell_el, ctx)
        if not blocks:
            return ""
        # If cell content contains a nested Markdown table, do not splice it into
        # the outer table cell (would produce garbled rendering). Return a placeholder
        # and let the nested table render as a separate block.
        if any(self._is_markdown_table_block(b) for b in blocks):
            return "[详见子表格]"
        if len(blocks) == 1:
            return blocks[0].replace("\n", "<br>")
        return "<br>".join(block.replace("\n", "<br>") for block in blocks)

    def cell_to_inline_text(self, cell_el: ET.Element, ctx: RenderContext) -> str:
        blocks = self.render_cell_blocks(cell_el, ctx)
        if not blocks:
            return ""
        # If cell contains a nested Markdown table, return a placeholder to avoid
        # splicing table rows into heading/label detection logic.
        if any(self._is_markdown_table_block(b) for b in blocks):
            return "[子表格]"
        return "<br>".join(block.replace("\n", "<br>") for block in blocks)

    def cell_has_nested_structure(self, cell_el: ET.Element) -> bool:
        return cell_el.find(".//one:Table", NS) is not None or cell_el.find(".//one:Image", NS) is not None

    def render_cell_blocks(self, cell_el: ET.Element, ctx: RenderContext) -> list[str]:
        oe_children = cell_el.find("one:OEChildren", NS)
        if oe_children is None:
            return []
        blocks = self.render_oe_children(oe_children, ctx)
        # Merge adjacent simple text lines into a list when the cell is a list-like cluster.
        return self.merge_listish_blocks(blocks)

    def merge_listish_blocks(self, blocks: list[str]) -> list[str]:
        if not blocks:
            return []
        out: list[str] = []
        i = 0
        while i < len(blocks):
            b = blocks[i].strip()
            if self.is_groupable_list_block(blocks[i]):
                group = [b]
                i += 1
                while i < len(blocks) and self.is_groupable_list_block(blocks[i]):
                    group.append(blocks[i].strip())
                    i += 1
                expanded: list[str] = []
                for item in group:
                    for line in item.split("\n"):
                        line = line.strip()
                        if line:
                            expanded.append(line)
                normalized = self.normalize_list_group(expanded)
                out.append("\n".join(normalized))
                continue
            out.append(b)
            i += 1
        return out

    def is_groupable_list_block(self, block: str) -> bool:
        lines = [ln for ln in block.split("\n") if ln.strip()]
        if not lines:
            return False
        # Nested lists (indented lines) are already structured; don't flatten them.
        if any(re.match(r"^\s+", ln) for ln in lines):
            return False
        return all(self.is_list_item_line(ln.strip()) for ln in lines)

    @staticmethod
    def is_list_item_line(line: str) -> bool:
        return bool(re.match(r"^(- |\d+\. |[①②③④⑤⑥⑦⑧⑨⑩]\s*)", line))

    def normalize_list_group(self, lines: list[str]) -> list[str]:
        if all(re.match(r"^\d+\. ", ln) for ln in lines):
            out = []
            for i, ln in enumerate(lines, 1):
                cleaned = re.sub(r"^\d+\.\s*", "", ln)
                cleaned = strip_leading_colored_dash_markup(cleaned).strip()
                if strip_markup(cleaned).strip():
                    out.append(f"{i}. {cleaned}")
            return out
        if all(re.match(r"^[①②③④⑤⑥⑦⑧⑨⑩]\s*", ln) for ln in lines):
            out = []
            for i, ln in enumerate(lines, 1):
                cleaned = re.sub(r"^[①②③④⑤⑥⑦⑧⑨⑩]\s*", "", ln)
                cleaned = strip_leading_colored_dash_markup(cleaned).strip()
                if strip_markup(cleaned).strip():
                    out.append(f"{i}. {cleaned}")
            return out
        out = []
        for ln in lines:
            cleaned = re.sub(r"^(?:-\s*|\d+\.\s*|[①②③④⑤⑥⑦⑧⑨⑩]\s*)", "", ln)
            cleaned = strip_leading_colored_dash_markup(cleaned).strip()
            if not strip_markup(cleaned).strip():
                continue
            out.append(f"- {cleaned}")
        return out

    @staticmethod
    def _is_markdown_table_block(text: str) -> bool:
        """Detect if a block is a Markdown table (all non-empty lines start with |)."""
        lines = [ln for ln in text.split("\n") if ln.strip()]
        return len(lines) >= 2 and all(ln.lstrip().startswith("|") for ln in lines)

    def indent_blocks(self, blocks: list[str], spaces: int = 4) -> list[str]:
        prefix = " " * spaces
        out: list[str] = []
        for block in blocks:
            if self._is_markdown_table_block(block):
                # Tables must not be indented: indented tables render as code blocks.
                # Break the parent list context with blank lines so the table renders as a standalone block.
                out.append("")
                for line in block.split("\n"):
                    out.append(line)
                out.append("")
            else:
                for line in block.split("\n"):
                    out.append(prefix + line if line else "")
        return out


def relativize_asset_links_for_output(
    md_text: str,
    out_file_parent: Path,
    output_root: Path,
    asset_dir_names: list[str],
    image_syntax: str = "markdown",
) -> str:
    # Image assets are stored under the output root (e.g. images/, _generated_images/).
    # When markdown files are written to nested subfolders, rewrite links to stay relative.
    rel_to_root = os.path.relpath(output_root, out_file_parent)
    if rel_to_root in {".", ""}:
        return md_text
    rel_prefix = rel_to_root.replace("\\", "/").rstrip("/")
    cleaned_roots = []
    for name in asset_dir_names:
        if not name:
            continue
        cleaned = name.strip("/\\")
        if cleaned:
            cleaned_roots.append(cleaned)
    valid_roots = tuple(f"{name}/" for name in cleaned_roots)
    if not valid_roots:
        return md_text

    def rewrite_path(link_path: str) -> str:
        if not link_path:
            return link_path
        if re.match(r"^(?:[a-zA-Z]+:|/)", link_path):
            return link_path
        if link_path.startswith(valid_roots):
            return f"{rel_prefix}/{link_path}"
        return link_path

    if image_syntax == "obsidian":
        def obs_repl(m: re.Match[str]) -> str:
            inner = m.group(1)
            return f"![[{rewrite_path(inner)}]]"

        return re.sub(r"!\[\[([^\]]+)\]\]", obs_repl, md_text)

    def md_repl(m: re.Match[str]) -> str:
        alt = m.group(1)
        path = m.group(2)
        return f"![{alt}](<{rewrite_path(path)}>)"

    return re.sub(r"!\[([^\]]*)\]\(<([^>]+)>\)", md_repl, md_text)


def convert_path(
    xml_input: Path,
    output_dir: Path,
    attachment_dir: Path | None,
    sort_key: str = "name",
    recursive: bool = False,
    image_syntax: str = "markdown",
    copy_attachments: bool = False,
    asset_dir_name: str = "images",
) -> tuple[list[Path], SequentialImageResolver]:
    output_dir.mkdir(parents=True, exist_ok=True)
    # By design, conversion is XML-driven: images come from XML-embedded binary data (if present),
    # otherwise OCR-based placeholders are generated so markdown keeps all image slots displayable.
    # attachment_dir is kept only for backward-compatible CLI calls and is intentionally ignored.
    attachment_dir = None
    attachment_link_prefix = ""
    copy_attachment_dir: Path | None = None
    placeholder_dir: Path
    placeholder_link_prefix: str

    if copy_attachments:
        asset_dir = output_dir / asset_dir_name
        copy_attachment_dir = asset_dir
        attachment_link_prefix = asset_dir_name.replace("\\", "/").rstrip("/")
        placeholder_dir = asset_dir
        placeholder_link_prefix = attachment_link_prefix
    else:
        placeholder_dir = output_dir / "_generated_images"
        placeholder_link_prefix = "_generated_images"

    resolver = SequentialImageResolver(
        attachment_dir,
        sort_key=sort_key,
        attachment_link_prefix=attachment_link_prefix,
        copy_attachment_dir=copy_attachment_dir,
        placeholder_dir=placeholder_dir,
        placeholder_link_prefix=placeholder_link_prefix,
    )
    converter = OneNoteXMLToMarkdownConverter(resolver, image_syntax=image_syntax)

    if xml_input.is_file():
        xml_files = [xml_input]
    else:
        pattern = "**/*.xml" if recursive else "*.xml"
        xml_files = sorted(xml_input.glob(pattern), key=section_sort_key)

    written: list[Path] = []
    for xml_file in xml_files:
        # Skip section hierarchy metadata exported by stage 1 (.one -> XML); only page XML should become markdown.
        if xml_file.name.lower() == "section-hierarchy.xml":
            continue
        md_text = converter.convert_file(xml_file)
        if xml_input.is_file():
            out_file = output_dir / (xml_file.stem + ".md")
        else:
            rel_xml = xml_file.relative_to(xml_input)
            out_file = output_dir / rel_xml.with_suffix(".md")
        out_file.parent.mkdir(parents=True, exist_ok=True)
        md_text = relativize_asset_links_for_output(
            md_text,
            out_file_parent=out_file.parent,
            output_root=output_dir,
            asset_dir_names=[asset_dir_name, "_generated_images"],
            image_syntax=image_syntax,
        )
        out_file.write_text(md_text, encoding="utf-8")
        written.append(out_file)
    return written, resolver


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="Convert OneNote XML (nested tables) to Markdown using heuristic rules learned from sample files."
    )
    p.add_argument("xml_input", help="XML file or directory (e.g. xml)")
    p.add_argument("-o", "--output-dir", default="converted_md", help="Output markdown directory")
    p.add_argument(
        "-a",
        "--attachment-dir",
        default=None,
        help="Reference sample resource directory (ignored by converter; kept only for backward compatibility)",
    )
    p.add_argument(
        "--image-sort",
        choices=["name", "mtime"],
        default="name",
        help="Deprecated compatibility option (ignored; images are not read from attachment resources)",
    )
    p.add_argument(
        "--image-syntax",
        choices=["markdown", "obsidian"],
        default="markdown",
        help="Image link format in output markdown",
    )
    p.add_argument(
        "--copy-attachments",
        action="store_true",
        help="Bundle generated/extracted image assets into the output directory for portable markdown output",
    )
    p.add_argument(
        "--asset-dir",
        default="images",
        help="Resource folder name under output-dir when --copy-attachments is enabled (default: images)",
    )
    p.add_argument("--recursive", action="store_true", help="Recursively scan xml_input when it is a directory")
    return p


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    xml_input = Path(args.xml_input)
    output_dir = Path(args.output_dir)
    attachment_dir = Path(args.attachment_dir) if args.attachment_dir else None

    if not xml_input.exists():
        parser.error(f"xml_input not found: {xml_input}")
    if attachment_dir and not attachment_dir.exists():
        print(
            f"Note: attachment-dir is ignored by the converter. Provided path does not exist and will be ignored: {attachment_dir}",
            file=sys.stderr,
        )
        attachment_dir = None

    written, resolver = convert_path(
        xml_input=xml_input,
        output_dir=output_dir,
        attachment_dir=attachment_dir,
        sort_key=args.image_sort,
        recursive=args.recursive,
        image_syntax=args.image_syntax,
        copy_attachments=args.copy_attachments,
        asset_dir_name=args.asset_dir,
    )
    print(f"Converted {len(written)} file(s) -> {output_dir}")
    for path in written:
        print(path)
    if args.attachment_dir:
        print(
            "Note: attachment-dir was ignored. Image rendering is produced from XML-embedded image data when available, otherwise OCR-based SVG placeholders are generated.",
            file=sys.stderr,
        )
    if resolver.copied_attachment_count:
        print(
            f"Copied {resolver.copied_attachment_count} attachment file(s) into output resources.",
            file=sys.stderr,
        )
    if resolver.extracted_from_xml_count:
        print(
            f"Extracted {resolver.extracted_from_xml_count} image file(s) directly from XML-embedded image data.",
            file=sys.stderr,
        )
    if resolver.placeholder_count:
        print(
            f"Info: {resolver.placeholder_count} image node(s) had no XML-embedded binary data and were written as SVG placeholders (theme-named when OCR is available).",
            file=sys.stderr,
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
