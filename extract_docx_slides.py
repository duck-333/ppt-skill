#!/usr/bin/env python
"""Parse a lesson DOCX into template-driven slide data.

The output is content-only JSON. Rendering is handled separately by the
artifact-tool presentation generator so the Word parser remains reusable.
"""

from __future__ import annotations

import argparse
import json
import re
import zipfile
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}

MAJOR_RE = re.compile(r"^[（(]\s*(\d+)\s*[）)]\s*(.+?)\s*[：:]?\s*$")
CIRCLED_RE = re.compile(r"^([①②③④⑤⑥⑦⑧⑨⑩])\s*(.+?)\s*[：:]?\s*$")
TOPIC_RE = re.compile(r"^\s*\d+\s*[.．、]\s*(.+?)\s*$")
WORD_RE = re.compile(
    r"^(?P<word>[A-Za-z][A-Za-z'\-]*)[\s\u00a0]*"
    r"(?P<phonetic>/[^/]+/)[\s\u00a0]*[，,]?\s*(?P<rest>.*)$"
)
POS_RE = re.compile(
    r"(?<![A-Za-z])(?:n|v|adj|adv|prep|conj|pron|num|art|interj)\."
)
SENTENCE_END_RE = re.compile(r"[。！？!?；;：:]$")


# Normalize the heading patterns explicitly. Earlier lessons contain a mix of
# parenthesized numbers, circled numbers, and decimal section headings.
MAJOR_RE = re.compile(r"^[（(]\s*(\d+)\s*[）)]\s*(.+?)\s*[：:]?\s*$")
CIRCLED_RE = re.compile(r"^([①②③④⑤⑥⑦⑧⑨⑩])\s*(.+?)\s*[：:]?\s*$")
TOPIC_RE = re.compile(r"^\s*\d+\s*[.．、]\s*(.+?)\s*$")
WORD_RE = re.compile(
    r"^(?P<word>[A-Za-z][A-Za-z'\-]*)[\s\u00a0]*"
    r"(?P<phonetic>/[^/]+/)[\s\u00a0]*[，,:：]?\s*(?P<rest>.*)$"
)
SENTENCE_END_RE = re.compile(r"[。！？；;]$")


@dataclass
class WordItem:
    word: str
    phonetic: str
    definition: str
    analysis: str
    source_paragraphs: list[int]


@dataclass
class Category:
    title: str
    words: list[WordItem]


def paragraph_text(paragraph: ET.Element) -> str:
    skipped = {
        f"{{{W_NS}}}pict",
        f"{{{W_NS}}}drawing",
        f"{{{W_NS}}}object",
        f"{{{W_NS}}}txbxContent",
        f"{{{W_NS}}}del",
    }
    text_tag = f"{{{W_NS}}}t"
    values: list[str] = []

    def visit(node: ET.Element) -> None:
        if node.tag in skipped:
            return
        if node.tag == text_tag:
            values.append(node.text or "")
            return
        for child in node:
            visit(child)

    visit(paragraph)
    return "".join(values).strip()


def extract_paragraphs(docx_path: Path) -> list[dict[str, object]]:
    with zipfile.ZipFile(docx_path) as archive:
        document_xml = archive.read("word/document.xml")

    root = ET.fromstring(document_xml)
    body = root.find("w:body", NS)
    if body is None:
        return []

    paragraphs: list[dict[str, object]] = []
    source_index = 0
    for child in body:
        if child.tag != f"{{{W_NS}}}p":
            continue
        source_index += 1
        text = paragraph_text(child).replace("\u00a0", " ").strip()
        if not text:
            continue
        style_node = child.find("./w:pPr/w:pStyle", NS)
        num_id = child.find("./w:pPr/w:numPr/w:numId", NS)
        ilvl = child.find("./w:pPr/w:numPr/w:ilvl", NS)
        paragraphs.append(
            {
                "index": source_index,
                "text": text,
                "style": style_node.get(f"{{{W_NS}}}val", "") if style_node is not None else "",
                "numId": num_id.get(f"{{{W_NS}}}val", "") if num_id is not None else "",
                "level": ilvl.get(f"{{{W_NS}}}val", "") if ilvl is not None else "",
            }
        )
    return paragraphs


def list_embedded_media(docx_path: Path) -> list[dict[str, object]]:
    with zipfile.ZipFile(docx_path) as archive:
        result = []
        for entry in archive.infolist():
            if not entry.filename.startswith("word/media/") or entry.is_dir():
                continue
            result.append(
                {
                    "name": Path(entry.filename).name,
                    "path": entry.filename,
                    "bytes": entry.file_size,
                }
            )
        return result


def extract_media(docx_path: Path, assets_dir: Path) -> list[str]:
    assets_dir.mkdir(parents=True, exist_ok=True)
    written: list[str] = []
    with zipfile.ZipFile(docx_path) as archive:
        for entry in archive.infolist():
            if not entry.filename.startswith("word/media/") or entry.is_dir():
                continue
            target = assets_dir / Path(entry.filename).name
            with archive.open(entry) as source:
                target.write_bytes(source.read())
            written.append(str(target.resolve()))
    return written


def clean_heading(text: str) -> str:
    text = text.strip().rstrip("：:").strip()
    match = MAJOR_RE.match(text)
    if match:
        return match.group(2).strip()
    match = CIRCLED_RE.match(text)
    if match:
        return match.group(2).strip()
    return text


def split_definition(rest: str) -> str:
    matches = list(POS_RE.finditer(rest))
    if not matches:
        return rest.strip(" ，,")
    return rest[matches[0].start() :].strip(" ，,")


def parse_word(text: str, paragraph_index: int) -> WordItem | None:
    match = WORD_RE.match(text.strip())
    if not match:
        return None
    word = match.group("word").strip()
    phonetic = match.group("phonetic").strip().strip("/")
    rest = match.group("rest").strip()
    if not rest:
        return None
    return WordItem(
        word=word,
        phonetic=phonetic,
        definition=split_definition(rest),
        analysis=rest,
        source_paragraphs=[paragraph_index],
    )


def append_note(word: WordItem, text: str, paragraph_index: int) -> None:
    note = text.strip()
    if not note:
        return
    word.analysis = f"{word.analysis}\n{note}".strip()
    word.source_paragraphs.append(paragraph_index)


def merge_intro_lines(lines: Iterable[str]) -> str:
    merged: list[str] = []
    for raw in lines:
        text = raw.strip()
        if not text:
            continue
        if not merged:
            merged.append(text)
            continue
        if SENTENCE_END_RE.search(merged[-1]):
            merged.append(text)
        else:
            merged[-1] = f"{merged[-1]}{text}"
    return "\n".join(merged)


def parse_categories(paragraphs: list[dict[str, object]], start_index: int) -> list[Category]:
    categories: list[Category] = []
    current_major = ""
    current_category: Category | None = None
    pending_notes: list[tuple[str, int]] = []
    last_word: WordItem | None = None

    def category_for(title: str) -> Category:
        nonlocal current_category
        if current_category is not None and not current_category.words:
            current_category.title = title
            return current_category
        current_category = Category(title=title, words=[])
        categories.append(current_category)
        return current_category

    for paragraph in paragraphs[start_index:]:
        text = str(paragraph["text"]).strip()
        paragraph_index = int(paragraph["index"])

        major_match = MAJOR_RE.match(text)
        topic_match = TOPIC_RE.match(text)
        if major_match or topic_match:
            current_major = clean_heading(text)
            current_category = None
            pending_notes = []
            last_word = None
            continue

        circled_match = CIRCLED_RE.match(text)
        if circled_match:
            category_for(clean_heading(text))
            last_word = None
            continue

        word = parse_word(text, paragraph_index)
        if word is not None:
            if current_category is None:
                category_for(current_major or "意项")
            assert current_category is not None
            if pending_notes:
                prefix = "\n".join(note for note, _ in pending_notes)
                word.analysis = f"{prefix}\n{word.analysis}".strip()
                word.source_paragraphs = [idx for _, idx in pending_notes] + word.source_paragraphs
                pending_notes = []
            current_category.words.append(word)
            last_word = word
            continue

        if text.startswith("【释义】"):
            if last_word is not None:
                append_note(last_word, text, paragraph_index)
            else:
                pending_notes.append((text, paragraph_index))
            continue

        if current_major:
            pending_notes.append((text, paragraph_index))

    return [category for category in categories if category.words]


def parse_simple_category(
    paragraphs: list[dict[str, object]],
    start_index: int,
    title: str,
) -> list[Category]:
    """Parse lessons whose word list is not wrapped in numbered headings."""
    words: list[WordItem] = []
    index = start_index
    while index < len(paragraphs):
        item = paragraphs[index]
        text = str(item["text"]).strip()
        paragraph_index = int(item["index"])
        if not WORD_RE.match(text):
            index += 1
            continue

        merged = text
        source_indexes = [paragraph_index]
        notes: list[tuple[str, int]] = []
        next_index = index + 1
        while next_index < len(paragraphs):
            next_item = paragraphs[next_index]
            next_text = str(next_item["text"]).strip()
            next_paragraph_index = int(next_item["index"])
            if WORD_RE.match(next_text) or MAJOR_RE.match(next_text) or CIRCLED_RE.match(next_text):
                break
            if next_text.startswith("【释义】"):
                notes.append((next_text, next_paragraph_index))
            else:
                merged = f"{merged}{next_text}"
                source_indexes.append(next_paragraph_index)
            next_index += 1

        word = parse_word(merged, paragraph_index)
        if word is not None:
            word.source_paragraphs = source_indexes
            for note, note_index in notes:
                append_note(word, note, note_index)
            words.append(word)
        index = next_index

    return [Category(title=title, words=words)] if words else []


def chunked(items: list[object], size: int) -> list[list[object]]:
    return [items[index : index + size] for index in range(0, len(items), size)]


def build_slides_model(docx_path: Path) -> dict[str, object]:
    paragraphs = extract_paragraphs(docx_path)
    if not paragraphs:
        raise ValueError(f"No readable paragraphs found in {docx_path}")

    first_major_index = next(
        (index for index, item in enumerate(paragraphs) if MAJOR_RE.match(str(item["text"]))),
        len(paragraphs),
    )

    topic_index = next(
        (
            index
            for index, item in enumerate(paragraphs[:first_major_index])
            if TOPIC_RE.match(str(item["text"])) and "字母" in str(item["text"])
        ),
        1 if len(paragraphs) > 1 else 0,
    )
    topic_match = TOPIC_RE.match(str(paragraphs[topic_index]["text"]))
    root_name = topic_match.group(1).strip() if topic_match else str(paragraphs[topic_index]["text"]).strip()

    first_word_index = next(
        (
            index
            for index, item in enumerate(paragraphs[topic_index + 1 :], start=topic_index + 1)
            if WORD_RE.match(str(item["text"]).strip())
        ),
        len(paragraphs),
    )
    intro_end_index = min(first_major_index, first_word_index)
    intro_lines = [
        str(item["text"])
        for item in paragraphs[topic_index + 1 : intro_end_index]
        if not TOPIC_RE.match(str(item["text"]))
    ]
    root_logic = merge_intro_lines(intro_lines)
    categories = parse_categories(paragraphs, first_major_index)
    if not categories and first_word_index < len(paragraphs):
        categories = parse_simple_category(paragraphs, first_word_index, root_name)

    letter_match = re.search(r"字母\s*([A-Za-z])", root_name)
    top_category = letter_match.group(1).upper() if letter_match else ""
    lesson_title = docx_path.stem

    concept_letters: list[str] = []
    for match in re.finditer(
        r"(?:字母\s*|通\s*)([A-Za-z])",
        f"{root_name}\n{root_logic}",
        flags=re.IGNORECASE,
    ):
        value = match.group(1).upper()
        if value not in concept_letters:
            concept_letters.append(value)
    if top_category and top_category not in concept_letters:
        concept_letters.insert(0, top_category)
    concept_word = " / ".join(concept_letters) if concept_letters else top_category
    phonetic_match = re.search(r"/([^/]+?)/", root_logic)
    concept_phonetic = phonetic_match.group(1).strip() if phonetic_match else ""

    slides: list[dict[str, object]] = [
        {
            "type": "cover",
            "sourceSlide": 1,
            "topCategory": top_category,
            "title": lesson_title,
        }
    ]

    intro_parts = [part for part in root_logic.splitlines() if part.strip()]
    intro_definition_count = 1 if len(intro_parts) <= 2 else 2
    slides.append(
        {
            "type": "concept",
            "sourceSlide": 2,
            "topCategory": top_category,
            "category": root_name,
            "word": {
                "word": concept_word,
                "phonetic": concept_phonetic,
                "definition": "\n".join(intro_parts[:intro_definition_count]),
                "analysis": "\n".join(intro_parts[intro_definition_count:]),
                "source_paragraphs": [
                    int(item["index"])
                    for item in paragraphs[topic_index + 1 : intro_end_index]
                    if not TOPIC_RE.match(str(item["text"]))
                ],
            },
        }
    )

    for page_index, group in enumerate(chunked(categories, 3), start=1):
        typed_group = [item for item in group if isinstance(item, Category)]
        slides.append(
            {
                "type": "overview",
                "sourceSlide": 1,
                "topCategory": top_category,
                "rootName": root_name,
                "rootLogic": "",
                "pageIndex": page_index,
                "pageCount": (len(categories) + 2) // 3,
                "items": [category.title for category in typed_group],
            }
        )

    for category in categories:
        for word in category.words:
            slides.append(
                {
                    "type": "detail",
                    "sourceSlide": 2,
                    "topCategory": top_category,
                    "category": category.title,
                    "word": asdict(word),
                }
            )

    summary_page_size = 6
    for page_index, group in enumerate(chunked(categories, summary_page_size), start=1):
        typed_group = [item for item in group if isinstance(item, Category)]
        slides.append(
            {
                "type": "summary",
                "sourceSlide": 3,
                "pageIndex": page_index,
                "pageCount": (len(categories) + summary_page_size - 1) // summary_page_size,
                "items": [
                    {
                        "title": category.title,
                        "words": [word.word for word in category.words],
                    }
                    for category in typed_group
                ],
            }
        )

    return {
        "schemaVersion": 1,
        "source": {
            "docx": str(docx_path.resolve()),
            "paragraphCount": len(paragraphs),
            "embeddedMedia": list_embedded_media(docx_path),
        },
        "lesson": {
            "title": lesson_title,
            "topCategory": top_category,
            "rootName": root_name,
            "rootLogic": root_logic,
        },
        "categories": [
            {
                "title": category.title,
                "words": [asdict(word) for word in category.words],
            }
            for category in categories
        ],
        "slides": slides,
        "counts": {
            "categories": len(categories),
            "words": sum(len(category.words) for category in categories),
            "slides": len(slides),
        },
        "paragraphs": paragraphs,
    }


MAJOR_RE_V2 = re.compile(
    r"^[\(\uff08]\s*([\d\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341]+)"
    r"\s*[\)\uff09]\s*(.+?)\s*[:\uff1a]?\s*$"
)
CIRCLED_RE_V2 = re.compile(r"^([\u2460-\u2469])\s*(.+?)\s*[:\uff1a]?\s*$")
TOPIC_RE_V2 = re.compile(r"^\s*\d+\s*[.\uff0e\u3001]\s*(.+?)\s*$")
WORD_RE_V2 = re.compile(
    r"^(?P<word>[A-Za-z][A-Za-z'\-]*)[\s\u00a0]*"
    r"(?P<phonetic>/[^/]+/)[\s\u00a0]*[,\uff0c:\uff1a]?\s*(?P<rest>.*)$"
)
SENTENCE_END_RE_V2 = re.compile(r"[\u3002\uff01\uff1f\uff1b;]$")


def clean_heading_v2(text: str) -> str:
    value = text.strip().rstrip(":\uff1a").strip()
    for pattern, group_index in (
        (MAJOR_RE_V2, 2),
        (CIRCLED_RE_V2, 2),
        (TOPIC_RE_V2, 1),
    ):
        match = pattern.match(value)
        if match:
            return match.group(group_index).strip()
    return value


def split_definition_v2(rest: str) -> str:
    matches = list(POS_RE.finditer(rest))
    if not matches:
        return rest.strip(" ,\uff0c")
    return rest[matches[0].start() :].strip(" ,\uff0c")


def parse_word_v2(text: str, paragraph_index: int) -> WordItem | None:
    match = WORD_RE_V2.match(text.strip())
    if not match:
        return None
    rest = match.group("rest").strip()
    if not rest:
        return None
    return WordItem(
        word=match.group("word").strip(),
        phonetic=match.group("phonetic").strip().strip("/"),
        definition=split_definition_v2(rest),
        analysis=rest,
        source_paragraphs=[paragraph_index],
    )


def merge_intro_lines_v2(lines: Iterable[str]) -> str:
    merged: list[str] = []
    for raw in lines:
        text = raw.strip()
        if not text:
            continue
        if not merged or SENTENCE_END_RE_V2.search(merged[-1]):
            merged.append(text)
        else:
            merged[-1] = f"{merged[-1]}{text}"
    return "\n".join(merged)


def parse_categories_v2(
    paragraphs: list[dict[str, object]], start_index: int
) -> list[Category]:
    categories: list[Category] = []
    current_major = ""
    current_category: Category | None = None
    pending_notes: list[tuple[str, int]] = []
    last_word: WordItem | None = None

    def category_for(title: str) -> Category:
        nonlocal current_category
        if current_category is not None and not current_category.words:
            current_category.title = title
            return current_category
        current_category = Category(title=title, words=[])
        categories.append(current_category)
        return current_category

    for paragraph in paragraphs[start_index:]:
        text = str(paragraph["text"]).strip()
        paragraph_index = int(paragraph["index"])

        if MAJOR_RE_V2.match(text) or TOPIC_RE_V2.match(text):
            current_major = clean_heading_v2(text)
            current_category = None
            pending_notes = []
            last_word = None
            continue

        if CIRCLED_RE_V2.match(text):
            category_for(clean_heading_v2(text))
            pending_notes = []
            last_word = None
            continue

        word = parse_word_v2(text, paragraph_index)
        if word is not None:
            if current_category is None:
                category_for(current_major or "\u8bcd\u7ec4")
            assert current_category is not None
            if pending_notes:
                prefix = "\n".join(note for note, _ in pending_notes)
                word.analysis = f"{prefix}\n{word.analysis}".strip()
                word.source_paragraphs = [
                    idx for _, idx in pending_notes
                ] + word.source_paragraphs
                pending_notes = []
            current_category.words.append(word)
            last_word = word
            continue

        if text.startswith("\u3010\u91ca\u4e49\u3011"):
            if last_word is not None:
                append_note(last_word, text, paragraph_index)
                last_word.definition = split_definition_v2(
                    last_word.analysis.replace("\n", "")
                )
            else:
                pending_notes.append((text, paragraph_index))
            continue

        if current_major or current_category is not None:
            if last_word is not None:
                append_note(last_word, text, paragraph_index)
                last_word.definition = split_definition_v2(
                    last_word.analysis.replace("\n", "")
                )
            else:
                pending_notes.append((text, paragraph_index))

    return [category for category in categories if category.words]


def parse_simple_category_v2(
    paragraphs: list[dict[str, object]], start_index: int, title: str
) -> list[Category]:
    words: list[WordItem] = []
    index = start_index
    while index < len(paragraphs):
        paragraph = paragraphs[index]
        text = str(paragraph["text"]).strip()
        paragraph_index = int(paragraph["index"])
        if not WORD_RE_V2.match(text):
            index += 1
            continue

        merged = text
        source_indexes = [paragraph_index]
        notes: list[tuple[str, int]] = []
        next_index = index + 1
        while next_index < len(paragraphs):
            next_item = paragraphs[next_index]
            next_text = str(next_item["text"]).strip()
            next_paragraph_index = int(next_item["index"])
            if (
                WORD_RE_V2.match(next_text)
                or MAJOR_RE_V2.match(next_text)
                or CIRCLED_RE_V2.match(next_text)
                or TOPIC_RE_V2.match(next_text)
            ):
                break
            if next_text.startswith("\u3010\u91ca\u4e49\u3011"):
                notes.append((next_text, next_paragraph_index))
            else:
                merged = f"{merged}{next_text}"
                source_indexes.append(next_paragraph_index)
            next_index += 1

        word = parse_word_v2(merged, paragraph_index)
        if word is not None:
            word.source_paragraphs = source_indexes
            for note, note_index in notes:
                append_note(word, note, note_index)
            words.append(word)
        index = next_index

    return [Category(title=title, words=words)] if words else []


def build_slides_model_v2(docx_path: Path) -> dict[str, object]:
    paragraphs = extract_paragraphs(docx_path)
    if not paragraphs:
        raise ValueError(f"No readable paragraphs found in {docx_path}")

    structural_indexes = [
        index
        for index, item in enumerate(paragraphs)
        if (
            MAJOR_RE_V2.match(str(item["text"]).strip())
            or TOPIC_RE_V2.match(str(item["text"]).strip())
        )
        and "\u5b57\u6bcd" in str(item["text"])
    ]
    topic_index = structural_indexes[0] if structural_indexes else (1 if len(paragraphs) > 1 else 0)
    root_name = clean_heading_v2(str(paragraphs[topic_index]["text"]))

    first_word_index = next(
        (
            index
            for index, item in enumerate(paragraphs[topic_index + 1 :], start=topic_index + 1)
            if WORD_RE_V2.match(str(item["text"]).strip())
        ),
        len(paragraphs),
    )
    first_section_index = next(
        (
            index
            for index, item in enumerate(paragraphs[topic_index + 1 :], start=topic_index + 1)
            if (
                MAJOR_RE_V2.match(str(item["text"]).strip())
                or TOPIC_RE_V2.match(str(item["text"]).strip())
                or CIRCLED_RE_V2.match(str(item["text"]).strip())
            )
        ),
        len(paragraphs),
    )
    intro_end_index = min(first_word_index, first_section_index)
    intro_lines = [
        str(item["text"])
        for item in paragraphs[topic_index + 1 : intro_end_index]
        if not (
            MAJOR_RE_V2.match(str(item["text"]).strip())
            or TOPIC_RE_V2.match(str(item["text"]).strip())
            or CIRCLED_RE_V2.match(str(item["text"]).strip())
        )
    ]
    root_logic = merge_intro_lines_v2(intro_lines)

    if first_section_index < first_word_index:
        categories = parse_categories_v2(paragraphs, first_section_index)
    elif first_word_index < len(paragraphs):
        categories = parse_simple_category_v2(paragraphs, first_word_index, root_name)
    else:
        categories = []

    letter_match = re.search(r"\u5b57\u6bcd\s*([A-Za-z])", root_name)
    top_category = letter_match.group(1).upper() if letter_match else ""
    lesson_title = docx_path.stem
    if not top_category:
        filename_letter_match = re.search(r"\u5b57\u6bcd\s*([A-Za-z])", lesson_title)
        if filename_letter_match:
            top_category = filename_letter_match.group(1).upper()

    concept_letters: list[str] = []
    for match in re.finditer(
        r"(?:\u5b57\u6bcd\s*|\u8bcd\u6839\s*)([A-Za-z])",
        f"{root_name}\n{root_logic}",
        flags=re.IGNORECASE,
    ):
        value = match.group(1).upper()
        if value not in concept_letters:
            concept_letters.append(value)
    if top_category and top_category not in concept_letters:
        concept_letters.insert(0, top_category)
    concept_word = " / ".join(concept_letters) if concept_letters else top_category
    phonetic_match = re.search(r"/([^/]+?)/", root_logic)
    concept_phonetic = phonetic_match.group(1).strip() if phonetic_match else ""

    slides: list[dict[str, object]] = [
        {
            "type": "cover",
            "sourceSlide": 1,
            "topCategory": top_category,
            "title": lesson_title,
        }
    ]

    intro_parts = [part for part in root_logic.splitlines() if part.strip()]
    intro_definition_count = 1 if len(intro_parts) <= 2 else 2
    slides.append(
        {
            "type": "concept",
            "sourceSlide": 2,
            "topCategory": top_category,
            "category": root_name,
            "word": {
                "word": concept_word,
                "phonetic": concept_phonetic,
                "definition": "\n".join(intro_parts[:intro_definition_count]),
                "analysis": "\n".join(intro_parts[intro_definition_count:]),
                "source_paragraphs": [
                    int(item["index"])
                    for item in paragraphs[topic_index + 1 : intro_end_index]
                ],
            },
        }
    )

    for page_index, group in enumerate(chunked(categories, 3), start=1):
        typed_group = [item for item in group if isinstance(item, Category)]
        slides.append(
            {
                "type": "overview",
                "sourceSlide": 1,
                "topCategory": top_category,
                "rootName": root_name,
                "rootLogic": "",
                "pageIndex": page_index,
                "pageCount": (len(categories) + 2) // 3,
                "items": [category.title for category in typed_group],
            }
        )

    for category in categories:
        for word in category.words:
            slides.append(
                {
                    "type": "detail",
                    "sourceSlide": 2,
                    "topCategory": top_category,
                    "category": category.title,
                    "word": asdict(word),
                }
            )

    summary_page_size = 6
    for page_index, group in enumerate(chunked(categories, summary_page_size), start=1):
        typed_group = [item for item in group if isinstance(item, Category)]
        slides.append(
            {
                "type": "summary",
                "sourceSlide": 3,
                "pageIndex": page_index,
                "pageCount": (len(categories) + summary_page_size - 1) // summary_page_size,
                "items": [
                    {
                        "title": category.title,
                        "words": [word.word for word in category.words],
                    }
                    for category in typed_group
                ],
            }
        )

    return {
        "schemaVersion": 1,
        "source": {
            "docx": str(docx_path.resolve()),
            "paragraphCount": len(paragraphs),
            "embeddedMedia": list_embedded_media(docx_path),
        },
        "lesson": {
            "title": lesson_title,
            "topCategory": top_category,
            "rootName": root_name,
            "rootLogic": root_logic,
        },
        "categories": [
            {
                "title": category.title,
                "words": [asdict(word) for word in category.words],
            }
            for category in categories
        ],
        "slides": slides,
        "counts": {
            "categories": len(categories),
            "words": sum(len(category.words) for category in categories),
            "slides": len(slides),
        },
        "paragraphs": paragraphs,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Parse a lesson DOCX into slides JSON.")
    parser.add_argument("--docx", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--assets-dir", type=Path)
    args = parser.parse_args()

    model = build_slides_model_v2(args.docx)
    if args.assets_dir:
        model["source"]["extractedMedia"] = extract_media(args.docx, args.assets_dir)

    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(model, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(model["counts"], ensure_ascii=False))


if __name__ == "__main__":
    main()
