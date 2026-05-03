#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
依赖安装（先在命令行执行）：

    pip install easyocr opencv-python

可选（若你已安装并配置好 tesseract，也可用 pytesseract 作为备选）：

    pip install pytesseract

说明：
- 按图片顺序 pic1.png -> pic2.png -> pic3.png 做 OCR
- 合并结果并做“初步结构识别”：标题 / Detail
- 最终写入同目录 content.txt
"""

from __future__ import annotations

import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional, Tuple

import numpy as np  # type: ignore


_EASYOCR_READER = None


def _safe_stdout_utf8() -> None:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass


def _imread_cv2_unicode_safe(image_path: Path):
    """
    在 Windows 中文路径下安全读取图片。
    cv2.imread 对部分非 ASCII 路径不稳定，改用 fromfile + imdecode。
    """
    import cv2  # type: ignore

    try:
        buf = np.fromfile(str(image_path), dtype=np.uint8)
        if buf.size == 0:
            return None
        img = cv2.imdecode(buf, cv2.IMREAD_COLOR)
        if img is not None:
            return img
    except Exception:
        pass

    # 兜底：常规读取
    return cv2.imread(str(image_path))


def _get_easyocr_reader():
    global _EASYOCR_READER
    if _EASYOCR_READER is None:
        # 兼容旧版 Pillow（easyocr 依赖 Image.Resampling）
        from PIL import Image as PILImage  # type: ignore

        if not hasattr(PILImage, "Resampling"):
            class _ResamplingCompat:
                NEAREST = PILImage.NEAREST
                BILINEAR = PILImage.BILINEAR
                BICUBIC = PILImage.BICUBIC
                BOX = getattr(PILImage, "BOX", PILImage.BILINEAR)
                HAMMING = getattr(PILImage, "HAMMING", PILImage.BILINEAR)
                LANCZOS = getattr(PILImage, "LANCZOS", PILImage.BICUBIC)

            PILImage.Resampling = _ResamplingCompat  # type: ignore[attr-defined]

        import easyocr  # type: ignore

        _EASYOCR_READER = easyocr.Reader(["ch_sim", "en"], gpu=False)
    return _EASYOCR_READER


def _read_with_easyocr(image_path: Path) -> List[str]:
    """
    使用 easyocr 读取图片，返回按从上到下的“行文本”列表。
    """
    reader = _get_easyocr_reader()
    img = _imread_cv2_unicode_safe(image_path)
    if img is None:
        raise FileNotFoundError(f"无法读取图片：{image_path}")

    # detail=1: (bbox, text, conf)
    results = reader.readtext(img, detail=1, paragraph=False)

    # 把 box 按 y 归并成行：用 bbox 的中心 y 做聚类
    items: List[Tuple[float, float, str]] = []
    for bbox, text, conf in results:
        if not text or not str(text).strip():
            continue
        # bbox = [[x1,y1],[x2,y2],[x3,y3],[x4,y4]]
        ys = [p[1] for p in bbox]
        xs = [p[0] for p in bbox]
        cy = float(sum(ys) / len(ys))
        cx = float(sum(xs) / len(xs))
        items.append((cy, cx, str(text).strip()))

    # 按 y 排序
    items.sort(key=lambda t: (t[0], t[1]))

    # 行聚类阈值：动态取图片高度的 1.2%，最小 10px
    h = img.shape[0]
    y_thresh = max(10.0, h * 0.012)

    lines: List[List[Tuple[float, str]]] = []
    for cy, cx, text in items:
        if not lines:
            lines.append([(cx, text)])
            last_y = cy
            continue
        # 用上一行的平均 y 来判断是否同一行
        last_line = lines[-1]
        # 估算上一行 y：用该行首个元素的 y（近似即可）
        # 为简单起见，比较与上一条 item 的 y 差
        if abs(cy - last_y) <= y_thresh:
            last_line.append((cx, text))
        else:
            lines.append([(cx, text)])
        last_y = cy

    # 行内按 x 排序并拼接
    out: List[str] = []
    for line in lines:
        line.sort(key=lambda t: t[0])
        merged = " ".join([t[1] for t in line]).strip()
        merged = re.sub(r"\s+", " ", merged)
        if merged:
            out.append(merged)
    return out


def _read_with_tesseract(image_path: Path) -> List[str]:
    """
    备选：使用 pytesseract（需要本机安装 tesseract 并配置 PATH）。
    """
    import pytesseract  # type: ignore

    img = _imread_cv2_unicode_safe(image_path)
    if img is None:
        raise FileNotFoundError(f"无法读取图片：{image_path}")

    # 中英混合
    txt = pytesseract.image_to_string(img, lang="chi_sim+eng")
    lines = [ln.strip() for ln in (txt or "").splitlines() if ln.strip()]
    return lines


_RE_DETAIL = re.compile(
    r"^\s*(?P<word>[A-Za-z][A-Za-z\-]*)\s*(?P<phonetic>/[^/]+/)?\s*[，,]?\s*(?P<rest>.*)$"
)


def _is_title(line: str) -> bool:
    s = (line or "").strip()
    if not s:
        return False
    return ("词根" in s) or ("释义" in s)


def _is_detail(line: str) -> bool:
    s = (line or "").strip()
    if not s:
        return False
    m = _RE_DETAIL.match(s)
    if not m:
        return False
    # detail 通常会包含中文释义或“即/adj./n./v.” 等
    rest = (m.group("rest") or "").strip()
    if any(k in rest for k in ("即", "n.", "v.", "adj.", "adv.", "释义")):
        return True
    # 或者含中文
    if re.search(r"[\u4e00-\u9fff]", rest):
        return True
    return False


def _normalize_lines(lines: List[str]) -> List[str]:
    out: List[str] = []
    for ln in lines:
        s = ln.strip()
        if not s:
            continue
        # 常见 OCR 错误修复：全角标点/空格
        s = s.replace("，", "，")
        s = re.sub(r"\s+", " ", s)
        out.append(s)
    return out


def structure_lines(lines: List[str]) -> List[str]:
    """
    初步结构化：标题行原样输出；Detail 行尽量规整为：
        word /phonetic/，<rest>
    其他行原样输出。
    """
    out: List[str] = []
    for ln in _normalize_lines(lines):
        if _is_title(ln):
            out.append(ln)
            continue
        if _is_detail(ln):
            m = _RE_DETAIL.match(ln)
            assert m is not None
            word = (m.group("word") or "").strip()
            phon = (m.group("phonetic") or "").strip()
            rest = (m.group("rest") or "").strip()
            # 统一标点
            if phon:
                out.append(f"{word} {phon}，{rest}".strip("，"))
            else:
                out.append(f"{word}，{rest}".strip("，"))
            continue
        out.append(ln)
    return out


def main() -> int:
    _safe_stdout_utf8()
    base = Path(__file__).resolve().parent
    images = [base / n for n in ("pic1.png", "pic2.png", "pic3.png")]
    content_path = base / "content.txt"

    missing = [p.name for p in images if not p.exists()]
    if missing:
        print(f"错误：找不到图片文件：{', '.join(missing)}（目录：{base}）")
        return 2

    all_lines: List[str] = []
    for p in images:
        print(f"[OCR] {p.name}")
        try:
            lines = _read_with_easyocr(p)
        except Exception as e_easy:
            print(f"easyocr 失败：{p.name} -> {e_easy}")
            try:
                lines = _read_with_tesseract(p)
                print(f"已回退到 tesseract：{p.name}")
            except Exception as e_tess:
                print(f"tesseract 也失败：{p.name} -> {e_tess}")
                return 3

        all_lines.extend(lines)
        all_lines.append("")  # 图片之间加空行分隔，便于人工检查

    structured = structure_lines(all_lines)
    merged_text = "\n".join(structured).rstrip() + "\n"

    content_path.write_text(merged_text, encoding="utf-8")
    print(f"已写入：{content_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

