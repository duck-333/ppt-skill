#!/usr/bin/env python
"""Download British/American word audio from Youdao and convert it to WAV."""

from __future__ import annotations

import argparse
import asyncio
import hashlib
import json
import re
import subprocess
import sys
from pathlib import Path
from urllib.parse import quote

import aiohttp
import imageio_ffmpeg


VARIANTS = {
    "uk": 1,
    "us": 2,
}


async def fetch_audio(
    session: aiohttp.ClientSession,
    semaphore: asyncio.Semaphore,
    word: str,
    variant: str,
) -> tuple[str, bytes]:
    audio_type = VARIANTS[variant]
    url = f"https://dict.youdao.com/dictvoice?audio={quote(word)}&type={audio_type}"
    last_error: Exception | None = None
    async with semaphore:
        for attempt in range(6):
            try:
                async with session.get(url) as response:
                    payload = await response.read()
                    content_type = response.headers.get("content-type", "")
                    if response.status != 200:
                        raise RuntimeError(f"HTTP {response.status}")
                    if "audio" not in content_type.lower() or len(payload) < 1_000:
                        raise RuntimeError(
                            f"Invalid audio response: {content_type}, {len(payload)} bytes"
                        )
                    return url, payload
            except Exception as error:
                last_error = error
                await asyncio.sleep(1.25 * (attempt + 1))
    raise RuntimeError(f"Could not download {word}/{variant}: {last_error}")


async def process_audio(
    session: aiohttp.ClientSession,
    semaphore: asyncio.Semaphore,
    word: str,
    variant: str,
    output_dir: Path,
) -> dict[str, object]:
    stem = word.lower()
    wav_path = output_dir / f"{stem}_{variant}.wav"
    if wav_path.exists() and wav_path.stat().st_size > 1_000:
        return {
            "word": word,
            "variant": variant,
            "wav": str(wav_path.resolve()),
            "bytes": wav_path.stat().st_size,
            "cached": True,
        }

    url, payload = await fetch_audio(session, semaphore, word, variant)
    mp3_path = output_dir / f"{stem}_{variant}.mp3"
    mp3_path.write_bytes(payload)

    ffmpeg = imageio_ffmpeg.get_ffmpeg_exe()
    subprocess.run(
        [
            ffmpeg,
            "-y",
            "-loglevel",
            "error",
            "-i",
            str(mp3_path),
            "-ac",
            "1",
            "-ar",
            "22050",
            "-sample_fmt",
            "s16",
            str(wav_path),
        ],
        check=True,
    )
    mp3_path.unlink()
    if not wav_path.exists() or wav_path.stat().st_size <= 1_000:
        raise RuntimeError(f"Invalid WAV output: {wav_path}")

    return {
        "word": word,
        "variant": variant,
        "sourceUrl": url,
        "sourceSha256": hashlib.sha256(payload).hexdigest(),
        "sourceBytes": len(payload),
        "wav": str(wav_path.resolve()),
        "bytes": wav_path.stat().st_size,
        "cached": False,
    }


async def run(args: argparse.Namespace) -> None:
    model = json.loads(args.slides_json.read_text(encoding="utf-8"))
    words: list[str] = []
    for slide in model["slides"]:
        if slide.get("type") not in {"detail", "concept"}:
            continue
        word = str(slide.get("word", {}).get("word", "")).strip()
        if not re.fullmatch(r"[A-Za-z][A-Za-z'\-]*", word):
            continue
        if word not in words:
            words.append(word)
    if not words:
        raise ValueError("No playable slide words found.")

    args.output_dir.mkdir(parents=True, exist_ok=True)
    timeout = aiohttp.ClientTimeout(total=45)
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
        "Referer": "https://dict.youdao.com/",
    }
    semaphore = asyncio.Semaphore(args.concurrency)
    async with aiohttp.ClientSession(timeout=timeout, headers=headers) as session:
        results = await asyncio.gather(
            *[
                process_audio(session, semaphore, word, variant, args.output_dir)
                for word in words
                for variant in ("uk", "us")
            ]
        )

    by_word: dict[str, dict[str, dict[str, object]]] = {}
    for item in results:
        by_word.setdefault(str(item["word"]), {})[str(item["variant"])] = item
    identical_variants = [
        word
        for word, variants in by_word.items()
        if variants.get("uk", {}).get("sourceSha256")
        and variants.get("uk", {}).get("sourceSha256")
        == variants.get("us", {}).get("sourceSha256")
    ]

    manifest = {
        "source": "Youdao online dictionary pronunciation",
        "variantMapping": {"uk": "type=1", "us": "type=2"},
        "wordCount": len(words),
        "audioCount": len(results),
        "identicalVariantWords": identical_variants,
        "items": results,
    }
    manifest_path = args.output_dir / "audio-manifest.json"
    manifest_path.write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(
        json.dumps(
            {
                "words": len(words),
                "wavFiles": len(results),
                "identicalVariantWords": identical_variants,
                "manifest": str(manifest_path.resolve()),
            },
            ensure_ascii=False,
        )
    )


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--slides-json", required=True, type=Path)
    parser.add_argument("--output-dir", required=True, type=Path)
    parser.add_argument("--concurrency", type=int, default=4)
    args = parser.parse_args()
    if args.concurrency < 1 or args.concurrency > 8:
        raise ValueError("--concurrency must be between 1 and 8.")
    asyncio.run(run(args))


if __name__ == "__main__":
    try:
        main()
    except Exception as error:
        print(str(error), file=sys.stderr)
        raise
