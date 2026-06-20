#!/usr/bin/env python
"""Generate British and American word pronunciations as PowerPoint-ready WAV."""

from __future__ import annotations

import argparse
import asyncio
import json
import subprocess
import sys
from pathlib import Path

import edge_tts
import imageio_ffmpeg


VOICES = {
    "uk": "en-GB-SoniaNeural",
    "us": "en-US-JennyNeural",
}


async def synthesize(
    word: str,
    variant: str,
    output_dir: Path,
    semaphore: asyncio.Semaphore,
) -> dict[str, object]:
    stem = word.lower()
    mp3_path = output_dir / f"{stem}_{variant}.mp3"
    wav_path = output_dir / f"{stem}_{variant}.wav"

    if wav_path.exists() and wav_path.stat().st_size > 1_000:
        return {
            "word": word,
            "variant": variant,
            "voice": VOICES[variant],
            "wav": str(wav_path.resolve()),
            "cached": True,
        }

    async with semaphore:
        last_error: Exception | None = None
        for attempt in range(5):
            try:
                communicate = edge_tts.Communicate(
                    text=word,
                    voice=VOICES[variant],
                    rate="-10%",
                )
                await communicate.save(str(mp3_path))
                break
            except Exception as error:  # Network failures are retried.
                last_error = error
                await asyncio.sleep(1.5 * (attempt + 1))
        else:
            raise RuntimeError(
                f"Could not synthesize {word}/{variant}: {last_error}"
            )

        ffmpeg = imageio_ffmpeg.get_ffmpeg_exe()
        command = [
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
        ]
        subprocess.run(command, check=True)
        mp3_path.unlink()

    if not wav_path.exists() or wav_path.stat().st_size <= 1_000:
        raise RuntimeError(f"Invalid WAV output: {wav_path}")

    return {
        "word": word,
        "variant": variant,
        "voice": VOICES[variant],
        "wav": str(wav_path.resolve()),
        "bytes": wav_path.stat().st_size,
        "cached": False,
    }


async def run(args: argparse.Namespace) -> None:
    model = json.loads(args.slides_json.read_text(encoding="utf-8"))
    words = [
        str(slide["word"]["word"]).strip()
        for slide in model["slides"]
        if slide.get("type") == "detail"
    ]
    if not words:
        raise ValueError("No detail-slide words found.")

    args.output_dir.mkdir(parents=True, exist_ok=True)
    semaphore = asyncio.Semaphore(args.concurrency)
    tasks = [
        synthesize(word, variant, args.output_dir, semaphore)
        for word in words
        for variant in ("uk", "us")
    ]
    results = await asyncio.gather(*tasks)

    manifest = {
        "source": "Microsoft Edge online neural text-to-speech",
        "voices": VOICES,
        "wordCount": len(words),
        "audioCount": len(results),
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
