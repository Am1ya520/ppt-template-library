"""Import .pptx files into a local deck library and refresh its catalog.

    python tools/import_decks.py "<file-or-folder>" [more ...] [--library DIR]

Each imported deck becomes `deckNN/{source.pptx,index.json,meta.json}` inside the
library root. The root is `--library`, else `$PPTL_LIB`, else `./library` beside
the repository.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import subprocess
import sys
from collections import Counter
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
PPTL = REPO / "pptl.py"


def library_root(explicit: str | None = None) -> Path:
    if explicit:
        return Path(explicit).expanduser().resolve()
    env = os.environ.get("PPTL_LIB")
    if env:
        return Path(env).expanduser().resolve()
    return REPO / "library"


def fillability(page: dict) -> float:
    """How suitable a page is for being re-filled (0..1)."""
    slots = [s for s in page["slots"] if s["paragraphs"] and any(p.strip() for p in s["paragraphs"])]
    if not slots:
        return 0.0
    texts = [" ".join(p for p in s["paragraphs"] if p.strip()) for s in slots]
    avg_len = sum(len(t) for t in texts) / len(texts)
    long_frac = sum(1 for t in texts if len(t) > 40) / len(texts)
    score = 1.0
    score -= min(avg_len / 120.0, 0.45)
    score -= long_frac * 0.35
    score -= max(0, len(slots) - 9) * 0.02
    return round(max(0.0, min(1.0, score)), 3)


def readable_pages(data: dict) -> list[int]:
    """Pages that can host real paragraphs: several text slots, one of them big."""
    out: list[int] = []
    for page in data["pages"]:
        texts = [s for s in page["slots"] if s["paragraphs"] and any(x.strip() for x in s["paragraphs"])]
        if not texts:
            continue
        big = [s for s in texts if (s.get("max_pt") or 0) >= 18]
        if len(texts) >= 3 and big and page.get("fillability", 0) >= 0.5:
            out.append(page["index"])
    return out


def next_deck_id(lib: Path) -> str:
    used = {p.name for p in lib.glob("deck*") if p.is_dir()}
    i = 1
    while f"deck{i:02d}" in used:
        i += 1
    return f"deck{i:02d}"


def collect(paths: list[str]) -> list[Path]:
    out: list[Path] = []
    for raw in paths:
        p = Path(raw)
        if p.is_dir():
            out.extend(sorted(p.rglob("*.pptx")))
        elif p.suffix.lower() == ".pptx":
            out.append(p)
    return [p for p in out if not p.name.startswith("~$")]


def file_sig(path: Path) -> tuple[int, str] | None:
    """Cheap fingerprint: size plus first/last 1 MB. Never reads whole files."""
    try:
        size = path.stat().st_size
        h = hashlib.sha1()
        with open(path, "rb") as f:
            h.update(f.read(1 << 20))
            if size > (2 << 20):
                f.seek(size - (1 << 20))
                h.update(f.read(1 << 20))
        return size, h.hexdigest()
    except OSError:
        return None


def import_one(src: Path, lib: Path, scenario: str = "未分类") -> dict | None:
    deck_id = next_deck_id(lib)
    dst = lib / deck_id
    dst.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, dst / "source.pptx")
    rc = subprocess.run(
        [sys.executable, str(PPTL), "index", str(src), "-o", str(dst)],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    produced = dst / f"{src.stem}.json"
    idx = dst / "index.json"
    if produced.exists():
        produced.replace(idx)
    if rc.returncode != 0 or not idx.exists():
        print(f"!! index failed for {src.name}: {rc.stderr[-300:]}")
        return None
    data = json.loads(idx.read_text(encoding="utf-8"))
    for page in data["pages"]:
        page["fillability"] = fillability(page)
    idx.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    mix = Counter(p["page_type"] for p in data["pages"])
    best = sorted(data["pages"], key=lambda p: -p["fillability"])[:6]
    meta = {
        "deck_id": deck_id,
        "source_name": src.stem,
        "scenario": scenario,
        "slides": data["slides"],
        "ratio": data["ratio"],
        "theme": data.get("theme", {}),
        "page_mix": dict(mix),
        "readable_pages": readable_pages(data),
        "fillability_overall": round(
            sum(p["fillability"] for p in data["pages"]) / max(len(data["pages"]), 1), 3
        ),
        "best_pages": [
            {"index": p["index"], "type": p["page_type"], "fillability": p["fillability"]} for p in best
        ],
    }
    (dst / "meta.json").write_text(json.dumps(meta, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"{deck_id}  {src.stem:<40} {data['slides']:>3}p  {dict(mix)}")
    return meta


def rewrite_catalog(lib: Path) -> int:
    metas = []
    for meta_file in sorted(lib.glob("deck*/meta.json")):
        metas.append(json.loads(meta_file.read_text(encoding="utf-8")))
    (lib / "catalog.json").write_text(
        json.dumps({"decks": metas}, ensure_ascii=False, indent=1), encoding="utf-8"
    )
    lines = [
        "# 模板库总表 / Template library catalog",
        "",
        "`可改页` = 能直接放正文的页数（正文槽 ≥18pt、≥3 个文本框、可填充度 ≥0.5）。",
        "",
        "| id | 模板 | 页数 | 可改页 | 封面 | 目录 | 章节 | 内容 | 密排 | 结尾 | 可填充度 |",
        "| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |",
    ]
    for m in metas:
        mix = m["page_mix"]
        lines.append(
            f"| {m['deck_id']} | {m['source_name']} | {m['slides']} | "
            f"{len(m.get('readable_pages') or [])} | "
            f"{mix.get('cover', 0)} | {mix.get('toc', 0)} | {mix.get('chapter', 0)} | "
            f"{mix.get('content', 0)} | {mix.get('dense', 0)} | {mix.get('ending', 0)} | "
            f"{m['fillability_overall']:.2f} |"
        )
    (lib / "catalog.md").write_text("\n".join(lines) + "\n", encoding="utf-8")
    return len(metas)


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="import .pptx templates into the deck library")
    ap.add_argument("paths", nargs="+", help="files or folders")
    ap.add_argument("--library", help="library root (default: $PPTL_LIB or ./library)")
    ap.add_argument("--scenario", default="未分类", help="label stored in meta.json")
    args = ap.parse_args(argv)

    lib = library_root(args.library)
    lib.mkdir(parents=True, exist_ok=True)
    files = collect(args.paths)
    if not files:
        print("no .pptx found")
        return 1

    existing = set()
    for deck in lib.glob("deck*"):
        sig = file_sig(deck / "source.pptx")
        if sig:
            existing.add(sig)

    imported = skipped = 0
    for src in files:
        sig = file_sig(src)
        if sig is None:
            continue
        if sig in existing:
            skipped += 1
            continue
        if import_one(src, lib, args.scenario):
            existing.add(sig)
            imported += 1

    total = rewrite_catalog(lib)
    print(f"\nimported {imported}, skipped {skipped} duplicates; library now {total} decks")
    print(f"catalog -> {lib/'catalog.md'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
