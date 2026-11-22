"""Export a small, git-friendly manifest of a deck library.

    python tools/export_manifest.py [--library DIR] [-o manifest.json]

The manifest holds metadata only (no .pptx bytes): deck id, template name,
page count, which pages can host real text, page-type mix, theme colours and
fonts. It is what a repository publishes so other people can see what a library
contains before downloading any template.
"""
from __future__ import annotations

import argparse
import json
import os
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent


def library_root(explicit: str | None = None) -> Path:
    if explicit:
        return Path(explicit).expanduser().resolve()
    env = os.environ.get("PPTL_LIB")
    if env:
        return Path(env).expanduser().resolve()
    return REPO / "library"


def main(argv=None) -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--library")
    ap.add_argument("-o", "--out", default=str(REPO / "manifest.json"))
    args = ap.parse_args(argv)

    lib = library_root(args.library)
    decks = []
    for meta_file in sorted(lib.glob("deck*/meta.json")):
        m = json.loads(meta_file.read_text(encoding="utf-8"))
        theme = m.get("theme") or {}
        decks.append(
            {
                "id": m["deck_id"],
                "name": m["source_name"],
                "slides": m["slides"],
                "ratio": round(m.get("ratio") or 0, 3),
                "scenario": m.get("scenario", ""),
                "readable_pages": m.get("readable_pages") or [],
                "page_mix": m.get("page_mix") or {},
                "fillability": m.get("fillability_overall"),
                "fonts": {k: v for k, v in theme.items() if k in ("major_latin", "major_ea", "minor_latin", "minor_ea")},
                "accents": [v for k, v in theme.items() if k.startswith("accent")][:6],
            }
        )
    decks.sort(key=lambda d: (-len(d["readable_pages"]), -d["slides"]))
    payload = {
        "schema": 1,
        "generated_from": "pptl.py index",
        "deck_count": len(decks),
        "page_count": sum(d["slides"] for d in decks),
        "usable_deck_count": sum(1 for d in decks if len(d["readable_pages"]) >= 6),
        "usable_page_count": sum(len(d["readable_pages"]) for d in decks),
        "decks": decks,
    }
    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(payload, ensure_ascii=False, indent=1), encoding="utf-8")
    print(
        f"{payload['deck_count']} decks / {payload['page_count']} pages "
        f"({payload['usable_deck_count']} usable, {payload['usable_page_count']} readable pages)"
    )
    print(f"manifest -> {out}  ({out.stat().st_size/1024:.0f} KB)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
