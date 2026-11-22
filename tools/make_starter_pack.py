"""Build a starter pack of templates for a GitHub Release.

    python tools/make_starter_pack.py --list                 # what would be picked
    python tools/make_starter_pack.py -n 24 -o pack          # copy .pptx into pack/
    python tools/make_starter_pack.py -n 24 -o pack --zip    # ... and zip it

Selection: decks with the most readable pages, at most `--per-family` decks that
share the same dominant accent colour, so the pack is not 24 copies of one style.

Nothing here uploads anything: attach the resulting zip to a Release yourself,
and only ship templates you are allowed to redistribute (see PUBLISH.md).
"""
from __future__ import annotations

import argparse
import json
import os
import shutil
import zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent


def library_root(explicit: str | None = None) -> Path:
    if explicit:
        return Path(explicit).expanduser().resolve()
    env = os.environ.get("PPTL_LIB")
    if env:
        return Path(env).expanduser().resolve()
    return REPO / "library"


def pick(lib: Path, limit: int, per_family: int, min_readable: int) -> list[dict]:
    rows = []
    for meta_file in lib.glob("deck*/meta.json"):
        m = json.loads(meta_file.read_text(encoding="utf-8"))
        readable = m.get("readable_pages") or []
        if len(readable) < min_readable:
            continue
        accents = (m.get("theme") or {}).get("accent1") or ""
        rows.append(
            {
                "id": m["deck_id"],
                "name": m["source_name"],
                "slides": m["slides"],
                "readable": len(readable),
                "family": (accents or "none").upper(),
            }
        )
    rows.sort(key=lambda r: (-r["readable"], -r["slides"]))
    out, fam = [], {}
    for row in rows:
        if len(out) >= limit:
            break
        if fam.get(row["family"], 0) >= per_family:
            continue
        fam[row["family"]] = fam.get(row["family"], 0) + 1
        out.append(row)
    return out


def main(argv=None) -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--library")
    ap.add_argument("-n", "--count", type=int, default=24)
    ap.add_argument("--per-family", type=int, default=4)
    ap.add_argument("--min-readable", type=int, default=8)
    ap.add_argument("-o", "--out", default="pack")
    ap.add_argument("--zip", action="store_true", help="also write <out>.zip")
    ap.add_argument("--list", action="store_true", help="only print the selection")
    args = ap.parse_args(argv)

    lib = library_root(args.library)
    chosen = pick(lib, args.count, args.per_family, args.min_readable)
    if not chosen:
        print("nothing qualifies; lower --min-readable or import more templates")
        return 1

    total_mb = 0.0
    for row in chosen:
        src = lib / row["id"] / "source.pptx"
        size = src.stat().st_size / 1024 / 1024 if src.exists() else 0
        total_mb += size
        print(f"{row['id']:<8}{row['readable']:>4}可改页  {size:>6.1f}MB  {row['name'][:50]}")
    print(f"\n{len(chosen)} decks, {total_mb:.0f} MB")
    if args.list:
        return 0

    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    for row in chosen:
        src = lib / row["id"] / "source.pptx"
        if src.exists():
            shutil.copy2(src, out / f"{row['id']}_{row['name'][:40]}.pptx")
    index = out / "PACK.md"
    index.write_text(
        "# 精选模板包\n\n| id | 模板 | 页数 | 可改页 |\n| --- | --- | --- | --- |\n"
        + "\n".join(
            f"| {r['id']} | {r['name']} | {r['slides']} | {r['readable']} |" for r in chosen
        )
        + "\n\n把这些 .pptx 放一个目录，然后：\n\n```bash\npython tools/import_decks.py <该目录>\n```\n",
        encoding="utf-8",
    )
    if args.zip:
        zip_path = out.with_suffix(".zip")
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for file in out.iterdir():
                zf.write(file, file.name)
        print(f"zip -> {zip_path}  ({zip_path.stat().st_size/1024/1024:.0f} MB)")
    print(f"pack -> {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
