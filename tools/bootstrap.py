"""One-shot bootstrap: check deps, create the library, import templates, print next steps.

    python tools/bootstrap.py                          # 只检查环境 + 建空库
    python tools/bootstrap.py --templates "D:/我的模板"   # 顺便把整个目录入库
    python tools/bootstrap.py --templates a.pptx b.pptx --library D:/pptl-lib

Exit code 0 = ready to use. Safe to run twice (import skips duplicates).
"""
from __future__ import annotations

import argparse
import importlib.util
import os
import subprocess
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent


def check_deps() -> bool:
    missing = [m for m in ("pptx", "lxml") if importlib.util.find_spec(m) is None]
    if not missing:
        print("[ok] python 依赖齐全 (python-pptx, lxml)")
        return True
    print(f"[!] 缺依赖: {', '.join(missing)}")
    print("    安装:  pip install -r requirements.txt")
    return False


def library_root(explicit: str | None) -> Path:
    if explicit:
        return Path(explicit).expanduser().resolve()
    env = os.environ.get("PPTL_LIB")
    if env:
        return Path(env).expanduser().resolve()
    return REPO / "library"


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(prog="bootstrap")
    ap.add_argument("paths", nargs="*", help="模板文件或目录（可多个，等价于 --templates）")
    ap.add_argument("--templates", nargs="*", help="模板文件或目录（可多个）")
    ap.add_argument("--library", help="库位置（默认 $PPTL_LIB 或仓库内 ./library）")
    args = ap.parse_args(argv)

    ok = check_deps()
    lib = library_root(args.library)
    lib.mkdir(parents=True, exist_ok=True)
    decks = [p for p in lib.glob("deck*") if (p / "meta.json").exists()]
    print(f"[ok] 库目录: {lib}  (现有 {len(decks)} 套)")

    if not ok:
        return 2

    inputs = (args.templates or []) + args.paths
    if inputs:
        rc = subprocess.run(
            [sys.executable, str(REPO / "tools" / "import_decks.py"), *inputs, "--library", str(lib)]
        )
        if rc.returncode != 0:
            return rc.returncode
    elif not decks:
        print("\n还没有模板。下一步任选其一：")
        print(f'  python tools/import_decks.py "<你的模板目录>" --library "{lib}"')
        print("  或先下载成品包 https://github.com/Am1ya520/ppt-template-library/releases 解压后导入")

    print("\n下一步：")
    print(f'  python pptl.py library --library "{lib}"        # 看库里有什么')
    print(f'  python pptl.py fill "{lib}/deckNN/source.pptx" examples/plan-demo.json -o out/demo.pptx')
    print('  python tools/qa_deck.py out/demo.pptx         # 0 error 才交付')
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
