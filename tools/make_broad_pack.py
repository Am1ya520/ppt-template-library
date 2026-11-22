"""Pick a broad, general-purpose template pack from the local library.

Rules (v1.2.0):
  * only decks whose pages are actually re-fillable (readable_pages >= 8, fillability >= 0.6)
  * general-purpose only -- political / seasonal / industry-specific / IP-character
    names are excluded, as are decks tied to one regulatory body (法院/检察院/水利...)
  * size 0.6-25 MB, 10-60 slides
  * at most 2 decks per visual family (same theme fonts + accent colours),
    at most 3 per name series, so the pack spans many looks instead of one shop
  * scenarios get a quota so every common use case has a few choices

    python tools/make_broad_pack.py --library ../library --out ../gh/dist/pack-v120 --list
"""
from __future__ import annotations

import argparse
import json
import re
import shutil
import unicodedata
from collections import Counter, defaultdict
from pathlib import Path

BAN = (
    "党建|党政|党史|入党|两学一做|四讲四有|精准扶贫|扶贫|国庆|建军|军|警|消防|两会|"
    "十九大|二十大|政府|正府|纪委|纪检|廉政|廉洁|雷锋|五一|春节|中秋|教师节|母亲节|"
    "儿童节|圣诞|元旦|清明|端午|七夕|重阳|悟空|库洛米|动漫|游戏|明星|招商|考研|招生|"
    "简历|婚礼|宝宝|军训|校庆|运动会|医院|医疗|银行|保险|微商|带货|直播|抖音|小红书|"
    "法院|检察|司法|公检法|水利|路政|电网|电力|工商|税务|环保|生态|创卫|安全生产|"
    "扫黑除恶|禁毒|人防|民兵|征兵|妇女|工会|团委|少先队|党委|党支部|党员|计划生育|"
    "人口普查|职业规划|大赛|动车|火车|铁路|航空|烟草|石油|矿山|钢铁|物业|酒店|餐饮|"
    "驾校|美容|健身|婚纱|景区|旅行社|乡村振兴|土地|拆迁|信访|应急|防汛|抗旱"
)

SCENARIOS = [
    ("汇报总结", ("总结", "汇报", "述职", "年报")),
    ("商业计划", ("创业", "计划书", "融资", "bp", "商业", "策划")),
    ("公司简介", ("公司简介", "企业宣传", "企业简介", "宣传")),
    ("企业培训", ("培训", "入职", "员工", "目标管理", "团建")),
    ("学术答辩", ("答辩", "论文", "学术", "开题", "毕业")),
    ("极简商务", ("极简", "简约", "白金", "黑白", "商务")),
    ("互联网科技", ("互联网", "科技", "信息", "数据", "智能", "数字")),
    ("教育文化", ("教育", "教学", "课程", "国学", "中国风", "文化")),
    ("项目方案", ("项目", "方案", "计划", "规划", "实施")),
]

# how many decks each common scenario gets in a 50-deck pack
QUOTA = defaultdict(
    lambda: 5,
    {
        "汇报总结": 8,
        "商业计划": 7,
        "极简商务": 7,
        "企业培训": 6,
        "学术答辩": 5,
        "项目方案": 5,
        "公司简介": 4,
        "互联网科技": 4,
        "教育文化": 4,
    },
)

SKIP_NAME = re.compile(r"^\s*(1|新建|演示文稿|ppt|pptx)\s*(\(\d+\))?\s*$", re.I)


def norm(name: str) -> str:
    return unicodedata.normalize("NFKC", name).strip().lower()


def scenario_of(name: str) -> str | None:
    low = norm(name)
    for label, keys in SCENARIOS:
        if any(k in low for k in keys):
            return label
    return None


def series_of(name: str) -> str:
    """Rough shop/series key: drop trailing (n), digits and punctuation."""
    s = re.sub(r"[\(（]\d+[\)）]", "", name)
    s = re.sub(r"[\s_\-·.]+", "", s)
    s = re.sub(r"\d+$", "", s)
    return s[:6] or name[:6]


def family_of(meta: dict) -> str:
    """Visual family: theme ink + the first two accents (that is what the eye sees)."""
    theme = meta.get("theme") or {}
    ink = (theme.get("dk1") or "")[:2]
    a1 = (theme.get("accent1") or "")[:4]
    a2 = (theme.get("accent2") or "")[:4]
    if not (ink or a1 or a2):
        return f"plain|{(meta.get('slides') or 0)//10}"
    return f"{ink}|{a1}|{a2}"


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(prog="make_broad_pack")
    ap.add_argument("--library", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--target", type=int, default=50)
    ap.add_argument("--list", action="store_true", help="dry run: print the selection")
    args = ap.parse_args(argv)

    lib = Path(args.library).expanduser().resolve()
    cands = []
    for meta_path in lib.glob("deck*/meta.json"):
        try:
            meta = json.loads(meta_path.read_text(encoding="utf-8"))
        except Exception:
            continue
        name = meta.get("source_name") or ""
        if not name or SKIP_NAME.match(name) or re.search(BAN, name):
            continue
        readable = len(meta.get("readable_pages") or [])
        if readable < 8 or (meta.get("fillability_overall") or 0) < 0.6:
            continue
        if not 10 <= meta.get("slides", 0) <= 60:
            continue
        src = meta_path.parent / "source.pptx"
        if not src.exists():
            continue
        size = src.stat().st_size
        if not 600 * 1024 <= size <= 25 * 1024 * 1024:
            continue
        scenario = scenario_of(name)
        if scenario is None:
            continue
        cands.append(
            {
                "id": meta["deck_id"],
                "name": name,
                "scenario": scenario,
                "readable": readable,
                "slides": meta["slides"],
                "size": size,
                "family": family_of(meta),
                "series": series_of(name),
                "src": src,
            }
        )

    cands.sort(key=lambda c: (-c["readable"], c["size"]))
    picked = []
    fam_used, ser_used = Counter(), Counter()
    for fam_cap, ser_cap, quota_scale in ((2, 3, 1.0), (3, 3, 1.25), (4, 3, 1.6), (99, 3, 99)):
        per_scenario = Counter(c["scenario"] for c in picked)
        for c in cands:
            if len(picked) >= args.target:
                break
            if any(c["id"] == p["id"] for p in picked):
                continue
            if fam_used[c["family"]] >= fam_cap or ser_used[c["series"]] >= ser_cap:
                continue
            if per_scenario[c["scenario"]] >= QUOTA[c["scenario"]] * quota_scale:
                continue
            fam_used[c["family"]] += 1
            ser_used[c["series"]] += 1
            per_scenario[c["scenario"]] += 1
            picked.append(c)
        if len(picked) >= args.target:
            break

    picked.sort(key=lambda c: (c["scenario"], -c["readable"]))
    per_scenario = Counter(c["scenario"] for c in picked)
    total = sum(c["size"] for c in picked)
    print(f"{len(picked)} decks, {total / 1048576:.1f} MB, scenarios: {dict(per_scenario)}")
    for i, c in enumerate(picked, 1):
        label = f"{i:02d} {c['scenario']}-{c['name'][:24]}"
        c["label"] = re.sub(r'[\\/:*?"<>|]', "_", label)
        print(f"  {c['label']:<44} {c['readable']:>3}改页 {c['slides']:>3}p {c['size']/1048576:>5.1f}MB  {c['id']}")

    if args.list:
        return 0

    out = Path(args.out).expanduser().resolve()
    if out.exists():
        shutil.rmtree(out)
    out.mkdir(parents=True)
    for c in picked:
        shutil.copy2(c["src"], out / f"{c['label']}.pptx")
    (Path(args.out).parent / "pack-v120.json").write_text(
        json.dumps(
            [{"label": c["label"], "id": c["id"], "name": c["name"], "slides": c["slides"],
              "readable": c["readable"], "size": c["size"], "scenario": c["scenario"]} for c in picked],
            ensure_ascii=False,
            indent=1,
        ),
        encoding="utf-8",
    )
    print(f"written -> {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
