---
name: ppt-template-library
description: 用本地自建模板库出稿：复制模板原有页面、只替换文字、删掉多余页，绝不从白纸自造版式。当用户要做 PPT、幻灯片、汇报、路演、培训材料、演示文稿，或要求"别太 AI""用我的模板"时使用。
---

# PPT 模板库出稿

用户的核心要求：**PPT 必须复用模板真实页面**。做法是复制模板里的原页 → 只替换文字 → 删掉没用到的页。
**绝不**用 python-pptx 从空白页画矩形/圆形自造版式——那正是"太 AI"的来源。
产出**一律去除切换动画和声音**（引擎已内置）。

## 0. 先确认库在哪

库位置解析顺序：`--library` 参数 → 环境变量 `PPTL_LIB` → 技能/仓库旁边的 `library/`。

```bash
python <repo>/pptl.py library            # 列库：按"可改页"排序
```

如果是空的，说明还没入库，走第 4 节。

## 1. 挑模板

1. 读 `library/catalog.md`，按场景 + `可改页` 列挑 **1 套**。
   **一份 deck 只用一套模板**——混用不同模板会立刻变回 AI 味。
2. 读该套 `library/deckNN/meta.json`：
   - `readable_pages`：能放正文的页码（正文槽 ≥18pt、≥3 个文本框）
   - `best_pages`：按可填充度排的推荐页
3. 需要精确槽位时读 `index.json` 或 `python <repo>/pptl.py slots <pptx> --slide N`。

## 2. 写 plan.json

```json
{
  "year": "2026",
  "pages": [
    {"type": "cover",  "source_slide": 1,  "title": "标题", "body": ["部门 · 2026 年 3 月"]},
    {"type": "toc",    "source_slide": 2,  "title": "目录", "body": ["一", "二", "三"]},
    {"type": "dense",  "source_slide": 12, "title": "小标题", "body": ["条目一……", "条目二……"]},
    {"type": "ending", "source_slide": 31, "title": "谢谢", "body": ["联系方式"]}
  ]
}
```

- `type` ∈ `cover / toc / chapter / content / dense / ending`
- 正文长度按模板原槽位写：标签 6–12 字，段落 20–40 字；超了会挤或缩
- 精确钉形状：`"slots": {"<id>": "文字"}`

## 3. 填充并检查

```bash
python <repo>/pptl.py fill "library/deckNN/source.pptx" plan.json -o out/x.pptx
```

**必须逐页看渲染**：转成 PNG 后一页页看，重点检查——

- [ ] 文字都在模板原有文字框里，没有新画的形状
- [ ] 没有残留 `Your title here` / `添加标题` / `Lorem`
- [ ] 没有大面积空白（多余槽位宁可强制填，也别留空）
- [ ] 标题没跑到装饰小框里（封面/结尾尤其容易）
- [ ] 无切换、无声音

不满意就调 `source_slide` 换页，或用 `slots` 钉死位置。

## 4. 追加模板

```bash
python <repo>/tools/import_decks.py "D:/我的模板"        # 目录或文件，可多个
```

自动去重（大小 + 首尾 1MB 哈希）、自动续编号、重写 `catalog.md`。

## 5. 交付时说明什么

- 用了哪套模板（deckNN + 模板名）
- 用了哪几页（可复现）
- 已知妥协（例如某页正文槽只有 10pt，只能放短句）

## 已知边界

- 正文槽只有 8–10pt 的页（五边形/图片卡片说明位）只当装饰页
- "单框多段 + 独立圆号"的列表页暂不支持整行删除，避开
- 模板 LOGO / 母版装饰属于版式图层，可能残留
- 缺字体时 PowerPoint 会替换，字形与原设计不同
