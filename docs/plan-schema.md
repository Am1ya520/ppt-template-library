# plan.json 字段说明

一份 plan 描述"要出几页、每页放什么字"。引擎按 `pages` 顺序出页，页数完全由 plan 决定。

```json
{
  "year": "2026",
  "pages": [
    {
      "type": "cover",
      "source_slide": 1,
      "title": "智能仓储解决方案",
      "body": ["产品中心 · 2026 年 3 月"],
      "slots": {"37": "技术白皮书"}
    }
  ]
}
```

## 顶层

| 字段 | 类型 | 说明 |
| --- | --- | --- |
| `year` | string | 可选。写到模板里年份占位（如 `201X`）的字；只在放得下 4 个字的框里写 |
| `pages` | array | 必填，至少 1 页 |

## 页面对象

| 字段 | 类型 | 说明 |
| --- | --- | --- |
| `type` | string | 必填：`cover` / `toc` / `chapter` / `content` / `dense` / `ending` |
| `source_slide` | int | 可选。用模板的第几页（1 起）。不给则按 `type` 自动选页 |
| `title` | string | 可选。页面主标题；`cover` / `chapter` / `content` / `dense` / `ending` 会用 |
| `body` | string[] | 可选。正文条目，一条落一个文本框（章节页可只有一条副标题） |
| `slots` | object | 可选。`{"形状id": "文字"}` 精确指定，优先于自动分配 |

### type 与模板页的关系

- 引擎会优先选 `page_type` 相同、且 `readable_pages` 里的页
- 计划写 `dense` 但模板该页是 `chapter` 时，正文不会被放进去（章节页只有标题位），所以**给 `source_slide` 时先确认那页的页型**

### body 条数 vs 模板槽位数

- 条数少：多余槽位会被清空；必要时引擎会把剩余条目并进最大的那个框
- 条数多：超出的条目会尝试并进大框，或按行删除整行（`dense`/`content` 的列表页）
- 想让版面正好，**让 body 条数等于该页可用的正文槽数**（`meta.json` 的 `readable_pages` + `index.json` 的 `slots` 可查）

## 输出与检查

`pptl.py fill` 会打印每页的填充报告，例如：

```
p 2 <- slide  8 [dense  ] filled=4 blanked=9 removed=0 recy=4 headline=#75 矩形 74 24.0pt
stripped 18 transition/animation/media nodes
```

- `filled`：落进去的条目数
- `blanked`：清掉的占位槽数
- `removed`：整行删掉的条数
- `recy`：自动改写文字颜色的处数
- `headline`：标题落在哪个形状
- `stripped`：清掉的切换/动画/音视频节点

`filled` 明显小于 `len(body)` 或某页显示 `filled=0`，说明选错页了，回去换 `source_slide`。
