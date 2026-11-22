# ppt-template-library

> 仓库：<https://github.com/Am1ya520/ppt-template-library>　成品模板包：<https://github.com/Am1ya520/ppt-template-library/releases>

**用你自己的 PPT 模板出稿，而不是让 AI 从白纸画版式。**

给 [DeepSeek Harness](https://github.com/deepseek-ai) 用的技能 + 独立 CLI：它把你硬盘里的一堆 `.pptx` 模板编成索引，
出稿时**复制模板原有的页面、只替换文字、删掉没用到的页**——所以成品看起来就是那套模板做的，没有"AI 味儿"。

- 引擎只依赖 `python-pptx`，单文件 `pptl.py`
- 一份 deck 只用一套模板（混用会立刻回到 AI 感）
- 产出一律**没有切换动画、没有声音**
- 附带一个索引/套壳工具链：`index`（逐页槽位）→ `fill`（套内容）→ `qa_deck`（交付前静态质检） → `catalog`（库总表）

> 本仓库不含任何模板文件。模板是你自己的资产，通过 `tools/import_decks.py` 从本地目录入库；
> 想把模板也分发出去，见 [PUBLISH.md](PUBLISH.md)。

## 它解决什么问题

| 常见做法 | 结果 |
| --- | --- |
| 让 AI 用 python-pptx 从空白页画矩形/圆形 | 版式简单、颜色一套模板味，一眼 AI |
| 让 AI 生成 HTML/图片式的"设计稿" | 看着花，但不是能改的 PPT |
| **复制模板真实页面，只换字**（本仓库） | 和模板原样一致，客户/评委看不出是生成的 |

## 安装（三条命令）

```bash
git clone https://github.com/Am1ya520/ppt-template-library.git
cd ppt-template-library
pip install -r requirements.txt          # 只有 python-pptx + lxml
python tools/bootstrap.py "D:/我的模板"    # 检查环境 + 建库 + 把模板全部入库（可反复跑，自动去重）
```

`bootstrap.py` 不带参数只检查环境；跑完会打印接下来要用的命令。

### 作为 DeepSeek Harness 技能使用

技能就是本仓库的 `SKILL.md`。放进 DSH 的 skills 目录后，任何对话里说"做个 PPT / 幻灯片 / 汇报"就会自动加载：

```bash
# Windows (PowerShell)
git clone https://github.com/Am1ya520/ppt-template-library.git "$env:USERPROFILE\.dsh\skills\ppt-template-library"
# macOS / Linux
git clone https://github.com/Am1ya520/ppt-template-library.git ~/.dsh/skills/ppt-template-library
```

装完**重启 DSH**（或新开一个会话）才会出现在技能列表里。

### 库很大？放到别的盘

仓库里不放模板。默认库是仓库内的 `library/`；模板多（几个 GB 以上）时建议放独立目录，用环境变量指过去，技能和 CLI 都会跟着用：

```bash
# Windows
setx PPTL_LIB "F:\pptl-lib"
# macOS / Linux
export PPTL_LIB="$HOME/pptl-lib"      # 写进 ~/.zshrc 或 ~/.bashrc
```

也可以每次显式传 `--library D:/pptl-lib`。

### 手上没有模板？

[Releases](https://github.com/Am1ya520/ppt-template-library/releases) 里有整理好的通用模板包（年终总结、财务报告、商务策划、创业 BP、公司简介、企业培训、毕业答辩等 28 套），下载解压后：

```bash
python tools/import_decks.py "解压出来的目录"      # 一次性全部入库
```

### 作为普通 CLI 使用

```bash
python pptl.py library --limit 20      # 按"可改页"排序看库
python pptl.py --help
```

## 一、把模板入库

```bash
python tools/import_decks.py "D:/我的模板"                 # 目录（递归）或单个文件，可多个
python tools/import_decks.py "D:/模板A" "D:/模板B.pptx" --scenario 汇报
```

- 自动去重：`文件大小 + 首尾 1MB 的 sha1`（**不整份读**，几 GB 的库也快）
- 自动跳过 `~$` 临时文件和不足 5 页的空文档
- 每套落成 `library/deckNN/{source.pptx, index.json, meta.json}`，并重写 `library/catalog.md`

看库：

```bash
python pptl.py library              # 按"可改页"排序
cat library/catalog.md              # 总表：页数 / 可改页 / 各类页数量 / 可填充度
```

## 二、出稿

写一份 `plan.json`：

```json
{
  "year": "2026",
  "pages": [
    {"type": "cover",   "source_slide": 1,  "title": "智能仓储解决方案", "body": ["产品中心 · 2026 年 3 月"]},
    {"type": "toc",     "source_slide": 2,  "title": "目录", "body": ["行业现状与痛点", "总体架构", "试点节奏"]},
    {"type": "dense",   "source_slide": 12, "title": "总体架构", "body": ["感知层：…", "执行层：…", "决策层：…"]},
    {"type": "ending",  "source_slide": 31, "title": "谢谢", "body": ["产品中心"]}
  ]
}
```

```bash
python pptl.py fill "library/deck02/source.pptx" plan.json -o out/demo.pptx
```

- `type` ∈ `cover / toc / chapter / content / dense / ending`
- 不给 `source_slide` 时按类型自动挑页（优先 `meta.json` 的 `readable_pages`）
- 想精确钉到某个形状：`"slots": {"<形状 id>": "文字"}`，形状 id 用 `pptl.py slots <pptx> --slide N` 查
- 字段细节见 [docs/plan-schema.md](docs/plan-schema.md)

**出稿后一定渲染逐页检查**（LibreOffice / PowerPoint / 任何转图工具都行）。文字溢出、空卡片这类问题只有看图才能发现。

## 三、引擎做了什么

`pptl.py` 的核心不是"排版"，而是**判断模板上哪个框是什么角色**，然后把字准确地放进去：

| 机制 | 作用 |
| --- | --- |
| `classify()` | 把每个文字框分成 `headline / text / year / num / design / furniture / filler / empty / rotated` |
| `furniture` | `目录 / CONTENTS / 谢谢 / THANK YOU / MANY THANKS` 是版式固定文字，原样保留，不被覆盖 |
| `is_filler()` | 识别中英占位符（`此处`、`添加标题`、`Your title here`、`CONTENT TITLE`、`Click to add…`），清掉但不留洞 |
| `pick_headline()` | 封面/结尾这类居中版式按页面水平中心取标题；年份只写进放得下 4 个字的框 |
| `fit_font()` | 按框的真实可用行数缩字号，缩到下限仍放不下就换一个框 |
| `paint_behind()` | 判断文字背后是什么颜色（含渐变、图片），自动换黑/白字，保住可读性 |
| 拥挤/强制放置 | 短句进矮框不算挤；整页候选全挤时强制放置——宁可紧一点，也不交空白页 |
| `strip_effects()` | 保存前清掉 `p:transition`、`p:timing`、音频视频关系：产出无切换无声音 |

## 想把模板也发出去

一条命令（需要能访问 GitHub 的网络）：

```powershell
powershell -ExecutionPolicy Bypass -File tools\publish.ps1
```

细看 [tools/RELEASE-HOWTO.md](tools/RELEASE-HOWTO.md)（含"机器连不上 GitHub"的退路）。手动做法：


```bash
python tools/make_starter_pack.py --list                 # 看看会挑哪几套（按可改页排序，同色系限量）
python tools/make_starter_pack.py -n 24 -o pack --zip    # 复制成 pack/ 并打包 zip
```

选出来的是"最能放正文"的那几套，并且限制了同色系套数，不会变成 24 份一个风格。
**发布前先读 [PUBLISH.md](PUBLISH.md) 的版权一节**——第三方商业模板不适合公开再分发。

## 目录结构

```
pptl.py                 引擎（index / slots / fill / library）
tools/import_decks.py   模板入库 + 目录重建
tools/export_manifest.py 导出 manifest.json（只含元数据，不含模板文件）
tools/make_starter_pack.py 挑精选模板打包（GitHub Release 用）
SKILL.md                DeepSeek Harness 技能定义
docs/plan-schema.md     plan.json 字段说明
examples/plan-demo.json 可直接跑的示例
manifest.json           **Release 那 28 套包的清单**（862 页 / 750 可改页）：包里有哪套、各多少页、哪些页能放正文。你自己的库用 tools/export_manifest.py 另导一份
library/                你的模板库（.gitignore，不入库）
```

## 可移植性

- **Python 3.9+**（开发与实测环境 3.12），只依赖 `python-pptx` + `lxml`，无其他服务
- Windows / macOS / Linux 都能跑（路径全部走 `pathlib`）；`.ppt` 老格式需要本机装 LibreOffice 做一次转换
- 全新机器实测流程：`git clone` → `pip install -r requirements.txt` → `tools/bootstrap.py` → `import_decks.py` → `pptl.py fill` → `qa_deck.py`，全程 5 分钟，无需任何配置

库位置解析顺序：`--library` 参数 → 环境变量 `PPTL_LIB` → 仓库旁边 `./library`。

```bash
export PPTL_LIB=/data/ppt-library     # 库放别处
python pptl.py library
```

## 已知边界

- 有些模板的正文槽只有 8–10pt（五边形/图片卡片的说明位）：这类页只当装饰页，别放段落
- "单框多段 + 独立圆号"的列表页暂不支持整行删除，出稿避开
- 模板 LOGO / 母版装饰属于版式图层，脚本读不到幻灯片上的文本，可能残留（需在母版里处理）
- 中文字体缺失会被替换，字形可能和模板原设计不同；把模板用到的字体装到出稿机器上更稳

## License

引擎与脚本：MIT（见 [LICENSE](LICENSE)）。本仓库**不含任何模板文件**。

[Releases](https://github.com/Am1ya520/ppt-template-library/releases) 里的模板包是**第三方作者的模板作品**，来自公开分享的合集，
版权归各自作者。仓库不主张版权、不以此收费，仅提供"复制模板原页、只换文字"这套方法的现成素材：

- 允许：个人学习、内部汇报、非公开场合使用
- 不允许：转售、二次打包收费、去作者信息后当成自己的模板出售
- 商用请向原作者或正规模板平台购买授权
- 作者/权利人要求下架：开 issue 说明是哪一套，24 小时内移除

对授权要求严格的组织，建议不要用模板包，改用 `tools/import_decks.py` 把**你们已授权的模板**导入本地库——
引擎本身不需要这些文件。
