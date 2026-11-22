# 发布到 GitHub：把技能给出去，模板怎么给

目标：别人 clone 仓库就能用这个技能，并且**也能用上你的那批模板**。这两件事要分开做，因为模板是几 GB 的文件。

## 一、仓库里放什么（不含模板）

```
pptl.py                   引擎
tools/import_decks.py     入库
tools/export_manifest.py  清单导出
SKILL.md                  技能定义
docs/, examples/
manifest.json             模板清单（元数据，178 KB 左右，适合进 git）
```

`manifest.json` 是给人看的"你都有哪些模板"：id、模板名、页数、可改页、页型构成、字体、主题色。
**不含任何 .pptx 字节**，所以可以提交、可以 diff。

```bash
python tools/export_manifest.py --library "F:/ppt-lib/library" -o manifest.json
```

## 二、模板怎么给（三种，建议 A + C 组合）

### A. GitHub Releases 附件（推荐给"精选包"）

挑 20–30 套最能打的，压成一个 zip 挂到 Release：

```bash
# 打包（Windows PowerShell）
$ids = @("deck02","deck07","deck10","deck18")   # 换成你的精选
foreach ($id in $ids) { Copy-Item "F:/ppt-lib/library/$id/source.pptx" "pack/$id.pptx" }
Compress-Archive -Path pack/* -DestinationPath ppt-templates-starter.zip
```

```bash
# 用户侧下载并入库
gh release download v1.0.0 -p ppt-templates-starter.zip
python tools/import_decks.py ./ppt-templates-starter
```

- 单文件 < 2 GB；要发全量就分卷 `pptl-pack-01.zip …`，每个 < 2 GB
- Release 附件不计入仓库体积，clone 的人不会被几 GB 拖死

### B. 网盘链接 + 校验和（推荐给"全量 3.7 GB"）

README 里给链接和提取码，同时给出 sha256 列表，避免"下到一半坏包"：

```bash
# 生成校验清单
python - <<'PY'
import hashlib, pathlib
for p in pathlib.Path("F:/ppt-lib/library").glob("deck*/source.pptx"):
    print(hashlib.sha256(p.read_bytes()).hexdigest()[:16], p.parent.name, p.name)
PY
```

配一个 `fetch_templates.py`：用户填自己的网盘直链/本地目录，脚本负责下载 → 解压 → `import_decks.py`。
把链接放在**仓库之外**（README 里的说明），这样模板不会被 GitHub 的 DMCA 直接删仓库。

### C. 让用户自带模板（默认，永远可用）

```bash
python tools/import_decks.py "D:/我的PPT模板"
```

这是最稳的路：技能本身与模板解耦，谁的模板都能用。README 里要写清这一点——**没有模板也能跑，只是库里是空的**。

## 三、版权：必须先想清楚

你库里那 328 套大部分是第三方商业模板（"小王公司简介"系列、公众号分享的合集等）。
**把 .pptx 原文件公开分发，属于再分发他人作品，风险在你身上**，被投诉时 GitHub 会连仓库一起处理。

安全的做法，按风险从低到高：

1. 只发引擎 + manifest + 自带模板流程（不放任何 .pptx）——零风险
2. 只发**自己做的**或**明确允许再分发**（CC0/CC-BY/公司自有/买断授权）的模板
3. 全量放网盘、README 只写"自备模板"，链接另置——仍是再分发，风险不减，只是不挂在仓库里

建议：仓库走 1，Release 走 2（挑你能确认授权的），全量留在自己网盘并明确写"仅供个人学习，请勿传播"。

## 四、发版步骤

```bash
cd ppt-template-library
git init && git add . && git commit -m "feat: template-based deck engine"
gh repo create <you>/ppt-template-library --public --source=. --push
git tag v1.0.0 && git push --tags
gh release create v1.0.0 ./ppt-templates-starter.zip --notes "starter templates + engine v1.0.0"
```

`.gitignore` 已排除 `library/`、`out/`、`*.pptx`，避免 3.7 GB 被误提交：

```bash
git status --short | head          # 确认没有 library/ 和 .pptx
du -sh .git                        # 应该在几 MB 量级
```

## 五、给别人用的三条命令

README 开头就该是这三行：

```bash
pip install python-pptx
python tools/import_decks.py "你的模板目录"
python pptl.py library
```
