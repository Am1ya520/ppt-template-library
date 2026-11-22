# 一条命令发布（需要能访问 GitHub 的网络）

在这个仓库目录里执行：

```powershell
powershell -ExecutionPolicy Bypass -File tools\publish.ps1
```

脚本会做三件事，每步失败都会停下来告诉你原因：

1. 检查 `gh` 是否安装、是否已登录（没装会提示 `winget install --id GitHub.cli`，没登录会自动拉起浏览器授权）
2. 建仓库 `ppt-template-library`（已存在则跳过）并把本分支推上去
3. 用 `dist/ppt-templates.zip` 建 Release `v1.0.0` 并上传附件

只想要其中一步：

```powershell
powershell -File tools\publish.ps1 -SkipRelease    # 只推代码
powershell -File tools\publish.ps1 -SkipPush       # 只传模板包
```

参数：`-RepoName`（默认 ppt-template-library）、`-Tag`（默认 v1.0.0）、
`-Asset`（默认 dist/ppt-templates.zip）、`-Visibility public|private`。

## 如果这台机器连不上 GitHub

本机的 `curl https://github.com` 返回 `000`，直连被拦。两种情况：

- **有代理/加速工具**：先打开它，再确认 `curl -I https://github.com` 有响应；
  如果代理只对浏览器生效（系统代理），给 git 单独配：
  `git config --global http.proxy http://127.0.0.1:<端口>`
- **完全没有**：把 `dist/ppt-templates.zip` 拷到能上网的机器，按
  [PUBLISH.md](PUBLISH.md) 的"网页上传"步骤做（建仓库 → 上传 12 个文件 → Releases → 上传 zip），
  这样不需要任何命令行工具。
