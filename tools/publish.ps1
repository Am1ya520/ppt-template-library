param(
    [string]$Owner = "Am1ya520",
    [string]$RepoName = "ppt-template-library",
    [string]$Tag = "v1.0.0",
    [string]$Asset = "dist/ppt-templates.zip",
    [ValidateSet("public", "private")][string]$Visibility = "public",
    [switch]$SkipPush,
    [switch]$SkipRelease
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
Set-Location $root

function Fail($msg) { Write-Host "✗ $msg" -ForegroundColor Red; exit 1 }
function Ok($msg)   { Write-Host "✓ $msg" -ForegroundColor Green }

if (-not (Get-Command gh -ErrorAction SilentlyContinue)) {
    Fail "没有 gh。先装：winget install --id GitHub.cli  然后重开一个终端。"
}

gh auth status 2>$null | Out-Null
if ($LASTEXITCODE -ne 0) {
    Write-Host "需要登录 GitHub（会打开浏览器）..." -ForegroundColor Yellow
    gh auth login -h github.com -w
    if ($LASTEXITCODE -ne 0) { Fail "登录没有完成。" }
}
Ok "gh 已登录"

$user = gh api user --jq .login
if ($LASTEXITCODE -ne 0 -or -not $user) { Fail "拿不到 GitHub 用户名，检查网络。" }
gh auth setup-git 2>$null | Out-Null
Ok "账号：$user（已配置 git 凭据）"
if ($Owner -and $user -ne $Owner) { Write-Host "⚠ 当前登录的是 $user，不是 $Owner；产物会落在 $user 名下。" -ForegroundColor Yellow }

if (-not $SkipPush) {
    $hasRemote = (git remote) -contains "origin"
    if (-not $hasRemote) {
        $prev = $ErrorActionPreference
        $ErrorActionPreference = "SilentlyContinue"
        gh repo view "$user/$RepoName" *> $null
        $exists = ($LASTEXITCODE -eq 0)
        $ErrorActionPreference = $prev
        if (-not $exists) {
            gh repo create "$user/$RepoName" --$Visibility --source=. --remote=origin --push
            if ($LASTEXITCODE -ne 0) { Fail "建仓库失败（多半是网络）。" }
            Ok "仓库已建并推送：https://github.com/$user/$RepoName"
        } else {
            git remote add origin "https://github.com/$user/$RepoName.git"
            git push -u origin HEAD
            Ok "已推到已有仓库"
        }
    } else {
        git push -u origin HEAD
        Ok "已推送"
    }
}

if (-not $SkipRelease) {
    if (-not (Test-Path $Asset)) {
        Fail "找不到 $Asset。先生成：python tools/make_starter_pack.py --ids ... -o dist/ppt-templates --zip"
    }
    $sizeMb = [math]::Round((Get-Item $Asset).Length / 1MB)
    Write-Host "上传 $Asset ($sizeMb MB)..." -ForegroundColor Yellow
    git tag $Tag 2>$null
    git push origin $Tag 2>$null
    $notes = "dist/RELEASE_NOTES.md"
    if (-not (Test-Path $notes)) { Fail "缺少 $notes（Release 说明），不要用内联字符串——PowerShell/ bash 都会吃掉反引号和尖括号。" }
    gh release create $Tag $Asset --title "$Tag · 18 套成品模板" --notes-file $notes
    if ($LASTEXITCODE -ne 0) { Fail "Release 上传失败。" }
    Ok "Release 已发布：https://github.com/$user/$RepoName/releases/tag/$Tag"
}

Write-Host "`n完成。" -ForegroundColor Cyan
