#Requires -Version 5.1
$ErrorActionPreference = "Stop"

# GitHub network via local proxy (see .cursor/skills/github-proxy)
$env:HTTP_PROXY  = 'http://127.0.0.1:7890'
$env:HTTPS_PROXY = 'http://127.0.0.1:7890'
$env:ALL_PROXY   = 'http://127.0.0.1:7890'
$env:NO_PROXY    = 'localhost,127.0.0.1'
$GitProxy = @('-c', 'http.proxy=http://127.0.0.1:7890', '-c', 'https.proxy=http://127.0.0.1:7890')

$Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
Set-Location $Root

$Version = "beta0.0.1"
$ExePath = Join-Path $Root "dist\MailBatch.exe"
$RepoName = "mailBatch"

function Ensure-GhAuth {
    gh auth status *> $null
    if ($LASTEXITCODE -ne 0) {
        Write-Host "[ERROR] 未登录 GitHub。请先运行: gh auth login" -ForegroundColor Red
        exit 1
    }
}

function Ensure-Exe {
    if (-not (Test-Path $ExePath)) {
        Write-Host "Building MailBatch.exe..." -ForegroundColor Yellow
        & (Join-Path $Root "build_exe.ps1")
        if ($LASTEXITCODE -ne 0 -or -not (Test-Path $ExePath)) {
            Write-Host "[ERROR] 打包失败，未找到 $ExePath" -ForegroundColor Red
            exit 1
        }
    }
}

Ensure-GhAuth
Ensure-Exe

$remoteUrl = gh repo view "hottersquash/$RepoName" --json url -q .url 2>$null
if ($LASTEXITCODE -ne 0) {
    Write-Host "Creating GitHub repository hottersquash/$RepoName ..."
    gh repo create $RepoName --public --source=. --remote=origin `
        --description "Windows bulk email tool from Excel with rich text templates"
    if ($LASTEXITCODE -ne 0) { exit 1 }
} else {
    Write-Host "Repository exists: $remoteUrl"
    git remote get-url origin 2>$null | Out-Null
    if ($LASTEXITCODE -ne 0) {
        git remote add origin "https://github.com/hottersquash/$RepoName.git"
    }
}

git branch -M main
Write-Host "Pushing code to GitHub..."
git @GitProxy push -u origin main
if ($LASTEXITCODE -ne 0) { exit 1 }

Write-Host "Creating release $Version ..."
git tag -f $Version
git @GitProxy push -f origin $Version

$notes = @"
## MailBatch $Version (Windows Beta)

- Excel 批量邮件发送
- 富文本模板编辑器（内嵌）
- 主题/正文支持 {{列名}} 占位符
- SMTP 预设与配置导入导出

### 安装
下载 \`MailBatch.exe\`，双击运行。无需安装 Python。

### 系统要求
- Windows 10/11
- 需要 Edge WebView2 运行时（Win10/11 通常已预装）
"@

gh release create $Version $ExePath `
    --title "MailBatch $Version" `
    --notes $notes `
    --prerelease
if ($LASTEXITCODE -ne 0) {
    gh release upload $Version $ExePath --clobber
}

$releaseUrl = gh release view $Version --json url -q .url
Write-Host ""
Write-Host "Done!" -ForegroundColor Green
Write-Host "Repository: https://github.com/hottersquash/$RepoName"
Write-Host "Release:    $releaseUrl"
