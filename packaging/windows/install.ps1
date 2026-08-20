$ErrorActionPreference = "Stop"

$PackagePath = Get-ChildItem -Path $PSScriptRoot -Filter "webmail-cli-*.tgz" | Select-Object -First 1
if (-not $PackagePath) {
    throw "安装包损坏：找不到 webmail-cli-*.tgz。"
}

$Node = Get-Command node.exe -ErrorAction SilentlyContinue
$Npm = Get-Command npm.cmd -ErrorAction SilentlyContinue
if (-not $Node -or -not $Npm) {
    throw "需要先安装 Node.js 24 或更高版本：https://nodejs.org/"
}

$NodeMajor = [int](& node.exe -p "process.versions.node.split('.')[0]")
if ($NodeMajor -lt 24) {
    throw "当前 Node.js 版本是 $(& node.exe --version)，需要 24 或更高版本。"
}

$InstallRoot = Join-Path $env:LOCALAPPDATA "webmail-cli"
$BinDirectory = Join-Path $InstallRoot "bin"
New-Item -ItemType Directory -Force -Path $InstallRoot, $BinDirectory | Out-Null

Write-Host "正在安装 webmail-cli，请保持网络连接……"
& npm.cmd install --prefix $InstallRoot --omit=dev --no-audit --no-fund $PackagePath.FullName
if ($LASTEXITCODE -ne 0) {
    throw "npm 安装失败，退出码：$LASTEXITCODE"
}

$WebmailLauncher = @'
@echo off
node "%LOCALAPPDATA%\webmail-cli\node_modules\webmail-cli\dist\cli.js" %*
'@
$McpLauncher = @'
@echo off
node "%LOCALAPPDATA%\webmail-cli\node_modules\webmail-cli\dist\mcp\stdio.js" %*
'@
Set-Content -Path (Join-Path $BinDirectory "webmail.cmd") -Value $WebmailLauncher -Encoding Ascii
Set-Content -Path (Join-Path $BinDirectory "webmail-mcp.cmd") -Value $McpLauncher -Encoding Ascii

$UserPath = [Environment]::GetEnvironmentVariable("Path", "User")
$PathEntries = @($UserPath -split ";" | Where-Object { $_ })
if ($PathEntries -notcontains $BinDirectory) {
    $NewPath = (($PathEntries + $BinDirectory) -join ";")
    [Environment]::SetEnvironmentVariable("Path", $NewPath, "User")
}
$env:Path = "$BinDirectory;$env:Path"

$InstalledVersion = & webmail.cmd --version
Write-Host ""
Write-Host "安装成功：webmail-cli $InstalledVersion" -ForegroundColor Green
Write-Host "下一步请打开新的 PowerShell 或 CMD，运行：webmail status --json"
Write-Host "如果浏览器提示登录，请完成登录后再次运行该命令。"
