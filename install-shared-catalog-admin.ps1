# 备用方案（需要管理员运行一次）：共享文件夹目录 + 信任目录注册
# 适用于「开发者注册表」方式没有生效的情况。
# 用法：右键"以管理员身份运行"或以管理员打开 PowerShell 执行本文件。
$ErrorActionPreference = "Stop"

$catalogDir = "C:\addin-catalog"
# 以脚本所在目录定位加载项：仓库放在任何路径都能直接运行
$addinDir = $PSScriptRoot

if (-not (Test-Path $catalogDir)) { New-Item -ItemType Directory -Path $catalogDir | Out-Null }
Copy-Item (Join-Path $addinDir "manifest.xml") $catalogDir -Force

net share addincatalog=$catalogDir /grant:Everyone,FULL 2>$null | Out-Null

$catRoot = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs"
New-Item -Path $catRoot -Force | Out-Null
$catKey = Join-Path $catRoot "{B4E8A2C1-9D3F-4E7A-8C5B-2F6D1A9E3B70}"
New-Item -Path $catKey -Force | Out-Null
Set-ItemProperty -Path $catKey -Name "Url" -Value "\\localhost\addincatalog\"
Set-ItemProperty -Path $catKey -Name "Flags" -Value 1 -Type DWord

Write-Output "共享目录与信任目录已配置。重启 PowerPoint 后：插入 > 获取加载项 > 共享文件夹。"
