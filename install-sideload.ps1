# 侧载注册（免管理员）
# 官方格式（office-addin-dev-settings / dev-settings-windows.ts）：
#   键:  HKCU\Software\Microsoft\Office\16.0\Wef\Developer
#   值名: 清单 <Id>（GUID，无花括号）
#   值数据: manifest.xml 的纯文件路径
# 卸载 = 删除该值
$ErrorActionPreference = "Stop"

$addinId = "a7f3d9e2-6c48-4b1a-9d2f-3e5b7c9a1f84"
$manifestPath = "C:\Users\hxsci\ZCodeProject\ai-ppt-addin\manifest.xml"

if (-not (Test-Path $manifestPath)) { throw "找不到清单文件: $manifestPath" }

$devRoot = "HKCU:\Software\Microsoft\Office\16.0\Wef\Developer"
New-Item -Path $devRoot -Force | Out-Null

# 清理旧版错误格式（子键形式），以防残留
Remove-Item (Join-Path $devRoot "{A7F3D9E2-6C48-4B1A-9D2F-3E5B7C9A1F84}") -ErrorAction SilentlyContinue

Set-ItemProperty -Path $devRoot -Name $addinId -Value $manifestPath -Type String

Write-Output "已写入: $devRoot -> $addinId = $manifestPath"
reg.exe query "HKCU\Software\Microsoft\Office\16.0\Wef\Developer" /s
