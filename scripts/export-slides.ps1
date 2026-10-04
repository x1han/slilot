# 将当前 PowerPoint 中打开的活动演示文稿渲染为逐页 PNG（供 Slilot 视觉审查）
# 用法: powershell -NoProfile -ExecutionPolicy Bypass -File scripts/export-slides.ps1 -OutDir <dir>
# 输出: OutDir\slide-<n>.png（文件名排序即页序）
# 注意:
#  - 只附着到已运行的 PowerPoint 实例（GetActiveObject），绝不拉起新实例，避免空实例残留
#  - PowerPoint COM 是单实例；用户正在操作（模态对话框等）时可能失败，此处带重试
param([string]$OutDir)
$ErrorActionPreference = "Stop"

try {
  $pp = [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application")
} catch {
  throw ("未检测到正在运行的 PowerPoint。请先打开 PowerPoint 并打开要审查的演示文稿。")
}

# 附着到用户正在使用的实例，取当前活动的演示文稿（含未保存修改的实时状态）
$pres = $null
$lastErr = ""
for ($try = 1; $try -le 3 -and ($null -eq $pres); $try++) {
  try {
    $pres = $pp.ActivePresentation
  } catch {
    $lastErr = $_.Exception.Message
    Write-Output ("active attempt " + $try + " failed: " + $lastErr)
    Start-Sleep -Seconds (2 * $try)
  }
}
if ($null -eq $pres) { throw ("no active presentation: " + $lastErr) }

# 只导出，不关闭、不退出（这是用户正在编辑的文档）
$count = $pres.Slides.Count
for ($i = 1; $i -le $count; $i++) {
  $png = Join-Path $OutDir ("slide-" + $i + ".png")
  $pres.Slides.Item($i).Export($png, "PNG", 1280, 720)
}
Write-Output ("exported " + $count + " slides")
