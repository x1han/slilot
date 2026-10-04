# 将当前 PowerPoint 中打开的活动演示文稿渲染为逐页 PNG（供 Slilot 视觉审查）
# 用法: powershell -NoProfile -ExecutionPolicy Bypass -File scripts/export-slides.ps1 -OutDir <dir>
# 输出: OutDir\slide-<n>.png（文件名排序即页序）
# 注意:
#  - PowerPoint 单实例。COM 拉起的实例不注册到 ROT，所以 GetActiveObject 失败时
#    必须回退 New-Object（CoCreateInstance 会附着到已运行进程）
#  - 若实例里没有任何打开的演示文稿（可能是本脚本拉起的空实例），退出并清理
param([string]$OutDir)
$ErrorActionPreference = "Stop"

$pp = $null
try { $pp = [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application") } catch {}
if ($null -eq $pp) {
  try { $pp = New-Object -ComObject PowerPoint.Application } catch {
    throw ("无法连接 PowerPoint: " + $_.Exception.Message)
  }
}

$pres = $null
$lastErr = ""
for ($try = 1; $try -le 3 -and ($null -eq $pres); $try++) {
  try { $pres = $pp.ActivePresentation } catch { $lastErr = $_.Exception.Message; Start-Sleep -Seconds (2 * $try) }
}
if ($null -eq $pres) {
  # 清理可能由本脚本拉起的空实例（没有任何打开的文稿时 Quit 是安全的）
  try { if ($pp.Presentations.Count -eq 0) { $pp.Quit() } } catch {}
  throw ("无法获取活动演示文稿（请确认 PowerPoint 已打开演示文稿）: " + $lastErr)
}

$count = $pres.Slides.Count
for ($i = 1; $i -le $count; $i++) {
  # JPG + 1024 宽：截图只供模型视觉审查，控制体积（PNG 全尺寸单张可达数百 KB）
  $png = Join-Path $OutDir ("slide-" + $i + ".jpg")
  $pres.Slides.Item($i).Export($png, "JPG", 1024, 576)
}
Write-Output ("exported " + $count + " slides")
