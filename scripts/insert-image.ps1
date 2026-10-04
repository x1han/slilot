# 向当前演示文稿的指定页插入图片（精确页 + 坐标，供 Slilot add_image 使用）
# 用法: powershell -NoProfile -ExecutionPolicy Bypass -File scripts/insert-image.ps1 -ImagePath <path> -SlideNumber <1-based> -Left <pt> -Top <pt> -Width <pt> -Height <pt>
param([string]$ImagePath, [int]$SlideNumber, [single]$Left, [single]$Top, [single]$Width, [single]$Height)
$ErrorActionPreference = "Stop"

if (-not (Test-Path $ImagePath)) { throw ("图片文件不存在: " + $ImagePath) }

try {
  $pp = [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application")
} catch {
  throw ("未检测到正在运行的 PowerPoint。请先打开 PowerPoint 并打开演示文稿。")
}

$pres = $null
$lastErr = ""
for ($try = 1; $try -le 3 -and ($null -eq $pres); $try++) {
  try { $pres = $pp.ActivePresentation } catch { $lastErr = $_.Exception.Message; Start-Sleep -Seconds (2 * $try) }
}
if ($null -eq $pres) { throw ("no active presentation: " + $lastErr) }

if ($SlideNumber -lt 1 -or $SlideNumber -gt $pres.Slides.Count) {
  throw ("slide index out of range: " + $SlideNumber + " / " + $pres.Slides.Count)
}

$slide = $pres.Slides.Item($SlideNumber)
$shape = $slide.Shapes.AddPicture($ImagePath, 0, -1, $Left, $Top, $Width, $Height)
Write-Output ("inserted: " + $shape.Name + " on slide " + $SlideNumber)
