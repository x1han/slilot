# 向当前演示文稿的指定页插入图片（精确页 + 坐标，供 Slilot add_image 使用）
# 关键行为：width/height 视为边界框，图片按原始宽高比 contain 缩放后在框内居中——
#          无论模型给的框比例是否与图片一致，都不会拉伸变形。
# 用法: powershell -NoProfile -ExecutionPolicy Bypass -File scripts/insert-image.ps1 -ImagePath <path> -SlideNumber <1-based> -Left <pt> -Top <pt> -Width <pt> -Height <pt>
# 注意: PowerPoint 单实例。COM 拉起的实例不注册到 ROT，GetActiveObject 失败时
#       必须回退 New-Object（CoCreateInstance 会附着到已运行进程）
param([string]$ImagePath, [int]$SlideNumber, [single]$Left, [single]$Top, [single]$Width, [single]$Height)
$ErrorActionPreference = "Stop"

if (-not (Test-Path $ImagePath)) { throw ("图片文件不存在: " + $ImagePath) }

Add-Type -AssemblyName System.Drawing
$img = [System.Drawing.Image]::FromFile($ImagePath)
$imgWpx = $img.Width
$imgHpx = $img.Height
$img.Dispose()

# 像素 -> pt（按 96dpi：1px = 0.75pt），得到图片自然尺寸
$natW = $imgWpx * 0.75
$natH = $imgHpx * 0.75

if ($Width -gt 0 -and $Height -gt 0) {
  # contain：取较小的缩放系数，保证整图都在框内且不变形
  $scale = [Math]::Min($Width / $natW, $Height / $natH)
} elseif ($Width -gt 0) {
  $scale = $Width / $natW
} elseif ($Height -gt 0) {
  $scale = $Height / $natH
} else {
  $scale = 1.0
}

$placeW = [Math]::Round($natW * $scale, 2)
$placeH = [Math]::Round($natH * $scale, 2)
# 在给定框内居中
$placeLeft = [Math]::Round($Left + ($Width - $placeW) / 2, 2)
$placeTop = [Math]::Round($Top + ($Height - $placeH) / 2, 2)

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
  try { if ($pp.Presentations.Count -eq 0) { $pp.Quit() } } catch {}
  throw ("无法获取活动演示文稿（请确认 PowerPoint 已打开演示文稿）: " + $lastErr)
}

if ($SlideNumber -lt 1 -or $SlideNumber -gt $pres.Slides.Count) {
  throw ("slide index out of range: " + $SlideNumber + " / " + $pres.Slides.Count)
}

$slide = $pres.Slides.Item($SlideNumber)
$shape = $slide.Shapes.AddPicture($ImagePath, 0, -1, $placeLeft, $placeTop, $placeW, $placeH)
Write-Output ("inserted: " + $shape.Name + " on slide " + $SlideNumber)
Write-Output ("PLACED left=$placeLeft top=$placeTop width=$placeW height=$placeH native=${imgWpx}x${imgHpx}px scale=" + [Math]::Round($scale, 3))
