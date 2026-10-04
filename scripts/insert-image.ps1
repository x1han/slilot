# 向当前演示文稿的指定页插入图片（精确页 + 坐标，供 Slilot add_image 使用）
# 用法: powershell -NoProfile -ExecutionPolicy Bypass -File scripts/insert-image.ps1 -ImagePath <path> -SlideNumber <1-based> -Left <pt> -Top <pt> -Width <pt> -Height <pt>
# 注意: PowerPoint 单实例。COM 拉起的实例不注册到 ROT，GetActiveObject 失败时
#       必须回退 New-Object（CoCreateInstance 会附着到已运行进程）
param([string]$ImagePath, [int]$SlideNumber, [single]$Left, [single]$Top, [single]$Width, [single]$Height)
$ErrorActionPreference = "Stop"

if (-not (Test-Path $ImagePath)) { throw ("图片文件不存在: " + $ImagePath) }

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
$shape = $slide.Shapes.AddPicture($ImagePath, 0, -1, $Left, $Top, $Width, $Height)
Write-Output ("inserted: " + $shape.Name + " on slide " + $SlideNumber)
