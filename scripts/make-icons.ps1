# Generate Slilot brand icons (transparent PNG): slide stack + 4-point sparkle
# Usage: powershell -ExecutionPolicy Bypass -File scripts/make-icons.ps1
Add-Type -AssemblyName System.Drawing
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$outDir = Join-Path (Split-Path -Parent $here) "public\icons"

foreach ($s in @(16, 32, 80, 256)) {
  $bmp = New-Object System.Drawing.Bitmap($s, $s)
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
  $g.Clear([System.Drawing.Color]::Transparent)
  $k = $s / 64.0

  # back slide (dark, offset)
  $p1 = New-Object System.Drawing.Drawing2D.GraphicsPath
  $bx = 7 * $k; $by = 13 * $k; $bw = 37 * $k; $bh = 44 * $k; $bd = 10 * $k
  $p1.AddArc($bx, $by, $bd, $bd, 180, 90)
  $p1.AddArc($bx + $bw - $bd, $by, $bd, $bd, 270, 90)
  $p1.AddArc($bx + $bw - $bd, $by + $bh - $bd, $bd, $bd, 0, 90)
  $p1.AddArc($bx, $by + $bh - $bd, $bd, $bd, 90, 90)
  $p1.CloseFigure()
  $c1 = [System.Drawing.Color]::FromArgb(255, 154, 52, 34)
  $b1 = New-Object System.Drawing.SolidBrush($c1)
  $g.FillPath($b1, $p1)

  # front slide (gradient rounded rect)
  $p2 = New-Object System.Drawing.Drawing2D.GraphicsPath
  $fx = 17 * $k; $fy = 7 * $k; $fw = 40 * $k; $fh = 50 * $k; $fd = 12 * $k
  $p2.AddArc($fx, $fy, $fd, $fd, 180, 90)
  $p2.AddArc($fx + $fw - $fd, $fy, $fd, $fd, 270, 90)
  $p2.AddArc($fx + $fw - $fd, $fy + $fh - $fd, $fd, $fd, 0, 90)
  $p2.AddArc($fx, $fy + $fh - $fd, $fd, $fd, 90, 90)
  $p2.CloseFigure()
  $rect2 = New-Object System.Drawing.RectangleF([single]$fx, [single]$fy, [single]$fw, [single]$fh)
  $colA = [System.Drawing.Color]::FromArgb(255, 232, 115, 77)
  $colB = [System.Drawing.Color]::FromArgb(255, 201, 63, 34)
  $grad = New-Object System.Drawing.Drawing2D.LinearGradientBrush($rect2, $colA, $colB, 90)
  $g.FillPath($grad, $p2)

  # 4-point sparkle (white), center (37,29) r=14
  $cx = 37 * $k; $cy = 29 * $k; $r = 14 * $k
  $a = 0.107 * $r; $b = 0.464 * $r
  $p3 = New-Object System.Drawing.Drawing2D.GraphicsPath
  $p3.AddBezier($cx, $cy - $r, $cx + $a, $cy - $b, $cx + $b, $cy - $a, $cx + $r, $cy)
  $p3.AddBezier($cx + $r, $cy, $cx + $b, $cy + $a, $cx + $a, $cy + $b, $cx, $cy + $r)
  $p3.AddBezier($cx, $cy + $r, $cx - $a, $cy + $b, $cx - $b, $cy + $a, $cx - $r, $cy)
  $p3.AddBezier($cx - $r, $cy, $cx - $b, $cy - $a, $cx - $a, $cy - $b, $cx, $cy - $r)
  $p3.CloseFigure()
  $c3 = [System.Drawing.Color]::FromArgb(255, 255, 247, 243)
  $b3 = New-Object System.Drawing.SolidBrush($c3)
  $g.FillPath($b3, $p3)

  # two dots (text placeholders)
  $c4 = [System.Drawing.Color]::FromArgb(255, 255, 217, 204)
  $b4 = New-Object System.Drawing.SolidBrush($c4)
  $dr = 2.4 * $k
  $g.FillEllipse($b4, [single](25 * $k - $dr), [single](49 * $k - $dr), [single](2 * $dr), [single](2 * $dr))
  $g.FillEllipse($b4, [single](32 * $k - $dr), [single](49 * $k - $dr), [single](2 * $dr), [single](2 * $dr))

  $outPath = Join-Path $outDir ("icon-" + $s + ".png")
  $bmp.Save($outPath, [System.Drawing.Imaging.ImageFormat]::Png)
  $g.Dispose()
  $bmp.Dispose()
  Write-Output ("saved " + $outPath)
}
