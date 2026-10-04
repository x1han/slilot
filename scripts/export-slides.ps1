# 将 pptx 渲染为逐页 PNG（供 Slilot 视觉审查）
# 用法: powershell -NoProfile -ExecutionPolicy Bypass -File scripts/export-slides.ps1 -PptxPath <path> -OutDir <dir>
# 输出: OutDir\slide-<n>.png + manifest.txt（每行一个 png 路径）
# 注意: PowerPoint COM 是单实例。若用户正在操作 PowerPoint（有模态对话框等），
#       Presentations.Open 可能抛伪错误（如 Filename cannot exceed 255），此处带重试。
param([string]$PptxPath, [string]$OutDir)
$ErrorActionPreference = "Stop"

if (-not (Test-Path $PptxPath)) { throw ("输入文件不存在: " + $PptxPath) }
Write-Output ("input: " + $PptxPath + " (len=" + $PptxPath.Length + ", exists=" + (Test-Path $PptxPath) + ")")

$pp = New-Object -ComObject PowerPoint.Application

$pres = $null
$lastErr = ""
for ($try = 1; $try -le 3 -and ($null -eq $pres); $try++) {
  try {
    # WithWindow 必须为 true：msoFalse 时会抛"Filename cannot exceed 255"的伪错误
    $pres = $pp.Presentations.Open($PptxPath, -1, 0, -1)
  } catch {
    $lastErr = $_.Exception.Message
    Write-Output ("open attempt " + $try + " failed: " + $lastErr)
    Start-Sleep -Seconds (2 * $try)
  }
}
if ($null -eq $pres) { throw ("Open failed after 3 attempts: " + $lastErr) }

try {
  $count = $pres.Slides.Count
  $files = @()
  for ($i = 1; $i -le $count; $i++) {
    $png = Join-Path $OutDir ("slide-" + $i + ".png")
    $pres.Slides.Item($i).Export($png, "PNG", 1280, 720)
    $files += $png
  }
  $files | Set-Content -Path (Join-Path $OutDir "manifest.txt") -Encoding ASCII
  Write-Output ("exported " + $count + " slides")
} finally {
  $pres.Close()
  $pp.Quit()
}
