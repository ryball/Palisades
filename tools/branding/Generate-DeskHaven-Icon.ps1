Add-Type -AssemblyName System.Drawing

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$resourcesDir = Join-Path $repoRoot 'Palisades.Application\Ressources'
$previewPath = Join-Path $resourcesDir 'icon-preview.png'
$iconPath = Join-Path $resourcesDir 'icon.ico'

$size = 256
$bitmap = New-Object System.Drawing.Bitmap $size, $size
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)
$graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
$graphics.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
$graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
$graphics.Clear([System.Drawing.Color]::Transparent)

function New-RoundedRectPath([float]$x, [float]$y, [float]$w, [float]$h, [float]$r) {
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $d = $r * 2
    $path.AddArc($x, $y, $d, $d, 180, 90)
    $path.AddArc($x + $w - $d, $y, $d, $d, 270, 90)
    $path.AddArc($x + $w - $d, $y + $h - $d, $d, $d, 0, 90)
    $path.AddArc($x, $y + $h - $d, $d, $d, 90, 90)
    $path.CloseFigure()
    return $path
}

$outerPath = New-RoundedRectPath 12 12 232 232 44
$backgroundBrush = New-Object System.Drawing.Drawing2D.LinearGradientBrush (
    [System.Drawing.Point]::new(0,0),
    [System.Drawing.Point]::new(256,256),
    [System.Drawing.Color]::FromArgb(255, 32, 47, 110),
    [System.Drawing.Color]::FromArgb(255, 42, 144, 214)
)
$graphics.FillPath($backgroundBrush, $outerPath)

$glowBrush = New-Object System.Drawing.Drawing2D.PathGradientBrush($outerPath)
$glowBrush.CenterColor = [System.Drawing.Color]::FromArgb(80, 255, 255, 255)
$glowBrush.SurroundColors = [System.Drawing.Color[]]@([System.Drawing.Color]::FromArgb(0, 255, 255, 255))
$graphics.FillPath($glowBrush, $outerPath)

$borderPen = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(110, 255, 255, 255)), 4
$graphics.DrawPath($borderPen, $outerPath)

$tabStripBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(85, 13, 22, 58))
$graphics.FillRectangle($tabStripBrush, 34, 40, 188, 28)

$tabColors = @(
    [System.Drawing.Color]::FromArgb(255, 76, 209, 186),
    [System.Drawing.Color]::FromArgb(255, 120, 126, 255),
    [System.Drawing.Color]::FromArgb(255, 255, 191, 93)
)

for ($i = 0; $i -lt 3; $i++) {
    $tabX = 44 + ($i * 58)
    $tabPath = New-RoundedRectPath $tabX 28 46 34 12
    $tabBrush = New-Object System.Drawing.SolidBrush ($tabColors[$i])
    $graphics.FillPath($tabBrush, $tabPath)
}

$panelLayouts = @(
    @{ X = 34; Y = 78; W = 84; H = 120; Color = [System.Drawing.Color]::FromArgb(230, 50, 193, 176) },
    @{ X = 126; Y = 78; W = 96; H = 54; Color = [System.Drawing.Color]::FromArgb(230, 119, 122, 255) },
    @{ X = 126; Y = 144; W = 96; H = 54; Color = [System.Drawing.Color]::FromArgb(230, 255, 187, 84) }
)

foreach ($panel in $panelLayouts) {
    $panelPath = New-RoundedRectPath $panel.X $panel.Y $panel.W $panel.H 16
    $panelBrush = New-Object System.Drawing.SolidBrush $panel.Color
    $graphics.FillPath($panelBrush, $panelPath)
    $graphics.DrawPath((New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(70, 255, 255, 255), 2)), $panelPath)

    $miniBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(210, 255, 255, 255))
    $miniSize = 12
    for ($row = 0; $row -lt 2; $row++) {
        for ($col = 0; $col -lt 2; $col++) {
            $mx = $panel.X + 16 + ($col * 20)
            $my = $panel.Y + 14 + ($row * 20)
            if ($mx + $miniSize -lt ($panel.X + $panel.W - 12) -and $my + $miniSize -lt ($panel.Y + $panel.H - 8)) {
                $graphics.FillEllipse($miniBrush, $mx, $my, $miniSize, $miniSize)
            }
        }
    }
}

$havenPen = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(210, 255, 255, 255), 8)
$graphics.DrawArc($havenPen, 56, 150, 144, 70, 180, 180)
$graphics.DrawLine($havenPen, 60, 186, 128, 118)
$graphics.DrawLine($havenPen, 128, 118, 196, 186)

$bitmap.Save($previewPath, [System.Drawing.Imaging.ImageFormat]::Png)

$iconHandle = $bitmap.GetHicon()
try {
    $icon = [System.Drawing.Icon]::FromHandle($iconHandle)
    $fileStream = [System.IO.File]::Create($iconPath)
    try {
        $icon.Save($fileStream)
    }
    finally {
        $fileStream.Dispose()
        $icon.Dispose()
    }
}
finally {
    $signature = @'
using System;
using System.Runtime.InteropServices;
public static class NativeMethods {
    [DllImport("user32.dll", CharSet=CharSet.Auto)]
    public static extern bool DestroyIcon(IntPtr handle);
}
'@
    Add-Type -TypeDefinition $signature -ErrorAction SilentlyContinue | Out-Null
    [NativeMethods]::DestroyIcon($iconHandle) | Out-Null
    $graphics.Dispose()
    $bitmap.Dispose()
}

Write-Output "Generated icon preview: $previewPath"
Write-Output "Generated application icon: $iconPath"
