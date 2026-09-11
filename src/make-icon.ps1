# Рисует иконку приложения (app.ico): синий квадрат с «блёстками».
Add-Type -AssemblyName System.Drawing
$out = Join-Path $PSScriptRoot 'app.ico'

function New-IconBitmap([int]$s) {
    $bmp = New-Object System.Drawing.Bitmap $s, $s, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = 'AntiAlias'
    $g.Clear([System.Drawing.Color]::Transparent)
    $pad = [math]::Max(0.5, $s * 0.04)
    $r = New-Object System.Drawing.RectangleF $pad, $pad, ($s - 2 * $pad), ($s - 2 * $pad)
    $d = $r.Width * 0.46
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $path.AddArc($r.X, $r.Y, $d, $d, 180, 90)
    $path.AddArc($r.Right - $d, $r.Y, $d, $d, 270, 90)
    $path.AddArc($r.Right - $d, $r.Bottom - $d, $d, $d, 0, 90)
    $path.AddArc($r.X, $r.Bottom - $d, $d, $d, 90, 90)
    $path.CloseFigure()
    $brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush $r, ([System.Drawing.Color]::FromArgb(37, 99, 235)), ([System.Drawing.Color]::FromArgb(6, 182, 212)), 45.0
    $g.FillPath($brush, $path)

    $sparkle = {
        param($cx, $cy, $R)
        $p = New-Object System.Drawing.Drawing2D.GraphicsPath
        $k = $R * 0.14
        $tips = @(
            (New-Object System.Drawing.PointF $cx, ($cy - $R)),
            (New-Object System.Drawing.PointF ($cx + $R), $cy),
            (New-Object System.Drawing.PointF $cx, ($cy + $R)),
            (New-Object System.Drawing.PointF ($cx - $R), $cy))
        $ctl = @(
            (New-Object System.Drawing.PointF ($cx + $k), ($cy - $k)),
            (New-Object System.Drawing.PointF ($cx + $k), ($cy + $k)),
            (New-Object System.Drawing.PointF ($cx - $k), ($cy + $k)),
            (New-Object System.Drawing.PointF ($cx - $k), ($cy - $k)))
        for ($i = 0; $i -lt 4; $i++) { $p.AddBezier($tips[$i], $ctl[$i], $ctl[$i], $tips[($i + 1) % 4]) }
        $p.CloseFigure()
        $g.FillPath([System.Drawing.Brushes]::White, $p)
    }
    & $sparkle ($s * 0.43) ($s * 0.56) ($s * 0.33)
    & $sparkle ($s * 0.72) ($s * 0.27) ($s * 0.15)
    if ($s -ge 32) { & $sparkle ($s * 0.75) ($s * 0.76) ($s * 0.09) }
    $g.Dispose()
    return $bmp
}

function Get-DibBytes($bmp) {
    $s = $bmp.Width
    $ms = New-Object IO.MemoryStream
    $w = New-Object IO.BinaryWriter $ms
    $w.Write([int32]40); $w.Write([int32]$s); $w.Write([int32]($s * 2)); $w.Write([int16]1); $w.Write([int16]32)
    $w.Write([int32]0); $w.Write([int32]($s * $s * 4)); $w.Write([int32]0); $w.Write([int32]0); $w.Write([int32]0); $w.Write([int32]0)
    for ($y = $s - 1; $y -ge 0; $y--) {
        for ($x = 0; $x -lt $s; $x++) {
            $c = $bmp.GetPixel($x, $y)
            $w.Write([byte]$c.B); $w.Write([byte]$c.G); $w.Write([byte]$c.R); $w.Write([byte]$c.A)
        }
    }
    $maskRow = [int]([math]::Ceiling($s / 32.0) * 4)
    $w.Write((New-Object byte[] ($maskRow * $s)))
    $w.Flush()
    return , $ms.ToArray()
}

$sizes = 16, 24, 32, 48, 64, 256
$blobs = @()
foreach ($s in $sizes) {
    $bmp = New-IconBitmap $s
    if ($s -ge 256) {
        $ms = New-Object IO.MemoryStream
        $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
        $blobs += , $ms.ToArray()
    } else {
        $blobs += , (Get-DibBytes $bmp)
    }
    $bmp.Dispose()
}

$fs = [IO.File]::Create($out)
$bw = New-Object IO.BinaryWriter $fs
$bw.Write([uint16]0); $bw.Write([uint16]1); $bw.Write([uint16]$sizes.Count)
$offset = 6 + 16 * $sizes.Count
for ($i = 0; $i -lt $sizes.Count; $i++) {
    $b = if ($sizes[$i] -ge 256) { 0 } else { $sizes[$i] }
    $bw.Write([byte]$b); $bw.Write([byte]$b); $bw.Write([byte]0); $bw.Write([byte]0)
    $bw.Write([uint16]1); $bw.Write([uint16]32); $bw.Write([uint32]$blobs[$i].Length); $bw.Write([uint32]$offset)
    $offset += $blobs[$i].Length
}
foreach ($blob in $blobs) { $bw.Write($blob) }
$bw.Close()
"OK: $out"
