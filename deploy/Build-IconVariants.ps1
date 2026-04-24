#Requires -Version 5.1
<#
.SYNOPSIS
    Renders chamfered "machined steel plate" icon candidates.

.DESCRIPTION
    Visual system per user feedback (saved in memory):
      - Chamfered (45°) corners — industrial / steel-plate look
      - Flat solid tile, no gloss / gradient / glow
      - White glyph on coloured tile — high contrast at 16px taskbar size
      - Hero glyph: NMT wordmark
      - Variant glyphs: iso wireframe cube (CAD), padlock (PDM vault),
                        drafting triangle + ruler (2D drafting)

    All glyphs use a shared stroke weight so the family is visually
    consistent.

    Output: NMTMacroGuide-v3-<color>-<glyph>-<size>.png
#>

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# ── Palette ─────────────────────────────────────────────────────────────
# All "tile" colours saturated enough that white glyph reads at 16px.
$Tiles = @{
    'blue'   = [System.Drawing.Color]::FromArgb(255,  30,  64, 175)   # #1e40af  corporate dark blue
    'steel'  = [System.Drawing.Color]::FromArgb(255,  71,  85, 105)   # #475569  steel grey
    'yellow' = [System.Drawing.Color]::FromArgb(255, 217, 119,   6)   # #d97706  industrial / safety
}
# Subtle dark inner border softens the chamfer edge against light backgrounds.
$Border = [System.Drawing.Color]::FromArgb(90, 0, 0, 0)
$White  = [System.Drawing.Color]::White

# Glyph occupies this fraction of the tile (slightly less for the wordmark
# since 3 letters need more horizontal room).
$GlyphScale     = 0.58
$WordmarkScale  = 0.74

# ── Canvas + tile ───────────────────────────────────────────────────────
function New-Canvas {
    param([int]$Size)
    $bmp = New-Object System.Drawing.Bitmap $Size, $Size, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g   = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode      = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.TextRenderingHint  = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
    $g.InterpolationMode  = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
    $g.Clear([System.Drawing.Color]::Transparent)
    return [pscustomobject]@{ Bitmap = $bmp; Graphics = $g }
}

# Octagonal "chamfered square" — like a plate with the corners machined off
# at 45°. Chamfer depth scales with tile size so the angle stays 45°.
function Add-ChamferedTile {
    param(
        [System.Drawing.Graphics]$g,
        [int]$Size,
        [System.Drawing.Color]$FillColor
    )
    $pad   = [Math]::Max(1, [int]($Size * 0.012))
    $L = [double]$pad
    $T = [double]$pad
    $R = [double]($Size - $pad)
    $B = [double]($Size - $pad)
    $cham = ($R - $L) * 0.13   # ≈ 13% — reads as a clear chamfer at 16px without dominating

    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $path.AddPolygon([System.Drawing.PointF[]]@(
        (New-Object System.Drawing.PointF ($L + $cham), $T),
        (New-Object System.Drawing.PointF ($R - $cham), $T),
        (New-Object System.Drawing.PointF $R, ($T + $cham)),
        (New-Object System.Drawing.PointF $R, ($B - $cham)),
        (New-Object System.Drawing.PointF ($R - $cham), $B),
        (New-Object System.Drawing.PointF ($L + $cham), $B),
        (New-Object System.Drawing.PointF $L, ($B - $cham)),
        (New-Object System.Drawing.PointF $L, ($T + $cham))
    ))

    $brush = New-Object System.Drawing.SolidBrush $FillColor
    $g.FillPath($brush, $path)
    $brush.Dispose()

    $pen = New-Object System.Drawing.Pen $Border, ([Math]::Max(1.0, $Size / 256.0))
    $g.DrawPath($pen, $path)
    $pen.Dispose()
    return $path
}

# Centre + scale a path to fit ($Size * $Scale), centred on tile.
function Fit-Path {
    param([System.Drawing.Drawing2D.GraphicsPath]$Path, [int]$Size, [double]$Scale = $GlyphScale)
    $b = $Path.GetBounds()
    if ($b.Width -lt 0.1 -or $b.Height -lt 0.1) { return }
    $target = $Size * $Scale
    $s = [Math]::Min($target / $b.Width, $target / $b.Height)
    $mat = New-Object System.Drawing.Drawing2D.Matrix
    $mat.Translate(($Size - $b.Width * $s) / 2, ($Size - $b.Height * $s) / 2)
    $mat.Scale($s, $s)
    $mat.Translate(-$b.X, -$b.Y)
    $Path.Transform($mat)
}

# Convert a stroked path into a filled path so it fills crisply at any size
# (avoids GDI+ subpixel pen-width quirks at tiny tile sizes).
function ConvertTo-FilledStroke {
    param([System.Drawing.Drawing2D.GraphicsPath]$Path, [double]$StrokeWidth)
    $pen = New-Object System.Drawing.Pen ([System.Drawing.Color]::Black), $StrokeWidth
    $pen.LineJoin = [System.Drawing.Drawing2D.LineJoin]::Miter
    $pen.MiterLimit = 4
    $pen.StartCap = [System.Drawing.Drawing2D.LineCap]::Flat
    $pen.EndCap   = [System.Drawing.Drawing2D.LineCap]::Flat
    $filled = $Path.Clone()
    $filled.Widen($pen)
    $pen.Dispose()
    return $filled
}

# ── Glyphs ──────────────────────────────────────────────────────────────

# NMT wordmark — bold sans, fits ~74% of tile width.
function Get-NMTPath {
    param([int]$Size)
    $candidates = @('Segoe UI Black', 'Arial Black', 'Segoe UI', 'Arial')
    $font = $null
    foreach ($fam in $candidates) {
        try {
            $font = New-Object System.Drawing.Font $fam, ($Size * 0.4), ([System.Drawing.FontStyle]::Bold), ([System.Drawing.GraphicsUnit]::Pixel)
            break
        } catch { continue }
    }
    $sf = New-Object System.Drawing.StringFormat
    $sf.Alignment     = [System.Drawing.StringAlignment]::Near
    $sf.LineAlignment = [System.Drawing.StringAlignment]::Near
    $layout = New-Object System.Drawing.RectangleF 0, 0, ($Size * 4), ($Size * 4)
    $p = New-Object System.Drawing.Drawing2D.GraphicsPath
    $p.AddString('NMT', $font.FontFamily, [int][System.Drawing.FontStyle]::Bold, ($Size * 0.4), $layout, $sf)
    $font.Dispose()
    $sf.Dispose()
    Fit-Path $p $Size $WordmarkScale
    return $p
}

# Isometric cube wireframe — 6 outer hexagon edges + 3 inner edges from
# the centre vertex (where the three visible faces meet). Stroked then
# Widened to a filled path for consistent rendering.
function Get-CubePath {
    param([int]$Size)
    $cx = $Size * 0.5
    $cy = $Size * 0.5
    $r  = $Size * 0.30
    $hx = $r * 0.866   # cos(30°)
    $hy = $r * 0.5     # sin(30°)

    $top    = New-Object System.Drawing.PointF $cx,        ($cy - $r)
    $tl     = New-Object System.Drawing.PointF ($cx - $hx), ($cy - $hy)
    $tr     = New-Object System.Drawing.PointF ($cx + $hx), ($cy - $hy)
    $bl     = New-Object System.Drawing.PointF ($cx - $hx), ($cy + $hy)
    $br     = New-Object System.Drawing.PointF ($cx + $hx), ($cy + $hy)
    $bottom = New-Object System.Drawing.PointF $cx,        ($cy + $r)
    $center = New-Object System.Drawing.PointF $cx,        $cy

    $lines = New-Object System.Drawing.Drawing2D.GraphicsPath
    # Outer hexagon (6 edges)
    $lines.AddLine($top, $tr);     $lines.StartFigure()
    $lines.AddLine($tr, $br);      $lines.StartFigure()
    $lines.AddLine($br, $bottom);  $lines.StartFigure()
    $lines.AddLine($bottom, $bl);  $lines.StartFigure()
    $lines.AddLine($bl, $tl);      $lines.StartFigure()
    $lines.AddLine($tl, $top);     $lines.StartFigure()
    # Three inner edges from centre vertex
    $lines.AddLine($center, $top);     $lines.StartFigure()
    $lines.AddLine($center, $bl);      $lines.StartFigure()
    $lines.AddLine($center, $br)

    # Stroke width scaled so it looks consistent at all sizes.
    $strokeWidth = $Size * 0.055
    $filled = ConvertTo-FilledStroke $lines $strokeWidth
    $lines.Dispose()
    return $filled
}

# Heavy-duty padlock — solid body with rounded corners, U-shaped shackle
# above it, small keyhole cut out of the body.
function Get-LockPath {
    param([int]$Size)
    $cx = $Size * 0.5

    # Body
    $bodyW = $Size * 0.50
    $bodyH = $Size * 0.36
    $bodyL = $cx - $bodyW / 2
    $bodyT = $Size * 0.50
    $bodyR = $bodyL + $bodyW
    $bodyB = $bodyT + $bodyH
    $cor   = $Size * 0.04
    $diam  = $cor * 2

    $body = New-Object System.Drawing.Drawing2D.GraphicsPath
    $body.AddArc($bodyL,            $bodyT,            $diam, $diam, 180, 90)
    $body.AddArc(($bodyR - $diam),  $bodyT,            $diam, $diam, 270, 90)
    $body.AddArc(($bodyR - $diam), ($bodyB - $diam),   $diam, $diam,   0, 90)
    $body.AddArc($bodyL,           ($bodyB - $diam),   $diam, $diam,  90, 90)
    $body.CloseFigure()

    # Shackle: half-circle on top + two vertical legs down to body top.
    $shackleOuterW = $Size * 0.36
    $shackleStroke = $Size * 0.075
    $sLeft  = $cx - $shackleOuterW / 2
    $sRight = $cx + $shackleOuterW / 2
    $sTop   = $Size * 0.18
    $arcRect = New-Object System.Drawing.RectangleF $sLeft, $sTop, $shackleOuterW, $shackleOuterW

    $shackleLine = New-Object System.Drawing.Drawing2D.GraphicsPath
    $shackleLine.AddArc($arcRect, 180, 180)            # top half-circle, ends at horizontal midline
    $arcMidY = $sTop + $shackleOuterW / 2
    $shackleLine.StartFigure()
    $shackleLine.AddLine($sLeft,  $arcMidY, $sLeft,  $bodyT)   # left leg
    $shackleLine.StartFigure()
    $shackleLine.AddLine($sRight, $arcMidY, $sRight, $bodyT)   # right leg

    $shackleFilled = ConvertTo-FilledStroke $shackleLine $shackleStroke
    $shackleLine.Dispose()

    # Combine body + shackle, then cut a keyhole circle from the body
    # using even-odd fill.
    $combined = New-Object System.Drawing.Drawing2D.GraphicsPath
    $combined.AddPath($body, $false)
    $combined.AddPath($shackleFilled, $false)

    $kr = $Size * 0.045
    $key = New-Object System.Drawing.RectangleF ($cx - $kr), ($bodyT + $bodyH * 0.32), ($kr * 2), ($kr * 2)
    $combined.AddEllipse($key)
    $combined.FillMode = [System.Drawing.Drawing2D.FillMode]::Alternate

    $body.Dispose()
    $shackleFilled.Dispose()
    return $combined
}

# Drafting triangle (45-45-90, outlined) sitting on a horizontal ruler
# with tick marks. Even-odd fill subtracts the inside of the triangle and
# the tick rectangles from the ruler bar.
function Get-DraftingPath {
    param([int]$Size)
    # Triangle vertices — right angle at bottom-left, hypotenuse top-right.
    $tBL = New-Object System.Drawing.PointF ($Size * 0.20), ($Size * 0.72)
    $tT  = New-Object System.Drawing.PointF ($Size * 0.20), ($Size * 0.18)
    $tBR = New-Object System.Drawing.PointF ($Size * 0.74), ($Size * 0.72)
    $tri = New-Object System.Drawing.Drawing2D.GraphicsPath
    $tri.AddPolygon([System.Drawing.PointF[]]@($tBL, $tT, $tBR))
    # Stroke width matches the cube wireframe (0.055) so the family looks
    # cohesive — was 0.045 originally, looked thinner than the others.
    $triFilled = ConvertTo-FilledStroke $tri ($Size * 0.055)
    $tri.Dispose()

    # Ruler: a long thin rectangle at the bottom + tick cutouts.
    $rL = $Size * 0.10
    $rR = $Size * 0.92
    $rT = $Size * 0.78
    $rH = $Size * 0.10
    $ruler = New-Object System.Drawing.Drawing2D.GraphicsPath
    $ruler.AddRectangle((New-Object System.Drawing.RectangleF $rL, $rT, ($rR - $rL), $rH))

    $tickCount = 9
    $tickW = $Size * 0.012
    $tickH = $Size * 0.045   # cut down from the top edge
    for ($i = 1; $i -lt $tickCount; $i++) {
        $tx = $rL + ($rR - $rL) * ($i / [double]$tickCount) - $tickW / 2
        $ruler.AddRectangle((New-Object System.Drawing.RectangleF $tx, $rT, $tickW, $tickH))
    }
    $ruler.FillMode = [System.Drawing.Drawing2D.FillMode]::Alternate

    $combined = New-Object System.Drawing.Drawing2D.GraphicsPath
    $combined.AddPath($triFilled, $false)
    $combined.AddPath($ruler, $false)
    $combined.FillMode = [System.Drawing.Drawing2D.FillMode]::Winding
    $triFilled.Dispose()
    $ruler.Dispose()
    return $combined
}

# ── Composite ───────────────────────────────────────────────────────────
function Render-Variant {
    param(
        [string]$ColorName,
        [string]$GlyphName,
        [int]$Size
    )
    $c = New-Canvas $Size
    $tile = Add-ChamferedTile $c.Graphics $Size $Tiles[$ColorName]

    $glyph = switch ($GlyphName) {
        'nmt'      { Get-NMTPath      $Size }
        'cube'     { Get-CubePath     $Size }
        'lock'     { Get-LockPath     $Size }
        'drafting' { Get-DraftingPath $Size }
    }

    $brush = New-Object System.Drawing.SolidBrush $White
    $c.Graphics.FillPath($brush, $glyph)
    $brush.Dispose()

    $glyph.Dispose()
    $tile.Dispose()
    $c.Graphics.Dispose()
    return $c.Bitmap
}

# ── Render ──────────────────────────────────────────────────────────────
$glyphs = @('nmt', 'cube', 'lock', 'drafting')

Write-Host ""
Write-Host "  Building chamfered-tile variants..." -ForegroundColor Cyan

# All 4 glyphs in primary blue at 256 + 16 (small-size sanity).
foreach ($glyph in $glyphs) {
    foreach ($size in @(256, 16)) {
        $bmp = Render-Variant -ColorName 'blue' -GlyphName $glyph -Size $size
        $out = Join-Path $ScriptDir ("NMTMacroGuide-v3-blue-{0}-{1}.png" -f $glyph, $size)
        $bmp.Save($out, [System.Drawing.Imaging.ImageFormat]::Png)
        $bmp.Dispose()
    }
    Write-Host ("    blue   / {0,-8}  (256 + 16 written)" -f $glyph) -ForegroundColor DarkGray
}

# NMT hero in alternate colours (steel, yellow) at 256 only.
foreach ($color in 'steel','yellow') {
    $bmp = Render-Variant -ColorName $color -GlyphName 'nmt' -Size 256
    $out = Join-Path $ScriptDir ("NMTMacroGuide-v3-{0}-nmt-256.png" -f $color)
    $bmp.Save($out, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
    Write-Host ("    {0,-6} / nmt       (256 written)" -f $color) -ForegroundColor DarkGray
}

Write-Host ""
Write-Host "  Done." -ForegroundColor Green
Write-Host ""
