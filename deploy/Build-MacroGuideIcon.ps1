#Requires -Version 5.1
<#
.SYNOPSIS
    Builds the NMT Macro Guide .ico file from scratch.

.DESCRIPTION
    Generates a multi-resolution Windows icon (16/32/48/64/128/256) in the
    style selected during icon review:
      - Rounded-square tile (≈ 22% corner radius — modern Win11 / iOS default)
      - Flat corporate dark blue fill (#1e40af), no gloss or gradient
      - Subtle dark inner border softens the edge
      - Bold white "NMT" wordmark centred in the plate

    Each size is rendered separately (not downscaled from 256) so 16/32 px
    stay legible on dark-mode taskbars. All sizes are PNG-encoded inside
    the .ico container (supported on Windows Vista+).

.NOTES
    Run on demand:
      powershell -ExecutionPolicy Bypass -File Build-MacroGuideIcon.ps1
    Output: NMTMacroGuide.ico in the same folder as this script.
    The deploy script picks it up from there.
#>

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

# ── Output ──────────────────────────────────────────────────────────────
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$OutPath   = Join-Path $ScriptDir 'NMTMacroGuide.ico'

# ── Palette ─────────────────────────────────────────────────────────────
$TileColor  = [System.Drawing.Color]::FromArgb(255,  30,  64, 175)  # #1e40af corporate dark blue
$BorderEdge = [System.Drawing.Color]::FromArgb( 90,   0,   0,   0)  # 35% black inner edge
$Glyph      = [System.Drawing.Color]::White

# ── Render one size into a Bitmap ───────────────────────────────────────
function New-IconBitmap {
    param([int]$Size)

    $bmp = New-Object System.Drawing.Bitmap $Size, $Size, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g   = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode      = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.TextRenderingHint  = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
    $g.InterpolationMode  = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
    $g.Clear([System.Drawing.Color]::Transparent)

    # ── Rounded-square tile ─────────────────────────────────────────────
    # 22% corner radius matches modern app-icon convention (iOS / Win11)
    # and reads as a soft "tile" at every size.
    $pad    = [Math]::Max(1, [int]($Size * 0.012))
    $rectF  = New-Object System.Drawing.RectangleF $pad, $pad, ($Size - 2*$pad), ($Size - 2*$pad)
    $radius = $rectF.Width * 0.22
    $diam   = $radius * 2
    $tile   = New-Object System.Drawing.Drawing2D.GraphicsPath
    $tile.AddArc($rectF.X,             $rectF.Y,              $diam, $diam, 180, 90)
    $tile.AddArc($rectF.Right - $diam, $rectF.Y,              $diam, $diam, 270, 90)
    $tile.AddArc($rectF.Right - $diam, $rectF.Bottom - $diam, $diam, $diam,   0, 90)
    $tile.AddArc($rectF.X,             $rectF.Bottom - $diam, $diam, $diam,  90, 90)
    $tile.CloseFigure()

    $fill = New-Object System.Drawing.SolidBrush $TileColor
    $g.FillPath($fill, $tile)
    $fill.Dispose()

    # Skip the 1-px border at the tiniest size — it just muddies the edge.
    if ($Size -ge 24) {
        $pen = New-Object System.Drawing.Pen $BorderEdge, ([Math]::Max(1.0, $Size / 256.0))
        $g.DrawPath($pen, $tile)
        $pen.Dispose()
    }

    # ── "NMT" wordmark ──────────────────────────────────────────────────
    # Bold sans that ships on Windows. Segoe UI Black is ideal but not
    # everywhere — fall back through a chain.
    $candidates = @('Segoe UI Black', 'Arial Black', 'Segoe UI', 'Arial')
    $font = $null
    foreach ($fam in $candidates) {
        try {
            $font = New-Object System.Drawing.Font $fam, ($Size * 0.4), ([System.Drawing.FontStyle]::Bold), ([System.Drawing.GraphicsUnit]::Pixel)
            break
        } catch { continue }
    }
    if ($null -eq $font) {
        throw "Could not load any of: $($candidates -join ', ')"
    }

    $sf = New-Object System.Drawing.StringFormat
    $sf.Alignment     = [System.Drawing.StringAlignment]::Near
    $sf.LineAlignment = [System.Drawing.StringAlignment]::Near
    # Large layout rect — we fit precisely by measuring the glyph path below.
    $layoutRect = New-Object System.Drawing.RectangleF 0, 0, ($Size * 4), ($Size * 4)

    $textPath = New-Object System.Drawing.Drawing2D.GraphicsPath
    $textPath.AddString('NMT', $font.FontFamily, [int][System.Drawing.FontStyle]::Bold, ($Size * 0.4), $layoutRect, $sf)
    $font.Dispose()
    $sf.Dispose()

    # Fit the glyph to ~74% of the tile width, centred on the tile. Uniform
    # scale so the letterforms stay proportional; recentres on the glyph's
    # actual bounds (font metrics leave more top padding than we want).
    $b = $textPath.GetBounds()
    if ($b.Width -gt 0.1 -and $b.Height -gt 0.1) {
        $target = $Size * 0.74
        $s = [Math]::Min($target / $b.Width, $target / $b.Height)
        $mat = New-Object System.Drawing.Drawing2D.Matrix
        $mat.Translate(($Size - $b.Width * $s) / 2, ($Size - $b.Height * $s) / 2)
        $mat.Scale($s, $s)
        $mat.Translate(-$b.X, -$b.Y)
        $textPath.Transform($mat)
    }

    $glyphBrush = New-Object System.Drawing.SolidBrush $Glyph
    $g.FillPath($glyphBrush, $textPath)
    $glyphBrush.Dispose()

    # Cleanup
    $textPath.Dispose()
    $tile.Dispose()
    $g.Dispose()
    return $bmp
}

# ── Encode a Bitmap as PNG bytes ────────────────────────────────────────
function ConvertTo-PngBytes {
    param([System.Drawing.Bitmap]$Bitmap)
    $ms = New-Object System.IO.MemoryStream
    $Bitmap.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
    $bytes = $ms.ToArray()
    $ms.Dispose()
    return ,$bytes
}

# ── Pack PNG byte arrays into a multi-resolution .ico file ──────────────
# ICO format: ICONDIR (6 bytes) + ICONDIRENTRY[] (16 bytes each) + image
# payloads. PNG-in-ICO is the Vista+ convention and what every modern
# Windows shell handles — simpler than the BMP-for-small / PNG-for-256
# mix that older Windows icons used.
function Save-Ico {
    param(
        [string]$Path,
        [hashtable]$ImagesBySize
    )
    $sizes = $ImagesBySize.Keys | Sort-Object -Descending
    $count = $sizes.Count

    $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write)
    $bw = New-Object System.IO.BinaryWriter $fs
    try {
        $bw.Write([uint16]0)         # Reserved
        $bw.Write([uint16]1)         # Type: 1 = icon
        $bw.Write([uint16]$count)    # Number of images

        $offset = 6 + (16 * $count)
        $entries = @()
        foreach ($s in $sizes) {
            $bytes = $ImagesBySize[$s]
            $entries += [pscustomobject]@{ Size = $s; Bytes = $bytes; Offset = $offset }
            $offset += $bytes.Length
        }

        foreach ($e in $entries) {
            $w = if ($e.Size -ge 256) { 0 } else { $e.Size }   # 0 means 256 in ICO spec
            $bw.Write([byte]$w)
            $bw.Write([byte]$w)
            $bw.Write([byte]0)        # Colors in palette (0 for >=256)
            $bw.Write([byte]0)        # Reserved
            $bw.Write([uint16]1)      # Colour planes
            $bw.Write([uint16]32)     # Bits per pixel
            $bw.Write([uint32]$e.Bytes.Length)
            $bw.Write([uint32]$e.Offset)
        }

        foreach ($e in $entries) {
            $bw.Write($e.Bytes)
        }
    } finally {
        $bw.Dispose()
        $fs.Dispose()
    }
}

# ── Build it ────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  Building NMT Macro Guide icon (rounded NMT wordmark)..." -ForegroundColor Cyan
$sizes = @(256, 128, 64, 48, 32, 16)
$images = @{}
foreach ($s in $sizes) {
    $bmp = New-IconBitmap -Size $s
    $images[$s] = ConvertTo-PngBytes -Bitmap $bmp
    $bmp.Dispose()
    Write-Host ("    rendered {0,3}px  ({1,5} bytes)" -f $s, $images[$s].Length) -ForegroundColor DarkGray
}

Save-Ico -Path $OutPath -ImagesBySize $images

$finalSize = (Get-Item $OutPath).Length
Write-Host ""
Write-Host ("  Wrote: {0}" -f $OutPath) -ForegroundColor Green
Write-Host ("  Total: {0:N0} bytes" -f $finalSize) -ForegroundColor Green
Write-Host ""
