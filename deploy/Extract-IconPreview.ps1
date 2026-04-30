# Re-render preview PNGs directly from the generator's drawing function.
# The .NET Icon class can't decode PNG-format ICO entries (a known
# limitation), so we bypass it and call the bitmap renderer ourselves.
# Windows Explorer renders the .ico fine — this is just for human inspection.

. 'C:\AllMacros\deploy\Build-MacroGuideIcon.ps1' *>$null  # only need the function

# Above dot-sources the generator, which also rebuilds the .ico. That's fine —
# but it suppresses output. If the function isn't in scope, fail loudly.
if (-not (Get-Command New-IconBitmap -ErrorAction SilentlyContinue)) {
    throw "New-IconBitmap not in scope after dot-sourcing"
}

foreach ($s in 256, 128, 64, 48, 32, 16) {
    $bmp = New-IconBitmap -Size $s
    $out = "C:\AllMacros\deploy\NMTMacroGuide-preview-$s.png"
    $bmp.Save($out, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
    Write-Host "  wrote $out"
}
