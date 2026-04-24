$c = 'C:\AllMacros\MacroGuide.html'
$y = 'Y:\Solidworks\Macros\Macro Data PDM\MacroGuide\MacroGuide.html'
Get-Item $c, $y | Select-Object FullName, Length, LastWriteTime | Format-List
Write-Host '---diff (C: vs Y:)---'
$diff = Compare-Object (Get-Content $c) (Get-Content $y)
if ($diff) {
    Write-Host ("Differences: " + $diff.Count + " lines")
    $diff | Select-Object -First 30 | Format-Table
} else {
    Write-Host 'IDENTICAL'
}
Write-Host '---server PS diff---'
$cs = 'C:\AllMacros\deploy\macro-guide-server.ps1'
$ys = 'Y:\Solidworks\Macros\Macro Data PDM\MacroGuide\macro-guide-server.ps1'
$sd = Compare-Object (Get-Content $cs) (Get-Content $ys)
if ($sd) { Write-Host ("server diff lines: " + $sd.Count); $sd | Select-Object -First 10 | Format-Table }
else { Write-Host 'server PS IDENTICAL' }
Write-Host '---running server---'
Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" |
    Where-Object { $_.CommandLine -like '*macro-guide-server*' } |
    Select-Object ProcessId, CreationDate | Format-List
