param([Parameter(Mandatory=$true)][string]$Path,[string]$OutDir="_slides")
$ErrorActionPreference = "Stop"
$full = (Resolve-Path $Path).Path
$dir = Join-Path (Split-Path $full) $OutDir
if (Test-Path $dir) { Remove-Item $dir -Recurse -Force }
New-Item -ItemType Directory -Path $dir | Out-Null
$pp = New-Object -ComObject PowerPoint.Application
try {
    $pres = $pp.Presentations.Open($full, $true, $false, $false)  # ReadOnly, no window
    $i = 1
    foreach ($sl in $pres.Slides) {
        $out = Join-Path $dir ("slide{0:D2}.png" -f $i)
        $sl.Export($out, "PNG", 1600, 900)
        $i++
    }
    $pres.Close()
    "exported $($i-1) slides -> $dir"
} finally {
    $pp.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pp) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
