param(
    [string]$Pptx = 'D:\AI-agent\codexworksection\american-character\build\American-Character-Forrest-Gump.pptx',
    [string]$OutDir = 'D:\AI-agent\codexworksection\american-character\build\slides'
)

$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.DisplayAlerts = 1
try {
    $pres = $ppt.Presentations.Open($Pptx, $true, $false, $false)
    Write-Output ("slides=" + $pres.Slides.Count)
    for ($i = 1; $i -le $pres.Slides.Count; $i++) {
        Write-Output ("slide " + $i +
            ": shapes=" + $pres.Slides.Item($i).Shapes.Count +
            " animation_effects=" + $pres.Slides.Item($i).TimeLine.MainSequence.Count)
    }
    $pres.Export($OutDir, 'PNG', 1920, 1080)
    $pdf = [System.IO.Path]::ChangeExtension($Pptx, '.pdf')
    $pres.SaveAs($pdf, 32)
    Write-Output ("exported -> " + $OutDir)
    Write-Output ("pdf -> " + $pdf)
    $pres.Close()
}
finally {
    $ppt.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
}
