param(
    [Parameter(Mandatory=$true)][string]$InputPath,
    [Parameter(Mandatory=$true)][string]$OutputDirectory,
    [ValidateRange(72,1200)][int]$Dpi = 300
)
$ErrorActionPreference = 'Stop'
$inputFile = (Resolve-Path -LiteralPath $InputPath).Path
if ([IO.Path]::GetExtension($inputFile) -ine '.vsdx') { throw 'Only .vsdx inputs are supported.' }
$out = [IO.Path]::GetFullPath($OutputDirectory)
if (Test-Path -LiteralPath $out) { throw 'OutputDirectory must be new; existing files are never overwritten.' }
$sourceHash = (Get-FileHash -LiteralPath $inputFile -Algorithm SHA256).Hash
[void](New-Item -ItemType Directory -Path $out)
$app=$null; $doc=$null; $settings=$null; $savedResolution=$false
try {
    $app = New-Object -ComObject Visio.InvisibleApp
    $app.EventsEnabled = $false
    $app.AlertResponse = 7
    $settings = $app.Settings
    [int]$oldResolution=0; [double]$oldWidth=0; [double]$oldHeight=0; [int]$oldUnits=0
    $settings.GetRasterExportResolution([ref]$oldResolution,[ref]$oldWidth,[ref]$oldHeight,[ref]$oldUnits)
    $savedResolution=$true
    $settings.SetRasterExportResolution(3, $Dpi, $Dpi, 0)
    # Read-only, unlisted, hidden, macros disabled, no workspace and no data refresh.
    $doc = $app.Documents.OpenEx($inputFile, (2+8+64+128+256+1024))
    $pages = @()
    for ($i=1; $i -le $doc.Pages.Count; $i++) {
        $page = $doc.Pages.Item($i)
        $png = Join-Path $out ('page-{0:D3}_preview.png' -f $i)
        $page.Export($png)
        if (!(Test-Path -LiteralPath $png) -or (Get-Item -LiteralPath $png).Length -eq 0) {
            throw "Empty export: $png"
        }
        $pages += [ordered]@{page=$i; name=$page.NameU; top_level_shapes=$page.Shapes.Count;
            png=$png; sha256=(Get-FileHash -LiteralPath $png -Algorithm SHA256).Hash}
    }
    if (!$pages.Count) { throw 'Visio returned no pages.' }
    if ((Get-FileHash -LiteralPath $inputFile -Algorithm SHA256).Hash -ne $sourceHash) {
        throw 'Input changed during rendering; discard this verification.'
    }
    [ordered]@{source=$inputFile; source_sha256=$sourceHash; renderer='Microsoft Visio';
        visio_version=$app.Version; dpi=$Dpi; opened=$true; pages=$pages;
        visual_review='Required separately; opening/export alone does not prove fidelity.'
    } | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $out 'render.json') -Encoding utf8
    Write-Output (Join-Path $out 'render.json')
} finally {
    try { if ($null -ne $doc) { $doc.Close() } }
    finally {
        try {
            if ($savedResolution) {
                $settings.SetRasterExportResolution($oldResolution,$oldWidth,$oldHeight,$oldUnits)
            }
        } finally {
            if ($null -ne $app) { $app.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) }
        }
    }
}
