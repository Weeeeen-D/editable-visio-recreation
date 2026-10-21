[CmdletBinding()]
param(
    [string]$OutputDirectory = (Join-Path $PSScriptRoot '..\examples')
)
$ErrorActionPreference = 'Stop'
$out = [IO.Path]::GetFullPath($OutputDirectory)
[void][IO.Directory]::CreateDirectory($out)
$vsdx = Join-Path $out 'demo.vsdx'
$renderDir = Join-Path $out 'render'
$reportPath = Join-Path $out 'interaction-check.json'
if ((Test-Path -LiteralPath $vsdx) -or (Test-Path -LiteralPath $renderDir) -or (Test-Path -LiteralPath $reportPath)) {
    throw 'Use a new output directory. Existing demo, render and report files are never overwritten.'
}
function mm([double]$value) { return $value / 25.4 }
function Set-Cell($shape, [string]$name, [string]$formula) { $shape.CellsU($name).FormulaU = $formula }
function Set-TextStyle($shape, [int]$fontId) {
    Set-Cell $shape 'Char.Font' ([string]$fontId)
    Set-Cell $shape 'Char.Size' '6 pt'
    Set-Cell $shape 'Char.Color' 'RGB(28,44,61)'
    Set-Cell $shape 'Para.HorzAlign' '1'
    Set-Cell $shape 'VerticalAlign' '1'
    foreach ($margin in @('LeftMargin','RightMargin','TopMargin','BottomMargin')) { Set-Cell $shape $margin '0.6 mm' }
}
function Add-Box($page, [string]$name, [string]$text, [double]$left, [double]$bottom, [double]$right, [double]$top, [int]$fontId, [bool]$labelOnly=$false) {
    $shape = $page.DrawRectangle((mm $left),(mm $bottom),(mm $right),(mm $top))
    $shape.NameU = $name
    $shape.Text = $text
    Set-TextStyle $shape $fontId
    if ($labelOnly) {
        Set-Cell $shape 'LinePattern' '0'; Set-Cell $shape 'FillPattern' '0'
    } else {
        Set-Cell $shape 'LineColor' 'RGB(69,101,129)'
        Set-Cell $shape 'LineWeight' '0.6 pt'
        Set-Cell $shape 'FillForegnd' 'RGB(235,243,249)'
        Set-Cell $shape 'FillPattern' '1'
    }
    return $shape
}
function Add-Connector($app, $page, [string]$name, $source, $destination) {
    $line = $page.Drop($app.ConnectorToolDataObject,0,0)
    $line.NameU = $name
    $line.CellsU('BeginX').GlueTo($source.CellsU('PinX'))
    $line.CellsU('EndX').GlueTo($destination.CellsU('PinX'))
    Set-Cell $line 'LineColor' 'RGB(69,101,129)'
    Set-Cell $line 'LineWeight' '0.6 pt'
    Set-Cell $line 'EndArrow' '13'
    return $line
}
function Get-Bounds($shape) {
    [double]$left=0; [double]$bottom=0; [double]$right=0; [double]$top=0
    # visBBoxExtents (4) checks visible geometry, rather than Width alone.
    $shape.BoundingBox(4,[ref]$left,[ref]$bottom,[ref]$right,[ref]$top)
    return [ordered]@{left_in=$left;bottom_in=$bottom;right_in=$right;top_in=$top;width_in=($right-$left);height_in=($top-$bottom)}
}
function Get-Endpoint($shape) {
    return [ordered]@{x_in=$shape.CellsU('EndX').ResultIU;y_in=$shape.CellsU('EndY').ResultIU}
}
function Close-PrivateVisio($app, $doc) {
    try { if ($null -ne $doc) { $doc.Close() } }
    finally { if ($null -ne $app) { try { $app.Quit() } finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) } } }
}
$app=$null; $doc=$null
$report = [ordered]@{schema_version=1;kind='synthetic editable-output demo';source='demo.vsdx';generation=[ordered]@{status='not_run';passed=$false};render=[ordered]@{status='not_run';passed=$false};interaction=[ordered]@{status='not_run';passed=$false};failure=$null}
try {
    $app = New-Object -ComObject Visio.InvisibleApp
    $app.EventsEnabled = $false; $app.AlertResponse = 7
    # Add a new blank document in our own invisible instance; never use GetActiveObject.
    $doc = $app.Documents.Add('')
    $doc.Title = 'Editable Visio recreation: synthetic output demo'
    $doc.Subject = 'Original synthetic diagram; no third-party source image'
    $doc.Creator = 'editable-visio-recreation contributors'
    $page = $doc.Pages.Item(1); $page.NameU = 'Synthetic demo'
    Set-Cell $page.PageSheet 'PageWidth' '80 mm'; Set-Cell $page.PageSheet 'PageHeight' '60 mm'
    $fontId = $doc.Fonts.Item('Times New Roman').ID
    $title = Add-Box $page 'DemoTitle' 'Editable Visio output' 6 49 74 55 $fontId $true
    Set-Cell $title 'Char.Style' '1'
    $source = Add-Box $page 'DefineLayout' "Define`nlayout" 5 27 25 40 $fontId
    $middle = Add-Box $page 'BuildObjects' "Build editable`nobjects" 30 27 50 40 $fontId
    $target = Add-Box $page 'EditableOutput' "Editable`nVSDX" 55 27 75 40 $fontId
    $first = Add-Connector $app $page 'LayoutToObjects' $source $middle
    $second = Add-Connector $app $page 'ObjectsToOutput' $middle $target
    $caption = Add-Box $page 'DemoCaption' "Synthetic editable-output demo`nShapes, text and native glued connectors" 5 8 75 20 $fontId $true
    # Eliminate workstation-specific author metadata from the public package.
    $doc.Creator = 'editable-visio-recreation contributors'
    $doc.SaveAs($vsdx)
    $ids = [ordered]@{resize_block=$source.ID;move_block=$target.ID;connector=$second.ID}
    $report.generation = [ordered]@{status='passed';passed=$true;page_mm=@(80,60);font='Times New Roman';font_pt=6;shape_count=$page.Shapes.Count;native_glued_connectors=2;visio_version=$app.Version;test_shape_ids=$ids}
} catch {
    $report.generation.status='failed'
    $report.failure = [ordered]@{stage='generation';message=($_.Exception.Message -replace [regex]::Escape($out),'[output]')}
    throw
} finally {
    Close-PrivateVisio $app $doc
    if (!$report.generation.passed) { $report | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $reportPath -Encoding utf8 }
}
try {
    # Visio sets lastModifiedBy from the workstation even when Creator is set.
    # Redact that author field before hashing/rendering the public demonstration.
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [IO.Compression.ZipFile]::Open($vsdx, [IO.Compression.ZipArchiveMode]::Update)
    try {
        $entry = $archive.GetEntry('docProps/core.xml')
        $reader = [IO.StreamReader]::new($entry.Open(), [Text.Encoding]::UTF8)
        try { [xml]$core = $reader.ReadToEnd() } finally { $reader.Dispose() }
        $lastAuthor = $core.SelectSingleNode("//*[local-name()='lastModifiedBy']")
        if ($null -ne $lastAuthor) { $lastAuthor.InnerText = 'editable-visio-recreation contributors' }
        $entry.Delete()
        $newEntry = $archive.CreateEntry('docProps/core.xml')
        $writer = [IO.StreamWriter]::new($newEntry.Open(), [Text.UTF8Encoding]::new($false))
        try { $writer.Write($core.OuterXml) } finally { $writer.Dispose() }
    } finally { $archive.Dispose() }
    # This invokes the actual-file renderer. The renderer requires a new directory.
    & (Join-Path $PSScriptRoot 'render_visio.ps1') -InputPath $vsdx -OutputDirectory $renderDir -Dpi 300
    $renderJsonPath = Join-Path $renderDir 'render.json'
    $render = Get-Content -Raw -LiteralPath $renderJsonPath | ConvertFrom-Json
    $render.source = '../demo.vsdx'
    $render | Add-Member -NotePropertyName 'path_base' -NotePropertyValue 'render.json directory'
    foreach ($item in $render.pages) { $item.png = [IO.Path]::GetFileName($item.png) }
    $render | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $renderJsonPath -Encoding utf8
    $report.render = [ordered]@{status='passed';passed=$true;renderer='Microsoft Visio';manifest='render/render.json';visual_review='Required separately'}
} catch {
    $report.render.status='failed'
    $report.failure = [ordered]@{stage='render';message=($_.Exception.Message -replace [regex]::Escape($out),'[output]')}
    $report | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $reportPath -Encoding utf8
    throw
}
$workRoot = Join-Path $out 'work'
$workDir = Join-Path $workRoot ('interaction-' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($workDir)
$copy = Join-Path $workDir 'disposable.vsdx'
Copy-Item -LiteralPath $vsdx -Destination $copy
$sourceHash = (Get-FileHash -LiteralPath $vsdx -Algorithm SHA256).Hash
$app=$null; $doc=$null
try {
    $app=New-Object -ComObject Visio.InvisibleApp; $app.EventsEnabled=$false; $app.AlertResponse=7
    # Hidden/unlisted, macros disabled, no workspace and no data refresh. Writable test copy.
    $doc=$app.Documents.OpenEx($copy,(8+64+128+256+1024))
    $page=$doc.Pages.Item(1)
    $resize=$page.Shapes.ItemFromID($ids.resize_block)
    $move=$page.Shapes.ItemFromID($ids.move_block)
    $connector=$page.Shapes.ItemFromID($ids.connector)
    $beforeBounds=Get-Bounds $resize
    $beforeWidth=$resize.CellsU('Width').ResultIU
    $beforeEndpoint=Get-Endpoint $connector
    $beforeTargetBounds=Get-Bounds $move
    $beforePin=[ordered]@{x_in=$move.CellsU('PinX').ResultIU;y_in=$move.CellsU('PinY').ResultIU}
    $resize.CellsU('Width').ResultIU=$beforeWidth*1.2
    # A displacement exceeding the 13 mm block height separates old and new bounds.
    # Dynamic walking glue may choose a different point on the block's boundary.
    $expectedMove=mm 18
    $move.CellsU('PinY').ResultIU=$beforePin.y_in+$expectedMove
    $doc.Save()
    $doc.Close(); $doc=$null
    # Reopen the saved copy to check persisted geometry/routing, not transient caches.
    $doc=$app.Documents.OpenEx($copy,(8+64+128+256+1024))
    $page=$doc.Pages.Item(1)
    $resize=$page.Shapes.ItemFromID($ids.resize_block); $move=$page.Shapes.ItemFromID($ids.move_block); $connector=$page.Shapes.ItemFromID($ids.connector)
    $afterBounds=Get-Bounds $resize
    $afterEndpoint=Get-Endpoint $connector
    $afterTargetBounds=Get-Bounds $move
    $afterPin=[ordered]@{x_in=$move.CellsU('PinX').ResultIU;y_in=$move.CellsU('PinY').ResultIU}
    $endFormula=$connector.CellsU('EndY').FormulaU
    $expectedIncrease=$beforeWidth*0.2
    $actualIncrease=$afterBounds.width_in-$beforeBounds.width_in
    $resizePass=([math]::Abs($actualIncrease-$expectedIncrease) -lt 0.001)
    $pinDelta=$afterPin.y_in-$beforePin.y_in
    $endpointDeltaX=$afterEndpoint.x_in-$beforeEndpoint.x_in
    $endpointDelta=$afterEndpoint.y_in-$beforeEndpoint.y_in
    $endpointDistance=[math]::Sqrt(($endpointDeltaX*$endpointDeltaX)+($endpointDelta*$endpointDelta))
    $tolerance=0.001
    $targetMoved=([math]::Abs($pinDelta-$expectedMove) -lt $tolerance -and
        [math]::Abs($afterPin.x_in-$beforePin.x_in) -lt $tolerance -and
        [math]::Abs(($afterTargetBounds.bottom_in-$beforeTargetBounds.bottom_in)-$expectedMove) -lt $tolerance -and
        [math]::Abs(($afterTargetBounds.top_in-$beforeTargetBounds.top_in)-$expectedMove) -lt $tolerance)
    $boundsSeparated=($afterTargetBounds.bottom_in -gt ($beforeTargetBounds.top_in+$tolerance))
    $endpointChanged=($endpointDistance -gt $tolerance)
    $insideNewBounds=($afterEndpoint.x_in -ge ($afterTargetBounds.left_in-$tolerance) -and
        $afterEndpoint.x_in -le ($afterTargetBounds.right_in+$tolerance) -and
        $afterEndpoint.y_in -ge ($afterTargetBounds.bottom_in-$tolerance) -and
        $afterEndpoint.y_in -le ($afterTargetBounds.top_in+$tolerance))
    $edgeDistances=@([math]::Abs($afterEndpoint.x_in-$afterTargetBounds.left_in),
        [math]::Abs($afterEndpoint.x_in-$afterTargetBounds.right_in),
        [math]::Abs($afterEndpoint.y_in-$afterTargetBounds.bottom_in),
        [math]::Abs($afterEndpoint.y_in-$afterTargetBounds.top_in))
    $boundaryDistance=($edgeDistances | Measure-Object -Minimum).Minimum
    $onNewBoundary=($insideNewBounds -and $boundaryDistance -le $tolerance)
    $outsideOldBounds=($afterEndpoint.x_in -lt ($beforeTargetBounds.left_in-$tolerance) -or
        $afterEndpoint.x_in -gt ($beforeTargetBounds.right_in+$tolerance) -or
        $afterEndpoint.y_in -lt ($beforeTargetBounds.bottom_in-$tolerance) -or
        $afterEndpoint.y_in -gt ($beforeTargetBounds.top_in+$tolerance))
    $glueRows=@()
    foreach($connect in $page.Connects) {
        if (($connect.FromSheet.ID -eq $ids.connector) -and ($connect.ToSheet.ID -eq $ids.move_block)) {
            $glueRows += [ordered]@{from_shape=$connect.FromSheet.ID;from_cell=$connect.FromCell.Name;to_shape=$connect.ToSheet.ID;to_cell=$connect.ToCell.Name}
        }
    }
    $gluePersisted=($glueRows.Count -gt 0)
    $connectorPass=($targetMoved -and $boundsSeparated -and $endpointChanged -and $onNewBoundary -and $outsideOldBounds -and $gluePersisted)
    $report.interaction=[ordered]@{status='completed';passed=($resizePass -and $connectorPass);tested_disposable_copy=$true;reopened_after_save=$true;resize=[ordered]@{passed=$resizePass;shape=$ids.resize_block;width_multiplier=1.2;before=$beforeBounds;after=$afterBounds;expected_visible_width_increase_in=$expectedIncrease;actual_visible_width_increase_in=$actualIncrease};connector_following=[ordered]@{passed=$connectorPass;connector=$ids.connector;target=$ids.move_block;before_target_bounds=$beforeTargetBounds;after_target_bounds=$afterTargetBounds;before_endpoint=$beforeEndpoint;after_endpoint=$afterEndpoint;expected_target_delta_y_in=$expectedMove;target_delta_y_in=$pinDelta;endpoint_delta_x_in=$endpointDeltaX;endpoint_delta_y_in=$endpointDelta;endpoint_distance_moved_in=$endpointDistance;distance_to_new_boundary_in=$boundaryDistance;tolerance_in=$tolerance;checks=[ordered]@{target_moved_expected_amount=$targetMoved;old_new_target_bounds_separated=$boundsSeparated;endpoint_position_changed=$endpointChanged;endpoint_on_new_target_boundary=$onNewBoundary;endpoint_outside_old_target_bounds=$outsideOldBounds;native_glue_relationship_persisted=$gluePersisted};persisted_end_y_formula=$endFormula;persisted_glue=$glueRows}}
    if (!$report.interaction.passed) { throw 'Visible resize or persisted native-connector following test failed.' }
} catch {
    $report.interaction.status='failed'
    $report.failure=[ordered]@{stage='interaction';message=($_.Exception.Message -replace [regex]::Escape($out),'[output]')}
    throw
} finally {
    Close-PrivateVisio $app $doc
    $report.source_unchanged=((Get-FileHash -LiteralPath $vsdx -Algorithm SHA256).Hash -eq $sourceHash)
    # Only our exact disposable file is removed; never recursively delete a computed path.
    if (Test-Path -LiteralPath $copy) { Remove-Item -LiteralPath $copy }
    if (Test-Path -LiteralPath $workDir) { Remove-Item -LiteralPath $workDir }
    if ((Test-Path -LiteralPath $workRoot) -and !(Get-ChildItem -LiteralPath $workRoot)) { Remove-Item -LiteralPath $workRoot }
    $report | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $reportPath -Encoding utf8
}
Write-Output ('Created demo.vsdx, render/render.json and interaction-check.json in ' + $out)

