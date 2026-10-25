# Optional local renderer. Requires Windows PowerShell 5.1+ and approved desktop PowerPoint.
# Never use -ExecutionPolicy Bypass to run this file. Follow your company policy.
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$RunPath,
    [ValidateRange(320, 4096)][int]$Width = 1920
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$application = $null
$presentations = $null
$presentation = $null
$slides = $null
$pageSetup = $null
$stagePath = $null
$finished = $false
$ownsApplication = $false
$previousSecurity = $null
$securityChanged = $false

function Assert-LocalPath([string]$Path) {
    if ($Path.StartsWith('\\') -or $Path.StartsWith('//') -or $Path -match '(^|[\\/])\.\.([\\/]|$)') {
        throw 'UNC/device paths and parent traversal are not allowed.'
    }
    $absolute = [IO.Path]::GetFullPath($Path)
    $anchor = [IO.Path]::GetPathRoot($absolute)
    $drive = New-Object IO.DriveInfo($anchor)
    if ($drive.DriveType -eq [IO.DriveType]::Network) { throw "Network drive is not allowed: $anchor" }
    $cursor = $absolute
    while ($cursor) {
        if (Test-Path -LiteralPath $cursor) {
            $item = Get-Item -LiteralPath $cursor -Force
            if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
                throw "Symlink/junction/reparse path is not allowed: $cursor"
            }
        }
        $parent = [IO.Directory]::GetParent($cursor)
        if ($null -eq $parent) { break }
        $cursor = $parent.FullName
    }
    return $absolute
}

function Release-Com($Object) {
    if ($null -ne $Object -and [Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
}

function Get-SafeXml([IO.Stream]$Stream) {
    $settings = New-Object Xml.XmlReaderSettings
    $settings.DtdProcessing = [Xml.DtdProcessing]::Prohibit
    $settings.XmlResolver = $null
    $settings.MaxCharactersInDocument = 8388608
    $reader = [Xml.XmlReader]::Create($Stream, $settings)
    try {
        $document = New-Object Xml.XmlDocument
        $document.XmlResolver = $null
        $document.Load($reader)
        return ,$document
    } finally { $reader.Dispose() }
}

try {
    $run = Assert-LocalPath $RunPath
    if (-not (Test-Path -LiteralPath $run -PathType Container)) { throw 'RunPath must be an existing extracted run directory.' }
    $manifestPath = Assert-LocalPath (Join-Path $run 'manifest.json')
    $sourcePath = Assert-LocalPath (Join-Path $run 'source\original.pptx')
    $assetsPath = Assert-LocalPath (Join-Path $run 'assets')
    $slidesPath = Assert-LocalPath (Join-Path $assetsPath 'slides')
    $renderManifestPath = Assert-LocalPath (Join-Path $run 'render-manifest.json')
    if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) { throw 'Missing source/original.pptx.' }
    if ((Get-Item -LiteralPath $manifestPath).Length -gt 10485760) { throw 'Manifest exceeds 10 MiB.' }
    $manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($manifest.source.path -ne 'source/original.pptx') { throw 'Unexpected manifest source path.' }
    $sourceHash = (Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash.ToLowerInvariant()
    if ($sourceHash -ne [string]$manifest.source.sha256) { throw 'Source SHA-256 does not match manifest. Stop and investigate.' }
    $sourceHashAlias = $manifest.PSObject.Properties['source_sha256']
    if ($null -ne $sourceHashAlias -and $sourceHashAlias.Value -ne $sourceHash) { throw 'Conflicting source SHA-256 fields in manifest.' }
    $slideCount = [int]$manifest.slide_count
    if ($slideCount -le 0 -or @($manifest.slides).Count -ne $slideCount) { throw 'Manifest has no slides or inconsistent slide count.' }
    for ($number = 1; $number -le $slideCount; $number++) {
        if ([int]$manifest.slides[$number - 1].slide_no -ne $number) { throw 'Manifest slide order must be contiguous and one-based.' }
        $target = Join-Path $slidesPath ('slide-{0:D4}.png' -f $number)
        [void](Assert-LocalPath $target)
        if (Test-Path -LiteralPath $target) { throw "Refusing to overwrite existing render: $target" }
    }
    if (Test-Path -LiteralPath $renderManifestPath) { throw 'render-manifest.json already exists. Use a new run; nothing was overwritten.' }
    if ((Test-Path -LiteralPath $assetsPath) -and -not (Test-Path -LiteralPath $assetsPath -PathType Container)) { throw 'assets is not a directory.' }
    if ((Test-Path -LiteralPath $slidesPath) -and -not (Test-Path -LiteralPath $slidesPath -PathType Container)) { throw 'assets/slides is not a directory.' }

    # Inspect ALL relationship files before invoking Office; no external links are allowed.
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::OpenRead($sourcePath)
    try {
        foreach ($entry in $zip.Entries) {
            $name = $entry.FullName.Replace('\', '/')
            if ($name.StartsWith('/') -or $name -match '(^|/)\.\.(/|$)' -or $name.Contains(':')) { throw 'Unsafe member path in PPTX archive.' }
            if ($name -match '(?i)(vbaProject|vbaData|/activeX/|/embeddings/)') {
                throw 'Active content or embedded OLE object found. Automatic Office rendering is blocked; use an approved manual export.'
            }
            if ($name -eq '[Content_Types].xml') {
                if ($entry.Length -gt 8388608) { throw 'Content type XML is unexpectedly large.' }
                $stream = $entry.Open()
                try { $typesXml = Get-SafeXml $stream } finally { $stream.Dispose() }
                foreach ($contentType in $typesXml.SelectNodes('//*[@ContentType]')) {
                    if ($contentType.GetAttribute('ContentType') -match '(?i)(macroEnabled|vbaProject|activeX|oleObject)') {
                        throw 'Active Office content type found. Automatic rendering is blocked; use an approved manual export.'
                    }
                }
            }
            if ($name.EndsWith('.rels', [StringComparison]::OrdinalIgnoreCase)) {
                if ($entry.Length -gt 8388608) { throw "Relationship XML is unexpectedly large: $name" }
                $stream = $entry.Open()
                try { $xml = Get-SafeXml $stream } finally { $stream.Dispose() }
                foreach ($relationship in $xml.SelectNodes("//*[local-name()='Relationship']")) {
                    if ($relationship.GetAttribute('TargetMode') -eq 'External') {
                        throw "External relationship found in $name (including hyperlinks). Automatic rendering is blocked. Use an approved manual export and register it with the pipeline."
                    }
                    if ($relationship.GetAttribute('Type') -match '(?i)/(vbaProject|vbaData|control|activeXControl|oleObject|package)$') {
                        throw "Active or embedded Office relationship found in $name. Automatic rendering is blocked; use an approved manual export."
                    }
                }
            }
        }
    } finally { $zip.Dispose() }

    # PowerPoint can reuse an existing automation server. Do not touch a user's open session.
    if (@(Get-Process -Name POWERPNT -ErrorAction SilentlyContinue).Count -gt 0) {
        throw 'PowerPoint is already running. Save and close it before optional COM rendering, or use manual export.'
    }
    $stagePath = Join-Path $run ('.render-staging-' + [Guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($stagePath)
    try { $application = New-Object -ComObject PowerPoint.Application }
    catch { throw 'Desktop PowerPoint COM is unavailable. No render completed. Install nothing automatically; use company-approved manual slide export.' }
    $presentations = $application.Presentations
    if ($presentations.Count -ne 0) { throw 'A PowerPoint session with open presentations was found. Rendering was cancelled.' }
    $ownsApplication = $true
    # Save/restore the prior value, including if Open fails.
    $previousSecurity = $application.AutomationSecurity
    # msoAutomationSecurityForceDisable = 3. Fail closed if setting is unsupported.
    $application.AutomationSecurity = 3
    $securityChanged = $true
    if ([int]$application.AutomationSecurity -ne 3) { throw 'Could not force-disable Office macros.' }
    # ReadOnly = msoTrue (-1), Untitled = msoFalse (0), WithWindow = msoFalse (0).
    $presentation = $presentations.Open($sourcePath, -1, 0, 0)
    $application.AutomationSecurity = $previousSecurity
    $securityChanged = $false
    $slides = $presentation.Slides
    if ($slides.Count -ne $slideCount) { throw 'PowerPoint slide count differs from manifest.' }
    $pageSetup = $presentation.PageSetup
    if ($pageSetup.SlideWidth -le 0 -or $pageSetup.SlideHeight -le 0) { throw 'Invalid slide dimensions.' }
    $height = [int][Math]::Round($Width * $pageSetup.SlideHeight / $pageSetup.SlideWidth)
    if ($height -le 0 -or $height -gt 16384) { throw 'Unsupported slide image height.' }
    Add-Type -AssemblyName System.Drawing
    $renders = New-Object 'System.Collections.Generic.List[object]'
    # Enumerating Slides includes hidden slides; presentation order is preserved.
    for ($number = 1; $number -le $slideCount; $number++) {
        $slide = $null
        $fileName = 'slide-{0:D4}.png' -f $number
        $imagePath = Join-Path $stagePath $fileName
        try {
            $slide = $slides.Item($number)
            $slide.Export($imagePath, 'PNG', $Width, $height)
        } finally { Release-Com $slide }
        if (-not (Test-Path -LiteralPath $imagePath -PathType Leaf)) { throw "PowerPoint did not export slide $number." }
        $bitmap = [Drawing.Image]::FromFile($imagePath)
        try { $actualWidth = $bitmap.Width; $actualHeight = $bitmap.Height } finally { $bitmap.Dispose() }
        if ($actualWidth -le 0 -or $actualHeight -le 0) { throw "Invalid image for slide $number." }
        $renders.Add([ordered]@{ slide_no = $number; path = "assets/slides/$fileName"; sha256 = (Get-FileHash -LiteralPath $imagePath -Algorithm SHA256).Hash.ToLowerInvariant(); width = $actualWidth; height = $actualHeight })
    }
    if ((Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash.ToLowerInvariant() -ne $sourceHash) { throw 'Source file changed during rendering.' }
    $result = [ordered]@{ schema_version = 1; source_sha256 = $sourceHash; method = 'powerpoint_com'; rendered_at_utc = [DateTime]::UtcNow.ToString('o'); renders = @($renders.ToArray()) }
    $stagedManifest = Join-Path $stagePath 'render-manifest.json'
    [IO.File]::WriteAllText($stagedManifest, ($result | ConvertTo-Json -Depth 8), (New-Object Text.UTF8Encoding($false)))
    [void](Assert-LocalPath $slidesPath)
    [void][IO.Directory]::CreateDirectory($slidesPath)
    # File.Move is exclusive on Windows PowerShell/.NET: it fails if destination exists.
    foreach ($render in $renders) {
        $destination = Join-Path $run $render.path
        [void](Assert-LocalPath $destination)
        [IO.File]::Move((Join-Path $stagePath ([IO.Path]::GetFileName($render.path))), $destination)
    }
    [void](Assert-LocalPath $renderManifestPath)
    [IO.File]::Move($stagedManifest, $renderManifestPath)
    $finished = $true
    Write-Output "Rendered $slideCount slides. Verify visuals and run pipeline validation. Manifest: $renderManifestPath"
    Write-Output "Empty staging directory retained: $stagePath"
} catch {
    [Console]::Error.WriteLine('ERROR: ' + $_.Exception.Message)
    if ($stagePath) { [Console]::Error.WriteLine('Staging files were retained for review: ' + $stagePath) }
    [Console]::Error.WriteLine('No success is claimed. Existing artifacts were never overwritten; any partial new outputs must be reviewed. Use a new run for retries.')
} finally {
    if ($securityChanged -and $null -ne $application) {
        try { $application.AutomationSecurity = $previousSecurity } catch { }
    }
    if ($null -ne $presentation) { try { $presentation.Close() } catch { } }
    $canQuit = $false
    if ($ownsApplication -and $null -ne $presentations) {
        try { $canQuit = ($presentations.Count -eq 0) } catch { }
    }
    Release-Com $pageSetup
    Release-Com $slides
    Release-Com $presentation
    Release-Com $presentations
    # Only the COM instance created above is asked to quit; never kill processes.
    if ($canQuit -and $null -ne $application) { try { $application.Quit() } catch { } }
    Release-Com $application
}
if (-not $finished) { exit 1 }
exit 0
