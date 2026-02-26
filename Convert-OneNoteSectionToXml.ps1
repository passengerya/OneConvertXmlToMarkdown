param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$InputOneFile,

    [Parameter(Mandatory = $false, Position = 1)]
    [string]$OutputDirectory = ".\output\xml",

    [int]$LoadTimeoutSeconds = 30,

    [switch]$IncludeEmptyPages,

    [switch]$PassThru
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$OneNoteNamespace = "http://schemas.microsoft.com/office/onenote/2013/onenote"
$HierarchyScopePages = 4     # hsPages
$PageInfoAll = 7             # piAll
$XmlSchema2013 = 2           # xs2013
$CreateFileTypeNone = 0      # cftNone

function Write-Info {
    param([string]$Message)
    Write-Host "[INFO] $Message"
}

function ConvertTo-SafeName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,
        [int]$MaxLength = 100
    )

    $safe = $Name.Trim()
    if ([string]::IsNullOrWhiteSpace($safe)) {
        $safe = "Untitled"
    }

    foreach ($char in [System.IO.Path]::GetInvalidFileNameChars()) {
        $safe = $safe.Replace([string]$char, "_")
    }

    $safe = $safe -replace "\s+", " "
    $safe = $safe.TrimEnd(@([char]'.', [char]' '))

    if ([string]::IsNullOrWhiteSpace($safe)) {
        $safe = "Untitled"
    }

    $reserved = @(
        "CON","PRN","AUX","NUL",
        "COM1","COM2","COM3","COM4","COM5","COM6","COM7","COM8","COM9",
        "LPT1","LPT2","LPT3","LPT4","LPT5","LPT6","LPT7","LPT8","LPT9"
    )
    if ($reserved -contains $safe.ToUpperInvariant()) {
        $safe = "_$safe"
    }

    if ($safe.Length -gt $MaxLength) {
        $safe = $safe.Substring(0, $MaxLength).TrimEnd(@([char]'.', [char]' '))
    }

    if ([string]::IsNullOrWhiteSpace($safe)) {
        $safe = "Untitled"
    }

    return $safe
}

function New-NamespaceManager {
    param([xml]$XmlDocument)
    $ns = New-Object System.Xml.XmlNamespaceManager($XmlDocument.NameTable)
    $ns.AddNamespace("one", $OneNoteNamespace)
    return $ns
}

function Get-HierarchyXml {
    param(
        [Parameter(Mandatory = $true)]$OneNoteApp,
        [Parameter(Mandatory = $true)][string]$SectionId
    )

    $xml = ""
    $OneNoteApp.GetHierarchy($SectionId, $HierarchyScopePages, [ref]$xml, $XmlSchema2013)
    return $xml
}

function Get-SectionHierarchyLoaded {
    param(
        [Parameter(Mandatory = $true)]$OneNoteApp,
        [Parameter(Mandatory = $true)][string]$SectionId,
        [int]$TimeoutSeconds = 30
    )

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $lastPageCount = -1
    $stableRounds = 0
    $stableRoundsRequired = 2
    $bestHierarchyXml = $null
    $bestPageCount = 0
    $lastAreAllPagesAvailable = $false

    while ($stopwatch.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
        try {
            $OneNoteApp.NavigateTo($SectionId)
        } catch {
        }

        Start-Sleep -Milliseconds 1000

        try {
            $OneNoteApp.SyncHierarchy($SectionId)
        } catch {
        }

        $hierarchyXml = Get-HierarchyXml -OneNoteApp $OneNoteApp -SectionId $SectionId
        [xml]$doc = $hierarchyXml
        $pages = $doc.SelectNodes("//*[local-name()='Page']")
        $pageCount = $pages.Count

        $sectionNode = $doc.SelectSingleNode("/*[local-name()='Section']")
        $areAllPagesAvailable = $true
        if ($null -ne $sectionNode) {
            $attr = $sectionNode.Attributes["areAllPagesAvailable"]
            if ($null -ne $attr -and [string]$attr.Value -eq "false") {
                $areAllPagesAvailable = $false
            }
        }
        $lastAreAllPagesAvailable = $areAllPagesAvailable

        if ($pageCount -gt $bestPageCount) {
            $bestPageCount = $pageCount
            $bestHierarchyXml = $hierarchyXml
        }

        if ($pageCount -gt 0 -and $pageCount -eq $lastPageCount) {
            $stableRounds++
        } else {
            $stableRounds = 0
        }

        # Wait until page count stops changing for a few checks, and OneNote no longer reports partial availability.
        if ($pageCount -gt 0 -and $stableRounds -ge $stableRoundsRequired -and $areAllPagesAvailable) {
            return @{
                Xml = $hierarchyXml
                PageCount = $pageCount
            }
        }

        $lastPageCount = $pageCount
    }

    if ($bestPageCount -gt 0 -and $null -ne $bestHierarchyXml) {
        Write-Warning "Hierarchy did not fully stabilize within timeout. Using best available hierarchy with $bestPageCount page(s). Consider increasing -LoadTimeoutSeconds."
        if (-not $lastAreAllPagesAvailable) {
            Write-Warning "OneNote still reported areAllPagesAvailable=false at timeout. Export may be incomplete."
        }
        return @{
            Xml = $bestHierarchyXml
            PageCount = $bestPageCount
        }
    }

    throw "Failed to load page hierarchy within $TimeoutSeconds seconds (last page count: $lastPageCount). Make sure OneNote is installed and can open this .one file."
}

function Test-PageHasBodyContent {
    param(
        [Parameter(Mandatory = $true)][xml]$PageXmlDoc
    )

    # Body content is typically under one:Outline. We also count non-text objects as content.
    $xpath = "//*[local-name()='Outline' or local-name()='InkDrawing' or local-name()='Image' or local-name()='InsertedFile' or local-name()='Audio' or local-name()='Video']"
    $bodyObject = $PageXmlDoc.SelectSingleNode($xpath)

    return ($null -ne $bodyObject)
}

function Save-Utf8File {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Content
    )

    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, $Content, $utf8NoBom)
}

function Get-UniqueFilePath {
    param(
        [Parameter(Mandatory = $true)][string]$Directory,
        [Parameter(Mandatory = $true)][string]$BaseName,
        [string]$Extension = ".xml"
    )

    $candidate = Join-Path $Directory ($BaseName + $Extension)
    if (-not (Test-Path -LiteralPath $candidate)) {
        return $candidate
    }

    $index = 2
    while ($true) {
        $candidate = Join-Path $Directory ("{0} ({1}){2}" -f $BaseName, $index, $Extension)
        if (-not (Test-Path -LiteralPath $candidate)) {
            return $candidate
        }
        $index++
    }
}

if (-not (Test-Path -LiteralPath $InputOneFile -PathType Leaf)) {
    throw "Input file not found: $InputOneFile"
}

$inputPath = [System.IO.Path]::GetFullPath((Resolve-Path -LiteralPath $InputOneFile).Path)
$outputRoot = [System.IO.Path]::GetFullPath($OutputDirectory)
# Create default/custom output root automatically when missing.
New-Item -ItemType Directory -Force -Path $outputRoot | Out-Null

Write-Info "Input file: $inputPath"
Write-Info "Output dir: $outputRoot"

$oneNote = New-Object -ComObject OneNote.Application
$sectionId = ""
$oneNote.OpenHierarchy($inputPath, "", [ref]$sectionId, $CreateFileTypeNone)

Write-Info "Opened section. Section ID: $sectionId"
Write-Info "Waiting for OneNote to load page hierarchy..."

$loaded = Get-SectionHierarchyLoaded -OneNoteApp $oneNote -SectionId $sectionId -TimeoutSeconds $LoadTimeoutSeconds
$hierarchyXml = [string]$loaded.Xml

[xml]$hierarchyDoc = $hierarchyXml
$sectionNode = $hierarchyDoc.SelectSingleNode("/*[local-name()='Section']")
if ($null -eq $sectionNode) {
    throw "Could not parse section node from hierarchy XML."
}

$sectionName = [string]$sectionNode.GetAttribute("name")
$sectionDirName = ConvertTo-SafeName -Name $sectionName
$sectionOutputDir = Join-Path $outputRoot $sectionDirName
if (Test-Path -LiteralPath $sectionOutputDir) {
    # Re-run safety: clear previous export for this section to avoid generating "(2)" duplicates.
    Write-Info "Cleaning existing section output directory..."
    Remove-Item -LiteralPath $sectionOutputDir -Recurse -Force
}
New-Item -ItemType Directory -Force -Path $sectionOutputDir | Out-Null

$hierarchyXmlPath = Join-Path $sectionOutputDir "section-hierarchy.xml"
Save-Utf8File -Path $hierarchyXmlPath -Content $hierarchyXml

$pages = $hierarchyDoc.SelectNodes("/*[local-name()='Section']/*[local-name()='Page']")
Write-Info ("Detected {0} page nodes" -f $pages.Count)

$folderStack = New-Object System.Collections.ArrayList
$numberStack = New-Object System.Collections.ArrayList
$childCounters = @{}
$seenFolderNamesByParent = @{}

$stats = [ordered]@{
    TotalPages       = 0
    ExportedXmlPages = 0
    SkippedEmpty     = 0
    CreatedFolders   = 0
}

for ($i = 0; $i -lt $pages.Count; $i++) {
    $pageNode = $pages.Item($i)
    $stats.TotalPages++

    $title = [string]$pageNode.GetAttribute("name")
    if ([string]::IsNullOrWhiteSpace($title)) {
        $title = "Untitled"
    }

    $levelRaw = [string]$pageNode.GetAttribute("pageLevel")
    $level = 1
    if (-not [int]::TryParse($levelRaw, [ref]$level)) {
        $level = 1
    }
    if ($level -lt 1) { $level = 1 }

    # Clamp malformed jumps (e.g., level 4 immediately after level 1) to keep a valid tree on disk.
    $maxValidLevel = $folderStack.Count + 1
    if ($level -gt $maxValidLevel) { $level = $maxValidLevel }

    $nextLevel = 0
    if ($i + 1 -lt $pages.Count) {
        $nextNode = $pages.Item($i + 1)
        $nextLevelRaw = [string]$nextNode.GetAttribute("pageLevel")
        if (-not [int]::TryParse($nextLevelRaw, [ref]$nextLevel)) {
            $nextLevel = 1
        }
        if ($nextLevel -lt 1) { $nextLevel = 1 }
    }
    $hasChildren = ($nextLevel -gt $level)

    while ($folderStack.Count -ge $level) {
        $folderStack.RemoveAt($folderStack.Count - 1)
    }
    while ($numberStack.Count -ge $level) {
        $numberStack.RemoveAt($numberStack.Count - 1)
    }

    $parentKey = ($numberStack | ForEach-Object { [string]$_ }) -join "."
    if (-not $childCounters.ContainsKey($parentKey)) {
        $childCounters[$parentKey] = 0
    }
    $childCounters[$parentKey] = [int]$childCounters[$parentKey] + 1
    [void]$numberStack.Add([int]$childCounters[$parentKey])

    $numLabel = ($numberStack | ForEach-Object { ([int]$_).ToString("D3") }) -join "."
    $baseFolderName = "{0} {1}" -f $numLabel, (ConvertTo-SafeName -Name $title)

    $parentDir = if ($folderStack.Count -eq 0) { $sectionOutputDir } else { [string]$folderStack[$folderStack.Count - 1] }
    $parentDirKey = $parentDir.ToLowerInvariant()
    if (-not $seenFolderNamesByParent.ContainsKey($parentDirKey)) {
        $seenFolderNamesByParent[$parentDirKey] = @{}
    }

    $folderName = $baseFolderName
    $suffix = 2
    while ($seenFolderNamesByParent[$parentDirKey].ContainsKey($folderName.ToLowerInvariant())) {
        $folderName = "{0} ({1})" -f $baseFolderName, $suffix
        $suffix++
    }
    $seenFolderNamesByParent[$parentDirKey][$folderName.ToLowerInvariant()] = $true

    $pageDir = $null
    $xmlOutputDir = $parentDir
    if ($hasChildren) {
        $pageDir = Join-Path $parentDir $folderName
        New-Item -ItemType Directory -Force -Path $pageDir | Out-Null
        $stats.CreatedFolders++
        [void]$folderStack.Add($pageDir)
        $xmlOutputDir = $pageDir
    }

    $pageId = [string]$pageNode.GetAttribute("ID")
    $pageXml = ""
    $oneNote.GetPageContent($pageId, [ref]$pageXml, $PageInfoAll, $XmlSchema2013)

    [xml]$pageDoc = $pageXml
    $hasBody = Test-PageHasBodyContent -PageXmlDoc $pageDoc

    if ($IncludeEmptyPages -or $hasBody) {
        $pageFileBaseName = ConvertTo-SafeName -Name $title
        $pageXmlPath = Get-UniqueFilePath -Directory $xmlOutputDir -BaseName $pageFileBaseName -Extension ".xml"
        Save-Utf8File -Path $pageXmlPath -Content $pageXml
        $stats.ExportedXmlPages++
    } else {
        $stats.SkippedEmpty++
    }
}

Write-Host ""
Write-Host "Done."
Write-Host ("- Section dir: {0}" -f $sectionOutputDir)
Write-Host ("- Total pages: {0}" -f $stats.TotalPages)
Write-Host ("- Exported XML pages: {0}" -f $stats.ExportedXmlPages)
Write-Host ("- Skipped empty pages: {0}" -f $stats.SkippedEmpty)
Write-Host ("- Hierarchy XML: {0}" -f $hierarchyXmlPath)

if ($PassThru) {
    [pscustomobject]@{
        InputFile         = $inputPath
        OutputRoot        = $outputRoot
        SectionName       = $sectionName
        SectionDirectory  = $sectionOutputDir
        HierarchyXmlPath  = $hierarchyXmlPath
        TotalPages        = [int]$stats.TotalPages
        ExportedXmlPages  = [int]$stats.ExportedXmlPages
        SkippedEmptyPages = [int]$stats.SkippedEmpty
    }
}
