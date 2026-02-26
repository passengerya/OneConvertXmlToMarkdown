param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$InputOneFile,

    [Parameter(Mandatory = $false)]
    [string]$XmlOutputDirectory = ".\output\xml",

    [Parameter(Mandatory = $false)]
    [string]$MarkdownOutputDirectory = ".\output\markdown",

    [int]$LoadTimeoutSeconds = 30,

    [switch]$IncludeEmptyPages,

    [ValidateSet("markdown", "obsidian")]
    [string]$ImageSyntax = "markdown",

    [switch]$CopyAssets = $true,

    [string]$AssetDirectoryName = "attachment",

    [switch]$SkipMarkdownStage,

    [switch]$PassThru
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Write-Step {
    param([string]$Message)
    Write-Host "[PIPELINE] $Message"
}

function Resolve-PythonCommand {
    $python = Get-Command python -ErrorAction SilentlyContinue
    if ($python) {
        return @($python.Source)
    }

    $py = Get-Command py -ErrorAction SilentlyContinue
    if ($py) {
        return @($py.Source, "-3")
    }

    throw "Python not found. Install Python 3 and ensure 'python' or 'py' is available in PATH."
}

$scriptRoot = Split-Path -Parent $PSCommandPath
$xmlConverterScript = Join-Path $scriptRoot "Convert-OneNoteSectionToXml.ps1"
$mdConverterScript = Join-Path $scriptRoot "convert_onenote_xml.py"

if (-not (Test-Path -LiteralPath $xmlConverterScript -PathType Leaf)) {
    throw "XML converter script not found: $xmlConverterScript"
}
if (-not (Test-Path -LiteralPath $mdConverterScript -PathType Leaf)) {
    throw "Markdown converter script not found: $mdConverterScript"
}

Write-Step "Stage 1/2: Converting .one to hierarchical XML..."

$xmlResult = & $xmlConverterScript `
    -InputOneFile $InputOneFile `
    -OutputDirectory $XmlOutputDirectory `
    -LoadTimeoutSeconds $LoadTimeoutSeconds `
    -IncludeEmptyPages:$IncludeEmptyPages `
    -PassThru

if ($null -eq $xmlResult) {
    throw "XML converter did not return pipeline metadata."
}

$xmlSectionDir = [string]$xmlResult.SectionDirectory
$sectionName = [string]$xmlResult.SectionName

Write-Step ("XML stage done. Section XML dir: {0}" -f $xmlSectionDir)

if ($SkipMarkdownStage) {
    Write-Step "Markdown stage skipped by parameter."
    if ($PassThru) {
        return [pscustomobject]@{
            Xml = $xmlResult
            Markdown = $null
        }
    }
    return
}

$mdRoot = [System.IO.Path]::GetFullPath($MarkdownOutputDirectory)
$mdSectionDir = Join-Path $mdRoot $sectionName
New-Item -ItemType Directory -Force -Path $mdRoot | Out-Null
# Re-run safety: clear previous markdown output for the same section to avoid duplicate files.
if (Test-Path -LiteralPath $mdSectionDir) {
    Write-Step ("Cleaning existing Markdown section directory: {0}" -f $mdSectionDir)
    Remove-Item -LiteralPath $mdSectionDir -Recurse -Force
}

$pythonCmd = @(Resolve-PythonCommand)
$pythonExe = $pythonCmd[0]
$pythonPrefixArgs = @()
if ($pythonCmd.Count -gt 1) {
    $pythonPrefixArgs = $pythonCmd[1..($pythonCmd.Count - 1)]
}

$mdArgs = @()
$mdArgs += $pythonPrefixArgs
$mdArgs += @(
    $mdConverterScript,
    $xmlSectionDir,
    "-o", $mdSectionDir,
    "--recursive",
    "--image-syntax", $ImageSyntax
)
if ($CopyAssets) {
    $mdAssetArg = $AssetDirectoryName
    if (-not [string]::IsNullOrWhiteSpace($mdAssetArg)) {
        $normalized = $mdAssetArg.Replace("\", "/")
        if ($normalized -notmatch "^\.\.?/" -and $normalized -notmatch "/") {
            # For a simple folder name like "attachment", convert it to ../attachment so assets are siblings of the markdown section folder.
            $mdAssetArg = "../$normalized"
        }
    }
    $mdArgs += "--copy-attachments"
    $mdArgs += @("--asset-dir", $mdAssetArg)
}

Write-Step "Stage 2/2: Converting hierarchical XML to hierarchical Markdown..."
Write-Step ("Python command: {0} {1}" -f $pythonExe, ($mdArgs -join " "))

& $pythonExe @mdArgs
if ($LASTEXITCODE -ne 0) {
    throw "Markdown converter failed with exit code $LASTEXITCODE"
}

Write-Host ""
Write-Host "Pipeline completed."
Write-Host ("- XML root: {0}" -f $xmlResult.OutputRoot)
Write-Host ("- XML section dir: {0}" -f $xmlSectionDir)
Write-Host ("- Markdown root: {0}" -f $mdRoot)
Write-Host ("- Markdown section dir: {0}" -f $mdSectionDir)

if ($PassThru) {
    [pscustomobject]@{
        Xml = $xmlResult
        Markdown = [pscustomobject]@{
            OutputRoot       = $mdRoot
            SectionDirectory = $mdSectionDir
            ImageSyntax      = $ImageSyntax
            CopyAssets       = [bool]$CopyAssets
            AssetDirectory   = $AssetDirectoryName
        }
    }
}
