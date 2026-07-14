param(
    [Parameter(Mandatory = $true)]
    [string[]]$Paths,
    [string]$PdfDir = "",
    [ValidateSet("Auto", "WPS", "Word")]
    [string]$Office = "Auto"
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

function New-OfficeApplication {
    param([string]$Preference)
    $attempts = switch ($Preference) {
        "WPS" { @("KWPS.Application") }
        "Word" { @("Word.Application") }
        default { @("KWPS.Application", "Word.Application") }
    }
    foreach ($progId in $attempts) {
        try {
            return [pscustomobject]@{
                App = New-Object -ComObject $progId
                ProgId = $progId
            }
        } catch {
            continue
        }
    }
    throw "Neither WPS Writer nor Microsoft Word COM automation is available. Do not claim pagination was verified."
}

function Update-DocumentFields {
    param($Document)
    $updated = 0
    if ([int]$Document.Fields.Count -gt 0) {
        $Document.Fields.Update() | Out-Null
        $updated += [int]$Document.Fields.Count
    }
    foreach ($section in $Document.Sections) {
        foreach ($collection in @($section.Headers, $section.Footers)) {
            foreach ($index in 1..3) {
                $story = $null
                try { $story = $collection.Item($index) } catch { continue }
                $exists = $true
                try { $exists = [bool]$story.Exists } catch {}
                if ($exists -and [int]$story.Range.Fields.Count -gt 0) {
                    # An existing field that cannot be updated is a hard error;
                    # otherwise cached page totals may remain stale.
                    $story.Range.Fields.Update() | Out-Null
                    $updated += [int]$story.Range.Fields.Count
                }
            }
        }
    }
    return $updated
}

function Invoke-Repaginate {
    param($Document)
    try {
        $Document.Repaginate() | Out-Null
        return $true
    } catch {
        # Some WPS COM versions do not expose Repaginate. ComputeStatistics and
        # PDF export below still force layout; the result records this fallback.
        return $false
    }
}

$resolved = foreach ($path in $Paths) {
    $resolvedPath = (Resolve-Path -LiteralPath $path).Path
    if ([IO.Path]::GetExtension($resolvedPath).ToLowerInvariant() -ne ".docx") {
        throw "Only .docx files are supported: $resolvedPath"
    }
    $resolvedPath
}
$duplicates = $resolved | Group-Object { $_.ToLowerInvariant() } | Where-Object Count -gt 1
if ($duplicates) {
    throw "The same DOCX path was supplied more than once."
}

if ($PdfDir) {
    $pdfRoot = [IO.Path]::GetFullPath($PdfDir)
    New-Item -ItemType Directory -Path $pdfRoot -Force | Out-Null
} else {
    $pdfRoot = ""
}

if ($pdfRoot) {
    $pdfNames = $resolved | ForEach-Object { ([IO.Path]::GetFileNameWithoutExtension($_) + ".pdf").ToLowerInvariant() }
    if ($pdfNames.Count -ne ($pdfNames | Select-Object -Unique).Count) {
        throw "DOCX basenames collide in PdfDir; use unique filenames before rendering."
    }
}

$officeInstance = New-OfficeApplication -Preference $Office
$app = $officeInstance.App
$results = @()

try {
    $app.Visible = $false
    try { $app.DisplayAlerts = 0 } catch {}

    foreach ($path in $resolved) {
        $doc = $null
        try {
            $doc = $app.Documents.Open($path, $false, $false)
            $measurements = @()
            $updatedFields = 0
            $repaginateSupported = $true
            for ($pass = 1; $pass -le 3; $pass++) {
                $updatedFields += Update-DocumentFields -Document $doc
                if (-not (Invoke-Repaginate -Document $doc)) { $repaginateSupported = $false }
                $doc.Save()
                $pages = [int]$doc.ComputeStatistics(2)
                if ($pages -le 0) { throw "Office returned a non-positive page count for $path" }
                $measurements += $pages
                if ($measurements.Count -ge 2 -and $measurements[-1] -eq $measurements[-2]) { break }
            }
            if ($measurements.Count -lt 2 -or $measurements[-1] -ne $measurements[-2]) {
                throw "Pagination did not stabilize after three Office passes: $path ($($measurements -join ', '))"
            }
            $pages = [int]$measurements[-1]
            $pdfPath = ""
            if ($pdfRoot) {
                $pdfPath = Join-Path $pdfRoot (([IO.Path]::GetFileNameWithoutExtension($path)) + ".pdf")
                $doc.ExportAsFixedFormat($pdfPath, 17)
                if (-not (Test-Path -LiteralPath $pdfPath) -or (Get-Item -LiteralPath $pdfPath).Length -le 0) {
                    throw "Office did not create a non-empty PDF preview: $pdfPath"
                }
                $afterExport = [int]$doc.ComputeStatistics(2)
                if ($afterExport -ne $pages) {
                    throw "Pagination changed during PDF export: $path ($pages -> $afterExport)"
                }
            }
            $results += [pscustomobject]@{
                path = $path
                pages = $pages
                paragraphs = [int]$doc.Paragraphs.Count
                tables = [int]$doc.Tables.Count
                pdf = $pdfPath
                fields_updated = $updatedFields
                repaginate_supported = $repaginateSupported
                stable_measurements = $measurements
            }
        } finally {
            if ($null -ne $doc) {
                $doc.Close($false)
            }
        }
    }
} finally {
    try {
        if ($null -ne $app) {
            $app.Quit()
        }
    } finally {
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

[pscustomobject]@{
    office = $officeInstance.ProgId
    results = $results
} | ConvertTo-Json -Depth 5
