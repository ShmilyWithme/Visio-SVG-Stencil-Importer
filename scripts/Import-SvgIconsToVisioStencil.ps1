param(
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$IconDirectory,

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$StencilPath,

    [switch]$Force,
    [switch]$KeepIntermediate
)

$ErrorActionPreference = 'Stop'

$IconDirectory = [System.IO.Path]::GetFullPath($IconDirectory)
$StencilPath = [System.IO.Path]::GetFullPath($StencilPath)

if (-not (Test-Path -LiteralPath $IconDirectory -PathType Container)) {
    throw "SVG directory does not exist: $IconDirectory"
}

$svgFiles = @(Get-ChildItem -LiteralPath $IconDirectory -Filter '*.svg' -File | Sort-Object Name)
if ($svgFiles.Count -eq 0) {
    throw "No SVG files found in: $IconDirectory"
}

if ((Test-Path -LiteralPath $StencilPath) -and -not $Force) {
    throw "Output already exists: $StencilPath. Re-run with -Force only if replacement is intended."
}

$outputDirectory = Split-Path -Parent $StencilPath
if (-not (Test-Path -LiteralPath $outputDirectory -PathType Container)) {
    New-Item -ItemType Directory -Path $outputDirectory | Out-Null
}

$visio = $null
$stencil = $null
$powerPoint = $null
$presentation = $null
$tempDirectory = Join-Path ([System.IO.Path]::GetTempPath()) ("visio-svg-stencil-$PID")
$temporaryStencilPath = Join-Path $tempDirectory 'svg-icons.vssx'

try {
    New-Item -ItemType Directory -Path $tempDirectory | Out-Null

    $powerPoint = New-Object -ComObject PowerPoint.Application
    $powerPoint.Visible = -1 # This Office build does not permit a hidden PowerPoint window.
    $presentation = $powerPoint.Presentations.Add()
    $slide = $presentation.Slides.Add(1, 12) # ppLayoutBlank

    $visio = New-Object -ComObject Visio.Application
    $visio.Visible = $true
    $visio.AlertResponse = 7 # Do not save prompts during cleanup.
    $stencil = $visio.Documents.Add('vssx')

    $usedNames = @{}
    foreach ($svg in $svgFiles) {
        $masterName = [System.IO.Path]::GetFileNameWithoutExtension($svg.Name)
        if ($usedNames.ContainsKey($masterName)) {
            $usedNames[$masterName]++
            $masterName = "$masterName ($($usedNames[$masterName]))"
        } else {
            $usedNames[$masterName] = 1
        }

        $emfPath = Join-Path $tempDirectory ($svg.BaseName + '.emf')
        $picture = $slide.Shapes.AddPicture($svg.FullName, 0, -1, 0, 0, -1, -1)
        $picture.Export($emfPath, 5) # ppShapeFormatEMF
        $picture.Delete()

        $master = $stencil.Masters.Add()
        $master.Name = $masterName
        $master.NameU = $masterName
        $null = $master.Import($emfPath)
    }

    # Visio COM can reject non-ASCII SaveAs paths. Save internally, then copy to the requested path.
    $stencil.SaveAs($temporaryStencilPath)
    $null = $stencil.Close()
    $stencil = $null
    Copy-Item -LiteralPath $temporaryStencilPath -Destination $StencilPath -Force

    [PSCustomObject]@{
        StencilPath = $StencilPath
        SvgCount = $svgFiles.Count
    }
} finally {
    if ($stencil -ne $null) {
        try { $null = $stencil.Close() } catch { }
    }
    if ($presentation -ne $null) {
        try { $presentation.Close() } catch { }
    }
    if ($powerPoint -ne $null) {
        try { $powerPoint.Quit() } catch { }
    }
    if ($visio -ne $null) {
        try { $visio.Quit() } catch { }
    }
    if ((Test-Path -LiteralPath $tempDirectory) -and -not $KeepIntermediate) {
        Remove-Item -LiteralPath $tempDirectory -Recurse -Force
    }
}
