<#
.SYNOPSIS
Converts the CV HTML master page to DOCX and PDF.

.DESCRIPTION
Opens the source HTML in Microsoft Word, removes web-only download links for
export, applies optional style and margins, and writes both DOCX and PDF files.

.PARAMETER HtmlPath
Path to the source HTML file. Relative paths resolve from the script folder.

.PARAMETER OutputBaseName
Base filename used for generated outputs.

.PARAMETER OutputDirectory
Directory where generated DOCX and PDF files are written.

.PARAMETER QuickStyleSet
Word Quick Style Set name to apply (for example: Black & White (Classic)).

.PARAMETER MarginInches
Page margin in inches. Default is 0.5 (narrow).

.PARAMETER ExportFontName
Font name forced for export documents.

.EXAMPLE
.\Convert-CV.ps1

Generates Ben-Madle-Jordan-CV.docx and Ben-Madle-Jordan-CV.pdf from index.html.

.EXAMPLE
.\Convert-CV.ps1 -OutputBaseName "CV-2026" -MarginInches 0.75 -ExportFontName "Arial Nova"

.NOTES
Author: Ben Madle-Jordan
Created: 04/07/2026
Version: 1.0
#>
[CmdletBinding()]
param(
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$HtmlPath = "index.html",

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$OutputBaseName = "Ben-Madle-Jordan-CV",

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$OutputDirectory = ".",

    [Parameter()]
    [string]$QuickStyleSet = "Black & White (Classic)",

    [Parameter()]
    [ValidateRange(0, 5)]
    [double]$MarginInches = 0.5,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$ExportFontName = "Arial Nova"
)

#region Runtime Settings
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
#endregion

#region Helper Functions
function Resolve-CvPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$BaseDirectory
    )

    if ([System.IO.Path]::IsPathRooted($Path)) {
        return [System.IO.Path]::GetFullPath($Path)
    }

    return [System.IO.Path]::GetFullPath((Join-Path $BaseDirectory $Path))
}

function New-ExportHtmlContent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$HtmlContent,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$FontName
    )

    $processedHtml = [System.Text.RegularExpressions.Regex]::Replace(
        $HtmlContent,
        '<div class="cv-downloads"[\s\S]*?</div>',
        '',
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    )

    if ($processedHtml -match '</style>') {
        $styleOverride = "`n        body { font-family: '$FontName', Arial, sans-serif !important; }`n    </style>"
        $processedHtml = $processedHtml -replace '</style>', $styleOverride
    }

    return $processedHtml
}

function Set-DocumentMargins {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $WordApplication,

        [Parameter(Mandatory = $true)]
        $Document,

        [Parameter(Mandatory = $true)]
        [double]$MarginInches
    )

    if ($MarginInches -le 0) {
        return
    }

    $marginPoints = $WordApplication.InchesToPoints($MarginInches)
    $Document.PageSetup.TopMargin = $marginPoints
    $Document.PageSetup.BottomMargin = $marginPoints
    $Document.PageSetup.LeftMargin = $marginPoints
    $Document.PageSetup.RightMargin = $marginPoints
}
#endregion

#region Path Resolution
$scriptDirectory = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$resolvedHtmlPath = Resolve-CvPath -Path $HtmlPath -BaseDirectory $scriptDirectory
$resolvedOutputDirectory = Resolve-CvPath -Path $OutputDirectory -BaseDirectory $scriptDirectory

if (-not (Test-Path -Path $resolvedHtmlPath)) {
    throw "HTML file not found: $resolvedHtmlPath"
}

if (-not (Test-Path -Path $resolvedOutputDirectory)) {
    New-Item -ItemType Directory -Path $resolvedOutputDirectory | Out-Null
}

$docxPath = Join-Path $resolvedOutputDirectory ("{0}.docx" -f $OutputBaseName)
$pdfPath = Join-Path $resolvedOutputDirectory ("{0}.pdf" -f $OutputBaseName)
$tempHtmlPath = Join-Path $resolvedOutputDirectory ("{0}.export-temp.html" -f [System.IO.Path]::GetFileNameWithoutExtension($OutputBaseName))
#endregion

#region Conversion
$word = $null
$document = $null

try {
    $htmlContent = Get-Content -Path $resolvedHtmlPath -Raw
    $exportHtml = New-ExportHtmlContent -HtmlContent $htmlContent -FontName $ExportFontName
    Set-Content -Path $tempHtmlPath -Value $exportHtml -Encoding UTF8

    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    $document = $word.Documents.Open($tempHtmlPath)

    if (-not [string]::IsNullOrWhiteSpace($ExportFontName)) {
        $document.Content.Font.Name = $ExportFontName
    }

    if (-not [string]::IsNullOrWhiteSpace($QuickStyleSet)) {
        try {
            $document.ApplyQuickStyleSet($QuickStyleSet)
        }
        catch {
            Write-Warning "Could not apply Quick Style Set '$QuickStyleSet'. Continuing without it."
        }
    }

    Set-DocumentMargins -WordApplication $word -Document $document -MarginInches $MarginInches

    $wdFormatXMLDocument = 16
    $wdExportFormatPDF = 17

    $document.SaveAs2($docxPath, $wdFormatXMLDocument)
    $document.ExportAsFixedFormat($pdfPath, $wdExportFormatPDF)

    [PSCustomObject]@{
        DocxPath = $docxPath
        PdfPath  = $pdfPath
    }
}
finally {
    if ($document -ne $null) {
        $document.Close()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($document) | Out-Null
    }

    if ($word -ne $null) {
        $word.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }

    if (Test-Path -Path $tempHtmlPath) {
        Remove-Item -Path $tempHtmlPath -Force
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
#endregion
