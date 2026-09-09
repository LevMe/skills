param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputDirectory
)

$ErrorActionPreference = "Stop"
$word = $null
$document = $null
$probe = $null

try {
    if ($PSVersionTable.PSEdition -ne "Desktop") {
        throw "Run this script with Windows PowerShell 5.1, not PowerShell 7."
    }

    $resolvedInput = (Resolve-Path -LiteralPath $InputPath).Path
    $resolvedOutput = [System.IO.Path]::GetFullPath($OutputDirectory)
    New-Item -ItemType Directory -Path $resolvedOutput -Force | Out-Null

    $stem = [System.IO.Path]::GetFileNameWithoutExtension($resolvedInput)
    $pdfPath = Join-Path $resolvedOutput ($stem + ".pdf")

    # Probe Word before opening the input document.
    $word = New-Object -ComObject Word.Application
    if ($null -eq $word) {
        throw "Microsoft Word COM could not be created."
    }
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $probe = $word.Documents.Add()
    if ($null -eq $probe) {
        throw "Microsoft Word could not create a temporary document."
    }
    $probe.Close()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($probe) | Out-Null
    $probe = $null

    # Open the input read-only, update fields, and export without saving it.
    $document = $word.Documents.Open($resolvedInput, $false, $true, $false)
    if ($null -eq $document) {
        throw "Microsoft Word could not open the input DOCX."
    }
    $document.Fields.Update() | Out-Null
    for ($index = 1; $index -le $document.TablesOfContents.Count; $index++) {
        $tableOfContents = $document.TablesOfContents.Item($index)
        $tableOfContents.Update() | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($tableOfContents) | Out-Null
    }
    $document.Repaginate()
    $document.ExportAsFixedFormat($pdfPath, 17)

    $popplerCommand = Get-Command pdftoppm.exe -ErrorAction SilentlyContinue
    $popplerPath = $null
    if ($null -ne $popplerCommand) {
        $popplerPath = $popplerCommand.Path
    }
    if ([string]::IsNullOrWhiteSpace($popplerPath)) {
        $runtimeCandidate = Join-Path ([Environment]::GetFolderPath("UserProfile")) ".cache\codex-runtimes\codex-primary-runtime\dependencies\native\poppler\Library\bin\pdftoppm.exe"
        if (Test-Path -LiteralPath $runtimeCandidate) {
            $popplerPath = $runtimeCandidate
        }
    }
    if ([string]::IsNullOrWhiteSpace($popplerPath)) {
        throw "Poppler pdftoppm.exe was not found; the PDF was exported but page images were not created."
    }

    $imagePrefix = Join-Path $resolvedOutput ($stem + "-page")
    & $popplerPath -png -r 150 $pdfPath $imagePrefix
    if ($LASTEXITCODE -ne 0) {
        throw ("Poppler failed with exit code " + $LASTEXITCODE + ".")
    }
    $pngPattern = $stem + "-page-*.png"
    $pngFiles = @(Get-ChildItem -LiteralPath $resolvedOutput -Filter $pngPattern -File)
    if ($pngFiles.Count -eq 0) {
        throw "Poppler returned success but no page images were created."
    }

    Write-Output ("PDF=" + $pdfPath)
    Write-Output ("PNG_COUNT=" + $pngFiles.Count)
}
catch {
    throw ("Word DOCX render failed: " + $_.Exception.Message)
}
finally {
    if ($null -ne $probe) {
        try { $probe.Close() } catch { }
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($probe) | Out-Null } catch { }
    }
    if ($null -ne $document) {
        try { $document.Close() } catch { }
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($document) | Out-Null } catch { }
    }
    if ($null -ne $word) {
        try { $word.Quit() } catch { }
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null } catch { }
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
