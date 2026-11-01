param(
    [Parameter(Mandatory = $true)][string]$InputDocx,
    [Parameter(Mandatory = $true)][string]$OutputPdf,
    [switch]$UpdateFields
)
$ErrorActionPreference = 'Stop'
$taskDocx = (Resolve-Path -LiteralPath $InputDocx).Path
$taskPdf = [System.IO.Path]::GetFullPath($OutputPdf)
$taskWord = $null
$taskDocument = $null
try {
    $taskWord = New-Object -ComObject Word.Application
    $taskWord.Visible = $false
    $taskWord.DisplayAlerts = 0
    $taskDocument = $taskWord.Documents.Open($taskDocx, $false, (-not $UpdateFields.IsPresent), $false)
    if ($UpdateFields) {
        $taskDocument.Fields.Update() | Out-Null
        foreach ($taskToc in $taskDocument.TablesOfContents) { $taskToc.Update() }
        $taskDocument.Repaginate()
        foreach ($taskToc in $taskDocument.TablesOfContents) { $taskToc.UpdatePageNumbers() }
        $taskDocument.Save()
    }
    $taskDocument.ExportAsFixedFormat($taskPdf, 17)
    Write-Output ('Rendered pages: ' + $taskDocument.ComputeStatistics(2))
} finally {
    if ($null -ne $taskDocument) {
        $taskDocument.Close(0)
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($taskDocument) | Out-Null
    }
    if ($null -ne $taskWord) {
        $taskWord.Quit(0)
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($taskWord) | Out-Null
    }
}
