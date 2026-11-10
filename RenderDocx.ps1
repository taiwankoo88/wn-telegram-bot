# Renders Word files with the Word installed on this PC so the bot can check the page layout.
# Job file: {"files": ["...docx", ...], "pdfDir": "..."}. Prints one JSON line per file, then {"done":true}.
# For each file: page count, the body text of page 1, and every text box (its text and whether it overflows).
param([Parameter(Mandatory)][string]$JobFile)
$ErrorActionPreference = 'Stop'
$job = Get-Content -Raw -LiteralPath $JobFile -Encoding UTF8 | ConvertFrom-Json
# Values straight from ConvertFrom-Json are PSObject-wrapped; handed to a Word COM method they make Word hang.
$files = @($job.files | ForEach-Object { [string]$_ })
$pdfDir = [string]$job.pdfDir
New-Item -ItemType Directory -Force -Path $pdfDir | Out-Null
# A progress log next to the PDFs shows where a run stopped if Word ever hangs.
$log = Join-Path $pdfDir 'render.log'
function Step($text) { Add-Content -LiteralPath $log -Value ("{0:HH:mm:ss} {1}" -f (Get-Date), $text) }
Step "start pid $PID"
# Word works on copies in a short temporary folder: Windows PowerShell cannot handle paths beyond 260 characters.
$work = Join-Path $env:TEMP ("steel-render-" + $PID)
New-Item -ItemType Directory -Force -Path $work | Out-Null
$wordBefore = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$word = New-Object -ComObject Word.Application
# The Word this run started (any other Word window belongs to the user and is left alone).
$mine = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $wordBefore -notcontains $_.Id } | ForEach-Object { $_.Id })
Step "Word started (pid $($mine -join ','))"
try {
  $word.Visible = $false
  $word.DisplayAlerts = 0          # wdAlertsNone
  $word.AutomationSecurity = 3     # msoAutomationSecurityForceDisable: never run macros
  foreach ($file in $files) {
    # Word is given a copy, so the delivered file is never locked or touched.
    $copy = [string](Join-Path $work ([IO.Path]::GetFileName($file)))
    Copy-Item -LiteralPath $file -Destination $copy -Force
    Step "opening $([IO.Path]::GetFileName($file))"
    $doc = $word.Documents.Open($copy, $false, $true)   # ConfirmConversions=false, ReadOnly=true
    try {
      Step "opened"
      $doc.Repaginate()
      $pages = [int]$doc.ComputeStatistics(2)           # wdStatisticPages
      $page1End = if ($pages -gt 1) { $doc.GoTo(1, 1, 2).Start } else { $doc.Content.End }   # start of page 2
      $page1 = $doc.Range(0, $page1End).Text
      $boxes = @()
      foreach ($shape in $doc.Shapes) {
        $hasText = $false
        try { $hasText = [bool]$shape.TextFrame.HasText } catch { }
        if ($hasText) {
          $boxes += [pscustomobject]@{
            name = $shape.Name
            page = [int]$shape.Anchor.Information(3)     # wdActiveEndPageNumber
            overflowing = [bool]$shape.TextFrame.Overflowing
            text = $shape.TextFrame.TextRange.Text
          }
        }
      }
      Step "checked: $pages page(s), $($boxes.Count) text box(es)"
      # Export beside the copy (short path, plain string), then move the PDF to where the bot wants it.
      $tmpPdf = [string][IO.Path]::ChangeExtension($copy, '.pdf')
      $doc.ExportAsFixedFormat($tmpPdf, 17)               # wdExportFormatPDF
      $pdf = [string](Join-Path $pdfDir ([IO.Path]::GetFileName($tmpPdf)))
      Move-Item -LiteralPath $tmpPdf -Destination $pdf -Force
      Step "pdf written"
      [pscustomobject]@{file = $file; pdf = $pdf; pages = $pages; page1Text = $page1; textBoxes = $boxes} | ConvertTo-Json -Compress -Depth 4
    } finally {
      $doc.Close(0)                                       # wdDoNotSaveChanges
    }
  }
} finally {
  try { $word.Quit() } catch { }
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  # A hidden Word that ignores Quit would otherwise linger for good, holding the work folder open.
  foreach ($i in 1..10) {
    if (-not @(Get-Process -Id $mine -ErrorAction SilentlyContinue).Count) { break }
    Start-Sleep -Milliseconds 500
  }
  foreach ($leftover in @(Get-Process -Id $mine -ErrorAction SilentlyContinue)) { Stop-Process -Id $leftover.Id -Force -ErrorAction SilentlyContinue; Step "Word pid $($leftover.Id) had to be stopped" }
  Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
  Step "Word closed"
}
'{"done":true}'
