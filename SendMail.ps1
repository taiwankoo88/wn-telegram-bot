# Sends one email with one attachment through the Outlook profile of the logged-in user.
# The job file holds to, subject, body and attachment; the last output line is JSON with the result.
# -Probe only connects to Outlook and reports the account; nothing is sent.
param([string]$JobFile, [switch]$Probe)
$ErrorActionPreference = 'Stop'
if (-not $Probe) {
  $job = Get-Content -Raw -LiteralPath $JobFile -Encoding UTF8 | ConvertFrom-Json
  if (-not (Test-Path -LiteralPath $job.attachment)) { throw "Attachment not found: $($job.attachment)" }
}
# Since 2026-10-07 Outlook started by COM activation (New-Object -ComObject) never registers with DCOM
# (0x80080005), while a normally started Outlook can still be attached to. So attach to the running Outlook,
# or start it minimized and attach once it has registered.
function Get-RunningOutlook { try { [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application') } catch { $null } }
# Office registers for attaching only after its window has been active and then lost focus (Microsoft KB 238610).
# An Outlook started minimized never has (2026-10-09), so activate its window briefly and minimize it again.
Add-Type -Namespace SendMail -Name Win -MemberDefinition '[DllImport("user32.dll")] public static extern bool ShowWindow(System.IntPtr h, int cmd);'
function Register-Outlook {
  $p = Get-Process OUTLOOK -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 } | Select-Object -First 1
  if (-not $p) { return }
  [void](New-Object -ComObject WScript.Shell).AppActivate($p.Id)
  Start-Sleep -Milliseconds 500
  [void][SendMail.Win]::ShowWindow($p.MainWindowHandle, 6)  # SW_MINIMIZE: Outlook loses focus and registers
}
$startedOutlook = -not (Get-Process OUTLOOK -ErrorAction SilentlyContinue)
$outlook = Get-RunningOutlook
if (-not $outlook) {
  if ($startedOutlook) {
    $exe = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE').'(default)'
    Start-Process -FilePath $exe -WindowStyle Minimized
  }
  for ($i = 0; $i -lt 60 -and -not $outlook; $i++) {
    Start-Sleep -Seconds 1
    $outlook = Get-RunningOutlook
    if (-not $outlook -and $i % 5 -eq 4) { Register-Outlook }
  }
  if (-not $outlook) { $outlook = New-Object -ComObject Outlook.Application }
}
try {
  if ($Probe) {
    $session = $outlook.GetNamespace('MAPI')
    @{status = 'probe'; account = $session.Accounts.Item(1).SmtpAddress; outbox = $session.GetDefaultFolder(4).Items.Count; startedOutlook = $startedOutlook} | ConvertTo-Json -Compress
    return
  }
  $session = $outlook.GetNamespace('MAPI')
  $outbox = $session.GetDefaultFolder(4)
  $waiting = $outbox.Items.Count
  $mail = $outlook.CreateItem(0)
  $mail.To = $job.to
  $mail.Subject = $job.subject
  $mail.Body = $job.body
  [void]$mail.Attachments.Add($job.attachment)
  $mail.Send()
  # Outlook started by automation may take a while to connect; the message is sent once it leaves the Outbox.
  $status = 'queued'
  for ($i = 0; $i -lt 90; $i++) {
    Start-Sleep -Seconds 1
    if ($outbox.Items.Count -le $waiting) { $status = 'sent'; break }
    if ($i % 15 -eq 10) { try { $session.SendAndReceive($false) } catch {} }
  }
  # Outlook keeps running while a message waits in the Outbox so that it still goes out.
  if ($status -eq 'sent' -and $startedOutlook) { $outlook.Quit() }
  @{status = $status} | ConvertTo-Json -Compress
} finally {
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($outlook)
}
