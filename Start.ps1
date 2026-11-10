$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot
$runningBot = Join-Path $PSScriptRoot 'bot.mjs'
if (Get-CimInstance Win32_Process -Filter "Name = 'node.exe'" | Where-Object { $_.CommandLine -and $_.CommandLine.Contains($runningBot) }) {
  Write-Host 'Bot already running.'
  exit 0
}
$tokenPath = Join-Path $PSScriptRoot 'token.dpapi'
if (!(Test-Path -LiteralPath $tokenPath)) { throw 'Run Setup.cmd first.' }
# Set-Content adds a trailing newline. DPAPI's serialized string must be trimmed.
$encryptedToken = (Get-Content -Raw -LiteralPath $tokenPath).Trim()
if ([string]::IsNullOrWhiteSpace($encryptedToken)) { throw 'Saved token is empty. Run Setup.cmd again.' }
$secret = ConvertTo-SecureString -String $encryptedToken
$ptr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secret)
try {
  $env:TELEGRAM_BOT_TOKEN = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr)
  $node = (Get-Command node -CommandType Application -ErrorAction Stop).Source
  $bot = Join-Path $PSScriptRoot 'bot.mjs'
  $botProcess = Start-Process -FilePath $node -ArgumentList ('"' + $bot + '"') -WorkingDirectory $PSScriptRoot -WindowStyle Hidden -RedirectStandardOutput (Join-Path $PSScriptRoot 'bot.log') -RedirectStandardError (Join-Path $PSScriptRoot 'error.log') -PassThru
  Write-Host 'Connecting to Telegram. Please wait...'
  $online = $false
  for ($attempt = 0; $attempt -lt 15; $attempt++) {
    Start-Sleep -Seconds 1
    $botProcess.Refresh()
    if ($botProcess.HasExited) { throw 'Bot stopped before connecting. Please share this window with support; error.log contains the details.' }
    $statusPath = Join-Path $PSScriptRoot 'status.json'
    if (Test-Path -LiteralPath $statusPath) {
      try {
        $botStatus = Get-Content -Raw -LiteralPath $statusPath | ConvertFrom-Json
        if ($botStatus.pid -eq $botProcess.Id) { $online = $true; break }
      } catch { }
    }
  }
  if ($online) {
    Write-Host ('ONLINE: @' + $botStatus.username) -ForegroundColor Green
    Write-Host 'Now send /help in your Telegram bot chat. You may close this window.'
  } else {
    Write-Host 'NOT ONLINE YET: the process is running and retrying the Telegram connection.' -ForegroundColor Yellow
    Write-Host 'Please share this window with support. Do not repeat Setup or enter your token again.'
  }
} finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr); Remove-Item Env:TELEGRAM_BOT_TOKEN -ErrorAction SilentlyContinue }
