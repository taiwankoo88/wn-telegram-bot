$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot
Write-Host 'Telegram Fastener Bot - local setup'
Write-Host 'Create a NEW bot with @BotFather /newbot. Paste its token below (hidden).'
$botSecret = Read-Host 'Bot token' -AsSecureString
$botPtr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($botSecret)
try {
  $botToken = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($botPtr)
  if ($botToken -notmatch '^\d+:[A-Za-z0-9_-]{20,}$') { throw 'Invalid token format.' }
  try {
    $botInfo = Invoke-RestMethod -Uri ('https://api.telegram.org/bot' + $botToken + '/getMe') -Method Post
    $hookInfo = Invoke-RestMethod -Uri ('https://api.telegram.org/bot' + $botToken + '/getWebhookInfo') -Method Post
  } catch { throw 'Telegram connection failed. Check token and internet access.' }
  if ($hookInfo.result.url) { throw 'This bot already has an integration. Create a new bot in BotFather.' }
  $botSecret | ConvertFrom-SecureString | Set-Content -LiteralPath (Join-Path $PSScriptRoot 'token.dpapi') -NoNewline
} finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($botPtr); $botToken = $null }
$codeBytes = New-Object byte[] 8
$rng = [Security.Cryptography.RandomNumberGenerator]::Create()
$rng.GetBytes($codeBytes)
$rng.Dispose()
$pairCode = -join ($codeBytes | ForEach-Object { $_.ToString('x2') })
@{pairCode=$pairCode} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $PSScriptRoot 'config.json') -Encoding UTF8
# Pairing failure lock can be reset locally without removing an existing owner.
$statePath = Join-Path $PSScriptRoot 'state.json'
if (Test-Path -LiteralPath $statePath) {
  $state = Get-Content -Raw -LiteralPath $statePath | ConvertFrom-Json
  $state.pairFailures = 0
  $state | ConvertTo-Json | Set-Content -LiteralPath $statePath -Encoding UTF8
}
Write-Host ('Open Telegram: https://t.me/' + $botInfo.result.username)
Write-Host ('First administrator: send /pair ' + $pairCode)
Write-Host 'Existing administrator pairing is retained on reconfiguration.'
Write-Host 'Keep this PC awake and connected. Run Start.cmd after setup.'
Read-Host 'Press Enter to close'
