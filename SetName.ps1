# Shows the bot's display name and, when -Name is given, changes it through the Telegram Bot API.
# The token is decrypted in memory like Start.ps1 does and is never shown or written anywhere.
param([string]$Name)
$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot
$tokenPath = Join-Path $PSScriptRoot 'token.dpapi'
if (!(Test-Path -LiteralPath $tokenPath)) { throw 'Run Setup.cmd first.' }
$secret = ConvertTo-SecureString -String ((Get-Content -Raw -LiteralPath $tokenPath).Trim())
$ptr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secret)
try {
  $token = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr)
  $api = "https://api.telegram.org/bot$token"
  $current = Invoke-RestMethod -Uri "$api/getMyName" -Method Post
  Write-Host ('Current name: ' + $current.result.name)
  if ($Name) {
    $result = Invoke-RestMethod -Uri "$api/setMyName" -Method Post -ContentType 'application/json; charset=utf-8' -Body ([Text.Encoding]::UTF8.GetBytes((@{name = $Name} | ConvertTo-Json -Compress)))
    if (-not $result.ok) { throw 'Telegram rejected the new name.' }
    $after = Invoke-RestMethod -Uri "$api/getMyName" -Method Post
    Write-Host ('New name:     ' + $after.result.name)
  }
} catch {
  # Telegram error bodies never contain the token; the request URL is not echoed.
  $detail = $_.ErrorDetails.Message
  if ($detail) { Write-Host ('Telegram error: ' + $detail) } else { Write-Host ('Error: ' + $_.Exception.Message) }
  exit 1
} finally {
  [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr); $token = $null; $api = $null
}
