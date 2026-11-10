$botPath = Join-Path $PSScriptRoot 'bot.mjs'
Get-CimInstance Win32_Process -Filter "Name = 'node.exe'" | Where-Object { $_.CommandLine -and $_.CommandLine.Contains($botPath) } | ForEach-Object { Stop-Process -Id $_.ProcessId }
Write-Host 'Stopped this Telegram bot.'
