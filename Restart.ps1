$ErrorActionPreference = 'Stop'
Write-Host 'Restarting the Fastener Order Bot...'
& (Join-Path $PSScriptRoot 'Stop.ps1')
& (Join-Path $PSScriptRoot 'Start.ps1')
