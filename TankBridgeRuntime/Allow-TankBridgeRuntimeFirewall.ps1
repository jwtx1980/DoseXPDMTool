param(
    [string]$BridgeExe
)

$ErrorActionPreference = "Stop"

$packageRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($BridgeExe)) {
    $BridgeExe = Join-Path $packageRoot "bridge\TankControllerBridge.exe"
}

if (-not (Test-Path -LiteralPath $BridgeExe)) {
    throw "Bridge executable was not found at $BridgeExe. Run this script from the unzipped TG-51 Tank Bridge Runtime package folder or pass -BridgeExe."
}

$principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw "Run this script as Administrator."
}

$ruleName = "TG-51 Tank Bridge Inbound"
$existing = Get-NetFirewallRule -DisplayName $ruleName -ErrorAction SilentlyContinue
if ($existing) {
    $existing | Remove-NetFirewallRule
}

New-NetFirewallRule `
    -DisplayName $ruleName `
    -Direction Inbound `
    -Action Allow `
    -Program $BridgeExe `
    -Protocol TCP `
    -RemoteAddress 169.254.0.0/16 `
    -Profile Any | Out-Null

Write-Host "Added Windows Defender Firewall inbound rule:"
Write-Host "  $ruleName"
Write-Host "  Program: $BridgeExe"
Write-Host "  Remote: 169.254.0.0/16"
Write-Host
Write-Host "Restart the TG-51 Tank Bridge, then connect again."
