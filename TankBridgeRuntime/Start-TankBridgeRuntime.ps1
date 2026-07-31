param(
    [string]$Url = "http://127.0.0.1:8770"
)

$ErrorActionPreference = "Stop"

$packageRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$bridgeExe = Join-Path $packageRoot "bridge\TankControllerBridge.exe"

if (-not (Test-Path -LiteralPath $bridgeExe)) {
    throw "Bridge executable was not found at $bridgeExe."
}

try {
    Invoke-RestMethod -Uri "$Url/health" -TimeoutSec 2 | Out-Null
    Write-Host "TG-51 Tank Bridge is already running at $Url."
    exit 0
}
catch {
}

$dataRoot = $env:TANK_BRIDGE_DATA_DIR
if ([string]::IsNullOrWhiteSpace($dataRoot)) {
    $dataRoot = Join-Path ([Environment]::GetFolderPath("LocalApplicationData")) "TG51TankBridge"
}
$logRoot = Join-Path $dataRoot "logs"
New-Item -ItemType Directory -Path $logRoot -Force | Out-Null
$out = Join-Path $logRoot "tank-bridge.out.log"
$err = Join-Path $logRoot "tank-bridge.err.log"

Start-Process `
    -FilePath $bridgeExe `
    -ArgumentList @("--urls", $Url) `
    -WorkingDirectory $packageRoot `
    -RedirectStandardOutput $out `
    -RedirectStandardError $err `
    -WindowStyle Hidden | Out-Null

$deadline = DateTimeOffset.Now.AddSeconds(15)
while (DateTimeOffset.Now -lt $deadline) {
    try {
        Invoke-RestMethod -Uri "$Url/health" -TimeoutSec 2 | Out-Null
        Write-Host "TG-51 Tank Bridge started at $Url."
        exit 0
    }
    catch {
        Start-Sleep -Milliseconds 300
    }
}

throw "TG-51 Tank Bridge did not answer health checks at $Url. See $out and $err."
