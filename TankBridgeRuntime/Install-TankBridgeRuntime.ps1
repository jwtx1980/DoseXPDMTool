param(
    [string]$InstallDir = (Join-Path $env:ProgramFiles "TG51TankBridge"),
    [switch]$Launch
)

$ErrorActionPreference = "Stop"

function Write-Step([string]$message) {
    Write-Host
    Write-Host $message -ForegroundColor Cyan
}

function Require-Administrator {
    $principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw "Run this installer as Administrator."
    }
}

function Stop-InstalledProcesses([string]$Root) {
    $fullRoot = [System.IO.Path]::GetFullPath($Root).TrimEnd("\")
    $processes = @(Get-CimInstance Win32_Process -ErrorAction SilentlyContinue |
        Where-Object {
            $_.ExecutablePath -and
            ([System.IO.Path]::GetFullPath($_.ExecutablePath).StartsWith($fullRoot, [System.StringComparison]::OrdinalIgnoreCase))
        })

    foreach ($process in $processes) {
        Write-Host ("Stopping running process {0} ({1})" -f $process.Name, $process.ProcessId)
        Invoke-CimMethod -InputObject $process -MethodName Terminate | Out-Null
    }
}

function Copy-Package([string]$Source, [string]$Destination) {
    $fullSource = [System.IO.Path]::GetFullPath($Source).TrimEnd("\")
    $fullDestination = [System.IO.Path]::GetFullPath($Destination).TrimEnd("\")

    if ($fullSource.Equals($fullDestination, [System.StringComparison]::OrdinalIgnoreCase)) {
        Write-Host "Source is already the install directory; skipping file copy."
        return
    }

    New-Item -ItemType Directory -Path $fullDestination -Force | Out-Null

    $args = @(
        $fullSource,
        $fullDestination,
        "/MIR",
        "/XD", "measurement-results", "queue-runs", "diagnostics",
        "/XF", "*.log", "tank-network*.json",
        "/R:2",
        "/W:1",
        "/NFL",
        "/NDL",
        "/NP"
    )
    & robocopy @args | Out-Host
    if ($LASTEXITCODE -gt 7) {
        throw "Robocopy failed with exit code $LASTEXITCODE."
    }
}

function Install-FirewallRule([string]$BridgeExe) {
    if (-not (Test-Path -LiteralPath $BridgeExe)) {
        throw "Bridge executable was not found at $BridgeExe."
    }

    $ruleName = "TG-51 Tank Bridge Inbound"
    Get-NetFirewallRule -DisplayName $ruleName -ErrorAction SilentlyContinue | Remove-NetFirewallRule

    New-NetFirewallRule `
        -DisplayName $ruleName `
        -Direction Inbound `
        -Action Allow `
        -Program $BridgeExe `
        -Protocol TCP `
        -RemoteAddress 169.254.0.0/16 `
        -Profile Any | Out-Null

    Write-Host "Firewall rule installed:"
    Write-Host "  $ruleName"
    Write-Host "  Program: $BridgeExe"
    Write-Host "  Remote: 169.254.0.0/16"
}

function New-Shortcut([string]$ShortcutPath, [string]$TargetPath, [string]$WorkingDirectory) {
    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($ShortcutPath)
    $shortcut.TargetPath = $TargetPath
    $shortcut.WorkingDirectory = $WorkingDirectory
    $shortcut.Description = "TG-51 Tank Bridge Runtime"
    $shortcut.Save()
}

function Install-Shortcuts([string]$Root) {
    $commonPrograms = [Environment]::GetFolderPath("CommonPrograms")
    $startMenuFolder = Join-Path $commonPrograms "TG-51 Tank Bridge"
    New-Item -ItemType Directory -Path $startMenuFolder -Force | Out-Null

    New-Shortcut `
        -ShortcutPath (Join-Path $startMenuFolder "Start TG-51 Tank Bridge.lnk") `
        -TargetPath (Join-Path $Root "Start-TankBridgeRuntime.cmd") `
        -WorkingDirectory $Root
}

Require-Administrator

$sourceRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$sourceRoot = [System.IO.Path]::GetFullPath($sourceRoot)
$InstallDir = [System.IO.Path]::GetFullPath($InstallDir)
$logPath = Join-Path $env:TEMP ("TG51TankBridge-install-{0:yyyyMMdd-HHmmss}.log" -f (Get-Date))

Start-Transcript -Path $logPath -Force | Out-Null
try {
    Write-Host "TG-51 Tank Bridge Runtime installer"
    Write-Host "Source:  $sourceRoot"
    Write-Host "Target:  $InstallDir"
    Write-Host "Log:     $logPath"

    $sourceBridge = Join-Path $sourceRoot "bridge\TankControllerBridge.exe"
    if (-not (Test-Path -LiteralPath $sourceBridge)) {
        throw "This installer must be run from the unzipped package folder containing bridge\TankControllerBridge.exe."
    }

    Write-Step "Stopping prior installed bridge processes"
    Stop-InstalledProcesses -Root $InstallDir

    Write-Step "Copying package to Program Files"
    Copy-Package -Source $sourceRoot -Destination $InstallDir

    $bridge = Join-Path $InstallDir "bridge\TankControllerBridge.exe"

    Write-Step "Installing firewall rule"
    Install-FirewallRule -BridgeExe $bridge

    Write-Step "Creating Start Menu shortcut"
    Install-Shortcuts -Root $InstallDir

    Write-Step "Verifying installed files"
    foreach ($required in @($bridge, (Join-Path $InstallDir "Start-TankBridgeRuntime.cmd"), (Join-Path $InstallDir "wwwroot\index.html"), (Join-Path $InstallDir "Test-TankNetworkAssumptions.ps1"))) {
        if (-not (Test-Path -LiteralPath $required)) {
            throw "Missing installed file: $required"
        }
        Write-Host "OK: $required"
    }

    Write-Step "Install complete"
    Write-Host "Installed bridge executable:"
    Write-Host "  $bridge"
    Write-Host
    Write-Host "TG-51 app launch command:"
    Write-Host "  $InstallDir\Start-TankBridgeRuntime.cmd"
    Write-Host
    Write-Host "Network diagnostic:"
    Write-Host "  powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$InstallDir\Test-TankNetworkAssumptions.ps1`" -LiveBridgeProbe -OutputJson `"$env:USERPROFILE\Desktop\tank-network-live.json`""

    if ($Launch) {
        Write-Step "Launching bridge"
        Start-Process -FilePath (Join-Path $InstallDir "Start-TankBridgeRuntime.cmd") -WorkingDirectory $InstallDir | Out-Null
    }
}
finally {
    Stop-Transcript | Out-Null
    Write-Host
    Write-Host "Install log: $logPath"
}
