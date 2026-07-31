param(
    [int]$DiscoverySeconds = 8,
    [int]$DiscoveryPort = 1221,
    [string]$FallbackCcuAddress = "169.254.40.1",
    [int]$FallbackCcuPort = 1222,
    [int]$ExpectedCallbackRemotePort = 1227,
    [string]$PackageRoot,
    [switch]$SkipTcpProbe,
    [switch]$LiveBridgeProbe,
    [string]$OutputJson
)

$ErrorActionPreference = "Stop"

$results = New-Object System.Collections.Generic.List[object]

function Add-Check {
    param(
        [string]$Name,
        [ValidateSet("PASS", "WARN", "FAIL", "INFO")]
        [string]$Status,
        [string]$Details,
        $Data = $null
    )

    $item = [pscustomobject]@{
        Time = (Get-Date).ToString("s")
        Name = $Name
        Status = $Status
        Details = $Details
        Data = $Data
    }
    $script:results.Add($item) | Out-Null
    $color = switch ($Status) {
        "PASS" { "Green" }
        "WARN" { "Yellow" }
        "FAIL" { "Red" }
        default { "Cyan" }
    }
    Write-Host ("[{0}] {1}: {2}" -f $Status, $Name, $Details) -ForegroundColor $color
}

function Get-PrintableAscii {
    param([byte[]]$Bytes)

    $chars = foreach ($b in $Bytes) {
        if ($b -ge 32 -and $b -le 126) {
            [char]$b
        }
        else {
            "."
        }
    }
    -join $chars
}

function Get-RouteLocalAddress {
    param(
        [string]$RemoteAddress,
        [int]$RemotePort
    )

    $socket = $null
    try {
        $socket = New-Object System.Net.Sockets.Socket(
            [System.Net.Sockets.AddressFamily]::InterNetwork,
            [System.Net.Sockets.SocketType]::Dgram,
            [System.Net.Sockets.ProtocolType]::Udp)
        $socket.Connect($RemoteAddress, $RemotePort)
        return $socket.LocalEndPoint.Address.ToString()
    }
    finally {
        if ($socket -ne $null) {
            $socket.Dispose()
        }
    }
}

function Get-AppPreferredLocalAddress {
    param(
        [string]$CcuAddress,
        [int]$CcuPort,
        $Addresses
    )

    if ($env:TANK_LOCAL_ADDRESS) {
        $configured = [System.Net.IPAddress]::Parse($env:TANK_LOCAL_ADDRESS)
        $available = @($Addresses | Where-Object { $_.IPAddress -eq $configured.ToString() -and $_.AddressState -eq "Preferred" })
        if ($available.Count -eq 0) {
            throw "TANK_LOCAL_ADDRESS $configured is not assigned as a Preferred IPv4 address on this computer."
        }

        return [pscustomobject]@{
            Address = $configured.ToString()
            Source = "TANK_LOCAL_ADDRESS"
            Note = "Environment override."
        }
    }

    $ccu = [System.Net.IPAddress]::Parse($CcuAddress)
    $ccuBytes = $ccu.GetAddressBytes()
    $preferred = @($Addresses | Where-Object { $_.AddressState -eq "Preferred" })
    $same169ClassC = @($preferred | Where-Object {
            $bytes = [System.Net.IPAddress]::Parse($_.IPAddress).GetAddressBytes()
            $bytes[0] -eq 169 -and $bytes[1] -eq 254 -and
            $ccuBytes[0] -eq 169 -and $ccuBytes[1] -eq 254 -and
            $bytes[2] -eq $ccuBytes[2]
        })
    if ($same169ClassC.Count -gt 0) {
        return [pscustomobject]@{
            Address = $same169ClassC[0].IPAddress
            Source = "same 169.254.x class-C"
            Note = "This is the app's first automatic preference."
        }
    }

    $routed = Get-RouteLocalAddress -RemoteAddress $CcuAddress -RemotePort $CcuPort
    if ($routed) {
        return [pscustomobject]@{
            Address = $routed
            Source = "Windows route"
            Note = "No preferred 169.254.x match was found, so the app would use the routed local address."
        }
    }

    $originalLab = @($preferred | Where-Object { $_.IPAddress -eq "169.254.104.137" })
    if ($originalLab.Count -gt 0) {
        return [pscustomobject]@{
            Address = "169.254.104.137"
            Source = "original lab fallback"
            Note = "Preferred original lab address is assigned."
        }
    }

    $linkLocal = @($preferred | Where-Object { $_.IPAddress -like "169.254.*" })
    if ($linkLocal.Count -gt 0) {
        return [pscustomobject]@{
            Address = $linkLocal[0].IPAddress
            Source = "first preferred 169.254.x"
            Note = "Fallback link-local selection."
        }
    }

    throw "No Preferred 169.254.x.x tank adapter address is assigned. Tentative or Deprecated addresses do not count."
}

function Receive-CcuDiscovery {
    param(
        [int]$Port,
        [int]$Seconds
    )

    $udp = $null
    try {
        $udp = New-Object System.Net.Sockets.UdpClient([System.Net.Sockets.AddressFamily]::InterNetwork)
        $udp.Client.SetSocketOption(
            [System.Net.Sockets.SocketOptionLevel]::Socket,
            [System.Net.Sockets.SocketOptionName]::ReuseAddress,
            $true)
        $udp.Client.Bind([System.Net.IPEndPoint]::new([System.Net.IPAddress]::Any, $Port))

        $deadline = (Get-Date).AddSeconds($Seconds)
        while ((Get-Date) -lt $deadline) {
            $remaining = [Math]::Max(250, [int]($deadline - (Get-Date)).TotalMilliseconds)
            $task = $udp.ReceiveAsync()
            if ($task.Wait([Math]::Min($remaining, 1000))) {
                $packet = $task.Result
                $ascii = Get-PrintableAscii $packet.Buffer
                $ipMatches = [regex]::Matches($ascii, "\b\d{1,3}(?:\.\d{1,3}){3}\b")
                $portMatches = [regex]::Matches($ascii, "\b\d{3,5}\b")
                $candidateIp = if ($ipMatches.Count -gt 0) { $ipMatches[$ipMatches.Count - 1].Value } else { $packet.RemoteEndPoint.Address.ToString() }
                $candidatePort = $null
                foreach ($m in $portMatches) {
                    $value = [int]$m.Value
                    if ($value -ge 1024 -and $value -le 65535) {
                        $candidatePort = $value
                    }
                }
                if ($candidatePort -eq $null) {
                    $candidatePort = $script:FallbackCcuPort
                }

                return [pscustomobject]@{
                    RemoteEndpoint = $packet.RemoteEndPoint.ToString()
                    Bytes = $packet.Buffer.Length
                    Printable = $ascii
                    CandidateAddress = $candidateIp
                    CandidatePort = $candidatePort
                }
            }
        }

        return $null
    }
    finally {
        if ($udp -ne $null) {
            $udp.Close()
            $udp.Dispose()
        }
    }
}

function Test-TcpControl {
    param(
        [string]$RemoteAddress,
        [int]$RemotePort,
        [string]$LocalAddress,
        [int]$TimeoutMs = 5000
    )

    $client = $null
    try {
        $client = New-Object System.Net.Sockets.TcpClient([System.Net.Sockets.AddressFamily]::InterNetwork)
        if (-not [string]::IsNullOrWhiteSpace($LocalAddress)) {
            $client.Client.Bind([System.Net.IPEndPoint]::new([System.Net.IPAddress]::Parse($LocalAddress), 0))
        }
        $task = $client.ConnectAsync($RemoteAddress, $RemotePort)
        if (-not $task.Wait($TimeoutMs)) {
            throw "TCP connect timed out after $TimeoutMs ms."
        }
        if ($task.IsFaulted) {
            throw $task.Exception.GetBaseException()
        }

        return [pscustomobject]@{
            LocalEndpoint = $client.Client.LocalEndPoint.ToString()
            RemoteEndpoint = $client.Client.RemoteEndPoint.ToString()
        }
    }
    finally {
        if ($client -ne $null) {
            $client.Close()
            $client.Dispose()
        }
    }
}

function Get-FreeAdjacentPortPair {
    param([string]$LocalAddress)

    for ($port = 59150; $port -lt 62000; $port += 2) {
        $control = $null
        $callback = $null
        try {
            $ip = [System.Net.IPAddress]::Parse($LocalAddress)
            $control = [System.Net.Sockets.TcpListener]::new($ip, $port)
            $callback = [System.Net.Sockets.TcpListener]::new($ip, $port + 1)
            $control.Start()
            $callback.Start()
            return [pscustomobject]@{
                ControlPort = $port
                CallbackPort = $port + 1
            }
        }
        catch {
        }
        finally {
            if ($control -ne $null) {
                $control.Stop()
            }
            if ($callback -ne $null) {
                $callback.Stop()
            }
        }
    }

    throw "No adjacent local control/callback port pair was available."
}

function Test-FirewallRule {
    param([string]$BridgeExe)

    if (-not (Get-Command Get-NetFirewallRule -ErrorAction SilentlyContinue)) {
        return [pscustomobject]@{
            Available = $false
            MatchingRules = @()
            Message = "NetSecurity cmdlets are not available in this PowerShell session."
        }
    }

    if (-not (Test-Path -LiteralPath $BridgeExe)) {
        return [pscustomobject]@{
            Available = $true
            MatchingRules = @()
            Message = "Bridge executable was not found at $BridgeExe."
        }
    }

    $rules = @(Get-NetFirewallRule -Direction Inbound -Action Allow -Enabled True -ErrorAction SilentlyContinue |
        ForEach-Object {
            $rule = $_
            $apps = @(Get-NetFirewallApplicationFilter -AssociatedNetFirewallRule $rule -ErrorAction SilentlyContinue)
            foreach ($app in $apps) {
                if ($app.Program -and ([System.IO.Path]::GetFullPath($app.Program) -ieq [System.IO.Path]::GetFullPath($BridgeExe))) {
                    [pscustomobject]@{
                        DisplayName = $rule.DisplayName
                        Profile = $rule.Profile.ToString()
                        Program = $app.Program
                    }
                }
            }
        })

    [pscustomobject]@{
        Available = $true
        MatchingRules = $rules
        Message = if ($rules.Count -gt 0) { "Found inbound allow rule for bridge executable." } else { "No inbound allow rule found for bridge executable." }
    }
}

function Resolve-PackageRoot {
    param([string]$RequestedRoot)

    if (-not [string]::IsNullOrWhiteSpace($RequestedRoot)) {
        return [System.IO.Path]::GetFullPath($RequestedRoot)
    }

    $scriptDir = Split-Path -Parent $PSCommandPath
    $parent = Split-Path -Parent $scriptDir
    $homeDir = [Environment]::GetFolderPath("UserProfile")
    $candidates = @(
        $scriptDir,
        (Join-Path $scriptDir "TankBridgeRuntime"),
        (Join-Path $scriptDir "LocalOPABQueueRunner"),
        (Join-Path $scriptDir "dist\TankBridgeRuntime"),
        (Join-Path $scriptDir "dist\LocalOPABQueueRunner"),
        (Join-Path $parent "TankBridgeRuntime"),
        (Join-Path $parent "LocalOPABQueueRunner"),
        (Join-Path $parent "dist\TankBridgeRuntime"),
        (Join-Path $parent "dist\LocalOPABQueueRunner"),
        (Join-Path $homeDir "Downloads\TankBridgeRuntime"),
        (Join-Path $homeDir "Downloads\TG51TankBridgeRuntime"),
        (Join-Path $homeDir "Downloads\LocalOPABQueueRunner"),
        (Join-Path $homeDir "Downloads\Local OPAB Queue Runner"),
        (Join-Path $homeDir "Desktop\TankBridgeRuntime"),
        (Join-Path $homeDir "Desktop\TG51TankBridgeRuntime"),
        (Join-Path $homeDir "Desktop\LocalOPABQueueRunner"),
        (Join-Path $homeDir "Desktop\Local OPAB Queue Runner")
    )

    foreach ($candidate in $candidates) {
        $full = [System.IO.Path]::GetFullPath($candidate)
        if (Test-Path -LiteralPath (Join-Path $full "bridge\TankControllerBridge.exe")) {
            return $full
        }
    }

    return [System.IO.Path]::GetFullPath((Join-Path $parent "dist\TankBridgeRuntime"))
}

function Invoke-LiveBridgeProbe {
    param(
        [string]$Root,
        [int]$Port = 18819
    )

    $bridgeExe = Join-Path $Root "bridge\TankControllerBridge.exe"
    if (-not (Test-Path -LiteralPath $bridgeExe)) {
        throw "Packaged bridge was not found at $bridgeExe."
    }

    $base = "http://127.0.0.1:$Port"
    $out = Join-Path $env:TEMP "tank-network-live-probe-$Port.out.log"
    $err = Join-Path $env:TEMP "tank-network-live-probe-$Port.err.log"
    Remove-Item $out, $err -ErrorAction SilentlyContinue
    $process = Start-Process -FilePath $bridgeExe -ArgumentList @("--urls", $base) -WorkingDirectory $Root -RedirectStandardOutput $out -RedirectStandardError $err -WindowStyle Hidden -PassThru
    $connected = $false

    try {
        $ready = $false
        for ($i = 0; $i -lt 40; $i++) {
            try {
                Invoke-RestMethod "$base/health" -TimeoutSec 2 | Out-Null
                $ready = $true
                break
            }
            catch {
                Start-Sleep -Milliseconds 500
            }
        }
        if (-not $ready) {
            throw "Bridge did not answer health at $base."
        }

        $connect = Invoke-RestMethod "$base/api/connect" -Method Post -TimeoutSec 45
        $connected = $true
        $state = $connect
        for ($i = 0; $i -lt 20; $i++) {
            Start-Sleep -Milliseconds 750
            $state = Invoke-RestMethod "$base/api/state" -TimeoutSec 5
            if (($state.callbackConnectionCount -gt 0 -or $state.callbackPeerEndpoint) -and $state.latestStatus -and $state.samples.Count -gt 0) {
                break
            }
        }

        $disconnect = Invoke-RestMethod "$base/api/disconnect" -Method Post -TimeoutSec 30
        $connected = $false

        return [pscustomobject]@{
            Connected = [bool]$connect.connected
            LocalEndpoint = $state.localEndpoint
            CallbackEndpoint = $state.callbackEndpoint
            CallbackPeerEndpoint = $state.callbackPeerEndpoint
            CallbackConnectionCount = $state.callbackConnectionCount
            LastCallbackAt = $state.lastCallbackAt
            LatestStatus = $state.latestStatus
            SampleCount = $state.samples.Count
            DisconnectLogTail = @($disconnect.logs | Select-Object -Last 6)
        }
    }
    finally {
        if ($connected) {
            try {
                Invoke-RestMethod "$base/api/disconnect" -Method Post -TimeoutSec 30 | Out-Null
            }
            catch {
            }
        }
        if ($process -and -not $process.HasExited) {
            Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
            $process.WaitForExit()
        }
    }
}

Write-Host
Write-Host "Tank network assumptions diagnostic" -ForegroundColor Cyan
Write-Host "Passive checks run first. Live bridge probe is only run with -LiveBridgeProbe." -ForegroundColor Cyan
Write-Host

$PackageRoot = Resolve-PackageRoot -RequestedRoot $PackageRoot
Add-Check "Package root" "INFO" "Using package root $PackageRoot." ([pscustomobject]@{ PackageRoot = $PackageRoot })

try {
    $adapters = @(Get-NetAdapter -ErrorAction Stop | Where-Object { $_.Status -eq "Up" } |
        Select-Object Name, InterfaceDescription, MacAddress, LinkSpeed, InterfaceIndex)
    Add-Check "Active adapters" "INFO" ("Found {0} active adapter(s)." -f $adapters.Count) $adapters
}
catch {
    Add-Check "Active adapters" "WARN" $_.Exception.Message
}

try {
    $addresses = @(Get-NetIPAddress -AddressFamily IPv4 -ErrorAction Stop |
        Where-Object { $_.IPAddress -notlike "127.*" } |
        Select-Object IPAddress, PrefixLength, InterfaceAlias, InterfaceIndex, AddressState)
    $linkLocal = @($addresses | Where-Object { $_.IPAddress -like "169.254.*" })
    $preferredLinkLocal = @($linkLocal | Where-Object { $_.AddressState -eq "Preferred" })
    if ($linkLocal.Count -gt 0) {
        $status = if ($preferredLinkLocal.Count -gt 0) { "PASS" } else { "WARN" }
        Add-Check "Link-local tank adapter address" $status ("Found {0} 169.254.x.x address(es); {1} are Preferred." -f $linkLocal.Count, $preferredLinkLocal.Count) $linkLocal
    }
    else {
        Add-Check "Link-local tank adapter address" "WARN" "No 169.254.x.x adapter address found. The CCU link usually needs one." $addresses
    }
}
catch {
    Add-Check "Link-local tank adapter address" "WARN" $_.Exception.Message
}

$discovery = $null
try {
    $discovery = Receive-CcuDiscovery -Port $DiscoveryPort -Seconds $DiscoverySeconds
    if ($discovery) {
        Add-Check "UDP CCU discovery" "PASS" ("Received {0} byte announcement from {1}; candidate endpoint {2}:{3}." -f $discovery.Bytes, $discovery.RemoteEndpoint, $discovery.CandidateAddress, $discovery.CandidatePort) $discovery
    }
    else {
        Add-Check "UDP CCU discovery" "WARN" ("No UDP announcement received on port {0} in {1} seconds. The app will fall back to {2}:{3}." -f $DiscoveryPort, $DiscoverySeconds, $FallbackCcuAddress, $FallbackCcuPort)
    }
}
catch {
    Add-Check "UDP CCU discovery" "FAIL" $_.Exception.Message
}

$ccuAddress = if ($discovery -and $discovery.CandidateAddress) { $discovery.CandidateAddress } else { $FallbackCcuAddress }
$ccuPort = if ($discovery -and $discovery.CandidatePort) { [int]$discovery.CandidatePort } else { $FallbackCcuPort }

$routeLocalAddress = $null
try {
    $routeLocalAddress = Get-RouteLocalAddress -RemoteAddress $ccuAddress -RemotePort $ccuPort
    $status = if ($routeLocalAddress -like "169.254.*") { "PASS" } else { "WARN" }
    Add-Check "Windows route local address" $status ("Windows would route to {0}:{1} using local address {2}." -f $ccuAddress, $ccuPort, $routeLocalAddress) ([pscustomobject]@{ CcuAddress = $ccuAddress; CcuPort = $ccuPort; LocalAddress = $routeLocalAddress })
}
catch {
    Add-Check "Windows route local address" "FAIL" $_.Exception.Message
}

$localAddress = $null
try {
    $allAddresses = @(Get-NetIPAddress -AddressFamily IPv4 -ErrorAction Stop |
        Where-Object { $_.IPAddress -notlike "127.*" } |
        Select-Object IPAddress, PrefixLength, InterfaceAlias, InterfaceIndex, AddressState)
    $selection = Get-AppPreferredLocalAddress -CcuAddress $ccuAddress -CcuPort $ccuPort -Addresses $allAddresses
    $localAddress = $selection.Address
    $status = if ($localAddress -like "169.254.*") { "PASS" } else { "WARN" }
    Add-Check "App preferred local address" $status ("App would bind CCU sockets to {0} via {1}." -f $selection.Address, $selection.Source) $selection
}
catch {
    Add-Check "App preferred local address" "FAIL" $_.Exception.Message
}

if ($localAddress) {
    try {
        $pair = Get-FreeAdjacentPortPair -LocalAddress $localAddress
        Add-Check "Adjacent local control/callback ports" "PASS" ("Can bind local myQA-style pair {0}/{1} on {2}." -f $pair.ControlPort, $pair.CallbackPort, $localAddress) $pair
    }
    catch {
        Add-Check "Adjacent local control/callback ports" "FAIL" $_.Exception.Message
    }
}

if (-not $SkipTcpProbe) {
    try {
        $tcp = Test-TcpControl -RemoteAddress $ccuAddress -RemotePort $ccuPort -LocalAddress $localAddress
        Add-Check "TCP control reachability" "PASS" ("Opened and closed TCP control socket {0} -> {1}." -f $tcp.LocalEndpoint, $tcp.RemoteEndpoint) $tcp
    }
    catch {
        Add-Check "TCP control reachability" "FAIL" $_.Exception.Message
    }
}
else {
    Add-Check "TCP control reachability" "INFO" "Skipped because -SkipTcpProbe was supplied."
}

try {
    $bridgeExe = Join-Path $PackageRoot "bridge\TankControllerBridge.exe"
    $fw = Test-FirewallRule -BridgeExe $bridgeExe
    if (-not $fw.Available) {
        Add-Check "Windows firewall bridge rule" "WARN" $fw.Message $fw
    }
    elseif ($fw.MatchingRules.Count -gt 0) {
        Add-Check "Windows firewall bridge rule" "PASS" $fw.Message $fw
    }
    else {
        Add-Check "Windows firewall bridge rule" "WARN" ($fw.Message + " If connect works but live position/chamber readings are idle, run Allow-TankBridgeRuntimeFirewall.cmd as Administrator from the unzipped bridge runtime package. The Windows allow-access popup may not create the exact bridge rule we need.") $fw
    }
}
catch {
    Add-Check "Windows firewall bridge rule" "WARN" $_.Exception.Message
}

try {
    $tcpToCcu = @(Get-NetTCPConnection -RemoteAddress $ccuAddress -ErrorAction SilentlyContinue |
        Select-Object LocalAddress, LocalPort, RemoteAddress, RemotePort, State, OwningProcess)
    if ($tcpToCcu.Count -eq 0) {
        Add-Check "Existing CCU TCP sessions" "PASS" "No existing TCP sessions to the selected CCU address." $tcpToCcu
    }
    else {
        Add-Check "Existing CCU TCP sessions" "WARN" ("Found {0} existing TCP session(s) to {1}. myQA or another runner may already own the CCU." -f $tcpToCcu.Count, $ccuAddress) $tcpToCcu
    }
}
catch {
    Add-Check "Existing CCU TCP sessions" "WARN" $_.Exception.Message
}

if ($LiveBridgeProbe) {
    try {
        $live = Invoke-LiveBridgeProbe -Root $PackageRoot
        $hasLive = $live.CallbackConnectionCount -gt 0 -and $live.LatestStatus -and $live.SampleCount -gt 0
        if ($hasLive) {
            Add-Check "Packaged bridge live callback" "PASS" ("Connected, received callback from {0}, saw {1} sample(s), then disconnected." -f $live.CallbackPeerEndpoint, $live.SampleCount) $live
        }
        else {
            Add-Check "Packaged bridge live callback" "FAIL" "Bridge connected but did not receive live callback samples before timeout." $live
        }
    }
    catch {
        Add-Check "Packaged bridge live callback" "FAIL" $_.Exception.Message
    }
}
else {
    Add-Check "Packaged bridge live callback" "INFO" "Skipped. Re-run with -LiveBridgeProbe when the CCU is ready and myQA Accept is disconnected."
}

Write-Host
Write-Host "Summary" -ForegroundColor Cyan
$results | Group-Object Status | Sort-Object Name | ForEach-Object {
    Write-Host ("  {0}: {1}" -f $_.Name, $_.Count)
}

if ($OutputJson) {
    $fullOutput = [System.IO.Path]::GetFullPath($OutputJson)
    $results | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $fullOutput -Encoding UTF8
    Write-Host
    Write-Host "Wrote JSON report: $fullOutput" -ForegroundColor Cyan
}

$failCount = @($results | Where-Object { $_.Status -eq "FAIL" }).Count
if ($failCount -gt 0) {
    exit 1
}
