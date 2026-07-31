TG-51 Tank Bridge Runtime
=========================

Purpose
-------
This package contains only the local tank bridge needed by the TG-51 tool. The
TG-51 tool talks to the bridge over localhost. The bridge talks to the IBA tank
CCU over the tank network.

Default local API:

    http://127.0.0.1:8770

Important endpoints:

    GET  /health
    GET  /capabilities
    GET  /api/state
    POST /api/connect
    POST /api/disconnect
    POST /api/move
    POST /api/move/isocenter

Recommended install
-------------------
1. Unzip this package.
2. Right-click Install-TankBridgeRuntime.cmd.
3. Choose Run as administrator.
4. Approve the Windows UAC prompt.

The installer copies the bridge to:

    C:\Program Files\TG51TankBridge

It also creates the Windows firewall rule needed for the CCU callback.

Why Administrator is required
-----------------------------
Administrator/elevated privileges are required only for setup actions:

- copying files into C:\Program Files
- creating/removing the Windows Defender Firewall rule

The TG-51 app does not need to run elevated during normal use after the bridge is
installed and the firewall rule exists. The TG-51 app can launch:

    C:\Program Files\TG51TankBridge\Start-TankBridgeRuntime.cmd

Firewall rule
-------------
The installer creates this rule:

    DisplayName:   TG-51 Tank Bridge Inbound
    Direction:     Inbound
    Action:        Allow
    Program:       C:\Program Files\TG51TankBridge\bridge\TankControllerBridge.exe
    Protocol:      TCP
    RemoteAddress: 169.254.0.0/16
    Profile:       Any

The rule is program-based rather than port-based because the bridge chooses a
local callback port dynamically and advertises that callback port to the CCU.
Some Windows APIs may display the same remote range as
169.254.0.0/255.255.0.0; the TG-51 app should accept either form.

If not using the installer
--------------------------
From the unzipped package folder, right-click:

    Allow-TankBridgeRuntimeFirewall.cmd

and run it as Administrator. This creates the same firewall rule for the bridge
exe in the current package folder.

TG-51 app behavior
------------------
At first run, the TG-51 app should:

1. Check http://127.0.0.1:8770/health.
2. If /health responds, check http://127.0.0.1:8770/capabilities before
   accepting the service. Require schemaVersion collector-capabilities-v1 and at
   least POST /api/connect plus POST /api/move.
3. If not running, start C:\Program Files\TG51TankBridge\Start-TankBridgeRuntime.cmd.
4. If the bridge is not installed, prompt the user to run Install-TankBridgeRuntime.cmd.
5. If connect succeeds but callbacks do not arrive, prompt the user to run the
   installer/firewall helper as Administrator.

The bridge may already be installed or used by another app. Reuse a valid
running bridge instead of reinstalling it casually; the installer may stop an
existing bridge process while updating files.

Tank/network setup
------------------
1. Connect the PC to the tank/CCU network.
2. Make sure myQA Accept is not actively connected to the CCU.
3. Launch the TG-51 app or Start-TankBridgeRuntime.cmd.
4. The bridge listens for CCU UDP discovery on port 1221.
5. If discovery is not heard, it falls back to 169.254.40.1:1222.

Useful environment overrides:

- TANK_CONTROLLER_ADDRESS, for example 169.254.40.1
- TANK_CONTROLLER_PORT, normally 1222
- TANK_LOCAL_ADDRESS, only if the bridge must use a specific PC adapter address
- TANK_DISCOVERY_PORT, normally 1221
- TANK_DISCOVERY_TIMEOUT_MS, normally 5000
- TANK_BRIDGE_DATA_DIR, defaults to %LOCALAPPDATA%\TG51TankBridge
- TANK_BRIDGE_LOG_DIR, defaults to %LOCALAPPDATA%\TG51TankBridge\logs

Runtime data
------------
Normal bridge runtime data is written to the user's local app data folder, not
beside the executable in Program Files:

    %LOCALAPPDATA%\TG51TankBridge

This includes bridge logs, callback frame logs, position CSVs, queue-run state,
and scan result artifacts. The TG-51 app can override this with
TANK_BRIDGE_DATA_DIR if a site needs a different writable location.

Diagnostic
----------
Run this from the installed or unzipped package folder:

    powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Unblock-File '.\Test-TankNetworkAssumptions.ps1'; & '.\Test-TankNetworkAssumptions.ps1' -OutputJson '.\tank-network.json'"

Live bridge callback check, with myQA Accept disconnected:

    powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Unblock-File '.\Test-TankNetworkAssumptions.ps1'; & '.\Test-TankNetworkAssumptions.ps1' -LiveBridgeProbe -OutputJson '.\tank-network-live.json'"

Callback troubleshooting
------------------------
If /api/connect succeeds but /api/state shows no recent lastCallbackAt or empty
latestStatus.x/y/z, the outbound tank control connection probably worked but
Windows blocked the inbound CCU callback. Run the installer or firewall helper
as Administrator, then restart the bridge and try again. callbackPeerEndpoint is
useful diagnostic evidence when populated, but it is not a hard movement gate by
itself.
