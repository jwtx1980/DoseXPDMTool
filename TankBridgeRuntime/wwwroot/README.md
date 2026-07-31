# Tank Controller Web App

Browser front end for the local tank controller bridge.

Run the real local bridge from the repo root:

```powershell
dotnet run --project .\TankControllerBridge\TankControllerBridge.csproj --urls http://127.0.0.1:8770
```

The browser calls the bridge API, and the bridge reuses the existing WinForms `TankSession`
protocol implementation. No browser-side simulation is used.
