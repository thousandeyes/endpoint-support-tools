# ThousandEyes Endpoint Agent Support Tools

#### [`Install-EndpointAgentTcpNetworkTests.ps1`](./scripts/Install-EndpointAgentTcpNetworkTests.ps1)

PowerShell script to install, re-install or upgrade ThousandEyes Endpoint Agent with support for TCP network tests. The script can be invoked with a command line like the following:

```
powershell.exe -ExecutionPolicy Bypass -NoProfile ^
    -File .\Path\To\Install-EndpointAgentTcpNetworkTests.ps1 ^
    -InstallerFilePath ".\Path\To\Endpoint Agent for Acme Enterprises-x64-1.80.0.msi"
```

The equivalent command, all one one line:
```
powershell.exe -ExecutionPolicy Bypass -NoProfile -File .\Path\To\Install-EndpointAgentTcpNetworkTests.ps1 -InstallerFilePath ".\Path\To\Endpoint Agent for Acme Enterprises-x64-1.80.0.msi"
```
