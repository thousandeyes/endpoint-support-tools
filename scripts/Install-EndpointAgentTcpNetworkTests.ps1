# Copyright (c) 2021 ThousandEyes, Inc.

#Requires -Version 5.1
#Requires -RunAsAdministrator

<#
    .SYNOPSIS
        Installs ThousandEyes Endpoint Agent with support for TCP network tests.

    .DESCRIPTION
        This script will install, re-install or upgrade ThousandEyes Endpoint Agent with support for TCP network
        tests.

        On success, the script exits with a return code of 0. If msiexec.exe fails, the script returns the return
        code from msiexec.exe.

        Unless otherwise specified, optional installer features such as the IE, Chrome and Edge browser extensions
        are left in their current state. To forcibly add or remove these features, pass one or more of
        -EnableIeExtension:<$true|$false>, -EnableChromeExtension:<$true|$false> or
        -EnableEdgeExtension:<$true|$false> to this script.

    .PARAMETER InstallerFilePath
        Path to a ThousandEyes Endpoint Agent installer.

    .PARAMETER EnableIeExtension
        Enables support for collecting network metrics for whitelisted pages visited in Internet Explorer.

    .PARAMETER EnableChromeExtension
        Enables support for collecting network metrics for whitelisted pages visited in Google Chrome.

    .PARAMETER EnableEdgeExtension
        Enables support for collecting network metrics for whitelisted pages visited in Microsoft Edge.

    .EXAMPLE
        PS> .\Install-EndpointAgentTcpNetworkTests.ps1 -InstallerFilePath ".\Endpoint Agent for Acme Enterprises-x64-1.80.0.msi"

        Ensures TCP network test support is enabled, and the IE, Chrome and Edge browser extension features are
        left in their current state.

    .EXAMPLE
        PS> .\Install-EndpointAgentTcpNetworkTests.ps1 -InstallerFilePath ".\Endpoint Agent for Acme Enterprises-x64-1.80.0.msi" -EnableIeExtension:$false -EnableChromeExtension:$false -EnableEdgeExtension:$false

        Ensures TCP network test support is enabled, but the IE, Chrome and Edge browser extension features are not
        enabled.

    .NOTES
        This script must be run with administrator privileges.
#>
[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$InstallerFilePath,

    [Parameter()]
    [switch]$EnableIeExtension,

    [Parameter()]
    [switch]$EnableChromeExtension,

    [Parameter()]
    [switch]$EnableEdgeExtension
)
process {
    Set-StrictMode -Version 3

    $ErrorActionPreference = "Stop"
    $PSDefaultParameterValues = @{
      "Write-Error:ErrorAction" = "Continue"
    }

    if (-not $PSCmdlet.MyInvocation.BoundParameters.ContainsKey("InformationAction")) {
        # If caller hasn't specified, then display Write-Information messages.
        $InformationPreference = "Continue"
    }

    # Returns the properties table of the given msi as a hash table.
    function getPropertiesFromMsi([string]$msiPath) {
        $msiOpenDatabaseModeReadOnly = 0

        $installer = New-Object -ComObject "WindowsInstaller.Installer"
        $db = $null
        try {
            $db = $installer.OpenDatabase($msiPath, $msiOpenDatabaseModeReadOnly)
        }
        catch [Runtime.InteropServices.COMException] {
            # Not a valid MSI.
            return @{}
        }

        try {
            $query = "SELECT Property, Value FROM Property"
            $view = $db.OpenView($query)
            try {
                $null = $view.Execute()

                $propTable = @{}
                while ($true) {
                    $record = $view.Fetch()
                    if ($null -eq $record) {
                        break;
                    }

                    $key = $record.StringData(1)
                    $value = $record.StringData(2)

                    $propTable[$key] = $value
                }

                return $propTable
            }
            finally {
                $null = $view.Close()
                $view = $null
            }
        }
        finally {
            $db = $null
            $null = [GC]::Collect()
        }
    }

    # Returns the product codes matching the given upgrade code.
    function getRelatedProducts([string]$upgradeCode) {
        $installer = New-Object -ComObject "WindowsInstaller.Installer"
        $productsList = $installer.RelatedProducts($upgradeCode)

        $products = @()
        foreach ($product in $productsList) {
            $products += $product;
        }
        return , $products
    }

    # Returns a product info attribute value for the given product.
    function getProductInfo([string]$productCode, [string]$attribute) {
        $installer = New-Object -ComObject "WindowsInstaller.Installer"
        return $installer.ProductInfo($productCode, $attribute)
    }

    # For the given product, returns a hashtable mapping feature name => feature install state.
    function getProductFeatureStates([string]$productCode) {
        $installer = New-Object -ComObject "WindowsInstaller.Installer"
        $features = $installer.Features($productCode)

        $featureStates = @{}
        foreach ($featureName in $features) {
            $state = $installer.FeatureState($productCode, $featureName)
            $featureStates[$featureName] = $state
        }
        return $featureStates
    }

    # ThousandEyes upgrade codes.
    $UPGRADE_CODES = @(
        "{9FE0CE31-553B-4712-8C27-2E4F941557D5}",
        "{23218815-2B98-4B24-B0C3-6D7D13EDFA14}",
        "{6EAE54F3-D704-4784-B652-6B26088898A7}",
        "{2E68BE1C-2D76-4FDE-A220-83E5046FEC3C}",
        "{22E5DBDA-D9FA-4A3A-8D2E-5426E4B11EC5}",
        "{3C3FB442-646E-433E-908B-C4396BE7543E}"
    )

    # MSI functions require an absolute path.
    try {
        $InstallerFilePath = [string](Resolve-Path -LiteralPath $InstallerFilePath)
    }
    catch {
        Write-Error "`"$InstallerFilePath`" was not found."
        Exit 1
    }

    $msiProps = getPropertiesFromMsi $InstallerFilePath

    $productName = $msiProps["ProductName"]
    $version = $msiProps["ProductVersion"]
    $upgradeCode = $msiProps["UpgradeCode"]

    if (($null -eq $productName) -or
        ($null -eq $version) -or
        ($null -eq $upgradeCode) -or
        -not ($UPGRADE_CODES -contains $upgradeCode)) {
        Write-Error "`"$InstallerFilePath`" does not appear to be a valid ThousandEyes Endpoint Agent installer."
        Exit 1
    }

    $relatedProducts = getRelatedProducts $upgradeCode

    if ($relatedProducts.Count -gt 1) {
        Write-Error "Cannot continue: Windows installer reports multiple existing installations of $productName."
        Exit 1
    }

    # Assume package name is the same as the input file until we determine otherwise.
    $packageName = ([IO.FileInfo]$InstallerFilePath).BaseName

    # Default feature states for optional features
    $proposedFeatures = [ordered]@{
        "TcpNetworkTestsSupport" = $false
        "IeExtension"            = $false
        "ChromeExtension"        = $false
        "EdgeExtension"          = $false
    }

    if ($relatedProducts.Count -eq 0) {
        Write-Information "No existing $productName installation detected."
    }
    else {
        $oldProductCode = $relatedProducts[0]

        $oldProductName = getProductInfo $oldProductCode "ProductName"
        $oldVersion = getProductInfo $oldProductCode "VersionString"
        $oldPackageName = getProductInfo $oldProductCode "PackageName"

        if (($null -eq $oldProductName) -or
            ($null -eq $oldVersion) -or
            ($null -eq $oldPackageName)) {
            Write-Error "Failed to gather details of existing $productName installation."
            Exit 1
        }

        if ([Version]$oldVersion -gt [Version]$version) {
            Write-Error "A newer version of $oldProductName ($oldVersion) is already installed."
            Exit 1
        }

        Write-Information "Found existing installation: $oldProductName v$oldVersion"

        $packageName = $oldPackageName

        $msiInstallStateLocal = 3

        # Respect existing feature selection.
        $oldFeatureStates = getProductFeatureStates $oldProductCode
        foreach ($featureName in $oldFeatureStates.Keys) {
            if ($null -ne $proposedFeatures[$featureName]) {
                $proposedFeatures[$featureName] = ($oldFeatureStates[$featureName] -eq $msiInstallStateLocal)
            }
        }
    }

    # Always enable TCP Network Tests feature.
    $proposedFeatures["TcpNetworkTestsSupport"] = $true

    # If specified on the command line, adjust browser extension features. Otherwise leave them alone.
    if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey("EnableIeExtension")) {
        $proposedFeatures["IeExtension"] = $EnableIeExtension
    }
    if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey("EnableChromeExtension")) {
        $proposedFeatures["ChromeExtension"] = $EnableChromeExtension
    }
    if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey("EnableEdgeExtension")) {
        $proposedFeatures["EdgeExtension"] = $EnableEdgeExtension
    }

    $userTempDir = [IO.Path]::GetTempPath()

    $installerTempDir = Join-Path $userTempDir ([Guid]::NewGuid())
    $null = New-Item -ItemType Directory -Path $installerTempDir
    try {
        $installerLogPath = Join-Path `
            $userTempDir "ThousandEyesInstall-$((Get-Date).ToUniversalTime().ToString("yyyy-MM-dd_HH-mm-ss")).log"
        $installerTempPath = Join-Path $installerTempDir $packageName

        Copy-Item $InstallerFilePath -Destination $installerTempPath

        $enabledFeatures = @()
        $disabledFeatures = @()
        foreach ($featureName in $proposedFeatures.Keys) {
            if ($proposedFeatures[$featureName]) {
                $enabledFeatures += $featureName
            } else {
                $disabledFeatures += $featureName
            }
        }

        # Start-Process does not escape arguments. Further, msiexec performs non-standard argument parsing,
        # requiring e.g. VAR="foo bar" and not "VAR=foo bar".
        $msiexecArgs = @(
            "/i",
            "`"$installerTempPath`"",
            "ADDLOCAL=`"$($enabledFeatures -Join ",")`"",
            "REMOVE=`"$($disabledFeatures -Join ",")`"",
            "/qn",
            "/quiet"
            "/norestart",
            "/l*vx",
            "`"$installerLogPath`""
        )

        Write-Information "Installing $productName v$version..."
        Write-Information "`tCommand line: msiexec $($msiexecArgs -join " ")"

        $proc = Start-Process -FilePath "msiexec.exe" -ArgumentList $msiexecArgs -PassThru -Wait

        $ERROR_SUCCESS_REBOOT_REQUIRED = 3010
        if (($proc.ExitCode -eq 0) -or ($proc.ExitCode -eq $ERROR_SUCCESS_REBOOT_REQUIRED)) {
            Write-Information ("Installation was successful.`n" +
                               "`tmsiexec exit status: $($proc.ExitCode)`n" +
                               "`tInstaller logs: $installerLogPath")
            Exit 0
        }
        else {
            Write-Error ("Installation was unsuccessful. msiexec exit status: $($proc.ExitCode).`n" +
                         "Please provide ThousandEyes support with the installer logs: $installerLogPath")
            Exit $proc.ExitCode
        }
    }
    finally {
        $null = Remove-Item -LiteralPath $installerTempDir -Force -Recurse -ErrorAction Continue
    }
}
