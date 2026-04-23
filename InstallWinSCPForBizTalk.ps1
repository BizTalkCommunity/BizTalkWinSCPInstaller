<#
.SYNOPSIS
    Installs the correct WinSCP version for Microsoft BizTalk Server 2016 or 2020.
  
.DESCRIPTION
    Detects the BizTalk Server edition and installed cumulative update, maps it to
    the required WinSCP version, and installs WinSCP.exe and WinSCPnet.dll to the
    BizTalk installation folder.

    Supports common PowerShell risk-mitigation parameters: -WhatIf and -Confirm.

    If the required version is already installed, no changes are made unless
    -ForceInstall is specified.

    Supports online and offline workflows:
    - Online: downloads nuget.exe and the WinSCP package automatically.
    - Offline: uses pre-downloaded NuGet/WinSCP files in -nugetDownloadFolder.

    Run in an elevated PowerShell session with write access to:
    - the BizTalk installation folder
    - the temporary download folder

    The download folder is not deleted after execution.
  
.PARAMETER nugetDownloadFolder
    Temporary folder used to store nuget.exe and extracted WinSCP package files.
    Default: $env:TEMP\nuget

    For offline installs, point this to a folder that already contains the
    required NuGet and WinSCP package contents.
.PARAMETER ForceInstall
    Reinstalls WinSCP even if the required version is already present in the
    BizTalk installation folder.
.EXAMPLE
    .\InstallWinSCPForBizTalk.ps1
    Detects BizTalk/CU and installs the required WinSCP version.
    Downloads NuGet and WinSCP if needed.
.EXAMPLE
    .\InstallWinSCPForBizTalk.ps1 -nugetDownloadFolder WinSCPTemp
    Uses WinSCPTemp as the package/download folder.
.EXAMPLE
    .\InstallWinSCPForBizTalk.ps1 -ForceInstall
    Forces reinstall of WinSCP.
.NOTES
    Authors: Thomas Canter, Sandro Pereira, Michael Stepensen, Niclas Oberg
    Last Updated: April 2026
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
#Parameters
Param(
    [Parameter(
        Mandatory = $false
    )]
    [ValidateNotNullOrEmpty()]
    [string]$nugetDownloadFolder = (Get-Item Env:TEMP).Value + "\nuget",
    [Parameter(
        Mandatory = $false
    )]
    [switch]$ForceInstall
)
#####################################################################
# Function to serach for specific Microsoft BizTalk Server Cumulative updates
#####################################################################
function Search-BTSCumulativeUpdate {
    Param(
        [string] $CumulativeUpdateID,
        [string] $BizTalkVersion
    )

    $uninstallPaths = @(
        "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    $installedApps = foreach ($path in $uninstallPaths) {
        Get-ItemProperty -Path $path -ErrorAction SilentlyContinue
    }

    return [bool]($installedApps | Where-Object {
        $name = [string]$_.DisplayName
        if ([string]::IsNullOrWhiteSpace($name)) {
            return $false
        }

        # Normalize whitespace and match key tokens with relaxed ordering.
        $normalized = ($name -replace '\s+', ' ').Trim()
        $hasBizTalkVersion = $normalized -match "(?i)\bBizTalk\b.*\b$BizTalkVersion\b"
        $hasKB = $normalized -match "(?i)\bKB\D*$CumulativeUpdateID\b"
        return ($hasBizTalkVersion -and $hasKB)
    })
}

#####################################################################
# Function to detect BizTalk cumulative update from DisplayName
#####################################################################
function Get-BTSCumulativeUpdateByDisplayName {
    Param(
        [string] $BizTalkVersion
    )

    $uninstallPaths = @(
        "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    $installedApps = foreach ($path in $uninstallPaths) {
        Get-ItemProperty -Path $path -ErrorAction SilentlyContinue
    }

    $cuMatches = foreach ($app in $installedApps) {
        $displayName = [string]$app.DisplayName
        if ([string]::IsNullOrWhiteSpace($displayName)) {
            continue
        }

        $normalized = ($displayName -replace '\s+', ' ').Trim()
        $hasBizTalkVersion = $normalized -match "(?i)\bBizTalk\b.*\b$BizTalkVersion\b"
        $hasCuMarker = $normalized -match '(?i)\b(Cumulative\s*Update|CU\s*\d+)\b'

        if ($hasBizTalkVersion -and $hasCuMarker) {
            $cuNumber = $null
            $kbNumber = $null

            if ($normalized -match '(?i)Cumulative\s*Update\s*(\d+)') {
                $cuNumber = [int]$Matches[1]
            }
            elseif ($normalized -match '(?i)\bCU\s*(\d+)\b') {
                $cuNumber = [int]$Matches[1]
            }

            if ($normalized -match '(?i)\bKB\D*(\d{6,8})\b') {
                $kbNumber = $Matches[1]
            }

            if ($cuNumber) {
                [pscustomobject]@{
                    CUNumber = $cuNumber
                    KB = $kbNumber
                    DisplayName = $displayName
                    InstallDate = $app.InstallDate
                }
            }
        }
    }

    $bestMatch = $cuMatches | Sort-Object -Property CUNumber, InstallDate -Descending | Select-Object -First 1
    if ($bestMatch) {
        return [pscustomobject]@{
            Found = $true
            CUNumber = $bestMatch.CUNumber
            KB = $bestMatch.KB
            DisplayName = $bestMatch.DisplayName
            InstallDate = $bestMatch.InstallDate
        }
    }

    return [pscustomobject]@{
        Found = $false
        CUNumber = 0
        KB = $null
        DisplayName = $null
        InstallDate = $null
    }
}
#####################################################################
# Function to write an error
#####################################################################
function Write-Error {
    Param([string] $ErrorMessage)
    Write-Host -ForegroundColor Red "$ErrorMessage";
}
#####################################################################
# Function to write success
#####################################################################
function Write-Success {
    Param([string] $SuccessMessage)
    Write-Host -ForegroundColor Green "$SuccessMessage";
}

#####################################################################
# Function to test if current session runs elevated
#####################################################################
function Test-IsAdministrator {
    $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

#####################################################################
# Function to resolve WinSCP package layout from extracted NuGet files
#####################################################################
function Resolve-WinSCPPackageLayout {
    Param(
        [string] $PackageRoot,
        [string] $ExeFileName,
        [string] $DllFileName
    )

    $resolvedExe = $null
    $resolvedDll = $null

    if (-not (Test-Path $PackageRoot)) {
        return [pscustomobject]@{
            ExePath = $null
            DllPath = $null
            IsResolved = $false
        }
    }

    $exeCandidates = @(
        (Join-Path $PackageRoot "tools\$ExeFileName"),
        (Join-Path $PackageRoot "content\$ExeFileName")
    )
    $dllCandidates = @(
        (Join-Path $PackageRoot "lib\netstandard2.0\$DllFileName"),
        (Join-Path $PackageRoot "lib\netstandard\$DllFileName"),
        (Join-Path $PackageRoot "lib\net\$DllFileName"),
        (Join-Path $PackageRoot "lib\$DllFileName")
    )

    foreach ($candidate in $exeCandidates) {
        if (Test-Path $candidate) {
            $resolvedExe = $candidate
            break
        }
    }

    foreach ($candidate in $dllCandidates) {
        if (Test-Path $candidate) {
            $resolvedDll = $candidate
            break
        }
    }

    if (-not $resolvedExe) {
        $exeFile = Get-ChildItem -Path $PackageRoot -Recurse -File -Filter $ExeFileName -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($exeFile) {
            $resolvedExe = $exeFile.FullName
        }
    }

    if (-not $resolvedDll) {
        $dllFile = Get-ChildItem -Path $PackageRoot -Recurse -File -Filter $DllFileName -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($dllFile) {
            $resolvedDll = $dllFile.FullName
        }
    }

    return [pscustomobject]@{
        ExePath = $resolvedExe
        DllPath = $resolvedDll
        IsResolved = [bool]($resolvedExe -and $resolvedDll)
    }
}
# Default $Continue flag to true, set to false to end the process
$Continue = $true;
$PrerequisiteFailure = $false;
  
$DebugPreference = "Continue";
$isAdministrator = Test-IsAdministrator
#####################################################################
# Default WinSCP configuration and package layout notes
# - Start with a safe default (BizTalk 2016 RTM -> WinSCP 5.7.7).
# - Actual WinSCP version is selected later from detected BizTalk CU.
# - NuGet package root format: WinSCP.<version>\
# - Typical EXE path: tools\WinSCP.exe (older packages may use content\WinSCP.exe)
# - Typical DLL paths:
#     lib\netstandard2.0\WinSCPnet.dll
#     lib\netstandard\WinSCPnet.dll
#     lib\net\WinSCPnet.dll
#     lib\WinSCPnet.dll
# - This script resolves EXE and DLL paths dynamically from extracted files.
#####################################################################
$winSCPVersion = "5.7.7"
$winSCPexeFile = "WinSCP.exe";
$winSCPdllFile = "WinSCPnet.dll";
$WinSCPexe = "WinSCP.$winSCPVersion\content\$winSCPexeFile"
$winSCPdll = "WinSCP.$winSCPVersion\lib\$winSCPdllFile"
$hashString = "##############################################################################";
$bangString = "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!";
$upString = "^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^";
##############################################################
# Checking for BizTalk Server
##############################################################
$bizTalkInstallFolder = (Get-Item Env:BTSINSTALLPATH).Value;
$bizTalkInstallFolderExists = Test-Path $bizTalkInstallFolder
$BizTalk2016ProductCode = '{B084F3A7-3E8F-4E7B-B673-EED1715D28ED}';
$BizTalk2020ProductCode = '{205F5836-7512-4A06-9E74-ADC8AFA0EEC5}';
$bizTalkProductCodeCurrent = (get-itemPropertyValue 'HKLM:\SOFTWARE\Microsoft\BizTalk Server\3.0' -Name 'ProductCodeCurrent')
$bizTalkProductName = (get-itemPropertyValue 'HKLM:\SOFTWARE\Microsoft\BizTalk Server\3.0' -Name 'ProductName')
$bizTalkProductVersion = (get-itemPropertyValue 'HKLM:\SOFTWARE\Microsoft\BizTalk Server\3.0' -Name 'ProductVersion')
if ($Continue) {
    Write-Success $hashString
    Write-Success 'Checking if Microsoft BizTalk Server is installed.';
    Write-Success $hashString
    if (-not $bizTalkInstallFolder) {
        Write-Error ('The Env:BTSINSTALLPATH doesn`t exist, checking to see if the path is in the registry HKLM:\SOFTWARE\Microsoft\BizTalk Server\3.0@InstallPath');
        $bizTalkInstallFolder = (get-itemPropertyValue 'HKLM:\SOFTWARE\Microsoft\BizTalk Server\3.0' -Name 'InstallPath')
          
        if (-not $bizTalkInstallFolder) {
            $Continue = false;
            Write-Error "Microsoft BizTalk Server was not located by checking the environment variable BTSINSTALLPATH and the Registry key for BizTalk, exiting the process";
            Write-Error "Please confirm that Microsoft BizTalk Server is installed on this system";
        }
    }
    if ($Continue -and -not $bizTalkInstallFolderExists) {
        $Continue = $false
        Write-Error "$bangString"
        Write-Error "Microsoft BizTalk Server installation was found in the registry.";
        Write-Error "Regardless, the $bizTalkInstallFolder folder does not exist.";
        Write-Error "Although it was found in the registry and/or the BTSINSTALLPATH environment setting";
        Write-Error "Exiting...";
        Write-Error "$bangString"
    }
    else {
        Write-Success "Located $bizTalkProductName version $bizTalkProductVersion.";
        Write-Success "Microsoft BizTalk Server is installed.";
        Write-Success "Located in the `'$bizTalkInstallFolder`' folder.";
        if ($bizTalkProductCodeCurrent -eq $BizTalk2020ProductCode) {
            $BizTalkVersion = "2020";
        }
        elseif ($bizTalkProductCodeCurrent -eq $BizTalk2016ProductCode) {
            $BizTalkVersion = "2016";
        }
        else {
            $Continue = $false;
            Write-Error "`n$bangString"
            Write-Error "Neither Microsoft BizTalk Server 2016 nor Microsoft BizTalk Server 2020 were found."
            Write-Error "Exiting";
            Write-Error "$bangString"
        }
  
    }
}
Write-Verbose  "The result of the search for the BizTalk Server:";
Write-Verbose "`$bizTalkInstallFolder       $bizTalkInstallFolder";
Write-Verbose "`$bizTalkInstallFolderExists $bizTalkInstallFolderExists";
Write-Verbose "`$BizTalk2016ProductCode     $BizTalk2016ProductCode";
Write-Verbose "`$BizTalk2020ProductCode     $BizTalk2020ProductCode";
Write-Verbose "`$bizTalkProductCodeCurrent  $bizTalkProductCodeCurrent";
Write-Verbose "`$bizTalkProductName         $bizTalkProductName";
Write-Verbose "`$bizTalkProductVersion      $bizTalkProductVersion";

# Fail fast for forced reinstall in non-elevated sessions.
if ($Continue -and -not $isAdministrator -and $ForceInstall) {
    Write-Error "`n$bangString"
    Write-Error "This PowerShell session is not running as Administrator."
    Write-Error "The BizTalk installation folder requires elevation for writes: $bizTalkInstallFolder"
    if (-not $WhatIfPreference) {
        Write-Error "ForceInstall requires write access and will fail without elevation."
        Write-Error "Please rerun from an elevated PowerShell session."
        Write-Error "$bangString"
        $PrerequisiteFailure = $true
        $Continue = $false
    }
    else {
        Write-Error "Continuing in read/dry-run mode. If a real install is required, rerun elevated."
        Write-Error "$bangString"
    }
}
  
$winSCPVersion = $null;
$btsKB = "none";
$bizTalkCUVer = "no CU";
$CUFound = $false;
if ($Continue) {
    Write-Success "`n$hashString"
    Write-Success "Determining what version of WinSCP to download"
    Write-Success "$hashString"
    ##############################################################
    # Deciding which is the correct version of WinSCP according
    # to Microsoft BizTalk Server version and cumulative update installed
    ##############################################################
    Write-Success "Detected Microsoft BizTalk Server $BizTalkVersion.";
    Write-Success "Testing to see which Cumulative Update is installed";
    if ($BizTalkVersion -eq "2020") {
        $winSCPVersion = "5.15.4"
        # Microsoft BizTalk Server 2020 CU mapping
        # CU   Build       KB(s)                 Release Date       WinSCP
        # CU6  3.13.895.0  5043408, 5048971      November 21, 2024  6.3.5
        # CU5  3.13.867.0  5032870               December 3, 2023   6.1.2
        # CU4  3.13.844.0  5009901               August 22, 2022    5.19.2
        # CU3  3.13.812.0  5007969               November 22, 2021  5.19.2
        # CU2  3.13.785.0  5003151               April 19, 2021     5.17.8
        # CU1  3.13.759.0  4538666               July 28, 2020      5.17.6
        # RTM  3.13.717.0  NA                    January 15, 2020   5.15.4
        $bts2020CUMap = @{
            6 = @{ KBs = @("5043408", "5048971"); WinSCP = "6.3.5" }
            5 = @{ KBs = @("5032870"); WinSCP = "6.1.2" }
            4 = @{ KBs = @("5009901"); WinSCP = "5.19.2" }
            3 = @{ KBs = @("5007969"); WinSCP = "5.19.2" }
            2 = @{ KBs = @("5003151"); WinSCP = "5.17.8" }
            1 = @{ KBs = @("4538666"); WinSCP = "5.17.6" }
        }

        $detected2020CU = Get-BTSCumulativeUpdateByDisplayName -BizTalkVersion $BizTalkVersion
        if ($detected2020CU.Found -and $bts2020CUMap.ContainsKey([int]$detected2020CU.CUNumber)) {
            $cuNumber = [int]$detected2020CU.CUNumber
            $cuEntry = $bts2020CUMap[$cuNumber]
            $bizTalkCUVer = "CU$cuNumber"
            $btsKB = if ($detected2020CU.KB) { $detected2020CU.KB } else { $cuEntry.KBs[0] }
            $winSCPVersion = $cuEntry.WinSCP
            $CUFound = $true
        }

        if (-not $CUFound) {
            foreach ($cuNumber in @(6, 5, 4, 3, 2, 1)) {
                $cuEntry = $bts2020CUMap[$cuNumber]
                foreach ($kb in $cuEntry.KBs) {
                    if (Search-BTSCumulativeUpdate -CumulativeUpdateID $kb -BizTalkVersion $BizTalkVersion) {
                        $bizTalkCUVer = "CU$cuNumber"
                        $btsKB = $kb
                        $winSCPVersion = $cuEntry.WinSCP
                        $CUFound = $true
                        break
                    }
                }
                if ($CUFound) {
                    break
                }
            }
        }
    }
    elseif ($BizTalkVersion -eq "2016") {
        # Microsoft BizTalk Server 2016 update mapping
        # Label         Build       KB        Release Date       WinSCP
        # CU9 and FP3   3.13.357.2  5005480   September 29, 2021 5.19.2
        # CU9           3.12.896.2  5005479   August 25, 2021    5.19.2
        # CU8 and FP3   3.13.349.2  4590075   January 6, 2021    5.15.9
        # CU8           3.12.880.2  4583530   December 7, 2020   5.15.9
        # CU7 and FP3   3.13.340.2  4536185   January 22, 2020   5.15.9
        # CU7           3.12.859.2  4528776   January 22, 2020   5.15.9
        # CU6 and FP3   3.12.843.2  4294900   February 7, 2019   5.13.1
        # CU6           3.12.843.2  4477494   February 28, 2019  5.13.1
        # CU5 and FP3   3.13.324.2  4103503   June 25, 2018      5.13.1
        # CU5 Hotfix    3.12.834.2  4345385   November 14, 2018  5.13.1
        # CU5           3.12.834.2  4132957   June 25, 2018      5.13.1
        # CU4 and FP2   3.13.252.2  4094130   April 2, 2018      5.7.7
        # CU4           3.12.823.2  4051353   January 30, 2018   5.7.7
        # CU3 and FP2   3.13.247.2  4054819   November 21, 2017  5.7.7
        # CU3 and FU1   3.13.177.2  4014788   November 15, 2017  5.7.7
        # CU3           3.12.815.2  4039664   September 1, 2017  5.7.7
        # CU2 / FU1     3.12.807.2  4021095   May 26, 2017       5.7.7
        # CU1           3.12.796.2  3208238   January 26, 2017   5.7.7
        # RTM           3.12.774.0  NA        September 30, 2016 5.7.7
        $bts2016UpdateMap = @(
            @{ Label = 'CU9 and FP3'; KB = '5005480'; WinSCP = '5.19.2' }
            @{ Label = 'CU9'; KB = '5005479'; WinSCP = '5.19.2' }
            @{ Label = 'CU8 and FP3'; KB = '4590075'; WinSCP = '5.15.9' }
            @{ Label = 'CU8'; KB = '4583530'; WinSCP = '5.15.9' }
            @{ Label = 'CU7 and FP3'; KB = '4536185'; WinSCP = '5.15.9' }
            @{ Label = 'CU7'; KB = '4528776'; WinSCP = '5.15.9' }
            @{ Label = 'CU6 and FP3'; KB = '4294900'; WinSCP = '5.13.1' }
            @{ Label = 'CU6'; KB = '4477494'; WinSCP = '5.13.1' }
            @{ Label = 'CU5 and FP3'; KB = '4103503'; WinSCP = '5.13.1' }
            @{ Label = 'CU5 Hotfix'; KB = '4345385'; WinSCP = '5.13.1' }
            @{ Label = 'CU5'; KB = '4132957'; WinSCP = '5.13.1' }
            @{ Label = 'CU4 and FP2'; KB = '4094130'; WinSCP = '5.7.7' }
            @{ Label = 'CU4'; KB = '4051353'; WinSCP = '5.7.7' }
            @{ Label = 'CU3 and FP2'; KB = '4054819'; WinSCP = '5.7.7' }
            @{ Label = 'CU3 and FU1'; KB = '4014788'; WinSCP = '5.7.7' }
            @{ Label = 'CU3'; KB = '4039664'; WinSCP = '5.7.7' }
            @{ Label = 'CU2 or FU1'; KB = '4021095'; WinSCP = '5.7.7' }
            @{ Label = 'CU1'; KB = '3208238'; WinSCP = '5.7.7' }
        )

        foreach ($update in $bts2016UpdateMap) {
            if (Search-BTSCumulativeUpdate -CumulativeUpdateID $update.KB -BizTalkVersion $BizTalkVersion) {
                $winSCPVersion = $update.WinSCP
                $btsKB = $update.KB
                $bizTalkCUVer = $update.Label
                $CUFound = $true
                break
            }
        }
    }
    if ($CUFound) {
        Write-Success "Detected Microsoft BizTalk Server $BizTalkVersion $bizTalkCUVer KB$btsKB";
    }
    else {
        # running Microsoft BizTalk Server without any Cumulative Updates, using original WinSCP Version
        Write-Success "Detected Microsoft BizTalk Server $BizTalkVersion with no cumulative updates.";
    }
    if ($ForceInstall) {
        Write-Success "ForceInstall was specified; this script will download/reuse and install WinSCP $winSCPVersion.";
    }
    else {
        Write-Success "If necessary, this script will download WinSCP $winSCPVersion";
    }
}
Write-Verbose  "The result of the search for the BizTalk Cumulative Update:";
Write-Verbose "`$winSCPVersion = $winSCPVersion";
Write-Verbose "`$btsKB         = $btsKB";
Write-Verbose "`$bizTalkCUVer  = $bizTalkCUVer";
Write-Verbose "`$CUFound       = $CUFound";
  
$btsWinSCPEXEProductVersionInstalled = "None";
$btsWinSCPDLLProductVersionInstalled = "None";
$btsWinSCPProductInstalledAndCorrect = $false;
$winSCPProductVersionRequired = $winSCPVersion;
$btsTargetWinSCPExe = $bizTalkInstallFolder + $winSCPexeFile;
$btsTargetWinSCPDll = $bizTalkInstallFolder + $winSCPdllFile;
if ($Continue) {
    Write-Success "`n$hashString"
    Write-Success "Checking to see if WinSCP $winSCPVersion is already";
    Write-Success "installed in the Microsoft BizTalk Server folder.";
    if ((Test-Path $btsTargetWinSCPExe) -and (Test-Path $btsTargetWinSCPDll)) {
        $btsWinSCPEXEProductVersionInstalled = (get-item $btsTargetWinSCPExe).VersionInfo.ProductVersion;
        $btsWinSCPDLLProductVersionInstalled = (get-item $btsTargetWinSCPDll).VersionInfo.ProductVersion;
        $winSCPProductVersionRequired = $winSCPVersion;
        if ($winSCPVersion.length -gt 2 -and $btsWinSCPEXEProductVersionInstalled.length -gt 2 -and $winSCPVersion.SubString($winSCPVersion.length - 2, 2) -ne '.0' -and $btsWinSCPEXEProductVersionInstalled.SubString($btsWinSCPEXEProductVersionInstalled.length - 2, 2) -eq '.0') {
            $winSCPProductVersionRequired = $winSCPVersion + '.0';
        }
        if ($winSCPProductVersionRequired -eq $btsWinSCPEXEProductVersionInstalled -and $winSCPProductVersionRequired -eq $btsWinSCPDLLProductVersionInstalled) {
            $btsWinSCPProductInstalledAndCorrect = $true;
        }
    }
    if ($btsWinSCPProductInstalledAndCorrect) {
        Write-Success "Detected WinSCP $winSCPVersion is already installed in Microsoft BizTalk Server.";
        if ($ForceInstall -and $btsWinSCPProductInstalledAndCorrect) {
            Write-Success "Reinstalling because ForceInstall was specified.";
            # Force a full reinstall path and only mark success after copy completes.
            $btsWinSCPProductInstalledAndCorrect = $false
        }
        else {
            Write-Success "Skipping installing the already installed version.";
            $Continue = $false;
        }
    }
    else {
        Write-Success "$bangstring"
        Write-Success "WinSCP $winSCPVersion is NOT installed in the";
        Write-Success "Microsoft BizTalk Server folder and needs to be installed.";
    }
    Write-Success "$hashString"
}
Write-Verbose "Check for existing WinSCP results were:";
Write-Verbose "`$btsWinSCPEXEProductVersionInstalled $btsWinSCPEXEProductVersionInstalled";
Write-Verbose "`$btsWinSCPDLLProductVersionInstalled $btsWinSCPDLLProductVersionInstalled";
Write-Verbose "`$winSCPProductVersionRequired        $winSCPProductVersionRequired";
Write-Verbose "`$btsWinSCPProductInstalledAndCorrect $btsWinSCPProductInstalledAndCorrect";
Write-Verbose "`$btsTargetWinSCPExe                  $btsTargetWinSCPExe";
Write-Verbose "`$btsTargetWinSCPDll                  $btsTargetWinSCPDll";

if ($Continue -and -not $isAdministrator -and -not $ForceInstall) {
    Write-Error "`n$bangString"
    Write-Error "This PowerShell session is not running as Administrator."
    Write-Error "The BizTalk installation folder requires elevation for writes: $bizTalkInstallFolder"
    Write-Error "Continuing in read/dry-run mode. If a real install is required, rerun elevated."
    Write-Error "$bangString"
}
  
  
if ($Continue) {
    $winSCPVersionInfo = $null
    if (-not [version]::TryParse($winSCPVersion, [ref]$winSCPVersionInfo)) {
        $Continue = $false
        if ([string]::IsNullOrEmpty($winSCPVersion)) {
            Write-Error "The WinSCP version was not set - the CU detection logic did not assign a value to `$winSCPVersion."
            Write-Error "This is likely a script bug. Review the BizTalk CU detection output above and confirm a CU or RTM baseline was matched."
        } else {
            Write-Error "The WinSCP version '$winSCPVersion' is not a valid dotted version number (e.g. 5.7.7 or 6.3.5)."
            Write-Error "This value came from the CU map table in this script. Check the WinSCP version string for BizTalk $BizTalkVersion $bizTalkCUVer in the table and correct it."
        }
        Write-Error $bangString
    }
}
if ($Continue) {
    # Resolve package file layout from what's already in the NuGet folder, if present.
    $winSCPPackageRoot = "$nugetDownloadFolder\WinSCP.$winSCPVersion"
    $winSCPLayout = Resolve-WinSCPPackageLayout -PackageRoot $winSCPPackageRoot -ExeFileName $winSCPexeFile -DllFileName $winSCPdllFile

    $WinSCPEXEDownload = $winSCPLayout.ExePath
    $WinSCPDllDownload = $winSCPLayout.DllPath
    $WinSCPEXEDownloadAlreadyExists = $false
    $WinSCPDllDownloadAlreadyExists = $false
    if ($WinSCPEXEDownload) {
        $WinSCPEXEDownloadAlreadyExists = Test-Path $WinSCPEXEDownload
    }
    if ($WinSCPDllDownload) {
        $WinSCPDllDownloadAlreadyExists = Test-Path $WinSCPDllDownload
    }
    $nugetDownloadFolderAlreadyExists = Test-Path $nugetDownloadFolder;
    $nugetDownloadFolderExists = $nugetDownloadFolderAlreadyExists;
    Write-Success "`n$hashString"
    Write-Success "Preparing the output folder to store the Nuget and WinSCP downloads"
    Write-Success "$hashString"
    if ($Continue -and -not $nugetDownloadFolderAlreadyExists) {
        ##############################################################
        # Prepare output folder
        ##############################################################
        if ($PSCmdlet.ShouldProcess("$nugetDownloadFolder", "Create Folder")) {
            Write-Success ("The target folder `'$nugetDownloadFolder`' doesn't exist, creating the folder.");
            New-Item -Path $nugetDownloadFolder -ItemType "Directory" -Force > $null
            $nugetDownloadFolderExists = Test-Path $nugetDownloadFolder
            if (-not $nugetDownloadFolderExists) {
                $Continue = $false
                Write-Error "An attempt to use the '" $nugetDownloadFolder "' directory for a download target failed.";
            }
        }
    }
    if ($Continue -and $nugetDownloadFolderExists) {
        Write-Success ("The target folder `'$nugetDownloadFolder`' is ready for use.");
    }
    Write-Verbose "Check and then potential creation of the target folder results:";
    Write-Verbose "`$nugetDownloadFolderAlreadyExists $nugetDownloadFolderExists";
    Write-Verbose "`$nugetDownloadFolderExists        $nugetDownloadFolderExists";
    Write-Verbose "`$WinSCPEXEDownload                $WinSCPEXEDownload";
    Write-Verbose "`$WinSCPDllDownload                $WinSCPDllDownload";
    Write-Verbose "`$WinSCPEXEDownloadAlreadyExists   $WinSCPEXEDownloadAlreadyExists";
    Write-Verbose "`$WinSCPDllDownloadAlreadyExists   $WinSCPDllDownloadAlreadyExists";
}
$targetNugetExe = "$nugetDownloadFolder\nuget.exe"
$targetNugetExeAlreadyExists = Test-Path $targetNugetExe
$targetNugetExeExists = $targetNugetExeAlreadyExists
if ($Continue) {
    ##############################################################
    # Download NuGet
    ##############################################################
    $sourceNugetExe = "https://dist.nuget.org/win-x86-commandline/latest/nuget.exe";
    if (-not $targetNugetExeAlreadyExists -or $ForceInstall) {
        Write-Success "`n$hashString";
        Write-Success "Downloading Nuget from:"
        Write-Success "`t`'$sourceNugetExe`'"
        Write-Success "Storing it in the folder";
        Write-Success "`t`'$nugetDownloadFolder`'"
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        if ($PSCmdlet.ShouldProcess("$sourceNugetExe -OutFile $targetNugetExe", "Run Invoke-WebRequest ")) {
            Invoke-WebRequest $sourceNugetExe -OutFile $targetNugetExe
            $targetNugetExeExists = Test-Path $targetNugetExe
            if (-not $targetNugetExeExists) {
                $Continue = $false
                Write-Error "`n$bangString";
                Write-Error "The download of the Nuget EXE from";
                Write-Error $sourceNugetExe;
                Write-Error "did not succeed";
                Write-Error "$bangString";
            }
        }
    }
    else {
        Write-Success "`n$hashString";
        Write-Success "NuGet already exists in the folder";
        Write-Success $nugetDownloadFolder;
        Write-Success "and doesn't need to be downloaded";
        Write-Success "$hashString";
    }
}
Write-Verbose "Check and then potential download NuGet results:";
Write-Verbose "`$targetNugetExeAlreadyExists $targetNugetExeAlreadyExists";
Write-Verbose "`$targetNugetExeExists        $targetNugetExeExists";
Write-Verbose "`$sourceNugetExe              $sourceNugetExe";
Write-Verbose "`$targetNugetExe              $targetNugetExe";
  
if ($Continue) {
    $winSCPPackageRoot = "$nugetDownloadFolder\WinSCP.$winSCPVersion"
    $winSCPLayout = Resolve-WinSCPPackageLayout -PackageRoot $winSCPPackageRoot -ExeFileName $winSCPexeFile -DllFileName $winSCPdllFile
    $WinSCPEXEDownload = $winSCPLayout.ExePath
    $WinSCPDllDownload = $winSCPLayout.DllPath
    $WinSCPEXEAlreadyExists = [bool]$WinSCPEXEDownload -and (Test-Path $WinSCPEXEDownload)
    $WinSCPDLLAlreadyExists = [bool]$WinSCPDllDownload -and (Test-Path $WinSCPDllDownload)
    $WinSCPEXEExists = $WinSCPEXEAlreadyExists
    $WinSCPDLLExists = $WinSCPDLLAlreadyExists
    $getWinSCP = "'$targetNugetExe' Install WinSCP -Version $winSCPVersion -NonInteractive -OutputDirectory '$nugetDownloadFolder'"
    Write-Success "`n$hashString";
    Write-Success "Downloading WinSCP version $winSCPVersion from NuGet";
    Write-Success "`t$getWinSCP";
    Write-Success "Storing it in the folder:"
    Write-Success "`t`'$nugetDownloadFolder`'";
    Write-Success "$hashString";
    if (-not $WinSCPEXEAlreadyExists -or -not $WinSCPDLLAlreadyExists) {
        if ($PSCmdlet.ShouldProcess("$getWinSCP", "Run Command")) {
            & $targetNugetExe Install WinSCP -Version $winSCPVersion -NonInteractive -OutputDirectory $nugetDownloadFolder
            $winSCPLayout = Resolve-WinSCPPackageLayout -PackageRoot $winSCPPackageRoot -ExeFileName $winSCPexeFile -DllFileName $winSCPdllFile
            $WinSCPEXEDownload = $winSCPLayout.ExePath
            $WinSCPDllDownload = $winSCPLayout.DllPath
            $WinSCPEXEExists = [bool]$WinSCPEXEDownload -and (Test-Path $WinSCPEXEDownload)
            $WinSCPDLLExists = [bool]$WinSCPDllDownload -and (Test-Path $WinSCPDllDownload)
            if (-not $WinSCPEXEExists -or -not $WinSCPDLLExists) {
                $Continue = $false
                Write-Error "`n$bangString";
                Write-Error "WinSCP $winSCPVersion was not properly downloaded.";
                Write-Error "Check the folder and error messages above:";
                Write-Error "$nugetDownloadFolder";
                Write-Error "And determine what files did download or did not download.";
                Write-Error "$bangString";
            }
        }
    }
}
Write-Verbose "Check and then potential download WinSCP results:";
Write-Verbose "`$getWinSCP              $getWinSCP";
Write-Verbose "`$WinSCPEXEAlreadyExists $WinSCPEXEAlreadyExists";
Write-Verbose "`$WinSCPDLLAlreadyExists $WinSCPDLLAlreadyExists";
Write-Verbose "`$WinSCPDllDownload      $WinSCPDllDownload";
Write-Verbose "`$WinSCPEXEExists        $WinSCPEXEExists";
Write-Verbose "`$WinSCPDLLExists        $WinSCPDLLExists";
  
##############################################################
# Installing WinSCP to Microsoft BizTalk Server Folder
##############################################################
$WinSCPTargetEXEExists = Test-Path $btsTargetWinSCPExe
$WinSCPDLLTargetExists = Test-Path $btsTargetWinSCPDll
if ($Continue -and -not $btsWinSCPProductInstalledAndCorrect) {
    Write-Success "`n$hashString";
    Write-Success "Installing WinSCP";
    Write-Success "$hashString";
    #Copy WinSCP items to Microsoft BizTalk Server Folder
    Write-Success "Copying WinSCP version $winSCPVersion to Microsoft BizTalk Server Folder:";
    Write-Success "`t`'$bizTalkInstallFolder'`.";
    $copySourceSummary = if ($WinSCPEXEDownload -and $WinSCPDllDownload) {
        "$WinSCPEXEDownload and $WinSCPDllDownload"
    }
    else {
        "WinSCP package files for version $winSCPVersion"
    }
    if ($PSCmdlet.ShouldProcess("$copySourceSummary to `'$bizTalkInstallFolder`'", "Copy Files")) {
        try {
            Copy-Item -Path $WinSCPEXEDownload -Destination $bizTalkInstallFolder -Force -ErrorAction Stop
            Copy-Item -Path $WinSCPDllDownload -Destination $bizTalkInstallFolder -Force -ErrorAction Stop
            $WinSCPTargetEXEExists = Test-Path $btsTargetWinSCPExe
            $WinSCPDLLTargetExists = Test-Path $btsTargetWinSCPDll
            if ($WinSCPTargetEXEExists -and $WinSCPDLLTargetExists) {
                $btsWinSCPProductInstalledAndCorrect = $true;
            }
            else {
                $Continue = $false
                if (-not $WinSCPTargetEXEExists) {
                    Write-Error "The $winSCPexeFile file version $winSCPVersion";
                    Write-Error "It was not properly copied to the target folder `'$bizTalkInstallFolder`'.";
                }
                if (-not $WinSCPDLLTargetExists) {
                    $Continue = $false
                    Write-Error "The $winSCPdllFile file version $winSCPVersion";
                    Write-Error "Was not properly copied to the target folder `'$bizTalkInstallFolder`'.";
                }
            }
        }
        catch {
            $Continue = $false
            Write-Error "Failed to copy WinSCP files to the BizTalk installation folder."
            Write-Error "$($_.Exception.Message)"
        }
    }
}
Write-Verbose "Check and then potential copy WinSCP to BizTalk results:";
Write-Verbose "`$bizTalkInstallFolder                $bizTalkInstallFolder";
Write-Verbose "`$WinSCPEXEDownload                   $WinSCPEXEDownload";
Write-Verbose "`$WinSCPDllDownload                   $WinSCPDllDownload";
Write-Verbose "`$WinSCPTargetEXEExists               $WinSCPTargetEXEExists";
Write-Verbose "`$WinSCPDLLTargetExists               $WinSCPDLLTargetExists";
Write-Verbose "`$btsWinSCPProductInstalledAndCorrect $btsWinSCPProductInstalledAndCorrect";
  
      
if ($btsWinSCPProductInstalledAndCorrect) {
    Write-Success "`n$bangString";
    Write-Success "WinSCP $winSCPVersion is installed.";
    Write-Success "Microsoft BizTalk Server`'s SFTP Adapter will use this version of WinSCP.";
    Write-Success "$upString";
}
elseif (-not $WhatIfPreference) {
    Write-Error "`n$bangString";
    if ($PrerequisiteFailure) {
        Write-Error "Installation did not run because one or more prerequisites were not met.";
        Write-Error "Please address the prerequisite errors above and rerun the script.";
    }
    else {
        Write-Error "Something went wrong during installation and the installation did not work.";
        Write-Error "Please inspect the errors above and resolve them.";
    }
    Write-Error "Exiting...";
    Write-Error "$bangString";
}
elseif ($WhatIfPreference) {
    Write-Success "`n$bangString";
    Write-Success "The parameter -WhatIf was set and this script executed without making";
    Write-Success "any changes and the output should checked to determine if it would have ";
    Write-Success "run correctly.";
    Write-Success "$bangString";
}

