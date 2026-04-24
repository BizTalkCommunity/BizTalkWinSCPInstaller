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
# Import core module functions used by this installer workflow
#####################################################################
$coreModulePath = Join-Path $PSScriptRoot "src\InstallWinSCPForBizTalk.Core.psm1"
if (-not (Test-Path $coreModulePath)) {
    throw "Required core module was not found: $coreModulePath"
}
Import-Module $coreModulePath

$utilsModulePath = Join-Path $PSScriptRoot "src\InstallWinSCPForBizTalk.Utils.psm1"
Import-Module $utilsModulePath

#####################################################################
# Default $Continue flag to true, set to false to end the process
$Continue = $true;
$PrerequisiteFailure = $false;

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
$hashString = Get-InstallerBannerLine -Name 'Hash'
$bangString = Get-InstallerBannerLine -Name 'Bang'
$upString = Get-InstallerBannerLine -Name 'Up'
##############################################################
# Checking for BizTalk Server
##############################################################
$bizTalkInstallFolderFromEnv = (Get-Item Env:BTSINSTALLPATH).Value
$bizTalkInstallFolderFromRegistry = (get-itemPropertyValue 'HKLM:\SOFTWARE\Microsoft\BizTalk Server\3.0' -Name 'InstallPath')
$bizTalkInstallFolderResolution = Resolve-BizTalkInstallFolder -EnvironmentInstallPath $bizTalkInstallFolderFromEnv -RegistryInstallPath $bizTalkInstallFolderFromRegistry
$bizTalkInstallFolder = $bizTalkInstallFolderResolution.InstallPath
$bizTalkInstallFolderExists = $bizTalkInstallFolderResolution.Exists
$bizTalkProductCodeCurrent = (get-itemPropertyValue 'HKLM:\SOFTWARE\Microsoft\BizTalk Server\3.0' -Name 'ProductCodeCurrent')
$bizTalkProductName = (get-itemPropertyValue 'HKLM:\SOFTWARE\Microsoft\BizTalk Server\3.0' -Name 'ProductName')
$bizTalkProductVersion = (get-itemPropertyValue 'HKLM:\SOFTWARE\Microsoft\BizTalk Server\3.0' -Name 'ProductVersion')
if ($Continue) {
    Write-InstallerSectionHeader -Title 'Checking if Microsoft BizTalk Server is installed.'
    if (-not $bizTalkInstallFolderResolution.IsFound) {
        Write-BizTalkRegistryFallbackNotice
        $Continue = $false;
        Write-BizTalkNotLocatedError
    }
    elseif ($bizTalkInstallFolderResolution.Source -eq 'Registry') {
        Write-BizTalkRegistryFallbackNotice
    }
    if ($Continue -and -not $bizTalkInstallFolderExists) {
        $Continue = $false
        Write-InstallerBangError -MessageLines @(
            'Microsoft BizTalk Server installation was found in the registry.'
            "Regardless, the $bizTalkInstallFolder folder does not exist."
            'Although it was found in the registry and/or the BTSINSTALLPATH environment setting'
            'Exiting...'
        )
    }
    else {
        Write-BizTalkLocatedSuccess -ProductName $bizTalkProductName -ProductVersion $bizTalkProductVersion -InstallFolder $bizTalkInstallFolder
        $BizTalkVersion = Get-BizTalkVersionFromProductCode -ProductCode $bizTalkProductCodeCurrent
        if ($null -eq $BizTalkVersion) {
            $Continue = $false;
            Write-InstallerBangError -LeadingNewLine -MessageLines @(
                'Neither Microsoft BizTalk Server 2016 nor Microsoft BizTalk Server 2020 were found.'
                'Exiting'
            )
        }
  
    }
}
Write-BizTalkSearchSnapshot -InstallFolder $bizTalkInstallFolder -InstallFolderExists $bizTalkInstallFolderExists -ProductCodeCurrent $bizTalkProductCodeCurrent -ProductName $bizTalkProductName -ProductVersion $bizTalkProductVersion

# Fail fast for forced reinstall in non-elevated sessions.
if ($Continue -and -not $isAdministrator -and $ForceInstall) {
    $nonAdminForceInstallMessages = @(
        'This PowerShell session is not running as Administrator.'
        "The BizTalk installation folder requires elevation for writes: $bizTalkInstallFolder"
    )
    if (-not $WhatIfPreference) {
        Write-InstallerBangError -LeadingNewLine -MessageLines ($nonAdminForceInstallMessages + @(
            'ForceInstall requires write access and will fail without elevation.'
            'Please rerun from an elevated PowerShell session.'
        ))
        $PrerequisiteFailure = $true
        $Continue = $false
    }
    else {
        Write-InstallerBangError -LeadingNewLine -MessageLines ($nonAdminForceInstallMessages + @(
            'Continuing in read/dry-run mode. If a real install is required, rerun elevated.'
        ))
    }
}
  
$winSCPVersion = $null;
$btsKB = "none";
$bizTalkCUVer = "no CU";
$CUFound = $false;
if ($Continue) {
    Write-InstallerSectionHeader -Title 'Determining what version of WinSCP to download' -LeadingNewLine
    ##############################################################
    # Deciding which is the correct version of WinSCP according
    # to Microsoft BizTalk Server version and cumulative update installed
    ##############################################################
    Write-BizTalkCuDetectionStart -BizTalkVersion $BizTalkVersion
    $cuResult = Get-WinSCPVersionForBizTalk -BizTalkVersion $BizTalkVersion
    if ($null -eq $cuResult) {
        $Continue = $false
        Write-InstallerBangError -LeadingNewLine -MessageLines @(
            "Could not determine the required WinSCP version for BizTalk Server $BizTalkVersion."
        )
    }
    else {
        $winSCPVersion = $cuResult.WinSCPVersion
        $btsKB         = $cuResult.KB
        $bizTalkCUVer  = $cuResult.CULabel
        $CUFound       = $cuResult.CUFound
    }
    if ($CUFound) {
        Write-InstallerSuccess "Detected Microsoft BizTalk Server $BizTalkVersion $bizTalkCUVer KB$btsKB";
    }
    else {
        # running Microsoft BizTalk Server without any Cumulative Updates, using original WinSCP Version
        Write-InstallerSuccess "Detected Microsoft BizTalk Server $BizTalkVersion with no cumulative updates.";
    }
    if ($ForceInstall) {
        Write-InstallerSuccess "ForceInstall was specified; this script will download/reuse and install WinSCP $winSCPVersion.";
    }
    else {
        Write-InstallerSuccess "If necessary, this script will download WinSCP $winSCPVersion";
    }
}
Write-BizTalkCuSearchSnapshot -WinSCPVersion $winSCPVersion -KB $btsKB -CULabel $bizTalkCUVer -CUFound $CUFound
  
$btsWinSCPEXEProductVersionInstalled = "None";
$btsWinSCPDLLProductVersionInstalled = "None";
$btsWinSCPProductInstalledAndCorrect = $false;
$installExecutionPlan = $null
$winSCPProductVersionRequired = $winSCPVersion;
$btsTargetWinSCPExe = $bizTalkInstallFolder + $winSCPexeFile;
$btsTargetWinSCPDll = $bizTalkInstallFolder + $winSCPdllFile;
if ($Continue) {
    Write-InstallerDelimitedMessage -MessageLines @(
        "Checking to see if WinSCP $winSCPVersion is already"
        "installed in the Microsoft BizTalk Server folder."
    ) -Delimiter 'Hash' -Level 'Success' -LeadingNewLine
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
    $installExecutionPlan = Get-InstallExecutionPlan -IsAdministrator $isAdministrator -ForceInstall ([bool]$ForceInstall) -WhatIf ([bool]$WhatIfPreference) -AlreadyInstalledCorrect $btsWinSCPProductInstalledAndCorrect
    if ($btsWinSCPProductInstalledAndCorrect) {
        Write-InstallerSuccess "Detected WinSCP $winSCPVersion is already installed in Microsoft BizTalk Server.";
        if ($installExecutionPlan.ShouldReinstall) {
            Write-InstallerSuccess "Reinstalling because ForceInstall was specified.";
            # Force a full reinstall path and only mark success after copy completes.
            $btsWinSCPProductInstalledAndCorrect = $false
        }
        elseif ($installExecutionPlan.ShouldSkipBecauseInstalled) {
            Write-InstallerSuccess "Skipping installing the already installed version.";
            $Continue = $false;
        }
    }
    else {
        Write-WinSCPNotInstalledNotice -WinSCPVersion $winSCPVersion
    }
    Write-InstallerSuccess "$hashString"
}
Write-ExistingWinSCPSnapshot -ExeProductVersionInstalled $btsWinSCPEXEProductVersionInstalled -DllProductVersionInstalled $btsWinSCPDLLProductVersionInstalled -ProductVersionRequired $winSCPProductVersionRequired -InstalledAndCorrect $btsWinSCPProductInstalledAndCorrect -TargetExePath $btsTargetWinSCPExe -TargetDllPath $btsTargetWinSCPDll

if ($Continue -and $installExecutionPlan -and $installExecutionPlan.RequiresElevationWarning -and -not $ForceInstall) {
    Write-InstallerBangError -LeadingNewLine -MessageLines @(
        'This PowerShell session is not running as Administrator.'
        "The BizTalk installation folder requires elevation for writes: $bizTalkInstallFolder"
        'Continuing in read/dry-run mode. If a real install is required, rerun elevated.'
    )
}
  
  
if ($Continue) {
    $versionValidation = Test-WinSCPVersionString -WinSCPVersion $winSCPVersion
    if (-not $versionValidation.IsValid) {
        $Continue = $false
        if ($versionValidation.ErrorCode -eq 'MissingVersion') {
            Write-InstallerError "The WinSCP version was not set - the CU detection logic did not assign a value to `$winSCPVersion."
            Write-InstallerError "This is likely a script bug. Review the BizTalk CU detection output above and confirm a CU or RTM baseline was matched."
        }
        else {
            Write-InstallerError "The WinSCP version '$winSCPVersion' is not a valid dotted version number (e.g. 5.7.7 or 6.3.5)."
            Write-InstallerError "This value came from the CU map table in this script. Check the WinSCP version string for BizTalk $BizTalkVersion $bizTalkCUVer in the table and correct it."
        }
        Write-InstallerError $bangString
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
    Write-InstallerSectionHeader -Title 'Preparing the output folder to store the Nuget and WinSCP downloads' -LeadingNewLine
    if ($Continue -and -not $nugetDownloadFolderAlreadyExists) {
        ##############################################################
        # Prepare output folder
        ##############################################################
        if ($PSCmdlet.ShouldProcess("$nugetDownloadFolder", "Create Folder")) {
            Write-InstallerSuccess ("The target folder `'$nugetDownloadFolder`' doesn't exist, creating the folder.");
            New-Item -Path $nugetDownloadFolder -ItemType "Directory" -Force > $null
            $nugetDownloadFolderExists = Test-Path $nugetDownloadFolder
            if (-not $nugetDownloadFolderExists) {
                $Continue = $false
                Write-InstallerError "An attempt to use the '" $nugetDownloadFolder "' directory for a download target failed.";
            }
        }
    }
    if ($Continue -and $nugetDownloadFolderExists) {
        Write-InstallerSuccess ("The target folder `'$nugetDownloadFolder`' is ready for use.");
    }
    Write-TargetFolderPreparationSnapshot -NugetDownloadFolderAlreadyExists $nugetDownloadFolderAlreadyExists -NugetDownloadFolderExists $nugetDownloadFolderExists -WinSCPEXEDownload $WinSCPEXEDownload -WinSCPDllDownload $WinSCPDllDownload -WinSCPEXEDownloadAlreadyExists $WinSCPEXEDownloadAlreadyExists -WinSCPDllDownloadAlreadyExists $WinSCPDllDownloadAlreadyExists
}
$targetNugetExe = "$nugetDownloadFolder\nuget.exe"
$targetNugetExeAlreadyExists = Test-Path $targetNugetExe
$targetNugetExeExists = $targetNugetExeAlreadyExists
if ($Continue) {
    ##############################################################
    # Download NuGet
    ##############################################################
    $sourceNugetExe = "https://dist.nuget.org/win-x86-commandline/latest/nuget.exe";
    $nugetDownloadPlan = Get-NuGetDownloadPlan -TargetNugetExeAlreadyExists ([bool]$targetNugetExeAlreadyExists) -ForceInstall ([bool]$ForceInstall) -SourceNugetExe $sourceNugetExe -TargetNugetExe $targetNugetExe
    if ($nugetDownloadPlan.ShouldDownload) {
        Write-InstallerSuccess "`n$hashString";
        Write-InstallerSuccess "Downloading Nuget from:"
        Write-InstallerSuccess "`t`'$sourceNugetExe`'"
        Write-InstallerSuccess "Storing it in the folder";
        Write-InstallerSuccess "`t`'$nugetDownloadFolder`'"
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        if ($PSCmdlet.ShouldProcess($nugetDownloadPlan.ShouldProcessTarget, "Run Invoke-WebRequest ")) {
            Invoke-WebRequest $sourceNugetExe -OutFile $targetNugetExe
            $targetNugetExeExists = Test-Path $targetNugetExe
            if (-not $targetNugetExeExists) {
                $Continue = $false
                Write-InstallerBangError -LeadingNewLine -MessageLines @(
                    'The download of the Nuget EXE from'
                    $sourceNugetExe
                    'did not succeed'
                )
            }
        }
    }
    else {
        Write-InstallerDelimitedMessage -MessageLines @(
            'NuGet already exists in the folder'
            $nugetDownloadFolder
            "and doesn't need to be downloaded"
        ) -Delimiter 'Hash' -Level 'Success' -LeadingNewLine
    }
}
Write-NuGetDownloadSnapshot -TargetNugetExeAlreadyExists $targetNugetExeAlreadyExists -TargetNugetExeExists $targetNugetExeExists -SourceNugetExe $sourceNugetExe -TargetNugetExe $targetNugetExe

$getWinSCP = $null
$WinSCPEXEAlreadyExists = $false
$WinSCPDLLAlreadyExists = $false
$WinSCPEXEExists = $false
$WinSCPDLLExists = $false

if ($Continue) {
    $winSCPPackageRoot = "$nugetDownloadFolder\WinSCP.$winSCPVersion"
    $winSCPLayout = Resolve-WinSCPPackageLayout -PackageRoot $winSCPPackageRoot -ExeFileName $winSCPexeFile -DllFileName $winSCPdllFile
    $WinSCPEXEDownload = $winSCPLayout.ExePath
    $WinSCPDllDownload = $winSCPLayout.DllPath
    $WinSCPEXEAlreadyExists = [bool]$WinSCPEXEDownload -and (Test-Path $WinSCPEXEDownload)
    $WinSCPDLLAlreadyExists = [bool]$WinSCPDllDownload -and (Test-Path $WinSCPDllDownload)
    $WinSCPEXEExists = $WinSCPEXEAlreadyExists
    $WinSCPDLLExists = $WinSCPDLLAlreadyExists
    $winSCPDownloadPlan = Get-WinSCPPackageDownloadPlan -WinSCPEXEAlreadyExists ([bool]$WinSCPEXEAlreadyExists) -WinSCPDLLAlreadyExists ([bool]$WinSCPDLLAlreadyExists) -TargetNugetExe $targetNugetExe -WinSCPVersion $winSCPVersion -NugetDownloadFolder $nugetDownloadFolder
    $getWinSCP = $winSCPDownloadPlan.CommandPreview
    Write-InstallerDelimitedMessage -MessageLines @(
        "Downloading WinSCP version $winSCPVersion from NuGet"
        "`t$getWinSCP"
        'Storing it in the folder:'
        "`t`'$nugetDownloadFolder`'"
    ) -Delimiter 'Hash' -Level 'Success' -LeadingNewLine
    if ($winSCPDownloadPlan.ShouldDownload) {
        if ($PSCmdlet.ShouldProcess("$getWinSCP", "Run Command")) {
            & $targetNugetExe Install WinSCP -Version $winSCPVersion -NonInteractive -OutputDirectory $nugetDownloadFolder
            $winSCPLayout = Resolve-WinSCPPackageLayout -PackageRoot $winSCPPackageRoot -ExeFileName $winSCPexeFile -DllFileName $winSCPdllFile
            $WinSCPEXEDownload = $winSCPLayout.ExePath
            $WinSCPDllDownload = $winSCPLayout.DllPath
            $WinSCPEXEExists = [bool]$WinSCPEXEDownload -and (Test-Path $WinSCPEXEDownload)
            $WinSCPDLLExists = [bool]$WinSCPDllDownload -and (Test-Path $WinSCPDllDownload)
            if (-not $WinSCPEXEExists -or -not $WinSCPDLLExists) {
                $Continue = $false
                Write-InstallerBangError -LeadingNewLine -MessageLines @(
                    "WinSCP $winSCPVersion was not properly downloaded."
                    'Check the folder and error messages above:'
                    $nugetDownloadFolder
                    'And determine what files did download or did not download.'
                )
            }
        }
    }
}
Write-WinSCPDownloadSnapshot -GetWinSCP $getWinSCP -WinSCPEXEAlreadyExists $WinSCPEXEAlreadyExists -WinSCPDLLAlreadyExists $WinSCPDLLAlreadyExists -WinSCPDllDownload $WinSCPDllDownload -WinSCPEXEExists $WinSCPEXEExists -WinSCPDLLExists $WinSCPDLLExists
  
##############################################################
# Installing WinSCP to Microsoft BizTalk Server Folder
##############################################################
$WinSCPTargetEXEExists = Test-Path $btsTargetWinSCPExe
$WinSCPDLLTargetExists = Test-Path $btsTargetWinSCPDll
if ($Continue -and -not $btsWinSCPProductInstalledAndCorrect) {
    Write-InstallerSectionHeader -Title 'Installing WinSCP' -LeadingNewLine
    #Copy WinSCP items to Microsoft BizTalk Server Folder
    Write-InstallerSuccess "Copying WinSCP version $winSCPVersion to Microsoft BizTalk Server Folder:";
    Write-InstallerSuccess "`t`'$bizTalkInstallFolder'`.";
    $copySourceSummary = if ($WinSCPEXEDownload -and $WinSCPDllDownload) {
        "$WinSCPEXEDownload and $WinSCPDllDownload"
    }
    else {
        "WinSCP package files for version $winSCPVersion"
    }
    if ($PSCmdlet.ShouldProcess("$copySourceSummary to `'$bizTalkInstallFolder`'", "Copy Files")) {
        $installResult = Invoke-WinSCPTargetInstall -SourceExePath $WinSCPEXEDownload -SourceDllPath $WinSCPDllDownload -TargetFolder $bizTalkInstallFolder -TargetExePath $btsTargetWinSCPExe -TargetDllPath $btsTargetWinSCPDll -WinSCPVersion $winSCPVersion -ExeFileName $winSCPexeFile -DllFileName $winSCPdllFile
        $WinSCPTargetEXEExists = $installResult.TargetExeExists
        $WinSCPDLLTargetExists = $installResult.TargetDllExists
        if ($installResult.InstalledSuccessfully) {
            $btsWinSCPProductInstalledAndCorrect = $true;
        }
        else {
            $Continue = $false
            foreach ($message in $installResult.ErrorMessages) {
                Write-InstallerError $message
            }
        }
    }
}
Write-WinSCPCopySnapshot -BizTalkInstallFolder $bizTalkInstallFolder -WinSCPEXEDownload $WinSCPEXEDownload -WinSCPDllDownload $WinSCPDllDownload -WinSCPTargetEXEExists $WinSCPTargetEXEExists -WinSCPDLLTargetExists $WinSCPDLLTargetExists -InstalledAndCorrect $btsWinSCPProductInstalledAndCorrect
  
      
$finalOutcome = Get-FinalExecutionOutcome -InstalledSuccessfully $btsWinSCPProductInstalledAndCorrect -WhatIf ([bool]$WhatIfPreference) -PrerequisiteFailure $PrerequisiteFailure -ContinueFlag $Continue
Write-InstallerFinalOutcome -Outcome $finalOutcome.Outcome -WinSCPVersion $winSCPVersion

