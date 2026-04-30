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
# Parameters
Param(
    [Parameter(
        Mandatory = $false
    )]
    [ValidateNotNullOrEmpty()]
    [string]$nugetDownloadFolder = (Get-Item Env:TEMP).Value + "\nuget",
    [Parameter(
        Mandatory = $false
    )]
    [switch]$ForceInstall,
    [Parameter(
        Mandatory = $false
    )]
    [string]$LogFolder,
    [Parameter(
        Mandatory = $false
    )]
    [ValidateSet('Info', 'Verbose', 'Debug')]
    [string]$LogLevel = 'Info',
    [Parameter(
        Mandatory = $false
    )]
    [switch]$EnableEventLog,
    [Parameter(
        Mandatory = $false
    )]
    [string]$EventLogName = 'Application',
    [Parameter(
        Mandatory = $false
    )]
    [string]$EventSource = 'BizTalkWinSCPInstaller',
    [Parameter(
        Mandatory = $false
    )]
    [switch]$CheckHash
)
# Import modules used by this installer workflow.
$coreModulePath = Join-Path $PSScriptRoot "src\InstallWinSCPForBizTalk.Core.psm1"
if (-not (Test-Path $coreModulePath)) {
    throw "Required core module was not found: $coreModulePath"
}
Import-Module $coreModulePath

$utilsModulePath = Join-Path $PSScriptRoot "src\InstallWinSCPForBizTalk.Utils.psm1"
Import-Module $utilsModulePath

$workflowModulePath = Join-Path $PSScriptRoot "src\InstallWinSCPForBizTalk.Workflow.psm1"
Import-Module $workflowModulePath

$initModulePath = Join-Path $PSScriptRoot "src\InstallWinSCPForBizTalk.Init.psm1"
Import-Module $initModulePath

# Read BizTalk registry values at script scope so test mocks can intercept them.
$bizTalkRegistryPath = 'HKLM:\SOFTWARE\Microsoft\BizTalk Server\3.0'
$bizTalkDetectionArgs = @{
    EnvironmentInstallPath = (Get-Item Env:BTSINSTALLPATH).Value
    RegistryInstallPath    = (Get-ItemPropertyValue $bizTalkRegistryPath -Name 'InstallPath')
    ProductCodeCurrent     = (Get-ItemPropertyValue $bizTalkRegistryPath -Name 'ProductCodeCurrent')
    ProductName            = (Get-ItemPropertyValue $bizTalkRegistryPath -Name 'ProductName')
    ProductVersion         = (Get-ItemPropertyValue $bizTalkRegistryPath -Name 'ProductVersion')
}

# Script bootstrap: initialize logging, workflow context, and supporting state.
$bootstrap = Initialize-InstallerBootstrap `
    -NuGetDownloadFolder $nugetDownloadFolder `
    -ForceInstall ([bool]$ForceInstall) `
    -LogFolder $LogFolder `
    -LogLevel $LogLevel `
    -EnableEventLog ([bool]$EnableEventLog) `
    -EventLogName $EventLogName `
    -EventSource $EventSource `
    -WhatIf ([bool]$WhatIfPreference) `
    -PSCmdlet $PSCmdlet `
    -DebugPreferenceValue ([string]$DebugPreference) `
    -VerbosePreferenceValue ([string]$VerbosePreference)

$workflowContext = $bootstrap.WorkflowContext

# Redefine InvokeWebRequest in script scope so test mocks can intercept Invoke-WebRequest.
$workflowContext['InvokeWebRequest'] = { param($Uri, $OutFile) Invoke-WebRequest -Uri $Uri -OutFile $OutFile }
$workflowContext['CheckHash'] = [bool]$CheckHash

# Phase 1: Environment detection and install target selection
Invoke-BizTalkDetectionPhase -Context $workflowContext @bizTalkDetectionArgs
Invoke-ForceInstallPrerequisitePhase -Context $workflowContext -WhatIf ([bool]$WhatIfPreference)
Invoke-CuDetectionPhase -Context $workflowContext

# Sync phase 1 outputs into script-scope variables used by later phases.
$PrerequisiteFailure = [bool]$workflowContext.PrerequisiteFailure

# Phase 2: Existing installation checks and version validation
Invoke-ExistingWinSCPCheckPhase -Context $workflowContext
Invoke-NonAdminWarningPhase -Context $workflowContext
Invoke-VersionValidationPhase -Context $workflowContext

# Sync phase 2 outputs before download/copy operations.
$btsWinSCPProductInstalledAndCorrect = [bool]$workflowContext.btsWinSCPProductInstalledAndCorrect

# Phase 3: Download preparation, package acquisition, and copy
Invoke-DownloadFolderPreparationPhase -Context $workflowContext
Invoke-NuGetDownloadPhase -Context $workflowContext
Invoke-WinSCPPackageDownloadPhase -Context $workflowContext
Invoke-WinSCPCopyPhase -Context $workflowContext
Invoke-WinSCPVerificationPhase -Context $workflowContext

# Sync phase 3 outcomes for final result calculation and reporting.
$Continue = [bool]$workflowContext.Continue
$winSCPVersion = $workflowContext.winSCPVersion
$btsWinSCPProductInstalledAndCorrect = [bool]$workflowContext.btsWinSCPProductInstalledAndCorrect

# Final outcome: single standardized success/warning/error path
$finalOutcome = Get-FinalExecutionOutcome -InstalledSuccessfully $btsWinSCPProductInstalledAndCorrect -WhatIf ([bool]$WhatIfPreference) -PrerequisiteFailure $PrerequisiteFailure -ContinueFlag $Continue
Write-InstallerFinalOutcome -Outcome $finalOutcome.Outcome -WinSCPVersion $winSCPVersion
Write-InstallerLogEntry -Level 'Info' -Message ("Installer execution completed with outcome '$($finalOutcome.Outcome)'.")
Disable-InstallerLogging

