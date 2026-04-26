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
    [string]$LogLevel = 'Info'
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

# Initialize shared flow flags.
$Continue = $true;
$PrerequisiteFailure = $false;

if ([string]::IsNullOrWhiteSpace($LogFolder)) {
    $LogFolder = Join-Path (Join-Path (Get-Item Env:TEMP).Value 'BizTalkWinSCPInstaller') 'logs'
}

$effectiveLogLevel = $LogLevel
if ($DebugPreference -ne 'SilentlyContinue') {
    $effectiveLogLevel = 'Debug'
}
elseif ($VerbosePreference -ne 'SilentlyContinue' -and $effectiveLogLevel -eq 'Info') {
    $effectiveLogLevel = 'Verbose'
}

$loggingSession = Initialize-InstallerLogging -LogFolder $LogFolder -LogLevel $effectiveLogLevel

$isAdministrator = Test-IsAdministrator
Write-InstallerLogEntry -Level 'Info' -Message 'Installer execution started.'
Write-InstallerLogEntry -Level 'Info' -Message ("Parameters: NuGetDownloadFolder='{0}'; ForceInstall={1}; WhatIf={2}" -f $nugetDownloadFolder, [bool]$ForceInstall, [bool]$WhatIfPreference)
# Default WinSCP configuration and package layout notes.
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

$winSCPVersion = "5.7.7"
$winSCPexeFile = "WinSCP.exe";
$winSCPdllFile = "WinSCPnet.dll";
$hashString = Get-InstallerBannerLine -Name 'Hash'
$bangString = Get-InstallerBannerLine -Name 'Bang'

# Shared mutable context passed across workflow phases.
# Each phase updates this hashtable so the main script can remain a linear orchestrator.
$workflowContext = @{
    Continue            = $Continue
    PrerequisiteFailure = $PrerequisiteFailure
    isAdministrator     = $isAdministrator
    ForceInstall        = [bool]$ForceInstall
    WhatIf              = [bool]$WhatIfPreference
    winSCPexeFile       = $winSCPexeFile
    winSCPdllFile       = $winSCPdllFile
    hashString          = $hashString
    bangString          = $bangString
    nugetDownloadFolder = $nugetDownloadFolder
    psCmdlet            = $PSCmdlet
    InvokeWebRequest    = { param($Uri, $OutFile) Invoke-WebRequest -Uri $Uri -OutFile $OutFile }
    logFilePath         = $loggingSession.LogFile
    logLevel            = $loggingSession.LogLevel
}

$bizTalkInstallFolderFromEnv = (Get-Item Env:BTSINSTALLPATH).Value
$bizTalkInstallFolderFromRegistry = (get-itemPropertyValue 'HKLM:\SOFTWARE\Microsoft\BizTalk Server\3.0' -Name 'InstallPath')
$bizTalkProductCodeCurrent = (get-itemPropertyValue 'HKLM:\SOFTWARE\Microsoft\BizTalk Server\3.0' -Name 'ProductCodeCurrent')
$bizTalkProductName = (get-itemPropertyValue 'HKLM:\SOFTWARE\Microsoft\BizTalk Server\3.0' -Name 'ProductName')
$bizTalkProductVersion = (get-itemPropertyValue 'HKLM:\SOFTWARE\Microsoft\BizTalk Server\3.0' -Name 'ProductVersion')

# Phase 1: Environment detection and install target selection
Invoke-BizTalkDetectionPhase -Context $workflowContext -EnvironmentInstallPath $bizTalkInstallFolderFromEnv -RegistryInstallPath $bizTalkInstallFolderFromRegistry -ProductCodeCurrent $bizTalkProductCodeCurrent -ProductName $bizTalkProductName -ProductVersion $bizTalkProductVersion
Invoke-ForceInstallPrerequisitePhase -Context $workflowContext -WhatIf ([bool]$WhatIfPreference)
Invoke-CuDetectionPhase -Context $workflowContext

# Sync phase 1 outputs into script-scope variables used by later phases.
$Continue = [bool]$workflowContext.Continue
$PrerequisiteFailure = [bool]$workflowContext.PrerequisiteFailure
$bizTalkInstallFolder = $workflowContext.bizTalkInstallFolder
$bizTalkInstallFolderExists = [bool]$workflowContext.bizTalkInstallFolderExists
$bizTalkProductCodeCurrent = $workflowContext.bizTalkProductCodeCurrent
$bizTalkProductName = $workflowContext.bizTalkProductName
$bizTalkProductVersion = $workflowContext.bizTalkProductVersion
$BizTalkVersion = $workflowContext.BizTalkVersion
$winSCPVersion = $workflowContext.winSCPVersion
$btsKB = $workflowContext.btsKB
$bizTalkCUVer = $workflowContext.bizTalkCUVer
$CUFound = [bool]$workflowContext.CUFound

# Phase 2: Existing installation checks and version validation
Invoke-ExistingWinSCPCheckPhase -Context $workflowContext
Invoke-NonAdminWarningPhase -Context $workflowContext
Invoke-VersionValidationPhase -Context $workflowContext

# Sync phase 2 outputs before download/copy operations.
$Continue = [bool]$workflowContext.Continue
$winSCPVersion = $workflowContext.winSCPVersion
$btsWinSCPEXEProductVersionInstalled = $workflowContext.btsWinSCPEXEProductVersionInstalled
$btsWinSCPDLLProductVersionInstalled = $workflowContext.btsWinSCPDLLProductVersionInstalled
$btsWinSCPProductInstalledAndCorrect = [bool]$workflowContext.btsWinSCPProductInstalledAndCorrect
$installExecutionPlan = $workflowContext.installExecutionPlan
$winSCPProductVersionRequired = $workflowContext.winSCPProductVersionRequired
$btsTargetWinSCPExe = $workflowContext.btsTargetWinSCPExe
$btsTargetWinSCPDll = $workflowContext.btsTargetWinSCPDll

# Phase 3: Download preparation, package acquisition, and copy
Invoke-DownloadFolderPreparationPhase -Context $workflowContext
Invoke-NuGetDownloadPhase -Context $workflowContext
Invoke-WinSCPPackageDownloadPhase -Context $workflowContext
Invoke-WinSCPCopyPhase -Context $workflowContext

# Sync phase 3 outcomes for final result calculation and reporting.
$Continue = [bool]$workflowContext.Continue
$winSCPVersion = $workflowContext.winSCPVersion
$btsWinSCPProductInstalledAndCorrect = [bool]$workflowContext.btsWinSCPProductInstalledAndCorrect
$WinSCPEXEDownload = $workflowContext.WinSCPEXEDownload
$WinSCPDllDownload = $workflowContext.WinSCPDllDownload
$WinSCPTargetEXEExists = [bool]$workflowContext.WinSCPTargetEXEExists
$WinSCPDLLTargetExists = [bool]$workflowContext.WinSCPDLLTargetExists

# Final outcome: single standardized success/warning/error path
$finalOutcome = Get-FinalExecutionOutcome -InstalledSuccessfully $btsWinSCPProductInstalledAndCorrect -WhatIf ([bool]$WhatIfPreference) -PrerequisiteFailure $PrerequisiteFailure -ContinueFlag $Continue
Write-InstallerFinalOutcome -Outcome $finalOutcome.Outcome -WinSCPVersion $winSCPVersion
Write-InstallerLogEntry -Level 'Info' -Message ("Installer execution completed with outcome '$($finalOutcome.Outcome)'.")
Disable-InstallerLogging

