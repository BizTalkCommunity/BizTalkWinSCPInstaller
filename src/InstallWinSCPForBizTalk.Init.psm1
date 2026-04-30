# InstallWinSCPForBizTalk.Init.psm1
# Bootstrap/initialization helpers for installer startup.

function Initialize-InstallerBootstrap {
    <#
    .SYNOPSIS
    Builds the initial installer workflow context and BizTalk detection arguments.

    .DESCRIPTION
    Performs the non-phase startup preparation used by InstallWinSCPForBizTalk.ps1,
    including logging setup, preference-aware log level selection, banner token
    resolution, and BizTalk discovery argument collection.

    .PARAMETER NuGetDownloadFolder
    Folder used for NuGet and WinSCP package acquisition.

    .PARAMETER ForceInstall
    Indicates whether reinstall should be forced.

    .PARAMETER LogFolder
    Folder to write installer logs to.

    .PARAMETER LogLevel
    Requested base log verbosity.

    .PARAMETER EnableEventLog
    Enables forwarding installer logs to Windows Event Log.

    .PARAMETER EventLogName
    Event Log name used when event logging is enabled.

    .PARAMETER EventSource
    Event source name used when event logging is enabled.

    .PARAMETER WhatIf
    Indicates whether script execution is in WhatIf mode.

    .PARAMETER PSCmdlet
    Current script cmdlet context for ShouldProcess delegation.

    .PARAMETER DebugPreferenceValue
    Caller debug preference used to auto-promote effective log level.

    .PARAMETER VerbosePreferenceValue
    Caller verbose preference used to auto-promote effective log level.

    #>
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$NuGetDownloadFolder,

        [Parameter(Mandatory = $true)]
        [bool]$ForceInstall,

        [Parameter(Mandatory = $false)]
        [string]$LogFolder,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Info', 'Verbose', 'Debug')]
        [string]$LogLevel,

        [Parameter(Mandatory = $true)]
        [bool]$EnableEventLog,

        [Parameter(Mandatory = $true)]
        [string]$EventLogName,

        [Parameter(Mandatory = $true)]
        [string]$EventSource,

        [Parameter(Mandatory = $true)]
        [bool]$WhatIf,

        [Parameter(Mandatory = $true)]
        $PSCmdlet,

        [Parameter(Mandatory = $true)]
        [string]$DebugPreferenceValue,

        [Parameter(Mandatory = $true)]
        [string]$VerbosePreferenceValue
    )

    $continue = $true
    $prerequisiteFailure = $false

    if ([string]::IsNullOrWhiteSpace($LogFolder)) {
        $LogFolder = Join-Path (Join-Path (Get-Item Env:TEMP).Value 'BizTalkWinSCPInstaller') 'logs'
    }

    $effectiveLogLevel = $LogLevel
    if ($DebugPreferenceValue -ne 'SilentlyContinue') {
        $effectiveLogLevel = 'Debug'
    }
    elseif ($VerbosePreferenceValue -ne 'SilentlyContinue' -and $effectiveLogLevel -eq 'Info') {
        $effectiveLogLevel = 'Verbose'
    }

    $loggingSession = Initialize-InstallerLogging -LogFolder $LogFolder -LogLevel $effectiveLogLevel -EnableEventLog:$EnableEventLog -EventLogName $EventLogName -EventSource $EventSource

    $isAdministrator = Test-IsAdministrator
    Write-InstallerLogEntry -Level 'Info' -Message 'Installer execution started.'
    Write-InstallerLogEntry -Level 'Info' -Message ("Parameters: NuGetDownloadFolder='{0}'; ForceInstall={1}; WhatIf={2}; EventLogEnabled={3}" -f $NuGetDownloadFolder, $ForceInstall, $WhatIf, $EnableEventLog)

    $winSCPexeFile = 'WinSCP.exe'
    $winSCPdllFile = 'WinSCPnet.dll'
    $hashString = Get-InstallerBannerLine -Name 'Hash'
    $bangString = Get-InstallerBannerLine -Name 'Bang'

    $workflowContext = @{
        # Flow state
        Continue            = $continue
        PrerequisiteFailure = $prerequisiteFailure
        ForceInstall        = $ForceInstall
        WhatIf              = $WhatIf
        isAdministrator     = $isAdministrator

        # Runtime operations
        psCmdlet            = $PSCmdlet

        # Package/file settings
        nugetDownloadFolder = $NuGetDownloadFolder
        winSCPexeFile       = $winSCPexeFile
        winSCPdllFile       = $winSCPdllFile

        # Console output tokens
        hashString          = $hashString
        bangString          = $bangString

        # Logging settings
        logFilePath         = $loggingSession.LogFile
        logLevel            = $loggingSession.LogLevel
        eventLogEnabled     = [bool]$loggingSession.EventLogEnabled
        eventLogName        = $loggingSession.EventLogName
        eventSource         = $loggingSession.EventSource
    }

    return [pscustomobject]@{
        WorkflowContext = $workflowContext
        LogFolder       = $LogFolder
    }
}

Export-ModuleMember -Function Initialize-InstallerBootstrap
