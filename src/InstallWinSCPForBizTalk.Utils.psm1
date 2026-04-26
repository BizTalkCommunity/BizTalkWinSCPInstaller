# InstallWinSCPForBizTalk.Utils.psm1
# Shared output, snapshot, and semantic messaging helpers.

# ---------------------------------------------------------------------------
# Logging subsystem — module-scope state
# ---------------------------------------------------------------------------
$script:InstallerLogState = @{
    Enabled         = $false
    LogFile         = $null
    LogLevel        = 'Info'
    EventLogEnabled = $false
    EventLogName    = 'Application'
    EventSource     = 'BizTalkWinSCPInstaller'
}

# Returns the numeric rank of a log level for threshold comparison.
function Get-InstallerLogLevelRank {
    [OutputType([int])]
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Error', 'Info', 'Verbose', 'Debug')]
        [string]$Level
    )
    switch ($Level) {
        'Error'   { return 0 }
        'Info'    { return 1 }
        'Verbose' { return 2 }
        'Debug'   { return 3 }
        default   { throw "Unsupported log level: $Level" }
    }
}

# Returns $true when the given level is at or below the active log threshold.
function Test-InstallerLogLevelEnabled {
    [OutputType([bool])]
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Error', 'Info', 'Verbose', 'Debug')]
        [string]$Level
    )
    if (-not $script:InstallerLogState.Enabled) { return $false }
    return (Get-InstallerLogLevelRank -Level $Level) -le (Get-InstallerLogLevelRank -Level $script:InstallerLogState.LogLevel)
}

# Maps installer log levels to Windows Event Log entry types.
function Convert-InstallerEventEntryType {
    [OutputType([string])]
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Error', 'Info', 'Verbose', 'Debug')]
        [string]$Level
    )

    if ($Level -eq 'Error') { return 'Error' }
    return 'Information'
}

# Returns whether the configured event source already exists.
function Test-InstallerEventSourceExists {
    [OutputType([bool])]
    Param(
        [Parameter(Mandatory = $true)]
        [string]$EventSource
    )

    return [System.Diagnostics.EventLog]::SourceExists($EventSource)
}

# Registers a Windows Event Log source when missing.
function New-InstallerEventSource {
    Param(
        [Parameter(Mandatory = $true)]
        [string]$EventLogName,

        [Parameter(Mandatory = $true)]
        [string]$EventSource
    )

    New-EventLog -LogName $EventLogName -Source $EventSource
}

# Writes a single Windows Event Log entry for installer output.
function Write-InstallerEventLogRecord {
    Param(
        [Parameter(Mandatory = $true)]
        [string]$EventLogName,

        [Parameter(Mandatory = $true)]
        [string]$EventSource,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Error', 'Information')]
        [string]$EntryType,

        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    Write-EventLog -LogName $EventLogName -Source $EventSource -EntryType $EntryType -EventId 1000 -Category 0 -Message $Message
}

# Writes an entry to Windows Event Log when event logging is enabled.
function Write-InstallerEventLogEntry {
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Error', 'Info', 'Verbose', 'Debug')]
        [string]$Level,

        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    if (-not $script:InstallerLogState.EventLogEnabled) { return }

    try {
        if (-not (Test-InstallerEventSourceExists -EventSource $script:InstallerLogState.EventSource)) {
            New-InstallerEventSource -EventLogName $script:InstallerLogState.EventLogName -EventSource $script:InstallerLogState.EventSource
        }

        $entryType = Convert-InstallerEventEntryType -Level $Level
        Write-InstallerEventLogRecord -EventLogName $script:InstallerLogState.EventLogName -EventSource $script:InstallerLogState.EventSource -EntryType $entryType -Message $Message
    }
    catch {
        if (-not [string]::IsNullOrWhiteSpace($script:InstallerLogState.LogFile)) {
            $timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            $warning = "[$timestamp] [WARN] Event log write failed: $($_.Exception.Message)"
            Add-Content -Path $script:InstallerLogState.LogFile -Value $warning -Encoding UTF8
        }
    }
}

# Appends a single timestamped entry to the active log file when level qualifies.
function Write-InstallerLogEntry {
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Error', 'Info', 'Verbose', 'Debug')]
        [string]$Level,

        [Parameter(Mandatory = $true)]
        [string]$Message
    )
    if (-not (Test-InstallerLogLevelEnabled -Level $Level)) { return }
    $timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $line = "[$timestamp] [$($Level.ToUpper())] $Message"
    Add-Content -Path $script:InstallerLogState.LogFile -Value $line -Encoding UTF8
    Write-InstallerEventLogEntry -Level $Level -Message $Message
}

# Creates a timestamped log file and activates the logging sink for this session.
function Initialize-InstallerLogging {
    [OutputType([hashtable])]
    Param(
        [Parameter(Mandatory = $true)]
        [string]$LogFolder,

        [Parameter(Mandatory = $false)]
        [ValidateSet('Info', 'Verbose', 'Debug')]
        [string]$LogLevel = 'Info',

        [Parameter(Mandatory = $false)]
        [switch]$EnableEventLog,

        [Parameter(Mandatory = $false)]
        [string]$EventLogName = 'Application',

        [Parameter(Mandatory = $false)]
        [string]$EventSource = 'BizTalkWinSCPInstaller'
    )

    if (-not (Test-Path $LogFolder)) {
        New-Item -Path $LogFolder -ItemType Directory -Force | Out-Null
    }

    $timestamp = (Get-Date).ToString('yyyy-MM-dd-HHmmss')
    $baseName = "BizTalkWinSCPInstaller-$timestamp"
    $logFile = Join-Path $LogFolder "$baseName.log"
    $suffix = 1
    while (Test-Path $logFile) {
        $suffix++
        $logFile = Join-Path $LogFolder ('{0}-{1:D2}.log' -f $baseName, $suffix)
    }

    $script:InstallerLogState.Enabled         = $true
    $script:InstallerLogState.LogFile         = $logFile
    $script:InstallerLogState.LogLevel        = $LogLevel
    $script:InstallerLogState.EventLogEnabled = [bool]$EnableEventLog
    $script:InstallerLogState.EventLogName    = $EventLogName
    $script:InstallerLogState.EventSource     = $EventSource

    @(
        '================================================================================'
        ' BizTalk WinSCP Installer log'
        ('  Machine   : ' + $env:COMPUTERNAME)
        ('  User      : ' + $env:USERNAME)
        ('  PSVersion : ' + $PSVersionTable.PSVersion)
        ('  LogLevel  : ' + $LogLevel)
        ('  LogFolder : ' + $LogFolder)
        '================================================================================'
    ) | Set-Content -Path $logFile -Encoding UTF8

    Write-InstallerLogEntry -Level 'Info' -Message ("Support log file: $logFile")
    if ($EnableEventLog) {
        Write-InstallerLogEntry -Level 'Info' -Message ("Windows Event Log sink enabled: LogName='{0}', Source='{1}'" -f $EventLogName, $EventSource)
    }

    return @{
        LogFile         = $logFile
        LogLevel        = $LogLevel
        EventLogEnabled = [bool]$EnableEventLog
        EventLogName    = $EventLogName
        EventSource     = $EventSource
    }
}

# Deactivates the logging sink and releases the log file reference.
function Disable-InstallerLogging {
    $script:InstallerLogState.Enabled         = $false
    $script:InstallerLogState.LogFile         = $null
    $script:InstallerLogState.EventLogEnabled = $false
    $script:InstallerLogState.EventLogName    = 'Application'
    $script:InstallerLogState.EventSource     = 'BizTalkWinSCPInstaller'
}

# ---------------------------------------------------------------------------
# Console output helpers
# ---------------------------------------------------------------------------

# Write an installer error line.
function Write-InstallerError {
    <#
    .SYNOPSIS
    Writes an installer error line using standard error color styling.

    .PARAMETER ErrorMessage
    Error message text to display.
    #>
    Param([string] $ErrorMessage)
    Write-InstallerLogEntry -Level 'Info' -Message $ErrorMessage
    Write-Host -ForegroundColor Red "$ErrorMessage";
}

# Write an installer success/info line.
function Write-InstallerSuccess {
    <#
    .SYNOPSIS
    Writes an installer success/info line using standard success color styling.

    .PARAMETER SuccessMessage
    Success/info message text to display.
    #>
    Param([string] $SuccessMessage)
    Write-InstallerLogEntry -Level 'Info' -Message $SuccessMessage
    Write-Host -ForegroundColor Green "$SuccessMessage";
}

# Resolve shared banner lines used in installer output.
function Get-InstallerBannerLine {
    <#
    .SYNOPSIS
    Returns a standard banner delimiter line used by installer output blocks.

    .PARAMETER Name
    Banner style name: Hash, Bang, or Up.
    #>
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Hash', 'Bang', 'Up')]
        [string] $Name
    )

    switch ($Name) {
        'Hash' { return '##############################################################################' }
        'Bang' { return '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!' }
        'Up' { return '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^' }
        default { throw "Unsupported banner name: $Name" }
    }
}

# Write a delimiter-wrapped message block.
function Write-InstallerDelimitedMessage {
    <#
    .SYNOPSIS
    Writes a structured, delimiter-wrapped message block.

    .DESCRIPTION
    Outputs one or more lines wrapped by a selected delimiter style and severity
    level, with optional leading blank line for section separation.

    .PARAMETER MessageLines
    Message lines to emit inside the delimiter block.

    .PARAMETER Delimiter
    Delimiter style name: Hash, Bang, or Up.

    .PARAMETER Level
    Output level style: Success or Error.

    .PARAMETER LeadingNewLine
    Adds a leading newline before the opening delimiter when specified.
    #>
    Param(
        [Parameter(Mandatory = $true)]
        [string[]] $MessageLines,

        [ValidateSet('Hash', 'Bang', 'Up')]
        [string] $Delimiter = 'Hash',

        [ValidateSet('Success', 'Error')]
        [string] $Level = 'Success',

        [switch] $LeadingNewLine
    )

    $delimiterLine = Get-InstallerBannerLine -Name $Delimiter
    $firstDelimiter = if ($LeadingNewLine) { "`n$delimiterLine" } else { $delimiterLine }

    if ($Level -eq 'Error') {
        Write-InstallerError $firstDelimiter
        foreach ($line in $MessageLines) {
            Write-InstallerError $line
        }
        Write-InstallerError $delimiterLine
        return
    }

    Write-InstallerSuccess $firstDelimiter
    foreach ($line in $MessageLines) {
        Write-InstallerSuccess $line
    }
    Write-InstallerSuccess $delimiterLine
}

# Write a section header block with hash delimiters.
function Write-InstallerSectionHeader {
    <#
    .SYNOPSIS
    Writes a standard installer section header block.

    .PARAMETER Title
    Section title text.

    .PARAMETER LeadingNewLine
    Adds a leading newline before the opening delimiter when specified.
    #>
    Param(
        [Parameter(Mandatory = $true)]
        [string] $Title,

        [switch] $LeadingNewLine
    )

    Write-InstallerDelimitedMessage -MessageLines @($Title) -Delimiter 'Hash' -Level 'Success' -LeadingNewLine:$LeadingNewLine
}

# Write a bang-delimited error block.
function Write-InstallerBangError {
    <#
    .SYNOPSIS
    Writes a standardized bang-delimited error block.

    .PARAMETER MessageLines
    Error lines to emit inside the delimiter block.

    .PARAMETER LeadingNewLine
    Adds a leading newline before the opening delimiter when specified.
    #>
    Param(
        [Parameter(Mandatory = $true)]
        [string[]] $MessageLines,

        [switch] $LeadingNewLine
    )

    Write-InstallerDelimitedMessage -MessageLines $MessageLines -Delimiter 'Bang' -Level 'Error' -LeadingNewLine:$LeadingNewLine
}

# Render standardized final installer outcome messages.
function Write-InstallerFinalOutcome {
    <#
    .SYNOPSIS
    Renders standardized final installer outcome messaging.

    .PARAMETER Outcome
    Final outcome classification string.

    .PARAMETER WinSCPVersion
    Resolved WinSCP version associated with the run.
    #>
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string] $Outcome,

        [Parameter(Mandatory = $false)]
        [string] $WinSCPVersion
    )

    $bangString = Get-InstallerBannerLine -Name 'Bang'

    switch ($Outcome) {
        'Success' {
            Write-InstallerSuccess "`n$bangString"
            Write-InstallerSuccess "WinSCP $WinSCPVersion is installed."
            Write-InstallerSuccess "Microsoft BizTalk Server`'s SFTP Adapter will use this version of WinSCP."
            Write-InstallerSuccess $bangString
        }
        'DryRun' {
            Write-InstallerDelimitedMessage -MessageLines @(
                'The parameter -WhatIf was set and this script executed without making'
                'any changes and the output should be checked to determine if it would have '
                'run correctly.'
            ) -Delimiter 'Bang' -Level 'Success' -LeadingNewLine
        }
        'PrerequisiteFailure' {
            Write-InstallerDelimitedMessage -MessageLines @(
                'Installation did not run because one or more prerequisites were not met.'
                'Please address the prerequisite errors above and rerun the script.'
                'Exiting...'
            ) -Delimiter 'Bang' -Level 'Error' -LeadingNewLine
        }
        default {
            Write-InstallerDelimitedMessage -MessageLines @(
                'Something went wrong during installation and the installation did not work.'
                'Please inspect the errors above and resolve them.'
                'Exiting...'
            ) -Delimiter 'Bang' -Level 'Error' -LeadingNewLine
        }
    }
}

# Wrapper for debug state snapshots emitted by installer phases.
function Write-InstallerStateSnapshot {
    <#
    .SYNOPSIS
    Emits a structured debug/verbose snapshot of installer state.

    .PARAMETER Title
    Snapshot title heading.

    .PARAMETER State
    Ordered state hashtable to emit.
    #>
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Title,

        [Parameter(Mandatory = $true)]
        [System.Collections.IDictionary]$State
    )

    if ($null -eq $State) { return }

    if (-not [string]::IsNullOrWhiteSpace($Title)) {
        Write-Verbose $Title
        Write-Debug $Title
        Write-InstallerLogEntry -Level 'Verbose' -Message $Title
    }

    foreach ($entry in $State.GetEnumerator()) {
        $line = ('${0} = {1}' -f $entry.Key, $entry.Value)
        Write-Verbose $line
        Write-Debug $line
        Write-InstallerLogEntry -Level 'Verbose' -Message $line
    }
}

# Emit discovery snapshot for BizTalk install path and product metadata.
function Write-BizTalkSearchSnapshot {
    <#
    .SYNOPSIS
    Emits snapshot state for BizTalk install discovery results.

    .PARAMETER InstallFolder
    Resolved BizTalk install folder path.

    .PARAMETER InstallFolderExists
    Indicates whether resolved install folder exists.

    .PARAMETER ProductCodeCurrent
    BizTalk product code value.

    .PARAMETER ProductName
    BizTalk product name value.

    .PARAMETER ProductVersion
    BizTalk product version value.
    #>
    Param(
        [string]$InstallFolder,
        [bool]$InstallFolderExists,
        [string]$ProductCodeCurrent,
        [string]$ProductName,
        [string]$ProductVersion
    )

    Write-InstallerStateSnapshot -Title 'The result of the search for the BizTalk Server:' -State ([ordered]@{
        bizTalkInstallFolder       = $InstallFolder
        bizTalkInstallFolderExists = $InstallFolderExists
        bizTalkProductCodeCurrent  = $ProductCodeCurrent
        bizTalkProductName         = $ProductName
        bizTalkProductVersion      = $ProductVersion
    })
}

# Emit CU-detection snapshot showing resolved WinSCP target mapping.
function Write-BizTalkCuSearchSnapshot {
    <#
    .SYNOPSIS
    Emits snapshot state for BizTalk CU detection results.

    .PARAMETER WinSCPVersion
    Resolved WinSCP version.

    .PARAMETER KB
    Detected KB identifier.

    .PARAMETER CULabel
    Detected CU label.

    .PARAMETER CUFound
    Indicates whether a CU was detected.
    #>
    Param(
        [string]$WinSCPVersion,
        [string]$KB,
        [string]$CULabel,
        [bool]$CUFound
    )

    Write-InstallerStateSnapshot -Title 'The result of the search for the BizTalk Cumulative Update:' -State ([ordered]@{
        winSCPVersion = $WinSCPVersion
        btsKB         = $KB
        bizTalkCUVer  = $CULabel
        CUFound       = $CUFound
    })
}

# Emit snapshot for existing WinSCP file/version checks in target folder.
function Write-ExistingWinSCPSnapshot {
    <#
    .SYNOPSIS
    Emits snapshot state for existing WinSCP target validation checks.

    .PARAMETER ExeProductVersionInstalled
    Existing installed EXE product version.

    .PARAMETER DllProductVersionInstalled
    Existing installed DLL product version.

    .PARAMETER ProductVersionRequired
    Required WinSCP product version.

    .PARAMETER InstalledAndCorrect
    Indicates whether installed WinSCP is already correct.

    .PARAMETER TargetExePath
    Target EXE path in BizTalk folder.

    .PARAMETER TargetDllPath
    Target DLL path in BizTalk folder.
    #>
    Param(
        [string]$ExeProductVersionInstalled,
        [string]$DllProductVersionInstalled,
        [string]$ProductVersionRequired,
        [bool]$InstalledAndCorrect,
        [string]$TargetExePath,
        [string]$TargetDllPath
    )

    Write-InstallerStateSnapshot -Title 'Check for existing WinSCP results were:' -State ([ordered]@{
        btsWinSCPEXEProductVersionInstalled = $ExeProductVersionInstalled
        btsWinSCPDLLProductVersionInstalled = $DllProductVersionInstalled
        winSCPProductVersionRequired        = $ProductVersionRequired
        btsWinSCPProductInstalledAndCorrect = $InstalledAndCorrect
        btsTargetWinSCPExe                  = $TargetExePath
        btsTargetWinSCPDll                  = $TargetDllPath
    })
}

# Emit snapshot for download-folder readiness and staged artifact discovery.
function Write-TargetFolderPreparationSnapshot {
    <#
    .SYNOPSIS
    Emits snapshot state for download-folder preparation decisions.

    .PARAMETER NugetDownloadFolderAlreadyExists
    Indicates whether download folder existed prior to preparation.

    .PARAMETER NugetDownloadFolderExists
    Indicates whether download folder exists after preparation.

    .PARAMETER WinSCPEXEDownload
    Resolved WinSCP EXE package path.

    .PARAMETER WinSCPDllDownload
    Resolved WinSCP DLL package path.

    .PARAMETER WinSCPEXEDownloadAlreadyExists
    Indicates whether package EXE already existed.

    .PARAMETER WinSCPDllDownloadAlreadyExists
    Indicates whether package DLL already existed.
    #>
    Param(
        [bool]$NugetDownloadFolderAlreadyExists,
        [bool]$NugetDownloadFolderExists,
        [string]$WinSCPEXEDownload,
        [string]$WinSCPDllDownload,
        [bool]$WinSCPEXEDownloadAlreadyExists,
        [bool]$WinSCPDllDownloadAlreadyExists
    )

    Write-InstallerStateSnapshot -Title 'Check and then potential creation of the target folder results:' -State ([ordered]@{
        nugetDownloadFolderAlreadyExists = $NugetDownloadFolderAlreadyExists
        nugetDownloadFolderExists        = $NugetDownloadFolderExists
        WinSCPEXEDownload                = $WinSCPEXEDownload
        WinSCPDllDownload                = $WinSCPDllDownload
        WinSCPEXEDownloadAlreadyExists   = $WinSCPEXEDownloadAlreadyExists
        WinSCPDllDownloadAlreadyExists   = $WinSCPDllDownloadAlreadyExists
    })
}

# Emit snapshot for NuGet acquisition state and source/target paths.
function Write-NuGetDownloadSnapshot {
    <#
    .SYNOPSIS
    Emits snapshot state for NuGet acquisition/readiness.

    .PARAMETER TargetNugetExeAlreadyExists
    Indicates whether nuget.exe existed before acquisition.

    .PARAMETER TargetNugetExeExists
    Indicates whether nuget.exe exists after acquisition.

    .PARAMETER SourceNugetExe
    NuGet download source URI.

    .PARAMETER TargetNugetExe
    NuGet target file path.
    #>
    Param(
        [bool]$TargetNugetExeAlreadyExists,
        [bool]$TargetNugetExeExists,
        [string]$SourceNugetExe,
        [string]$TargetNugetExe
    )

    Write-InstallerStateSnapshot -Title 'Check and then potential download NuGet results:' -State ([ordered]@{
        targetNugetExeAlreadyExists = $TargetNugetExeAlreadyExists
        targetNugetExeExists        = $TargetNugetExeExists
        sourceNugetExe              = $SourceNugetExe
        targetNugetExe              = $TargetNugetExe
    })
}

# Emit snapshot for WinSCP package acquisition and extracted artifact presence.
function Write-WinSCPDownloadSnapshot {
    <#
    .SYNOPSIS
    Emits snapshot state for WinSCP package acquisition.

    .PARAMETER GetWinSCP
    Command preview used for WinSCP package acquisition.

    .PARAMETER WinSCPEXEAlreadyExists
    Indicates whether package EXE existed before acquisition.

    .PARAMETER WinSCPDLLAlreadyExists
    Indicates whether package DLL existed before acquisition.

    .PARAMETER WinSCPDllDownload
    Resolved package DLL path.

    .PARAMETER WinSCPEXEExists
    Indicates whether package EXE exists after acquisition.

    .PARAMETER WinSCPDLLExists
    Indicates whether package DLL exists after acquisition.
    #>
    Param(
        [string]$GetWinSCP,
        [bool]$WinSCPEXEAlreadyExists,
        [bool]$WinSCPDLLAlreadyExists,
        [string]$WinSCPDllDownload,
        [bool]$WinSCPEXEExists,
        [bool]$WinSCPDLLExists
    )

    Write-InstallerStateSnapshot -Title 'Check and then potential download WinSCP results:' -State ([ordered]@{
        getWinSCP              = $GetWinSCP
        WinSCPEXEAlreadyExists = $WinSCPEXEAlreadyExists
        WinSCPDLLAlreadyExists = $WinSCPDLLAlreadyExists
        WinSCPDllDownload      = $WinSCPDllDownload
        WinSCPEXEExists        = $WinSCPEXEExists
        WinSCPDLLExists        = $WinSCPDLLExists
    })
}

# Emit snapshot for copy/install results into the BizTalk installation folder.
function Write-WinSCPCopySnapshot {
    <#
    .SYNOPSIS
    Emits snapshot state for WinSCP copy/install result validation.

    .PARAMETER BizTalkInstallFolder
    Target BizTalk install folder path.

    .PARAMETER WinSCPEXEDownload
    Source WinSCP EXE path.

    .PARAMETER WinSCPDllDownload
    Source WinSCP DLL path.

    .PARAMETER WinSCPTargetEXEExists
    Indicates whether target EXE exists after copy.

    .PARAMETER WinSCPDLLTargetExists
    Indicates whether target DLL exists after copy.

    .PARAMETER InstalledAndCorrect
    Indicates whether final install state is considered correct.
    #>
    Param(
        [string]$BizTalkInstallFolder,
        [string]$WinSCPEXEDownload,
        [string]$WinSCPDllDownload,
        [bool]$WinSCPTargetEXEExists,
        [bool]$WinSCPDLLTargetExists,
        [bool]$InstalledAndCorrect
    )

    Write-InstallerStateSnapshot -Title 'Check and then potential copy WinSCP to BizTalk results:' -State ([ordered]@{
        bizTalkInstallFolder                = $BizTalkInstallFolder
        WinSCPEXEDownload                   = $WinSCPEXEDownload
        WinSCPDllDownload                   = $WinSCPDllDownload
        WinSCPTargetEXEExists               = $WinSCPTargetEXEExists
        WinSCPDLLTargetExists               = $WinSCPDLLTargetExists
        btsWinSCPProductInstalledAndCorrect = $InstalledAndCorrect
    })
}

# Semantic installer message wrappers that keep workflow code concise.

# Warn that BTS path is missing in environment and registry fallback is being used.
function Write-BizTalkRegistryFallbackNotice {
    <#
    .SYNOPSIS
    Writes standard warning when BTSINSTALLPATH is missing and registry fallback is used.
    #>
    Write-InstallerError 'The Env:BTSINSTALLPATH does not exist, checking to see if the path is in the registry HKLM:\SOFTWARE\Microsoft\BizTalk Server\3.0@InstallPath'
}

# Emit terminal message when BizTalk installation cannot be located.
function Write-BizTalkNotLocatedError {
    <#
    .SYNOPSIS
    Writes standard terminal error messaging for missing BizTalk installation.
    #>
    Write-InstallerError 'Microsoft BizTalk Server was not located by checking the environment variable BTSINSTALLPATH and the Registry key for BizTalk, exiting the process'
    Write-InstallerError 'Please confirm that Microsoft BizTalk Server is installed on this system'
}

# Emit standardized success lines for detected BizTalk product details.
function Write-BizTalkLocatedSuccess {
    <#
    .SYNOPSIS
    Writes standardized success messaging for detected BizTalk product details.

    .PARAMETER ProductName
    Detected BizTalk product name.

    .PARAMETER ProductVersion
    Detected BizTalk product version.

    .PARAMETER InstallFolder
    Detected BizTalk install folder.
    #>
    Param(
        [Parameter(Mandatory = $true)]
        [string]$ProductName,
        [Parameter(Mandatory = $true)]
        [string]$ProductVersion,
        [Parameter(Mandatory = $true)]
        [string]$InstallFolder
    )

    Write-InstallerSuccess "Located $ProductName version $ProductVersion."
    Write-InstallerSuccess 'Microsoft BizTalk Server is installed.'
    Write-InstallerSuccess "Located in the '$InstallFolder' folder."
}

# Emit CU-detection kickoff message for the detected BizTalk major version.
function Write-BizTalkCuDetectionStart {
    <#
    .SYNOPSIS
    Writes standardized CU-detection kickoff lines.

    .PARAMETER BizTalkVersion
    Detected BizTalk major version label.
    #>
    Param(
        [Parameter(Mandatory = $true)]
        [string]$BizTalkVersion
    )

    Write-InstallerSuccess "Detected Microsoft BizTalk Server $BizTalkVersion."
    Write-InstallerSuccess 'Testing to see which Cumulative Update is installed'
}

# Emit standardized not-installed notice prior to package acquisition flow.
function Write-WinSCPNotInstalledNotice {
    <#
    .SYNOPSIS
    Writes standardized message for missing required WinSCP installation.

    .PARAMETER WinSCPVersion
    Required WinSCP version that is not currently installed.
    #>
    Param(
        [Parameter(Mandatory = $true)]
        [string]$WinSCPVersion
    )

    $bangString = Get-InstallerBannerLine -Name 'Bang'
    Write-InstallerSuccess $bangString
    Write-InstallerSuccess "WinSCP $WinSCPVersion is not installed in the"
    Write-InstallerSuccess 'Microsoft BizTalk Server folder and needs to be installed.'
    Write-InstallerSuccess $bangString
}
