#####################################################################
# InstallWinSCPForBizTalk.Utils.psm1
#
# Shared output utility functions for the WinSCP BizTalk installer
#####################################################################

#####################################################################
# Function to write an error
#####################################################################
function Write-InstallerError {
    Param([string] $ErrorMessage)
    Write-Host -ForegroundColor Red "$ErrorMessage";
}

#####################################################################
# Function to write success
#####################################################################
function Write-InstallerSuccess {
    Param([string] $SuccessMessage)
    Write-Host -ForegroundColor Green "$SuccessMessage";
}

#####################################################################
# Function to resolve shared banner lines used in script output
#####################################################################
function Get-InstallerBannerLine {
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

#####################################################################
# Function to write a delimited message block
#####################################################################
function Write-InstallerDelimitedMessage {
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

#####################################################################
# Function to write a section header block with hash delimiters
#####################################################################
function Write-InstallerSectionHeader {
    Param(
        [Parameter(Mandatory = $true)]
        [string] $Title,

        [switch] $LeadingNewLine
    )

    Write-InstallerDelimitedMessage -MessageLines @($Title) -Delimiter 'Hash' -Level 'Success' -LeadingNewLine:$LeadingNewLine
}

#####################################################################
# Function to write a bang-delimited error block
#####################################################################
function Write-InstallerBangError {
    Param(
        [Parameter(Mandatory = $true)]
        [string[]] $MessageLines,

        [switch] $LeadingNewLine
    )

    Write-InstallerDelimitedMessage -MessageLines $MessageLines -Delimiter 'Bang' -Level 'Error' -LeadingNewLine:$LeadingNewLine
}

#####################################################################
# Function to render final installer outcome messages
#####################################################################
function Write-InstallerFinalOutcome {
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string] $Outcome,

        [Parameter(Mandatory = $false)]
        [string] $WinSCPVersion
    )

    $bangString = Get-InstallerBannerLine -Name 'Bang'
    $upString = Get-InstallerBannerLine -Name 'Up'

    switch ($Outcome) {
        'Success' {
            Write-InstallerSuccess "`n$bangString"
            Write-InstallerSuccess "WinSCP $WinSCPVersion is installed."
            Write-InstallerSuccess "Microsoft BizTalk Server`'s SFTP Adapter will use this version of WinSCP."
            Write-InstallerSuccess $upString
        }
        'DryRun' {
            Write-InstallerDelimitedMessage -MessageLines @(
                'The parameter -WhatIf was set and this script executed without making'
                'any changes and the output should checked to determine if it would have '
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

#####################################################################
# Debug utility class for compact state snapshot logging
#####################################################################
class InstallerDebugUtility {
    static [void] WriteState([string]$Title, [hashtable]$State) {
        if ($null -eq $State) {
            return
        }

        if (-not [string]::IsNullOrWhiteSpace($Title)) {
            Write-Verbose $Title
            Write-Debug $Title
        }

        foreach ($entry in $State.GetEnumerator()) {
            $line = ('${0} = {1}' -f $entry.Key, $entry.Value)
            Write-Verbose $line
            Write-Debug $line
        }
    }
}

#####################################################################
# Function wrapper for debug state snapshots from installer script
#####################################################################
function Write-InstallerStateSnapshot {
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Title,

        [Parameter(Mandatory = $true)]
        [hashtable]$State
    )

    [InstallerDebugUtility]::WriteState($Title, $State)
}

function Write-BizTalkSearchSnapshot {
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

function Write-BizTalkCuSearchSnapshot {
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

function Write-ExistingWinSCPSnapshot {
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

function Write-TargetFolderPreparationSnapshot {
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

function Write-NuGetDownloadSnapshot {
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

function Write-WinSCPDownloadSnapshot {
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

function Write-WinSCPCopySnapshot {
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

#####################################################################
# Semantic installer message wrappers to reduce script output noise
#####################################################################
function Write-BizTalkRegistryFallbackNotice {
    Write-InstallerError 'The Env:BTSINSTALLPATH doesn`t exist, checking to see if the path is in the registry HKLM:\SOFTWARE\Microsoft\BizTalk Server\3.0@InstallPath'
}

function Write-BizTalkNotLocatedError {
    Write-InstallerError 'Microsoft BizTalk Server was not located by checking the environment variable BTSINSTALLPATH and the Registry key for BizTalk, exiting the process'
    Write-InstallerError 'Please confirm that Microsoft BizTalk Server is installed on this system'
}

function Write-BizTalkLocatedSuccess {
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

function Write-BizTalkCuDetectionStart {
    Param(
        [Parameter(Mandatory = $true)]
        [string]$BizTalkVersion
    )

    Write-InstallerSuccess "Detected Microsoft BizTalk Server $BizTalkVersion."
    Write-InstallerSuccess 'Testing to see which Cumulative Update is installed'
}

function Write-WinSCPNotInstalledNotice {
    Param(
        [Parameter(Mandatory = $true)]
        [string]$WinSCPVersion
    )

    $bangString = Get-InstallerBannerLine -Name 'Bang'
    Write-InstallerSuccess $bangString
    Write-InstallerSuccess "WinSCP $WinSCPVersion is NOT installed in the"
    Write-InstallerSuccess 'Microsoft BizTalk Server folder and needs to be installed.'
}
