# InstallWinSCPForBizTalk.Core.psm1
# Core decision and validation helpers used by the installer workflow.

# Resolve WinSCP package layout from extracted NuGet files.
function Resolve-WinSCPPackageLayout {
    <#
    .SYNOPSIS
    Resolves WinSCP.exe and WinSCPnet.dll paths from an extracted NuGet package.

    .DESCRIPTION
    Searches common WinSCP NuGet package folder structures for EXE and DLL files.
    Supports both known paths and fallback recursive search.

    .PARAMETER PackageRoot
    Root path of the extracted WinSCP package (e.g., C:\temp\nuget\WinSCP.5.19.2)

    .PARAMETER ExeFileName
    Name of the EXE file to find (typically WinSCP.exe)

    .PARAMETER DllFileName
    Name of the DLL file to find (typically WinSCPnet.dll)

    .OUTPUTS
    [pscustomobject] with properties ExePath, DllPath, and IsResolved

    .EXAMPLE
    Resolve-WinSCPPackageLayout -PackageRoot "C:\temp\WinSCP.5.19.2" -ExeFileName "WinSCP.exe" -DllFileName "WinSCPnet.dll"
    #>
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string] $PackageRoot,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string] $ExeFileName,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
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
    # Keep both netstandard2.0 and netstandard for historical WinSCP NuGet layouts.
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

# Search for specific BizTalk cumulative updates by KB value.
function Search-BTSCumulativeUpdate {
    <#
    .SYNOPSIS
    Detects whether a specific BizTalk CU KB is present in uninstall entries.
    #>
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string] $CumulativeUpdateID,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
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

        $normalized = ($name -replace '\s+', ' ').Trim()
        $hasBizTalkVersion = $normalized -match "(?i)\bBizTalk\b.*\b$BizTalkVersion\b"
        $hasKB = $normalized -match "(?i)\bKB\D*$CumulativeUpdateID\b"
        return ($hasBizTalkVersion -and $hasKB)
    })
}

# Detect the highest CU from uninstall DisplayName entries.
function Get-BTSCumulativeUpdateByDisplayName {
    <#
    .SYNOPSIS
    Returns the most recent detected BizTalk CU based on uninstall DisplayName.
    #>
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
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

# Test whether the current PowerShell session runs elevated.
function Test-IsAdministrator {
    <#
    .SYNOPSIS
    Returns true when the current identity is in the local Administrators role.

    .DESCRIPTION
    Supports dependency injection for unit tests by allowing callers to provide
    custom scriptblocks for identity and principal construction.
    #>
    Param(
        [scriptblock]$GetCurrentIdentity = { [Security.Principal.WindowsIdentity]::GetCurrent() },
        [scriptblock]$NewPrincipal = {
            param($Identity)
            New-Object Security.Principal.WindowsPrincipal($Identity)
        }
    )

    $currentIdentity = & $GetCurrentIdentity
    $principal = & $NewPrincipal $currentIdentity
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Evaluate install-gating behavior from elevation and execution flags.
function Get-InstallExecutionPlan {
    <#
    .SYNOPSIS
    Evaluates whether install steps should proceed based on current flags/state.
    #>
    Param(
        [Parameter(Mandatory = $true)]
        [bool]$IsAdministrator,
        [Parameter(Mandatory = $true)]
        [bool]$ForceInstall,
        [Parameter(Mandatory = $true)]
        [bool]$WhatIf,
        [Parameter(Mandatory = $true)]
        [bool]$AlreadyInstalledCorrect
    )

    $canProceed = $true
    $prerequisiteFailure = $false
    $shouldSkipBecauseInstalled = $false
    $shouldReinstall = $false
    $requiresElevationWarning = $false

    # Match script behavior: non-admin + ForceInstall is a hard stop unless WhatIf.
    if (-not $IsAdministrator -and $ForceInstall) {
        if (-not $WhatIf) {
            $canProceed = $false
            $prerequisiteFailure = $true
        }
        else {
            $requiresElevationWarning = $true
        }
    }

    if ($AlreadyInstalledCorrect) {
        if ($ForceInstall) {
            $shouldReinstall = $true
        }
        else {
            $canProceed = $false
            $shouldSkipBecauseInstalled = $true
        }
    }

    if ($canProceed -and -not $IsAdministrator -and -not $ForceInstall) {
        $requiresElevationWarning = $true
    }

    return [pscustomobject]@{
        CanProceed = $canProceed
        PrerequisiteFailure = $prerequisiteFailure
        ShouldSkipBecauseInstalled = $shouldSkipBecauseInstalled
        ShouldReinstall = $shouldReinstall
        RequiresElevationWarning = $requiresElevationWarning
    }
}

# Validate WinSCP version string values.
function Test-WinSCPVersionString {
    <#
    .SYNOPSIS
    Validates a WinSCP version string and returns structured results.
    #>
    Param(
        [string]$WinSCPVersion
    )

    $parsedVersion = $null

    if ([string]::IsNullOrWhiteSpace($WinSCPVersion)) {
        return [pscustomobject]@{
            IsValid = $false
            ErrorCode = "MissingVersion"
            ParsedVersion = $null
        }
    }

    if (-not [version]::TryParse($WinSCPVersion, [ref]$parsedVersion)) {
        return [pscustomobject]@{
            IsValid = $false
            ErrorCode = "InvalidFormat"
            ParsedVersion = $null
        }
    }

    return [pscustomobject]@{
        IsValid = $true
        ErrorCode = "None"
        ParsedVersion = $parsedVersion
    }
}

# Resolve BizTalk install folder from environment/registry candidates.
function Resolve-BizTalkInstallFolder {
    <#
    .SYNOPSIS
    Selects the BizTalk install folder candidate from environment and registry.

    .DESCRIPTION
    Prefers the environment value when present; otherwise falls back to registry.
    Returns the selected path and basic source/existence metadata.
    #>
    Param(
        [string]$EnvironmentInstallPath,
        [string]$RegistryInstallPath
    )

    $selectedPath = $null
    $source = 'None'

    if (-not [string]::IsNullOrWhiteSpace($EnvironmentInstallPath)) {
        $selectedPath = $EnvironmentInstallPath
        $source = 'Environment'
    }
    elseif (-not [string]::IsNullOrWhiteSpace($RegistryInstallPath)) {
        $selectedPath = $RegistryInstallPath
        $source = 'Registry'
    }

    return [pscustomobject]@{
        InstallPath = $selectedPath
        Source = $source
        IsFound = -not [string]::IsNullOrWhiteSpace($selectedPath)
        Exists = [bool]($selectedPath -and (Test-Path $selectedPath))
    }
}

# Evaluate whether NuGet should be downloaded or reused.
function Get-NuGetDownloadPlan {
    <#
    .SYNOPSIS
    Returns a plan object for NuGet acquisition behavior.
    #>
    Param(
        [Parameter(Mandatory = $true)]
        [bool]$TargetNugetExeAlreadyExists,
        [Parameter(Mandatory = $true)]
        [bool]$ForceInstall,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$SourceNugetExe,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$TargetNugetExe
    )

    $shouldDownload = (-not $TargetNugetExeAlreadyExists -or $ForceInstall)

    return [pscustomobject]@{
        ShouldDownload = $shouldDownload
        ShouldReuseExisting = (-not $shouldDownload)
        ShouldProcessTarget = "$SourceNugetExe -OutFile $TargetNugetExe"
    }
}

# Evaluate whether the WinSCP package should be downloaded or reused.
function Get-WinSCPPackageDownloadPlan {
    <#
    .SYNOPSIS
    Returns a plan object for WinSCP package acquisition behavior.
    #>
    Param(
        [Parameter(Mandatory = $true)]
        [bool]$WinSCPEXEAlreadyExists,
        [Parameter(Mandatory = $true)]
        [bool]$WinSCPDLLAlreadyExists,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$TargetNugetExe,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$WinSCPVersion,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$NugetDownloadFolder
    )

    $commandPreview = "'$TargetNugetExe' Install WinSCP -Version $WinSCPVersion -NonInteractive -OutputDirectory '$NugetDownloadFolder'"
    $shouldDownload = (-not $WinSCPEXEAlreadyExists -or -not $WinSCPDLLAlreadyExists)

    return [pscustomobject]@{
        ShouldDownload = $shouldDownload
        CommandPreview = $commandPreview
    }
}

# Copy and validate WinSCP files in the BizTalk installation folder.
function Invoke-WinSCPTargetInstall {
    <#
    .SYNOPSIS
    Copies WinSCP binaries to BizTalk folder and validates target presence.
    #>
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$SourceExePath,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$SourceDllPath,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$TargetFolder,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$TargetExePath,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$TargetDllPath,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$WinSCPVersion,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$ExeFileName,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$DllFileName
    )

    try {
        Copy-Item -Path $SourceExePath -Destination $TargetFolder -Force -ErrorAction Stop
        Copy-Item -Path $SourceDllPath -Destination $TargetFolder -Force -ErrorAction Stop

        $targetExeExists = Test-Path $TargetExePath
        $targetDllExists = Test-Path $TargetDllPath

        if ($targetExeExists -and $targetDllExists) {
            return [pscustomobject]@{
                InstalledSuccessfully = $true
                TargetExeExists = $targetExeExists
                TargetDllExists = $targetDllExists
                ErrorMessages = @()
            }
        }

        $errorMessages = @()
        if (-not $targetExeExists) {
            $errorMessages += "The $ExeFileName file version $WinSCPVersion"
            $errorMessages += "It was not properly copied to the target folder '$TargetFolder'."
        }
        if (-not $targetDllExists) {
            $errorMessages += "The $DllFileName file version $WinSCPVersion"
            $errorMessages += "Was not properly copied to the target folder '$TargetFolder'."
        }

        return [pscustomobject]@{
            InstalledSuccessfully = $false
            TargetExeExists = $targetExeExists
            TargetDllExists = $targetDllExists
            ErrorMessages = $errorMessages
        }
    }
    catch {
        return [pscustomobject]@{
            InstalledSuccessfully = $false
            TargetExeExists = $false
            TargetDllExists = $false
            ErrorMessages = @(
                'Failed to copy WinSCP files to the BizTalk installation folder.'
                "$($_.Exception.Message)"
            )
        }
    }
}

# Classify package/download readiness from file existence state.
function Get-PackageReadinessState {
    <#
    .SYNOPSIS
    Classifies package acquisition state from known file-existence flags.
    #>
    Param(
        [Parameter(Mandatory = $true)]
        [bool]$NuGetExeExists,
        [Parameter(Mandatory = $true)]
        [bool]$WinSCPExeExists,
        [Parameter(Mandatory = $true)]
        [bool]$WinSCPDllExists
    )

    if (-not $NuGetExeExists) {
        return [pscustomobject]@{
            IsReady = $false
            State = "MissingNuGet"
        }
    }

    if (-not $WinSCPExeExists -and -not $WinSCPDllExists) {
        return [pscustomobject]@{
            IsReady = $false
            State = "MissingWinSCPPackage"
        }
    }

    if (-not $WinSCPExeExists -or -not $WinSCPDllExists) {
        return [pscustomobject]@{
            IsReady = $false
            State = "IncompleteWinSCPPackage"
        }
    }

    return [pscustomobject]@{
        IsReady = $true
        State = "Ready"
    }
}

# Classify final execution outcome from terminal state flags.
function Get-FinalExecutionOutcome {
    <#
    .SYNOPSIS
    Classifies final script outcome using terminal state flags.
    #>
    Param(
        [Parameter(Mandatory = $true)]
        [bool]$InstalledSuccessfully,
        [Parameter(Mandatory = $true)]
        [bool]$WhatIf,
        [Parameter(Mandatory = $true)]
        [bool]$PrerequisiteFailure,
        [Parameter(Mandatory = $true)]
        [bool]$ContinueFlag
    )

    if ($InstalledSuccessfully) {
        return [pscustomobject]@{
            Outcome = "Success"
            IsError = $false
        }
    }

    if ($WhatIf) {
        return [pscustomobject]@{
            Outcome = "DryRun"
            IsError = $false
        }
    }

    if ($PrerequisiteFailure) {
        return [pscustomobject]@{
            Outcome = "PrerequisiteFailure"
            IsError = $true
        }
    }

    if (-not $ContinueFlag) {
        return [pscustomobject]@{
            Outcome = "InstallFailure"
            IsError = $true
        }
    }

    return [pscustomobject]@{
        Outcome = "Unknown"
        IsError = $true
    }
}

# Map a BizTalk product code GUID to a major version label.
function Get-BizTalkVersionFromProductCode {
    <#
    .SYNOPSIS
    Maps a BizTalk Server product code GUID to a version label string.

    .OUTPUTS
    '2016', '2020', or $null for unrecognised product codes.
    #>
    Param(
        [Parameter(Mandatory = $true)]
        [string] $ProductCode
    )

    $BizTalk2016ProductCode = '{B084F3A7-3E8F-4E7B-B673-EED1715D28ED}'
    $BizTalk2020ProductCode = '{205F5836-7512-4A06-9E74-ADC8AFA0EEC5}'

    if ($ProductCode -eq $BizTalk2020ProductCode) { return '2020' }
    if ($ProductCode -eq $BizTalk2016ProductCode) { return '2016' }
    return $null
}

# Select the correct WinSCP version for a BizTalk installation.
function Get-WinSCPVersionForBizTalk {
    <#
    .SYNOPSIS
    Detects the installed BizTalk CU and returns the required WinSCP version.

    .DESCRIPTION
    Combines display-name CU detection and KB-based fallback search to select
    the correct WinSCP version for the given BizTalk major version.
    Returns a structured object, or $null for unsupported BizTalk versions.

    .OUTPUTS
    [pscustomobject] with WinSCPVersion, CULabel, KB, CUFound — or $null.
    #>
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string] $BizTalkVersion
    )

    $winSCPVersion = $null
    $btsKB         = 'none'
    $bizTalkCUVer  = 'no CU'
    $CUFound       = $false

    if ($BizTalkVersion -eq '2020') {
        $winSCPVersion = '5.15.4'
        # CU   Build       KB(s)                 Release Date       WinSCP
        # CU6  3.13.895.0  5043408, 5048971      November 21, 2024  6.3.5
        # CU5  3.13.867.0  5032870               December 3, 2023   6.1.2
        # CU4  3.13.844.0  5009901               August 22, 2022    5.19.2
        # CU3  3.13.812.0  5007969               November 22, 2021  5.19.2
        # CU2  3.13.785.0  5003151               April 19, 2021     5.17.8
        # CU1  3.13.759.0  4538666               July 28, 2020      5.17.6
        # RTM  3.13.717.0  NA                    January 15, 2020   5.15.4
        $bts2020CUMap = @{
            6 = @{ KBs = @('5043408', '5048971'); WinSCP = '6.3.5' }
            5 = @{ KBs = @('5032870');             WinSCP = '6.1.2' }
            4 = @{ KBs = @('5009901');             WinSCP = '5.19.2' }
            3 = @{ KBs = @('5007969');             WinSCP = '5.19.2' }
            2 = @{ KBs = @('5003151');             WinSCP = '5.17.8' }
            1 = @{ KBs = @('4538666');             WinSCP = '5.17.6' }
        }

        $detected = Get-BTSCumulativeUpdateByDisplayName -BizTalkVersion $BizTalkVersion
        if ($detected.Found -and $bts2020CUMap.ContainsKey([int]$detected.CUNumber)) {
            $cuNumber     = [int]$detected.CUNumber
            $cuEntry      = $bts2020CUMap[$cuNumber]
            $bizTalkCUVer = "CU$cuNumber"
            $btsKB        = if ($detected.KB) { $detected.KB } else { $cuEntry.KBs[0] }
            $winSCPVersion = $cuEntry.WinSCP
            $CUFound      = $true
        }

        if (-not $CUFound) {
            foreach ($cuNumber in @(6, 5, 4, 3, 2, 1)) {
                $cuEntry = $bts2020CUMap[$cuNumber]
                foreach ($kb in $cuEntry.KBs) {
                    if (Search-BTSCumulativeUpdate -CumulativeUpdateID $kb -BizTalkVersion $BizTalkVersion) {
                        $bizTalkCUVer  = "CU$cuNumber"
                        $btsKB         = $kb
                        $winSCPVersion = $cuEntry.WinSCP
                        $CUFound       = $true
                        break
                    }
                }
                if ($CUFound) { break }
            }
        }
    }
    elseif ($BizTalkVersion -eq '2016') {
        $winSCPVersion = '5.7.7'
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
        # CU2 or FU1    3.12.807.2  4021095   May 26, 2017       5.7.7
        # CU1           3.12.796.2  3208238   January 26, 2017   5.7.7
        # RTM           3.12.774.0  NA        September 30, 2016 5.7.7
        $bts2016UpdateMap = @(
            @{ Label = 'CU9 and FP3'; KB = '5005480'; WinSCP = '5.19.2' }
            @{ Label = 'CU9';         KB = '5005479'; WinSCP = '5.19.2' }
            @{ Label = 'CU8 and FP3'; KB = '4590075'; WinSCP = '5.15.9' }
            @{ Label = 'CU8';         KB = '4583530'; WinSCP = '5.15.9' }
            @{ Label = 'CU7 and FP3'; KB = '4536185'; WinSCP = '5.15.9' }
            @{ Label = 'CU7';         KB = '4528776'; WinSCP = '5.15.9' }
            @{ Label = 'CU6 and FP3'; KB = '4294900'; WinSCP = '5.13.1' }
            @{ Label = 'CU6';         KB = '4477494'; WinSCP = '5.13.1' }
            @{ Label = 'CU5 and FP3'; KB = '4103503'; WinSCP = '5.13.1' }
            @{ Label = 'CU5 Hotfix';  KB = '4345385'; WinSCP = '5.13.1' }
            @{ Label = 'CU5';         KB = '4132957'; WinSCP = '5.13.1' }
            @{ Label = 'CU4 and FP2'; KB = '4094130'; WinSCP = '5.7.7' }
            @{ Label = 'CU4';         KB = '4051353'; WinSCP = '5.7.7' }
            @{ Label = 'CU3 and FP2'; KB = '4054819'; WinSCP = '5.7.7' }
            @{ Label = 'CU3 and FU1'; KB = '4014788'; WinSCP = '5.7.7' }
            @{ Label = 'CU3';         KB = '4039664'; WinSCP = '5.7.7' }
            @{ Label = 'CU2 or FU1';  KB = '4021095'; WinSCP = '5.7.7' }
            @{ Label = 'CU1';         KB = '3208238'; WinSCP = '5.7.7' }
        )

        foreach ($update in $bts2016UpdateMap) {
            if (Search-BTSCumulativeUpdate -CumulativeUpdateID $update.KB -BizTalkVersion $BizTalkVersion) {
                $winSCPVersion = $update.WinSCP
                $btsKB         = $update.KB
                $bizTalkCUVer  = $update.Label
                $CUFound       = $true
                break
            }
        }
    }
    else {
        return $null
    }

    return [pscustomobject]@{
        WinSCPVersion = $winSCPVersion
        CULabel       = $bizTalkCUVer
        KB            = $btsKB
        CUFound       = $CUFound
    }
}

Export-ModuleMember -Function Resolve-WinSCPPackageLayout, Search-BTSCumulativeUpdate, Get-BTSCumulativeUpdateByDisplayName, Test-IsAdministrator, Get-InstallExecutionPlan, Test-WinSCPVersionString, Resolve-BizTalkInstallFolder, Get-NuGetDownloadPlan, Get-WinSCPPackageDownloadPlan, Invoke-WinSCPTargetInstall, Get-PackageReadinessState, Get-FinalExecutionOutcome, Get-BizTalkVersionFromProductCode, Get-WinSCPVersionForBizTalk
