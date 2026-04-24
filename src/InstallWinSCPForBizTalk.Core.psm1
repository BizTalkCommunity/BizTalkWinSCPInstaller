#####################################################################
# InstallWinSCPForBizTalk.Core.psm1
# 
# Core functions extracted for unit testing
#####################################################################

#####################################################################
# Function to resolve WinSCP package layout from extracted NuGet files
#####################################################################
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

#####################################################################
# Function to search for specific BizTalk cumulative updates
#####################################################################
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

#####################################################################
# Function to detect the highest CU from uninstall DisplayName entries
#####################################################################
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

#####################################################################
# Function to test if current session runs elevated
#####################################################################
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

#####################################################################
# Function to evaluate install-gating behavior
#####################################################################
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

Export-ModuleMember -Function Resolve-WinSCPPackageLayout, Search-BTSCumulativeUpdate, Get-BTSCumulativeUpdateByDisplayName, Test-IsAdministrator, Get-InstallExecutionPlan
