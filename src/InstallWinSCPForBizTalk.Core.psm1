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

Export-ModuleMember -Function Resolve-WinSCPPackageLayout
