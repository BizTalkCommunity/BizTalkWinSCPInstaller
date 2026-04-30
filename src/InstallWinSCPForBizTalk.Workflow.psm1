# InstallWinSCPForBizTalk.Workflow.psm1
# Workflow/orchestration phase helpers for the installer script.

# Detect BizTalk installation metadata and normalize install target details.
function Invoke-BizTalkDetectionPhase {
    <#
    .SYNOPSIS
    Detects BizTalk installation metadata and resolves install location details.

    .DESCRIPTION
    Uses environment/registry inputs to resolve the BizTalk installation folder,
    validates required product metadata, and updates the shared workflow context.

    .PARAMETER Context
    Mutable hashtable used to share installer state across workflow phases.

    .PARAMETER EnvironmentInstallPath
    BTSINSTALLPATH value from environment lookup.

    .PARAMETER RegistryInstallPath
    InstallPath value from BizTalk registry lookup.

    .PARAMETER ProductCodeCurrent
    BizTalk ProductCodeCurrent registry value.

    .PARAMETER ProductName
    BizTalk ProductName registry value.

    .PARAMETER ProductVersion
    BizTalk ProductVersion registry value.
    #>
    Param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Context,

        [string]$EnvironmentInstallPath,
        [string]$RegistryInstallPath,
        [string]$ProductCodeCurrent,
        [string]$ProductName,
        [string]$ProductVersion
    )

    $Context.bizTalkInstallFolderFromEnv = $EnvironmentInstallPath
    $Context.bizTalkInstallFolderFromRegistry = $RegistryInstallPath
    $Context.bizTalkInstallFolderResolution = Resolve-BizTalkInstallFolder -EnvironmentInstallPath $Context.bizTalkInstallFolderFromEnv -RegistryInstallPath $Context.bizTalkInstallFolderFromRegistry
    $Context.bizTalkInstallFolder = $Context.bizTalkInstallFolderResolution.InstallPath
    $Context.bizTalkInstallFolderExists = $Context.bizTalkInstallFolderResolution.Exists
    $Context.bizTalkProductCodeCurrent = $ProductCodeCurrent
    $Context.bizTalkProductName = $ProductName
    $Context.bizTalkProductVersion = $ProductVersion

    if ($Context.Continue) {
        Write-InstallerSectionHeader -Title 'Checking if Microsoft BizTalk Server is installed.'
        if (-not $Context.bizTalkInstallFolderResolution.IsFound) {
            Write-BizTalkRegistryFallbackNotice
            $Context.Continue = $false
            Write-BizTalkNotLocatedError
        }
        elseif ($Context.bizTalkInstallFolderResolution.Source -eq 'Registry') {
            Write-BizTalkRegistryFallbackNotice
        }

        if ($Context.Continue -and -not $Context.bizTalkInstallFolderExists) {
            $Context.Continue = $false
            Write-InstallerBangError -MessageLines @(
                'Microsoft BizTalk Server installation was found in the registry.'
                "Regardless, the $($Context.bizTalkInstallFolder) folder does not exist."
                'Although it was found in the registry and/or the BTSINSTALLPATH environment setting'
                'Exiting...'
            )
        }
        else {
            Write-BizTalkLocatedSuccess -ProductName $Context.bizTalkProductName -ProductVersion $Context.bizTalkProductVersion -InstallFolder $Context.bizTalkInstallFolder
            $Context.BizTalkVersion = Get-BizTalkVersionFromProductCode -ProductCode $Context.bizTalkProductCodeCurrent
            if ($null -eq $Context.BizTalkVersion) {
                $Context.Continue = $false
                Write-InstallerBangError -LeadingNewLine -MessageLines @(
                    'Neither Microsoft BizTalk Server 2016 nor Microsoft BizTalk Server 2020 were found.'
                    'Exiting...'
                )
            }
        }
    }

    Write-BizTalkSearchSnapshot -InstallFolder $Context.bizTalkInstallFolder -InstallFolderExists $Context.bizTalkInstallFolderExists -ProductCodeCurrent $Context.bizTalkProductCodeCurrent -ProductName $Context.bizTalkProductName -ProductVersion $Context.bizTalkProductVersion
}

# Enforce ForceInstall elevation rules before any mutating action is attempted.
function Invoke-ForceInstallPrerequisitePhase {
    <#
    .SYNOPSIS
    Applies elevation prerequisites for ForceInstall scenarios.

    .DESCRIPTION
    Prevents real install execution for non-admin ForceInstall runs unless WhatIf
    is enabled, and records prerequisite failure state in workflow context.

    .PARAMETER Context
    Mutable hashtable used to share installer state across workflow phases.

    .PARAMETER WhatIf
    Indicates whether script execution is in WhatIf (dry-run) mode.
    #>
    Param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Context,

        [Parameter(Mandatory = $true)]
        [bool]$WhatIf
    )

    if ($Context.Continue -and -not $Context.isAdministrator -and $Context.ForceInstall) {
        $nonAdminForceInstallMessages = @(
            'This PowerShell session is not running as Administrator.'
            "The BizTalk installation folder requires elevation for writes: $($Context.bizTalkInstallFolder)"
        )

        if (-not $WhatIf) {
            Write-InstallerBangError -LeadingNewLine -MessageLines ($nonAdminForceInstallMessages + @(
                'ForceInstall requires write access and will fail without elevation.'
                'Please rerun from an elevated PowerShell session.'
            ))
            $Context.PrerequisiteFailure = $true
            $Context.Continue = $false
        }
        else {
            Write-InstallerBangError -LeadingNewLine -MessageLines ($nonAdminForceInstallMessages + @(
                'Continuing in read/dry-run mode. If a real install is required, rerun elevated.'
            ))
        }
    }
}

# Resolve required WinSCP version from detected BizTalk version/CU mapping.
function Invoke-CuDetectionPhase {
    <#
    .SYNOPSIS
    Resolves the required WinSCP version for the detected BizTalk installation.

    .DESCRIPTION
    Calls CU mapping logic, sets WinSCP/CU fields in workflow context, and emits
    standardized output for detected CU or RTM baseline.

    .PARAMETER Context
    Mutable hashtable used to share installer state across workflow phases.
    #>
    Param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Context
    )

    $Context.winSCPVersion = $null
    $Context.btsKB = 'none'
    $Context.bizTalkCUVer = 'no CU'
    $Context.CUFound = $false

    if ($Context.Continue) {
        Write-InstallerSectionHeader -Title 'Determining what version of WinSCP to download' -LeadingNewLine
        Write-BizTalkCuDetectionStart -BizTalkVersion $Context.BizTalkVersion
        $cuResult = Get-WinSCPVersionForBizTalk -BizTalkVersion $Context.BizTalkVersion
        if ($null -eq $cuResult) {
            $Context.Continue = $false
            Write-InstallerBangError -LeadingNewLine -MessageLines @(
                "Could not determine the required WinSCP version for BizTalk Server $($Context.BizTalkVersion)."
            )
        }
        else {
            $Context.winSCPVersion = $cuResult.WinSCPVersion
            $Context.btsKB         = $cuResult.KB
            $Context.bizTalkCUVer  = $cuResult.CULabel
            $Context.CUFound       = $cuResult.CUFound
        }

        if ($Context.CUFound) {
            Write-InstallerSuccess "Detected Microsoft BizTalk Server $($Context.BizTalkVersion) $($Context.bizTalkCUVer) KB$($Context.btsKB)"
        }
        else {
            Write-InstallerSuccess "Detected Microsoft BizTalk Server $($Context.BizTalkVersion) with no cumulative updates."
        }

        if ($Context.ForceInstall) {
            Write-InstallerSuccess "ForceInstall was specified; this script will download/reuse and install WinSCP $($Context.winSCPVersion)."
        }
        else {
            Write-InstallerSuccess "If necessary, this script will download WinSCP $($Context.winSCPVersion)"
        }
    }

    Write-BizTalkCuSearchSnapshot -WinSCPVersion $Context.winSCPVersion -KB $Context.btsKB -CULabel $Context.bizTalkCUVer -CUFound $Context.CUFound
}

# Check for an already-correct WinSCP install and decide skip/reinstall behavior.
function Invoke-ExistingWinSCPCheckPhase {
    <#
    .SYNOPSIS
    Checks existing WinSCP files and computes install/reinstall execution plan.

    .DESCRIPTION
    Reads existing BizTalk target file versions, compares with required version,
    and sets execution gating flags in workflow context.

    .PARAMETER Context
    Mutable hashtable used to share installer state across workflow phases.
    #>
    Param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Context
    )

    $Context.btsWinSCPEXEProductVersionInstalled = 'None'
    $Context.btsWinSCPDLLProductVersionInstalled = 'None'
    $Context.btsWinSCPProductInstalledAndCorrect = $false
    $Context.installExecutionPlan = $null
    $Context.winSCPProductVersionRequired = $Context.winSCPVersion
    $Context.btsTargetWinSCPExe = $Context.bizTalkInstallFolder + $Context.winSCPexeFile
    $Context.btsTargetWinSCPDll = $Context.bizTalkInstallFolder + $Context.winSCPdllFile

    if ($Context.Continue) {
        Write-InstallerDelimitedMessage -MessageLines @(
            "Checking to see if WinSCP $($Context.winSCPVersion) is already"
            'installed in the Microsoft BizTalk Server folder.'
        ) -Delimiter 'Hash' -Level 'Success' -LeadingNewLine

        if ((Test-Path $Context.btsTargetWinSCPExe) -and (Test-Path $Context.btsTargetWinSCPDll)) {
            $Context.btsWinSCPEXEProductVersionInstalled = (get-item $Context.btsTargetWinSCPExe).VersionInfo.ProductVersion
            $Context.btsWinSCPDLLProductVersionInstalled = (get-item $Context.btsTargetWinSCPDll).VersionInfo.ProductVersion
            $Context.winSCPProductVersionRequired = $Context.winSCPVersion

            if ($Context.winSCPVersion.length -gt 2 -and $Context.btsWinSCPEXEProductVersionInstalled.length -gt 2 -and $Context.winSCPVersion.SubString($Context.winSCPVersion.length - 2, 2) -ne '.0' -and $Context.btsWinSCPEXEProductVersionInstalled.SubString($Context.btsWinSCPEXEProductVersionInstalled.length - 2, 2) -eq '.0') {
                $Context.winSCPProductVersionRequired = $Context.winSCPVersion + '.0'
            }

            if ($Context.winSCPProductVersionRequired -eq $Context.btsWinSCPEXEProductVersionInstalled -and $Context.winSCPProductVersionRequired -eq $Context.btsWinSCPDLLProductVersionInstalled) {
                $Context.btsWinSCPProductInstalledAndCorrect = $true
            }
        }

        $Context.installExecutionPlan = Get-InstallExecutionPlan -IsAdministrator $Context.isAdministrator -ForceInstall ([bool]$Context.ForceInstall) -WhatIf ([bool]$Context.WhatIf) -AlreadyInstalledCorrect $Context.btsWinSCPProductInstalledAndCorrect
        if ($Context.btsWinSCPProductInstalledAndCorrect) {
            Write-InstallerSuccess "Detected WinSCP $($Context.winSCPVersion) is already installed in Microsoft BizTalk Server."
            if ($Context.installExecutionPlan.ShouldReinstall) {
                Write-InstallerSuccess 'Reinstalling because ForceInstall was specified.'
                $Context.btsWinSCPProductInstalledAndCorrect = $false
            }
            elseif ($Context.installExecutionPlan.ShouldSkipBecauseInstalled) {
                Write-InstallerSuccess 'Skipping installing the already installed version.'
                $Context.Continue = $false
            }
        }
        else {
            Write-WinSCPNotInstalledNotice -WinSCPVersion $Context.winSCPVersion
        }
    }

    Write-ExistingWinSCPSnapshot -ExeProductVersionInstalled $Context.btsWinSCPEXEProductVersionInstalled -DllProductVersionInstalled $Context.btsWinSCPDLLProductVersionInstalled -ProductVersionRequired $Context.winSCPProductVersionRequired -InstalledAndCorrect $Context.btsWinSCPProductInstalledAndCorrect -TargetExePath $Context.btsTargetWinSCPExe -TargetDllPath $Context.btsTargetWinSCPDll
}

# Emit non-admin warning when continuing in non-elevated non-force scenarios.
function Invoke-NonAdminWarningPhase {
    <#
    .SYNOPSIS
    Emits non-admin warning messaging for read/dry-run continuation paths.

    .DESCRIPTION
    Uses install execution plan fields in workflow context to display guidance
    when elevation is recommended but execution can continue safely.

    .PARAMETER Context
    Mutable hashtable used to share installer state across workflow phases.
    #>
    Param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Context
    )

    if ($Context.Continue -and $Context.installExecutionPlan -and $Context.installExecutionPlan.RequiresElevationWarning -and -not $Context.ForceInstall) {
        Write-InstallerBangError -LeadingNewLine -MessageLines @(
            'This PowerShell session is not running as Administrator.'
            "The BizTalk installation folder requires elevation for writes: $($Context.bizTalkInstallFolder)"
            'Continuing in read/dry-run mode. If a real install is required, rerun elevated.'
        )
    }
}

# Validate resolved WinSCP version format before downloads/install proceed.
function Invoke-VersionValidationPhase {
    <#
    .SYNOPSIS
    Validates resolved WinSCP version string format.

    .DESCRIPTION
    Uses core version parsing checks and stops workflow progression when version
    mapping yields missing or invalid values.

    .PARAMETER Context
    Mutable hashtable used to share installer state across workflow phases.
    #>
    Param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Context
    )

    if ($Context.Continue) {
        $versionValidation = Test-WinSCPVersionString -WinSCPVersion $Context.winSCPVersion
        if (-not $versionValidation.IsValid) {
            $Context.Continue = $false
            if ($versionValidation.ErrorCode -eq 'MissingVersion') {
                Write-InstallerError 'The WinSCP version was not set - the CU detection logic did not assign a value to `$winSCPVersion.'
                Write-InstallerError 'This is likely a script bug. Review the BizTalk CU detection output above and confirm a CU or RTM baseline was matched.'
            }
            else {
                Write-InstallerError "The WinSCP version '$($Context.winSCPVersion)' is not a valid dotted version number (e.g. 5.7.7 or 6.3.5)."
                Write-InstallerError "This value came from the CU map table in this script. Check the WinSCP version string for BizTalk $($Context.BizTalkVersion) $($Context.bizTalkCUVer) in the table and correct it."
            }
            Write-InstallerError $Context.bangString
        }
    }
}

# Ensure download folder state is ready and capture initial package-layout snapshot.
function Invoke-DownloadFolderPreparationPhase {
    <#
    .SYNOPSIS
    Prepares download folder and resolves pre-existing package layout state.

    .DESCRIPTION
    Creates the NuGet download folder when needed, resolves initial package file
    locations, and stores folder/package readiness details in workflow context.

    .PARAMETER Context
    Mutable hashtable used to share installer state across workflow phases.
    #>
    Param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Context
    )

    if ($Context.Continue) {
        $winSCPPackageRoot = "$($Context.nugetDownloadFolder)\WinSCP.$($Context.winSCPVersion)"
        $winSCPLayout = Resolve-WinSCPPackageLayout -PackageRoot $winSCPPackageRoot -ExeFileName $Context.winSCPexeFile -DllFileName $Context.winSCPdllFile

        $Context.WinSCPEXEDownload = $winSCPLayout.ExePath
        $Context.WinSCPDllDownload = $winSCPLayout.DllPath
        $Context.WinSCPEXEDownloadAlreadyExists = $false
        $Context.WinSCPDllDownloadAlreadyExists = $false
        if ($Context.WinSCPEXEDownload) {
            $Context.WinSCPEXEDownloadAlreadyExists = Test-Path $Context.WinSCPEXEDownload
        }
        if ($Context.WinSCPDllDownload) {
            $Context.WinSCPDllDownloadAlreadyExists = Test-Path $Context.WinSCPDllDownload
        }
        $Context.nugetDownloadFolderAlreadyExists = Test-Path $Context.nugetDownloadFolder
        $Context.nugetDownloadFolderExists = $Context.nugetDownloadFolderAlreadyExists

        Write-InstallerSectionHeader -Title 'Preparing the output folder to store the NuGet and WinSCP downloads' -LeadingNewLine
        if ($Context.Continue -and -not $Context.nugetDownloadFolderAlreadyExists) {
            if ($Context.psCmdlet.ShouldProcess($Context.nugetDownloadFolder, 'Create Folder')) {
                Write-InstallerSuccess "The target folder '$($Context.nugetDownloadFolder)' doesn't exist, creating the folder."
                New-Item -Path $Context.nugetDownloadFolder -ItemType 'Directory' -Force > $null
                $Context.nugetDownloadFolderExists = Test-Path $Context.nugetDownloadFolder
                if (-not $Context.nugetDownloadFolderExists) {
                    $Context.Continue = $false
                    Write-InstallerError "An attempt to use the '$($Context.nugetDownloadFolder)' directory for a download target failed."
                }
            }
        }

        if ($Context.Continue -and $Context.nugetDownloadFolderExists) {
            Write-InstallerSuccess "The target folder '$($Context.nugetDownloadFolder)' is ready for use."
        }

        Write-TargetFolderPreparationSnapshot -NugetDownloadFolderAlreadyExists $Context.nugetDownloadFolderAlreadyExists -NugetDownloadFolderExists $Context.nugetDownloadFolderExists -WinSCPEXEDownload $Context.WinSCPEXEDownload -WinSCPDllDownload $Context.WinSCPDllDownload -WinSCPEXEDownloadAlreadyExists $Context.WinSCPEXEDownloadAlreadyExists -WinSCPDllDownloadAlreadyExists $Context.WinSCPDllDownloadAlreadyExists
    }
}

# Download (or reuse) nuget.exe and snapshot resulting readiness state.
function Invoke-NuGetDownloadPhase {
    <#
    .SYNOPSIS
    Ensures nuget.exe is available for package acquisition.

    .DESCRIPTION
    Reuses or downloads nuget.exe based on current state and ForceInstall flags,
    then records resulting readiness state in workflow context.

    .PARAMETER Context
    Mutable hashtable used to share installer state across workflow phases.
    #>
    Param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Context
    )

    $Context.targetNugetExe = "$($Context.nugetDownloadFolder)\nuget.exe"
    $Context.targetNugetExeAlreadyExists = Test-Path $Context.targetNugetExe
    $Context.targetNugetExeExists = $Context.targetNugetExeAlreadyExists
    $Context.sourceNugetExe = 'https://dist.nuget.org/win-x86-commandline/latest/nuget.exe'

    if ($Context.Continue) {
        $nugetDownloadPlan = Get-NuGetDownloadPlan -TargetNugetExeAlreadyExists ([bool]$Context.targetNugetExeAlreadyExists) -ForceInstall ([bool]$Context.ForceInstall) -SourceNugetExe $Context.sourceNugetExe -TargetNugetExe $Context.targetNugetExe
        if ($nugetDownloadPlan.ShouldDownload) {
            Write-InstallerSuccess "`n$($Context.hashString)"
            Write-InstallerSuccess 'Downloading NuGet from:'
            Write-InstallerSuccess "`t'$($Context.sourceNugetExe)'"
            Write-InstallerSuccess 'Storing it in the folder'
            Write-InstallerSuccess "`t'$($Context.nugetDownloadFolder)'"
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

            if ($Context.psCmdlet.ShouldProcess($nugetDownloadPlan.ShouldProcessTarget, 'Run Invoke-WebRequest ')) {
                & $Context.InvokeWebRequest $Context.sourceNugetExe $Context.targetNugetExe
                $Context.targetNugetExeExists = Test-Path $Context.targetNugetExe
                if (-not $Context.targetNugetExeExists) {
                    $Context.Continue = $false
                    Write-InstallerBangError -LeadingNewLine -MessageLines @(
                        'The download of the NuGet EXE from'
                        $Context.sourceNugetExe
                        'did not succeed'
                    )
                }
            }
        }
        else {
            Write-InstallerDelimitedMessage -MessageLines @(
                'NuGet already exists in the folder'
                $Context.nugetDownloadFolder
                "and doesn't need to be downloaded"
            ) -Delimiter 'Hash' -Level 'Success' -LeadingNewLine
        }
    }

    Write-NuGetDownloadSnapshot -TargetNugetExeAlreadyExists $Context.targetNugetExeAlreadyExists -TargetNugetExeExists $Context.targetNugetExeExists -SourceNugetExe $Context.sourceNugetExe -TargetNugetExe $Context.targetNugetExe
}

# Download (or reuse) WinSCP package files and verify required artifacts exist.
function Invoke-WinSCPPackageDownloadPhase {
    <#
    .SYNOPSIS
    Ensures required WinSCP package artifacts exist in the download folder.

    .DESCRIPTION
    Reuses or downloads WinSCP package files, resolves extracted EXE/DLL paths,
    and records package completeness state in workflow context.

    .PARAMETER Context
    Mutable hashtable used to share installer state across workflow phases.
    #>
    Param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Context
    )

    $Context.getWinSCP = $null
    $Context.WinSCPEXEAlreadyExists = $false
    $Context.WinSCPDLLAlreadyExists = $false
    $Context.WinSCPEXEExists = $false
    $Context.WinSCPDLLExists = $false

    if ($Context.Continue) {
        $winSCPPackageRoot = "$($Context.nugetDownloadFolder)\WinSCP.$($Context.winSCPVersion)"
        $winSCPLayout = Resolve-WinSCPPackageLayout -PackageRoot $winSCPPackageRoot -ExeFileName $Context.winSCPexeFile -DllFileName $Context.winSCPdllFile
        $Context.WinSCPEXEDownload = $winSCPLayout.ExePath
        $Context.WinSCPDllDownload = $winSCPLayout.DllPath
        $Context.WinSCPEXEAlreadyExists = [bool]$Context.WinSCPEXEDownload -and (Test-Path $Context.WinSCPEXEDownload)
        $Context.WinSCPDLLAlreadyExists = [bool]$Context.WinSCPDllDownload -and (Test-Path $Context.WinSCPDllDownload)
        $Context.WinSCPEXEExists = $Context.WinSCPEXEAlreadyExists
        $Context.WinSCPDLLExists = $Context.WinSCPDLLAlreadyExists

        $winSCPDownloadPlan = Get-WinSCPPackageDownloadPlan -WinSCPEXEAlreadyExists ([bool]$Context.WinSCPEXEAlreadyExists) -WinSCPDLLAlreadyExists ([bool]$Context.WinSCPDLLAlreadyExists) -TargetNugetExe $Context.targetNugetExe -WinSCPVersion $Context.winSCPVersion -NugetDownloadFolder $Context.nugetDownloadFolder
        $Context.getWinSCP = $winSCPDownloadPlan.CommandPreview

        Write-InstallerDelimitedMessage -MessageLines @(
            "Downloading WinSCP version $($Context.winSCPVersion) from NuGet"
            "`t$($Context.getWinSCP)"
            'Storing it in the folder:'
            "`t'$($Context.nugetDownloadFolder)'"
        ) -Delimiter 'Hash' -Level 'Success' -LeadingNewLine

        if ($winSCPDownloadPlan.ShouldDownload) {
            if ($Context.psCmdlet.ShouldProcess($Context.getWinSCP, 'Run Command')) {
                & $Context.targetNugetExe Install WinSCP -Version $Context.winSCPVersion -NonInteractive -OutputDirectory $Context.nugetDownloadFolder
                $winSCPLayout = Resolve-WinSCPPackageLayout -PackageRoot $winSCPPackageRoot -ExeFileName $Context.winSCPexeFile -DllFileName $Context.winSCPdllFile
                $Context.WinSCPEXEDownload = $winSCPLayout.ExePath
                $Context.WinSCPDllDownload = $winSCPLayout.DllPath
                $Context.WinSCPEXEExists = [bool]$Context.WinSCPEXEDownload -and (Test-Path $Context.WinSCPEXEDownload)
                $Context.WinSCPDLLExists = [bool]$Context.WinSCPDllDownload -and (Test-Path $Context.WinSCPDllDownload)
                if (-not $Context.WinSCPEXEExists -or -not $Context.WinSCPDLLExists) {
                    $Context.Continue = $false
                    Write-InstallerBangError -LeadingNewLine -MessageLines @(
                        "WinSCP $($Context.winSCPVersion) was not properly downloaded."
                        'Check the folder and error messages above:'
                        $Context.nugetDownloadFolder
                        'And determine what files did download or did not download.'
                    )
                }
            }
        }
    }

    Write-WinSCPDownloadSnapshot -GetWinSCP $Context.getWinSCP -WinSCPEXEAlreadyExists $Context.WinSCPEXEAlreadyExists -WinSCPDLLAlreadyExists $Context.WinSCPDLLAlreadyExists -WinSCPDllDownload $Context.WinSCPDllDownload -WinSCPEXEExists $Context.WinSCPEXEExists -WinSCPDLLExists $Context.WinSCPDLLExists
}

# Copy resolved WinSCP artifacts into BizTalk folder and record install outcome.
function Invoke-WinSCPCopyPhase {
    <#
    .SYNOPSIS
    Copies WinSCP binaries into BizTalk installation folder.

    .DESCRIPTION
    Executes final copy operation with ShouldProcess support, validates target
    artifact presence, and records install outcome in workflow context.

    .PARAMETER Context
    Mutable hashtable used to share installer state across workflow phases.
    #>
    Param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Context
    )

    $Context.WinSCPTargetEXEExists = Test-Path $Context.btsTargetWinSCPExe
    $Context.WinSCPDLLTargetExists = Test-Path $Context.btsTargetWinSCPDll

    if ($Context.Continue -and -not $Context.btsWinSCPProductInstalledAndCorrect) {
        Write-InstallerSectionHeader -Title 'Installing WinSCP' -LeadingNewLine
        Write-InstallerSuccess "Copying WinSCP version $($Context.winSCPVersion) to Microsoft BizTalk Server Folder:"
        Write-InstallerSuccess "`t'$($Context.bizTalkInstallFolder)'."

        $copySourceSummary = if ($Context.WinSCPEXEDownload -and $Context.WinSCPDllDownload) {
            "$($Context.WinSCPEXEDownload) and $($Context.WinSCPDllDownload)"
        }
        else {
            "WinSCP package files for version $($Context.winSCPVersion)"
        }

        if ($Context.psCmdlet.ShouldProcess("$copySourceSummary to '$($Context.bizTalkInstallFolder)'", 'Copy Files')) {
            $installResult = Invoke-WinSCPTargetInstall -SourceExePath $Context.WinSCPEXEDownload -SourceDllPath $Context.WinSCPDllDownload -TargetFolder $Context.bizTalkInstallFolder -TargetExePath $Context.btsTargetWinSCPExe -TargetDllPath $Context.btsTargetWinSCPDll -WinSCPVersion $Context.winSCPVersion -ExeFileName $Context.winSCPexeFile -DllFileName $Context.winSCPdllFile
            $Context.WinSCPTargetEXEExists = $installResult.TargetExeExists
            $Context.WinSCPDLLTargetExists = $installResult.TargetDllExists
            if ($installResult.InstalledSuccessfully) {
                $Context.btsWinSCPProductInstalledAndCorrect = $true
            }
            else {
                $Context.Continue = $false
                foreach ($message in $installResult.ErrorMessages) {
                    Write-InstallerError $message
                }
            }
        }
    }

    Write-WinSCPCopySnapshot -BizTalkInstallFolder $Context.bizTalkInstallFolder -WinSCPEXEDownload $Context.WinSCPEXEDownload -WinSCPDllDownload $Context.WinSCPDllDownload -WinSCPTargetEXEExists $Context.WinSCPTargetEXEExists -WinSCPDLLTargetExists $Context.WinSCPDLLTargetExists -InstalledAndCorrect $Context.btsWinSCPProductInstalledAndCorrect
}

# Verify installed WinSCP file versions and optionally check copy integrity.
function Invoke-WinSCPVerificationPhase {
    <#
    .SYNOPSIS
    Verifies installed WinSCP file versions and optionally checks copy integrity.

    .DESCRIPTION
    Reads ProductVersion from installed WinSCP files and compares against the
    expected version. When CheckHash is set in the workflow context, also verifies
    the SHA256 of each installed file against its source in the download folder.
    Records verification results in workflow context. Runs whenever
    btsWinSCPProductInstalledAndCorrect is true — covering both the
    just-installed and already-correct-skip paths.

    .PARAMETER Context
    Mutable hashtable used to share installer state across workflow phases.
    #>
    Param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Context
    )

    $Context.verificationResult = $null

    if ($Context.btsWinSCPProductInstalledAndCorrect) {
        Write-InstallerSectionHeader -Title 'Verifying installed WinSCP files' -LeadingNewLine

        $verificationArgs = @{
            TargetExePath   = $Context.btsTargetWinSCPExe
            TargetDllPath   = $Context.btsTargetWinSCPDll
            ExpectedVersion = $Context.winSCPVersion
            CheckHash       = [bool]$Context.CheckHash
        }
        if ($Context.CheckHash -and $Context.WinSCPEXEDownload) {
            $verificationArgs['SourceExePath'] = $Context.WinSCPEXEDownload
        }
        if ($Context.CheckHash -and $Context.WinSCPDllDownload) {
            $verificationArgs['SourceDllPath'] = $Context.WinSCPDllDownload
        }

        $verification = Get-WinSCPInstallVerification @verificationArgs
        $Context.verificationResult = $verification

        Write-WinSCPVersionCheckResults -Verification $verification -ExpectedVersion $Context.winSCPVersion
        Write-WinSCPHashCheckResults -Verification $verification

        if ($verification.VerifiedSuccessfully) {
            Write-InstallerSuccess 'Verification passed.'
        }
        else {
            Write-InstallerBangError -LeadingNewLine -MessageLines @(
                'Post-install verification failed. One or more installed WinSCP files'
                'do not match the expected version or failed the integrity check.'
                'Review the errors above. You may need to rerun with -ForceInstall.'
            )
        }

        Write-WinSCPVerificationSnapshot -Verification $verification -ExpectedVersion $Context.winSCPVersion -CheckHash ([bool]$Context.CheckHash)
    }
}

Export-ModuleMember -Function Invoke-BizTalkDetectionPhase, Invoke-ForceInstallPrerequisitePhase, Invoke-CuDetectionPhase, Invoke-ExistingWinSCPCheckPhase, Invoke-NonAdminWarningPhase, Invoke-VersionValidationPhase, Invoke-DownloadFolderPreparationPhase, Invoke-NuGetDownloadPhase, Invoke-WinSCPPackageDownloadPhase, Invoke-WinSCPCopyPhase, Invoke-WinSCPVerificationPhase
