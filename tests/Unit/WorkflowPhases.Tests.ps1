<#
.SYNOPSIS
Unit tests for workflow phase functions covering uncovered error and edge-case branches.

.DESCRIPTION
Tests each Invoke-*Phase function in isolation using a mutable context hashtable and
module-scoped mocks. Targets branches that are unreachable through the integration-level
InstallerScript.Execution tests: BizTalk-not-found, folder-missing, ForceInstall
prerequisite failure, null CU result, already-installed detection, version validation
failures, folder creation failure, and download failure paths.
#>

if (-not (Get-Module -ListAvailable Pester | Where-Object { $_.Version.Major -ge 5 })) {
    throw "Pester 5+ is required to run this test file. Install with: Install-Module Pester -Scope CurrentUser -RequiredVersion 5.0 -Force"
}

Describe "Workflow phase functions" {
    BeforeAll {
        $script:repoRoot = Resolve-Path (Join-Path (Split-Path -Parent $PSCommandPath) "../..")
        Import-Module (Join-Path $script:repoRoot "src/InstallWinSCPForBizTalk.Core.psm1") -Force
        Import-Module (Join-Path $script:repoRoot "src/InstallWinSCPForBizTalk.Utils.psm1") -Force
        Import-Module (Join-Path $script:repoRoot "src/InstallWinSCPForBizTalk.Workflow.psm1") -Force

        # Create a temporary no-op batch file used as a fake nuget.exe in download tests.
        $script:fakenugetBat = [System.IO.Path]::ChangeExtension([System.IO.Path]::GetTempFileName(), 'bat')
        '@echo off' | Set-Content $script:fakenugetBat -Encoding ASCII

        function script:New-TestContext {
            $mockCmdlet = [PSCustomObject]@{ ShouldProcessResult = $true }
            $mockCmdlet | Add-Member -MemberType ScriptMethod -Name ShouldProcess -Value {
                return $this.ShouldProcessResult
            }

            @{
                Continue                            = $true
                isAdministrator                     = $false
                ForceInstall                        = $false
                WhatIf                              = $false
                bizTalkInstallFolder                = 'C:\BizTalk\'
                bizTalkInstallFolderExists          = $true
                winSCPVersion                       = '5.15.4'
                winSCPexeFile                       = 'WinSCP.exe'
                winSCPdllFile                       = 'WinSCPnet.dll'
                btsTargetWinSCPExe                  = 'C:\BizTalk\WinSCP.exe'
                btsTargetWinSCPDll                  = 'C:\BizTalk\WinSCPnet.dll'
                hashString                          = '##############################################################################'
                bangString                          = '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!'
                BizTalkVersion                      = '2020'
                bizTalkCUVer                        = 'CU6'
                btsKB                               = '5048971'
                CUFound                             = $true
                nugetDownloadFolder                 = 'C:\fake\nuget'
                targetNugetExe                      = $script:fakenugetBat
                sourceNugetExe                      = 'https://dist.nuget.org/win-x86-commandline/latest/nuget.exe'
                WinSCPEXEDownload                   = $null
                WinSCPDllDownload                   = $null
                installExecutionPlan                = $null
                btsWinSCPProductInstalledAndCorrect = $false
                psCmdlet                            = $mockCmdlet
                InvokeWebRequest                    = { param($src, $dst) }
            }
        }
    }

    AfterAll {
        Remove-Module InstallWinSCPForBizTalk.Workflow -ErrorAction SilentlyContinue
        Remove-Module InstallWinSCPForBizTalk.Utils -ErrorAction SilentlyContinue
        Remove-Module InstallWinSCPForBizTalk.Core -ErrorAction SilentlyContinue
        if (Test-Path $script:fakenugetBat) {
            Remove-Item $script:fakenugetBat -Force -ErrorAction SilentlyContinue
        }
    }

    # -------------------------------------------------------------------------
    Context "Invoke-BizTalkDetectionPhase" {
        BeforeEach {
            Mock -ModuleName InstallWinSCPForBizTalk.Workflow Write-InstallerSectionHeader {}
            Mock -ModuleName InstallWinSCPForBizTalk.Workflow Write-BizTalkRegistryFallbackNotice {}
            Mock -ModuleName InstallWinSCPForBizTalk.Workflow Write-BizTalkNotLocatedError {}
            Mock -ModuleName InstallWinSCPForBizTalk.Workflow Write-InstallerBangError {}
            Mock -ModuleName InstallWinSCPForBizTalk.Workflow Write-BizTalkLocatedSuccess {}
            Mock -ModuleName InstallWinSCPForBizTalk.Workflow Write-BizTalkSearchSnapshot {}
            Mock -ModuleName InstallWinSCPForBizTalk.Workflow Get-BizTalkVersionFromProductCode { return '2020' }
        }

        It "stops Continue and calls registry-fallback and not-located helpers when BizTalk is not found" {
            Mock -ModuleName InstallWinSCPForBizTalk.Workflow Resolve-BizTalkInstallFolder {
                [PSCustomObject]@{ IsFound = $false; Source = 'None'; InstallPath = 'C:\BizTalk\'; Exists = $false }
            }

            $ctx = script:New-TestContext
            Invoke-BizTalkDetectionPhase -Context $ctx `
                -ProductCodeCurrent '{205F5836-7512-4A06-9E74-ADC8AFA0EEC5}' `
                -ProductName 'Microsoft BizTalk Server 2020' -ProductVersion '3.13.717.0'

            $ctx.Continue | Should -BeFalse
            Should -Invoke Write-BizTalkRegistryFallbackNotice -ModuleName InstallWinSCPForBizTalk.Workflow -Times 1
            Should -Invoke Write-BizTalkNotLocatedError -ModuleName InstallWinSCPForBizTalk.Workflow -Times 1
        }

        It "stops Continue and emits bang error when BizTalk folder is missing despite registry entry" {
            Mock -ModuleName InstallWinSCPForBizTalk.Workflow Resolve-BizTalkInstallFolder {
                [PSCustomObject]@{ IsFound = $true; Source = 'Registry'; InstallPath = 'C:\BizTalk\'; Exists = $false }
            }

            $ctx = script:New-TestContext
            Invoke-BizTalkDetectionPhase -Context $ctx `
                -EnvironmentInstallPath '' -RegistryInstallPath 'C:\BizTalk\' `
                -ProductCodeCurrent '{205F5836-7512-4A06-9E74-ADC8AFA0EEC5}' `
                -ProductName 'Microsoft BizTalk Server 2020' -ProductVersion '3.13.717.0'

            $ctx.Continue | Should -BeFalse
            Should -Invoke Write-InstallerBangError -ModuleName InstallWinSCPForBizTalk.Workflow -Times 1
        }
    }

    # -------------------------------------------------------------------------
    Context "Invoke-ForceInstallPrerequisitePhase" {
        BeforeEach {
            Mock -ModuleName InstallWinSCPForBizTalk.Workflow Write-InstallerBangError {}
        }

        It "sets PrerequisiteFailure and stops Continue for non-admin ForceInstall without WhatIf" {
            $ctx = script:New-TestContext
            $ctx.isAdministrator = $false
            $ctx.ForceInstall = $true

            Invoke-ForceInstallPrerequisitePhase -Context $ctx -WhatIf $false

            $ctx.PrerequisiteFailure | Should -BeTrue
            $ctx.Continue | Should -BeFalse
            Should -Invoke Write-InstallerBangError -ModuleName InstallWinSCPForBizTalk.Workflow -Times 1
        }
    }

    # -------------------------------------------------------------------------
    Context "Invoke-CuDetectionPhase" {
        BeforeEach {
            Mock -ModuleName InstallWinSCPForBizTalk.Workflow Write-InstallerSectionHeader {}
            Mock -ModuleName InstallWinSCPForBizTalk.Workflow Write-BizTalkCuDetectionStart {}
            Mock -ModuleName InstallWinSCPForBizTalk.Workflow Write-InstallerSuccess {}
            Mock -ModuleName InstallWinSCPForBizTalk.Workflow Write-InstallerBangError {}
            Mock -ModuleName InstallWinSCPForBizTalk.Workflow Write-BizTalkCuSearchSnapshot {}
        }

        It "stops Continue when WinSCP version cannot be determined for the BizTalk version" {
            Mock -ModuleName InstallWinSCPForBizTalk.Workflow Get-WinSCPVersionForBizTalk { return $null }

            $ctx = script:New-TestContext
            Invoke-CuDetectionPhase -Context $ctx

            $ctx.Continue | Should -BeFalse
            Should -Invoke Write-InstallerBangError -ModuleName InstallWinSCPForBizTalk.Workflow -Times 1
        }
    }

    # -------------------------------------------------------------------------
    Context "Invoke-ExistingWinSCPCheckPhase" {
        BeforeEach {
            Mock -ModuleName InstallWinSCPForBizTalk.Workflow Write-InstallerDelimitedMessage {}
            Mock -ModuleName InstallWinSCPForBizTalk.Workflow Write-InstallerSuccess {}
            Mock -ModuleName InstallWinSCPForBizTalk.Workflow Write-WinSCPNotInstalledNotice {}
            Mock -ModuleName InstallWinSCPForBizTalk.Workflow Write-ExistingWinSCPSnapshot {}
        }

        It "sets btsWinSCPProductInstalledAndCorrect when both target files exist with matching version" {
            $mockFileInfo = [PSCustomObject]@{
                VersionInfo = [PSCustomObject]@{ ProductVersion = '5.15.4' }
            }
            Mock -ModuleName InstallWinSCPForBizTalk.Workflow Test-Path { return $true }
            Mock -ModuleName InstallWinSCPForBizTalk.Workflow Get-Item { return $mockFileInfo }
            Mock -ModuleName InstallWinSCPForBizTalk.Workflow Get-InstallExecutionPlan {
                [PSCustomObject]@{ ShouldReinstall = $false; ShouldSkipBecauseInstalled = $false; CanProceed = $true; RequiresElevationWarning = $false }
            }

            $ctx = script:New-TestContext
            Invoke-ExistingWinSCPCheckPhase -Context $ctx

            $ctx.btsWinSCPProductInstalledAndCorrect | Should -BeTrue
        }

        It "appends .0 suffix to required version when installed file version has trailing zero" {
            $mockFileInfo = [PSCustomObject]@{
                VersionInfo = [PSCustomObject]@{ ProductVersion = '5.15.4.0' }
            }
            Mock -ModuleName InstallWinSCPForBizTalk.Workflow Test-Path { return $true }
            Mock -ModuleName InstallWinSCPForBizTalk.Workflow Get-Item { return $mockFileInfo }
            Mock -ModuleName InstallWinSCPForBizTalk.Workflow Get-InstallExecutionPlan {
                [PSCustomObject]@{ ShouldReinstall = $false; ShouldSkipBecauseInstalled = $false; CanProceed = $true; RequiresElevationWarning = $false }
            }

            $ctx = script:New-TestContext
            Invoke-ExistingWinSCPCheckPhase -Context $ctx

            $ctx.winSCPProductVersionRequired | Should -Be '5.15.4.0'
            $ctx.btsWinSCPProductInstalledAndCorrect | Should -BeTrue
        }

        It "resets btsWinSCPProductInstalledAndCorrect when ShouldReinstall is set" {
            $mockFileInfo = [PSCustomObject]@{
                VersionInfo = [PSCustomObject]@{ ProductVersion = '5.15.4' }
            }
            Mock -ModuleName InstallWinSCPForBizTalk.Workflow Test-Path { return $true }
            Mock -ModuleName InstallWinSCPForBizTalk.Workflow Get-Item { return $mockFileInfo }
            Mock -ModuleName InstallWinSCPForBizTalk.Workflow Get-InstallExecutionPlan {
                [PSCustomObject]@{ ShouldReinstall = $true; ShouldSkipBecauseInstalled = $false; CanProceed = $true; RequiresElevationWarning = $false }
            }

            $ctx = script:New-TestContext
            Invoke-ExistingWinSCPCheckPhase -Context $ctx

            $ctx.btsWinSCPProductInstalledAndCorrect | Should -BeFalse
        }

        It "stops Continue when ShouldSkipBecauseInstalled is set" {
            $mockFileInfo = [PSCustomObject]@{
                VersionInfo = [PSCustomObject]@{ ProductVersion = '5.15.4' }
            }
            Mock -ModuleName InstallWinSCPForBizTalk.Workflow Test-Path { return $true }
            Mock -ModuleName InstallWinSCPForBizTalk.Workflow Get-Item { return $mockFileInfo }
            Mock -ModuleName InstallWinSCPForBizTalk.Workflow Get-InstallExecutionPlan {
                [PSCustomObject]@{ ShouldReinstall = $false; ShouldSkipBecauseInstalled = $true; CanProceed = $false; RequiresElevationWarning = $false }
            }

            $ctx = script:New-TestContext
            Invoke-ExistingWinSCPCheckPhase -Context $ctx

            $ctx.Continue | Should -BeFalse
        }
    }

    # -------------------------------------------------------------------------
    Context "Invoke-VersionValidationPhase" {
        BeforeEach {
            Mock -ModuleName InstallWinSCPForBizTalk.Workflow Write-InstallerError {}
        }

        It "stops Continue and writes MissingVersion error when version is null" {
            $ctx = script:New-TestContext
            $ctx.winSCPVersion = $null

            Invoke-VersionValidationPhase -Context $ctx

            $ctx.Continue | Should -BeFalse
            Should -Invoke Write-InstallerError -ModuleName InstallWinSCPForBizTalk.Workflow -Times 3
        }

        It "stops Continue and writes InvalidFormat error when version string is malformed" {
            $ctx = script:New-TestContext
            $ctx.winSCPVersion = 'not-a-version'

            Invoke-VersionValidationPhase -Context $ctx

            $ctx.Continue | Should -BeFalse
            Should -Invoke Write-InstallerError -ModuleName InstallWinSCPForBizTalk.Workflow -Times 3
        }
    }

    # -------------------------------------------------------------------------
    Context "Invoke-DownloadFolderPreparationPhase" {
        BeforeEach {
            Mock -ModuleName InstallWinSCPForBizTalk.Workflow Write-InstallerSectionHeader {}
            Mock -ModuleName InstallWinSCPForBizTalk.Workflow Write-InstallerSuccess {}
            Mock -ModuleName InstallWinSCPForBizTalk.Workflow Write-InstallerError {}
            Mock -ModuleName InstallWinSCPForBizTalk.Workflow Write-TargetFolderPreparationSnapshot {}
            Mock -ModuleName InstallWinSCPForBizTalk.Workflow Resolve-WinSCPPackageLayout {
                [PSCustomObject]@{ ExePath = $null; DllPath = $null; IsResolved = $false }
            }
        }

        It "stops Continue when folder creation is attempted via ShouldProcess but the folder does not exist afterward" {
            Mock -ModuleName InstallWinSCPForBizTalk.Workflow Test-Path { return $false }
            Mock -ModuleName InstallWinSCPForBizTalk.Workflow New-Item {}

            $ctx = script:New-TestContext
            Invoke-DownloadFolderPreparationPhase -Context $ctx

            $ctx.Continue | Should -BeFalse
            Should -Invoke Write-InstallerError -ModuleName InstallWinSCPForBizTalk.Workflow -Times 1
        }
    }

    # -------------------------------------------------------------------------
    Context "Invoke-NuGetDownloadPhase" {
        BeforeEach {
            Mock -ModuleName InstallWinSCPForBizTalk.Workflow Write-InstallerSuccess {}
            Mock -ModuleName InstallWinSCPForBizTalk.Workflow Write-InstallerDelimitedMessage {}
            Mock -ModuleName InstallWinSCPForBizTalk.Workflow Write-InstallerBangError {}
            Mock -ModuleName InstallWinSCPForBizTalk.Workflow Write-NuGetDownloadSnapshot {}
        }

        It "stops Continue when nuget.exe download runs but the file is still absent afterward" {
            Mock -ModuleName InstallWinSCPForBizTalk.Workflow Test-Path { return $false }

            $ctx = script:New-TestContext
            Invoke-NuGetDownloadPhase -Context $ctx

            $ctx.Continue | Should -BeFalse
            Should -Invoke Write-InstallerBangError -ModuleName InstallWinSCPForBizTalk.Workflow -Times 1
        }
    }

    # -------------------------------------------------------------------------
    Context "Invoke-WinSCPPackageDownloadPhase" {
        BeforeEach {
            Mock -ModuleName InstallWinSCPForBizTalk.Workflow Write-InstallerDelimitedMessage {}
            Mock -ModuleName InstallWinSCPForBizTalk.Workflow Write-InstallerBangError {}
            Mock -ModuleName InstallWinSCPForBizTalk.Workflow Write-WinSCPDownloadSnapshot {}
            Mock -ModuleName InstallWinSCPForBizTalk.Workflow Resolve-WinSCPPackageLayout {
                [PSCustomObject]@{ ExePath = $null; DllPath = $null; IsResolved = $false }
            }
        }

        It "stops Continue when package files are not found after the nuget download command runs" {
            $ctx = script:New-TestContext
            Invoke-WinSCPPackageDownloadPhase -Context $ctx

            $ctx.Continue | Should -BeFalse
            Should -Invoke Write-InstallerBangError -ModuleName InstallWinSCPForBizTalk.Workflow -Times 1
        }
    }

    # -------------------------------------------------------------------------
    Context "Invoke-WinSCPCopyPhase" {
        BeforeEach {
            Mock -ModuleName InstallWinSCPForBizTalk.Workflow Write-InstallerSectionHeader {}
            Mock -ModuleName InstallWinSCPForBizTalk.Workflow Write-InstallerSuccess {}
            Mock -ModuleName InstallWinSCPForBizTalk.Workflow Write-InstallerError {}
            Mock -ModuleName InstallWinSCPForBizTalk.Workflow Write-WinSCPCopySnapshot {}
            Mock -ModuleName InstallWinSCPForBizTalk.Workflow Test-Path { return $false }
        }

        It "sets btsWinSCPProductInstalledAndCorrect when Invoke-WinSCPTargetInstall reports success" {
            Mock -ModuleName InstallWinSCPForBizTalk.Workflow Invoke-WinSCPTargetInstall {
                [PSCustomObject]@{
                    InstalledSuccessfully = $true
                    TargetExeExists       = $true
                    TargetDllExists       = $true
                    ErrorMessages         = @()
                }
            }

            $ctx = script:New-TestContext
            $ctx.WinSCPEXEDownload = 'C:\fake\WinSCP.exe'
            $ctx.WinSCPDllDownload = 'C:\fake\WinSCPnet.dll'

            Invoke-WinSCPCopyPhase -Context $ctx

            $ctx.btsWinSCPProductInstalledAndCorrect | Should -BeTrue
        }
    }
}
