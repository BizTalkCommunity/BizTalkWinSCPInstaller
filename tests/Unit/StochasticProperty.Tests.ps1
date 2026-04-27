<#
.SYNOPSIS
Seeded stochastic/property tests for core helper behavior.

.DESCRIPTION
Uses deterministic randomized inputs (fixed seed) to exercise broader valid,
invalid, normal, and abnormal paths while remaining repeatable in CI.
#>

if (-not (Get-Module -ListAvailable Pester | Where-Object { $_.Version.Major -ge 5 })) {
    throw "Pester 5+ is required to run this test file. Install with: Install-Module Pester -Scope CurrentUser -RequiredVersion 5.0 -Force"
}

Describe "Stochastic property tests" {
    BeforeAll {
        $modulePath = Join-Path (Split-Path -Parent $PSCommandPath) "../../src/InstallWinSCPForBizTalk.Core.psm1"
        Import-Module $modulePath -Force
        $script:rng = [System.Random]::new(20260424)
    }

    AfterAll {
        Remove-Module InstallWinSCPForBizTalk.Core -ErrorAction SilentlyContinue
    }

    Context "Test-WinSCPVersionString properties" {
        It "matches TryParse/whitespace expectations across randomized inputs" {
            $alphabet = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789.-_+xv ".ToCharArray()

            1..200 | ForEach-Object {
                $roll = $script:rng.Next(0, 10)
                $candidate = $null

                if ($roll -eq 0) {
                    $candidate = $null
                }
                elseif ($roll -eq 1) {
                    $candidate = " " * $script:rng.Next(1, 5)
                }
                elseif ($roll -le 4) {
                    # Mostly-valid generated dotted numeric versions.
                    $major = $script:rng.Next(0, 20)
                    $minor = $script:rng.Next(0, 20)
                    $build = $script:rng.Next(0, 20)
                    if ($script:rng.Next(0, 2) -eq 0) {
                        $candidate = "$major.$minor.$build"
                    }
                    else {
                        $revision = $script:rng.Next(0, 20)
                        $candidate = "$major.$minor.$build.$revision"
                    }
                }
                else {
                    # Free-form random strings that are often invalid.
                    $len = $script:rng.Next(1, 18)
                    $chars = 1..$len | ForEach-Object { $alphabet[$script:rng.Next(0, $alphabet.Length)] }
                    $candidate = -join $chars
                }

                $result = Test-WinSCPVersionString -WinSCPVersion $candidate
                $parsed = $null
                $isMissing = [string]::IsNullOrWhiteSpace($candidate)
                $isParsable = if ($isMissing) { $false } else { [version]::TryParse($candidate, [ref]$parsed) }

                if ($isMissing) {
                    $result.IsValid | Should -BeFalse
                    $result.ErrorCode | Should -Be "MissingVersion"
                    $result.ParsedVersion | Should -Be $null
                }
                elseif ($isParsable) {
                    $result.IsValid | Should -BeTrue
                    $result.ErrorCode | Should -Be "None"
                    $result.ParsedVersion | Should -Not -Be $null
                }
                else {
                    $result.IsValid | Should -BeFalse
                    $result.ErrorCode | Should -Be "InvalidFormat"
                    $result.ParsedVersion | Should -Be $null
                }
            }
        }
    }

    Context "Get-PackageReadinessState properties" {
        It "returns expected state across randomized boolean combinations" {
            1..200 | ForEach-Object {
                $nuget = [bool]$script:rng.Next(0, 2)
                $exe = [bool]$script:rng.Next(0, 2)
                $dll = [bool]$script:rng.Next(0, 2)

                $result = Get-PackageReadinessState -NuGetExeExists $nuget -WinSCPExeExists $exe -WinSCPDllExists $dll

                $expectedState = if (-not $nuget) {
                    "MissingNuGet"
                }
                elseif (-not $exe -and -not $dll) {
                    "MissingWinSCPPackage"
                }
                elseif (-not $exe -or -not $dll) {
                    "IncompleteWinSCPPackage"
                }
                else {
                    "Ready"
                }

                $result.State | Should -Be $expectedState
                $result.IsReady | Should -Be ($expectedState -eq "Ready")
            }
        }
    }

    Context "Get-InstallExecutionPlan properties" {
        It "preserves key invariants across randomized boolean combinations" {
            1..200 | ForEach-Object {
                $isAdmin = [bool]$script:rng.Next(0, 2)
                $force = [bool]$script:rng.Next(0, 2)
                $whatIf = [bool]$script:rng.Next(0, 2)
                $installed = [bool]$script:rng.Next(0, 2)

                $plan = Get-InstallExecutionPlan -IsAdministrator $isAdmin -ForceInstall $force -WhatIf $whatIf -AlreadyInstalledCorrect $installed

                if (-not $isAdmin -and $force -and -not $whatIf) {
                    $plan.CanProceed | Should -BeFalse
                    $plan.PrerequisiteFailure | Should -BeTrue
                }

                if ($installed -and -not $force) {
                    $plan.CanProceed | Should -BeFalse
                    $plan.ShouldSkipBecauseInstalled | Should -BeTrue
                }

                if ($installed -and $force) {
                    $plan.ShouldReinstall | Should -BeTrue
                }

                if ($plan.PrerequisiteFailure) {
                    $plan.CanProceed | Should -BeFalse
                }
            }
        }
    }

    Context "Get-FinalExecutionOutcome properties" {
        It "applies precedence rules consistently across randomized terminal states" {
            1..200 | ForEach-Object {
                $installed = [bool]$script:rng.Next(0, 2)
                $whatIf = [bool]$script:rng.Next(0, 2)
                $prereq = [bool]$script:rng.Next(0, 2)
                $continueFlag = [bool]$script:rng.Next(0, 2)

                $result = Get-FinalExecutionOutcome -InstalledSuccessfully $installed -WhatIf $whatIf -PrerequisiteFailure $prereq -ContinueFlag $continueFlag

                $expectedOutcome = if ($installed) {
                    "Success"
                }
                elseif ($whatIf) {
                    "DryRun"
                }
                elseif ($prereq) {
                    "PrerequisiteFailure"
                }
                elseif (-not $continueFlag) {
                    "InstallFailure"
                }
                else {
                    "Unknown"
                }

                $result.Outcome | Should -Be $expectedOutcome
                $result.IsError | Should -Be ($expectedOutcome -in @("PrerequisiteFailure", "InstallFailure", "Unknown"))
            }
        }
    }

    Context "Resolve-BizTalkInstallFolder properties" {
        BeforeEach {
            Mock -ModuleName InstallWinSCPForBizTalk.Core Test-Path {
                param([string]$Path)
                return ($Path -like "C:\\exists*")
            }
        }

        It "honors environment-over-registry precedence across randomized inputs" {
            $candidates = @(
                $null,
                "",
                "   ",
                "C:\exists\\BizTalk",
                "C:\missing\\BizTalk",
                "C:\exists\\BizTalk\\Alt",
                "C:\missing\\BizTalk\\Alt"
            )

            1..200 | ForEach-Object {
                $envPath = $candidates[$script:rng.Next(0, $candidates.Count)]
                $regPath = $candidates[$script:rng.Next(0, $candidates.Count)]

                $result = Resolve-BizTalkInstallFolder -EnvironmentInstallPath $envPath -RegistryInstallPath $regPath

                $expectedPath = if (-not [string]::IsNullOrWhiteSpace($envPath)) {
                    $envPath
                }
                elseif (-not [string]::IsNullOrWhiteSpace($regPath)) {
                    $regPath
                }
                else {
                    $null
                }

                $expectedSource = if (-not [string]::IsNullOrWhiteSpace($envPath)) {
                    'Environment'
                }
                elseif (-not [string]::IsNullOrWhiteSpace($regPath)) {
                    'Registry'
                }
                else {
                    'None'
                }

                $result.InstallPath | Should -Be $expectedPath
                $result.Source | Should -Be $expectedSource
                $result.IsFound | Should -Be (-not [string]::IsNullOrWhiteSpace($expectedPath))
                $result.Exists | Should -Be ([bool]($expectedPath -and $expectedPath -like "C:\\exists*"))
            }
        }
    }

    Context "Get-BizTalkVersionFromProductCode properties" {
        It "maps known codes and rejects random unknown codes" {
            $known2020 = '{205F5836-7512-4A06-9E74-ADC8AFA0EEC5}'
            $known2016 = '{B084F3A7-3E8F-4E7B-B673-EED1715D28ED}'

            Get-BizTalkVersionFromProductCode -ProductCode $known2020 | Should -Be '2020'
            Get-BizTalkVersionFromProductCode -ProductCode $known2016 | Should -Be '2016'

            1..100 | ForEach-Object {
                $unknown = ('{' + [guid]::NewGuid().ToString().ToUpperInvariant() + '}')
                if ($unknown -ne $known2020 -and $unknown -ne $known2016) {
                    Get-BizTalkVersionFromProductCode -ProductCode $unknown | Should -Be $null
                }
            }
        }
    }
}
