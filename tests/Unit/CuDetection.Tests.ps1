<#
.SYNOPSIS
Unit tests for BizTalk CU detection and WinSCP version mapping helpers.

.DESCRIPTION
Validates KB/display-name CU detection logic and mapping behavior used to select
the required WinSCP version for BizTalk 2016/2020 scenarios.
#>

if (-not (Get-Module -ListAvailable Pester | Where-Object { $_.Version.Major -ge 5 })) {
    throw "Pester 5+ is required to run this test file. Install with: Install-Module Pester -Scope CurrentUser -RequiredVersion 5.0 -Force"
}

Describe "CU detection functions" {
    BeforeAll {
        $modulePath = Join-Path (Split-Path -Parent $PSCommandPath) "../../src/InstallWinSCPForBizTalk.Core.psm1"
        Import-Module $modulePath -Force

        # Mirrors current script map for BizTalk 2020 CU selection.
        $script:Bts2020Map = @{
            6 = @{ KBs = @("5043408", "5048971"); WinSCP = "6.3.5" }
            5 = @{ KBs = @("5032870"); WinSCP = "6.1.2" }
            4 = @{ KBs = @("5009901"); WinSCP = "5.19.2" }
            3 = @{ KBs = @("5007969"); WinSCP = "5.19.2" }
            2 = @{ KBs = @("5003151"); WinSCP = "5.17.8" }
            1 = @{ KBs = @("4538666"); WinSCP = "5.17.6" }
        }

        function script:Get-MappedWinSCP2020 {
            param($DetectedCU)
            if ($DetectedCU.Found -and $script:Bts2020Map.ContainsKey([int]$DetectedCU.CUNumber)) {
                return $script:Bts2020Map[[int]$DetectedCU.CUNumber].WinSCP
            }
            return "5.15.4"
        }
    }

    AfterAll {
        Remove-Module InstallWinSCPForBizTalk.Core -ErrorAction SilentlyContinue
        Remove-Item function:\Get-MappedWinSCP2020 -ErrorAction SilentlyContinue
    }

    Context "Search-BTSCumulativeUpdate" {
        It "finds BizTalk 2020 CU6 by KB 5048971" {
            Mock -ModuleName InstallWinSCPForBizTalk.Core Get-ItemProperty {
                @(
                    [pscustomobject]@{ DisplayName = "Microsoft BizTalk Server 2020 Cumulative Update 6 KB5048971" }
                )
            }

            $found = Search-BTSCumulativeUpdate -CumulativeUpdateID "5048971" -BizTalkVersion "2020"
            $found | Should -BeTrue
        }

        It "finds BizTalk 2020 CU6 by alternate KB 5043408" {
            Mock -ModuleName InstallWinSCPForBizTalk.Core Get-ItemProperty {
                @(
                    [pscustomobject]@{ DisplayName = "Microsoft BizTalk Server 2020 Cumulative Update 6 KB5043408" }
                )
            }

            $found = Search-BTSCumulativeUpdate -CumulativeUpdateID "5043408" -BizTalkVersion "2020"
            $found | Should -BeTrue
        }

        It "returns false when uninstall entries have empty display names" {
            Mock -ModuleName InstallWinSCPForBizTalk.Core Get-ItemProperty {
                @(
                    [pscustomobject]@{ DisplayName = $null },
                    [pscustomobject]@{ DisplayName = "   " }
                )
            }

            $found = Search-BTSCumulativeUpdate -CumulativeUpdateID "5043408" -BizTalkVersion "2020"
            $found | Should -BeFalse
        }
    }

    Context "Get-BTSCumulativeUpdateByDisplayName" {
        It "parses Cumulative Update wording and maps CU6 to WinSCP 6.3.5" {
            Mock -ModuleName InstallWinSCPForBizTalk.Core Get-ItemProperty {
                @(
                    [pscustomobject]@{
                        DisplayName = "Microsoft BizTalk Server 2020 Cumulative Update 6 KB5048971"
                        InstallDate = "20241121"
                    }
                )
            }

            $detected = Get-BTSCumulativeUpdateByDisplayName -BizTalkVersion "2020"
            $detected.Found | Should -BeTrue
            $detected.CUNumber | Should -Be 6
            $detected.KB | Should -Be "5048971"
            (Get-MappedWinSCP2020 -DetectedCU $detected) | Should -Be "6.3.5"
        }

        It "parses CUX shorthand and maps CU6 to WinSCP 6.3.5" {
            Mock -ModuleName InstallWinSCPForBizTalk.Core Get-ItemProperty {
                @(
                    [pscustomobject]@{
                        DisplayName = "BizTalk Server 2020 CU6 KB5043408"
                        InstallDate = "20241120"
                    }
                )
            }

            $detected = Get-BTSCumulativeUpdateByDisplayName -BizTalkVersion "2020"
            $detected.Found | Should -BeTrue
            $detected.CUNumber | Should -Be 6
            $detected.KB | Should -Be "5043408"
            (Get-MappedWinSCP2020 -DetectedCU $detected) | Should -Be "6.3.5"
        }

        It "returns Found false for non-matching or empty uninstall entries" {
            Mock -ModuleName InstallWinSCPForBizTalk.Core Get-ItemProperty {
                @(
                    [pscustomobject]@{ DisplayName = $null; InstallDate = $null },
                    [pscustomobject]@{ DisplayName = "Unrelated Product"; InstallDate = "20240101" }
                )
            }

            $detected = Get-BTSCumulativeUpdateByDisplayName -BizTalkVersion "2020"
            $detected.Found | Should -BeFalse
            $detected.CUNumber | Should -Be 0
            (Get-MappedWinSCP2020 -DetectedCU $detected) | Should -Be "5.15.4"
        }

        It "selects highest CU when multiple are present (2016 CU9 over CU8)" {
            Mock -ModuleName InstallWinSCPForBizTalk.Core Get-ItemProperty {
                @(
                    [pscustomobject]@{
                        DisplayName = "Microsoft BizTalk Server 2016 Cumulative Update 8 KB4583530"
                        InstallDate = "20201207"
                    },
                    [pscustomobject]@{
                        DisplayName = "Microsoft BizTalk Server 2016 Cumulative Update 9 KB5005480"
                        InstallDate = "20210929"
                    }
                )
            }

            $detected = Get-BTSCumulativeUpdateByDisplayName -BizTalkVersion "2016"
            $detected.Found | Should -BeTrue
            $detected.CUNumber | Should -Be 9
            $detected.KB | Should -Be "5005480"
        }
    }
}
