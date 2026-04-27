<#
.SYNOPSIS
Unit tests for target installation copy/verification helper.

.DESCRIPTION
Validates copy behavior, post-copy target verification, and user-facing error
messages returned by Invoke-WinSCPTargetInstall.
#>

if (-not (Get-Module -ListAvailable Pester | Where-Object { $_.Version.Major -ge 5 })) {
    throw "Pester 5+ is required to run this test file. Install with: Install-Module Pester -Scope CurrentUser -RequiredVersion 5.0 -Force"
}

Describe "Invoke-WinSCPTargetInstall" {
    BeforeAll {
        $modulePath = Join-Path (Split-Path -Parent $PSCommandPath) "../../src/InstallWinSCPForBizTalk.Core.psm1"
        Import-Module $modulePath -Force
    }

    AfterAll {
        Remove-Module InstallWinSCPForBizTalk.Core -ErrorAction SilentlyContinue
    }

    It "returns success when both target files exist after copy" {
        Mock -ModuleName InstallWinSCPForBizTalk.Core Copy-Item { }
        Mock -ModuleName InstallWinSCPForBizTalk.Core Test-Path { $true }

        $result = Invoke-WinSCPTargetInstall -SourceExePath 'C:\pkg\WinSCP.exe' -SourceDllPath 'C:\pkg\WinSCPnet.dll' -TargetFolder 'C:\BizTalk' -TargetExePath 'C:\BizTalk\WinSCP.exe' -TargetDllPath 'C:\BizTalk\WinSCPnet.dll' -WinSCPVersion '5.15.4' -ExeFileName 'WinSCP.exe' -DllFileName 'WinSCPnet.dll'

        $result.InstalledSuccessfully | Should -BeTrue
        $result.TargetExeExists | Should -BeTrue
        $result.TargetDllExists | Should -BeTrue
        $result.ErrorMessages.Count | Should -Be 0
    }

    It "returns failure with exe message when exe is missing" {
        Mock -ModuleName InstallWinSCPForBizTalk.Core Copy-Item { }
        Mock -ModuleName InstallWinSCPForBizTalk.Core Test-Path {
            param([string]$Path)
            if ($Path -eq 'C:\BizTalk\WinSCP.exe') { return $false }
            return $true
        }

        $result = Invoke-WinSCPTargetInstall -SourceExePath 'C:\pkg\WinSCP.exe' -SourceDllPath 'C:\pkg\WinSCPnet.dll' -TargetFolder 'C:\BizTalk' -TargetExePath 'C:\BizTalk\WinSCP.exe' -TargetDllPath 'C:\BizTalk\WinSCPnet.dll' -WinSCPVersion '5.15.4' -ExeFileName 'WinSCP.exe' -DllFileName 'WinSCPnet.dll'

        $result.InstalledSuccessfully | Should -BeFalse
        $result.TargetExeExists | Should -BeFalse
        $result.TargetDllExists | Should -BeTrue
        ($result.ErrorMessages -join "`
") | Should -Match 'WinSCP\.exe file version 5\.15\.4'
    }

    It "returns failure with generic copy error when copy throws" {
        Mock -ModuleName InstallWinSCPForBizTalk.Core Copy-Item { throw 'copy failed' }

        $result = Invoke-WinSCPTargetInstall -SourceExePath 'C:\pkg\WinSCP.exe' -SourceDllPath 'C:\pkg\WinSCPnet.dll' -TargetFolder 'C:\BizTalk' -TargetExePath 'C:\BizTalk\WinSCP.exe' -TargetDllPath 'C:\BizTalk\WinSCPnet.dll' -WinSCPVersion '5.15.4' -ExeFileName 'WinSCP.exe' -DllFileName 'WinSCPnet.dll'

        $result.InstalledSuccessfully | Should -BeFalse
        $result.TargetExeExists | Should -BeFalse
        $result.TargetDllExists | Should -BeFalse
        $result.ErrorMessages[0] | Should -Be 'Failed to copy WinSCP files to the BizTalk installation folder.'
        $result.ErrorMessages[1] | Should -Match 'copy failed'
    }
}
