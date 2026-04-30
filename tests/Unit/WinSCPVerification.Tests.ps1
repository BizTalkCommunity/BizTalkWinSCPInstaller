<#
.SYNOPSIS
Unit tests for post-install WinSCP verification logic.

.DESCRIPTION
Covers Get-WinSCPInstallVerification (Core) and Invoke-WinSCPVerificationPhase
(Workflow): version matching, .0 suffix normalization, hash check, skip behavior,
and workflow context recording.
#>

if (-not (Get-Module -ListAvailable Pester | Where-Object { $_.Version.Major -ge 5 })) {
    throw "Pester 5+ is required to run this test file. Install with: Install-Module Pester -Scope CurrentUser -RequiredVersion 5.0 -Force"
}

Describe "Get-WinSCPInstallVerification" {
    BeforeAll {
        $modulePath = Join-Path (Split-Path -Parent $PSCommandPath) "../../src/InstallWinSCPForBizTalk.Core.psm1"
        Import-Module $modulePath -Force
    }

    AfterAll {
        Remove-Module InstallWinSCPForBizTalk.Core -ErrorAction SilentlyContinue
    }

    It "returns VerifiedSuccessfully true when both file versions match" {
        Mock -ModuleName InstallWinSCPForBizTalk.Core Test-Path { $true }
        Mock -ModuleName InstallWinSCPForBizTalk.Core Get-Item {
            [pscustomobject]@{ VersionInfo = [pscustomobject]@{ ProductVersion = '5.19.2' } }
        }

        $result = Get-WinSCPInstallVerification `
            -TargetExePath 'C:\BizTalk\WinSCP.exe' `
            -TargetDllPath 'C:\BizTalk\WinSCPnet.dll' `
            -ExpectedVersion '5.19.2'

        $result.ExeVersionMatch      | Should -BeTrue
        $result.DllVersionMatch      | Should -BeTrue
        $result.VerifiedSuccessfully | Should -BeTrue
    }

    It "normalizes trailing .0 suffix when installed version has it and expected does not" {
        Mock -ModuleName InstallWinSCPForBizTalk.Core Test-Path { $true }
        Mock -ModuleName InstallWinSCPForBizTalk.Core Get-Item {
            [pscustomobject]@{ VersionInfo = [pscustomobject]@{ ProductVersion = '5.19.2.0' } }
        }

        $result = Get-WinSCPInstallVerification `
            -TargetExePath 'C:\BizTalk\WinSCP.exe' `
            -TargetDllPath 'C:\BizTalk\WinSCPnet.dll' `
            -ExpectedVersion '5.19.2'

        $result.ExeVersionMatch      | Should -BeTrue
        $result.DllVersionMatch      | Should -BeTrue
        $result.VerifiedSuccessfully | Should -BeTrue
    }

    It "does not normalize when both expected and installed share the .0 suffix" {
        Mock -ModuleName InstallWinSCPForBizTalk.Core Test-Path { $true }
        Mock -ModuleName InstallWinSCPForBizTalk.Core Get-Item {
            [pscustomobject]@{ VersionInfo = [pscustomobject]@{ ProductVersion = '5.19.2.0' } }
        }

        $result = Get-WinSCPInstallVerification `
            -TargetExePath 'C:\BizTalk\WinSCP.exe' `
            -TargetDllPath 'C:\BizTalk\WinSCPnet.dll' `
            -ExpectedVersion '5.19.2.0'

        $result.ExeVersionMatch      | Should -BeTrue
        $result.DllVersionMatch      | Should -BeTrue
        $result.VerifiedSuccessfully | Should -BeTrue
    }

    It "returns ExeVersionMatch false when EXE version is wrong" {
        Mock -ModuleName InstallWinSCPForBizTalk.Core Test-Path { $true }
        Mock -ModuleName InstallWinSCPForBizTalk.Core Get-Item {
            param($Path)
            if ($Path -like '*WinSCP.exe') {
                [pscustomobject]@{ VersionInfo = [pscustomobject]@{ ProductVersion = '5.15.4' } }
            }
            else {
                [pscustomobject]@{ VersionInfo = [pscustomobject]@{ ProductVersion = '5.19.2' } }
            }
        }

        $result = Get-WinSCPInstallVerification `
            -TargetExePath 'C:\BizTalk\WinSCP.exe' `
            -TargetDllPath 'C:\BizTalk\WinSCPnet.dll' `
            -ExpectedVersion '5.19.2'

        $result.ExeVersionMatch      | Should -BeFalse
        $result.DllVersionMatch      | Should -BeTrue
        $result.VerifiedSuccessfully | Should -BeFalse
    }

    It "returns DllVersionMatch false when DLL version is wrong" {
        Mock -ModuleName InstallWinSCPForBizTalk.Core Test-Path { $true }
        Mock -ModuleName InstallWinSCPForBizTalk.Core Get-Item {
            param($Path)
            if ($Path -like '*WinSCPnet.dll') {
                [pscustomobject]@{ VersionInfo = [pscustomobject]@{ ProductVersion = '5.15.4' } }
            }
            else {
                [pscustomobject]@{ VersionInfo = [pscustomobject]@{ ProductVersion = '5.19.2' } }
            }
        }

        $result = Get-WinSCPInstallVerification `
            -TargetExePath 'C:\BizTalk\WinSCP.exe' `
            -TargetDllPath 'C:\BizTalk\WinSCPnet.dll' `
            -ExpectedVersion '5.19.2'

        $result.ExeVersionMatch      | Should -BeTrue
        $result.DllVersionMatch      | Should -BeFalse
        $result.VerifiedSuccessfully | Should -BeFalse
    }

    It "returns ExeVersionMatch false and ExeActualVersion 'not found' when EXE file is missing" {
        Mock -ModuleName InstallWinSCPForBizTalk.Core Test-Path {
            param($Path)
            return $Path -notlike '*WinSCP.exe'
        }
        Mock -ModuleName InstallWinSCPForBizTalk.Core Get-Item {
            [pscustomobject]@{ VersionInfo = [pscustomobject]@{ ProductVersion = '5.19.2' } }
        }

        $result = Get-WinSCPInstallVerification `
            -TargetExePath 'C:\BizTalk\WinSCP.exe' `
            -TargetDllPath 'C:\BizTalk\WinSCPnet.dll' `
            -ExpectedVersion '5.19.2'

        $result.ExeVersionMatch      | Should -BeFalse
        $result.ExeActualVersion     | Should -Be 'not found'
        $result.VerifiedSuccessfully | Should -BeFalse
    }

    It "returns DllVersionMatch false and DllActualVersion 'not found' when DLL file is missing" {
        Mock -ModuleName InstallWinSCPForBizTalk.Core Test-Path {
            param($Path)
            return $Path -notlike '*WinSCPnet.dll'
        }
        Mock -ModuleName InstallWinSCPForBizTalk.Core Get-Item {
            [pscustomobject]@{ VersionInfo = [pscustomobject]@{ ProductVersion = '5.19.2' } }
        }

        $result = Get-WinSCPInstallVerification `
            -TargetExePath 'C:\BizTalk\WinSCP.exe' `
            -TargetDllPath 'C:\BizTalk\WinSCPnet.dll' `
            -ExpectedVersion '5.19.2'

        $result.DllVersionMatch      | Should -BeFalse
        $result.DllActualVersion     | Should -Be 'not found'
        $result.VerifiedSuccessfully | Should -BeFalse
    }

    It "does not call Get-FileHash and leaves hash results null when CheckHash is false" {
        Mock -ModuleName InstallWinSCPForBizTalk.Core Test-Path { $true }
        Mock -ModuleName InstallWinSCPForBizTalk.Core Get-Item {
            [pscustomobject]@{ VersionInfo = [pscustomobject]@{ ProductVersion = '5.19.2' } }
        }
        Mock -ModuleName InstallWinSCPForBizTalk.Core Get-FileHash { }

        $result = Get-WinSCPInstallVerification `
            -TargetExePath 'C:\BizTalk\WinSCP.exe' `
            -TargetDllPath 'C:\BizTalk\WinSCPnet.dll' `
            -ExpectedVersion '5.19.2' `
            -CheckHash $false

        $result.CheckedHash | Should -BeFalse
        $result.ExeHashMatch | Should -BeNullOrEmpty
        $result.DllHashMatch | Should -BeNullOrEmpty
        Should -Not -Invoke Get-FileHash -ModuleName InstallWinSCPForBizTalk.Core
    }

    It "returns ExeHashMatch and DllHashMatch true when hashes match source" {
        Mock -ModuleName InstallWinSCPForBizTalk.Core Test-Path { $true }
        Mock -ModuleName InstallWinSCPForBizTalk.Core Get-Item {
            [pscustomobject]@{ VersionInfo = [pscustomobject]@{ ProductVersion = '5.19.2' } }
        }
        Mock -ModuleName InstallWinSCPForBizTalk.Core Get-FileHash {
            [pscustomobject]@{ Hash = 'AABBCC1234567890' }
        }

        $result = Get-WinSCPInstallVerification `
            -TargetExePath 'C:\BizTalk\WinSCP.exe' `
            -TargetDllPath 'C:\BizTalk\WinSCPnet.dll' `
            -ExpectedVersion '5.19.2' `
            -SourceExePath 'C:\nuget\WinSCP.exe' `
            -SourceDllPath 'C:\nuget\WinSCPnet.dll' `
            -CheckHash $true

        $result.CheckedHash          | Should -BeTrue
        $result.ExeHashMatch         | Should -BeTrue
        $result.DllHashMatch         | Should -BeTrue
        $result.VerifiedSuccessfully | Should -BeTrue
    }

    It "returns ExeHashMatch false and VerifiedSuccessfully false when EXE hash differs from source" {
        Mock -ModuleName InstallWinSCPForBizTalk.Core Test-Path { $true }
        Mock -ModuleName InstallWinSCPForBizTalk.Core Get-Item {
            [pscustomobject]@{ VersionInfo = [pscustomobject]@{ ProductVersion = '5.19.2' } }
        }
        Mock -ModuleName InstallWinSCPForBizTalk.Core Get-FileHash -ParameterFilter { $Path -like '*nuget*' } {
            [pscustomobject]@{ Hash = 'SOURCE_HASH_AABBCC' }
        }
        Mock -ModuleName InstallWinSCPForBizTalk.Core Get-FileHash -ParameterFilter { $Path -like '*BizTalk*' } {
            [pscustomobject]@{ Hash = 'CORRUPTED_HASH_XXYYZZ' }
        }

        $result = Get-WinSCPInstallVerification `
            -TargetExePath 'C:\BizTalk\WinSCP.exe' `
            -TargetDllPath 'C:\BizTalk\WinSCPnet.dll' `
            -ExpectedVersion '5.19.2' `
            -SourceExePath 'C:\nuget\WinSCP.exe' `
            -SourceDllPath 'C:\nuget\WinSCPnet.dll' `
            -CheckHash $true

        $result.ExeHashMatch         | Should -BeFalse
        $result.VerifiedSuccessfully | Should -BeFalse
    }

    It "skips hash check and does not fail when CheckHash is true but source paths are absent" {
        Mock -ModuleName InstallWinSCPForBizTalk.Core Test-Path { $true }
        Mock -ModuleName InstallWinSCPForBizTalk.Core Get-Item {
            [pscustomobject]@{ VersionInfo = [pscustomobject]@{ ProductVersion = '5.19.2' } }
        }
        Mock -ModuleName InstallWinSCPForBizTalk.Core Get-FileHash { }

        $result = Get-WinSCPInstallVerification `
            -TargetExePath 'C:\BizTalk\WinSCP.exe' `
            -TargetDllPath 'C:\BizTalk\WinSCPnet.dll' `
            -ExpectedVersion '5.19.2' `
            -CheckHash $true

        $result.CheckedHash          | Should -BeFalse
        $result.VerifiedSuccessfully | Should -BeTrue
        Should -Not -Invoke Get-FileHash -ModuleName InstallWinSCPForBizTalk.Core
    }
}

Describe "Invoke-WinSCPVerificationPhase" {
    BeforeAll {
        $corePath     = Join-Path (Split-Path -Parent $PSCommandPath) "../../src/InstallWinSCPForBizTalk.Core.psm1"
        $utilsPath    = Join-Path (Split-Path -Parent $PSCommandPath) "../../src/InstallWinSCPForBizTalk.Utils.psm1"
        $workflowPath = Join-Path (Split-Path -Parent $PSCommandPath) "../../src/InstallWinSCPForBizTalk.Workflow.psm1"
        Import-Module $corePath -Force
        Import-Module $utilsPath -Force
        Import-Module $workflowPath -Force
    }

    AfterAll {
        Remove-Module InstallWinSCPForBizTalk.Core, InstallWinSCPForBizTalk.Utils, InstallWinSCPForBizTalk.Workflow -ErrorAction SilentlyContinue
    }

    It "records verificationResult in context and reports pass when versions match" {
        Mock -ModuleName InstallWinSCPForBizTalk.Core Test-Path { $true }
        Mock -ModuleName InstallWinSCPForBizTalk.Core Get-Item {
            [pscustomobject]@{ VersionInfo = [pscustomobject]@{ ProductVersion = '5.19.2' } }
        }
        Mock -ModuleName InstallWinSCPForBizTalk.Workflow Write-InstallerSectionHeader { }
        Mock -ModuleName InstallWinSCPForBizTalk.Workflow Write-InstallerSuccess { }
        Mock -ModuleName InstallWinSCPForBizTalk.Workflow Write-InstallerError { }
        Mock -ModuleName InstallWinSCPForBizTalk.Workflow Write-InstallerBangError { }
        Mock -ModuleName InstallWinSCPForBizTalk.Workflow Write-WinSCPVersionCheckResults { }
        Mock -ModuleName InstallWinSCPForBizTalk.Workflow Write-WinSCPHashCheckResults { }
        Mock -ModuleName InstallWinSCPForBizTalk.Workflow Write-WinSCPVerificationSnapshot { }

        $ctx = @{
            btsWinSCPProductInstalledAndCorrect = $true
            btsTargetWinSCPExe = 'C:\BizTalk\WinSCP.exe'
            btsTargetWinSCPDll = 'C:\BizTalk\WinSCPnet.dll'
            winSCPVersion      = '5.19.2'
            CheckHash          = $false
            WinSCPEXEDownload  = ''
            WinSCPDllDownload  = ''
        }

        Invoke-WinSCPVerificationPhase -Context $ctx

        $ctx.verificationResult                          | Should -Not -BeNull
        $ctx.verificationResult.VerifiedSuccessfully     | Should -BeTrue
        $ctx.verificationResult.ExeVersionMatch          | Should -BeTrue
        $ctx.verificationResult.DllVersionMatch          | Should -BeTrue
        Should -Not -Invoke Write-InstallerBangError -ModuleName InstallWinSCPForBizTalk.Workflow
    }

    It "records verificationResult with VerifiedSuccessfully false and calls bang error on version mismatch" {
        Mock -ModuleName InstallWinSCPForBizTalk.Core Test-Path { $true }
        Mock -ModuleName InstallWinSCPForBizTalk.Core Get-Item {
            [pscustomobject]@{ VersionInfo = [pscustomobject]@{ ProductVersion = '5.15.4' } }
        }
        Mock -ModuleName InstallWinSCPForBizTalk.Workflow Write-InstallerSectionHeader { }
        Mock -ModuleName InstallWinSCPForBizTalk.Workflow Write-InstallerSuccess { }
        Mock -ModuleName InstallWinSCPForBizTalk.Workflow Write-InstallerError { }
        Mock -ModuleName InstallWinSCPForBizTalk.Workflow Write-InstallerBangError { }
        Mock -ModuleName InstallWinSCPForBizTalk.Workflow Write-WinSCPVersionCheckResults { }
        Mock -ModuleName InstallWinSCPForBizTalk.Workflow Write-WinSCPHashCheckResults { }
        Mock -ModuleName InstallWinSCPForBizTalk.Workflow Write-WinSCPVerificationSnapshot { }

        $ctx = @{
            btsWinSCPProductInstalledAndCorrect = $true
            btsTargetWinSCPExe = 'C:\BizTalk\WinSCP.exe'
            btsTargetWinSCPDll = 'C:\BizTalk\WinSCPnet.dll'
            winSCPVersion      = '5.19.2'
            CheckHash          = $false
            WinSCPEXEDownload  = ''
            WinSCPDllDownload  = ''
        }

        Invoke-WinSCPVerificationPhase -Context $ctx

        $ctx.verificationResult.VerifiedSuccessfully | Should -BeFalse
        Should -Invoke Write-InstallerBangError -Times 1 -ModuleName InstallWinSCPForBizTalk.Workflow
    }

    It "does not run and leaves verificationResult null when install did not succeed" {
        Mock -ModuleName InstallWinSCPForBizTalk.Workflow Write-InstallerSectionHeader { }

        $ctx = @{
            btsWinSCPProductInstalledAndCorrect = $false
        }

        Invoke-WinSCPVerificationPhase -Context $ctx

        $ctx.verificationResult | Should -BeNull
        Should -Not -Invoke Write-InstallerSectionHeader -ModuleName InstallWinSCPForBizTalk.Workflow
    }
}
