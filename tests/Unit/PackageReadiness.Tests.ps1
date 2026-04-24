if (-not (Get-Module -ListAvailable Pester | Where-Object { $_.Version.Major -ge 5 })) {
    throw "Pester 5+ is required to run this test file. Install with: Install-Module Pester -Scope CurrentUser -RequiredVersion 5.0 -Force"
}

Describe "Get-PackageReadinessState" {
    BeforeAll {
        $modulePath = Join-Path (Split-Path -Parent $PSCommandPath) "../../src/InstallWinSCPForBizTalk.Core.psm1"
        Import-Module $modulePath -Force
    }

    AfterAll {
        Remove-Module InstallWinSCPForBizTalk.Core -ErrorAction SilentlyContinue
    }

    It "returns MissingNuGet when nuget.exe is missing" {
        $result = Get-PackageReadinessState -NuGetExeExists $false -WinSCPExeExists $true -WinSCPDllExists $true

        $result.IsReady | Should -BeFalse
        $result.State | Should -Be "MissingNuGet"
    }

    It "returns MissingWinSCPPackage when both WinSCP files are missing" {
        $result = Get-PackageReadinessState -NuGetExeExists $true -WinSCPExeExists $false -WinSCPDllExists $false

        $result.IsReady | Should -BeFalse
        $result.State | Should -Be "MissingWinSCPPackage"
    }

    It "returns IncompleteWinSCPPackage when one WinSCP file is missing" -ForEach @(
        @{ Exe = $true; Dll = $false },
        @{ Exe = $false; Dll = $true }
    ) {
        param($Exe, $Dll)

        $result = Get-PackageReadinessState -NuGetExeExists $true -WinSCPExeExists $Exe -WinSCPDllExists $Dll

        $result.IsReady | Should -BeFalse
        $result.State | Should -Be "IncompleteWinSCPPackage"
    }

    It "returns Ready when all package artifacts exist" {
        $result = Get-PackageReadinessState -NuGetExeExists $true -WinSCPExeExists $true -WinSCPDllExists $true

        $result.IsReady | Should -BeTrue
        $result.State | Should -Be "Ready"
    }
}
