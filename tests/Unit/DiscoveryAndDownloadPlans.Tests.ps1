<#
.SYNOPSIS
Unit tests for discovery and package-download planning helpers.

.DESCRIPTION
Validates BizTalk install path resolution and download-plan decision helpers for
NuGet and WinSCP package acquisition behavior.
#>

if (-not (Get-Module -ListAvailable Pester | Where-Object { $_.Version.Major -ge 5 })) {
    throw "Pester 5+ is required to run this test file. Install with: Install-Module Pester -Scope CurrentUser -RequiredVersion 5.0 -Force"
}

Describe "Discovery and download planning helpers" {
    BeforeAll {
        $modulePath = Join-Path (Split-Path -Parent $PSCommandPath) "../../src/InstallWinSCPForBizTalk.Core.psm1"
        Import-Module $modulePath -Force
    }

    AfterAll {
        Remove-Module InstallWinSCPForBizTalk.Core -ErrorAction SilentlyContinue
    }

    Context "Resolve-BizTalkInstallFolder" {
        It "prefers environment path when both env and registry are set" {
            Mock -ModuleName InstallWinSCPForBizTalk.Core Test-Path { $true } -ParameterFilter { $Path -eq 'C:\EnvPath' }

            $result = Resolve-BizTalkInstallFolder -EnvironmentInstallPath 'C:\EnvPath' -RegistryInstallPath 'C:\RegistryPath'

            $result.InstallPath | Should -Be 'C:\EnvPath'
            $result.Source | Should -Be 'Environment'
            $result.IsFound | Should -BeTrue
            $result.Exists | Should -BeTrue
        }

        It "falls back to registry path when env is missing" {
            Mock -ModuleName InstallWinSCPForBizTalk.Core Test-Path { $true } -ParameterFilter { $Path -eq 'C:\RegistryPath' }

            $result = Resolve-BizTalkInstallFolder -EnvironmentInstallPath '' -RegistryInstallPath 'C:\RegistryPath'

            $result.InstallPath | Should -Be 'C:\RegistryPath'
            $result.Source | Should -Be 'Registry'
            $result.IsFound | Should -BeTrue
            $result.Exists | Should -BeTrue
        }

        It "returns None when both env and registry are missing" {
            $result = Resolve-BizTalkInstallFolder -EnvironmentInstallPath '' -RegistryInstallPath ''

            $result.InstallPath | Should -Be $null
            $result.Source | Should -Be 'None'
            $result.IsFound | Should -BeFalse
            $result.Exists | Should -BeFalse
        }
    }

    Context "Get-NuGetDownloadPlan" {
        It "downloads when nuget.exe does not exist" {
            $plan = Get-NuGetDownloadPlan -TargetNugetExeAlreadyExists $false -ForceInstall $false -SourceNugetExe 'https://dist.nuget.org/win-x86-commandline/latest/nuget.exe' -TargetNugetExe 'C:\nuget\nuget.exe'

            $plan.ShouldDownload | Should -BeTrue
            $plan.ShouldReuseExisting | Should -BeFalse
            $plan.ShouldProcessTarget | Should -Be 'https://dist.nuget.org/win-x86-commandline/latest/nuget.exe -OutFile C:\nuget\nuget.exe'
        }

        It "downloads when force install is true even if nuget.exe exists" {
            $plan = Get-NuGetDownloadPlan -TargetNugetExeAlreadyExists $true -ForceInstall $true -SourceNugetExe 'https://dist.nuget.org/win-x86-commandline/latest/nuget.exe' -TargetNugetExe 'C:\nuget\nuget.exe'

            $plan.ShouldDownload | Should -BeTrue
            $plan.ShouldReuseExisting | Should -BeFalse
        }

        It "reuses existing nuget.exe when present and not force install" {
            $plan = Get-NuGetDownloadPlan -TargetNugetExeAlreadyExists $true -ForceInstall $false -SourceNugetExe 'https://dist.nuget.org/win-x86-commandline/latest/nuget.exe' -TargetNugetExe 'C:\nuget\nuget.exe'

            $plan.ShouldDownload | Should -BeFalse
            $plan.ShouldReuseExisting | Should -BeTrue
        }
    }

    Context "Get-WinSCPPackageDownloadPlan" {
        It "downloads when one package artifact is missing" {
            $plan = Get-WinSCPPackageDownloadPlan -WinSCPEXEAlreadyExists $true -WinSCPDLLAlreadyExists $false -TargetNugetExe 'C:\nuget\nuget.exe' -WinSCPVersion '5.15.4' -NugetDownloadFolder 'C:\nuget'

            $plan.ShouldDownload | Should -BeTrue
            $plan.CommandPreview | Should -Be "'C:\nuget\nuget.exe' Install WinSCP -Version 5.15.4 -NonInteractive -OutputDirectory 'C:\nuget'"
        }

        It "skips download when both package artifacts already exist" {
            $plan = Get-WinSCPPackageDownloadPlan -WinSCPEXEAlreadyExists $true -WinSCPDLLAlreadyExists $true -TargetNugetExe 'C:\nuget\nuget.exe' -WinSCPVersion '5.15.4' -NugetDownloadFolder 'C:\nuget'

            $plan.ShouldDownload | Should -BeFalse
        }
    }
}
