<#
.SYNOPSIS
Unit tests for scripts/Build-ProductionPackage.ps1.
#>

if (-not (Get-Module -ListAvailable Pester | Where-Object { $_.Version.Major -ge 5 })) {
    throw "Pester 5+ is required to run this test file. Install with: Install-Module Pester -Scope CurrentUser -RequiredVersion 5.0 -Force"
}

Describe "Build-ProductionPackage script" {
    BeforeAll {
        $script:repoRoot = Resolve-Path (Join-Path (Split-Path -Parent $PSCommandPath) '../..')
        $script:builderPath = Join-Path $script:repoRoot 'scripts/Build-ProductionPackage.ps1'
    }

    BeforeEach {
        $script:testRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('BizTalkWinSCPInstaller.BundleTests.' + [guid]::NewGuid().ToString('N'))
        New-Item -Path $script:testRoot -ItemType Directory -Force | Out-Null
    }

    AfterEach {
        Remove-Item -Path $script:testRoot -Recurse -Force -ErrorAction SilentlyContinue
    }

    It "creates a minimal bundle with only runtime essentials" {
        $outputFolder = Join-Path $script:testRoot 'bundle'

        & $script:builderPath -OutputFolder $outputFolder -Clean | Out-Null

        Test-Path (Join-Path $outputFolder 'InstallWinSCPForBizTalk.ps1') | Should -BeTrue
        Test-Path (Join-Path $outputFolder 'src/InstallWinSCPForBizTalk.Core.psm1') | Should -BeTrue
        Test-Path (Join-Path $outputFolder 'src/InstallWinSCPForBizTalk.Utils.psm1') | Should -BeTrue
        Test-Path (Join-Path $outputFolder 'src/InstallWinSCPForBizTalk.Workflow.psm1') | Should -BeTrue
        Test-Path (Join-Path $outputFolder 'Run-Installer.ps1') | Should -BeTrue
        Test-Path (Join-Path $outputFolder 'PRODUCTION-README.md') | Should -BeTrue
        Test-Path (Join-Path $outputFolder 'package-manifest.json') | Should -BeTrue

        Test-Path (Join-Path $outputFolder 'tests') | Should -BeFalse
        Test-Path (Join-Path $outputFolder 'docs') | Should -BeFalse
    }

    It "includes offline nuget payload when supplied" {
        $outputFolder = Join-Path $script:testRoot 'bundle-with-payload'
        $payloadFolder = Join-Path $script:testRoot 'payload'
        $payloadPackage = Join-Path $payloadFolder 'WinSCP.5.15.4/tools'

        New-Item -Path $payloadPackage -ItemType Directory -Force | Out-Null
        Set-Content -Path (Join-Path $payloadFolder 'nuget.exe') -Value 'fake nuget' -Encoding ASCII
        Set-Content -Path (Join-Path $payloadPackage 'WinSCP.exe') -Value 'fake exe' -Encoding ASCII

        & $script:builderPath -OutputFolder $outputFolder -Clean -NuGetPayloadFolder $payloadFolder | Out-Null

        Test-Path (Join-Path $outputFolder 'nuget/nuget.exe') | Should -BeTrue
        Test-Path (Join-Path $outputFolder 'nuget/WinSCP.5.15.4/tools/WinSCP.exe') | Should -BeTrue

        $wrapperContent = Get-Content -Path (Join-Path $outputFolder 'Run-Installer.ps1') -Raw
        $expectedLiteral = [regex]::Escape("Join-Path `$PSScriptRoot 'nuget'")
        $wrapperContent | Should -Match $expectedLiteral
    }
}
