<#
.SYNOPSIS
Unit tests for scripts/New-BizTalkProbeReport.ps1.
#>

if (-not (Get-Module -ListAvailable Pester | Where-Object { $_.Version.Major -ge 5 })) {
    throw "Pester 5+ is required to run this test file. Install with: Install-Module Pester -Scope CurrentUser -RequiredVersion 5.0 -Force"
}

Describe "New-BizTalkProbeReport script" {
    BeforeAll {
        $script:repoRoot = Resolve-Path (Join-Path (Split-Path -Parent $PSCommandPath) '../..')
        $script:probePath = Join-Path $script:repoRoot 'scripts/New-BizTalkProbeReport.ps1'
    }

    BeforeEach {
        $script:testRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('BizTalkWinSCPInstaller.ProbeTests.' + [guid]::NewGuid().ToString('N'))
        New-Item -Path $script:testRoot -ItemType Directory -Force | Out-Null
    }

    AfterEach {
        Remove-Item -Path $script:testRoot -Recurse -Force -ErrorAction SilentlyContinue
    }

    It "writes probe report JSON even when BizTalk is not detected" {
        $outputPath = Join-Path $script:testRoot 'probe.json'

        & $script:probePath -OutputPath $outputPath | Out-Null

        Test-Path $outputPath | Should -BeTrue
        $report = Get-Content -Path $outputPath -Raw | ConvertFrom-Json
        $report.schemaVersion | Should -Be '1.0'
        $report | Get-Member -Name selectedWinSCPVersion | Should -Not -BeNullOrEmpty
    }
}
