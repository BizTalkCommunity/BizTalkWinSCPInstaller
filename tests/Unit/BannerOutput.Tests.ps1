<#
.SYNOPSIS
Focused banner/output tests for InstallWinSCPForBizTalk.ps1.

.DESCRIPTION
Bootstraps minimal Pester coverage for banner constants and final WhatIf output
behavior in the current monolithic script layout.
#>

if (-not (Get-Module -ListAvailable Pester | Where-Object { $_.Version.Major -ge 5 })) {
    throw "Pester 5+ is required to run this test file. Install with: Install-Module Pester -Scope CurrentUser -RequiredVersion 5.0 -Force"
}

Describe "Installer banner and output behavior" {
    BeforeAll {
        $script:repoRoot = Resolve-Path (Join-Path (Split-Path -Parent $PSCommandPath) "../..")
        $script:installerPath = Join-Path $script:repoRoot "InstallWinSCPForBizTalk.ps1"
        $script:hashBanner = "##############################################################################"
        $script:bangBanner = "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
        $script:upBanner = "^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^"
    }

    BeforeEach {
        $script:testRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("BizTalkWinSCPInstaller.BannerTests." + [guid]::NewGuid().ToString("N"))
        $script:bizTalkFolder = Join-Path $script:testRoot "BizTalk"
        $script:nugetFolder = Join-Path $script:testRoot "nuget"

        New-Item -Path $script:bizTalkFolder -ItemType Directory -Force | Out-Null
        New-Item -Path $script:nugetFolder -ItemType Directory -Force | Out-Null

        $script:previousTemp = $env:TEMP
        $script:previousBtsInstallPath = $env:BTSINSTALLPATH
        $env:TEMP = $script:testRoot
        $env:BTSINSTALLPATH = $script:bizTalkFolder

        Mock Get-ItemPropertyValue {
            param([string]$Path, [string]$Name)
            switch ($Name) {
                "ProductCodeCurrent" { return "{205F5836-7512-4A06-9E74-ADC8AFA0EEC5}" }
                "ProductName" { return "Microsoft BizTalk Server 2020" }
                "ProductVersion" { return "3.13.717.0" }
                "InstallPath" { return $script:bizTalkFolder }
                default { return $null }
            }
        } -ParameterFilter { $Path -eq "HKLM:\SOFTWARE\Microsoft\BizTalk Server\3.0" }

        Mock Get-ChildItem { @() } -ParameterFilter { $Path -like "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall*" }
    }

    AfterEach {
        if ($null -ne $script:previousTemp) {
            $env:TEMP = $script:previousTemp
        }
        else {
            Remove-Item Env:TEMP -ErrorAction SilentlyContinue
        }

        if ($null -ne $script:previousBtsInstallPath) {
            $env:BTSINSTALLPATH = $script:previousBtsInstallPath
        }
        else {
            Remove-Item Env:BTSINSTALLPATH -ErrorAction SilentlyContinue
        }

        Remove-Item -Path $script:testRoot -Recurse -Force -ErrorAction SilentlyContinue
    }

    It "emits hash and bang banners in WhatIf mode" {
        $output = (& $script:installerPath -nugetDownloadFolder $script:nugetFolder -WhatIf -Confirm:$false 6>&1) | Out-String

        $output | Should -Match ([regex]::Escape($script:hashBanner))
        $output | Should -Match ([regex]::Escape($script:bangBanner))
    }

    It "does not emit legacy up-banner footer" {
        $output = (& $script:installerPath -nugetDownloadFolder $script:nugetFolder -WhatIf -Confirm:$false 6>&1) | Out-String

        $output | Should -Not -Match ([regex]::Escape($script:upBanner))
    }
}
