if (-not (Get-Module -ListAvailable Pester | Where-Object { $_.Version.Major -ge 5 })) {
    throw "Pester 5+ is required to run this test file. Install with: Install-Module Pester -Scope CurrentUser -RequiredVersion 5.0 -Force"
}

Describe "InstallWinSCPForBizTalk script execution" {
    BeforeAll {
        $script:repoRoot = Resolve-Path (Join-Path (Split-Path -Parent $PSCommandPath) "../..")
        $script:installerPath = Join-Path $script:repoRoot "InstallWinSCPForBizTalk.ps1"
    }

    BeforeEach {
        $script:testRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("BizTalkWinSCPInstaller.Tests." + [guid]::NewGuid().ToString("N"))
        $script:bizTalkFolder = Join-Path $script:testRoot "BizTalk"
        $script:nugetFolder = Join-Path $script:testRoot "nuget"
        New-Item -Path $script:bizTalkFolder -ItemType Directory -Force | Out-Null
        New-Item -Path $script:nugetFolder -ItemType Directory -Force | Out-Null

        $script:previousBtsInstallPath = $env:BTSINSTALLPATH
        $env:BTSINSTALLPATH = $script:bizTalkFolder

        # Stub registry reads used by the installer script.
        Mock Get-ItemPropertyValue {
            switch ($Name) {
                "ProductCodeCurrent" { return "{205F5836-7512-4A06-9E74-ADC8AFA0EEC5}" }
                "ProductName" { return "Microsoft BizTalk Server 2020" }
                "ProductVersion" { return "3.13.717.0" }
                "InstallPath" { return $script:bizTalkFolder }
                default { return $null }
            }
        } -ParameterFilter { $Path -eq "HKLM:\SOFTWARE\Microsoft\BizTalk Server\3.0" }
    }

    AfterEach {
        if ($null -ne $script:previousBtsInstallPath) {
            $env:BTSINSTALLPATH = $script:previousBtsInstallPath
        }
        else {
            Remove-Item Env:BTSINSTALLPATH -ErrorAction SilentlyContinue
        }

        Remove-Item -Path $script:testRoot -Recurse -Force -ErrorAction SilentlyContinue
    }

    It "installs from pre-staged package files without invoking NuGet download" {
        $packageRoot = Join-Path $script:nugetFolder "WinSCP.5.15.4"
        $toolsPath = Join-Path $packageRoot "tools"
        $dllPath = Join-Path $packageRoot "lib/netstandard2.0"
        New-Item -Path $toolsPath -ItemType Directory -Force | Out-Null
        New-Item -Path $dllPath -ItemType Directory -Force | Out-Null
        Set-Content -Path (Join-Path $toolsPath "WinSCP.exe") -Value "fake exe" -Encoding ASCII
        Set-Content -Path (Join-Path $dllPath "WinSCPnet.dll") -Value "fake dll" -Encoding ASCII
        Set-Content -Path (Join-Path $script:nugetFolder "nuget.exe") -Value "fake nuget" -Encoding ASCII

        Mock Get-ItemProperty { @() }

        & $script:installerPath -nugetDownloadFolder $script:nugetFolder -Confirm:$false

        Test-Path (Join-Path $script:bizTalkFolder "WinSCP.exe") | Should -BeTrue
        Test-Path (Join-Path $script:bizTalkFolder "WinSCPnet.dll") | Should -BeTrue
    }

    It "chooses CU6 package version when CU6 uninstall entry is detected" {
        $packageRoot = Join-Path $script:nugetFolder "WinSCP.6.3.5"
        $toolsPath = Join-Path $packageRoot "tools"
        $dllPath = Join-Path $packageRoot "lib/netstandard2.0"
        New-Item -Path $toolsPath -ItemType Directory -Force | Out-Null
        New-Item -Path $dllPath -ItemType Directory -Force | Out-Null
        Set-Content -Path (Join-Path $toolsPath "WinSCP.exe") -Value "fake exe" -Encoding ASCII
        Set-Content -Path (Join-Path $dllPath "WinSCPnet.dll") -Value "fake dll" -Encoding ASCII
        Set-Content -Path (Join-Path $script:nugetFolder "nuget.exe") -Value "fake nuget" -Encoding ASCII

        Mock Get-ItemProperty {
            @(
                [pscustomobject]@{
                    DisplayName = "Microsoft BizTalk Server 2020 Cumulative Update 6 KB5048971"
                    InstallDate = "20241121"
                }
            )
        }

        & $script:installerPath -nugetDownloadFolder $script:nugetFolder -Confirm:$false

        Test-Path (Join-Path $script:bizTalkFolder "WinSCP.exe") | Should -BeTrue
        Test-Path (Join-Path $script:bizTalkFolder "WinSCPnet.dll") | Should -BeTrue
    }

    It "uses 2016 CU mapping when a BizTalk 2016 CU is detected" {
        Mock Get-ItemPropertyValue {
            switch ($Name) {
                "ProductCodeCurrent" { return "{B084F3A7-3E8F-4E7B-B673-EED1715D28ED}" }
                "ProductName" { return "Microsoft BizTalk Server 2016" }
                "ProductVersion" { return "3.12.896.2" }
                "InstallPath" { return $script:bizTalkFolder }
                default { return $null }
            }
        } -ParameterFilter { $Path -eq "HKLM:\SOFTWARE\Microsoft\BizTalk Server\3.0" }

        Mock Get-ItemProperty {
            @(
                [pscustomobject]@{ DisplayName = "Microsoft BizTalk Server 2016 Cumulative Update 9 KB5005479" }
            )
        }

        $packageRoot = Join-Path $script:nugetFolder "WinSCP.5.19.2"
        $toolsPath = Join-Path $packageRoot "tools"
        $dllPath = Join-Path $packageRoot "lib/netstandard2.0"
        New-Item -Path $toolsPath -ItemType Directory -Force | Out-Null
        New-Item -Path $dllPath -ItemType Directory -Force | Out-Null
        Set-Content -Path (Join-Path $toolsPath "WinSCP.exe") -Value "fake exe" -Encoding ASCII
        Set-Content -Path (Join-Path $dllPath "WinSCPnet.dll") -Value "fake dll" -Encoding ASCII
        Set-Content -Path (Join-Path $script:nugetFolder "nuget.exe") -Value "fake nuget" -Encoding ASCII

        & $script:installerPath -nugetDownloadFolder $script:nugetFolder -Confirm:$false

        Test-Path (Join-Path $script:bizTalkFolder "WinSCP.exe") | Should -BeTrue
        Test-Path (Join-Path $script:bizTalkFolder "WinSCPnet.dll") | Should -BeTrue
    }

    It "creates nuget download folder when the provided folder does not exist" {
        $newNugetFolder = Join-Path $script:testRoot "nuget-missing"
        $packageRoot = Join-Path $newNugetFolder "WinSCP.5.15.4"
        $toolsPath = Join-Path $packageRoot "tools"
        $dllPath = Join-Path $packageRoot "lib/netstandard2.0"

        Mock Get-ItemProperty { @() }

        # Invoke-WebRequest is only used to hydrate nuget.exe when not present.
        Mock Invoke-WebRequest {
            param($Uri, $OutFile)
            New-Item -Path (Split-Path -Parent $OutFile) -ItemType Directory -Force | Out-Null
            Set-Content -Path $OutFile -Value "fake nuget" -Encoding ASCII
        }

        New-Item -Path $toolsPath -ItemType Directory -Force | Out-Null
        New-Item -Path $dllPath -ItemType Directory -Force | Out-Null
        Set-Content -Path (Join-Path $toolsPath "WinSCP.exe") -Value "fake exe" -Encoding ASCII
        Set-Content -Path (Join-Path $dllPath "WinSCPnet.dll") -Value "fake dll" -Encoding ASCII

        & $script:installerPath -nugetDownloadFolder $newNugetFolder -Confirm:$false

        Test-Path $newNugetFolder | Should -BeTrue
        Test-Path (Join-Path $script:bizTalkFolder "WinSCP.exe") | Should -BeTrue
        Assert-MockCalled Invoke-WebRequest -Times 1
    }

    It "completes successfully in WhatIf mode without creating target binaries" {
        Set-Content -Path (Join-Path $script:nugetFolder "nuget.exe") -Value "fake nuget" -Encoding ASCII
        Mock Get-ItemProperty { @() }

        & $script:installerPath -nugetDownloadFolder $script:nugetFolder -Confirm:$false -WhatIf

        Test-Path (Join-Path $script:bizTalkFolder "WinSCP.exe") | Should -BeFalse
        Test-Path (Join-Path $script:bizTalkFolder "WinSCPnet.dll") | Should -BeFalse
    }

    It "supports ForceInstall in non-admin WhatIf mode" {
        $packageRoot = Join-Path $script:nugetFolder "WinSCP.5.15.4"
        $toolsPath = Join-Path $packageRoot "tools"
        $dllPath = Join-Path $packageRoot "lib/netstandard2.0"
        New-Item -Path $toolsPath -ItemType Directory -Force | Out-Null
        New-Item -Path $dllPath -ItemType Directory -Force | Out-Null
        Set-Content -Path (Join-Path $toolsPath "WinSCP.exe") -Value "fake exe" -Encoding ASCII
        Set-Content -Path (Join-Path $dllPath "WinSCPnet.dll") -Value "fake dll" -Encoding ASCII
        Set-Content -Path (Join-Path $script:nugetFolder "nuget.exe") -Value "fake nuget" -Encoding ASCII

        Mock Get-ItemProperty { @() }

        & $script:installerPath -nugetDownloadFolder $script:nugetFolder -ForceInstall -WhatIf -Confirm:$false

        Test-Path (Join-Path $script:bizTalkFolder "WinSCP.exe") | Should -BeFalse
        Test-Path (Join-Path $script:bizTalkFolder "WinSCPnet.dll") | Should -BeFalse
    }

    It "stops early for unsupported BizTalk product code" {
        Mock Get-ItemPropertyValue {
            switch ($Name) {
                "ProductCodeCurrent" { return "{00000000-0000-0000-0000-000000000000}" }
                "ProductName" { return "Microsoft BizTalk Server Unknown" }
                "ProductVersion" { return "0.0.0.0" }
                "InstallPath" { return $script:bizTalkFolder }
                default { return $null }
            }
        } -ParameterFilter { $Path -eq "HKLM:\SOFTWARE\Microsoft\BizTalk Server\3.0" }

        Mock Get-ItemProperty { @() }

        & $script:installerPath -nugetDownloadFolder $script:nugetFolder -Confirm:$false

        Test-Path (Join-Path $script:bizTalkFolder "WinSCP.exe") | Should -BeFalse
        Test-Path (Join-Path $script:bizTalkFolder "WinSCPnet.dll") | Should -BeFalse
    }

}
