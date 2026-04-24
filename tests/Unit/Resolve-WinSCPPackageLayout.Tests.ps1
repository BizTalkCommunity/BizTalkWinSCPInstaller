if (-not (Get-Module -ListAvailable Pester | Where-Object { $_.Version.Major -ge 5 })) {
    throw "Pester 5+ is required to run this test file. Install with: Install-Module Pester -Scope CurrentUser -RequiredVersion 5.0 -Force"
}

Describe "Resolve-WinSCPPackageLayout" {
    BeforeAll {
        # Import the core module
        $modulePath = Join-Path (Split-Path -Parent $PSCommandPath) "../../src/InstallWinSCPForBizTalk.Core.psm1"
        Import-Module $modulePath -Force
    }

    AfterAll {
        Remove-Module InstallWinSCPForBizTalk.Core -ErrorAction SilentlyContinue
    }

    Context "When package root exists with known folder structure" {
        BeforeEach {
            Mock -ModuleName InstallWinSCPForBizTalk.Core Test-Path {
                param($Path)
                if ($Path -eq "C:\temp\WinSCP.5.19.2") { return $true }
                # Mock: tools\WinSCP.exe and lib\netstandard2.0\WinSCPnet.dll exist
                $Path -match "\\tools\\WinSCP\.exe$" -or `
                $Path -match "\\lib\\netstandard2\.0\\WinSCPnet\.dll$"
            }
        }

        It "Should resolve both EXE and DLL from standard paths" {
            $result = Resolve-WinSCPPackageLayout `
                -PackageRoot "C:\temp\WinSCP.5.19.2" `
                -ExeFileName "WinSCP.exe" `
                -DllFileName "WinSCPnet.dll"

            $result.IsResolved | Should -Be $true
            $result.ExePath | Should -Match "tools\\WinSCP\.exe$"
            $result.DllPath | Should -Match "lib\\netstandard2\.0\\WinSCPnet\.dll$"
        }
    }

    Context "When package root does not exist" {
        BeforeEach {
            Mock -ModuleName InstallWinSCPForBizTalk.Core Test-Path { $false }
        }

        It "Should return unresolved with null paths" {
            $result = Resolve-WinSCPPackageLayout `
                -PackageRoot "C:\nonexistent\WinSCP.5.19.2" `
                -ExeFileName "WinSCP.exe" `
                -DllFileName "WinSCPnet.dll"

            $result.IsResolved | Should -Be $false
            $result.ExePath | Should -Be $null
            $result.DllPath | Should -Be $null
        }
    }

    Context "When standard paths fail but recursive search finds files" {
        BeforeEach {
            Mock -ModuleName InstallWinSCPForBizTalk.Core Test-Path {
                param($Path)
                if ($Path -eq "C:\temp\WinSCP.5.19.2") { return $true }
                return $false
            }
            Mock -ModuleName InstallWinSCPForBizTalk.Core Get-ChildItem {
                param($Path, $Filter, [switch]$Recurse, [switch]$File)
                
                if ($Filter -eq "WinSCP.exe" -and $Recurse) {
                    return [pscustomobject]@{
                        FullName = "C:\temp\WinSCP.5.19.2\somedeep\folder\WinSCP.exe"
                    }
                }
                if ($Filter -eq "WinSCPnet.dll" -and $Recurse) {
                    return [pscustomobject]@{
                        FullName = "C:\temp\WinSCP.5.19.2\somedeep\folder\WinSCPnet.dll"
                    }
                }
                return $null
            }
        }

        It "Should resolve via recursive fallback search" {
            $result = Resolve-WinSCPPackageLayout `
                -PackageRoot "C:\temp\WinSCP.5.19.2" `
                -ExeFileName "WinSCP.exe" `
                -DllFileName "WinSCPnet.dll"

            $result.IsResolved | Should -Be $true
            $result.ExePath | Should -Match "WinSCP\.exe$"
            $result.DllPath | Should -Match "WinSCPnet\.dll$"
        }
    }

    Context "When only EXE is found" {
        BeforeEach {
            Mock -ModuleName InstallWinSCPForBizTalk.Core Test-Path {
                param($Path)
                if ($Path -eq "C:\temp\WinSCP.5.19.2") { return $true }
                $Path -match "\\tools\\WinSCP\.exe$"
            }
            Mock -ModuleName InstallWinSCPForBizTalk.Core Get-ChildItem { $null }
        }

        It "Should return unresolved because DLL is missing" {
            $result = Resolve-WinSCPPackageLayout `
                -PackageRoot "C:\temp\WinSCP.5.19.2" `
                -ExeFileName "WinSCP.exe" `
                -DllFileName "WinSCPnet.dll"

            $result.IsResolved | Should -Be $false
            $result.ExePath | Should -Not -Be $null
            $result.DllPath | Should -Be $null
        }
    }
}
