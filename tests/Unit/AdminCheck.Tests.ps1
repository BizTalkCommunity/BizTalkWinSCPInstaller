if (-not (Get-Module -ListAvailable Pester | Where-Object { $_.Version.Major -ge 5 })) {
    throw "Pester 5+ is required to run this test file. Install with: Install-Module Pester -Scope CurrentUser -RequiredVersion 5.0 -Force"
}

Describe "Test-IsAdministrator" {
    BeforeAll {
        $modulePath = Join-Path (Split-Path -Parent $PSCommandPath) "../../src/InstallWinSCPForBizTalk.Core.psm1"
        Import-Module $modulePath -Force
    }

    AfterAll {
        Remove-Module InstallWinSCPForBizTalk.Core -ErrorAction SilentlyContinue
    }

    It "returns true when principal reports Administrator role" {
        $identityGetter = { [pscustomobject]@{ Name = "UnitTestUser" } }
        $principalFactory = {
            param($Identity)
            $principal = New-Object psobject
            $principal | Add-Member -MemberType ScriptMethod -Name IsInRole -Value {
                param($Role)
                return $true
            }
            return $principal
        }

        $result = Test-IsAdministrator -GetCurrentIdentity $identityGetter -NewPrincipal $principalFactory
        $result | Should -BeTrue
    }

    It "returns false when principal does not report Administrator role" {
        $identityGetter = { [pscustomobject]@{ Name = "UnitTestUser" } }
        $principalFactory = {
            param($Identity)
            $principal = New-Object psobject
            $principal | Add-Member -MemberType ScriptMethod -Name IsInRole -Value {
                param($Role)
                return $false
            }
            return $principal
        }

        $result = Test-IsAdministrator -GetCurrentIdentity $identityGetter -NewPrincipal $principalFactory
        $result | Should -BeFalse
    }

    It "returns a boolean when using default identity/principal providers" {
        $result = Test-IsAdministrator
        ($result -is [bool]) | Should -BeTrue
    }
}
