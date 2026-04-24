if (-not (Get-Module -ListAvailable Pester | Where-Object { $_.Version.Major -ge 5 })) {
    throw "Pester 5+ is required to run this test file. Install with: Install-Module Pester -Scope CurrentUser -RequiredVersion 5.0 -Force"
}

Describe "Test-WinSCPVersionString" {
    BeforeAll {
        $modulePath = Join-Path (Split-Path -Parent $PSCommandPath) "../../src/InstallWinSCPForBizTalk.Core.psm1"
        Import-Module $modulePath -Force
    }

    AfterAll {
        Remove-Module InstallWinSCPForBizTalk.Core -ErrorAction SilentlyContinue
    }

    It "returns MissingVersion for null" {
        $result = Test-WinSCPVersionString -WinSCPVersion $null

        $result.IsValid | Should -BeFalse
        $result.ErrorCode | Should -Be "MissingVersion"
        $result.ParsedVersion | Should -Be $null
    }

    It "returns MissingVersion for empty/whitespace" -ForEach @("", "   ") {
        $result = Test-WinSCPVersionString -WinSCPVersion $_

        $result.IsValid | Should -BeFalse
        $result.ErrorCode | Should -Be "MissingVersion"
        $result.ParsedVersion | Should -Be $null
    }

    It "returns InvalidFormat for non-version strings" -ForEach @("abc", "5.x.1", "v6.3.5") {
        $result = Test-WinSCPVersionString -WinSCPVersion $_

        $result.IsValid | Should -BeFalse
        $result.ErrorCode | Should -Be "InvalidFormat"
        $result.ParsedVersion | Should -Be $null
    }

    It "accepts valid dotted versions" -ForEach @("5.7.7", "6.3.5", "5.19.2.0") {
        $result = Test-WinSCPVersionString -WinSCPVersion $_

        $result.IsValid | Should -BeTrue
        $result.ErrorCode | Should -Be "None"
        $result.ParsedVersion | Should -Not -Be $null
    }
}
