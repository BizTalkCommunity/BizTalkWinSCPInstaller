if (-not (Get-Module -ListAvailable Pester | Where-Object { $_.Version.Major -ge 5 })) {
    throw "Pester 5+ is required to run this test file. Install with: Install-Module Pester -Scope CurrentUser -RequiredVersion 5.0 -Force"
}

Describe "Get-FinalExecutionOutcome" {
    BeforeAll {
        $modulePath = Join-Path (Split-Path -Parent $PSCommandPath) "../../src/InstallWinSCPForBizTalk.Core.psm1"
        Import-Module $modulePath -Force
    }

    AfterAll {
        Remove-Module InstallWinSCPForBizTalk.Core -ErrorAction SilentlyContinue
    }

    It "classifies successful install as Success" {
        $result = Get-FinalExecutionOutcome -InstalledSuccessfully $true -WhatIf $false -PrerequisiteFailure $false -ContinueFlag $true

        $result.Outcome | Should -Be "Success"
        $result.IsError | Should -BeFalse
    }

    It "classifies WhatIf execution as DryRun when not installed" {
        $result = Get-FinalExecutionOutcome -InstalledSuccessfully $false -WhatIf $true -PrerequisiteFailure $false -ContinueFlag $true

        $result.Outcome | Should -Be "DryRun"
        $result.IsError | Should -BeFalse
    }

    It "classifies prerequisite stop as PrerequisiteFailure" {
        $result = Get-FinalExecutionOutcome -InstalledSuccessfully $false -WhatIf $false -PrerequisiteFailure $true -ContinueFlag $false

        $result.Outcome | Should -Be "PrerequisiteFailure"
        $result.IsError | Should -BeTrue
    }

    It "classifies non-prerequisite stop as InstallFailure" {
        $result = Get-FinalExecutionOutcome -InstalledSuccessfully $false -WhatIf $false -PrerequisiteFailure $false -ContinueFlag $false

        $result.Outcome | Should -Be "InstallFailure"
        $result.IsError | Should -BeTrue
    }

    It "classifies unresolved state as Unknown" {
        $result = Get-FinalExecutionOutcome -InstalledSuccessfully $false -WhatIf $false -PrerequisiteFailure $false -ContinueFlag $true

        $result.Outcome | Should -Be "Unknown"
        $result.IsError | Should -BeTrue
    }
}
