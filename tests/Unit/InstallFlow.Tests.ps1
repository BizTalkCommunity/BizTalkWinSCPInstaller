<#
.SYNOPSIS
Unit tests for install-flow planning helper behavior.

.DESCRIPTION
Validates Get-InstallExecutionPlan outcomes for elevation, ForceInstall, WhatIf,
and already-installed decision paths.
#>

if (-not (Get-Module -ListAvailable Pester | Where-Object { $_.Version.Major -ge 5 })) {
    throw "Pester 5+ is required to run this test file. Install with: Install-Module Pester -Scope CurrentUser -RequiredVersion 5.0 -Force"
}

Describe "Get-InstallExecutionPlan" {
    BeforeAll {
        $modulePath = Join-Path (Split-Path -Parent $PSCommandPath) "../../src/InstallWinSCPForBizTalk.Core.psm1"
        Import-Module $modulePath -Force
    }

    AfterAll {
        Remove-Module InstallWinSCPForBizTalk.Core -ErrorAction SilentlyContinue
    }

    It "fails fast for non-admin ForceInstall when not WhatIf" {
        $plan = Get-InstallExecutionPlan -IsAdministrator $false -ForceInstall $true -WhatIf $false -AlreadyInstalledCorrect $false

        $plan.CanProceed | Should -BeFalse
        $plan.PrerequisiteFailure | Should -BeTrue
        $plan.RequiresElevationWarning | Should -BeFalse
    }

    It "warns but proceeds for non-admin ForceInstall in WhatIf mode" {
        $plan = Get-InstallExecutionPlan -IsAdministrator $false -ForceInstall $true -WhatIf $true -AlreadyInstalledCorrect $false

        $plan.CanProceed | Should -BeTrue
        $plan.PrerequisiteFailure | Should -BeFalse
        $plan.RequiresElevationWarning | Should -BeTrue
    }

    It "skips install when already installed and not ForceInstall" {
        $plan = Get-InstallExecutionPlan -IsAdministrator $true -ForceInstall $false -WhatIf $false -AlreadyInstalledCorrect $true

        $plan.CanProceed | Should -BeFalse
        $plan.ShouldSkipBecauseInstalled | Should -BeTrue
        $plan.ShouldReinstall | Should -BeFalse
    }

    It "reinstalls when already installed and ForceInstall is true" {
        $plan = Get-InstallExecutionPlan -IsAdministrator $true -ForceInstall $true -WhatIf $false -AlreadyInstalledCorrect $true

        $plan.ShouldReinstall | Should -BeTrue
        $plan.ShouldSkipBecauseInstalled | Should -BeFalse
        $plan.PrerequisiteFailure | Should -BeFalse
    }

    It "warns for non-admin normal install path" {
        $plan = Get-InstallExecutionPlan -IsAdministrator $false -ForceInstall $false -WhatIf $false -AlreadyInstalledCorrect $false

        $plan.CanProceed | Should -BeTrue
        $plan.RequiresElevationWarning | Should -BeTrue
        $plan.PrerequisiteFailure | Should -BeFalse
    }
}
