if (-not (Get-Module -ListAvailable Pester | Where-Object { $_.Version.Major -ge 5 })) {
    throw "Pester 5+ is required to run this test file. Install with: Install-Module Pester -Scope CurrentUser -RequiredVersion 5.0 -Force"
}

Describe "Cyclomatic complexity gates" {
    BeforeAll {
        $script:repoRoot = Resolve-Path (Join-Path (Split-Path -Parent $PSCommandPath) "../..")
        $script:complexityTargets = @(
            @{
                Path = Join-Path $script:repoRoot "InstallWinSCPForBizTalk.ps1"
                MaxScriptBodyComplexity = 60
                MaxFunctionComplexity = 12
            },
            @{
                Path = Join-Path $script:repoRoot "src\InstallWinSCPForBizTalk.Core.psm1"
                MaxScriptBodyComplexity = 1
                MaxFunctionComplexity = 12
            }
        )

        function script:Get-DecisionCount {
            param(
                [Parameter(Mandatory = $true)]
                [System.Management.Automation.Language.Ast]$Ast,
                [switch]$TopLevelOnly
            )

            $decisionCount = 0
            $nodes = $Ast.FindAll({
                param($node)
                if ($node -eq $Ast) {
                    return $false
                }

                if ($TopLevelOnly) {
                    $ancestor = $node.Parent
                    while ($null -ne $ancestor) {
                        if ($ancestor -is [System.Management.Automation.Language.FunctionDefinitionAst]) {
                            return $false
                        }
                        if ($ancestor -eq $Ast) {
                            break
                        }
                        $ancestor = $ancestor.Parent
                    }
                }

                return (
                    $node -is [System.Management.Automation.Language.IfStatementAst] -or
                    $node -is [System.Management.Automation.Language.ForEachStatementAst] -or
                    $node -is [System.Management.Automation.Language.ForStatementAst] -or
                    $node -is [System.Management.Automation.Language.WhileStatementAst] -or
                    $node -is [System.Management.Automation.Language.DoWhileStatementAst] -or
                    $node -is [System.Management.Automation.Language.DoUntilStatementAst] -or
                    $node -is [System.Management.Automation.Language.CatchClauseAst] -or
                    $node -is [System.Management.Automation.Language.TrapStatementAst]
                )
            }, $true)

            foreach ($node in $nodes) {
                if ($node -is [System.Management.Automation.Language.IfStatementAst]) {
                    $decisionCount += $node.Clauses.Count
                }
                else {
                    $decisionCount += 1
                }
            }

            return $decisionCount
        }

        function script:Get-CyclomaticComplexityReport {
            param(
                [Parameter(Mandatory = $true)]
                [string]$Path
            )

            $tokens = $null
            $parseErrors = $null
            $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$parseErrors)
            if ($parseErrors.Count -gt 0) {
                throw "Cannot calculate complexity for '$Path' because it does not parse."
            }

            $functionReports = foreach ($functionAst in $ast.FindAll({
                param($node)
                $node -is [System.Management.Automation.Language.FunctionDefinitionAst]
            }, $true)) {
                $decisionCount = Get-DecisionCount -Ast $functionAst.Body
                [pscustomobject]@{
                    Name = $functionAst.Name
                    DecisionCount = $decisionCount
                    Complexity = $decisionCount + 1
                }
            }

            $scriptDecisionCount = Get-DecisionCount -Ast $ast -TopLevelOnly
            [pscustomobject]@{
                Path = $Path
                ScriptBodyComplexity = $scriptDecisionCount + 1
                Functions = $functionReports
            }
        }
    }

    It "keeps script-body complexity within threshold for each source file" {
        foreach ($target in $script:complexityTargets) {
            $report = Get-CyclomaticComplexityReport -Path $target.Path
            $report.ScriptBodyComplexity | Should -BeLessOrEqual $target.MaxScriptBodyComplexity -Because "script body complexity for $($target.Path) must stay within the pipeline gate"
        }
    }

    It "keeps each function within the per-function complexity threshold" {
        foreach ($target in $script:complexityTargets) {
            $report = Get-CyclomaticComplexityReport -Path $target.Path
            foreach ($functionReport in $report.Functions) {
                $functionReport.Complexity | Should -BeLessOrEqual $target.MaxFunctionComplexity -Because "$($functionReport.Name) in $($target.Path) must stay within the pipeline gate"
            }
        }
    }
}