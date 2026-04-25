if (-not (Get-Module -ListAvailable Pester | Where-Object { $_.Version.Major -ge 5 })) {
    throw "Pester 5+ is required to run this test file. Install with: Install-Module Pester -Scope CurrentUser -RequiredVersion 5.0 -Force"
}

Describe "Cyclomatic complexity gates" {
    BeforeAll {
        $script:repoRoot = Resolve-Path (Join-Path (Split-Path -Parent $PSCommandPath) "../..")
        $script:complexityTargets = @(
            @{
                Path = Join-Path $script:repoRoot "InstallWinSCPForBizTalk.ps1"
                MaxScriptBodyComplexity = 46
                MaxFunctionComplexity = 12
                MaxScriptBodyCognitiveComplexity = 95
                MaxFunctionCognitiveComplexity = 35
            },
            @{
                Path = Join-Path $script:repoRoot "src\InstallWinSCPForBizTalk.Core.psm1"
                MaxScriptBodyComplexity = 1
                MaxFunctionComplexity = 12
                MaxScriptBodyCognitiveComplexity = 1
                MaxFunctionCognitiveComplexity = 35
            },
            @{
                Path = Join-Path $script:repoRoot "src\InstallWinSCPForBizTalk.Workflow.psm1"
                MaxScriptBodyComplexity = 1
                MaxFunctionComplexity = 10
                MaxScriptBodyCognitiveComplexity = 1
                MaxFunctionCognitiveComplexity = 20
            },
            @{
                Path = Join-Path $script:repoRoot "src\InstallWinSCPForBizTalk.Utils.psm1"
                MaxScriptBodyComplexity = 1
                MaxFunctionComplexity = 6
                MaxScriptBodyCognitiveComplexity = 1
                MaxFunctionCognitiveComplexity = 8
            }
        )

        function script:Get-DecisionNodes {
            param(
                [Parameter(Mandatory = $true)]
                [System.Management.Automation.Language.Ast]$Ast,
                [switch]$TopLevelOnly
            )

            return $Ast.FindAll({
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
        }

        function script:Get-DecisionCount {
            param(
                [Parameter(Mandatory = $true)]
                [System.Management.Automation.Language.Ast]$Ast,
                [switch]$TopLevelOnly
            )

            $decisionCount = 0
            $nodes = Get-DecisionNodes -Ast $Ast -TopLevelOnly:$TopLevelOnly

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

        function script:Get-NestingDepth {
            param(
                [Parameter(Mandatory = $true)]
                [System.Management.Automation.Language.Ast]$Node,
                [Parameter(Mandatory = $true)]
                [System.Management.Automation.Language.Ast]$BoundaryAst,
                [switch]$TopLevelOnly
            )

            $depth = 0
            $ancestor = $Node.Parent
            while ($null -ne $ancestor -and $ancestor -ne $BoundaryAst) {
                if ($TopLevelOnly -and $ancestor -is [System.Management.Automation.Language.FunctionDefinitionAst]) {
                    break
                }

                if (
                    $ancestor -is [System.Management.Automation.Language.IfStatementAst] -or
                    $ancestor -is [System.Management.Automation.Language.ForEachStatementAst] -or
                    $ancestor -is [System.Management.Automation.Language.ForStatementAst] -or
                    $ancestor -is [System.Management.Automation.Language.WhileStatementAst] -or
                    $ancestor -is [System.Management.Automation.Language.DoWhileStatementAst] -or
                    $ancestor -is [System.Management.Automation.Language.DoUntilStatementAst] -or
                    $ancestor -is [System.Management.Automation.Language.CatchClauseAst] -or
                    $ancestor -is [System.Management.Automation.Language.TrapStatementAst]
                ) {
                    $depth += 1
                }

                $ancestor = $ancestor.Parent
            }

            return $depth
        }

        function script:Get-CognitiveComplexity {
            param(
                [Parameter(Mandatory = $true)]
                [System.Management.Automation.Language.Ast]$Ast,
                [switch]$TopLevelOnly
            )

            $complexity = 0
            $nodes = Get-DecisionNodes -Ast $Ast -TopLevelOnly:$TopLevelOnly
            foreach ($node in $nodes) {
                $nestingDepth = Get-NestingDepth -Node $node -BoundaryAst $Ast -TopLevelOnly:$TopLevelOnly
                if ($node -is [System.Management.Automation.Language.IfStatementAst]) {
                    $complexity += ($node.Clauses.Count * (1 + $nestingDepth))
                }
                else {
                    $complexity += (1 + $nestingDepth)
                }
            }

            return $complexity
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
                $cognitiveComplexity = Get-CognitiveComplexity -Ast $functionAst.Body
                [pscustomobject]@{
                    Name = $functionAst.Name
                    DecisionCount = $decisionCount
                    Complexity = $decisionCount + 1
                    CognitiveComplexity = $cognitiveComplexity
                }
            }

            $scriptDecisionCount = Get-DecisionCount -Ast $ast -TopLevelOnly
            $scriptCognitiveComplexity = Get-CognitiveComplexity -Ast $ast -TopLevelOnly
            [pscustomobject]@{
                Path = $Path
                ScriptBodyComplexity = $scriptDecisionCount + 1
                ScriptBodyCognitiveComplexity = $scriptCognitiveComplexity
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

    It "keeps script-body cognitive complexity within threshold for each source file" {
        foreach ($target in $script:complexityTargets) {
            $report = Get-CyclomaticComplexityReport -Path $target.Path
            $report.ScriptBodyCognitiveComplexity | Should -BeLessOrEqual $target.MaxScriptBodyCognitiveComplexity -Because "script body cognitive complexity for $($target.Path) must stay within the pipeline gate"
        }
    }

    It "keeps each function within the per-function cognitive complexity threshold" {
        foreach ($target in $script:complexityTargets) {
            $report = Get-CyclomaticComplexityReport -Path $target.Path
            foreach ($functionReport in $report.Functions) {
                $functionReport.CognitiveComplexity | Should -BeLessOrEqual $target.MaxFunctionCognitiveComplexity -Because "$($functionReport.Name) in $($target.Path) must stay within the cognitive complexity gate"
            }
        }
    }
}