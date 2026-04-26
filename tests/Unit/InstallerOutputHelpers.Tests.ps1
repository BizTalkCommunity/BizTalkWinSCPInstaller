<#
.SYNOPSIS
Unit tests for installer output and snapshot helper functions.

.DESCRIPTION
Validates banner/delimited output helpers, final-outcome messaging, state snapshot
wrappers, and semantic output helper behavior in the Utils module.
#>

if (-not (Get-Module -ListAvailable Pester | Where-Object { $_.Version.Major -ge 5 })) {
    throw "Pester 5+ is required to run this test file. Install with: Install-Module Pester -Scope CurrentUser -RequiredVersion 5.0 -Force"
}

Describe "Installer output helpers" {
    BeforeAll {
        $modulePath = Join-Path (Split-Path -Parent $PSCommandPath) "../../src/InstallWinSCPForBizTalk.Utils.psm1"
        Import-Module $modulePath -Force
    }

    AfterAll {
        Disable-InstallerLogging
        Remove-Module InstallWinSCPForBizTalk.Utils -ErrorAction SilentlyContinue
    }

    Context "Initialize-InstallerLogging" {
        BeforeEach {
            $script:logRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("InstallerLogging.Tests." + [guid]::NewGuid().ToString("N"))
        }

        AfterEach {
            Disable-InstallerLogging
            Remove-Item -Path $script:logRoot -Recurse -Force -ErrorAction SilentlyContinue
        }

        It "creates a timestamped support log file in the requested folder" {
            Initialize-InstallerLogging -LogFolder $script:logRoot -LogLevel 'Info' | Out-Null

            $logFiles = Get-ChildItem -Path $script:logRoot -Filter 'BizTalkWinSCPInstaller-*.log'
            $logFiles.Count | Should -Be 1
            $logFiles[0].Name | Should -Match '^BizTalkWinSCPInstaller-\d{4}-\d{2}-\d{2}-\d{6}(?:-\d{2})?\.log$'
            $content = Get-Content -Path $logFiles[0].FullName -Raw
            $content | Should -Match 'BizTalk WinSCP Installer log'
            $content | Should -Match 'Support log file:'
        }

        It "records snapshot details only when the log level includes verbose detail" {
            Initialize-InstallerLogging -LogFolder $script:logRoot -LogLevel 'Info' | Out-Null
            Write-InstallerStateSnapshot -Title 'Test snapshot' -State ([ordered]@{ key = 'value' })
            $infoLog = Get-ChildItem -Path $script:logRoot -Filter '*.log' | Get-Content -Raw

            Disable-InstallerLogging
            Remove-Item -Path $script:logRoot -Recurse -Force -ErrorAction SilentlyContinue
            New-Item -Path $script:logRoot -ItemType Directory -Force | Out-Null

            Initialize-InstallerLogging -LogFolder $script:logRoot -LogLevel 'Verbose' | Out-Null
            Write-InstallerStateSnapshot -Title 'Test snapshot' -State ([ordered]@{ key = 'value' })
            $verboseLog = Get-ChildItem -Path $script:logRoot -Filter '*.log' | Get-Content -Raw

            $infoLog    | Should -Not -Match '\[VERBOSE\]'
            $verboseLog | Should -Match '\[VERBOSE\] Test snapshot'
        }

        It "writes event records when event log sink is enabled" {
            Mock -ModuleName InstallWinSCPForBizTalk.Utils Test-InstallerEventSourceExists { $true }
            Mock -ModuleName InstallWinSCPForBizTalk.Utils Write-InstallerEventLogRecord {}

            Initialize-InstallerLogging -LogFolder $script:logRoot -LogLevel 'Info' -EnableEventLog -EventLogName 'Application' -EventSource 'BizTalkWinSCPInstaller.Tests' | Out-Null
            Write-InstallerLogEntry -Level 'Info' -Message 'Event sink test line'

            Should -Invoke Write-InstallerEventLogRecord -ModuleName InstallWinSCPForBizTalk.Utils -Times 3 -ParameterFilter {
                $EventLogName -eq 'Application' -and $EventSource -eq 'BizTalkWinSCPInstaller.Tests'
            }
        }

        It "falls back to warning line in file log when event log write fails" {
            Mock -ModuleName InstallWinSCPForBizTalk.Utils Test-InstallerEventSourceExists { $true }
            Mock -ModuleName InstallWinSCPForBizTalk.Utils Write-InstallerEventLogRecord { throw 'Event sink unavailable' }

            Initialize-InstallerLogging -LogFolder $script:logRoot -LogLevel 'Info' -EnableEventLog | Out-Null
            Write-InstallerLogEntry -Level 'Info' -Message 'Event sink fallback test'

            $logFile = Get-ChildItem -Path $script:logRoot -Filter 'BizTalkWinSCPInstaller-*.log' | Select-Object -First 1
            $content = Get-Content -Path $logFile.FullName -Raw
            $content | Should -Match '\[WARN\] Event log write failed:'
        }

        It "creates a numbered log file name when a timestamp collision occurs" {
            Mock -ModuleName InstallWinSCPForBizTalk.Utils Get-Date { [datetime]'2026-04-25T23:30:00' }

            $existingLog = Join-Path $script:logRoot 'BizTalkWinSCPInstaller-2026-04-25-233000.log'
            New-Item -Path $script:logRoot -ItemType Directory -Force | Out-Null
            Set-Content -Path $existingLog -Value 'existing log' -Encoding UTF8

            $session = Initialize-InstallerLogging -LogFolder $script:logRoot -LogLevel 'Info'
            Split-Path -Leaf $session.LogFile | Should -Be 'BizTalkWinSCPInstaller-2026-04-25-233000-02.log'
        }
    }

    Context "Level mapping helpers" {
        It "maps installer levels to numeric ranks" {
            (Get-InstallerLogLevelRank -Level 'Error')   | Should -Be 0
            (Get-InstallerLogLevelRank -Level 'Info')    | Should -Be 1
            (Get-InstallerLogLevelRank -Level 'Verbose') | Should -Be 2
            (Get-InstallerLogLevelRank -Level 'Debug')   | Should -Be 3
        }

        It "maps event log entry types with error as Error and others as Information" {
            (Convert-InstallerEventEntryType -Level 'Error')   | Should -Be 'Error'
            (Convert-InstallerEventEntryType -Level 'Info')    | Should -Be 'Information'
            (Convert-InstallerEventEntryType -Level 'Verbose') | Should -Be 'Information'
            (Convert-InstallerEventEntryType -Level 'Debug')   | Should -Be 'Information'
        }
    }

    Context "Event log wrappers" {
        It "forwards event-source registration to New-EventLog" {
            Mock -ModuleName InstallWinSCPForBizTalk.Utils New-EventLog {}

            New-InstallerEventSource -EventLogName 'Application' -EventSource 'BizTalkWinSCPInstaller.Tests'

            Should -Invoke New-EventLog -ModuleName InstallWinSCPForBizTalk.Utils -Times 1 -ParameterFilter {
                $LogName -eq 'Application' -and $Source -eq 'BizTalkWinSCPInstaller.Tests'
            }
        }

        It "forwards event entry writes to Write-EventLog" {
            Mock -ModuleName InstallWinSCPForBizTalk.Utils Write-EventLog {}

            Write-InstallerEventLogRecord -EventLogName 'Application' -EventSource 'BizTalkWinSCPInstaller.Tests' -EntryType 'Information' -Message 'hello'

            Should -Invoke Write-EventLog -ModuleName InstallWinSCPForBizTalk.Utils -Times 1 -ParameterFilter {
                $LogName -eq 'Application' -and
                $Source -eq 'BizTalkWinSCPInstaller.Tests' -and
                $EntryType -eq 'Information' -and
                $EventId -eq 1000 -and
                $Category -eq 0 -and
                $Message -eq 'hello'
            }
        }

        It "creates event source when missing before writing event record" {
            $script:eventSourceCheckCount = 0
            Mock -ModuleName InstallWinSCPForBizTalk.Utils Test-InstallerEventSourceExists {
                $script:eventSourceCheckCount++
                return $script:eventSourceCheckCount -gt 1
            }
            Mock -ModuleName InstallWinSCPForBizTalk.Utils New-InstallerEventSource {}
            Mock -ModuleName InstallWinSCPForBizTalk.Utils Write-InstallerEventLogRecord {}

            $eventRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('InstallerLogging.EventSource.' + [guid]::NewGuid().ToString('N'))
            try {
                Initialize-InstallerLogging -LogFolder $eventRoot -EnableEventLog -EventSource 'BizTalkWinSCPInstaller.Tests' | Out-Null
                Write-InstallerEventLogEntry -Level 'Info' -Message 'event write test'
            }
            finally {
                Disable-InstallerLogging
                Remove-Item -Path $eventRoot -Recurse -Force -ErrorAction SilentlyContinue
            }

            Should -Invoke New-InstallerEventSource -ModuleName InstallWinSCPForBizTalk.Utils -Times 1
            Should -Invoke Write-InstallerEventLogRecord -ModuleName InstallWinSCPForBizTalk.Utils -Times 3
        }
    }

    Context "Get-InstallerBannerLine" {
        It "returns the hash banner for Hash" {
            (Get-InstallerBannerLine -Name 'Hash') | Should -Be '##############################################################################'
        }

        It "returns the bang banner for Bang" {
            (Get-InstallerBannerLine -Name 'Bang') | Should -Be '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!'
        }

        It "returns the up banner for Up" {
            (Get-InstallerBannerLine -Name 'Up') | Should -Be '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^'
        }
    }

    Context "Write-InstallerSectionHeader" {
        BeforeEach {
            $script:successMessages = @()
            Mock -ModuleName InstallWinSCPForBizTalk.Utils Write-InstallerSuccess {
                param([string] $SuccessMessage)
                $script:successMessages += $SuccessMessage
            }
        }

        It "writes section header as delimiter, title, delimiter" {
            Write-InstallerSectionHeader -Title 'Checking BizTalk'

            $script:successMessages.Count | Should -Be 3
            $script:successMessages[0] | Should -Be '##############################################################################'
            $script:successMessages[1] | Should -Be 'Checking BizTalk'
            $script:successMessages[2] | Should -Be '##############################################################################'
        }

        It "adds leading newline to first delimiter when requested" {
            Write-InstallerSectionHeader -Title 'Checking BizTalk' -LeadingNewLine

            $script:successMessages.Count | Should -Be 3
            $script:successMessages[0] | Should -Be "`n##############################################################################"
            $script:successMessages[1] | Should -Be 'Checking BizTalk'
            $script:successMessages[2] | Should -Be '##############################################################################'
        }
    }

    Context "Write-InstallerBangError" {
        It "writes a bang-delimited error block" {
            Mock -ModuleName InstallWinSCPForBizTalk.Utils Write-InstallerDelimitedMessage {}

            Write-InstallerBangError -LeadingNewLine -MessageLines @('line 1', 'line 2')

            Should -Invoke Write-InstallerDelimitedMessage -ModuleName InstallWinSCPForBizTalk.Utils -Times 1 -ParameterFilter {
                $Delimiter -eq 'Bang' -and
                $Level -eq 'Error' -and
                $LeadingNewLine -eq $true -and
                $MessageLines.Count -eq 2 -and
                $MessageLines[0] -eq 'line 1' -and
                $MessageLines[1] -eq 'line 2'
            }
        }
    }

    Context "Write-InstallerFinalOutcome" {
        BeforeEach {
            $script:successMessages = @()
            $script:errorMessages = @()

            Mock -ModuleName InstallWinSCPForBizTalk.Utils Write-InstallerSuccess {
                param([string] $SuccessMessage)
                $script:successMessages += $SuccessMessage
            }

            Mock -ModuleName InstallWinSCPForBizTalk.Utils Write-InstallerError {
                param([string] $ErrorMessage)
                $script:errorMessages += $ErrorMessage
            }
        }

        It "writes success outcome with paired bang delimiters" {
            Write-InstallerFinalOutcome -Outcome 'Success' -WinSCPVersion '5.15.4'

            $script:errorMessages.Count | Should -Be 0
            $script:successMessages.Count | Should -Be 4
            $script:successMessages[0] | Should -Be "`n!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
            $script:successMessages[1] | Should -Be 'WinSCP 5.15.4 is installed.'
            $script:successMessages[3] | Should -Be '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!'
        }

        It "writes dry-run outcome as success delimited message" {
            Write-InstallerFinalOutcome -Outcome 'DryRun' -WinSCPVersion '5.15.4'

            $script:errorMessages.Count | Should -Be 0
            $script:successMessages.Count | Should -Be 5
            $script:successMessages[0] | Should -Be "`n!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
            $script:successMessages[4] | Should -Be '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!'
        }

        It "writes prerequisite-failure outcome as error delimited message" {
            Write-InstallerFinalOutcome -Outcome 'PrerequisiteFailure' -WinSCPVersion '5.15.4'

            $script:successMessages.Count | Should -Be 0
            $script:errorMessages.Count | Should -Be 5
            $script:errorMessages[0] | Should -Be "`n!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
            $script:errorMessages[4] | Should -Be '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!'
        }

        It "writes unknown outcome as generic error delimited message" {
            Write-InstallerFinalOutcome -Outcome 'SomethingElse'

            $script:successMessages.Count | Should -Be 0
            $script:errorMessages.Count | Should -Be 5
            ($script:errorMessages -join "`n") | Should -Match 'Something went wrong during installation'
        }
    }

    Context "Write-InstallerStateSnapshot" {
        BeforeEach {
            $script:verboseMessages = @()
            $script:debugMessages = @()

            Mock -ModuleName InstallWinSCPForBizTalk.Utils Write-Verbose {
                param([string] $Message)
                $script:verboseMessages += $Message
            }

            Mock -ModuleName InstallWinSCPForBizTalk.Utils Write-Debug {
                param([string] $Message)
                $script:debugMessages += $Message
            }
        }

        It "writes title and each key/value to verbose and debug streams" {
            Write-InstallerStateSnapshot -Title 'Snapshot title' -State ([ordered]@{
                firstVar = 'alpha'
                secondVar = 42
            })

            $script:verboseMessages.Count | Should -Be 3
            $script:debugMessages.Count | Should -Be 3
            $script:verboseMessages[0] | Should -Be 'Snapshot title'
            $script:verboseMessages[1] | Should -Be '$firstVar = alpha'
            $script:verboseMessages[2] | Should -Be '$secondVar = 42'
            $script:debugMessages[0] | Should -Be 'Snapshot title'
        }
    }

    Context "Snapshot wrapper helpers" {
        It "maps BizTalk search fields into a snapshot payload" {
            Mock -ModuleName InstallWinSCPForBizTalk.Utils Write-InstallerStateSnapshot {}

            Write-BizTalkSearchSnapshot -InstallFolder 'C:\BizTalk\' -InstallFolderExists $true -ProductCodeCurrent 'CODE' -ProductName 'BizTalk' -ProductVersion '3.13'

            Should -Invoke Write-InstallerStateSnapshot -ModuleName InstallWinSCPForBizTalk.Utils -Times 1 -ParameterFilter {
                $Title -eq 'The result of the search for the BizTalk Server:' -and
                $State.bizTalkInstallFolder -eq 'C:\BizTalk\' -and
                $State.bizTalkInstallFolderExists -eq $true -and
                $State.bizTalkProductCodeCurrent -eq 'CODE' -and
                $State.bizTalkProductName -eq 'BizTalk' -and
                $State.bizTalkProductVersion -eq '3.13'
            }
        }

        It "maps WinSCP copy fields into a snapshot payload" {
            Mock -ModuleName InstallWinSCPForBizTalk.Utils Write-InstallerStateSnapshot {}

            Write-WinSCPCopySnapshot -BizTalkInstallFolder 'C:\BizTalk\' -WinSCPEXEDownload 'C:\temp\WinSCP.exe' -WinSCPDllDownload 'C:\temp\WinSCPnet.dll' -WinSCPTargetEXEExists $true -WinSCPDLLTargetExists $false -InstalledAndCorrect $false

            Should -Invoke Write-InstallerStateSnapshot -ModuleName InstallWinSCPForBizTalk.Utils -Times 1 -ParameterFilter {
                $Title -eq 'Check and then potential copy WinSCP to BizTalk results:' -and
                $State.bizTalkInstallFolder -eq 'C:\BizTalk\' -and
                $State.WinSCPEXEDownload -eq 'C:\temp\WinSCP.exe' -and
                $State.WinSCPDllDownload -eq 'C:\temp\WinSCPnet.dll' -and
                $State.WinSCPTargetEXEExists -eq $true -and
                $State.WinSCPDLLTargetExists -eq $false -and
                $State.btsWinSCPProductInstalledAndCorrect -eq $false
            }
        }
    }

    Context "Semantic message wrappers" {
        It "writes registry fallback notice as installer error" {
            Mock -ModuleName InstallWinSCPForBizTalk.Utils Write-InstallerError {}

            Write-BizTalkRegistryFallbackNotice

            Should -Invoke Write-InstallerError -ModuleName InstallWinSCPForBizTalk.Utils -Times 1 -ParameterFilter {
                $ErrorMessage -match 'Env:BTSINSTALLPATH does not exist'
            }
        }

        It "writes BizTalk located success triplet" {
            Mock -ModuleName InstallWinSCPForBizTalk.Utils Write-InstallerSuccess {}

            Write-BizTalkLocatedSuccess -ProductName 'BizTalk X' -ProductVersion '1.2.3' -InstallFolder 'C:\BizTalk\'

            Should -Invoke Write-InstallerSuccess -ModuleName InstallWinSCPForBizTalk.Utils -Times 3
            Should -Invoke Write-InstallerSuccess -ModuleName InstallWinSCPForBizTalk.Utils -Times 1 -ParameterFilter {
                $SuccessMessage -eq 'Located BizTalk X version 1.2.3.'
            }
            Should -Invoke Write-InstallerSuccess -ModuleName InstallWinSCPForBizTalk.Utils -Times 1 -ParameterFilter {
                $SuccessMessage -eq "Located in the 'C:\BizTalk\' folder."
            }
        }

        It "writes WinSCP not-installed notice with bang line and details" {
            Mock -ModuleName InstallWinSCPForBizTalk.Utils Get-InstallerBannerLine { '!!!' }
            Mock -ModuleName InstallWinSCPForBizTalk.Utils Write-InstallerSuccess {}

            Write-WinSCPNotInstalledNotice -WinSCPVersion '5.15.4'

            Should -Invoke Write-InstallerSuccess -ModuleName InstallWinSCPForBizTalk.Utils -Times 4
            Should -Invoke Write-InstallerSuccess -ModuleName InstallWinSCPForBizTalk.Utils -Times 2 -ParameterFilter {
                $SuccessMessage -eq '!!!'
            }
            Should -Invoke Write-InstallerSuccess -ModuleName InstallWinSCPForBizTalk.Utils -Times 1 -ParameterFilter {
                $SuccessMessage -eq 'Microsoft BizTalk Server folder and needs to be installed.'
            }
        }

        It "writes two installer errors describing missing BizTalk installation" {
            $script:errorMessages = @()
            Mock -ModuleName InstallWinSCPForBizTalk.Utils Write-InstallerError {
                param([string] $ErrorMessage)
                $script:errorMessages += $ErrorMessage
            }

            Write-BizTalkNotLocatedError

            $script:errorMessages.Count | Should -Be 2
            $script:errorMessages[0] | Should -Match 'BTSINSTALLPATH'
            $script:errorMessages[1] | Should -Match 'Please confirm'
        }
    }
}