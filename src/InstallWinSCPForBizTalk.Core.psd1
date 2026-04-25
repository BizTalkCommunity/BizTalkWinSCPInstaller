@{
    RootModule = 'InstallWinSCPForBizTalk.Core.psm1'
    ModuleVersion = '1.0.0'
    GUID = 'ec1d6f27-2e27-45f5-89a3-8c6f9ca5f001'
    Author = 'BizTalk Community'
    CompanyName = 'BizTalk Community'
    Copyright = '(c) BizTalk Community. All rights reserved.'
    Description = 'Core decision, validation, and installation helper functions for InstallWinSCPForBizTalk.'
    PowerShellVersion = '5.1'
    FunctionsToExport = @(
        'Resolve-WinSCPPackageLayout',
        'Search-BTSCumulativeUpdate',
        'Get-BTSCumulativeUpdateByDisplayName',
        'Test-IsAdministrator',
        'Get-InstallExecutionPlan',
        'Test-WinSCPVersionString',
        'Resolve-BizTalkInstallFolder',
        'Get-NuGetDownloadPlan',
        'Get-WinSCPPackageDownloadPlan',
        'Invoke-WinSCPTargetInstall',
        'Get-PackageReadinessState',
        'Get-FinalExecutionOutcome',
        'Get-BizTalkVersionFromProductCode',
        'Get-WinSCPVersionForBizTalk'
    )
    CmdletsToExport = @()
    VariablesToExport = @()
    AliasesToExport = @()
    PrivateData = @{
        PSData = @{
            Tags = @('BizTalk', 'WinSCP', 'Installer')
            ProjectUri = 'https://github.com/BizTalkCommunity/BizTalkWinSCPInstaller'
        }
    }
}