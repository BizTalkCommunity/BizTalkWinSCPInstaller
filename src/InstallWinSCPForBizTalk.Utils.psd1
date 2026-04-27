@{
    RootModule = 'InstallWinSCPForBizTalk.Utils.psm1'
    ModuleVersion = '2.0.0'
    GUID = 'fc3948e4-72ec-4e34-8b9d-7f4f2d79a1af'
    Author = 'BizTalk Community'
    CompanyName = 'BizTalk Community'
    Copyright = '(c) BizTalk Community. All rights reserved.'
    Description = 'Output, diagnostics, and semantic message helpers for InstallWinSCPForBizTalk.'
    PowerShellVersion = '5.1'
    FunctionsToExport = @(
        'Write-InstallerError',
        'Write-InstallerSuccess',
        'Get-InstallerBannerLine',
        'Write-InstallerDelimitedMessage',
        'Write-InstallerSectionHeader',
        'Write-InstallerBangError',
        'Write-InstallerFinalOutcome',
        'Write-InstallerStateSnapshot',
        'Write-BizTalkSearchSnapshot',
        'Write-BizTalkCuSearchSnapshot',
        'Write-ExistingWinSCPSnapshot',
        'Write-TargetFolderPreparationSnapshot',
        'Write-NuGetDownloadSnapshot',
        'Write-WinSCPDownloadSnapshot',
        'Write-WinSCPCopySnapshot',
        'Write-BizTalkRegistryFallbackNotice',
        'Write-BizTalkNotLocatedError',
        'Write-BizTalkLocatedSuccess',
        'Write-BizTalkCuDetectionStart',
        'Write-WinSCPNotInstalledNotice'
    )
    CmdletsToExport = @()
    VariablesToExport = @()
    AliasesToExport = @()
    PrivateData = @{
        PSData = @{
            Tags = @('BizTalk', 'WinSCP', 'Installer', 'Diagnostics')
            ProjectUri = 'https://github.com/BizTalkCommunity/BizTalkWinSCPInstaller'
        }
    }
}