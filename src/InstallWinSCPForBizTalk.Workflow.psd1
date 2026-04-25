@{
    RootModule = 'InstallWinSCPForBizTalk.Workflow.psm1'
    ModuleVersion = '1.0.0'
    GUID = '5a1ddf6a-ff6a-4e45-a2bf-5f1794789f52'
    Author = 'BizTalk Community'
    CompanyName = 'BizTalk Community'
    Copyright = '(c) BizTalk Community. All rights reserved.'
    Description = 'Workflow phase orchestration helper functions for InstallWinSCPForBizTalk.'
    PowerShellVersion = '5.1'
    FunctionsToExport = @(
        'Invoke-BizTalkDetectionPhase',
        'Invoke-ForceInstallPrerequisitePhase',
        'Invoke-CuDetectionPhase',
        'Invoke-ExistingWinSCPCheckPhase',
        'Invoke-NonAdminWarningPhase',
        'Invoke-VersionValidationPhase',
        'Invoke-DownloadFolderPreparationPhase',
        'Invoke-NuGetDownloadPhase',
        'Invoke-WinSCPPackageDownloadPhase',
        'Invoke-WinSCPCopyPhase'
    )
    CmdletsToExport = @()
    VariablesToExport = @()
    AliasesToExport = @()
    PrivateData = @{
        PSData = @{
            Tags = @('BizTalk', 'WinSCP', 'Installer', 'Workflow')
            ProjectUri = 'https://github.com/BizTalkCommunity/BizTalkWinSCPInstaller'
        }
    }
}