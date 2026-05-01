@{
    RootModule = 'InstallWinSCPForBizTalk.Init.psm1'
    ModuleVersion = '2.1.0'
    GUID = '1f998d11-8b0f-4669-9f77-b3a9e089de56'
    Author = 'BizTalk Community'
    CompanyName = 'BizTalk Community'
    Copyright = '(c) BizTalk Community. All rights reserved.'
    Description = 'Bootstrap initialization helper functions for InstallWinSCPForBizTalk.'
    PowerShellVersion = '5.1'
    FunctionsToExport = @(
        'Initialize-InstallerBootstrap'
    )
    CmdletsToExport = @()
    VariablesToExport = @()
    AliasesToExport = @()
    PrivateData = @{
        PSData = @{
            Tags = @('BizTalk', 'WinSCP', 'Installer', 'Init')
            ProjectUri = 'https://github.com/BizTalkCommunity/BizTalkWinSCPInstaller'
        }
    }
}
