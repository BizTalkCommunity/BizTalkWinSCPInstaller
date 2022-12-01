
NAME
    InstallWinSCPForBizTalk.ps1
    
SYNOPSIS
    Installs WinSCP by detecting the version of Microsoft BizTalk Server 2016 or 2020
    Determines which cumulative update is installed
    Downloads WinSCP using NuGet and installs WinSCP to
    the Microsoft BizTalk Server installation folder.
    
    Credits to Michael Stepensen who created the original script.
    Credits to Nicolas Blatter for testing each new release of the cumulative updates.
    Credits to Sandro Pereira for updating the script and helping improve it.
    Credits to Niclas Öberg for updating the script for BTS 2020 CU4
    
    Supported versions of BizTalk Server and cumulative updates:
    Microsoft BizTalk Server 2016
        CU/FU name Build version  KB number  Release date       WinSCP Version
        CU9 FP3    3.13.357.2     5005480    September 29, 2021 WinSCP 5.19.2
        CU9        3.12.896.2     5005479    August 25, 2021    WinSCP 5.19.2
        CU8 FP3    3.13.349.2     4590075    January 6, 2021    WinSCP 5.15.9
        CU8        3.12.880.2     4583530    December 7, 2020   WinSCP 5.15.9
        CU7 FP3    3.13.340.2     4536185    January 22, 2020   WinSCP 5.15.9
        CU7        3.12.859.2     4528776    January 22, 2020   WinSCP 5.15.9
        CU6 FP3    3.12.843.2     4294900    February 7, 2019   WinSCP 5.13.1
        CU6        3.12.843.2     4477494    February 28, 2019  WinSCP 5.13.1
        CU5 FP3    3.13.324.2     4103503    June 25, 2018      WinSCP 5.13.1
        CU5 Hotfix 3.12.834.2     4132957    November 14, 2018  WinSCP 5.13.1
        CU5        3.12.834.2     4132957    June 25, 2018      WinSCP 5.13.1
        CU4 FP2    3.13.252.2     4094130    April 2, 2018       WinSCP 5.7.7
        CU4        3.12.823.2     4051353    January 30, 2018    WinSCP 5.7.7
        CU3 FP2    3.13.247.2     4054819    November 21, 2017   WinSCP 5.7.7
        CU3 FU1    3.13.177.2     4014788    November 15, 2017   WinSCP 5.7.7
        CU3        3.12.815.2     4039664    September 01, 2017  WinSCP 5.7.7
        CU2        3.12.807.2     4021095    May 26, 2017        WinSCP 5.7.7
        CU1        3.12.796.2     3208238    January 26, 2017    WinSCP 5.7.7
        RTM        3.12.774.2                December 1, 2016    WinSCP 5.7.7
    
    Microsoft BizTalk Server 2020
        CU name Build version KB number  Release day       WinSCP Version
        CU3     3.13.812.0    5007969    November 22, 2021 WinSCP 5.19.2
        CU2     3.13.785.0    5003151    April 19, 2021    WinSCP 5.17.6
        CU1     3.13.759.0    4538666    July 28, 2020     WinSCP 5.17.6
        RTM     3.13.717.0    NA         January 15, 2020  WinSCP 5.15.4
    
    
SYNTAX
    InstallWinSCPForBizTalk-v7.ps1 [[-nugetDownloadFolder] <String>] [-ForceInstall] [-WhatIf] [-Confirm] [<CommonParameters>]
    
    
DESCRIPTION
    Determining which WinSCP version to install depends on your current version of
    Microsoft BizTalk Server 2016 or 2020 you have 
    which Cumulative Update and Feature Pack/Update you have installed.
    Making sure that you have the right version has been historically difficult.
    This script detects the version of the the most recent Cumulative Update
    and Feature Pack and then downloads NuGet; gets the right version of WinSCP
    using NuGet; and then finally copies WinSCP to 
    the Microsoft BizTalk Server installation folder.
    The user must have write access to the Microsoft BizTalk Server 
    installation folder and the temporary folder used during 
    the process, which implies should be running under 
    an administrative PowerShell session.
    By default, the user's TEMP folder  is used, and a subfolder 'nuget' 
    is created below that and used to store the temporary files.
    
    NOTE: This script does not delete the temporary folder after use.
    
    If the correct version of WinSCP is already installed in the BizTalk
    directory, then the script detects the installation and does nothing.
    
    This script supports the -WhatIf parameter to show what 
    would happen but makes no changes to the system.
    
    This script supports a -ForceInstall parameter that overrides the
    WinSCP detection and installs even if the correct WinSCP version
    is installed.
    
    PRODUCTION USE
    This script is designed to be used in a production environment
    when the system does not have access to the internet.
    During testing you will run this script on a system that has
    the same version and cumulative update versions as production.
    Specify the output folder using the -nugetDownloadFolder.
    
    Copy the output folder to production and run the script with
    the -nugetDownloadFolder setting pointing to the folder.
    
    The script will detect that the folder contains NuGet and WinSCP
    in the folder and will not attempt to go to the internet
    to download WinSCP.
    Instead, the script will use the existing files and copy them to
    the Microsoft BizTalk Server installation folder.
    

PARAMETERS
    -nugetDownloadFolder <String>
        nugetDownloadFolder is the temporary folder that will be used 
        to save the NuGet program and the WinSCP components 
        during download.
        This folder is not deleted after the script finishes.
        
        Required?                    false
        Position?                    1
        Default value                (Get-Item Env:TEMP).Value + "\nuget"
        Accept pipeline input?       false
        Accept wildcard characters?  false
        
    -ForceInstall [<SwitchParameter>]
        ForceInstall is used to indicate that even if the existing 
        version of WinSCP is correct in the Microsoft BizTalk Server 
        installation folder it forces the install from the web.
        
        Required?                    false
        Position?                    named
        Default value                False
        Accept pipeline input?       false
        Accept wildcard characters?  false
        
    -WhatIf [<SwitchParameter>]
        WhatIf will show you what version of Microsoft BizTalk Server is installed.
        Then the script will show you what version of WinSCP would install.
        No changes are made to the system.
        
        Required?                    false
        Position?                    named
        Default value                
        Accept pipeline input?       false
        Accept wildcard characters?  false
        
    -Confirm [<SwitchParameter>]
        
        Required?                    false
        Position?                    named
        Default value                
        Accept pipeline input?       false
        Accept wildcard characters?  false
        
    <CommonParameters>
        This cmdlet supports the common parameters: Verbose, Debug,
        ErrorAction, ErrorVariable, WarningAction, WarningVariable,
        OutBuffer, PipelineVariable, and OutVariable. For more information, see 
        about_CommonParameters (http://go.microsoft.com/fwlink/?LinkID=113216). 
    
INPUTS
    
OUTPUTS
    
NOTES
    
    
        Author: Thomas Canter, Sandro Pereira, Michael Stepensen
        Date:   December 13th, 2021
    
    -------------------------- EXAMPLE 1 --------------------------
    
    .\InstallWinSCPForBizTalk.ps1
    
    This is the fully automated installation which will install
    WinSCP based on the installed version of 
    Microsoft BizTalk Server, Cumulative Update and Feature Pack.
    If needed, it will download NuGet and WinSCP to install it.
    If the current version of WinSCP is installed it will do nothing.
    
    
    
    
    -------------------------- EXAMPLE 2 --------------------------
    
    .\InstallWinSCPForBizTalk.ps1 -nugetDownloadFolder WinSCPTemp
    
    This will install WinSCP using WinSCPTemp as the 
    temporary folder to store the files dowloaded.
    
    
    
    
    -------------------------- EXAMPLE 3 --------------------------
    
    .\InstallWinSCPForBizTalk.ps1 -ForceInstall
    
    This will install WinSCP even if the existing version of WinSCP is correct.
    
    
    
    
