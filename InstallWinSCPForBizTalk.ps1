<#
.SYNOPSIS
    Version 8
        Added support for Verbose flag to debug issues
  
    Installs WinSCP by detecting the version of Microsoft BizTalk Server 2016 or 2020
    Determines which cumulative update is installed
    Downloads NuGet and then WinSCP using NuGet and installs WinSCP to
    the Microsoft BizTalk Server installation folder.
  
    Credits to Michael Stepensen who created the original script.
    Credits to Nicolas Blatter for testing each new release of the cumulative updates.
    Credits to Sandro Pereira for updating the script and helping improve it.
    Latest credit to Niclas Öberg for updating the script to support BizTalk Server 2020 CU4.
    Supported versions of BizTalk Server and cumulative updates:
    Microsoft BizTalk Server 2016
        CU/FU name Build version  KB number  Release date       WinSCP Version
        CU9 FP3    3.13.357.2     5005480    September 29, 2021 WinSCP 5.19.2
        CU9        3.12.896.2     5005479    August 25, 2021    WinSCP 5.19.2
        CU8 FP3    3.13.349.2     4590075    January 6, 2021    WinSCP 5.15.9
        CU8        3.12.880.2     4583530    December 7, 2020   WinSCP 5.17.8
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
        CU4     3.13.844.0    5009901    August 22, 2022   WinSCP 5.19.2
        CU3     3.13.812.0    5007969    November 22, 2021 WinSCP 5.17.8
        CU2     3.13.785.0    5003151    April 19, 2021    WinSCP 5.17.8
        CU1     3.13.759.0    4538666    July 28, 2020     WinSCP 5.17.6
        RTM     3.13.717.0    NA         January 15, 2020  WinSCP 5.15.4
  
.DESCRIPTION
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
      
    This script supports the -ForceInstall parameter that overrides the
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
  
.PARAMETER nugetDownloadFolder
    nugetDownloadFolder is the temporary folder that will be used
    to save the NuGet program and the WinSCP components
    during download.
    This folder is not deleted after the script finishes.
.PARAMETER ForceInstall
    ForceInstall is used to indicate that even if the existing
    version of WinSCP is correct in the Microsoft BizTalk Server
    installation folder it forces the install from the web.
.PARAMETER WhatIf
    WhatIf will show you what version of Microsoft BizTalk Server is installed.
    Then the script will show you what version of WinSCP would install.
    No changes are made to the system.
.EXAMPLE
    .\InstallWinSCPForBizTalk.ps1
    This is the fully automated installation which will install
    WinSCP based on the installed version of
    Microsoft BizTalk Server, Cumulative Update and Feature Pack.
    If needed, it will download NuGet and WinSCP to install it.
    If the correct version of WinSCP is installed it will do nothing.
.EXAMPLE
    .\InstallWinSCPForBizTalk.ps1 -nugetDownloadFolder WinSCPTemp
    This will install WinSCP using WinSCPTemp as the
    temporary folder to store the files dowloaded.
.EXAMPLE
    .\InstallWinSCPForBizTalk.ps1 -ForceInstall
    This will install WinSCP even if the existing version of WinSCP is correct.
.NOTES
    Author: Thomas Canter, Sandro Pereira, Michael Stepensen, Niclas Öberg
    Date:   September 20, 2022
#>
[cmdletbinding(SupportsShouldProcess)]
#Parameters
Param(
    [Parameter(
        Mandatory = $false
    )]
    [ValidateNotNullOrEmpty()]
    [string]$nugetDownloadFolder = (Get-Item Env:TEMP).Value + "\nuget",
    [Parameter(
        Mandatory = $false
    )]
    [switch]$ForceInstall
)
#####################################################################
# Function to serach for specific Microsoft BizTalk Server Cumulative updates
#####################################################################
function Search-BTSCumulativeUpdate {
    Param(
        [string] $CumulativeUpdateID,
        [string] $BizTalkVersion
    )
      
    $CUFound = $false
    $CUNameTemplate = "*BizTalk *" + $BizTalkVersion + "*KB*" + $CumulativeUpdateID + "*";
    $CUFound = [bool](Get-ChildItem -path HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.Name -like $CUNameTemplate });
    return $CUFound
}
#####################################################################
# Function to write an error
#####################################################################
function Write-Error {
    Param([string] $ErrorMessage)
    Write-Success -ForegroundColor Red "$ErrorMessage";
}
#####################################################################
# Function to write success
#####################################################################
function Write-Success {
    Param([string] $SuccessMessage)
    Write-Host -ForegroundColor Green "$SuccessMessage";
}
# Default $Continue flag to true, set to false to end the process
$Continue = $true;
  
$DebugPreference = "Continue";
#####################################################################
# Initialize the base default versions
# Assume BizTalk Server 2016 with no Cumulative Update installed
# WinSCP has stored the required files for BizTalk in different
# folders for each minor version
# In the Nuget Archive the WinSCP files are stored in the format:
# WinSCP + Version \ EXE Folder \ EXE File
# Here is the example for WinSCP 5.7.7
# WinSCP.5.7.7\tools\WinSCP.exe
# WinSCP + Version \ EXE Folder \ EXE File
# WinSCP.5.7.7\lib\netstandard2.0\WinSCPnet.dll
# WinSCP + Version \ DLL Folder \ DLL File
# WinSCP 5.7.7
# 5.7.7
#   EXE folder = 'content'
#   DLL folder = 'lib'
# 5.11.*
#   EXE folder = 'content'
#   DLL folder = 'lib'
# 5.13.1
#   EXE folder = 'tools'
#   DLL folder = 'lib\net'
# 5.15.*
#   EXE folder = 'tools'
#   DLL folder = 'lib\netstandard'
# 5.15.9
#   EXE folder = 'tools'
#   DLL folder = 'lib\netstandard'
# 5.17.6
#   EXE folder = 'tools'
#   DLL folder = 'lib\netstandard2.0'
# 5.17.8
#   EXE folder = 'tools'
#   DLL folder = 'lib\netstandard2.0'
# 5.19.2
#   EXE folder = 'tools'
#   DLL folder = 'lib\netstandard2.0'
#####################################################################
$winSCPVersion = "5.7.7"
$winSCPexeFile = "WinSCP.exe";
$winSCPdllFile = "WinSCPnet.dll";
$WinSCPexe = "WinSCP.$winSCPVersion\content\$winSCPexeFile"
$winSCPdll = "WinSCP.$winSCPVersion\lib\$winSCPdllFile"
$hashString = "##############################################################################";
$bangString = "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!";
$upString = "^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^";
##############################################################
# Checking for BizTalk Server
##############################################################
$bizTalkInstallFolder = (Get-Item Env:BTSINSTALLPATH).Value;
$bizTalkInstallFolderExists = Test-Path $bizTalkInstallFolder
$BizTalk2016ProductCode = '{B084F3A7-3E8F-4E7B-B673-EED1715D28ED}';
$BizTalk2020ProductCode = '{205F5836-7512-4A06-9E74-ADC8AFA0EEC5}';
$bizTalkProductCodeCurrent = (get-itemPropertyValue 'HKLM:\SOFTWARE\Microsoft\BizTalk Server\3.0' -Name 'ProductCodeCurrent')
$bizTalkProductName = (get-itemPropertyValue 'HKLM:\SOFTWARE\Microsoft\BizTalk Server\3.0' -Name 'ProductName')
$bizTalkProductVersion = (get-itemPropertyValue 'HKLM:\SOFTWARE\Microsoft\BizTalk Server\3.0' -Name 'ProductVersion')
if ($Continue) {
    Write-Success $hashString
    Write-Success 'Checking if Microsoft BizTalk Server is installed.';
    Write-Success $hashString
    if (-not $bizTalkInstallFolder) {
        Write-Error ('The Env:BTSINSTALLPATH doesn`t exist, checking to see if the path is in the registry HKLM:\SOFTWARE\Microsoft\BizTalk Server\3.0@InstallPath');
        $bizTalkInstallFolder = (get-itemPropertyValue 'HKLM:\SOFTWARE\Microsoft\BizTalk Server\3.0' -Name 'InstallPath')
          
        if (-not $bizTalkInstallFolder) {
            $Continue = false;
            Write-Error "Microsoft BizTalk Server was not located by checking the environment variable BTSINSTALLPATH and the Registry key for BizTalk, exiting the process";
            Write-Error "Please confirm that Microsoft BizTalk Server is installed on this system";
        }
    }
    if ($Continue -and -not $bizTalkInstallFolderExists) {
        $Continue = $false
        Write-Error "$bangString"
        Write-Error "Microsoft BizTalk Server installation was found in the registry.";
        Write-Error "Regardless, the $bizTalkInstallFolder folder does not exist.";
        Write-Error "Although it was found in the registry and/or the BTSINSTALLPATH environment setting";
        Write-Error "Exiting...";
        Write-Error "$bangString"
    }
    else {
        Write-Success "Located $bizTalkProductName version $bizTalkProductVersion.";
        Write-Success "Microsoft BizTalk Server is installed.";
        Write-Success "Located in the `'$bizTalkInstallFolder`' folder.";
        if ($bizTalkProductCodeCurrent -eq $BizTalk2020ProductCode) {
            $BizTalkVersion = "2020";
        }
        elseif ($bizTalkProductCodeCurrent -eq $BizTalk2016ProductCode) {
            $BizTalkVersion = "2016";
        }
        else {
            $Continue = $false;
            Write-Error "`n$bangString"
            Write-Error "Neither Microsoft BizTalk Server 2016 nor Microsoft BizTalk Server 2020 were found."
            Write-Error "Exiting";
            Write-Error "$bangString"
        }
  
    }
}
Write-Verbose  "The result of the search for the BizTalk Server:";
Write-Verbose "`$bizTalkInstallFolder       $bizTalkInstallFolder";
Write-Verbose "`$bizTalkInstallFolderExists $bizTalkInstallFolderExists";
Write-Verbose "`$BizTalk2016ProductCode     $BizTalk2016ProductCode";
Write-Verbose "`$BizTalk2020ProductCode     $BizTalk2020ProductCode";
Write-Verbose "`$bizTalkProductCodeCurrent  $bizTalkProductCodeCurrent";
Write-Verbose "`$bizTalkProductName         $bizTalkProductName";
Write-Verbose "`$bizTalkProductVersion      $bizTalkProductVersion";
  
$winSCPVersion = $null;
$btsKB = "none";
$bizTalkCUVer = "no CU";
$CUFound = $false;
if ($Continue) {
    Write-Success "`n$hashString"
    Write-Success "Determining what version of WinSCP to download"
    Write-Success "$hashString"
    ##############################################################
    # Deciding which is the correct version of WinSCP according
    # to Microsoft BizTalk Server version and cumulative update installed
    ##############################################################
    Write-Success "Detected Microsoft BizTalk Server $BizTalkVersion.";
    Write-Success "Testing to see which Cumulative Update is installed";
    if ($BizTalkVersion -eq "2020") {
        $winSCPVersion = "5.15.4"
        # Microsoft BizTalk Server 2020
                                        #CU name    Build version KB number   Release day         WinSCP Version
        $bts2020_CU4 = "5009901";       #CU4        3.13.844.0    5009901     August 22, 2022     WinSCP 5.19.2
        $bts2020_CU3 = "5007969";       #CU3        3.13.812.0    5007969     November 22, 2021   WinSCP 5.19.2
        $bts2020_CU2 = "5003151";       #CU2        3.13.785.0    5003151     April 19, 2021      WinSCP 5.17.8
        $bts2020_CU1 = "4538666";       #CU1        3.13.759.0    4538666     July 28, 2020       WinSCP 5.17.6
                                        #NonUC      3.13.717.0    NA          January 15, 2020    WinSCP 5.15.4
        if (Search-BTSCumulativeUpdate -CumulativeUpdateID $bts2020_CU4 -BizTalkVersion $BizTalkVersion) {
            # Microsoft BizTalk Server 2020 CU4
            $bizTalkCUVer = 'CU4'
            $btsKB = $bts2020_CU4
            $winSCPVersion = "5.19.2"
            $CUFound = $true
        }
        elseif (Search-BTSCumulativeUpdate -CumulativeUpdateID $bts2020_CU3 -BizTalkVersion $BizTalkVersion) {
            # Microsoft BizTalk Server 2020 CU3
            $bizTalkCUVer = 'CU3'
            $btsKB = $bts2020_CU3
            $winSCPVersion = "5.19.2"
            $CUFound = $true
        }
        elseif (Search-BTSCumulativeUpdate -CumulativeUpdateID $bts2020_CU2 -BizTalkVersion $BizTalkVersion) {
            # Microsoft BizTalk Server 2020 CU2
            $bizTalkCUVer = 'CU2'
            $btsKB = $bts2020_CU2
            $winSCPVersion = "5.17.8"
            $CUFound = $true
        }
        elseif (Search-BTSCumulativeUpdate -CumulativeUpdateID $bts2020_CU1 -BizTalkVersion $BizTalkVersion) {
            # Microsoft BizTalk Server 2020 CU1
            $bizTalkCUVer = 'CU1'
            $btsKB = $bts2020_CU1
            $winSCPVersion = "5.17.6"
            $CUFound = $true
        }
    }
    elseif ($BizTalkVersion -eq "2016") {
        # Microsoft BizTalk Server 2016
        #                               CU name     Build version  KB number  Release date        WinSCP Version
        $bts2016_CU9_FP3 = "5005480";   #CU9 FP3    3.13.357.2     5005480    September 29, 2021  WinSCP 5.19.2
        $bts2016_CU9 = "5005479";       #CU9        3.12.896.2     5005479    August 25, 2021     WinSCP 5.19.2
        $bts2016_CU8_FP3 = "4590075";   #CU8 FP3    3.13.349.2     4590075    January 6, 2021     WinSCP 5.15.9
        $bts2016_CU8 = "4583530";       #CU8        3.12.880.2     4583530    December 7, 2020    WinSCP 5.15.9
        $bts2016_CU7_FP3 = "4536185";   #CU7 FP3    3.13.340.2     4536185    January 22, 2020    WinSCP 5.15.9
        $bts2016_CU7 = "4528776";       #CU7        3.12.859.2     4528776    January 22, 2020    WinSCP 5.15.9
        $bts2016_CU6_FP3 = "4294900";   #CU6 FP3    3.12.843.2     4294900    February 7, 2019    WinSCP 5.13.1
        $bts2016_CU6 = "4477494";       #CU6        3.12.843.2     4477494    February 28, 2019   WinSCP 5.13.1
        $bts2016_CU5_FP3 = "4103503";   #CU5 FP3    3.13.324.2     4103503    June 25, 2018       WinSCP 5.13.1
        $bts2016_CU5Hotfix = "4345385"; #CU5 Hotfix 3.12.834.2     4345385    November 14, 2018   WinSCP 5.13.1
        $bts2016_CU5 = "4132957";       #CU5        3.12.834.2     4132957    June 25, 2018       WinSCP 5.13.1
        $bts2016_CU4_FP2 = "4094130";   #CU4 FP2    3.13.252.2     4094130    April 2, 2018       WinSCP 5.7.7
        $bts2016_CU4 = "4051353";       #CU4        3.12.823.2     4051353    January 30, 2018    WinSCP 5.7.7
        $bts2016_CU3_FP2 = "4054819";   #CU3 FP2    3.13.247.2     4054819    November 21, 2017   WinSCP 5.7.7
        $bts2016_CU3_FU1 = "4014788";   #CU3 FU1    3.13.177.2     4014788    November 15, 2017   WinSCP 5.7.7
        $bts2016_CU3 = "4039664";       #CU3        3.12.815.2     4039664    September 01, 2017  WinSCP 5.7.7
        $bts2016_CU2 = "4021095";       #CU2        3.12.807.2     4021095    May 26, 2017        WinSCP 5.7.7
        $bts2016_CU1 = "3208238";       #CU1        3.12.796.2     3208238    January 26, 2017    WinSCP 5.7.7
                                        #NonCU      3.12.774.0     NA         September 30, 2016  WinSCP 5.7.7
        $CUfound = $false
        if (Search-BTSCumulativeUpdate -CumulativeUpdateID $bts2016_CU9_FP3 -BizTalkVersion $BizTalkVersion) {
            # running Microsoft BizTalk Server 2016 CU9 and FP3 with new WinSCP Version
            $winSCPVersion = "5.19.2"
            $btsKB = $bts2016_CU9_FP3
            $bizTalkCUVer = 'CU9 and FP3'
            $CUFound = $true
        }
        elseif (Search-BTSCumulativeUpdate -CumulativeUpdateID $bts2016_CU9 -BizTalkVersion $BizTalkVersion) {
            # running Microsoft BizTalk Server 2016 CU9 with new WinSCP Version
            $winSCPVersion = "5.19.2"
            $btsKB = $bts2016_CU9
            $bizTalkCUVer = 'CU9'
            $CUFound = $true
        }
        elseif (Search-BTSCumulativeUpdate -CumulativeUpdateID $bts2016_CU8_FP3 -BizTalkVersion $BizTalkVersion) {
            # running Microsoft BizTalk Server 2016 CU8 and FP3 with new WinSCP Version
            $winSCPVersion = "5.15.9"
            $btsKB = $bts2016_CU8_FP3
            $bizTalkCUVer = 'CU8 and FP3'
            $CUFound = $true
        }
        elseif (Search-BTSCumulativeUpdate -CumulativeUpdateID $bts2016_CU8 -BizTalkVersion $BizTalkVersion) {
            # running Microsoft BizTalk Server 2016 CU8 with new WinSCP Version
            $winSCPVersion = "5.15.9"
            $btsKB = $bts2016_CU8
            $bizTalkCUVer = 'CU8'
            $CUFound = $true
        }
        elseif (Search-BTSCumulativeUpdate -CumulativeUpdateID $bts2016_CU7_FP3 -BizTalkVersion $BizTalkVersion) {
            # running Microsoft BizTalk Server 2016 CU7 and FP3 with new WinSCP Version
            $winSCPVersion = "5.15.9"
            $btsKB = $bts2016_CU7_FP3
            $bizTalkCUVer = 'CU7 and FP3'
            $CUFound = $true
        }
        elseif (Search-BTSCumulativeUpdate -CumulativeUpdateID $bts2016_CU7 -BizTalkVersion $BizTalkVersion) {
            # running Microsoft BizTalk Server 2016 CU7 with new WinSCP Version
            $winSCPVersion = "5.15.9"
            $btsKB = $bts2016_CU7
            $bizTalkCUVer = 'CU7'
            $CUFound = $true
        }
        elseif (Search-BTSCumulativeUpdate -CumulativeUpdateID $bts2016_CU6_FP3 -BizTalkVersion $BizTalkVersion) {
            # running Microsoft BizTalk Server 2016 CU6 and FP 3 with new WinSCP Version
            $winSCPVersion = "5.13.1"
            $btsKB = $bts2016_CU6_FP3
            $bizTalkCUVer = 'CU6 and FP3'
            $CUFound = $true
        }
        elseif (Search-BTSCumulativeUpdate -CumulativeUpdateID $bts2016_CU6 -BizTalkVersion $BizTalkVersion) {
            # running Microsoft BizTalk Server 2016 CU6 with new WinSCP Version
            $winSCPVersion = "5.13.1"
            $btsKB = $bts2016_CU6
            $bizTalkCUVer = 'CU6'
            $CUFound = $true
        }
        elseif (Search-BTSCumulativeUpdate -CumulativeUpdateID $bts2016_CU5_FP3 -BizTalkVersion $BizTalkVersion) {
            # running Microsoft BizTalk Server 2016 CU5 and FP3 with new WinSCP Version
            $winSCPVersion = "5.13.1"
            $btsKB = $bts2016_CU5_FP3
            $bizTalkCUVer = 'CU5 and FP3'
            $CUFound = $true
        }
        elseif (Search-BTSCumulativeUpdate -CumulativeUpdateID $bts2016_CU5Hotfix -BizTalkVersion $BizTalkVersion) {
            # running Microsoft BizTalk Server 2016 CU5 Hotfix with new WinSCP Version
            $winSCPVersion = "5.13.1"
            $btsKB = $bts2016_CU5HotFix
            $bizTalkCUVer = 'CU5 Hotfix'
            $CUFound = $true
        }
        elseif (Search-BTSCumulativeUpdate -CumulativeUpdateID $bts2016_CU5 -BizTalkVersion $BizTalkVersion) {
            # running Microsoft BizTalk Server 2016 CU5 with new WinSCP Version
            $winSCPVersion = "5.13.1"
            $btsKB = $bts2016_CU5
            $bizTalkCUVer = 'CU5'
            $CUFound = $true
        }
        # The remainder of the CU/FU combinations use the default WinSCP Version 5.7.7
        elseif (Search-BTSCumulativeUpdate -CumulativeUpdateID $bts2016_CU4_FP2 -BizTalkVersion $BizTalkVersion) {
            # running Microsoft BizTalk Server 2016 CU4 and FP2, using original WinSCP Version
            $bizTalkCUVer = 'CU4 and FP2'
            $btsKB = $bts2016_CU4_FP2
            $CUFound = $true
        }
        elseif (Search-BTSCumulativeUpdate -CumulativeUpdateID $bts2016_CU4 -BizTalkVersion $BizTalkVersion) {
            # running Microsoft BizTalk Server 2016 CU4, using original WinSCP Version
            $bizTalkCUVer = 'CU4'
            $btsKB = $bts2016_CU4
            $CUFound = $true
        }
        elseif (Search-BTSCumulativeUpdate -CumulativeUpdateID $bts2016_CU3_FP2 -BizTalkVersion $BizTalkVersion) {
            # running Microsoft BizTalk Server 2016 CU3 and FP2, using original WinSCP Version
            $bizTalkCUVer = 'CU3 and FP2'
            $btsKB = $bts2016_CU3_FP2
            $CUFound = $true
        }
        elseif (Search-BTSCumulativeUpdate -CumulativeUpdateID $bts2016_CU3_FU1 -BizTalkVersion $BizTalkVersion) {
            # running Microsoft BizTalk Server 2016 CU3 and FU1, using original WinSCP Version
            $bizTalkCUVer = 'CU3 and FU1'
            $btsKB = $bts2016_CU3_FU1
            $CUFound = $true
        }
        elseif (Search-BTSCumulativeUpdate -CumulativeUpdateID $bts2016_CU3 -BizTalkVersion $BizTalkVersion) {
            # running Microsoft BizTalk Server 2016 CU3, using original WinSCP Version
            $bizTalkCUVer = 'CU3'
            $btsKB = $bts2016_CU3
            $CUFound = $true
        }
        elseif (Search-BTSCumulativeUpdate -CumulativeUpdateID $bts2016_CU2 -BizTalkVersion $BizTalkVersion) {
            # running Microsoft BizTalk Server 2016 CU2, using original WinSCP Version
            $bizTalkCUVer = 'CU2 or FU1'
            $btsKB = $bts2016_CU2
            $CUFound = $true
        }
        elseif (Search-BTSCumulativeUpdate -CumulativeUpdateID $bts2016_CU1 -BizTalkVersion $BizTalkVersion) {
            # running Microsoft BizTalk Server 2016 CU1, using original WinSCP Version
            $bizTalkCUVer = 'CU1'
            $btsKB = $bts2016_CU1
            $CUFound = $true
        }
    }
    if ($CUFound) {
        Write-Success "Detected Microsoft BizTalk Server $BizTalkVersion $bizTalkCUVer KB$btsKB";
    }
    else {
        # running Microsoft BizTalk Server without any Cumulative Updates, using original WinSCP Version
        Write-Success "Detected Microsoft BizTalk Server $BizTalkVersion with no cumulative updates.";
    }
    Write-Success "If necessary, this script will download WinSCP $winSCPVersion";
}
Write-Verbose  "The result of the search for the BizTalk Cumulative Update:";
Write-Verbose "`$winSCPVersion = $winSCPVersion";
Write-Verbose "`$btsKB         = $btsKB";
Write-Verbose "`$bizTalkCUVer  = $bizTalkCUVer";
Write-Verbose "`$CUFound       = $CUFound";
  
$btsWinSCPEXEProductVersionInstalled = "None";
$btsWinSCPDLLProductVersionInstalled = "None";
$btsWinSCPProductInstalledAndCorrect = $false;
$winSCPProductVersionRequired = $winSCPVersion;
$btsTargetWinSCPExe = $bizTalkInstallFolder + $winSCPexeFile;
$btsTargetWinSCPDll = $bizTalkInstallFolder + $winSCPdllFile;
if ($Continue) {
    Write-Success "`n$hashString"
    Write-Success "Checking to see if WinSCP $winSCPVersion is already";
    Write-Success "installed in the Microsoft BizTalk Server folder.";
    if ((Test-Path $btsTargetWinSCPExe) -and (Test-Path $btsTargetWinSCPDll)) {
        $btsWinSCPEXEProductVersionInstalled = (get-item $btsTargetWinSCPExe).VersionInfo.ProductVersion;
        $btsWinSCPDLLProductVersionInstalled = (get-item $btsTargetWinSCPDll).VersionInfo.ProductVersion;
        $winSCPProductVersionRequired = $winSCPVersion;
        if ($winSCPVersion.length -gt 2 -and $btsWinSCPEXEProductVersionInstalled.length -gt 2 -and $winSCPVersion.SubString($winSCPVersion.length - 2, 2) -ne '.0' -and $btsWinSCPEXEProductVersionInstalled.SubString($btsWinSCPEXEProductVersionInstalled.length - 2, 2) -eq '.0') {
            $winSCPProductVersionRequired = $winSCPVersion + '.0';
        }
        if ($winSCPProductVersionRequired -eq $btsWinSCPEXEProductVersionInstalled -and $winSCPProductVersionRequired -eq $btsWinSCPDLLProductVersionInstalled) {
            $btsWinSCPProductInstalledAndCorrect = $true;
        }
    }
    if ($btsWinSCPProductInstalledAndCorrect) {
        Write-Success "Detected WinSCP $winSCPVersion is already installed in Microsoft BizTalk Server.";
        if ($ForceInstall -and $btsWinSCPProductInstalledAndCorrect) {
            Write-Success "Reinstalling because ForceInstall was specified.";
        }
        else {
            Write-Success "Skipping installing the already installed version.";
            $Continue = $false;
        }
    }
    else {
        Write-Success "$bangstring"
        Write-Success "WinSCP $winSCPVersion is NOT installed in the";
        Write-Success "Microsoft BizTalk Server folder and needs to be installed.";
    }
    Write-Success "$hashString"
}
Write-Verbose "Check for existing WinSCP results were:";
Write-Verbose "`$btsWinSCPEXEProductVersionInstalled $btsWinSCPEXEProductVersionInstalled";
Write-Verbose "`$btsWinSCPDLLProductVersionInstalled $btsWinSCPDLLProductVersionInstalled";
Write-Verbose "`$winSCPProductVersionRequired        $winSCPProductVersionRequired";
Write-Verbose "`$btsWinSCPProductInstalledAndCorrect $btsWinSCPProductInstalledAndCorrect";
Write-Verbose "`$btsTargetWinSCPExe                  $btsTargetWinSCPExe";
Write-Verbose "`$btsTargetWinSCPDll                  $btsTargetWinSCPDll";
  
  
if ($Continue) {
    $winSCPVersionArray = $winSCPVersion.Split('.');
    if ($winSCPVersionArray.Count -gt 1) {
        [int]$winSCPMajorVer = $winSCPVersionArray[0];
        [int]$winSCPMinorVer = $winSCPVersionArray[1];
    }
    else {
        $Continue = $false;
        Write-Error "WinSCP Version $winSCPVersion is not a recogized version such as 5.7.7.";
        Write-Error $bangString;
    }
    if ($winSCPMajorVer -ne 5) {
        $Continue = $false;
        Write-Error "WinSCP Version $winSCPVersion is not a supported version, only version 5.x.x is supported.";
        Write-Error $bangString;
    }
}
if ($Continue) {
    #use the right version of WinSCP
  
    if ($winSCPMinorVer -eq 17 -or $winSCPMinorVer -eq 19) {
        $winSCPexe = "WinSCP.$winSCPVersion\tools\$winSCPexeFile"
        $winSCPdll = "WinSCP.$winSCPVersion\lib\netstandard2.0\$winSCPdllFile"
    }
    elseif ($winSCPMinorVer -eq 15) {
        $winSCPexe = "WinSCP.$winSCPVersion\tools\$winSCPexeFile"
        $winSCPdll = "WinSCP.$winSCPVersion\lib\netstandard\$winSCPdllFile"
    }
    elseif ($winSCPMinorVer -eq 13) {
        $winSCPexe = "WinSCP.$winSCPVersion\tools\$winSCPexeFile"
        $winSCPdll = "WinSCP.$winSCPVersion\lib\net\$winSCPdllFile"
    }
    elseif ($winSCPVersion -le 11) {
        $WinSCPexe = "WinSCP.$winSCPVersion\content\$winSCPexeFile"
        $winSCPdll = "WinSCP.$winSCPVersion\lib\$winSCPdllFile"
    }
      
    $WinSCPEXEDownload = "$nugetDownloadFolder\$winSCPexe"
    $WinSCPDllDownload = "$nugetDownloadFolder\$winSCPdll"
    $WinSCPEXEDownloadAlreadyExists = Test-Path $WinSCPexe;
    $WinSCPDllDownloadAlreadyExists = Test-Path $$WinSCPDllDownload;
    $nugetDownloadFolderAlreadyExists = Test-Path $nugetDownloadFolder;
    $nugetDownloadFolderExists = $nugetDownloadFolderAlreadyExists;
    Write-Success "`n$hashString"
    Write-Success "Preparing the output folder to store the Nuget and WinSCP downloads"
    Write-Success "$hashString"
    if ($Continue -and -not $nugetDownloadFolderAlreadyExists -and -not $btsWinSCPProductInstalledAndCorrect) {
        ##############################################################
        # Prepare output folder
        ##############################################################
        if ($PSCmdlet.ShouldProcess("$nugetDownloadFolder", "Create Folder")) {
            Write-Success ("The target folder `'$nugetDownloadFolder`' doesn't exist, creating the folder.");
            New-Item -Path $nugetDownloadFolder -ItemType "Directory" > $null
            $nugetDownloadFolderExists = Test-Path $nugetDownloadFolder
            if (-not $nugetDownloadFolderExists) {
                $Continue = $false
                Write-Error "An attempt to use the '" $nugetDownloadFolder "' directory for a download target failed.";
            }
        }
    }
    if ($Continue -and $nugetDownloadFolderExists) {
        Write-Success ("The target folder `'$nugetDownloadFolder`' is ready for use.");
    }
    Write-Verbose "Check and then potential creation of the target folder results:";
    Write-Verbose "`$nugetDownloadFolderAlreadyExists $nugetDownloadFolderExists";
    Write-Verbose "`$nugetDownloadFolderExists        $nugetDownloadFolderExists";
    Write-Verbose "`$WinSCPEXEDownload                $WinSCPEXEDownload";
    Write-Verbose "`$WinSCPDllDownload                $WinSCPDllDownload";
    Write-Verbose "`$WinSCPEXEDownloadAlreadyExists   $WinSCPEXEDownloadAlreadyExists";
    Write-Verbose "`$WinSCPDllDownloadAlreadyExists   $WinSCPDllDownloadAlreadyExists";
}
$targetNugetExe = "$nugetDownloadFolder\nuget.exe"
$targetNugetExeAlreadyExists = Test-Path $targetNugetExe
$targetNugetExeExists = $targetNugetExeAlreadyExists
if ($Continue) {
    ##############################################################
    # Download NuGet
    ##############################################################
    $sourceNugetExe = "https://dist.nuget.org/win-x86-commandline/latest/nuget.exe";
    if (-not $targetNugetExeAlreadyExists -or $ForceInstall) {
        Write-Success "`n$hashString";
        Write-Success "Downloading Nuget from:"
        Write-Success "`t`'$sourceNugetExe`'"
        Write-Success "Storing it in the folder";
        Write-Success "`t`'$nugetDownloadFolder`'"
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        if ($PSCmdlet.ShouldProcess("$sourceNugetExe -OutFile $targetNugetExe", "Run Invoke-WebRequest ")) {
            Invoke-WebRequest $sourceNugetExe -OutFile $targetNugetExe
            $targetNugetExeExists = Test-Path $targetNugetExe
            if (-not $targetNugetExeExists) {
                $Continue = $false
                Write-Error "`n$bangString";
                Write-Error "The download of the Nuget EXE from";
                Write-Error $sourceNugetExe;
                Write-Error "did not succeed";
                Write-Error "$bangString";
            }
        }
    }
    else {
        Write-Success "`n$hashString";
        Write-Success "NuGet already exists in the folder";
        Write-Success $nugetDownloadFolder;
        Write-Success "and doesn't need to be downloaded";
        Write-Success "$hashString";
    }
}
Write-Verbose "Check and then potential download NuGet results:";
Write-Verbose "`$targetNugetExeAlreadyExists $targetNugetExeAlreadyExists";
Write-Verbose "`$targetNugetExeExists        $targetNugetExeExists";
Write-Verbose "`$sourceNugetExe              $sourceNugetExe";
Write-Verbose "`$targetNugetExe              $targetNugetExe";
  
if ($Continue) {
    $WinSCPEXEAlreadyExists = Test-Path $WinSCPEXEDownload
    $WinSCPDLLAlreadyExists = Test-Path $WinSCPDllDownload
    $WinSCPEXEExists = Test-Path $WinSCPEXEDownload
    $WinSCPDLLExists = Test-Path $WinSCPDllDownload
    $getWinSCP = "'$targetNugetExe' Install WinSCP -Version $winSCPVersion -NonInteractive -OutputDirectory '$nugetDownloadFolder'"
    Write-Success "`n$hashString";
    Write-Success "Downloading WinSCP version $winSCPVersion from NuGet";
    Write-Success "`t$getWinSCP";
    Write-Success "Storing it in the folder:"
    Write-Success "`t`'$nugetDownloadFolder`'";
    Write-Success "$hashString";
    if (-not $WinSCPEXEAlreadyExists -or -not $WinSCPDLLAlreadyExists) {
        if ($PSCmdlet.ShouldProcess("$getWinSCP", "Run Command")) {
            Invoke-Expression "& $getWinSCP";
            $WinSCPEXEExists = Test-Path $WinSCPEXEDownload
            $WinSCPDLLExists = Test-Path $WinSCPDllDownload
            if (-not $WinSCPEXEExists -or -not $WinSCPDLLExists) {
                $Continue = $false
                Write-Error "`n$bangString";
                Write-Error "WinSCP $winSCPVersion was not properly downloaded.";
                Write-Error "Check the folder and error messages above:";
                Write-Error "$nugetDownloadFolder";
                Write-Error "And determine what files did download or did not download.";
                Write-Error "$bangString";
            }
        }
    }
}
Write-Verbose "Check and then potential download WinSCP results:";
Write-Verbose "`$getWinSCP              $getWinSCP";
Write-Verbose "`$WinSCPEXEAlreadyExists $WinSCPEXEAlreadyExists";
Write-Verbose "`$WinSCPDLLAlreadyExists $WinSCPDLLAlreadyExists";
Write-Verbose "`$WinSCPDllDownload      $WinSCPDllDownload";
Write-Verbose "`$WinSCPEXEExists        $WinSCPEXEExists";
Write-Verbose "`$WinSCPDLLExists        $WinSCPDLLExists";
  
##############################################################
# Installing WinSCP to Microsoft BizTalk Server Folder
##############################################################
$WinSCPTargetEXEExists = Test-Path $btsTargetWinSCPExe
$WinSCPDLLTargetExists = Test-Path $btsTargetWinSCPDll
if ($Continue -and -not $btsWinSCPProductInstalledAndCorrect) {
    Write-Success "`n$hashString";
    Write-Success "Installing WinSCP";
    Write-Success "$hashString";
    #Copy WinSCP items to Microsoft BizTalk Server Folder
    Write-Success "Copying WinSCP version $winSCPVersion to Microsoft BizTalk Server Folder:";
    Write-Success "`t`'$bizTalkInstallFolder'`.";
    if ($PSCmdlet.ShouldProcess("$WinSCPexe and $WinSCPdll to `'$bizTalkInstallFolder`'", "Copy Files")) {
        Copy-Item $WinSCPEXEDownload $bizTalkInstallFolder
        Copy-Item $WinSCPDllDownload $bizTalkInstallFolder
        $WinSCPTargetEXEExists = Test-Path $btsTargetWinSCPExe
        $WinSCPDLLTargetExists = Test-Path $btsTargetWinSCPDll
        if ($WinSCPTargetEXEExists -and $WinSCPDLLTargetExists) {
            $btsWinSCPProductInstalledAndCorrect = $true;
        }
        else {
            $Continue = $false
            if (-not $WinSCPTargetEXEExists) {
                Write-Error "The $winSCPexeFile file version $winSCPVersion";
                Write-Error "It was not properly copied to the target folder `'$bizTalkInstallFolder`'.";
            }
            if (-not $WinSCPDLLTargetExists) {
                $Continue = $false
                Write-Error "The $winSCPdllFile file version $winSCPVersion";
                Write-Error "Was not properly copied to the target folder `'$bizTalkInstallFolder`'.";
            }  
        }
    }
}
Write-Verbose "Check and then potential copy WinSCP to BizTalk results:";
Write-Verbose "`$bizTalkInstallFolder                $bizTalkInstallFolder";
Write-Verbose "`$WinSCPEXEDownload                   $WinSCPEXEDownload";
Write-Verbose "`$WinSCPDllDownload                   $WinSCPDllDownload";
Write-Verbose "`$WinSCPTargetEXEExists               $WinSCPTargetEXEExists";
Write-Verbose "`$WinSCPDLLTargetExists               $WinSCPDLLTargetExists";
Write-Verbose "`$btsWinSCPProductInstalledAndCorrect $btsWinSCPProductInstalledAndCorrect";
  
      
if ($btsWinSCPProductInstalledAndCorrect) {
    Write-Success "`n$bangString";
    Write-Success "WinSCP $winSCPVersion is installed.";
    Write-Success "Microsoft BizTalk Server`'s SFTP Adapter will use this version of WinSCP.";
    Write-Success "$upString";
}
elseif (-not $WhatIfPreference) {
    Write-Error "`n$bangString";
    Write-Error "Something went wrong during installation and the installation did not work.";
    Write-Error "Please inspect the errors above and resolve them.";
    Write-Error "Exiting...";
    Write-Error "$bangString";
}
elseif ($WhatIfPreference) {
    Write-Success "`n$bangString";
    Write-Success "The parameter -WhatIf was set and this script executed without making";
    Write-Success "any changes and the output should checked to determine if it would have ";
    Write-Success "run correctly.";
    Write-Success "$bangString";
}
