<#
.SYNOPSIS
This script is designed to Download and isntall the latest version of Office365 on a terminal server using the current ODT.

.DESCRIPTION
[Insert a more detailed description of what the script does here.]

.PARAMETER Creator
Name: Phillip Anderson
Email: panderson@integrate-it.com.au

.PARAMETER DateCreated
The date when this script was created.
07/11/2023

.EXAMPLE
Invoke-Expression(New-Object Net.WebClient).DownloadString('https://raw.githubusercontent.com/integrateitaus/Public/main/DeployOffice365Apps.psm1'); DeployOffice365Apps


.NOTES
[Insert any additional notes about the script here.]


#>
##########

$WorkingDir = "C:\Support"
$ErrorActionPreference = "Stop"

$LogPath = "$WorkingDir\$env:computername-365AppInstall.log"
<#function Set-ErrorLogDestination {
    $LocalErrorLog = "$WorkingDir\$env:computername-365AppInstall.log"
    #$NetErrorLog = "$SourceDir\$env:computername-365AppInstall.log"

    #Set logfile destination
    if (Test-Path -Path $NetErrorLog) {
        $LogPath = $NetErrorLog

    } else {
        $LogPath = $LocalErrorLog
    }

    if (-not (Test-Path -Path $WorkingDir -PathType Container)) {
        New-Item -Path $WorkingDir -ItemType Directory
    }
}
#>
#Set-ErrorLogDestination



function InstallOfficeODT {

    $odt_url = "https://www.microsoft.com/en-us/download/confirmation.aspx?id=49117"
    $page = Invoke-WebRequest -Uri $odt_url
    $latest_version_url = $page.Links | Where-Object { $_.href -like "*officedeploymenttool*.exe" } | Select-Object -ExpandProperty href -First 1

    $ODT_selfextractor = "$WorkingDir\officedeploymenttool_*.exe"
    $odt_extracted_folder = "$WorkingDir\ODT"

    # Download Office ODT
    try { 
        Write-Output "Downloading Office ODT"
        Start-BitsTransfer -Source $latest_version_url -Destination $WorkingDir
    } catch {
        Write-Output "Error Downloading Office ODT: $_"
        Add-Content -Path $LogPath -Value "Error Downloading Office ODT: $_"
    } 

    # Extract Office ODT
    try { 
        Write-Output "Extracting Office ODT"
        Start-Process $ODT_selfextractor -ArgumentList "/extract:$odt_extracted_folder /quiet /passive" -Wait
    } catch {
        Write-Output "Error Extracting Office ODT: $_"
        Add-Content -Path $LogPath -Value "Error Extracting Office ODT: $_"
    } 
}


function CreateOfficeConfigxml {
    # Create an XML configuration file for the Office 365 installation
    try { 
# Create an XML configuration file for the Office 365 installation
@"
<Configuration>
    <Add OfficeClientEdition="64" Channel="Current" MigrateArch="TRUE">
        <Product ID="O365ProPlusRetail">
            <Language ID="MatchOS" />
            <ExcludeApp ID="Groove" />
            <ExcludeApp ID="Lync" />
        </Product>
    </Add>
    <Property Name="SharedComputerLicensing" Value="1" />
    <Property Name="AUTOACTIVATE" Value="1" /> 
    <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
    <Property Name="DeviceBasedLicensing" Value="0" />
    <Property Name="SCLCacheOverride" Value="0" />
    <Updates Enabled="TRUE" />
    <RemoveMSI />
    <AppSettings>
        <User Key="software\microsoft\office\16.0\common" Name="autoorgidgetkey" Value="1" Type="REG_DWORD" App="office16" Id="L_AutoOrgIDGetKey" />
        <User Key="software\microsoft\office\16.0\excel\options" Name="defaultformat" Value="51" Type="REG_DWORD" App="excel16" Id="L_SaveExcelfilesas" />
        <User Key="software\microsoft\office\16.0\powerpoint\options" Name="defaultformat" Value="27" Type="REG_DWORD" App="ppt16" Id="L_SavePowerPointfilesas" />
        <User Key="software\microsoft\office\16.0\word\options" Name="defaultformat" Value="" Type="REG_SZ" App="word16" Id="L_SaveWordfilesas" />
    </AppSettings>
    <Display Level="Full" AcceptEULA="TRUE" />
</Configuration>
"@ | Out-File -FilePath "c:\Support\odt\soe_configuration.xml" -Encoding ASCII

    } catch {
        Write-Host "Error: $_"
        Add-Content -Path "$LogPath" -Value "Error: $_"
        exit 1
    }
}


function InstallOffice365 {
    # Install Office 365 using the configuration file
    try {
        Change user /install
        Start-Process -FilePath "$WorkingDir\ODT\setup.exe" -ArgumentList "/configure $WorkingDir\ODT\soe_configuration.xml" -Wait -ErrorAction Stop
        Write-Output "Office 365 installed successfully"
    } catch {
        Write-Output "Error installing Office 365: $_"
        Add-Content -Path "$LogPath" -Value "Error installing Office 365: $_"
        exit 1
    } finally {
    Change user /Execute
 
}

}

function removeOffice365setupfiles {
    
    # Remove the Office Deployment Tool and configuration file
    try {
        Remove-Item -Path "$WorkingDir\Setup.exe" -ErrorAction Stop
        Remove-Item -Path "$WorkingDir\configuration.xml" -ErrorAction Stop
        Write-Output "Office Deployment Tool and configuration file removed successfully"
    } catch {
        Write-Output "Error removing Office Deployment Tool and configuration file: $_"
        Add-Content -Path "$LogPath" -Value "Error removing Office Deployment Tool and configuration file: $_"
        exit 1
    }
}
function DeployOffice365Apps {
    try {
        InstallOfficeODT
        CreateOfficeConfigxml
        InstallOffice365
        RemoveOfficeODT
    } catch {
        Write-Output "Error deploying Office 365 apps: $_"
        Add-Content -Path "$LogPath" -Value "Error deploying Office 365 apps: $_"
        exit 1
    }
}






