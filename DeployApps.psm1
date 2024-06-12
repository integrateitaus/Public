
<#
Applications to install:
Install-Chrome -Wait
Install-Adobe -Wait
Install-Teams -Wait
Install-Firefox -Wait
Install-Zoom -Wait
Install-EdgeEnt
#>

###########################################
#one liner to run from powershell
#Invoke-Expression(New-Object Net.WebClient).DownloadString('https://raw.githubusercontent.com/integrateitaus/Public/main/DeployApps.psm1'); & $using:Install-Bluebeam
#Invoke-Expression(New-Object Net.WebClient).DownloadString('https://raw.githubusercontent.com/integrateitaus/Public/main/DeployApps.psm1'); Install-Bluebeam

<#Todo:
22.11.2023: Test bluebeam installer

###########################################
#>
Invoke-Expression (Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/integrateitaus/Public/b5245c289179890600506f6fa9aff1ddf8ab0a27/Functions/Write-Log.psm1' -UseBasicParsing).Content

$WorkingDir = "C:\Support\"
$ErrorActionPreference = "Stop"

$LogPath = "$WorkingDir\$env:computername-AppInstall.log"


#Enable RDS Install Mode
#Change User /Install

###########################################
#Install Google Chrome

    $chrome_url = "https://dl.google.com/tag/s/appguid%3D%7B8A69D345-D564-463C-AFF1-A69D9E530F96%7D%26iid%3D%7B9F1E6C2C-6E2C-7D4D-9D2B-3D4D9C5A5E3D%7D%26lang%3Den%26browser%3D5%26usagestats%3D1%26appname%3DGoogle%2520Chrome%26needsadmin%3Dprefers%26ap%3Dx64-stable-statsdef_1%26brand%3DGCEA/dl/chrome/install/googlechromestandaloneenterprise64.msi"
    $chrome_msi = "$WorkingDir\googlechromestandaloneenterprise64.msi"

    function Install-Chrome {
       # Set-ErrorLogDestination
        #Download Google Chrome
        try { 
              Write-Output "Downloading Google Chrome MSI"
              Start-BitsTransfer -Source $chrome_url  -Destination $chrome_msi
              Add-Content -Path $LogPath -Value "Sucessfully Downloading Google Chrome: $_"
        } catch {
                 Write-Output "Error Downloading Google Chrome: $_"
                 Add-Content -Path $LogPath -Value "Error Downloading Google Chrome: $_"
        } 
    
        #Install Google Chrome
            try { 
                Change User /Install
                  Write-Output "Installing Google Chrome"
                  Start-Process msiexec.exe -Wait -ArgumentList "/i $chrome_msi /qn /norestart ALLUSERS=1"
                  Write-Output "Google Chrome installed successfully"
                  Add-Content -Path $LogPath -Value "Google Chrome installed successfully: $_"
        } catch {
                Write-Output "Error installing Google Chrome: $_"
                Add-Content -Path $LogPath -Value "Error installing Google Chrome: $_"
        } finally {
                Change user /Execute
         
        }
    }


######################################
$WorkingDir = "C:\Support\"



function Install-Teams {

    $sourceUrl = "https://statics.teams.cdn.office.net/production-windows-x64/enterprise/webview2/lkg/MSTeams-x64.msix"
    $msixPath = "$WorkingDir\MSTeams-x64.msix"

    # Check if the registry value exists and its data is 1
    $RegistryPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\AppModelUnlock"
    $RegistryValueName = "AllowAllTrustedApps"
    $RegistryValueData = 1

    if ((Get-ItemProperty -Path $RegistryPath -Name $RegistryValueName -ErrorAction SilentlyContinue) -and ((Get-ItemProperty -Path $RegistryPath -Name $RegistryValueName).$RegistryValueName -eq $RegistryValueData)) {
        Write-Host "Windows Sideload apps feature is already enabled."
    } else {
        # Check if the registry path exists, and create any missing keys
        if (-not (Test-Path -Path $RegistryPath -PathType Container)) {
            New-Item -Path $RegistryPath -Force | Out-Null
        }

        # Check if the DWORD registry value exists
        if (-not (Test-Path -Path "$RegistryPath\$RegistryValueName")) {
            Try {
                # Create the new DWORD registry value
                New-ItemProperty -Path $RegistryPath -Name $RegistryValueName -Value $RegistryValueData -PropertyType DWord | Out-Null
                Write-Host "Windows Sideload apps feature has been enabled."
            } Catch {
                Write-Host "Failed to enable Windows Sideload apps feature."
            }
        } else {
            Write-Host "Windows Sideload apps feature is enabled."
        }
    }



    # Check if the MSIX file already exists
    if (Test-Path $msixPath) {
        # Get the size of the existing file
        $existingSize = (Get-Item $msixPath).Length

        # Get the size of the file from the source URL
        $webRequest = [System.Net.WebRequest]::Create($sourceUrl)
        $webResponse = $webRequest.GetResponse()
        $sourceSize = $webResponse.ContentLength

        # Compare the sizes
        if ($existingSize -eq $sourceSize) {
            Write-Host "$msixPath already exists and its size matches the source."
        } else {
            Write-Host "$msixPath already exists but its size does not match the source. Re-downloading..."
            try {
                Remove-Item $msixPath
                write-host "Removed $msixPath"
            } catch {
                Write-Host "Failed to remove $msixPath."
            }
        }
    } else {
        try {
            Write-Host "Downloading the Modern Teams MSIX file from $sourceUrl"
            Start-BitsTransfer -Source $sourceUrl -Destination $msixPath
            Write-Host "Download successful."
        } catch {
            Write-Host "Failed to download the Modern Teams MSIX file."
        }
    }

    # Install the new version of MS Teams
    try {
        Write-Log "Installing Modern Teams"
        Change user /install
        Dism /Online /Add-ProvisionedAppxPackage /PackagePath:$msixPath /SkipLicense
        Change user /execute
    } catch {
        Write-Host "Failed to add provisioned Appx package: $_"
    }
    # Uninstall Classic Teams
    function Uninstall-ClassicTeams {
        $classicTeamsExe = "C:\Program Files (x86)\Teams Installer\Teams.exe"
        if (Test-Path $classicTeamsExe) {
            try {
                Write-Output "Uninstalling Classic Teams"
                Start-Process -FilePath $classicTeamsExe -ArgumentList "--uninstall", "--s"
                Write-Output "Classic Teams uninstalled successfully"
                Add-Content -Path $LogPath -Value "Classic Teams uninstalled successfully"
            } catch {
                Write-Output "Error uninstalling Classic Teams: $_"
                Add-Content -Path $LogPath -Value "Error uninstalling Classic Teams: $_"
            }
        } else {
            Write-Output "Classic Teams is not installed"
            Add-Content -Path $LogPath -Value "Classic Teams is not installed"
        }
    }

    # Call the Uninstall-ClassicTeams function
    Uninstall-ClassicTeams
}





##############
# Install Zoom 64-bit for all users
$Zoom_url = "https://zoom.us/client/5.16.6.24712/ZoomInstallerFull.msi?archType=x64"
$zoom_msi = "$WorkingDir\ZoomInstallerFull.msi"

function Install-Zoom {
  #  Set-ErrorLogDestination
    #Download Zoom


    try { 
        Write-Output "Downloading Zoom 64-bit Installer"
        Start-BitsTransfer -Source $zoom_url -Destination $zoom_msi
    } catch {
        Write-Output "Error Downloading Zoom: $_"
        Add-Content -Path $LogPath -Value "Error Downloading Zoom: $_"
    } 

    #Install Zoom
    try { 
        Change user /install
        Write-Output "Installing Zoom x64 for all users"
        Start-Process -FilePath $zoom_msi -ArgumentList "/quiet", "/qn", "/norestart" -Wait
        Write-Output "Zoom 64-bit installed successfully for all users"
        Add-Content -Path $LogPath -Value "Zoom 64-bit installed successfully for all users: $_"
    } catch {
        Write-Output "Error installing Zoom: $_"
        Add-Content -Path $LogPath -Value "Error installing Zoom: $_"
    } 
    finally {
        Change user /execute
    }
}
##########################



# Install Firefox for all users
$firefox_url = "https://download.mozilla.org/?product=firefox-msi-latest-ssl&os=win64&lang=en-US"
$firefox_msi = "$WorkingDir\FirefoxSetup.msi"
function Install-Firefox {
   # Set-ErrorLogDestination
    #Download Firefox
    try { 
        Write-Output "Downloading Firefox Installer"
        Start-BitsTransfer -Source $firefox_url -Destination $firefox_msi
    } catch {
        Write-Output "Error Downloading Firefox: $_"
        Add-Content -Path $LogPath -Value "Error Downloading Firefox: $_"
    } 

    #Install Firefox
    try { 
        Change user /install
        Write-Output "Installing Firefox"
        Start-Process msiexec.exe -Wait -ArgumentList "/i $firefox_msi /qn /norestart ALLUSERS=1" 
        Write-Output "Firefox installed successfully"
        Add-Content -Path $LogPath -Value "Firefox installed successfully: $_"
    } catch {
        Write-Output "Error installing Firefox: $_"
        Add-Content -Path $LogPath -Value "Error installing Firefox: $_"
    } finally {
        Change user /execute
    }
    
}

################

# Install Adobe Acrobat Reader DC for all users

$adobe_url = "https://ardownload2.adobe.com/pub/adobe/reader/win/AcrobatDC/2300620360/AcroRdrDC2300620360_en_US.exe"
$adobe_exe = "$WorkingDir\AcroRdrDC_en_US.exe"
function Install-Adobe {
  #  Set-ErrorLogDestination
    #Download Adobe Acrobat Reader DC
    try { 
        Write-Output "Downloading Adobe Acrobat Reader DC Installer"
        Start-BitsTransfer -Source  $adobe_url -Destination $adobe_exe
    } catch {
        Write-Output "Error Downloading Adobe Acrobat Reader DC: $_"
        Add-Content -Path $LogPath -Value "Error Downloading Adobe Acrobat Reader DC: $_"
    } 

    #Install Adobe Acrobat Reader DC
    try { 
     Change user /install
        Write-Output "Installing Adobe Acrobat Reader DC"
        Start-Process -FilePath $adobe_exe -ArgumentList "/sAll", "/rs", "/l", "en_US", "/msi", "/norestart", "EULA_ACCEPT=YES", "SUPPRESS_APP_LAUNCH=YES", "SUPPRESS_APP_RESTART=YES", "ALLUSERS=1" -Wait
        Write-Output "Adobe Acrobat Reader DC installed successfully"
        Add-Content -Path $LogPath -Value "Adobe Acrobat Reader DC installed successfully: $_"
    } catch {
        Write-Output "Error installing Adobe Acrobat Reader DC: $_"
        Add-Content -Path $LogPath -Value "Error installing Adobe Acrobat Reader DC: $_"
    } finally {
       Change user /execute
    }
}


function Install-EdgeEnt {
    # Install Microsoft Edge Enterprise
    $edge_url = "https://msedge.sf.dl.delivery.mp.microsoft.com/filestreamingservice/files/3a274bcd-d249-4afb-9c76-b19dc2016be2/MicrosoftEdgeEnterpriseX64.msi"
    $edge_msi = "$WorkingDir\MicrosoftEdgeEnterpriseX64.msi"

    try {
        Write-Output "Downloading Microsoft Edge Enterprise Installer"
        Start-BitsTransfer -Source $edge_url -Destination $edge_msi
    } catch {
        Write-Output "Error Downloading Microsoft Edge Enterprise: $_"
        Add-Content -Path $LogPath -Value "Error Downloading Microsoft Edge Enterprise: $_"
        exit
    }

    try {
        Change user /install
        Write-Output "Installing Microsoft Edge Enterprise"
        Start-Process msiexec.exe -Wait -ArgumentList "/i $edge_msi /qn /norestart ALLUSERS=1"
        Write-Output "Microsoft Edge Enterprise installed successfully"
        Add-Content -Path $LogPath -Value "Microsoft Edge Enterprise installed successfully: $_"
    } catch {
        Write-Output "Error installing Microsoft Edge Enterprise: $_"
        Add-Content -Path $LogPath -Value "Error installing Microsoft Edge Enterprise: $_"
    } finally {
        Change user /execute
    }
}
