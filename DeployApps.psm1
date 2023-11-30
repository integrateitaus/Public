
<#
Applications to install:
Install-Chrome -Wait
Install-Adobe -Wait
Install-Teams -Wait
Install-Firefox -Wait
Install-Zoom -Wait
Install-FSLogix -wait
#>

###########################################
#one liner to run from powershell
#Invoke-Expression(New-Object Net.WebClient).DownloadString('https://raw.githubusercontent.com/integrateitaus/Public/main/DeployApps.psm1'); & $using:Install-Bluebeam
#Invoke-Expression(New-Object Net.WebClient).DownloadString('https://raw.githubusercontent.com/integrateitaus/Public/main/DeployApps.psm1'); Install-Bluebeam

<#Todo:
22.11.2023: Test bluebeam installer
#>
###########################################

$WorkingDir = "C:\Support\"
$ErrorActionPreference = "Stop"

$LogPath = "$WorkingDir\$env:computername-AppInstall.log"
<#function Set-ErrorLogDestination {
    $LocalErrorLog = "$WorkingDir\$env:computername-AppInstall.log"
    $NetErrorLog = "$SourceDir\$env:computername-AppInstall.log"

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

Set-ErrorLogDestination
#>

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

$teams_url = "https://teams.microsoft.com/downloads/desktopurl?env=production&plat=windows&arch=x64&managedInstaller=true&download=true"
$teams_msi = "$WorkingDir\Teams_windows_x64.msi"


function Install-Teams {
   # Set-ErrorLogDestination
    #Download Microsoft Teams
    try { 
        Write-Output "Downloading Microsoft Teams Installer"
        Start-BitsTransfer -Source $teams_url -Destination $teams_msi
        } catch {
                 Write-Output "Error Downloading Microsoft Teams Installer: $_"
                }


    #Install Microsoft Teams
    try { 
       Change user /install
        Write-Output "Installing Microsoft Teams"
        Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$teams_msi`" /qn ALLUSERS=1" -Wait -NoNewWindow
        #Start-Process "$teams_exe" -ArgumentList "/s" -Wait -ErrorAction Stop
        Write-Output "Microsoft Teams installed successfully"
        } catch {
                Write-Output "Error installing Microsoft Teams: $_"
                } finally {
                            Change user /Execute
                        }
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

# Install FSLogix for all users
$fslogix_zip_url = "https://aka.ms/fslogix_download"
$fslogix_zip = "$WorkingDir\FSLogix.zip"
$fslogix_extracted_folder = "C:\Support\FSLogix_Apps_*"
$fslogix_exe = "$fslogix_extracted_folder\x64\Release\FSLogixAppsSetup.exe"
$fslogix_log = "$WorkingDir\FSlogix\fslogix.txt"

function Install-FSLogix {
    # Download FSLogix
    try { 
        Write-Output "Downloading FSLogix Installer"
        Start-BitsTransfer -Source $fslogix_zip_url -Destination $fslogix_zip
    } catch {
        Write-Output "Error Downloading FSLogix: $_"
        Add-Content -Path $LogPath -Value "Error Downloading FSLogix: $_"
    } 

    # Extract FSLogix
    try { 
        Write-Output "Extracting FSLogix"
        Expand-Archive -Path $fslogix_zip -DestinationPath "$WorkingDir" -Force
    } catch {
        Write-Output "Error Extracting FSLogix: $_"
        Add-Content -Path $LogPath -Value "Error Extracting FSLogix: $_"
    } 

    # Install FSLogix
    try { 
        Change user /install
        Write-Output "Installing FSLogix"
        Start-Process $fslogix_exe -ArgumentList "/install", "/quiet", "/norestart", "/log", $fslogix_log -Wait
        Write-Output "FSLogix installed successfully"
        Add-Content -Path $LogPath -Value "FSLogix installed successfully: $_"
    } catch {
        Write-Output "Error installing FSLogix: $_"
        Add-Content -Path $LogPath -Value "Error installing FSLogix: $_"
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


