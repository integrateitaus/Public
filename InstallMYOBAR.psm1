<#
.SYNOPSIS
This script installs Microsoft SQL Server Compact 4.0 SP1 and MYOB AccountRight.

.DESCRIPTION
The script checks if Microsoft SQL Server Compact 4.0 SP1 is installed. If not, it downloads and installs it.
Then, it retrieves the latest download link for MYOB AccountRight from the MYOB website and downloads the MSI file.
After that, it installs MYOB AccountRight and moves the shortcuts to the MYOB folder on the public desktop if the folder exists.

.PARAMETER LogPath
The path where the log files will be created.

.EXAMPLE
New-Log -LogPath "C:\Support\"
Invoke-Expression(New-Object Net.WebClient).DownloadString('https://raw.githubusercontent.com/integrateitaus/Public/main/InstallMYOBAR.psm1'); InstallMYOB
This example creates a log file in the specified path.

.NOTES
Author: Phillip Anderson
Date: 01/12/2023
Version: 1.0
#>
# FILEPATH: InstallMYOBAR.ps1

# Create the log directory if it doesn't exist
$logPath = "C:\Support\"
if (-not (Test-Path -Path $logPath)) {
    $null = New-Item -Path $logPath -ItemType Directory
}

# Import the module
function Write-Log {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $timestamp = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
    $logEntry = "$timestamp - $Message"
    Add-Content -Path $logFile -Value $logEntry
    Write-Output $Message
}

function New-Log {
    # Get the script name
    $scriptName = $MyInvocation.MyCommand.Name.Replace(".ps1", "")

    # Create a log file
    $logFile = Join-Path -Path $LogPath -ChildPath "$scriptName.txt"
    $null = New-Item -Path $logFile -ItemType File
}

New-Log -LogPath "$logPath"

function IsAppInstalled {
    param (
        [string]$appName
    )

    # Get the list of installed applications from the registry
    $installedApps = Get-ChildItem 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall' | ForEach-Object { Get-ItemProperty $_.PsPath }

    # Filter the list to find the application
    $foundApp = $installedApps | Where-Object { $_.DisplayName -like "*$appName*" }

    # Check if the application is installed
    if ($foundApp) {
        Write-Log -Message "$appName is installed."
    } else {
        Write-Log -Message "$appName is not installed."
        Write-Log -Message "Starting install process"
    }
}

function InstallSQLCompact {
    # Define the application name
    $appName = "Microsoft SQL Server Compact 4.0 SP1"

    # Check if the application is installed
    IsAppInstalled $appName

    if ($foundApp) {
        Write-Log -Message "$appName is already installed."
        return
    }

    try {
        $Downloadurl = "https://download.microsoft.com/download/F/F/D/FFDF76E3-9E55-41DA-A750-1798B971936C/ENU/SSCERuntime_x64-ENU.exe"
        $downloadPath = "C:\Support\SSCERuntime_x64-ENU.exe"
        $MSIPath = "C:\support\SSCERuntime_x64-ENU.msi"

        Write-Log -Message "Downloading $appName..."
        # Download the exe file
        Start-BitsTransfer -Source $Downloadurl -Destination $downloadPath

        # Log the download success
        Write-Log -Message "$appName downloaded successfully"

        Write-Log "Extracting $appName..."
        cmd /c $downloadPath /i /x:C:\support /q

        Write-Log -Message "Installing $appName..."
        $installArgs = "/i $MSIPath /qn /norestart"
        $msiExecProcess = Start-Process -FilePath "msiexec.exe" -ArgumentList $installArgs -PassThru
        $msiExecProcess | Wait-Process

        # Check if the application is installed after installation
        IsAppInstalled $appName

        if ($foundApp) {
            Write-Log -Message "$appName installed successfully."
        } else {
            Write-Log -Message "$appName failed to install."
            exit
        }
    } catch {
        Write-Log -Message "Failed to install $appName : $_"
        exit
    }

    Write-Log -Message "Starting MYOB AR Installation..."
    InstallMYOB
}

function GetDownloadLink {
    $pageurl = "https://www.myob.com/au/support/downloads"
    
    try {
        # Download the HTML of the webpage
        $html = Invoke-WebRequest -Uri $pageurl -UseBasicParsing
        
        # Parse the HTML to find the download link
        $downloadLink = $html.Links | Where-Object { $_.href -match "MYOB_AccountRight_Client.*.msi" } | Select-Object -First 1
        Write-Output "$url"

        # Extract the download link URL
        $Url = $downloadLink.href

        
    } catch {
        # Log the error
        Write-Log -Message "Failed to retrieve the download link: $_"
        return
    }
}


Function GetOnlineVersion {
Try {
    $OnlineVersion = GetDownloadLink

        # Extract the download link URL
        $OnlineVersion = $downloadLink.href
        $OnlineVersion = $OnlineVersion.Substring($url.IndexOf("20"))
        $OnlineVersion = $OnlineVersion.Replace(".msi", "")
        
        # Print the modified download link
        Write-Output "Latest available MYOB AccountRight version is $OnlineVersion"
        
    } catch {
        # Log the error
        Write-Log -Message "Failed to retrieve the latest version number: $_"
        return
       }   
}


function CheckMYOBVersion {

    # Get the online version of MYOB AccountRight
    $OnlineVersion = GetOnlineVersion
    try {
        # Get the installed version of MYOB AccountRight
        $installedVersion = (Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -like "MYOB AccountRight*" }).Version

        if ($installedVersion -eq $OnlineVersion) {
            Write-Log -Message "Installed version of MYOB AccountRight matches the online version. exiting script."
            exit
        } else {
            Write-Log -Message "Installed version of MYOB AccountRight does not match the online version."
            InstallMYOB
        }
    } catch {
        Write-Log -Message "Failed to check MYOB version: $_"
        
    }
}


function DownloadMYOBAccountright {
    # Get the download link
    $Downloadurl = GetDownloadLink

    try {
        # Download the MSI file
        Start-BitsTransfer -Source $Downloadurl -Destination $global:downloadPath

        # Log the download success
        Write-Log -Message "MYOB Accountright downloaded successfully"
    } catch {
        # Log the download failure
        Write-Log -Message "Failed to download MYOB Accountright: $_"
        exit
    }
}

function MoveMYOBShortcut {
    param (
        [string]$Publicdesktop = "C:\Users\Public\Desktop",
        [string]$shortcutPattern = "AccountRight 202*.*.lnk"
    )

    # Step 1: Check if the MYOB folder exists on the desktop
    $myobFolder = Join-Path -Path $Publicdesktop -ChildPath "MYOB"

    if (-not (Test-Path -Path $myobFolder -PathType Container)) {
        Write-Log -Message "$myobFolder folder doesn't exist on the desktop. Exiting script."
        exit
    }

    Write-Log -Message "$myobFolder folder exists on the desktop."
    Write-Log -Message "Moving MYOB AR shortcuts to $myobFolder"

    # Step 2: Move shortcuts to the MYOB folder
    $myobShortcuts = Get-ChildItem -Path $PublicDesktop -Filter $shortcutPattern

    if ($myobShortcuts) {
        foreach ($shortcut in $myobShortcuts) {
            $destinationPath = Join-Path -Path $myobFolder -ChildPath $shortcut.Name
            Move-Item -Path $shortcut.FullName -Destination $destinationPath -Force
            Write-Log -Message "Moved shortcut $($shortcut.Name) to $($destinationPath)"
        }
    } else {
        Write-Log -Message "No MYOB shortcuts found on the desktop."
        exit
    }
}

function InstallMYOB {
    # Define the application name
    $appName = "MYOB AccountRight"
    CheckMYOBVersion
    # Check if the application is already installed
    if (IsAppInstalled $appName) {
        Write-Log -Message "$appName is already installed..."
        
        return

    } else {
        Write-Log -Message "Starting $appName installation..."
    }

    # Check if the installer file exists
    if (Test-Path $downloadPath) {
        try {
            Write-Log -Message "$appName Installer found, starting installation..."
            Write-Log -Message "Changing to Install Mode"
            cmd.exe /c "Change user /install"

            Write-Log -Message "Installing $appName"
            # Install MYOB AR
            Start-Process msiexec.exe -Wait -ArgumentList "/i $downloadPath /qn ALLUSERS=1"

            Write-Log -Message "$appName installed successfully"
            MoveMYOBShortcut
        } catch {
            Write-Log -Message "Error occurred during installation: $_"
            return
        } finally {
            Write-Log -Message "Re-enabling execute mode"
            cmd.exe /c "Change user /execute"
        }
    } else {
        DownloadMYOBAccountright
        Write-Log -Message "$downloadPath not found, Starting download"
        return
    }
}

InstallSQLCompact
#DownloadMYOBAccountright
#InstallMYOB

