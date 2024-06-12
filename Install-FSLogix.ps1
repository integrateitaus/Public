<#
.SYNOPSIS
    Installs FSLogix if it is not already installed or if a newer version is available.

.DESCRIPTION
    This script checks if FSLogix is already installed on the system. If it is, it checks if a newer version is available online. If a newer version is available, it downloads and installs FSLogix. If FSLogix is not installed, it downloads and installs FSLogix.

.PARAMETER WorkingDir
    The directory where the script will download and extract FSLogix.

.PARAMETER LogPath
    The path to the log file where installation logs and errors will be recorded.

.EXAMPLE
    Install-FSLogix -WorkingDir "C:\Support" -LogPath "C:\Support\FSLogixInstall.log"
#>

$WorkingDir = "C:\Support"
$LogPath = Join-Path $WorkingDir "$env:computername-FSLogixInstall.log"
$onlineVersion = "2.9.8884.27471"

function Install-FSLogix {


    # Check if FSLogix is already installed
    $installedVersion = Get-InstalledFSLogixVersion
    if ($installedVersion) {
        Write-Output "FSLogix is installed. Installed version: $installedVersion"
        if ($onlineVersion -gt $installedVersion) {
            Write-Output "Online version of FSLogix is newer than the installed version. Downloading FSLogix Installer..."
            Invoke-FSLogixDownload
        } else {
            Write-Output "FSLogix is already up to date. Skipping installation."
            return
        }
    } else {
        Write-Output "FSLogix is not installed. Downloading FSLogix Installer..."
        Invoke-FSLogixDownload
    }
}

function Get-InstalledFSLogixVersion {


    try {
        $installedVersion = $null
        $fslogixRegistryPath = "HKLM:\SOFTWARE\FSLogix\Apps"
        if (Test-Path -Path $fslogixRegistryPath) {
            $installedVersion = Get-ItemProperty -Path $fslogixRegistryPath
            $installedVersion = $installedVersion.Version
        }
        return $installedVersion
    } catch {
        $errorMessage = "Error getting installed FSLogix version: $_"
        Write-Output $errorMessage
        Add-Content -Path $LogPath -Value $errorMessage
        return $null
    }
}

function Invoke-FSLogixDownload {

    try { 
        Write-Output "Downloading FSLogix Installer"
        Start-BitsTransfer -Source "https://aka.ms/fslogix_download" -Destination (Join-Path $WorkingDir "FSLogix.zip") -ErrorAction Stop
    } catch {
        $errorMessage = "Error Downloading FSLogix: $_"
        Write-Output $errorMessage
        Add-Content -Path $LogPath -Value $errorMessage
        return
    } 

    # Extract FSLogix
    try { 
        # Create FSLogix extracted folder if it doesn't exist
        if (-not (Test-Path -Path (Join-Path $WorkingDir "FSLogix_Apps"))) {
            New-Item -ItemType Directory -Path (Join-Path $WorkingDir "FSLogix_Apps") | Out-Null
        }
        Write-Output "Extracting FSLogix"
        Expand-Archive -Path (Join-Path $WorkingDir "FSLogix.zip") -DestinationPath (Join-Path $WorkingDir "FSLogix_Apps") -Force
    } catch {
        $errorMessage = "Error Extracting FSLogix: $_"
        Write-Output $errorMessage
        Add-Content -Path $LogPath -Value $errorMessage
        return
    } 

    # Install FSLogix
    try { 
        Write-Output "Installing FSLogix"
        Start-Process (Join-Path $WorkingDir "FSLogix_Apps\x64\Release\FSLogixAppsSetup.exe") -ArgumentList "/install", "/quiet", "/norestart", "/log", (Join-Path $WorkingDir "fslogix.txt") -Wait
        $fslogixLogContent = Get-Content -Path (Join-Path $WorkingDir "fslogix.txt")
        Add-Content -Path $LogPath -Value $fslogixLogContent
        Write-Output "FSLogix installed successfully"
        Add-Content -Path $LogPath -Value "FSLogix installed successfully"
    } catch {
        $errorMessage = "Error installing FSLogix: $_"
        Write-Output $errorMessage
        Add-Content -Path $LogPath -Value $errorMessage
    }
}

Install-FSLogix -WorkingDir $WorkingDir -LogPath $LogPath
