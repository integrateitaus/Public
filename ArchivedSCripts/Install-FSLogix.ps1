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

$onlineVersion = "2.9.8884.27471"
$WorkingDir = "C:\temp"
$LogPath = "C:\temp\FSLogixUpdate-Version-$onlineVersion.log"


function Install-FSLogix {


    try { 
        Write-Output "Installing FSLogix"
        $fslogixInstallerPath = Join-Path $WorkingDir "FSLogix_Apps\x64\Release\FSLogixAppsSetup.exe"
        Start-Process $fslogixInstallerPath -ArgumentList "/install", "/quiet", "/norestart", "/log", (Join-Path $WorkingDir "fslogix.txt") -Wait

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



function Invoke-FSLogixDownload {

    try { 
        Write-Output "Downloading FSLogix Installer"
        $fslogixInstallerUrl = "https://aka.ms/fslogix_download"
        Start-BitsTransfer -Source $fslogixInstallerUrl -Destination (Join-Path $WorkingDir "FSLogix.zip") -ErrorAction Stop
    } catch {
        $errorMessage = "Error Downloading FSLogix: $_"
        Write-Output $errorMessage
        Add-Content -Path $LogPath -Value $errorMessage
        return
    } 

    try { 
        # Create FSLogix extracted folder if it doesn't exist
        $fslogixExtractedFolderPath = Join-Path $WorkingDir "FSLogix_Apps"
        if (-not (Test-Path -Path $fslogixExtractedFolderPath)) {
            New-Item -ItemType Directory -Path $fslogixExtractedFolderPath | Out-Null
        }
        Write-Output "Extracting FSLogix"
        Expand-Archive -Path (Join-Path $WorkingDir "FSLogix.zip") -DestinationPath $fslogixExtractedFolderPath -Force
    } catch {
        $errorMessage = "Error Extracting FSLogix: $_"
        Write-Output $errorMessage
        Add-Content -Path $LogPath -Value $errorMessage
        return
    } 
}



try {
    $fslogixRegistryPath = "HKLM:\SOFTWARE\FSLogix\Apps"
    if (Test-Path -Path $fslogixRegistryPath) {
        $installedVersion = (Get-ItemProperty -Path "HKLM:\SOFTWARE\FSLogix\Apps").InstallVersion   
        $installedVersion
    }

} catch {
    $errorMessage = "Error getting installed FSLogix version: $_"
    Write-Output $errorMessage
    Add-Content -Path $LogPath -Value $errorMessage

}

if ($installedVersion) {
    Write-Output "FSLogix is installed. Installed version: $installedVersion"
    if ($onlineVersion -gt $installedVersion) {
        Write-Output "Online version of FSLogix is newer than the installed version. Downloading FSLogix Installer..."
        Invoke-FSLogixDownload -WorkingDir $WorkingDir -LogPath $LogPath
    } else {
        Write-Output "FSLogix is already up to date. Skipping installation."
        Add-Content -Path $LogPath -Value $errorMessage
        exit 0
    }
} else {
    Write-Output "FSLogix is not installed. Downloading FSLogix Installer..."
    Invoke-FSLogixDownload -WorkingDir $WorkingDir -LogPath $LogPath
}


Install-FSLogix -WorkingDir $WorkingDir -LogPath $LogPath
