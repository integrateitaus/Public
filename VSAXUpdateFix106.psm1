# Install VSAX Fix1

function ApplyVSAXFix {
    param (
        [string]$WorkingDir = "C:\Support\",
        [string]$LogPath = "$WorkingDir\$env:computername-AppInstall.log",
        [string]$VSAXWinUpdate_zip_url = "https://integrate-it1.vsax.net/Updates/win_agent_update.zip",
        [string]$VSAXWinUpdate_zip = "$WorkingDir\VSAXWinUpdate.zip",
        [string]$processPath = "C:\Program Files\VSA X\PCMonitorSrv.exe"
    )

    try { 
        Write-Output "Downloading VSAXWinUpdate Installer"
        Start-BitsTransfer -Source $VSAXWinUpdate_zip_url -Destination $VSAXWinUpdate_zip
    } catch {
        Write-Output "Error Downloading VSAXWinUpdate: $_"
        Add-Content -Path $LogPath -Value "Error Downloading VSAXWinUpdate: $_"
    } 

    try {
        Write-Output "Stopping the VSAX service"
        Stop-Service -Name "VSAX"
        $process = Get-Process | Where-Object { $_.Path -eq $processPath }
        
        if ($process) {
            Stop-Process -Id $process.Id -Force
            Write-Output "Process stopped"
            
        } else {
            Write-Output "Process not running"
        }

    } catch {
        Write-Output "Error stopping the VSAX service: $_"
        Add-Content -Path $LogPath -Value "Error stopping the VSAX service: $_"
    }

    try { 
        Write-Output "Extracting VSAXWinUpdate"
        Expand-Archive -Path $VSAXWinUpdate_zip -DestinationPath "C:\Program Files\VSA X" -Force
    } catch {
        Write-Output "Error Extracting VSAXWinUpdate: $_"
        Add-Content -Path $LogPath -Value "Error Extracting VSAXWinUpdate: $_"
    } 

    Write-Output "Starting the VSAX service"
    try {
        #Start-Process -FilePath "C:\Program Files\VSA X\PCMonitorSrv.exe"
        Start-Service -Name "VSAX"
    } catch {
        Write-Output "Error starting the VSAX service: $_"
        Add-Content -Path $LogPath -Value "Error starting the VSAX service: $_"
    }
}

#ApplyVSAXFix
