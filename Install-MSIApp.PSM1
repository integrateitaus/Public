<#
.SYNOPSIS
This script is used to download and install MSI Applications from a URL.

.DESCRIPTION
The script downloads the MSI from a specified URL and installs it on the local machine.

.PARAMETER DownloadURL
The URL from which the MSI file will be downloaded.

.EXAMPLE
# Assign the DownloadURL value
$DownloadURL = ""

# Invoke the expression
Invoke-Expression (New-Object Net.WebClient).DownloadString('https://raw.githubusercontent.com/integrateitaus/Public/main/Install-MSIApp.psm1'); DownloadMSI -DownloadURL $DownloadURL

.NOTES
- This script requires administrative privileges to install the MSI Application.
- The script creates a log file in the specified log directory to track the progress and status of the installation.
#>

#Invoke-Expression (Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/integrateitaus/Public/b5245c289179890600506f6fa9aff1ddf8ab0a27/Functions/Write-Log.psm1' -UseBasicParsing).Content
Invoke-Expression(New-Object Net.WebClient).DownloadString('https://raw.githubusercontent.com/integrateitaus/Public/b5245c289179890600506f6fa9aff1ddf8ab0a27/Functions/Write-Log.psm1')

# Create the log directory if it doesn't exist
$logDirectory = "C:\Support"
if (-not (Test-Path -Path $logDirectory)) {
    New-Item -ItemType Directory -Path $logDirectory | Out-Null
}

# Define the log file path
$logFile = Join-Path -Path $logDirectory -ChildPath "$(Get-Date -Format 'dd.MM.yyyy.HH.mm.ss')-log.txt"

$MSIFileName = $DownloadURL -replace ".*msi/"
$downloadPath = Join-Path $logDirectory $MSIFileName



function Write-Log {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message
    )
    
    $timestamp = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
    $logMessage = "$timestamp - $Message"
    
    Add-Content -Path $logFile -Value $logMessage
    Write-Output $logMessage
}

function DownloadMSI ($DownloadURL){
    Write-Log "Downloading $MSIFileName from $DownloadURL to $downloadPath"
    Start-BitsTransfer -Source $DownloadURL -Destination $downloadPath

    # Wait for the file to download
    while (!(Test-Path $downloadPath)) {
        Start-Sleep -Seconds 10
    }

    # Continue with the script
    Write-Log -Message "Download finished. Installing now."
        
    # Install the MYOB Application
    if (Test-Path $downloadPath) {
        InstallMSI
    }
    else {
        Write-Log -Message "Install Failed: File not found"
    }
}

function InstallMSI {
    # Step 3: Install $MSIFileName
    if (Test-Path "C:\support\$MSIFileName") {
        Write-Log -Message "Changing to Install Mode" 
        cmd.exe /c "Change user /install"

        Write-Log "Installing $MSIFileName"

        # Install MYOBAR 
        try {
            Write-Log -Message "Installing $MSIFileName"  
            Start-Process msiexec.exe -Wait -ArgumentList "/i C:\support\$MSIFileName /qn ALLUSERS=1"
            Write-Log -Message "$MSIFileName installed successfully"
        }
        catch {
            Write-Log -Message "Error occurred during installation: $_" 
            exit
        }
        finally {
            Write-Log -Message "Re-enabling execute mode" 
            cmd.exe /c "Change user /execute"
        }
    }
    else {
        Write-Log -Message "MSI not found, starting download" 
        DownloadMSI
        return
    }
}

# Create the log directory if it doesn't exist
$logDirectory = "C:\Support"
if (-not (Test-Path -Path $logDirectory)) {
    New-Item -ItemType Directory -Path $logDirectory | Out-Null
}

# Define the log file path
$logFile = Join-Path -Path $logDirectory -ChildPath "$(Get-Date -Format 'dd.MM.yyyy.HH.mm.ss')-log.txt"

$MSIFileName = $DownloadURL -replace ".*msi/"
$downloadPath = Join-Path $logDirectory $MSIFileName

# Invoke the script
DownloadMSI
