# Define the application name
$appName = "Microsoft SQL Server Compact 4.0 SP1"

$ErrorActionPreference = "Stop"

# Create the log directory if it doesn't exist
$logDirectory = "C:\Support"
if (-not (Test-Path -Path $logDirectory)) {
    New-Item -ItemType Directory -Path $logDirectory | Out-Null
}

# Define the log file path
$logFile = Join-Path -Path $logDirectory -ChildPath "$(Get-Date -Format 'dd.MM.yyyy.HH.mm.ss')-log.txt"

# Function to write log messages
function Write-Log {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message
    )
    
    $timestamp = Get-Date -Format "dd-mm-yyyy HH:mm:ss"
    $logMessage = "$timestamp - $Message"
    
    Add-Content -Path $logFile -Value $logMessage
    Write-Output $logMessage
}

    # Get the list of installed applications from the registry
    $installedApps = Get-ChildItem 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall' | ForEach-Object { Get-ItemProperty $_.PsPath }

    # Filter the list to find the application
    $foundApp = $installedApps | Where-Object { $_.DisplayName -like "*$appName*" }

    # Check if the application is installed
    if ($foundApp) {
        Write-Log -Message "$appName is installed."
        Write-Log -Message "Exiting Script"
        exit 0
    } else {
        Write-Log -Message "$appName is not installed."
        $downloadPath = "C:\Support\SSCERuntime_x64-ENU.exe"

        # Get the download link    
        $Downloadurl = "https://download.microsoft.com/download/F/F/D/FFDF76E3-9E55-41DA-A750-1798B971936C/ENU/SSCERuntime_x64-ENU.exe"
        
        try {
            # Download the MSI file
            Write-Log -Message "Downloading $appName..."
            Start-BitsTransfer -Source $Downloadurl -Destination $downloadPath
            
            # Log the download success
            Write-Log -Message "$appName downloaded successfully"
        } catch {
            # Log the download failure
            Write-Log -Message "Failed to download $appName  $_"
            exit
        }
    
        try {
            Write-Log -Message "Installing $appName..."
            cmd /c $downloadPath /i /x:C:\support /q
            $MSIPath = "C:\support\SSCERuntime_x64-ENU.msi"
            Start-Process msiexec.exe -Wait -ArgumentList "/i $MSIPath /qn /norestart" 
            Write-Log -Message "$appName has been installed."
        } catch {
            Write-Log -Message "Failed to install $appName $_"
        }
    }

