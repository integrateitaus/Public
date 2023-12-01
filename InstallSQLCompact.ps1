# Define the application name
$appName = "Microsoft SQL Server Compact 4.0 SP1"

function IsAppInstalled {
    # Get the list of installed applications from the registry
    $installedApps = Get-ChildItem 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall' | ForEach-Object { Get-ItemProperty $_.PsPath }

    # Filter the list to find the application
    $foundApp = $installedApps | Where-Object { $_.DisplayName -like "*$appName*" }

    # Check if the application is installed
    if ($foundApp) {
        Write-Output "$appName is installed."
    } else {
        Write-Output "$appName is not installed."
        Write-Output "Starting install process"
    }
}
        
    $downloadPath = "C:\Support\SSCERuntime_x64-ENU.exe"

    # Function to download MYOB Accountright
    function DownloadApp {
        # Get the download link    
       
        $Downloadurl = "https://download.microsoft.com/download/F/F/D/FFDF76E3-9E55-41DA-A750-1798B971936C/ENU/SSCERuntime_x64-ENU.exe"
        
        try {
           
            # Download the MSI file
            write-host "Downloading $AppName..."
            
            Start-BitsTransfer -Source $Downloadurl -Destination $downloadPath
            
            # Log the download success
            Write-Log -Message "$AppName downloaded successfully"
        } catch {
            # Log the download failure
            Write-Log -Message "Failed to download $AppName : $_"
            exit
        }
    }

    #Download Microsoft SQL Server Compact Edition 4.0 SP1 x64
    DownloadApp

    Write-Host "Installing $AppName..."
    cmd /c $downloadPath /i /x:C:\support /q

    $MSIPath = "C:\support\SSCERuntime_x64-ENU.msi"
    Start-Process msiexec.exe -Wait -ArgumentList "/i $MSIPath /qn /norestart" 
    Write-Host "$AppName has been installed."
}


