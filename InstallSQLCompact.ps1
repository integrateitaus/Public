$regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
$AppName = "Microsoft SQL Server Compact Edition 4.0 SP1 x64"
$installed = Get-ItemProperty -Path $regPath | Where-Object { $_.DisplayName -eq $AppName }

if ($installed) {
    Write-Host "Microsoft SQL Server Compact Edition 4.0 SP1 x64 is already installed."
} else {
        
    $downloadPath = "C:\Support\SSCERuntime_x64-ENU.exe"

    # Function to download MYOB Accountright
    function DownloadApp {
        # Get the download link    
        $Downloadurl = "https://download.microsoft.com/download/F/F/D/FFDF76E3-9E55-41DA-A750-1798B971936C/ENU/SSCERuntime_x64-ENU.exe"
        
        try {
            # Download the MSI file
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



cmd /c SSCERuntime_x64-ENU.exe /i /x:C:\support /q

