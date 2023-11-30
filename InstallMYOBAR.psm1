# FILEPATH: script.ps1


#Invoke-Expression(New-Object Net.WebClient).DownloadString('https://raw.githubusercontent.com/integrateitaus/Public/main/InstallMYOBAR.psm1'); InstallMYOB
# Create the log directory if it doesn't exist
$logDirectory = "C:\Support"
if (-not (Test-Path -Path $logDirectory)) {
    New-Item -ItemType Directory -Path $logDirectory | Out-Null
}

# Define the log file path
$logFile = Join-Path -Path $logDirectory -ChildPath "log.txt"


# Function to write log messages
function Write-Log {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message
    )
    
    $timestamp = Get-Date -Format "dd-mm-yyyy HH:mm:ss"
    $logMessage = "$timestamp - $Message"
    
    Add-Content -Path $logFile -Value $logMessage
}

function InstallSQLCompact {
    $regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
    $AppName = "Microsoft SQL Server Compact Edition 4.0 SP1 x64"
    $installed = Get-ItemProperty -Path $regPath | Where-Object { $_.DisplayName -eq $AppName }
    $Downloadurl = "https://download.microsoft.com/download/F/F/D/FFDF76E3-9E55-41DA-A750-1798B971936C/ENU/SSCERuntime_x64-ENU.exe"
    $downloadPath = "C:\Support\SSCERuntime_x64-ENU.exe"
    $MSIPath = "C:\support\SSCERuntime_x64-ENU.msi"
    

    if ($installed) {
        Write-Output "Microsoft SQL Server Compact Edition 4.0 SP1 x64 is already installed."
        InstallMYOB
    } else {
        Write-Output "$appname is not installed."
        Write-Log -Message "$appname is not installed."

        Write-Output "Downloading $appname..."  
        Write-Log -Message "Downloading $appname..."
        
        try {
            # Download the exe file
            Start-BitsTransfer -Source $Downloadurl -Destination $downloadPath
            
            # Log the download success
            Write-Log -Message "$AppName downloaded successfully" 
            write-output "$AppName downloaded successfully"
        } catch {
            # Log the download failure
            Write-Log -Message "Failed to download $AppName : $_" 
            Write-Output "Failed to download $AppName : $_"
            exit
        }

try {
        Write-Output "Extracting $AppName..."
        cmd /c $downloadPath /i /x:C:\support /q
} catch {
        Write-Output "Failed to extract $AppName : $_" 
        Write-Log -Message "Failed to extract $AppName : $_"
        exit
}
try{
        Write-Output "Installing $AppName..." 
        Write-Log -Message "Installing $AppName..."
       
        Start-Process msiexec.exe -Wait -ArgumentList "/i $MSIPath /qn /norestart" 
} catch {
        Write-Output "Failed to install $AppName : $_" 
        Write-Log -Message "Failed to install $AppName : $_"
        exit
}
<#
        # Check if the application is installed
        $installed = Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -eq $AppName }

        if ($installed) {
            Write-Output "$AppName has been installed." 
            Write-Log -Message "$AppName has been installed."
        } else {
            Write-Output "$AppName is not installed." 
            Write-Log -Message "$AppName is not installed."
            exit
        }
   #>
        # Call the installation function
        InstallMYOB
    }
}
InstallSQLCompact

# Function to get the download link
function GetDownloadLink {
    $pageurl = "https://www.myob.com/au/support/downloads"
    
    try {
        # Download the HTML of the webpage
        $html = Invoke-WebRequest -Uri $pageurl -UseBasicParsing
        
        # Parse the HTML to find the download link
        $downloadLink = $html.Links | Where-Object { $_.href -like "*MYOB_AccountRight_Client*.msi" } | Select-Object -First 1
        
        # Print the download link
        $Url = $downloadLink.href
        Write-Output "$Url"
        
    } catch {
        # Log the error
        Write-Log -Message "Failed to retrieve the download link: $_"
        Write-Output "Failed to retrieve the download link: $_"
    }
}

# Call the function
GetDownloadLink
$downloadPath = "c:\support\MYOB_AccountRight_Client.msi"



# Function to download MYOB Accountright
function DownloadMYOBAccountright {
    # Get the download link    
    $Downloadurl = Get-DownloadLink
    
    try {
        # Download the MSI file
        
        Start-BitsTransfer -Source $Downloadurl -Destination $downloadPath
        
        # Log the download success
        Write-Log -Message "MYOB Accountright downloaded successfully"
        write-output "MYOB Accountright downloaded successfully"
    } catch {
        # Log the download failure
        Write-Log -Message "Failed to download MYOB Accountright: $_"
        Write-Output "Failed to download MYOB Accountright: $_"
        exit
    }
}
DownloadMYOBAccountright

# Function to move MYOB shortcuts
function MoveMYOBShortcut {
    param (
        [string]$Publicdesktop = "C:\Users\Public\Desktop",
        [string]$shortcutPattern = "AccountRight 202*.*.lnk"
        )

    # Step 1: Check if the MYOB folder exists on the desktop
    $myobFolder = "C:\Users\Public\Desktop\MYOB"

    if (-not (Test-Path -Path $myobFolder -PathType Container)) {
       Write-Log -Message "$myobFolder folder doesn't exist on the desktop. Exiting script." | Write-Output "$myobFolder folder doesn't exist on the desktop. Exiting script."
        exit
    }

    # Step 2: Move shortcuts to the MYOB folder
    $myobShortcuts = Get-ChildItem -Path $PublicDesktop -Filter $shortcutPattern

    if ($myobShortcuts.Count -gt 0) {
       

        foreach ($shortcut in $myobShortcuts) {
            $destinationPath = Join-Path -Path $myobFolder -ChildPath $shortcut.Name
            Move-Item -Path $shortcut.FullName -Destination $destinationPath -Force
            
            Write-Log -Message "Moved shortcut $($shortcut.Name) to $($destinationPath)"
            write-output "Moved shortcut $($shortcut.Name) to $($destinationPath)"
        }
    } else {
                Write-Log -Message "No MYOB shortcuts found on the desktop."
                Write-Output "No MYOB shortcuts found on the desktop."
    }
}
       



# Function to install MYOB AccountRight
function InstallMYOB {
    # Step 3: Install MYOB AccountRight
    If ((Test-Path "C:\support\MYOB_AccountRight_Client.msi") -eq $true) {
        try { 
            Write-Log -Message "Changing to Install Mode"
            Write-Output "Changing to Install Mode"
            
            cmd.exe /c "Change user /install"

            #Write-Output "Installing MYOB AccountRight"
            Write-Log -Message "Installing MYOB AccountRight" 
            Write-Output "Installing MYOB AccountRight"

            #Install the VSA Agent
            Start-process msiexec.exe -Wait -ArgumentList "/i C:\support\MYOB_AccountRight_Client.msi /qn ALLUSERS=1"
            Write-Log -Message "MYOB AccountRight installed successfully" | Write-Output "MYOB AccountRight installed successfully"
            MoveMYOBShortcut
        }
        catch {
            Write-Log -Message "Error occurred during installation: $_"
            Write-Output "Error occurred during installation: $_"
            #exit
        }
        finally {
            Write-Log -Message "Re-enabling execute mode"
            Write-Output "Re-enabling execute mode"
            cmd.exe /c "Change user /execute"
        }
    }     Else {
        
       Write-Log -Message "MYOB Installer not found, exiting installation"
       Write-Output "MYOB Installer not found, exiting installation"
    }
}

# Call the function
InstallMYOB
