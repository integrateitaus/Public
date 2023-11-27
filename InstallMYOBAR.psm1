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


# Function to get the download link
function Get-DownloadLink {
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
    }
}

# Call the function
Get-DownloadLink
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
    } catch {
        # Log the download failure
        Write-Log -Message "Failed to download MYOB Accountright: $_"
        exit
    }
}


# Function to move MYOB shortcuts
function MoveMYOBShortcut {
    param (
        [string]$Publicdesktop = "C:\Users\Public\Desktop",
        [string]$shortcutPattern = "AccountRight 202*.*.lnk"
        )

    # Step 1: Check if the MYOB folder exists on the desktop
    $myobFolder = "C:\Users\Public\Desktop\MYOB"

    if (-not (Test-Path -Path $myobFolder -PathType Container)) {
        Write-Output "$myobFolder doesn't exist. Exiting script."
       Write-Log -Message "$myobFolder folder doesn't exist on the desktop. Exiting script."
    }

    # Step 2: Move shortcuts to the MYOB folder
    $myobShortcuts = Get-ChildItem -Path $PublicDesktop -Filter $shortcutPattern

    if ($myobShortcuts.Count -gt 0) {
       

        foreach ($shortcut in $myobShortcuts) {
            $destinationPath = Join-Path -Path $myobFolder -ChildPath $shortcut.Name
            Move-Item -Path $shortcut.FullName -Destination $destinationPath -Force
            Write-Output "Moved shortcut $($shortcut.Name) to $($destinationPath)"
            Write-Log -Message "Moved shortcut $($shortcut.Name) to $($destinationPath)"
        }
    } else {
        Write-Output "No MYOB shortcuts found on the desktop."
        Write-Log -Message "No MYOB shortcuts found on the desktop."
    }
}
DownloadMYOBAccountright       



# Function to install MYOB AccountRight
function InstallMYOB {
    # Step 3: Install MYOB AccountRight
    If ((Test-Path "C:\support\MYOB_AccountRight_Client.msi") -eq $true) {
        try { 
            Write-Output "Changing to Install Mode"
            Write-Log -Message "Changing to Install Mode"
            
            cmd.exe /c "Change user /install"

            #Write-Output "Installing MYOB AccountRight"
            Write-Log -Message "Installing MYOB AccountRight" 

            #Install the VSA Agent
            Start-process msiexec.exe -Wait -ArgumentList "/i C:\support\MYOB_AccountRight_Client.msi /qn ALLUSERS=1"
            Write-Output "MYOB AccountRight installed successfully"
            Write-Log -Message "MYOB AccountRight installed successfully"
            MoveMYOBShortcut
        }
        catch {
            Write-Output "Error occurred during installation: $_" 
            Write-Log -Message "Error occurred during installation: $_"
            #exit
        }
        finally {
            Write-Output "Re-enabling execute mode"
            Write-Log -Message "Re-enabling execute mode"
            cmd.exe /c "Change user /execute"
        }
    }     Else {
        Write-Output "MYOB Installer not found, exiting installation"
        Write-Log -Message "MYOB Installer not found, exiting installation"
    }
}

# Call the function
InstallMYOB
