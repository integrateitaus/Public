

#example usage
# Assign the DownloadURL value
$DownloadURL = "https://download.myob.com/arl/msi/MYOB_AccountRight_Client_2024.2.msi"

# Invoke the expression
Invoke-Expression (New-Object Net.WebClient).DownloadString('https://raw.githubusercontent.com/integrateitaus/Public/main/InstallMYOBAR_Manual.psm1'); DownloadandInstallMYOBAR -DownloadURL $DownloadURL



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
       # [Parameter(Mandatory = $true)]
        [string]$Message
    )
    
    $timestamp = Get-Date -Format "dd-mm-yyyy HH:mm:ss"
    $logMessage = "$timestamp - $Message"
    
    Add-Content -Path $logFile -Value $logMessage
    Write-Output $logMessage
}

# Define the MYOB AccountRight version

#Download MYOB AccountRight Client



function DownloadandInstallMYOBAR($DownloadURL) {
    #Download the MYOB AR Client
    
$MYOBARFilename = $downloadURL -replace ".*msi/"
$downloadPath = Join-Path $Logdirectory $MYOBARFilename  

Write-Log "Downloading $MYOBARFilenane from $DownloadURL to $downloadPath"
    Start-BitsTransfer -Source $DownloadURL -Destination $downloadPath

    # Wait for the file to download
    while (!(Test-Path "$downloadPath")) {
        Start-Sleep -Seconds 10
    }

    # Continue with the script
    Write-Log -Message "Download finished. Installing now."
        
    
    #Install the MYOB Application
    If ((Test-Path "$downloadpath") -eq $true) {
            
        InstallMYOB
    }
    Else {
        Write-Log -Message "Install Failed: File not found"
    }
}
    


function MoveMYOBShortcut {
    param (
        [string]$Publicdesktop = "C:\Users\Public\Desktop",
        [string]$shortcutPattern = "AccountRight 202*.*.lnk"
        )

    # Step 1: Check if the MYOB folder exists on the desktop
    $myobFolder = "C:\Users\Public\Desktop\MYOB"

    if (-not (Test-Path -Path $myobFolder -PathType Container)) {
       Write-Log -Message "$myobFolder folder doesn't exist on the desktop. Exiting script." 
        exit
    } else {
        Write-Log -Message "$myobFolder folder found on the desktop. Continuing with the script."}

    # Step 2: Move shortcuts to the MYOB folder
    $myobShortcuts = Get-ChildItem -Path $PublicDesktop -Filter $shortcutPattern

    if ($myobShortcuts.Count -gt 0) {
       

        foreach ($shortcut in $myobShortcuts) {
            $destinationPath = Join-Path -Path $myobFolder -ChildPath $shortcut.Name
            Move-Item -Path $shortcut.FullName -Destination $destinationPath -Force
            
            Write-Log -Message "Moved shortcut $($shortcut.Name) to $($destinationPath)" 
        }
    } else {
                Write-Log -Message "No MYOB shortcuts found on the desktop." 
    }
}
       



# Function to install MYOB AccountRight
function InstallMYOB {
    #InstallSQLCompact
    #DownloadMYOBAccountright
    # Step 3: Install MYOB AccountRight
    If ((Test-Path "C:\support\$MYOBARFilename") -eq $true) {
            Write-Log -Message "Changing to Install Mode" 
            
            cmd.exe /c "Change user /install"

            #Write-Output "Installing MYOB AccountRight"
            

            #Install MYOBAR 
            try {
                Write-Log -Message "Installing MYOB AccountRight"  
                Start-process msiexec.exe -Wait -ArgumentList "/i C:\support\$MYOBARFilename /qn ALLUSERS=1"
                MoveMYOBShortcut
                
            }
            catch {
                Write-Log -Message "Error occurred during installation: $_" 
                exit
            }
            finally {
                Write-Log -Message "Re-enabling execute mode" 
                cmd.exe /c "Change user /execute"
            }
            Write-Log -Message "MYOB AccountRight installed successfully" 
        
    }     Else {
        Write-Log -Message "MYOB Installer not found, starting download" 
        DownloadMYOB

    }
}