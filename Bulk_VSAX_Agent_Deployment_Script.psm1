#!ps
<#

# Bulk_VSAX_Agent_Deployment_Script.ps1
#
# Created 24/10/2023 by Phillip Anderson
# IntegrateIT Australia
#>


##########################
# Script Variables
##########################
#AgentDownloadURL is the VSA Agent download link for the client - Unique per client & location


[CmdletBinding()]
param (
    [Parameter(Mandatory=$true)]
    [string]$AgentDownloadURL
)

$path = "C:\temp"


##########################

#Check if the C:\temp directory exists
function CreateDirectory() {
    if (!(Test-Path -PathType Container $path)) {
        Write-Output "Creating $path directory"
        New-Item -ItemType Directory -Path $path
    }
    
}
    
function DownloadAndInstallAgent($AgentDownloadURL) {
    #Download the client specific VSA Agent
    Write-Output "Downloading VSA-x Agent"
    Invoke-WebRequest -Uri $AgentDownloadURL -OutFile "C:\temp\VSAX_x64.msi"

    # Wait for the file to download
    while (!(Test-Path "C:\temp\VSAX_x64.msi")) {
        Start-Sleep -Seconds 30
    }

    # Continue with the script
    Write-Output "Download finished. Installing now."
        
    
    #Install the VSA Agent
    If ((Test-Path "$path\VSAX_x64.msi") -eq $true) {
            
        Start-process msiexec.exe -Wait -ArgumentList "/i C:\temp\VSAX_x64.msi /qn"
        Write-Output "Install Finished"
    }
    Else {
        Write-Output "Install Failed: File not found"
    }
}
    
function RemoveDesktopIcon() {
    $Files = Get-ChildItem -Path "C:\Users\*" -Filter "VSA X Manager.lnk" -Recurse -ErrorAction SilentlyContinue -Force
    
    # Remove each found shortcut.
    Foreach ($File in $Files) {
        Remove-Item "$File" -Force
        Write-Output "Policy: Computers: Remove VSA X Manager Desktop Icon Removed: $File"
    }
    
    Write-Output "Policy: Computers: Remove VSA X Manager Desktop Icon: Completed"
}
    
function DeployVSAXAgent($AgentDownloadURL) {
    CreateDirectory
    DownloadAndInstallAgent $AgentDownloadURL
    RemoveDesktopIcon
}


#Export-ModuleMember -Function DeployVSAXAgent
