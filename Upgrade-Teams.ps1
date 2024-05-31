
# Download the new version of MS Teams from the CDN
Set-Location "C:\Temp"


$TeamsClassic = Test-Path C:\Users\*\AppData\Local\Microsoft\Teams\current\Teams.exe
$TeamsNew = Get-ChildItem "C:\Program Files\WindowsApps" -Filter "MSTeams_*"

$sourceUrl = "https://statics.teams.cdn.office.net/production-windows-x64/enterprise/webview2/lkg/MSTeams-x64.msix"
$msixPath = "C:\Temp\MSTeams-x64.msix"

function Install-MSTeams {
    # Check if the registry value exists and its data is 1
    $RegistryPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\AppModelUnlock"
    $RegistryValueName = "AllowAllTrustedApps"
    $RegistryValueData = 1

    if ((Get-ItemProperty -Path $RegistryPath -Name $RegistryValueName -ErrorAction SilentlyContinue) -and ((Get-ItemProperty -Path $RegistryPath -Name $RegistryValueName).$RegistryValueName -eq $RegistryValueData)) {
        Write-Host "Windows Sideload apps feature is already enabled."
    } else {
        # Check if the registry path exists, and create any missing keys
        if (-not (Test-Path -Path $RegistryPath -PathType Container)) {
            New-Item -Path $RegistryPath -Force | Out-Null
        }

        # Check if the DWORD registry value exists
        if (-not (Test-Path -Path "$RegistryPath\$RegistryValueName")) {
            Try {
                # Create the new DWORD registry value
                New-ItemProperty -Path $RegistryPath -Name $RegistryValueName -Value $RegistryValueData -PropertyType DWord | Out-Null
                Write-Host "Windows Sideload apps feature has been enabled."
            } Catch {
                Write-Host "Failed to enable Windows Sideload apps feature."
            }
        } else {
            Write-Host "Windows Sideload apps feature is enabled."
        }
    }



    # Check if the MSIX file already exists
    if (Test-Path $msixPath) {
        # Get the size of the existing file
        $existingSize = (Get-Item $msixPath).Length

        # Get the size of the file from the source URL
        $webRequest = [System.Net.WebRequest]::Create($sourceUrl)
        $webResponse = $webRequest.GetResponse()
        $sourceSize = $webResponse.ContentLength

        # Compare the sizes
        if ($existingSize -eq $sourceSize) {
            Write-Host "$msixPath already exists and its size matches the source."
        } else {
            Write-Host "$msixPath already exists but its size does not match the source. Re-downloading..."
            try {
                Remove-Item $msixPath
                write-host "Removed $msixPath"
            } catch {
                Write-Host "Failed to remove $msixPath."
            }
        }
    } else {
        try {
            Write-Host "Downloading the Modern Teams MSIX file from $sourceUrl"
            Start-BitsTransfer -Source $sourceUrl -Destination $msixPath
            Write-Host "Download successful."
        } catch {
            Write-Host "Failed to download the Modern Teams MSIX file."
        }
    }

    # Install the new version of MS Teams
    try {
        Write-Host "Installing Modern Teams"
        Dism /Online /Add-ProvisionedAppxPackage /PackagePath:$msixPath /SkipLicense
    } catch {
        Write-Host "Failed to add provisioned Appx package: $_"
    }
}

if($TeamsNew){
    Write-Host "Modern Teams is already installed"  
    exit 0
} else {
    Write-Host "Modern Teams is not installed"
    try {
        Install-MSTEAMS
    } catch {
        Write-Host "Failed to install Modern Teams"
    }
}


# Check if Teams Classic is installed

if($TeamsClassic){
function Uninstall-TeamsClassic($TeamsPath) {
    try {
        $process = Start-Process -FilePath "$TeamsPath\Update.exe" -ArgumentList "--uninstall /s" -PassThru -Wait -ErrorAction STOP

        if ($process.ExitCode -ne 0) {
            Write-Error "Uninstallation failed with exit code $($process.ExitCode)."
        }
    } catch {
        Write-Error $_.Exception.Message
    }
}

# Remove Teams Machine-Wide Installer
Write-Host "Removing Teams Machine-wide Installer"

#Windows Uninstaller Registry Path
$registryPath = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"

# Get all subkeys and match the subkey that contains "Teams Machine-Wide Installer" DisplayName.
$MachineWide = Get-ItemProperty -Path $registryPath | Where-Object -Property DisplayName -eq "Teams Machine-Wide Installer"

if ($MachineWide) {
    Start-Process -FilePath "msiexec.exe" -ArgumentList "/x ""$($MachineWide.PSChildName)"" /qn" -NoNewWindow -Wait
} else {
    Write-Host "Teams Machine-Wide Installer not found"
}

# Get all Users
$AllUsers = Get-ChildItem -Path "$($ENV:SystemDrive)\Users"

# Process all Users
foreach ($User in $AllUsers) {
    Write-Host "Processing user: $($User.Name)"

    # Locate installation folder
    $localAppData = "$($ENV:SystemDrive)\Users\$($User.Name)\AppData\Local\Microsoft\Teams"
    $programData = "$($env:ProgramData)\$($User.Name)\Microsoft\Teams"

    if (Test-Path "$localAppData\Current\Teams.exe") {
        Write-Host "  Uninstall Teams for user $($User.Name)"
        Uninstall-TeamsClassic -TeamsPath $localAppData
    } elseif (Test-Path "$programData\Current\Teams.exe") {
        Write-Host "  Uninstall Teams for user $($User.Name)"
        Uninstall-TeamsClassic -TeamsPath $programData
    } else {
        Write-Host "  Teams installation not found for user $($User.Name)"
    }
}



# Remove old Teams folders and icons
$TeamsFolder_old = "$($ENV:SystemDrive)\Users\*\AppData\Local\Microsoft\Teams"
$TeamsIcon_old = "$($ENV:SystemDrive)\Users\*\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Microsoft Teams*.lnk"
Get-Item $TeamsFolder_old | Remove-Item -Force -Recurse
Get-Item $TeamsIcon_old | Remove-Item -Force -Recurse

} else {
    Write-Host "Teams Classic is not installed"
}

