<#
.SYNOPSIS
    Upgrades Microsoft Teams to the latest version.

.DESCRIPTION
    This script checks if Microsoft Teams is already installed and if not, it installs the latest version of Microsoft Teams.
    It also checks if the classic version of Teams is installed and uninstalls it if found.

.PARAMETER None

.EXAMPLE
    Upgrade-Teams

.NOTES
    Author: 
    Date: 12/06/2024
#>

Set-Location "C:\Temp"

$isTeamsClassicInstalled = Test-Path C:\Users\*\AppData\Local\Microsoft\Teams\current\Teams.exe
#$isTeamsNewInstalled = Get-ChildItem "C:\Program Files\WindowsApps" -Filter "MSTeams_241*"

$sourceUrl = "https://statics.teams.cdn.office.net/production-windows-x64/enterprise/webview2/lkg/MSTeams-x64.msix"
$msixPath = "C:\Temp\MSTeams-x64.msix"

function Install-MSTeams {
    <#
    .SYNOPSIS
        Installs the latest version of Microsoft Teams.

    .DESCRIPTION
        This function enables the Windows Sideload apps feature, downloads the MSIX file for the latest version of Microsoft Teams,
        and installs it using the Dism command.

    .PARAMETER None

    .EXAMPLE
        Install-MSTeams
    #>

    # Check if the registry value exists and its data is 1
    $RegistryPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\AppModelUnlock"
    $RegistryValueName = "AllowAllTrustedApps"
    $RegistryValueData = 1

    if ((Get-ItemProperty -Path $RegistryPath -Name $RegistryValueName -ErrorAction SilentlyContinue) -and ((Get-ItemProperty -Path $RegistryPath -Name $RegistryValueName).$RegistryValueName -eq $RegistryValueData)) {
        Write-Output "Windows Sideload apps feature is already enabled."
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
                Write-Output "Windows Sideload apps feature has been enabled."
            } Catch {
                Write-Output "Failed to enable Windows Sideload apps feature."
            }
        } else {
            Write-Output "Windows Sideload apps feature is enabled."
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
            Write-Output "$msixPath already exists and its size matches the source."
        } else {
            Write-Output "$msixPath already exists but its size does not match the source. Re-downloading..."
            try {
                Remove-Item $msixPath -ErrorAction SilentlyContinue
                Write-Output "Removed $msixPath"
            } catch {
                Write-Output "Failed to remove $msixPath."
            }
        }
    } else {
        try {
            Write-Output "Downloading the Modern Teams MSIX file from $sourceUrl"
            Start-BitsTransfer -Source $sourceUrl -Destination $msixPath
            Write-Output "Download successful."
        } catch {
            Write-Output "Failed to download the Modern Teams MSIX file."
        }
    }

    # Install the new version of MS Teams
    try {
        Write-Output "Installing Modern Teams"
        Dism /Online /Add-ProvisionedAppxPackage /PackagePath:$msixPath /SkipLicense
    } catch {
        Write-Output "Failed to add provisioned Appx package: $_"
    }
}


<#if ($isTeamsNewInstalled) {
    Write-Output "Modern Teams is already installed"
    
} else {
    Write-Output "Modern Teams is not installed"
 #>   try {
        Install-MSTeams
        Write-Output "Modern Teams has been installed successfully."
    } catch {
        Write-Output "Failed to install Modern Teams"
    }
#}

# Check if Teams Classic is installed
if ($isTeamsClassicInstalled) {
    function Uninstall-TeamsClassic($TeamsPath) {
        <#
        .SYNOPSIS
            Uninstalls the classic version of Microsoft Teams.

        .DESCRIPTION
            This function uninstalls the classic version of Microsoft Teams by running the Update.exe with the "--uninstall /s" argument.

        .PARAMETER $TeamsPath
            The path to the Teams installation directory.

        .EXAMPLE
            Uninstall-TeamsClassic "C:\Program Files (x86)\Microsoft\Teams"

        #>

        try {
            $process = Start-Process -FilePath "$TeamsPath\Update.exe" -ArgumentList "--uninstall /s" -PassThru -Wait -ErrorAction SilentlyContinue

            if ($process.ExitCode -ne 0) {
                Write-Output "Uninstallation failed with exit code $($process.ExitCode)."
            }
        } catch {
            Write-Output $_.Exception.Message
        }
    }

    # Remove Teams Machine-Wide Installer
    Write-Output "Removing Teams Machine-wide Installer"

    # Windows Uninstaller Registry Path
    $registryPath = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"

    # Get all subkeys and match the subkey that contains "Teams Machine-Wide Installer" DisplayName.
    $MachineWide = Get-ItemProperty -Path $registryPath -ErrorAction SilentlyContinue | Where-Object -Property DisplayName -eq "Teams Machine-Wide Installer"

    if ($MachineWide) {
        Start-Process -FilePath "msiexec.exe" -ArgumentList "/x ""$($MachineWide.PSChildName)"" /qn" -NoNewWindow -Wait -ErrorAction SilentlyContinue
    } else {
        Write-Output "Teams Machine-Wide Installer not found"
    }

    # Get all Users
    $AllUsers = Get-ChildItem -Path "$($ENV:SystemDrive)\Users" -ErrorAction SilentlyContinue

    # Process all Users
    foreach ($User in $AllUsers) {
        $programData = Join-Path $env:ProgramData $User.Name 
        $programData = Join-Path $programData "Microsoft\Teams"
        
        $localAppData = Join-Path $env:LocalAppData $User.Name 
        $localAppData = Join-Path $localAppData "Microsoft\Teams"

        Uninstall-TeamsClassic $localAppData
        Uninstall-TeamsClassic $programData
    }

    # Remove old Teams folders and icons
    $TeamsFolder_old = Join-Path $ENV:SystemDrive "Users\*\AppData\Local\Microsoft\Teams"
    $TeamsIcon_old = Join-Path $ENV:SystemDrive "Users\*\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Microsoft Teams*.lnk"
    $TeamsShortcut_old = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Teams.lnk"
    $TeamsShortcut_PublicDesktop = "C:\Users\Public\Desktop\Microsoft Teams.lnk"
    Get-Item -Path $TeamsShortcut_PublicDesktop -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
    Get-Item -Path $TeamsShortcut_old -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
    Get-Item -Path $TeamsFolder_old -ErrorAction SilentlyContinue | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
    Get-Item -Path $TeamsIcon_old -ErrorAction SilentlyContinue | Remove-Item -Force -Recurse 

} else {
    Write-Output "Teams Classic is not installed"
}
