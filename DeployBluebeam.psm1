#Invoke-Expression(New-Object Net.WebClient).DownloadString('https://raw.githubusercontent.com/integrateitaus/Public/main/DeployBluebeam.psm1'); Install-Bluebeam -serial "1234567" -product "12345-1234567"

$LogPath = "C:\Support\$env:computername-AppInstall.log"

function Get-LatestVersion {
    # URL of the webpage
    $url = "https://www.bluebeam.com/download"

    # Download the webpage
    $page = Invoke-WebRequest -Uri $url

    # Regular expression pattern to match URLs
    $pattern = 'https://downloads.bluebeam.com/software/downloads/20\.[^"]*'

    # Find all matches
    $match = $page.Content | Select-String -Pattern $pattern -AllMatches

    # Output the URLs
    $downloadURL = $match.Matches.Value

    # Remove the specified text from the start of the string
    $newString = $downloadURL -replace "^https://downloads.bluebeam.com/software/downloads/", ""

    # Remove everything after "/BbRevu20." including "/BbRevu20."
    $newString = $newString -replace "/BbRevu20..*", ""

    # Output the new string
    return $newString
}

# Install Bluebeam version 20.3.20 for all users
function Install-Bluebeam {
    Param
    (
        #[Parameter(Mandatory=$true)]
        #[string] $version,
        [Parameter(Mandatory=$true)]
        [string] $serial,
        [Parameter(Mandatory=$true)]
        [string] $product
    )

    # Get latest version
    $version = Get-LatestVersion
    Write-Output "Returned latest Bluebeam Version: $version"
    Add-Content -Path $LogPath -Value "Returned latest Bluebeam Version: $version"

    # Install Bluebeam
    $installed = Get-WmiObject -Query "SELECT * FROM Win32_Product WHERE (Name LIKE 'Bluebeam Revu%')" | Where-Object { $_.Version -ge $version }
    if (!$installed) {

        # Variables
        #$LogPath = "C:\Support\$env:computername-AppInstall.log"
        $Bluebeam_zip_url = "https://downloads.bluebeam.com/software/downloads/" + "$version" + "/MSIBluebeamRevu" + "$version" + "x64.zip"
        $Bluebeam_zip = "C:\Support\MSIBluebeamRevu" + "$version" + "x64.zip"
        $Bluebeam_extracted_folder = "C:\Support\MSIBluebeamRevu" + "$version"
        $Bluebeam_msi = "C:\Support\MSIBluebeamRevu" + "$version" + "\Bluebeam Revu x64 20.msi"

        $file = "Bluebeam Revu x64 20.msi"

        # Create the required folders if they don't exist
        if (!(Test-Path -Path "C:\Support")) {
            New-Item -ItemType Directory -Path "C:\Support"
        }
        if (!(Test-Path -Path $Bluebeam_extracted_folder)) {
            New-Item -ItemType Directory -Path $Bluebeam_extracted_folder
        }

        # Download Bluebeam
        try { 
            Write-Output "Downloading Bluebeam Installer"
            Start-BitsTransfer -Source $Bluebeam_zip_url -Destination $Bluebeam_zip
        } catch {
            Write-Output "Error Downloading Bluebeam: $_"
            Add-Content -Path $LogPath -Value "Error Downloading Bluebeam: $_"
        } 

        # Extract Bluebeam
        try { 
            Write-Output "Extracting Bluebeam"
            Add-Type -AssemblyName System.IO.Compression.FileSystem

            [System.IO.Compression.ZipFile]::OpenRead($Bluebeam_zip).Entries |
            Where-Object { $_.FullName -eq $file } |
            ForEach-Object { [System.IO.Compression.ZipFileExtensions]::ExtractToFile($_, "$Bluebeam_extracted_folder\$file", $true) }
        } catch {
            Write-Output "Error Extracting Bluebeam: $_"
            Add-Content -Path $LogPath -Value "Error Extracting Bluebeam: $_"
        } 


        try { 
            #Change user /install
            Write-Output "Installing Bluebeam"

            Start-Process msiexec.exe -ArgumentList "/i `"$Bluebeam_msi`" BB_SERIALNUMBER=$serial BB_PRODUCTKEY=$product /qn" -Wait
                        
            $installed = Get-WmiObject -Query "SELECT * FROM Win32_Product WHERE (Name LIKE 'Bluebeam Revu%'')" | Where-Object { $_.Version -ge $version }
            if ($installed) {
                Write-Output "Bluebeam installed successfully"
                Add-Content -Path $LogPath -Value "Bluebeam installed successfully: $_"
            } else {
                Write-Output "Bluebeam installed failed"
                Add-Content -Path $LogPath -Value "Bluebeam installed failed: $_"

                # attempt to uninstall old version
                # Get the installed application
                $app = Get-WmiObject -Query "SELECT * FROM Win32_Product WHERE (Name LIKE 'Bluebeam%')"

                # Check if the application exists
                if ($app -ne $null) {
                    Write-Output "Removing old version of Bluebeam"
                    # Uninstall the application
                    $app.Uninstall()
                    Write-Output "Uninstallation completed."

                    Write-Output "Second attempt at Installing Bluebeam"
                    Start-Process msiexec.exe -ArgumentList "/i `"$Bluebeam_msi`" BB_SERIALNUMBER=$serial BB_PRODUCTKEY=$product /qn" -Wait Start-Process msiexec.exe -ArgumentList "/i `"$Bluebeam_msi`" BB_SERIALNUMBER=$serial BB_PRODUCTKEY=$product /qn" -Wait
                    $installed = Get-WmiObject -Query "SELECT * FROM Win32_Product WHERE (Name LIKE 'Bluebeam Revu%'')" | Where-Object { $_.Version -ge $version }
                    if ($installed) {
                        Write-Output "Bluebeam installed successfully on second attempt."
                        Add-Content -Path $LogPath -Value "Bluebeam installed successfully on second attempt: $_"
                    } else {
                        Write-Output "Bluebeam installed failed on second attempt. Manual investigation required."
                        Add-Content -Path $LogPath -Value "Bluebeam installed failed on second attempt: $_"
                    }

                } else {
                    Write-Output "No existing Bluebeam version is not installed. Manual investigation required."
                }
            }

        } catch {
            Write-Output "Error installing Bluebeam: $_"
            Add-Content -Path $LogPath -Value "Error installing Bluebeam: $_"
        }

    } else {
        Write-Output "Bluebeam Version or newer is already installed"
        Add-Content -Path $LogPath -Value "Bluebeam $version or newer is already installed."
    }
}


