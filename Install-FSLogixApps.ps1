<#
.SYNOPSIS
    Install FSLogix Apps if not installed or if the installed version is older than the current version.

.DESCRIPTION
    This script checks the current version of FSLogix based on the short URL redirected filename and installs FSLogix Apps if it is not installed or if the installed version is older than the current version. The script uses [System.Version] object type cast for version comparison to ensure any major version numbers are accounted for. It also extracts the 64-bit FSLogix Apps installer exe from the downloaded zip.

    Caveats:
    - If Microsoft changes the short URL, the redirection, or the path/filename within the zip, the script will break.
    - It is strongly recommended not to autorun this script due to unexpected bugs in FSLogix.

.PARAMETER None

.EXAMPLE
    Install-FSLogixApps

.NOTES

#>

$FSLogixURL = "https://aka.ms/fslogix_download"
$FSLogixDownload = "FSLogixSetup.zip"
$FSLogixInstaller = "FSLogixAppsSetup.exe"
$ZipFileToExtract = "x64/Release/FSLogixAppsSetup.exe"
$Zip = "$env:temp\$FSLogixDownload"
$Installer = "$env:temp\$FSLogixInstaller"
$downloadAndInstall = $false

$ProductName = "Microsoft FSLogix Apps"

Write-Host "Checking registry for $ProductName"

# Get FSLogix version number if installed
$fslogixsearch = (get-wmiobject Win32_Product | where-object name -eq "Microsoft FSLogix Apps" | select-object Version)

switch ($fslogixsearch.count) {
    0 {
        # Not found
        $fslogixver = $null
        $downloadAndInstall = $true
    }
    1 {
        # One entry returned
        $fslogixver = [System.Version]$fslogixsearch.Version
        Write-Host "FSLogix version installed: $fslogixver"
    }
    {$_ -gt 1} {
        # two or more returned
        $fslogixver = [System.Version]$fslogixsearch[0].Version
        Write-Host "FSLogix version installed: $fslogixver"
    }

}

# Find current FSLogix version from short URL:
$WebRequest = [System.Net.WebRequest]::create($FSLogixURL)
$WebResponse = $WebRequest.GetResponse()
$ActualDownloadURL = $WebResponse.ResponseUri.AbsoluteUri
$WebResponse.Close()

$FSLogixCurrentVersion = [System.Version]((Split-Path $ActualDownloadURL -leaf).Split("_")[2]).Replace(".zip","")



#$FSLogixCurrentVersion = "2.9.8884.27471"
#$FSLogixCurrentVersion
Write-Host "Current FSLogix version: $FSLogixCurrentVersion"

# See if the current version is newer than the installed version:
if ($FSLogixCurrentVersion -gt $fslogixver) {
    # Current version greater than installed version, install new version
    Write-Host "New version will be downloaded and installed. ($FSLogixCurrentVersion > $fslogixver)"
    $downloadAndInstall = $true
}

# If $downloadAndInstall has been toggled true, download and install.
if ($downloadAndInstall)
{
    Write-Host "Not installed... beginning install..."
    # Download installer
    Import-Module BitsTransfer
    Write-Host "Downloading from: $FSLogixURL"
    Write-Host "Saving file to: $Zip"


    try { 
        Write-Output "Downloading FSLogix Installer"
        
        Start-BitsTransfer -Source $FSLogixURL -Destination "$env:temp\$FSLogixDownload" -ErrorAction Stop
    } catch {
        $errorMessage = "Error Downloading FSLogix: $_"
        Write-Output $errorMessage
    } 

    # Extract file from zip: x64\Release\FSLogixAppsSetup.exe to $env:temp\FSLogixAppsSetup.exe
    
    # Open zip
    Add-Type -Assembly System.IO.Compression.FileSystem
    $zipFile = [IO.Compression.ZipFile]::OpenRead($Zip)
    
    # Retrieve the $ZipFileToExtract and extract to $Installer
    $filetoextract = ($zipFile.Entries | Where-Object {$_.FullName -eq $ZipFileToExtract})
    [System.IO.Compression.ZipFileExtensions]::ExtractToFile($filetoextract[0], $Installer, $true)
    
    # Run installer
    Write-Host "Running $Installer /install /quiet /norestart"
    Start-Process $Installer -wait -ArgumentList "/install /quiet /norestart"

    <# # Wait for 5 minutes so that the files can be deleted because despite -wait being specified, it doesn't actually wait for all processes to finish
    Start-Sleep -Seconds 300

    # Close the zip file so it can be deleted
    $zipFile.Dispose()

    # Clean up
    Write-Host "Cleaning up, deleting $Installer and $Zip."
    Remove-Item -Path $Installer -Force
    Remove-Item -Path $Zip -Force
    #>
} else {
    Write-Host "FSLogix already installed and up to date."

}