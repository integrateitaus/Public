<#
.SYNOPSIS
    Uninstalls Microsoft Teams application packages for all users and the current user.

.DESCRIPTION
    This script uninstalls Microsoft Teams application packages from the system. It removes both the AppPackage and AppxPackage versions of Microsoft Teams.

    The script performs the following steps:
    1. Uninstalls Microsoft Teams AppPackage for all users.
    2. Uninstalls Microsoft Teams AppPackage for the current user.
    3. Uninstalls Microsoft Teams AppxPackage for all users.
    4. Uninstalls Microsoft Teams AppxPackage for the current user.
    5. Uninstalls Microsoft Teams package using DISM (Deployment Image Servicing and Management) for the current user.
    6. Displays a list of all provisioned apps.

.NOTES
    - This script requires administrative privileges to uninstall the application packages.
    - The script may display errors if the application packages are not found or if there are permission issues.

.EXAMPLE
    Uninstall-TeamsModern.ps1
    Uninstalls Microsoft Teams application packages for all users and the current user.

#>

try {
    Write-Host "Uninstalling Microsoft Teams AppPackage for All Users..."
    Get-AppPackage *teams* -AllUsers | Remove-AppPackage -AllUsers
} catch {
    Write-Host "Failed to uninstall Microsoft Teams AppPackage for All Users. Error: $_"
}

try {
    Write-Host "Uninstalling Microsoft Teams AppPackage for Current User..."
    Get-AppPackage *teams* | Remove-AppPackage
} catch {
    Write-Host "Failed to uninstall Microsoft Teams AppPackage for Current User. Error: $_"
}

try {
    Write-Host "Uninstalling Microsoft Teams AppxPackage for All Users..."
    Get-AppxPackage -AllUsers *teams* | Remove-AppxPackage -AllUsers
    Get-AppxPackage -allusers MSTeams | Remove-AppxPackage -AllUsers
} catch {
    Write-Host "Failed to uninstall Microsoft Teams AppxPackage for All Users. Error: $_"
}

try {
    Write-Host "Uninstalling Microsoft Teams AppxPackage for Current User..."
    Get-AppxPackage *teams* | Remove-AppxPackage
    Get-AppxPackage MSTeams | Remove-AppxPackage
} catch {
    Write-Host "Failed to uninstall Microsoft Teams AppxPackage for Current User. Error: $_"
}

try {
    $packageName = "MicrosoftTeams"
    $package = Get-AppxPackage -Name $packageName -ErrorAction SilentlyContinue

    if ($package) {
        Write-Host "Uninstalling Teams package using DISM..."
        $result = dism.exe /Online /Remove-ProvisionedAppxPackage /PackageName:$package.PackageFullName /PackageFamilyName:$package.PackageFamilyName
        if ($result -eq 0) {
            Write-Host "Teams package uninstalled successfully."
        } else {
            Write-Host "Failed to uninstall Teams package. Error code: $result"
        }
    } else {
        Write-Host "Teams package is not installed."
    }
} catch {
    Write-Host "Failed to uninstall Teams package. Error: $_"
}

Write-Host "List of all provisioned apps:"
Get-ProvisionedAppxPackage -Online | Select-Object DisplayName, PackageName, PackageFamilyName, PackageFullName
