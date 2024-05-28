<#
.SYNOPSIS
This script uninstalls a list of Dell packages and Dell Optimizer.

.DESCRIPTION
The script iterates through an array of Dell package names and attempts to uninstall each package using the Get-Package and Uninstall-Package cmdlets. If an error occurs during the uninstallation process, the script falls back to an alternate uninstallation method for Dell Optimizer.

.PARAMETER None

.EXAMPLE
.\Uninstall-DellOEMApps.ps1
Runs the script to uninstall the specified Dell packages and Dell Optimizer.

.NOTES
Author: Phillip Anderson
Date: 28/05/2024
#>

function Uninstall-DellOEMApps {

    $packages = @(
        "*Dell Command*",
        "*Dell Power Manager Service*",
        "*Dell Mobile Connect*",
        "*Dell Update*",
        "Dell Optimizer*",
        "Dell SupportAssist Remediation*",
        "Dell SupportAssist OS Recovery*",
        "Clipchamp.Clipchamp",
        "DellInc.DellDigitalDelivery",
        #"DellInc.DellSupportAssistforPCs",
        "DellInc.PartnerPromo",
        "Microsoft.GamingApp",
        "Microsoft.GetHelp",
        "Microsoft.Getstarted",
        "Microsoft.MicrosoftSolitaireCollection",
        "Microsoft.MixedReality.Portal",
        "Microsoft.SkypeApp",
        #"Microsoft.StorePurchaseApp",
        "Microsoft.Windows.DevHome",
        "Microsoft.WindowsFeedbackHub",
        "Microsoft.Xbox.TCUI",
        "Microsoft.XboxApp",
        "Microsoft.XboxGameOverlay",
        "Microsoft.XboxGamingOverlay",
        "Microsoft.XboxIdentityProvider",
        "Microsoft.XboxSpeechToTextOverlay",
        "Microsoft.ZuneMusic",
        "Microsoft.ZuneVideo",
        "MicrosoftTeams",
        "Microsoft.communicationsapps"
    )

    foreach ($package in $packages) {
        try {
            Write-Host "Attempting to uninstall $package"

            Get-Package -Name $package -ErrorAction SilentlyContinue | Uninstall-Package -Force -ErrorAction SilentlyContinue | Out-Null
            Get-AppxPackage -Name "*Dell Digital Delivery*" -AllUsers | Remove-AppxPackage -ErrorAction SilentlyContinue | Out-Null
            Get-AppxPackage -Name $package | Remove-AppxPackage -ErrorAction SilentlyContinue | Out-Null
            Get-AppXProvisionedPackage -Online | Where-Object DisplayName -EQ $package | Remove-AppxProvisionedPackage -Online  -ErrorAction SilentlyContinue | Out-Null

            Write-Host "Successfully uninstalled $package" -ForegroundColor Green

        } catch {
            Write-Host "Failed to uninstall $package $_" -ForegroundColor Red
        }
    }

    try {
        Write-Host "Uninstalling Dell Optimizer"
        Get-Package -Name "Dell Optimizer*" | Uninstall-Package -Force -ErrorAction SilentlyContinue
    }
    catch {
        Write-Host "An error occurred while uninstalling Dell Optimizer: $_ Attempting alternate uninstall method"
        # Remove Dell optimizer using registry and file system
        $unins = Get-ChildItem "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall" -ErrorAction SilentlyContinue | Get-ItemProperty | Where-Object {$_.DisplayName -Like "*Dell Optimizer*"} | Select-Object Displayname, Uninstallstring
        Start-Process -FilePath $unins.UninstallString -ArgumentList "/S" -Wait -ErrorAction SilentlyContinue

        if (Test-Path -Path "C:\Program Files (x86)\InstallShield Installation Information\{286A9ADE-A581-43E8-AA85-6F5D58C7DC88}\DellOptimizer.exe") {
            Invoke-Command -ScriptBlock { "C:\Program Files (x86)\InstallShield Installation Information\{286A9ADE-A581-43E8-AA85-6F5D58C7DC88}\DellOptimizer.exe" } -ArgumentList "-remove -runfromtemp"
        }
    }
}

