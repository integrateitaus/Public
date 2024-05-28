function ChangeLicenseKey {
    param (
        [Parameter(Mandatory=$true)]
        [string]$NewLicenseKey
    )

    try {
        # Change the license key
        $Result = slmgr /ipk $NewLicenseKey

        # Check if the license key change was successful
        if ($Result -match "successfully installed") {
            # Upgrade from Home to Pro
            $UpgradeResult = slmgr /upk
            if ($UpgradeResult -match "successfully uninstalled") {
                $UpgradeResult = slmgr /ipk $NewLicenseKey
                if ($UpgradeResult -match "successfully installed") {
                    # Activate the new license key
                    $ActivationResult = slmgr /ato
                    if ($ActivationResult -match "successfully activated") {
                        Write-Host "License key changed and upgraded to Windows Pro successfully."
                    } else {
                        Write-Host "Failed to activate the new license key."
                    }
                } else {
                    Write-Host "Failed to install the new license key."
                }
            } else {
                Write-Host "Failed to uninstall the existing license key."
            }
        } else {
            Write-Host "Failed to change the license key."
        }
    } catch {
        Write-Host "An error occurred: $_"
    }
}

# Call the function with the new license key
$NewLicenseKey = "XXXXX-XXXXX-XXXXX-XXXXX-XXXXX"
ChangeLicenseKey -NewLicenseKey $NewLicenseKey
