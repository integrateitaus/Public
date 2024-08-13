$regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings"
$regName = "EnableNegotiate"
$regValue = 0

Set-ItemProperty -Path $regPath -Name $regName -Value $regValue

$adapters = Get-NetAdapter | Where-Object { $_.Status -eq 'Up' }

foreach ($adapter in $adapters) {
    Set-NetConnectionProfile -InterfaceAlias $adapter.Name -NetworkCategory Private
}

winrm quickconfig -force


Set-NetFirewallProfile -Profile Domain,Public,Private -Enabled False

