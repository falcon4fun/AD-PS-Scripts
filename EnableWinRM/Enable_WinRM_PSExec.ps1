# Enables PSRemoting and ability to use PSExec

# Set all network adapters to Private
Get-NetConnectionProfile | Set-NetConnectionProfile -NetworkCategory Private

# Enable PSRemoting
Enable-PSRemoting –force
Set-NetFirewallRule -DisplayGroup “File And Printer Sharing” -Enabled True -Profile Private
