# Prints all printers on devices

# Local admin credential with 1\ or any other fake computername
$creds = Get-Credential -UserName 1\Administrator -Message 1

$Computer = "Computer1.domain.lan", "Computer2.domain.lan", "Computer3.domain.lan", "Computer4.domain.lan"

$Computer | ForEach-Object {
    If (Test-Connection $_ -Count 1 -ErrorAction SilentlyContinue) {
        New-PSSession $_ -Credential $creds
    }
}

$Sessions = Get-PSSession
Invoke-Command -Session $Sessions -ScriptBlock { hostname; $printers = Get-Printer; $printers | Format-Table Name, DriverName, PortName }
Remove-PSSession -Session $Sessions
