# Installs new print driver for network print without using a printserver
# Mostly usable on non-domain (workgroup) machines

# Local admin credential with 1\ or any other fake computername
$creds = Get-Credential -UserName 1\Administrator -Message 1

$Computer = "Computer1.domain.lan", "Computer2.domain.lan", "Computer3.domain.lan", "Computer4.domain.lan"

$Computer | ForEach-Object {
    If (Test-Connection $_ -Count 1 -ErrorAction SilentlyContinue) {
        New-PSSession $_ -Credential $creds
    }
}

$Sessions = Get-PSSession
Invoke-Command -FilePath C:\Temp\Install_Printer_PsRemote_Child.ps1 -Session $Sessions
Invoke-Command -Session $Sessions -ScriptBlock { hostname; Remove-Printer -ErrorAction SilentlyContinue -Name "iR-ADV C255", "Canon.iR-ADV C255/355 UFR II"; }
Remove-PSSession -Session $Sessions