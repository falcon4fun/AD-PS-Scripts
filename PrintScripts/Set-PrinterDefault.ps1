# Tryes to set default printer on remote machine
# Don't remember if it will work normally


$creds = Get-Credential -UserName 1\Administrator -Message 1

$Computer = "Computer1.domain.lan", "Computer2.domain.lan", "Computer3.domain.lan", "Computer4.domain.lan"

$Computer | ForEach-Object {
    If (Test-Connection $_ -Count 1 -ErrorAction SilentlyContinue) {
        New-PSSession $_ -Credential $creds
    }
}

$Sessions = Get-PSSession
$scriptBlock = {
    hostname;
    Get-CimInstance -ClassName CIM_Printer | Select-Object Name, SystemName, Default | Format-Table;
}
Invoke-Command -Session $Sessions -ScriptBlock $scriptBlock

$scriptBlock = {
    $printer = Get-CimInstance -ClassName Win32_Printer -Filter "name LIKE 'Canon Office Printer'";
    Invoke-CimMethod -InputObject $printer -MethodName SetDefaultPrinter;
}
Invoke-Command -Session $Sessions -ScriptBlock $scriptBlock

$scriptBlock = {
    hostname;
        (New-Object -ComObject WScript.Network).SetDefaultPrinter('Canon Office Printer');
}
Invoke-Command -Session $Sessions -ScriptBlock $scriptBlock

Remove-PSSession -Session $Sessions
