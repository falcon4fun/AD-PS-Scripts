# Reg2CI (c) 2022 by Roger Zander
if ((Test-Path -LiteralPath "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp") -ne $true) {
    New-Item "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp" -force -ea SilentlyContinue 
};

New-ItemProperty -LiteralPath 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp' -Name 'UserAuthentication' -Value 1 -PropertyType DWord -Force -ea SilentlyContinue;
