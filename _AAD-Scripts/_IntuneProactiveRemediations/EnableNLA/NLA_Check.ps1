try {
    if (-NOT (Test-Path -LiteralPath "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp")) { exit 1 };
    if ((Get-ItemPropertyValue -LiteralPath 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp' -Name 'UserAuthentication' -ea SilentlyContinue) -eq 1) { 
    } else { 
        exit 1 
    };
} catch { 
    exit 1 
}
exit 0
