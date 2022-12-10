# Sets preferable Explorer configuration for Admin

$Hidden = (Get-ItemProperty -path HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced).Hidden
If ($Hidden -ne 1) {
    Set-ItemProperty -Path HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced -Name Hidden -Value "1" -Type DWord -Force
}

$ShowSuperHidden = (Get-ItemProperty -path HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced).ShowSuperHidden
If ($ShowSuperHidden -ne 1) {
    Set-ItemProperty -Path HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced -Name ShowSuperHidden -Value "1" -Type DWord -Force
}

$HideFileExt = (Get-ItemProperty -path HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced).HideFileExt
If ($HideFileExt -ne 0) {
    Set-ItemProperty -Path HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced -Name HideFileExt -Value "0" -Type DWord -Force
}

$PersistBrowsers = (Get-ItemProperty -path HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced).PersistBrowsers
If ($PersistBrowsers -ne 1) {
    Set-ItemProperty -Path HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced -Name PersistBrowsers -Value "1" -Type DWord -Force
}

$LaunchTo = (Get-ItemProperty -path HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced).LaunchTo
If ($LaunchTo -ne 1) {
    Set-ItemProperty -Path HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced -Name LaunchTo -Value "1" -Type DWord -Force
}
