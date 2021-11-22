set /p UserName=Enter samAccountName:
runas /noprofile /user:trans\%UserName% "notepad"
pause