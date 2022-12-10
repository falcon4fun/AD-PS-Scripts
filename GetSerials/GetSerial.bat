@echo off
hostname >> "%~dp0\Serials.txt"
echo %username% >> "%~dp0\Serials.txt"
whoami >> "%~dp0\Serials.txt"
wmic baseboard get product,Manufacturer,version,serialnumber |more >> "%~dp0\Serials.txt"
wmic bios get serialnumber |more >> "%~dp0\Serials.txt"
getmac /fo csv /v >> "%~dp0\Serials.txt"
echo ---------------------------- >> "%~dp0\Serials.txt"
echo.  >> "%~dp0\Serials.txt"
