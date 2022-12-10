# Invoked from Install_Printer_PsRemote

# Installs new print driver for network print without using a printserver
# Mostly usable on non-domain (workgroup) machines

# Drivers should be on local share with anonymous access
# Otherwise use user and password hardcoded. 

# Local admin credential with 1\ or any other fake computername. Will not work with only username otherwise

$printAddr = "10.0.1.100"
$portName = "IP:10.0.1.100"
# Printer Driver Name can be extracted from INF file
$printDriverName = "Canon Generic Plus PCL6"
$printName = "Canon Office Printer"

$portExists = Get-Printerport -Name $portname -ErrorAction SilentlyContinue
if (-not $portExists) {
    Add-PrinterPort -Name $portName -PrinterHostAddress $printAddr
}

$printDriverExists = Get-PrinterDriver -name $printDriverName -ErrorAction SilentlyContinue
if ($printDriverExists) {
    Add-Printer -Name $printName -PortName $portname -DriverName $printDriverName
} else {
    net use S: \\share1.domain.lan\PrinterDriverShare /user:ADSAD
    pnputil.exe -i -a "S:\*.inf"
    Add-PrinterDriver -Name $printDriverName
    Add-Printer -Name $printName -PortName $portname -DriverName $printDriverName
    net use S: /delete
    net stop spooler
    net start spooler
}
