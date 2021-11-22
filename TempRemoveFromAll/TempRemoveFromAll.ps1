# Temp REMOVER

# Removes Temp, Recycle.bins, Caches from All AD servers.
# Folders defined in Dirs1 and Dirs2.
# By default it only calculates storage size without removing.
# To remove simply uncomment "Remove-Item" lines

Start-Transcript -path output.txt

$global:TotalOfAll = 0
$global:TotalOfRecycle = 0
$global:TotalOfWinTemp = 0

#W2008 & W2012
$Dirs1 = @(
'\AppData\Local\Microsoft\Windows\Temporary Internet Files\Content.IE5'
'\AppData\Local\Microsoft\Windows\Temporary Internet Files\Content.MSO'
'\AppData\Local\Microsoft\Windows\Temporary Internet Files\Content.Outlook'
'\AppData\Local\Microsoft\Windows\Temporary Internet Files\Content.Word'
'\AppData\Local\Microsoft\Windows\Temporary Internet Files\Low'
'\AppData\Local\Google\Chrome\User Data\Default\Cache'
'\AppData\Local\Google\Chrome\User Data\Default\Code Cache'
'\AppData\LocalLow\Adobe\Acrobat\DC\ConnectorIcons'
'\AppData\Local\Temp'
)

$Dirs2 = @(
'\AppData\Local\Microsoft\Windows\INetCache\IE'
'\AppData\Local\Microsoft\Windows\INetCache\Content.MSO'
'\AppData\Local\Microsoft\Windows\INetCache\Content.Outlook'
'\AppData\Local\Microsoft\Windows\INetCache\Content.Word'
'\AppData\Local\Microsoft\Windows\INetCache\Low\IE'
'\AppData\Local\Google\Chrome\User Data\Default\Cache'
'\AppData\Local\Google\Chrome\User Data\Default\Code Cache'
'\AppData\LocalLow\Adobe\Acrobat\DC\ConnectorIcons'
'\AppData\Local\Temp'
)

function Get-DirectorySize
{

  param(
    [Parameter(ValueFromPipeline)] [Alias('PSPath')]
    [string] $LiteralPath = '.',
    [switch] $Recurse,
    [switch] $ExcludeSelf,
    [int] $Depth = -1,
    [int] $__ThisDepth = 0 # internal use only
  )

  process {

    # Resolve to a full filesystem path, if necessary
    $fullName = if ($__ThisDepth) { $LiteralPath } else { Convert-Path -ErrorAction SilentlyContinue -LiteralPath $LiteralPath }

    if ($ExcludeSelf) { # Exclude the input dir. itself; implies -Recurse

      $Recurse = $True
      $ExcludeSelf = $False

    } else { # Process this dir.

      # Calculate this dir's total logical size.
      # Note: [System.IO.DirectoryInfo].EnumerateFiles() would be faster, 
      # but cannot handle inaccessible directories.
      $size = [Linq.Enumerable]::Sum(
        [long[]] (Get-ChildItem -Force -Recurse -File -LiteralPath $fullName).ForEach('Length')
      )

      # Create a friendly representation of the size.
      $decimalPlaces = 2
      $padWidth = 8
      $scaledSize = switch ([double] $size) {
        {$_ -ge 1tb } { $_ / 1tb; $suffix='tb'; break }
        {$_ -ge 1gb } { $_ / 1gb; $suffix='gb'; break }
        {$_ -ge 1mb } { $_ / 1mb; $suffix='mb'; break }
        {$_ -ge 1kb } { $_ / 1kb; $suffix='kb'; break }
        default       { $_; $suffix='b'; $decimalPlaces = 0; break }
      }

      # Construct and output an object representing the dir. at hand.
      [pscustomobject] @{
        FullName = $fullName
        FriendlySize = ("{0:N${decimalPlaces}}${suffix}" -f $scaledSize).PadLeft($padWidth, ' ')
        Size = $size
      }

    }

    # Recurse, if requested.
    if ($Recurse -or $Depth -ge 1) {
      if ($Depth -lt 0 -or (++$__ThisDepth) -le $Depth) {
        # Note: This top-down recursion is inefficient, because any given directory's
        #       subtree is processed in full.
        Get-ChildItem -Force -Directory -LiteralPath $fullName |
          ForEach-Object { Get-DirectorySize -LiteralPath $_.FullName -Recurse -Depth $Depth -__ThisDepth $__ThisDepth }
      }
    }

  }

}

Function Remove-Folders($computerName, $Dirs) {

    #$computerName | ft

    $Total = 0
    $Profiles = Get-ChildItem "\\$computerName\C$\Users"

    
    ForEach ($Profile in $Profiles) {
        #Write-Host $Profile
        $ProfileSize = 0
        ForEach ($Dir in $Dirs) {
            $str = "\\$computerName\C$\Users\"+$Profile+$Dir+"\*"
            $str
            $DirSize = Get-DirectorySize -LiteralPath $str.Substring(0,$str.Length-1) -ErrorAction SilentlyContinue
            $DirSize
            $Total += $DirSize.Size
            #Remove-Item $str -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    Write-Host "`nProfileDir: "
    $Total / 1MB
   

    $Recycle = "\\$computerName\C$\`$Recycle.Bin\*"
    $Recycle | fl
    $DirSize = Get-DirectorySize -LiteralPath $Recycle.Substring(0,$Recycle.Length-1) -ErrorAction SilentlyContinue
    Write-Host "`nRecycle: "
    $($DirSize.Size / 1MB)
    $Total += $DirSize.Size
    $global:TotalOfRecycle += $DirSize.Size
    #Remove-Item $Recycle -Recurse -Force -ErrorAction SilentlyContinue

    $TempDir = "\\$computerName\C$\Windows\Temp\*"
    $TempDir | fl
    $DirSize1 = Get-DirectorySize -LiteralPath $TempDir.Substring(0,$TempDir.Length-1) -ErrorAction SilentlyContinue
    Write-Host "`nTempDir: "
    $($DirSize1.Size / 1MB)
    $Total += $DirSize1.Size
    $global:TotalOfWinTemp += $DirSize1.Size
    #Remove-Item $TempDir -Recurse -Force -ErrorAction SilentlyContinue

    Write-Host "`nTotal: "
    Write-Host $Total
    Write-Host $($Total / 1MB) " MB"
    

    $global:TotalOfAll += $Total
    Write-Host "TotalOfAll: $global:TotalOfAll"
    Write-Host "****************************************"
    Write-Host "****************************************"
    Write-Host "****************************************"
    Write-Host "****************************************"
}


$computers = Get-ADComputer -Filter { OperatingSystem -Like '*Windows Server*' -and Enabled -eq $True } -Properties OperatingSystem | Sort-Object

foreach ($computer in $computers.name)
{
Write-Host $computer
    if (Test-Connection -ComputerName $computer -Count 2 -Quiet)
    {
        $ThisVer = (Get-WmiObject Win32_OperatingSystem -ComputerName $computer -Errorvariable err -ErrorAction SilentlyContinue).Version
        #Write-Host "Err: $err"

        if (!$err) {
            #Write-Host "ThisVer: $ThisVer"
            if ($ThisVer -as [version] -ge (new-object 'Version' 10,0)) {
                #Write-Host "10.0"
                Remove-Folders $computer $Dirs2
            } elseif ($ThisVer -as [version] -ge (new-object 'Version'6,3)) {
                #Write-Host "6.3"
                Remove-Folders $computer $Dirs1
            } elseif ($ThisVer -as [version] -ge (new-object 'Version'6,1)) {
                #Write-Host "6,1"
                Remove-Folders $computer $Dirs1
            } else {
                Write-Host "Error"
            }
        } else {
            Write-Host "Err: $err"
            Write-Host "****************************************"
            Write-Host "****************************************"
            Write-Host "****************************************"
            Write-Host "****************************************"
        }
    } else {
        Write-Host "No Ping"
        Write-Host "****************************************"
        Write-Host "****************************************"
        Write-Host "****************************************"
        Write-Host "****************************************"
    }
}

Write-Host "`nTotalOfAll: "
Write-Host $global:TotalOfAll
Write-Host $($global:TotalOfAll / 1MB) " MB"

Write-Host "`nTotalOfRecycle: "
Write-Host $global:TotalOfRecycle
Write-Host $($global:TotalOfRecycle / 1MB) " MB"

Write-Host "`nTotalOfWinTemp: "
Write-Host $global:TotalOfWinTemp
Write-Host $($global:TotalOfWinTemp / 1MB) " MB"

Stop-Transcript