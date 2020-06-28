### Module - zTools

function tPing {
    <#
    .NOTES
        ######################
         mail@nimbus117.co.uk
        ######################
    
    .SYNOPSIS

        C:\PS> tPing google.com

        TimeStamp            Destination  Response  Time
        ---------            -----------  --------  ----
        03/04/2015 22:35:28  google.com       True    18
        03/04/2015 22:35:29  google.com       True    18
        03/04/2015 22:35:30  google.com       True    19
        03/04/2015 22:35:31  google.com       True    29

    .PARAMETER Destination
        IP/Name of the host you want to ping.

    .PARAMETER Count
        The number of times to ping the host (alias -n).

    .PARAMETER Continuous
        Ping the host continuously, ctrl+c to exit (alias -t).

    .PARAMETER Quiet
        Returns $True if any pings succeeded and $False if all failed.

    .PARAMETER Resolve
        Resolve IP to hostname or display IP when destination is a name (alias -a).

    .PARAMETER Delay
        Time between pings in milliseconds.

    .PARAMETER Timeout
        Timeout in milliseconds to wait for a reply.

    .PARAMETER Port
        The tcp port to ping.
    
    .EXAMPLE
        PS C:\> tPing 192.168.0.1 -TimeOut 100 -Delay 500 -Count 10
        Ping 192.168.0.1 10 times, with a timeout of 100ms and a delay between pings of 500ms.

        TimeStamp            Destination  Response  Time
        ---------            -----------  --------  ----
        03/04/2015 22:39:28  192.168.0.1     False     0
    
    .EXAMPLE
        PS C:\> tPing COMPUTER1 -Delay 5000 -t | tee .\tPing.txt | ogv
        Ping COMPUTER1 every 5 seconds continuously.
        Pipe the output into Tee-Object then Out-GridView to write to a txt file and display in a grid.
    
    .EXAMPLE
        PS C:\> tPing 8.8.8.8 -Resolve
        Resolve a name to an IP address or an IP address to a name (may slow things down).

        TimeStamp            Destination                     Response  Time
        ---------            -----------                     --------  ----
        03/04/2015 22:46:08  google-public-dns-a.google.com      True    18
    .EXAMPLE
        PS C:\> tPing google.com 80 -TimeOut 500
        Ping google.com on tcp port 80 and set the timeout to 500ms.

        TimeStamp            Destination    Response   Time
        ---------            -----------    --------   ----
        03/04/2015 22:56:28  google.com:80      True     22
    .EXAMPLE
        PS C:\> tPing 8.8.8.8 -Port 53 -Resolve

        TimeStamp            Destination                        Response   Time
        ---------            -----------                        --------   ----
        03/04/2015 22:56:28  google-public-dns-a.google.com:53      True     22
    .EXAMPLE
        PS C:\> tPing 192.168.0.40 22 -Quiet
        True

        Returns true or false.
    #>

    [CmdletBinding(DefaultParameterSetName='Count')]

    param(
        [Parameter(Mandatory=$true,Position=0)][ValidateNotNullOrEmpty()][string]$Destination,
        [Parameter(ParameterSetName='Count')][Parameter(ParameterSetName='Quiet')][Alias('n')][Int]$Count = 4,
        [Parameter(ParameterSetName='Continuous')][Alias('t')][Switch]$Continuous,
        [Parameter(ParameterSetName='Quiet')][Switch]$Quiet,
        [Parameter(ParameterSetName='Count')][Parameter(ParameterSetName='Continuous')][Alias('a')][Switch]$Resolve,
        [Parameter(ParameterSetName='Count')][Parameter(ParameterSetName='Continuous')][ValidateRange(100,60000)][Int]$Delay = 1000,
        [Parameter()][ValidateRange(100,60000)][Alias('w')][Int]$TimeOut = 1000,
        [Parameter(Position=1)][Int]$Port
    )
    
    function icmpPinger {
        $TimeStamp = (Get-Date)
        try {$Ping = (New-Object System.Net.NetworkInformation.Ping).Send("$Destination",$TimeOut)}
        catch {Write-Verbose $_}
        finally {
            if ($Resolve -and (($Destination -As [IPAddress]) -As [Bool])) {
                try {$Resolved = ([Net.DNS]::GetHostEntry("$Destination")).HostName} catch {Write-Verbose $_} 
                if ($Resolved) {$Destination = $Resolved}
            }
            elseif ($Resolve) {$Destination = $Ping.Address}

            if ($Ping.Status -eq 'Success'){$Status = $true} else {$Status = $false}
            if ($Ping.RoundtripTime) {$Time = $Ping.RoundtripTime} else {$Time = 0}
            
            [PSCustomObject]@{
                TimeStamp = $TimeStamp
                Destination = $Destination
                Response = $Status
                Time = $Time
            }
        }
    }

    function tcpPinger {
        try {
            $TcpClient = New-Object System.Net.Sockets.TCPClient
            $StopWatch = New-Object System.Diagnostics.Stopwatch
            $TimeStamp = Get-Date
            $Connect = $TcpClient.BeginConnect("$Destination",$Port,$null,$null)
            $StopWatch.Start()
            while ($TcpClient.Connected -ne $true) {
                if ($StopWatch.ElapsedMilliseconds -ge $Timeout) {$Time = 0 ; Break}
                else {$Time = $StopWatch.ElapsedMilliseconds}
            }
            $StopWatch.Stop()
            if (!$Time) {$Time = 0}
        }
        catch {$Time = 0 ; Write-Verbose $_}

        if ($Resolve -and (($Destination -As [IPAddress]) -As [Bool])) {
            try {$Resolved = ([Net.DNS]::GetHostEntry("$Destination")).HostName} catch {Write-Verbose $_} 
            if ($Resolved) {$Destination = "$Resolved`:$Port"}
        }
        elseif ($Resolve) {$Destination = $TcpClient.Client.RemoteEndPoint}
        else {$Destination = "$Destination`:$Port"}

        [PSCustomObject]@{
            TimeStamp = $TimeStamp
            Destination = $Destination
            Response = $TcpClient.Connected
            Time = $Time
        }
        $TcpClient.Close()
    }
    
    if ($Quiet) {
        $Response = @()
        for ($x=1 ; $x -le $Count ; $x++) {
            if ($Port) {$Response += (tcpPinger).Response}
            else {$Response += (icmpPinger).Response}
            if ($x -lt $Count) {Start-Sleep -Milliseconds $Delay}
        }
        if ($Response -contains $true) {$true}
        else {$false}
    }

    elseif ($Continuous) {
        while ($true) {
            if ($Port) {tcpPinger}
            else {icmpPinger}
            Start-Sleep -Milliseconds $Delay
        }
    }
    
    else {
        for ($x=1 ; $x -le $Count ; $x++) {
            if ($Port) {tcpPinger}
            else {icmpPinger}
            if ($x -lt $Count) {Start-Sleep -Milliseconds $Delay}
        }
    }
}

function Invoke-NetworkScan {

    <#
    .NOTES
        ######################
         mail@nimbus117.co.uk
        ######################
    
    .SYNOPSIS
        Network and port scanner.

    .PARAMETER StartIPAddress

        Starting IP address in range.

    .PARAMETER EndIPAddress

        Ending Ip address in range.

    .PARAMETER IPAddress

        Ip address when using SubnetMask or PrefixLength parameters.

    .PARAMETER SubnetMask
        
        Subnet mask to use with IPAddress parameter.

    .PARAMETER PrefixLength

        Prefix length (CIDR) to use with IPAddress parameter. 

    .PARAMETER JobLimit

        Max number of jobs to run concurrently.

    .PARAMETER ScanPorts

        Scan ports specified by Ports parameter.

    .PARAMETER Ports
    
        List of ports to scan.

    .PARAMETER ResolveHost

        Attempt to resolve host name from IP.

    .PARAMETER IncludeFailedPing

        Display results that fail ping. ResolveHost and ScanPorts will not be run against IP addresses thsat fail ping unless this parameter is set.

    .PARAMETER Timeout

        ICMP and TCP ping timeout value.

    .EXAMPLE

        Invoke-NetworkScan -StartIPAddress 172.31.1.1 -EndIPAddress 172.31.1.10

    .EXAMPLE
        
        Get-NetAdapter -Name Management* | Get-NetIPAddress -AddressFamily IPv4 | Invoke-NetworkScan -ScanPorts

    .EXAMPLE

        Invoke-NetworkScan -IPAddress 172.31.1.1 -PrefixLength 24 -ResolveHost | Out-GridView

    #>

    param(
        [parameter(ParameterSetName='Range',position=0)][IPAddress]$StartIPAddress,
        [parameter(ParameterSetName='Range',position=1)][IPAddress]$EndIPAddress,  
        [parameter(ParameterSetName='Mask',ValueFromPipelineByPropertyName=$true)][parameter(ParameterSetName='Prefix')][IPAddress]$IPAddress,
        [parameter(ParameterSetName='Mask')][IPAddress]$SubnetMask,  
        [parameter(ParameterSetName='Prefix',ValueFromPipelineByPropertyName=$true)][Int]$PrefixLength,
        [parameter()][Int]$JobLimit = 10,
        [parameter()][Int[]]$Ports = @(20,21,22,23,53,69,80,110,137,138,139,143,389,443,445,587,1025,1433,8080,3306,3389,5985),
        [parameter()][Switch]$ResolveHost,
        [parameter()][Switch]$ScanPorts,
        [parameter()][Switch]$IncludeFailedPing,
        [parameter()][Int]$Timeout = 100
    )  

    begin {

        function IP-toINT64 ($IP) {  
            $Octets = $IP.Split(".")  
            [Long]([Long]$Octets[0]*16777216 +[Long]$Octets[1]*65536 +[Long]$Octets[2]*256 +[Long]$Octets[3])
        }

        function INT64-toIP ([Long]$Int) {
            (([Math]::Truncate($Int/16777216)).ToString()+"."+([Math]::Truncate(($Int%16777216)/65536)).ToString()+"."+([Math]::Truncate(($Int%65536)/256)).ToString()+"."+([Math]::truncate($Int%256)).ToString())
        }

        $Global:ScanResults = @()
    }
  
    process {

        if ($PrefixLength) {$SubnetMask = [IPAddress]::Parse((INT64-toIP -Int ([Convert]::ToInt64(("1"*$PrefixLength+"0"*(32-$PrefixLength)),2))))}
        if ($IPAddress) {
            $NetworkAddress = New-Object Net.IPAddress ($SubnetMask.address -band $IPAddress.Address)
            $BroadcastAddress = New-Object Net.IPAddress (([IPAddress]::Parse("255.255.255.255").Address -bxor $SubnetMask.Address -bor $NetworkAddress.Address))
            $Start = IP-toINT64 -IP $NetworkAddress.IPAddressToString
            $End = IP-toINT64 -IP $BroadcastAddress.IPAddressToString
        }
        else {$Start = IP-toINT64 -IP $StartIPAddress.IPAddressToString ; $End = IP-toINT64 -IP $EndIPAddress.IPAddressToString}
        
        $IPRange = @()
        for ($i = $Start; $i -le $End; $i++) {$IPRange += INT64-toIP -Int $i} 
        
        $IPRangeCount = ($IPRange | Measure-Object).Count
        $JobCounter = 0
        $CompletedCounter = 0

        $IPRange | ForEach-Object {
            
            $JobCounter++
            Write-Progress -Activity "Scan-Network" -Status "$JobCounter of $IPRangeCount Started" -PercentComplete (($JobCounter / $IPRangeCount) * 100) -CurrentOperation "Starting Scan - $_" -Id 1

            Start-Job -Name "Scan-Network$JobCounter" {

                $IPAddress = $Args[0] ; $Ports = $Args[1] ; $Timeout = $Args[2] ; $ResolveHost = $Args[3] ; $ScanPorts = $Args[4] ; $IncludeFailedPing = $Args[5]
                $PingResults = (New-Object System.Net.NetworkInformation.Ping).Send("$IPAddress",$Timeout)
                if (($PingResults.Status -eq 'Success') -or $IncludeFailedPing) {
                    if ($ResolveHost) {try {$HostName = ([Net.DNS]::GetHostEntry("$IPAddress")).HostName} catch {$HostName = ""}}
                    if ($ScanPorts) {
                        $OpenPorts = @()
                        foreach ($Port in $Ports) {
                            $TcpClient = New-Object System.Net.Sockets.TcpClient
                            $BeginConnect = $TcpClient.BeginConnect("$IPAddress",$Port,$null,$null)
                            if ($TcpClient.Connected) {$OpenPorts += $Port}
                            else {
                            
                                Start-Sleep -Milliseconds $Timeout
                                if ($TcpClient.Connected) {$OpenPorts += $Port}     
                            }
                            $TcpClient.Close()
                        }
                    }
                    
                    New-Object PSCustomObject -Property @{IPAddress = $IPAddress ; Ping = $PingResults.Status ; HostName = $HostName ; OpenPorts = $OpenPorts}
                }
            } -ArgumentList $_, $Ports, $Timeout, $ResolveHost, $ScanPorts, $IncludeFailedPing  | Out-Null

            if ($JobCounter -eq $IPRangeCount) {Get-Job -Name "Scan-Network*" | Wait-Job | Out-Null}
            $CompletedJobs = Get-Job -Name "Scan-Network*" | Where-Object {$_.State -eq 'Completed'}
            $Result = $CompletedJobs | Receive-Job | Select-Object IPAddress, Ping, HostName, OpenPorts
            $Result
            $Global:ScanResults = $Global:ScanResults += $Result
            $CompletedJobs | Remove-Job

            While((Get-Job -State 'Running' | Measure-Object).Count -ge $JobLimit) {
                Write-Progress -Activity "Scan-Network" -Status "$JobCounter of $IPRangeCount Started" -PercentComplete (($JobCounter / $IPRangeCount) * 100) -CurrentOperation "Reached job limit" -id 1
                Start-Sleep -Seconds 1
            }
        }
    }
}

function Invoke-PortListener {

    <#
    .NOTES
        ######################
            mail@nimbus117.co.uk
        ######################
    
    .SYNOPSIS
        Creates a TCP listener on the specified port. Ctrl-C to quit and stop the listener.

    .PARAMETER Port

        Port to listen on.

    .EXAMPLE

        Invoke-PortListener 8080
    #>

    param(
        [parameter(Mandatory=$true,position=0)]
        [ValidateRange(1,65536)]
        [int]$Port
    )
    
    $ActiveListeners = ([System.Net.NetworkInformation.IPGlobalProperties]::GetIPGlobalProperties().GetActiveTcpListeners()).Port
    if ($ActiveListeners -eq $Port) {Throw "Already listening on port $Port"}

    try {
        
        $Listener = [System.Net.Sockets.TcpListener]$Port
        $Listener.Start()
        Write-Warning "Listening on Port $Port"
        while ($true) {}
    }
    catch {Write-Error $_.Exception.Message}
    finally {$Listener.Stop()}
}
Set-Alias listen -Value Invoke-PortListener

function Download-File {
    <#
    .NOTES
        ######################
         mail@nimbus117.co.uk
        ######################
    
    .SYNOPSIS
        Download a file over http or ftp

        PS C:\Users\john> Download-File "http://download.microsoft.com/download/1/8/E/18EA4843-C596-4542-9236-DE46F780806E/Windows8.1-KB2693643-x64.msu"


        Url      : http://download.microsoft.com/download/1/8/E/18EA4843-C596-4542-9236-DE46F780806E/Windows8.1-KB2693643-x64.msu
        Path     : C:\Users\john
        FileName : Windows8.1-KB2693643-x64.msu
        Size     : 67.59MB
        Silent   : False

        Press Enter To Continue: 

    .EXAMPLE
        PS C:\> Download-File ftp://speedtest:speedtest@ftp.otenet.gr/test10Mb.db -Path C:\temp
        Save file in path C:\temp\test10Mb.db (default .\test10Mb.db).
    .EXAMPLE
        PS C:\> Download-File http://download.linnrecords.com/test/flac/test192.aspx -FileName test192.flac
        Save file as .\test192.flac (default .\test192.aspx).
    .EXAMPLE
        PS C:\> Download-File http://download.thinkbroadband.com/100MB.zip -Quiet
        Using '-Quiet' suppresses the prompt to continue and disables the progress display.
    #>

    Param(
        [parameter(Mandatory=$true,Position=0)]
        [ValidateNotNullOrEmpty()]
        [string]$Url,
        [parameter(Position=1)]
        [string]$Path = (Get-Location),
        [string]$FileName = ($Url.TrimEnd('/') -split '/')[-1],
        [switch]$Quiet
    )

    $ErrorActionPreference = 'Stop'
    
    try {
        $Uri = New-Object "Uri" "$Url"
        $Request = [Net.HttpWebRequest]::Create($Uri)
        $Request.Set_Timeout(5000)
        $Response = $Request.GetResponse()
        $TotalBytes = $Response.Get_ContentLength()
        $ResponseStream = $Response.GetResponseStream()

        $Path = Resolve-Path $Path
        $FullPath = Join-Path -Path $Path -ChildPath $FileName
        if (Test-Path $FullPath) {
            $x = Get-ChildItem $FullPath
            $Count = (Get-ChildItem "$(($x.Fullname).Replace($x.Extension,'*'))" | Measure-Object).Count
            $FullPath = Join-Path -Path $x.DirectoryName -ChildPath ($x.BaseName + "($Count)" + $x.Extension)
        }

        if (-not($Quiet)) {

            New-Object PSCustomObject -Property @{
                Url = $Url
                FileName = ($FullPath -split '\\')[-1]
                Path = $Path
                Size = "$([math]::Round(($TotalBytes / 1MB),2))MB"
                Silent = $Quiet
            } | Select-Object Url,Path,FileName,Size,Silent

            Read-Host "Press Enter To Continue"
            $LastTime = Get-Date
            $StopWatch = [System.Diagnostics.Stopwatch]::StartNew()
        }

        $FileStream = New-Object IO.FileStream $FullPath, Create
        $Buffer = New-Object byte[] 10KB
        $Count = $ResponseStream.Read($Buffer,0,$Buffer.Length)
        $DownloadedBytes = $Count

        while ($Count -gt 0) {
            $FileStream.Write($Buffer, 0, $Count)
            $Count = $ResponseStream.Read($Buffer,0,$Buffer.length)
            $DownloadedBytes = $DownloadedBytes + $Count
    
            if ((-not($Quiet)) -and ($StopWatch.Elapsed.TotalMilliseconds -ge 500)) {
                $Rate = (($DownloadedBytes - $LastDLBytes) / (New-TimeSpan –Start $LastTime –End $(Get-date)).TotalSeconds)
                $RateMB = "{0:N2}" -f ($Rate / 1MB)
                $Status = "$("{0:N2}" -f ($DownloadedBytes / 1MB)) of $([math]::round(($TotalBytes / 1MB), 2)) MB @ $RateMB MB/s"
                $Percent = (($DownloadedBytes / $TotalBytes) * 100)
                $TimeRemaining = ($TotalBytes - $DownloadedBytes) / $Rate
                Write-Progress -Activity "$Url -> $FullPath" -Status $Status -PercentComplete $Percent -SecondsRemaining $TimeRemaining
                $LastDLBytes = $DownloadedBytes
                $LastTime = Get-Date
                $StopWatch.Reset()
                $StopWatch.Start()
            }
        }
    }

    catch {Write-Error $_}

    finally {
        if ($FileStream -ne $null) {
            $FileStream.Flush()
            $FileStream.Close()
            $FileStream.Dispose()
        }
        if ($ResponseStream -ne $null) {
            $ResponseStream.Dispose()
        }
    }
}
Set-Alias dlf -Value Download-File

function Get-Size {

    <#
    .NOTES
        ######################
         mail@nimbus117.co.uk
        ######################

    .SYNOPSIS
        Get folder and file sizes.
    .DESCRIPTION
        Returns the size of a file or the total size of all files in a folder.

        C:\PS>Get-Size .

        FullName       Recurse     Size FileCount
        --------       -------     ---- ---------
        C:\Users\User  False   52436067         6
    .PARAMETER Path
        Path to folder or file.
    .PARAMETER Unit
        Units to display size in (Size is displayed in Bytes if no units are specified).
        KB, MB, GB, TB
    .PARAMETER Round
        Round size to this many decimal places.
    .PARAMETER Filter
        Only include files that match the filter (has no affect if the path parameter is a file). 
    .PARAMETER Recurse
        Include sub folders (has no affect if the path parameter is a file).
    .PARAMETER Force
        Include hidden and system files.
    .EXAMPLE
        Get-Size . MB -Recurse -Force

        FullName       Recurse  SizeMB FileCount
        --------       -------  ------ ---------
        C:\Users\User  True    1830.43      2005
    .EXAMPLE
        Get-Size -Path .\Dropbox -Unit GB -Recurse

        FullName               Recurse SizeGB FileCount
        --------               ------- ------ ---------
        C:\Users\User\Dropbox  True      0.45      1644
    .EXAMPLE
        Get-ChildItem -Directory | Get-Size -Unit KB -Recurse

        FullName                      Recurse    SizeKB FileCount
        --------                      -------    ------ ---------
        C:\Users\User\.VirtualBox     True     25823.45        18
        C:\Users\User\Contacts        True            0         0
        C:\Users\User\Desktop         True            0         0
        C:\Users\User\Documents       True    152366.39       201
    .EXAMPLE
        Get-ChildItem -File | Get-Size -Unit KB -Recurse

        FullName                           Recurse SizeKB FileCount
        --------                           ------- ------ ---------
        C:\Users\User\.recently-used.xbel  True      0.21          
        C:\Users\User\50MB.zip             True     51200          
        C:\Users\User\diskAlignment.txt    True       1.3          
        C:\Users\User\diskAlignmentXP.txt  True       1.3          
    .EXAMPLE
        Get-Size ~ KB -Recurse -Filter *.ps1 -Round 4

        FullName       Recurse   SizeKB FileCount
        --------       -------   ------ ---------
        C:\Users\User  True    635.0293       157
    .EXAMPLE
        Get-Size .,documents,downloads MB -re

        FullName                 Recurse  SizeMB FileCount
        --------                 -------  ------ ---------
        C:\Users\User            True    1752.94      2141
        C:\Users\User\documents  True      148.8       202
        C:\Users\User\downloads  True     201.39         5
    #>

    param(
        [Parameter(Position=0,Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [Alias('FullName')]
        [string[]]$Path,
        [Parameter(Position=1)][validateset("KB","MB","GB","TB")][String]$Unit,
        [Parameter(Position=2)][ValidateRange(0,15)][Int]$Round = 2,
        [Parameter()][String]$Filter,
        [Parameter()][Switch]$Recurse,
        [Parameter()][Switch]$Force
    )

    process {
            
        $Path | ForEach-Object {
            
            $ItemParams = @{Path = $_}
            if ($Force) {$ItemParams += @{Force = $true}}

            $Item = Get-Item @ItemParams

            if ($Item.PSIsContainer) {

                $ChildItemParams = @{Path = $Item}
                if ($Recurse) {$ChildItemParams += @{Recurse = $true}}
                if ($Force) {$ChildItemParams += @{Force = $true}}
                if ($Filter) {$ChildItemParams += @{Filter = $Filter}}

                $ChildItems = Get-ChildItem @ChildItemParams
                $Size = ($ChildItems | Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
                $Count = ($ChildItems | Where-Object {!$_.PSIsContainer} | Measure-Object -ErrorAction SilentlyContinue).Count
            }

            else {$Size = $Item.Length ; $Count = $null}

            if ($Item) {
                [PsCustomObject]@{
                    FullName = $Item.FullName
                    Recurse = $Recurse
                    "Size$($Unit.ToUpper())" = [math]::Round(($Size / ('1' + $Unit)), $Round)
                    FileCount = $Count
                }
            }

            Clear-Variable Item
        }
    }
}
Set-Alias gs -Value Get-Size

function Extract-Zip {

    param(
        [Parameter(Mandatory=$true,Position=0)][ValidateNotNullOrEmpty()]$Source,
        [parameter(Position=1)]$Destination = (Get-Location),
        [Parameter(Position=2)][switch]$Silent
    )

    $ErrorActionPreference = 'Stop'

    if ($Silent) {$Option = 20} else {$Option = 0}

    try {
        $Shell = New-Object -ComObject Shell.Application
        $Zip = $Shell.NameSpace((Resolve-Path $Source).Path)
        foreach($Item in $Zip.Items()) {
            $Shell.Namespace((Resolve-Path $Destination).Path).CopyHere($Item,$Option)
        }
    }
    catch {"ERROR: $_"}
}
Set-Alias exz -Value Extract-Zip

function Get-LicenseStatus {
        
    param([
        Parameter(Position=0,Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [alias('ComputerName')]
        [ValidateNotNullOrEmpty()]
        [string[]]$Name
    )

    begin {$LicenseStatus=@{0="Unlicensed"; 1="Licensed"; 2="OOBGrace"; 3="OOTGrace"; 4="NonGenuineGrace"; 5="Notification"; 6="ExtendedGrace"}}

    process {

        $Name | ForEach-Object {

            $Command = Invoke-Command -ComputerName $_ -ScriptBlock {
            
                Get-WmiObject -Query  "SELECT * FROM SoftwareLicensingProduct WHERE PartialProductKey <> null AND ApplicationId='55c92734-d682-4d71-983e-d6ec3f16059f' AND LicenseIsAddon=False"
            }
            
            if ($Command) {$Command | Select-Object @{l='ComputerName';e={$_.PSComputerName}}, @{l='OS';e={$_.Name}}, @{l='LicenseStatus';e={$LicenseStatus[[int]$_.LicenseStatus]}}, PartialProductKey}
        }
    }
}

function Install-LicenseKey {

    param(
        [Parameter(Position=0,Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [alias('ComputerName')]
        [ValidateNotNullOrEmpty()]
        [string[]]$Name,
        [Parameter(Position=1,Mandatory=$true)]
        [ValidatePattern("^\S{5}-\S{5}-\S{5}-\S{5}-\S{5}$")]
        [ValidateNotNullOrEmpty()]
        [string]$Key
    )

    begin {$LicenseStatus=@{0="Unlicensed"; 1="Licensed"; 2="OOBGrace"; 3="OOTGrace"; 4="NonGenuineGrace"; 5="Notification"; 6="ExtendedGrace"}}

    process {

        $Name | ForEach-Object {

            $Command = Invoke-Command -ComputerName $_ -ArgumentList $Key -ScriptBlock {

                $LicensingService = Get-WmiObject -Class SoftwareLicensingService
                $LicensingService.InstallProductKey($args[0]) | Out-Null
                $LicensingService.RefreshLicenseStatus() | Out-Null

                $Status = Get-WmiObject -Class SoftwareLicensingProduct | Where-Object {$_.PartialProductKey -and $_.ApplicationId -eq '55c92734-d682-4d71-983e-d6ec3f16059f'}
                $x = 0

                while ($Status.LicenseStatus -ne 1 -and $x -lt 5) {
                    
                    $x++
                    Start-sleep 5
                    $Status = Get-WmiObject -Query  "SELECT * FROM SoftwareLicensingProduct WHERE PartialProductKey <> null AND ApplicationId='55c92734-d682-4d71-983e-d6ec3f16059f' AND LicenseIsAddon=False"
                }

                $Status
            }
            
            if ($Command) {$Command | Select-Object @{l='ComputerName';e={$_.PSComputerName}}, @{l='OS';e={$_.Name}}, @{l='LicenseStatus';e={$LicenseStatus[[int]$_.LicenseStatus]}}, PartialProductKey}
        }
    }
}

function Set-ServiceLogOn {

    param(
        
        [Parameter(position=0,Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [string[]]
        $Name,
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [PSCredential]
        [System.Management.Automation.CredentialAttribute()]
        $ServiceCredential,
        [String]
        $ComputerName,
        [PSCredential]
        [System.Management.Automation.CredentialAttribute()]
        $Credential,
        [Switch]
        $StartService = $true
    )

    process {

        $Name | ForEach-Object {
            
            try {

                $Params = @{Class = 'Win32_Service' ; Filter = "name='$_'" ; ErrorAction = 'Stop'}
                if ($ComputerName) {$Params.Add('ComputerName',$ComputerName)}
                if ($Credential) {$Params.Add('Credential',$Credential)}

                $Service = Get-WmiObject @Params

                if ($Service) {

                    $Change = $Service.Change($null,$null,$null,$null,$null,$null,$ServiceCredential.UserName,$ServiceCredential.GetNetworkCredential().Password)
                    if ($Change.ReturnValue -gt 0) {throw "Error changing $Name service user/password - Win32_Service return value $($Change.ReturnValue)."}
                    
                    if ($Service.Started) {

                        $Stop = $Service.StopService()
                        if ($Stop.ReturnValue -gt 0) {throw "Error stopping $Name service - Win32_Service return value $($Stop.ReturnValue)."}

                        $Service = Get-WmiObject @Params
                        while ($Service.Started) {Start-Sleep 1 ; $Service = Get-WmiObject @Params}
                    }
                
                    if ($StartService) {

                        $Start = $Service.StartService()
                        if ($Start.ReturnValue -gt 0) {throw "Error starting $Name service - Win32_Service return value $($Start.ReturnValue)."}
                    }
                
                    Get-WmiObject @Params | Select-Object PSComputerName, Name, State, StartName
                }

                else {throw "Cannot find any service with service name '$Name'."}
            }

            catch {Write-Error $_}

            finally {if ($Service) {Clear-Variable Service}}
        }
    }
}

function Invoke-PSMonitor {

    param(
        
        [parameter(position=0)][string]$ComputerName,
        [Parameter(Position=1)][System.Management.Automation.CredentialAttribute()][PSCredential]$Credential,
        [int]$UpdateDelayMs = 500
    )

    $CPUParams = @{Class = 'Win32_Processor'}
    $MemParams = @{Class = 'Win32_OperatingSystem'}
    $VolParams = @{Class = 'win32_Volume'}

    if ($ComputerName) {$CPUParams.Add('ComputerName',$ComputerName) ; $MemParams.Add('ComputerName',$ComputerName) ; $VolParams.Add('ComputerName',$ComputerName)}
    if ($Credential) {$CPUParams.Add('Credential',$Credential) ; $MemParams.Add('Credential',$Credential) ; $VolParams.Add('Credential',$Credential)}

    while ($true) {

        $CPULoad = (Get-CimInstance @CPUParams | Measure-Object -Property LoadPercentage -Average).Average
        $Memory = Get-CimInstance @MemParams
        $TotalMemoryGB = [math]::Round($Memory.TotalVisibleMemorySize / 1MB,1)
        $UsedMemoryGB = [math]::Round(($Memory.TotalVisibleMemorySize - $Memory.FreePhysicalMemory) / 1MB, 1)
        $UsedMemoryPer = [math]::Round(($UsedMemoryGB / $TotalMemoryGB) * 100,0)

        Write-Progress -Activity "CPU Load" -Status "$CPULoad%" -PercentComplete $CPULoad -Id 1
        Write-Progress -Activity "Memory Usage" -Status "$UsedMemoryGB/$TotalMemoryGB`GB ($UsedMemoryPer%)" -PercentComplete $UsedMemoryPer -Id 2

        $Volumes = Get-CimInstance @VolParams | 
        Where-Object {($_.Capacity -ne $null) -and $_.DriveLetter} |
        Sort-Object DriveLetter | Select-Object -First 3

        $x = 2
        $Volumes | ForEach-Object {
            $x++
            $CapacityGB = [math]::Round($_.Capacity / 1GB, 1)
            $FreeSpaceGB = [math]::Round($_.FreeSpace / 1GB, 1)
            $UsedSpaceGB = [math]::Round(($_.Capacity - $_.FreeSpace) / 1GB, 1)
            $UsedSpacePer = ($UsedSpaceGB / $CapacityGB) * 100
            Write-Progress -Activity $_.Name -Status "$UsedSpaceGB/$CapacityGB`GB ($UsedSpacePer%)" -PercentComplete $UsedSpacePer -Id $x
        }

        Start-Sleep -Milliseconds $UpdateDelayMs
    }
}

function Get-Memory {

    param(
        
        [Parameter(position=0)][string[]]$ComputerName,
        [Parameter()][PSCredential][System.Management.Automation.CredentialAttribute()]$Credential,
        [Parameter()][validateset("KB","MB","GB","TB")][String]$Unit = 'GB'
    )

    $CorrectUnit = ('1' + $Unit) / 1024

    $WmiParams = @{Class = 'Win32_OperatingSystem'}

    if ($ComputerName) {$WmiParams += @{ComputerName = $ComputerName}}

    if ($Credential) {$WmiParams += @{Credential = $Credential}}

    Get-CimInstance @WmiParams | 

    Select-Object `
        PSComputerName,
        @{l="Total$($Unit.ToUpper())";e={[math]::Round($_.TotalVisibleMemorySize / $CorrectUnit, 2)}},
        @{l="Used$($Unit.ToUpper())";e={[math]::Round(($_.TotalVisibleMemorySize - $_.FreePhysicalMemory) / $CorrectUnit, 2)}},
        @{l="Free$($Unit.ToUpper())";e={[math]::Round($_.FreePhysicalMemory / $CorrectUnit, 2)}},
        @{l='PercentFree';e={[math]::Round((($_.FreePhysicalMemory / $_.TotalVisibleMemorySize) * 100), 0)}}

}

function Get-SMARTStatus {

    param(
        
        [Parameter(position=0)][string[]]$ComputerName,
        [Parameter()][PSCredential][System.Management.Automation.CredentialAttribute()]$Credential
    )

    $WmiParams = @{NameSpace = 'root\wmi' ; Class = 'MSStorageDriver_FailurePredictStatus'}

    if ($ComputerName) {$WmiParams += @{ComputerName = $ComputerName}}

    if ($Credential) {$WmiParams += @{Credential = $Credential}}

    Get-CimInstance @WmiParams | 

    Select PSComputerName, InstanceName, PredictFailure, Reason
}
Set-Alias SMART -Value Get-SMARTStatus

function Get-RDGCurrentConnections {

    param(
        
        [Parameter(position=0)][string[]]$ComputerName,
        [Parameter(position=1)][PSCredential][System.Management.Automation.CredentialAttribute()]$Credential
    )

    $TProtocol = @{0 = "RPC/HTTP" ; 1 = "HTTP" ; 2 = "UDP"}

    $WmiParams = @{
        
        Namespace = 'root\cimv2\terminalservices'
        Query = "select * from  Win32_TSGatewayConnection"
    }

    if ($ComputerName) {$WmiParams += @{ComputerName = $ComputerName}}

    if ($Credential) {$WmiParams += @{Credential = $Credential}}

    Get-CimInstance @WmiParams |
    
    Select-Object `
        @{label='ConnectedTime';expression={[System.Management.ManagementDateTimeconverter]::ToDateTime($_.ConnectedTime)}},
        @{label='Gateway';expression={$_.PsComputerName}},
        ClientAddress, ConnectedResource, UserName,
        @{label='ConnectionDuration';expression={[System.Management.ManagementDateTimeconverter]::ToTimeSpan($_.ConnectionDuration)}}, `
        @{label='IdleTime';expression={[System.Management.ManagementDateTimeconverter]::ToTimeSpan($_.IdleTime)}}, `
        @{label='Transport';expression={$TProtocol[[int]$_.TransportProtocol]}}
}

function New-FileCleanUpTask {
    <#
    .SYNOPSIS 
    Deploys scripts (.ps1 and .bat files) and creates a scheduled task on a remote computer to remove files on a schedule. 
    e.g. Remove log files from a specific directory older than 7 days every day at 7:00am.

    .PARAMETER ComputerName
    Name of remote computer.

    .PARAMETER FilePath
    File/Folder path to be cleaned. e.g. C:\path or C:\path\*.txt

    .PARAMETER OlderThanDays
    Files older than x days will be removed.
    
    .PARAMETER RunAt
    Time of day the scheduled task is run.

    .PARAMETER Credential
    Specifies a user account that has permission to perform this action.

    .PARAMETER Recurse
    Removes items in the specified locations and in all child items of the locations.

    .PARAMETER TaskName
    Name of scheduled task.
    
    .PARAMETER ScriptPath
    Path scripts are deplyed to.

    .PARAMETER SmtpFrom
    Smtp from address for error email.

    .PARAMETER SmtpTo
    Smtp to address for error email.

    .PARAMETER SmtpServer
    Smtp server address for error email.

    .EXAMPLE
    PS C:\> New-FileCleanUpTask -ComputerName SRV01 -FilePath C:\temp\Test\*.log -OlderThanDays 7 -RunAt 06:00

    .EXAMPLE
    PS C:\> New-FileCleanUpTask -ComputerName SRV02 -FilePath 'C:\temp\*.log','C:\inetpub\logs' -OlderThanDays 28 -RunAt 10:00pm
    #>

    param (
        [parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][string[]]$ComputerName,

        [parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][string[]]$FilePath,

        [parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][uint32]$OlderThanDays,

        [parameter(Mandatory=$true)][DateTime]$RunAt,

        [PSCredential][System.Management.Automation.CredentialAttribute()]$Credential,

        [switch]$Recurse,

        [ValidateNotNullOrEmpty()][String]$TaskName = 'FileCleanUpTask',

        [ValidateNotNullOrEmpty()][string]$ScriptPath = 'c:\zts\scripts\FileCleanUpTask',

        [string]$SmtpFrom = 'ztsReports@zonalconnect.com',

        [string]$SmtpTo = 'zts@zonal.co.uk',

        [string]$SmtpServer = 'mail.zonalconnect.local'
    )

    $ErrorActionPreference = 'Stop'

    $sb = {

        $BatFile = @"
@echo off
setlocal enableextensions
cd /d "%~dp0"

::Run ps script
powershell.exe -executionpolicy bypass -NoProfile -NoLogo -file "%~dpn0.ps1" *>&1 > "%~dpn0.log"
"@

        $PSScript = @"

function Remove-Files {

    param(
        
        [parameter(Mandatory=`$true)][ValidateNotNullOrEmpty()][string]`$Path,
        [parameter(Mandatory=`$true)][ValidateNotNullOrEmpty()][uint32]`$OlderThanDays,
        [switch]`$Recurse
    )

    `$ErrorActionPreference = 'Stop'

    try {

        Get-Date

        `$GciParams = @{Path = `$Path ; File = `$true}
        if (`$Recurse) {`$GciParams += @{Recurse = `$True}}
        
        `$OlderThan = (Get-Date).Date.AddDays(-`$OlderThanDays)
        "Deleting Files older than - `$(Get-Date `$OlderThan -Format 'dd/MM/yyyy')"

        `$FilesToDelete = Get-ChildItem @GciParams | 
        Where-Object {`$_.LastWriteTime -lt `$OlderThan} 
        
        `$FilesToDelete | Format-Table -AutoSize

        `$FilesToDelete | Remove-Item -Force -Confirm:`$False
    }

    catch {
    
        `$MessageParameters = @{
            Subject = "Remove-Files - `$env:COMPUTERNAME `$Path"
            From = '$Using:SmtpFrom'
            To = '$Using:SmtpTo'
            Body = `$_
            SmtpServer = '$Using:SmtpServer'
        }

        `$_
        Send-MailMessage @MessageParameters
    }
}

"@

        $Using:FilePath | ForEach-Object {$PSScript += "`nRemove-Files -Path '$_' -OlderThanDays $Using:OlderThanDays -Recurse:`$$Using:Recurse"}

        if (-not(Test-Path $Using:ScriptPath)) {mkdir $Using:ScriptPath | Out-Null}

        $PSScript | Out-File -FilePath "$Using:ScriptPath\$Using:TaskName.ps1" -Force

        $BatFile | Out-File -FilePath "$Using:ScriptPath\$Using:TaskName.bat" -Encoding ascii -Force

        $STAction = New-ScheduledTaskAction -Execute "$Using:ScriptPath\$Using:TaskName.bat"

        $STTrigger = New-ScheduledTaskTrigger -Daily -At $Using:RunAt

        Register-ScheduledTask -Action $STAction -Trigger $STTrigger -User 'System' -TaskName $Using:TaskName -RunLevel Highest -Force

    }

    $params = @{ScriptBlock = $sb ; ComputerName = $ComputerName}
    if ($Credential) {$params += @{Credential = $Credential}}

    Invoke-Command @params
}

function Resume-DfsReplication {

    param([Parameter(Position=0)][string]$ComputerName)

    if (Test-Connection $ComputerName -Quiet) {

        $WMI = Get-WmiObject -Class DfsrVolumeConfig -Namespace root\microsoftdfs -ComputerName $ComputerName

        $WMI.ResumeReplication()
    }

}

function Uninstall-MSIProduct {

    [CmdletBinding(SupportsShouldProcess = $true,ConfirmImpact = 'High',DefaultParameterSetName=0)]
    
    param(
        
        [Parameter(ParameterSetName=0,Mandatory=$true,Position=0,ValueFromPipelineByPropertyName=$true)]
        [Alias('DisplayName')]
        [ValidateNotNullOrEmpty()]
        [string[]]$ProductName,

        [Parameter(ParameterSetName=0)][Switch]$Quiet,

        [Parameter(ParameterSetName=0)][Switch]$NoRestart,

        [Parameter(ParameterSetName=1)][Switch]$List
    )

    begin {

        $ErrorActionPreference = 'Stop'

        $RegPath = 'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'

        if ($List) {Get-ChildItem $RegPath | ForEach-Object {Get-ItemProperty $_.PSPath} |
     
        Where-Object {$_.DisplayName} | Sort-Object DisplayName | Select-Object DisplayName,DisplayVersion,Publisher,UninstallString}
    }

    process {

        ForEach ($product in $ProductName) {
        
            $Uninstall = Get-ChildItem $RegPath | ForEach-Object {Get-ItemProperty $_.PSPath} | Where-Object {$_.DisplayName -match $Product}

            if ($Uninstall) {

                $Uninstall | ForEach-Object {

                    if ($PSCmdlet.ShouldProcess($_.DisplayName,'Uninstall')) {

                        $Args = @()

                        if ($Quiet) {$Args += '/q'}

                        if ($NoRestart) {$Args += '/norestart'}

                        $Args += "/X $($_.PSChildName)"

                        Start-Process 'msiexec.exe' -ArgumentList $Args -Wait
                    }
                }
            }
        
            else {Write-Warning "Could not find product matching '$ProductName'"}
        }
    }
}

function Parse-Logs {

    param(
        
        [Parameter(Mandatory = $True, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string]$SearchString,
        
        [Parameter(Position = 1)]
        [ValidateNotNullOrEmpty()]
        [string]$Path = (Get-Location),

        [switch]$Recurse
    )

    $Gci_Params = @{Path = $Path;File = $true;Recurse = $Recurse}

    Get-ChildItem @Gci_Params | ForEach-Object {

        "###";"$($_.FullName)";"###"

        $_ | Get-Content | Where-Object {$_ | Select-String $SearchString}
    }
}

function New-RandomPassword {

    <#
        https://makemeapassword.org/Api
    #>

    [CmdletBinding(DefaultParameterSetName='Password')]

    param(
        
        [ValidateSet('plain','xml','json')][string]$Result = 'plain',
        [ValidateRange(1,128)][int]$Length = 12,
        [ValidateRange(1,50)][int]$Count = 1,
        [Parameter(ParameterSetName='Password')][switch]$Special,
        [Parameter(ParameterSetName='Pin')][switch]$Pin
    )

    try {

        if ($Pin) {$Type = 'pin'}
    
        else {$Type = 'alphanumeric'}

        $URI = "https://makemeapassword.org/api/v1/$Type/$Result`?c=$Count&l=$Length"

        if ($Special) {$URI = $URI + '&sym=True'}

        $Request = Invoke-WebRequest $URI -UseBasicParsing

        $Request.Content -split '\r\n'
    }

    catch {Write-Warning $_.Exception.Message}
}
Set-Alias password -Value New-RandomPassword

function Get-DfsrReplicatedFolder {
    
    <#
    .SYNOPSIS
        Gets a DFS replicated folder.

    .Example
        Get-DfsrReplicatedFolder -ComputerName dca-ior-api1
    
    .Example
        Get-DfsrReplicatedFolder -ReplicationGroupName Main -ReplicatedFolderName FolderRedirect
    #>

    param(
        
        [string]$ReplicationGroupName,

        [string]$ReplicatedFolderName,

        [Alias('PSComputerName')][string[]]$ComputerName,

        [PSCredential]
        [System.Management.Automation.CredentialAttribute()]$Credential
    )

    $ErrorActionPreference = 'Stop'

    try {

        $Session_Params = @{}

        if ($ComputerName) {$Session_Params += @{ComputerName = $ComputerName}}
        if ($Credential) {$Session_Params += @{Credential = $Credential}}
    
        $Session = New-CimSession @Session_Params

        $ReplicatedFolder = Get-CimInstance -Namespace 'root\microsoftdfs' -Class 'DfsrReplicatedFolderInfo' -CimSession $Session

        if ($ReplicationGroupName) {$ReplicatedFolder = $ReplicatedFolder | Where-Object {$_.ReplicationGroupName -eq $ReplicationGroupName}}
        if ($ReplicatedFolderName) {$ReplicatedFolder = $ReplicatedFolder | Where-Object {$_.ReplicatedFolderName -eq $ReplicatedFolderName}}

        $ReplicatedFolder
    }

    catch {Write-Error $_.Exception.Message}
}

function Clear-DfsrConflictFolder {

    <#
    .SYNOPSIS
        Deletes the contents of the 'Conflict and Deleted' folder for a given DFS replicated folder.

    .Example
        Clear-DfsrConflictFolder -ReplicatedFolderGuid 24C51DAC-A00B-4FF3-BFAC-85F977471A17
    
    .Example
        Get-DfsrReplicatedFolder -ComputerName dca-ior-api1,dca-ior-api2 | Clear-DfsrConflictFolder
    #>

    [CmdletBinding(SupportsShouldProcess = $true,ConfirmImpact = 'High')]

    param(
        
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [Alias('Identifier')]
        [string[]]$ReplicatedFolderGuid,
        
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [Alias('PSComputerName')]
        [string]$ComputerName,

        [PSCredential]
        [System.Management.Automation.CredentialAttribute()]$Credential
    )

    Begin {$ErrorActionPreference = 'Stop'}

    Process {

        try {

            $WMI_Params = @{
        
                Namespace = 'root\microsoftdfs'
                Class = 'DfsrReplicatedFolderInfo'
            }

            if ($ComputerName) {$WMI_Params += @{ComputerName = $ComputerName}}
            if ($Credential) {$WMI_Params += @{Credential = $Credential}}
        
            $ReplicatedFolder = Get-WmiObject @WMI_Params -Filter "ReplicatedFolderGuid='$ReplicatedFolderGuid'"

            if ($PSCmdlet.ShouldProcess($ReplicatedFolder.PSComputerName + ':' + $ReplicatedFolder.ReplicationGroupName + '\' + $ReplicatedFolder.ReplicatedFolderName,'CleanupConflictDirectory')) {

                $ReplicatedFolder.CleanupConflictDirectory()
            }
        }

        catch {Write-Error $_.Exception.Message}
    }
}

function Reset-WindowsUpdate {

    param(
        
        [parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [Alias('ComputerName','PSComputerName')]
        [string[]]$Name,

        [PSCredential][System.Management.Automation.CredentialAttribute()]$Credential,

        [switch]$Force,

        [switch]$AsJob,

        [ValidateRange(1,64)]
        [uint32]$ThrottleLimit = $WU_ThrottleLimit
    )

    begin {

        $sb = {

            if ($using:Force) {

                if ((Get-Service 'wuauserv').Status -eq 'Running') {Stop-Service 'wuauserv' -Force}

                Remove-Item 'C:\Windows\SoftwareDistribution' -Recurse -Force

                'SusClientId','PingID','AccountDomainSid','SusClientIDValidation' | ForEach-Object {

                    if (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate' -Name $_ -ErrorAction 'SilentlyContinue') {
        
                        Remove-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate' -Name $_ -Force
                    }
                }

                'LastWaitTimeout','DetectionStartTime','NextDetectionTime','AUState' | ForEach-Object {

                    if (Get-ItemProperty -Path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update' -Name $_ -ErrorAction 'SilentlyContinue') {
        
                        Remove-ItemProperty -Path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update' -Name $_ -Force
                    }
                }

                if ((Get-Service 'wuauserv').Status -eq 'Stopped') {Start-Service 'wuauserv'}

                if ((Get-Service BITS).Status -eq 'Stopped') {Start-Service 'BITS'}
            }

            gpupdate /force /target:computer > $null

            wuauclt /resetauthorization /detectnow

            Start-Sleep 60

            wuauclt /reportnow
        }
    }

    process {

        $Invoke_Params = @{ScriptBlock = $sb ; ThrottleLimit = $ThrottleLimit ; ComputerName = $Name ; AsJob = $AsJob}
    
        if ($Credential) {$Invoke_Params += @{Credential = $Credential}}

        Invoke-Command @Invoke_Params
    }
}

function Invoke-WsusCleanup {

    [CmdletBinding()]

    param(

        [parameter(Position = 0,Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string[]]$WsusServerName,

        [uint32]$WsusPort = 8530,

        [switch]$UseSsl
    )

    $ErrorActionPreference = 'Stop'

    foreach ($Server in $WsusServerName) {

        Write-Verbose "Connecting to WSUS server - $Server"
        $wsusServer = Get-WsusServer -Name  $Server -PortNumber $WsusPort -UseSsl:$UseSsl

        Write-Verbose "Starting cleanup on WSUS server - $WsusServerName"
        $WsusServer | Invoke-WsusServerCleanup -CleanupObsoleteUpdates -CleanupUnneededContentFiles -DeclineExpiredUpdates -CompressUpdates -DeclineSupersededUpdates -Confirm:$false

        $Subscription = $WsusServer.GetSubscription()

        Write-Verbose "Starting Synchronization on WSUS server - $Server"

        $Timer = [System.Diagnostics.Stopwatch]::StartNew()

        $Subscription.StartSynchronization()

         While ($Subscription.GetSynchronizationStatus() -ne 'NotProcessing') {
             
             Start-Sleep -Seconds 5

             Write-Verbose "Waiting for Synchronization to finish on WSUS server - $Server - Runtime $($Timer.Elapsed.Hours):$($Timer.Elapsed.Minutes):$($Timer.Elapsed.Seconds)"
         }
    }
}

function Get-WsusNeededUpdates {

    param(

        [parameter(Position = 0,Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$WsusServerName,

        [parameter(Position = 1,Mandatory=$true)]
        $GroupName,

        [switch]$Approve,

        [uint32]$WsusPort = 8530,

        [switch]$UseSsl
    )

    $wsusServer = Get-WsusServer -Name  $WsusServerName -PortNumber $WsusPort -UseSsl:$UseSsl

    $TargetGroup = ($wsusServer.GetComputerTargetGroups() | where name -eq $GroupName)

    if ($TargetGroup) {

        $ComputerScope = New-Object Microsoft.UpdateServices.Administration.ComputerTargetScope

        $ComputerScope.IncludeSubgroups = $true

        $ComputerScope.IncludeDownstreamComputerTargets = $true

        $NeededUnapprovedUpdates = $TargetGroup.GetComputerTargets($ComputerScope).GetUpdateInstallationInfoPerUpdate() | 

        Where-Object {$_.UpdateInstallationState -eq 'NotInstalled' -and $_.UpdateApprovalAction -eq 'NotApproved'} | 

        Select-Object UpdateId -Unique 
        
        if (($NeededUnapprovedUpdates | Measure-Object).Count -gt 0) {

            $NeededUnapprovedUpdates | ForEach-Object {
    
                if ($Approve) {Get-WsusUpdate -UpdateServer $wsusServer -UpdateId $_.UpdateId | Approve-WsusUpdate -Action Install -TargetGroupName $GroupName}

                else {Get-WsusUpdate -UpdateServer $wsusServer -UpdateId $_.UpdateId}
            }
        }

        else {Write-Warning "No updates to approve."}
    }

    else {Write-Warning "No computer group named '$GroupName'."}
}

function Get-PhysicalDiskUsage {

    param(
        
        [Parameter(position=0)][string[]]$ComputerName,
        [Parameter()][PSCredential][System.Management.Automation.CredentialAttribute()]$Credential,
        [Parameter()][validateset("KB","MB","GB","TB")][String]$Unit = 'GB'
    )

    $WmiParams = @{
        
        Namespace = "root/Microsoft/Windows/Storage" 
        Class = 'MSFT_PhysicalDisk'
    }

    if ($ComputerName) {$WmiParams += @{ComputerName = $ComputerName}}

    if ($Credential) {$WmiParams += @{Credential = $Credential}}

    Get-CimInstance @WmiParams | 

    Select-Object `
        PSComputerName,
        FriendlyName,
        Model,
        @{l="Size$($Unit.ToUpper())";e={[math]::Round($_.Size / "1$Unit", 2)}},
        @{l="Allocatedsize$($Unit.ToUpper())";e={[math]::Round($_.Allocatedsize / "1$Unit", 2)}},
        HealthStatus
}

## Export all functions and all aliases
Export-ModuleMember -Function * -Alias * 