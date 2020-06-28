### Module - zMisc

function Show-Colours {[Enum]::GetValues([System.ConsoleColor]) | Foreach-Object {Write-Host $_ -ForegroundColor $_}}

function Flash {
    $CurrentFG = $Host.UI.RawUI.BackgroundColor
    [Enum]::GetValues([System.ConsoleColor]) | 
    ForEach-Object {$Host.UI.RawUI.BackgroundColor=$_ ; cls ; Start-Sleep -Milliseconds 100} #Write-Host (" " * ($Host.UI.RawUI.WindowSize.Width - 1))
    $Host.UI.RawUI.BackgroundColor = $CurrentFG
    cls
}

function Get-InternalIP {

    try{Get-NetAdapter -Physical -ea Stop | Where-Object Status -eq Up | Get-NetIPAddress -AddressFamily IPv4 -ea Stop | Select-Object IPAddress} 
    catch{((ipconfig | sls IPv4) -split ' : ') | ForEach-Object {if (($_ -as [IPAddress]) -as [Bool]) {$_ | Select-Object @{l='IPAddress';e={$_}}}}}
}

function Get-ExternalIP {

    try {if (tPing ident.me 80 -Quiet) {Invoke-WebRequest ident.me -TimeoutSec 30 -UseBasicParsing | Select-Object @{l='IPAddress';e={$_.Content}}}} catch {}
}

function Get-IPGeoLocation {

    param([Parameter(ValueFromPipelineByPropertyName)][String]$IPAddress)

    Invoke-RestMethod http://ip-api.com/json/$IPAddress -TimeoutSec 10 | 
    
    Select-Object `
        @{l='Status';e={$_.status}},`
        @{l='IP';e={$_.query}},`
        @{l='City';e={$_.city}},`
        @{l='Country';e={$_.country}},`
        @{l='RegionName';e={$_.regionName}},`
        @{l='Region';e={$_.region}},`
        @{l='ZIP';e={$_.zip}},`
        @{l='TimeZone';e={$_.timeZone}},`
        @{l='Longitude';e={$_.lon}},`
        @{l='Latitude';e={$_.lat}},`
        @{l='ISP';e={$_.isp}},`
        @{l='ORG';e={$_.org}},`
        @{l='AS number';e={$_.as}}
}

Function Get-Weather {

    [CmdletBinding(DefaultParameterSetName="CityName")]

    param(
        [Parameter(Position=0,ValueFromPipelineByPropertyName)]
        [parameter(ParameterSetName="CityName")]
        $City,

        [parameter(ParameterSetName="GPS")]
        [ValidateRange(-180,180)]
        $Longitude,

        [parameter(ParameterSetName="GPS")]
        [ValidateRange(-90,90)]
        $Latitude,

        [parameter(HelpMessage="Get an API key at https://openweathermap.org")]
        $OpenWeatherMapAPIKey = "aa0a7819912626c3cb7c2df230b7f8c0",

        [parameter(HelpMessage="Get an API key at https://timezonedb.com")]
        $TimeZoneDbApiKey = "NRJKZURL5DQ7"
    )

    try {
        # Base openweathermap call with api key
        $URI = "http://api.openweathermap.org/data/2.5/weather?APPID=$OpenWeatherMapAPIKey"
    
        # zone used to convert local time (UTC)
        $UTCTimeZone = "Atlantic/Reykjavik"

        # Weather per city or GPS or GeoLoc if not specified
        IF ($City) {$URI += "&q=$City"}
        ELSEIF ($Longitude -and $Latitude) {$URI += "&lon=$Longitude&lat=$Latitude"}
    
        # Call to openweathermap
        $ApiResponse = Invoke-WebRequest -Uri $URI | ConvertFrom-Json
    
        $city = "$($ApiResponse.name) ($($ApiResponse.sys.country))"

        # Call Timezondb to get the local time info (Local time and offset from UTC)
        $LocalTime = Invoke-RestMethod -Method get -ContentType json -uri "http://api.timezonedb.com/v2/get-time-zone?key=$TimeZoneDbApiKey&format=json&by=position&lat=$($ApiResponse.coord.lat)&lng=$($ApiResponse.coord.lon)"

        [pscustomobject]@{
            City = $City
            Weather = $ApiResponse.weather.description
            Temperature = "$($ApiResponse.main.temp - 273.15) c"
            Humidity = "$($ApiResponse.main.humidity) %"
            Wind = "$($ApiResponse.wind.speed * 3.6) km/h"
            LocalTime = $LocalTime.formatted
            Sunrise = Get-date ((get-date 01/01/1970).AddSeconds($ApiResponse.sys.sunrise).AddSeconds($LocalTime.gmtOffset)) -Format T
            Sunset = Get-date ((get-date 01/01/1970).AddSeconds($ApiResponse.sys.sunset).AddSeconds($LocalTime.gmtOffset)) -Format T
        }
    } 

    catch {Write-Error $_.Exception -ErrorAction Stop}
} 

function Get-WeatherOld {

    param(
        [Parameter(Position=0,ValueFromPipelineByPropertyName,Mandatory)][ValidateNotNullOrEmpty()][string]$Country,
        [Parameter(Position=1,ValueFromPipelineByPropertyName)][string]$City
    )
        $Proxy = New-WebServiceProxy -uri http://www.webservicex.com/globalweather.asmx?WSDL
        if (!$City) {([xml]$Proxy.GetCitiesByCountry("$Country")).NewDataSet.Table | Sort-Object Country,City}
        else{([xml]$Proxy.GetWeather("$City","$Country")).CurrentWeather}
        $Proxy.Abort()

}

function Get-LocalWeather {Get-IPGeoLocation | Get-Weather}

function Get-WhoIs {

    param([parameter(Mandatory=$true,Position=0)][string]$DomainName)

    $WebProxy = New-WebServiceProxy ‘http://www.webservicex.net/whois.asmx?WSDL’
    $WebProxy.GetWhoIs($DomainName)

}
Set-Alias WhoIs -Value Get-WhoIs

function Invoke-Speech {

    param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Message,
        [ValidateNotNullOrEmpty()]
        [ValidateRange(0,100)]
        [int]$Volume = 100,
        [ValidateNotNullOrEmpty()]
        [ValidateRange(-10,10)]
        [int]$Rate,
        [ValidateNotNullOrEmpty()]
        [ValidateSet('Microsoft Hazel Desktop','Microsoft Zira Desktop')]
        [String]$Voice
    )

    begin {Add-Type -AssemblyName System.speech}
    process {

        $Say = New-Object System.Speech.Synthesis.SpeechSynthesizer

        $Say.Rate = $Rate
        $Say.Volume = $Volume
        if ($Voice) {$Say.SelectVoice($Voice)}
        $Say.Speak("$($Message | Out-String)")
        $Say.Dispose()
    }
}
Set-Alias say -Value Invoke-Speech

function Search-Google {

    param([parameter(Mandatory=$true,position=0,ValueFromPipeline=$true)][ValidateNotNullOrEmpty()][string[]]$SearchTerm)

    process {
    
        $SearchTerm | ForEach-Object {
            
            $Query = $_.Replace(" ","+")
            
            Start-Process http://google.com/search?q=$Query
        }
    }
}
Set-Alias google -Value Search-Google

function Invoke-WOL {

    param([string]$Mac = "94:de:80:2f:82:a2")

    $MacByteArray = $Mac -split "[:-]" | ForEach-Object { [Byte] "0x$_"}
    [Byte[]] $MagicPacket = (,0xFF * 6) + ($MacByteArray  * 16)
    $UdpClient = New-Object System.Net.Sockets.UdpClient
    $UdpClient.Connect(([System.Net.IPAddress]::Broadcast),9)
    $UdpClient.Send($MagicPacket,$MagicPacket.Length)
    $UdpClient.Close()
}
Set-Alias wol -Value Invoke-WOL

function Get-Info {

    # Get some information to display
    try {$Internal = Get-InternalIP} catch {}
    try {$External = Get-ExternalIP} catch {}
    $CS = Get-WmiObject Win32_ComputerSystem
    $CPU = Get-WmiObject Win32_Processor
    $CPUCount = ($CPU | Measure-Object).Count
    $CPU = $CPU | Select-Object -Unique @{l='Name';e={$_.Name -replace '\s+',' '}},@{l='CoreCount';e={$_.NumberOfCores * $CPUCount}},@{l='ProcCount';e={$_.NumberOfLogicalProcessors * $CPUCount}}
    $OS = Get-WmiObject win32_OperatingSystem 
    if ($OS.CSDVersion -eq $null) {$OSVersion = $OS.Caption} else {$OSVersion = "$($OS.Caption)($($OS.CSDVersion))"}
    $UpTime = New-TimeSpan -Start ([Management.ManagementDateTimeConverter]::ToDateTime($OS.LastBootUpTime)) -End (Get-Date)

    # Build display object
    [PSCustomObject]@{
        UpTime = "{0} {1}:{2}:{3}" -f $UpTime.Days, $UpTime.Hours.ToString("00"), $UpTime.Minutes.ToString("00"), $UpTime.Seconds.ToString("00")
        OSVersion = $OSVersion
        Model = "$($CS.Manufacturer) - $($CS.Model)"
        CPU = "$($CPU.Name) P:$($CPU.CoreCount) L:$($CPU.ProcCount)"
        MemoryGB = “{0:N2}” -f ($OS.TotalVisibleMemorySize/1MB)
        InternalIP = $Internal.IPAddress
        ExternalIP = $External.IPAddress
    }
}

function Get-WinException {

    param([parameter(mandatory=$true,position=0,ValueFromPipeline=$true)][ValidateNotNullOrEmpty()][int]$ErrorCode)

    ([ComponentModel.Win32Exception]$ErrorCode).Message

}
Set-Alias winex -Value Get-WinException

function Get-ProxyCommand {

    param([parameter(position=0,mandatory=$true)][ValidateNotNullOrEmpty()][string]$Command)

    $MetaData = New-Object system.management.automation.commandmetadata (Get-Command $Command)
    [System.Management.Automation.ProxyCommand]::Create($MetaData)

}

function Get-ParameterAliases {

# http://blogs.technet.com/b/heyscriptingguy/archive/2011/01/15/weekend-scripter-discovering-powershell-cmdlet-parameter-aliases.aspx

    param([parameter(position=0,mandatory=$true)][ValidateNotNullOrEmpty()][string]$Command)

    (Get-Command $Command).Parameters.Values | Where-Object {$_.aliases} | Select-Object Name, Aliases
}

Function ConvertFrom-Bytes {

param(
        [Parameter(Position = 0,Mandatory = $True)]
        [ValidateNotNullOrEmpty()]
        [double]$Bytes,

        [Parameter(Position=1)]
        [ValidateSet('KB','MB','GB','TB')]
        [string]$Units = 'MB',

        [Parameter(Position=2)]
        [ValidateRange(0,10)]
        [int]$Round = 2
)

    [math]::Round(($Bytes / "1$Units"), $Round)
}
Set-Alias cfb -Value ConvertFrom-Bytes

function Set-TSSessionSettings {

    Param(

        [Parameter(Position=0,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [Alias('IPAddress','ComputerName')]
        [string[]]
        $Name = 'Localhost',
        [Parameter(Position=1)]
        [PSCredential][System.Management.Automation.CredentialAttribute()]
        $Credential,
        [ValidateSet('Disconnect','Terminate')]
        [string]
        $BrokenConnectionAction,
        [int]
        $ActiveSessionLimit,
        [int]
        $DisconnectedSessionLimit,
        [int]
        $IdleSessionLimit

    )

    begin {$ErrorActionPreference = 'Stop'}

    process {

        try {

            $Name | ForEach-Object {

                $ComputerName = $_

                $params = @{
                    
                    Namespace = 'Root/CIMV2/TerminalServices'
                    Class = 'Win32_TSSessionSetting'
                    ComputerName = $ComputerName
                }

                if ($Credential) {$params += @{Credential = $Credential}}

                $TSSessionSetting = Get-WmiObject @params | Where-Object {$_.TerminalName -like "*RDP*"}

                if ($BrokenConnectionAction -or $ActiveSessionLimit -or $DisconnectedSessionLimit -or $IdleSessionLimit) {
                
                    $TSSessionSetting.TimeLimitPolicy = 0
                    $TSSessionSetting.Put() | Out-Null
                }


                if ($BrokenConnectionAction) {

                    if ($BrokenConnectionAction -eq 'Disconnect') {$BCAction = 0}
                    else {$BCAction = 1}

                    $TSSessionSetting.BrokenConnectionPolicy = 0
                    $TSSessionSetting.Put() | Out-Null
                    $TSSessionSetting.BrokenConnection($BCAction) | Out-Null
                    
                }

                if ($ActiveSessionLimit) {$TSSessionSetting.TimeLimit('ActiveSessionLimit',($ActiveSessionLimit * 1000)) | Out-Null}

                if ($DisconnectedSessionLimit) {$TSSessionSetting.TimeLimit('DisconnectedSessionLimit',($DisconnectedSessionLimit * 1000)) | Out-Null}

                if ($IdleSessionLimit) {$TSSessionSetting.TimeLimit('IdleSessionLimit',($IdleSessionLimit * 1000)) | Out-Null}

                Get-WmiObject @params | Where-Object {$_.TerminalName -like "*RDP*"}
            }
        }
        
        catch {Write-Error "$ComputerName - $_"}
    }
}

function Edit-File {

    [CmdletBinding()]

    param(

        [parameter(position=0,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [Alias('Fullname')]
        [string[]]$Path
    )

    process {

        $Path | ForEach-Object {

            $TextEditors = "${env:ProgramFiles(x86)}\Notepad++\notepad++.exe", "$env:SystemRoot\System32\notepad.exe"
        
            $PsEditor = @("$env:SystemRoot\system32\WindowsPowerShell\v1.0\PowerShell_ISE.exe")
            
            if ($_ -like '*.ps*') {$TextEditors = $PsEditor += $TextEditors}

            foreach ($TextEditor in $TextEditors) {

                if (Test-Path $TextEditor) {

                    Write-Verbose "Using - $TextEditor"
                    
                    $params = @{FilePath = $TextEditor}

                    if ($_) {

                        if (Test-Path $_) {Write-Verbose "File exists - $_"}

                        else {
                            
                            Write-Verbose "Creating file - $_"
                            
                            New-Item -Path $_ -ItemType File | Out-Null    
                        }
                
                        $params += @{ArgumentList = $_}
                    }

                    Start-Process @params
                    
                    break
                }

                else {Write-Verbose "Can't find - $TextEditor"}
            }
        }
    }
}
Set-Alias edit -Value Edit-File

function Test-DcConnection {

    param(
        
        [string[]]$IPs = @('8.8.8.8','109.233.117.109','172.31.6.150','172.31.3.101','172.31.3.37'),
        [ValidateRange(100,60000)][Int]$Delay = 2000,
        [ValidateRange(100,60000)][Alias('w')][Int]$TimeOut = 1000
    )

    function icmpPinger {

        param([parameter(position=0)]$Destination)

        $TimeStamp = (Get-Date)
        try {$Ping = (New-Object System.Net.NetworkInformation.Ping).Send("$Destination",$TimeOut)}
        catch {Write-Verbose $_}
        finally {

            if ($Ping.Status -eq 'Success'){$Status = $true} else {$Status = $false}
            if ($Ping.RoundtripTime) {$Time = $Ping.RoundtripTime} else {$Time = 0}
        
            [int]$ID = $Destination.split('.')[-1]
        
            Write-Progress -Activity $Destination -Status $Status -Id $ID -PercentComplete ((!$Status -as [int]) * 100)
 
            [PSCustomObject]@{
                TimeStamp = $TimeStamp
                Destination = $Destination
                Response = $Status
                Time = $Time
            }
        }
    }

    while ($true) {
    
        $IPs | ForEach-Object {

            icmpPinger $_

        }

        Start-Sleep -Milliseconds $Delay
    }
}

function Get-AndInvokeHistory {
    
    param([uint32]$Count = 500)

    $ErrorActionPreference = 'Stop'

    try {

        $Command = Get-History -Count $Count | Out-GridView -OutputMode Single -Title 'Select command to run'
        
        if ($Command) {$Command | Invoke-History}
    }

    catch {Write-Error $_.Exception.Message}
}
Set-Alias -Name hr -Value Get-AndInvokeHistory

function Start-Steam {

    param([switch]$TenFoot)

    $Path = "C:\Program Files (x86)\Steam\Steam.exe"

    $Process_Params = @{FilePath = $Path ; ErrorAction = 'Stop'}

    if ($TenFoot) {$Process_Params += @{ArgumentList = '-tenfoot'}}

    Start-Process @Process_Params
}
Set-Alias -Name Steam -Value Start-Steam

function PopUp {
    
    param(
        
        [string]$Title = 'Title',
        [string]$Message = 'Message'
    )

    $PopUpResult = (New-Object -ComObject Wscript.Shell).Popup($Message,0,$Title,0x1)

    if ($PopUpResult -eq 1) {$true}
    else {$false}
}

function Get-WordOfTheDay {
    
    (Invoke-RestMethod "http://www.oed.com/rss/wordoftheday")[-1].Description.TrimStart("OED Word of the Day: ")
}
Set-Alias -Name wod -Value Get-WordOfTheDay

## Export all functions and all aliases
Export-ModuleMember -Function * -Alias *