### Module - zSQL

function Invoke-RegisteredServerQuery {
    
    param(
        [Parameter(Mandatory=$true)][string]$Query,
        [int]$Port = 1433,
        [int]$Timeout = 1000
    )

    Import-Module SqlServer -Force -ErrorAction Stop

    Get-ChildItem 'SQLSERVER:\SQLRegistration' -Recurse | 

    Where-Object {$_ -is [Microsoft.SqlServer.Management.RegisteredServers.RegisteredServer]} | 

    ForEach-Object {

        $TcpClient = New-Object System.Net.Sockets.TCPClient
        $StopWatch = New-Object System.Diagnostics.Stopwatch
        $TimeStamp = Get-Date
        $Connect = $TcpClient.BeginConnect($_.ServerName,$Port,$null,$null)
        $StopWatch.Start()
        while ($TcpClient.Connected -ne $true) {if ($StopWatch.ElapsedMilliseconds -ge $Timeout) {Break}}
        $StopWatch.Stop()
    
        if ($TcpClient.Connected) {Invoke-Sqlcmd -Query $Query -ConnectionString $_.ConnectionString}

        else {Write-Warning "Connection failed against '$($_.Name)'"}
    }
}

Set-Alias -Name RSQuery -Value Invoke-RegisteredServerQuery

## Export all functions and all aliases
Export-ModuleMember -Function * -Alias *