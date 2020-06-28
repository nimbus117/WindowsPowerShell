### Module - zKodi

###################################################################################################

$Kodi_Host = 'halsey'
$Kodi_Port = 8080
$env:Kodi_TimeOut = 10
$env:Kodi_Uri = "http://$Kodi_Host`:$Kodi_Port/jsonrpc" # Build Uri - "http://tv:8888/jsonrpc"
$Kodi_Headers = @{'content-type' = 'application/json'} # Required headers for Kodi JSON-RPC

###################################################################################################

function Show-KodiReturnedValue {

    param([Parameter(Position=0)][Object]$Returned)

    if ($Returned.result) {
    
        $Result = ($Returned.result | gm -MemberType NoteProperty | where name -ne limits).Name
        if ($Result) {$Returned.result.$Result}
        else {$Returned.result}
    }
    elseif ($Returned.error) {$Returned.error.message}
    else {$Returned}

}

###################################################################################################

function Get-KodiActivePlayer {

    # Invoke Player.GetActivePlayers method - returns details of active player
    $JSON_GetActivePlayer = @{jsonrpc = "2.0" ; method = "Player.GetActivePlayers" ; id = 1} | ConvertTo-Json
    Invoke-RestMethod -Uri $env:Kodi_Uri -Method Post -Body $JSON_GetActivePlayer -Headers $Kodi_Headers -TimeoutSec $env:Kodi_TimeOut |
    Select-Object -ExpandProperty result -ErrorAction SilentlyContinue
}

function Get-KodiPlayerItem {

    # Invoke Player.GetItem method - returns currently playing media type and label
    $ActivePlayer = Get-KodiActivePlayer
    $JSON_GetItem = @{jsonrpc = "2.0" ; method = "Player.GetItem" ; params = @{playerid = $ActivePlayer.playerid} ; id = 1} | ConvertTo-Json
    Invoke-RestMethod -Uri $env:Kodi_Uri -Method Post -Body $JSON_GetItem -Headers $Kodi_Headers -TimeoutSec $env:Kodi_TimeOut |
    Select-Object -ExpandProperty result -ErrorAction SilentlyContinue
}

function Get-KodiPlayerProperty {

    param([parameter(Mandatory=$true)][ValidateSet('speed','percentage','time','totaltime')][Object[]]$Property)
    # Invoke Player.GetProperties method - return properties
    $ActivePlayer = Get-KodiActivePlayer
    $params = @{playerid = $ActivePlayer.playerid ; properties = $Property}
    $JSON_GetProperties_Speed = @{jsonrpc = "2.0" ; method = "Player.GetProperties" ; params = $params ; id = 1} | ConvertTo-Json
    Invoke-RestMethod -Uri $env:Kodi_Uri -Method Post -Body $JSON_GetProperties_Speed -Headers $Kodi_Headers -TimeoutSec $env:Kodi_TimeOut |
    Select-Object -ExpandProperty result -ErrorAction SilentlyContinue
}

###################################################################################################

function Invoke-KodiPlayPause {

   try {
        $ActivePlayer = Get-KodiActivePlayer
        if ($ActivePlayer) {
            # Invoke Player.PlayPause method - plays or pauses the current player
            $JSON_PlayPause = @{jsonrpc = "2.0" ; method = "Player.PlayPause" ; params = @{playerid = $ActivePlayer.playerid} ; id = 1} | ConvertTo-Json
            $PlayPause = Invoke-RestMethod -Uri $env:Kodi_Uri -Method Post -Body $JSON_PlayPause -Headers $Kodi_Headers -TimeoutSec $env:Kodi_TimeOut
            if ($PlayPause.result) {$PlayPause.result}
            elseif ($PlayPause.error) {Write-Warning "$($MyInvocation.MyCommand) - $($PlayPause.error.message)"}
        }
        else {Write-Warning "$($MyInvocation.MyCommand) - No Active Player"}
    }
    catch {Write-Warning "$($MyInvocation.MyCommand) - $($_.Exception.Message)" ; break}
}
Set-Alias kpp -Value Invoke-KodiPlayPause

function Invoke-KodiStop {

   try {
        $ActivePlayer = Get-KodiActivePlayer
        if ($ActivePlayer) {
            # Invoke Player.Stop method - stops the current player
            $JSON_Stop = @{jsonrpc = "2.0" ; method = "Player.Stop" ; params = @{playerid = $ActivePlayer.playerid} ; id = 1} | ConvertTo-Json
            $Stop = Invoke-RestMethod -Uri $env:Kodi_Uri -Method Post -Body $JSON_Stop -Headers $Kodi_Headers -TimeoutSec $env:Kodi_TimeOut
            if ($Stop.result) {$Stop.result}
            elseif ($Stop.error) {Write-Warning "$($MyInvocation.MyCommand) - $($Stop.error.message)"}
        }
        else {Write-Warning "$($MyInvocation.MyCommand) - No Active Player"}
    }
    catch {Write-Warning "$($MyInvocation.MyCommand) - $($_.Exception.Message)" ; break}
}
Set-Alias ksp -Value Invoke-KodiStop

function Invoke-KodiSeek {

    param([Parameter(Position=0)]$Value) # "smallforward", "smallbackward", "bigforward", "bigbackward", 0..100, @{time}

    try {
        $ActivePlayer = Get-KodiActivePlayer
        if ($ActivePlayer) {
            # Invoke Player.Seek method - jumps to specific point in media - beginning = 0, end = 100 
            $params = @{playerid = $ActivePlayer.playerid ; value = $Value}
            $JSON_Seek = @{jsonrpc = "2.0" ; method = "Player.Seek" ; params = $params ; id = 1} | ConvertTo-Json
            $Seek = Invoke-RestMethod -Uri $env:Kodi_Uri -Method Post -Body $JSON_Seek -Headers $Kodi_Headers -TimeoutSec $env:Kodi_TimeOut
            if ($Seek.result) {$Seek.result}
            elseif ($Seek.error) {Write-Warning "$($MyInvocation.MyCommand) - $($Seek.error.message)"}
        }
        else {Write-Warning "$($MyInvocation.MyCommand) - No Active Player"}
    }
    catch {Write-Warning "$($MyInvocation.MyCommand) - $($_.Exception.Message)" ; break}
}
Set-Alias ksk -Value Invoke-KodiSeek

function Invoke-KodiVolume {

    param([Parameter(Position=0)]$Value) # increment, decrement,0..100

    try {
        # Invoke Application.SetVolume - sets application volume to given value
        $params = @{volume = $Value}
        $JSON_SetVolume = @{jsonrpc = "2.0" ; method = "Application.SetVolume" ; params = $params ; id = 1} | ConvertTo-Json
        $SetVolume = Invoke-RestMethod -Uri $env:Kodi_Uri -Method Post -Body $JSON_SetVolume -Headers $Kodi_Headers -TimeoutSec $env:Kodi_TimeOut
        if ($SetVolume.result) {$SetVolume.result}
        elseif ($SetVolume.error) {Write-Warning "$($MyInvocation.MyCommand) - $($SetVolume.error.message)"}
    }
    catch {Write-Warning "$($MyInvocation.MyCommand) - $($_.Exception.Message)" ; break}
}
Set-Alias kvl -Value Invoke-KodiVolume

function Send-KodiNotification {

    param([string]$Title='Title',[string]$Message='Message')
    try {
        # Invoke GUI.ShowNotification - play item
        $params = @{title = $Title ; message = $Message}
        $JSON_ShowNotification = @{jsonrpc = "2.0" ; method = "GUI.ShowNotification" ; params = $params ; id = 1} | ConvertTo-Json
        $ShowNotification = Invoke-RestMethod -Uri $env:Kodi_Uri -Method Post -Body $JSON_ShowNotification -Headers $Kodi_Headers -TimeoutSec $env:Kodi_TimeOut
        if ($ShowNotification.result) {$ShowNotification.result}
        elseif ($ShowNotification.error) {Write-Warning "$($MyInvocation.MyCommand) - $($ShowNotification.error.message)"}
    }
    catch {Write-Warning "$($MyInvocation.MyCommand) - $($_.Exception.Message)" ; break}
}
Set-Alias ksn -Value Send-KodiNotification

function Get-KodiMovie {

    param(
        [Parameter(Position=0)]$FilterTitle,
        [Parameter(Position=1)][Object[]]$Properties = ("runtime","plot","year")
    )

    try {
        # Invoke VideoLibrary.GetMovies method - returns a list of movies filterd by value
        $Sort = @{order= "ascending" ; method = "title"}
        $Filter = @{operator = "contains" ; field = "title" ; value = "$FilterTitle"}
        $params = @{sort = $Sort ; filter = $Filter ; properties = $Properties}
        $JSON_GetMovies = @{jsonrpc = "2.0" ; method = "VideoLibrary.GetMovies" ; params = $params ; id = 1} | ConvertTo-Json
        $Movies = Invoke-RestMethod -Uri $env:Kodi_Uri -Method Post -Body $JSON_GetMovies -Headers $Kodi_Headers -TimeoutSec $env:Kodi_TimeOut
        if ($Movies.result) {$Movies.result.movies}
        elseif ($Movies.error) {Write-Warning "$($MyInvocation.MyCommand) - $($Movies.error.message)"}
    }
    catch {Write-Warning "$($MyInvocation.MyCommand) - $($_.Exception.Message)" ; break}
}
Set-Alias kgm -Value Get-KodiMovie

function Get-KodiTVShow {

    param(
        [Parameter(Position=0)]$FilterTitle,
        [Parameter(Position=1)][Object[]]$Properties = ("plot","episode")
    )

    try {
        # Invoke VideoLibrary.GetTVShows method - returns a list of TV Shows filterd by value
        $Sort = @{order= "ascending" ; method = "title"}
        $Filter = @{operator = "contains" ; field = "title" ; value = "$FilterTitle"}
        $params = @{sort = $Sort ; filter = $Filter ; properties = $Properties}
        $JSON_GetTVShows = @{jsonrpc = "2.0" ; method = "VideoLibrary.GetTVShows" ; params = $params ; id = 1} | ConvertTo-Json
        $TVShows = Invoke-RestMethod -Uri $env:Kodi_Uri -Method Post -Body $JSON_GetTVShows -Headers $Kodi_Headers -TimeoutSec $env:Kodi_TimeOut
        if ($TVShows.result) {$TVShows.result.tvshows}
        elseif ($TVShows.error) {Write-Warning "$($MyInvocation.MyCommand) - $($TVShows.error.message)"}
    }
    catch {Write-Warning "$($MyInvocation.MyCommand) - $($_.Exception.Message)" ; break}
}
Set-Alias kgtsh -Value Get-KodiTVShow

function Get-KodiTVSeason {

    param(
        [Parameter(Position=0,ValueFromPipelineByPropertyName=$true)][int]$TVShowId,
        [Parameter(Position=1)][Object[]]$Properties = ("episode","season","showtitle")
    )

    try {
        # Invoke VideoLibrary.GetSeasons method - returns a list of seasons for a tvshow
        $Sort = @{order= "ascending" ; method = "title"}
        $params = @{sort = $Sort ; tvshowid = $TVShowId ; properties = $Properties}
        $JSON_GetSeasons = @{jsonrpc = "2.0" ; method = "VideoLibrary.GetSeasons" ; params = $params ; id = 1} | ConvertTo-Json
        $Seasons = Invoke-RestMethod -Uri $env:Kodi_Uri -Method Post -Body $JSON_GetSeasons -Headers $Kodi_Headers -TimeoutSec $env:Kodi_TimeOut
        if ($Seasons.result) {$Seasons.result.seasons}
        elseif ($Seasons.error) {Write-Warning "$($MyInvocation.MyCommand) - $($Seasons.error.message)"}
    }
    catch {Write-Warning "$($MyInvocation.MyCommand) - $($_.Exception.Message)" ; break}
}
Set-Alias kgtse -Value Get-KodiTVSeason

function Get-KodiTVEpisode {

    param(
        [Parameter(Position=0)]$FilterTitle,
        [Parameter(Position=1)][Object[]]$Properties = ("runtime","plot","firstaired","showtitle","season")
    )

    try {
        # Invoke VideoLibrary.GetEpisodes method - returns a list of Episodes filterd by value
        $Sort = @{order= "ascending" ; method = "title"}
        $Filter = @{operator = "contains" ; field = "title" ; value = "$FilterTitle"}
        $params = @{sort = $Sort ; filter = $Filter ; properties = $Properties}
        $JSON_GetEpisodes = @{jsonrpc = "2.0" ; method = "VideoLibrary.GetEpisodes" ; params = $params ; id = 1} | ConvertTo-Json
        $Episodes = Invoke-RestMethod -Uri $env:Kodi_Uri -Method Post -Body $JSON_GetEpisodes -Headers $Kodi_Headers -TimeoutSec $env:Kodi_TimeOut
        if ($Episodes.result) {$Episodes.result.episodes}
        elseif ($Episodes.error) {Write-Warning "$($MyInvocation.MyCommand) - $($Episodes.error.message)"}
    }
    catch {Write-Warning "$($MyInvocation.MyCommand) - $($_.Exception.Message)" ; break}
}
Set-Alias kgte -Value Get-KodiTVEpisode

function Open-KodiID {

    param(
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true,ParameterSetName=1)][int]$MovieId,
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true,ParameterSetName=2)][int]$EpisodeId,
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true,ParameterSetName=3)][int]$AlbumId,
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true,ParameterSetName=4)][int]$SongId,
        [switch]$Resume = $true
    )

    try {
      
        if ($MovieId) {$Id = $MovieId ; $IdType = 'movieid'}
        elseif ($EpisodeId) {$Id = $EpisodeId ; $IdType = 'episodeid'}
        elseif ($AlbumId) {$Id = $AlbumId ; $IdType = 'albumid'}
        elseif ($SongId) {$Id = $SongId ; $IdType = 'songid'}

        # Invoke Player.Open - play item
        $params = @{item = @{"$IdType" = $Id} ; options = @{resume = [bool]$Resume}}
        $JSON_Open = @{jsonrpc = "2.0" ; method = "Player.Open" ; params = $params ; id = 1} | ConvertTo-Json
        $Open = Invoke-RestMethod -Uri $env:Kodi_Uri -Method Post -Body $JSON_Open -Headers $Kodi_Headers -TimeoutSec $env:Kodi_TimeOut
        if ($Open.result) {$Open.result}
        elseif ($Open.error) {Write-Warning "$($MyInvocation.MyCommand) - $($Open.error.message)"}
    }
    catch {Write-Warning "$($MyInvocation.MyCommand) - $($_.Exception.Message)" ; break}
}
Set-Alias koi -Value Open-KodiId

###################################################################################################

## Export all functions and all aliases
Export-ModuleMember -Function * -Alias *

###################################################################################################

# Invoke-RestMethod -Uri $env:Kodi_Uri

# http://kodi.wiki/view/JSON-RPC_API/v6#Player.Property.Name

# Thanks to 
# http://www.foo.co.za/xbmcs-json-rpc-api-really-pausing-a-video
# http://powershell.com/cs/forums/t/16905.aspx
# http://www.xbmcbrasil.net/archive/index.php?thread-1655.html
# http://forum.kodi.tv/showthread.php?tid=171843