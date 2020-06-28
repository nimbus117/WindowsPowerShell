<#

------------------------------------------------

Name: PowerVLI
Author: Xavier Avrillier
LAst Update: 17/02/2017
http://vxav.fr

------------------------------------------------

Log Insight API links : 
Global (no ingest) - TODO
    https://vmw-loginsight.github.io/
Ingest API
    http://pubs.vmware.com/log-insight-36/index.jsp#com.vmware.log-insight.developer.doc/GUID-DCD0CB8A-EB78-4112-8785-34E27D782E75.html

------------------------------------------------

CONSTRAINTS:

All queries have to be URL-encoded.

URL-encoded format for numeric operators
--------------
%3D     EQ
%21%3D  NE   
%3C     LT   
%3C%3D  LE   
%3E     GT 
%3E%3D  GE   

URL-encoded format for string operators
--------------
CONTAINS       If white space, treated as a phrase 
NOT_CONTAINS   If white space, treated as a phrase
HAS            Contains every string delimited by white space
NOT_HAS        Doesn't Contain every string delimited by white space
=~             MATCHES_REGEX
!=~            NOT_MATCHES_REGEX
LAST           Only used with timestamp. displays events in the last [timestamp] milliseconds
XXX/EXIST      Limit the result to event that contain the field XXX

Other
-------------
%20            Space
%2A            Star *

------------------------------------------------

#>

Function Connect-VLIServer {

<#

    .SYNOPSIS
        
        Establishes an API session with the vRealize Log Insight Server.

    .DESCRIPTION

        This command establishes an API session on the log insight server. The default session TTL is 30 minutes.
        If a session is established there is no output. The relevant informations of the session are stored in a global variable.
        The properties of the variable are the following:
            -VLIServer : IP or hostname of the targeted Log Insight server
            -SessionEndDate : Exact date at which the session expires
            -Bearer : Token used to authenticate every call to the API

    .EXAMPLE

        PS> Connect-VLIServer -Server 10.39.0.50 -Creds $DomainCredentials -SSLCheckIgnore

        This command will connect to the VLI server with a domain account with no certificate verification (self-signed).

    .EXAMPLE

        PS> Connect-VLIServer -Server 10.39.0.50 -Creds $LocalCredentials -AuthProvider "Local" -SSLCheckIgnore

        This command will connect to the VLI server with an account local to VLI with no certificate verification (self-signed)

#>

Param(
    [switch]
    $SSLCheckIgnore,

    [Parameter(Mandatory=$True,Position=0,ValueFromPipeline=$True)]
    [string]
    $VLIServer,
    
    [Parameter(Mandatory=$True,Position=1,ValueFromPipeline=$True)]
    [System.Management.Automation.PSCredential]
    $Creds = (get-credential),

    [ValidateSet("ActiveDirectory","Local")]
    [string]
    $AuthProvider = "ActiveDirectory"
)

TRY {

    ##### Certificate verification doesn't work with self-signed certificates (default).
    ##### Enabled by default, disabled is essentially like ignoring the warning in the browser. Be careful where you use it.

    IF ($SSLCheckIgnore) {
    add-type @" 
    using System.Net; 
    using System.Security.Cryptography.X509Certificates; 

    public class NoSSLCheckPolicy : ICertificatePolicy { 
        public NoSSLCheckPolicy() {} 
        public bool CheckValidationResult( 
            ServicePoint sPoint, X509Certificate cert, 
            WebRequest wRequest, int certProb) { 
            return true; 
        } 
    } 
"@ 
    [System.Net.ServicePointManager]::CertificatePolicy = new-object NoSSLCheckPolicy 
    }

    ##### The body of the rest call is not stored in a variable on purpose to avoid having the password in clear text somewhere. (logging, ...)
    ##### Hence the long command

    $RestURISession = ("https://"+$VLIServer+":9543/api/v1/sessions")

    $RESTSession = Invoke-RestMethod -Method Post -ContentType 'application/json' -Body (
        [ordered]@{
            username = $Creds.UserName
            password = $Creds.GetNetworkCredential().password
            provider = $AuthProvider
        } | ConvertTo-Json) -Uri $RestURISession


    ##### Variable that stores relevant informations about the REST session. Made global to be usable outside the scope of this function.

    $Global:RESTStore  = [pscustomobject]@{
        VLIServer      = $VLIServer
        SessionEndDate = (get-date).AddSeconds($RESTSession.ttl)
        Bearer         = @{Authorization = "Bearer $($RestSession.sessionId)"}
    }

    $RESTStore

} CATCH {

    Write-Error $_.Exception -ErrorAction stop

}
}


Function Get-VLIEvent {

<#

    .SYNOPSIS
        
        Query the vRealize Log Insight server for individual logs records.

    .PARAMETER ORDER

        Use the "Order" parameter to sort the log events over the queried period. If " | Sort-Object" is used it will only sort the returned objects.
        Example: MaxSample=5 Start May Finish JUNE
            Order parameter=ASC     MAY|>>>>>-----------------|JUNE
            | Sort-Object ASC       MAY|----------------->>>>>|JUNE

    .PARAMETER CONTAINS

        f the * character (star) is not appended to the constraint only absolute matches will be returned.
        Example:
            > Contains iscsi    returns only "dca-utl-vbr1 in-guest iscsi blah".
            > Contains iscsi*   returns "dca-utl-vbr1 in-guest iscsi blah" and "zts1 iscsid error".

    .PARAMETER NOTCONTAINS

        If the * character (star) is not appended to the constraint only absolute matches will be returned.
        Example:
            > Not_Contains iscsi    can return "zts1 iscsid error" but not "dca-utl-vbr1 in-guest iscsi blah".
            > Not_Contains iscsi*   won't return either of them.

    .PARAMETER FIELDEXIST
        
        Returns only events that contain the specified field (appname, VC_Username, Hostname ...).

    .PARAMETER TIMEOUTSECONDS
        
        By default the queries will timeout after 30 seconds. This parameter allows you to increase this limit.

    .PARAMETER OUTPUT
        
        This parameter controls the formatting of the returned object.
            Human will return an object with name properties, easier to read.
            Ingest will return an object ready to be ingested by the Log Insight API.

    .EXAMPLE

        PS> Get-VLIEvent -Start 01/01/2017 -MaxSample 5 -Contains "zts5*","zts7*" -Type error

        This command will return the first 5 events since 1st january 2017 containing zts5 or zts7 and error in the text.

#>

param(
    [Parameter(Position=0,ValueFromPipeline=$True)]
    [string]
    $VLIServer = ($RESTStore.VLIServer),

    $Start = "",

    $Finish = "",

    [ValidateRange(0,20000)]
    [int]
    $MaxSamples = 100,

    [ValidateSet("OLDEST","NEWEST")]
    [string]
    $Order = "NEWEST",

    [string[]]
    $Contains = "",

    [string[]]
    $NotContains = "",

    [string]
    $FieldExists = "",

    [ValidateSet("Error","Warning","Info","Verbose")]
    [string]
    $Type = "",

    [int]
    $TimeOutSeconds = 30,

    [ValidateSet("HUMAN","INGEST")]
    [string]
    $Output = "HUMAN"

)

TRY {

    ##### Check the existence and validity of the token
    IF ($RESTStore.SessionEndDate -lt (get-date) -or !$RESTStore) {Throw "API session expired or non existant, please (re)connect to a VLI server using the Connect-VLIServer command"}

    ##### Prepare the bits that will make the URI
    $RestURICall = ("https://"+$VLIServer+":9543/api/v1/events")
    $RESTDate    = Get-URIDate -Start $Start -Finish $Finish
    $RESTString  = Get-URIFilter -Contains $Contains -NotContains $NotContains -Type $Type -FieldExists $FieldExists
    
    ##### Manages the display type. simple returns a user friendly object with properties. Default returns a string that can be used to ingest back in the API.
    Switch ($Output) {
        "HUMAN" {$viewtype = "simple" ;$viewSelect = "results"}
        "INGEST"{$viewtype = "default";$viewSelect = "events"}
    }
    Switch ($Order) {
        "OLDEST" {$NewOrder = "ASC"}
        "NEWEST" {$NewOrder = "DESC"}
    }

    ##### Put the bits of the URI together
    $RestURICall = $RestURICall + $RESTDate + $RESTString + "?limit=$MaxSamples&order-by-direction=$NewOrder&view=$viewtype&timeout=" + ($TimeOutSeconds*1000)


    ##### Invoke the REST query to the VLI server
    $RESTResponse = Invoke-RestMethod -Method Get -Uri $RestURICall -Headers $RESTStore.Bearer

    $RESTResponse | select -ExpandProperty $viewSelect

    $RESTResponse | Check-ResponseComplete

} CATCH {

    Write-Error $_.Exception -ErrorAction stop

}

}



Function Check-ResponseComplete {


<# 

.DESCRIPTION
    If the API response is incomplete the error is printed on the console
    The response is still usable in a variable or in a pipe as write host is "just display"

.NOTES
    In case of incomplete API response the returned object contains a "warnings" property composed of:
        -ID of the error
        -DETAILS of the error (brief)
        -PROGRESS status of the query in fraction

#>

Param(
    [Parameter(Position=0,ValueFromPipeline=$True)]
    $RESTResponse
)

    IF ($RESTResponse.complete -eq $false) {

        Write-Host "Incomplete API Response WARNING ID:"$RESTResponse.warnings.id"-"$RESTResponse.warnings.details -ForegroundColor Yellow -BackgroundColor Black

    }

}



Function Get-URIDate {

<#

.NOTES
    The API uses timestamps in milliseconds to set date boundaries.

#>

Param(
    $Start,

    $Finish
)

    IF ($Start) {
        
        IF (!$Finish) {$Finish = Get-date} ELSE {$Finish = get-date $Finish}
        $Start = Get-date $Start
                
        IF ($Start -gt $Finish) {Throw "Start date must be older than finish date"}

        "/timestamp/%3E" + ([string][math]::round((New-TimeSpan -Start 01/01/1970 -End $Start).TotalMilliseconds,0)) + "/timestamp/%3C" + ([string][math]::round((New-TimeSpan -Start 01/01/1970 -End $Finish).TotalMilliseconds,0))

    } ELSEIF ($Finish) {

        "/timestamp/%3C" + ([string][math]::round((New-TimeSpan -Start 01/01/1970 -End $Finish).TotalMilliseconds,0))

    }

}






Function Get-URIFilter {

<#

.NOTES
    2%A is the URL notation of * (star). If it is not appended to the constraint only absolute matches will be returned
    Example:
        > Contains iscsi    returns only "dca-utl-vbr1 in-guest iscsi blah"
        > Contains iscsi* returns "dca-utl-vbr1 in-guest iscsi blah" and "zts1 iscsid error"

    Filters of the same type (CONTAINS, NOT_CONTAINS, HAS ...) are treated with an OR operator but treated with an AND operator with filters of a different type. 
    Example:
        > CONTAINS zts5,zts9;HAS error   =   ("zts5" OR "zts9") AND "error"

#>

Param(

    [string[]]
    $Contains,

    [string[]]
    $NotContains,

    [string]
    $Type,

    [string]
    $FieldExists

)

    $URIFilter = ""

    IF ($FieldExists) {$URIFilter += "/$FieldExists/EXISTS"}


    IF ($Type) {$URIFilter += "/text/HAS%20$Type%2A"}

    IF ($Contains) {
        $Contains | ForEach-Object {
            $URIFilter += ("/text/CONTAINS%20" + ($_ -replace " ","%20")).Replace('*',"%2A")
        }
    }

    IF ($NotContains) {
        $NotContains | ForEach-Object {
            $URIFilter += ("/text/NOT_CONTAINS%20" + ($_ -replace " ","%20")).Replace('*',"%2A")
        }
    }

    $URIFilter

}




###################################

###################################

###################################