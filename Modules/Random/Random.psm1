### Module - Random

function Get-NetStat {
<#
.SYNOPSIS
    This function will get the output of netstat -n and parse the output
.DESCRIPTION
    This function will get the output of netstat -n and parse the output
.LINK
    http://www.lazywinadmin.com/2014/08/powershell-parse-this-netstatexe.html
#>
    PROCESS
    {
        # Get the output of netstat
        $data = netstat -n
        
        # Keep only the line with the data (we remove the first lines)
        $data = $data[4..$data.count]
        
        # Each line need to be splitted and get rid of unnecessary spaces
        foreach ($line in $data)
        {
            # Get rid of the first whitespaces, at the beginning of the line
            $line = $line -replace '^\s+', ''
            
            # Split each property on whitespaces block
            $line = $line -split '\s+'
            
            # Define the properties
            $properties = @{
                Protocole = $line[0]
                LocalAddressIP = ($line[1] -split ":")[0]
                LocalAddressPort = ($line[1] -split ":")[1]
                ForeignAddressIP = ($line[2] -split ":")[0]
                ForeignAddressPort = ($line[2] -split ":")[1]
                State = $line[3]
            }
            
            # Output the current line
            New-Object -TypeName PSObject -Property $properties
        }
    }
}

function Get-ComObject {

# https://gallery.technet.microsoft.com/Get-ComObject-Function-to-50a92047

    [CmdletBinding(DefaultParameterSetName='Filter')]
     
    param(
        [Parameter(ParameterSetName='Filter',Mandatory=$true,Position=0)]
        [ValidateNotNullOrEmpty()]
        [string]$Filter,
        [Parameter(ParameterSetName='List')]
        [switch]$ListAll
    )
 
    $ComObjects = Get-ChildItem HKLM:\Software\Classes -ErrorAction SilentlyContinue | Where-Object {
        $_.PSChildName -match '^\w+\.\w+$' -and (Test-Path -Path "$($_.PSPath)\CLSID")
    } | Select-Object -ExpandProperty PSChildName
 
    if ($Filter) {$ComObjects | Where-Object {$_ -like $Filter}} 
    else {$ComObjects}
}

function Get-LoggedOnUser {

<#
.Synopsis
Queries a computer to check for interactive sessions

.DESCRIPTION
This script takes the output from the quser program and parses this to PowerShell objects

.NOTES   
Name: Get-LoggedOnUser
Author: Jaap Brasser
Version: 1.2.1
DateUpdated: 2015-09-23

.LINK
http://www.jaapbrasser.com

.PARAMETER ComputerName
The string or array of string for which a query will be executed

.EXAMPLE
.\Get-LoggedOnUser.ps1 -ComputerName server01,server02

Description:
Will display the session information on server01 and server02

.EXAMPLE
'server01','server02' | .\Get-LoggedOnUser.ps1

Description:
Will display the session information on server01 and server02
#>

    param(

        [Parameter(
            Position=0,ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true
        )]
        [alias('ComputerName')]
        [string[]]$Name = 'localhost'
    )

    begin {
        
        $ErrorActionPreference = 'Stop'
        
        Function Convert-IdleTimeStringToTimeSpan {
            # Thanks to kbgeoff

            param($IdleTime)
  
            $Days, $Hours, $Minutes = 0, 0, 0
        
            If ($IdleTime -eq 'none') {$null}

            Else {

                If ( $IdleTime -ne '.') {

                    If ( $IdleTime -like '*+*' ) {

                        $Days, $IdleTime = $IdleTime.Split('+')
                    }
      
                    If ( $IdleTime -like '*:*' ) {

                        $Hours, $Minutes = $IdleTime.Split(':')
                    }

                    Else {

                        $Minutes = $IdleTime
                    }
                }
            
            New-Timespan -Days $Days -Hours $Hours -Minutes $Minutes
               
            } 
        }
    }

    process {

        foreach ($Computer in $Name) {

            try {

                if (Test-Connection $Computer -Quiet -Count 1) {

                    quser /server:$Computer 2>&1 | Select-Object -Skip 1 | ForEach-Object {

                        $CurrentLine = $_.Trim() -Replace '\s+',' ' -Split '\s'
                    
                        $HashProps = @{
                            UserName = $CurrentLine[0]
                            ComputerName = $Computer
                        }

                        # If session is disconnected different fields will be selected
                        if ($CurrentLine[2] -eq 'Disc') {

                            $HashProps.SessionName = $null
                            $HashProps.Id = [int]$CurrentLine[1]
                            $HashProps.State = $CurrentLine[2]
                            $HashProps.IdleTime = Convert-IdleTimeStringToTimeSpan($CurrentLine[3])
                            $HashProps.LogonTime = Get-Date ($CurrentLine[4..6] -join ' ')
                            $HashProps.LogonTime = Get-date ($CurrentLine[4..($CurrentLine.GetUpperBound(0))] -join ' ')                        } 
                    
                        else {

                            $HashProps.SessionName = $CurrentLine[1]
                            $HashProps.Id = [int]$CurrentLine[2]
                            $HashProps.State = $CurrentLine[3]
                            $HashProps.IdleTime = Convert-IdleTimeStringToTimeSpan($CurrentLine[4])
                            $HashProps.LogonTime = Get-Date ($CurrentLine[5..($CurrentLine.GetUpperBound(0))] -join ' ')                        }

                        New-Object -TypeName PSCustomObject -Property $HashProps |
                        Select-Object -Property UserName,ComputerName,SessionName,Id,State,IdleTime,LogonTime,Error
                    }
                }

                else {throw "Failed Ping."}

            } 
            
            catch {

                New-Object -TypeName PSCustomObject -Property @{
                    
                    ComputerName = $Computer
                    Error = $_.Exception.Message
                } | Select-Object -Property UserName,ComputerName,SessionName,Id,State,IdleTime,LogonTime,Error
            }
        }
    }
}

function Disconnect-LoggedOnUser {

<#
.SYNOPSIS   
Function to disconnect a RDP session remotely
    
.DESCRIPTION 
This function provides the functionality to disconnect a RDP session remotely by providing the ComputerName and the SessionId
	
.PARAMETER ComputerName
This can be a single computername or an array where the RDP sessions will be disconnected

.PARAMETER Id
The Session Id that that will be disconnected

.NOTES   
Name: Disconnect-LoggedOnUser
Author: Jaap Brasser
DateUpdated: 2015-06-03
Version: 1.0
Blog: http://www.jaapbrasser.com

.LINK
http://www.jaapbrasser.com

.EXAMPLE   
. .\Disconnect-LoggedOnUser.ps1
    
Description 
-----------     
This command dot sources the script to ensure the Disconnect-LoggedOnUser function is available in your current PowerShell session

.EXAMPLE
Disconnect-LoggedOnUser -ComputerName server01 -Id 5

Description
-----------
Disconnect session id 5 on server01

.EXAMPLE
.\Get-LoggedOnUser.ps1 -ComputerName server01,server02 | Where-Object {$_.UserName -eq 'JaapBrasser'} | Disconnect-LoggedOnUser -Verbose

Description
-----------
Use the Get-LoggedOnUser script to gather the user sessions on server01 and server02. Where-Object filters out only the JaapBrasser user account and then disconnects the session by piping the results into Disconnect-LoggedOnUser while displaying verbose information.
#>
    param(
        [Parameter(
            Mandatory=$true,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true,
            Position=0
        )]
        [alias('ComputerName')]
        [string[]]$Name,
        [Parameter(
            Mandatory=$true,
            ValueFromPipelineByPropertyName=$true
        )]
        [int[]]$Id
    )

    begin {$ErrorActionPreference = 'Stop'}

    process {
        
        foreach ($Computer in $Name) {

            $Id | ForEach-Object {

                Write-Verbose "Attempting to disconnect session $Id on $Computer"

                try {

                    rwinsta $_ /server:$Computer
                    Write-Verbose "Session $Id on $Computer successfully disconnected"
                } 
                
                catch {Write-Warning "Error on $Computer, $($_.Exception.Message)"}
            }
        }
    }
}

function Invoke-Rick {

    iex (New-Object Net.WebClient).DownloadString("http://bit.ly/e0Mw9w")
}

function Invoke-CDDrive {

    # http://techibee.com/powershell/eject-or-close-cddvd-drive-using-powershellalternative-to-windows-media-objects/2176

    [CmdletBinding()]            
    param(            
    
        [parameter(mandatory=$true,position=0)][ValidateSet('Eject','Close')]$Action
    )

    try {

        $Diskmaster = New-Object -ComObject IMAPI2.MsftDiscMaster2
        $DiskRecorder = New-Object -ComObject IMAPI2.MsftDiscRecorder2
        $DiskRecorder.InitializeDiscRecorder($DiskMaster)
        
        if ($Action -eq 'Eject') {$DiskRecorder.EjectMedia()}
        
        elseif ($Action -eq 'Close') {$DiskRecorder.CloseTray()}            
    } 
    
    catch {Write-Error "Failed to operate the disk. Details : $_"}
}

function Win32Restart-Computer {
    <# 
        .SYNOPSIS
        Restarts one or more computers using the WMI Win32_OperatingSystem method.
        
        .DESCRIPTION
        Restarts, shuts down, logs off, or powers down one or more computers. This relies on WMI's Win32_OperatingSystem class. 
        Supports common parameters -verbose, -whatif, and -confirm.
        
        .PARAM ComputerName
        One or more computer names to operate against. Accepts pipeline input ByValue and ByPropertyName.
        
        .PARAM Action
        Can be Restart, LogOff, Shutdown, or PowerOff.
        
        .PARAM Force
        Force the action.

        .EXAMPLE
        'localhost','server1' | Win32Restart-Computer -action LogOff -whatif
    #>

    # http://windowsitpro.com/blog/advanced-functions-part-2-shouldprocess-your-script-cmdlets
    
    [CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact="High")]
    
    param (
        
        [parameter(Position=0,Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]       
        [string[]]$ComputerName,
        
        [parameter(Position=1,Mandatory=$true)]
        [ValidateSet("Restart","LogOff","Shutdown","PowerOff")]
        [string]$Action,

        [Parameter(Position=2)]
        [System.Management.Automation.CredentialAttribute()]
        [PSCredential]$Credential,
        
        [switch]$Force
    )
    
    BEGIN {
        
        # translate action to numeric value required by the method
        
        switch ($Action) {
            
            "Restart" {$ActionArg = 2 ; break}
            
            "LogOff" {$ActionArg = 0 ; break}
            
            "Shutdown" {$ActionArg = 1 ; break}
            
            "PowerOff" {$ActionArg = 8 ; break}
        
        }

        # to force, add 4 to the value
        
        if ($Force) {$ActionArg += 4}

        Write-Verbose "Action set to $Action, $ActionArg"
    }
 
    
    PROCESS {
        
        Write-Verbose "Attempting to connect to $ComputerName"

        # this is how we support -whatif and -confirm
        
        # which are enabled by the SupportsShouldProcess
        
        # parameter in the cmdlet bindnig
        

        $params = @{ComputerName = $ComputerName}

        if ($Credential) {$params += @{Credential = $Credential}}

        if ($pscmdlet.ShouldProcess($ComputerName)) {
            
            (Get-WmiObject win32_operatingsystem @params).Win32Shutdown($ActionArg)
        
        }
    }
}

function Invoke-Repadmin {repadmin.exe /showrepl * /csv | ConvertFrom-Csv}

function Get-ProductKey {
     <#   
    .SYNOPSIS   
        Retrieves the product key and OS information from a local or remote system/s.
         
    .DESCRIPTION   
        Retrieves the product key and OS information from a local or remote system/s. Queries of 64bit OS from a 32bit OS will result in 
        inaccurate data being returned for the Product Key. You must query a 64bit OS from a system running a 64bit OS.
        
    .PARAMETER Computername
        Name of the local or remote system/s.
         
    .NOTES   
        Author: Boe Prox
        Version: 1.1       
            -Update of function from http://powershell.com/cs/blogs/tips/archive/2012/04/30/getting-windows-product-key.aspx
            -Added capability to query more than one system
            -Supports remote system query
            -Supports querying 64bit OSes
            -Shows OS description and Version in output object
            -Error Handling
     
    .EXAMPLE 
     Get-ProductKey -Computername Server1
     
    OSDescription                                           Computername OSVersion ProductKey                   
    -------------                                           ------------ --------- ----------                   
    Microsoft(R) Windows(R) Server 2003, Enterprise Edition Server1       5.2.3790  bcdfg-hjklm-pqrtt-vwxyy-12345     
         
        Description 
        ----------- 
        Retrieves the product key information from 'Server1'
    #>         
    [cmdletbinding()]
    Param (
        [parameter(ValueFromPipeLine=$True,ValueFromPipeLineByPropertyName=$True)]
        [Alias("CN","__Server","IPAddress","Server")]
        [string[]]$Computername = $Env:Computername
    )
    Begin {   
        $map="BCDFGHJKMPQRTVWXY2346789" 
    }
    Process {
        ForEach ($Computer in $Computername) {
            Write-Verbose ("{0}: Checking network availability" -f $Computer)
            If (Test-Connection -ComputerName $Computer -Count 1 -Quiet) {
                Try {
                    Write-Verbose ("{0}: Retrieving WMI OS information" -f $Computer)
                    $OS = Get-WmiObject -ComputerName $Computer Win32_OperatingSystem -ErrorAction Stop                
                } Catch {
                    $OS = New-Object PSObject -Property @{
                        Caption = $_.Exception.Message
                        Version = $_.Exception.Message
                    }
                }
                Try {
                    Write-Verbose ("{0}: Attempting remote registry access" -f $Computer)
                    $remoteReg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine,$Computer)
                    If ($OS.OSArchitecture -eq '64-bit') {
                        $value = $remoteReg.OpenSubKey("SOFTWARE\Microsoft\Windows NT\CurrentVersion").GetValue('DigitalProductId4')[0x34..0x42]
                    } Else {                        
                        $value = $remoteReg.OpenSubKey("SOFTWARE\Microsoft\Windows NT\CurrentVersion").GetValue('DigitalProductId')[0x34..0x42]
                    }
                    $ProductKey = ""  
                    Write-Verbose ("{0}: Translating data into product key" -f $Computer)
                    for ($i = 24; $i -ge 0; $i--) { 
                      $r = 0 
                      for ($j = 14; $j -ge 0; $j--) { 
                        $r = ($r * 256) -bxor $value[$j] 
                        $value[$j] = [math]::Floor([double]($r/24)) 
                        $r = $r % 24 
                      } 
                      $ProductKey = $map[$r] + $ProductKey 
                      if (($i % 5) -eq 0 -and $i -ne 0) { 
                        $ProductKey = "-" + $ProductKey 
                      } 
                    }
                } Catch {
                    $ProductKey = $_.Exception.Message
                }        
                $object = New-Object PSObject -Property @{
                    Computername = $Computer
                    ProductKey = $ProductKey
                    OSDescription = $os.Caption
                    OSVersion = $os.Version
                } 
                $object.pstypenames.insert(0,'ProductKey.Info')
                $object
            } Else {
                $object = New-Object PSObject -Property @{
                    Computername = $Computer
                    ProductKey = 'Unreachable'
                    OSDescription = 'Unreachable'
                    OSVersion = 'Unreachable'
                }  
                $object.pstypenames.insert(0,'ProductKey.Info')
                $object                           
            }
        }
    }
}

function Show-Object {

#############################################################################
##
## Show-Object
##
## From Windows PowerShell Cookbook (O'Reilly)
## by Lee Holmes (http://www.leeholmes.com/guide)
##
##############################################################################

<#

.SYNOPSIS

Provides a graphical interface to let you explore and navigate an object.


.EXAMPLE

PS > $ps = { Get-Process -ID $pid }.Ast
PS > Show-Object $ps

#>

param(
    ## The object to examine
    [Parameter(ValueFromPipeline = $true)]
    $InputObject
)

Set-StrictMode -Version 3

Add-Type -Assembly System.Windows.Forms

## Figure out the variable name to use when displaying the
## object navigation syntax. To do this, we look through all
## of the variables for the one with the same object identifier.
$rootVariableName = dir variable:\* -Exclude InputObject,Args |
    Where-Object {
        $_.Value -and
        ($_.Value.GetType() -eq $InputObject.GetType()) -and
        ($_.Value.GetHashCode() -eq $InputObject.GetHashCode())
}

## If we got multiple, pick the first
$rootVariableName = $rootVariableName| % Name | Select -First 1

## If we didn't find one, use a default name
if(-not $rootVariableName)
{
    $rootVariableName = "InputObject"
}

## A function to add an object to the display tree
function PopulateNode($node, $object)
{
    ## If we've been asked to add a NULL object, just return
    if(-not $object) { return }

    ## If the object is a collection, then we need to add multiple
    ## children to the node
    if([System.Management.Automation.LanguagePrimitives]::GetEnumerator($object))
    {
        ## Some very rare collections don't support indexing (i.e.: $foo[0]).
        ## In this situation, PowerShell returns the parent object back when you
        ## try to access the [0] property.
        $isOnlyEnumerable = $object.GetHashCode() -eq $object[0].GetHashCode()

        ## Go through all the items
        $count = 0
        foreach($childObjectValue in $object)
        {
            ## Create the new node to add, with the node text of the item and
            ## value, along with its type
            $newChildNode = New-Object Windows.Forms.TreeNode
            $newChildNode.Text = "$($node.Name)[$count] = $childObjectValue : " +
                $childObjectValue.GetType()

            ## Use the node name to keep track of the actual property name
            ## and syntax to access that property.
            ## If we can't use the index operator to access children, add
            ## a special tag that we'll handle specially when displaying
            ## the node names.
            if($isOnlyEnumerable)
            {
                $newChildNode.Name = "@"
            }

            $newChildNode.Name += "[$count]"
            $null = $node.Nodes.Add($newChildNode)               

            ## If this node has children or properties, add a placeholder
            ## node underneath so that the node shows a '+' sign to be
            ## expanded.
            AddPlaceholderIfRequired $newChildNode $childObjectValue

            $count++
        }
    }
    else
    {
        ## If the item was not a collection, then go through its
        ## properties
        foreach($child in $object.PSObject.Properties)
        {
            ## Figure out the value of the property, along with
            ## its type.
            $childObject = $child.Value
            $childObjectType = $null
            if($childObject)
            {
                $childObjectType = $childObject.GetType()
            }

            ## Create the new node to add, with the node text of the item and
            ## value, along with its type
            $childNode = New-Object Windows.Forms.TreeNode
            $childNode.Text = $child.Name + " = $childObject : $childObjectType"
            $childNode.Name = $child.Name
            $null = $node.Nodes.Add($childNode)

            ## If this node has children or properties, add a placeholder
            ## node underneath so that the node shows a '+' sign to be
            ## expanded.
            AddPlaceholderIfRequired $childNode $childObject
        }
    }
}

## A function to add a placeholder if required to a node.
## If there are any properties or children for this object, make a temporary
## node with the text "..." so that the node shows a '+' sign to be
## expanded.
function AddPlaceholderIfRequired($node, $object)
{
    if(-not $object) { return }

    if([System.Management.Automation.LanguagePrimitives]::GetEnumerator($object) -or
        @($object.PSObject.Properties))
    {
        $null = $node.Nodes.Add( (New-Object Windows.Forms.TreeNode "...") )
    }
}

## A function invoked when a node is selected.
function OnAfterSelect
{
    param($Sender, $TreeViewEventArgs)

    ## Determine the selected node
    $nodeSelected = $Sender.SelectedNode

    ## Walk through its parents, creating the virtual
    ## PowerShell syntax to access this property.
    $nodePath = GetPathForNode $nodeSelected

    ## Now, invoke that PowerShell syntax to retrieve
    ## the value of the property.
    $resultObject = Invoke-Expression $nodePath
    $outputPane.Text = $nodePath

    ## If we got some output, put the object's member
    ## information in the text box.
    if($resultObject)
    {
        $members = Get-Member -InputObject $resultObject | Out-String       
        $outputPane.Text += "`n" + $members
    }
}

## A function invoked when the user is about to expand a node
function OnBeforeExpand
{
    param($Sender, $TreeViewCancelEventArgs)

    ## Determine the selected node
    $selectedNode = $TreeViewCancelEventArgs.Node

    ## If it has a child node that is the placeholder, clear
    ## the placeholder node.
    if($selectedNode.FirstNode -and
        ($selectedNode.FirstNode.Text -eq "..."))
    {
        $selectedNode.Nodes.Clear()
    }
    else
    {
        return
    }

    ## Walk through its parents, creating the virtual
    ## PowerShell syntax to access this property.
    $nodePath = GetPathForNode $selectedNode 

    ## Now, invoke that PowerShell syntax to retrieve
    ## the value of the property.
    Invoke-Expression "`$resultObject = $nodePath"

    ## And populate the node with the result object.
    PopulateNode $selectedNode $resultObject
}

## A function to handle keypresses on the form.
## In this case, we capture ^C to copy the path of
## the object property that we're currently viewing.
function OnKeyPress
{
    param($Sender, $KeyPressEventArgs)

    ## [Char] 3 = Control-C
    if($KeyPressEventArgs.KeyChar -eq 3)
    {
        $KeyPressEventArgs.Handled = $true

        ## Get the object path, and set it on the clipboard
        $node = $Sender.SelectedNode
        $nodePath = GetPathForNode $node
        [System.Windows.Forms.Clipboard]::SetText($nodePath)

        $form.Close()
    }
}

## A function to walk through the parents of a node,
## creating virtual PowerShell syntax to access this property.
function GetPathForNode
{
    param($Node)

    $nodeElements = @()

    ## Go through all the parents, adding them so that
    ## $nodeElements is in order.
    while($Node)
    {
        $nodeElements = ,$Node + $nodeElements
        $Node = $Node.Parent
    }

    ## Now go through the node elements
    $nodePath = ""
    foreach($Node in $nodeElements)
    {
        $nodeName = $Node.Name

        ## If it was a node that PowerShell is able to enumerate
        ## (but not index), wrap it in the array cast operator.
        if($nodeName.StartsWith('@'))
        {
            $nodeName = $nodeName.Substring(1)
            $nodePath = "@(" + $nodePath + ")"
        }
        elseif($nodeName.StartsWith('['))
        {
            ## If it's a child index, we don't need to
            ## add the dot for property access
        }
        elseif($nodePath)
        {
            ## Otherwise, we're accessing a property. Add a dot.
            $nodePath += "."
        }

        ## Append the node name to the path
        $nodePath += $nodeName
    }

    ## And return the result
    $nodePath
}

## Create the TreeView, which will hold our object navigation
## area.
$treeView = New-Object Windows.Forms.TreeView
$treeView.Dock = "Top"
$treeView.Height = 500
$treeView.PathSeparator = "."
$treeView.Add_AfterSelect( { OnAfterSelect @args } )
$treeView.Add_BeforeExpand( { OnBeforeExpand @args } )
$treeView.Add_KeyPress( { OnKeyPress @args } )

## Create the output pane, which will hold our object
## member information.
$outputPane = New-Object System.Windows.Forms.TextBox
$outputPane.Multiline = $true
$outputPane.ScrollBars = "Vertical"
$outputPane.Font = "Consolas"
$outputPane.Dock = "Top"
$outputPane.Height = 300

## Create the root node, which represents the object
## we are trying to show.
$root = New-Object Windows.Forms.TreeNode
$root.Text = "$InputObject : " + $InputObject.GetType()
$root.Name = '$' + $rootVariableName
$root.Expand()
$null = $treeView.Nodes.Add($root)

## And populate the initial information into the tree
## view.
PopulateNode $root $InputObject

## Finally, create the main form and show it.
$form = New-Object Windows.Forms.Form
$form.Text = "Browsing " + $root.Text
$form.Width = 1000
$form.Height = 800
$form.Controls.Add($outputPane)
$form.Controls.Add($treeView)
$null = $form.ShowDialog()
$form.Dispose()
}

Function Send-NetMessage{ 
<#   
.SYNOPSIS   
    Sends a message to network computers 
  
.DESCRIPTION   
    Allows the administrator to send a message via a pop-up textbox to multiple computers 
  
.EXAMPLE   
    Send-NetMessage "This is a test of the emergency broadcast system.  This is only a test." 
  
    Sends the message to all users on the local computer. 
  
.EXAMPLE   
    Send-NetMessage "Updates start in 15 minutes.  Please log off." -Computername testbox01 -Seconds 30 -VerboseMsg -Wait 
  
    Sends a message to all users on Testbox01 asking them to log off.   
    The popup will appear for 30 seconds and will write verbose messages to the console.  
 
.EXAMPLE 
    ".",$Env:Computername | Send-NetMessage "Fire in the hole!" -Verbose 
     
    Pipes the computernames to Send-NetMessage and sends the message "Fire in the hole!" with verbose output 
     
    VERBOSE: Sending the following message to computers with a 5 delay: Fire in the hole! 
    VERBOSE: Processing . 
    VERBOSE: Processing MyPC01 
    VERBOSE: Message sent. 
     
.EXAMPLE 
    Get-ADComputer -filter * | Send-NetMessage "Updates are being installed tonight. Please log off at EOD." -Seconds 60 
     
    Queries Active Directory for all computers and then notifies all users on those computers of updates.   
    Notification stays for 60 seconds or until user clicks OK. 
     
.NOTES   
    Author: Rich Prescott   
    Blog: blog.richprescott.com 
    Twitter: @Rich_Prescott 
#> 
 
Param( 
    [Parameter(Mandatory=$True)] 
    [String]$Message, 
     
    [String]$Session="*", 
     
    [Parameter(ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)] 
    [Alias("Name")] 
    [String[]]$Computername=$env:computername, 
     
    [Int]$Seconds="5", 
    [Switch]$VerboseMsg, 
    [Switch]$Wait 
    ) 
     
Begin 
    { 
    Write-Verbose "Sending the following message to computers with a $Seconds second delay: $Message" 
    } 
     
Process 
    { 
    ForEach ($Computer in $ComputerName) 
        { 
        Write-Verbose "Processing $Computer" 
        $cmd = "msg.exe $Session /Time:$($Seconds)" 
        if ($Computername){$cmd += " /SERVER:$($Computer)"} 
        if ($VerboseMsg){$cmd += " /V"} 
        if ($Wait){$cmd += " /W"} 
        $cmd += " $($Message)" 
 
        Invoke-Expression $cmd 
        } 
    } 
End 
    { 
    Write-Verbose "Message sent." 
    } 
}

## Export all functions and all aliases
Export-ModuleMember -Function * -Alias *