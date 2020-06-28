### Module - zAccounts

# Local

Function  Get-LocalGroupMember2 {

    param(
        
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$GroupName,

        [string]$ComputerName = 'LocalHost'
    )

    ([adsi]"WinNT://$ComputerName/$GroupName,group").Invoke('members')  | ForEach {

        $_.GetType().InvokeMember("ADSPath","GetProperty",$null,$_,$null) | Select-Object @{n='ComputerName';e={$ComputerName}},@{n='MemberName';e={$_.ToString().Replace('WinNT://','').Replace('/','\')}}
    
    } | Sort-Object MemberName
}

function Add-LocalGroupMember2 {

    param(
        
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$GroupName,
        
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [string[]]$MemberName,

        [string]$ComputerName = 'LocalHost'
    )

    process {
        
        foreach ($Member in $MemberName) {

            $Member = $Member -replace '\\','/'

            ([adsi]"WinNT://$ComputerName/$GroupName,group").Add("WinNT://$Member")
        }
    }
}

function Remove-LocalGroupMember2 {

    param(
        
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$GroupName,
        
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [string[]]$MemberName,

        [string]$ComputerName = 'LocalHost'
    )

    process {
        
        foreach ($Member in $MemberName) {

            $Member = $Member -replace '\\','/'

            ([adsi]"WinNT://$ComputerName/$GroupName,group").Remove("WinNT://$Member")
        }
    }
}

# Domain

function Get-ADLastLogon {

    param(
        
        [parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true)]
        [ValidateNotNullOrEmpty()]
        [System.Object]$ADObject,
        [parameter()]
        [ValidateNotNullOrEmpty()]
        $DCFilter = '*'
    )

    begin {if (-not (Get-Module ActiveDirectory)) {throw "Requires ActiveDirectory module."}}

    process {

        Get-ADDomainController -Filter $DCFilter | 

        ForEach-Object {$ADObject | Get-ADObject -Properties LastLogon,LastLogonTimeStamp -Server $_.Name} | 

        Measure-Object -Maximum -Property LastLogon,LastLogonTimeStamp | 

        Measure-Object -Maximum -Property Maximum |

        Select-Object @{l='SamAccountName';e={$ADObject.SamAccountName}},@{l='LastLogon';e={[DateTime]::FromFileTime($_.Maximum)}}
    }
}

# SID

function Get-SIDFromNTAccount {

    param([parameter(position=0,mandatory=$true)][ValidateNotNullOrEmpty()][string]$NTAccount)

    New-Object PSCustomObject -Property @{
        
        NTAccount = $NTAccount
        SID = (New-Object System.Security.Principal.NTAccount("$NTAccount")).Translate([System.Security.Principal.SecurityIdentifier]).Value
    }
}

function Get-NTAccountFromSID {

    param([parameter(position=0,mandatory=$true)][ValidateNotNullOrEmpty()][string]$SID)

    New-Object PSCustomObject -Property @{
        
        SID = $SID
        NTAccount = (New-Object System.Security.Principal.SecurityIdentifier("$SID")).Translate([System.Security.Principal.NTAccount]).Value
    }
}

## Export all functions and all aliases
Export-ModuleMember -Function * -Alias * 