function New-VMPrompt {

    [CmdletBinding()]

    param(
        
        [switch]$SetNameAndIP,
        [switch]$SetLocalAdminPassword,
        [switch]$JoinDomain,
        [switch]$NewRDGCustomer
    )

    $ErrorActionPreference = 'Stop'

    while (-not$Name) {
            
        $Name = Read-Host "Enter VM name"

        if ($Name) {if ($Check = Get-VM $Name -ErrorAction SilentlyContinue) {Write-Warning "$($Check.Name) already exists" ; Clear-Variable Name}}
    }

    while (-not$Template) {$Template = (Get-Template | Out-GridView -OutputMode Single -Title "Select Template")}

    while (-not$Datastore) {
    
        $Datastore = (Get-Datastore | 
        Select-Object Name,@{l="NumberVMs";e={($_ | Get-VM).Count}},FreeSpaceGB,CapacityGB | 
        Sort-Object @{e={[int]$_.FreeSpaceGB};d={$true}} | 
        Out-GridView -OutputMode Single -Title "Select Datastore")
    }
    $Datastore = Get-datastore $Datastore.name

    while (-not$VLan) {$VLan = (Get-VirtualPortGroup | Sort-Object -Unique | Out-GridView -PassThru -Title "Select Network(s)")}

    while (-not$ResourcePool) {$ResourcePool = (Get-ResourcePool | Out-GridView -OutputMode Single -Title "Select Resource Pool")}

    $VMHost = (Get-VMHost | Select-Object name,@{l="NumberVMs";e={($_ | Get-VM).Count}},NumCpu,CpuUsageMhz,MemoryUsageGB,MemoryTotalGB | Out-GridView -OutputMode Single -Title "Select host")
    $VMHost = get-vmhost $VMhost.name

    $Tags = (Get-Tag | Out-GridView -PassThru -Title "Select tag(s)")

    $Location = (Get-Folder -Type VM | Out-GridView -OutputMode Single -Title "Select folder")

    $VMDetails = [PSCustomObject]@{

        Name = $Name
        Template = $Template
        Datastore = $Datastore
        VLan = $VLan
        ResourcePool = $ResourcePool
        VMhost = $VMHost
        Tags = $Tags
        Location = $Location
    }

    if ($SetNameAndIP -or $SetLocalAdminPassword -or $JoinDomain) {

        $Settings = @{}

        $GuestCreds = Get-Credential "Administrator" -Message "Local Admin Account"

        $Settings += @{GuestCreds = $GuestCreds}
    }

    if ($SetNameAndIP) {
        
        $IP = Read-Host "Enter IP"
        while (-not(($IP -as [ipaddress]) -as [bool])) {$IP = Read-Host "Invalid IP, re-enter"}
        $Mask = Read-Host "Enter Subnet Mask"
        while (-not(($Mask -as [ipaddress]) -as [bool])) {$Mask = Read-Host "Invalid Subnet Mask, re-enter"}
        $Gateway = Read-Host "Enter Default Gateway"
        while (-not(($Gateway -as [ipaddress]) -as [bool])) {$Gateway = Read-Host "Invalid Default Gateway, re-enter"}
        $Hostname = Read-Host "Enter Hostname"
        while (($Hostname.Length -lt 1) -or ($Hostname.Length -gt 15)) {$Hostname = Read-Host "Hostname must be between 1-15 characters, re-enter"}

        $Settings += @{IP = $IP ; Mask = $Mask ; Gateway = $Gateway ; Hostname = $Hostname}
    }

    if ($SetLocalAdminPassword) {
        
        $NewGuestCreds = Get-Credential "Administrator" -Message "New local admin account"

        $Settings += @{NewGuestCreds = $NewGuestCreds}
    }

    if ($JoinDomain) {

        $DomainCreds = Get-Credential "zonalconnect\" -Message "Zonalconnect domain details"

        $Settings += @{DomainCreds = $DomainCreds}
    }

    $VMDetails | New-ViVMFromTemplate

    if ($SetNameAndIP -or $SetLocalAdminPassword -or $JoinDomain) {
        
        $VM = Get-VM $VMDetails.Name

        $Settings += @{VM = $VM}
    }

    $SettingsObj = [PSCustomObject]$Settings

    if ($SetNameAndIP) {$SettingsObj | Set-NameAndIPAddress}

    if ($SetLocalAdminPassword) {

        $SettingsObj | Set-LocalAdminPassword

        $SettingsObj.GuestCreds = $SettingsObj.NewGuestCreds
    }

    if ($JoinDomain) {$SettingsObj | Join-Domain}
}

###########################

function New-ViVMFromTemplate {
    
    [CmdletBinding()]

    param(

        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Name,

        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [VMware.VimAutomation.ViCore.Types.V1.Inventory.Template]
        $Template,

        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [VMware.VimAutomation.ViCore.Types.V1.DatastoreManagement.Datastore]
        $Datastore,

        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [VMware.VimAutomation.ViCore.Types.V1.Host.Networking.VirtualPortGroupBase[]]
        $VLan,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [VMware.VimAutomation.ViCore.Types.V1.Inventory.VIContainer]
        $ResourcePool,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [VMware.VimAutomation.ViCore.Types.V1.Inventory.VMHost]
        $VMHost,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [VMware.VimAutomation.ViCore.Types.V1.Tagging.Tag[]]
        $Tags,

        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [VMware.VimAutomation.ViCore.Types.V1.Inventory.Folder]
        $Location
    )

    begin {
        
        Write-Verbose "Invoking - $($MyInvocation.InvocationName)"
        
        $ErrorActionPreference = 'Stop'
    }

    process {

        $Params = @{

            Name = $Name
            Template = $Template  
            Datastore = $Datastore 
            confirm = $False
        }

        if ($VMHost) {$Params += @{VMhost = $VMHost}}
        if ($Location) {$Params += @{Location = $Location}}
        if ($ResourcePool) {$Params += @{ResourcePool = $ResourcePool}}
            
        try {
            
            Write-Verbose "Cloning new VM from Template - $($Template.Name)"
            $NewVM = New-VM @Params -ErrorAction Stop

            Write-Verbose "Setting network adapter - $($VLan.Name)"
            $VMNic = Get-NetworkAdapter -VM $NewVM
            Set-NetworkAdapter -NetworkAdapter $VMNic -NetworkName $VLan.Name -Confirm:$False

            if ($Tags) {
                
                Write-Verbose "Assigning Tags"
                $Tags | ForEach-Object {New-TagAssignment -Tag $_ -Entity $NewVM -confirm:$False}
            }

            Write-Verbose "Starting VM - $($NewVM.Name)"
            $NewVM | Start-VM -confirm:$False | Wait-tools

            Write-Verbose "Waiting for VMtools update"
            While (((Get-view $NewVM).Guest.ToolsVersionStatus) -ne "guestToolsCurrent") {sleep 5}

            Start-Sleep 5
        }
        catch {Write-Error $_}

        Clear-Variable NewVM,VMNic
    }
}

function Set-NameAndIPAddress {
    
    [CmdletBinding()]

    param(

        [Parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [VMware.VimAutomation.ViCore.Types.V1.Inventory.VirtualMachine]
        $VM,

        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [PSCredential][System.Management.Automation.CredentialAttribute()]
        $GuestCreds,

        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [ipaddress]
        $IP,

        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [ipaddress]
        $Mask,

        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [ipaddress]
        $Gateway,

        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [ValidateLength(1,15)]
        [string]
        $Hostname
    )

    Write-Verbose "Invoking - $($MyInvocation.InvocationName)"

    $ErrorActionPreference = 'Stop'

    $ScriptText = @'

        Param(
            [Parameter(Position=0)][string]$IP,
            [Parameter(Position=1)][string]$Mask,
            [Parameter(Position=2)][String]$Gateway,
            [Parameter(Position=3)][string]$Hostname
        )

        $ErrorActionPreference = "Stop"

        $NetworkAdapter = Get-WmiObject win32_networkadapterconfiguration | Where-Object {$_.IPEnabled -eq "True"}

        $NetworkAdapter.EnableStatic($IP, $Mask)

        $NetworkAdapter.SetGateways($Gateway, 1)

        $Rename = Get-WmiObject win32_computersystem
        $Rename.Rename("$Hostname")

        Restart-Computer -Force
'@

    $params = "$IP $Mask $Gateway $($Hostname.ToUpper())"
    
    try {

        Write-Verbose "Waiting for tools"
        $VM | Wait-Tools | Out-Null
        
        Write-Verbose "Copying Script to VM"
        Invoke-VMScript -VM $VM -ScriptText "'$ScriptText' | out-file -filepath C:\temp\Set-NameAndIPAddress.ps1" -GuestCredential $GuestCreds
        Write-Verbose "Running Script"
        try {Invoke-VMScript -VM $VM -ScriptText "Powershell.exe -executionpolicy bypass -File C:\temp\Set-NameAndIPAddress.ps1 $params" -GuestCredential $GuestCreds}
        catch {}
        Write-Verbose "Waiting for reboot"
        Start-Sleep 10
        Get-VM $VM | Wait-Tools
        Start-Sleep 10
        Write-Verbose "Deleting Script"
        Invoke-VMScript -VM $VM -ScriptText "Remove-Item c:\temp\Set-NameAndIPAddress.ps1" -GuestCredential $GuestCreds
    }
    catch {Write-Error $_}
}

function Set-LocalAdminPassword {

    [CmdletBinding()]

    param(

        [Parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [VMware.VimAutomation.ViCore.Types.V1.Inventory.VirtualMachine]
        $VM,

        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [PSCredential][System.Management.Automation.CredentialAttribute()]
        $GuestCreds,

        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [PSCredential][System.Management.Automation.CredentialAttribute()]
        $NewGuestCreds
    )

    Write-Verbose "Invoking - $($MyInvocation.InvocationName)"

    $ErrorActionPreference = 'Stop'

    [string]$NewGuestPassword = $NewGuestCreds.GetNetworkCredential().Password

    $ScriptText = @'

        Param(
            [Parameter(Position=0)][string]$NewGuestPassword
        )

        $ErrorActionPreference = "Stop"

        $user = [adsi]"WinNT://localhost/administrator,user"
        $user.SetPassword($NewGuestPassword)
        $user.SetInfo()
'@

    $params = "'$NewGuestPassword'"
    
    try {

        Write-Verbose "Waiting for tools"
        $VM | Wait-Tools | Out-Null
        
        Write-Verbose "Copying Script to VM"
        Invoke-VMScript -VM $VM -ScriptText "'$ScriptText' | out-file -filepath C:\temp\Set-LocalAdminPassword.ps1" -GuestCredential $GuestCreds
        Write-Verbose "Running Script"
        Invoke-VMScript -VM $VM -ScriptText "Powershell.exe -executionpolicy bypass -File C:\temp\Set-LocalAdminPassword.ps1 $params" -GuestCredential $GuestCreds -ErrorAction SilentlyContinue
        Write-Verbose "Deleting Script"
        Invoke-VMScript -VM $VM -ScriptText "Remove-Item c:\temp\Set-LocalAdminPassword.ps1" -GuestCredential $NewGuestCreds
    }
    catch {Write-Error $_}
}

function Join-Domain {

    [CmdletBinding()]

    param(

        [Parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [VMware.VimAutomation.ViCore.Types.V1.Inventory.VirtualMachine]
        $VM,

        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [PSCredential][System.Management.Automation.CredentialAttribute()]
        $GuestCreds,

        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [PSCredential][System.Management.Automation.CredentialAttribute()]
        $DomainCreds
    )
    
    Write-Verbose "Invoking - $($MyInvocation.InvocationName)"

    $ScriptText = @'

        Param(
            [Parameter(Position=0)]$DomainUser,
            [Parameter(Position=1)]$DomainPass
        )

        $ErrorActionPreference = "Stop"
  
        $DomainPass = ConvertTo-SecureString -String $DomainPass -asPlainText -Force

        $Credentials = New-object System.Management.Automation.PSCredential($DomainUser,$DomainPass)

        Add-computer -DomainName zonalconnect.local -restart -credential $Credentials -Force
'@

    $ErrorActionPreference = 'Stop'

    $DomainUser = $DomainCreds.username
    $DomainPass = $DomainCreds.GetNetworkCredential().password
    $params = "'$DomainUser' '$DomainPass'"

    try {

        Write-Verbose "Waiting for tools"
        $VM | Wait-Tools | Out-Null

        Write-Verbose "Copying Script to VM"
        Invoke-VMScript -VM $VM -ScriptText "'$ScriptText' | out-file -filepath C:\temp\Join-Domain.ps1" -GuestCredential $GuestCreds
        Write-Verbose "Running script"
        Invoke-VMScript -VM $VM -ScriptText "Powershell.exe -executionpolicy bypass -File C:\temp\Join-Domain.ps1 $params" -GuestCredential $GuestCreds
        #Write-Verbose "Deleting script"
        #Invoke-VMScript -VM $VM -ScriptText "Remove-Item c:\temp\Join-Domain.ps1" -GuestCredential $GuestCreds
    }
    catch {Write-Error $_}
}

function New-RDGCustomer {

}

function Activate-Windows {

    [CmdletBinding()]

    Param(

        [Parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Hostname,

        [Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $ProductKey
    )

    Write-Verbose "Invoking - $($MyInvocation.InvocationName)"

    if ($ProductKey) {slmgr.vbs $Hostname /ipk $ProductKey}
    else {slmgr.vbs $Hostname /ato}
}

function Check-WindowsActivation {

    [CmdletBinding()]

    Param(

        [Parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Hostname
    )

    Write-Verbose "Invoking - $($MyInvocation.InvocationName)"

    Get-CimInstance -ClassName SoftwareLicensingProduct -ComputerName $Hostname |
    Where PartialProductKey | Select-Object Pscomputername,Name,@{l='LicenseStatus';e={
        switch ($_.LicenseStatus)
        {
            0 {'Unlicensed'}
            1 {'licensed'}
            2 {'OOBGrace'}
            3 {'OOTGrace'}
            4 {'NonGenuineGrace'}
            5 {'Notification'}
            6 {'ExtendedGrace'}
            Default {'Undetected'}
        }
    }
}
}