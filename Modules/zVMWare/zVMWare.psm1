### Module - zVMWare

function Get-VMInfo {

    param(
        
        [Parameter(position=0,Mandatory=$true,ValueFromPipeLine=$true)]
        [ValidateNotNullOrEmpty()]
        [VMware.VimAutomation.ViCore.Types.V1.Inventory.VirtualMachine[]]$VM
    )

    process {
    
        $VM | ForEach-Object {

            $CurrentFolder = $_.Folder
        
            While ($CurrentFolder.name -ne "vm") {

                $Parent = Get-Folder $CurrentFolder | Select-Object Parent
                $Path = "\" + $CurrentFolder.name + $Path
                $CurrentFolder = $Parent.Parent
                if ($CurrentFolder.count -gt 0 ) {$CurrentFolder= $CurrentFolder[0]}
            }

            $_ | Select-Object `
                Name,
                @{l='HostName';e={$_.ExtensionData.Guest.HostName}},
                @{l='IpAddress';e={$_.Guest.IpAddress}},
                @{l='Folder';e={$Path}},
                @{l='Host';e={$_.VMHost.Name}},
                PowerState,
                @{l='OS';e={$_.Guest.OSFullName.Replace('Microsoft ','')}},
                @{l='Hardware';e={$_.ExtensionData.Config.Version}},
                @{l='Tools';e={$_.ExtensionData.Guest.ToolsStatus}},
                NumCpu,
                MemoryGB
        
            Clear-variable -Name path
        }
    }
}

function Get-VMDiskMap {

    <#
    .NOTES
        ######################
         mail@nimbus117.co.uk
        ######################

        Based on a script by NiTRo - http://www.hypervisor.fr/?p=5070
    
    .SYNOPSIS
        Map VMWare hard disks to Windows guest disks.

    .DESCRIPTION
        Get-VMDiskMap uses vSphere PowerCLI and WMI to map VMWare hard disks to Windows guest disks by matching UUID's.
        It requires vSphere PowerCLI and an established connection to a vCenter server.
        By default remote WMI queries are made over RPC.
        When the guest VM can not be reached over the network the UseVIX switch parameter allows for the WMI query to be run via VMWare Tools.
        The current session credentials will be used to authenticate against the guest VM whether using RPC or VIX.
        Alternative guest credentials can by specified using the GuestCreds parameter.

    .PARAMETER VMName
        VM Object.

    .PARAMETER GuestCreds
        Windows guest credentials.

    .PARAMETER UseVIX
        Connect to the Windows guest via VIX (VMware Tools).

    .EXAMPLE
        PS C:\>Get-VM VM1 | Get-VMDiskMap | Format-Table

        VMName VMScsiId VMDisk      WinDisk VMSize WinSize VMPath
        ------ -------  ------      ------- ------ ------- ------
        VM1    0:0      Hard disk 1 Disk 0      30      30 [DS1] VM1/VM1.vmdk
        VM1    1:0      Hard disk 2 Disk 2      20      20 [DS1] VM1/VM1_1.vmdk
        VM1    2:0      Hard disk 3 Disk 3      20      20 [DS1] VM1/VM1_2.vmdk
        VM1    3:15     Hard disk 4 Disk 1      10      10 [DS1] VM1/VM1_3.vmdk


        Description

        -----------

        This command maps VMWare hard disks to Windows disks for VM1 and displays the results as a table.

    .EXAMPLE
        PS C:\>$creds = Get-Credential
        PS C:\>Get-VM *sql* | Get-VMDiskMap -GuestCreds $creds


        Description

        -----------

        The first comand prompts for and saves credentials to the variable $creds.
        The seccond command maps VMWare hard disks to Windows disks for all VM's with sql in their name, using the GuestCreds parameter with the saved credentials.

    .EXAMPLE
        PS C:\>$VM = Get-VM VM2,VM3
        PS C:\>Get-VMDiskMap -VM $VM -UseVIX -GuestCreds domain\user | Out-GridView


        Description

        -----------

        The first command saves the VirtualMachine objects to the variable $VM.
        The second command maps VMWare hard disks to Windows disks for VM2 and VM3 then displays the results in a grid.
        The UseVIX parameter runs the WMI query using VMWare Tools. The GuestCreds parameter prompts for a password when a username is specified.
    #>

    param (
    
        [Parameter(Position=0,Mandatory=$true,ValueFromPipeline=$true)]
        [ValidateNotNullOrEmpty()]
        [VMware.VimAutomation.ViCore.Types.V1.Inventory.VirtualMachine[]]
        $VM,
        [Parameter(Position=1)]
        [PSCredential][System.Management.Automation.CredentialAttribute()]
        $GuestCreds,
        [Switch]
        $UseVIX
    )

    process {

        $VM | ForEach-Object {

            try {

                $VMDevice = ($_ | Get-View -ea Stop).Config.Hardware.Device

                if (!$VMDevice) {throw 'No VM hard disks returned.'}

                if ($UseVIX) {
        
                    $ScriptText = "powershell.exe -NoProfile -Command `"Get-WmiObject Win32_DiskDrive ^| Select-Object SerialNumber,Index,Size ^| ConvertTo-CSV -NoTypeInformation`""

                    $Invoke_VMScriptParams = @{ScriptText = $ScriptText ; ScriptType = 'Bat'}

	                if ($GuestCreds) {$Invoke_VMScriptParams += @{GuestCredential = $GuestCreds}}
        
                    $WinDisks = ($_ | Invoke-VMScript @Invoke_VMScriptParams -ea Stop).ScriptOutput | ConvertFrom-Csv
	            }

                else {

                    $Get_WmiObjectParams = @{Class = 'Win32_DiskDrive' ; ComputerName = $_.Guest.HostName}

                    if ($GuestCreds) {$Get_WmiObjectParams += @{Credential = $GuestCreds}}

                    $WinDisks = Get-WmiObject @Get_WmiObjectParams -ea Stop | Select-Object SerialNumber,Index,Size
	            }

                if (!$WinDisks) {throw 'No WMI data returned.'}

                foreach ($SCSIController in ($VMDevice | Where-Object {$_.DeviceInfo.Label -match "SCSI Controller"})) {

                    foreach ($VMDisk in ($VMDevice | Where-Object {$_.ControllerKey -eq $SCSIController.Key})) {

	                    $WinDisk = $WinDisks | Where-Object {$_.SerialNumber -eq $VMDisk.Backing.Uuid.Replace('-','')}

	                    if ($WinDisk) {

                            [PSCustomObject]@{

                                VMName = $_.Name
                                VMScsiId = "{0}:{1}" -f $SCSIController.BusNumber,$VMDisk.UnitNumber
		                        VMDisk = $VMDisk.DeviceInfo.label
		                        WinDisk = 'Disk {0}' -f $WinDisk.Index
		                        VMSize = [Math]::Round($VMDisk.CapacityInKB/1MB, 1)
		                        WinSize = [Math]::Round($WinDisk.Size/1GB, 1)
		                        VMPath = $VMDisk.Backing.FileName
                            }
		                }
                    }
	            }
            }
            catch {Write-Error $_}
        }
    }
}

function Get-VMFiles {

    param(
        [parameter(position=0,ValueFromPipeline=$true,mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [VMware.VimAutomation.ViCore.Types.V1.DatastoreManagement.Datastore[]]
        $DataStore,
        [parameter(position=1)]
        [int]
        $DaysSinceLastWrite = 0,
        [String]
        $Filter = '*'
    )

    process {

        $DataStore | ForEach-Object {

            $_ | Select-Object @{l='Path';e={$_.DatastoreBrowserPath}} | Get-ChildItem -Recurse -Filter $Filter | 
            Where-Object {($_.PSIsContainer -eq $false) -and (($_.LastWriteTime) -lt (Get-Date).AddDays("-$DaysSinceLastWrite"))}
        }
    }
}

function Start-VMRC {
    
    <#
    .NOTES
        ######################
         mail@nimbus117.co.uk
        ######################
    
    .SYNOPSIS
        
        Launches 'VMware Remote Console' (https://my.vmware.com/en/web/vmware/details?downloadGroup=VMRC800&productId=491).
        
    .PARAMETER Name
    
    .PARAMETER VM

    .PARAMETER Server

    .EXAMPLE
        PS C:\> Start-VMRC VM1

    .EXAMPLE
        PS C:\> Get-VM VM* | Start-VMRC

    .EXAMPLE
        PS C:\> $Server = Get-Item vis:\172.16.1.10@443\
        PS C:\> Start-VMRC -Name VM2 -Server $Server
    #>

    [CmdletBinding(DefaultParameterSetName='Name')]

    param(

        [Parameter(ParameterSetName='Name', Mandatory=$true, Position=0)]
        [Alias('VMName')]
        [ValidateNotNullOrEmpty()]
        [string[]]
        $Name,

        [Parameter(ParameterSetName='VMObject', Mandatory=$true, ValueFromPipeline=$true)]
        [ValidateNotNullOrEmpty()]
        [VMware.VimAutomation.ViCore.Types.V1.Inventory.VirtualMachine[]]
        $VM,

        [VMware.VimAutomation.ViCore.Types.V1.VIServer]
        $Server = $global:DefaultVIServer
    )

    begin {
            
            $ErrorActionPreference = 'Stop'

            $VMRCPath = "C:\Program Files (x86)\VMware\VMware Remote Console\vmrc.exe"
            if (-not(Test-Path $VMRCPath)) {throw "Cannot find $VMRCPath."}

            $Session = Get-View $Server.ExtensionData.Client.ServiceContent.SessionManager
        }
    

    process {

        if ($Name) {
        
            $GetVM = Get-VM -Name $Name -Server $Server -ErrorAction Continue
            if ($GetVM) {$VM = $GetVM}
            else {return}
        }

        $VM | ForEach-Object {

            try {
        
                $Ticket =  $Session.AcquireCloneTicket()
                $HostAddress = $Server.ServiceUri.Host
                $VMID = $_.ExtensionData.MoRef.Value
                $VMRCArgs = "vmrc://clone:$Ticket@$HostAddress/?moid=$VMID"

                Start-Process -FilePath $VMRCPath -ArgumentList $VMRCArgs
                $_
                Start-Sleep -Milliseconds 500
            }

            catch {Write-Error $_}
        }
    }
}
Set-Alias vmrc -Value Start-VMRC

function Regenerate-MacAddress {

    param(
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
        [ValidateNotNullOrEmpty()]
        [VMware.VimAutomation.ViCore.Types.V1.VirtualDevice.NetworkAdapter[]]
        $NetworkAdapter,
        [switch]$Confirm
    )

    begin {
            
            $ErrorActionPreference = 'Stop'
        }
    

    process {

        $NetworkAdapter | ForEach-Object {

            try {
        
                $_.ExtensionData.AddressType = 'Generated'
                $_.ExtensionData.MacAddress = ""
                Set-NetworkAdapter $_ -confirm:$Confirm
            }
            catch {Write-Error $_}
        }
    }
}

Function Get-QuickVMStats {

    param(
        [Parameter(Mandatory = $True,ValueFromPipeline=$True)]
        [ValidateNotNullOrEmpty()]
        [VMware.VimAutomation.ViCore.Types.V1.Inventory.VirtualMachine[]]
        $VM,

        [ValidateNotNullOrEmpty()]
        [datetime]$Start = (Get-date).AddDays(-1)
    )

    Process {

        $VM | ForEach-Object {

            $Cpu  = (Get-Stat -Entity $_ -Stat cpu.usagemhz.average -Start $Start).value | Measure-Object -Average -Maximum

            $Memory  = (Get-Stat -Entity $_ -Stat mem.usage.average -Start $Start).value | Measure-Object -Average -Maximum

            [PSCustomObject]@{

                VM = $_.Name
                CpuAverageMhz = [math]::Round($Cpu.Average,0)
                CpuMaximumMHz = [math]::Round($Cpu.Maximum,0)
                MemAverageGB = [math]::Round(($Memory.Average / 100 * $_.MemoryGB),2)
                MemMaximumGB = [math]::Round(($Memory.Maximum / 100 * $_.MemoryGB),2)
            }

            Clear-Variable Cpu,Memory
        }

    }
}

Function Set-VMCopyPaste {

    <#
    .DESCRIPTION
    Set $Disabled to false to enable copy/paste on the VM.
    Set $Disabled to true to disable copy/paste on the VM.

    The VM will require a restart for the changes to be effective.

    The outputs are the 2 advanced settings.
    #>

    param(

        [parameter(position=0,ValueFromPipeline=$True,ValueFromPipelineByPropertyname=$True,Mandatory=$True)]
        [VMware.VimAutomation.ViCore.Impl.V1.VM.UniversalVirtualMachineImpl[]]
        $VM,    

        [parameter(ParameterSetName=0)]
        [validateset("true","false")]
        [string]
        $Disabled

    )

    Process{

    $Copy = "isolation.tools.copy.disable"
    $Paste= "isolation.tools.paste.disable"

        $VM | ForEach-Object {
            $CurrentVM = $_
            $Copy,$Paste | ForEach-Object {

                $CP = $_    
                $setting = $CurrentVM | Get-AdvancedSetting -Name $CP

                IF   ($setting) {$setting | Set-AdvancedSetting -Value $Disabled -confirm:$false}
                ELSE {$CurrentVM | New-AdvancedSetting -Name $CP -Value $Disabled -confirm:$false}

            }
        }

    }

} 

## Export all functions and all aliases
Export-ModuleMember -Function * -Alias *