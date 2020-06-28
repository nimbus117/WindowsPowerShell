### Module - zHyperV

function New-VMFromIso {
        
    [cmdletbinding()]
    param(
        [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][string]$Name,
	    [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][string]$ISO,
        [Parameter()][string]$Path = 'D:',
        [Parameter()][int64]$StartBytes = 1024MB,
        [Parameter()][int64]$MinBytes = 512MB,
        [Parameter()][int64]$MaxBytes = 2048MB,
        [Parameter()][int64]$VHDSizeBytes = 20GB,
        [Parameter()][string]$VSwitchName = 'vExternal'
    )

    $ErrorActionPreference = 'Stop'
    $NewVHDFilePath = "$Path\$Name\Virtual Hard Disks"
    $NewVHDFile =  "$NewVHDFilePath\$Name.vhdx"

    try {
        Write-Verbose "Creating new VM: '$Name'"
        $NewVM = New-VM -Name $Name -MemoryStartupBytes $StartBytes -Path $Path -Generation 1

        Write-Verbose "Creating Directory: '$NewVHDFilePath'"
        $NewDir = New-Item -Path $NewVHDFilePath -ItemType directory -Force

        Write-Verbose "Creating new vhd: '$NewVHDFile $VHDSizeBytes'"
        $NewVHD = New-VHD -Path $NewVHDFile -SizeBytes $VHDSizeBytes -Dynamic
        
        Write-Verbose "Adding vhd: '$NewVHDFile'"
        $AddVHD = Add-VMHardDiskDrive -VMName $Name -Path $NewVHDFile

        Write-Verbose "Setting memory: 'Start=$StartBytes Min=$MinBytes Max=$MaxBytes'"
        $SetMemory = Set-VMMemory -VMName $Name -DynamicMemoryEnabled $true -StartupBytes $StartBytes -MinimumBytes $MinBytes -MaximumBytes $MaxBytes

        Get-VMNetworkAdapter $Name | Remove-VMNetworkAdapter
        Write-Verbose "Adding network adapter: '$VSwitchName'"
        $AddNetwork = Add-VMNetworkAdapter -VMName $Name -SwitchName $VSwitchName

        Write-Verbose "Setting dvd drive: '$ISO'"
        $SetDvd = Set-VMDvdDrive -VMName $Name -ControllerNumber 1 -Path $ISO
    }
    catch {Write-Error $_}
}

function New-VMFromTemplate {
    
    [cmdletbinding()]
    param(
        [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][string]$Name,
	    [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][string]$TemplateVHD,
        [Parameter()][string]$Path = 'D:',
        [Parameter()][int64]$StartBytes = 1024MB,
        [Parameter()][int64]$MinBytes = 512MB,
        [Parameter()][int64]$MaxBytes = 2048MB,
        [Parameter()][string]$VSwitchName = 'vExternal'
    )

    $ErrorActionPreference = 'Stop'

    try {
        $NewVHDFilePath = "$Path\$Name\Virtual Hard Disks"
        $NewVHDFile =  "$NewVHDFilePath\$Name.vhdx"

        Write-Verbose "Creating Directory: '$NewVHDFilePath'"
        $NewDir = New-Item -Path $NewVHDFilePath -ItemType directory -Force

        Write-Verbose "Converting template vhd: '$TemplateVHD'"
        $ConvertVHD = Convert-VHD -Path $TemplateVHD -DestinationPath $NewVHDFile -VHDType Dynamic

        Write-Verbose "Creating new VM: '$Name'"
        $NewVM = New-VM -Name $Name -MemoryStartupBytes $StartBytes -Path $Path -Generation 1

        Write-Verbose "Adding vhd: '$NewVHDFile'"
        $AddVHD = Add-VMHardDiskDrive -VMName $Name -Path $NewVHDFile

        Write-Verbose "Setting memory: 'Start=$StartBytes,Min=$MinBytes,Max=$MaxBytes'"
        $SetMemory = Set-VMMemory -VMName $Name -DynamicMemoryEnabled $true -StartupBytes $StartBytes -MinimumBytes $MinBytes -MaximumBytes $MaxBytes

        Get-VMNetworkAdapter $Name | Remove-VMNetworkAdapter
        Write-Verbose "Adding network adapter: '$VSwitchName'"
        $AddNetwork = Add-VMNetworkAdapter -VMName $Name -SwitchName $VSwitchName
    }
    catch {Write-Error $_}
}

function New-VMFromTemplateDomain {
    
    [cmdletbinding()]
    param(
        [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][string[]]$VMName,
	    [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][string]$TemplateVHD,
        [Parameter()][string]$Path = 'D:\',
        [Parameter()][int64]$ProcessorCount = 1,
        [Parameter()][int64]$StartBytes = 1024MB,
        [Parameter()][int64]$MinBytes = 512MB,
        [Parameter()][int64]$MaxBytes = 2048MB,
        [Parameter()][string]$VSwitchName = 'vExternal',
        [Parameter()][string]$Key = '----',
        [Parameter()][string]$Domain = 'base.nimbus117.co.uk',
        [Parameter()][PSCredential][System.Management.Automation.CredentialAttribute()]$SetLocalAdminCred = (Get-Credential local\administrator -Message 'Local Admin user credential'),
        [Parameter()][PSCredential][System.Management.Automation.CredentialAttribute()]$DomainJoinCred = (Get-Credential -Message 'Domain join credential')
    )

    $ErrorActionPreference = 'Stop'

    foreach ($Name in $VMName) {

$AnswerFile = @"
<!--*************************************************
Windows Server 2016 Answer File Generator
Created using Windows AFG found at:
http://windowsafg.no-ip.org

Installation Notes
Location: 
Notes: 
**************************************************-->

<?xml version="1.0" encoding="utf-8"?>
<unattend xmlns="urn:schemas-microsoft-com:unattend">
<settings pass="windowsPE">
<component name="Microsoft-Windows-International-Core-WinPE" processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS" xmlns:wcm="http://schemas.microsoft.com/WMIConfig/2002/State" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<SetupUILanguage>
<UILanguage>en-US</UILanguage>
</SetupUILanguage>
<InputLocale>0c09:00000409</InputLocale>
<SystemLocale>en-US</SystemLocale>
<UILanguage>en-US</UILanguage>
<UILanguageFallback>en-US</UILanguageFallback>
<UserLocale>en-US</UserLocale>
</component>
<component name="Microsoft-Windows-Setup" processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS" xmlns:wcm="http://schemas.microsoft.com/WMIConfig/2002/State" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<ImageInstall>
<OSImage>
<InstallTo>
<DiskID>0</DiskID>
<PartitionID>2</PartitionID>
</InstallTo>
</OSImage>
</ImageInstall>
<UserData>
<AcceptEula>true</AcceptEula>
<FullName>administrator</FullName>
<Organization>t55</Organization>
<ProductKey>
<Key>$Key</Key>
</ProductKey>
</UserData>
<EnableFirewall>true</EnableFirewall>
<DiskConfiguration>
<Disk wcm:action="add">
<CreatePartitions>
<CreatePartition wcm:action="add">
<Order>1</Order>
<Size>350</Size>
<Type>Primary</Type>
</CreatePartition>
<CreatePartition wcm:action="add">
<Extend>true</Extend>
<Order>2</Order>
<Type>Primary</Type>
</CreatePartition>
</CreatePartitions>
<ModifyPartitions>
<ModifyPartition wcm:action="add">
<Format>NTFS</Format>
<Label>System</Label>
<Order>1</Order>
<PartitionID>1</PartitionID>
<TypeID>0x27</TypeID>
</ModifyPartition>
<ModifyPartition wcm:action="add">
<Order>2</Order>
<PartitionID>2</PartitionID>
<Letter>C</Letter>
<Label>OS</Label>
<Format>NTFS</Format>
</ModifyPartition>
</ModifyPartitions>
<DiskID>0</DiskID>
<WillWipeDisk>true</WillWipeDisk>
</Disk>
</DiskConfiguration>
</component>
</settings>
<settings pass="offlineServicing">
<component name="Microsoft-Windows-LUA-Settings" processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS" xmlns:wcm="http://schemas.microsoft.com/WMIConfig/2002/State" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<EnableLUA>true</EnableLUA>
</component>
</settings>
<settings pass="generalize">
<component name="Microsoft-Windows-Security-SPP" processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS" xmlns:wcm="http://schemas.microsoft.com/WMIConfig/2002/State" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<SkipRearm>1</SkipRearm>
</component>
</settings>
<settings pass="specialize">
<component name="Microsoft-Windows-International-Core" processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS" xmlns:wcm="http://schemas.microsoft.com/WMIConfig/2002/State" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<InputLocale>0809:00000809</InputLocale>
<SystemLocale>en-GB</SystemLocale>
<UILanguage>en-GB</UILanguage>
<UILanguageFallback>en-GB</UILanguageFallback>
<UserLocale>en-GB</UserLocale>
</component>
<component name="Microsoft-Windows-Security-SPP-UX" processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS" xmlns:wcm="http://schemas.microsoft.com/WMIConfig/2002/State" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<SkipAutoActivation>true</SkipAutoActivation>
</component>
<component name="Microsoft-Windows-SQMApi" processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS" xmlns:wcm="http://schemas.microsoft.com/WMIConfig/2002/State" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<CEIPEnabled>0</CEIPEnabled>
</component>
<component name="Microsoft-Windows-Shell-Setup" processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS" xmlns:wcm="http://schemas.microsoft.com/WMIConfig/2002/State" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<ComputerName>$Name</ComputerName>
</component>
</settings>
<settings pass="oobeSystem">
<component name="Microsoft-Windows-Shell-Setup" processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS" xmlns:wcm="http://schemas.microsoft.com/WMIConfig/2002/State" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<OOBE>
<HideEULAPage>true</HideEULAPage>
<HideLocalAccountScreen>true</HideLocalAccountScreen>
<HideOEMRegistrationScreen>true</HideOEMRegistrationScreen>
<HideOnlineAccountScreens>true</HideOnlineAccountScreens>
<HideWirelessSetupInOOBE>true</HideWirelessSetupInOOBE>
<NetworkLocation>Work</NetworkLocation>
<ProtectYourPC>1</ProtectYourPC>
<SkipMachineOOBE>true</SkipMachineOOBE>
<SkipUserOOBE>true</SkipUserOOBE>
</OOBE>
<UserAccounts>
<AdministratorPassword>
<Value>$($SetLocalAdminCred.GetNetworkCredential().Password)</Value>
<PlainText>true</PlainText>
</AdministratorPassword>
<LocalAccounts>
<LocalAccount wcm:action="add">
<Description>administrator</Description>
<DisplayName>administrator</DisplayName>
<Group>Administrators</Group>
<Name>administrator</Name>
</LocalAccount>
</LocalAccounts>
</UserAccounts>
<RegisteredOrganization>t55</RegisteredOrganization>
<RegisteredOwner>administrator</RegisteredOwner>
<DisableAutoDaylightTimeSet>false</DisableAutoDaylightTimeSet>
<TimeZone>GMT Standard Time</TimeZone>
</component>
</settings>
</unattend>
"@

        try {
            $NewVHDFilePath = "$Path\$Name\Virtual Hard Disks"
            $NewVHDFile =  "$NewVHDFilePath\$Name.vhdx"

            Write-Verbose "Creating Directory: '$NewVHDFilePath'"
            $NewDir = New-Item -Path $NewVHDFilePath -ItemType directory -Force

            Write-Verbose "Converting template vhd: '$TemplateVHD'"
            $ConvertVHD = Convert-VHD -Path $TemplateVHD -DestinationPath $NewVHDFile -VHDType Dynamic

            Write-Verbose "Mounting Disk Image and copying setup files"
            Mount-DiskImage $NewVHDFile
            $DriveLetter=((Get-DiskImage $NewVHDFile | Get-Disk | Get-Partition | where DriveLetter).DriveLetter) +":"
            $null = Get-PSDrive
            $AnswerFile | Out-File -FilePath "$Driveletter\Windows\System32\Sysprep\unattend.xml"
            Dismount-DiskImage $NewVHDFile

            Write-Verbose "Creating new VM: '$Name'"
            $NewVM = New-VM -Name $Name -MemoryStartupBytes $StartBytes -Path $Path -Generation 1

            Write-Verbose "Adding vhd: '$NewVHDFile'"
            $AddVHD = Add-VMHardDiskDrive -VMName $Name -Path $NewVHDFile

            Write-Verbose "Setting memory: 'Start=$StartBytes,Min=$MinBytes,Max=$MaxBytes'"
            $SetMemory = Set-VMMemory -VMName $Name -DynamicMemoryEnabled $true -StartupBytes $StartBytes -MinimumBytes $MinBytes -MaximumBytes $MaxBytes

            Write-Verbose "Setting processor count to: $ProcessorCount"
            $SetProcessorCount = Set-VMProcessor -VMName $Name -Count $ProcessorCount

            Get-VMNetworkAdapter $Name | Remove-VMNetworkAdapter
            Write-Verbose "Adding network adapter: '$VSwitchName'"
            $AddNetwork = Add-VMNetworkAdapter -VMName $Name -SwitchName $VSwitchName
    
            Write-Verbose "Starting VM and waiting for conectivity"
            Start-VM -VMName $Name -Passthru | Wait-VM -For IPAddress -Timeout 300

            Write-Verbose "Joining to domain: $Domain"
            Invoke-Command -VMName $Name -Credential $SetLocalAdminCred -ScriptBlock {
            
                Remove-Item "$env:SystemRoot\System32\Sysprep\unattend.xml"

                Add-Computer -DomainName $using:Domain -Credential $using:DomainJoinCred -Restart
            }
        }
        catch {Write-Error $_ -ErrorAction Continue}
    }
}

function Get-VHDSize {

    param(
        [Parameter(ValueFromPipeline=$true)][Microsoft.HyperV.PowerShell.VirtualMachine]$VM,
        [Parameter(Position=0)][validateset("KB","MB","GB","TB")][String]$Units = 'GB'
    )

    process{

        foreach ($x in $VM) {
    
            $x | Select-Object -ExpandProperty HardDrives | Get-VHD | Select-Object `
                @{l='Name';e={$x.Name}},`
                Path,`
                @{l='Size';e={[math]::round($_.Size / "1$Units", 2)}},`
                @{l='FileSize';e={[math]::round($_.FileSize / "1$Units", 2)}}
        }
    }
}

function Get-VMMemoryInfo {

    param(
        [Parameter(ValueFromPipeline=$true)][Microsoft.HyperV.PowerShell.VirtualMachine]$VM,
        [Parameter(Position=0)][validateset("KB","MB","GB")][String]$Units = 'MB'
    )

    process{

        foreach ($x in $VM) {
    
            $x | Select-Object `
                Name,`
                @{l='Dynamic';e={$_.DynamicMemoryEnabled}},
                @{l='Min';e={[math]::round($_.MemoryMinimum / "1$Units", 2)}},`
                @{l='Start';e={[math]::round($_.MemoryStartup / "1$Units", 2)}},`
                @{l='Max';e={[math]::round($_.MemoryMaximum / "1$Units", 2)}},`
                @{l='Assigned';e={[math]::round($_.MemoryAssigned / "1$Units", 2)}},`
                @{l='Demand';e={[math]::round($_.MemoryDemand / "1$Units", 2)}},`
                @{l='Status';e={$_.MemoryStatus}}
        }
    }
}

function Start-VMConnect {
    
    [CmdletBinding()]

    param(
        [Parameter(Position=0,Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [string[]]$VMName,

        [Parameter(Position=1,ValueFromPipelineByPropertyName=$true)]
        [string]$ComputerName = 'localhost',

        [ValidateNotNullOrEmpty()]
        [string]$VMConnectPath = 'C:\Windows\System32\vmconnect.exe'
    )

    begin {

        Write-Verbose "Start: $($MyInvocation.MyCommand)"
        if (-not (Test-Path $VMConnectPath)) {Write-Warning "Requires $VMConnectPath" ; Break}
        $Count = 0
    }

    process{

        foreach ($VM in $VMName) {

            try {

                Write-Verbose "Getting VM '$VM'"
                $GetVM = Get-VM $VM -ComputerName $ComputerName -ErrorAction Stop
                $ArgumetList = @($ComputerName, $GetVM.Name, "-G $($GetVM.Id)", "-C $Count")
                
                Write-Verbose "Starting Process '$VMConnectPath $ArgumetList'"
                Start-Process -FilePath $VMConnectPath -ArgumentList $ArgumetList -ErrorAction Stop
                $Count++
                Start-Sleep -Milliseconds 500
            } 

            catch {Write-Error $_}
        }
    }

    end {Write-Verbose "End: $($MyInvocation.MyCommand)"}
}

### 5nine

function Start-VMConsole {
    
    [CmdletBinding()]
    param(
        [Parameter(Position=0,Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [string[]]$Name,
        [Parameter()][string]$HostID = "542795b9-ff1d-4265-ab6c-f950f1a8a2a9",
        [Parameter()][string]$HostName = $env:COMPUTERNAME
    )

    begin {
        Write-Verbose "Start: $($MyInvocation.MyCommand)"
        $5nineGuestConsole = "C:\Program Files\5nine\5nine Manager for Hyper-V\5nine.GuestConsole.exe"
        if (-not (Test-Path $5nineGuestConsole)) {Write-Warning "Requires $5nineGuestConsole" ; Break}
        $Count = 0
    }

    process{
        foreach ($VM in $Name) {
            try {
                Write-Verbose "Getting VM '$VM'"
                $GetVM = Get-VM $VM -ErrorAction Stop
                $ArgumetList = "/cert-ignore /vmconnect:$($GetVM.Id) /v:$HostName`:2179 /hypervplugin:5nine.GuestConsole.Plugin.dll /hostid:$HostID"
                Write-Verbose "Starting Process '$5nineGuestConsole $ArgumetList'"
                Start-Process -FilePath $5nineGuestConsole -ArgumentList $ArgumetList -ErrorAction Stop
                $Count++
                Start-Sleep -Milliseconds 500
            } 
            catch {Write-Error $_}
        }
    }

    end {Write-Verbose "End: $($MyInvocation.MyCommand)"}
}

## Export all functions and all aliases
Export-ModuleMember -Function * -Alias * 