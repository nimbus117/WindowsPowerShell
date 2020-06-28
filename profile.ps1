### PowerShell All Hosts Profile

# Set location to C:\Users\user
Set-Location $env:USERPROFILE

# Set windows and buffer size for ConsoleHost
if ($Host.Name -eq "ConsoleHost") {
    try {
        $pf_BufferSize = New-Object System.Management.Automation.Host.Size(128,3000)
        $Host.ui.rawui.BufferSize = $pf_BufferSize
        $pf_WindowSize = New-Object System.Management.Automation.Host.Size(128,50)
        $Host.ui.rawui.WindowSize = $pf_WindowSize
    } 
    
    catch {
        try {
            $pf_WindowSize = New-Object System.Management.Automation.Host.Size(128,50)
            $Host.ui.rawui.WindowSize = $pf_WindowSize
            $pf_BufferSize = New-Object System.Management.Automation.Host.Size(128,3000)
            $Host.ui.rawui.BufferSize = $pf_BufferSize
        } catch {}
    }
}

# Set back and foreground colours
$Host.UI.RawUI.BackgroundColor = 'Black'
$Host.UI.RawUI.ForegroundColor = 'DarkCyan' 
Clear-Host

# Get current user
$pf_CurrentUser = New-Object System.Security.Principal.WindowsPrincipal([System.Security.Principal.WindowsIdentity]::GetCurrent())

# Check if running as Administrator
$pf_IsAdmin = $pf_CurrentUser.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
if($pf_IsAdmin) {$pf_User = $pf_CurrentUser.Identities.Name + " (Administrator)"} 
else {$pf_User = $pf_CurrentUser.Identities.Name}

# Set host title - Domain\User - ComputerName (Version)
$pf_HostTitle = "{0} - {1} (v{2}.{3})" -f $pf_User,$env:COMPUTERNAME,$Host.Version.Major,$Host.Version.Minor
$Host.UI.RawUI.WindowTitle = $pf_HostTitle

# Custom prompt
function Prompt {

    $pf_BatRemaining = (Get-WmiObject -Class Win32_Battery).EstimatedChargeRemaining
    if ($pf_BatRemaining) {

        $pf_DashOffset = ($pf_BatRemaining | Measure-Object -Character).Characters + 2
        if ($pf_BatRemaining -ge 50) {$pf_BatColour = 'Green'} 
        elseif ($pf_BatRemaining -ge 25) {$pf_BatColour = 'Yellow'} 
        else {$pf_BatColour = 'Red'}
        Write-Host "[" -NoNewline -ForegroundColor $Host.UI.RawUI.ForegroundColor
        Write-Host $pf_BatRemaining -NoNewline -ForegroundColor $pf_BatColour
        Write-Host "]" -NoNewline -ForegroundColor $Host.UI.RawUI.ForegroundColor
    }

    $pf_DashOffset = $pf_DashOffset + 6
    $pf_Pwd = Get-Location
    $pf_Time = Get-Date -Format HH:mm:ss
    Write-Host "[" -NoNewline -ForegroundColor $Host.UI.RawUI.ForegroundColor
    Write-Host $pf_Time -NoNewline -ForegroundColor DarkGray
    Write-Host "][" -NoNewline -ForegroundColor $Host.UI.RawUI.ForegroundColor
    Write-Host $pf_Pwd -NoNewline -ForegroundColor DarkRed
    Write-Host "]" -NoNewline -ForegroundColor $Host.UI.RawUI.ForegroundColor
    
    if ($global:DefaultVIServer) {

        $pf_DashOffset = $pf_DashOffset + $global:DefaultVIServer.Name.Length + 2
        Write-Host "[" -NoNewline -ForegroundColor $Host.UI.RawUI.ForegroundColor
        Write-Host $global:DefaultVIServer.Name -NoNewline -ForegroundColor Yellow
        Write-Host "]" -NoNewline -ForegroundColor $Host.UI.RawUI.ForegroundColor
    }
    
    $pf_Width = ($Host.UI.RawUI.BufferSize.Width - $pf_DashOffset - $pf_Pwd.ToString().Length - $pf_Time.ToString().Length)
    if ($pf_IsAdmin) {$pf_DashColour = 'DarkRed'} else {$pf_DashColour = $Host.UI.RawUI.ForegroundColor}
    Write-Host " $('-' * $pf_Width)" -ForegroundColor $pf_DashColour #(Get-Random @(1..8))
    'PS> '
}

# Path Variables
$HostsFile = "$env:SystemRoot\System32\drivers\etc\hosts"
$PSScripts = "$env:USERPROFILE\Dropbox\Computer\code\powershell"
$Tools = "$env:USERPROFILE\Dropbox\Computer\tools"

# Session Transcript
1..10 | ForEach-Object {try{Stop-Transcript | Out-Null} catch{}}
$TranscriptPartialPath = "$env:TEMP\PSTranscript-"
Get-ChildItem "$TranscriptPartialPath*" | Sort-Object LastWriteTime -Descending | 
Select-Object -Skip 2 | Remove-Item -Force -Confirm:$false -ErrorAction SilentlyContinue
Start-Transcript -Path (“$TranscriptPartialPath$(Get-Date -Format yyMMddHHmm).txt”) -Force -Append | Out-Null

# Service Fabric command aliases
Get-Command -Module ServiceFabric -ErrorAction SilentlyContinue | where name -like '*ServiceFabric*' | ForEach-Object {
    
    $Alias = $_.Name -replace 'ServiceFabric','SF'

    Set-Alias -Name $Alias -Value $_.Name

    Clear-Variable Alias
}

# NetFirewall command aliases
Get-Command -Module NetSecurity -ErrorAction SilentlyContinue | where name -like '*NetFirewall*' | ForEach-Object {
    
    $Alias = $_.Name -replace 'NetFirewall','NF'

    Set-Alias -Name $Alias -Value $_.Name

    Clear-Variable Alias
}

# Set Default Parameter Values
if ($PSVersionTable.PSVersion.Major -gt 2) {

    $PSDefaultParameterValues['Format-[wt]*:Autosize'] = $true
    #$PSDefaultParameterValues['*-vm*:ComputerName'] = 'ts-5500'
}

# Functions

function Connect-vCenter {

    param(
        [parameter(position=0)]
        [ValidateNotNullOrEmpty()]
        $Server = 'dca-vcenter',

        [Parameter(Position=1)]
        [PSCredential][System.Management.Automation.CredentialAttribute()]
        $Credential,

        [parameter()]
        [ValidateNotNullOrEmpty()]
        $Protocol = 'https'
    )

    $ErrorActionPreference = 'Stop'

    Import-Module VMware.VimAutomation.Core

    if (($Credential -eq $null) -and ($zcred)) {$Credential = $zcred}

    $params = @{Server = $Server ; Protocol = $Protocol}
    if ($Credential) {$params.Add('Credential',$Credential)}

    Connect-VIServer @params 3>&1 | Out-Null

    if ($Global:DefaultVIServer) {

        if (!$Global:pf_HostTitle) {$Global:pf_HostTitle = $Host.UI.RawUI.WindowTitle}

        $Host.UI.RawUI.WindowTitle = $Global:pf_HostTitle + " - $($Global:DefaultVIServer.Name)"
    
        $MOTD = Get-AdvancedSetting -Name "vpxd.motd" -Entity $Global:DefaultVIServer -ErrorAction SilentlyContinue

        if ($MOTD.Value) {Write-Host ("`n{0}`n" -f $MOTD.Value) -ForegroundColor Yellow}
    }
}
Set-Alias cvc -Value Connect-vCenter