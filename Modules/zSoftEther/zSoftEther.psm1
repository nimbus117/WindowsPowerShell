### Module - zSoftEther

function Connect-SoftEtherVPN {

    param(
        
        [string]$ConnectionName = 'oni',
        [string]$VpncmdPath = 'C:\Program Files\SoftEther VPN Client\vpncmd.exe'
    )

    if ($ConnectionName -eq 'oni') {
        if (-not (Test-Connection flat.nimbus117.co.uk -Quiet -Count 2)) {
            throw "Unable to ping $ConnectionName"
        }
    }

    Invoke-Expression "& `"$VpncmdPath`" localhost /client /CMD AccountConnect $ConnectionName"
}
Set-Alias -Name cvpn -Value Connect-SoftEtherVPN


function Disconnect-SoftEtherVPN {

    param(
        
        [string]$ConnectionName = 'oni',
        [string]$VpncmdPath = 'C:\Program Files\SoftEther VPN Client\vpncmd.exe'
    )

    Invoke-Expression "& `"$VpncmdPath`" localhost /client /CMD AccountDisconnect $ConnectionName"
}
Set-Alias -Name dvpn -Value Disconnect-SoftEtherVPN


function Get-SoftEtherVPNConnection {

    param(
        
        [string]$VpncmdPath = 'C:\Program Files\SoftEther VPN Client\vpncmd.exe'
    )

    Invoke-Expression "& `"$VpncmdPath`" localhost /client /CMD Accountlist"
}
Set-Alias -Name gvpn -Value Get-SoftEtherVPNConnection

## Export all functions and all aliases
Export-ModuleMember -Function * -Alias *