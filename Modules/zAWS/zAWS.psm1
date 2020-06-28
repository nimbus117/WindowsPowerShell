### Module - zAWS

function Import-AWSModule {Import-Module "C:\Program Files (x86)\AWS Tools\PowerShell\AWSPowerShell\AWSPowerShell.psd1"}

# Download aws keys, import csv, set aws credentials, set default region
# $AWSCreds = Import-Csv 'C:\Users\James\Dropbox\Computer\AWS\zonal-credentials.csv'
# Set-AWSCredentials -AccessKey $AWSCreds.'Access Key Id' -SecretKey $AWSCreds.'Secret Access Key'
# Initialize-AWSDefaults -Region eu-west-1

<#
$KeyName = 'jg-aws'
$PemPath = "$env:USERPROFILE\$KeyName.pem"
$KeyPair = New-EC2KeyPair -KeyName $KeyName
$KeyPair.KeyMaterial | Out-File -Encoding ascii -FilePath $PemPath
"KeyName: {0}" -f $KeyPair.KeyName | Out-File -Encoding ascii -FilePath $PemPath -Append
"KeyFingerprint: {0}" -f $KeyPair.KeyFingerprint | Out-File -Encoding ascii -FilePath $PemPath -Append

$SG = 'win-sg'
New-EC2SecurityGroup $SG -Description "Windows Remote Access"
Grant-EC2SecurityGroupIngress -GroupName $SG -IpPermissions @{IpProtocol = "icmp"; FromPort = -1; ToPort = -1; IpRanges = @("94.175.82.61/32")}
Grant-EC2SecurityGroupIngress -GroupName $SG -IpPermissions @{IpProtocol = "tcp"; FromPort = 3389; ToPort = 3389; IpRanges = @("94.175.82.61/32")}
Grant-EC2SecurityGroupIngress -GroupName $SG -IpPermissions @{IpProtocol = "udp"; FromPort = 3389; ToPort = 3389; IpRanges = @("94.175.82.61/32")}
Grant-EC2SecurityGroupIngress -GroupName $SG -IpPermissions @{IpProtocol = "tcp"; FromPort = 5985; ToPort = 5986; IpRanges = @("94.175.82.61/32")}
#>

function DeployNew-EC2Instance {
    
    [CmdletBinding()]
    param(
        [parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][string]$InstanceName,
        [string]$Region = 'eu-west-1',
        [string]$KeyName = 'jg-aws',
        [string]$InstanceType = 'm1.small',
        [string]$SGName = 'win-sg',
        [string]$PemPath = "$env:USERPROFILE\Dropbox\Computer\AWS\jg-aws.pem"
    )

    $ErrorActionPreference = 'Stop'

    $MS2012Id = Get-EC2ImageByName WINDOWS_2012R2_BASE | Select-Object -First 1
    Write-Verbose ("Using Image: '{0}'" -f $MS2012Id.Name)

    $UserData = @"
<powershell>
    Set-NetFirewallRule -Name WINRM-HTTP-In-TCP-PUBLIC -RemoteAddress Any
    "test" | out-file C:\Users\Administrator\test.txt
</powershell>
"@

    $UserDataBase64 = [System.Convert]::ToBase64String([System.Text.ASCIIEncoding]::UTF8.GetBytes($UserData))

    Write-Verbose ("New Instance: ImageId '{0}', Type '{1}', Key '{2}', SG '{3}'" -f $MS2012Id.ImageId,$InstanceType,$KeyName,$SGName)
    $Instance = New-EC2Instance -ImageId $MS2012Id.ImageId -InstanceType $InstanceType -KeyName $KeyName -SecurityGroup $SGName -UserData $UserDataBase64 -Region $Region

    $InstanceId = ($Instance | Select-Object -ExpandProperty Instances).InstanceId
    Write-Verbose "InstanceId: '$InstanceId'"

    $InstanceState = (Get-EC2Instance -Instance $InstanceId).Instances.State.Name
    while ($InstanceState -ne 'running') {
        
        Write-Verbose "Wating for running state - $InstanceState"
        Start-Sleep 10
        $InstanceState = (Get-EC2Instance -Instance $InstanceId).Instances.State.Name
    }

    Write-Verbose "Naming Instance: '$InstanceName'"
    $InstanceTag = New-EC2Tag -Resource $InstanceId -Tag @{Key = 'Name' ; Value = $InstanceName}

    $Public = (Get-EC2Instance -Instance $InstanceId).Instances
    Write-Verbose ("Public Address: '{0}', {1}" -f $Public.PublicDnsName,$Public.PublicIpAddress)

    while (-not(tPing $Public.PublicIpAddress -Quiet)) {
    
        Write-Verbose "Wating for ping response"
        Start-Sleep 10
    }

    if (Test-Path $PemPath) {

        $Password = $null
        while ($Password -eq $null) {
        
            try {$Password = Get-EC2PasswordData -InstanceId $InstanceId -PemFile $PemPath -Decrypt}
            catch {
                Write-Verbose "Waiting for password"
                Start-Sleep -Seconds 10
            }
        }
        Write-Verbose "Admin Password: $Password"
        $SecurePass = ConvertTo-SecureString $Password -AsPlainText -Force
        $Credential = New-Object System.Management.Automation.PSCredential ("Administrator", $SecurePass)

        New-Variable -Scope 'Global' -Name "Creds_$InstanceName" -Value $Credential
    }
}