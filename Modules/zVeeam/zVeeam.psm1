### Module - zMisc

function Get-AvgBackupSizes {

    param([parameter(Mandatory=$true)][ValidateScript({Test-Path $_})][string]$Path)

    Get-ChildItem $Path | ForEach-Object {

        $Files = Get-ChildItem $_.FullName -Recurse

        $Full = $Files | where Extension -eq '.vbk'
        $FullCount = ($Full | Measure-Object).Count
        $FullAvgSize = ($Full | Measure-Object -Average -Property Length).Average

        $Inc = $Files | Where-Object {($_.Extension -eq '.vrb') -or ($_.Extension -eq '.vib')}
        $IncCount = ($Inc | Measure-Object).Count
        $IncAvgSize = ($Inc | Measure-Object -Average -Property Length).Average

        [pscustomobject]@{

            Job = $_.Name
            FullCount = $FullCount
            FullAvgSize = [math]::round($FullAvgSize / 1GB, 1)
            IncCount = $IncCount
            IncAvgSize = [math]::Round($IncAvgSize / 1GB, 1)
        }
    }
}


function Get-VMChurn {
    
    param(
        [parameter()]
        [int]$DaysBack = 7,
        [parameter()]
        [string]$JobFilter = "*-daily*"
    )

    Add-PSSnapin -name "VeeamPSSnapIn"

    (Get-VBRBackupSession |

    Where-Object {

        ($_.EndTime -ge (Get-Date).AddDays(-$DaysBack)) -and 
        ($_.isfullmode -eq $false) -and 
        ($_.JobName -like $JobFilter)
    }

    ).GetTaskSessions() | 
    
    Select-Object Name, `
        @{label='StartTime';expression={$_.info.QueuedTime}}, `
        @{label='ProcessedGB';expression={[math]::Round($_.Progress.ProcessedSize / 1GB,2)}}, `
        @{label='ReadGB';expression={[math]::Round($_.Progress.ReadSize / 1GB,2)}}, `
        @{label='TransferedGB';expression={[math]::Round($_.Progress.TransferedSize / 1GB,2)}} |

    Sort-Object @{e='StartTime';descending=$true}, Name
}