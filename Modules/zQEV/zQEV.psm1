### Module - zQEV

# Function to list logs
function Get-QEVEventLogList {

        param(
		    [string[]]$ComputerName,
            [string[]]$Filter = '*'
        )

        $sb = {

            Get-WinEvent -ListLog $args | Where-Object {$_.RecordCount} |

            Select-Object LogName,RecordCount,LastWriteTime,@{label='ActLastWriteTime';expression={(Get-WinEvent -LogName $_.LogName -MaxEvents 1).TimeCreated}}
        }

        $Parameters = @{ScriptBlock = $sb ; ArgumentList = $Filter}

        if ($ComputerName) {$Parameters.Add('ComputerName',$ComputerName)}

        if ($Credential) {$Parameters.Add('Credential',$Credential)}

        Invoke-Command @Parameters | 
        
        Sort-Object @{expression={if ($_.LogName -like "*Microsoft-*") {"zz$($_.LogName)"} else {"$($_.LogName)"}}} -Unique

    }#Get-QEVEventLogList

# Function to return events
function Get-QEVEvents {

	param(
		[string[]]$ComputerName,
		[datetime]$StartTime,
		[datetime]$EndTime,
		[int[]]$Level,
        [int[]]$IDs,
        [string]$Message,
        [string]$UserName,
		[string[]]$Logs,
        [int]$MaxEvents
	)

	$sb = {

		$Logs = $args[0]
		$StartTime = $args[1]
		$EndTime = $args[2]
		$Level = $args[3]
        $IDs = $args[4]
        $UserName = $args[5]
        $Message = $args[6]
        $MaxEvents = $args[7]
    
        $CurrentCulture = Get-Culture
        [System.Threading.Thread]::CurrentThread.CurrentCulture = New-Object "System.Globalization.CultureInfo" "en-US"

        if ($level -contains 4) {$level += 0 ; $level += 5}

        $EventFilter = @{
                
            Logname=@($Logs)
            StartTime=$StartTime
            EndTime=$EndTime
            Level=@($level)
        }

        if ($IDs) {$EventFilter.Add('ID',@($IDs))}

        if ($UserName) {

            try {

                $UserObject = New-Object System.Security.Principal.NTAccount("$UserName") 
                $UserSID = $UserObject.Translate([System.Security.Principal.SecurityIdentifier]) 
                if ($UserSID) {$EventFilter.Add('UserID',$UserSID.Value)}
            } 
            catch {Write-Warning "Unable to translate NTAccount to SID - $UserName."}
        }
                
        $Events = Get-WinEvent -FilterHashtable $EventFilter -MaxEvents $MaxEvents | 
            
        Select-Object *, @{Label='NTAccount';Expression={
            
            $UserId = $_.UserId
            try {
                if ($UserId) {
                    $SID = New-Object System.Security.Principal.SecurityIdentifier ($UserId)
                    $User = $SID.Translate([System.Security.Principal.NTAccount])
                    $User.Value
                }
                else {"N/A"}
            }
            catch {$UserId ; Write-Warning "Unable to translate SID to NTAccount - $UserId"}
            finally {if ($UserId) {Clear-Variable UserId}}
        }}

        if ($Message) {$Events | Where-Object {$_.Message -match $Message}}
        else {$Events}

        [System.Threading.Thread]::CurrentThread.CurrentCulture = $CurrentCulture
	}

    $Parameters = @{
            
        ScriptBlock = $sb
            
        ArgumentList = @($Logs, $StartTime, $EndTime, $Level, $IDs, $UserName, $Message, $MaxEvents)
    }

    if ($ComputerName) {$Parameters.Add('ComputerName',$ComputerName)}

    if ($Credential) {$Parameters.Add('Credential',$Credential)}

    Invoke-Command @Parameters |
        
    Sort-Object @{Expression='TimeCreated';Descending='True'}, @{Expression='RecordId';Descending='True'}

}#Get-QEVEvents

# Function to build Main form
function QEV-Form-Main {

    $Form_Main = New-Object Windows.Forms.Form
    $Form_Main.Text = $Title 
    $Form_Main.Size = New-Object System.Drawing.Size @(680,410)
    $Form_Main.StartPosition = "CenterScreen"
    $Form_Main.FormBorderStyle = "FixedSingle"
    $Form_Main.MaximizeBox = $false
    $Form_Main.MinimizeBox = $false
    $Form_Main.Activate()

    $DateTimePicker_StartTime = New-Object System.Windows.Forms.DateTimePicker
    $DateTimePicker_StartTime.Location = New-Object System.Drawing.Point(60,30)
    $DateTimePicker_StartTime.Size = New-Object System.Drawing.Size(195,20)
    $DateTimePicker_StartTime.Format = "custom"
    $DateTimePicker_StartTime.CustomFormat = $DateTimeFormatLong
    $DateTimePicker_StartTime.Value = (Get-Date (Get-Date).AddHours(-$DefaultStartTime) -Format $DateTimeFormatLong)
    $Form_Main.Controls.Add($DateTimePicker_StartTime)

    $DateTimePicker_EndTime = New-Object System.Windows.Forms.DateTimePicker
    $DateTimePicker_EndTime.Location = New-Object System.Drawing.Point(60,74)
    $DateTimePicker_EndTime.Size = New-Object System.Drawing.Size(195,20)
    $DateTimePicker_EndTime.Format = "custom"
    $DateTimePicker_EndTime.CustomFormat = $DateTimeFormatLong
    $Form_Main.Controls.Add($DateTimePicker_EndTime)

    $CheckedListBox_Level = New-Object System.Windows.Forms.CheckedListBox
    $CheckedListBox_Level.Location = New-Object System.Drawing.Point(260,30)
    $CheckedListBox_Level.Size = New-Object System.Drawing.Size(80,64)
    "Critical","Error","Warning","Information" | ForEach-Object {$CheckedListBox_Level.Items.Add($_) | Out-Null}
    0..2 | ForEach-Object {$CheckedListBox_Level.SetItemChecked($_, $true)}
    $CheckedListBox_Level.CheckOnClick = $true
    $Form_Main.Controls.Add($CheckedListBox_Level)

    $TextBox_ComputerName = New-Object System.Windows.Forms.TextBox
    $TextBox_ComputerName.Location = New-Object System.Drawing.Point(345,30) 
    $TextBox_ComputerName.Size = New-Object System.Drawing.Size(126,20)
    $TextBox_ComputerName.Text = $ComputerName
    $Form_Main.Controls.Add($TextBox_ComputerName)

    $TextBox_UserName = New-Object System.Windows.Forms.TextBox
    $TextBox_UserName.Location = New-Object System.Drawing.Point(476,30) 
    $TextBox_UserName.Size = New-Object System.Drawing.Size(126,20)
    $TextBox_UserName.Text = $UserName
    $Form_Main.Controls.Add($TextBox_UserName)

    $TextBox_EventID = New-Object System.Windows.Forms.TextBox
    $TextBox_EventID.Location = New-Object System.Drawing.Point(345,74) 
    $TextBox_EventID.Size = New-Object System.Drawing.Size(126,20)
    $Form_Main.Controls.Add($TextBox_EventID)

    $TextBox_Message = New-Object System.Windows.Forms.TextBox
    $TextBox_Message.Location = New-Object System.Drawing.Point(476,74) 
    $TextBox_Message.Size = New-Object System.Drawing.Size(126,20)
    $Form_Main.Controls.Add($TextBox_Message)

    $ListView_Logs = New-Object System.Windows.Forms.ListView
    $ListView_Logs.Location = New-Object System.Drawing.Point(60, 126)
    $ListView_Logs.Size = New-Object System.Drawing.Size(544, 190)
    $ListView_Logs.HideSelection = $false
    $ListView_Logs.FullRowSelect = $true
    $ListView_Logs.AllowColumnReorder = $true
    $ListView_Logs.View = "Details"
    $ListView_Logs.Columns.Add('LogName', 375) | Out-Null
    $ListView_Logs.Columns.Add('Count', 48) | Out-Null
    $ListView_Logs.Columns.Add('LastWrite', 100) | Out-Null

    # Add logs to the ListView
    Get-QEVEventLogList -ComputerName $TextBox_ComputerName.Text.Split(' ') -Filter 'Application', 'System', 'Security' | 

    ForEach-Object {

        $Item = New-Object System.Windows.Forms.ListViewItem($_.LogName)
        $Item.SubItems.Add($_.RecordCount)
        $Item.SubItems.Add("$(Get-Date $_.ActLastWriteTime -Format $DateTimeFormatShort)")
        $ListView_Logs.Items.Add($Item)
    } | Out-Null

    # Add column click - sort rows
    $ListView_Logs.add_ColumnClick({

        if ($LastColumn -eq $_.Column) {$Script:Toggle = -not $Toggle}
        else {$Script:Toggle = 0}
        $Script:LastColumn = $_.Column
        $List = @()
        $SortAsDate = $true
        $SortAsInt = $true
        $Date = [datetime]"00:00"
        $Int = [int]1

        $ListView_Logs.Items | ForEach-Object {

            $CheckDateTime = [datetime]::TryParse($_.SubItems.text[$LastColumn], [ref]$Date)
            $CheckInt = [int]::TryParse($_.SubItems.text[$LastColumn], [ref]$Int)
            if (-not ($CheckDateTime)) {$SortAsDate = $false}
            if (-not ($CheckInt)) {$SortAsInt = $false}
            $List += $_
        }

        $ListView_Logs.BeginUpdate()
        $ListView_Logs.Items.Clear()
        if ($SortAsDate) {

            $ListView_Logs.Items.AddRange(@($List | Sort-Object @{Expression={Get-Date $_.SubItems.Text[$LastColumn]};Ascending=$Toggle}))
        }
        elseif ($SortAsInt) {
            
            $ListView_Logs.Items.AddRange(@($List | Sort-Object @{Expression={[int]$_.SubItems.Text[$LastColumn]};Ascending=$Toggle}))
        }
        else {
            
            $ListView_Logs.Items.AddRange(@($List | Sort-Object @{Expression={
                if ($_.SubItems.Text[$LastColumn] -like "*Microsoft-*") {"zz$($_.SubItems.Text[$LastColumn])"}
                else {"$($_.SubItems.Text[$LastColumn])"}
            };Ascending=!$Toggle}))
        }
        $ListView_Logs.EndUpdate()
    })
    $Form_Main.Controls.Add($ListView_Logs)

    $Label_DTP_Start = New-Object System.Windows.Forms.Label
    $Label_DTP_Start.Location = New-Object System.Drawing.Point(60,15) 
    $Label_DTP_Start.Size = New-Object System.Drawing.Size(60,13)
    $Label_DTP_Start.Text = "Start Time"
    $Form_Main.Controls.Add($Label_DTP_Start)

    $Label_DTP_End = New-Object System.Windows.Forms.Label
    $Label_DTP_End.Location = New-Object System.Drawing.Point(60,59)
    $Label_DTP_End.Size = New-Object System.Drawing.Size(60,13)
    $Label_DTP_End.Text = "End Time"
    $Form_Main.Controls.Add($Label_DTP_End)

    $Label_CLB_Level = New-Object System.Windows.Forms.Label
    $Label_CLB_Level.Location = New-Object System.Drawing.Point(260,15) 
    $Label_CLB_Level.Size = New-Object System.Drawing.Size(60,13)
    $Label_CLB_Level.Text = "Level"
    $Form_Main.Controls.Add($Label_CLB_Level)

    $Label_TB_ComputerName = New-Object System.Windows.Forms.Label
    $Label_TB_ComputerName.Location = New-Object System.Drawing.Point(345,15) 
    $Label_TB_ComputerName.Size = New-Object System.Drawing.Size(100,13)
    $Label_TB_ComputerName.Text = "ComputerName(s)"
    $Form_Main.Controls.Add($Label_TB_ComputerName)

    $Label_TB_UserName = New-Object System.Windows.Forms.Label
    $Label_TB_UserName.Location = New-Object System.Drawing.Point(476,15) 
    $Label_TB_UserName.Size = New-Object System.Drawing.Size(100,13)
    $Label_TB_UserName.Text = "UserName"
    $Form_Main.Controls.Add($Label_TB_UserName)

    $Label_TB_EventID = New-Object System.Windows.Forms.Label
    $Label_TB_EventID.Location = New-Object System.Drawing.Point(345,59) 
    $Label_TB_EventID.Size = New-Object System.Drawing.Size(85,13)
    $Label_TB_EventID.Text = "Event ID(s)"
    $Form_Main.Controls.Add($Label_TB_EventID)

    $Label_TB_Message = New-Object System.Windows.Forms.Label
    $Label_TB_Message.Location = New-Object System.Drawing.Point(476,59) 
    $Label_TB_Message.Size = New-Object System.Drawing.Size(85,13)
    $Label_TB_Message.Text = "Message"
    $Form_Main.Controls.Add($Label_TB_Message)

    $LinkLabel_ShowAll = New-Object System.Windows.Forms.LinkLabel
    $LinkLabel_ShowAll.Location = New-Object System.Drawing.Point(60,111)
    $LinkLabel_ShowAll.Size = New-Object System.Drawing.Size(51,13)
    $LinkLabel_ShowAll.Text = "Show All"
    $LinkLabel_ShowAll.Add_Click({

        $ListView_Logs.Items.Clear()
        $ListView_Logs.BeginUpdate()
        Get-QEVEventLogList -ComputerName $TextBox_ComputerName.Text.Split(' ') | ForEach-Object {

            $Item = New-Object System.Windows.Forms.ListViewItem($_.LogName)
            $Item.SubItems.Add($_.RecordCount)
            $Item.SubItems.Add("$(Get-Date $_.ActLastWriteTime -Format $DateTimeFormatShort)")
            $ListView_Logs.Items.Add($Item)

        } | Out-Null

        $ListView_Logs.EndUpdate()
        $ListView_Logs.Focus()
    })
    $Form_Main.Controls.Add($LinkLabel_ShowAll)

    $LinkLabel_SelectAll = New-Object System.Windows.Forms.LinkLabel
    $LinkLabel_SelectAll.Location = New-Object System.Drawing.Point(464,111)
    $LinkLabel_SelectAll.Size = New-Object System.Drawing.Size(51,13)
    $LinkLabel_SelectAll.Text = "Select All"
    $LinkLabel_SelectAll.Add_click({

        $ListView_Logs.BeginUpdate()
        $ListView_Logs.Items | ForEach-Object {$_.Selected = $true}
        $ListView_Logs.EndUpdate()
        $ListView_Logs.Focus()
    })
    $Form_Main.Controls.Add($LinkLabel_SelectAll)

    $LinkLabel_Invert = New-Object System.Windows.Forms.LinkLabel
    $LinkLabel_Invert.Location = New-Object System.Drawing.Point(521,111)
    $LinkLabel_Invert.Size = New-Object System.Drawing.Size(82,13)
    $LinkLabel_Invert.Text = "Invert Selection"
    $LinkLabel_Invert.Add_click({

        $ListView_Logs.BeginUpdate()
        $ListView_Logs.Items | ForEach-Object {$_.Selected = -not $_.Selected}
        $ListView_Logs.EndUpdate()
        $ListView_Logs.Focus()
    })
    $Form_Main.Controls.Add($LinkLabel_Invert)

    $Button_View = New-Object System.Windows.Forms.Button
    $Button_View.Location = New-Object System.Drawing.Point(257,336)
    $Button_View.Size = New-Object System.Drawing.Size(75,23)
    $Button_View.Text = "View"
    $Button_View.Add_Click({QEV-Form-Viewer})
    $Form_Main.Controls.Add($Button_View)

    $Button_Exit = New-Object System.Windows.Forms.Button
    $Button_Exit.Location = New-Object System.Drawing.Point(340,336)
    $Button_Exit.Size = New-Object System.Drawing.Size(75,23)
    $Button_Exit.Text = "Exit"
    $Button_Exit.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $Form_Main.CancelButton = $Button_Exit
    $Form_Main.Controls.Add($Button_Exit)

    $Form_Main.ShowDialog() | Out-Null

}#QEV-Form-Main

# Function to build Viewer form
function QEV-Form-Viewer {

    $Button_View.Enabled = $false

    $Paramaters = @{

		ComputerName = $TextBox_ComputerName.Text.Split(' ')
		StartTime = Get-Date $DateTimePicker_StartTime.Text
		EndTime = Get-Date $DateTimePicker_EndTime.Text
		Level = $CheckedListBox_Level.CheckedIndices | ForEach-Object {$_ + 1}
		Logs = @($ListView_Logs.SelectedItems.Text)
        IDs = $TextBox_EventID.Text.Split(' ')
        Message = $TextBox_Message.Text
        UserName = $TextBox_UserName.Text
        MaxEvents = $MaxEvents
    }

    $Events = Get-QEVEvents @Paramaters

    $EventProperties = @{property = @(
            
        @{label='Level';expression={$_.LevelDisplayName}},
        'TimeCreated',
        @{label='ProviderName';expression={$_.ProviderName.replace('Microsoft-Windows-','')}},
        @{label='EventID';expression={$_.ID}},
        @{label='TaskCategory';expression={$_.TaskDisplayName}}
    )}

    # Needed if only 1 event is returned
    $ArrayList_Events = New-Object System.collections.ArrayList
    if (($Events| Measure-Object).Count -eq 1) {$ArrayList_Events.Add(($Events | Select-Object @EventProperties))}
    else {$ArrayList_Events.AddRange(($Events | Select-Object @EventProperties))}

    $Form_Viewer = New-Object System.Windows.Forms.Form
    $Form_Viewer.Size = New-Object System.Drawing.Size(800,600)
    $Form_Viewer.StartPosition = "CenterScreen"
    $Form_Viewer.FormBorderStyle = "FixedSingle"
    $Form_Viewer.Text = $Title + " - $(($Events | Measure-Object).Count)"
    $Form_Viewer.MaximizeBox = $false
    $Form_Viewer.MinimizeBox = $false
    $Form_Viewer.Activate()

    $DataGridView_Events = New-Object System.Windows.Forms.DataGridView
    $DataGridView_Events.Location = New-Object System.Drawing.Point(12,12)
    $DataGridView_Events.Size = New-Object System.Drawing.Size(760,308)
    $DataGridView_Events.VirtualMode = $true
    $DataGridView_Events.AutoSizeColumnsMode = "Fill"
    $DataGridView_Events.ColumnHeadersVisible = $true
    $DataGridView_Events.ColumnHeadersHeightSizeMode = 'DisableResizing'
    $DataGridView_Events.RowHeadersVisible = $false
    $DataGridView_Events.AllowUserToOrderColumns = $true
    $DataGridView_Events.AllowUserToResizeRows = $false
    $DataGridView_Events.MultiSelect = $false
    $DataGridView_Events.DataSource = $ArrayList_Events
    $DataGridView_Events.ReadOnly = $true
    $DataGridView_Events.SelectionMode = 'FullRowSelect'
    $DataGridView_Events.BackgroundColor = 'White'
    $DataGridView_Events.AlternatingRowsDefaultCellStyle.BackColor = 'WhiteSmoke'
    $Form_Viewer.Controls.Add($DataGridView_Events)

    $Label_LogName = New-Object System.Windows.Forms.Label
    $Label_LogName.Location = New-Object System.Drawing.Point(12,332) 
    $Label_LogName.Size = New-Object System.Drawing.Size(350,13)
    $Form_Viewer.Controls.Add($Label_LogName)

    $Label_User = New-Object System.Windows.Forms.Label
    $Label_User.Location = New-Object System.Drawing.Point(12,355) 
    $Label_User.Size = New-Object System.Drawing.Size(350,13)
    $Form_Viewer.Controls.Add($Label_User)

    $Label_Computer = New-Object System.Windows.Forms.Label
    $Label_Computer.Location = New-Object System.Drawing.Point(412,332) 
    $Label_Computer.Size = New-Object System.Drawing.Size(350,13)
    $Form_Viewer.Controls.Add($Label_Computer)

    $Label_RecordId = New-Object System.Windows.Forms.Label
    $Label_RecordId.Location = New-Object System.Drawing.Point(412,355) 
    $Label_RecordId.Size = New-Object System.Drawing.Size(350,13)
    $Form_Viewer.Controls.Add($Label_RecordId)

    $TextBox_EventMessage = New-Object System.Windows.Forms.TextBox
    $TextBox_EventMessage.Location = New-Object System.Drawing.Point(12,379)
    $TextBox_EventMessage.Size = New-Object System.Drawing.Size(760,170)
    $TextBox_EventMessage.ReadOnly = $true
    $TextBox_EventMessage.Multiline = $true
    $TextBox_EventMessage.AcceptsReturn = $true
    $TextBox_EventMessage.ScrollBars = 'Vertical'
    $TextBox_EventMessage.Text =  $Events[0].Message
    $Form_Viewer.Controls.Add($TextBox_EventMessage)

    $DataGridView_Events.Add_SelectionChanged({
                
        if ($Events[$DataGridView_Events.SelectedRows.Index].Properties.Value) {
        
            $TextBox_EventMessage.Text = `
                ($Events[$DataGridView_Events.SelectedRows.Index].Message | Out-String) `
                + "`r`n" + "`r`n" + "Event Details:" + "`r`n" `
                + ($Events[$DataGridView_Events.SelectedRows.Index].Properties.Value | Out-String)
        }

        # | Where-Object {$_ -isnot [byte]}

        else {($TextBox_EventMessage.Text = $Events[$DataGridView_Events.SelectedRows.Index].Message | Out-String)}

        $Label_LogName.Text = "Log: $($Events[$DataGridView_Events.SelectedRows.Index].LogName.Replace('Microsoft-Windows-',''))"

        $Label_User.Text = "User: $($Events[$DataGridView_Events.SelectedRows.Index].NTAccount)"

        $Label_Computer.Text = "Computer: $($Events[$DataGridView_Events.SelectedRows.Index].MachineName)"

        $Label_RecordId.Text = "Record: $($Events[$DataGridView_Events.SelectedRows.Index].RecordId)"
    })

    $Form_Viewer.ShowDialog()

    $Button_View.Enabled = $true

}#QEV-Form-Viewer

# Main function
function Start-QuickEventViewer {

    <#
    .NOTES
        ######################
         mail@nimbus117.co.uk
        ######################
        
        ## lastwritetime from 'get-winevent -listlog' wrong - workaround - check latest event lastwritetime in each log

        ## Returned list will exclude empty logs - Where-Object {$_.RecordCount}

        ## Security log events are level 0 (not 1-4), treat as infomation(4) - if $level contains 4 add 0 to array

        ## Level and message not returned on windows 7/2008 - workaround = change culture to US before get-winevent and then back after - partially works

    .SYNOPSIS
        QEV - Quick Event Viewer

        PS C:\>Start-QuickEventViewer
    #>

    param(
        [Parameter(Position=0)][string[]]$ComputerName,
        [Parameter(Position=1)][PSCredential][System.Management.Automation.CredentialAttribute()]$Credential,
        [int]$MaxEvents = 200
    )

    $Title = "QEV - Quick Event Viewer"

    $DefaultStartTime = 24 # Hours back

    $CurrentCulture = Get-Culture
    $DateTimeFormatLong = $CurrentCulture.DateTimeFormat.LongDatePattern + " " + $CurrentCulture.DateTimeFormat.ShortTimePattern
    $DateTimeFormatShort = $CurrentCulture.DateTimeFormat.ShortDatePattern + " " + $CurrentCulture.DateTimeFormat.ShortTimePattern

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    QEV-Form-Main

}#Start-QuickEventViewer
Set-Alias QEV -Value Start-QuickEventViewer

Export-ModuleMember -Function Start-QuickEventViewer,Get-QEVEventLogList,Get-QEVEvents -Alias * 