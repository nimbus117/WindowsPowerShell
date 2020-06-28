# Main form - connect to vcenter, resource limits, excluded hosts, csv/manual buttons

# csv form - launch file browser, automatically run validate csv/resources and then prompt 
# with results and option to continue with warning that some vms may fail to deploy ---- option to preview in out-gridview first????

# Main Form
function Deploy-VMGui {

    Add-Type -AssemblyName System.Windows.Forms

    $Form_Main = New-Object system.Windows.Forms.Form
    $Form_Main.Text = "Deploy-VM"
    $Form_Main.StartPosition = "CenterScreen"
    $Form_Main.FormBorderStyle = "FixedSingle"
    $Form_Main.Width = 270
    $Form_Main.Height = 430

    $TextBox_vCenter = New-Object system.windows.Forms.TextBox
    $TextBox_vCenter.Text = "vCenter IP or name"
    $TextBox_vCenter.Width = 132
    $TextBox_vCenter.Height = 20
    $TextBox_vCenter.location = new-object system.drawing.point(61,10)
    $TextBox_vCenter.Font = "Microsoft Sans Serif,10"
    $Form_Main.controls.Add($TextBox_vCenter)

    $Button_Connect = New-Object system.windows.Forms.Button
    $Button_Connect.Text = "Connect"
    $Button_Connect.Width = 70
    $Button_Connect.Height = 30
    $Button_Connect.location = new-object system.drawing.point(91,35)
    $Button_Connect.Font = "Microsoft Sans Serif,10"
    $Button_Connect.Add_Click({

        #$Button_Connect.Enabled = $false
    
        if ($global:DefaultVIServer.Name -ne $TextBox_vCenter.Text) {Connect-vCenter -Server $TextBox_vCenter.Text}
        if ($global:DefaultVIServer.Name -eq $TextBox_vCenter.Text) {
            
            $ListBox_ExcludedHosts.DataSource = (Get-VMHost | Sort-Object Name)
            $ListBox_ExcludedHosts.SelectedIndex = -1
            $NumericUpDown_MaxVMHostRamProv.Enabled = $true
            $NumericUpDown_DsUsagePercentThreshold.Enabled = $true
            $NumericUpDown_MaxVCpuPerCore.Enabled = $true
            $NumericUpDown_DSMaxLiveVM.Enabled = $true
            $ListBox_ExcludedHosts.Enabled = $true
            $Button_CSV.Enabled = $true
            $Button_Manual.Enabled = $true
        }

        #$Button_Connect.Enabled = $true
    })
    $Form_Main.controls.Add($Button_Connect)

    $Label_MaxVMHostRamProv = New-Object system.windows.Forms.Label
    $Label_MaxVMHostRamProv.Text = "Max Memory Usage (%)"
    $Label_MaxVMHostRamProv.AutoSize = $true
    $Label_MaxVMHostRamProv.Width = 25
    $Label_MaxVMHostRamProv.Height = 10
    $Label_MaxVMHostRamProv.location = new-object system.drawing.point(10,80)
    $Label_MaxVMHostRamProv.Font = "Microsoft Sans Serif,10"
    $Form_Main.controls.Add($Label_MaxVMHostRamProv)

    $Label_DsUsagePercentThreshold = New-Object system.windows.Forms.Label
    $Label_DsUsagePercentThreshold.Text = "Max Datastore usage (%)"
    $Label_DsUsagePercentThreshold.AutoSize = $true
    $Label_DsUsagePercentThreshold.Width = 25
    $Label_DsUsagePercentThreshold.Height = 10
    $Label_DsUsagePercentThreshold.location = new-object system.drawing.point(10,110)
    $Label_DsUsagePercentThreshold.Font = "Microsoft Sans Serif,10"
    $Form_Main.controls.Add($Label_DsUsagePercentThreshold)

    $Label_MaxVCpuPerCore = New-Object system.windows.Forms.Label
    $Label_MaxVCpuPerCore.Text = "Max vCPU/Core"
    $Label_MaxVCpuPerCore.AutoSize = $true
    $Label_MaxVCpuPerCore.Width = 25
    $Label_MaxVCpuPerCore.Height = 10
    $Label_MaxVCpuPerCore.location = new-object system.drawing.point(10,140)
    $Label_MaxVCpuPerCore.Font = "Microsoft Sans Serif,10"
    $Form_Main.controls.Add($Label_MaxVCpuPerCore)

    $Label_DSMaxLiveVM = New-Object system.windows.Forms.Label
    $Label_DSMaxLiveVM.Text = "Max VM/Datastore"
    $Label_DSMaxLiveVM.AutoSize = $true
    $Label_DSMaxLiveVM.Width = 25
    $Label_DSMaxLiveVM.Height = 10
    $Label_DSMaxLiveVM.location = new-object system.drawing.point(10,170)
    $Label_DSMaxLiveVM.Font = "Microsoft Sans Serif,10"
    $Form_Main.controls.Add($Label_DSMaxLiveVM)

    $ListBox_ExcludedHosts = New-Object system.windows.Forms.ListBox
    $ListBox_ExcludedHosts.Text = "listBox"
    $ListBox_ExcludedHosts.Width = 142
    $ListBox_ExcludedHosts.Height = 82
    $ListBox_ExcludedHosts.enabled = $false
    $ListBox_ExcludedHosts.ScrollAlwaysVisible = $true
    $ListBox_ExcludedHosts.SelectionMode = 'MultiExtended'
    $ListBox_ExcludedHosts.location = new-object system.drawing.point(55,232)
    $Form_Main.controls.Add($ListBox_ExcludedHosts)

    $NumericUpDown_MaxVMHostRamProv = New-Object system.windows.Forms.NumericUpDown
    $NumericUpDown_MaxVMHostRamProv.Text = "80"
    $NumericUpDown_MaxVMHostRamProv.Width = 42
    $NumericUpDown_MaxVMHostRamProv.Height = 20
    $NumericUpDown_MaxVMHostRamProv.enabled = $false
    $NumericUpDown_MaxVMHostRamProv.location = new-object system.drawing.point(190,80)
    $NumericUpDown_MaxVMHostRamProv.Font = "Microsoft Sans Serif,10"
    $NumericUpDown_MaxVMHostRamProv.Maximum = 100
    $Form_Main.controls.Add($NumericUpDown_MaxVMHostRamProv)

    $NumericUpDown_DsUsagePercentThreshold = New-Object system.windows.Forms.NumericUpDown
    $NumericUpDown_DsUsagePercentThreshold.Text = "80"
    $NumericUpDown_DsUsagePercentThreshold.Width = 42
    $NumericUpDown_DsUsagePercentThreshold.Height = 20
    $NumericUpDown_DsUsagePercentThreshold.enabled = $false
    $NumericUpDown_DsUsagePercentThreshold.location = new-object system.drawing.point(190,110)
    $NumericUpDown_DsUsagePercentThreshold.Font = "Microsoft Sans Serif,10"
    $NumericUpDown_DsUsagePercentThreshold.Maximum = 100
    $Form_Main.controls.Add($NumericUpDown_DsUsagePercentThreshold)

    $NumericUpDown_MaxVCpuPerCore = New-Object system.windows.Forms.NumericUpDown
    $NumericUpDown_MaxVCpuPerCore.Text = "5"
    $NumericUpDown_MaxVCpuPerCore.Width = 42
    $NumericUpDown_MaxVCpuPerCore.Height = 20
    $NumericUpDown_MaxVCpuPerCore.enabled = $false
    $NumericUpDown_MaxVCpuPerCore.location = new-object system.drawing.point(190,140)
    $NumericUpDown_MaxVCpuPerCore.Font = "Microsoft Sans Serif,10"
    $NumericUpDown_MaxVCpuPerCore.Maximum = 20
    $Form_Main.controls.Add($NumericUpDown_MaxVCpuPerCore)

    $NumericUpDown_DSMaxLiveVM = New-Object system.windows.Forms.NumericUpDown
    $NumericUpDown_DSMaxLiveVM.Text = "20"
    $NumericUpDown_DSMaxLiveVM.Width = 42
    $NumericUpDown_DSMaxLiveVM.Height = 20
    $NumericUpDown_DSMaxLiveVM.enabled = $false
    $NumericUpDown_DSMaxLiveVM.location = new-object system.drawing.point(190,170)
    $NumericUpDown_DSMaxLiveVM.Font = "Microsoft Sans Serif,10"
    $NumericUpDown_DSMaxLiveVM.Maximum = 100
    $Form_Main.controls.Add($NumericUpDown_DSMaxLiveVM)

    $Button_CSV = New-Object system.windows.Forms.Button
    $Button_CSV.Text = "CSV"
    $Button_CSV.Width = 60
    $Button_CSV.Height = 30
    $Button_CSV.enabled = $false
    $Button_CSV.location = new-object system.drawing.point(10,337)
    $Button_CSV.Font = "Microsoft Sans Serif,10"
    $Button_CSV.Add_Click({Form-CSV})
    $Form_Main.controls.Add($Button_CSV)

    $Button_Manual = New-Object system.windows.Forms.Button
    $Button_Manual.Text = "Manual"
    $Button_Manual.Width = 60
    $Button_Manual.Height = 30
    $Button_Manual.enabled = $false
    $Button_Manual.location = new-object system.drawing.point(180,336)
    $Button_Manual.Font = "Microsoft Sans Serif,10"
    $Button_Manual.Add_Click({Form-Manual})
    $Form_Main.controls.Add($Button_Manual)

    $Label_ExcludedHosts = New-Object system.windows.Forms.Label
    $Label_ExcludedHosts.Text = "Excluded Hosts"
    $Label_ExcludedHosts.AutoSize = $true
    $Label_ExcludedHosts.Width = 25
    $Label_ExcludedHosts.Height = 10
    $Label_ExcludedHosts.location = new-object system.drawing.point(80,211)
    $Label_ExcludedHosts.Font = "Microsoft Sans Serif,10"
    $Form_Main.controls.Add($Label_ExcludedHosts)

    [void]$Form_Main.ShowDialog()
    $Form_Main.Dispose()

}

# Deploy from CSV
function Form-CSV {

    $OpenFile_CSV = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFile_CSV.InitialDirectory = (Get-location).Path
    $OpenFile_CSV.Title = "Open CSV"
    $OpenFile_CSV.MultiSelect = $false
    $OpenFile_CSV.CheckFileExists = $true
    $OpenFile_CSV.CheckPathExists = $true
    $OpenFile_CSV.Filter = "Txt Files|*.csv"
    $OpenFile_CSV.ShowDialog() | Out-Null
    $CSVPath = $OpenFile_CSV.FileName
    $OpenFile_CSV.Dispose()

    if ($CSVPath) {

        $Form_CSVValidate = New-Object system.Windows.Forms.Form
        $Form_CSVValidate.Text = "Validate"
        $Form_CSVValidate.TopMost = $true
        $Form_CSVValidate.Width = 600
        $Form_CSVValidate.Height = 500
        $Form_CSVValidate.StartPosition = "CenterScreen"
        $Form_CSVValidate.FormBorderStyle = "FixedSingle"
        $Form_CSVValidate.TopMost = $false

        $TextBox_CSVValidate = New-Object system.windows.Forms.TextBox
        $TextBox_CSVValidate.Multiline = $true
        $TextBox_CSVValidate.Width = 560
        $TextBox_CSVValidate.Height = 340
        $TextBox_CSVValidate.location = new-object system.drawing.point(10,20)
        $TextBox_CSVValidate.Font = "Microsoft Sans Serif,10"
        $TextBox_CSVValidate.ScrollBars = 'Vertical'
        $TextBox_CSVValidate.ReadOnly = $true
        $Form_CSVValidate.controls.Add($TextBox_CSVValidate)

        $Button_CSVValidate = New-Object system.windows.Forms.Button
        $Button_CSVValidate.Text = "Validate"
        $Button_CSVValidate.Width = 65
        $Button_CSVValidate.Height = 30
        $Button_CSVValidate.location = new-object system.drawing.point(150,395)
        $Button_CSVValidate.Font = "Microsoft Sans Serif,10"
        $Button_CSVValidate.Add_Click({
        
            $Button_CSVValidate.Enabled = $false

            $TextBox_CSVValidate.Text = "Validating CSV"

            $ValidateCSV = Validate-CSVContent -CSV $CSVPath -ErrorAction SilentlyContinue | select -Unique

            if ($ValidateCSV) {$TextBox_CSVValidate.Text = $ValidateCSV}
        
            else {
            
                $TextBox_CSVValidate.Text = "Validating Resources"

                $ValidateResources_Params = @{

                    Collection = (Import-Csv $CSVPath) 
                    VMHostExcluded = $ListBox_ExcludedHosts.SelectedItems.Name
                    DSMaxLiveVM = $NumericUpDown_DSMaxLiveVM.Text
                    DsUsagePercentThreshold = $NumericUpDown_DsUsagePercentThreshold.Text
                    MaxVCpuPerCore = $NumericUpDown_MaxVCpuPerCore.Text
                    MaxVMHostRamProv = $NumericUpDown_MaxVMHostRamProv.Text
                }

                $ValidateResources = Validate-Resources @ValidateResources_Params

                if ($ValidateResources -eq $true) {
            
                    $TextBox_CSVValidate.Text = "Validation successful"

                    $Button_CSVDeploy.Enabled = $true
                }
                else {$TextBox_CSVValidate.Text = $ValidateResources}
            }

            $Button_CSVValidate.Enabled = $true
        })
        $Form_CSVValidate.controls.Add($Button_CSVValidate)

        $Button_CSVDeploy = New-Object system.windows.Forms.Button
        $Button_CSVDeploy.Text = "Deploy"
        $Button_CSVDeploy.Width = 65
        $Button_CSVDeploy.Height = 30
        $Button_CSVDeploy.location = new-object system.drawing.point(350,396)
        $Button_CSVDeploy.Font = "Microsoft Sans Serif,10"
        $Button_CSVDeploy.Enabled = $false
        $Button_CSVDeploy.Add_Click({
    
            $Button_CSVDeploy.Enabled = $false

            $CSV_Params = @{

                DSMaxLiveVM = $NumericUpDown_DSMaxLiveVM.Text
                DsUsagePercentThreshold = $NumericUpDown_DsUsagePercentThreshold.Text
                MaxVCpuPerCore = $NumericUpDown_MaxVCpuPerCore.Text
                MaxVMHostRamProv = $NumericUpDown_MaxVMHostRamProv.Text
            }

            if ($ListBox_ExcludedHosts.SelectedItem) {$CSV_Params += @{ExcludedVMHost = $ListBox_ExcludedHosts.SelectedItems.Name}}

            $CSVDomainCreds = try {Get-Credential -Message "Enter Domain Credential"} catch{}

            if ($CSVDomainCreds) {
        
                $CSV_Params += @{DomainCredential = $CSVDomainCreds}
            }
        
            Import-Csv $CSVPath | Deploy-VM @CSV_Params
            #Import-Csv $CSVPath | ogv
            #$CSV_Params | out-file C:\Users\jamesg\Desktop\test.txt

            $Form_CSVValidate.Dispose()
        })
        $Form_CSVValidate.controls.Add($Button_CSVDeploy)

        [void]$Form_CSVValidate.ShowDialog()
        $Form_CSVValidate.Dispose()
    }
}

# Deploy Manually
function Form-Manual {
    
    $Form_Manual = New-Object system.Windows.Forms.Form
    $Form_Manual.Text = "Manual"
    $Form_Manual.TopMost = $false
    $Form_Manual.Width = 380
    $Form_Manual.Height = 580
    $Form_Manual.StartPosition = "CenterScreen"
    $Form_Manual.FormBorderStyle = "FixedSingle"

    $TextBox_VMName = New-Object system.windows.Forms.TextBox
    $TextBox_VMName.Width = 180
    $TextBox_VMName.Height = 20
    $TextBox_VMName.location = new-object system.drawing.point(120,20)
    $TextBox_VMName.Font = "Microsoft Sans Serif,10"
    $TextBox_VMName.MaxLength = 15
    $Form_Manual.controls.Add($TextBox_VMName)

    $Label_VMName = New-Object system.windows.Forms.Label
    $Label_VMName.Text = "VMName *"
    $Label_VMName.AutoSize = $true
    $Label_VMName.Width = 25
    $Label_VMName.Height = 10
    $Label_VMName.location = new-object system.drawing.point(10,20)
    $Label_VMName.Font = "Microsoft Sans Serif,10"
    $Form_Manual.controls.Add($Label_VMName)

    $ComboBox_Template = New-Object system.windows.Forms.ComboBox
    $ComboBox_Template.Width = 180
    $ComboBox_Template.Height = 20
    $ComboBox_Template.location = new-object system.drawing.point(120,50)
    $ComboBox_Template.Font = "Microsoft Sans Serif,10"
    $ComboBox_Template.DropDownStyle = "DropDownList"
    $ComboBox_Template.DataSource = (Get-Template | Sort-Object Name)
    $Form_Manual.controls.Add($ComboBox_Template)

    $Label_Template = New-Object system.windows.Forms.Label
    $Label_Template.Text = "Template *"
    $Label_Template.AutoSize = $true
    $Label_Template.Width = 25
    $Label_Template.Height = 10
    $Label_Template.location = new-object system.drawing.point(10,50)
    $Label_Template.Font = "Microsoft Sans Serif,10"
    $Form_Manual.controls.Add($Label_Template)

    $Label_Datastore = New-Object system.windows.Forms.Label
    $Label_Datastore.Text = "Datastore *"
    $Label_Datastore.AutoSize = $true
    $Label_Datastore.Width = 25
    $Label_Datastore.Height = 10
    $Label_Datastore.location = new-object system.drawing.point(10,80)
    $Label_Datastore.Font = "Microsoft Sans Serif,10"
    $Form_Manual.controls.Add($Label_Datastore)

    $ComboBox_Datastore = New-Object system.windows.Forms.ComboBox
    $ComboBox_Datastore.Width = 180
    $ComboBox_Datastore.Height = 20
    $ComboBox_Datastore.location = new-object system.drawing.point(120,80)
    $ComboBox_Datastore.Font = "Microsoft Sans Serif,10"
    $ComboBox_Datastore.DataSource = (Get-Datastore | Sort-Object Name)
    $Form_Manual.controls.Add($ComboBox_Datastore)

    $Label_Folder = New-Object system.windows.Forms.Label
    $Label_Folder.Text = "Folder *"
    $Label_Folder.AutoSize = $true
    $Label_Folder.Width = 25
    $Label_Folder.Height = 10
    $Label_Folder.location = new-object system.drawing.point(10,110)
    $Label_Folder.Font = "Microsoft Sans Serif,10"
    $Form_Manual.controls.Add($Label_Folder)

    $ComboBox_Folder = New-Object system.windows.Forms.ComboBox
    $ComboBox_Folder.Width = 180
    $ComboBox_Folder.Height = 20
    $ComboBox_Folder.location = new-object system.drawing.point(120,110)
    $ComboBox_Folder.Font = "Microsoft Sans Serif,10"
    $ComboBox_Folder.DropDownStyle = "DropDownList"
    $ComboBox_Folder.DataSource = (Get-Folder | Sort-Object Name)
    $Form_Manual.controls.Add($ComboBox_Folder)

    $Label_OSType = New-Object system.windows.Forms.Label
    $Label_OSType.Text = "OSType *"
    $Label_OSType.AutoSize = $true
    $Label_OSType.Width = 25
    $Label_OSType.Height = 10
    $Label_OSType.location = new-object system.drawing.point(10,140)
    $Label_OSType.Font = "Microsoft Sans Serif,10"
    $Form_Manual.controls.Add($Label_OSType)

    $ComboBox_OSType = New-Object system.windows.Forms.ComboBox
    $ComboBox_OSType.Width = 180
    $ComboBox_OSType.Height = 20
    $ComboBox_OSType.location = new-object system.drawing.point(120,140)
    $ComboBox_OSType.Font = "Microsoft Sans Serif,10"
    $ComboBox_OSType.DropDownStyle = "DropDownList"
    $ComboBox_OSType.DataSource = "Windows","Linux"
    $Form_Manual.controls.Add($ComboBox_OSType)

    $Label_Cluster = New-Object system.windows.Forms.Label
    $Label_Cluster.Text = "Cluster"
    $Label_Cluster.AutoSize = $true
    $Label_Cluster.Width = 25
    $Label_Cluster.Height = 10
    $Label_Cluster.location = new-object system.drawing.point(10,170)
    $Label_Cluster.Font = "Microsoft Sans Serif,10"
    $Form_Manual.controls.Add($Label_Cluster)

    $ComboBox_Cluster = New-Object system.windows.Forms.ComboBox
    $ComboBox_Cluster.Width = 180
    $ComboBox_Cluster.Height = 20
    $ComboBox_Cluster.location = new-object system.drawing.point(120,170)
    $ComboBox_Cluster.Font = "Microsoft Sans Serif,10"
    $ComboBox_Cluster.DropDownStyle = "DropDownList"
    $ComboBox_Cluster.DataSource = (Get-Cluster | Sort-Object Name),""
    $Form_Manual.controls.Add($ComboBox_Cluster)

    $Label_NetworkName = New-Object system.windows.Forms.Label
    $Label_NetworkName.Text = "NetworkName"
    $Label_NetworkName.AutoSize = $true
    $Label_NetworkName.Width = 25
    $Label_NetworkName.Height = 10
    $Label_NetworkName.location = new-object system.drawing.point(10,200)
    $Label_NetworkName.Font = "Microsoft Sans Serif,10"
    $Form_Manual.controls.Add($Label_NetworkName)

    $ComboBox_NetworkName = New-Object system.windows.Forms.ComboBox
    $ComboBox_NetworkName.Width = 180
    $ComboBox_NetworkName.Height = 20
    $ComboBox_NetworkName.location = new-object system.drawing.point(120,200)
    $ComboBox_NetworkName.Font = "Microsoft Sans Serif,10"
    $ComboBox_NetworkName.DropDownStyle = "DropDownList"
    $ComboBox_NetworkName.DataSource = (Get-VirtualPortGroup | Sort-Object Name -Unique)
    $Form_Manual.controls.Add($ComboBox_NetworkName)

    $TextBox_IPAddress = New-Object system.windows.Forms.TextBox
    $TextBox_IPAddress.Width = 180
    $TextBox_IPAddress.Height = 20
    $TextBox_IPAddress.location = new-object system.drawing.point(120,230)
    $TextBox_IPAddress.Font = "Microsoft Sans Serif,10"
    $Form_Manual.controls.Add($TextBox_IPAddress)

    $Label_IPAddress = New-Object system.windows.Forms.Label
    $Label_IPAddress.Text = "IPAddress"
    $Label_IPAddress.AutoSize = $true
    $Label_IPAddress.Width = 25
    $Label_IPAddress.Height = 10
    $Label_IPAddress.location = new-object system.drawing.point(10,230)
    $Label_IPAddress.Font = "Microsoft Sans Serif,10"
    $Form_Manual.controls.Add($Label_IPAddress)

    $Label_DefaultGateway = New-Object system.windows.Forms.Label
    $Label_DefaultGateway.Text = "DefaultGateway"
    $Label_DefaultGateway.AutoSize = $true
    $Label_DefaultGateway.Width = 25
    $Label_DefaultGateway.Height = 10
    $Label_DefaultGateway.location = new-object system.drawing.point(10,260)
    $Label_DefaultGateway.Font = "Microsoft Sans Serif,10"
    $Form_Manual.controls.Add($Label_DefaultGateway)

    $TextBox_DefaultGateway = New-Object system.windows.Forms.TextBox
    $TextBox_DefaultGateway.Width = 180
    $TextBox_DefaultGateway.Height = 20
    $TextBox_DefaultGateway.location = new-object system.drawing.point(120,260)
    $TextBox_DefaultGateway.Font = "Microsoft Sans Serif,10"
    $Form_Manual.controls.Add($TextBox_DefaultGateway)

    $Label_SubnetMask = New-Object system.windows.Forms.Label
    $Label_SubnetMask.Text = "SubnetMask"
    $Label_SubnetMask.AutoSize = $true
    $Label_SubnetMask.Width = 25
    $Label_SubnetMask.Height = 10
    $Label_SubnetMask.location = new-object system.drawing.point(10,290)
    $Label_SubnetMask.Font = "Microsoft Sans Serif,10"
    $Form_Manual.controls.Add($Label_SubnetMask)

    $TextBox_SubnetMask = New-Object system.windows.Forms.TextBox
    $TextBox_SubnetMask.Width = 180
    $TextBox_SubnetMask.Height = 20
    $TextBox_SubnetMask.location = new-object system.drawing.point(120,290)
    $TextBox_SubnetMask.Font = "Microsoft Sans Serif,10"
    $Form_Manual.controls.Add($TextBox_SubnetMask)

    $Label_DNSServer = New-Object system.windows.Forms.Label
    $Label_DNSServer.Text = "DNSServer"
    $Label_DNSServer.AutoSize = $true
    $Label_DNSServer.Width = 25
    $Label_DNSServer.Height = 10
    $Label_DNSServer.location = new-object system.drawing.point(10,320)
    $Label_DNSServer.Font = "Microsoft Sans Serif,10"
    $Form_Manual.controls.Add($Label_DNSServer)

    $TextBox_DNSServer = New-Object system.windows.Forms.TextBox
    $TextBox_DNSServer.Width = 180
    $TextBox_DNSServer.Height = 20
    $TextBox_DNSServer.location = new-object system.drawing.point(120,320)
    $TextBox_DNSServer.Font = "Microsoft Sans Serif,10"
    $Form_Manual.controls.Add($TextBox_DNSServer)

    $Label_VMNote = New-Object system.windows.Forms.Label
    $Label_VMNote.Text = "VMNote"
    $Label_VMNote.AutoSize = $true
    $Label_VMNote.Width = 25
    $Label_VMNote.Height = 10
    $Label_VMNote.location = new-object system.drawing.point(10,350)
    $Label_VMNote.Font = "Microsoft Sans Serif,10"
    $Form_Manual.controls.Add($Label_VMNote)

    $TextBox_VMNote = New-Object system.windows.Forms.TextBox
    $TextBox_VMNote.Width = 180
    $TextBox_VMNote.Height = 20
    $TextBox_VMNote.location = new-object system.drawing.point(120,350)
    $TextBox_VMNote.Font = "Microsoft Sans Serif,10"
    $Form_Manual.controls.Add($TextBox_VMNote)

    $Label_Domain = New-Object system.windows.Forms.Label
    $Label_Domain.Text = "Domain"
    $Label_Domain.AutoSize = $true
    $Label_Domain.Width = 25
    $Label_Domain.Height = 10
    $Label_Domain.location = new-object system.drawing.point(10,380)
    $Label_Domain.Font = "Microsoft Sans Serif,10"
    $Form_Manual.controls.Add($Label_Domain)

    $TextBox_Domain = New-Object system.windows.Forms.TextBox
    $TextBox_Domain.Width = 180
    $TextBox_Domain.Height = 20
    $TextBox_Domain.location = new-object system.drawing.point(120,380)
    $TextBox_Domain.Font = "Microsoft Sans Serif,10"
    $Form_Manual.controls.Add($TextBox_Domain)

    $Label_AdminPassword = New-Object system.windows.Forms.Label
    $Label_AdminPassword.Text = "AdminPassword"
    $Label_AdminPassword.AutoSize = $true
    $Label_AdminPassword.Width = 25
    $Label_AdminPassword.Height = 10
    $Label_AdminPassword.location = new-object system.drawing.point(10,410)
    $Label_AdminPassword.Font = "Microsoft Sans Serif,10"
    $Form_Manual.controls.Add($Label_AdminPassword)

    $TextBox_AdminPassword = New-Object System.Windows.Forms.MaskedTextBox
    $TextBox_AdminPassword.PasswordChar = '*'
    $TextBox_AdminPassword.Width = 180
    $TextBox_AdminPassword.Height = 20
    $TextBox_AdminPassword.location = new-object system.drawing.point(120,410)
    $TextBox_AdminPassword.Font = "Microsoft Sans Serif,10"
    $Form_Manual.controls.Add($TextBox_AdminPassword)

    $Label_ProductKey = New-Object system.windows.Forms.Label
    $Label_ProductKey.Text = "ProductKey"
    $Label_ProductKey.AutoSize = $true
    $Label_ProductKey.Width = 25
    $Label_ProductKey.Height = 10
    $Label_ProductKey.location = new-object system.drawing.point(10,440)
    $Label_ProductKey.Font = "Microsoft Sans Serif,10"
    $Form_Manual.controls.Add($Label_ProductKey)

    $TextBox_ProductKey = New-Object System.Windows.Forms.TextBox
    $TextBox_ProductKey.Width = 180
    $TextBox_ProductKey.Height = 20
    $TextBox_ProductKey.location = New-Object system.drawing.point(120,440)
    $TextBox_ProductKey.Font = "Microsoft Sans Serif,10"
    $Form_Manual.controls.Add($TextBox_ProductKey)

    $Button_Deploy = New-Object System.Windows.Forms.Button
    $Button_Deploy.Text = "Deploy"
    $Button_Deploy.Width = 60
    $Button_Deploy.Height = 30
    $Button_Deploy.location = New-Object system.drawing.point(150,483)
    $Button_Deploy.Font = "Microsoft Sans Serif,10"
    $Button_Deploy.Add_Click({

        $Button_Deploy.Enabled = $false

        $Manual_Params = @{
    
            VMName = $TextBox_VMName.Text
            Template = $ComboBox_Template.Text
            Datastore = $ComboBox_Datastore.Text
            Folder = $ComboBox_Folder.Text
            OSType = $ComboBox_OSType.Text
            Cluster = $ComboBox_Cluster.Text
            NetworkName = $ComboBox_NetworkName.Text
            IPAddress = $TextBox_IPAddress.Text
            DefaultGateway = $TextBox_DefaultGateway.Text
            SubnetMask = $TextBox_SubnetMask.Text
            DNSServer = $TextBox_DNSServer.Text
            VMNote = $TextBox_VMNote.Text
            Domain = $TextBox_Domain.Text
            AdminPassword = $TextBox_AdminPassword.Text
            ProductKey = $TextBox_ProductKey.Text
            DSMaxLiveVM = $NumericUpDown_DSMaxLiveVM.Text
            DsUsagePercentThreshold = $NumericUpDown_DsUsagePercentThreshold.Text
            MaxVCpuPerCore = $NumericUpDown_MaxVCpuPerCore.Text
            MaxVMHostRamProv = $NumericUpDown_MaxVMHostRamProv.Text
        }

        if ($ListBox_ExcludedHosts.SelectedItem) {$Manual_Params += @{ExcludedVMHost = $ListBox_ExcludedHosts.SelectedItems.Name}}

        if (($ComboBox_OSType.Text -eq 'Windows') -and ($TextBox_Domain.Text)) {
        
            $Manual_Params += @{DomainCredential = (Get-Credential -Message "Enter Domain Credential")}
        }
        
        #$Manual_Params | ogv
        Deploy-VM @Manual_Params
        #$Manual_Params | out-file C:\Users\jamesg\Desktop\test.txt

        $Button_Deploy.Enabled = $true
    })
    $Form_Manual.controls.Add($Button_Deploy)

    [void]$Form_Manual.ShowDialog()
    $Form_Manual.Dispose()
}

######

function Deploy-VM {
    
    <#
    .NOTES
    
    .SYNOPSIS
        Deploy VMs using custom specs.

    .DESCRIPTION

    .PARAMETER VMName
        Specifies a name for the new virtual machine as well as the OS hostname.

    .PARAMETER Template
        Specifies the virtual machine template you want to use for the creation of the new virtual machine

    .PARAMETER Datastore
        Specifies the datastore where you want to place the new virtual machine. Matches against part of a name e.g. use 'SAAS' to have the script select the most appropriate SAAS datastore.

    .PARAMETER Folder
        Specifies the folder where you want to place the new virtual machine.

    .PARAMETER OSType
        Specifies the type of the operating system. The valid values are Linux and Windows.

    .PARAMETER Cluster
        Specifies the names of the cluster you want to use.

    .PARAMETER NetworkName
        Specifies the name of the network to which you want to connect the virtual network adapter.

    .PARAMETER IPAddress
        Specifies an IP address.

    .PARAMETER DefaultGateway
        Specifies a default gateway.

    .PARAMETER SubnetMask
        Specifies a subnet mask.

    .PARAMETER DNSServer
        Specifies the DNS server settings. When using a CSV seperate DNS servers using a space.

    .PARAMETER VMNote
        Provide a description for the virtual machine.

    .PARAMETER Domain
        Specifies a domain name. Mandatory when using OSType Linux.

    .PARAMETER AdminPassword
        Specifies a new OS administrator's password. This parameter applies only to Windows operating systems.

    .PARAMETER ProductKey
        Specifies the MS product key.

    .PARAMETER DomainCredential
        Specifies the credentials you want to use for domain authentication. This parameter applies only to Windows operating system.

    .PARAMETER ExcludedVMHost
        Specify VMhosts to be excluded. These hosts will not be used when deploying VMs.

    .PARAMETER OrgName
        Specifies the name of the organization to which the administrator belongs.

    .PARAMETER TimeZone
        Specifies the name or ID of the time zone for a Windows guest OS only.

    .PARAMETER DSMaxLiveVM
        LEAVE AS DEFAULT! Max live VMs per datastore.

    .PARAMETER DsUsagePercentThreshold
        LEAVE AS DEFAULT! Datastore max usage percent.

    .PARAMETER MaxVCpuPerCore
        LEAVE AS DEFAULT! Max virtual CPUs per physical core.

    .PARAMETER MaxVMHostRamProv
        LEAVE AS DEFAULT! Max provisioned RAM percent per host.

    .PARAMETER LogFilePath
    
    .EXAMPLE
    
        Deploy-VM -VMName test-win -Template 2012r2-STD-GUI -Datastore SAAS -Folder "NO BACKUP" -OSType Windows -NetworkName "VM Customer 52" -IPAddress 172.31.11.68 -DefaultGateway 172.31.11.65 -SubnetMask 255.255.255.224 -DNSServer "172.31.6.136","172.31.6.137" -Domain domain.com -DomainCredential domain\user

    .EXAMPLE

        Deploy-VM -VMName test-lin2 -Template ubuntu-16.04.2-lvm -Datastore 15KRAID6 -Folder nonprod -OSType linux -NetworkName "VLAN 1024" -IPAddress 10.39.24.30 -DefaultGateway 10.39.24.3 -SubnetMask 255.255.255.0 -DNSServer "10.39.2.10","10.39.2.11" -Domain domain.com

    .EXAMPLE
        
        Import-Csv -Path '.\Deploy-Module.csv' | Deploy-VM -DomainCredential domain\user
    #>
    param(

        [parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateScript({Validate-AllowedSpecialChar $_})]
        [ValidateLength(3,15)]
        [string]$VMName,

        [parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Template,

        [parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Datastore,

        [parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Folder,

        [parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateSet('Windows','Linux')]
        [string]$OSType,

        ##

        [parameter(ValueFromPipelineByPropertyName=$true)]
        [string]$Cluster,

        [parameter(ValueFromPipelineByPropertyName=$true)]
        [string]$NetworkName,

        [parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateScript({($_ -eq $null) -or ($_ -eq "") -or ($_ -match [IPAddress]$_)})]
        [string]$IPAddress,

        [parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateScript({($_ -eq $null) -or ($_ -eq "") -or ($_ -match [IPAddress]$_)})]
        [string]$DefaultGateway,

        [parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateScript({($_ -eq $null) -or ($_ -eq "") -or ($_ -match [IPAddress]$_)})]
        [string]$SubnetMask,

        [parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateScript({($_ -eq $null) -or ($_ -eq "") -or ($_ -split " " | ForEach-Object {$_ -match [IPAddress]$_})})]
        [string[]]$DNSServer,

        [parameter(ValueFromPipelineByPropertyName=$true)]
        [string]$VMNote,

        [parameter(ValueFromPipelineByPropertyName=$true)]
        [string]$Domain,

        [parameter(ValueFromPipelineByPropertyName=$true)]
        [string]$AdminPassword,

        [parameter(ValueFromPipelineByPropertyName=$true)]
        [string]$ProductKey,

        ##
        
        [PSCredential]
        [System.Management.Automation.CredentialAttribute()]$DomainCredential,

        [ValidateNotNullOrEmpty()]
        [string[]]$ExcludedVMHost,

        [validateNotNullOrEmpty()]
        [string]$OrgName = 'Zonal Retail Data Systems',

        [validateNotNullOrEmpty()]
        [int]$TimeZone = 085,

        [validateNotNullOrEmpty()]
        [int]$DSMaxLiveVM = 20,
        
        [validateRange(10,100)]
        [int]$DsUsagePercentThreshold = 80,
        
        [validateNotNullOrEmpty()]
        [int]$MaxVCpuPerCore = 5,
        
        [validateRange(10,100)]
        [int]$MaxVMHostRamProv = 80,

        [validateNotNullOrEmpty()]
        [string]$LogFilePath = "$($MyInvocation.mycommand -replace '.ps1','').log"
    )

    begin {

        $ErrorActionPreference = 'Stop'

        if (!$DefaultVIServer) {throw "No DefaultVIServer. Connect to vCenter using 'Connect-VIServer' and try again."}

        Start-Transcript $LogFilePath

        if ($ExcludedVMHost) {$ExcludedVMHost = Get-VMHost $ExcludedVMHost}

    }

    Process {

        try {

            ## Check Parameters

            if (($OSType -eq 'Windows') -and $Domain -and (!$DomainCredential)) {throw "Domain credentials not provided."}

            if ($OSType -eq 'Linux' -and (!$Domain)) {throw "The Domain parameter is mandatory when using the OSType 'Linux'."}

            if ($IPAddress -and (Test-Connection $IPAddress -Quiet -Count 2)) {throw "IPAddress '$IPAddress' is already in use."}

            if (($OSType -eq 'Windows') -and (!$AdminPassword)) {
            
                $AdminPassword = New-RandomPassword -Special:$false
                $VMNote = $VMNote + " - Admin Pass - $AdminPassword"
            }
            
            ## Splat parameters for OSCustomizationSpec

            $Guid = [guid]::NewGuid()

            $OSCustomizationSpec_Params = @{
                
                Name = $Guid
                Description = "$VMName"
                OSType = $OSType
                Type = 'NonPersistent'
            }
         
            if ($Domain) {$OSCustomizationSpec_Params += @{Domain = $Domain}}

            if ($OSType -eq 'Windows') {$OSCustomizationSpec_Params += @{ChangeSid = $true ; AdminPassword = $AdminPassword ; TimeZone = $TimeZone ; FullName = $OrgName ; OrgName = $OrgName}}

            if ($DomainCredential -and $Domain -and $OSType -eq 'Windows') {$OSCustomizationSpec_Params += @{DomainCredential = $DomainCredential}}

            if ($OSType -eq 'Windows' -and (!($Domain -and $DomainCredential))) {$OSCustomizationSpec_Params += @{Workgroup = 'WORKGROUP'}}

            if ($OSType -eq 'Windows' -and $ProductKey) {$OSCustomizationSpec_Params += @{ProductKey = $ProductKey}}

            if ($DNSServer) {$DNSServer = $DNSServer -split ' '}

            if ($DNSServer -and $OSType -eq 'Linux') {$OSCustomizationSpec_Params += @{DnsServer = $DNSServer}}

            ## Splat parameters for OSCustomizationNicMapping

            if ($IPAddress -and $SubnetMask -and $DefaultGateway) {

                $OSCustomizationNicMapping_Params = @{
                    
                    IpMode = 'UseStaticIP'
                    IpAddress = $IPAddress
                    SubnetMask = $SubnetMask
                    DefaultGateway = $DefaultGateway 
                }
            }
            
            else {$OSCustomizationNicMapping_Params = @{IpMode = 'UseDhcp'}}

            if ($DNSServer -and $OSType -eq 'Windows') {$OSCustomizationNicMapping_Params += @{Dns = $DNSServer}}

            $CustomSpec = New-OSCustomizationSpec @OSCustomizationSpec_Params
            
            $CustomSpec | Get-OSCustomizationNicMapping | Set-OSCustomizationNicMapping @OSCustomizationNicMapping_Params -Confirm:$false | Out-Null
            
            ## Gather parameters for New-VM

            $ChosenTemplate = Get-Template $Template

            $Folder = Get-Folder $Folder

            if ($Cluster) {$VMHosts = Get-Cluster $Cluster | Get-VMHost | Where-Object {$_.Name -notin $ExcludedVMHost}}

            else {$VMHosts = Get-VMHost | Where-Object {$_.Name -notin $ExcludedVMHost -and $_.ConnectionState -eq 'Connected'}}
            
            $VMHostList = @()

            ## Choose VMHost

            foreach ($VMHost in $VMHosts) {

                $LiveVM = $VMHost | Get-VM | Where-Object {$_.PowerState -eq 'PoweredOn'}
                $VCpuPerCore = (($LiveVM | Measure-Object -Sum -Property NumCpu).Sum + $Templatenumcpu) / $VMHost.NumCpu
                $VRamProvPercent = (($LiveVM | Measure-Object -Sum -Property MemoryMB).Sum + $Templatememorymb) / $VMHost.MemoryTotalMB

                IF (($VCpuPerCore -le $MaxVCpuPerCore) -and ($VRamProvPercent -le ($MaxVMHostRamProv/100))) {
                    $VMHostList += $VMHost
                }
            }

            if (!$VMHostList) {throw "No host to satisfy the requirements"}
        
            $ChosenVMHost = $VMHostList | Sort-Object CpuUsageMhz | select -First 1

            if ($NetworkName) {
                
                $VLAN = Get-VMHost $ChosenVMHost | Get-VirtualPortGroup | Where-Object {$_.Name -eq $NetworkName}
                if (!$VLAN) {throw "VlanID $NetworkName is unavailable on VMHost $ChosenVMHost."}
            }

            ## Choose Datastore

            $DatastoreList = Get-Datastore | Where-Object {$_.Name -match $Datastore}

            $TemplateDiskGB = ($ChosenTemplate.ExtensionData.Summary.Storage.committed + $ChosenTemplate.ExtensionData.Summary.Storage.uncommitted)/1GB

            $ChosenDS = $DatastoreList | Where-Object {

                ($_.FreeSpaceGB - $TemplateDiskGB)/$_.CapacityGB -ge (1-($DsUsagePercentThreshold/100)) -and `
                ((($_ | Get-VM | where PowerState -eq PoweredOn) | Measure-Object).count + 1) -le $DSMaxLiveVM
            
            } | Sort-Object FreeSpaceGB -Descending | select -First 1

            IF (!$ChosenDS) {throw "No selected datastore to satisfy the requirements"}

            ## Splat parameters for New-VM

            $VM_Params = @{

                Name = $VMName
                ResourcePool = $ChosenVMHost
                Template = $ChosenTemplate
                Location = $Folder
                Datastore = $ChosenDS
                OSCustomizationSpec = $CustomSpec
            }

            if ($VMNote) {$VM_Params += @{Notes = $VMNote}}

            New-VM @VM_Params | Out-Null

            if ($VLAN) {Get-VM $VMName | Get-NetworkAdapter | Set-NetworkAdapter -NetworkName $VLAN.Name -StartConnected $true -Confirm:$false | Out-Null}

            Get-VM $VMName | Start-VM
        }

        catch {
        
            Write-Error $_.Exception.Message -ErrorAction Continue
        }

        Finally {
            
            try {Get-OSCustomizationSpec -Name $Guid | Remove-OSCustomizationSpec -Confirm:$false} catch {}
        }
    }

    end {Stop-Transcript}
}

function New-DeployCSV {
    
    param([parameter(position=0)][string]$Path = (Get-Location).Path + "\Deploy-VM.csv")

    $CSVContent = @"
VMName,Template,Datastore,Folder,OSType,Cluster,NetworkName,IPAddress,DefaultGateway,SubnetMask,DNSServer,VMNote,Domain,AdminPassword,ProductKey
test-win,2012_R2_Core_SaaS,15KRAID6,NonProd,Windows,,TF Utility 1039,10.39.2.166,10.39.2.3,255.255.254.0,10.39.2.10 10.39.2.11,test note,,,
test-lin,Ubuntu_16.04,Capacity,NonProd,Linux,,TF Utility 1039,10.39.2.167,10.39.2.3,255.255.254.0,10.39.2.10 10.39.2.11,,zonalconnect.local,,
test-domain,2012_R2_Core_SaaS,15KRAID6,NonProd,Windows,,TF Utility 1039,10.39.2.168,10.39.2.3,255.255.254.0,10.39.2.10 10.39.2.11,,zonalconnect.local,Badger777,XXXXX-XXXXX-XXXXX-XXXXX-XXXXX
"@

    $CSVContent | Out-File -FilePath $Path
}

Function Validate-CSVContent {

<#

.DESCRIPTION

Returns nothing if OK

Returns the list of what is wrong if not OK

#>

[cmdletbinding()]

param(
    $CSVPath,
    $RequiredFields = @("VMName","Template","Datastore","Folder","OSType","Cluster","NetworkName","IPAddress","DefaultGateway","SubnetMask","DNSServer","VMNote","Domain","AdminPassword","ProductKey")
)

$CSVImport = Import-Csv $CSVPath

if ( !(Compare-Object -ReferenceObject ($CSVImport | Get-Member -MemberType NoteProperty).Name -DifferenceObject $RequiredFields) ) {

    ## Verification that all VIObjects exist

    foreach ($temp in $CSVImport.template) {

        if (!(Get-template $temp)) {"Template $temp doesn't exist`r`n"}

    }

    foreach ($ds in $CSVImport.datastore) {

        if (!(Get-datastore "*$ds*")) {"Datastore matching `"$ds`" doesn't exist`r`n"}

    }

    foreach ($fol in $CSVImport.folder) {

        if (!(Get-folder $fol)) {"Folder $fol doesn't exist`r`n"}

    }

    foreach ($pg in $CSVImport.NetworkName) {

        if (!(Get-virtualportgroup -name $pg)) {"Port group $pg doesn't exist`r`n"}

    }

    foreach ($clus in $CSVImport.cluster) {

        if (!(Get-cluster $clus)) {"Cluster $clus doesn't exist`r`n"}

    }

    ## Verification that all IPs are correct
    
    foreach ($ip in @($CSVImport.ipaddress,$CSVImport.defaultgateway,$CSVImport.subnetmask,($CSVImport.DNSServer -split " "))) {
        $IP | ForEach-Object {
            if (($_ -ne '') -and (!($_ -as [ipaddress]))) {"Invalid ip : $_`r`n"}
        }
    }

    ## Verification that product keys are correct

    foreach ($pkey in $CSVImport.ProductKey) {if ($pkey -notlike "?????-?????-?????-?????-?????" -and $pkey) {"Invalid product key : $pkey`r`n"}}

    ## Verification that the VM names don't contain unallowed special characters

    foreach ($VMName in $CSVImport.VMName) {if (!(Validate-AllowedSpecialChar -allowedSpecialChar "-",".","_" -String $VMName)) {"`"$VMName`" contains prohibited characters`r`n"}}

} else {"Invalid CSV Headers!`n"}

}

Function Validate-Resources {

<#

.DESCRIPTION

Returns $true if validation OK.
Returns the resource limitations as an error if validation not OK.
No use of pipeline

#>

param(
    [parameter(Mandatory=$True)]
    [ValidateNotNullOrEmpty()]
    $Collection,

    [ValidateNotNullOrEmpty()]
    [int]
    $DSMaxLiveVM = 20,

    [ValidateNotNullOrEmpty()]
    [int]
    $DsUsagePercentThreshold = 80,

    [ValidateNotNullOrEmpty()]
    [int]
    $MaxVCpuPerCore = 5,

    [ValidateNotNullOrEmpty()]
    [int]
    $MaxVMHostRamProv = 80,

    $VMHostExcluded
)

TRY {

    # Check for template and datastore properties in the objects collection

    $Cmember = $Collection | Get-Member | select -ExpandProperty name

    IF ($Cmember -notcontains "datastore" -or $Cmember -notcontains "template") {Write-Output "The collection of objects must contain the fields datastore and template to run the validation`r`n"}
 
    IF ($ResourceNONOK) {Clear-Variable ResourceNONOK}
    $tempMemMB  = 0
    $tempNumCPU = 0
   
    ################################ BUILD DATASTORE ARRAY OBJECT

    $DsArray = @()
    $Datastores = @()

    $Collection.Datastore | Select -Unique | ForEach-Object {
        $CurrentDS = Get-datastore "*$_*" -ErrorAction Stop
        $Datastores += $CurrentDS
        $DsArray += [pscustomobject]@{
            Datastore = $_
            DatastoreVIObject= $CurrentDS
            CapacityGB       = ($CurrentDS | Measure-Object -Property capacitygb -Sum).sum
            ExpectedDSFreeGB = ($CurrentDS | Measure-Object -Property FreeSpaceGB -Sum).sum
            ExpectedLiveVMDS = ($CurrentDS | Get-VM -ErrorAction Stop | where powerstate -eq poweredon | Measure-Object).count
        }
    }

    ################################ TEMPLATE EXPECTED RESOURCES

    foreach ($CsvRecord in $Collection) {
        # Template
        $CurTemp = Get-template $CsvRecord.Template -ErrorAction Stop

        # Template disk
        $curTempDiskGB = ($CurTemp.ExtensionData.Summary.Storage.committed + $CurTemp.ExtensionData.Summary.Storage.uncommitted )/1gb
        ($DsArray | where Datastore -match $CsvRecord.datastore).ExpectedDSFreeGB = ($DsArray | where Datastore -match $CsvRecord.datastore).ExpectedDSFreeGB - $curTempDiskGB
        ($DsArray | where Datastore -match $CsvRecord.datastore).ExpectedLiveVMDS++
         
        # Template compute
        $tempMemMB  += $CurTemp.ExtensionData.Summary.Config.MemorySizeMB | Out-Null
        $tempNumCPU += $CurTemp.ExtensionData.Summary.Config.numcpu | Out-Null
    }
    
    ################################ COMPARE EXPECTED RESOURCES WITH THRESHOLDS
    
    $DeployVMHost = Get-VMHost -ErrorAction Stop | where {$_ -notin $VMHostExcluded}

    $NumberVM = ($Collection | Measure-Object).count

    $LiveVm = $DeployVMHost | get-vm | where powerstate -eq poweredon

    # Compute resources check

    $ExpectedMemoryPercent = (($LiveVm | Measure-Object -Property memorymb -Sum).sum + $tempMemMB) / ($DeployVMHost | Measure-Object -Property memorytotalmb -Sum).sum
    $ExpectedCpuperCore    = (($livevm | Measure-Object -Sum -Property numcpu).sum + $tempNumCPU) / ($DeployVMHost | Measure-Object -Property numcpu -Sum).sum
    
    IF ($ExpectedMemoryPercent -gt ($MaxVMHostRamProv/100)) {$ResourceNONOK += "Memory: $([math]::round($ExpectedMemoryPercent*100,2))% > $([math]::round($MaxVMHostRamProv,2))%`r`n"}
    IF ($ExpectedCpuperCore -gt $MaxVCpuPerCore) {$ResourceNONOK += "CPU cores: $([math]::round($ExpectedCpuperCore,2)) to 1 > $MaxVCpuPerCore to 1`r`n"}   
    
    # Datastore resource check

    foreach ($DSgroup in $DsArray) {

        # Get expected datastore % free space after the deployment from this DSarray object
        $ExpectedDSFreePercent = $DSgroup.ExpectedDSFreeGB / $DSgroup.CapacityGB

        IF ($ExpectedDSFreePercent -lt (1-($DsUsagePercentThreshold/100))) {
            $ResourceNONOK += "[$($DSgroup.Datastore)] Storage free space : $([math]::round($ExpectedDSFreePercent * 100,2))% < $(100 - $DsUsagePercentThreshold)%`r`n"
        }

        # Get expected average number of live VM per datastore after the deployment from this DSarray object
        $ExpectedLiveVMDS = $DSgroup.ExpectedLiveVMDS / (($DSgroup.DatastoreVIObject | Measure-Object).count)

        IF ($ExpectedLiveVMDS -gt $DSMaxLiveVM) {$ResourceNONOK += "[$($DSgroup.Datastore)] VM per Datastore: $([math]::round($ExpectedLiveVMDS,2)) > $DSMaxLiveVM`r`n"}

    }

    ################################ FUNCTION RETURN

    IF ($ResourceNONOK) {

       Write-Output ($ResourceNONOK | select -Unique)

    } ELSE {

       return $true

    }

} CATCH {

    Write-Output $_.Exception.Message

}

}

Function New-RandomPassword {

    param(
        [ValidateRange(6,30)]
        [int]$Length = 12,
        [switch]$UpperCase = $true,
        [switch]$LowerCase = $true,
        [switch]$Numbers = $true,
        [switch]$Special = $true,
        [string[]]$ExcludedChars = @("$","[","]","^","{","}",'`','~'," ","|","#","I","l","O","0","'","`"")
    )

    IF (!$UpperCase -and !$LowerCase -and !$Numbers -and !$Special) {write-warning "Please specify characters to use";break}

    IF ($UpperCase) {$Range += 65..90}
    IF ($LowerCase) {$Range += 97..122}
    IF ($Numbers) {$Range += 48..57}
    IF ($Special) {$Range += 33..47+58..64+91..96+123..126}

    $AllowedChar = @()

    foreach ($i in $Range) {IF ([char]$i -cnotin $ExcludedChars) {$AllowedChar += [char]$i}}

    $BadPass = $true

    while ($BadPass) {
    
        if ($BadPass) {Clear-Variable BadPass}
        if ($RandomPassword) {Clear-Variable RandomPassword}

        for ($i = 1; $i –le $Length; $i++) {

            $RandomIndex = Get-Random -Maximum $AllowedChar.Count

            $RandomPassword += $AllowedChar[$RandomIndex]
        }

        IF ($LowerCase) {IF ($RandomPassword -cnotmatch "[a-z]") {$BadPass = $true}}
        IF ($UpperCase) {IF ($RandomPassword -cnotmatch "[A-Z]") {$BadPass = $true}}
        IF ($Numbers)   {IF ($RandomPassword -notmatch "[0-9]") {$BadPass = $true}}
        IF ($Special)   {IF ($RandomPassword -cnotmatch '[^a-zA-Z0-9]') {$BadPass = $true}}
    }

    $RandomPassword
}

Function Validate-AllowedSpecialChar {

    param(
        [ValidateNotNullOrEmpty()]
        [string]$String,
        [char[]]$AllowedSpecialChar = @("-","_",".")
    )

    $CharArray = $String.ToCharArray() | where {$_ -match '[^a-zA-Z0-9]'}

    foreach ($Char in $CharArray) {If ($Char -notin $AllowedSpecialChar) {$NOK++}}

    if (!$NOK) {$true}
    else {$false}
}

Function Get-VMCustomizationStatus {

param(
    [VMware.VimAutomation.ViCore.types.V1.Inventory.VirtualMachine]
    $VM
)

    $custoStarted   = "CustomizationStartedEvent"
    $custoSucceeded = "CustomizationSucceeded"
    $custoFailed    = "CustomizationFailed"
    $CustoTask      = "Task: Customize virtual machine guest OS"

    $events     = Get-VIEvent -Entity $VM 

    IF ($events | where {$_.GetType().name -eq $custoSucceeded}) {Return "Success"}
    IF ($events | where {$_.GetType().name -eq $custoFailed})    {Return "Failed"}
    IF ($events | where {$_.GetType().name -eq $custoStarted})   {Return "Started"}

    IF ($events | where fullformattedmessage -eq $CustoTask) {Return "NotStarted"}
    ELSE {Return "Unknown"}

}

function Install-Zabbix {

    param(
        
        [parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$VMName,
        
        [PSCredential][System.Management.Automation.CredentialAttribute()]$Credential,

        [ValidateNotNullOrEmpty()]
        [string]$ZabbixFolderPath = 'Z:\zts\tools\VM Setup\Zabbix\Current'
    )

    begin {
        
        $ErrorActionPreference = 'Stop'

        if (!$DefaultVIServer) {Write-Error "No DefaultVIServer. Connect to vCenter using 'Connect-VIServer' and try again."}

        if (!(Test-Path $ZabbixFolderPath)) {Write-Error "Cannot find ZabbixFolder - '$ZabbixFolderPath'."}
    }

    process {
        
        try {
            
            $Source = "$ZabbixFolderPath\*"
            $Destination = "C:\Temp\zabbix\"
            $VM = Get-VM $VMName

            $Copy_Params = @{VM = $VM;Source = $Source;Destination = $Destination;LocalToGuest = $true;Force = $true}
            
            $VMScript_Params = @{VM = $VM;ScriptType = 'Bat'}

            if ($Credential) {$Copy_Params += @{GuestCredential = $Credential} ; $VMScript_Params += @{GuestCredential = $Credential}}

            Copy-VMGuestFile @Copy_Params -WarningAction SilentlyContinue

            $ScriptText = "cd $Destination & PowerShell -NoProfile -ExecutionPolicy Bypass -File zabbix_agent-install.ps1 & cd.. & rmdir $Destination /S /Q"

            $result = Invoke-VMScript @VMScript_Params -ScriptText $ScriptText -WarningAction SilentlyContinue

            if (!($result | Select-String "Status      : Running" -SimpleMatch)) {Write-Error "Zabbix Agent failed to install."}
        }

        catch {Write-Error $_.Exception.Message}
    }
}

Function Check-InstalledSoft {

Param(

    [ValidateNotNullOrEmpty()]
    [String[]]$SoftToCheck,

    [ValidateScript({Test-Connection -Quiet $_})]
    [String[]]$Computer,

    [System.Management.Automation.PSCredential]
    $credential
)

    $params = @{ScriptBlock = {Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*}}

    IF ($Computer) { $params += @{ComputerName = $Computer} }

    IF ($Credential) { $params += @{Credential = $Credential} }



    $CheckInstalled = Invoke-Command @params



    $SoftToCheck | ForEach-Object {
        $soft = $_
        
        $CheckInstalled | ForEach-Object {
            
            IF ($_.DisplayName -like "*$soft*") {
                
                [PSCustomObject]@{
                    Computer = $_.PSComputerName
                    Software = $_.DisplayName
                    Version  = $_.DisplayVersion
                }
            } # IF
        } #$CheckInstalled | ForEach-Object 
    } #$SoftToCheck | ForEach-Object 
}

Function Install-Office {

<#
.SYNOPSIS
Function : Install-Office
Version  : 1.5
EMail    : Xavier.avrillier@zonal.co.uk
Date     : 15/01/2016
Info     : Remotely install Microsoft Office 2013 (Word, Excel) on remote machine.
          That script uses administrative shares to copy the MSP file on the box so it needs to be on the domain

By default installs Office 2013 with the key stated in the description.
99% of the time there is no need to specify any parameter (except the VM one).
Added parameters in MSP file to skip the menus (First things first) at first launch of office.

.DESCRIPTION
---------------------------------------------------------------------------------------
NOTES FOR LATER UPDATE
---------------------------------------------------------------------------------------
To create an MSP file : 

    - Mount the Office ISO or get the file in a folder
    - At the root the ISO, run "Setup.exe /admin" to launch the Office Customization Tool (OCT)
    - Configure all the properties as expected with all the silent properties checked
    - In "Modify Setup Properties" add the followings :
        + AUTO_ACTIVATE = 1
        + REBOOT        = ReallySuppress
        + SETUP_REBOOT  = Never


Office Standard 2013 MAK Key : N7F4B-GDYM8-WW8XC-36RKR-7H37T

.PARAMETER MSP
path to the MSP file (all silenting option must have been enabled chen creating it)

.PARAMETER OfficeISO
Full path to the Microsoft Office ISO file  

.PARAMETER VM
Virtual Machin object

.EXAMPLE 
PS C:\> Get-VM "Zonal-HR-*" | Install-Office

This command will install office on the HR VMs using the default values specified in the parameters.

.EXAMPLE
PS C:\> Get-VM "Zonal-HR-*" | Install-Office -OfficeISO (Get-item 'vmstore:\*\RDG1\*Office_2010.ISO') -MSP '\\dca-utl-nas\main\zts\resources\Office2010-AnswerFile.MSP'

This command will install office on the HR VMs using the Office 2010 ISO and MSP file specified in the input.
#>

    param (
        [parameter(position=0,ValueFromPipeline=$True,Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({Test-Connection -Quiet $_.Guest.IPAddress})]
        [VMware.VimAutomation.ViCore.types.V1.Inventory.VirtualMachine[]]
        $VM,

        [System.Management.Automation.PSCredential]
        $Credential,

        [parameter(position=1)]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({(get-item $_).ItemType -eq "IsoImageFile"})]
        [String]
        $OfficeISO = "vmstores:\*@443\DCA\RDG1\ISOs\SW_DVD5_Office_2013w_SP1_64Bit_English_MLF_X19-34904.ISO",

        [parameter(position=2)]
        [ValidateNotNullOrEmpty()]
        [ValidatePattern('.msp$')]
        [ValidateScript({Test-Path $_})]
        [String]
        $MSP = 'Z:\zts\Scripts\WindowsPowerShell\Modules\zDeploy\Office2013STDAnswerFile.MSP'
    )

    begin {
      $MSPLocalPath = "C:\Temp\" + ($MSP.Split("\") | select -Last 1)
      $OfficeISOPath    = (Get-Item $OfficeISO).DatastoreFullPath
    }

    process {

        $VM | ForEach-Object {
            
            $params = @{ComputerName = $_.Guest.Hostname}
            IF ($Credential) { $params += @{Credential = $Credential} }

            $RemoteSession = New-PSSession @params

            $Computer = $_.Guest.Hostname
            $CD       = Get-CDDrive $_
            $Proceed  = "n"
            Write-Output "---------------------------------------"
            Write-Output "    Processing $($VM.Name)"
            Write-Output "---------------------------------------"

            TRY {
            
            #Error checks --------------------------------------------------------------------------------------------------
                Write-Verbose "Checking for compliancy"

                IF (-not$CD) { Throw "No CD drive on $($_.name)"}

                IF ($CD.IsoPath) { 

                    $Proceed = Read-Host "$($CD.IsoPath) currently mounted, proceed anyway ? [Y/N]"

                    IF ($Proceed -ne "Y") {Throw "$($CD.IsoPath) mounted : Installation on $($VM.Name) aborted by user"}
                }
                
                $params = @{Computer = $Computer ; SoftToCheck = "Microsoft Office Standard"}

                IF ($Credential) { $params += @{Credential = $Credential} }

                IF (Check-InstalledSoft @params) {Throw "Microsoft Office Standard is already installed on $($_.name)"}

                IF ( Invoke-Command -Session ($RemoteSession) -ScriptBlock {get-process} | where processname -eq "Setup"  ) {Throw "An install is currently being performed on $($_.name)"}

            #Prep VM -------------------------------------------------------------------------------------------------------
                Write-Output "Mounting ISO and transfering Office answer file"

                $CD | Set-CDDRive -IsoPath $OfficeISOPath -Connected $true -Confirm:$False | Out-Null
            
                Start-Sleep 5
                
                $params = @{path = "\\$Computer\C$\Temp"}
                IF ($Credential) { $params += @{Credential = $Credential} }

                if (-not(Test-Path @params)) {mkdir @params | Out-Null}

                $params = @{path = $MSP ; destination = "\\$Computer\C$\Temp" ; force = $true} 
                IF ($Credential) { $params += @{Credential = $Credential} }

                Copy-Item @params


            #Invoke script on remote computer ------------------------------------------------------------------------------
                Write-Output "Invoke Office Installation remotely"

                $params = @{ComputerName=$Computer ; ArgumentList=$MSPLocalPath ; ScriptBlock = {
                    
                    $CDletter = Get-WMIObject -Class Win32_CDROMDrive | select -ExpandProperty drive
                    
                    #Check if ISO mounted is Office. Office CD has a specific readme.txt file in \Updates\
                    $CheckOfficeA = Test-Path "$($CDletter)\Updates\readme.txt"
                    IF ($CheckOfficeA) {$CheckOfficeB = (Get-content "$($CDletter)\Updates\readme.txt") -notlike "*Any patches placed in this folder will be applied during initial install*"}

                    IF (-not$CheckOfficeA -or $CheckOfficeB) {throw "ISO file mounted is not Microsoft Office"}

                    #ISO mounted is Office. Installation can begin
                    Start-Process -FilePath "$CDletter\setup.exe" -ArgumentList "/adminfile $($args[0])" -Wait

                 } }

                IF ($Credential) { $params += @{Credential = $Credential} }

                Invoke-Command @params

            #Clean remote computer and check for installation --------------------------------------------------------------
                Write-Output "Cleaning $($VM.Name)"

                $CD | Set-CDDrive -NoMedia -Confirm:$False | Out-Null -ErrorAction SilentlyContinue

                $params = @{path = ("\\$Computer\$MSPLocalPath").Replace(':','$')}
                IF ($Credential) { $params += @{Credential = $Credential} }

                Remove-Item @params                

                $params = @{Computer = $Computer ; SoftToCheck = "Microsoft Office Standard"}
                IF ($Credential) { $params += @{Credential = $Credential} }

                IF (Check-InstalledSoft @params) {

                    Write-Host "Microsoft Office Standard installed on $($_.name)" -ForegroundColor Green
                    Write-Host "/!\ SPLA spreadsheet must be updated /!\" -ForegroundColor Green

                } ELSE {Throw "Microsoft Office installation failed on $($_.name)"}

            } 
                               
            CATCH {
                 Write-error $_.Exception.message
            }
        }
    } 
} 

#####

function Create-TescoADObjects {

    param(
        
        [parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateScript({Validate-AllowedSpecialChar $_})]
        [ValidateLength(3,15)]
        [string]$VMName,
        [string]$OUPath = "OU=Tesco,OU=zCustomers,DC=zonalconnect,DC=local"
    )

    process {

        $ComputerName = $VMName
        $UserName = "$VMName-user"
        $GroupName = "$VMName-users"
        $OUName = $VMName
        $PolicyName = "$VMName-policy"
        $Description = $VMName



        $OU = NEW-ADOrganizationalUnit -Name $OUName –path $OUPath -Description $Description -PassThru

        $UserPassword = New-RandomPassword -Special:$false
        $SecUserPass = ConvertTo-SecureString -String $UserPassword -Force -AsPlainText

        $Computer = New-ADComputer -Name $ComputerName -SAMAccountName $ComputerName -Path $OU.DistinguishedName -Description $Description -PassThru

        $Group = New-ADGroup `
            -Name $GroupName `
            -SamAccountName $GroupName `
            -GroupCategory Security `
            -GroupScope Global `
            -DisplayName $GroupName `
            -Description $Description `
            -Path $OU.DistinguishedName `
            -PassThru

        $User = New-ADUser `
            -Name $UserName `
            -SamAccountName $UserName `
            -DisplayName $UserName `
            -Path $OU.DistinguishedName `
            -AccountPassword $SecUserPass `
            -PasswordNeverExpires $True `
            -CannotChangePassword $True `
            -Enabled $True `
            -Description $Description `
            -PassThru

        Add-ADGroupMember $Group.Name $User.Name

        Start-Sleep 1

        New-RdgGPO dca-utl-dc1 $PolicyName $OU.DistinguishedName $Group.SID $Description "D:\Windows\SYSVOL\sysvol\zonalconnect.local\Policies"

        [pscustomobject]@{
    
            OU = $OU.DistinguishedName
            ComputerName = $Computer.Name
            Group = $Group.Name
            UserName = "zonalconnect\$($User.Name)"
            UserPassword = $UserPassword
        }
    }
}

function New-RdgGPO ($gpoDC,$gpoName,$gpoOU,$gpoGroupSID,$gpoDescription,$gpoPathToPolicies) {


    Try {
        $gpoDetails = Copy-GPO -SourceName rdg-template -TargetName $gpoName -ErrorAction Stop
        $gpoPath = $gpoPathToPolicies + "\{" + $gpoDetails.Id + "}\Machine\Microsoft\Windows NT\SecEdit"
        

        $remoteSession = New-PSSession -computername $gpoDC
        $remoteInvoke = (Invoke-Command –Session $remoteSession –ScriptBlock { 

            $fileContent = @"
            [Unicode]
            Unicode=yes
            [Version]
            signature="`$CHICAGO$"
            Revision=1
            [Group Membership]
            *S-1-5-32-555__Memberof =
            *S-1-5-32-555__Members = *$($args[0])
            BUILTIN\Power Users__Memberof =
            BUILTIN\Power Users__Members = *$($args[0])
"@
            $fileContent | Out-File "$($args[1])\GptTmpl.inf"

        } -ArgumentList $gpoGroupSID, $gpoPath)

        $errors = (Invoke-Command –Session $remoteSession –ScriptBlock {$error})
        $remoteSession | Remove-PSSession

        if ($errors -ne $NULL) {
            foreach ($error in $errors) {
                Write-Output "`nERROR: New-RdgGPO - $($error.Exception.Message)"
                Write-Host "`nERROR: New-RdgGPO - $($error)`n" -ForegroundColor 'Red'
            }
        }

        $gpoDetails = Get-GPO $gpoName
        $gpoDetails.description = "$gpoDescription"

        $link = New-GPLink -name $gpoName -target $gpoOU -enforced yes
    }

    Catch [system.exception] {
        
        Write-Error $_.Exception.Message
    }
}


#Validate-Resources -Collection (Import-Csv -Path '.\Deploy-Module.csv')

#Import-Csv -Path '.\Deploy-Module.csv' | Create-TescoADObjects | Export-Csv -Path .\Create-TescoADObjects.csv

#Import-Csv -Path '.\Deploy-Module.csv' | Deploy-VM -DomainCredentials zonalconnect\james

#Import-Csv -Path '.\Deploy-Module.csv' | ForEach-Object {$name = $_.VMName ; Get-VMCustomizationStatus (Get-VM $Name) | select @{l='VMName';e={$Name}},@{l='Status';e={$_}}}

#Import-Csv -Path '.\Deploy-Module.csv' | Install-Zabbix

Export-ModuleMember -Function Deploy-VMGui,Deploy-VM,Get-VMCustomizationStatus,Install-Zabbix,Connect-vCenter,Install-Office,New-DeployCSV,New-RandomPassword