# !!!!!!!!!!!!!!!!
$ADTVersion = "2.10.2"
# !!!!!!!!!!!!!!!!

# Version 2.10.2
# Scott P. Morton
# 9/23/2019
# Changed from Remove-ADComputer to Remove-ADObject to avoid failures where bitlocker keys are stored in the computer object as a container
# This requires a recursive delete of the object
# More cosmetic corrections
# 9/24/2019 Set cancel buttons to disabled after process completes, cosmetic issue with form name
# 9/26/2019 Other various improvements

# Version 2.9.1
# Scott P. Morton
# 9/23/2019
# Added a throttle function to the delete process in COT to avoid excessive actions on the target DC
# Corected a cosmetic issue and a minor bug

# Version 2.8
# Scott P. Morton
# 9/20/2019
# corrected some logic around using current creds during disable/delete of CO's

# Version 2.7
# Scott P. Morton
# 9/20/2019
# Removed remnants of VMware checks for CO's
# Added a processed counter to the COT (Comp Obj Tool)
# Need to sync User object tool methods

# Version 2.6
# Scott P. Morton
# 9/18/2019
# Removed the addition PWD Age from returned data in computer object tool
# Need to sync User object tool methods

# Version 2.5
# Scott P. Morton
# 9/17/2019
# Corrected a display issue with importing CSV files for computer objects

# Version 2.4
# Scott P. Morton
# 9/17/2019
# Added a cancel button for operations under computer object tool
# Need to add one for user object tool as well

# Version 2.3
# Scott P. Morton
# 9/16/2019
# added a cancel button for aged object scanning

# Version 2.2
# Scott P. Morton
# 8/29/2019
# Added the computer object tool functionality
# enbedded all aging tools under a single top level tab
# rearranged for more estatically pleasing appearance

# Version 1.2
# Scott P. Morton
# Added User tools tab w/ User account lockout tool
# 8/28/2019 


# Version 1.1
# Written by Scott P. Morton
# 8/27/2019

Import-Module ActiveDirectory

# Global Variables
$script:date = Get-Date
$creds = $null

# User Tab global vars
$array_Usr = @()            # used for selected and displayed data
$listMatching_Usr = @{}
$failures_Usr = @{}
$filters_Usr = $false

[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
[reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.drawing

# form objects
$Form1 = New-Object System.Windows.Forms.Form 
$Tabcontrol1 =  New-Object System.Windows.Forms.TabControl
$TabcontrolAging =  New-Object System.Windows.Forms.TabControl
$objectAgingTab = New-Object System.Windows.Forms.TabPage
$userToolsTab = New-Object System.Windows.Forms.TabPage
$userObjTab = New-Object System.Windows.Forms.TabPage
$computerObjTab = New-Object System.Windows.Forms.TabPage
$Server = New-Object System.Windows.Forms.TextBox
$Server_Label = New-Object System.Windows.Forms.Label
$CredsButton = New-Object System.Windows.Forms.Button
$CurrentCreds_Check = New-Object System.Windows.Forms.CheckBox

# User Tools Tab
$Server_LabelUT = New-Object System.Windows.Forms.Label
$CredsButtonUT = New-Object System.Windows.Forms.Button
$CurrentCreds_CheckUT = New-Object System.Windows.Forms.CheckBox
$UserID_UT = New-Object System.Windows.Forms.TextBox
$UserId_Label_UT = New-Object System.Windows.Forms.Label
$LockoutButton_UT = New-Object System.Windows.Forms.Button

# User Objects Tab
$numOfDays_DrpText_Usr= New-Object System.Windows.Forms.ComboBox
$numOfDays_Label_Usr = New-Object System.Windows.Forms.Label
$LastModifiedDate_Check_Usr = New-Object System.Windows.Forms.CheckBox
$ModifiedDate_DrpText_Usr= New-Object System.Windows.Forms.ComboBox
$ModifiedDateDrpBx_Label_Usr = New-Object System.Windows.Forms.Label
$Disabled_Check_Usr = New-Object System.Windows.Forms.CheckBox
$Matches_Usr = New-Object System.Windows.Forms.Label
$Matches_Usr_Label = New-Object System.Windows.Forms.Label
$Operation_Label_Usr = New-Object System.Windows.Forms.Label
$disableObject_Usr = New-Object System.Windows.Forms.RadioButton
$deleteObject_Usr = New-Object System.Windows.Forms.RadioButton
$ScanButton_Usr = New-Object System.Windows.Forms.Button
$ModifyButton_Usr = New-Object System.Windows.Forms.Button
$ImportCSVButton_Usr = New-Object System.Windows.Forms.Button
$DisplayButton_Usr = New-Object System.Windows.Forms.Button
$ExportCSVButton_Usr = New-Object System.Windows.Forms.Button
$ResetButton_Usr = New-Object System.Windows.Forms.Button
$CancelButton_Usr = New-Object System.Windows.Forms.Button

# Computer Objects tab
$ScanButton_Comp = New-Object System.Windows.Forms.Button
$ModifyButton_Comp = New-Object System.Windows.Forms.Button
$ImportCSVButton_Comp = New-Object System.Windows.Forms.Button
$DisplayButton_Comp = New-Object System.Windows.Forms.Button
$ExportCSVButton_Comp = New-Object System.Windows.Forms.Button
$ResetButton_Comp = New-Object System.Windows.Forms.Button
$PingCheck_Comp = New-Object System.Windows.Forms.CheckBox
$Filters_Label_Comp = New-Object System.Windows.Forms.Label
$numOfDays_DrpText_Comp= New-Object System.Windows.Forms.ComboBox
$numOfDays_Label_Comp = New-Object System.Windows.Forms.Label
$LastModifiedDate_Check_Comp = New-Object System.Windows.Forms.CheckBox
$ModifiedDate_DrpText_Comp= New-Object System.Windows.Forms.ComboBox
$ModifiedDateDrpBx_Label = New-Object System.Windows.Forms.Label
$Disabled_Check_Comp = New-Object System.Windows.Forms.CheckBox
$OSlist_Comp = New-Object System.Windows.Forms.ListBox
$OSlist_Comp_Label = New-Object System.Windows.Forms.Label
$Matches_Comp = New-Object System.Windows.Forms.Label
$Matches_Comp_Label = New-Object System.Windows.Forms.Label
$Selected_Comp = New-Object System.Windows.Forms.Label
$Selected_Comp_Label = New-Object System.Windows.Forms.Label
$Processed_Comp = New-Object System.Windows.Forms.Label
$Processed_Comp_Label = New-Object System.Windows.Forms.Label
$Operation_Label_Comp = New-Object System.Windows.Forms.Label
$disableObject_Comp = New-Object System.Windows.Forms.RadioButton
$deleteObject_Comp = New-Object System.Windows.Forms.RadioButton
$Validation_Label_Comp = New-Object System.Windows.Forms.Label
$CancelScanButton_Comp = New-Object System.Windows.Forms.Button
$CancelOpsButton_Comp = New-Object System.Windows.Forms.Button


$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState

# form specs
$buttonSize = New-Object System.Drawing.Size(105,25)
$Form1.Text = "AD Tools - " + $ADTVersion
$Form1.Name = "adtools"
#$Form1.AutoScaleMode = 1
#$Form1.AutoScale = $true
$Form1.AutoSizeMode = 1
$Form1.AutoSize = $true
$Form1.Padding = New-Object System.Windows.Forms.Padding(10)
$Form1.DataBindings.DefaultDataSourceUpdateMode = 0
$Form1.StartPosition = "CenterScreen"

# tab control specs
$Tabcontrol1.Name = "tabControl"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 10
$System_Drawing_Point.Y = 15
$Tabcontrol1.Location = $System_Drawing_Point
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 640
$System_Drawing_Size.Width = 820
$Tabcontrol1.Size = $System_Drawing_Size
$Form1.Controls.Add($Tabcontrol1)

# Object Aging tab control specs
$TabcontrolAging.Name = "tabAgingControl"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 5
$System_Drawing_Point.Y = 85
$TabcontrolAging.Location = $System_Drawing_Point
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 520
$System_Drawing_Size.Width = 800
$TabcontrolAging.Size = $System_Drawing_Size
$objectAgingTab.Controls.Add($TabcontrolAging)


$userToolsTab.AutoSize = $true
$userToolsTab.TabIndex = 0
$userToolsTab.Text = "User Tools"
$userToolsTab.Enabled = $true
$Tabcontrol1.Controls.Add($userToolsTab)

$objectAgingTab.AutoSize = $true
$objectAgingTab.TabIndex = 1
$objectAgingTab.Text = "Object Aging Tools"
$objectAgingTab.Enabled = $true
$Tabcontrol1.Controls.Add($objectAgingTab)

$userObjTab.AutoSize = $true
$userObjTab.TabIndex = 0
$userObjTab.Text = "User Object Aging"
$userObjTab.Enabled = $true
$TabcontrolAging.Controls.Add($userObjTab)

$computerObjTab.AutoSize = $true
$computerObjTab.TabIndex = 1
$computerObjTab.Text = "Computer Object Aging"
$computerObjTab.Enabled = $true
$TabcontrolAging.Controls.Add($computerObjTab)

###################################################################
# Begin Build .Net Objects Functions

Function Check_Obj_Aging_Start()
{
    if($server.Text -eq "" -or ($script:creds -eq $null -and -not $CurrentCreds_Check.Checked ))
    {
        $ScanButton_Usr.Enabled = $false
        $ScanButton_Comp.Enabled = $false
        $ModifyButton_Usr.Enabled = $false
        $ModifyButton_Comp.Enabled = $false
        return $true
    }
    $ScanButton_Usr.Enabled = $false
    $ScanButton_Comp.Enabled = $false
    $ModifyButton_Usr.Enabled = $false
    $ModifyButton_Comp.Enabled = $false
}
Function ObjectAging()
{
    $Server.Location = New-Object System.Drawing.Size(5,35)
    $Server.Size = New-Object System.Drawing.Size(270,25)
    $Server.Text = ""
    $Server.TabIndex = 0
    $objectAgingTab.Controls.Add($Server)

    $Server_Label.Location = New-Object System.Drawing.Size(5,16) 
    $Server_Label.Size = New-Object System.Drawing.Size(270,20) 
    $Server_Label.Text = "Server Name or IP address to query"
    $objectAgingTab.Controls.Add($Server_Label) 

    $CredsButton.Location = New-Object System.Drawing.Size(285,35)
    $CredsButton.Size = New-Object System.Drawing.Size(100,20)
    $CredsButton.Text = "Get Credentials"
    $CredsButton.Enabled = $true
    $CredsButton.Add_Click(
    {
        $script:creds = Get-Credential
        if ($script:creds -ne $null)
        {
            $ScanButton_Usr.Enabled = $true
            $ScanButton_Comp.Enabled = $true
        }
    })
    $objectAgingTab.Controls.Add($CredsButton)

    $CurrentCreds_Check.Location = New-Object System.Drawing.Size(395,35)
    $CurrentCreds_Check.Size = New-Object System.Drawing.Size(120,20)
    $CurrentCreds_Check.Text = "Use Current Creds"
    $CurrentCreds_Check.Add_CheckStateChanged({
    
        if ($CurrentCreds_Check.Checked)
        {
            $CredsButton.Enabled = $false
            $ScanButton_Usr.Enabled = $true
            $ScanButton_Comp.Enabled = $true
        }
        else
        {
            $CredsButton.Enabled = $true
            $ScanButton_Usr.Enabled = $false
            $ScanButton_Comp.Enabled = $false
        }
    })
    $objectAgingTab.Controls.Add($CurrentCreds_Check)
}

Function UserToolsTab()
{
    $Server_LabelUT.Location = New-Object System.Drawing.Size(10,10) 
    $Server_LabelUT.Size = New-Object System.Drawing.Size(270,20) 
    $Server_LabelUT.Text = "PDC Emulator is used for all actions on this tab"
    $userToolsTab.Controls.Add($Server_LabelUT) 

    $CredsButtonUT.Location = New-Object System.Drawing.Size(10,40)
    $CredsButtonUT.Size = New-Object System.Drawing.Size(100,20)
    $CredsButtonUT.Text = "Get Credentials"
    $CredsButtonUT.Enabled = $true
    $CredsButtonUT.Add_Click(
    {
        $script:creds = Get-Credential
        if($script:creds -ne $null)
        {
            $LockoutButton_UT.enabled = $true
        }
    })
    $userToolsTab.Controls.Add($CredsButtonUT)

    $CurrentCreds_CheckUT.Location = New-Object System.Drawing.Size(120,40)
    $CurrentCreds_CheckUT.Size = New-Object System.Drawing.Size(120,20)
    $CurrentCreds_CheckUT.Text = "Use Current Creds"
    $CurrentCreds_CheckUT.Add_CheckStateChanged({
    
        if ($CurrentCreds_CheckUT.Checked)
        {
            $CredsButtonUT.Enabled = $false
            $LockoutButton_UT.enabled = $true
        }
        else
        {
            $CredsButtonUT.Enabled = $true
            $LockoutButton_UT.enabled = $false
        }
    })
    $userToolsTab.Controls.Add($CurrentCreds_CheckUT)

    $UserId_Label_UT.Location = New-Object System.Drawing.Size(10,80) 
    $UserId_Label_UT.Size = New-Object System.Drawing.Size(170,20) 
    $UserId_Label_UT.Text = "User ID to query"
    $userToolsTab.Controls.Add($UserId_Label_UT) 

    $UserID_UT.Location = New-Object System.Drawing.Size(10,100)
    $UserID_UT.Size = New-Object System.Drawing.Size(120,25)
    $UserID_UT.Text = ""
    $userToolsTab.Controls.Add($UserID_UT)

    $LockoutButton_UT.Location = New-Object System.Drawing.Size(10,135)
    $LockoutButton_UT.Size = New-Object System.Drawing.Size(120,25)
    $LockoutButton_UT.Text = "Scan for Lockouts"
    $LockoutButton_UT.enabled = $false
    $LockoutButton_UT.Add_Click(
        {
            $LockoutButton_UT.enabled = $false
            Lockout_UsrTool
        })
    $userToolsTab.Controls.Add($LockoutButton_UT)
}

Function UserObjectsTab()
{

    $numOfDays_DrpText_Usr.Location = New-Object System.Drawing.Size(10,15)
    $numOfDays_DrpText_Usr.Size = New-Object System.Drawing.Size(50,20)
    $numOfDays_DrpText_Usr.DropDownHeight = 100
    [Void] $numOfDays_DrpText_Usr.Items.Add("180")
    [Void] $numOfDays_DrpText_Usr.Items.Add("90")
    [Void] $numOfDays_DrpText_Usr.Items.Add("60")
    [Void] $numOfDays_DrpText_Usr.Items.Add("45")
    [Void] $numOfDays_DrpText_Usr.Items.Add("30")
    [Void] $numOfDays_DrpText_Usr.Items.Add("7 Yrs.")
    $numOfDays_DrpText_Usr.SelectedIndex = 0
    $userObjTab.Controls.Add($numOfDays_DrpText_Usr)

    $numOfDays_Label_Usr.Location = New-Object System.Drawing.Size(70,20) 
    $numOfDays_Label_Usr.Size = New-Object System.Drawing.Size(280,20) 
    $numOfDays_Label_Usr.Text = "Days since last logon (per MS LastLogonTimestamp)"
    $userObjTab.Controls.Add($numOfDays_Label_Usr) 

    $LastModifiedDate_Check_Usr.Location = New-Object System.Drawing.Size(10,50)
    $LastModifiedDate_Check_Usr.Size = New-Object System.Drawing.Size(120,30)
    $LastModifiedDate_Check_Usr.Text = "Enable LastModifiedDate"
    $LastModifiedDate_Check_Usr.Add_CheckStateChanged({
        if ($LastModifiedDate_Check_Usr.Checked)
        {
            $ModifiedDate_DrpText_Usr.Enabled = $true
            $ModifiedDateDrpBx_Label_Usr.Enabled = $true
        }
        else
        {
            $ModifiedDate_DrpText_Usr.Enabled = $false
            $ModifiedDateDrpBx_Label_Usr.Enabled = $false
        }
        })
    $userObjTab.Controls.Add($LastModifiedDate_Check_Usr)

    $ModifiedDate_DrpText_Usr.Location = New-Object System.Drawing.Size(130,55)
    $ModifiedDate_DrpText_Usr.Size = New-Object System.Drawing.Size(50,20)
    $ModifiedDate_DrpText_Usr.DropDownHeight = 100
    [Void] $ModifiedDate_DrpText_Usr.Items.Add("180")
    [Void] $ModifiedDate_DrpText_Usr.Items.Add("90")
    [Void] $ModifiedDate_DrpText_Usr.Items.Add("60")
    [Void] $ModifiedDate_DrpText_Usr.Items.Add("45")
    [Void] $ModifiedDate_DrpText_Usr.Items.Add("30")
    [Void] $ModifiedDate_DrpText_Usr.Items.Add("7 yrs.")
    $ModifiedDate_DrpText_Usr.SelectedIndex = 0
    $ModifiedDate_DrpText_Usr.Enabled = $false
    $userObjTab.Controls.Add($ModifiedDate_DrpText_Usr)

    $ModifiedDateDrpBx_Label_Usr.Location = New-Object System.Drawing.Size(180,50) 
    $ModifiedDateDrpBx_Label_Usr.Size = New-Object System.Drawing.Size(160,30) 
    $ModifiedDateDrpBx_Label_Usr.Text = "Select the number of days since last modified"
    $ModifiedDateDrpBx_Label_Usr.Enabled = $false
    $userObjTab.Controls.Add($ModifiedDateDrpBx_Label_Usr) 

    $Disabled_Check_Usr.Location = New-Object System.Drawing.Size(10,90)
    $Disabled_Check_Usr.Size = New-Object System.Drawing.Size(300,30)
    $Disabled_Check_Usr.Text = "Search for disabled accounts with LastModifiedDate date greater than selected number of days"
    $Disabled_Check_Usr.Add_CheckStateChanged({
        if ($Disabled_Check_Usr.Checked)
        {
            $LastModifiedDate_Check_Usr.Checked = $true
            $numOfDays_DrpText_Usr.Enabled = $false
            $deleteObject_Usr.Checked = $true
            $disableObject_Usr.Enabled = $false
        }
        else
        {
            $LastModifiedDate_Check_Usr.Checked = $false
            $numOfDays_DrpText_Usr.Enabled = $true
            $disableObject_Usr.Enabled = $true
            $disableObject_Usr.Checked = $true
        }
        })
    $userObjTab.Controls.Add($Disabled_Check_Usr)

    $Matches_Usr.Location = New-Object System.Drawing.Size(10,170) 
    $Matches_Usr.Size = New-Object System.Drawing.Size(60,20) 
    $Matches_Usr.Text = "0"
    $userObjTab.Controls.Add($Matches_Usr) 

    $Matches_Usr_Label.Location = New-Object System.Drawing.Size(100,170) 
    $Matches_Usr_Label.Size = New-Object System.Drawing.Size(250,20) 
    $Matches_Usr_Label.Text = "- Matching users"
    $userObjTab.Controls.Add($Matches_Usr_Label) 

    $Operation_Label_Usr.Location = New-Object System.Drawing.Size(10,310) 
    $Operation_Label_Usr.Size = New-Object System.Drawing.Size(250,20) 
    $Operation_Label_Usr.Text = "Select the desired operation"
    $userObjTab.Controls.Add($Operation_Label_Usr) 

    $disableObject_Usr.Location = New-Object System.Drawing.Size(10,330)
    $disableObject_Usr.Text = "Disable Objects"
    $disableObject_Usr.Checked = $true
    $userObjTab.Controls.Add($disableObject_Usr)

    $deleteObject_Usr.Location = New-Object System.Drawing.Size(160,330) 
    $deleteObject_Usr.Text = "Delete Objects"
    $userObjTab.Controls.Add($deleteObject_Usr)

    $ScanButton_Usr.Location = New-Object System.Drawing.Size(10,130)
    $ScanButton_Usr.Size = New-Object System.Drawing.Size(75,25)
    $ScanButton_Usr.Text = "Scan"
    $ScanButton_Usr.Add_Click(
        {
            if(Check_Obj_Aging_Start)
            {
                [System.Windows.Forms.MessageBox]::Show("Missing server info or credentials", "Status")
                return
            }
            $script:Cancel_Usr = $false
            $CancelButton_Usr.Enabled = $true
            Scan_Usr
            $CancelButton_Usr.Enabled = $false

            $script:array_Usr = @()
            foreach ( $child in $listMatching_Usr.Values )
            {
                $script:array_Usr += $child
                [System.Windows.Forms.Application]::DoEvents()
            }

            $Matches_Usr.Text = $listMatching_Usr.Count.ToString()

            $DisplayButton_Usr.Enabled = $true
            $ExportCSVButton_Usr.Enabled = $true

            [System.Windows.Forms.MessageBox]::Show("Scan completed", "Status")
        })

    $userObjTab.Controls.Add($ScanButton_Usr)

    $CancelButton_Usr.Location = New-Object System.Drawing.Size(95,130)
    $CancelButton_Usr.Size = New-Object System.Drawing.Size(75,25)
    $CancelButton_Usr.Text = "Cancel"
    $CancelButton_Usr.Add_Click({
        $script:Cancel_Usr = $true
        $CancelButton_Usr.Enabled = $false
    })
    $userObjTab.Controls.Add($CancelButton_Usr)

    $ModifyButton_Usr.Location = New-Object System.Drawing.Size(10,375)
    $ModifyButton_Usr.Size = New-Object System.Drawing.Size(140,25)
    $ModifyButton_Usr.Text = "Perform Operation"
    $ModifyButton_Usr.Enabled = $false
    $ModifyButton_Usr.Add_Click(
        {
            if(Check_Obj_Aging_Start)
            {
                [System.Windows.Forms.MessageBox]::Show("Missing server info or credentials", "Status")
                return
            }
            Perform_Operation_Usr
        })
    $userObjTab.Controls.Add($ModifyButton_Usr)

    $ImportCSVButton_Usr.Location = New-Object System.Drawing.Size(155,375)
    $ImportCSVButton_Usr.Size = New-Object System.Drawing.Size(140,25)
    $ImportCSVButton_Usr.Text = "Import CSV"
    $ImportCSVButton_Usr.Add_Click({Import_CSV_Usr;$ScanButton_Usr.Enabled = $false;$DisplayButton_Usr.Enabled = $false;$ExportCSVButton_Usr.Enabled = $false;$ModifyButton_Usr.Enabled = $true})
    $userObjTab.Controls.Add($ImportCSVButton_Usr)

    $DisplayButton_Usr.Location = New-Object System.Drawing.Size(10,210)
    $DisplayButton_Usr.Size = New-Object System.Drawing.Size(140,25)
    $DisplayButton_Usr.Text = "Display Accounts"
    $DisplayButton_Usr.Add_Click({Display_Selections_Usr;$ModifyButton_Usr.Enabled = $true; $ExportCSVButton_Usr.Enabled = $true})
    $DisplayButton_Usr.Enabled = $false
    $userObjTab.Controls.Add($DisplayButton_Usr)

    $ExportCSVButton_Usr.Location = New-Object System.Drawing.Size(160,210)
    $ExportCSVButton_Usr.Size = New-Object System.Drawing.Size(140,25)
    $ExportCSVButton_Usr.Text = "Export CSV"
    #$ExportCSVButton_Usr.Enabled = $false
    $ExportCSVButton_Usr.Add_Click({Export_CSV_Usr})
    $userObjTab.Controls.Add($ExportCSVButton_Usr)

    $ResetButton_Usr.Location = New-Object System.Drawing.Size(10,245)
    $ResetButton_Usr.Size = New-Object System.Drawing.Size(140,25)
    $ResetButton_Usr.Text = "Reset"
    $ResetButton_Usr.Enabled = $true
    $ResetButton_Usr.Add_Click({Init_Sys_Usr})
    $userObjTab.Controls.Add($ResetButton_Usr)
}

Function CompObjectsTab()
{
    $ScanButton_Comp.Location = New-Object System.Drawing.Size(10,260)
    $ScanButton_Comp.Size = $buttonSize
    $ScanButton_Comp.Text = "Scan"
    $ScanButton_Comp.Enabled = $false
    $ScanButton_Comp.Add_Click(
        {
            if(Check_Obj_Aging_Start)
            {
                [System.Windows.Forms.MessageBox]::Show("Missing server info or credentials", "Status")
                return
            }
            $script:Cancel_Comp = $false
            $CancelScanButton_Comp.Enabled = $true
            $ScanButton_Comp.Enabled = $false
            $ImportCSVButton_Comp.Enabled = $false
            Scan_Comp
            $ScanButton_Comp.Enabled = $true
            $CancelScanButton_Comp.Enabled = $false

            if ($LastModifiedDate_Check_Comp.Checked)
            {
                Filters_Comp
            } 

            if ($PingCheck_Comp.Checked)
            {
                Validate_Comp
            }

            LoadOSs_Comp

            $Matches_Comp.Text = $listMatching_Comp.Count.ToString()

            $DisplayButton_Comp.Enabled = $true
            $ExportCSVButton_Comp.Enabled = $true

            [System.Windows.Forms.MessageBox]::Show("Scan completed", "Status")

        })

    $computerObjTab.Controls.Add($ScanButton_Comp)

    $CancelScanButton_Comp.Location = New-Object System.Drawing.Size(125,260)
    $CancelScanButton_Comp.Size = $buttonSize
    $CancelScanButton_Comp.Text = "Cancel"
    $CancelScanButton_Comp.Add_Click(
                            {
                                $script:Cancel_Comp = $true
                                $CancelScanButton_Comp.Enabled = $false
                                $ScanButton_Comp.Enabled = $true
                            })

    $computerObjTab.Controls.Add($CancelScanButton_Comp)

    $ModifyButton_Comp.Location = New-Object System.Drawing.Size(10,445)
    $ModifyButton_Comp.Size = $buttonSize
    $ModifyButton_Comp.Text = "Perform Operation"
    $ModifyButton_Comp.Enabled = $false
    $ModifyButton_Comp.Add_Click(
        {
            if(Check_Obj_Aging_Start)
            {
                [System.Windows.Forms.MessageBox]::Show("Missing server info or credentials", "Status")
                return
            }
            $script:Cancel_Comp = $false
            $CancelOpsButton_Comp.Enabled = $true
            $ImportCSVButton_Comp.Enabled = $false
            Perform_Operation_Comp
            $CancelOpsButton_Comp.Enabled = $false
        })
    $computerObjTab.Controls.Add($ModifyButton_Comp)

    $CancelOpsButton_Comp.Location = New-Object System.Drawing.Size(125,445)
    $CancelOpsButton_Comp.Size = $buttonSize
    $CancelOpsButton_Comp.Text = "Cancel Ops"
    $CancelOpsButton_Comp.Enabled = $false
    $CancelOpsButton_Comp.Add_Click({
            $script:Cancel_Comp = $true
            $CancelOpsButton_Comp.Enabled = $false
    })
    $computerObjTab.Controls.Add($CancelOpsButton_Comp)

    $ImportCSVButton_Comp.Location = New-Object System.Drawing.Size(240,445)
    $ImportCSVButton_Comp.Size = $buttonSize
    $ImportCSVButton_Comp.Text = "Import CSV"
    $ImportCSVButton_Comp.Add_Click(
        {
            Import_CSV_Comp;
            LoadOSs_Comp
            #$Matches_Comp.Text = $listMatching_Comp.Count.ToString()
            $DisplayButton_Comp.Enabled = $true
            $ScanButton_Comp.Enabled = $false;
            $ModifyButton_Comp.Enabled = $true
        })
    $computerObjTab.Controls.Add($ImportCSVButton_Comp)

    $DisplayButton_Comp.Location = New-Object System.Drawing.Size(355,445)
    $DisplayButton_Comp.Size = $buttonSize
    $DisplayButton_Comp.Text = "Select / Display"
    $DisplayButton_Comp.Add_Click(
        {
            Display_Selections_Comp
            $ModifyButton_Comp.Enabled = $true
            $script:exportall_Comp = $false
            $ExportCSVButton_Comp.Text = "Export Selected"
        })
    $DisplayButton_Comp.Enabled = $false
    $computerObjTab.Controls.Add($DisplayButton_Comp)

    
    $ExportCSVButton_Comp.Location = New-Object System.Drawing.Size(470,445)
    $ExportCSVButton_Comp.Size = $buttonSize
    $ExportCSVButton_Comp.Text = "Export All"
    $ExportCSVButton_Comp.Enabled = $false
    $ExportCSVButton_Comp.Add_Click({Export_CSV_Comp})
    $computerObjTab.Controls.Add($ExportCSVButton_Comp)

    $ResetButton_Comp.Location = New-Object System.Drawing.Size(675,445)
    $ResetButton_Comp.Size = $buttonSize
    $ResetButton_Comp.Text = "Reset"
    $ResetButton_Comp.Enabled = $true
    $ResetButton_Comp.Add_Click({Init_Sys_Comp})
    $computerObjTab.Controls.Add($ResetButton_Comp)

    $PingCheck_Comp.Location = New-Object System.Drawing.Size(10,75)
    $PingCheck_Comp.Size = New-Object System.Drawing.Size(70,30)
    $PingCheck_Comp.Text = "Ping Check"
    $computerObjTab.Controls.Add($PingCheck_Comp)

    $Validation_Label_Comp.Location = New-Object System.Drawing.Size(10,55) 
    $Validation_Label_Comp.Size = New-Object System.Drawing.Size(330,20) 
    $Validation_Label_Comp.Text = "----------------- Validation Checks"
    $computerObjTab.Controls.Add($Validation_Label_Comp) 

    $Filters_Label_Comp.Location = New-Object System.Drawing.Size(10,115) 
    $Filters_Label_Comp.Size = New-Object System.Drawing.Size(330,20) 
    $Filters_Label_Comp.Text = "----------------- Filters"
    $computerObjTab.Controls.Add($Filters_Label_Comp) 

    $numOfDays_DrpText_Comp.Location = New-Object System.Drawing.Size(10,145)
    $numOfDays_DrpText_Comp.Size = New-Object System.Drawing.Size(50,20)
    $numOfDays_DrpText_Comp.DropDownHeight = 100
    [Void] $numOfDays_DrpText_Comp.Items.Add("180")
    [Void] $numOfDays_DrpText_Comp.Items.Add("90")
    [Void] $numOfDays_DrpText_Comp.Items.Add("60")
    [Void] $numOfDays_DrpText_Comp.Items.Add("45")
    [Void] $numOfDays_DrpText_Comp.Items.Add("30")
    $numOfDays_DrpText_Comp.SelectedIndex = 0
    $computerObjTab.Controls.Add($numOfDays_DrpText_Comp)

    $numOfDays_Label_Comp.Location = New-Object System.Drawing.Size(70,140) 
    $numOfDays_Label_Comp.Size = New-Object System.Drawing.Size(270,30) 
    $numOfDays_Label_Comp.Text = "Days since LastLogonTimestamp for unused accounts (Primary search parameter per MicroSoft)"
    $computerObjTab.Controls.Add($numOfDays_Label_Comp) 

    $LastModifiedDate_Check_Comp.Location = New-Object System.Drawing.Size(10,180)
    $LastModifiedDate_Check_Comp.Size = New-Object System.Drawing.Size(120,30)
    $LastModifiedDate_Check_Comp.Text = "Enable LastModifiedDate"
    $LastModifiedDate_Check_Comp.Add_CheckStateChanged({
        if ($LastModifiedDate_Check_Comp.Checked)
        {
            $ModifiedDate_DrpText_Comp.Enabled = $true
            $ModifiedDateDrpBx_Label.Enabled = $true
        }
        else
        {
            $ModifiedDate_DrpText_Comp.Enabled = $false
            $ModifiedDateDrpBx_Label.Enabled = $false
        }
        })
    $computerObjTab.Controls.Add($LastModifiedDate_Check_Comp)

    $ModifiedDate_DrpText_Comp.Location = New-Object System.Drawing.Size(130,185)
    $ModifiedDate_DrpText_Comp.Size = New-Object System.Drawing.Size(50,20)
    $ModifiedDate_DrpText_Comp.DropDownHeight = 100
    [Void] $ModifiedDate_DrpText_Comp.Items.Add("180")
    [Void] $ModifiedDate_DrpText_Comp.Items.Add("90")
    [Void] $ModifiedDate_DrpText_Comp.Items.Add("60")
    [Void] $ModifiedDate_DrpText_Comp.Items.Add("45")
    [Void] $ModifiedDate_DrpText_Comp.Items.Add("30")
    $ModifiedDate_DrpText_Comp.SelectedIndex = 0
    $ModifiedDate_DrpText_Comp.Enabled = $false
    $computerObjTab.Controls.Add($ModifiedDate_DrpText_Comp)

    $ModifiedDateDrpBx_Label.Location = New-Object System.Drawing.Size(180,180) 
    $ModifiedDateDrpBx_Label.Size = New-Object System.Drawing.Size(160,30) 
    $ModifiedDateDrpBx_Label.Text = "Select the number of days since last modified"
    $ModifiedDateDrpBx_Label.Enabled = $false
    $computerObjTab.Controls.Add($ModifiedDateDrpBx_Label) 

    $Disabled_Check_Comp.Location = New-Object System.Drawing.Size(10,220)
    $Disabled_Check_Comp.Size = New-Object System.Drawing.Size(300,30)
    $Disabled_Check_Comp.Text = "Search for disabled accounts with LastModifiedDate date greater than Selected number of days"
    $Disabled_Check_Comp.Add_CheckStateChanged({
        if ($Disabled_Check_Comp.Checked)
        {
            $PingCheck_Comp.Enabled = $false
            $LastModifiedDate_Check_Comp.Checked = $true
            $numOfDays_DrpText_Comp.Enabled = $false
            $PingCheck_Comp.Enabled = $false
            $deleteObject_Comp.Checked = $true
            $disableObject_Comp.Enabled = $false
        }
        else
        {
            $LastModifiedDate_Check_Comp.Checked = $false
            $numOfDays_DrpText_Comp.Enabled = $true
            $PingCheck_Comp.Enabled = $true
            $disableObject_Comp.Enabled = $true
            $disableObject_Comp.Checked = $true
        }
        })
    $computerObjTab.Controls.Add($Disabled_Check_Comp)

    $OSlist_Comp.Location = New-Object System.Drawing.Size(340,25)
    $OSlist_Comp.Size = New-Object System.Drawing.Size(440,415)
    $OSlist_Comp.SelectionMode = [System.Windows.Forms.SelectionMode]::MultiExtended
    $computerObjTab.Controls.Add($OSlist_Comp) 

    $OSlist_Comp_Label.Location = New-Object System.Drawing.Size(340,6) 
    $OSlist_Comp_Label.Size = New-Object System.Drawing.Size(430,20) 
    $OSlist_Comp_Label.Text = "Select the Operating Systems to modify (ctrl-click for multiple)"
    $computerObjTab.Controls.Add($OSlist_Comp_Label) 

    $Matches_Comp.Location = New-Object System.Drawing.Size(10,300) 
    $Matches_Comp.Size = New-Object System.Drawing.Size(60,20) 
    $Matches_Comp.Text = "0"
    $computerObjTab.Controls.Add($Matches_Comp) 

    $Matches_Comp_Label.Location = New-Object System.Drawing.Size(100,300) 
    $Matches_Comp_Label.Size = New-Object System.Drawing.Size(250,20) 
    $Matches_Comp_Label.Text = "- Matching objects"
    $computerObjTab.Controls.Add($Matches_Comp_Label) 

    $Selected_Comp.Location = New-Object System.Drawing.Size(10,325) 
    $Selected_Comp.Size = New-Object System.Drawing.Size(60,20) 
    $Selected_Comp.Text = "0"
    $computerObjTab.Controls.Add($Selected_Comp) 

    $Selected_Comp_Label.Location = New-Object System.Drawing.Size(100,325) 
    $Selected_Comp_Label.Size = New-Object System.Drawing.Size(250,20) 
    $Selected_Comp_Label.Text = "- Selected objects"
    $computerObjTab.Controls.Add($Selected_Comp_Label)

    $Processed_Comp.Location = New-Object System.Drawing.Size(10,350) 
    $Processed_Comp.Size = New-Object System.Drawing.Size(60,20) 
    $Processed_Comp.Text = "0"
    $computerObjTab.Controls.Add($Processed_Comp) 

    $Processed_Comp_Label.Location = New-Object System.Drawing.Size(100,350) 
    $Processed_Comp_Label.Size = New-Object System.Drawing.Size(250,20) 
    $Processed_Comp_Label.Text = "- Processed objects"
    $computerObjTab.Controls.Add($Processed_Comp_Label) 


    $Operation_Label_Comp.Location = New-Object System.Drawing.Size(10,380) 
    $Operation_Label_Comp.Size = New-Object System.Drawing.Size(250,20) 
    $Operation_Label_Comp.Text = "Select the desired operation"
    $computerObjTab.Controls.Add($Operation_Label_Comp) 

    $disableObject_Comp.Location = New-Object System.Drawing.Size(10,400)
    $disableObject_Comp.Text = "Disable Objects"
    $disableObject_Comp.Checked = $true
    $computerObjTab.Controls.Add($disableObject_Comp)

    $deleteObject_Comp.Location = New-Object System.Drawing.Size(160,400) 
    $deleteObject_Comp.Text = "Delete Objects"
    $computerObjTab.Controls.Add($deleteObject_Comp)

}

# End Build .Net Objects Functions
###################################################################

###################################################################
# Begin User Tools Functions

Function Lockout_UsrTool()
{
    try
    {
	    [string] $ComputerNameUT = ((Get-ADDomainController -Discover -Service PrimaryDC).HostName)
	    #[string] $UserName = $(
		#    Add-Type -AssemblyName Microsoft.VisualBasic
		#    [Microsoft.VisualBasic.Interaction]::InputBox('Enter the username to check','Check lockouts', $env:USERNAME)
	    #)
		
	    if (!$UserID_UT.Text) { exit }
        $user = $UserID_UT.Text

	    $filter = "*[System[EventID=4740] and EventData[Data[@Name='TargetUserName']='$user']]"

        if ($CurrentCreds_CheckUT.Checked)
        {
            $Events = Get-WinEvent -ComputerName $ComputerNameUT -Logname Security -FilterXPath $filter -ErrorAction Stop
        }
        else
        {
            $Events = Get-WinEvent -Credential $creds -ComputerName $ComputerNameUT -Logname Security -FilterXPath $filter -ErrorAction Stop
        }
	    $Events | Select-Object TimeCreated,
	    @{Name='User Name';Expression={$_.Properties[0].Value}},
	    @{Name='Source Host';Expression={$_.Properties[1].Value}} | Out-GridView -Wait -Title "ADM Lockouts"
    }
	
    catch
    {
	    if ($_.Exception -match "No events were found that match the specified selection criteria") {
		    (new-object -ComObject wscript.shell).Popup("No recent lockouts were found",0,"None Found")
	    } else {
		    Throw $_.Exception
	    }
    }
    $LockoutButton_UT.enabled = $true
}
# End User Tools Functions
###################################################################


###################################################################
# Begin User Object Functions

Function Init_Sys_Usr()
{
    $script:date = Get-Date
    $filters_Usr = $false


    $array_Usr = @()            # used for selected and displayed data
    $listMatching_Usr = @{}
    $failures_Usr = @{}
    $ScanButton_Usr.Enabled = $false
    $CancelButton_Usr.Enabled = $false
    $ModifyButton_Usr.Enabled = $false
    $ExportCSVButton_Usr.Enabled = $false
    $DisplayButton_Usr.Enabled = $false

    #$Selected.Text = $script:array.Count.ToString()
    $Matches_Usr.Text = $listMatching_Usr.Count.ToString()

}

Function Scan_Usr()
{
    $ScanButton_Usr.Enabled = $false

    switch ($numOfDays_DrpText_Usr.SelectedIndex) 
    { 
        0 {$daysOld = 180} 
        1 {$daysOld = 90} 
        2 {$daysOld = 60} 
        3 {$daysOld = 45} 
        4 {$daysOld = 30} 
        5 {$daysOld = 2555} 
        default {$daysOld = 180}
    }

    Write-Host "Processing"
    try{

        
        if($CurrentCreds_Check.Checked)
        {
            Get-ADUser -Filter * -Server $Server.Text -Properties Name,CanonicalName,Description,Enabled,LastLogonDate,lastLogonTimeStamp,Modified,modifyTimeStamp,PasswordLastSet,pwdLastSet | 
            ForEach-Object {

                if ($listMatching_Usr.Count%10 -eq 0){
                        #Write-Host "." -NoNewline
                    }

                if ($Disabled_Check_Usr.Checked)
                {
                        switch ($ModifiedDate_DrpText_Usr.SelectedIndex) 
                    { 
                        0 {$LastModifiedDate = 180} 
                        1 {$LastModifiedDate = 90} 
                        2 {$LastModifiedDate = 60} 
                        3 {$LastModifiedDate = 45} 
                        4 {$LastModifiedDate = 30} 
                        5 {$LastModifiedDate = 2555} 
                        default {$LastModifiedDate = 180}
                    }
                    if (($script:date - $_.modifyTimeStamp).Days -ge $LastModifiedDate -AND $_.Enabled -eq $false)
                    {
                        $_ | Add-Member -MemberType NoteProperty -Name "Days Since Last Mod" -Value ($script:date - $_.modifyTimeStamp).Days -Force
                        $listMatching_Usr.add($_.CanonicalName,$_) # Use CanonicalName to capture duplicate entries
                    }
                }

                else
                {
                    if (($script:date - ([datetime]::FromFileTime($_.lastLogonTimeStamp))).Days -ge $daysOld -AND $_.Enabled -eq $true)
                    {
                        if ($_.pwdLastSet -ne $null)
                        {
                            $_ | Add-Member -MemberType NoteProperty -Name "Pwd Age" -Value ($script:date - ([datetime]::FromFileTime($_.pwdLastSet))).Days -Force
                        }
                        else
                        {
                            $_ | Add-Member -MemberType NoteProperty -Name "Pwd Age" -Value $null -Force
                        }

                        $listMatching_Usr.add($_.CanonicalName,$_) # Use CanonicalName to capture duplicate entries
                    }
                }
                # uncomment to debug
                #if ($listMatching_Usr.Count -ge 50) { break } 
                $Matches_Usr.Text = $listMatching_Usr.Count.ToString()
                if($script:Cancel_Usr -eq $true){
                    break
                }
                [System.Windows.Forms.Application]::DoEvents()
            }
        }
        else
        {
            Get-ADUser -Filter * -Credential $creds -Server $Server.Text -Properties Name,CanonicalName,Description,Enabled,LastLogonDate,lastLogonTimeStamp,Modified,modifyTimeStamp,PasswordLastSet,pwdLastSet | 
            
            ForEach-Object {

                if ($listMatching_Usr.Count%10 -eq 0){
                    Write-Host "." -NoNewline
                    }

                if ($Disabled_Check_Usr.Checked)
                {
                        switch ($ModifiedDate_DrpText_Usr.SelectedIndex) 
                    { 
                        0 {$LastModifiedDate = 180} 
                        1 {$LastModifiedDate = 90} 
                        2 {$LastModifiedDate = 60} 
                        3 {$LastModifiedDate = 45} 
                        4 {$LastModifiedDate = 30} 
                        5 {$LastModifiedDate = 2555} 
                        default {$LastModifiedDate = 180}
                    }
                    if (($script:date - $_.modifyTimeStamp).Days -ge $LastModifiedDate -AND $_.Enabled -eq $false)
                    {
                        $_ | Add-Member -MemberType NoteProperty -Name "Days Since Last Mod" -Value ($script:date - $_.modifyTimeStamp).Days -Force
                        $listMatching_Usr.add($_.CanonicalName,$_) # Use CanonicalName to capture duplicate entries
                    }
                }

                else
                {
                    if (($script:date - ([datetime]::FromFileTime($_.lastLogonTimeStamp))).Days -ge $daysOld -AND $_.Enabled -eq $true)
                    {
                        if ($_.pwdLastSet -ne $null)
                        {
                            $_ | Add-Member -MemberType NoteProperty -Name "Pwd Age" -Value ($script:date - ([datetime]::FromFileTime($_.pwdLastSet))).Days -Force
                        }
                        else
                        {
                            $_ | Add-Member -MemberType NoteProperty -Name "Pwd Age" -Value $null -Force
                        }

                        $listMatching_Usr.add($_.CanonicalName,$_) # Use CanonicalName to capture duplicate entries
                    }
                }
                # uncomment to debug
                #if ($listMatching_Usr.Count -ge 50) { break } 
                $Matches_Usr.Text = $listMatching_Usr.Count.ToString()
                if($script:Cancel_Usr -eq $true){
                    break
                }
                [System.Windows.Forms.Application]::DoEvents()
            }
        }
    }
    Catch{
      Write-Host "Scan broke out of processing, continuing back to main program" -ForegroundColor Red -BackgroundColor Gray
      Write-Host $_.ScriptStackTrace
    }
}

Function Perform_Operation_Usr()
{

    foreach ($child in $script:array_Usr ) 
    {
        if ($disableObject_Usr.Checked)
        {
            try
            {
                write-host "Disabling - "$child.SamAccountName
                Set-ADUser -Identity $child.SamAccountName -Credential $creds -Server $Server.Text -enabled $False
            }

            catch
            {
                $failures_Usr.Add($child.SamAccountName,$child)
            }

        }
        elseif ($deleteObject_Usr.Checked)
        {
            try
            {
                Remove-ADUser -Identity $child.SamAccountName -Credential $creds -Server $Server.Text -Confirm:$False
            }
            catch
            {
                $failures_Usr.Add($child.SamAccountName,$child)
            }
        }
    }

    if ($failures_Usr.Count)
    {
        $OUTPUT = [System.Windows.Forms.MessageBox]::Show("Modification failures detected, click Yes to select destination file for report and no to disregard", "Status", 4)
        if ($OUTPUT -eq "YES")
        {
            # Request the filename to write data to
            $fd = New-Object system.windows.forms.savefiledialog
            $fd.showdialog()
            $fd.filename

            $failures_Usr.Values | Export-Csv -Path $fd.filename –NoTypeInformat
        }
    }

    [System.Windows.Forms.MessageBox]::Show("Modification process completed", "Status")

}

Function Import_CSV_Usr()
{
    # Get the file containing the server list
    $fd = New-Object system.windows.forms.openfiledialog
    $fd.showdialog()
    $fd.filename


    # Setup the data
    $script:array_Usr = Import-Csv -Path $fd.FileName

    [System.Windows.Forms.MessageBox]::Show("CSV import completed", "Status")

}

Function Export_CSV_Usr()
{
    # Request the filename to write data to
    $fd = New-Object system.windows.forms.savefiledialog
    $fd.showdialog()
    $fd.filename

    $array_Usr | Export-Csv -Path $fd.filename –NoTypeInformation

    [System.Windows.Forms.MessageBox]::Show("Export CSV completed", "Status")

}


Function Display_Selections_Usr()
{
    #$Selected.Text = $script:array.Count.ToString()
    $script:array_Usr | Out-GridView
}


# End User Object Functions
###################################################################

###################################################################
# Begin Computer Object Functions
function Init_Sys_Comp()
{
    $script:Cancel_Comp = $false
    $CancelOpsButton_Comp.Enabled = $flase
    $script:date = Get-Date
    $script:Cancel_Comp = $false
    $CancelScanButton_Comp.Enabled = $false
    $script:array_Comp = @{}            # used for Selected and displayed data
    $script:listMatching_Comp = @{}
    $script:listOS_Comp = @{}
    $script:failures_Comp = @{}
    $OSlist_Comp.Items.Clear()
    $DisplayButton_Comp.Enabled = $false

    $rc = Check_Obj_Aging_Start

    $Selected_Comp.Text = $array_Comp.Count.ToString()
    $Matches_Comp.Text = $listMatching_Comp.Count.ToString()

}

function Scan_Comp()
{
    switch ($numOfDays_DrpText_Comp.SelectedIndex) 
    { 
        0 {$daysOld = 180} 
        1 {$daysOld = 90} 
        2 {$daysOld = 60} 
        3 {$daysOld = 45} 
        4 {$daysOld = 30} 
        default {$daysOld = 180}
    }

    try{

        
        if($CurrentCreds_Check.Checked)
        {
            Get-ADComputer -filter * -Server $Server.Text -Properties Name,CanonicalName,Description,Enabled,IPv4Address,LastLogonDate,lastLogonTimeStamp,Modified,modifyTimeStamp,OperatingSystem,PasswordLastSet,pwdLastSet |
                ForEach-Object {

                    if ($Disabled_Check_Comp.Checked)
                    {
                        switch ($ModifiedDate_DrpText_Comp.SelectedIndex) 
                        { 
                            0 {$LastModifiedDate = 180} 
                            1 {$LastModifiedDate = 90} 
                            2 {$LastModifiedDate = 60} 
                            3 {$LastModifiedDate = 45} 
                            4 {$LastModifiedDate = 30} 
                            default {$LastModifiedDate = 180}
                        }

                        if (($script:date - $_.modifyTimeStamp).Days -ge $LastModifiedDate -and $_.Enabled -eq $false)
                        {
                            $_ | Add-Member -MemberType NoteProperty -Name "Days Since Last Mod" -Value ($script:date - $_.modifyTimeStamp).Days -Force
                            $listMatching_Comp.add($_.CanonicalName,$_) # Use CanonicalName to capture duplicate entries
                        }
                    }

                    else
                    {
                        if (($script:date - ([datetime]::FromFileTime($_.lastLogonTimeStamp))).Days -ge $daysOld -and $_.Enabled -eq $true)
                        {
                            $listMatching_Comp.add($_.CanonicalName,$_) # Use CanonicalName to capture duplicate entries
                        }
                    }
                # uncomment to debug
                #if ($listMatching_Comp.Count -ge 150) { break } 
                $Matches_Comp.Text = $listMatching_Comp.Count.ToString()
                if($script:Cancel_Comp -eq $true){
                    break
                }
                [System.Windows.Forms.Application]::DoEvents()
            }
        }
        else
        {
            Get-ADComputer -filter * -Credential $creds -Server $Server.Text -Properties Name,CanonicalName,Description,Enabled,IPv4Address,LastLogonDate,lastLogonTimeStamp,Modified,modifyTimeStamp,OperatingSystem,PasswordLastSet,pwdLastSet |
                ForEach-Object {

                    if ($Disabled_Check_Comp.Checked)
                    {
                        switch ($ModifiedDate_DrpText_Comp.SelectedIndex) 
                        { 
                            0 {$LastModifiedDate = 180} 
                            1 {$LastModifiedDate = 90} 
                            2 {$LastModifiedDate = 60} 
                            3 {$LastModifiedDate = 45} 
                            4 {$LastModifiedDate = 30} 
                            default {$LastModifiedDate = 180}
                        }

                        if (($script:date - $_.modifyTimeStamp).Days -ge $LastModifiedDate -and $_.Enabled -eq $false)
                        {
                            $_ | Add-Member -MemberType NoteProperty -Name "Days Since Last Mod" -Value ($script:date - $_.modifyTimeStamp).Days -Force
                            $listMatching_Comp.add($_.CanonicalName,$_) # Use CanonicalName to capture duplicate entries
                        }
                    }

                    else
                    {
                        if (($script:date - ([datetime]::FromFileTime($_.lastLogonTimeStamp))).Days -ge $daysOld -and $_.Enabled -eq $true)
                        {
                            $listMatching_Comp.add($_.CanonicalName,$_) # Use CanonicalName to capture duplicate entries
                        }
                    }
                # uncomment to debug
                #if ($listMatching_Comp.Count -ge 150) { break } 
                $Matches_Comp.Text = $listMatching_Comp.Count.ToString()
                if($script:Cancel_Comp -eq $true){
                    break
                }
                [System.Windows.Forms.Application]::DoEvents()
            }
        }
    }
    Catch{
        Write-Host "Scan broke out of processing, continuing back to main program" -ForegroundColor Red -BackgroundColor Gray  
        Write-Host $_.ScriptStackTrace
      }
  
}

function Perform_Operation_Comp()
{
    if ($disableObject_Comp.Checked)
    {
        Write-Host "Processing 'Disable' objects"
    }
    elseif ($deleteObject_Comp.Checked)
    {
        Write-Host "Processing 'Delete' objects"
    }
    $counter = 0
    try{
        foreach ($child in $array_Comp.Values ) 
        {
            if ($disableObject_Comp.Checked)
            {
                try
                {
                    if($CurrentCreds_Check.Checked){
                        set-ADComputer -Identity $child.SamAccountName -Server $Server.Text -enabled $False
                    }
                    else{
                        set-ADComputer -Identity $child.SamAccountName -Credential $creds -Server $Server.Text -enabled $False
                    }
                }

                catch
                {
                    $failures_Comp.Add($child.SamAccountName,$child)
                }

            }
            elseif ($deleteObject_Comp.Checked)
            {
                try
                {
                    $x = Get-ADComputer $child.SamAccountName  -Server $Server.Text
                    if($CurrentCreds_Check.Checked){
                        #Remove-ADComputer -Identity $child.SamAccountName -Server $Server.Text -Confirm:$False
                        Remove-ADObject -Identity $x -Server $Server.Text -Recursive -Confirm:$False
                    }
                    else{
                        #Remove-ADComputer -Identity $child.SamAccountName -Credential $creds -Server $Server.Text -Confirm:$False
                        Remove-ADObject -Identity $x -Server $Server.Text -Recursive -Confirm:$False -Credential $creds
                    }
                }
                catch
                {
                    Write-Host "Processing failed on host $child.SamAccountName"
                    $failures_Comp.Add($child.SamAccountName,$child)
                }
            }
            if($script:Cancel_Comp -eq $true){
                break
            }
            $counter += 1
            $Processed_Comp.Text = $counter.ToString()
            [System.Windows.Forms.Application]::DoEvents()
            if($counter%100 -eq 0)
            {
                $Time = [System.Diagnostics.Stopwatch]::StartNew()
                while($Time.ElapsedMilliseconds -lt 10000)
                {
                    [System.Windows.Forms.Application]::DoEvents()
                }
            }
        }
    }
    catch{
        Write-Host('Operation broke out of processing, continuing back to main program"')
        Write-Host $_.ScriptStackTrace
    }

    if ($failures_Comp.Count)
    {
        $OUTPUT_Comp = [System.Windows.Forms.MessageBox]::Show("Modification failures detected, click Yes to select destination file for report and no to disregard", "Status", 4)
        if ($OUTPUT_Comp -eq "YES")
        {
            # Request the filename to write data to
            $fd = New-Object system.windows.forms.savefiledialog
            $fd.showdialog()
            $fd.filename

            $failures_Comp.Values | Export-Csv -Path $fd.filename –NoTypeInformat
        }
    }

    [System.Windows.Forms.MessageBox]::Show("Modification process completed", "Status")

}

function Import_CSV_Comp()
{
    # Get the file containing the server list
    $fd = New-Object system.windows.forms.openfiledialog
    $fd.showdialog()
    $fd.filename


    # Setup the data
    $array_import = @()
    $array_import = Import-Csv -Path $fd.FileName
    $listMatching_Comp.Clear()
    foreach ($child in $array_import)
    {
        $listMatching_Comp.Add($child.CanonicalName,$child)
        $Matches_Comp.Text = $listMatching_Comp.Count.ToString()
        [System.Windows.Forms.Application]::DoEvents()
    }

    [System.Windows.Forms.MessageBox]::Show("CSV import completed", "Status")

}

function Export_CSV_Comp()
{
    # Request the filename to write data to
    $fd = New-Object system.windows.forms.savefiledialog
    $fd.showdialog()
    $fd.filename

    if ($exportall_Comp)
    {
        $listMatching_Comp.Values | Export-Csv -Path $fd.filename –NoTypeInformation
    }

    else
    {
        $array_Comp.Values | Export-Csv -Path $fd.filename –NoTypeInformation
    }

    [System.Windows.Forms.MessageBox]::Show("Export CSV completed", "Status")

}

function LoadOSs_Comp()
{
    foreach ($child in $listMatching_Comp.Values)
    {
        try
        {
            if ($child.OperatingSystem -eq $null)
            {
                $listOS_Comp.add("No OS","No OS")
            }
            else
            {
                $listOS_Comp.add($child.OperatingSystem,$child.OperatingSystem)
            }
        }
        catch
        {
            continue
        }
        [System.Windows.Forms.Application]::DoEvents()
    }

    foreach ($child in $listOS_Comp.Values)
    {
        $OSlist_Comp.Items.Add($child)
    }
    
    $OSlist_Comp.Sorted = $true    
}

function Display_Selections_Comp()
{
    $script:array_Comp = @{}
    foreach ( $child in $listMatching_Comp.Values )
    {
        foreach ($item in $OSlist_Comp.SelectedItems)
        {
           if ($child.OperatingSystem -eq $item)
           {
                $array_Comp.Add($child.CanonicalName,$child)
                break
           }
           elseif (($child.OperatingSystem -eq $null) -and ($item -eq "No OS"))
           {
                $array_Comp.Add($child.CanonicalName,$child)
                break
           }
        }
        [System.Windows.Forms.Application]::DoEvents()
    }

    $Selected_Comp.Text = $array_Comp.Count.ToString()
    $array_Comp.Values | Out-GridView
}

function Validate_Comp()
{
    $removal_List = @{}

    foreach ($child in $listMatching_Comp.Values)
    {
        Write-Host "Validating -" $child.Name
        
        if ($child.DNSHostName -ne $null)
        {
            try
            {
                [System.Net.Dns]::Resolve($child.DNSHostName)
                if ((Test-Connection -ComputerName $child.DNSHostName -Quiet -Count 2 -TimeToLive 5 ))
                {
                    Write-Host "Adding -" $child.Name "to removal list"
                    $removal_List.add($child.CanonicalName,$child.CanonicalName)
                    continue
                }
            }

            catch
            {
               continue 
            }
        }
        else
        {
            try
            {
                [System.Net.Dns]::Resolve($child.Name)
                if ((Test-Connection -ComputerName $child.Name -Quiet -Count 2 -TimeToLive 5 ))
                {
                    Write-Host "Adding -" $child.Name "to removal list"
                    $removal_List.add($child.CanonicalName,$child.CanonicalName)
                    continue
                }
            }

            catch
            {
                continue
            }
        }
    }

    foreach ($child in $removal_List.Values)
    {
        Write-Host "Removing -" $child
        $listMatching_Comp.Remove($child)
    }
}

function Filters_Comp()
{
    $removal_List = @{}
    Write-Host "Applying Filters"
    if ($ModifiedDate_DrpText_Comp.Enabled)
    {
        switch ($ModifiedDate_DrpText_Comp.SelectedIndex) 
        { 
            0 {$LastModifiedDate = 180} 
            1 {$LastModifiedDate = 90} 
            2 {$LastModifiedDate = 60} 
            3 {$LastModifiedDate = 45} 
            4 {$LastModifiedDate = 30} 
            default {$LastModifiedDate = 180}
        }
        
        foreach ($child in $listMatching_Comp.Values)
        {
            if (($script:date - $child.modifyTimeStamp).Days -le $LastModifiedDate)
            {
                Write-Host "Adding -" $child.Name "to removal list"
                $removal_List.add($child.CanonicalName,$child.CanonicalName)
            }
        }
    }

    foreach ($child in $removal_List.Values)
    {
        Write-Host "Removing -" + $child
        $listMatching_Comp.Remove($child)
    }
}


# End Computer Object Functions
###################################################################


UserToolsTab
ObjectAging
UserObjectsTab
Init_Sys_Usr
CompObjectsTab
Init_Sys_Comp

[void]$Form1.ShowDialog()