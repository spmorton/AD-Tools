
# Global Variables
$date = Get-Date
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

#. .\User-Object-Tool.ps1

$ADTVersion = 1.1

# form objects
$Form1 = New-Object System.Windows.Forms.Form 
$Tabcontrol1 =  New-Object System.Windows.Forms.TabControl
$userObjTab = New-Object System.Windows.Forms.TabPage
$computerObjTab = New-Object System.Windows.Forms.TabPage
$Server = New-Object System.Windows.Forms.TextBox
$Server_Label = New-Object System.Windows.Forms.Label
$CredsButton = New-Object System.Windows.Forms.Button
$CurrentCreds_Check = New-Object System.Windows.Forms.CheckBox

# User Tab Objects
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


$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState

# form specs
$Form1.Text = "AD Tools - " + $ADTVersion
$Form1.Name = "adtools"
$Form1.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 725
$System_Drawing_Size.Height = 750
$Form1.ClientSize = $System_Drawing_Size

# tab control specs
$Tabcontrol1.Name = "tabControl"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 75
$System_Drawing_Point.Y = 85
$Tabcontrol1.Location = $System_Drawing_Point
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 600
$System_Drawing_Size.Width = 575
$Tabcontrol1.Size = $System_Drawing_Size
$Form1.Controls.Add($Tabcontrol1)


$userObjTab.AutoSize = $true
$userObjTab.TabIndex = 0
$userObjTab.Text = "User Objects"
$userObjTab.Enabled = $true
$Tabcontrol1.Controls.Add($userObjTab)

$computerObjTab.AutoSize = $true
$computerObjTab.TabIndex = 1
$computerObjTab.Text = "Computer Objects"
$computerObjTab.Enabled = $false
$Tabcontrol1.Controls.Add($computerObjTab)

$Server.Location = New-Object System.Drawing.Size(80,35)
$Server.Size = New-Object System.Drawing.Size(270,25)
$Server.Text = ""
$Form1.Controls.Add($Server)

$Server_Label.Location = New-Object System.Drawing.Size(80,16) 
$Server_Label.Size = New-Object System.Drawing.Size(270,20) 
$Server_Label.Text = "Server Name or IP address to query"
$Form1.Controls.Add($Server_Label) 

$CredsButton.Location = New-Object System.Drawing.Size(360,35)
$CredsButton.Size = New-Object System.Drawing.Size(100,20)
$CredsButton.Text = "Get Credentials"
$CredsButton.Enabled = $true
$CredsButton.Add_Click({
    $script:creds = Get-Credential
    $ScanButton_Usr.Enabled = $true
    })
$Form1.Controls.Add($CredsButton)

$CurrentCreds_Check.Location = New-Object System.Drawing.Size(470,35)
$CurrentCreds_Check.Size = New-Object System.Drawing.Size(120,20)
$CurrentCreds_Check.Text = "Use Current Creds"
$CurrentCreds_Check.Add_CheckStateChanged({
    
    if ($CurrentCreds_Check.Checked)
    {
        $CredsButton.Enabled = $false
        $ScanButton_Usr.Enabled = $true
    }
    else
    {
        $CredsButton.Enabled = $true
        $ScanButton_Usr.Enabled = $false
    }
})
$Form1.Controls.Add($CurrentCreds_Check)



Function UserTabObjects()
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
                                Scan_Usr
 
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

    $ModifyButton_Usr.Location = New-Object System.Drawing.Size(10,375)
    $ModifyButton_Usr.Size = New-Object System.Drawing.Size(140,25)
    $ModifyButton_Usr.Text = "Perform Operation"
    $ModifyButton_Usr.Enabled = $false
    $ModifyButton_Usr.Add_Click({Perform_Operation_Usr})
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

###################################################################
# Begin User Object Functions

Function Init_Sys_Usr()
{
    $filters_Usr = $false


    $array_Usr = @()            # used for selected and displayed data
    $listMatching_Usr = @{}
    $failures_Usr = @{}
    $ScanButton_Usr.Enabled = $false
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
                if (($date - $_.modifyTimeStamp).Days -ge $LastModifiedDate -AND $_.Enabled -eq $false)
                {
                    $_ | Add-Member -MemberType NoteProperty -Name "Days Since Last Mod" -Value ($date - $_.modifyTimeStamp).Days -Force
                    $listMatching_Usr.add($_.CanonicalName,$_) # Use CanonicalName to capture duplicate entries
                }
            }

            else
            {
                if (($date - ([datetime]::FromFileTime($_.lastLogonTimeStamp))).Days -ge $daysOld -AND $_.Enabled -eq $true)
                {
                    if ($_.pwdLastSet -ne $null)
                    {
                        $_ | Add-Member -MemberType NoteProperty -Name "Pwd Age" -Value ($date - ([datetime]::FromFileTime($_.pwdLastSet))).Days -Force
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
                if (($date - $_.modifyTimeStamp).Days -ge $LastModifiedDate -AND $_.Enabled -eq $false)
                {
                    $_ | Add-Member -MemberType NoteProperty -Name "Days Since Last Mod" -Value ($date - $_.modifyTimeStamp).Days -Force
                    $listMatching_Usr.add($_.CanonicalName,$_) # Use CanonicalName to capture duplicate entries
                }
            }

            else
            {
                if (($date - ([datetime]::FromFileTime($_.lastLogonTimeStamp))).Days -ge $daysOld -AND $_.Enabled -eq $true)
                {
                    if ($_.pwdLastSet -ne $null)
                    {
                        $_ | Add-Member -MemberType NoteProperty -Name "Pwd Age" -Value ($date - ([datetime]::FromFileTime($_.pwdLastSet))).Days -Force
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
            [System.Windows.Forms.Application]::DoEvents()
        }
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

# End Computer Object Functions
###################################################################



UserTabObjects
Init_Sys_Usr

[void]$Form1.ShowDialog()