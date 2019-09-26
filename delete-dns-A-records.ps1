# Scott P. Morton
# 9/26/2019
# Takes an export from AD-Tools computer scan and deletes associated DNS A records if they exist

    [reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
    [reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null

    [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')

    $server = [Microsoft.VisualBasic.Interaction]::InputBox('Target Server?', 'Server')
    $domain = [Microsoft.VisualBasic.Interaction]::InputBox('Domain?(ie. us.chs.net)', 'Domain')

    # Get the file containing the server list
    $fd = New-Object system.windows.forms.openfiledialog
    $fd.showdialog()
    $fd.filename

    # Setup the data
    $array_import = @()
    $array_import = Import-Csv -Path $fd.FileName
    
    foreach ($child in $array_import)
    {
        $x = Get-DnsServerResourceRecord -Name $child.SamAccountName.split("$")[0] -RRType "A" -ComputerName $server -ZoneName $domain -ErrorAction:SilentlyContinue
        if($x)
        {
            write-host "Deleting DNS 'A' record for:" $x.HostName
            Remove-DnsServerResourceRecord -name $x.HostName -ComputerName $server -ZoneName $domain -RRType "A" -Force
        }
    }
    