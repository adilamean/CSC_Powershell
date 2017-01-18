param([int]$FreeSpace, [string]$AppName, [string]$RegHive, [string]$Region, [string]$ApplyOUExcl);


$instance="CSCDBSNDC003"
$database='invdb'
$userid='inventory'
$password='inventory'
$connectionString="User ID=$userid;Password=$password;Initial Catalog=$database;Data Source=$instance"

$target ="select distinct upper(inv_systems.name) Name from inv_systems
	where DateDiff(day,inv_systems.scantime,getdate()) < 30"



$lowdsk ="select distinct inv_systems.name Name from inv_systems
	where 
	inv_systems.free_space <= "+$FreeSpace+"
	and DateDiff(day,inv_systems.scantime,getdate()) < 30"


$excl ="select distinct inv_systems.name from inv_systems, inv_machinegroups
	where upper(inv_machinegroups.[group]) like 'APPL-%-BPR-XPT-"+$AppName+"%'
	and inv_systems.name in
	("+$target+")
	and inv_systems.name not in
	("+$ouexcl+" union all "+$lowdsk+")
	and inv_machinegroups.machine_id=inv_systems.id
	and DateDiff(day,inv_systems.scantime,getdate()) < 30"



$notin ="select distinct inv_systems.name from inv_systems
	where not exists (
	   select * from inv_registryentry 
	   where upper(inv_registryentry.registry_key) like '\HKEY_LOCAL_MACHINE\SOFTWARE%\CSC\PACKAGES\"+$RegHive+"'
	   and upper(inv_registryentry.attribute_name) = 'INSTALLED'
	   and upper(inv_registryentry.attribute_value) = 'YES'
	   and inv_registryentry.machine_id = inv_systems.id
	)"

$instl="select distinct inv_systems.name from inv_systems
	where inv_systems.name not in
	("+$notin+")"



$connection = New-Object System.Data.SqlClient.SqlConnection
$connection.ConnectionString = $connectionString
#$connection.ConnectionString = "Server=$dataSource;Database=$database;Integrated Security=True;"
$connection.Open()
$command = $connection.CreateCommand()


### Calculate Target ###

$command.CommandText = $target
$command.CommandTimeout = 0

$targetcount = 0

$result = $command.ExecuteReader()

$table = new-object “System.Data.DataTable”
$table.Load($result)

$table | Select @{Name="Name"; Expression={$_.Name.ToUpper()}} | Export-CSV .\"$RegHive"-Target.txt  -NoTypeInformation -Encoding UTF8


#Declaring Join Object

. .\Join-Object.ps1



#Calculating XPT Exclusions


Remove-Item .\"$RegHive"-XPTExcl.txt -Force

$ADGroupData = Import-CSV .\ADGroupData.csv


$XPTGroup = "Appl-U-bPr-Xpt-" +$AppName + "-MST"


$grp = Get-ADGroup -Filter {Name -eq $XPTGroup}  -Server "CSCDC8BLY001.amer.globalcsc.net" -SearchBase "OU=Global Apps,OU=ChannelMasterGroups,OU=BMC CM,OU=BMC Software,DC=amer,DC=globalcsc,DC=net" -Properties Members

$grp.members | Out-File .\"$RegHive"-XPTExcl.txt -Encoding UTF8 -Append


ForEach ($item in $ADGroupData)

{

$XPTGroup = "Appl-G-bPr-Xpt-" +$AppName + "-"+$item.GroupName 


$grp = Get-ADGroup -Filter {Name -eq $XPTGroup}  -Server $item.ADServer -SearchBase $item.GroupLocation -Properties Members

$grp.members | Out-File .\"$RegHive"-XPTExcl.txt -Encoding UTF8 -Append

}

$XPTExcl = Import-CSV .\"$RegHive"-XPTExcl.txt -Header DistinguishedName | Select -Property @{label = 'Name';expression= {$_.DistinguishedName -replace '^CN=|,.*$'}}


#Comparing the machines in bPower with AD BUInfo So that we include Managed Workstations only


$BuInfo = Import-CSV "D:\AMI Reports\BUInfo.csv"
$Target = Import-CSV .\"$RegHive"-Target.txt

Join-Object –Left $Target –Right $BuInfo -LeftJoinProperty Name -RightJoinProperty Name –LeftProperties Name -Type OnlyIfInBoth | Select-Object "Name"| Export-CSV .\"$RegHive"-Target2.txt -NoTypeInformation -Encoding UTF8

$total = (Import-CSV .\"$RegHive"-Target2.txt | Measure-Object)

$targetstring = $AppName+",Count"

Write-Output $targetstring | out-file .\"$RegHive"-Report.txt

$targetstring = "Target Machines,"+$total.count

Write-Output $targetstring | out-file .\"$RegHive"-Report.txt -Append


### Calculate XPT Exclusions ###

$Target2 = Import-CSV .\"$RegHive"-Target2.txt

Join-Object –Left $Target2 –Right $XPTExcl -LeftJoinProperty Name -RightJoinProperty Name –LeftProperties Name -Type OnlyIfInBoth | Select-Object "Name"| Export-CSV .\"$RegHive"-XPTExcl2.txt -NoTypeInformation -Encoding UTF8

$xptexclcount = (Import-CSV .\"$RegHive"-XPTExcl2.txt | Measure-Object).Count

$targetstring = "XPT Exclusions,"+$xptexclcount

Write-Output $targetstring | out-file .\"$RegHive"-Report.txt -Append


$Target3 = Compare-Object –ReferenceObject $Target2 –DifferenceObject $XPTExcl -Property Name | Where-Object {$_.SideIndicator -ne "=>"} | Select-Object "Name"
$Target3.count

### Calculate OU Exclusions ###


Import-CSV "D:\AMI Reports\BUInfo.csv"| Where-Object {$_.MainSite -eq "IS&S"} | Select-Object "Name" | Export-CSV .\"$RegHive"-OUExcl.txt -NoTypeInformation -Encoding UTF8

$OUExcl = Import-CSV .\"$RegHive"-OUExcl.txt | Select Name

Join-Object –Left $Target3 –Right $OUExcl -LeftJoinProperty Name -RightJoinProperty Name –LeftProperties Name -Type OnlyIfInBoth | Select-Object "Name"| Export-CSV .\"$RegHive"-OUExcl2.txt -NoTypeInformation -Encoding UTF8

$ouexclcount = (Import-CSV .\"$RegHive"-OUExcl2.txt | Measure-Object).Count

$targetstring = "OU Exclusions,"+$ouexclcount

Write-Output $targetstring | out-file .\"$RegHive"-Report.txt -Append


$Target4 = Compare-Object –ReferenceObject $Target3 –DifferenceObject $OUExcl -Property Name | Where-Object {$_.SideIndicator -ne "=>"} | Select-Object "Name"
$Target4.count




## Calculate Low Disk Space ###

$command.CommandText = $lowdsk
$command.CommandTimeout = 0

$targetcount = 0

$result = $command.ExecuteReader()

$table = new-object “System.Data.DataTable”
$table.Load($result)

$table | Select Name | Export-CSV .\"$RegHive"-FreeSpace.txt  -NoTypeInformation -Encoding UTF8

$LessSpc = Import-CSV .\"$RegHive"-FreeSpace.txt

Join-Object –Left $Target4 –Right $LessSpc -LeftJoinProperty Name -RightJoinProperty Name –LeftProperties Name -Type OnlyIfInBoth | Select-Object "Name"| Export-CSV .\"$RegHive"-FreeSpace2.txt -NoTypeInformation -Encoding UTF8

$lessspace = (Import-CSV .\"$RegHive"-FreeSpace2.txt | Measure-Object).Count

$targetstring = "Low Disk,"+$lessspace


Write-Output $targetstring | out-file .\"$RegHive"-Report.txt -Append

$Target5 = Compare-Object –ReferenceObject $Target4 –DifferenceObject $LessSpc -Property Name | Where-Object {$_.SideIndicator -ne "=>"} | Select-Object "Name" 


### Calculate Adjusted Target ###


$Target5 | Export-CSV .\"$RegHive"-AdjTarget.txt -NoTypeInformation -Encoding UTF8

$targetstring = "Adjusted Target,"+$Target5.count

Write-Output $targetstring | out-file .\"$RegHive"-Report.txt -Append


### Calculate Not Installed ###

$command.CommandText = $notin
$command.CommandTimeout = 0

$targetcount = 0

$result = $command.ExecuteReader()

$table = new-object “System.Data.DataTable”
$table.Load($result)

$table | Select Name | Export-CSV .\"$RegHive"-NotInstalled.txt  -NoTypeInformation -Encoding UTF8

$NotInstall = Import-CSV .\"$RegHive"-NotInstalled.txt

Join-Object –Left $Target5 –Right $NotInstall -LeftJoinProperty Name -RightJoinProperty Name –LeftProperties Name -Type OnlyIfInBoth | Select-Object "Name"| Export-CSV .\"$RegHive"-NotInstalled2.txt -NoTypeInformation -Encoding UTF8

$notinstalled = (Import-CSV .\"$RegHive"-NotInstalled2.txt | Measure-Object).Count

$targetstring = "Not Installed,"+$notinstalled

Write-Output $targetstring | out-file .\"$RegHive"-Report.txt -Append

$Target6 = Compare-Object –ReferenceObject $Target5 –DifferenceObject $NotInstall -Property Name | Where-Object {$_.SideIndicator -ne "=>"} | Select-Object "Name" 


### Calculate Installed Machines ###

$Target6 | Export-CSV .\"$RegHive"-Installed.txt -NoTypeInformation -Encoding UTF8


$targetstring = "Installed,"+$Target6.count

Write-Output $targetstring | out-file .\"$RegHive"-Report.txt -Append


$percentage =  [decimal]::round((($Target6.count / $Target5.count) * 100),2)

$targetstring = "Install Percentage,"+$percentage+"%"

Write-Output $targetstring | out-file .\"$RegHive"-Report.txt -Append


$connection.Close()

####### Mail Alert ##########

$SmtpClient = new-object system.net.mail.smtpClient -ArgumentList "20.137.2.88"

$MailMessage = New-Object system.net.mail.mailmessage

$mailmessage.from = "edm_csci@csc.com"

$mailmessage.CC.add("arathore7@csc.com")
$mailmessage.To.add("jmongia@csc.com")
$mailmessage.To.add("cmathew2@csc.com")
$mailmessage.To.add("vvenkatesh4@csc.com")
$mailmessage.Cc.add("akumar333@csc.com")
$mailmessage.Cc.add("mjoshi23@csc.com")


$AdjTarget = "D:\Package Reports\"+$RegHive+"-AdjTarget.txt"

$Installed = "D:\Package Reports\"+$RegHive+"-Installed.txt"

$NotInstall = "D:\Package Reports\"+$RegHive+"-NotInstalled.txt"

$attachment1 = New-Object System.Net.Mail.Attachment($AdjTarget, 'text/plain')

$attachment2 = New-Object System.Net.Mail.Attachment($Installed, 'text/plain')

$attachment3 = New-Object System.Net.Mail.Attachment($NotInstall, 'text/plain')

$mailmessage.Subject = ""+$RegHive+" Package Report - "+$Region+""

$mailmessage.IsBodyHTML = $true

$style = "< style>BODY{font-family, Arial; font-size, 10pt;}"
$style = $style + "TABLE{border, 1px solid black; border-collapse, collapse;}"
$style = $style + "TH{border, 1px solid black; background, #dddddd; padding, 5px; }"
$style = $style + "TD{border, 1px solid black; padding, 5px; }"
$style = $style + "< /style>"


$body = Import-CSV .\"$RegHive"-Report.txt| Select-Object $AppName,Count  | ConvertTo-Html 

$mailmessage.Body = $body

$mailmessage.Attachments.Add($attachment1)
$mailmessage.Attachments.Add($attachment2)
$mailmessage.Attachments.Add($attachment3)

#$smtpclient.Send($mailmessage) 