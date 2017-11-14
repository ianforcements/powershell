#Inactive AD Computer Export
#By Ian Hutchinson 25/09/2015

#Searches AD for computer objects that have not logged on for greater than $inactiveDays


$inactiveDays = 134
$ADComputerData = get-adcomputer -filter * -properties * | Where-Object{$_.LastLogonDate -le (get-date).adddays(-$inactiveDays)}

#find container OU for each object
$ADComputerData | ForEach-Object {
	$ObjectCanonicalNameElements = $_.CanonicalName -split '/'
	$ObjectJoinedCanonicalName = $ObjectCanonicalNameelements[0..($ObjectCanonicalNameElements.Count - 2)] -join '/'
	$_ | Add-Member -NotePropertyName 'ParentOU' -NotePropertyValue $ObjectJoinedCanonicalName -force
}
#filter out required data
$ADComputerData = $ADComputerData | select ParentOU,cn,Enabled,LastLogonDate,Description

#export to CSV
$ADComputerData | export-csv ".\ADComputerExport134days.csv"



