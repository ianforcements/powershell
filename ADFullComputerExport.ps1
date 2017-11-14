#AD Computer Export
#By Ian Hutchinson 25/09/2015

$ADComputerData = get-adcomputer -filter * -properties *

#find container OU for each object
$ADComputerData | ForEach-Object {
	$ObjectCanonicalNameElements = $_.CanonicalName -split '/'
	$ObjectJoinedCanonicalName = $ObjectCanonicalNameelements[0..($ObjectCanonicalNameElements.Count - 2)] -join '/'
	$_ | Add-Member -NotePropertyName 'ParentOU' -NotePropertyValue $ObjectJoinedCanonicalName -force
}
#filter out required data
$ADComputerData = $ADComputerData | select name,description,distinguishedname,dnshostname,enabled,lastlogondate,modified,objectclass,objectguid,operatingsystem,samaccountname,sid,userprincipalname,whenchanged,whencreated,parentou

#export to CSV
$ADComputerData | export-csv ".\ADFullComputerExport.csv"



