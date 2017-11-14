# Bulk get user SID
# By Ian Hutchinson 7/10/15

# For an input CSV containing usernames, retrieves the SIDs associated with those usernames. 

$inputData = import-csv ".\userNameList.csv"
$outputData = ".\userNameListWithSIDS.csv"
$foundUsers = 0
$notfoundUsers = 0

$inputData | forEach-object {
	$ADUserInstance = get-aduser $_.userName
	if ($? -eq $False) {
		write-host $_.userName "not found."
		$_ | Add-Member -NotePropertyName 'ADPresent' -NotePropertyValue $False -force
		$notFoundUsers += 1
	}
	else {
		write-host $_.userName "SID is:" $ADUserInstance.SID
		$_ | Add-Member -NotePropertyName 'ADPresent' -NotePropertyValue $True -force
		$_ | Add-Member -NotePropertyName 'SID' -NotePropertyValue $ADUserInstance.SID -force
		$_ | Add-Member -NotePropertyName 'DisplayName' -NotePropertyValue $AdUserInstance.Name -force
		$foundUsers += 1
	}
}
$inputData | export-csv $outputData
write-host "Completed."
write-host $foundUsers "SIDs found." $notFoundUsers "SIDs not found."
write-host "Results written to" $outputdata
