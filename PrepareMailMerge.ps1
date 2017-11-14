#Prepare Mail Merge
#Created 29/09/2015 by Ian Hutchinson
#When given a list of PCs, this will search for a username in the 'description' field
#If it can match this to an active user, it will add the user to a mail merge list
#intended for automating emails sent to owners of computers related to inactive computer objects in AD
# This had some server names or company specific details hardcoded into it. These are removed for sharing online

import-module activedirectory

$inactiveComputerData = import-csv ".\InactiveComputerAuditReport.csv"
$inactiveOtherData = import-csv ".\OtherComputerAuditReport.csv"
$mailMerge = @()

$inactiveComputerData | foreach-object {
	#attempt to pull user from AD
	$ADOwner = get-aduser -identity $_.Description
	#if this is successful, and user is enabled, add data to mail merge
	if ($? -eq $True){
		if($ADOwner.Enabled) {
			$mailMergeLine = new-object PSObject
			$mailMergeLine | Add-member -membertype NoteProperty -name 'Asset' -value $_.cn -force
			$mailMergeLine | Add-member -membertype NoteProperty -name 'sAMAccountName' -value $ADOwner.SamAccountName -force
			$mailMergeLine | Add-member -membertype NoteProperty -name 'GivenName' -value $ADOwner.GivenName -force
			$mailMergeLine | Add-member -membertype NoteProperty -name 'Surname' -value $ADOwner.Surname -force
			$mailMergeLine | Add-member -membertype NoteProperty -name 'UserPrincipalName' -value $ADOwner.UserPrincipalName -force
			$mailMerge += $MailMergeLine
		}
	}
}

$inactiveOtherData | foreach-object {
	#attempt to pull user from AD
	$ADOwner = get-aduser -identity $_.Description
	#if this is successful, and user is enabled, add data to mail merge
	if ($? -eq $True){
		if($ADOwner.Enabled) {
			$mailMergeLine = new-object PSObject
			$mailMergeLine | Add-member -membertype NoteProperty -name 'Asset' -value $_.cn -force
			$mailMergeLine | Add-member -membertype NoteProperty -name 'sAMAccountName' -value $ADOwner.SamAccountName -force
			$mailMergeLine | Add-member -membertype NoteProperty -name 'GivenName' -value $ADOwner.GivenName -force
			$mailMergeLine | Add-member -membertype NoteProperty -name 'Surname' -value $ADOwner.Surname -force
			$mailMergeLine | Add-member -membertype NoteProperty -name 'UserPrincipalName' -value $ADOwner.UserPrincipalName -force
			$mailMerge += $MailMergeLine
		}
	}
}

$mailMerge | export-csv ".\MailMerge.csv"
