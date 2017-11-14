#BulkGetEmailAddresses
#by Ian Hutchinson 23/11/2015
#For a csv list of user display names, grab the email addresses
#takes an input: userDisplayNames.csv
# This had some server names and company specific details hardcoded into it. These are removed for sharing online

#initialise some stuff
$userDisplayNameList = import-csv ".\userDisplayNames.csv"
$outputFileName = ".\output.csv"
$foundUsers = 0
$notfoundUsers = 0

#import Exchange snap-in
$server = ""
Get-PSSnapin | where {$_.Name -eq "microsoft.exchange.management.powershell.e2010"}
add-pssnapin "microsoft.exchange.management.powershell.e2010" -ErrorVariable errSnapin ;. $env:ExchangeInstallPath\bin\RemoteExchange-mod.ps1
Connect-ExchangeServer -Server $server -allowclobber

$userDisplayNameList | foreach-object {
    $searchTerm = $_.DisplayName
    $data = get-mailbox -identity $searchTerm
    if ($? -eq $False) {
        write-host $_.DisplayName "not found."
        $_ | Add-Member -NotePropertyName 'ADPresent' -NotePropertyValue $False -force
		$notFoundUsers += 1
    }
    else {
		write-host $_.userName "Email is:" $data.PrimarySMTPAddress.toString()
		$_ | Add-Member -NotePropertyName 'ADPresent' -NotePropertyValue $True -force
		$_ | Add-Member -NotePropertyName 'EmailAddress' -NotePropertyValue $data.PrimarySMTPAddress.toString() -force
        $foundUsers += 1
	}
}


$userDisplayNameList | export-csv $outputFileName
write-host "Completed"
write-host "data written to " $outputFileName
write-host "Found users: " $foundUsers
write-host "Not found users: " $notfoundUsers

forEach ($session in $PSSessions) {
Remove-PSSession $session
}