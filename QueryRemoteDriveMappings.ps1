###Query remote mapped drives v1.0
#For a given username on a given remote machine, displays all current drive mappings
##By Ian Hutchinson 1/10/2015

## to do: make it so that the script authenticates to CompanyB domain  to query CompanyB devices
#set up some things
Import-Module ActiveDirectory
$outFilePath = ".\remoteDriveMappings.txt"
$driveMappingArray = @()
$longestFilePath = 0

#talk at you and get input
clear
write-host "Query Remote Drive Mappings v1.0"
write-host "by Ian Hutchinson 1/10/2015"
write-host "********************************"
write-host "This script will contact a remote computer and query its registry for all`n network drive mappings set by a specified user of that machine."
write-host "The remote machine must be online and connected to the network."
write-host "The specified user must have the network drives mapped on the machine, but they need not be logged on."
write-host "********************************`n"

write-host "Please Enter Username: " -nonewline
$UID = read-host
write-host "Please Enter Computer name: " -nonewline
$computer = read-host
write-host ""
#validate user and grab the SID
$userData = get-aduser $UID
if ($? -eq $False){
	write-host "Username" $UID "not found. Exiting." -foreground "YELLOW"
	Write-Host 'Press any key to continue...'
	$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
	exit
} else {write-host "Username" $UID "found."}

#I tried to make the script contact the CompanyB machines directly
#However this fails as it appears our CompanyA accounts are not permitted for remote registry on their machines
#This section warns to run this script on CompanyB domain for CompanyB users
if ($userData.DistinguishedName -like "*OU=Linked CompanyB Users*"){
write-host "User appears to be a CompanyB user - this script may need to be run from CompanyB domain!" -foreground "RED"
write-host "A copy of this script is located at C:\Admin\Scripts on the CompanyB jump host. " -foreground "RED"
write-host "A planned future version of this script will not have this limitation." -foreground "RED"
}

#set up remote query, make sure remote machine is responding, then get the list of mapped drives
$regKey = $userData.SID.value + "\Network"
if (Test-Connection -ComputerName $Computer -Count 1 -Quiet) {
	#If remote machine is talking, say so, then get the data
	write-host "Remote PC" $Computer "found."
	$remoteRegistry = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('Users',$computer)
	$openKey = $remoteRegistry.OpenSubKey($regKey)
	$openKey.getSubKeyNames() | foreach-object {
		$openDriveMapKey = $openKey.OpenSubKey($_)
		$driveMapPath = $openDriveMapKey.getvalue('RemotePath')
		$driveMapping = new-object PSObject
		$driveMapping | add-member -membertype NoteProperty -name 'DriveLetter' -value $_
		$driveMapping | add-member -membertype NoteProperty -name 'Path' -value $driveMapPath
		$driveMappingArray += $driveMapping
		if($longestDrivePath -lt $driveMapping.Path.Length) {$longestDrivePath = $driveMapping.Path.Length}
		}
	
}
else {
	#If machine isn't talking, say so.
	write-host "Remote computer " $Computer "not contactable! Exiting."
	Write-Host 'Press any key to continue...'
	$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
	exit
}
#We have the data, now do some formatting things.
#we want to make sure the file path isn't truncated by the console width
#having done this, spit it out to text file and display to user.
$dataWidth = $longestDrivePath + 12
if ($dataWidth -lt 80) {$dataWidth = 80}
$driveMappingArray | Format-table -autosize | out-file $outFilePath -width $dataWidth -force 
write-host "`nResults written to" $outFilePath
write-host "This file will now be opened.`n"
Write-Host 'Press any key to continue...'
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
notepad.exe $outFilePath
exit
