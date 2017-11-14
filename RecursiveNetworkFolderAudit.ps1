#RecursiveNetworkfolderAudit
#
#version 0.5 by Ian Hutchinson 4/7/16
#
#this script starts at a network folder specified by yourself and grabs every subfolder, recursing until no subfolders remain.
#for each folder this script will grab the ACL for that folder and list each entry on that ACL
#it will then spit out this list of folders and ACL entries
#For each ACL entry that is an AD group, it will run a recursive listing of all members of that group
#This was created as an immediate solution to a specific problem and may need to be modified for different circumstances.
#It can take a very long time indeed for large folders. It is not recommended for auditing very large network shares in its current state.
#It also has problems with very long paths.

#TODO

#for each folder, check that its ACL entries are different to the parent folder
#if not, NBD, remove entries
#if so, retain them
#for each ACL entry remaining, get the group members via AD and put them in a CSV in a folder
#try to zip them if possible?
#produce some text that can be copied and pasted into an email reply
#currently fails for very long paths
#potential fix here:
# http://powershell.com/cs/forums/p/12697/22651.aspx


Import-Module ActiveDirectory

$initialSearchPath = ""
$totalFoldersCount = 0
$folderList = ""
$totalACLData = @()
$uniqueACE = @()

function doEnumerateFolders() {
write-host "Searching for subfolders. This may take a very long time."
write-host "You'll have to wait it out."
$global:folderlist = get-childitem -Recurse -Directory $initialSearchPath
$global:totalFoldersCount = $folderlist.count
write-host "Located $global:totalFoldersCount folders and subfolders"
write-host "The script will now check the ACL on each folder.`n`n"

}

function doListACLentries() {
    forEach($folder in $global:folderList) {
        write-host ('FOLDER:' + $folder.fullname)
        $global:folderACLData = get-acl $folder.fullname
        forEach($ACE in $folderACLData.access) {
            if($global:uniqueACE -notContains $ACE.identityReference) {$global:uniqueACE += $ACE.IdentityReference}
            write-host "`tACE: " $ACE.identityReference "Has" $ACE.AccessControlType "rights:" $ACE.FileSystemRights
        }
    }            

}


write-host "Recursive Network Folder Audit" -ForegroundColor Yellow
write-host "version 0.5 by Ian Hutchinson 04/07/2016`n"

write-host "Please enter the starting path." -ForegroundColor Yellow
$initialSearchPath = read-host "Starting path"
write-host "Please enter a short name to identify this report" -ForegroundColor Yellow
$reportName = read-host "Report name"
$scriptDate = get-date -Format "yyyyMMdd_hh-mm-ss"
$logPath = "$PSScriptRoot\logs\$reportName-$scriptdate.txt"
write-host "Starting Transcript..."
Start-Transcript -Path $logPath
if($?) {} else {
    write-host "Error starting transcript!" -ForegroundColor Red
    write-host "Output will not be logged!`n`n" -ForegroundColor Red
write-host "`nTesting that this location is reachable..." -NoNewline
if(test-path $initialSearchPath) {
    write-host "OK!" -ForegroundColor Green
    doEnumerateFolders
    doListACLentries
    
} else {
    write-host "Not OK!" -ForegroundColor Red
    write-host "Unable to locate folder. Exiting"
}

Stop-Transcript
