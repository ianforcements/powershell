# Create Distribution List
# by Ian Hutchinson 08/02/2016
# Script to follow the requirements for creating a new distribution list
# DL refers to 'Distribution List' in this context, and not Domain Local Group as in AD Canon.
# This had some server names and company specific details hardcoded into it. These are removed for sharing online
# the script is *not* complete


#initiate Exchange connection
$mailserver = ""
Get-PSSnapin | where {$_.Name -eq "microsoft.exchange.management.powershell.e2010"}
add-pssnapin "microsoft.exchange.management.powershell.e2010" -ErrorVariable errSnapin ;. $env:ExchangeInstallPath\bin\RemoteExchange-mod.ps1
Connect-ExchangeServer -Server $mailserver -allowclobber

$DLOU = "CompanyA.net.int/Client/Distribution Lists"
$DLmemberList = @()
$DLmanager = ''

#request alias and check that it conforms to requirements
#notes: mailbox alias may have no spaces, max length of 64
$DLnameIsApproved = $false
while($DLnameIsApproved -eq $false) {

    write-host "DLs should have a name of the form 'DL_Name'"
    $DLname = read-host "Please enter the name of the distribution group"
    if ($DLname.StartsWith("dl_") -eq $true) {
        write-host "`nName entered does not have capitals for the prefix. Correcting this."
        $DLname = "DL" + $DLname.Substring(2,$DLname.Length -2)
    }
    if ($DLname.StartsWith("DL_") -ne $true) {
        write-host "Name entered does not start with 'DL_', prepending these characters."
        $DLname = "DL_" + $DLname
    }
    write-host "`nDL name will be $DLname" -ForegroundColor Yellow
    $userResponse = read-host "If this is OK, press y to continue, or any other key to try again"
    if ($userResponse -eq 'y') {$DLnameIsApproved = $true}
    $userResponse = $null
}
write-host "$DLname approved as DL name."
write-host "`nThe mailbox will also have an alias, or nickname."
write-host "The mailbox alias has stricter restrictions than the DL name. It must contain no spaces or illegal characters"
write-host "it has a maximum length of 64 characters. Any spaces will be replaced with underscores."
Write-Host "Any other illegal characters will be replaced with asterisks."

#remove spaces from DL alias, replace with underscores
$DLalias = $DLname -replace ' ', '_'
#replace illegal characters with asterisks
$DLalias = $DLalias -ireplace '[^a-z0-9!#$%&*+-/=?_{|}~]', "*"
#cut the string down to the maximum length, if it exceeds this
if ($DLalias.Length -ge 64) {write-host "DL name is longer than the maximum length for aliases, cutting this down."}
$DLalias = $DLalias.Substring(0,[math]::min(63,$DLalias.Length))
write-host "`nthe DL alias will be: $DLalias" -ForegroundColor Yellow

#DL needs a manager. Request this from user.
$DLmanagerSelected = $false
$managerNameInput = ''
while($DLmanagerSelected -eq $false) {
    while($managerNameInput.length -eq 0) {
        $managerNameInput = read-host "`nPlease enter the name of the mailbox owner"
    }
    $managerNameInput += "*"
    $managerSearchResults = @()
    $managerSearchResults = Get-ADUser -Filter{name -like $managerNameInput -and enabled -eq $true} -Properties company, department
    
    forEach ($searchResult in $managerSearchResults) {
        $searchResultIndex = [array]::IndexOf($managerSearchResults, $searchResult)
        $searchResultIndex++
        write-host $searchResultIndex".`tName: " -NoNewline -ForegroundColor Yellow
        write-host $searchResult.name -NoNewline
        write-host "`tCompany: " -NoNewline -ForegroundColor Yellow
        write-host $searchResult.company -NoNewline
        write-host "`tDepartment: " -NoNewline -ForegroundColor Yellow
        Write-Host $searchResult.department
    }
    
    $userResponseisValid = $false
    while ($userResponseisValid -eq $false) {
        write-host "Please select the manager by typing the number of the selection.`nr to retry, x to exit:" -foreground "yellow"
        $userResponse = read-host
        if ($userResponse -ne $null) {
            if($userResponse -eq 'x') {exit}
            if($userResponse -eq 'r') {$userResponseisValid = $true}
            if(([int]$userResponse -gt 0) -and ([int]$userResponse -le ($managerSearchResults.count +1))) {
                $DLmanager = $managerSearchResults[([int]$userResponse - 1)]
                $userResponseisValid = $true
                $DLmanagerSelected = $true
            }
        }
    $userResponse = $null
    }
}
write-host "Manager selected is " $DLmanager.name

$userResponseisValid = $false
$importListOfUsers = $false
while($userResponseisValid -eq $false) {
    $userResponse = read-host "Do you have a list of members for this DL? (y/n):"
    if ($userResponse -eq 'y') {
        $userResponseisValid = $true
        $importListOfUsers = $true
    }
    if ($userResponse -eq 'n') {$userResponseisValid = $true}
    if ($userResponseisValid -eq $false) {write-host "Invalid entry."}
}

if($importListOfUsers) {
    write-host "Please enter the full filename of a CSV containing a list of DL members."
    write-host "Filename must be located at " $PSScriptRoot
    write-host "See example csv file in script directory for an example of the format required."
    $filename = read-host
    $filename = $PSScriptRoot + $filename
    $usersToImport = import-csv $filename
    if($? -eq $true) {
        forEach ($user in $usersToImport) {
            $userADInstance = get-aduser $user.name -ErrorAction SilentlyContinue
            if($? -eq $true) {
                $DLMemberlist += $userADInstance.name
            } else { write-host $user.name "not found. Check manually."}
        }
    }

}


write-host "Please enter any notes relevant to the creation of this DL."
$DLcreationNotes = read-host


#create group

New-DistributionGroup -Name $DLname -Alias $DLalias -ManagedBy $DLmanager -Members $DLmemberList -Notes $DLcreationNotes -OrganizationalUnit $DLOU -WhatIf

#terminate session
$openPSSessions = Get-PSSession
foreach ($session in $openPSSessions) {Remove-PSSession $session.Id}

