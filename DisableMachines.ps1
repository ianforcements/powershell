#Disable machines on Alinta domain
#by Ian Hutchinson 30/09/2015
#Machine will take a CSV file containing machine names and disable them, then moved them to the disabled OU
# This had some server names or company specific details hardcoded into it. These are removed for sharing online

import-module activedirectory
$disabledOU = "OU=Disabled,OU=Workstations,OU=Client,DC=companyA,DC=net,DC=int"
$disabledDescription = "Disabled per $$$$$$$$$$"
$computerlist = import-csv ".\ComputersToBeDisabled.csv"
$computerlist | foreach-object {
$computerInstance = get-adcomputer $_.cn
set-adcomputer $_.cn -Description $disabledDescription -Enabled 0
move-adobject -Identity $computerInstance.objectGUID.GUID -TargetPath $disabledOU
$_ | Add-Member -NotePropertyName 'Disabled' -NotePropertyValue 'TRUE'
}

write-host "Completed."
$computerlist | export-csv ".\DisabledMachines.csv"
