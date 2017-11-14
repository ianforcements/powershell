#GetMailboxPermission
#retrieves mailbox permissions for the specified mailbox
#version 1.0 by Ian Hutchinson 2016/01/15

## to do - check against excluded groups so that it does not list domain servers or service accounts
$mailserver = ""
."C:\Program Files\Microsoft\Exchange Server\V14\Bin\RemoteExchange.ps1"
connect-ExchangeServer -server $mailserver -allowclobber

#introduction, get mailbox alias and grab the data from Exchange
write-host "`n`n`n`n`nGetMailboxPermission" -ForegroundColor Yellow
write-host "Version 1.0 by Ian Hutchinson 15/1/2016"
write-host "This script will grab the permissions that relate to a mailbox object."
write-host "It will exclude any inherited permissions such as exchange admin permissions etc."
write-host ""
write-host "Please specify the alias of the target mailbox:"
$targetMailbox = read-host
write-host ""
$outFileName = ".\Output\" + $targetMailbox + " " + (get-date -format "yyyyMMdd_hh-mm-ss") + ".csv"
$permissionData = Get-MailboxPermission -identity $targetMailbox -ErrorAction stop
if($?) { write-host "Mailbox found!"} else { write-host "An error ocurred, exiting." -ForegroundColor Yellow; exit;}

#prepare the output array, loop through all the permission data we have and extract the valuable bits
$outputData = @()
$outputRowCount = 0
ForEach ($permissionRecord in $permissionData) {
if($permissionRecord.isInherited -eq $false) {
    $ADObject = Get-ADObject -Filter {(ObjectSid -eq $permissionRecord.User.SecurityIdentifier)} -properties SamAccountName
    
    
    if ($?) {
        if($ADObject.ObjectClass -eq 'group'){
            write-host "The group " $permissionRecord.User.RawIdentity "has rights to this folder. Recursing through this folder." -ForegroundColor Cyan
            Get-ADGroupMember $permissionRecord.User.SecurityIdentifier -Recursive | ForEach {
                $outputDataRow = New-Object PSObject
                $outputDataRow | Add-Member -MemberType NoteProperty -name "User" -value $_.Name -Force
                $outputDataRow | Add-Member -MemberType NoteProperty -name "Username" -value $_.SamAccountName -force
                $outputDataRow | Add-Member -MemberType NoteProperty -name "Rights" -value ($permissionRecord.AccessRights -join ',') -force
                $outputDataRow | Add-Member -MemberType NoteProperty -name "Group Membership" -value $ADObject.SamAccountName -force
                $outputDataRow | Add-Member -MemberType NoteProperty -name "Deny" -value $permissionRecord.Deny -force
                $outputData += $OutputDataRow
                $outputRowCount ++
            }
        }
    

        
        if($ADObject.ObjectClass -eq 'user'){
            $outputDataRow = New-Object PSObject
            $outputDataRow | Add-Member -MemberType NoteProperty -name "User" -value $ADObject.Name -Force
            $outputDataRow | Add-Member -MemberType NoteProperty -name "Username" -value $ADObject.SamAccountName -Force
            $outputDataRow | Add-Member -MemberType NoteProperty -name "Rights" -value ($permissionRecord.AccessRights -join ',') -Force
            $outputDataRow | Add-Member -MemberType NoteProperty -name "Group Membership" -value "" -Force
            $outputDataRow | Add-Member -MemberType NoteProperty -name "Deny" -value $permissionRecord.Deny -Force
            $outputData += $OutputDataRow
            $outputRowCount ++
        }

    }
}
}

write-host "Found $outputRowCount non-inherited permission records."

if ($outputRowCount -gt 0) {
    $outputData | export-csv $outFileName -NoTypeInformation -Force
    write-host "Data written to $outFileName"
    write-host "This file will now be opened."
    invoke-item "$outFileName"

} else { write-host "No data to write." }
forEach ($session in $PSSessions) {
Remove-PSSession $session
}
write-host "Press any key to continue."

$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')

Stop-Transcript

exit
