# EditCalendarPermission
# Version 1.0 by Ian Hutchinson 31/03/2016

# For a given user and mailbox alias, removes or adds or changes access
# This had some server names and company specific details hardcoded into it. These are removed for sharing online

$server = ""
Get-PSSnapin | where {$_.Name -eq "microsoft.exchange.management.powershell.e2010"}
add-pssnapin "microsoft.exchange.management.powershell.e2010" -ErrorVariable errSnapin ;. $env:ExchangeInstallPath\bin\RemoteExchange-mod.ps1
Connect-ExchangeServer -Server $server -allowclobber
Import-Module ActiveDirectory

$permissionLevels = @{
"o" = "Owner";
"a" = "Author";
"e" = "Editor";
"r" = "Reviewer";
}


$mailboxNameValid = $false
while ($mailboxNameValid -eq $false) {
    $targetCalendar = read-host "`nPlease enter the alias of the target calendar"
    try{$mailboxNameTest = Get-Mailbox $targetCalendar -errorAction SilentlyContinue} catch {}
    if ($MailboxNameTest) {
        write-host "Calendar"$mailboxNameTest.Name"found!" -ForegroundColor Cyan
        $mailboxNameValid = $true
    } else {
        write-host "Calendar is not found!" -ForegroundColor red
        write-host "Please try again!"
    }
}

$userNameValid = $false
while ($userNameValid -eq $false) {
    $userToEdit = read-host "Please enter the username of the user to change"
    try{$userNameTest = Get-AdUser $userToEdit -errorAction SilentlyContinue} catch {}
    if ($UserNameTest) {
        write-host "User"$userNameTest.name"found!" -ForegroundColor Cyan
        $userNameValid = $true
    } else {
        write-host "user is not found!" -ForegroundColor red
        write-host "Please try again!"
    }
}

$targetQuery = $targetCalendar + ":\Calendar"
#get the user's access to the mailbox
try{$calendarUserPermissionTest = Get-MailboxFolderPermission -identity $targetQuery -user $userToEdit -errorAction SilentlyContinue} catch {}
if ($calendarUserPermissionTest) {
    write-host "User currently has the below rights to this calendar:"
    write-host $calendarUserPermissionTest.AccessRights -ForegroundColor Yellow
    $permissionLevels.Add("x", "REMOVE ACCESS")
} else {
    write-host "User does not currently have a permission entry on this calendar."
}

write-host "`n`nThe script allows you to set one of the below permission levels:"
write-host "`nKey`t`tPermission Level"
$permissionLevels.GetEnumerator() | ForEach-Object {
    write-host $_.Key "`t`t" $_.Value -ForegroundColor Yellow
}
$selectionIsValid = $false
while($selectionIsValid -eq $false) {
    $permissionSelection = read-host "Please select the permission required"
    if($permissionLevels.$permissionSelection) {$selectionIsValid = $true}
}

$transcriptOutputDirectory = $PSScriptRoot + "\logs\" + $targetCalendar + "-" + "-" + $userToEdit + "_" + (get-date -format "yyyyMMdd_hh-mm-ss") + ".txt"
Start-Transcript -path $script:transcriptOutputDirectory
$transcriptStarted = $true

write-host $permissionLevels.$permissionSelection "access selected for"$userToEdit "on calendar" $targetCalendar -ForegroundColor Green

write-host "`nWriting selected permission..." -NoNewline

if($calendarUserPermissionTest) {
    if($permissionSelection = 'x') {
        Remove-MailboxFolderPermission -Identity $targetQuery -user $userToEdit
    } else {
        set-mailboxFolderPermission -Identity $targetQuery -User $userToEdit -AccessRights $permissionLevels.$permissionSelection
    }
} else {
    Add-MailboxFolderPermission -Identity $targetQuery -User $userToEdit -AccessRights $permissionLevels.$permissionSelection
}

write-host "Done" -ForegroundColor Green
forEach ($session in $PSSessions) {
Remove-PSSession $session
}
Stop-Transcript
write-host "`nPress any key to continue."

$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')


exit

