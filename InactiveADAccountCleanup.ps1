# Inactive AD Account Cleanup
# By Ian Hutchinson 22/12/2015

# This script will search AD for users who have not logged on for the last 180 days.
# it will check to confirm if the accounts have been created recently and exclude these from the process
# Any remaining user accounts are disabled and relocated to the inactive OU

# It is worth noting that the AD attribute for last logon is not guaranteed to be updated at every logon all the time
# This is according to Microsoft design which seeks to reduce load on domain controllers
# Lastlogondate data may be up to 14 days out of date by design.
# This means that in practice the script is cleaning users whose last logon is 180-194 days ago.
# Microsoft blog article for more info:
# http://blogs.technet.com/b/askds/archive/2009/04/15/the-lastlogontimestamp-attribute-what-it-was-designed-for-and-how-it-works.aspx

write-host "Inactive AD account cleanup v0.1"
write-host "by Ian Hutchinson 22/12/2015`n"
write-host "This script will disable and relocate AD user objects which have not been logged on to in over 180 days.`n"
write-host "+++ NOTES +++"
write-host "There are some complications regarding the accuracy of last logon times in AD"
write-host "as a result account information may be up to 14 days out of date."
write-host "to prevent unnecessary account disablement the cutoff date is set to 194 days in the past."
write-host "Additionally, the last logon time is only updated by certain authentication events"
write-host "this script examines the lastLogonTimeStamp attribute to determine stale accounts.`n`n"

#Initialise some stuff
import-module activedirectory
$inactiveDays = 60
$nominalInactiveDays = $inactiveDays -14
$inactiveCutoffDate = (get-date).AddDays(-$inactiveDays).ToFileTime()
$exclusionReasons = @("Not Excluded","Recently Created", "Whitelisted") #whitelisting / exclusion list is currently not implemented
$companyAOutFileName = ".\CompanyA"+$nominalInactiveDays+"DaysInactiveADUsers_" + (get-date -format "yyyyMMdd_hh-mm-ss") + ".csv"
$companyBOutFileName = ".\CompanyB"+$nominalInactiveDays+"DaysInactiveADUsers_" + (get-date -format "yyyyMMdd_hh-mm-ss") + ".csv"
$companyAInactiveUsers = @()
$companyBInactiveUsers = @()
$companyATargetOUs = @("OU=CompanyA,OU=Users,OU=Client,DC=companyA,DC=net,DC=int") #note that this script will automatically search subtrees
$companyBTargetOUs = @("OU=Standard Users,OU=Users,OU=CompanyB,DC=com,DC=au") #note that this script will automatically search subtrees
$companyAInactiveOU = "OU=Inactive,OU=Users,OU=Client,DC=companyA,DC=net,DC=int"
$companyBInactiveOU = "OU=Inactive Users,OU=Users,OU=CompanyB,DC=com,DC=au"
$companyBDomainController = 'serverName.companyB.com.au'

#get CompanyB credentials
write-host "You will be prompted for CompanyB credentials"
do {
    $companyBCredentialIsValid = $false
    $companyBCredentials = $host.ui.PromptForCredential("CompanyB Credentials", "Please enter your CompanyB Admin user name and password.", "","")
    write-host "`nTesting CompanyB Credentials"
    get-aduser administrator -server $companyBDomainController -credential $companyBCredentials | out-null
    if ($? -eq $true) {
        write-host "Success!"
        $companyBCredentialIsValid = $true
    } else {
        do {
            write-host "failed. (r)etry or e(x)it?"
            $usrInput = read-host
        } until (($usrInput -eq 'r') -or ($usrInput -eq 'x'))
        if ($usrInput -eq 'x') {exit}
    }
} while ($companyBCredentialIsValid -eq $false)


#for each target OU, get a list of stale users

#start with CompanyA
write-host "Examining accounts in CompanyA domain"

forEach ($targetOU in $companyATargetOUs) {
    $OUInactiveUsers = Get-ADUser -SearchBase $targetOU -SearchScope Subtree -Filter {LastLogonTimeStamp -le $inactiveCutoffDate} -properties LastLogonTimeStamp, CanonicalName, whenCreated, title, distinguishedName
    
    
    #now loop through the list of users we have found
    forEach ($inactiveUser in $OUInactiveUsers) {        
        #exclude users who suit exclusion criteria
        if($inactiveUser.whenCreated.ToFileTime() -ge $inactiveCutoffDate) { #check that the user was not created within the cutoff date
        $inactiveUser | Add-Member -NotePropertyName 'Excluded' -NotePropertyValue $exclusionReasons[1] -force #recently created
        } else {
        $inactiveUser | Add-Member -NotePropertyName 'Excluded' -NotePropertyValue $exclusionReasons[0] -force #not excluded
        }

        #Find Parent OU & put the last logon timestamp in human-readable format
        $ObjectCanonicalNameElements = $inactiveUser.CanonicalName -split '/'
        $ObjectCanonicalName = $ObjectCanonicalNameElements[0..($ObjectCanonicalNameElements.Count -2)] -join '/'
        $inactiveUser | Add-Member -NotePropertyName 'ParentOU' -NotePropertyValue $ObjectCanonicalName -force
        $lastLogonTimeStampHumanReadable = [datetime]::FromFileTime($inactiveUser.lastLogonTimeStamp)
        $inactiveUser | Add-Member -NotePropertyName 'LastLogon' -NotePropertyValue $lastLogonTimeStampHumanReadable -force
        
    }
    $companyAInactiveUsers += $OUInactiveUsers | select Surname, GivenName, Enabled, LastLogon, Title, UserPrincipalName, SAMAccountName, ParentOU, Excluded, whenCreated, distinguishedName
}

$companyAInactiveUsers | export-csv $companyAOutFileName -NoTypeInformation -Force
write-host $companyAInactiveUsers.count " inactive users found in CompanyA domain."
write-host ($companyAInactiveUsers | Where-Object {$_.Excluded -eq $exclusionReasons[0]}).count " of these will be disabled."
Write-Host "CompanyA list of inactive users saved to " $companyAOutFileName


write-host "Examining accounts in CompanyB domain"
forEach ($targetOU in $companyBTargetOUs) {
    $OUInactiveUsers = Get-ADUser -SearchBase $targetOU -SearchScope Subtree -Filter {LastLogonTimeStamp -le $inactiveCutoffDate} -properties LastLogonTimeStamp, CanonicalName, whenCreated, title, distinguishedName -server $companyBDomainController -Credential $companyBCredentials
    
    
    #now loop through the list of users we have found
    forEach ($inactiveUser in $OUInactiveUsers) {        
        #exclude users who suit exclusion criteria
        if($inactiveUser.whenCreated.ToFileTime() -ge $inactiveCutoffDate) { #check that the user was not created within the cutoff date
        $inactiveUser | Add-Member -NotePropertyName 'Excluded' -NotePropertyValue $exclusionReasons[1] -force #recently created
        } else {
        $inactiveUser | Add-Member -NotePropertyName 'Excluded' -NotePropertyValue $exclusionReasons[0] -force #not excluded
        }

        #Find Parent OU & put the last logon timestamp in human-readable format
        $ObjectCanonicalNameElements = $inactiveUser.CanonicalName -split '/'
        $ObjectCanonicalName = $ObjectCanonicalNameElements[0..($ObjectCanonicalNameElements.Count -2)] -join '/'
        $inactiveUser | Add-Member -NotePropertyName 'ParentOU' -NotePropertyValue $ObjectCanonicalName -force
        $lastLogonTimeStampHumanReadable = [datetime]::FromFileTime($inactiveUser.lastLogonTimeStamp)
        $inactiveUser | Add-Member -NotePropertyName 'LastLogon' -NotePropertyValue $lastLogonTimeStampHumanReadable -force
        
    }
    $companyBInactiveUsers += $OUInactiveUsers | select Surname, GivenName, Enabled, LastLogon, Title, UserPrincipalName, SAMAccountName, ParentOU, Excluded, whenCreated, distinguishedName
}

$companyBInactiveUsers | export-csv $companyBOutFileName -NoTypeInformation -Force
write-host $companyBInactiveUsers.count " inactive users found in CompanyA domain."
write-host ($companyBInactiveUsers | Where-Object {$_.Excluded -eq $exclusionReasons[0]}).count " of these will be disabled."
Write-Host "CompanyB list of inactive users saved to " $companyBOutFileName

write-host "`n`n+++NOTE+++" -ForegroundColor RED
write-host "Please check the output files to confirm the list of users to be disabled and relocated." -ForegroundColor RED
write-host "Users can be excluded from this process by editing the 'Excluded' column to any value other than " $exclusionReasons[0]
write-host "Be sure to save any changes before continuing."

do {
    write-host "`nTo continue, type " -NoNewline; write-host "'continue'. " -f Yellow -NoNewline; write-host "To exit, press " -NoNewline; write-host "x" -f Yellow
    $usrInput = read-host
} until (($usrInput -eq 'continue') -or ($usrInput -eq 'x'))
if ($usrInput -eq 'x') {exit}


$companyAUsersToDisable = import-csv $companyAOutFileName | Where-Object{ $_.Excluded -eq $exclusionReasons[0]}
$companyBUsersToDisable = import-csv $companyBOutFileName | Where-Object{ $_.Excluded -eq $exclusionReasons[0]}
$disabledUsers = 0
$movedUsers = 0

forEach ($user in $companyAUsersToDisable) {
    #these lines are commented out until I get to go-ahead to deploy the script
    #confirm best practice for moving user - might be best to get-aduser then pipe this to move-adobject
    #also create log file output of all changes
    #move-adobject -identity $user.distinguishedname -TargetPath $companyAInactiveOU
    #Disable-ADAccount -Identity $user.distinguishedName
}
forEach ($user in $companyBUsersToDisable) {
    #these lines are commented out until I get to go-ahead to deploy the script
    #confirm best practice for moving user - might be best to get-aduser then pipe this to move-adobject
    #also create log file output of all changes
    #move-adobject -identity $user.distinguishedname -TargetPath $companyBInactiveOU -Server $companyBDomainController -Credential $companyBCredentials
    #Disable-ADAccount -Identity $user.distinguishedName -Server $companyBDomainController -Credential $companyBCredentials
}

#to do: close out script and announce output to user, save log files, return to script menu, handle any errors that appeared
# write test cases, have service account created with rights only to required OUs, 
# - pipe to select, then flag first or last 10
# next step to make change in subOU in companyA
