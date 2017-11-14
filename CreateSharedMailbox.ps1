# CreateSharedMailbox
# Version 0.5 by Ian Hutchinson

#set sent items behaviour by default to sender and mailbox
#when deployed this script had a few hardcoded values specific to the organisation. These have been replaced with $$$$$ and should be adjusted accordingly
function doInitialise {
    Import-Module ActiveDirectory
    
    #used for password generation
    [Reflection.Assembly]::LoadWithPartialName("System.Web")
    
    #initiate Exchange connection

    $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $$$$$$$$$$$$$$$$$$$$$ -Authentication Kerberos
    Import-PSSession $exchangeSession
    
    clear-host
    write-host "Create Shared mailbox Script"
    write-host "Version 0.5 by Ian Hutchinson 08/03/16"
    write-host "This version will create mailboxes but will not change permissions."

    #pick a mail DB randomly. The environment in which this was deployed had a large range of mail and archive DBs and new mailboxes were intended to be allotted randomly

    $mailboxMailDBnumber = get-random -Minimum 01 -Maximum $$
    $numDigits = $mailboxMailDBnumber.ToString().Length
    $mailboxMailDB = "$$$$$$$$$" -replace ".{$numDigits}$"
    $mailboxMailDB += $mailboxMailDBnumber
    
    #pick an archive DB randomly
    $mailboxArchiveDBNumber = get-random -Minimum 01 -Maximum $$
    $numDigits = $mailboxArchiveDBNumber.ToString().Length
    $mailboxArchiveDB = "$$$$$$$" -replace ".{$numDigits}$"
    $mailboxArchiveDB += $mailboxArchiveDBNumber

    #randomly generate a mailbox password. The password needed to match the company password policy, which is the cause of the limitations you see below
    
    $passwordOK  = $false
    while ($passwordOK -eq $false) {
        $candidatePassword = [System.Web.Security.Membership]::GeneratePassword(12,2)
        $passwordOK = $true
        if ($candidatePassword -cnotmatch '[0-9]') {$passwordOK = $false}
        if ($candidatePassword -cnotmatch '[a-z]') {$passwordOK = $false}
        if ($candidatePassword -cnotmatch '[A-Z]') {$passwordOK = $false}
        $candidatePassword = $candidatePassword | ConvertTo-SecureString -AsPlainText -Force
    }
    


    $script:mailboxParameters = @{
        "Name" = ""
        "Alias" = ""
        "UserPrincipalName" = ""
        "OrganizationalUnit" = "OU=Resource Mailboxes,OU=Client,DC=alinta,DC=net,DC=int"
        "Database" = $mailboxMailDB
        "ArchiveDatabase" = $mailboxArchiveDB
        "Password" = $candidatePassword
    }
    $script:mailboxAdditionalParameters = @{
        "Manager" = ""
        "Notes" = ""
    }
    
}

function doStartTranscript {
    #mailbox name must be set before calling this function

    $script:transcriptOutputDirectory = $PSScriptRoot + "\logs\" + $mailboxParameters.Alias + "-" + (get-date -format "yyyyMMdd_hh-mm-ss") + ".txt"
    Start-Transcript -path $script:transcriptOutputDirectory
    $script:transcriptStarted = $true
}

function doGetmailboxName {
    write-host "`nRemember to respect the naming conventions for room mailbox names."
    $mailboxNameValid = $false
    while ($mailboxNameValid -eq $false) {
        $mailboxParameters.Name = read-host "Please enter the name of the new shared mailbox"
        try{$mailboxNameTest = Get-Mailbox $mailboxParameters.Name -errorAction SilentlyContinue} catch {}
        if ($MailboxNameTest) {
            write-host "Name already taken" -ForegroundColor Red
            write-host "Please try again"
        } else {
            write-host "Name is free!" -ForegroundColor Cyan
            $mailboxNameValid = $true
        }
    }
    write-host "mailbox display name will be "$mailboxParameters.Name -ForegroundColor Yellow
    
    $mailboxAliasValid = $false
    while ($mailboxAliasValid -eq $false) {
        $mailboxParameters.Alias = read-host "Please enter the desired mailbox alias"
    
        write-host "Now checking that the mailbox alias is valid, adjusting as necessary:"
        #remove spaces from alias, replace with underscores
        $mailboxParameters.Alias = $mailboxParameters.Alias -replace ' ', '_'
        #replace illegal characters with asterisks
        $mailboxParameters.Alias = $mailboxParameters.Alias -ireplace '[^a-z0-9!#$%&*+-/=?_{|}~]', "*"
        #cut the string down to the maximum length, if it exceeds this
        if ($mailboxParameters.Alias.Length -ge 64) {write-host "Alias is longer than the maximum length for aliases, cutting this down.`n`n"}
        $mailboxParameters.Alias = $mailboxParameters.Alias.Substring(0,[math]::min(63,$mailboxParameters.Alias.Length))
        try{$mailboxAliasTest = Get-Mailbox $mailboxParameters.Name -errorAction SilentlyContinue} catch {}
        if ($MailboxAliasTest) {
            write-host "Alias already taken" -ForegroundColor Red
            write-host "Please try again"
        } else {
            write-host "Alias is free!" -ForegroundColor Cyan
            $mailboxAliasValid = $true
        }
    }
    write-host "mailbox alias will be "$mailboxParameters.Alias -ForegroundColor Yellow
    $mailboxParameters.UserPrincipalName = $mailboxParameters.Alias + "@$$$$$.com"
    write-host "UserPrincipalName will be "$mailboxParameters.UserPrincipalName -ForegroundColor Yellow

}

function doGetMailboxManager {
    $managerSelected = $false
    $managerNameInput = ''
    $managerNameInput = read-host "`nPlease enter the name of the mailbox owner"
    $managerNameInput += "*"

    $managerSearchResults = @(Get-ADUser -Filter{name -like $managerNameInput -and enabled -eq $true} -Properties company, department)

    if ($managerSearchResults.Count -eq 0) {
        write-host "No results found for search." -ForegroundColor Red
    } else {
        write-host "`nList of search results:"
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
            write-host "`nPlease select the manager by typing the number of the selection.`nOr press r to retry, x to exit:" -foreground "yellow"
            $userResponse = read-host
            if ($userResponse -ne $null) {
                if($userResponse -eq 'x') {doExit}
                if($userResponse -eq 'r') {
                    $userResponseisValid = $true
                } elseif(([int]$userResponse -gt 0) -and ([int]$userResponse -le $managerSearchResults.count)) {
                    $mailboxAdditionalParameters.manager = $managerSearchResults[([int]$userResponse - 1)]
                    $userResponseisValid = $true
                    $managerSelected = $true
                }
            }
            $userResponse = $null
        }
    }
    
    $managerSelected
}

function doGetmailboxNotes {
    $userResponseIsValid = $false
    $userResponse = $null
    while ($userResponseIsValid -eq $false) {
        write-host "`nPlease enter any relevant notes for this mailbox."
        write-host "These will be copied onto the object in exchange."
        write-host "e.g. put in the Request or Task number"
        $userResponse = read-host
        if ($userResponse -ne $null) {
            $userResponseIsValid = $true
            $mailboxAdditionalParameters.Notes = $userResponse
            Write-Host "`nmailbox notes will be:`n" $userResponse
        }
    }
}

function doCreatemailbox {
    write-host "Creating the mailbox" -ForegroundColor Green
    $mailbox = New-Mailbox @mailboxParameters -Shared
    
    #set manager
    $mailbox | set-user -manager $mailboxAdditionalParameters.Manager

    #set notes
    $mailbox | set-user -Notes $mailboxAdditionalParameters.Notes

}

function doExit {
    #terminate session
    $openPSSessions = Get-PSSession
    foreach ($session in $openPSSessions) {Remove-PSSession $session.Id}

    write-host "`nExiting" -foreground "RED"
    Stop-Transcript
    Write-Host -NoNewLine 'Press any key to continue...'
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    clear
    exit
}

function doPromptToContinue {
    $userResponseisValid = $false
    $userResponse = $null
    $userWantsToContinue = $false
    while ($userResponseisValid -eq $false) {

        $userResponse = read-host "y for yes, n for no, x to exit script"

        if ($userResponse -ne $null) {
            if($userResponse -eq 'x') {doExit}
            if($userResponse -eq 'n') {
                $userResponseisValid = $true
                $userWantsToContinue = $false
                write-host "Ok, let's try again`n" -ForegroundColor Yellow
            }
            if($userResponse -eq 'y') {
                $userResponseisValid = $true
                $userWantsToContinue = $true
                write-host "Ok, moving on.`n" -ForegroundColor Yellow
            }
        }
    }

    $userWantsToContinue
}

function doSummary {
    write-host "`nHere are the mailbox details selected:"
    write-host "mailbox name: "$mailboxParameters.Name
    write-host "mailbox alias: "$mailboxParameters.Alias
    Write-host "UserPrincipalName: "$mailboxParameters.UserPrincipalName
    write-host "mailbox Manager: "$mailboxAdditionalParameters.manager
    write-host "notes: "$mailboxAdditionalParameters.Notes
    
    Write-Host 'Press any key to continue...'
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
}

doInitialise

$mailboxNamesAccepted = $false
$mailboxManagerAccepted = $false

while ($mailboxNamesAccepted -eq $false) {
    doGetmailboxName
    write-host "Do you want to continue with this mailbox name and alias?"
    $mailboxNamesAccepted = doPromptToContinue
}
while ($mailboxManagerAccepted -eq $false) {
    if(doGetmailboxManager) {
        write-host "`nManager selected is " $mailboxAdditionalParameters.manager -ForegroundColor Yellow
        write-host "Do you want to continue with this selected mailbox manager?"
        $mailboxManagerAccepted = doPromptToContinue
    }
}
doGetmailboxNotes

doStartTranscript
doSummary
doCreatemailbox
doExit
