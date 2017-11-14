# CreateSharedCalendar
# Version 0.5 by Ian Hutchinson
# This had some server names and company specific details hardcoded into it. These are replaced with dollar signs $ for uploading to github

function doInitialise {
    Import-Module ActiveDirectory
    
    #used for password generation
    [Reflection.Assembly]::LoadWithPartialName("System.Web")
    
    #initiate Exchange connection

    $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $$$$$$$$$$$$ -Authentication Kerberos
    Import-PSSession $exchangeSession
    
    clear-host
    write-host "Create Shared Calendar Script"
    write-host "Version 0.5 by Ian Hutchinson 08/03/16"
    write-host "This version will create calendars but will not change permissions."

    #pick a mail DB randomly. The company where this was deployed had 25 different mail DBs named sequentially.

    $calendarMailDBnumber = get-random -Minimum 01 -Maximum 25
    $numDigits = $calendarMailDBnumber.ToString().Length
    $calendarMailDB = "$$$$$$$" -replace ".{$numDigits}$"
    $calendarMailDB += $calendarMailDBnumber
    
    #randomly generate a mailbox password. This had to comply with the company password policy.
    
    $passwordOK  = $false
    while ($passwordOK -eq $false) {
        $candidatePassword = [System.Web.Security.Membership]::GeneratePassword(12,2)
        $passwordOK = $true
        if ($candidatePassword -cnotmatch '[0-9]') {$passwordOK = $false}
        if ($candidatePassword -cnotmatch '[a-z]') {$passwordOK = $false}
        if ($candidatePassword -cnotmatch '[A-Z]') {$passwordOK = $false}
        $candidatePassword = $candidatePassword | ConvertTo-SecureString -AsPlainText -Force
    }
    


    $script:CalendarParameters = @{
        "Name" = ""
        "Alias" = ""
        "UserPrincipalName" = ""
        "OrganizationalUnit" = "$$$$$$$$$$$$$$$$$$"
        "Database" = $calendarMailDB
        "Password" = $candidatePassword
    }
    $script:CalendarAdditionalParameters = @{
        "Manager" = ""
        "Private" = $false
        "Notes" = ""
        "Shared" = $false
        "Room" = $false
        "Equipment" = $false
    }
}

function doStartTranscript {
    #Calendar name must be set before calling this function

    $script:transcriptOutputDirectory = $PSScriptRoot + "\logs\" + $calendarParameters.Alias + "-" + (get-date -format "yyyyMMdd_hh-mm-ss") + ".txt"
    Start-Transcript -path $script:transcriptOutputDirectory
    $script:transcriptStarted = $true
}

function doGetCalendarName {
    write-host "`nRemember to respect the naming conventions for room calendar names."
    $calendarNameValid = $false
    while ($calendarNameValid -eq $false) {
        $calendarParameters.Name = read-host "Please enter the name of the new shared calendar"
        try{$mailboxNameTest = Get-Mailbox $calendarParameters.Name -errorAction SilentlyContinue} catch {}
        if ($MailboxNameTest) {
            write-host "Name already taken" -ForegroundColor Red
            write-host "Please try again"
        } else {
            write-host "Name is free!" -ForegroundColor Cyan
            $calendarNameValid = $true
        }
    }
    write-host "Calendar display name will be "$calendarParameters.Name -ForegroundColor Yellow
    
    $calendarAliasValid = $false
    while ($calendarAliasValid -eq $false) {
        $calendarParameters.Alias = read-host "Please enter the desired calendar alias"
    
        write-host "Now checking that the calendar alias is valid, adjusting as necessary:"
        #remove spaces from DL alias, replace with underscores
        $calendarParameters.Alias = $calendarParameters.Alias -replace ' ', '_'
        #replace illegal characters with asterisks
        $calendarParameters.Alias = $calendarParameters.Alias -ireplace '[^a-z0-9!#$%&*+-/=?_{|}~]', "*"
        #cut the string down to the maximum length, if it exceeds this
        if ($calendarParameters.Alias.Length -ge 64) {write-host "Alias is longer than the maximum length for aliases, cutting this down.`n`n"}
        $calendarParameters.Alias = $calendarParameters.Alias.Substring(0,[math]::min(63,$calendarParameters.Alias.Length))
        try{$mailboxAliasTest = Get-Mailbox $calendarParameters.Name -errorAction SilentlyContinue} catch {}
        if ($MailboxAliasTest) {
            write-host "Alias already taken" -ForegroundColor Red
            write-host "Please try again"
        } else {
            write-host "Alias is free!" -ForegroundColor Cyan
            $calendarAliasValid = $true
        }
    }
    write-host "Calendar alias will be "$calendarParameters.Alias -ForegroundColor Yellow
    $calendarParameters.UserPrincipalName = $calendarParameters.Alias + "@$$$$$$$$$.com"
    write-host "UserPrincipalName will be "$calendarParameters.UserPrincipalName -ForegroundColor Yellow

}

function doGetCalendarManager {
    $managerSelected = $false
    $managerNameInput = ''
    $managerNameInput = read-host "`nPlease enter the name of the calendar owner"
    $managerNameInput += "*"
    $managerSearchResults = @(Get-ADUser -Filter{name -like $managerNameInput -and enabled -eq $true} -Properties company, department)}

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
                    $calendarAdditionalParameters.manager = $managerSearchResults[([int]$userResponse - 1)]
                    $userResponseisValid = $true
                    $managerSelected = $true
                }
            }
            $userResponse = $null
        }
    }
    
    $managerSelected
}

function doGetPublicOrPrivate {
    $userResponseisValid = $false
    $userResponse = $null
    
    while ($userResponseisValid -eq $false) {

        $userResponse = read-host "Will this be a private calendar? (y/n)"

        if ($userResponse -ne $null) {
            if($userResponse -eq 'n') {
                $userResponseisValid = $true
                $calendarAdditionalParameters.private = $false
                write-host "Calendar will be made publicly visible" -ForegroundColor Yellow
            }
            if($userResponse -eq 'y') {
                $userResponseisValid = $true
                $calendarAdditionalParameters.private = $true
                write-host "Calendar will be made made private" -ForegroundColor Yellow
            }
        }
    }
}

function doGetCalendarResourceType {
    $userResponseIsValid = $false
    $userResponse = $null
    while ($userResponseIsValid -eq $false) {
        write-host "`nPlease select the calendar resource type."
        write-host "Select 'r' for room, 'e' for equipment, 's' for Shared"
        $userResponse = read-host
        if ($userResponse -ne $null) {
            switch ($userResponse) {
                'r' {
                    $userResponseIsValid = $true
                    $script:calendarAdditionalParameters.Room = $true
                    $script:calendarAdditionalParameters.Equipment= $true
                    $script:calendarAdditionalParameters.Shared= $true
                    write-host "Calendar will be a Room resource" -ForegroundColor Yellow
                }
                'e' {
                    $userResponseIsValid = $true
                    $script:calendarAdditionalParameters.Room= $true
                    $script:calendarAdditionalParameters.Equipment= $true
                    $script:calendarAdditionalParameters.Shared= $true
                    write-host "Calendar will be a Equipment resource" -ForegroundColor Yellow
                }
                's' {
                    $userResponseIsValid = $true
                    $script:calendarAdditionalParameters.Room= $true
                    $script:calendarAdditionalParameters.Equipment= $true
                    $script:calendarAdditionalParameters.Shared= $true
                    write-host "Calendar will be a Shared resource" -ForegroundColor Yellow
                }
                default {
                    $userResponseIsValid = $false
                    $script:calendarAdditionalParameters.Room= $true
                    $script:calendarAdditionalParameters.Equipment= $true
                    $script:calendarAdditionalParameters.Shared= $true
                    write-host "Inavlid entry." -ForegroundColor Yellow
                }
            }
        }
    }
}

function doGetCalendarNotes {
    $userResponseIsValid = $false
    $userResponse = $null
    while ($userResponseIsValid -eq $false) {
        write-host "`nPlease enter any relevant notes for this calendar."
        write-host "These will be copied onto the object in exchange."
        write-host "e.g. put in the Request or Task number"
        $userResponse = read-host
        if ($userResponse -ne $null) {
            $userResponseIsValid = $true
            $calendarAdditionalParameters.Notes = $userResponse
            Write-Host "`nCalendar notes will be:`n" $userResponse
        }
    }
}

function doCreateCalendar {
    write-host "Creating the calendar mailbox" -ForegroundColor Green
    if ($calendarAdditionalParameters.Room -eq $true) {
        $mailbox = New-Mailbox @calendarParameters -Room
    }
    if ($calendarAdditionalParameters.Equipment -eq $true) {
        $mailbox = New-Mailbox @calendarParameters -Equipment
    }
    if ($calendarAdditionalParameters.Shared -eq $true) {
        $mailbox = New-Mailbox @calendarParameters -Shared
    }
    #set manager
    $mailbox | set-user -manager $calendarAdditionalParameters.Manager

    #set notes
    $mailbox | set-user -Notes $calendarAdditionalParameters.Notes

    #default is public, set private if this is required
    if ($calendarAdditionalParameters.Private -eq $true) {
        $query = $mailbox.alias + ":\calendar"
        set-MailboxFolderPermission $query -user default -accessrights none
    }
}

function doExit {
    #terminate session
    $openPSSessions = Get-PSSession
    foreach ($session in $openPSSessions) {Remove-PSSession $session.Id}

    write-host "`nExiting" -foreground "RED"
    if ($script:transcriptStarted -eq $true) {Stop-Transcript}
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

    $userWantsToContinue # This is needed for the function to work. Puts the value into the pipeline which is then returned by the function
}

function doSummary {
    write-host "`nHere are the calendar details selected:"
    write-host "Calendar name: "$calendarParameters.Name
    write-host "Calendar alias: "$calendarParameters.Alias
    Write-host "UserPrincipalName: "$calendarParameters.UserPrincipalName
    write-host "Calendar Manager: "$calendarAdditionalParameters.manager
    write-host "room: " $calendaAdditionalParameters.Room
    write-host "equipment: "$calendarAdditionalParameters.Equipment
    write-host "shared: "$calendarAdditionalParameters.Shared
    write-host "notes: "$calendarAdditionalParameters.Notes
    if ($calendarAdditionalParameters.private) {write-host "Calendar will be private."} else {Write-Host "calendar will be publicly visible"}

    Write-Host 'Press any key to continue...'
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
}

doInitialise

$calendarNamesAccepted = $false
$calendarManagerAccepted = $false

while ($calendarNamesAccepted -eq $false) {
    doGetCalendarName
    write-host "Do you want to continue with this calendar name and alias?"
    $calendarNamesAccepted = doPromptToContinue
}
while ($calendarManagerAccepted -eq $false) {
    if(doGetCalendarManager) {
        write-host "`nManager selected is " $calendarAdditionalParameters.manager -ForegroundColor Yellow
        write-host "Do you want to continue with this selected calendar manager?"
        $calendarManagerAccepted = doPromptToContinue
    }
}
doGetPublicOrPrivate
doGetCalendarResourceType
doGetCalendarNotes

doStartTranscript
doSummary
doCreateCalendar
doExit
