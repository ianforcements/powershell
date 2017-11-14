#Set Mailbox sent items configuration
#by Ian Hutchinson 19/11/2015
#
#For a given mailbox name, sets the sent items configuration.

#import Exchange snap-in
#mailserver = ""
Get-PSSnapin | where {$_.Name -eq "microsoft.exchange.management.powershell.e2010"}
add-pssnapin "microsoft.exchange.management.powershell.e2010" -ErrorVariable errSnapin ;. $env:ExchangeInstallPath\bin\RemoteExchange-mod.ps1
Connect-ExchangeServer -Server $mailserver -allowclobber

#initialise variables
$mailboxID = ""
$mailboxSentItemConfigState = ""
$mailboxdata = ""
$sendAsChangeRequired = ""
$sendAsSenderAndFromRequired = ""
$sendOnBehalfChangeRequired = ""
$sendOnBehalfChangeRequired = ""

#Intro and Request user to provide name of mailbox
write-host "Set Mailbox Sent Items Configuration"
write-host "Script by Ian Hutchinson, 19/11/2015"
write-host "************************************`n"
write-host "This script will alter the behaviour of Exchange when an item is sent as a shared mailbox using 'send-as' or 'send-on-behalf'."
write-host "Under these circumstances the following behaviours are possible:"
write-host "1. Sender: Place the sent item in the 'Sent Items' folder belonging to the person who sent the email"
write-host "2. SenderAndFrom: Place the sent item in the 'Sent Items' folders of both the person who sent the email, and the mailbox that they sent-as"
write-host "It is not possible to set the behaviour to place the item in the sent-as mailbox only."
write-host "`nThese behaviours can be set independently for 'send-as' and 'send-on-behalf' behaviour.`n"
write-host "Please enter the alias of the target mailbox: " -NoNewline
$mailboxID = read-host

#Get the data
$mailboxSentItemConfigState = Get-MailboxSentItemsConfiguration $mailboxID
$mailboxData = get-Mailbox $mailboxID
#Display the data
write-host "`n***Selected mailbox details***"
write-host "Name: " $mailboxData.Name
write-host "Alias: " $mailboxdata.Alias
write-host "Send-As behaviour: " $mailboxSentItemConfigState.SendAsItemsCopiedTo
write-host "Send-On-Behalf behaviour: " $mailboxSentItemConfigState.SendOnBehalfOfItemsCopiedTo
write-host ""
#Prompt to check if changes to Send-As behaviour are required.
do{
	write-host "Do you wish to make changes to the Send-As behaviour?: (Y/N) " -nonewline
	$sendAsChangeRequired = read-host
}
until ($sendAsChangeRequired -eq "Y" -or $sendAsChangeRequired -eq "N")

if($sendAsChangeRequired -eq "Y") {
    do{
        write-host "Do you require SenderAndFrom to be enabled for Send-As?: (Y/N) " -nonewline
		$sendAsSenderAndFromRequired = read-host
	}
	until ($sendAsSenderAndFromRequired -eq "Y" -or $sendAsSenderAndFromRequired -eq "N")
}

do{
	write-host "Do you wish to make changes to the Send-On-Behalf behaviour?: (Y/N) " -nonewline
	$sendOnBehalfChangeRequired = read-host
}
until ($sendOnBehalfChangeRequired -eq "Y" -or $sendOnBehalfChangeRequired -eq "N")

if($sendOnBehalfChangeRequired -eq "Y") {
    do{
        write-host "Do you require SenderAndFrom to be enabled for Send-On-Behalf?: (Y/N) " -nonewline
		$sendOnBehalfSenderAndFromRequired = read-host
	}
	until ($sendOnBehalfSenderAndFromRequired -eq "Y" -or $sendOnBehalfSenderAndFromRequired -eq "N")
}

write-host ""

if($mailboxSentItemConfigState.SendAsItemsCopiedTo -eq "Sender" -and $sendAsSenderAndFromRequired -eq "Y") {
    write-host "Setting Send-As configuration on " $mailboxData.Name " to SenderAndFrom"
    Set-MailboxSentItemsConfiguration $mailboxdata.Name -SendAsItemsCopiedTo SenderAndFrom
} elseif($mailboxSentItemConfigState.SendAsItemsCopiedTo -eq "SenderAndFrom" -and $sendAsSenderAndFromRequired -eq "N") {
    write-host "Setting Send-As configuration on " $mailboxData.Name " to Sender Only"
    Set-MailboxSentItemsConfiguration $mailboxdata.Name -SendAsItemsCopiedTo Sender
} else {
    write-host "No changes made to Send-As behaviour"
}

if($mailboxSentItemConfigState.SendOnBehalfOfItemsCopiedTo -eq "Sender" -and $sendOnBehalfSenderAndFromRequired -eq "Y") {
    write-host "Setting Send-On-Behalf configuration on " $mailboxData.Name " to SenderAndFrom"
    Set-MailboxSentItemsConfiguration $mailboxdata.Name -SendOnBehalfOfItemsCopiedTo SenderAndFrom
} elseif($mailboxSentItemConfigState.SendOnBehalfOfItemsCopiedTo -eq "SenderAndFrom" -and $sendOnBehalfSenderAndFromRequired -eq "N") {
    write-host "Setting Send-On-Behalf configuration on " $mailboxData.Name " to Sender Only"
    Set-MailboxSentItemsConfiguration $mailboxdata.Name -SendOnBehalfOfItemsCopiedTo Sender
} else {
write-host "No changes made to Send-On-Behalf behaviour"
}

write-host "`n`nCompleted.`n"
#Get the data again
$mailboxSentItemConfigState = Get-MailboxSentItemsConfiguration $mailboxID
$mailboxData = get-Mailbox $mailboxID
#Display the data
write-host "***Current Mailbox State***"
write-host "Name: " $mailboxData.Name
write-host "Alias: " $mailboxdata.Alias
write-host "Send-As behaviour: " $mailboxSentItemConfigState.SendAsItemsCopiedTo
write-host "Send-On-Behalf behaviour: " $mailboxSentItemConfigState.SendOnBehalfOfItemsCopiedTo
write-host ""

forEach ($session in $PSSessions) {
Remove-PSSession $session
}