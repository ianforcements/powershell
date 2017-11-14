# Obtain UPN Mismatch List
# By Ian Hutchinson 2/2/2016
# Filters through AD data to determine if any users have incorrectly set UPN
# Any users whose UPN does not match their primary SMTP are flagged and output to CSV

# Introduce
write-host "Obtain UPN Mismatch List"
write-host "by Ian Hutchinson 02/02/2016"
write-host "this will obtain a list of AD users whose UPN does not match their primary SMTP"

# Initialise things
import-module ActiveDirectory
$CompanyAUserData = @()
$CompanyAUPNmismatchData = @()
$CompanyAOutFileName = ".\companyAUPNmismatches" + (get-date -format "yyyyMMdd_hh-mm-ss") + ".csv"

write-host "Importing AD data"
$CompanyAUserData = get-aduser -filter * -properties ProxyAddresses, CanonicalName, whenCreated, title
write-host "Filtering data"
forEach ($user in $CompanyAUserData) {
    $userPrimarySMTP = $null
    $upnMismatch = $null
    forEach ($address in $user.proxyAddresses) {
        if ($address -cmatch "SMTP:") {
            $userPrimarySMTP = $address.substring(5)
        }
    }
    if ($userPrimarySMTP -ne $user.userPrincipalName) {$upnMismatch = $true} else {$upnMismatch = $false}

    if ($upnMismatch) {
        #get OU by splitting Canonicalname and reassembling without name at end
        $userCanonicalNameElements = $user.CanonicalName -split '/'
        $userOU = $userCanonicalNameElements[0..($userCanonicalNameElements.Count -2)] -join '/'
        $user | Add-Member -NotePropertyName 'ParentOU' -NotePropertyValue $userOU -force
        #generate proposed UPN
        $user | Add-Member -NotePropertyName 'ProposedUPN' -NotePropertyValue $userPrimarySMTP -force
        #append entry to mismatch data
        $CompanyAUPNmismatchData += $user | Select Surname, GivenName, Title, UserPrincipalName, ProposedUPN, SAMaccountName, enabled, ParentOU
    }

}

$CompanyAUPNmismatchData | export-csv $CompanyAOutFileName -NoTypeInformation -Force
write-host "Number of UPN mismatches found in companyA domain: " $CompanyAUPNmismatchData.count
write-host "companyA UPN mismatches written to " $CompanyAOutFileName
