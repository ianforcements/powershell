#Ping Machines in local Domain
#by Ian Hutchinson 28/09/2015
#This script takes an input csv of AD computer objects, and tests a connection with each.
#It will note the results of the test in an output CSV


#get ready
import-module activedirectory


#get input
$machinesToTest = import-csv ".\InactiveMachines.csv"

#get going
$machinesToTest | foreach-object {
$didPing = test-connection -ComputerName $_.cn -count 1 -quiet -delay 1
write-host $_.cn "ping result:" $didping
$_ | Add-Member -NotePropertyName 'DidPing' -NotePropertyValue $didPing -force
}

#get out
$machinesToTest | export-csv ".\InactiveMachinesPingResults.csv"


