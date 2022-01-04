

# Check if Log exists
# Ref: http://msdn.microsoft.com/en-us/library/system.diagnostics.eventlog.exists(v=vs.110).aspx

[System.Diagnostics.EventLog]::Exists('Breakroom Display');

# make event log - can skip
{New-eventlog -logname "Breakroom Display" -Source $_} 


# Ref: http://msdn.microsoft.com/en-us/library/system.diagnostics.eventlog.sourceexists(v=vs.110).aspx
# Check if Source exists

'FSCreated','FSChanged','FSRenamed','FSDeleted' | ForEach-Object {[System.Diagnostics.EventLog]::SourceExists($_)}


# Create Event Log Source 
'FSCreated','FSChanged','FSRenamed','FSDeleted' | ForEach-Object {New-eventlog -logname "Breakroom Display" -Source $_} 

# Write Sample Logs

Write-EventLog -LogName 'Breakroom Display' -Source 'FSCreated' -Message "init sample event entry FSCreated" -EventId 1 -EntryType information
Write-EventLog -LogName 'Breakroom Display' -Source 'FSChanged' -Message "init sample event entry FSChanged" -EventId 2 -EntryType information
Write-EventLog -LogName 'Breakroom Display' -Source 'FSRenamed' -Message "init sample event entry FSRenamed" -EventId 3 -EntryType information
Write-EventLog -LogName 'Breakroom Display' -Source 'FSDeleted' -Message "init sample event entry FSDeleted" -EventId 4 -EntryType information








$logFileExists = Get-EventLog -list | Where-Object {$_.logdisplayname -eq 'Breakroom Display'} 
if (! $logFileExists) {
Write-Host "does not exist "
}



# Build Event Logs
$logFileExists = Get-EventLog -list | Where-Object {$_.logdisplayname -eq 'Breakroom Display'} 
if (! $logFileExists) {
    # Create Event Log Source 
    'FSCreated', 'FSChanged','FSRenamed','FSDeleted' | ForEach-Object {New-eventlog -logname "Breakroom Display" -Source $_} 
}

# Test Eventlogs
Write-EventLog -LogName 'Breakroom Display' -Source 'FSCreated' -Message "init sample event entry FSCreated" -EventId 1 -EntryType information
Write-EventLog -LogName 'Breakroom Display' -Source 'FSChanged' -Message "init sample event entry FSChanged" -EventId 2 -EntryType information
Write-EventLog -LogName 'Breakroom Display' -Source 'FSRenamed' -Message "init sample event entry FSRenamed" -EventId 3 -EntryType information
Write-EventLog -LogName 'Breakroom Display' -Source 'FSDeleted' -Message "init sample event entry FSDeleted" -EventId 4 -EntryType information


# output the Eventlog and each of it's sources.
Get-WmiObject win32_nteventlogfile -Filter "logfilename='Breakroom Display'" | foreach {$_.sources}

# Remove the eventlog sources 
Remove-Eventlog -Source "FSCreated"
Remove-Eventlog -Source "FSChanged"
Remove-Eventlog -Source "FSRenamed"
Remove-Eventlog -Source "FSDeleted"
# Remove the Event log 
Remove-EventLog -LogName 'Breakroom Display'







#######
## Cheak and create event logs
#######

'FSCreated','FSChanged','FSRenamed','FSDeleted' | % {
    $x = IndexOf($_)
    if ([System.Diagnostics.EventLog]::SourceExists($_))
    {write-host "eventlog source $_ already exists"}
    else
    {write-host (New-eventlog -logname "Breakroom Display" -Source $_ )} 
  }



