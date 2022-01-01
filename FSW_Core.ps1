<#	
	.NOTES
	===========================================================================
   	Find a way to set this script as a windows service and add restart 
	===========================================================================
#>
Clear-Host
$ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
$fileSystemWatcherDirPath = 'C:\Presentation'
$fileSystemWatcherFilter = '*.*'
$fileSystemWatcher = [System.IO.FileSystemWatcher]::new($fileSystemWatcherDirPath, $fileSystemWatcherFilter)
$fileSystemWatcher.IncludeSubdirectories = $false
$fileSystemWatcher.EnableRaisingEvents = $true
$fileSystemWatcher.NotifyFilter = [System.IO.NotifyFilters]::FileName -bor [System.IO.NotifyFilters]::DirectoryName -bor [System.IO.NotifyFilters]::LastWrite # [System.Linq.Enumerable]::Sum([System.IO.NotifyFilters].GetEnumValues())

# Create syncronized hashtable
$syncdFsItemEventHashT = [hashtable]::Synchronized([hashtable]::new())

$fileSystemWatcherAction = {
	try
	{
		$fsItemEvent = [pscustomobject]@{
			EventIdentifier  = $Event.EventIdentifier
			SourceIdentifier = $Event.SourceIdentifier
			TimeStamp	     = $Event.TimeGenerated
			FullPath	     = $Event.SourceEventArgs.FullPath
			ChangeType	     = $Event.SourceEventArgs.ChangeType
		}
		
		# Collecting event in synchronized hashtable (overrides existing keys so that only the latest event details are available)
		$syncdFsItemEventHashT[$fsItemEvent.FullPath] = $fsItemEvent
	}
	catch
	{
		Write-Host ($_ | Format-List * | Out-String) -ForegroundColor red
	}
}
# Script block which processes collected events and do further actions like copying for backup, etc...
# NoGo "Start-Job" it's not possible to access and modify the synchronized hashtable created within this scope.
$fSItemEventProcessingJob = {
	$keys = [string[]]$syncdFsItemEventHashT.psbase.Keys
	
	foreach ($key in $keys)
	{
		$fsEvent = $syncdFsItemEventHashT[$key]
		
		try
		{
			# in case changetype eq DELETED or the item can't be found on the filesystem by the script -> remove the item from hashtable without any further actions.
			# This affects temporary files from applications. BUT: Could also affect files with file permission issues.
			if (($fsEvent.ChangeType -eq [System.IO.WatcherChangeTypes]::Deleted) -or (! (Test-Path -LiteralPath $fsEvent.FullPath)))
			{
				$syncdFsItemEventHashT.Remove($key)
				Write-Host ("==> Item '$key' with changetype '$($fsEvent.ChangeType)' removed from hashtable without any further actions!") -ForegroundColor Blue
				continue
			}
			
			# get filesystem object
			$fsItem = Get-Item -LiteralPath $fsEvent.FullPath -Force
			
			if ($fsItem -is [System.IO.FileInfo])
			{
				# file processing
				
				try
				{
					# Check whether the file is still locked / in use by another process
					[System.IO.FileStream]$fileStream = [System.IO.File]::Open($fsEvent.FullPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
					$fileStream.Close()
				}
				catch [System.IO.IOException] {
					Write-Host ("==> Item '$key' with changetype '$($fsEvent.ChangeType)' is still in use and can't be read!") -ForegroundColor Yellow
					continue
				}
			}
			elseIf ($fsItem -is [System.IO.DirectoryInfo])
			{
				# directory processing
			}
			
			$syncdFsItemEventHashT.Remove($key)
			Write-Host ("==> Item '$key' with changetype '$($fsEvent.ChangeType)' has been processed and removed from hashtable.") -ForegroundColor Blue
			
		}
		catch
		{
			Write-Host ($_ | Format-List * | Out-String) -ForegroundColor red
		}
	}
}

[void] (Register-ObjectEvent -InputObject $fileSystemWatcher -EventName 'Created' -SourceIdentifier 'FSCreated' -Action $fileSystemWatcherAction)
[void] (Register-ObjectEvent -InputObject $fileSystemWatcher -EventName 'Changed' -SourceIdentifier 'FSChanged' -Action $fileSystemWatcherAction)
[void] (Register-ObjectEvent -InputObject $fileSystemWatcher -EventName 'Renamed' -SourceIdentifier 'FSRenamed' -Action $fileSystemWatcherAction)
[void] (Register-ObjectEvent -InputObject $fileSystemWatcher -EventName 'Deleted' -SourceIdentifier 'FSDeleted' -Action $fileSystemWatcherAction)

Write-Host "Watching for changes in '$fileSystemWatcherDirPath'.`r`nPress CTRL+C to exit!"
try
{
	do
	{
		Wait-Event -Timeout 1
		
		if ($syncdFsItemEventHashT.Count -gt 0)
		{
			Write-Host "`r`n"
			Write-Host ('-' * 50) -ForegroundColor Green
			Write-Host "Collected events in hashtable queue:" -ForegroundColor Green
			$syncdFsItemEventHashT.Values | Format-Table | Out-String

			$EventLog = 'Breakroom Display'
			$SourceType = $syncdFsItemEventHashT[$syncdFsItemEventHashT.Keys].SourceIdentifier
			
			$EventData = [ordered]@{
				Program		     = 'BreakRoom Dispaly Automation Script';
				EventIdentifier  = $syncdFsItemEventHashT[$syncdFsItemEventHashT.Keys].EventIdentifier;
				SourceIdentifier = $syncdFsItemEventHashT[$syncdFsItemEventHashT.Keys].SourceIdentifier;
				TimeStamp	     = $syncdFsItemEventHashT[$syncdFsItemEventHashT.Keys].TimeStamp;
				FullPath		 = $syncdFsItemEventHashT[$syncdFsItemEventHashT.Keys].FullPath;
				ChangeType	     = $syncdFsItemEventHashT[$syncdFsItemEventHashT.Keys].ChangeType;
			}
			
			# New-EventSource -EventLog $EventLog -Source $SourceType  - build a function that checks to see if exist on startup, 
			# incorp starting the last known powerpoint aswell. 
			function New-EventSource 
			{
				[CmdLetBinding()]
				param (
					[string]$EventLog,
					[string]$Source
				)
				
				if ([System.Diagnostics.EventLog]::SourceExists($Source) -eq $false)
				{
					try
					{
						[System.Diagnostics.EventLog]::CreateEventSource($Source, $EventLog)
					}
					catch
					{
						$PSCmdlet.ThrowTerminatingError($_)
					}
				}
				else
				{
					'Source {0} for event log {1} already exists' -f $Source, $EventLog | Write-Warning
				}
			}
			
			function Write-WinEvent
			{
				[CmdLetBinding()]
				param (
					[string]$LogName,
					[string]$Provider,
					[int64]$EventId,
					[System.Diagnostics.EventLogEntryType]$EventType,
					[System.Collections.Specialized.OrderedDictionary]$EventData,
					[ValidateSet('JSON', 'CSV', 'XML')]
					[string]$MessageFormat = 'JSON'
				)
				
				$EventMessage = @()
				
				switch ($MessageFormat)
				{
					'JSON' { $EventMessage += $EventData | ConvertTo-Json }
					'CSV' { $EventMessage += ($EventData.GetEnumerator() | Select-Object -Property Key, Value | ConvertTo-Csv -NoTypeInformation) -join "`n" }
					'XML' { $EventMessage += ($EventData | ConvertTo-Xml).OuterXml }
				}
				
				$EventMessage += foreach ($Key in $EventData.Keys)
				{
					'{0}:{1}' -f $Key, $EventData.$Key
				}
				
				try
				{
					$Event = [System.Diagnostics.EventInstance]::New($EventId, $null, $EventType)
					$EventLog = [System.Diagnostics.EventLog]::New()
					$EventLog.Log = $LogName
					$EventLog.Source = $Provider
					$EventLog.WriteEvent($Event, $EventMessage)
				}
				catch
				{
					$PSCmdlet.ThrowTerminatingError($_)
				}
			}
			
			# Switch case used to provide the ability to assign individual Event ID's to each file modification action. 
			# This will later be used as a windows task scheduler trigger to fire off a Powershell script when detected 
			
			switch ($SourceType)
			{
				'FSChanged' { Write-WinEvent -LogName $EventLog -Provider $SourceType -EventId 1 -EventType Information -EventData $EventData }
				'FSCreated' { Write-WinEvent -LogName $EventLog -Provider $SourceType -EventId 2 -EventType Information -EventData $EventData }
				'FSDeleted' { Write-WinEvent -LogName $EventLog -Provider $SourceType -EventId 3 -EventType Information -EventData $EventData }
				'FSRenamed' { Write-WinEvent -LogName $EventLog -Provider $SourceType -EventId 4 -EventType Information -EventData $EventData }
			}
			
		}
		
		# Process hashtable items and do something with them (like copying, ..)
		.$fSItemEventProcessingJob
		
		# Garbage collector
		[GC]::Collect()
		
	}
	while ($true)
	
}
finally
{
	# unregister
	Unregister-Event -SourceIdentifier 'FSChanged'
	Unregister-Event -SourceIdentifier 'FSCreated'
	Unregister-Event -SourceIdentifier 'FSDeleted'
	Unregister-Event -SourceIdentifier 'FSRenamed'
	
	# dispose
	$FileSystemWatcher.Dispose()
	Write-Host "`r`nEvent Handler removed."
}
