<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2021 v5.8.194
	 Created on:   	10/25/2021 5:12 PM
	 Created by:   	gchilders
	 Organization: 	
	 Filename:     	
	===========================================================================
	.DESCRIPTION
		A description of the file.
#>
<#
function Get-EventRecord
{
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[string]$eventRecordID,
		[Parameter(Mandatory = $true)]
		[string]$eventChannel
	)
	
	$event = get-winevent -LogName $eventChannel -FilterXPath "<QueryList><Query Id='0' Path='$eventChannel'><Select Path='$eventChannel'>*[System [(EventRecordID=$eventRecordID)]]</Select></Query></QueryList>"
	$JsonObject = $event.Message | ConvertFrom-Json
	
	[System.Windows.MessageBox]::Show( "$($JsonObject.FullPath) removed")
}


Get-EventRecord $eventRecordID $eventChannel
#Get-EventRecord $eventParams $event.Message


# -eventRecordID $(eventRecordID) -eventChannel $(eventChannel)
#>



param($eventRecordID, $eventChannel)

Add-Type -AssemblyName PresentationFramework

    
    write-host $eventRecordID
    write-host $eventChannel
    write-host $eventSeverity

    $eventChannel = 'Breakroom Display'

    write-host $eventChannel

	$event = get-winevent -LogName $eventChannel -FilterXPath "<QueryList><Query Id='0' Path='$eventChannel'><Select Path='$eventChannel'>*[System [(EventRecordID=$eventRecordID)]]</Select></Query></QueryList>"
	$JsonObject = $event.Message | ConvertFrom-Json
	
	[System.Windows.MessageBox]::Show( "$($JsonObject.FullPath) removed")