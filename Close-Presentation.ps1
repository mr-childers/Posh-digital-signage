try
{
	If (Get-Process | ?{ $_.ProcessName -eq "POWERPNT" })
	{
        # Bind to PowerPoint COM-based object model.count number of open presentations and iterate closing all 
		$a = [System.Runtime.Interopservices.Marshal]::GetActiveObject('powerpoint.application')
		for ($ia = $a.Presentations.count; $ia -gt 0; $ia--)
		{
			$a.Presentations[$ia].Close()
		}
	}
}
catch
{
	Write-Host "Powerpoint not running"
}
finally
{
    # close the PowerPoint program and release COM object(s) from powershell
	Write-Host "cleaning up ..."
	$a.quit()
	$a = $null
    # if i have more then one powerpoint open this only closes one at a time 
    # [System.Runtime.Interopservices.Marshal]::ReleaseComObject($a)
    # Remove-Variable a
	1..2 | ForEach-Object { 
		#[System.GC]::Collect();
        [System.GC]::GetTotalMemory(‘forcefullcollection’) | out-null;
		[System.GC]::WaitForPendingFinalizers();
		}

}


<#
.Synopsis
   Short description
.DESCRIPTION
   Long description
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
#>
function Close-Presentation
{
    [CmdletBinding()]
    [OutputType([System.Void])]
    Param()

    Begin
    {

    }
    Process
    {

    }
    End
    {
    }
}



if (Get-Process | ?{ $_.ProcessName -eq "POWERPNT" })
{
    try
    {
        # Bind to PowerPoint COM-based object model.count number of open presentations and iterate closing all 
        $a = [System.Runtime.Interopservices.Marshal]::GetActiveObject('powerpoint.application')
		for ($ia = $a.Presentations.count; $ia -gt 0; $ia--)
		{
			$a.Presentations[$ia].Close()
		}
    }
    catch [DoSomethingCrazy]
    {
        Write-Host "Powerpoint not running"
    }
    catch [System.Net.WebException],[System.Exception]
    {
        Write-Host "Other exception"
    }
    finally
    {
        # close the PowerPoint program and release COM object(s) from powershell
	    Write-Host "cleaning up ..."
	    $a.quit()
	    $a = $null
        # if i have more then one powerpoint open this only closes one at a time 
        # [System.Runtime.Interopservices.Marshal]::ReleaseComObject($a)
        # Remove-Variable a
	    1..2 | ForEach-Object { 
		    #[System.GC]::Collect();
            [System.GC]::GetTotalMemory(‘forcefullcollection’) | out-null;
		    [System.GC]::WaitForPendingFinalizers();
		    }
        }
}



# cheak to see if any Presentations are open

$c = [System.Runtime.Interopservices.Marshal]::GetActiveObject('powerpoint.application')
$null -eq $c.Presentations[1]