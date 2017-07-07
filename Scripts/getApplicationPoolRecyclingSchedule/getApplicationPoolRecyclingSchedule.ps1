<#
.FILENAME
	GetApplicationPoolRecyclingSchedule.ps1

.SYNOPSIS
	List Application Pools Recycling Schedule and Generate a CSV File

.DESCRIPTION	
	List Application Pool Recycling Schedule and Generate a CSV File
	
.LINK
	https://github.com/OhNotreDame/SPPS
	
.NOTES
	Author: JBO
	Modified by:JBON
	Modified:  0/07/2017
	Schedule : None

#>


function CreateRow($AppPool, $AppPoolState, $AppPoolSchedule)
{
    $objRow = New-Object System.Object
    $objRow | Add-Member -type NoteProperty -name ApplicationPool -value $AppPool
    $objRow | Add-Member -type NoteProperty -name State -value $AppPoolState
    $objRow | Add-Member -type NoteProperty -name RecyclingSchedule -value $AppPoolSchedule
    return $objRow
}


Clear-Host
Remove-Module *


$scriptName = $([System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Definition))
$cmdpath = $MyInvocation.MyCommand.Path
Write-Host "************************************************************************" -ForegroundColor Gray `r
Write-Host "$scriptName # Script started on $(get-date -format 'dd/MM/yyyy HH:mm:ss')" -ForegroundColor Gray `r
Write-Host "************************************************************************" -ForegroundColor Gray `r		

# Loading SharePoint Assembly and PS Snapin
Add-PsSnapin Microsoft.Sharepoint.PowerShell -EA silentlycontinue
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint") | Out-Null
[void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Web.Administration")

Import-Module WebAdministration

# Starting SPAssignment
Start-SPAssignment -Global

# CSV file Settings
$fileName = "$scriptName$(get-date -format 'yyyyMMddHHmm').csv"
$fileDestLocation = "D:\Scripts\$scriptName"

try
{
	#Intiate Results Object
	$rows = @()
	
	Get-Item IIS:\AppPools\* | ForEach-Object {									
								$appPoolName = $_.Name
								$appPoolState = $_.State
								$appPoolLocation = "IIS:\AppPools\$appPoolName"
								$appPoolRecyclingTime = (Get-ItemProperty ($appPoolLocation) -Name Recycling.periodicRestart.schedule.collection).value
							
								if ($appPoolRecyclingTime)
								{
									$AppPoolSchedule = "Every day at $appPoolRecyclingTime";
								}
								else
								{
									$AppPoolSchedule = "Disabled";
								}
							
								$rows += CreateRow $appPoolName $appPoolState $AppPoolSchedule
							}
							
	$rows
} 
catch 
{
	Write-Host "/!\ $scriptName An exception has been caught /!\ "  -ForegroundColor Red `r
	Write-Host "Type: " $_.Exception.GetType().FullName   -ForegroundColor Red `r
	Write-Host "Message: " $_.Exception.Message  -ForegroundColor Red `r
	Write-Host "Stacktrace: " $_.Exception.Stacktrace  -ForegroundColor Red `r
}
finally
{
	# Export CSV file
	$rows | Export-Csv "$fileDestLocation\$fileName" -NoTypeInformation -Delimiter ";" -Encoding UTF8
	Write-Host "CSV file available at $fileDestLocation\$fileName"  -ForegroundColor Green `r

	# Stopping SPAssignment and Transcript
	Stop-SPAssignment -Global

	Write-Host "************************************************************************" -ForegroundColor Gray `r
	Write-Host "$scriptName # Script ended on $(get-date -format 'dd/MM/yyyy HH:mm:ss')" -ForegroundColor Gray `r
	Write-Host "************************************************************************" -ForegroundColor Gray `r	
}	