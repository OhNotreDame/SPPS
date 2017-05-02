<#

.SYNOPSIS
	Get all Workflows Instances from a specific status and generate dedicated CSV files.
	
.DESCRIPTION
	Script used to collect general information about the workflows deployed and running on the SharePoint Farm by Status
	
.PARAMETER status
	status of the Workflow Instance to be collected
	Should be 'Running','Completed','Cancelled','Error'
	
.EXAMPLE
	GetWorkflowsInstancesBystatus.ps1 -status Running
	GetWorkflowsInstancesBystatus.ps1 -status Completed
	GetWorkflowsInstancesBystatus.ps1 -status Cancelled
	GetWorkflowsInstancesBystatus.ps1 -status Error
	
.OUTPUTS
	Two folders (Logs\ and CSV\)
	1 log file  [NC: GetWorkflowsInstancesBystatus_<MachineName>_<Status>.log]
	for all Workflows Instances: 1 CSV file (Semicolon delimited) [NC: GetWorkflowsInstancesBystatus_<MachineName>_<Status>.csv]

.LINK
	https://community.nintex.com/community/build-your-own/blog/2015/04/09/finding-all-of-the-workflows-in-your-farm-using-powershell
	
.NOTES
	Created by: JBO
	Created: 11.01.2017
	Last Updated by: JBO
	Last Updated: 26.04.2017
	
#>

param
(
	[Parameter(Mandatory=$true, Position=1)]
	[ValidateNotNullOrEmpty()]
	[ValidateSet('Running','Completed','Cancelled','Error')]
	[string]$status
)

Clear-Host
Remove-Module *

########################
### GLOBAL VARIABLES ###
########################
$scriptName = $([System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Definition))


########################
####### SETTINGS #######
########################
$logsFolderName = $PSScriptRoot+"\Logs\"
$csvFolderName = $PSScriptRoot+"\CSV\"

try 
{
	$serverName = $env:computername
	
	$logsFileName = $logsFolderName + $scriptName  + "_"+ $serverName + "_" + $status +  ".log"
	$csvWFInstanceFileName = $csvFolderName + $scriptName + "_" + $serverName  +"_" + $status + ".csv"
	
	#by default, will grab the Running Workflow
	$statusInt = 2; 
	
	switch ($status)
	{
		'Running' 	{ $statusInt = 2; break; }
		'Completed' { $statusInt = 4; break; }
		'Cancelled' { $statusInt = 8; break; }
		'Error'		{ $statusInt = 64; break; }
		default 	{ $statusInt = 2; break; }
	}
	

	########################
	#### PRE-REQUISITES ####
	########################
	
	if(!(Test-Path $logsFolderName))
	{
		New-Item $logsFolderName -type Directory -Force | Out-Null
	}
	
	if(!(Test-Path $csvFolderName))
	{
		New-Item $csvFolderName -type Directory -Force | Out-Null
	} 

	########################
	## MONITORING & LOGS ###
	########################
	
	
	# Loading SharePoint and Nintex Objects into the PS session
	Add-PSSnapin Microsoft.SharePoint.PowerShell -EA silentlycontinue
	[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint");
	[void][System.Reflection.Assembly]::LoadWithPartialName("Nintex.Workflow");
	[void][System.Reflection.Assembly]::LoadWithPartialName("Nintex.Workflow.SupportConsole");
	[void][System.Reflection.Assembly]::LoadWithPartialName("Nintex.Workflow.Administration");
	[void][System.Reflection.Assembly]::LoadWithPartialName("Nintex.Forms.SharePoint.Administration");

	
	Start-SPAssignment -Global
	Start-Transcript -path $logsFileName | Out-Null
	
	Write-Host "#**************************************************************#" -Foregroundcolor Cyan `r
	Write-Host "# Script $scriptName started" -Foregroundcolor Cyan `r  
	Write-Host "#**************************************************************#" -Foregroundcolor Cyan `r 
	Write-Host "# Transcript file: `n$logsFileName"  -ForegroundColor Cyan `r  
	Write-Host "# CSV file: `n$csvWFInstanceFileName"  -ForegroundColor Cyan `r  
	Write-Host "# Status: $status ($statusInt)"  -ForegroundColor Cyan `r  
	Write-Host "#**************************************************************#" -Foregroundcolor Cyan `r
	Write-Host ""
	
	# Grab Nintex Config database name
	Write-Host "> Grab Nintex Config database name" -noNewLine
	$CFGDB = [Nintex.Workflow.Administration.ConfigurationDatabase]::OpenConfigDataBase().Database
	Write-Host -f Green " [Done]"

	
	Write-Host "#**************************************************************#" -Foregroundcolor Yellow `r
	Write-Host "# Running Workflows Instances" -Foregroundcolor Yellow `r
	Write-Host "#**************************************************************#" -Foregroundcolor Yellow `r
	
	# Creating instance of .NET SQL client
	Write-Host "> Prepare SQL Query" -noNewLine
	$cmd = New-Object -TypeName System.Data.SqlClient.SqlCommand
	$cmd.CommandType = [System.Data.CommandType]::Text

	# Begin SQL Query 
	$cmd.CommandText = "SELECT	I.[WorkflowName]
								,I.[SiteID]
								,I.[WebID]
								,I.[ListID]
								,I.[ItemID]
								,I.[WorkflowInitiator]
								,I.StartTime							
						FROM [$CFGDB].dbo.WorkflowInstance I
						WHERE I.[STATE] = " + $statusInt +" ;";
	
	Write-Host -f Green " [Done]"	
	
	#CSV Content for WorkflowInstances list
	$wfInstanceList = @()
	
	Write-Host "> Parsing SQL Query results" -noNewLine
	#Call to find all Nintex Content Databases in the Nintex Configuration Database, then execute the above query against each.  
	foreach ($database in [Nintex.Workflow.Administration.ConfigurationDatabase]::GetConfigurationDatabase().ContentDatabases)
	{
		$reader = $database.ExecuteReader($cmd)

		while($reader.Read())
		{
			$wfInstRow = New-Object System.Object
			
			if([string]::IsNullOrEmpty($reader["SiteID"])) 
			{
				Write-Debug "SiteID returned by SQL Statement is null."
			}
			else
			{
				# Convert string SiteID into Guid SiteID
				$siteID = [Guid] $reader["SiteID"];
				
				# Get SPSite 
				$site = $(Get-SPSite -identity $siteID  -EA SilentlyContinue)
				
				if ($site -eq $null)
				{
					Write-Debug "SPSite not found for the SiteID '$siteID'."
				}
				else
				{
					if([string]::IsNullOrEmpty($reader["WebID"])) 
					{
						Write-Debug "WebID returned by SQL Statement is null."
					}
					else
					{
						Write-Debug "Workflow '$($reader['WorkflowName'])' on site '$($site.URL)'" 
						
						# Convert string WebID into Guid WebID
						$webID = [Guid] $reader["WebID"];
						
						# Get SPWeb 
						$web = $site.AllWebs[$webID]
						
						if ($web -eq $null)
						{
							Write-Debug "SPWeb not found for this WebID '$webID'."
						}
						else
						{
							if([string]::IsNullOrEmpty($reader["ListID"])) 
							{
								Write-Debug "ListID returned by SQL Statement is null."
							}
							else
							{
								# Convert string ListID into Guid ListID
								$listID = [Guid] $reader["ListID"];
								
								# Get SPList
								if ($listID -eq [system.guid]::empty)
								{
									Write-Debug "List GUID is empty."
									$listName = ""
									$workflowType = "Site"
								}
								else
								{
									$list = $web.Lists[$listID]								
									if ($list -eq $null)
									{
										Write-Debug "SPList not found for this ListID '$listID'."
										$workflowType = "Site"
									}
									else
									{
										$listName = $List.title;
										$workflowType = "List"
									}
									
									#Adding Query results to table object
									$wfInstRow | Add-Member -MemberType NoteProperty -Name "WorkflowName" -Value $reader["WorkflowName"]
									$wfInstRow | Add-Member -MemberType NoteProperty -Name "Type" -Value $workflowType
									$wfInstRow | Add-Member -MemberType NoteProperty -Name "SiteName" -Value $web.Title
									$wfInstRow | Add-Member -MemberType NoteProperty -Name "SiteURL" -Value $web.URL
									$wfInstRow | Add-Member -MemberType NoteProperty -Name "List" -Value $listName
									$wfInstRow | Add-Member -MemberType NoteProperty -Name "ItemID" -Value $reader["ItemID"]
									$wfInstRow | Add-Member -MemberType NoteProperty -Name "WorkflowInitiator" -Value $reader["WorkflowInitiator"]
									$wfInstRow | Add-Member -MemberType NoteProperty -Name "StartTime" -Value $reader["StartTime"]

									if ($statusInt -eq 2)
									{
										#Calculating Duration of the Workflow
										$now = Get-Date
										$starDate = [datetime] $reader["StartTime"]
										$duration = NEW-TIMESPAN $starDate $now
										$wfInstRow | Add-Member -MemberType NoteProperty -Name "Duration (Days)" -Value $duration.Days
									}
									else
									{
										Write-Debug "No need to compute duration up to now, workflow is terminated."
									}
									
									$wfInstanceList += $wfInstRow

									}
							}#end if ($lsit)
							
						}#end if ($web)

					}				
				}#end if ($site)

			}
			if ($web -ne $null) {
				$web.Dispose()
			}
			
			if ($site -ne $null) {
				$site.Dispose()
			}
			
		}#end while($reader.Read())

	}#end #foreach
	Write-Host -f Green " [Done]"	
	
	Write-Host "> Generate CSV file for Running Workflow Instance" -noNewLine
	$wfInstanceList | Export-CSV -Path $csvWFInstanceFileName -NoTypeInformation -Delimiter ";" -Encoding UTF8
	Write-Host -f Green " [Done]"	

}
catch [Exception]
{	
	Write-Host ""
	Write-Host "/!\ [$scriptName] An exception has been caught /!\ " -Foregroundcolor Red `r
	Write-Host "Type: " $_.Exception.GetType().FullName -Foregroundcolor Red `r
	Write-Host "Message: " $_.Exception.Message -Foregroundcolor Red `r
	Write-Host "Stacktrace: `n" $_.Exception.Stacktrace -Foregroundcolor Red `r
}
finally
{
	if ($web -ne $null)
	{
		$web.Dispose()
	}
	
	if ($site -ne $null)
	{
		$site.Dispose()
	}
	
	Write-Host ""
	Write-Host "#**************************************************************#" -Foregroundcolor Cyan `r
	Write-Host "# Script $scriptName ended" -ForegroundColor Cyan `r
	Write-Host "#**************************************************************#" -Foregroundcolor Cyan `r
	
	Stop-SPAssignment -Global 
	Stop-Transcript | Out-Null
}
