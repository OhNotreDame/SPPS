<#

.SYNOPSIS
	Get all Published workflows and generate dedicated CSV files (using Nintex DBs)
	
.DESCRIPTION
	Script used to collect general information about the Published workflows deployed on the SharePoint Farm	
	
.PARAMETER 
	None
	
.EXAMPLE
	getPublishedWorkflows.ps1
	
.OUTPUTS
	Two folders (Logs\ and CSV\)
	1 LOG file  [NC: GetPublishedWorkflows_<MachineName>.log]
	1 CSV file (Semicolon delimited) [NC: GetPublishedWorkflows_<MachineName>.csv]
	
.LINK
	https://community.nintex.com/community/build-your-own/blog/2015/04/09/finding-all-of-the-workflows-in-your-farm-using-powershell
	
.NOTES
	Source: https://github.com/OhNotreDame/SPPS
	Created by: JBO
	Created: 11.01.2017
	Last Updated by: JBO
	Last Updated: 12.01.2017
	
#>

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
	
	$logsFileName = $logsFolderName + $scriptName  + "_"+ $serverName + ".log"
	$csvWFFileName = $csvFolderName + $scriptName + "_" + $serverName + ".csv"
	
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
	[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint") 
	[void][System.Reflection.Assembly]::LoadWithPartialName("Nintex.Workflow")
	[void][System.Reflection.Assembly]::LoadWithPartialName("Nintex.Workflow.SupportConsole") 
	[void][System.Reflection.Assembly]::LoadWithPartialName("Nintex.Workflow.Administration")
	[void][System.Reflection.Assembly]::LoadWithPartialName("Nintex.Forms.SharePoint.Administration")

	
	Start-SPAssignment -Global
	Start-Transcript -path $logsFileName | Out-Null
	
	Write-Host "#**************************************************************#" -Foregroundcolor Cyan `r
	Write-Host "# Script $scriptName started" -Foregroundcolor Cyan `r  
	Write-Host "#**************************************************************#" -Foregroundcolor Cyan `r 
	Write-Host "# Transcript file: `n$logsFileName"  -ForegroundColor Cyan `r  
	Write-Host "#**************************************************************#" -Foregroundcolor Cyan `r
	Write-Host ""
	
	# Grab Nintex Config database name
	Write-Host "> Grab Nintex Config database name" -noNewLine
	$CFGDB = [Nintex.Workflow.Administration.ConfigurationDatabase]::OpenConfigDataBase().Database
	Write-Host -f Green " [Done]"
	
	Write-Host "#**************************************************************#" -Foregroundcolor Yellow `r
	Write-Host "# Published Workflows" -Foregroundcolor Yellow `r
	Write-Host "#**************************************************************#" -Foregroundcolor Yellow `r
	
	Write-Host "> Prepare SQL Query" -noNewLine
	$cmd = New-Object -TypeName System.Data.SqlClient.SqlCommand
	$cmd.CommandType = [System.Data.CommandType]::Text

	# Begin SQL Query 
	$cmd.CommandText = "SELECT  PW.[WorkflowId],
								PW.[WorkflowName],
								PW.[WorkflowType],
								PW.[PublishTime],
								PW.[Version],
								PW.[Author],
								PW.[SiteID],
								PW.[WebID],
								PW.[ListID]
						FROM [$CFGDB].dbo.PublishedWorkflows PW
						INNER JOIN (SELECT [WorkflowId],MAX([Version]) AS MaxVersion
									FROM [$CFGDB].dbo.PublishedWorkflows
									GROUP BY WorkflowId) PWMax
							ON PWMax.[WorkflowId]=PW.[WorkflowId]
							AND PWMax.MaxVersion=PW.[Version];"
	
	Write-Host -f Green " [Done]"	
	
	#CSV Content for Workflows list
	$wfList = @()
	
	Write-Host "> Parsing SQL Query results" -noNewLine
	#Call to find all Nintex Content Databases in the Nintex Configuration Database, then execute the above query against each.  
	foreach ($database in [Nintex.Workflow.Administration.ConfigurationDatabase]::GetConfigurationDatabase().ContentDatabases)
	{
		$reader = $database.ExecuteReader($cmd)

		while($reader.Read())
		{
			$wfPubRow = New-Object System.Object
			$webTitle= ""
			$webURL= ""
			$listTitle= ""

			if(![string]::IsNullOrEmpty($reader["SiteID"])) {
				$siteID = [Guid] $reader["SiteID"];
				$site = Get-SPSite -identity $siteID -ea silentlycontinue
			}
			
			if((![string]::IsNullOrEmpty($reader["WebID"])) -and ($site -ne $null)) {
				$webID = [Guid] $reader["WebID"];
				$web = $site.AllWebs[$webID]
				$webTitle= $web.Title
				$webURL= $web.URL
			}

			if((![string]::IsNullOrEmpty($reader["ListID"]))  -and ($web -ne $null)){
				$listID = [Guid] $reader["ListID"];
				$list = $web.Lists[$listID]
				$listTitle= $list.Title
			}
			
			#Adding Query results to table object
			$wfPubRow | Add-Member -MemberType NoteProperty -Name "WorkflowName" -Value $reader["WorkflowName"]
			$wfPubRow | Add-Member -MemberType NoteProperty -Name "WorkflowType" -Value $reader["WorkflowType"]
			$wfPubRow | Add-Member -MemberType NoteProperty -Name "Published" -Value $reader["PublishTime"]
			$wfPubRow | Add-Member -MemberType NoteProperty -Name "Version" -Value $reader["Version"]
			$wfPubRow | Add-Member -MemberType NoteProperty -Name "SiteName" -Value $webTitle
			$wfPubRow | Add-Member -MemberType NoteProperty -Name "SiteURL" -Value $webURL
			$wfPubRow | Add-Member -MemberType NoteProperty -Name "List" -Value $listTitle
			$wfPubRow | Add-Member -MemberType NoteProperty -Name "Author" -Value $reader["Author"]

			$wfList += $wfPubRow
			
			if ($web -ne $null)
			{
				$web.Dispose()
			}
	
			if ($site -ne $null)
			{
				$site.Dispose()
			}
		}

	}
	Write-Host -f Green " [Done]"	
	
	Write-Host "> Generate CSV file for Published Workflows" -noNewLine
	$wfList | Export-CSV -Path $csvWFFileName -NoTypeInformation -Delimiter ";" 
	Write-Host -f Green " [Done]"	
}
catch [Exception]
{	
	Write-Host ""
	Write-Host "/!\ $scriptName An exception has been caught /!\ "
	Write-Host "Type: " $_.Exception.GetType().FullName
	Write-Host "Message: " $_.Exception.Message
	Write-Host "Stacktrace: `n" $_.Exception.Stacktrace
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