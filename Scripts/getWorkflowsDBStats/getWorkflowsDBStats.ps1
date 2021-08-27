<#

.SYNOPSIS
	Get Stats about Nintex Workflows
	
.DESCRIPTION
	Get Stats about Nintex Workflows
	
.EXAMPLE
	GetWorkflowsDBStats.ps1
	
.OUTPUTS
	None
	
.LINK
	https://community.nintex.com/community/build-your-own/blog/2015/04/09/finding-all-of-the-workflows-in-your-farm-using-powershell
	
.NOTES
	Source: https://github.com/OhNotreDame/SPPS
	Created by: JBO
	Created: 29.03.2017
	Last Updated by: JBO
	Last Updated: 29.03.2017
	
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
	# Loading SharePoint and Nintex Objects into the PS session
	Add-PSSnapin Microsoft.SharePoint.PowerShell -EA silentlycontinue
	[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint");
	[void][System.Reflection.Assembly]::LoadWithPartialName("Nintex.Workflow");
	[void][System.Reflection.Assembly]::LoadWithPartialName("Nintex.Workflow.SupportConsole");
	[void][System.Reflection.Assembly]::LoadWithPartialName("Nintex.Workflow.Administration");
	[void][System.Reflection.Assembly]::LoadWithPartialName("Nintex.Forms.SharePoint.Administration");
	
	Start-SPAssignment -Global

	Write-Host "#**************************************************************#" -Foregroundcolor Cyan `r
	Write-Host "# $scriptName started" -Foregroundcolor Cyan `r  
	Write-Host "#**************************************************************#" -Foregroundcolor Cyan `r 
	Write-Host ""
	
	# Grab Nintex Config database name
	$CFGDB = [Nintex.Workflow.Administration.ConfigurationDatabase]::OpenConfigDataBase().Database

	
	Write-Host "#**************************************************************#" -Foregroundcolor Yellow `r
	Write-Host "# Nintex Workflows Instances COUNT" -Foregroundcolor Yellow `r
	Write-Host "#**************************************************************#" -Foregroundcolor Yellow `r
	
	# Creating instance of .NET SQL client
	$cmd = New-Object -TypeName System.Data.SqlClient.SqlCommand
	$cmd.CommandType = [System.Data.CommandType]::Text

	# Begin SQL Query 
	$cmd.CommandText = "SELECT COUNT (*) as 'Workflow Instance Count'			
						FROM [$CFGDB].[dbo].WorkflowInstance;";
	
	$reader = [Nintex.Workflow.Administration.ConfigurationDatabase]::GetConfigurationDatabase().ExecuteReader($cmd)
	while($reader.Read())
	{
		$Count_WfInstance = $reader["Workflow Instance Count"]
		Write-Host "Workflow Instance Count: $Count_WfInstance"  -Foregroundcolor Green `r
	}
	
	Write-Host "#**************************************************************#" -Foregroundcolor Yellow `r
	Write-Host "# Nintex Workflows Log COUNT" -Foregroundcolor Yellow `r
	Write-Host "#**************************************************************#" -Foregroundcolor Yellow `r
	
	# Creating instance of .NET SQL client
	$cmd = New-Object -TypeName System.Data.SqlClient.SqlCommand
	$cmd.CommandType = [System.Data.CommandType]::Text

	# Begin SQL Query 
	$cmd.CommandText = "SELECT COUNT (*) as 'Workflow Log Count'			
						FROM [$CFGDB].[dbo].WorkflowLog;";
	
	$reader = [Nintex.Workflow.Administration.ConfigurationDatabase]::GetConfigurationDatabase().ExecuteReader($cmd)
	while($reader.Read())
	{
		$Count_WfLog = $reader["Workflow Log Count"]
		Write-Host "Workflow Log Count: $Count_WfLog"  -Foregroundcolor Green `r
	}
	
	Write-Host "#**************************************************************#" -Foregroundcolor Yellow `r
	Write-Host "# Nintex Workflows Progress COUNT" -Foregroundcolor Yellow `r
	Write-Host "#**************************************************************#" -Foregroundcolor Yellow `r
	
	# Creating instance of .NET SQL client
	$cmd = New-Object -TypeName System.Data.SqlClient.SqlCommand
	$cmd.CommandType = [System.Data.CommandType]::Text

	# Begin SQL Query 
	$cmd.CommandText = "SELECT COUNT (*) as 'Workflow Progress Count'			
						FROM [$CFGDB].[dbo].WorkflowProgress;";
	
	$reader = [Nintex.Workflow.Administration.ConfigurationDatabase]::GetConfigurationDatabase().ExecuteReader($cmd)
	while($reader.Read())
	{
		$Count_WfProgress = $reader["Workflow Progress Count"]
		Write-Host "Workflow Progress Count: $Count_WfProgress"  -Foregroundcolor Green `r
	}
	
	Write-Host "#**************************************************************#" -Foregroundcolor Yellow `r
	Write-Host "# Nintex Database Fragmented Indexes" -Foregroundcolor Yellow `r
	Write-Host "#**************************************************************#" -Foregroundcolor Yellow `r
	
	# Creating instance of .NET SQL client
	$cmd = New-Object -TypeName System.Data.SqlClient.SqlCommand
	$cmd.CommandType = [System.Data.CommandType]::Text

	# Begin SQL Query 
	$cmd.CommandText = "SELECT OBJECT_NAME(i.object_id) AS TableName ,i.name AS IndexName ,phystat.avg_fragmentation_in_percent as IndexFragentation FROM sys.dm_db_index_physical_stats(DB_ID(), NULL, NULL, NULL, 'DETAILED') phystat inner JOIN sys.indexes i ON i.object_id = phystat.object_id AND i.index_id = phystat.index_id WHERE phystat.avg_fragmentation_in_percent > 10 ORDER BY phystat.avg_fragmentation_in_percent DESC"


	$reader = [Nintex.Workflow.Administration.ConfigurationDatabase]::GetConfigurationDatabase().ExecuteReader($cmd)
	$indexes = @()
	while($reader.Read())
	{
		$roundedIndex =  [math]::Round($reader["IndexFragentation"],2)
		$roundedPercent = $roundedIndex.ToString() + " %"
		
		$row = New-Object System.Object
		$row | Add-Member -MemberType NoteProperty -Name "TableName" -Value $reader["TableName"]
		$row | Add-Member -MemberType NoteProperty -Name "IndexName" -Value $reader["IndexName"]
		$row | Add-Member -MemberType NoteProperty -Name "IndexFragentation" -Value $roundedPercent
		$indexes += $row
	}
	$indexes

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
	Write-Host ""
	Write-Host "#**************************************************************#" -Foregroundcolor Cyan `r
	Write-Host "# $scriptName ended" -ForegroundColor Cyan `r
	Write-Host "#**************************************************************#" -Foregroundcolor Cyan `r
	
	Stop-SPAssignment -Global 
}