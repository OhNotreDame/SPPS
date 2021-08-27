<#

.SYNOPSIS
	Clean Nintex Workflows History (SharePoint List & SQL Progress Table) on SPWeb
	
.DESCRIPTION
	Clean Nintex Workflows History (SharePoint List & SQL Progress Table) on SPWeb
	
.PARAMETER status
	Status of the Workflow Instance to be collected
	Should be 'Completed','Cancelled','Error'
	
.PARAMETER siteUrl
	(Mandatory) URL of the SharePoint Site (SPWeb)
	
.EXAMPLE
	cleanWorkflowsHistoryOnSite.ps1 siteUrl <siteUrl> -status <'Completed','Cancelled','Error'>
	
.OUTPUTS
	None
	
.LINK
	None
	
.NOTES
	Source: https://github.com/OhNotreDame/SPPS
	Created by: JBO
	Created: 29.03.2017
	Last Updated by: JBO
	Last Updated: 30.03.2017
	
#>
param
(
	[Parameter(Mandatory=$true, Position=1)]
	[ValidateNotNullOrEmpty()]
	[string]$siteURL,
	[Parameter(Mandatory=$true, Position=2)]
	[ValidateNotNullOrEmpty()]
	[ValidateSet('Completed','Cancelled','Error')]
	[string]$status
)


Clear-Host
Remove-Module *

########################
### GLOBAL VARIABLES ###
########################
$scriptName = $([System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Definition))

Write-Host "#**************************************************************#" -Foregroundcolor Cyan `r
Write-Host "# $scriptName started" -ForegroundColor Cyan `r
Write-Host "#**************************************************************#" -Foregroundcolor Cyan `r		
Write-Host "# status: $status" -Foregroundcolor Cyan `r	# 
Write-Host "# siteURL: $siteURL" -Foregroundcolor Cyan `r	


#####################################################
# Loading SharePoint Assembly and PS Snapin
#####################################################
Add-PsSnapin Microsoft.Sharepoint.PowerShell -EA silentlycontinue
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint") | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName("Nintex.Workflow") | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName("Nintex.Workflow.SupportConsole") | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName("Nintex.Workflow.Administration") | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName("Nintex.Forms.SharePoint.Administration") | Out-Null
Import-Module "$PSScriptRoot\CleanWorkflowsHistoryCommon.psm1"


########################
####### SETTINGS #######
########################
$logsFolderName = $PSScriptRoot+"\Logs\"
$csvFolderName = $PSScriptRoot+"\CSV\"



Start-SPAssignment -Global

try 
{	
	$web = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
	if($web -ne $null) 
	{
		$siteName = $web.Title
		$siteID = $web.ID
		
		Write-Host "# siteName: $siteName"  -Foregroundcolor Cyan `r	
		Write-Host "# siteID: $siteID" -Foregroundcolor Cyan `r	 
		Write-Host "#**************************************************************#" -Foregroundcolor Cyan `r	
	
		Write-Host ""
		Write-Host "# 1 - Get Workflows Stats (prior to cleaning operations)"  -Foregroundcolor Cyan `r	
		Write-Host "#**************************************************************#" -Foregroundcolor Cyan `r	
	
		
		# Workflow History List on this site
		$Count_WfHistoryList = getNintexWorklowsHistoryItemsCount -siteURL $siteURL
		if ($Count_WfHistoryList -ne -1)
		{
			Write-Host "Workflow History List (SP) Count: $Count_WfHistoryList"  -Foregroundcolor Green `r
		}
		else
		{
			Write-Host "Something went wrong with getNintexWorklowsHistoryItemsCount()" -Foregroundcolor Red `r	
		}
		
		
		# Workflow Instances (DB - Nintex) on this site
		$Count_WfInstance = getSPWebNintexWorklowsInstanceDBCount -siteURL $siteURL
		if ($Count_WfInstance -ne -1)
		{
			Write-Host "Workflow Instance (DB - Nintex) Count: $Count_WfInstance"  -Foregroundcolor Green `r
		}
		else
		{
			Write-Host "Something went wrong with getNintexWorklowsInstanceDBrowsCount()" -Foregroundcolor Red `r	
		}
	
	
		# Workflow Progress (DB - Nintex) on this site
		$Count_WfProgress = getSPWebNintexWorklowsProgressDBCount -siteURL $siteURL
		if ($Count_WfProgress -ne -1)
		{
			Write-Host "Workflow Progress (DB - Nintex) Count: $Count_WfProgress"  -Foregroundcolor Green `r
		}
		else
		{
			Write-Host "Something went wrong with getNintexWorklowsProgressDBCount()" -Foregroundcolor Red `r	
		}
	
		Write-Host ""
		Write-Host "# 2 - Clear Workflow Data by Workflow Status"  -Foregroundcolor Cyan `r	
		Write-Host "#**************************************************************#" -Foregroundcolor Cyan `r	

		# Clear Workflow History List
		clearWorkflowHistory -siteURL $siteURL -status $status

		# Clear Workflow Progress Data
		clearWorkflowData -siteURL $siteURL -status $status

		Write-Host ""
		Write-Host "# 3 - Clear Workflow Data for Deleted Lists (if any)"  -Foregroundcolor Cyan `r	
		Write-Host "#**************************************************************#" -Foregroundcolor Cyan `r	
				
		# Clear Workflow History List for deleted lists
		clearWorkflowHistoryForDeletedLists -siteURL $siteURL 
		
		# Clear Workflow Progress Data for deleted lists
		clearWorkflowDataForDeletedLists -siteURL $siteURL 

		Write-Host ""
		Write-Host "# 4 - Get Workflows Stats (after cleaning operations)"  -Foregroundcolor Cyan `r	
		Write-Host "#**************************************************************#" -Foregroundcolor Cyan `r	
	
		# Workflow History List on this site
		$Count_WfHistoryList = getNintexWorklowsHistoryItemsCount -siteURL $siteURL
		if ($Count_WfHistoryList -ne -1)
		{
			Write-Host "Workflow History List (SP) Count: $Count_WfHistoryList"  -Foregroundcolor Green `r
		}
		else
		{
			Write-Host "Something went wrong with getNintexWorklowsHistoryItemsCount()" -Foregroundcolor Red `r	
		}

		# Workflow Instances (DB - Nintex) on this site		
		$Count_WfInstance = getSPWebNintexWorklowsInstanceDBCount -siteURL $siteURL
		if ($Count_WfInstance -ne -1)
		{
			Write-Host "Workflow Instance (DB - Nintex) Count: $Count_WfInstance"  -Foregroundcolor Green `r
		}
		else
		{
			Write-Host "Something went wrong with getNintexWorklowsInstanceDBrowsCount()" -Foregroundcolor Red `r	
		}
	
		#####################################################
		# Workflow Progress (DB - Nintex) on this site
		#####################################################
		$Count_WfProgress = getSPWebNintexWorklowsProgressDBCount -siteURL $siteURL
		if ($Count_WfProgress -ne -1)
		{
			Write-Host "Workflow Progress (DB - Nintex) Count: $Count_WfProgress"  -Foregroundcolor Green `r
		}
		else
		{
			Write-Host "Something went wrong with getNintexWorklowsProgressDBCount()" -Foregroundcolor Red `r	
		}
	}
	else
	{
		Write-Warning "Site '$siteURL' does not exist."
	}
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
	if($web -ne $null) 
	{
		$web.Dispose();
	}
	Write-Host ""
	Write-Host "#**************************************************************#" -Foregroundcolor Cyan `r
	Write-Host "# $scriptName ended" -ForegroundColor Cyan `r
	Write-Host "#**************************************************************#" -Foregroundcolor Cyan `r
	
	Stop-SPAssignment -Global 
}



