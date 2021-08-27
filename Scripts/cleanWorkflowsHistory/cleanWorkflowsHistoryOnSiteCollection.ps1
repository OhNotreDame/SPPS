<#

.SYNOPSIS
	Clean Nintex Workflows History (SharePoint List & SQL Progress Table) on all SPWeb of a Site Collection
	
.DESCRIPTION
	Clean Nintex Workflows History (SharePoint List & SQL Progress Table) on all SPWeb of a Site Collection
	
.PARAMETER status
	Status of the Workflow Instance to be collected
	Should be 'Completed','Cancelled','Error'
	
.PARAMETER siteUrl
	(Mandatory) URL of the SharePoint Site (SPSite)
	
.EXAMPLE
	cleanWorkflowsHistoryOnSiteCollection.ps1 siteUrl <siteUrl> -status <'Completed','Cancelled','Error'>
	
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

Start-SPAssignment -Global

try 
{	
	#### PRE-REQUISITES ####
	if(!(Test-Path $logsFolderName))
	{
		New-Item $logsFolderName -type Directory -Force | Out-Null
	}
	
	#### TRANSCRIPT / LOGS ####
	$logsFileName = $logsFolderName + $scriptName  + "_"+ $(get-date -format 'yyyyMMdd_HHmmss') + ".log"
	Start-Transcript -path $logsFileName -noclobber -Append

	$site = Get-SPSite $siteURL -ErrorAction SilentlyContinue
	if($site -ne $null) 
	{
		foreach($subWeb in $site.AllWebs)		
		{
			.\CleanWorkflowsHistoryOnSite.ps1 -siteURL $subWeb.URL -status $status
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
	if($site -ne $null) 
	{
		$site.Dispose();
	}
	
	Write-Host ""
	Write-Host "#**************************************************************#" -Foregroundcolor Cyan `r
	Write-Host "# $scriptName ended" -ForegroundColor Cyan `r
	Write-Host "#**************************************************************#" -Foregroundcolor Cyan `r
	
	#####################################################
	# Stopping SPAssignment and Transcript
	#####################################################
	Stop-SPAssignment -Global
	Stop-Transcript
	
}



