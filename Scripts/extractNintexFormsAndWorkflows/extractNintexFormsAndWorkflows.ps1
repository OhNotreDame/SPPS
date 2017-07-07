<#

.SYNOPSIS
	Extract Nintex Forms and Workflows and generate XML files to be capable of redeploying them later
	
.DESCRIPTION
	Extract Nintex Forms and Workflows and generate XML files to be capable of redeploying them later
	
.PARAMETER siteUrl
	(Mandatory) URL of the SharePoint Site
	
.EXAMPLE
	CreateIntranetPermissionsLevel.ps1 -siteUrl <siteURL>
	
.OUTPUTS
	None

.LINK
	https://github.com/OhNotreDame/SPPS
	
.NOTES
	Created by: JBO
	Created: 27.03.2017
	Last Updated by: JBO
	Last Updated: 27.03.2017
	
#>

param
(
 [Parameter(Mandatory=$true, Position=1)]
 [string]$siteURL,
 [Parameter(Mandatory=$false, Position=2)]
 [string]$destinationFolderPath
)

Clear-Host
Remove-Module *

$scriptName = $([System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Definition))
Write-Host "******************************************************************************" -ForegroundColor Gray `r
Write-Host "$scriptName # Script started." -ForegroundColor Gray `r
Write-Host "******************************************************************************" -ForegroundColor Gray `r		

#####################################################
# Loading SharePoint Assembly and PS Snapin
#####################################################
Add-PsSnapin Microsoft.Sharepoint.PowerShell -EA silentlycontinue
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint") | Out-Null

#####################################################
# Starting SPAssignment
#####################################################
Start-SPAssignment -Global

#####################################################
# Setting Path Variables
#####################################################
$scriptdir = $PSScriptRoot
Set-Variable -Name "scriptPath" -Value $scriptdir -Scope Global




try
{
	if([string]::IsNullOrEmpty($destinationFolderPath)) 
	{
		$destinationFolderPath = Get-Location
		Write-Warning "[$functionName] Paramater destinationFolderPath is empty, will set it to the current location." 
		Write-Host "[$functionName] destinationFolderPath: $destinationFolderPath" -foregroundcolor Cyan
	}
	
	$ModuleFolderPath = "D:\QuickDeployFW\Modules"
	$destFolderArtefacts = "$destinationFolderPath\Artefacts"
	$destFolderLogs = "$destinationFolderPath\Logs"
	$destFolderWF = "$destinationFolderPath\NintexWorkflows" 
	$destFolderForms = "$destinationFolderPath\NintexForms" 

	Write-Host "" `r
	Write-Host "******************************************************************************" -ForegroundColor Magenta `r
	Write-Host "$scriptName # Parameters and Settings" -ForegroundColor Magenta `r
	Write-Host "******************************************************************************" -ForegroundColor Magenta `r		
	Write-Host "siteURL: $siteURL" -ForegroundColor Gray `r
	Write-Host "solutionFolderPath: $solutionFolderPath" -ForegroundColor Gray `r
	Write-Host "configFilePath: $configFilePath" -ForegroundColor Gray `r
	Write-Host "artefactFolderPath: $artefactFolderPath" -ForegroundColor Gray `r
	Write-Host "defaultConfigFolderPath: $defaultConfigFolderPath" -ForegroundColor Gray `r
	Write-Host "ModuleFolderPath: $ModuleFolderPath" -ForegroundColor Gray `r
	Write-Host "scriptdir: $scriptdir" -ForegroundColor Gray `r
	Write-Host "******************************************************************************" -ForegroundColor Magenta `r	
	
	#####################################################
	# Loading QuickDeploy Custom Modules
	#####################################################
	Write-Host "" `r
	Write-Host "******************************************************************************" -ForegroundColor Magenta `r
	Write-Host "About to import QuickDeploy Modules " -ForegroundColor Magenta `r
	Write-Host "******************************************************************************" -ForegroundColor Magenta `r
	Import-Module "$ModuleFolderPath\SPHelpers\SPHelpers.psm1"
	Import-Module "$ModuleFolderPath\SPFileUploader\SPFileUploader.psm1"
	Import-Module "$ModuleFolderPath\SPSiteColumns\SPSiteColumns.psm1"
	Import-Module "$ModuleFolderPath\SPSiteContentTypes\SPSiteContentTypes.psm1"
	Import-Module "$ModuleFolderPath\SPLists\SPLists.psm1"
	Import-Module "$ModuleFolderPath\NintexWorkflows\NintexWorkflows.psm1"
	Import-Module "$ModuleFolderPath\NintexForms\NintexForms.psm1"
	Write-Host "QuickDeploy Modules Successfully Imported" -ForegroundColor Green `r
	Write-Host "******************************************************************************" -ForegroundColor Magenta `r
	
	###########################
	#### CREATING FOLDERS  ####
	###########################
	if(!(Test-Path $destFolderLogs))
	{
		New-Item $destFolderLogs -type Directory -Force | Out-Null
	}
	if (!(Test-Path $destFolderArtefacts))
	{
		New-Item $destFolderArtefacts -type Directory -Force | Out-Null
	}
	if(!(Test-Path $destFolderWF))
	{
		New-Item $destFolderWF -type Directory -Force | Out-Null
	}
	if(!(Test-Path $destFolderForms))
	{
		New-Item $destFolderForms -type Directory -Force | Out-Null
	}
	
	
	###########################
	#### TRANSCRIPT / LOGS ####
	###########################
	$logsFileName = $destFolderLogs + "\" + $scriptName  + "_"+ $(get-date -format 'yyyyMMdd_HHmmss') + ".log"
	$logsFileName
	Start-Transcript -path $logsFileName -noclobber -Append
	
	$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
	if($curWeb -ne $null)
	{
		Write-Host "" `r
		Write-Host "******************************************************************************" -ForegroundColor Magenta `r
		Write-Host "About to extract Nintex Workflows " -ForegroundColor Magenta `r
		Write-Host "******************************************************************************" -ForegroundColor Magenta `r
		ExportAllWorkflowsFromSite -siteURL $siteURL -destFolder $destinationFolderPath
		
		Write-Host "" `r
		Write-Host "******************************************************************************" -ForegroundColor Magenta `r
		Write-Host "About to extract Nintex Forms " -ForegroundColor Magenta `r
		Write-Host "******************************************************************************" -ForegroundColor Magenta `r
		ExportAllActiveFormsFromSite -siteURL $siteURL -destFolder $destinationFolderPath
	
	}
	else
	{
		Write-Warning "Site '$siteURL' does not exist."
	}		
}
catch [Exception]
{	
	Write-Host "/!\ $scriptName An exception has been caught /!\ "  -ForegroundColor Red `r
	Write-Host "Type: " $_.Exception.GetType().FullName   -ForegroundColor Red `r
	Write-Host "Message: " $_.Exception.Message  -ForegroundColor Red `r
	Write-Host "Stacktrace: " $_.Exception.Stacktrace  -ForegroundColor Red `r
}
finally
{
	if($curWeb -ne $null)
	{
		$curWeb.Dispose()
	}

	Write-Host "******************************************************************************" -ForegroundColor Gray `r		
	Write-Host "$scriptName # Script ended." -ForegroundColor Gray `r
	Write-Host "******************************************************************************" -ForegroundColor Gray `r					
	
	#####################################################
	# Stopping SPAssignment and Transcript
	#####################################################
	Stop-SPAssignment -Global
	Stop-Transcript
}