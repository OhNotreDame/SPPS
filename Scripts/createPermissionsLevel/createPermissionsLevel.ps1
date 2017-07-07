<#

.SYNOPSIS
	Create all  Permissions Levels on the site siteURL
	
.DESCRIPTION
	Create all PROD Permissions Levels on the site siteURL
	
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
	[string]$siteURL
)
Clear-Host
Remove-Module *

################################################################
# Setting Global Variables
################################################################
$ModuleFolderPath = "D:\QuickDeployFW\Modules"
$scriptName = $([System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Definition))
Write-Host "******************************************************************************" -ForegroundColor Magenta `r
Write-Host "$scriptName # Script started."
Write-Host "******************************************************************************" -ForegroundColor Magenta `r		
Write-Host "siteURL : $siteURL"
Write-Host "ModuleFolderPath : $ModuleFolderPath" 
Write-Host "******************************************************************************" -ForegroundColor Magenta `r			

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


################################################################
# Loading QuickDeploy Custom Modules
################################################################
Import-Module "$ModuleFolderPath\SPHelpers\SPHelpers.psm1"
Import-Module "$ModuleFolderPath\SPSiteGroups\SPSiteGroups.psm1"
Import-Module "$ModuleFolderPath\SPSitePermissions\SPSitePermissions.psm1"
Import-Module "$ModuleFolderPath\SPPermissionsLevels\SPPermissionsLevels.psm1"


########################
####### SETTINGS #######
########################
$logsFolderName = $PSScriptRoot+"\Logs\"
try 
{
	
	#### PRE-REQUISITES ####
	if(!(Test-Path $logsFolderName))
	{
		New-Item $logsFolderName -type Directory -Force | Out-Null
	}
	
	#### TRANSCRIPT / LOGS ####
	$logsFileName = $logsFolderName + $scriptName  + "_"+ $(get-date -format 'yyyyMMdd_HHmmss') + ".log"
	$logsFileName
	Start-Transcript -path $logsFileName -noclobber -Append

	#### CREATE PERM.LEVEL ####	
	$sitePermLevelXMLFilePath    = "$scriptdir\createPermissionsLevel.xml"
	if(Test-Path $sitePermLevelXMLFilePath)
	{
		$permissionsLevelsXML = LoadXMLFile -xmlPath $sitePermLevelXMLFilePath
		if($permissionsLevelsXML -ne $null -and $permissionsLevelsXML.HasChildNodes)
		{
			browseAndParseSPPermissionsLevelXML -siteURL $siteURL -permissionsLevelsXML $permissionsLevelsXML
		}  
		else
		{
			Write-Warning "XML File for <SPSitePermissions> is empty." 
		}
		  
	}
	else
	{
		Write-Host "XML File for <SPSitePermissions> does not exist." -ForegroundColor Cyan `r
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
	Write-Host "`r"
	Write-Host "******************************************************************************" -ForegroundColor Cyan `r		
	Write-Host "$scriptName # Script ended." -ForegroundColor Cyan `r
	Write-Host "******************************************************************************" -ForegroundColor Cyan `r					
	
	#####################################################
	# Stopping SPAssignment and Transcript
	#####################################################
	Stop-SPAssignment -Global
	Stop-Transcript
}