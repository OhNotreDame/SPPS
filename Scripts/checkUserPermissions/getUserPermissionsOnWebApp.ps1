<#

.SYNOPSIS
	Get all the permissions for all the securable objects (SPSite, SPWeb, SPList, SPFolder or SPListItem) for a specific User on a specific SPWebApplication (aka Web Application), and generate CSV files with the results.
	
.DESCRIPTION
	Get all the permissions for all the securable objects (SPSite, SPWeb, SPList, SPFolder or SPListItem) for a specific User on a specific SPWebApplication (aka Web Application), and generate CSV files with the results.
	By default, will only look after WindowsClaims User permissions.
	If -notClaims is used, will also look after Windows User permissions.
	
.PARAMETER siteUrl
	(Mandatory) URL of the SharePoint Site
	
.PARAMETER userName
	(Mandatory) userName of the User to check permissions 

.PARAMETER notClaims
	(Optional) Switch to generate report for non-Claims user
	
.EXAMPLE
	getUserPermissionsOnWebApp.ps1 -siteUrl <siteURL> -userName <userName> -notClaims
	
.OUTPUTS
	One file for Claims User Permissions
	if -notClaims, One additional file for Windows User Permissions 

.LINK
	http://www.sharepointdiary.com/2013/01/permission-report-for-specific-user.html
	https://gallery.technet.microsoft.com/scriptcenter/SharePoint-Permission-2840f327
	
.NOTES
	Source:  https://github.com/OhNotreDame/SPPS
	Created by: JBO
	Created: 28.02.2017
	Last Updated by: JBO
	Last Updated: 01.03.2017
	
#>
param
(
	[Parameter(Mandatory=$true, Position=1)]
	[string]$siteURL,
	[Parameter(Mandatory=$false, Position=2)]
	[ValidateScript({
			if ($_.StartsWith("domain\", $true, $null))
			{
				$true
			}
			else {
				throw "`r`n$_ is not a valid username.`nPlease use following format <domain\userName>."
			}
		})]
	[string]$userName,
	[Parameter(Mandatory=$false, Position=3)]
	[switch]$notClaims
)


$webAppURI = [System.Uri] $siteURL
$computedSiteName =  $webAppURI.Host +"_" +$webAppURI.Port

Clear-Host
Remove-Module *

#####################################################
# Loading SharePoint Assembly and PS Snapin
#####################################################
Add-PsSnapin Microsoft.Sharepoint.PowerShell -EA silentlycontinue
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint") | Out-Null
Import-Module "$PSScriptRoot\CheckUserPermissionsCommon.psm1"
Start-SPAssignment -Global

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
	
	#### PRE-REQUISITES ####
	if(!(Test-Path $logsFolderName))
	{
		New-Item $logsFolderName -type Directory -Force | Out-Null
	}
	
	if(!(Test-Path $csvFolderName))
	{
		New-Item $csvFolderName -type Directory -Force | Out-Null
	} 

	#### PARSING USER NAME ####
	$userName = $userName.ToLower()
	$userNameClaims = "i:0#.w|$userName"
	
	#### TRANSCRIPT / LOGS ####
	$logsFileName = $logsFolderName + $scriptName  + "_"+ $(get-date -format 'yyyyMMdd_HHmmss') + ".log"
	$logsFileName
	Start-Transcript -path $logsFileName -noclobber -Append

	#### CSV FILE PATHs ####
	$pos = $userName.IndexOf("\")
	$userNameForCSVfile = $userName.Substring($pos+1)
	$csvFileNameClaims = $csvFolderName + $scriptName  + "_"+ $computedSiteName + "_" + $userNameForCSVfile + "_Claims" +  ".csv"
	$csvFileNameWindows = $csvFolderName + $scriptName + "_"+ $computedSiteName + "_" + $userNameForCSVfile + "_Windows" +  ".csv"
	
	Write-Host "******************************************************************************" -ForegroundColor Cyan `r		
	Write-Host "# Script $scriptName started" -Foregroundcolor Cyan `r  
	Write-Host "# siteURL:" -Foregroundcolor Cyan  `r
	Write-Host "# $siteURL" -Foregroundcolor Cyan  `r
	Write-Host "******************************************************************************" -ForegroundColor Cyan `r		
	Write-Host "`r"
	
	
	$webApp = Get-SPWebApplication $siteURL
	if($webApp -ne $null)
	{

		#### WINDOWS USER ####
		if ($notClaims)
		{
			$permList = @()
			Write-Host "#**************************************************************#" -Foregroundcolor Magenta `r
			Write-Host "Permissions for User: $userName" -Foregroundcolor Magenta `r
			Write-Host "#**************************************************************#" -Foregroundcolor Magenta `r
			
			#Step 1: WebApp Policies
			$permList += checkWebAppUserPolicy -webAppURL $siteURL -userName $userName
			
			#Step 2: Site Collection and Subsites
			$permList += getUserPermissionsOnSPWebApplication -webAppURL $siteURL -userName $userName
			
			#Step 3: Export
			$permList | Export-CSV -Path $csvFileNameWindows -NoTypeInformation -Delimiter ";" -Encoding UTF8
			Write-Host "`r"
		}
		
		#### CLAIMS USER ####
		$permList = @()
		Write-Host "#**************************************************************#" -Foregroundcolor Magenta `r
		Write-Host "Permissions for User:  $userNameClaims" -Foregroundcolor Magenta `r
		Write-Host "#**************************************************************#" -Foregroundcolor Magenta `r
		#Step 1: WebApp Policies
		$permList += checkWebAppUserPolicy -webAppURL $siteURL -userName $userNameClaims
		
		#Step 2: Site Collection and Subsites
		$permList += getUserPermissionsOnSPWebApplication -webAppURL $siteURL -userName $userNameClaims
		
		#Step 3: Export
		$permList | Export-CSV -Path $csvFileNameClaims -NoTypeInformation -Delimiter ";" -Encoding UTF8
		
	
		
	}
	else
	{
		Write-Warning "SPWeb '$siteURL' does not exist."
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