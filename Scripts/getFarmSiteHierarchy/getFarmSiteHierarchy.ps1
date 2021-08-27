<#

.SYNOPSIS
	Get all Site Collections and Subsites available on the Farm and generate dedicated CSV file.
	
.DESCRIPTION
	Script used to collect general and specific information about the SharePoint Farm	
	Mainly focused on :
	1. Get and Export all Farm Settings/Details (one file per Setting)
	2. Get all SPSite and SPWeb objects of all SPWebApplications and Export some info into a CSV file (one file per WebApplication).
	
.PARAMETER 
	None
	
.EXAMPLE
	GetFarmSiteHierarchy.ps1 
	
.OUTPUTS
	Two folders (Logs\ and CSV\)
	1 log file [NC: GetFarmSiteHierarchy_<MachineName>.log]
	1 CSV file by Web Application: (Semicolon delimited) [NC: GetFarmSiteHierarchy_<MachineName>_<WebAppHost>_<WebAppPort>.csv]

.LINK
	None
	
.NOTES
	Source: https://github.com/OhNotreDame/SPPS
	Created by: JBO
	Created: 13.01.2017
	Last Updated by: JBO
	Last Updated: 13.01.2017
	
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




$webAppURI = [System.Uri] $webAppURL
$webAppHost = $webAppURI.Host



try 
{
	$serverName = $env:computername
	
	# $execDate = Get-Date -format "yyyyMMddHHmmss"
	# $logsFileName = $logsFolderName + $scriptName + "_" + $execDate  + "_"+ $serverName + ".log"
	# $csvFileNameRoot = $csvFolderName + $scriptName + "_" + $execDate  + "_Farm_" + $serverName
	
	$logsFileName = $logsFolderName + $scriptName  + "_"+ $serverName + ".log"
	$csvFileNameSPSiteRoot = $csvFolderName + $scriptName + "_"+ $serverName  + "_" 
	
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
	Add-PSSnapin Microsoft.SharePoint.PowerShell -EA silentlycontinue
	
	Start-SPAssignment -Global
	
	Start-Transcript -path $logsFileName | Out-Null
	
	Write-Host "#**************************************************************#" -Foregroundcolor Cyan `r
	Write-Host "# Script $scriptName started" -Foregroundcolor Cyan `r  
	Write-Host "#**************************************************************#" -Foregroundcolor Cyan `r 
	Write-Host "# Transcript file: `n$logsFileName"  -ForegroundColor Cyan `r  
	Write-Host "#**************************************************************#" -Foregroundcolor Cyan `r
	Write-Host ""
		
	Write-Host "#-------------------------------------------#" `r 
	Write-Host "# Site Collections and Subsites"
	Write-Host "#-------------------------------------------#" `r 
	
	#$SPWebList = @() <=> CSV Content for SPSite/SPWeb list
	$SPWebList = @()
	
	$webAppCollection = Get-SPWebApplication 
	foreach ($webApp in $webAppCollection)
	{
		#$webApp.URL
		Write-Host "`nWeb Application:  $($webApp.URL)" -Foregroundcolor Magenta
		$SPWebList = @()
		
		$webAppURI = [System.Uri] $webApp.URL

		$webAppHost = $webAppURI.Host + "_" + $webAppURI.Port
		$csvFileName = $csvFileNameSPSiteRoot + $webAppHost +  ".csv"
	
		foreach($site in $webApp.Sites)
		{
			Write-Host "> Browsing site collection $($site.URL)" -Foregroundcolor Gray -noNewLine
		
			foreach ($web in $site.AllWebs)
			{
		
				$siteName = $web.Title
				$siteURL = $web.URL
				$siteTemplate = $web.WebTemplate
				$siteIsRootWeb = $web.IsRootWeb
				$siteCreated = $web.Created
				$siteLastModified = $web.LastItemModifiedDate
				$isPublishingWeb= [Microsoft.SharePoint.Publishing.PublishingWeb]::IsPublishingWeb($web)
				$siteHasUniquePerm = $web.HasUniqueRoleAssignments
				
				$sitePrimaryOwner = ""
				$siteSecondaryOwner = ""
				$siteAdministrators = ""
				$siteContentDB = ""
				$siteContentDB = ""
				$siteSize = ""
				$hitsMensuel = ""
				$uniqUsersMensuel = ""
					
				if ($web.Url -eq $site.URL)
				{
					$siteContentDB = $($site.ContentDatabase.Name)
					$siteSize = $site.Usage.Storage/1GB
					
					$sitePrimaryOwner = $site.Owner.UserLogin
					$siteSecondaryOwner = $site.SecondaryContact.UserLogin
	
					$siteAdmins = $site.RootWeb.SiteAdministrators
					$siteAdministrators = "";
					foreach ($scAdmin in $siteAdmins)	{
						$siteAdministrators += $scAdmin.UserLogin + "," ;
					}
					$siteAdministrators = $siteAdministrators -replace ",$"
				}
				
				#Prepare CSV Structure
				$infoWeb = New-Object PSObject
				Add-Member -inputObject $infoWeb -memberType NoteProperty -name "SiteName" -value $siteName
				Add-Member -inputObject $infoWeb -memberType NoteProperty -name "SiteURL" -value $siteURL				
				Add-Member -inputObject $infoWeb -memberType NoteProperty -name "isRootWeb" -value $siteIsRootWeb
				Add-Member -inputObject $infoWeb -memberType NoteProperty -name "isPublishingWeb" -value $isPublishingWeb
				Add-Member -inputObject $infoWeb -memberType NoteProperty -name "SiteCollectionAdmin" -value $siteAdministrators
				Add-Member -inputObject $infoWeb -memberType NoteProperty -name "PrimaryOwner" -value $sitePrimaryOwner
				Add-Member -inputObject $infoWeb -memberType NoteProperty -name "SecondaryOwner" -value $siteSecondaryOwner
				Add-Member -inputObject $infoWeb -memberType NoteProperty -name "SiteTemplate" -value $siteTemplate
				Add-Member -inputObject $infoWeb -memberType NoteProperty -name "hasUniqPerms" -value $siteHasUniquePerm
				Add-Member -inputObject $infoWeb -memberType NoteProperty -name "ContentDB" -value $siteContentDB
				Add-Member -inputObject $infoWeb -memberType NoteProperty -name "Size" -value $siteSize
				Add-Member -inputObject $infoWeb -memberType NoteProperty -name "Created" -value $siteCreated
				Add-Member -inputObject $infoWeb -memberType NoteProperty -name "LastModified" -value $siteLastModified
					
				# Append infoWeb to SPWebList
				$SPWebList += $infoWeb

				$web.Dispose()
				
			} #end foreach AllWebs
			
			$site.Dispose()
			Write-Host -f Green " [Done]"
		} #end foreach Sites
		
		# Export the SPSite/SPWeb results in csv file.
		$SPWebList | Export-CSV -Path $csvFileName -NoTypeInformation -Delimiter ";" 

	} #end foreach webApp

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