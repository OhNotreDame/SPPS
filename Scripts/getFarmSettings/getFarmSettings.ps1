<#

.SYNOPSIS
	Get all Farm Settings and generate dedicated CSV files.
	
.DESCRIPTION
	Script used to collect general and specific information about the SharePoint Farm	
	Get and Export all Farm Settings/Details (one file per Setting)
		
.PARAMETER 
	None
	
.EXAMPLE
	GetFarmSettings.ps1 
	
.OUTPUTS
	Two folders (Logs\ and CSV\)
	1 log file [NC: GetFarmSettings_<MachineName>_<SPFarmSettingsName>.log]
	1 CSV file by Farm Setting: (Semicolon delimited) [NC: GetFarmSettings_<MachineName>_<SPFarmSettingsName>.csv]

.LINK
	None
	
.NOTES
	Source: https://github.com/OhNotreDame/SPPS
	Created by: JBO
	Created: 11.01.2017
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
	$csvFileNameRoot = $csvFolderName + $scriptName + "_"+ $serverName  + "_" 
	
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
	Write-Host "# 0. Farm Settings"
	Write-Host "#-------------------------------------------#" `r 
	Write-Host "> SP Farm Settings" -Foregroundcolor Yellow -NoNewline
	$csvFileNameFarm = $csvFileNameRoot + "SPFarm.csv"
	$farm = Get-SPFarm 
	$farm | Select DisplayName, Name, Id, Status, Version, BuildVersion,
				@{Name="TimerService";Expression={$_.TimerService.Name}} | Export-CSV $csvFileNameFarm -NoTypeInformation -Delimiter ";"
	Write-Host -f Green " [Done]"
	
	Write-Host "> SP Servers" -Foregroundcolor Yellow -NoNewline
	$csvFileNameServers = $csvFileNameRoot + "SPFarmServers.csv"
	
	#$serversList= @() <=> CSV Content for SPSite/SPWeb list
	$serversList = @()
	
	foreach ($svr in $farm.Servers) {
		$infoServer = New-Object PSObject
		Add-Member -inputObject $infoServer -memberType NoteProperty -name "ServerName" -value $svr.DisplayName
		Add-Member -inputObject $infoServer -memberType NoteProperty -name "ServerRole" -value $svr.Role
		Add-Member -inputObject $infoServer -memberType NoteProperty -name "ServerStatus" -value $svr.Status

		# Append infoWeb to WebAppMetrics
		$serversList += $infoServer
	}
	# Export the SPServers results in csv file.
	$serversList | Export-CSV -Path $csvFileNameServers -NoTypeInformation -Delimiter ";"
	Write-Host -f Green " [Done]"

	Write-Host "#-------------------------------------------#" `r 
	Write-Host "# 1. Service Accounts and other identities"
	Write-Host "#-------------------------------------------#" `r 
	
	Write-Host "> SP Process Accounts" -Foregroundcolor Yellow -NoNewline
	$csvFileNameProcAcc = $csvFileNameRoot + "SPProcessAccounts.csv"
	Get-SPProcessAccount | select Name | Export-CSV $csvFileNameProcAcc -NoTypeInformation -Delimiter ";"
	Write-Host -f Green " [Done]"
	
	Write-Host "> SP Managed Accounts" -Foregroundcolor Yellow -NoNewline
	$csvFileNameMngdAcc = $csvFileNameRoot + "SPManagedAccounts.csv"
	Get-SPManagedAccount | select TypeName, UserName, DiplayName, Sid, AutomaticChange, Name, Id, Status | Export-CSV $csvFileNameMngdAcc -NoTypeInformation -Delimiter ";" 
	Write-Host -f Green " [Done]"
	
	Write-Host "#-------------------------------------------#" `r 
	Write-Host "# 2. Databases"
	Write-Host "#-------------------------------------------#" `r 
	
	Write-Host "> SP Database" -Foregroundcolor Yellow -NoNewline
	$csvFileNameDB = $csvFileNameRoot + "SPDatabases.csv"
	Get-SPDatabase | select DisplayName, Server, ID, Status| Export-CSV $csvFileNameDB -NoTypeInformation -Delimiter ";"
	Write-Host -f Green " [Done]"
	
	Write-Host "> SP Content Database" -Foregroundcolor Yellow -NoNewline
	$csvFileNameContentDB = $csvFileNameRoot + "SPContentDB.csv"
	Get-SPContentDatabase | select 	Name, DisplayName, Id, Server, PreferredTimerServiceInstance, 
									CurrentSiteCount, WarningSiteCount, MaximumSiteCount, 
									@{Name="SPWebApplication";Expression={$_.SPWebApplication.Name}} , 
									Exists, IsReadOnly, DiskSizeRequired, NeedsUpgradeIncludeChildren, 
									NeedsUpgrade, Status, Version | Export-CSV $csvFileNameContentDB -NoTypeInformation -Delimiter ";" 
	Write-Host -f Green " [Done]"
	
	Write-Host "#-------------------------------------------#" `r 
	Write-Host "# 3. Services & Service Applications"
	Write-Host "#-------------------------------------------#" `r 
	
	Write-Host "> SP Services Instances" -Foregroundcolor Yellow -NoNewline
	$csvFileSrvInstances = $csvFileNameRoot + "SPServiceInstances.csv"
	Get-SPServiceInstance | select Service, TypeName, Id, 
							@{Name="ServerName";Expression={$_.Server.Name}}, 
							NeedsUpgradeIncludeChildren, NeedsUpgrade, Status, 
							Version | Export-CSV $csvFileSrvInstances -NoTypeInformation -Delimiter ";" 
	Write-Host -f Green " [Done]"
	 
	Write-Host "> SP Services Applications" -Foregroundcolor Yellow -NoNewline
	$csvFileSrvApp = $csvFileNameRoot + "SPServiceApplications.csv"
	Get-SPServiceApplication | select DisplayName, Name, TypeName, Id, Service, 
	IisVirtualDirectoryPath, 
	@{Name="ApplicationPool";Expression={$_.ApplicationPool.Name}},
	@{Name="DefaultEndpoint";Expression={$_.DefaultEndpoint.Name}},
	Version | Export-CSV $csvFileSrvApp -NoTypeInformation -Delimiter ";" 
	Write-Host -f Green " [Done]"
	
	Write-Host "> SP Services Application Pools" -Foregroundcolor Yellow -NoNewline
	$csvFileSrvAppPool = $csvFileNameRoot + "SPServiceAppPools.csv"
	Get-SPServiceApplicationPool | select Id, Name, ProcessAccountName | Export-CSV $csvFileSrvAppPool -NoTypeInformation -Delimiter ";" 
	Write-Host -f Green " [Done]"
	
	Write-Host "#-------------------------------------------#" `r 
	Write-Host "# 4. Farm Solutions & Features"
	Write-Host "#-------------------------------------------#" `r 
	
	Write-Host "> SP Farm Solution" -Foregroundcolor Yellow -NoNewline
	$csvFileNameFarmSol = $csvFileNameRoot + "SPFarmSolutions.csv"
	Get-SPSolution | select SolutionId, Name, DisplayName, Deployed, DeploymentState, ContainsGlobalAssembly, ContainsWebApplicationResource, JobExists, Version  | Export-CSV $csvFileNameFarmSol -NoTypeInformation -Delimiter ";" 
	Write-Host -f Green " [Done]"
	
	Write-Host "> SP Features" -Foregroundcolor Yellow -NoNewline
	$csvFileNameFeatures = $csvFileNameRoot + "SPFeatures.csv"
	Get-SPFeature | select SolutionId, Name, DisplayName, Scope, RootDirectory, Status | Sort -Property Scope,DisplayName | Export-CSV $csvFileNameFeatures -NoTypeInformation -Delimiter ";" 
	Write-Host -f Green " [Done]"
	
	Write-Host "> WebTemplates available" -Foregroundcolor Yellow -NoNewline
	$csvFileNameWebTemplates = $csvFileNameRoot + "SPWebTemplates.csv"
	Get-SPWebTemplate | select ID, Title, Name, Description, Lcid, IsCustomTemplate, IsRootWebOnly, IsSubWebOnly | Export-CSV $csvFileNameWebTemplates -NoTypeInformation -Delimiter ";" 
	Write-Host -f Green " [Done]"
	
	Write-Host "#-------------------------------------------#" `r 
	Write-Host "# 5. Web Applications & URLs"
	Write-Host "#-------------------------------------------#" `r 
	
	Write-Host "> SP Alternate URLs" -Foregroundcolor Yellow -NoNewline
	$csvFileNameAltURLs = $csvFileNameRoot + "SPAlternateURL.csv"
	Get-SPAlternateUrl | select IncomingUrl, UrlZone, PublicUrl | Export-CSV $csvFileNameAltURLs -NoTypeInformation -Delimiter ";" 
	Write-Host -f Green " [Done]"
	
	Write-Host "> SP Web Applications" -Foregroundcolor Yellow -NoNewline
	$csvFileNameWebApps = $csvFileNameRoot + "SPWebApps.csv"
	$webAppCollection = Get-SPWebApplication -IncludeCentralAdministration 
	$webAppCollection | select 	DisplayName, Id, Url, Status, 
								@{Name="ApplicationPoolName";Expression={$_.ApplicationPool.Name}},
								@{Name="ApplicationPoolIdentityName";Expression={$_.ApplicationPool.UserName}}, 
								@{Name="HostHeader";Expression={$_.IisSettings["Default"].ServerBindings.HostHeader}}, 
								@{Name="Port";Expression={$_.IisSettings["Default"].ServerBindings.Port}}, 
								DefaultQuotaTemplate, AllowDesigner, 
								MaximumFileSize, MaxItemsPerThrottledOperation, MaxItemsPerThrottledOperationOverride, 
								RecycleBinCleanupEnabled, RecycleBinEnabled, RecycleBinRetentionPeriod, 
								UnusedSiteNotificationPeriod, UnusedSiteNotificationsBeforeDeletion | Export-CSV $csvFileNameWebApps -NoTypeInformation -Delimiter ";" 
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