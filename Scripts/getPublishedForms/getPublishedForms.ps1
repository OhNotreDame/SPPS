<#

.SYNOPSIS
	Get all Published Forms and generate dedicated CSV files
	
.DESCRIPTION
	Script used to collect general information about the Published Forms deployed on the SharePoint Farm	
	
.PARAMETER 
	None
	
.EXAMPLE
	getPublishedForms.ps1
	
.OUTPUTS
	Two folders (Logs\ and CSV\)
	1 log file [NC: GetPublishedForms_<MachineName>.log]
	for all Published Forms: 1 CSV file (Semicolon delimited) [NC: GetPublishedForms_<MachineName>.csv]		
	
.LINK
	https://community.nintex.com/community/build-your-own/blog/2015/06/09/locate-nintex-forms-in-farm-via-powershell
	
.NOTES
	Source: https://github.com/OhNotreDame/SPPS
	Created by: JBO
	Created: 12.01.2017
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
	$csvFormsFileName = $csvFolderName + $scriptName + "_" + $serverName + ".csv"
	
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
	Write-Host "# CSV file: `n$csvFormsFileName"  -ForegroundColor Cyan `r  
	Write-Host "#**************************************************************#" -Foregroundcolor Cyan `r
	Write-Host ""
	

	Write-Host "#**************************************************************#" -Foregroundcolor Yellow `r
	Write-Host "# Published Forms" -Foregroundcolor Yellow `r
	Write-Host "#**************************************************************#" -Foregroundcolor Yellow `r
	
	#CSV Content for Forms list
	$formsList = @()
	
	
	$webAppCollection = Get-SPWebApplication
	foreach ($webApp in $webAppCollection)
	{
		
		if (!$webApp.IsAdministrationWebApplication )
		{
			#$webApp.URL
			Write-Host "`nWeb Application:  $($webApp.URL)" -Foregroundcolor Magenta
			
			foreach($site in $webApp.Sites)
			{
				Write-Host "> Site collection $($site.URL)" -Foregroundcolor Gray -NoNewLine
			
				foreach ($web in $site.AllWebs)
				{
					$siteName = $web.Title
					$siteURL = $web.URL
					Write-Verbose "$siteName  - $siteURL "
					$nintexFormsList = $web.Lists.TryGetList("NintexForms")
					
					if ($nintexFormsList -ne $null)
					{
						$nintexFormsListURL = $siteURL + "/" + $nintexFormsList.DefaultView.Url
						Write-Verbose "nintexFormsListURL: $nintexFormsListURL"
						
						foreach ($listItem in $nintexFormsList.Items)
						{
							
							# Form general information from SPListItem object
							$formName= $listItem.Name
							$formCreated = $listItem["Created"];
							$formModified = $listItem["Modified"];
							$formHasPublishedVersion = $listItem.HasPublishedVersion
							$formStatus = $listItem.Level
							$formVersion = $listItem["_UIVersionString"];
							$formLastVersion = $null
							$parentListName = [string]::Empty
							$parentCTName = [string]::Empty
							
							Write-Verbose "Before Last Published Version (if any)"
							# Compute Last Published Version (if any)
							if ($formHasPublishedVersion -and ($formStatus -eq "Published"))
							{
								$formLastVersion = $formVersion							
							}
							
							if ($formHasPublishedVersion -and ($formStatus -eq "Draft"))
							{
								$versionDecimal = [decimal]	$listItem["_UIVersionString"]
								$formLastVersion = [int] [math]::truncate($versionDecimal) 
							}
							
							if ((!$formHasPublishedVersion) -and ($formStatus -eq "Draft"))
							{
								$formLastVersion = -1
							}
							
							Write-Verbose "Before Getting List"
							# Get Parent List and ContentType Names
							$parentListID = $listItem["FormListId"];
							if ($parentListID -ne $null)
							{
								$parentListID = [Guid] $parentListID.ToString()	
								$parentList = $web.Lists  | where {$_.ID -eq $parentListID} 
								if ($parentList -ne $null)
								{
									$parentListName = $parentList.Title
									$parentListItemCount = $parentList.Items.Count
									$parentListCreated = $parentList.Created
									$parentListLastItemModifiedDate = $parentList.LastItemModifiedDate
									$parentCTID = $listItem["FormContentTypeId"];

									if ($parentCTID -ne $null)
									{	
										$parentCT = $parentList.ContentTypes | where {$_.ID -eq $parentCTID} 
										if ($parentCT -ne $null)
										{
											$parentCTName = $parentCT.Name
										}
										else
										{
											Write-Verbose "Warning: Content-Type '$parentCTID' does not exist anymore on List '$listName'.";
										}
										
										#Prepare CSV Structure
										$infoForms = New-Object PSObject
										Add-Member -inputObject $infoForms -memberType NoteProperty -name "SiteName" -value $siteName
										Add-Member -inputObject $infoForms -memberType NoteProperty -name "SiteURL" -value $siteURL
										Add-Member -inputObject $infoForms -memberType NoteProperty -name "List" -value $parentListName
										Add-Member -inputObject $infoForms -memberType NoteProperty -name "ListItemCount" -value $parentListItemCount
										Add-Member -inputObject $infoForms -memberType NoteProperty -name "ListCreated" -value $parentListCreated
										Add-Member -inputObject $infoForms -memberType NoteProperty -name "ListLastUpdated" -value $parentListLastItemModifiedDate
										Add-Member -inputObject $infoForms -memberType NoteProperty -name "Content-Type" -value $parentListName
										Add-Member -inputObject $infoForms -memberType NoteProperty -name "Name" -value $formName
										Add-Member -inputObject $infoForms -memberType NoteProperty -name "Created" -value $formCreated
										Add-Member -inputObject $infoForms -memberType NoteProperty -name "Modified" -value $formModified
										Add-Member -inputObject $infoForms -memberType NoteProperty -name "Status" -value $formStatus
										Add-Member -inputObject $infoForms -memberType NoteProperty -name "Version" -value $formVersion
										Add-Member -inputObject $infoForms -memberType NoteProperty -name "HasPublishedVersion" -value $formHasPublishedVersion
										Add-Member -inputObject $infoForms -memberType NoteProperty -name "LastPublishedVersion" -value $formLastVersion
																			
										#Append infoWeb to formsList
										$formsList += $infoForms	
										
										
									}
									else
									{
										Write-Verbose "Warning: 'FormContentTypeId' is empty.";
									}
										
								}
								else
								{
									Write-Verbose "Warning: List '$FormListId' does not exist anymore.";
								}
								
							}
							else
							{
								Write-Verbose "Warning: 'FormListId' is empty.";
							}
							
							

						}#end foreach items						

					}
					else
					{
						Write-Verbose "Warning: No NintexForms list on $siteURL."
					}
					$web.Dispose()
				} #end foreach AllWebs
				
				$site.Dispose()
				Write-Host -f Green " [Done]"
				
			} #end foreach Sites
			
			
		}
		else
		{
			Write-Verbose "Warning: Central Administration ignored."
		}
	} #end foreach webApp
	
	# Export the SPSite/SPWeb results in csv file.
	$formsList | Export-CSV -Path $csvFormsFileName -NoTypeInformation -Delimiter ";" -Encoding "UTF8"

}
catch
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
	
	Write-Host "#**************************************************************#" -Foregroundcolor Cyan `r
	Write-Host "# Script $scriptName ended" -ForegroundColor Cyan `r
	Write-Host "#**************************************************************#" -Foregroundcolor Cyan `r
	
	Stop-SPAssignment -Global 
	Stop-Transcript | Out-Null
}