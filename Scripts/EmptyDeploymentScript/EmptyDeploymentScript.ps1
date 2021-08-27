param
(
 [Parameter(Mandatory=$true, Position=1)]
 [string]$siteURL,
 [Parameter(Mandatory=$false, Position=2)]
 [string]$solutionFolderPath
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
	if([string]::IsNullOrEmpty($solutionFolderPath)) 
	{
		$solutionFolderPath = Get-Location
		Write-Warning "[$scriptName] Paramater solutionFolderPath is empty, will set it to the current location." 
		Write-Host "[$scriptName] solutionFolderPath: $solutionFolderPath" -foregroundcolor Cyan
	}
	
	$moduleFolderPath = "<PATH_TO_FRAMEWORK_ROOT_FOLDER>\Modules"
	$destFolderArtefacts = "$solutionFolderPath\Artefacts"
	$destFolderLogs = "$solutionFolderPath\Logs"
	$destFolderWF = "$solutionFolderPath\NintexWorkflows" 
	$destFolderForms = "$solutionFolderPath\NintexForms" 

	Write-Host "" `r
	Write-Host "******************************************************************************" -ForegroundColor Magenta `r
	Write-Host "$scriptName # Parameters and Settings" -ForegroundColor Magenta `r
	Write-Host "******************************************************************************" -ForegroundColor Magenta `r		
	Write-Host "siteURL: $siteURL" -ForegroundColor Gray `r
	Write-Host "solutionFolderPath: $solutionFolderPath" -ForegroundColor Gray `r
	Write-Host "destFolderArtefacts: $destFolderArtefacts" -ForegroundColor Gray `r
	Write-Host "moduleFolderPath: $moduleFolderPath" -ForegroundColor Gray `r
	Write-Host "scriptdir: $scriptdir" -ForegroundColor Gray `r
	Write-Host "******************************************************************************" -ForegroundColor Magenta `r	
		
	#####################################################
	# Loading SPPS Custom Modules
	#####################################################
	Write-Host "" `r
	Write-Host "******************************************************************************" -ForegroundColor Magenta `r
	Write-Host "About to import SPPS Modules " -ForegroundColor Magenta `r
	Write-Host "******************************************************************************" -ForegroundColor Magenta `r
	Import-Module "$ModuleFolderPath\SPHelpers\SPHelpers.psm1"
	Import-Module "$ModuleFolderPath\SPFileUploader\SPFileUploader.psm1"
	Import-Module "$ModuleFolderPath\SPCommonStructure\SPCommonStructure.psm1"
	Import-Module "$ModuleFolderPath\SPSiteColumns\SPSiteColumns.psm1"
	Import-Module "$ModuleFolderPath\SPSiteContentTypes\SPSiteContentTypes.psm1"
	Import-Module "$ModuleFolderPath\SPLists\SPLists.psm1"
	Import-Module "$ModuleFolderPath\SPListViews\SPListViews.psm1"
	Import-Module "$ModuleFolderPath\SPListColumnsTranslation\SPListColumnsTranslation.psm1"
	Import-Module "$ModuleFolderPath\NintexWorkflows\NintexWorkflows.psm1"
	Import-Module "$ModuleFolderPath\NintexForms\NintexForms.psm1"
	Import-Module "$ModuleFolderPath\SPSiteGroups\SPSiteGroups.psm1"
	Import-Module "$ModuleFolderPath\SPSitePermissions\SPSitePermissions.psm1"
	Import-Module "$ModuleFolderPath\SPListPermissions\SPListPermissions.psm1"
	Import-Module "$ModuleFolderPath\SPSitePages\SPSitePages.psm1"
	Write-Host "SPPS Modules Successfully Imported" -ForegroundColor Green `r
	Write-Host "******************************************************************************" -ForegroundColor Magenta `r
	
	###########################
	#### CREATING FOLDERS  ####
	###########################
	if(!(Test-Path $destFolderLogs))
	{
		New-Item $destFolderLogs -type Directory -Force | Out-Null
	}
	
	###########################
	#### TRANSCRIPT / LOGS ####
	###########################
	$logsFileName = $destFolderLogs + "\"+ $scriptName  + "_"+ $(get-date -format 'yyyyMMdd_HHmmss') + ".log"
	$logsFileName
	Start-Transcript -path $logsFileName -noclobber -Append
	
	$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
	if($curWeb -ne $null)
	{
		$parentURL = ""
		if ($curWeb.IsRootWeb)
		{
			$parentURL = $siteURL
		}
		else
		{
			$parentURL = $curWeb.site.RootWeb.URL
		}
		Write-Host "" `r
		Write-Host "******************************************************************************" -ForegroundColor Magenta `r
		Write-Host "About to activate Nintex Workflows Features " -ForegroundColor Magenta `r
		Write-Host "******************************************************************************" -ForegroundColor Magenta `r
		# Nintex Workflow 2013 / SPSite / NintexWorkflow / 0561d315-d5db-4736-929e-26da142812c5 
		Enable-SPFeature -Identity "0561d315-d5db-4736-929e-26da142812c5" -URL $parentURL -ErrorAction Stop -Force -Confirm:$False

		# Enabling Nintex Workflows SPWeb Features
		# Nintex Workflow / SPWeb / NintexWorkflowWeb / 9bf7bf98-5660-498a-9399-bc656a61ed5d
		Enable-SPFeature -Identity "9bf7bf98-5660-498a-9399-bc656a61ed5d" -URL $siteUrl -ErrorAction Stop -Force -Confirm:$False
		
		Write-Host "Nintex Workflows Features activated " -ForegroundColor Green `r
		Write-Host "******************************************************************************" -ForegroundColor Magenta `r
		
		
		Write-Host "" `r
		Write-Host "******************************************************************************" -ForegroundColor Magenta `r
		Write-Host "About to activate Nintex Forms Features " -ForegroundColor Magenta `r
		Write-Host "******************************************************************************" -ForegroundColor Magenta `r
		# Nintex Forms Prerequisites Feature / SPSite / NintexFormsSitePrerequisites / 716f0ee9-e2b0-41f0-a73c-47ed73f135de
		Enable-SPFeature -Identity "716f0ee9-e2b0-41f0-a73c-47ed73f135de" -URL $parentURL -ErrorAction Stop -Force -Confirm:$False

		#Nintex Forms for SharePoint List Forms / SPSite / NintexFormsListSite / 202afc3c-7384-4700-978d-6da3d3cce192
		Enable-SPFeature -Identity "202afc3c-7384-4700-978d-6da3d3cce192" -URL $parentURL -ErrorAction Stop -Force -Confirm:$False

		#Nintex Forms for Nintex Workflow / SPSite / NintexFormsWorkflowSite / ac8addc7-7252-4136-8dcb-9887a277ae2c
		Enable-SPFeature -Identity "ac8addc7-7252-4136-8dcb-9887a277ae2c" -URL $parentURL -ErrorAction Stop -Force -Confirm:$False
		
		Write-Host "Nintex Forms Features activated " -ForegroundColor Green `r
		Write-Host "******************************************************************************" -ForegroundColor Magenta `r
		
		Write-Host "" `r
		Write-Host "******************************************************************************" -ForegroundColor Magenta `r
		Write-Host "About to deploy SharePoint Artefacts " -ForegroundColor Magenta `r
		Write-Host "******************************************************************************" -ForegroundColor Magenta `r
		

		############################################
		## Common Columns, CTs, Lists for Nintex Solution
		############################################
		$commonXmlFilePath = "$destFolderArtefacts\SPCommonStructure.xml"
		Write-Host "******************************************************************************" -ForegroundColor Cyan `r
		Write-Host "Deploying Common Columns, CTs, Lists for Nintex Solution..." -ForegroundColor Cyan  `r
		Write-Host "***************************************" -ForegroundColor Cyan `r
		Write-Host "SiteURL:"   -ForegroundColor Cyan  `r
		Write-Host "$siteURL"   -ForegroundColor Gray  `r
		Write-Host "File:"   -ForegroundColor Cyan  `r		
		Write-Host "$commonXmlFilePath"   -ForegroundColor Gray  `r		
		Write-Host "******************************************************************************" -ForegroundColor Cyan `r
		
		if(Test-Path $commonXmlFilePath)
		{
			$commonXmlFile = LoadXMLFile -xmlPath $commonXmlFilePath
			if($commonXmlFile -ne $null -and $commonXmlFile.HasChildNodes)
			{
				browseAndParseSPCommonStructureXML -siteURL $siteURL -SPCommonStructureDescriptionXML $commonXmlFile 
			}
			else
			{
				Write-Warning "XML File for <SPSiteColumns> is empty." 
			}
		}
		else
		{
			Write-Warning "XML File for <SPSiteColumns> does not exist." 
		}

		############################################
		# Upload files on site.
		############################################
		$fileToUploadFilePath = "$destFolderArtefacts\SPFileUploader.xml"
		$fileToUploadLocation = "$destFolderArtefacts\SPFileUploader\"
		Write-Host "******************************************************************************" -ForegroundColor Cyan `r
		Write-Host "Uploading file(s)..."  -ForegroundColor Cyan `r
		Write-Host "***************************************" -ForegroundColor Cyan `r
		Write-Host "SiteURL:"   -ForegroundColor Cyan  `r
		Write-Host "$siteURL"   -ForegroundColor Gray  `r
		Write-Host "File:"   -ForegroundColor Cyan  `r		
		Write-Host "$fileToUploadFilePath"   -ForegroundColor Gray  `r	
		Write-Host "******************************************************************************" -ForegroundColor Cyan `r
		
		if(Test-Path $fileToUploadFilePath)
		{
			$uploadFilesXML = LoadXMLFile -xmlPath  $fileToUploadFilePath
			if($uploadFilesXML -ne $null -and $uploadFilesXML.HasChildNodes)
			{
				browseFilesAndFoldersToUpload -siteURL $siteURL -sourceFolderPath $fileToUploadLocation -uploadFilesXML $uploadFilesXML 
			}
			else
			{
				Write-Warning "XML File for <SPFileUploader> is empty." 
			}
		}
		else
		{
			Write-Warning "XML File for <SPFileUploader> does not exist."
		}
			

		############################################
		## Add/Update Site Columns in site.
		############################################
		$SiteCollXmlFilePath = "$destFolderArtefacts\SPSiteColumns.xml"
		Write-Host "******************************************************************************" -ForegroundColor Cyan `r
		Write-Host "Deploying Site Columns..." -ForegroundColor Cyan  `r
		Write-Host "***************************************" -ForegroundColor Cyan `r
		Write-Host "SiteURL:"   -ForegroundColor Cyan  `r
		Write-Host "$siteURL"   -ForegroundColor Gray  `r
		Write-Host "File:"   -ForegroundColor Cyan  `r		
		Write-Host "$SiteCollXmlFilePath"   -ForegroundColor Gray  `r		
		Write-Host "******************************************************************************" -ForegroundColor Cyan `r
		
		if(Test-Path $SiteCollXmlFilePath)
		{
			$SiteCollXmlFile = LoadXMLFile -xmlPath $SiteCollXmlFilePath
			if($SiteCollXmlFile -ne $null -and $SiteCollXmlFile.HasChildNodes)
			{
				browseAndParseSiteColumnsXML -siteURL $siteURL -siteColumnsDescriptionXML $SiteCollXmlFile 
			}
			else
			{
				Write-Warning "XML File for <SPSiteColumns> is empty." 
			}
		}
		else
		{
			Write-Warning "XML File for <SPSiteColumns> does not exist." 
		}
		
		
		############################################
		## Add/Update Site Content Types in site.
		############################################
		$SiteContentTypeXmlFilePath = "$destFolderArtefacts\SPSiteContentTypes.xml"
		Write-Host "*******************************************************************************" -ForegroundColor Gray `r
		Write-Host "Deploying Site Content-Types..."  -ForegroundColor Cyan  `r
		Write-Host "***************************************" -ForegroundColor Cyan `r
		Write-Host "SiteURL:"   -ForegroundColor Cyan  `r
		Write-Host "$siteURL"   -ForegroundColor Gray  `r
		Write-Host "File:"   -ForegroundColor Cyan  `r		
		Write-Host "$SiteContentTypeXmlFilePath"   -ForegroundColor Gray  `r		
		Write-Host "******************************************************************************" -ForegroundColor Cyan `r
		
		if(Test-Path $SiteContentTypeXmlFilePath)
		{
			$SiteContentTypeXmlFile = LoadXMLFile -xmlPath $SiteContentTypeXmlFilePath	
			if($SiteContentTypeXmlFile -ne $null -and $SiteContentTypeXmlFile.HasChildNodes)
			{
				browseAndParseSiteContentTypesXML -siteURL $siteURL -SiteContentTypesDescriptionXML $SiteContentTypeXmlFile 
			}
			else
			{
				Write-Warning "XML File for <SPSiteContentTypes> is empty." 
			}
		}
		else
		{
			Write-Warning "XML File for <SPSiteContentTypes> does not exist." -ForegroundColor Cyan `r
		}	
		
		
		############################################
		## Add/Update Lists in site.
		############################################
		$SiteListsXmlFilePath = "$destFolderArtefacts\SPLists.xml" 
		Write-Host "*******************************************************************************" -ForegroundColor Gray `r
		Write-Host "Deploying Lists..."  -ForegroundColor Cyan  `r
		Write-Host "***************************************" -ForegroundColor Cyan `r
		Write-Host "SiteURL:"   -ForegroundColor Cyan  `r
		Write-Host "$siteURL"   -ForegroundColor Gray  `r
		Write-Host "File:"   -ForegroundColor Cyan  `r		
		Write-Host "$SiteListsXmlFilePath"   -ForegroundColor Gray  `r		
		Write-Host "******************************************************************************" -ForegroundColor Cyan `r
		
		if(Test-Path $SiteListsXmlFilePath)
		{
			$SiteListsXmlFile = LoadXMLFile -xmlPath $SiteListsXmlFilePath
			if($SiteListsXmlFile -ne $null -and $SiteListsXmlFile.HasChildNodes)
			{
				browseAndParseListsXML -siteURL $siteURL -listDescriptionXML $SiteListsXmlFile 
			}
			else
			{
				Write-Warning "XML File for <SPLists> is empty." 
			}
		}
		else
		{
			Write-Host "XML File for <SPLists> does not exist." -ForegroundColor Cyan `r
		}
		
		############################################
		## Add/Update Nintex Forms in site.
		############################################
		$formFilePath = "$destFolderArtefacts\NintexForms.xml" 
		Write-Host "*******************************************************************************" -ForegroundColor Gray `r
		Write-Host "Deploying Nintex Forms..."  -ForegroundColor Cyan  `r
		Write-Host "***************************************" -ForegroundColor Cyan `r
		Write-Host "SiteURL:"   -ForegroundColor Cyan  `r
		Write-Host "$siteURL"   -ForegroundColor Gray  `r
		Write-Host "File:"   -ForegroundColor Cyan  `r		
		Write-Host "$formFilePath"   -ForegroundColor Gray  `r		
		Write-Host "Folder containing all forms:"   -ForegroundColor Cyan  `r		
		Write-Host "$destFolderForms"   -ForegroundColor Gray  `r	
		Write-Host "******************************************************************************" -ForegroundColor Cyan `r
		
		if(Test-Path $formFilePath)
		{
			$formFileXML = LoadXMLFile -xmlPath  $formFilePath
			if($formFileXML -ne $null -and $formFileXML.HasChildNodes)
			{
				browseAndDeployFormsXML -siteURL $siteURL -formsDescriptionXML $formFileXML -srcFolder $destFolderForms
			}
			else
			{
				Write-Warning "XML File for <NintexForms> is empty." 
			}
		}
		else
		{
			Write-Host "XML File for <NintexForms> does not exist." -ForegroundColor Cyan `r
		}
		
		############################################
		## Add/Update Nintex Workflows in site.
		############################################
		$wfFilePath = "$destFolderArtefacts\NintexWorkflows.xml" 
		Write-Host "*******************************************************************************" -ForegroundColor Gray `r
		Write-Host "Deploying Nintex Workflows..."  -ForegroundColor Cyan  `r
		Write-Host "***************************************" -ForegroundColor Cyan `r
		Write-Host "SiteURL:"   -ForegroundColor Cyan  `r
		Write-Host "$siteURL"   -ForegroundColor Gray  `r
		Write-Host "File:"   -ForegroundColor Cyan  `r		
		Write-Host "$wfFilePath"   -ForegroundColor Gray  `r	
		Write-Host "Folder containing all workflows:"   -ForegroundColor Cyan  `r		
		Write-Host "$destFolderWF"   -ForegroundColor Gray  `r			
		Write-Host "******************************************************************************" -ForegroundColor Cyan `r
		
		if(Test-Path $wfFilePath)
		{
			$wfFileXML = LoadXMLFile -xmlPath  $wfFilePath
			if($wfFileXML -ne $null -and $wfFileXML.HasChildNodes)
			{
				browseAndDeployWorkflowsXML -siteURL $siteURL -workflowDescriptionXML $wfFileXML -srcFolder $destFolderWF
			}
			else
			{
				Write-Warning "XML File for <NintexWorkflows> is empty." 
			}
		}
		else
		{
			Write-Host "XML File for <NintexWorkflows> does not exist." -ForegroundColor Cyan `r
		}	
	
	}
	else
	{
		Write-Warning "Site '$siteURL' does not exist."
	}	
	
		############################################
		## Add/Update List Views in site.
		############################################
		$listViewXmlFilePath    = "$destFolderArtefacts\SPListViews.xml"
		Write-Host "*******************************************************************************" -ForegroundColor Cyan `r
		Write-Host "Deploying List Views..."  -ForegroundColor Cyan  `r
		Write-Host "***************************************" -ForegroundColor Cyan `r
		Write-Host "SiteURL:"   -ForegroundColor Cyan  `r
		Write-Host "$siteURL"   -ForegroundColor Cyan  `r
		Write-Host "File:"   -ForegroundColor Cyan  `r		
		Write-Host "$listViewXmlFilePath"   -ForegroundColor Cyan  `r		
		Write-Host "************************************************************************" -ForegroundColor Cyan `r
		$listViewsXML = LoadXMLFile -xmlPath $listViewXmlFilePath
		if($listViewsXML -ne $null -and $listViewsXML.HasChildNodes)
		{
			browseAndParseListViewsXML -siteURL $siteURL -listViewsXML $listViewsXML
		} 
		else
		{
			Write-Warning "XML File for <SPListViews> is empty." 
		} 
	
		############################################
		## Add/Update Groups in Site
		############################################
		$groupFilePath = "$destFolderArtefacts\SPSiteGroups.xml" 
		Write-Host "*******************************************************************************" -ForegroundColor Cyan `r
		Write-Host "Deploying SP Site Groups..."  -ForegroundColor Cyan  `r
		Write-Host "***************************************" -ForegroundColor Cyan `r
		Write-Host "SiteURL:"   -ForegroundColor Cyan  `r
		Write-Host "$siteURL"   -ForegroundColor Cyan  `r
		Write-Host "Folder containing all groups:"   -ForegroundColor Cyan  `r		
		Write-Host "$groupFilePath"   -ForegroundColor Cyan  `r			
		Write-Host "************************************************************************" -ForegroundColor Cyan `r
		
		if(Test-Path $groupFilePath)
		{
			$spGroupFileXML = LoadXMLFile -xmlPath $groupFilePath
			if($spGroupFileXML -ne $null -and $spGroupFileXML.HasChildNodes)
			{
				browseAndParseSiteGroupsXML -siteURL $siteURL -siteGroupsXML $spGroupFileXML
			}
			else
			{
				Write-Warning "XML File for <SPSiteGroups> is empty." 
			}
			  
		}
		else
		{
			Write-Host "XML File for <SPSiteGroups> does not exist." -ForegroundColor Cyan `r
		}	

		###########################################
		## Add/Update Permissions on Site
		###########################################
		$sitePermXMLFilePath    = "$destFolderArtefacts\SPSitePermissions.xml"
		if(Test-Path $sitePermXMLFilePath)
		{
			$sitePermissionXML = LoadXMLFile -xmlPath $sitePermXMLFilePath
			if($sitePermissionXML -ne $null -and $sitePermissionXML.HasChildNodes)
			{
				browseAndParseSPSitePermissionsXML -siteURL $siteURL -sitePermissionXML $sitePermissionXML
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


		###########################################
		## Add/Update Permissions on Lists
		###########################################
		$sitePermXMLFilePath    = "$destFolderArtefacts\SPListPermissions.xml"
		if(Test-Path $sitePermXMLFilePath)
		{
			$sitePermissionXML = LoadXMLFile -xmlPath $sitePermXMLFilePath
			if($sitePermissionXML -ne $null -and $sitePermissionXML.HasChildNodes)
			{
				browseAndParseSPListPermissionsXML -siteURL $siteURL -listPermissionXML $sitePermissionXML
			}  
			else
			{
				Write-Warning "XML File for <SPListPermissions> is empty." 
			}
			  
		}
		else
		{
			Write-Host "XML File for <SPListPermissions> does not exist." -ForegroundColor Cyan `r
		}	

		############################################
		## Add/Update Pages in Site
		############################################
		$sitePagesFilePath = "$destFolderArtefacts\SPSitePages.xml" 
		Write-Host "*******************************************************************************" -ForegroundColor Cyan `r
		Write-Host "Deploying SP Site Pages..."  -ForegroundColor Cyan  `r
		Write-Host "***************************************" -ForegroundColor Cyan `r
		Write-Host "SiteURL:"   -ForegroundColor Cyan  `r
		Write-Host "$siteURL"   -ForegroundColor Cyan  `r
		Write-Host "Folder containing all pages:"   -ForegroundColor Cyan  `r		
		Write-Host "$sitePagesFilePath"   -ForegroundColor Cyan  `r			
		Write-Host "************************************************************************" -ForegroundColor Cyan `r
		
		if(Test-Path $sitePagesFilePath)
		{
			$sitePagesFileXML = LoadXMLFile -xmlPath $sitePagesFilePath
			if($sitePagesFileXML -ne $null -and $sitePagesFileXML.HasChildNodes)
			{
				browseAndParseSPSitePagesXML -siteURL $siteURL -sitePagesDescriptionXML $sitePagesFileXML
			}
			else
			{
				Write-Warning "XML File for <SPSiteGroups> is empty." 
			}
			  
		}
		else
		{
			Write-Host "XML File for <SPSiteGroups> does not exist." -ForegroundColor Cyan `r
		}	

	############################################
	## Translate lists columns
	############################################
	$colTranslationFilePath = "$destFolderArtefacts\SPListColumnsTranslation.xml" 
	Write-Host "*******************************************************************************" -ForegroundColor Cyan `r
	Write-Host "Translating lists columns..."  -ForegroundColor Cyan  `r
	Write-Host "***************************************" -ForegroundColor Cyan `r
	Write-Host "SiteURL:"   -ForegroundColor Cyan  `r
	Write-Host "$siteURL"   -ForegroundColor Cyan  `r
	Write-Host "Folder containing all groups:"   -ForegroundColor Cyan  `r		
	Write-Host "$colTranslationFilePath"   -ForegroundColor Cyan  `r			
	Write-Host "************************************************************************" -ForegroundColor Cyan `r
		
	if(Test-Path $colTranslationFilePath)
	{
		$listColumnsTransXML = LoadXMLFile -xmlPath $colTranslationFilePath
		if($listColumnsTransXML -ne $null -and $listColumnsTransXML.HasChildNodes)
		{
			browseAndParseListColumnsTranslationXML -siteURL $siteURL -listColumnsTransXML $listColumnsTransXML
		}
		else
		{
			Write-Warning "XML File for <SPListColumnsTranslation> is empty." 
		}
	}
	else
	{
		Write-Host "XML File for <SPListColumnsTranslation> does not exist." -ForegroundColor Cyan `r
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