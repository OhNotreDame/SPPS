##############################################################################################
#              
# NAME: SPLists.psm1 
# PURPOSE: 
#	Manage Lists
#	Relies on an XML Configuration file to identify and handle the Lists.
#	See SPSiteLists.xml for Schema
#
# SOURCE : https://github.com/OhNotreDame/SPPS
#
##############################################################################################


<#
.SYNOPSIS
	Browse and parse the listDescriptionXML XML object
	
.DESCRIPTION
	Browse and parse the file listDescriptionXML
		For each node, Check if List exists in site siteURL
		If does not exist, create a new List	
		If exists, call update the List.		
	
.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER listDescriptionXML
	XML Object representing the lists to be created/updated
	
.EXAMPLE
	browseAndParseListsXML -siteURL <SiteURL> -listDescriptionXML <XMLObjectToParse>
	
.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 16.01.2017
	Last Updated by: JBO
	Last Updated: 16.01.2017
#>
function browseAndParseListsXML()
{
	[CmdletBinding()]
	param
    (
        [Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[XML]$listDescriptionXML
	)

    try
    {
		$functionName = $MyInvocation.MyCommand.Name
		Write-Debug "[$functionName] Entering function" 
		Write-Debug "[$functionName] Parameter / siteURL: $siteURL"
	
		if($listDescriptionXML.HasChildNodes)
		{
			$lists = $listDescriptionXML.SelectNodes("/Lists")
			foreach($list in $lists.List)
			{
				$listName = $list.Title.Trim()
				$isListExist = existList -siteURL $siteURL -listName  $listName

				if($isListExist -eq $false)
				{
					createList -siteURL $siteURL -listDefinitionXML $list
				}
				else
				{
					#Update List is not implemented yet.
					#updateList -siteURL $siteURL -listDefinitionXML $listDefinitionXML 

					#Refresh Site Columns used in list
					refreshSiteColumnsOnList -siteURL $siteURL -listName $listName 
				}
			}#end foreach
		}
		else
		{
			Write-Warning "[$functionName] List XML defintion file is empty."
		}

    }
    catch [Exception]
    {
		Write-Host "/!\ $functionName An exception has been caught /!\ "  -ForegroundColor Red `r
		Write-Host "Type: " $_.Exception.GetType().FullName   -ForegroundColor Red `r
		Write-Host "Message: " $_.Exception.Message  -ForegroundColor Red `r
		Write-Host "Stacktrace: " $_.Exception.Stacktrace  -ForegroundColor Red `r
    }
	finally
	{
		Write-Debug "[$functionName] Exiting function" 
	} 
}


<#
.SYNOPSIS
	Check if the list listName exists.
	
.DESCRIPTION
	Check if List $listName exists in site siteURL
		If exists, return true
		If not, return false
	
.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER listName
	Name of the list
	
.EXAMPLE
	existList -siteURL <SiteURL> -listName <listName>
	
.OUTPUTS
	True if List exists, false otherwise

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 16.01.2017
	Last Updated by: JBO
	Last Updated: 16.01.2017
#>
function existList()
{
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[String]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[String]$listName
	)
    
	$functionName = $MyInvocation.MyCommand.Name
	Write-Debug "[$functionName] Entering function" 
	Write-Debug "[$functionName] Parameter / siteURL: $siteURL"
	Write-Debug "[$functionName] Parameter / listName: $listName"
	$existList = $false;
    
	try
    {
		
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if($curWeb -ne $null)
		{
			$list=$curWeb.Lists.TryGetList($listName)
			if($list -ne $null)
			{
				$existList = $true
			}
			else
			{
				$existList = $false
			}
		}
		else
		{
			Write-Warning "[$functionName] Site '$siteURL' not found."
			$existList = $false
		}
	}
	catch [Exception]
	{
		Write-Host "/!\ $functionName An exception has been caught /!\ "  -ForegroundColor Red `r
		Write-Host "Type: " $_.Exception.GetType().FullName   -ForegroundColor Red `r
		Write-Host "Message: " $_.Exception.Message  -ForegroundColor Red `r
		Write-Host "Stacktrace: " $_.Exception.Stacktrace  -ForegroundColor Red `r
		$existList = $false
	}
	finally
	{
		if($curWeb -ne $null)
		{
			$curWeb.Dispose()
		}
		Write-Debug "[$functionName] Exiting function" 
	}
	return $existList;
 }


<#
.SYNOPSIS
	Check if the list listName exists and return it
	
.DESCRIPTION
	Check if List $listName exists in site siteURL
		If exists, return it (SPList Object)
		If not, return null
	
.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER listName
	Name of the list
	
.EXAMPLE
	getList -siteURL <SiteURL> -listName <listName>
	
.OUTPUTS
	SPList Object if List exists, null otherwise

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 16.01.2017
	Last Updated by: JBO
	Last Updated: 16.01.2017
#>
function getList()
{
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[String]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[String]$listName
	)
    
	$functionName = $MyInvocation.MyCommand.Name
	Write-Debug "[$functionName] Entering function" 
	Write-Debug "[$functionName] Parameter / siteURL: $siteURL"
	Write-Debug "[$functionName] Parameter / listName: $listName"

	try
	{
		
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if($curWeb -ne $null)
		{
			$list = $curWeb.Lists.TryGetList($listName)
			if($list -ne $null)
			{
				Write-Debug "[$functionName] List '$listName' found on site '$siteURL'."
				$listToReturn =  $list
			}
			else
			{
				Write-Warning "[$functionName] List '$listName' not found on site '$siteURL'."
			}
		}
		else
		{
			Write-Warning "[$functionName] Site '$siteURL' not found."
		}
	}
	catch [Exception]
	{
		Write-Host "/!\ $functionName An exception has been caught /!\ "  -ForegroundColor Red `r
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
		Write-Debug "[$functionName] Exiting function" 
    }
	return $listToReturn;
 }





 <#
.SYNOPSIS
	Rename the list listName with new name
	
.DESCRIPTION
	Check if List $listName exists in site siteURL
		If exists, return it (SPList Object)
		If not, return null
	
.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER listName
	Name of the list

.PARAMETER newName
	New Name of the list
		
.EXAMPLE
	renameList -siteURL <SiteURL> -listName <listName> -newListName <newListName>
	
.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 16.01.2017
	Last Updated by: JBO
	Last Updated: 16.01.2017
#>
function renameList()
{
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[String]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[String]$listName,
		[Parameter(Mandatory=$true, Position=3)]
		[String]$newListName
	)
    
	$functionName = $MyInvocation.MyCommand.Name
	Write-Debug "[$functionName] Entering function" 
	Write-Debug "[$functionName] Parameter / siteURL: $siteURL"
	Write-Debug "[$functionName] Parameter / listName: $listName"
	Write-Debug "[$functionName] Parameter / newListName: $newListName"

	try
	{
		
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if($curWeb -ne $null)
		{
			$list = $curWeb.Lists.TryGetList($listName)
			if($list -ne $null)
			{				
				$list.Title = $newListName
				$list.Update();
				$curWeb.Update();
			}
			else
			{
				Write-Warning "[$functionName] List '$listName' not found on site '$siteURL'."
			}
		}
		else
		{
			Write-Warning "[$functionName] Site '$siteURL' not found."
		}
	}
	catch [Exception]
	{
		Write-Host "/!\ $functionName An exception has been caught /!\ "  -ForegroundColor Red `r
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
		Write-Debug "[$functionName] Exiting function" 
    }
 }

<#
.SYNOPSIS
	Create the list based on its XML definition on site siteURL
	
.DESCRIPTION
	Create the list based on its XML definition on site siteURL
	
.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER listDefinitionXML
	List XML Definition
	
.EXAMPLE
	createList -siteURL <SiteURL> -listDefinitionXML <listXML>
	
.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 16.01.2017
	Last Updated by: JBO
	Last Updated: 16.01.2017
#>
function createList()
{
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[String]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[XML.XmlElement]$listDefinitionXML
	)
    
	try
	{
		$functionName = $MyInvocation.MyCommand.Name
		Write-Debug "[$functionName] Entering function" 
		Write-Debug "[$functionName] Parameter / siteURL: $siteURL"
		Write-Debug "[$functionName] Parameter / listTitle: $($listDefinitionXML.Title)"
		Write-Debug "[$functionName] Parameter / listName: $($listDefinitionXML.Name)"

		#Get the web object  
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if($curWeb -ne $null)
		{ 
			$curWeb.AllowUnsafeUpdates = $true
			# Primary List Settings	
			$listDescription = $listDefinitionXML.Description
			$listTitle = $listDefinitionXML.Title
			$listName = $listDefinitionXML.Name
			$listURL = $listDefinitionXML.Url
			$listTemplateType = $listDefinitionXML.Type
			
			# Versioning Settings
			$listVersioningEnabled = [System.Convert]::ToBoolean($listDefinitionXML.Attributes["VersioningEnabled"].value)
			$listMajorVersionLimit = $listDefinitionXML.MajorVersionLimit
			$listEnableMinorVersions = [System.Convert]::ToBoolean($listDefinitionXML.Attributes["EnableMinorVersions"].value)
			$listMajorWithMinorVersionsLimit = $listDefinitionXML.MajorWithMinorVersionsLimit
		 
			# Global List Settings
			$listQuickLaunch = [System.Convert]::ToBoolean($listDefinitionXML.Attributes["QuickLaunch"].value)
			$listDisableAttachments = [System.Convert]::ToBoolean($listDefinitionXML.Attributes["DisableAttachments"].value)
			$listFolderCreation = [System.Convert]::ToBoolean($listDefinitionXML.Attributes["FolderCreation"].value)
			$listEnableContentTypes = [System.Convert]::ToBoolean($listDefinitionXML.Attributes["EnableContentTypes"].value)
		
			# Other Settings
			$disableGridEditing  = [System.Convert]::ToBoolean($listDefinitionXML.Attributes["DisableGridEditing"].value)
			$enableDialogs  = [System.Convert]::ToBoolean($listDefinitionXML.Attributes["NavigateForFormsPages"].value)
			$disableCrawling  = [System.Convert]::ToBoolean($listDefinitionXML.Attributes["NoCrawl"].value)


			# Draft Version Visibility Settings
			####################################
			# Value = 0 (Reader) A reader can view the minor version of a document draft. 
			# Value = 1 (Author) An author can view the minor version of a document draft.
			# Value = 2 (Approver) An approver can view the minor version of a document draft.
			$listDraftVisibility = $listDefinitionXML.DraftVersionVisibility
			switch ($listDraftVisibility)
			{
			"0"	 {
					 $listDraftVersionVisibility = [Microsoft.SharePoint.DraftVisibilityType]::Reader;
					 break;
				 }
			"1"	 {
					 $listDraftVersionVisibility = [Microsoft.SharePoint.DraftVisibilityType]::Author;
					 break;
				 }
			"2"	 {
					 $listDraftVersionVisibility = [Microsoft.SharePoint.DraftVisibilityType]::Approver;
					 break;
				 }
			default
				 {
					 $listDraftVersionVisibility = [Microsoft.SharePoint.DraftVisibilityType]::Reader;
					 break;
				 }
			}

			# Moderation Settings
			$listModeratedList  = [System.Convert]::ToBoolean($listDefinitionXML.Attributes["ModeratedList"].value)
			$listModerationType  = [System.Convert]::ToBoolean($listDefinitionXML.Attributes["ModerationType"].value)

			# isPrivate & Ordered.... IGNORED .....
			#$listOrderedList  = [System.Convert]::ToBoolean($listDefinitionXML.Attributes["OrderedList"].value)
			#$listPrivateList  = [System.Convert]::ToBoolean($listDefinitionXML.Attributes["PrivateList"].value)
			

			####################################
			# DISPLAY FOR DEBUG
			####################################
			
			# Primary List Settings
			#----------------------------------#
			#Write-Host "[$functionName] XML- listTitle: $listTitle" -ForegroundColor Magenta `r	
			#Write-Host "[$functionName] XML- listDescription: $listDescription" -ForegroundColor Magenta `r	
			#Write-Host "[$functionName] XML- listURL: $listURL" -ForegroundColor Magenta `r	
			#Write-Host "[$functionName] XML- listTemplateType: $listTemplateType" -ForegroundColor Magenta `r	
			# Versioning Settings
			#----------------------------------#
			#Write-Host "[$functionName] XML- listVersioningEnabled: $listVersioningEnabled" -ForegroundColor Magenta `r	
			#Write-Host "[$functionName] XML- listEnableMinorVersions: $listEnableMinorVersions" -ForegroundColor Magenta `r	
			# Global List Settings
			#----------------------------------#
			#Write-Host "[$functionName] XML- listQuickLaunch: $listQuickLaunch" -ForegroundColor Magenta `r	
			#Write-Host "[$functionName] XML- listDisableAttachments: $listDisableAttachments" -ForegroundColor Magenta `r	
			#Write-Host "[$functionName] XML- listFolderCreation: $listFolderCreation" -ForegroundColor Magenta `r	
			#Write-Host "[$functionName] XML- listEnableContentTypes: $listEnableContentTypes" -ForegroundColor Magenta `r	
			# Draft Version Visibility Settings
			#----------------------------------#
			#Write-Host "[$functionName] XML- listDraftVersionVisibility: $listDraftVersionVisibility" -ForegroundColor Magenta `r	
			# Moderation Settings
			#----------------------------------#
			#Write-Host "[$functionName] XML- listModeratedList: $listModeratedList" -ForegroundColor Magenta `r	
			#Write-Host "[$functionName] XML- listModerationType: $listModerationType" -ForegroundColor Magenta `r	
			# isPrivate & Ordered.... IGNORED .....
			#----------------------------------#
			#Write-Host "[$functionName] XML- listPrivateList: $listPrivateList" -ForegroundColor Magenta `r	
			#Write-Host "[$functionName] XML- listOrderedList: $listOrderedList" -ForegroundColor Magenta `r	

			if(![string]::IsNullOrEmpty($listName) -and  ![string]::IsNullOrEmpty($listTitle))
			{

				$spList = $curWeb.Lists.TryGetList($listTitle)
				if(($spList -eq $null))
				{
					Write-Host "[$functionName] About to add List '$listName'." -ForegroundColor Magenta `r	
	
					#Get List Template
					$listTemplateFullText = $curWeb.ListTemplates | Where-Object {$_.Type -eq $listTemplateType}
					
					#Create List
					$curWeb.Lists.Add($listName, $listDescription, $listTemplateFullText)
					$curWeb.Update();
			         
					#Get newly created List and apply settings
					$spList = $curWeb.Lists.TryGetList($listName)
					if(($spList -ne $null))
					{
						Write-Host "[$functionName] List '$listName' successfully created." -ForegroundColor Green `r
						Write-Host "[$functionName] About to setup additional settings on List '$listTitle'." -ForegroundColor Magenta `r	
					
						#----------------------------------------
						# Set final title
						#----------------------------------------
						Write-Host "[$functionName] List '$listName' - Set final title : $listTitle"
						ForEach($culture in $curWeb.SupportedUICultures)
						{
							[System.Threading.Thread]::CurrentThread.CurrentUICulture=$culture;
							$spList.Title = $listTitle;
							$spList.Update();
						}

						#----------------------------------------
						# Enable/Disable on QuickLaunch
						#----------------------------------------
						$spList.onQuickLaunch = $listQuickLaunch;
					
						#----------------------------------------
						# Global List Settings
						#----------------------------------------

						if ($listEnableContentTypes -ne $null){
							$spList.ContentTypesEnabled = $listEnableContentTypes
						}

						if ($listDisableAttachments -eq $true){
							$listEnableAttachments = !$listDisableAttachments;
							$spList.EnableAttachments = $listEnableAttachments
						}

						if ($listFolderCreation -ne $null){
							$spList.EnableFolderCreation = $listFolderCreation
						}

						#----------------------------------------
						# Major Versioning Settings (activation + limits)
						#----------------------------------------
						if ($listVersioningEnabled -ne $null)
						{
							$spList.EnableVersioning = $listVersioningEnabled;
							if ($listMajorVersionLimit -ne $null)
							{
								$spList.MajorVersionLimit = $listMajorVersionLimit;
							}
						}
				
						#----------------------------------------
						# Minor Versioning Settings  (activation + limits)
						#----------------------------------------
						if ( ($listTemplateType -eq "101") -and ($listEnableMinorVersions -ne $null) )
						{
							$spList.EnableMinorVersions = $listEnableMinorVersions;
							if ($listMajorWithMinorVersionsLimit -ne $null)
							{
								$spList.EnableMinorVersions = $listEnableMinorVersions;
							}		
						}

						if ($listDraftVersionVisibility -ne $null){
							$spList.DraftVersionVisibility = $listDraftVersionVisibility;
						}
										
						#----------------------------------------
						# Moderation Settings
						#----------------------------------------
						if ($listModeratedList -ne $null){
							$spList.EnableModeration = $listModeratedList;
						}		

						#----------------------------------------
						# Forms/Dialogs Settings
						#----------------------------------------
						if ($disableGridEditing -ne $null){
							Write-Host "disableGridEditing: $disableGridEditing"
							$spList.DisableGridEditing = $disableGridEditing
						}
						
						#----------------------------------------
						# Crawl Settings
						#----------------------------------------
						if ($enableDialogs -ne $null){
							Write-Host "enableDialogs: $enableDialogs"
							$spList.NavigateForFormsPages = $enableDialogs
						}

						#----------------------------------------
						# Crawl Settings
						#----------------------------------------
						if ($disableCrawling -ne $null){
							Write-Host "disableCrawling: $disableCrawling"
							$spList.NoCrawl  = $disableCrawling
						}


						$spList.Update();
						Write-Host "[$functionName] List '$listTitle' successfully configured." -ForegroundColor Green `r

						Write-Host "[$functionName] About to manage Content-Types on List '$listTitle'" -ForegroundColor Magenta `r
						#----------------------------------------
						#Add Content-Types to List
						#----------------------------------------
						foreach($contentType in $listDefinitionXML.ContentTypes.ContentType)
						{
							$ctName = $contentType.Name
							if(![string]::IsNullOrEmpty($ctName))
							{
								# Add Content-Type to list
								addContentTypeToList -siteURL $siteURL -listName $listTitle -ctName $ctName
						
								# Set Default Content-Type
								$setDefaultCtType = [System.Convert]::ToBoolean($contentType.Attributes["SetAsDefault"].value)
								if($setDefaultCtType -eq $true)
								{
									setDefaultContentTypeToList -siteURL $siteURL -listName $listTitle -ctName $ctName
								}  
                  
							}
						}

						#----------------------------------------
						#Remove Content-Types to List
						#----------------------------------------
						Write-Host "[$functionName] About to remove Content-Types from List '$listTitle'" -ForegroundColor Magenta `r
						foreach($contentType in $listDefinitionXML.ContentTypesToRemove.ContentType)
						{
						
							$ctName = $contentType.Name
							if(![string]::IsNullOrEmpty($ctName))
							{
								# Add Content-Type to list
								removeContentTypeFromList -siteURL $siteURL -listName $listTitle -ctName $ctName
							}
						}
						Write-Host "[$functionName] Content-Types managed on List '$listTitle'" -ForegroundColor Green `r
					
					}
					else
					{
						Write-Warning "[$functionName] List '$listName' does not exist on site '$siteURL'."						
					}

				}
				else
				{
					Write-Warning "[$functionName] List '$listName' already exists on site '$siteURL'."
				}

			 }
			else
			{
				Write-Warning "[$functionName] List XML parameters are empty."
			}
		}
		else
		{
			Write-Warning "[$functionName] Site '$siteURL' not found."
		}

    }
    catch [Exception]
    {
		Write-Host "/!\ $functionName An exception has been caught /!\ "  -ForegroundColor Red `r
		Write-Host "Type: " $_.Exception.GetType().FullName   -ForegroundColor Red `r
		Write-Host "Message: " $_.Exception.Message  -ForegroundColor Red `r
		Write-Host "Stacktrace: " $_.Exception.Stacktrace  -ForegroundColor Red `r
    }
    finally
    {
		$curWeb.AllowUnsafeUpdates = $false
		if($curWeb -ne $null)
		{
			$curWeb.Dispose()
		}
		Write-Debug "[$functionName] Exiting function" 
    }

 }


<#
.SYNOPSIS
	Will set AS DEFAULT the Content-Type $ctName on the list $listName
	
.DESCRIPTION
	Will set AS DEFAULT the Content-Type $ctName on the list $listName
	
.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER listName
	Name of the list

.PARAMETER ctName
	Name of the content-type to set as default
	
.EXAMPLE
	setDefaultContentType -siteURL <SiteURL> -listName <listName> -ctName <ctName>
	
.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 16.01.2017
	Last Updated by: JBO
	Last Updated: 16.01.2017
#>
function setDefaultContentTypeToList()
{
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[String]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[String]$listName,
		[Parameter(Mandatory=$true, Position=3)]
		[String]$ctName
	)
	$functionName = $MyInvocation.MyCommand.Name
	Write-Debug "[$functionName] Entering function" 
	Write-Debug "[$functionName] Parameter / siteURL: $siteURL"
	Write-Debug "[$functionName] Parameter / listName: $listName"
	Write-Debug "[$functionName] ctName / listName: $ctName"
 
	try
	{
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if($curWeb -ne $null)
		{
			$spList=$curWeb.Lists.TryGetList($listName)
			if($spList -ne $null)
			{
				#Add site Content-Types to the list
				$ctOnList = $spList.ContentTypes[$ctName]
				if($ctOnList -ne $null)
				{
					
					Write-Host "[$functionName] About to set Content-Type '$ctName' as Default Content-Type of List '$listName'." -ForegroundColor Magenta `r
					$currentCTOrder = $spList.ContentTypes
					$result= New-Object System.Collections.Generic.List[Microsoft.SharePoint.SPContentType]
					foreach ($ct in $currentCTOrder)
					{
							if ($ct.Name.Contains($ctName))
							{
								$result.Add($ct)
							}
					}
					$spList.RootFolder.UniqueContentTypeOrder = $result
					$spList.RootFolder.Update() 
					
					$spList.Update() 
					$curWeb.Update() 
					Write-Host "[$functionName] Content-Type '$ctName' successfully set As Default Content-Type for List '$listName'." -ForegroundColor Green `r
				}
				else
				{
					Write-Warning "[$functionName] Content-Type '$ctName' is not associated with List '$listName'."
				}

			}
			else
			{
				Write-Warning "[$functionName] List '$listName' not found on site '$siteURL'."
			}
		} 
		else
		{
			Write-Warning "[$functionName] Site '$siteURL' not found."
		}
	}
	catch [Exception]
	{
		Write-Host "/!\ $functionName An exception has been caught /!\ "  -ForegroundColor Red `r
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
	}
}



<#
.SYNOPSIS
	Refresh all the fields of a list based on their Site Column definition
	
.DESCRIPTION
	Refresh all the fields of a list based on their Site Column definition
	
.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER listName
	Name of the list
	
.EXAMPLE
	refreshSiteColumnsOnList -siteURL <SiteURL> -listName <ListName> 
	
.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 28.04.2017
	Last Updated by: JBO
	Last Updated: 28.04.2017
#>
function refreshSiteColumnsOnList()
{
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[String]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[String]$listName
	)
    
	$functionName = $MyInvocation.MyCommand.Name
	Write-Debug "[$functionName] Entering function" 
	Write-Debug "[$functionName] Parameter / siteURL: $siteURL"
	Write-Debug "[$functionName] Parameter / listName: $listName"
		
	try
	{
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if($curWeb -ne $null)
		{	
			$curList = $curWeb.Lists.TryGetList($listName)
			if($curList -ne $null)
			{

				Write-Host "[$functionName] About to refresh List '$listName' from site '$siteURL'." -ForegroundColor Magenta  `r
				$myFieldsArray = New-Object System.Collections.ArrayList

				$curList.Fields  | ForEach-Object { 
						$internalName = $_.InternalName
						$staticName = $_.StaticName
						Write-Host "[$functionName] Adding '$internalName' ('$staticName') to myFieldsArray"
						$myFieldsArray.Add($_.InternalName) > $null
				}

				[System.Threading.Thread]::CurrentThread.CurrentUICulture=$curWeb.UICulture;
				foreach ($fieldName in $myFieldsArray) {
					Write-Host "[$functionName] Refreshing column '$fieldName' on List" -ForegroundColor Cyan  `r
					$fieldToUpdate = $curList.Fields.GetFieldByInternalName($fieldName)
					Write-Host "[$functionName] Sealed: " $fieldToUpdate.Sealed
					Write-Host "[$functionName] ReadOnlyField: " $fieldToUpdate.ReadOnlyField
				
					if (($fieldToUpdate -ne $null) -and (!$fieldToUpdate.ReadOnlyField))
					#if (($fieldToUpdate -ne $null) -and ((!$fieldToUpdate.Sealed) -or (!$fieldToUpdate.ReadOnlyField)))
					{
						$parentWebField = $curWeb.Fields.TryGetFieldByStaticName($fieldName)
						if($parentWebField -ne $null)
						{
							#$parentWebField.SchemaXml
							$fieldToUpdate.SchemaXml = $parentWebField.SchemaXml;
							$fieldToUpdate.Update();
						}
						else
						{
							Write-Warning "[$functionName] Field '$fieldName' not found in SPWeb."
						}
					}
					else
					{
						Write-Warning "[$functionName] SPField object '$fieldName' on list '$listName' is null or read-only."
					}
				}		
				$curList.Update();

				Write-Host "[$functionName] List '$listName' successfully refreshed from site '$siteURL'."  -ForegroundColor Green  `r
			}
			else
			{
				Write-Warning "[$functionName] List '$listName' does not exist on site '$siteURL'."
			}
		}
		else
		{
			Write-Warning "[$functionName] Site '$siteURL' not found."
		}

	}
	catch [Exception]
	{
		Write-Host "/!\ $functionName An exception has been caught /!\ "  -ForegroundColor Red `r
		Write-Host "Type: " $_.Exception.GetType().FullName   -ForegroundColor Red `r
		Write-Host "Message: " $_.Exception.Message  -ForegroundColor Red `r
		Write-Host "Stacktrace: " $_.Exception.Stacktrace  -ForegroundColor Red `r
	}
	finally
	{
		Write-Debug "[$functionName] Exiting function" 
	}
}

<#
.SYNOPSIS
	Update the list $listName based on its XML definition on site siteURL (NOT IMPLEMENTED)
	
.DESCRIPTION
	Update the list $listName based on its XML definition on site siteURL (NOT IMPLEMENTED)
	
.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER listName
	Name of the list
	
.PARAMETER listDefinitionXML
	List XML Definition
	
.EXAMPLE
	updateList -siteURL <SiteURL> -listName <ListName> -listDefinitionXML <listXML>
	
.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 16.01.2017
	Last Updated by: JBO
	Last Updated: 16.01.2017
#>
function updateList()
{
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[String]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[String]$listName,
		[Parameter(Mandatory=$false, Position=3)]
		[XML.XmlElement]$listDefinitionXML
	)
    
	$functionName = $MyInvocation.MyCommand.Name
	Write-Debug "[$functionName] Entering function" 
	Write-Debug "[$functionName] Parameter / siteURL: $siteURL"
	Write-Debug "[$functionName] Parameter / listName: $listName"
		
	try
	{
		Write-Host "#------------------------------------------------#" -ForegroundColor Red `r
		Write-Host "#- Function '$functionName' is not implemented. -#" -ForegroundColor Red `r
		Write-Host "#------------------------------------------------#" -ForegroundColor Red `r
	}
	catch [Exception]
	{
		Write-Host "/!\ $functionName An exception has been caught /!\ "  -ForegroundColor Red `r
		Write-Host "Type: " $_.Exception.GetType().FullName   -ForegroundColor Red `r
		Write-Host "Message: " $_.Exception.Message  -ForegroundColor Red `r
		Write-Host "Stacktrace: " $_.Exception.Stacktrace  -ForegroundColor Red `r
	}
	finally
	{
		Write-Debug "[$functionName] Exiting function" 
	}
}


<#
.SYNOPSIS
	Add the content-type $ctName to the list $listName
	
.DESCRIPTION
	Add the content-type $ctName to the list $listName
	If not already associated to the list, add the content-type $ctName
	
.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER listName
	Name of the list
	
.PARAMETER ctName
	Name of the content-type to add to the list
	
.EXAMPLE
	addContentTypeToList -siteURL <SiteURL> -listName <ListName> -ctName <ctName>
	
.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 16.01.2017
	Last Updated by: JBO
	Last Updated: 16.01.2017
#>
function addContentTypeToList()
{
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[String]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[String]$listName,
		[Parameter(Mandatory=$true, Position=3)]
		[String]$ctName
	)
    
	$functionName = $MyInvocation.MyCommand.Name
	Write-Debug "[$functionName] Entering function" 
	Write-Debug "[$functionName] Parameter / siteURL: $siteURL"
	Write-Debug "[$functionName] Parameter / listName: $listName"
	Write-Debug "[$functionName] Parameter / ctName: $ctName"
	
    try
    {
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if($curWeb -ne $null)
		{

			# Search Site Content-Type '$ctName' at Parent Site Collection Level
			Write-Host "[$functionName] Looking for '$ctName' at Parent Site Collection Level."
			$ctToAdd = $curWeb.Site.RootWeb.ContentTypes[$ctName]
				
			# If Site Content-Type is not available at SiteCollection Level, Search Site Column '$ctName' at Current Subsite Level
            if($ctToAdd -eq $null)
            {
                Write-Host "[$functionName] Looking for '$ctName' at Current Subsite Level."
				$ctToAdd = $curWeb.ContentTypes[$ctName]
	        }

			if ($ctToAdd -ne $null)
			{
				#Check the list existance in site and Content-Type not associated with the list.
				$curList = $curWeb.Lists.TryGetList($listName)
				if($curList -ne $null)
				{
					if (!$curList.ContentTypesEnabled)
					{
						$curList.ContentTypesEnabled = $true
						$curList.Update()
					}

					$ctOnList = $curList.ContentTypes[$ctName]
					#Add site Content-Types to the list
					if($ctOnList -eq $null)
					{
						Write-Host "[$functionName] About to add Content-Type '$ctName' to List '$listName'." -ForegroundColor Magenta `r
						$ct = $curList.ContentTypes.Add($ctToAdd)
						$curList.Update()
						$curWeb.Update()
						Write-Host "[$functionName] Content-Type '$ctName' added to List '$listName'." -ForegroundColor Green `r
					}
					else
					{
						Write-Warning "[$functionName] Content-Type '$ctName' already associated with List '$listName'."
					}
				}
				else
				{
					Write-Warning "[$functionName] List '$listName' does not exist on site '$siteURL'."		  
				}
			}
			else
			{
				Write-Warning "[$functionName] Content-Type '$ctName' does not exist on site '$siteURL' or on its Parent Site Collection."
			}
		} 
		else
        {
           Write-Warning "[$functionName] Site '$siteURL' not found."
        }
    }
    catch [Exception]
    {
		Write-Host "/!\ $functionName An exception has been caught /!\ "  -ForegroundColor Red `r
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
		Write-Debug "[$functionName] Exiting function" 
    }
 }


<#
.SYNOPSIS
	Remove the content-type $ctName from the list $listName
	
.DESCRIPTION
	Remove the content-type $ctName from the list $listName
	If already associated to the list, remove the content-type $ctName
	
.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER listName
	Name of the list
	
.PARAMETER ctName
	Name of the content-type to add to the list
	
.EXAMPLE
	removeContentTypeFromList -siteURL <SiteURL> -listName <ListName> -ctName <ctName>
	
.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 16.01.2017
	Last Updated by: JBO
	Last Updated: 16.01.2017
#>
function removeContentTypeFromList()
{
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[String]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[String]$listName,
		[Parameter(Mandatory=$true, Position=3)]
		[String]$ctName 
	)
    
	$functionName = $MyInvocation.MyCommand.Name
	Write-Debug "[$functionName] Entering function" 
	Write-Debug "[$functionName] Parameter / siteURL: $siteURL"
	Write-Debug "[$functionName] Parameter / listName: $listName"
	Write-Debug "[$functionName] Parameter / ctName: $ctName"
   
   try
    {
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if($curWeb -ne $null)
		{
			$curList=$curWeb.Lists.TryGetList($listName)
			if(($curList -ne $null))
			{
                $ctToRemove = $curList.ContentTypes[$ctName]
                if($ctToRemove -ne $null)
                {
					Write-Host "[$functionName] About to remove Content-Type '$ctName' from List '$listName'." -ForegroundColor Magenta `r
					$curList.ContentTypes.Delete($ctToRemove.Id)
					$curList.Update()
					Write-Host "[$functionName] Content-Type '$ctName' has been removed from List '$listName'." -ForegroundColor Green `r
                }
                else
                {
					Write-Warning "[$functionName] Content-Type '$ctName' does not exist in List '$listName'."
                }
            }
			else
			{
				Write-Warning "[$functionName] List '$listName' does not exist on Site '$siteURL'." 
			}
		}
		else
		{
			Write-Warning "[$functionName] Site '$siteURL' not found."
		}
    }
    catch [Exception]
    {
		Write-Host "/!\ $functionName An exception has been caught /!\ "  -ForegroundColor Red `r
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
		Write-Debug "[$functionName] Exiting function" 
    }
 }
 

<#
.SYNOPSIS
	Add the field $fieldStaticName to list $listName
	
.DESCRIPTION
	Add the field $fieldStaticName to list $listName
	
.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER listName
	Name of the list
	
.PARAMETER fieldStaticName
	Static Name of the SiteColumn to add to the list
	
.EXAMPLE
	addSiteColumnToList -siteURL <SiteURL> -listName <ListName> -fieldStaticName <fieldStaticName>
	
.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 16.01.2017
	Last Updated by: JBO
	Last Updated: 16.01.2017
#>
function addSiteColumnToList()  
{
    [CmdletBinding()]
	Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[string]$listName,
		[Parameter(Mandatory=$true, Position=3)]
		[string]$fieldStaticName
	)
    
	$functionName = $MyInvocation.MyCommand.Name
	Write-Debug "[$functionName] Entering function" 
	Write-Debug "[$functionName] Parameter / siteURL: $siteURL"
	Write-Debug "[$functionName] Parameter / listName: $listName"
	Write-Debug "[$functionName] Parameter / fieldStaticName: $fieldStaticName"
	
	try
    {
		
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
        if($curWeb -ne $null)
        {
            $list = $curWeb.Lists.TryGetList($listName)
            if($list -ne $null)
            {
			    # Search Site Column '$fieldStaticName' at Parent Site Collection Level
				Write-Debug "[$functionName] Looking for '$fieldStaticName' at Parent Site Collection Level."
				$listFieldToAdd = $curWeb.Site.RootWeb.Fields.TryGetFieldByStaticName($fieldStaticName)
				
				# If Column is not available at SiteCollection Level, Search Site Column '$fieldStaticName' at Current Subsite Level
                if($listFieldToAdd -eq $null)
                {
                    Write-Host "[$functionName] Looking for '$fieldStaticName' at Current Subsite Level."
					$listFieldToAdd = $curWeb.Fields.TryGetFieldByStaticName($fieldStaticName)
	            }
				
                if($listFieldToAdd -ne $null)
                {
					$fieldName = $listFieldToAdd.Title;

					if($list.Fields.TryGetFieldByStaticName($fieldStaticName) -eq $null)
                    {
                        Write-Host "[$functionName] About to add SiteColumn '$fieldName' to list '$listName'."  -ForegroundColor Magenta `r
 					    $list.Fields.Add($listFieldToAdd)
                        $list.Update()
                        Write-Host "[$functionName] SiteColumn '$fieldName' is added to list '$listName'."  -ForegroundColor Green  `r
                    }
                    else
                    {
                        Write-Warning "[$functionName] SiteColumn '$fieldName' already exists on list '$listName'."
                    }
                }
				else
				{
					Write-Warning "[$functionName] SiteColumn '$fieldName' does not exist on site '$siteURL' or on its Parent Site Collection."
				}
            }
			else
			{
				Write-Warning "[$functionName] List '$listName' does not exist on site '$siteURL'"  -ForegroundColor Cyan  `r
			}
        }
		else
		{
			Write-Warning "[$functionName] Site '$siteURL' not found."  -ForegroundColor Cyan  `r
		}
    }
    catch [Exception]
    {
		Write-Host "/!\ $functionName An exception has been caught /!\ "  -ForegroundColor Red `r
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
		Write-Debug "[$functionName] Exiting function" 
    }
}


<#
.SYNOPSIS
	Remove the field $fieldStaticName from the list $listName
	
.DESCRIPTION
	Remove the field $fieldStaticName from the list $listName
	
.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER listName
	Name of the list
	
.PARAMETER fieldStaticName
	Static Name of the SiteColumn to be removed from the list
	
.EXAMPLE
	removeSiteColumnFromList -siteURL <SiteURL> -listName <ListName> -fieldStaticName <fieldStaticName>
	
.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 16.01.2017
	Last Updated by: JBO
	Last Updated: 16.01.2017
#>
function removeSiteColumnFromList()  
{
    [CmdletBinding()]
	Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[string]$listName,
		[Parameter(Mandatory=$true, Position=3)]
		[string]$fieldStaticName
	)	
	    
	$functionName = $MyInvocation.MyCommand.Name
	Write-Debug "[$functionName] Entering function" 
	Write-Debug "[$functionName] Parameter / siteURL: $siteURL"
	Write-Debug "[$functionName] Parameter / listName: $listName"
	Write-Debug "[$functionName] Parameter / fieldStaticName: $fieldStaticName"
	
    try
    {
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if($curWeb -ne $null)
        {
            $list = $curWeb.Lists.TryGetList($listName)
            if($list -ne $null)
            {                
				$listFieldToRemove=$list.Fields.TryGetFieldByStaticName($fieldStaticName)
                if($listFieldToRemove -ne $null)
                {
					$fieldName = $listFieldToRemove.Title;
					
					Write-Host "[$functionName] About to remove SiteColumn '$fieldName' from List '$listName'." -ForegroundColor Magenta  `r
					$listFieldToRemove.Sealed = $false;
					$listFieldToRemove.AllowDeletion = $true;
					$listFieldToRemove.ReadOnlyField = $false
					$listFieldToRemove.Update();
					$list.Fields.Delete($fieldStaticName)
					$list.Update();
					Write-Host "[$functionName] SiteColumn '$fieldName' successfully removed from List '$listName'."-ForegroundColor Green  `r
				}
				else
				{
					Write-Warning "[$functionName] SiteColumn '$fieldName' already removed from List '$listName'."
				}
			}
			else
			{
				Write-Warning "[$functionName] List '$listName' does not exist on site '$siteURL'."
			}
        }
		else
		{
			Write-Warning "[$functionName] Site '$siteURL' not found."
		}

    }
    catch [Exception]
    {
		Write-Host "/!\ $functionName An exception has been caught /!\ "  -ForegroundColor Red `r
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
		Write-Debug "[$functionName] Exiting function" 
    }
}


<#
.SYNOPSIS
	Check if the list listName exists and then delete it.
	
.DESCRIPTION
	Check if the list listName exists and then delete it.
	
.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER listName
	Name of the list
	
.EXAMPLE
	deleteList -siteURL <SiteURL> -listName <listName>
	
.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 16.01.2017
	Last Updated by: JBO
	Last Updated: 16.01.2017
#>
function deleteList()
{
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[String]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[String]$listName
	)
    
	$functionName = $MyInvocation.MyCommand.Name
	Write-Debug "[$functionName] Entering function" 
	Write-Debug "[$functionName] Parameter / siteURL: $siteURL"
	Write-Debug "[$functionName] Parameter / listName: $listName"
    
	try
	{
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if($curWeb -ne $null)
		{
			$list = getList -siteURL $siteURL -listName $listName
			if($list -ne $null)
			{
				Write-Host "[$functionName] About to delete List '$listName' from site '$siteURL'." -ForegroundColor Magenta  `r
				$list.AllowDeletion = $true
				$list.Update()
				$list.Delete()
				Write-Host "[$functionName] List '$listName' successfully deleted from site '$siteURL'."  -ForegroundColor Green  `r
			}
			else
			{
				Write-Warning "[$functionName] List '$listName' does not exist on site '$siteURL'."
			}
		}
		else
		{
			Write-Warning "[$functionName] Site '$siteURL' not found."
		}

	}
	catch [Exception]
	{
		Write-Host "/!\ $functionName An exception has been caught /!\ "  -ForegroundColor Red `r
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
		Write-Debug "[$functionName] Exiting function" 
	}
}


<#
.SYNOPSIS
	Set Item Level Security $securityLevel on List $listName
	
.DESCRIPTION
	Set Item Level Security $securityLevel on List $listName
	
.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER listName
	Name of the list
	
.PARAMETER readSecurity
	Security Level to Apply for Read actions
		1 <> All users have Read access to all items.
		2 <> Users have Read access only to items that they create.
		
.PARAMETER readSecurity
	Security Level to Apply for Write actions
		1 <> All users can modify all items.
		2 <> Users can modify only items that they create.
		4 <> Users cannot modify any list item.
	
.EXAMPLE
	setItemLevelSecurity -siteURL <SiteURL> -listName <listName> -readSecurity <1|2> -writeSecurity <1|2|4>
	
.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 16.01.2017
	Last Updated by: JBO
	Last Updated: 16.01.2017
#>
 function setItemLevelSecurity()
 {
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[String]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[String]$listName,
		[Parameter(Mandatory=$true, Position=3)]
		[int]$readSecurity,
		[Parameter(Mandatory=$true, Position=4)]
		[int]$writeSecurity
	)
	    
	$functionName = $MyInvocation.MyCommand.Name
	Write-Debug "[$functionName] Entering function" 
	Write-Debug "[$functionName] Parameter / siteURL: $siteURL"
	Write-Debug "[$functionName] Parameter / listName: $listName"
	Write-Debug "[$functionName] Parameter / readSecurity: $readSecurity"
	Write-Debug "[$functionName] Parameter / writeSecurity: $writeSecurity"
	
	try
	{
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if($curWeb -ne $null)
		{
			$list = getList -siteURL $siteURL -listName $listName
			if($list -ne $null)
			{
				Write-Host "[$functionName] About to set item level security on List '$listName' from site '$siteURL'." -ForegroundColor Magenta  `r
				$list.ReadSecurity = $readSecurity;
				$list.WriteSecurity = $writeSecurity;
				$list.Update();
				Write-Host "[$functionName] Item level security successfully applied on List '$listName' from site '$siteURL'."  -ForegroundColor Green  `r
			}
			else
			{
				Write-Warning "[$functionName] List '$listName' does not exist on site '$siteURL'."
			}
		}
		else
		{
			Write-Warning "[$functionName] Site '$siteURL' not found."
		}

	}
	catch [Exception]
	{
		Write-Host "/!\ $functionName An exception has been caught /!\ "  -ForegroundColor Red `r
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
	}

}