##############################################################################################
#              
# NAME: SPDocumentSetSettings.psm1 
# PURPOSE: 
#	Manage Document Sets
#	Relies on an XML Configuration file to identify and handle the Lists.
#	See SPDocumentSetSettings.xml for Schema
#
# SOURCE : https://github.com/OhNotreDame/SPPS
#
##############################################################################################



<#
.SYNOPSIS
	Set Document Folder Settings
	
.DESCRIPTION
	Set Document Folder Settings
	
.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER docSetSettingsDescriptionXML
	XML Object representing the document sets to be created/updated
	
.EXAMPLE
	setDocumentFolderSettings -siteURL <SiteURL> -docSetSettingsDescriptionXML <XMLObjectToParse>
	
.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 07.06.2017
	Last Updated by: JBO
	Last Updated: 07.06.2017
#>
function setDocumentFolderSettings()
{

	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[XML]$docSetSettingsDescriptionXML
	)

    try
    {
		$functionName = $MyInvocation.MyCommand.Name
		Write-Debug "[$functionName] Entering function" 
		Write-Debug "[$functionName] Parameter / siteURL: $siteURL"
	
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
        if($curWeb -ne $null)
        {
			$ctTypes= $docSetSettingsDescriptionXML.ContentTypes.ContentType
			foreach($ctType in $ctTypes)
			{
				$ctName = $ctType.Name
				Write-Host "[$functionName] About to customize Document Set (content-type) '$ctName'." -ForegroundColor Magenta `r
				
				$ctExists = $curWeb.ContentTypes[$ctName]
				if($ctExists -ne $null)
				{					
					$currentCT = $curWeb.ContentTypes[$ctName]
					$docSet = [Microsoft.Office.DocumentManagement.DocumentSets.DocumentSetTemplate]::GetDocumentSetTemplate($currentCT)
					# Allow Content Types
					if($ctType.AllowedContentTypes -ne $null -and $ctType.AllowedContentTypes.HasChildNodes)
					{
						Write-Host "[$functionName] About to define 'AllowedContentTypes' for Content-Type '$ctName'." -ForegroundColor Cyan `r
						
						$docSet.AllowedContentTypes.Clear()
						foreach($contentTypeToAdd in $ctType.AllowedContentTypes.AllowedContentType)
						{
							$contentTypeToAddName=$curWeb.ContentTypes[$contentTypeToAdd.Name]
							$docSet.AllowedContentTypes.Add($contentTypeToAddName.Id)			
						}
                        Write-Host "[$functionName] Document Set 'AllowedContentTypes' have been set." -ForegroundColor Green `r
					}
					else
					{
						Write-Debug "[$functionName] 'AllowedContentTypes' node is NOT defined in XML for Content-Type '$ctName'."
					}
					
					# Set Shared Fields
					if($ctType.SharedFields -ne $null -and $ctType.SharedFields.HasChildNodes)
					{
						Write-Host "[$functionName] About to define 'SharedFields' for Content-Type '$ctName'." -ForegroundColor Cyan `r
						foreach($SharedField in $ctType.SharedFields.SharedField)
						{
							$SharedField.Name
                            $fieldToShare= New-Object Microsoft.SharePoint.SpField -ArgumentList @($curWeb.Fields,$SharedField.Name)
							$docSet.SharedFields.Add($fieldToShare)
						}
                        Write-Host "[$functionName] Document Set 'SharedFields' have been set." -ForegroundColor Green `r
					}
					else
					{
						Write-Debug "[$functionName] 'SharedFields' node is NOT defined in XML for Content-Type '$ctName'."
					}

					# Add Fields at Welcome page
					if($ctType.WelcomePageFields -ne $null -and $ctType.WelcomePageFields.HasChildNodes)
					{
						Write-Host "[$functionName] About to define 'WelcomePageFields' for Content-Type '$ctName'." -ForegroundColor Cyan `r
						
						foreach($welcomeField in $ctType.WelcomePageFields.WelcomePageField)
						{
							$welcomeFieldToAdd= New-Object Microsoft.SharePoint.SpField -ArgumentList @($curWeb.Fields,$welcomeField.Name)
							$docSet.welcomePageFields.Add($welcomeFieldToAdd)
						}
                        Write-Host "[$functionName] Document Set 'WelcomePageFields' have been set." -ForegroundColor Green `r
					}
					else
					{
						Write-Debug "[$functionName] 'WelcomePageFields' node is NOT defined in XML for Content-Type '$ctName'."
					}
					
					#Update DocSet Content-Type Object
					$docSet.Update($true)
					$curWeb.Update();
					Write-Host "[$functionName] Document Set settings have been fully updated." -ForegroundColor Green `r
				}
				else
				{
					Write-Debug "[$functionName] Document Set (content-type) '$ctName' does not exist."
					#Content-Type creation should have been done prior to customization, using SPSiteContentTypeModules
				}			
			}#End foreach	
		}
		else
		{
			Write-Warning "[$functionName] Site '$siteURL' could not be found." 
		}	
	}
	 catch [Exception]
    {
		Write-Host "/!\ [$functionName] An exception has been caught /!\ "  -ForegroundColor Red `r
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
	Setup default view for a List DocSet
	
.DESCRIPTION
	Setup default view for a List DocSet
	
.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER listName
	Name of the List where the DocSet must be configured
	
.PARAMETER ctName
	Name of the DocSet Content-Type to be configured
		
.PARAMETER viewName
	Name of the View to be set as default for the Welcome Page
	
.EXAMPLE
	setDocumentSetDefaultView -siteURL <SiteURL> -listName <ListName> -ctName <ContentTypeName> -viewName <ViewName>
	
.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 14.12.2020
	Last Updated by: JBO
	Last Updated: 14.12.2020
#>
function setDocumentSetDefaultView()
{

	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[string]$listName,
		[Parameter(Mandatory=$true, Position=3)]
		[string]$ctName,
		[Parameter(Mandatory=$true, Position=4)]
		[string]$viewName
	)

    try
    {
		$functionName = $MyInvocation.MyCommand.Name
		Write-Debug "[$functionName] Entering function" 
		Write-Debug "[$functionName] Parameter / siteURL: $siteURL"
	
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
        if($curWeb -ne $null)
        {
			$curList= $curWeb.Lists[$listName]
			if($curList -ne $null)
			{
				$defaultView = $curList.Views[$viewName]
				if($defaultView -ne $null)
				{
					$curCT_DocSet = $curList.ContentTypes[$ctName]
					if($curCT_DocSet -ne $null)
					{
						Write-Host "[$functionName] About to set '$viewName' as default view for the Document Set '$ctName' on list '$listName'." -ForegroundColor Cyan `r
						$docSet = [Microsoft.Office.DocumentManagement.DocumentSets.DocumentSetTemplate]::GetDocumentSetTemplate($curCT_DocSet)
						$docSet.WelcomePageView = $defaultView 
						$docSet.Update($true)
						Write-Host "[$functionName] Default view successfully set to '$viewName' on the Document Set '$ctName' of list '$listName'." -ForegroundColor Green `r
					}
					else
					{
						Write-Warning "[$functionName] Content-type '$ctName' could not be found on List '$listName'."
					}
				}
				else
				{
					Write-Warning "[$functionName] View '$viewName' could not be found on List '$listName'." 
				}
			}
			else
			{
				Write-Warning "[$functionName] List '$listName' could not be found." 
			}
		}
		else
		{
			Write-Warning "[$functionName] Site '$siteURL' could not be found." 
		}	
	}
	 catch [Exception]
    {
		Write-Host "/!\ [$functionName] An exception has been caught /!\ "  -ForegroundColor Red `r
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








