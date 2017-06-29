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
				
				$ctExists = $curWeb.AvailableContentTypes[$ctName]
				if($ctExists -eq $true)
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
							$fieldToShare= New-Object Microsoft.SharePoint.SpField -ArgumentList @($curWeb.Fields,$SharedField.Name)
							$docSet.SharedFields.Add($fieldToShare)
						}
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
					}
					else
					{
						Write-Debug "[$functionName] 'WelcomePageFields' node is NOT defined in XML for Content-Type '$ctName'."
					}
					
					#Update DocSet Content-Type Object
					$docSet.Update($true)
					$curWeb.Update();
					Write-Host "[$functionName] Document Set settings have been updated." -ForegroundColor Green `r
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