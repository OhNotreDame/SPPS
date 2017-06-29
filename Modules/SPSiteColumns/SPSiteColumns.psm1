##############################################################################################
#              
# NAME: SPSiteColumns.psm1 
# PURPOSE: 
#	Manage Site Columns (Creation, Deletion, Update)
#	Relies on an XML Configuration file for Site Columns description.
#	See SPSiteColumns.xml for Schema
#
# SOURCE : https://github.com/OhNotreDame/SPPS
#
##############################################################################################


<#
.SYNOPSIS
	Parse the file SiteColumnsDescriptionXML XML object and loop across all the Site Columns to add/update
	
.DESCRIPTION
	Will parse the file SiteColumnsDescriptionXML XML object and loop across all the Site Columns to add/update.  
	For each node, check if the field exist in the site siteURL
		If exists, call updateSiteColumn()
		If does not exist, call updateSiteColumn()
		
.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER siteColumnsDescriptionXML
	XML object of the file to parse
	
.EXAMPLE
	browseAndParseSiteColumnsXML -siteURL <SiteURL> -siteColumnsDescriptionXML <siteColumnsDescriptionXML>
	
.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 13.01.2017
	Last Updated by: JBO
	Last Updated: 13.01.2017
#>
function browseAndParseSiteColumnsXML()
{
	Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[XML]$siteColumnsDescriptionXML
	)
    
	$functionName = $MyInvocation.MyCommand.Name
	Write-Debug "Entering $functionName"
	
	try
	{
		if($siteColumnsDescriptionXML.HasChildNodes)
		{
			$siteColumns = $siteColumnsDescriptionXML.SelectNodes("/Fields")
			foreach($siteColumn in $siteColumns.Field)
			{
				$isSiteColumnExist = existSiteColumn -siteURL $siteURL -siteColumnID $siteColumn.ID
				if($isSiteColumnExist -eq $false)
				{
					createSiteColumn -siteURL $siteURL -fieldDefinitionXML $siteColumn
				}
				else
				{
					updateSiteColumn -siteURL $siteURL -fieldDefinitionXML $siteColumn               
				}
			}#end foreach
		}
		else
		{
			Write-Warning "[$functionName] XML file is empty."
		}
	}
	catch [Exception]
	{
		Write-Host "/!\ [$functionName] An exception has been caught /!\ "  -ForegroundColor Red `r
		Write-Host "Type: " $_.Exception.GetType().FullName   -ForegroundColor Red `r
		Write-Host "Message: " $_.Exception.Message  -ForegroundColor Red `r
		Write-Host "Stacktrace: " $_.Exception.Stacktrace  -ForegroundColor Red `r
	}
}

<#
.SYNOPSIS
	Check if a Site Column, with ID siteColumnID exists on site siteURL.
	
.DESCRIPTION
	Check if a Site Column, with ID siteColumnID exists on site siteURL.
	If exists, return true. If does not exist, return false.
		
.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER siteColumnID
	ID of the Site Column
	
.EXAMPLE
	existSiteColumn -siteURL <SiteURL> -siteColumnID <siteColumnID>
	
.OUTPUTS
	True, if exists
	False, if not

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 13.01.2017
	Last Updated by: JBO
	Last Updated: 13.01.2017
#>
function existSiteColumn()
{
	Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[GUID]$siteColumnID
	)
    
	$functionName = $MyInvocation.MyCommand.Name
	Write-Debug "Entering $functionName"
	Write-Debug "[$functionName] Parameter/ siteURL: $siteURL"
	Write-Debug "[$functionName] Parameter/ siteColumnID: $siteColumnID"
	
	$existsSiteColumn = $false
    try
    {
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if ($curWeb -ne $null)
		{
			# if (!($curWeb.IsRootWeb))
			# {
				# Write-Debug "[$functionName] Site '$siteURL' is not a RootWeb. Getting RootWeb."
				# $curWeb = $curWeb.site.RootWeb
			# }
			# else
			# {
				# Write-Debug "[$functionName] Site '$siteURL' is already a RootWeb."
			# }
			
			$column = $curWeb.Fields[$siteColumnID]
			if ($column)
			{
				$existsSiteColumn = $true
			}
			else
			{
				$existsSiteColumn = $false
			}
		}
		else
		{
			$existsSiteColumn = $false
			Write-Warning "[$functionName] Site '$siteURL' does not exist."
		}    
    }
    catch [Exception]
    {
		Write-Host ""
		Write-Host "/!\ [$functionName] An exception has been caught /!\ "  -ForegroundColor Red `r
		Write-Host "Type: " $_.Exception.GetType().FullName   -ForegroundColor Red `r
		Write-Host "Message: " $_.Exception.Message  -ForegroundColor Red `r
		Write-Host "Stacktrace: " $_.Exception.Stacktrace  -ForegroundColor Red `rr
        $existsSiteColumn = $false
    }
    finally
    {
        if($curWeb -ne $null)
        {
           $curWeb.Dispose()
        }
		Write-Debug "Exiting $functionName"
    }
    return $existsSiteColumn
}


<#
.SYNOPSIS
	Check if a Site Column $columnName exists on site siteURL.
	
.DESCRIPTION
	Check if a Site Column, with name columnName exists on site siteURL.
	If exists, return the Site Columns. If does not exist, return null.
		
.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER columnName
	Name of the Site Column
	
.EXAMPLE
	getSiteColumnByName -siteURL <SiteURL> -columnName <columnName>
	
.OUTPUTS
	True, if exists
	False, if not

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 13.01.2017
	Last Updated by: JBO
	Last Updated: 13.01.2017
#>
function getSiteColumnByName()
{
	Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[string]$columnName
	)
    
	$functionName = $MyInvocation.MyCommand.Name
	Write-Debug "Entering $functionName"
	Write-Debug "[$functionName] Parameter/ siteURL: $siteURL"
	Write-Debug "[$functionName] Parameter/ columnName: $columnName"
	
	$siteColumn = $null
    try
    {
	
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if ($curWeb -ne $null)
		{
			# if (!($curWeb.IsRootWeb))
			# {
				# Write-Debug "[$functionName] Site '$siteURL' is not a RootWeb. Getting RootWeb."
				# $curWeb = $curWeb.site.RootWeb
			# }
			# else
			# {
				# Write-Debug "[$functionName] Site '$siteURL' is already a RootWeb."
			# }
			
			$siteColumn = $curWeb.Fields[$columnName]
		}
		else
		{
			$siteColumn = $null
			Write-Warning "[$functionName] Site '$siteURL' does not exist."
		}    
    }
    catch [Exception]
    {
		Write-Host ""
		Write-Host "/!\ [$functionName] An exception has been caught /!\ "  -ForegroundColor Red `r
		Write-Host "Type: " $_.Exception.GetType().FullName   -ForegroundColor Red `r
		Write-Host "Message: " $_.Exception.Message  -ForegroundColor Red `r
		Write-Host "Stacktrace: " $_.Exception.Stacktrace  -ForegroundColor Red `r
		$siteColumn = $null
    }
    finally
    {
        if($curWeb -ne $null)
        {
           $curWeb.Dispose()
        }
		Write-Debug "Exiting $functionName"
    }
    return $siteColumn
 }



<#
.SYNOPSIS
	Check if a Site Column, with ID siteColumnID exists on site siteURL, and return its Static Name
	
.DESCRIPTION
	Check if a Site Column, with ID siteColumnID exists on site siteURL, and return its Static Name
	If exists, return its Static Name. If does not exist, return Empty.
		
.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER siteColumnID
	ID of the Site Column
	
.EXAMPLE
	getStaticNameFromColumnID -siteURL <SiteURL> -siteColumnID <siteColumnID>
	
.OUTPUTS
	Return Static Name (if exists)
	Return Empty String (if not)

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 13.01.2017
	Last Updated by: JBO
	Last Updated: 13.01.2017
#>
function getStaticNameFromColumnID()
{
	Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[GUID]$siteColumnID
	)
    
	$functionName = $MyInvocation.MyCommand.Name
	Write-Debug "Entering $functionName"
	Write-Debug "[$functionName] Parameter/ siteURL: $siteURL"
	Write-Debug "[$functionName] Parameter/ siteColumnID: $siteColumnID"
	
	$staticName = [String]::Empty
    try
    {
		$curWeb = GetSPWebBySiteUrl -siteUrl $siteURL 
		if($curWeb -ne $null)
		{
			$column = $curWeb.Fields[$siteColumnID]
			if ($column)
			{
				$staticName = $column.StaticName
				Write-Debug "[$functionName] Parameter/ staticName: $staticName"
			}
			else
			{
				Write-Warning "[$functionName] Column '$siteColumnID' does not exist."
				$staticName = ""
			}
		}
		else
		{
			Write-Warning "[$functionName] Site '$siteURL' does not exist."
			$staticName = ""
		}      
    }
    catch [Exception]
    {
		Write-Host ""
		Write-Host "/!\ [$functionName] An exception has been caught /!\ "  -ForegroundColor Red `r
		Write-Host "Type: " $_.Exception.GetType().FullName   -ForegroundColor Red `r
		Write-Host "Message: " $_.Exception.Message  -ForegroundColor Red `r
		Write-Host "Stacktrace: " $_.Exception.Stacktrace  -ForegroundColor Red `r
		$staticName = ""
    }
    finally
    {
        if($curWeb -ne $null)
        {
			$curWeb.Dispose()
        }
		Write-Debug "Exiting $functionName"
    }
    return $staticName
 }


<#
.SYNOPSIS
	Create the Site Column based on its XML definition fieldDefinitionXML on site siteURL.
	
.DESCRIPTION
	Create the Site Column based on its XML definition fieldDefinitionXML on site siteURL.
		
.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER siteColumnID
	XML description of the Site Column to create
	
.EXAMPLE
	createSiteColumn -siteURL <SiteURL> -fieldDefinitionXML <fieldDefinitionXML>
	
.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 13.01.2017
	Last Updated by: JBO
	Last Updated: 13.01.2017
#>
function createSiteColumn()
{
	Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[System.XML.XmlElement]$fieldDefinitionXML
		)
    
	$functionName = $MyInvocation.MyCommand.Name
	Write-Debug "Entering $functionName"
	Write-Debug "[$functionName] Parameter/ siteURL: $siteURL"
	
    try
    {
		
		$siteColumnName = $fieldDefinitionXML.DisplayName 
		Write-Debug "[$functionName] Parameter/ siteColumnName: $siteColumnName"
		
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if ($curWeb -ne $null)
		{
			# if (!($curWeb.IsRootWeb))
			# {
				# Write-Debug "[$functionName] Site '$siteURL' is not a RootWeb. Getting RootWeb."
				# $curWeb = $curWeb.site.RootWeb
			# }
			# else
			# {
				# Write-Debug "[$functionName] Site '$siteURL' is already a RootWeb."
			# }
			
			$exists = existSiteColumn -siteURL $curWeb.URL -siteColumnID $fieldDefinitionXML.ID
			if ($exists -eq $false)
			{
				Write-Host "[$functionName] About to Create Site Column '$siteColumnName'." -ForegroundColor Magenta `r
				$curWeb.Fields.AddFieldAsXml($fieldDefinitionXML.OuterXml)
				Write-Host "[$functionName] Site Column '$siteColumnName' successfully Created."  -ForegroundColor Green `r
			}
			else
			{
				Write-Warning "[$functionName] Site Column '$siteColumnName' already exists."
			}
		}
		else
		{
			$siteColumn = $null
			Write-Warning "[$functionName] Site '$siteURL' does not exist."
		} 	
    }
    catch [Exception]
    {
		Write-Host ""
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
		Write-Debug "Exiting $functionName"
    }
 }

 
<#
.SYNOPSIS
	Update the Site Column based on its XML definition fieldDefinitionXML on site siteURL.
	
.DESCRIPTION
	Update the Site Column based on its XML definition fieldDefinitionXML on site siteURL.
		
.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER siteColumnID
	XML description of the site column to create
	
.EXAMPLE
	updateSiteColumn -siteURL <SiteURL> -fieldDefinitionXML <fieldDefinitionXML>
	
.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 13.01.2017
	Last Updated by: JBO
	Last Updated: 13.01.2017
#>
function updateSiteColumn()
{
	Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[System.XML.XmlElement]$fieldDefinitionXML
	)
    
	$functionName = $MyInvocation.MyCommand.Name
	Write-Debug "Entering $functionName"
	Write-Debug "[$functionName] Parameter/ siteURL: $siteURL"
	
    try
    {
		$siteColumnName = $fieldDefinitionXML.DisplayName 
		Write-Debug "[$functionName] Parameter/ siteColumnName: $siteColumnName"
		
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if ($curWeb -ne $null)
		{
			# if (!($curWeb.IsRootWeb))
			# {
				# Write-Debug "[$functionName] Site '$siteURL' is not a RootWeb. Getting RootWeb."
				# $curWeb = $curWeb.site.RootWeb
			# }
			# else
			# {
				# Write-Debug "[$functionName] Site '$siteURL' is already a RootWeb."
			# }
			
			$column = $curWeb.Fields[[GUID]$fieldDefinitionXML.ID]
			if($column)
			{
				Write-Host "[$functionName] About to update Site Column '$siteColumnName'."  -ForegroundColor Magenta `r
				$column.SchemaXml = $fieldDefinitionXML.OuterXml 
				$column.Update()
				Write-Host "[$functionName] Site Column '$siteColumnName' successfully updated."  -ForegroundColor Green `r
			}
			else
			{
				Write-Warning "[$functionName] Site Column '$siteColumnName' does not exist."
			}
		}
		else
		{
			Write-Warning "[$functionName] Site '$siteURL' does not exist."
		} 
		
    }
    catch [Exception]
    {
		Write-Host ""
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
		Write-Debug "Exiting $functionName"
    }
}


####################################################################
######################### LAST TO REFACTOR #########################
####################################################################
<#
.SYNOPSIS
	Delete Site Column siteColumnID from site siteURL.
	
.DESCRIPTION
	Delete Site Column siteColumnID from site siteURL.
	i.e. remove it from all content-types, all lists and libraries from all the SPWeb of the Site Collection.
	
.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER siteColumnID
	ID of the site column to delete
	
.EXAMPLE
	deleteSiteColumn -siteURL <SiteURL> -siteColumnID <siteColumnID>
	
.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 13.01.2017
	Last Updated by: JBO
	Last Updated: 13.01.2017
#>
function deleteSiteColumn()
{
	Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[GUID]$siteColumnID
	)
    
	$functionName = $MyInvocation.MyCommand.Name
	Write-Debug "Entering $functionName"
	
    try
    {

		$exists = existSiteColumn -siteURL $siteURL -siteColumnID $siteColumnID
		if ($exists -eq $true)
		{
			#Write-Host "[$functionName] Site Column '$siteColumnID' exists on site '$siteURL'." -ForegroundColor White `r

			
			$isRootWeb = isRootWeb -siteUrl $siteURL
			if($isRootWeb -eq $true)
			{
				#Write-Host "[$functionName] '$siteURL' is RootWeb." -ForegroundColor Magenta `r
				$rootWeb = GetSPWebBySiteUrl -siteUrl $siteURL 		
				$curSPSite = Get-SPSite $siteUrl	
			}
			else
			{
				#Write-Host "[$functionName] '$siteURL' is SubWeb. Getting site.RootWeb." -ForegroundColor Magenta `r
				$curWeb = GetSPWebBySiteUrl -siteUrl $siteURL 
				$rootWeb = $curWeb.site.RootWeb
				$curSPSite = $curWeb.site
			}
				
			$column = $rootWeb.Fields[[GUID]$siteColumnID]
			$columnDisplayName = $column.Title
			$columnStaticName = $column.StaticName
			$columnInternalName = $column.InternalName
			

		
			if ($column.UsedInWebContentTypes -eq $true)		
			{
				Write-Host "[$functionName] 'Site Column '$columnDisplayName' is implemented in any site collection content type" -ForegroundColor Magenta `r
			
				foreach($loopWeb in $curSPSite.AllWebs)
				{
					Write-Host "[$functionName] About to delete Site Column '$columnDisplayName' from SPWeb '" $loopWeb.Url "' ... " -ForegroundColor Magenta `r
					
					#Delete all references to the field in lists
					foreach($curList in $loopWeb.Lists)
					{
						#Fucntion to Write removeFieldFromList
						$listFieldToRemove=$curList.Fields.TryGetFieldByStaticName($columnStaticName)
						if($listFieldToRemove -ne $null)
						{
							#Write-Host "[$functionName] Deleting Site Column '$columnDisplayName' from List " $curList.Title -ForegroundColor White `r
							removeSiteColumnFromList -siteURL $loopWeb.Url -listName $curList.Title -fieldStaticName $columnStaticName
							Write-Host "[$functionName] Site Column '$columnDisplayName' succesfully deleted from Content Type " $curList.Name -ForegroundColor Green `r
						}

						#Delete field from all content types of this list
						foreach($ct in $curList.ContentTypes) 
						{
							$fieldInCT=$ct.Fields.TryGetFieldByStaticName($columnStaticName)
							if($fieldInCT) 
							{
								#Write-Host "[$functionName] Deleting Site Column '$columnDisplayName' from Content Type " $ct.Name -ForegroundColor White `r
								removeSiteColumnFromContentType -siteURL $loopWeb.Url -ct $ct -fieldStaticName $columnStaticName
								Write-Host "[$functionName] Site Column '$columnDisplayName' succesfully deleted from Content Type " $ct.Name -ForegroundColor Green `r
							}
						}
						
					}
					
					#Delete field from all content types
					foreach($ct in $loopWeb.ContentTypes) 
					{
						$fieldInCT=$ct.Fields.TryGetFieldByStaticName($columnStaticName)
						if($fieldInCT) 
						{
							#Write-Host "[$functionName] Deleting Site Column '$columnDisplayName' from Content Type " $ct.Name -ForegroundColor White `r
							if ($ct.Sealed -eq $false)
							{
							
								#Write-Host "[$functionName] {AllWebs} - Deleting Site Column '$columnDisplayName' from Content Type " $ct.Name -ForegroundColor Magenta `r
								removeSiteColumnFromContentType -siteURL $loopWeb.Url -ct $ct -fieldStaticName $columnStaticName
								Write-Host "[$functionName] {AllWebs} - Site Column '$columnDisplayName' succesfully deleted from Content Type " $ct.Name -ForegroundColor Green `r
							}
							else
							{
								Write-Host "[$functionName] {AllWebs} - Site Content Type"$ct.Name"is sealed" -ForegroundColor Cyan `r
							}						
						}
						else
						{
							#Write-Host "[$functionName] {AllWebs} - Site Column '$columnDisplayName' not present in Content Type " $ct.Name -ForegroundColor Cyan `r
						}
					}
				
					$loopWeb.update();
					#EndLoop on SubWeb
				}
				
				Write-Host "[$functionName] About to delete '$columnDisplayName' on each CTs, Lists of every Subsites of '$siteURL'." -ForegroundColor Magenta `r
				$rootWeb = GetSPWebBySiteUrl -siteUrl $siteURL 
				
				#Delete field from all content types
				foreach($ct in $rootWeb.ContentTypes) 
				{				
					$fieldInCT=$ct.Fields.TryGetFieldByStaticName($columnStaticName)
					if($fieldInCT) 
					{
						#Write-Host "[$functionName] Deleting Site Column '$columnDisplayName' from Content Type " $ct.Name -ForegroundColor White `r
						if ($ct.Sealed -eq $false)
						{
							Write-Host "[$functionName] {RootWeb} - Deleting Site Column '$columnDisplayName' from Content Type " $ct.Name -ForegroundColor White `r
							removeSiteColumnXMLFromContentType -siteURL $rootWeb.Url -ct $ct -fieldStaticName $columnInternalName
							Write-Host "[$functionName] {RootWeb} - Site Column '$columnDisplayName' succesfully deleted from Content Type " $ct.Name -ForegroundColor Green `r
						}
						else
						{
							Write-Host "[$functionName] {SubWeb} - Site Content Type"$ct.Name"is sealed" -ForegroundColor Cyan `r
						}						
					}
					else
					{
						Write-Host "[$functionName] {SubWeb} - Site Column '$columnDisplayName' not present in Content Type " $ct.Name -ForegroundColor Cyan `r
					}
				}

				
				$rootWeb = GetSPWebBySiteUrl -siteUrl $siteURL 
				$column = $rootWeb.Fields[[GUID]$siteColumnID]
				if ($column.UsedInWebContentTypes -eq $true)		
				{
					Write-Warning "[$functionName] '$columnDisplayName' is STILL implemented in any site collection content type"
				}
				else
				{
					#Write-Host "[$functionName] About to delete SiteColumn '$columnDisplayName' from RootWeb." -ForegroundColor Magenta `r
					$rootWeb.Fields.Delete($columnStaticName)
					$rootWeb.Update()
					Write-Host "[$functionName] SiteColumn '$columnDisplayName' succesfully deleted from RootWeb." -ForegroundColor Green `
				}
			}
			else
			{
				Write-Host "[$functionName] '$columnDisplayName' is NOT implemented in any site collection content type" -ForegroundColor Cyan `r
				$rootWeb.Fields.Delete($columnStaticName)
				$rootWeb.Update()
				Write-Host "[$functionName] SiteColumn '$columnDisplayName' succesfully deleted from RootWeb." -ForegroundColor Green `
			}
		}
		else
		{
			Write-Host "[$functionName] '$siteColumnID' does not exist on site '$siteURL'." -ForegroundColor Cyan `r
		}
    }
    catch [Exception]
    {
		Write-Host ""
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
		Write-Debug "Exiting $functionName"
    }
 }
