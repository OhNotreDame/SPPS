##############################################################################################
#              
# NAME: SPLookupFields.psm1 
# PURPOSE: 
#	Manage Lookups
#	Relies on an XML Configuration file to identify and handle the Lists.
#	See SPLookupFields.xml for Schema
#
# SOURCE : https://github.com/OhNotreDame/SPPS
#
##############################################################################################


<#
.SYNOPSIS
	Browse and parse the SiteLookupsDescriptionXML XML object
	
.DESCRIPTION
	Browse and parse the file SiteLookupsDescriptionXML
		For each node, will initiate Lookup creation	
	
.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER lookupsDescriptionXML
	XML Object representing ALL the lookups to be created/updated
	
.EXAMPLE
	browseAndParseSiteLookupsXML -siteURL <SiteURL> -lookupsDescriptionXML <XMLObjectToParse>
	
.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 06.03.2017
	Last Updated by: JBO
	Last Updated: 06.03.2017
#>
function browseAndParseSiteLookupsXML()
{ 
    Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[xml]$lookupsDescriptionXML
	)

	try
    {		
		$functionName = $MyInvocation.MyCommand.Name
		Write-Debug "[$functionName] Entering function" 
		Write-Debug "[$functionName] Parameter / siteURL: $siteURL"
     
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
        if($curWeb -ne $null)
		{
			$LookupFieldXml = $lookupsDescriptionXML.Lookups
			if($LookupFieldXml -ne $null -and $LookupFieldXml.HasChildNodes)
			{
				foreach($lookupInfo in $LookupFieldXml.Lookup)
				{
					Write-Debug "[$functionName] lookupInfo:`n$lookupInfo"
					$listName = $lookupInfo.Source.ListName
					$list = $curWeb.Lists.TryGetList($listName)
					if($list -ne $null)
					{
						$lkpStaticName = $lookupInfo.Field.StaticName
						$field = $curWeb.Fields.TryGetFieldByStaticName($lkpStaticName)
						if($field -eq $null)
						{
							Write-Debug "[$functionName] - calling createSiteLookup" 
							createSiteLookup -siteURL $siteURL -lookupDefinitionXML $lookupInfo
						}
						else
						{				
							#-NOT IMPLEMENTED- Update the Lookup Column
							#updateSiteLookup -siteURL $siteURL -lookupDefinitionXML $lookupInfo
						}
					}
					else
					{
						Write-Host "[$functionName] Lookup list '$listName' does not exist in site '$siteURL'"
					}
				}#foreach
			}
			else
			{
				Write-Warning "[$functionName] Lookups XML defintion is empty."
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
	Browse and parse the browseAndParseSiteLookupsAssociationsXML XML object
	
.DESCRIPTION
	Browse and parse the file browseAndParseSiteLookupsAssociationsXML
		For each node, will association with the related List or Content-Type	
	
.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER $lookupsAssociationsXML
	XML Object representing ALL the lookups to be associated with List or Content-Type
	
.EXAMPLE
	browseAndParseSiteLookupsAssociationsXML -siteURL <SiteURL> -$lookupsAssociationsXML <XMLObjectToParse>
	
.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 08.03.2017
	Last Updated by: JBO
	Last Updated: 08.03.2017
#>
function browseAndParseSiteLookupsAssociationsXML()
{ 
    Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[xml]$lookupsAssociationsXML
	)

	try
    {		
		$functionName = $MyInvocation.MyCommand.Name
		Write-Debug "[$functionName] Entering function" 
		Write-Debug "[$functionName] Parameter / siteURL: $siteURL"
     
		$lookupsAssociationsXML

		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
        if($curWeb -ne $null)
		{

			$allAssocXML = $lookupsAssociationsXML.SelectNodes("/LookupsAssociations")
			if($allAssocXML.HasChildNodes)
			{
				#--- Associate Lookups to Content-Types
				Write-Host "[$functionName] About to Associate Lookups to Content-Types."-ForegroundColor Cyan `r
				$lkpToAssocToCTs =  $allAssocXML.AssociateLookupToContentTypes
				foreach($lkpToAssoc in $lkpToAssocToCTs.Lookup)
				{
					$lkpToAssoc.ContentTypeName
					addSiteLookupToSiteContentType -siteURL $siteURL -ctName $lkpToAssoc.ContentTypeName -lookupStaticName $lkpToAssoc.InternalName 
				}

				#--- Associate Lookups to Lists
				Write-Host "[$functionName] About to Associate Lookups to Lists."-ForegroundColor Cyan `r
				$lkpToAssocToList =  $allAssocXML.AssociateLookupToLists
				foreach($lkpToAssoc in $lkpToAssocToList.Lookup)
				{
					$lkpToAssoc.ContentTypeName
					addSiteLookupToList -siteURL $siteURL -listName $lkpToAssoc.ListName -lookupStaticName $lkpToAssoc.InternalName 
				}

			}
			else
			{
				Write-Warning "[$functionName] Lookups XML defintion is empty."
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
	Create the Lookup Site Column and all of it dependant fields
	
.DESCRIPTION
	Create the Lookup Site Column and all of it dependant fields
	
.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER lookupDefinitionXML
	XML Object representing the lookup to be created/updated
	
.EXAMPLE
	createSiteLookup -siteURL <SiteURL> -lookupDefinitionXML <lookupDefinitionXML>
	
.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 06.03.2017
	Last Updated by: JBO
	Last Updated: 06.03.2017
#>
function createSiteLookup()
{
    Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[XML.XMLElement]$lookupDefinitionXML
	)
    
	try
	{			
		$functionName = $MyInvocation.MyCommand.Name
		Write-Debug "[$functionName] Entering function" 
		
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
        if($curWeb -ne $null)
		{
			$siteColumnName = $lookupDefinitionXML.Field.DisplayName
			$siteColumnStaticName =  $lookupDefinitionXML.Field.StaticName
			$siteColumnRequired = [System.Convert]::ToBoolean($lookupDefinitionXML.Field.Required)

			$exists = existSiteColumn -siteURL $siteURL -siteColumnID $lookupDefinitionXML.Field.ID
			if ($exists -eq $false)
			{
				# Get Parent List Info
				
				# IF SOURCE
				$lkpDisplayName =  $lookupDefinitionXML.Source.DisplayName
				$sourceListName = $lookupDefinitionXML.Source.ListName
				$sourceFieldName = $lookupDefinitionXML.Source.FieldName
				# Get Parent List
				
				$sourceList =   getList -siteURL $siteURL -listName $sourceListName
				if($sourceList -ne $null)
                {
					$sourceListID = $sourceList.ID
					Write-Host "[$functionName] About to Create Lookup (Site Column) '$siteColumnName' based on its XML definition." -ForegroundColor Magenta  `r
					# .... not working .... $curWeb.Fields.AddFieldAsXml($lookupDefinitionXML.Field.OuterXml)
					$curWeb.Fields.AddLookup($siteColumnStaticName,$sourceListID,$siteColumnRequired)  
					$curWeb.Update()
					
					$lookupField = $curWeb.Fields.TryGetFieldByStaticName($siteColumnStaticName)
					if ($lookupField -ne $null)
					{
						Write-Host "[$functionName] Lookup (Site Column) '$siteColumnName' successfully created based on its XML definition."  -ForegroundColor Green `r
						#
						#
						Write-Host "[$functionName] About to Customize Lookup (Site Column) '$siteColumnName' based on its {Source} XML definition." -ForegroundColor Magenta  `r
						$lookupField.LookupField = $sourceList.Fields[$sourceFieldName]
						$lookupField.Title = $siteColumnName
						$lookupField.Group = $lookupDefinitionXML.Field.Group;
						$lookupField.Update()
						Write-Host "[$functionName] Lookup (Site Column) '$siteColumnName' successfully created based on its {Source} XML definition."  -ForegroundColor Green `r
						#
						#
						
						if ($lookupDefinitionXML.Source.DependantFields.HasChildNodes)
						{
							#if has a childnodes
							foreach($carryOnField in $lookupDefinitionXML.Source.DependantFields.DependantField)
							{
								$srcFieldName = $carryOnField.srcFieldName
								$dstFieldStaticName = $carryOnField.dstStaticName
								$dstFieldDisplayName = $carryOnField.dstDisplayName

								Write-Host "[$functionName] About to add Dependant Field '$srcFieldName' to Lookup (Site Column) '$siteColumnName'." -ForegroundColor Magenta  `r
								$depLookUp = $curWeb.Fields.AddDependentLookup($dstFieldStaticName, $lookupField.Id)
                                $fieldDepLookup = [Microsoft.SharePoint.SPFieldLookup] $curWeb.Fields.GetFieldByInternalName($depLookUp)
                                $fieldDepLookup.LookupField = $sourceList.Fields[$srcFieldName]
								$fieldDepLookup.Title = $dstFieldDisplayName
								$fieldDepLookup.Update()
								Write-Host "[$functionName] Dependant Field '$srcFieldName' successfully added to Lookup (Site Column) '$siteColumnName'." -ForegroundColor Green `r
							}#foreach							
						}
						else
						{
							Write-Warning "[$functionName] No Dependant Fields to add to Lookup (Site Column) '$siteColumnName'."
						}						
					}
					else
					{
						Write-Warning "[$functionName] Lookup (Site Column) '$siteColumnName' not retrieved after creation."
					}
				}
				else
				{
					Write-Warning "[$functionName] Parent List '$listName' does not exist on site '$siteURL'."
				} 
			}
			else
			{
				Write-Warning "[$functionName] Site Column '$siteColumnName' already exists."
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
	-NOT IMPLEMENTED- Update the Lookup Column
	
.DESCRIPTION
	-NOT IMPLEMENTED- Update the Lookup Column
	
.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER lookupDefinitionXML
	XML Object representing the lookup to be created/updated
	
.EXAMPLE
	updateSiteLookup -siteURL <SiteURL> -lookupDefinitionXML <lookupDefinitionXML>
	
.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 06.03.2017
	Last Updated by: JBO
	Last Updated: 06.03.2017
#>
function updateSiteLookup()
{
    Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[Xml.XmlElement]$lookupDefinitionXML
	)

    try
    {
        $functionName = $MyInvocation.MyCommand.Name
		Write-Debug "[$functionName] Entering function" 
		
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
        if($curWeb -ne $null)
        {
				Write-Host "#-----------------------------------------#"  -ForegroundColor Red `r
				Write-Host "/!\ $functionName not implemented /!\ "  -ForegroundColor Red `r
				Write-Host "#-----------------------------------------#"  -ForegroundColor Red `r
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
	Delete the Lookup Column, relying on function deleteSiteColumn() from SPSiteColumns  
	
.DESCRIPTION
	-NOT FULLY FUNCTIONAL-
	Delete the Lookup Column, relying on function deleteSiteColumn() from SPSiteColumns 
	
.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER lookupFieldID
	ID of the Site Column to be deleted
	
.EXAMPLE
	deleteSiteLookup -siteURL <SiteURL> -lookupFieldID <lookupFieldID>
	
.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 06.03.2017
	Last Updated by: JBO
	Last Updated: 06.03.2017
#>
function deleteSiteLookup()
{
    Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[string]$lookupFieldID
	)

    try
    {
		$functionName = $MyInvocation.MyCommand.Name
		Write-Debug "[$functionName] Entering function" 
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
        if($curWeb -ne $null)
        {
			$siteColumnID = [GUID]$lookupFieldID
			$existLookup = existSiteColumn -siteURL $siteURL -siteColumnID $siteColumnID
			if($existLookup)
			{
				deleteSiteColumn -siteURL $siteURL -siteColumnID $siteColumnID
            }
            else
            {
                Write-Warning "[$functionName] Lookup (Site Column) '$siteColumnID' does not exist on site '$siteURL'."
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
	Add the Lookup Site Column to the content-type $ctName
	
.DESCRIPTION
	Add the Lookup Site Column to the content-type $ctName
	
.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER ctName
	Name of the content-type

.PARAMETER lookupStaticName
	StaticName of the Lookup Site Column
	
.EXAMPLE
	addSiteLookupToSiteContentType -siteURL <SiteURL> -ctName <ctName> -lookupStaticName <lookupStaticName>
	
.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 06.03.2017
	Last Updated by: JBO
	Last Updated: 06.03.2017
#>
function addSiteLookupToSiteContentType()  
{
    Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[string]$ctName,
		[Parameter(Mandatory=$true, Position=3)]
		[string]$lookupStaticName
	)	
	
	try
	{
		$functionName = $MyInvocation.MyCommand.Name
		Write-Debug "[$functionName] Entering function" 
		
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue 
        if($curWeb -ne $null)
        {
			$listFieldToAdd=$curWeb.Fields.TryGetFieldByStaticName($lookupStaticName)
			if($listFieldToAdd -ne $null)
			{
				$ctToUpdate = $curWeb.ContentTypes[$ctName]
				if ($ctToUpdate -ne $null)
				{
					addContentTypeFieldWithStaticName -siteURL $siteURL -ct $ctToUpdate -fieldStaticName $lookupStaticName
				}
				else
				{
					Write-Warning "[$functionName] Content-Type '$ctName' does not exist on site '$siteURL'."
				}
			}
			else
			{
				Write-Warning "[$functionName] Lookup (Site Column) '$lookupStaticName' does not exist on site '$siteURL'."
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
	Remove the Lookup Site Column from the content-type $ctName
	
.DESCRIPTION
	Remove the Lookup Site Column from the content-type $ctName
	
.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER ctName
	Name of the content-type

.PARAMETER lookupStaticName
	StaticName of the Lookup Site Column
	
.EXAMPLE
	removeSiteLookupToSiteContentType -siteURL <SiteURL> -ctName <ctName> -lookupStaticName <lookupStaticName>
	
.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 06.03.2017
	Last Updated by: JBO
	Last Updated: 06.03.2017
#>
function removeSiteLookupFromSiteContentType()  
{
    Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[string]$ctName,
		[Parameter(Mandatory=$true, Position=3)]
		[string]$lookupStaticName
	)
	
	try
	{
		$functionName = $MyInvocation.MyCommand.Name
		Write-Debug "[$functionName] Entering function" 
		
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue 
        if($curWeb -ne $null)
        {
			$listFieldToAdd=$curWeb.Fields.TryGetFieldByStaticName($lookupStaticName)
			if($listFieldToAdd -ne $null)
			{
				$ct = getSiteContentTypeByName -siteURL $siteURL -ctName $ctName
				if($ct -ne $null)
				{
					removeContentTypeFieldWithStaticName -siteURL $siteURL -ct $ct -fieldStaticName $lookupStaticName
				}
				else
				{
					Write-Warning "[$functionName] Content-Type '$ctName' does not exist on site '$siteURL'."
				}
			}
			else
			{
				Write-Warning "[$functionName] Lookup (Site Column) '$lookupStaticName' does not exist on site '$siteURL'."
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
	Add a Site Lookup to the list $listName
	
.DESCRIPTION
	Add a Site Lookup to the list $listName
	
.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER listName
	Name of the List

.PARAMETER lookupStaticName
	StaticName of the Lookup Site Column
	
.EXAMPLE
	addSiteLookupToList -siteURL <SiteURL> -listName <listName> -lookupStaticName <lookupStaticName>
	
.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 06.03.2017
	Last Updated by: JBO
	Last Updated: 06.03.2017
#>
function addSiteLookupToList()  
{
    Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[string]$listName,
		[Parameter(Mandatory=$true, Position=3)]
		[string]$lookupStaticName
	)
 
	try
    {
        $functionName = $MyInvocation.MyCommand.Name
		Write-Debug "[$functionName] Entering function" 
		
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue 
        if($curWeb -ne $null)
        {
            $list = getList -siteURL $siteURL -listName $listName
            if($list -ne $null)
            {
				$listName = $list.Title 
			
				$listFieldToAdd=$curWeb.Fields.TryGetFieldByStaticName($lookupStaticName)
				if($listFieldToAdd -ne $null)
				{
					addSiteColumnToList  -siteURL $siteURL -listName $listName -fieldStaticName $lookupStaticName
				}
				else
				{
					Write-Warning "[$functionName] Lookup (Site Column) '$lookupStaticName' does not exist on site '$siteURL'."
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

