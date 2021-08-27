##############################################################################################
#              
# NAME: SPCommonStructure.psm1 
# PURPOSE: 
#	Manage Common Structure (Creation, Update of Site Columns, Site Content-Types and Lists)
#	Relies on an XML Configuration file for Site Columns description.
#	See SPCommonStructure.xml for Schema
#
##############################################################################################


<#
.SYNOPSIS

.DESCRIPTION

.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER siteColumnsDescriptionXML
	XML object of the file to parse
	
.EXAMPLE
	browseAndParseSPCommonStructureXML -siteURL <SiteURL> -SPCommonStructureDescriptionXML <siteColumnsDescriptionXML>
	
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
function browseAndParseSPCommonStructureXML()
{
	Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[XML]$SPCommonStructureDescriptionXML
	)
    
	$functionName = $MyInvocation.MyCommand.Name
	Write-Debug "Entering $functionName"
	
	try
	{
		$parentSiteURL = ""
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if ($curWeb -ne $null)
		{
			if (!($curWeb.IsRootWeb))
			{
				$parentSiteURL = $curWeb.site.RootWeb.URL
			}
			else
			{
				$parentSiteURL = $siteURL
			}
			
			$SPCommonStructure
			
			$SPCommonStructure = $SPCommonStructureDescriptionXML.SelectNodes("/SPCommonStructure")
			if($SPCommonStructure -ne $null -and $SPCommonStructure.HasChildNodes)
			{	
				#--- Add Common Columns (SPCommonStructure.Fields node)			
				Write-Host "[$functionName] About to add Common Site Columns."-ForegroundColor Gray `r
				$FieldsXML =  $SPCommonStructure.Fields			
				if($FieldsXML -ne $null -and $FieldsXML.HasChildNodes)
				{									
					foreach($siteColumn in $FieldsXML.Field)
					{
						$doesSiteColumnExist = existSiteColumn -siteURL $parentSiteURL -siteColumnID $siteColumn.ID
						if($doesSiteColumnExist -eq $false)
						{
							createSiteColumn -siteURL $parentSiteURL -fieldDefinitionXML $siteColumn
						}
						else
						{
							updateSiteColumn -siteURL $parentSiteURL -fieldDefinitionXML $siteColumn               
						}
					}#end foreach
				}
				else
				{
					Write-Warning "[$functionName] 'Fields' node is empty."
				}

				#--- Add Common Content-Types (SPCommonStructure.ContentTypes node)			
				Write-Host "[$functionName] About to add Common Content-Types."-ForegroundColor Gray `r
				$contentTypesXML =  $SPCommonStructure.ContentTypes
				if($contentTypesXML -ne $null -and $contentTypesXML.HasChildNodes)
				{									
					foreach($ctType in $contentTypesXML.ContentType)
					{
						$doesCTypeExist = existSiteContentTypeByName -siteURL $parentSiteURL -ctName $ctType.Name
						if($doesCTypeExist -eq $false)
						{
							createSiteContentType -siteURL $parentSiteURL -ctDescriptionXML $ctType -Verbose
						}
						else
						{
							updateSiteContentType -siteURL $parentSiteURL -ctDescriptionXML $ctType
						}
					}
				}
				else
				{
					Write-Warning "[$functionName] 'Content-Types' node is empty."
				}
				
				#--- Add Common Lists (SPCommonStructure.Lists node)			
				Write-Host "[$functionName] About to add Common Lists."-ForegroundColor Gray `r
				$listsXML =  $SPCommonStructure.Lists
				if($listsXML -ne $null -and $listsXML.HasChildNodes)
				{									
					foreach($list in $listsXML.List)
					{
						$listName = $list.Title.Trim()
						$doesListExist = existList -siteURL $siteURL -listName  $listName
						if($doesListExist -eq $false)
						{
							createList -siteURL $siteURL -listDefinitionXML $list
						}
						else
						{
							#Update List is not implemented yet.
							#updateList -siteURL $siteURL -listDefinitionXML $listDefinitionXML 
						}
					}#end foreach
				}
				else
				{
					Write-Warning "[$functionName] 'Lists' node is empty."
				}
				
			}
			else
			{
				Write-Warning "[$functionName] XML file is empty."
			}
		}
		else
		{
			Write-Warning "[$functionName] Site '$siteURL' does not exist."
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
		Write-Debug "Exiting $functionName"
	
	}
}