##############################################################################################
#              
# NAME: SPListColumnsTranslation.psm1 
# PURPOSE: 
#	Translate List Column Names
#	Relies on an XML Configuration file for column names
#	See SPListColumnsTranslation.xml for Schema
#
# SOURCE : https://github.com/OhNotreDame/SPPS
#
##############################################################################################

<#
.SYNOPSIS
	Change the display name of list columns. Relies on an XML Configuration file for column name
	
.DESCRIPTION
	Change the display name of list columns. Relies on an XML Configuration file for column name
		
.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER listColumnsTransXML
  XML object with list of columns name.

.EXAMPLE
	browseAndParseListColumnsTranslationXML -siteURL <SiteURL> -listColumnsTransXML <listColumnsTransXML>
	
.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 04.05.2017
	Last Updated by: JBO
	Last Updated: 04.05.2017
#>
function browseAndParseListColumnsTranslationXML()
{
	Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[XML]$listColumnsTransXML
	)
	
	$functionName = $MyInvocation.MyCommand.Name	
	Write-Debug "Entering $functionName"
	Write-Debug "[$functionName] Parameter / siteURL: $siteURL"
	
	try
	{		
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if ($curWeb -ne $null)
		{
			# main node SPListColumnsTranslation
			$listColumnsTrans = $listColumnsTransXML.SelectNodes("/SPListColumnsTranslation")
			if(($listColumnsTrans.Count -gt 0) -and $listColumnsTrans.HasChildNodes)
			{
				# Parse lists
				foreach($list in $listColumnsTrans.List)
				{
					$listName = $list.Attributes["listName"].value
					if(-Not $list.HasChildNodes)
					{
						Write-Warning "[$functionName] 'SPListColumnsTranslation/List[$listName]' is empty."
						continue
					}
					# Parse ListColumn
					foreach($listColumn in $list.ListColumn)
					{
						$internalName = $listColumn.Attributes["internalName"].value
						if(-Not $listColumn.HasChildNodes)
						{
							Write-Warning "[$functionName] 'SPListColumnsTranslation/List[$listName]/ListColumn[$internalName]' is empty."
							continue
						}
						# Parse Translation
						foreach($translation in $listColumn.Translation)
						{
							$culture = $translation.Attributes["culture"].value
							$displayName = $translation.InnerText
							if(($culture.Length -eq 0) -or ($displayName.Length -eq 0))
							{
								Write-Warning "[$functionName] 'SPListColumnsTranslation/List[$listName]/ListColumn[$internalName]/Translation' empty or missing culture [$culture] or display name [$displayName]."
								continue
							}
							# Translate column
							Write-Host "[$functionName] About to call translateColumnDisplayNameOnList -siteURL $siteURL -listName $listName -fieldStaticName $internalName -newDisplayName $displayName -culture $culture"
							translateColumnDisplayNameOnList -siteURL $siteURL -listName $listName -fieldStaticName $internalName -newDisplayName $displayName -culture $culture
						}
					}
				}
			}
			else
			{
				Write-Warning "[$functionName] 'SPListColumnsTranslation' XML node is missing or empty."
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

<#
.SYNOPSIS
	Change the display name of the field $fieldStaticName on list $listName
	
.DESCRIPTION
	Change the display name of the field $fieldStaticName on list $listName

		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[string]$listName,	
		[Parameter(Mandatory=$true, Position=3)]
		[string]$fieldStaticName,
		[Parameter(Mandatory=$true, Position=4)]
		[string]$newDisplayName,
		[Parameter(Mandatory=$true, Position=5)]
		[string]$culture
	
.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER listName
	Name of the list
	
.PARAMETER fieldStaticName
	Static Name of the SiteColumn

.PARAMETER newDisplayName
	New Display Name of the SiteColumn

.PARAMETER culture
  culture where newDisplayName will be changed

.EXAMPLE
	changeColumnDisplayNameOnList -siteURL <SiteURL> -listName <ListName> -fieldStaticName <fieldStaticName> -newDisplayName <newDisplayName> -culture <culture>
	
.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 04.05.2017
	Last Updated by: JBO
	Last Updated: 04.05.2017
#>
function translateColumnDisplayNameOnList()  
{
  [CmdletBinding()]
	Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[string]$listName,	
		[Parameter(Mandatory=$true, Position=3)]
		[string]$fieldStaticName,
		[Parameter(Mandatory=$true, Position=4)]
		[string]$newDisplayName,
		[Parameter(Mandatory=$true, Position=5)]
		[string]$culture
	)
    
	$functionName = $MyInvocation.MyCommand.Name
	Write-Debug "[$functionName] Entering function" 
	Write-Debug "[$functionName] Parameter / siteURL: $siteURL"
	Write-Debug "[$functionName] Parameter / listName: $listName"
	Write-Debug "[$functionName] Parameter / fieldStaticName: $fieldStaticName"
	Write-Debug "[$functionName] Parameter / newDisplayName: $newDisplayName"
	Write-Debug "[$functionName] Parameter / culture: $culture"

	try
  {
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
    if($curWeb -ne $null)
    {
      $list = $curWeb.Lists.TryGetList($listName)
      if($list -ne $null)
      {
				[System.Threading.Thread]::CurrentThread.CurrentUICulture=$culture;
				$curField = $list.Fields.TryGetFieldByStaticName($fieldStaticName);
				if($curField -ne $null)
				{
					Write-Host "[$functionName] About to rename '$fieldStaticName' to '$newDisplayName' on list $listName for $culture."  -ForegroundColor Magenta `r
					$curField.Title = $newDisplayName;
					$curField.Update($true);
					$list.Update();
					Write-Host "[$functionName] ListColumn '$fieldStaticName' has been renamed to '$newDisplayName' on list $listName for $culture."  -ForegroundColor Green  `r
				}
				else
				{
					Write-Warning "[$functionName] ListColumn '$fieldStaticName' does not exist on list '$listName'."
				}
      }
			else
			{
				Write-Warning "[$functionName] List '$listName' does not exist on site '$siteURL'"
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
