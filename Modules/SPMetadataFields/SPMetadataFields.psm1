##############################################################################################
#              
# NAME: SPMetadataFields.psm1 
# PURPOSE: 
#	Manage Metadata fields
#	Relies on an XML Configuration file to identify and handle the MMS fields.
#	See SPMetadataFields.xml for Schema
#
# SOURCE : https://github.com/OhNotreDame/SPPS
#
##############################################################################################

<#
.SYNOPSIS
	Browse and parse the MMSFieldsDescriptionXML XML object
	
.DESCRIPTION
	Browse and parse the file MMSFieldsDescriptionXML
		For each node, will initiate SPField (MMS) creation	
	
.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER MMSFieldsDescriptionXML
	XML Object representing ALL the MMS fields to be created/updated
	
.EXAMPLE
	browseAndParseMMSFieldsXML -siteURL <SiteURL> -MMSFieldsDescriptionXML <XMLObjectToParse>
	
.OUTPUTS
	None
.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 06.03.2016
	Last Updated by: JBO
	Last Updated: 06.03.2016
#>
function browseAndParseMMSFieldsXML()
{
    Param
	([Parameter(Mandatory=$true, Position=1)]
	 [string]$siteURL,
     [Parameter(Mandatory=$true, Position=2)]
	 [XML]$MMSFieldsDescriptionXML
    )
    
	$functionName = $MyInvocation.MyCommand.Name
	Write-Debug "[$functionName] Entering function" 
	Write-Debug "[$functionName] Parameter / siteURL: $siteURL"
	Write-Debug "[$functionName] Parameter / MMSFieldsDescriptionXML: $MMSFieldsDescriptionXML"

    try
    {
       $curWeb = GetSPWebBySiteUrl -siteUrl $siteURL       
       if($curWeb -ne $null)
       {
           if($MMSFieldsDescriptionXML -ne $null -and $MMSFieldsDescriptionXML.HasChildNodes -eq $true)
           {
                $fieldsToMap = $MMSFieldsDescriptionXML.MMSColumnsMapping.FieldToMap
            
                foreach($FieldToMap in $fieldsToMap)
                {
                    #$existField = $curWeb.Fields[$FieldToMap.ID]
                    $existField = $curWeb.Fields | Where {$_.Id -eq $FieldToMap.ID }

                    if($existField -ne $null)
                    {
                        if($existField.TermSetId -eq $null -or $existField.TermSetId -eq "00000000-0000-0000-0000-000000000000")
                        {
                	        bindMMSFieldToTermSet -siteURL $siteURL -FieldToMapXML $FieldToMap
                        }
                        else
                        {
                            Write-Warning "[$functionName] Field '$($existField.Name)' already bound with MMS." 
                        }
                    }
                    else
                    {
                        Write-Warning "[$functionName] Field '$($FieldToMap.ID)' does not exist."
                    }

                }
           }
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
	Bind the MMS site column to the Term Set
	
.DESCRIPTION
	Bind the MMS site column to the Term Set
	
.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER FieldToMapXML
	XML Object representing the SPField column
	
.EXAMPLE
	bindMMSFieldToTermSet -siteURL <SiteURL> -FieldToMapXML <XMLObjectToParse>
	
.OUTPUTS
	None
.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 06.03.2016
	Last Updated by: JBO
	Last Updated: 06.03.2016
#>
function bindMMSFieldToTermSet()
{
    Param
	([Parameter(Mandatory=$true, Position=1)]
	 [string]$siteURL,
     [Parameter(Mandatory=$true, Position=2)]
	 [XML.XMLElement]$FieldToMapXML
    )
   
	$functionName = $MyInvocation.MyCommand.Name
	Write-Debug "[$functionName] Entering function" 
	Write-Debug "[$functionName] Parameter / siteURL: $siteURL"
	Write-Debug "[$functionName] Parameter / FieldToMapXML: $FieldToMapXML"
	
    try
    {
        $curWeb = GetSPWebBySiteUrl -siteUrl $siteURL
        
        if($curWeb -ne $null)
        {
            $termStoreName = $FieldToMapXML.Attributes["TermStoreName"].value.Trim()
            $termSetName = $FieldToMapXML.Attributes["TermSetName"].value.Trim()
            $fieldId = $FieldToMapXML.Attributes["ID"].value.Trim()
            $groupName = $FieldToMapXML.Attributes["TermGroupName"].value.Trim()

            if($termStoreName -ne $null -and $termSetName -ne $null -and $fieldId -ne $null -and $termGroupName -ne $null)
            { 
                    $taxonomySession = Get-SPTaxonomySession -Site $siteURL
                    $termStore = $taxonomySession.TermStores | Where {$_.Name -eq $termStoreName}
                    if($termStore -ne $null)
                    {
                        $currentGroup = $termStore.Groups | Where {$_.Name -eq $termGroupName}
                        if($currentGroup -ne $null)
                        {
                            $termSet = $currentGroup.TermSets | Where {$_.Name -eq $termSetName}
                            if($termSet -ne $null)
                            {
                                $taxonomyField = $curWeb.Fields | Where-object{$_.Id -eq $fieldId}
                                $taxonomyField.sspId = $termSet.TermStore.Id
                                $taxonomyField.TermSetId = $termSet.Id
                                $taxonomyField.Update()
                                $curWeb.Update()
                                Write-Host "$functionName Field $($taxonomyField.Name) has be successfully bound to termset '$($termSet.Name)'." -ForegroundColor Green
                            }
                            else
                            {
                              Write-Warning "[$functionName] Term Set '$termSetName' does not exist." 
                            }
                        }
                        else
                        {
                           Write-Warning "[$functionName] Term Set Group '$termGroupName' does not exist."
                        }
                      
                    }                     
                    else
                    {
                        Write-Warning "[$functionName] Term Store '$termStore' does not exist."
                    }
            }
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
