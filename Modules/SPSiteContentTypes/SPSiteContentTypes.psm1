##############################################################################################
#              
# NAME: SPSiteContentTyoes.psm1 
# PURPOSE: 
#	Manage Site Content-Types
#	Relies on an XML Configuration file to identify and handle the Site Content-types.
#	See SPSiteContentTypes.xml for Schema
#
# SOURCE : https://github.com/OhNotreDame/SPPS
#
##############################################################################################

<#
.SYNOPSIS
	Browse and parse the SiteContentTypesDescriptionXML XML object
	
.DESCRIPTION
	Browse and parse the file SiteColumnsDescriptionXML
	For each node, Check if site Content-Type exists in site siteURL
		If does not exist, create a new Content-Type.	
		If exists, call update the Content-Type.		
	
.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER SiteContentTypesDescriptionXML
	XML Object representing the content-types to be created/updated
	
.EXAMPLE
	browseAndParseSiteContentTypesXML -siteURL <SiteURL> -SiteContentTypesDescriptionXML <XMLObjectToParse>
	
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
function browseAndParseSiteContentTypesXML()
{
    [CmdletBinding()]
	param
    (
        [Parameter(Mandatory=$true, Position=1)]
        [string]$siteURL,
        [Parameter(Mandatory=$true, Position=2)]
        [XML]$SiteContentTypesDescriptionXML
    )
	
	$functionName = $MyInvocation.MyCommand.Name
    Write-Debug "[$functionName] Entering function" 
    Write-Debug "[$functionName] Parameter / siteURL: $siteURL"
	
    try
    {
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
        if($curWeb -ne $null)
        {
                        
            $ctTypes= $SiteContentTypesDescriptionXML.ContentTypes.ContentType
            foreach($ctType in $ctTypes)
            {
                $IsContentTypeExist = existSiteContentTypeByName -siteURL $siteURL -ctName $ctType.Name
				Write-Debug "[$functionName] Content-Type exists? $IsContentTypeExist"
                if($IsContentTypeExist -eq $false)
                {
					createSiteContentType -siteURL $siteURL -ctDescriptionXML $ctType
                }
                else
                {
					updateSiteContentType -siteURL $siteURL -ctDescriptionXML $ctType
                }
            }
		}
		else
		{
			Write-Warning "[$functionName] Site '$siteURL' could not be found." 
		    return;
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
	Check if content-type siteContentTypeName exist on site siteURL
	
.DESCRIPTION
Check if content-type siteContentTypeName exist on site siteURL	
	
.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER ctName
	Name of the content-type
	
.EXAMPLE
	existSiteContentTypeByName -siteURL <SiteURL> -ctName <ctName>
	
.OUTPUTS
	True if Content-Type exists, false otherwise

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 16.01.2017
	Last Updated by: JBO
	Last Updated: 16.01.2017
#>
function existSiteContentTypeByName()
{
	[CmdletBinding()]
	param
    (
        [Parameter(Mandatory=$true, Position=1)]
        [string]$siteURL,
        [Parameter(Mandatory=$true, Position=2)]
        [string]$ctName
    )
	
	$functionName = $MyInvocation.MyCommand.Name
    Write-Debug "[$functionName] Entering function" 
    Write-Debug "[$functionName] Parameter / siteURL: $siteURL"
    Write-Debug "[$functionName] Parameter / ctName: $ctName"
	
	$existCtype= $false
    try
    { 
        $curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if($curWeb -ne $null)
		{ 
			$currentCT = $curWeb.AvailableContentTypes[$ctName]
			if($currentCT -ne $null)
			{
				$existCtype= $true
				Write-Debug "[$functionName] Content-Type '$ctName' exists on site '$siteURL'."
			}
			else
			{
				$existCtype= $false
				Write-Debug "[$functionName] Content-Type '$ctName' does not exist on site '$siteURL'."
			}
		}
		else
		{
			Write-Warning "[$functionName] Site '$siteURL' could not be found." 
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
    return $existCtype
}


<#
.SYNOPSIS
	Check if content-type ctID exist on site siteURL
	
.DESCRIPTION
	Check if content-type ctID exist on site siteURL	
	
.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER ctID
	ID of the content-type
	
.EXAMPLE
	existSiteContentTypeByID -siteURL <SiteURL> -ctID <ctID>
	
.OUTPUTS
	True if Content-Type exists, false otherwise

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 16.01.2017
	Last Updated by: JBO
	Last Updated: 16.01.2017
#>
function existSiteContentTypeByID()
{
    [CmdletBinding()]
	param
    (
        [Parameter(Mandatory=$true, Position=1)]
        [string]$siteURL,
        [Parameter(Mandatory=$true, Position=2)]
        [string]$ctID
    )
    $functionName = $MyInvocation.MyCommand.Name
    Write-Debug "[$functionName] Entering function" 
    Write-Debug "[$functionName] Parameter / siteURL: $siteURL"
    Write-Debug "[$functionName] Parameter / ctID: $ctID"
	
	$existCtype= $false
    try
    {
		$ctContentTypeId = [Microsoft.SharePoint.SPContentTypeId] $ctID
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if($curWeb -ne $null)
		{ 
			$currentCT = $curWeb.AvailableContentTypes[$ctContentTypeId]
			if($currentCT -ne $null)
			{
				$existCtype= $true
				Write-Debug "[$functionName] Content-Type '$ctID' exists on site '$siteURL'." 
			}
			else
			{
				$existCtype= $false
				Write-Debug "[$functionName] Content-Type '$ctID' does not exist on site '$siteURL'."
			}
		}
		else
		{
			Write-Warning "[$functionName] Site '$siteURL' not found."
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
        if($contentType -ne $null)
        {
            $contentType.Dispose()
        }
        if($curWeb -ne $null)
        {
           $curWeb.Dispose()
        }
        Write-Debug "[$functionName] Exiting function" 
    }
    return $existCtype
}


<#
.SYNOPSIS
	Get (if exists on site siteURL) content-type ctName and return it
	
.DESCRIPTION
	Get (if exists on site siteURL) content-type ctName and return it
	
.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER ctName
	Name of the content-type
	
.EXAMPLE
	getSiteContentTypeByName -siteURL <SiteURL> -ctName <ctName>
	
.OUTPUTS
	if exists, the content-type $ctName
	if not, null

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 16.01.2017
	Last Updated by: JBO
	Last Updated: 16.01.2017
#>
function getSiteContentTypeByName()
{
    [CmdletBinding()]
	param
    (
        [Parameter(Mandatory=$true, Position=1)]
        [string]$siteURL,
        [Parameter(Mandatory=$true, Position=2)]
        [string]$ctName
    )
    $currentCT= $null
	
	$functionName = $MyInvocation.MyCommand.Name
    Write-Debug "[$functionName] Entering function" 
	Write-Debug "[$functionName] Parameter / siteURL: $siteURL"
    Write-Debug "[$functionName] Parameter / ctName: $ctName"
	
	try
    {

		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if($curWeb -ne $null)
		{ 
			$existCtype = existSiteContentTypeByName -siteURL $siteURL -ctName $ctName
			if($existCtype -eq $true)
        	{
				$currentCT = $curWeb.AvailableContentTypes[$ctName]
				Write-Debug "[$functionName] Content-Type '$ctID' exists on site '$siteURL'."
			}
			else
			{
				Write-Debug "[$functionName] Content-Type '$ctName' does not exist on site '$siteURL'."
			}	
		}
		else
		{
			Write-Warning "[$functionName] Site '$siteURL' not found."
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
    return $currentCT
}


<#
.SYNOPSIS
	Get (if exists on site siteURL) content-type ctID and return it
	
.DESCRIPTION
	Get (if exists on site siteURL) content-type ctID and return it
	
.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER ctID
	ID of the content-type
	
.EXAMPLE
	getSiteContentTypeByID -siteURL <SiteURL> -ctID <ctID>
	
.OUTPUTS
	if exists, the content-type $ctName
	if not, null

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 16.01.2017
	Last Updated by: JBO
	Last Updated: 16.01.2017
#>
function getSiteContentTypeByID()
{
	[CmdletBinding()]
	param
    (
        [Parameter(Mandatory=$true, Position=1)]
        [string]$siteURL,
        [Parameter(Mandatory=$true, Position=2)]
        [string]$ctID
    )
    
	$currentCT = $null;
	$functionName = $MyInvocation.MyCommand.Name
    Write-Debug "[$functionName] Entering function" 
	Write-Debug "[$functionName] Parameter / siteURL: $siteURL"
    Write-Debug "[$functionName] Parameter / ctID: $ctID"
	
	try
	{
		$ctContentTypeId = [Microsoft.SharePoint.SPContentTypeId] $ctID
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if($curWeb -ne $null)
		{ 
			$existCT = existSiteContentTypeByID -siteURL $siteURL -ctID  $ctContentTypeId
			if($existCT -eq $true)
			{
				$currentCT = $curWeb.AvailableContentTypes[$ctContentTypeId]
				Write-Debug "[$functionName] Content-Type '$ctID' exists on site '$siteURL'." 
			}
			else
			{
				Write-Debug "[$functionName] Content-Type '$ctID' does not exist on site '$siteURL'."
			}	
		}
		else
		{
			Write-Warning "[$functionName] Site '$siteURL' not found."
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
    return $currentCT
}


<#
.SYNOPSIS
	Create a content-type based on its XML Definition
	
.DESCRIPTION
	Create a content-type based on its XML Definition
	
.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER ctDescriptionXML
	XML Definition of Content-Type
	
.EXAMPLE
	createSiteContentType -siteURL <SiteURL> -ctDescriptionXML <ctToCreateXML>
	
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
function createSiteContentType()
{
    [CmdletBinding()]
	param
    (
        [Parameter(Mandatory=$true, Position=1)]
        [string]$siteURL,
        [Parameter(Mandatory=$true, Position=2)]
        [XML.XMLElement]$ctDescriptionXML
    )
    
	$functionName = $MyInvocation.MyCommand.Name
    Write-Debug "[$functionName] Entering function" 
	Write-Debug "[$functionName] Parameter / siteURL: $siteURL"

    try
    {
        $curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if($curWeb -ne $null)
        {            

			$ctName = $ctDescriptionXML.Attributes["Name"].value
            
			if($curWeb.AvailableContentTypes[$ctName] -eq $null)
			{				
				$parentCTName = $ctDescriptionXML.Attributes["ParentContentType"].value;
				Write-Host "[$functionName] Initial parentCTName : $parentCTName"

				$culture = $curWeb.UICulture.Name
				$culture = $culture.ToLower()

				switch -Wildcard ($culture)
				{
					'en*' 
					{
						$item_cultOk = "Item"
						$document_cultOk = "Document" 		
						$documentSet_cultOk = "Document Set" 		
					}
					'fr*' 
					{
						$item_cultOk = "Élément"
						$document_cultOk = "Document" 		
						$documentSet_cultOk = "Ensemble de documents"						
					}
					'de*' 
					{
						$item_cultOk = "Element"
						$document_cultOk = "Dokument" 	
						$documentSet_cultOk = "Dokumentenmappe"						
					}
					'it*' 
					{
						$item_cultOk = "Elemento"
						$document_cultOk = "Documento" 	
						$documentSet_cultOk = "Set di documenti"
					}
					default
					{
						$item_cultOk = "Élément"
						$document_cultOk = "Document" 	
						$documentSet_cultOk = "Ensemble de documents"			
					}
				}

				[string] $parentCT_cultOK = ""
				switch ($parentCTName)  
				{
					"Item" { $parentCT_cultOK = $item_cultOk; }
					"Document" { $parentCT_cultOK = $document_cultOk; }
					"Document Set" { $parentCT_cultOK = $documentSet_cultOk; }
				}
				Write-Host "[$functionName] Parent Content-Type ($parentCTName) resolved : $parentCT_cultOK"
				$ctypeParent =$curWeb.AvailableContentTypes[$parentCT_cultOK];
				if($ctypeParent -ne $null)
				{
					Write-Host "[$functionName] About to create Content-Type '$ctName' on site '$siteURL'." -ForegroundColor Magenta `r 
					$ctypeGroup = $ctDescriptionXML.Attributes["Group"].value
					$ctypeDescription = $ctDescriptionXML.Attributes["Description"].value
					
					$ctype =  New-Object Microsoft.SharePoint.SPContentType -ArgumentList @($ctypeParent,$curWeb.ContentTypes,$ctName)
					$ctype.Group= $ctypeGroup
					$ctype.Description= $ctypeDescription
					$curWeb.ContentTypes.add($ctype) | Out-Null
					$curWeb.Update()
					Write-Host "[$functionName] Content-Type '$ctName' has been successfully created on site '$siteURL'." -ForegroundColor Green `r 
	
					if($ctDescriptionXML.FieldRefs.HasChildNodes -eq $true)
					{
						foreach($addFieldRef in $ctDescriptionXML.FieldRefs.FieldRef)
						{
							addContentTypeFieldFromXML -siteURL $siteURL -ct $ctype -fieldToAddXML $addFieldRef
						}
				
						foreach($removeFieldRef in $ctDescriptionXML.FieldRefs.RemoveFieldRef)
						{
							removeContentTypeFieldFromXML -siteURL $siteURL -ct $ctype -fieldToRemoveXML $removeFieldRef
						}
					}
					else
					{
						Write-Warning "[$functionName] XML Content-Type FieldsRefs node does not have child nodes. No FieldRef to add/remove."
					}
				}
				else
				{
					Write-Warning "[$functionName] Parent Content-Type '$parentCTName' does not exist."
				}
			}
			else
			{
				Write-Warning "[$functionName] Content-Type '$ctName' already exists."
			}
		}
		else
		{
			Write-Warning "[$functionName] Site '$siteURL' not found."
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
	Update an existing content-type based on the XML Definition
	
.DESCRIPTION
	Update an existing content-type based on the XML Definition
	
.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER ctDescriptionXML
	XML Definition of Content-Type
	
.EXAMPLE
	updateSiteContentType -siteURL <SiteURL> -ctDescriptionXML <ctToUpdateXML>
	
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
function updateSiteContentType()
{
    [CmdletBinding()]
	param
    (
        [Parameter(Mandatory=$true, Position=1)]
        [string]$siteURL,
        [Parameter(Mandatory=$true, Position=2)]
        [XML.XMLElement]$ctDescriptionXML
    )
   
	$functionName = $MyInvocation.MyCommand.Name
    Write-Debug "[$functionName] Entering function" 
	Write-Debug "[$functionName] Parameter / siteURL: $siteURL"

    try
    {
		$ctName= $ctDescriptionXML.Attributes["Name"].value
		$ctGroup= $ctDescriptionXML.Attributes["Group"].value
		$ctDescription= $ctDescriptionXML.Attributes["Description"].value
		$ctParent= $ctDescriptionXML.Attributes["Parent"].value

		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if($curWeb -ne $null)
        {   
			$ctToUpdate = $curWeb.ContentTypes[$ctName]
			if ($ctToUpdate -ne $null)
			{
				$ctToUpdate.Group = $ctGroup
				$ctToUpdate.Description = $ctDescription
				$ctToUpdate.UpdateIncludingSealedAndReadOnly($true);
				
				#$ctToUpdate.Update() <=> Updates the content type definition that is stored in the database with any changes you have made programmatically.
				#$ctToUpdate.Update($true) <=>  Updates the content type definition that is stored in the database and, optionally, updates all content types that inherit from this content type.
				#$ctToUpdate.UpdateIncludingSealedAndReadOnly ($true) <=> Updates the content type definition that is stored in the database and, applies changes to all derived content types, including sealed or read-only CTs.
				

				if($ctDescriptionXML.FieldRefs.HasChildNodes -eq $true)
				{
					Write-Host "[$functionName] About to update Content-Type '$ctName' on site '$siteURL'." -ForegroundColor Magenta `r
					foreach($addFieldRef in $ctDescriptionXML.FieldRefs.FieldRef)
					{
					   addContentTypeFieldFromXML -siteURL $siteURL -ct $ctToUpdate -fieldToAddXML $addFieldRef
					}
				
					foreach($removeFieldRef in $ctDescriptionXML.FieldRefs.RemoveFieldRef)
					{
						removeContentTypeFieldFromXML -siteURL $siteURL -ct $ctToUpdate -fieldToRemoveXML $removeFieldRef
					}
					$ctToUpdate.UpdateIncludingSealedAndReadOnly($true);

					Write-Host "[$functionName] Content-Type '$ctName' has been updated on site '$siteURL'." -ForegroundColor Green `r
				}
				else
				{
					Write-Warning "[$functionName] XML Content-Type FieldsRefs node does not have child nodes. No FieldRef to add/remove."
				}
			}
			else
			{
				Write-Warning "[$functionName] Content-Type '$ctName' does not exist."
			}
		}
		else
		{
			Write-Warning "[$functionName] Site '$siteURL' not found."
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
	Add field to site Content-Type $ct based on field $fieldToAddXML XML definition
	
.DESCRIPTION
	Add field to site Content-Type $ct based on field $fieldToAddXML XML definition
	
.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER ct
	Content-Type
	
.PARAMETER fieldToAddXML
	XML Field Description
	
.EXAMPLE
	addFieldToContentTypeFromXML -siteURL <SiteURL> -ct <ctToUpdate> -fieldToAddXML <fieldDescription>
	
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
function addContentTypeFieldFromXML()
{
	[CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true, Position=1)]
        [string]$siteURL,
        [Parameter(Mandatory=$true, Position=2)]
        [Microsoft.SharePoint.SPContentType]$ct,
        [Parameter(Mandatory=$true, Position=3)]
        [Xml.XmlElement]$fieldToAddXML 
    )
    
	$functionName = $MyInvocation.MyCommand.Name
    Write-Debug "[$functionName] Entering function" 
	Write-Debug "[$functionName] Parameter / siteURL: $siteURL"
	Write-Debug "[$functionName] Parameter / ct: $($ct.Name)"
	
	try
    {

       if($fieldToAddXML -ne $null)
       {
            $ctFieldId= $fieldToAddXML.ID
            $ctFieldName=$fieldToAddXML.Name
            $ctFieldDspName= $fieldToAddXML.DisplayName
			$ctFieldStaticName = $fieldToAddXML.StaticName
			$ctFieldRequired = [System.Convert]::ToBoolean($fieldToAddXML.Attributes["Required"].value)
			$ctFieldHidden = [System.Convert]::ToBoolean($fieldToAddXML.Attributes["Hidden"].value)
			
			
			if(([string]::IsNullOrEmpty($ctFieldStaticName))  -and (!([string]::IsNullOrEmpty($ctFieldName))))
			{ 
				$ctFieldStaticName =$ctFieldName
			}
	

            $curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
			if($curWeb -ne $null)
			{            
				$parentSiteURL = $curWeb.site.RootWeb.URL
				$existField = existSiteColumn -siteURL $siteURL -siteColumnID $fieldToAddXML.ID
				if($existField -eq $True)
				{
					Write-Host "[$functionName] Site Column '$ctFieldDspName' found on site '$siteUrl'."  -ForegroundColor Green `
					#Check the field in Content-Type.
					$fieldInCT = $ct.Fields.TryGetFieldByStaticName($ctFieldStaticName)
					$ctName =$ct.Name
					if($fieldInCT -eq $null)
					{

						Write-Host "[$functionName] About to add Site column '$ctFieldName' to Content-Type '$ctName'."  -ForegroundColor Magenta `r
						
						#Add field to Content-Type
						$fieldToAdd = $curWeb.Fields[[GUID]$fieldToAddXML.ID]
						
						$SPFieldLinkToAdd = new-object Microsoft.SharePoint.SPFieldLink $fieldToAdd
						$SPFieldLinkToAdd.Required = $ctFieldRequired 
						$SPFieldLinkToAdd.Hidden = $ctFieldHidden 
						$SPFieldLinkToAdd.DisplayName = $ctFieldDspName 	
						
						$ct.FieldLinks.Add($SPFieldLinkToAdd)
						$ct.Update($true)
						$curWeb.Update()
						Write-Host "[$functionName] Site column '$ctFieldName' has been successfully added to Content-Type '$ctName'."  -ForegroundColor Green `r

					}
					else 
					{
						#Make sure that Required, Hidden and Display Name are properly setup
						Write-Debug "[$functionName] Site Column '$ctFieldDspName' is already present in content-type '$ctName'." 
						Write-Host "[$functionName] About to change Required, Hidden, DisplayName settings on Site Column '$ctFieldDspName' of Content-Type '$ctName'." -ForegroundColor Magenta `r
						
						Write-Debug "[$functionName] Changing Required setting for '$ctFieldDspName'." 
						$ct.FieldLinks[$fieldInCT.Id].Required = $ctFieldRequired 
					
						Write-Debug "[$functionName] Changing Hidden setting for '$ctFieldDspName'." 
						$ct.FieldLinks[$fieldInCT.Id].Hidden = $ctFieldHidden 
				
						Write-Debug "[$functionName] Changing Title value for '$ctFieldDspName'."
						$ct.FieldLinks[$fieldInCT.Id].DisplayName = $ctFieldDspName

						$ct.Update($true)
						$curWeb.Update()
						Write-Host "[$functionName] Required, Hidden, DisplayName settings successfully updated for Site column '$ctFieldName' on Content-Type '$ctName'."  -ForegroundColor Green `r
					}
				}
				else
				{
					Write-Warning "[$functionName] Site Column '$ctFieldDspName' does not exist on site '$siteUrl'." 
					
					
					$existField = existSiteColumn -siteURL $parentSiteURL -siteColumnID $fieldToAddXML.ID
					if($existField -eq $True)
					{
						Write-Host "[$functionName] Site Column '$ctFieldDspName' found on site '$parentSiteURL'."  -ForegroundColor Green `
						
						#Check the field in Content-Type.
						$fieldInCT = $ct.Fields.TryGetFieldByStaticName($ctFieldStaticName)
						$ctName =$ct.Name
						if($fieldInCT -eq $null)
						{

							Write-Host "[$functionName] About to add Site column '$ctFieldName' to Content-Type '$ctName'."  -ForegroundColor Magenta `r
							
							#Add field to Content-Type
							$fieldToAdd = $curWeb.Site.RootWeb.Fields[[GUID]$fieldToAddXML.ID]
							
							$SPFieldLinkToAdd = new-object Microsoft.SharePoint.SPFieldLink $fieldToAdd
							$SPFieldLinkToAdd.Required = $ctFieldRequired 
							$SPFieldLinkToAdd.Hidden = $ctFieldHidden 
							$SPFieldLinkToAdd.DisplayName = $ctFieldDspName 	
							
							$ct.FieldLinks.Add($SPFieldLinkToAdd)
							$ct.Update($true)
							$curWeb.Update()
							Write-Host "[$functionName] Site column '$ctFieldName' has been successfully added to Content-Type '$ctName'."  -ForegroundColor Green `r

						}
						else
						{
							#Make sure that Required, Hidden and Display Name are properly setup
							Write-Debug "[$functionName] Site Column '$ctFieldDspName' is already present in content-type '$ctName'." 
							Write-Host "[$functionName] About to change Required, Hidden, DisplayName settings on Site Column '$ctFieldDspName' of Content-Type '$ctName'." -ForegroundColor Magenta `r
							
							Write-Debug "[$functionName] Changing Required setting for '$ctFieldDspName'." 
							$ct.FieldLinks[$fieldInCT.Id].Required = $ctFieldRequired 
						
							Write-Debug "[$functionName] Changing Hidden setting for '$ctFieldDspName'." 
							$ct.FieldLinks[$fieldInCT.Id].Hidden = $ctFieldHidden 
					
							Write-Debug "[$functionName] Changing Title value for '$ctFieldDspName'."
							$ct.FieldLinks[$fieldInCT.Id].DisplayName = $ctFieldDspName

							$ct.Update($true)
							$curWeb.Update()
							Write-Host "[$functionName] Required, Hidden, DisplayName settings successfully updated for Site column '$ctFieldName' on Content-Type '$ctName'."  -ForegroundColor Green `r
						}
					}
					else
					{
						Write-Warning "[$functionName] Site Column '$ctFieldDspName' does not exist on site '$parentSiteURL'." 
					}
				}
            }
			else
			{
				Write-Warning "[$functionName] Site '$siteURL' not found." 
			}
		}
		else
		{
			Write-Warning "[$functionName] XML description is null. Impossible to add field to content-type."
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
	Add field $fieldStaticName to site Content-Type
	
.DESCRIPTION
	Add field $fieldStaticName to site Content-Type
	
.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER ct
	Content-Type
	
.PARAMETER fieldStaticName
	XML Field Description
	
.EXAMPLE
	addContentTypeFieldWithStaticName -siteURL <SiteURL> -ct <ctToUpdate> -fieldToAddXML <fieldStaticName>
	
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
function addContentTypeFieldWithStaticName()
{
    [CmdletBinding()]
	param
    (
        [Parameter(Mandatory=$true, Position=1)]
        [string]$siteURL,
        [Parameter(Mandatory=$true, Position=2)]
        [Microsoft.SharePoint.SPContentType]$ct,
        [Parameter(Mandatory=$true, Position=3)]
        [String]$fieldStaticName
    )

	$functionName = $MyInvocation.MyCommand.Name
    Write-Debug "[$functionName] Entering function" 
	Write-Debug "[$functionName] Parameter / siteURL: $siteURL"
	Write-Debug "[$functionName] Parameter / ct: $($ct.Name)"
	Write-Debug "[$functionName] Parameter / fieldStaticName: $fieldStaticName"
	
	try
    {
	
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if($curWeb -ne $null)
		{		
			$ctName = $ct.Name
			$fieldToAdd = $curWeb.Fields.TryGetFieldByStaticName($fieldStaticName)
			if($fieldToAdd -ne $null)
			{
				#Check the field in Content-Type.
				$fieldInCT = $ct.Fields.TryGetFieldByStaticName($fieldStaticName)
				if($fieldInCT -eq $null)
				{
					Write-Host "[$functionName] About to add Site Column '$fieldStaticName' to Content-Type '$ctName'." -ForegroundColor Magenta `r
					$SPFieldLinkToAdd = new-object Microsoft.SharePoint.SPFieldLink $fieldToAdd
					$ct.FieldLinks.Add($SPFieldLinkToAdd)
					$ct.Update($true)
					$curWeb.Update()
					
					Write-Host "[$functionName] Site Column '$fieldStaticName' has been successfully added to Content-Type '$ctName'."
				}
				else
				{
					Write-Debug "[$functionName] Site Column '$fieldStaticName' is already added to Content-Type '$ctName'."
				}
			}
			else
			{
				Write-Debug "[$functionName] Site Column '$fieldStaticName' does not exist on site '$siteUrl'."
			}
		}
		else
		{
			Write-Debug "[$functionName] Site '$siteURL' not found."
		}
    }
    catch [Exception]
    {
        Write-Host "$functionNam Error: $_.Exception.Message" -ForegroundColor Red `r
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
	Remove field from Content-Type $ct based on its XML definition $fieldToRemoveXML
	
.DESCRIPTION
	Remove field from Content-Type $ct based on its XML definition $fieldToRemoveXML
	
.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER ct
	Content-Type
	
.PARAMETER fieldToRemoveXML
	XML description of the field to be removed
	
.EXAMPLE
	removeContentTypeFieldFromXML -siteURL <SiteURL> -ct <ctToUpdate> -fieldToRemoveXML <xmlFieldDescription>
	
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
function removeContentTypeFieldFromXML()
{
    [CmdletBinding()]
	param
    (
        [Parameter(Mandatory=$true, Position=1)]
        [string]$siteURL,
        [Parameter(Mandatory=$true, Position=2)]
        [Microsoft.SharePoint.SPContentType]$ct,
        [Parameter(Mandatory=$true, Position=3)]
        [Xml.XmlElement]$fieldToRemoveXML
    )
	
	$functionName = $MyInvocation.MyCommand.Name
    Write-Debug "[$functionName] Entering function" 
	Write-Debug "[$functionName] Parameter / siteURL: $siteURL"
	Write-Debug "[$functionName] Parameter / ct: $($ct.Name)"
	Write-Debug "[$functionName] Parameter / fieldToRemoveXML: $($fieldToRemoveXML.StaticName)"
		
    try
    {
        $ctFieldId= $fieldToRemoveXML.Attributes["Id"].value
        $ctFieldName=$fieldToRemoveXML.Attributes["Name"].value
		$ctFieldStaticName = $fieldToRemoveXML.StaticName
		
		if(([string]::IsNullOrEmpty($ctFieldStaticName))  -and (!([string]::IsNullOrEmpty($ctFieldName))))
		{ 
			$ctFieldStaticName =$ctFieldName
		}
		
		
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if($curWeb -ne $null)
		{		
			$existField = existSiteColumn -siteURL $siteURL -siteColumnID $fieldToRemoveXML.ID
			if($existField -eq $True)
			{
				$ctName = $ct.Name
				#Check the field in Content-Type.
				$fieldInCT = $ct.Fields.TryGetFieldByStaticName($ctFieldStaticName)
				if($fieldInCT -ne $null)
				{
					Write-Host "[$functionName] About to remove Site Column '$ctFieldStaticName' from Content-Type '$ctName'." -ForegroundColor Magenta `r
					#Write-Host "[$functionName] fieldInCT is not null"  -ForegroundColor Magenta `r
					$fieldToRemove = $curWeb.Fields[[GUID]$fieldToRemoveXML.ID]
					$spFieldLink = New-Object Microsoft.SharePoint.SPFieldLink ($fieldToRemove);
					#Write-Host "[$functionName] SPFieldLinkToAdd:"  $spFieldLink.ID -ForegroundColor Magenta `r
					$ct.FieldLinks.Delete($spFieldLink.ID)
					$ct.Update()
					$curWeb.Update()
					Write-Host "[$functionName] Site Column '$ctFieldName' has been removed successfully from Content-Type '$ctName'."  -ForegroundColor Green `r
				}
				else
				{
					Write-Warning "[$functionName] Site Column '$ctFieldDspName' is already removed from Content-Type '$ctName'."
				}
			}
			else
			{
				Write-Warning "[$functionName] Site Column '$ctFieldDspName' does not exist on site '$siteUrl'."
			}
		}
		else
		{
			Write-Warning "[$functionName] Site '$siteURL' not found." 
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
	Remove field $fieldStaticName from Content-Type $ct 
	
.DESCRIPTION
	Remove field $fieldStaticName from Content-Type $ct 
	
.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER ct
	Content-Type
	
.PARAMETER fieldStaticName
	Static Name of the field to be removed
	
.EXAMPLE
	removeContentTypeFieldWithStaticName -siteURL <SiteURL> -ct <ctToUpdate> -fieldStaticName <fieldStaticName>
	
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
function removeContentTypeFieldWithStaticName()
{
	[CmdletBinding()]
	param
    (
        [Parameter(Mandatory=$true, Position=1)]
        [string]$siteURL,
        [Parameter(Mandatory=$true, Position=2)]
        [Microsoft.SharePoint.SPContentType]$ct,
        [Parameter(Mandatory=$true, Position=3)]
        [String]$fieldStaticName
    )
   
	$functionName = $MyInvocation.MyCommand.Name
    Write-Debug "[$functionName] Entering function" 
	Write-Debug "[$functionName] Parameter / siteURL: $siteURL"
	Write-Debug "[$functionName] Parameter / ct: $($ct.Name)"
	Write-Debug "[$functionName] Parameter / fieldStaticName: $fieldStaticName"
	
	try
    {
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if($curWeb -ne $null)
		{		
		
			$ctName = $ct.Name
			$fieldToRemove = $curWeb.Fields.TryGetFieldByStaticName($fieldStaticName)
			if($fieldToRemove -ne $null)
			{
				#Check the field in Content-Type.
				$fieldInCT = $ct.Fields.TryGetFieldByStaticName($fieldStaticName)
				if($fieldInCT -ne $null)
				{
					Write-Host "[$functionName] About to remove Site Column '$fieldStaticName' from Content-Type '$ctName'." -ForegroundColor Magenta `r
					$spFieldLink = New-Object Microsoft.SharePoint.SPFieldLink ($fieldToRemove);
					#Write-Host "[removeSiteColumnFromContentType] spFieldLink:"  $spFieldLink.ID -ForegroundColor Magenta `r
					$ct.FieldLinks.Delete($spFieldLink.ID)
					$ct.Update()
					$curWeb.Update()
					Write-Host "[$functionName] Site Column '$fieldStaticName' has been removed successfully from Content-Type '$ctName'." -ForegroundColor Green `r
				}
				else
				{
					Write-Warning "[$functionName] Site Column '$fieldStaticName' is already removed from Content-Type '$ctName'."
				}
			}
			else
			{
				Write-Warning "[$functionName] Site Column '$fieldStaticName' does not exist on site '$siteUrl'."
			}
		}
		else
		{
			Write-Warning "[$functionName] Site '$siteURL' not found."
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
	Edit field from Content-Type $ct based on its XML definition $fieldToEditXML
	
.DESCRIPTION
	Edit field from Content-Type $ct based on its XML definition $fieldToEditXML
	
.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER ct
	Content-Type
	
.PARAMETER fieldToEditXML
	XML Description on the field to update
	
.EXAMPLE
	editContentTypeFieldFromXML -siteURL <SiteURL> -ct <ctToUpdate> -fieldToEditXML <fieldToEditXML>
	
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
function editContentTypeFieldFromXML()
{
	[CmdletBinding()]
	param
    (
        [Parameter(Mandatory=$true, Position=1)]
        [string]$siteURL,
        [Parameter(Mandatory=$true, Position=2)]
        [ContentType]$ct,
        [Parameter(Mandatory=$true, Position=3)]
        [XML]$fieldToEditXML
    )
	
	$functionName = $MyInvocation.MyCommand.Name
    Write-Debug "[$functionName] Entering function" 
	Write-Debug "[$functionName] Parameter / siteURL: $siteURL"
	Write-Debug "[$functionName] Parameter / ct: $($ct.Name)"
	Write-Debug "[$functionName] Parameter / fieldToEditXML: $($fieldToEditXML.Name)"
   
   try
    {
		#Parsing XML Element
        $ctFieldId= $fieldToEditXML.Id
        $ctFieldName=$fieldToEditXML.Name
        $ctFieldDspName= $fieldToEditXML.DisplayName
		$ctFieldStaticName = $fieldToEditXML.StaticName
        $ctFieldRequired= $fieldToEditXML.Attributes["Required"].value
        $ctFieldReadOnly= $fieldToEditXML.Attributes["ReadOnly"].value
        $field = $ct.Fields.TryGetFieldByStaticName($ctFieldStaticName)
        if($field -ne $null)
        {
			Write-Host "[$functionName] About to update Site Column '$ctFieldDspName' on Content-Type '$($ct.Name)'." -ForegroundColor Magenta `r
			$field.SchemaXml = $fieldToEditXML
		    $field.Update()
			$ct.FieldLinks.Update($field)
			$ct.Update($true)
			Write-Host "[$functionName] Site Column '$ctFieldDspName' has been successfully updated on Content-Type '$($ct.Name)'." -ForegroundColor Green `r
        }
		else
		{
			Write-Host "[$functionName] Site Column '$ctFieldDspName' does not exist on Content-Type '$($ct.Name)'" -ForegroundColor Cyan `r
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
	Delete site Content-Type based on its name $ctName
	
.DESCRIPTION
	Delete site Content-Type based on its name $ctName
	(NOT FULLY IMPLEMENTED)
	
.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER ctName
	Name of the Content-Type to be deleted

.EXAMPLE
	deleteContentType -siteURL <SiteURL> -ctName <ctName>
	
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
function deleteContentType()
{
	[CmdletBinding()]
	param
    (
        [Parameter(Mandatory=$true, Position=1)]
        [string]$siteURL,
        [Parameter(Mandatory=$true, Position=2)]
        [string]$ctName
    )
	
	$functionName = $MyInvocation.MyCommand.Name
    Write-Debug "[$functionName] Entering function" 
	Write-Debug "[$functionName] Parameter / siteURL: $siteURL"
	Write-Debug "[$functionName] Parameter / ct: $ctName"

    try
    {
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if($curWeb -ne $null)
		{		
		
			Write-Host "#------------------------------------------------#" -ForegroundColor Red `r
			Write-Host "#- Function '$functionName' is not implemented. -#" -ForegroundColor Red `r
			Write-Host "#------------------------------------------------#" -ForegroundColor Red `r

			#foreach($loopWeb in $curSPSite.AllWebs)
			#{
				#-------------------------------------------------------------------------------
				#Step 1: Delete Content-Type on each and every lists of each and every subsites
				#-------------------------------------------------------------------------------
				#$loopWebTitle= $loopWeb.Title
				#$loopWeb.Url
				#$ctToRemove = getSiteContentTypeByName	-siteURL $siteURL -ctName $ctName
				#if ($ctToRemove -ne $null)
				#{
				#	Write-Host "[$functionName] About to delete Site Content-Type '$ctName' from SPWeb '$siteURL' ... " -ForegroundColor White `r
				#	$ctusage = [Microsoft.SharePoint.SPContentTypeUsage]::GetUsages($ctToRemove)
				#	foreach ($ctuse in $ctusage)
				#	{
				#		$ctuse
				#		if ($ctuse.IsUrlToList)
				#		{
				#			$list = $loopWeb.Lists.TryGetList($ctuse.Url);
				#			$listName = $list.Title 
				#			Write-Host "[$functionName] About to delete Site Content-Type '$ctName' from List '$listName'."-ForegroundColor Yellow `r
				#			$list.ContentTypes.Delete($ctToRemove.ID);
				#			$list.Update();
				#			Write-Host "[$functionName] Site Content-Type '$ctName' successfully deleted from List '$listName'."-ForegroundColor Green `r
				#		}
				#	}
				#}
				#else
				#{
				#	Write-Host "[$functionName] Site Content-Type '$ctName' not found on SPWeb '$siteURL'."-ForegroundColor Cyan `r
				#}
			#}

			#-------------------------------------------------------------------------------
			#Step 2: Delete Site Content-Type
			#-------------------------------------------------------------------------------
			#$ct = getSiteContentTypeByName	-siteURL $curSPSite.Url  -ctName $ctName
			#if ($ct -ne $null)
			#{
			#	# Check for read-only and sealed.
			#	if ($ct.ReadOnly)
			#	{ 
			#		$ct.ReadOnly = $false
			#	}

			#	# Check for sealed.
			#	if ($ct.Sealed)
			#	{ 
			#		$ct.Sealed = $false
			#	}
				
			#	$ct.Update()
			#	$ct.Delete()
			#}
			#else
			#{
			#	Write-Host "[$functionName] Site Content-Type '$ctName' not found on '$curSPSite.Title'."-ForegroundColor Cyan `r
			#}

		}
		else
		{
			Write-Warning "[$functionName] Site '$siteURL' not found."
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