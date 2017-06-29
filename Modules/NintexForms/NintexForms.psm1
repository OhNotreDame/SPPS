##############################################################################################
#              
# NAME: NintexForms.psm1 
# PURPOSE: 
#	Manage Nintex Forms (Export, Deploy, Publish...)
# 
# SOURCE : https://github.com/OhNotreDame/SPPS
#
##############################################################################################



<#
.SYNOPSIS
	Export All Active Forms from Site $siteURL and generate and XML file with all Forms settings
	
.DESCRIPTION
	Export All Active Forms fom Site $siteURL and generate and XML file with all Forms settings
	By Active Forms, we mean Published and Fully Associated Forms.
	Previews are not exported
	
.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER destFolder
	(Optional) Location of the exported form
	If not specified, will get the current location 
		
.EXAMPLE
	ExportAllFormsFromSite -siteURL <SiteURL> [-destFolder <dstFolderPath>]
	
.OUTPUTS
	One XML file containing the list of all the Forms on site $siteURL and their settings (Name, List (ID+Name), Content-Type (ID+Name) and FileName)
	One folder NintexForms\ to store all forms files
	One .XML file by Workflow

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 22.02.2017
	Last Updated by: JBO
	Last Updated: 22.02.2017
#>
function ExportAllActiveFormsFromSite()
{
	[CmdletBinding()]
	param
	(
	    [Parameter(Mandatory=$true, Position=1)]
	    [string]$siteURL, 
		[Parameter(Mandatory=$false, Position=2)]
	    [string]$destFolder
	)
		
	$functionName = $MyInvocation.MyCommand.Name
	Write-Debug "[$functionName] Entering function" 
	Write-Debug "[$functionName] Parameter / siteURL: $siteURL" 

	try
	{
		if([string]::IsNullOrEmpty($destFolder)) 
		{
			$destFolder = Get-Location
			Write-Warning "[$functionName] Paramater destFolder is empty, will set it to the current location." 
			Write-Host "[$functionName] destFolder: $destFolder" -foregroundcolor Cyan
		}
		
		Write-Debug "[$functionName] Before creating Export folder"
		$exportFolder = $destFolder + "\NintexForms\"
		if (!(Test-Path $exportFolder))
		{
			New-Item $exportFolder -type Directory -Force | Out-Null
		}
		
		$artefactFolder = $destFolder + "\Artefacts\"
		if (!(Test-Path $artefactFolder))
		{
			New-Item $artefactFolder -type Directory -Force | Out-Null
		}
		
		Write-Debug "[$functionName] Before Loading Nintex DLLs"
		#Loading Nintex DLLs		
		[System.Reflection.Assembly]::LoadWithPartialName("Nintex.Forms.SharePoint") | Out-Null
		[System.Reflection.Assembly]::LoadWithPartialName("Nintex.Forms") | Out-Null
		
		Write-Debug "[$functionName] Before getting SPWeb"
		$curWeb= Get-SPWeb $siteURL -EA SilentlyContinue 
		if ($curWeb -ne $null)
		{
			$webID = $curWeb.ID
			Write-Debug "[$functionName] Before getting SPLibrary NintexForms"
			$nintexFormsList = $curWeb.Lists.TryGetList("NintexForms")
			if ($nintexFormsList -ne $null)
			{
				$nintexFormsListURL = $siteURL + "/" + $nintexFormsList.DefaultView.Url
				Write-Host "[$functionName] nintexFormsListURL: $nintexFormsListURL" -foregroundcolor Cyan
				
				#Preparing XML Output and File Location
				Write-Debug "[$functionName] Preparing XML Output and File Location"
				$webTitleTrim = $curWeb.Title -replace " ", ""
				#$xmlFilePath  = $destFolder + "\NintexForms_"+ $webTitleTrim + ".xml"
				$xmlFilePath  = $artefactFolder + "\NintexForms.xml"
				Write-Host "[$functionName] xmlFilePath: $xmlFilePath"  -foregroundcolor Cyan
	
				Write-Debug "[$functionName] Get an XMLTextWriter to create the XML"
				# get an XMLTextWriter to create the XML
				$xmlWriter = New-Object System.XMl.XmlTextWriter($xmlFilePath,$Null)
				$xmlWriter.Formatting = 'Indented'
				$xmlWriter.Indentation = 1
				$XmlWriter.IndentChar = "`t"

				# write xml header
				$xmlWriter.WriteStartDocument()

				# create root element "Forms" and add some attributes to it
				$xmlWriter.WriteStartElement('Forms')
			
				#############################
				##### LOOPING INTO LIST #####
				#############################
				foreach ($listItem in $nintexFormsList.Items)
				{
					
					# Form general information from SPListItem object
					$formName= $listItem.Name
					$formStatus = $listItem.Level
					$formID = $listItem["FormId"];
					$parentListID = $listItem["FormListId"];					
					$parentListName = [string]::Empty
					$parentCTID = $listItem["FormContentTypeId"];
					$parentCTName = [string]::Empty

					# Get Parent List and ContentType Names
					if ($parentListID -ne $null)
					{
						$parentListID = [Guid] $parentListID.ToString()	
						$parentList = $curWeb.Lists  | where {$_.ID -eq $parentListID} 
						if ($parentList -ne $null)
						{
							$parentListName = $parentList.Title
							if ($parentCTID -ne $null)
							{	
								$parentCT = $parentList.ContentTypes | where {$_.ID -eq $parentCTID} 
								if ($parentCT -ne $null)
								{
									$parentCTName = $parentCT.Name
									$comprehensiveFileName= ($parentListName -replace  " ", "" ) + ".xml"
									# Download file
									$oItemFile = $listItem.File
									$oItemFileBinary = $oItemFile.OpenBinary()
									$oItemFileStream = New-Object System.IO.FileStream($exportFolder + "/" + $comprehensiveFileName), Create
									$writer = New-Object System.IO.BinaryWriter($oItemFileStream)
									$writer.write($oItemFileBinary)
									$writer.Close()
									
									
									# Create the 'Form' node with attributes
									$xmlWriter.WriteStartElement('Form')
									$xmlWriter.WriteAttributeString("Name", $formName);
									$xmlWriter.WriteAttributeString("Status", $formStatus);
									$xmlWriter.WriteAttributeString("FormId", $formId);
									$xmlWriter.WriteAttributeString("ListId", $parentListID);
									$xmlWriter.WriteAttributeString("ListName", $parentListName);
									$xmlWriter.WriteAttributeString("ContentTypeID", $parentCTID);
									$xmlWriter.WriteAttributeString("ContentTypeName", $parentCTName);
									$xmlWriter.WriteAttributeString("FileName", $comprehensiveFileName );

									# Close the 'Form' node
									$xmlWriter.WriteEndElement()
									
								}
								else
								{
									Write-Warning "Content-Type '$parentCTID' does not exist anymore on List '$listName'.";
								}
							}
							else
							{
								Write-Warning "Field 'FormContentTypeId' is empty.";
							}
						}
						else
						{
							Write-Warning "List '$FormListId' does not exist anymore.";
						}
					}
					else
					{
						Write-Warning "Field 'FormListId' is empty.";
					}
				}
			}
			else
			{
				Write-Warning "[$functionName] Library 'NintexForms' not found on site '$siteURL'."
			}
			
			Write-Debug "[$functionName] Close root element 'Forms'"
			# close the "Forms" node
			$xmlWriter.WriteEndElement()
			 
			Write-Debug "[$functionName] Finalize file"
			# finalize the document
			$xmlWriter.WriteEndDocument()
			$xmlWriter.Flush()
			$xmlWriter.Close()

		}
		else
		{
			Write-Warning "[$functionName] Site '$siteURL' not found."
		}
	}
	catch
	{
		Write-Host ""
		Write-Host "/!\ [$functionName] An exception has been caught /!\ " -Foregroundcolor Red `r
		Write-Host "Type: " $_.Exception.GetType().FullName -Foregroundcolor Red `r
		Write-Host "Message: " $_.Exception.Message -Foregroundcolor Red `r
		Write-Host "Stacktrace: `n" $_.Exception.Stacktrace -Foregroundcolor Red `r
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
	Export All Forms fom Site $siteURL and generate and XML file with all Forms settings
	
.DESCRIPTION
	Export All Forms fom Site $siteURL and generate and XML file with all Forms settings
	
.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER destFolder
	(Optional) Location of the exported form
	If not specified, will get the current location 
		
.EXAMPLE
	ExportAllFormsFromSite -siteURL <SiteURL> [-destFolder <dstFolderPath>]
	
.OUTPUTS
	One XML file containing the list of all the Forms on site $siteURL and their settings (Name, List (ID+Name), Content-Type (ID+Name) and FileName)
	One folder NintexForms\ to store all forms files
	One .XML file by Workflow

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 17.01.2017
	Last Updated by: JBO
	Last Updated: 22.02.2017
#>
function ExportAllFormsFromSite()
{
	[CmdletBinding()]
	param
	(
	    [Parameter(Mandatory=$true, Position=1)]
	    [string]$siteURL, 
		[Parameter(Mandatory=$false, Position=2)]
	    [string]$destFolder
	)
		
	$functionName = $MyInvocation.MyCommand.Name
	Write-Debug "[$functionName] Entering function" 
	Write-Debug "[$functionName] Parameter / siteURL: $siteURL" 

	try
	{
		if([string]::IsNullOrEmpty($destFolder)) 
		{
			$destFolder = Get-Location
			Write-Warning "[$functionName] Paramater destFolder is empty, will set it to the current location." 
			Write-Host "[$functionName] destFolder: $destFolder" -foregroundcolor Cyan
		}
		
		Write-Debug "[$functionName] Before creating Export folder"
		$exportFolder = $destFolder + "\NintexForms\"
		if (!(Test-Path $exportFolder))
		{
			New-Item $exportFolder -type Directory -Force | Out-Null
		}
		
		$artefactFolder = $destFolder + "\Artefacts\"
		if (!(Test-Path $exportFolder))
		{
			New-Item $artefactFolder -type Directory -Force | Out-Null
		}
		
		Write-Debug "[$functionName] Before Loading Nintex DLLs"
		#Loading Nintex DLLs		
		[System.Reflection.Assembly]::LoadWithPartialName("Nintex.Forms.SharePoint") | Out-Null
		[System.Reflection.Assembly]::LoadWithPartialName("Nintex.Forms") | Out-Null
		
		Write-Debug "[$functionName] Before getting SPWeb"
		$curWeb= Get-SPWeb $siteURL -EA SilentlyContinue 
		if ($curWeb -ne $null)
		{
			$webID = $curWeb.ID
			Write-Debug "[$functionName] Before getting SPLibrary NintexForms"
			$nintexFormsList = $curWeb.Lists.TryGetList("NintexForms")
			if ($nintexFormsList -ne $null)
			{
				$nintexFormsListURL = $siteURL + "/" + $nintexFormsList.DefaultView.Url
				Write-Host "[$functionName] nintexFormsListURL: $nintexFormsListURL" -foregroundcolor Cyan
				
				#Preparing XML Output and File Location
				Write-Debug "[$functionName] Preparing XML Output and File Location"
				$webTitleTrim = $curWeb.Title -replace " ", ""
				#$xmlFilePath  = $destFolder + "\NintexForms_"+ $webTitleTrim + ".xml"
				$xmlFilePath  = $artefactFolder + "\NintexForms.xml"
				Write-Host "[$functionName] xmlFilePath: $xmlFilePath"  -foregroundcolor Cyan
	
				Write-Debug "[$functionName] Get an XMLTextWriter to create the XML"
				# get an XMLTextWriter to create the XML
				$xmlWriter = New-Object System.XMl.XmlTextWriter($xmlFilePath,$Null)
				$xmlWriter.Formatting = 'Indented'
				$xmlWriter.Indentation = 1
				$XmlWriter.IndentChar = "`t"

				# write xml header
				$xmlWriter.WriteStartDocument()

				# create root element "Forms" and add some attributes to it
				$xmlWriter.WriteStartElement('Forms')
			
				#############################
				##### LOOPING INTO LIST #####
				#############################
				foreach ($listItem in $nintexFormsList.Items)
				{
					
					# Form general information from SPListItem object
					$formName= $listItem.Name
					$formStatus = $listItem.Level
					$formID = $listItem["FormId"];
					$parentListID = $listItem["FormListId"];					
					$parentListName = [string]::Empty
					$parentCTID = $listItem["FormContentTypeId"];
					$parentCTName = [string]::Empty
	
					$oItemFile = $listItem.File
					$oItemFileBinary = $oItemFile.OpenBinary()
					$oItemFileStream = New-Object System.IO.FileStream($exportFolder + "/" + $oItemFile.Name), Create
					$writer = New-Object System.IO.BinaryWriter($oItemFileStream)
					$writer.write($oItemFileBinary)
					$writer.Close()
	
					# Get Parent List and ContentType Names
					if ($parentListID -ne $null)
					{
						$parentListID = [Guid] $parentListID.ToString()	
						$parentList = $curWeb.Lists  | where {$_.ID -eq $parentListID} 
						if ($parentList -ne $null)
						{
							$parentListName = $parentList.Title
							if ($parentCTID -ne $null)
							{	
								$parentCT = $parentList.ContentTypes | where {$_.ID -eq $parentCTID} 
								if ($parentCT -ne $null)
								{
									$parentCTName = $parentCT.Name
								}
								else
								{
									Write-Warning "Content-Type '$parentCTID' does not exist anymore on List '$listName'.";
								}
							}
							else
							{
								Write-Warning "Field 'FormContentTypeId' is empty.";
							}
						}
						else
						{
							Write-Warning "List '$FormListId' does not exist anymore.";
						}
					}
					else
					{
						Write-Warning "Field 'FormListId' is empty.";
					}
				
					# Create the 'Form' node with attributes
					$xmlWriter.WriteStartElement('Form')
					$xmlWriter.WriteAttributeString("Name", $formName);
					$xmlWriter.WriteAttributeString("Status", $formStatus);
					$xmlWriter.WriteAttributeString("FormId", $formId);
					$xmlWriter.WriteAttributeString("ListId", $parentListID);
					$xmlWriter.WriteAttributeString("ListName", $parentListName);
					$xmlWriter.WriteAttributeString("ContentTypeID", $parentCTID);
					$xmlWriter.WriteAttributeString("ContentTypeName", $parentCTName);
					$xmlWriter.WriteAttributeString("FileName", $oItemFile.Name );


					# Close the 'Form' node
					$xmlWriter.WriteEndElement()
				
				}
			}
			else
			{
				Write-Warning "[$functionName] Library 'NintexForms' not found on site '$siteURL'."
			}
			
			Write-Debug "[$functionName] Close root element 'Forms'"
			# close the "Forms" node
			$xmlWriter.WriteEndElement()
			 
			Write-Debug "[$functionName] Finalize file"
			# finalize the document
			$xmlWriter.WriteEndDocument()
			$xmlWriter.Flush()
			$xmlWriter.Close()

		}
		else
		{
			Write-Warning "[$functionName] Site '$siteURL' not found."
		}
	}
	catch
	{
		Write-Host ""
		Write-Host "/!\ [$functionName] An exception has been caught /!\ " -Foregroundcolor Red `r
		Write-Host "Type: " $_.Exception.GetType().FullName -Foregroundcolor Red `r
		Write-Host "Message: " $_.Exception.Message -Foregroundcolor Red `r
		Write-Host "Stacktrace: `n" $_.Exception.Stacktrace -Foregroundcolor Red `r
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
	Generate a form digest to ensure Forms deployment
	
.DESCRIPTION
	Generate a form digest to ensure Forms deployment
	
.PARAMETER siteURL
	URL of the SharePoint site

.EXAMPLE
	GetFormDigest -siteURL <siteURL>
	
.OUTPUTS
	String $newFormDigest if call to _api/contextinfo successful, null otherwise

.LINK
	Inspired by https://spalexandre.wordpress.com/2014/08/05/automatiser-les-deploiements-nintex-workflow-et-nintex-forms/
	
.NOTES
	Created by: JBO
	Created: 18.01.2017
	Last Updated by: JBO
	Last Updated: 19.01.2017
#>
function GetFormDigest()
{
	[CmdletBinding()]
	param
	(
	    [Parameter(Mandatory=$true, Position=1)]
	    [string]$siteURL
	)

	$functionName = $MyInvocation.MyCommand.Name
	Write-Debug "[$functionName] Entering function" 
	Write-Debug "[$functionName] Parameter / siteURL: $siteURL" 

	[string] $newFormDigest = $null
	try
	{
		Write-Debug "[$functionName] Building form digest URL"
		[System.Reflection.Assembly]::LoadWithPartialName("System.IO") | Out-Null
		$formDigestRequest = [Microsoft.SharePoint.Utilities.SPUtility]::ConcatUrls($siteURL, "_api/contextinfo")
		Write-Host "[$functionName] formDigestRequest: $formDigestRequest"

		$formDigestUri = New-Object System.Uri($formDigestRequest)
	 
		$credCache = New-Object System.Net.CredentialCache
		$credCache.Add($formDigestUri, "NTLM", [System.Net.CredentialCache]::DefaultNetworkCredentials)
		$spRequest = [System.Net.HttpWebRequest] [System.Net.HttpWebRequest]::Create($formDigestRequest)
		$spRequest.Credentials = $credCache
		$spRequest.Method = "POST"
		$spRequest.Accept = "application/json;odata=verbose"
		$spRequest.ContentLength = 0
	 
		[System.Net.HttpWebResponse] $endpointResponse = [System.Net.HttpWebResponse] $spRequest.GetResponse()
		[System.IO.Stream]$postStream = $endpointResponse.GetResponseStream()
		[System.IO.StreamReader] $postReader = New-Object System.IO.StreamReader($postStream)
		$results = $postReader.ReadToEnd()
	 
		$postReader.Close()
		$postStream.Close()
	 
		#Get the FormDigest Value
		$startTag = "FormDigestValue"
		$endTag = "LibraryVersion"
		$startTagIndex = $results.IndexOf($startTag) + 1
		$endTagIndex = $results.IndexOf($endTag, $startTagIndex)
		
		if (($startTagIndex -ge 0) -and  ($endTagIndex -gt $startTagIndex))
		{
			$newFormDigest = $results.Substring($startTagIndex + $startTag.Length + 2, $endTagIndex - $startTagIndex - $startTag.Length - 5)
		}
	 
		Write-Host "[$functionName] newFormDigest: $newFormDigest"
	}
	catch
	{
		$newFormDigest = $null
		Write-Host ""
		Write-Host "/!\ [$functionName] An exception has been caught /!\ " -Foregroundcolor Red `r
		Write-Host "Type: " $_.Exception.GetType().FullName -Foregroundcolor Red `r
		Write-Host "Message: " $_.Exception.Message -Foregroundcolor Red `r
		Write-Host "Stacktrace: `n" $_.Exception.Stacktrace -Foregroundcolor Red `r
	}
	finally
	{
		Write-Debug "[$functionName] Exiting function" 
	}
    return $newFormDigest
}


<#
.SYNOPSIS
	Convert XML Form file $filename into a JSON object
	
.DESCRIPTION
	Convert XML Form file $filename into a JSON object
	
.PARAMETER fileName
	Name of the XML Form File to be converted

.EXAMPLE
	ConvertXMLFormFileIntoJSON -fileName <fileName>
	
.OUTPUTS
	JSON object if file successfully loaded, null otherwise

.LINK
	Inspired by https://spalexandre.wordpress.com/2014/08/05/automatiser-les-deploiements-nintex-workflow-et-nintex-forms/
	
.NOTES
	Created by: JBO
	Created: 18.01.2017
	Last Updated by: JBO
	Last Updated: 19.01.2017
#>
function ConvertXMLFormFileIntoJSON()
{
	[CmdletBinding()]
	param
	(
	    [Parameter(Mandatory=$true, Position=1)]
	    [string]$fileName
	)

	$functionName = $MyInvocation.MyCommand.Name
	Write-Debug "[$functionName] Entering function" 
	Write-Debug "[$functionName] Parameter / fileName: $fileName" 

	$json = $null
	
	try
	{
		Write-Debug "[$functionName] Getting Json from Nintex Form Xml" 
 		[System.Reflection.Assembly]::LoadWithPartialName("Nintex.Forms.SharePoint") | Out-Null
		[System.Reflection.Assembly]::LoadWithPartialName("Nintex.Forms") | Out-Null
    	
		if (Test-Path $fileName)
		{
			[byte[]] $fileBytes = [System.IO.File]::ReadAllBytes($FileName)
			$form = [Nintex.Forms.FormsHelper]::XmlToObject([Nintex.Forms.NFUtilities]::ConvertByteArrayToString($fileBytes, [System.Text.Encoding]::UTF8))
			
			$form.LiveSettings.Url = ""
			$form.LiveSettings.ShortUrl = ""
			$form.RefreshLayoutDisplayNames()
			$form.Id = [guid]::NewGuid()
	 
			$json = [Nintex.Forms.FormsHelper]::ObjectToJson($form)
			#Write-Host "[$functionName] json: $json"
		}
		else		
		{
			Write-Warning "[$functionName] File not found at '$fileName'."	
			$json = $null
		}
	}
	catch
	{
		$json = $null
		Write-Host ""
		Write-Host "/!\ [$functionName] An exception has been caught /!\ " -Foregroundcolor Red `r
		Write-Host "Type: " $_.Exception.GetType().FullName -Foregroundcolor Red `r
		Write-Host "Message: " $_.Exception.Message -Foregroundcolor Red `r
		Write-Host "Stacktrace: `n" $_.Exception.Stacktrace -Foregroundcolor Red `r
	}
	finally
	{
		Write-Debug "[$functionName] Exiting function" 
	}
    return $json
}


<#
.SYNOPSIS
	Browse and parse the formsDescriptionXML XML object and deploy the forms
	
.DESCRIPTION
	Browse and parse the file formsDescriptionXML
		For each node, Deploy the Forms	
	
.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER formsDescriptionXML
	XML Object representing the Forms to be deployed
	
.PARAMETER destFolder
	(Optional) Location of the exported form
	If not specified, will get the current location 
	
.EXAMPLE
	browseAndDeployFormsXML -siteURL <SiteURL> -formsDescriptionXML <XMLObjectToParse> [-destFolder <destFolderPath>]
	
.OUTPUTS
	None

.LINK
	Inspired by https://spalexandre.wordpress.com/2014/08/05/automatiser-les-deploiements-nintex-workflow-et-nintex-forms/
	
.NOTES
	Created by: JBO
	Created: 18.01.2017
	Last Updated by: JBO
	Last Updated: 19.01.2017
#>
function browseAndDeployFormsXML()
{
	[CmdletBinding()]
	param
    (
        [Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[XML]$formsDescriptionXML, 
		[Parameter(Mandatory=$false, Position=3)]
	    [string]$srcFolder
	)

    try
    {
		$functionName = $MyInvocation.MyCommand.Name
		Write-Debug "[$functionName] Entering function" 
		Write-Debug "[$functionName] Parameter / siteURL: $siteURL"
	
		if([string]::IsNullOrEmpty($srcFolder)) 
		{
			$srcFolder = Get-Location
			Write-Warning "[$functionName] Paramater srcFolder is empty, will set it to the current location." 
			Write-Host "[$functionName] srcFolder: $srcFolder" -foregroundcolor Cyan
		}
		
		$curWeb = Get-SPWeb $siteURL -EA SilentlyContinue
		if ($curWeb -ne $null)
		{
			if($formsDescriptionXML.HasChildNodes)
			{
				Write-Debug "[$functionName] About to loop on Form nodes." 
				$formLists = $formsDescriptionXML.SelectNodes("/Forms")
				foreach($curForm in $formLists.Form)
				{					
					$formDigest = GetFormDigest -siteURL $siteURL
					if(!([string]::IsNullOrEmpty($formDigest)))
					{ 
						Write-Debug "[$functionName] Form Digest OK ... "
						Write-Debug "[$functionName] Prepare prerequisites for WS call"
						
						#Step1: WS address URI
						$addressUrl = [Microsoft.SharePoint.Utilities.SPUtility]::ConcatUrls($siteURL, "_vti_bin/NintexFormsServices/NfRestService.svc/PublishForm")
						$addressUri = New-Object System.Uri($addressUrl)
						
						#Step2: Create the web request
						[System.Net.HttpWebRequest] $request = [System.Net.WebRequest]::Create($addressUri)
						$request.Method = "POST";
						$request.ContentType = "application/json; charset=utf-8";
						$request.Accept = "application/json, text/javascript, */*; q=0.01"
						$request.Headers.Add("X-RequestDigest", $formDigest);
						$request.Headers.Add("X-Requested-With", "XMLHttpRequest")

						#Step3: Add authentication to request (Prompt ? Impersonnalisation)
						$credCache = New-Object System.Net.CredentialCache
						$cred = [System.Net.CredentialCache]::DefaultNetworkCredentials
						$credCache.Add($addressUri, "NTLM", [System.Net.CredentialCache]::DefaultNetworkCredentials)
						$request.Credentials = $credCache 					
					
						$fileName = $curForm.FileName
						Write-Host "[$functionName] About to deploy form '$($curForm.FileName)' on list '$($curForm.ListName)'" -ForegroundColor Magenta
						#Write-Host "[$functionName] ContentTypeID $($curForm.ContentTypeID)"
						#Write-Host "[$functionName] ContentTypeName $($curForm.ContentTypeName)"
						
						#Step4: Get JSON for the file 
						$fullPathToFile = $srcFolder + "\" + $fileName
						$form = ConvertXMLFormFileIntoJSON -FileName $fullPathToFile						
						
						#Step5: Get listID and ctID						
						$listName = $curForm.ListName
						if(!([string]::IsNullOrEmpty($listName)))
						{
							$parentList = $curWeb.Lists.TryGetList($listName);
							if ($parentList -ne $null)
							{
								$targetListId = "{$($parentList.ID)}"
								
								$ctName = $curForm.ContentTypeName
								if(!([string]::IsNullOrEmpty($ctName)))
								{
									$parentCT = $parentList.ContentTypes | where {$_.Name -eq $ctName} 
									if ($parentCT -ne $null)
									{
										$targetCTId = $parentCT.ID
									}
									else
									{
										Write-Host "Warning: Content-Type '$parentCTID' does not exist anymore on List '$listName'.";
									}
								}
								else
								{
									Write-Host "Warning: ContentTypeName attribute in the XML is empty.";
								}
																
								#Step6: Create the message
								
								#Create the data we want to send
								$data = "{`"contentTypeId`": `"$targetCTId`", `"listId`": `"$targetListId`", `"form`": $form }"	
								
								#Create a byte array of the data we want to send
								$utf8 = New-Object System.Text.UTF8Encoding
								[byte[]] $byteData = $utf8.GetBytes($data.ToString())

								#Set the content length in the request headers
								$request.ContentLength = $byteData.Length;

								#Write data
								$postStream = $request.GetRequestStream()
								$postStream.Write($byteData, 0, $byteData.Length);
																
								#Step7: Call the WS and get the response
								[System.Net.HttpWebResponse] $response = [System.Net.HttpWebResponse] $request.GetResponse()
 
								#Step8: Parse the response stream
								[System.IO.StreamReader] $reader = New-Object System.IO.StreamReader($response.GetResponseStream())
								
								$strResult = $reader.ReadToEnd()
								$jsonResult = ConvertFrom-Json $strResult
								Write-Host "[$functionName] Form '$($curForm.Name)' successfully deployed on list '$($curForm.ListName)'" -ForegroundColor Green
								Write-Host "[$functionName] Current version is $($jsonResult.PublishFormResult.Version)" -ForegroundColor Green
								
								# try {
									# $strResult = $reader.ReadToEnd()
									# $jsonResult = ConvertFrom-Json $strResult
									# Write-Host "[$functionName] Form '$($curForm.Name)' successfully deployed on list '$($curForm.ListName)'" -ForegroundColor Green
									# Write-Host "[$functionName] Form '$($curForm.Name)' current version is $($jsonResult.PublishFormResult.Version)" -ForegroundColor Green
								# }
								# catch [Exception] {
									# Write-Host "[$functionName] Warning: Unable to get version fron Publish Form Result." -ForegroundColor Yellow
									# Write-Host "[$functionName] Form Publishing may have failed." -ForegroundColor Yellow
									# Write-Host "[$functionName] PublishFormResult is : $($strResult)" -ForegroundColor Yellow
								# }									
							}
							else
							{
								Write-Host "Warning: ListName '$listName' not found on site '$siteURL'.";
							}							
						}
						else
						{
							Write-Host "Warning: ListName attribute in the XML is empty.";
						}					
					}
					else
					{
						Write-Host "Warning: FormDigest is empty.";
					}

					if($postStream -ne $null) 
					{ 
						$postStream.Dispose() 
					}
				}#end foreach
			}
			else
			{
				Write-Warning "[$functionName] Forms XML definition file is empty."
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
		if($postStream -ne $null) 
		{ 
			$postStream.Dispose() 
		}
		
		if ($curWeb -ne $null)
		{
			$curWeb.Dispose();
		}
		Write-Debug "[$functionName] Exiting function" 
	} 
}