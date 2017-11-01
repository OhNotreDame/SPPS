param
(
 [Parameter(Mandatory=$true, Position=1)]
 [string]$siteURL,
 [Parameter(Mandatory=$true, Position=2)]
 [string]$structureGroupName,
 [Parameter(Mandatory=$false, Position=3)]
 [string]$solutionFolderPath
 
)

Clear-Host
Remove-Module *

$scriptName = $([System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Definition))
Write-Host "************************************************************************" -ForegroundColor Gray `r
Write-Host "$scriptName # Script started." -ForegroundColor Gray `r
Write-Host "************************************************************************" -ForegroundColor Gray `r		

#####################################################
# Loading SharePoint Assembly and PS Snapin
#####################################################
Add-PsSnapin Microsoft.Sharepoint.PowerShell -EA silentlycontinue
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint") | Out-Null

#####################################################
# Starting SPAssignment
#####################################################
Start-SPAssignment -Global

#####################################################
# Setting Path Variables
#####################################################
$scriptdir = $PSScriptRoot
Set-Variable -Name "scriptPath" -Value $scriptdir -Scope Global


try
{
	if([string]::IsNullOrEmpty($solutionFolderPath)) 
	{
		$solutionFolderPath = Get-Location
		Write-Warning "[$scriptName] Paramater solutionFolderPath is empty, will set it to the current location." 
		Write-Host "[$scriptName] solutionFolderPath: $solutionFolderPath" -foregroundcolor Cyan
	}
	
	$ModuleFolderPath = "D:\QuickDeployFW\Modules"
	$destFolderArtefacts = "$solutionFolderPath\Artefacts"
	$destFolderLogs = "$solutionFolderPath\Logs"
	$destFolderWF = "$solutionFolderPath\NintexWorkflows" 
	$destFolderForms = "$solutionFolderPath\NintexForms" 

	Write-Host "" `r
	Write-Host "************************************************************************" -ForegroundColor Magenta `r
	Write-Host "$scriptName # Parameters and Settings" -ForegroundColor Magenta `r
	Write-Host "************************************************************************" -ForegroundColor Magenta `r		
	Write-Host "siteURL: $siteURL" -ForegroundColor Gray `r
	Write-Host "structureGroupName: $structureGroupName" -ForegroundColor Gray `r
	Write-Host "solutionFolderPath: $solutionFolderPath" -ForegroundColor Gray `r
	Write-Host "************************************************************************" -ForegroundColor Magenta `r	
	Write-Host "destFolderArtefacts: $destFolderArtefacts" -ForegroundColor Gray `r
	Write-Host "ModuleFolderPath: $ModuleFolderPath" -ForegroundColor Gray `r
	Write-Host "scriptdir: $scriptdir" -ForegroundColor Gray `r
	Write-Host "************************************************************************" -ForegroundColor Magenta `r	
		
	#####################################################
	# Loading QuickDeploy Custom Modules
	#####################################################
	Write-Host "" `r
	Write-Host "************************************************************************" -ForegroundColor Magenta `r
	Write-Host "About to import QuickDeploy Modules " -ForegroundColor Magenta `r
	Write-Host "************************************************************************" -ForegroundColor Magenta `r
	Import-Module "$ModuleFolderPath\SPHelpers\SPHelpers.psm1"
	Import-Module "$ModuleFolderPath\SPFileUploader\SPFileUploader.psm1"
	Import-Module "$ModuleFolderPath\SPCommonStructure\SPCommonStructure.psm1"
	Import-Module "$ModuleFolderPath\SPSiteColumns\SPSiteColumns.psm1"
	Import-Module "$ModuleFolderPath\SPSiteContentTypes\SPSiteContentTypes.psm1"
	Import-Module "$ModuleFolderPath\SPLists\SPLists.psm1"
	Import-Module "$ModuleFolderPath\SPListViews\SPListViews.psm1"
	Import-Module "$ModuleFolderPath\NintexWorkflows\NintexWorkflows.psm1"
	Import-Module "$ModuleFolderPath\NintexForms\NintexForms.psm1"
	Import-Module "$ModuleFolderPath\SPSiteGroups\SPSiteGroups.psm1"
	Import-Module "$ModuleFolderPath\SPListPermissions\SPListPermissions.psm1"
	Import-Module "$ModuleFolderPath\SPSitePages\SPSitePages.psm1"
	
	Write-Host "QuickDeploy Modules Successfully Imported" -ForegroundColor Green `r
	Write-Host "************************************************************************" -ForegroundColor Magenta `r
	
	###########################
	#### CREATING FOLDERS  ####
	###########################
	
	
	#### Logs Folder ####
	if(!(Test-Path $destFolderLogs))
	{
		New-Item $destFolderLogs -type Directory -Force | Out-Null
	}
	
	#### Artefacts Folder ####
	if(!(Test-Path $destFolderArtefacts))
	{
		New-Item $destFolderArtefacts -type Directory -Force | Out-Null
	}
	
	###########################
	#### TRANSCRIPT / LOGS ####
	###########################
	$logsFileName = $destFolderLogs + "\"+ $scriptName  + "_"+ $(get-date -format 'yyyyMMdd_HHmmss') + ".log"
	$logsFileName
	Start-Transcript -path $logsFileName -noclobber -Append
	
	$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
	if($curWeb -ne $null)
	{
		
		######################################################
		#### Extract Site Columns
		######################################################
		
		#Create a XML File to Export Fields
		$SPSiteColumnsFile = "$destFolderArtefacts\SPSiteColumns.xml"
		
		Write-Host ""
		Write-Host "******************************************************************************" -ForegroundColor Magenta `r	
		Write-Host "$scriptName # About to create XML file for <SPSiteColumns>" -ForegroundColor Magenta `r
		Write-Host "******************************************************************************" -ForegroundColor Magenta `r	

		# XML File -- Get an XMLTextWriter to create the XML
		Write-Debug "[$scriptName] Get an XMLTextWriter to create the XML"
		
		$xmlWriter = New-Object System.XMl.XmlTextWriter($SPSiteColumnsFile,$Null)
		
		# XML File -- Choose a pretty formatting
		Write-Debug "[$scriptName] Choose a pretty formatting"
		$xmlWriter.Formatting = 'Indented'
		$xmlWriter.Indentation = 1
		$XmlWriter.IndentChar = "`t"
		
		# XML File -- Write header
		Write-Debug "[$scriptName] Write Header"
		$xmlWriter.WriteStartDocument()
	
		# XML File -- Create root element "Fields" and add some attributes to it
		$xmlWriter.WriteStartElement('Fields')


		$curWeb.Fields | ForEach-Object {
			if ($_.Group -eq $structureGroupName) {
					$xmlWriter.WriteRaw($_.SchemaXml) 
			}
		}
		
		# XML File -- Close the "SPSiteColumns" node
		Write-Debug "[$scriptName] Close root element 'Fields'"
		$xmlWriter.WriteEndElement()
		
		# XML File -- Finalize the document 
		$xmlWriter.WriteEndDocument()
		$xmlWriter.Flush()
		$xmlWriter.Close()
		Write-Host "$scriptName # XML file successfully finalized."  -ForegroundColor Green `r		

		######################################################
		#### Extract Site Content-Types
		######################################################
		
		#Create a XML File to Export Content-Types
		$SPSiteContentTypesFile = "$destFolderArtefacts\SPSiteContentTypes.xml"
		
		Write-Host ""
		Write-Host "******************************************************************************" -ForegroundColor Magenta `r	
		Write-Host "$scriptName # About to create XML file for <SPSiteContentTypes>" -ForegroundColor Magenta `r
		Write-Host "******************************************************************************" -ForegroundColor Magenta `r	

		# XML File -- Get an XMLTextWriter to create the XML
		Write-Debug "[$scriptName] Get an XMLTextWriter to create the XML"
		
		$xmlWriter = New-Object System.XMl.XmlTextWriter($SPSiteContentTypesFile,$Null)
		
		# XML File -- Choose a pretty formatting
		Write-Debug "[$scriptName] Choose a pretty formatting"
		$xmlWriter.Formatting = 'Indented'
		$xmlWriter.Indentation = 1
		$XmlWriter.IndentChar = "`t"
		
		# XML File -- Write header
		Write-Debug "[$scriptName] Write Header"
		$xmlWriter.WriteStartDocument()
	
		# XML File -- Create root element "ContentTypes" and add some attributes to it
		$xmlWriter.WriteStartElement('ContentTypes')


		foreach ($ct in $curWeb.ContentTypes)
        {
			if ($ct.Group -eq $structureGroupName) {
				
				$parentCtId = $ct.Parent.Id 
				
				Write-Host "$($ct.Name) - $parentCtId"
				#$xmlWriter.WriteRaw($_.SchemaXml) 
				
				$parentCt = "Item"
								
				if ($parentCtId-eq "0x0101") {
					$parentCt = "Document"
				}
				else {
					$parentCt = "Item"
				}
				
				
				# <ContentType Name="Demande" Group="RH - Places de Parc" Description="" ParentContentType="Item" >
					# <FieldRefs>
						# <Field ID="F049FF1D-879F-43DA-9CDE-DDC991026745" Name="linkedToMAJInfoBenef" DisplayName="_linkedToMAJInfoBenef" Required="FALSE"/>
					# </FieldRefs>
				# </ContentType>
				
				$xmlWriter.WriteStartElement('ContentType')
				$xmlWriter.WriteAttributeString("Name", $ct.Name);
				$xmlWriter.WriteAttributeString("Group", $ct.Group);
				$xmlWriter.WriteAttributeString("Description", $ct.Description);
				$xmlWriter.WriteAttributeString("ParentContentType", $parentCt);
				
				$xmlWriter.WriteStartElement('FieldRefs')
				
				foreach ($field in $ct.Fields)
                {
						
					#<FieldRef ID="53101f38-dd2e-458c-b245-0c236cc13d1a" Name="AssignedTo" DisplayName="Assigné à" Required="TRUE"/>
					
					$xmlWriter.WriteStartElement('FieldRef')
					$xmlWriter.WriteAttributeString("ID", $field.ID);
					$xmlWriter.WriteAttributeString("Name", $field.StaticName);
					$xmlWriter.WriteAttributeString("DisplayName", $field.Title);
					$xmlWriter.WriteAttributeString("Required", $field.Required);
					
					#Closing Previous <FieldRef> 
					$xmlWriter.WriteEndElement()
				}

				#Closing Previous <FieldRefs> 
				$xmlWriter.WriteEndElement()
				
				#Closing Previous <ContentType> 
				$xmlWriter.WriteEndElement()	
			}
		}

		# XML File -- Close the "ContentTypes" node
		Write-Debug "[$scriptName] Close root element 'ContentTypes'"
		$xmlWriter.WriteEndElement()
		
		# XML File -- Finalize the document 
		$xmlWriter.WriteEndDocument()
		$xmlWriter.Flush()
		$xmlWriter.Close()
		Write-Host "$scriptName # XML file successfully finalized."  -ForegroundColor Green `r	
		
	}
	else
	{
		Write-Warning "Site '$siteURL' does not exist."
	}		
}
catch [Exception]
{	
	Write-Host "/!\ $scriptName An exception has been caught /!\ "  -ForegroundColor Red `r
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

	Write-Host "************************************************************************" -ForegroundColor Gray `r		
	Write-Host "$scriptName # Script ended." -ForegroundColor Gray `r
	Write-Host "************************************************************************" -ForegroundColor Gray `r					
	
	#####################################################
	# Stopping SPAssignment and Transcript
	#####################################################
	Stop-SPAssignment -Global
	Stop-Transcript
}