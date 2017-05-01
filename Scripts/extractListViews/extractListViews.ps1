<#
.SYNOPSIS
	Extract all the views of all the lists of site $siteURL and generate an XML file to import them somewhere else
	
.DESCRIPTION
	Extract all the views of all the lists of site $siteURL and generate an XML file to import them somewhere else
	
.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER destinationFolderPath
	(Optional) Location of the exported XML files
	If not specified, will get the current location
	
.EXAMPLE
	extractListViews.ps1 -siteURL <siteURL> [-destinationFolderPath <destinationFolderPath>]
	
.OUTPUTS
	1 file containg all Site Lists Views (SPListViews.xml)

.LINK
	https://github.com/OhNotreDame/SPPS
	
.NOTES
	Created by: JBO
	Created: 13.03.2017
	Last Updated by: JBO
	Last Updated: 13.03.2017
#>

param
(
	[Parameter(Mandatory=$true, Position=1)]
	[ValidateNotNullOrEmpty()]
	[string]$siteURL,
	[Parameter(Mandatory=$false, Position=2)]
	[string]$destinationFolderPath 
)

Clear-Host
Remove-Module *

$scriptName = $([System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Definition))
Write-Host "******************************************************************************" -ForegroundColor Gray `r
Write-Host "$scriptName # Script started." -ForegroundColor Gray `r
Write-Host "******************************************************************************" -ForegroundColor Gray `r		

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
	if([string]::IsNullOrEmpty($destinationFolderPath)) 
	{
		$destinationFolderPath = Get-Location
		Write-Warning "[$scriptName] Paramater destinationFolderPath is empty, will set it to the current location." 
		Write-Host "[$scriptName] destinationFolderPath: $destinationFolderPath" -foregroundcolor Cyan
	}
	
	$ModuleFolderPath = "D:\QuickDeployFW\Modules"
	$destFolderArtefacts = "$destinationFolderPath\Artefacts"
	$destFolderLogs = "$destinationFolderPath\Logs"
	
	# XML Files Settings
	$fileNameForListsViews = "SPListViews.xml"
	
	# Excluded Lists
	$excludedLists = "Documents", "Forms Templates", "MicroFeed", "Site Assets", "Site Pages", "Style Library", "Workflow Tasks", "ContentAssets"
	
	Write-Host "" `r
	Write-Host "******************************************************************************" -ForegroundColor Magenta `r
	Write-Host "$scriptName # Parameters " -ForegroundColor Magenta `r
	Write-Host "******************************************************************************" -ForegroundColor Magenta `r		
	Write-Host "siteURL: $siteURL" -ForegroundColor Gray `r
	Write-Host "fileNameForListsViews: $fileNameForListsViews" -ForegroundColor Gray `r
	Write-Host "excludedLists: $excludedLists" -ForegroundColor Gray `r
	Write-Host "******************************************************************************" -ForegroundColor Magenta `r	
	Write-Host "$scriptName # Other Settings" -ForegroundColor Magenta `r
	Write-Host "******************************************************************************" -ForegroundColor Magenta `r
	Write-Host "scriptdir: $scriptdir" -ForegroundColor Gray `r
	Write-Host "destFolderArtefacts: $destFolderArtefacts" -ForegroundColor Gray `r
	Write-Host "ModuleFolderPath: $ModuleFolderPath" -ForegroundColor Gray `r
	Write-Host "******************************************************************************" -ForegroundColor Magenta `r	
	
	#####################################################
	# Loading QuickDeploy Custom Modules
	#####################################################
	Write-Host "" `r
	Write-Host "******************************************************************************" -ForegroundColor Magenta `r
	Write-Host "$scriptName # About to import QuickDeploy Modules " -ForegroundColor Magenta `r
	Write-Host "******************************************************************************" -ForegroundColor Magenta `r
	Import-Module "$ModuleFolderPath\SPHelpers\SPHelpers.psm1"
	Write-Host "$scriptName # QuickDeploy Modules Successfully Imported" -ForegroundColor Green `r
	Write-Host "******************************************************************************" -ForegroundColor Magenta `r
	
	###########################
	#### CREATING FOLDERS  ####
	###########################
	if(!(Test-Path $destFolderLogs))
	{
		New-Item $destFolderLogs -type Directory -Force | Out-Null
	}
	if (!(Test-Path $destFolderArtefacts))
	{
		New-Item $destFolderArtefacts -type Directory -Force | Out-Null
	}
	
	###########################
	#### TRANSCRIPT / LOGS ####
	###########################
	$logsFileName = $destFolderLogs + "\" + $scriptName  + "_"+ $(get-date -format 'yyyyMMdd_HHmmss') + ".log"
	$logsFileName
	Start-Transcript -path $logsFileName -noclobber -Append
	
	$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
	if($curWeb -ne $null)
	{
		Write-Host ""
		Write-Host "******************************************************************************" -ForegroundColor Magenta `r	
		Write-Host "$scriptName # About to browse Lists in order to prepare XML" -ForegroundColor Magenta `r
		Write-Host "******************************************************************************" -ForegroundColor Magenta `r	
		# XML File -- Preparing View Details
		$AddViewArray = @()
		$UpdateViewArray = @()
		
		# Get All Visible (not hidden) lists that are not excluded
		$toExportLists = $curWeb.Lists | Where-Object { $_.Hidden -eq $false -and $excludedLists -notcontains $_.Title} #| Sort-Object $_.Title 
	
		foreach($curList in $toExportLists)
		{
			$curListDisplayName = $curList.Title.ToString()
			$curListInternalName = $curList.RootFolder.ToString()
			Write-Host "List: $curListDisplayName " -ForegroundColor Yellow `r
			
			$ViewNamePos = $curListInternalName.Length

			# Get All Views that are not empty-named
			#----------------------------------------
			$toExportViews = $curList.Views | Where-Object { $_.Hidden -eq $false -and  $_.Title -ne ""} #| Sort-Object $_.Title
			foreach($curView in $toExportViews )
			{
				$curViewDisplayName = $curView.Title.ToString()
				$curViewURL = $curView.URL.ToString()
				Write-Host "... View: '$curViewDisplayName'" -ForegroundColor Cyan `r
				
				$curViewShortURL= $curViewURL.Substring($ViewNamePos +1)
				$curViewInternalName= [IO.Path]::GetFileNameWithoutExtension($curViewShortURL)
				
				$schema = [XML]$curView.SchemaXml			
				$schema.View.Url = $curViewInternalName
				$schema.View.Name = $curViewDisplayName

				# Output -- Preparing View Details
				#----------------------------------------
				$viewDetails = New-Object -TypeName PSObject
				$viewDetails | Add-Member -Name 'ListName' -MemberType Noteproperty -Value $curListDisplayName
				$viewDetails | Add-Member -Name 'ViewName' -MemberType Noteproperty -Value $curViewDisplayName
				$viewDetails | Add-Member -Name 'ViewSchema' -MemberType Noteproperty -Value $schema.InnerXML
				
				# Output -- Append to Array
				#----------------------------------------
				if ($curViewInternalName -eq "AllItems")
				{
					# Append SchemaXml to UpdateView Node
					$UpdateViewArray += $viewDetails
				}
				else
				{
					# Append SchemaXml to AddView Node
					$AddViewArray += $viewDetails
				}
				
			}
		}
		Write-Host "$scriptName # Browsing Lists and View completed" -ForegroundColor Green `r
		
		
		$xmlFilePath  = $destFolderArtefacts + "\$fileNameForListsViews"

		Write-Host ""
		Write-Host "******************************************************************************" -ForegroundColor Magenta `r	
		Write-Host "$scriptName # About to create XML file." -ForegroundColor Magenta `r
		Write-Host "******************************************************************************" -ForegroundColor Magenta `r	

		# XML File -- Get an XMLTextWriter to create the XML
		Write-Debug "[$scriptName] Get an XMLTextWriter to create the XML"
		
		$xmlWriter = New-Object System.XMl.XmlTextWriter($xmlFilePath,$Null)
		
		# XML File -- Choose a pretty formatting
		Write-Debug "[$scriptName] Choose a pretty formatting"
		$xmlWriter.Formatting = 'Indented'
		$xmlWriter.Indentation = 1
		$XmlWriter.IndentChar = "`t"
		
		# XML File -- Write header
		Write-Debug "[$scriptName] Write Header"
		$xmlWriter.WriteStartDocument()
	
		# XML File -- Create root element "SPListViews" and add some attributes to it
		$xmlWriter.WriteStartElement('SPListViews')

		
		# XML File -- Building <AddView> node and child nodes
		Write-Host "... Building <AddView> node and child nodes"  -ForegroundColor Cyan `r
		$xmlWriter.WriteStartElement('AddView')
		
		$previousListName = ""
		$loopCount = 0
		
		foreach ($viewDetail in $AddViewArray)
		{
			Write-Debug "[$scriptName] AddViewArray// $($viewDetail.ListName) - $($viewDetail.ViewName)"
			
			if ($previousListName -ne $viewDetail.ListName)
			{
				if ($loopCount -ne 0)
				{
					#Closing Previous <List>
					$xmlWriter.WriteEndElement()
				}
				else
				{
					$loopCount = 1
				}
				
				#Openin new one
				$xmlWriter.WriteStartElement('List')
				$xmlWriter.WriteAttributeString("Title", $viewDetail.ListName);
				$previousListName = $viewDetail.ListName;			
			}

			Write-Debug "[$scriptName] AddViewArray// before WriteRaw"
			$xmlWriter.WriteRaw($viewDetail.ViewSchema)
		}
		#Closing Previous <List> 
		$xmlWriter.WriteEndElement()
		
		#Closing Previous <AddView> 
		$xmlWriter.WriteEndElement()
		Write-Debug "[$scriptName] <AddView> node and child nodes completed" 
		
		
		# XML File -- Building <UpdateView> node and child nodes
		Write-Host "... Building <UpdateView> node and child nodes"  -ForegroundColor Cyan `r
		$xmlWriter.WriteStartElement("UpdateView")
		$previousListName = ""
		$loopCount = 0
		
		foreach ($viewDetail in $UpdateViewArray)
		{
			Write-Debug "[$scriptName] UpdateViewArray// $($viewDetail.ListName) - $($viewDetail.ViewName)"
			
			if ($previousListName -ne $viewDetail.ListName)
			{
				if ($loopCount -ne 0)
				{
					#Closing Previous <List> ^
					$xmlWriter.WriteEndElement()
				}
				else
				{
					$loopCount = 1
				}	
				#Openin new one
				$xmlWriter.WriteStartElement('List')
				$xmlWriter.WriteAttributeString("Title", $viewDetail.ListName);
				$previousListName = $viewDetail.ListName;	
			}


			Write-Debug "[$scriptName] UpdateViewArray// before WriteRaw"
			$xmlWriter.WriteRaw($viewDetail.ViewSchema)
		}
		#Closing Previous <List> 
		$xmlWriter.WriteEndElement()
		
		#Closing Previous <UpdateView> 
		$xmlWriter.WriteEndElement()
		Write-Debug "[$scriptName] <UpdateView> node and child nodes completed" 
		
		# XML File -- Close the "SPListViews" node
		Write-Debug "[$scriptName] Close root element 'SPListViews'"
		$xmlWriter.WriteEndElement()
		
		# XML File -- Finalize the document 
		$xmlWriter.WriteEndDocument()
		$xmlWriter.Flush()
		$xmlWriter.Close()
		Write-Host "$scriptName # XML file successfully finalized."  -ForegroundColor Green `r

		Write-Host ""
		Write-Host "******************************************************************************" -ForegroundColor White `r	
		Write-Host "$scriptName # File location: `n$xmlFilePath"  -ForegroundColor White `r
		Write-Host "******************************************************************************" -ForegroundColor White `r	
		Write-Host ""
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

	Write-Host "******************************************************************************" -ForegroundColor Gray `r		
	Write-Host "$scriptName # Script ended." -ForegroundColor Gray `r
	Write-Host "******************************************************************************" -ForegroundColor Gray `r					
	
	#####################################################
	# Stopping SPAssignment and Transcript
	#####################################################
	Stop-SPAssignment -Global
	Stop-Transcript
}