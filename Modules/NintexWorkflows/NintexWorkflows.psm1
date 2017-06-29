##############################################################################################
#              
# NAME: NintexWorkflows.psm1 
# PURPOSE: 
#	Manage Nintex Workflows (Export, Deploy, ...)
#
# SOURCE : https://github.com/OhNotreDame/SPPS
#
##############################################################################################


<#
.SYNOPSIS
	Export List Workflow calling NWAdmin.exe and save the NWF file into \NintexWorkflows\ folder
	
.DESCRIPTION
	Export List Workflow calling NWAdmin.exe and save the NWF file into \NintexWorkflows\ folder
	
.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER workflowName
	Name of the List Workflow to be exported
	
.PARAMETER listName
	Name of the list where the workflow is currently activated
	
.PARAMETER fileName
	Name of the expected workflow file
	
.PARAMETER destFolder
	(Optional) Location of the exported workflow
	If not specified, will get the current location 
		
.EXAMPLE
	ExportListWorkflow -siteURL <SiteURL> -workflowName <workflowName> -listName <listName> -fileName <fileName> [-destFolder <destFolderPath>]
	
.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 17.01.2017
	Last Updated by: JBO
	Last Updated: 19.01.2017
#>
function ExportListWorkflow()
{
	[CmdletBinding()]
	param
	(
	    [Parameter(Mandatory=$true, Position=1)]
	    [string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
	    [string]$workflowName,
		[Parameter(Mandatory=$true, Position=3)]
	    [string]$listName,
		[Parameter(Mandatory=$true, Position=4)]	
		[ValidateScript({
			if ($_ -imatch "\.(?-i:)(?-i:nwf)$") {
				$true
			}
			else {
				throw "`n$_ is not a valid filename for an Nintex Workflow File. `nFilename should end with '.nwf' (lowercase)."
			}
		})]
		[string]$fileName, 
		[Parameter(Mandatory=$false, Position=5)]
	    [string]$destFolder
	)
	
	$workflowType = "List"
	$functionName = $MyInvocation.MyCommand.Name
	Write-Debug "[$functionName] Entering function" 
	Write-Debug "[$functionName] Parameter / siteURL: $siteURL" 
	Write-Debug "[$functionName] Parameter / workflowName: $workflowName" 
	Write-Debug "[$functionName] Parameter / listName: $listName" 
	Write-Debug "[$functionName] Parameter / fileName: $fileName" 
	Write-Debug "[$functionName] Parameter / workflowType: $workflowType" 
	Write-Debug "[$functionName] Parameter / destFolder: $destFolder" 

	try
	{
		if([string]::IsNullOrEmpty($destFolder)) 
		{
			$destFolder = Get-Location
			Write-Warning "[$functionName] Paramater destFolder is empty, will set it to the current location." 
			Write-Host "[$functionName] destFolder: $destFolder" -foregroundcolor Cyan
		}
		
		$exportFolder = $destFolder + "\NintexWorkflows\"
		if (!(Test-Path $exportFolder))
		{
			New-Item $exportFolder -type Directory -Force | Out-Null
		}
		
		$curWeb= Get-SPWeb $siteURL -EA SilentlyContinue 
		if ($curWeb -ne $null)
		{
			$list=$curWeb.Lists.TryGetList($listName)
			if($list -ne $null)
			{
				Write-Host "[$functionName] About to export $workflowType Workflow '$workflowName' from List '$listName'." -ForegroundColor Magenta `r
				$finalFilePath = $exportFolder + "\" + $fileName
					
				$cmdNWAdmin = "NWAdmin.exe -o ExportWorkflow -siteURL '" + $siteURL + "' -workflowType '" + $workflowType + "' -workflowName '" + $workflowName + "' -list '" + $listName + "' -fileName '" + $finalFilePath + "'"
			
				Write-Host "$cmdNWAdmin" -ForegroundColor White `r
				$outputNWA = Invoke-Expression -Command:$cmdNWAdmin	
				if ($outputNWA -match "error")
				{
					throw [Exception] $outputNWA
				}
				else
				{
					Write-Host "[$functionName] $workflowType Workflow '$workflowName' successfully exported from List '$listName'." -ForegroundColor Green `r	
				}				
			}
			else
			{
				Write-Warning "[$functionName] List '$listName' not found on site '$siteURL'."	
			}
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
	Deploy List Workflow calling NWAdmin.exe from the \NintexWorkflows\ folder
	
.DESCRIPTION
	Deploy List Workflow calling NWAdmin.exe from the \NintexWorkflows\ folder
	
.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER workflowName
	Name of the List Workflow to be exported
	
.PARAMETER listName
	Name of the list where the workflow is currently activated
	
.PARAMETER fileName
	Name of the expected workflow file
	
.PARAMETER srcFolder
	(Optional) Location of the exported workflow
	If not specified, will get the current location
	
.PARAMETER overwrite
	(Optional) If exist on the target list, overwrite the workflow

.PARAMETER saveOnly
	(Optional) Save only the workflow on the target list

.PARAMETER saveIfCannotPublish
	(Optional) Save only the workflow if it cannot be published
	
.PARAMETER skipValidation
	(Optional) Skip the workflow validation
	
.EXAMPLE
	DeployListWorkflow -siteURL <SiteURL> -workflowName <workflowName> -listName <listName> -fileName <fileName> [-srcFolder <destFolderPath>] [-overwrite] [-saveOnly] [-saveIfCannotPublish] [-skipValidation]
	
.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 17.01.2017
	Last Updated by: JBO
	Last Updated: 19.01.2017
#>
function DeployListWorkflow()
{
	[CmdletBinding()]
	param
	(
	    [Parameter(Mandatory=$true, Position=1)]
	    [string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
	    [string]$workflowName,
		[Parameter(Mandatory=$true, Position=3)]
	    [string]$listName,
		[Parameter(Mandatory=$true, Position=4)]	
		[ValidateScript({
			if ($_ -imatch "\.(?-i:)(?-i:nwf)$") {
				$true
			}
			else {
				throw "`n$_ is not a valid filename for an Nintex Workflow File. `nFilename should end with '.nwf' (lowercase)."
			}
		})]
		[string]$fileName,
		[Parameter(Mandatory=$false, Position=5)]
	    [string]$srcFolder,
		[Parameter(Mandatory=$false, Position=6)]	
		[switch]$overwrite,
		[Parameter(Mandatory=$false, Position=7)]	
		[switch]$saveOnly,
		[Parameter(Mandatory=$false, Position=8)]	
		[switch]$saveIfCannotPublish,
		[Parameter(Mandatory=$false, Position=9)]	
		[switch]$skipValidation
	)
	
	$workflowType = "List"
	$functionName = $MyInvocation.MyCommand.Name
	Write-Debug "[$functionName] Entering function" 
	Write-Debug "[$functionName] Parameter / siteURL: $siteURL" 
	Write-Debug "[$functionName] Parameter / workflowName: $workflowName" 
	Write-Debug "[$functionName] Parameter / listName: $listName" 
	Write-Debug "[$functionName] Parameter / fileName: $fileName" 
	Write-Debug "[$functionName] Parameter / workflowType: $workflowType" 

	try
	{
		if([string]::IsNullOrEmpty($srcFolder)) 
		{
			$exportFolder = Get-Location
			Write-Warning "[$functionName] Paramater srcFolder is empty, will set it to the current location." 
		}
		else
		{
			$exportFolder = $srcFolder
		}
		Write-Host "[$functionName] exportFolder: $exportFolder" -foregroundcolor Cyan

		$additionalArgs = ""
		if ($overwrite)
		{
			$additionalArgs += "-overwrite "
		}
		
		if ($saveOnly)
		{
			$additionalArgs += "-saveOnly "
		}
		
		if ($saveIfCannotPublish)
		{
			$additionalArgs += "-saveIfCannotPublish "
		}
		
		if ($skipValidation)
		{
			$additionalArgs += "-skipValidation "
		}
		Write-Debug "[$functionName] additionalArgs: $additionalArgs"
		
		$curWeb= Get-SPWeb $siteURL -EA SilentlyContinue 
		if ($curWeb -ne $null)
		{
			$list=$curWeb.Lists.TryGetList($listName)
			if($list -ne $null)
			{
				Write-Host "[$functionName] About to deploy $workflowType Workflow '$workflowName' on List '$listName'." -ForegroundColor Magenta `r
				
				$finalFilePath = $exportFolder + "\" + $fileName
				if (Test-Path $finalFilePath)
				{
					$cmdNWAdmin = "NWAdmin.exe -o DeployWorkflow -siteURL '" + $siteURL + "' -workflowName '" + $workflowName + "' -targetList '"+ $listName + "' -nwfFile '" + $finalFilePath + "' " +  $additionalArgs
					Write-Host "$cmdNWAdmin" -ForegroundColor White `r
					$outputNWA = Invoke-Expression -Command:$cmdNWAdmin
					if ($outputNWA -match "error")
					{
						throw [Exception] $outputNWA
					}
					else
					{
						Write-Host "[$functionName] $workflowType Workflow '$workflowName' successfully deployed on List '$listName'." -ForegroundColor Green `r	
					}					
				}
				else
				{
					Write-Warning "[$functionName] File '$fileName' not found at '$finalFilePath'."	
				}	
			}
			else
			{
				Write-Warning "[$functionName] List '$listName' not found on site '$siteURL'."	
			}
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
	Export Reusable Workflow calling NWAdmin.exe and save the NWF file into \NintexWorkflows\ folder
	
.DESCRIPTION
	Export Reusable Workflow calling NWAdmin.exe and save the NWF file into \NintexWorkflows\ folder
	
.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER workflowName
	Name of the Workflow to be exported
	
.PARAMETER workflowType
	Type of the Workflow to be exported

.PARAMETER fileName
	Name of the expected workflow file

.PARAMETER srcFolder
	(Optional) Location of the exported workflow
	If not specified, will get the current location
	
.EXAMPLE
	ExportReusableWorkflow -siteURL <SiteURL> -workflowName <workflowName> -workflowType <workflowType> -fileName <fileName>  [-destFolder <destFolderPath>]
	
.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 17.01.2017
	Last Updated by: JBO
	Last Updated: 19.01.2017
#>
function ExportReusableWorkflow ()
{
	[CmdletBinding()]
	param
	(
	    [Parameter(Mandatory=$true, Position=1)]
	    [string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
	    [string]$workflowName,
		[Parameter(Mandatory=$true, Position=3)]
	    [ValidateSet("Reusable", "GloballyReusable", "Site", "Workflowapproval", "UserDefinedAction", IgnoreCase = $true)]
		[string]$workflowType,
		[Parameter(Mandatory=$true, Position=4)]	
		[ValidateScript({
			if ($_ -imatch "\.(?-i:)(?-i:nwf)$") {
				$true
			}
			else {
				throw "`n$_ is not a valid filename for an Nintex Workflow File. `nFilename should end with '.nwf' (lowercase)."
			}
		})]
		[string]$fileName,
		[Parameter(Mandatory=$false, Position=5)]
	    [string]$destFolder
	)
	
	$functionName = $MyInvocation.MyCommand.Name
	
	
	Write-Debug "[$functionName] Entering function" 
	Write-Debug "[$functionName] Parameter / siteURL: $siteURL" 
	Write-Debug "[$functionName] Parameter / workflowName: $workflowName" 
	Write-Debug "[$functionName] Parameter / fileName: $fileName" 
	Write-Debug "[$functionName] Parameter / workflowType: $workflowType" 
		
	try
	{
		if([string]::IsNullOrEmpty($destFolder)) 
		{
			$destFolder = Get-Location
			Write-Warning "[$functionName] Paramater srcFolder is empty, will set it to the current location." 
			Write-Host "[$functionName] srcFolder: $srcFolder" -foregroundcolor Cyan
		}
		
		$exportFolder = $destFolder + "\NintexWorkflows\"
		if (!(Test-Path $exportFolder))
		{
			New-Item $exportFolder -type Directory -Force | Out-Null
		}
		
		$curWeb= Get-SPWeb $siteURL -EA SilentlyContinue 
		if ($curWeb -ne $null)
		{
			Write-Host "[$functionName] About to export $workflowType Workflow '$workflowName' from site '$siteURL'." -ForegroundColor Magenta `r		
			$finalFilePath = $exportFolder + "\" + $fileName
			$cmdNWAdmin = "NWAdmin.exe -o ExportWorkflow -siteURL '" + $siteURL + "' -workflowType '" + $workflowType + "' -workflowName '" + $workflowName +"' -fileName '" + $finalFilePath + "'"
		
			Write-Host "$cmdNWAdmin" -ForegroundColor White `r
			$outputNWA = Invoke-Expression -Command:$cmdNWAdmin				
			if ($outputNWA -match "error")
			{
				throw [Exception] $outputNWA
			}
			else
			{
				Write-Host "[$functionName] $workflowType Workflow '$workflowName' successfully exported site '$siteURL'." -ForegroundColor Green `r	
			}			
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
	Deploy Reusable Workflow calling NWAdmin.exe from the \NintexWorkflows\ folder
	
.DESCRIPTION
	Deploy Reusable Workflow calling NWAdmin.exe from the \NintexWorkflows\ folder
	
.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER workflowName
	Name of the Workflow to be deployed
	
.PARAMETER workflowType
	Type of the Workflow to be deployed

.PARAMETER fileName
	Name of the expected workflow file
	
.PARAMETER srcFolder
	(Optional) Location of the exported workflow
	If not specified, will get the current location
	
.PARAMETER overwrite
	(Optional) If exist on the target list, overwrite the workflow

.PARAMETER saveOnly
	(Optional) Save only the workflow on the target list

.PARAMETER saveIfCannotPublish
	(Optional) Save only the workflow if it cannot be published
	
.PARAMETER skipValidation
	(Optional) Skip the workflow validation
	
.EXAMPLE
	DeployReusableWorkflow -siteURL <SiteURL> -workflowName <workflowName> -fileName <fileName> [-srcFolder <destFolderPath>] [-overwrite] [-saveOnly] [-saveIfCannotPublish] [-skipValidation]		
	
.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 17.01.2017
	Last Updated by: JBO
	Last Updated: 19.01.2017
#>
function DeployReusableWorkflow ()
{
	[CmdletBinding()]
	param
	(
	    [Parameter(Mandatory=$true, Position=1)]
	    [string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
	    [string]$workflowName,
		[Parameter(Mandatory=$true, Position=3)]	
		[ValidateScript({
			if ($_ -imatch "\.(?-i:)(?-i:nwf)$") {
				$true
			}
			else {
				throw "`n$_ is not a valid filename for an Nintex Workflow File. `nFilename should end with '.nwf' (lowercase)."
			}
		})]
		[string]$fileName,
		[Parameter(Mandatory=$false, Position=5)]
	    [string]$srcFolder,
		[Parameter(Mandatory=$false, Position=6)]	
		[switch]$overwrite,
		[Parameter(Mandatory=$false, Position=7)]	
		[switch]$saveOnly,
		[Parameter(Mandatory=$false, Position=8)]	
		[switch]$saveIfCannotPublish,
		[Parameter(Mandatory=$false, Position=9)]	
		[switch]$skipValidation
	)

	$workflowType = "Reusable"
	$functionName = $MyInvocation.MyCommand.Name
	Write-Debug "[$functionName] Entering function" 
	Write-Debug "[$functionName] Parameter / siteURL: $siteURL" 
	Write-Debug "[$functionName] Parameter / workflowName: $workflowName" 
	Write-Debug "[$functionName] Parameter / fileName: $fileName" 

	try
	{
		if([string]::IsNullOrEmpty($srcFolder)) 
		{
			$exportFolder = Get-Location
			Write-Warning "[$functionName] Paramater srcFolder is empty, will set it to the current location." 
		}
		else
		{
			$exportFolder = $srcFolder
		}
		Write-Host "[$functionName] exportFolder: $exportFolder" -foregroundcolor Cyan
	
		$additionalArgs = ""
		if ($overwrite)
		{
			$additionalArgs += "-overwrite "
		}
		
		if ($saveOnly)
		{
			$additionalArgs += "-saveOnly "
		}
		
		if ($saveIfCannotPublish)
		{
			$additionalArgs += "-saveIfCannotPublish "
		}
		
		if ($skipValidation)
		{
			$additionalArgs += "-skipValidation "
		}
		Write-Debug "[$functionName] additionalArgs: $additionalArgs"
		
		
		$curWeb= Get-SPWeb $siteURL -EA SilentlyContinue 
		if ($curWeb -ne $null)
		{
			Write-Host "[$functionName] About to deploy $workflowType Workflow '$workflowName' on site '$siteURL'." -ForegroundColor Magenta `r
			$finalFilePath = $exportFolder + "\" + $fileName
			if (Test-Path $finalFilePath)
			{
				#$outputNWA =  [string] (NWAdmin.exe -o DeployWorkflow -siteURL $siteURL -workflowName $workflowName -nwfFile $fileName)
				$cmdNWAdmin = "NWAdmin.exe -o DeployWorkflow -siteURL '" + $siteURL + "' -workflowName '" + $workflowName + "' -nwfFile '" + $finalFilePath + "' " +  $additionalArgs
				Write-Host "$cmdNWAdmin" -ForegroundColor White `r
				$outputNWA = Invoke-Expression -Command:$cmdNWAdmin				
				if ($outputNWA -match "error")
				{
					throw [Exception] $outputNWA
				}
				else
				{
					Write-Host "[$functionName] $workflowType Workflow '$workflowName' successfully deploy on site '$siteURL'." -ForegroundColor Green `r	
				}				
			}
			else
			{
				Write-Warning "[$functionName] File '$fileName' not found at '$finalFilePath'."	
			}		
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
	Export All Workflows From Site $siteURL and generate and XML file with all WF settings
	
.DESCRIPTION
	Export All Workflows From Site $siteURL and generate and XML file with all WF settings
	
.PARAMETER siteUrl
	URL of the SharePoint Site
		
.EXAMPLE
	ExportAllWorkflowsFromSite -siteURL <SiteURL>
	
.OUTPUTS
	One XML file containing the list of all the Workflows on site $siteURL and their settings (Name, Type, List (if it's a List WF) and FileName)
	One folder NintexWorkflows\ to store all workflows files
	One .NWF file by Workflow

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 17.01.2017
	Last Updated by: JBO
	Last Updated: 19.01.2017
#>
function ExportAllWorkflowsFromSite()
{
	[CmdletBinding()]
	param
	(
	    [Parameter(Mandatory=$true, Position=1)]
	    [string]$siteURL, 
		[Parameter(Mandatory=$false, Position=1)]
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
			Write-Host "[$functionName] destFolder: $destFolder"
		}
		
		$exportFolder = $destFolder + "\NintexWorkflows\"
		if (!(Test-Path $exportFolder))
		{
			New-Item $exportFolder -type Directory -Force | Out-Null
		}
		
		Write-Debug "[$functionName] Before creating Export folder"
		$exportFolder = $destFolder + "\NintexWorkflows\"
		if (!(Test-Path $exportFolder))
		{
			New-Item $exportFolder -type Directory -Force | Out-Null
		}
		
		$artefactFolder = $destFolder + "\Artefacts\"
		if (!(Test-Path $artefactFolder))
		{
			New-Item $artefactFolder -type Directory -Force | Out-Null
		}

		$curDir = Get-Location
		#Loading Nintex DLLs		
		[void][System.Reflection.Assembly]::LoadWithPartialName("Nintex.Workflow") 
		[void][System.Reflection.Assembly]::LoadWithPartialName("Nintex.Workflow.SupportConsole")
		[void][System.Reflection.Assembly]::LoadWithPartialName("Nintex.Workflow.Administration")
		
		$curWeb= Get-SPWeb $siteURL -EA SilentlyContinue 
		if ($curWeb -ne $null)
		{
			$webID = $curWeb.ID

			# Grab Nintex Config database name
			$CFGDB = [Nintex.Workflow.Administration.ConfigurationDatabase]::OpenConfigDataBase().Database
			$cmd = New-Object -TypeName System.Data.SqlClient.SqlCommand
			$cmd.CommandType = [System.Data.CommandType]::Text

			# Begin SQL Query 
			$cmd.CommandText = "SELECT  PW.[WorkflowId],
										PW.[WorkflowName],
										PW.[WorkflowType],
										PW.[SiteID],
										PW.[WebID],
										PW.[ListID]
								FROM [$CFGDB].dbo.PublishedWorkflows PW
								INNER JOIN (SELECT [WorkflowId],MAX([Version]) AS MaxVersion
											FROM [$CFGDB].dbo.PublishedWorkflows
											GROUP BY WorkflowId) PWMax
									ON PWMax.[WorkflowId]=PW.[WorkflowId]
									AND PWMax.MaxVersion=PW.[Version]
								WHERE PW.[WebID] = '" + $webID +"' ;"
			
			
			#Preparing XML Output and File Location
			Write-Debug "[$functionName] Preparing XML Output and File Location"
			$webTitleTrim = $curWeb.Title -replace " ", ""
			#$xmlFilePath = $exportFolder +"NintexWorkflows_"+ $webTitleTrim + ".xml"
		
			$xmlFilePath  = $artefactFolder + "\NintexWorkflows.xml"
			
			Write-Host "[$functionName] xmlFilePath: $xmlFilePath"
			Write-Debug "[$functionName] Get an XMLTextWriter to create the XML"
			# get an XMLTextWriter to create the XML
			$xmlWriter = New-Object System.XMl.XmlTextWriter($xmlFilePath,$Null)
			
			Write-Debug "[$functionName] Choose a pretty formatting"
			# Choose a pretty formatting
			$xmlWriter.Formatting = 'Indented'
			$xmlWriter.Indentation = 1
			$XmlWriter.IndentChar = "`t"
			 
			Write-Debug "[$functionName] Write header"
			# write header
			$xmlWriter.WriteStartDocument()

			Write-Debug "[$functionName] Create root element 'Workflows'"
			# create root element "Workflows" and add some attributes to it
			$xmlWriter.WriteStartElement('Workflows')
				
			foreach ($database in [Nintex.Workflow.Administration.ConfigurationDatabase]::GetConfigurationDatabase().ContentDatabases)
			{
				$reader = $database.ExecuteReader($cmd)
				while($reader.Read())
				{					
					$wfName = $reader['WorkflowName']
					$wfType = $reader['WorkflowType']
					Write-Host "[$functionName] Workflow: $wfName - Type:$wfType"
					Write-Debug "[$functionName] Workflow: $wfName - Type:$wfType - List:$($reader['ListID'])"
								
					# Create the 'Workflow' node with attributes
					$xmlWriter.WriteStartElement('Workflow')
					$xmlWriter.WriteAttributeString("Name", $reader["WorkflowName"]);
					$xmlWriter.WriteAttributeString("Type", $reader["WorkflowType"]);
					
					# If it's a List Workflow, add the List Attribute to the XML
					if($reader["ListID"] -ne "00000000-0000-0000-0000-000000000000") {
						$listID = [Guid] $reader["ListID"];
						$list = $curWeb.Lists[$listID]
						$listName = $list.Title;
						$xmlWriter.WriteAttributeString("List", $listName);
					}
					
					# Build fileName prior to export
					 $fileName = ($wfName -replace  " ", "" ) + ".nwf"
					# $fileName = $exportFolder + $fileName
					
					# Add the FileName Attribute to the XML
					$xmlWriter.WriteAttributeString("FileName", $fileName );
					
					# Close the 'Workflow' node
					$xmlWriter.WriteEndElement()
										
					# Export Workflow to file
					if ($wfType -eq "List")
					{ 
						ExportListWorkflow -siteURL $siteURL -workflowName $wfName -listName $listName -fileName $fileName -destFolder $destFolder
					}
					else
					{
						ExportReusableWorkflow -siteURL $siteURL -workflowName $wfName -workflowType $wfType -fileName $fileName -destFolder $destFolder
					}
				}#end while($reader.Read())
			}#end foreach ($database)

			Write-Debug "[$functionName] Close root element 'Workflows'"
			# close the "Workflows" node
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
	Browse and parse the workflowDescriptionXML XML object
	
.DESCRIPTION
	Browse and parse the file listDescriptionXML
		For each node, Deploy the Workflow	
	
.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER workflowDescriptionXML
	XML Object representing the workflow to be deployed
	
.EXAMPLE
	browseAndDeployWorkflowsXML -siteURL <SiteURL> -workflowDescriptionXML <XMLObjectToParse>
	
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
function browseAndDeployWorkflowsXML()
{
	[CmdletBinding()]
	param
    (
        [Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[XML]$workflowDescriptionXML, 
		[Parameter(Mandatory=$false, Position=3)]
	    [string]$srcFolder
	)

    try
    {
		$functionName = $MyInvocation.MyCommand.Name
		Write-Host "[$functionName] Entering function" 
		Write-Host "[$functionName] Parameter / siteURL: $siteURL"
	
		if([string]::IsNullOrEmpty($srcFolder)) 
		{
			$srcFolder = Get-Location
			Write-Warning "[$functionName] Paramater srcFolder is empty, will set it to the current location." 
			Write-Host "[$functionName] srcFolder: $srcFolder" -foregroundcolor Cyan
		}
		
		if($workflowDescriptionXML.HasChildNodes)
		{
			$wfLists = $workflowDescriptionXML.SelectNodes("/Workflows")
			foreach($wf in $wfLists.Workflow)
			{
				Write-Debug "[$functionName] Name: $($wf.Name)"
				Write-Debug "[$functionName] Type: $($wf.Type)"
				
				if ($wf.Type -eq "List")
				{
					DeployListWorkflow -siteURL $siteURL -workflowName $wf.Name -listName $wf.List -fileName $wf.FileName -srcFolder $srcFolder -overwrite
				}
				else
				{
					DeployReusableWorkflow -siteURL $siteURL -workflowName $wf.Name -fileName $wf.FileName -srcFolder $srcFolder -overwrite
				}

			}#end foreach
		}
		else
		{
			Write-Warning "[$functionName] List XML defintion file is empty."
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
		Write-Debug "[$functionName] Exiting function" 
	} 
}