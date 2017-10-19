<#

.SYNOPSIS
	List all SP Solutions deployed on farm and extract their WSP in the folder <destinationFolderPath>

.DESCRIPTION	
	List all SP Solutions deployed on farm and extract their WSP in the folder <destinationFolderPath>
	
.PARAMETER destinationFolderPath [Optional]
	Path of the desination folder

.EXAMPLE
	extractAndExportWSP.ps1 -destinationFolderPath C:\folder1\folder2
	extractAndExportWSP.ps1 

.OUTPUTS
	One XML file to recap all extracted WSP
	One WSP by solution deployed on SPFarm
	
.LINK
	https://github.com/OhNotreDame/SPPS
	
.NOTES
	Author: JBO
	Created: 02/10/2017
	Modified by: JBO
	Modified: 02/10/2017
	
#>
param
(
 [Parameter(Mandatory=$false, Position=1)]
 [string]$destinationFolderPath
)


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
$ModuleFolderPath = "D:\QuickDeployFW\Modules"



try
{
	if([string]::IsNullOrEmpty($destinationFolderPath)) 
	{
		$destinationFolderPath = Get-Location
		Write-Warning "[$functionName] Paramater destinationFolderPath is empty, will set it to the current location." 
		Write-Host "[$functionName] destinationFolderPath: $destinationFolderPath" -foregroundcolor Cyan
	}
	
	$destFolderLogs = "$destinationFolderPath\Logs"
	$destFolderWSP = "$destinationFolderPath\WSP_"+ $(get-date -format 'yyyyMMdd_HHmmss')

	$xmlFilePath  = $destFolderWSP + "\extractAndExportWSP_"+ $(get-date -format 'yyyyMMdd_HHmmss')+".xml"
	
	Write-Host "" `r
	Write-Host "******************************************************************************" -ForegroundColor Magenta `r
	Write-Host "$scriptName # Parameters & Settings" -ForegroundColor Magenta `r
	Write-Host "******************************************************************************" -ForegroundColor Magenta `r		
	Write-Host "ModuleFolderPath:" -ForegroundColor Gray `r
	Write-Host "$ModuleFolderPath" -ForegroundColor Gray `r
	Write-Host "******************************************************************************" -ForegroundColor Gray `r	
	Write-Host "scriptdir: "
	Write-Host "$scriptdir" -ForegroundColor Gray `r
	Write-Host "******************************************************************************" -ForegroundColor Gray `r
	Write-Host "destinationFolderPath:" -ForegroundColor Gray `r
	Write-Host "$destinationFolderPath" -ForegroundColor Gray `r
	Write-Host "******************************************************************************" -ForegroundColor Gray `r	
	Write-Host "scriptdir: "
	Write-Host "$scriptdir" -ForegroundColor Gray `r
	Write-Host "******************************************************************************" -ForegroundColor Magenta `r	
	
	Write-Host "" `r
	Write-Host "******************************************************************************" -ForegroundColor Magenta `r
	Write-Host "$scriptName # Other Settings" -ForegroundColor Magenta `r
	Write-Host "******************************************************************************" -ForegroundColor Magenta `r	
	Write-Host "destFolderLogs: "
	Write-Host "$destFolderLogs" -ForegroundColor Gray `r
	Write-Host "******************************************************************************" -ForegroundColor Gray `r	
	Write-Host "destFolderWSP: "
	Write-Host "$destFolderWSP" -ForegroundColor Gray `r
	Write-Host "******************************************************************************" -ForegroundColor Magenta `r	
	Write-Host "" `r
	
	
	###########################
	#### CREATING FOLDERS  ####
	###########################
	if(!(Test-Path $destFolderLogs))
	{
		New-Item $destFolderLogs -type Directory -Force | Out-Null
	}
	
	if(!(Test-Path $destFolderWSP))
	{
		New-Item $destFolderWSP -type Directory -Force | Out-Null
	}
	
	###########################
	#### TRANSCRIPT / LOGS ####
	###########################
	$logsFileName = $destFolderLogs + "\" + $scriptName  + "_"+ $(get-date -format 'yyyyMMdd_HHmmss') + ".log"
	Start-Transcript -path $logsFileName -noclobber -Append

	
	###########################
	####### OUTPUT FILE #######
	###########################
	$xmlWriter = New-Object System.XMl.XmlTextWriter($xmlFilePath,$Null)
	Write-Debug "[$scriptName] Choose a pretty formatting"
	# Choose a pretty formatting
	$xmlWriter.Formatting = 'Indented'
	$xmlWriter.Indentation = 1
	$XmlWriter.IndentChar = "`t"
	 
	Write-Debug "[$scriptName] Write header"
	# write header
	$xmlWriter.WriteStartDocument()
	# create root element "Workflows" and add some attributes to it
	$xmlWriter.WriteStartElement('SPSolutions')
	
	Write-Host "" `r
	Write-Host "About to export all SP solutions." -ForegroundColor Cyan `r	  
	foreach ($solution in Get-SPSolution)  
	{  
	

		$id = $Solution.SolutionID  
		$name = $Solution.Name  
		$title = $Solution.DisplayName  
		$version = $Solution.Version  
		$status = $Solution.Status  
		
		$containsGAC = $Solution.ContainsGlobalAssembly  
		$filename = $Solution.SolutionFile.Name 
		
		$deployed = $Solution.Deployed  
		$deployedWebApps = $Solution.DeployedWebApplications  
		$deployedWebAppsTXT = "";
		
		foreach ($webApp in	$deployedWebApps)
		{
			$newWebAppLine = $webApp.Name + ";"
			$deployedWebAppsTXT += $newWebAppLine
		}
		
		Write-Host "Exporting '$title' .." -ForegroundColor White `r	
		$solution.SolutionFile.SaveAs("$destFolderWSP\$filename") 		
		
		# Create the 'WSP' node with attributes
		$xmlWriter.WriteStartElement('WSP')
		$xmlWriter.WriteAttributeString("Name", $name);
		$xmlWriter.WriteAttributeString("DisplayName", $title);
		$xmlWriter.WriteAttributeString("SolutionID", $id);
		$xmlWriter.WriteAttributeString("ContainsGlobalAssembly", $containsGAC);
		$xmlWriter.WriteAttributeString("Version", $version);
		$xmlWriter.WriteAttributeString("Status", $status);
		$xmlWriter.WriteAttributeString("FileName", $fileName);
		$xmlWriter.WriteAttributeString("Deployed", $deployed);
		$xmlWriter.WriteAttributeString("DeployedWebApplications", $deployedWebAppsTXT);
		
		$xmlWriter.WriteEndElement()

	}
	
	Write-Debug "[$scriptName] Close root element 'SPSolutions'"
	# close the "SPSolutions" node
	$xmlWriter.WriteEndElement()
			 
	Write-Debug "[$scriptName] Finalize file"
	# finalize the document
	$xmlWriter.WriteEndDocument()
	$xmlWriter.Flush()
	$xmlWriter.Close()
	
	Write-Host "All SP solutions have been successfully exported." -ForegroundColor Green `r	  
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
