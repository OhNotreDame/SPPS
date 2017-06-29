<#
.SYNOPSIS
	Refresh all the fields of a list based on their Site Column definition
	
.DESCRIPTION
	Refresh all the fields of a list based on their Site Column definition
	
.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER listName
	Name of the list
	
.EXAMPLE
	refreshSiteColumsOnList.ps1 -siteURL <SiteURL> -listName <ListName> 
	
.OUTPUTS
	None

.LINK
	https://github.com/OhNotreDame/SPPS
	
.NOTES
	Created by: JBO
	Created: 28.04.2017
	Last Updated by: JBO
	Last Updated: 28.04.2017
#>


param
(
	[Parameter(Mandatory=$true, Position=1)]
	[String]$siteURL,
	[Parameter(Mandatory=$true, Position=2)]
	[String]$listName
)

Clear-Host
Remove-Module *

$scriptName = $([System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Definition))
Write-Host "************************************************************************" -ForegroundColor Gray `r
Write-Host "$scriptName # Script started." -ForegroundColor Gray `r
Write-Host "************************************************************************" -ForegroundColor Gray `r		

# Loading SharePoint Assembly and PS Snapin
Add-PsSnapin Microsoft.Sharepoint.PowerShell -EA silentlycontinue
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint") | Out-Null

# Starting SPAssignment
Start-SPAssignment -Global

try
{
	$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
	if($curWeb -ne $null)
	{	
		$curList = $curWeb.Lists.TryGetList($listName)
		if($curList -ne $null)
		{
			Write-Host "[$functionName] About to refresh List '$listName' from site '$siteURL'." -ForegroundColor Magenta  `r
			$myFieldsArray = New-Object System.Collections.ArrayList

			$curList.Fields  | ForEach-Object { 
				$myFieldsArray.Add($_.InternalName) > $null
			}

			[System.Threading.Thread]::CurrentThread.CurrentUICulture=$curWeb.UICulture;
			foreach ($fieldName in $myFieldsArray) {
				Write-Host "[$functionName] Refreshing column '$fieldName' on List" -ForegroundColor Cyan  `r
				$fieldToUpdate = $curList.Fields.GetFieldByInternalName($fieldName)
				if($fieldToUpdate -ne $null)
				{
					$parentWebField = $curWeb.Fields.GetFieldByInternalName($fieldName)
					if($parentWebField -ne $null)
					{
						#$parentWebField.SchemaXml
						$fieldToUpdate.SchemaXml = $parentWebField.SchemaXml;
						$fieldToUpdate.Update();
					}
					else
					{
						Write-Warning "[$functionName] Field '$fieldName' not found in SPWeb."
					}
				}
				else
				{
					Write-Warning "[$functionName] Field '$fieldName' not found in List."
				}
			}
			$curList.Update();

			Write-Host "[$scriptName] List '$listName' successfully refreshed from site '$siteURL'."  -ForegroundColor Green  `r
		}
		else
		{
			Write-Warning "[$scriptName] List '$listName' does not exist on site '$siteURL'."
		}
	}
	else
	{
		Write-Warning "[$scriptName] Site '$siteURL' not found."
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
}