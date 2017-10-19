param
(
	[Parameter(Mandatory=$true, Position=1)]
	[string]$siteURL,
	[Parameter(Mandatory=$true, Position=2)]
	[string]$libraryName, 
	[Parameter(Mandatory=$true, Position=3)]
	[string]$filePath
)


<#
.SYNOPSIS
	Upload file to a SharePoint Library remotely without using WebClient/WebDav

.DESCRIPTION	
	Upload file to a SharePoint Library remotely without using WebClient/WebDav

.PARAMETER siteURL
	[Mandatory] URL of the SharePoint site
	
.PARAMETER libraryName
	[Mandatory] Internal Name of the document library

.PARAMETER filePath
	[Mandatory] Path of the file to be uploaded

.EXAMPLE
	uploadFileToSPLibrary.ps1 -siteURL <siteURL> -libraryName <libraryName> -filePath <filePath>
	
.OUTPUTS
	None
	
.LINK
	https://github.com/OhNotreDame/SPPS
	
.NOTES
	Created by: JBO
	Created: 21.08.2017
	Last Updated by: JBO
	Last Updated: 21.08.2017
	
#>

Clear-Host
Remove-Module *

$scriptName = $([System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Definition))
Write-Host "************************************************************************" -ForegroundColor Gray `r
Write-Host "$scriptName # Script started." -ForegroundColor Gray `r
Write-Host "************************************************************************" -ForegroundColor Gray `r		


try
{
	if(Test-Path $filePath)
	{
		$destination = $siteURL + "/" + $libraryName
		Write-Host "destination: $destination";
		
		# Get Current user
		$currentUser = [Environment]::UserName
		Write-Host "currentUser: $currentUser";
		
		$file = get-childitem $filePath
		Write-Host "fileName: $($file.FullName)";
		
		# Prepare the upload
		$webclient = New-Object System.Net.WebClient 
		
		###### OPTION 1:  PROMPT FOR CREDENTIALS ######
		$webclient.Credentials = Get-Credential $currentUser
		################################################
		
		###### OPTION 2: NO PROMPT / DefaultNetworkCredentials ######
		$credCache = New-Object System.Net.CredentialCache
		$credCache.Add($siteURL, "NTLM", [System.Net.CredentialCache]::DefaultNetworkCredentials)
		$webclient.Credentials = $credCache
		#############################################################
	
		# Upload the file
		$webclient.UploadFile($destination + "/" + $file.Name, "PUT", $file.FullName)
		
	}
	else
	{
		Write-Warning "File <$filePath> does not exist." 
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
	$scriptName = $([System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Definition))
	Write-Host "************************************************************************" -ForegroundColor Gray `r
	Write-Host "$scriptName # Script ended." -ForegroundColor Gray `r
	Write-Host "************************************************************************" -ForegroundColor Gray `r		
}
