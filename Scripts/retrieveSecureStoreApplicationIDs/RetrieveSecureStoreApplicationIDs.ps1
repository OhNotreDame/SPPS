<#
.FILENAME
	RetrieveSecureStoreApplicationIDs.ps1

.SYNOPSIS
	Retrieve all Secure Store Applications IDs

.LINK
	https://github.com/OhNotreDame/SPPS
	
.NOTES
	Author: Somebody from the Internet
	Modified by: Julie Bonnard (JBON)
	Modified:  07/07/2017
	Schedule : None

#>

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

$SecureStoreProvider=[Microsoft.Office.SecureStoreService.Server.SecureStoreProviderFactory]::Create()
Write-Host "Get Central Administration ... "  -ForegroundColor Cyan `r	
$site = Get-SPSite -Identity $(Get-SPWebApplication -IncludeCentralAdministration | ?{ $_.IsAdministrationWebApplication}).Url
Write-Host "Get Secure Store Service Context ... " -ForegroundColor Cyan `r	
$SecureStoreProvider.Context = Get-SPServiceContext -Site ($site)
$SecureStoreProvider.GetTargetApplications() |  ForEach-Object {
	Write-Host "Target Application Name: $($_.Name)" -ForegroundColor Green `r	
	Write-Host "Target Application ID: $($_.ID)" -ForegroundColor Green `r	
	try {
		$SecureStoreProvider.GetCredentials($_.ApplicationId) | ForEach-Object {
			$Credential = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($_.Credential))
			Write-Host "$($_.CredentialType): $($Credential)"
		}
	} catch  {
		Write-Host "`t$($_)"  -ForegroundColor yellow
	}
	Write-Host ""
}

	
# Stopping SPAssignment and Transcript
Stop-SPAssignment -Global

Write-Host "************************************************************************" -ForegroundColor Gray `r		
Write-Host "$scriptName # Script ended." -ForegroundColor Gray `r
Write-Host "************************************************************************" -ForegroundColor Gray `r					
