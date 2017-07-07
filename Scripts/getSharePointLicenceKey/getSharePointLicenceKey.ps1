<#
.SYNOPSIS 
    Get the SharePoint License Key
.DESCRIPTION
    Get the SharePoint License Key  
	
.PARAMETER version
    Version of SharePoint
	
.EXAMPLE
    GetSharePointLicenceKey.ps1 -version 2013
  
.NOTES
    Author: System Center Automation Team
	Modified by: Julie Bonnard (JBON)
	Modified:  06/07/2017
	Schedule : None
#>
Param(
    [Parameter(Mandatory=$True,Position=1)]
    [string] $version
)



Clear-Host
Remove-Module *

#####################################################
# Loading SharePoint Assembly and PS Snapin
#####################################################
Add-PsSnapin Microsoft.Sharepoint.PowerShell -EA silentlycontinue
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint") | Out-Null

########################
### GLOBAL VARIABLES ###
########################
$scriptName = $([System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Definition))
$cmdpath = $MyInvocation.MyCommand.Path
Write-Host "************************************************************************" -ForegroundColor Gray `r
Write-Host "$scriptName # Script started on $(get-date -format 'dd/MM/yyyy HH:mm:ss')" -ForegroundColor Gray `r
Write-Host "************************************************************************" -ForegroundColor Gray `r		

try
{

	$map = "BCDFGHJKMPQRTVWXY2346789"
	$property = @{"2007"="12.0";"2010"="14.0";"2013"="15.0"}

	# Get Property
	$value = (get-itemproperty "HKLM:\SOFTWARE\Microsoft\Office\$($property[$version])\Registration\{90$(($property[$version] -replace '\.',''))000-110D-0000-1000-0000000FF1CE}").digitalproductid[0x34..0x42] 

	# Begin Parsing
	$ProductKey = "" 
	for ($i = 24; $i -ge 0; $i--) 
	{
		$r = 0
		for ($j = 14; $j -ge 0; $j--) {
			$r = ($r * 256) -bxor $value[$j]
			$value[$j] = [math]::Floor([double]($r/24))
			$r = $r % 24
		}
		$ProductKey = $map[$r] + $ProductKey
		if (($i % 5) -eq 0 -and $i -ne 0) {
			$ProductKey = "-" + $ProductKey
		}
	}
	
	Write-Host "ProductKey for SharePoint $version :"  $ProductKey -ForegroundColor Green `r

} 
catch 
{
	Write-Host "/!\ $scriptName An exception has been caught /!\ "  -ForegroundColor Red `r
	Write-Host "Type: " $_.Exception.GetType().FullName   -ForegroundColor Red `r
	Write-Host "Message: " $_.Exception.Message  -ForegroundColor Red `r
	Write-Host "Stacktrace: " $_.Exception.Stacktrace  -ForegroundColor Red `r
}
finally
{
	Write-Host "************************************************************************" -ForegroundColor Gray `r
	Write-Host "$scriptName # Script ended on $(get-date -format 'dd/MM/yyyy HH:mm:ss')" -ForegroundColor Gray `r
	Write-Host "************************************************************************" -ForegroundColor Gray `r	
}	
