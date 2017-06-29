##############################################################################################
#              
# NAME: SPHelpers.psm1 
# PURPOSE: Global and Generic functions related to SharePoint
#
# SOURCE : https://github.com/OhNotreDame/SPPS
#
##############################################################################################



<#
.SYNOPSIS
	Ensure the user userName on site siteURL and return its SPUser Object
	
.DESCRIPTION
	Ensure the user userName on site siteURL and return its SPUser Object, by computing the userName Windows Claism
	Return the SPUser object obtained by the SPWeb method EnsureUser()

.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER userName
	Username (domain\loginName) to check and ensure
	
.EXAMPLE
	ensureClaimsUser -siteURL <SiteURL> -userName <cid\userName>
	
.OUTPUTS
	SPUser

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 01.03.2017
	Last Updated by: JBO
	Last Updated: 01.03.2017
#>
function ensureClaimsUser 
{

	Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[ValidateScript({
			if ($_.StartsWith("domain\", $true, $null))
			{
				$true
			}
			else {
				throw "`r`n$_ is not a valid username.`nPlease use following format <domain\userName>."
			}
		})]
		[string]$userName
	)


	$functionName = $MyInvocation.MyCommand.Name	
	Write-Debug "Entering $functionName"
	$claimsEncodedUser = [string]::Empty;
	$spUser = $null;

	try
	{		
		
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if ($curWeb -ne $null)
		{
			if(-not[string]::IsNullOrEmpty($userName))				
			{
				$claim = New-SPClaimsPrincipal -Identity $userName -IdentityType WindowsSamAccountName
				$claimsEncodedUser = $claim.ToEncodedString();
				Write-Host "[$functionName] claimsEncodedUser: '$claimsEncodedUser'" -ForegroundColor DarkYellow `r
				$spUser = $curWeb.EnsureUser($claimsEncodedUser);
			}
			else
			{
				Write-Warning "[$functionName] 'UserName' is empty."
			}
		}
		else
		{
			Write-Warning "[$functionName] Site '$siteURL' does not exist." 
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
		Write-Debug "Exiting $functionName"
    }
	return $spUser;
}



<#
.SYNOPSIS
	Load the XML file into XML object
	
.DESCRIPTION
	Check whether the file path is valid or not.
	Then load the XML file into XML object.
		
.PARAMETER xmlPath
	Path of the XML file
					
.EXAMPLE
	LoadXMLFile -xmlPath <FilePathOfXMLFile>
	
.OUTPUTS
	XML Object containing the XML file

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 13.01.2017
	Last Updated by: JBO
	Last Updated: 13.01.2017
#>
function LoadXMLFile()
{
    [CmdletBinding()]
	param
    (
	    [Parameter(Mandatory=$true, Position=1)]
	    [string]$xmlPath
    ) 
    
	$functionName = $MyInvocation.MyCommand.Name
	Write-Debug "[$functionName] Entering function" 
	Write-Debug "[$functionName] Parameter/ FilePath: $xmlPath" 
	
	$xmlFile = $null
	
	try
	{		
		if(Test-Path $xmlPath)
		{			
			$xmlFile = [xml](Get-Content($xmlPath) -encoding utf8) 
		}
		else
		{
			Write-Warning  "[$functionName] File does not exist.`nPlease check the file path."
		}		
	}
	catch [Exception]
	{
		$xmlFile = $null
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
	
    return $xmlFile
}



<#
.SYNOPSIS
	Update the Regional Settings to set the specified culture
	
.DESCRIPTION
	Update the Regional Settings to set the specified culture
	
.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER cultureText
	Culture to be applied (Examples: en-GB, fr-FR, fr-CH, ...)
	
.EXAMPLE
	updateRegionalSettings -siteURL <siteURL> -cultureText <newCulture>
	
.OUTPUTS
	None
	
.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 13.01.2017
	Last Updated by: JBO
	Last Updated: 13.01.2017
#>
function updateRegionalSettings() 
{
	[CmdletBinding()]
	param
    (
	    [Parameter(Mandatory=$true, Position=1)]
	    [string]$siteURL,
        [Parameter(Mandatory=$true, Position=2)]
	    [string]$cultureText
    ) 
		
	$functionName = $MyInvocation.MyCommand.Name
	Write-Debug "[$functionName] Entering function" 
	Write-Debug "[$functionName] Parameter/ siteURL: $siteURL"
	Write-Debug "[$functionName] Parameter/ cultureText: $cultureText"

	try
	{
		$curWeb= GetSPWebBySiteUrl -siteURL $siteURL -EA SilentlyContinue 
		if ($curWeb -ne $null)
		{
			Write-Host "[$functionName] About to change culture to '$cultureText' on site '$siteURL'." -ForegroundColor Yellow `r
			$culture=[System.Globalization.CultureInfo]::CreateSpecificCulture($cultureText)
			$curWeb.Locale = $culture
			$curWeb.Update();
        	Write-Host "[$functionName] Site culture successfully changed to '$cultureText' on site '$siteURL'." -ForegroundColor Green `r
		}
		else
		{
			Write-Warning "[$functionName] Site '$siteURL' does not exist."
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
}