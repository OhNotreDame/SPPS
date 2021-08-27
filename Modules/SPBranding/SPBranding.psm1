##############################################################################################
#              
# NAME: SPBranding.psm1 
# PURPOSE: 
#	Manage Branding on a SharePoint site
#
##############################################################################################


<#
.SYNOPSIS
	Set the master page on site
	
.DESCRIPTION
	Set the master page on site
		
.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER masterPageFilePath
	Location/Path (Server Relative) of the masterpage to be applied
	
.EXAMPLE
	setMasterPage -siteURL <SiteURL> -masterPageFilePath <masterPageFilePath>
	
.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 04.04.2017
	Last Updated by: JBO
	Last Updated: 04.04.2017
#>
function setMasterPage()
{
	Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[string]$masterPageFilePath
	)
    
	$functionName = $MyInvocation.MyCommand.Name
    Write-Debug "[$functionName] Entering function" 
    Write-Debug "[$functionName] Parameter / siteURL: $siteURL" 
    Write-Debug "[$functionName] Parameter / masterPageFilePath: $masterPageFilePath" 
	
    try
    {
        $curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
        if($curWeb -ne $null)
        {     
			if ($curWeb.isRootWeb)
			{
				$customMasterFile = $curWeb.GetFile($masterPageFilePath)
			}
			else
			{
				$customMasterFile = $curWeb.site.RootWeb.GetFile($masterPageFilePath)
			}

            if($customMasterFile -ne $null -and $customMasterFile.Length -gt 0)
            {
                Write-Host "[$functionName] About to apply Masterpage on site '$siteURL'." -ForegroundColor Magenta `r  
				$curWeb.AllowUnsafeUpdates  = $true;
                $curWeb.MasterUrl = $masterPageFilePath
                $curWeb.CustomMasterUrl = $masterPageFilePath
                $curWeb.Update();
                $curWeb.AllowUnsafeUpdates  = $false;                
				Write-Host "[$functionName] Masterpage successfully updated on '$siteURL'." -ForegroundColor Green `r
            }
            else
            {
                Write-Warning "[$functionName] Masterpage '$masterPageFilePath' not found on '$siteURL'." 
            }
        }
        else
        {
            Write-Warning "[$functionName] Site '$siteURL' not found." 
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
}


<#
.SYNOPSIS
	Set the master page on site
	
.DESCRIPTION
	Set the master page on site
		
.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER logoFilePath
	Location/Path (Server Relative) of the logo to be applied
	
.EXAMPLE
	setLogo -siteURL <SiteURL> -logoFilePath <logoFilePath>
	
.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 04.04.2017
	Last Updated by: JBO
	Last Updated: 04.04.2017
#>
function setLogo()
{
	Param
	(   [Parameter(Mandatory=$true, Position=1)]
	    [string]$siteURL,
        [Parameter(Mandatory=$true, Position=2)]
	    [string]$logoFilePath
	)

	$functionName = $MyInvocation.MyCommand.Name
    Write-Debug "[$functionName] Entering function" 
    Write-Debug "[$functionName] Parameter / siteURL: $siteURL" 
    Write-Debug "[$functionName] Parameter / masterPageFilePath: $masterPageFilePath" 
	
    try
    {
        $curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
        if($curWeb -ne $null)
        {     
			if ($curWeb.isRootWeb)
			{
				$logoFile = $curWeb.GetFile($logoFilePath)
			}
			else
			{
				$logoFile =  $curWeb.site.RootWeb.GetFile($logoFilePath)
			}

            if($logoFile -ne $null -and $logoFile.Length -gt 0)
            {
                Write-Host "$functionName About to apply Logo on site $siteName." -ForegroundColor White `r  
				$curWeb.AllowUnsafeUpdates  = $true;
                $curWeb.SiteLogoUrl = $logoFile.ServerRelativeUrl
                $curWeb.Update()
                $curWeb.AllowUnsafeUpdates  = $false;

                Write-Host "$functionName Logo applied on '$siteURL'." -ForegroundColor Green `r
            }
            else
            {
                Write-Warning "[$functionName] Logo not found on '$siteURL'."
            }
        }
        else
        {
            Write-Warning "[$functionName] Site '$siteURL' not found." 
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
}

