##############################################################################################
# 
# NAME: SPPermissionsLevels.psm1  
# PURPOSE: Manage Role definitions in SharePoint site.
#	See SPPermissionsLevels.xml for Schema
#
# SOURCE : https://github.com/OhNotreDame/SPPS
#
##############################################################################################
<#
.SYNOPSIS
	Parse the file SPPermissionsLevels XML object and initiate Permissions Level customization
	
.DESCRIPTION
	Will parse the file SPPermissionsLevels XML object making the difference between the permissions level to add (AddPermissionsLevel node),
	to change (ChangePermissionsLevel node) and the permissions to be deleted (DeletePermissionsLevel node)

.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER permissionsLevelsXML
	XML object of Permissions Level to manage.
	
.EXAMPLE
	browseAndParseSPPermissionsLevelXML -siteURL <SiteURL> -permissionsLevelsXML <permissionsLevelsXML>
	
.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 01.03.2017
	Last Updated by: JBO
	Last Updated: 01.03.2017
#>
function browseAndParseSPPermissionsLevelXML()
{
    Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[xml]$permissionsLevelsXML
	)

	$functionName = $MyInvocation.MyCommand.Name	
	Write-Debug "Entering $functionName"
	
    try
    {
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if ($curWeb -ne $null)
		{
			$SPPermissionsLevels = $permissionsLevelsXML.SelectNodes("/SPPermissionsLevels")
			if($SPPermissionsLevels -ne $null -and $SPPermissionsLevels.HasChildNodes)
			{	
				#--- Add Permissions (SPPermissionsLevels.AddPermissionsLevel node)			
				Write-Host "`n[$functionName] About to call browseAndCreateListPermissionsXML()."-ForegroundColor Cyan `r
				$addPermissionLevelXML =  $SPPermissionsLevels.AddPermissionsLevel
				if($addPermissionLevelXML -ne $null -and $addPermissionLevelXML.HasChildNodes)
				{									
					browseAndCreatePermissionsLevelsXML -siteURL $siteURL -addPermissionsLevelXML $addPermissionLevelXML
				}
				else
				{
					Write-Warning "[$functionName] 'AddPermissionsLevel' node is empty."
				}

				#--- Change Permissions (SPPermissionsLevels.ChangePermissionsLevel node)
				Write-Host "`n[$functionName] About to call browseAndChangePermissionsLevelsXML()."-ForegroundColor Cyan `r
				$changePermissionLevelXML =  $SPPermissionsLevels.ChangePermissionsLevel
				if($changePermissionLevelXML -ne $null -and $changePermissionLevelXML.HasChildNodes)
				{									
					browseAndChangePermissionsLevelsXML -siteURL $siteURL -changePermissionsLevelXML $changePermissionLevelXML
				}
				else
				{
					Write-Warning "`n[$functionName] 'ChangePermissionsLevel' node is empty." 
				}

				#--- Delete Permissions (SPPermissionsLevels.DeletePermissionsLevel node)
				Write-Host "`n[$functionName] About to call browseAndDeletePermissionsLevelsXML()."-ForegroundColor Cyan `r
				$deletePermissionLevelXML =  $SPPermissionsLevels.DeletePermissionsLevel
				if($deletePermissionLevelXML -ne $null -and $deletePermissionLevelXML.HasChildNodes)
				{									
					browseAndDeletePermissionsLevelsXML -siteURL $siteURL -deletePermissionsLevelXML $deletePermissionLevelXML
				}
				else
				{
					Write-Warning "[$functionName] 'DeletePermissionsLevel' node is empty."
				}
			
			}
			else
			{
				Write-Warning "[$functionName] 'SPPermissionsLevels' node is empty."
			}	   
		}
		else
		{
			Write-Warning "[$functionName] Site '$siteURL' not found."
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
		Write-Debug "Exiting $functionName"
    }
}


<#
.SYNOPSIS
	Parse the node AddPermission and initiate Permissions Level creation
	
.DESCRIPTION
	Parse the node AddPermission and initiate Permissions Level creation

.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER addPermissionsLevelXML
	XML object of Permissions Level to add.
	
.EXAMPLE
	browseAndCreatePermissionsLevelsXML -siteURL <SiteURL> -addPermissionsLevelXML <addPermissionXML>

.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 01.03.2017
	Last Updated by: JBO
	Last Updated: 01.03.2017
#>
function browseAndCreatePermissionsLevelsXML()
{
	Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[XML.XmlElement]$addPermissionsLevelXML
	)

	$functionName = $MyInvocation.MyCommand.Name	
	Write-Debug "Entering $functionName"
	
	try
	{				
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if ($curWeb -ne $null)
		{			
			if($addPermissionsLevelXML -ne $null -and $addPermissionsLevelXML.HasChildNodes)
			{
				foreach($plToCreate in $addPermissionsLevelXML.PermissionsLevel)
				{
					$plName = $plToCreate.Name
					$plBasePermissions = $plToCreate.BasePermissions
					if(-not[string]::IsNullOrEmpty($plName.Trim()) -and -not[string]::IsNullOrEmpty($plBasePermissions))
					{	
						createPermissionsLevel -siteURL $siteURL -plName $plName -plBasePermissions $plBasePermissions -plDescription $plToCreate.Description 
					}
					else
					{
						Write-Warning "[$functionName] Permission Level Name or Base Permissions could not be empty." 
					}
				} #foreach
			}
			else
			{
				Write-Warning "[$functionName] 'addPermissionsLevelXML' XML node is empty."
			}
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
		Write-Debug "Exiting $functionName"
    }
}


<#
.SYNOPSIS
	Parse the node changePermissionsLevels and initiate Permissions Level updates (description and basepermissions)
	
.DESCRIPTION
	Parse the node changePermissionsLevels and initiate Permissions Level updates (description and basepermissions)

.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER changePermissionsLevelXML
	XML object of Permissions Level to update.
	
.EXAMPLE
	browseAndChangePermissionsLevelsXML -siteURL <SiteURL> -changePermissionsLevelXML <changePermissionsLevelXML>

.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 01.03.2017
	Last Updated by: JBO
	Last Updated: 01.03.2017
#>
function browseAndChangePermissionsLevelsXML()
{
	Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[XML.XmlElement]$changePermissionsLevelXML
	)

	$functionName = $MyInvocation.MyCommand.Name	
	Write-Debug "Entering $functionName"
	
	try
	{				
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if ($curWeb -ne $null)
		{			
			if($changePermissionsLevelXML -ne $null -and $changePermissionsLevelXML.HasChildNodes)
			{
				foreach($plToChange in $changePermissionsLevelXML.PermissionsLevel)
				{
					$plName = $plToChange.Name
					$plBasePermissions = $plToChange.BasePermissions
					$plDescription = $plToChange.Description 
					if(-not[string]::IsNullOrEmpty($plName.Trim()) -and -not[string]::IsNullOrEmpty($plBasePermissions))
					{	
						changePermissionsLevel -siteURL $siteURL -plName $plName -plBasePermissions $plBasePermissions -plDescription $plDescription 
					}
					else
					{
						Write-Warning "[$functionName] Permission Level Name or Base Permissions could not be empty." 
					}
				} #foreach
			}
			else
			{
				Write-Warning "[$functionName] 'changePermissionsLevelXML' XML node is empty."
			}
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
		Write-Debug "Exiting $functionName"
    }
}



<#
.SYNOPSIS
	Parse the node DeletePermissionsLevels and initiate Permissions Level deletion
	
.DESCRIPTION
	Parse the node DeletePermissionsLevels and initiate Permissions Level deletion

.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER deletePermissionsLevelXML
	XML object of Permissions Level to delete.
	
.EXAMPLE
	browseAndDeletePermissionsLevelsXML -siteURL <SiteURL> -deletePermissionsLevelXML <deletePermissionsLevelXML>

.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 01.03.2017
	Last Updated by: JBO
	Last Updated: 01.03.2017
#>
function browseAndDeletePermissionsLevelsXML()
{
	Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[XML.XmlElement]$deletePermissionsLevelXML
	)

	$functionName = $MyInvocation.MyCommand.Name	
	Write-Debug "Entering $functionName"
	
	try
	{				
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if ($curWeb -ne $null)
		{			
			if($deletePermissionsLevelXML -ne $null -and $deletePermissionsLevelXML.HasChildNodes)
			{
				foreach($plToDelete in $deletePermissionsLevelXML.PermissionsLevel)
				{
					$plName = $plToDelete.Name
					if(-not[string]::IsNullOrEmpty($plName.Trim())) 
					{	
						RemovePermissionsLevel -siteURL $siteURL -plName $plName
					}
					else
					{
						Write-Warning "[$functionName] Permission Level Name could not be empty." 
					}
				} #foreach
			}
			else
			{
				Write-Warning "[$functionName] 'deletePermissionsLevelXML' XML node is empty."
			}
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
		Write-Debug "Exiting $functionName"
    }
}


<#
.SYNOPSIS
	Create a specific Permissions Level on a Site 
	
.DESCRIPTION
	Create a specific Permissions Level on a Site 

.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER plName
	Name of the Permission Level
	
.PARAMETER plBasePermissions
	Base Permissions of the Permission Level
	
.PARAMETER plDescription
	(Optional) Description of the Permission Level

.EXAMPLE
	createPermissionsLevel -siteURL <SiteURL> -plName <plName> -plBasePermissions <plBasePermissions> -plDescription <plDescription>
	
.OUTPUTS
	None
	
.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 01.03.2017
	Last Updated by: JBO
	Last Updated: 01.03.2017
#>
function createPermissionsLevel()
{
    Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[string]$plName,
		[Parameter(Mandatory=$true, Position=3)]
		[string]$plBasePermissions,
		[Parameter(Mandatory=$false, Position=4)]
		[string]$plDescription
	)
	
	$functionName = $MyInvocation.MyCommand.Name	
	Write-Debug "Entering $functionName"
	
    try
    {
	
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if($curWeb -ne $null)
		{
			if(-not[string]::IsNullOrEmpty($plName.Trim()) -and -not[string]::IsNullOrEmpty($plBasePermissions))
			{
				if($curWeb.RoleDefinitions[$plName] -eq $null)
				{
					Write-Host "[$functionName] About to create Permission Level '$plName' on site '$siteURL'." -foregroundcolor Magenta `r
					$spRoleDefinition = New-Object Microsoft.SharePoint.SPRoleDefinition  
					$spRoleDefinition.Name = $plName
					$spRoleDefinition.Description = $plDescription
					$spRoleDefinition.BasePermissions = $plBasePermissions
					$curWeb.RoleDefinitions.Add($spRoleDefinition)
					Write-Host "[$functionName] Permission Level '$plName' successfully created on site '$siteURL'." -foregroundcolor Green `r
				}
				else
				{
					Write-Warning "[$functionName] Permission Level '$plName' already exists on site '$siteURL'."
				}
			}
			else
			{
				Write-Warning "[$functionName] Name or Base Permission could not be empty." 
			}
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
    }
}


<#
.SYNOPSIS
	Change a specific Permissions Level on a Site 
	
.DESCRIPTION
	Change a specific Permissions Level on a Site 

.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER plName
	Name of the Permission Level
	
.PARAMETER basePermissions
	Base Permissions of the Permission Level
	
.PARAMETER plDescription
	(Optional) Description of the Permission Level

.EXAMPLE
	changePermissionsLevel -siteURL <SiteURL> -plName <plName> -basePermissions <basePermissions> -plDescription <plDescription>
	
.OUTPUTS
	None
	
.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 01.03.2017
	Last Updated by: JBO
	Last Updated: 01.03.2017
#>
function changePermissionsLevel()
{
    Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[string]$plName,
		[Parameter(Mandatory=$true, Position=3)]
		[string]$plBasePermissions,
		[Parameter(Mandatory=$false, Position=4)]
		[string]$plDescription
	)
	
	$functionName = $MyInvocation.MyCommand.Name	
	Write-Debug "Entering $functionName"
	
    try
    {
	
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if($curWeb -ne $null)
		{
			if(-not[string]::IsNullOrEmpty($plName.Trim()) -and -not[string]::IsNullOrEmpty($plBasePermissions))
			{
			
				$plToChange = $curWeb.RoleDefinitions[$plName]
				if($plToChange -ne $null)
				{
					Write-Host "[$functionName] About to change Permission Level '$plName' on site '$siteURL'." -foregroundcolor Magenta `r
					$plToChange.Description = $plDescription
					$plToChange.BasePermissions = $plBasePermissions
					$plToChange.Update()
					Write-Host "[$functionName] Permission Level '$plName' successfully created on site '$siteURL'." -foregroundcolor Green `r
				}
				else
				{
					Write-Warning "[$functionName] Permission Level '$plName' does not exist on site '$siteURL'."
				}
			}
			else
			{
				Write-Warning "[$functionName] Name or Base Permission could not be empty." 
			}
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
    }
}





<#
.SYNOPSIS
	Delete a specific Permission Level from a Site (without checking if it's already in use)
	
.DESCRIPTION
	Delete a specific Permission Level from a Site (without checking if it's already in use)

.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER plName
	Name of the Permission Level

.EXAMPLE
	removePermissionsLevel -siteURL <SiteURL> -plName <plName>
	
.OUTPUTS
	None
	
.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 01.03.2017
	Last Updated by: JBO
	Last Updated: 01.03.2017
#>
function removePermissionsLevel()
{
    Param(
    [Parameter(Mandatory=$true, Position=1)]
    [string]$siteURL,
    [Parameter(Mandatory=$true, Position=2)]
    [string]$plName 
	)
	
	$functionName = $MyInvocation.MyCommand.Name	
	Write-Debug "Entering $functionName"
	
    try
    {
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if($curWeb -ne $null)
		{
			if(-not[string]::IsNullOrEmpty($plName.Trim())) 
			{
				if($curWeb.RoleDefinitions[$plName] -ne $null)
				{
					Write-Host "[$functionName] About to remove Permission Level '$plName on site '$siteURL'." -foregroundColor Magenta `r 
					$spRoleDefinition = New-Object Microsoft.SharePoint.SPRoleDefinition    
					$spRoleDefinition.Name = $plName   
					$curWeb.RoleDefinitions.Delete($spRoleDefinition.Name) 
					Write-Host "[$functionName] Permission Level '$plName' successfully removed on site '$siteURL'." -foregroundcolor Green `r
				}
				else
				{
					Write-Warning "[$functionName] Permission Level '$plName' does not exist on site '$siteURL'."
				}
			}
			else
			{
				Write-Warning "[$functionName] Permission Level Name could not be empty."
			}
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
		Write-Debug "Exiting $functionName"
    }
}