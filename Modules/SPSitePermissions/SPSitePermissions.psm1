##############################################################################################
#              
# NAME: SPSitePermissions.psm1 
# PURPOSE: 
#	Manage Site Permissions (Add, Delete, Change)
#	Relies on an XML Configuration file for Site Group description.
#	See SPSitePermissions.xml for Schema
#
##############################################################################################



<#
.SYNOPSIS
	Parse the file SPSitePermissions XML object and initiate the Site Permissions customization
	
.DESCRIPTION
	Will parse the file SPSitePermissions XML object making the difference between the permissions to add (AddPermission node),
	the permissions to change (ChangePermission node) and the permissions to be deleted (DeletePermission node)

.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER sitePermissionXML
	XML object of Site Permissions to manage.
	
.EXAMPLE
	browseAndParseSiteGroupsXML -siteURL <SiteURL> -sitePermissionXML <sitePermissionXML>
	
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
function browseAndParseSPSitePermissionsXML()
{
    Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[xml]$sitePermissionXML
	)

	$functionName = $MyInvocation.MyCommand.Name	
	Write-Debug "Entering $functionName"
	
    try
    {
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if ($curWeb -ne $null)
		{
			$SPSitePermissions = $sitePermissionXML.SelectNodes("/SPSitePermissions")
			if($SPSitePermissions.HasChildNodes)
			{
				#--- Add Permissions (SPSitePermissions.AddPermission node)
				Write-Host "[$functionName] About to call browseAndCreateSitePermissionsXML()."-ForegroundColor Cyan `r
				$addPermissionXML =  $SPSitePermissions.AddPermission
				if($addPermissionXML -ne $null -and $addPermissionXML.HasChildNodes)
				{									
					browseAndCreateSitePermissionsXML -siteURL $siteURL -addPermissionXML $addPermissionXML
				}
				else
				{
					Write-Warning "[$functionName] 'AddPermission' node is empty."
				}

				#--- Change Permissions (SPSitePermissions.ChangePermission node)
				Write-Host ""
				Write-Host "[$functionName] About to call browseAndChangeSitePermissionsXML()."-ForegroundColor Cyan `r
				$changePermissionXML =  $SPSitePermissions.ChangePermission
				if($changePermissionXML -ne $null -and $changePermissionXML.HasChildNodes)
				{									
					browseAndChangeSitePermissionsXML -siteURL $siteURL -changePermissionXML $changePermissionXML
				}
				else
				{
					Write-Warning "[$functionName] 'ChangePermission' node is empty."
				}

				#--- Delete Permissions (SPSitePermissions.DeletePermission node)
				Write-Host ""
				Write-Host "[$functionName] About to call browseAndDeleteSitePermissionsXML()."-ForegroundColor Cyan `r
				$deletePermissionXML =  $SPSitePermissions.DeletePermission
				if($deletePermissionXML -ne $null -and $deletePermissionXML.HasChildNodes)
				{									
					browseAndDeleteSitePermissionsXML -siteURL $siteURL -deletePermissionXML $deletePermissionXML
				}
				else
				{
					Write-Warning "[$functionName] 'DeletePermission' node is empty."
				}
			
			}
		    else
			{
				Write-Warning "[$functionName] 'SitePermissions' node is empty."
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



##############################################################################################
######################### HANDLING SITE PERMISSIONS FROM NODES FILE ##########################
##############################################################################################



<#
.SYNOPSIS
	Parse the node AddPermission and initiate Site Permissions creation
	
.DESCRIPTION
	Parse the node AddPermission and initiate Site Permissions creation

.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER addPermissionXML
	XML object of Site Permissions to add.
	
.EXAMPLE
	browseAndCreateSitePermissionsXML -siteURL <SiteURL> -addPermissionXML <addPermissionXML>
	
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
function browseAndCreateSitePermissionsXML()
{
	Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[XML.XmlElement]$addPermissionXML

	)
	
	$functionName = $MyInvocation.MyCommand.Name	
	Write-Debug "Entering $functionName"
	
	try
	{				
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if ($curWeb -ne $null)
		{			
			if($addPermissionXML -ne $null -and $addPermissionXML.HasChildNodes)
			{
				foreach($permissionToAdd in $addPermissionXML.Permission)
				{
					Write-Host "[$functionName] About to assign new permissions level for {$type} '$groupName'." -ForegroundColor Cyan `r
					$groupName = $permissionToAdd.Name
					$permissionLevel = $permissionToAdd.PermissionLevel
					$type = $permissionToAdd.Type
					if($groupName -ne $null -and $permissionLevel -ne $null)
					{						
						if($type.ToLower() -eq "group")
						{
							assignSitePermissionsToGroup -siteUrl $siteURL -groupName $groupName -permissionLevelName $permissionLevel
						}
						elseif($type.ToLower() -eq "user")
						{
							assignSitePermissionsToUser -siteUrl $siteURL -userName $groupName -permissionLevelName $permissionLevel
						}
						else
						{
							Write-Warning "[$functionName] 'Permission' XML node contains invalid type."
						}
					}
					else
					{
						Write-Warning "[$functionName] 'Permission' XML node contains invalid groupName or permissionLevel."
					}
				} #foreach
			}
			else
			{
				Write-Warning "[$functionName] 'AddPermission' XML node is empty."
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
	Parse the node ChangePermission and initiate Site Permissions update
	
.DESCRIPTION
	Parse the node ChangePermission and initiate Site Permissions update
	Will remove ALL existing Site Permissions before setting the expected permission Level

.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER changePermissionXML
	XML object of Site Permissions to change.
	
.EXAMPLE
	browseAndChangeSitePermissionsXML -siteURL <SiteURL> -changePermissionXML <changePermissionXML>
	
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
function browseAndChangeSitePermissionsXML()
{
	Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[XML.XmlElement]$changePermissionXML

	)
	
	$functionName = $MyInvocation.MyCommand.Name	
	Write-Debug "Entering $functionName"
	
	try
	{				
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if ($curWeb -ne $null)
		{
			if($changePermissionXML -ne $null -and $changePermissionXML.HasChildNodes)
			{
				foreach($permissionToAdd in $changePermissionXML.Permission)
				{
					$groupName = $permissionToAdd.Name
					$permissionLevel = $permissionToAdd.PermissionLevel
					$type = $permissionToAdd.Type

					######################################################
					#### STEP 1: Clear current/existing permissions 
					######################################################					
					Write-Host "[$functionName] About to remove all existing site permissions for {$type} '$groupName'." -ForegroundColor Cyan `r
					if($type.ToLower() -eq "group")
					{
						removeAllSitePermissionsToGroup -siteUrl $siteURL -groupName $groupName
					}
					elseif($type.ToLower() -eq "user")
					{
						removeAllSitePermissionsToUser -siteUrl $siteURL -userName $groupName
					}
					else
					{
						Write-Warning "[$functionName] 'Permission' XML node contains invalid type."
					}					
					
					######################################################
					#### STEP 2: Assign new permissions
					######################################################
					Write-Host "[$functionName] About to assign new permissions level for {$type} '$groupName'." -ForegroundColor Cyan `r
					if($groupName -ne $null -and $permissionLevel -ne $null)
					{								
						if($type.ToLower() -eq "group")
						{
							assignSitePermissionsToGroup -siteUrl $siteURL -groupName $groupName -permissionLevelName $permissionLevel
						}
						elseif($type.ToLower() -eq "user")
						{
							assignSitePermissionsToUser -siteUrl $siteURL -userName $groupName -permissionLevelName $permissionLevel
						}
						else
						{
							Write-Warning "[$functionName] 'Permission' XML node contains invalid type."
						}
					}
					else
					{
						Write-Warning "[$functionName] 'ChangePermission' XML node contains invalid groupName or permissionLevel."
					}
				}#foreach
			}
			else
			{
				Write-Warning "[$functionName] 'Permission' XML node is empty."
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
	Parse the node DeletePermission and initiate Site Permissions deletion.
	
.DESCRIPTION
	Parse the node DeletePermission and initiate Site Permissions deletion.
	Will remove ONLY the expected permission Level.

.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER deletePermissionXML
	XML object of Site Permissions to delete.
	
.EXAMPLE
	browseAndDeleteSitePermissionsXML -siteURL <SiteURL> -deletePermissionXML <deletePermissionXML>
	
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
function browseAndDeleteSitePermissionsXML()
{
	Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[XML.XmlElement]$deletePermissionXML
	)
	
	$functionName = $MyInvocation.MyCommand.Name	
	Write-Debug "Entering $functionName"
	
	try
	{				
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if ($curWeb -ne $null)
		{
			if($deletePermissionXML -ne $null -and $deletePermissionXML.HasChildNodes)
			{
				foreach($permissionToDelete in $deletePermissionXML.Permission)
				{					
					$groupName = $permissionToDelete.Name
					$type = $permissionToDelete.Type
					$permissionLevel = $permissionToDelete.PermissionLevel

					if($groupName -ne $null -and $permissionLevel -ne $null)
					{
						
						Write-Host "[$functionName] About to remove '$permissionLevel' site permissions for {$type} '$groupName'." -ForegroundColor Cyan `r
						if($type.ToLower() -eq "group")
						{
							removeSitePermissionsToGroup -siteUrl $siteURL -groupName $groupName -permissionLevelName $permissionLevel
						}
						elseif($type.ToLower() -eq "user")
						{
							removeSitePermissionsToUser -siteUrl $siteURL -userName $groupName -permissionLevelName $permissionLevel
						}
						else
						{
							Write-Warning "[$functionName] 'Permission' XML node contains invalid type."
						}
				
					}
					else
					{
						Write-Warning "[$functionName] 'Permission' XML node contains invalid groupName or permissionLevel."
					}					
			
				} #foreach
			}
			else
			{
				Write-Warning "[$functionName] 'DeletePermission' XML node is empty."
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



##############################################################################################
########################### MANAGING SITE PERMISSIONS INHERITANCE ############################
##############################################################################################



<#
.SYNOPSIS
	Break Permissions inheritance at Site Level, and copy or not, existing permissions

.DESCRIPTION
	Break Permissions inheritance at Site Level, and copy or not, existing permissions

.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER keepExistingSitePermissions
	Whether Site is keeping a copy of the existing permissions after breaking the inheritance
	
.EXAMPLE
	breakSitePermissionsInheritance -siteURL <SiteURL> -keepExistingSitePermissions <TRUE|FALSE>
	
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
function breakSitePermissionsInheritance()
{
    param
	(
		[Parameter(Mandatory = $true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory = $true, Position=2)]
		[boolean]$keepExistingSitePermissions
	)
	
	$functionName = $MyInvocation.MyCommand.Name	
	Write-Debug "Entering $functionName"
	
	try
	{
        $curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if ($curWeb -ne $null)		
		{
			if ($curWeb.HasUniqueRoleAssignments -eq $false)
			{
				Write-Debug "[$functionName] Permissions are inherited on site '$siteURL'."
				$curWeb.BreakRoleInheritance($keepExistingSitePermissions);
				$curWeb.Update();
				Write-Host "[$functionName] Permissions Inheritance successfully brokenon site '$siteURL'." -foregroundcolor Green `r
			}
			else
			{
				Write-Warning "[$functionName] Permissions Inheritance already broken on site '$siteURL'."
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
	Restore Permissions inheritance at Site Level
	
.DESCRIPTION
	Restore Permissions inheritance at Site Level

.PARAMETER siteUrl
	URL of the SharePoint Site
	
.EXAMPLE
	restoreSitePermissionsInheritance -siteURL <SiteURL> 
	
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
function restoreSitePermissionsInheritance()
{
    param
	(
		[Parameter(Mandatory = $true, Position=1)]
		[string]$siteURL
	)
	
	$functionName = $MyInvocation.MyCommand.Name	
	Write-Debug "Entering $functionName"
	
	try
	{
        $curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if ($curWeb -ne $null)		
		{
			if ($curWeb.HasUniqueRoleAssignments -eq $true)
			{
				Write-Host "[$functionName] Permissions Inheritance already broken on site '$siteURL'."
				$curWeb.ResetRoleInheritance()
				$curWeb.Update();
				Write-Host "[$functionName] Permissions Inheritance successfully restored on site '$siteURL'."  -foregroundcolor Green `r
			}
			else
			{
				Write-Warning "[$functionName] Permissions already inherited on " $curWeb.Title
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



##############################################################################################
########################### MANAGING SITE PERMISSIONS for GROUPS #############################
##############################################################################################



<#
.SYNOPSIS
	Check if SharePoint Group has permission on the Site. 
	
.DESCRIPTION
	Check if SharePoint Group has permission on the Site. 
	Return true if SharePoint Group has permission on the Site, false instead

.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER groupName
	Name of the SharePoint Group	
	
.EXAMPLE
	existSitePermissionsForGroup -siteURL <SiteURL> -groupName <groupName>
	
.OUTPUTS
	True, if Group has permissions
	False, if not

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 01.03.2017
	Last Updated by: JBO
	Last Updated: 01.03.2017
#>
function existSitePermissionsForGroup()
{
	param
	(
		[Parameter(Mandatory = $true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory = $true, Position=2)]
		[string]$groupName
	)
	
    $functionName = $MyInvocation.MyCommand.Name	
	Write-Debug "Entering $functionName"
	
	$hasPermission= $false	
	try
	{	
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if ($curWeb -ne $null)		
		{
			$doesGroupExist= existGroup -siteURL $siteURL -groupName $groupName
			if($doesGroupExist= $true)
			{
				$currentGroup=$curWeb.sitegroups[$groupName]
				$roleAssignment = New-Object Microsoft.SharePoint.SPRoleAssignment($currentGroup)
				if($roleAssignment.RoleDefinitionBindings.Count -gt 0)
				{
					$hasPermission= $true
					Write-Debug "[$functionName] Group '$groupName' has permissions on site '$siteURL'."
				}
				else
				{
					Write-Warning "[$functionName] Group '$groupName' does not have permissions on site '$siteURL'."
				}
			}	
			else
			{
				Write-Warning "[$functionName] Group '$groupName' not found on site '$siteURL'." 
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
	return $hasPermission
}


<#
.SYNOPSIS
	Assign the permissions Level $permissionLevelName to group $groupName
	
.DESCRIPTION
	Assign the permissions Level $permissionLevelName to group $groupName

.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER groupName
	Name of the SharePoint Group	
	
.PARAMETER permissionLevelName
	Name of the Permission level
	
.EXAMPLE
	assignSitePermissionsToGroup -siteURL <SiteURL> -groupName <groupName> -permissionLevelName <permissionLevelName>
	
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
function assignSitePermissionsToGroup()
{
	param
	(
		[Parameter(Mandatory = $true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory = $true, Position=2)]
		[string]$groupName,
        [Parameter(Mandatory = $true, Position=3)]
		[string]$permissionLevelName
	)
   
	$functionName = $MyInvocation.MyCommand.Name	
	Write-Debug "Entering $functionName"
	
	try
	{
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if($curWeb -ne $null)
        {
			$doesGroupExist= existGroup -siteURL $siteURL -groupName $groupName
			if($doesGroupExist -eq $true)
			{
				$currentGroup= getGroup -siteURL $siteURL -groupName $groupName
				$roleAssignment = New-Object Microsoft.SharePoint.SPRoleAssignment($currentGroup)
				if($roleAssignment.RoleDefinitionBindings[$permissionLevelName] -eq $null)
				{
					$roleDefinition=$curWeb.RoleDefinitions[$permissionLevelName] 
					$roleAssignment.RoleDefinitionBindings.Add($roleDefinition) 
				
					if ($curWeb.HasUniqueRoleAssignments -eq $false)
					{
						breakSitePermissionsInheritance -siteURL $siteURL -keepExistingSitePermissions $true
						$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
					}
					Write-Host "[$functionName] About to assign Group '$groupName Permissions '$permissionLevelName' on site '$siteURL'." -ForegroundColor Magenta `r 
					$curWeb.RoleAssignments.Add($roleAssignment); 
					$curWeb.Update() 
					Write-Host "[$functionName] Permission '$permissionLevelName' has been assigned to group '$groupName' on site '$siteURL'." -ForegroundColor Green `r 
				}			
				else
				{
					Write-Warning "[$functionName] Permission Level '$permissionLevelName' not found on site '$siteURL'." 
				}
			}	
			else
			{
				Write-Warning "[$functionName] Group '$groupName' not found on Site '$siteURL'." 
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
	If the group groupName exists and if has permissions on site, Remove only his Permissions 'permissionLevelName' on site
	
.DESCRIPTION
	If the group groupName exists and if has permissions on site, Remove only his Permissions 'permissionLevelName' on site

.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER groupName
	Name of the SharePoint Group	
	
.PARAMETER permissionLevelName
	Name of the Permission level
	
.EXAMPLE
	removeSitePermissionsToGroup -siteURL <SiteURL> -groupName <groupName> -permissionLevelName <permissionLevelName>
	
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
function removeSitePermissionsToGroup()
{
    param
	(
		[Parameter(Mandatory = $true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory = $true, Position=2)]
		[string]$groupName,
		[Parameter(Mandatory = $true, Position=3)]
		[string]$permissionLevelName			
	)
	
	$functionName ="[removeSitePermissionsToGroup]"
    try
    {
	
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if ($curWeb -ne $null)		
		{
			$doesGroupExist= existGroup -siteURL $siteURL -groupName $groupName
			if($doesGroupExist -eq $true)
			{
		
				if ($curWeb.HasUniqueRoleAssignments -eq $false)
				{
					breakSitePermissionsInheritance -siteURL $siteURL -keepExistingSitePermissions $true
					$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
				}
				
				$currentGroup= getGroup -siteURL $siteURL -groupName $groupName
				$roleAssignments = $currentGroup.ParentWeb.RoleAssignments.GetAssignmentByPrincipal($currentGroup)
				if( $roleAssignments.RoleDefinitionBindings.Count -gt 0)
				{
					$roleAssignments = $curWeb.RoleAssignments.GetAssignmentByPrincipal($currentGroup)
					Write-Host "[$functionName] About to remove Group '$groupName Permissions '$permissionLevelName' on site '$siteURL'." -ForegroundColor Magenta `r 
					foreach ($permission in $roleAssignments.RoleDefinitionBindings)
					{
						# Write-Host "[removeSitePermissionsToGroup] About to remove"$permission.Name
						$rd = $curWeb.RoleDefinitions[$permissionLevelName]
						$roleAssignments.RoleDefinitionBindings.Remove($rd)
						$roleAssignments.Update()
					}
					#$roleAssignments.Update()
					#$currentGroup.Update()
					$curWeb.Update()
					Write-Host "[$functionName] Group '$groupName' Permissions '$permissionLevelName' has been removed for '$groupName' on site '$siteURL'." -ForegroundColor Green `r 
				}
				else
				{
					Write-Warning "[$functionName] Group '$groupName' has no permissions on site '$siteURL'." 
				}
			}	
			else
			{
				Write-Warning "[$functionName] Group '$groupName' not found on Site '$siteURL'." 
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
	If the group groupName exists and if has permissions on site, Remove ALL of its permissions on site
	
.DESCRIPTION
	If the group groupName exists and if has permissions on site, Remove ALL of its permissions on site

.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER groupName
	Name of the SharePoint Group	
	
.EXAMPLE
	removeAllSitePermissionsToGroup -siteURL <SiteURL> -groupName <groupName>
	
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
function removeAllSitePermissionsToGroup()
{
    param
	(
		[Parameter(Mandatory = $true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory = $true, Position=2)]
		[string]$groupName
	)
   
   	$functionName = $MyInvocation.MyCommand.Name	
	Write-Debug "Entering $functionName"
	
    try
    {	
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
        if($curWeb -ne $null)
        {		
			$doesGroupExist= existGroup -siteURL $siteURL -groupName $groupName
			if($doesGroupExist -eq $true)
			{
				$currentGroup= getGroup -siteURL $siteURL -groupName $groupName
				if ($curWeb.HasUniqueRoleAssignments -eq $false)
				{
					breakSitePermissionsInheritance -siteURL $siteURL -keepExistingSitePermissions $true
					$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
				}
				
				$roleAssignments = $currentGroup.ParentWeb.RoleAssignments.GetAssignmentByPrincipal($currentGroup)
				if($roleAssignments.RoleDefinitionBindings.Count -gt 0)
				{
					$roleAssignments = $curWeb.RoleAssignments.GetAssignmentByPrincipal($currentGroup)
					Write-Host "[$functionName]  About to remove all permissions successfully for '$groupName' on Site '$siteURL'." -ForegroundColor Yellow `r 
					foreach ($permission in $roleAssignments.RoleDefinitionBindings)
					{				
						removeSitePermissionsToGroup -siteURL $siteURL -groupName $groupName -permissionLevelName $permission.Name
					}
					Write-Host "[$functionName]  All permissions successfully removed for '$groupName' on Site '$siteURL'." -ForegroundColor Green `r 
				}
			}	
			else
			{
				Write-Warning "[$functionName] Group '$groupName' not found on Site '$siteURL'." 
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



##############################################################################################
########################### MANAGING SITE PERMISSIONS for USERS ##############################
##############################################################################################



<#
.SYNOPSIS
	Check if User exists and has Explicit or Implicit (as group member) permissions on SharePoint site.
	
.DESCRIPTION
	Check if User exists and has Explicit or Implicit (as group member) permissions on SharePoint site.
	Return true if SharePoint Group has permission on the Site, false instead

.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER userName
	Name of the User
	
.EXAMPLE
	existSitePermissionsForUser -siteURL <SiteURL> -userName <userName>
	
.OUTPUTS
	True, if User has permissions
	False, if not

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 01.03.2017
	Last Updated by: JBO
	Last Updated: 01.03.2017
#>
function existSitePermissionsForUser()
{
	param
	(
		[Parameter(Mandatory = $true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory = $true, Position=2)]
		[string]$userName
	)
	
	$functionName = $MyInvocation.MyCommand.Name	
	Write-Debug "Entering $functionName"
	$hasPermission= $false	
    
	try
	{	
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if ($curWeb -ne $null)		
		{
			$user = $curWeb.EnsureUser($userName)			
			if($user -ne $null)
			{
				#Check if the user is site collection administrator  
				$getuser = get-spuser -identity $user.LoginName -web $siteURL
				if($getuser.isSiteAdmin)  
				{  
					Write-Debug "[$functionName] User '$userName' is a site collection administrator on site '$siteURL'."
					$hasPermission= $true
				}  
				
				#Check if the user is site auditor  
				if($getuser.isSiteAuditor)  
				{  
					Write-Debug "[$functionName] User '$userName' is a site collection auditor on site '$siteURL'."  
					$hasPermission= $true
				}  
				
				#Check if User's permissions
				$permissionInfo = $curWeb.GetUserEffectivePermissionInfo($user.LoginName)  
				$roles = $permissionInfo.RoleAssignments  
				for ($i = 0; $i -lt $roles.Count; $i++)  
				{  
					$bRoles = $roles[$i].RoleDefinitionBindings  
					foreach ($roleDefinition in $bRoles)  
					{  
						$roleDefName = $roleDefinition.Name 
						if ($roles[$i].Member.ToString().Contains('\'))  
						{  
							Write-Debug "[$functionName] User '$userName' has direct {$roleDefName} permissions on site '$siteURL'."
							$hasPermission= $true
						}  
						else  
						{  
							$rolesGroupMember= $roles[$i].Member.ToString();
							Write-Debug "The User '$userName' has permissions {$roleDefName} given via Group {$rolesGroupMember} on site '$siteURL'." 
							$hasPermission= $true
						}  
					}  
				}  
			}	
			else
			{
				Write-Host "[$functionName] User '$userName' not found on Site '$siteURL'." 
			}
		}
		else
		{
			Write-Host "[$functionName] Site '$siteURL' does not exist."
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
	return $hasPermission
}


<#
.SYNOPSIS
	Assign the permissions Level $permissionLevelName to user $userName
	
.DESCRIPTION
	Assign the permissions Level $permissionLevelName to group $userName

.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER userName
	Name of the SharePoint Group	
	
.PARAMETER permissionLevelName
	Name of the Permission level
	
.EXAMPLE
	assignSitePermissionsToGroup -siteURL <SiteURL> -userName <userName> -permissionLevelName <permissionLevelName>
	
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
function assignSitePermissionsToUser()
{
	param
	(
		[Parameter(Mandatory = $true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory = $true, Position=2)]
		[string]$userName,
        [Parameter(Mandatory = $true, Position=3)]
		[string]$permissionLevelName
	)
	
	$functionName = $MyInvocation.MyCommand.Name	
	Write-Debug "Entering $functionName"
	
	try
	{
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue		
        if($curWeb -ne $null)
        {
			$user = $curWeb.EnsureUser($userName)			
			if($user -ne $null)
			{
				$roleDefinition=$curWeb.RoleDefinitions[$permissionLevelName] 
				if ($roleDefinition -ne $null)
				{
					Write-Host "[$functionName] About to assign Permission '$permissionLevelName' to User '$userName' on Site '$siteURL'." -foregroundcolor Magenta `r
					$spUserToAdd= Set-SPUser -web $curWeb -Identity $user -AddPermissionLevel	$permissionLevelName
					Write-Host "[$functionName] Permission '$permissionLevelName' has been assigned to '$userName' with Set-SPUser." -ForegroundColor Green `r 
				}
				else
				{
				   Write-Warning "[$functionName] Permission '$permissionLevelName' does not exist on Site '$siteURL'." 
				}
			}
			else
			{
				Write-Warning "[$functionName] User '$userName' not found on Site '$siteURL'." 
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
	Remove only his Permissions 'permissionLevelName' on site If the user userName exists and if has permissions
	
.DESCRIPTION
	Remove only his Permissions 'permissionLevelName' on site If the user userName exists and if has permissions

.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER userName
	Name of the SharePoint Group	
	
.PARAMETER permissionLevelName
	Name of the Permission level
	
.EXAMPLE
	removeSitePermissionsToUser -siteURL <SiteURL> -userName <userName> -permissionLevelName <permissionLevelName>
	
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
function removeSitePermissionsToUser()
{
    param
	(
		[Parameter(Mandatory = $true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory = $true, Position=2)]
		[string]$userName,
		[Parameter(Mandatory = $true, Position=3)]
		[string]$permissionLevelName			
	)
   
	$functionName = $MyInvocation.MyCommand.Name	
	Write-Debug "Entering $functionName"

    try
    {
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
        if($curWeb -ne $null)
        {
			$user = $curWeb.EnsureUser($userName)
			if($user -ne $null)
			{
				$roleDefinition=$curWeb.RoleDefinitions[$permissionLevelName] 
				if ($roleDefinition -ne $null)
				{
					$spUserToAdd= Set-SPUser -web $curWeb -Identity $user -RemovePermissionLevel $permissionLevelName
					Write-Host "[$functionName] Permission '$permissionLevelName' has been removed for '$userName' with Set-SPUser." -ForegroundColor Green `r 
				}
				else
				{
				   Write-Warning "[$functionName] Permissions '$permissionLevelName' does not exist. Impossible to perform action." 
				}
			}
			else
			{
				Write-Warning "[$functionName] User '$userName' not found on Site '$siteURL'." 
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
	If the user userName exists and if has permissions on site, Remove ALL his Permissions on Site
	
.DESCRIPTION
	If the user userName exists and if has permissions on site, Remove ALL his Permissions on Site

.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER userName
	Name of the SharePoint Group	

.EXAMPLE
	removeAllSitePermissionsToUser -siteURL <SiteURL> -userName <userName>
	
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
function removeAllSitePermissionsToUser()
{
    param
	(
		[Parameter(Mandatory = $true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory = $true, Position=2)]
		[string]$userName
	)
   
	$functionName = $MyInvocation.MyCommand.Name	
	Write-Debug "Entering $functionName"
	
    try
    {	
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
        if($curWeb -ne $null)
        {
			if ($curWeb.HasUniqueRoleAssignments -eq $false)
			{
				breakSitePermissionsInheritance -siteURL $siteURL -keepExistingSitePermissions $true
				$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
			}

			$user = $curWeb.EnsureUser($userName)
            if($user -ne $null)
			{
				$roleAssignments = $curWeb.RoleAssignments.GetAssignmentByPrincipal($user)
				if($roleAssignments.RoleDefinitionBindings.Count -gt 0)
				{			 
					Write-Host "[$functionName] About to remove All permissions for '$userName' on site '$siteURL'." -ForegroundColor Magenta `r 
					foreach ($permission in $roleAssignments.RoleDefinitionBindings)
					{				
						removeSitePermissionsToUser -siteURL $siteURL -userName $userName -permissionLevelName $permission.Name
					}
					Write-Host "[$functionName] All permissions successfully removed for '$userName' on site '$siteURL'." -ForegroundColor Green `r 
				}
				else
				{
					Write-Warning "[$functionName] No permissions for '$userName' on site." 
				}
			}
			else
			{
				Write-Warning "[$functionName] User '$userName' not found on Site '$siteURL'." 
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


