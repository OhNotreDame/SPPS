##############################################################################################
#              
# NAME: SPListPermissions.psm1 
# PURPOSE: 
#	Manage List Permissions (Add, Delete, Change)
#	Relies on an XML Configuration file for ListPermissions description.
#	See SPListPermissions.xml for Schema
#
# SOURCE : https://github.com/OhNotreDame/SPPS
#
##############################################################################################


<#
.SYNOPSIS
	Parse the file SPListPermissions XML object and initiate List Permissions customization
	
.DESCRIPTION
	Will parse the file SPListPermissions XML object making the difference between the permissions to add (AddPermission node),
	the permissions to change (ChangePermission node) and the permissions to be deleted (DeletePermission node)

.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER listPermissionXML
	XML object of List Permissions to manage.
	
.EXAMPLE
	browseAndParseSPListPermissionsXML -siteURL <SiteURL> -listPermissionXML <listPermissionXML>
	
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
function browseAndParseSPListPermissionsXML()
{
    Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[xml]$listPermissionXML
	)

	$functionName = $MyInvocation.MyCommand.Name	
	Write-Debug "Entering $functionName"
	
    try
    {
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if ($curWeb -ne $null)
		{
			$SPListPermissions = $listPermissionXML.SelectNodes("/SPListPermissions")
			if($SPListPermissions -ne $null -and $SPListPermissions.HasChildNodes)
			{	
				#--- Add Permissions (SPListPermissions.AddPermission node)			
				Write-Host "`n[$functionName] About to call browseAndCreateListPermissionsXML()." -ForegroundColor Cyan `r
				$addPermissionXML =  $SPListPermissions.AddPermission
				if($addPermissionXML -ne $null -and $addPermissionXML.HasChildNodes)
				{									
					browseAndCreateListPermissionsXML -siteURL $siteURL -addPermissionXML $addPermissionXML
				}
				else
				{
					Write-Warning "[$functionName] 'AddPermission' node is empty."
				}

				#--- Change Permissions (SPListPermissions.ChangePermission node)
				Write-Host "`n[$functionName] About to call browseAndChangeListPermissionsXML()." -ForegroundColor Cyan `r
				$changePermissionXML =  $SPListPermissions.ChangePermission
				if($changePermissionXML -ne $null -and $changePermissionXML.HasChildNodes)
				{									
					browseAndChangeListPermissionsXML -siteURL $siteURL -changePermissionXML $changePermissionXML
				}
				else
				{
					Write-Warning "[$functionName] 'ChangePermission' node is empty." 
				}

				#--- Delete Permissions (SPListPermissions.DeletePermission node)
				Write-Host "`n[$functionName] About to call browseAndDeleteListPermissionsXML()." -ForegroundColor Cyan `r
				$deletePermissionXML =  $SPListPermissions.DeletePermission
				if($deletePermissionXML -ne $null -and $deletePermissionXML.HasChildNodes)
				{									
					browseAndDeleteListPermissionsXML -siteURL $siteURL -deletePermissionXML $deletePermissionXML
				}
				else
				{
					Write-Warning "[$functionName] 'DeletePermission' node is empty."
				}
			
			}
			else
			{
				Write-Warning "[$functionName] 'SPListPermissions' node is empty."
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
######################### HANDLING LIST PERMISSIONS FROM NODES FILE ##########################
##############################################################################################



<#
.SYNOPSIS
	Parse the node AddPermission and initiate List Permissions creation
	
.DESCRIPTION
	Parse the node AddPermission and initiate List Permissions creation

.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER listPermissionXML
	XML object of List Permissions to create.
	
.EXAMPLE
	browseAndCreateListPermissionsXML -siteURL <SiteURL> -addPermissionXML <addPermissionXML>
	
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
function browseAndCreateListPermissionsXML()
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
					$groupName = $permissionToAdd.Name
					$listName = $permissionToAdd.ListName
					$permissionLevel = $permissionToAdd.PermissionLevel
					$type = $permissionToAdd.Type

					Write-Debug "[$functionName] About to assign new permissions level for group '$groupName' on list '$listName'."
					if(($groupName -ne $null) -and ($permissionLevel -ne $null) -and ($listName -ne $null))
					{
						if($type.ToLower() -eq "group")
						{
							assignListPermissionsToGroup -siteUrl $siteURL -groupName $groupName -permissionLevelName $permissionLevel -listName $listName
						}
						elseif($type.ToLower() -eq "user")
						{
							assignListPermissionsToUser -siteUrl $siteURL -userName $groupName -permissionLevelName $permissionLevel -listName $listName
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
}


<#
.SYNOPSIS
	Parse the node ChangePermission and initiate List Permissions update
	
.DESCRIPTION
	Parse the node ChangePermission and initiate List Permissions update
	Will remove all existing List Permissions before setting the expected permission Level

.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER changePermissionXML
	XML object of List Permissions to change.
	
.EXAMPLE
	browseAndChangeListPermissionsXML -siteURL <SiteURL> -changePermissionXML <changePermissionXML>
	
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
function browseAndChangeListPermissionsXML()
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
					$listName = $permissionToAdd.ListName
					$permissionLevel = $permissionToAdd.PermissionLevel				
					$type = $permissionToAdd.Type

					######################################################
					#### STEP 1: Clear current/existing permissions 
					######################################################					
					Write-Debug "[$functionName] About to remove all existing site permissions for {$type} '$groupName' on list '$listName'."
					if($type.ToLower() -eq "group")
					{
							removeAllListPermissionsToGroup -siteUrl $siteURL -groupName $groupName -listName $listName
					}
					elseif($type.ToLower() -eq "user")
					{
							removeAllListPermissionsToUser -siteUrl $siteURL -userName $groupName -listName $listName
					}
					else
					{
						Write-Warning "[$functionName] 'Permission' XML node contains invalid type."
					}

					######################################################
					#### STEP 2: Assign new permissions
					######################################################
					Write-Debug "[$functionName] About to assign new permissions level for {$type} '$groupName' on list '$listName'."
					if($groupName -ne $null -and $permissionLevel -ne $null)
					{
						if($type.ToLower() -eq "group")
						{
								assignListPermissionsToGroup -siteUrl $siteURL -groupName $groupName -permissionLevelName $permissionLevel -listName $listName
						}
						elseif($type.ToLower() -eq "user")
						{
								assignListPermissionsToUser -siteUrl $siteURL -userName $groupName -permissionLevelName $permissionLevel -listName $listName
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
				Write-Warning "[$functionName] 'ChangePermission' XML node is empty."
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
	Parse the node DeletePermission and initiate List Permissions update
	
.DESCRIPTION
	Parse the node ChangePermission and initiate List Permissions update
	Will remove all existing List Permissions before setting the expected permission Level

.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER deletePermissionXML
	XML object of List Permissions to delete.
	
.EXAMPLE
	browseAndChangeListPermissionsXML -siteURL <SiteURL> -deletePermissionXML <deletePermissionXML>
	
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
function browseAndDeleteListPermissionsXML()
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
					$listName = $permissionToDelete.ListName
					$permissionLevel = $permissionToDelete.PermissionLevel
					$type = $permissionToDelete.Type

					Write-Debug "[$functionName] About to remove permissions level '$permissionLevel' for {$type} '$groupName' on list '$listName'."
					if(($groupName -ne $null) -and ($permissionLevel -ne $null) -and ($listName -ne $null))
					{
						if($type.ToLower() -eq "group")
						{			
							removeListPermissionsToGroup  -siteUrl $siteURL -groupName $groupName -permissionLevelName $permissionLevel -listName $listName
						}
						elseif($type.ToLower() -eq "user")
						{							
							removeListPermissionsToUser -siteUrl $siteURL -userName $groupName -permissionLevelName $permissionLevel -listName $listName
						}
						else
						{
							Write-Host "[$functionName] 'Permission' XML node contains invalid type."
						}
					}
					else
					{
						Write-Host "[$functionName] 'Permission' XML node contains invalid groupName or permissionLevel."
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
}


##############################################################################################
########################### MANAGING LIST PERMISSIONS INHERITANCE ############################
##############################################################################################


<#
.SYNOPSIS
	Break Permissions inheritance at List Level, and copy or not, existing permissions

.DESCRIPTION
	Break Permissions inheritance at List Level, and copy or not, existing permissions

.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER listName
	Name of the List
	
.PARAMETER keepExistingSitePermissions
	Whether List is keeping a copy of the existing permissions after breaking the inheritance
	
.EXAMPLE
	breakListPermissionsInheritance -siteURL <SiteURL> -listName <listName> -keepExistingSitePermissions <TRUE|FALSE>
	
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
function breakListPermissionsInheritance()
{
    param
	(
		[Parameter(Mandatory = $true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory = $true, Position=2)]
		[string]$listName,
		[Parameter(Mandatory = $true, Position=3)]
		[Boolean]$keepExistingSitePermissions
	)

	$functionName = $MyInvocation.MyCommand.Name	
	Write-Debug "Entering $functionName"
	
    try
    {
        $curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if ($curWeb -ne $null)
		{
			$list = $curWeb.Lists.TryGetList($listName)
			if($list -ne $null)
			{
				$unique = $list.hasuniqueroleassignments
				if($unique -eq $false)
				{	
						Write-Host "[$functionName] Permissions are inherited on $listName." 
						$list.BreakRoleInheritance($keepExistingSitePermissions)
						Write-Host "[$functionName] Permissions Inheritance successfully broken on $listName." -foregroundcolor Green `r
				}
				else
				{
					Write-Warning "[$functionName] Permissions Inheritance is already broken on $listName."  
				}
			}
			else
			{
				Write-Warning "[$functionName] List '$listName' does not exist." 
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
}


<#
.SYNOPSIS
	Restore Permissions inheritance at List Level
	
.DESCRIPTION
	Restore Permissions inheritance at List Level

.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER listName
	Name of the List
	
.EXAMPLE
	restoreSitePermissionsInheritance -siteURL <SiteURL> -listName <listName>
	
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
function restoreListPermissionsInheritance()
{
    param
	(
		[Parameter(Mandatory = $true, Position=1)]
		[string]$siteURL,
        [Parameter(Mandatory = $true, Position=2)]
		[string]$listName
	)

	$functionName = $MyInvocation.MyCommand.Name	
	Write-Debug "Entering $functionName"
	
    try
    {
        $curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if ($curWeb -ne $null)
		{
			$list = $curWeb.Lists.TryGetList($listName)
			if($list -ne $null)
			{
				if ($list.HasUniqueRoleAssignments -eq $true)
				{
					Write-Host "[$functionName]  About to restore Permission inheritance on '$list'." -ForegroundColor Magenta `r
					$list.ResetRoleInheritance();
					$list.Update();
					$curWeb.Update();
					Write-Host "[$functionName] Permissions Inheritance successfully restored on '$list'." -foregroundcolor Green `r
				}
				else
				{
					Write-Warning "[$functionName] Permissions already inherited on '$list'." 
				}
			}
			else
			{
				Write-Warning "[$functionName] List '$listName' does not exist." 
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
########################### MANAGING LIST PERMISSIONS for GROUPS #############################
##############################################################################################

<#
.SYNOPSIS
	Check if SharePoint Group has permission on the List. 
	
.DESCRIPTION
	Check if SharePoint Group has permission on the List. 
	Return true if SharePoint Group has permission on the List, false instead

.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER groupName
	Name of the SharePoint Group	
	
.PARAMETER listName
	Name of the List
	
.EXAMPLE
	existSitePermissionsForGroup -siteURL <SiteURL> -groupName <groupName> -listName <listName>
	
.OUTPUTS
	true, if Group has permissions
	false, if not

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 01.03.2017
	Last Updated by: JBO
	Last Updated: 01.03.2017
#>
function existListPermissionsForGroup()
{
    param
	(
		[Parameter(Mandatory = $true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory = $true, Position=2)]
		[string]$groupName,
		[Parameter(Mandatory = $true, Position=3)]
		[string]$listName
	)
    
	$functionName = $MyInvocation.MyCommand.Name	
	Write-Debug "Entering $functionName"
	$existListPermission= $false
 
    try
    {
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if ($curWeb -ne $null)
		{
			$list = $curWeb.Lists.TryGetList($listName)
			if ($list -ne  $null)
			{
				$isGroupExist= existGroup -siteURL $siteURL -groupName $groupName
				if ($isGroupExist -eq $true)
				{
					#$unique = $list.hasuniqueroleassignments
					#if($unique -eq $false)
					#{
						$currentGroup=$curWeb.sitegroups[$groupName]
						$roleAssignments = $list.RoleAssignments.GetAssignmentByPrincipal($currentGroup)
						if ($roleAssignments -ne $null) 
						{
							if($roleAssignments.RoleDefinitionBindings.Count -gt 0)
							{	
								$existListPermission= $true
								$curPerm = $roleAssignments.RoleDefinitionBindings.Name
								Write-Debug "$functionName Group '$groupName' has direct permissions {$curPerm} on List '$listName'."
							}
							else
							{
								Write-Warning "$functionName Group '$groupName' does not have direct permissions on List '$listName'." 
							}
						}
						else
						{
							Write-Warning "$functionName Group '$groupName' does not have direct permissions on List '$listName'." 
						}
					#}
					#else
					#{
					#	$existListPermission= $false
					#	Write-Host "[$functionName] List '$listName' inherits permissions from parent Site '$siteURL'. No explicit permissions for '$groupName'."  
					#}
				}
				else
				{
					Write-Warning "[$functionName] Group '$groupName' does not exist on Site '$siteURL'." 
				}
			}
			else
			{
				$existListPermission= $false
				Write-Warning "[$functionName] List '$listName' does not exist on Site '$siteURL'." 
			}
		}
		else
		{
			$existListPermission= $false
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
    return $existListPermission
}


<#
.SYNOPSIS
	Assign the permissions Level $permissionLevelName to group $groupName on the list $listName
	
.DESCRIPTION
	Assign the permissions Level $permissionLevelName to group $groupName on the list $listName

.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER groupName
	Name of the SharePoint Group	
	
.PARAMETER permissionLevelName
	Name of the Permission level
	
.PARAMETER listName
	Name of the List
		
.EXAMPLE
	assignSitePermissionsToGroup -siteURL <SiteURL> -groupName <groupName> -permissionLevelName <permissionLevelName> -listName <listName>
	
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
function assignListPermissionsToGroup()
{
    param
	(
		[Parameter(Mandatory = $true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory = $true, Position=2)]
		[string]$groupName,
		[Parameter(Mandatory = $true, Position=3)]
		[string]$permissionLevelName,
		[Parameter(Mandatory = $true, Position=4)]
		[string]$listName
	)   
	
	$functionName = $MyInvocation.MyCommand.Name	
	Write-Debug "Entering $functionName"
 
    try
    {
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if ($curWeb -ne $null)
		{
			$list = $curWeb.Lists.TryGetList($listName)		
			if ($list -ne $null)
			{
				$unique = $list.HasUniqueRoleAssignments
				if(!$unique)
				{
					$list.BreakRoleInheritance($true) #Break List Permission, keeping current permissions assignments
				}
				
				$isGroupExist= existGroup -siteURL $siteURL -groupName $groupName
				if($isGroupExist -eq $true)
				{
					$currentGroup = $curWeb.SiteGroups[$groupName]
					$roleAssignment = new-object Microsoft.SharePoint.SPRoleAssignment($currentGroup)
					$roleDefinition = $curWeb.RoleDefinitions[$permissionLevelName];
					if($roleDefinition -ne $null)
					{
						Write-Host "[$functionName] About to assign Permission '$permissionLevelName' to Group '$groupName' on List '$listName'." -ForegroundColor Magenta `r
						$roleAssignment.RoleDefinitionBindings.Add($roleDefinition);
						$list.RoleAssignments.Add($roleAssignment)
						$list.Update();
						Write-Host "[$functionName] Permission '$permissionLevelName' successfully added to Group '$groupName' on List '$listName'." -foregroundcolor Green `r
					}
					else
					{
						Write-Warning "[$functionName] Permissions '$permissionLevelName' does not exist on site '$siteURL'." 
					}
				}
				else
				{
				   Write-Warning "[$functionName] Group '$groupName' does not exist on site '$siteURL'." 
				}
			}
			else
			{
				Write-Warning "[$functionName] List '$listName' does not exist on site '$siteURL'." 
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
	If the group groupName exists and if has permissions on list, Remove ALL of its permissions on list
	
.DESCRIPTION
	If the group groupName exists and if has permissions on list, Remove ALL of its permissions on list

.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER groupName
	Name of the SharePoint Group	
	
.PARAMETER listName
	Name of the List
	
.EXAMPLE
	removeAllListPermissionsToGroup -siteURL <SiteURL> -groupName <groupName> -listName <listName>
	
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
function removeAllListPermissionsToGroup()
{
    param
	(
		[Parameter(Mandatory = $true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory = $true, Position=2)]
		[string]$groupName,
		[Parameter(Mandatory = $true, Position=3)]
		[string]$listName
	)
 
    $functionName = $MyInvocation.MyCommand.Name	
	Write-Debug "Entering $functionName"
	
	try
    {
          $curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		  if ($curWeb -ne $null)
        {
			 $list = $curWeb.Lists.TryGetList($listName)
			 if ($list -ne $null)
			{
				$isGroupExist= existGroup -siteURL $siteURL -groupName $groupName
				if($isGroupExist -eq $true)
				{
					$currentGroup = $curWeb.SiteGroups[$groupName]
					$roleAssignment = new-object Microsoft.SharePoint.SPRoleAssignment($currentGroup)
					if($roleAssignment -ne $null)
					{                    
						Write-Host "[$functionName] About to assign all Permissions of group '$groupName' on List '$listName'." -ForegroundColor Magenta `r
						$list.RoleAssignments.Remove($currentGroup)
						$list.Update();
						Write-Host "[$functionName] All Permissions of group '$groupName' successfully removed from '$listName'." -foregroundcolor Green `r
					}
					else
					{
						Write-Warning "[$functionName] Group '$groupName' has no permissions on '$listName'."
					}           
				}
				else
				{
				   Write-Warning "[$functionName] Group '$groupName' does not exist." 
				}
			}
			else
			{
				Write-Warning "[$functionName] List '$listName' does not exist." 
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
# removeListPermissionsToGroup (TO FINISH)
##############################################################################################
<#
.SYNOPSIS
	Remove the permissions Level $permissionLevelName to group $groupName on the list $listName
	
.DESCRIPTION
	Remove the permissions Level $permissionLevelName to group $groupName on the list $listName

.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER groupName
	Name of the SharePoint Group	
	
.PARAMETER permissionLevelName
	Name of the Permission level
	
.PARAMETER listName
	Name of the List
		
.EXAMPLE
	removeListPermissionsToGroup -siteURL <SiteURL> -groupName <groupName> -permissionLevelName <permissionLevelName> -listName <listName>
	
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
function removeListPermissionsToGroup()
{
    param
	(
		[Parameter(Mandatory = $true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory = $true, Position=2)]
		[string]$groupName,
		[Parameter(Mandatory = $true, Position=3)]
		[string]$permissionLevelName,
		[Parameter(Mandatory = $true, Position=4)]
		[string]$listName			
	)
   
	$functionName = $MyInvocation.MyCommand.Name	
	Write-Debug "Entering $functionName"
	
    try
    {
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
        if($curWeb -ne $null)
        {
			$list = $curWeb.Lists.TryGetList($listName)		
			if ($list -ne $null)
			{
				$isGroupExist= existGroup -siteURL $siteURL -groupName $groupName
				if($isGroupExist -eq $true)
				{
					$currentGroup = $curWeb.SiteGroups[$groupName]
					$roleAssignments = $list.RoleAssignments.GetAssignmentByPrincipal($currentGroup)
					if($roleAssignments.RoleDefinitionBindings.Count -gt 0)
					{		 
						$rd = $curWeb.RoleDefinitions[$permissionLevelName]
						if($rd -ne $null)
						{
							Write-Host "[$functionName] About to remove '$permissionLevelName' Permissions to Group '$groupName' on List '$listName'." -ForegroundColor Magenta `r
							$roleAssignments.RoleDefinitionBindings.Remove($rd)
							$roleAssignments.Update()	
							$list.Update()			
							Write-Host "[$functionName] Permissions '$permissionLevelName' successfully removed for Group '$groupName' on List '$listName'." -foregroundcolor Green `r
						}
						else
						{
							Write-Warning "[$functionName] Permission '$permissionLevelName' not found on Site '$siteURL'."
						}
					}
					else
					{
						Write-Warning "[$functionName] Group '$groupName' has no permission on on List '$listName'."
					}
				}
				else
				{
					Write-Warning "[$functionName] Group '$groupName' not found on Site '$siteURL'."
				}
			}
			else
			{
				Write-Warning "[$functionName] List '$listName' does not exist on Site '$siteURL."
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



##############################################################################################
# existListPermissionsForUser (MAKE IT CLAIMS ROBUST)
# CLAIMS: https://cann0nf0dder.wordpress.com/2014/08/04/get-the-user-claim-token-string-in-powershell-from-windows-domainname/
##############################################################################################
<#
.SYNOPSIS
	Check if SharePoint Group has permission on the Site. 
	
.DESCRIPTION
	Check if SharePoint Group has permission on the Site. 
	Return true if SharePoint Group has permission on the Site, false instead

.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER userName
	Login Name of the User

.PARAMETER listName
	Name of the List
	
.EXAMPLE
	existListPermissionsForUser -siteURL <SiteURL> -userName <userName> -listName <listName>
	
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
function existListPermissionsForUser()
{
    param
	(
		[Parameter(Mandatory = $true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory = $true, Position=2)]
		[string]$userName,
		[Parameter(Mandatory = $true, Position=3)]
		[string]$listName
	)
    
	$functionName = $MyInvocation.MyCommand.Name	
	Write-Debug "Entering $functionName"
	$existListPermission= $false
 
	#Write-Host "[$functionName] [PARAM] userName: $userName" -ForegroundColor Cyan `r
	#Write-Host "[$functionName] [PARAM] listName: $listName" -ForegroundColor Cyan `r
    try
    {
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if ($curWeb -ne $null)
		{
			$list = $curWeb.Lists.TryGetList($listName)
			if ($list -ne  $null)
			{
				#$unique = $list.HasUniqueRoleAssignments
				#if($unique)
				#{
					$userEnsured = $curWeb.EnsureUser($userName)			
					$user = Get-SPUser -web $siteURL -Identity $userEnsured.LoginName 
					if($user -ne $null)
					{
						$roleAssignments = $list.RoleAssignments.GetAssignmentByPrincipal($user)
						if($roleAssignments.RoleDefinitionBindings.Count -gt 0)
						{	
							$existListPermission= $true
							$curPerm = $roleAssignments.RoleDefinitionBindings.Name
							Write-Host  "$functionName User '$userName' has direct permissions {$curPerm} on List '$listName'." -ForegroundColor Green `r
						}
						else
						{
							Write-Warning  "$functionName User '$userName' does not have direct permissions on List '$listName'." 
						}
					}
					else
					{
						Write-Warning "[$functionName] User '$userName' does not exist." 
					}
				#}
				#else
				#{
				#	Write-Host "[$functionName] List '$listName' inherits permissions from parent Site '$siteURL'. No explicit permissions for '$userGroup'."  
				#}
			}
			else
			{
				Write-Warning "[$functionName] List '$listName' does not exist on Site '$siteURL'." 
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
    return $existListPermission
}


##############################################################################################
# assignListPermissionsToUser (TO FINISH)
##############################################################################################
<#
.SYNOPSIS
	Assign permission $permissionLevelName to user $userName on list $listName
	
.DESCRIPTION
	Assign permission $permissionLevelName to user $userName on list $listName

.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER userName
	Login Name of the User
	
.PARAMETER permissionLevelName
	Name of the Permission Level
	
.PARAMETER listName
	Name of the List
	
.EXAMPLE
	assignListPermissionsToUser -siteURL <SiteURL> -userName <userName> -permissionLevelName <permissionLevelName> -listName <listName>
	
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
function assignListPermissionsToUser()
{
    param
	(
		[Parameter(Mandatory = $true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory = $true, Position=2)]
		[string]$userName,
		[Parameter(Mandatory = $true, Position=3)]
		[string]$permissionLevelName,
		[Parameter(Mandatory = $true, Position=4)]
		[string]$listName
	)
   
	$functionName = $MyInvocation.MyCommand.Name	
	Write-Debug "Entering $functionName"
	
    try
    {
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if ($curWeb -ne $null)
		{
			$list = $curWeb.Lists.TryGetList($listName)		
			if ($list -ne $null)
			{
				$unique = $list.HasUniqueRoleAssignments
				if(!$unique)
				{
					$list.BreakRoleInheritance($true) #Break List Permission, keeping current permissions assignments
				}
				
				$userEnsured = $curWeb.EnsureUser($userName)				
				if($userEnsured -ne $null)
				{
					$roleAssignment = new-object Microsoft.SharePoint.SPRoleAssignment($userEnsured)
					$roleDefinition = $curWeb.RoleDefinitions[$permissionLevelName];
					if($roleDefinition -ne $null)
					{
						Write-Host "[$functionName] About to assign Permission '$permissionLevelName' to User '$userName' on List '$listName'." -ForegroundColor Magenta `r
						$roleAssignment.RoleDefinitionBindings.Add($roleDefinition);
						$list.RoleAssignments.Add($roleAssignment)
						$list.Update();
						$curWeb.Update();
						Write-Host "[$functionName] Permission '$permissionLevelName' successfully added to User '$userName' on List '$listName'." -foregroundcolor Green `r
					}
					else
					{
						Write-Warning "[$functionName] Permission '$permissionLevelName' does not exist on site '$siteURL'." 
					}
				}
				else
				{
					Write-Warning "[$functionName] User '$userName' not found on Site '$siteURL'."
				}	
			}
			else
			{
				Write-Warning "[$functionName] List '$listName' does not exist on site '$siteURL'." 
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
# removeListPermissionsToUser (TO FINISH)
###############################################################################################
<#
.SYNOPSIS
	Remove permission $permissionLevelName to user $userName on list $listName
	
.DESCRIPTION
	Remove permission $permissionLevelName to user $userName on list $listName

.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER userName
	Login Name of the User
	
.PARAMETER permissionLevelName
	Name of the Permission Level
	
.PARAMETER listName
	Name of the List
	
.EXAMPLE
	removeListPermissionsToUser -siteURL <SiteURL> -userName <userName> -permissionLevelName <permissionLevelName> -listName <listName>
	
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
function removeListPermissionsToUser()
{
    param
	(
		[Parameter(Mandatory = $true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory = $true, Position=2)]
		[string]$userName,
		[Parameter(Mandatory = $true, Position=3)]
		[string]$permissionLevelName,
		[Parameter(Mandatory = $true, Position=4)]
		[string]$listName
	)
   
	$functionName = $MyInvocation.MyCommand.Name	
	Write-Debug "Entering $functionName"
	
    try
    {
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
        if($curWeb -ne $null)
        {
			$list = $curWeb.Lists.TryGetList($listName)		
			if ($list -ne $null)
			{
				$user = $curWeb.EnsureUser($userName)
				if($user -ne $null)
				{
					$roleAssignments = $list.RoleAssignments.GetAssignmentByPrincipal($user)
					if($roleAssignments.RoleDefinitionBindings.Count -gt 0)
					{		 
						$rd = $curWeb.RoleDefinitions[$permissionLevelName]
						if($rd -ne $null)
						{
							Write-Host "[$functionName] About to remove '$permissionLevelName' Permissions to User '$userName' on List '$listName'." -ForegroundColor Magenta `r
							$roleAssignments.RoleDefinitionBindings.Remove($rd)
							$roleAssignments.Update()	
							$list.Update()			
							Write-Host "[$functionName] Permissions '$permissionLevelName' successfully removed for User '$userName' on List '$listName'." -foregroundcolor Green `r
						}
						else
						{
							Write-Warning "[$functionName] Permission '$permissionLevelName' not found on Site '$siteURL'."
						}

					}
					else
					{
						Write-Warning "[$functionName] User '$userName' does not have ANY DIRECT permission on List '$listName'."
					}
					#
					# Do Something
					#
					#Write-Host "[$functionName] NOT IMPLEMENTED YET."				
				}
				else
				{
					Write-Warning "[$functionName] User '$userName' not found on Site '$siteURL'."
				}
			}
			else
			{
				Write-Warning "[$functionName] List '$listName' does not exist on Site '$siteURL."
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
# removeAllListPermissionsToUser (TO FINISH)
##############################################################################################
<#
.SYNOPSIS
	Remove all User Permissions on list $listName
	
.DESCRIPTION
	Remove all User Permissions on list $listName

.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER userName
	Name of the SharePoint Group	
	
.PARAMETER listName
	Name of the List
	
.EXAMPLE
	removeAllListPermissionsToUser -siteURL <SiteURL> -userName <userName> -listName <listName>
	
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
function removeAllListPermissionsToUser()
{
    param
	(
		[Parameter(Mandatory = $true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory = $true, Position=2)]
		[string]$userName,
		[Parameter(Mandatory = $true, Position=3)]
		[string]$listName
	)
 
    $functionName = $MyInvocation.MyCommand.Name	
	Write-Debug "Entering $functionName"
	
	try
    {
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if ($curWeb -ne $null)
		{
			$list = $curWeb.Lists.TryGetList($listName)
			if ($list -ne $null)
			{
				$user = $curWeb.EnsureUser($userName)
				if($user -ne $null)
				{
					Write-Host "[$functionName] About to assign all Permissions to User '$userName' on List '$listName'." -ForegroundColor Magenta `r
					$list.RoleAssignments.Remove($user)
					Write-Host "[$functionName] All User '$userName' Permissions have been removed on List '$listName'." -foregroundcolor Green `r		
				}					
				else
				{
				   Write-Warning "[$functionName] User '$userName' does not exist." 
				}
			}
			else
			{
				Write-Warning "[$functionName] List '$listName' does not exist." 
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