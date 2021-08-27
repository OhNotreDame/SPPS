##############################################################################################
#              
# NAME: SPSiteGroups.psm1 
# PURPOSE: 
#	Manage Site Groups (Creation, Deletion, Update)
#	Relies on an XML Configuration file for Site Group description.
#	See SPSiteGroups.xml for Schema
#
##############################################################################################


<#
.SYNOPSIS
	Parse the file siteGroupsXML XML object and initiate the Site Group customization
	
.DESCRIPTION
	Will parse the file siteGroupsXML XML object making the difference between the groups to add (AddGroup node),
	the groups to be updated (UpdateGroup node) and the groups to be deleted (DeleteGroup node)

.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER siteGroupsXML
	XML object of Site Groups to manage.
	
.EXAMPLE
	browseAndParseSiteGroupsXML -siteURL <SiteURL> -siteGroupsXML <siteColumnsDescriptionXML>
	
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
function browseAndParseSiteGroupsXML()
{
	Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[XML]$SiteGroupsXML
	)
	
	$functionName = $MyInvocation.MyCommand.Name	
	Write-Debug "Entering $functionName"
	
	try
	{		
		
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if ($curWeb -ne $null)
		{
		
			$SPGroupsNode = $SiteGroupsXML.SelectNodes("/SPSiteGroups")
			if($SPGroupsNode -ne $null -and $SPGroupsNode.HasChildNodes)
			{
				################################
				####### Groups to create #######
				################################
				Write-Debug "[$functionName] About to call browseAndCreateSiteGroupsXML()."
				$grpToAddXML =  $SPGroupsNode.AddGroup
				if($grpToAddXML -ne $null -and $grpToAddXML.HasChildNodes)
				{									
					browseAndCreateSiteGroupsXML -siteURL $siteURL -GroupsToAddXML $grpToAddXML
				}
				else
				{
					Write-Warning "[$functionName] 'AddGroup' node is empty."
				}


				################################
				####### Groups to edit #########
				################################
				Write-Debug "[$functionName] About to call browseAndUpdateSiteGroupsXML()."
				$grpToEditXML =  $SPGroupsNode.UpdateGroup
				if($grpToEditXML -ne $null -and $grpToEditXML.HasChildNodes)
				{									
					browseAndUpdateSiteGroupsXML -siteURL $siteURL -GroupsToUpdateXML $grpToEditXML
				}
				else
				{
					Write-Warning "[$functionName] 'EditGroup' node is empty."
				}


				################################
				####### Groups to delete #######
				################################
				Write-Debug "[$functionName] About to call browseAndDeleteSiteGroupsXML()."
				$grpToDeleteXML =  $SPGroupsNode.DeleteGroup
				if($grpToDeleteXML -ne $null -and $grpToDeleteXML.HasChildNodes)
				{									
					browseAndDeleteSiteGroupsXML -siteURL $siteURL -GroupsToDeleteXML $grpToDeleteXML
				}
				else
				{
					Write-Warning "[$functionName] 'DeleteGroup' node is empty."
				}
			
			}
			else
			{
				Write-Warning "[$functionName] 'SPSiteGroups' XML defintion file is empty."
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
}


<#
.SYNOPSIS
	Parse the file GroupsToAddXML XML object and initiate the Site Group creation
	
.DESCRIPTION
	Will parse the file GroupsToAddXML XML object and loop accross all the groups to be created

.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER GroupsToAddXML
	XML object of the file to parse
	
.EXAMPLE
	browseAndCreateSiteGroupXML -siteURL <SiteURL> -GroupsToAddXML <siteColumnsDescriptionXML>
	
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
function browseAndCreateSiteGroupsXML()
{
	Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[XML.XmlElement]$GroupsToAddXML
	)
	
	$functionName = $MyInvocation.MyCommand.Name	
	Write-Debug "Entering $functionName"
	try
	{		
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if ($curWeb -ne $null)
		{
			if($GroupsToAddXML -ne $null -and $GroupsToAddXML.HasChildNodes)
			{
				foreach($groupToAdd in $GroupsToAddXML.Group)
				{

					$groupName = $groupToAdd.Name
					$groupOwner = $groupToAdd.GroupOwner
					$groupDefaultUser = $groupToAdd.DefaultUser

					$groupDescription = $groupToAdd.Description
					$EditMembership = $groupToAdd.EditMembership
					$ViewMembership = $groupToAdd.ViewMembership
					$JoinLeaveRequest = $groupToAdd.JoinLeaveRequest

					$doesGroupExist= existGroup -siteURL $siteURL -groupName $groupName
					if($doesGroupExist -eq $false)
					{
						
						#---- Create Group.
						if($groupName -ne $null)
						{
							Write-Host "[$functionName] About to call createGroup() for '$groupName'." -ForegroundColor Cyan `r
							createGroup -siteUrl $siteURL -groupName $groupName -groupOwner $groupOwner -defaultUser $groupDefaultUser
						}
						else
						{
							Write-Warning "[$functionName] 'GroupName' is empty. Impossible to create group"
						}
						
						#---- Set Group Description.
						if($groupDescription -ne $null)
						{
							Write-Host "[$functionName] About to call setDescription() on '$groupName'." -ForegroundColor Cyan `r
							setDescription -siteURL $siteURL -groupName $groupName -description $groupDescription
						}
						else
						{
							Write-Warning "[$functionName] 'Description' setting for '$groupName' is empty."
						}

						#---- Set Group Edit Membership for group.
						if($EditMembership -ne $null)
						{
							Write-Host "[$functionName] About to call setEditMembership() on '$groupName'." -ForegroundColor Cyan `r
							setEditMembership -siteURL $siteURL -groupName $groupName -editMembership  $EditMembership
						}
						else
						{
							Write-Warning "[$functionName] 'EditMembership' setting for '$groupName' is empty."
						}

						#---- Set View Membership for group.
						if($ViewMembership -ne $null)
						{
							Write-Host "[$functionName] About to call setViewMembership() on '$groupName'." -ForegroundColor Cyan `r
							setViewMembership -siteURL $siteURL -groupName $groupName -viewMembership  $ViewMembership
						}
						else
						{
							Write-Warning "[$functionName] 'ViewMembership' setting for '$groupName' is empty."
						}

						#---- Set Join Leave Request for group.
						if($JoinLeaveRequest -ne $null)
						{
							Write-Host "[$functionName] About to call setJoinLeaveRequest() on '$groupName'." -ForegroundColor Cyan `r
							$joinLeaveRequest = [System.Convert]::ToBoolean($JoinLeaveRequest)
							setJoinLeaveRequest -siteURL $siteURL -groupName $groupName -allowRequestToJoinLeave $joinLeaveRequest
						}
						else
						{
							Write-Warning "[$functionName] 'JoinLeaveRequest' setting for '$groupName' is empty."
						}
					}					
					else
					{
						Write-Warning "[$functionName] Group '$groupName' already exists in site '$siteURL'."
						Write-Warning "Please use browseAndUpdateSiteGroupXML method."
					}
				} #foreach
			}
			else
			{
				Write-Warning "[$functionName] 'AddGroup' XML node is empty."
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
	Parse the file GroupsToUpdateXML XML object and initiate the Site Group update
	
.DESCRIPTION
	Will parse the file GroupsToUpdateXML XML object and loop accross all the groups to be updated

.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER GroupsToUpdateXML
	XML object of the file to parse
	
.EXAMPLE
	browseAndUpdateSiteGroupsXML -siteURL <SiteURL> -GroupsToUpdateXML <GroupsToUpdateXML>
	
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
function browseAndUpdateSiteGroupsXML()
{
	Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[XML.XmlElement]$GroupsToUpdateXML
	)
	
	$functionName = $MyInvocation.MyCommand.Name	
	Write-Debug "Entering $functionName"
	
	try
	{				
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if ($curWeb -ne $null)
		{
			if($GroupsToUpdateXML -ne $null -and $GroupsToUpdateXML.HasChildNodes)
			{
				foreach($groupToEdit in $GroupsToUpdateXML.Group)
				{

					$groupName = $groupToEdit.Name
					$groupNewName = $groupToEdit.NewName
					$groupOwner = $groupToEdit.Owner					
					$groupDefaultUser = $groupToEdit.DefaultUser
					$groupDescription = $groupToEdit.Description
					$EditMembership = $groupToEdit.EditMembership
					$ViewMembership = $groupToEdit.ViewMembership
					$JoinLeaveRequest = $groupToEdit.JoinLeaveRequest

					$doesGroupExist= existGroup -siteURL $siteURL -groupName $groupName
					if($doesGroupExist -eq $true)
					{									
						#---- Rename Group.
						if($groupNewName -ne $null)
						{
							Write-Host "[$functionName] About to call renameGroup() on '$groupName'." -ForegroundColor Cyan `r
							renameGroup -siteURL $siteURL -groupName $groupName -newGroupName $groupNewName
						}
						else
						{
							Write-Warning "[$functionName] 'NewName' setting for '$groupName' is empty."
						}
						
						#---- Set Group Description.
						if($groupDescription -ne $null)
						{
							Write-Host "[$functionName] About to call setDescription() on '$groupName'." -ForegroundColor Cyan `r
							setDescription -siteURL $siteURL -groupName $groupName -description $groupDescription
						}
						else
						{
							Write-Warning "[$functionName] 'Description' setting for '$groupName' is empty."
						}

						#---- Set Group Edit Membership for group.
						if($EditMembership -ne $null)
						{
							Write-Host "[$functionName] About to call setEditMembership() on '$groupName'." -ForegroundColor Cyan `r
							setEditMembership -siteURL $siteURL -groupName $groupName -editMembership  $EditMembership
						}
						else
						{
							Write-Warning "[$functionName] 'EditMembership' setting for '$groupName' is empty."
						}

						#---- Set View Membership for group.
						if($ViewMembership -ne $null)
						{
							Write-Host "[$functionName] About to call setViewMembership() on '$groupName'." -ForegroundColor Cyan `r
							setViewMembership -siteURL $siteURL -groupName $groupName -viewMembership  $ViewMembership
						}
						else
						{
							Write-Warning "[$functionName] 'ViewMembership' setting for '$groupName' is empty."
						}

						#---- Set Join Leave Request for group.
						if($JoinLeaveRequest -ne $null)
						{
							Write-Host "[$functionName] About to call setJoinLeaveRequest() on '$groupName'." -ForegroundColor Cyan `r
							$joinLeaveRequest = [System.Convert]::ToBoolean($JoinLeaveRequest)
							setJoinLeaveRequest -siteURL $siteURL -groupName $groupName -allowRequestToJoinLeave $joinLeaveRequest
						}
						else
						{
							Write-Warning "[$functionName] 'JoinLeaveRequest' setting for '$groupName' is empty."
						}
					}					
					else
					{
						Write-Warning "[$functionName] Group '$groupName' does not exist in site '$siteURL'."
						Write-Warning "Please use browseAndCreateSiteGroupXML method."
					}
				}#foreach
			}
			else
			{
				Write-Warning "[$functionName] 'EditGroup' XML node is empty."
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
	Parse the file GroupsToDeleteXML XML object and initiate the Site Group deletion
	
.DESCRIPTION
	Will parse the file GroupsToDeleteXML XML object and loop accross all the groups to be deleted

.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER GroupsToDeleteXML
	XML object of the file to parse
	
.EXAMPLE
	browseAndDeleteSiteGroupsXML -siteURL <SiteURL> -GroupsToDeleteXML <GroupsToDeleteXML>
	
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
function browseAndDeleteSiteGroupsXML()
{
	Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[XML.XmlElement]$GroupsToDeleteXML
	)
	
	$functionName = $MyInvocation.MyCommand.Name	
	Write-Debug "Entering $functionName"
	
	try
	{		
		
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if ($curWeb -ne $null)
		{		
			if($GroupsToDeleteXML -ne $null -and $GroupsToDeleteXML.HasChildNodes)
			{
				foreach($groupToDelete in $GroupsToDeleteXML.Group)
				{
					$groupName = $groupToDelete.Name				
					Write-Host "[$functionName] About to delete '$groupName' in site '$siteURL'." -ForegroundColor Cyan `r
					
					$doesGroupExist= existGroup -siteURL $siteURL -groupName $groupName
					if($doesGroupExist -eq $true)
					{									
						deleteGroup -siteURL $siteURL -groupName $groupName
					}					
					else
					{
						Write-Warning "[$functionName] Group '$groupName' does not exist in site '$siteURL'."
						Write-Warning "Impossible to delete an unexisting group."
					}
				}#foreach
			}
			else
			{
				Write-Warning "[$functionName] 'DeleteGroup' XML node is empty."
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
	Check if the group $groupName exists in site $siteURL
	
.DESCRIPTION
	Will check if the SharePoint group $groupName exists in site $siteURL
	Return true if the SharePoint Group exists, false instead

.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER groupName
	Name of the SharePoint Group
	
.EXAMPLE
	existGroup -siteURL <SiteURL> -groupName <groupName>
	
.OUTPUTS
	True, if exists
	False, if not

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 01.03.2017
	Last Updated by: JBO
	Last Updated: 01.03.2017
#>
function existGroup()
{
	param
	(
        [Parameter(Mandatory=$true, Position=1)]
	    [string]$siteURL,
        [Parameter(Mandatory=$true, Position=2)]
	    [string]$groupName
	)

    $functionName = $MyInvocation.MyCommand.Name	
	Write-Debug "Entering $functionName"
	
	try
	{
		$existGroup = $false
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if ($curWeb -ne $null)
		{
			if($curWeb.SiteGroups.GetCollection(@($groupName)).Count -gt 0)
			{
				$existGroup= $true
			}
			else
			{
				$existGroup= $false
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
	return $existGroup
}


<#
.SYNOPSIS
	If exists, return the group $groupName
	
.DESCRIPTION
	Will check if the SharePoint group $groupName exists
	Return the SPGroup object if the SharePoint Group exists, null instead

.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER groupName
	Name of the SharePoint Group
	
.EXAMPLE
	getGroup -siteURL <SiteURL> -groupName <groupName>
	
.OUTPUTS
	<SPGroup> object, if exists
	NULL, if not

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 01.03.2017
	Last Updated by: JBO
	Last Updated: 01.03.2017
#>
function getGroup()
{
	param
	(
        [Parameter(Mandatory=$true, Position=1)]
	    [string]$siteURL,
        [Parameter(Mandatory=$true, Position=2)]
	    [string]$groupName
	)
    	
    $functionName = $MyInvocation.MyCommand.Name	
	Write-Debug "Entering $functionName"

	try
	{
		$group= $null
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if ($curWeb -ne $null)
		{
			if($curWeb.SiteGroups.GetCollection(@($groupName)).Count -gt 0)
			{
				$group = $curWeb.SiteGroups[$groupName]
			}
			else
			{
				$group = $null
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
	return $group
}


<#
.SYNOPSIS
	Create the SharePoint group
	
.DESCRIPTION
	Create the SharePoint group

.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER groupName
	Name of the SharePoint Group	
	
.PARAMETER groupOwner
	Username of the Owner of the SharePoint group
	
.PARAMETER defaultUser
	Username of the Default User of the SharePoint group
		
.PARAMETER groupDescription
	(Optional) Description of the SharePoint Group
	
.EXAMPLE
	createGroup -siteURL <SiteURL> -groupName <groupName> -groupOwner <groupOwner> -defaultUser <defaultUser> -groupDescription <groupDescription>
	
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
function createGroup()
{
	param
	(
	    [Parameter(Mandatory=$true, Position=1)]
	    [string]$siteURL,
        [Parameter(Mandatory=$true, Position=2)]
	    [string]$groupName,
        [Parameter(Mandatory=$true, Position=3)]
	    [string]$groupOwner,
        [Parameter(Mandatory=$true, Position=4)]
	    [string]$defaultUser,
        [Parameter(Mandatory=$false, Position=5)]
	    [string]$groupDescription

	)
	
	$functionName = $MyInvocation.MyCommand.Name	
	Write-Debug "Entering $functionName"
	
	try
	{
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if ($curWeb -ne $null)
		{
			#----------------------------------------------------
			# DEBUG* Create group
			#----------------------------------------------------
			# Write-Debug "[createGroup]Parameter siteURL $siteURL."
			# Write-Debug "[createGroup]Parameter groupName $groupName."
			# Write-Debug "[createGroup]Parameter groupOwner $groupOwner." 
			# Write-Debug "[createGroup]Parameter defaultUser $defaultUser." 

			$doesGroupExist= existGroup -siteURL $siteURL -groupName $groupName
			if($doesGroupExist -eq $false)
			{
				$ownerGroupExist = existGroup  -siteURL $siteURL -groupName $groupOwner
				if($ownerGroupExist -eq $false)
				{
					#groupOwner is a User
					Write-Host "[$functionName] Owner '$groupOwner' is a User." 
					$grpOwner = ensureClaimsUser -siteURL $siteURL -userName $groupOwner
					Write-Host "[$functionName] Owner '$groupOwner' converted to claims: '$grpOwner'." 
				}
				else
				{
					# groupOwner is a Group
					Write-Host "[$functionName] Owner $groupOwner is a Group."
					$grpOwner = $curWeb.SiteGroups[$groupOwner]
				}		
		
				#old: $groupOwnerDefaultUser = $curWeb.EnsureUser($defaultUser)
				$groupOwnerDefaultUser = ensureClaimsUser -siteURL $siteURL -userName $defaultUser
				if($groupOwnerDefaultUser -ne $null)
				{
					if($grpOwner -ne $null)
					{
						Write-Host "[$functionName] About to create Group '$groupName' on site '$siteURL'." -ForegroundColor Magenta `r 
						$curWeb.SiteGroups.Add($groupName, $grpOwner, $groupOwnerDefaultUser, $groupDescription) 
						Write-Host "[$functionName] Group '$groupName' sucessfully created on site '$siteURL'." -ForegroundColor Green `r 
					}
					else
					{
						Write-Warning "[$functionName] GroupOwner '$groupOwner' is null."	
					}
				
				}
				else
				{
					Write-Warning "[$functionName] DefaultUser '$defaultUser' is  null."
				}			
			}
			else
			{
				Write-Warning "[$functionName] Group '$groupName' already exists." 
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
        if($curWeb -ne $Null)
		{
			$curWeb.Dispose()
		}
		Write-Debug "Exiting $functionName"
    }
}


<#
.SYNOPSIS
	Delete the SharePoint group
	
.DESCRIPTION
	Delete the SharePoint group

.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER groupName
	Name of the SharePoint Group	
	
.EXAMPLE
	deleteGroup -siteURL <SiteURL> -groupName <groupName>
	
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
function deleteGroup()
{
	param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[string]$groupName
	)    
	
	$functionName = $MyInvocation.MyCommand.Name	
	Write-Debug "Entering $functionName"
	
	try
	{
        $curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if ($curWeb -ne $null)
		{
			$doesGroupExist= existGroup -siteURL $siteURL -groupName $groupName
			#if($curWeb.SiteGroups.GetCollection(@($groupName)).Count -gt 0)
			if($doesGroupExist -eq $true)
			{
				Write-Host "[$functionName] About to delete Group '$groupName' on site '$siteURL'." -ForegroundColor Magenta `r 
				$curWeb.SiteGroups.Remove($groupName)
				$curWeb.Update()
				Write-Host "[$functionName] Group '$groupName' has been deleted on site '$siteURL'." -ForegroundColor Green `r 
			}
			else
			{
				Write-Warning "[$functionName] Group '$groupName' does not exist."
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
	Add a list of Users to a SharePoint Group
	
.DESCRIPTION
	Add a list of Users to a SharePoint Group

.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER groupName
	Name of the SharePoint Group		
	
.PARAMETER usersToAddList
	Users to be added to the SharePoint group (semicolon separated)	
	
.EXAMPLE
	addUsersToGroup -siteURL <SiteURL> -groupName <groupName> -usersToAddList <usersToAddList>
	
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
function addUsersToGroup()
{
	param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[string]$groupName,
		[Parameter(Mandatory=$true, Position=3)]
		[string]$usersToAddList 
	)

    $functionName = $MyInvocation.MyCommand.Name	
	Write-Debug "Entering $functionName"
	
	try
	{
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if ($curWeb -ne $null)
		{
			$doesGroupExist= existGroup -siteURL $siteURL -groupName $groupName
			if($doesGroupExist -eq $true)
			{
				[string[]]$sepratedUsers =$usersToAddList.Split(";")
				foreach($userToAdd in $sepratedUsers)
				{ 
					if(-not[string]::IsNullOrEmpty($userToAdd))
					{
						addUserToGroup -siteURL $siteURL -groupName $groupName -userName $userToAdd
					}
				}    
			} 
			else
			{
				Write-Warning "[$functionName] Group '$groupName' does not exist on site."
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
        if($curWeb -ne $Null)
        {
			$curWeb.Dispose()
        }
		Write-Debug "Exiting $functionName"
    }
}


<#
.SYNOPSIS
	Add a specific User to a SharePoint Group
	
.DESCRIPTION
	Add a specific User to a SharePoint Group

.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER groupName
	Name of the SharePoint Group		
	
.PARAMETER userName
	Username to add (domain\userName)
	
.EXAMPLE
	addUserToGroup -siteURL <SiteURL> -groupName <groupName> -userName <userName>
	
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
function addUserToGroup()
{
	param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[string]$groupName,
		[Parameter(Mandatory=$true, Position=3)]
		[string]$userName 
	)
    
	$functionName = $MyInvocation.MyCommand.Name	
	Write-Debug "Entering $functionName"
	
	try
	{
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if ($curWeb -ne $null)
		{
			$doesGroupExist= existGroup -siteURL $siteURL -groupName $groupName
			if($doesGroupExist -eq $true)
			{
				#Safety Measure (1) : Always Capslock !!
				#$userName = $userName.ToUpper()
				#Safety Measure (1) : Ensure User
				#$spUserToAdd = $curWeb.EnsureUser($userName)
				$spUserToAdd = ensureClaimsUser -siteURL $siteURL -userName $userName

				if($spUserToAdd -ne $null)
				{
					$doesUserExist = $curWeb.sitegroups[$groupName].Users[$spUserToAdd]
					if($doesUserExist -eq $null)
					{
						$curWeb.sitegroups[$groupName].AddUser($spUserToAdd)
						Write-Host "[$functionName] User '$spUserToAdd' has been added to $groupName group with AddUser()." -ForegroundColor Green `r  
					}
					else
					{
						Write-Warning "[$functionName] User '$spUserToAdd' already exists in $groupName group." 
					}
				}
				else
				{
					$spUserToAdd= Set-SPUser -web $curWeb -Identity $spUserToAdd -Group $groupName
					Write-Host "[$functionName] User '$spUserToAdd' has been added to $groupName group with Set-SPUser." -ForegroundColor Green `r  
				}    
			} 
			else
			{
				Write-Warning "[$functionName] Group '$groupName' does not exist on site." 
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
        if($curWeb -ne $Null)
        {
            $curWeb.Dispose()
        }
		Write-Debug "Exiting $functionName"
    }
}


<#
.SYNOPSIS
	Remove a list of Users from a SharePoint Group
	
.DESCRIPTION
	Remove a list of Users from a SharePoint Group

.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER groupName
	Name of the SharePoint Group		
	
.PARAMETER usersToRemoveList
	Users to be deleted from the SharePoint group (semicolon separated)	
	
.EXAMPLE
	removeUsersFromGroup -siteURL <SiteURL> -groupName <groupName> -usersToRemoveList <usersToRemoveList>
	
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
function removeUsersFromGroup()
{
	param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[string]$groupName,
		[Parameter(Mandatory=$true, Position=3)]
		[string]$usersToRemoveList
	)

    $functionName = $MyInvocation.MyCommand.Name	
	Write-Debug "Entering $functionName"

	try
	{
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if ($curWeb -ne $null)
		{
			$doesGroupExist= existGroup -siteURL $siteURL -groupName $groupName
			if($doesGroupExist -eq $true)
			{
				[string[]]$sepratedUsers =$usersToRemoveList.Split(";")
				foreach($userToRemove in $sepratedUsers)
				{ 				
					if (-not[string]::IsNullOrEmpty($userToRemove))
					{
						removeUserFromGroup -siteURL $siteURL -groupName $groupName -userName $userToRemove
					}			
				} 
			} 
			else
			{
				Write-Warning "[$functionName] Group '$groupName' does not exist on site."
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
		if($curWeb -ne $Null)
		{
			$curWeb.Dispose()
		}
		Write-Debug "Exiting $functionName"
	}
	
} 


<#
.SYNOPSIS
	Remove a specific User from a SharePoint Group
	
.DESCRIPTION
	Remove a specific User from a SharePoint Group

.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER groupName
	Name of the SharePoint Group		
	
.PARAMETER userName
	Username to add (domain\userName)
	
.EXAMPLE
	removeUserFromGroup -siteURL <SiteURL> -groupName <groupName> -userName <userName>
	
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
function removeUserFromGroup()
{
	param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[string]$groupName,
		[Parameter(Mandatory=$true, Position=3)]
		[string]$userName
	)
	
	$functionName = $MyInvocation.MyCommand.Name	
	Write-Debug "Entering $functionName"
	
	try
	{
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if ($curWeb -ne $null)
		{
			$doesGroupExist= existGroup -siteURL $siteURL -groupName $groupName
			if($doesGroupExist -eq $true)
			{
				#Safety Measure (1) : Always Capslock !!
				#$userName = $userName.ToUpper()
				#Safety Measure (1) : Always Capslock !!
				$spUserToRemove = ensureClaimsUser -siteURL $siteURL -userName $userName

				if($spUserToRemove -ne $null)
				{
					$doesUserExist = $curWeb.sitegroups[$groupName].Users[$spUserToRemove]
					if($doesUserExist -ne $null)
					{
						$curWeb.sitegroups[$groupName].RemoveUser($spUserToRemove)
						Write-Host "[$functionName] User '$userName' has been Removed from $groupName group." -ForegroundColor Green `r  
					}
					else
					{
						Write-Warning "[$functionName] User '$userName' does not exist in $groupName group." 
					}
				}
				else
				{
					Write-Warning "[$functionName] User '$userName' does not exist on site." 
				}
			} 
			else
			{
				Write-Warning "[$functionName] Group '$groupName' does not exist on site." 
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
		if($curWeb -ne $Null)
		{
			$curWeb.Dispose()
		}
		Write-Debug "Exiting $functionName"
	}
	
} 


<#
.SYNOPSIS
	Change who (Owners or Members) is capable of editing group membership 
	
.DESCRIPTION
	Change who (Owners or Members) is capable of editing group membership 

.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER groupName
	Name of the SharePoint Group		
	
.PARAMETER editMembership
	Who will be capable of changing the group (Possible values: Owners or Members)
	
.EXAMPLE
	setEditMembership -siteURL <SiteURL> -groupName <groupName> -editMembership <editMembership>
	
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
function setEditMembership()
{
	param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[string]$groupName,
		[Parameter(Mandatory=$true, Position=3)]
		[string]$editMembership
	)

	$functionName = $MyInvocation.MyCommand.Name	
	Write-Debug "Entering $functionName"
	
	try
	{
		# For String Casting/Testing
		$membersLowerCase = "members";
		$ownersLowerCase = "owners";
		$editMembershipLowerCase = $editMembership.ToLower();

		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if ($curWeb -ne $null)
		{		
			Write-Debug "[$functionName] PARAM: groupName: '$groupName'."
			Write-Debug "[$functionName] PARAM: editMembership: '$editMembership'."
			Write-Debug "[$functionName] PARAM: editMembership.LowerCase: '$editMembershipLowerCase'."

			$group= getGroup -siteURL $siteURL -groupName $groupName
			if($group -ne $null)
			{
				if (($editMembershipLowerCase -ne $membersLowerCase) -and ($editMembershipLowerCase -ne $ownersLowerCase) )
				{
					Write-Warning "[$functionName] '$editMembership' is not a recognized value."
				}
				else
				{
					if ($editMembership.ToLower() -eq "Members".ToLower())
					{				
						$group.AllowMembersEditMembership = $true;
					}
					elseif($editMembership.ToLower() -eq "Owners".ToLower())
					{
						$group.AllowMembersEditMembership = $false;
					}

					Write-Host "[$functionName] About to allow '$editMembership' to edit membership of the group '$groupName'." -ForegroundColor Magenta `r
					$group.Update()
					Write-Host "[$functionName] '$editMembership' allowed to edit membership of the group '$groupName'." -ForegroundColor Green `r
				}			
			} 
			else
			{
				Write-Warning "[$functionName] Group '$groupName' does not exist on site." 
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
		if($curWeb -ne $Null)
		{
			$curWeb.Dispose()
		}
		Write-Debug "Exiting $functionName"
	}
	
} 


<#
.SYNOPSIS
	Change who (Everyones or Members) is capable of viewing group membership 
	
.DESCRIPTION
	Change who (Everyones or Members) is capable of viewing group membership 

.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER groupName
	Name of the SharePoint Group		
	
.PARAMETER viewMembership
	Who will be capable of viewing the group members (Possible values: Everyone or Members)
	
.EXAMPLE
	setViewMembership -siteURL <SiteURL> -groupName <groupName> -viewMembership <viewMembership>
	
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
function setViewMembership()
{
	param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[string]$groupName,
		[Parameter(Mandatory=$true, Position=3)]
		[string]$viewMembership
	)

	$functionName = $MyInvocation.MyCommand.Name	
	Write-Debug "Entering $functionName"
	
	try
	{
		# For String Casting/Testing
		$membersLowerCase = "members";
		$everyoneLowerCase = "everyone";
		$viewMembershipLowerCase = $viewMembership.ToLower();

		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if ($curWeb -ne $null)
		{
			#Write-Host "[$functionName] PARAM: groupName: '$groupName'."
			#Write-Host "[$functionName] PARAM: viewMembership: '$viewMembership'."
			#Write-Host "[$functionName] PARAM: viewMembership.LowerCase: '$viewMembershipLowerCase'."

			$group= getGroup -siteURL $siteURL -groupName $groupName
			if($group -ne $null)
			{
				if (($viewMembershipLowerCase -ne $membersLowerCase) -and ($viewMembershipLowerCase -ne $everyoneLowerCase) )
				{
					Write-Warning "[$functionName] '$viewMembership' is not a recognized value."
				}
				else
				{
					if ($viewMembership.ToLower() -eq "Members".ToLower())
					{				
						$group.OnlyAllowMembersViewMembership = $true;
					}
					elseif($viewMembership.ToLower() -eq "Everyone".ToLower())
					{
						$group.OnlyAllowMembersViewMembership = $false;
					}

					Write-Host "[$functionName] About to allow '$viewMembership' to view membership of the group '$groupName'." -ForegroundColor Magenta `r
					$group.Update()
					Write-Host "[$functionName] '$viewMembership' allowed to view membership of the group '$groupName'." -ForegroundColor Green `r
				}			
			} 
			else
			{
				Write-Warning "[$functionName] Group '$groupName' does not exist on site." 
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
		if($curWeb -ne $Null)
		{
			$curWeb.Dispose()
		}
		Write-Debug "Exiting $functionName"
	}	
} 


<#
.SYNOPSIS
	Allow users to request membership in the group and to allow users to request to leave the group
	
.DESCRIPTION
	Allow users to request membership in the group and to allow users to request to leave the group

.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER groupName
	Name of the SharePoint Group		
	
.PARAMETER allowRequestToJoinLeave
	true to allow users to request membership in the group or request to leave the group; otherwise, false.
	
.EXAMPLE
	setJoinLeaveRequest -siteURL <SiteURL> -groupName <groupName> -allowRequestToJoinLeave <TRUE|FALSE>
	
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
function setJoinLeaveRequest()
{
	param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[string]$groupName,
		[Parameter(Mandatory=$true, Position=3)]
		[boolean]$allowRequestToJoinLeave 
	)

	$functionName = $MyInvocation.MyCommand.Name	
	Write-Debug "Entering $functionName"
	
	try
	{
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if ($curWeb -ne $null)
		{
			#Write-Host "[$functionName] PARAM: groupName: '$groupName'."
			#Write-Host "[$functionName] PARAM: allowRequestToJoinLeave: '$allowRequestToJoinLeave'."

			$group= getGroup -siteURL $siteURL -groupName $groupName
			if($group -ne $null)
			{
				Write-Host "[$functionName] About to set AllowRequestToJoinLeave to '$allowRequestToJoinLeave' for the group '$groupName'." -ForegroundColor Magenta `r
				$group.AllowRequestToJoinLeave = $allowRequestToJoinLeave 
				$group.Update()
				Write-Host "[$functionName] AllowRequestToJoinLeave set to '$allowRequestToJoinLeave' for the group '$groupName'." -ForegroundColor Green `r		
			} 
			else
			{
				Write-Warning "[$functionName] Group '$groupName' does not exist on site."
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
		if($curWeb -ne $Null)
		{
			$curWeb.Dispose()
		}
		Write-Debug "Exiting $functionName"
	}	
} 


<#
.SYNOPSIS
	Allow/prevent users to be automatically added or removed when they make a request.
	
.DESCRIPTION
	Allow/prevent users to be automatically added or removed when they make a request.

.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER groupName
	Name of the SharePoint Group		
	
.PARAMETER autoAcceptRequestToJoinLeave
	true to specify that users are automatically added or removed when they make a request; otherwise, false.
	
.EXAMPLE
	setAutoAcceptRequest -siteURL <SiteURL> -groupName <groupName> -autoAcceptRequestToJoinLeave <TRUE|FALSE>
	
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
function setAutoAcceptRequest()
{
	param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[string]$groupName,
		[Parameter(Mandatory=$true, Position=3)]
		[boolean]$autoAcceptRequestToJoinLeave 
	)

	$functionName = $MyInvocation.MyCommand.Name	
	Write-Debug "Entering $functionName"
	
	try
	{
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if ($curWeb -ne $null)
		{
			#Write-Host "[$functionName] PARAM: groupName: '$groupName'."
			#Write-Host "[$functionName] PARAM: autoAcceptRequestToJoinLeave: '$autoAcceptRequestToJoinLeave'."

			$group= getGroup -siteURL $siteURL -groupName $groupName
			if($group -ne $null)
			{
				Write-Host "[$functionName] About to set AutoAcceptRequestToJoinLeave to '$autoAcceptRequestToJoinLeave' for the group '$groupName'." -ForegroundColor Magenta `r
				$group.AutoAcceptRequestToJoinLeave  = $autoAcceptRequestToJoinLeave 
				$group.Update()
				Write-Host "[$functionName] AutoAcceptRequestToJoinLeave set to '$autoAcceptRequestToJoinLeave' for the group '$groupName'." -ForegroundColor Green `r		
			} 
			else
			{
				Write-Warning "[$functionName] Group '$groupName' does not exist on site."
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
		if($curWeb -ne $Null)
		{
			$curWeb.Dispose()
		}
		Write-Debug "Exiting $functionName"
	}
} 

<#
.SYNOPSIS
	Rename the group
	
.DESCRIPTION
	Rename the group

.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER groupName
	Name of the SharePoint Group		
	
.PARAMETER newGroupName
	New name of the SharePoint Group	
	
.EXAMPLE
	renameGroup -siteURL <SiteURL> -groupName <groupName> -newGroupName <newGroupName>
	
.OUTPUTS
	None
	
.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 01.03.2017
	Last Updated by: JBO
	Last Updated: 10.12.2020
#>
function renameGroup()
{
	
	param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[string]$groupName,
		[Parameter(Mandatory=$true, Position=3)]
		[string]$newGroupName
	)

	$functionName = $MyInvocation.MyCommand.Name	
	Write-Debug "Entering $functionName"
	
	
	try
	{
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if ($curWeb -ne $null)
		{
			Write-Debug "[$functionName] PARAM: groupName: '$groupName'." 
			Write-Debug "[$functionName] PARAM: description: '$description'."

			$group= getGroup -siteURL $siteURL -groupName $groupName
			if($group -ne $null)
			{
				Write-Host "[$functionName] About to rename Group '$groupName' to '$newGroupName'." -ForegroundColor Magenta `r				
				$curWeb.AllowUnsafeUpdates = $true;
				$group.Name = $newGroupName;
				$group.Update();
				$curWeb.Update();
				Write-Host "[$functionName] Group '$groupName' has been renamed to '$newGroupName'." -ForegroundColor Green `r		

				$curWeb.AllowUnsafeUpdates = $false;
			} 
			else
			{
				Write-Warning "[$functionName] Group '$groupName' does not exist on site."
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
		if($curWeb -ne $Null)
		{
			$curWeb.Dispose()
		}
		Write-Debug "Exiting $functionName"
	}	
}

<#
.SYNOPSIS
	Set the group description
	
.DESCRIPTION
	Set the group description

.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER groupName
	Name of the SharePoint Group		
	
.PARAMETER description
	Description of the SharePoint Group	
	
.EXAMPLE
	setDescription -siteURL <SiteURL> -groupName <groupName> -description <description>
	
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
function setDescription()
{
	param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[string]$groupName,
		[Parameter(Mandatory=$true, Position=3)]
		[string]$description
	)

	$functionName = $MyInvocation.MyCommand.Name	
	Write-Debug "Entering $functionName"
	
	try
	{
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if ($curWeb -ne $null)
		{
			Write-Debug "[$functionName] PARAM: groupName: '$groupName'." 
			Write-Debug "[$functionName] PARAM: description: '$description'."

			$group= getGroup -siteURL $siteURL -groupName $groupName
			if($group -ne $null)
			{
				Write-Host "[$functionName] About to change Group '$groupName' description." -ForegroundColor Magenta `r				
				$curWeb.AllowUnsafeUpdates = $true;
				$descriptionField = [Microsoft.SharePoint.SPFieldMultiLineText]$curWeb.SiteUserInfoList.Fields[[Microsoft.SharePoint.SPBuiltInFieldId]::Notes];
				$groupItem = $curWeb.SiteuserInfoList.GetItemById($group.ID);
				$groupItem[$descriptionField.InternalName] = $description;
				$groupItem.Update();
				$group.Update()
				$curWeb.Update()
				Write-Host "[$functionName] Group '$groupName' description has been updated." -ForegroundColor Green `r		

				$curWeb.AllowUnsafeUpdates = $false;
			} 
			else
			{
				Write-Warning "[$functionName] Group '$groupName' does not exist on site."
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
		if($curWeb -ne $Null)
		{
			$curWeb.Dispose()
		}
		Write-Debug "Exiting $functionName"
	}	
}