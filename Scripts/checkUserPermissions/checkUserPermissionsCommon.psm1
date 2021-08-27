<#
.SYNOPSIS
	Get the Permissions informations for a user ($userName) on a specific SharePoint Securable object ($object)
	
.DESCRIPTION
	Get the Permissions informations for a user ($userName) on a specific SharePoint Securable object ($object)
	Will also mention explicitly if the user has NO ACCESS at all on a SPWeb or a SPList
	
.PARAMETER siteUrl
	(Mandatory) URL of the SharePoint Site
	
.PARAMETER object
	(Mandatory) SharePoint Securable object

.EXAMPLE
	getPermissionInfo -siteUrl <siteURL> -object <object>
	
.OUTPUTS
	Array $permissionDataCollection containing user permissions on the SharePoint Securable object	(if any)  
#>
function getPermissionInfo()
{
	param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$userName,
		[Parameter(Mandatory=$true, Position=2)]
		[Microsoft.SharePoint.SPSecurableObject]$object
	)
	
	#Object Array to hold Permission data
	$permissionDataCollection = @()

	#Determine the given Object type and Get URL of it
	switch($object.GetType().FullName)
	{
		"Microsoft.SharePoint.SPWeb"  { $objectType = "SPWeb" ; $objectURL = $object.URL }
		"Microsoft.SharePoint.SPListItem" {
			if($object.Folder -ne $null)
			{
				$objectType = "Folder" ; $objectURL = "$($object.Web.Url)/$($object.Url)"
			}
			else
			{
				$objectType = "Item"; $objectURL = "$($object.Web.Url)/$($object.Url)"
			}
		}
		#Microsoft.SharePoint.SPList, Microsoft.SharePoint.SPDocumentLibrary, "Microsoft.SharePoint.SPPictureLibrary",etc
		default { $objectType = "List"; $objectURL = "$($object.ParentWeb.Url)/$($object.RootFolder.URL)" }
	}
  
	#Write-Host "ObjectTitle: '$($object.Title)' `nObjectType: '$objectType'" -Foregroundcolor Cyan `r
	
	#Get Permissions of the user on given object - Such as: Web, List, Folder, ListItem
	$userPermissionInfo = $object.GetUserEffectivePermissionInfo($userName)
	if ($userPermissionInfo.RoleAssignments -ne $null)
	{
		
		#Iterate through each permission and get the details
		foreach($assignment in $userPermissionInfo.RoleAssignments)
		{
			$userPermissions = ($assignment.RoleDefinitionBindings | Where-Object { $_.Name -ne "Limited Access"} | Select-Object -ExpandProperty name ) -join "," 
			
			if($assignment.Member -is [Microsoft.SharePoint.SPGroup])  
			{ 					
				$permissionType = "Group" 
				$identity = $assignment.Member.Name    
			} 
			else
			{
				$permissionType = "User"	
				$identity = $userName
			}
		
			if ($assignment.RoleDefinitionBindings.Name -ne "Limited Access")
			{
				#Create an object to hold storage data
				$permissionData = New-Object PSObject
				Add-Member -inputObject $permissionData -memberType NoteProperty -name "Object" -value $objectType
				Add-Member -inputObject $permissionData -memberType NoteProperty -name "Name" -value $object.Title
				Add-Member -inputObject $permissionData -memberType NoteProperty -name "URL" -value $objectURL 
				Add-Member -inputObject $permissionData -memberType NoteProperty -name "Perm. Type" -value $permissionType
				Add-Member -inputObject $permissionData -memberType NoteProperty -name "Perm. Identity" -value $identity	
				Add-Member -inputObject $permissionData -memberType NoteProperty -name "Perm. Level" -value $userPermissions		
				
				$permissionDataCollection += $permissionData
			}	
		}
	
	}
	else
	{
		#Write-Host "'UserEffectivePermissionInfo' object is null." -Foregroundcolor Red `r
		
		if ( ($objectType -eq "List") -or ($objectType -eq "SPWeb") )
		{
			$permissionData = New-Object PSObject
			Add-Member -inputObject $permissionData -memberType NoteProperty -name "Object" -value $objectType
			Add-Member -inputObject $permissionData -memberType NoteProperty -name "Name" -value $object.Title
			Add-Member -inputObject $permissionData -memberType NoteProperty -name "URL" -value $objectURL 		
			Add-Member -inputObject $permissionData -memberType NoteProperty -name "Perm. Type" -value "No Access"
			Add-Member -inputObject $permissionData -memberType NoteProperty -name "Perm. Identity" -value $userName	
			Add-Member -inputObject $permissionData -memberType NoteProperty -name "Perm. Level" -value "No Access"			
			$permissionDataCollection += $permissionData
		}
	}	
	return $permissionDataCollection
}


<#
.SYNOPSIS
	Check if the user ($username) is member of the Site Collection Administrators group
	
.DESCRIPTION
	Check if the user ($username) is member of the Site Collection Administrators group
	
.PARAMETER siteCollURL
	(Mandatory) URL of the SharePoint Site Collection
	
.PARAMETER userName
	(Mandatory) SharePoint Securable object

.EXAMPLE
	checkSiteCollectionAdministrators -siteCollURL <siteCollURL> -userName <userName>
	
.OUTPUTS
	Array $permList containing user permissions on the Site Collection (if any)
#>
function checkSiteCollectionAdministrators
{
	param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteCollURL,
		[Parameter(Mandatory=$true, Position=2)]
		[string]$userName
	)

	#Object Array to hold Permission data
	$permList = @()

	$site = Get-SPSite $siteCollURL
	if($site -ne $null)
	{
		foreach ($sca in $site.RootWeb.SiteAdministrators)
		{
			if ($sca.UserLogin -eq $userName)
			{
				Write-Host "'$userName' is Site Collection Administrator." -Foregroundcolor Green `r
				Write-Host "Other existing permissions will be overriden by Site Collection Admin rights." -Foregroundcolor Green `r
				$siteName = $site.RootWeb.Title
				$siteURL = $site.RootWeb.URL
				$identity = $userName
				$scope = "SPSite"
				$permType = "Site Collection Administrators"
				
				$permissionData = New-Object PSObject
				Add-Member -inputObject $permissionData -memberType NoteProperty -name "Object" -value $scope
				Add-Member -inputObject $permissionData -memberType NoteProperty -name "Name" -value $siteName
				Add-Member -inputObject $permissionData -memberType NoteProperty -name "URL" -value $siteURL				
				Add-Member -inputObject $permissionData -memberType NoteProperty -name "Perm. Type" -value "Group"
				Add-Member -inputObject $permissionData -memberType NoteProperty -name "Perm. Identity" -value $permType
				Add-Member -inputObject $permissionData -memberType NoteProperty -name "Perm. Level" -value "Full Control"
				# Append permissionData to permList
				$permList += $permissionData	
			}
			else
			{
				#Write-Host "'$userName' is not Site Collection Administrator." -Foregroundcolor Red `r
			}
		}		
	}
	else
	{
		Write-Host "Site Collection '$siteCollURL' does not exist."	
	}
	if($site -ne $null)
	{
		$site.Dispose()
	}
	return $permList;
}



<#
.SYNOPSIS
	Check if the user ($username) is member of a Web-Application Policy
	
.DESCRIPTION
	Check if the user ($username) is member of a Web-Application Policy
	
.PARAMETER webAppURL
	(Mandatory) URL of the SharePoint Web-Application
	
.PARAMETER userName
	(Mandatory) SharePoint Securable object

.EXAMPLE
	checkWebAppUserPolicy -webAppURL <webAppURL> -userName <userName>
	
.OUTPUTS
	Array $permList containing user policy on the Web Application (if any)
#>
function checkWebAppUserPolicy
{
	param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$webAppURL,
		[Parameter(Mandatory=$true, Position=2)]
		[string]$userName
	)
	
	#Object Array to hold Permission data
	$permList = @()

	$webApp = Get-SPWebApplication $webAppURL
	if($webApp -ne $null)
	{
		foreach ($policy in $webApp.Policies)
		{
			if($policy.UserName.EndsWith($userName,1))
			{
				$userPolicies = ($policy.PolicyRoleBindings | Select-Object -ExpandProperty name ) -join "," 
				
				Write-Host "'$userName' is member of a Web-Application Policy." -Foregroundcolor Green `r
				Write-Host "Other existing permissions will be overriden by the policies." -Foregroundcolor Green `r
				$siteName = $webApp.Title
				$siteURL = $webApp.URL
				$scope = "SPWebApplication"
				$permType = "Policy"
				
				$permissionData = New-Object PSObject
				Add-Member -inputObject $permissionData -memberType NoteProperty -name "Object" -value $scope
				Add-Member -inputObject $permissionData -memberType NoteProperty -name "Name" -value $siteName
				Add-Member -inputObject $permissionData -memberType NoteProperty -name "URL" -value $siteURL				
				Add-Member -inputObject $permissionData -memberType NoteProperty -name "Perm. Type" -value $permType
				Add-Member -inputObject $permissionData -memberType NoteProperty -name "Perm. Identity" -value $userName
				Add-Member -inputObject $permissionData -memberType NoteProperty -name "Perm. Level" -value $userPolicies
				# Append permissionData to permList
				$permList += $permissionData				
			}
			else
			{
				#Write-Host "'$userName' is not Site Collection Administrator." -Foregroundcolor Red `r
			}
		}		
	}
	else
	{
		Write-Host "Web-Application '$webAppURL' does not exist."	
	}
	return $permList;
}


<#
.SYNOPSIS
	Get User Permissions on a single SPWeb (and all its child SPList, SPFolder and SPListItem)
	
.DESCRIPTION
	Get User Permissions on a single SPWeb (and all its child SPList, SPFolder and SPListItem)
	
.PARAMETER siteUrl
	(Mandatory) URL of the SharePoint Site
	
.PARAMETER object
	(Mandatory) SharePoint Securable object

.EXAMPLE
	getUserPermissionsOnSPWeb -siteUrl <siteURL> -object <object>
	
.OUTPUTS
	Array $globalPermList containing all User Permissions on the single SPWeb
#>
function getUserPermissionsOnSPWeb()
{
	param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$false, Position=2)]
		[string]$userName
	)
	
	$globalPermList = @() # $permList = @() <=> CSV Content for SPSite/SPWeb list
	
	$web = Get-SPWeb $siteURL
	if($web -ne $null)
	{
		if($web.HasUniqueRoleAssignments -eq $true)
		{
			$globalPermList += getPermissionInfo -userName $userName -Object $web
		}
		
		$listsToParse =  $web.Lists | where { $_.Hidden -eq $false } 
		foreach($list in $listsToParse)
		{
			Write-Host ">> SPList '$($list.Title)'" -Foregroundcolor Cyan
			
			#Step 3: Permissions on SPList
			$globalPermList += getPermissionInfo -userName $userName -Object $list
			
			#Step 4: Permissions on SPFolder
			#Write-Host "Check SPFolder Permissions" -Foregroundcolor Yellow `r
			$uniqueFolders = $list.Folders | where { $_.HasUniqueRoleAssignments -eq $true }                   
			foreach($folder in $uniqueFolders)
			{
				$globalPermList += getPermissionInfo -userName $userName -Object $folder
			}
			
			#Step 4: Permissions on SPListItem
			#Write-Host "Check SPFolder SPListItem" -Foregroundcolor Yellow `r
			$uniqueItems = $list.Items | where { $_.HasUniqueRoleAssignments -eq $true }
			#Get Item level permissions
			foreach($item in $uniqueItems)
			{
				$globalPermList += getPermissionInfo -userName $userName -Object $item
			}
		}#end foreach list
	}
	else
	{
		Write-Warning "SPWeb '$siteURL' does not exist."
	}
	
	if($web -ne $null)
	{
		$web.Dispose()
	}
	return $globalPermList;
}



<#
.SYNOPSIS
	Get User Permissions on a single SPSite (and all its childs)
	
.DESCRIPTION
	Get User Permissions on a single SPWeb (and all its childs)
	
.PARAMETER siteCollURL
	(Mandatory) URL of the SharePoint Site Collection
	
.PARAMETER object
	(Mandatory) SharePoint Securable object

.EXAMPLE
	getUserPermissionsOnSPWeb -siteCollURL <siteCollURL> -object <object>
	
.OUTPUTS
	Array $globalPermList containing all User Permissions on the single SPWeb
#>
function getUserPermissionsOnSPSite()
{
	param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteCollURL,
		[Parameter(Mandatory=$false, Position=2)]
		[string]$userName
	)
	
	$globalPermList = @() # $permList = @() <=> CSV Content for SPSite/SPWeb list
	
	$site = Get-SPSite $siteCollURL
	if($site -ne $null)
	{
		$globalPermList += checkSiteCollectionAdministrators -siteCollURL $siteCollURL -userName $userName
		
		foreach($subWeb in $site.AllWebs)		
		{
			Write-Host "> SPWeb '$($subWeb.URL)'" -Foregroundcolor Yellow
			$globalPermList += getUserPermissionsOnSPWeb -siteURL $subWeb.URL -userName $userName
		}
		
	}
	else
	{
		Write-Warning "SPSite '$siteCollURL' does not exist."
	}
	
	if($site -ne $null)
	{
		$site.Dispose()
	}
	return $globalPermList;
}


<#
.SYNOPSIS
	Get User Permissions on Web-Application (and all its childs: SPSite, SPWeb)
	
.DESCRIPTION
	Get User Permissions on Web-Application (and all its childs: SPSite, SPWeb)
	
.PARAMETER webAppURL
	(Mandatory) URL of the SharePoint Web-Application
	
.PARAMETER object
	(Mandatory) SharePoint Securable object

.EXAMPLE
	getUserPermissionsOnSPWeb -webAppURL <webAppURL> -object <object>
	
.OUTPUTS
	Array $globalPermList containing all User Permissions on the single SPWeb
#>
function getUserPermissionsOnSPWebApplication()
{
	param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$webAppURL,
		[Parameter(Mandatory=$false, Position=2)]
		[string]$userName
	)
	
	$globalPermList = @() # $permList = @() <=> CSV Content for SPSite/SPWeb list
	
	$webApp = Get-SPWebApplication $webAppURL
	if($webApp -ne $null)
	{
		foreach($site in $webApp.Sites)		
		{
			Write-Host "SPSite '$($site.URL)'" -Foregroundcolor Magenta
			$globalPermList += getUserPermissionsOnSPSite -siteCollURL $site.URL -userName $userName
		}
		
	}
	else
	{
		Write-Warning "Web-Application '$webAppURL' does not exist."
	}
	return $globalPermList;
}

