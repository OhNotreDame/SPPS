##############################################################################################
#              
# NAME: SPListViews.psm1 
# PURPOSE: Manage List Views
#	Relies on an XML Configuration file to identify and handle the List Views.
#	See SPSiteListViews.xml for Schema
# SOURCE : https://github.com/OhNotreDame/SPPS
##############################################################################################


<#
.SYNOPSIS
	Parse the file listViewsXML XML object and initiate the List Views customization
	
.DESCRIPTION
	Will parse the file listViewsXML XML object making the difference between the Views to add (AddView node),
	the views to be updated (UpdateView node) and the views to be deleted (DeleteView node)

.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER listViewsXML
	XML object of Site Groups to manage.
	
.EXAMPLE
	browseAndParseListViewsXML -siteURL <SiteURL> -listsViewXML <lsitViewDescriptionXML>
	
.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 16.03.2017
	Last Updated by: JBO
	Last Updated: 16.03.2017
#>
function browseAndParseListViewsXML()
{
    Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[xml]$listViewsXML
	)

	$functionName = $MyInvocation.MyCommand.Name
	Write-Debug "[$functionName] Entering function" 

    try
    {
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if ($curWeb -ne $null)
		{
			Write-Debug "[$functionName] Getting SPListViews node"
			$SPListViewsNode = $listViewsXML.SelectNodes("/SPListViews")
			if($SPListViewsNode -ne $null -and $SPListViewsNode.HasChildNodes)
			{
				################################
				####### Views to create #######
				################################
				Write-Host "[$functionName] About to call browseAndCreateViewsXML()." -ForegroundColor Cyan `r
				$viewToAddXML =  $SPListViewsNode.AddView 
				if($viewToAddXML -ne $null -and $viewToAddXML.HasChildNodes)
				{									
					browseAndCreateListViewsXML -siteURL $siteURL -ViewsToAddXML $viewToAddXML
				}
				else
				{
					Write-Warning "[$functionName] 'AddView' node is empty."
				}


				################################
				####### Views to edit #########
				################################
				Write-Host "[$functionName] About to call browseAndUpdateListViewsXML()." -ForegroundColor Cyan `r
				$viewToEditXML =  $SPListViewsNode.UpdateView
				if($viewToEditXML -ne $null -and $viewToEditXML.HasChildNodes)
				{									
					browseAndUpdateListViewsXML -siteURL $siteURL -ViewsToUpdateXML $viewToEditXML
				}
				else
				{
					Write-Warning "[$functionName] 'UpdateView' node is empty."
				}
				
				################################
				####### Views to delete #######
				################################
				Write-Host "[$functionName] About to call browseAndDeleteListViewsXML()."  -ForegroundColor Cyan `r
				$viewToDeleteXML =  $SPListViewsNode.DeleteView
				if($viewToDeleteXML -ne $null -and $viewToDeleteXML.HasChildNodes)
				{									
					browseAndDeleteListViewsXML -siteURL $siteURL -ViewsToDeleteXML $viewToDeleteXML
				}
				else
				{
					Write-Warning "[$functionName] 'DeleteView' node is empty."
				}
			
			}
			else
			{
				Write-Warning "[$functionName] 'SPListViews' XML defintion file is empty."
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
	Parse the node ViewsToAddXML XML object and initiate the View creation
	
.DESCRIPTION
	Will parse the node ViewsToAddXML XML object and loop accross all the views to be created

.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER ViewsToAddXML
	XML object of the views to add
	
.EXAMPLE
	browseAndCreateListViewsXML -siteURL <SiteURL> -ViewsToAddXML <viewToAddXML>
	
.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 16.03.2017
	Last Updated by: JBO
	Last Updated: 16.03.2017
#>
function browseAndCreateListViewsXML()
{
    Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[XML.XmlElement]$ViewsToAddXML
	)

	$functionName = $MyInvocation.MyCommand.Name
	Write-Debug "[$functionName] Entering function" 

    try
    {
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if ($curWeb -ne $null)
		{
			Write-Debug "[$functionName] Getting ViewsToAddXML node"
			if($ViewsToAddXML -ne $null -and $ViewsToAddXML.HasChildNodes)
			{
				Write-Debug "[$functionName] ViewsToAddXML node has child"
				foreach($parentListXML in $ViewsToAddXML.List)
				{
					$listName = $parentListXML.Title;
					Write-Debug "[$functionName] Current ListName '$listName'"
					$spList = $curWeb.Lists.TryGetList($listName);
					if($spList -ne $null)
					{
						foreach($view in $parentListXML.View)
						{
							$doesViewExist = existListView -siteURL $siteURL -listName $listName -viewName $view.Name
							if($doesViewExist -ne $true)
							{
								Write-Debug "[$functionName] About to call createListView() for '$($view.Name)' on List '$listName'."
								createListView -siteURL $siteURL -listName $listName -listViewDefinitionXML $view
							}
							else
							{
								Write-Warning "[$functionName] View '$($view.Name)' already created on List '$listName'."
							}
						}#foreach $view
					}
					else
					{
						Write-Warning "[$functionName] List '$listName' not found on site '$siteURL'."
					}						
				}#foreach $parentListXML
			}
			else
			{
				Write-Warning "[$functionName] 'AddView' XML node is empty."
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
		Write-Debug "[$functionName] Exiting function" 
    }
}


<#
.SYNOPSIS
	Parse the node ViewsToUpdateXML XML object and initiate the View update
	
.DESCRIPTION
	Will parse the node ViewsToUpdateXML XML object and loop accross all the views to be updated

.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER ViewsToUpdateXML
	XML object of the views to update
	
.EXAMPLE
	browseAndUpdateListViewsXML -siteURL <SiteURL> -ViewsToAddXML <viewToAddXML>
	
.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 16.03.2017
	Last Updated by: JBO
	Last Updated: 16.03.2017
#>
function browseAndUpdateListViewsXML()
{
    Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[XML.XmlElement]$ViewsToUpdateXML
	)

	$functionName = $MyInvocation.MyCommand.Name
	Write-Debug "[$functionName] Entering function" 

    try
    {
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if ($curWeb -ne $null)
		{
			if($ViewsToUpdateXML -ne $null -and $ViewsToUpdateXML.HasChildNodes)
			{
				foreach($parentListXML in $ViewsToUpdateXML.List)
				{
					$listName = $parentListXML.Title;
					$spList = $curWeb.Lists.TryGetList($listName);
					if($spList -ne $null)
					{
						foreach($view in $parentListXML.View)
						{
							$doesViewExist = existListView -siteURL $siteURL -listName $listName -viewName $view.Name
							if($doesViewExist -eq $true)
							{
								Write-Debug "[$functionName] About to call createListView() for '$($view.Name)' on List '$listName'."
								updateListView -siteURL $siteURL -listName $listName -viewName $view.Name -listViewDefinitionXML $view
							}
							else
							{
								Write-Warning "[$functionName] View '$($view.Name)' does not exist on List '$listName'."
							}
						}#foreach $view
					}
					else
					{
						Write-Warning "[$functionName] List '$listName' not found on site '$siteURL'."
					}						
				}#foreach $parentListXML
			}
			else
			{
				Write-Warning "[$functionName] 'AddView' XML node is empty."
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
		Write-Debug "[$functionName] Exiting function" 
    }
}


<#
.SYNOPSIS
	Parse the node viewToDeleteXML XML object and initiate the View deletion
	
.DESCRIPTION
	Will parse the node viewToDeleteXML XML object and loop accross all the views to be deleted

.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER ViewsToUpdateXML
	XML object of the views to delete
	
.EXAMPLE
	browseAndDeleteListViewsXML -siteURL <SiteURL> -ViewsToDeleteXML <viewToDeleteXML>
	
.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 16.03.2017
	Last Updated by: JBO
	Last Updated: 16.03.2017
#>
function browseAndDeleteListViewsXML()
{
    Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[XML.XmlElement]$ViewsToDeleteXML
	)

	$functionName = $MyInvocation.MyCommand.Name
	Write-Debug "[$functionName] Entering function" 

    try
    {
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if ($curWeb -ne $null)
		{
			if($ViewsToDeleteXML -ne $null -and $ViewsToDeleteXML.HasChildNodes)
			{
				foreach($parentListXML in $ViewsToDeleteXML.List)
				{
					$listName = $parentListXML.Title;
					$spList = $curWeb.Lists.TryGetList($listName);
					if($spList -ne $null)
					{
						foreach($view in $parentListXML.View)
						{
							$doesViewExist = existListView -siteURL $siteURL -listName $listName -viewName $view.Name
							if($doesViewExist -eq $true)
							{
								Write-Debug "[$functionName] About to call deleteListView() for '$($view.Name)' on List '$listName'."
								deleteListView -siteURL $siteURL -listName $listName -viewName $view.Name
							}
							else
							{
								Write-Warning "[$functionName] View '$($view.Name)' does not exist on List '$listName'."
							}
						}#foreach $view
					}
					else
					{
						Write-Warning "[$functionName] List '$listName' not found on site '$siteURL'."
					}						
				}#foreach $parentListXML
			}
			else
			{
				Write-Warning "[$functionName] 'AddView' XML node is empty."
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
		Write-Debug "[$functionName] Exiting function" 
    }
}






<#
.SYNOPSIS
	Check if a List View $viewName exists in List $listName
	
.DESCRIPTION
	Check if a List View $viewName exists in List $listName

.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER listName
	Name of the List
		
.PARAMETER viewName
	Name of the View
	
.EXAMPLE
	existListView -siteURL <SiteURL> -listName <listName> -viewName <viewName>
	
.OUTPUTS
	Return true if exists, false instead

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 16.03.2017
	Last Updated by: JBO
	Last Updated: 16.03.2017
#>
function existListView()
{
    Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[String]$listName,
		[Parameter(Mandatory=$true, Position=3)]
		[String]$viewName
	)

	$functionName = $MyInvocation.MyCommand.Name
	Write-Debug "[$functionName] Entering function" 
	
    try
    {
        $exist = $false
        $curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
        if($curWeb -ne $null)
        {
             $spList = $curWeb.Lists.TryGetList($listName)
             if($spList -ne $null)
             {
				$view = $spList.Views[$ViewName]
				if($view -ne $null)
				{
					$exist = $true
				}
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
		Write-Debug "[$functionName] Exiting function" 
    }
    return $exist
}



<#
.SYNOPSIS
	Create the List View $viewName in List $listName based on its definition (Query, Fields, Settings) described in $listViewDefinitionXML 
	
.DESCRIPTION
	Create the List View $viewName in List $listName based on its definition (Query, Fields, Settings) described in $listViewDefinitionXML 

.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER listName
	Name of the List
		
.PARAMETER listViewDefinitionXML
	Definition of the View
	
.EXAMPLE
	createListView -siteURL <SiteURL> -listName <listName> -$listViewDefinitionXML <ViewDefinitionXML>
	
.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 16.03.2017
	Last Updated by: JBO
	Last Updated: 16.03.2017
#>
function createListView()
{
    Param
	([Parameter(Mandatory=$true, Position=1)]
	 [string]$siteURL,
     [Parameter(Mandatory=$true, Position=2)]
	 [String]$listName,
     [Parameter(Mandatory=$true, Position=3)]
	 [XML.XMLElement]$listViewDefinitionXML)

    $functionName = $MyInvocation.MyCommand.Name
	Write-Debug "[$functionName] Entering function" 
	
	try
    {
		$listViewDefinitionXML
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if($curWeb -ne $null)
		{
            $spList = $curWeb.Lists.TryGetList($listName)
			if ( $spList -ne $null )
			{
				$existView = existListView -siteURL $siteURL -listName $listName -viewName $listViewDefinitionXML.Name
				if($existView -ne $true)
				{					
					$viewTitle = $listViewDefinitionXML.Name 
					$viewURL= $listViewDefinitionXML.URL
					
					#------------------------------------------------------------------#
					#---------------------- VIEW - EXPTECTED FIELDS -------------------#
					#------------------------------------------------------------------#	
					#Write-Host "$functionName parsing columns to add from XML. Building intermediate viewField string collection" -ForegroundColor Magenta `r
					$viewFields = New-Object System.Collections.Specialized.StringCollection
					foreach($viewField in $listViewDefinitionXML.ViewFields.FieldRef)
					{
						$viewFields.Add($viewField.Name) > $null
					}
					#Write-Host "$functionName Intermediate viewField string collection successfully buit." -ForegroundColor Magenta `r

					#------------------------------------------------------------------#
					#--------------------- VIEW - EXPTECTED SETTINGS ------------------#
					#------------------------------------------------------------------#	
					#Query property
					$viewQuery =$listViewDefinitionXML.Query.InnerXml
					
					#RowLimit property
					if (![string]::IsNullOrEmpty($listViewDefinitionXML.RowLimit.InnerXml))
					{
						$viewRowLimit = $listViewDefinitionXML.RowLimit.InnerXml
					}
					else
					{
						$viewRowLimit = $listViewDefinitionXML.RowLimit
					}
					Write-Debug "[$functionName] viewRowLimit: $viewRowLimit" 
					
					#Paged property
					if (![string]::IsNullOrEmpty($listViewDefinitionXML.RowLimit.Paged))
					{
						$viewPaged =[System.Convert]::ToBoolean($listViewDefinitionXML.RowLimit.Paged)
					}
					else
					{
						$viewPaged =[System.Convert]::ToBoolean($listViewDefinitionXML.Attributes["Paged"].value)
					}
					Write-Host "[$functionName] viewPaged: $viewPaged" 
					  
					#DefaultView property
					if (![string]::IsNullOrEmpty($listViewDefinitionXML.Attributes["DefaultView"].value))
					{
						$viewDefaultView  = $false
					}
					else
					{
						$viewDefaultView =[System.Convert]::ToBoolean($listViewDefinitionXML.Attributes["DefaultView"].value)
					}
				
					#Create the view in the destination list					
					$curWeb.AllowUnsafeUpdates=$true

					Write-Host "[$functionName] About to create View '$viewTitle' in list '$listName'." -ForegroundColor Magenta `r
					$newview = $spList.Views.Add($viewURL, $viewFields, $viewQuery,$viewRowLimit, $viewPaged, $viewDefaultView)
					$newview.Update()
					$newview.Title= $viewTitle
					$newview.Update()
					$spList.Update()
					Write-Host "[$functionName] View '$viewTitle' successfully created in list '$listName'." -ForegroundColor Green `r

					$curWeb.AllowUnsafeUpdates=$false
					$curWeb.Update()
				}
				else
				{
					Write-Warning "[$functionName] View '$viewTitle' already exists in list '$listName'." 
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
		Write-Debug "[$functionName] Exiting function" 
    }
}


<#
.SYNOPSIS
	Delete the view $viewName from list $listName
	
.DESCRIPTION
	Delete the view $viewName from list $listName

.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER listName
	Name of the List
		
.PARAMETER viewName
	Name of the View
	
.EXAMPLE
	deleteListView -siteURL <SiteURL> -listName <listName> -viewName <viewName>
	
.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 16.03.2017
	Last Updated by: JBO
	Last Updated: 16.03.2017
#>
function deleteListView()
{
    Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[String]$listName,
		[Parameter(Mandatory=$true, Position=3)]
		[String]$viewName
	)
   
	$functionName = $MyInvocation.MyCommand.Name
	Write-Debug "[$functionName] Entering function" 
	
    try
    {
        $curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
        if($curWeb -ne $null)
        {
             $spList = $curWeb.Lists.TryGetList($listName)
             if($spList -ne $null)
             {
                $spView = $spList.Views[$ViewName]
                if($spView -ne $null)
                {
                    Write-Host "$functionName About to delete View '$viewName' from List '$listName'." -ForegroundColor Magenta `r
					$spList.Views.Delete($spView.ID)
                    $spList.Update()
                    $curWeb.Update()
                    Write-Host "$functionName View '$viewName' removed from List '$listName'." -foregroundcolor Green `r
                }
                else
                {
                    Write-Warning "$functionName View '$viewName' does not exist on List '$listName'."
                }
             }
             else
             {
                Write-Warning "$functionName List '$listName' does not exist on site '$siteURL'."
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
		Write-Debug "[$functionName] Exiting function" 
    }
}

<#
.SYNOPSIS
	Update the List View $viewName in List $listName
	
.DESCRIPTION
	Update the List View $viewName in List $listName by replacing its fields and its current query and settings by the one from $listViewDefinitionXML ML 

.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER listName
	Name of the List
		
.PARAMETER listViewDefinitionXML
	Definition of the View
	
.EXAMPLE
	updateListView -siteURL <SiteURL> -listName <listName> -$listViewDefinitionXML <ViewDefinitionXML>
	
.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 16.03.2017
	Last Updated by: JBO
	Last Updated: 16.03.2017
#>
function updateListView()
{
    Param
	([Parameter(Mandatory=$true, Position=1)]
	 [string]$siteURL,
     [Parameter(Mandatory=$true, Position=2)]
	 [String]$listName,
     [Parameter(Mandatory=$true, Position=3)]
	 [String]$viewName,
     [Parameter(Mandatory=$true, Position=4)]
	 [XML.XMLElement]$listViewDefinitionXML
	)

	$functionName = $MyInvocation.MyCommand.Name
	Write-Debug "[$functionName] Entering function" 
	
    try
    {		
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if($curWeb -ne $null)
		{
            $spList = $curWeb.Lists.TryGetList($listName)
			if ( $spList -ne $null )
			{
				$existView = existListView -siteURL $siteURL -listName $listName -viewName $viewName
				if($existView -eq $true)
				{	
					$view = $spList.Views[$viewName]
					if($view -ne $null)
					{
						Write-Host "$functionName About to update view '$viewName' in list '$listName'." -ForegroundColor Magenta `r
						
						#------------------------------------------------------------------#
						#----------------- GETTING VIEW SETTINGS FROM XML -----------------#
						#------------------------------------------------------------------#						
						#Query property
						$viewQuery =$listViewDefinitionXML.Query.InnerXml
						
						#RowLimit property
						if (![string]::IsNullOrEmpty($listViewDefinitionXML.RowLimit.InnerXml))
						{
							$viewRowLimit = $listViewDefinitionXML.RowLimit.InnerXml
						}
						else
						{
							$viewRowLimit = $listViewDefinitionXML.RowLimit
						}
						Write-Debug "[$functionName] viewRowLimit: $viewRowLimit" 
						
						#Paged property
						if (![string]::IsNullOrEmpty($listViewDefinitionXML.RowLimit.Paged))
						{
							$viewPaged =[System.Convert]::ToBoolean($listViewDefinitionXML.RowLimit.Paged)
						}
						else
						{
							$viewPaged =[System.Convert]::ToBoolean($listViewDefinitionXML.Attributes["Paged"].value)
						}
						Write-Host "[$functionName] viewPaged: $viewPaged" 
						  
						#DefaultView property
						if (![string]::IsNullOrEmpty($listViewDefinitionXML.Attributes["DefaultView"].value))
						{
							$viewDefaultView  = $false
						}
						else
						{
							$viewDefaultView =[System.Convert]::ToBoolean($listViewDefinitionXML.Attributes["DefaultView"].value)
						}

						#------------------------------------------------------------------#
						#---------------------- CHANING VIEW SETTINGS ---------------------#
						#------------------------------------------------------------------#	
						#Write-Host "$functionName About to apply new settings on view '$viewName'." -ForegroundColor Magenta `r
						$view.Query = $viewQuery
						$view.RowLimit = $viewRowLimit
						$view.Paged = $viewPaged
						$view.DefaultView = $viewDefaultView
						$view.Update()
						#Write-Host "$functionName New settings applied on view '$viewName'." -ForegroundColor Green `r

						#------------------------------------------------------------------#
						#------------------------ CLEARING ALL FIELDS ---------------------#
						#------------------------------------------------------------------#	
						#Write-Host "$functionName About to clear all columns from view '$viewName'." -ForegroundColor Magenta `r
						$view.ViewFields.DeleteAll();
						$view.Update();
						#Write-Host "$functionName All columns successfully removed from view '$viewName'." -ForegroundColor Green `r

						#------------------------------------------------------------------#
						#----------------------- ADDING REQUIRED FIELDS -------------------#
						#------------------------------------------------------------------#	
						#Write-Host "$functionName About to add columns to view '$viewName'." -ForegroundColor Magenta `r
						$viewFields = New-Object System.Collections.Specialized.StringCollection
						foreach($viewField in $listViewDefinitionXML.ViewFields.FieldRef)
						{
							$curFieldName = $viewField.Name
							#Write-Host "$functionName Adding column '$curFieldName' from view '$viewName'." -ForegroundColor Magenta `r
							$view.ViewFields.Add($curFieldName);
							#Write-Host "$functionName Column '$curFieldName' added to view '$viewName'." -ForegroundColor Magenta `r
						}
						$view.Update();
						#Write-Host "$functionName Columns successfully added to view '$viewName'." -ForegroundColor Green `r

						Write-Host "$functionName View '$viewName' successfully updated in list '$listName'." -ForegroundColor Green `r

					}
					else
					{
						Write-Warning "$functionName SPView object for '$viewName' is null."
					}
				}
				else
				{
					Write-Warning "$functionName View '$viewName' does not exist on List '$listName'."
				}
			}
			else
			{
				Write-Warning "$functionName List '$listName' does not exist on site '$siteURL'."
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
		Write-Debug "[$functionName] Exiting function" 
    }
}



<#
.SYNOPSIS
	Update an existing List View $viewName of List $listName replacing its viewScope by the one passed via $viewScope paramater
	
.DESCRIPTION
	Update an existing List View $viewName of List $listName replacing its viewScope by the one passed via $viewScope paramater

.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER listName
	Name of the List

.PARAMETER viewName
	Name of the View
			
.PARAMETER viewScope
	Name of the expected viewScope
	
.EXAMPLE
	changeListViewScope -siteURL <SiteURL> -listName <listName> -viewName <viewName> -viewScope <viewScope>
	
.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 16.03.2017
	Last Updated by: JBO
	Last Updated: 16.03.2017
#>
function changeListViewScope()
{
    Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[String]$listName,
		[Parameter(Mandatory=$true, Position=3)]
		[String]$viewName,
		[Parameter(Mandatory=$true, Position=4)]
		[String]$viewScope
	)

	$functionName = $MyInvocation.MyCommand.Name
	Write-Debug "[$functionName] Entering function" 
	
    try
    {
		
		switch ($viewScope) 
		{ 
			"Default" { $SPViewScope= [Microsoft.SharePoint.SPViewScope]::Default } 
			"Recursive" { $SPViewScope= [Microsoft.SharePoint.SPViewScope]::Recursive } 
			"RecursiveAll" { $SPViewScope= [Microsoft.SharePoint.SPViewScope]::RecursiveAll } 
			"FilesOnly" { $SPViewScope= [Microsoft.SharePoint.SPViewScope]::FilesOnly } 
			default { $SPViewScope= [Microsoft.SharePoint.SPViewScope]::Default }
		}

		Write-Debug "$functionName viewScope/ Paramater: '$viewScope'."
		Write-Debug "$functionName viewScope/ Microsoft.SharePoint.SPViewScope(equivalent): '$SPViewScope'."

      
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if($curWeb -ne $null)
		{
            $spList = $curWeb.Lists.TryGetList($listName)
			if ( $spList -ne $null )
			{
				$existView = existListView -siteURL $siteURL -listName $listName -viewName $viewName
				if($existView -eq $true)
				{	
					$view = $spList.Views[$viewName]
					if($view -ne $null)
					{
						Write-Host "$functionName About to set viewScope '$SPViewScope' to view '$viewName' in list '$listName'." -ForegroundColor Magenta `r
						$view.Scope = $SPViewScope;
                        $view.Update()
                        $spList.Update()						
						Write-Host "$functionName ViewScope '$SPViewScope' in view '$viewName' successfully updated in list '$listName'." -ForegroundColor Green `r

					}
					else
					{
						Write-Warning "$functionName SPView object for '$viewName' is null."
					}
				}
				else
				{
					Write-Warning "$functionName View '$viewName' does not exist on List '$listName'."
				}
			}
			else
			{
				Write-Warning "$functionName List '$listName' does not exist on site '$siteURL'."
			}
		}
		else
		{
			Write-Warning "$functionName Site '$siteURL' does not exist." 
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