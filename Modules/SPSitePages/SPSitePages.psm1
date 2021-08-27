##############################################################################################
#              
# NAME: SPSitePages.psm1 
# PURPOSE: 
#	Manage Site Pages
#	Relies on an XML Configuration file to identify and handle Site Pages.
#	See SPSitePages.xml for Schema
#
##############################################################################################


<#
.SYNOPSIS
	Parse the file sitePagesDescriptionXML XML object and initiate the Site Pages customization

.DESCRIPTION
	Will parse the file sitePagesDescriptionXML XML object making the difference between the pages to add (AddPage node),
	the pages to be updated (UpdatePage node) and the pages to be deleted (DeletePage node)

.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER sitePagesDescriptionXML
	XML object of Site Pages to manage.

.EXAMPLE
	browseAndParseSPSitePagesXML -siteURL <SiteURL> -sitePagesDescriptionXML <sitePagesDescriptionXML>

.OUTPUTS
	None

.LINK
	None

.NOTES
	Created by: JBO
	Created: 16.01.2017
	Last Updated by: JBO
	Last Updated: 16.01.2017
#>
function browseAndParseSPSitePagesXML()
{
    Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[XML]$sitePagesDescriptionXML
	)

	try
	{
		$functionName = $MyInvocation.MyCommand.Name
		Write-Debug "[$functionName] Entering function" 
		Write-Debug "[$functionName] Parameter / siteURL: $siteURL"

		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if ($curWeb -ne $null)
		{
			$sitePages =  $sitePagesDescriptionXML.SelectNodes("/SPSitePages")
			if($sitePagesDescriptionXML.HasChildNodes  -and $sitePages.HasChildNodes)
			{

				# Create Pages
				$createPages = $sitePages.AddPage
				if($createPages -ne $null -and $createPages.HasChildNodes)
				{
					browseAndCreateSitePagesXML -siteURL $siteURL -pagesToAddXML $createPages 
				}
				else
				{
					Write-Warning "[$functionName] 'AddPage' node is empty."
				}

				# Delete Pages
				$deletePages = $sitePages.DeletePage
				if($deletePages -ne $null -and $deletePages.HasChildNodes)
				{
					browseAndDeleteSitePagesXML -siteURL $siteURL -pagesToDeleteXML $deletePages
				}
				else
				{
					Write-Warning "[$functionName] 'DeletePage' node is empty."
				}
			}
			else
			{
				Write-Warning "[$functionName] 'SPSitePages' XML defintion file is empty."
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
		Write-Debug "[$functionName] Exiting function" 
    }
}


<#
.SYNOPSIS
	Parse the file pagesToAddXML XML object and initiate the Site Page creation
	
.DESCRIPTION
	Will parse the file pagesToAddXML XML object and loop accross all the Pages to be created

.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER pagesToAddXML
	XML Element of the pages to be created
	
.EXAMPLE
	browseAndCreateSitePagesXML -siteURL <SiteURL> -pagesToAddXML <pagesToAddXML>
	
.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 20.03.2017
	Last Updated by: JBO
	Last Updated: 20.03.2017
#>
function browseAndCreateSitePagesXML()
{
    Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[XML.XmlElement]$pagesToAddXML
	)

    try
    {
       
		$functionName = $MyInvocation.MyCommand.Name
		Write-Debug "[$functionName] Entering function" 
		Write-Debug "[$functionName] Parameter / siteURL: $siteURL"
		Write-Host $pagesToAddXML
		foreach($pageToCreate in $pagesToAddXML.Page)
        {
            #Get the properties of site page
            $pageName = $pageToCreate.Name.Trim()
            $pageTitle = $pageToCreate.Title.Trim()
            $pageType = $pageToCreate.PageType.Trim()
            $pageLayout = $pageToCreate.PageLayout.Trim()
            $pageLibrary = $pageToCreate.PageLibrary.Trim()

            if(![string]::IsNullOrEmpty($pageName) -and ![string]::IsNullOrEmpty($pageType) -and ![string]::IsNullOrEmpty($pageLibrary))
            {
               createPage -siteURL $siteURL -pageLibraryName $pageLibrary -pageName $pageName -pageTitle $pageTitle -pageType $pageType -pageLayout $pageLayout
            }
            else
            {
               Write-Warning  "[$functionName] Page minimal informations (Name, Type or Location) are empty."      
            }
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
		Write-Debug "[$functionName] Exiting function" 
    }
}


<#
.SYNOPSIS
	Parse the file pagesToDeleteXML XML object and initiate the Site Page deletion
	
.DESCRIPTION
	Will parse the file pagesToAddXML XML object and loop accross all the Pages to be deleted

.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER pagesToDeleteXML
	XML Element of the pages to be deleted
	
.EXAMPLE
	browseAndDeleteSitePagesXML -siteURL <SiteURL> -pagesToAddXML <pagesToAddXML>
	
.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 20.03.2017
	Last Updated by: JBO
	Last Updated: 20.03.2017
#>
function browseAndDeleteSitePagesXML()
{
    Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		 [string]$siteURL,
		 [Parameter(Mandatory=$true, Position=2)]
		 [XML.XMLElement]$pagesToDeleteXML
	)

    try
    {
		$functionName = $MyInvocation.MyCommand.Name
		Write-Debug "[$functionName] Entering function" 
		Write-Debug "[$functionName] Parameter / siteURL: $siteURL"
		
		foreach($pageToDelete in $pagesToDeleteXML.Page)
		{
			#Get the properties of site page
			$pageName = $pageToDelete.Name.Trim()
			$pageURL = $pageToDelete.URL.Trim()

			if(![string]::IsNullOrEmpty($pageName) -and ![string]::IsNullOrEmpty($pageURL))
			{
				deletePage -siteURL $siteURL -pageName $pageName -pageURL $pageURL
			}
			else
			{
				 Write-Warning  "[$functionName] Page minimal informations (Name or URL) are empty."             
			}
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
		Write-Debug "[$functionName] Exiting function" 
    }
}


<#
.SYNOPSIS
	 Set the page pageName as Home Page (or Welcome Page) of the site
	
.DESCRIPTION
	 Set the page pageName as Home Page (or Welcome Page) of the site

.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER pageName
	Name of the page to be set as home page
		
.PARAMETER pageURL
	URL of the page to be set as home page

.PARAMETER webPartsToRemoveXML
	XML Element of the webparts to be added to the page
	
.EXAMPLE
	setHomePage -siteURL <SiteURL> -pageName <pageName> -pageURL <pageURL>

.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 20.03.2017
	Last Updated by: JBO
	Last Updated: 20.03.2017
#>
function setHomePage()
{
    Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[String]$pageName,
		[Parameter(Mandatory=$true, Position=3)]
		[String]$pageUrl
	)

    try
    {      
		$functionName = $MyInvocation.MyCommand.Name
		Write-Debug "[$functionName] Entering function" 
		
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if($curWeb -ne $null)
		{
			#Preparing pageURL to be append with ServerRelativeURL of current SPWeb
			if($pageUrl.toLower().StartsWith('/'))
			{
				$pageUrl = $pageUrl.Substring(1,$pageUrl.Length-1)
			}
			$fullURLtoPage = $curWeb.RootFolder.ServerRelativeURL + $pageUrl
			Write-Debug "[$functionName] Page Full URL: '$fullURLtoPage'" 

			#Get File
			$file = $curWeb.GetFile($fullURLtoPage)						
			if ( ($file -ne $null) -and ($file.Exists) )
			{					
				#Page Url
				Write-Host "[$functionName] About to set Page '$pageName' as Home Page" -ForegroundColor Magenta  `r
				$curWebrootFolder = $curWeb.RootFolder
				$curWebrootFolder.WelcomePage = $pageUrl
				$curWebrootFolder.Update()
				$curWeb.Update()
				Write-Host "[$functionName] Page '$pageName' has been set as Home Page" -ForegroundColor Green  `r
			}
			else
			{
				Write-Host "[$functionName] Page '$pageName' not found on site '$siteURL'."
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
		Write-Debug "[$functionName] Exiting function" 
    }
}





###########################################
###########################################
########## Check Page Existence ########### 
###########################################
###########################################

<#
.SYNOPSIS
	 Check if the page pageName exists in library pageLibraryName of site siteurl
	
.DESCRIPTION
	 Set the page pageName as Home Page (or Welcome Page) of the site

.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER pageLibraryName
	Name of the page library containg the page
			
.PARAMETER pageName
	Name of the page
	
.EXAMPLE
	existPageByName -siteURL <SiteURL> -pageLibraryName <pageLibraryName> -pageName <pageName>

.OUTPUTS
	true, if page exists, false if not

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 20.03.2017
	Last Updated by: JBO
	Last Updated: 20.03.2017
#>
function existPageByName()
{
	Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[String]$pageLibraryName,
		[Parameter(Mandatory=$true, Position=3)]
		[String]$pageName
	)

    try
    {
		      
		$functionName = $MyInvocation.MyCommand.Name
		Write-Debug "[$functionName] Entering function" 
		$pageExists = $false 

		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if($curWeb -ne $null)
		{
			$pageLibrary = $curWeb.Lists.TryGetList($pageLibraryName)
			if($pageLibrary -ne $null)
			{
				if(!$pageName.toLower().EndsWith('.aspx'))
				{
					$pageName = $pageName.Trim() + ".aspx"
				}

				#Open file
				$pageUrl = $pageLibrary.RootFolder.ServerRelativeUrl + "/" + $pageName.Trim()
				$file = $curWeb.GetFile($pageUrl)			
				if ( ($file -ne $null) -and ($file.Exists) )
				{
					$pageExists = $true
				}
				else
				{
					#Write-Warning "[$functionName] Page '$pageName' does not exist in Page Library '$pageLibraryName'."
					$pageExists = $false 
				}
			}
			else
			{
				Write-Warning "[$functionName] Page Library '$pageLibraryName' not found on site '$siteURL'."
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
		$pageExists = $false 
    }
	finally
    {
        if($curWeb -ne $null)
        {
			$curWeb.Dispose()
        }
		Write-Debug "[$functionName] Exiting function" 
    }
    return $pageExists
}


<#
.SYNOPSIS
	 Check if the page pageName exists in library pageLibraryName of site siteurl, using it's direct URL
	
.DESCRIPTION
	 Check if the page pageName exists in library pageLibraryName of site siteurl, using it's direct URL

.PARAMETER siteUrl
	URL of the SharePoint Site
			
.PARAMETER pageName
	Name of the page
				
.PARAMETER pageURL
	URL of the page

.EXAMPLE
	existPageByURL -siteURL <SiteURL> -pageName <pageName> -pageURL <pageURL>

.OUTPUTS
	true, if page exists, false if not

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 20.03.2017
	Last Updated by: JBO
	Last Updated: 20.03.2017
#>
function existPageByURL()
{
    Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[String]$pageName,
		[Parameter(Mandatory=$true, Position=3)]
		[String]$pageUrl
	)

    try
    {
		$pageExists = $false 
		$functionName = $MyInvocation.MyCommand.Name
		Write-Debug "[$functionName] Entering function" 

       	$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if($curWeb -ne $null)
		{
			#Preparing pageURL to be append with ServerRelativeURL of current SPWeb
			if($pageUrl.toLower().StartsWith('/'))
			{
				$pageUrl = $pageUrl.Substring(1,$pageUrl.Length-1)
			}
			$fullURLtoPage = $curWeb.RootFolder.ServerRelativeURL + $pageUrl
			Write-Debug "[$functionName] Page Full URL: '$fullURLtoPage'" 

			#Get File
			$file = $curWeb.GetFile($fullURLtoPage)					
			if ( ($file -ne $null) -and ($file.Exists) )
			{
				$pageExists = $true
			}
			else
			{
				#Write-Warning "[$functionName] Page '$pageName' not found on site '$siteURL'."
				$pageExists = $false 
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
		$pageExists = $false 
    }
	finally
    {
        if($curWeb -ne $null)
        {
			$curWeb.Dispose()
        }
		Write-Debug "[$functionName] Exiting function" 
    }
    return $pageExists
}




###########################################
###########################################
########## Create & Delete Page ########### 
###########################################
###########################################


<#
.SYNOPSIS
	Initiate creation of page PageName on library PageLibraryName depending on pageType
	
.DESCRIPTION
	Will initiate creation of page PageName on library PageLibraryName depending on pageType

.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER pageLibraryName
	Name of the page library
			
.PARAMETER pageName
	Name of the page
			
.PARAMETER pageTitle
	Title (DisplayName) of the page
			
.PARAMETER pageType
	Type of the page (Wiki, Webpart or Publishing)
					
.PARAMETER pageLayout
	Layout of the page

.EXAMPLE
	createPage -siteURL <SiteURL> -pageLibraryName <pageLibraryName> -pageName <pageName> -pageTitle <pageTitle> -pageType <pageType> -pageLayout <pageLayout>
	
.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 20.03.2017
	Last Updated by: JBO
	Last Updated: 20.03.2017
#>
function createPage()
{
	Param
	(
			[Parameter(Mandatory=$true, Position=1)]
			[string]$siteURL,
			[Parameter(Mandatory=$true, Position=2)]
			[String]$pageLibraryName,
			[Parameter(Mandatory=$true, Position=3)]
			[String]$pageName,
			[Parameter(Mandatory=$true, Position=4)]
			[String]$pageTitle,
			[Parameter(Mandatory=$true, Position=5)]
			[String]$pageType,
			[Parameter(Mandatory=$true, Position=6)]
			[String]$pageLayout
	)

    try
    {
		$functionName = $MyInvocation.MyCommand.Name
		Write-Debug "[$functionName] Entering function" 
		Write-Debug "[$functionName] Page Name: '$pageName'" 
		Write-Debug "[$functionName] Page Type: '$pageType'" 
		Write-Debug "[$functionName] Page Layout: '$pageLayout'" 

		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if($curWeb -ne $null)
		{			
			$pageLibrary = $curWeb.Lists.TryGetList($pageLibraryName)
			if($pageLibrary -ne $null)
			{
				$pageExists = existPageByName -siteURL $siteURL -pageLibraryName $pageLibraryName -pageName $pageName
				if($pageExists -eq $false)
				{
					$pageTypeLower = $pageType.Trim().ToLower();
					switch ($pageTypeLower)
					{
						"wiki" {
								createWikiPage -siteUrl $siteURL -pageLibraryName $pageLibraryName -pageName $pageName
								break;
						 }
						"publishing" {
								 createPublishingPage -siteUrl $siteURL -pageLibraryName $pageLibraryName -pageName $pageName -pageLayout $pageLayout
								 break;
						 }
						"webpart" {
								createWebPartPage -siteUrl $siteURL -pageLibraryName $pageLibraryName -pageName $pageName -pageTitle $pageTitle -pageLayout $pageLayout
								break;
						 }
						default {
								createWikiPage -siteUrl $siteURL -pageLibraryName $pageLibraryName -pageName $pageName
								break;
						 }
					}		
				}
				else
				{
					Write-Warning "[$functionName] Page '$pageName' already present on site '$siteURL'."
				}
			}
			else
			{
				Write-Warning "[$functionName] Page Library '$pageLibraryName' not found on site '$siteURL'."
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
		Write-Debug "[$functionName] Exiting function" 
    }
}


<#
.SYNOPSIS
	Initiate creation of Wiki page PageName in library PageLibraryName
	
.DESCRIPTION
	Initiate creation of Wiki page PageName in library PageLibraryName

.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER pageLibraryName
	Name of the page library
			
.PARAMETER pageName
	Name of the page

.EXAMPLE
	createWikiPage -siteURL <SiteURL> -pageLibraryName <pageLibraryName> -pageName <pageName> 
	
.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 20.03.2017
	Last Updated by: JBO
	Last Updated: 20.03.2017
#>
function createWikiPage()
{
	Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[String]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[String]$pageLibraryName,
		[Parameter(Mandatory=$true, Position=3)]
		[String]$pageName
	)

    try
    {
		$functionName = $MyInvocation.MyCommand.Name
		Write-Debug "[$functionName] Entering function" 
		Write-Debug "[$functionName] Page Library Name: '$pageLibraryName'" 
		Write-Debug "[$functionName] Page Name: '$pageName'" 

       $curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
       if($curWeb -ne $null)
       {
            $pageLibrary = $curWeb.Lists.TryGetList($pageLibraryName)
	        if($pageLibrary -ne $null)
	        {
                $rootFolder = $pageLibrary.RootFolder
                $files = $rootFolder.Files;
                $pageName = $pageName.Trim()

                if(!$pageName.toLower().EndsWith('.aspx'))
                {
                   $pageName = $pageName + ".aspx"
                }
				$pageExists = existPageByName -siteURL $siteURL -pageLibraryName $pageLibraryName -pageName $pageName 
				if($pageExists -eq $false)            
				{  
					Write-Host "[$functionName] About to create the Wiki Page '$pageName'." -ForegroundColor Magenta `r
					$newWikiPage = $files.Add($rootFolder.ServerRelativeUrl + "/$pageName", [Microsoft.SharePoint.SPTemplateFileType]::WikiPage)
					Write-Host "[$functionName] Wiki Page '$($newWikiPage.Name)' successfully created." -ForegroundColor Green `r
				}
				else
				{
					Write-Warning "[$functionName] Page '$pageName' already exists in Library '$pageLibraryName'."
				}
            }
			else
			{
				Write-Warning "[$functionName] Page Library '$pageLibraryName' not found on site '$siteURL'."
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
		Write-Debug "[$functionName] Exiting function" 
    }
}


<#
.SYNOPSIS
	Initiate creation of Wiki page PageName in library PageLibraryName
	
.DESCRIPTION
	Initiate creation of Wiki page PageName in library PageLibraryName

.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER pageLibraryName
	Name of the page library
			
.PARAMETER pageName
	Name of the page
			
.PARAMETER pageTitle
	Title (DisplayName) of the page
						
.PARAMETER pageLayout
	Layout of Page from (1 to 8).

.EXAMPLE
	createWebPartPage -siteURL <SiteURL> -pageLibraryName <pageLibraryName> -pageName <pageName> -pageTitle <pageTitle> -pageLayout <pageLayout>
	
.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 20.03.2017
	Last Updated by: JBO
	Last Updated: 20.03.2017
#>
function createWebPartPage()
{
	Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[String]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[String]$pageLibraryName,
		[Parameter(Mandatory=$true, Position=3)]
		[String]$pageName,
		[Parameter(Mandatory=$true, Position=4)]
		[String]$pageTitle,
		[Parameter(Mandatory=$true, Position=5)]
		[String]$pageLayout
	)

    try
    {
      	$functionName = $MyInvocation.MyCommand.Name
		Write-Debug "[$functionName] Entering function" 
		Write-Debug "[$functionName] Page Library Name: '$pageLibraryName'" 
		Write-Debug "[$functionName] Page Name: '$pageName'" 

       $curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
       if($curWeb -ne $null)
       {
            $pageLibrary = $curWeb.Lists.TryGetList($pageLibraryName)
	        if($pageLibrary -ne $null)
	        {
				$pageExists = existPageByName -siteURL $siteURL -pageLibraryName $pageLibraryName -pageName $pageName 
				if($pageExists -eq $false)            
				{   
					Write-Debug "[$functionName] PageLib (ID): $($pageLibrary.ID)"
			
					if(!$pageName.toLower().EndsWith('.aspx'))
					{
						$pageName = $pageName.Replace(".aspx","")
					}     
					       
					$xml = '<?xml version="1.0" encoding="UTF-8"?>
					<Method ID="0,NewWebPage">
						<SetList Scope="Request">' + $pageLibrary.ID + '</SetList>
						<SetVar Name="Cmd">NewWebPage</SetVar>
						<SetVar Name="ID">New</SetVar>
						<SetVar Name="Type">WebPartPage</SetVar>
						<SetVar Name="WebPartPageTemplate">' + $pageLayout + '</SetVar>
						<SetVar Name="Overwrite">true</SetVar>
						<SetVar Name="Title">'+ $pageName +'</SetVar>
					</Method>'

					Write-Host "[$functionName] About to create the Web-Part Page '$pageName'." -ForegroundColor Magenta `r
					$curWeb.ProcessBatchData($xml) | Out-Null 
					$pageLibrary.Update();
					Write-Host "[$functionName] Web-Part Page '$pageName' successfully created." -ForegroundColor Green `r
				
				}			
				else
				{
					Write-Warning "[$functionName] Page '$pageName' already exists in Library '$pageLibraryName'."
				}

			}
			else
			{
				Write-Warning "[$functionName] Page Library '$pageLibraryName' not found on site '$siteURL'."
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
		Write-Debug "[$functionName] Exiting function" 
    }
}


<#
.SYNOPSIS
	Initiate creation of Publishing page PageName in library PageLibraryName
	
.DESCRIPTION
	Initiate creation of Publishing page PageName in library PageLibraryName

.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER pageLibraryName
	Name of the page library
			
.PARAMETER pageName
	Name of the page
 
.PARAMETER pageLayout
	Layout of Page from (1 to 8).

.EXAMPLE
	createPublishingPage -siteURL <SiteURL> -pageLibraryName <pageLibraryName> -pageName <pageName> -pageLayout <pageLayout>
	
.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 20.03.2017
	Last Updated by: JBO
	Last Updated: 20.03.2017
#>
function createPublishingPage()
{
    Param
	([Parameter(Mandatory=$true, Position=1)]
	 [String]$siteURL,
     [Parameter(Mandatory=$true, Position=2)]
	 [String]$pageLibraryName,
     [Parameter(Mandatory=$true, Position=3)]
	 [String]$pageName,
     [Parameter(Mandatory=$true, Position=4)]
	 [String]$pageLayout)

    try
    {      
		$functionName = $MyInvocation.MyCommand.Name
		Write-Debug "[$functionName] Entering function" 
		Write-Debug "[$functionName] Page Name: '$pageName'" 
		
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
        $pubWeb = [Microsoft.SharePoint.Publishing.PublishingWeb]::GetPublishingWeb($curWeb)
        if($pubWeb -ne $null)
        {
            if(!$pageName.toLower().EndsWith('.aspx'))
            {
                $pageName = $pageName + ".aspx"
            }


            # Finds the appropriate page layout
            $pl = $pubWeb.GetAvailablePageLayouts() | Where { $_.Title -eq $PageLayout.Trim()}
            if($pl -ne $null)
            {
                # Add the new publishing page
                Write-Host "[createPublishingPage]Creating new publishing page" $pageName  
                $newPage = $pubWeb.AddPublishingPage($PageName, $pl)   
                $newPage.Title = $pageName  
                $newPage.Update();    
                # Check in the Page with Comments  
                $newPage.CheckIn("System Checked-IN")    
                # Publish the Page With Comments  
                $newPage.ListItem.File.Publish("UNHCR System Published")
                Write-Host "[createPublishingPage]Publishing Page" $pageName " Created Successfully" -ForegroundColor Green `r
            }
            else
            {
                Write-Host "[createPublishingPage] Publishing Page layout $PageLayout is not found. Impossible to create page." -ForegroundColor Cyan `r
            }
        }
        else
        {  
          Write-Host "[createPublishingPage] $siteURL is not a valid publishing web" -ForegroundColor  Cyan `r
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

        if($pubWeb -ne $null)
        {
			$pubWeb.Dispose()
        }
		Write-Debug "[$functionName] Exiting function" 
    }
}



<#
.SYNOPSIS
	Initiate deletion of page PageName on library PageLibraryName 
	
.DESCRIPTION
	Will initiate deletion of page PageName on library PageLibraryName

.PARAMETER siteUrl
	URL of the SharePoint Site

.PARAMETER pageName
	Name of the page
			
.PARAMETER pageURL
	URL of the page

.EXAMPLE
	deletePage -siteURL <SiteURL> -pageName <pageName> -pageURL <pageURL> 
	
.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 20.03.2017
	Last Updated by: JBO
	Last Updated: 20.03.2017
#>
function deletePage()
{
    Param
	([Parameter(Mandatory=$true, Position=1)]
	 [String]$siteURL,
     [Parameter(Mandatory=$true, Position=2)]
	 [String]$pageName,
     [Parameter(Mandatory=$false,Position=3)]
	 [String]$pageUrl)

    try
    {      
		$functionName = $MyInvocation.MyCommand.Name
		Write-Debug "[$functionName] Entering function" 
		Write-Debug "[$functionName] Page Name: '$pageName'" 
		Write-Debug "[$functionName] Page URL: '$pageUrl'" 

		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if($curWeb -ne $null)
		{			
			#Preparing pageURL to be append with ServerRelativeURL of current SPWeb
			if($pageUrl.toLower().StartsWith('/'))
			{
				$pageUrl = $pageUrl.Substring(1,$pageUrl.Length-1)
			}
			$fullURLtoPage = $curWeb.RootFolder.ServerRelativeURL + $pageUrl
			Write-Debug "[$functionName] Page Full URL: '$fullURLtoPage'" 

			#Get File
			$file = $curWeb.GetFile($fullURLtoPage)			
			if ( ($file -ne $null) -and ($file.Exists) )
			{
				Write-Host "[$functionName] About to delete Page '$pageName'" -ForegroundColor Magenta  `r
				$file.Delete()
				Write-Host "[$functionName] Page '$pageName' deleted successfully." -ForegroundColor Green  `r
			}
			else
			{
				Write-Warning "[$functionName] Page '$pageName' not found."
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
		Write-Debug "[$functionName] Exiting function" 
    }
}
