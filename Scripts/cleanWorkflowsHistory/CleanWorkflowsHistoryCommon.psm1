<#

.SYNOPSIS
	Get Nintex Workflow History List-Items Count
	
.DESCRIPTION
	Get Nintex Workflow History List-Items Count
	
.PARAMETER siteUrl
	(Mandatory) URL of the SharePoint Site
	
.EXAMPLE
	getNintexWorklowsHistoryItemsCount -siteUrl <siteUrl>
	
.OUTPUTS
	Integer : Item/Row Count (>=0) or -1 if an error occured
	
.LINK
	None
	
.NOTES
	Source: https://github.com/OhNotreDame/SPPS
	Created by: JBO
	Created: 29.03.2017
	Last Updated by: JBO
	Last Updated: 30.03.2017
	
#>
function getNintexWorklowsHistoryItemsCount
{
	param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[ValidateNotNullOrEmpty()]
		[string]$siteURL
	)
	
	$functionName = $MyInvocation.MyCommand.Name
	Write-Debug "[$functionName] Entering function" 
	Write-Debug "[$functionName] Parameter / siteURL: $siteURL"
	
	try
	{
		$itemCount = -1;
		
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if($curWeb -ne $null) 
		{
			Write-Debug "[$functionName] About to get Workflow History 'List Item' Count"
			$wfListName ="NintexWorkflowHistory"
			$spList = $curWeb.Lists.TryGetList($wfListName)
			if($spList -ne $null)
			{
				$itemCount = $spList.ItemCount
				Write-Debug "[$functionName] Workflow History List (SP) Count: $itemCount"
			}		
			else
			{
				Write-Warning "[$functionName] List '$wfListName' does not exist on site '$siteName'."					
			}
		}
		else
		{
			Write-Warning "[$functionName] Site '$siteURL' does not exist."
		}
	}
	catch [Exception]
	{	
		Write-Host ""
		Write-Host "/!\ [$scriptName] An exception has been caught /!\ " -Foregroundcolor Red `r
		Write-Host "Type: " $_.Exception.GetType().FullName -Foregroundcolor Red `r
		Write-Host "Message: " $_.Exception.Message -Foregroundcolor Red `r
		Write-Host "Stacktrace: `n" $_.Exception.Stacktrace -Foregroundcolor Red `r
	}
	finally
	{
		if($curWeb -ne $null) 
		{
			$curWeb.Dispose();
		}
		Write-Debug "[$functionName] Exiting function" 
	}
	return $itemCount;
}


<#

.SYNOPSIS
	Get Nintex Table 'Workflow Instance' Rows Count for a specific SPWeb
	
.DESCRIPTION
	Get Nintex Table 'Workflow Instance' Rows Count  for a specific SPWeb
	
.PARAMETER siteUrl
	(Mandatory) URL of the SharePoint Site (SPWeb)
	
.EXAMPLE
	getSPWebNintexWorklowsInstanceDBCount -siteUrl <siteUrl>
	getSPWebNintexWorklowsInstanceDBCount -siteUrl http://intranet-assura/FIN
	getSPWebNintexWorklowsInstanceDBCount -siteUrl http://intranet-assura
	
.OUTPUTS
	Integer : Item/Row Count (>=0) or -1 if an error occured
	
.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 29.03.2017
	Last Updated by: JBO
	Last Updated: 30.03.2017
	
#>
function getSPWebNintexWorklowsInstanceDBCount
{
	param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[ValidateNotNullOrEmpty()]
		[string]$siteUrl
	)
	
	$functionName = $MyInvocation.MyCommand.Name
	Write-Debug "[$functionName] Entering function" 
	Write-Debug "[$functionName] Parameter / siteURL: $siteURL"
	
	try
	{
		$rowsCount = -1;

		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if($curWeb -ne $null) 
		{
			$siteID = $curWeb.ID
			Write-Debug "" 
			Write-Debug "About to get Workflow Instances 'Table Rows' Count"
			$CFGDB = [Nintex.Workflow.Administration.ConfigurationDatabase]::OpenConfigDataBase().Database
			# Creating instance of .NET SQL client
			$cmd = New-Object -TypeName System.Data.SqlClient.SqlCommand
			$cmd.CommandType = [System.Data.CommandType]::Text

			# Begin SQL Query 
			$cmd.CommandText = "SELECT COUNT (*) as 'Workflow Instance Count'			
								FROM [$CFGDB].[dbo].WorkflowInstance
								WHERE WebID = '"+$siteID+"'";
			
			$reader = [Nintex.Workflow.Administration.ConfigurationDatabase]::GetConfigurationDatabase().ExecuteReader($cmd)
			while($reader.Read())
			{
				$rowsCount = $reader["Workflow Instance Count"]
				Write-Debug "Workflow Instance (DB - Nintex) Count: $rowsCount"
			}
		}
		else
		{
			Write-Warning "[$functionName] Site '$siteURL' does not exist."
		}
	}
	catch [Exception]
	{	
		Write-Host ""
		Write-Host "/!\ [$scriptName] An exception has been caught /!\ " -Foregroundcolor Red `r
		Write-Host "Type: " $_.Exception.GetType().FullName -Foregroundcolor Red `r
		Write-Host "Message: " $_.Exception.Message -Foregroundcolor Red `r
		Write-Host "Stacktrace: `n" $_.Exception.Stacktrace -Foregroundcolor Red `r
	}
	finally
	{
		if($curWeb -ne $null) 
		{
			$curWeb.Dispose();
		}
		Write-Debug "[$functionName] Exiting function" 
	}
	return $rowsCount;
}


<#

.SYNOPSIS
	Get Nintex Table 'Workflow Instance' Rows Count for a specific SPSite
	
.DESCRIPTION
	Get Nintex Table 'Workflow Instance' Rows Count for a specific SPSite
	
.PARAMETER siteUrl
	(Mandatory) URL of the SharePoint Site (SPSite)
	
.EXAMPLE
	getSPSiteNintexWorklowsInstanceDBCount -siteUrl <siteUrl>
	getSPSiteNintexWorklowsInstanceDBCount -siteUrl http://intranet-assura
	
.OUTPUTS
	Integer : Item/Row Count (>=0) or -1 if an error occured
	
.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 29.03.2017
	Last Updated by: JBO
	Last Updated: 30.03.2017
	
#>
function getSPSiteNintexWorklowsInstanceDBCount
{
	param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[ValidateNotNullOrEmpty()]
		[string]$siteUrl
	)
	
	$functionName = $MyInvocation.MyCommand.Name
	Write-Debug "[$functionName] Entering function" 
	Write-Debug "[$functionName] Parameter / siteURL: $siteURL"
	
	try
	{
		$rowsCount = -1;

		$curSite = Get-SPSite $siteURL -ErrorAction SilentlyContinue
		if($curSite -ne $null) 
		{
			$siteID = $curSite.ID
			Write-Debug "" 
			Write-Debug "About to get Workflow Instances 'Table Rows' Count"
			$CFGDB = [Nintex.Workflow.Administration.ConfigurationDatabase]::OpenConfigDataBase().Database
			# Creating instance of .NET SQL client
			$cmd = New-Object -TypeName System.Data.SqlClient.SqlCommand
			$cmd.CommandType = [System.Data.CommandType]::Text

			# Begin SQL Query 
			$cmd.CommandText = "SELECT COUNT (*) as 'Workflow Instance Count'			
								FROM [$CFGDB].[dbo].WorkflowInstance
								WHERE SiteID = '"+$siteID+"'";
			
			$reader = [Nintex.Workflow.Administration.ConfigurationDatabase]::GetConfigurationDatabase().ExecuteReader($cmd)
			while($reader.Read())
			{
				$rowsCount = $reader["Workflow Instance Count"]
				Write-Debug "Workflow Instance (DB - Nintex) Count: $rowsCount"
			}
		}
		else
		{
			Write-Warning "[$functionName] Site '$siteURL' does not exist."
		}
	}
	catch [Exception]
	{	
		Write-Host ""
		Write-Host "/!\ [$scriptName] An exception has been caught /!\ " -Foregroundcolor Red `r
		Write-Host "Type: " $_.Exception.GetType().FullName -Foregroundcolor Red `r
		Write-Host "Message: " $_.Exception.Message -Foregroundcolor Red `r
		Write-Host "Stacktrace: `n" $_.Exception.Stacktrace -Foregroundcolor Red `r
	}
	finally
	{
		if($curSite -ne $null) 
		{
			$curSite.Dispose();
		}
		Write-Debug "[$functionName] Exiting function" 
	}
	return $rowsCount;
}



<#

.SYNOPSIS
	Get Nintex Table 'Workflow Progress' Rows Count for a specific SPWeb
	
.DESCRIPTION
	Get Nintex Table 'Workflow Progress' Rows Count  for a specific SPWeb
	
.PARAMETER siteUrl
	(Mandatory) URL of the SharePoint Site (SPWeb)
	
.EXAMPLE
	getSPWebNintexWorklowsProgressDBCount -siteUrl <siteUrl>
	getSPWebNintexWorklowsProgressDBCount -siteUrl http://intranet-assura/FIN
	getSPWebNintexWorklowsProgressDBCount -siteUrl http://intranet-assura
	
.OUTPUTS
	Integer : Item/Row Count (>=0) or -1 if an error occured
	
.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 29.03.2017
	Last Updated by: JBO
	Last Updated: 30.03.2017
	
#>
function getSPWebNintexWorklowsProgressDBCount
{
	param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[ValidateNotNullOrEmpty()]
		[string]$siteURL
	)
	
	$functionName = $MyInvocation.MyCommand.Name
	Write-Debug "[$functionName] Entering function" 
	Write-Debug "[$functionName] Parameter / siteURL: $siteURL"
	
	try
	{
		$rowsCount = -1;
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if($curWeb -ne $null) 
		{
			$siteID = $curWeb.ID
			Write-Debug "" 
			Write-Debug "About to get Workflow Instances 'Table Rows' Count"
			$CFGDB = [Nintex.Workflow.Administration.ConfigurationDatabase]::OpenConfigDataBase().Database
			# Creating instance of .NET SQL client
			$cmd = New-Object -TypeName System.Data.SqlClient.SqlCommand
			$cmd.CommandType = [System.Data.CommandType]::Text

			# Begin SQL Query 
			$cmd.CommandText = "SELECT COUNT([WorkflowProgressID]) as 'Workflow Progress Count'	
								FROM [$CFGDB].[dbo].[WorkflowProgress] WP
								WHERE WP.[InstanceID] IN (SELECT [InstanceID]
															FROM  [$CFGDB].[dbo].[WorkflowInstance]
															WHERE WebID = '"+$siteID+"');";
			
			$reader = [Nintex.Workflow.Administration.ConfigurationDatabase]::GetConfigurationDatabase().ExecuteReader($cmd)
			while($reader.Read())
			{
				$rowsCount = $reader["Workflow Progress Count"]
				Write-Debug "Workflow Progress (DB - Nintex) Count: $Count_WfProgress" 
			}
		}
		else
		{
			Write-Warning "[$functionName] Site '$siteURL' does not exist."
		}
	}
	catch [Exception]
	{	
		Write-Host ""
		Write-Host "/!\ [$scriptName] An exception has been caught /!\ " -Foregroundcolor Red `r
		Write-Host "Type: " $_.Exception.GetType().FullName -Foregroundcolor Red `r
		Write-Host "Message: " $_.Exception.Message -Foregroundcolor Red `r
		Write-Host "Stacktrace: `n" $_.Exception.Stacktrace -Foregroundcolor Red `r
	}
	finally
	{
		if($curWeb -ne $null) 
		{
			$curWeb.Dispose();
		}
		Write-Debug "[$functionName] Exiting function" 
	}
	return $rowsCount;
}



<#

.SYNOPSIS
	Get Nintex Table 'Workflow Progress' Rows Count for a specific SPSite
	
.DESCRIPTION
	Get Nintex Table 'Workflow Progress' Rows Count  for a specific SPSite
	
.PARAMETER siteUrl
	(Mandatory) URL of the SharePoint Site (SPSite)
	
.EXAMPLE
	getSPSiteNintexWorklowsProgressDBCount -siteUrl <siteUrl>
	getSPSiteNintexWorklowsProgressDBCount -siteUrl http://intranet-assura
	
.OUTPUTS
	Integer : Item Count or -1 if an error occured
	
.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 29.03.2017
	Last Updated by: JBO
	Last Updated: 30.03.2017
	
#>
function getSPSiteNintexWorklowsProgressDBCount
{
	param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[ValidateNotNullOrEmpty()]
		[string]$siteURL
	)
	
	$functionName = $MyInvocation.MyCommand.Name
	Write-Debug "[$functionName] Entering function" 
	Write-Debug "[$functionName] Parameter / siteURL: $siteURL"
	
	try
	{
		$rowsCount = -1;
		$curSite = Get-SPSite $siteURL -ErrorAction SilentlyContinue
		if($curSite -ne $null) 
		{
			$siteID = $curSite.ID
			Write-Debug "" 
			Write-Debug "About to get Workflow Instances 'Table Rows' Count"
			$CFGDB = [Nintex.Workflow.Administration.ConfigurationDatabase]::OpenConfigDataBase().Database
			# Creating instance of .NET SQL client
			$cmd = New-Object -TypeName System.Data.SqlClient.SqlCommand
			$cmd.CommandType = [System.Data.CommandType]::Text

			# Begin SQL Query 
			$cmd.CommandText = "SELECT COUNT([WorkflowProgressID]) as 'Workflow Progress Count'	
								FROM [$CFGDB].[dbo].[WorkflowProgress] WP
								WHERE WP.[InstanceID] IN (SELECT [InstanceID]
															FROM  [$CFGDB].[dbo].[WorkflowInstance]
															WHERE SiteID = '"+$siteID+"');";
			
			$reader = [Nintex.Workflow.Administration.ConfigurationDatabase]::GetConfigurationDatabase().ExecuteReader($cmd)
			while($reader.Read())
			{
				$rowsCount = $reader["Workflow Progress Count"]
				Write-Debug "Workflow Progress (DB - Nintex) Count: $Count_WfProgress" 
			}
		}
		else
		{
			Write-Warning "[$functionName] Site '$siteURL' does not exist."
		}
	}
	catch [Exception]
	{	
		Write-Host ""
		Write-Host "/!\ [$scriptName] An exception has been caught /!\ " -Foregroundcolor Red `r
		Write-Host "Type: " $_.Exception.GetType().FullName -Foregroundcolor Red `r
		Write-Host "Message: " $_.Exception.Message -Foregroundcolor Red `r
		Write-Host "Stacktrace: `n" $_.Exception.Stacktrace -Foregroundcolor Red `r
	}
	finally
	{
		if($curSite -ne $null) 
		{
			$curSite.Dispose();
		}
		Write-Debug "[$functionName] Exiting function" 
	}
	return $rowsCount;
}




<#

.SYNOPSIS
	Clear Nintex Workflow History List-Items for all workflows in the state $status
	
.DESCRIPTION
	Clear Nintex Workflow History List-Items for all workflows in the state $status
	
.PARAMETER siteUrl
	(Mandatory) URL of the SharePoint Site (SPWeb)
		
.PARAMETER status
	(Mandatory) Status of the Workflows
	
.EXAMPLE
	clearWorkflowHistory -siteUrl <siteUrl> -status <'Completed','Cancelled','Error'>
	clearWorkflowHistory -siteUrl http://intranet-assura/FIN -status Completed

.OUTPUTS
	None
	
.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 29.03.2017
	Last Updated by: JBO
	Last Updated: 30.03.2017
	
#>
function clearWorkflowHistory
{
	param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[ValidateNotNullOrEmpty()]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[ValidateNotNullOrEmpty()]
		[ValidateSet('Completed','Cancelled','Error')]
		[string]$status
	)
	
	$functionName = $MyInvocation.MyCommand.Name
	Write-Debug "[$functionName] Entering function" 
	Write-Debug "[$functionName] Parameter / siteURL: $siteURL"
	
	try
	{
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if($curWeb -ne $null) 
		{
			Write-Host ""
			Write-Host "About to Clear Workflow History List"  -ForegroundColor Yellow  `r
			$cmdNWAdmin = "NWAdmin.exe -o PurgeHistoryListData -siteURL '" + $siteURL + "' -state " + $status + " -silent"
			Write-Host "$cmdNWAdmin" -ForegroundColor White `r
			$outputNWA = Invoke-Expression -Command:$cmdNWAdmin	
			if ($outputNWA -match "error")
			{
				throw $outputNWA
			}
			else
			{
				Write-Host "Workflow History List successfully cleared for '$status' workflows " -ForegroundColor Green `r	
			}	
		}
		else
		{
			Write-Warning "[$functionName] Site '$siteURL' does not exist."
		}
	}
	catch [Exception]
	{	
		Write-Host ""
		Write-Host "/!\ [$scriptName] An exception has been caught /!\ " -Foregroundcolor Red `r
		Write-Host "Type: " $_.Exception.GetType().FullName -Foregroundcolor Red `r
		Write-Host "Message: " $_.Exception.Message -Foregroundcolor Red `r
		Write-Host "Stacktrace: `n" $_.Exception.Stacktrace -Foregroundcolor Red `r
	}
	finally
	{
		if($curWeb -ne $null) 
		{
			$curWeb.Dispose();
		}
		Write-Debug "[$functionName] Exiting function" 
	}
}


<#

.SYNOPSIS
	Clear Nintex Workflow Data for all workflows in the state $status
	
.DESCRIPTION
	Clear Nintex Workflow Data for all workflows in the state $status
	
.PARAMETER siteUrl
	(Mandatory) URL of the SharePoint Site (SPWeb)
		
.PARAMETER status
	(Mandatory) Status of the Workflows
	
.EXAMPLE
	clearWorkflowData -siteUrl <siteUrl> -status <'Completed','Cancelled','Error'>
	clearWorkflowData -siteUrl http://intranet-assura/FIN -status Completed

.OUTPUTS
	None
	
.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 29.03.2017
	Last Updated by: JBO
	Last Updated: 30.03.2017
	
#>
function clearWorkflowData
{
	param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[ValidateNotNullOrEmpty()]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[ValidateNotNullOrEmpty()]
		[ValidateSet('Completed','Cancelled','Error')]
		[string]$status
	)
	
	$functionName = $MyInvocation.MyCommand.Name
	Write-Debug "[$functionName] Entering function" 
	Write-Debug "[$functionName] Parameter / siteURL: $siteURL"
	
	try
	{
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if($curWeb -ne $null) 
		{
			Write-Host "" 
			Write-Host "About to Clear Workflow History List for '$status' workflows."  -ForegroundColor Yellow  `r
			$cmdNWAdmin = "NWAdmin.exe -o PurgeWorkflowData -siteURL '"  + $siteURL + "' -state " + $status + " -silent"
			Write-Host "$cmdNWAdmin" -ForegroundColor White `r
			$outputNWA = Invoke-Expression -Command:$cmdNWAdmin	
			if ($outputNWA -match "error")
			{
				throw $outputNWA
			}
			else
			{
				Write-Host "Workflow Progress Data successfully cleared for '$status' workflows." -ForegroundColor Green `r	
			}	
		}
		else
		{
			Write-Warning "[$functionName] Site '$siteURL' does not exist."
		}
	}
	catch [Exception]
	{	
		Write-Host ""
		Write-Host "/!\ [$scriptName] An exception has been caught /!\ " -Foregroundcolor Red `r
		Write-Host "Type: " $_.Exception.GetType().FullName -Foregroundcolor Red `r
		Write-Host "Message: " $_.Exception.Message -Foregroundcolor Red `r
		Write-Host "Stacktrace: `n" $_.Exception.Stacktrace -Foregroundcolor Red `r
	}
	finally
	{
		if($curWeb -ne $null) 
		{
			$curWeb.Dispose();
		}
		Write-Debug "[$functionName] Exiting function" 
	}
}



<#

.SYNOPSIS
	Clear Nintex Workflow History List-Items for Deleted Lists
	
.DESCRIPTION
	Clear Nintex Workflow History List-Items for Deleted Lists
	
.PARAMETER siteUrl
	(Mandatory) URL of the SharePoint Site (SPWeb)

	
.EXAMPLE
	clearWorkflowHistoryForDeletedLists -siteUrl <siteUrl> 

.OUTPUTS
	None
	
.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 29.03.2017
	Last Updated by: JBO
	Last Updated: 30.03.2017
	
#>
function clearWorkflowHistoryForDeletedLists
{
	param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[ValidateNotNullOrEmpty()]
		[string]$siteURL
	)
	
	$functionName = $MyInvocation.MyCommand.Name
	Write-Debug "[$functionName] Entering function" 
	Write-Debug "[$functionName] Parameter / siteURL: $siteURL"
	
	try
	{
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if($curWeb -ne $null) 
		{
			Write-Host ""
			Write-Host "About to Clear Workflow History List for 'Deleted Lists'."  -ForegroundColor Yellow  `r
			$cmdNWAdmin = "NWAdmin.exe -o PurgeHistoryListData -siteURL '" + $siteURL + "' -DeletedLists -silent"
			Write-Host "$cmdNWAdmin" -ForegroundColor White `r
			$outputNWA = Invoke-Expression -Command:$cmdNWAdmin	
			if ($outputNWA -match "error")
			{
				throw $outputNWA
			}
			else
			{
				Write-Host "Workflow History List successfully cleared for 'Deleted Lists'." -ForegroundColor Green `r	
			}	
		}
		else
		{
			Write-Warning "[$functionName] Site '$siteURL' does not exist."
		}
	}
	catch [Exception]
	{	
		Write-Host ""
		Write-Host "/!\ [$scriptName] An exception has been caught /!\ " -Foregroundcolor Red `r
		Write-Host "Type: " $_.Exception.GetType().FullName -Foregroundcolor Red `r
		Write-Host "Message: " $_.Exception.Message -Foregroundcolor Red `r
		Write-Host "Stacktrace: `n" $_.Exception.Stacktrace -Foregroundcolor Red `r
	}
	finally
	{
		if($curWeb -ne $null) 
		{
			$curWeb.Dispose();
		}
		Write-Debug "[$functionName] Exiting function" 
	}
}


<#

.SYNOPSIS
	Clear Nintex Workflow Data for all Deleted Lists
	
.DESCRIPTION
	Clear Nintex Workflow Data for all Deleted Lists
	
.PARAMETER siteUrl
	(Mandatory) URL of the SharePoint Site (SPWeb)

	
.EXAMPLE
	clearWorkflowDataForDeletedLists -siteUrl <siteUrl> 

.OUTPUTS
	None
	
.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 29.03.2017
	Last Updated by: JBO
	Last Updated: 30.03.2017
	
#>
function clearWorkflowDataForDeletedLists
{
	param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[ValidateNotNullOrEmpty()]
		[string]$siteURL
	)
	
	$functionName = $MyInvocation.MyCommand.Name
	Write-Debug "[$functionName] Entering function" 
	Write-Debug "[$functionName] Parameter / siteURL: $siteURL"
	
	try
	{
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if($curWeb -ne $null) 
		{
			Write-Host "" 
			Write-Host "About to Clear Workflow History List for deleted lists"  -ForegroundColor Yellow  `r
			$cmdNWAdmin = "NWAdmin.exe -o PurgeWorkflowData -siteURL '" + $siteURL + "' -deletedLists -silent"
			Write-Host "$cmdNWAdmin" -ForegroundColor White `r
			$outputNWA = Invoke-Expression -Command:$cmdNWAdmin	
			if ($outputNWA -match "error")
			{
				throw $outputNWA
			}
			else
			{
				Write-Host "Workflow Progress Data successfully cleared for deleted lists." -ForegroundColor Green `r	
			}	
		}
		else
		{
			Write-Warning "[$functionName] Site '$siteURL' does not exist."
		}
	}
	catch [Exception]
	{	
		Write-Host ""
		Write-Host "/!\ [$scriptName] An exception has been caught /!\ " -Foregroundcolor Red `r
		Write-Host "Type: " $_.Exception.GetType().FullName -Foregroundcolor Red `r
		Write-Host "Message: " $_.Exception.Message -Foregroundcolor Red `r
		Write-Host "Stacktrace: `n" $_.Exception.Stacktrace -Foregroundcolor Red `r
	}
	finally
	{
		if($curWeb -ne $null) 
		{
			$curWeb.Dispose();
		}
		Write-Debug "[$functionName] Exiting function" 
	}
}