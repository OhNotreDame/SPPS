<#
.SYNOPSIS
	Start the Site Workflow $workflowName using Nintex Webservice (NintexWorkflow/Workflow.asmx)
	
.DESCRIPTION
	Start the Site Workflow $workflowName using Nintex Webservice (NintexWorkflow/Workflow.asmx)
	
.PARAMETER siteUrl
	[string] URL of the SharePoint Site

.PARAMETER workflowName
	[string] Name of the workflow to start
	
.PARAMETER workflowData
	[string] Optional Workflow data
	
.EXAMPLE
	startSiteWorkflow -siteURL <siteURL> -workflowName <workflowName> [-workflowData <workflowData>]
	startSiteWorkflow -siteURL http://sharepoint_siteurl/sites/debug -workflowName "Send Mail Site"
	
.OUTPUTS
	None

.LINK
	http://vadimtabakman.com/nintex-workflow-start-workflow-with-powershell.aspx
	
.NOTES
	Created by: JBO
	Created: 13.03.2017
	Last Updated by: JBO
	Last Updated: 13.03.2017
#>
function startSiteWorkflow()
{
    Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[string]$workflowName,
		[Parameter(Mandatory=$false, Position=3)]
		[string]$workflowData
	)

	$functionName = $MyInvocation.MyCommand.Name	
	Write-Debug "Entering $functionName"
	
    try
    {
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if ($curWeb -ne $null)
		{
			$curWeb.AllowUnsafeUpdates = $true;
			$parentSite = $curWeb.Site
			if ($parentSite -ne $null)
			{
				$workflowManager = $parentSite.WorkFlowManager;
				if ($workflowManager  -ne $null)
				{
					#################
					# ToDo: Manage WF Association Data (ie WF Parameters, if any)
					#################
					$transferData ="Test By Powershell"
					[String]$associationData = "<Data>" + $workflowData + "</Data>"

					#Step1: WS address URI
					$proxyWSURL= [Microsoft.SharePoint.Utilities.SPUtility]::ConcatUrls($siteURL, "_vti_bin/NintexWorkflow/workflow.asmx")
					
					# $proxyWSUri = New-Object System.Uri($proxyWSURL)
					# $credCache = New-Object System.Net.CredentialCache
					# $cred = [System.Net.CredentialCache]::DefaultNetworkCredentials
					# $credCache.Add($proxyWSUri, "NTLM", [System.Net.CredentialCache]::DefaultNetworkCredentials)
					
					#Step2: Create WS Proxy
					$proxy = New-WebServiceProxy -Uri $proxyWSURL -UseDefaultCredential
					$proxy.Url = $proxyWSUrl
					
					#Step3: Call WS Method 'StartWorkflowOnListItem'
					Write-Host "About to start Workflow '$workflowName' on site '$siteURL'" -ForegroundColor Magenta `r 
					$workflow = $proxy.StartSiteWorkflow($workflowName,$associationData)
					Write-Host "Workflow '$workflowName' started on site '$siteURL'." -ForegroundColor Green `r 

				}
				else
				{
					Write-Warning "[$functionName] Workflow Manager object is null."
				}
			}
			else
			{
				Write-Warning "[$functionName] Parent SPSite not found."
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
			$curWeb.AllowUnsafeUpdates = $false;
			$curWeb.Dispose()
        }
		Write-Debug "Exiting $functionName"
    }
}


<#
.SYNOPSIS
	Start the List Workflow $workflowName on the list item $itemID using Nintex Webservice (NintexWorkflow/Workflow.asmx)
	
.DESCRIPTION
	Start the List Workflow $workflowName on the list item $itemID using Nintex Webservice (NintexWorkflow/Workflow.asmx)
	
.PARAMETER siteUrl
	[string] URL of the SharePoint Site

.PARAMETER workflowName
	[string] Name of the workflow to start

.PARAMETER listName
	[string] Name of the list where the item is located

.PARAMETER itemID
	[string] ID of the item where the workflow should be started
	
.PARAMETER workflowData
	[string] Optional Workflow data
	
.EXAMPLE
	startListWorkflow -siteURL <siteURL> -workflowName <workflowName> -listName <listName> -itemID <itemID> [-workflowData <workflowData>]
	startListWorkflow -siteURL http://sharepoint_siteurl/sites/debug -workflowName "Send Mail" -listName "Tasks" -itemID 1
	
.OUTPUTS
	None

.LINK
	http://vadimtabakman.com/nintex-workflow-start-workflow-with-powershell.aspx
	
.NOTES
	Created by: JBO
	Created: 13.03.2017
	Last Updated by: JBO
	Last Updated: 13.03.2017
#>
function startListWorkflow()
{
    Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[string]$workflowName,
		[Parameter(Mandatory=$true, Position=3)]
		[string]$listName,
		[Parameter(Mandatory=$true, Position=4)]
		[string]$itemID,
		[Parameter(Mandatory=$false, Position=5)]
		[string]$workflowData
	)

	$functionName = $MyInvocation.MyCommand.Name	
	Write-Debug "Entering $functionName"
	
    try
    {
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if ($curWeb -ne $null)
		{
			$curWeb.AllowUnsafeUpdates = $true;
			$parentSite = $curWeb.Site
			if ($parentSite -ne $null)
			{
				$workflowManager = $parentSite.WorkFlowManager;
				if ($workflowManager  -ne $null)
				{
				
					$list=$curWeb.Lists.TryGetList($listName)
					if ($list -ne $null)
					{
						#Step 1: Get Workflow
						$wfAssociation=$list.WorkFlowAssociations | where {$_.Name -eq $workflowName};
						$wfAssociation.AllowAsyncManualStart = $true
						$wfAssociation.AllowManual = $true
						
						$data = $wfAssociation.AssociationData;
						
						#Step 2: Get Item and Start Workflow
						$item = $list.GetItemById($itemID);
						if ($item -ne $null)
						{
							#################
							# ToDo: Manage WF Association Data (ie WF Parameters, if any)
							#################
							$transferData =""
							[String]$associationData = "<Data><inputXML>" + $transferData + "</inputXML></Data>"

							#Step1: WS address URI
							$proxyWSURL= [Microsoft.SharePoint.Utilities.SPUtility]::ConcatUrls($siteURL, "_vti_bin/NintexWorkflow/workflow.asmx")
							
							# $proxyWSUri = New-Object System.Uri($proxyWSURL)
							# $credCache = New-Object System.Net.CredentialCache
							# $cred = [System.Net.CredentialCache]::DefaultNetworkCredentials
							# $credCache.Add($proxyWSUri, "NTLM", [System.Net.CredentialCache]::DefaultNetworkCredentials)
							
							#Step2: Create WS Proxy
							$proxy = New-WebServiceProxy -Uri $proxyWSURL -UseDefaultCredential
							$proxy.Url = $proxyWSUrl
							
							#Step3: Call WS Method 'StartWorkflowOnListItem'
							Write-Host "About to start Workflow '$workflowName' on Item '$itemID' of List '$listName'" -ForegroundColor Magenta `r 
							$workflow = $proxy.StartWorkflowOnListItem($itemID,$list,$workflowName,$associationData)
							Write-Host "Workflow '$workflowName' started on Item '$itemID' of List '$listName'" -ForegroundColor Green `r 

						}				
						else
						{
							Write-Warning "[$functionName] Item '$itemID' not found on list '$listName'."
						}					
					}				
					else
					{
						Write-Warning "[$functionName] List '$listName' not found on site '$siteURL'."
					}
				}
				else
				{
					Write-Warning "[$functionName] Workflow Manager object is null."
				}
			}
			else
			{
				Write-Warning "[$functionName] Parent SPSite not found."
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
			$curWeb.AllowUnsafeUpdates = $false;
			$curWeb.Dispose()
        }
		Write-Debug "Exiting $functionName"
    }
}

<#
.SYNOPSIS
	Start the List Workflow $workflowName on the list item $itemID using Worflow Manager
	
.DESCRIPTION
	Start the List Workflow $workflowName on the list item $itemID using Worflow Manager
	
.PARAMETER siteUrl
	[string] URL of the SharePoint Site

.PARAMETER workflowName
	[string] Name of the workflow to start

.PARAMETER listName
	[string] Name of the list where the item is located

.PARAMETER itemID
	[string] ID of the item where the workflow should be started
	
.EXAMPLE
	startListWorkflowViaWorkflowManager -siteURL <siteURL> -workflowName <workflowName> -listName <listName> -itemID <itemID>
	startListWorkflowViaWorkflowManager -siteURL http://sharepoint_siteurl/sites/debug -workflowName "Send Mail" -listName "Tasks" -itemID 1
	
.OUTPUTS
	None

.LINK
	http://vadimtabakman.com/nintex-workflow-start-workflow-with-powershell.aspx
	
.NOTES
	Created by: JBO
	Created: 13.03.2017
	Last Updated by: JBO
	Last Updated: 13.03.2017
#>
function startListWorkflowViaWorkflowManager()
{
    Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[string]$workflowName,
		[Parameter(Mandatory=$true, Position=3)]
		[string]$listName,
		[Parameter(Mandatory=$true, Position=4)]
		[string]$itemID
	)

	$functionName = $MyInvocation.MyCommand.Name	
	Write-Debug "Entering $functionName"
	
    try
    {
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if ($curWeb -ne $null)
		{
			$curWeb.AllowUnsafeUpdates = $true;
			$parentSite = $curWeb.Site
			if ($parentSite -ne $null)
			{
				$workflowManager = $parentSite.WorkFlowManager;
				if ($workflowManager  -ne $null)
				{
				
					$list=$curWeb.Lists.TryGetList($listName)
					if ($list -ne $null)
					{
						#Step 1: Get Workflow
						$wfAssociation=$list.WorkFlowAssociations | where {$_.Name -eq $workflowName};
						$wfAssociation.AllowAsyncManualStart = $true
						$wfAssociation.AllowManual = $true
						
						$data = $wfAssociation.AssociationData;
						
						#Step 2: Get Item and Start Workflow
						$item = $list.GetItemById($itemID);
						if ($item -ne $null)
						{
							Write-Host "About to start Workflow '$workflowName' on Item '$itemID' of List '$listName'" -ForegroundColor Magenta `r 
							#$wfRunOption=[Microsoft.Sharepoint.Workflow.SPWorkflowRunoptions]::synchronous
							#$wf = $workflowManager.StartWorkFlow($item,$wfAssociation,$data, $wfRunOption);
							$wf = $workflowManager.StartWorkFlow($item,$wfAssociation,$data, $true);
							Write-Host "Workflow '$workflowName' started on Item '$itemID' of List '$listName'" -ForegroundColor Green `r 

						}				
						else
						{
							Write-Warning "[$functionName] Item '$itemID' not found on list '$listName'."
						}					
					}				
					else
					{
						Write-Warning "[$functionName] List '$listName' not found on site '$siteURL'."
					}
				}
				else
				{
					Write-Warning "[$functionName] Workflow Manager object is null."
				}
			}
			else
			{
				Write-Warning "[$functionName] Parent SPSite not found."
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
			$curWeb.AllowUnsafeUpdates = $false;
			$curWeb.Dispose()
        }
		Write-Debug "Exiting $functionName"
    }


}
