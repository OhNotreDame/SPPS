<#

.SYNOPSIS
	Remove missing assembly from a Content-Database in SharePoint
	
.DESCRIPTION
	Remove missing assembly  from a Content-Database in SharePoint
	If ReportOnly is used, will generate a report on assembly usage accross Content-Database

.PARAMETER serverName [Mandatory]
	Name of the SQL Server

.PARAMETER filePath [Mandatory]
	FilePath of the list of Missing Assemblies

.PARAMETER ContentDb [Mandatory]
	Name of the Content DB
	
.PARAMETER ReportOnly [Optional]
	Used to generate a report
	
.EXAMPLE
	.\RemoveAssemblyFromContentDB.ps1 -serverName "DBServer\DBInstance" -filePath D:\Scripts\RemoveAssemblyFromContentDB\MissingAssembly-PWA.txt -ContentDb "WSS_CONTENT_XYZ" -ReportOnly
	.\RemoveSPFeatureFromContentDB.ps1 -serverName "DBServer\DBInstance" -filePath D:\Scripts\RemoveAssemblyFromContentDB\MissingAssembly-PWA.txt -ContentDb "WSS_CONTENT_XYZ"
	
.OUTPUTS
	1 CSV file [NC: RemoveAssemblyFromContentDB_<contentDB>_<TIMESTAMPE>_.csv]

.LINK
	https://github.com/OhNotreDame/SPPS
	
.NOTES
	Created by: JBO
	Created: 17.07.2017
	Last Updated by: JBO
	Last Updated:  17.07.2017
	Inspired by http://etienne-sharepoint.blogspot.ch/2011/10/solving-missingassembly-errors-from.html
	
#>

param
(
	[Parameter(Mandatory=$true, Position=1)]
	[String]$serverName, 
	[Parameter(Mandatory=$true, Position=2)]
	[String]$filePath,
	[Parameter(Mandatory=$true, Position=3)]
	[String]$contentDB,
    [Parameter(Mandatory=$false, Position=4)]
	[switch]$ReportOnly
)

Clear-Host
Remove-Module *

function CreateRow ($Assembly, $ERName, $ERClass, $Object, $Name, $URL)
{
    $objRow = New-Object System.Object
    $objRow | Add-Member -type NoteProperty -name Assembly -value $Assembly
    $objRow | Add-Member -type NoteProperty -name EvtReceiver -value $ERName
    $objRow | Add-Member -type NoteProperty -name Class -value $ERClass
    $objRow | Add-Member -type NoteProperty -name Object -value $Object
    $objRow | Add-Member -type NoteProperty -name ObjName -value $Name
    $objRow | Add-Member -type NoteProperty -name ObjURL -value $URL
    return $objRow
}


#Declare SQL Query function
function Run-SQLQuery ($SqlServer, $SqlDatabase, $SqlQuery)
{
	$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
	$SqlConnection.ConnectionString = "Server =" + $SqlServer + "; Database =" + $SqlDatabase + "; Integrated Security = True"
	$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
	$SqlCmd.CommandText = $SqlQuery
	$SqlCmd.Connection = $SqlConnection
	$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
	$SqlAdapter.SelectCommand = $SqlCmd
	$DataSet = New-Object System.Data.DataSet
	$SqlAdapter.Fill($DataSet)
	$SqlConnection.Close()
	$DataSet.Tables[0]
}


function GetAssemblyDetails ($assembly, $DBname)
{

    #Define SQL Query and set in Variable
	$Query = "SELECT * from EventReceivers where Assembly = '"+$assembly+"'"

	#Runing SQL Query to get information about Assembly (looking in EventReceiver Table) and store it in a Table
	$QueryReturn = @(Run-SQLQuery -SqlServer $ServerName -SqlDatabase $DBname -SqlQuery $Query | select Id, Name, SiteId, WebId, HostId, HostType)

	#Actions for each element in the table returned
	foreach ($event in $QueryReturn)
	{   

        #Write-Host "Row: $($event.Id) - $($event.Name) - $($event.SiteId) - $($event.WebId)  - $($event.HostId) - $($event.HostType)"
	    #SPSite Event Receiver
		if ($event.HostType -eq 0)
		{
			#$site = Get-SPSite -Limit all | where {$_.Id -eq $event.SiteId}
            $site = Get-SPSite -Identity $event.SiteId
			if ($site)
            {
                #Get the EventReceiver Site Object
			    $er = $site.EventReceivers | where {$_.Id -eq $event.Id}

			    #Write-Host "SPSite:  $($site.Url) - ER: $($er.Name) - ERClass:  $($er.Class) `n##"			
                $global:allRows += CreateRow $assembly $er.Name $er.Class "SPSite" $site.RootWeb.Title $site.Url
                if ($report -eq $false) {
			        $er.Delete()
                }
            }
		}

        #SPWeb Event Receiver
		if ($event.HostType -eq 1)
		{
		    $site = Get-SPSite -Identity $event.SiteId
		    if ($site)
            {                         
                $web = $site | Get-SPWeb -Identity $event.WebId            
                if ($web)
                {
                    #Get the EventReceiver Site Object
		            $er = $web.EventReceivers | where {$_.Id -eq $event.Id}
		            $er.Name

		            #Write-Host "SPWeb:  $($site.Url) - ER: $($er.Name) - ERClass:  $($er.Class) `n##"			
                    $global:allRows += CreateRow $assembly $er.Name $er.Class "SPWeb" $web.Title $web.Url
		            if ($report -eq $false) {
			            $er.Delete()
                    }
                }
            }
		}

        #SPList Event Receiver
		if ($event.HostType -eq 2)
		{
			$site = Get-SPSite -Identity $event.SiteId
			if ($site)
            {                         
                $web = $site | Get-SPWeb -Identity $event.WebId      
			    if ($web)
                {
                    $list = $web.Lists | where {$_.Id -eq $event.HostId}
			        if ($list)
                    {
                        #Get the EventReceiver List Object
			            $er = $list.EventReceivers | where {$_.Id -eq $event.Id}
                        $listURL = $web.Url+"/"+$list.RootFolder

		           	    #Write-Host "SPList:  $listURL - ER: $($er.Name) - ERClass:  $($er.Class) `n##"	
                        $global:allRows += CreateRow $assembly $er.Name $er.Class "SPList" $list.Title $listURL
		                if ($report -eq $false) {
			                $er.Delete()
                        }
                    }
                }
            }
		}

	}
}


$scriptName = $([System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Definition))
Write-Host "************************************************************************" -ForegroundColor Gray `r
Write-Host "$scriptName # Script started." -ForegroundColor Gray `r
Write-Host "************************************************************************" -ForegroundColor Gray `r		
Write-Host "PARAM/serverName : $serverName" -ForegroundColor Gray `r
Write-Host "PARAM/filePath : $filePath" -ForegroundColor Gray `r
Write-Host "PARAM/contentDB : $contentDB" -ForegroundColor Gray `r
Write-Host "PARAM/ReportOnly : $ReportOnly" -ForegroundColor Gray `r
Write-Host "************************************************************************" -ForegroundColor Gray `r		


# Loading SharePoint Assembly and PS Snapin
Add-PsSnapin Microsoft.Sharepoint.PowerShell -EA silentlycontinue
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint") | Out-Null
[void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Web.Administration")

# Starting SPAssignment
Start-SPAssignment -Global

# CSV file Settings
$fileName = $scriptName+"_"+$contentDB+"_"+"$(get-date -format 'yyyyMMddHHmm').csv"
$fileDestLocation = "D:\Scripts\$scriptName"

try
{
	[bool]$report = $false
    if ($ReportOnly) { $report = $true }
    
    #Intiate Results Object
	$global:allRows = @()
	
	if(Test-Path $filePath)
	{
		$fileContent = @(Get-Content $filePath)
		if($fileContent -ne $null)
		{
			foreach ($missingAssembly in $fileContent)
			{
				Write-Host "$missingAssembly"  -foregroundcolor Magenta
				Write-Host "> Looking for it in '$contentDB'"
				GetAssemblyDetails $missingAssembly $contentDB
			}					
		}
		else
		{
			Write-Warning "File <$filePath> is empty." 
		}
	}
	else
	{
		Write-Warning "File <$filePath> does not exist."
	}
}
catch
{
	Write-Host ""
	Write-Host "/!\ [$scriptName] An exception has been caught /!\ " -Foregroundcolor Red `r
	Write-Host "Type: " $_.Exception.GetType().FullName -Foregroundcolor Red `r
	Write-Host "Message: " $_.Exception.Message -Foregroundcolor Red `r
	Write-Host "Stacktrace: `n" $_.Exception.Stacktrace -Foregroundcolor Red `r
}
finally
{
	$allRows

    # Export CSV file
	$allRows | Export-Csv "$fileDestLocation\$fileName" -NoTypeInformation -Delimiter ";" -Encoding UTF8
	Write-Host "CSV file available at $fileDestLocation\$fileName"  -ForegroundColor Green `r

    #Dispose SPSite
    if ($site -ne $null) {
        $site.Dispose();
    }

    #Dispose SPWeb
    if ($web -ne $null) {
        $web.Dispose();
    }
	# Stopping SPAssignment and Transcript
	Stop-SPAssignment -Global

	Write-Host "************************************************************************" -ForegroundColor Gray `r		
	Write-Host "$scriptName # Script ended." -ForegroundColor Gray `r
	Write-Host "************************************************************************" -ForegroundColor Gray `r					
}



