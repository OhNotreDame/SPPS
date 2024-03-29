##############################################################################################
#              
# NAME: SPFileUploader.psm1 
# PURPOSE: Upload any file into SharePoint Site
#	Relies on an XML Configuration file to identify which file to upload where.
#	See SPFileUploader.xml for Schema
#
# SOURCE : https://github.com/OhNotreDame/SPPS
#
##############################################################################################


<#
.SYNOPSIS
	Browse and parse the uploadFilesXML XML object
	
.DESCRIPTION
	Browse and parse the uploadFilesXML XML object in order to 
	upload all referenced files into the right location (library, folder, subfolder, ...) of site siteURL.
		
.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER sourceFolderPath
	Root Location/Path where are stored the files to be uploaded
	
.PARAMETER uploadFilesXML
	XML Object representing the list of file to upload
	
.EXAMPLE
	browseFilesAndFoldersToUpload -siteURL <SiteURL> -uploadFilesXML <XMLObjectToParse>
	
.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 13.01.2017
	Last Updated by: JBO
	Last Updated: 13.01.2017
#>
function browseFilesAndFoldersToUpload()
{
	[CmdletBinding()]
	Param
	(	
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,		
		[Parameter(Mandatory=$true, Position=2)]
		[string]$sourceFolderPath,
		[Parameter(Mandatory=$true, Position=3)]
		[XML]$uploadFilesXML
	)
    
	$functionName = $MyInvocation.MyCommand.Name
    Write-Debug "[$functionName] Entering function" 
    Write-Debug "[$functionName] Parameter / siteURL: $siteURL" 
    Write-Debug "[$functionName] Parameter / sourceFolderPath: $sourceFolderPath" 
	
	try
    {
		Write-Debug "[$functionName] Parameter/ siteURL: $siteURL"		
		foreach($file in $uploadFilesXML.FilesUpload.File)
		{
			$targetLibName = $file.TargetLibrary.Trim()
			$targetFolder = $file.TargetFolder.Trim()
			$fileName = $file.Name.Trim()
			$sourceFolder = $file.SourceFolder.Trim()
			
			# Building Final File Path
			if ([string]::IsNullOrEmpty($sourceFolder)){
				$finalFilePath = $sourceFolderPath +"\"+ $fileName
			}
			else{
				$finalFilePath = $sourceFolderPath +"\"+ $sourceFolder +"\"+ $fileName
			}
			
			if(Test-Path $finalFilePath)	
			{
				uploadFile -siteURL $siteURL -targetLibName $targetLibName -targetPath $targetFolder -fileName $fileName -filePath $finalFilePath
			}
			else	
			{
				Write-Warning "[$functionName] File '$fileName' does not exist at this location $finalFilePath."
			}			
		}#end foreach
    }
    catch [exception]
    {
		Write-Host "/!\ [$functionName] An exception has been caught /!\ "  -ForegroundColor Red `r
		Write-Host "Type: " $_.Exception.GetType().FullName   -ForegroundColor Red `r
		Write-Host "Message: " $_.Exception.Message  -ForegroundColor Red `r
		Write-Host "Stacktrace: " $_.Exception.Stacktrace  -ForegroundColor Red `r
    }
	finally
	{
		Write-Debug "[$functionName] Exiting function" 
	}
}


<#

.SYNOPSIS
	Browse and parse the uploadFilesXML XML object
	
.DESCRIPTION
	Browse and parse the uploadFilesXML XML object
	Upload all references files into the right location (library, folder, subfolder, ...) of site siteURL.

.PARAMETER siteUrl
URL of the SharePoint Site
				
.PARAMETER targetLibName
Target library name of file to be uploaded on site

.PARAMETER targetPath
Target folder of file to be uploaded on site (could be empty if the file is supposed to be uploaded at RootFolder)

.PARAMETER fileName
Name of file to be uploaded

.PARAMETER	filePath
Location/Physical path of the file to be uploaded
	
.EXAMPLE
	uploadFile -siteURL <SiteURL> -targetLibName <LibraryName> -targetPath <FolderName> -fileName <FileName> -filePath <filePath>
	
.OUTPUTS
	None

.LINK
	None
	
.NOTES
	Created by: JBO
	Created: 13.01.2017
	Last Updated by: JBO
	Last Updated: 13.01.2017

#>
function uploadFile()
{
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[string]$targetLibName,
		[Parameter(Mandatory=$false, Position=3)]
		[AllowEmptyString()]
		[string]$targetPath,
		[Parameter(Mandatory=$true, Position=4)]
		[string]$fileName,
		[Parameter(Mandatory=$true, Position=5)]
		[string]$filePath
	)
    
	$functionName = $MyInvocation.MyCommand.Name
	Write-Debug "[$functionName] Entering function" 
	Write-Debug "[$functionName] Parameter siteURL: $siteURL" 
	Write-Debug "[$functionName] Parameter targetLibName: $targetLibName" 
	Write-Debug "[$functionName] Parameter targetPath: $targetPath" 
	Write-Debug "[$functionName] Parameter fileName: $fileName" 
	Write-Debug "[$functionName] Parameter filePath: $filePath" 

    try
    {
        $curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
        if($curWeb -ne $null)
        {
            $targetLibrary = $curWeb.Lists.TryGetList($targetLibName)
	        if($targetLibrary -ne $null)
	        {
				# Write-Host "[$functionName] Library:  $targetLibName "
				# Write-Host "[$functionName] Versioning enabled: " $targetLibrary.EnableVersioning
				# Write-Host "[$functionName] Minor Versioning enabled: " $targetLibrary.EnableMinorVersions
				
				# Open file
				$fileStream = ([System.IO.FileInfo] (Get-Item $filePath)).OpenRead()				
                if($fileStream -ne $null -and $fileStream.Length -gt 0)
                {
                    # Get the target folder.
                    $folderUrl= [String]::Format("{0}/{1}/{2}", $curWeb.Url,$targetLibrary.RootFolder.Url,$targetPath)			
					#Write-Host "$functionName Computed/ folderUrl: $folderUrl" `r
					
					if([string]::IsNullOrEmpty($targetPath)) 
					{   
						# Build final File Path	
						$finalFilePath = $targetLibrary.RootFolder.Url+ "/" + $fileName
						
						# Check & Setup folder hierarchy
						$newFolder = $targetLibrary.RootFolder.ServerRelativeUrl			
					}					
					else
					{
						# Build final File Path	
						$finalFilePath = $targetLibrary.RootFolder.Url+ "/" + $targetPath + "/" + $fileName
						
						# Check & Setup folder hierarchy
						$destinationFolder = ensureListFolderTree -siteURL $siteURL -listName $targetLibName -folderName $targetPath
					}
					
					Write-Debug "[$functionName] finalFilePath: $finalFilePath" 
				   
					# Copying file $fileName to $targetLibName...
					#$destinationFolder = $curWeb.GetFolder($newFolder.Folder.ServerRelativeUrl)

					$fileToBePublished = $curWeb.GetFile($finalFilePath)
					if ($fileToBePublished.Exists)
					{
						Write-Warning "[$functionName] File '$fileName' already exist on site." 
						if(($fileToBePublished.Level -eq "Checkout"))
						{
							Write-Warning "[$functionName] File '$fileName' already checked-out by another user." 
							Write-Host "[$functionName] About to discard Check-Out on '$fileName'." -ForegroundColor Magenta `r
							$fileToBePublished.UndoCheckOut() 
							Write-Host "[$functionName] Check-Out discarded by Powershell on File '$fileName' ." -ForegroundColor Green `r
						}
						
						#Write-Host "[$functionName] About to check-out '$fileName'." -ForegroundColor Magenta `r
						#$fileToBePublished.CheckOut()
						#Write-Host "[$functionName] File '$fileName' checked-out by Powershell." -ForegroundColor Green `r
					}
					
					Write-Host "[$functionName] About to upload" -ForegroundColor Cyan `r
					Write-Host "... file: $fileName `n... into: $newFolder." -ForegroundColor Cyan `r
				    $spFile = $destinationFolder.Files.Add($finalFilePath, [System.IO.Stream]$fileStream, $true)
                    Write-Host "[$functionName] File '$fileName' sucessfully uploaded." -ForegroundColor Green `r
 
				

					if(($spFile.Level -eq "Checkout"))
					{
						Write-Host "[$functionName] About to check-in '$fileName'." -ForegroundColor Magenta `r
						$spFile.CheckIn("File Checked-In by Powershell", [Microsoft.SharePoint.SPCheckinType]::MajorCheckIn);
						Write-Host "[$functionName] File '$fileName' checked-in by Powershell." -ForegroundColor Green `r
					}
 
					if (($targetLibrary.EnableMinorVersions -eq $true) -and ($spFile.Level -ne "Published") )
					{
						Write-Host "[$functionName] About to Publish '$fileName'." -ForegroundColor Magenta `r
						$spFile.Publish("File Published by Powershell");
						Write-Host "[$functionName] File '$fileName' Published by Powershell." -ForegroundColor Green `r
					}

				    #Close file stream
				    $fileStream.Close();
                }
                else
                {
					Write-Warning "[$functionName] Empty file object." 
                }		
	        }
			else
			{
		        Write-Warning "[$functionName] Library '$targetLibName' not found on site '$siteURL'." 
		        return;
	        }
        }
		else
		{
			Write-Warning "[$functionName] Site '$siteURL' not found." 
		    return;
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
	Checks the target folder existence within the list, if not creates it.
	
.DESCRIPTION
	Will create (if not exist) the folder folder name within the list (or library) listName within site siteURL.
	This function is recursive and will Check & Setup the complete folder hierarchy (if needed).
		
.PARAMETER siteUrl
	URL of the SharePoint Site
	
.PARAMETER listName
	Target list/library name

.PARAMETER folderName
	Target folder name
	
.EXAMPLE
	ensureListFolderTree -siteURL <SiteURL> -listName <ListName> -folderName <folderName>
	
.OUTPUTS
	None

.LINK
	Source code inspired by https://platinumdogs.me/2014/07/23/create-folder-hierarchies-for-sharepoint-lists-and-libraries-using-powershell/
	
.NOTES
	Created by: JBO
	Created: 13.01.2017
	Last Updated by: JBO
	Last Updated: 13.01.2017
#>
function ensureListFolderTree()
{
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory=$true, Position=1)]
		[string]$siteURL,
		[Parameter(Mandatory=$true, Position=2)]
		[AllowEmptyString()]
		[string]$listName,
		[Parameter(Mandatory=$true, Position=3)]
		[AllowEmptyString()]
		[string]$folderName
	)

	$functionName = $MyInvocation.MyCommand.Name
	Write-Debug "[$functionName] Entering function" 
    try
    {
		Write-Host "siteURL: $siteURL"
		$curWeb = Get-SPWeb $siteURL -ErrorAction SilentlyContinue
		if ($curWeb -ne $null)
		{
			#Get the list
			$list = $curWeb.Lists.TryGetList($listName)
			if ($list -ne $null)
			{
				$parentUrl = $list.RootFolder.ServerRelativeUrl		
				Write-Host "parentUrl: $parentUrl"
				$folders = $folderName.Replace("^/+","").Split("/")
	
				foreach($curFolderName in $folders) 
				{
					#Write-Host "curFolderName: $curFolderName" 
					$folderUrl = $parentUrl + "/" + $curFolderName					
					#Write-Host "folderUrl: $folderUrl" 
					$folder = $curWeb.GetFolder($folderUrl)
					
					if ($folder.Exists -eq $false) 
					{
						
						$folder = $list.AddItem($parentUrl, [Microsoft.SharePoint.SPFileSystemObjectType]::Folder, $curFolderName)
						$folder.Update()
						$folder = $curWeb.GetFolder($folder.Folder.ServerRelativeUrl)
					}
					else
					{
						Write-Debug "[$functionName] Folder '$curFolderName' already exists in list '$listName'." 
					}
					$parentUrl = $folder.ServerRelativeUrl				
				} #end foreach	
			   $folder;			
			}
			else
			{
				Write-Warning "[$functionName] List '$listName' could not be found on site '$siteURL'."
			}
		}
		else
		{
			Write-Warning "[$functionName] Site '$siteURL' could not be found."
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
		
		Write-Debug "[$functionName] Entering function" 
    }
}