# SPPS
Powershell scripts to interact with SharePoint


##  Refresh SPFields on SPList using refreshSiteColumsOnList()
Refresh all the fields of a list based on their Site Column definition
```Powershell
refreshSiteColumsOnList.ps1 -siteURL "http://spweb_url" -listName "name of the list"
```

##  Upload files on site using browseFilesAndFoldersToUpload()
Relies on an XML File SPFileUploader.xml describing the list and location of file to be uploaded.
```XML
<?xml version="1.0" encoding="utf-8"?>
<FilesUpload>
	<File Name="fileA.master" SourceFolder="PhysicalPath\MPs" TargetLibrary="Master Page Gallery" TargetFolder="CustomFolder"></File>
	<File Name="fileA.html" SourceFolder="PhysicalPath\MPs" TargetLibrary="Master Page Gallery" TargetFolder="CustomFolder"></File>
</FilesUpload>
```

To ensure that these files are uploaded, you need to call browseFilesAndFoldersToUpload().
This function will browse and parse the SPFileUploader.xml and initiate the upload of all referenced files into the right location (library, folder, subfolder, ...) on site siteURL.

```Powershell
Import-Module "$PathToModuleFolder\SPHelpers\SPHelpers.psm1"
Import-Module "PathToModuleFolder\SPFileUploader\SPFileUploader.psm1"
$fileToUploadFilePath = "configFileToUpload_FilePath\SPFileUploader.xml"
$fileToUploadLocation = "uploadFilesLocation\SPFileUploader\"
if(Test-Path $fileToUploadFilePath)
{
  $uploadFilesXML = LoadXMLFile -xmlPath  $fileToUploadFilePath
  if($uploadFilesXML -ne $null -and $uploadFilesXML.HasChildNodes)
  {
    browseFilesAndFoldersToUpload -siteURL $siteURL -sourceFolderPath $fileToUploadLocation -uploadFilesXML $uploadFilesXML 
  }
  else
  {
    Write-Warning "XML File for <SPFileUploader> is empty." 
  }
}
else
{
  Write-Warning "XML File for <SPFileUploader> does not exist."
}
```
