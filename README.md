# SPPS
Powershell scripts to interact with SharePoint


##  Refresh SPFields on SPList using refreshSiteColumsOnList.ps1
Refresh all the fields of a list based on their Site Column definition
```Powershell
refreshSiteColumsOnList.ps1 -siteURL "http://spweb_url" -listName "name of the list"
```

##  SPFileUploader

### XML File Structure
Relies on an XML File SPFileUploader.xml describing the list and location of file to be uploaded.
```XML
<?xml version="1.0" encoding="utf-8"?>
<FilesUpload>
	<File Name="fileA.master" SourceFolder="PhysicalPath\MPs" TargetLibrary="Master Page Gallery" TargetFolder="CustomFolder"></File>
	<File Name="fileA.html" SourceFolder="PhysicalPath\MPs" TargetLibrary="Master Page Gallery" TargetFolder="CustomFolder"></File>
</FilesUpload>
```

### Upload files on site using browseFilesAndFoldersToUpload()
To ensure that these files are uploaded, you need to call browseFilesAndFoldersToUpload().
This function will browse and parse the SPFileUploader.xml and initiate the upload of all referenced files into the right location (library, folder, subfolder, ...) on site siteURL.

```Powershell
Import-Module "pathTo_ModuleFolder\SPHelpers\SPHelpers.psm1"
Import-Module "pathTo_ModuleFolder\SPFileUploader\SPFileUploader.psm1"
# Configuration file 'SPFileUploader.xml'
$fileToUploadFilePath = "pathTo_ConfigFile\SPFileUploader.xml"
# Location of the file to be uploaded
$fileToUploadLocation = "pathTo_SourceFolder\SPFileUploader\"
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

## SPListViews

### XML File Structure
```XML
<?xml version="1.0" encoding="utf-8" ?>
<SPListViews>
 <AddView>
    <List Title="List-1">
	<View Name="My Super View 1" URL="View1" DefaultView="TRUE" Hidden="False" ReadOnly ="False" RowLimit="50" Paged="True">
		<!-- STANDARD CAML QUERY VIEW with Query, ViewFields, ... nodes -->
	</View>
	<View Name="View 2" URL="View2" DefaultView="FALSE" Hidden="False" ReadOnly ="False" RowLimit="50" Paged="True">
		<!-- STANDARD CAML QUERY VIEW with Query, ViewFields, ... nodes -->
	</View>
    </List>
 </AddView>
 <UpdateView></UpdateView>
 <DeleteView></DeleteView>
</SPListViews> 
```

### Export all Lists Views from a SPWeb using extractListViews.ps1
Refresh all the fields of a list based on their Site Column definition
```Powershell
extractListViews.ps1 -siteURL "http://spweb_url" [-destinationFolderPath <LocalFolderPathForGeneratedFile>]
```

###  Import all Lists Views to another SPWeb using browseAndParseListViewsXML()
To ensure that these files are uploaded, you need to call browseFilesAndFoldersToUpload().
This function will browse and parse the SPFileUploader.xml and initiate the upload of all referenced files into the right location (library, folder, subfolder, ...) on site siteURL.

```Powershell
Import-Module "pathTo_ModuleFolder\SPHelpers\SPHelpers.psm1"
Import-Module "pathTo_ModuleFolder\SPListViews\SPListViews.psm1"
# Configuration file 'SPListViews.xml'
$listViewXmlFilePath = "pathTo_ConfigFile\SPListViews.xml"
if(Test-Path $listViewXmlFilePath)
{
  $listViewsXML = LoadXMLFile -xmlPath $listViewXmlFilePath
  if($listViewsXML -ne $null -and $listViewsXML.HasChildNodes)
  {
    browseAndParseListViewsXML -siteURL $siteURL -listViewsXML $listViewsXML
  }
  else
  {
    Write-Warning "XML File for <SPListViews> is empty." 
  }
}
else
{
  Write-Warning "XML File for <SPListViews> does not exist."
}
```


