# SPPS
Powershell scripts to interact with SharePoint

## Powershell Scripts

### Get all Workflows Instance By Status usingGetWorkflowsInstancesBystatus.ps1 
Get all Workflows Instances from a specific status and generate dedicated CSV files. This script is a fork of [this one](https://community.nintex.com/community/build-your-own/blog/2015/04/09/finding-all-of-the-workflows-in-your-farm-using-powershell)
```Powershell
GetWorkflowsInstancesBystatus.ps1 -status Running
```

##  Refresh SPFields on SPList using refreshSiteColumsOnList.ps1
Refresh all the fields of a list based on their Site Column definition
```Powershell
refreshSiteColumsOnList.ps1 -siteURL "http://spweb_url" -listName "name of the list"
```


## Powershell Modules

###  SPFileUploader
See [SPFileUploader Wiki Page](https://github.com/OhNotreDame/SPPS/wiki/SPFileUploader)

### SPListViews
See [SPListViews Wiki Page](https://github.com/OhNotreDame/SPPS/wiki/SPListViews)



