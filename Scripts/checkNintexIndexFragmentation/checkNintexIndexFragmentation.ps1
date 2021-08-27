Add-PSSnapin Microsoft.SharePoint.PowerShell

[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint")
[void][System.Reflection.Assembly]::LoadWithPartialName("Nintex.Workflow")
[void][System.Reflection.Assembly]::LoadWithPartialName("Nintex.Workflow.Administration")

$cmd = New-Object -TypeName System.Data.SqlClient.SqlCommand

$cmd.CommandType = [System.Data.CommandType]::Text

$cmd.CommandText = "SELECT OBJECT_NAME(i.object_id) AS TableName ,i.name AS TableIndexName ,phystat.avg_fragmentation_in_percent FROM sys.dm_db_index_physical_stats(DB_ID(), NULL, NULL, NULL, 'DETAILED') phystat inner JOIN sys.indexes i ON i.object_id = phystat.object_id AND i.index_id = phystat.index_id WHERE phystat.avg_fragmentation_in_percent > 10"


$reader = [Nintex.Workflow.Administration.ConfigurationDatabase]::GetConfigurationDatabase().ExecuteReader($cmd)

$indexes = @()

while($reader.Read())
{
    $row = New-Object System.Object

    $row | Add-Member -MemberType NoteProperty -Name "TableName" -Value $reader["TableName"]
    $row | Add-Member -MemberType NoteProperty -Name "TableIndexName" -Value $reader["TableIndexName"]
    $row | Add-Member -MemberType NoteProperty -Name "avg_fragmentation_in_percent" -Value $reader["avg_fragmentation_in_percent"]


$indexes += $row
}

$indexes