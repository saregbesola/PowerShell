function Get-ItemInfo($filePath)
{
$itemProperties = Get-ChildItem -Path $filePath | Select-Object Name,FullName,CreationTime,LastWriteTime
$owner = (Get-Acl -Path $filePath).Owner
$result = [PSCustomObject]@{
Name = $itemProperties.Name
FullName = $itemProperties.FullName
CreationTime = $itemProperties.CreationTime
ModifiedDate = $itemProperties.LastWriteTime
Owner = $owner
}
return $result
}
Get-ItemInfo "c:\Krishna\PowerShell Scripts\Get-fileInfo.ps1"