Function Show-NextAssetID
{
	<#
	.SYNOPSIS
	Retrieve the next available Asset ID from SQL
	#>
	
    ##Set Location of PS instance to SQL Database
	Set-Location SQLSERVER:<Your_SQL_Server>Databases\Assets\Tables
	
	##Set variable for asset id
	$Assetted = Invoke-Sqlcmd "SELECT MAX(asset_id)+1 FROM dbo.AssetList Where asset_id IS NOT NULL AND asset_id < 9990000;"
	$Assetted = $Assetted.column1

	##Return to default PS location
	Set-Location C:\WINDOWS\system32

    ##Display Asset ID to user
	$wshell = New-Object -ComObject Wscript.Shell
	$wshell.Popup("The next available Asset ID is $($Assetted).",0,"Progress Update",0x0)
}