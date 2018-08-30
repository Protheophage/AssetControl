Function Get-AssetInfo
{
	<#
	.SYNOPSIS
	Query SQL for Asset(s)
	
	.DESCRIPTION
	Pull asset entries by name, ID, or department code.
	Note: If names and IDs are both given, this will search for the names.
	
	.PARAMETER AssetName
	Looks for a string from the input.
	.PARAMETER AssetID
	Looks for an Integer from the input.
	.PARAMETER LogPath
	Defaults to desktop if no value is entered with -LogPath
	
	.EXAMPLE
	Get-AssetInfo- AssetName "PC-42"
	Search by name.
	
	.EXAMPLE
	Get-AssetInfo -AssetName "PC-42","PC-69","PC-99" -LogPath "C:\logs\thislog.txt"
	Search for multiple assets by name, and set custom log path.
	
	.EXAMPLE
	Get-AssetInfo -AssetID 1942
	Search by ID.
	
	.EXAMPLE
	Get-AssetInfo -AssetID 1942,1969,1999 -LogPath "C:\logs\thislog.txt"
	Search for multiple assets by ID, and set custom log path.
	
	.EXAMPLE
	Get-AssetInfo -AssetName "IT-%"
	Searh for a group of assets by department code
	
	.EXAMPLE
	Get-AssetInfo -AssetName "IT-%","MK-%","HO-%"
	Search for multiple groups of assets by department code
	#>
	[CmdletBinding()]
	Param
	(
		[parameter(ValueFromPipeline=$True)]
		[String[]]$AssetName,
		[parameter(ValueFromPipeline=$True)]
		[Int[]]$AssetID,
		$LogPath
	)
	BEGIN
	{
		IF(!$LogPath) {$LogPath = "$($env:USERPROFILE)\Desktop\Search_Log.txt"}
		$ChangeLog = @()
		Set-Location SQLSERVER:\SQL\Anu\SQLEXPRESS\Databases\Assets\tables
	}
	PROCESS
	{
		If(!$AssetName)
		{
			Foreach($ID in $AssetID)
			{

			$cmpFnd = Invoke-Sqlcmd "SELECT * FROM [dbo].[AssetList] Where asset_id LIKE '$ID';"
			$ChangeLog = $ChangeLog + ($cmpFnd)
			}
		}
		ELSE
		{
			Foreach($Name in $AssetName)
			{

			$cmpFnd = Invoke-Sqlcmd "SELECT * FROM [dbo].[AssetList] Where asset_name LIKE '$Name';"
			$ChangeLog = $ChangeLog + ($cmpFnd)
			}
		}
	}
	END
	{
		$ChangeLog = [PSObject]$ChangeLog
		
		Set-Location C:\WINDOWS\system32
		
			If (!(Test-Path $LogPath))
		{
			New-Item $LogPath -ItemType File
			$ChangeLog | ForEach-Object {"`r`nDate Added: $($_.date_added)`r`nDate Updated: $($_.date_updated)`r`nAsset name: $($_.asset_name)`r`nAsset ID: $($_.asset_id)`r`nSerial Number: $($_.serial_number)`r`nManufacturer: $($_.manufacturer)`r`nModel: $($_.model)`r`nDescription: $($_.description)`r`nProduct Key: $($_.product_key)`r`nStatus: $($_.status)`r`nPurchase Price: $($_.purch_price)" | Add-Content -Path $LogPath}
			Add-Content -Path $LogPath -Value "`r`nPLEASE REMEMBER TO DELETE THIS FILE WHEN YOU ARE FINISHED WITH IT."
			notepad.exe $LogPath
		}
		Else
		{
			$ChangeLog | ForEach-Object {"`r`nDate Added: $($_.date_added)`r`nDate Updated: $($_.date_updated)`r`nAsset name: $($_.asset_name)`r`nAsset ID: $($_.asset_id)`r`nSerial Number: $($_.serial_number)`r`nManufacturer: $($_.manufacturer)`r`nModel: $($_.model)`r`nDescription: $($_.description)`r`nProduct Key: $($_.product_key)`r`nStatus: $($_.status)`r`nPurchase Price: $($_.purch_price)" | Add-Content -Path $LogPath}
			Add-Content -Path $LogPath -Value "`r`nPLEASE REMEMBER TO DELETE THIS FILE WHEN YOU ARE FINISHED WITH IT."
			notepad.exe $LogPath
		}
		$ChangeLog		
	}
}