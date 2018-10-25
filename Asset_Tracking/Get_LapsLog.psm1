Function Get-LapsLog
{
	<#
	.SYNOPSIS
	Query SQL for LAPS
	
	.DESCRIPTION
	Pull LAPS entries for an asset, multiple assets, or a group of assets by department code.
	
	.PARAMETER AssetName
	Looks for a string from the input.
	.PARAMETER AssetID
	Looks for an Integer from the input.
	.PARAMETER LogPath
	Defaults to desktop if no value is entered with -LogPath
	
	.EXAMPLE
	Get-LapsLog -AssetName "PC-42"
	Search by name.
	
	.EXAMPLE
	Get-LapsLog -AssetName "PC-42","PC-69","PC-99" -LogPath "C:\logs\thislog.txt"
	Search for multiple assets by name, and set custom log path.

	.EXAMPLE
	Get-LapsLog -AssetName "IT-%"
	Searh for a group of assets by department code
	
	.EXAMPLE
	Get-LapsLog -AssetName "IT-%","MK-%","HO-%"
	Search for multiple groups of assets by department code
	#>
	
	[CmdletBinding()]
	Param
	(
		[parameter(ValueFromPipeline=$True)]
		[String[]]$AssetName,
		$LogPath
	)
	BEGIN
	{
		IF(!$LogPath) {$LogPath = "$($env:USERPROFILE)\Desktop\Search_Log.txt"}
		$ChangeLog = @()
		Set-Location SQLSERVER:\SQL\PROMETHEUS\DEFAULT\Databases\Assets\Tables
	}
	PROCESS
	{
		Foreach($Name in $AssetName)
			{

			$cmpFnd = Invoke-Sqlcmd "SELECT * FROM [dbo].[Laps_Log] Where asset_name LIKE '$Name';"
			$ChangeLog = $ChangeLog + ($cmpFnd)
			}
	}
	END
	{
		$ChangeLog = [PSObject]$ChangeLog
		
		Set-Location C:\WINDOWS\system32
		
			If (!(Test-Path $LogPath))
		{
			New-Item $LogPath -ItemType File
			$ChangeLog | ForEach-Object {"`r`nDate Logged: $($_.date_logged)`r`nAsset name: $($_.asset_name)`r`nAsset ID: $($_.asset_id)`r`nSerial Number: $($_.serial_number)`r`nLAPS Password: $($_.laps_pw)" | Add-Content -Path $LogPath}
			Add-Content -Path $LogPath -Value "`r`nPLEASE REMEMBER TO DELETE THIS FILE WHEN YOU ARE FINISHED WITH IT."
			notepad.exe $LogPath
		}
		Else
		{
			$ChangeLog | ForEach-Object {"`r`nDate Logged: $($_.date_logged)`r`nAsset name: $($_.asset_name)`r`nAsset ID: $($_.asset_id)`r`nSerial Number: $($_.serial_number)`r`nLAPS Password: $($_.laps_pw)" | Add-Content -Path $LogPath}
			Add-Content -Path $LogPath -Value "`r`nPLEASE REMEMBER TO DELETE THIS FILE WHEN YOU ARE FINISHED WITH IT."
			notepad.exe $LogPath
		}
		$ChangeLog		
	}
}