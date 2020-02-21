Function Confirm-ITComp
{
	<#
	.SYNOPSIS
	Check if this is being run on an IT computer and quit if not
	#>
	If ($env:computername -NotLike "IT-*")
	{
		##Show popup for user
		$wshell = New-Object -ComObject Wscript.Shell
		$wshell.Popup("Please run this from an IT department computer.",0,$env:computername,0x0)
		Break
	}
}

Function Confirm-inSQL
{
	<#
	.SYNOPSIS
	Check if Asset Name is already in SQL
	
	.PARAMETER CompName

	.EXAMPLE
	Confirm-inSQL -CompName "PC-42"
	Checks if there is already an asset named PC-42. Prompts with pop-up if found.	
	#>
	
	[CmdletBinding()]
    Param
	(
		[parameter(ValueFromPipeline=$True)]
		[String[]]$CompName
	)

	BEGIN
	{
		##Set Location of PS instance to SQL Database
		Set-Location SQLSERVER:<Your_SQL_Server>Databases\Assets\Tables
	}
	
	PROCESS
	{
		ForEach($C in $CompName)
		{
			$cmpFnd = Invoke-Sqlcmd "SELECT * FROM dbo.AssetList Where asset_name COLLATE SQL_Latin1_General_CP1_CI_AS = '$C';"
			
			if($null -ne $cmpFnd)
			{
				##Show popup for user
				$wshell = New-Object -ComObject Wscript.Shell
				$wshell.Popup(($cmpFnd | ForEach-Object { "$($_.asset_name) is already registered as follows:`nAsset ID: $($_.asset_id)`nSerial Number: $($_.serial_number)"}),0,$cmpFnd.asset_name,0x0)
			}
		}
	}
	
	END
	{
		##Return to default PS location
		Set-Location C:\WINDOWS\system32
	}
}

