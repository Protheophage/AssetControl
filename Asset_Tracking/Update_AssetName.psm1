Function Update-AssetName
{
	<#
	.SYNOPSIS
	Update the asset name in SQL
	
	.DESCRIPTION
	Finds the asset by ($Name) or ($Serial) and updates with the asset ($NewName)
	
	.PARAMETER Name
	This is the current name of the asset.
	.PARAMETER NewName
	This is the desired new asset name
	.PARAMETER Serial
	This is the serial number of the asset
	
	.EXAMPLE
	Update-AssetName('PC-42')('PC-24')
	Update one asset.
	
	.EXAMPLE
	Update-AssetName('PC-42','PC-69','PC-99')('PC-24','PC-96','PC-66')
	Update multiple assets seaching by current Name.
	

	.EXAMPLE
	Update-AssetName -Serial '1JXNRZ2','1JXHBZ2' -NewName 'PC-42','PC-13'
	Update multiple assets serching by Serial Number
	#>
	
	[CmdletBinding()]
	Param
	(
		[parameter(ValueFromPipeline=$True)]
		[String[]]$Name,
		[parameter(ValueFromPipeline=$True)]
		[String[]]$NewName,
		[String]$Serial,
		[Bool]$IsOnline
	)
	
	BEGIN
	{
		Set-Location SQLSERVER:\SQL\PROMETHEUS\DEFAULT\Databases\Assets\Tables
		$i = 0
	}
	PROCESS
	{
		If(!$Serial)
		{
			Foreach($N in $Name)
			{
				$NN = $NewName[$i]
				If(!$IsOnline)
				{
					[String]$CmpDsc = (Get-WmiObject -ComputerName "$NN" -Class Win32_OperatingSystem).Description
					Invoke-Sqlcmd "
					BEGIN TRANSACTION
					UPDATE [dbo].[AssetList]
					SET description = '$CmpDsc'
					WHERE asset_name COLLATE SQL_Latin1_General_CP1_CI_AS = '$N';
					COMMIT TRANSACTION;
					"
				}
				ElseIf(Test-Connection -Cn $NN -BufferSize 16 -Count 1 -ea 0 -quiet)
				{
					[String]$CmpDsc = (Get-WmiObject -ComputerName "$NN" -Class Win32_OperatingSystem).Description
					Invoke-Sqlcmd "
					BEGIN TRANSACTION
					UPDATE [dbo].[AssetList]
					SET description = '$CmpDsc'
					WHERE asset_name COLLATE SQL_Latin1_General_CP1_CI_AS = '$N';
					COMMIT TRANSACTION;
					"
				}
				Invoke-Sqlcmd "
				BEGIN TRANSACTION
				UPDATE [dbo].[AssetList]
				SET date_updated = GetDate()
				WHERE asset_name COLLATE SQL_Latin1_General_CP1_CI_AS = '$N';
				UPDATE [dbo].[AssetList]
				SET asset_name = '$NN'
				WHERE asset_name COLLATE SQL_Latin1_General_CP1_CI_AS = '$N';
				COMMIT TRANSACTION;
				"
				$i = $i + 1
			}
		}
		Else
		{
			Foreach($S in $Serial)
			{
				$NN = $NewName[$i]
				If(!$IsOnline)
				{
					[String]$CmpDsc = (Get-WmiObject -ComputerName "$NN" -Class Win32_OperatingSystem).Description
					Invoke-Sqlcmd "
					BEGIN TRANSACTION
					UPDATE [dbo].[AssetList]
					SET description = '$CmpDsc'
					WHERE serial_number COLLATE SQL_Latin1_General_CP1_CI_AS = '$S';
					COMMIT TRANSACTION;
					"
				}
				ElseIf(Test-Connection -Cn $NN -BufferSize 16 -Count 1 -ea 0 -quiet)
				{
					[String]$CmpDsc = (Get-WmiObject -ComputerName "$NN" -Class Win32_OperatingSystem).Description
					Invoke-Sqlcmd "
					BEGIN TRANSACTION
					UPDATE [dbo].[AssetList]
					SET description = '$CmpDsc'
					WHERE serial_number COLLATE SQL_Latin1_General_CP1_CI_AS = '$S';
					COMMIT TRANSACTION;
					"
				}
				Invoke-Sqlcmd "
				BEGIN TRANSACTION
				UPDATE [dbo].[AssetList]
				SET date_updated = GetDate()
				WHERE serial_number COLLATE SQL_Latin1_General_CP1_CI_AS = '$S';
				UPDATE [dbo].[AssetList]
				SET asset_name = '$NN'
				WHERE serial_number COLLATE SQL_Latin1_General_CP1_CI_AS = '$S';
				COMMIT TRANSACTION;
				"
				$i = $i + 1
			}
		}
	}
	END
	{
		Set-Location C:\WINDOWS\system32
	}
}