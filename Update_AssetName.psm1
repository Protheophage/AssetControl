Function Update-AssetName
 {
	<#
	.SYNOPSIS
	Update the asset name in SQL
	
	.DESCRIPTION
	Finds the asset ($Name) and updates with the asset ($NewName)
	
	.PARAMETER Name
	This is the current name of the asset.
	.PARAMETER NewID
	This is the desired new asset name
	
	.EXAMPLE
	Update-AssetName('PC-42')('PC-24')
	Update one asset.
	
	.EXAMPLE
	Update-AssetName('PC-42','PC-69','PC-99')('PC-24','PC-96','PC-66')
	Update multiple assets.
	#>
	
	[CmdletBinding()]
	Param
	(
		[parameter(ValueFromPipeline=$True)]
		[String[]]$Name,
		[parameter(ValueFromPipeline=$True)]
		[String[]]$NewName
	)
	
	BEGIN{
		Set-Location SQLSERVER:\SQL\Anu\SQLEXPRESS\Databases\Assets\tables
		$i = 0
	}
	PROCESS{
		Foreach($N in $Name)
		{
			$NN = $NewName[$i]
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
	END{
		Set-Location C:\WINDOWS\system32
		$NewName | Get-AssetInfo
	}
 }