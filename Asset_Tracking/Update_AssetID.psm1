Function Update-AssetID
 {
	<#
	.SYNOPSIS
	Update the asset ID in SQL
	
	.DESCRIPTION
	Finds the asset ($Name) and updates the asset ID ($NewID)
	
	.PARAMETER Name
	This is the current name of the asset.
	.PARAMETER NewID
	This is the desired new asset ID
	
	.EXAMPLE
	Update-AssetID("PC-42")(1234)
	Update one asset.
	
	.EXAMPLE
	Update-AssetID("PC-42","PC-69","PC-99")(1234,4321,5678)
	Update multiple assets.
	#>
	
	[CmdletBinding()]
	Param
	(
		[parameter(ValueFromPipeline=$True)]
		[String[]]$Name,
		[parameter(ValueFromPipeline=$True)]
		[INT[]]$NewID
	)
	
	BEGIN{
		Set-Location SQLSERVER:<Your_SQL_Server>Databases\Assets\Tables
		$i = 0
	}
	PROCESS{
		Foreach($N in $Name)
		{
			$ID = $NewID[$i]
			Invoke-Sqlcmd "
			BEGIN TRANSACTION
			UPDATE [dbo].[AssetList]
			SET date_updated = GetDate()
			WHERE asset_name COLLATE SQL_Latin1_General_CP1_CI_AS = '$N';
			UPDATE [dbo].[AssetList]
			SET asset_id = '$ID'
			WHERE asset_name COLLATE SQL_Latin1_General_CP1_CI_AS = '$N';
			COMMIT TRANSACTION;
			"
			$i = $i + 1
		}	
	}
	END{
		Set-Location C:\WINDOWS\system32
	}
 }