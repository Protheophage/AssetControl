Function UnRegister-Asset
 {
	<#
	.SYNOPSIS
	Retire an asset
	
	.DESCRIPTION
	Finds the asset ($AssetID), sets status to 0 and adds entry to Retired log
	
	.PARAMETER AssetID
	This is Asset ID assigned when registered
	
	.EXAMPLE
	UnRegister-Asset('1892')
	Retire one asset.
	
	.EXAMPLE
	UnRegister-Asset('1892','1968','1969')
    Retire multiple assets.
    
    .EXAMPLE
	Retire-Asset('1892')
	Retire one asset.
	
	.EXAMPLE
	Retire-Asset('1892','1968','1969')
    Retire multiple assets.
	#>
	
	[CmdletBinding()]
	Param
	(
		[parameter(ValueFromPipeline=$True)]
		[Int[]]$AssetID
	)
	
	BEGIN{
		Set-Location SQLSERVER:\SQL\Anu\SQLEXPRESS\Databases\Assets\tables
	}
	PROCESS{
		Foreach($ID in $AssetID)
		{
            $cmpFnd = Invoke-Sqlcmd "SELECT * FROM [dbo].[AssetList] Where asset_id LIKE '$ID';"
            $rdReg = $cmpFnd.date_added
            [Int]$raID = $cmpFnd.asset_id
            $ratName = $cmpFnd.asset_type_name
            $rsNum = $cmpFnd.serial_number
            $rMan = $cmpFnd.manufacturer
            $rMod = $cmpFnd.model
            $[Double]rpPrice = $cmpFnd.purch_price
            
            $tDate = (GET-DATE)
            $span = NEW-TIMESPAN -Start $cmpFnd.date_added -End $tDate
            [Int]$age = $span.days/365
            IF($age -lt 4 -and $age -gt 0)
                {[Double]$rVal = $cmpFnd.purch_price/$age}
            ELSE
                {[Double]$rVal = 0.00}

            Invoke-Sqlcmd "
			BEGIN TRANSACTION
			UPDATE [dbo].[AssetList]
			SET date_updated = GetDate()
			WHERE asset_id = '$ID';
			UPDATE [dbo].[AssetList]
			SET asset_name = 'Retired'
            WHERE asset_id = '$ID';
            UPDATE [dbo].[AssetList]
            SET status = '0'
            WHERE asset_id = '$ID';
			COMMIT TRANSACTION;
            "
            Invoke-Sqlcmd "
			INSERT INTO [dbo].[Retired]
           ([date_registered]
           ,[date_retired]
           ,[asset_id]
           ,[asset_type_name]
           ,[serial_number]
           ,[manufacturer]
           ,[model]
           ,[purch_price]
           ,[retired_value])
        VALUES
           ('$rdReg'
           ,GetDate()
           ,'$raID'
           ,'$ratName'
           ,'$rsNum'
           ,'$rMan'
           ,'$rMod'
           ,'$rpPrice'
           ,'$rVal')
        GO
            "
		}	
	}
	END{
		Set-Location C:\WINDOWS\system32
	}
 }

 New-Alias -Name Retire-Asset -Value UnRegister-Asset

 Export-ModuleMember -Alias * -Function *