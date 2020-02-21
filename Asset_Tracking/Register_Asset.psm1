Function Register-Asset
{
    <#
	.SYNOPSIS
	Register a new asset into our system
	
	.DESCRIPTION
	Register a new asset into our system
	
	.PARAMETER AssetName
	
	.PARAMETER PurchaseValue
	
	.PARAMETER Manufacturer
	
	.PARAMETER Model

	.PARAMETER SerialNumber

	.PARAMETER Description

	.PARAMETER ProductKey
	
	.PARAMETER AssetTypeName

	.PARAMETER NotPC

	.PARAMETER IsOnline
	
	.EXAMPLE
	Register-Asset -AssetName "PC-42" -PurchaseValue "123.45" -IsOnline
	Registers a new asset with the name "PC-42", sets the purchase value to $123.45, and states that the PC is online.
	Only the asset name and purchase value are required for PCs. All other info will be gathered automatically.
	
	.EXAMPLE
	Register-Asset -AssetName "iPod-42" -PurchaseValue "123.45" -Manufacturer "Apple" -Model "iPod Touch" -SerialNumber "987654321" -Description "The 42nd iPod" -AssetTypeName "ios" -NotPC
	Registers a non-PC asset.
	All information must be input manually for a non PC asset, and the -NotPC switch MUST be set.
	
	#>
	
    [CmdletBinding()]
    Param
	(
		[parameter(ValueFromPipeline=$True)]
		[parameter(Mandatory=$true)]
		[String[]]$AssetName,
		[parameter(Mandatory=$true)]
		[Single[]]$PurchaseValue,
		[String[]]$Manufacturer,
		[String[]]$Model,
		[String[]]$SerialNumber,
		[String[]]$Description,
		[String[]]$ProductKey,
		[String[]]$AssetTypeName,
        [switch]$NotPC,
		[switch]$IsOnline
    )
    
    BEGIN
    {
		Set-Location SQLSERVER:<Your_SQL_Server>Databases\Assets\Tables
		$i = 0
    }
    PROCESS
    {
		If(!$NotPC)
		{
			Foreach($N in $AssetName)
			{
				$PV = $PurchaseValue[$i]
				If(!$IsOnline)  
				{  
					#Check if PC is Online
					If(!(Test-Connection -Cn $N -BufferSize 16 -Count 1 -ea 0 -quiet))
					{
						Write-Host $N" is not online"
						Continue
					}
					Else
					{
						$cmpFnd = Invoke-Sqlcmd "SELECT * FROM dbo.AssetList Where asset_name COLLATE SQL_Latin1_General_CP1_CI_AS = '$N';"
						if($null -ne $cmpFnd)
						{
							##Show popup for user
							$wshell = New-Object -ComObject Wscript.Shell
							$wshell.Popup(($cmpFnd | ForEach-Object { "$($_.asset_name) is already registered as follows:`nAsset ID: $($_.asset_id)`nSerial Number: $($_.serial_number)"}),0,$cmpFnd.asset_name,0x0)
							Continue
						}
						Else
						{
							##Set variable for asset id
							$AssettID = Invoke-Sqlcmd "SELECT MAX(asset_id)+1 FROM dbo.AssetList Where asset_id IS NOT NULL AND asset_id < 9990000;"
							$AssettID = $AssettID.column1
							##Assign $AssettID to asset_id now in case of new query before this finishes
							Invoke-Sqlcmd "INSERT INTO dbo.AssetList (asset_id) VALUES ('$AssettID')"

							#Gather Info From PC
							[String]$Manufacturer = (Get-WmiObject -ComputerName "$N" Win32_SystemEnclosure).Manufacturer
							[String]$Model = (Get-WmiObject -ComputerName "$N" -Class Win32_ComputerSystem).model
							[String]$SerialNumber = (get-wmiobject -computername "$N" win32_bios).serialnumber
							[String]$Description = (Get-WmiObject -ComputerName "$N" -Class Win32_OperatingSystem).Description
							[String]$Description = $Description -replace '[\W]', ' '
							[String]$ProductKey = (get-wmiObject -computername "$N" -Class SoftwareLicensingService).OA3xOriginalProductKey
							[String]$AssetTypeName = (Get-WmiObject -ComputerName "$N"-Class Win32_OperatingSystem).caption

							##Send Asset ID to AD Attribute "comment", Serial Number to AD Attribute "serialNumber", and Product Key to AD Attribute carLicense
							$CompAD = Get-ADComputer -Identity $N -Properties comment,carLicense,ms-Mcs-AdmPwd
							$CompAD.serialNumber = $SerialNumber
							$CompAD.comment = $AssettID
							$CompAD.carLicense = $ProductKey
							Set-ADComputer -Instance $CompAD

							[String]$LAPSpw = $CompAD."ms-Mcs-AdmPwd"

							##Update remaining info in SQL
							Invoke-Sqlcmd "UPDATE dbo.AssetList SET date_added = GetDate(), date_updated = GetDate(), asset_name = '$N', asset_type_name = '$AssetTypeName', serial_number = '$SerialNumber', manufacturer = '$Manufacturer', model = '$Model', description = '$Description', product_key = '$ProductKey', status = '1', purch_price = '$PV' WHERE asset_id = '$AssettID';"

							Invoke-Sqlcmd "INSERT INTO [dbo].[Laps_Log]
								([date_logged]
								,[asset_name]
								,[asset_id]
								,[serial_number]
								,[laps_pw])
							VALUES
								(GetDate()
								,'$N'
								,'$AssettID'
								,'$SerialNumber'
								,'$LAPSpw')
							GO"

							Set-BiosAsset -Name $N -IsOnline

							New-DymoLabel -AstNum $AssettID -AstModel $Model

							$i = $i + 1
						}
					}
				}
				Else
				{
					$cmpFnd = Invoke-Sqlcmd "SELECT * FROM dbo.AssetList Where asset_name COLLATE SQL_Latin1_General_CP1_CI_AS = '$N';"
					if($null -ne $cmpFnd)
                	{
						##Show popup for user
						$wshell = New-Object -ComObject Wscript.Shell
						$wshell.Popup(($cmpFnd | ForEach-Object { "$($_.asset_name) is already registered as follows:`nAsset ID: $($_.asset_id)`nSerial Number: $($_.serial_number)"}),0,$cmpFnd.asset_name,0x0)
						Continue
					}
					Else
					{
						##Set variable for asset id
						$AssettID = Invoke-Sqlcmd "SELECT MAX(asset_id)+1 FROM dbo.AssetList Where asset_id IS NOT NULL AND asset_id < 9990000;"
						$AssettID = $AssettID.column1
						##Assign $AssettID to asset_id now in case of new query before this finishes
						Invoke-Sqlcmd "INSERT INTO dbo.AssetList (asset_id) VALUES ('$AssettID')"

						#Gather Info From PC
						[String]$Manufacturer = (Get-WmiObject -ComputerName "$N" Win32_SystemEnclosure).Manufacturer
						[String]$Model = (Get-WmiObject -ComputerName "$N" -Class Win32_ComputerSystem).model
						[String]$SerialNumber = (get-wmiobject -computername "$N" win32_bios).serialnumber
						[String]$Description = (Get-WmiObject -ComputerName "$N" -Class Win32_OperatingSystem).Description
						[String]$Description = $Description -replace '[\W]', ' '
						[String]$ProductKey = (get-wmiObject -computername "$N" -Class SoftwareLicensingService).OA3xOriginalProductKey
						[String]$AssetTypeName = (Get-WmiObject -ComputerName "$N"-Class Win32_OperatingSystem).caption

						##Send Asset ID to AD Attribute "comment", Serial Number to AD Attribute "serialNumber", and Product Key to AD Attribute carLicense
						$CompAD = Get-ADComputer -Identity $N -Properties comment,carLicense,ms-Mcs-AdmPwd
						$CompAD.serialNumber = $SerialNumber
						$CompAD.comment = $AssettID
						$CompAD.carLicense = $ProductKey
						Set-ADComputer -Instance $CompAD

						[String]$LAPSpw = $CompAD."ms-Mcs-AdmPwd"

						##Update remaining info in SQL
						Invoke-Sqlcmd "UPDATE dbo.AssetList SET date_added = GetDate(), date_updated = GetDate(), asset_name = '$N', asset_type_name = '$AssetTypeName', serial_number = '$SerialNumber', manufacturer = '$Manufacturer', model = '$Model', description = '$Description', product_key = '$ProductKey', status = '1', purch_price = '$PV' WHERE asset_id = '$AssettID';"

						Invoke-Sqlcmd "INSERT INTO [dbo].[Laps_Log]
							([date_logged]
							,[asset_name]
							,[asset_id]
							,[serial_number]
							,[laps_pw])
						VALUES
							(GetDate()
							,'$N'
							,'$AssettID'
							,'$SerialNumber'
							,'$LAPSpw')
						GO"

						Set-BiosAsset -Name $N -IsOnline

						New-DymoLabel -AstNum $AssettID -AstModel $Model

						$i = $i + 1
					}
				}
			}
		}
		Else
		{
			##Set variable for asset id
			$AssettID = Invoke-Sqlcmd "SELECT MAX(asset_id)+1 FROM dbo.AssetList Where asset_id IS NOT NULL AND asset_id < 9990000;"
			$AssettID = $AssettID.column1
			##Assign $AssettID to asset_id now in case of new query before this finishes
			Invoke-Sqlcmd "INSERT INTO dbo.AssetList (asset_id) VALUES ('$AssettID')"

			##Update remaining info in SQL
			Invoke-Sqlcmd "UPDATE dbo.AssetList SET date_added = GetDate(), date_updated = GetDate(), asset_name = '$N', asset_type_name = '$AssetTypeName', serial_number = '$SerialNumber', manufacturer = '$Manufacturer', model = '$Model', description = '$Description', product_key = '$ProductKey', status = '1', purch_price = '$PV' WHERE asset_id = '$AssettID';"

			New-DymoLabel -AstNum $AssettID -AstModel $Model

			$i = $i + 1
		}
    }
    END
    {
        Set-Location C:\Windows\system32
    }
}