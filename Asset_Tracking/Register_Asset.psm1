Function Register-Asset
{
Function Is-ITComp
{
	<#
	.SYNOPSIS
	Check if this is being run on an IT computer and quit if not
	#>
	If ($env:computername -like "IT-*")
	{
        Get-AssetName
	}
	Else
	{
		echo "Please run this from an IT department computer."
		Break
	}
}

 
Function Get-AssetName
{
	<#
	.SYNOPSIS
	Get Asset name from user and assign to $comp
	Get Purchase Value from user and assign to $purval
	#>
	$uInp = Invoke-2InpBox -formTitle "Asset Information" -formPrompt "Please enter the asset name and purchase value." -b1Text "Asset Name:" -b2Text "Value:"
	$comp = $uInp.Box1
	[single]$purval = $uInp.Box2
    Is-inSQL
}

<#
#>
 
Function Is-inSQL
{
	<#
	.SYNOPSIS
	Check if Asset Name is already in SQL
	#>

	##Set Location of PS instance to SQL Database
	Set-Location SQLSERVER:\SQL\PROMETHEUS\DEFAULT\Databases\Assets\Tables
	$cmpFnd = Invoke-Sqlcmd "SELECT * FROM dbo.AssetList Where asset_name COLLATE SQL_Latin1_General_CP1_CI_AS = '$comp';"

	##Return to default PS location
	Set-Location C:\WINDOWS\system32
	
	if($cmpFnd -ne $null)
	{
		##Show popup for user
		$wshell = New-Object -ComObject Wscript.Shell
		$wshell.Popup(($cmpFnd | ForEach-Object { "$($_.asset_name) is already registered as follows:`nAsset ID: $($_.asset_id)`nSerial Number: $($_.serial_number)"}),0,$cmpFnd.asset_name,0x0)
		Break
	}
	ELSE
	{
		Get-AssetType
	}		

}

<#
#>
 
Function Get-AssetType
{

	[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
	[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
	$objFormA = New-Object System.Windows.Forms.Form 
	$objFormA.Text = "Asset Type:"
	$objFormA.Size = New-Object System.Drawing.Size(200,140) 
	$objFormA.StartPosition = "CenterScreen"
	
	$objFormA.KeyPreview = $True
	$objFormA.Add_KeyDown({
		if ($_.KeyCode -eq "Enter" -or $_.KeyCode -eq "Escape"){
			$objFormA.Close()
		}
	})
	
	$PCbutton = New-Object System.Windows.Forms.Button
	$PCbutton.Location = New-Object System.Drawing.Size(20,60)
	$PCbutton.Size = New-Object System.Drawing.Size(75,23)
	$PCbutton.Text = "PC"
	$PCbutton.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$objFormA.AcceptButton = $PCbutton
    $objFormA.Controls.Add($PCbutton)
    
	$OtherButton = New-Object System.Windows.Forms.Button
	$OtherButton.Location = New-Object System.Drawing.Size(90,60)
	$OtherButton.Size = New-Object System.Drawing.Size(75,23)
	$OtherButton.Text = "Other"
	$OtherButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
	$objFormA.CancelButton = $OtherButton
    $objFormA.Controls.Add($OtherButton)
	
	$objLabelA = New-Object System.Windows.Forms.Label
	$objLabelA.Location = New-Object System.Drawing.Size(20,20) 
	$objLabelA.Size = New-Object System.Drawing.Size(180,100)
	$objLabelA.Text = "Is the asset a Windows PC,`nor another device type?"
	$objFormA.Controls.Add($objLabelA)
	
		$objFormA.Topmost = $True
	
	$objFormA.Add_Shown({$objFormA.Activate()})
	
	$WhtKnd = $objFormA.ShowDialog()

	If ($WhtKnd -eq "OK")
	{
		Is-Online
	}
	Else
	{
		Get-AssetID
	}
}

<#
#>
 
Function Is-Online
{
	<#
	.SYNOPSIS
	Check if PC is online
	#>
	
	##Return to default PS location
	Set-Location C:\WINDOWS\system32

	if(!(Test-Connection -Cn $comp -BufferSize 16 -Count 1 -ea 0 -quiet))
	{

		[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
		[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
		$objFormA = New-Object System.Windows.Forms.Form 
		$objFormA.Text = "Warning"
		$objFormA.Size = New-Object System.Drawing.Size(200,140) 
		$objFormA.StartPosition = "CenterScreen"
		
		$objFormA.KeyPreview = $True
		$objFormA.Add_KeyDown({
			if ($_.KeyCode -eq "Enter" -or $_.KeyCode -eq "Escape"){
				$objFormA.Close()
			}
		})
		
		$PCbutton = New-Object System.Windows.Forms.Button
		$PCbutton.Location = New-Object System.Drawing.Size(20,60)
		$PCbutton.Size = New-Object System.Drawing.Size(75,23)
		$PCbutton.Text = "Continue"
		$PCbutton.DialogResult = [System.Windows.Forms.DialogResult]::OK
		$objFormA.AcceptButton = $PCbutton
		$objFormA.Controls.Add($PCbutton)
		
		$OtherButton = New-Object System.Windows.Forms.Button
		$OtherButton.Location = New-Object System.Drawing.Size(90,60)
		$OtherButton.Size = New-Object System.Drawing.Size(75,23)
		$OtherButton.Text = "Exit"
		$OtherButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
		$objFormA.CancelButton = $OtherButton
		$objFormA.Controls.Add($OtherButton)
		
		$objLabelA = New-Object System.Windows.Forms.Label
		$objLabelA.Location = New-Object System.Drawing.Size(20,20) 
		$objLabelA.Size = New-Object System.Drawing.Size(180,100)
		$objLabelA.Text = "The PC appears to be offline.`nWould you like to continue?"
		$objFormA.Controls.Add($objLabelA)
		
			$objFormA.Topmost = $True
		
		$objFormA.Add_Shown({$objFormA.Activate()})
		
		$WhtKnd = $objFormA.ShowDialog()
	
		If ($WhtKnd -eq "OK")
		{
			Get-AssetID
		}
		Else
		{
			BREAK
		}
	}
	ELSE
	{
		Get-PcInfo
	}
}	

<#
#>
 
Function Get-PcInfo
{
	<#
	.SYNOPSIS
	Retrieve an Asset ID from SQL, reserve spot and, get info from PC
	#>
	
    ##Tell user what is happening
	$wshell = New-Object -ComObject Wscript.Shell
	$wshell.Popup("$($comp) pre-checks completed.`nGathering PC information.",3,"Progress Update",0x0)

	##Set Location of PS instance to SQL Database
	Set-Location SQLSERVER:\SQL\PROMETHEUS\DEFAULT\Databases\Assets\Tables
	
	##Set variable for asset id
	$Assetted = Invoke-Sqlcmd "SELECT MAX(asset_id)+1 FROM dbo.AssetList Where asset_id IS NOT NULL AND asset_id < 9990000;"
	$Assetted = $Assetted.column1

	##Assign $Assetted to asset_id now in case of new query before this finishes
	Invoke-Sqlcmd "INSERT INTO dbo.AssetList (asset_id) VALUES ('$Assetted')"

	##Return to default PS location
	Set-Location C:\WINDOWS\system32

	##Store asset ID on PC
    ##Tell user what is happening
	$wshell = New-Object -ComObject Wscript.Shell
	$wshell.Popup("Sending ID to $($comp).",3,"Progress Update",0x0)

	##Create AssetID.txt
	New-Item -ItemType directory -path "\\$comp\C$\Windows\Help\AssetID"
	New-Item "\\$comp\C$\Windows\Help\AssetID\AssetID.txt" -ItemType file
	set-content -path "\\$comp\C$\Windows\Help\AssetID\AssetID.txt" -Value $Assetted

	##Get computer description and assign to CmpDescription
	$CmpDescription = Get-WmiObject -ComputerName "$comp" -Class Win32_OperatingSystem | Select Description
	
	##Get Serial number and assign to srlnmbr
	$srlnmbr = get-wmiobject -computername "$comp" win32_bios serialnumber

	##Get Windows Product Key, and assign to $ProdKey
	$GetPK = get-wmiObject -computername "$comp" -query 'select * from SoftwareLicensingService'
	$ProdKey = $GetPK.OA3xOriginalProductKey
	Send-ADinfo
}

<#
#>
 
Function Send-ADinfo
{
	<#
	.SYNOPSIS
	Update Active Directory asset information
	#>
	
	##Return to default PS location
	Set-Location C:\WINDOWS\system32

    ##Tell user what is happening
	$wshell = New-Object -ComObject Wscript.Shell
	$wshell.Popup("$($comp) information gathered.`nUpdating AD and SQL.",3,"Progress Update",0x0)	

	##Send Asset ID to AD Attribute "comment", Serial Number to AD Attribute "serialNumber", and Product Key to AD Attribute carLicense
	$CompAD = Get-ADComputer -Identity $comp -Properties comment,carLicense,ms-Mcs-AdmPwd
	$CompAD.serialNumber = $srlnmbr.SerialNumber
	$CompAD.comment = $Assetted
	$CompAD.carLicense = $ProdKey
	Set-ADComputer -Instance $CompAD

	##Assign Variables to send to SQL
	$ComInfo = get-wmiobject -computername "$comp" win32_computersystem
	$OSInfo = Get-WmiObject -ComputerName "$comp" -Class Win32_OperatingSystem
	$SQName = $CompAD.Name
	$SQTypeName = $OSInfo.caption
	$SQSerial = $srlnmbr.SerialNumber
	$SQManufacturer = $ComInfo.Manufacturer
	$SQModel = $ComInfo.Model
	$SQProdKey = $ProdKey
	$SQLAPS = $CompAD."ms-Mcs-AdmPwd"
	
	Send-SQLAPS
	Send-SQLinfo
}

<#
#>

 Function Send-SQLaps
 {
	<#
	.SYNOPSIS
	Create new entry in SQL LAPS Log
	#>
	
	##Set Location of PS instance to SQL Database
	Set-Location SQLSERVER:\SQL\PROMETHEUS\DEFAULT\Databases\Assets\Tables

	##Update remaining info in SQL
	Invoke-Sqlcmd "INSERT INTO [dbo].[Laps_Log]
           ([date_logged]
           ,[asset_name]
           ,[asset_id]
           ,[serial_number]
           ,[laps_pw])
		VALUES
           (GetDate()
           ,'$SQName'
           ,$($Assetted)
           ,'$SQSerial'
           ,'$SQLAPS')
		GO"
 }
  
<#
#>
 
Function Send-SQLinfo
{
	<#
	.SYNOPSIS
	Update SQL with asset information
	#>
	
	##Set Location of PS instance to SQL Database
	Set-Location SQLSERVER:\SQL\PROMETHEUS\DEFAULT\Databases\Assets\Tables

	##Update remaining info in SQL
	Invoke-Sqlcmd "UPDATE dbo.AssetList SET date_added = GetDate(), date_updated = GetDate(), asset_name = '$SQName', asset_type_name = '$SQTypeName', serial_number = '$SQSerial', manufacturer = '$SQManufacturer', model = '$SQModel', description = '$($CmpDescription).description', product_key = '$SQProdKey', status = '1', purch_price = $($purval) WHERE asset_id = '$Assetted';"
	
	##Show popup for user
	$AsstLst = Invoke-Sqlcmd "SELECT * FROM dbo.AssetList WHERE asset_id = '$Assetted';"
	$wshell = New-Object -ComObject Wscript.Shell
	$wshell.Popup(($AsstLst | ForEach-Object { "Asset ID: $($_.asset_id)`nSerial Number: $($_.serial_number)"}),0,$AsstLst.asset_name,0x0)
	
	Out-Text
}
  
<#
#>
 

Function Out-Text
{
	<#
	.SYNOPSIS
	Create text doc on users desktop with registered asset information.
	#>
	Param
	(
		$Path
	)
	$Path = "$($env:USERPROFILE)\Desktop\Registered_Asset.txt"
	If (!(Test-Path $Path))
	{
		New-Item $Path -ItemType File
		$AsstLst | ForEach-Object {"`r`nDate Added: $($_.date_added)`r`nDate Updated: $($_.date_updated)`r`nAsset name: $($_.asset_name)`r`nAsset ID: $($_.asset_id)`r`nSerial Number: $($_.serial_number)`r`nManufacturer: $($_.manufacturer)`r`nModel: $($_.model)`r`nDescription: $($_.description)`r`nProduct Key: $($_.product_key)`r`nStatus: $($_.status)`r`nPurchase Price: $($_.purch_price)" | Add-Content -Path $Path}
		Add-Content -Path $Path -Value "`r`nPLEASE REMEMBER TO DELETE THIS FILE WHEN YOU ARE FINISHED WITH IT."
		notepad.exe $Path
	}
	Else
	{
		$AsstLst | ForEach-Object {"`r`nDate Added: $($_.date_added)`r`nDate Updated: $($_.date_updated)`r`nAsset name: $($_.asset_name)`r`nAsset ID: $($_.asset_id)`r`nSerial Number: $($_.serial_number)`r`nManufacturer: $($_.manufacturer)`r`nModel: $($_.model)`r`nDescription: $($_.description)`r`nProduct Key: $($_.product_key)`r`nStatus: $($_.status)`r`nPurchase Price: $($_.purch_price)" | Add-Content -Path $Path}
		Add-Content -Path $Path -Value "`r`nPLEASE REMEMBER TO DELETE THIS FILE WHEN YOU ARE FINISHED WITH IT."
		notepad.exe $Path
	}
	$AsstLst | ForEach-Object {$assetnumber = $($_.asset_id);$assetmodel = $($_.model);New-DymoLabel -AstNum $assetnumber -AstModel $assetmodel}
}

<#
#>
 
Function Get-ManualRegistration
{
	[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
	[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
	
	$objForm = New-Object System.Windows.Forms.Form 
	$objForm.Text = "Asset Manual Registration Form"
	$objForm.Size = New-Object System.Drawing.Size(400,300) 
	$objForm.StartPosition = "CenterScreen"
	
	$objForm.KeyPreview = $True
	$objForm.Add_KeyDown({
		if ($_.KeyCode -eq "Enter"){
			$objForm.Close()
		}
		ELSEIF ($_.KeyCode -eq "Escape"){
		BREAK
		}
	})
	
	$OKButton = New-Object System.Windows.Forms.Button
	$OKButton.Location = New-Object System.Drawing.Size(105,220)
	$OKButton.Size = New-Object System.Drawing.Size(75,23)
	$OKButton.Text = "OK"
	$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$objForm.AcceptButton = $OKButton
	$objForm.Controls.Add($OKButton)
	
	$CancelButton = New-Object System.Windows.Forms.Button
	$CancelButton.Location = New-Object System.Drawing.Size(180,220)
	$CancelButton.Size = New-Object System.Drawing.Size(75,23)
	$CancelButton.Text = "Cancel"
	$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
	$objForm.CancelButton = $CancelButton
	$objForm.Controls.Add($CancelButton)
	
	$objLabel = New-Object System.Windows.Forms.Label
	$objLabel.Location = New-Object System.Drawing.Size(20,20) 
	$objLabel.Size = New-Object System.Drawing.Size(380,20) 
	$objLabel.Text = "Please fill in the fields below as appropriate for the asset type:"
	$objForm.Controls.Add($objLabel)
	
	$objAname = New-Object System.Windows.Forms.Label
	$objAname.Location = New-Object System.Drawing.Size(10,40) 
	$objAname.Size = New-Object System.Drawing.Size(80,20) 
	$objAname.Text = "Asset Name:"
	$objForm.Controls.Add($objAname)

	$objTxtAname = New-Object System.Windows.Forms.Label
	$objTxtAname.Location = New-Object System.Drawing.Size(90,40) 
	$objTxtAname.Size = New-Object System.Drawing.Size(280,15) 
	$objTxtAname.Text = $comp
	$objForm.Controls.Add($objTxtAname)	
	
	$objAsID = New-Object System.Windows.Forms.Label
	$objAsID.Location = New-Object System.Drawing.Size(10,60) 
	$objAsID.Size = New-Object System.Drawing.Size(80,20) 
	$objAsID.Text = "Asset ID:"
	$objForm.Controls.Add($objAsID) 
	
	$objTxtAsID = New-Object System.Windows.Forms.Label
	$objTxtAsID.Location = New-Object System.Drawing.Size(90,62) 
	$objTxtAsID.Size = New-Object System.Drawing.Size(280,15) 
	$objTxtAsID.Text = $Assetted
	$objForm.Controls.Add($objTxtAsID) 
	
	$objAtype = New-Object System.Windows.Forms.Label
	$objAtype.Location = New-Object System.Drawing.Size(10,80) 
	$objAtype.Size = New-Object System.Drawing.Size(80,20) 
	$objAtype.Text = "Asset Type:"
	$objForm.Controls.Add($objAtype)
	
	$objTxtAtype = New-Object System.Windows.Forms.TextBox 
	$objTxtAtype.Location = New-Object System.Drawing.Size(90,80) 
	$objTxtAtype.Size = New-Object System.Drawing.Size(280,20) 
	$objForm.Controls.Add($objTxtAtype) 
	
	$objSnum = New-Object System.Windows.Forms.Label
	$objSnum.Location = New-Object System.Drawing.Size(10,100) 
	$objSnum.Size = New-Object System.Drawing.Size(80,20) 
	$objSnum.Text = "Serial Number:"
	$objForm.Controls.Add($objSnum)
	
	$objTxtSnum = New-Object System.Windows.Forms.TextBox 
	$objTxtSnum.Location = New-Object System.Drawing.Size(90,100) 
	$objTxtSnum.Size = New-Object System.Drawing.Size(280,20) 
	$objForm.Controls.Add($objTxtSnum) 
	
	$objManu = New-Object System.Windows.Forms.Label
	$objManu.Location = New-Object System.Drawing.Size(10,120) 
	$objManu.Size = New-Object System.Drawing.Size(80,20) 
	$objManu.Text = "Manufacturer:"
	$objForm.Controls.Add($objManu)
	
	$objTxtManu = New-Object System.Windows.Forms.TextBox 
	$objTxtManu.Location = New-Object System.Drawing.Size(90,120) 
	$objTxtManu.Size = New-Object System.Drawing.Size(280,20) 
	$objForm.Controls.Add($objTxtManu) 
	
	$objModel = New-Object System.Windows.Forms.Label
	$objModel.Location = New-Object System.Drawing.Size(10,140) 
	$objModel.Size = New-Object System.Drawing.Size(80,20) 
	$objModel.Text = "Asset Model:"
	$objForm.Controls.Add($objModel)
	
	$objTxtModel = New-Object System.Windows.Forms.TextBox 
	$objTxtModel.Location = New-Object System.Drawing.Size(90,140) 
	$objTxtModel.Size = New-Object System.Drawing.Size(280,20) 
	$objForm.Controls.Add($objTxtModel) 
	
	$objDesc = New-Object System.Windows.Forms.Label
	$objDesc.Location = New-Object System.Drawing.Size(10,160) 
	$objDesc.Size = New-Object System.Drawing.Size(80,20) 
	$objDesc.Text = "Description:"
	$objForm.Controls.Add($objDesc)
	
	$objTxtDesc = New-Object System.Windows.Forms.TextBox 
	$objTxtDesc.Location = New-Object System.Drawing.Size(90,160) 
	$objTxtDesc.Size = New-Object System.Drawing.Size(280,20) 
	$objForm.Controls.Add($objTxtDesc) 
	
	$objPkey = New-Object System.Windows.Forms.Label
	$objPkey.Location = New-Object System.Drawing.Size(10,180) 
	$objPkey.Size = New-Object System.Drawing.Size(80,20) 
	$objPkey.Text = "Product Key:"
	$objForm.Controls.Add($objPkey)
	
	$objTxtPkey = New-Object System.Windows.Forms.TextBox 
	$objTxtPkey.Location = New-Object System.Drawing.Size(90,180) 
	$objTxtPkey.Size = New-Object System.Drawing.Size(280,20) 
	$objForm.Controls.Add($objTxtPkey) 
	
	$objStat = New-Object System.Windows.Forms.Label
	$objStat.Location = New-Object System.Drawing.Size(10,200) 
	$objStat.Size = New-Object System.Drawing.Size(80,20)
	$objStat.Text = "Asset Status:"
	$objForm.Controls.Add($objStat)
	
	$objTxtStat = New-Object System.Windows.Forms.Label
	$objTxtStat.Location = New-Object System.Drawing.Size(90,200) 
	$objTxtStat.Size = New-Object System.Drawing.Size(280,20) 
	$objTxtStat.Text = "1"
	$objForm.Controls.Add($objTxtStat)
	
	$objForm.Topmost = $True
	
	$objForm.Add_Shown({$objForm.Activate()})
	$formResult = $objForm.ShowDialog()
	
	IF ($formResult -eq [System.Windows.Forms.DialogResult]::Cancel)
	{
	BREAK
	}

	##Assign Variables to send to SQL
	$SQName = $objTxtAname.text
	$SQTypeName = $objTxtAtype.text
	$SQSerial = $objTxtSnum.text
	$SQManufacturer = $objTxtManu.text
	$SQModel = $objTxtModel.text
	$CmpDescription = $objTxtDesc.text
	$SQProdKey = $objTxtPkey.text
	[int]$SQStat = 1
	
	Send-SQLinfo
}

<#
#>
 
Function Get-AssetID
{
	<#
	.SYNOPSIS
	Retrieve an Asset ID from SQL, and reserve spot
	#>
	
    ##Tell user what is happening
	$wshell = New-Object -ComObject Wscript.Shell
	$wshell.Popup("Retrieving Asset ID for $($comp).",3,"Progress Update",0x0)

	##Set Location of PS instance to SQL Database
	Set-Location SQLSERVER:\SQL\PROMETHEUS\DEFAULT\Databases\Assets\Tables
	
	##Set variable for asset id
	$Assetted = Invoke-Sqlcmd "SELECT MAX(asset_id)+1 FROM dbo.AssetList Where asset_id IS NOT NULL AND asset_id < 9990000;"
	$Assetted = $Assetted.column1

	##Assign $Assetted to asset_id now in case of new query before this finishes
	Invoke-Sqlcmd "INSERT INTO dbo.AssetList (asset_id) VALUES ('$Assetted')"

	##Return to default PS location
	Set-Location C:\WINDOWS\system32
	
	Get-ManualRegistration
}

<#
#>
 
Is-ITComp
}