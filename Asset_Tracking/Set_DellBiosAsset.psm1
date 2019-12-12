#Check if PC is Online
if(!(Test-Connection -Cn $comp -BufferSize 16 -Count 1 -ea 0 -quiet))
{
    $wshell = New-Object -ComObject Wscript.Shell
    $wshell.Popup("$($comp) is not online.",3,$cmpFnd.asset_name,0x0)
    Break
}
$manufact = (Get-WmiObject -ComputerName "$comp" Win32_SystemEnclosure).Manufacturer
If($manufact = "*dell*")
{
    echo "its a dell"
}
#Check if PC is Dell
##if not
###Exit
#Else 
#Gather Serial number
#Gather Description
#Check BIOS for Asset Tag
##If Not
###check SQL for Serial Number and get Asset ID
####If Asset ID less than 90000 & !Null
#####Send to DellBIOS
#####Update SQL Asset Name by Serial Number
#####Update SQL Description by Serial Number
####Else
#####Warn that Asset ID not recorded in SQL
#####Exit
##Else
###Update SQL Asset Name by Serial Number
###Update SQL Asset ID by Serial Number
###Update SQL Description by Serial Number

##From Get-AssetInfo
ForEach($num in $Serial)
			{
			$cmpFnd = Invoke-Sqlcmd "SELECT * FROM [dbo].[AssetList] Where serial_number LIKE '$num';"
			$ChangeLog = $ChangeLog + ($cmpFnd)

Is-Online


##Get computer description and assign to CmpDescription
	$CmpDescription = (Get-WmiObject -ComputerName "$comp" -Class Win32_OperatingSystem).Description
	
	##Get Serial number and assign to srlnmbr
	$srlnmbr = get-wmiobject -computername "$comp" win32_bios serialnumber

	##Get Windows Product Key, and assign to $ProdKey
	$GetPK = get-wmiObject -computername "$comp" -query 'select * from SoftwareLicensingService'
	$ProdKey = $GetPK.OA3xOriginalProductKey
    
    ##Check for BIOS Asset Tag
    $abox = (Get-WmiObject -ComputerName "$comp" Win32_SystemEnclosure).SMBiosAssetTag
    If (-Not $abox) 
    {
        $srlnmbr = get-wmiobject -computername "$comp" win32_bios serialnumber
    }
    Else 
    {
        $abox
    }


(Get-WmiObject -ComputerName "$comp" Win32_SystemEnclosure).SMBiosAssetTag

$abox = (Get-WmiObject -ComputerName "$comp" Win32_SystemEnclosure).SMBiosAssetTag
If (-Not $abox) {echo "its empty"} Else {$abox}

C:\sysinternals\PsExec.exe \\UC-1jxnrz2 -accepteula -s powershell.exe '"Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned -Force"; "Import-Module DellBIOSProvider -force"; "get-childitem DellSmbios:\systeminformation\asset | select CurrentValue"; "$env:computername"'