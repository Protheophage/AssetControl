Function Set-BiosAsset
{
    <#
    .SYNOPSIS
    Set the BIOS/UEFI Asset Tag on a Dell PC

    .DESCRIPTION
    Checks if the Asset Tag value is set. If not quiry SQL (using Serial Number) for Asset ID, and set the BIOS/UEFI Asset Tag. Then updates Asset Name and Description in SQL.

    .PARAMETER Name
    This is the current name of the asset
    .PARAMETER IsOnline
    Switch. If set will force skipping of check that PC is online

    .EXAMPLE
    Set-BiosAsset -Name 'PC-42','PC-13','PC-88'
    Run process on multiple computers
    
    .EXAMPLE
    Set-BiosAsset -Name 'PC-42-SURFACE' -IsOnline
    Set-BiosAsset does a ping check to see if the asset is online by default.  The firewall settings on Surfaces blocks ICMP Ping requests, and causes the process to end.
    Set -IsOnline to run the process without checking if the PC is online 
    #>

    [CmdletBinding()]
    Param
    (
        [Parameter(ValueFromPipeline=$true)]
        [String[]]$Name,
        [switch]$IsOnline
    )

    Process
    {
        Foreach($N in $Name)
        {
            If(!$IsOnline)  
            {  
                #Check if PC is Online
                If(!(Test-Connection -Cn $N -BufferSize 16 -Count 1 -ea 0 -quiet))
                {
                    Write-Host $N" is not online"
                    Continue
                }
            }
            Else
            {
                #Gather Info From PC
                $manufact = (Get-WmiObject -ComputerName "$N" Win32_SystemEnclosure).Manufacturer
                $PcModel = (Get-WmiObject -ComputerName "$N" -Class Win32_ComputerSystem).model
                $srlnmbr = (get-wmiobject -computername "$N" win32_bios).serialnumber
                $biosAtag = (Get-WmiObject -ComputerName "$N" Win32_SystemEnclosure).SMBiosAssetTag
                If($manufact -like "*dell*")
                {
                    If([string]::IsNullOrWhiteSpace($biosAtag))
                    {
                        [INT]$biosAtag = (Get-AssetInfo -Serial $srlnmbr -LogPath 'None').asset_id
                        If($biosAtag -lt 90000)
                        {
                            #Set Asset Tag in BIOS
                            C:\sysinternals\PsExec.exe \\$N -accepteula -s powershell.exe "Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned -Force;Install-Module DellBIOSProvider -force;Import-Module DellBIOSProvider -force;si DellSmbios:\systeminformation\asset -Value $($biosAtag)"
                            #Update Asset Name & Description in SQL
                            Update-AssetName -Serial $srlnmbr -NewName $N -IsOnline 1
                            Write-Host $N" completed"
                        }
                        Else
                        {
                            Write-Host $N" Asset ID is not set in SQL. Please update the Asset ID in SQL."
                            Continue
                        } 
                    }
                    Else
                    {
                        Update-AssetName -Serial $srlnmbr -NewName $N -IsOnline 1
                        Update-AssetID -Name $N -NewID $biosAtag
                        Write-Host $N" completed"
                    }
                }
                ElseIf($PcModel -like "*surface*")
                {
                    If([string]::IsNullOrWhiteSpace($biosAtag) -OR $biosAtag -eq "0")
                    {
                        [INT]$biosAtag = (Get-AssetInfo -Serial $srlnmbr -LogPath 'None').asset_id
                        If($biosAtag -lt 90000)
                        {
                            #Set Asset Tag in BIOS
                            C:\sysinternals\PsExec.exe \\$N -accepteula -s powershell.exe "New-Item -Path 'c:\Program Files\' -Name 'SurfaceAssetTag' -ItemType 'directory' -Force"
                            Copy-Item "\\kite\IT Dept\Applications\Surface\Surface-Utils-Gen-Independent\Surface_Asset_Tag\AssetTag.exe" "\\$N\C$\Program Files\SurfaceAssetTag\AssetTag.exe" -Force
                            C:\sysinternals\PsExec.exe \\$N -accepteula -s powershell.exe "Start-Process -filepath 'C:\Program Files\SurfaceAssetTag\AssetTag.exe' -ArgumentList '-s $($biosAtag)'"
                            #Update Asset Name & Description in SQL
                            Update-AssetName -Serial $srlnmbr -NewName $N -IsOnline 1
                            Write-Host $N" completed"
                        }
                        Else
                        {
                            Write-Host $N" Asset ID is not set in SQL. Please update the Asset ID in SQL."
                            Continue
                        } 
                    }
                    Else
                    {
                        Update-AssetName -Serial $srlnmbr -NewName $N -IsOnline 1
                        Update-AssetID -Name $N -NewID $biosAtag
                        Write-Host $N" completed"
                    }
                }
                Else
                {
                    Write-Host $N" is not a Dell or a Surface."
                    Continue
                }
            }
        }
    }
    End
    {
        Write-Host "Process Complete"
    }
}