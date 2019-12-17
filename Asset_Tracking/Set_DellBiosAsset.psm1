Function Set-DellBiosAsset
{
    <#
    .SYNOPSIS
    Set the BIOS/UEFI Asset Tag on a Dell PC

    .DESCRIPTION
    Checks if the Asset Tag value is set. If not quiry SQL (using Serial Number) for Asset ID, and set the BIOS/UEFI Asset Tag. Then updates Asset Name and Description in SQL.

    .PARAMETER Name

    .EXAMPLE
    Set-DellBiosAsset -Name 'PC-42','PC-13','PC-88'
    Run process on multiple computers
    #>

    [CmdletBinding()]
    Param
    (
        [Parameter(ValueFromPipeline=$true)]
        [String[]]$Name
    )

    Process
    {
        Foreach($N in $Name)
        {
            #Check if PC is Online
            If(!(Test-Connection -Cn $N -BufferSize 16 -Count 1 -ea 0 -quiet))
            {
                Continue
            }
            Else
            {
                #Check if PC is Dell
                $manufact = (Get-WmiObject -ComputerName "$N" Win32_SystemEnclosure).Manufacturer
                $biosAtag = ''
                If($manufact = "*dell*")
                {
                    #Gather Serial number
                    $srlnmbr = (get-wmiobject -computername "$N" win32_bios).serialnumber
                    #Gather Description
                    $CmpDescription = (Get-WmiObject -ComputerName "$N" -Class Win32_OperatingSystem).Description
                    #Check BIOS for Asset Tag
                    $biosAtag = (Get-WmiObject -ComputerName "$N" Win32_SystemEnclosure).SMBiosAssetTag
                    If([string]::IsNullOrWhiteSpace($biosAtag))
                    {
                        [INT]$biosAtag = (Get-AssetInfo -Serial $srlnmbr -LogPath 'None').asset_id
                        If($biosAtag -lt 90000)
                        {
                            #Set Asset Tag in BIOS
                            C:\sysinternals\PsExec.exe \\$N -accepteula -s powershell.exe "Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned -Force;Install-Module DellBIOSProvider -force;Import-Module DellBIOSProvider -force;si DellSmbios:\systeminformation\asset -Value $($biosAtag)"
                            #Update Asset Name & Description in SQL
                            Update-AssetName -Serial $srlnmbr -NewName $N
                        }
                        Else
                        {
                            Write-Host $N"Asset ID is not set in SQL. Please update the Asset ID in SQL."
                            Continue
                        }

                    }
                    Else
                    {
                        Update-AssetName -Serial $srlnmbr -NewName $N
                        Update-AssetID -Name $N -NewID $biosAtag
                    }
                }
                Else
                {
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