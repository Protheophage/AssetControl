Function Install-DellBIOSProvider
{
	<#
	.SYNOPSIS
	Check if DellBIOSProvider is installed and install if not
	#>
	[CmdletBinding()]
    $DBPMod = Get-Module -ListAvailable -Name "DellBIOSProvider"
    If (-Not $DBPMod)
	{
		Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned -Force
		Install-PackageProvider -Name "Nuget" -force
		Install-Module -Name "DellBIOSProvider" -force
		Import-Module -Name "DellBIOSProvider"
	}
	Else
	{
		Write-Verbose "DellBIOSProvider module is already installed. Importing now."
		Import-Module -Name "DellBIOSProvider"
	}
}