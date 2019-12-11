Function Install-SQLModule
{
	<#
	.SYNOPSIS
	Check if SqlServer is installed and install if not
	#>
	[CmdletBinding()]
    $SQMod = Get-Module -ListAvailable -Name "SqlServer"
    If (-Not $SQMod)
	{
		Install-Module -Name "SqlServer"
		Import-Module -Name "SqlServer"
	}
	Else
	{
		Write-Verbose "SqlServer module is already installed. Importing now."
		Import-Module -Name "SqlServer"
	}
}