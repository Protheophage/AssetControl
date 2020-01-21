Function Confirm-ITComp
{
	<#
	.SYNOPSIS
	Check if this is being run on an IT computer and quit if not
	#>
	If ($env:computername -NotLike "IT-*")
	{
        ##Needs changed to a pop-up warning for GUI compatibility
		Write-Output "Please run this from an IT department computer."
		Break
	}
}