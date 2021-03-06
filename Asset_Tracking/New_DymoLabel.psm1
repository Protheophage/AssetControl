Function New-DymoLabel
{
	<#
	.SYNOPSIS
	Print an Asset Label from the Dymo printer
	
	.Description
	Check if the Dymo printer and Dymo software are installed.  Install if not. Then, print the asset label.
	
	.PARAMETER AstNum
	The asset number to be printed.
	.PARAMETER AstModel
	The description (asset model) for the label.
	
	.EXAMPLE
	New-DymoLabel -AstNum '1942' AstModel 'Surface Pro 4'
	#>
	
	[CmdletBinding()]
	Param
	(
		[parameter(Mandatory=$true)]
		[String[]]$AstNum,
		[parameter(Mandatory=$true)]
		[String[]]$AstModel
	)
	BEGIN
	{
		$dymoPrint = Get-Printer -Name "*dymo*"
		If(!$dymoPrint)
		{
			Write-Output 'You do not have a DYMO printer'
			Write-Output 'Adding DYMO printer <Path-to-shared-dymoPrinter>'
			Add-Printer -ConnectionName "<Path-to-shared-dymoPrinter>"
		}
		$dymmoSoft = get-package -Name "*Dymo Label*" -ErrorAction SilentlyContinue
		If(!$dymmoSoft)
		{
			start-process msiexec.exe -argumentList '/i', '"<Path-to-DymoLabelMsi>\DYMO Label.msi"', '/passive' -PassThru -verb runas -wait
		}
	}
	PROCESS
	{
		$scrbck = 
		{
			param($ANum,$AModel)
			[reflection.assembly]::LoadFile('C:\Program Files (x86)\DYMO\DYMO Label Software\Framework\DYMO.DLS.Runtime.dll')
			[reflection.assembly]::LoadFile('C:\Program Files (x86)\DYMO\DYMO Label Software\Framework\DYMO.Label.Framework.dll')
			[reflection.assembly]::LoadFile('C:\Program Files (x86)\DYMO\DYMO Label Software\Framework\DYMO.Common.dll')

			$printername = [DYMO.Label.Framework.Framework]::GetLabelWriterPrinters() | Select-Object -ExpandProperty name

			$labelfile = '<Path-to-DymoLabel-Template> Template.label'

			$label = [Dymo.label.framework.label]::open($labelfile)
			$label.SetObjectText('Asset Number',$ANum)
			$label.SetObjectText('Model',$AModel)
			$label.Print($printername)
		}
		$i = 0
		ForEach($AN in $AstNum)
		{
			$AM = $AstModel[$i]
			Start-Job -ScriptBlock $scrbck -RunAs32 -Args $AN,$AM
			$i = $i + 1
		}
	}
}