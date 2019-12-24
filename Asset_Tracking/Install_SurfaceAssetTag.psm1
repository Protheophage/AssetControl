Function Install-SurfaceAssetTag
{
	<#
	.SYNOPSIS
	Check if Surface AssetTag is installed and install if not
	#>
	[CmdletBinding()]
	$SurfAstTagPath = "C:\Program Files\SurfaceAssetTag"
	If (!(Test-Path $SurfAstTagPath))
	{
		$SurfAstTagzip = "$SurfAstTagPath\SurfaceAssetTag.zip"
        New-Item $SurfAstTagPath -ItemType Directory
        $URL = "https://download.microsoft.com/download/8/2/E/82EEFB07-1AB3-4557-B654-B34D64C9DD94/Surface%20Asset%20Tag.zip" 
        $WebClient = New-Object System.Net.WebClient
		$WebClient.DownloadFile($URL,$SurfAstTagzip)
		$WebClient.Dispose()
		Expand-Archive -LiteralPath $SurfAstTagzip -DestinationPath $SurfAstTagPath
	}
	Else
	{
		Write-Verbose "Surface Asset Tag is already installed. Adding to Path."
	}
	[Environment]::SetEnvironmentVariable("Path",$Env:Path + ";C:\Program Files\SurfaceAssetTag\Surface Asset Tag", "Machine")
}