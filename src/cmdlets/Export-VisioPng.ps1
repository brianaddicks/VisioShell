function Export-VisioPng {
    [CmdletBinding()]
	<#
        .SYNOPSIS
            Exports a Visio document as a PNG.
	#>

	Param (
		[Parameter(Mandatory=$true,Position=0)]
		[string]$PngPath,
		
		[Parameter(Mandatory=$false)]
		[ValidatePattern('\d+x\d+')]
		[string]$Resolution
	)
	
	$VerbosePrefix = "Export-VisioPng:"
	
	# Import Visio Stencil
	if ($Resolution) {
		$SplitRes = $Resolution.Split('x')
		$XRes = [double]$SplitRes[0]
		$YRes = [double]$SplitRes[1]
		Write-Verbose "$VerbosePrefix $XRes x $YRes"
		$Global:VisioShellInstance.App.Settings.SetRasterExportResolution(1, $XRes, $YRes, 0)
	}
	
	$Global:VisioShellInstance.ActivePage.Export($PngPath)
}