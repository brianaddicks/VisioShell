function Add-VisioStencil {
    [CmdletBinding()]
	<#
        .SYNOPSIS
            Drops a given Stencil onto the ActivePage at the given coordinates.
	#>

	Param (
		[Parameter(Mandatory=$true,Position=0)]
		[object]$StencilObject,
		
		[Parameter(Mandatory=$true,Position=1)]
		[double]$PinX,
		
		[Parameter(Mandatory=$true,Position=2)]
		[double]$PinY
	)
	
	$VerbosePrefix = "Add-VisioStencil:"
	
	# Drop Shape
	$Shape = $global:VisioShellInstance.ActivePage.Drop($StencilObject, $PinX, $PinY)
	return $Shape
}