function Set-VisioShapeFont {
    [CmdletBinding()]
	<#
        .SYNOPSIS
            Configures various Font Properties of a given shape.
	#>

	Param (
		[Parameter(Mandatory=$true,Position=0)]
		[object]$ShapeObject,
		
		[Parameter(Mandatory=$false)]
		[ValidatePattern('[a-fA-F0-9]{6}')]
		[string]$ColorInHex
	)
	
	$VerbosePrefix = "Set-VisioShapeFont:"
	
	$Red   = [Convert]::ToInt32($ColorInHex.SubString(0,2),16)
	Write-Verbose "$VerbosePrefix Red: $([string]$Red)"
		
	$Green = [Convert]::ToInt32($ColorInHex.SubString(2,2),16)
	Write-Verbose "$VerbosePrefix Green: $Green"
		
	$Blue  = [Convert]::ToInt32($ColorInHex.SubString(4,2),16)
	Write-Verbose "$VerbosePrefix Blue: $Blue"
		
	$UpdateColor = $ShapeObject.CellsSRC(3,0,1).Formula = "THEMEGUARD(RGB($Red,$Green,$Blue))"
}