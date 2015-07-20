function Add-VisioRectangle {
    [CmdletBinding()]
	<#
        .SYNOPSIS
            Draws a Rectangle with the given Coordinates.
	#>

	Param (
		[Parameter(Mandatory=$true,Position=0)]
		[double]$x1,
		
		[Parameter(Mandatory=$true,Position=1)]
		[double]$y1,
		
		[Parameter(Mandatory=$true,Position=2)]
		[double]$x2,
		
		[Parameter(Mandatory=$true,Position=3)]
		[double]$y2,
		
		[Parameter(Mandatory=$false)]
		[switch]$TextBox,
		
		[Parameter(Mandatory=$false)]
		[int]$FontSize,
		
		[Parameter(Mandatory=$false)]
		[string]$Text
	)
	
	$VerbosePrefix = "Add-VisioRectangle:"
	
	# Draw Rectangle
	$Rectangle = $global:VisioShellInstance.ActivePage.DrawRectangle($x1, $y1, $x2, $y2)
	
	# Set Textbox settings
	if ($TextBox) {
		$Rectangle.TextStyle = "Normal"
		$Rectangle.LineStyle = "Text Only"
		$Rectangle.FillStyle = "Text Only"
	}
	
	# Set Font Size
	if ($FontSize) {
		$Rectangle.CellsSRC(3,0,7).Formula = [string]$FontSize + " pt"
	}
	
	# Set Text
	if ($Text) {
		$Rectangle.Text = $Text
	}
	
	return $Rectangle
}