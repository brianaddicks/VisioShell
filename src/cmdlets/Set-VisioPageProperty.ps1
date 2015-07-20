function Set-VisioPageProperty {
    [CmdletBinding()]
	<#
        .SYNOPSIS
            Sets various page properties for a Visio Document.
	#>

	Param (
		[Parameter(Mandatory=$false)]
		[switch]$ResizeToFitContents
	)
	
	# Resize To Fit Contents
	$Resize = $Global:VisioShellInstance.ActivePage.ResizeToFitContents()
}