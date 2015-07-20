function Save-VisioDocument {
    [CmdletBinding()]
	<#
        .SYNOPSIS
            Saves active Document as a VSD file. For other formats, see "gcm -module VisioShell -Verb Export"
	#>

	Param (
		[Parameter(Mandatory=$true,Position=0)]
		[string]$VsdPath
	)
	
	# Save File
	$Save = $Global:VisioShellInstance.ActiveDocument.SaveAs($VsdPath)
}