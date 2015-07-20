function Import-VisioStencilFile {
    [CmdletBinding()]
	<#
        .SYNOPSIS
            Imports a stencil into the Visio Instance.
	#>

	Param (
		[Parameter(Mandatory=$true)]
		[string]$StencilPath
	)
	
	# Import Visio Stencil
	$LoadStencil = $Global:VisioShellInstance.Documents.Add($StencilPath)
	return $LoadStencil
}