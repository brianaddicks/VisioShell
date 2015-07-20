function Select-VisioStencil {
    [CmdletBinding()]
	<#
        .SYNOPSIS
            Selects a Visio Stencil by name.
	#>

	Param (
		[Parameter(Mandatory=$true)]
		[string]$StencilName
	)
	
	$VerbosePrefix = "Select-VisioStencil:"
	
	# Check for and Select Stencil
	foreach ($d in ($global:VisioShellInstance.Documents | ? { $_.FullName -match "stencil" } )) {
		Write-Verbose "$VerbosePrefix Searching document stencil: $($d.FullName)."
		foreach ($m in $d.Masters) {
			Write-Verbose "$VerbosePrefix Checking master: $($m.Name)."
			if ($m.Name -eq $StencilName) {
				Write-Verbose "$VerbosePrefix Match Found."
				$SelectedStencil = $d.Masters.Item($StencilName)
				break
			}
		}
	}
	
	if ($SelectedStencil) {
		return $SelectedStencil
	} else {
		Throw "$VerbosePrefix Stencil named `"$StencilName`" not found."
	}		
}