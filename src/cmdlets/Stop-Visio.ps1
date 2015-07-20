function Stop-Visio {
    [CmdletBinding()]
	<#
        .SYNOPSIS
            Stops the instance of Microsoft Visio referenced in $global:VisioShellInstance
	#>

	Param (
		[Parameter(Mandatory=$false)]
		[switch]$Quiet
	)
	
	# Quit Visio
	$Global:VisioShellInstance.App.Quit()
	Remove-Variable -name VisioShellInstance -Scope Global
}