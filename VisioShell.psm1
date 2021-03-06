###############################################################################
## Start Powershell Cmdlets
###############################################################################

###############################################################################
# Add-VisioRectangle

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

###############################################################################
# Add-VisioStencil

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

###############################################################################
# Export-VisioPng

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

###############################################################################
# Import-VisioStencilFile

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

###############################################################################
# Save-VisioDocument

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

###############################################################################
# Select-VisioStencil

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

###############################################################################
# Set-VisioPageProperty

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

###############################################################################
# Set-VisioShapeFont

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

###############################################################################
# Start-Visio

function Start-Visio {
    [CmdletBinding()]
	<#
        .SYNOPSIS
            Starts Microsoft Visio, Creates a new document, and selects the first page.  Returns the Visio Object, but also sets the global variable $global:VisioShellInstance so you don't have to pass the variable for every subsequent cmdlet.
	#>

	Param (
		[Parameter(Mandatory=$false)]
		[switch]$Visible,
		
		[Parameter(Mandatory=$false)]
		[switch]$Quiet
	)
	
	$VerbosePrefix = "Start-Visio:"
	
	# Custom return object
	$ReturnObject = @{}
	
	#Start Application
	Write-Verbose "$VerbosePrefix Create Visio App Instance"
	$ReturnObject.App = New-Object -ComObject Visio.Application
	
	#Set Visible as desired, default is hidden
	if ($Visible) {
		Write-Verbose "$VerbosePrefix Set Visible to True"
		$ReturnObject.App.Visible = $true
	} else {
		Write-Verbose "$VerbosePrefix Set Visible to False"
		$ReturnObject.App.Visible = $false
	}
	
	# Create New Document
	Write-Verbose "$VerbosePrefix Create new Document"
	$ReturnObject.Documents      = $ReturnObject.App.Documents
	$ReturnObject.ActiveDocument = $ReturnObject.App.Documents.Add("")

	# Select first page
	Write-Verbose "$VerbosePrefix Select Page 1 as Active Page"
	$ReturnObject.ActivePage = $ReturnObject.App.ActiveDocument.Pages.Item(1)
	
	Write-Verbose "$VerbosePrefix Set Global Variable"
	$Global:VisioShellInstance = $ReturnObject
	
	if (!($Quiet)) {
		return $ReturnObject
	}
}

###############################################################################
# Stop-Visio

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

###############################################################################
## Export Cmdlets
###############################################################################

Export-ModuleMember *-*
