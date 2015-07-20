[CmdletBinding()]
Param (
    [Parameter(Mandatory=$False,Position=0)]
	[switch]$PushToStrap
)

$VerbosePreference = "Continue"

if ($PushToStrap) {
    & ".\buildmodule.ps1" -PushToStrap
} else {
    & ".\buildmodule.ps1"
}

ipmo .\*.psd1

$ChassisName = "S4-CHASSIS"

Start-Visio -Visible -Quiet -Verbose
$global:stencilimport = Import-VisioStencilFile "C:\Users\brian.ADDICKS\Documents\My Shapes\S-Series_040215 - Visio Stencils.vss"
$global:stencil = select-visiostencil $ChassisName -verbose
$global:shape = Add-VisioStencil $Stencil 2.7984 8.4697
$global:rectangle = Add-VisioRectangle 0.265625 6.3125 5.328125 5.9375 -TextBox -FontSize 18 -Text $ChassisName