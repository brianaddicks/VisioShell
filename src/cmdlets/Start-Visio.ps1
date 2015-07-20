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
	$ReturnObject.Documents = $ReturnObject.App.Documents
	$AddDoc = $ReturnObject.App.Documents.Add("")

	# Select first page
	Write-Verbose "$VerbosePrefix Select Page 1 as Active Page"
	$ReturnObject.ActivePage = $ReturnObject.App.ActiveDocument.Pages.Item(1)
	
	Write-Verbose "$VerbosePrefix Set Global Variable"
	$Global:VisioShellInstance = $ReturnObject
	
	if (!($Quiet)) {
		return $ReturnObject
	}
}