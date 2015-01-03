<#
.Synopsis
	Bulk-import from folder any component, supported by PowerShell (script, module, source code, .Net assembly). 

.Description
	Bulk-import from folder any component, supported by PowerShell. Supported components:

	* Script (.ps1) - imported using Dot-Sourcing.
	* Module (.psm1) - imported using Import-Module cmdlet
	* Source code (.cs, .vb, .js) - imported using Add-Type cmdlet
	* .Net assembly (.dll) - imported using Add-Type cmdlet

	WARNING: To import .PS1 scripts this function itself has to be dot-sourced! See examples.

.Parameter Path
	This parameter is required.

	A string representing the path, where components located. Wildcards are not supported.
	This is the folder, where you could keep your functions in .ps1 scripts, modules in .psm1 files,
	C# (VB.Net, JavaScript) code in .cs, .vb, .js files and .Net assemblies (.dll).

.Parameter Recurse
	This parameter is optional.

	A switch indicating whether or not to recurse into subdirectories. Default is no recursion.

.Parameter Type
	This parameter is optional.

	An array of component types, that would be imported. If this parameter is not specified,
	any supported component will be imported.

	WARNING: To import .PS1 scripts this function itself has to be dot-sourced! See examples.

	The value can be any combination of the following:

	Ps
		Script (.ps1) - imported using Dot-Sourcing.

	Psm
		Module (.psm1) - imported using Import-Module cmdlet

	Cs
		Source code (.cs) - imported using Add-Type cmdlet

	Vb
		Source code (.vb) - imported using Add-Type cmdlet

	Js
		Source code (.js) - imported using Add-Type cmdlet

	Asm
		.Net assembly (.dll) - imported using Add-Type cmdlet

.Parameter Include
	This parameter is optional.

	An array of strings. Only file names that match will be imported. Wildcards are permitted.

.Parameter Exclude
	This parameter is optional.

	An array of strings. File names that match will not be imported. Has priority over the Include parameter.
	Wildcards are permitted.

.Example
	. Import-Component 'C:\PsLib'

		Description
		-----------
		Import all supported components (.ps1, .psm1, .cs, .vb, .js, .dll) from folder 'C:\PsLib'.
		Note, that to be able to import .ps1 scripts into the current scope, this function is dot-sourced.

.Example
	'C:\PsLib', 'C:\MyLib' | . Import-Component 

		Description
		-----------
		Import all supported components (.ps1, .psm1, .cs, .vb, .js, .dll) from folders 'C:\PsLib' and 'C:\MyLib'.
		Note, that to be able to import .ps1 scripts into the current scope, this function is dot-sourced.

.Example
	. Import-Component 'C:\PsLib' -Recurse

		Description
		-----------
		Import all supported components (.ps1, .psm1, .cs, .vb, .js, .dll), recurse into subdirectories.
		Note, that to be able to import .ps1 scripts into the current scope, this function is dot-sourced.

.Example
	. Import-Component 'C:\PsLib' -Recurse -Include 'MyScript*'

		Description
		-----------
		Import all supported components (.ps1, .psm1, .cs, .vb, .js, .dll), recurse into subdirectories.
		Include only files with names that match wildcard 'MyScript*'. Note, that to be able to import .ps1
		scripts into the current scope, this function is dot-sourced.

.Example
	. Import-Component 'C:\PsLib' -Recurse -Include 'MyScript*' -Exclude '*_backup*'

		Description
		-----------
		Import all supported components (.ps1, .psm1, .cs, .vb, .js, .dll), recurse into subdirectories.
		Include only files with names that match wildcard 'MyScript*'.
		Exclude files with names that match '*_backup*' wildcard.
		Note, that to be able to import .ps1 scripts into the current scope, this function is dot-sourced.

.Example
	. Import-Component 'C:\PsLib' -Recurse -Include 'MyScript*','*MyLib*' -Exclude '*_backup*','*_old*'

		Description
		-----------
		Import all supported components (.ps1, .psm1, .cs, .vb, .js, .dll), recurse into subdirectories.
		Include only files with names that match wildcards 'MyScript*' and '*MyLib*'.
		Exclude files with names that match '*_backup*' and '*_old*' wildcards.
		Note, that to be able to import .ps1 scripts into the current scope, this function is dot-sourced.

.Example
	Import-Component 'C:\PsLib' -Type Ps

		Description
		-----------
		Import Powershell modules only (.psm1).

.Example
	Import-Component 'C:\PsLib' -Type Ps,Cs,Asm

		Description
		-----------
		Import Powershell modules (.psm1), C# source code (.cs), .Net assemblies (.dll).
#>
function Import-Component
{
	[CmdletBinding()]
	Param(
		[Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Position = 0)]
		[ValidateScript({Test-Path -LiteralPath $_ -PathType Container})]
		[string]$Path,

		[Parameter(Mandatory = $false, ValueFromPipeline = $false, ValueFromPipelineByPropertyName = $true)]
		[ValidateSet('Ps','Psm','Cs','Vb','Js', 'Asm')]
		[array]$Type,

		[Parameter(Mandatory = $false, ValueFromPipeline = $false, ValueFromPipelineByPropertyName = $true)]
		[array]$Include = '*',

		[Parameter(Mandatory = $false, ValueFromPipeline = $false, ValueFromPipelineByPropertyName = $true)]
		[array]$Exclude = $null,

		[Parameter(Mandatory = $false, ValueFromPipeline = $false, ValueFromPipelineByPropertyName = $true)]
		[switch]$Recurse
	)

	Begin
	{

		$MyCommand =  (Get-PSCallStack)[1].Position.Text
		$MyCommandName = $MyInvocation.MyCommand.Name
		$Split = $MyCommand -split "($MyCommandName)"
		$IsDotSourced = $Split[0] -match '\.\s*$'

		$DotSource = {. $_.FullName}
		$ImportModule = {Import-Module $_.FullName -ErrorAction SilentlyContinue}
		$AddType = {Add-Type -LiteralPath $_.FullName -ErrorAction SilentlyContinue}

		$FileType = @{
			Ps = @{Extension = '.ps1' ; Command = $DotSource}
			Psm = @{Extension = '.psm1' ; Command = $ImportModule}
			Cs = @{Extension = '.cs' ; Command = $AddType}
			Vb = @{Extension = '.vb' ; Command = $AddType}
			Js = @{Extension = '.js' ; Command = $AddType}
			Asm = @{Extension = '.dll' ; Command = $AddType}
		}
	}

	Process
	{
		if($Type)
		{
			[array]$FileTypeToImport =  $Type | Sort-Object -Unique | ForEach-Object {$FileType.$_}
		}
		else
		{
			[array]$FileTypeToImport = $FileType.GetEnumerator() | ForEach-Object {$_.Value}
		}

		if(($FileTypeToImport).Extension -contains '.ps1' -and !$IsDotSourced)
		{
			Write-Warning "To import .PS1 scripts this function itself has to be dot-sourced! Example: $($Split[$Split.IndexOf($MyCommandName)] = ". $MyCommandName" ; $Split -join '')"
		}

		filter Skip-File {
			if
			(
				!($_.PSIsContainer) -and 
				$($tmp = $_.Name ; $Exclude | ForEach-Object {if($tmp -notlike $_){$true}}) -and
				$($tmp = $_.Name ; $Include | ForEach-Object {if($tmp -like $_){$true}})
			)
			{
				$_
			}
		}

		$FileTypeToImport |
			ForEach-Object {
				$Private:currFT = $_
				Get-ChildItem -LiteralPath $Path -Filter ('*' + $_.Extension) -Recurse:$Recurse |
					Skip-File |
						ForEach-Object {
							Try
							{
								. $Private:currFT.Command
								$Success = $true
								$ErrMsg = $null
							}
							Catch
							{
								$Success = $false
								$ErrMsg = $_
							}

							$ret = @{
								Name = $_.Name
								Loaded = $Success
								Path = $_.FullName
								Error = $ErrMsg
							}

							New-Object PSObject -Property $ret | Select-Object Name, Path, Loaded, Error
						}
			}
	}
}