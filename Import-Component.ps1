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
	This is the folder, where you could keep your functions in .ps1 scripts, modules,
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
		Module - imported using Import-Module cmdlet

		This function will only try to import well-formed modules. A "well-formed" module is a module
		that is stored in a directory that has the same name as the base name of at least one file in
		the module directory. If a module is not well-formed, Windows PowerShell does not recognize it
		as a module.

		The "base name" of a file is the name without the file name extension. In a well-formed module,
		the name of the directory that contains the module files must match the base name of at least
		one file in the module.

		More info: https://msdn.microsoft.com/en-us/library/dd878350.aspx

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

	An array of strings. Only file names without extension that match will be imported. Wildcards are permitted.

.Parameter Exclude
	This parameter is optional.

	An array of strings. File names without extension that match will not be imported. Has priority over the Include parameter.
	Wildcards are permitted.

.Example
	. Import-Component 'C:\PsLib'

		Description
		-----------
		Import all supported components (.ps1, modules, .cs, .vb, .js, .dll) from folder 'C:\PsLib'.
		Note, that to be able to import .ps1 scripts into the current scope, this function is dot-sourced.

.Example
	'C:\PsLib', 'C:\MyLib' | . Import-Component

		Description
		-----------
		Import all supported components (.ps1, modules, .cs, .vb, .js, .dll) from folders 'C:\PsLib' and 'C:\MyLib'.
		Note, that to be able to import .ps1 scripts into the current scope, this function is dot-sourced.

.Example
	. Import-Component 'C:\PsLib' -Recurse

		Description
		-----------
		Import all supported components (.ps1, modules, .cs, .vb, .js, .dll), recurse into subdirectories.
		Note, that to be able to import .ps1 scripts into the current scope, this function is dot-sourced.

.Example
	. Import-Component 'C:\PsLib' -Recurse -Include 'MyScript*'

		Description
		-----------
		Import all supported components (.ps1, modules, .cs, .vb, .js, .dll), recurse into subdirectories.
		Include only files with names without extension that match wildcard 'MyScript*'.
		Note, that to be able to import .ps1 scripts into the current scope, this function is dot-sourced.

.Example
	. Import-Component 'C:\PsLib' -Recurse -Include 'MyScript*' -Exclude '*_backup*'

		Description
		-----------
		Import all supported components (.ps1, modules, .cs, .vb, .js, .dll), recurse into subdirectories.
		Include only files with names without extension that match wildcard 'MyScript*'.
		Exclude files with names without extension that match '*_backup*' wildcard.
		Note, that to be able to import .ps1 scripts into the current scope, this function is dot-sourced.

.Example
	. Import-Component 'C:\PsLib' -Recurse -Include 'MyScript*','*MyLib*' -Exclude '*_backup*','*_old*'

		Description
		-----------
		Import all supported components (.ps1, modules, .cs, .vb, .js, .dll), recurse into subdirectories.
		Include only files with names without extension that match wildcards 'MyScript*' and '*MyLib*'.
		Exclude files with names without extension that match '*_backup*' and '*_old*' wildcards.
		Note, that to be able to import .ps1 scripts into the current scope, this function is dot-sourced.

.Example
	Import-Component 'C:\PsLib' -Type Psm

		Description
		-----------
		Import Powershell modules only.

.Example
	Import-Component 'C:\PsLib' -Type Ps,Cs,Asm

		Description
		-----------
		Import Powershell modules, C# source code (.cs), .Net assemblies (.dll).
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
		Write-Verbose "Function dot-sourced: $IsDotSourced"

		# Define scriptblocks that will import verious file types
		$DotSource = {. $_.FullName}
		$AddType = {Add-Type -LiteralPath $_.FullName -ErrorAction SilentlyContinue}
		$ImportModule = {
			if(Get-ChildItem -Path ($_.FullName + '\*') -Filter ($_.Basename + '.*'))
			{
				Write-Verbose "Folder '$($_.Basename)' looks like well-formed module"
				# https://msdn.microsoft.com/en-us/library/dd878350.aspx
				Import-Module $_.FullName -ErrorAction SilentlyContinue
			}
		}

		# Define extensions for file types and assign import commands
		$FileType = @{
			Ps = @{Extension = '.ps1' ; Command = $DotSource}
			Psm = @{Extension = [string]::Empty ; Command = $ImportModule}
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
			Write-Verbose "Trying to import file types $($Type -join ',')"
			[array]$FileTypeToImport =  $Type | Sort-Object -Unique | ForEach-Object {$FileType.$_}
		}
		else
		{
			Write-Verbose 'Trying to import all supported file types'
			[array]$FileTypeToImport = $FileType.GetEnumerator() | ForEach-Object {$_.Value}
		}

		if(($FileTypeToImport).Extension -contains '.ps1' -and !$IsDotSourced)
		{
			Write-Warning "To import .PS1 scripts this function itself has to be dot-sourced! Example: $($Split[$Split.IndexOf($MyCommandName)] = ". $MyCommandName" ; $Split -join '')"
		}

		# Filter function to process Include\Exclude parameters
		filter Skip-File {
			if
			(
				# No file extension and directory = module, else - file
				((!$currFT.Extension -and $_.PSIsContainer) -or ($currFT.Extension -and !$_.PSIsContainer)) -and
				$($tmp = $_.BaseName ; $Exclude | ForEach-Object {if($tmp -notlike $_){$true}}) -and
				$($tmp = $_.BaseName ; $Include | ForEach-Object {if($tmp -like $_){$true}})
			)
			{
				$_
			}
		}

		$FileTypeToImport |
			ForEach-Object {
				# We need to pass current file type to the filter function later
				Set-Variable -Name currFT -Value $_ -Option AllScope
				Write-Verbose "Searching path '$Path' for file type '$(('*' + $_.Extension))'"
				Get-ChildItem -LiteralPath $Path -Filter ('*' + $_.Extension) -Recurse:$Recurse |
					Skip-File |
						ForEach-Object {
							Try
							{
								Write-Verbose "Trying to import: $_"
								. $currFT.Command
								$Success = $true
								$ErrMsg = $null
								Write-Verbose 'Success'
							}
							Catch
							{
								Write-Verbose 'Failed'
								$Success = $false
								$ErrMsg = $_
							}

							$ret = @{
								Name = $_.Name
								Loaded = $Success
								Path = $_.FullName
								Error = $ErrMsg
							}

							Write-Verbose "Writing import status for '$_' to pipeline"
							New-Object PSObject -Property $ret | Select-Object Name, Path, Loaded, Error
						}
			}
	}
}