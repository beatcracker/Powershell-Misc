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
	Param()
	DynamicParam
	{
		# Stripped down version of the New-DynamicParameter function
		# https://gallery.technet.microsoft.com/scriptcenter/New-DynamicParameter-63389a46
		New-Module -OutVariable null -ReturnResult -AsCustomObject -ScriptBlock {
			$DynamicParameters = @(
				@{
					Name = 'Path'
					Type = [string]
					Position = 0
					Mandatory = $true
					ValueFromPipeline = $true
					ValueFromPipelineByPropertyName = $true
					ValidateScript = {Test-Path -LiteralPath $_ -PathType Container}
				},
				@{
					Name = 'Type'
					Type = [string[]]
					Position = 1
					ValueFromPipelineByPropertyName = $true
					ValidateSet = 'Ps','Psm','Cs','Vb','Js', 'Asm'
				},
				@{
					Name = 'Include'
					Type = [string]
					Position = 2
					ValueFromPipelineByPropertyName = $true
				}
				@{
					Name = 'Exclude'
					Type = [string]
					Position = 3
					ValueFromPipelineByPropertyName = $true
				},
				@{
					Name = 'Recurse'
					Type = [switch]
					Position = 4
					ValueFromPipelineByPropertyName = $true
				}
			)
			# Creating new dynamic parameters dictionary
			$Dictionary = New-Object -TypeName System.Management.Automation.RuntimeDefinedParameterDictionary

			# Strings to match attributes and validation arguments
			$AttributeRegex = '^(Mandatory|Position|ParameterSetName|DontShow|ValueFromPipeline|ValueFromPipelineByPropertyName|ValueFromRemainingArguments)$'
			$ValidationRegex = '^(AllowNull|AllowEmptyString|AllowEmptyCollection|ValidateCount|ValidateLength|ValidatePattern|ValidateRange|ValidateScript|ValidateSet|ValidateNotNull|ValidateNotNullOrEmpty)$'
			$AliasRegex = '^Alias$'

			$DynamicParameters | ForEach-Object {
				# Creating new parameter''s attirubutes object'
				$ParameterAttribute = New-Object -TypeName System.Management.Automation.ParameterAttribute

				# Looping through the bound parameters, setting attirubutes
				$CurrentParam = $_
				switch -regex ($_.Keys)
				{
					$AttributeRegex
					{
						Try
						{
							# Adding new parameter attribute
							$ParameterAttribute.$_ = $CurrentParam[$_]
						}
						Catch {$_}
						continue
					}
				}

				# Looping through the bound parameters, setting attirubutes
				switch -regex ($_.Keys)
				{
					$AttributeRegex
					{
						Try
						{
							# Adding new parameter attribute
							$ParameterAttribute.$_ = $CurrentParam[$_]
						}
						Catch {$_}
						continue
					}
				}

				# Creating new attribute collection object
				$AttributeCollection = New-Object -TypeName Collections.ObjectModel.Collection[System.Attribute]

				# Looping through the bound parameters, adding attributes
				switch -regex ($_.Keys)
				{
					$ValidationRegex
					{
						Try
						{
							# Adding attribute
							$ParameterOptions = New-Object -TypeName "System.Management.Automation.$_`Attribute" -ArgumentList $CurrentParam[$_] -ErrorAction SilentlyContinue
							$AttributeCollection.Add($ParameterOptions)
						}
						Catch {$_}
						continue
					}

					$AliasRegex
					{
						Try
						{
							# Adding alias
							$ParameterAlias = New-Object -TypeName System.Management.Automation.AliasAttribute -ArgumentList $CurrentParam[$_] -ErrorAction SilentlyContinue
							$AttributeCollection.Add($CurrentParam[$_])
							continue
						}
						Catch {$_}
					}
				}

				# Adding attributes to the attribute collection
				$AttributeCollection.Add($ParameterAttribute)

				# Finishing creation of the new dynamic parameter
				$Parameter = New-Object -TypeName System.Management.Automation.RuntimeDefinedParameter -ArgumentList @($_.Name, $_.Type, $AttributeCollection)

				# Adding dynamic parameter to the dynamic parameter dictionary
				$Dictionary.Add($_.Name, $Parameter)
			}

			# Writing dynamic parameter dictionary to the pipeline
			$Dictionary
		}
	}

	Process
	{
	New-Module -OutVariable null -ReturnResult -AsCustomObject -ScriptBlock {
			Param
			(
				$Path,
				$Type,
				$Exclude,
				$Include,
				$Recurse,
				$DotSource,
				$AddType,
				$ImportModule,
				$IsDotSourced,
				$SessionState
			)

			# Inherit parent's function preferences
			# More info: https://gallery.technet.microsoft.com/scriptcenter/Inherit-Preference-82343b9d
			$PreferenceVars = @(
				'ErrorView', 'FormatEnumerationLimit', 'LogCommandHealthEvent', 'LogCommandLifecycleEvent',
				'LogEngineHealthEvent', 'LogEngineLifecycleEvent', 'LogProviderHealthEvent',
				'LogProviderLifecycleEvent', 'MaximumAliasCount', 'MaximumDriveCount', 'MaximumErrorCount',
				'MaximumFunctionCount', 'MaximumHistoryCount', 'MaximumVariableCount', 'OFS', 'OutputEncoding',
				'ProgressPreference', 'PSDefaultParameterValues', 'PSEmailServer', 'PSModuleAutoLoadingPreference',
				'PSSessionApplicationName', 'PSSessionConfigurationName', 'PSSessionOption', 'ErrorActionPreference',
				'DebugPreference', 'ConfirmPreference', 'WhatIfPreference', 'VerbosePreference', 'WarningPreference'
			)
			$PreferenceVars | ForEach-Object {
				Set-Variable -Name $_ -Value $SessionState.PSVariable.GetValue($_) -Force
			}

			# Set default value for Include
			if(!$Include){
				$Include = '*'
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

			if($Type)
			{
				Write-Verbose "Trying to import file types $($Type -join ',')"
				[array]$FileTypeToImport = $Type | Sort-Object -Unique | ForEach-Object {$FileType.$_}
			}
			else
			{
				Write-Verbose 'Trying to import all supported file types'
				[array]$FileTypeToImport = $FileType.GetEnumerator() | ForEach-Object {$_.Value}
			}

			if(($FileTypeToImport).Extension -contains '.ps1' -and !$IsDotSourced.True)
			{
				Write-Warning "To import .PS1 scripts this function itself has to be dot-sourced! Example: $($IsDotSourced.Example)"
			}

			$FileTypeToImport |
				ForEach-Object {
					# We need to pass current file type to the filter function later
					$Private:currFT = $_
					Write-Verbose "Searching path '$($Path)' for file type '$(('*' + $_.Extension))'"
					Get-ChildItem -LiteralPath $Path -Filter ('*' + $_.Extension) -Recurse:$Recurse |
						# Process Include\Exclude parameters
						Where-Object {
							# No file extension and directory = module, else - file
							((!$Private:currFT.Extension -and $_.PSIsContainer) -or ($Private:currFT.Extension -and !$_.PSIsContainer)) -and
							# Include\Exclude filters with wildcard support
							$($_ |
								Where-Object {$tmp = $_.Basename ; ($Include | Where-Object {$tmp -like $_})} |
									Where-Object {$tmp = $_.Basename ; !($Exclude | Where-Object {$tmp -like $_})})
						} |
							ForEach-Object {
								Try
								{
									Write-Verbose "Trying to import: $_"
									. $Private:currFT.Command $_
									$Success = $true
									$ErrorMessage = $null
									Write-Verbose 'Success'
								}
								Catch
								{
									Write-Verbose 'Failure'
									$Success = $false
									$ErrorMessage = $_
								}

								$ret = @{
									Name = $_.Name
									Path = $_.FullName
									ErrorMessage = $ErrorMessage
									Success = $Success
								}

								Write-Verbose "Writing import status for '$_' to pipeline"
								New-Object -TypeName PSObject -Property $ret | Select-Object -Property Name, Path, ErrorMessage, Success
							}
				}
			} -ArgumentList (
				$PSBoundParameters.Path,
				$PSBoundParameters.Type,
				$PSBoundParameters.Exclude,
				$PSBoundParameters.Include,
				[bool]$PSBoundParameters.Recurse,
				# Execution in this scope is required to import PS1 files
				# To be executed in this scope (even when they passed to the module) scriptblocks have to be defined here
				# http://stackoverflow.com/questions/2193410/strange-behavior-with-powershell-scriptblock-variable-scope-and-modules-any-sug/27495377#27495377
				{. $args[0].FullName},
				{Add-Type -LiteralPath $args[0].FullName -ErrorAction SilentlyContinue},
				{
					if(Get-ChildItem -Path ($args[0].FullName + '\*') -Filter ($args[0].Basename + '.*'))
					{
						Write-Verbose "Folder '$($args[0].Basename)' looks like well-formed module"
						# https://msdn.microsoft.com/en-us/library/dd878350.aspx
						Import-Module -Name $args[0].FullName -ErrorAction SilentlyContinue
					}
					else
					{throw "Folder '$($args[0].Basename)' is not a well-formed module"}
				},
				(
					New-Module -OutVariable null -ReturnResult -AsCustomObject -ScriptBlock {
						# Is function dot-sourced?
						$MyCommand = (Get-PSCallStack)[2].Position.Text
						$MyCommandName = (Get-PSCallStack)[1].Command
		
						@{
							True = $MyCommand -match "\.\s+$MyCommandName\s+"
							Example = $MyCommand -replace "(^.*)($MyCommandName)(.*$)", '$1. $2$3'
						}
					}
				),
				$PSCmdlet.SessionState
			)
	}
}