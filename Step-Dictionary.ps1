<#
.Synopsis
	Recursively walk through each item in a dictionary and execute scriptblock against lowest level keys.

.Description
	Recursively walk through each item in a dictionary and execute scriptblock against lowest level keys.
	You can modify and remove lowest level keys while iterating over dictionary.

.Parameter Dictionary
	Dictionaty to iterate over. Must implement 'IDictionary' interface (hashtable, various .NET dictionaries).

.Parameter Scriptblock
	Scriptblock to execute. To access dictionary's key and it's value, two variables are exposed to the scriptblock:

	$Dictionary - current dictionary entry
	$key - key name

.Parameter Include
	Execute scriptblock only if lowest level key name matches specified wildcard. Accepts array of wildcards.

.Parameter Exclude
	Do not execute scriptblock if lowest level key name matches specified wildcard .Accepts array of wildcards.

.Parameter Depth
	Maximum nested node depth. Default value is [Int32]::MaxValue (2147483647).
	Keys with greater depth are not processed.

.Parameter CurrentDepth
	Internal parameter to support recursion, do not use it.

.Example
	Step-Dictionary -Dictionary $Dictionary -ScriptBlock {"${key}: $($Dictionary[$key])" | Write-Host}

	Print each lowest level key name and value.

.Example
	Step-Dictionary -Dictionary $Dictionary -ScriptBlock {$Dictionary[$key] = Get-Random}

	Set each lowest level key to random value.

.Example
	Step-Dictionary -Dictionary $Dictionary -ScriptBlock {$Dictionary.Remove($key)}

	Remove every lowest level key.

.Example

	#Search dictionary:

	$Dictionary = @{
		Alfa = @{
			Bravo = @{
				Charlie = 'FooBar'
				Delta = 'BarFoo'
				Echo = 'FarBoo'
			}
		}
	}

	# Search for lowest level keys with name 'Charlie' and output their values
	Step-Dictionary -Dictionary $Dictionary -ScriptBlock {$Dictionary[$key]} -Include 'Charlie'

	# Search for all lowest level keys except 'Charlie' and output their values
	Step-Dictionary -Dictionary $Dictionary -ScriptBlock {$Dictionary[$key]} -Exclude 'Charlie'
#>
function Step-Dictionary
{
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
		[ValidateScript({
			if($_.GetType().GetInterfaces().Name -contains 'IDictionary')
			{
				$true
			}
			else
			{
				throw 'The supplied object is not a dictionary.'
			}
		})]
		[ValidateNotNullOrEmpty()]
		$Dictionary,

		[Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
		[ValidateNotNullOrEmpty()]
		[scriptblock[]]$ScriptBlock,

		[Parameter(ValueFromPipelineByPropertyName = $true)]
		[ValidateNotNullOrEmpty()]
		[string[]]$Exclude,

		[Parameter(ValueFromPipelineByPropertyName = $true)]
		[ValidateNotNullOrEmpty()]
		[string[]]$Include = '*',

		[Parameter(ValueFromPipelineByPropertyName = $true)]
		[ValidateNotNullOrEmpty()]
		[int]$Depth = [Int32]::MaxValue,

		# Parameter below is to support recursion, do not use it
		[ValidateNotNullOrEmpty()]
		[int]$CurrentDepth = 0
	)

	Process
	{
		Write-Verbose "Current depth: $CurrentDepth"

		foreach($key in @($Dictionary.Keys))
		{
			Write-Verbose "Dictionary key: $key"

			if($Dictionary[$key].GetType().GetInterfaces().Name -contains 'IDictionary')
			{
				Write-Verbose "The '$key' contains dictionary"

				if(($CurrentDepth + 1) -ge $Depth)
				{
					Write-Verbose "Skipping, reached maximum depth: $Depth"
					continue
				}

				Write-Verbose "Recursively calling '$($PSCmdlet.MyInvocation.MyCommand.Name)'"

				$PSBoundParameters.Dictionary = $Dictionary[$key]
				$PSBoundParameters.CurrentDepth = $CurrentDepth + 1
				& $PSCmdlet.MyInvocation.MyCommand.Name @PSBoundParameters
			}
			else
			{
				if
				(
					# Include\Exclude filter with wildcard support
					$($key |
						Where-Object {$tmp = $_ ; ($Include | Where-Object {$tmp -like $_})} |
							Where-Object {$tmp = $_ ; !($Exclude | Where-Object {$tmp -like $_})})
				)
				{
					Write-Debug "Original '$key' value: $($Dictionary[$key])"

					# Execute scriptblocks
					foreach($sb in $ScriptBlock){
						. $sb
					}

					Write-Debug "New '$key' value: $($Dictionary[$key])"
				}
				else
				{
					Write-Verbose "Skipping key: $key"
				}
			}
		}
	}
}