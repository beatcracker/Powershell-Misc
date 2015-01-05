<#
.Synopsis
	Generate authors file for SVN to Git migration.
	Can map SVN authors to domain accounts and get full names and emails from Active Directiry.

.Description
	Generate authors file for one or more SVN repositories.
	Can map SVN authors to domain accounts and get full names and emails from Active Directiry
	Requires Subversion binaries and Get-SvnAuthor function:
	https://github.com/beatcracker/Powershell-Misc/blob/master/Get-SvnAuthor.ps1

.Notes
	Author: beatcracker (http://beatcracker.wordpress.com, https://github.com/beatcracker)
	License: Microsoft Public License (http://opensource.org/licenses/MS-PL)

.Component
	Requires Subversion binaries and Get-SvnAuthor function:
	https://github.com/beatcracker/Powershell-Misc/blob/master/Get-SvnAuthor.ps1

.Parameter Url
	This parameter is required.

	An array of strings representing URLs to the SVN repositories.

.Parameter Path
	This parameter is optional.

	A string representing path, where to create authors file.
	If not specified, new authors file will be created in the script directory.

.Parameter ShowOnly
	This parameter is optional.
	If this switch is specified, no file will be created and script will output collection of author names and emails.

.Parameter QueryActiveDirectory
	This parameter is optional.

	A switch indicating whether or not to query Active Directory for author full name and email.
	Supports the following formats for SVN author name: john, domain\john, john@domain

.Parameter User
	This parameter is optional.

	A string specifying username for SVN repository.

.Parameter Password
	This parameter is optional.

	A string specifying password for SVN repository.

.Parameter SvnPath
	This parameter is optional.

	A string specifying path to the svn.exe. Use it if Subversion binaries is not in your path variable, or you wish to use specific version.

.Example
	New-GitSvnAuthorsFile -Url 'http://svnserver/svn/project'

	Description
	-----------
	Create authors file for SVN repository http://svnserver/svn/project.
	New authors file will be created in the script directory.

.Example
	New-GitSvnAuthorsFile -Url 'http://svnserver/svn/project' -QueryActiveDirectory

	Description
	-----------
	Create authors file for SVN repository http://svnserver/svn/project.
	Map SVN authors to domain accounts and get full names and emails from Active Directiry.
	New authors file will be created in the script directory.

.Example
	New-GitSvnAuthorsFile -Url 'http://svnserver/svn/project' -ShowOnly

	Description
	-----------
	Create authors list for SVN repository http://svnserver/svn/project.
	Map SVN authors to domain accounts and get full names and emails from Active Directiry.
	No authors file will be created, instead script will return collection of objects.

.Example
	New-GitSvnAuthorsFile -Url 'http://svnserver/svn/project' -Path c:\authors.txt

	Description
	-----------
	Create authors file for SVN repository http://svnserver/svn/project.
	New authors file will be created as c:\authors.txt

.Example
	New-GitSvnAuthorsFile -Url 'http://svnserver/svn/project' -User john -Password doe

	Description
	-----------
	Create authors file for SVN repository http://svnserver/svn/project using username and password.
	New authors file will be created in the script directory.

.Example
	New-GitSvnAuthorsFile -Url 'http://svnserver/svn/project' -SvnPath 'C:\Program Files (x86)\VisualSVN Server\bin\svn.exe'

	Description
	-----------
	Create authors file for SVN repository http://svnserver/svn/project using custom svn.exe binary.
	New authors file will be created in the script directory.

.Example
	New-GitSvnAuthorsFile -Url 'http://svnserver/svn/project_1', 'http://svnserver/svn/project_2'

	Description
	-----------
	Create authors file for two SVN repositories: http://svnserver/svn/project_1 and http://svnserver/svn/project_2.
	New authors file will be created in the script directory.

.Example
	'http://svnserver/svn/project_1', 'http://svnserver/svn/project_2' | New-GitSvnAuthorsFile

	Description
	-----------
	Create authors file for two SVN repositories: http://svnserver/svn/project_1 and http://svnserver/svn/project_2.
	New authors file will be created in the script directory.
#>
[CmdletBinding()]
Param
(
	[Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Save')]
	[Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Show')]
	[string[]]$Url,

	[Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Save')]
	[ValidateScript({
		$ParentFolder = Split-Path -LiteralPath $_
		if(!(Test-Path -LiteralPath $ParentFolder  -PathType Container))
		{
			throw "Folder doesn't exist: $ParentFolder"
		}
		else
		{
			$true
		}
	})]
	[ValidateNotNullOrEmpty()]
	[string]$Path = (Join-Path -Path (Split-Path -Path $script:MyInvocation.MyCommand.Path) -ChildPath 'authors'),

	[Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Show')]
	[switch]$ShowOnly,

	[Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Save')]
	[Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Show')]
	[switch]$QueryActiveDirectory,

	[Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Save')]
	[Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Show')]
	[string]$User,

	[Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Save')]
	[Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Show')]
	[string]$Password,

	[Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Save')]
	[Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Show')]
	[string]$SvnPath
)

# Dotsource 'Get-SvnAuthor' function:
# https://github.com/beatcracker/Powershell-Misc/blob/master/Get-SvnAuthor.ps1
$ScriptDir = Split-Path $script:MyInvocation.MyCommand.Path
. (Join-Path -Path $ScriptDir -ChildPath 'Get-SvnAuthor.ps1')

# Strip extra parameters or splatting will fail
$Param = @{} + $PSBoundParameters
'ShowOnly', 'QueryActiveDirectory', 'Path' | ForEach-Object {$Param.Remove($_)}

# Get authors in SVN repo
$Names = Get-SvnAuthor @Param
[System.Collections.SortedList]$ret = @{}

# Exit, if no authors found
if(!$Names)
{
	Exit
}

# Find full name and email for every author
foreach($name in $Names)
{
	$Email = ''

	if($QueryActiveDirectory)
	{
		# Get account name from commit author name in any of the following formats:
		# john, domain\john, john@domain
		$Local:tmp = $name -split '(@|\\)'
		switch ($Local:tmp.Count)
		{
			1 { $SamAccountName = $Local:tmp[0] ; break }
			3 {
				if($Local:tmp[1] -eq '\')
				{
					[array]::Reverse($Local:tmp)
				}

				$SamAccountName = $Local:tmp[0]
				break
			}
			default {$SamAccountName = $null}
		}

		# Lookup account details
		if($SamAccountName)
		{
			$UserProps = ([adsisearcher]"(samaccountname=$SamAccountName)").FindOne().Properties

			if($UserProps)
			{
				Try
				{
					$Email = '{0} <{1}>' -f $UserProps.displayname[0], $UserProps.mail[0]
				}
				Catch{}
			}
		}
	}

	$ret += @{$name = $Email}
}

if($ShowOnly)
{
	$ret
}
else
{
	# Use System.IO.StreamWriter to write a file with Unix newlines.
	# It's also significally faster then Add\Set-Content Cmdlets.
	Try
	{
		#StreamWriter Constructor (String, Boolean, Encoding): http://msdn.microsoft.com/en-us/library/f5f5x7kt.aspx
		$StreamWriter = New-Object -TypeName System.IO.StreamWriter -ArgumentList $Path, $false,  ([System.Text.Encoding]::ASCII)
	}
	Catch
	{
		throw "Can't create file: $Path"
	}
	$StreamWriter.NewLine = "`n"

	foreach($item in $ret.GetEnumerator())
	{
		$Local:tmp = '{0} = {1}' -f $item.Key, $item.Value
		$StreamWriter.WriteLine($Local:tmp)
	}

	$StreamWriter.Flush()
	$StreamWriter.Close()
}