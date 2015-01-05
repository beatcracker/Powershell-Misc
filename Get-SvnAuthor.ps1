<#
.Synopsis
	Get list of unique commit authors in SVN repository.

.Description
	Get list of unique commit authors in one or more SVN repositories. Requires Subversion binaries.

.Parameter Url
	This parameter is required.

	An array of strings representing URLs to the SVN repositories.

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
	Get-SvnAuthor -Url 'http://svnserver/svn/project'

	Description
	-----------
	Get list of unique commit authors for SVN repository http://svnserver/svn/project

.Example
	Get-SvnAuthor -Url 'http://svnserver/svn/project' -User john -Password doe

	Description
	-----------
	Get list of unique commit authors for SVN repository http://svnserver/svn/project using username and password.

.Example
	Get-SvnAuthor -Url 'http://svnserver/svn/project' -SvnPath 'C:\Program Files (x86)\VisualSVN Server\bin\svn.exe'

	Description
	-----------
	Get list of unique commit authors for SVN repository http://svnserver/svn/project using custom svn.exe binary.

.Example
	Get-SvnAuthor -Url 'http://svnserver/svn/project_1', 'http://svnserver/svn/project_2'

	Description
	-----------
	Get list of unique commit authors for two SVN repositories: http://svnserver/svn/project_1 and http://svnserver/svn/project_2.

.Example
	'http://svnserver/svn/project_1', 'http://svnserver/svn/project_2' | Get-SvnAuthor

	Description
	-----------
	Get list of unique commit authors for two SVN repositories: http://svnserver/svn/project_1 and http://svnserver/svn/project_2.
#>
function Get-SvnAuthor
{
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
		[ValidateNotNullOrEmpty()]
		[string[]]$Url,

		[Parameter(ValueFromPipelineByPropertyName = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$User,

		[Parameter(ValueFromPipelineByPropertyName = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$Password,

		[ValidateScript({
			if(Test-Path -LiteralPath $_ -PathType Leaf)
			{
				$true
			}
			else
			{
				throw "$_ not found!"
			}
		})]
		[ValidateNotNullOrEmpty()]
		[string]$SvnPath = 'svn.exe'
	)

	Begin
	{
		if(!(Get-Command -Name $SvnPath -CommandType Application -ErrorAction SilentlyContinue))
		{
			throw "$SvnPath not found!"
		}
		$ret = @()
	}

	Process
	{
		$Url | ForEach-Object {
			$SvnCmd = @('log', $_, '--xml', '--quiet', '--non-interactive') + $(if($User){@('--username', $User)}) + $(if($Password){@('--password', $Password)})
			$SvnLog = &$SvnPath $SvnCmd *>&1

			if($LastExitCode)
			{
				Write-Error ($SvnLog | Out-String)
			}
			else
			{
				$ret += [xml]$SvnLog | ForEach-Object {$_.log.logentry.author}
			}
		}
	}

	End
	{
		$ret | Sort-Object -Unique
	}
}