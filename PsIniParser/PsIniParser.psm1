<# Not used in module. Uncomment and rename this file to *.ps1 if you want to run this file as script or dotsource it.

#region Load IniFileParser assembly
$Obj = Add-Type -Path '.\INIFileParser.dll' -PassThru

# For debugging
# $global:DebugPreference = 'Continue'
# $global:VerbosePreference = 'Continue'

#endregion

#>

#region Parsing\serializing functions

<#
    .Synopsis
    Parse INI from file or string.
    This function is later splitted and exported as Import-Ini ('File' parameter set) and ConvertFrom-Ini ('String' parameter set)
#>
function Parse-Ini
{
    [CmdletBinding(DefaultParameterSetName = 'String')]
    Param
    (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'File', Position = 0)]
        [ValidateNotNullOrEmpty()]
        [Alias('FullName')]
        [string[]]$Path,

        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'String', Position = 0)]
        [AllowEmptyString()]
        [string[]]$InputObject,

        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'File', Position = 1)]
        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'String', Position = 1)]
        [ValidateNotNullOrEmpty()]
        [string[]]$CommentStrings = @(';', '#'),

        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'File', Position = 2)]
        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'String', Position = 2)]
        [ValidateLength(1, 1)]
        [ValidateNotNullOrEmpty()]
        [string]$SectionStartChar = '[',

        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'File', Position = 3)]
        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'String', Position = 3)]
        [ValidateLength(1, 1)]
        [ValidateNotNullOrEmpty()]
        [string]$SectionEndChar = ']',

        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'File', Position = 4)]
        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'String', Position = 4)]
        [ValidateLength(1, 1)]
        [ValidateNotNullOrEmpty()]
        [string]$KeyValueAssigmentChar = '=',

        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'File', Position = 5)]
        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'String', Position = 5)]
        [switch]$OverrideDuplicateKeys,

        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'File', Position = 6)]
        [string]$Encoding,

        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'File', Position = 7)]
        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'String', Position = 6)]
        [switch]$AsObject

        <# Reserved for future use

        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$AssigmentSpacer = '',

        #>
    )

    Begin
    {
        Write-Verbose 'Creating new IniParser objects...'
        try
        {
            $FileIniParser = [IniParser.FileIniDataParser]::new()
            $StringIniParser = [IniParser.StringIniParser]::new()
        }
        catch
        {
            throw $_
        }

    }

    Process
    {
        try
        {
            if($Path)
            {
                Write-Verbose 'Using FileIniDataParser'
                $IniParser = $FileIniParser
            }
            else
            {
                Write-Verbose 'Using StringIniParser'
                $IniParser = $StringIniParser
            }
        }
        catch
        {
            throw $_
        }

        #region Configure IniParser

        Write-Verbose 'Mapping parameters to IniParser configuration'
        $CommentRegexTemplate = '^({0})(.*)'
        $IniParserCfg = @{
            CommentRegex = $CommentRegexTemplate -f (
                # Build a CommentRegex from the CommentString values
                (($CommentStrings | Select-Object -First ($CommentStrings.Count - 1) | ForEach-Object {[regex]::Escape($_) ; '|'}) + [regex]::Escape($CommentStrings[-1])) -join ''
            )
            SectionStartChar = $SectionStartChar
            SectionEndChar = $SectionEndChar
            CaseInsensitive = $true
            KeyValueAssigmentChar = $KeyValueAssigmentChar
            AllowKeysWithoutSection = $true
            AllowDuplicateKeys = $true
            OverrideDuplicateKeys = $OverrideDuplicateKeys
            AllowDuplicateSections = $true
            SkipInvalidLines = $true

            <# Not used

            SectionRegex = '^(\s*?)\[{1}\s*[_\{\}\#\+\;\%\(\)\=\?\&\$\,\:\/\.\-\w\d\s\\\~]+\s*](\s*?)'
            CommentChar = ';'
            CommentString = ';'
            AssigmentSpacer = $AssigmentSpacer
            ThrowExceptionsOnError = $true

            #>
        }

        Write-Verbose 'Setting new IniParser configuration'
        Set-IniParserConfiguration -InputObject $IniParser.Parser.Configuration -Configuration $IniParserCfg

        #endregion

        #region Process ini

        if($Path)
        {
            $InputObject = $Path
        }

        # Process multiple files via pipeline or array of strings
        $InputObject |
        ForEach-Object {
            try
            {
                if($Path)
                {
                    Write-Verbose 'Parsing INI from file'
                    $IniObject = $IniParser.ReadFile($_, (New-Encoding -Encoding $Encoding))
                }
                else
                {
                    Write-Verbose 'Parsing INI from string'
                    $IniObject = $IniParser.ParseString($_)
                }
            }
            catch
            {
                Write-Error $_
                return
            }

            if($AsObject)
            {
                Write-Verbose 'Writing output as object to the pipeline'
                $IniObject
            }
            else
            {
                # 'Ordered' hashtable, PS 2.0 compatible
                $IniHashtable = [System.Collections.Specialized.OrderedDictionary]::new()

                if($IniObject.Global.Count)
                {
                    Write-Verbose 'Processing keys without section ("Global" section)'
                    $IniObject.Global.GetEnumerator() |
                    ForEach-Object {
                        Write-Debug "Adding key: $($_.KeyName), value: $($_.Value)"
                        $IniHashtable.Add($_.KeyName, $_.Value)
                    }
                }

                if($IniObject.Sections.Count)
                {
                    Write-Verbose 'Processing keys within sections'
                    $IniObject.Sections.GetEnumerator() |
                    ForEach-Object {

                        # 'Ordered' hashtable, PS 2.0 compatible
                        $IniSectionHashtable = [System.Collections.Specialized.OrderedDictionary]::new()
                        $_.Keys.GetEnumerator() |
                        ForEach-Object {
                            $IniSectionHashtable.Add($_.KeyName, $_.Value)
                        }

                        Write-Debug "Section $($_.SectionName), adding key(s): $(($IniSectionHashtable | Format-Table -AutoSize | Out-String).TrimEnd())"
                        $IniHashtable.Add($_.SectionName, $IniSectionHashtable)
                    }
                }

                Write-Verbose 'Writing output as ordered hashtable to the pipeline'
                $IniHashtable
            }
        }
        #endregion
    }
}


<#
    .Synopsis
    Serialize INI to file or string.
    This function is later splitted and exported as Export-Ini ('File' parameter set) and ConvertTo-Ini ('Object' parameter set)
#>
function Serialize-Ini # Export-Ini and ConvertTo-Ini
{
    [CmdletBinding(DefaultParameterSetName = 'Object')]
    Param
    (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Object', Position = 0)]
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'File', Position = 0)]
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Merge', Position = 0)]
        [ValidateNotNullOrEmpty()]
        [System.Object[]]$InputObject,

        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'File', Position = 1)]
        [ValidateNotNullOrEmpty()]
        [string]$Path,

        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Object', Position = 1)]
        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'File', Position = 2)]
        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Merge', Position = 1)]
        [ValidateNotNullOrEmpty()]
        [string]$CommentString = ';',

        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Object', Position = 2)]
        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'File', Position = 3)]
        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Merge', Position = 2)]
        [ValidateLength(1, 1)]
        [ValidateNotNullOrEmpty()]
        [string]$SectionStartChar = '[',

        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Object', Position = 3)]
        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'File', Position = 4)]
        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Merge', Position = 3)]
        [ValidateLength(1, 1)]
        [ValidateNotNullOrEmpty()]
        [string]$SectionEndChar = ']',

        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Object', Position = 4)]
        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'File', Position = 5)]
        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Merge', Position = 4)]
        [ValidateLength(1, 1)]
        [ValidateNotNullOrEmpty()]
        [string]$KeyValueAssigmentChar = '=',

        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Object', Position = 5)]
        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Merge', Position = 5)]
        [switch]$Merge,

        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'File', Position = 6)]
        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Merge', Position = 6)]
        [switch]$OverrideDuplicateKeys,

        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'File', Position = 7)]
        [string]$Encoding,

        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'File', Position = 8)]
        [switch]$NoClobber

        <# Reserved for future use

        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Object')]
        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'File')]
        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Merge')]
        [ValidateNotNullOrEmpty()]
        [string]$AssigmentSpacer = '',

        #>
    )

    Begin
    {
        Write-Verbose 'Creating new IniParser objects...'
        try
        {
            $IniParser = [IniParser.FileIniDataParser]::new()
            $IniData = [IniParser.Model.IniDataCaseInsensitive]::new()
        }
        catch
        {
            throw $_
        }

    }

    Process
    {
        if($Path)
        {
            if($NoClobber -and (Test-Path -LiteralPath $Path -PathType Leaf))
            {
                Write-Error "File already exist: $Path"
                return
            }

            Write-Verbose 'Forcing Merge for the file output'
            $Merge = $true
        }

        Write-Verbose 'Mapping parameters to IniParser configuration'
        $IniParserCfg = @{
            SectionStartChar = $SectionStartChar
            SectionEndChar = $SectionEndChar
            CaseInsensitive = $true
            CommentString = $CommentString
            KeyValueAssigmentChar = $KeyValueAssigmentChar
            AssigmentSpacer = ''

            <# Not used

            CommentRegex = 
            SectionRegex = '^(\s*?)\[{1}\s*[_\{\}\#\+\;\%\(\)\=\?\&\$\,\:\/\.\-\w\d\s\\\~]+\s*](\s*?)'
            CommentChar = ';'
            AllowKeysWithoutSection = $true
            AllowDuplicateKeys = $true
            OverrideDuplicateKeys = $OverrideDuplicateKeys
            ThrowExceptionsOnError = $true
            AllowDuplicateSections = $true
            SkipInvalidLines = $true

            #>
        }

        Write-Verbose 'Setting new IniParser configuration'
        Set-IniParserConfiguration -InputObject $IniData.Configuration -Configuration $IniParserCfg

        $InputObject |
        ForEach-Object {

            if(Is-Hashtable -InputObject $_)
            {
                Write-Verbose 'Input is hashtable, converting to IniData'
                $currIniData = ConvertHashtable-ToIni -Hashtable $_
            }
            elseif(Is-IniData -InputObject $_)
            {
                Write-Verbose 'Input is IniData, skipping conversion'
                $currIniData = $_
            }
            else
            {
                Write-Error 'Input is not hashtable or IniData'
                return
            }

            if($OverrideDuplicateKeys)
            {
                Write-Verbose 'Processing data from the current object, overriding duplicate keys if any'
                $IniData.Merge($currIniData)
            }
            else
            {
                Write-Verbose 'Processing data from the current object'

                # This doesn't work: $IniData.Merge($currIniData.Clone().Merge($IniData.Clone()))

                $cloneCID = $currIniData.Clone()
                $cloneID = $IniData.Clone()
                $cloneCID.Merge($cloneID)

                $IniData.Merge($cloneCID)
            }

            if(!$Merge) # means we're also not writing a file
            {
                Write-Verbose 'Writing output as string to the pipeline'
                $IniData.ToString()

                Write-Verbose 'Resetting IniParser object for the next loop'
                try
                {
                    $IniData.Global.RemoveAllKeys()
                    $IniData.Sections.Clear()
                }
                catch
                {
                    throw $_
                }
            }
        }
    }

    End
    {
        if($Merge)
        {
            if($Path)
            {
                Write-Verbose "Writing merged output as file: $Path"
                $IniParser.WriteFile($Path, $IniData, (New-Encoding -Encoding $Encoding))
            }
            else
            {
                Write-Verbose 'Writing merged output as string to the pipeline'
                $IniData.ToString()
            }
        }
    }
}
#endregion


#region Helper functions
<#
    .Synopsis
    Helper function for Serialize-Ini, converts hashtable to [IniParaser.Model.IniData] object.

    .Parameter Hashtable
    Hashtable to convert.

    .Parameter IniData
    Optional [IniParaser.Model.IniData] object, that will hold converted hashtable.
    If not specified, new [IniParser.Model.IniDataCaseInsensitive] will be created and used.

    .Parameter Depth
    Optional, specifies maximum depth of recursion. Default is 2, which corresponds to INI file structure:
    first level are keys without section, second level are keys within section. Values more then 2 are currently not supported.

    .Parameter CurrentDepth
    Optional, incrementend for each recursive call. Used internally in function to support recursion.

    .Parameter Section
    Optional, default section for non-nested key-value pairs in hashtable. Used internally in function to support recursion.
    Default value is 'Global', and then in each recursive call it's set to the name of the hashtable key that holds nested hashtable.
    Used as section name in [IniParaser.Model.IniData] object.
#>
function ConvertHashtable-ToIni
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        $Hashtable,

        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        $IniData = [IniParser.Model.IniDataCaseInsensitive]::new(),

        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [int]$Depth = 2,

        # Parameters below are to support recursion
        [ValidateNotNullOrEmpty()]
        [int]$CurrentDepth = 1,

        [ValidateNotNullOrEmpty()]
        [string]$Section = 'Global'
    )

    Begin
    {
        # Scriptblocks for common messages
        $AddKey = {"Adding key: '$($_.Key)' to section: '$Section', with value: '$($_.Value)'"}
        $AddKeyFailed = {"Failed to add key: '$($_.Key)', to section: '$Section'"}
    }

    Process
    {
        Write-Debug "Current depth: $CurrentDepth"

        if($CurrentDepth -gt $Depth)
        {
            Write-Debug "Exiting, reached maximum: $Depth"
            return
        }
        $Hashtable.GetEnumerator() |
        ForEach-Object {
            if($_.Value | Is-Hashtable)
            {
                Write-Debug "Key contains hashtable: $($_.Key)"
                Write-Debug "Recursively calling $($PSCmdlet.MyInvocation.MyCommand.Name)"
                $PSBoundParameters.IniData = $IniData
                $PSBoundParameters.CurrentDepth = $CurrentDepth + 1
                $PSBoundParameters.Section = $_.Key
                $PSBoundParameters.Hashtable = $_.Value
                & $PSCmdlet.MyInvocation.MyCommand.Name @PSBoundParameters
            }
            else
            {
                if($Section -eq 'Global')
                {
                    Write-Debug $(. $AddKey)
                    if(!$IniData.Global.AddKey($_.Key, $_.Value))
                    {
                        throw $(. $AddKeyFailed)
                    }
                }
                else
                {
                    if(!$IniData.Sections.ContainsSection($Section))
                    {
                        Write-Debug "Adding ini section: $Section"
                        if(!$IniData.Sections.AddSection($Section))
                        {
                            throw "Failed to add section: $Section"
                        }
                    }

                    Write-Debug $(. $AddKey)
                    if(!$IniData.Sections[$Section].AddKey($_.Key, $_.Value))
                    {
                        throw $(. $AddKeyFailed)
                    }
                }                    
            }
        }

        # Do not return if runnnig recursively
        if($CurrentDepth -eq 1)
        {
            Write-Debug 'Returning result'
            $IniData
        }
    }
}

<#
    .Synopsis
    Checks if value is Hashtable or Ordered hashtable ([System.Collections.Specialized.OrderedDictionary]) and returns True or False accordingly.

    .Parameter InputObject
    Object to check.
#>
function Is-Hashtable
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        $InputObject
    )

    Process
    {
        if
        (
            ($InputObject -is [hashtable]) -or
            ($InputObject -is [System.Collections.Specialized.OrderedDictionary])
        )
        {
            Write-Debug 'Object is Hashtable or OrderedDictionary'
            $true
        }
        else
        {
            Write-Debug 'Object is not Hashtable or OrderedDictionary'
            $false
        }
    }
}

<#
    .Synopsis
    Checks if value is  [IniParser.Model.IniData] or [IniParser.Model.IniDataCaseInsensitive] and returns True or False accordingly.

    .Parameter InputObject
    Object to check.
#>
function Is-IniData
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        $InputObject
    )

    Process
    {
        if
        (
            ($InputObject -is [IniParser.Model.IniData]) -or
            ($InputObject -is [IniParser.Model.IniDataCaseInsensitive])
        )
        {
            Write-Debug 'Object is IniData'
            $true
        }
        else
        {
            Write-Debug 'Object is not IniData'
            $false
        }
    }
}

function Set-IniParserConfiguration
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [ValidateNotNullOrEmpty()]
        $InputObject,

        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        $Configuration
    )

    Process
    {
        $Configuration.GetEnumerator() |
        ForEach-Object {
            Write-Debug "$($_.Key) = $($_.Value)"
            $InputObject."$($_.Key)" = $_.Value
        }
    }
}

<#
    .Synopsis
    Try to create user-specified encoding, return default (OS ANSI codepage) if failed.

    .Parameter Encoding
    Optional, .Net encoding  name. If not specified, function returns OS ANSI codepage. Use [System.Text.Encoding]::GetEncodings() to get the list of supported encodings.
#>
function New-Encoding
{
    [CmdletBinding()]
    Param
    (
        [Parameter(ValueFromPipeline = $true)]
        [string]$Encoding
    )

    Process
    {
        $ret = [System.Text.Encoding]::Default

        if($Encoding)
        {
            Write-Debug 'Trying to get user-specified encoding'
            try
            {
                $ret = [System.Text.Encoding]::GetEncoding($Encoding)
            }
            catch
            {
                Write-Error "Not valid encoding: '$Encoding', using default"
            }
        }

        $ret
    }
}

<#
    .Synopsis
    Creates comment based help from hashtable.

    .Description
    This function converts hashtable to comment based help. It used internally when creating proxy functions.

    Limitations:

    You can't specifiy same keys in Per-Function and Common hashtables.
    No attempt to merge them is made, so it will result in error.

    Hashtable format:

    $HelpData = @{
        # Help data for function named 'New-BlackHole'
        # Repeat for every function that needs help generated.
        'New-BlackHole' = @{
            # Nested hashtable that holds help for function parameters
            Parameter = @{
                # Each of the nested keys will be converted to comment block, like this:
                #
                # .Parameter ParameterName
                # Description of ParameterName

                ParameterName = 'Description of ParameterName'
            }

            # Each of the keys will be converted to comment block, like this:
            #
            # .Key
            # Value

            # .Synopsis
            # Function synopis
            Synopsis = 'Function synopis'

            # .Description
            # Function description
            Description = 'Function description'
        }

        # This key holds common items, that will be added to any function
        Common = @{
            # Each of the keys will be converted to comment block, like this:
            #
            # .Key
            # Value
            #
            # If value is an array, multiple blocks will be generated.
            # It can be used for ex. for .Link items.

            # .Link
            # http://example.net
            Link = 'http://example.com', 'http://example.net'

            # .Note
            # Some note
            Note = 'Some note'

            # .Parameter ParameterName
            # Description of ParameterName
            CommonParameterName = 'Description of CommonParameterName'
        }
    }

    .Parameter HelpData
    Hashtable that holds help data. For hashtable format, see function description.

    .Parameter CommandName
    Function name. Corresponding hashtable from HelpData will be used to generate help.
#>
function New-CommentBasedHelp
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$CommandName,

        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [hashtable]$HelpData
    )

    Process
    {
        if($Command = Get-Command -Name $CommandName -ErrorAction SilentlyContinue)
        {
            Write-Debug "Getting command metadata for: $CommandName"
            $CommandMetaData = [System.Management.Automation.CommandMetaData]::new($Command)
        }
        else
        {
            Write-Debug "Command not found: $CommandName, so redundant parameters wouldn't be skipped"
        }

        Write-Debug 'Generating help topics'
        $Topics = ($HelpData[$CommandName].Keys + $HelpData['Common'].Keys) | Sort-Object -Unique

        $NewLine = [System.Environment]::NewLine
        $CommandHelp = @()
        foreach($topic in $Topics){

            $HelpTopic = $HelpData[$CommandName][$topic]

            if($HelpData['Common'][$topic])
            {
                $HelpTopic += $HelpData['Common'][$topic]
            }
                     
            if($HelpTopic -is [hashtable])
            {
                $HelpTopic.GetEnumerator() |
                ForEach-Object {
                    if
                    (
                        $topic -eq 'Parameter' -and
                        ($Command -and !$CommandMetaData.Parameters.ContainsKey($_.Key))
                    )
                    {
                        Write-Debug "Skipping topic for Parameter: $($_.Key)"
                        return
                    }

                    Write-Debug "Processing topic: $('{0} {1}' -f $topic, $_.Key)"
                    $CommandHelp += '.{0} {1}{3}{2}' -f $topic, $_.Key, $_.Value, $NewLine

                }
            }
            else
            {
                Write-Debug "Processing topic: $topic"
                $HelpTopic |
                ForEach-Object {
                    $CommandHelp += '.{0}{2}{1}' -f $topic, $_, $NewLine
                }
            }
        }

        return ($CommandHelp -join $NewLine)
    }
}

<#
    .Synopsis
    Creates new proxy function by removing ParameterSets from base function.

    .Description
    Creates new proxy function by removing ParameterSets from base function.
    It used internally for creating proxy functions (Export-Ini, ConvertFrom-Ini and Import-Ini, ConvertTo-Ini) from Parse-Ini and Serialize-Ini functions.

    .Parameter CommandName
    Name of the base function, that will be used for proxy function creation.

    .Parameter ProxyCommandName
    Name of the new proxy function.

    .Parameter ExcludeParameterSet
    Array of parameter set names to exculde from base function, when creating new proxy function.

    .Parameter AsScriptblock
    Optional, if this switch is specified, new proxy function will be returned as scriptblock.
    By default, new proxy function is returned as string.

    .Parameter HelpData
    Optional, hashtable from which is help for new proxy function will be generated.
    For details see help for New-CommentBasedHelp function.

    If not specified, no help for the new proxy function will be created.
#>
function New-ProxyCommandFromParameterSet
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateScript({
            Get-Command -Name $_ -CommandType Function, Cmdlet
        })]
        [ValidateNotNullOrEmpty()]
        [string]$CommandName,

        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$ProxyCommandName,

        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [string[]]$ExcludeParameterSet,

        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [switch]$AsScriptblock,

        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [hashtable]$HelpData
    )

    Process
    {
        Write-Debug "Getting command metadata for: $CommandName"
        $BaseCommandMetaData = [System.Management.Automation.CommandMetaData]::new((Get-Command -Name $CommandName))

        Write-Debug "Deep cloning command metadata for proxy command: $ProxyCommandName"
        $ProxyCommandMetaData = [System.Management.Automation.CommandMetadata]::new($BaseCommandMetaData)

        Write-Debug 'Enumerating parameters'
        ($ProxyCommandMetaData.Parameters | Select-Object -ExpandProperty Keys) |
        ForEach-Object {
            Write-Debug "Processing parameter: $_"
            $Parameter = $_
            $ExcludeParameterSet |
            ForEach-Object {
                if($ProxyCommandMetaData.Parameters[$Parameter].ParameterSets.ContainsKey($_))
                {
                    Write-Debug "Parameter contains excluded ParameterSet: $_"
                    if($ProxyCommandMetaData.Parameters[$Parameter].ParameterSets.Count -eq 1)
                    {
                        Write-Debug 'Removing parameter because it doesn''t contains any other ParameterSets'
                        if(!$ProxyCommandMetaData.Parameters.Remove($Parameter))
                        {
                            throw "Failed to remove parameter: $Parameter"
                        }
                    }
                    else
                    {
                        Write-Debug "Removing ParameterSet: $_"
                        if(!$ProxyCommandMetaData.Parameters[$Parameter].ParameterSets.Remove($_))
                        {
                            throw "Failed to remove parameter set: $_, for parameter: $_"
                        }
                    }
                }
            }
        }

        $NewLine = [System.Environment]::NewLine

        if($HelpData)
        {
            $ProxyCommandHelp = New-CommentBasedHelp -CommandName $ProxyCommandName -HelpData $HelpData
        }

        if($ProxyCommandHelp)
        {
            $ProxyCommandBody = [System.Management.Automation.ProxyCommand]::Create($ProxyCommandMetaData, $ProxyCommandHelp)
        }
        else
        {
            $ProxyCommandBody = [System.Management.Automation.ProxyCommand]::Create($ProxyCommandMetaData)
        }

        Write-Debug 'Creating new proxy function'
        #$ProxyCommandFunctionBody = "function $ProxyCommandName {$NewLine$ProxyCommandBody$NewLine}"
        $ProxyCommandFunctionBody = 'function {0} {{{2}{1}{2}}}' -f $ProxyCommandName, $ProxyCommandBody, $NewLine

        if($AsScriptblock)
        {
            Write-Debug 'Returning new proxy function as scriptblock'
            [ScriptBlock]::Create($ProxyCommandFunctionBody)
        }
        else
        {
            Write-Debug 'Returning new proxy function as string'
            $ProxyCommandFunctionBody
        }
    }
}
#endregion

#region Help

$HelpData = @{
    'Import-Ini' = @{
        Parameter = @{
            Path = 'Specifies the path to the INI file to import. You can also pipe a path to Import-Ini.'
        }
        Synopsis = 'Creates hashtables from the items in a INI file.'
        Description = @'
The INI file format is an informal standard for configuration files for some platforms or software. INI files are simple text files with a basic structure composed of sections, properties, and values. The Import-Ini cmdlet creates ordered hashtables from the items in INI file, preserving order of the keys and sections. Each section in the INI file becomes a key of the hashtable and the propertie-value pairs in the section become nested hashtable with the corresponding key-value pairs. Property-value pairs without section in the INI file become topmost keys and values in the hashtable. Import-Ini should work on almost any INI file, including files that are generated by the Export-Ini cmdlet.

You can use the parameters of the Import-Ini cmdlet to specify various parsing options (INI files are not standardized), or direct Import-Ini to use specific encoding while reading a file.
    
You can also use the ConvertTo-Ini and ConvertFrom-Ini cmdlets to convert hashtables to INI strings (and back). These cmdlets are the same as the Export-Ini and Import-Ini cmdlets, except that they do not deal with files.
'@
        Example = @'
Import-Ini -Path C:\Windows\System.ini
 
    Description
    -----------
    Import single file, return hashtable.
'@,
@'
Import-Ini -Path 'C:\Windows\System.ini', 'C:\Windows\Win.ini'
 
    Description
    -----------
    Import multiple files, return hashtables.
'@,
@'
C:\Windows\System.ini | Import-Ini
 
    Description
    -----------
    Import single file specified via pipeline input, return hashtable.
'@,
@'
'C:\Windows\System.ini', 'C:\Windows\Win.ini' | Import-Ini
 
    Description
    -----------
    Import multiple files specified via pipeline input, return hashtables.
'@,
@'
Import-Ini -Path C:\Windows\System.ini -CommentStrings '%' -SectionStartChar '{' -SectionEndChar '}' -KeyValueAssigmentChar '@'
 
    Description
    -----------
    Import single file with non-standard structure, return hashtable.

    File example:

    {Section}
    Key@Value
    %Comment
'@,
@'
$IniData = Import-Ini -Path C:\Windows\System.ini -AsObject
 
    Description
    -----------
    Import single file, return [IniParser.IniData] object.


    • View keys without sections:

        $IniData.Global

    • Get value of the key 'MyGlobalKey':

        $IniData.Global['MyGlobalKey']

    • Set value of the key 'MyGlobalKey':

        $IniData.Global['MyGlobalKey'] = 'NewValue'


    • View INI sections:

        $IniData.Sections


    • View leading\trailing comments for section:

        $IniData.Sections.GetSectionData('Section').LeadingComments
        $IniData.Sections.GetSectionData('Section').TrailingComments

    • View all comments for section:

        $IniData.Sections.GetSectionData('Section').Comments


    • Get value of the key 'woafont' in section '386Enh'

        $IniData.Sections['386Enh']['woafont']

    • Set value of the key 'woafont' in section '386Enh'

        $IniData.Sections['386Enh']['woafont'] = 'NewValue'


    • Convert modified object to file:

        $IniData | Export-Ini -Path C:\Windows\System.ini
'@
    }
    'ConvertFrom-Ini' = @{
        Parameter = @{
            InputObject = 'Specifies the INI strings to be converted to hashtables. Enter a variable that contains the INI strings or type a command or expression that gets the INI strings. You can also pipe the INI strings to ConvertFrom-Ini.'
        }
        Synopsis = 'Creates hashtables from sections, properties, and values in INI strings.'
        Description = @'
The ConvertFrom-Ini cmdlet creates ordered hashtables from the INI-formatted strings that are generated by the ConvertTo-Ini cmdlet (or provided by user), preserving order of the keys and sections. Each section in the INI file becomes a key of the hashtable and the propertie-value pairs in the section become nested hashtable with the corresponding key-value pairs. Property-value pairs without section in the INI file become topmost keys and values in the hashtable.

You can use the parameters of the ConvertFrom-Ini cmdlet to specify various parsing options (INI files are not standardized).
    
You can also use the Export-Ini and Import-Ini cmdlets to convert hashtables to INI strings in a file (and back). These cmdlets are the same as the ConvertTo-Ini and ConvertFrom-Ini cmdlets, except that they save the INI strings in a file.
'@
        Example = @'
ConvertFrom-Ini -InputObject $IniString
 
    Exemplary variable
    ------------------

$IniString = @'
[Section]
Key=Value
''@

    Description
    -----------
    Convert single INI string, return hashtable.
'@,
@'
ConvertFrom-Ini -InputObject $IniStringA, $IniStringB
 
    Exemplary variables
    -------------------

$IniStringA = @'
[SectionA]
KeyA=ValueA
''@

$IniStringB = @'
[SectionB]
KeyB=ValueB
''@

    Description
    -----------
    Convert multiple INI strings, return hashtables.
'@,
@'
$IniString | ConvertFrom-Ini
 
    Exemplary variable
    ------------------

$IniString = @'
[Section]
Key=Value
''@

    Description
    -----------
    Convert single INI string specified via pipeline input, return hashtable.
'@,
@'
$IniStringA, $IniStringB | ConvertFrom-Ini
 
    Exemplary variables
    -------------------

$IniStringA = @'
[SectionA]
KeyA=ValueA
''@

$IniStringB = @'
[SectionB]
KeyB=ValueB
''@

    Description
    -----------
    Convert multiple INI strings specified via pipeline input, return hashtables.
'@,
@'
ConvertFrom-Ini -InputObject $IniString -CommentStrings '%' -SectionStartChar '{' -SectionEndChar '}' -KeyValueAssigmentChar '@'
 
    Exemplary variable
    ------------------

$IniString = @'
{Section}
Key@Value
%Comment
''@

    Description
    -----------
    Convert single INI string with non-standard structure, return hashtable.
'@,
@'
$IniData = ConvertFrom-Ini -InputObject $IniString -AsObject
 
    Exemplary variable
    ------------------

$IniString = @'
;Leading comment
[Section]
Key=Value
;Trailing comment
''@

    Description
    -----------
    Convert single INI string, return [IniParser.IniData] object.


    • View keys without sections:

        $IniData.Global

    • Get value of the key 'MyGlobalKey':

        $IniData.Global['MyGlobalKey']

    • Set value of the key 'MyGlobalKey':

        $IniData.Global['MyGlobalKey'] = 'NewValue'


    • View INI sections:

        $IniData.Sections


    • View leading\trailing comments for section:

        $IniData.Sections.GetSectionData('Section').LeadingComments
        $IniData.Sections.GetSectionData('Section').TrailingComments

    • View all comments for section:

        $IniData.Sections.GetSectionData('Section').Comments


    • Get value of the key 'woafont' in section '386Enh'

        $IniData.Sections['386Enh']['woafont']

    • Set value of the key 'woafont' in section '386Enh'

        $IniData.Sections['386Enh']['woafont'] = 'NewValue'


    • Convert modified object to INI string:

        $IniData | ConvertTo-Ini
'@
    }
    'Export-Ini' = @{
        Parameter = @{
            Path = 'Specifies the path to the INI file to export. You can also pipe a path to Export-Ini.'
        }
        Description = @'
The INI file format is an informal standard for configuration files for some platforms or software. INI files are simple text files with a basic structure composed of sections, properties, and values. The Export-Ini cmdlet creates a INI file from the hashtable that you submit, preserving order of the keys and sections. Each key of the hashtable becomes section in the INI file and it's nested key-value pairs become corresponding propertie-value pairs in the section. Topmost keys in the hashtable become property-value pairs without section in the INI file.

You can also use the ConvertTo-Ini and ConvertFrom-Ini cmdlets to convert hashtables to INI strings (and back). These cmdlets are the same as the Export-Ini and Import-Ini cmdlets, except that they do not deal with files.
'@
        Example = @'
Export-Ini -InputObject $Hashtable -Path '.\My.ini'
 
    Exemplary variable
    ------------------

    $Hashtable = @{
        Section = @{
            Key = 'Value'
        }
    }

    Description
    -----------
    Export single hashtable to file 'My.ini'
'@,
@'
Export-Ini -InputObject $Hashtable -Path '.\My.ini' -NoClobber
 
    Exemplary variable
    ------------------

    $Hashtable = @{
    Section = @{
    Key = 'Value'
    }
    }

    Description
    -----------
    Export single hashtable to file 'My.ini', do not overwrite file if it exists.
'@,
@'
Export-Ini -InputObject $Hashtable -Path '.\My.ini' -Encoding 'UTF-8'
 
    Exemplary variable
    ------------------

    $Hashtable = @{
    Section = @{
    Key = 'Value'
    }
    }

    Description
    -----------
    Export single hashtable to file 'My.ini', use UTF-8 encoding.
'@,
@'
Export-Ini -InputObject $HashtableA, $HashtableB -Path '.\My.ini'
 
    Exemplary variables
    -------------------

    $HashtableA = @{
    SectionA = @{
    KeyA = 'ValueA'
    }
    }

    $HashtableB = @{
    SectionB = @{
    KeyB = 'ValueB'
    }
    }

    Description
    -----------
    Export multiple hashtables to file 'My.ini'
'@,
@'
$HashtableA, $HashtableB | Export-Ini -Path '.\My.ini'
 
    Exemplary variables
    -------------------

    $HashtableA = @{
    SectionA = @{
    KeyA = 'ValueA'
    }
    }

    $HashtableB = @{
    SectionB = @{
    KeyB = 'ValueB'
    }
    }

    Description
    -----------
    Export multiple hashtables specified via pipeline input to file 'My.ini'
'@,
@'
Export-Ini -InputObject $HashtableA, $HashtableB -Path '.\My.ini'
 
    Exemplary variables
    -------------------

    $HashtableA = @{
    Section = @{
    Key = 'ValueA'
    }
    }

    $HashtableB = @{
    Section = @{
    Key = 'ValueB'
    }
    }

    Description
    -----------
    Export multiple hashtables specified via pipeline input to file 'My.ini'. Duplicate keys are not overriden.

    Result:

    [Section]
    Key=ValueA
'@,
@'
Export-Ini -InputObject $Hashtable  -SectionStartChar '{' -SectionEndChar '}' -KeyValueAssigmentChar '@' -Path '.\My.ini'
 
    Exemplary variable
    ------------------

    $Hashtable = @{
    Section = @{
    Key = 'Value'
    }
    }

    Description
    -----------
    Export single hashtable to file 'My.ini' with non-standard structure.

    Resulting file:

    {Section}
    Key@Value
'@,
@'
Export-Ini -InputObject $HashtableA, $HashtableB -Path '.\My.ini' -OverrideDuplicateKeys
 
    Exemplary variables
    -------------------

    $HashtableA = @{
    Section = @{
    Key = 'ValueA'
    }
    }

    $HashtableB = @{
    Section = @{
    Key = 'ValueB'
    }
    }

    Description
    -----------
    Export multiple hashtables specified via pipeline input to file 'My.ini'. Duplicate keys are overriden.

    Resulting file:

    [Section]
    Key=ValueB
'@,
@'
Export-Ini -InputObject $HashtableA, $HashtableB, $IniData -Path '.\My.ini'
 
    Exemplary variables
    -------------------

    $HashtableA = @{
    SectionA = @{
    KeyA = 'ValueA'
    }
    }

    $HashtableB = @{
    SectionB = @{
    KeyB = 'ValueB'
    }
    }

    $IniString = @'
    [SectionC]
    KeyC=ValueC
    ''@

    $IniData = $IniString | ConvertFrom-Ini -AsObject

    Description
    -----------
    You can mix and match hashtables and [IniParser.IniData] objects. This example will export them to file 'My.ini.

    Resulting file:

    [SectionA]
    KeyA=ValueA

    [SectionB]
    KeyB=ValueB

    [SectionC]
    KeyC=ValueC
'@
    }
    'ConvertTo-Ini' = @{
        Parameter = @{
            InputObject = 'Specifies the hashtable to be converted to INI strings. Keys and values converted to property-value pairs, if hashtable is nested, parent key is treated as section name. Hashtable could be nested up to two levels. Enter a variable that contains the hashtable or type a command or expression that gets the hashtable. You can also pipe the hashtable to ConvertTo-Ini.'
        }
        Synopsis = 'Creates INI strings from hashtables from sections, properties, and values in INI-strings.'
        Description = @'
The ConvertTo-Ini cmdlet creates INI-strings from the hashtable that is generated by the ConvertFrom-Ini cmdlet (or provided by user), preserving order of the keys and sections. Each key of the hashtable becomes section in the INI-string and it's nested key-value pairs become corresponding propertie-value pairs in the section. Topmost keys in the hashtable become property-value pairs without section in the INI strings.

You can use the parameters of the ConvertFrom-Ini cmdlet to specify various parsing options (INI files are not standardized).
    
You can also use the Import-Ini and Export-Ini cmdlets to convert INI strings in a file to hashtables (and back). These cmdlets are the same as the ConvertTo-Ini and ConvertFrom-Ini cmdlets, except that they save the INI-strings in a file.
'@
        Example = @'
ConvertTo-Ini -InputObject $Hashtable
 
    Exemplary variable
    ------------------

    $Hashtable = @{
        Section = @{
            Key = 'Value'
        }
    }

    Description
    -----------
    Convert single hashtable, return single INI string.
'@,
@'
ConvertTo-Ini -InputObject $HashtableA, $HashtableB
 
    Exemplary variables
    -------------------

    $HashtableA = @{
    SectionA = @{
    KeyA = 'ValueA'
    }
    }

    $HashtableB = @{
    SectionB = @{
    KeyB = 'ValueB'
    }
    }

    Description
    -----------
    Convert multiple hashtables, return multiple INI strings.
'@,
@'
$HashtableA, $HashtableB | ConvertTo-Ini
 
    Exemplary variables
    -------------------

    $HashtableA = @{
    SectionA = @{
    KeyA = 'ValueA'
    }
    }

    $HashtableB = @{
    SectionB = @{
    KeyB = 'ValueB'
    }
    }

    Description
    -----------
    Convert multiple hashtables specified via pipeline input, return multiple INI strings.
'@,
@'
ConvertTo-Ini -InputObject $HashtableA, $HashtableB -Merge
 
    Exemplary variables
    -------------------

    $HashtableA = @{
    Section = @{
    Key = 'ValueA'
    }
    }

    $HashtableB = @{
    Section = @{
    Key = 'ValueB'
    }
    }

    Description
    -----------
    Convert multiple hashtables, merge them, return single INI string. Duplicate keys are not overriden.

    Result:

    [Section]
    Key=ValueA
'@,
@'
ConvertTo-Ini -InputObject $HashtableA, $HashtableB -Merge -OverrideDuplicateKeys
 
    Exemplary variables
    -------------------

    $HashtableA = @{
    Section = @{
    Key = 'ValueA'
    }
    }

    $HashtableB = @{
    Section = @{
    Key = 'ValueB'
    }
    }

    Description
    -----------
    Convert multiple hashtables, merge them, return single INI string. Duplicate keys are overriden.

    Result:

    [Section]
    Key=ValueB
'@,
@'
ConvertTo-Ini -InputObject $HashtableA, $HashtableB, $IniData -Merge
 
    Exemplary variables
    -------------------

    $HashtableA = @{
    SectionA = @{
    KeyA = 'ValueA'
    }
    }

    $HashtableB = @{
    SectionB = @{
    KeyB = 'ValueB'
    }
    }

    $IniString = @'
    [SectionC]
    KeyC=ValueC
    ''@

    $IniData = $IniString | ConvertFrom-Ini -AsObject

    Description
    -----------
    You can mix and match hashtables and [IniParser.IniData] objects. This example will convert and merge them to a single INI string.

    Result:

    [SectionA]
    KeyA=ValueA

    [SectionB]
    KeyB=ValueB

    [SectionC]
    KeyC=ValueC
'@
    }
    Common = @{
        Parameter = @{
            CommentStrings = @'
This parameter is optional. Default value is ';','#'.

Sets the strings that defines the start of a comment in INI file. A comment spans from the comment character to the end of the line.
'@
            SectionStartChar = @'
This parameter is optional. Default value is '['.

Sets the character that defines the start of a section name.
'@
            SectionEndChar = @'
This parameter is optional. Default value is ']'.

Sets the character that defines the end of a section name.
'@
            KeyValueAssigmentChar = @'
This parameter is optional. Default value is '='.

Sets the character that defines a value assigned to a key.
'@
            OverrideDuplicateKeys = @'
This parameter is optional. Default value is False.

If this switch is specified, when the parser finds a duplicate key, it overwrites the previous value, so the key will always contain the value of the last key read in the file. If set to false the first read value is preserved, so the key will always contain the value of the first key read in the file.
'@
            Encoding = @'
This parameter is optional. Default value is the encoding for the operating system''s current ANSI code page.

Specifies the type of character encoding that was used in the INI file. Accepted values include: US-ASCII, UTF-7, UTF-8, UTF-16, UTF-32 or any other valid .NET encoding name. To get the list run [System.Text.Encoding]::GetEncodings() .
'@
            AsObject = 'If this switch is specified, cmdlet will return [IniParser.IniData] object, instead of hashtable. It preserves additional metadata in INI file, such as comments. This object can be passed to Export-Ini and ConvertTo-Ini cmdlets.'
        }
        Notes = 'This cmdlet uses INI File Parser by Ricardo Amores Hernández. It''s a .NET, Mono and Unity3d compatible library for reading/writing INI data from IO streams, file streams, and strings written in C#.'
        Link = 'https://github.com/rickyah/ini-parser', 'https://github.com/beatcracker/Powershell-Misc'
    }
}

#endregion

#region Boostrap

# Create proxy functions, that will wrap Parse\Serialize functions.
# This code runs on module import.

$ProxyFunctionsMap = @{
    'Parse-Ini' = @{
        'Import-Ini' = 'String'
        'ConvertFrom-Ini' = 'File'
    }
    'Serialize-Ini' = @{
        'Export-Ini' = 'Object', 'Merge'
        'ConvertTo-Ini' = 'File'
    }
}

# Can't use ForEach-Object due to scoping: we need to execute returned scriptblock in the current scope
foreach($BaseFunction in $ProxyFunctionsMap.GetEnumerator()) {

    foreach($ProxyFunction in $BaseFunction.Value.GetEnumerator()) {
        Write-Debug "Creating new proxy function: $($ProxyFunction.Key)"
        . (New-ProxyCommandFromParameterSet -CommandName $BaseFunction.Key -ProxyCommandName $ProxyFunction.Key -ExcludeParameterSet $ProxyFunction.Value -HelpData $HelpData -AsScriptblock)

        Write-Debug "Exporting proxy function: $($ProxyFunction.Key)"
        Export-ModuleMember -Function $ProxyFunction.Key
    }
}

#endregion