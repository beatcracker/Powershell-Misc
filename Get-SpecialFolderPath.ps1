<#
.Synopsis
	Gets the path to the system special folder that is identified by the specified enumeration (CSIDL, KNOWNFOLDERID or best match).

.Description
    Gets the path to the system special folder that is identified by the specified enumeration. On pre .NET 4.0 systems tries to map
    unknown KNOWNFOLDERID to CSIDLs. This, for example allows to query for 'ProgramFilesx86' directory when PowerShell is running
    in .Net 3.5, where SpecialFolder enumeration contains only KNOWNFOLDERID for 'ProgramFiles'.

.Parameter Name
	An array of strings (SpecialFolder, CSIDL or both) to query. Accepts pipeline input.

.Parameter Csidl
	An array of strings containing CSIDLs to query.

.Parameter KnownFolderId
	An array of strings containing KNOWNFOLDERIDs to query.

.Example
    'Favorites' | Get-SpecialFolderPath

    Get folder path for 'Favorites' folder.

.Example
    Get-SpecialFolderPath -Name 'Favorites'

    Get folder path for 'Favorites' folder.

.Example
    'Favorites', 'CSIDL_DESKTOP' | Get-SpecialFolderPath

    Get folder paths for 'Favorites' and Desktop folders

.Example
    Get-SpecialFolderPath -Name 'Favorites', 'CSIDL_DESKTOP'

    Get folder paths for 'Favorites' and Desktop folders

.Example
    Get-SpecialFolderPath -Csidl 'CSIDL_DESKTOP'

    Get folder path for Desktop folder by CSIDL

.Example
    Get-SpecialFolderPath -KnownFolderId 'Favorites'

    Get folder path for 'Favorites' folder by KNOWNFOLDERID

.Example
    Get-SpecialFolderPath -Name 'NetHood'

    Get folder path for 'NetHood' KnownFolderId. On my system, there is no 'NetHood' KNOWNFOLDERID in SpecialFolder enumeration,
    so the function will fallback to CSIDL mapping. Example of verbose output in this situation:

    VERBOSE: Checking if [System.Environment]::GetFolderPath available
    VERBOSE: Result: True
    VERBOSE: No KnownFolderId for: 'NetHood', trying to map to CSIDL
    VERBOSE: KnownFolderId 'NetHood' is mapped to CSIDL(s): CSIDL_NETHOOD
    VERBOSE: Registering type: Shell32.Tools
    VERBOSE: Processing CSIDL(s): CSIDL_NETHOOD
    C:\Users\beatcracker\AppData\Roaming\Microsoft\Windows\Network Shortcuts
#>
function Get-SpecialFolderPath
{
    [CmdletBinding()]
    Param()
    DynamicParam
    {
        # CSIDLs are taken from the latest Shlobj.h (Win10 SDK)
        # https://gist.github.com/beatcracker/4b154d46cc26776b50e7/raw/a317160dad57157f100e0f6e6d68c692c2bee7f1/ShlObj.h
        $CsidlEnum = @{
            CSIDL_ADMINTOOLS = 0x0030 # <user name>\Start Menu\Programs\Administrative Tools
            CSIDL_ALTSTARTUP = 0x001d # non localized startup
            CSIDL_APPDATA = 0x001a # <user name>\Application Data
            CSIDL_BITBUCKET = 0x000a # <desktop>\Recycle Bin
            CSIDL_CDBURN_AREA = 0x003b # USERPROFILE\Local Settings\Application Data\Microsoft\CD Burning
            CSIDL_COMMON_ADMINTOOLS = 0x002f # All Users\Start Menu\Programs\Administrative Tools
            CSIDL_COMMON_ALTSTARTUP = 0x001e # non localized common startup
            CSIDL_COMMON_APPDATA = 0x0023 # All Users\Application Data
            CSIDL_COMMON_DESKTOPDIRECTORY = 0x0019 # All Users\Desktop
            CSIDL_COMMON_DOCUMENTS = 0x002e # All Users\Documents
            CSIDL_COMMON_FAVORITES = 0x001f
            CSIDL_COMMON_MUSIC = 0x0035 # All Users\My Music
            CSIDL_COMMON_OEM_LINKS = 0x003a # Links to All Users OEM specific apps
            CSIDL_COMMON_PICTURES = 0x0036 # All Users\My Pictures
            CSIDL_COMMON_PROGRAMS = 0X0017 # All Users\Start Menu\Programs
            CSIDL_COMMON_STARTMENU = 0x0016 # All Users\Start Menu
            CSIDL_COMMON_STARTUP = 0x0018 # All Users\Startup
            CSIDL_COMMON_TEMPLATES = 0x002d # All Users\Templates
            CSIDL_COMMON_VIDEO = 0x0037 # All Users\My Video
            CSIDL_COMPUTERSNEARME = 0x003d # Computers Near Me (computered from Workgroup membership)
            CSIDL_CONNECTIONS = 0x0031 # Network and Dial-up Connections
            CSIDL_CONTROLS = 0x0003 # My Computer\Control Panel
            CSIDL_COOKIES = 0x0021
            CSIDL_DESKTOP = 0x0000 # <desktop>
            CSIDL_DESKTOPDIRECTORY = 0x0010 # <user name>\Desktop
            CSIDL_DRIVES = 0x0011 # My Computer
            CSIDL_FAVORITES = 0x0006 # <user name>\Favorites
            CSIDL_FONTS = 0x0014 # windows\fonts
            CSIDL_HISTORY = 0x0022
            CSIDL_INTERNET = 0x0001 # Internet Explorer (icon on desktop)
            CSIDL_INTERNET_CACHE = 0x0020
            CSIDL_LOCAL_APPDATA = 0x001c # <user name>\Local Settings\Applicaiton Data (non roaming)
            CSIDL_MYDOCUMENTS = 0x0005 # My Documents
            CSIDL_MYMUSIC = 0x000d # "My Music" folder
            CSIDL_MYPICTURES = 0x0027 # C:\Program Files\My Pictures
            CSIDL_MYVIDEO = 0x000e # "My Videos" folder
            CSIDL_NETHOOD = 0x0013 # <user name>\nethood
            CSIDL_NETWORK = 0x0012 # Network Neighborhood (My Network Places)
            CSIDL_PERSONAL = 0x0005 # Personal was just a silly name for My Documents
            CSIDL_PRINTERS = 0x0004 # My Computer\Printers
            CSIDL_PRINTHOOD = 0x001b # <user name>\PrintHood
            CSIDL_PROFILE = 0x0028 # USERPROFILE
            CSIDL_PROGRAM_FILES = 0x0026 # C:\Program Files
            CSIDL_PROGRAM_FILES_COMMON = 0x002b # C:\Program Files\Common
            CSIDL_PROGRAM_FILES_COMMONX86 = 0x002c # x86 Program Files\Common on RISC
            CSIDL_PROGRAM_FILESX86 = 0x002a # x86 C:\Program Files on RISC
            CSIDL_PROGRAMS = 0x0002 # Start Menu\Programs
            CSIDL_RECENT = 0x0008 # <user name>\Recent
            CSIDL_RESOURCES = 0x0038 # Resource Direcotry
            CSIDL_RESOURCES_LOCALIZED = 0x0039 # Localized Resource Direcotry
            CSIDL_SENDTO = 0x0009 # <user name>\SendTo
            CSIDL_STARTMENU = 0x000b # <user name>\Start Menu
            CSIDL_STARTUP = 0x0007 # Start Menu\Programs\Startup
            CSIDL_SYSTEM = 0x0025 # GetSystemDirectory()
            CSIDL_SYSTEMX86 = 0x0029 # x86 system directory on RISC
            CSIDL_TEMPLATES = 0x0015
            CSIDL_WINDOWS = 0x0024 # GetWindowsDirectory()
        }

        # KNOWNFOLDERID CSIDL equivalents:
        # https://msdn.microsoft.com/en-us/library/windows/desktop/dd378457.aspx
        $KnownFolderIdToCsidl = @{
            AdminTools = 'CSIDL_ADMINTOOLS'
            CDBurning = 'CSIDL_CDBURN_AREA'
            CommonAdminTools = 'CSIDL_COMMON_ADMINTOOLS'
            CommonOEMLinks = 'CSIDL_COMMON_OEM_LINKS'
            CommonPrograms = 'CSIDL_COMMON_PROGRAMS'
            CommonStartMenu = 'CSIDL_COMMON_STARTMENU'
            CommonStartup = 'CSIDL_COMMON_STARTUP', 'CSIDL_COMMON_ALTSTARTUP'
            CommonTemplates = 'CSIDL_COMMON_TEMPLATES'
            ComputerFolder = 'CSIDL_DRIVES'
            ConnectionsFolder = 'CSIDL_CONNECTIONS'
            ControlPanelFolder = 'CSIDL_CONTROLS'
            Cookies = 'CSIDL_COOKIES'
            Desktop = 'CSIDL_DESKTOP', 'CSIDL_DESKTOPDIRECTORY'
            Documents = 'CSIDL_MYDOCUMENTS', 'CSIDL_PERSONAL'
            Favorites = 'CSIDL_FAVORITES', 'CSIDL_COMMON_FAVORITES'
            Fonts = 'CSIDL_FONTS'
            History = 'CSIDL_HISTORY'
            InternetCache = 'CSIDL_INTERNET_CACHE'
            InternetFolder = 'CSIDL_INTERNET'
            LocalAppData = 'CSIDL_LOCAL_APPDATA'
            LocalizedResourcesDir = 'CSIDL_RESOURCES_LOCALIZED'
            Music = 'CSIDL_MYMUSIC'
            NetHood = 'CSIDL_NETHOOD'
            NetworkFolder = 'CSIDL_NETWORK', 'CSIDL_COMPUTERSNEARME'
            Pictures = 'CSIDL_MYPICTURES'
            PrintersFolder = 'CSIDL_PRINTERS'
            PrintHood = 'CSIDL_PRINTHOOD'
            Profile = 'CSIDL_PROFILE'
            ProgramData = 'CSIDL_COMMON_APPDATA'
            ProgramFiles = 'CSIDL_PROGRAM_FILES'
            ProgramFilesCommon = 'CSIDL_PROGRAM_FILES_COMMON'
            ProgramFilesCommonX86 = 'CSIDL_PROGRAM_FILES_COMMONX86'
            ProgramFilesX86 = 'CSIDL_PROGRAM_FILESX86'
            Programs = 'CSIDL_PROGRAMS'
            PublicDesktop = 'CSIDL_COMMON_DESKTOPDIRECTORY'
            PublicDocuments = 'CSIDL_COMMON_DOCUMENTS'
            PublicMusic = 'CSIDL_COMMON_MUSIC'
            PublicPictures = 'CSIDL_COMMON_PICTURES'
            PublicVideos = 'CSIDL_COMMON_VIDEO'
            Recent = 'CSIDL_RECENT'
            RecycleBinFolder = 'CSIDL_BITBUCKET'
            ResourceDir = 'CSIDL_RESOURCES'
            RoamingAppData = 'CSIDL_APPDATA'
            SendTo = 'CSIDL_SENDTO'
            StartMenu = 'CSIDL_STARTMENU'
            Startup = 'CSIDL_STARTUP', 'CSIDL_ALTSTARTUP'
            System = 'CSIDL_SYSTEM'
            SystemX86 = 'CSIDL_SYSTEMX86'
            Templates = 'CSIDL_TEMPLATES'
            Videos = 'CSIDL_MYVIDEO'
            Windows = 'CSIDL_WINDOWS'
        }

        $DynamicParameters = @{
			Name = 'Name'
			Type = [string[]]
            Parameter = @{
			    Position = 0
                ValueFromPipeline = $true
                ValueFromPipelineByPropertyName = $true
                ParameterSetName = 'Name'
            }
			ValidateSet = $CsidlEnum.Keys + $KnownFolderIdToCsidl.Keys
        },
        @{
			Name = 'Csidl'
			Type = [string[]]
            Parameter = @{
			    Position = 0
                ValueFromPipelineByPropertyName = $true
                ParameterSetName = 'Csidl'
            }
			ValidateSet = $CsidlEnum.Keys
		},
        @{
			Name = 'KnownFolderId'
            Type = [Environment+SpecialFolder[]]
            Parameter = @{
			    Position = 0
                ValueFromPipelineByPropertyName = $true
                ParameterSetName = 'KnownFolderId'
            }
        }

        # Create the dictionary
        $RuntimeParameterDictionary = New-Object -TypeName System.Management.Automation.RuntimeDefinedParameterDictionary

        foreach($dp in $DynamicParameters)
        {
            # Create the collection of attributes
            $AttributeCollection = New-Object -TypeName System.Collections.ObjectModel.Collection[System.Attribute]

            # Create and set the parameters' attributes (Parameter, ValidateSet)

            # [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName)]
            $AttributeCollection.Add(
                (New-Object -TypeName System.Management.Automation.ParameterAttribute -Property $dp.Parameter)
            )

            # [ValidateSet()]
            if($dp.ValidateSet)
            {
                $AttributeCollection.Add(
                    (New-Object -TypeName System.Management.Automation.ValidateSetAttribute -ArgumentList $dp.ValidateSet)
                )
            }

            # Create and return the dynamic parameter
            $RuntimeParameter = New-Object -TypeName System.Management.Automation.RuntimeDefinedParameter(
                $dp.Name,
                $dp.Type,
                $AttributeCollection
            )
            $RuntimeParameterDictionary.Add($dp.Name, $RuntimeParameter)
        }

        # Return dicitonary
        $RuntimeParameterDictionary
    }

    Begin
    {
        Write-Verbose 'Checking if [System.Environment]::GetFolderPath available'
        $GetFolderPath = $false
        try
        {
            $GetFolderPath = [bool][System.Environment].GetMethod('GetFolderPath')
        }
        catch [System.Reflection.AmbiguousMatchException]
        {
            $GetFolderPath = $true
        }

        Write-Verbose "Result: $GetFolderPath"

        $hwnd = New-Object -TypeName UIntPtr
        $lpData = New-Object -TypeName System.Text.StringBuilder(260) # MAX_PATH = 256

        function Register-Shell32Tools
        {
            if(![bool]('Shell32.Tools' -as [Type]))
            {
                $Enum = $(foreach($key in $CsidlEnum.GetEnumerator()){'{0}{1} = {2}' -f [System.Environment]::NewLine, $key.Name, $key.Value}) -join ','
                $Shell32Tools = @'
                    [DllImport("Shell32.dll")]
                    public static extern int SHGetSpecialFolderPath(
                        UIntPtr hwndOwner,
                        System.Text.StringBuilder lpszPath,
                        CSIDL iCsidl,
                        int fCreate
                    );

                    public enum CSIDL : int {{{0}
                    }}
'@
                Write-Verbose 'Registering type: Shell32.Tools'
                Add-Type -MemberDefinition ($Shell32Tools -f $Enum) -Name Tools -Namespace Shell32 -Using System.Text -ErrorAction Stop
            }
            else
            {
                Write-Verbose 'Type already registered: Shell32.Tools'
            }
        }

        function Get-FolderPathByCsidl
        {
            Param
            (
                [Parameter(Mandatory=$true)]
                [UIntPtr]$hwnd,

                [Parameter(Mandatory=$true)]
                [System.Text.StringBuilder]$lpData,

                [Parameter(Mandatory=$true)]
                [string[]]$Csidl
            )

            Register-Shell32Tools

            Write-Verbose "Processing CSIDL(s): $($Csidl -join ', ')"

            foreach($item in $Csidl)
            {
                if([Shell32.Tools]::SHGetSpecialFolderPath($hwnd, $lpData, $item, 0))
                {
                    $lpData.ToString()
                }
                else
                {
                    Write-Error "Folder not found: $item"
                }
            }
        }

        function Get-FolderPathByKnownFolderId
        {
            Param
            (
                [Parameter(Mandatory = $true)]
                [System.Environment+SpecialFolder[]]$Path
            )

            Write-Verbose "Processing KnownFolderId(s): $($Path -join ', ')"
            foreach($item in $Path)
            {
                [Environment]::GetFolderPath($item)
            }
        }
    }

    Process
    {
        if($PSBoundParameters.Name)
        {
            foreach($item in $PSBoundParameters.Name)
            {
                if($item -like 'CSIDL_*')
                {
                    Get-FolderPathByCsidl -hwnd $hwnd -lpData $lpData -Csidl $item
                }
                else
                {
                    if($GetFolderPath -and [System.Enum]::GetNames('System.Environment+SpecialFolder') -contains $item)
                    {
                        Get-FolderPathByKnownFolderId -Path $item
                    }
                    else
                    {
                        Write-Verbose "No KnownFolderId for: '$item', trying to map to CSIDL"
                        $MappedCsidl = $KnownFolderIdToCsidl[$item]

                        if($MappedCsidl)
                        {
                            Write-Verbose "KnownFolderId '$item' is mapped to CSIDL(s): $($MappedCsidl -join ', ')"
                            Get-FolderPathByCsidl -hwnd $hwnd -lpData $lpData -Csidl $MappedCsidl | Select-Object -Unique
                        }
                        else
                        {
                            Write-Error "Folder not found. There is no known CSIDL mapping for '$item'"
                        }
                    }

                }
            }
        }
        elseif($PSBoundParameters.Csidl)
        {
            Get-FolderPathByCsidl -hwnd $hwnd -lpData $lpData -Csidl $PSBoundParameters.Csidl
        }
        elseif($PSBoundParameters.KnownFolderId)
        {
            Get-FolderPathByKnownFolderId -Path $PSBoundParameters.KnownFolderId
        }
        else
        {
            throw 'No valid parameters supplied!'
        }
    }
}