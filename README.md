# Miscellaneous PowerShell goodies

## Table of Contents

- [How to use Functions\Scripts](#how-to-use-functionsscripts)
  - [In PowerShell console\PowerShell ISE script pane](#in-powershell-consolepowershell-ise-script-pane)
  - [In your own script](#in-your-own-script)
- [How to use Modules](#how-to-use-modules)
- [Functions](#functions)
  - [Get-TerminologyTranslation](#get-terminologytranslation)
  - [Split-CommandLine](#split-commandline)
  - [Import-Component](#import-component)
  - [New-DynamicParameter](#new-dynamicparameter)
  - [Get-SvnAuthor](#get-svnauthor)
  - [Step-Dictionary](#step-dictionary)
  - [Get-SpecialFolderPath](#get-specialfolderpath)
  - [Start-ConsoleProcess](#start-consoleprocess)
  - [Remove-ComObject](#remove-comobject)
- [Scripts](#scripts)
  - [New-GitSvnAuthorsFile](#new-gitsvnauthorsfile)
- [Modules](#modules)
  - [PsIniParser](#psiniparser)

## How to use Functions\Scripts

The best way to use provided functions is a [dot-sourcing](http://ss64.com/ps/source.html).

Dot-sourcing runs a script file in the current scope so that any functions, aliases, and variables that the script file creates are added to the current scope. 

### In PowerShell console\PowerShell ISE script pane

*Note the space between first dot and path!*

```powershell
. c:\scripts\Get-TerminologyTranslation.ps1
```

Or, navigate to the folder, where you downloaded script file

```
cd c:\scripts
```

And then type:

```powershell
. .\Get-TerminologyTranslation.ps1
```

To verify, that script file is loaded, try to view help for the loaded function. The function name is usually the same as script file name (without extension), if not stated otherwise in readme.

```powershell
Get-Help -Name Get-TerminologyTranslation -Full
```

Congratulations, you can now make use of newly added function!

```powershell
'Windows Update' | Get-TerminologyTranslation -From 'en-us' -To 'ru-ru' -Source Terms

```

### In your own script

*Note the space between first dot and path!*

Dot-source from arbitrary location before calling the function:

```powershell
. c:\scripts\Get-TerminologyTranslation.ps1
'Windows Update' | Get-TerminologyTranslation -From 'en-us' -To 'ru-ru' -Source Terms
```

Or put it alongside with your script, get path to the script folder programmatically, and then dot-source:

```powershell
$ScriptDir = Split-Path $script:MyInvocation.MyCommand.Path
. (Join-Path -Path $ScriptDir -ChildPath 'Get-TerminologyTranslation.ps1')
'Windows Update' | Get-TerminologyTranslation -From 'en-us' -To 'ru-ru' -Source Terms

```

## How to use Modules
Download\clone this repository and copy module folder to your PowerShell modules folder. If you downloaded repository as ZIP file, you need to unblock it first:

* Using GUI: right-click ZIP file, click `Properties` and then click the `Unblock` button.
* Using PowerShell: `Unblock-File 'X:\Path\to\file.zip'`

PowerShell will look in the paths specified in the `$env:PSModulePath` environment variable when searching for available modules on a system. Default locations are:
  * System-wide:
    * `C:\Program Files\WindowsPowerShell\Modules`
    * `C:\Windows\system32\WindowsPowerShell\v1.0\Modules\`
  * Per-user:
    * `C:\Users\USERNAME\Documents\WindowsPowerShell\Modules`

Modules stored in those locations are easily discoverable and autoloaded with PowerShell 3.0 and higher. If you're not sure, copy module to your *Per-user* folder.

* To list all available modules, use:
```powershell
Get-Module -ListAvailable
```
* To import available module, use:
```powershell
Import-Module -Name 'ModuleName'
```

Alternatively, you can import module from any location:
```powershell
Import-Module -Name 'X:\Path\to\module_folder'
```

## Functions

### [Get-TerminologyTranslation](Get-TerminologyTranslation.ps1)

Enables user to look up terminology translations and user-interface translations from actual Microsoft products via [Microsoft Terminology Service API](http://www.microsoft.com/Language/en-US/Microsoft-Terminology-API.aspx). For details see [ Terminology Service API SDK PDF](http://download.microsoft.com/download/1/5/D/15D3DDC6-7403-4366-BE99-AF5247ADEF1C/Microsoft-Terminology-API-SDK.pdf).

Features

  * Any-to-any language translation searches, e.g. Japanese to/from French or any other language combination.
  * Filter searches with string case and hotkey sensitivity.
  * Filter searches by product name and version.
  * Get list of languages supported by the Terminology Service API.
  * Get list of products supported by the Terminology Service API.
  * Full comment-based help and usage examples.

Usage examples:

* Ever wonder how [Cherokee](http://en.wikipedia.org/wiki/Cherokee) spell `Start menu`?
```powershell
'Start menu' | Get-TerminologyTranslation -From 'en-us' -To 'chr-cher-us' -Source Both

ᎠᏂᎩᏍᏙᏗ ᏗᏑᏰᏍᏗᎢ
```

* What is `Control panel` in French?
```powershell
'Control panel' | Get-TerminologyTranslation -From 'en-us' -To 'fr-fr' -Source Both

Panneau de configuration
```

* How about `Fatal Error` in German?
```powershell
'Fatal error' | Get-TerminologyTranslation -From 'en-us' -To 'de-de' -Source Both

Schwerwiegender Fehler
```

* You've got a cyrillic SharePoint installation and wonder where the hell is `Application Management` in SharePoint Central Administration?
```powershell
'Application Management' | Get-TerminologyTranslation -From 'en-us' -To 'ru-ru' -Source 'UiStrings' -Name 'SharePoint Server'

Управление приложениями
```

### [Split-CommandLine](Split-CommandLine.ps1)

PowerShell version of [EchoArgs](http://blogs.technet.com/b/heyscriptingguy/archive/2011/09/20/solve-problems-with-external-command-lines-in-powershell.aspx). This is the Cmdlet version of the code from the article [PowerShell and external commands done right](http://edgylogic.com/blog/powershell-and-external-commands-done-right). It can parse command-line arguments using Win32 API function [CommandLineToArgvW](http://msdn.microsoft.com/en-us/library/windows/desktop/bb776391.aspx) and echo command line arguments back out to the console for your review.

Features

  * Parse arbitrary command-line, or if none specified, the command-line of the current PowerShell host.
  * Full comment-based help and usage examples.

Usage examples:

* Parse user-specified command-line
```powershell
'"c:\windows\notepad.exe" test.txt' | Split-CommandLine

c:\windows\notepad.exe
test.txt
```

* Parse command-line of the live process
```powershell
Get-WmiObject Win32_Process -Filter "Name='notepad.exe'" | Split-CommandLine

c:\windows\notepad.exe
test.txt
```

### [Import-Component](Import-Component.ps1)

Bulk-import from folder any component, supported by PowerShell (script, module, source code, .Net assembly).

Features

  * Supported components:
    * Script (.ps1) - imported using [Dot-Sourcing](http://ss64.com/ps/source.html).
    * Module - imported using [Import-Module](http://technet.microsoft.com/en-us/library/hh849725.aspx) cmdlet  
        _This function will only try to import well-formed modules. A "well-formed" module is a module that is stored in a directory that has the same name as the base name of at least one file in the module directory. If a module is not well-formed, Windows PowerShell does not recognize it as a module. [More info](https://msdn.microsoft.com/en-us/library/dd878350.aspx)_
    * Source code (.cs, .vb, .js) - imported using [Add-Type](http://technet.microsoft.com/en-us/library/hh849914.aspx) cmdlet
    * .Net assembly (.dll) - imported using [Add-Type](http://technet.microsoft.com/en-us/library/hh849914.aspx) cmdlet
  * Full comment-based help and usage examples.

__WARNING: To import .PS1 scripts this function itself has to be dot-sourced!__ Example:

*Note the space between first dot and function name!*
```powershell
. Import-Component 'C:\PsLib'
```

Usage example

* Import all supported components (`.ps1`, `module`, `.cs`, `.vb`, `.js`, `.dll`), recurse into subdirectories. Include only files with names without extension that match wildcards `MyScript*` and `*MyLib*`. Exclude files with names without extension that match `*_backup*` and `*_old*` wildcards.

```powershell
. Import-Component 'C:\PsLib' -Recurse -Include 'MyScript*','*MyLib*' -Exclude '*_backup*','*_old*'
```
### [New-DynamicParameter](New-DynamicParameter.ps1)

Helper function to simplify creating [dynamic parameters](https://technet.microsoft.com/en-us/library/hh847743.aspx).
Example use cases:
 - Include parameters only if your environment dictates it
 - Include parameters depending on the value of a user-specified parameter
 - Provide tab completion and intellisense for parameters, depending on the environment

Credits to Justin Rich ([blog](http://jrich523.wordpress.com), [GitHub](https://github.com/jrich523)) and Warren F. ([blog](http://ramblingcookiemonster.github.io), [GitHub](https://github.com/RamblingCookieMonster)) for their initial code and inspiration:
 - [New-DynamicParam.ps1](https://github.com/RamblingCookieMonster/PowerShell/New-DynamicParam.ps1)
 - [Credentials and Dynamic Parameters](http://ramblingcookiemonster.wordpress.com/2014/11/27/quick-hits-credentials-and-dynamic-parameters/)
 - [PowerShell: Simple way to add dynamic parameters to advanced function](http://jrich523.wordpress.com/2013/05/30/powershell-simple-way-to-add-dynamic-parameters-to-advanced-function/)

Credit to BM for alias and type parameters and their handling.

Features

  * Create dynamic parameters for your functions on the fly.
  * Full comment-based help and usage examples.

Usage example

* Create one dynamic parameter. This example illustrates the use of `New-DynamicParameter` to create a single dynamic parameter. The `Drive`'s parameter `ValidateSet` is populated with all available volumes on the computer for handy tab completion / intellisense. 

```powershell
function Get-FreeSpace
{
	[CmdletBinding()]
	Param()
	DynamicParam
	{
		# Get drive names for ValidateSet attribute
		$DriveList = ([System.IO.DriveInfo]::GetDrives()).Name

		# Create new dynamic parameter
		New-DynamicParameter -Name Drive -ValidateSet $DriveList -Type ([array]) -Position 0 -Mandatory
	}

	Process
	{
		# Dynamic parameters don't have corresponding variables created,
		# you need to call New-DynamicParameter with CreateVariables switch to fix that.
		New-DynamicParameter -CreateVariables -BoundParameters $PSBoundParameters

		$DriveInfo = [System.IO.DriveInfo]::GetDrives() | Where-Object {$Drive -contains $_.Name}
		$DriveInfo |
			ForEach-Object {
				if(!$_.TotalFreeSpace)
				{
					$FreePct = 0
				}
				else
				{
					$FreePct = [System.Math]::Round(($_.TotalSize / $_.TotalFreeSpace), 2)
				}
				New-Object -TypeName psobject -Property @{
					Drive = $_.Name
					DriveType = $_.DriveType
					'Free(%)' = $FreePct
				}
			}
	}
}
```

### [Get-SvnAuthor](Get-SvnAuthor.ps1)

Get list of unique commit authors in one or more SVN repositories. Requires Subversion binaries. Can be used to create authors file for SVN to Git migrations.

Features

  * Get list of unique commit authors in one or more SVN repositories.
  * Full comment-based help and usage examples.

Usage example

* Get list of unique commit authors for SVN repository `http://svnserver/svn/project`

```powershell
'http://svnserver/svn/project' | Get-SvnAuthor

John Doe
Jane Doe

```

### [Step-Dictionary](Step-Dictionary.ps1)

Recursively walk through each item in a dictionary and execute scriptblock against lowest level keys. You can modify and remove lowest level keys while iterating over dictionary.

Features

  * Recursively walk through each item in a dictionary and execute scriptblock against lowest level keys.
  * Full comment-based help and usage examples.

Usage examples


```powershell

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

# Print each lowest level key name and value
Step-Dictionary -Dictionary $Dictionary -ScriptBlock {"${key}: $($Dictionary[$key])" | Write-Host}

# Set each lowest level key to random value
Step-Dictionary -Dictionary $Dictionary -ScriptBlock {$Dictionary[$key] = Get-Random}

# Remove every lowest level key
Step-Dictionary -Dictionary $Dictionary -ScriptBlock {$Dictionary.Remove($key)}
```

### [Get-SpecialFolderPath](Get-SpecialFolderPath.ps1)

Gets the path to the system special folder that is identified by the specified enumeration. On pre .NET 4.0 systems tries to map unknown [KNOWNFOLDERID](https://msdn.microsoft.com/en-us/library/windows/desktop/dd378457.aspx) to [CSIDLs](https://gist.github.com/beatcracker/4b154d46cc26776b50e7/raw/a317160dad57157f100e0f6e6d68c692c2bee7f1/ShlObj.h). This, for example allows to query for `ProgramFilesx86` directory when PowerShell is running in .Net 3.5, where [SpecialFolder enumeration](https://msdn.microsoft.com/en-us/library/system.environment.specialfolder.aspx) contains only `KNOWNFOLDERID` for `ProgramFiles`.

Features

  * Gets the path to the system special folder that is identified by the specified enumeration (`CSIDL`, `KNOWNFOLDERID` or best match).
  * Full comment-based help and usage examples.

Usage examples


* Get folder paths for 'Favorites' and Desktop folders using both `CSIDL` and `KNOWNFOLDERID`
```powershell
'Favorites', 'CSIDL_DESKTOP' | Get-SpecialFolderPath
Get-SpecialFolderPath -Name 'Favorites', 'CSIDL_DESKTOP'
```

* Get folder path for Desktop folder by `CSIDL`
```powershell
Get-SpecialFolderPath -Csidl 'CSIDL_DESKTOP'
```

* Get folder path for 'Favorites' folder by `KNOWNFOLDERID`

```powershell
Get-SpecialFolderPath -KnownFolderId 'Favorites'
```

* Get folder path for `NetHood` `KNOWNFOLDERID`.
```powershell
Get-SpecialFolderPath -Name 'NetHood'
```
On my system, there is no `NetHood` `KNOWNFOLDERID` in SpecialFolder enumeration, so the function will fallback to `CSIDL` mapping. Example of verbose output in this situation:

    VERBOSE: Checking if [System.Environment]::GetFolderPath available
    VERBOSE: Result: True
    VERBOSE: No KnownFolderId for: 'NetHood', trying to map to CSIDL
    VERBOSE: KnownFolderId 'NetHood' is mapped to CSIDL(s): CSIDL_NETHOOD
    VERBOSE: Registering type: Shell32.Tools
    VERBOSE: Processing CSIDL(s): CSIDL_NETHOOD
    C:\Users\beatcracker\AppData\Roaming\Microsoft\Windows\Network Shortcuts

### [Start-ConsoleProcess](Start-ConsoleProcess.ps1)

This function will start console executable, pipe any user-specified strings to it and capture `StandardOutput`/`StandardError` streams and `exit code`.

Features

  * Returns object with following properties:
    * `StdOut` - array of strings captured from `StandardOutput`
    * `StdErr` - array of strings captured from `StandardError`
    * `ExitCode` - exit code set by executable
  * Full comment-based help and usage examples.

Usage examples

* Start `find.exe` and capture its output. Because no arguments specified, `find.exe` prints error to `StandardError` stream, which is captured by the function.
```powershell
Start-ConsoleProcess -FilePath find

StdOut StdErr                               ExitCode
------ ------                               --------
{}     {FIND: Parameter format not correct}        2
```

* Start `robocopy.exe` with arguments  and capture its output. `Robocopy.exe` will mirror contents of the `C:\Src` folder to `C:\Dst` and print log to `StandardOutput` stream, which is captured by the function.
```powershell
$Result = Start-ConsoleProcess -FilePath robocopy -ArgumentList 'C:\Src', 'C:\Dst', '/mir'
$Result.StdOut

-------------------------------------------------------------------------------
   ROBOCOPY     ::     Robust File Copy for Windows                              
-------------------------------------------------------------------------------

  Started : 01 January 2016 y. 00:00:01
   Source : C:\Src\
     Dest : C:\Dst\

    Files : *.*
    
  Options : *.* /S /E /DCOPY:DA /COPY:DAT /PURGE /MIR /R:1000000 /W:30 

------------------------------------------------------------------------------

                       1	C:\Src\
        New File  		       6	Readme.txt
  0%  
100%  

------------------------------------------------------------------------------

               Total    Copied   Skipped  Mismatch    FAILED    Extras
    Dirs :         1         0         0         0         0         0
   Files :         1         1         0         0         0         0
   Bytes :         6         6         0         0         0         0
   Times :   0:00:00   0:00:00                       0:00:00   0:00:00


   Speed :                 103 Bytes/sec.
   Speed :               0.005 MegaBytes/min.
   Ended : 01 January 2016 y. 00:00:01
```


* Start `diskpart.exe`, pipe strings to its StandardInput and capture its output. `Diskpart.exe` will accept piped strings as if they were typed in the interactive session and list all disks and volumes on the PC.

Note that running `diskpart` requires already elevated PowerShell console. Otherwise, you will recieve elevation request and `diskpart` will run, however, no strings would be piped to it.
```
$Result = 'list disk', 'list volume' | Start-ConsoleProcess -FilePath diskpart
$Result.StdOut

Microsoft DiskPart version 6.3.9600

Copyright (C) 1999-2013 Microsoft Corporation.
On computer: HAL9000

DISKPART> 
  Disk ###  Status         Size     Free     Dyn  Gpt
  --------  -------------  -------  -------  ---  ---
  Disk 0    Online          298 GB      0 B         

DISKPART> 
  Volume ###  Ltr  Label        Fs     Type        Size     Status     Info
  ----------  ---  -----------  -----  ----------  -------  ---------  --------
  Volume 0     E                       DVD-ROM         0 B  No Media           
  Volume 1     C   System       NTFS   Partition    100 GB  Healthy    System  
  Volume 2     D   Storage      NTFS   Partition    198 GB  Healthy            

DISKPART> 
```

### [Remove-ComObject](Remove-ComObject.ps1)

Release COM object and remove associated variable. COM object is released using [Marshal.FinalReleaseComObject](https://msdn.microsoft.com/en-us/library/system.runtime.interopservices.marshal.finalreleasecomobject.aspx) method call. Optionally you can force garbage collection.

Features

  * Release COM object and remove associated variable.
  * Full comment-based help and usage examples.

Usage examples

* Removes COM object stored in variable `$Ie` and variable itself.

```powershell
# Create Internet Explorer COM object
$Ie = New-Object -ComObject InternetExplorer.Application

# ... do stuff ...

# Remove Internet Explorer COM object
Remove-ComObject -Name Ie
```

## Scripts

### [New-GitSvnAuthorsFile](New-GitSvnAuthorsFile.ps1)

Generate authors file for one or more SVN repositories to assist SVN to Git migrations.	Can map SVN authors to domain accounts and get full names and emails from Active Directory. Requires Subversion binaries and [Get-SvnAuthor](#get-svnauthor) function.

Features

  * Generate authors file for one or more SVN repositories
  * Map SVN authors to domain accounts and get full names and emails from Active Directory
  * Full comment-based help and usage examples.

Usage example

* Create authors file for SVN repository `http://svnserver/svn/project`. New authors file will be created as `c:\authors.txt`

```powershell
'http://svnserver/svn/project' | New-GitSvnAuthorsFile -Path c:\authors.txt

```

## Modules

### [PsIniParser](PsIniParser)

This module allows to import, export and convert INI files (and strings) to hashtables (or objects) and vice versa. You can specify various parsing options (INI files are not standardized), or use specific encoding while reading a file.

Features

  * Provided cmdlets mimic native PowerShell ones _(Import-\*, Export-\*, ConvertTo-\*, ConvertFrom-\*)_
  * Uses actively supported [INI File Parser](https://github.com/rickyah/ini-parser) by Ricardo Amores Hernández.
  * Highly configurable, can read and write non-standard INI files
  * Supports any available .Net encoding for reading and writing files (unlike Windows' native functions and their PInvoke wrappers that have [very limited Unicode support](http://www.siao2.com/2006/09/15/754992.aspx).)
  * Full comment-based help and usage examples.

Usage examples:

###### Import
* Single file:
```powershell
'C:\Windows\System.ini' | Import-Ini
```
* Multiple files:
```powershell
'C:\Windows\System.ini', 'C:\Windows\Win.ini' | Import-Ini
```
* All INI files in the directory:
```powershell
'C:\Windows\*.ini' | Get-ChildItem | Import-Ini
```
* Single file with non-standard structure:
```
{Section}
Key@Value
%Comment
```

```powershell
'Weird.ini' | Import-Ini -CommentStrings '%' -SectionStartChar '{' -SectionEndChar '}' -KeyValueAssigmentChar '@'
```  

###### Convert from
* INI string:
```powershell
"[Section]`nKey=Value" | ConvertFrom-Ini
```
* Multiple INI strings:
```powershell
"[Section]`nKeyA=ValueA", "[SectionB]`nKeyB=ValueB" | ConvertFrom-Ini
```
* Single INI string with non-standard structure:
```powershell
"{Section}`nKey@Value`n%Comment" | ConvertFrom-Ini -CommentStrings '%' -SectionStartChar '{' -SectionEndChar '}' -KeyValueAssigmentChar '@'
```

###### Export
* Hashtable to INI file:
```powershell
@{Section = @{Key = 'Value'}} | Export-Ini -Path '.\My.ini'
```
* Merge multiple hashtables to INI file:
```powershell
@{SectionA = @{KeyA = 'ValueA'}}, @{SectionB = @{KeyB = 'ValueB'}} | Export-Ini -Path '.\My.ini'
```
* Hashtable to INI file with non-standard structure:
```powershell
@{Section = @{Key = 'Value'}} | Export-Ini -Path '.\My.ini' -SectionStartChar '{' -SectionEndChar '}' -KeyValueAssigmentChar '@'
```

###### Convert to
* Hashtable to INI string:
```powershell
@{Section = @{Key = 'Value'}} | ConvertTo-Ini
```
* Multiple hashtables to INI string(s):
```powershell
@{SectionA = @{KeyA = 'ValueA'}}, @{SectionB = @{KeyB = 'ValueB'}} | ConvertTo-Ini
@{SectionA = @{KeyA = 'ValueA'}}, @{SectionB = @{KeyB = 'ValueB'}} | ConvertTo-Ini -Merge
```
* Hashtable to INI string with non-standard structure:
```powershell
@{Section = @{Key = 'Value'}} | ConvertTo-Ini -SectionStartChar '{' -SectionEndChar '}' -KeyValueAssigmentChar '@'
```
