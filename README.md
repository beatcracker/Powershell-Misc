#Miscellaneous PowerShell goodies

####Table of Contents

- [How to use Functions\Scripts](#how-to-use-functionsscripts)
  - [In PowerShell console\PowerShell ISE script pane](#in-powershell-consolepowershell-ise-script-pane)
  - [In your own script](#in-your-own-script)
- [How to use Modules](#how-to-use-modules)
- [Functions](#functions)
  - [Get-TerminologyTranslation](#get-terminologytranslationps1)
  - [Split-CommandLine](#split-commandlineps1)
  - [Import-Component](#import-componentps1)
  - [Get-SvnAuthor](#get-svnauthorps1)
- [Modules](#modules)
  - [PsIniParser](#psiniparser)
- [Scripts](#scripts)
  - [New-GitSvnAuthorsFile](#new-gitsvnauthorsfileps1)

####How to use Functions\Scripts

The best way to use provided functions is a [dot-sourcing](http://ss64.com/ps/source.html).

Dot-sourcing runs a script in the current scope so that any functions, aliases, and variables that the script creates are added to the current scope. 

#####In PowerShell console\PowerShell ISE script pane

*Note the space between first dot and path!*

```powershell
. c:\scripts\Get-TerminologyTranslation.ps1
```

Or, navigate to the folder, where you downloaded script

```
cd c:\scripts
```

And then type:

```powershell
. .\Get-TerminologyTranslation.ps1
```

To verify, that script is loaded, try to view help for the loaded function. The function name is usually the same as script name (without extension), if not stated otherwise in readme.

```powershell
Get-Help -Full Get-TerminologyTranslation
```

Congratulations, you can now make use of newly added function!

```powershell
Get-TerminologyTranslation -Text 'Control Panel' -From 'en-us' -To 'ru-ru' -Source Terms

```

#####In your own script

*Note the space between first dot and path!*

Dot-source from arbitrary location before calling the function:

```powershell
. c:\scripts\Get-TerminologyTranslation.ps1
Get-TerminologyTranslation -Text 'Control Panel' -From 'en-us' -To 'ru-ru' -Source Terms
```

Or put it alongside with your script, get path to the script folder programmatically, and then dot-source:

```powershell
$ScriptDir = Split-Path $script:MyInvocation.MyCommand.Path
. (Join-Path -Path $ScriptDir -ChildPath 'Get-TerminologyTranslation.ps1')
Get-TerminologyTranslation -Text 'Control Panel' -From 'en-us' -To 'ru-ru' -Source Terms

```

####How to use Modules
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
Get-Module -ListAvailable`
```
* To import available module, use:
```powershell
Import-Module -Name 'ModuleName'
```

Alternatively, you can import module from any location:
```powershell
Import-Module -Name 'X:\Path\to\module_folder'
```

####Functions

#####`Get-TerminologyTranslation.ps1`

Enables user to look up terminology translations and user-interface translations from actual Microsoft products via [Microsoft Terminology Service API](http://www.microsoft.com/Language/en-US/Microsoft-Terminology-API.aspx). For details see [ Terminology Service API SDK PDF](http://download.microsoft.com/download/1/5/D/15D3DDC6-7403-4366-BE99-AF5247ADEF1C/Microsoft-Terminology-API-SDK.pdf).

Features

  * Any-to-any language translation searches, e.g. Japanese to/from French or any other language combination.
  * Filter searches with string case and hotkey sensitivity.
  * Filter searches by product name and version.
  * Get list of languages supported by the Terminology Service API.
  * Get list of products supported by the Terminology Service API.
  * Full comment-based help and usage examples.

#####`Split-CommandLine.ps1`

This is the Cmdlet version of the code from the article [PowerShell and external commands done right](http://edgylogic.com/blog/powershell-and-external-commands-done-right). It can parse command-line arguments using Win32 API function [CommandLineToArgvW](http://msdn.microsoft.com/en-us/library/windows/desktop/bb776391.aspx).

Features

  * Parse arbitrary command-line, or if none specified, the command-line of the current PowerShell host.
  * Full comment-based help and usage examples.

#####`Import-Component.ps1`

Bulk-import from folder any component, supported by PowerShell (script, module, source code, .Net assembly).

Features

  * Supported components:
    * Script (.ps1) - imported using [Dot-Sourcing](http://ss64.com/ps/source.html).
    * Module (.psm1) - imported using [Import-Module](http://technet.microsoft.com/en-us/library/hh849725.aspx) cmdlet
    * Source code (.cs, .vb, .js) - imported using [Add-Type](http://technet.microsoft.com/en-us/library/hh849914.aspx) cmdlet
    * .Net assembly (.dll) - imported using [Add-Type](http://technet.microsoft.com/en-us/library/hh849914.aspx) cmdlet
  * Full comment-based help and usage examples.

__WARNING: To import .PS1 scripts this function itself has to be dot-sourced!__ Example:

*Note the space between first dot and function name!*
```powershell
. Import-Component 'C:\PsLib'
```

#####`Get-SvnAuthor.ps1`

Get list of unique commit authors in one or more SVN repositories. Requires Subversion binaries. Can be used to create authors file for SVN to Git migrations.

Features

  * Get list of unique commit authors in one or more SVN repositories.
  * Full comment-based help and usage examples.

####Modules

#####`PsIniParser`

This module allows to import, export and convert INI files (and strings) to hashtables (or objects) and vice versa. You can specify various parsing options (INI files are not standardized), or use specific encoding while reading a file.

Features

  * Provided cmdlets mimic native PowerShell ones _(Import-\*, Export-\*, ConvertTo-\*, ConvertFrom-\*)_
  * Uses actively supported [INI File Parser](https://github.com/rickyah/ini-parser) by Ricardo Amores Hernández.
  * Highly configurable, can read and write non-standard INI files
  * Supports any available .Net encoding for reading and writing files (unlike Windows' native functions and their PInvoke wrappers that have [very limited Unicode support](http://www.siao2.com/2006/09/15/754992.aspx).)
  * Full comment-based help and usage examples.

Usage examples:

######Import
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

######Convert from
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

######Export
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

######Convert to
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
####Scripts

#####`New-GitSvnAuthorsFile.ps1`

Generate authors file for one or more SVN repositories to assist SVN to Git migrations.	Can map SVN authors to domain accounts and get full names and emails from Active Directory. Requires Subversion binaries and [Get-SvnAuthor](#get-svnauthorps1) function.

Features

  * Generate authors file for one or more SVN repositories
  * Map SVN authors to domain accounts and get full names and emails from Active Directory
  * Full comment-based help and usage examples.
