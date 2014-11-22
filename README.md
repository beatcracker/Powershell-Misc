#Miscellaneous Powershell goodies

##Table of Contents

- [How to use](#How-To)
  - [In PowerShell console\PowerShell ISE script pane](#In PowerShell console\PowerShell ISE script pane)
  - [In your own script](#In your own script)
- [Functions](#Functions)
  - [Get-TerminologyTranslation](#Get-TerminologyTranslation.ps1)
  - [Split-CommandLine](#Split-CommandLine.ps1)

####How to use

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
. $ScriptDir\Get-TerminologyTranslation.ps1
Get-TerminologyTranslation -Text 'Control Panel' -From 'en-us' -To 'ru-ru' -Source Terms

```
####Functions

#####Get-TerminologyTranslation.ps1

Enables user to look up terminology translations and user-interface translations from actual Microsoft products via [Microsoft Terminology Service API](http://www.microsoft.com/Language/en-US/Microsoft-Terminology-API.aspx). For details see [ Terminology Service API SDK PDF](http://download.microsoft.com/download/1/5/D/15D3DDC6-7403-4366-BE99-AF5247ADEF1C/Microsoft-Terminology-API-SDK.pdf).

Features

  * Any-to-any language translation searches, e.g. Japanese to/from French or any other language combination.
  * Filter searches with string case and hotkey sensitivity.
  * Filter searches by product name and version.
  * Get list of languages supported by the Terminology Service API.
  * Get list of products supported by the Terminology Service API.
  * Full comment-based help and usage examples.

#####Split-CommandLine.ps1

This is the Cmdlet version of the code from the article [PowerShell and external commands done right](http://edgylogic.com/blog/powershell-and-external-commands-done-right). It can parse command-line arguments using Win32 API function [CommandLineToArgvW](http://msdn.microsoft.com/en-us/library/windows/desktop/bb776391.aspx).

Features

  * Parse arbitrary command-line, or if none specified, the command-line of the current PowerShell host.
  * Full comment-based help and usage examples.
