#Miscellaneous Powershell goodies

##Table of Contents

- [Get-TerminologyTranslation](####Get-TerminologyTranslation.ps1)
- [Split-CommandLine](####Split-CommandLine.ps1)

####Get-TerminologyTranslation.ps1

Enables user to look up terminology translations and user-interface translations from actual Microsoft products via [Microsoft Terminology Service API](http://www.microsoft.com/Language/en-US/Microsoft-Terminology-API.aspx). For details see [ Terminology Service API SDK PDF](http://download.microsoft.com/download/1/5/D/15D3DDC6-7403-4366-BE99-AF5247ADEF1C/Microsoft-Terminology-API-SDK.pdf).

Features

  * Any-to-any language translation searches, e.g. Japanese to/from French or any other language combination.
  * Filter searches with string case and hotkey sensitivity.
  * Filter searches by product name and version.
  * Get list of languages supported by the Terminology Service API.
  * Get list of products supported by the Terminology Service API.
  * Full comment-based help and usage examples.

####Split-CommandLine.ps1

This is the Cmdlet version of the code from the article [PowerShell and external commands done right](http://edgylogic.com/blog/powershell-and-external-commands-done-right). It can parse command-line arguments using Win32 API function [CommandLineToArgvW](http://msdn.microsoft.com/en-us/library/windows/desktop/bb776391.aspx).

Features

  * Parse arbitrary command-line, or if none specified, the command-line of the current PowerShell host.
  * Full comment-based help and usage examples.