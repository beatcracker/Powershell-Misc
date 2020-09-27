<#
.Synopsis
    Write-Host but with ANSI colors!

.Description
    Drop-in Write-Host replacement that uses ANSI escape codes to render colors.
    Allows for colorized output in CI systems.

.Parameter Object
    Objects to display in the host.

.Parameter ForegroundColor
    Specifies the text color. There is no default.

.Parameter BackgroundColor
    Specifies the background color. There is no default.

.Parameter Separator
    Specifies a separator string to insert between objects displayed by the host.

.Parameter NoNewline
    The string representations of the input objects are concatenated to form the output.
    No spaces or newlines are inserted between the output strings.
    No newline is added after the last output string.

.Example
    Write-Host 'Double rainbow!' -ForegroundColor Magenta -BackgroundColor Yellow

.Notes
    Author : beatcracker (https://github.com/beatcracker)
    License: MS-PL (https://opensource.org/licenses/MS-PL)
    Source : https://github.com/beatcracker/Powershell-Misc
#>
function Write-Host {
    [CmdletBinding()]
    Param(
        [Parameter(ValueFromPipeline)]
        [Alias('Msg', 'Message')]
        [System.Object[]]$Object,
        [System.Object]$Separator,
        [System.ConsoleColor]$ForegroundColor,
        [System.ConsoleColor]$BackgroundColor,
        [switch]$NoNewline
    )

    Begin {
        # Map ConsoleColor enum values to ANSI colors
        # https://en.wikipedia.org/wiki/ANSI_escape_code#3/4_bit
        $AnsiColor = @(
            30, 34, 32, 36, 31, 35, 33, 37, 90, 94, 92, 96, 91, 95, 93, 97
        )
        # PS < 6.0 doesn't have `e escape character
        $Esc = [char]27
        $AnsiTemplate = "$Esc[{0}m{1}$Esc[{2}m"
    }

    Process {
        # https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_special_characters#escape-e
        # https://docs.microsoft.com/en-us/powershell/scripting/windows-powershell/wmf/whats-new/console-improvements#vt100-support
        if ($Host.UI.SupportsVirtualTerminal) {
            $Method = if ($NoNewline) { 'Write' } else { 'WriteLine' }
            $Output = if ($Separator) { $Object -join $Separator } else { "$Object" }

            # Splitting by regex ensures that this will work on files from Windows/Linux/macOS
            # Get-Content .\Foobar.txt -Raw | Write-Host -ForegroundColor Red
            foreach ($item in $Output -split '\r\n|\r|\n') {
                if ("$BackgroundColor") {
                    $item = $AnsiTemplate -f ($AnsiColor[$BackgroundColor.value__] + 10), $item, 49
                }
                if ("$ForegroundColor") {
                    $item = $AnsiTemplate -f $AnsiColor[$ForegroundColor.value__], $item, 39
                }

                [System.Console]::$Method($item)
            }
        }
        else {
            Microsoft.PowerShell.Utility\Write-Host @PSBoundParameters
        }
    }
}