<#
.Synopsis
    Release COM object and remove associated variable.

.Description
    Release COM object and remove associated variable.
    COM object is released using [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject method call.
    Optionally you can force garbage collection.

.Parameter Name
    Array of variables names, that contain COM objects.

.Parameter Force
    Switch. Allows the function to remove a variable even if it is read-only.
    Note, that even using the Force parameter, the function cannot remove a constant.

.Parameter ForceGC
    Switch. Force garbage collection after removing all COM objects and associated variables.

.Example
    Remove-ComObject -Name Ie

    Removes COM object stored in variable $Ie and variable itself.

    Example:

    # Create Internet Explorer COM object
    $Ie = New-Object -ComObject InternetExplorer.Application

    # ... do stuff ...

    # Remove Internet Explorer COM object
    Remove-ComObject -Name Ie

.Example
    'Ie' | Remove-ComObject

    Removes COM object stored in variable $Ie and variable itself.
    Variable name is supplied via pipeline.

.Example
    Remove-ComObject -Name Ie, Excel, Word

    Removes COM objects stored in variables $Ie, $Excel, and $Word and variables themselves.

.Example
    'Ie', 'Excel', 'Word' | Remove-ComObject

    Removes COM objects stored in variables $Ie, $Excel, and $Word and variables themselves.
    Variable names are supplied via pipeline.

.Example
    Remove-ComObject -Name Ie -ForceGC

    Removes COM object stored in variable $Ie and variable itself.
    Force immediate garbage collection after variable and COM object has been removed.

.Example
    Remove-ComObject -Name Ie -Force

    Removes COM object stored in variable $Ie and variable itself.
    Variable removed even if it's ReadOnly.

    Example:

    # Create Internet Explorer COM object
    New-Variable -Name Ie -Option ReadOnly -Value (New-Object -ComObject InternetExplorer.Application)

    # ... do stuff ...

    # Remove Internet Explorer COM object
    Remove-ComObject -Name Ie -Force
#>
function Remove-ComObject
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [string[]]$Name,

        [switch]$Force,

        [switch]$ForceGC
    )

    Process
    {
        foreach($var in $Name)
        {
            Write-Verbose 'Trying to release COM object'
            if
            (
                [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject(
                    [System.__ComObject](Get-Variable -Name $var -Scope 1 -ValueOnly -ErrorAction Stop)
                )
            )
            {
                Write-Error 'Failed to release COM object!'
            }

            Write-Verbose "Removing variable from parent scope: $var"
            Remove-Variable -Name $var -Scope 1 -Force:$Force
        }
    }

    End
    {
        if($ForceGC)
        {
            Write-Verbose 'Forcing garbage collection'
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
    }
}