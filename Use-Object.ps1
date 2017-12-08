<#
.Synopsis
    PowerShell-style version of C# 'using' statement.
    This function will take care of disposing .NET and releasing COM objects for you.

.Description
    PowerShell-style version of C# 'using' statement.

    I felt that C# syntax is no quite fit for PowerShell, so I've made a 'pipelined' version.

    The object is passed via the pipeline and scriptblock is passed as parameter.
    The object is available to the scriptblock via $_ variable, similarly to 'ForEach-Obect'.

    More details here: https://beatcracker.wordpress.com/2017/12/09/yet-another-using-statement/

.Parameter ScriptBlock
    Scriptblock to execute, use '$_' to access object.

.Example
    New-Object -TypeName System.IO.StreamWriter -ArgumentList 'c:\foo.txt' | Use-Object {$_.WriteLine('BAR')}

    Use StreamWriter to write text to file. Stream will be disposed and closed after scriptblock is executed.

.Example
    New-Object -ComObject InternetExplorer.Application | Use-Object {
        $_.Visible = $true
        $_.navigate('https://bing.com')
        Start-Sleep -Seconds 10
        $_.Quit()
    }

    Use Internet Explorer to show website and release IE COM object afterwards.
#>
filter Use-Object {
    Param (
        [ValidateNotNullOrEmpty()]
        [scriptblock]$ScriptBlock
    )

    try {
        . $ScriptBlock
    }
    finally {
        if ($_ -is [System.IDisposable]) {
            'Disposing: {0}' -f $_.GetType().Name | Write-Verbose
            if ($null -eq $DisposableObject.psbase) {
                $_.Dispose()
            }
            else {
                $_.psbase.Dispose()
            }
        }
        elseif ($_ -is [System.__ComObject]) {
            Write-Verbose 'Releasing COM object'
            if ([System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($_)) {
                Write-Error 'Failed to release COM object!'
            }
        }
    }
}