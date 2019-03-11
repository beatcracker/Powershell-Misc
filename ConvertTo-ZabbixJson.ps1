<#
.Synopsis
    Convert object to JSON supported by Zabbix low-level discovery.

.Description
    Convert object to JSON supported by Zabbix low-level discovery.
    Object property names will be converted to uppercase, prefixed by #
    and wrapped in curly braces: Name -> {#NAME}

.Parameter InputObject
    Object to be converted. You can supply multiple objects.

.Parameter Compress
    Omits white space and indented formatting in the output string.
    If not set, the ConvertTo-Json defaults are used.

.Example
    [pscustomobject]@{a=1 ; b = 2}, @{c = 3 ; d = 4} | ConvertTo-ZabbixJson

    Converts PSCustomObject and hashtable to Zabbix LLD JSON:

    {
        "data":  [
                    {
                        "{#A}":  1,
                        "{#B}":  2
                    },
                    {
                        "{#D}":  4,
                        "{#C}":  3
                    }
                ]
    }

.Example
    Get-Website | Select-Object name, physicalPath | ConvertTo-ZabbixJson

    Converts website object to Zabbix LLD JSON:

    {
        "data":  [
                    {
                        "{#NAME}":  "Default Web Site",
                        "{#PHYSICALPATH}":  "C:\\inetpub\\wwwroot"
                    }
                ]
    }
#>
function ConvertTo-ZabbixJson {
    [CmdletBinding()]
    Param (
        [Parameter(ValueFromPipeline = $true, Position = 0)]
        $InputObject,

        [Parameter(Position = 1)]
        [switch]$Compress
    )

    Begin {
        $Result = New-Object -TypeName System.Collections.Generic.List[PSCustomObject]
    }

    Process {
        foreach ($item in $InputObject) {
            # if item is hashtable, convert it to PSCustomObject
            $item = [pscustomobject]$item

            if ($InvalidPropertyName = @($item.PsObject.Properties.Name) -notmatch '[0-9A-Z_\.]') {
                throw "Invalid property name: $InvalidPropertyName . Allowed symbols for LLD macro names are: 0-9 , A-Z , _ , ."
            }

            # Build calculated properties for Zabbix LLD object
            $Property = foreach ($name in $item.PsObject.Properties.Name) {
                @{
                    n = '{{#{0}}}' -f $name.ToUpper()
                    e = [scriptblock]::Create(('$_.''{0}''.ToString()' -f $name))
                }
            }

            [void]$Result.Add(
                ($item | Select-Object -Property $Property)
            )
        }
    }

    End {
        if ($Result.Count) {
            @{data = $Result} | ConvertTo-Json -Compress:$Compress
        }
    }
}