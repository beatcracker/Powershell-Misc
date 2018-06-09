<#
.Synopsis
    Creates clustered MSMQ role.

.Description
    Creates clustered MSMQ role with correct group type and dependencies.
    Can optionally add services to the created group.

.Parameter Name
    Role name.

.Parameter Disk
    Cluster disk name to use for shared storage.
    Must exist and be available.

.Parameter IpAddress
    Array of IP addresses that should be added to the group.

.Parameter Services
    Array of Windows services that should be added to the group.

.Parameter Cluster
    Cluster name to operate on. Netbios/FQDN.
    If the input for this parameter is omitted, then the all operations are run on the local cluster.

.Parameter Network
    Cluster network name.
    If the input for this parameter is omitted, first available network will be used.

.Parameter Start
    Start role group after it's been created.

.Example
    Add-ClusterMsmqRole -Name 'MSMQ' -Disk 'Cluster Disk 1' -IpAddress '10.20.30.40' -Start

    Create new MSMQ role with network name 'MSMQ' and IP address '10.20.30.40' using 'Cluster Disk 1' for shared storage.
    Start 'MSMQ' group after it's been created.

.Example
    Add-ClusterMsmqRole -Name 'MSMQ' -Disk 'Cluster Disk 1' -IpAddress '10.20.30.40' -Service 'SomeService'

    Create new MSMQ role with network name 'MSMQ' and IP address '10.20.30.40' using 'Cluster Disk 1' for shared storage.
    Add windows service 'SomeService' to 'MSMQ' group. Do not start 'MSMQ' group.
#>
function Add-ClusterMsmqRole {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string]$Name,

        [Parameter(Mandatory = $true, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [string]$Disk,

        [Parameter(Mandatory = $true, Position = 2)]
        [ValidateNotNullOrEmpty()]
        [ipaddress[]]$IpAddress,

        [Parameter(Position = 3)]
        [ValidateNotNullOrEmpty()]
        [string[]]$Service,

        [Parameter(Position = 4)]        
        [ValidateScript({
            Get-Cluster -Name $_
        })]
        [string]$Cluster = '.',

        [Parameter(Position = 5)]
        [string]$Network,

        [Parameter(Position = 6)]
        [switch]$Start
    )

    Begin {
        # Set default cluster name
        $PSDefaultParameterValues = @{
            '*-Cluster*:Cluster' = $Cluster
        }
    }

    End {
        # Using $Network = if ($Network){...} breaks pipeline somehow
        if ($Network) {
            [Microsoft.FailoverClusters.PowerShell.ClusterNetwork]$Network = Get-ClusterNetwork -Name $Network
        } else {
            # Using Select-Object stops pipeline too...
            [Microsoft.FailoverClusters.PowerShell.ClusterNetwork]$Network = @(Get-ClusterNetwork)[0]
        }
        Write-Verbose "[*] Cluster network: $Network"

        [Microsoft.FailoverClusters.PowerShell.ClusterResource]$Disk = Get-ClusterResource -Name $Disk |
        Where-Object {
            $_.ResourceType -eq 'Physical Disk' -and
            $_.OwnerGroup -eq 'Available Storage'
        }

        if (-not $Disk) {
            throw "Cluster disk '$Disk' is not available!"
        }

        Write-Verbose "[*] Creating group: $Name"
        $Group = Add-ClusterGroup -Name $Name -GroupType Msmq

        $IpAddressList = foreach ($ip in $IpAddress) {
            $ipName = "IP Address $ip"
            Write-Verbose "[*] Adding '$ipName' to '$Group' using '$Network'"
            $Group |
                Add-ClusterResource -Name $ipName -ResourceType 'IP Address' |
                Set-ClusterParameter -Multiple @{
                    Network    = $Network.Name
                    Address    = "$ip"
                    SubnetMask = $Network.AddressMask
                } | Write-Verbose
        }

        Write-Verbose "[*] Moving cluster disk '$Disk' to group '$Group'"
        $Disk | Move-ClusterResource -Group $Group.Name | Write-Verbose

        Write-Verbose "[*] Adding '$Name (Network Name)' to group '$Group'"
        $NetworkName = $Group | Add-ClusterResource -Name $Name -ResourceType 'Network Name'

        Write-Verbose "[*] Setting 'DnsName' on '$NetworkName (Network Name)'"
        $NetworkName | Set-ClusterParameter -Name 'DnsName' -Value $Name | Write-Verbose

        $NetworkNameDep = (
            $IpAddressList | ForEach-Object {"([$_])"}
        ) -join ' and '
    
        Write-Verbose "[*] Adding dependencies to '$NetworkName (Network Name)': $NetworkNameDep"
        $NetworkName | Set-ClusterResourceDependency -Dependency $NetworkNameDep | Write-Verbose

        $MsmqName = "$Name-MSMQ"
        Write-Verbose "[*] Adding resource of type 'Msmq' to group '$Group': $MsmqName"
        $Msmq = $Group | Add-ClusterResource -Name $MsmqName -ResourceType Msmq

        $MsmqDep = "([$Disk]) and ([$NetworkName])"
        Write-Verbose "[*] Adding dependencies to '$Msmq (MSMQ)': $MsmqDep"
        $Msmq | Set-ClusterResourceDependency -Dependency $MsmqDep | Write-Verbose

        foreach ($svc in $Service) {
            Write-Verbose "[*] Adding resource of type 'Generic Service' to group '$Group': $svc"
            $Group |
                Add-ClusterResource -Name $svc -ResourceType 'Generic Service' |
                Set-ClusterParameter -Name 'ServiceName' -Value $svc | Write-Verbose
        }

        if ($Start) {
            Write-Verbose "[*] Starting group: $Group"
            $Group | Start-ClusterGroup | Write-Verbose
        }
    }
}