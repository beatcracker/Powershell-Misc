<#
.Synopsis

Wrapper around Win32 API functions for managing (Group) Managed Service Accounts

.Description

Wrapper around Win32 API functions for managing (Group) Managed Service Accounts.
Allows to test/add/remove (G)MSAs without 'Active Directory' module.

.Parameter Add

Installs an existing Active Directory managed service account on the computer on which the cmdlet is run.

.Parameter AccountPassword

Specifies the account password as a secure string.
This parameter enables you to specify the password of a standalone managed service account that you have provisioned and is ignored for group managed service accounts.
This is required when you are installing a standalone managed service account on a server located on a segmented network (site) with read-only domain controllers (for example, a perimeter network or DMZ).
In this case you should create the standalone managed service account, link it with the appropriate computer account, and assign a well-known password that must be passed when installing the standalone managed service account on the server on the read-only domain controller site with no access to writable domain controllers.

.Parameter PromptForPassword

Indicates that you can enter the password of a standalone managed service account that you have pre-provisioned and ignored for group managed service accounts.
This is required when you are installing a standalone managed service account on a server located on a segmented network (site) with no access to writable domain controllers, but only read-only domain controllers (RODCs) (e.g. perimeter network or DMZ).
In this case you should create the standalone managed service account, link it with the appropriate computer account, and assign a well-known password that must be passed when installing the standalone managed service account on the server on the RODC-only site.

.Parameter Test

Tests whether the specified standalone managed service account (sMSA) or group managed service account (gMSA) exists in the Netlogon store on the specified server.

.Parameter Query

Queries the specified service account from the local computer.
The result indicates whether the account is ready for use, which means it can be authenticated and that it can access the domain using its current credentials.

.Parameter Detailed

Return MSA_INFO_STATE enumeration containing detailed information on (G)MSA state.
See: https://msdn.microsoft.com/en-us/library/windows/desktop/dd894396.aspx

.Parameter Remove

Removes an Active Directory standalone managed service account (MSA) on the computer on which the cmdlet is run.
For group MSAs, the cmdlet removes the group MSA from the cache.
However, if a service is still using the group MSA and the host has permission to retrieve the password, then a new cache entry is created.
The specified MSA must be installed on the computer.

.Parameter ForceRemoveLocal

Indicates that you can remove the account from the local security authority (LSA) if there is no access to a writable domain controller.
This is required if you are uninstalling the MSA from a server that is placed in a segmented network such as a perimeter network with access only to a read-only domain controller.
If you specify this parameter and the server has access to a writable domain controller, the account is also un-linked from the computer account in the directory.


.Parameter AccountName

Specifies the Active Directory MSA to uninstall.
You can identify an MSA by its Security Account Manager (SAM) account name.

.Example

'GMSA_Acount' | Use-ServiceAccount -Add

Install Group Managed Service Account with SAM account name 'GMSA_Account' on the computer on which the cmdlet is run

.Example

Use-ServiceAccount -AccountName 'GMSA_Acount' -Add

Install Group Managed Service Account with SAM account name 'GMSA_Account' on the computer on which the cmdlet is run

.Example

'GMSA_Acount' | Use-ServiceAccount -Test

Test whether the specified standalone managed service account (sMSA) or group managed service account (gMSA) exists in the Netlogon store on the this server.

.Example

'GMSA_Acount' | Use-ServiceAccount -Query

Queries the specified service account from the local computer.

#>

function Use-ServiceAccount
{
    [CmdletBinding(DefaultParameterSetName = 'Add')]
    Param
    (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Add')]
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'AccountPassword')]
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'PromptForPassword')]
        [switch]$Add,

        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'AccountPassword')]
        [string]$AccountPassword,

        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'PromptForPassword')]
        [switch]$PromptForPassword,

        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Test')]
        [switch]$Test,

        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Query')]
        [switch]$Query,

        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Query')]
        [switch]$Detailed,

        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Remove')]
        [switch]$Remove,

        [Parameter(ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Remove')]
        [switch]$ForceRemoveLocal,

        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateLength(1, 15)]
        [string[]]$AccountName
    )

    Begin
    {

        $DllImport = @'

            // Service accounts

            [DllImport("logoncli.dll", CharSet = CharSet.Auto)]
            public static extern uint NetQueryServiceAccount(
                [In] string ServerName,
                [In] string AccountName,
                [In] uint InfoLevel,
                out IntPtr Buffer
            );

            [DllImport("logoncli.dll", CharSet = CharSet.Auto)]
            public static extern uint NetIsServiceAccount(
                string ServerName,
                string AccountName,
                ref bool IsService
            );

            [DllImport("logoncli.dll", CharSet = CharSet.Auto)]
            public static extern uint NetAddServiceAccount(
                string ServerName,
                string AccountName,
                string Reserved,
                int Flags
            );

            [DllImport("logoncli.dll", CharSet = CharSet.Auto)]
            public static extern uint NetRemoveServiceAccount(
                string ServerName,
                string AccountName,
                int Flags
            );

	        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
	        public struct MSA_INFO
	        {
		        public MSA_INFO_STATE State;
	        }

	        [Flags]
	        public enum MSA_INFO_STATE : uint
	        {
		        MsaInfoNotExist = 1u,
		        MsaInfoNotService = 2u,
		        MsaInfoCannotInstall = 3u,
		        MsaInfoCanInstall = 4u,
		        MsaInfoInstalled = 5u
	        }
        

            // FormatMessage
            // https://github.com/PowerShell/PowerShell/blob/master/src/Microsoft.PowerShell.Commands.Diagnostics/CommonUtils.cs

            private const uint FORMAT_MESSAGE_ALLOCATE_BUFFER = 0x00000100;
            private const uint FORMAT_MESSAGE_IGNORE_INSERTS = 0x00000200;
            private const uint FORMAT_MESSAGE_FROM_SYSTEM = 0x00001000;
            private const uint LOAD_LIBRARY_AS_DATAFILE = 0x00000002;
            private const uint FORMAT_MESSAGE_FROM_HMODULE = 0x00000800;

            [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
            private static extern uint FormatMessage(uint dwFlags, IntPtr lpSource,
                uint dwMessageId, uint dwLanguageId,
                [MarshalAs(UnmanagedType.LPWStr)]
                StringBuilder lpBuffer,
                uint nSize, IntPtr Arguments);

            [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
            private static extern IntPtr LoadLibraryEx(
                [MarshalAs(UnmanagedType.LPWStr)] string lpFileName,
                IntPtr hFile,
                uint dwFlags
                );

            [DllImport("kernel32.dll")]
            private static extern bool FreeLibrary(IntPtr hModule);

            public static uint FormatMessageFromModule(uint lastError, string moduleName, out String msg)
            {
                uint formatError = 0;
                msg = String.Empty;
                IntPtr moduleHandle = IntPtr.Zero;

                moduleHandle = LoadLibraryEx(moduleName, IntPtr.Zero, LOAD_LIBRARY_AS_DATAFILE);
                if (moduleHandle == IntPtr.Zero)
                {
                    return (uint)Marshal.GetLastWin32Error();
                }

                try
                {
                    uint dwFormatFlags = FORMAT_MESSAGE_IGNORE_INSERTS | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_FROM_HMODULE;
                    uint LANGID = (uint)System.Globalization.CultureInfo.CurrentUICulture.LCID;

                    StringBuilder outStringBuilder = new StringBuilder(1024);
                    uint nChars = FormatMessage(dwFormatFlags,
                        moduleHandle,
                        lastError,
                        LANGID,
                        outStringBuilder,
                        (uint)outStringBuilder.Capacity,
                        IntPtr.Zero);

                    if (nChars == 0)
                    {
                        formatError = (uint)Marshal.GetLastWin32Error();
                    }
                    else
                    {
                        msg = outStringBuilder.ToString();
                        if (msg.EndsWith(Environment.NewLine, StringComparison.Ordinal))
                        {
                            msg = msg.Substring(0, msg.Length - 2);
                        }
                    }
                }
                finally
                {
                    FreeLibrary(moduleHandle);
                }
                return formatError;
            }
'@
        Add-Type -MemberDefinition $DllImport -Name ServiceAccount -Namespace LogonCli -ErrorAction Stop -UsingNamespace 'System.Text'

        function Format-MessageFromModule
        {
            [CmdletBinding()]
            Param
            (
                [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
                [uint32[]]$LastError,

                [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
                [string]$ModuleName
            )

            Process
            {
                foreach ($err in $LastError) {
                    $result = $null
                    if (! ($ret = [LogonCli.ServiceAccount]::FormatMessageFromModule($err, $ModuleName, [ref]$result))) {
                        $result
                    } else {
                        Write-Error "Can't get message ID: $err. Error code: $ret"
                    }

                }
            }
        }
    }


    Process
    {
        foreach ($name in $AccountName) {
            Write-Verbose "Using account: $name"

            $result = 0

            if ($Test) {
                Write-Verbose 'Testing account using NetIsServiceAccount'
                if (!($ret = [LogonCli.ServiceAccount]::NetIsServiceAccount($null, $name, [ref]$result))) {
                    $result
                }
            } elseif ($Query) {
                Write-Verbose 'Querying account detail using NetQueryServiceAccount'
                if (!($ret = [LogonCli.ServiceAccount]::NetQueryServiceAccount($null, $name, 0, [ref]$result))) {
                    $result = [System.Runtime.InteropServices.Marshal]::PtrToStructure($result, [System.Type][LogonCli.ServiceAccount+MSA_INFO])

                    if ($Detailed) {
                        Write-Verbose 'Returning detailed result'
                        $result.State
                    } else {
                        if ($result.State -eq [LogonCli.ServiceAccount+MSA_INFO_STATE]::MsaInfoInstalled) {
                            $true
                        } else {
                            switch ($result.State) {
                                ([LogonCli.ServiceAccount+MSA_INFO_STATE]::MsaInfoNotExist) {
                                    Write-Warning "Cannot find Managed Service Account $name in the directory. Verify the Managed Service Account identity and call the cmdlet again."
                                }

                                ([LogonCli.ServiceAccount+MSA_INFO_STATE]::MsaInfoNotService) {
                                    Write-Warning "The $name is not a Managed Service Account. Verify the identity and call the cmdlet again."
                                }

                                ([LogonCli.ServiceAccount+MSA_INFO_STATE]::MsaInfoCannotInstall) {
                                    Write-Warning "Test failed for Managed Service Account $name. If standalone Managed Service Account, the account is linked to another computer object in the Active Directory. If group Managed Service Account, either this computer does not have permission to use the group MSA or this computer does not support all the Kerberos encryption types required for the gMSA. See the MSA operational log for more information."
                                }

                                ([LogonCli.ServiceAccount+MSA_INFO_STATE]::MsaInfoCanInstall) {
                                    Write-Warning "The Managed Service Account $name is not linked with any computer object in the directory."
                                }
                            }

                            $false
                        }
                    }
                }
            } elseif ($Add) {
                Write-Verbose 'Installing account using NetAddServiceAccount'
                if ($PromptForPassword-or $AccountPassword) {
                    if ($PromptForPassword) {
                        $AccountPassword = Read-Host
                    }

                    $ret = [LogonCli.ServiceAccount]::NetAddServiceAccount( $null, $name, $AccountPassword, 2)
                }  else {
                    $ret = [LogonCli.ServiceAccount]::NetAddServiceAccount($null, $name, $null, 1)
                }
            } elseif ($Remove) {
                Write-Verbose 'Uninstalling account using NetRemoveServiceAccount'
                $ret = [LogonCli.ServiceAccount]::NetRemoveServiceAccount($null, $AccountName, 1)

                # fom winnt.h
                # define STATUS_NO_SUCH_DOMAIN 0xC00000DF
                if ((0xC00000DF -eq $ret) -and ($ForceRemoveLocal)) {
                    Write-Verbose 'Can''t contact domain, removing local account data'
                    $ret = [LogonCli.ServiceAccount]::NetRemoveServiceAccount($null, $AccountName, 2)
                }
            }

            if ($ret) {
                Write-Verbose "Returning user-friendly error message for status code: $ret"
                $ret | Format-MessageFromModule -ModuleName 'Ntdll.dll' | Write-Error 
            }
        }
    }
}