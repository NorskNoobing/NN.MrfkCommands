# "-Enabled $true" and "-ChangePasswordAtLogon $true" doesn't seem to be working
function New-LocalAdmin {
    param (
        [Parameter(Mandatory=$true)][string]$username,
        [string]$hostname
    )
    $admUsername = "adm_$username"
    $UPN = "$admUsername@intern.mrfylke.no"
    if (!(Get-ADUser -Credential $(Get-AdmCreds) -filter {UserPrincipalName -eq $UPN})) {
        Import-Module 'C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1'
        Set-Location ps1:
        
        $result = Get-ADuser $username
        $firstName = $result.GivenName
        $lastName = $result.Surname

        if (!$hostname) {
            #FIX: Need to use adm creds
            $hostname = (Get-CimInstance -ComputerName sccm-ps.intern.mrfylke.no -Namespace root/SMS/site_PS1 -Query "select Name from sms_r_system where LastLogonUserName='$username'").Name[0]
            <#
                $var = Invoke-Command -Credential $(Get-AdmCreds) -ComputerName "wintools04" -ScriptBlock {
                    (Get-CimInstance -ComputerName sccm-ps.intern.mrfylke.no -Namespace root/SMS/site_PS1 -Query "select Name from sms_r_system where LastLogonUserName='$username'").Name[0]
                }

                Get-CimInstance: WinRM cannot process the request. The following error with errorcode 0x8009030e occurred while using Kerberos authentication: A specified logon session does not exist. It may already have been terminated.
                Possible causes are:
                -The user name or password specified are invalid.
                -Kerberos is used when no authentication method and no user name are specified.
                -Kerberos accepts domain user names, but not local user names.
                -The Service Principal Name (SPN) for the remote computer name and port does not exist.
                -The client and remote computers are in different domains and there is no trust between the two domains.
                After checking for the above issues, try the following:
                -Check the Event Viewer for events related to authentication.
                -Change the authentication method; add the destination computer to the WinRM TrustedHosts configuration setting or use HTTPS transport.
                Note that computers in the TrustedHosts list might not be authenticated.
                -For more information about WinRM configuration, run the following command: winrm help config.
                InvalidOperation: Cannot index into a null array.
            #>
        }
        
        New-ADUser -Credential $(Get-AdmCreds) -name $admUsername -Path "OU=Admin brukere,OU=\#Tilgangskontroll - Drift,OU=SADM,OU=MRFYLKE,DC=intern,DC=mrfylke,DC=no" -UserPrincipalName $UPN -Description "Lokaladmin på maskin $hostname" -GivenName "$firstName" -Surname "$lastName" -DisplayName "$firstName $lastName (Admin)" -AccountPassword (Read-Host -AsSecureString) -ChangePasswordAtLogon $true -Enabled $true
        $Ping = (Test-Connection $hostname -Count 1).Status
        if ($Ping -eq "Success") {
            Invoke-Command -Credential $(Get-AdmCreds) -ComputerName $hostname -ScriptBlock {Add-LocalGroupMember -Group Administratorer -Member "intern\$admUsername"}
        } else {
            "Ping status: $Ping"
        }
    } else {
        "UPN $UPN already exists"
    }
}