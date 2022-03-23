# "-Enabled $true" and "-ChangePasswordAtLogon $true" doesn't seem to be working
function New-LocalAdmin {
    param (
        [Parameter(Mandatory=$true)][string]$username,
        [string]$hostname
    )
    $admUsername = "adm_$username"
    $UPN = "$admUsername@intern.mrfylke.no"
    if (!(Get-ADUser -filter {UserPrincipalName -eq $UPN})) {
        Import-Module 'C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1'
        Set-Location ps1:
        
        $result = Get-ADuser $username
        $firstName = $result.GivenName
        $lastName = $result.Surname

        if (!$hostname) {
            $hostname = (Get-CimInstance -ComputerName sccm-ps.intern.mrfylke.no -Namespace root/SMS/site_PS1 -Query "select Name from sms_r_system where LastLogonUserName='$username'").Name[0]
        }
        
        New-ADUser -name $admUsername -Path "OU=Admin brukere,OU=\#Tilgangskontroll - Drift,OU=SADM,OU=MRFYLKE,DC=intern,DC=mrfylke,DC=no" -UserPrincipalName $UPN -Description "Lokaladmin på maskin $hostname" -GivenName "$firstName" -Surname "$lastName" -DisplayName "$firstName $lastName (Admin)" -AccountPassword (Read-Host -AsSecureString) -ChangePasswordAtLogon $true -Enabled $true
        $Ping = (Test-Connection $hostname -Count 1).Status
        if ($Ping -eq "Success") {
            Invoke-Command -ComputerName $hostname -ScriptBlock {Add-LocalGroupMember -Group Administratorer -Member "intern\$admUsername"}
        } else {
            "Ping status: $Ping"
        }
    } else {
        "UPN $UPN already exists"
    }
}