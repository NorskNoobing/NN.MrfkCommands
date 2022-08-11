# "-Enabled $true" and "-ChangePasswordAtLogon $true" doesn't seem to be working
function New-LocalAdmin {
    param (
        [Parameter(Mandatory=$true)][string]$username,
        [Parameter(Mandatory=$true)][string]$hostname
    )
    $admUsername = "adm_$username"
    $UPN = "$admUsername@intern.mrfylke.no"

    #Create user in AD
    if (!(Get-ADUser -filter {UserPrincipalName -eq $UPN})) {
        $result = Get-ADuser $username
        $firstName = $result.GivenName
        $lastName = $result.Surname
        
        New-ADUser -name $admUsername -Path "OU=Admin brukere,OU=\#Tilgangskontroll - Drift,OU=SADM,OU=MRFYLKE,DC=intern,DC=mrfylke,DC=no" -UserPrincipalName $UPN -Description "Lokaladmin på maskin $hostname" -GivenName "$firstName" -Surname "$lastName" -DisplayName "$firstName $lastName (Admin)" -AccountPassword (Read-Host -AsSecureString) -ChangePasswordAtLogon $true -Enabled $true
    } else {
        "UPN $UPN already exists"
    }

    #Add ADUser to admin role on the given hostname
    $Ping = (Test-Connection $hostname -Count 1).Status
    if ($Ping -eq "Success") {
        Invoke-Command -ComputerName $hostname -ScriptBlock {Add-LocalGroupMember -Group Administratorer -Member "intern\$admUsername"}
    } else {
        "Ping status: $Ping"
    }
}