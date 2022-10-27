function New-LocalAdmin {
    param (
        [Parameter(Mandatory)][string]$username,
        [Parameter(Mandatory)][string]$hostname
    )
    $admUsername = "adm_$username"
    $UPN = "$admUsername@intern.mrfylke.no"

    #Create user in AD
    if (!(Get-ADUser -Credential $(Get-AdmCreds) -filter {UserPrincipalName -eq $UPN})) {
        $result = Get-ADuser -Credential $(Get-AdmCreds) $username
        $firstName = $result.GivenName
        $lastName = $result.Surname
        
        New-ADUser -Credential $(Get-AdmCreds) -name "$firstName $lastName (Admin)" -SamAccountName $admUsername -Path "OU=Lokaladmins,OU=Adminbrukere,DC=intern,DC=mrfylke,DC=no" -UserPrincipalName $UPN -Description "Lokaladmin på maskin $hostname" -GivenName "$firstName" -Surname "$lastName" -DisplayName "$firstName $lastName (Admin)" -AccountPassword ("Molde123!" | ConvertTo-SecureString -AsPlainText) -ChangePasswordAtLogon $true -Enabled $true
    } else {
        Write-Information -MessageData "UPN `"$UPN`" already exists" -InformationAction Continue
    }

    $discordHookPath = "$env:USERPROFILE\.creds\Discord\workHookUri.xml"
    if (!(Test-Path $discordHookPath)) {
        $discordHookDir = $discordHookPath.Substring(0, $discordHookPath.lastIndexOf('\'))
        if (!(Test-Path $discordHookDir)) {
            New-Item -ItemType Directory $discordHookDir | Out-Null
        }
        Read-Host -AsSecureString "Enter Discord webhook uri" | Export-Clixml $discordHookPath
    }

    $taskName = "Localadmin for $admUsername on $hostname"
    $task = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue

    #Add ADUser to admin role on the given hostname
    if (Test-Connection $hostname -Count 1 -Quiet) {
        Invoke-Command -Credential $(Get-AdmCreds) -ComputerName $hostname -ScriptBlock {
            if (Get-LocalGroupMember -Group Administratorer -Member "intern\$using:admUsername") {
                "User `"$using:admUsername`" is already an administrator on `"$using:hostname`""
            } else {
                Add-LocalGroupMember -Group Administratorer -Member "intern\$using:admUsername"
            }
        }

        if ($task) {
            Send-DiscordMessage -uri (Import-Clixml $discordHookPath | ConvertFrom-SecureString -AsPlainText) -message "Task `"$taskName`" has been completed."
            $task | Unregister-ScheduledTask -Confirm:$false
        }
    } elseif (!$task) {
        $module1 = "C:\Users\danhol13\repos\mrfk-commands\Output\mrfk-commands.psm1"
        $module2 = "C:\Users\danhol13\repos\NN.Notifications\Output\NN.Notifications.psm1"
        $action = New-ScheduledTaskAction -Execute "pwsh.exe" -Argument "-WindowStyle Hidden -command `"Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force;Import-Module `"$module1`";Import-Module `"$module2`";New-LocalAdmin -username $username -hostname $hostname`""
        #Run every 5 minutes
        $trigger = New-ScheduledTaskTrigger -Once -At (Get-Date) -RepetitionInterval (New-TimeSpan -Minutes 5)
        $null = Register-ScheduledTask -Action $action -Trigger $trigger -TaskName $taskName -Description "Add `"$admUsername`" as localadmin on `"$hostname`""

        Send-DiscordMessage -uri (Import-Clixml $discordHookPath | ConvertFrom-SecureString -AsPlainText) -message "Task `"$taskName`" has been created."
    }
}