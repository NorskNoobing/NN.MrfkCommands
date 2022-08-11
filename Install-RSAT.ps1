function Install-RSAT {
    param (
        [switch]$WUServerBypass
    )

    try {
        Get-ADUser -Filter "Name -eq 0" | Out-Null
    }
    catch [System.Management.Automation.CommandNotFoundException] {
        #todo: launch the following code with admin privledges
        if ($using:WUServerBypass) {
            Invoke-WUServerBypass
        }
    
        $RSATNames = (Get-WindowsCapability -Name RSAT* -online | Where State -NotLike 'Installed').Name
        
        foreach ($item in $RSATNames) {
            Get-WindowsCapability -Name $item -Online | Add-WindowsCapability -Online
        }
    }
}