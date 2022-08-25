function Install-RSAT {
    param (
        [switch]$WUServerBypass,
        [string]$ModulePathPassthrough
    )

    try {
        Get-ADUser -Filter "Name -eq 0" | Out-Null
    }
    catch [System.Management.Automation.CommandNotFoundException] {
        try {
            gsudo
        }
        catch [System.Management.Automation.CommandNotFoundException] {
            throw "Gsudo needs to be installed before running this command. Please run Install-gsudo"
        }

        Invoke-gsudo {
            Import-Module "$using:ModulePathPassthrough"

            if ($using:WUServerBypass) {
                Invoke-WUServerBypass
            }
        
            $RSATNames = (Get-WindowsCapability -Name RSAT* -online | Where-Object State -NotLike 'Installed').Name 
            
            foreach ($item in $RSATNames) {
                Get-WindowsCapability -Name $item -Online | Add-WindowsCapability -Online
            }
        }
    }
}