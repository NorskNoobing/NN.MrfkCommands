function Remove-MrfkComputer {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory,ValueFromPipeline,Position=0)][string]$Hostname,
        [string]$MECMNameSpace = "root/SMS/site_PS1",
        [string]$MECMHost = "sccm-ps.intern.mrfylke.no"
    )

    begin {
        try {
            $null = Get-ADUser -Filter "Name -eq 0"
        }
        catch [System.Management.Automation.CommandNotFoundException] {
            Write-Error -ErrorAction Stop -Message @"
Please install RSAT before running this function. You can install RSAT by following this guide:
https://github.com/NorskNoobing/NN.MrfkCommands#prerequisites
"@
        }

        $splat = @{
            "Credential" = Get-MrfkAdmCreds
            "ComputerName" = $MECMHost
            "ErrorAction" = "Stop"
        }
        New-CimSession @splat
    }

    process {
        $MECMComputer = Get-CimInstance -CimSession $CimSession -Namespace $MECMNameSpace -Query @"
Select * from SMS_R_System where name = `"$Hostname`"
"@

        if ($MECMComputer) {
            Remove-CimInstance -CimInstance $MECMComputer
        } else {
            Write-Warning -Message "Couldn't find any computer in MECM with the name `"$Hostname`""
        }

        try {
            $ADComputer = Get-ADComputer $Hostname
        } catch {
            Write-Warning -Message "Couldn't find any computer in AD with the name `"$Hostname`""
        }

        if ($ADComputer) {
            (Get-ADObject -Filter * -SearchBase $ADComputer.DistinguishedName).ForEach({
                if($_.DistinguishedName -ne $ADComputer.DistinguishedName) {
                    Remove-ADObject $_.DistinguishedName -Credential (Get-MrfkAdmCreds) -Confirm:$false
                }
            })

            Remove-ADComputer -Identity $ADComputer.DistinguishedName -Credential (Get-MrfkAdmCreds) -Confirm:$false
        }
    }
}