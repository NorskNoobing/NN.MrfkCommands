function Get-MrfkComputerInfo {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory,ParameterSetName="Get computer by hostname",ValueFromPipeline,Position=0)]
        [string]$Hostname,
        [Parameter(Mandatory,ParameterSetName="Get computers by username")][string]$Username,
        [string]$MECMNameSpace = "root/SMS/site_PS1",
        [string]$MECMHost = "sccm-ps.intern.mrfylke.no",
        [string]$DC = "dc01.intern.mrfylke.no"
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
        $CimSession = New-CimSession @splat
    }

    process {
        $ComputerExportArr = New-Object -TypeName System.Collections.ArrayList
        $Notes = New-Object -TypeName System.Collections.ArrayList

        switch ($PsCmdlet.ParameterSetName) {
            "Get computer by hostname" {
                [array]$HostnameArr = $Hostname
            }
            "Get computers by username" {
                $splat = @{
                    "Query" = "Select * from SMS_R_System where LastLogonUserName = `"$Username`""
                    "Namespace" = $MECMNameSpace
                    "CimSession" = $CimSession
                }
                [array]$HostnameArr = (Get-CimInstance @splat).Name
            }
        }

        $HostnameArr.ForEach({
            $splat = @{
                "Query" = "Select * from SMS_R_System where name = `"$_`""
                "Namespace" = $MECMNameSpace
                "CimSession" = $CimSession
            }
            $MecmComputer = Get-CimInstance @splat
            
            if ($MecmComputer) {
                $splat = @{
                    "Query" = @"
select distinct SMS_G_System_PROCESSOR.*
from SMS_R_System
inner join SMS_G_System_PROCESSOR
on SMS_G_System_PROCESSOR.ResourceID = SMS_R_System.ResourceId
where ResourceId = $($MecmComputer.ResourceId)
"@
                    "Namespace" = $MECMNameSpace
                    "CimSession" = $CimSession
                }
                $CPUInfo = Get-CimInstance @splat
                
                $splat = @{
                    "Query" = @"
Select * from SMS_G_System_Computer_System_Product 
where ResourceId = $($MecmComputer.ResourceId)
"@
                    "Namespace" = $MECMNameSpace
                    "CimSession" = $CimSession
                }
                $ModelInfo = Get-CimInstance @splat
            } else {
                $Notes.Add("Couldn't find `"$Hostname`" in MECM.")
            }

            try {
                $ADComputer = Get-ADComputer $_
            
                $splat = @{
                    "Filter" = {objectclass -eq "msFVE-RecoveryInformation"}
                    "SearchBase" = $ADComputer.DistinguishedName
                    "Properties" = "msFVE-RecoveryPassword"
                    "Credential" = Get-MrfkAdmCreds
                    "Server" = $DC
                }
                $BitlockerRecoveryKeys = (Get-ADObject @splat)."msFVE-RecoveryPassword"
            }
            catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
                $Notes.Add("Couldn't find `"$Hostname`" in AD.")
            }
            
            
            if ($MecmComputer.agentname) {
                $HeartbeatIndex = $MecmComputer.agentname.IndexOf("Heartbeat Discovery")
                $LastHeartbeat = $MecmComputer.agenttime[$HeartbeatIndex]
            }
            
            $null = $ComputerExportArr.Add(
                [PSCustomObject]@{
                    "Hostname" = $_
                    "LastLoggedOnUser" = $MecmComputer.LastLogonUserName
                    "MACAddresses" = $MecmComputer.MACAddresses
                    "Model" = $ModelInfo.Name
                    "CPUName" = $CPUInfo.Name
                    "SN" = $ModelInfo.IdentifyingNumber
                    "LastHeartbeat" = $LastHeartbeat
                    "BitlockerRecoveryKeys" = $BitlockerRecoveryKeys
                    "Notes" = $Notes
                }
            )
        })
        $ComputerExportArr
    }
}