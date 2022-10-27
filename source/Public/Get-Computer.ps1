#Takes approx 130ms per entry when piping inputs into the function
function Get-Computer {
    [CmdletBinding(DefaultParameterSetName="Get computer by hostname")]
    param (
        [Parameter(Mandatory,ValueFromPipeline,Position=0,ParameterSetName="Get computer by hostname")][string]$hostname,
        [Parameter(ParameterSetName="Get computer by hostname")][pscredential]$credential
    )
    
    begin {
        $allCpuInfoQry = @"
select distinct SMS_G_System_PROCESSOR.*
from SMS_R_System
inner join SMS_G_System_PROCESSOR
on SMS_G_System_PROCESSOR.ResourceID = SMS_R_System.ResourceId
"@

        $cimsession = New-CimSession -Credential $credential -ComputerName sccm-ps.intern.mrfylke.no -ErrorAction Stop
        
        $allAdComputers = Get-ADComputer -Filter *
        $allMecmComputers = Get-CimInstance -Query "Select * from SMS_R_System" -Namespace root/SMS/site_PS1 -CimSession $cimsession
        $allCpuInfo = Get-CimInstance -Query $allCpuInfoQry -Namespace root/SMS/site_PS1 -CimSession $cimsession
        $allModelInfo = Get-CimInstance -Query "Select * from SMS_G_System_Computer_System_Product" -Namespace root/SMS/site_PS1 -CimSession $cimsession
    
        $compInfExport = New-Object -TypeName System.Collections.ArrayList
    }
    
    process {
        $notes = New-Object -TypeName System.Collections.ArrayList
        $mecmComputer = $allMecmComputers.where({$_.Name -eq $hostname})
        $cpuInfo = $allCpuInfo.where({$_.SystemName -eq $hostname})
        $adComputer = $allAdComputers.where({$_.Name -eq $hostname})
        $modelInfo = $allModelInfo.where({$_.ResourceID -eq $mecmComputer.ResourceId})

        if ($mecmComputer.LastLogonUserName) {
            try {
                $adUser = Get-ADUser $mecmComputer.LastLogonUserName -Properties DisplayName
            }
            catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
                $null = $notes.Add("Can't find user with the identity `"$($mecmComputer.LastLogonUserName)`"")
            }
        } else {
            $null = $notes.Add("There's no last logged on user")
        }

        if ($mecmComputer.agentname) {
            $heartbeatIndex = $mecmComputer.agentname.IndexOf("Heartbeat Discovery")
            $lastHeartbeat = $mecmComputer.agenttime[$heartbeatIndex]
        }
        
        $IPv4 = $mecmComputer.IPAddresses.ForEach({
            if ($_ -like "*.*") {
                $_
            }
        })

        #Get online status
        try {
            $tnc = Test-Connection $hostname -Count 1 -ErrorAction "Stop"

            if ($tnc.Status -eq "TimedOut") {
                [string]$status = "Offline"
            }
            elseif ($tnc.status -eq "Success") {
                [string]$status = "Online"
            }
        }
        catch [System.Net.NetworkInformation.PingException] {
            [string]$status = "Can't resolve hostname"
        }

        $null = $compInfExport.Add(
            [PSCustomObject]@{
                Hostname          = $mecmComputer.Name
                DisplayName       = $adUser.DisplayName
                Username          = $mecmComputer.LastLogonUserName
                MACAddress        = $mecmComputer.MACAddresses
                Model             = $modelInfo.Name
                SN                = $modelInfo.IdentifyingNumber
                GUID              = $mecmComputer.SMBIOSGUID
                CreatedAt         = $mecmComputer.CreationDate
                DeviceOSBuild     = $mecmComputer.BuildExt
                OS                = $mecmComputer.OperatingSystemNameandVersion
                IPv4              = $IPv4
                "Enabled(AD)"     = $adComputer.Enabled
                Status            = $status
                LastHeartbeat     = $lastHeartbeat
                CpuName           = $cpuInfo.Name
                Notes             = $notes
            }
        )
    }

    end {
        $compInfExport
    }
}