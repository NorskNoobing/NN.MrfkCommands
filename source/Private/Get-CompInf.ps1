function Get-CompInf {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)][array]$allMecmComputers,
        [Parameter(Mandatory)][array]$allCpuInfo,
        [Parameter(Mandatory)][array]$allAdComputers,
        [Parameter(Mandatory)][array]$allModelInfo,
        [Parameter(Mandatory)][string]$hostname
    )
    
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
            catch [System.Management.Automation.ParameterBindingException] {
                $null = $notes.Add("Can't convert parameter with name `"$($mecmComputer.LastLogonUserName)`" and type 'System.Object[]' to the type 'Microsoft.ActiveDirectory.Management.ADUser'")
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
    }
}