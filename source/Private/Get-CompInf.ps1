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

        if ($mecmComputer.agentname) {
            $heartbeatIndex = $mecmComputer.agentname.IndexOf("Heartbeat Discovery")
            $lastHeartbeat = $mecmComputer.agenttime[$heartbeatIndex]
        }

        [PSCustomObject]@{
            Hostname          = $mecmComputer.Name
            Username          = $mecmComputer.LastLogonUserName
            MACAddress        = $mecmComputer.MACAddresses
            Model             = $modelInfo.Name
            SN                = $modelInfo.IdentifyingNumber
            GUID              = $mecmComputer.SMBIOSGUID
            CreatedAt         = $mecmComputer.CreationDate
            DeviceOSBuild     = $mecmComputer.BuildExt
            OS                = $mecmComputer.OperatingSystemNameandVersion
            "Enabled(AD)"     = $adComputer.Enabled
            LastHeartbeat     = $lastHeartbeat
            CpuName           = $cpuInfo.Name
            Notes             = $notes
        }
    }
}