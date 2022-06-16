<#
    1. Add CollectionName
    2. Call on functions within function to prevent using $var = $null
    3. Set required params and param types
    4. Foreach-Object -Parallel
    5. Add everything to PSCustomObject then output (write to cli or append object to file)
#>
function CompInf {
    param (
        $hostname
    )
    $collectionsarray = New-Object -TypeName System.Collections.ArrayList

    $cmdeviceinfo = get-cmdevice -name $hostname
    $adcomputer = Get-ADComputer $hostname -Properties *
    $collections = Get-CimInstance -ComputerName sccm-ps.intern.mrfylke.no -Namespace root/SMS/site_PS1 -Query "SELECT SMS_Collection.* FROM SMS_FullCollectionMembership, SMS_Collection where name = '$hostname' and SMS_FullCollectionMembership.CollectionID = SMS_Collection.CollectionID"
    # FILTER OUT DUPLICATES IN MECM
    if ($cmdeviceinfo -is [array]) {
        $cmdeviceinfo = $cmdeviceinfo[0]
    }
    $model = ($collections.Name | Select-String "All HP Elite*", "All HP Pro*", "All HP Z*", "All Lenovo Think*") -replace "All ", ""
    $mac = ($cmdeviceinfo.MACAddress) -replace ',', ''
    $sn = $cmdeviceinfo.serialnumber
    $guid = $cmdeviceinfo.SMBIOSGUID
    $IsActive = $cmdeviceinfo.IsActive
    $DeviceOSBuild = $cmdeviceinfo.DeviceOSBuild
    $created = $adcomputer.Created
    $OperatingSystem = $adcomputer.OperatingSystem
    $IPv4Address = $adcomputer.IPv4Address
    $Enabled = $adcomputer.Enabled
    $LastHardwareScan = $cmdeviceinfo.LastHardwareScan
    $LastPolicyRequest = $cmdeviceinfo.LastPolicyRequest
    $LastDDR = $cmdeviceinfo.LastDDR
    $CPUGen = ($collections.Name | Select-String " Gen - ") -replace "All ", ""

    # WIN11 COMPATIBILITY
    $win11incompatcollection = "PS1002E2", "PS1002E3", "PS1002E4", "PS1002E5", "PS1002E8"
    $win11compatcollection = "PS1002EA", "PS1002E9", "PS1002E6", "PS1002E7"
    foreach ($item in $($collections.CollectionID)) {
        if ($win11incompatcollection -contains $item) {
            $win11incompatcounter++
        }
        elseif ($win11compatcollection -contains $item) {
            $win11compatcounter++
        }
    }
    if ($win11incompatcounter -gt 0) {
        $win11compat = $false
    }
    elseif ($win11compatcounter -gt 0) {
        $win11compat = $true
    }
    else {
        $win11compat = "UNKNOWN"
    }

    #Convert collectionlist from string
    $var = $collections | Select-Object CollectionID, Name
    foreach ($item in $var) {
        $item = ("$item") -replace "@{CollectionID=", "" -replace "}", "" -replace ";", "" -replace "Name=", ""
        $collectionsarray.Add($item) | Out-Null
    }

    if (!($Global:displayname -or $Global:username)) {
        $displayname = $null
        $username = $null
        $username = $cmdeviceinfo.username
        $displayname = (Get-ADUser "$username" -Properties DisplayName).DisplayName
    }

    #Get online status
    try {
        $tnc = Test-Connection $hostname -Count 1 -ErrorAction "Stop"

        if ($tnc.Status -eq "TimedOut") {
            $status = "Offline"
        }
        elseif ($tnc.status -eq "Success") {
            $status = "Online"
        }
    }
    catch [System.Net.NetworkInformation.PingException] {
        $status = "Can't resolve hostname"
    }

    $computerinfo = [PSCustomObject]@{
        Hostname          = $hostname
        DisplayName       = $displayname
        Username          = $username
        MACAddress        = $mac
        Model             = $model
        SN                = $sn
        GUID              = $guid
        CreatedAt         = $created
        IsActive          = $IsActive
        DeviceOSBuild     = $DeviceOSBuild
        OS                = $OperatingSystem
        IPv4              = $IPv4Address
        Enabled           = $Enabled
        Collections       = $collectionsarray
        Status            = $status
        LastHardwareScan  = $LastHardwareScan
        LastPolicyRequest = $LastPolicyRequest
        LastDDR           = $LastDDR
        Win11Compatible   = $win11compat
        CPUGeneration     = $CPUGen
    }

    # OUTPUT TYPE
    if ($OutFile) {
        Add-Content -Path $OutFile -Value "$hostname,$displayname,$username,$mac,$model,$sn,$guid,$created,$IsActive,$DeviceOSBuild,$OperatingSystem,$IPv4Address,$Enabled,$collectionsarray,$LastHardwareScan,$LastPolicyRequest,$LastDDR,$win11compat,$CPUGen"
    }
    else {
        $computerinfo
    }
}

function Get-Computer {
    param (
        $hostname, $importpath, $collectionIDs, $OutFile, $displayname, $username, $filter, $inputpath, $collectionNames
    )
    $Global:GetComputer = $null
    Import-Module 'C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1'
    Set-Location ps1:
    $allcomputers = New-Object -TypeName System.Collections.ArrayList
    $allfilteredcomputers = New-Object -TypeName System.Collections.ArrayList

    #
    #
    # - INPUTS
    #
    #

    # XLSX
    if ($importpath) {
        $computerdoc = Import-Excel -Path $importpath
        foreach ($computer in $computerdoc) {
            $hostname = $computer.hostname
            $allcomputers.Add($hostname) | Out-Null
        }
    }
    # HOSTNAME
    elseif ($hostname) {
        $allcomputers.Add($hostname) | Out-Null
    }
    # COLLECTIONID
    elseif ($collectionIDs) {
        foreach ($item in $collectionIDs) {
            $hostnamearray = (Get-CMCollectionMember -CollectionId $item).Name
            foreach ($hostname in $hostnamearray) {
                $allcomputers.Add($hostname) | Out-Null
            }
        }
    }
    # DISPLAYNAME
    elseif ($displayname) {
        $Global:displayname = $displayname
        $username = (Get-ADUser -Filter { DisplayName -like $displayname }).Name
        $Global:username = $username
        $allcomputers = (Get-CimInstance -ComputerName sccm-ps.intern.mrfylke.no -Namespace root/SMS/site_PS1 -Query "select Name from sms_r_system where LastLogonUserName='$username'").Name
    }
    # USERNAME
    elseif ($username) {
        $Global:username = $username
        $Global:displayname = (Get-ADUser $username -Properties DisplayName).DisplayName
        $allcomputers = (Get-CimInstance -ComputerName sccm-ps.intern.mrfylke.no -Namespace root/SMS/site_PS1 -Query "select Name from sms_r_system where LastLogonUserName='$username'").Name
    }

    #
    #
    # - OUTPUT TYPE (cli or file)
    #
    #

    if ($OutFile) {
        $Outpath = "C:\Users\adm_danhol13\Desktop\"
        $OutFile = "$Outpath$OutFile$('.csv')"

        $DoesOutFileExist = Test-Path -Path $OutFile -PathType Leaf
        if (!$DoesOutFileExist) {
            New-Item -Path $OutFile -ItemType File | Out-Null
            Set-Content -Path $OutFile -Value "Hostname,DisplayName,Username,MAC,Model,S/N,GUID,Created,IsActive,DeviceOSBuild,OperatingSystem,IPv4Address,Enabled,Collections,LastHardwareScan,LastPolicyRequest,LastDDR,Win11Compatible,CPUGeneration"
        }
        else {
            # STOP CHECK FOR ALL HOSTNAMES ALREADY IN FILE
            $inputcsv = (Import-Csv $OutFile).Hostname
            foreach ($hostname in $inputcsv) {
                $allcomputers.Remove($hostname)
            }
        }
        
    }
    else {
        $Global:GetComputer = New-Object -TypeName System.Collections.ArrayList
    }

    #
    #
    # - FILTERS
    #
    #

    foreach ($hostname in $allcomputers) {
        if ($hostname -like "*$filter*") {
            $allfilteredcomputers.Add($hostname) | Out-Null
        }
    }
    
    # FILTER OUT DUPLICATES IN MECM
    $allfilteredcomputers = $allfilteredcomputers | Select-Object -Unique #side effect: Changes arraylist to array

    #
    #
    # - CODE
    #
    #

    foreach ($hostname in $allfilteredcomputers) {
        $hostname
        CompInf -hostname $hostname
    }
    #
    # - END
    #
    if ($GetComputer) {
        $GetComputer
    }
    $Global:displayname = $null
    $Global:username = $null
    $filter = $null
    $inputpath = $null
}