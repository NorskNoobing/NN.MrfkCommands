<#
    1. Add CollectionName
    2. Define type at every new variable, then remove all the if($var is [type])
    3. Set requirement of PS7
#>
function CompInf {
    param (
        [Parameter(Mandatory)][string]$hostname,
        [Parameter(Mandatory)][string]$MECMModulePath,
        [string]$OutFile,
        [bool]$skipdb
    )

    Import-Module $MECMModulePath
    Set-Location ps1:

    $cmdeviceinfo = ([array](get-cmdevice -name $hostname))[0] #FILTER OUT DUPLICATES OF THE SAME HOSTNAME IN MECM
    $adcomputer = Get-ADComputer $hostname -Properties *

    if (!$skipdb) {
        [array]$collections = Get-CimInstance -ComputerName sccm-ps.intern.mrfylke.no -Namespace root/SMS/site_PS1 -Query "SELECT SMS_Collection.* FROM SMS_FullCollectionMembership, SMS_Collection where name = '$hostname' and SMS_FullCollectionMembership.CollectionID = SMS_Collection.CollectionID"

        [string]$model = ([array](($collections.Name | Select-String "All HP*", "All Lenovo*") -replace "All ", ""))[0] #Filter out $model outputting multiple names by selecting the first object in array. E.g. "HP Elitebook 840 G3 HP EliteBook 840 G3 = N75 Ver 01.53"
    
        [string]$CPUGenName = ($collections.Name | Select-String " Gen - ") -replace "All ", ""
        [string]$CPUGenNumber = $CPUGenName -replace "[^0-9]"
    }

    [string]$username = $cmdeviceinfo.username
    [string]$displayname = (Get-ADUser "$username" -Properties DisplayName).DisplayName

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
        Hostname          = $hostname
        DisplayName       = $displayname
        Username          = $username
        MACAddress        = $cmdeviceinfo.MACAddress
        Model             = $model
        SN                = $cmdeviceinfo.serialnumber
        GUID              = $cmdeviceinfo.SMBIOSGUID
        CreatedAt         = $adcomputer.Created
        "IsActive(MECM)"  = $cmdeviceinfo.IsActive
        DeviceOSBuild     = $cmdeviceinfo.DeviceOSBuild
        OS                = $adcomputer.OperatingSystem
        IPv4              = $adcomputer.IPv4Address
        "Enabled(AD)"     = $adcomputer.Enabled
        Status            = $status
        LastHardwareScan  = $cmdeviceinfo.LastHardwareScan
        LastPolicyRequest = $cmdeviceinfo.LastPolicyRequest
        LastDDR           = $cmdeviceinfo.LastDDR
        CPUGeneration     = $CPUGenNumber
    }
}

function Get-Computer {
    param (
        [array]$hostnames,
        [array]$importExcelPaths,
        [array]$collectionIDs,
        [array]$displaynames,
        [array]$usernames,
        [string]$OutFile,
        [string]$filter,
        [switch]$skipdb
    )

    if (-not ($hostnames -or $importExcelPaths -or $collectionIDs -or $displaynames -or $usernames)) {
        throw "Input missing"
    }

    switch ($env:COMPUTERNAME) {
        "wintools03" {
            [string]$MECMModulePath = 'C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1'
            
        }
        "wintools04" {
            [string]$MECMModulePath = 'C:\Program Files (x86)\Microsoft Endpoint Manager\AdminConsole\bin\ConfigurationManager.psd1'
        }
        Default {
            throw "Please run this function on a supported computer."
        }
    }

    Import-Module $MECMModulePath

    $allcomputers = New-Object -TypeName System.Collections.ArrayList
    $allCompInf = [System.Collections.Concurrent.ConcurrentBag[psobject]]::new()

    #INPUTS
    # XLSX
    if ($importExcelPaths) {
        $importExcelPaths | Foreach-Object {
            Import-Excel -Path $_ | Foreach-Object {
                $allcomputers.Add($_.hostname) | Out-Null
            }
        }
    }
    # HOSTNAME
    if ($hostnames) {
        $hostnames | Foreach-Object {
            $allcomputers.Add($_) | Out-Null
        }
    }
    # COLLECTIONID
    if ($collectionIDs) {
        Set-Location ps1:
        $collectionIDs | Foreach-Object {
            [array]$cidHostnamearray = (Get-CMCollectionMember -CollectionId $_).Name
            $cidHostnamearray | Foreach-Object {
                $allcomputers.Add($_) | Out-Null
            }
        }
        Set-Location "C:"
    }
    # DISPLAYNAME
    if ($displaynames) {
        $displaynames | Foreach-Object {
            [string]$dnUsername = (Get-ADUser -Filter { DisplayName -like $_ }).Name

            [array]$dnHostname = (Get-CimInstance -ComputerName sccm-ps.intern.mrfylke.no -Namespace root/SMS/site_PS1 -Query "select Name from sms_r_system where LastLogonUserName='$dnUsername'").Name
            $dnHostname | Foreach-Object {$allcomputers.Add($_) | Out-Null}
        }
    }
    # USERNAME
    if ($usernames) {
        $usernames | Foreach-Object {
            [array]$unHostname = (Get-CimInstance -ComputerName sccm-ps.intern.mrfylke.no -Namespace root/SMS/site_PS1 -Query "select Name from sms_r_system where LastLogonUserName='$_'").Name
            $unHostname | Foreach-Object {$allcomputers.Add($_) | Out-Null}
        }
    }

    #FILTERS
    if ($OutFile) {
        [bool]$DoesOutFileExist = Test-Path -Path $OutFile -PathType Leaf
        
        if ($DoesOutFileExist) {
            #Filter out all hostnames already in file
            [array]$importHostnames = (Import-Csv $OutFile -Delimiter ";").Hostname
            $importHostnames | ForEach-Object {
                if ($allcomputers -contains $_) {
                    $allcomputers.Remove($_)
                }
            }
        }
    }

    if ($filter) {
        $allcomputers | ForEach-Object {
            if ($_ -like "*$filter*") {
                $allcomputers.Remove($_) | Out-Null
            }
        }
    }
    
    #Removes duplicate hostnames from array
    [array]$allcomputers = $allcomputers | Select-Object -Unique

    #CODE
    $CompInfDef = ${function:CompInf}.ToString()

    if ($skipdb) {
        $throttlelimit = 1000
    } else {
        $throttlelimit = 10
    }

    $allcomputers | Foreach-Object -ThrottleLimit $throttlelimit -Parallel {
        ${function:CompInf} = $using:CompInfDef
        $currentCompInf = CompInf -hostname $_ -OutFile $using:OutFile -skipdb $using:skipdb -MECMModulePath $using:MECMModulePath
        ($using:allCompInf).Add($currentCompInf)
    }

    # OUTPUT TYPE
    if ($OutFile) {
        $allCompInf | Export-Csv -Path $OutFile -Delimiter ";" -Append
    }
    else {
        $allCompInf
    }
}