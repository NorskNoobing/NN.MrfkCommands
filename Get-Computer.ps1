<#
    1. Add CollectionName
    2. Call on functions within function to prevent using $var = $null
    3. Set required params and param types
    4. Foreach-Object -Parallel
    5. Add everything to PSCustomObject then output (write to cli or append object to file)
    6. Define type at every new variable, then remove all the if($var is [type])
#>
function CompInf {
    param (
        [Parameter(Mandatory)][string]$hostname,
        [string]$OutFile
    )
    Import-Module 'C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1'
    Set-Location ps1:

    $cmdeviceinfo = ([array](get-cmdevice -name $hostname))[0] #FILTER OUT DUPLICATES IN MECM
    $adcomputer = Get-ADComputer $hostname -Properties *
    $collections = Get-CimInstance -ComputerName sccm-ps.intern.mrfylke.no -Namespace root/SMS/site_PS1 -Query "SELECT SMS_Collection.* FROM SMS_FullCollectionMembership, SMS_Collection where name = '$hostname' and SMS_FullCollectionMembership.CollectionID = SMS_Collection.CollectionID"

    #todo: ask the db for model instead of relying on collectionname
    $model = ([array](($collections.Name | Select-String "All HP*", "All Lenovo*") -replace "All ", ""))[0] #Filter out $model outputting multiple names by selecting the first object in array. E.g. "HP Elitebook 840 G3 HP EliteBook 840 G3 = N75 Ver 01.53"

    [string]$CPUGen = ($collections.Name | Select-String " Gen - ") -replace "All ", ""

    # WIN11 COMPATIBILITY
    [int]$CPUGenNumber = $CPUGen -replace "[^0-9]"
    if (($CPUGenNumber -le 7) -and ($CPUGenNumber -gt 0)) {
        $win11compat = $false
    } elseif ($CPUGenNumber -ge 8) {
        $win11compat = $true
    } else {
        $win11compat = "UNKNOWN"
    }

    $username = $cmdeviceinfo.username
    $displayname = (Get-ADUser "$username" -Properties DisplayName).DisplayName

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
        MACAddress        = $cmdeviceinfo.MACAddress
        Model             = $model
        SN                = $cmdeviceinfo.serialnumber
        GUID              = $cmdeviceinfo.SMBIOSGUID
        CreatedAt         = $adcomputer.Created
        IsActive          = $cmdeviceinfo.IsActive
        DeviceOSBuild     = $cmdeviceinfo.DeviceOSBuild
        OS                = $adcomputer.OperatingSystem
        IPv4              = $adcomputer.IPv4Address
        Enabled           = $adcomputer.Enabled
        Status            = $status
        LastHardwareScan  = $cmdeviceinfo.LastHardwareScan
        LastPolicyRequest = $cmdeviceinfo.LastPolicyRequest
        LastDDR           = $cmdeviceinfo.LastDDR
        Win11Compatible   = $win11compat
        CPUGeneration     = $CPUGen
    }

    # OUTPUT TYPE
    if ($OutFile) {
        $computerinfo | Export-Csv -Path $OutFile -Delimiter ";" -Append #fix: file is in use by another process
    }
    else {
        $computerinfo
    }
}
function Get-Computer {
    param (
        [array]$hostnames, 
        [array]$importExcelPaths, 
        [array]$collectionIDs, 
        [string]$OutFile, 
        [array]$displaynames, 
        [array]$usernames, 
        [string]$filter
    )
    Import-Module 'C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1'
    Set-Location ps1:
    $allcomputers = New-Object -TypeName System.Collections.ArrayList
    $allfilteredcomputers = New-Object -TypeName System.Collections.ArrayList

    # - INPUTS
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
        $collectionIDs | Foreach-Object {
            $cidHostnamearray = (Get-CMCollectionMember -CollectionId $_).Name
            $cidHostnamearray | Foreach-Object {
                $allcomputers.Add($_) | Out-Null
            }
        }
    }
    # DISPLAYNAME
    if ($displaynames) {
        $displaynames | Foreach-Object {
            $dnUsername = (Get-ADUser -Filter { DisplayName -like $_ }).Name
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

    # - FILTERS
    if ($OutFile) {
        $DoesOutFileExist = Test-Path -Path $OutFile -PathType Leaf
        
        if ($DoesOutFileExist) {
            #Filter out all hostnames already in file
            $importHostnames = (Import-Csv $OutFile -Delimiter ";").Hostname
            $importHostnames | ForEach-Object {
                if ($allcomputers -contains $_) {
                    $allcomputers.Remove($_)
                }
            }
        }
    }

    $allcomputers | ForEach-Object {
        if ($_ -like "*$filter*") {
            $allfilteredcomputers.Add($_) | Out-Null
        }
    }
    
    # FILTER OUT DUPLICATES IN MECM
    $allfilteredcomputers = $allfilteredcomputers | Select-Object -Unique #side effect: Changes arraylist to array

    # - CODE
    $allfilteredcomputers | Foreach-Object {
        CompInf -hostname $_ -OutFile $OutFile
    }
}