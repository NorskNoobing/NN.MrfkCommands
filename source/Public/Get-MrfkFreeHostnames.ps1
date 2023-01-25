function Get-MrfkFreeHostnames {
    param (
        [Parameter(Mandatory)][ValidateSet("LT","PC","TB")][string]$Prefix,
        [Parameter(Mandatory)][string]$LocationCode,
        [Parameter(Mandatory)][int]$Digits,
        [int]$Count,
        [string]$MECMNameSpace = "root/SMS/site_PS1",
        [string]$MECMHost = "sccm-ps.intern.mrfylke.no"
    )
    $CimSession = New-CimSession -Credential (Get-MrfkAdmCreds) -ComputerName $MECMHost -ErrorAction Stop
    $AllMecmComputers = Get-CimInstance -Namespace $MECMNameSpace -CimSession $CimSession -Query @"
Select * from SMS_R_System
"@
    
    $Filter = "$Prefix-$LocationCode-"

    $PCNumArr = $AllMecmComputers.where({
        ($_.Name -like "$Filter*") -and (
            $_.Name -like ($Filter + ((
                $_.Name.Replace("$Filter","","OrdinalIgnoreCase") -replace "^0+",""
            ).PadLeft($Digits,'0')))
        )
    }).Name

    if ($PCNumArr) {
        $PCNumArr = $PCNumArr.Replace("$Filter","","OrdinalIgnoreCase") -replace "^0+","" | Sort-Object
    }
    
    $FreePCNumArr = (1..("9" * $Digits)).where({$_ -notin $PCNumArr})

    if ($Count) {
        $FreePCNumArr = $FreePCNumArr | Select-Object -First $Count
    }
    
    $FreePCNumArr.ForEach({
        $Num = ([string]$_).PadLeft($Digits,'0')
        "$Filter$Num"
    })
}