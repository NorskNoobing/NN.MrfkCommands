function Get-FreeHostnames {
    param (
        [Parameter(Mandatory,HelpMessage="Example inputs: `"LT-SADM-`" or `"PC-SADM-`"")][ValidateScript({
            $_ -like "*-*-"
        }, ErrorMessage="Please provide a valid input. Valid inputs are `"<device-type>-<locationname>-`" Example: `"LT-SADM-`" or `"PC-VROM-`""
        )][string]$filter,
        [pscredential]$credential,
        [int]$count
    )
    $cimsession = New-CimSession -Credential $credential -ComputerName sccm-ps.intern.mrfylke.no -ErrorAction Stop
    $allMecmComputers = Get-CimInstance -Query "Select * from SMS_R_System" -Namespace root/SMS/site_PS1 -CimSession $cimsession
    
    $pcNumArr = $allMecmComputers.where({$_.Name -like "*$filter*"}).Name.Replace("$filter","","OrdinalIgnoreCase") -replace "^0+","" | Sort-Object
    $freePcNumArr = (1..999).where({$_ -notin $pcNumArr})

    if ($count) {
        $freePcNumArr = $freePcNumArr | Select-Object -First $count
    }
    
    $freePcNumArr.ForEach({
        $num = ([string]$_).PadLeft(3,'0')
        "$filter$num"
    })
}
#todo: add $HostnamePrefix="PC","TB","LT" and $LocationName="SADM","VROM",etc.
#todo: Get all locaiton shorts from hostnames currently in MECM and validatescript