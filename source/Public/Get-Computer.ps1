function Get-Computer {
    [CmdletBinding(DefaultParameterSetName="List computers")]
    param (
        [Parameter(Mandatory,ValueFromPipeline,Position=0,ParameterSetName="Get computer by hostname")][string]$hostname,
        [pscredential]$credential,
        [Parameter(ParameterSetName="List computers")][switch]$ListComputers
    )
    
    begin {
        $allCpuInfoQry = @"
select distinct SMS_G_System_PROCESSOR.*
from SMS_R_System
inner join SMS_G_System_PROCESSOR
on SMS_G_System_PROCESSOR.ResourceID = SMS_R_System.ResourceId
"@
        Write-Information "$(Get-Date -Format "hh:mm:ss") Creating CimSession"
        $cimsession = New-CimSession -Credential $credential -ComputerName sccm-ps.intern.mrfylke.no -ErrorAction Stop

        Write-Information "$(Get-Date -Format "hh:mm:ss") Fetching all AD computers"
        $allAdComputers = Get-ADComputer -Filter *
        Write-Information "$(Get-Date -Format "hh:mm:ss") Fetching MECM computers"
        $allMecmComputers = Get-CimInstance -Query "Select * from SMS_R_System" -Namespace root/SMS/site_PS1 -CimSession $cimsession
        Write-Information "$(Get-Date -Format "hh:mm:ss") Fetching all CPU info"
        $allCpuInfo = Get-CimInstance -Query $allCpuInfoQry -Namespace root/SMS/site_PS1 -CimSession $cimsession
        Write-Information "$(Get-Date -Format "hh:mm:ss") Fetching all model info"
        $allModelInfo = Get-CimInstance -Query "Select * from SMS_G_System_Computer_System_Product" -Namespace root/SMS/site_PS1 -CimSession $cimsession
    
        $compInfExport = New-Object -TypeName System.Collections.ArrayList
    }
    
    process {
        switch ($PsCmdlet.ParameterSetName) {
            "List computers" {
                $CompInfDef = ${function:Get-CompInf}.ToString()

                $compInfExport = $allMecmComputers | ForEach-Object -ThrottleLimit 5000 -Parallel {
                    ${function:Get-CompInf} = $using:CompInfDef
                    Write-Information "$(Get-Date -Format "hh:mm:ss") Fetching CompInf for `"$($_.Name)`""

                    $CompInfSplat = @{
                        "allMecmComputers" = $using:allMecmComputers
                        "allCpuInfo" = $using:allCpuInfo
                        "allAdComputers" = $using:allAdComputers
                        "allModelInfo" = $using:allModelInfo
                        "hostname" = $_.Name
                    }
                    Get-CompInf @CompInfSplat
                    Write-Information "$(Get-Date -Format "hh:mm:ss") Completed fetching info for computer `"$($_.Name)`""
                }
            }
            "Get computer by hostname" {
                $CompInfSplat = @{
                    "allMecmComputers" = $allMecmComputers
                    "allCpuInfo" = $allCpuInfo
                    "allAdComputers" = $allAdComputers
                    "allModelInfo" = $allModelInfo
                    "hostname" = $hostname
                }
                $currentCompInf = Get-CompInf @CompInfSplat
                $compInfExport.Add($currentCompInf)
            }
        }
    }

    end {
        $compInfExport
    }
}