function Disable-OutlookElements {
    param (
    [Parameter(Mandatory=$true)][string]$hostname
    )
    $tnc = tnc $hostname
    if ($tnc.PingSucceeded) {
        if ((Invoke-command -Credential $(Get-AdmCreds) -computername $hostname {Get-ItemProperty 'HKLM:\Software\Microsoft\Office\Outlook\Addins\Gecko.Ephorte.Outlook.Main' -Name LoadBehavior}) -eq 0) {
            "Outlook Elements is already deactivated"
        } else {

            $response = Invoke-command -Credential $(Get-AdmCreds) -computername $hostname {New-Item 'HKLM:\Software\Microsoft\Office\Outlook\Addins\Gecko.Ephorte.Outlook.Main' -Force | Out-Null; (New-ItemProperty 'HKLM:\Software\Microsoft\Office\Outlook\Addins\Gecko.Ephorte.Outlook.Main' -Name 'LoadBehavior' -Value 0 -PropertyType dword).LoadBehavior}
            if ($response -eq 0) {
                "Success"
            } else {
                $response
                "Failed"
            }
        }
    } else {
        "Computer is offline"
    }
}
