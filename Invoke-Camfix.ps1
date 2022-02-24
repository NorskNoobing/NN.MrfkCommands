function Invoke-Camfix {
    param (
        [Parameter(Mandatory=$true)][string]$hostname
    )
    $tnc = tnc $hostname
    if ($tnc.PingSucceeded) {
        if (Invoke-command -computername $hostname {Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows Media Foundation\Platform\' -Name EnableFrameServerMode}) {
            "Camfix is already added"
        } else {
            $response = Invoke-command -computername $hostname {(New-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows Media Foundation\Platform\' -Name 'EnableFrameServerMode' -Value 0 -PropertyType dword).EnableFrameServerMode}
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
Invoke-Command -ComputerName 'LT-SADM-955' -ScriptBlock {Get-ItemProperty -Path HKLM:\Software\Policies\Microsoft\Windows\AppPrivacy\LetAppsAccessMicrophone}