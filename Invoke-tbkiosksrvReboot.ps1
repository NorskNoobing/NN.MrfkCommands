function Invoke-tbkiosksrvReboot {
    if (!(Get-InstalledModule -Name Posh-SSH)) {
        Install-Module -Name Posh-SSH -Force
    } else {
        Import-Module Posh-SSH -Force
    }

    $user = "sys_tbkiosk"
    $Credential = Get-Credential -UserName $user -Message "Enter password for $user" 

    $SshSession = New-SSHSession -ComputerName "sr-sikt-001" -Credential $Credential -Force

    Invoke-SSHCommand -SessionId $SshSession.SessionId -Command "sudo reboot now"
}