Import-Module "C:\Users\Public\Documents\PowerShell\Modules\Posh-SSH\3.0.6\Posh-SSH.psd1" -Force

$user = "sys_tbkiosk"
$Credential = Get-Credential -UserName $user -Message "Enter password for $user" 

$SshSession = New-SSHSession -ComputerName "sr-sikt-001" -Credential $Credential -Force

Invoke-SSHCommand -SessionId $SshSession.SessionId -Command "sudo reboot now"