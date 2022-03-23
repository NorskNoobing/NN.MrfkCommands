function Remove-Userprofile {
    param (
        [Parameter(Mandatory)][string]$username,
        [Parameter(Mandatory)][array]$hostnames
    )
    
    #TODO: Get hostnames
    $hostnames | ForEach-Object {
        $hostname = "$_"
        #Connect to remote computer
        Invoke-Command -Credential $(Get-AdmCreds) -ComputerName $hostname -ScriptBlock {
            #Force user to logoff
            $sessionId = ((quser.exe | Where-Object {$_ -match $username}) -split ' +')[2]
            logoff.exe $sessionId
            
            Move-Item -Path "C:\Users\$username" -Destination "C:\Users\$username.old" #Rename userprofile
            Get-CimInstance -Class Win32_UserProfile | Where-Object { $_.LocalPath.split('\')[-1] -eq "$username"} | Remove-WmiObject #Remove userprofile object
        }
    }
}