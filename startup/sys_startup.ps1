<#
Documentation
https://docs.microsoft.com/en-us/windows-hardware/customize/power-settings/power-button-and-lid-settings-lid-switch-close-action
https://superuser.com/questions/874849/change-what-closing-the-lid-does-from-the-commandline/874858
#>

function Send-DiscordMessage {
    param (
        $Content = "" ,
        $HookURI = $(Import-Clixml -Path "$env:USERPROFILE\.creds\WUHook.xml" | ConvertFrom-SecureString -AsPlainText)
    )

    $InvokeRestMethodSplat = @{
        Body = [PSCustomObject]@{
                content = $Content
            } | ConvertTo-Json
            
        ContentType = 'Application/Json'
        Method = 'POST'
        Uri = $HookURI
    }
    Invoke-RestMethod @InvokeRestMethodSplat
}

#Set background
$BGPath = "$env:windir\Resources\Themes\SADM-Bakgrunn_Ny.jpg"
$CurrentBGSize = (Get-Item -Path $BGPath).Length/1MB
[int]$winver = (Get-CimInstance -ClassName Win32_OperatingSystem).name.split("|")[0] -replace "[^0-9]"
if ($CurrentBGSize -lt 1) {
    if ($winver -eq 11) {
        Copy-Item -Path "$env:SystemRoot\web\wallpaper\Windows\img19.jpg" -Destination $BGPath
    } 
    elseif (($winver -gt 0) -and !($winver -eq 11)) {
        Invoke-WebRequest -Uri "https://wallpaperaccess.com/download/windows-11-4k-6233787" -OutFile $BGPath
    }
}

#Set startscreen
$StartscreenPath = "$env:windir\Resources\Themes\Startbilde.jpg"
$CurrentStartscreenSize = (Get-Item -Path $StartscreenPath).Length/1MB
if ($CurrentStartscreenSize -lt 1) {
    #[Only on win11] Copy-Item -Path "$env:SystemRoot\web\wallpaper\Windows\img19.jpg" -Destination $StartscreenPath
    Invoke-WebRequest -Uri https://wallpaperaccess.com/download/windows-11-4k-6233787 -OutFile $StartscreenPath
}

#Enable WU
New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" -Force
Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" -name "DisableDualScan" -value 0
Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" -name "DisableWindowsUpdateAccess" -value 0
Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" -name "SetPolicyDrivenUpdateSourceForDriverUpdates" -value 0
Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" -name "SetPolicyDrivenUpdateSourceForFeatureUpdates" -value 0
Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" -name "SetPolicyDrivenUpdateSourceForOtherUpdates" -value 0
Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" -name "SetPolicyDrivenUpdateSourceForQualityUpdates" -value 0
Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" -name "UseWUServer" -value 0
Restart-Service wuauserv

powercfg -change -standby-timeout-ac 0
powercfg -change -hibernate-timeout-ac 0
powercfg -change -monitor-timeout-ac 0

#Do nothing when closing lid while plugged into AC
$powerSchemeGUID = (powercfg -getactivescheme).split(" ")[3]
$subgroupGUID = ((powercfg -Q $powerSchemeGUID | Select-String "Power buttons and lid") -split(" "))[4] #Test EN
if (!$subgroupGUID) {
    $subgroupGUID = ((powercfg -Q $powerSchemeGUID | Select-String "Av/på-knapper og deksel") -split(" "))[4] #If EN fails, use NO
}
powercfg -SETACVALUEINDEX $powerSchemeGUID $subgroupGUID 5ca83367-6e45-459f-a27b-476b1d01c936 0

if (!(Get-InstalledModule -Name "PSWindowsUpdate")) {
    Install-Module -Name "PSWindowsUpdate" -Force
}

if (Get-WindowsUpdate) {
    Install-WindowsUpdate -AcceptAll -IgnoreReboot
    Send-DiscordMessage -Content "$(hostname) is updated."
}

$LastNotificationPath = "C:\Users\Public\WUNotification.xml"
if (!(Test-Path -Path $LastNotificationPath)) {
    Get-Date 0 | Export-Clixml -Path $LastNotificationPath
}

$CurrentDate = Get-Date
$LastNotification = Import-Clixml -Path $LastNotificationPath

if ((($CurrentDate - $LastNotification).Hours -ge 2) -and (Get-WURebootStatus -Silent)) {
    $Min = Get-Date '04:00'
    $Max = Get-Date '07:00'
    $CurrentDate = Get-Date
    $isDateBetweenMinAndMax = $min.TimeOfDay -le $CurrentDate.TimeOfDay -and $Max.TimeOfDay -ge $CurrentDate.TimeOfDay

    Get-Date | Export-Clixml -Path $LastNotificationPath
    if ($isDateBetweenMinAndMax) {
        Send-DiscordMessage -Content "Rebooting $(hostname)."
        shutdown /r /t 0
        exit
    }
    
    Send-DiscordMessage -Content "A reboot of $(hostname) is required."
}