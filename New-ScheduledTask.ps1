function New-ScheduledTask {
    #TODO: Repeat every 1min, for a duration of 450min
    $Action = New-ScheduledTaskAction -Execute 'pwsh.exe' -Argument '-File "C:\Users\Public\Get-CallStatus.ps1"'
    $Trigger = New-ScheduledTaskTrigger -Daily -At 08:00 -RepetitionInterval (New-TimeSpan -Minutes 1) -RepetitionDuration (New-TimeSpan -Hours 7 -Minutes 30) #The flag daily doesn't work with the repetition flags
    $Settings = New-ScheduledTaskSettingsSet
    $Task = New-ScheduledTask -Action $Action -Trigger $Trigger -Settings $Settings
    #TODO: Run whether user is logged on or not
    Register-ScheduledTask -TaskName 'Startup' -InputObject $Task
}
