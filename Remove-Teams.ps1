function Remove-Teams {
    Stop-Process -Name "Outlook"
    Stop-Process -Name "Teams"

    Remove-Item -Path "$env:APPDATA/Teams" -Force
    Remove-Item -Path "$env:APPDATA/Microsoft/Teams" -Force
    Remove-Item -Path "$env:LOCALAPPDATA/Microsoft/Teams" -Force
    Remove-Item -Path "$env:LOCALAPPDATA/Microsoft/TeamsMeetingAddin" -Force
    Remove-Item -Path "$env:LOCALAPPDATA/Microsoft/TeamsPresenceAddin" -Force
}
