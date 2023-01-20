function Get-MrfkAdmCreds {
    param (
        [string]$admCredsPath = "$env:USERPROFILE\.creds\MRFK\adm_creds.xml"
    )

    if (!(Test-Path $admCredsPath)) {
        New-MrfkAdmCreds
    }
    
    Import-Clixml $admCredsPath
}