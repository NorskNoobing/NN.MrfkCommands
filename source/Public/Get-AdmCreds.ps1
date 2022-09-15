function Get-AdmCreds {
    param (
        [string]$admCredsPath = "$env:USERPROFILE\.creds\Windows\adm_creds.xml"
    )

    if (!(Test-Path $admCredsPath)) {
        #Create directories for adm_creds
        $admCredsDir = $admCredsPath.Substring(0, $admCredsPath.lastIndexOf('\'))
        if (!(Test-Path $admCredsDir)) {
            New-Item -ItemType Directory $admCredsDir | Out-Null
        }
        
        #Create adm_creds file
        Get-Credential -Message "Enter admin credentials" | Export-Clixml $admCredsPath
    } else {
        Import-Clixml $admCredsPath
    }
}