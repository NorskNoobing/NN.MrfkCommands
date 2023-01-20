function New-MrfkAdmCreds {
    param (
        [string]$admCredsPath = "$env:USERPROFILE\.creds\MRFK\adm_creds.xml"
    )

    #Create parent folders for the file
    $admCredsDir = $admCredsPath.Substring(0, $admCredsPath.lastIndexOf('\'))
    if (!(Test-Path $admCredsDir)) {
        $null = New-Item -ItemType Directory $admCredsDir
    }
    
    #Create adm_creds file
    Get-Credential -Message "Enter your mrfk admin credentials" | Export-Clixml $admCredsPath
}