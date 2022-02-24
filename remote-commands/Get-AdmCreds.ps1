function Get-AdmCreds {
    $CredsPath = "$env:USERPROFILE\.creds"
    $AdmCredsPath = "$CredsPath\adm.xml"

    if (!$CredsPath) {
        mkdir -p $CredsPath
    }

    if (Test-Path -Path $AdmCredsPath) {
        $AdmCreds = Import-Clixml -Path $AdmCredsPath
    } else {
        $AdmCreds = Get-Credential
        $AdmCreds | Export-Clixml -Path $AdmCredsPath
    }
    $AdmCreds
}