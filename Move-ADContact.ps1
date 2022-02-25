function Move-ADContact {
    param (
        [Parameter(Mandatory=$true)][string]$mail
    )
    Get-ADObject -Credential $(Get-AdmCreds) -Filter {mail -eq "$mail"} | Move-ADObject -Credential $(Get-AdmCreds) -TargetPath "OU=Exchange_Kontakt,OU=\#Brukergrupper,OU=SADM,OU=MRFYLKE,DC=intern,DC=mrfylke,DC=no"
}