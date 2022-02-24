function Move-ADContact {
    param (
        [Parameter(Mandatory=$true)][string]$mail
    )
    Get-ADObject -Filter {mail -eq "$mail"} | Move-ADObject -TargetPath "OU=Exchange_Kontakt,OU=\#Brukergrupper,OU=SADM,OU=MRFYLKE,DC=intern,DC=mrfylke,DC=no"
}