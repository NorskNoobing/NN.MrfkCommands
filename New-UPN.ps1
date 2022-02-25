<#
Automate applying new UPN
    1. Get all tickets in queue
    2. See if any of the tickets has "Epost* *har blitt endret til*"
    3. Assign ticket to myself
    4. Get ticket
    5. Import ticket contents
    6. Ask for confirmation in Discord with the reaction ✅ or 🚫 https://discord.com/developers/docs/resources/channel#create-message and https://discord.com/developers/docs/resources/channel#create-reaction
    7. If (✅.count -gt 1) continue, elseif (🚫.count -gt 1) exit https://discord.com/developers/docs/resources/channel#get-reactions
    8. Delete discord message
    9. Send confirmation of action in an internal note in the ticket
#>
function New-UPN {
    param (
        [Parameter(Mandatory=$true)][string]$username,
        [Parameter(Mandatory=$true)][string]$newupn,
        [Parameter(Mandatory=$true)][string]$oldupn
    )
    if ($newupn -like '*@mrfylke.no') {
        Set-ADUser -Credential $(Get-AdmCreds) -Identity $username -Replace @{UserPrincipalName="$newupn"}
        "UPN for user $username has been changed to $newupn"
    } else {
        "New UPN doesn't contain @mrfylke.no"
    }
}