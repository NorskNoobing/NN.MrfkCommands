function Invoke-PCCheck {
    param (
        [Parameter(Mandatory)][string]$displayName
    )
    $notes = New-Object -TypeName System.Collections.ArrayList

    # Check if there is a active EXO sessions
    $psSessions = Get-PSSession | Select-Object -Property State, Name
    If (!(($psSessions -like '@{State=Opened; Name=ExchangeOnlineInternalSession*').Count -gt 0)) {
        throw "Run Connect-ExchangeOnline before running this function."
    }
    
    #Get userinfo
    $adUser = Get-ADUser -Credential $(Get-AdmCreds) -Filter {DisplayName -eq $displayName} -Properties *

    if (!$adUser) {
        throw "Can't find user `"$displayName`" in AD."
    }

    $adUsername = $aduser.Name
    $homeDir = $adUser.HomeDirectory
    
    #Create home dir
    if ($homeDir) {
        [bool]$homeDirExists = Invoke-gsudo {
            if (!(Test-Path $using:homeDir)) {
                New-Item -ItemType Directory $using:homeDir | Out-Null
            }
            # Get the ACL for an existing folder
            $ExistingACL = Get-Acl -Path $using:homeDir
            # Sets Full Control permission for the given user on This Folder, Subfolders and Files
            $Permissions = "intern\$using:adUsername", 'FullControl', 'ContainerInherit,ObjectInherit', 'None', 'Allow'
            # Create a new FileSystemAccessRule object
            $Rule = New-Object -TypeName System.Security.AccessControl.FileSystemAccessRule -ArgumentList $Permissions
            # Modify the existing ACL to include the new rule
            $ExistingACL.SetAccessRule($Rule)
            # Apply the modified access rule to the folder
            $ExistingACL | Set-Acl -Path $using:homeDir | Out-Null
    
            Test-Path $using:homeDir
        }
    } else {
        [bool]$homeDirExists = $false
        $notes.Add("There's no homedir path on the AD user.")
    }

    #Check if mailbox is in Exchange Online
    if (Get-InstalledModule ExchangeOnlineManagement) {
        Import-Module ExchangeOnlineManagement
    } else {
        Install-Module ExchangeOnlineManagement -Force
    }
    
    [bool]$migrated = Get-EXOMailbox $adUser.mail -ErrorAction SilentlyContinue
    
    #Reset AD password
    Set-ADAccountPassword $adUser.Name -Credential $(Get-AdmCreds) -NewPassword ("Molde123!" | ConvertTo-SecureString -AsPlainText) -Reset -Server "dc01.intern.mrfylke.no"

    [PSCustomObject]@{
        DisplayName = $adUser.DisplayName
        Username = $adUser.Name
        UPN = $adUser.UserPrincipalName
        ADLocation = $adUser.DistinguishedName
        HomeDirExists = $homeDirExists
        Migrated = $migrated
        Notes = $notes
    }
}