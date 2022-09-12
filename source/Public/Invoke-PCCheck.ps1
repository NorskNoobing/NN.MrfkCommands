function Invoke-PCCheck {
    param (
        [Parameter(Mandatory)][ValidateSet("Atlanten VGS","Borgund VGS","Fagerlia VGS","Fagskolen i Alesund",
        "Gjermundnes VGS","Haram VGS","Herøy VGS","Hustadvika VGS","Kristiansund VGS","Rauma VGS","Romsdal VGS",
        "Spjelkavik VGS","Sunndal VGS","Sykkylven VGS","Surnadal VGS","Tingvoll VGS","Ulstein VGS","Volda VGS","Orsta VGS",
        "Alesund VGS (Volsdalsberga)","Alesund VGS (Fagerlia)","Campus Kristiansund","Olsvika","Carolus","Fylkeshuset","Iteam","Skjeltene")][string]$location,
        [Parameter(Mandatory)][string]$displayName
    )
    # Check if there is a active EXO sessions
    $psSessions = Get-PSSession | Select-Object -Property State, Name
    If (!(($psSessions -like '@{State=Opened; Name=ExchangeOnlineInternalSession*').Count -gt 0)) {
        throw "Run Connect-ExchangeOnline before running this function."
    }

    #Fetch admCreds for running AD functions
    $admCredsPath = "$env:USERPROFILE\.creds\Windows\adm_creds.xml"
    if (!(Test-Path $admCredsPath)) {
        New-Item -ItemType Directory $($admCredsPath.Substring(0, $admCredsPath.lastIndexOf('\'))) | Out-Null
        
        $admCreds = Get-Credential -Message "Enter admin credentials"
        $admCreds | Export-Clixml $admCredsPath
    } else {
        $admCreds = Import-Clixml $admCredsPath
    }
    
    #Get userinfo
    $adUser = Get-ADUser -Credential $admCreds -Filter {DisplayName -eq $displayName} -Properties *
    $adUsername = $aduser.Name
    $homeDir = $adUser.HomeDirectory
    
    #Create home dir
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

    #Check if mailbox is in Exchange Online
    if (Get-InstalledModule ExchangeOnlineManagement) {
        Import-Module ExchangeOnlineManagement
    } else {
        Install-Module ExchangeOnlineManagement -Force
    }
    
    [bool]$migrated = Get-EXOMailbox $adUser.mail -ErrorAction SilentlyContinue
    
    #Reset AD password
    Set-ADAccountPassword $adUser.Name -Credential $admCreds -NewPassword ("Molde123!" | ConvertTo-SecureString -AsPlainText) -Reset

    [PSCustomObject]@{
        DisplayName = $adUser.DisplayName
        Username = $adUser.Name
        UPN = $adUser.UserPrincipalName
        ADLocation = $adUser.DistinguishedName
        HomeDirExists = $homeDirExists
        Migrated = $migrated
    }
    
    #Print shippinglabel
    New-ShippingLabel -displayname $displayName -location $location
}