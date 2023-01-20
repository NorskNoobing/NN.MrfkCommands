function Get-MrfkUserInfo {
    [CmdletBinding()]
    param (
        [Parameter(ParameterSetName="username")][string]$Username,
        [Parameter(ParameterSetName="displayname")][string]$DisplayName,
        [Parameter(ParameterSetName="mobilephone")][string]$MobilePhone,
        [switch]$IncludeComputerInfo,
        [switch]$ExpandComputerInfo
    )

    begin {
        $RequiredModulesNameArray = @("NN.WindowsSetup")
        $RequiredModulesNameArray.ForEach({
            if (Get-InstalledModule $_ -ErrorAction SilentlyContinue) {
                Import-Module $_ -Force
            } else {
                Install-Module $_ -Force
            }
        })

        #Install RSAT
        Install-RSAT -WUServerBypass
    }

    process {
        switch ($PsCmdlet.ParameterSetName) {
            "username" {
                $filter = "SamAccountName -like `"$Username`""
            }
            "displayname" {
                $filter = "DisplayName -like `"$DisplayName`""
            }
            "mobilephone" {
                $filter = "MobilePhone -like `"$MobilePhone`""
            }
        }

        #Install required modules
        $RequiredModulesNameArray = @('NN.MrfkCommands')
        $RequiredModulesNameArray.ForEach({
            if (Get-InstalledModule $_ -ErrorAction SilentlyContinue) {
                Import-Module $_ -Force
            } else {
                Install-Module $_ -Force -Repository PSGallery
            }
        })

        #Get userinfo of the ADusers
        $ADUser = Get-ADUser -filter $filter -Properties MobilePhone,DisplayName | Select-Object @(
            "DisplayName","Name","SamAccountName","MobilePhone",
            "UserPrincipalName","Enabled","DistinguishedName"
        )

        #Pick an ADUser if we get multiple hits on the search query
        if ($ADUser -is [array]) {
            $splat = @{
                "Title" = "Found multiple hits on the input. Please select the user."
                "OutputMode" = "Single"
            }
            $ADUser = $ADUser | Out-GridView @splat
        }

        if (!$ADUser) {
            Write-Error -ErrorAction "Stop" -Message "Please select a user before continuing."
        }

        if ($IncludeComputerInfo) {
            $ComputerExportArr = Get-MrfkComputerInfo -Username $ADUser.SamAccountName
            if (!$ExpandComputerInfo) {
                $ADUser | Add-Member -MemberType NoteProperty -Name "Computers" -Value $ComputerExportArr
            }
        }

        #Post output
        $ADUser
        if ($ExpandComputerInfo) {
            $ComputerExportArr
        }
    }
}