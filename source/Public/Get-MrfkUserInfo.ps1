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
                $filter = "name -like `"*$Username*`""
            }
            "displayname" {
                $filter = "displayname -like `"*$DisplayName*`""
            }
            "mobilephone" {
                $filter = "mobilephone -like `"*$MobilePhone*`""
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

        #Get userinfo of the ADuser (limited to the first hit, and user can't pick manually "yet")
        $ADUser = ([array](Get-ADUser -filter $filter -Properties mobilephone,displayName))[0] |
        Select-Object -ExcludeProperty @(
            "ObjectClass","ObjectGUID","SID","AddedProperties","RemovedProperties",
            "ModifiedProperties","PropertyCount","PropertyNames","Line"
        )

        if ($IncludeComputerInfo) {
            $ComputerExportArr = Get-MrfkComputerInfo -Username $ADUser.Name
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