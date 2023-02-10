function New-MrfkShippingLabel {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)][ValidateScript({
            $_ -in (Get-SnipeLocations).name
        })][string]$location,
        [Parameter(Mandatory)][string]$displayname,
        [int]$copies,
        [string]$mobile,
        [string]$PrinterNetworkPath = "\\sr-safecom-sla1\PR-STORLABEL-SSDSK"
    )

    process {
        #Get the selected locations address
        $locationResult = Get-SnipeLocations -name $location

        $locationName = $locationResult.address
        $address = $locationResult.address2
        $postalCode = $locationResult.zip
        $city = $locationResult.city

        if (!$mobile) {
            try {
                $null = Get-ADUser -Filter "Name -eq 0"
            }
            catch [System.Management.Automation.CommandNotFoundException] {
                Write-Error -ErrorAction Stop -Message @"
Please install RSAT before running this function. You can install RSAT by following this guide:
https://github.com/NorskNoobing/NN.MrfkCommands#prerequisites
"@
            }

            #Search AD for phonenumber
            [string]$mobile = (Get-ADUser -Filter {DisplayName -like $displayname} -Properties mobile).mobile
        }

        if ($location) {
            $splat = @{
                "Recipient" = "$displayname`v$mobile"
                "Location" = "$locationName`v$address`v$postalCode $city"
                "PrinterNetworkPath" = $PrinterNetworkPath
                "copies" = $copies
            }
            Invoke-MrfkShippingLabelPrint @splat
        }
    }
}