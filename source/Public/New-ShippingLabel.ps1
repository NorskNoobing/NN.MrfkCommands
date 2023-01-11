function New-ShippingLabel {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)][ValidateScript({
            $_ -in (Get-SnipeLocations).name
        })][string]$location,
        [Parameter(Mandatory)][string]$displayname,
        [int]$copies = 1,
        [string]$mobile,
        [string]$PrinterNetworkPath = "\\sr-safecom-sla1\PR-STORLABEL-SSDSK"
    )
    
    begin {
        $RequiredModulesNameArray = @("NN.SnipeIT","NN.WindowsSetup")
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
        #Get the selected locations address
        $locationResult = Get-SnipeLocations -name $location

        $locationName = $locationResult.address
        $address = $locationResult.address2
        $postalCode = $locationResult.zip
        $city = $locationResult.city

        if (!$mobile) {
            #Search AD for phonenumber
            [string]$mobile = (Get-ADUser -Filter {DisplayName -like $displayname} -Properties mobile).mobile
        }

        if ($location) {
            #Get current default printer
            $defaultPrinter = (Get-CimInstance -Class Win32_Printer).where({$_.Default -eq $true}).Name

            #Set printer named PR-STORLABEL-SSDSK to default printer
            if (!(Get-Printer -Name $PrinterNetworkPath -ErrorAction SilentlyContinue)) {
                (New-Object -ComObject WScript.Network).AddWindowsPrinterConnection($PrinterNetworkPath)
            }
            (New-Object -ComObject WScript.Network).SetDefaultPrinter($PrinterNetworkPath)

            #Create new Word document
            $WordObj = New-Object -ComObject Word.Application
            $null = $WordObj.Documents.Add()

            #Set page size and margins
            $WordObj.Selection.PageSetup.PageHeight = "192mm"
            $WordObj.Selection.PageSetup.PageWidth = "102mm"
            $WordObj.Selection.PageSetup.TopMargin = "12,7mm"
            $WordObj.Selection.PageSetup.BottomMargin = "12,7mm"
            $WordObj.Selection.PageSetup.LeftMargin = "12,7mm"
            $WordObj.Selection.PageSetup.RightMargin = "12,7mm"

            #Insert content
            $WordObj.Selection.Font.Bold = 1
            $WordObj.Selection.TypeText("Mottaker:
            ")
            $WordObj.Selection.Font.Bold = 0
            $WordObj.Selection.TypeText("$displayname
            $mobile")
            $WordObj.Selection.TypeParagraph()
            $WordObj.Selection.Font.Bold = 1
            $WordObj.Selection.TypeText("Addresse:
            ")
            $WordObj.Selection.Font.Bold = 0
            $WordObj.Selection.TypeText("$locationName
            $address
            $postalCode $city")

            (1..$copies).ForEach({
                #Send To Default Printer
                $WordObj.PrintOut()
            })

            #Change default printer back to the previous value
            (New-Object -ComObject WScript.Network).SetDefaultPrinter("$defaultPrinter")

            #Close File without saving
            $WordObj.ActiveDocument.Close(0)
            $WordObj.quit() 
        }
    }
}