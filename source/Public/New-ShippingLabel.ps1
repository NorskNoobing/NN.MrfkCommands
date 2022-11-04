function New-ShippingLabel {
    #Requires -Module NN.SnipeIT,NN.WindowsSetup
    param (
        [Parameter(Mandatory)][ValidateScript({
            $_ -in (Get-SnipeLocations).name
        })][string]$location,
        [Parameter(Mandatory)][string]$displayname,
        [int]$copies = 1,
        [string]$mobile
    )

    #Get the selected locations address
    $locationResult = Get-SnipeLocation -name $location

    $locationName = $locationResult.address
    $address = $locationResult.address2
    $postalCode = $locationResult.zip
    $city = $locationResult.city

    if ($displayname -and !$mobile) {
        #Install RSAT
        Install-RSAT -WUServerBypass

        #Search AD for phonenumber
        [string]$mobile = (Get-ADUser -Filter {DisplayName -like $displayname} -Properties mobile).mobile
    }

    if ($displayname -and $location) {
        #Get current default printer
        $defaultPrinter = (Get-CimInstance -Class Win32_Printer).where{$_.Default -eq $true}.Name

        #Set printer named PR-STORLABEL-SSDSK to default printer
        if (!(Get-Printer -Name "\\sr-safecom-sla1\PR-STORLABEL-SSDSK" -ErrorAction SilentlyContinue)) {
            (New-Object -ComObject WScript.Network).AddWindowsPrinterConnection("\\sr-safecom-sla1\PR-STORLABEL-SSDSK")
        }
        (New-Object -ComObject WScript.Network).SetDefaultPrinter("\\sr-safecom-sla1\PR-STORLABEL-SSDSK")

        #Create new Word document
        $WordObj = New-Object -ComObject Word.Application
        $WordObj.Documents.Add() | Out-Null

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

        <#
            https://books.google.no/books?id=rbpNppFdipkC&pg=PT114&lpg=PT114&dq=Application.PrintOut+%22copies%22+%22powershell%22&source=bl&ots=5_iiXja8EA&sig=ACfU3U11_KmhwFHlsOgEXNFcSHXx3rTvww&hl=en&sa=X&ved=2ahUKEwi-td6gzcb6AhXwlosKHTtAA9IQ6AF6BAgsEAM#v=onepage&q=Application.PrintOut%20%22copies%22%20%22powershell%22&f=false
            https://learn.microsoft.com/en-us/office/vba/api/word.application.printout#parameters
            I didn't manage to get it working through using the COM-object "copies" parameter, so I used a workaround within PowerShell
        #>
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