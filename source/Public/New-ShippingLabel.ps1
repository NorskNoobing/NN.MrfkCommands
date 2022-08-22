function New-ShippingLabel {
    param (
        [ValidateSet("Atlanten VGS","Borgund VGS","Fagerlia VGS","Fagskolen i Alesund",
        "Gjermundnes VGS","Haram VGS","Herøy VGS","Hustadvika VGS","Kristiansund VGS","Rauma VGS","Romsdal VGS",
        "Spjelkavik VGS","Sunndal VGS","Sykkylven VGS","Surnadal VGS","Tingvoll VGS","Ulstein VGS","Volda VGS","Orsta VGS",
        "Alesund VGS (Volsdalsberga)","Alesund VGS (Fagerlia)","Campus Kristiansund","Olsvika","Carolus","Fylkeshuset")][string]$location,
        [string]$displayname,
        [string]$mobile
    )

    #LOCATION
    #Get the selected locations address 
    switch ($location) {
        "Atlanten VGS" { 
            $locationName = "Atlanten videregående skole"
            $address = "Dalaveien 25"
            $postalCode = "6511"
            $city = "Kristiansund"
        }
        "Borgund VGS" {
            $locationName = "Borgund vidaregåande skule"
            $address = "Yrkesskolevegen 20"
            $postalCode = "6011"
            $city = "Ålesund"
        }
        "Campus Kristiansund" {
            $locationName = "Kristiansund Rådhus Campus"
            $address = "Fosnagata 13"
            $postalCode = "6509"
            $city = "Kristiansund"
        }
        "Fagskolen i Alesund" {
            $locationName = "Fagskolen Møre og Romsdal"
            $address = "Fogd Greves veg 9"
            $postalCode = "6009"
            $city = "Ålesund"
        }
        "Gjermundnes VGS" {
            $locationName = "Gjermundnes vidaregåande skule"
            $address = "Gjermundnesvegen 200"
            $postalCode = "6392"
            $city = "Vikebukt"
        }
        "Haram VGS" {
            $locationName = "Haram vidaregåande skule"
            $address = "Skuleråsa 10"
            $postalCode = "6270"
            $city = "Brattvåg"
        }
        "Herøy VGS" {
            $locationName = "Herøy vidaregåande skule"
            $address = "Lisjebøveien 4"
            $postalCode = "6091"
            $city = "Fosnavåg"
        }
        "Hustadvika VGS" {
            $locationName = "Hustadvika vidaregåande skole"
            $address = "Bøen 22"
            $postalCode = "6440"
            $city = "Elnesvågen"
        }
        "Kristiansund VGS" {
            $locationName = "Kristiansund videregående skole"
            $address = "Sankthanshaugen 2"
            $postalCode = "6514"
            $city = "Kristiansund"
        }
        "Olsvika" {
            $locationName = "Møre og Romsdal Fylkeskommune"
            $address = "Vestre Olsvikveg 13"
            $postalCode = "6019"
            $city = "Ålesund"
        }
        "Rauma VGS" {
            $locationName = "Rauma videregående skole"
            $address = "Ringgata 35"
            $postalCode = "6300"
            $city = "Åndalsnes"
        }
        "Romsdal VGS" {
            $locationName = "Romsdal videregående skole"
            $address = "Langmyrvegen 83"
            $postalCode = "6415"
            $city = "Molde"
        }
        "Spjelkavik VGS" {
            $locationName = "Spjelkavik vidaregåande skole"
            $address = "Nedre Langhaugen 32"
            $postalCode = "6011"
            $city = "Ålesund"
        }
        "Sunndal VGS" {
            $locationName = "Sunndal VGS"
            $address = "Skoleveien 14"
            $postalCode = "6600"
            $city = "Sunndalsøra"
        }
        "Sykkylven VGS" {
            $locationName = "Sykkylven vidaregåande skule"
            $address = "Kyrkjevegen 6"
            $postalCode = "6230"
            $city = "Sykkylven"
        }
        "Surnadal VGS" {
            $locationName = "Surnadal vidaregåande skole"
            $address = "Øyatrøvegen 30"
            $postalCode = "6650"
            $city = "Surnadal"
        }
        "Tingvoll VGS" {
            $locationName = "Tingvoll videregående skole"
            $address = "Skolevegen 35"
            $postalCode = "6630"
            $city = "Tingvoll"
        }
        "Ulstein VGS" {
            $locationName = "Ulstein vidaregåande skule"
            $address = "Holsekerdalen 180"
            $postalCode = "6065"
            $city = "Ulsteinvik"
        }
        "Volda VGS" {
            $locationName = "Volda vidaregåande skule"
            $address = "Vevendelvegen 35"
            $postalCode = "6102"
            $city = "Volda"
        }
        "Orsta VGS" {
            $locationName = "Ørsta vidaregåande skule"
            $address = "Holmegata 14"
            $postalCode = "6153"
            $city = "Ørsta"
        }
        "Alesund VGS (Volsdalsberga)" {
            $locationName = "Ålesund videregående skole avd. Volsdalsberga"
            $address = "Sjømannsvegen 27"
            $postalCode = "6008"
            $city = "Ålesund"
        }
        "Alesund VGS (Fagerlia)" {
            $locationName = "Ålesund videregående skole avd. Fagerlia"
            $address = "Gangstøvikvegen 27"
            $postalCode = "6009"
            $city = "Ålesund"
        }
        "Carolus" {
            $locationName = "Møre og Romsdal Fylkeskommune"
            $address = "Bjørnstjerne Bjørnsons veg 6"
            $postalCode = "6412"
            $city = "Molde"
        }
        "Fylkeshuset" {
            $locationName = "Møre og Romsdal Fylkeskommune"
            $address = "Julsundvegen 9"
            $postalCode = "6412"
            $city = "Molde"
        }
    }

    if ($displayname -and !$mobile) {
        <#
            todo: check for RSAT without needing elevation
        #>
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

        #Send To Default Printer
        $WordObj.PrintOut()

        #Change default printer back to the previous value
        (New-Object -ComObject WScript.Network).SetDefaultPrinter("$defaultPrinter")

        #Close File without saving
        $WordObj.ActiveDocument.Close(0)
        $WordObj.quit() 
    }
}