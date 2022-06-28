function New-ShippingLabel {
    param (
        [ValidateSet("Atlanten VGS","Borgund VGS","Fagerlia VGS","Fagskolen i Alesund",
        "Gjermundnes VGS","Haram VGS","Herøy VGS","Hustadvika VGS","Kristiansund VGS","Rauma VGS","Romsdal VGS",
        "Spjelkavik VGS","Sunndal VGS","Sykkylven VGS","Surnadal VGS","Tingvoll VGS","Ulstein VGS","Volda VGS","Orsta VGS",
        "Alesund VGS (Volsdalsberga)","Campus Kristiansund","Olsvika")][string]$location,
        [string]$displayname
    )

    if (!$location -and !$displayname) {
            #UI   
            Add-Type -AssemblyName System.Windows.Forms
            Add-Type -AssemblyName System.Drawing
            
            $form = New-Object System.Windows.Forms.Form
            $form.Text = 'Shipping label'
            $form.Size = New-Object System.Drawing.Size(300,250)
            $form.StartPosition = 'CenterScreen'
            
            $okButton = New-Object System.Windows.Forms.Button
            $okButton.Location = New-Object System.Drawing.Point(75,180)
            $okButton.Size = New-Object System.Drawing.Size(75,23)
            $okButton.Text = 'Print'
            $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $form.AcceptButton = $okButton
            $form.Controls.Add($okButton)
            
            $cancelButton = New-Object System.Windows.Forms.Button
            $cancelButton.Location = New-Object System.Drawing.Point(150,180)
            $cancelButton.Size = New-Object System.Drawing.Size(75,23)
            $cancelButton.Text = 'Cancel'
            $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $form.CancelButton = $cancelButton
            $form.Controls.Add($cancelButton)
            
            $label = New-Object System.Windows.Forms.Label
            $label.Location = New-Object System.Drawing.Point(10,20)
            $label.Size = New-Object System.Drawing.Size(280,20)
            $label.Text = 'Displayname:'
            $form.Controls.Add($label)
            
            $textBox = New-Object System.Windows.Forms.TextBox
            $textBox.Location = New-Object System.Drawing.Point(10,40)
            $textBox.Size = New-Object System.Drawing.Size(260,20)
            $form.Controls.Add($textBox)
            
            $label = New-Object System.Windows.Forms.Label
            $label.Location = New-Object System.Drawing.Point(10,70)
            $label.Size = New-Object System.Drawing.Size(280,20)
            $label.Text = 'Location:'
            $form.Controls.Add($label)
            
            $listBox = New-Object System.Windows.Forms.ListBox
            $listBox.Location = New-Object System.Drawing.Point(10,90)
            $listBox.Size = New-Object System.Drawing.Size(260,20)
            $listBox.Height = 80
            
            [void] $listBox.Items.Add("Atlanten VGS")
            [void] $listBox.Items.Add("Borgund VGS")
            [void] $listBox.Items.Add("Campus Kristiansund")
            [void] $listBox.Items.Add("Fagerlia VGS")
            [void] $listBox.Items.Add("Fagskolen i Alesund")
            [void] $listBox.Items.Add("Gjermundnes VGS")
            [void] $listBox.Items.Add("Haram VGS")
            [void] $listBox.Items.Add("Herøy VGS")
            [void] $listBox.Items.Add("Hustadvika VGS")
            [void] $listBox.Items.Add("Kristiansund VGS")
            [void] $listBox.Items.Add("Olsvika")
            [void] $listBox.Items.Add("Rauma VGS")
            [void] $listBox.Items.Add("Romsdal VGS")
            [void] $listBox.Items.Add("Spjelkavik VGS")
            [void] $listBox.Items.Add("Sunndal VGS")
            [void] $listBox.Items.Add("Sykkylven VGS")
            [void] $listBox.Items.Add("Surnadal VGS")
            [void] $listBox.Items.Add("Tingvoll VGS")
            [void] $listBox.Items.Add("Ulstein VGS")
            [void] $listBox.Items.Add("Volda VGS")
            [void] $listBox.Items.Add("Orsta VGS")
            [void] $listBox.Items.Add("Alesund VGS (Volsdalsberga)")
            
            $form.Controls.Add($listBox)
            
            $form.Topmost = $true
            
            $form.Add_Shown({$textBox.Select()})
            $result = $form.ShowDialog()
            
            if ($result -eq [System.Windows.Forms.DialogResult]::OK)
            {
                $displayname = $textBox.Text
                $location = $listBox.Text
            }
    }

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
                $locationName = "Ålesund videregående skole"
                $address = "Sjømannsvegen 27"
                $postalCode = "6008"
                $city = "Ålesund"
            }
        }

    if ($displayname) {
        #Enable WU
        New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" -Force
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" -name "DisableDualScan" -value 0
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" -name "DisableWindowsUpdateAccess" -value 0
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" -name "SetPolicyDrivenUpdateSourceForDriverUpdates" -value 0
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" -name "SetPolicyDrivenUpdateSourceForFeatureUpdates" -value 0
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" -name "SetPolicyDrivenUpdateSourceForOtherUpdates" -value 0
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" -name "SetPolicyDrivenUpdateSourceForQualityUpdates" -value 0
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" -name "UseWUServer" -value 0
        Restart-Service wuauserv

        $RSATState = Get-WindowsCapability -Name RSAT* -online

        foreach ($item in $RSATState) {
            if (!($item.State -eq "Installed")) {
                Get-WindowsCapability -Name $item.Name -Online | Add-WindowsCapability -Online
            }
        }

        #Search AD for phonenumber
        [string]$mobile = (Get-ADUser -Filter {DisplayName -like $displayname} -Properties mobile).mobile
    }

    if ($displayname -and $location) {
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

        #Close File without saving
        $WordObj.ActiveDocument.Close(0)
        $WordObj.quit() 
    }
}