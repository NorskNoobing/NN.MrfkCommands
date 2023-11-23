function Invoke-MrfkShippingLabelPrint {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)][string]$Recipient,
        [Parameter(Mandatory)][string]$Location,
        [int]$copies = 1,
        [string]$PrinterNetworkPath = "\\print01\PR-SIT-STORLABEL"
    )

    process {
        #Get current default printer
        $defaultPrinter = (Get-CimInstance -Class Win32_Printer).where({ $_.Default -eq $true }).Name

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
        $WordObj.Selection.TypeText("Mottaker:`r`n")
        $WordObj.Selection.Font.Bold = 0
        $WordObj.Selection.TypeText($Recipient)
        $WordObj.Selection.TypeParagraph()
        $WordObj.Selection.Font.Bold = 1
        $WordObj.Selection.TypeText("Addresse:`r`n")
        $WordObj.Selection.Font.Bold = 0
        $WordObj.Selection.TypeText($Location)

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