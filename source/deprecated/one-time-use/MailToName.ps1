$testflow = Import-Excel Z:/'Test Flow.xlsx'
$allmails = $testflow.Ansatt
$namearray = New-Object -TypeName System.Collections.ArrayList
foreach ($mail in $allmails) {
    $displayname = Get-ADUser -Filter {mail -eq $mail}
    $namearray.Add($displayname.givenname + " " + $displayname.surname)
}
Out-File -FilePath Z:/out.txt -InputObject $namearray