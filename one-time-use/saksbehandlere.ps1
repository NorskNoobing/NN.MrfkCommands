$allcaseworkers = Import-Excel -Path "Z:\saksbehandlere pureservice.xlsx"
$OutFile = "Z:\saksbehandlere.csv"
$DoesOutfileExist = Test-Path -Path $OutFile -PathType Leaf
if ($DoesOutfileExist -eq $false) {
    New-Item -Path $OutFile -ItemType File | Out-Null
    Set-Content -Path $OutFile -Value "Saksbehandler,Brukernavn,Avdeling,Enabled"
}

foreach ($displayname in $allcaseworkers) {
    $name = $displayname.name
    $ADobject = Get-ADUser -Filter{DisplayName -eq $name} -Properties Department
    $department = $null
    $department = $ADobject.Department
    if ($department -like "*,*"){
        $department = $department -replace ","," "
    }

    if ($($ADobject.SamAccountName) -like "adm_*") {} 
    else {
        Add-Content -Path $OutFile -Value "$name,$($ADobject.SamAccountName),$department,$($ADobject.Enabled)"
    }
}