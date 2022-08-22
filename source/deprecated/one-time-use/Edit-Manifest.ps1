    <#
    .SYNOPSIS
        Changes the manifest of an exe file from "requireAdministrator" to "asInvoker". Essentially stop a program from forcing administrative privledges
    
    .PARAMETER inputpath
        .exe file of the program you want to change the manifest for.

    .PARAMETER reverse
        Reverses the command, so the user has to have admin privledges to open the program.

    .EXAMPLE
        Edit-Manifest -inputpath "C:\Program Files (x86)\Steam\steam.exe"

    .INPUTS
        String, Switch
    #>
# https://developer.microsoft.com/en-us/windows/downloads/windows-sdk/
function Edit-Manifest {
    param (
        [string]$inputpath,
        [switch]$reverse
    )

    $mtlocation = "C:\Program Files (x86)\Windows Kits\10\bin\10.0.22000.0\x86" #WINDOWS SDK IN $mtlocation

    Start-Process "$mtlocation\mt.exe" -inputResource:"$inputpath" -out:"$inputpath.manifest"
    $manifestcontent = Get-Content -Path "$inputpath.manifest"
    if ($reverse) {
        $manifestcontent = $manifestcontent.Replace("asInvoker","requireAdministrator")
    } else {
        $manifestcontent = $manifestcontent.Replace("requireAdministrator","asInvoker")
    }
    $manifestcontent | Out-File "$inputpath.manifest"
    Start-Process "$mtlocation\mt.exe" -manifest "$inputpath.manifest" -outputResource:"$inputpath"
    Remove-Item -Path "$inputpath.manifest"
}