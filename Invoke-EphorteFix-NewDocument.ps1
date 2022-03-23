function Invoke-EphorteFix-NewDocument {
    param (
        [Parameter(Mandatory)][string]$hostname
    )
    
    Invoke-Command -ComputerName $hostname -ScriptBlock {
        if (!(Test-Path -Path "HKLM:\SOFTWARE\Microsoft\.NETFramework")) {
            New-Item -Force -Path "HKLM:\SOFTWARE\Microsoft\.NETFramework"
        }

        if (!(Test-Path -Path "HKLM:\SOFTWARE\Wow6432Node\Microsoft\.NETFramework")) {
            New-Item -Force -Path "HKLM:\SOFTWARE\Wow6432Node\Microsoft\.NETFramework"
        }

        New-ItemProperty -Force -Path "HKLM:\SOFTWARE\Microsoft\.NETFramework" -Name "EnableIEHosting" -PropertyType "dword" -Value "1"
        New-ItemProperty -Force -Path "HKLM:\SOFTWARE\Wow6432Node\Microsoft\.NETFramework" -Name "EnableIEHosting" -PropertyType "dword" -Value "1"
    }
}