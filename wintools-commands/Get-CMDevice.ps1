function Get-CMDevice-wt {
    Invoke-Command -Credential $(Get-AdmCreds) -ComputerName wintools04 -ScriptBlock {
        import-module 'C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1'
        set-location ps1:
        #BUG: Import-Module: The specified module was not loaded because no valid module file was found in any module directory.

        Get-CMDevice @args
    }
}