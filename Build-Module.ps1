#Requires -Module ModuleBuilder
[string]$moduleName = "mrfk-commands"
[version]$version = "0.0.1"
[string]$author = "NorskNoobing"
[string]$ProjectUri = "https://github.com/$author/$moduleName"
[string]$releaseNotes = "Initial commit"
[string]$description = "Automate the workflow of caseworkers"
[array]$tags = @("MRFK","Mrfylke","Møre og Romsdal Fylkeskommune")
[version]$PSversion = "7.2"

$manifestSplat = @{
    "Description" = $description
    "PowerShellVersion" = $PSversion
    "Tags" = $tags
    "ReleaseNotes" = $releaseNotes
    "Path" = "$PSScriptRoot\source\$moduleName.psd1"
    "RootModule" = "$moduleName.psm1"
    "Author" = $author
    "ProjectUri" = $ProjectUri
}
New-ModuleManifest @manifestSplat

$buildSplat = @{
    "SourcePath" = "$PSScriptRoot\source\$moduleName.psd1"
    "Version" = $version
}
Build-Module @buildSplat