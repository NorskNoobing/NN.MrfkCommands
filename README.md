# NN.MrfkCommands
![PowerShell Gallery](https://img.shields.io/powershellgallery/dt/NN.MrfkCommands) ![GitHub last commit](https://img.shields.io/github/last-commit/NorskNoobing/NN.MrfkCommands)

## Prerequisites
This module requires RSAT in many of the custom functions. It can be installed by running PowerShell as admin, and pasting the following.
```powershell
Get-WindowsCapability -Name "RSAT*" -Online | Add-WindowsCapability -Online
```
## Installation
Run the following command in your PowerShell terminal to install the module.
```powershell
Install-Module NN.MrfkCommands -Repository PSGallery -Force
```
## Documentation
You can see the documentation for all the functions on the wiki tab in this repo.
