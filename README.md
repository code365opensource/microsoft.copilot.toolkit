# The readme about yourmodulename

> [!NOTE]
> Replace `yourmodulename` with the name of the module, and delete this note. change the GUID in the psd1 file as well.
> Follow the guidance (https://learn.microsoft.com/en-us/powershell/gallery/concepts/publishing-guidelines) to create a good module.

## Description


[![PowerShell Gallery Version](https://img.shields.io/powershellgallery/v/yourmodulename?label=yourmodulename)](https://www.powershellgallery.com/packages/yourmodulename) [![PowerShell Gallery](https://img.shields.io/powershellgallery/dt/yourmodulename)](https://www.powershellgallery.com/packages/yourmodulename) [![](https://img.shields.io/badge/change-logs-blue)](CHANGELOG.md) ![https://img.shields.io/powershellgallery/p/yourmodulename.svg](https://img.shields.io/powershellgallery/p/yourmodulename.svg)


## Install the module

```powershell
Install-Module -Name yourmodulename -Scope CurrentUser

# or if you want to install the preview version
Install-Module -Name yourmodulename -Scope CurrentUser -AllowPrerelease
```

## Update the module

```powershell
Update-Module -Name yourmodulename -Scope CurrentUser
```

## Uninstall the module

```powershell
Uninstall-Module -Name yourmodulename -Scope CurrentUser
```