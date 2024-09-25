The unofficial PowerShell module for Microsoft Copilot, currently it provides the following features:

1. Create your own [Declarative Copilot](https://learn.microsoft.com/en-us/microsoft-365-copilot/extensibility/overview-declarative-copilot) by using `New-DeclarativeCopilot` cmdlet or `ndc` alias.

[![PowerShell Gallery Version](https://img.shields.io/powershellgallery/v/microsoft.copilot.toolkit?label=microsoft.copilot.toolkit)](https://www.powershellgallery.com/packages/microsoft.copilot.toolkit) [![PowerShell Gallery](https://img.shields.io/powershellgallery/dt/microsoft.copilot.toolkit)](https://www.powershellgallery.com/packages/microsoft.copilot.toolkit) [![](https://img.shields.io/badge/change-logs-blue)](CHANGELOG.md) ![https://img.shields.io/powershellgallery/p/microsoft.copilot.toolkit.svg](https://img.shields.io/powershellgallery/p/microsoft.copilot.toolkit.svg)


## Install the module

```powershell
Install-Module -Name microsoft.copilot.toolkit -Scope CurrentUser

```
> [!TIP]
> if this is the first time you use PowerShell, you should trust the official PSGallery so that you can download our module from the gallery
> `Set-PSRepository -InstallationPolicy Trusted -Name PSGallery`
>
> You need to allow the RemoteSigned scripts to run on your machine, so our module won't be blocked.
> `Set-ExecutionPolicy RemoteSigned -Scope CurrentUser`


## Create your own Declarative Copilot

You can use the `New-DeclarativeCopilot` cmdlet to create your own Declarative Copilot app package. Below are some examples of how to use the `New-DeclarativeCopilot` cmdlet. When you get the app package, you can upload it to your Microsoft 365 Copilot environment. You can also download the [sample package](Private/assets/Product%20Copilot.zip) if you want to try it out first.

![](Private/assets/sideload.png)

Once you get it done, you can use the `Product Copilot` in your Copilot app.

![](Private/assets/productcopilot.jpg)



-------------------------- EXAMPLE 1 --------------------------

PS > New-DeclarativeCopilot -name "Product Copilot" -instructions "You are an experienced product manager, you help users to ideation, planning, and delivering great product from zero to one."

This is the simplest example, since only the name and instructions parameter are mandatory for this command.




-------------------------- EXAMPLE 2 --------------------------

PS > New-DeclarativeCopilot -name "Product Copilot" -instructions "You are an experienced product manager, you help users to ideation, planning, and delivering great product from zero to one." -starterPrompts "Write PM spec, Please help me write spec about the idea below`n"

This example creates a Declarative Copilot app package named "Product Copilot" with the instructions "You are an experienced product manager, you help users to ideation, planning, and delivering great product from zero to one." and a starter prompt "Write PM spec, Please help me write spec about the idea below".




-------------------------- EXAMPLE 3 --------------------------

PS > New-DeclarativeCopilot -name "Product Copilot" -instructions "You are an experienced product manager, you help users to ideation, planning, and delivering great product from zero to one." -starterPrompts "Write PM spec, Please help me write spec about the idea below`n" -enableWebSearch -enableGraphicArt -enableCodeInterpreter

This example creates a Declarative Copilot app package named "Product Copilot" with the instructions "You are an experienced product manager, you help users to ideation, planning, and delivering great product from zero to one." and a starter prompt "Write PM spec, Please help me write spec about the idea below". It also enables the Web Search, Graphic Art, and Code Interpreter capabilities.




-------------------------- EXAMPLE 4 --------------------------

PS > New-DeclarativeCopilot -name "Product Copilot" -instructions "You are an experienced product manager, you help users to ideation, planning, and delivering great product from zero to one." -starterPrompts "Write PM spec, Please help me write spec about the idea below`n" -enableWebSearch -enableGraphicArt -enableCodeInterpreter -onedriveOrSharePointUrls "https://contoso.sharepoint.com/sites/teamsite", "https://contoso-my.sharepoint.com/personal/user_contoso_com", "https://contoso-my.sharepoint.com/personal/user_contoso_com/Documents/Shared%20with%20Everyone"

This example creates a Declarative Copilot app package named "Product Copilot" with the instructions "You are an experienced product manager, you help users to ideation, planning, and delivering great product from zero to one." and a starter prompt "Write PM spec, Please help me write spec about the idea below". It also enables the Web Search, Graphic Art, and Code Interpreter capabilities, and specifies the OneDrive or SharePoint URLs.




-------------------------- EXAMPLE 5 --------------------------

PS > New-DeclarativeCopilot -name "Product Copilot" -instructions "You are an experienced product manager, you help users to ideation, planning, and delivering great product from zero to one." -starterPrompts "Write PM spec, Please help me write spec about the idea below`n" -enableWebSearch -enableGraphicArt -enableCodeInterpreter -onedriveOrSharePointUrls "https://contoso.sharepoint.com/sites/teamsite", "https://contoso-my.sharepoint.com/personal/user_contoso_com", "https://contoso-my.sharepoint.com/personal/user_contoso_com/Documents/Shared%20with%20Everyone" -graphConnectorIds "12345678-abcd-1234-abcd-1234567890ab", "23456789-abcd-1234-abcd-1234567890ab"

This example creates a Declarative Copilot app package named "Product Copilot" with the instructions "You are an experienced product manager, you help users to ideation, planning, and delivering great product from zero to one." and a starter prompt "Write PM spec, Please help me write spec about the idea below". It also enables the Web Search, Graphic Art, and Code Interpreter capabilities, specifies the OneDrive or SharePoint URLs, and specifies the Graph Connector IDs.




-------------------------- EXAMPLE 6 --------------------------

PS > New-DeclarativeCopilot -name "Product Copilot" -instructions "You are an experienced product manager, you help users to ideation, planning, and delivering great product from zero to one." -starterPrompts "Write PM spec, Please help me write spec about the idea below`n" -enableWebSearch -enableGraphicArt -enableCodeInterpreter -onedriveOrSharePointUrls "https://contoso.sharepoint.com/sites/teamsite", "https://contoso-my.sharepoint.com/personal/user_contoso_com", "https://contoso-my.sharepoint.com/personal/user_contoso_com/Documents/Shared%20with%20Everyone" -graphConnectorIds "12345678-abcd-1234-abcd-1234567890ab", "23456789-abcd-1234-abcd-1234567890ab" -actionFiles "C:\path\to\action1.json", "C:\path\to\action2.json"

This example creates a Declarative Copilot app package named "Product Copilot" with the instructions "You are an experienced product manager, you help users to ideation, planning, and delivering great product from zero to one." and a starter prompt "Write PM spec, Please help me write spec about the idea below". It also enables the Web Search, Graphic Art, and Code Interpreter capabilities, specifies the OneDrive or SharePoint URLs, specifies the Graph Connector IDs, and specifies the action files.




-------------------------- EXAMPLE 7 --------------------------

PS > New-DeclarativeCopilot -name "Product Copilot" -instructions "You are an experienced product manager, you help users to ideation, planning, and delivering great product from zero to one."  -outlineIcon192x192 "C:\path\to\outline.png" -colorIcon32x32 "C:\path\to\color.png" -author "Your name"

This example creates a Declarative Copilot app package named "Product Copilot" with the instructions "You are an experienced product manager, you help users to ideation, planning, and delivering great product from zero to one." and specifies the outline icon, color icon, and author.



## Update the module

```powershell
Update-Module -Name microsoft.copilot.toolkit -Scope CurrentUser
```

## Uninstall the module

```powershell
Uninstall-Module -Name microsoft.copilot.toolkit -Scope CurrentUser
```
