<#
.SYNOPSIS
    Create a Declarative Copilot app package.
.DESCRIPTION
    This cmdlet creates a Declarative Copilot app package, which includes a Teams app with a Declarative Copilot, you can sideload for yourself, or publish to your organization and share with more people. 
    
    You can learn more about Declarative Copilot here (https://learn.microsoft.com/en-us/microsoft-365-copilot/extensibility/overview-declarative-copilot).
.PARAMETER author
    The author of the Declarative Copilot app. It is optional, if not specified, the author will be Ares Chen.
.PARAMETER name
    The name of the Declarative Copilot.
.PARAMETER description
    The description of the Declarative Copilot.
.PARAMETER instructions
    The instructions of the Declarative Copilot, this is very important to guide the user to use the Declarative Copilot. Think this as the system prompt when you interact with the ChatGPT or other AI models.
.PARAMETER outlineIcon192x192
    The path of the outline icon of the Declarative Copilot, it should be a 192x192 png file. It is optional, if not specified, the default icon will be used.
.PARAMETER colorIcon32x32
    The path of the color icon of the Declarative Copilot, it should be a 32x32 png file. It is optional, if not specified, the default icon will be used.
.PARAMETER starterPrompts
    The starter prompts of the Declarative Copilot. Think this as templates of prompts you suggest to the user so that they can start the conversation easily.

    You can define at most 6 starter prompts. Each one should be a string, format as "title,text", for example, "Create a new document,Create a new document in Word". It means the title is "Create a new document" and the text is "Create a new document in Word". If you don't specify "," the title and text will be the same.
.PARAMETER enableWebSearch
    Enable the Web Search capability, so that the Declarative Copilot can search the web (for example, query the weather or the recent news) for the user.

.PARAMETER enableGraphicArt
    Enable the Graphic Art capability which allows the user to generate image or video.

.PARAMETER enableCodeInterpreter
    Enable the Code Interpreter capability which allows the user to run code.
.PARAMETER onedriveOrSharePointUrls
    The URLs of the OneDrive or SharePoint files which the Declarative Copilot can access. You can specify multiple URLs, the Url can be a SharePoint site, a SharePoint document library, or a OneDrive folder.

    For example, "https://contoso.sharepoint.com/sites/teamsite", "https://contoso-my.sharepoint.com/personal/user_contoso_com", "https://contoso-my.sharepoint.com/personal/user_contoso_com/Documents/Shared%20with%20Everyone".

    Please make sure the user who uses the Declarative Copilot has the permission to access the files in the specified URLs.

.PARAMETER graphConnectorIds
    The IDs of the Graph Connectors which the Declarative Copilot can access. You can learn more about Graph Connectors here (https://learn.microsoft.com/en-us/microsoft-365-copilot/extensibility/overview-graph-connector).

.PARAMETER actionFiles
    The action (plugin) files which the Declarative Copilot can access. 

.EXAMPLE
    New-DeclarativeCopilot -name "Product Copilot" -instructions "You are an experienced product manager, you help users to ideation, planning, and delivering great product from zero to one." 

    This is the simplest example, since only the name and instructions parameter are mandatory for this command.

.EXAMPLE
    New-DeclarativeCopilot -name "Product Copilot" -instructions "You are an experienced product manager, you help users to ideation, planning, and delivering great product from zero to one." -starterPrompts "Write PM spec, Please help me write spec about the idea below`n" 

    This example creates a Declarative Copilot app package named "Product Copilot" with the instructions "You are an experienced product manager, you help users to ideation, planning, and delivering great product from zero to one." and a starter prompt "Write PM spec, Please help me write spec about the idea below". 

.EXAMPLE
    New-DeclarativeCopilot -name "Product Copilot" -instructions "You are an experienced product manager, you help users to ideation, planning, and delivering great product from zero to one." -starterPrompts "Write PM spec, Please help me write spec about the idea below`n" -enableWebSearch -enableGraphicArt -enableCodeInterpreter

    This example creates a Declarative Copilot app package named "Product Copilot" with the instructions "You are an experienced product manager, you help users to ideation, planning, and delivering great product from zero to one." and a starter prompt "Write PM spec, Please help me write spec about the idea below". It also enables the Web Search, Graphic Art, and Code Interpreter capabilities.

.EXAMPLE
    New-DeclarativeCopilot -name "Product Copilot" -instructions "You are an experienced product manager, you help users to ideation, planning, and delivering great product from zero to one." -starterPrompts "Write PM spec, Please help me write spec about the idea below`n" -enableWebSearch -enableGraphicArt -enableCodeInterpreter -onedriveOrSharePointUrls "https://contoso.sharepoint.com/sites/teamsite", "https://contoso-my.sharepoint.com/personal/user_contoso_com", "https://contoso-my.sharepoint.com/personal/user_contoso_com/Documents/Shared%20with%20Everyone"

    This example creates a Declarative Copilot app package named "Product Copilot" with the instructions "You are an experienced product manager, you help users to ideation, planning, and delivering great product from zero to one." and a starter prompt "Write PM spec, Please help me write spec about the idea below". It also enables the Web Search, Graphic Art, and Code Interpreter capabilities, and specifies the OneDrive or SharePoint URLs.

.EXAMPLE
    New-DeclarativeCopilot -name "Product Copilot" -instructions "You are an experienced product manager, you help users to ideation, planning, and delivering great product from zero to one." -starterPrompts "Write PM spec, Please help me write spec about the idea below`n" -enableWebSearch -enableGraphicArt -enableCodeInterpreter -onedriveOrSharePointUrls "https://contoso.sharepoint.com/sites/teamsite", "https://contoso-my.sharepoint.com/personal/user_contoso_com", "https://contoso-my.sharepoint.com/personal/user_contoso_com/Documents/Shared%20with%20Everyone" -graphConnectorIds "12345678-abcd-1234-abcd-1234567890ab", "23456789-abcd-1234-abcd-1234567890ab"

    This example creates a Declarative Copilot app package named "Product Copilot" with the instructions "You are an experienced product manager, you help users to ideation, planning, and delivering great product from zero to one." and a starter prompt "Write PM spec, Please help me write spec about the idea below". It also enables the Web Search, Graphic Art, and Code Interpreter capabilities, specifies the OneDrive or SharePoint URLs, and specifies the Graph Connector IDs.

.EXAMPLE

    New-DeclarativeCopilot -name "Product Copilot" -instructions "You are an experienced product manager, you help users to ideation, planning, and delivering great product from zero to one." -starterPrompts "Write PM spec, Please help me write spec about the idea below`n" -enableWebSearch -enableGraphicArt -enableCodeInterpreter -onedriveOrSharePointUrls "https://contoso.sharepoint.com/sites/teamsite", "https://contoso-my.sharepoint.com/personal/user_contoso_com", "https://contoso-my.sharepoint.com/personal/user_contoso_com/Documents/Shared%20with%20Everyone" -graphConnectorIds "12345678-abcd-1234-abcd-1234567890ab", "23456789-abcd-1234-abcd-1234567890ab" -actionFiles "C:\path\to\action1.json", "C:\path\to\action2.json"

    This example creates a Declarative Copilot app package named "Product Copilot" with the instructions "You are an experienced product manager, you help users to ideation, planning, and delivering great product from zero to one." and a starter prompt "Write PM spec, Please help me write spec about the idea below". It also enables the Web Search, Graphic Art, and Code Interpreter capabilities, specifies the OneDrive or SharePoint URLs, specifies the Graph Connector IDs, and specifies the action files.

.EXAMPLE
    New-DeclarativeCopilot -name "Product Copilot" -instructions "You are an experienced product manager, you help users to ideation, planning, and delivering great product from zero to one."  -outlineIcon192x192 "C:\path\to\outline.png" -colorIcon32x32 "C:\path\to\color.png" -author "Your name"

    This example creates a Declarative Copilot app package named "Product Copilot" with the instructions "You are an experienced product manager, you help users to ideation, planning, and delivering great product from zero to one." and specifies the outline icon, color icon, and author.

.Link
    https://github.com/code365opensource/microsoft.copilot.toolkit
#>
function New-DeclarativeCopilot {
    [CmdletBinding()][Alias("ndc")]
    param (
        [string]$author,
        [Parameter(Mandatory = $true)]
        [string]$name,
        [string]$description,
        [Parameter(Mandatory = $true)]
        [string]$instructions,
        [string]$outlineIcon192x192,
        [string]$colorIcon32x32,
        [string[]]$starterPrompts,
        [switch]$enableWebSearch,
        [switch]$enableGraphicArt,
        [switch]$enableCodeInterpreter,
        [string[]]$onedriveOrSharePointUrls,
        [string[]]$graphConnectorIds,
        [string[]]$actionFiles
    )

    Send-AppInsightsTrace -Message "microsoft.copilot.toolkit" -Properties @{ "command" = "New-DeclarativeCopilot" } -ErrorAction SilentlyContinue
    
    # copy the content of private\assets\declarativecopilot to the temp folder
    $tempFolder = Join-Path $env:TEMP "microsoft.copilot.toolkit-$([Guid]::NewGuid().ToString())"

    $sourceFolder = Join-Path $PSScriptRoot -ChildPath "../private/assets/declarativecopilot"
    Copy-Item -Path $sourceFolder -Destination $tempFolder -Recurse -Force

    # update the manifest file
    $manifest = Get-Content (Join-Path $tempFolder "manifest.json") | ConvertFrom-Json
    $manifest.id = [Guid]::NewGuid().ToString()

    $manifest.name.short = $name
    $manifest.name.full = $name
    $manifest.description.short = "A Teams app inclduing a Declarative Copilot, $name"
    $manifest.description.full = "A Teams app inclduing a Declarative Copilot, $name, generated by the microsoft.copilot.toolkit module (https://www.powershellgallery.com/packages/microsoft.copilot.toolkit), created by Ares Chen."

    # update the author
    if ($author) {
        $manifest.developer.name = $author
    }

    # save the updated manifest file to the same file
    $manifest | ConvertTo-Json -Depth 10 | Set-Content -Path (Join-Path $tempFolder "manifest.json") -Force -Encoding UTF8

    # update the outline icon
    if ($outlineIcon192x192 -and (Test-Path $outlineIcon192x192) -and ($outlineIcon192x192 -match "\.png$")) {
        # make sure it is a 192x192 png file
        $outlineIcon = Join-Path $tempFolder "outline.png"
        Copy-Item -Path $outlineIcon192x192 -Destination $outlineIcon -Force
    }

    # update the color icon
    if ($colorIcon32x32 -and (Test-Path $colorIcon32x32) -and ($colorIcon32x32 -match "\.png$")) {
        $colorIcon = Join-Path $tempFolder "color.png"
        Copy-Item -Path $colorIcon32x32 -Destination $colorIcon -Force
    }

    # update the content of the declarativecopilot.json
    $copilot = Get-Content (Join-Path $tempFolder "declarativecopilot.json") | ConvertFrom-Json
    $copilot.name = $name
    $copilot.instructions = $instructions

    # parse the starter prompts
    if ($starterPrompts -and $starterPrompts.Count -gt 0) {       
        $copilot | Add-Member -MemberType NoteProperty -Name "conversation_starters" -Value @(
            foreach ($starterPrompt in $starterPrompts) {
                $parts = $starterPrompt.Split(",")
                if ($parts.Count -eq 1) {
                    [ordered]@{
                        "title" = $starterPrompt
                        "text"  = $starterPrompt
                    }
                }
                else {
                    [ordered]@{
                        "title" = $parts[0]
                        "text"  = $starterPrompt.SubString($parts[0].Length + 1)
                    }
                }
            }
        )
    }

    $capabilities = @()
    if ($enableWebSearch) {
        $capabilities += @{
            "name" = "WebSearch"
        }
    }

    if ($enableGraphicArt) {
        $capabilities += @{
            "name" = "GraphicArt"
        }
    }

    if ($enableCodeInterpreter) {
        $capabilities += @{
            "name" = "CodeInterpreter"
        }
    }

    if ($onedriveOrSharePointUrls -and $onedriveOrSharePointUrls.Count -gt 0) {
        $capabilities += [ordered]@{
            "name"         = "OneDriveAndSharePoint"
            "items_by_url" = @(
                foreach ($url in $onedriveOrSharePointUrls) {
                    @{
                        "url" = $url
                    }
                }
            )
        }
    }

    if ($graphConnectorIds -and $graphConnectorIds.Count -gt 0) {
        $capabilities += [ordered]@{
            "name"        = "GraphConnectors"
            "connections" = @(
                foreach ($id in $graphConnectorIds) {
                    @{
                        "connection_id" = $id
                    }
                }
            )
        }
    }

    # if any capibility is enabled, update the capabilities
    if ($capabilities.Count -gt 0) {
        $copilot | Add-Member -MemberType NoteProperty -Name "capabilities" -Value $capabilities
    }

    if ($actionFiles -and $actionFiles.Count -gt 0) {
        $actions = @(
            foreach ($actionFile in $actionFiles) {
                # copy the action file to the temp folder
                $actionFileName = Split-Path $actionFile -Leaf
                Copy-Item -Path $actionFile -Destination (Join-Path $tempFolder -ChildPath $actionFileName) -Force

                [ordered]@{
                    "id"   = [Guid]::NewGuid().ToString()
                    "file" = $actionFileName
                }
            }
        )

        $copilot | Add-Member -MemberType NoteProperty -Name "actions" -Value $actions
    }

    # save the updated declarativecopilot.json to the same file
    $copilot | ConvertTo-Json -Depth 10 | Set-Content -Path (Join-Path $tempFolder "declarativecopilot.json") -Force -Encoding UTF8


    # create a zip file and save it in current folder, named as $name.zip
    $zipFile = "$name.zip"

    # compress the temp folder to a zip file
    Compress-Archive -Path "$tempFolder\\*" -DestinationPath $zipFile -Force


    # remove the temp folder
    Remove-Item -Path $tempFolder -Recurse -Force

    Write-Host "The Declarative Copilot app has been created successfully. The zip file is saved as $zipFile."

}