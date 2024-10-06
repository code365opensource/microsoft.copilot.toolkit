<#
.SYNOPSIS
    Create a Declarative Agent app package.
.DESCRIPTION
    This cmdlet creates a Declarative Agent app package, which includes a Teams app with a Declarative Agent, you can sideload for yourself, or publish to your organization and share with more people. 
    
    You can learn more about Declarative Agent here (https://learn.microsoft.com/en-us/microsoft-365-copilot/extensibility/overview-declarative-copilot).
.PARAMETER author
    The author of the Declarative Agent app. It is optional, if not specified, the author will be Ares Chen.
.PARAMETER name
    The name of the Declarative Agent.
.PARAMETER description
    The description of the Declarative Agent.
.PARAMETER instructions
    The instructions of the Declarative Agent, this is very important to guide the user to use the Declarative Agent. Think this as the system prompt when you interact with the ChatGPT or other AI models.

    You can specify the instructions as a string, or a file path which contains the instructions. If it is a file path, the content of the file will be used as the instructions. The file should be a UTF-8 encoded text file.
.PARAMETER outlineIcon192x192
    The path of the outline icon of the Declarative Agent, it should be a 192x192 png file. It is optional, if not specified, the default icon will be used.
.PARAMETER colorIcon32x32
    The path of the color icon of the Declarative Agent, it should be a 32x32 png file. It is optional, if not specified, the default icon will be used.
.PARAMETER starterPrompts
    The starter prompts of the Declarative Agent. Think this as templates of prompts you suggest to the user so that they can start the conversation easily.

    You can define at most 6 starter prompts. Each one should be a string, format as "title,text", for example, "Create a new document,Create a new document in Word". It means the title is "Create a new document" and the text is "Create a new document in Word". If you don't specify "," the title and text will be the same.
.PARAMETER enableWebSearch
    Enable the Web Search capability, so that the Declarative Agent can search the web (for example, query the weather or the recent news) for the user.

.PARAMETER enableGraphicArt
    Enable the Graphic Art capability which allows the user to generate image or video.

.PARAMETER enableCodeInterpreter
    Enable the Code Interpreter capability which allows the user to run code.
.PARAMETER onedriveOrSharePointUrls
    The URLs of the OneDrive or SharePoint files which the Declarative Agent can access. You can specify multiple URLs, the Url can be a SharePoint site, a SharePoint document library, or a OneDrive folder.

    For example, "https://contoso.sharepoint.com/sites/teamsite", "https://contoso-my.sharepoint.com/personal/user_contoso_com", "https://contoso-my.sharepoint.com/personal/user_contoso_com/Documents/Shared%20with%20Everyone".

    Please make sure the user who uses the Declarative Agent has the permission to access the files in the specified URLs.

    if you specify "all", it means the Declarative Agent can access all the files in the user's OneDrive or SharePoint sites you have permission to access.

.PARAMETER graphConnectorIds
    The IDs of the Graph Connectors which the Declarative Agent can access. You can learn more about Graph Connectors here (https://learn.microsoft.com/en-us/microsoft-365-copilot/extensibility/overview-graph-connector).

    if you specify "all", it means the Declarative Agent can access all the Graph Connectors you have permission to access.

.PARAMETER actionFiles
    The action (plugin) files which the Declarative Agent can access. 

.EXAMPLE
    New-DeclarativeAgent -name "Product Copilot" -instructions "You are an experienced product manager, you help users to ideation, planning, and delivering great product from zero to one." 

    This is the simplest example, since only the name and instructions parameter are mandatory for this command.

.EXAMPLE
    New-DeclarativeAgent -name "Product Copilot" -instructions "yourinstructions.txt"

    This example creates a Declarative Agent app package named "Product Copilot" with the instructions from the file "yourinstructions.txt".

.EXAMPLE
    New-DeclarativeAgent -name "Product Copilot" -instructions "You are an experienced product manager, you help users to ideation, planning, and delivering great product from zero to one." -starterPrompts "Write PM spec, Please help me write spec about the idea below`n" 

    This example creates a Declarative Agent app package named "Product Copilot" with the instructions "You are an experienced product manager, you help users to ideation, planning, and delivering great product from zero to one." and a starter prompt "Write PM spec, Please help me write spec about the idea below". 

.EXAMPLE
    New-DeclarativeAgent -name "Product Copilot" -instructions "You are an experienced product manager, you help users to ideation, planning, and delivering great product from zero to one." -starterPrompts "Write PM spec, Please help me write spec about the idea below`n" -enableWebSearch -enableGraphicArt -enableCodeInterpreter

    This example creates a Declarative Agent app package named "Product Copilot" with the instructions "You are an experienced product manager, you help users to ideation, planning, and delivering great product from zero to one." and a starter prompt "Write PM spec, Please help me write spec about the idea below". It also enables the Web Search, Graphic Art, and Code Interpreter capabilities.

.EXAMPLE
    New-DeclarativeAgent -name "Product Copilot" -instructions "You are an experienced product manager, you help users to ideation, planning, and delivering great product from zero to one." -starterPrompts "Write PM spec, Please help me write spec about the idea below`n" -enableWebSearch -enableGraphicArt -enableCodeInterpreter -onedriveOrSharePointUrls "https://contoso.sharepoint.com/sites/teamsite", "https://contoso-my.sharepoint.com/personal/user_contoso_com", "https://contoso-my.sharepoint.com/personal/user_contoso_com/Documents/Shared%20with%20Everyone"

    This example creates a Declarative Agent app package named "Product Copilot" with the instructions "You are an experienced product manager, you help users to ideation, planning, and delivering great product from zero to one." and a starter prompt "Write PM spec, Please help me write spec about the idea below". It also enables the Web Search, Graphic Art, and Code Interpreter capabilities, and specifies the OneDrive or SharePoint URLs.

.EXAMPLE
    New-DeclarativeAgent -name "Product Copilot" -instructions "You are an experienced product manager, you help users to ideation, planning, and delivering great product from zero to one." -starterPrompts "Write PM spec, Please help me write spec about the idea below`n" -enableWebSearch -enableGraphicArt -enableCodeInterpreter -onedriveOrSharePointUrls "https://contoso.sharepoint.com/sites/teamsite", "https://contoso-my.sharepoint.com/personal/user_contoso_com", "https://contoso-my.sharepoint.com/personal/user_contoso_com/Documents/Shared%20with%20Everyone" -graphConnectorIds "12345678-abcd-1234-abcd-1234567890ab", "23456789-abcd-1234-abcd-1234567890ab"

    This example creates a Declarative Agent app package named "Product Copilot" with the instructions "You are an experienced product manager, you help users to ideation, planning, and delivering great product from zero to one." and a starter prompt "Write PM spec, Please help me write spec about the idea below". It also enables the Web Search, Graphic Art, and Code Interpreter capabilities, specifies the OneDrive or SharePoint URLs, and specifies the Graph Connector IDs.


.EXAMPLE
    New-DeclarativeAgent -name "Product Copilot" -instructions "You are an experienced product manager, you help users to ideation, planning, and delivering great product from zero to one."  -outlineIcon192x192 "C:\path\to\outline.png" -colorIcon32x32 "C:\path\to\color.png" -author "Your name"

    This example creates a Declarative Agent app package named "Product Copilot" with the instructions "You are an experienced product manager, you help users to ideation, planning, and delivering great product from zero to one." and specifies the outline icon, color icon, and author.

.Link
    https://github.com/code365opensource/microsoft.copilot.toolkit
#>
function New-DeclarativeCopilot {
    [CmdletBinding()][Alias("ndc", "nda", "New-DeclarativeAgent")]
    param (
        [Parameter(ParameterSetName = "default")]
        [string]$author,
        [Parameter(ParameterSetName = "default", Mandatory = $true)]
        [string]$name,
        [Parameter(ParameterSetName = "default")]
        [string]$description,
        [Parameter(ParameterSetName = "default", Mandatory = $true)]
        [string]$instructions,
        [Parameter(ParameterSetName = "default")]
        [string]$outlineIcon192x192,
        [Parameter(ParameterSetName = "default")]
        [string]$colorIcon32x32,
        [Parameter(ParameterSetName = "default")]
        [string[]]$starterPrompts,
        [Parameter(ParameterSetName = "default")]
        [switch]$enableWebSearch,
        [Parameter(ParameterSetName = "default")]
        [switch]$enableGraphicArt,
        [Parameter(ParameterSetName = "default")]
        [switch]$enableCodeInterpreter,
        [Parameter(ParameterSetName = "default")]
        [string[]]$onedriveOrSharePointUrls,
        [Parameter(ParameterSetName = "default")]
        [string[]]$graphConnectorIds,
        [Parameter(ParameterSetName = "quickstart", Mandatory = $true)]
        [switch]$helloworld
    )

    Send-AppInsightsTrace -Message "microsoft.copilot.toolkit" -Properties @{ "command" = "New-DeclarativeAgent" } -ErrorAction SilentlyContinue
    
    # copy the content of private\assets\declarativecopilot to the temp folder
    $tempFolder = Join-Path $env:TEMP "microsoft.copilot.toolkit-$([Guid]::NewGuid().ToString())"

    $sourceFolder = Join-Path $PSScriptRoot -ChildPath "../private/assets/declarativecopilot"
    Copy-Item -Path $sourceFolder -Destination $tempFolder -Recurse -Force

    # if user specific the helloworld switch, then use the default values for instructions and name parameters
    if ($helloworld) {
        $instructions = "Hello, I am a Declarative Agent, I can help you with your daily work. You can ask me anything, I will try my best to help you."
        $name = "Hello World"
    }

    # update the manifest file
    $manifest = Get-Content (Join-Path $tempFolder "manifest.json") | ConvertFrom-Json
    $manifest.id = [Guid]::NewGuid().ToString()

    $manifest.name.short = $name
    $manifest.name.full = $name
    $manifest.description.short = "A Teams app inclduing a Declarative Agent, $name"
    $manifest.description.full = "A Teams app inclduing a Declarative Agent, $name, generated by the microsoft.copilot.toolkit module (https://www.powershellgallery.com/packages/microsoft.copilot.toolkit), created by Ares Chen."

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

    # if instructions is a file, read the content of the file
    if ($instructions -and (Test-Path $instructions -PathType Leaf)) {
        $instructions = Get-Content $instructions -Raw -Encoding UTF8 
    }
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
        # if $onedriveOrSharePointUrls contains "all", then add the capability "OneDriveAndSharePoint" without any url
        if ($onedriveOrSharePointUrls -contains "all") {
            $capabilities += @{
                "name" = "OneDriveAndSharePoint"
            }
        }
        else {
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


    }

    if ($graphConnectorIds -and $graphConnectorIds.Count -gt 0) {

        # if $graphConnectorIds contains "all", then add the capability "GraphConnectors" without any id
        if ($graphConnectorIds -contains "all") {
            $capabilities += @{
                "name" = "GraphConnectors"
            }
        }
        else {

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
    }

    # if any capibility is enabled, update the capabilities
    if ($capabilities.Count -gt 0) {
        $copilot | Add-Member -MemberType NoteProperty -Name "capabilities" -Value $capabilities
    }

    # if ($actionFiles -and $actionFiles.Count -gt 0) {
    #     $actions = @(
    #         foreach ($actionFile in $actionFiles) {
    #             # copy the action file to the temp folder
    #             $actionFileName = Split-Path $actionFile -Leaf
    #             Copy-Item -Path $actionFile -Destination (Join-Path $tempFolder -ChildPath $actionFileName) -Force

    #             [ordered]@{
    #                 "id"   = [Guid]::NewGuid().ToString()
    #                 "file" = $actionFileName
    #             }
    #         }
    #     )

    #     $copilot | Add-Member -MemberType NoteProperty -Name "actions" -Value $actions
    # }

    # save the updated declarativecopilot.json to the same file
    $copilot | ConvertTo-Json -Depth 10 | Set-Content -Path (Join-Path $tempFolder "declarativecopilot.json") -Force -Encoding UTF8


    # create a zip file and save it in current folder, named as $name.zip
    $zipFile = "$name.zip"

    # compress the temp folder to a zip file
    Compress-Archive -Path "$tempFolder\\*" -DestinationPath $zipFile -Force


    # remove the temp folder
    Remove-Item -Path $tempFolder -Recurse -Force

    Write-Host "The Declarative Agent app has been created successfully. The zip file is saved as $zipFile."

}