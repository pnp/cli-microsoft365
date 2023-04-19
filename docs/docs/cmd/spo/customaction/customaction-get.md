# spo customaction get

Gets information about a user custom action for site or site collection

## Usage

```sh
m365 spo customaction get [options]
```

## Options

`-i, --id [id]`
: ID of the user custom action to retrieve information for. Specify either `id`, `title` or `clientSideComponentId`

`-t, --title [title]`
: Title of the user custom action to retrieve information for. Specify either `id`, `title` or `clientSideComponentId`

`-c, --clientSideComponentId [clientSideComponentId]`
: clientSideComponentId of the user custom action to retrieve information for. Specify either `id`, `title` or `clientSideComponentId`

`-u, --webUrl <webUrl>`
: Url of the site or site collection to retrieve the custom action from

`-s, --scope [scope]`
: Scope of the custom action. Allowed values `Site,Web,All`. Default `All`

--8<-- "docs/cmd/_global.md"

## Remarks

If the command finds multiple user custom actions with the specified `title` or `clientSideComponentId`, it will prompt you to disambiguate which user custom action it should get, listing the discovered IDs.

## Examples

Return details about the user custom action based on the id and a given url

```sh
m365 spo customaction get --id 058140e3-0e37-44fc-a1d3-79c487d371a3 --webUrl https://contoso.sharepoint.com/sites/test
```

Return details about the user custom action based on the title and a given url

```sh
m365 spo customaction get --title "YourAppCustomizer" --webUrl https://contoso.sharepoint.com/sites/test
```

Return details about the user custom action based on the clientSideComponentId and a given url

```sh
m365 spo customaction get --clientSideComponentId "34a019f9-6198-4053-a3b6-fbdea9a107fd" --webUrl https://contoso.sharepoint.com/sites/test
```

Return details about the user custom action based on the id and a given url and the scope

```sh
m365 spo customaction get --id "058140e3-0e37-44fc-a1d3-79c487d371a3" --webUrl https://contoso.sharepoint.com/sites/test --scope Site
```

Return details about the user custom action based on the id and a given url and the scope

```sh
m365 spo customaction get --id "058140e3-0e37-44fc-a1d3-79c487d371a3" --webUrl https://contoso.sharepoint.com/sites/test --scope Web
```

Return details about the user custom action based on the id and a given url and the scope

```sh
m365 spo customaction get --id "058140e3-0e37-44fc-a1d3-79c487d371a3" --webUrl https://contoso.sharepoint.com/sites/test --scope Web
```

## Response

=== "JSON"

    ```json
    {
      "ClientSideComponentId": "34a019f9-6198-4053-a3b6-fbdea9a107fd",
      "ClientSideComponentProperties": "{\"sampleTextOne\":\"One item is selected in the list.\", \"sampleTextTwo\":\"This command is always visible.\"}",
      "CommandUIExtension": null,
      "Description": null,
      "Group": null,
      "Id": "158cb0d1-8703-4a36-866d-84aed8233bd3",
      "ImageUrl": null,
      "Location": "ClientSideExtension.ListViewCommandSet.CommandBar",
      "Name": "{158cb0d1-8703-4a36-866d-84aed8233bd3}",
      "RegistrationId": "100",
      "RegistrationType": 1,
      "Rights": "{\"High\":0,\"Low\":0}",
      "Scope": "Web",
      "ScriptBlock": null,
      "ScriptSrc": null,
      "Sequence": 65536,
      "Title": "ExtensionTraining",
      "Url": null,
      "VersionOfUserCustomAction": "1.0.1.0"
    }
    ```

=== "Text"

    ```text
    ClientSideComponentId        : 34a019f9-6198-4053-a3b6-fbdea9a107fd
    ClientSideComponentProperties: {"sampleTextOne":"One item is selected in the list.", "sampleTextTwo":"This command is always visible."}
    CommandUIExtension           : null
    Description                  : null
    Group                        : null
    Id                           : 158cb0d1-8703-4a36-866d-84aed8233bd3
    ImageUrl                     : null
    Location                     : ClientSideExtension.ListViewCommandSet.CommandBar
    Name                         : {158cb0d1-8703-4a36-866d-84aed8233bd3}
    RegistrationId               : 100
    RegistrationType             : 1
    Rights                       : {"High":0,"Low":0}
    Scope                        : Web
    ScriptBlock                  : null
    ScriptSrc                    : null
    Sequence                     : 65536
    Title                        : ExtensionTraining
    Url                          : null
    VersionOfUserCustomAction    : 1.0.1.0
    ```

=== "CSV"

    ```csv
    ClientSideComponentId,ClientSideComponentProperties,CommandUIExtension,Description,Group,Id,ImageUrl,Location,Name,RegistrationId,RegistrationType,Rights,Scope,ScriptBlock,ScriptSrc,Sequence,Title,Url,VersionOfUserCustomAction
    34a019f9-6198-4053-a3b6-fbdea9a107fd,"{""sampleTextOne"":""One item is selected in the list."", ""sampleTextTwo"":""This command is always visible.""}",,,,158cb0d1-8703-4a36-866d-84aed8233bd3,,ClientSideExtension.ListViewCommandSet.CommandBar,{158cb0d1-8703-4a36-866d-84aed8233bd3},100,1,"{""High"":0,""Low"":0}",Web,,,65536,ExtensionTraining,,1.0.1.0
    ```

=== "Markdown"

    ```md
    # spo customaction get --webUrl "https://contoso.sharepoint.com" --clientSideComponentId "34a019f9-6198-4053-a3b6-fbdea9a107fd" --scope "Web"

    Date: 27/1/2023

    ## ExtensionTraining (158cb0d1-8703-4a36-866d-84aed8233bd3)

    Property | Value
    ---------|-------
    ClientSideComponentId | 34a019f9-6198-4053-a3b6-fbdea9a107fd
    ClientSideComponentProperties | {"sampleTextOne":"One item is selected in the list.", "sampleTextTwo":"This command is always visible."}
    CommandUIExtension | null
    Description | null
    Group | null
    Id | 158cb0d1-8703-4a36-866d-84aed8233bd3
    ImageUrl | null
    Location | ClientSideExtension.ListViewCommandSet.CommandBar
    Name | {158cb0d1-8703-4a36-866d-84aed8233bd3}
    RegistrationId | 100
    RegistrationType | 1
    Rights | {"High":0,"Low":0}
    Scope | Web
    ScriptBlock | null
    ScriptSrc | null
    Sequence | 65536
    Title | ExtensionTraining
    Url | null
    VersionOfUserCustomAction | 1.0.1.0
    ```
