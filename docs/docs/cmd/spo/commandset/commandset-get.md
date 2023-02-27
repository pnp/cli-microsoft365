# spo commandset get

Get a ListView Command Set that is added to a site.

## Usage

```sh
m365 spo commandset get [options]
```

## Options

`-u, --webUrl <webUrl>`
: Url of the site.

`-t, --title [title]`
: The title of the ListView Command Set. Specify either `title`, `id` or `clientSideComponentId`.

`-i, --id [id]`
: The id of the ListView Command Set. Specify either `title`, `id` or `clientSideComponentId`.

`-c, --clientSideComponentId [clientSideComponentId]`
: The id of the ListView Command Set. Specify either `title`, `id` or `clientSideComponentId`.

`-s, --scope [scope]`
: Scope of the ListView Command Set. Allowed values: `Site`, `Web`, `All`. Defaults to `All`.

--8<-- "docs/cmd/_global.md"

## Remarks

If the command finds multiple command sets with the specified title, it will prompt you to disambiguate which command set it should use, listing the discovered IDs.

## Examples

Retrieves an ListView Command Set by title.

```sh
m365 spo commandset get --title "Some customizer" --webUrl https://contoso.sharepoint.com/sites/sales
```

Retrieves an ListView Command Set by id with scope 'Web'.

```sh
m365 spo commandset get --id 14125658-a9bc-4ddf-9c75-1b5767c9a337 --webUrl https://contoso.sharepoint.com/sites/sales -scope Web
```

Retrieves an ListView Command Set by clientSideComponentId with scope 'Site'.

```sh
m365 spo commandset get --clientSideComponentId c1cbd896-5140-428d-8b0c-4873be19f5ac --webUrl https://contoso.sharepoint.com/sites/sales --scope Site
```

## Response

=== "JSON"

    ```json
    {
      "ClientSideComponentId": "c1cbd896-5140-428d-8b0c-4873be19f5ac",
      "ClientSideComponentProperties": "{\"sampleTextOne\":\"One item is selected in the list.\", \"sampleTextTwo\":\"This command is always visible.\"}",
      "CommandUIExtension": null,
      "Description": null,
      "Group": null,
      "HostProperties": "",
      "Id": "9a0674de-2f3d-4a26-ba79-62b460ddd327",
      "ImageUrl": null,
      "Location": "ClientSideExtension.ListViewCommandSet.CommandBar",
      "Name": "{9a0674de-2f3d-4a26-ba79-62b460ddd327}",
      "RegistrationId": "100",
      "RegistrationType": 1,
      "Rights": {
        "High": "0",
        "Low": "0"
      },
      "Scope": 3,
      "ScriptBlock": null,
      "ScriptSrc": null,
      "Sequence": 65536,
      "Title": "Notification",
      "Url": null,
      "VersionOfUserCustomAction": "1.0.1.0"
    }
    ```

=== "Text"

    ```text
    ClientSideComponentId        : c1cbd896-5140-428d-8b0c-4873be19f5ac
    ClientSideComponentProperties: {"sampleTextOne":"One item is selected in the list.", "sampleTextTwo":"This command is always visible."}
    CommandUIExtension           : null
    Description                  : null
    Group                        : null
    HostProperties               :
    Id                           : 9a0674de-2f3d-4a26-ba79-62b460ddd327
    ImageUrl                     : null
    Location                     : ClientSideExtension.ListViewCommandSet.CommandBar
    Name                         : {9a0674de-2f3d-4a26-ba79-62b460ddd327}
    RegistrationId               : 100
    RegistrationType             : 1
    Rights                       : {"High":"0","Low":"0"}
    Scope                        : 3
    ScriptBlock                  : null
    ScriptSrc                    : null
    Sequence                     : 65536
    Title                        : Notification
    Url                          : null
    VersionOfUserCustomAction    : 1.0.1.0
    ```

=== "CSV"

    ```csv
    ClientSideComponentId,ClientSideComponentProperties,CommandUIExtension,Description,Group,HostProperties,Id,ImageUrl,Location,Name,RegistrationId,RegistrationType,Rights,Scope,ScriptBlock,ScriptSrc,Sequence,Title,Url,VersionOfUserCustomAction
    c1cbd896-5140-428d-8b0c-4873be19f5ac,"{""sampleTextOne"":""One item is selected in the list."", ""sampleTextTwo"":""This command is always visible.""}",,,,,9a0674de-2f3d-4a26-ba79-62b460ddd327,,ClientSideExtension.ListViewCommandSet.CommandBar,{9a0674de-2f3d-4a26-ba79-62b460ddd327},100,1,"{""High"":""0"",""Low"":""0""}",3,,,65536,Notification,,1.0.1.0
    ```

=== "Markdown"

    ```md
    # spo commandset get --webUrl "https://contoso.sharepoint.com/sites/sales" --id "9a0674de-2f3d-4a26-ba79-62b460ddd327"

    Date: 27/02/2023

    ## Notification (9a0674de-2f3d-4a26-ba79-62b460ddd327)

    Property | Value
    ---------|-------
    ClientSideComponentId | c1cbd896-5140-428d-8b0c-4873be19f5ac
    ClientSideComponentProperties | {"sampleTextOne":"One item is selected in the list.", "sampleTextTwo":"This command is always visible."}
    CommandUIExtension | null
    Description | null
    Group | null
    HostProperties |
    Id | 9a0674de-2f3d-4a26-ba79-62b460ddd327
    ImageUrl | null
    Location | ClientSideExtension.ListViewCommandSet.CommandBar
    Name | {9a0674de-2f3d-4a26-ba79-62b460ddd327}
    RegistrationId | 100
    RegistrationType | 1
    Rights | {"High":"0","Low":"0"}
    Scope | 3
    ScriptBlock | null
    ScriptSrc | null
    Sequence | 65536
    Title | Notification
    Url | null
    VersionOfUserCustomAction | 1.0.1.0
    ```
