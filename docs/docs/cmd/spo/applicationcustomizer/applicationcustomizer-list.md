# spo applicationcustomizer list

Get a list of application customizers that are added to a site.

## Usage

```sh
m365 spo applicationcustomizer list [options]
```

## Options

`-u, --webUrl <webUrl>`
: The url of the site.

`-s, --scope [scope]`
: Scope of the application customizers. Allowed values `Site`, `Web`, `All`. Defaults to `All`

--8<-- "docs/cmd/_global.md"

## Examples

Retrieves a list of application customizers.

```sh
m365 spo applicationcustomizer list --webUrl https://contoso.sharepoint.com/sites/sales
```

## Response

=== "JSON"

    ```json
    [
      {
        "ClientSideComponentId": "4358e70e-ec3c-4713-beb6-39c88f7621d1",
        "ClientSideComponentProperties": "{\"listTitle\":\"News\",\"listViewTitle\":\"Published News\"}",
        "CommandUIExtension": null,
        "Description": null,
        "Group": null,
        "HostProperties": "",
        "Id": "f405303c-6048-4636-9660-1b7b2cadaef9",
        "ImageUrl": null,
        "Location": "ClientSideExtension.ApplicationCustomizer",
        "Name": "{f405303c-6048-4636-9660-1b7b2cadaef9}",
        "RegistrationId": null,
        "RegistrationType": 0,
        "Rights": {
          "High": 0,
          "Low": 0
        },
        "Scope": 3,
        "ScriptBlock": null,
        "ScriptSrc": null,
        "Sequence": 65536,
        "Title": "NewsTicker",
        "Url": null,
        "VersionOfUserCustomAction": "1.0.1.0"
      }
    ]
    ```

=== "Text"

    ```text
    Name                                    Location                                   Scope  Id
    --------------------------------------  -----------------------------------------  -----  ------------------------------------
    {f405303c-6048-4636-9660-1b7b2cadaef9}  ClientSideExtension.ApplicationCustomizer  3      f405303c-6048-4636-9660-1b7b2cadaef9
    ```

=== "CSV"

    ```csv
    Name,Location,Scope,Id
    {f405303c-6048-4636-9660-1b7b2cadaef9},ClientSideExtension.ApplicationCustomizer,3,f405303c-6048-4636-9660-1b7b2cadaef9
    ```

=== "Markdown"

    ```md
    # spo applicationcustomizer list --webUrl "https://contoso.sharepoint.com"

    Date: 28/2/2023

    ## NewsTicker (f405303c-6048-4636-9660-1b7b2cadaef9)

    Property | Value
    ---------|-------
    ClientSideComponentId | 4358e70e-ec3c-4713-beb6-39c88f7621d1
    ClientSideComponentProperties | {"listTitle":"News","listViewTitle":"Published News"}
    CommandUIExtension | null
    Description | null
    Group | null
    HostProperties |
    Id | f405303c-6048-4636-9660-1b7b2cadaef9
    ImageUrl | null
    Location | ClientSideExtension.ApplicationCustomizer
    Name | {f405303c-6048-4636-9660-1b7b2cadaef9}
    RegistrationId | null
    RegistrationType | 0
    Rights | {"High":0,"Low":0}
    Scope | 3
    ScriptBlock | null
    ScriptSrc | null
    Sequence | 65536
    Title | NewsTicker
    Url | null
    VersionOfUserCustomAction | 1.0.1.0
    ```
