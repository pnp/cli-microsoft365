# spo app upgrade

Upgrades app in the specified site

## Usage

```sh
spo app upgrade [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --id <id>`|ID of the app to upgrade
`-s, --siteUrl <siteUrl>`|Absolute URL of the site to upgrade the app in
`--scope [scope]`|Scope of the app catalog: `tenant|sitecollection`. Default `tenant`
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to a SharePoint Online site, using the [spo login](../login.md) command.

## Remarks

To upgrade an app in the site, you have to first log in to a SharePoint site using the [spo login](../login.md) command, eg. `spo login https://contoso.sharepoint.com`.

If the app with the specified ID doesn't exist in the app catalog, the command will fail with an error.

## Examples

Upgrade the app with ID _b2307a39-e878-458b-bc90-03bc578531d6_ in the _https://contoso.sharepoint.com_ site.

```sh
spo app upgrade --id b2307a39-e878-458b-bc90-03bc578531d6 --siteUrl https://contoso.sharepoint.com
```

Upgrade the app with ID _b2307a39-e878-458b-bc90-03bc578531d6_ in the _https://contoso.sharepoint.com_ site from site collection app catalog.

```sh
spo app upgrade --id b2307a39-e878-458b-bc90-03bc578531d6 --siteUrl https://contoso.sharepoint.com --scope sitecollection
```

## More information

- Application Lifecycle Management (ALM) APIs: [https://docs.microsoft.com/en-us/sharepoint/dev/apis/alm-api-for-spfx-add-ins](https://docs.microsoft.com/en-us/sharepoint/dev/apis/alm-api-for-spfx-add-ins)