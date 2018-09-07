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
`-i, --id <id>`|ID of the app to retrieve information for
`-s, --siteUrl <siteUrl>`|Absolute URL of the site to install the app in
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to a SharePoint Online site, using the [spo login](../login.md) command.

## Remarks

To upgrade an app in the site, you have to first log in to a SharePoint site using the [spo login](../login.md) command, eg. `spo login https://contoso.sharepoint.com`.

## Examples

Upgrade the app with ID _b2307a39-e878-458b-bc90-03bc578531d6_ in the _https://contoso.sharepoint.com_ site.

```sh
spo app upgrade -i b2307a39-e878-458b-bc90-03bc578531d6 -s https://contoso.sharepoint.com
```

## More information

- Application Lifecycle Management (ALM) APIs: [https://docs.microsoft.com/en-us/sharepoint/dev/apis/alm-api-for-spfx-add-ins](https://docs.microsoft.com/en-us/sharepoint/dev/apis/alm-api-for-spfx-add-ins)