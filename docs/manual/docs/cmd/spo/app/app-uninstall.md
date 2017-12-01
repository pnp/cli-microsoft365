# spo app uninstall

Uninstalls an app from the site

## Usage

```sh
spo app uninstall [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --id <id>`|ID of the app to retrieve information for
`-s, --siteUrl <siteUrl>`|Absolute URL of the site to uninstall the app from
`--confirm`|Don't prompt for confirming uninstalling the app
`-o, --output <output>`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online site, using the [spo connect](../connect.md) command.

## Remarks

To uninstall an app from the site, you have to first connect to a SharePoint site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso.sharepoint.com`.

## Examples

```sh
spo app uninstall -i b2307a39-e878-458b-bc90-03bc578531d6 -s https://contoso.sharepoint.com
```

Uninstalls the app with ID _b2307a39-e878-458b-bc90-03bc578531d6_ from the _https://contoso.sharepoint.com_ site.

```sh
spo app uninstall -i b2307a39-e878-458b-bc90-03bc578531d6 -s https://contoso.sharepoint.com
```

Uninstalls the app with ID _b2307a39-e878-458b-bc90-03bc578531d6_ from the _https://contoso.sharepoint.com_ site without prompting for confirmation.

## More information

- Application Lifecycle Management (ALM) APIs: [https://docs.microsoft.com/en-us/sharepoint/dev/apis/alm-api-for-spfx-add-ins](https://docs.microsoft.com/en-us/sharepoint/dev/apis/alm-api-for-spfx-add-ins)